package Win32::PowerShell::IPC;
use Moo 2;
use Win32;
use Win32::API;
use Win32::Process;
use Win32API::File 'FdGetOsFHandle';
use Try::Tiny;
use Carp;

# ABSTRACT: Set up IPC between Perl and a PowerShell child process

=head1 SYNOPSIS

  my $ps= Win32::PowerShell::IPC->new();
  $ps->run_or_die('$encrypted_pass = "'.$password.'" | ConvertTo-SecureString -AsPlainText -Force');
  $ps->run_or_die('$credential = New-Object System.Management.Automation.PSCredential'
    .' -ArgumentList "'.$username.'", $encrypted_pass');
  $ps->run_or_die('$sess = New-PSSession -ConfigurationName Microsoft.Exchange'
    .' -ConnectionUri https://ps.outlook.com/powershell'
    .' -Credential $credential -Authentication Basic -AllowRedirection');
  $ps->run_or_die('Import-PSSession $sess');
  # Followed by interactive use of the Exchange methods that PowerShell now has access to

=head1 DESCRIPTION

There's a lot of things in the Microsoft world that can't be done with perl.
This is even more true with many Microsoft services offering PowerShell
integration instead of more accessible Web APIs.  While you can certainly
write out a PowerShell script file and execute it, the session setup can be
extremely slow and you might want to run numerous commands and take action
based on the outcome.  And you might rather do the logic in Perl than write
all that in PowerShell, especially if it involves your database.

This module fires up a captive child PowerShell process to which you can
submit commands, and receive the results.   It's all text for now, but Perl
excels at messy stuff like this.

PowerShell also seems to offer an option to exchange commands and results as
XML, which would be a lot more reliable than text, but I haven't explored
this yet.  Patches welcome.  (and good grief, haven't they learned JSON over
at Redmond, yet?)

=head1 ATTRIBUTES

=head2 running

Whether PowerShell is running

=head2 exe_path

Full path to PowerShell.exe, lazy-built on demand from $PATH if you don't
specify it.  This attribute is writeable, but won't have any effect once the
child process is started.

=head2 proc

The L<Win32::Process> of PowerShell, initialized by L</spawn> or by calling any
of the run/begin methods.

=head2 cleanup_delay

When this object is destroyed, wait up to this many milliseconds for the
PowerShell process to exit.  (we send it an "exit" command)  Default is
2000. (2 seconds).

=head2 stdin

The write-side of PowerShell's STDIN pipe

=head2 stdout

The read-side of PowerShell's STDOUT pipe

=head2 stdout_h

The Windows Handle for the read handle of PowerShell's STDOUT pipe.
The Windows Handle is needed for calling Win32 API calls via i.e.
L<Win32::API> wrappers, which can't accept a perl globref.

=head2 rbuf

The accumulated read buffer from STDOUT of PowerShell.  Used to hold leftover
stream contents that might follow the end of a command output.

=cut

sub running       { defined shift->proc }
has exe_path      => ( is => 'rw', lazy => 1, builder => 1 );
has cleanup_delay => ( is => 'rw', default => sub { 2000 } );
has proc          => ( is => 'rwp' );

has stdin         => ( is => 'rwp' );
has stdout        => ( is => 'rwp' );
has stdout_h      => ( is => 'lazy', clearer => 1 );

has rbuf          => ( is => 'rw', default => sub { '' } );

has _command_boundary   => ( is => 'rw', default => sub { [] } );
has _command_seq_number => ( is => 'rw', default => sub { 0 } );

=head1 METHODS

=head2 new

Standard Moo constructor.  All attributes are optional.
You might consider setting L</cleanup_delay> or L</exe_path>

=head2 spawn

Start the PowerShell process.  Dies on any failure.  Returns true.

=cut

sub _build_exe_path {
	# Find PowerShell executable in PATH.  (CreateProcess ought to do this but doesn't for some reason?)
	my ($exe)= grep { -f $_  } map { "${_}\\PowerShell.exe" } split /;/, $ENV{PATH};
	defined $exe or croak "Can't locate PowerShell.exe in PATH: $ENV{PATH}";
	return $exe;
}

sub spawn {
	my $self= shift;
	if ($self->running) {
		carp "Subprocess is already running\n";
		return;
	}
	
	# make sure we have this before mucking around with file handles
	my $exe= $self->exe_path;
	
	my ($in_r, $in_w, $out_r, $out_w, $save_stdin, $save_stdout, $save_stderr);
	
	# Create pipes
	pipe($in_r, $in_w) or die "pipe: $!";
	pipe($out_r, $out_w) or die "pipe: $!";
	
	# Temporarily overwrite the main handles of this process (dangerous... but
	# Win32::Process doesn't support the other arguments to CreateProcess
	# to pass the file handles directly.)
	my @err;
	try {
		open $save_stdin, '<&', \*STDIN or die "Can't save STDIN: $!\n";
		open STDIN, '<&', $in_r or die "Can't redirect STDIN: $!\n";
		open $save_stdout, '>&', \*STDOUT or die "Can't save STDOUT: $!\n";
		open STDOUT, '>&', $out_w or die "Can't redirect STDOUT: $!\n";
		open $save_stderr, '>&', \*STDERR or die "Can't save STDERR: $!\n";
		open STDERR, '>&', $out_w or die "Can't redirect STDERR: $!\n";
		Win32::Process::Create(
			my $proc, $exe, "PowerShell -Command -",
			1, # inherit handles
			Win32::Process->NORMAL_PRIORITY_CLASS,
			'.' # cur dir
		) or die "Can't launch PowerShell: ".Win32::FormatMessage(Win32::GetLastError())."\n";
		$self->_set_proc($proc);
		$self->_set_stdin($in_w);
		$self->_set_stdout($out_r);
		$self->rbuf('');
		@{ $self->_command_boundary }= ();
	}
	catch {
		chomp;
		push @err, $_;
	};
	# Now restore handles before throwing exception
	open(STDIN, '<&', $save_stdin) or push(@err, "Can't restore STDIN: $!")
		if defined $save_stdin;
	open(STDOUT, '>&', $save_stdout) or push(@err, "Can't restore STDOUT: $!")
		if defined $save_stdout;
	open(STDERR, '>&', $save_stderr) or push(@err, "Can't restore STDERR: $!")
		if defined $save_stderr;
	croak join("; ", @err) if @err;
	return 1;
}

# Check if the process has exited, or wait for it to exit, or optionally kill
# it if the timeout expires.
sub _wait_or_kill {
	my ($self, $timeout, $kill)= @_;
	return 1 unless $self->proc;
	my $exit_code;
	if ($self->proc->Wait($timeout)) {
		$self->proc->GetExitCode($exit_code);
	} elsif ($kill) {
		$self->proc->Kill(255);
	}
	if ($kill or defined $exit_code) {
		$self->_set_proc(undef);
		$self->_set_stdin(undef);
		$self->_set_stdout(undef);
		$self->clear_stdout_h;
		return 1;
	}
	return 0;
}

sub DESTROY {
	my $self= shift;
	# If powershell is still running, tell it to exit, and give it 2 seconds
	# to exit gracefully before killing it.
	if ($self->running) {
		try { $self->write_all("exit\r\n"); };
		$self->_wait_or_kill($self->cleanup_delay, 1)
			if $self->running; # write_all can also reap the process
	}
}

=head2 begin_command

  $ps->begin_command("Text to execute");

This sends the text to the PowerShell instance, or dies trying.
It does not wait for a result.  It actually also sends an "echo"
command that is used to detect the end of the output from your
command.

=cut

sub begin_command {
	my ($self, $command)= @_;
	$self->spawn unless $self->running;
	my $boundary;
	do {
		$boundary= sprintf('END_COMMAND_%d_%X_%s', ++$self->{_command_seq_number}, time, rand)
	} while index($command, $boundary) >= 0;
	push @{ $self->_command_boundary }, $boundary;
	$command .= "\r\n" unless $command =~ /\r\n$/;
	$self->write_all($command);
	$self->write_all("echo $boundary;\r\n");
}

=head2 collect_command

  my $output= $ps->collect_command;

This blocks until it receives the full output from your oldest pending
command, and no other command. (you may have multiple pending commands)
This module delimits the output with "echo" statements so that it can tell
where the output of a command ends, but you shouldn't ever see signs of this
implementation detail.  I hope.

If you don't want to block, see L</stdout_readable> and L</read_more>.
Not that there's a complete solution there... but it will get you a little
farther.

=cut

sub collect_command {
	my ($self)= @_;
	my $next_boundary= $self->_command_boundary->[0]
		or croak "No command is pending";
	while ($self->{rbuf} !~ /\Q$next_boundary\E\r?$/) {
		$self->read_more;
	}
	$self->{rbuf} =~ s/(.*)\Q$next_boundary\E\r?$//ms;
	shift @{$self->_command_boundary};
	return $1;
}

=head2 run_command

  my $output= $ps->run_command("My Command;");

Send the command and then wait for the result.
This is like L</begin_command> + L</collect_command> except that it also
discards the output of any previous commands to make sure that you're getting
the output from I<this> command.

=cut

sub run_command {
	my ($self, $command)= @_;
	$self->begin_command($command);
	my $ret;
	$ret= $self->collect_command
		while @{ $self->_command_boundary };
	return $ret;
}

=head2 run_or_die

Like L</run_command>, but if the output looks like a PowerShell exception report,
die instead of return.

=cut

sub run_or_die {
	my ($self, $command)= @_;
	my $result= $self->run_command($command);
	croak $result if $result =~ /^\s*\+ FullyQualifiedErrorId\s*:/m;
	return $result;
}

=head2 write_all

Lower-level method to write all data to the PowerShell pipe, but also die if
PowerShell isn't running.

=cut

sub write_all {
	my ($self, $buf)= @_;
	my $ret;
	$self->running or croak "Powershell not started";
	$self->_wait_or_kill(0, 0) and croak "Powershell exited";
	while (length($buf) && (($ret= syswrite($self->stdin, $buf))||0) > 0) {
		#print STDERR "wrote '".substr($buf, 0, $ret)."'\n";
		substr($buf, 0, $ret)= '';
	}
	croak "syswrite: $!" unless defined $ret;
	return $ret > 0;
}

=head2 read_more

Lower-level method to read more data from the pipe into L</rbuf> but also
die if PowerShell isn't running.

=cut

sub read_more {
	my $self= shift;
	$self->_wait_or_kill(0, 0) and croak "Powershell exited";
	my $ret= sysread($self->stdout, my $buf, 4096);
	defined $ret or croak "sysread: $!";
	#print STDERR "read '".$buf."\n";
	$self->{rbuf} .= $buf;
	return $ret;
}

=head2 stdout_readable

  if ($ps->stdout_readable) {
    $ps->read_more;
    ... # now inspect $ps->rbuf
  }

Calls L</PeekNamedPipe> on C<< $ps->stdout_h >> file handle, and returns
true if there are bytes available.

=head1 FUNCTIONS

=head2 PeekNamedPipe

  Win32::PowerShell::IPC::PeekNamedPipe(
    $win32_handle, $buffer, $buffer_size,
    $bytes_read, $bytes_available, $bytes_remaining_always_zero
  );

The Windows API is really dismissive of the concept of non-blocking
single-threaded operation that most of the Unix world loves so much.
There is no way to perform a non-blocking read or select() on any windows
handle other than a socket.  Your options are either to dive into the crazy
mess of the overlapped (asynchronous) I/O API, or make one thread per handle,
or mess around with WaitForMultipleObjects which can only listen to 64 things
at once.

The one little workaround I found available for pipes is the PeekNamedPipe
function, which can tell you if there is any data on the pipe.  You can't wait
with a timeout, but it at least gives the option of a crude check/sleep loop.

This is not a method, but a regular function.  The first argument is a Win32
handle (not perl globref), which you can get from L<Win32API::File/FdGetOsFHandle>.

The second and third are the buffer and number of bytes to read.  If the first
is not defined I enlarge it to the specified size, and if the size is undefined
I use the existing size of the buffer.  (if both are undef, I pass NULL which
doesn't read anything).

The C<$bytes_read> I<receives> the value of the number of bytes written to the
buffer, but you can ignore it because I resize the buffer to match.  Set to
undef if you like.

The C<$bytes_available> is the useful one, and I<receives> the value of the
number of bytes in the pipe.

The final argument isn't useful for byte stream pipes, but I included it anyway.

Returns true if it succeeded.  Check C<< Win32::FormatMessage(Win32::GetLastError()) >>
otherwise.

=cut

sub _build_stdout_h {
	my $self= shift;
	my $out= $self->stdout or croak "Stdout not open";
	return FdGetOsFHandle(fileno($out));
}

sub stdout_readable {
	my $self= shift;
	my $n;
	return PeekNamedPipe($self->stdout_h, undef, 0, undef, $n, undef) && $n > 0;
}

my $peek_named_pipe;
sub PeekNamedPipe {
	$peek_named_pipe ||= Win32::API->new("kernel32", 'PeekNamedPipe', 'NPIPPP', 'N')
		|| die "Can't load PeekNamedPipe from kernel32.dll";
	# hNamedPipe  - Windows Handle (integer)
	# lpBuffer    - Destination buffer where bytes will be written. NULL if not needed.
	# nBufferSize - size of buffer.
	# lpBytesRead - destination DWORD of number of bytes stored in buffer.  NULL if not needed.
	# lpTotalBytesAvail - destination DWORD of number of bytes available.   NULL if not needed.
	# lpBytesLeftThisMessage - (not relevant for named or anonymous pipes)  NULL if not needed.
	
	if ($_[2]) { # if buffer length specified and nonzero, make buffer that long.
		$_[1]= "\0" x $_[2];
	}
	my $ret= $peek_named_pipe->Call(
		$_[0],
		$_[1],
		(defined $_[1]? length($_[1]) : 0), # use actual length of buffer
		(my $got_buf  = pack('L', 0)),
		(@_ > 4? (my $avail_buf= pack('L', 0)) : undef),
		(@_ > 5? (my $left_buf = pack('L', 0)) : undef),
	);
	if ($ret) {
		# Decode each of the DWORDs that were not passed as NULL
		my $got= unpack('L', $got_buf);
		substr($_[1], $got > 0? $got : 0)= '' if defined $_[1];
		$_[3]= $got if @_ > 3;
		$_[4]= unpack('L', $avail_buf) if @_ > 4;
		$_[5]= unpack('L', $left_buf ) if @_ > 5;
	}
	return $ret;
}

=head1 SEE ALSO

=head2 L<IPC::Run>

A well-maintained module for running child processes and bi-directional
communication with them.  However, the list of caveats for Win32 platform is
somewhat alarming.  (but I totally understand the difficulty it solves)
I wasn't comfortable with using that in production, so I wrote this module with
read/write on regular pipes.

=head2 L<PowerShell>

A module to produce PowerShell command syntax using perl object methods.

=head2 L<Win32::Process>

Used by this module.  Wraps a background process for Win32 environments where
fork/exec is broken, such as Strawberry Perl.

=head2 L<Win32::Job>

I wish I'd found this module sooner, since it is something of an improvement
over Win32::Process.  However, it doesn't provide a method to check if a
process is still running without killing the process, so ultimately not usable
for this module.

=cut

1;
