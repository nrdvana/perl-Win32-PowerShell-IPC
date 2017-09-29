#! perl
use strict;
use warnings;
use Win32;
use Test::More;
use Log::Any::Adapter 'TAP';

use_ok 'Win32::PowerShell::IPC';

my $ps= new_ok( 'Win32::PowerShell::IPC', [], 'IPC instance' );

ok( $ps->spawn, 'Start child process' );

is( $ps->run_command("echo Foo"), "Foo\r\n", 'Run echo command' );

undef $ps; # should not throw exception
pass( 'no exception during DESTROY' );

# TODO: find a way to test that the powershell process is no longer running

done_testing;
