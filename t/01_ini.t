#!/usr/bin/perl
use strict;
use warnings;
use Test::More tests => 9;
use File::Find;

my ($agent, @langs, @voices, $systemroot, %chars, $char, $c);

BEGIN { 

# 1: use the installed module    
use_ok( 'Win32::MSAgent' ); 

}

# 2. test the constructor
ok( $agent = Win32::MSAgent->new(),                           'Can create a new Agent object');
# 3. test if $agent isa Win32::MSAgent
isa_ok($agent, 'Win32::MSAgent');

# 4. test if the languages can be retrieved
ok( @langs = $agent->Characters->ListLanguages(),             'Can fetch languages');
# 5. test if the voices for 'English (US)' can be retrieved
ok( @voices = $agent->Characters->ListVoices('English (US)'), 'Can fetch a voice');

# get the systemroot (under which the MS Agent characterfiles are installed somewhere
$systemroot = $ENV{SYSTEMROOT} || $ENV{WINDIR} || 'C:\WINDOWS';
print "I'm going to look for MS Agent Character files (.ACS files) under $systemroot now\n";

# Find installed characters on this system
find(sub {$chars{$_} = $File::Find::name if $_=~ /.*?\.acs$/i; }, $systemroot);
print "I did not find any Microsoft Agent Character files on your system\n" unless %chars;
$char = (keys %chars)[0];
$char =~ s/(.*)?\..*/$1/;
print "I found at least character $char; going to continue testing with $char\n" if $char;

SKIP:{
skip("Skipping character tests", 4) unless $char;

# 6. Test if we can load the $char
ok( $agent->Characters->Load($char, "$char.acs"),           "The $char character is installed and loaded" );
# 7. Test if we can use the $char
ok( $c = $agent->Characters($char),                         "We can use the character $char");
# 8. Test if we can Show the $char
ok( $c->Show(),                                             "We can show the character $char");
# 9. Test if we can move the $char
ok( $c->MoveTo(300,300),                                    "We can move to another position");
sleep(5);
}
