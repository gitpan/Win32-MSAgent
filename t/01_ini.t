#!/usr/bin/perl
use strict;
use warnings;
use Win32::MSAgent;
use Test::Simple tests => 8;
use File::Find;

my $agent = Win32::MSAgent->new();
my (@langs, @voices);

ok( defined $agent ,                                          'Constructor works');
ok( ref($agent) eq 'Win32::MSAgent',                          'Constructor creates the right class' );
ok( @langs = $agent->Characters->ListLanguages(),             'Can fetch languages');
ok( @voices = $agent->Characters->ListVoices('English (US)'), 'Can fetch a voice');

# Find installed characters on this system
my %chars;
find(sub {$chars{$_} = $File::Find::name if $_=~ /.*?\.acs$/i; }, "c:\\windows");
print "I did not find any Microsoft Agent Character files on your system\n" unless %chars;
my $char = (keys %chars)[0];
$char =~ s/(.*)?\..*/$1/;
print "I found character $char\n";
my $c;
ok( $agent->Characters->Load($char, "$char.acs"),           "The $char character is installed and loaded" );
ok( $c = $agent->Characters($char),                         "We can use the character $char");
ok( $c->Show(),                                             "We can show the character $char");
ok( $c->MoveTo(300,300),                                    "We can move to another position");
sleep(5);
