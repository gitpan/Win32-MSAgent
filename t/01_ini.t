#!/usr/bin/perl
use strict;
use warnings;
use Test::Simple tests => 3;
use Win32::MSAgent;

my $agent = Win32::MSAgent->new();

ok( defined $agent , 'Constructor works');
ok( ref($agent) eq 'Win32::MSAgent', 'Constructor creates the right class' );
ok( $agent->Characters->Load('Genie', "genie.acs"), 'The Genie character is installed and loaded' );
