use ExtUtils::MakeMaker;
# See lib/ExtUtils/MakeMaker.pm for details of how to influence
# the contents of the Makefile that is written.
WriteMakefile(
    NAME         => 'Win32::MSAgent',
    VERSION_FROM => 'lib/Win32/MSAgent.pm', # finds $VERSION
    AUTHOR       => 'Jouke Visser (jouke@cpan.org)',
    ABSTRACT     => 'Interface module for the Microsoft Agent'
);
