package Win32::MSAgent;
use strict;
use warnings;
use Win32::OLE;
use Win32::OLE::Variant;

BEGIN {
	use Exporter ();
	use vars qw ($VERSION @ISA @EXPORT @EXPORT_OK %EXPORT_TAGS);
	$VERSION     = 0.01;
	@ISA         = qw (Exporter);
	#Give a hoot don't pollute, do not export more than needed by default
	@EXPORT      = qw ();
	@EXPORT_OK   = qw ();
	%EXPORT_TAGS = ();
}

sub new
{
    my $proto = shift;
    my $char = shift;
    my $self = {};
    my $class = ref($proto) || $proto;
    bless $self, $class;
    Win32::OLE->Initialize(Win32::OLE::COINIT_MULTITHREADED);
    $self->{_agent} = Win32::OLE->new('Agent.Control.2') || die Win32::OLE->LastError();
    $self->Connected(1);
    if ($char)
    {
        $self->Characters->Load($char, "$char.acs");
    }

    return $self;
}

sub ShowDefaultCharacterProperties
{
    my $self = shift;
    my %opts = @_;
    my $x = $opts{x} || $opts{X};
    my $y = $opts{y} || $opts{y};
    if ((defined $x) && (defined $y))
    {
        $self->{_agent}->ShowDefaultCharacterProperties($x, $y);
    }
    else
    {
        $self->{_agent}->ShowDefaultCharacterProperties();
    }
}

sub Connected
{
    my $self = shift;
    my $bool = shift;
    $self->{_agent}->{Connected} = Variant(VT_BOOL, $bool);
    return $self->{_agent}->{Connected}
}

sub Name
{
    # not implemented
    return undef;
}

sub RaiseRequestErrors
{
    my $self = shift;
    my $bool = shift;
    $self->{_agent}->{RaiseRequestErrors} = Variant(VT_BOOL, $bool);
    return $self->{_agent}->{RaiseRequestErrors}
}

sub Characters
{
    my $self = shift;
    $self->{_Characters} = Win32::MSAgent::Characters->new($self->{_agent}, @_) 
        if ((@_) or (not exists $self->{_Characters}));
    return $self->{_Characters};
}

package Win32::MSAgent::Characters;

use strict;
use warnings;
use Win32::OLE;
use Win32::OLE::Variant;
use Win32::OLE::Enum;

our $Voices ={'US English'        => {'Adult Female #1' => '{CA141FD0-AC7F-11D1-97A3-006008273008}', 
                                      'Adult Female #2' => '{CA141FD0-AC7F-11D1-97A3-006008273009}', 
                                      'Adult Male #1'   => '{CA141FD0-AC7F-11D1-97A3-006008273000}',
                                      'Adult Male #2'   => '{CA141FD0-AC7F-11D1-97A3-006008273001}',
                                      'Adult Male #3'   => '{CA141FD0-AC7F-11D1-97A3-006008273002}',
                                      'Adult Male #4'   => '{CA141FD0-AC7F-11D1-97A3-006008273003}',
                                      'Adult Male #5'   => '{CA141FD0-AC7F-11D1-97A3-006008273004}',
                                      'Adult Male #6'   => '{CA141FD0-AC7F-11D1-97A3-006008273005}',
                                      'Adult Male #7'   => '{CA141FD0-AC7F-11D1-97A3-006008273006}',
                                      'Adult Male #8'   => '{CA141FD0-AC7F-11D1-97A3-006008273007}'},
            'British English'     => {'Carol'    => '{227A0E40-A92A-11d1-B17B-0020AFED142E}',
                                      'Peter'    => '{227A0E41-A92A-11d1-B17B-0020AFED142E}'},
            'Dutch'               => {'Linda'    => '{A0DDCA40-A92C-11d1-B17B-0020AFED142E}',
                                      'Alexander'=> '{A0DDCA41-A92C-11d1-B17B-0020AFED142E}'},
            'French'              => {'Véronique'=> '{0879A4E0-A92C-11d1-B17B-0020AFED142E}',
                                      'Pierre'   => '{0879A4E1-A92C-11d1-B17B-0020AFED142E}'},
            'German'              => {'Anna'     => '{3A1FB760-A92B-11d1-B17B-0020AFED142E}',
                                      'Stefan'   => '{3A1FB761-A92B-11d1-B17B-0020AFED142E}'},
            'Italian'             => {'Barbara'  => '{7EF71700-A92D-11d1-B17B-0020AFED142E}',
                                      'Stefano'  => '{7EF71701-A92D-11d1-B17B-0020AFED142E}'},
            'Japanese'            => {'Naoko'    => '{A778E060-A936-11d1-B17B-0020AFED142E}',
                                      'Kenji'    => '{A778E061-A936-11d1-B17B-0020AFED142E}'},
            'Korean'              => {'Shin-Ah'  => '{12E0B720-A936-11d1-B17B-0020AFED142E}',
                                      'Jun-Ho'   => '{12E0B721-A936-11d1-B17B-0020AFED142E}'},
            'Portuguese (Brazil)' => {'Juliana'  => '{8AA08CA0-A1AE-11d3-9BC5-00A0C967A2D1}',
                                      'Alexandre'=> '{8AA08CA1-A1AE-11d3-9BC5-00A0C967A2D1}'},
            'Russian'             => {'Svetlana' => '{06377F80-D48E-11d1-B17B-0020AFED142E}',
                                      'Boris'    => '{06377F81-D48E-11d1-B17B-0020AFED142E}'},
            'Spanish'             => {'Carmen'   => '{2CE326E0-A935-11d1-B17B-0020AFED142E}',
                                      'Julio'    => '{2CE326E1-A935-11d1-B17B-0020AFED142E}'}};

our $Languages = {'Arabic'                => 0x0401,
                'Basque'                  => 0x042D,
                'Chinese (Simplified)'    => 0x0804,
                'Chinese (Traditional)'   => 0x0404,
                'Croatian'                => 0x041A,
                'Czech'                   => 0x0405,
                'Danish'                  => 0x0406,
                'Dutch'                   => 0x0413,
                'English (British)'       => 0x0809,
                'English (US)'            => 0x0409,
                'Finnish'                 => 0x040B,
                'French'                  => 0x040C,
                'German'                  => 0x0407,
                'Greek'                   => 0x0408,
                'Hebrew'                  => 0x040D,
                'Hungarian'               => 0x040E,
                'Italian'                 => 0x0410,
                'Japanese'                => 0x0411,
                'Korean'                  => 0x0412,
                'Norwegian'               => 0x0414,
                'Polish'                  => 0x0415,
                'Portuguese (Portugal)'   => 0x0816,
                'Portuguese (Brazil)'     => 0x0416,
                'Romanian'                => 0x0418,
                'Russian'                 => 0x0419,
                'Slovakian'               => 0x041B,
                'Slovenian'               => 0x0424,
                'Spanish'                 => 0x0C0A,
                'Swedish'                 => 0x041D,
                'Thai'                    => 0x041E,
                'Turkish'                 => 0x041F};


sub new
{
    my $proto = shift;
    my $agent = shift;
    my $self = {};
    my $class = ref($proto) || $proto;
    bless $self, $class;
    $self->{_Characters} = $agent->Characters(@_);
    return $self;
}

sub Load
{
    my $self = shift;
    $self->{_Characters}->Load(@_);
}

sub Show
{
    my $self = shift;
    $self->{_Characters}->Show();
}

sub Hide
{
    my $self = shift;
    $self->{_Characters}->Hide();
}

sub Play
{
    my $self = shift;
    $self->{_Characters}->Play(@_);
}

sub Interrupt
{
    my $self = shift;
    $self->{_Characters}->Interrupt(@_);
}

sub Stop
{
    my $self = shift;
    $self->{_Characters}->Stop(@_);
}

sub StopAll
{
    my $self = shift;
    $self->{_Characters}->StopAll(@_);
}

sub Activate
{
    my $self = shift;
    $self->{_Characters}->Activate(@_);
}

sub GestureAt
{
    my $self = shift;
    $self->{_Characters}->GestureAt(@_);
}

sub Get
{
    my $self = shift;
    $self->{_Characters}->Get(@_);
}

sub Listen
{
    my $self = shift;
    my $bool = shift;
    $self->{_Characters}->Listen(Variant(VT_BOOL, $bool));
}

sub MoveTo
{
    my $self = shift;
    $self->{_Characters}->MoveTo(@_);
}

sub ShowPopupMenu
{
    my $self = shift;
    $self->{_Characters}->ShowPopupMenu(@_);
}

sub Wait
{
    my $self = shift;
    $self->{_Characters}->Wait(@_);
}

sub Think
{
    my $self = shift;
    $self->{_Characters}->Think(@_);
}

sub Speak
{
    my $self = shift;
    $self->{_Characters}->Speak(@_);
}

sub AnimationNames
{
    my $self = shift;
    my $olenum = $self->{_Characters}->AnimationNames();
    my $names = Win32::OLE::Enum->new($olenum);
    return $names->All();
}

sub Voice
{
    my $self = shift;
    my $language = shift;
    my $voice = shift;
    $self->{_Characters}->{LanguageID} = $Languages->{$language};
    $self->{_Characters}->{TTSModeID} = $Voices->{$language}->{$voice};
}

sub Active
{
    my $self = shift;
    my $bool = shift;
    return $self->{_Characters}->{Active} = Variant(VT_BOOL, $bool);
}

sub AutoPopupMenu
{
    my $self = shift;
    my $bool = shift;
    return $self->{_Characters}->{AutoPopupMenu} = Variant(VT_BOOL, $bool);
}

sub Description
{
    my $self = shift;
    my $param = shift;
    $self->{_Characters}->{Description} = $param if $param;
    return $self->{_Characters}->{Description};
}

sub Visible
{
    my $self = shift;
    return $self->{_Characters}->{Visible};
}

sub Speed
{
    my $self = shift;
    return $self->{_Characters}->{Speed};
}

sub Pitch
{
    my $self = shift;
    return $self->{_Characters}->{Pitch};
}

sub LanguageID
{
    my $self = shift;
    my $language = shift;
    $self->{_Characters}->{LanguageID} = $Languages->{$language} if $language;
    return $self->{_Characters}->{LanguageID};
}

sub TTSModeID
{
    my $self = shift;
    my $TTSModeID = shift;
    $self->{_Characters}->{TTSModeID} = $Voices->{$self->LanguageID}->{$TTSModeID} if $TTSModeID;
    return $self->{_Characters}->{TTSModeID};
}

sub Height
{
    my $self = shift;
    my $height = shift;
    $self->{_Characters}->{Height} = $height if $height;
    return $self->{_Characters}->{Height};
}

sub Width
{
    my $self = shift;
    my $width = shift;
    $self->{_Characters}->{Width} = $width if $width;
    return $self->{_Characters}->{Width};
}

sub HelpModeOn
{
    my $self = shift;
    my $bool = shift;
    return $self->{_Characters}->{HelpModeOn} = Variant(VT_BOOL, $bool);
}

sub SoundEffectsOn
{
    my $self = shift;
    my $bool = shift;
    return $self->{_Characters}->{SoundEffectsOn} = Variant(VT_BOOL, $bool);
}


########################################### main pod documentation begin ##
# Below is the stub of documentation for your module. You better edit it!

=head1 NAME

Win32::MSAgent - 

=head1 SYNOPSIS

    use Win32::MSAgent;
    my $agent = Win32::MSAgent->new('Genie');

    my $char = $agent->Characters('Genie');
    $char->Voice('Dutch', 'Alexander');
    $char->SoundEffectsOn(1);
    $char->Show();

    $char->MoveTo(300,300);
    sleep(5);

    foreach my $animation ($char->AnimationNames)
    {
        my $request = $char->Play($animation);
        $char->Speak($animation);
        my $i = 0;
        while (($request->Status == 2) || ($request->Status == 4))
        { $char->Stop($request) if $i >10; sleep(1);  $i++}
    }

=head1 DESCRIPTION

Win32::MSAgent allows you to use the Microsoft Agent 2.0 OLE control in your
perl scripts. From the Microsoft Website: "With the Microsoft Agent set of 
software services, developers can easily enhance the user interface of their 
applications and Web pages with interactive personalities in the form of animated 
characters. These characters can move freely within the computer display, speak 
aloud (and by displaying text onscreen), and even listen for spoken voice commands."

Since the MS Agent itself is only available on MS Windows platforms, this module
will only work on those.

=head1 PREREQUISITES

In order to use the MSAgent in your scripts, you need to download and install some
components. They can all be downloaded for free from http://www.microsoft.com/msagent/devdownloads.htm

=over 4

=item 1. Microsoft Agent Core Components

=item 2. Localized Language Components 

=item 3. MS Agent Character files (.acs files)

=item 4. Text To Speech engine for your language

=item 5. SAPI 4.0 runtime

=back

Optionally you can install the Speech Recognition Engines and the Speech Control Panel. The Speech
Recognition part of MS Agent is not supported in this version of Win32::MSAgent

=head1 USAGE

There is no POD documentation on the methods and stuff yet. I will write this for upcoming versions.
In this version the Agent control and the Characters object are implemented. No events are implemented yet,
nor any of the Speech Recognition stuff.
Please take the MS Agent Platform SDK helpfile (download from http://www.microsoft.com/msagent/devdownloads.htm)
for documentation on the various methods and properties.

=head1 BUGS

Hey! Give me some slack! This is only version 0.01! Sure there are bugs!

=head1 SUPPORT

The MS Agent itself is supported on MS' public newsgroup news://microsoft.public.msagent
You can email the author for support on this module.

=head1 AUTHOR

	Jouke Visser
	jouke@cpan.org
	http://jouke.pvoice.org

=head1 COPYRIGHT

Copyright (c) 2002 Jouke Visser. All rights reserved.
This program is free software; you can redistribute
it and/or modify it under the same terms as Perl itself.

The full text of the license can be found in the
LICENSE file included with this module.

=head1 SEE ALSO

perl(1).

"Yet Another True Value";
__END__
