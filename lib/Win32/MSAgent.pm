package Win32::MSAgent;
use strict;
use warnings;
use Win32::OLE;
use Win32::OLE::Variant;

BEGIN {
	use Exporter ();
	use vars qw ($VERSION @ISA @EXPORT @EXPORT_OK %EXPORT_TAGS);
	$VERSION     = 0.02;
	@ISA         = qw ();
	#Give a hoot don't pollute, do not export more than needed by default
	@EXPORT      = qw ();
	@EXPORT_OK   = qw ();
	%EXPORT_TAGS = ();
    use vars qw { $Agent } ;
    Win32::OLE->Initialize(Win32::OLE::COINIT_MULTITHREADED);
    $Agent = Win32::OLE->new('Agent.Control.2') || die Win32::OLE->LastError();
    $Agent->{Connected} = 1;
}


sub new
{
    my $proto = shift;
    my $char = shift;
    my $self = {};
    my $class = ref($proto) || $proto;
    bless $self, $class;
    if ($char)
    {
        $self->Characters->Load($char, "$char.acs");
    }

    return $self;
}


sub ShowDefaultCharacterProperties
{
    my $self = shift;
    $Agent->ShowDefaultCharacterProperties(@_);
}


sub Connected
{
    my $self = shift;
    my $bool = shift;
    $Agent->{Connected} = Variant(VT_BOOL, $bool);
    return $Agent->{Connected}
}


sub RaiseRequestErrors
{
    my $self = shift;
    my $bool = shift;
    $Agent->{RaiseRequestErrors} = Variant(VT_BOOL, $bool);
    return $Agent->{RaiseRequestErrors}
}

sub Characters
{
    my $self = shift;
    $self->{_Characters} = Win32::MSAgent::Characters->new($Agent, @_) 
        if ((@_) or (not exists $self->{_Characters}));
    return $self->{_Characters};
}

package Win32::MSAgent::Characters;

use strict;
use warnings;
use Win32::OLE;
use Win32::OLE::Variant;
use Win32::OLE::Enum;

our $Voices ={'English (US)'        => {'Adult Female #1' => '{CA141FD0-AC7F-11D1-97A3-006008273008}', 
                                      'Adult Female #2' => '{CA141FD0-AC7F-11D1-97A3-006008273009}', 
                                      'Adult Male #1'   => '{CA141FD0-AC7F-11D1-97A3-006008273000}',
                                      'Adult Male #2'   => '{CA141FD0-AC7F-11D1-97A3-006008273001}',
                                      'Adult Male #3'   => '{CA141FD0-AC7F-11D1-97A3-006008273002}',
                                      'Adult Male #4'   => '{CA141FD0-AC7F-11D1-97A3-006008273003}',
                                      'Adult Male #5'   => '{CA141FD0-AC7F-11D1-97A3-006008273004}',
                                      'Adult Male #6'   => '{CA141FD0-AC7F-11D1-97A3-006008273005}',
                                      'Adult Male #7'   => '{CA141FD0-AC7F-11D1-97A3-006008273006}',
                                      'Adult Male #8'   => '{CA141FD0-AC7F-11D1-97A3-006008273007}'},
            'English (British)'     => {'Carol'    => '{227A0E40-A92A-11d1-B17B-0020AFED142E}',
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

sub ListVoices
{
    my $self = shift;
    my $language = shift;
    return if not $language;
    return keys %{$Voices->{$language}}
}

sub ListLanguages
{
    my $self = shift;
    return keys %{$Languages};
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


=pod

=head1 NAME

Win32::MSAgent - Interface module for the Microsoft Agent

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

The POD documentation here is not complete. I will continue to write more documentation in upcoming versions.
In this version the Agent control and the Characters object are implemented. No events are implemented yet,
nor any of the Speech Recognition stuff.

=head2 $agent = Win32::MSAgent->new([charactername])

The constructor optionally takes the name of the character to load. It 
loads the MS Agent OLE control, connects to it, and if the character name
is supplied, it loads that character in the Characters object already.
It returns the Win32::MSAgent object itself.


=head2 $agent->ShowDefaultCharacterProperties([X,Y])

Calling this method displays the default character properties window 
(not the Microsoft Agent property sheet). If you do not specify the X 
and Y coordinates, the window appears at the last location it was displayed


=head2 $agent->Connected(boolean)

Returns or sets whether the current object is connected to the Microsoft Agent server.


=head2 $agent->RaiseRequestErrors(boolean)

This method enables you to determine whether the server raises errors that occur 
with methods that support Request objects. For example, if you specify an animation 
name that does not exist in a Play method, the server raises an error (displaying 
the error message) unless you set this property to False. 


=head2 $char = $agent->Characters(character)

This method returns the character object for the character you specify. 
If you don't specify a charactername, it returns the object for the 
character you previously 'Load'-ed. 

=head2 $agent->Characters->ListVoices(languagename)

This method can be called on a Character object. It lists all possible voices for a certain
language. It does not list voices nor languages that are installed on your system! Just the
ones that you could install and use with MS Agent. You can query the correct id for a language
from ListLanguages

=head2 $agent->Characters->ListLanguages

This method can be called on a Character object. It lists all possible languages you
could use with MS Agent. Again, these are not the installed languages, just the ones MS Agent
could recognize.

=head2 $agent->Characters->Load(id, acsfile)

You can load characters from the Agent subdirectory by specifying a relative path 
(one that does not include a colon or leading slash character). This prefixes the 
path with Agent's characters directory (located in the localized Windows\msagent 
directory). For example, specifying the following would load Genie.acs from Agent's 
Chars directory:

   $agent->Characters->Load("genie", "genie.acs")

You can also specify your own directory in Agent's Chars directory.

   $agent->Characters->Load("genie", "MyCharacters\\genie.acs")

=head2 $char->Show

This method shows the character on whose object it is called

=head2 $char->Hide

This method hides the character on whose object it is called

=head2 $char->Play(animation)

This method lets the character do the animation you tell it to.
You can get a list of all animations that a certain character supports
by calling AnimationNames

=head2 $char->Interrupt(request)

The Play, Speak, Move, and other 'animation' methods return request objects.
If you want to interrupt the animation, call this method on the character
with the returned request object from the action you want to interrupt.

=head2 $char->Stop(request)

Like the Interrupt method, this does not only interrupt the action, it 
completely stops it. The author of this module does not completely see the
difference between the two, since there is no 'Resume' method to call on
an interrupted action.
With $request->Status however you can find out _why_ an action is stopped...
See below for more info on the Request object.

=head2 $char->StopAll([type])

If you call this method without any parameters, it stops every action for this
character. You can however supply a 'type' parameter, indicating which types
of actions should be stopped:

(You can also specify multiple types by separating them with commas. )

=over 4

=item "Get"

To stop all queued Get requests.

=item "NonQueuedGet"

To stop all non-queued Get requests (Get method with Queue parameter set to False).

=item "Move"

To stop all queued MoveTo requests.

=item "Play" 

To stop all queued Play requests.

=item "Speak" 

To stop all queued Speak requests.

=back

=head2 Activate

=head2 GestureAt

=head2 Get

=head2 Listen

=head2 MoveTo

=head2 ShowPopupMenu

=head2 Wait

=head2 Think

=head2 Speak

=head2 @animations = $char->AnimationNames()

=head2 Voice

=head2 Active

=head2 AutoPopupMenu

=head2 Visible

=head2 Speed

=head2 Pitch

=head2 LanguageID

=head2 TTSModeID

=head2 Height

=head2 Width

=head2 HelpModeOn

=head2 SoundEffectsOn

Please take the MS Agent Platform SDK helpfile (download from http://www.microsoft.com/msagent/devdownloads.htm)
for more documentation on the various methods and properties.

=head1 The Request object

yadayadayada

=head1 BUGS

Hey! Give me some slack! This is only version 0.02! Sure there are bugs!

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

=cut

"Yet Another True Value";

__END__
