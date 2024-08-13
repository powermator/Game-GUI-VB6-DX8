DirectX 7 Direct Sound Module for Visual Basic
By, Jonathan Enzinna
    Lazer Light Ltd.
    enzinna@hotpop.com

INSTRUCTIONS:
Add the module to your project and add the following lines to the form(s) that are going to use Direct Sound:

cSoundBuffers=10 			'Select the number of buffers you wish to use
SetupDX7Sound Me              		
CreateBuffers cSoundBuffers, App.Path+"\Default.wav"

PLEASE NOTE: The default wave file should include the full path and should be blank and very short to make the program run smoother.

Then you are all set to use the following functions to play wave files in PCM format (no compression):

PLAY SOUND:

DX7LoadSound(Buffer as Integer, sfile as String)

	Buffer = Sound Buffer to use
	sfile  = Full path to the wave file


PLAY SOUND IN NON-CENTER CHANNEL:

PlaySoundWithPan(Buffer as Integer, sfile as String, Optional Volume as Integer, Opyional 
		 PanValue as Byte)

	Buffer and sfile see above
	Volume   = volume level (0 (mute) to 100 (full))
	PanValue = pan value (0 (left) to 100 (right) where 50=center)


PANNING A SOUND WHILE IT IS STILL PLAYING:

PanSound(Buffer as Integer, PanValue as Byte)

	Buffer and PanValue see above


SETTING VOLUME WHILE A SOUND IS STILL PLAYING:

VolumeLevel(Buffer as Integer, Volume as Integer)

	Buffer and Volume see above


DETERMINING WHETHER A BUFFER IS PLAYING A SOUND OR NOT:

Function IsPlaying(Buffer as Integet) as Long

	Passes a true or false according to if the buffer (stated in Buffer) is playing.




HELP:

If you need help, feel free to e-mail me at enzinna@hotpop.com