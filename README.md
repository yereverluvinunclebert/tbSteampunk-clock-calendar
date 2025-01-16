# Steampunk Clock Calendar-TwinBasic
 
A FOSS Desktop Utility TwinBasic WoW64 Widget for Windows XP, ReactOS, Vista, Win7, 8 and 10/11+.

My current TwinBasic/RC6 PSD program being worked upon now, in progress, you can download it but do expect some functionality to be incomplete and unpolished, especially as it is using a Beta-grade Compiler (TwinBasic). Estimated at 99% graphically complete, 92% functionally complete but only 82% code-complete, mostly due to a need for an efficiency-rewrite and removal of some TwinBasic workarounds that need to be fixed properly in TwinBasic. I am always adding in new functionality as this is an educational project for me, there are code improvements to be made, alternative more efficient methods found &c.

What does the program do already in its unfinished state?

	* Chimes the hours and quarters.
	* Has up to five working alarms. 
	* All the steampunk controls working and functioning as designed, responds to keyboard and mouse events.
 	* The time slider now advances time using a logarthmic scale.
 	* Alarms can be set via the time slider.  
	* Has a fully functioning preference utility.
	* Has full documentation and help.
	* Demonstrates the use of TwinBasic and Cairo Graphics working together.
	* Demonstrates how to put a transparent form on your desktop using TwinBasic.

When finished, this Steampunk Clock Calendar widget will be an attractive steampunk widget for your desktop. It is a deliberately-complex clock with a skeumorphic interface. Functional but gorgeous at the same time. This VB6/TwinBasic Widget is a moveable widget that you can move anywhere around the desktop as you require. The design is based upon the Yahoo widget of the same design which I also designed. This is its replacement.

TwinBASIC bugs:

There are a few bugs specific to TwinBasic that need to be rectified by Wayne Phillips in TwinBasic itself, that directly affect the Steampunk Clock/calendar.

	* BorderStyle value on Frames requires a .Refresh https://github.com/twinbasic/twinbasic/issues/1970	
 	* Image JPGs loaded from RC collection upside down https://github.com/twinbasic/twinbasic/issues/1939
 	* Image control - mouseUp event not activated when the control is made invisible https://github.com/twinbasic/twinbasic/issues/1938
   	* TwinBasic is slower than VB6 when populating ComboBox controls https://github.com/twinbasic/twinbasic/issues/1968

These bugs will no doubt be fixed before TwinBasic leaves Beta. They are just worth mentioning.
 
If you just want to install the Steampunk Clock Calendar on your desktop, there will be a clockCalendarInstaller.exe in the release section that you download and run. If you want to compile this yourself, clone/download the lot. You will of course, need the TwinBasic IDE to make changes which is downloadable and available for free. If you have VB6 already installed then you can use that, of course.

 Why TwinBasic? Well, with a 64 bit, modern-language upgrade improvement being delivered now, in the form of "100% compatible" TwinBasic, coupled with support for transparent PNGs via RC/Cairo, TwinBasic native code has an  amazing future. 

 ![vb6-logo-200](https://github.com/yereverluvinunclebert/Panzer-JustClock-VB6/assets/2788342/7986e544-0b94-4a10-90bb-2d9fb60c294a)
 
 I created this as a development from the original Yahoo widget/ Konfabulator version I had previously created for the steampunk 
 communities. This widget is an attractive steampunk TwinBasic widget for your desktop. It is almost all my code with some help from the chaps at VBForums (credits given).
 
 ![about-image001](https://github.com/yereverluvinunclebert/Steampunk-clock-calendar-vb6/assets/2788342/c6a5962d-ccc3-43ad-8316-607c122026ee)

 This widget can be increased in size, animation speed can be changed, 
 opacity/transparency may be set as to the users discretion. The widget can 
 also be made to hide for a pre-determined period.

 Right clicking will bring up a menu of options. Double-clicking on the widget will cause a personalised Windows application to 
 fire up. The first time you run it there will be no assigned function and so it 
 will state as such and then pop up the preferences so that you can enter the 
 command of your choice. The widget takes command line-style commands for 
 windows. Mouse hover over the widget and press CTRL+mousewheel up/down to resize. It works well on Windows Vista
 to Windows 11. There will be another version for ReactOS and Windows XP, not yet built, watch this space.

![steampunk_clock_calendar_mkii__2_9__rc_by_yereverluvinuncleber-d4l5xny](https://github.com/yereverluvinunclebert/Steampunk-clock-calendar-version-2.9/assets/2788342/f2dc5337-0c98-418c-9b68-2374ae7c4222)


 This widget is currently Beta-grade software, under development, not yet 
 ready to use on a production system - use at your own risk.

 This version was developed on Windows 10 64bit using 32 bit version of TwinBasic. Created as a FOSS 
 project creating a WoW64 widget for the desktop. 

 The tool has two modes, Clock mode and Alarm mode. In clock mode the clock ticks, the calendar shows the date.
In alarm mode you can set alarms and when the time has passed the alarm will sound.

Instructions for use:
 
![lookatme](https://github.com/yereverluvinunclebert/Steampunk-clock-calendar-version-2.9/assets/2788342/d8878f9f-a95a-46f0-8fad-d3cf2573aa1a)

 It is open source to allow easy configuration, bug-fixing, enhancement and 
 community contribution towards free-and-useful TwinBasic utilities that can be created
 by anyone. The first step was the creation of a template program to form the 
 basis for the conversion of other desktop utilities or widgets. A future step 
 are new TwinBasic widgets with more functionality and then hopefully, conversion of 
 each to RADBasic/TwinBasic for future-proofing and 64bit-ness. 

![wotw-clock-help-image](https://github.com/yereverluvinunclebert/Steampunk-clock-calendar-version-2.9/assets/2788342/00887907-e663-448a-b322-7d6584d95512)

 By the left of the calendar are five brass toggles/keys. Pressing on each will have the following effect:

H Key - will show the first help canvas indicated by the brass number 1 on the top left of the wooden
bar. clicking on the brassnumber 1 will select the next drop down help canvas.
Clicking on the ring pull at the bottom will make the current canvas go away.

A Key - will activate the alarm mode and will also show the help canvas the first time
it is pressed. Click on the ring pull at the bottom to make the canvas go away
(f you do this note that it will still be in alarm mode). Clicking on the bell set will also cause
the clock to go into alarm mode.

When you have pressed the A key it will release the slider and you may move it to the right
or left and change time. When you have selected the date/time you want then move the slider
to the central position and click on the bell set. The alarm will set. You can set up to five alarms.

Alarm mode -  Normal operation is this: When the slider is released the further you move the slider from the
centre position the more quickly the date/time will change.

When you are ready to set the alarm, click the bellset, two bells will sound and the alarm is set.

* Please note that while the timepiece in Alarm Mode all clock functions are switched off *
* Alarms will not sound whilst in alarm mode *

To cancel an alarm setting or viewing operation just click on the clock face.
To cancel a ringing alarm - just click on the bellset.

Each time you press the alarm bell to set an alarm, a pop-up will display indicating
which alarm you are going to set. Each time you press the A key, it will select the next alarm.
To the right of the clock there are from zero to five alarm toggles depending on how many alarms you
have previously set. f you click on the toggle it will display the date and time set for this alarm.
If you then click on the associated 'cash-register-style' pop-up it will allow you to delete this alarm.

L Key - gently quietens the whole clock: ticking, chimes, alarm sounds all reduced by 21db.

M Key - leaves all other sounds alone but turns off only the chimes. Another click turns the
         chimes back on again. You will see the bell clapper move to/from the bell set.

P Key - Turns off the pendulum. Another click turns it on again. Single-click on the pendulum itself
         also turns off the pendulum.

Crank - The hand crank is the master volume control. Crank it down to mute all sounds and crank
         it up to restore the sound back to the level it was prior to muting.

To the left of the digital clock is another brass toggle:

S Key - Raises the transparent screen logging the various controls you select. Because the screen is
         transparent the text may be hard to see when used on a dark desktop background. A new 'B' key
         will also appear on the right of the screen frame that allows you to raise/lower the back screen.
         This will allow you to read the text.
	 
![wotw-clock-help-imageII](https://github.com/yereverluvinunclebert/Steampunk-clock-calendar-version-2.9/assets/2788342/ca4d4f68-ee8c-4d93-a684-3ee90907192a)

The screen currently only displays clock/calendar operations but may do more in the future.

At the back-end there are more preferences that may be changed, all are documented by an associated description.



 This utility is one of a set of steampunk and dieselpunk widgets. That you can 
 find here on Deviantart: https://www.deviantart.com/yereverluvinuncleber/gallery
 
 I do hope you enjoy using this utility and others. Your own software 
 enhancements and contributions will be gratefully received if you choose to 
 contribute.

 BUILD: The program runs without any Additional plugins.
 
 Built using: TwinBasic and RichClient 6 
 
 Links:
 
	https://twinbasic.com/
 	https://www.vbrichclient.com/#/en/About/
	MZ-TOOLS https://www.mztools.com/  
	CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62468&lngWId=1  
	Rubberduck http://rubberduckvba.com/  
	Rocketdock https://punklabs.com/  
	Registry code ALLAPI.COM  
	La Volpe http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1  
	PrivateExtractIcons code http://www.activevb.de/rubriken/  
	Subclassing code - Elroy  
	Open File common dialog code without dependent OCX - http://forums.codeguru.com/member.php?92278-rxbagain  
	VBAdvance  
 
 Tested on :
 
	Windows 7 Professional 32bit on Intel    
	Windows 7 Ultimate 64bit on Intel    
	Windows 7 Professional 64bit on Intel    
	Windows 10 Home 64bit on Intel    
	Windows 10 Home 64bit on AMD    
	Windows 11 64bit on Intel  
   
 ![tb6IDE-001](https://github.com/user-attachments/assets/97e1cc71-50c8-4625-b0b8-e5a0d9326413)
 The TwinBasic IDE displaying the Steampunk Clock/Calendar in code view on Windows 10.
 
 CREDITS:
 
 I have really tried to maintain the credits as the project has progressed. If I 
 have made a mistake and left someone out then do forgive me. I will make amends 
 if anyone points out my mistake in leaving someone out.
 
 MicroSoft in the 90s - MS built good, lean and useful tools in the late 90s and 
 early 2000s. Thanks for VB6 Microsoft, what a pity we can't download it anymore, 
 use TwinBasic instead...
 
 Olaf Schmidt - This tool was built using the RichClient RC6 Cairo wrapper for 
 VB6 that also works for TwinBasic. Specifically the components using transparency and reading images directly 
 from PSD. Thanks for the massive effort Olaf in creating Cairo counterparts for 
 all VB6 native controls and giving us access to advanced features on controls 
 such as transparency.
 
 Shuja Ali @ codeguru for his settings.ini code.
 
 ALLAPI.COM        For the registry reading code.
 
 Rxbagain on codeguru for his Open File common dialog code without a dependent 
 OCX - http://forums.codeguru.com/member.php?92278-rxbagain
 
 si_the_geek       for his special folder code
 
 Elroy on VB forums for the balloon tooltips and his essential subclassing code.
 
 Harry Whitfield for his quality testing, brain stimulation and being an 
 unwitting source of inspiration. 
 
 Dependencies:
 
 o A windows-alike o/s such as Windows Vista 7-11 or Apple Mac OSX 11. 
 
 o TwinBasic IDE installed from https://twinbasic.com/
 
 ![vb6-logo](https://github.com/yereverluvinunclebert/Panzer-JustClock-VB6/assets/2788342/861f1ce1-7058-4d09-af94-daab9206ca86)
 
 * Uses the latest version of the RC6 Cairo framework from Olaf Schmidt.
 
 During development the RC6 components need to be registered. These scripts are 
 used to register. Run each by double-clicking on them. These reside in the project's BIN folder.
 
	RegisterRC6inPlace.vbs
	RegisterRC6WidgetsInPlace.vbs
 
 During runtime on the users system, the RC6 components are dynamically 
 referenced using modRC6regfree.bas which is compiled into the binary.	
 
 Requires a SteampunkClockCalendar folder in C:\Users\<user>\AppData\Roaming\ 
 eg: C:\Users\<user>\AppData\Roaming\SteampunkClockCalendar
 Requires a settings.ini file to exist in C:\Users\<user>\AppData\Roaming\PzJustclock
 The above will be created automatically by the compiled program when run for the 
 first time.
 
o Krool's replacement for the Microsoft Windows Common Controls found in
mscomctl.ocx (slider) are replicated by the addition of one
dedicated OCX file that are shipped with this package.

The CCR slider should appear in the TwinBasic IDE toolbar automatically.
The OCX will reside in the program folder. The program reference 
to this OCX is contained within the supplied resource file, Steampunk Clock Calendar.RES. The reference 
to this file is already compiled into the binary. As long as the OCX is in the same 
folder as the binary the program will run without the need to register the OCX manually.

![clockPrefs](https://github.com/user-attachments/assets/63b56785-fbc3-4e71-9acc-cf714be80507)

One of the preference screens for this utility.


 * SETUP.EXE - The program is currently distributed using setup2go, a very useful 
 and comprehensive installer program that builds a .exe installer. You'll have to 
 find a copy of setup2go on the web as it is now abandonware. Contact me
 directly for a copy. The file "install steampunk-clock-calendar 0.1.0.s2g" is the configuration 
 file for setup2go. When you build it will report any errors in the build. Look in the releases
 folder for a release.
 
 * HELP.CHM - the program documentation is built using the NVU HTML editor and 
 compiled using the Microsoft supplied CHM builder tools (HTMLHelp Workshop) and 
 the HTM2CHM tool from Yaroslav Kirillov. Both are abandonware but still do
 the job admirably. The HTML files exist alongside the compiled CHM file in the 
 HELP folder.
 
 VB6 Project References in the IDE:

	TwinBasic VBA Compatibility Package 
	TwinBasic VBRun Compatibility Package 	
 	TwinBasic VBRun Compatibility Package (Forms) 
	RC6Widgets
 	RC6
  	CCRSlider.ocx
 	TwinBasic Common Controls Compatibility Package 
   

 LICENCE AGREEMENTS:
 
 Copyright © 2023 Dean Beedell
 
 In addition to the GNU General Public Licence please be aware that you may use 
 any of my own imagery in your own creations but commercially only with my 
 permission. In all other non-commercial cases I require a credit to the 
 original artist using my name or one of my pseudonyms and a link to my site. 
 With regard to the commercial use of incorporated images, permission and a 
 licence would need to be obtained from the original owner and creator, ie. me.

 ![wotw-clock-help-preview](https://github.com/yereverluvinunclebert/Steampunk-clock-calendar-version-2.9/assets/2788342/81d32fa2-5b79-4615-b31b-ce46c767ee87)

![desktop](https://github.com/yereverluvinunclebert/Panzer-JustClock-VB6/assets/2788342/8cf592a3-968f-4bf1-ab98-c734ff1cc261)


 
