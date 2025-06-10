PROJECT STATUS WARNING!!!
-------------------------
THIS IS AN ANCIENT PROJECT AND IS UNMAINTAINED! I'm uploading it for its archival value since it was a relative hit a couple of decades ago, when controlling your home computer (especially games) with your speech was bordering science fiction.

[![IMAGE ALT TEXT HERE](https://img.youtube.com/vi/cfQu2nuDt8A/0.jpg)](https://www.youtube.com/watch?v=cfQu2nuDt8A)

PiLfIuS!: Speech recognition with Gamers in mind
`'`'`'`''`'`'`'`'`'`'`'`'`'`'`'`'`'`'`''`'`'`'`'

INDEX
-----
1. What is PiLfIuS?
2. Installation
3. Quick start
4. Command-line arguments
5. Known bugs and limitations
6. Author, notes, support and contact info
7. Legal stuff
8. Acknowledgements

1. What is PiLfIuS?
--------------------

PiLfIuS! is a FREE speech recognition and command application specifically written with gamers in mind. It will identify whatever commands you speak into your microphone, and send keystrokes accordingly. Commands, and their respective keystrokes are fully configurable by the user.

Unlike many other speech recognition applications, keystrokes sent by PiLfIuS! should be recognized by ANY application, since these programs will not tell the difference between a keystroke sent by your keyboard, and a keystroke sent as a result of your voice commands.



2. Installation (Probably won't work on anything newer than Windows XP)
-----------------

- Download and install a Speech Recognition Engine for Windows XP (SAPI5) if you haven't already.

- To install PiLfIuS! just run the included EXE file and follow instructions.



3. Quick start
--------------

To get things going ASAP, follow these instructions:

- Run PiLfIuS!

- Click on CREATE NEW COMMAND LIST

- To add Command Groups, click on the "+" button in the "Commands Group" frame. Enter its name, and click "+" to add it.
Command groups can be used to better organize commands, grouping them by their general function (i.e.: commands "Move North" and "Move South" can be grouped as "Movement Commands", whilst "Engage at will" and "Hold your fire" can be grouped as "Fire Commands").

- To add voice commands, click the "+" in the Voice Commands List. Type the voice command, select which group you want it to be assigned to, and click "+" to add it. This is the command that will be recognized when you say it in your microphone.

- The Keystrokes panel should now slide into view.

- To assign keystrokes to the voice command you just created, select the textbox in the "Add keystrokes to command" frame. Enter a key combination (you should see it appear on the textbox), and click the "+" at the bottom of this frame to add a new keystroke. You may add as many keystrokes as you wish.
These are the keystrokes that will be "pressed" when you say the command you just created. Notice that their order does matter.

- THAT'S IT. You can now load your favorite game.
When you speak the new command into your microphone, the keystrokes you configured will be sent.

- Don't forget to SAVE, with the buttons at the bottom left of the screen.

NOTE: Obviously, PiLfIuS! must be running in order to identify the voice commands. You may minimize it, but don't close it. If you close the Command List, speech recognition will stop.



4. Command-line arguments
-------------------------

-list:"<commandlist path>"
Loads a PiLfIuS! Command List on startup. If only a filename is specified, with no path, it will search the folder where PiLfIuS! is installed for that file.
Example: pilfius.exe -list:"c:\Program Files\PiLfIuS\mylist.lcl"



5. Known bugs and limitations
-----------------------------

- Animations produce artifacts.

- If PUSH-TO-ACTIVATE key combination has a Shift, Alt and/or Ctrl in it, PiLfIuS! will stop listening after a command has been recognized even if you didn't release the PTA keys. You'll have to release and press them again for the program to start listening again.



6. Author, notes and contact info
---------------------------------

PiLfIuS! was written and designed by Diego Wasser, a.k.a. Razorwings18.
Website and support forums: http://www.pilfius.com.ar/

Funny note I wrote back then; no longer applies: "PiLfIuS! is Donationware, which means you can use it at its full extent for as long as you like. However, users who have not donated will get a guilt pop-up in their conscience each time the application is run. Donating will effectively remove this feature."
I'm now on welfare and have dibs on any garbage bags my local McDonald's throws away at night, so no need for donations anymore.



7. Legal stuff
--------------

PiLfIuS! is released under the Creative Commons Attribution-NonCommercial-NoDerivs 3.0 Unported license. Details at http://creativecommons.org/licenses/by-nc-nd/3.0/ . Copyright (c) 2007 Diego Wasser.

In short, that means you may freely distribute this application for non-commercial use, but this readme file MUST be included, unmodified, along with it. PiLfIuS! shall not be included in any commercial distribution (even as a "gift" accompanying a commercial product) without the express consent of the author.

I will not be held liable for any damage or problem usage of this software may cause. This application is offered as-is, with no warranties or guarantees. Use at your own risk.


Oh. And PiLfIuS! does not contain spyware in any shape or form. It's a shame I even have to mention that.



8. Acknowledgements
-------------------

Skularach, for testing and feedback in the closed beta for v0.7
Bashar711, carlo maker and GinSoakedBoy for their extensive testing and suggestions in the closed beta for v0.9
Everyone who's posted bug reports, suggestions and comments at the forums (yeah, this thing had an active forum even).

PiLfIuS! has become a better application thanks to you; the users and I appreciate your involvement in the project.