pptx2ds
=======

pptx2ds.ps1 is a PowerShell script to convert PowerPoint presentations to DigitalSky scripts.

The goal of the program is to make it as easy as possible for presenters to create presentations that make use of the dome. Ease of use is more important than features: converting a presentation should only take a minute with no (or only a few) manual steps.

Together with other solutions this software can be used provide different level of presentation support in Digital Sky.

1. Simple: through streaming or separate projector.
2. Intermediate: some dome features through software like this.
3. Full: hand-tailored DigitalSky script.
4. (Extreme: J-Walt.)

Currently the software is rather simple: it takes a .pptx file, converts it to images and creates a 'button' to display 3 slides at the same time. The slides can either flip together or one after the other. Slide animations and flow control are not supported.

It is written in PowerShell, because 1) this should work on any windows machine out of the box and 2) it can automatically work with Office. It has not yet been tested on a machine that has both Office and DigitalSky, but this should work fine.

Usage
=====
First time Usage:

1. Create the "Slides" button set in DigitalSky.
   - "File" menu -> "Preferences" item -> "Scripts" tab -> "Add New Folder" button -> Type "Slides"-> "OK"button.
2. Change the paths in "General Settings" to match your local installation.
   - Ensure that the directories exist, in particular the Shows/Slides one.
3. Ensure that you can run PowerShell scripts.
   - E.g. run: `Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser`
   - See http://technet.microsoft.com/nl-NL/library/hh847748.aspx

Usage per presentation.

1. Place the presentation in the $path_powerpoint directory.
2. Change the "Presentation Settings".
3. Run the script.
4. Reload the "Slides" button set in DigitalSky. E.g. by pressing F5.
5. Start the presentation by pressing the corresponding button.

Improvements
============

Feel free to create issues with proposals for future functionality or even fork it. Possible improvements are:
- Better flow control.
- Better animations between/within slides.
- Movie support.
- 3D support.
- Easier way to handle settings.
- Automatic batch mode.
- etc.

Related
=======
See the PowerPoint dome templates by Andrew Hazelden: http://www.andrewhazelden.com/blog/2013/06/powerpoint-dome-template/ (Which unfortunately cannot be used with pptx2ds yet.)
