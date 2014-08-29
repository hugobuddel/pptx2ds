pptx2ds
=======

pptx2ds.ps1 is a PowerShell script to convert PowerPoint presentations to DigitalSky scripts.

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
