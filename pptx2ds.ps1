# Script to create DigitalSky buttons from PowerPoint presentations.
#
# Creator: Hugo Buddelmeijer <hugo@buddelmeijer.nl>
# Creation Date: 2014-08-29
# Version: 0.1.0
#
# First time Usage:
# 1) Create the "Slides" button set in DigitalSky.
#    - "File" menu -> "Preferences" item -> "Scripts" tab
#       -> "Add New Folder" button -> Type "Slides"-> "OK"button.
# 2) Change the paths in "General Settings" to match your local installation.
#    Ensure that the directories exist, in particular the Shows/Slides one.
# 3) Ensure that you can run PowerShell scripts. E.g. run:
#        Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser                                                               
#    See http://technet.microsoft.com/nl-NL/library/hh847748.aspx
#
# Usage per presentation.
# 1) Place the presentation in the $path_powerpoint directory.
# 2) Change the "Presentation Settings".
# 3) Run the script.
# 4) Reload the "Slides" button set in DigitalSky. E.g. by pressing F5.
# 5) Start the presentation by pressing the corresponding button.
#



#####
# Presentation Settings
#

$filename_powerpoint = "testpresentation.pptx"

$script_number = "55"

$flipmode = "alltogether"
#$flipmode = "onebyone"



#####
# General Settings
#

# Change these paths to your local installation.
[string] $pwd_real = Convert-Path ( Get-Location )
$path_powerpoint = $pwd_real + "\Powerpoint\"
$path_digitalsky = $pwd_real + "\DigitalSky\"

# Create a Slides button set and Shows directory first.
$path_buttons = $path_digitalsky + "Buttons\Slides\"
$path_shows = $path_digitalsky + "Shows\Slides\"

#$path_shows_script = $path_shows
$path_shows_script = "C:\DigitalSky\Shows\Slides\"


#####
# Automatically Generated Configuration.
#

$width_slide = 30

$height_slide = ($width_slide * 3) /4

$filename_button = "F" + $script_number + ".sct"



#####
# Generating Images From Powerpoint
#
# Based on
# http://tfl09.blogspot.nl/2012/11/saving-powerpoint-slides-to-pdf-with.html
# 

# The full location of the Powerpoint file.
$fullfilename_powerpoint = $path_powerpoint + $filename_powerpoint

Write-Host "Opening PowerPoint presentation."
Write-Host $fullfilename_powerpoint 

# The directory where the slides will be saved is based on the number of the button.
$fullfilename_png = $path_shows + $script_number

# Is edding these assemblies necessary?
Add-type -AssemblyName office -ErrorAction SilentlyContinue 
Add-Type -AssemblyName microsoft.office.interop.powerpoint -ErrorAction SilentlyContinue

# Starting powerpoint.
[object] $powerpoint = New-Object -ComObject PowerPoint.Application
[object] $presentation = $powerpoint.Presentations.Open($fullfilename_powerpoint)

# Get the number of slides, which is used in the creation of the DS script.
$number_of_slides = $presentation.Slides.Count

# Convert the PowerPoint presentation to PNGs
Write-Host "Converting presentation to PNG."
$opt= [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPNG
$presentation.SaveAs($fullfilename_png, $opt)



#####
# Creating the DigitalSky button.
#

# The full location of the Button file.
$fullfilename_button = $path_buttons + $filename_button
$filename = $fullfilename_button

# The files that are created by PowerPoint.
$filenamebase = $path_shows_script + $script_number + "\Slide"

Write-Host "Creating script" $script_number "for" $number_of_slides "slides."

# Script Header.
';----------------------------------------------------------------------' > $filename
'; Script number = F' + $script_number >> $filename
'; Title         = "testpres"' >> $filename
'; Color         = "255,0,0"' >> $filename
'; Created on    : 2014-08-27 09:46 Z' >> $filename
'; Modified      : 2014-08-27 09:47 Z' >> $filename
'; Version       : 2.3' >> $filename
'; Created by    : "hugo@buddelmeijer.nl"' >> $filename
'; Keywords      : ' >> $filename
'; Description   : ' >> $filename
';----------------------------------------------------------------------' >> $filename
';' >> $filename

# Load all the slides.
for ($s = 1; $s -le $number_of_slides; $s++) {
    #Write-Host $s
    $t = 'Text Add "slide' + $s + '" "' + $filenamebase + $s + '.PNG" 0   0      0       0          0     0       0     0' 
    $t >>  $filename

}

# Flip all three slides at the same time.
if ($flipmode -eq "alltogether") {
    $slides_done = 0
    do {

        #Write-Host $slides_done 

        '+.1    Text Locate "slide' + ($slides_done + 1) + '"   0       -60    45   0  ' + $width_slide + ' ' + $height_slide >> $filename
        '       Text Locate "slide' + ($slides_done + 2) + '"   0         0    45   0  ' + $width_slide + ' ' + $height_slide >> $filename
        '       Text Locate "slide' + ($slides_done + 3) + '"   0        60    45   0  ' + $width_slide + ' ' + $height_slide >> $filename
        '       Text View   "slide' + ($slides_done + 1) + '"   3       100   100 100 100' >> $filename
        '       Text View   "slide' + ($slides_done + 2) + '"   3       100   100 100 100' >> $filename
        '       Text View  "slide' + ($slides_done + 3) + '"    3       100   100 100 100' >> $filename
        'ButtonText  "Next Slides"' >> $filename
        'STOP' >> $filename
        '       Text View  "slide' + ($slides_done + 1) + '"     2       0     0   0   0' >> $filename
        '       Text View  "slide' + ($slides_done + 2) + '"     2       0     0   0   0' >> $filename
        '       Text View  "slide' + ($slides_done + 3) + '"     2       0     0   0   0' >> $filename

        $slides_done += 3
    } while ($slides_done -lt $number_of_slides)
}

# Flip slides one by one.
if ($flipmode -eq "onebyone") {

    for ($s = 1; $s -le $number_of_slides; $s++) {
        $az = -60 + ( ( $s - 1 ) % 3) * 60

        '       Text Locate "slide' + $s + '"   0       ' + $az + '    45   0  ' + $width_slide + ' ' + $height_slide >> $filename
        '       Text View   "slide' + $s + '"   1       100   100 100 100' >> $filename
        'ButtonText  "Next Slides"' >> $filename
        'STOP' >> $filename
    #    '       Text View  "slide' + ($slides_done + 1) + '"     2       0     0   0   0' >> $filename

    }
}

# Remove all slides.
for ($s = 1; $s -le $number_of_slides; $s++) {
    #Write-Host $s
    $t = 'Text Remove "slide' + $s
    $t >> $filename
}
