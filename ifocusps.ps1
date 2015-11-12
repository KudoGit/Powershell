Param(
  [bool]$print = $false,
  [string]$subject,
  [int]$grade,
  [char]$gradec,
  [string]$type,
  [int]$lesson,
)

$USER = $Env:userprofile

$STAFFPATH   = $USER + "\Dropbox\Staff Print files\"

#Math worksheet directories
$CCSS        = "\CCSS "

$CCSSMATH    = "Math\CCSS Math\New_CC_8-25-2015-01-22 (eVER) STAMPED\Math\Gr."          
$CCML        = "M Lesson "

$SM          = "Math\SM\Text Books and Work Books (PRINT FROM HERE)\Easy Print Files (PRINT FROM HERE)\"

$FM          = "Math\FM\Focus Math (PRINT FROM HERE!!!)\Level "
$FM2         = "\FM-"

#Reading worksheet directories
$CCSSELA     = "ELA\CCSS Reading (STAMPLED)\New_CC_8-25-2015-01-22 (eVER) STAMPED\Reading - Gr."
$CCRL        = "R Lesson "

$LF          = "ELA\LF (Watermarked, Stamped)\LF"
$LF2         = " (Use this)\"

$VF          = "ELA\VF (Easy Print, Stamped, Watermarked)\Grade "
$VF2         = "Vocab Fundamentals - Grade "

$PH          = "ELA\PH\SB\Phonics "
$PH2         = " (With Name, Date, Time)\"
$PH3         = " - Individual Lessons\"

$FR          = "ELA\FR\Level "
$FR2         = "\FR "

$SV          = "ELA\SV\Student Work\Water Marked\TOC\"
$SV2         = "Passwords Science Vocab "

$STAMS       = "Math\STAMS\Water Marked\"
$STARS       = "ELA\STARS\Grayscale\Water Marked\STARS "


$DIRECTORY = $null
$FILE = $null

$FULLTYPE = $null

#Converts user inputted shorthand into the actual filename counterparts
if($type -eq "RW") {
  $FULLTYPE = "Rhyming Words"
}elseif($type -eq "CW") {
  $FULLTYPE = "Compound Words"
}elseif($type -eq "CONT") {
  $FULLTYPE = "Contractions"
}elseif($type -eq "SYN") {
  $FULLTYPE = "Synonyms"
}elseif($type -eq "ANT") {
  $FULLTYPE = "Antonyms"
}elseif($type -eq "HP") {
  $FULLTYPE = "Homophones"
}elseif($type -eq "HG") {
  $FULLTYPE = "Homographs"
}elseif($type -eq "PREF") {
  $FULLTYPE = "Prefixes"
}elseif($type -eq "WP") {
  $FULLTYPE = "Word Play"
}elseif($type -eq "PL") {
  $FULLTYPE = "Precise Language"
}elseif($type -eq "SUF") {
  $FULLTYPE = "Suffixes"
}elseif($type -eq "ROOT") {
  $FULLTYPE = "Roots"
}elseif($type -eq "IDI") {
  $FULLTYPE = "Idioms"
}elseif($type -eq "BLEN") {
  $FULLTYPE = "Blended Words"
}elseif($type -eq "CLIP") {
  $FULLTYPE = "Clipped Words"
}elseif($type -eq "HET") {
  $FULLTYPE = "Heteronyms"
}elseif($type -eq "Sentences") {
  $FULLTYPE = "Sent"
}elseif($type -eq "Pro") {
  $FULLTYPE = "Pron"
}elseif($type -eq "Pronoun") {
  $FULLTYPE = "Pron"
}elseif($type -eq "Adverbs") {
  $FULLTYPE = "Adv"
}elseif($type -eq "Vocab") {
  $FULLTYPE = "Voc"
}elseif($type -eq "Vocabulary") {
  $FULLTYPE = "Voc"
}else {
  $FULLTYPE = $type
}

#building the worksheet filepath by worksheet type
if($subject -eq "ccssm") {
  $DIRECTORY = $STAFFPATH + $CCSSMATH + $grade + $CCSS + $grade + "M - SB\"
  $FILE      = "CCSS " + $grade + $CCML + $lesson + " SB.pdf"
}elseif($subject -eq "sm") {
  #Regular SM Levels are 1A to 6B. 7A-8B are special and are all in their own directories
  #File paths missing...
  if($lesson -eq "7A") {
    $DIRECTORY = $STAFFPATH + "Math\SM\Discovering(7A & B) Math\TB 7A (eVer - STAMPED)\"

  }elseif($lesson -eq "7B") {
    $DIRECTORY = $STAFFPATH + "Math\SM\Discovering(7A & B) Math\TB 7B (eVer - STAMPED)\"

  }elseif($lesson -eq "8A") {
    $DIRECTORY = $STAFFPATH + "Math\SM\Dimensions(8A & B) Math\TB 8A (eVer - STAMPED)\"

  }elseif($lesson -eq "8B") {
    $DIRECTORY = $STAFFPATH + "Math\SM\Dimensions(8A & B) Math\TB 8B (eVer - STAMPED)\" 

  } else {
    $DIRECTORY = $STAFFPATH + $SM + $grade + "\"
    $FILE      = $grade + " Unit " + $lesson + " (STAMPED).pdf"
  }
}elseif($subject -eq "fm") {
  $DIRECTORY = $STAFFPATH + $FM + $grade + "\"
  $FILE      = "FM-" + $gradec + "-" + $type + ".pdf"
}elseif($subject -eq "stams") {
  $DIRECTORY = $STAFFPATH + $STAMS
  $FILE      = "STAMS- " + $grade + " (Water Marked).pdf"
}elseif($subject -eq "ccssr") {
  $DIRECTORY = $STAFFPATH + $CCSSELA + $grade + $CCSS + $grade + "R - SB\"
  $FILE      = "CCSS " + $grade + $CCRL + $lesson + " SB.pdf"
}elseif($subject -eq "lf") {
  $DIRECTORY = $STAFFPATH + $LF + $grade + $LF2 
  $FILE      = "LF" + $grade + " (*) " + $FULLTYPE + ".pdf"
}elseif($subject -eq "vf") {
  $DIRECTORY = $STAFFPATH + $VF + $grade + "\"
  $FILE      = $VF2 + $grade + " - (*) " + $FULLTYPE + " - Unit " $lesson + ".pdf"
}elseif($subject -eq "ph") {
  $DIRECTORY = $STAFFPATH + $PH + $grade + $PH2 + "PH" + $grade + $PH3
  $FILE      = "Phonics " + $grade + " - Lesson " + $lesson + ".pdf"
}elseif($subject -eq "fr") {
  $DIRECTORY = $STAFFPATH + $FR + $grade + "\"
  $FILE      = "FR " + $gradec + " - " + $type + ".pdf"
}elseif($subject -eq "sv") {
  $DIRECTORY = $STAFFPATH + $SV
  $FILE      = $SV2 + $gradec + ".pdf"
}elseif($subject -eq "stars") {
  $DIRECTORY = $STAFFPATH + $STARS + $grade + "\"
  $FILE      = "STARS " + $gradec + " - *.pdf" #doublecheck this. '*' might be too vague here
}

#If the script was supposed to find a pdf (every case except SMTB 7A-8B), print or open those files
if($FILE -like "*.pdf") {
  $FILEPDIR = $DIRECTORY + $FILE

  #Abuse gci's pattern matching. Note some of the built filenames have a '*' in them
  $FILES = Get-ChildItem $FILEPDIR
  foreach ($file in $files) {
    if($print) {
      #Print the worksheet (open file, print file, quit file)
      Start-Process -FilePath $file.FullName -Verb Print -PassThru | %{sleep 10;$_} | kill
    } else {
      #Used for testing the script's file finding capabilities without wasting paper
      Start-Process -FilePath $file.FullName
    }
  }
} else {
  #If the right file couldn't be parsed by this script (yet)
  #Open the directory where it should be and let the worker manually navigate to it
  ii $DIRECTORY
}

#Print out what was looked for, whether it exists or not.
Write-Host $DIRECTORY
Write-Host $FILE
