.\loadDialog.ps1 -XamlPath 'test.xaml'

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
$CCSSELA     = "ELA\CCSS Reading (STAMPED)\New_CC_8-25-2015-01-22 (eVER) STAMPED\Reading - Gr."
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

$SV          = "ELA\SV\Student Work\SV CUT\"
$SV2         = "SV "

$STAMS       = "Math\STAMS\Water Marked\"
$STARS       = "ELA\STARS\Grayscale\Water Marked\STARS "

$comboBox1.SelectedIndex = 0

function Array-Num {
  param($Array)
  for($i=0; $i -lt $Array.length; $i++) {
    $Array[$i] = ($i+1)
  }
}

$PHT1 = New-Object int[] 30
$PHT2 = New-Object int[] 32
$PHT3 = New-Object int[] 36
Array-Num $PHT1
Array-Num $PHT2
Array-Num $PHT3

$CCSMG1   = New-Object int[] 35
$CCSMG2   = New-Object int[] 28
$CCSMG347 = New-Object int[] 33
$CCSMG58  = New-Object int[] 31
$CCSMG6   = New-Object int[] 29
$CCSMGK   = New-Object int[] 32
Array-Num $CCSMG1
Array-Num $CCSMG2
Array-Num $CCSMG347
Array-Num $CCSMG58
Array-Num $CCSMG6
Array-Num $CCSMGK

$CCSRGK   = New-Object int[] 18
$CCSRG18  = New-Object int[] 21
$CCSRG235 = New-Object int[] 22
$CCSRG4   = New-Object int[] 26
$CCSRG6   = New-Object int[] 20
$CCSRG7   = New-Object int[] 19
Array-Num $CCSRGK
Array-Num $CCSRG18
Array-Num $CCSRG235
Array-Num $CCSRG4
Array-Num $CCSRG6
Array-Num $CCSRG7


$gradep      = 1, 2, 3
$gradef      = 1, 2, 3, 4, 5, 6
$gradecc     = 'K', 1, 2, 3, 4, 5, 6, 7, 8
$gradec      = 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'
$gradesm     = "1A", "1B", "2A", "2B", "3A", "3B", "4A", "4B", "5A", "5B", "6A", "6B",  "7A", "7B", "8A", "8B"

#Note: Unused: "WP"- Word Play
$VFT1        = "RW", "CW", "CONT", "SYN", "ANT", "HP", "HG", "PREF"
$VFT2        = "CW", "SYN", "PL", "ANT", "HP", "HG", "PREF", "SUF"
$VFT3        = "CW", "SYN", "PL", "ANT", "HP", "HG", "PREF", "SUF", "ROOT", "IDI"
$VFT456      = "CW", "SYN", "PL", "ANT", "HP", "HG", "HET", "PREF", "SUF", "ROOT", "IDI", "BLEN", "CLIP"

$VFU2        = 1, 2
$VFU3        = 1, 2, 3
$VFU4        = 1, 2, 3, 4
$VFU5        = 1, 2, 3, 4, 5
$VFU6        = 1, 2, 3, 4, 5, 6
$VFU7        = 1, 2, 3, 4, 5, 6, 7
$VFU10       = 1, 2, 3, 4, 5, 6, 7, 8, 9, 10

$LF123       = "Nouns", "Adj", "Pron", "Verbs", "Adv", "Sents", "Cap", "Abbrev", "Punct", "Usage", "Vocab", "Sent Ed"
$LF456       = "Nouns", "Adj", "Pron", "Verbs", "Adv", "Prep", "Sents", "Cap", "Abbrev", "Punct", "Usage", "Vocab", "Para Ed"

$FRT         = "C&C", "DCMI", "MP", "C&E", "MID", "SEQ"
$FMT         = "BNS", "DPA", "IGC", "UA", "UE", "UG"

$SV9         = "L1-3", "L5-6", "L7-9"
$SV12        = "L1-4", "L5-8", "L9-12"
$SV15        = "L1-5", "L6-10", "L11-15"


$STARS6      = "L1", "L2", "R1-2", "L3", "L4", "R3-4", "L5", "L6", "R5-6", "FR6"
$STARS8      = "L1", "L2", "R1-2", "L3", "L4", "R3-4", "L5", "L6", "R5-6", "L7", "L8", "R7-8", "FR8"
$STARS12     = "L1", "L2", "L3", "R1-3", "L4", "L5", "L6", "R4-6", "L7", "L8", "L9", "R7-9", "L10", "L11", "L12", "R10-12", "FR12"

function Box-Change {
  param($ComboBoxNum, $Array)
  for($i=0; $i -lt $Array.length; $i++) {
    $ComboBoxNum.Items.add($array[$i])
  }
  $ComboBoxNum.SelectedIndex = 0
}


function Grade-Select{
  param($GradeBox, $TypeBox)
  $Subject = $comboBox1.SelectedItem.ToString()
  $Grade = $null
  try {
    $Grade = $GradeBox.SelectedItem.ToString()
  } catch [system.exception] {
    return
  } finally {
    #do nothing
  }
  if ($Subject -like "*CCSSM") {
    $TypeBox.Items.Clear()
    if($Grade -like "*1") {
      Box-Change $TypeBox $CCSMG1
    } elseif ($Grade -like "*2") {
      Box-Change $TypeBox $CCSMG2
    } elseif ($Grade -like "*6") {
      Box-Change $TypeBox $CCSMG6
    } elseif ($Grade -like "*K") {
      Box-Change $TypeBox $CCSMGK
    } elseif (($Grade -like "*5") -or ($Grade -like "*8")) {
      Box-Change $TypeBox $CCSMG58
    } else {
      Box-Change $TypeBox $CCSMG347
    }
  } elseif ($Subject -like "*CCSSR") {
    $TypeBox.Items.Clear()
    if($Grade -like "*K") {
      Box-Change $TypeBox $CCSRGK
    } elseif ($Grade -like "*4") {
      Box-Change $TypeBox $CCSRG4
    } elseif ($Grade -like "*6") {
      Box-Change $TypeBox $CCSRG6
    } elseif ($Grade -like "*7") {
      Box-Change $TypeBox $CCSRG7
    } elseif (($Grade -like "*1") -or ($Grade -like "*8")) {
      Box-Change $TypeBox $CCSR18
    } else {
      Box-Change $TypeBox $CCSRG235
    }
  } elseif ($Subject -like "*PH") {
    $TypeBox.Items.Clear()
    if($Grade -like "*1") {
      Box-Change $TypeBox $PHT1
    } elseif ($Grade -like "*2") {
      Box-Change $TypeBox $PHT2
    } elseif ($Grade -like "*3") {
      Box-Change $TypeBox $PHT3
    }
  } elseif ($Subject -like "*SV") {
    $TypeBox.Items.Clear()
    if($Grade -like "*A") {
      Box-Change $TypeBox $SV9
    } elseif ($Grade -like "*B") {
      Box-Change $TypeBox $SV12
    } else {
      Box-Change $TypeBox $SV15
    }
  } elseif ($Subject -like "*STARS") {
    $TypeBox.Items.Clear()
    if(($Grade -like "*AA") -or ($Grade -like "*K")) {
      Box-Change $TypeBox $STARS6
    } elseif ($Grade -like "*A") {
      Box-Change $TypeBox $STARS8
    } else {
      Box-Change $TypeBox $STARS12
    }
  } elseif ($Subject -like "*LF") {
    $TypeBox.Items.Clear()
    if(($Grade -like "*1") -or ($Grade -like "*2") -or ($Grade -like "*3")) {
      Box-Change $TypeBox $LF123
    } else {
      Box-Change $TypeBox $LF456
    }
  } elseif ($Subject -like "*VF") {
    $TypeBox.Items.Clear()
    if($Grade -like "*1") {
      Box-Change $TypeBox $VFT1
    } elseif ($Grade -like "*2") {
      Box-Change $TypeBox $VFT2
    } elseif ($Grade -like "*3") {
      Box-Change $TypeBox $VFT3
    } else {
      Box-Change $Typebox $VFT456
    }
  }
}

$comboBox2.add_SelectionChanged({
  Grade-Select $comboBox2 $comboBox3
  if($comboBox2.SelectedIndex -gt $comboBox4.SelectedIndex) {
    $comboBox4.SelectedIndex = $comboBox2.SelectedIndex
  }
})

$comboBox4.add_SelectionChanged({
  Grade-Select $comboBox4 $comboBox5
})

function Type-Select{
  param($GradeBox, $TypeBox, $UnitBox)
  $Subject = $comboBox1.SelectedItem.ToString()
  $Grade = $null
  $Type = $null
  try {
    $Grade = $GradeBox.SelectedItem.ToString()
    $Type  = $TypeBox.SelectedItem.ToString()
  } catch [system.exception] {
    return
  } finally {
    #do nothing
  }
  if($Subject -notlike "*VF") {
    return
  }
  $UnitBox.Items.Clear()
  if($Type -like "*RW") {
    Box-Change $UnitBox $VFU7
  } elseif ($Type -like "*CW") {
    if($Grade -like "*1") {
      Box-Change $UnitBox $VFU6
    } else {
      Box-Change $UnitBox $VFU3
    }
  } elseif ($Type -like "*CONT") {
    Box-Change $UnitBox $VFU3
  } elseif ($Type -like "*SYN") {
    if($Grade -like "*1") {
      Box-Change $UnitBox $VFU6
    } elseif ($Grade -like "*2") {
      Box-Change $UnitBox $VFU7
    } elseif($Grade -like "*3" -or $Grade -like "*4") {
      Box-Change $UnitBox $VFU3
    } else {
      Box-Change $UnitBox $VFU2
    }
  } elseif ($Type -like "*PL") {
    if($Grade -like "*3") {
      Box-Change $UnitBox $VFU10
    } else {
      Box-Change $UnitBox $VFU6
    }
  } elseif ($Type -like "*ANT") {
    if($Grade -like "*1") {
      Box-Change $UnitBox $VFU7
    } elseif ($Grade -like "*2") {
      Box-Change $UnitBox $VFU5
    } elseif($Grade -like "*3") {
      Box-Change $UnitBox $VFU3
    } else {
      Box-Change $UnitBox $VFU2
    }
  } elseif ($Type -like "*HP") {
    if($Grade -like "*1") {
      Box-Change $UnitBox $VFU7
    } elseif ($Grade -like "*2") {
      Box-Change $UnitBox $VFU6
    } else {
      Box-Change $UnitBox $VFU3
    }
  } elseif ($Type -like "*HG") {
    if($Grade -like "*1" -or $Grade -like "*2" -or $Grade -like "*3") {
      Box-Change $UnitBox $VFU4
    } else {
      Box-Change $UnitBox $VFU3
    }
  } elseif ($Type -like "*HET") {
    $UnitBox.Items.add(1)
  } elseif ($Type -like "*PREF") {
    if($Grade -like "*1") {
      Box-Change $UnitBox $VFU2
    } elseif ($Grade -like "*2") {
      Box-Change $UnitBox $VFU7
    } elseif($Grade -like "*3") {
      Box-Change $UnitBox $VFU5
    } else {
      Box-Change $UnitBox $VFU6
    }
  } elseif ($Type -like "*SUF") {
    if($Grade -like "*2" -or $Grade -like "*3") {
      Box-Change $UnitBox $VFU4
    } else {
      Box-Change $UnitBox $VFU6
    }
  } elseif ($Type -like "*ROOT") {
    if($Grade -like "*3" -or $Grade -like "*4") {
      Box-Change $UnitBox $VFU4
    } else {
      Box-Change $UnitBox $VFU5
    }
  } elseif ($Type -like "*IDI") {
    Box-Change $UnitBox $VFU3
  } elseif ($Type -like "*BLEN") {
    $UnitBox.Items.add(1)
  } elseif ($Type -like "*CLIP") {
    $UnitBox.Items.add(1)
  } elseif ($Type -like "*WP") {
    $UnitBox.Items.add(1)
  }
}

$comboBox3.add_SelectionChanged({
  Type-Select $comboBox2 $comboBox3 $comboBox6
})

$comboBox5.add_SelectionChanged({
  Type-Select $comboBox4 $comboBox5 $comboBox7
})

$comboBox1.add_SelectionChanged({
  $comboBox2.Items.Clear()
  $comboBox3.Items.Clear()
  $comboBox4.Items.Clear()
  $comboBox5.Items.Clear()
  $Subject = $comboBox1.SelectedItem.ToString()
  if($Subject -like "*CCSSM") {
    Box-Change $comboBox2 $gradecc
    Box-Change $comboBox4 $gradecc
    $textBlock4.Text = "CCSSM"
  } elseif ($Subject -like "*CCSSR") {
    Box-Change $comboBox2 $gradecc
    Box-Change $comboBox4 $gradecc
    $textBlock4.Text = "CCSSR"
  } elseif ($Subject -like "*SM") {
    Box-Change $comboBox2 $gradesm
    Box-Change $comboBox4 $gradesm
    $textBlock4.Text = "SM"
  } elseif ($Subject -like "*LF") {
    Box-Change $comboBox2 $gradef
    Box-Change $comboBox4 $gradef
    $textBlock4.Text = "LF"
  } elseif ($Subject -like "*VF") {
    Box-Change $comboBox2 $gradef
    Box-Change $comboBox4 $gradef
    $textBlock4.Text = "VF"
  } elseif ($Subject -like "*PH") {
    Box-Change $comboBox2 $gradep
    Box-Change $comboBox4 $gradep
    $textBlock4.Text = "PH"
  } elseif ($Subject -like "*FR") {
    Box-Change $comboBox2 $gradec
    Box-Change $comboBox4 $gradec
    Box-Change $comboBox3 $FRT
    Box-Change $comboBox5 $FRT
    $textBlock4.Text = "FR"
  } elseif ($Subject -like "*FM") {
    Box-Change $comboBox2 $gradec
    Box-Change $comboBox4 $gradec
    Box-Change $comboBox3 $FMT
    Box-Change $comboBox5 $FMT
    $textBlock4.Text = "FM"
  } elseif ($Subject -like "*SV") {
    Box-Change $comboBox2 $gradec
    Box-Change $comboBox4 $gradec
    $textBlock4.Text = "SV"
  } elseif ($Subject -like "*STARS") {
    $comboBox2.Items.add("AA")
    $comboBox4.Items.add("AA")
    Box-Change $comboBox2 $gradec
    Box-Change $comboBox4 $gradec
    $textBlock4.Text = "STARS"
  } elseif ($Subject -like "*STAMS") {
    Box-Change $comboBox2 $gradec
    Box-Change $comboBox4 $gradec
    $textBlock4.Text = "STAMS"
  }
})

function Find-File{
  param([bool]$print)
  $DIRECTORY = $null
  $FILE = $null

  $FULLTYPE = $null

  $LFNUM  = $comboBox2.SelectedIndex+1
  $VFUNIT = $null

  $CBOX1 = $comboBox1.SelectedItem.ToString()
  $CBOX2 = $comboBox2.SelectedItem.ToString()
  $CBOX3 = $null
  $NullBox3  = $false
  try {
    $CBOX3  = $comboBox3.SelectedItem.ToString()
    $VFUNIT = $comboBox6.SelectedItem.ToString()
  } catch [system.exception] {
    $NullBox3 = $true
  } finally {
    #do nothing
  }
  
#Converts user inputted shorthand into the actual filename counterparts
  if($NullBox3) {
    Write-Host "NullBox3"
  } elseif($CBOX3 -like "*RW") {
    $FULLTYPE = "Rhyming Words"
  } elseif ($CBOX3 -like "*CW") {
    $FULLTYPE = "Compound Words"
  } elseif ($CBOX3 -like "*CONT") {
    $FULLTYPE = "Contractions"
  } elseif ($CBOX3 -like "*SYN") {
    $FULLTYPE = "Synonyms"
  } elseif ($CBOX3 -like "*ANT") {
    $FULLTYPE = "Antonyms"
  } elseif ($CBOX3 -like "*HP") {
    $FULLTYPE = "Homophones"
  } elseif ($CBOX3 -like "*HG") {
    $FULLTYPE = "Homographs"
  } elseif ($CBOX3 -like "*PREF") {
    $FULLTYPE = "Prefixes"
  } elseif ($CBOX3 -like "*WP") {
    $FULLTYPE = "Word Play"
  } elseif ($CBOX3 -like "*PL") {
    $FULLTYPE = "Precise Language"
  } elseif ($CBOX3 -like "*SUF") {
    $FULLTYPE = "Suffixes"
  } elseif ($CBOX3 -like "*ROOT") {
    $FULLTYPE = "Roots"
  } elseif ($CBOX3 -like "*IDI") {
    $FULLTYPE = "Idioms"
  } elseif ($CBOX3 -like "*BLEN") {
    $FULLTYPE = "Blended Words"
  } elseif ($CBOX3 -like "*CLIP") {
    $FULLTYPE = "Clipped Words"
  } elseif ($CBOX3 -like "*HET") {
    $FULLTYPE = "Heteronyms"
  } elseif ($CBOX3 -like "*Nouns") {
    $FULLTYPE = "Nouns"
  } elseif ($CBOX3 -like "*Adj") {
    $FULLTYPE = "Adj"
  } elseif ($CBOX3 -like "*Pron") {
    $FULLTYPE = "Pron"
  } elseif ($CBOX3 -like "*Verbs") {
    $FULLTYPE = "Verbs"
  } elseif ($CBOX3 -like "*Adverbs") {
    $FULLTYPE = "Adv"
  } elseif ($CBOX3 -like "*Sents") {
    $FULLTYPE = "Sent"
  } elseif ($CBOX3 -like "*Vocab") {
    $FULLTYPE = "Voc"
  } elseif ($CBOX3 -like "*Cap") {
    $FULLTYPE = "Cap"
  } elseif ($CBOX3 -like "*Abbrev") {
    $FULLTYPE = "Abbrev"
  } elseif ($CBOX3 -like "*Punct") {
    $FULLTYPE = "Punct"
  } elseif ($CBOX3 -like "*Usage") {
    $FULLTYPE = "Usage"
  } elseif ($CBOX3 -like "*Vocab") {
    $FULLTYPE = "Voc"
  } elseif ($CBOX3 -like "*Sent Ed") {
    $FULLTYPE = "Sent Ed"
  } elseif ($CBOX3 -like "*Para Ed") {
    $FULLTYPE = "Para Ed"
  } elseif ($CBOX3 -like "*C&C") {
    $FULLTYPE = "C&C"
  } elseif ($CBOX3 -like "*DCMI") {
    $FULLTYPE = "DCMI"
  } elseif ($CBOX3 -like "*MP") {
    $FULLTYPE = "MP"
  } elseif ($CBOX3 -like "*C&E") {
    $FULLTYPE = "C&E"
  } elseif ($CBOX3 -like "*MID") {
    $FULLTYPE = "MID"
  } elseif ($CBOX3 -like "*SEQ") {
    $FULLTYPE = "SEQ"
  } elseif ($CBOX3 -like "*BNS") {
    $FULLTYPE = "BNS"
  } elseif ($CBOX3 -like "*DPA") {
    $FULLTYPE = "DPA"
  } elseif ($CBOX3 -like "*IGC") {
    $FULLTYPE = "IGC"
  } elseif ($CBOX3 -like "*UA") {
    $FULLTYPE = "UA"
  } elseif ($CBOX3 -like "*UE") {
    $FULLTYPE = "UE"
  } elseif ($CBOX3 -like "*UG") {
    $FULLTYPE = "UG"
  } elseif ($CBOX3 -like "*L10") {
    $FULLTYPE = "Lesson 10"
  } elseif ($CBOX3 -like "*L11") {
    $FULLTYPE = "Lesson 11"
  } elseif ($CBOX3 -like "*L12") {
    $FULLTYPE = "Lesson 12"
  } elseif ($CBOX3 -like "*L1") {
    $FULLTYPE = "Lesson 1"
  } elseif ($CBOX3 -like "*L2") {
    $FULLTYPE = "Lesson 2"
  } elseif ($CBOX3 -like "*L3") {
    $FULLTYPE = "Lesson 3"
  } elseif ($CBOX3 -like "*L4") {
    $FULLTYPE = "Lesson 4"
  } elseif ($CBOX3 -like "*L5") {
    $FULLTYPE = "Lesson 5"
  } elseif ($CBOX3 -like "*L6") {
    $FULLTYPE = "Lesson 6"
  } elseif ($CBOX3 -like "*L7") {
    $FULLTYPE = "Lesson 7"
  } elseif ($CBOX3 -like "*L8") {
    $FULLTYPE = "Lesson 8"
  } elseif ($CBOX3 -like "*L9") {
    $FULLTYPE = "Lesson 9"
  } elseif ($CBOX3 -like "*R1-2") {
    $FULLTYPE = "Review 1-2"
  } elseif ($CBOX3 -like "*R3-4") {
    $FULLTYPE = "Review 3-4"
  } elseif ($CBOX3 -like "*R5-6") {
    $FULLTYPE = "Review 5-6"
  } elseif ($CBOX3 -like "*R7-8") {
    $FULLTYPE = "Review 7-8"
  } elseif ($CBOX3 -like "*R1-3") {
    $FULLTYPE = "Review 1-3"
  } elseif ($CBOX3 -like "*R4-6") {
    $FULLTYPE = "Review 4-6"
  } elseif ($CBOX3 -like "*R7-9") {
    $FULLTYPE = "Review 7-9"
  } elseif ($CBOX3 -like "*R10-12") {
    $FULLTYPE = "Review 10-12"
  } elseif ($CBOX3 -like "*FR6") {
    $FULLTYPE = "Final Review 1-6"
  } elseif ($CBOX3 -like "*FR8") {
    $FULLTYPE = "Final Review 1-8"
  } elseif ($CBOX3 -like "*FR12") {
    $FULLTYPE = "Final Review 1-12"
  } elseif ($CBOX3 -like "*L1-4") {
    $FULLTYPE = "Lessons 1-4"
  } elseif ($CBOX3 -like "*L5-8") {
    $FULLTYPE = "Lessons 5-8"
  } elseif ($CBOX3 -like "*R9-12") {
    $FULLTYPE = "Lessons 9-12"
  } elseif ($CBOX3 -like "*L1-3") {
    $FULLTYPE = "Lessons 1-3"
  } elseif ($CBOX3 -like "*L4-6") {
    $FULLTYPE = "Lessons 4-6"
  } elseif ($CBOX3 -like "*L7-9") {
    $FULLTYPE = "Lessons 7-9"
  } elseif ($CBOX3 -like "*L1-5") {
    $FULLTYPE = "Lessons 1-5"
  } elseif ($CBOX3 -like "*L6-10") {
    $FULLTYPE = "Lessons 6-10"
  } elseif ($CBOX3 -like "*L11-15") {
    $FULLTYPE = "Lessons 11-15"
  }else {
    $FULLTYPE = $CBOX3
  }

#building the worksheet filepath by worksheet type
  if($CBOX1 -like "*ccssm") {
    $DIRECTORY = $STAFFPATH + $CCSSMATH + $CBOX2 + $CCSS + $CBOX2 + "M - SB\"
    $FILE      = "CCSS " + $CBOX2 + $CCML + $CBOX3 + " SB.pdf"
  } elseif ($CBOX1 -like "*sm") {
    #Regular SM Levels are 1A to 6B. 7A-8B are special and are all in their own directories
    #File paths missing...
    if($CBOX2 -like "*7A") {
      $DIRECTORY = $STAFFPATH + "Math\SM\Discovering(7A & B) Math\TB 7A (eVer - STAMPED)\"

    } elseif ($CBOX2 -like "*7B") {
      $DIRECTORY = $STAFFPATH + "Math\SM\Discovering(7A & B) Math\TB 7B (eVer - STAMPED)\"

    } elseif ($CBOX2 -like "*8A") {
      $DIRECTORY = $STAFFPATH + "Math\SM\Dimensions(8A & B) Math\TB 8A (eVer - STAMPED)\"

    } elseif ($CBOX2 -like "*8B") {
      $DIRECTORY = $STAFFPATH + "Math\SM\Dimensions(8A & B) Math\TB 8B (eVer - STAMPED)\" 

    } else {
      $DIRECTORY = $STAFFPATH + $SM + $CBOX2 + "\"
      $FILE      = $CBOX2 + " Unit " + $CBOX3 + " (STAMPED).pdf"
    }
  } elseif ($CBOX1 -like "*fm") {
    $DIRECTORY = $STAFFPATH + $FM + $CBOX2 + "\"
    $FILE      = "FM-" + $CBOX2 + "-" + $CBOX3 + ".pdf"
  } elseif ($CBOX1 -like "*stams") {
    $DIRECTORY = $STAFFPATH + $STAMS
    $FILE      = "STAMS- " + $CBOX2 + " (Water Marked).pdf"
  } elseif ($CBOX1 -like "*ccssr") {
    $DIRECTORY = $STAFFPATH + $CCSSELA + $CBOX2 + $CCSS + $CBOX2 + "R - SB\"
    $FILE      = "CCSS " + $CBOX2 + $CCRL + $CBOX3 + " SB.pdf"
  } elseif ($CBOX1 -like "*lf") {
    $DIRECTORY = $STAFFPATH + $LF + $CBOX2 + $LF2 
    $FILE      = "LF" + $CBOX2 + " (" + $LFNUM +") " + $FULLTYPE + ".pdf"
  } elseif ($CBOX1 -like "*vf") {
    $DIRECTORY = $STAFFPATH + $VF + $CBOX2 + "\"
    $FILE      = $VF2 + $CBOX2 + " - (*) " + $FULLTYPE + " - Unit " + $VFUNIT + "*.pdf"
  } elseif ($CBOX1 -like "*ph") {
    $DIRECTORY = $STAFFPATH + $PH + $CBOX2 + $PH2 + "PH" + $CBOX2 + $PH3
    $FILE      = "Phonics " + $CBOX2 + " - Lesson " + $CBOX3 + ".pdf"
  } elseif ($CBOX1 -like "*fr") {
    $DIRECTORY = $STAFFPATH + $FR + $CBOX2 + "\"
    $FILE      = "FR " + $CBOX2 + " - " + $CBOX3 + ".pdf"
  } elseif ($CBOX1 -like "*sv") {
    $DIRECTORY = $STAFFPATH + $SV
    $FILE      = $SV2 + $CBOX2 + " " + $FULLTYPE + ".pdf"
  } elseif ($CBOX1 -like "*stars") {
    $DIRECTORY = $STAFFPATH + $STARS + $CBOX2 + "\"
    $FILE      = "STARS " + $CBOX2 + " - " + $FULLTYPE + ".pdf"
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
}

$button1.add_Click({
  Find-File $false
})

$button2.add_Click({
  Find-File $true
})

function Find-Range {
  param($print)
  $start2 = $comboBox2.SelectedIndex
  $start3 = $comboBox3.SelectedIndex
  $start6 = $comboBox6.SelectedIndex
  $end2   = $comboBox4.SelectedIndex
  $end3   = $comboBox5.SelectedIndex
  $end6   = $comboBox7.SelectedIndex
  
  $ISVF   = ($comboBox1.SelectedItem.ToString() -like "*VF")
  for($i=0; $i -lt $comboBox2.Items.Count; $i++) {
    if($end2 -lt $i) {
      break
    }
    for($j=0; $j -lt $comboBox3.Items.Count; $j++) {
      if(($end2 -eq $i -and $end3 -lt $j) -or ($end2 -lt $i) ) {
        break
      }
      if($ISVF) {
        for($k=0; $k -lt $comboBox6.Items.Count; $k++) {
          if(($end2 -eq $i -and $end3 -eq $j -and $end6 -lt $k) -or ($end3 -lt $j)) {
            break
          }
          if(($start2 -eq $i -and $start3 -eq $j -and $start6 -le $k) -or ($start2 -le $i -and $start3 -lt $j)) {
            $comboBox2.SelectedIndex = $i
            $comboBox3.SelectedIndex = $j
            $comboBox6.SelectedIndex = $k
            Find-File $print
          }
        }
      } elseif(($start2 -eq $i -and $start3 -le $j) -or ($start2 -lt $i) ) {
        $comboBox2.SelectedIndex = $i
        $comboBox3.SelectedIndex = $j
        Find-File $print
      }
    }
  }
  $comboBox2.SelectedIndex = $start2
  $comboBox3.SelectedIndex = $start3
  $comboBox6.SelectedIndex = $start6
}

$button3.add_Click({
  Find-Range $false
})

$button4.add_Click({
  Find-Range $true
})

#Launch the window
$xamGUI.ShowDialog() | out-null
