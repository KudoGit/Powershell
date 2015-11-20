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
$VFT         = "RW", "CW", "CONT", "SYN", "ANT", "HP", "HG", "PREF", "PL", "SUF", "ROOT", "IDI", "BLEN", "CLIP", "WP", "HET"
$LFT         = "Nouns", "Adj", "Pron", "Verbs", "Adv", "Sents", "Cap", "Abbrev", "Punct", "Usage", "Vocab", "Sent Ed", "Para Ed"
$FRT         = "C&C", "DCMI", "MP", "C&E", "MID", "SEQ"
$FMT         = "BNS", "Filler"

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
}

$comboBox2.add_SelectionChanged({
  if ($comboBox1.SelectedItem.ToString() -like "*CCSSM") {
    $comboBox3.Items.Clear()
    if($comboBox2.SelectedItem.ToString() -like "*1") {
      Box-Change $comboBox3 $CCSMG1
    } elseif($comboBox2.SelectedItem.ToString() -like "*2") {
      Box-Change $comboBox3 $CCSMG2
    } elseif($comboBox2.SelectedItem.ToString() -like "*6") {
      Box-Change $comboBox3 $CCSMG6
    } elseif($comboBox2.SelectedItem.ToString() -like "*K") {
      Box-Change $comboBox3 $CCSMGK
    } elseif(($comboBox2.SelectedItem.ToString() -like "*5") -or ($comboBox2.SelectedItem.ToString() -like "*8")) {
      Box-Change $comboBox3 $CCSMG58
    } else {
      Box-Change $comboBox3 $CCSMG347
    }
  }
  if ($comboBox1.SelectedItem.ToString() -like "*CCSSR") {
    $comboBox3.Items.Clear()
    if($comboBox2.SelectedItem.ToString() -like "*K") {
      Box-Change $comboBox3 $CCSRGK
    } elseif($comboBox2.SelectedItem.ToString() -like "*4") {
      Box-Change $comboBox3 $CCSRG4
    } elseif($comboBox2.SelectedItem.ToString() -like "*6") {
      Box-Change $comboBox3 $CCSRG6
    } elseif($comboBox2.SelectedItem.ToString() -like "*7") {
      Box-Change $comboBox3 $CCSRG7
    } elseif(($comboBox2.SelectedItem.ToString() -like "*1") -or ($comboBox2.SelectedItem.ToString() -like "*8")) {
      Box-Change $comboBox3 $CCSR18
    } else {
      Box-Change $comboBox3 $CCSRG235
    }
  }
  if ($comboBox1.SelectedItem.ToString() -like "*PH") {
    $comboBox3.Items.Clear()
    if($comboBox2.SelectedItem.ToString() -like "*1") {
      Box-Change $comboBox3 $PHT1
    } elseif($comboBox2.SelectedItem.ToString() -like "*2") {
      Box-Change $comboBox3 $PHT2
    } elseif($comboBox2.SelectedItem.ToString() -like "*3") {
      Box-Change $comboBox3 $PHT3
    }
  }
  if ($comboBox1.SelectedItem.ToString() -like "*SV") {
    $comboBox3.Items.Clear()
    if($comboBox2.SelectedItem.ToString() -like "*A") {
      Box-Change $comboBox3 $SV9
    } elseif($comboBox2.SelectedItem.ToString() -like "*B") {
      Box-Change $comboBox3 $SV12
    } else {
      Box-Change $comboBox3 $SV15
    }
  }
  if ($comboBox1.SelectedItem.ToString() -like "*STARS") {
    $comboBox3.Items.Clear()
    if(($comboBox2.SelectedItem.ToString() -like "*AA") -or ($comboBox2.SelectedItem.ToString() -like "*K")) {
      Box-Change $comboBox3 $STARS6
    } elseif($comboBox2.SelectedItem.ToString() -like "*A") {
      Box-Change $comboBox3 $STARS8
    } else {
      Box-Change $comboBox3 $STARS12
    }
  }

  
})

$comboBox4.add_SelectionChanged({
  if ($comboBox1.SelectedItem.ToString() -like "*CCSSM") {
    $comboBox5.Items.Clear()
    if($comboBox4.SelectedItem.ToString() -like "*1") {
      Box-Change $comboBox5 $CCSMG1
    } elseif($comboBox4.SelectedItem.ToString() -like "*2") {
      Box-Change $comboBox5 $CCSMG2
    } elseif($comboBox4.SelectedItem.ToString() -like "*6") {
      Box-Change $comboBox5 $CCSMG6
    } elseif($comboBox4.SelectedItem.ToString() -like "*K") {
      Box-Change $comboBox5 $CCSMGK
    } elseif(($comboBox4.SelectedItem.ToString() -like "*5") -or ($comboBox4.SelectedItem.ToString() -like "*8")) {
      Box-Change $comboBox5 $CCSMG58
    } else {
      Box-Change $comboBox5 $CCSMG347
    }
  }
  if ($comboBox1.SelectedItem.ToString() -like "*CCSSR") {
    $comboBox5.Items.Clear()
    if($comboBox4.SelectedItem.ToString() -like "*K") {
      Box-Change $comboBox5 $CCSRGK
    } elseif($comboBox4.SelectedItem.ToString() -like "*4") {
      Box-Change $comboBox5 $CCSRG4
    } elseif($comboBox4.SelectedItem.ToString() -like "*6") {
      Box-Change $comboBox5 $CCSRG6
    } elseif($comboBox4.SelectedItem.ToString() -like "*7") {
      Box-Change $comboBox5 $CCSRG7
    } elseif(($comboBox4.SelectedItem.ToString() -like "*1") -or ($comboBox4.SelectedItem.ToString() -like "*8")) {
      Box-Change $comboBox5 $CCSR18
    } else {
      Box-Change $comboBox5 $CCSRG235
    }
  }
  if ($comboBox1.SelectedItem.ToString() -like "*PH") {i
    $comboBox5.Items.Clear()
    if($comboBox4.SelectedItem.ToString() -like "*1") {
      Box-Change $comboBox5 $PHT1
    } elseif($comboBox4.SelectedItem.ToString() -like "*2") {
      Box-Change $comboBox5 $PHT2
    } elseif($comboBox4.SelectedItem.ToString() -like "*3") {
      Box-Change $comboBox5 $PHT3
    }
  }
  if ($comboBox1.SelectedItem.ToString() -like "*SV") {
    $comboBox5.Items.Clear()
    if($comboBox4.SelectedItem.ToString() -like "*A") {
      Box-Change $comboBox5 $SV9
    } elseif($comboBox2.SelectedItem.ToString() -like "*B") {
      Box-Change $comboBox5 $SV12
    } else {
      Box-Change $comboBox5 $SV15
    }
  }
  if ($comboBox1.SelectedItem.ToString() -like "*STARS") {
    $comboBox5.Items.Clear()
    if($comboBox4.SelectedItem.ToString() -like "*AA" -or $comboBox4.SelectedItem.ToString() -like "*K") {
      Box-Change $comboBox5 $STARS6
    } elseif($comboBox4.SelectedItem.ToString() -like "*A") {
      Box-Change $comboBox5 $STARS8
    } else {
      Box-Change $comboBox5 $STARS12
    }
  }
})

$comboBox1.add_SelectionChanged({
  Write-Host "Subject Changed"
  $comboBox2.Items.Clear()
  $comboBox3.Items.Clear()
  $comboBox4.Items.Clear()
  $comboBox5.Items.Clear()
  if($comboBox1.SelectedItem.ToString() -like "*CCSSM") {
    Box-Change $comboBox2 $gradecc
    Box-Change $comboBox4 $gradecc
    $textBlock4.Text = "CCSSM"
  } elseif ($comboBox1.SelectedItem.ToString() -like "*CCSSR") {
    Box-Change $comboBox2 $gradecc
    Box-Change $comboBox4 $gradecc
    $textBlock4.Text = "CCSSR"
  } elseif ($comboBox1.SelectedItem.ToString() -like "*SM") {
    Box-Change $comboBox2 $gradesm
    Box-Change $comboBox4 $gradesm
    $textBlock4.Text = "SM"
  } elseif ($comboBox1.SelectedItem.ToString() -like "*LF") {
    Box-Change $comboBox2 $gradef
    Box-Change $comboBox4 $gradef
    Box-Change $comboBox3 $LFT
    Box-Change $comboBox5 $LFT
    $textBlock4.Text = "LF"
  } elseif ($comboBox1.SelectedItem.ToString() -like "*VF") {
    Box-Change $comboBox2 $gradef
    Box-Change $comboBox4 $gradef
    Box-Change $comboBox3 $VFT
    Box-Change $comboBox5 $VFT
    $textBlock4.Text = "VF"
  } elseif ($comboBox1.SelectedItem.ToString() -like "*PH") {
    Box-Change $comboBox2 $gradep
    Box-Change $comboBox4 $gradep
    $textBlock4.Text = "PH"
  } elseif ($comboBox1.SelectedItem.ToString() -like "*FR") {
    Box-Change $comboBox2 $gradec
    Box-Change $comboBox4 $gradec
    Box-Change $comboBox3 $FRT
    Box-Change $comboBox5 $FRT
    $textBlock4.Text = "FR"
  } elseif ($comboBox1.SelectedItem.ToString() -like "*FM") {
    Box-Change $comboBox2 $gradec
    Box-Change $comboBox4 $gradec
    Box-Change $comboBox3 $FMT
    Box-Change $comboBox5 $FMT
    $textBlock4.Text = "FM"
  } elseif ($comboBox1.SelectedItem.ToString() -like "*SV") {
    Box-Change $comboBox2 $gradec
    Box-Change $comboBox4 $gradec
    $textBlock4.Text = "SV"
  } elseif ($comboBox1.SelectedItem.ToString() -like "*STARS") {
    Box-Change $comboBox2 $gradec
    Box-Change $comboBox4 $gradec
    $textBlock4.Text = "STARS"
  } elseif ($comboBox1.SelectedItem.ToString() -like "*STAMS") {
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

  $VFUNIT = $null

#Converts user inputted shorthand into the actual filename counterparts
  if($comboBox3.SelectedItem.ToString() -like "*RW") {
    $FULLTYPE = "Rhyming Words"
  }elseif($comboBox3.SelectedItem.ToString() -like "*CW") {
    $FULLTYPE = "Compound Words"
  }elseif($comboBox3.SelectedItem.ToString() -like "*CONT") {
    $FULLTYPE = "Contractions"
  }elseif($comboBox3.SelectedItem.ToString() -like "*SYN") {
    $FULLTYPE = "Synonyms"
  }elseif($comboBox3.SelectedItem.ToString() -like "*ANT") {
    $FULLTYPE = "Antonyms"
  }elseif($comboBox3.SelectedItem.ToString() -like "*HP") {
    $FULLTYPE = "Homophones"
  }elseif($comboBox3.SelectedItem.ToString() -like "*HG") {
    $FULLTYPE = "Homographs"
  }elseif($comboBox3.SelectedItem.ToString() -like "*PREF") {
    $FULLTYPE = "Prefixes"
  }elseif($comboBox3.SelectedItem.ToString() -like "*WP") {
    $FULLTYPE = "Word Play"
  }elseif($comboBox3.SelectedItem.ToString() -like "*PL") {
    $FULLTYPE = "Precise Language"
  }elseif($comboBox3.SelectedItem.ToString() -like "*SUF") {
    $FULLTYPE = "Suffixes"
  }elseif($comboBox3.SelectedItem.ToString() -like "*ROOT") {
    $FULLTYPE = "Roots"
  }elseif($comboBox3.SelectedItem.ToString() -like "*IDI") {
    $FULLTYPE = "Idioms"
  }elseif($comboBox3.SelectedItem.ToString() -like "*BLEN") {
    $FULLTYPE = "Blended Words"
  }elseif($comboBox3.SelectedItem.ToString() -like "*CLIP") {
    $FULLTYPE = "Clipped Words"
  }elseif($comboBox3.SelectedItem.ToString() -like "*HET") {
    $FULLTYPE = "Heteronyms"
  }elseif($comboBox3.SelectedItem.ToString() -like "*Nouns") {
    $FULLTYPE = "Nouns"
  }elseif($comboBox3.SelectedItem.ToString() -like "*Adj") {
    $FULLTYPE = "Adj"
  }elseif($comboBox3.SelectedItem.ToString() -like "*Pron") {
    $FULLTYPE = "Pron"
  }elseif($comboBox3.SelectedItem.ToString() -like "*Verbs") {
    $FULLTYPE = "Verbs"
  }elseif($comboBox3.SelectedItem.ToString() -like "*Adverbs") {
    $FULLTYPE = "Adv"
  }elseif($comboBox3.SelectedItem.ToString() -like "*Sents") {
    $FULLTYPE = "Sent"
  }elseif($comboBox3.SelectedItem.ToString() -like "*Vocab") {
    $FULLTYPE = "Voc"
  }elseif($comboBox3.SelectedItem.ToString() -like "*Cap") {
    $FULLTYPE = "Cap"
  }elseif($comboBox3.SelectedItem.ToString() -like "*Abbrev") {
    $FULLTYPE = "Abbrev"
  }elseif($comboBox3.SelectedItem.ToString() -like "*Punct") {
    $FULLTYPE = "Punct"
  }elseif($comboBox3.SelectedItem.ToString() -like "*Usage") {
    $FULLTYPE = "Usage"
  }elseif($comboBox3.SelectedItem.ToString() -like "*Vocab") {
    $FULLTYPE = "Voc"
  }elseif($comboBox3.SelectedItem.ToString() -like "*Sent Ed") {
    $FULLTYPE = "Sent Ed"
  }elseif($comboBox3.SelectedItem.ToString() -like "*Para Ed") {
    $FULLTYPE = "Para Ed"
  }elseif($comboBox3.SelectedItem.ToString() -like "*C&C") {
    $FULLTYPE = "C&C"
  }elseif($comboBox3.SelectedItem.ToString() -like "*DCMI") {
    $FULLTYPE = "DCMI"
  }elseif($comboBox3.SelectedItem.ToString() -like "*MP") {
    $FULLTYPE = "MP"
  }elseif($comboBox3.SelectedItem.ToString() -like "*C&E") {
    $FULLTYPE = "C&E"
  }elseif($comboBox3.SelectedItem.ToString() -like "*MID") {
    $FULLTYPE = "MID"
  }elseif($comboBox3.SelectedItem.ToString() -like "*SEQ") {
    $FULLTYPE = "SEQ"
  }elseif($comboBox3.SelectedItem.ToString() -like "*BNS") {
    $FULLTYPE = "BNS"
  }elseif($comboBox3.SelectedItem.ToString() -like "*DPA") {
    $FULLTYPE = "DPA"
  }elseif($comboBox3.SelectedItem.ToString() -like "*IGC") {
    $FULLTYPE = "IGC"
  }elseif($comboBox3.SelectedItem.ToString() -like "*UA") {
    $FULLTYPE = "UA"
  }elseif($comboBox3.SelectedItem.ToString() -like "*UE") {
    $FULLTYPE = "UE"
  }elseif($comboBox3.SelectedItem.ToString() -like "*UG") {
    $FULLTYPE = "UG"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L10") {
    $FULLTYPE = "Lesson 10"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L11") {
    $FULLTYPE = "Lesson 11"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L12") {
    $FULLTYPE = "Lesson 12"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L1") {
    $FULLTYPE = "Lesson 1"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L2") {
    $FULLTYPE = "Lesson 2"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L3") {
    $FULLTYPE = "Lesson 3"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L4") {
    $FULLTYPE = "Lesson 4"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L5") {
    $FULLTYPE = "Lesson 5"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L6") {
    $FULLTYPE = "Lesson 6"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L7") {
    $FULLTYPE = "Lesson 7"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L8") {
    $FULLTYPE = "Lesson 8"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L9") {
    $FULLTYPE = "Lesson 9"
  }elseif($comboBox3.SelectedItem.ToString() -like "*R1-2") {
    $FULLTYPE = "Review 1-2"
  }elseif($comboBox3.SelectedItem.ToString() -like "*R3-4") {
    $FULLTYPE = "Review 3-4"
  }elseif($comboBox3.SelectedItem.ToString() -like "*R5-6") {
    $FULLTYPE = "Review 5-6"
  }elseif($comboBox3.SelectedItem.ToString() -like "*R7-8") {
    $FULLTYPE = "Review 7-8"
  }elseif($comboBox3.SelectedItem.ToString() -like "*R1-3") {
    $FULLTYPE = "Review 1-3"
  }elseif($comboBox3.SelectedItem.ToString() -like "*R4-6") {
    $FULLTYPE = "Review 4-6"
  }elseif($comboBox3.SelectedItem.ToString() -like "*R7-9") {
    $FULLTYPE = "Review 7-9"
  }elseif($comboBox3.SelectedItem.ToString() -like "*R10-12") {
    $FULLTYPE = "Review 10-12"
  }elseif($comboBox3.SelectedItem.ToString() -like "*FR6") {
    $FULLTYPE = "Final Review 1-6"
  }elseif($comboBox3.SelectedItem.ToString() -like "*FR8") {
    $FULLTYPE = "Final Review 1-8"
  }elseif($comboBox3.SelectedItem.ToString() -like "*FR12") {
    $FULLTYPE = "Final Review 1-12"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L1-4") {
    $FULLTYPE = "Lessons 1-4"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L5-8") {
    $FULLTYPE = "Lessons 5-8"
  }elseif($comboBox3.SelectedItem.ToString() -like "*R9-12") {
    $FULLTYPE = "Lessons 9-12"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L1-3") {
    $FULLTYPE = "Lessons 1-3"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L4-6") {
    $FULLTYPE = "Lessons 4-6"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L7-9") {
    $FULLTYPE = "Lessons 7-9"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L1-5") {
    $FULLTYPE = "Lessons 1-5"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L6-10") {
    $FULLTYPE = "Lessons 6-10"
  }elseif($comboBox3.SelectedItem.ToString() -like "*L11-15") {
    $FULLTYPE = "Lessons 11-15"
  }else {
    $FULLTYPE = $comboBox3.SelectedItem.ToString()
  }

#building the worksheet filepath by worksheet type
  if($comboBox1.SelectedItem.ToString() -like "*ccssm") {
    $DIRECTORY = $STAFFPATH + $CCSSMATH + $comboBox2.SelectedItem.ToString() + $CCSS + $comboBox2.SelectedItem.ToString() + "M - SB\"
    $FILE      = "CCSS " + $comboBox2.SelectedItem.ToString() + $CCML + $comboBox3.SelectedItem.ToString() + " SB.pdf"
  }elseif($comboBox1.SelectedItem.ToString() -like "*sm") {
    #Regular SM Levels are 1A to 6B. 7A-8B are special and are all in their own directories
    #File paths missing...
    if($comboBox2.SelectedItem.ToString() -like "*7A") {
      $DIRECTORY = $STAFFPATH + "Math\SM\Discovering(7A & B) Math\TB 7A (eVer - STAMPED)\"

    }elseif($comboBox2.SelectedItem.ToString() -like "*7B") {
      $DIRECTORY = $STAFFPATH + "Math\SM\Discovering(7A & B) Math\TB 7B (eVer - STAMPED)\"

    }elseif($comboBox2.SelectedItem.ToString() -like "*8A") {
      $DIRECTORY = $STAFFPATH + "Math\SM\Dimensions(8A & B) Math\TB 8A (eVer - STAMPED)\"

    }elseif($comboBox2.SelectedItem.ToString() -like "*8B") {
      $DIRECTORY = $STAFFPATH + "Math\SM\Dimensions(8A & B) Math\TB 8B (eVer - STAMPED)\" 

    } else {
      $DIRECTORY = $STAFFPATH + $SM + $comboBox2.SelectedItem.ToString() + "\"
      $FILE      = $comboBox2.SelectedItem.ToString() + " Unit " + $comboBox3.SelectedItem.ToString() + " (STAMPED).pdf"
    }
  }elseif($comboBox1.SelectedItem.ToString() -like "*fm") {
    $DIRECTORY = $STAFFPATH + $FM + $comboBox2.SelectedItem.ToString() + "\"
    $FILE      = "FM-" + $comboBox2.SelectedItem.ToString() + "-" + $comboBox3.SelectedItem.ToString() + ".pdf"
  }elseif($comboBox1.SelectedItem.ToString() -like "*stams") {
    $DIRECTORY = $STAFFPATH + $STAMS
    $FILE      = "STAMS- " + $comboBox2.SelectedItem.ToString() + " (Water Marked).pdf"
  }elseif($comboBox1.SelectedItem.ToString() -like "*ccssr") {
    $DIRECTORY = $STAFFPATH + $CCSSELA + $comboBox2.SelectedItem.ToString() + $CCSS + $comboBox2.SelectedItem.ToString() + "R - SB\"
    $FILE      = "CCSS " + $comboBox2.SelectedItem.ToString() + $CCRL + $comboBox3.SelectedItem.ToString() + " SB.pdf"
  }elseif($comboBox1.SelectedItem.ToString() -like "*lf") {
    $DIRECTORY = $STAFFPATH + $LF + $comboBox2.SelectedItem.ToString() + $LF2 
    $FILE      = "LF" + $comboBox2.SelectedItem.ToString() + " (*) " + $FULLTYPE + ".pdf"
  }elseif($comboBox1.SelectedItem.ToString() -like "*vf") {
    $DIRECTORY = $STAFFPATH + $VF + $comboBox2.SelectedItem.ToString() + "\"
    $FILE      = $VF2 + $comboBox2.SelectedItem.ToString() + " - (*) " + $FULLTYPE + " - Unit " + $VFUNIT + "*.pdf"
  }elseif($comboBox1.SelectedItem.ToString() -like "*ph") {
    $DIRECTORY = $STAFFPATH + $PH + $comboBox2.SelectedItem.ToString() + $PH2 + "PH" + $comboBox2.SelectedItem.ToString() + $PH3
    $FILE      = "Phonics " + $comboBox2.SelectedItem.ToString() + " - Lesson " + $comboBox3.SelectedItem.ToString() + ".pdf"
  }elseif($comboBox1.SelectedItem.ToString() -like "*fr") {
    $DIRECTORY = $STAFFPATH + $FR + $comboBox2.SelectedItem.ToString() + "\"
    $FILE      = "FR " + $comboBox2.SelectedItem.ToString() + " - " + $comboBox3.SelectedItem.ToString() + ".pdf"
  }elseif($comboBox1.SelectedItem.ToString() -like "*sv") {
    $DIRECTORY = $STAFFPATH + $SV
    $FILE      = $SV2 + $comboBox2.SelectedItem.ToString() + " " + $FULLTYPE + ".pdf"
  }elseif($comboBox1.SelectedItem.ToString() -like "*stars") {
    $DIRECTORY = $STAFFPATH + $STARS + $comboBox2.SelectedItem.ToString() + "\"
    $FILE      = "STARS " + $comboBox2.SelectedItem.ToString() + " - " + $FULLTYPE + ".pdf"
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

#Launch the window
$xamGUI.ShowDialog() | out-null
