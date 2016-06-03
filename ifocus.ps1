.\loadDialog.ps1 -XamlPath 'GUI.xaml'

$USER = $Env:userprofile

$STAFFPATH   = $USER + "\Dropbox\Staff Print files\"

#Math worksheet directories
$CCSS        = "\CCSS "

$CCSSMATH    = "Math\CCSS Math\New_CC_8-25-2015-01-22 (eVER) STAMPED\Math\Gr."          
$CCML        = "M Lesson "

$SM          = "Math\SM\SM CC Edition\CC Edition Merged\"
$SMOLD       = "Math\SM\Text Books and Work Books (PRINT FROM HERE)\Easy Print Files (PRINT FROM HERE)\"

$FM          = "Math\FM\Focus Math (PRINT FROM HERE!!!)\Level "
$FM2         = "\FM-"

$MA          = "Math\MA\"

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

$LHB         = "ELA\CCSS Reading (STAMPED)\NEW CC LHB (ALL SPLIT) STAMPED\LHB - Grade "
$LHB2        = " (SPLIT)\LHB "
$LHB3        = " SB\LHB "

$STAMS       = "Math\STAMS\Water Marked\"
$STARS       = "ELA\STARS\Grayscale\Water Marked\STARS "

$gradec      = 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'
$gradesm     = "KM-A", "KM-B", "1A", "1B", "2A", "2B", "3A", "3B", "4A", "4B", "5A", "5B", "6A", "6B",  "TB7A", "TB7B", "TB8A", "TB8B", "WB7A", "WB7B", "WB8A", "WB8B"

$SM7A        = 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 2.1, 2.2, 2.3, 2.4, 2.5, 2.6, 2.7, 2.8, 3.1, 3.2, 3.3, 3.4, 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 5.1, 5.2, 5.3, 5.4, 5.5, 6.1, 6.2, 6.3, 6.4, 7.1, 7.2, 7.3, 7.4, 7.5, 8.1, 8.2, 8.3, 8.4, 8.5, 8.6
$SM7B        = 9.1, 9.2, 9.3, 10.1, 10.2, 10.3, 10.4, 11.1, 11.2, 11.3, 11.4, 11.5, 12.1, 12.2, 12.3, 12.4, 12.5, 12.6, 13.1, 13.2, 13.3, 13.4, 14.1, 14.2, 14.3, 14.4, 14.5, 15.1, 15.2, 15.3, 15.4, 15.5, 15.6, 16.1, 16.2, 16.3, 16.4, 17.1, 17.2, 17.3, 17.4, 17.5
$SM8A        = 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7, 1.8, 2.1, 2.2, 2.3, 2.4, 2.5, 2.6, 3.1, 3.2, 3.3, 3.4, 4.1, 4.2, 4.3, 4.4, 5.1, 5.2, 5.3, 5.4, 5.5, 5.6, 6.1, 6.2, 6.3, 6.4, 7.1, 7.2, 7.3, 7.4, 7.5
$SM8B        = 8.1, 8.2, 8.3, 8.1, 9.2, 9.3, 10.1, 10.2, 10.3, 10.4, 11.1, 11.2, 11.3, 11.4, 12.1, 12.2, 12.3, 12.4, 12.5, 13.1, 13.2, 13.3, 13.4, 14.1, 14.2, 14.3, 14.4, 14.5, 14.6

#Not ready yet
$WB7B = 9.1, 9.2
$WB8A = 1.1, 1.2
$WB8B = 9.1, 9.2

#Note: Unused: "WP"- Word Play
$VFT1        = "RW", "CW", "CONT", "SYN", "ANT", "HP", "HG", "PREF"
$VFT2        = "CW", "SYN", "PL", "ANT", "HP", "HG", "PREF", "SUF"
$VFT3        = "CW", "SYN", "PL", "ANT", "HP", "HG", "PREF", "SUF", "ROOT", "IDI"
$VFT456      = "CW", "SYN", "PL", "ANT", "HP", "HG", "HET", "PREF", "SUF", "ROOT", "IDI", "BLEN", "CLIP"

$PHT1        = "Word List", "1-4", "5-7", "8-10", "11-13", "14-16", "17-19", "20-23", "24-26", "27-30"

$LF123       = "Nouns", "Adj", "Pron", "Verbs", "Adv", "Sents", "Cap", "Abbrev", "Punct", "Usage", "Vocab", "Sent Ed"
$LF456       = "Nouns", "Adj", "Pron", "Verbs", "Adv", "Prep", "Sents", "Cap", "Abbrev", "Punct", "Usage", "Vocab", "Para Ed"

$FRT         = "C&C", "DCMI", "MP", "C&E", "MID", "SEQ"
$FMT         = "BNS", "DPA", "IGC", "UA", "UE", "UG"

$MAG         = "K (A Level)", "1 (B Level)", "2 (C Level)", "3 (D Level)", "4 (E Level)", "5 (F Level)", "6 (G Level)", "7 (H Level)", "8 (I Level)", "9 (J Level)", "10 (K Level)", "11 (L Level)", "12 (M Level)"

$SV9         = "L1-3", "L4-6", "L7-9"
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

function Box-Num {
  param($ComboBoxNum, $start, $finish)
  for($i=$start; $i -le $finish; $i++) {
    $ComboBoxNum.Items.add($i)
  }
  $ComboBoxNum.SelectedIndex = 0
}

function Grade-Select{
  param($GradeBox, $TypeBox, $UnitBox)
  $Subject = $comboBox1.SelectedItem.ToString()
  $Grade = $null
  try {
    $Grade = $GradeBox.SelectedItem.ToString()
  } catch [system.exception] {
    return
  } finally {
    #do nothing
  }
  
  $UnitBox.Items.Clear()
  if ($Subject -like "*FM" -or $Subject -like "*FR") {
    return
  }

  $TypeBox.Items.Clear()
  if ($Subject -like "*CCSSM") {
    if($Grade -like "*1") {
      Box-Num $TypeBox 1 35
    } elseif ($Grade -like "*2") {
      Box-Num $TypeBox 1 28
    } elseif ($Grade -like "*6") {
      Box-Num $TypeBox 1 29
    } elseif ($Grade -like "*K") {
      Box-Num $TypeBox 1 32
    } elseif (($Grade -like "*5") -or ($Grade -like "*8")) {
      Box-Num $TypeBox 1 31
    } else {
      Box-Num $TypeBox 1 33
    }
  } elseif ($Subject -like "*SM") {
    if($Grade -like "*KM-A") {
      Box-Num $TypeBox 1 8
    } elseif ($Grade -like "*KM-B") {
      Box-Num $TypeBox 9 15
    } elseif ($Grade -like "*1A") {
      Box-Num $TypeBox 1 10
    } elseif ($Grade -like "*1B") {
      Box-Num $TypeBox 11 17
    } elseif ($Grade -like "*2A") {
      Box-Num $TypeBox 1 5
    } elseif ($Grade -like "*2B") {
      Box-Num $TypeBox 6 12
    } elseif ($Grade -like "*3A") {
      Box-Num $TypeBox 1 5
    } elseif ($Grade -like "*3B") {
      Box-Num $TypeBox 6 13
    } elseif ($Grade -like "*4A") {
      Box-Num $TypeBox 1 5
    } elseif ($Grade -like "*4B") {
      Box-Num $TypeBox 6 11
    } elseif ($Grade -like "*5A") {
      Box-Num $TypeBox 1 6
    } elseif ($Grade -like "*5B") {
      Box-Num $TypeBox 7 15
    } elseif ($Grade -like "*6A") {
      Box-Num $TypeBox 1 6
    } elseif ($Grade -like "*6B") {
      Box-Num $TypeBox 7 13
    } elseif ($Grade -like "*TB7A") {
      Box-Change $TypeBox $SM7A
    } elseif ($Grade -like "*TB7B") {
      Box-Change $TypeBox $SM7B
    } elseif ($Grade -like "*TB8A") {
      Box-Change $TypeBox $SM8A
    } elseif ($Grade -like "*TB8B") {
      Box-Change $TypeBox $SM8B
    } elseif ($Grade -like "*WB7A") {
      Box-Num $TypeBox 1 8
    } elseif ($Grade -like "*WB7B") {
      Box-Num $TypeBox 9 17
    } elseif ($Grade -like "*WB8A") {
      Box-Num $TypeBox 1 7
    } elseif ($Grade -like "*WB8B") {
      Box-Num $TypeBox 8 14
    } else {
      #This shouldn't happen though...?
      Box-Num $TypeBox 1 5
    }
  } elseif ($Subject -like "*CCSSR") {
    if($Grade -like "*K") {
      Box-Num $TypeBox 1 18
    } elseif ($Grade -like "*4") {
      Box-Num $TypeBox 1 26
    } elseif ($Grade -like "*6") {
      Box-Num $TypeBox 1 20
    } elseif ($Grade -like "*7") {
      Box-Num $TypeBox 1 19
    } elseif (($Grade -like "*1") -or ($Grade -like "*8")) {
      Box-Num $TypeBox 1 21
    } else {
      Box-Num $TypeBox 1 22
    }
  } elseif ($Subject -like "*PH") {
    if($Grade -like "*1") {
      Box-Change $Typebox $PHT1
    } elseif ($Grade -like "*2") {
      $TypeBox.Items.add("Word List")
      Box-Num $TypeBox 1 32
    } elseif ($Grade -like "*3") {
      $TypeBox.Items.add("Word List")
      Box-Num $TypeBox 1 36
    }
  } elseif ($Subject -like "*SV") {
    if($Grade -like "*A") {
      Box-Change $TypeBox $SV9
    } elseif ($Grade -like "*B") {
      Box-Change $TypeBox $SV12
    } else {
      Box-Change $TypeBox $SV15
    }
  } elseif ($Subject -like "*STARS") {
    if(($Grade -like "*AA") -or ($Grade -like "*K")) {
      Box-Change $TypeBox $STARS6
    } elseif ($Grade -like "*A") {
      Box-Change $TypeBox $STARS8
    } else {
      Box-Change $TypeBox $STARS12
    }
  } elseif ($Subject -like "*LF") {
    if(($Grade -like "*1") -or ($Grade -like "*2") -or ($Grade -like "*3")) {
      Box-Change $TypeBox $LF123
    } else {
      Box-Change $TypeBox $LF456
    }
  } elseif ($Subject -like "*VF") {
    if($Grade -like "*1") {
      Box-Change $TypeBox $VFT1
    } elseif ($Grade -like "*2") {
      Box-Change $TypeBox $VFT2
    } elseif ($Grade -like "*3") {
      Box-Change $TypeBox $VFT3
    } else {
      Box-Change $Typebox $VFT456
    }
  } elseif ($Subject -like "*LHB") {
    if($Grade -like "*2") {
      Box-Num $TypeBox 1 26
    } elseif ($Grade -like "*3") {
      Box-Num $TypeBox 1 33
    } elseif ($Grade -like "*4") {
      Box-Num $TypeBox 1 24
    } elseif ($Grade -like "*5") {
      Box-Num $TypeBox 1 23
    } elseif ($Grade -like "*8") {
      Box-Num $TypeBox 1 19
    } else { #6 & 7
      Box-Num $TypeBox 1 17
    }
  }
    
}

$comboBox2.add_SelectionChanged({
  Grade-Select $comboBox2 $comboBox3 $comboBox6
  if($comboBox2.SelectedIndex -gt $comboBox4.SelectedIndex) {
    $comboBox4.SelectedIndex = $comboBox2.SelectedIndex
  }
})

$comboBox4.add_SelectionChanged({
  Grade-Select $comboBox4 $comboBox5 $comboBox7
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
    Box-Num $UnitBox 1 7
  } elseif ($Type -like "*CW") {
    if($Grade -like "*1") {
      Box-Num $UnitBox 1 6
    } else {
      Box-Num $UnitBox 1 3
    }
  } elseif ($Type -like "*CONT") {
    Box-Num $UnitBox 1 3
  } elseif ($Type -like "*SYN") {
    if($Grade -like "*1") {
      Box-Num $UnitBox 1 6
    } elseif ($Grade -like "*2") {
      Box-Num $UnitBox 1 7
    } elseif($Grade -like "*3" -or $Grade -like "*4") {
      Box-Num $UnitBox 1 3
    } else {
      Box-Num $UnitBox 1 2
    }
  } elseif ($Type -like "*PL") {
    if($Grade -like "*3") {
      Box-Num $UnitBox 1 10
    } else {
      Box-Num $UnitBox 1 6
    }
  } elseif ($Type -like "*ANT") {
    if($Grade -like "*1") {
      Box-Num $UnitBox 1 7
    } elseif ($Grade -like "*2") {
      Box-Num $UnitBox 1 5
    } elseif($Grade -like "*3") {
      Box-Num $UnitBox 1 3
    } else {
      Box-Num $UnitBox 1 2
    }
  } elseif ($Type -like "*HP") {
    if($Grade -like "*1") {
      Box-Num $UnitBox 1 7
    } elseif ($Grade -like "*2") {
      Box-Num $UnitBox 1 6
    } else {
      Box-Num $UnitBox 1 3
    }
  } elseif ($Type -like "*HG") {
    if($Grade -like "*1" -or $Grade -like "*2" -or $Grade -like "*3") {
      Box-Num $UnitBox 1 4
    } else {
      Box-Num $UnitBox 1 3
    }
  } elseif ($Type -like "*HET") {
    $UnitBox.Items.add(1)
  } elseif ($Type -like "*PREF") {
    if($Grade -like "*1") {
      Box-Num $UnitBox 1 2
    } elseif ($Grade -like "*2") {
      Box-Num $UnitBox 1 7
    } elseif($Grade -like "*3") {
      Box-Num $UnitBox 1 5
    } else {
      Box-Num $UnitBox 1 6
    }
  } elseif ($Type -like "*SUF") {
    if($Grade -like "*2" -or $Grade -like "*3") {
      Box-Num $UnitBox 1 4
    } else {
      Box-Num $UnitBox 1 6
    }
  } elseif ($Type -like "*ROOT") {
    if($Grade -like "*3" -or $Grade -like "*4") {
      Box-Num $UnitBox 1 4
    } else {
      Box-Num $UnitBox 1 5
    }
  } elseif ($Type -like "*IDI") {
    Box-Num $UnitBox 1 3
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
  if($comboBox2.SelectedIndex -eq $comboBox4.SelectedIndex) {
    if($comboBox3.SelectedIndex -gt $comboBox5.SelectedIndex) {
      $comboBox5.SelectedIndex = $comboBox3.SelectedIndex
    }
  }
})

$comboBox5.add_SelectionChanged({
  Type-Select $comboBox4 $comboBox5 $comboBox7
})

$comboBox6.add_SelectionChanged({
  if($comboBox2.SelectedIndex -eq $comboBox4.SelectedIndex) {
    if($comboBox3.SelectedIndex -eq $comboBox5.SelectedIndex) {
      if($comboBox6.SelectedIndex -gt $comboBox7.SelectedIndex) {
        $comboBox7.SelectedIndex = $comboBox6.SelectedIndex
      }
    }
  }
})
  
$comboBox1.add_SelectionChanged({
  $comboBox2.Items.Clear()
  $comboBox3.Items.Clear()
  $comboBox4.Items.Clear()
  $comboBox5.Items.Clear()
  $Subject = $comboBox1.SelectedItem.ToString()
  if($Subject -like "*CCSSM") {
    $comboBox2.Items.add('K')
    $comboBox4.Items.add('K')
    Box-Num $comboBox2 1 8
    Box-Num $comboBox4 1 8
    $textBlock4.Text = "CCSSM"
  } elseif ($Subject -like "*CCSSR") {
    $comboBox2.Items.add('K')
    $comboBox4.Items.add('K')
    Box-Num $comboBox2 1 8
    Box-Num $comboBox4 1 8
    $textBlock4.Text = "CCSSR"
  } elseif ($Subject -like "*SM") {
    Box-Change $comboBox2 $gradesm
    Box-Change $comboBox4 $gradesm
    $textBlock4.Text = "SM"
  } elseif ($Subject -like "*LF") {
    Box-Num $comboBox2 1 6
    Box-Num $comboBox4 1 6
    $textBlock4.Text = "LF"
  } elseif ($Subject -like "*VF") {
    Box-Num $comboBox2 1 6
    Box-Num $comboBox4 1 6
    $textBlock4.Text = "VF"
  } elseif ($Subject -like "*PH") {
    Box-Num $comboBox2 1 3
    Box-Num $comboBox4 1 3
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
  } elseif ($Subject -like "*MA") {
    Box-Change $comboBox2 $MAG
    Box-Change $comboBox4 $MAG
    $textBlock4.Text = "MA"
  } elseif ($Subject -like "*LHB") {
    Box-Num $comboBox2 2 8
    Box-Num $comboBox4 2 8
    $textBlock4.Text = "LHB"
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
  } catch [system.exception] {
    $NullBox3 = $true
  } finally {
    #do nothing
  }
  try {
    $VFUNIT = $comboBox6.SelectedItem.ToString()
  } catch [system.exception] {
    Write-Host "No VF"
  } finally {
    #do nothing
  }
  
#Converts user inputted shorthand into the actual filename counterparts
#TO-DO TESTING: Remove the FR, FM, and LF cases. 
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
    #Regular SM Levels are 1A to 5B. 6A-6B use the old directories, 7A-8B are special and are all in their own directories
    if($CBOX2 -like "*6A" -or $CBOX2 -like "*6B") {
      $DIRECTORY = $SMOLD + $CBOX2 + "\"
      $FILE      = $CBOX2 + " Unit " + $CBOX3 + " (STAMPED).pdf"
    } elseif($CBOX2 -like "*TB7A") {
      $DIRECTORY = $STAFFPATH + "Math\SM\Discovering(7A & B) Math\TB 7A (eVer - STAMPED)\"
      $FILE = "TB 7A " + $CBOX3 + " *.pdf"
    } elseif ($CBOX2 -like "*TB7B") {
      $DIRECTORY = $STAFFPATH + "Math\SM\Discovering(7A & B) Math\TB 7B (eVer - STAMPED)\"
      $FILE = "TB 7B " + $CBOX3 + " *.pdf"
    } elseif ($CBOX2 -like "*TB8A") {
      $DIRECTORY = $STAFFPATH + "Math\SM\Dimensions(8A & B) Math\TB 8A (eVer - STAMPED)\"
      $FILE = "TB 8A " + $CBOX3 + " *.pdf"
    } elseif ($CBOX2 -like "*TB8B") {
      $DIRECTORY = $STAFFPATH + "Math\SM\Dimensions(8A & B) Math\TB 8B (eVer - STAMPED)\" 
      $FILE = "TB 8B " + $CBOX3 + " *.pdf"
    } elseif($CBOX2 -like "*WB7A") {
      $DIRECTORY = $STAFFPATH + "Math\SM\Discovering(7A & B) Math\WB 7A - CUT + STAMPED\WB 7A Masters (by chapter)\"
      $FILE = "WB 7A " + $CBOX3 + " *.pdf"
    } elseif($CBOX2 -like "*WB7B") {
      $DIRECTORY = $STAFFPATH + "Math\SM\Discovering(7A & B) Math\WB 7B - CUT + STAMPED\WB 7B Masters (by chapter)\"
      $FILE = "WB 7B " + $CBOX3 + " *.pdf"
    } elseif($CBOX2 -like "*WB8A") {
      $DIRECTORY = $STAFFPATH + "Math\SM\Dimensions(8A & B) Math\WB 8A - CUT + STAMPED\WB 8A Masters (by chapter)\"
      $FILE = "WB 8A " + $CBOX3 + " *.pdf"
    } elseif($CBOX2 -like "*WB8B") {
      $DIRECTORY = $STAFFPATH + "Math\SM\Dimensions(8A & B) Math\WB 8B - CUT + STAMPED\WB 8B Masters (by chapter)\"
      $FILE = "WB 8B " + $CBOX3 + " *.pdf"
    } else {
      $DIRECTORY = $STAFFPATH + $SM + $CBOX2 + "\"
      $FILE      = $CBOX2 + " Unit " + $CBOX3 + ".pdf"
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
    $FILE      = "LF" + $CBOX2 + " (*) " + $FULLTYPE + ".pdf"
  } elseif ($CBOX1 -like "*vf") {
    $DIRECTORY = $STAFFPATH + $VF + $CBOX2 + "\"
    $FILE      = $VF2 + $CBOX2 + " - (*) " + $FULLTYPE + " - Unit " + $VFUNIT + "*.pdf"
  } elseif ($CBOX1 -like "*ph") {
    if($CBOX2 -like "*1") {
      $DIRECTORY = $STAFFPATH + $PH + $CBOX2 + $PH2
      $FILE      = "*" + $CBOX3 + "*.pdf"
    } else {
      $DIRECTORY = $STAFFPATH + $PH + $CBOX2 + $PH2 + "PH" + $CBOX2 + $PH3
      if($CBOX3 -like "*Word List") {
        $DIRECTORY = $STAFFPATH + $PH + $CBOX2 + $PH2
        $FILE      = "*" + $CBOX3 + ".pdf"
      } else {
        $DIRECTORY = $STAFFPATH + $PH + $CBOX2 + $PH2 + "PH" + $CBOX2 + $PH3
        $FILE      = "Phonics " + $CBOX2 + " - Lesson " + $CBOX3 + ".pdf"
      }
    }
  } elseif ($CBOX1 -like "*fr") {
    $DIRECTORY = $STAFFPATH + $FR + $CBOX2 + "\"
    $FILE      = "FR " + $CBOX2 + " - " + $CBOX3 + ".pdf"
  } elseif ($CBOX1 -like "*sv") {
    $DIRECTORY = $STAFFPATH + $SV
    $FILE      = $SV2 + $CBOX2 + " " + $FULLTYPE + ".pdf"
  } elseif ($CBOX1 -like "*stars") {
    $DIRECTORY = $STAFFPATH + $STARS + $CBOX2 + "\"
    $FILE      = "STARS " + $CBOX2 + " - " + $FULLTYPE + ".pdf"
  } elseif ($CBOX1 -like "*lhb") {
    $DIRECTORY = $STAFFPATH + $LHB + $CBOX2 + $LHB2 + $CBOX2 + " SB\"
    $FILE      = "LHB " + $CBOX2 + " - Lesson " +$CBOX3 + " SB.pdf"
  }

  #Print or open desired files
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
    #If the right file couldn't be parsed by this script (due to a bug)
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

$comboBox1.SelectedIndex = 0

#Launch the window
$xamGUI.ShowDialog() | out-null
