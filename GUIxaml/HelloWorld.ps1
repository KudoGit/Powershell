.\loadDialog.ps1 -XamlPath 'test.xaml'

#EVENT Handler
$button1.add_Click({
  Write-Host "Print Button"
})

$comboBox1.add_SelectionChanged({
  Write-Host "Subject Changed"
})

$comboBox2.add_SelectionChanged({
  Write-Host "Grade Changed"
})

$comboBox3.add_SelectionChanged({
  Write-Host "Lesson Changed"
})

#Launch the window
$xamGUI.ShowDialog() | out-null