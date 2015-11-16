$SCRIPTPATH    = "C:\Users\Kudo\Documents\WindowsPowerShell\Scripts"
$GITPATH       = "C:\Users\Kudo\Documents\GitHub\Powershell"
$VIMPATH       = "C:\Program Files (x86)\Vim\vim74\vim.exe"
$NOTEPADPPPATH = "C:\Program Files (x86)\Notepad++\notepad++.exe"

Set-Alias vi   $VIMPATH
Set-Alias vim  $VIMPATH
Set-Alias npp  $NOTEPADPPPATH
Set-Alias n++  $NOTEPADPPPATH

# for editing your PowerShell profile

Function cd-Scripts
{
    cd $SCRIPTPATH
}

Function cd-Git
{
    cd $GITPATH
}

Function Edit-Profile
{
    vim $profile
}

# for editing your Vim settings
Function Edit-Vimrc
{
    vim $home\_vimrc
}
