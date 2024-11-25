#Adds the library's.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#OrderedDictionary which stores the file types and their extensions.
$fileTypes = [ordered]@{
    "Video files" = @(".mp4", ".avi", ".mov", ".wmv", ".mkv", ".flv", ".mpeg", ".mpg", ".3gp", ".webm", ".vob", ".mts")
    "Image files" = @(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif", ".svg", ".webp", ".heic")
    "Word documents" = @(".doc", ".docx")
    "Powerpoint presentations" = @(".ptt", ".pttx", ".pps", ".ppsx")
    "Excel sheets" = @(".xls", ".xlsx", ".xlsm", ".xlsb")
}

#This creates the base GUI.
$form = New-Object System.Windows.Forms.Form
$form.Text = "File system organizer"
$form.Size = New-Object System.Drawing.Size(600, 380)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $false

#This creates the "Done" button.
$okayButton = New-Object System.Windows.Forms.Button
$okayButton.Size = New-Object System.Drawing.Size(60, 20)
$okayButton.Location = New-Object System.Drawing.Point(($form.Size.Width / 2 - $okayButton.Size.Width / 2), 300)
$okayButton.Text = "Done"

#This creates the drop down menu, in which you can select the file types.
$dropDownMenu = New-Object System.Windows.Forms.ComboBox
$dropDownMenu.Size = New-Object System.Drawing.Size(300, 50)
$dropDownMenu.Location = New-Object System.Drawing.Point(($form.Size.Width / 2 - $dropDownMenu.Size.Width / 2), 50)
foreach ($fileExtension in $fileTypes.Keys) {
    $dropDownMenu.Items.Add($fileExtension)
}
$form.Controls.Add($dropDownMenu)

#This creates a simple text label on top of the drop down menu.
$textLabel = New-Object System.Windows.Forms.Label
$textLabel.Size = New-Object System.Drawing.Size($dropDownMenu.Size)
$textLabel.Location = New-Object System.Drawing.Point($dropDownMenu.Location.X, ($dropDownMenu.Location.Y - 30))
$textLabel.Text = "Filter file types:"
$textLabel.Font = New-Object System.Drawing.Font("Arial", 10)
$textLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

#This creates the error label, which tells the user what went wrong or if it succeeded.
$errorLabel = New-Object System.Windows.Forms.Label
$errorLabel.Size = New-Object System.Drawing.Size(550, 15)
$errorLabel.Location = New-Object System.Drawing.Point(($form.Size.Width / 2 - $errorLabel.Size.Width / 2), ($okayButton.Location.Y - 20))
$errorLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

#Function to create the file explorer buttons and labels.
function AddFileExplorerSelection {
    param (
        [string]$text,
        [int]$positionY
    )

    #Creates the text box in which you can put the file directory.
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Size = New-Object System.Drawing.Size(550, 80)
    $textBox.Location = New-Object System.Drawing.Point(($form.ClientSize.Width / 2 - $textBox.ClientSize.Width / 2), $positionY)
    $textBox.Text = ""
    $textBox.Font = New-Object System.Drawing.Font("Arial", 10)
    $textBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    #Creates the label which tells you if it's the destination or origin directory.
    $label = New-Object System.Windows.Forms.Label
    $label.Size = New-Object System.Drawing.Size($textBox.Size)
    $label.Location = New-Object System.Drawing.Point($textBox.Location.X, ($positionY - 30))
    $label.Text = $text
    $label.Font = New-Object System.Drawing.Font("Arial", 10)
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

    #This creates the button which will open the file explorer window.
    $openFileExplorerButton = New-Object System.Windows.Forms.Button
    $openFileExplorerButton.Size = New-Object system.Drawing.Size(120, $textBox.Size.Height)
    $openFileExplorerButton.Location = New-Object System.Drawing.Point(($textBox.Size.Width - $openFileExplorerButton.Size.Width + $textBox.Location.X), $positionY)
    $openFileExplorerButton.Text = "Open File Explorer"
    $openFileExplorerButton.Font = New-Object System.Drawing.Font("Arial", 8)  

    #This adds controls to the button
    $form.Controls.Add($openFileExplorerButton)
    $form.Controls.Add($textBox)
    $form.Controls.Add($label)

    #Returns the textbox, which is necessary to update later on. Also returns the button, which we need to see if it gets clicked.
    return $textBox, $openFileExplorerButton
}

#Calls the function
$fromFolder, $fromFolderButton = AddFileExplorerSelection -text "Original folder" -positionY 150
$destinationFolder, $destinationFolderButton = AddFileExplorerSelection -text "Destination folder" -positionY 250

#Gives the button a functionality.
$fromFolderButton.Add_Click({
    $fileExplorerMenu = New-Object System.Windows.Forms.OpenFileDialog
    $fileExplorerMenu.Filter = "File folder|*."
    $fileExplorerMenu.FileName = "Select Folder"
    $fileExplorerMenu.CheckFileExists = $false
    $fileExplorerMenu.CheckPathExists = $true
    
    if ($fileExplorerMenu.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        $fromFolder.Text = [System.IO.Path]::GetDirectoryName($fileExplorerMenu.FileName)
    }
})

#Same as previous but with the destination directory instead.
$destinationFolderButton.Add_Click({
    $fileExplorerMenu = New-Object System.Windows.Forms.OpenFileDialog
    $fileExplorerMenu.Filter = "File folder|*."
    $fileExplorerMenu.FileName = "Select Folder"
    $fileExplorerMenu.CheckFileExists = $false
    $fileExplorerMenu.CheckPathExists = $true
    
    if ($fileExplorerMenu.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        $destinationFolder.Text = [System.IO.Path]::GetDirectoryName($fileExplorerMenu.FileName)
    }
})

#Gives the "Okay button" functionality.
$okayButton.Add_Click({
    #The counter to keep track of the amount of files succeeded
    $itemsTransfered = 0
    $itemsFailed = 0

    #This looks checks if the selected item from the drop down menu is the same as the item which is in the ordered dictionary.
    $filesOption = $dropDownMenu.SelectedItem
    if ($fileTypes.Contains($filesOption)) {
        
        #This gets an array of files which match the file extension defined in the ordered dictionary
        $extensions = $fileTypes[$filesOption]
        $files = Get-ChildItem -Path $fromFolder.Text | Where-Object { $extensions -contains $_.Extension}
        
        #Loops trough each of the files.
        foreach ($file in $files) {
            if ($destinationFolder.Text -ne "") {
                if (Test-Path -Path $destinationFolder.Text -PathType Container) {
                    try {
                        #This moves each file to their destination
                        Move-Item -Path $file.FullName -Destination $destinationFolder.Text -ErrorAction Stop
                        $itemsTransfered++
                    }
                    catch {
                        #Error handling
                        Write-Error "Error: $_"
                        $itemsFailed++
                    }
                }
                else {
                    #Error handling
                    Write-Error "'$($destinationFolder.Text)' is an invalid directory!"
                    $errorLabel.ForeColor = [System.Drawing.Color]::Red
                    $errorLabel.Text = "'$($destinationFolder.Text)' is an invalid directory!"
                    break
                }
            }
            else {
                #Error handling
                Write-Error "Destination directory cannot remain empty!"
                $errorLabel.ForeColor = [System.Drawing.Color]::Red
                $errorLabel.Text = "Destination directory cannot remain empty!"
                break
            }
        }
    #Error handling
    $errorLabel.ForeColor = [System.Drawing.Color]::Green
    $errorLabel.Text = "Files succesfully moved: $itemsTransfered, Files failed: $itemsFailed"
    }

})

#Adding more controls
$form.Controls.Add($okayButton)
$form.Controls.Add($textLabel)
$form.Controls.Add($errorLabel)

#Finally so the GUI is actually visible.
$form.TopMost = $true
$form.ShowDialog()
