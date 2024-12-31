#Adds the library's.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#OrderedDictionary which stores the file types and their extensions.
$fileTypes = [ordered]@{
    "General Windows files"        = [ordered]@{
        "Compressed files"         = @(".zip", ".rar", ".7z", ".tar.gz")
        "Document files"           = @(".pdf", ".txt", ".rtf", ".odt")
    }
    "Media files"                  = [ordered]@{
        "Video files"              = @(".mp4", ".avi", ".mov", ".wmv", ".mkv", ".flv", ".mpeg", ".mpg", ".3gp", ".webm", ".vob", ".mts")
        "Image files"              = @(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif", ".svg", ".webp", ".heic")
        "Audio files"              = @(".mp3", ".wav", ".aiff", "aac", ".ogg", ".flac", ".alac", ".wma", ".pcm")
    }

    "Installation files"           = [ordered]@{
        "Executable files"         = @(".exe", ".bat", ".cmd", ".msi", ".com", ".jar", ".pi", ".pl")
        "System files"             = @(".dll", ".sys", ".ini")
    }

    "3D models"                    = [ordered]@{
        "General files"            = @(".obj", ".3ds", ".ply", ".blend")
        "3D printing files"        = @(".stl", ".amf", ".3mf")
        "Game development files"   = @(".fbx", ".dae", ".gltf", ".x3d")
        "CAD files"                = @(".step", ".iges", ".dwg", "dxf")
        "All 3d files"             = @(".obj", ".3ds", ".ply", ".blend", ".stl", ".amf", ".3mf", ".fbx", ".dae", ".gltf", ".x3d", ".step", ".iges", ".dwg", ".dxf")
    }

    "Microsoft Office files"       = [ordered]@{
        "Word documents"           = @(".doc", ".docx")
        "Powerpoint presentations" = @(".ptt", ".pttx", ".pps", ".ppsx")
        "Excel sheets"             = @(".xls", ".xlsx", ".xlsm", ".xlsb")
    }
    "Imported files"               = [ordered]@{ <# Import file extensions which aren't already implemented
        "example"                 = @(".exm", ".exam", ".mpl") #>
    }
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
$dropDownMenu.Size = New-Object System.Drawing.Size(250, 50)
$dropDownMenu.Location = New-Object System.Drawing.Point(20, 50)

foreach ($fileExtension in $fileTypes.Keys) {
    [void]$dropDownMenu.Items.Add($fileExtension)
}
$form.Controls.Add($dropDownMenu)

#This creates a simple text label on top of the drop down menu.
$textLabelDDM = New-Object System.Windows.Forms.Label
$textLabelDDM.Size = New-Object System.Drawing.Size($dropDownMenu.Size)
$textLabelDDM.Location = New-Object System.Drawing.Point($dropDownMenu.Location.X, ($dropDownMenu.Location.Y - 30))
$textLabelDDM.Text = "Filter file types:"
$textLabelDDM.Font = New-Object System.Drawing.Font("Arial", 10)
$textLabelDDM.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

#This creates the second drop down menu.
$selectFileExtensions = New-Object System.Windows.Forms.ComboBox
$selectFileExtensions.Size = New-Object System.Drawing.Size($dropDownMenu.ClientSize.Width, $dropDownMenu.ClientSize.Height)
$selectFileExtensions.Location = New-Object System.Drawing.Point(($dropDownMenu.Location.X + $dropDownMenu.Size.Width + 20), $dropDownMenu.Location.Y)
$form.Controls.Add($selectFileExtensions)

$textLabelSFE = New-Object System.Windows.Forms.Label
$textLabelSFE.Size = New-Object System.Drawing.Size($selectFileExtensions.Size)
$textLabelSFE.Location = New-Object System.Drawing.Point($selectFileExtensions.Location.X, ($selectFileExtensions.Location.Y - 30))
$textLabelSFE.Text = "Filter file extensions:"
$textLabelSFE.Font = New-Object System.Drawing.Font("Arial", 10)
$textLabelSFE.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter


#This creates a simple text label on top of the drop down menu.
$textLabelDDM = New-Object System.Windows.Forms.Label
$textLabelDDM.Size = New-Object System.Drawing.Size($dropDownMenu.Size)
$textLabelDDM.Location = New-Object System.Drawing.Point($dropDownMenu.Location.X, ($dropDownMenu.Location.Y - 30))
$textLabelDDM.Text = "Filter file types:"
$textLabelDDM.Font = New-Object System.Drawing.Font("Arial", 10)
$textLabelDDM.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

#This creates the error label, which tells the user what went wrong or if it succeeded.
$errorLabel = New-Object System.Windows.Forms.Label
$errorLabel.Size = New-Object System.Drawing.Size(550, 15)
$errorLabel.Location = New-Object System.Drawing.Point(($form.Size.Width / 2 - $errorLabel.Size.Width / 2), ($okayButton.Location.Y - 20))
$errorLabel.Text = ""
$errorLabel.Font = New-Object System.Drawing.Font("Arial", 7)
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

function DropDownMenuFunctionality {
    param (
        $Object,
        $Selectedoptions
    )
        
    $selected = $fileTypes[$Selectedoptions]
    $items = New-Object System.Collections.ArrayList
    foreach ($item in $Object.Items) {
        $items.Add($item)
    }

    if ($Object.Items.Length -ne 0) {
        foreach ($item in $items) {
            try {
                [void]$Object.Items.Remove($item)
            }
            catch {
                Write-Error "Error: $_"
            }
        }
    }

    ForEach ($option in $selected.Keys) {
        try {
            [void]$Object.Items.Add($option)
        }
        catch {
            Write-Error "Error: $_"
        }
        finally {
            $form.Controls.Add($Object)
        }
    }
    return $newObject
}

#Function used for logging errors
function ErrorLogging {
    param (
        [string]$errorMSG,
        [string]$errorDesc,
        [string]$state
    )
    if ( $errorDesc ) {
        Write-Error $errorDesc 
    }
    if ($state -eq "Error") {
        $errorLabel.ForeColor = [System.Drawing.Color]::Red
    }
    elseif ($state -eq "Success") {
        $errorLabel.ForeColor = [System.Drawing.Color]::Green
    }
    $errorLabel.Text = $errorMSG
}

#Gives the "Okay button" functionality.
function MoveFilesClick {
    param (
    )
    $itemsTransfered = 0
    $itemsFailed = 0

    if ($fromFolder.Text -eq "") { return ErrorLogging -errorMSG "Origin folder cannot remain empty!" -errorDesc "Origin directory cannot remain empty!" -state "Error" }
    if ($null -eq $selectFileExtensions.SelectedItem) { return ErrorLogging -errorMSG "Please select file extension" -errorDesc "Combobox does not have a selected item" -state "Error" }

    if ($fileTypes.Values.Contains($selectFileExtensions.SelectedItem)) {
        $selected = $fileTypes[$dropDownMenu.SelectedItem]
        $extensions = $selected[$selectFileExtensions.SelectedItem]
        $files = Get-ChildItem -Path $fromFolder.Text | Where-Object { $extensions -contains $_.Extension}

        if ($destinationFolder.Text -eq "") { return ErrorLogging -errorMSG "Destination directory cannot remain empty!" -errorDesc "Destination directory cannot remain empty!" -state "Error" }
        if (-Not (Test-Path -Path $destinationFolder.Text -PathType Container)) { return ErrorLogging -errorMSG "'$($destinationFolder.Text)' is an invalid directory!" -errorDesc "'$($destinationFolder.Text)' is an invalid directory!" -state "Error" }
    
        foreach ($file in $files) {

            try {
                Move-Item -Path $file.FullName -Destination $destinationFolder.Text -ErrorAction Stop
                $itemsTransfered++
            }
            catch {
                Write-Error "Error: $_"
                $itemsFailed++
            }
            finally {
                #Error handling
                ErrorLogging -errorMSG "Files successfully moved: $itemsTransfered, Files failed: $itemsFailed" -state "Success"
            }
        }
    } 
}

#Fires when the selected item of the first dropdown menu changes.
$dropDownMenu.Add_SelectedIndexChanged({
    DropDownMenuFunctionality -Object $selectFileExtensions -selectedOptions $dropDownMenu.SelectedItem
})

#Fires when the okay button has been clicked
$okayButton.Add_Click({
    MoveFilesClick
})

#Adding more controls
$form.Controls.Add($okayButton)
$form.Controls.Add($textLabelDDM)
$form.Controls.Add($errorLabel)
$form.Controls.Add($textLabelSFE)

#Finally so the GUI is actually visible.
$form.TopMost = $true
$form.ShowDialog()
