 <# 
 
#------------------------------------------------------------
# Copyright Idit Bnaya  All rights reserved.
#------------------------------------------------------------      

.DESCRIPTION 
This script protect a PDF document with a password and send it by email using a ps form


.NOTES
Written by: Idit Bnaya 02/2019

Find me on:

* My Blog:	https://itblog.bnaya.co.il
* LinkedIn:	https://www.linkedin.com/in/idit-bnaya/                    

.requires
1. Smtp server
2. Itextsharp.dll - https://sourceforge.net/projects/itextsharp
3. Save the pdfPasswordProtect folder  under c:\temp or change the location in the code under $mtpath

#>
 


##### Change The following: ######

$Mysubject = "Encrypted pdf file"
$FromEmailAddress = "EncryptPDF@Domain.com"
$MySmtp = "Smtp@domain.com"
$mypath = "C:\temp" #only if you save the pdfPasswordProtect folder on a diferrent location


##### Load itextsharp.dll ######

Set-ExecutionPolicy -Scope Process -ExecutionPolicy unrestricted -force
$DllPath = "$mypath\pdfPasswordProtect\itextsharp.dll"
[System.Reflection.Assembly]::LoadFrom($DllPath)

$protectPDF = 0



 
Function Get-FileName($initialDirectory)
{   
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
    Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "PDF Files (*.pdf)| *.pdf"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
} 

Function Send-Email($outFile, $emailAddress)
{
    $mconsole.Text=""
    
    if ($protectPDF -eq 0)
    {
        try
        {
            #Send-email logic here: protected file name is in $outFile
            ############################################################################### 
 
            ###### Define Variables ######
 
            $fromaddress = $FromEmailAddress 
            $toaddress = $emailAddress 
            #$bccaddress = "" 
            #$CCaddress = "" 
            $Subject = "$Mysubject" 
            $body = get-content C:\temp\pdfPasswordProtect\MailBody.html
            $attachment = $outFile
            $smtpserver = $MySmtp 
 
            ##############################
 
            $message = new-object System.Net.Mail.MailMessage 
            $message.From = $fromaddress 
            $message.To.Add($toaddress) 
            #$message.CC.Add($CCaddress) 
            #$message.Bcc.Add($bccaddress) 
            $message.IsBodyHtml = $True
            $message.Subject = $Subject 
            $attach = new-object Net.Mail.Attachment($attachment)
            $message.Attachments.Add($attach) 
            $message.body = $body
            $smtp = new-object Net.Mail.SmtpClient($smtpserver) 
            $smtp.Send($message)
 
            ##############################
            $mfile_input.Text=""
            $mphone_input.Text=""
            $memail_input.Text=""
                     $mconsole.Text += "Success!"

            #Remove-Item $outFile -ErrorAction Stop
        }
        catch [System.Management.Automation.ActionPreferenceStopException]
        {
            $mconsole.Text = "Error Deleting file"
        }
        catch
        {
            $mconsole.Text = "Error - Mail Did not send"
        }

        $protectPDF = 0
    } else {
        $mconsole.Text = "PDF Encrypt failed"
    }
} 

Function Password-ProtectPDF($fileName, $phoneNumber)
{
    try
    {
        $file = New-Object System.IO.FileInfo $fileName
        $outFile = $env:TEMP + "\" + $file.Name
        $fileWithPassword = New-Object System.IO.FileInfo $outFile
        $password = $phoneNumber
        $fileStreamIn = $file.OpenRead()
        $fileStreamOut = New-Object System.IO.FileStream($fileWithPassword.FullName,[System.IO.FileMode]::Create,[System.IO.FileAccess]::Write,[System.IO.FileShare]::None)
        $reader = New-Object iTextSharp.text.pdf.PdfReader $fileStreamIn
        [iTextSharp.text.pdf.PdfEncryptor]::Encrypt($reader, $fileStreamOut, $true, $password, $password, [iTextSharp.text.pdf.PdfWriter]::ALLOW_PRINTING)
        $protectPDF = 1
        return $outFile
    }
    catch
    {
        $protectPDF = 0
    }
}


     
Add-Type -AssemblyName System.Windows.Forms 
Add-Type -AssemblyName System.Drawing 
$MyForm = New-Object System.Windows.Forms.Form 
$MyForm.Text="Encrypt PDF" 
$MyForm.Size = New-Object System.Drawing.Size(320,360) 
$MyForm.RightToLeft="Yes"
$myform.RightToLeftLayout="Yes"
     
$img = [System.Drawing.Image]::Fromfile('C:\temp\pdfPasswordProtect\image.png')

#create resized bitmap
$imageSize = 70
$newWidth = $imageSize
$newHeight = $imageSize
$bmpResized = New-Object System.Drawing.Bitmap($newWidth, $newHeight)
$graph = [System.Drawing.Graphics]::FromImage($bmpResized)
$graph.Clear([System.Drawing.Color]::White)
$graph.DrawImage($img,0,0 , $newWidth, $newHeight)

$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Width = $imageSize
$pictureBox.Height = $imageSize
$pictureBox.Image = $bmpResized
$pictureBox.Top="80" 
$pictureBox.Left="20" 
$myForm.controls.add($pictureBox)


 #first input box
 #pdf label
 $mfile_label = New-Object System.Windows.Forms.Label 
        $mfile_label.Text="Select an unprotected PDF file" 
        $mfile_label.Top="20" 
        $mfile_label.Left ="140" 
        $mfile_label.Anchor="right,Top" 
$mfile_label.Size = New-Object System.Drawing.Size(200,20) 
$MyForm.Controls.Add($mfile_label) 




#input  file
$mfile_button = New-Object System.Windows.Forms.Button 
        $mfile_button.Text="Browse"
        $mfile_button.Top="39" 
        $mfile_button.Left="19" 
        $mfile_button.Anchor="Left,Top"
        $mfile_button.Add_Click(
        {    
		    $mfile_input.Text = Get-FileName
        })

$mfile_button.Size = New-Object System.Drawing.Size(100,23) 
$MyForm.Controls.Add($mfile_button)


$mfile_input = New-Object System.Windows.Forms.TextBox 
        $mfile_input.Text="" 
        $mfile_input.Top="40" 
        $mfile_input.Left="120" 
        $mfile_input.Anchor="Left,Top" 
      
$mfile_input.Size = New-Object System.Drawing.Size(170,23) 
$MyForm.Controls.Add($mfile_input) 

     
## Email Settings
$memail_label = New-Object System.Windows.Forms.Label 
        $memail_label.Text="Enter a valid Email address" 
        $memail_label.Top="80" 
        $memail_label.Left="150" 
        $memail_label.Anchor="Left,Top" 
$memail_label.Size = New-Object System.Drawing.Size(160,23) 
$MyForm.Controls.Add($memail_label) 
         

$memail_input = New-Object System.Windows.Forms.TextBox 
        $memail_input.Text="" 
        $memail_input.Top="105" 
        $memail_input.Left="120" 
        $memail_input.Anchor="Left,Top" 
$memail_input.Size = New-Object System.Drawing.Size(150,23) 
$MyForm.Controls.Add($memail_input)
         
  
#phone Label 
 $mphone_label = New-Object System.Windows.Forms.Label 
        $mphone_label.Text="Enter A phone number          (will be used to open the file)"
        $mphone_label.RightToLeft="no"
        $mphone_label.Top="140" 
        $mphone_label.Left="140" 
        $mphone_label.Anchor="Left,Top" 
$mphone_label.Size = New-Object System.Drawing.Size(150,23) 
$MyForm.Controls.Add($mphone_label)

$mphone_input = New-Object System.Windows.Forms.TextBox 
        $mphone_input.Text="" 
        $mphone_input.Top="172" 
        $mphone_input.Left="120" 
        $mphone_input.Anchor="Left,Top" 
$mphone_input.Size = New-Object System.Drawing.Size(150,23) 
$MyForm.Controls.Add($mphone_input)


$mconsole = New-Object System.Windows.Forms.TextBox 
        $mconsole.Text="" 
        $mconsole.Top="200" 
        $mconsole.Left="10" 
        $mconsole.Multiline="true"
        $mconsole.Anchor="Left,Top" 
$mconsole.Size = New-Object System.Drawing.Size(260,92) 
$MyForm.Controls.Add($mconsole) 
        

$msend = New-Object System.Windows.Forms.Button 
        $msend.Text="Send Email" 
        $msend.Top="169" 
        $msend.Left="19" 
        $msend.Anchor="Left,Top" 
        $msend.Add_Click(
        {   
            $errorStack = @()
            $fileName = ""
            $phoneNumber = ""
            $emailAddress = ""
            
            #Validate phone number
            If($mphone_input.Text -match "[0-9]") # Valid Phone number)
            {
                $phoneNumber = $mphone_input.Text
            }
            Else # Invalid Phone number
            {
                $errorStack += "Only numbers are allowed `r`n`r`n"
            }#>

            # Validate file
            If($mfile_input.Text -match ".\.pdf$") # Valid File name (ends with.pdf)
            {
                $fileName = $mfile_input.Text
            }
            Else # Invalid file
            {
                $errorStack += "Error select pdf file `r`n pdf only `r`n`r`n"
            }

            # Validate E-mail
            $mailRegexp = '^(?:[a-z0-9!#$%&'+"'"+'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'+"'"+'*+/=?^_`{|}~-]+)*|"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])$'
            If($memail_input.Text -match $mailRegexp) # Valid E-mail address
            {
                $emailAddress = $memail_input.Text
            }
            Else # Invalid E-mail
            {
                $errorStack += "The email Address is not valid `r`n`r`n"
            }

            # Display errors
            if ($errorStack.Length -ne 0)
            {
                [System.Windows.Forms.MessageBox]::Show($errorStack , "Error")
            }

            # Sending the mail if all the needed data is available
            if (($phoneNumber -ne "") -and ($fileName -ne "") -and ($emailAddress -ne ""))
            {
                $protectPDF = 0
                $mconsole.Text += "Creating PDF encryted file: `r`n" + $outFile + "`r`n"
                $outFile = Password-ProtectPDF -fileName $fileName -phoneNumber $phoneNumber
		        Send-Email -outFile $outFile -emailAddress $emailAddress
            }
        })

       
$msend.Size = New-Object System.Drawing.Size(100,23) 
$MyForm.Controls.Add($msend) 
$MyForm.ShowDialog()