 <# 
 
#------------------------------------------------------------
# Copyright Idit Bnaya  All rights reserved.
#------------------------------------------------------------      

.DESCRIPTION 
This script protect a PDF document with a password and send it by email using a ps form



.NOTES
Written by: Idit Bnaya 03/2019 - מותאם לעברית בעקבות דרישה :)

Find me on:

* My Blog:	https://itblog.bnaya.co.il
* LinkedIn:	https://www.linkedin.com/in/idit-bnaya/  

 

iTextPdf is licensed under AGPL
Download it from - https://github.com/itext/itextsharp
More information: https://itextpdf.com/en/search?query=itextsharp.dll             

.requires
1. SMTP server
2  Download Itextsharp.dll and place it in the pdfPasswordProtect
3. Itextsharp.dll - https://sourceforge.net/projects/itextsharp
   https://github.com/WolfeReiter/iTextSharp/blob/master/README
4. Save the pdfPasswordProtect folder  under c:\temp or change the location in the code under $mtpath

#>
 


##### Change The following: ######

$Mysubject = "מסמך מוצפן"
$FromEmailAddress = "EncryptPDF@Domain.com"
$MySmtp = "Smtp@domain.com"
$mypath = "C:\temp\pdfPasswordProtect" #only if you save the pdfPasswordProtect folder on a diferrent location
$MailPath  = "$mypath\MailHE.html"


##### Load itextsharp.dll ######

Set-ExecutionPolicy -Scope Process -ExecutionPolicy unrestricted -force
$DllPath = "$mypath\itextsharp.dll"
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

Function Send-Email($fileName, $phoneNumber, $emailAddress)
{
    $mconsole.Text=""
    $outFile = Password-ProtectPDF -fileName $fileName -phoneNumber $phoneNumber -emailAddress $emailAddress
    if ($protectPDF -eq 0)
    {
        $mconsole.Text += "יוצר קובץ זמני מוגן בסיסמא: `r`n" + $outFile + "`r`n"

        try
        {
            #Send-email logic here: protected file name is in $outFile
            ############################################################################### 
 
            ##### Define Variables ##### 
 
            $fromaddress = $FromEmailAddress 
            $toaddress = $emailAddress 
            $bccaddress = "" 
            $CCaddress = "" 
            $Subject = "$Mysubject" 
            $body = get-content $MailPath -Encoding utf8
            $attachment = $outFile
            $smtpserver = $MySmtp 
 
            ############################ 
 
            $message = new-object System.Net.Mail.MailMessage 
            $message.From = $fromaddress 
            $message.To.Add($toaddress) 
            $message.CC.Add($CCaddress) 
            $message.Bcc.Add($bccaddress) 
            $message.IsBodyHtml = $True 
            $message.Subject = $Subject 
            $attach = new-object Net.Mail.Attachment($attachment) 
            $message.Attachments.Add($attach) 
            $message.body = $body 
            $smtp = new-object Net.Mail.SmtpClient($smtpserver) 
            $smtp.Send($message) 
 
            ############################

            Remove-Item $outFile -ErrorAction Stop
        }
        catch [System.Management.Automation.ActionPreferenceStopException]
        {
            $mconsole.Text = "תקלה במחיקת הקובץ, פנה לצוות המחשוב לעזרה"
        }
        catch
        {
            $mconsole.Text = "שליחת המייל נכשלה, פנה לצוות המחשוב לעזרה"
        }

        $protectPDF = 0
    } else {
        $mconsole.Text = "הצפנת הקובץ נכשלה, פנה לצוות המחשוב לעזרה"
    }
} 

Function Password-ProtectPDF($fileName, $phoneNumber, $emailAddress)
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
$MyForm.Text="הצפנת PDF" 
$MyForm.Size = New-Object System.Drawing.Size(320,360) 
$MyForm.RightToLeft="Yes"
$myform.RightToLeftLayout="Yes"
     
$img = [System.Drawing.Image]::Fromfile("$mypath\image.png")

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
$pictureBox.Left="200" 
$myForm.controls.add($pictureBox)
 
$mphone_input = New-Object System.Windows.Forms.TextBox 
        $mphone_input.Text="" 
        $mphone_input.Top="40" 
        $mphone_input.Left="20" 
        $mphone_input.Anchor="Left,Top" 
$mphone_input.Size = New-Object System.Drawing.Size(100,23) 
$MyForm.Controls.Add($mphone_input) 
         
 
$mphone_label = New-Object System.Windows.Forms.Label 
        $mphone_label.Text="מספר טלפון של הלקוח:" 
        $mphone_label.RightToLeft="Yes"
        $mphone_label.Top="20" 
        $mphone_label.Left="20" 
        $mphone_label.Anchor="Left,Top" 
$mphone_label.Size = New-Object System.Drawing.Size(150,23) 
$MyForm.Controls.Add($mphone_label) 
         
 
$memail_label = New-Object System.Windows.Forms.Label 
        $memail_label.Text="מייל למשלוח הקובץ:" 
        $memail_label.Top="80" 
        $memail_label.Left="20" 
        $memail_label.Anchor="Left,Top" 
$memail_label.Size = New-Object System.Drawing.Size(160,23) 
$MyForm.Controls.Add($memail_label) 
         
 
$memail_input = New-Object System.Windows.Forms.TextBox 
        $memail_input.Text="" 
        $memail_input.Top="105" 
        $memail_input.Left="20" 
        $memail_input.Anchor="Left,Top" 
$memail_input.Size = New-Object System.Drawing.Size(150,23) 
$MyForm.Controls.Add($memail_input) 
         
 
$mfile_button = New-Object System.Windows.Forms.Button 
        $mfile_button.Text="בחר קובץ..." 
        $mfile_button.Top="159" 
        $mfile_button.Left="195" 
        $mfile_button.Anchor="Left,Top"
        $mfile_button.Add_Click(
        {    
		    $mfile_input.Text = Get-FileName
        })
$mfile_button.Size = New-Object System.Drawing.Size(100,23) 
$MyForm.Controls.Add($mfile_button) 
         
 
$mfile_label = New-Object System.Windows.Forms.Label 
        $mfile_label.Text="בחר קובץ PDF ללא סיסמא:" 
        $mfile_label.Top="140" 
        $mfile_label.Left ="20" 
        $mfile_label.Anchor="right,Top" 
$mfile_label.Size = New-Object System.Drawing.Size(200,20) 
$MyForm.Controls.Add($mfile_label) 
         
 
$mfile_input = New-Object System.Windows.Forms.TextBox 
        $mfile_input.Text="" 
        $mfile_input.Top="160" 
        $mfile_input.Left="20" 
        $mfile_input.Anchor="Left,Top" 
        $mfile_input.Add_Click(
        {    
		    $mfile_input.Text = Get-FileName
        })
$mfile_input.Size = New-Object System.Drawing.Size(170,23) 
$MyForm.Controls.Add($mfile_input) 
         
 
$mconsole = New-Object System.Windows.Forms.TextBox 
        $mconsole.Text="" 
        $mconsole.Top="200" 
        $mconsole.Left="20" 
        $mconsole.Multiline="true"
        $mconsole.Anchor="Left,Top" 
$mconsole.Size = New-Object System.Drawing.Size(260,92) 
$MyForm.Controls.Add($mconsole) 
         
 
$msend = New-Object System.Windows.Forms.Button 
        $msend.Text="שלח מייל" 
        $msend.Top="39" 
        $msend.Left="195" 
        $msend.Anchor="Left,Top" 
        $msend.Add_Click(
        {   
            $errorStack = @()
            $fileName = ""
            $phoneNumber = ""
            $emailAddress = ""
            
            # Validate phone number
            If($mphone_input.Text -match "^0[0-9]{9}$") # Valid Phone number (only 10 numbers with 0 first)
            {
                $phoneNumber = $mphone_input.Text
            }
            Else # Invalid Phone number
            {
                $errorStack += "מספר הטלפון שגוי אנא הכנס מספר תקין עם ספרות בלבד `r`n ללא סימני פיסוק `r`n`r`n"
            }

            # Validate file
            If($mfile_input.Text -match ".\.pdf$") # Valid File name (ends with.pdf)
            {
                $fileName = $mfile_input.Text
            }
            Else # Invalid file
            {
                $errorStack += "לא נבחר קובץ `r`n נא לבחור קובץ מתאים `r`n`r`n"
            }

            # Validate E-mail
            $mailRegexp = '^(?:[a-z0-9!#$%&'+"'"+'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'+"'"+'*+/=?^_`{|}~-]+)*|"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])$'
            If($memail_input.Text -match $mailRegexp) # Valid E-mail address
            {
                $emailAddress = $memail_input.Text
            }
            Else # Invalid E-mail
            {
                $errorStack += "מייל לא תקין `r`n`r`n"
            }

            # Display errors
            if ($errorStack.Length -ne 0)
            {
                [System.Windows.Forms.MessageBox]::Show($errorStack , "תקלה בחלק מהנתונים")
            }

            # Sending the mail if all the needed data is available
            if (($phoneNumber -ne "") -and ($fileName -ne "") -and ($emailAddress -ne ""))
            {
                $protectPDF = 0
		        Send-Email -fileName $fileName -phoneNumber $phoneNumber -emailAddress $emailAddress
            }
        })
$msend.Size = New-Object System.Drawing.Size(100,23) 
$MyForm.Controls.Add($msend) 
$MyForm.ShowDialog()
