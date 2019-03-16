# PDFValidation
Encrypt PDF files and send them by Email.
There are two folders, one with an English based form and the other with an Hebrew based form.

Extract the folders and save the "PDFpassprotect" folder in the location you choose, the default is c:\temp but you can change it inside the code. 

In this script the password is the Client phone number - in my case it was the best option but you can change it to what ever you need. 
You will need to make some adjustments for your environment, look for the variables at the beginning of the script


#Requires:
1. SMTP server
2  Download Itextsharp.dll and place it in the pdfPasswordProtect
 Download it from - https://github.com/itext/itextsharp
 https://github.com/WolfeReiter/iTextSharp/blob/master/README
 iTextPdf is licensed under AGPL - More information: https://itextpdf.com/en/search?query=itextsharp.dll

4. Save the pdfPasswordProtect folder  under c:\temp or change the location in the code under $mtpath

*************************************************
Windows 7 works only with old versions of itextsharp.dll, I had to download some verions and test them on win7 and win10 before it worked.
*************************************************


Have fun :)
 
#------------------------------------------------------------
# Copyright #Idit Bnaya  #Itext  All rights reserved.
#------------------------------------------------------------      

DESCRIPTION 
This script protect a PDF document with a password and send it by email using a PowerShell form

NOTES
Written by: Idit Bnaya 02/2019

Find me on:

* My Blog:	https://itblog.bnaya.co.il
* LinkedIn:	https://www.linkedin.com/in/idit-bnaya/                    

