# ExportLabbook
OneNote labbook export and backup tool

You need to download the pdfsharp dll file from here: https://sourceforge.net/projects/pdfsharp/ and place it in the same folder than the export script. 

OneNote must run and a user with valid reading credentials on the notebook in question must be logged in. 

To make the magic happen, you need to install the onenote-powershell adaptor. You can get the *.msi file for that from: http://cid-83dfdc4530b17dd8.skydrive.live.com/self.aspx/Open%20Source%20Projects/OneNotePowershell.msi

More details about the provider can be found here: https://bdewey.com/2007/07/18/onenote-powershell-provider/

The script now checks the date of the most recently exported summary file and exports all pages created after this date. 

Furthermore the Schtasks command is used to schedule the execution of the script. In order to make this work you have to execute the following command once: 

`Schtasks /create /sc weekly /d sun /tn "Export Labbook" /tr "PowerShell -File [Path to Script]\ExportLabbook.ps1" /st 23:00:00`

This will create a scheduled task called "Export Labbook" which will run every Sunday 23:00.