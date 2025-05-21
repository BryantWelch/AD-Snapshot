#region Output Options Tab
# Create File Options Group Box
$grpFileOptions = New-Object System.Windows.Forms.GroupBox
$grpFileOptions.Text = "File Options"
$grpFileOptions.Location = New-Object System.Drawing.Point(10, 20)
$grpFileOptions.Size = New-Object System.Drawing.Size(740, 150)
$tabOutput.Controls.Add($grpFileOptions)

# Create File
$chkCreateFile = New-Object System.Windows.Forms.CheckBox
$chkCreateFile.Text = "Save Report as File"
$chkCreateFile.Location = New-Object System.Drawing.Point(20, 25)
$chkCreateFile.Size = New-Object System.Drawing.Size(200, 20)
$chkCreateFile.Checked = $true
$toolTip.SetToolTip($chkCreateFile, "When checked, the report will be saved as an HTML file. This is recommended for sortable user and computer lists.")
$grpFileOptions.Controls.Add($chkCreateFile)

# Output Path
$lblOutputPath = New-Object System.Windows.Forms.Label
$lblOutputPath.Text = "Output Path (blank = Desktop):"
$lblOutputPath.Location = New-Object System.Drawing.Point(20, 55)
$lblOutputPath.Size = New-Object System.Drawing.Size(200, 20)
$grpFileOptions.Controls.Add($lblOutputPath)

$txtOutputPath = New-Object System.Windows.Forms.TextBox
$txtOutputPath.Location = New-Object System.Drawing.Point(220, 55)
$txtOutputPath.Size = New-Object System.Drawing.Size(400, 20)
$toolTip.SetToolTip($txtOutputPath, "Specify the folder where reports will be saved. If left blank, reports will be saved to your desktop.")
$grpFileOptions.Controls.Add($txtOutputPath)

$btnBrowseOutputPath = New-Object System.Windows.Forms.Button
$btnBrowseOutputPath.Text = "..."
$btnBrowseOutputPath.Location = New-Object System.Drawing.Point(630, 55)
$btnBrowseOutputPath.Size = New-Object System.Drawing.Size(30, 20)
$grpFileOptions.Controls.Add($btnBrowseOutputPath)

# PDF Options
$chkWantPDFFile = New-Object System.Windows.Forms.CheckBox
$chkWantPDFFile.Text = "Create PDF Report"
$chkWantPDFFile.Location = New-Object System.Drawing.Point(20, 85)
$chkWantPDFFile.Size = New-Object System.Drawing.Size(200, 20)
$chkWantPDFFile.Checked = $false
$toolTip.SetToolTip($chkWantPDFFile, "When checked, a PDF version of the report will be created. Requires wkhtmltopdf to be installed.")
$grpFileOptions.Controls.Add($chkWantPDFFile)

# PDF Converter Path
$lblPDFConverter = New-Object System.Windows.Forms.Label
$lblPDFConverter.Text = "PDF Converter Path:"
$lblPDFConverter.Location = New-Object System.Drawing.Point(20, 115)
$lblPDFConverter.Size = New-Object System.Drawing.Size(200, 20)
$grpFileOptions.Controls.Add($lblPDFConverter)

$txtPDFConverter = New-Object System.Windows.Forms.TextBox
$txtPDFConverter.Location = New-Object System.Drawing.Point(220, 115)
$txtPDFConverter.Size = New-Object System.Drawing.Size(400, 20)
$txtPDFConverter.Text = "c:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
$toolTip.SetToolTip($txtPDFConverter, "Path to the wkhtmltopdf executable. Required for PDF generation. Download from http://wkhtmltopdf.org/downloads.html")
$grpFileOptions.Controls.Add($txtPDFConverter)

$btnBrowsePDFConverter = New-Object System.Windows.Forms.Button
$btnBrowsePDFConverter.Text = "..."
$btnBrowsePDFConverter.Location = New-Object System.Drawing.Point(630, 115)
$btnBrowsePDFConverter.Size = New-Object System.Drawing.Size(30, 20)
$grpFileOptions.Controls.Add($btnBrowsePDFConverter)

# Email Options Group Box
$grpEmailOptions = New-Object System.Windows.Forms.GroupBox
$grpEmailOptions.Text = "Email Options"
$grpEmailOptions.Location = New-Object System.Drawing.Point(10, 180)
$grpEmailOptions.Size = New-Object System.Drawing.Size(740, 200)
$tabOutput.Controls.Add($grpEmailOptions)

# Send Email
$chkSendEmail = New-Object System.Windows.Forms.CheckBox
$chkSendEmail.Text = "Send Email Report"
$chkSendEmail.Location = New-Object System.Drawing.Point(20, 25)
$chkSendEmail.Size = New-Object System.Drawing.Size(200, 20)
$chkSendEmail.Checked = $false
$toolTip.SetToolTip($chkSendEmail, "When checked, the report will be emailed to the specified recipients.")
$grpEmailOptions.Controls.Add($chkSendEmail)

# From Email
$lblFromEmail = New-Object System.Windows.Forms.Label
$lblFromEmail.Text = "From Email:"
$lblFromEmail.Location = New-Object System.Drawing.Point(20, 55)
$lblFromEmail.Size = New-Object System.Drawing.Size(200, 20)
$grpEmailOptions.Controls.Add($lblFromEmail)

$txtFromEmail = New-Object System.Windows.Forms.TextBox
$txtFromEmail.Location = New-Object System.Drawing.Point(220, 55)
$txtFromEmail.Size = New-Object System.Drawing.Size(400, 20)
$txtFromEmail.Text = "user1@example.com"
$toolTip.SetToolTip($txtFromEmail, "The email address that will appear as the sender of the report.")
$grpEmailOptions.Controls.Add($txtFromEmail)

# To Email
$lblToEmail = New-Object System.Windows.Forms.Label
$lblToEmail.Text = "To Email:"
$lblToEmail.Location = New-Object System.Drawing.Point(20, 85)
$lblToEmail.Size = New-Object System.Drawing.Size(200, 20)
$grpEmailOptions.Controls.Add($lblToEmail)

$txtToEmail = New-Object System.Windows.Forms.TextBox
$txtToEmail.Location = New-Object System.Drawing.Point(220, 85)
$txtToEmail.Size = New-Object System.Drawing.Size(400, 20)
$txtToEmail.Text = "user2@example.com"
$toolTip.SetToolTip($txtToEmail, "The primary recipient email address for the report.")
$grpEmailOptions.Controls.Add($txtToEmail)

# CC List
$lblCcList = New-Object System.Windows.Forms.Label
$lblCcList.Text = "CC List (comma separated):"
$lblCcList.Location = New-Object System.Drawing.Point(20, 115)
$lblCcList.Size = New-Object System.Drawing.Size(200, 20)
$grpEmailOptions.Controls.Add($lblCcList)

$txtCcList = New-Object System.Windows.Forms.TextBox
$txtCcList.Location = New-Object System.Drawing.Point(220, 115)
$txtCcList.Size = New-Object System.Drawing.Size(400, 20)
$toolTip.SetToolTip($txtCcList, "Additional recipients for the report email. For multiple recipients, use comma-separated email addresses.")
$grpEmailOptions.Controls.Add($txtCcList)

# SMTP Server
$lblSmtpServer = New-Object System.Windows.Forms.Label
$lblSmtpServer.Text = "SMTP Server:"
$lblSmtpServer.Location = New-Object System.Drawing.Point(20, 145)
$lblSmtpServer.Size = New-Object System.Drawing.Size(200, 20)
$grpEmailOptions.Controls.Add($lblSmtpServer)

$txtSmtpServer = New-Object System.Windows.Forms.TextBox
$txtSmtpServer.Location = New-Object System.Drawing.Point(220, 145)
$txtSmtpServer.Size = New-Object System.Drawing.Size(400, 20)
$txtSmtpServer.Text = "smtp.example.org"
$toolTip.SetToolTip($txtSmtpServer, "The SMTP server that will be used to send the email report. Example: smtp.office365.com")
$grpEmailOptions.Controls.Add($txtSmtpServer)

# Attach PDF to Email
$chkAttachPDF = New-Object System.Windows.Forms.CheckBox
$chkAttachPDF.Text = "Attach PDF to Email (requires PDF option enabled)"
$chkAttachPDF.Location = New-Object System.Drawing.Point(20, 175)
$chkAttachPDF.Size = New-Object System.Drawing.Size(350, 20)
$chkAttachPDF.Checked = $false
$toolTip.SetToolTip($chkAttachPDF, "When checked, the PDF report will be attached to the email. Requires the 'Create PDF Report' option to be enabled.")
$grpEmailOptions.Controls.Add($chkAttachPDF)
