#region Output Options Tab
# Create File Options Group Box
$grpFileOptions = New-Object System.Windows.Forms.GroupBox
$grpFileOptions.Text = "File Options"
$grpFileOptions.Location = New-Object System.Drawing.Point(10, 20)
$grpFileOptions.Size = New-Object System.Drawing.Size(740, 105)
$tabOutput.Controls.Add($grpFileOptions)

# --- Output formats ---
$lblFormats = New-Object System.Windows.Forms.Label
$lblFormats.Text = "Generate:"
$lblFormats.Location = New-Object System.Drawing.Point(20, 31)
$lblFormats.Size = New-Object System.Drawing.Size(70, 20)
$grpFileOptions.Controls.Add($lblFormats)

$chkCreateFile = New-Object System.Windows.Forms.CheckBox
$chkCreateFile.Text = "HTML report"
$chkCreateFile.Location = New-Object System.Drawing.Point(95, 30)
$chkCreateFile.Size = New-Object System.Drawing.Size(120, 22)
$chkCreateFile.Checked = $true
$toolTip.SetToolTip($chkCreateFile, "Save the interactive, self-contained HTML report (recommended - supports sortable user and computer lists).")
$grpFileOptions.Controls.Add($chkCreateFile)

$chkWantPDFFile = New-Object System.Windows.Forms.CheckBox
$chkWantPDFFile.Text = "PDF report"
$chkWantPDFFile.Location = New-Object System.Drawing.Point(225, 30)
$chkWantPDFFile.Size = New-Object System.Drawing.Size(110, 22)
$chkWantPDFFile.Checked = $false
$toolTip.SetToolTip($chkWantPDFFile, "Also produce a PDF using your installed Microsoft Edge or Google Chrome. No extra software or configuration is required.")
$grpFileOptions.Controls.Add($chkWantPDFFile)

$chkExportCSV = New-Object System.Windows.Forms.CheckBox
$chkExportCSV.Text = "CSV (users && computers)"
$chkExportCSV.Location = New-Object System.Drawing.Point(345, 30)
$chkExportCSV.Size = New-Object System.Drawing.Size(220, 22)
$chkExportCSV.Checked = $false
$toolTip.SetToolTip($chkExportCSV, "Also export the user and computer lists as CSV files alongside the report. Useful for spreadsheets and auditing.")
$grpFileOptions.Controls.Add($chkExportCSV)

# --- Destination folder ---
$lblOutputPath = New-Object System.Windows.Forms.Label
$lblOutputPath.Text = "Output folder:"
$lblOutputPath.Location = New-Object System.Drawing.Point(20, 70)
$lblOutputPath.Size = New-Object System.Drawing.Size(90, 20)
$grpFileOptions.Controls.Add($lblOutputPath)

$txtOutputPath = New-Object System.Windows.Forms.TextBox
$txtOutputPath.Location = New-Object System.Drawing.Point(115, 68)
$txtOutputPath.Size = New-Object System.Drawing.Size(540, 22)
$toolTip.SetToolTip($txtOutputPath, "Folder where reports are saved. Leave blank to save to your Desktop.")
$grpFileOptions.Controls.Add($txtOutputPath)

$btnBrowseOutputPath = New-Object System.Windows.Forms.Button
$btnBrowseOutputPath.Text = "..."
$btnBrowseOutputPath.Location = New-Object System.Drawing.Point(660, 68)
$btnBrowseOutputPath.Size = New-Object System.Drawing.Size(30, 22)
$grpFileOptions.Controls.Add($btnBrowseOutputPath)

# Email Options Group Box
$grpEmailOptions = New-Object System.Windows.Forms.GroupBox
$grpEmailOptions.Text = "Email Options"
$grpEmailOptions.Location = New-Object System.Drawing.Point(10, 135)
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
$txtFromEmail.Text = ""
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
$txtToEmail.Text = ""
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
$txtSmtpServer.Text = ""
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
