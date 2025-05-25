Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms

Public Class ExternalRecipientsWarningForm
    Inherits Form

    Private _externalRecipients As List(Of String)

    Public Sub New(ByVal externalRecipients As List(Of String))
        InitializeComponent()
        _externalRecipients = externalRecipients
        PopulateRecipientsList()
    End Sub

    Private Sub InitializeComponent()
        Me.lblWarning = New System.Windows.Forms.Label()
        Me.lstRecipients = New System.Windows.Forms.ListBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnEncrypt = New System.Windows.Forms.Button()
        Me.btnSend = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        ' lblWarning
        '
        Me.lblWarning.AutoSize = True
        Me.lblWarning.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWarning.Location = New System.Drawing.Point(12, 9)
        Me.lblWarning.Name = "lblWarning"
        Me.lblWarning.Size = New System.Drawing.Size(450, 30)
        Me.lblWarning.TabIndex = 0
        Me.lblWarning.Text = "The email will be sent to following external recipients:" & vbCrLf & _
                             "Are you sure that your email doesn't contain sensitive information?"
        '
        ' lstRecipients
        '
        Me.lstRecipients.FormattingEnabled = True
        Me.lstRecipients.Location = New System.Drawing.Point(12, 42)
        Me.lstRecipients.Name = "lstRecipients"
        Me.lstRecipients.Size = New System.Drawing.Size(450, 147)
        Me.lstRecipients.TabIndex = 1
        '
        ' btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(12, 205)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(120, 35)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        ' btnEncrypt
        '
        Me.btnEncrypt.DialogResult = System.Windows.Forms.DialogResult.Yes
        Me.btnEncrypt.Location = New System.Drawing.Point(177, 205)
        Me.btnEncrypt.Name = "btnEncrypt"
        Me.btnEncrypt.Size = New System.Drawing.Size(120, 35)
        Me.btnEncrypt.TabIndex = 3
        Me.btnEncrypt.Text = "Encrypt"
        Me.btnEncrypt.UseVisualStyleBackColor = True
        '
        ' btnSend
        '
        Me.btnSend.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnSend.Location = New System.Drawing.Point(342, 205)
        Me.btnSend.Name = "btnSend"
        Me.btnSend.Size = New System.Drawing.Size(120, 35)
        Me.btnSend.TabIndex = 4
        Me.btnSend.Text = "Send"
        Me.btnSend.UseVisualStyleBackColor = True
        '
        ' ExternalRecipientsWarningForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(474, 252)
        Me.Controls.Add(Me.btnSend)
        Me.Controls.Add(Me.btnEncrypt)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.lstRecipients)
        Me.Controls.Add(Me.lblWarning)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ExternalRecipientsWarningForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "External Recipients Warning"
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Private Sub PopulateRecipientsList()
        lstRecipients.Items.Clear()
        For Each recipient In _externalRecipients
            lstRecipients.Items.Add(recipient)
        Next
    End Sub

    Private WithEvents lblWarning As Label
    Private WithEvents lstRecipients As ListBox
    Private WithEvents btnCancel As Button
    Private WithEvents btnEncrypt As Button
    Private WithEvents btnSend As Button
End Class