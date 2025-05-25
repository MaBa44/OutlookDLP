Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms

Public Class EncryptionForm
    Inherits Form

    Public Property Password As String
    Public Property PasswordDescription As String

    Public Sub New()
        InitializeComponent()
        
        ' Apply localization
        ApplyLocalization()
    End Sub

    Private Sub ApplyLocalization()
        ' Set form title and labels from resources
        Me.Text = LocalizationManager.GetString("EncryptTitle")
        lblPassword.Text = LocalizationManager.GetString("PasswordLabel")
        lblPasswordDescription.Text = LocalizationManager.GetString("PasswordDescriptionLabel")
        btnOK.Text = "OK"
        btnCancel.Text = LocalizationManager.GetString("ButtonCancel")
    End Sub

    Private Sub InitializeComponent()
        Me.lblPassword = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.lblPasswordDescription = New System.Windows.Forms.Label()
        Me.txtPasswordDescription = New System.Windows.Forms.TextBox()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        ' lblPassword
        '
        Me.lblPassword.AutoSize = True
        Me.lblPassword.Location = New System.Drawing.Point(12, 15)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(56, 13)
        Me.lblPassword.TabIndex = 0
        Me.lblPassword.Text = "Password:"
        '
        ' txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(142, 12)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = "*"c
        Me.txtPassword.Size = New System.Drawing.Size(230, 20)
        Me.txtPassword.TabIndex = 1
        '
        ' lblPasswordDescription
        '
        Me.lblPasswordDescription.AutoSize = True
        Me.lblPasswordDescription.Location = New System.Drawing.Point(12, 41)
        Me.lblPasswordDescription.Name = "lblPasswordDescription"
        Me.lblPasswordDescription.Size = New System.Drawing.Size(124, 13)
        Me.lblPasswordDescription.TabIndex = 2
        Me.lblPasswordDescription.Text = "Password Description:"
        '
        ' txtPasswordDescription
        '
        Me.txtPasswordDescription.Location = New System.Drawing.Point(142, 38)
        Me.txtPasswordDescription.Name = "txtPasswordDescription"
        Me.txtPasswordDescription.Size = New System.Drawing.Size(230, 20)
        Me.txtPasswordDescription.TabIndex = 3
        '
        ' btnOK
        '
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnOK.Location = New System.Drawing.Point(142, 74)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(110, 30)
        Me.btnOK.TabIndex = 4
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        AddHandler Me.btnOK.Click, AddressOf Me.btnOK_Click
        '
        ' btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(262, 74)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(110, 30)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        ' EncryptionForm
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(384, 116)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.txtPasswordDescription)
        Me.Controls.Add(Me.lblPasswordDescription)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.lblPassword)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "EncryptionForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Encrypt Email"
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs)
        ' Validate inputs
        If String.IsNullOrWhiteSpace(txtPassword.Text) Then
            MessageBox.Show(
                LocalizationManager.GetString("PasswordRequired"),
                LocalizationManager.GetString("ValidationError"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
            txtPassword.Focus()
            Me.DialogResult = DialogResult.None
            Return
        End If

        If String.IsNullOrWhiteSpace(txtPasswordDescription.Text) Then
            MessageBox.Show(
                LocalizationManager.GetString("PasswordDescriptionRequired"),
                LocalizationManager.GetString("ValidationError"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
            txtPasswordDescription.Focus()
            Me.DialogResult = DialogResult.None
            Return
        End If

        ' Set properties
        Password = txtPassword.Text
        PasswordDescription = txtPasswordDescription.Text
    End Sub

    Private WithEvents lblPassword As Label
    Private WithEvents txtPassword As TextBox
    Private WithEvents lblPasswordDescription As Label
    Private WithEvents txtPasswordDescription As TextBox
    Private WithEvents btnOK As Button
    Private WithEvents btnCancel As Button
End Class