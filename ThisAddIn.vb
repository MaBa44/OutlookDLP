Imports System
Imports System.IO
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook
Imports System.Text.RegularExpressions
Imports System.Linq
Imports System.Globalization
Imports Microsoft.Office.Tools.Outlook
 
Public Class ThisAddIn
    Inherits Microsoft.Office.Tools.Outlook.AddInBase
    Private WithEvents inspectors As Inspectors
    Private WithEvents mailItem As MailItem
    Private acceptedDomains As List(Of String) = New List(Of String)()

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        ' Initialize localization
        InitializeLocalization()
        
        inspectors = Me.Application.Inspectors
        LoadAcceptedDomains()
    End Sub

    Private Sub InitializeLocalization()
        Try
            ' Get the Outlook UI language
            Dim languageSettings = Me.Application.LanguageSettings
            Dim lcid As Integer = languageSettings.LanguageID(Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI)
            Dim culture As New CultureInfo(lcid)
            
            ' Initialize the localization manager with the Outlook UI culture
            LocalizationManager.Initialize(culture)
        Catch ex As Exception
            ' If there's an error, initialize with default culture
            LocalizationManager.Initialize()
        End Try
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        ' Clean up resources
    End Sub

    Private Sub LoadAcceptedDomains()
        Try
            Dim configPath As String = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "accepted_domains.txt")
            If File.Exists(configPath) Then
                acceptedDomains = File.ReadAllLines(configPath).ToList()
            Else
                ' Create default file with example domains
                Dim defaultDomains As String() = {"yourdomain.com", "company.com", "internal.org"}
                File.WriteAllLines(configPath, defaultDomains)
                acceptedDomains = defaultDomains.ToList()
            End If
        Catch ex As Exception
            MessageBox.Show(
                LocalizationManager.GetString("ConfigurationError", ex.Message),
                LocalizationManager.GetString("ConfigurationErrorTitle"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub inspectors_NewInspector(ByVal Inspector As Inspector) Handles inspectors.NewInspector
        If Inspector.CurrentItem IsNot Nothing AndAlso TypeOf Inspector.CurrentItem Is MailItem Then
            mailItem = TryCast(Inspector.CurrentItem, MailItem)
            AddHandler mailItem.Send, AddressOf mailItem_Send
        End If
    End Sub

    Private Sub mailItem_Send(ByRef Cancel As Boolean)
        ' Check if sender is internal
        If Not IsSenderInternal() Then
            ' If sender is not internal, don't process
            Return
        End If

        ' Get external recipients
        Dim externalRecipients As List(Of String) = GetExternalRecipients(mailItem)

        ' If there are external recipients, show warning
        If externalRecipients.Count > 0 Then
            Cancel = True ' Cancel sending temporarily
            
            ' Show warning form
            Using warningForm As New ExternalRecipientsWarningForm(externalRecipients)
                Dim result As DialogResult = warningForm.ShowDialog()
                
                Select Case result
                    Case DialogResult.Cancel
                        ' User chose Cancel - do nothing, email won't be sent
                        Return
                        
                    Case DialogResult.Yes
                        ' User chose Encrypt
                        Using encryptForm As New EncryptionForm()
                            If encryptForm.ShowDialog() = DialogResult.OK Then
                                ' Encrypt and send
                                EncryptAndSendEmail(mailItem, encryptForm.Password, encryptForm.PasswordDescription)
                            End If
                        End Using
                        
                    Case DialogResult.OK
                        ' User chose Send without encryption
                        mailItem.Send()
                End Select
            End Using
        End If
    End Sub

    Private Function IsSenderInternal() As Boolean
        Try
            Dim senderEmail As String = mailItem.SenderEmailAddress
            If String.IsNullOrEmpty(senderEmail) Then
                senderEmail = Me.Application.Session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
            End If
            
            Dim domain As String = GetDomainFromEmail(senderEmail)
            Return acceptedDomains.Contains(domain.ToLower())
        Catch ex As Exception
            MessageBox.Show("Error checking sender domain: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Function GetExternalRecipients(ByVal mail As MailItem) As List(Of String)
        Dim externalRecipients As New List(Of String)()
        
        ' Check all recipient types (To, CC, BCC)
        For Each recipient As Recipient In mail.Recipients
            Try
                Dim email As String = GetEmailFromRecipient(recipient)
                Dim domain As String = GetDomainFromEmail(email)
                
                If Not String.IsNullOrEmpty(domain) AndAlso Not acceptedDomains.Contains(domain.ToLower()) Then
                    externalRecipients.Add(email)
                End If
            Catch ex As Exception
                ' Skip problematic recipients
            End Try
        Next
        
        Return externalRecipients
    End Function

    Private Function GetEmailFromRecipient(ByVal recipient As Recipient) As String
        Try
            ' Try to get SMTP address
            If recipient.AddressEntry IsNot Nothing AndAlso recipient.AddressEntry.AddressEntryUserType = OlAddressEntryUserType.olExchangeUserAddressEntry Then
                Dim exchangeUser As ExchangeUser = recipient.AddressEntry.GetExchangeUser()
                If exchangeUser IsNot Nothing Then
                    Return exchangeUser.PrimarySmtpAddress
                End If
            End If
            
            ' Fallback to recipient address
            Return recipient.Address
        Catch ex As Exception
            Return recipient.Address
        End Try
    End Function

    Private Function GetDomainFromEmail(ByVal email As String) As String
        Try
            Dim match As Match = Regex.Match(email, "@([^@]+)$")
            If match.Success Then
                Return match.Groups(1).Value.ToLower()
            End If
            Return String.Empty
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    Private Sub EncryptAndSendEmail(ByVal originalMail As MailItem, ByVal password As String, ByVal passwordDescription As String)
        Try
            ' Create a new mail item
            Dim encryptedMail As MailItem = TryCast(Me.Application.CreateItem(OlItemType.olMailItem), MailItem)
            
            ' Copy recipients
            For Each recipient As Recipient In originalMail.Recipients
                encryptedMail.Recipients.Add(recipient.Address)
            Next
            
            ' Set subject with [Encrypted] prefix
            encryptedMail.Subject = LocalizationManager.GetString("EncryptedSubjectPrefix") & originalMail.Subject
            
            ' Set body with encryption notice
            encryptedMail.HTMLBody = LocalizationManager.GetString("EncryptedEmailBody", passwordDescription)
            
            ' Create temporary folder for files to encrypt
            Dim tempFolder As String = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString())
            Directory.CreateDirectory(tempFolder)
            
            Try
                ' Save email body to file
                Dim bodyFileName As String = Path.Combine(tempFolder, "message.html")
                File.WriteAllText(bodyFileName, originalMail.HTMLBody)
                
                ' Save attachments
                For Each attachment As Attachment In originalMail.Attachments
                    attachment.SaveAsFile(Path.Combine(tempFolder, attachment.FileName))
                Next
                
                ' Create encrypted ZIP
                Dim zipFileName As String = Path.Combine(Path.GetTempPath(), "Encrypted_" & Guid.NewGuid().ToString() & ".zip")
                ZipHelper.CreateEncryptedZip(tempFolder, zipFileName, password)
                
                ' Add ZIP as attachment
                encryptedMail.Attachments.Add(zipFileName)
                
                ' Send the email
                encryptedMail.Send()
            Finally
                ' Clean up temporary files
                Try
                    If Directory.Exists(tempFolder) Then
                        Directory.Delete(tempFolder, True)
                    End If
                Catch
                    ' Ignore cleanup errors
                End Try
            End Try
        Catch ex As Exception
            MessageBox.Show(
                LocalizationManager.GetString("EncryptionError", ex.Message),
                LocalizationManager.GetString("EncryptionErrorTitle"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub
End Class