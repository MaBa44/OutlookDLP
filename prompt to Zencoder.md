This is the prompt to Zencoder which was used to generate he application.

Extend the sample Outlook VSTO add-in code (using VB.NET) to perform the following functions before sending an email:
1. checks, if any of email's recipiens is an external recipient (the definition of an external recipient is given below)
2. if no - send an email
3. if yes - display warning "The email will be sent to following external recipients: <list of external recipients>". Below the list should be displayed text: "Are you sure that your email doesn't contain sensitive information? If so, please click on "Cancel" button and then remove confidential information or encrypt the message by clicking on "Encrypt" button." Below the text three buttons should be placed, each one with its own description: button "Cancel" with description "to go back to editing the message", button "Encrypt" with description "to create an ecrypted email", button "Send" with description "to send en unencrypted message".
4. if user chooses "Encrypt" option - ask user for a <password> and <password description> and then create an encrypted ZIP archive containing the message body and all attachments. <password description> cannot be empty.Then send an email containing this ZIP archive as an attachment. The subject line of the encrypted email should start with "[Encrypted]" prefix. The message body of the email should contain the text "This email contains confidential information. It was encrypted to ensure, that only authorized persons can access the information contained within this email. The passwort for decryption is:" <password description>.
5. if user chooses "Send" option - display a confirmation dialog "You are going to send an unencrypted message. Are you positive, that doing so you are not violating data protection rules?". After confirmation send an email without encryption. After decline get back to previous dialog.
6. if user chooses "Cancel" option - do not send an email and return to mail compose window.

Definition of an external recipient:
- a recipient with domain name not found in a list of accepted domains stored in a configuration file called "accepted_domains.txt".

Note:
- the add-in should work only when the sender is inside the organization (i.e. has domain name specified in the file "accepted_domains.txt").
- the add-in should be able to handle multiple recipients.
- the add-in should be able to handle emails with attachments.
- the add-in should be able to handle emails with HTML content.
- the add-in should be able to handle emails with plain text content.
- the add-in should be able to handle emails with both HTML and plain text content.
- for creating ZIP files should be used the library ProDotNetZip version 1.20.0
