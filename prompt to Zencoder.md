This is the prompt to Zencoder which was used to generate he application.

Create an Outlook VSTO add-in (using VB.NET) which performs the following functions before sending an email:
1. checks, if any of email's recipiens is an external recipient (the definition of an external recipient is given below)
2. if no - send an email
3. if yes - display warning "The email will be sent to following external recipients: <list of external recipients>. Are you sure that your email doesn't contain sensitive information? If so, please click on "Cancel" button and then remove confidential information or encrypt the message by clicking on "Encrypt" button " and ask user for choosing between three options: "Cancel", "Encrypt" and "Send".
4. if user chooses "Encrypt" option - ask user for a password and a password description and then create an encrypted ZIP archive containing the message body and all attachments. The passwort for decryption is: " followed by the password description. The password description cannot be empty.Then send an email containing this ZIP archive as an attachment. The subject line of the encrypted email should start with "[Encrypted]" prefix. The message body of the email should contain the text "This email contains confidential information. It was encrypted to ensure, that only authorized persons can access the information contained within this email. 
5. if user chooses "Send" option - send an email without encryption.
6. if user chooses "Cancel" option - do not send an email and return to mail compose window.

Note:
- the add-in should work only when the sender is inside the organization (i.e. has domain name specified in the organization profile settings).
- the add-in should be able to handle multiple recipients.
- the add-in should be able to handle emails with attachments.
- the add-in should be able to handle emails with HTML content.
- the add-in should be able to handle emails with plain text content.
- the add-in should be able to handle emails with both HTML and plain text content.

Definition of an external recipient:
- a recipient with domain name not found in a list of accepted domains stored in a configuration file called "accepted_domains.txt". 