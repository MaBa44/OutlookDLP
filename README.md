# Outlook DLP Add-in

This VSTO add-in for Microsoft Outlook provides Data Loss Prevention (DLP) functionality by detecting and warning users when they are about to send emails to external recipients.

## Features

- Automatically detects external recipients based on domain names
- Warns users when sending emails to external recipients
- Provides options to cancel, encrypt, or send the email as-is
- Encrypts sensitive emails with password protection
- Customizable list of internal domains
- Multilingual support (English, German, Polish)

## Requirements

- Microsoft Outlook 2013 or later
- .NET Framework 4.7.2 or later
- Visual Studio 2017 or later (for development)

## Installation

1. Build the solution in Visual Studio
2. Install the add-in using the VSTO installer
3. Configure the accepted domains in the `accepted_domains.txt` file

## Configuration

The add-in uses a configuration file called `accepted_domains.txt` to determine which email domains are considered internal. This file should contain one domain name per line.

Example:
```
yourdomain.com
company.com
internal.org
```

## Internationalization (i18n)

The add-in supports multiple languages:
- English (default)
- German (de-DE)
- Polish (pl-PL)

The language is automatically selected based on the Outlook UI language settings. If the Outlook language is not supported, the add-in will fall back to English.

## How It Works

1. When a user attempts to send an email, the add-in checks if any recipients have domain names not listed in the `accepted_domains.txt` file.
2. If external recipients are detected, a warning dialog is displayed with three options:
   - **Cancel**: Return to the email composition window
   - **Encrypt**: Encrypt the email with a password
   - **Send**: Send the email without encryption
3. If the user chooses to encrypt the email, they are prompted to enter a password and password description.
4. The encrypted email is sent with the subject prefixed with "[Encrypted]" and the message body and attachments are packaged in a password-protected ZIP file.

## Development

The add-in is built using VB.NET and the Microsoft Office VSTO framework. It uses the DotNetZip library for creating encrypted ZIP archives.

### Project Structure

- `ThisAddIn.vb`: Main add-in class that handles Outlook events
- `Forms/ExternalRecipientsWarningForm.vb`: Form for warning about external recipients
- `Forms/EncryptionForm.vb`: Form for entering encryption password
- `ZipHelper.vb`: Helper class for creating encrypted ZIP archives
- `LocalizationManager.vb`: Class for handling multilingual support
- `Resources/Strings.resx`: Default (English) string resources
- `Resources/Strings.de-DE.resx`: German string resources
- `Resources/Strings.pl-PL.resx`: Polish string resources
- `accepted_domains.txt`: Configuration file for internal domains

## Adding a New Language

To add support for a new language:

1. Create a new resource file named `Strings.[culture-code].resx` in the Resources folder
2. Copy all string entries from the default `Strings.resx` file
3. Translate the values to the target language
4. Build the project

The add-in will automatically detect the new language if it matches the Outlook UI language.

## License

This project is licensed under the MIT License.