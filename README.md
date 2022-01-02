# form2docx

## What's this?

"form2docx" is a [Google Apps Script](https://workspace.google.co.jp/intl/en/products/apps-script/) (GAS) that converts contents retrieved by Google Forms into Microsoft Word file (.docx) or OpenDocument Text file (.odt) and sends it as email attachment.

## Requires

* Your [Google account](https://www.google.com/account/about/)
* Application for creating templates, such as [Microsoft Word](https://www.microsoft.com/en-us/microsoft-365/word),  [LibreOffice](https://www.libreoffice.org/) Writer and etc.

## Usage

1. Create a template file includes placeholder(s) and save as .docx or .odt format (see template-sample.odt file)
2. Upload the template file to the Google Drive
3. Open the uploaded file with Google Docs application (converts automatically from .docx / .odt to Google Docs format)
4. Get the ID of the converted template file
5. Create a new input form according to the template using Google Forms
6. Open the GAS script editor of the form and put "form2docx" script to the editor
7. Edit user settings of the script (e.g. the ID of the converted template file)
8. Set trigger event and exec permission to the script
9. Open the form to the public

## Notice

* Please use at your own risk
* Generated .docx / .odt files are archived in Google Drive

## Changelog

* 1.0.0 (2022-01-01)
  * Opening to the public

## License

[The MIT License](https://opensource.org/licenses/mit-license.php)