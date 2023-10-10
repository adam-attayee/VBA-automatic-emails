# VBA Save and Attach to Email

This VBA script is designed to automate the process of saving the current workbook and attaching it to a new email in Microsoft Outlook. The script simplifies the task of sending files as email attachments directly from Excel.

## Overview

### Project Objective

Have you ever needed to quickly send the current Excel workbook as an email attachment without the hassle of manually saving it and then composing an email? This VBA script aims to streamline that process. It saves the current workbook, attaches it to a new Outlook email, and opens the email for further review or sending.

### Key Features

- **Automated File Saving**: The script automatically saves the current workbook, eliminating the need for manual saving before sending an email.

- **Outlook Integration**: It leverages Microsoft Outlook to create a new email, attach the saved file, and populate the email with a predefined message.

- **Customization**: Users can customize email recipients, subject, message body, and more, making it adaptable to various use cases.

- **Ease of Use**: Running the script is as simple as executing a macro within Excel, making it accessible to users with basic VBA knowledge.

## Getting Started

To use this script, follow these steps:

1. **Open Excel**: Ensure you have Microsoft Excel installed on your computer.

2. **Enable Macros**: If macros are not already enabled, go to Excel Options > Trust Center > Trust Center Settings > Macro Settings, and select "Enable all macros" or "Enable macros with notification," depending on your security preferences.

3. **Access the VBA Editor**: Press `Alt` + `F11` to open the Visual Basic for Applications (VBA) editor in Excel.

4. **Copy and Paste Code**: Copy the provided VBA code and paste it into a new or existing module within the VBA editor.

5. **Configure Email Settings**: Customize the script by adjusting the recipient's email address, CC recipient(s), subject, and email body as needed within the code.

6. **Run the Script**: Close the VBA editor and return to your Excel workbook. You can then run the script by going to the Developer tab (if not visible, enable it in Excel options) and click on "Run."

7. **Review and Send**: The script will save the current workbook, open a new Outlook email with the saved file attached, and display the email for further review. You can send it manually if everything looks correct.

## Code Structure

The script consists of the following components:

- Variable declarations for Outlook and file-related objects.
- File-saving functionality.
- Creation of a new Outlook email.
- Customizable email settings.
- Cleanup code to release resources.

