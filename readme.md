# <img src="./assets/icon.png" width="50" alt=""> Email Extraction Utility 

## Introduction

This utility takes a list of emails, as Outlook typically creates, cleans the email list removing everything but the emails themselves. The utility uses the Windows clipboard to transfer the original email list, processes them, then posts the new email list back to the clipboard for the user to 'paste' them where desired.

Converts this: ``` FirstName1 LastName1 - C00000 <Employee.number1@company.com>;```

Into this: ``` Employee.number1@company.com;```

## Building

1. If the Python module 'pyperclip' is not installed use 'pip' to install it.
   ``` pip install pyperclip ```
2. Execute the command script ```build.cmd```
3. If succesful, the new executable will be located here ``` .\dist\extEmails\extEmails.exe ```

## Installation

1. Copy the ``` .\dist\extEmails\extEmails.exe ``` file to any desired folder.
2. Create a shortcut to the ```extEmails.exe ``` file and place it on your toolbar for easy, global access.
## Execution

1. Highlight and copy the original email list using the Windows 'copy' function. [CTRL+C]
2. Execute the program, "extEmails.exe"
3. Then use the Windows 'paste' function, [CTRL-V], to paste the new clean email list where desired.

### 1. Copy the original Outlook email list [CTRL-C]

```text
FirstName1 LastName1 - C00000 <Employee.number1@company.com>;FirstName2.LastName2 - A000000 <FirstName2.LastName2@company.com>
```

### 2. Run the Executable
![Alt text](assets/exe.png)
### 3. Paste the re-factored email list where desired [CTRL-V]

```text
Employee.number1@company.com;FirstName2.FirstName2.LastName2@company.com
```

