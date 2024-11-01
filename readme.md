# <img src="./assets/icon.png" width="50" alt=""> Email Extraction Utility 

## Introduction

This utility takes a list of emails, as Outlook typically creates, cleans the email list removing everything but the emails themselves. The utility uses the Windows clipboard to transfer the original email list, processes them, then posts the new email list back to the clipboard for the user to 'paste' them where desired.

Converts this: ``` FirstName1 LastName1 - C00000 <FirstName1.LastName1@company.com>; ```
Into this: ``` FirstName1.LastName1@company.com; ```

## Building

1. If the Python module 'pyperclip' is not installed use 'pip' to install it.
   ``` pip install pyperclip ```
2. Execute the command script ```build.cmd```
3. If successful, the new executable will be located here ``` .\release\extEmails.exe ```

## Installation

1. Copy the ```.\release\extEmails.exe``` file to any desired folder.
2. Create a shortcut to the ```extEmails.exe``` file, and place it on your toolbar for easy, global access.

## Execution

1. Highlight and copy the original email list using the Windows 'copy' function. **[CTRL+C]**
2. Execute the program, "extEmails.exe"
3. Then use the Windows 'paste' function, **[CTRL-V]**, to paste the new clean email list where desired.

### 1. Copy the original Outlook email list [CTRL-C]

```text
FirstName1 LastName1 - C000 <FirstName1.LastName1@company.com>; FirstName2 LastName2 - A000 <FirstName2.LastName2@company.com>
```

### 2. Run the Executable ![Alt text](assets/exe.png)

### 3. Paste the re-factored email list where desired [CTRL-V]

```text
FirstName1.LastName1@company.com; FirstName2.LastName2@company.com
```
## Adding extEmail to Right-Click Menu Windows 11

[Original Web Link](https://answers.microsoft.com/en-us/windows/forum/all/how-to-add-an-option-in-windows-11-main-context/a6c6ba72-e17b-44d5-b957-46af79b47b60)

Here are the steps to add an option to the Windows 11 main context menu:

1. Open the registry editor. You can do this by pressing the Win + R key combination to open the Run dialog, then type "regedit" and press Enter to open the registry editor.
2. Navigate to the following path: ```HKEY\_CLASSES\_ROOT\\Directory\\Background\\shell```
3. Right-click on the "shell" key, select "New" > "Key".
4. Name the new key, which will be the name of the new option you're adding.
5. Right-click on the new key, select "New" > "Key".
6. Name the new subkey "command".
7. Right-click on the "command" key, select "Modify".
8. In the "Value data" field, enter the full path of the command you want to run. For example, if you want to run a program called "myprogram.exe", you can enter `"C:\\Program Files\\MyProgram\\myprogram.exe"`.
9. Click "OK" to save the changes.
10. Close the registry editor.

### You should now see your new option in the Windows 11 main context menu.

To ensure that your new option appears in the main context menu instead of "Show More Options", try following these steps:

1. Follow the above steps to create a new subkey and command value data in the registry.
2. Right-click on the new key you created and select "New" > "Key".
3. Name the new subkey "Position".
4. Right-click on the "Position" key and select "Modify".
5. In the "Value data" field, enter "Top", and then click "OK" to save the changes.
6. Close the registry editor.
7. Log out or restart your computer to apply the changes.

At this point, your new option should appear in the main context menu. If it still appears under "Show More Options", make sure the "Position" subkey is set to "Top".

#### Regedit Export ("extEmail.reg") 

```registry
Windows Registry Editor Version 5.00

[HKEY_CLASSES_ROOT\Directory\background\shell\extEmail]
"Icon"=hex(2):22,00,44,00,3a,00,5c,00,53,00,74,00,6f,00,72,00,61,00,67,00,65,\
  00,5c,00,5f,00,41,00,70,00,70,00,73,00,5c,00,65,00,78,00,74,00,45,00,6d,00,\
  61,00,69,00,6c,00,73,00,5c,00,65,00,78,00,74,00,45,00,6d,00,61,00,69,00,6c,\
  00,73,00,2e,00,65,00,78,00,65,00,22,00,00,00

[HKEY_CLASSES_ROOT\Directory\background\shell\extEmail\command]
@="\"D:\\Storage\\_Apps\\extEmails\\extEmails.exe\""
```
