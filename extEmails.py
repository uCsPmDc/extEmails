# This utility takes a list of emails, as Outlook typically creates, cleans 
# the email list removing everything but the emails themselves. The utility 
# uses the Windows clipboard to transfer the original email list, processes 
# them, then posts the new email list back to the clipboard for the user to 
# 'paste' them where desired.
# 
# Ref: https://pyinstaller.org/en/stable/spec-files.html
#
# Author:   uCsPmDc
# Date:     Jan 5, 2024
# Version:  1.0

try:
    # importing regex module
    import re           # Regex module
    import pyperclip    # For creating an EXE (pip install)

except ModuleNotFoundError:

    print(f'\n\n----------------------------------------------')
    print(f' Error! Python module not found.')
    print(f'   Please run "pip install -r requirements.txt"')
    print(f'   from the command line. Then try again.')
    print(f'----------------------------------------------')
    exit(1)

def substr_swap(str_msg: str, dic: dict) -> str:
    ''' Replaces multiple substrings passed in with a dictionary
          Syntax: str = substr_swap(str, {">": "", "<": ""})
    '''
    for char in dic.keys():
        str_msg = str_msg.replace(char, dic[char])
    return str_msg

def main():
    emailStr: str = ""
    s = pyperclip.paste()                         # Get current clipboard contents
    result = re.findall(r'\S+@\S+', s)            # Regex parses all email address text

    for r in result:                              # Loop and eliminate '<', '>' chars
        emailStr += substr_swap(r, {">": "", "<": "", ";": "; "})

    # print(emailStr, end='')
    pyperclip.copy(emailStr)                      # Put the clean email string back on the clipboard

if __name__ == '__main__':
    main()
