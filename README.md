# Bulk Email Automation
This script sends bulk emails in Outlook to a list of recipients. It connects to Outlook's SMTP server and uses standard TLS encryption to send messages.

## Dependencies
In order for the script to run properly, there are a few prerequisites that must be upheld. The simplest solution is to configure a _virtual environment_ with the following dependencies:

1. Python 3 or above

2. 'xlrd' package

3. 'cx_freeze' package

## Instructions for Usage
In order to properly use the script, a few formatting requirements need to be followed.

1. A spreadsheet file named '[data.xlsx](https://github.com/ryanrashid/automail/blob/master/data.xlsx)' should store recipient information.

    - The first column should have the **names** of recipients
    
    - The second column should have the **email addresses** of recipients
    
    - The script does not read the first row since it is typically column labels.
    
<p align="center"><img width="346" alt="data" src="https://user-images.githubusercontent.com/64723727/89721411-d1711580-d9a2-11ea-8912-68d6cc16a2da.png"></p>

2. An HTML file named '[message.html](https://github.com/ryanrashid/automail/blob/master/message.html)' should store the content of the email.

    - The first line should have the **subject** of the email
    
    - The HTML section should have the **body** of the email
    
    - There is a placeholder for the recipients name. Any additional placeholders in the message should be amended in the Python script as well.

<p align="center"><img width="421" alt="Screen Shot 2020-08-08 at 3 52 40 PM" src="https://user-images.githubusercontent.com/64723727/89719559-328eee00-d98f-11ea-84c9-02d1b4284a87.png"></p>
    
>Note: If you change any of the naming conventions or formatting for these files, you should update the script to reflect them.

## Creating a Standalone Executable

Before running the setup script, be sure that any file naming changes are reflected in the '[setup.py](https://github.com/ryanrashid/automail/blob/master/setup.py)' file.

1. Change directory into the 'automail' folder.

2. In the terminal, run the following command:
```
  $ python3 setup.py build
```
> Note: You may need to replace 'python3' with 'python' depending on how Python was configured.

You should now see a folder called 'build' with another folder that stores the .exe file. You can move the folder anywhere, but **do not** change any of the folder's internal structure. 
However, you can still edit the spreadsheet and message files with updated content as long as their locations remain the same.
