import csv
import docx
import time
import smtplib
import os
from time import sleep
from tkinter import *
from tkinter import filedialog
import tkinter as tk
import tkinter
# New User Automation Script #
import docx
from docx import Document
root = tk.Tk()
canvas = tk.Canvas(root, width=300, height=100)
canvas.grid(columnspan=3,rowspan=3)

# Returns a string which is the filepath where the file is located
def openFile():
    filepath = filedialog.askopenfilename(parent=root,title="Open the Word Document",)
    return filepath
# Splits the Word  document into a string
def word_document_input(file):
    doc = docx.Document(file)

    completedText = []

    for paragraph in doc.paragraphs:
        completedText.append(paragraph.text)
    new_string =''.join(map(str, completedText))
    automatic_input = new_string.split(':')
    print(automatic_input)
    return automatic_input
def browse_window():
    #Instructions
    instruction = tk.Label(root, text ="Select the New Starter Word Document: ",)
    instruction.grid(columnspan=3, column=0, row=1)

    #browse button
    browse_text = tk.StringVar()
    browse_btn = tk.Button(root, textvariable=browse_text, command=lambda:word_document_input(openFile()), font="Raleway",
                           bg="#20bebe", fg="white", height=2, width=15)
    browse_text.set("Browse")
    browse_btn.grid(column=1,row=2)

    canvas = tk.Canvas(root, width=600, height=250)
    canvas.grid(columnspan=3)
    root.mainloop()
    return None
#Checks if the user already exists in the Active Directory Database

def active_directory_account_check(sam_account):
    with open('activeDirectory.txt', 'r') as file:
        data = file.readlines()
        for i in range(len(data)):
            if sam_account in data[i]:
                print(f'User already exists: {data[i]}', end='')
                while True:
                    sam_account = str(input("Enter a new username: "))
                    if (sam_account in data[i]) == False:
                        break
    return sam_account
#Flow control username  - TO DO FOR CLONE ( IF CLONE in database == False, Provide another clone )
#Query through AD database (imported in SQL/TXT/CSV file ) -> Currently as .txt
def active_directory_account_check(sam_account):
    with open('activeDirectory.txt', 'r') as file:
        data = file.readlines()
        for i in range(len(data)):
            if sam_account in data[i]:
                print(f'User already exists: {data[i]}', end='')
                while True:
                    sam_account = str(input("Enter a new username: "))
                    if (sam_account in data[i]) == False:
                        break
    return sam_account

# Query through the SQL /CSV  or text file (Currently only text file) database for account checking
def _line_manager_check(manager_check):
    with open('activeDirectory.txt', 'r') as file:
        data = file.readlines()
        for i in range(len(data)):
            if manager_check in data[i]:
                print(f'The manager account is: {data[i]}')
    return manager_check

#Totally useless function
def __useless_function():
    print("=== INVALID ORGANIZATIONAL UNIT ===")

# Output as a string the time when the new user has been created in Active Directory
def _time_():
    now = time.gmtime()
    time_map = f'{now[2]}-{now[1]}-{now[0]}'
    return time_map

# First email notification template
def outlook_propco_notification():
    subject = f'REQ{request_number} - Test_Application Account Creation Notification - {user_first_name} {user_last_name}\n'
    body = f' \n Hello,\n\n Please create a Test_Application Account for the following user:\n' \
           f'\n Email address:{user_SAM_account}{main_SMTP_extension}\n'

    with open('application.txt', 'w') as auto_email:
        auto_email.write(subject)
        auto_email.write(body)
    os.system('application.txt')

# New user notification email
def outlook_email():
    subject = f'REQ{request_number} - New Starter Notification - {user_first_name} {user_last_name}\n'
    body = f' \n Hello,\n\n We have created as a new starter with the following details:\n' \
           f'\n Job Title: {user_job_title}\n' \
           f'\n Department: {user_department}\n' \
           f'\n Reporting to: {user_line_manager}\n'

    with open('email_notification.txt', 'w') as auto_email:
        auto_email.write(subject)
        auto_email.write(body)

# Powershell Script Output -> The powershell script based on which the user is created.
def active_directory():

    remote_mailbox_address = '@onecwmail.onmicrosoft.com'
    # Not
    new_remote_mailbox = f'New-RemoteMailbox {user_SAM_account} -RemoteRoutingAddress {user_SAM_account}{remote_mailbox_address} ' \
     f'-UserPrincipalName {user_SAM_account} -PrimarySmtpAddress {main_smtp}\n'
    import_AD_module = 'Import-Module ActiveDirectory\n'
    # CHANGED
    user_path_defined = path = "$user =" +"'"+clone_user+"'"+'\n'\
                        "Get-ADUser -Filter 'samAccountName -like $user' | ForEach-Object{ $DN=$_.distinguishedname -split',' \n" \
                        "$clone_location =$DN[1..($DN.count -1)] -join ','} \n"
    # Distinguished Name Path
    ou_path = f'$ou_path = $clone_location '
    create_new_ad_user = f'$New_Starter = New-ADUser -Name "{user_first_name} {user_last_name}" ' \
                         f' -ChangePasswordAtLogon $true ' \
                         f' -GivenName {user_first_name} ' \
                         f' -Surname {user_last_name} ' \
                         f' -SamAccountName {user_SAM_account} ' \
                         f' -UserPrincipalName {user_SAM_account}{main_SMTP_extension} ' \
                         f' -Path $ou_path ' \
                         f' -AccountPassword(ConvertTo-SecureString -AsPlainText "ValidPassword1234CZ!" -Force) ' \
                         f' -PassThru | Enable-ADAccount \n'
    new_starter_sam_account = f'$new_starter_sam_account = "{user_SAM_account}"\n'
    new_starter_name = f'$new_starter_name = "{user_first_name} {user_last_name}"\n'

    source_user_groups = f'$SourceUsersGroup = "{clone_user}" \n'
    destination_user = f'$DestinationUser = $new_starter_sam_account \n'
    get_ad_source_user_member_of = f'Get-ADUser $SourceUsersGroup -Properties MemberOf | Select-Object -ExpandProperty MemberOf \n'
    copy_memberof_from_user = f'$sourceUserMemberOf ={get_ad_source_user_member_of}\n'

    # Loop through the member groups in ad (  activedir + loop member of = same command splitted in two chunks)#
    active_diectory_group = "{Get-ADGroup -Identity $group | Add-ADGroupMember -Members $DestinationUser}"
    loop_memberof = f'foreach($group in $SourceUserMemberOf){active_diectory_group}\n'

    set_AD_employee_number = f'Set-ADUser {user_SAM_account} -EmployeeNumber Unknown \n'
    ad_user_description = f'Set-ADUser {user_SAM_account} -description "{set_ad_user_description}" \n'
    ad_user_street = f'Set-ADUser {user_SAM_account} -StreetAddress "{user_location}" \n'
    ad_user_manager = f'Set-ADUser {user_SAM_account} -Manager {user_line_manager}\n'
    ad_user_job_title = f'Set-ADUSer {user_SAM_account} -Title "{user_job_title}"\n'
    ad_user_office = f'Set-AdUser {user_SAM_account} -Office "{user_office}"\n'
    ad_user_department = f'Set-ADUser {user_SAM_account} -Department {user_department}\n'
    ad_user_display_name = f'Set-ADUser {user_SAM_account} -Displayname "{user_first_name} {user_last_name}"\n'
    ad_user_email_address = f'Set-ADUser {user_SAM_account} -EmailAddress {user_SAM_account}{main_SMTP_extension}\n'
    ad_source_user_member_of = f'$SourceUserMemberof = $Get-AdUser $sourceUserGroup -Properties MemberOf ' \
                               f'| Select-Object -ExpandProperty MemberOf \n'
    ad_groups_loop = 'foreach($group in $sourceUserMemberof)' \
                     '{Get-AdGroup -Identity $Group | Add-ADGroupMember -Members' \
                     '$DestinationUser}\n'
    destination_user_member_of = f'$SourceUsersMemberOf = Get-ADUser $DestinationUser -Properties MemberOf ' \
                                 f'| Select-Object -ExpandProperty memberof \n'
    ad_user_groups_copied = "foreach($group in $SourceUsersMemberOf){Get-ADGroup -Identity $group | Select-Object -ExpandProperty samAccountName}\n"
    space_between_rows = "\n"
        # Final script output file #
    with open('reap.ps1', 'w') as module:
        # module.write(powershell_exectuable)

        # module.write(new_remote_mailbox)
        module.write(import_AD_module)
        module.write(user_path_defined)
        module.write(ou_path)
        module.write(space_between_rows)

        module.write(create_new_ad_user)
        module.write(new_starter_sam_account)
        module.write(new_starter_name)
        module.write(space_between_rows)

        module.write(source_user_groups)
        module.write(destination_user)
        module.write(copy_memberof_from_user)
        module.write(loop_memberof)
        # module.write(destination_user)#
        module.write(destination_user_member_of)
        module.write(ad_user_groups_copied)
        module.write(space_between_rows)

        module.write(ad_user_description)
        module.write(set_AD_employee_number)
        module.write(ad_user_job_title)
        module.write(ad_user_manager)
        module.write(ad_user_street)
        module.write(ad_user_office)
        module.write(ad_user_display_name)
        module.write(ad_user_department)
        module.write(ad_user_email_address)


def file_script():
    script_output_file = ""
    with open('reap.ps1', 'a') as file:
        file.write(script_output_file.join(all_extra_alias))
        file.write(script_output_file.join(final_output))
        file.write(script_output_file.join(main_smtp))

# Not implemented yet - Boundry check for white spaces - If needed
def line_remove():
    with open('reap.txt', 'r') as q:
        lines = q.readlines()
        lines = [line.replace(' ', '') for line in lines]

    with open('reap.txt', 'w') as q:
        q.writelines(lines)


def email_extension():
    cw = automatic_input[15]
    return cw


def new_starter():
    name = input("User name: ")
    database = input("RPS Database: ")
    extension = input("Email Extension: ")
    rps_exe = '@rpscountrywide.co.uk'
    print(name + '-' + database + rps_exe)
    print(name + extension)



ad_OU = ['IT', 'Finance', 'Sales', 'Human Resources', 'Legal']
yes = ['yes', 'Yes', 'y', 'Y', ]
no = ['no', 'N', 'No', 'n', 'nope']
approvall = (yes, no)
reapit_databases = (['ABB', 'BDS', 'BRI', 'CSX', 'CSW', 'CTW', 'NTH', 'SHH', 'CWC', 'CWX', 'NWE', 'CWE'])
ad_path_ou = []
# F
print('>>>-- $-> New User Automation Script <-$ --<<<')
###########---> Word File Input <---##############
automatic_input = word_document_input(openFile())
##################################################
request_number = str(input("Request Number: "))
agent_name = str(input("Enter your initials:"))
####  TO DO -> Flow control - Automate relevant data extraction
user_first_name = automatic_input[0+1]
user_last_name = automatic_input[2+1]
##########################################################
user_SAM_account = f'{user_first_name}.{user_last_name}'
active_directory_account_check(user_SAM_account)
user_SAM_account = active_directory_account_check(user_SAM_account)
###-> Clone user to copy the Ative Directory groups from
clone_name = automatic_input[5]
clone_user = clone_name.replace(" ",".")
print(clone_user)
###########################################################
user_job_title = automatic_input[13]
manager = automatic_input[11]
user_line_manager = _line_manager_check(manager).replace(" ",".")
print(user_line_manager)
###########################################################
user_office = automatic_input[6]
user_department = automatic_input[17]
user_location = automatic_input[9]
set_ad_user_description = f'Manual Setup {_time_()} {agent_name}{request_number}'
main_SMTP_extension = email_extension()
user_OU = user_department
#if user_OU in ad_OU:
#    ad_path_ou = user_OU
#else:
#    while True:
#        user_OU = str(input("User Organizational Unit: "))
#        if user_OU in ad_OU:
#                ad_path_ou = user_OU
#                print(f'Valid OU has been found: {user_OU}')
#                break
#
set_the_main_smtp = 'Set-ADUser ' + user_SAM_account + ' -add @{ProxyAddresses=''"SMTP:' \
                    + user_SAM_account + main_SMTP_extension + '"' + '}'
main_smtp = []
main_smtp.append(set_the_main_smtp)

rps12_check_question = str(input("RPS 12 Y/N? "))

final_output = []
alias_extension = []
user_sam_account_alias = (user_SAM_account + '-')
extra_alias = ('@ExtraAlias.com')
all_extra_alias = []
if rps12_check_question in yes:
    for a in range(len(reapit_databases)):
        print(f'{user_sam_account_alias}{reapit_databases[a]}{extra_alias}')
        all_allias = 'Set-ADUser ' + user_SAM_account + ' -add @{ProxyAddresses="smtp:' + user_sam_account_alias + \
                     reapit_databases[a] + extra_alias + '"}\n'
        all_extra_alias.append(all_allias)

    multiple_rps_database = str(input("Additional email aliases required(Y/N): "))
    rps_database_string = []

    if multiple_rps_database in yes:
        number_of_rps_databases = int(input("How many additional email aliases?(Enter a number): "))
        for i in range(number_of_rps_databases):
            rps_database_input = str(input("Select Extra Aliases: "))
            rps_database_string.append(rps_database_input)
            alias_extension = '@testExtraAlias.co.uk'
        for j in range(number_of_rps_databases):
            print(user_SAM_account + '-' + rps_database_string[j] + alias_extension)
            setup_rps12_smtp = 'Set-ADUser ' + user_SAM_account + ' -add @{ProxyAddresses=''"smtp:' \
                               + user_SAM_account + '-' + rps_database_string[j] + alias_extension + '"' + '}' + '\n'
            final_output.append(setup_rps12_smtp)
        print(f'{final_output}\n')



prop_check = str(input("App1 account needed: "))

if prop_check in yes:
    outlook_propco_notification()
    print("Email request has been sent")

vex_account = str(input("App2 account needed: "))

if vex_account in yes:
    print("Email request sent")

active_directory()
file_script()
outlook_email()

os.system('application.txt')
os.system('reap.ps1')
os.system('email_notification.txt')
os.system('scripthtat.bat')
