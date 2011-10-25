YajHFC Outlook phone book plugin
================================

This plugin adds a phone book to YajHFC that reads its contents from Outlook contacts.
It does this by invoking the Simple MAPI COM interface of Outlook (using the JACOB library).

This plugin is for Windows only.

INSTALLATION
============

If you have installed YajHFC using the Windows installer:
- Use the provided setup exe for the plugin
*or*
- Unpack the archive into the YajHFC installation directory
- Start YajHFC using the start menu or desktop icon, go to Options->Plugins&JDBC and add the yajhfc-outlook-pb-plugin.jar as plugin.

If you manually installed YajHFC:
- Unpack the archive into the YajHFC installation directory
- Start YajHFC using the provided yajhfc-with-outlook.cmd batch file

USAGE
=====

- Go to the phone book window, select Phonebook->Add to list->Microsoft Outlook contacts.
- Select the contacts folder this phone book should display 
- Click OK
You should now have a new phone book containing the Outlook contacts from the selected folder.

Please note that Outlook may display a security warning if you select the "Read email address and comment" or "Read distribution lists" options.
If you wish to disable these warnings, please follow the following steps (however, this will allow other programs to access this information or send mails without warning, too):
1. Open Outlook 2007.
2. Click "Tools", "Trust Center". The "Trust Center" dialog box will open.
3. Click "Programmatic Access"
4. Click on the circle next to the wording "Never warn me about suspicious activity".
5. Click "OK". The security warnings within Outlook 2007 are now disabled.

If you have multiple folders with contacts in Outlook, you can simply add multiple phone books.

SUPPORT
=======

If you have any problems or feedback, please write a mail to support@yajhfc.de
