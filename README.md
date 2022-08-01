# Office - Auto Email Send

The goal of this project is to send customized emails to a given mail list. The solution was created to help Mines Nancy school's administration send transcripts to the students faster. It was also used to get customized informations from the students as regards Mines Nancy N18's graduation ceremony, where a Powerpoint presentation was created. The strategy of the solution is the following : 
- The solution uses Microsoft Outlook, that Excel can connect to
- Each mail sent by Outlook is a form that Excel can fill with VBA : email adress, message, embedded files etc...
- Each row of the mail list corresponds to a mail that will be sent, the columnn elements will be used to send the mail
- For each mail, a template message is filled with the informations contained in a given row of the mail list
- Iteratively, the Outlook form is filled thanks to the informations contained in each row of the mail list, and the mail is sent

This process is illustrated below. The scenario of a school administration, sending GPAs and transcripts to the students, is the one shared in this repository :

![AutoSendMail](https://user-images.githubusercontent.com/106969232/182213171-d7812203-b056-4396-afc4-647929481204.JPG)

HOW TO USE :
- Set-up Microsoft Outlook : a Microsoft compatible address will be required /!\.
- Open "SendReport.xlsm" and enable content and macros.
- Replace the email addresses in the mail list and check to "Mail Body" Tab.
- To check the VBA code, go to "View" on the upper ribbon and select "Macros".
- Select and edit "Rectangle1_Click" (VBA code also saved as "macro.bas"). This macro will be launched when clicking on the "Mail list" tab blue button.
- Outlook may encounter issues if a metered connection is used /!\ (it will wait for a non-metered connection to send the emails). It can be disabled in your current connection's Windows properties.
- Click on the "START" blue button. The "Message sent" row should start filling with OKs, for each mail sent.
- If there is a crash during the process, changing the "Start:" number can enable to restart from where the crash occured.

The solution was tested on : Microsoft Windows 10 Home, Version 10.0.19043 Build 19043 (x64)
