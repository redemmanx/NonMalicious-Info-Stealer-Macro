# NonMalicious-Info-Stealer-Macro
A VBA scriptings for use in POC only to test whether your Antivirus able to catch the non-malicious scripting operation using Windows built-in tools to steal information and saved it in a local working folder or upload to an FTP server. There are no third party tools or executables downloaded from this code. Some might call it fileless malware, some call it in-memory attack. For me, it is just scripting attack by utilizing windows built-in command and tools for a bad intentions.

To start, open a new word document, Alt+F11, copy and paste the code save as Word 97 ".doc".
This word VBA macro is able to steal information such as browser saved password/credentials, save to a working directory example, a USB Pen Drive.external hard disk.
The attacker can use the code, pretending to borrow the victim's computer to edit his fake document, but in the background,
the code is stealing informations or search for documents(future update). The attacker also can upload the stolen files onto his own FTP server.
All tools used in the code are non-malicious and only uses Microsoft Windows built-in tools and operation such as the cmd.exe, ftp.exe, 
copy and paste operation. Only advanced behavioral detection tools may catch this activity. No third party tools or executable downloaded from the document.

An attacker may also send the info-stealing document through email as an attachment, victim opens it and the code executes in the background and upload all stolen information and data to a remote FTP file server. 

Tested with
Windows Defender 10 - may catch but changing the content of the doc will do the trick.
Panda Adaptive Defense 360 - will miss it.
SentinelOne - Will catch it due to the behavior of the document.

sample document: harmless-document.doc (macro password: infostealer)

3 VBA code samples:
local-stealing-mode-poc-simple.txt
-very simple code which will even bypass the most Advanced behavioral AV
-For USB drive local mode use

local-stealing-mode-poc.txt
-bypassed most AVs
-for USB drive local mode use

credentials-info-stealer-ftp and local-poc.txt
-bypassed most AVs
-can be used for USB drive local mode and connect to FTP server to upload



