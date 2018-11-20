# NonMalicious-Info-Stealer-Macro
Sample VBA Code, open a new word document, Alt+F11, copy and paste the code save as Word 97 ".doc".
This word VBA macro is able to steal information such as browser saved password/credentials, save to a working directory example: Pen Drive.
The attacker can use the code as if the attacker is using the victim's computer to edit his fake document but in the background,
the code is stealing informations or search for documents(future update). The attacker also can upload the stolen files onto his own FTP server.
All tools used in the code are non-malicious and only uses Microsoft Windows built-in tools and operation such as the cmd.exe, ftp.exe, 
copy and paste operation. Only advanced behavioral detection tools may catch this activity. No third party tools or executable downloaded
from the document.

Tested with
Windows Defender 10 - may catch but changing the content of the doc will do the trick.
Panda Adaptive Defense 360 - will miss it.
SentinelOne - Will catch it due to the behavior of the document.
