# AutoHotkeyNotify

Client.ahk goes out on a 5 minute interval and checks a website for messages and displays them.

The message should appear as follows

SUBJECT|MESSAGE|GC=FFFFFF GR=3 TC=Black MC=Black BC=FF0000 BW=4 BF=600 SC=500 SI=600 Image=245 C:\Windows\system32\shell32.dll|TARGET

Target can be ALL or a specific computer. If multiple computers are targeted, they are comma separated.

If you wish to have multiple alerts separate them with a ¿ (ALT+168). 
The database is in MySQL and a php page pulls from it. The database table structure is:

```
Column  Type    Comment
id  int(11) Auto Increment   
Title   varchar(64)  
Message varchar(256)     
Options varchar(256)     
Target  varchar(256)     
Enabled bit(1)
```
  
And it ends up looking like ![this](http://i.imgur.com/0Rd88Ct.png "Table Structure")
