# VBA RunPE

## Description 
A simple yet effective implementation of the RunPE technique in VBA. This code can be used to run executables from the memory of Word or Excel. It is compatible with both 32 bits and 64 bits versions of Microsoft Office 2010 and above.   


![Win10_x64_Office2016_x64_PowerShell](https://github.com/itm4n/VBA-RunPE/raw/master/screenshots/01_Win10_x64_Office2016_x64_PowerShell.png)

## Usage
1) In the ___Exploit___ procedure at the end of the code, set the path of the file you want to execute. 
```
strSrcFile = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
```
__/!\\__ If you're using a 32 bits version of Microsoft Office on a 64 bits OS, you must use 32 bits binaries. 
```
strSrcFile = "C:\Windows\SysWOW64\cmd.exe"
strSrcFile = "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
```
2) Specify the command line arguments (optional).
```
strArguments = "-exec Bypass"
```
This will be used to form a command line equivalent to:
```
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -exec Bypass
```
3) Enable __View__ > __Immediate Window__ (Ctrl + G) to check execution and error logs.
4) Run the macro!

## Credits
This code is mainly a VBA adaptation of the C++ implementation published by @Zer0Mem0ry (32 bits only).
https://github.com/Zer0Mem0ry/RunPE

## Misc

### Tests
This code was tested on the following platforms:
- Windows 7 Pro 32 bits + Office 2010 32 bits
- Windows 7 Pro 64 bits + Office 2016 32 bits
- Windows 2008 R2 64 bits + Office 2010 64 bits
- Windows 10 Pro 64 bits + Office 2016 64 bits

Currently, this doesn't work with all Windows binaries. For example, it can't be used to run _regedit.exe_. I guess I need to do some manual imports of missing DLLs.

### Side notes
Here is a table of correspondence between some C++ and VBA types:

| C++ | VBA | Arch |
| --- | --- | --- |
| BYTE | Byte | 32 & 64 |
| WORD | Integer | 32 & 64 |
| DWORD, ULONG, LONG | Long | 32 & 64 |
| DWORD64 | LongLong | 64 |
| HANDLE | LongPtr(\*) | 32 & 64
| LPSTR | String | 32 & 64 |
| LPBYTE | LongPtr(\*) | 32 & 64 |

(\*) LongPtr is a "dynamic" type, it is 4 Bytes long in Office 32 bits and 8 Bytes long in Office 64 bits. 
https://msdn.microsoft.com/fr-fr/library/office/ee691831(v=office.14).aspx 
