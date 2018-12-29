# VBA RunPE

## Description 
A simple yet effective implementation of the RunPE technique in VBA. This code can be used to run executables from the memory of Word or Excel. It is compatible with both 32 bits and 64 bits versions of Microsoft Office 2010 and above. 

![Win10_x64_Office2016_x64_PowerShell](https://github.com/itm4n/VBA-RunPE/raw/master/screenshots/00_runpe-demo.gif)

## Usage 1 - PE file on disk 
1) In the `Exploit` procedure at the end of the code, set the path of the file you want to execute. 
```
strSrcFile = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
```
__/!\\__ If you're using a 32 bits version of Microsoft Office on a 64 bits OS, you must specify 32 bits binaries. 
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
3) (Optional) Enable __View__ > __Immediate Window__ (`Ctrl+G`) to check execution and error logs.
4) Run the `Exploit` macro!

## Usage 2 - Embedded PE 
1) Use `pe2vba.py` to convert a PE file to VBA. This way, it can be directly embedded into the macro. 
```
user@host:~$ python pe2vba.py meterpreter.exe 
[+] Created file 'meterpreter.exe.vba'.
```
2) Replace the following code in `RunPE.vba` with the the content of the `.vba` file which was generated in the previous step.
```
' ================================================================================
'                                ~~~ EMBEDDED PE ~~~
' ================================================================================

' CODE GENRATED BY PE2VBA
Private Function PE() As String
    Dim strPE As String
    strPE = ""
    PE = strPE
End Function
```
3) (Optional) Enable __View__ > __Immediate Window__ (`Ctrl+G`) to check execution and error logs.
4) Run the `Exploit` macro!

__/!\\__ When using an embedded PE, the macro will automatically switch to this mode because the `PE()` method will return a non-empty string. 

## Credits
This code is mainly a VBA adaptation of the C++ implementation published by @Zer0Mem0ry (32 bits only).
https://github.com/Zer0Mem0ry/RunPE

The PE embedding method was inspired by @DidierStevens' work. https://blog.didierstevens.com/

## Misc

### Tests
This code was tested on the following platforms:
- Windows 7 Pro 32 bits + Office 2010 32 bits
- Windows 7 Pro 64 bits + Office 2016 32 bits
- Windows 2008 R2 64 bits + Office 2010 64 bits
- Windows 10 Pro 64 bits + Office 2016 64 bits

Currently, this doesn't work with all Windows binaries. For example, it can't be used to run _regedit.exe_. I guess I need to do some manual imports of missing DLLs.

### Side notes
Here is a table of correspondence between some Win32 and VBA types:

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

### What about older versions of Microsoft Office (<=2007)?

As mentionned in the description, this code only works with Office 2010 and above. The reason for this is that the `LongPtr` type is extensively used. It was first introduced in Office 2010 to help developers make architecture independant code. Indeed, as described above, its size will be automatically adapted depending on the architecture of the Office process (32-bits / 64-bits).

So, if you try to run this code in Office 2007, you will get a `User-defined type not defined` error message for each variable using the `LongPtr` type. To work around this issue, you can replace all the `LongPtr` occurences with `Long` (32-bits) or `LongLong` (64-bits). Use `Ctrl+H` in your favorite text editor! ;)

Note: the code could be updated to take this compatibility issue into account but it would require too much effort for relatively little gain.
