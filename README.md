# VBA RunPE

## Description 
A simple yet effective implementation of the RunPE technique in VBA. This code can be used to run executables from the memory of Word or Excel. It is compatible with both 32 bits and 64 bits versions of Microsoft Office 2010 and above. 

More info here:  
https://itm4n.github.io/vba-runpe-part1/  
https://itm4n.github.io/vba-runpe-part2/  

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

1. Use `pe2vba.py` to convert a PE file to VBA. This way, it can be directly embedded into the macro.

```
user@host:~/Tools/VBA-RunPE$ ./pe2vba.py meterpreter.exe 
[+] Created file 'meterpreter.exe.vba'.
```

2. ~~Replace the following code in `RunPE.vba` with the the content of the `.vba` file which was generated in the previous step.~~ The Python script converts the PE to VBA and applies the RunPE template automatically (no need to copy/paste manually).

```
' ================================================================================
'                                ~~~ EMBEDDED PE ~~~
' ================================================================================

' CODE GENRATED BY PE2VBA
' ===== BEGIN PE2VBA =====
Private Function PE() As String
    Dim strPE As String
    strPE = ""
    PE = strPE
End Function
' ===== END PE2VBA =====
```
3. (Optional) Enable __View__ > __Immediate Window__ (`Ctrl+G`) to check execution and error logs.

4. Run the `Exploit` macro!

__/!\\__ When using an embedded PE, the macro will automatically switch to this mode because the `PE()` method will return a non-empty string.


## Known issues

- __`GetThreadContext()` fails with error code 998.__

You might get this error if you run this macro from a __64-bits version of Office__. ~~__As a workaround__, you can move the code to __a module__ rather than executing it from the Word Object references. Thanks [@joeminicucci](https://github.com/joeminicucci) for the tip.~~

```
================================================================================
[*] Source file: 'C:\Windows\System32\cmd.exe'
[*] Checking source PE...
[*] Creating new process in suspended state...
[*] Retrieving the context of the main thread...
    |__ GetThreadContext() failed (Err: 998)
```

~~I have no idea why this workaround works for the moment. I've investigated this a bit though.~~ This error seems to be caused by the `CONTEXT` structure not being properly aligned in the 64-bits version. I noticed that the size of the structure is also incorrect (`[VBA] LenB(CONTEXT) != [C++] sizeof(CONTEXT)`) whereas it's fine in the 32-bits version. I have a working solution that allows the `GetThreadContext()` to return properly but then it breaks some other stuff further in the execution. 

__Edit 2019-12-15__: the definition of the 64-bits version of the `CONTEXT` structure was indeed incorrect but fixing this didn't fix the bug. So, I implemented a workaround for the 64-bits version. I replaced the `CONTEXT` structure argument of the `GetThreadContext()` and `SetThreadContext()` functions with a `Byte` Array of the same size. 

__Edit 2019-12-17__: I finally found the problem. My first assumption was correct, the `CONTEXT` structure must be 16-Bytes aligned in memory. This is something you can control in C by using `align(16)` in the definition of the structure but you can't control that in VBA. Therefore, `GetThreadContext()` and `SetThreadContext()` may "randomly" fail. `Byte` Arrays on the other hand seem to always be 16-Bytes aligned, that's why this workaround is effective but there is no guarantee, unless I reverse engineer the VBA interpreter/compiler and figure it out?!

- __`LongPtr` - _User Defined Type Not Defined___

If you get this error, it means that you are running the macro from an old version of Office (<=2007). The `LongPtr` type was introduced in VBA7 (Office 2010) along with the support of the 64-bits Windows API. It's very useful for handling pointers without having to worry about the architecture (32-bits / 64-bits).

As a workaround, you can replace all the `LongPtr` occurences with `Long` (32-bits) or `LongLong` (64-bits). Use `Ctrl+H` in your favorite text editor.


## Credits

[@hasherezade](https://twitter.com/hasherezade) - Complete RunPE implementation (https://github.com/hasherezade/)

[@Zer0Mem0ry](https://github.com/Zer0Mem0ry) - 32 bits RunPE written in C++ (https://github.com/Zer0Mem0ry/RunPE)

[@DidierStevens](https://twitter.com/didierstevens) - PE embedding in VBA


## Misc

### Tests

This code was tested on the following platforms:
- Windows 7 Pro 32 bits + Office 2010 32 bits
- Windows 7 Pro 64 bits + Office 2016 32 bits
- Windows 2008 R2 64 bits + Office 2010 64 bits
- Windows 10 Pro 64 bits + Office 2016 64 bits

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
