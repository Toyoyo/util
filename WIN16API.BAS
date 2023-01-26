Attribute VB_Name = "Module1"
' -----------------------------------------------------------------------
'
'    WIN31API.TXT -- Windows 3.1 API Declarations for Visual Basic
'
'              Copyright (C) 1994 Microsoft Corporation
'
'
'  This file contains only the Const, Type, and Declare statements for
'  Windows 3.1 APIs.
'
'  The Global Constants ABORTDOC, STARTDOC, and ENDDOCAPI are
'  Windows 3.0 specific and have been superseded by the ABORTDOC,
'  STARTDOC, and ENDDOCAPI functions in Windows 3.1.  These constants
'  have been made into comments in this file.  In order to use them you will
'  need to remove the comment symbol preceding each constant.
'
'  You have a royalty-free right to use, modify, reproduce and distribute
'  this file (and/or any modified version) in any way you find useful,
'  provided that you agree that Microsoft has no warranty, obligation or
'  liability for its contents.  Refer to the Microsoft Windows Programmer's
'  Reference for further information.
'
' ------------------------------------------------------------------------


' General Purpose Types

Type RECT
        left As Integer
        top As Integer
        right As Integer
        bottom As Integer
End Type

Type POINTAPI
        x As Integer
        y As Integer
End Type


' ----------------
'  Kernel Section
' ----------------

' ParameterBlock description structure for use with LoadModule
Type PARAMETERBLOCK
        wEnvSeg As Integer
        lpCmdLine As Long
        lpCmdShow As Long
        dwReserved As Long
End Type

' Loader Routines
Declare Function GetNumTasks Lib "Kernel" () As Integer
Declare Function GetModuleHandle Lib "Kernel" (ByVal lpModuleName As String) As Integer
Declare Function GetModuleUsage Lib "Kernel" (ByVal hModule As Integer) As Integer
Declare Function GetModuleFileName Lib "Kernel" (ByVal hModule As Integer, ByVal lpFilename As String, ByVal nSize As Integer) As Integer
Declare Function GetInstanceData Lib "Kernel" (ByVal hInstance As Integer, ByVal pData As Integer, ByVal nCount As Integer) As Integer
Declare Function LoadLibrary Lib "Kernel" (ByVal lpLibFileName As String) As Integer
Declare Function LoadModule Lib "Kernel" (ByVal lpModuleName As String, lpParameterBlock As PARAMETERBLOCK) As Integer
Declare Sub FreeModule Lib "Kernel" (ByVal hModule As Integer)
Declare Sub FreeLibrary Lib "Kernel" (ByVal hLibModule As Integer)
Declare Function GetFreeSpace Lib "Kernel" (ByVal wFlags As Integer) As Long
Declare Function WinExec Lib "Kernel" (ByVal lpCmdLine As String, ByVal nCmdShow As Integer) As Integer
Declare Sub DebugBreak Lib "Kernel" ()
Declare Sub OutputDebugString Lib "Kernel" (ByVal lpOutputString As String)
Declare Sub SwitchStackBack Lib "Kernel" ()
Declare Sub SwitchStackTo Lib "Kernel" (ByVal wStackSegment As Integer, ByVal wStackPointer As Integer, ByVal wStackTop As Integer)
Declare Function GetCurrentPDB Lib "Kernel" () As Integer
Declare Function GetVersion Lib "Kernel" () As Long

' OpenFile() Structure
Type OFSTRUCT
        cBytes As String * 1
        fFixedDisk As String * 1
        nErrCode As Integer
        reserved As String * 4
        szPathName As String * 128
End Type

' OpenFile() Flags
Public Const OF_READ = &H0
Public Const OF_WRITE = &H1
Public Const OF_READWRITE = &H2
Public Const OF_SHARE_COMPAT = &H0
Public Const OF_SHARE_EXCLUSIVE = &H10
Public Const OF_SHARE_DENY_WRITE = &H20
Public Const OF_SHARE_DENY_READ = &H30
Public Const OF_SHARE_DENY_NONE = &H40
Public Const OF_PARSE = &H100
Public Const OF_DELETE = &H200
Public Const OF_VERIFY = &H400
Public Const OF_CANCEL = &H800
Public Const OF_CREATE = &H1000
Public Const OF_PROMPT = &H2000
Public Const OF_EXIST = &H4000
Public Const OF_REOPEN = &H8000

Declare Function OpenFile Lib "Kernel" (ByVal lpFilename As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Integer) As Integer

' GetTempFileName() Flags
'
Public Const TF_FORCEDRIVE = &H80

Declare Function GetTempDrive Lib "Kernel" (ByVal cDriveLetter As Integer) As Integer
Declare Function GetTempFileName Lib "Kernel" (ByVal cDriveLetter As Integer, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFileName As String) As Integer
Declare Function SetHandleCount Lib "Kernel" (ByVal wNumber As Integer) As Integer

Declare Function GetDriveType Lib "Kernel" (ByVal nDrive As Integer) As Integer

' GetDriveType return values
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4

' Global Memory Flags
Public Const GMEM_FIXED = &H0
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_NOCOMPACT = &H10
Public Const GMEM_NODISCARD = &H20
Public Const GMEM_ZEROINIT = &H40
Public Const GMEM_MODIFY = &H80
Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_NOT_BANKED = &H1000
Public Const GMEM_SHARE = &H2000
Public Const GMEM_DDESHARE = &H2000
Public Const GMEM_NOTIFY = &H4000
Public Const GMEM_LOWER = GMEM_NOT_BANKED

Public Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Declare Function GlobalAlloc Lib "Kernel" (ByVal wFlags As Integer, ByVal dwBytes As Long) As Integer
Declare Function GlobalCompact Lib "Kernel" (ByVal dwMinFree As Long) As Long
Declare Function GlobalFree Lib "Kernel" (ByVal hMem As Integer) As Integer
Declare Function GlobalHandle Lib "Kernel" (ByVal wMem As Integer) As Long
Declare Function GlobalLock Lib "Kernel" (ByVal hMem As Integer) As Long
Declare Function GlobalReAlloc Lib "Kernel" (ByVal hMem As Integer, ByVal dwBytes As Long, ByVal wFlags As Integer) As Integer

'NOTE: instead of declaring the function GlobalDiscard and calling
'      GlobalDiscard(hMem), call GlobalReAlloc(hMem, 0, GMEM_MOVEABLE)

Declare Function GlobalSize Lib "Kernel" (ByVal hMem As Integer) As Long
Declare Function GlobalUnlock Lib "Kernel" (ByVal hMem As Integer) As Integer
Declare Function UnlockResource Lib "Kernel" Alias "GlobalUnlock" (ByVal hMem As Integer) As Integer
Declare Function GlobalFlags Lib "Kernel" (ByVal hMem As Integer) As Integer
Declare Function GlobalWire Lib "Kernel" (ByVal hMem As Integer) As Long
Declare Function GlobalUnWire Lib "Kernel" (ByVal hMem As Integer) As Integer
Declare Function GlobalLRUNewest Lib "Kernel" (ByVal hMem As Integer) As Integer
Declare Function GlobalLRUOldest Lib "Kernel" (ByVal hMem As Integer) As Integer
Declare Function GlobalPageLock Lib "Kernel" (ByVal wSelector As Integer) As Integer
Declare Function GlobalPageUnlock Lib "Kernel" (ByVal wSelector As Integer) As Integer
Declare Sub GlobalFix Lib "Kernel" (ByVal hMem As Integer)
Declare Function GlobalUnfix Lib "Kernel" (ByVal hMem As Integer) As Integer

' Flags returned by GlobalFlags (in addition to GMEM_DISCARDABLE)
Public Const GMEM_DISCARDED = &H4000
Public Const GMEM_LOCKCOUNT = &HFF

Declare Function LockSegment Lib "Kernel" (ByVal wSegment As Integer) As Integer
Declare Function UnlockSegment Lib "Kernel" (ByVal wSegment As Integer) As Integer

' Local Memory Flags
Public Const LMEM_FIXED = &H0
Public Const LMEM_MOVEABLE = &H2
Public Const LMEM_NOCOMPACT = &H10
Public Const LMEM_NODISCARD = &H20
Public Const LMEM_ZEROINIT = &H40
Public Const LMEM_MODIFY = &H80
Public Const LMEM_DISCARDABLE = &HF00

Public Const LHND = (LMEM_MOVEABLE + LMEM_ZEROINIT)
Public Const LPTR = (LMEM_FIXED + LMEM_ZEROINIT)

Public Const NONZEROLHND = (LMEM_MOVEABLE)
Public Const NONZEROLPTR = (LMEM_FIXED)

Public Const LNOTIFY_OUTOFMEM = 0
Public Const LNOTIFY_MOVE = 1
Public Const LNOTIFY_DISCARD = 2

Declare Function LocalAlloc Lib "Kernel" (ByVal wFlags As Integer, ByVal wBytes As Integer) As Integer
Declare Function LocalCompact Lib "Kernel" (ByVal wMinFree As Integer) As Integer
Declare Function LocalFree Lib "Kernel" (ByVal hMem As Integer) As Integer
Declare Function LocalHandle Lib "Kernel" (ByVal wMem As Integer) As Integer
Declare Function LocalInit Lib "Kernel" (ByVal wSegment As Integer, ByVal pStart As Integer, ByVal pEnd As Integer) As Integer
Declare Function LocalLock Lib "Kernel" (ByVal hMem As Integer) As Integer '(returns a near  pointer)
Declare Function LocalReAlloc Lib "Kernel" (ByVal hMem As Integer, ByVal wBytes As Integer, ByVal wFlags As Integer) As Integer

'NOTE: instead of declaring the function LocalDiscard and calling
'      LocalDiscard(hMem), call LocalReAlloc(hMem, 0, LMEM_MOVEABLE)

Declare Function LocalSize Lib "Kernel" (ByVal hMem As Integer) As Integer
Declare Function LocalUnlock Lib "Kernel" (ByVal hMem As Integer) As Integer
Declare Function LocalFlags Lib "Kernel" (ByVal hMem As Integer) As Integer
Declare Function LocalShrink Lib "Kernel" (ByVal hSeg As Integer, ByVal wSize As Integer) As Integer

' Flags returned by LocalFlags (in addition to LMEM_DISCARDABLE)
Public Const LMEM_DISCARDED = &H4000
Public Const LMEM_LOCKCOUNT = &HFF

Declare Function SetSwapAreaSize Lib "Kernel" (ByVal rsSize As Integer) As Long
Declare Function ValidateFreeSpaces Lib "Kernel" () As Long
Declare Sub LimitEmsPages Lib "Kernel" (ByVal dwKbytes As Long)
Declare Function SetErrorMode Lib "Kernel" (ByVal wMode As Integer) As Integer
Declare Sub ValidateCodeSegments Lib "Kernel" ()

Declare Function GetDOSEnvironment Lib "Kernel" () As Long

Declare Function AllocDStoCSAlias Lib "Kernel" (ByVal wSelector As Integer) As Integer
Declare Function AllocSelector Lib "Kernel" (ByVal wSelector As Integer) As Integer
Declare Function ChangeSelector Lib "Kernel" (ByVal wDestSelector As Integer, ByVal wSourceSelector As Integer) As Integer
Declare Function FreeSelector Lib "Kernel" (ByVal wSelector As Integer) As Integer
Declare Function FindResource Lib "Kernel" (ByVal hInstance As Integer, ByVal lpName As String, ByVal lpType As Any) As Integer
Declare Function LoadResource Lib "Kernel" (ByVal hInstance As Integer, ByVal hResInfo As Integer) As Integer
Declare Function FreeResource Lib "Kernel" (ByVal hResData As Integer) As Integer
Declare Function LockResource Lib "Kernel" (ByVal hResData As Integer) As Long
Declare Function AllocResource Lib "Kernel" (ByVal hInstance As Integer, ByVal hResInfo As Integer, ByVal dwSize As Long) As Integer
Declare Function SizeofResource Lib "Kernel" (ByVal hInstance As Integer, ByVal hResInfo As Integer) As Integer
Declare Function AccessResource Lib "Kernel" (ByVal hInstance As Integer, ByVal hResInfo As Integer) As Integer

' Predefined Resource Types
Public Const RT_CURSOR = 1&
Public Const RT_BITMAP = 2&
Public Const RT_ICON = 3&
Public Const RT_MENU = 4&
Public Const RT_DIALOG = 5&
Public Const RT_STRING = 6&
Public Const RT_FONTDIR = 7&
Public Const RT_FONT = 8&
Public Const RT_ACCELERATOR = 9&
Public Const RT_RCDATA = 10&

Declare Sub Yield Lib "Kernel" ()
Declare Function GetCurrentTask Lib "Kernel" () As Integer

Declare Function InitAtomTable Lib "Kernel" (ByVal nSize As Integer) As Integer
Declare Function AddAtom Lib "Kernel" (ByVal lpString As String) As Integer
Declare Function DeleteAtom Lib "Kernel" (ByVal nAtom As Integer) As Integer
Declare Function FindAtom Lib "Kernel" (ByVal lpString As String) As Integer
Declare Function GetAtomName Lib "Kernel" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function GlobalAddAtom Lib "User" (ByVal lpString As String) As Integer
Declare Function GlobalDeleteAtom Lib "User" (ByVal nAtom As Integer) As Integer
Declare Function GlobalFindAtom Lib "User" (ByVal lpString As String) As Integer
Declare Function GlobalGetAtomName Lib "User" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function GetAtomHandle Lib "Kernel" (ByVal wAtom As Integer) As Integer


' User Profile Routines
' NOTE: The lpKeyName argument for GetProfileString, WriteProfileString,
'       GetPrivateProfileString, and WritePrivateProfileString can be either
'       a string or NULL.  For example, to pass a string specify ByVal "wallpaper"
'       To pass NULL specify ByVal 0&
'       You can also pass NULL for the lpString argument for WriteProfileString
'       and WritePrivateProfileString
Declare Function GetProfileInt Lib "Kernel" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Integer) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer
Declare Function WriteProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any) As Integer
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFilename As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFilename As String) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Integer
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

'NOTE: Catch/Throw expect a long pointer to an 18-byte buffer (lpCatchBuf)
' eg:
'       Dim Buffer(1 To 9) As Integer
'       result% = Catch (Buffer(1))

Declare Function Catch Lib "Kernel" (lpCatchBuf As Any) As Integer
Declare Sub Throw Lib "Kernel" (lpCatchBuf As Any, ByVal nThrowBack As Integer)

Declare Sub FatalExit Lib "Kernel" (ByVal code As Integer)
Declare Sub SwapRecording Lib "Kernel" (ByVal wFlag As Integer)

' Character Translation Routines
Declare Function AnsiToOem Lib "Keyboard" (ByVal lpAnsiStr As String, ByVal lpOemStr As String) As Integer
Declare Function OemToAnsi Lib "Keyboard" (ByVal lpOemStr As String, ByVal lpAnsiStr As String) As Integer
Declare Sub AnsiToOemBuff Lib "Keyboard" (ByVal lpAnsiStr As String, ByVal lpOemStr As String, ByVal nLength As Integer)
Declare Sub OemToAnsiBuff Lib "Keyboard" (ByVal lpOemStr As String, ByVal lpAnsiStr As String, ByVal nLength As Integer)
Declare Function AnsiUpper Lib "User" (ByVal lpString As String) As String
Declare Function AnsiUpperBuff Lib "User" (ByVal lpString As String, ByVal aWORD As Integer) As Integer
Declare Function AnsiLower Lib "User" (ByVal lpString As String) As Long
Declare Function AnsiLowerBuff Lib "User" (ByVal lpString As String, ByVal aWORD As Integer) As Integer
Declare Function AnsiNext Lib "User" (ByVal lpString As String) As Long
Declare Function AnsiPrev Lib "User" (ByVal lpString As String, ByVal lpString As String) As Long

' Keyboard Information Routines
Declare Function OemKeyScan Lib "Keyboard" (ByVal wOemChar As Integer) As Long
Declare Function VkKeyScan Lib "Keyboard" (ByVal cChar As Integer) As Integer
Declare Function GetKeyboardType Lib "Keyboard" (ByVal nTypeFlag As Integer) As Integer
Declare Function MapVirtualKey Lib "Keyboard" (ByVal wCode As Integer, ByVal wMapType As Integer) As Integer
Declare Function GetKBCodePage Lib "Keyboard" ()
Declare Function GetKeyNameText Lib "Keyboard" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function ToAscii Lib "Keyboard" (ByVal wVirtKey As Integer, ByVal wScanCode As Integer, lpKeyState As Any, lpChar As Any, ByVal wFlags As Integer) As Integer

' Language dependent Routines
Declare Function IsCharAlpha Lib "User" (ByVal cChar As Integer) As Integer
Declare Function IsCharAlphaNumeric Lib "User" (ByVal cChar As Integer) As Integer
Declare Function IsCharUpper Lib "User" (ByVal cChar As Integer) As Integer
Declare Function IsCharLower Lib "User" (ByVal cChar As Integer) As Integer

Declare Function GetWinFlags Lib "Kernel" () As Long

Public Const WF_PMODE = &H1
Public Const WF_CPU286 = &H2
Public Const WF_CPU386 = &H4
Public Const WF_CPU486 = &H8
Public Const WF_STANDARD = &H10
Public Const WF_WIN286 = &H10
Public Const WF_ENHANCED = &H20
Public Const WF_WIN386 = &H20
Public Const WF_CPU086 = &H40
Public Const WF_CPU186 = &H80
Public Const WF_LARGEFRAME = &H100
Public Const WF_SMALLFRAME = &H200
Public Const WF_80x87 = &H400

' WEP fSystemExit flag values
Public Const WEP_SYSTEM_EXIT = 1
Public Const WEP_FREE_DLL = 0

' OEM Resource Ordinal Numbers
Public Const OBM_CLOSE = 32754
Public Const OBM_UPARROW = 32753
Public Const OBM_DNARROW = 32752
Public Const OBM_RGARROW = 32751
Public Const OBM_LFARROW = 32750
Public Const OBM_REDUCE = 32749
Public Const OBM_ZOOM = 32748
Public Const OBM_RESTORE = 32747
Public Const OBM_REDUCED = 32746
Public Const OBM_ZOOMD = 32745
Public Const OBM_RESTORED = 32744
Public Const OBM_UPARROWD = 32743
Public Const OBM_DNARROWD = 32742
Public Const OBM_RGARROWD = 32741
Public Const OBM_LFARROWD = 32740
Public Const OBM_MNARROW = 32739
Public Const OBM_COMBO = 32738

Public Const OBM_OLD_CLOSE = 32767
Public Const OBM_SIZE = 32766
Public Const OBM_OLD_UPARROW = 32765
Public Const OBM_OLD_DNARROW = 32764
Public Const OBM_OLD_RGARROW = 32763
Public Const OBM_OLD_LFARROW = 32762
Public Const OBM_BTSIZE = 32761
Public Const OBM_CHECK = 32760
Public Const OBM_CHECKBOXES = 32759
Public Const OBM_BTNCORNERS = 32758
Public Const OBM_OLD_REDUCE = 32757
Public Const OBM_OLD_ZOOM = 32756
Public Const OBM_OLD_RESTORE = 32755

Public Const OCR_NORMAL = 32512
Public Const OCR_IBEAM = 32513
Public Const OCR_WAIT = 32514
Public Const OCR_CROSS = 32515
Public Const OCR_UP = 32516
Public Const OCR_SIZE = 32640
Public Const OCR_ICON = 32641
Public Const OCR_SIZENWSE = 32642
Public Const OCR_SIZENESW = 32643
Public Const OCR_SIZEWE = 32644
Public Const OCR_SIZENS = 32645
Public Const OCR_SIZEALL = 32646
Public Const OCR_ICOCUR = 32647

Public Const OIC_SAMPLE = 32512
Public Const OIC_HAND = 32513
Public Const OIC_QUES = 32514
Public Const OIC_BANG = 32515
Public Const OIC_NOTE = 32516


' -------------
'  GDI Section
' -------------

' Binary raster ops
Public Const R2_BLACK = 1       '   0
Public Const R2_NOTMERGEPEN = 2 '  DPon
Public Const R2_MASKNOTPEN = 3  '  DPna
Public Const R2_NOTCOPYPEN = 4  '  PN
Public Const R2_MASKPENNOT = 5  '  PDna
Public Const R2_NOT = 6 '  Dn
Public Const R2_XORPEN = 7      '  DPx
Public Const R2_NOTMASKPEN = 8  '  DPan
Public Const R2_MASKPEN = 9     '  DPa
Public Const R2_NOTXORPEN = 10  '  DPxn
Public Const R2_NOP = 11        '  D
Public Const R2_MERGENOTPEN = 12        '  DPno
Public Const R2_COPYPEN = 13    '  P
Public Const R2_MERGEPENNOT = 14        '  PDno
Public Const R2_MERGEPEN = 15   '  DPo
Public Const R2_WHITE = 16      '   1

'  Ternary raster operations
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Public Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)
Public Const MERGECOPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
Public Const MERGEPAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
Public Const PATCOPY = &HF00021 ' (DWORD) dest = pattern
Public Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
Public Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Public Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
Public Const BLACKNESS = &H42&  ' (DWORD) dest = BLACK
Public Const WHITENESS = &HFF0062       ' (DWORD) dest = WHITE


' StretchBlt() Modes
Public Const BLACKONWHITE = 1
Public Const WHITEONBLACK = 2
Public Const COLORONCOLOR = 3

' PolyFill() Modes
Public Const ALTERNATE = 1
Public Const WINDING = 2

' Text Alignment Options
Public Const TA_NOUPDATECP = 0
Public Const TA_UPDATECP = 1

Public Const TA_LEFT = 0
Public Const TA_RIGHT = 2
Public Const TA_CENTER = 6

Public Const TA_TOP = 0
Public Const TA_BOTTOM = 8
Public Const TA_BASELINE = 24

Public Const ETO_GRAYED = 1
Public Const ETO_OPAQUE = 2
Public Const ETO_CLIPPED = 4

Public Const ASPECT_FILTERING = &H1

' Metafile Functions
Public Const META_SETBKCOLOR = &H201
Public Const META_SETBKMODE = &H102
Public Const META_SETMAPMODE = &H103
Public Const META_SETROP2 = &H104
Public Const META_SETRELABS = &H105
Public Const META_SETPOLYFILLMODE = &H106
Public Const META_SETSTRETCHBLTMODE = &H107
Public Const META_SETTEXTCHAREXTRA = &H108
Public Const META_SETTEXTCOLOR = &H209
Public Const META_SETTEXTJUSTIFICATION = &H20A
Public Const META_SETWINDOWORG = &H20B
Public Const META_SETWINDOWEXT = &H20C
Public Const META_SETVIEWPORTORG = &H20D
Public Const META_SETVIEWPORTEXT = &H20E
Public Const META_OFFSETWINDOWORG = &H20F
Public Const META_SCALEWINDOWEXT = &H400
Public Const META_OFFSETVIEWPORTORG = &H211
Public Const META_SCALEVIEWPORTEXT = &H412
Public Const META_LINETO = &H213
Public Const META_MOVETO = &H214
Public Const META_EXCLUDECLIPRECT = &H415
Public Const META_INTERSECTCLIPRECT = &H416
Public Const META_ARC = &H817
Public Const META_ELLIPSE = &H418
Public Const META_FLOODFILL = &H419
Public Const META_PIE = &H81A
Public Const META_RECTANGLE = &H41B
Public Const META_ROUNDRECT = &H61C
Public Const META_PATBLT = &H61D
Public Const META_SAVEDC = &H1E
Public Const META_SETPIXEL = &H41F
Public Const META_OFFSETCLIPRGN = &H220
Public Const META_TEXTOUT = &H521
Public Const META_BITBLT = &H922
Public Const META_STRETCHBLT = &HB23
Public Const META_POLYGON = &H324
Public Const META_POLYLINE = &H325
Public Const META_ESCAPE = &H626
Public Const META_RESTOREDC = &H127
Public Const META_FILLREGION = &H228
Public Const META_FRAMEREGION = &H429
Public Const META_INVERTREGION = &H12A
Public Const META_PAINTREGION = &H12B
Public Const META_SELECTCLIPREGION = &H12C
Public Const META_SELECTOBJECT = &H12D
Public Const META_SETTEXTALIGN = &H12E
Public Const META_DRAWTEXT = &H62F

Public Const META_CHORD = &H830
Public Const META_SETMAPPERFLAGS = &H231
Public Const META_EXTTEXTOUT = &HA32
Public Const META_SETDIBTODEV = &HD33
Public Const META_SELECTPALETTE = &H234
Public Const META_REALIZEPALETTE = &H35
Public Const META_ANIMATEPALETTE = &H436
Public Const META_SETPALENTRIES = &H37
Public Const META_POLYPOLYGON = &H538
Public Const META_RESIZEPALETTE = &H139

Public Const META_DIBBITBLT = &H940
Public Const META_DIBSTRETCHBLT = &HB41
Public Const META_DIBCREATEPATTERNBRUSH = &H142
Public Const META_STRETCHDIB = &HF43

Public Const META_DELETEOBJECT = &H1F0

Public Const META_CREATEPALETTE = &HF7
Public Const META_CREATEBRUSH = &HF8
Public Const META_CREATEPATTERNBRUSH = &H1F9
Public Const META_CREATEPENINDIRECT = &H2FA
Public Const META_CREATEFONTINDIRECT = &H2FB
Public Const META_CREATEBRUSHINDIRECT = &H2FC
Public Const META_CREATEBITMAPINDIRECT = &H2FD
Public Const META_CREATEBITMAP = &H6FE
Public Const META_CREATEREGION = &H6FF


' GDI Escapes
Public Const NEWFRAME = 1
'Public Const ABORTDOC = 2
Public Const NEXTBAND = 3
Public Const SETCOLORTABLE = 4
Public Const GETCOLORTABLE = 5
Public Const FLUSHOUTPUT = 6
Public Const DRAFTMODE = 7
Public Const QUERYESCSUPPORT = 8
Public Const SETABORTPROC = 9
'Public Const STARTDOC = 10
'Public Const ENDDOCAPI = 11
Public Const GETPHYSPAGESIZE = 12
Public Const GETPRINTINGOFFSET = 13
Public Const GETSCALINGFACTOR = 14
Public Const MFCOMMENT = 15
Public Const GETPENWIDTH = 16
Public Const SETCOPYCOUNT = 17
Public Const SELECTPAPERSOURCE = 18
Public Const DEVICEDATA = 19
Public Const PASSTHROUGH = 19
Public Const GETTECHNOLGY = 20
Public Const GETTECHNOLOGY = 20
Public Const SETENDCAP = 21
Public Const SETLINEJOIN = 22
Public Const SETMITERLIMIT = 23
Public Const BANDINFO = 24
Public Const DRAWPATTERNRECT = 25
Public Const GETVECTORPENSIZE = 26
Public Const GETVECTORBRUSHSIZE = 27
Public Const ENABLEDUPLEX = 28
Public Const GETSETPAPERBINS = 29
Public Const GETSETPRINTORIENT = 30
Public Const ENUMPAPERBINS = 31
Public Const SETDIBSCALING = 32
Public Const EPSPRINTING = 33
Public Const ENUMPAPERMETRICS = 34
Public Const GETSETPAPERMETRICS = 35
Public Const POSTSCRIPT_DATA = 37
Public Const POSTSCRIPT_IGNORE = 38
Public Const GETEXTENDEDTEXTMETRICS = 256
Public Const GETEXTENTTABLE = 257
Public Const GETPAIRKERNTABLE = 258
Public Const GETTRACKKERNTABLE = 259
Public Const EXTTEXTOUTAPI = 512
Public Const ENABLERELATIVEWIDTHS = 768
Public Const ENABLEPAIRKERNING = 769
Public Const SETKERNTRACK = 770
Public Const SETALLJUSTVALUES = 771
Public Const SETCHARSET = 772

Public Const STRETCHBLTAPI = 2048
Public Const BEGIN_PATH = 4096
Public Const CLIP_TO_PATH = 4097
Public Const END_PATH = 4098
Public Const EXT_DEVICE_CAPS = 4099
Public Const RESTORE_CTM = 4100
Public Const SAVE_CTM = 4101
Public Const SET_ARC_DIRECTION = 4102
Public Const SET_BACKGROUND_COLOR = 4103
Public Const SET_POLY_MODE = 4104
Public Const SET_SCREEN_ANGLE = 4105
Public Const SET_SPREAD = 4106
Public Const TRANSFORM_CTM = 4107
Public Const SET_CLIP_BOX = 4108
Public Const SET_BOUNDS = 4109
Public Const SET_MIRROR_MODE = 4110

' Spooler Error Codes
Public Const SP_NOTREPORTED = &H4000
Public Const SP_ERROR = (-1)
Public Const SP_APPABORT = (-2)
Public Const SP_USERABORT = (-3)
Public Const SP_OUTOFDISK = (-4)
Public Const SP_OUTOFMEMORY = (-5)

Public Const PR_JOBSTATUS = &H0

' Bitmap Header Definition
Type BITMAP '14 bytes
        bmType As Integer
        bmWidth As Integer
        bmHeight As Integer
        bmWidthBytes As Integer
        bmPlanes As String * 1
        bmBitsPixel As String * 1
        bmBits As Long
End Type

Type RGBTRIPLE
        rgbtBlue As String * 1
        rgbtGreen As String * 1
        rgbtRed As String * 1
End Type

Type RGBQUAD
        rgbBlue As String * 1
        rgbGreen As String * 1
        rgbRed As String * 1
        rgbReserved As String * 1
End Type

' structures for defining DIBs
Type BITMAPCOREHEADER '12 bytes
        bcSize As Long
        bcWidth As Integer
        bcHeight As Integer
        bcPlanes As Integer
        bcBitCount As Integer
End Type

Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

' constants for the biCompression field
Public Const BI_RGB = 0&
Public Const BI_RLE8 = 1&
Public Const BI_RLE4 = 2&

Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As String * 128 ' Array length is arbitrary; may be changed
End Type

Type BITMAPCOREINFO
        bmciHeader As BITMAPCOREHEADER
        bmciColors As String * 96 ' Array length is arbitrary; may be changed
End Type

Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type


' Clipboard Metafile Picture Structure
Type HANDLETABLE
        objectHandle As String * 512 ' Array length is arbitrary; may be changed
End Type

Type METARECORD
        rdSize As Long
        rdFunction As Integer
        rdParm As String * 512 ' Array length is arbitrary; may be changed
End Type


Type METAFILEPICT
        mm As Integer
        xExt As Integer
        yExt As Integer
        hMF As Integer
End Type

Type METAHEADER
        mtType As Integer
        mtHeaderSize As Integer
        mtVersion As Integer
        mtSize As Long
        mtNoObjects As Integer
        mtMaxRecord As Long
        mtNoParameters As Integer
End Type

Type TEXTMETRIC
        tmHeight As Integer
        tmAscent As Integer
        tmDescent As Integer
        tmInternalLeading As Integer
        tmExternalLeading As Integer
        tmAveCharWidth As Integer
        tmMaxCharWidth As Integer
        tmWeight As Integer
        tmItalic As String * 1
        tmUnderlined As String * 1
        tmStruckOut As String * 1
        tmFirstChar As String * 1
        tmLastChar As String * 1
        tmDefaultChar As String * 1
        tmBreakChar As String * 1
        tmPitchAndFamily As String * 1
        tmCharSet As String * 1
        tmOverhang As Integer
        tmDigitizedAspectX As Integer
        tmDigitizedAspectY As Integer
End Type

' GDI Logical Objects:

'  Pel Array
Type PELARRAY
        paXCount As Integer
        paYCount As Integer
        paXExt As Integer
        paYExt As Integer
        paRGBs As Integer
End Type

' Logical Brush (or Pattern)
Type LOGBRUSH
        lbStyle As Integer
        lbColor As Long
        lbHatch As Integer
End Type

' Logical Pen
Type LOGPEN
        lopnStyle As Integer
        lopnWidth As POINTAPI
        lopnColor As Long
End Type

Type PALETTEENTRY
        peRed As String * 1
        peGreen As String * 1
        peBlue As String * 1
        peFlags As String * 1
End Type

' Logical Palette
Type LOGPALETTE
        palVersion As Integer
        palNumEntries As Integer
        palPalEntry As String * 252 ' Array length is arbitrary; may be changed
End Type

' Logical Font
Public Const LF_FACESIZE = 32

Type LOGFONT
        lfHeight As Integer
        lfWidth As Integer
        lfEscapement As Integer
        lfOrientation As Integer
        lfWeight As Integer
        lfItalic As String * 1
        lfUnderline As String * 1
        lfStrikeOut As String * 1
        lfCharSet As String * 1
        lfOutPrecision As String * 1
        lfClipPrecision As String * 1
        lfQuality As String * 1
        lfPitchAndFamily As String * 1
        lfFaceName As String * LF_FACESIZE
End Type

Public Const OUT_DEFAULT_PRECIS = 0
Public Const OUT_STRING_PRECIS = 1
Public Const OUT_CHARACTER_PRECIS = 2
Public Const OUT_STROKE_PRECIS = 3

Public Const CLIP_DEFAULT_PRECIS = 0
Public Const CLIP_CHARACTER_PRECIS = 1
Public Const CLIP_STROKE_PRECIS = 2

Public Const DEFAULT_QUALITY = 0
Public Const DRAFT_QUALITY = 1
Public Const PROOF_QUALITY = 2

Public Const DEFAULT_PITCH = 0
Public Const FIXED_PITCH = 1
Public Const VARIABLE_PITCH = 2

Public Const ANSI_CHARSET = 0
Public Const SYMBOL_CHARSET = 2
Public Const SHIFTJIS_CHARSET = 128
Public Const OEM_CHARSET = 255

' Font Families
'
Public Const FF_DONTCARE = 0    '  Don't care or don't know.
Public Const FF_ROMAN = 16      '  Variable stroke width, serifed.

' Times Roman, Century Schoolbook, etc.
Public Const FF_SWISS = 32      '  Variable stroke width, sans-serifed.

' Helvetica, Swiss, etc.
Public Const FF_MODERN = 48     '  Constant stroke width, serifed or sans-serifed.

' Pica, Elite, Courier, etc.
Public Const FF_SCRIPT = 64     '  Cursive, etc.
Public Const FF_DECORATIVE = 80 '  Old English, etc.

' Font Weights
Public Const FW_DONTCARE = 0
Public Const FW_THIN = 100
Public Const FW_EXTRALIGHT = 200
Public Const FW_LIGHT = 300
Public Const FW_NORMAL = 400
Public Const FW_MEDIUM = 500
Public Const FW_SEMIBOLD = 600
Public Const FW_BOLD = 700
Public Const FW_EXTRABOLD = 800
Public Const FW_HEAVY = 900

Public Const FW_ULTRALIGHT = FW_EXTRALIGHT
Public Const FW_REGULAR = FW_NORMAL
Public Const FW_DEMIBOLD = FW_SEMIBOLD
Public Const FW_ULTRABOLD = FW_EXTRABOLD
Public Const FW_BLACK = FW_HEAVY

' Background Modes
Public Const TRANSPARENT = 1
Public Const OPAQUE = 2

' Mapping Modes
Public Const MM_TEXT = 1
Public Const MM_LOMETRIC = 2
Public Const MM_HIMETRIC = 3
Public Const MM_LOENGLISH = 4
Public Const MM_HIENGLISH = 5
Public Const MM_TWIPS = 6
Public Const MM_ISOTROPIC = 7
Public Const MM_ANISOTROPIC = 8

' Coordinate Modes
Public Const ABSOLUTE = 1
Public Const RELATIVE = 2

' Stock Logical Objects
Public Const WHITE_BRUSH = 0
Public Const LTGRAY_BRUSH = 1
Public Const GRAY_BRUSH = 2
Public Const DKGRAY_BRUSH = 3
Public Const BLACK_BRUSH = 4
Public Const NULL_BRUSH = 5
Public Const HOLLOW_BRUSH = NULL_BRUSH
Public Const WHITE_PEN = 6
Public Const BLACK_PEN = 7
Public Const NULL_PEN = 8
Public Const OEM_FIXED_FONT = 10
Public Const ANSI_FIXED_FONT = 11
Public Const ANSI_VAR_FONT = 12
Public Const SYSTEM_FONT = 13
Public Const DEVICE_DEFAULT_FONT = 14
Public Const DEFAULT_PALETTE = 15
Public Const SYSTEM_FIXED_FONT = 16

' Brush Styles
Public Const BS_SOLID = 0
Public Const BS_NULL = 1
Public Const BS_HOLLOW = BS_NULL
Public Const BS_HATCHED = 2
Public Const BS_PATTERN = 3
Public Const BS_INDEXED = 4
Public Const BS_DIBPATTERN = 5

' Hatch Styles
Public Const HS_HORIZONTAL = 0  '  -----
Public Const HS_VERTICAL = 1    '  |||||
Public Const HS_FDIAGONAL = 2   '  \\\\\
Public Const HS_BDIAGONAL = 3   '  /////
Public Const HS_CROSS = 4       '  +++++
Public Const HS_DIAGCROSS = 5   '  xxxxx

' Pen Styles
Public Const PS_SOLID = 0
Public Const PS_DASH = 1        '  -------
Public Const PS_DOT = 2 '  .......
Public Const PS_DASHDOT = 3     '  _._._._
Public Const PS_DASHDOTDOT = 4  '  _.._.._
Public Const PS_NULL = 5
Public Const PS_INSIDEFRAME = 6

' Device Parameters for GetDeviceCaps()
Public Const DRIVERVERSION = 0  '  Device driver version
Public Const TECHNOLOGY = 2     '  Device classification
Public Const HORZSIZE = 4       '  Horizontal size in millimeters
Public Const VERTSIZE = 6       '  Vertical size in millimeters
Public Const HORZRES = 8        '  Horizontal width in pixels
Public Const VERTRES = 10       '  Vertical width in pixels
Public Const BITSPIXEL = 12     '  Number of bits per pixel
Public Const PLANES = 14        '  Number of planes
Public Const NUMBRUSHES = 16    '  Number of brushes the device has
Public Const NUMPENS = 18       '  Number of pens the device has
Public Const NUMMARKERS = 20    '  Number of markers the device has
Public Const NUMFONTS = 22      '  Number of fonts the device has
Public Const NUMCOLORS = 24     '  Number of colors the device supports
Public Const PDEVICESIZE = 26   '  Size required for device descriptor
Public Const CURVECAPS = 28     '  Curve capabilities
Public Const LINECAPS = 30      '  Line capabilities
Public Const POLYGONALCAPS = 32 '  Polygonal capabilities
Public Const TEXTCAPS = 34      '  Text capabilities
Public Const CLIPCAPS = 36      '  Clipping capabilities
Public Const RASTERCAPS = 38    '  Bitblt capabilities
Public Const ASPECTX = 40       '  Length of the X leg
Public Const ASPECTY = 42       '  Length of the Y leg
Public Const ASPECTXY = 44      '  Length of the hypotenuse

Public Const LOGPIXELSX = 88    '  Logical pixels/inch in X
Public Const LOGPIXELSY = 90    '  Logical pixels/inch in Y

Public Const SIZEPALETTE = 104  '  Number of entries in physical palette
Public Const NUMRESERVED = 106  '  Number of reserved entries in palette
Public Const COLORRES = 108     '  Actual color resolution


' Device Capability Masks:

' Device Technologies
Public Const DT_PLOTTER = 0     '  Vector plotter
Public Const DT_RASDISPLAY = 1  '  Raster display
Public Const DT_RASPRINTER = 2  '  Raster printer
Public Const DT_RASCAMERA = 3   '  Raster camera
Public Const DT_CHARSTREAM = 4  '  Character-stream, PLP
Public Const DT_METAFILE = 5    '  Metafile, VDM
Public Const DT_DISPFILE = 6    '  Display-file

' Curve Capabilities
Public Const CC_NONE = 0        '  Curves not supported
Public Const CC_CIRCLES = 1     '  Can do circles
Public Const CC_PIE = 2 '  Can do pie wedges
Public Const CC_CHORD = 4       '  Can do chord arcs
Public Const CC_ELLIPSES = 8    '  Can do ellipese
Public Const CC_WIDE = 16       '  Can do wide lines
Public Const CC_STYLED = 32     '  Can do styled lines
Public Const CC_WIDESTYLED = 64 '  Can do wide styled lines
Public Const CC_INTERIORS = 128 '  Can do interiors

' Line Capabilities
Public Const LC_NONE = 0        '  Lines not supported
Public Const LC_POLYLINE = 2    '  Can do polylines
Public Const LC_MARKER = 4      '  Can do markers
Public Const LC_POLYMARKER = 8  '  Can do polymarkers
Public Const LC_WIDE = 16       '  Can do wide lines
Public Const LC_STYLED = 32     '  Can do styled lines
Public Const LC_WIDESTYLED = 64 '  Can do wide styled lines
Public Const LC_INTERIORS = 128 '  Can do interiors

' Polygonal Capabilities
Public Const PC_NONE = 0        '  Polygonals not supported
Public Const PC_POLYGON = 1     '  Can do polygons
Public Const PC_RECTANGLE = 2   '  Can do rectangles
Public Const PC_WINDPOLYGON = 4 '  Can do winding polygons
Public Const PC_TRAPEZOID = 4   '  Can do trapezoids
Public Const PC_SCANLINE = 8    '  Can do scanlines
Public Const PC_WIDE = 16       '  Can do wide borders
Public Const PC_STYLED = 32     '  Can do styled borders
Public Const PC_WIDESTYLED = 64 '  Can do wide styled borders
Public Const PC_INTERIORS = 128 '  Can do interiors

' Polygonal Capabilities
Public Const CP_NONE = 0        '  No clipping of output
Public Const CP_RECTANGLE = 1   '  Output clipped to rects

' Text Capabilities
Public Const TC_OP_CHARACTER = &H1      '  Can do OutputPrecision       CHARACTER
Public Const TC_OP_STROKE = &H2 '  Can do OutputPrecision       STROKE
Public Const TC_CP_STROKE = &H4 '  Can do ClipPrecision STROKE
Public Const TC_CR_90 = &H8     '  Can do CharRotAbility        90
Public Const TC_CR_ANY = &H10   '  Can do CharRotAbility        ANY
Public Const TC_SF_X_YINDEP = &H20      '  Can do ScaleFreedom  X_YINDEPENDENT
Public Const TC_SA_DOUBLE = &H40        '  Can do ScaleAbility  DOUBLE
Public Const TC_SA_INTEGER = &H80       '  Can do ScaleAbility  INTEGER
Public Const TC_SA_CONTIN = &H100       '  Can do ScaleAbility  CONTINUOUS
Public Const TC_EA_DOUBLE = &H200       '  Can do EmboldenAbility       DOUBLE
Public Const TC_IA_ABLE = &H400 '  Can do ItalisizeAbility      ABLE
Public Const TC_UA_ABLE = &H800 '  Can do UnderlineAbility      ABLE
Public Const TC_SO_ABLE = &H1000        '  Can do StrikeOutAbility      ABLE
Public Const TC_RA_ABLE = &H2000        '  Can do RasterFontAble        ABLE
Public Const TC_VA_ABLE = &H4000        '  Can do VectorFontAble        ABLE
Public Const TC_RESERVED = &H8000

' Raster Capabilities
Public Const RC_BITBLT = 1      '  Can do standard BLT.
Public Const RC_BANDING = 2     '  Device requires banding support
Public Const RC_SCALING = 4     '  Device requires scaling support
Public Const RC_BITMAP64 = 8    '  Device can support >64K bitmap
Public Const RC_GDI20_OUTPUT = &H10     '  has 2.0 output calls
Public Const RC_DI_BITMAP = &H80        '  supports DIB to memory
Public Const RC_PALETTE = &H100 '  supports a palette
Public Const RC_DIBTODEV = &H200        '  supports DIBitsToDevice
Public Const RC_BIGFONT = &H400 '  supports >64K fonts
Public Const RC_STRETCHBLT = &H800      '  supports StretchBlt
Public Const RC_FLOODFILL = &H1000      '  supports FloodFill
Public Const RC_STRETCHDIB = &H2000     '  supports StretchDIBits

' palette entry flags
Public Const PC_RESERVED = &H1  '  palette index used for animation
Public Const PC_EXPLICIT = &H2  '  palette index is explicit to device
Public Const PC_NOCOLLAPSE = &H4        '  do not match color to system palette

' DIB color table identifiers
Public Const DIB_RGB_COLORS = 0 '  color table in RGBTriples
Public Const DIB_PAL_COLORS = 1 '  color table in palette indices

' constants for Get/SetSystemPaletteUse()
Public Const SYSPAL_STATIC = 1
Public Const SYSPAL_NOSTATIC = 2

' constants for CreateDIBitmap
Public Const CBM_INIT = &H4&    '  initialize bitmap

' DrawText() Format Flags
Public Const DT_TOP = &H0
Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_BOTTOM = &H8
Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20
Public Const DT_EXPANDTABS = &H40
Public Const DT_TABSTOP = &H80
Public Const DT_NOCLIP = &H100
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_CALCRECT = &H400
Public Const DT_NOPREFIX = &H800
Public Const DT_INTERNAL = &H1000

Declare Function DrawText Lib "User" (ByVal hdc As Integer, ByVal lpStr As String, ByVal nCount As Integer, lpRect As RECT, ByVal wFormat As Integer) As Integer
Declare Function DrawIcon Lib "User" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal hicon As Integer) As Integer

' ExtFloodFill style flags
Public Const FLOODFILLBORDER = 0
Public Const FLOODFILLSURFACE = 1

Declare Function GetWindowDC Lib "User" (ByVal hwnd As Integer) As Integer
Declare Function GetDC Lib "User" (ByVal hwnd As Integer) As Integer
Declare Function ReleaseDC Lib "User" (ByVal hwnd As Integer, ByVal hdc As Integer) As Integer
Declare Function CreateDC Lib "GDI" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As String) As Integer
Declare Function CreateIC Lib "GDI" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As String) As Integer
Declare Function CreateCompatibleDC Lib "GDI" (ByVal hdc As Integer) As Integer
Declare Function DeleteDC Lib "GDI" (ByVal hdc As Integer) As Integer
Declare Function SaveDC Lib "GDI" (ByVal hdc As Integer) As Integer
Declare Function RestoreDC Lib "GDI" (ByVal hdc As Integer, ByVal nSavedDC As Integer) As Integer
Declare Function MoveTo Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer) As Long
Declare Function GetCurrentPosition Lib "GDI" (ByVal hdc As Integer) As Long
Declare Function LineTo Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer) As Integer
Declare Function GetDCOrg Lib "GDI" (ByVal hdc As Integer) As Long

Declare Function MulDiv Lib "GDI" (ByVal nNumber As Integer, ByVal nNumerator As Integer, ByVal nDenominator As Integer) As Integer

Declare Function ExtTextOut Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal wOptions As Integer, lpRect As Any, ByVal lpString As String, ByVal nCount As Integer, lpDx As Any) As Integer

Declare Function Polyline Lib "GDI" (ByVal hdc As Integer, lpPoints As POINTAPI, ByVal nCount As Integer) As Integer
Declare Function Polygon Lib "GDI" (ByVal hdc As Integer, lpPoints As POINTAPI, ByVal nCount As Integer) As Integer
Declare Function PolyPolygon Lib "GDI" (ByVal hdc As Integer, lpPoints As POINTAPI, lpPolyCounts As Integer, ByVal nCount As Integer) As Integer

Declare Function Rectangle Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
Declare Function RoundRect Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal X3 As Integer, ByVal Y3 As Integer) As Integer
Declare Function Ellipse Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
Declare Function Arc Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal X3 As Integer, ByVal Y3 As Integer, ByVal X4 As Integer, ByVal Y4 As Integer) As Integer
Declare Function Chord Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal X3 As Integer, ByVal Y3 As Integer, ByVal X4 As Integer, ByVal Y4 As Integer) As Integer
Declare Function Pie Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal X3 As Integer, ByVal Y3 As Integer, ByVal X4 As Integer, ByVal Y4 As Integer) As Integer
Declare Function PatBlt Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal dwRop As Long) As Integer
Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
Declare Function StretchBlt% Lib "GDI" (ByVal hdc%, ByVal x%, ByVal y%, ByVal nWidth%, ByVal nHeight%, ByVal hSrcDC%, ByVal XSrc%, ByVal YSrc%, ByVal nSrcWidth%, ByVal nSrcHeight%, ByVal dwRop&)
Declare Function TextOut Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal lpString As String, ByVal nCount As Integer) As Integer
Declare Function TabbedTextOut Lib "User" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal lpString As String, ByVal nCount As Integer, ByVal nTabPositions As Integer, lpnTabStopPositions As Integer, ByVal nTabOrigin As Integer) As Long
Declare Function GetCharWidth Lib "GDI" (ByVal hdc As Integer, ByVal wFirstChar As Integer, ByVal wLastChar As Integer, lpBuffer As Integer) As Integer
Declare Function SetPixel Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer) As Long
Declare Function FloodFill Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal crColor As Long) As Integer
Declare Function ExtFloodFill Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal crColor As Long, ByVal wFillType As Integer) As Integer

Declare Function GetStockObject Lib "GDI" (ByVal nIndex As Integer) As Integer

Declare Function CreatePen Lib "GDI" (ByVal nPenStyle As Integer, ByVal nWidth As Integer, ByVal crColor As Long) As Integer
Declare Function CreatePenIndirect Lib "GDI" (lpLogPen As LOGPEN) As Integer

Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function CreateHatchBrush Lib "GDI" (ByVal nIndex As Integer, ByVal crColor As Long) As Integer
Declare Function SetBrushOrg Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer) As Long
Declare Function GetBrushOrg Lib "GDI" (ByVal hdc As Integer) As Long
Declare Function CreatePatternBrush Lib "GDI" (ByVal hBitmap As Integer) As Integer
Declare Function CreateBrushIndirect Lib "GDI" (lpLogBrush As LOGBRUSH) As Integer

Declare Function CreateBitmap Lib "GDI" (ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal nPlanes As Integer, ByVal nBitCount As Integer, ByVal lpBits As Any) As Integer
Declare Function CreateBitmapIndirect Lib "GDI" (lpBitmap As BITMAP) As Integer
Declare Function CreateCompatibleBitmap Lib "GDI" (ByVal hdc As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
Declare Function CreateDiscardableBitmap Lib "GDI" (ByVal hdc As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer

Declare Function SetBitmapBits Lib "GDI" (ByVal hBitmap As Integer, ByVal dwCount As Long, ByVal lpBits As Any) As Long
Declare Function GetBitmapBits Lib "GDI" (ByVal hBitmap As Integer, ByVal dwCount As Long, ByVal lpBits As Any) As Long
Declare Function SetBitmapDimension Lib "GDI" (ByVal hBitmap As Integer, ByVal x As Integer, ByVal y As Integer) As Long
Declare Function GetBitmapDimension Lib "GDI" (ByVal hBitmap As Integer) As Long

Declare Function CreateFont% Lib "GDI" (ByVal h%, ByVal W%, ByVal E%, ByVal O%, ByVal W%, ByVal I%, ByVal U%, ByVal S%, ByVal C%, ByVal OP%, ByVal CP%, ByVal Q%, ByVal PAF%, ByVal F$)
Declare Function CreateFontIndirect Lib "GDI" (lpLogFont As LOGFONT) As Integer

Declare Function SelectClipRgn Lib "GDI" (ByVal hdc As Integer, ByVal hRgn As Integer) As Integer
Declare Function CreateRectRgn Lib "GDI" (ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
Declare Sub SetRectRgn Lib "GDI" (ByVal hRgn As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Function CreateRectRgnIndirect Lib "GDI" (lpRect As RECT) As Integer
Declare Function CreateEllipticRgnIndirect Lib "GDI" (lpRect As RECT) As Integer
Declare Function CreateEllipticRgn Lib "GDI" (ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
Declare Function CreatePolygonRgn Lib "GDI" (lpPoints As POINTAPI, ByVal nCount As Integer, ByVal nPolyFillMode As Integer) As Integer
Declare Function CreatePolyPolygonRgn Lib "GDI" (lpPoints As POINTAPI, lpPolyCounts As Integer, ByVal nCount As Integer, ByVal nPolyFillMode As Integer) As Integer
Declare Function CreateRoundRectRgn Lib "GDI" (ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer, ByVal X3 As Integer, ByVal Y3 As Integer) As Integer

'Declare Function GetObject Lib "GDI" (ByVal hObject As Integer, ByVal nCount As Integer, ByVal lpObject As Long) As Integer
Declare Function GetObjectGDI Lib "GDI" Alias "GetObject" (ByVal hObject As Integer, ByVal nCount As Integer, ByVal lpObject As Long) As Integer

Declare Function DeleteObject Lib "GDI" (ByVal hObject As Integer) As Integer
Declare Function SelectObject Lib "GDI" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer
Declare Function UnrealizeObject Lib "GDI" (ByVal hObject As Integer) As Integer

Declare Function SetBkColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long) As Long
Declare Function GetBkColor Lib "GDI" (ByVal hdc As Integer) As Long
Declare Function SetBkMode Lib "GDI" (ByVal hdc As Integer, ByVal nBkMode As Integer) As Integer
Declare Function GetBkMode Lib "GDI" (ByVal hdc As Integer) As Integer
Declare Function SetTextColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long) As Long
Declare Function GetTextColor Lib "GDI" (ByVal hdc As Integer) As Long
Declare Function SetTextAlign Lib "GDI" (ByVal hdc As Integer, ByVal wFlags As Integer) As Integer
Declare Function GetTextAlign Lib "GDI" (ByVal hdc As Integer) As Integer
Declare Function SetMapperFlags Lib "GDI" (ByVal hdc As Integer, ByVal dwFlag As Long) As Long
Declare Function GetAspectRatioFilter Lib "GDI" (ByVal hdc As Integer) As Long
Declare Function GetNearestColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long) As Long
Declare Function SetROP2 Lib "GDI" (ByVal hdc As Integer, ByVal nDrawMode As Integer) As Integer
Declare Function GetROP2 Lib "GDI" (ByVal hdc As Integer) As Integer
Declare Function SetStretchBltMode Lib "GDI" (ByVal hdc As Integer, ByVal nStretchMode As Integer) As Integer
Declare Function GetStretchBltMode Lib "GDI" (ByVal hdc As Integer) As Integer
Declare Function SetPolyFillMode Lib "GDI" (ByVal hdc As Integer, ByVal nPolyFillMode As Integer) As Integer
Declare Function GetPolyFillMode Lib "GDI" (ByVal hdc As Integer) As Integer
Declare Function SetMapMode Lib "GDI" (ByVal hdc As Integer, ByVal nMapMode As Integer) As Integer
Declare Function GetMapMode Lib "GDI" (ByVal hdc As Integer) As Integer
Declare Function SetWindowOrg Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer) As Long
Declare Function GetWindowOrg Lib "GDI" (ByVal hdc As Integer) As Long
Declare Function SetWindowExt Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer) As Long
Declare Function GetWindowExt Lib "GDI" (ByVal hdc As Integer) As Long
Declare Function SetViewportOrg Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer) As Long
Declare Function GetViewportOrg Lib "GDI" (ByVal hdc As Integer) As Long
Declare Function SetViewportExt Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer) As Long
Declare Function GetViewportExt Lib "GDI" (ByVal hdc As Integer) As Long
Declare Function OffsetViewportOrg Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer) As Long
Declare Function ScaleViewportExt Lib "GDI" (ByVal hdc As Integer, ByVal Xnum As Integer, ByVal Xdenom As Integer, ByVal Ynum As Integer, ByVal Ydenom As Integer) As Long
Declare Function OffsetWindowOrg Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer) As Long
Declare Function ScaleWindowExt Lib "GDI" (ByVal hdc As Integer, ByVal Xnum As Integer, ByVal Xdenom As Integer, ByVal Ynum As Integer, ByVal Ydenom As Integer) As Long

Declare Function GetClipBox Lib "GDI" (ByVal hdc As Integer, lpRect As RECT) As Integer
Declare Function IntersectClipRect Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
Declare Function OffsetClipRgn Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer) As Integer
Declare Function ExcludeClipRect Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Integer
Declare Function PtVisible Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer) As Integer
Declare Function CombineRgn Lib "GDI" (ByVal hDestRgn As Integer, ByVal hSrcRgn1 As Integer, ByVal hSrcRgn2 As Integer, ByVal nCombineMode As Integer) As Integer
Declare Function EqualRgn Lib "GDI" (ByVal hSrcRgn1 As Integer, ByVal hSrcRgn2 As Integer) As Integer
Declare Function OffsetRgn Lib "GDI" (ByVal hRgn As Integer, ByVal x As Integer, ByVal y As Integer) As Integer
Declare Function GetRgnBox Lib "GDI" (ByVal hRgn As Integer, lpRect As RECT) As Integer

Declare Function SetTextJustification Lib "GDI" (ByVal hdc As Integer, ByVal nBreakExtra As Integer, ByVal nBreakCount As Integer) As Integer
Declare Function GetTextExtent Lib "GDI" (ByVal hdc As Integer, ByVal lpString As String, ByVal nCount As Integer) As Long
Declare Function GetTabbedTextExtent Lib "GDI" (ByVal hdc As Integer, ByVal lpString As String, ByVal nCount As Integer, ByVal nTabPositions As Integer, lpnTabStopPositions As Integer) As Long
Declare Function SetTextCharacterExtra Lib "GDI" (ByVal hdc As Integer, ByVal nCharExtra As Integer) As Integer
Declare Function GetTextCharacterExtra Lib "GDI" (ByVal hdc As Integer) As Integer

Declare Function GetMetaFile Lib "GDI" (ByVal lpFilename As String) As Integer
Declare Function DeleteMetaFile Lib "GDI" (ByVal hMF As Integer) As Integer
Declare Function CopyMetaFile Lib "GDI" (ByVal hMF As Integer, ByVal lpFilename As String) As Integer

Declare Function PlayMetaFile Lib "GDI" (ByVal hdc As Integer, ByVal hMF As Integer) As Integer
Declare Sub PlayMetaFileRecord Lib "GDI" (ByVal hdc As Integer, lpHandletable As Integer, lpMetaRecord As METARECORD, ByVal nHandles As Integer)

Declare Function Escape Lib "GDI" (ByVal hdc As Integer, ByVal nEscape As Integer, ByVal nCount As Integer, ByVal lplnData As Any, ByVal lpOutData As Any) As Integer
Declare Function GetTextFace Lib "GDI" (ByVal hdc As Integer, ByVal nCount As Integer, ByVal lpFacename As String) As Integer

Declare Function GetTextMetrics Lib "GDI" (ByVal hdc As Integer, lpMetrics As TEXTMETRIC) As Integer

Declare Function GetDeviceCaps Lib "GDI" (ByVal hdc As Integer, ByVal nIndex As Integer) As Integer

Declare Function SetEnvironment Lib "GDI" (ByVal lpPortName As String, ByVal lpEnviron As String, ByVal nCount As Integer) As Integer
Declare Function GetEnvironment Lib "GDI" (ByVal lpPortName As String, lpEnviron As Any, ByVal nMaxCount As Integer) As Integer

Declare Function DPtoLP Lib "GDI" (ByVal hdc As Integer, lpPoints As POINTAPI, ByVal nCount As Integer) As Integer
Declare Function LPtoDP Lib "GDI" (ByVal hdc As Integer, lpPoints As POINTAPI, ByVal nCount As Integer) As Integer

Declare Function CreateMetaFile Lib "GDI" (lpString As Any) As Integer
Declare Function CloseMetaFile Lib "GDI" (ByVal hMF As Integer) As Integer
Declare Function GetMetaFileBits Lib "GDI" (ByVal hMF As Integer) As Integer
Declare Function SetMetaFileBits Lib "GDI" (ByVal hMem As Integer) As Integer

Declare Function SetDIBits Lib "GDI" (ByVal aHDC As Integer, ByVal hBitmap As Integer, ByVal nStartScan As Integer, ByVal nNumScans As Integer, ByVal lpBits As String, lpBI As BITMAPINFO, ByVal wUsage As Integer) As Integer

Declare Function GetDIBits Lib "GDI" (ByVal aHDC As Integer, ByVal hBitmap As Integer, ByVal nStartScan As Integer, ByVal nNumScans As Integer, ByVal lpBits As String, lpBI As BITMAPINFO, ByVal wUsage As Integer) As Integer

Declare Function SetDIBitsToDevice% Lib "GDI" (ByVal hdc#, ByVal x#, ByVal y#, ByVal dx#, ByVal dy#, ByVal SrcX#, ByVal SrcY#, ByVal Scan#, ByVal NumScans#, ByVal Bits As String, BitsInfo As BITMAPINFO, ByVal wUsage#)

Declare Function CreateDIBitmap Lib "GDI" (ByVal hdc As Integer, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, ByVal lpInitBits As String, lpInitInfo As BITMAPINFO, ByVal wUsage As Integer) As Integer

Declare Function CreateDIBPatternBrush Lib "GDI" (ByVal hPackedDIB As Integer, ByVal wUsage As Integer) As Integer

Declare Function StretchDIBits# Lib "GDI" (ByVal hdc#, ByVal x#, ByVal y#, ByVal dx#, ByVal dy#, ByVal SrcX#, ByVal SrcY#, ByVal wSrcWidth#, ByVal wSrcHeight#, ByVal lpBits As String, lpBitsInfo As BITMAPINFO, ByVal wUsage#, ByVal dwRop&)


Declare Function CreatePalette Lib "GDI" (lpLogPalette As LOGPALETTE) As Integer
Declare Function SelectPalette Lib "User" (ByVal hdc As Integer, ByVal hPalette As Integer, ByVal bForceBackground As Integer) As Integer
Declare Function RealizePalette Lib "User" (ByVal hdc As Integer) As Integer
Declare Function UpdateColors Lib "GDI" (ByVal hdc As Integer) As Integer
Declare Sub AnimatePalette Lib "GDI" (ByVal hPalette As Integer, ByVal wStartIndex As Integer, ByVal wNumEntries As Integer, lpPaletteColors As PALETTEENTRY)
Declare Function SetPaletteEntries Lib "GDI" (ByVal hPalette As Integer, ByVal wStartIndex As Integer, ByVal wNumEntries As Integer, lpPaletteEntries As PALETTEENTRY) As Integer
Declare Function GetPaletteEntries Lib "GDI" (ByVal hPalette As Integer, ByVal wStartIndex As Integer, ByVal wNumEntries As Integer, lpPaletteEntries As PALETTEENTRY) As Integer
Declare Function GetNearestPaletteIndex Lib "GDI" (ByVal hPalette As Integer, ByVal crColor As Long) As Integer
Declare Function ResizePalette Lib "GDI" (ByVal hPalette As Integer, ByVal nNumEntries As Integer) As Integer

Declare Function GetSystemPaletteEntries Lib "GDI" (ByVal hdc As Integer, ByVal wStartIndex As Integer, ByVal wNumEntries As Integer, lpPaletteEntries As PALETTEENTRY) As Integer
Declare Function GetSystemPaletteUse Lib "GDI" (ByVal hdc As Integer) As Integer
Declare Function SetSystemPaletteUse Lib "GDI" (ByVal hdc As Integer, ByVal wUsage As Integer) As Integer


' --------------
'  USER Section
' --------------

Declare Function wvsprintf Lib "User" (ByVal lpOutput As String, ByVal lpFormat As String, lpArglist As Integer) As Integer

' Scroll Bar Constants
Public Const SB_HORZ = 0
Public Const SB_VERT = 1
Public Const SB_CTL = 2
Public Const SB_BOTH = 3

' Scroll Bar Commands
Public Const SB_LINEUP = 0
Public Const SB_LINEDOWN = 1
Public Const SB_PAGEUP = 2
Public Const SB_PAGEDOWN = 3
Public Const SB_THUMBPOSITION = 4
Public Const SB_THUMBTRACK = 5
Public Const SB_TOP = 6
Public Const SB_BOTTOM = 7
Public Const SB_ENDSCROLL = 8

' ShowWindow() Commands
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9

' Old ShowWindow() Commands
Public Const HIDE_WINDOW = 0
Public Const SHOW_OPENWINDOW = 1
Public Const SHOW_ICONWINDOW = 2
Public Const SHOW_FULLSCREEN = 3
Public Const SHOW_OPENNOACTIVATE = 4

' Identifiers for the WM_SHOWWINDOW message
Public Const SW_PARENTCLOSING = 1
Public Const SW_OTHERZOOM = 2
Public Const SW_PARENTOPENING = 3
Public Const SW_OTHERUNZOOM = 4

' Region Flags
Public Const ERRORAPI = 0
Public Const NULLREGION = 1
Public Const SIMPLEREGION = 2
Public Const COMPLEXREGION = 3

' CombineRgn() Styles
Public Const RGN_AND = 1
Public Const RGN_OR = 2
Public Const RGN_XOR = 3
Public Const RGN_DIFF = 4
Public Const RGN_COPY = 5

' Virtual Keys, Standard Set
Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const VK_CANCEL = &H3
Public Const VK_MBUTTON = &H4   '  NOT contiguous with L & RBUTTON
Public Const VK_BACK = &H8
Public Const VK_TAB = &H9
Public Const VK_CLEAR = &HC
Public Const VK_RETURN = &HD
Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12
Public Const VK_PAUSE = &H13
Public Const VK_CAPITAL = &H14
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_SELECT = &H29
Public Const VK_PRINT = &H2A
Public Const VK_EXECUTE = &H2B
Public Const VK_SNAPSHOT = &H2C
'Public Const VK_COPY = &H2C    not used by keyboards.
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_HELP = &H2F

' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' VK_0 thru VK_9 are the same as their ASCII equivalents: '0' thru '0'

Public Const VK_NUMPAD0 = &H60
Public Const VK_NUMPAD1 = &H61
Public Const VK_NUMPAD2 = &H62
Public Const VK_NUMPAD3 = &H63
Public Const VK_NUMPAD4 = &H64
Public Const VK_NUMPAD5 = &H65
Public Const VK_NUMPAD6 = &H66
Public Const VK_NUMPAD7 = &H67
Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_F13 = &H7C
Public Const VK_F14 = &H7D
Public Const VK_F15 = &H7E
Public Const VK_F16 = &H7F

Public Const VK_NUMLOCK = &H90

' SetWindowsHook() codes
Public Const WH_MSGFILTER = (-1)
Public Const WH_JOURNALRECORD = 0
Public Const WH_JOURNALPLAYBACK = 1
Public Const WH_KEYBOARD = 2
Public Const WH_GETMESSAGE = 3
Public Const WH_CALLWNDPROC = 4
Public Const WH_CBT = 5
Public Const WH_SYSMSGFILTER = 6
Public Const WH_WINDOWMGR = 7

' Hook Codes
Public Const HC_LPLPFNNEXT = (-2)
Public Const HC_LPFNNEXT = (-1)
Public Const HC_ACTION = 0
Public Const HC_GETNEXT = 1
Public Const HC_SKIP = 2
Public Const HC_NOREM = 3
Public Const HC_NOREMOVE = 3
Public Const HC_SYSMODALON = 4
Public Const HC_SYSMODALOFF = 5

' CBT Hook Codes
Public Const HCBT_MOVESIZE = 0
Public Const HCBT_MINMAX = 1
Public Const HCBT_QS = 2

' WH_MSGFILTER Filter Proc Codes
Public Const MSGF_DIALOGBOX = 0
Public Const MSGF_MESSAGEBOX = 1
Public Const MSGF_MENU = 2
Public Const MSGF_MOVE = 3
Public Const MSGF_SIZE = 4
Public Const MSGF_SCROLLBAR = 5
Public Const MSGF_NEXTWINDOW = 6

' Window Manager Hook Codes
Public Const WC_INIT = 1
Public Const WC_SWP = 2
Public Const WC_DEFWINDOWPROC = 3
Public Const WC_MINMAX = 4
Public Const WC_MOVE = 5
Public Const WC_SIZE = 6
Public Const WC_DRAWCAPTION = 7

' Message Structure used in Journaling
Type EVENTMSG
        message As Integer
        paramL As Integer
        paramH As Integer
        time As Long
End Type

' Message structure
Type msg
        hwnd As Integer
        message As Integer
        wParam As Integer
        lParam As Long
        time As Long
        pt As POINTAPI
End Type


' Window field offsets for GetWindowLong() and GetWindowWord()
Public Const GWL_WNDPROC = (-4)
Public Const GWW_HINSTANCE = (-6)
Public Const GWW_HWNDPARENT = (-8)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)

' Class field offsets for GetClassLong() and GetClassWord()
Public Const GCL_MENUNAME = (-8)
Public Const GCW_HBRBACKGROUND = (-10)
Public Const GCW_HCURSOR = (-12)
Public Const GCW_HICON = (-14)
Public Const GCW_HMODULE = (-16)
Public Const GCW_CBWNDEXTRA = (-18)
Public Const GCW_CBCLSEXTRA = (-20)
Public Const GCL_WNDPROC = (-24)
Public Const GCW_STYLE = (-26)

' Window Messages
Public Const WM_NULL = &H0
Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
Public Const WM_ACTIVATE = &H6
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_ENABLE = &HA
Public Const WM_SETREDRAW = &HB
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_PAINT = &HF
Public Const WM_CLOSE = &H10
Public Const WM_QUERYENDSESSION = &H11
Public Const WM_QUIT = &H12
Public Const WM_QUERYOPEN = &H13
Public Const WM_ERASEBKGND = &H14
Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_ENDSESSION = &H16
Public Const WM_SHOWWINDOW = &H18
Public Const WM_CTLCOLOR = &H19
Public Const WM_WININICHANGE = &H1A
Public Const WM_DEVMODECHANGE = &H1B
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_FONTCHANGE = &H1D
Public Const WM_TIMECHANGE = &H1E
Public Const WM_CANCELMODE = &H1F
Public Const WM_SETCURSOR = &H20
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_CHILDACTIVATE = &H22
Public Const WM_QUEUESYNC = &H23
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_PAINTICON = &H26
Public Const WM_ICONERASEBKGND = &H27
Public Const WM_NEXTDLGCTL = &H28
Public Const WM_SPOOLERSTATUS = &H2A
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_DELETEITEM = &H2D
Public Const WM_VKEYTOITEM = &H2E
Public Const WM_CHARTOITEM = &H2F
Public Const WM_SETFONT = &H30
Public Const WM_GETFONT = &H31

Public Const WM_QUERYDRAGICON = &H37

Public Const WM_COMPAREITEM = &H39
Public Const WM_COMPACTING = &H41

Public Const WM_NCCREATE = &H81
Public Const WM_NCDESTROY = &H82
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCHITTEST = &H84
Public Const WM_NCPAINT = &H85
Public Const WM_NCACTIVATE = &H86
Public Const WM_GETDLGCODE = &H87
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMBUTTONDBLCLK = &HA9

Public Const WM_KEYFIRST = &H100
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Const WM_DEADCHAR = &H103
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSDEADCHAR = &H107
Public Const WM_KEYLAST = &H108

Public Const WM_INITDIALOG = &H110
Public Const WM_COMMAND = &H111
Public Const WM_SYSCOMMAND = &H112
Public Const WM_TIMER = &H113
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115
Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_MENUSELECT = &H11F
Public Const WM_MENUCHAR = &H120
Public Const WM_ENTERIDLE = &H121

Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSELAST = &H209

Public Const WM_PARENTNOTIFY = &H210
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDINEXT = &H224
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDITILE = &H226
Public Const WM_MDICASCADE = &H227
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDISETMENU = &H230

Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_UNDO = &H304
Public Const WM_RENDERFORMAT = &H305
Public Const WM_RENDERALLFORMATS = &H306
Public Const WM_DESTROYCLIPBOARD = &H307
Public Const WM_DRAWCLIPBOARD = &H308
Public Const WM_PAINTCLIPBOARD = &H309
Public Const WM_VSCROLLCLIPBOARD = &H30A
Public Const WM_SIZECLIPBOARD = &H30B
Public Const WM_ASKCBFORMATNAME = &H30C
Public Const WM_CHANGECBCHAIN = &H30D
Public Const WM_HSCROLLCLIPBOARD = &H30E
Public Const WM_QUERYNEWPALETTE = &H30F
Public Const WM_PALETTEISCHANGING = &H310
Public Const WM_PALETTECHANGED = &H311

' NOTE: All Message Numbers below 0x0400 are RESERVED.

' Private Window Messages Start Here:
Public Const WM_USER = &H400

' WM_SYNCTASK Commands
Public Const ST_BEGINSWP = 0
Public Const ST_ENDSWP = 1

' WinWhere() Area Codes
Public Const HTERROR = (-2)
Public Const HTTRANSPARENT = (-1)
Public Const HTNOWHERE = 0
Public Const HTCLIENT = 1
Public Const HTCAPTION = 2
Public Const HTSYSMENU = 3
Public Const HTGROWBOX = 4
Public Const HTSIZE = HTGROWBOX
Public Const HTMENU = 5
Public Const HTHSCROLL = 6
Public Const HTVSCROLL = 7
Public Const HTREDUCE = 8
Public Const HTZOOM = 9
Public Const HTLEFT = 10
Public Const HTRIGHT = 11
Public Const HTTOP = 12
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17
Public Const HTSIZEFIRST = HTLEFT
Public Const HTSIZELAST = HTBOTTOMRIGHT

' WM_MOUSEACTIVATE Return Codes
Public Const MA_ACTIVATE = 1
Public Const MA_ACTIVATEANDEAT = 2
Public Const MA_NOACTIVATE = 3

Declare Function RegisterWindowMessage Lib "User" (ByVal lpString As String) As Integer

' Size Message Commands
Public Const SIZENORMAL = 0
Public Const SIZEICONIC = 1
Public Const SIZEFULLSCREEN = 2
Public Const SIZEZOOMSHOW = 3
Public Const SIZEZOOMHIDE = 4

' Key State Masks for Mouse Messages
Public Const MK_LBUTTON = &H1
Public Const MK_RBUTTON = &H2
Public Const MK_SHIFT = &H4
Public Const MK_CONTROL = &H8
Public Const MK_MBUTTON = &H10

' Window Styles
Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_DISABLED = &H8000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_CAPTION = &HC00000      '  WS_BORDER Or WS_DLGFRAME
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000

Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000

Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_SIZEBOX = WS_THICKFRAME

' Common Window Styles
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_TILEDWINDOW = (WS_OVERLAPPEDWINDOW)

' Extended Window Styles
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&

' Class styles
Public Const CS_VREDRAW = &H1
Public Const CS_HREDRAW = &H2
Public Const CS_KEYCVTWINDOW = &H4
Public Const CS_DBLCLKS = &H8
Public Const CS_OWNDC = &H20
Public Const CS_CLASSDC = &H40
Public Const CS_PARENTDC = &H80
Public Const CS_NOKEYCVT = &H100
Public Const CS_NOCLOSE = &H200
Public Const CS_SAVEBITS = &H800
Public Const CS_BYTEALIGNCLIENT = &H1000
Public Const CS_BYTEALIGNWINDOW = &H2000
Public Const CS_GLOBALCLASS = &H4000    '  Global window class

' Predefined Clipboard Formats
Public Const CF_TEXT = 1
Public Const CF_BITMAP = 2
Public Const CF_METAFILEPICT = 3
Public Const CF_SYLK = 4
Public Const CF_DIF = 5
Public Const CF_TIFF = 6
Public Const CF_OEMTEXT = 7
Public Const CF_DIB = 8
Public Const CF_PALETTE = 9

Public Const CF_OWNERDISPLAY = &H80
Public Const CF_DSPTEXT = &H81
Public Const CF_DSPBITMAP = &H82
Public Const CF_DSPMETAFILEPICT = &H83

' "Private" formats don't get GlobalFree()'d
Public Const CF_PRIVATEFIRST = &H200
Public Const CF_PRIVATELAST = &H2FF

' "GDIOBJ" formats do get DeleteObject()'d
Public Const CF_GDIOBJFIRST = &H300
Public Const CF_GDIOBJLAST = &H3FF

Type PAINTSTRUCT
        hdc As Integer
        fErase As Integer
        rcPaint As RECT
        fRestore As Integer
        fIncUpdate As Integer
        rgbReserved As String * 16
End Type

Type CREATESTRUCT
        lpCreateParams As Long
        hInstance As Integer
        hmenu As Integer
        hWndParent As Integer
        cy As Integer
        cx As Integer
        y As Integer
        x As Integer
        style As Long
        lpszName As Long
        lpszClass As Long
        ExStyle As Long
End Type

' Owner draw control types
Public Const ODT_MENU = 1
Public Const ODT_LISTBOX = 2
Public Const ODT_COMBOBOX = 3
Public Const ODT_BUTTON = 4

' Owner draw actions
Public Const ODA_DRAWENTIRE = &H1
Public Const ODA_SELECT = &H2
Public Const ODA_FOCUS = &H4

' Owner draw state
Public Const ODS_SELECTED = &H1
Public Const ODS_GRAYED = &H2
Public Const ODS_DISABLED = &H4
Public Const ODS_CHECKED = &H8
Public Const ODS_FOCUS = &H10

' MEASUREITEMSTRUCT for ownerdraw
Type MEASUREITEMSTRUCT
        CtlType As Integer
        CtlID As Integer
        itemID As Integer
        itemWidth As Integer
        itemHeight As Integer
        itemData As Long
End Type

' DRAWITEMSTRUCT for ownerdraw
Type DRAWITEMSTRUCT
        CtlType As Integer
        CtlID As Integer
        itemID As Integer
        itemAction As Integer
        itemState As Integer
        hWndItem As Integer
        hdc As Integer
        rcItem As RECT
        itemData As Long
End Type

' DELETEITEMSTRUCT for ownerdraw
Type DELETEITEMSTRUCT
        CtlType As Integer
        CtlID As Integer
        itemID As Integer
        hWndItem As Integer
        itemData As Long
End Type

' COMPAREITEMSTRUCT for ownerdraw sorting
Type COMPAREITEMSTRUCT
        CtlType As Integer
        CtlID As Integer
        hWndItem As Integer
        itemID1 As Integer
        itemData1 As Long
        itemID2 As Integer
        itemData2 As Long
End Type

' Message Function Templates
Declare Function GetMessage Lib "User" (lpMsg As msg, ByVal hwnd As Integer, ByVal wMsgFilterMin As Integer, ByVal wMsgFilterMax As Integer) As Integer
Declare Function TranslateMessage Lib "User" (lpMsg As msg) As Integer
Declare Function DispatchMessage Lib "User" (lpMsg As msg) As Long
Declare Function PeekMessage Lib "User" (lpMsg As msg, ByVal hwnd As Integer, ByVal wMsgFilterMin As Integer, ByVal wMsgFilterMax As Integer, ByVal wRemoveMsg As Integer) As Integer

' PeekMessage() Options
Public Const PM_NOREMOVE = &H0
Public Const PM_REMOVE = &H1
Public Const PM_NOYIELD = &H2

Declare Function lstrcmp Lib "User" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Integer
Declare Function lstrcmpi Lib "User" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Integer
Declare Function lstrcpy Lib "Kernel" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare Function lstrcat Lib "Kernel" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare Function lstrlen Lib "Kernel" (ByVal lpString As Any) As Integer

Declare Function lopen Lib "Kernel" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Integer) As Integer
Declare Function lclose Lib "Kernel" Alias "_lclose" (ByVal hFile As Integer) As Integer
Declare Function lcreat Lib "Kernel" Alias "_lcreat" (ByVal lpPathName As String, ByVal iAttribute As Integer) As Integer
Declare Function llseek Lib "Kernel" Alias "_llseek" (ByVal hFile As Integer, ByVal lOffset As Long, ByVal iOrigin As Integer) As Long
Declare Function lread Lib "Kernel" Alias "_lread" (ByVal hFile As Integer, ByVal lpBuffer As String, ByVal wBytes As Integer) As Integer
Declare Function lwrite Lib "Kernel" Alias "_lwrite" (ByVal hFile As Integer, ByVal lpBuffer As String, ByVal wBytes As Integer) As Integer

Public Const READAPI = 0        '  Flags for _lopen
Public Const WRITEAPI = 1
Public Const READ_WRITE = 2

Declare Function ExitWindows Lib "User" (ByVal dwReturnCode As Long, ByVal wReserved As Integer) As Integer

Declare Function SwapMouseButton Lib "User" (ByVal bSwap As Integer) As Integer
Declare Function GetMessagePos Lib "User" () As Long
Declare Function GetMessageTime Lib "User" () As Long

Declare Function GetSysModalWindow Lib "User" () As Integer
Declare Function SetSysModalWindow Lib "User" (ByVal hwnd As Integer) As Integer

Declare Function SendMessage Lib "User" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Function PostMessage Lib "User" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Integer
Declare Function PostAppMessage Lib "User" (ByVal htask As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Integer
Declare Sub ReplyMessage Lib "User" (ByVal lReply As Long)
Declare Sub WaitMessage Lib "User" ()
Declare Function DefWindowProc Lib "User" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Sub PostQuitMessage Lib "User" (ByVal nExitCode As Integer)
Declare Function InSendMessage Lib "User" () As Integer

' Special HWND value for use with PostMessage and SendMessage
Public Const HWND_BROADCAST = &HFFFF

Declare Function GetDoubleClickTime Lib "User" () As Integer
Declare Sub SetDoubleClickTime Lib "User" (ByVal wCount As Integer)

Declare Function UnregisterClass Lib "User" (ByVal lpClassName As String, ByVal hInstance As Integer) As Integer

Declare Function SetMessageQueue Lib "User" (ByVal cMsg As Integer) As Integer

Public Const CW_USEDEFAULT = &H8000
Public Const HWND_DESKTOP = 0

Declare Function CreateWindow% Lib "User" (ByVal lpClassName$, ByVal lpWindowName$, ByVal dwStyle&, ByVal x%, ByVal y%, ByVal nWidth%, ByVal nHeight%, ByVal hWndParent%, ByVal hmenu%, ByVal hInstance%, ByVal lpParam$)
Declare Function CreateWindowEx% Lib "User" (ByVal dwExStyle&, ByVal lpClassName$, ByVal lpWindowName$, ByVal dwStyle&, ByVal x%, ByVal y%, ByVal nWidth%, ByVal nHeight%, ByVal hWndParent%, ByVal hmenu%, ByVal hInstance%, ByVal lpParam$)
Declare Function IsWindow Lib "User" (ByVal hwnd As Integer) As Integer
Declare Function IsChild Lib "User" (ByVal hWndParent As Integer, ByVal hwnd As Integer) As Integer
Declare Function DestroyWindow Lib "User" (ByVal hwnd As Integer) As Integer

Declare Function ShowWindow Lib "User" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer
Declare Function FlashWindow Lib "User" (ByVal hwnd As Integer, ByVal bInvert As Integer) As Integer
Declare Sub ShowOwnedPopups Lib "User" (ByVal hwnd As Integer, ByVal fShow As Integer)

Declare Function OpenIcon Lib "User" (ByVal hwnd As Integer) As Integer
Declare Sub CloseWindow Lib "User" (ByVal hwnd As Integer)
Declare Sub MoveWindow Lib "User" (ByVal hwnd As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Integer)
Declare Sub SetWindowPos Lib "User" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)

Declare Function BeginDeferWindowPos Lib "User" (ByVal nNumWindows As Integer) As Integer
Declare Function DeferWindowPos Lib "User" (ByVal hWinPosInfo As Integer, ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x, ByVal y, ByVal cx, ByVal cy, ByVal wFlags As Integer) As Integer
Declare Sub EndDeferWindowPos Lib "User" (ByVal hWinPosInfo As Integer)

Declare Function IsWindowVisible Lib "User" (ByVal hwnd As Integer) As Integer
Declare Function IsIconic Lib "User" (ByVal hwnd As Integer) As Integer
Declare Function AnyPopup Lib "User" () As Integer
Declare Sub BringWindowToTop Lib "User" (ByVal hwnd As Integer)
Declare Function IsZoomed Lib "User" (ByVal hwnd As Integer) As Integer

' SetWindowPos Flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_DRAWFRAME = &H20
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOREPOSITION = &H200

' SetWindowPos() hWndInsertAfter values
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Sub EndDialog Lib "User" (ByVal hDlg As Integer, ByVal nResult As Integer)
Declare Function GetDlgItem Lib "User" (ByVal hDlg As Integer, ByVal nIDDlgItem As Integer) As Integer
Declare Sub SetDlgItemInt Lib "User" (ByVal hDlg As Integer, ByVal nIDDlgItem As Integer, ByVal wValue As Integer, ByVal bSigned As Integer)
Declare Function GetDlgItemInt Lib "User" (ByVal hDlg As Integer, ByVal nIDDlgItem As Integer, lpTranslated As Integer, ByVal bSigned As Integer) As Integer
Declare Sub SetDlgItemText Lib "User" (ByVal hDlg As Integer, ByVal nIDDlgItem As Integer, ByVal lpString As String)
Declare Function GetDlgItemText Lib "User" (ByVal hDlg As Integer, ByVal nIDDlgItem As Integer, ByVal lpString As String, ByVal nMaxCount As Integer) As Integer
Declare Sub CheckDlgButton Lib "User" (ByVal hDlg As Integer, ByVal nIDButton As Integer, ByVal wCheck As Integer)
Declare Sub CheckRadioButton Lib "User" (ByVal hDlg As Integer, ByVal nIDFirstButton As Integer, ByVal nIDLastButton As Integer, ByVal nIDCheckButton As Integer)
Declare Function IsDlgButtonChecked Lib "User" (ByVal hDlg As Integer, ByVal nIDButton As Integer) As Integer
Declare Function SendDlgItemMessage Lib "User" (ByVal hDlg As Integer, ByVal nIDDlgItem As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Function GetNextDlgGroupItem Lib "User" (ByVal hDlg As Integer, ByVal hCtl As Integer, ByVal bPrevious As Integer) As Integer
Declare Function GetNextDlgTabItem Lib "User" (ByVal hDlg As Integer, ByVal hCtl As Integer, ByVal bPrevious As Integer) As Integer
Declare Function GetDlgCtrlID Lib "User" (ByVal hwnd As Integer) As Integer
Declare Function GetDialogBaseUnits Lib "User" () As Long
Declare Function DefDlgProc Lib "User" (ByVal hDlg As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long

Public Const DLGWINDOWEXTRA = 30        '  Window extra bytes needed for private dialog classes

Declare Function CallMsgFilter Lib "User" (lpMsg As msg, ByVal ncode As Integer) As Integer

' Clipboard Manager Functions
Declare Function OpenClipboard Lib "User" (ByVal hwnd As Integer) As Integer
Declare Function CloseClipboard Lib "User" () As Integer
Declare Function GetClipboardOwner Lib "User" () As Integer
Declare Function SetClipboardViewer Lib "User" (ByVal hwnd As Integer) As Integer
Declare Function GetClipboardViewer Lib "User" () As Integer
Declare Function ChangeClipboardChain Lib "User" (ByVal hwnd As Integer, ByVal hWndNext As Integer) As Integer
Declare Function SetClipboardData Lib "User" (ByVal wFormat As Integer, ByVal hMem As Integer) As Integer
Declare Function GetClipboardData Lib "User" (ByVal wFormat As Integer) As Integer
Declare Function RegisterClipboardFormat Lib "User" (ByVal lpString As String) As Integer
Declare Function CountClipboardFormats Lib "User" () As Integer
Declare Function EnumClipboardFormats Lib "User" (ByVal wFormat As Integer) As Integer
Declare Function GetClipboardFormatName Lib "User" (ByVal wFormat As Integer, ByVal lpString As String, ByVal nMaxCount As Integer) As Integer
Declare Function EmptyClipboard Lib "User" () As Integer
Declare Function IsClipboardFormatAvailable Lib "User" (ByVal wFormat As Integer) As Integer
Declare Function GetPriorityClipboardFormat Lib "User" (lpPriorityList As Integer, ByVal nCount As Integer) As Integer

Declare Function SetFocusAPI Lib "User" Alias "SetFocus" (ByVal hwnd As Integer) As Integer
Declare Function GetFocus Lib "User" () As Integer
Declare Function GetActiveWindow Lib "User" () As Integer
Declare Function GetKeyState Lib "User" (ByVal nVirtKey As Integer) As Integer
Declare Function GetAsyncKeyState Lib "User" (ByVal vKey As Integer) As Integer
Declare Sub GetKeyboardState Lib "User" (lpKeyState As Any)
Declare Sub SetKeyboardState Lib "User" (lpKeyState As Any)
Declare Function EnableHardwareInput Lib "User" (ByVal bEnableInput As Integer) As Integer
Declare Function GetInputState Lib "User" () As Integer
Declare Function GetCapture Lib "User" () As Integer
Declare Function SetCapture Lib "User" (ByVal hwnd As Integer) As Integer
Declare Sub ReleaseCapture Lib "User" ()

' Windows Functions
Declare Function KillTimer Lib "User" (ByVal hwnd As Integer, ByVal nIDEvent As Integer) As Integer

Declare Function EnableWindow Lib "User" (ByVal hwnd As Integer, ByVal aBOOL As Integer) As Integer
Declare Function IsWindowEnabled Lib "User" (ByVal hwnd As Integer) As Integer

Declare Function LoadAccelerators Lib "User" (ByVal hInstance As Integer, ByVal lpTableName As String) As Integer

Declare Function TranslateAccelerator Lib "User" (ByVal hwnd As Integer, ByVal hAccTable As Integer, lpMsg As msg) As Integer


' GetSystemMetrics() codes
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Const SM_CXVSCROLL = 2
Public Const SM_CYHSCROLL = 3
Public Const SM_CYCAPTION = 4
Public Const SM_CXBORDER = 5
Public Const SM_CYBORDER = 6
Public Const SM_CXDLGFRAME = 7
Public Const SM_CYDLGFRAME = 8
Public Const SM_CYVTHUMB = 9
Public Const SM_CXHTHUMB = 10
Public Const SM_CXICON = 11
Public Const SM_CYICON = 12
Public Const SM_CXCURSOR = 13
Public Const SM_CYCURSOR = 14
Public Const SM_CYMENU = 15
Public Const SM_CXFULLSCREEN = 16
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYKANJIWINDOW = 18
Public Const SM_MOUSEPRESENT = 19
Public Const SM_CYVSCROLL = 20
Public Const SM_CXHSCROLL = 21
Public Const SM_DEBUG = 22
Public Const SM_SWAPBUTTON = 23
Public Const SM_RESERVED1 = 24
Public Const SM_RESERVED2 = 25
Public Const SM_RESERVED3 = 26
Public Const SM_RESERVED4 = 27
Public Const SM_CXMIN = 28
Public Const SM_CYMIN = 29
Public Const SM_CXSIZE = 30
Public Const SM_CYSIZE = 31
Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33
Public Const SM_CXMINTRACK = 34
Public Const SM_CYMINTRACK = 35
Public Const SM_CMETRICS = 36

Declare Function GetSystemMetrics Lib "User" (ByVal nIndex As Integer) As Integer

Declare Function LoadMenu Lib "User" (ByVal hInstance As Integer, ByVal lpString As String) As Integer
Declare Function LoadMenuIndirect Lib "User" (lpMenuTemplate As Any) As Integer
Declare Function GetMenu Lib "User" (ByVal hwnd As Integer) As Integer
Declare Function SetMenu Lib "User" (ByVal hwnd As Integer, ByVal hmenu As Integer) As Integer
Declare Function ChangeMenu Lib "User" (ByVal hmenu As Integer, ByVal wID As Integer, ByVal lpszNew As String, ByVal wIDNew As Integer, ByVal wChange As Integer) As Integer
Declare Function HiliteMenuItem Lib "User" (ByVal hwnd As Integer, ByVal hmenu As Integer, ByVal wIDHiliteItem As Integer, ByVal wHilite As Integer) As Integer
Declare Function GetMenuString Lib "User" (ByVal hmenu As Integer, ByVal wIDItem As Integer, ByVal lpString As String, ByVal nMaxCount As Integer, ByVal wFlag As Integer) As Integer
Declare Function GetMenuState Lib "User" (ByVal hmenu As Integer, ByVal wID As Integer, ByVal wFlags As Integer) As Integer
Declare Sub DrawMenuBar Lib "User" (ByVal hwnd As Integer)
Declare Function GetSystemMenu Lib "User" (ByVal hwnd As Integer, ByVal bRevert As Integer) As Integer
Declare Function CreateMenu Lib "User" () As Integer
Declare Function CreatePopupMenu Lib "User" () As Integer
Declare Function DestroyMenu Lib "User" (ByVal hmenu As Integer) As Integer
Declare Function CheckMenuItem Lib "User" (ByVal hmenu As Integer, ByVal wIDCheckItem As Integer, ByVal wCheck As Integer) As Integer
Declare Function EnableMenuItem Lib "User" (ByVal hmenu As Integer, ByVal wIDEnableItem As Integer, ByVal wEnable As Integer) As Integer
Declare Function GetSubMenu Lib "User" (ByVal hmenu As Integer, ByVal nPos As Integer) As Integer
Declare Function GetMenuItemID Lib "User" (ByVal hmenu As Integer, ByVal nPos As Integer) As Integer
Declare Function GetMenuItemCount Lib "User" (ByVal hmenu As Integer) As Integer

Declare Function InsertMenu Lib "User" (ByVal hmenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
Declare Function AppendMenu Lib "User" (ByVal hmenu As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
Declare Function ModifyMenu Lib "User" (ByVal hmenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpString As Any) As Integer
Declare Function RemoveMenu Lib "User" (ByVal hmenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
Declare Function DeleteMenu Lib "User" (ByVal hmenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
Declare Function SetMenuItemBitmaps Lib "User" (ByVal hmenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal hBitmapUnchecked As Integer, ByVal hBitmapChecked As Integer) As Integer
Declare Function GetMenuCheckMarkDimensions Lib "User" () As Long
Declare Function TrackPopupMenu Lib "User" (ByVal hmenu As Integer, ByVal wFlags As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nReserved As Integer, ByVal hwnd As Integer, lpReserved As Any) As Integer

Declare Sub UpdateWindow Lib "GDI" (ByVal hwnd As Integer)
Declare Function SetActiveWindow Lib "User" (ByVal hwnd As Integer) As Integer

Declare Function BeginPaint Lib "GDI" (ByVal hwnd As Integer, lpPaint As PAINTSTRUCT) As Integer
Declare Sub EndPaint Lib "GDI" (ByVal hwnd As Integer, lpPaint As PAINTSTRUCT)
Declare Function GetUpdateRect Lib "GDI" (ByVal hwnd As Integer, lpRect As RECT, ByVal bErase As Integer) As Integer
Declare Function GetUpdateRgn Lib "GDI" (ByVal hwnd As Integer, ByVal hRgn As Integer, ByVal fErase As Integer) As Integer

Declare Function ExcludeUpdateRgn Lib "GDI" (ByVal hdc As Integer, ByVal hwnd As Integer) As Integer

Declare Sub InvalidateRect Lib "User" (ByVal hwnd As Integer, lpRect As RECT, ByVal bErase As Integer)
Declare Sub ValidateRect Lib "GDI" (ByVal hwnd As Integer, lpRect As RECT)

Declare Sub InvalidateRgn Lib "GDI" (ByVal hwnd As Integer, ByVal hRgn As Integer, ByVal bErase As Integer)
Declare Sub ValidateRgn Lib "GDI" (ByVal hwnd As Integer, ByVal hRgn As Integer)

Declare Sub ScrollWindow Lib "User" (ByVal hwnd As Integer, ByVal XAmount As Integer, ByVal YAmount As Integer, lpRect As RECT, lpClipRect As RECT)
Declare Function ScrollDC Lib "User" (ByVal hdc As Integer, ByVal dx As Integer, ByVal dy As Integer, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Integer, lprcUpdate As RECT) As Integer

Declare Function SetScrollPos Lib "User" (ByVal hwnd As Integer, ByVal nBar As Integer, ByVal nPos As Integer, ByVal bRedraw As Integer) As Integer
Declare Function GetScrollPos Lib "User" (ByVal hwnd As Integer, ByVal nBar As Integer) As Integer
Declare Sub SetScrollRange Lib "User" (ByVal hwnd As Integer, ByVal nBar As Integer, ByVal nMinPos As Integer, ByVal nMaxPos As Integer, ByVal bRedraw As Integer)
Declare Sub GetScrollRange Lib "User" (ByVal hwnd As Integer, ByVal nBar As Integer, lpMinPos As Integer, lpMaxPos As Integer)
Declare Sub ShowScrollBar Lib "User" (ByVal hwnd As Integer, ByVal wBar As Integer, ByVal bShow As Integer)

Declare Function SetProp Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal hData As Integer) As Integer
Declare Function GetProp Lib "User" (ByVal hwnd As Integer, ByVal lpString As Any) As Integer
Declare Function RemoveProp Lib "User" (ByVal hwnd As Integer, ByVal lpString As String) As Integer
Declare Sub SetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String)
Declare Function GetWindowText Lib "User" (ByVal hwnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
Declare Function GetWindowTextLength Lib "User" (ByVal hwnd As Integer) As Integer

Declare Sub GetClientRect Lib "GDI" (ByVal hwnd As Integer, lpRect As RECT)
Declare Sub GetWindowRect Lib "GDI" (ByVal hwnd As Integer, lpRect As RECT)
Declare Sub AdjustWindowRect Lib "GDI" (lpRect As RECT, ByVal dwStyle As Long, ByVal bMenu As Integer)
Declare Sub AdjustWindowRectEx Lib "GDI" (lpRect As RECT, ByVal dsStyle As Long, ByVal bMenu As Integer, ByVal dwEsStyle As Long)


' MessageBox() Flags
Public Const MB_OK = &H0
Public Const MB_OKCANCEL = &H1
Public Const MB_ABORTRETRYIGNORE = &H2
Public Const MB_YESNOCANCEL = &H3
Public Const MB_YESNO = &H4
Public Const MB_RETRYCANCEL = &H5

Public Const MB_ICONHAND = &H10
Public Const MB_ICONQUESTION = &H20
Public Const MB_ICONEXCLAMATION = &H30
Public Const MB_ICONASTERISK = &H40

Public Const MB_ICONINFORMATION = MB_ICONASTERISK
Public Const MB_ICONSTOP = MB_ICONHAND

Public Const MB_DEFBUTTON1 = &H0
Public Const MB_DEFBUTTON2 = &H100
Public Const MB_DEFBUTTON3 = &H200

Public Const MB_APPLMODAL = &H0
Public Const MB_SYSTEMMODAL = &H1000
Public Const MB_TASKMODAL = &H2000

Public Const MB_NOFOCUS = &H8000

Public Const MB_TYPEMASK = &HF
Public Const MB_ICONMASK = &HF0
Public Const MB_DEFMASK = &HF00
Public Const MB_MODEMASK = &H3000
Public Const MB_MISCMASK = &HC000

Declare Function MessageBox Lib "User" (ByVal hwnd As Integer, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Integer) As Integer
Declare Sub MessageBeep Lib "User" (ByVal wType As Integer)

Declare Function ShowCursor Lib "User" (ByVal bShow As Integer) As Integer
Declare Sub SetCursorPos Lib "User" (ByVal x As Integer, ByVal y As Integer)
Declare Function SetCursor Lib "User" (ByVal hCursor As Integer) As Integer
Declare Sub GetCursorPos Lib "User" (lpPoint As POINTAPI)
Declare Sub ClipCursor Lib "User" (lpRect As Any)

Declare Sub CreateCaret Lib "User" (ByVal hwnd As Integer, ByVal hBitmap As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer)
Declare Function GetCaretBlinkTime Lib "User" () As Integer
Declare Sub SetCaretBlinkTime Lib "User" (ByVal wMSeconds As Integer)
Declare Sub DestroyCaret Lib "User" ()
Declare Sub HideCaret Lib "User" (ByVal hwnd As Integer)
Declare Sub ShowCaret Lib "User" (ByVal hwnd As Integer)
Declare Sub SetCaretPos Lib "User" (ByVal x As Integer, ByVal y As Integer)
Declare Sub GetCaretPos Lib "User" (lpPoint As POINTAPI)

Declare Sub ClientToScreen Lib "User" (ByVal hwnd As Integer, lpPoint As POINTAPI)
Declare Sub ScreenToClient Lib "User" (ByVal hwnd As Integer, lpPoint As POINTAPI)

Declare Function WindowFromPoint Lib "User" (ByVal yPoint As Integer, ByVal xPoint As Integer) As Integer
Declare Function ChildWindowFromPoint Lib "User" (ByVal hwnd As Integer, ByVal ptScreen As Any) As Integer


' Color Types
Public Const CTLCOLOR_MSGBOX = 0
Public Const CTLCOLOR_EDIT = 1
Public Const CTLCOLOR_LISTBOX = 2
Public Const CTLCOLOR_BTN = 3
Public Const CTLCOLOR_DLG = 4
Public Const CTLCOLOR_SCROLLBAR = 5
Public Const CTLCOLOR_STATIC = 6
Public Const CTLCOLOR_MAX = 8   '  three bits max

Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_MENU = 4
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_ENDCOLORS = COLOR_BTNTEXT

Declare Function GetSysColor Lib "User" (ByVal nIndex As Integer) As Long
Declare Sub SetSysColors Lib "User" (ByVal nChanges As Integer, lpSysColor As Integer, lpColorValues As Long)

Declare Function FillRgn Lib "GDI" (ByVal hdc As Integer, ByVal hRgn As Integer, ByVal hBrush As Integer) As Integer
Declare Function FrameRgn Lib "GDI" (ByVal hdc As Integer, ByVal hRgn As Integer, ByVal hBrush As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
Declare Function InvertRgn Lib "GDI" (ByVal hdc As Integer, ByVal hRgn As Integer) As Integer
Declare Function PaintRgn Lib "GDI" (ByVal hdc As Integer, ByVal hRgn As Integer) As Integer
Declare Function PtInRegion Lib "GDI" (ByVal hRgn As Integer, ByVal x As Integer, ByVal y As Integer) As Integer

Declare Sub DrawFocusRect Lib "GDI" (ByVal hdc As Integer, lpRect As RECT)
Declare Function FillRect Lib "GDI" (ByVal hdc As Integer, lpRect As RECT, ByVal hBrush As Integer) As Integer
Declare Function FrameRect Lib "GDI" (ByVal hdc As Integer, lpRect As RECT, ByVal hBrush As Integer) As Integer
Declare Sub InvertRect Lib "GDI" (ByVal hdc As Integer, lpRect As RECT)
Declare Sub SetRect Lib "GDI" (lpRect As RECT, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Sub SetRectEmpty Lib "GDI" (lpRect As RECT)
Declare Function CopyRect Lib "GDI" (lpDestRect As RECT, lpSourceRect As RECT) As Integer
Declare Sub InflateRect Lib "GDI" (lpRect As RECT, ByVal x As Integer, ByVal y As Integer)
Declare Function IntersectRect Lib "GDI" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Integer
Declare Function UnionRect Lib "GDI" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Integer
Declare Sub OffsetRect Lib "GDI" (lpRect As RECT, ByVal x As Integer, ByVal y As Integer)
Declare Function IsRectEmpty Lib "GDI" (lpRect As RECT) As Integer
Declare Function EqualRect Lib "GDI" (lpRect1 As RECT, lpRect2 As RECT) As Integer
Declare Function PtInRect Lib "GDI" (lpRect As RECT, ptRect As Any) As Integer
Declare Function RectVisible Lib "GDI" (ByVal hdc As Integer, lpRect As RECT) As Integer
Declare Function RectInRegion Lib "GDI" (ByVal hRgn As Integer, lpRect As RECT) As Integer

Declare Function GetCurrentTime Lib "User" () As Long
Declare Function GetTickCount Lib "User" () As Long

Declare Function GetWindowWord Lib "User" (ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer
Declare Function SetWindowWord Lib "User" (ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal wNewWord As Integer) As Integer
Declare Function GetWindowLong Lib "User" (ByVal hwnd As Integer, ByVal nIndex As Integer) As Long
Declare Function SetWindowLong Lib "User" (ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Long) As Long
Declare Function GetClassWord Lib "User" (ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer
Declare Function SetClassWord Lib "User" (ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal wNewWord As Integer) As Integer
Declare Function GetClassLong Lib "User" (ByVal hwnd As Integer, ByVal nIndex As Integer) As Long
Declare Function SetClassLong Lib "User" (ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Long) As Long
Declare Function GetDesktopHwnd Lib "User" () As Integer
Declare Function GetDesktopWindow Lib "User" () As Integer


Declare Function GetParent Lib "User" (ByVal hwnd As Integer) As Integer
Declare Function SetParent Lib "User" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer
Declare Function FindWindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Function GetClassName Lib "User" (ByVal hwnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
Declare Function GetTopWindow Lib "User" (ByVal hwnd As Integer) As Integer
Declare Function GetNextWindow Lib "User" (ByVal hwnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function GetWindowTask Lib "User" (ByVal hwnd As Integer) As Integer
Declare Function GetLastActivePopup Lib "User" (ByVal hWndOwnder As Integer) As Integer

' GetWindow() Constants
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5

Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer

' Menu flags for Add/Check/EnableMenuItem()
Public Const MF_INSERT = &H0
Public Const MF_CHANGE = &H80
Public Const MF_APPEND = &H100
Public Const MF_DELETE = &H200
Public Const MF_REMOVE = &H1000

Public Const MF_BYCOMMAND = &H0
Public Const MF_BYPOSITION = &H400

Public Const MF_SEPARATOR = &H800

Public Const MF_ENABLED = &H0
Public Const MF_GRAYED = &H1
Public Const MF_DISABLED = &H2

Public Const MF_UNCHECKED = &H0
Public Const MF_CHECKED = &H8
Public Const MF_USECHECKBITMAPS = &H200

Public Const MF_STRING = &H0
Public Const MF_BITMAP = &H4
Public Const MF_OWNERDRAW = &H100

Public Const MF_POPUP = &H10
Public Const MF_MENUBARBREAK = &H20
Public Const MF_MENUBREAK = &H40

Public Const MF_UNHILITE = &H0
Public Const MF_HILITE = &H80

Public Const MF_SYSMENU = &H2000
Public Const MF_HELP = &H4000
Public Const MF_MOUSESELECT = &H8000

' Menu item resource format
Type MENUITEMTEMPLATEHEADER
        versionNumber As Integer
        offset As Integer
End Type

Type MENUITEMTEMPLATE
        mtOption As Integer
        mtID As Integer
        mtString As Long
End Type

Public Const MF_END = &H80

' System Menu Command Values
Public Const SC_SIZE = &HF000
Public Const SC_MOVE = &HF010
Public Const SC_MINIMIZE = &HF020
Public Const SC_MAXIMIZE = &HF030
Public Const SC_NEXTWINDOW = &HF040
Public Const SC_PREVWINDOW = &HF050
Public Const SC_CLOSE = &HF060
Public Const SC_VSCROLL = &HF070
Public Const SC_HSCROLL = &HF080
Public Const SC_MOUSEMENU = &HF090
Public Const SC_KEYMENU = &HF100
Public Const SC_ARRANGE = &HF110
Public Const SC_RESTORE = &HF120
Public Const SC_TASKLIST = &HF130

Public Const SC_ICON = SC_MINIMIZE
Public Const SC_ZOOM = SC_MAXIMIZE

' Resource Loading Routines
Declare Function LoadBitmap Lib "User" (ByVal hInstance As Integer, ByVal lpBitmapName As Any) As Integer
Declare Function LoadCursor Lib "User" (ByVal hInstance As Integer, ByVal lpCursorName As Any) As Integer
Declare Function CreateCursor Lib "User" (ByVal hInstance%, ByVal nXhotspot%, ByVal nYhotspot%, ByVal nWidth%, ByVal nHeight%, ByVal lpANDbitPlane As Any, ByVal lpXORbitPlane As Any) As Integer
Declare Function DestroyCursor Lib "User" (ByVal hCursor As Integer) As Integer

' Standard Cursor IDs
Public Const IDC_ARROW = 32512&
Public Const IDC_IBEAM = 32513&
Public Const IDC_WAIT = 32514&
Public Const IDC_CROSS = 32515&
Public Const IDC_UPARROW = 32516&
Public Const IDC_SIZE = 32640&
Public Const IDC_ICON = 32641&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_SIZENS = 32645&

Declare Function LoadIcon Lib "User" (ByVal hInstance As Integer, ByVal lpIconName As Any) As Integer
Declare Function CreateIcon Lib "User" (ByVal hInstance%, ByVal nWidth%, ByVal nHeight%, ByVal nPlanes%, ByVal nBitsPixel%, ByVal lpANDbits As Any, ByVal lpXORbits As Any) As Integer
Declare Function DestroyIcon Lib "User" (ByVal hicon As Integer) As Integer


Public Const ORD_LANGDRIVER = 1 '  The ordinal number for the entry point of
                                '  language drivers.

' Standard Icon IDs
Public Const IDI_APPLICATION = 32512&
Public Const IDI_HAND = 32513&
Public Const IDI_QUESTION = 32514&
Public Const IDI_EXCLAMATION = 32515&
Public Const IDI_ASTERISK = 32516&

Declare Function LoadString Lib "User" (ByVal hInstance As Integer, ByVal wID As Integer, ByVal lpBuffer As Any, ByVal nBufferMax As Integer) As Integer

Declare Function AddFontResource Lib "GDI" (ByVal lpFilename As Any) As Integer
Declare Function RemoveFontResource Lib "GDI" (ByVal lpFilename As Any) As Integer

' Dialog Box Command IDs
Public Const IDOK = 1
Public Const IDCANCEL = 2
Public Const IDABORT = 3
Public Const IDRETRY = 4
Public Const IDIGNORE = 5
Public Const IDYES = 6
Public Const IDNO = 7


' Control Manager Structures and Definitions

' Edit Control Styles
Public Const ES_LEFT = &H0&
Public Const ES_CENTER = &H1&
Public Const ES_RIGHT = &H2&
Public Const ES_MULTILINE = &H4&
Public Const ES_UPPERCASE = &H8&
Public Const ES_LOWERCASE = &H10&
Public Const ES_PASSWORD = &H20&
Public Const ES_AUTOVSCROLL = &H40&
Public Const ES_AUTOHSCROLL = &H80&
Public Const ES_NOHIDESEL = &H100&
Public Const ES_OEMCONVERT = &H400&

' Edit Control Notification Codes
Public Const EN_SETFOCUS = &H100
Public Const EN_KILLFOCUS = &H200
Public Const EN_CHANGE = &H300
Public Const EN_UPDATE = &H400
Public Const EN_ERRSPACE = &H500
Public Const EN_MAXTEXT = &H501
Public Const EN_HSCROLL = &H601
Public Const EN_VSCROLL = &H602

' Edit Control Messages
Public Const EM_GETSEL = WM_USER + 0
Public Const EM_SETSEL = WM_USER + 1
Public Const EM_GETRECT = WM_USER + 2
Public Const EM_SETRECT = WM_USER + 3
Public Const EM_SETRECTNP = WM_USER + 4
Public Const EM_SCROLL = WM_USER + 5
Public Const EM_LINESCROLL = WM_USER + 6
Public Const EM_GETMODIFY = WM_USER + 8
Public Const EM_SETMODIFY = WM_USER + 9
Public Const EM_GETLINECOUNT = WM_USER + 10
Public Const EM_LINEINDEX = WM_USER + 11
Public Const EM_SETHANDLE = WM_USER + 12
Public Const EM_GETHANDLE = WM_USER + 13
Public Const EM_GETTHUMB = WM_USER + 14
Public Const EM_LINELENGTH = WM_USER + 17
Public Const EM_REPLACESEL = WM_USER + 18
Public Const EM_SETFONT = WM_USER + 19
Public Const EM_GETLINE = WM_USER + 20
Public Const EM_LIMITTEXT = WM_USER + 21
Public Const EM_CANUNDO = WM_USER + 22
Public Const EM_UNDO = WM_USER + 23
Public Const EM_FMTLINES = WM_USER + 24
Public Const EM_LINEFROMCHAR = WM_USER + 25
Public Const EM_SETWORDBREAK = WM_USER + 26
Public Const EM_SETTABSTOPS = WM_USER + 27
Public Const EM_SETPASSWORDCHAR = WM_USER + 28
Public Const EM_EMPTYUNDOBUFFER = WM_USER + 29
Public Const EM_MSGMAX = WM_USER + 30

' Button Control Styles
Public Const BS_PUSHBUTTON = &H0&
Public Const BS_DEFPUSHBUTTON = &H1&
Public Const BS_CHECKBOX = &H2&
Public Const BS_AUTOCHECKBOX = &H3&
Public Const BS_RADIOBUTTON = &H4&
Public Const BS_3STATE = &H5&
Public Const BS_AUTO3STATE = &H6&
Public Const BS_GROUPBOX = &H7&
Public Const BS_USERBUTTON = &H8&
Public Const BS_AUTORADIOBUTTON = &H9&
Public Const BS_PUSHBOX = &HA&
Public Const BS_OWNERDRAW = &HB&
Public Const BS_LEFTTEXT = &H20&

' User Button Notification Codes
Public Const BN_CLICKED = 0
Public Const BN_PAINT = 1
Public Const BN_HILITE = 2
Public Const BN_UNHILITE = 3
Public Const BN_DISABLE = 4
Public Const BN_DOUBLECLICKED = 5

' Button Control Messages
Public Const BM_GETCHECK = WM_USER + 0
Public Const BM_SETCHECK = WM_USER + 1
Public Const BM_GETSTATE = WM_USER + 2
Public Const BM_SETSTATE = WM_USER + 3
Public Const BM_SETSTYLE = WM_USER + 4

' Static Control Constants
Public Const SS_LEFT = &H0&
Public Const SS_CENTER = &H1&
Public Const SS_RIGHT = &H2&
Public Const SS_ICON = &H3&
Public Const SS_BLACKRECT = &H4&
Public Const SS_GRAYRECT = &H5&
Public Const SS_WHITERECT = &H6&
Public Const SS_BLACKFRAME = &H7&
Public Const SS_GRAYFRAME = &H8&
Public Const SS_WHITEFRAME = &H9&
Public Const SS_USERITEM = &HA&
Public Const SS_SIMPLE = &HB&
Public Const SS_LEFTNOWORDWRAP = &HC&
Public Const SS_NOPREFIX = &H80&        '  Don't do "&" character translation

' Dialog Manager Routines
Declare Function IsDialogMessage Lib "User" (ByVal hDlg As Integer, lpMsg As msg) As Integer

Declare Sub MapDialogRect Lib "User" (ByVal hDlg As Integer, lpRect As RECT)

Declare Function DlgDirList Lib "User" (ByVal hDlg As Integer, ByVal lpPathSpec As String, ByVal nIDListBox As Integer, ByVal nIDStaticPath As Integer, ByVal wFileType As Integer) As Integer
Declare Function DlgDirSelect Lib "User" (ByVal hDlg As Integer, ByVal lpString As String, ByVal nIDListBox As Integer) As Integer
Declare Function DlgDirListComboBox Lib "User" (ByVal hDlg As Integer, ByVal lpPathSpec As String, ByVal nIDComboBox As Integer, ByVal nIDStaticPath As Integer, ByVal wFileType As Integer) As Integer
Declare Function DlgDirSelectComboBox Lib "User" (ByVal hDlg As Integer, ByVal lpString As String, ByVal nIDComboBox As Integer) As Integer

' Dialog Styles
Public Const DS_ABSALIGN = &H1&
Public Const DS_SYSMODAL = &H2&
Public Const DS_LOCALEDIT = &H20&       '  Edit items get Local storage.
Public Const DS_SETFONT = &H40& '  User specified font for Dlg controls
Public Const DS_MODALFRAME = &H80&      '  Can be combined with WS_CAPTION
Public Const DS_NOIDLEMSG = &H100&      '  WM_ENTERIDLE message will not be sent

Public Const DM_GETDEFID = WM_USER + 0
Public Const DM_SETDEFID = WM_USER + 1
Public Const DC_HASDEFID = &H534      '0x534B

' Dialog Codes
Public Const DLGC_WANTARROWS = &H1      '  Control wants arrow keys
Public Const DLGC_WANTTAB = &H2 '  Control wants tab keys
Public Const DLGC_WANTALLKEYS = &H4     '  Control wants all keys
Public Const DLGC_WANTMESSAGE = &H4     '  Pass message to control
Public Const DLGC_HASSETSEL = &H8       '  Understands EM_SETSEL message
Public Const DLGC_DEFPUSHBUTTON = &H10  '  Default pushbutton
Public Const DLGC_UNDEFPUSHBUTTON = &H20        '  Non-default pushbutton
Public Const DLGC_RADIOBUTTON = &H40    '  Radio button
Public Const DLGC_WANTCHARS = &H80      '  Want WM_CHAR messages
Public Const DLGC_STATIC = &H100        '  Static item: don't include
Public Const DLGC_BUTTON = &H2000       '  Button item: can be checked

Public Const LB_CTLCODE = 0&

' Listbox Return Values
Public Const LB_OKAY = 0
Public Const LB_ERR = (-1)
Public Const LB_ERRSPACE = (-2)

'
' The idStaticPath parameter to DlgDirList can have the following values
' ORed if the list box should show other details of the files along with
' the name of the files;

' all other details also will be returned

' Listbox Notification Codes
Public Const LBN_ERRSPACE = (-2)
Public Const LBN_SELCHANGE = 1
Public Const LBN_DBLCLK = 2
Public Const LBN_SELCANCEL = 3
Public Const LBN_SETFOCUS = 4
Public Const LBN_KILLFOCUS = 5

' Listbox messages
Public Const LB_ADDSTRING = (WM_USER + 1)
Public Const LB_INSERTSTRING = (WM_USER + 2)
Public Const LB_DELETESTRING = (WM_USER + 3)
Public Const LB_RESETCONTENT = (WM_USER + 5)
Public Const LB_SETSEL = (WM_USER + 6)
Public Const LB_SETCURSEL = (WM_USER + 7)
Public Const LB_GETSEL = (WM_USER + 8)
Public Const LB_GETCURSEL = (WM_USER + 9)
Public Const LB_GETTEXT = (WM_USER + 10)
Public Const LB_GETTEXTLEN = (WM_USER + 11)
Public Const LB_GETCOUNT = (WM_USER + 12)
Public Const LB_SELECTSTRING = (WM_USER + 13)
Public Const LB_DIR = (WM_USER + 14)
Public Const LB_GETTOPINDEX = (WM_USER + 15)
Public Const LB_FINDSTRING = (WM_USER + 16)
Public Const LB_GETSELCOUNT = (WM_USER + 17)
Public Const LB_GETSELITEMS = (WM_USER + 18)
Public Const LB_SETTABSTOPS = (WM_USER + 19)
Public Const LB_GETHORIZONTALEXTENT = (WM_USER + 20)
Public Const LB_SETHORIZONTALEXTENT = (WM_USER + 21)
Public Const LB_SETCOLUMNWIDTH = (WM_USER + 22)
Public Const LB_SETTOPINDEX = (WM_USER + 24)
Public Const LB_GETITEMRECT = (WM_USER + 25)
Public Const LB_GETITEMDATA = (WM_USER + 26)
Public Const LB_SETITEMDATA = (WM_USER + 27)
Public Const LB_SELITEMRANGE = (WM_USER + 28)
Public Const LB_MSGMAX = (WM_USER + 33)

' Listbox Styles
Public Const LBS_NOTIFY = &H1&
Public Const LBS_SORT = &H2&
Public Const LBS_NOREDRAW = &H4&
Public Const LBS_MULTIPLESEL = &H8&
Public Const LBS_OWNERDRAWFIXED = &H10&
Public Const LBS_OWNERDRAWVARIABLE = &H20&
Public Const LBS_HASSTRINGS = &H40&
Public Const LBS_USETABSTOPS = &H80&
Public Const LBS_NOINTEGRALHEIGHT = &H100&
Public Const LBS_MULTICOLUMN = &H200&
Public Const LBS_WANTKEYBOARDINPUT = &H400&
Public Const LBS_EXTENDEDSEL = &H800&
Public Const LBS_STANDARD = (LBS_NOTIFY Or LBS_SORT Or WS_VSCROLL Or WS_BORDER)

' Combo Box return Values
Public Const CB_OKAY = 0
Public Const CB_ERR = (-1)
Public Const CB_ERRSPACE = (-2)

' Combo Box Notification Codes
Public Const CBN_ERRSPACE = (-1)
Public Const CBN_SELCHANGE = 1
Public Const CBN_DBLCLK = 2
Public Const CBN_SETFOCUS = 3
Public Const CBN_KILLFOCUS = 4
Public Const CBN_EDITCHANGE = 5
Public Const CBN_EDITUPDATE = 6
Public Const CBN_DROPDOWN = 7

' Combo Box styles
Public Const CBS_SIMPLE = &H1&
Public Const CBS_DROPDOWN = &H2&
Public Const CBS_DROPDOWNLIST = &H3&
Public Const CBS_OWNERDRAWFIXED = &H10&
Public Const CBS_OWNERDRAWVARIABLE = &H20&
Public Const CBS_AUTOHSCROLL = &H40&
Public Const CBS_OEMCONVERT = &H80&
Public Const CBS_SORT = &H100&
Public Const CBS_HASSTRINGS = &H200&
Public Const CBS_NOINTEGRALHEIGHT = &H400&

' Combo Box messages
Public Const CB_GETEDITSEL = (WM_USER + 0)
Public Const CB_LIMITTEXT = (WM_USER + 1)
Public Const CB_SETEDITSEL = (WM_USER + 2)
Public Const CB_ADDSTRING = (WM_USER + 3)
Public Const CB_DELETESTRING = (WM_USER + 4)
Public Const CB_DIR = (WM_USER + 5)
Public Const CB_GETCOUNT = (WM_USER + 6)
Public Const CB_GETCURSEL = (WM_USER + 7)
Public Const CB_GETLBTEXT = (WM_USER + 8)
Public Const CB_GETLBTEXTLEN = (WM_USER + 9)
Public Const CB_INSERTSTRING = (WM_USER + 10)
Public Const CB_RESETCONTENT = (WM_USER + 11)
Public Const CB_FINDSTRING = (WM_USER + 12)
Public Const CB_SELECTSTRING = (WM_USER + 13)
Public Const CB_SETCURSEL = (WM_USER + 14)
Public Const CB_SHOWDROPDOWN = (WM_USER + 15)
Public Const CB_GETITEMDATA = (WM_USER + 16)
Public Const CB_SETITEMDATA = (WM_USER + 17)
Public Const CB_GETDROPPEDCONTROLRECT = (WM_USER + 18)
Public Const CB_MSGMAX = (WM_USER + 19)

' Scroll Bar Styles
Public Const SBS_HORZ = &H0&
Public Const SBS_VERT = &H1&
Public Const SBS_TOPALIGN = &H2&
Public Const SBS_LEFTALIGN = &H2&
Public Const SBS_BOTTOMALIGN = &H4&
Public Const SBS_RIGHTALIGN = &H4&
Public Const SBS_SIZEBOXTOPLEFTALIGN = &H2&
Public Const SBS_SIZEBOXBOTTOMRIGHTALIGN = &H4&
Public Const SBS_SIZEBOX = &H8&


' Sound Functions
'
Declare Function OpenSound Lib "Sound" () As Integer
Declare Sub CloseSound Lib "Sound" ()
Declare Function SetVoiceQueueSize Lib "Sound" (ByVal nVoice As Integer, ByVal nBytes As Integer) As Integer
Declare Function SetVoiceNote Lib "Sound" (ByVal nVoice As Integer, ByVal nValue As Integer, ByVal nLength As Integer, ByVal nCdots As Integer) As Integer
Declare Function SetVoiceAccent Lib "Sound" (ByVal nVoice As Integer, ByVal nTempo As Integer, ByVal nVolume As Integer, ByVal nMode As Integer, ByVal nPitch As Integer) As Integer
Declare Function SetVoiceEnvelope Lib "Sound" (ByVal nVoice As Integer, ByVal nShape As Integer, ByVal nRepeat As Integer) As Integer
Declare Function SetSoundNoise Lib "Sound" (ByVal nSource As Integer, ByVal nDuration As Integer) As Integer
Declare Function SetVoiceSound Lib "Sound" (ByVal nVoice As Integer, ByVal lFrequency As Long, ByVal nDuration As Integer) As Integer
Declare Function StartSound Lib "Sound" () As Integer
Declare Function StopSound Lib "Sound" () As Integer
Declare Function WaitSoundState Lib "Sound" (ByVal nState As Integer) As Integer
Declare Function SyncAllVoices Lib "Sound" () As Integer
Declare Function CountVoiceNotes Lib "Sound" (ByVal nVoice As Integer) As Integer
Declare Function GetThresholdEvent Lib "Sound" () As Integer
Declare Function GetThresholdStatus Lib "Sound" () As Integer
Declare Function SetVoiceThreshold Lib "Sound" (ByVal nVoice As Integer, ByVal nNotes As Integer) As Integer

' WaitSoundState() Constants
Public Const S_QUEUEEMPTY = 0
Public Const S_THRESHOLD = 1
Public Const S_ALLTHRESHOLD = 2

' Accent Modes
Public Const S_NORMAL = 0
Public Const S_LEGATO = 1
Public Const S_STACCATO = 2

' SetSoundNoise() Sources
Public Const S_PERIOD512 = 0    '  Freq = N/512 high pitch, less coarse hiss
Public Const S_PERIOD1024 = 1   '  Freq = N/1024
Public Const S_PERIOD2048 = 2   '  Freq = N/2048 low pitch, more coarse hiss
Public Const S_PERIODVOICE = 3  '  Source is frequency from voice channel (3)
Public Const S_WHITE512 = 4     '  Freq = N/512 high pitch, less coarse hiss
Public Const S_WHITE1024 = 5    '  Freq = N/1024
Public Const S_WHITE2048 = 6    '  Freq = N/2048 low pitch, more coarse hiss
Public Const S_WHITEVOICE = 7   '  Source is frequency from voice channel (3)

Public Const S_SERDVNA = (-1)   '  Device not available
Public Const S_SEROFM = (-2)    '  Out of memory
Public Const S_SERMACT = (-3)   '  Music active
Public Const S_SERQFUL = (-4)   '  Queue full
Public Const S_SERBDNT = (-5)   '  Invalid note
Public Const S_SERDLN = (-6)    '  Invalid note length
Public Const S_SERDCC = (-7)    '  Invalid note count
Public Const S_SERDTP = (-8)    '  Invalid tempo
Public Const S_SERDVL = (-9)    '  Invalid volume
Public Const S_SERDMD = (-10)   '  Invalid mode
Public Const S_SERDSH = (-11)   '  Invalid shape
Public Const S_SERDPT = (-12)   '  Invalid pitch
Public Const S_SERDFQ = (-13)   '  Invalid frequency
Public Const S_SERDDR = (-14)   '  Invalid duration
Public Const S_SERDSR = (-15)   '  Invalid source
Public Const S_SERDST = (-16)   '  Invalid state

' COMM declarations
'
Public Const NOPARITY = 0
Public Const ODDPARITY = 1
Public Const EVENPARITY = 2
Public Const MARKPARITY = 3
Public Const SPACEPARITY = 4

Public Const ONESTOPBIT = 0
Public Const ONE5STOPBITS = 1
Public Const TWOSTOPBITS = 2

Public Const IGNORE = 0 '  Ignore signal
Public Const INFINITE = &HFFFF  '  Infinite timeout

' Error Flags
Public Const CE_RXOVER = &H1    '  Receive Queue overflow
Public Const CE_OVERRUN = &H2   '  Receive Overrun Error
Public Const CE_RXPARITY = &H4  '  Receive Parity Error
Public Const CE_FRAME = &H8     '  Receive Framing error
Public Const CE_BREAK = &H10    '  Break Detected
Public Const CE_CTSTO = &H20    '  CTS Timeout
Public Const CE_DSRTO = &H40    '  DSR Timeout
Public Const CE_RLSDTO = &H80   '  RLSD Timeout
Public Const CE_TXFULL = &H100  '  TX Queue is full
Public Const CE_PTO = &H200     '  LPTx Timeout
Public Const CE_IOE = &H400     '  LPTx I/O Error
Public Const CE_DNS = &H800     '  LPTx Device not selected
Public Const CE_OOP = &H1000    '  LPTx Out-Of-Paper
Public Const CE_MODE = &H8000   '  Requested mode unsupported

Public Const IE_BADID = (-1)    '  Invalid or unsupported id
Public Const IE_OPEN = (-2)     '  Device Already Open
Public Const IE_NOPEN = (-3)    '  Device Not Open
Public Const IE_MEMORY = (-4)   '  Unable to allocate queues
Public Const IE_DEFAULT = (-5)  '  Error in default parameters
Public Const IE_HARDWARE = (-10)        '  Hardware Not Present
Public Const IE_BYTESIZE = (-11)        '  Illegal Byte Size
Public Const IE_BAUDRATE = (-12)        '  Unsupported BaudRate

' Events
Public Const EV_RXCHAR = &H1    '  Any Character received
Public Const EV_RXFLAG = &H2    '  Received certain character
Public Const EV_TXEMPTY = &H4   '  Transmitt Queue Empty
Public Const EV_CTS = &H8       '  CTS changed state
Public Const EV_DSR = &H10      '  DSR changed state
Public Const EV_RLSD = &H20     '  RLSD changed state
Public Const EV_BREAK = &H40    '  BREAK received
Public Const EV_ERR = &H80      '  Line status error occurred
Public Const EV_RING = &H100    '  Ring signal detected
Public Const EV_PERR = &H200    '  Printer error occured

' Escape Functions
Public Const SETXOFF = 1        '  Simulate XOFF received
Public Const SETXON = 2 '  Simulate XON received
Public Const SETRTS = 3 '  Set RTS high
Public Const CLRRTS = 4 '  Set RTS low
Public Const SETDTR = 5 '  Set DTR high
Public Const CLRDTR = 6 '  Set DTR low
Public Const RESETDEV = 7       '  Reset device if possible
Public Const GETMAXLPT = 8
Public Const GETMAXCOM = 9
Public Const GETBASEIRQ = 10


Public Const LPTx = &H80        '  Set if ID is for LPT device

Type DCB
        Id As String * 1
        BaudRate As Integer
        ByteSize As String * 1
        Parity As String * 1
        StopBits As String * 1
        RlsTimeout As Integer
        CtsTimeout As Integer
        DsrTimeout As Integer

        Bits1 As String * 1 ' The fifteen actual DCB bit-sized data fields
        Bits2 As String * 1 ' within these two bytes can be manipulated by
                          ' bitwise logical And/Or operations.  Refer to
                          ' SDKWIN.HLP for location/meaning of specific bits

        XonChar As String * 1
        XoffChar As String * 1
        XonLim As Integer
        XoffLim As Integer
        PeChar As String * 1
        EofChar As String * 1
        EvtChar As String * 1
        TxDelay As Integer
End Type

Type COMSTAT
        Bits As String * 1 ' For specific bit flags and their
                         ' meanings, refer to SDKWIN.HLP.
        cbInQue As Integer
        cbOutQue As Integer
End Type

Declare Function OpenComm Lib "User" (ByVal lpComName As String, ByVal wInQueue As Integer, ByVal wOutQueue As Integer) As Integer
Declare Function SetCommState Lib "User" (lpDCB As DCB) As Integer
Declare Function GetCommState Lib "User" (ByVal nCid As Integer, lpDCB As DCB) As Integer
Declare Function ReadComm Lib "User" (ByVal nCid As Integer, ByVal lpBuf As String, ByVal nSize As Integer) As Integer
Declare Function UngetCommChar Lib "User" (ByVal nCid As Integer, ByVal cChar As Integer) As Integer
Declare Function WriteComm Lib "User" (ByVal nCid As Integer, ByVal lpBuf As String, ByVal nSize As Integer) As Integer
Declare Function CloseComm Lib "User" (ByVal nCid As Integer) As Integer
Declare Function BuildCommDCB Lib "User" (ByVal lpDef As String, lpDCB As DCB) As Integer
Declare Function TransmitCommChar Lib "User" (ByVal nCid As Integer, ByVal cChar As Integer) As Integer
Declare Function SetCommEventMask Lib "User" (ByVal nCid As Integer, nEvtMask As Integer) As Long
Declare Function GetCommEventMask Lib "User" (ByVal nCid As Integer, ByVal nEvtMask As Integer) As Integer
Declare Function SetCommBreak Lib "User" (ByVal nCid As Integer) As Integer
Declare Function ClearCommBreak Lib "User" (ByVal nCid As Integer) As Integer
Declare Function FlushComm Lib "User" (ByVal nCid As Integer, ByVal nQueue As Integer) As Integer
Declare Function EscapeCommFunction Lib "User" (ByVal nCid As Integer, ByVal nFunc As Integer) As Integer

Type MDICREATESTRUCT
        szClass As Long
        szTitle As Long
        hOwner As Integer
        x As Integer
        y As Integer
        cx As Integer
        cy As Integer
        style As Long
        lParam As Long
End Type

Type CLIENTCREATESTRUCT
        hWindowMenu As Integer
        idFirstChild As Integer
End Type

Declare Function DefFrameProc Lib "User" (ByVal hwnd As Integer, ByVal hWndMDIClient As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function DefMDIChildProc Lib "User" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Long) As Long

Declare Function TranslateMDISysAccel Lib "User" (ByVal hWndClient As Integer, lpMsg As msg) As Integer

Declare Function ArrangeIconicWindows Lib "User" (ByVal hwnd As Integer) As Integer

'  Help engine section.

' Commands to pass WinHelp()
Public Const HELP_CONTEXT = &H1 '  Display topic in ulTopic
Public Const HELP_QUIT = &H2    '  Terminate help
Public Const HELP_INDEX = &H3   '  Display index
Public Const HELP_HELPONHELP = &H4      '  Display help on using help
Public Const HELP_SETINDEX = &H5        '  Set the current Index for multi index help
Public Const HELP_KEY = &H101           '  Display topic for keyword in offabData
Public Const HELP_MULTIKEY = &H201

Declare Function WinHelp Lib "User" (ByVal hwnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData As Any) As Integer

Type MULTIKEYHELP
        mkSize As Integer
        mkKeylist As String * 1
        szKeyphrase As String * 253 ' Array length is arbitrary; may be changed

End Type

' Window 3.1 commands to pass WinHelp()
Public Const HELP_CONTENTS = &H3     ' Display Help for a particular topic
Public Const HELP_SETCONTENTS = &H5  ' Display Help contents topic
Public Const HELP_CONTEXTPOPUP = &H8 ' Display Help topic in popup window
Public Const HELP_FORCEFILE = &H9    ' Ensure correct Help file is displayed
Public Const HELP_COMMAND = &H102    ' Execute Help macro
Public Const HELP_PARTIALKEY = &H105 ' Display topic found in keyword list
Public Const HELP_SETWINPOS = &H203  ' Display and position Help window

Declare Function GetFreeSystemResources Lib "User" (ByVal fuSysResource As Integer) As Integer

Public Const GFSR_SYSTEMRESOURCES = &H0
Public Const GFSR_GDIRESOURCES = &H1
Public Const GFSR_USERRESOURCES = &H2

Declare Sub LogError Lib "Kernel" (ByVal uErr As Integer, lpvInfo As Any)

' function declarations for profiler routines contained in Windows libraries
Declare Function ProfInsChk Lib "User" () As Integer
Declare Sub ProfSetup Lib "User" (ByVal nBufferSize As Integer, ByVal nSamples As Integer)
Declare Sub ProfSampRate Lib "User" (ByVal nRate286 As Integer, ByVal nRate386 As Integer)
Declare Sub ProfStart Lib "User" ()
Declare Sub ProfStop Lib "User" ()
Declare Sub ProfClear Lib "User" ()
Declare Sub ProfFlush Lib "User" ()
Declare Sub ProfFinish Lib "User" ()

' ***** LogParamError/LogError values *****

' Error modifier bits
Public Const ERR_WARNING = &H8000
Public Const ERR_PARAM = &H4000

Public Const ERR_SIZE_MASK = &H3000
Public Const ERR_BYTE = &H1000
Public Const ERR_WORD = &H2000
Public Const ERR_DWORD = &H3000

' ***** LogParamError() values *****

' Generic parameter values
Public Const ERR_BAD_VALUE = &H6001
Public Const ERR_BAD_FLAGS = &H6002
Public Const ERR_BAD_INDEX = &H6003
Public Const ERR_BAD_DVALUE = &H7004
Public Const ERR_BAD_DFLAGS = &H7005
Public Const ERR_BAD_DINDEX = &H7006
Public Const ERR_BAD_PTR = &H7007
Public Const ERR_BAD_FUNC_PTR = &H7008
Public Const ERR_BAD_SELECTOR = &H6009
Public Const ERR_BAD_STRING_PTR = &H700A
Public Const ERR_BAD_HANDLE = &H600B

' KERNEL parameter errors
Public Const ERR_BAD_HINSTANCE = &H6020
Public Const ERR_BAD_HMODULE = &H6021
Public Const ERR_BAD_GLOBAL_HANDLE = &H6022
Public Const ERR_BAD_LOCAL_HANDLE = &H6023
Public Const ERR_BAD_ATOM = &H6024
Public Const ERR_BAD_HFILE = &H6025

' USER parameter errors
Public Const ERR_BAD_HWND = &H6040
Public Const ERR_BAD_HMENU = &H6041
Public Const ERR_BAD_HCURSOR = &H6042
Public Const ERR_BAD_HICON = &H6043
Public Const ERR_BAD_HDWP = &H6044
Public Const ERR_BAD_CID = &H6045
Public Const ERR_BAD_HDRVR = &H6046

' GDI parameter errors
Public Const ERR_BAD_COORDS = &H7060
Public Const ERR_BAD_GDI_OBJECT = &H6061
Public Const ERR_BAD_HDC = &H6062
Public Const ERR_BAD_HPEN = &H6063
Public Const ERR_BAD_HFONT = &H6064
Public Const ERR_BAD_HBRUSH = &H6065
Public Const ERR_BAD_HBITMAP = &H6066
Public Const ERR_BAD_HRGN = &H6067
Public Const ERR_BAD_HPALETTE = &H6068
Public Const ERR_BAD_HMETAFILE = &H6069

' ***** LogError() values *****

' KERNEL errors
Public Const ERR_GALLOC = &H1
Public Const ERR_GREALLOC = &H2
Public Const ERR_GLOCK = &H3
Public Const ERR_LALLOC = &H4
Public Const ERR_LREALLOC = &H5
Public Const ERR_LLOCK = &H6
Public Const ERR_ALLOCRES = &H7
Public Const ERR_LOCKRES = &H8
Public Const ERR_LOADMODULE = &H9

' USER errors
Public Const ERR_CREATEDLG = &H40
Public Const ERR_CREATEDLG2 = &H41
Public Const ERR_REGISTERCLASS = &H42
Public Const ERR_DCBUSY = &H43
Public Const ERR_CREATEWND = &H44
Public Const ERR_STRUCEXTRA = &H45
Public Const ERR_LOADSTR = &H46
Public Const ERR_LOADMENU = &H47
Public Const ERR_NESTEDBEGINPAINT = &H48
Public Const ERR_BADINDEX = &H49
Public Const ERR_CREATEMENU = &H4A

' GDI errors
Public Const ERR_CREATEDC = &H80
Public Const ERR_CREATEMETA = &H81
Public Const ERR_DELOBJSELECTED = &H82
Public Const ERR_SELBITMAP = &H83

' Debugging support (DEBUG SYSTEM ONLY)
Type WINDEBUGINFO
        flags As Integer
        dwOptions As Long
        dwFilter As Long
        achAllocModule As String * 8
        dwAllocBreak As Long
        dwAllocCount As Long
End Type

Declare Function GetWinDebugInfo Lib "Kernel" (lpwdi As WINDEBUGINFO, ByVal flags As Integer) As Integer
Declare Function SetWinDebugInfo Lib "Kernel" (lpwdi As WINDEBUGINFO) As Integer

' WINDEBUGINFO flags values
Public Const WDI_OPTIONS = &H1
Public Const WDI_FILTER = &H2
Public Const WDI_ALLOCBREAK = &H4

' dwOptions values
Public Const DBO_CHECKHEAP = &H1
Public Const DBO_BUFFERFILL = &H4
Public Const DBO_DISABLEGPTRAPPING = &H10
Public Const DBO_CHECKFREE = &H20
Public Const DBO_SILENT = &H8000
Public Const DBO_TRACEBREAK = &H2000
Public Const DBO_WARNINGBREAK = &H1000
Public Const DBO_NOERRORBREAK = &H800
Public Const DBO_NOFATALBREAK = &H400
Public Const DBO_INT3BREAK = &H100

' dwFilter values
Public Const DBF_KERNEL = &H1000
Public Const DBF_KRN_MEMMAN = &H1
Public Const DBF_KRN_LOADMODULE = &H2
Public Const DBF_KRN_SEGMENTLOAD = &H4
Public Const DBF_USER = &H800
Public Const DBF_GDI = &H400
Public Const DBF_MMSYSTEM = &H40
Public Const DBF_PENWIN = &H20
Public Const DBF_APPLICATION = &H8
Public Const DBF_DRIVER = &H10

Public Const EW_REBOOTSYSTEM = &H43

Public Const EW_RESTARTWINDOWS = &H42

Declare Function ExitWindowsExec Lib "User" (ByVal lpszExe As String, ByVal lpszParams As String) As Integer
Declare Function IsBadReadPtr Lib "Kernel" (ByVal lp As String, ByVal cb As Integer) As Integer
Declare Function IsBadWritePtr Lib "Kernel" (ByVal lp As String, ByVal cb As Integer) As Integer
Declare Function IsBadStringPtr Lib "Kernel" (ByVal lpsz As String, ByVal cchMax As Integer) As Integer
Declare Function IsTask Lib "Kernel" (ByVal htask As Integer) As Integer

Public Const OBM_UPARROWI = 32737
Public Const OBM_DNARROWI = 32736
Public Const OBM_RGARROWI = 32735
Public Const OBM_LFARROWI = 32734

Declare Function lstrcpyn Lib "Kernel" (ByVal lpszString1 As String, ByVal lpszString2 As String, ByVal cChars As Integer) As String
Declare Function IsDBCSLeadByte Lib "Kernel" (ByVal bTestChar As Integer) As Integer

Type SIZE
        cx As Integer
        cy As Integer
End Type

' Drawing bounds accumulation APIs
Declare Function SetBoundsRect Lib "GDI" (ByVal hdc As Integer, lprcBounds As RECT, ByVal flags As Integer) As Integer
Declare Function GetBoundsRect Lib "GDI" (ByVal hdc As Integer, lprcBounds As RECT, ByVal flags As Integer) As Integer

Public Const DCB_RESET = &H1
Public Const DCB_ACCUMULATE = &H2
Public Const DCB_DIRTY = DCB_ACCUMULATE
Public Const DCB_SET = (DCB_RESET Or DCB_ACCUMULATE)
Public Const DCB_ENABLE = &H4
Public Const DCB_DISABLE = &H8

Declare Function SetWindowOrgEx Lib "GDI" (ByVal hdc As Integer, ByVal nX As Integer, ByVal nY As Integer, lpPoint As POINTAPI) As Integer
Declare Function GetWindowOrgEx Lib "GDI" (ByVal hdc As Integer, lpPoint As POINTAPI) As Integer
Declare Function SetWindowExtEx Lib "GDI" (ByVal hdc As Integer, ByVal nX As Integer, ByVal nY As Integer, lpSize As SIZE) As Integer
Declare Function GetWindowExtEx Lib "GDI" (ByVal hdc As Integer, lpSize As SIZE) As Integer
Declare Function OffsetWindowOrgEx Lib "GDI" (ByVal hdc As Integer, ByVal nX As Integer, ByVal nY As Integer, lpPoint As POINTAPI) As Integer
Declare Function ScaleWindowExtEx% Lib "GDI" (ByVal hdc%, ByVal nXnum%, ByVal nXdenom%, ByVal nYnum%, ByVal nYdenom%, lpSize As SIZE)
Declare Function SetViewportExtEx Lib "GDI" (ByVal hdc As Integer, ByVal nX As Integer, ByVal nY As Integer, lpSize As SIZE) As Integer
Declare Function GetViewportExtEx Lib "GDI" (ByVal hdc As Integer, lpSize As SIZE) As Integer
Declare Function SetViewportOrgEx Lib "GDI" (ByVal hdc As Integer, ByVal nX As Integer, ByVal nY As Integer, lpPoint As POINTAPI) As Integer
Declare Function GetViewportOrgEx Lib "GDI" (ByVal hdc As Integer, lpPoint As POINTAPI) As Integer
Declare Function OffsetViewportOrgEx Lib "GDI" (ByVal hdc As Integer, ByVal nX As Integer, ByVal nY As Integer, lpPoint As POINTAPI) As Integer
Declare Function ScaleViewportExtEx% Lib "GDI" (ByVal hdc%, ByVal nXnum%, ByVal nXdenom%, ByVal nYnum%, ByVal nYdenom%, lpSize As SIZE)

Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_BTNHIGHLIGHT = 20

Declare Function GetBrushOrgEx Lib "GDI" (ByVal hdc As Integer, lpPoint As POINTAPI) As Integer
Declare Function MoveToEx Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, lpPoint As POINTAPI) As Integer
Declare Function GetCurrentPositionEx Lib "GDI" (ByVal hdc As Integer, lpPoint As POINTAPI) As Integer
Declare Function GetTextExtentPoint Lib "GDI" (ByVal hdc As Integer, ByVal lpszString As String, ByVal cbString As Integer, lpSize As SIZE) As Integer

Public Const OUT_TT_PRECIS = 4
Public Const OUT_DEVICE_PRECIS = 5
Public Const OUT_RASTER_PRECIS = 6
Public Const OUT_TT_ONLY_PRECIS = 7

Public Const CLIP_LH_ANGLES = &H10
Public Const CLIP_TT_ALWAYS = &H20
Public Const CLIP_EMBEDDED = &H80

Declare Function GetAspectRatioFilterEx Lib "GDI" (ByVal hdc As Integer, lpAspectRatio As SIZE) As Integer

Public Const TMPF_TRUETYPE = &H4

Type PANOSE
        bFamilyType As String * 1
        bSerifStyle As String * 1
        bWeight As String * 1
        bProportion As String * 1
        bContrast As String * 1
        bStrokeVariation As String * 1
        bArmStyle As String * 1
        bLetterform As String * 1
        bMidline As String * 1
        bXHeight As String * 1
End Type

Type OUTLINETEXTMETRIC
        otmSize As Integer
        otmTextMetrics As TEXTMETRIC
        otmFiller As String * 1
        otmPanoseNumber As PANOSE
        otmfsSelection As Integer
        otmfsType As Integer
        otmsCharSlopeRise As Integer
        otmsCharSlopeRun As Integer
        otmItalicAngle As Integer
        otmEMSquare As Integer
        otmAscent As Integer
        otmDescent As Integer
        otmLineGap As Integer
        otmsCapEmHeight As Integer
        otmsXHeight As Integer
        otmrcFontBox As RECT
        otmMacAscent As Integer
        otmMacDescent As Integer
        otmMacLineGap As Integer
        otmusMinimumPPEM As Integer
        otmptSubscriptSize As POINTAPI
        otmptSubscriptOffset As POINTAPI
        otmptSuperscriptSize As POINTAPI
        otmptSuperscriptOffset As POINTAPI
        otmsStrikeoutSize As Integer
        otmsStrikeoutPosition As Integer
        otmsUnderscorePosition As Integer
        otmsUnderscoreSize As Integer
        otmpFamilyName As Long
        otmpFaceName As Long
        otmpStyleName As Long
        otmpFullName As Long
End Type

Declare Function GetOutlineTextMetrics Lib "GDI" (ByVal hdc As Integer, ByVal cbData As Integer, lpotm As OUTLINETEXTMETRIC) As Integer

' Structure passed to FONTENUMPROC
' NOTE: NEWTEXTMETRIC is the same as TEXTMETRIC plus 4 new fields
Type NEWTEXTMETRIC
        tmHeight As Integer
        tmAscent As Integer
        tmDescent As Integer
        tmInternalLeading As Integer
        tmExternalLeading As Integer
        tmAveCharWidth As Integer
        tmMaxCharWidth As Integer
        tmWeight As Integer
        tmItalic As String * 1
        tmUnderlined As String * 1
        tmStruckOut As String * 1
        tmFirstChar As String * 1
        tmLastChar As String * 1
        tmDefaultChar As String * 1
        tmBreakChar As String * 1
        tmPitchAndFamily As String * 1
        tmCharSet As String * 1
        tmOverhang As Integer
        tmDigitizedAspectX As Integer
        tmDigitizedAspectY As Integer
        ntmFlags As Long
        ntmSizeEM As Integer
        ntmCellHeight As Integer
        ntmAvgWidth As Integer
End Type

' ntmFlags field flags
Public Const NTM_REGULAR = &H40&
Public Const NTM_BOLD = &H20&
Public Const NTM_ITALIC = &H1&

Public Const LF_FULLFACESIZE = 64

Public Const TRUETYPE_FONTTYPE = &H4

Declare Function GetFontData& Lib "GDI" (ByVal hdc%, ByVal dwTable&, ByVal dwOffset&, ByVal lpvBuffer$, ByVal cbData&)
Declare Function CreateScalableFontResource% Lib "GDI" (ByVal fHidden%, ByVal lpszResourceFile$, ByVal lpszFontFile$, ByVal lpszCurrentPath$)

Type GLYPHMETRICS
        gmBlackBoxX As Integer
        gmBlackBoxY As Integer
        gmptGlyphOrigin As POINTAPI
        gmCellIncX As Integer
        gmCellIncY As Integer
End Type

Type FIXED
        fract As Integer
        value As Integer
End Type

Type MAT2
        eM11 As FIXED
        eM12 As FIXED
        eM21 As FIXED
        eM22 As FIXED
End Type

Declare Function GetGlyphOutline& Lib "GDI" (ByVal hdc%, ByVal uChar%, ByVal fuFormat%, lpgm As GLYPHMETRICS, ByVal cbBuffer&, ByVal lpBuffer$, lpmat2 As MAT2)

' GetGlyphOutline constants
Public Const GGO_METRICS = 0
Public Const GGO_BITMAP = 1
Public Const GGO_NATIVE = 2

Public Const TT_POLYGON_TYPE = 24

Public Const TT_PRIM_LINE = 1
Public Const TT_PRIM_QSPLINE = 2

Type POINTFX
        x As FIXED
        y As FIXED
End Type

Type TTPOLYCURVE
        wType As Integer
        cpfx As Integer
        apfx As POINTFX
End Type

Type TTPOLYGONHEADER
        cb As Long
        dwType As Long
        pfxStart As POINTFX
End Type

Type ABC
        abcA As Integer
        abcB As Integer
        abcC As Integer
End Type

Declare Function GetCharABCWidths Lib "GDI" (ByVal hdc As Integer, ByVal uFirstChar As Integer, ByVal uLastChar As Integer, lpabc As ABC) As Integer

Type KERNINGPAIR
        wFirst As Integer
        wSecond As Integer
        iKernAmount As Integer
End Type

Declare Function GetKerningPairs Lib "GDI" (ByVal hdc As Integer, ByVal cPairs As Integer, lpkrnpair As KERNINGPAIR) As Integer

Type RASTERIZER_STATUS
        nSize As Integer
        wFlags As Integer
        nLanguageID As Integer
End Type

' bits defined in wFlags of RASTERIZER_STATUS
Public Const TT_AVAILABLE = &H1
Public Const TT_ENABLED = &H2

Declare Function GetRasterizerCaps Lib "GDI" (lpraststat As RASTERIZER_STATUS, ByVal cb As Integer) As Integer
Declare Function SetBitmapDimensionEx Lib "GDI" (ByVal hbm As Integer, ByVal nX As Integer, ByVal nY As Integer, lpSize As SIZE) As Integer
Declare Function GetBitmapDimensionEx Lib "GDI" (ByVal hBitmap As Integer, lpDimension As SIZE) As Integer
Declare Function SetMetaFileBitsBetter Lib "GDI" (ByVal hmt As Integer) As Integer

Type DOCINFO
        cbSize As Integer
        lpszDocName As Long
        lpszOutput As Long
End Type

Declare Function StartDoc Lib "GDI" (ByVal hdc As Integer, lpdi As DOCINFO) As Integer
Declare Function StartPage Lib "GDI" (ByVal hdc As Integer) As Integer
Declare Function EndPage Lib "GDI" (ByVal hdc As Integer) As Integer
Declare Function EndDocAPI Lib "GDI" Alias "EndDoc" (ByVal hdc As Integer) As Integer
Declare Function AbortDoc Lib "GDI" (ByVal hdc As Integer) As Integer
Declare Function SpoolFile Lib "GDI" (ByVal lpszPrinter As String, ByVal lpszPort As String, ByVal lpszJob As String, ByVal lpszFile As String) As Integer

Public Const SM_CXDOUBLECLK = 36
Public Const SM_CYDOUBLECLK = 37
Public Const SM_CXICONSPACING = 38
Public Const SM_CYICONSPACING = 39
Public Const SM_MENUDROPALIGNMENT = 40
Public Const SM_PENWINDOWS = 41
Public Const SM_DBCSENABLED = 42

Declare Function SystemParametersInfo Lib "User" (ByVal uAction As Integer, ByVal uParam As Integer, ByVal lpvParam As Any, ByVal fuWinIni As Integer) As Integer

Public Const SPI_GETBEEP = 1
Public Const SPI_SETBEEP = 2
Public Const SPI_GETMOUSE = 3
Public Const SPI_SETMOUSE = 4
Public Const SPI_GETBORDER = 5
Public Const SPI_SETBORDER = 6
Public Const SPI_GETKEYBOARDSPEED = 10
Public Const SPI_SETKEYBOARDSPEED = 11
Public Const SPI_LANGDRIVER = 12
Public Const SPI_ICONHORIZONTALSPACING = 13
Public Const SPI_GETSCREENSAVETIMEOUT = 14
Public Const SPI_SETSCREENSAVETIMEOUT = 15
Public Const SPI_GETSCREENSAVEACTIVE = 16
Public Const SPI_SETSCREENSAVEACTIVE = 17
Public Const SPI_GETGRIDGRANULARITY = 18
Public Const SPI_SETGRIDGRANULARITY = 19
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPI_SETDESKPATTERN = 21
Public Const SPI_GETKEYBOARDDELAY = 22
Public Const SPI_SETKEYBOARDDELAY = 23
Public Const SPI_ICONVERTICALSPACING = 24
Public Const SPI_GETICONTITLEWRAP = 25
Public Const SPI_SETICONTITLEWRAP = 26
Public Const SPI_GETMENUDROPALIGNMENT = 27
Public Const SPI_SETMENUDROPALIGNMENT = 28
Public Const SPI_SETDOUBLECLKWIDTH = 29
Public Const SPI_SETDOUBLECLKHEIGHT = 30
Public Const SPI_GETICONTITLELOGFONT = 31
Public Const SPI_SETDOUBLECLICKTIME = 32
Public Const SPI_SETMOUSEBUTTONSWAP = 33
Public Const SPI_SETICONTITLELOGFONT = 34
Public Const SPI_GETFASTTASKSWITCH = 35
Public Const SPI_SETFASTTASKSWITCH = 36

' SystemParametersInfo flags
Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPIF_SENDWININICHANGE = &H2

Declare Function GetMessageExtraInfo Lib "User" () As Long
Declare Function GetQueueStatus Lib "User" (ByVal fuFlags As Integer) As Long

' GetQueueStatus flags
Public Const QS_KEY = &H1
Public Const QS_MOUSEMOVE = &H2
Public Const QS_MOUSEBUTTON = &H4
Public Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Public Const QS_POSTMESSAGE = &H8
Public Const QS_TIMER = &H10
Public Const QS_PAINT = &H20
Public Const QS_SENDMESSAGE = &H40

Public Const QS_ALLINPUT = &H7F


' ***** Power management *****

Public Const WM_POWER = &H48

' wParam for WM_POWER window message and DRV_POWER driver notification
Public Const PWR_OK = 1
Public Const PWR_FAIL = -1
Public Const PWR_SUSPENDREQUEST = 1
Public Const PWR_SUSPENDRESUME = 2
Public Const PWR_CRITICALRESUME = 3

Public Const GCW_ATOM = -32

Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_TRANSPARENT = &H20&

Type WINDOWPLACEMENT
        length As Integer
        flags As Integer
        showCmd As Integer
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type

Public Const WPF_SETMINPOSITION = &H1
Public Const WPF_RESTORETOMAXIMIZED = &H2

Declare Function GetWindowPlacement Lib "User" (ByVal hwnd As Integer, lpwndpl As WINDOWPLACEMENT) As Integer
Declare Function SetWindowPlacement Lib "User" (ByVal hwnd As Integer, lpwndpl As WINDOWPLACEMENT) As Integer

Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47

' WM_WINDOWPOSCHANGING/CHANGED struct pointed to by lParam
Type WINDOWPOS
        hwnd As Integer
        hWndInsertAfter As Integer
        x As Integer
        y As Integer
        cx As Integer
        cy As Integer
        flags As Integer
End Type

Declare Sub MapWindowPoints Lib "User" (ByVal hWndFrom As Integer, ByVal hWndTo As Integer, lppt As Any, ByVal cPoints As Integer)
Declare Function GetDCEx Lib "User" (ByVal hwnd As Integer, ByVal hrgnclip As Integer, ByVal fdwOptions As Long) As Integer

Public Const DCX_WINDOW = &H1&
Public Const DCX_CACHE = &H2&
Public Const DCX_CLIPCHILDREN = &H8&
Public Const DCX_CLIPSIBLINGS = &H10&
Public Const DCX_PARENTCLIP = &H20&
Public Const DCX_EXCLUDERGN = &H40&
Public Const DCX_INTERSECTRGN = &H80&
Public Const DCX_LOCKWINDOWUPDATE = &H400&
Public Const DCX_USESTYLE = &H10000

Declare Function LockWindowUpdate Lib "User" (ByVal hWndLock As Integer) As Integer
Declare Function RedrawWindow Lib "User" (ByVal hwnd As Integer, lprcUpdate As RECT, ByVal hrgnUpdate As Integer, ByVal fuRedraw As Integer) As Integer

Public Const RDW_INVALIDATE = &H1
Public Const RDW_INTERNALPAINT = &H2
Public Const RDW_ERASE = &H4
Public Const RDW_VALIDATE = &H8
Public Const RDW_NOINTERNALPAINT = &H10
Public Const RDW_NOERASE = &H20
Public Const RDW_NOCHILDREN = &H40
Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_UPDATENOW = &H100
Public Const RDW_ERASENOW = &H200
Public Const RDW_FRAME = &H400
Public Const RDW_NOFRAME = &H800

Declare Function ScrollWindowEx% Lib "User" (ByVal hwnd%, ByVal dx%, ByVal dy%, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate%, lprcUpdate As RECT, ByVal fuScroll%)

Public Const SW_SCROLLCHILDREN = &H1
Public Const SW_INVALIDATE = &H2
Public Const SW_ERASE = &H4

' WM_NCCALCSIZE return flags
Public Const WVR_ALIGNTOP = &H10
Public Const WVR_ALIGNLEFT = &H20
Public Const WVR_ALIGNBOTTOM = &H40
Public Const WVR_ALIGNRIGHT = &H80
Public Const WVR_HREDRAW = &H100
Public Const WVR_VREDRAW = &H200
Public Const WVR_REDRAW = (WVR_HREDRAW Or WVR_VREDRAW)
Public Const WVR_VALIDRECTS = &H400

' WM_NCCALCSIZE parameter structure
Type NCCALCSIZE_PARAMS
        rgrc As Long
        lppos As Long
End Type

Public Const MA_NOACTIVATEANDEAT = 4

Declare Function IsMenu Lib "User" (ByVal hmenu As Integer) As Integer

Public Const TPM_RIGHTBUTTON = &H2
Public Const TPM_LEFTALIGN = &H0
Public Const TPM_CENTERALIGN = &H4
Public Const TPM_RIGHTALIGN = &H8

Declare Function GetOpenClipboardWindow Lib "User" () As Integer
Declare Function CopyCursor Lib "User" (ByVal hinst As Integer, ByVal hcur As Integer) As Integer
Declare Function GetCursor Lib "User" () As Integer
Declare Sub GetClipCursor Lib "User" (lprc As RECT)
Declare Function CopyIcon Lib "User" (ByVal hinst As Integer, ByVal hicon As Integer) As Integer

Public Const MDIS_ALLCHILDSTYLES = &H1

' wParam values for WM_MDITILE and WM_MDICASCADE messages.
Public Const MDITILE_VERTICAL = &H0
Public Const MDITILE_HORIZONTAL = &H1
Public Const MDITILE_SKIPDISABLED = &H2

' Static Control Mesages
Public Const STM_SETICON = (WM_USER + 0)
Public Const STM_GETICON = (WM_USER + 1)

Public Const ES_READONLY = &H800&
Public Const ES_WANTRETURN = &H1000&

Public Const EM_GETFIRSTVISIBLELINE = (WM_USER + 30)
Public Const EM_SETREADONLY = (WM_USER + 31)
Public Const EM_SETWORDBREAKPROC = (WM_USER + 32)
Public Const EM_GETWORDBREAKPROC = (WM_USER + 33)
Public Const EM_GETPASSWORDCHAR = (WM_USER + 34)

' EDITWORDBREAKPROC code values
Public Const WB_LEFT = 0
Public Const WB_RIGHT = 1
Public Const WB_ISDELIMITER = 2

Public Const LBS_DISABLENOSCROLL = &H1000&

Public Const LB_SETITEMHEIGHT = (WM_USER + 33)
Public Const LB_GETITEMHEIGHT = (WM_USER + 34)
Public Const LB_FINDSTRINGEXACT = (WM_USER + 35)

Declare Function DlgDirSelectEx% Lib "User" (ByVal hWndDlg%, ByVal lpszPath$, ByVal cbPath%, ByVal idListBox%)
Declare Function DlgDirSelectComboBoxEx% Lib "User" (ByVal hWndDlg%, ByVal lpszPath$, ByVal cbPath%, ByVal idComboBox%)

Public Const CBS_DISABLENOSCROLL = &H800&

Public Const CB_SETITEMHEIGHT = (WM_USER + 19)
Public Const CB_GETITEMHEIGHT = (WM_USER + 20)
Public Const CB_SETEXTENDEDUI = (WM_USER + 21)
Public Const CB_GETEXTENDEDUI = (WM_USER + 22)
Public Const CB_GETDROPPEDSTATE = (WM_USER + 23)
Public Const CB_FINDSTRINGEXACT = (WM_USER + 24)

Public Const CBN_CLOSEUP = 8
Public Const CBN_SELENDOK = 9
Public Const CBN_SELENDCANCEL = 10

Declare Function UnhookWindowsHookEx Lib "User" (ByVal hHook As Integer) As Integer
Declare Function CallNextHookEx Lib "User" (ByVal hHook As Integer, ByVal ncode As Integer, ByVal wParam As Integer, lParam As Any) As Long

' HCBT_CREATEWND parameters pointed to by lParam
Type CBT_CREATEWND
        lpcs As Long
        hWndInsertAfter As Integer
End Type

' HCBT_ACTIVATE structure pointed to by lParam
Type CBTACTIVATESTRUCT
        fMouse As Integer
        hWndActive As Integer
End Type


Public Const WH_HARDWARE = 8

Type HARDWAREHOOKSTRUCT
        hwnd As Integer
        wMessage As Integer
        wParam As Integer
        lParam As Long
End Type

' SetWindowsHook() Shell hook code
Public Const WH_SHELL = 10

Public Const HSHELL_WINDOWCREATED = 1
Public Const HSHELL_WINDOWDESTROYED = 2
Public Const HSHELL_ACTIVATESHELLWINDOW = 3


' SetWindowsHook debug hook support
Public Const WH_DEBUG = 9

Type DEBUGHOOKINFO
        hModuleHook As Integer
        reserved As Long
        lParam As Long
        wParam As Integer
        code As Integer
End Type

Declare Function QuerySendMessage Lib "User" (ByVal hreserved1 As Integer, ByVal hreserved2 As Integer, ByVal hreserved3 As Integer, ByVal lpMessage As String) As Integer
Declare Function LockInput Lib "User" (ByVal hReserved As Integer, ByVal hWndInput As Integer, ByVal fLock As Integer) As Integer
Declare Function GetSystemDebugState Lib "User" () As Long

' Flags returned by GetSystemDebugState.
Public Const SDS_MENU = &H1
Public Const SDS_SYSMODAL = &H2
Public Const SDS_NOTASKQUEUE = &H4
Public Const SDS_DIALOG = &H8
Public Const SDS_TASKLOCKED = &H10

' Comm Baud Rate indices
Public Const CBR_110 = &HFF10
Public Const CBR_300 = &HFF11
Public Const CBR_600 = &HFF12
Public Const CBR_1200 = &HFF13
Public Const CBR_2400 = &HFF14
Public Const CBR_4800 = &HFF15
Public Const CBR_9600 = &HFF16
Public Const CBR_14400 = &HFF17
Public Const CBR_19200 = &HFF18
Public Const CBR_38400 = &HFF1B
Public Const CBR_56000 = &HFF1F
Public Const CBR_128000 = &HFF23
Public Const CBR_256000 = &HFF27

' notifications passed in low word of lParam on WM_COMMNOTIFY messages
Public Const CN_RECEIVE = &H1
Public Const CN_TRANSMIT = &H2
Public Const CN_EVENT = &H4

Declare Function EnableCommNotification Lib "User" (ByVal idComDev As Integer, ByVal hwnd As Integer, ByVal cbWriteNotify As Integer, ByVal OutQueue As Integer) As Integer

Public Const WM_COMMNOTIFY = &H44

' Driver messages
Public Const DRV_LOAD = &H1
Public Const DRV_ENABLE = &H2
Public Const DRV_OPEN = &H3
Public Const DRV_CLOSE = &H4
Public Const DRV_DISABLE = &H5
Public Const DRV_FREE = &H6
Public Const DRV_CONFIGURE = &H7
Public Const DRV_QUERYCONFIGURE = &H8
Public Const DRV_INSTALL = &H9
Public Const DRV_REMOVE = &HA
Public Const DRV_EXITSESSION = &HB
Public Const DRV_EXITAPPLICATION = &HC
Public Const DRV_POWER = &HF
Public Const DRV_RESERVED = &H800
Public Const DRV_USER = &H4000

' LPARAM of DRV_CONFIGURE message
Type DRVCONFIGINFO
        dwDCISize As Long
        lpszDCISectionName As Long
        lpszDCIAliasName As Long
End Type

' Supported return values for DRV_CONFIGURE message
Public Const DRVCNF_CANCEL = &H0
Public Const DRVCNF_OK = &H1
Public Const DRVCNF_RESTART = &H2

' Supported lParam1 of DRV_EXITAPPLICATION notification
Public Const DRVEA_NORMALEXIT = &H1
Public Const DRVEA_ABNORMALEXIT = &H2

Declare Function DefDriverProc Lib "User" (ByVal dwDriverIdentifier As Long, ByVal hdrvr As Integer, ByVal umsg As Integer, lParam1 As Any, lParam2 As Any) As Long
Declare Function OpenDriver Lib "User" (ByVal lpDriverName As String, ByVal lpSectionName As String, lParam As Any) As Integer
Declare Function CloseDriver Lib "User" (ByVal hdrvr As Integer, lParam1 As Any, lParam2 As Any) As Long
Declare Function SendDriverMessage Lib "User" (ByVal hdrvr As Integer, ByVal msg As Integer, lParam1 As Any, lParam2 As Any) As Long
Declare Function GetDriverModuleHandle Lib "User" (ByVal hdrvr As Integer) As Integer
Declare Function GetNextDriver Lib "User" (ByVal hdrvr As Integer, ByVal fdwFlag As Long) As Integer

' GetNextDriver flags
Public Const GND_FIRSTINSTANCEONLY = &H1

Public Const GND_FORWARD = &H0
Public Const GND_REVERSE = &H2

Type DRIVERINFOSTRUCT
        length As Integer
        hDriver As Integer
        hModule As Integer
        szAliasName As String * 128
End Type

Declare Function GetDriverInfo Lib "User" (ByVal hdrvr As Integer, lpdis As DRIVERINFOSTRUCT) As Integer


' ***** LZEXPAND error return codes *****

Public Const LZERROR_BADINHANDLE = (-1)  '  invalid input handle
Public Const LZERROR_BADOUTHANDLE = (-2) '  invalid output handle
Public Const LZERROR_READ = (-3)         '  corrupt compressed file format
Public Const LZERROR_WRITE = (-4)        '  out of space for output file
Public Const LZERROR_GLOBALLOC = (-5)    '  insufficient memory for LZFile struct
Public Const LZERROR_GLOBLOCK = (-6)     '  bad global handle
Public Const LZERROR_BADVALUE = (-7)     '  input parameter out of range
Public Const LZERROR_UNKNOWNALG = (-8)   '  compression algorithm not recognized

Declare Function LZStart Lib "LZexpand.dll" () As Integer
Declare Sub LZDone Lib "LZexpand.dll" ()
Declare Function CopyLZFile Lib "LZexpand.dll" (ByVal hfSource As Integer, ByVal hfDest As Integer) As Long
Declare Function LZCopy Lib "LZexpand.dll" (ByVal hfSource As Integer, ByVal hfDest As Integer) As Long
Declare Function LZInit Lib "LZexpand.dll" (ByVal hfSrc As Integer) As Integer
Declare Function GetExpandedName Lib "LZexpand.dll" (ByVal lpszSource As String, ByVal lpszBuffer As String) As Integer
Declare Function LZOpenFile Lib "LZexpand.dll" (ByVal lpszFile As String, lpOf As OFSTRUCT, ByVal style As Integer) As Integer
Declare Function LZSeek Lib "LZexpand.dll" (ByVal hfFile As Integer, ByVal lOffset As Long, ByVal nOrigin As Integer) As Long
Declare Function LZRead Lib "LZexpand.dll" (ByVal hfFile As Integer, ByVal lpvBuf As String, ByVal cb As Integer) As Integer
Declare Sub LZClose Lib "LZexpand.dll" (ByVal hfFile As Integer)


' ***** Windows Network support *****

Public Const WN_SUCCESS = &H0
Public Const WN_NOT_SUPPORTED = &H1
Public Const WN_NET_ERROR = &H2
Public Const WN_MORE_DATA = &H3
Public Const WN_BAD_POINTER = &H4
Public Const WN_BAD_VALUE = &H5
Public Const WN_BAD_PASSWORD = &H6
Public Const WN_ACCESS_DENIED = &H7
Public Const WN_FUNCTION_BUSY = &H8
Public Const WN_WINDOWS_ERROR = &H9
Public Const WN_BAD_USER = &HA
Public Const WN_OUT_OF_MEMORY = &HB
Public Const WN_CANCEL = &HC
Public Const WN_CONTINUE = &HD

Public Const WN_NOT_CONNECTED = &H30
Public Const WN_OPEN_FILES = &H31
Public Const WN_BAD_NETNAME = &H32
Public Const WN_BAD_LOCALNAME = &H33
Public Const WN_ALREADY_CONNECTED = &H34
Public Const WN_DEVICE_ERROR = &H35
Public Const WN_CONNECTION_CLOSED = &H36

Declare Function WNetAddConnection Lib "User" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Integer
Declare Function WNetCancelConnection Lib "User" (ByVal lpszName As String, ByVal bForce As Integer) As Integer
Declare Function WNetGetConnection Lib "User" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Integer) As Integer

