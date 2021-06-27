Attribute VB_Name = "mAPIConstants"

Option Explicit


Public Enum enumBorderFlags
    BF_ADJUST = &H2000
    BF_BOTTOM = &H8
    BF_DIAGONAL = &H10
    BF_FLAT = &H4000
    BF_LEFT = &H1
    BF_MIDDLE = &H800
    BF_MONO = &H8000
    BF_RIGHT = &H4
    BF_SOFT = &H1000
    BF_TOP = &H2
    BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
    BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
    BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
    BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
    BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
    BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    BF_TOPLEFT = (BF_TOP Or BF_LEFT)
    BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
End Enum

Public Enum enumBorderEdges
    BDR_RAISEDINNER = &H4
    BDR_RAISEDOUTER = &H1
    BDR_SUNKENINNER = &H8
    BDR_SUNKENOUTER = &H2
    EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
    EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
End Enum


Public Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Declare Function DrawEdge _
                                    Lib "user32" ( _
                                ByVal hdc As Long, _
                                qrc As Rect, _
                                ByVal edge As enumBorderEdges, _
                                ByVal grfFlags As enumBorderFlags) _
                            As Long


' List of different styles of keyboard entry allowed.
' Goes with the function ctlKeyPress()
Public Enum enumKeyPressAllowTypes
    NumbersOnly = 2 ^ 0
    Uppercase = 2 ^ 1
    NoSpaces = 2 ^ 2
    NoSingleQuotes = 2 ^ 3
    NoDoubleQuotes = 2 ^ 4
    AllowDecimal = 2 ^ 5
    AllowNegative = 2 ^ 6
    DatesOnly = 2 ^ 7
    TimesOnly = 2 ^ 8
    LettersOnly = 2 ^ 9
    AllowSpaces = 2 ^ 10
    AllowStars = 2 ^ 11
    AllowPounds = 2 ^ 12
End Enum


Public Const OFS_MAXPATHNAME = 128


Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Enum enumFileAttributes
    efaARCHIVE = &H20
    efaCOMPRESSED = &H800
    efaDIRECTORY = &H10
    efaHIDDEN = &H2
    efaNORMAL = &H80
    efaREADONLY = &H1
    efaSYSTEM = &H4
    efaTEMPORARY = &H100
End Enum

Public Enum enumDriveTypes
    DRIVE_CDROM = 5
    DRIVE_FIXED = 3
    DRIVE_RAMDISK = 6
    DRIVE_REMOTE = 4
    DRIVE_REMOVABLE = 2
End Enum

Public Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Public Enum enumFileNameParts
    efpFileName = 2 ^ 0
    efpFileExt = 2 ^ 1
    efpFilePath = 2 ^ 2
    efpFileNameAndExt = efpFileName + efpFileExt
    efpFileNameAndPath = efpFilePath + efpFileName
    
End Enum

Public Type typeVolumeInformation
    sRootPathName As String
    sVolumeName As String
    lVolumeSerialNo As Long
    lMaximumComponentLength As Long
    lFileSystemFlags As Long
    sFileSystemName As String
End Type

Public Const MAX_PATH = 400

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


Public Declare Function FindFirstFile _
                            Lib "kernel32" _
                            Alias "FindFirstFileA" ( _
                        ByVal lpFileName As String, _
                        lpFindFileData As WIN32_FIND_DATA) _
                    As Long

Public Declare Function FindNextFile _
                            Lib "kernel32" _
                            Alias "FindNextFileA" ( _
                        ByVal hFindFile As Long, _
                        lpFindFileData As WIN32_FIND_DATA) _
                    As Long

Public Declare Function FindClose _
                                    Lib "kernel32" ( _
                                ByVal hFindFile As Long) _
                            As Long

Public Declare Function GetVolumeInformation _
                                    Lib "kernel32" _
                                    Alias "GetVolumeInformationA" ( _
                                ByVal lpRootPathName As String, _
                                ByVal lpVolumeNameBuffer As String, _
                                ByVal nVolumeNameSize As Long, _
                                lpVolumeSerialNumber As Long, _
                                lpMaximumComponentLength As Long, _
                                lpFileSystemFlags As Long, _
                                ByVal lpFileSystemNameBuffer As String, _
                                ByVal nFileSystemNameSize As Long) _
                            As Long

Public Declare Function PathGetDriveNumber _
                                    Lib "SHLWAPI.DLL" _
                                    Alias "PathGetDriveNumberA" ( _
                                ByVal pszPath As String) _
                            As Long

Public Declare Function PathStripToRoot _
                                    Lib "SHLWAPI.DLL" _
                                    Alias "PathStripToRootA" ( _
                                ByVal pszPath As String) _
                            As Long
                                
Public Declare Function PathIsNetworkPath _
                                    Lib "SHLWAPI.DLL" _
                                    Alias "PathIsNetworkPathA" ( _
                                ByVal pszPath As String) _
                            As Boolean
            

Public Declare Function PathIsUNCServerShare _
                                    Lib "SHLWAPI.DLL" _
                                    Alias "PathIsUNCServerShareA" ( _
                                ByVal pszPath As String) _
                            As Boolean


Public Declare Function PathIsUNCServer _
                                    Lib "SHLWAPI.DLL" _
                                    Alias "PathIsUNCServerA" ( _
                                ByVal pszPath As String) _
                            As Boolean

Public Declare Function PathIsUNC _
                                    Lib "SHLWAPI.DLL" _
                                    Alias "PathIsUNCA" ( _
                                ByVal pszPath As String) _
                            As Boolean

Public Declare Function OpenFile _
                            Lib "kernel32" ( _
                        ByVal lpFileName As String, _
                        lpReOpenBuff As OFSTRUCT, _
                        ByVal wStyle As Long) _
                    As Long

Public Declare Function CloseHandle _
                            Lib "kernel32" ( _
                        ByVal hObject As Long) _
                    As Long

Public Declare Function GetFileInformationByHandle _
                            Lib "kernel32" ( _
                        ByVal hFile As Long, _
                        lpFileInformation As BY_HANDLE_FILE_INFORMATION) _
                    As Long

Public Declare Function FileTimeToSystemTime _
                            Lib "kernel32" ( _
                        lpFileTime As FILETIME, _
                        lpSystemTime As SYSTEMTIME) _
                    As Long

Public Declare Function SystemTimeToFileTime _
                                    Lib "kernel32" ( _
                                lpSystemTime As SYSTEMTIME, _
                                lpFileTime As FILETIME) _
                            As Long

Public Declare Function GetExpandedName _
                            Lib "lz32.dll" _
                            Alias "GetExpandedNameA" ( _
                        ByVal lpszSource As String, _
                        ByVal lpszBuffer As String) _
                    As Long


Public Declare Function GetShortPathName _
                            Lib "kernel32" _
                            Alias "GetShortPathNameA" ( _
                        ByVal lpszLongPath As String, _
                        ByVal lpszShortPath As String, _
                        ByVal cchBuffer As Long) _
                    As Long

Public Declare Function SetFileAttributes _
                            Lib "kernel32" _
                            Alias "SetFileAttributesA" ( _
                        ByVal lpFileName As String, _
                        ByVal dwFileAttributes As Long) _
                    As Long

Public Declare Function GetFileAttributes _
                            Lib "kernel32" _
                            Alias "GetFileAttributesA" ( _
                        ByVal lpFileName As String) _
                    As enumFileAttributes

Public Declare Function GetDriveType _
                            Lib "kernel32" _
                            Alias "GetDriveTypeA" ( _
                        ByVal nDrive As String) _
                    As Long

