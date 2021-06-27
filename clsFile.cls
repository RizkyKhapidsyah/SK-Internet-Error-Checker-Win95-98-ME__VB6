VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsFileName As String 'local copy

Public Property Get bFixedDisk() As Boolean
    bFixedDisk = ts.fileOpenStructure(mvarsFileName).fFixedDisk
End Property

Public Property Get dCreated() As Date
    dCreated = ts.timeFileToDate(ts.fileInformation(mvarsFileName).ftCreationTime)
End Property

Public Property Get dLastAccessed() As Date
    dLastAccessed = ts.timeFileToDate(ts.fileInformation(mvarsFileName).ftLastAccessTime)
End Property

Public Property Get sExpandedName() As String
    sExpandedName = ts.fileExpandedName(mvarsFileName)
End Property

Public Property Get sShortName() As String
    sShortName = ts.fileShortName(mvarsFileName)
End Property

Public Property Let sFilename(ByVal vData As String)
    mvarsFileName = vData
End Property

Public Property Get sFilename() As String
Attribute sFilename.VB_UserMemId = 0
    sFilename = mvarsFileName
End Property

'''Public Function Edit() As Boolean
'''End Function
'''
'''Public Function Move(sDestination As String, Optional bVerbose As Boolean) As Boolean
'''End Function
'''
'''Public Function Copy(sNewFileName As String, Optional bVerbose As Boolean) As Boolean
'''End Function
'''
'''Public Function Delete() As Boolean
'''End Function
'''
'''Public Function Launch() As Boolean
'''End Function
'''
'''Public Function Rename(ByVal sNewFileName As String, Optional ByVal bVerbose As Boolean) As Boolean
'''End Function

Public Property Let eAttributes(ByVal vData As enumFileAttributes)
    SetFileAttributes Me.sFilename, vData
End Property

Public Property Get eAttributes() As enumFileAttributes
    eAttributes = ts.fileAttributes(mvarsFileName)
End Property

Public Property Get lSize() As Long
    On Error Resume Next
    lSize = FileLen(mvarsFileName)
    If Err.Number <> 0 Then
        lSize = fileLength(mvarsFileName)
    End If
End Property


Public Property Get dLastModified() As Date
    dLastModified = ts.timeFileToDate(ts.fileInformation(ts.fileShortName(mvarsFileName)).ftLastWriteTime)
End Property

Public Property Get sName() As String
    sName = ts.sFilename(mvarsFileName, efpFileName)
End Property

Public Property Get sNameAndExtension() As String
    sNameAndExtension = ts.sFilename(mvarsFileName, efpFileNameAndExt)
End Property


Public Property Get sExtension() As String
    sExtension = ts.sFilename(mvarsFileName, efpFileExt)
End Property

Public Property Get sPath() As String
    sPath = ts.sFilename(mvarsFileName, efpFilePath)
End Property

Public Property Get sDriveLetter() As String
    Dim lNo As Long
    lNo = PathGetDriveNumber(mvarsFileName)
    If lNo > -1 Then
        sDriveLetter = Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ", 1 + lNo, 1) & ":"
    Else
        sDriveLetter = ""
    End If
    
End Property

Public Property Get sPathRoot() As String
    Dim sRet As String
    
    sRet = ts.fileRoot(mvarsFileName)
    If Right(sRet, 1) <> "\" And Trim(sRet) <> "" Then
        sRet = sRet & "\"
    End If
    sPathRoot = sRet
End Property

Public Property Get bNetworkPath() As Boolean
    bNetworkPath = PathIsNetworkPath(mvarsFileName)
End Property

Public Property Get bUNCServerShare() As Boolean
    bUNCServerShare = PathIsUNCServerShare(mvarsFileName)
End Property

Public Property Get bUNCServer() As Boolean
    bUNCServer = PathIsUNCServer(mvarsFileName)
End Property

Public Property Get bUNC() As Boolean
    bUNC = PathIsUNC(mvarsFileName)
End Property

Public Property Get eVolumeType() As enumDriveTypes
    eVolumeType = GetDriveType(Me.sPathRoot)
    
End Property

Public Property Get lVolumeSerialNo() As Long
    lVolumeSerialNo = ts.fileInformation(mvarsFileName).dwVolumeSerialNumber
End Property

Public Property Get sVolumeName() As String
    sVolumeName = ts.volumeInformation(Me.sPathRoot).sVolumeName
End Property


