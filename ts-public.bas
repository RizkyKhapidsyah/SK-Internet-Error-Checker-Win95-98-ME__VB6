Attribute VB_Name = "ts"

Option Explicit

Public Function rectMake(ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long) As Rect
    Dim tRet As Rect
    tRet.Bottom = lBottom
    tRet.Top = lTop
    tRet.Left = lLeft
    tRet.Right = lRight
    rectMake = tRet
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlKeyPress
'    This function is handy for wrapping input to textboxes
'    or other controls that have the KeyPress event to implement
'    standard types of input masks.
'       Example:
'            Private Sub txtPlaceOfEmployment_KeyPress(KeyAscii As Integer)
'                KeyAscii = ts.wrapKeyPress(KeyAscii, Uppercase + NoDoubleQuotes)
'            End Sub
Public Function ctlKeyPress(ByVal KeyAscii As KeyCodeConstants, ByVal TypeToAllow As enumKeyPressAllowTypes) As Integer
    
    Dim ltrKeyAscii As Integer
    ltrKeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    ' By default pass the keystroke through and then optionally kill it
    ctlKeyPress = KeyAscii
    
    ' Default Keystrokes to allow (enter, backspace, delete, escape)
    If _
        KeyAscii = vbKeyReturn Or _
        KeyAscii = vbKeyEscape Or _
        KeyAscii = vbKeyBack Or _
        KeyAscii = vbKeyDelete Then
        
        Exit Function
    End If
    
    ' NumbersOnly
    If (TypeToAllow And NumbersOnly) Then
        Select Case True
            Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9
            Case (KeyAscii = vbKeySubtract Or KeyAscii = Asc("-")) And (TypeToAllow And AllowNegative)
            Case KeyAscii = Asc("#") And (TypeToAllow And AllowPounds)
            Case KeyAscii = Asc("*") And (TypeToAllow And AllowStars)
            Case KeyAscii = vbKeyDecimal And (TypeToAllow And AllowDecimal)
            Case KeyAscii = vbKeySpace And (TypeToAllow And AllowSpaces)
            Case Else
                KeyAscii = 0
        End Select
    End If
    
    ' DatesOnly
    If (TypeToAllow And DatesOnly) Then
        Select Case True
            Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9
            Case KeyAscii = vbKeyDivide Or KeyAscii = Asc("/")
            Case Else
                KeyAscii = 0
        End Select
    End If
    
    ' TimesOnly
    If (TypeToAllow And TimesOnly) Then
        Select Case True
            Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9
            Case KeyAscii = Asc(":") Or KeyAscii = Asc(";")
                ctlKeyPress = Asc(":")
            Case ltrKeyAscii = vbKeyA Or ltrKeyAscii = vbKeyP Or ltrKeyAscii = vbKeyM
            Case Else
                KeyAscii = 0
        End Select
    End If
            
    ' LettersOnly
    If (TypeToAllow And LettersOnly) Then
        Select Case True
            Case ltrKeyAscii >= vbKeyA And ltrKeyAscii <= vbKeyZ
            Case Else
                KeyAscii = 0
        End Select
    End If
            
    ' UpperCase
    If (TypeToAllow And Uppercase) Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    
    ' No Spaces
    If (TypeToAllow And NoSpaces) And KeyAscii = vbKeySpace Then
        KeyAscii = 0
    End If
    
    ' No Double Quotes
    If (TypeToAllow And NoDoubleQuotes) And KeyAscii = Asc("""") Then
        KeyAscii = Asc("'")
    End If
    
    ' No Single Quotes
    If (TypeToAllow And NoSingleQuotes) And KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
    
    ctlKeyPress = KeyAscii
    
End Function

Public Function fileOpenStructure( _
                                ByVal sFilename As String) _
                As OFSTRUCT
    
    Dim tOF As OFSTRUCT
    Dim lHandle As Long
    lHandle = OpenFile(sFilename, tOF, 0)
    CloseHandle lHandle
    fileOpenStructure = tOF
    
End Function


Public Function fileInformation( _
                            ByVal sFilename As String) _
                As BY_HANDLE_FILE_INFORMATION
                    
    Dim tInfo As BY_HANDLE_FILE_INFORMATION
    Dim tOF As OFSTRUCT
    Dim lHandle As Long
    lHandle = OpenFile(sFilename, tOF, 0)
    If lHandle > 0 Then
        GetFileInformationByHandle lHandle, tInfo
    End If
    fileInformation = tInfo
    CloseHandle lHandle
    
End Function


Public Function timeFileToDate( _
                                ft As FILETIME) _
                As Date
                    
    Dim tSysTime As SYSTEMTIME
    FileTimeToSystemTime ft, tSysTime
    timeFileToDate = ts.timeSysToDate(tSysTime)
                    
End Function
                            
Public Function timeDateToFile( _
                                ByVal dDate As Date) _
                As FILETIME
                
    Dim tRet As FILETIME
    Dim tSys As SYSTEMTIME
    tSys = timeDateToSys(dDate)
    SystemTimeToFileTime tSys, tRet
    timeDateToFile = tRet
    
End Function

Public Function timeSysToDate( _
                            st As SYSTEMTIME) _
                As Date
    
    timeSysToDate = CDate(Format(st.wMonth, "00") & "/" & Format(st.wDay, "00") & "/" & Format(st.wYear, "0000") & " " & Format(st.wHour, "00") & ":" & Format(st.wMinute, "00") & ":" & Format(st.wSecond, "00"))
    
End Function

Public Function timeDateToSys( _
                            ByVal dDateTime As Date) _
                As SYSTEMTIME
                    
    Dim tRet As SYSTEMTIME
    tRet.wDay = Day(dDateTime)
    tRet.wMonth = Month(dDateTime)
    tRet.wYear = Year(dDateTime)
    tRet.wHour = Hour(dDateTime)
    tRet.wMinute = Minute(dDateTime)
    tRet.wSecond = Second(dDateTime)
    timeDateToSys = tRet
    
End Function


Public Function fileExpandedName( _
                                ByVal sFilename As String) _
                As String
                    
    Dim sBuffer As String
    sBuffer = Space(1024)
    GetExpandedName sFilename, sBuffer
    fileExpandedName = ts.sNT(sBuffer)
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sNT
'    Standing for NullTrim, this function will take in a null terminated string
'    and clip of the extra junk.  Useful for DLL calls that return results in
'    a string buffer.
Public Function sNT( _
                        ByVal sString As String) _
                As String
                
    Dim iNullLoc As Integer
    iNullLoc = InStr(sString, Chr(0))
    If iNullLoc > 0 Then
        sNT = Left(sString, iNullLoc - 1)
    Else
        sNT = sString
    End If
End Function

Public Function fileShortName( _
                                ByVal sFilename As String) _
                As String
                    
    Dim sBuffer As String
    sBuffer = Space(1024)
    GetShortPathName sFilename, sBuffer, Len(sBuffer)
    fileShortName = ts.sNT(sBuffer)
    
End Function

Public Function fileAttributes( _
                            ByVal sFilename As String) _
                As enumFileAttributes
    
    fileAttributes = GetFileAttributes(sFilename)
    
End Function

Public Function fileLength(ByVal sFilename As String) As Long

    Dim FileHandle As Integer
    
    FileHandle = FreeFile
    On Error Resume Next
    Open sFilename For Input As #FileHandle
    fileLength = LOF(FileHandle)
    Close #FileHandle
    On Error GoTo 0
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sFileName
'    This function is used to parse keys peices of info from a
'    filename that is passed into it.
Public Function sFilename( _
                            ByVal sFIle As String, _
                            ByVal ePortions As enumFileNameParts) _
                As String
    
    Dim lFirstPeriod As Long, lFirstBackSlash As Long
    lFirstPeriod = InStrRev(sFIle, ".")
    lFirstBackSlash = InStrRev(sFIle, "\")
    Dim sPath As String, sName As String, sExt As String
    If lFirstBackSlash > 0 Then
        sPath = Left(sFIle, lFirstBackSlash)
    End If
    If lFirstPeriod > 0 And lFirstPeriod > lFirstBackSlash Then
        sExt = Mid(sFIle, lFirstPeriod + 1)
        sName = Mid(sFIle, lFirstBackSlash + 1, lFirstPeriod - lFirstBackSlash - 1)
    Else
        sName = Mid(sFIle, lFirstBackSlash + 1)
    End If
    Dim sRet As String
    If ePortions And efpFilePath Then
        sRet = sRet & sPath
    End If
    If ePortions And efpFileName Then
        sRet = sRet & sName
    End If
    If ePortions And efpFileExt Then
        If sRet <> "" Then
            sRet = sRet & "." & sExt
        Else
            sRet = sRet & sExt
        End If
    End If
    sFilename = sRet
    
End Function

Public Function fileRoot(ByVal sFilename As String) As String
    
    Dim lngResult As Long
    lngResult = PathStripToRoot(sFilename)
    If lngResult <> 0 Then
        If InStr(sFilename, vbNullChar) > 0 Then
            fileRoot = Left$(sFilename, InStr(sFilename, vbNullChar) - 1)
        Else
            fileRoot = sFilename
        End If
    End If
    
End Function


Public Function volumeInformation(ByVal sDrive As String) As typeVolumeInformation
    
    Dim Ret As typeVolumeInformation
    Ret.sRootPathName = sDrive
    Ret.sFileSystemName = Space(1024)
    Ret.sVolumeName = Space(1024)
    GetVolumeInformation Ret.sRootPathName, Ret.sVolumeName, Len(Ret.sVolumeName), Ret.lVolumeSerialNo, Ret.lMaximumComponentLength, Ret.lFileSystemFlags, Ret.sFileSystemName, Len(Ret.sFileSystemName)
    Ret.sFileSystemName = ts.sNT(Ret.sFileSystemName)
    Ret.sVolumeName = ts.sNT(Ret.sVolumeName)
    volumeInformation = Ret
    
End Function

Public Function fileCount( _
                            ByVal sSpec As String) _
                As Long
                    
    Dim tInfo As WIN32_FIND_DATA
    Dim lCnt As Long
    Dim lFind As Long, lMatch As Long
    lFind = FindFirstFile(sSpec, tInfo)
    lMatch = 99
    Do While lFind > 0 And lMatch > 0
        lCnt = lCnt + 1
        lMatch = FindNextFile(lFind, tInfo)
    Loop
    fileCount = lCnt
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' sAppend
'    This function will append a string to another string when it is not already the
'    last character or characters in the string (useful for ensuring a string is ended
'    with a vbCrLf or when building paths, a backslash \).
Public Function sAppend(ByVal s2AppendTo As String, ByVal sChars2Append As String) As String
    
    If Right(s2AppendTo, Len(sChars2Append)) <> sChars2Append Then
        sAppend = s2AppendTo & sChars2Append
    Else
        sAppend = s2AppendTo
    End If
    
End Function

'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' ctlSetFocus
'    This function will set the current focus to the specified
'    control or screen object without throwing an error if the
'    object cannot receive focus.
Public Function ctlSetFocus(ByRef ObjToSetFocusTo As Object) As Boolean
    On Error Resume Next
    ObjToSetFocusTo.SetFocus
    ctlSetFocus = Err.Number = 0
    On Error GoTo 0
End Function
