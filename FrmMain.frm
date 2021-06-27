VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Internet Error Checker For Win95/98/ME"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   1530
      Left            =   2880
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CheckBox chkSubDirs 
      Caption         =   "Search In Subdirectories?"
      Height          =   195
      Left            =   1920
      TabIndex        =   5
      Top             =   4440
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.ListBox List1 
      Height          =   1530
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Undo Fix"
      Height          =   375
      Left            =   3758
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fix"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2318
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   375
      Left            =   878
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Current OS:"
      Height          =   255
      Left            =   158
      TabIndex        =   8
      Top             =   2160
      Width           =   5535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   158
      TabIndex        =   6
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"FrmMain.frx":000C
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   158
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetVersionEx& Lib "kernel32" Alias _
    "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) 'As Long
    Private Const VER_PLATFORM_WIN32_NT = 2
    Private Const VER_PLATFORM_WIN32_WINDOWS = 1
    Private Const VER_PLATFORM_WIN32s = 0


Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 'Maintenance string For PSS usage
    End Type

Public cFiles As New colFiles
Public cFiles2 As New colFiles
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
On Error Resume Next
Command3.Enabled = False
If Label3.Caption = "Current OS: Windows 9x" Then
Command1.Enabled = False
Label2.Caption = "Status: Please Wait Searching Hard Drive For Files..."
DoEvents
List1.Clear
List2.Clear
DoEvents
    cFiles.Clear
    cFiles.LoadFiles ts.sAppend("C:", "\") & "Winsock.dll", Me.chkSubDirs.Value

    With Me.List1
        Dim l As Long
        For l = 1 To cFiles.Count
            List1.AddItem cFiles(l).sPath & cFiles(l).sNameAndExtension
            DoEvents
        Next l
        DoEvents
    End With
    
cFiles.Clear
cFiles.LoadFiles ts.sAppend("C:", "\") & "Wsock32.dll", Me.chkSubDirs.Value
DoEvents
With Me.List2
    Dim ll As Long
    For ll = 1 To cFiles.Count
        List2.AddItem cFiles(ll).sPath & cFiles(ll).sNameAndExtension
        DoEvents
        Next ll
        DoEvents
    End With

    If List1.ListCount <= 1 And List2.ListCount <= 1 Then
    Label2.Caption = "Status: No Fix Needed!"
    Command2.Enabled = False
    Else
    Label2.Caption = "Status: A Fix Is Needed!"
    Command2.Enabled = True
    End If
    DoEvents
    Command1.Enabled = True
Else
MsgBox "This program is for Windows 95/98/ME ONLY!", vbCritical
End If
Command3.Enabled = True
End Sub

Private Sub FindMyOS()
    Dim MsgEnd As String
    Dim junk
    Dim osvi As OSVERSIONINFO
    osvi.dwOSVersionInfoSize = 148
    junk = GetVersionEx(osvi)


    If junk <> 0 Then


        Select Case osvi.dwPlatformId
            Case VER_PLATFORM_WIN32s '0
            MsgEnd = "Microsoft Win32s"
            Case VER_PLATFORM_WIN32_WINDOWS '1
            If ((osvi.dwMajorVersion > 4) Or _
            ((osvi.dwMajorVersion = 4) And (osvi.dwMinorVersion > 0))) Then
            MsgEnd = "Windows 9x"
        Else
            MsgEnd = "Windows 9x"
        End If
        Case VER_PLATFORM_WIN32_NT '2
        If osvi.dwMajorVersion <= 4 Then _
        MsgEnd = "Windows NT"
        If osvi.dwMajorVersion = 5 Then _
        MsgEnd = "Windows 2000"
    End Select
End If
Label3.Caption = "Current OS: " & MsgEnd
End Sub


Private Sub Command2_Click()
On Error Resume Next
Dim x As Long
Dim fFile As Integer
fFile = FreeFile
List1.ListIndex = 0
Open App.Path & "\Fix.bat" For Output As fFile
Print #fFile, "@echo off"
Close fFile

If List1.ListCount > 1 Then
Do Until List1.ListIndex = List1.ListCount - 1
If Label3.Caption = "Current OS: Windows 9x" Then
If List1.Text = "C:\WINDOWS\WINSOCK.DLL" Then
    List1.ListIndex = List1.ListIndex + 1
    Else
    Open App.Path & "\Fix.bat" For Append As fFile
    Print #fFile, "ren " & List1.Text & " WINSOCK.OLD"
    Close fFile
    List1.ListIndex = List1.ListIndex + 1
End If
End If
    Loop
If Label3.Caption = "Current OS: Windows 9x" Then
If List1.Text = "C:\WINDOWS\WINSOCK.DLL" Then
Else
    Open App.Path & "\Fix.bat" For Append As fFile
    Print #fFile, "ren " & List1.Text & " WINSOCK.OLD"
    Close fFile
End If
End If
End If

List2.ListIndex = 0
DoEvents
If List2.ListCount > 1 Then
Do Until List2.ListIndex = List2.ListCount - 1
If Label3.Caption = "Current OS: Windows 9x" Then
If List2.Text = "C:\WINDOWS\SYSTEM\WSOCK32.DLL" Then
    List2.ListIndex = List2.ListIndex + 1
    Else
    Open App.Path & "\Fix.bat" For Append As fFile
    Print #fFile, "ren " & List2.Text & " WSOCK32.OLD"
    Close fFile
    List2.ListIndex = List2.ListIndex + 1
End If
End If
    Loop
If Label3.Caption = "Current OS: Windows 9x" Then
If List2.Text = "C:\WINDOWS\SYSTEM\WSOCK32.DLL" Then
Else
    Open App.Path & "\Fix.bat" For Append As fFile
    Print #fFile, "ren " & List2.Text & " WSOCK32.OLD"
    Close fFile
End If
End If
End If
    'Open App.Path & "\Fix.bat" For Append As fFile
    'Print #fFile, "@echo Please restart your computer when finished."
    'Close fFile


'Make the Unfix file

Open App.Path & "\UnFix.bat" For Output As fFile
Print #fFile, "@echo off"
Close fFile

DoEvents
List1.ListIndex = 0
List2.ListIndex = 0
DoEvents

If List1.ListCount > 1 Then
Do Until List1.ListIndex = List1.ListCount - 1
If Label3.Caption = "Current OS: Windows 9x" Then
If List1.Text = "C:\WINDOWS\WINSOCK.DLL" Then
    List1.ListIndex = List1.ListIndex + 1
    Else
    x = Len(List1.Text)
    Open App.Path & "\UnFix.bat" For Append As fFile
    Print #fFile, "ren " & Left(List1.Text, x - 3) & "OLD" & " WINSOCK.DLL"
    Close fFile
    List1.ListIndex = List1.ListIndex + 1
End If
End If
    Loop
If Label3.Caption = "Current OS: Windows 9x" Then
If List1.Text = "C:\WINDOWS\WINSOCK.DLL" Then
Else
    x = Len(List1.Text)
    Open App.Path & "\UnFix.bat" For Append As fFile
    Print #fFile, "ren " & Left(List1.Text, x - 3) & "OLD" & " WINSOCK.DLL"
    Close fFile
End If
End If
End If

List2.ListIndex = 0
DoEvents
If List2.ListCount > 1 Then
Do Until List2.ListIndex = List2.ListCount - 1
If Label3.Caption = "Current OS: Windows 9x" Then
If List2.Text = "C:\WINDOWS\SYSTEM\WSOCK32.DLL" Then
    List2.ListIndex = List2.ListIndex + 1
    Else
    x = Len(List2.Text)
    Open App.Path & "\UnFix.bat" For Append As fFile
    Print #fFile, "ren " & Left(List2.Text, x - 3) & "OLD" & " WSOCK32.DLL"
    Close fFile
    List2.ListIndex = List2.ListIndex + 1
End If
End If
    Loop
If Label3.Caption = "Current OS: Windows 9x" Then
If List2.Text = "C:\WINDOWS\SYSTEM\WSOCK32.DLL" Then
Else
    x = Len(List2.Text)
    Open App.Path & "\UnFix.bat" For Append As fFile
    Print #fFile, "ren " & Left(List2.Text, x - 3) & "OLD" & " WSOCK32.DLL"
    Close fFile
End If
End If
End If


Call ShellExecute(hwnd, "Open", App.Path & "\Fix.bat", "", App.Path, 1)
DoEvents
Label2.Caption = "Status: Please Restart Your Computer"
End Sub

Private Sub Command3_Click()
Call ShellExecute(hwnd, "Open", App.Path & "\UnFix.bat", "", App.Path, 1)
DoEvents
Label2.Caption = "Status: Please Restart Your Computer"
End Sub

Private Sub Form_Load()
Me.Caption = "Internet Error Checker For Win95/98/ME v" & App.Major & "." & App.Minor & "." & App.Revision
Call FindMyOS
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
