VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  About this product"
   ClientHeight    =   2040
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6405
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1408.045
   ScaleMode       =   0  'User
   ScaleWidth      =   6014.626
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAbout 
      Height          =   2115
      Left            =   30
      TabIndex        =   0
      Top             =   -75
      Width           =   6315
      Begin VB.PictureBox picPaul 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   1320
         Left            =   5295
         ScaleHeight     =   1290
         ScaleWidth      =   915
         TabIndex        =   9
         Top             =   180
         Width           =   945
         Begin VB.Image imgPaul 
            Height          =   1290
            Left            =   0
            Picture         =   "frmAbout.frx":0C82
            Stretch         =   -1  'True
            Top             =   0
            Width           =   915
         End
      End
      Begin VB.Frame fraLine 
         Height          =   75
         Left            =   15
         TabIndex        =   7
         Top             =   1530
         Width           =   6285
      End
      Begin VB.PictureBox picK1PL 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   1320
         Left            =   60
         ScaleHeight     =   1290
         ScaleWidth      =   915
         TabIndex        =   6
         Top             =   180
         Width           =   945
         Begin VB.Image imgK1PL 
            Appearance      =   0  'Flat
            Height          =   1290
            Left            =   0
            Picture         =   "frmAbout.frx":6CF6
            Stretch         =   -1  'True
            Top             =   0
            Width           =   915
         End
      End
      Begin VB.CommandButton cmdSysInfo 
         Caption         =   "&System Info..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   1245
      End
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "OK, whatever"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4935
         TabIndex        =   1
         Top             =   1665
         Width           =   1260
      End
      Begin VB.Timer tmrAbout 
         Interval        =   20
         Left            =   45
         Top             =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hand Written By Brian Battles, WS1O"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   1305
         TabIndex        =   8
         Top             =   1755
         Width           =   3705
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         Height          =   210
         Left            =   2715
         TabIndex        =   5
         Top             =   525
         Width           =   1200
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application Title"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   1830
         TabIndex        =   4
         Top             =   225
         Width           =   2865
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDescription 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1980
         TabIndex        =   3
         Top             =   1095
         Width           =   2670
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Module     : frmAbout
' Description:
' Procedures : GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String)
'              StartSysInfo()
'              cmdOK_Click()
'              cmdSysInfo_Click()
'              Form_Load()
'              tmrAbout_Timer()

Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private bGoLeft  As Boolean
Private bGoRight As Boolean
Private bSpeedUp As Integer

Private intSpeed As Integer
Private intBng   As Integer

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    
    ' --------------------------------------------------
    ' Comments  :
    ' Parameters: KeyRoot
    '             KeyName
    '             SubKeyRef
    '             KeyVal -
    ' Returns   : Boolean -
    ' --------------------------------------------------
    
    On Error GoTo Err_GetKeyValue
    
    Dim I          As Long    ' Loop Counter
    Dim rc         As Long    ' Return Code
    Dim hKey       As Long    ' Handle To An Open Registry Key
    Dim hDepth     As Long
    Dim KeyValType As Long    ' Data Type Of A Registry Key
    Dim tmpVal     As String  ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long    ' Size Of Registry Key Variable
    
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    If (rc <> ERROR_SUCCESS) Then
        GoTo GetKeyError          ' Handle Error...
    End If
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    ' Retrieve Registry Key Value...
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                    KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
    If (rc <> ERROR_SUCCESS) Then
        GoTo GetKeyError          ' Handle Errors
    End If
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    ' Determine Key Value Type For Conversion...
    Select Case KeyValType                                  ' Search Data Types...
        Case REG_SZ                                             ' String Registry Key Data Type
            KeyVal = tmpVal                                     ' Copy String Value
        Case REG_DWORD                                          ' Double Word Registry Key Data Type
            For I = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, I, 1)))   ' Build Value Char. By Char.
            Next
            KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:
    
    ' Cleanup After An Error Has Occurred...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    
Exit_GetKeyValue:
    
    Exit Function
    
Err_GetKeyValue:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during GetKeyValue, in frmAbout", vbInformation, "Advisory"
            Resume Exit_GetKeyValue
    End Select
    
End Function
Public Sub StartSysInfo()
    
    ' --------------------------------------------------
    ' Comments  :
    ' --------------------------------------------------
    
    On Error GoTo SysInfoErr
    
    Dim rc          As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
            ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
        ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    Shell SysInfoPath, vbNormalFocus
    Exit Sub
    
SysInfoErr:

    MsgBox "System Information Is Unavailable At This Time", vbOKOnly

End Sub
Private Sub cmdOK_Click()
    
    ' --------------------------------------------------
    ' Comments  :
    ' --------------------------------------------------
    
    On Error GoTo Err_cmdOK_Click
    
    Unload Me
    
Exit_cmdOK_Click:
    
    End
    Exit Sub
    
Err_cmdOK_Click:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during cmdOK_Click, in frmAbout", vbInformation, "Advisory"
            Resume Exit_cmdOK_Click
    End Select
    
End Sub
Private Sub cmdSysInfo_Click()
    
    ' --------------------------------------------------
    ' Comments  :
    ' --------------------------------------------------
    
    On Error GoTo Err_cmdSysInfo_Click
    
    StartSysInfo
    
Exit_cmdSysInfo_Click:
    
    Exit Sub
    
Err_cmdSysInfo_Click:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during cmdSysInfo_Click, in frmAbout", vbInformation, "Advisory"
            Resume Exit_cmdSysInfo_Click
    End Select
    
End Sub
Private Sub Form_Activate()
   
    '---------------------------------------------------------------
    ' Purpose   :
    ' Modified  : 6/15/2001 By BB
    '---------------------------------------------------------------

    On Error GoTo Err_Form_Activate

    'Start_Scroll Me, picK1PL, 1

Exit_Form_Activate:
    
    On Error GoTo 0
    Exit Sub

Err_Form_Activate:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmAbout" & " during " & "Form_Activate" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_Form_Activate
    End Select
    
End Sub
Private Sub Form_load()
    
    ' --------------------------------------------------
    ' Comments  :
    ' --------------------------------------------------
    
    On Error GoTo Err_Form_Load
    
    LoadForm
    
Exit_Form_Load:
    
    Exit Sub
    
Err_Form_Load:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during Form_Load, in frmAbout", vbInformation, "Advisory"
            Resume Exit_Form_Load
    End Select
    
End Sub
Private Sub tmrAbout_Timer()
    
    ' --------------------------------------------------
    ' Comments  :
    ' --------------------------------------------------
    
    On Error GoTo Err_tmrAbout_Timer
    
    Randomize
    If (Rnd * 99) > 98 Then
        Label1.ForeColor = QBColor(Int(Rnd * 15))
    End If
    If (Rnd * 99) > 98 Then
        lblTitle.ForeColor = QBColor(Int(Rnd * 15))
    End If
    'IMAGE ON THE LEFT
    If bGoRight = True Then
        ' we're moving to the right
        Select Case picK1PL.Left
            Case Is >= picPaul.Left
                intBng = Int(Rnd * 99)
                Select Case intBng
                    Case Is > 98
                        PlaySound App.Path & "\" & "boing.wav", 0, &H1 'SND_ASYNC
                    Case 96 To 98
                        PlaySound App.Path & "\" & "boink.wav", 0, &H1
                    Case 94 To 96
                        PlaySound App.Path & "\" & "boingg.wav", 0, &H1
                    Case Else
                        PlaySound App.Path & "\" & "boin1.wav", 0, &H1
                End Select
                If intSpeed >= 3000 Then
                    bSpeedUp = False
                ElseIf intSpeed < 20 Then
                    bSpeedUp = True
                End If
                If bSpeedUp Then
                    intSpeed = intSpeed * 2  '+ 10
                Else
                    intSpeed = intSpeed / 2   '- 10
                End If
                If picK1PL.Left >= fraAbout.Left + fraAbout.Width Then
                    picK1PL.Left = picK1PL.Left - intSpeed
                    bGoRight = False
                Else
                    picK1PL.Left = picK1PL.Left - intSpeed
                    bGoRight = False
                End If
            Case Is < fraAbout.Left + fraAbout.Width ' right side of form
                If picK1PL.Left >= fraAbout.Left + fraAbout.Width Then
                    picK1PL.Left = picK1PL.Left - intSpeed
                    bGoRight = False
                Else
                    picK1PL.Left = picK1PL.Left + intSpeed
                    bGoRight = True
                End If
            Case Else
                ' huh?
                If picK1PL.Left >= fraAbout.Left + fraAbout.Width Then
                    picK1PL.Left = picK1PL.Left - intSpeed
                    bGoRight = False
                Else
                    picK1PL.Left = picK1PL.Left + intSpeed
                    bGoRight = True
                End If
        End Select
    Else
        ' we're moving to the left
        If picK1PL.Left + picK1PL.Width <= fraAbout.Left Then ' left side of form
            picK1PL.Left = picK1PL.Left + intSpeed
            bGoRight = True
                If picK1PL.Left >= fraAbout.Left + fraAbout.Width Then
                    picK1PL.Left = picK1PL.Left - intSpeed
                    bGoRight = False
                End If
        Else
            picK1PL.Left = picK1PL.Left - intSpeed
            bGoRight = False
        End If
    End If
    If picPaul.Left > fraAbout.Left + fraAbout.Width + 2000 Then
        picPaul.Left = picPaul.Left - intSpeed
    End If
    If bGoRight Then
        picPaul.Left = picPaul.Left - intSpeed
    Else
        picPaul.Left = picPaul.Left + intSpeed
    End If
    
Exit_tmrAbout_Timer:
    
    Exit Sub
    
Err_tmrAbout_Timer:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during tmrAbout_Timer, in frmAbout", vbInformation, "Advisory"
            Resume Exit_tmrAbout_Timer
    End Select
    
End Sub
Private Sub LoadForm()
   
    '---------------------------------------------------------------
    ' Purpose   :
    '---------------------------------------------------------------

    On Error GoTo Err_LoadForm

    bGoRight = True
    bGoLeft = True
    intSpeed = 30
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title

Exit_LoadForm:
    
    On Error GoTo 0
    Exit Sub

Err_LoadForm:

    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In frmAbout" & " during " & "LoadForm" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_LoadForm
    End Select

End Sub
