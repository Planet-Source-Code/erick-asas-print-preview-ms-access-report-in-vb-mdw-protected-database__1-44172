VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.OCX"
Begin VB.Form frm_LogIn 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1905
   ClientLeft      =   2835
   ClientTop       =   3195
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125.537
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1380
      Width           =   1140
   End
   Begin MSComDlg.CommonDialog cmdlgFindAccessFile 
      Left            =   120
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   495
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1380
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   885
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   510
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   900
      Width           =   1080
   End
   Begin VB.Label LogIn 
      BackStyle       =   0  'Transparent
      Caption         =   "User LogIn"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   0
      Picture         =   "frm_LogIn.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3840
   End
End
Attribute VB_Name = "frm_LogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This API function enables us to search for the MSAccess.exe file or any other file
Private Declare Function SearchTreeForFile _
  Lib "imagehlp" (ByVal RootPath As String, _
  ByVal InputPathName As String, _
  ByVal OutputPathBuffer As String) As Long
  
Private cnn As ADODB.Connection

Private Sub cmdCancel_Click()
    Unload frm_LogIn
End Sub

Private Sub cmdOK_Click()
On Error GoTo ET
Dim strConnection As String
Dim blnPathFound As Boolean
Dim intCounter As Integer
Dim strFilePath As String
Dim strFileBuffer As String

  'Establish Connection
  Set cnn = New ADODB.Connection
  'DSN Connection - ODBC
  'Set cnn = Establish_Connection_DSN(txtUserName, txtPassword)
  'DSNLess Connection
  Set cnn = Establish_Connection_DSNLess(txtUserName, txtPassword)
  
  'Connection is successfull, so,
  'Get MS Access Path, or find the file in the hard disk
  strFilePath = GetSetting(App.Title, "AccessPath", "Path", "")
  If strFilePath = "" Then 'MSAccess.exe path not yet defined, so find it
    
    'OPTION 1, Automatically search for the MSAccess.exe file
    'Comment OPTION 2 and UnComment all code under this option if you want to use the common dialog
    
'    MsgBox "The application will now search your computer for MS Access Path!", vbInformation, "Searching For MS Access Path..."
    'Change mouse cursor
'    LogIn.Caption = "Searching..."
'    Me.MousePointer = vbHourglass
    'Iterate through all drives in the computer. In this project, I
    'used the lazy approach since I still don't know what API to use
    'to automatically determine number of drives in a computer. If you
    'know, please email me. Another alternative is to show common dialog box
    'here...
    
    'Assign initial buffer for returned value
'    strFileBuffer = Space(1028)
'    For intCounter = 67 To 90 'Drive C To Z
'      If SearchTreeForFile(Chr(intCounter) & ":\", "MSAccess.exe", strFileBuffer) Then
        'It seems that there is a space at the end of the file returned,
        'so we'll take it out
'        intCounter = InStr(1, strFileBuffer, vbNullChar, vbTextCompare)
'        strFilePath = Left$(strFileBuffer, intCounter - 1)
        'Save setting to registry so that we'll not do this again
'        SaveSetting App.Title, "AccessPath", "Path", strFilePath
'        Exit For
'      Else
'        strFilePath = vbNullString
'      End If
'    Next intCounter
    
    'OPTION 2 - Use the common dialog  to find the MSAcces.exe file
    On Error Resume Next
    With cmdlgFindAccessFile
      .DialogTitle = "Please Specify The Location Of MSAccess.exe!"
      .InitDir = "F:\Program Files\Microsoft Office\Office\MSACCESS.EXE"
      .Filter = "MSAccess Executable (MSAccess.exe)|MSAccess.exe"
      .ShowOpen
    End With
    If cmdlgFindAccessFile.FileName <> "" Then
      'Save setting to registry so that we'll not do this again
      SaveSetting App.Title, "AccessPath", "Path", cmdlgFindAccessFile.FileName
      strFilePath = cmdlgFindAccessFile.FileName
    Else
      'Restore mouse cursor
      strFilePath = vbNullString
      Me.MousePointer = vbDefault
      Exit Sub
    End If
  End If
  On Error GoTo ET
  'Restore mouse cursor
  Me.MousePointer = vbDefault
  'Verify strFilePath if null
  If strFilePath = vbNullString Then
    MsgBox "Microsoft Access was not installed in your system.Please install it first.", vbCritical, "Microsoft Access Not Installed!"
    Unload Me
    Exit Sub
  End If
  'Since we arrive here, all checking was successfull,
  'so we'll pass the UserName, Password Connection
  'And other parameters to frm_ReportDemo
  With frm_ReportDemo
    Set .AdoConnection = cnn
    .UserID = txtUserName
    .UserPassword = txtPassword
    .MSAccessPath = strFilePath
    .MDBPath = App.Path & "\ESIS.MDB"
    .MDWPath = App.Path & "\ESIS.MDW"
    .Visible = True
  End With
  'Save User Name
  SaveSetting App.Title, "AccessPath", "LastUser", txtPassword
  Unload Me
  Exit Sub
ET:
  MsgBox Err.Description, vbInformation, "Access Denied!"
  Err = False
  Exit Sub
End Sub

Private Sub Form_Load()
On Error Resume Next
  txtUserName = GetSetting(App.Title, "AccessPath", "LastUser", "")
  If txtUserName.Text <> "" Then
    txtPassword.TabIndex = 0
  End If
End Sub
