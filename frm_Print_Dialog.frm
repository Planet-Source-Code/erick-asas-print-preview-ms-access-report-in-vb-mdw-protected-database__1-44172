VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_Print_Dialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5055
   ControlBox      =   0   'False
   Icon            =   "frm_Print_Dialog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4815
      Begin VB.CommandButton cmdCancel 
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
         Left            =   3120
         TabIndex        =   14
         Top             =   960
         Width           =   1380
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Ok"
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
         Left            =   3120
         TabIndex        =   13
         Top             =   480
         Width           =   1380
      End
      Begin VB.Frame FrameArray 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date Interval"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2025
         Index           =   6
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   4455
         Begin VB.CheckBox chkFields 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Include This Criteria"
            DataField       =   "Posted"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   2535
         End
         Begin MSComCtl2.DTPicker DTPickerDate 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   22740993
            CurrentDate     =   37690
         End
         Begin MSComCtl2.DTPicker DTPickerTime 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   1440
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   22740994
            CurrentDate     =   37690
         End
         Begin MSComCtl2.DTPicker DTPickerDate 
            Height          =   375
            Index           =   1
            Left            =   2400
            TabIndex        =   9
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   22740993
            CurrentDate     =   37690
         End
         Begin MSComCtl2.DTPicker DTPickerTime 
            Height          =   375
            Index           =   1
            Left            =   2400
            TabIndex        =   10
            Top             =   1440
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   22740994
            CurrentDate     =   37690
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00FFFFFF&
            Caption         =   "End Date/Time"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   12
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Start Date/Time"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.Frame FrameArray 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1185
         Index           =   7
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
         Begin VB.OptionButton chkForPostOnly 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Print"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton chkForPostOnly 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Preview"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Image Image2 
            Height          =   825
            Left            =   1920
            Picture         =   "frm_Print_Dialog.frx":000C
            Stretch         =   -1  'True
            Top             =   240
            Width           =   825
         End
      End
   End
   Begin VB.Label issues 
      BackStyle       =   0  'Transparent
      Caption         =   "Print Options"
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
      TabIndex        =   1
      Top             =   0
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   0
      Picture         =   "frm_Print_Dialog.frx":0C4E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5160
   End
End
Attribute VB_Name = "frm_Print_Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_blnPrint As Boolean
Private m_blnPrintCancelled As Boolean
Private m_DateInclusive As Boolean
Private m_dtDateStart As Date
Private m_dtDateEnd As Date
Private strDateStart As String
Private strDateEnd As String

Public Property Get dtDateStart() As Date
  dtDateStart = m_dtDateStart
End Property

Public Property Let dtDateStart(ByVal vNewValue As Date)
  m_dtDateStart = vNewValue
End Property

Public Property Get dtDateEnd() As Date
  dtDateEnd = m_dtDateEnd
End Property

Public Property Let dtDateEnd(ByVal vNewValue As Date)
  m_dtDateEnd = vNewValue
End Property

'Assign action from frm_explorer
Public Property Get blnPrintCancelled() As Boolean
  blnPrintCancelled = m_blnPrintCancelled
End Property

Public Property Let blnPrintCancelled(ByVal vNewValue As Boolean)
  m_blnPrintCancelled = vNewValue
End Property

'Assign action from frm_explorer
Public Property Get blnPrint() As Boolean
  blnPrint = m_blnPrint
End Property

Public Property Let blnPrint(ByVal vNewValue As Boolean)
  m_blnPrint = vNewValue
End Property

'Assign action from frm_explorer
Public Property Get DateInclusive() As Boolean
  DateInclusive = m_DateInclusive
End Property

Public Property Let DateInclusive(ByVal vNewValue As Boolean)
  m_DateInclusive = vNewValue
End Property

Private Sub chkFields_Click()
Dim blnEnable As Boolean
  If chkFields.Value = vbChecked Then
    blnEnable = True
  Else
    blnEnable = False
  End If
  With DTPickerDate(0)
    .Enabled = blnEnable
  End With
  With DTPickerTime(0)
    .Enabled = blnEnable
  End With
  With DTPickerDate(1)
    .Enabled = blnEnable
  End With
  With DTPickerTime(1)
    .Enabled = blnEnable
  End With
End Sub

Private Sub chkForPostOnly_Click(Index As Integer)
  Select Case Index
    Case 0
      m_blnPrint = False
    Case 1
      m_blnPrint = True
  End Select
End Sub

Private Sub cmdCancel_Click()
  blnPrintCancelled = True
  Me.Visible = False
End Sub

Private Sub cmdClose_Click()
On Error GoTo ET
  blnPrintCancelled = False
  If m_DateInclusive Then
    If chkFields.Value = vbChecked Then
      strDateStart = DTPickerDate(0).Value & " " & DTPickerTime(0).Value
      m_dtDateStart = Format(CDate(strDateStart), "General Date")
      strDateEnd = DTPickerDate(1).Value & " " & DTPickerTime(1).Value
      m_dtDateEnd = Format(CDate(strDateEnd), "General Date")
    Else
      m_DateInclusive = False
    End If
  Else
  m_DateInclusive = False
  End If
  Me.Visible = False
  Exit Sub
ET:
  
End Sub

Private Sub Form_Load()
On Error Resume Next
  m_blnPrint = False
  With DTPickerDate(0)
    .Value = Format(Now - 1, "Short Date")
    .Enabled = False
  End With
  With DTPickerTime(0)
    .Value = Format(Now, "Long Time")
    .Enabled = False
  End With
  With DTPickerDate(1)
    .Value = Format(Now, "Short Date")
    .Enabled = False
  End With
  With DTPickerTime(1)
    .Value = Format(Now, "Long Time")
    .Enabled = False
  End With
  If m_DateInclusive Then
    FrameArray(6).Enabled = True
    chkFields.Value = vbUnchecked
  Else
    FrameArray(6).Enabled = False
  End If
End Sub
