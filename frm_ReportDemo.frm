VERSION 5.00
Begin VB.Form frm_ReportDemo 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MS Access Report Print/Preview Demo"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frm_ReportDemo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Dbl_Click The Report to Print/Preview"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.ListBox lstReports 
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
         ForeColor       =   &H00FF0000&
         Height          =   4110
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5775
      End
   End
End
Attribute VB_Name = "frm_ReportDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_UserID As String 'module level user name
Private m_UserPassword As String 'module level user password
Private m_MDBPath As String 'module level .mdb path
Private m_MDWPath As String 'module level .mdw path
Private m_MSAccessPath As String 'module level msaccess.exe path

Private m_cnn As ADODB.Connection 'ado connection
Private m_rst As ADODB.Recordset 'ado recordset
Private lngAccessID As Long 'MSAccess ID upon launching or shelling out

Public Property Get UserID() As String
  UserID = m_UserID
End Property

Public Property Let UserID(ByVal vNewValue As String)
  m_UserID = vNewValue
End Property

Public Property Get UserPassword() As String
  UserPassword = m_UserPassword
End Property

Public Property Let UserPassword(ByVal vNewValue As String)
  m_UserPassword = vNewValue
End Property

Public Property Get MDBPath() As String
  MDBPath = m_MDBPath
End Property

Public Property Let MDBPath(ByVal vNewValue As String)
  m_MDBPath = vNewValue
End Property

Public Property Get MDWPath() As String
  MDWPath = m_MDWPath
End Property

Public Property Let MDWPath(ByVal vNewValue As String)
  m_MDWPath = vNewValue
End Property

Public Property Get MSAccessPath() As String
  MSAccessPath = m_MSAccessPath
End Property

Public Property Let MSAccessPath(ByVal vNewValue As String)
  m_MSAccessPath = vNewValue
End Property

Public Property Get AdoConnection() As ADODB.Connection
  AdoConnection = m_cnn
End Property

Public Property Set AdoConnection(ByVal vNewValue As ADODB.Connection)
  Set m_cnn = vNewValue
End Property

Private Sub Form_Load()
On Error GoTo ET

Dim blnPathFound As Boolean
Dim intCounter As Integer
   
  'Get List Of Reports From tbl_Report_Listing
  Set m_rst = New ADODB.Recordset
  Set m_rst = m_cnn.Execute("Select * From tbl_Report_Listing")
  If Not (m_rst.BOF And m_rst.EOF) Then
    With lstReports
      Do While Not m_rst.EOF
        .AddItem m_rst("Report_Name")
        'Thanks to Jos Groen for pointing this out
        .ItemData(.NewIndex) = Val(m_rst("ID"))
        m_rst.MoveNext
      Loop
      'Select first item of the list
      .Text = .List(0)
    End With
  End If
  Exit Sub
ET:
  MsgBox Err.Description, vbCritical, "Error!"
  Err = False
End Sub

'This procedure prints/preview the ms access report
'Arguments:
'strReportName - The name of the report
'strFilter - filter of the recordsource of the report
'DateInclusive - Determines if the report supports date interval viewing
Private Sub PrintAccessReport_Protected(ByVal strReportName As String, ByVal strFilter As String, ByVal DateInclusive As Boolean)

On Error Resume Next

Dim appAccess As Access.Application

Dim strShellCommand As String
Dim blnPrint As Boolean
Dim blnCancelled As Boolean
Dim blnDateInclusive As Boolean
Dim blnSuccess As Boolean
Dim strDateCriteria As String
Dim dteDateStart As Date
Dim dteDateEnd As Date
  
  'Get Instance of Access, if no instance, create one
  Set appAccess = GetObject(, "Access.Application")
  If Err <> 0 Then
    Err = False
    On Error GoTo ET
    'Create new instance of access by using the shell command
    strShellCommand = """" & MSAccessPath & """" & " " & _
      """" & MDBPath & """" & _
      " /wrkgrp " & """" & MDWPath & """" & _
      " /user " & """" & UserID & """" & _
      " /pwd " & """" & UserPassword & """"
    'Activate application and record the ID for Activating the application later on
    lngAccessID = Shell(strShellCommand, vbNormalFocus)
    'Set Flag to false so that the newly created instance can be captured later
    blnSuccess = False
    'Simulate an ALT TAB to return to our application
    SendKeys "%TAB"
  Else
    blnSuccess = True
  End If
  
  On Error GoTo ET
  'Show Print Dialog Options
  'This dialog was placed here to give ms access enough time to load, otherwise
  'An error occurs
  With frm_Print_Dialog
    .DateInclusive = DateInclusive
    .Show vbModal
    blnPrint = .blnPrint
    blnCancelled = .blnPrintCancelled
    blnDateInclusive = .DateInclusive
    If blnDateInclusive Then
      dteDateStart = .dtDateStart
      dteDateEnd = .dtDateEnd
      strDateCriteria = "[Date] Between #" & dteDateStart & "# And #" & dteDateEnd & "#"
    End If
  End With
  Unload frm_Print_Dialog
  
  'If a new instance was created, capture the newly created access instance
  If blnSuccess = False Then
    Set appAccess = GetObject(, "Access.Application")
  End If
  
  
  If Not blnCancelled Then 'Cancel button of the frm_Print_Dialog
    'Open Reports
    With appAccess
      If Not blnPrint Then
        If blnDateInclusive Then
          If strFilter <> "" Then
            'Construct filter
            strDateCriteria = strFilter & " AND (" & strDateCriteria & ")"
          End If
          'Opens report with date interval viewing in preview mode
          .DoCmd.OpenReport strReportName, acViewPreview, , strDateCriteria
        Else
          'Opens report without date interval viewing in preview mode
          .DoCmd.OpenReport strReportName, acViewPreview, , strFilter
        End If
      Else
        If blnDateInclusive Then
          If strFilter <> "" Then
            'Construct Filter
            strDateCriteria = strFilter & " AND (" & strDateCriteria & ")"
          End If
          'Opens report with date interval viewing in print mode
          .DoCmd.OpenReport strReportName, acViewNormal, , strDateCriteria
        Else
          'Opens report without date interval viewing in print mode
          .DoCmd.OpenReport strReportName, acViewNormal, , strFilter
        End If
      End If
      .DoCmd.Maximize
    End With
    On Error Resume Next
    'Activate Access
    If lngAccessID <> 0 Then
      AppActivate lngAccessID
    Else
      AppActivate "Microsoft Access"
    End If
  End If
  'Clean up
  Set appAccess = Nothing
  Exit Sub
ET:
  MsgBox Err.Description, vbCritical, "Error!"
  Set appAccess = Nothing
  Err = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set m_rst = Nothing
  m_cnn.Close
End Sub

Private Sub lstReports_DblClick()
On Error GoTo ET
Dim bln As Boolean
  If lstReports.Text <> "" Then
    'Filter the recordset to determine if the report has the date interval option checked
    m_rst.Filter = "ID = " & Val(lstReports.ListIndex) + 1
    bln = CBool(m_rst("Has_Filter"))
    PrintAccessReport_Protected lstReports.Text, "", bln
  End If
  Exit Sub
ET:
  MsgBox Err.Description, vbCritical, "Error!"
  Err = False
End Sub
