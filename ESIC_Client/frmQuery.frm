VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmQuery 
   Caption         =   "Query"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11055
   Icon            =   "frmQuery.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6495
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   10335
      ExtentX         =   18230
      ExtentY         =   11456
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.ComboBox cbxReport 
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton cmdGetReport 
      Caption         =   "Get "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox cbxProject 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Task:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblClient 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   7695
   End
   Begin VB.Label Label1 
      Caption         =   "Project:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsProject, lsConStrProject As String
Dim lssql As String

Private Sub cbxProject_Click()
    Dim rstTemp As New ADODB.Recordset
    
    cbxReport.Clear
    
    lsProject = cbxProject.Text
    
    If IsNull(lsProject) Or lsProject = "" Then
        MsgBox "Please select project!", vbExclamation
        
        Exit Sub
    End If
    
    If DB_Master_RecordSet("exec usp_GetProject_Info '" & lsProject & "'", rstTemp) Then
        If Not rstTemp.EOF Then
            lsConStrProject = IIf(IsNull(rstTemp.Fields("ProjectDBConString")), "", Trim(rstTemp.Fields("ProjectDBConString")))
            
        Else
            lsConStrProject = ""
        End If
    Else
        'Me.MousePointer = vbNormal
        MsgBox "Error Code 3003: Connect to Master database failed, usp_GetProject_Info!!!"
        Unload Me
    End If
    
    If rstTemp.State = adStateOpen Then
        rstTemp.Close
    End If
            
    lssql = "exec usp_ClientApp_Report " & cbxProject.ItemData(cbxProject.ListIndex) & ",'" & gClientName & "','" & gAppName & "','" & gAppVersion & "','" & gEmpID & "','" & App.Path & "','" & App.EXEName & "'"
    
    If DB_Project_RecordSet_WithConStr(lssql, rstTemp, lsConStrProject) Then
        While Not rstTemp.EOF
            cbxReport.AddItem (rstTemp.Fields("ReportName"))
            cbxReport.ItemData(cbxReport.ListCount - 1) = rstTemp.Fields("ReportID")
            
            rstTemp.MoveNext
        Wend
        
    Else
        'Me.MousePointer = vbNormal
        MsgBox "Error Code 3015: Connect to Project database failed, usp_ClientApp_Report!!!"
    End If
    
    If rstTemp.State = adStateOpen Then
        rstTemp.Close
    End If
    
End Sub

Private Sub cmdGetReport_Click()
On Error GoTo Local_Error

    Dim lsReportURL As String
    
           
    If IsNull(cbxProject.Text) Or cbxProject.Text = "" Then
        MsgBox "Please select project!", vbExclamation
        
        Exit Sub
    End If
    
    If IsNull(cbxReport.Text) Or cbxReport.Text = "" Then
        MsgBox "Please select report!", vbExclamation
        
        Exit Sub
    End If
    
    Dim rstTemp As New ADODB.Recordset
    
    lssql = "exec usp_ClientApp_ReportURL " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & "," & CStr(cbxReport.ItemData(cbxReport.ListIndex)) & "," & gLineID & ",'" & gClientName & "','" & gAppName & "','" & gAppVersion & "','" & gEmpID & "','" & App.Path & "','" & App.EXEName & "'"
    
    If DB_Project_RecordSet_WithConStr(lssql, rstTemp, lsConStrProject) Then
        If Not rstTemp.EOF Then
            lsReportURL = IIf(IsNull(rstTemp.Fields("ReportURL")), "", rstTemp.Fields("ReportURL"))
            
            If rstTemp.State = adStateOpen Then
                rstTemp.Close
            End If
            
            WebBrowser1.Stop
            
            WebBrowser1.Navigate (lsReportURL)
        End If
        
    Else
        'Me.MousePointer = vbNormal
        MsgBox "Error Code 3016: Connect to Project database failed, usp_ClientApp_Report!!!"
    End If

    Exit Sub
    
Local_Error:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
    lblClient.Caption = "Client: " & gClientName & "      User: " & gEmpID & vbTab & "      Line: " & gLineName
   
    Dim rstTemp As New ADODB.Recordset
    
    WebBrowser1.Navigate (gClientReportDefaultURL)
    
    cbxProject.Clear
    
    'gConStrProject = "Driver={SQL Server};server=suz-spcs-01;database=Northwind;uid=sa;pwd=spcs01"
    
    'rstTemp.Open "select distinct project from SIC", gConStrProject, adOpenDynamic
    If DB_Master_RecordSet("exec usp_GetProject_ByUser_Client '" & gEmpID & "'", rstTemp) Then
        While Not rstTemp.EOF
            cbxProject.AddItem (rstTemp.Fields("ProjectCode"))
            cbxProject.ItemData(cbxProject.ListCount - 1) = rstTemp.Fields("ProjectID")
            
            rstTemp.MoveNext
        Wend
    
        rstTemp.Close
   End If
        'Me.MousePointer = vbNormal
'        MsgBox "Error Code 3002: Connect to Master database failed, usp_GetProject_ByUser!!!"
'        Unload Me
'    End If
'    If DB_Master_RecordSet("exec usp_GetClientVersion", rstTemp) Then
'        If Not rstTemp.EOF Then
'
'             Label3.Caption = rstTemp.Fields("Appversion")
'        End If
'    Else
'        'Me.MousePointer = vbNormal
'        MsgBox "Error Code 3002: Connect to Master database failed, usp_GetClientVersion!!!"
'        Unload Me
'    End If
        
    

End Sub

Private Sub Form_Resize()
    WebBrowser1.Width = 0.92 * Me.Width
    If Me.Height > 1920 Then
        WebBrowser1.Height = 0.9 * (Me.Height - 1920)
    End If
End Sub
