VERSION 5.00
Begin VB.Form frmInfoCenter 
   Caption         =   "Client Information"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   Icon            =   "frmInfoCenter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAddStation 
      Caption         =   "<<"
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
      Left            =   4200
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdRemoveStation 
      Caption         =   ">>"
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
      Left            =   4200
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save(保存)"
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
      Left            =   6480
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox lstNot 
      Height          =   4545
      Left            =   4680
      TabIndex        =   6
      Top             =   1800
      Width           =   3855
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Query(查找)"
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
      Left            =   4680
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox lstResult 
      Height          =   4545
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   3615
   End
   Begin VB.ComboBox cbxProject 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Stations Not Assigned(未分配的站别):"
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
      Left            =   4680
      TabIndex        =   7
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Stations Assigned(已分配的站别):"
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
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   3015
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
      Left            =   480
      TabIndex        =   2
      Top             =   840
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
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "frmInfoCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsProject, lsConStrProject As String
    
Private Sub cbxProject_Click()
    Call cmdQuery_Click
End Sub

Private Sub cmdAddStation_Click()
    If lstNot.ListIndex < 0 Then
        MsgBox "Please select Station to be added!", vbExclamation
        
        Exit Sub
    End If
    
    lstResult.AddItem lstNot.List(lstNot.ListIndex)
    lstResult.ItemData(lstResult.ListCount - 1) = lstNot.ItemData(lstNot.ListIndex)
    
    lstNot.RemoveItem (lstNot.ListIndex)
End Sub


Private Sub cmdRemoveStation_Click()
    If lstResult.ListIndex < 0 Then
        MsgBox "Please select Station to be removed!", vbExclamation
        Exit Sub
    End If
    
    lstNot.AddItem lstResult.List(lstResult.ListIndex)
    lstNot.ItemData(lstResult.ListCount - 1) = lstResult.ItemData(lstResult.ListIndex)
    
    lstResult.RemoveItem (lstResult.ListIndex)
End Sub

Private Sub cmdQuery_Click()
'    Dim lsProject, lsConStrProject As String
    Dim rstTemp As New ADODB.Recordset
    
    lstResult.Clear
    lstNot.Clear
    
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
    
    If DB_Project_RecordSet_WithConStr("exec usp_GetClientInfo " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & ",'" & gClientName & "','" & gEmpID & "','Y'", rstTemp, lsConStrProject) Then

        While Not rstTemp.EOF
            lstResult.AddItem (IIf(IsNull(rstTemp.Fields("StationName")), "-", rstTemp.Fields("StationName")))
            lstResult.ItemData(lstResult.ListCount - 1) = rstTemp.Fields("StationID")
            
            rstTemp.MoveNext
        Wend
        rstTemp.Close
    Else
        MsgBox "Error Code 3003: Connect to project database failed(with str), usp_GetClientInfo!!!"
    End If

    If DB_Project_RecordSet_WithConStr("exec usp_GetClientInfo " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & ",'" & gClientName & "','" & gEmpID & "','N'", rstTemp, lsConStrProject) Then

        While Not rstTemp.EOF
            lstNot.AddItem (IIf(IsNull(rstTemp.Fields("StationName")), "-", rstTemp.Fields("StationName")))
            lstNot.ItemData(lstNot.ListCount - 1) = rstTemp.Fields("StationID")
            rstTemp.MoveNext
        Wend
        rstTemp.Close
    Else
        MsgBox "Error Code 3003: Connect to project database failed(with str), usp_GetClientInfo!!!"
    End If
    
    If DB_Project_RecordSet_WithConStr("exec usp_ClientInfo_Validate " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & ",'" & gClientName & "','" & gEmpID & "'", rstTemp, lsConStrProject) Then

        If Not rstTemp.EOF Then
            If rstTemp.Fields(0) = "Y" Then
                cmdSave.Visible = True
                cmdAddStation.Visible = True
                cmdRemoveStation.Visible = True
            End If
        End If
        
        If rstTemp.State = adStateOpen Then
            rstTemp.Close
        End If
    Else
        MsgBox "Error Code 3003: Connect to project database failed(with str), usp_ClientInfo_Validate!!!"
    End If
End Sub

Private Sub cmdSave_Click()
    Dim i As Integer
    
    If Not DB_Project_ExecSQL_WithConStr("exec Client_StationDelete " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & ",'" & gClientName & "'", lsConStrProject) Then
        MsgBox "Error Code 3013: Connect to project database failed(with str), Client_StationDelete!!!"
    End If
    
    If lstResult.ListCount > 0 Then
        For i = 0 To lstResult.ListCount - 1
                If Not DB_Project_ExecSQL_WithConStr("exec Client_StationUpdate " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & ",'" & gClientName & "'," & lstResult.ItemData(i) & ",'" & gEmpID & "'", lsConStrProject) Then
                    MsgBox "Error Code 3014: Connect to project database failed(with str), Client_StationUpdate!!!"
                End If
        Next i
    End If

End Sub

Private Sub Form_Load()
lblClient.Caption = "Client: " & gClientName & "      User: " & gEmpID & vbTab & "      Line: " & gLineName

    Dim rstTemp As New ADODB.Recordset
    
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
    Else
        'Me.MousePointer = vbNormal
        MsgBox "Error Code 3002: Connect to Master database failed, usp_GetProject_ByUser!!!"
        Unload Me
    End If
End Sub

