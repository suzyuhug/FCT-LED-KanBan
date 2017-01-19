VERSION 5.00
Begin VB.Form frmSchedule 
   Caption         =   "Schedule"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   Icon            =   "frmSchedule.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   11085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCurrent 
      Caption         =   "Current Time Schedule"
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
      Left            =   1320
      TabIndex        =   6
      Top             =   1320
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.ListBox lstResult 
      Height          =   4935
      Left            =   480
      TabIndex        =   5
      Top             =   2160
      Width           =   10215
   End
   Begin VB.ComboBox cbxProject 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Query"
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
      TabIndex        =   2
      Top             =   840
      Width           =   1335
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
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Schedule for crrent line:"
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
      TabIndex        =   1
      Top             =   1800
      Width           =   3615
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
Attribute VB_Name = "frmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsProject, lsConStrProject As String
Dim strSQL As String
Dim strTemp As String

Private Sub cbxProject_Click()
    Call cmdQuery_Click
End Sub

Private Sub cmdQuery_Click()
    Dim rstTemp As New ADODB.Recordset
    Dim i As Integer
    
    lstResult.Clear
    
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
    
    strSQL = "exec usp_Get_Schedule_ByLine " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & "," & gLineID & ","
    If chkCurrent.Value = Checked Then
        strSQL = strSQL & "1,"
    Else
        strSQL = strSQL & "0,"
    End If
    strSQL = strSQL & "'" & gClientName & "','" & gEmpID & "'"
    
    If DB_Project_RecordSet_WithConStr(strSQL, rstTemp, lsConStrProject) Then
        strTemp = ""
        For i = 0 To rstTemp.Fields.Count - 1

            strTemp = strTemp & rstTemp.Fields(i).Name & vbTab
            
'            If rstTemp.Fields(i).ActualSize > Len(rstTemp.Fields(i).Name) Then
'                'If rstTemp.Fields(i).ActualSize > 10 Then
'                    strTemp = strTemp & String(rstTemp.Fields(i).ActualSize - Len(rstTemp.Fields(i).Name), "-") & vbTab
'                'Else
'                '    strTemp = strTemp & String(rstTemp.Fields(i).ActualSize - Len(rstTemp.Fields(i).Name) + 5, " ")
'                'End If
'            End If
            'strTemp = strTemp & vbTab

        Next i
        lstResult.AddItem strTemp
        
        While Not rstTemp.EOF
            'lstResult.Columns = rstTemp.Fields.Count
            'Field1   &   "|"   Field2   &   "|"   &   Field3
            strTemp = ""
            For i = 0 To rstTemp.Fields.Count - 1
                'lstResult.List(0, i) = rstTemp.Fields(i)
                strTemp = strTemp & rstTemp.Fields(i) '& vbTab
'                If rstTemp.Fields(i).ActualSize > Len(rstTemp.Fields(i)) Then
'                'If rstTemp.Fields(i).ActualSize > 10 Then
'                    strTemp = strTemp & String(rstTemp.Fields(i).ActualSize - Len(rstTemp.Fields(i)), " ")
'                End If
                strTemp = strTemp & vbTab
            Next i
            
            strTemp = Left(strTemp, Len(strTemp) - 1)
            lstResult.AddItem strTemp
            rstTemp.MoveNext
        Wend
        rstTemp.Close
    Else
        MsgBox "Error Code 3003: Connect to project database failed(with str), usp_GetClientInfo!!!"
    End If

End Sub

Private Sub Form_Load()
lblClient.Caption = "Client: " & gClientName & "      User: " & gEmpID & vbTab & "      Line: " & gLineName

    Dim rstTemp As New ADODB.Recordset
    
    cbxProject.Clear
    
    'gConStrProject = "Driver={SQL Server};server=suz-spcs-01;database=Northwind;uid=sa;pwd=spcs01"
    
    'rstTemp.Open "select distinct project from SIC", gConStrProject, adOpenDynamic
    If DB_Master_RecordSet("exec usp_GetProject_ByUser '" & gEmpID & "'", rstTemp) Then
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
