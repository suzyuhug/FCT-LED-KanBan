VERSION 5.00
Begin VB.Form frmESICSelect 
   Caption         =   "Select ESIC"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   ControlBox      =   0   'False
   Icon            =   "frmESICSelect.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   7725
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   7080
      Top             =   2640
   End
   Begin VB.CommandButton Advance 
      Caption         =   "Advance(高级)..."
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
      Left            =   5880
      TabIndex        =   27
      Top             =   960
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   1200
      TabIndex        =   25
      Top             =   2760
      Visible         =   0   'False
      Width           =   4335
      Begin VB.ComboBox CboGroup 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.CheckBox ChkAllModel 
      Caption         =   "By Group(根据组)"
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
      TabIndex        =   24
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   240
      TabIndex        =   18
      Top             =   5160
      Width           =   7215
      Begin VB.ComboBox cbxStation 
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Top             =   1200
         Width           =   4335
      End
      Begin VB.ComboBox cbxModel 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   4335
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
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
         Left            =   5640
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
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
         Left            =   5640
         TabIndex        =   12
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox cbxSide 
         Height          =   315
         Left            =   960
         TabIndex        =   7
         Top             =   720
         Width           =   4335
      End
      Begin VB.CheckBox chkHistorySIC 
         Caption         =   "History SIC"
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
         Left            =   5640
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtHisSICReason 
         Height          =   615
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1680
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label Label3 
         Caption         =   "Station:"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Model:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Side:"
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
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblHisSICReason 
         Caption         =   "History SIC Reason:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.ComboBox cbxSICCategory 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   1200
      TabIndex        =   13
      Top             =   3840
      Visible         =   0   'False
      Width           =   4335
      Begin VB.ComboBox cbxExternalType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   3375
      End
      Begin VB.CommandButton Get 
         Caption         =   "Get"
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
         Left            =   3600
         TabIndex        =   17
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtOrder 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   3375
      End
   End
   Begin VB.CheckBox ChkGetData 
      Caption         =   "External Data(外部条件)"
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
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txtLine 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   4215
   End
   Begin VB.ComboBox cbxProject 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Label Label7 
      Caption         =   "Category(种类):"
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
      Left            =   0
      TabIndex        =   23
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Line(线别):"
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
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Select ESIC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Project(项目):"
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
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "frmESICSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Advance_Click()
    Unload Me
    frmESICGet.Show vbModal
End Sub

Private Sub CboGroup_Change()
     Dim rstTemp As New ADODB.Recordset
    
    If DB_Project_RecordSet("exec usp_GetSICModel_ByProjectCateGoryID " + CStr(cbxProject.ItemData(cbxProject.ListIndex)) + "," + gLineID + "," + CStr(cbxSICCategory.ItemData(cbxSICCategory.ListIndex)) & ",'" & gEmpID & "',1,'" & Trim(CboGroup.ItemData(CboGroup.ListIndex)) & "','" & gClientName & "'", rstTemp) Then
            cbxModel.Clear
            While Not rstTemp.EOF
                cbxModel.AddItem (rstTemp.Fields("ModelName"))
                cbxModel.ItemData(cbxModel.ListCount - 1) = rstTemp.Fields("ModelID")
                'cbxmodel.ItemData(
                rstTemp.MoveNext
                
            Wend
            If cbxModel.ListCount = 1 Then
                cbxModel.ListIndex = 0
            End If
            rstTemp.Close
        Else
            MsgBox "Error Code 3003: Connect to project database failed, usp_GetSICModel_ByProjectCateGoryID!!!"
        End If
End Sub

Private Sub CboGroup_Click()
   CboGroup_Change
End Sub

Private Sub cbxModel_Click()
   ' Call cbxModel_Change
   If cbxStation.Text = "" Then
        If gSelect = 1 Then
            gSelect = 0
        End If
    End If
    
    If gSelect = 1 Then
        Exit Sub
    End If
    
    Dim rstTemp As New ADODB.Recordset
    
    'If DB_Project_RecordSet("exec usp_GetSICSide_ByProjectModelID " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & "," & CStr(cbxModel.ItemData(cbxModel.ListIndex)) & "," & gLineID, rsttemp) Then
     If DB_Project_RecordSet("exec usp_GetSICSide_ByProjectModelID " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & "," & CStr(cbxModel.ItemData(cbxModel.ListIndex)) & "," & gLineID & "," & CStr(cbxSICCategory.ItemData(cbxSICCategory.ListIndex)) & ",'" & gEmpID & "'", rstTemp) Then
        cbxSide.Clear
        While Not rstTemp.EOF
            cbxSide.AddItem (rstTemp.Fields("SideName"))
            cbxSide.ItemData(cbxSide.ListCount - 1) = rstTemp.Fields("SideID")
            rstTemp.MoveNext
        Wend
         If cbxSide.ListCount = 1 Then
                cbxSide.ListIndex = 0
            End If
        rstTemp.Close
    Else
        MsgBox "Error Code 3003: Connect to project database failed, usp_GetSICSide_ByProjectModelID!!!"
    End If
End Sub

Private Sub cbxProject_Change()
    Dim rstTemp As New ADODB.Recordset
    
    'gConStrProject = "Driver={SQL Server};server=suz-spcs-01;database=Northwind;uid=sa;pwd=spcs01"
    
    gProject = cbxProject.Text
    cbxModel.Clear
    cbxSide.Clear
    cbxStation.Clear
    cbxSICCategory.Clear
    ChkAllModel.Enabled = True
    
    If DB_Master_RecordSet("exec usp_GetProject_Info '" & gProject & "'", rstTemp) Then
        If Not rstTemp.EOF Then
            gConStrProject = IIf(IsNull(rstTemp.Fields("ProjectDBConString")), "", Trim(rstTemp.Fields("ProjectDBConString")))
            gProjectWebURL = IIf(IsNull(rstTemp.Fields("ProjectWebURL")), "", Trim(rstTemp.Fields("ProjectWebURL")))
            gProjectDefaultURL = IIf(IsNull(rstTemp.Fields("ProjectDefaultURL")), "", Trim(rstTemp.Fields("ProjectDefaultURL")))
            If Right(gProjectWebURL, 1) <> "/" Then
                gProjectWebURL = gProjectWebURL + "/"
            End If
        Else
            gConStrProject = ""
        End If
    
        If rstTemp.State = adStateOpen Then
            rstTemp.Close
        End If
        
        If IsNull(gConStrProject) Or gConStrProject = "" Then
            MsgBox "Error Code 4001: Project NOT configured in system, usp_GetProject_Info!!!"
            Unload Me
        End If
        
        If DB_Project_RecordSet("exec usp_GetESICRoleByProjectIDUser " + CStr(cbxProject.ItemData(cbxProject.ListIndex)) + ",'" + gEmpID + "'", rstTemp) Then
            gSICViewer = rstTemp.Fields("SICViewer")
            gSICViewerHistory = rstTemp.Fields("SICViewerHistory")
            
            If gSICViewer < 1 Then
                MsgBox "You are NOT authorized to view SIC for this project." & vbCr & "没有权限查看此项目的SIC.", vbExclamation
                
                Exit Sub
            End If
            
            If gSICViewerHistory = 1 Then
                chkHistorySIC.Visible = True
            End If

        Else
            MsgBox "Error Code 3010: Connect to project database failed, usp_GetESICRoleByProjectIDUser!!!"
        End If
        
        If rstTemp.State = adStateOpen Then
            rstTemp.Close
        End If
          
    Dim SQLTemp As String
    Set rstTemp = New ADODB.Recordset

     SQLTemp = "exec usp_Check_Operator '" & gEmpID & "'"
     If DB_Project_RecordSet(SQLTemp, rstTemp) Then
         SQLTemp = "exec usp_Check_Operator '" & gEmpID & "'"
    End If
                   
                    
     
            
         '--Get ESIC Category
        cbxSICCategory.Clear
        Set rstTemp = New ADODB.Recordset
        If DB_Project_RecordSet("exec usp_GetSICCategory_ByProjectID " + CStr(cbxProject.ItemData(cbxProject.ListIndex)) & ",'" & gEmpID & "'", rstTemp) Then
            cbxSICCategory.Clear
            While Not rstTemp.EOF

                cbxSICCategory.AddItem (rstTemp.Fields("Strvalues"))
                cbxSICCategory.ItemData(cbxSICCategory.ListCount - 1) = rstTemp.Fields("ID")
    
                rstTemp.MoveNext
                
            Wend
            If cbxSICCategory.ListCount = 0 Then
                MsgBox "Error Code 3004: No ESIC you can open"
                Exit Sub
            End If
            
             cbxSICCategory.ListIndex = 0
             
            rstTemp.Close
        Else
            MsgBox "Error Code 3003: Connect to project database failed, usp_GetSICCategory_ByProjectID!!!"
        End If
        
        
       '--Get ESIC Category
       '--display external data
       If DB_Project_RecordSet("exec usp_Check_ExternalDate_Exists " + CStr(cbxProject.ItemData(cbxProject.ListIndex)), rstTemp) Then
        
            If Not rstTemp.EOF Then
                ChkGetData.Visible = True
                
            Else
                ChkGetData.Visible = False
                
            End If
           
            rstTemp.Close
        Else
            MsgBox "Error Code 3003: Connect to project database failed, usp_GetModel_ByProject!!!"
        End If
        
       '--display external data
        
'        If DB_Project_RecordSet("exec usp_GetSICModel_ByProjectID " + CStr(cbxProject.ItemData(cbxProject.ListIndex)) + "," + gLineID + ",'" & gEmpID & "'", rstTemp) Then
'            cbxModel.Clear
'            While Not rstTemp.EOF
'                cbxModel.AddItem (rstTemp.Fields("ModelName"))
'                cbxModel.ItemData(cbxModel.ListCount - 1) = rstTemp.Fields("ModelID")
'                'cbxmodel.ItemData(
'                rstTemp.MoveNext
'
'            Wend
'            If cbxModel.ListCount = 1 Then
'                cbxModel.ListIndex = 0
'            End If
'            rstTemp.Close
'        Else
'            MsgBox "Error Code 3003: Connect to project database failed, usp_GetModel_ByProject!!!"
'        End If
       
       
       '--Get External Type
        cbxExternalType.Clear
        
        If DB_Project_RecordSet("exec usp_GetSICExternalType_ByProjectID " + CStr(cbxProject.ItemData(cbxProject.ListIndex)), rstTemp) Then
            cbxExternalType.Clear
            While Not rstTemp.EOF

                cbxExternalType.AddItem (rstTemp.Fields("ExternalData"))
                cbxExternalType.ItemData(cbxExternalType.ListCount - 1) = rstTemp.Fields("ExternalID")
    
                rstTemp.MoveNext
                
            Wend
             cbxExternalType.ListIndex = -1
             
            rstTemp.Close
        Else
            MsgBox "Error Code 3003: Connect to project database failed, usp_GetSICCategory_ByProjectID!!!"
        End If
        
        
       '--Get External Type
        
'        If DB_Project_RecordSet("exec usp_GetSICStation_ByProjectID '" + gClientName + "'," + CStr(cbxProject.ItemData(cbxProject.ListIndex)) + "," + gLineID, rstTemp) Then
'            cbxStation.Clear
'            While Not rstTemp.EOF
'                CbxSICCategory.AddItem (rstTemp.Fields("Value"))
'                cbxStation.ItemData(cbxStation.ListCount - 1) = rstTemp.Fields("ID")
'
'                rstTemp.MoveNext
'            Wend
'
'            rstTemp.Close
'        Else
'            MsgBox "Error Code 3003: Connect to project database failed, usp_GetSICStation_ByProjectID!!!"
'        End If
        
        'Get Side
        
'        If DB_Project_RecordSet("exec gCode_Get 'ESICSide'", rsttemp) Then
'            cbxSide.Clear
'            While Not rsttemp.EOF
'                cbxSide.AddItem (rsttemp.Fields("strValues"))
'                cbxSide.ItemData(cbxSide.ListCount - 1) = rsttemp.Fields("ID")
'
'                rsttemp.MoveNext
'            Wend
'             If cbxSide.ListCount = 1 Then
'                cbxSide.ListIndex = 0
'            End If
'            rsttemp.Close
'        Else
'            MsgBox "Error Code 3003: Connect to project database failed, gCode_Get!!!"
'        End If
        
        
        'gRefreshTime
        Dim oldgRefreshTime As Integer
        oldgRefreshTime = gRefreshTime
        
        gSQL = "exec SystemRule_Get " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & ",3"
        
        If DB_Project_RecordSet(gSQL, rstTemp) Then
            If rstTemp.EOF Then
                MsgBox "Error Code 4002: Project rule RefreshTime NOT configured in system!!!"
            Else
                gRefreshTime = CInt(rstTemp.Fields("ProjectValue"))
            End If
            
            If rstTemp.State = adStateOpen Then rstTemp.Close
        Else
            MsgBox "Error Code 3010: Connect to project database failed, SystemRule_Get RefreshTime!!!"
        End If
        If oldgRefreshTime <= gRefreshTime Then
            gRefreshTime = oldgRefreshTime
        End If
        
        'Get CertificateTime for project
        gSQL = "exec SystemRule_Get " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & ",1"
        
        If DB_Project_RecordSet(gSQL, rstTemp) Then
            If rstTemp.EOF Then
                MsgBox "Error Code 4002: Project rule CertificateTime NOT configured in system!!!"
            Else
                gCertificateTime = CInt(rstTemp.Fields("ProjectValue"))
            End If
            
            If rstTemp.State = adStateOpen Then rstTemp.Close
        Else
            MsgBox "Error Code 3010: Connect to project database failed, SystemRule_Get CertificateTime!!!"
        End If
        
        gSQL = "exec SystemRule_Get " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & ",6"
        
        If DB_Project_RecordSet(gSQL, rstTemp) Then
            If rstTemp.EOF Then
                MsgBox "Error Code 4002: Project rule ClientQueryButtonEnable NOT configured in system!!!"
            Else
               gClientQueryButtonEnable = UCase(IIf(IsNull(rstTemp.Fields("ProjectValue")), "NO", rstTemp.Fields("ProjectValue")))
            End If
            
            If rstTemp.State = adStateOpen Then rstTemp.Close
        Else
            MsgBox "Error Code 3010: Connect to project database failed, SystemRule_Get ClientQueryButtonEnable!!!"
        End If
        
    Else
        'Me.MousePointer = vbNormal
        MsgBox "Error Code 3003: Connect to Master database failed, usp_GetProject_Info!!!"
        Unload Me
    End If
End Sub

Private Sub cbxProject_Click()
    Call cbxProject_Change
End Sub

Private Sub cbxSICCategory_Click()
     Dim rstTemp As New ADODB.Recordset
    cbxModel.Clear
    cbxSide.Clear
    cbxStation.Clear
    If (cbxSICCategory.ItemData(cbxSICCategory.ListIndex) = "2") Then
        ChkAllModel.Value = False
        ChkAllModel.Enabled = False
        Label2.Caption = "Title:"
        Label3.Visible = False
        cbxStation.Visible = False
        txtHisSICReason.Visible = False
        lblHisSICReason.Visible = False
        chkHistorySIC.Visible = False
        ChkGetData.Enabled = False
        
        
        
    Else
        If (cbxSICCategory.ItemData(cbxSICCategory.ListIndex) = "4") Then
            
            Label2.Caption = "Model:"
            Label3.Caption = "Title:"
            Label3.Visible = True
            cbxStation.Visible = True
            'txtHisSICReason.Visible = True
            'lblHisSICReason.Visible = True
            'chkHistorySIC.Visible = True
            ChkAllModel.Enabled = True
            ChkGetData.Enabled = True
            
            
            
            
        Else
    
            Label2.Caption = "Model:"
            Label3.Caption = "Station:"
            Label3.Visible = True
            cbxStation.Visible = True
            'txtHisSICReason.Visible = True
            'lblHisSICReason.Visible = True
            'chkHistorySIC.Visible = True
            ChkAllModel.Enabled = True
            ChkGetData.Enabled = True
        End If
        
    End If
     If DB_Project_RecordSet("exec usp_GetSICModel_ByProjectCateGoryID " + CStr(cbxProject.ItemData(cbxProject.ListIndex)) + "," + gLineID + "," + CStr(cbxSICCategory.ItemData(cbxSICCategory.ListIndex)) & ",'" & gEmpID & "',0,0,'" & gClientName & "'", rstTemp) Then
            cbxModel.Clear
            While Not rstTemp.EOF
                cbxModel.AddItem (rstTemp.Fields("ModelName"))
                cbxModel.ItemData(cbxModel.ListCount - 1) = rstTemp.Fields("ModelID")
                'cbxmodel.ItemData(
                rstTemp.MoveNext
                
            Wend
            If cbxModel.ListCount = 1 Then
                cbxModel.ListIndex = 0
            End If
            rstTemp.Close
        Else
            MsgBox "Error Code 3003: Connect to project database failed, usp_GetSICModel_ByProjectCateGoryID!!!"
        End If
        Set rstTemp = New ADODB.Recordset
        
        If cbxStation.ListCount = 0 And cbxModel.ListIndex = -1 Then
            
       
         If DB_Project_RecordSet("exec usp_GetSICStation_ByProjectCateGoryID " + CStr(cbxProject.ItemData(cbxProject.ListIndex)) + "," + gLineID + "," + CStr(cbxSICCategory.ItemData(cbxSICCategory.ListIndex)) & ",'" & gEmpID & "',0,0,'" & gClientName & "'", rstTemp) Then
            cbxStation.Clear
            While Not rstTemp.EOF
                cbxStation.AddItem (rstTemp.Fields("StationName"))
                cbxStation.ItemData(cbxStation.ListCount - 1) = rstTemp.Fields("StationID")
                'cbxmodel.ItemData(
                rstTemp.MoveNext
                
            Wend
            If cbxStation.ListCount = 1 Then
                cbxStation.ListIndex = 0
            End If
            rstTemp.Close
        Else
            MsgBox "Error Code 3003: Connect to project database failed, usp_GetSICStation_ByProjectCateGoryID!!!"
        End If
        End If
        
       
End Sub

Private Sub cbxSide_Change()
    On Error GoTo Errth
    Dim rstTemp As New ADODB.Recordset

    If gSelect = 0 Then
        
        If DB_Project_RecordSet("exec usp_GetSICStation_ByProjectModeSideID '" + gClientName + "'," + CStr(cbxProject.ItemData(cbxProject.ListIndex)) + "," + CStr(cbxModel.ItemData(cbxModel.ListIndex)) + "," + CStr(cbxSide.ItemData(cbxSide.ListIndex)) + "," + gLineID & "," & CStr(cbxSICCategory.ItemData(cbxSICCategory.ListIndex)) & ",'" & gEmpID & "'", rstTemp) Then
            cbxStation.Clear
            While Not rstTemp.EOF
                cbxStation.AddItem (rstTemp.Fields("StationName"))
                cbxStation.ItemData(cbxStation.ListCount - 1) = rstTemp.Fields("StationID")
                'cbxmodel.ItemData(
                rstTemp.MoveNext
            Wend
            If cbxStation.ListCount = 1 Then
                cbxStation.ListIndex = 0
            End If
            rstTemp.Close
        Else
            MsgBox "Error Code 3003: Connect to project database failed, usp_GetSICStation_ByProjectID!!!"
        End If
       
    Else
       

        If DB_Project_RecordSet("exec usp_GetSICModel_ByProjectStationSideID '" + gClientName + "'," + CStr(cbxProject.ItemData(cbxProject.ListIndex)) + "," + CStr(cbxStation.ItemData(cbxStation.ListIndex)) + "," + CStr(cbxSide.ItemData(cbxSide.ListIndex)) + "," + gLineID & "," & CStr(cbxSICCategory.ItemData(cbxSICCategory.ListIndex)) & ",'" & gEmpID & "'", rstTemp) Then
            cbxModel.Clear
            While Not rstTemp.EOF
                cbxModel.AddItem (rstTemp.Fields("ModelName"))
                cbxModel.ItemData(cbxModel.ListCount - 1) = rstTemp.Fields("ModelID")
                'cbxmodel.ItemData(
                rstTemp.MoveNext
            Wend
            If cbxModel.ListCount = 1 Then
                cbxModel.ListIndex = 0
            End If
            rstTemp.Close
        Else
            MsgBox "Error Code 3003: Connect to project database failed, usp_GetSICStation_ByProjectID!!!"
        End If
      
    End If
        Exit Sub
Errth:
    MsgBox "Error Code 3004: Please choose full info,usp_GetSICStation_ByProjectModeSideID", vbInformation
    
End Sub

Private Sub cbxSide_Click()
    Call cbxSide_Change
End Sub

Private Sub cbxStation_Click()
 If cbxModel.Text = "" Then
        If gSelect = 0 Then
            gSelect = 1
        End If
    End If
    
    If gSelect = 0 Then
        Exit Sub
    End If
    

    Dim rstTemp As New ADODB.Recordset
    
     If DB_Project_RecordSet("exec usp_GetSICSide_ByProjectStationID " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & "," & CStr(cbxStation.ItemData(cbxStation.ListIndex)) & "," & gLineID & "," & CStr(cbxSICCategory.ItemData(cbxSICCategory.ListIndex)) & ",'" & gEmpID & "'", rstTemp) Then
        cbxSide.Clear
        While Not rstTemp.EOF
            cbxSide.AddItem (rstTemp.Fields("SideName"))
            cbxSide.ItemData(cbxSide.ListCount - 1) = rstTemp.Fields("SideID")
            rstTemp.MoveNext
        Wend
         If cbxSide.ListCount = 1 Then
                cbxSide.ListIndex = 0
            End If
        rstTemp.Close
    Else
        MsgBox "Error Code 3003: Connect to project database failed, usp_GetSICSide_ByProjectModelID!!!"
    End If
End Sub

Private Sub Check1_Click()
    
End Sub

Private Sub ChkAllModel_Click()
    If cbxProject.Text <> "" Then
        

    
    If ChkAllModel.Value = 1 Then
        
        Label2.Caption = "TITLE:"
        Frame3.Visible = True
        cbxModel.Clear
        cbxStation.Clear
        cbxSide.Clear
        Fix_Position
        
        Dim rstTemp As New ADODB.Recordset
        CboGroup.Clear
        
      If DB_Project_RecordSet("exec usp_GetSICGroup_ByProjectCateGoryID " + CStr(cbxProject.ItemData(cbxProject.ListIndex)) + "," + gLineID + "," + CStr(cbxSICCategory.ItemData(cbxSICCategory.ListIndex)) & ",'" & gEmpID & "',1", rstTemp) Then
            cbxModel.Clear
            While Not rstTemp.EOF
                CboGroup.AddItem (rstTemp.Fields("ModelName"))
                CboGroup.ItemData(CboGroup.ListCount - 1) = rstTemp.Fields("ModelID")
                'cbxmodel.ItemData(
                rstTemp.MoveNext
                
            Wend
            If CboGroup.ListCount = 1 Then
                CboGroup.ListIndex = 0
            End If
            rstTemp.Close
        Else
            MsgBox "Error Code 3003: Connect to project database failed, usp_GetSICModel_ByProjectCateGoryID!!!"
        End If
        
        
       
        
    Else
        Frame3.Visible = False
        Label2.Caption = "MODEL:"
        cbxSICCategory_Click
        Fix_Position
    End If
        End If
    
End Sub
Private Sub Fix_Position()
    If ChkAllModel.Value = 1 Then
        If ChkGetData.Value = 1 Then
            frmESICSelect.Height = 8265
            Frame2.Top = 5160
            ChkGetData.Top = 3480
            Frame1.Top = 3840
            
        Else
            frmESICSelect.Height = 6990
            Frame2.Top = 3960
            ChkGetData.Top = 3480
            Frame1.Top = 3840
        End If
        
    Else
        If ChkGetData.Value = 1 Then
            frmESICSelect.Height = 7530
            Frame2.Top = 4440
            ChkGetData.Top = 2760
            Frame1.Top = 3120
        Else
            frmESICSelect.Height = 6300
            Frame2.Top = 3240
            ChkGetData.Top = 2760
            Frame1.Top = 3120
        End If
    End If
    
End Sub
Private Sub ChkGetData_Click()
    If ChkGetData.Value = 1 Then
        Frame1.Visible = True
      
        cbxExternalType.ListIndex = -1
        txtOrder.Text = ""
        cbxModel.Clear
        cbxStation.Clear
        cbxSide.Clear
        Fix_Position
    Else
        Frame1.Visible = False
        
        Fix_Position
    End If
    
End Sub


Private Sub chkHistorySIC_Click()
    If chkHistorySIC.Value = Checked Then
        lblHisSICReason.Visible = True
        txtHisSICReason.Visible = True
    Else
        lblHisSICReason.Visible = False
        txtHisSICReason.Visible = False
    End If
End Sub

Private Sub cmdCancel_Click()
    gURL = "NO_CHANGE"
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    Dim rstTemp As New ADODB.Recordset
    Dim strURL As String
    Dim intHistorySIC As Integer
    
    If (cbxProject.ListIndex < 0) Then
        MsgBox "Please select Project!", vbExclamation
        
        Exit Sub
    End If
    
    If (cbxModel.ListIndex < 0) Then
        MsgBox "Please select Model!", vbExclamation
        
        Exit Sub
    End If
    
    If (cbxStation.ListIndex < 0) Then
        MsgBox "Please select Station!", vbExclamation
        
        Exit Sub
    End If
    If (cbxSide.ListIndex < 0) Then
        'cbxSide.ListIndex = 2
        MsgBox "Please select Side!", vbExclamation
        
        Exit Sub
    End If

    intHistorySIC = 0
    If chkHistorySIC.Value = Checked Then
        intHistorySIC = 1
        txtHisSICReason.Text = LTrim(RTrim(txtHisSICReason.Text))
        If IsNull(txtHisSICReason.Text) Or txtHisSICReason.Text = "" Then
            MsgBox "Please input reason for opening History SIC!", vbExclamation
            
            Exit Sub
        End If
    End If
    
    If IsNull(txtHisSICReason.Text) Then
        txtHisSICReason.Text = ""
    End If
        
    gSQL = "exec usp_GetESIC_ByProjectModelStationIDSide " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & "," & CStr(cbxModel.ItemData(cbxModel.ListIndex)) & "," & CStr(cbxStation.ItemData(cbxStation.ListIndex)) & "," & CStr(cbxSide.ItemData(cbxSide.ListIndex)) & "," & CStr(intHistorySIC) & ",'" & txtHisSICReason.Text & "','" & gEmpID & "','" & CStr(cbxSICCategory.ItemData(cbxSICCategory.ListIndex)) & "'"
    'gConStrProject = "Driver={SQL Server};server=suz-spcs-01;database=Northwind;uid=sa;pwd=spcs01"
    
    If DB_Project_RecordSet(gSQL, rstTemp) Then
        If rstTemp.EOF Then
            rstTemp.Close
            
            MsgBox "Error Code 5001: Can NOT find SIC in system!", vbExclamation
            
            Exit Sub
        End If
        
        strURL = IIf(IsNull(rstTemp.Fields("FileUrl")), "", Trim(rstTemp.Fields("FileUrl")))
        gSICID = CStr(rstTemp.Fields("EsicID"))
        
        rstTemp.Close
    Else
        'Me.MousePointer = vbNormal
        MsgBox "Error Code 3002: Connect to project database failed, usp_GetESIC_ByProjectModelStationID!!!"
        Unload Me
    End If
    
    If IsNull(strURL) Or strURL = "" Then
        MsgBox "Error Code 5002, SIC file NOT configured in system!", vbExclamation
        Exit Sub
    End If
    
    'WebBrowser1.Navigate ("http://localhost/kanban/extend/pdf/F.pdf")
    'WebBrowser1.Navigate2 strURL
    If (cbxSICCategory.ItemData(cbxSICCategory.ListIndex) = "2") Then
        gURL = Replace(gProjectWebURL, LCase(Trim(cbxProject.Text)), "generalsic") + strURL
   
    Else
        gURL = gProjectWebURL + strURL
    End If
    gSIC = cbxProject.Text + " - " + cbxStation.Text + " - " + cbxModel.Text + " - " + cbxSide.Text + " - " + cbxSICCategory.Text
    
    
    If intHistorySIC = 1 Then
        gSIC = gSIC + "(OLD)"
    End If
    
    If IsNull(gURL) Or Len(gURL) = 0 Then
        MsgBox "SIC Not FOUND!", vbExclamation
    Else
        Unload Me
    End If
  
End Sub

Private Sub Command1_Click()
Dim inifile As String
inifile = App.Path & "\SICSetting.ini"

Dim i As Integer
Dim boolproject, boolsic, boolmodel, boolstation As Boolean
   
    For i = 0 To cbxProject.ListCount - 1
        If cbxProject.List(i) = GetPrivateStringValue("setting", "Project", inifile) Then
        cbxProject.ListIndex = i
        boolproject = True
            Exit For
        End If
    Next
    For i = 0 To cbxSICCategory.ListCount - 1
        If cbxSICCategory.List(i) = GetPrivateStringValue("setting", "Category", inifile) Then
        cbxSICCategory.ListIndex = i
        boolsic = True
            Exit For
        End If
    Next
    For i = 0 To cbxModel.ListCount - 1
        If cbxModel.List(i) = GetPrivateStringValue("setting", "Model", inifile) Then
        cbxModel.ListIndex = i
        boolmodel = True
            Exit For
        End If
    Next
    'ASSY. EM.IF/UF CB ASSY-1
    
    For i = 0 To cbxStation.ListCount - 1
        If cbxStation.List(i) = GetPrivateStringValue("setting", "Station", inifile) Then
        cbxStation.ListIndex = i
        boolstation = True
            Exit For
        End If
    Next
    
    
    If boolproject And boolsic And boolmodel And boolstation Then
    cmdOpen_Click
    Else
    MsgBox "SIC权限不足，请认证", vbExclamation
    End If
    



End Sub

Private Sub Form_Load()
    Dim rstTemp As New ADODB.Recordset
    ChkAllModel.Enabled = False
    Dim buttontext As String
    
    buttontext = "&Open" & vbCr & "打开"
    cmdOpen.Caption = buttontext
    buttontext = "&Cancel" & vbCr & "取消"
    cmdCancel.Caption = buttontext
    txtLine.Text = gLineName
    cbxProject.Clear
    Frame1.Visible = False
    ChkGetData.Top = 2760
    frmESICSelect.Height = 6300
    Frame2.Top = 3240
    
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
    If cbxProject.ListCount = 1 Then
                cbxProject.ListIndex = 0
    End If
    
    
  
    
End Sub


Private Sub Get_Click()
    If Trim(txtOrder.Text) = "" Then
        MsgBox "Error Code 3005: Please input the data"
        Exit Sub
    End If
    If Trim(cbxExternalType.Text) = "" Then
        MsgBox "Please select ExternalType"
        Exit Sub
    End If
    
    Dim rstTemp As New ADODB.Recordset
    cbxSide.Clear
    cbxStation.Clear
    
    
    'get model by external type
    
     If DB_Project_RecordSet("exec usp_GetSICModel_ByExternalDate " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & "," & CStr(cbxProject.ItemData(cbxExternalType.ListIndex)) & "," & Trim(txtOrder.Text), rstTemp) Then
            cbxModel.Clear
            While Not rstTemp.EOF
                cbxModel.AddItem (rstTemp.Fields("ModelName"))
                cbxModel.ItemData(cbxModel.ListCount - 1) = rstTemp.Fields("ModelID")
                'cbxmodel.ItemData(
                rstTemp.MoveNext
                
            Wend
            If cbxModel.ListCount = 1 Then
                cbxModel.ListIndex = 0
            End If
            rstTemp.Close
        Else
            MsgBox "Error Code 3003: Connect to project database failed, usp_GetSICModel_ByExternalDate!!!"
        End If
    
    
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
  If gstartid Then
   gstartid = False
   Command1_Click
    
    
    
  End If
End Sub
