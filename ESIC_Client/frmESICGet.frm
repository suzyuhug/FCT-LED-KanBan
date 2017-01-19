VERSION 5.00
Begin VB.Form frmESICGet 
   Caption         =   "Select ESIC "
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Commom 
      Caption         =   "Common(正常模式)..."
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
      Left            =   6120
      TabIndex        =   15
      Top             =   360
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Query(查询条件)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   8055
      Begin VB.TextBox txtArg 
         Height          =   375
         Index           =   2
         Left            =   1440
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtArg 
         Height          =   375
         Index           =   1
         Left            =   5760
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton BtnGet 
         Caption         =   "&Get(查询)"
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
         Left            =   6600
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtArg 
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Arg 
         Caption         =   "Label2"
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
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Arg 
         Caption         =   "Label2"
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
         Index           =   1
         Left            =   4200
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Arg 
         Caption         =   "Label2"
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
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SICList(SIC列表)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   8175
      Begin VB.ListBox LstSIC 
         Height          =   4560
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   18
         Top             =   240
         Width           =   6495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Remove All"
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
         Left            =   6840
         TabIndex        =   17
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Select All"
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
         Left            =   6840
         TabIndex        =   16
         Top             =   2040
         Width           =   1095
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
         Left            =   6840
         TabIndex        =   6
         Top             =   4320
         Width           =   1095
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
         Left            =   6840
         TabIndex        =   5
         Top             =   3600
         Width           =   1095
      End
   End
   Begin VB.ComboBox cbxProject 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin VB.TextBox txtLine 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   4215
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
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmESICGet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnGet_Click()
    Dim tempstr As String
    If cbxProject.Text = "" Then
        MsgBox "Please choose project first", vbInformation
        Exit Sub
    End If
    
    Dim i As Integer
    tempstr = ""
    For i = 0 To 2
        If Arg(i).Visible = True Then
            tempstr = tempstr & txtArg(i).Text
        End If
    Next i
    If tempstr = "" Then
        MsgBox "Please input the condition", vbInformation
        Exit Sub
    End If
    LstSIC.Clear
    
    i = 0
    
    Dim rstTemp As New ADODB.Recordset
    If DB_Project_RecordSet("exec usp_GetSICList_ByAdvance " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & ",'" & txtArg(0).Text & "','" & txtArg(1).Text & "','" & txtArg(2).Text & "','" & gClientName & "','" & gEmpID & "'", rstTemp) Then
        While Not rstTemp.EOF
            
            LstSIC.AddItem Trim(rstTemp(1))
            LstSIC.ItemData(i) = rstTemp(0)
            i = i + 1
            rstTemp.MoveNext
            
        Wend
       rstTemp.MoveFirst
    End If
    
    
    If i = 1 Then
        LstSIC.Selected(0) = True
        cmdOpen_Click
    End If
    
    
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    For i = 0 To LstSIC.ListCount - 1
        LstSIC.Selected(i) = True
    Next i
End Sub

Private Sub cbxProject_Click()
    Dim rstTemp As New ADODB.Recordset
    Dim i As Integer
    For i = 0 To 2
        Arg(i).Visible = False
        txtArg(i).Visible = False
    Next i
    
    gProject = cbxProject.Text
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
     End If
    If rstTemp.State = adStateOpen Then
            rstTemp.Close
        End If
    i = 0
    If DB_Project_RecordSet("exec usp_GetClinetArg_ByProject " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & ",'" & gClientName & "','" & gEmpID & "'", rstTemp) Then
        While Not rstTemp.EOF
            Arg(i).Visible = True
            txtArg(i).Visible = True
            Arg(i).Caption = rstTemp(1) + ":"
            i = i + 1
            rstTemp.MoveNext
        Wend
        rstTemp.Close
    End If
    
     If txtArg(0).Visible = True Then
        txtArg(0).SetFocus
        
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    'open SIC
    Dim i As Integer
    Dim k As Integer
    
    Dim IsOpened As Boolean
    IsOpened = False
    Dim rstTemp As New ADODB.Recordset
    For m = 0 To LstSIC.ListCount - 1
         If LstSIC.Selected(m) = True Then
            If InStr(LstSIC.List(m), "NotCertified") > 0 Then
                MsgBox "There are any SIC which is not certified,can not be opened", vbInformation
                
                Exit Sub
                
            End If
            
            
         End If
    Next m
         
    For i = 0 To LstSIC.ListCount - 1
         If LstSIC.Selected(i) = True Then
         
            
            gSIC = LstSIC.List(i)
            gSICID = LstSIC.ItemData(i)
            
            'get URL by ESICID
            
            Set rstTemp = New ADODB.Recordset
            
            If DB_Project_RecordSet("exec usp_GetSICByProjectSIC " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & ",'" & gSICID & "'", rstTemp) Then
                If Trim(rstTemp!Category) = 2 Then
                    gURL = Replace(gProjectWebURL, LCase(Trim(cbxProject.Text)), "generalsic") + Trim(rstTemp!FileURL)
                Else
                    gURL = gProjectWebURL + Trim(rstTemp!FileURL)
                End If
            Else
                MsgBox "Can not get URL", vbInformation
                Exit Sub
            End If
           IsOpened = False
           If frmMain.LstSIC.ListCount > 0 Then
            For j = 0 To frmMain.LstSIC.ListCount - 1
                If frmMain.LstSIC.ItemData(j) = gSICID Then
                    IsOpened = True
                    
                    Exit For
                End If
            Next j
            End If
    
            If IsOpened = False Then
                     For k = 0 To 9
                        If (pSICCertList(k, 0) = "NOSIC") Then
                        pSICCertList(k, 0) = gSICID
                        pSICCertList(k, 1) = "0"
            
                        Exit For
                        End If
                    Next k
            
                frmMain.LstSIC.AddItem (gSIC)
                frmMain.LstSIC.ItemData(frmMain.LstSIC.ListCount - 1) = gSICID
            End If
         End If
    Next i
     Unload Me
    'LstSIC.AddItem gSIC
    'LstSIC.ItemData(LstSIC.ListCount - 1) = CLng(gSICID)
    frmMain.LstSIC.ListIndex = frmMain.LstSIC.ListCount - 1
    frmMain.lstSIC_Click
    
    
    
  
    
            
End Sub

Private Sub Command2_Click()
     Dim i As Integer
    For i = 0 To LstSIC.ListCount - 1
        LstSIC.Selected(i) = False
    Next i
End Sub

Private Sub Commom_Click()
    Unload Me
    frmESICSelect.Show vbModal
End Sub

Private Sub Form_Activate()
     If txtArg(0).Visible = True Then
        txtArg(0).SetFocus
        
    End If
    Dim buttontext As String
    buttontext = "&Select All" & vbCr & "全选"
    Command1.Caption = buttontext
    buttontext = "&Remove All" & vbCr & "全不选"
    Command2.Caption = buttontext
    buttontext = "&Open" & vbCr & "打开"
    cmdOpen.Caption = buttontext
    buttontext = "&Cancel" & vbCr & "取消"
    cmdCancel.Caption = buttontext
End Sub

Private Sub Form_Load()
     Dim rstTemp As New ADODB.Recordset
    
    txtArg(0).Text = ""
    txtArg(1).Text = ""
    txtArg(2).Text = ""
    txtLine.Text = gLineName
    cbxProject.Clear
    Arg(0).Visible = False
    Arg(1).Visible = False
    Arg(2).Visible = False
    txtArg(0).Visible = False
    txtArg(1).Visible = False
    txtArg(2).Visible = False
    
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

Private Sub txtSN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Trim(txtSN.Text) = "" Then
            MsgBox "Please input model or serialnumber", vbInformation
            Exit Sub
        End If
         Dim rstTemp As New ADODB.Recordset
    
       If DB_Project_RecordSet("exec usp_GetSpecialSIC_BySN " & CStr(cbxProject.ItemData(cbxProject.ListIndex)) & "," & CStr(cbxModel.ItemData(cbxModel.ListIndex)) & "," & gLineID & "," & CStr(cbxSICCategory.ItemData(cbxSICCategory.ListIndex)) & ",'" & gEmpID & "'", rstTemp) Then
        While Not rstTemp.EOF
            cbxSide.AddItem (rstTemp.Fields("SideName"))
            cbxSide.ItemData(cbxSide.ListCount - 1) = rstTemp.Fields("SideID")
            rstTemp.MoveNext
        Wend
        End If
        
    End If
End Sub

Private Sub txtArg_KeyPress(Index As Integer, KeyAscii As Integer)
     If KeyAscii = 13 Then
        
        BtnGet_Click
    End If
End Sub
