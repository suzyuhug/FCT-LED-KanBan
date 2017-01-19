VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmArrangeSchedule 
   Caption         =   "Schedule"
   ClientHeight    =   7860
   ClientLeft      =   6045
   ClientTop       =   3195
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   12180
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3600
      TabIndex        =   26
      Top             =   840
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
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
      Left            =   9240
      TabIndex        =   25
      Top             =   840
      Width           =   375
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   7800
      TabIndex        =   24
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   21364738
      CurrentDate     =   40564
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   3120
      TabIndex        =   23
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   21364738
      CurrentDate     =   40564
   End
   Begin VB.TextBox Txtqty 
      Height          =   375
      Left            =   7440
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Txtwo 
      Height          =   375
      Left            =   4560
      TabIndex        =   19
      Top             =   1920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox combProject 
      Height          =   315
      ItemData        =   "Schedule arrange.frx":0000
      Left            =   960
      List            =   "Schedule arrange.frx":0002
      TabIndex        =   17
      Top             =   360
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5295
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   9340
      _Version        =   393216
      Cols            =   7
      FormatString    =   $"Schedule arrange.frx":0004
   End
   Begin VB.CommandButton Commd_Save 
      Caption         =   "Save New"
      Height          =   375
      Left            =   10320
      TabIndex        =   15
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Commd_Query 
      Caption         =   "Update"
      Height          =   375
      Left            =   10320
      TabIndex        =   14
      Top             =   600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   21364737
      CurrentDate     =   40563
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   21364737
      CurrentDate     =   40563
   End
   Begin VB.TextBox Txt_remark 
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox Comb_side 
      Height          =   315
      Left            =   960
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.ComboBox Comb_line 
      Height          =   315
      Left            =   6120
      TabIndex        =   6
      Top             =   360
      Width           =   3015
   End
   Begin VB.ComboBox Comb_model 
      Height          =   315
      Left            =   3600
      TabIndex        =   5
      Top             =   840
      Width           =   5535
   End
   Begin VB.ComboBox Comb_shift 
      Height          =   315
      Left            =   3600
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Qty"
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
      Left            =   6840
      TabIndex        =   22
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "WorkOrder"
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
      Left            =   3240
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label L_project 
      Caption         =   "Project"
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
      Left            =   240
      TabIndex        =   18
      Top             =   360
      Width           =   735
   End
   Begin VB.Label L_Endtime 
      Caption         =   "EndTime"
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
      Left            =   5040
      TabIndex        =   11
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label L_remark 
      Caption         =   "Remark"
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
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label L_Begintime 
      Caption         =   "BeginTime"
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
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label L_side 
      Caption         =   "Side"
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
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.Label L_line 
      Caption         =   "Line"
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
      Left            =   5640
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.Label L_model 
      Caption         =   "Model"
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
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
   Begin VB.Label L_shitft 
      Caption         =   "Shift"
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
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "FrmArrangeSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsConStrProject As String
Dim scheduleid As String

Private Sub Comb_shift_Click()
  If Comb_shift.ListIndex = 0 Then
    DTPicker3.Value = "00:00"
    DTPicker4.Value = "07:00"
  End If
  If Comb_shift.ListIndex = 1 Then
    DTPicker3.Value = "07:00"
    DTPicker4.Value = "15:30"
  End If
  If Comb_shift.ListIndex = 2 Then
    DTPicker3.Value = "15:30"
    DTPicker4.Value = "23:59"
  End If
End Sub

Private Sub combProject_Click()
  Dim rststemp As New ADODB.Recordset
  
  Dim rstTemp As New ADODB.Recordset
  If IsNull(combProject.Text) Or combProject.Text = "" Then
                 MsgBox "Please select project!", vbExclamation
                 Exit Sub
             End If

    If DB_Master_RecordSet("exec usp_GetProject_Info '" & combProject.Text & "'", rstTemp) Then
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
   Comb_line.Clear
  If DB_Project_RecordSet_WithConStr("exec Line_GetProject_Enable '" & CStr(combProject.ItemData(combProject.ListIndex)) & "'", rststemp, lsConStrProject) Then
       While Not rststemp.EOF
            Comb_line.AddItem (rststemp.Fields("LineName"))
            Comb_line.ItemData(Comb_line.ListCount - 1) = rststemp.Fields("LineID")
            rststemp.MoveNext
       Wend
   End If
   rststemp.Close
   Comb_model.Clear
   Dim rststemp2 As New ADODB.Recordset
   If DB_Project_RecordSet_WithConStr("exec Model_Get_Enabled '" & CStr(combProject.ItemData(combProject.ListIndex)) & "',0", rststemp2, lsConStrProject) Then
       While Not rststemp2.EOF
            Comb_model.AddItem (rststemp2.Fields("ModelName"))
            Comb_model.ItemData(Comb_model.ListCount - 1) = rststemp2.Fields("ModelID")
            rststemp2.MoveNext
       Wend
   End If
   rststemp2.Close
   Comb_shift.Clear
   Dim rststemp3 As New ADODB.Recordset
   If DB_Project_RecordSet_WithConStr("exec gCode_Get 'ShiftKind'", rststemp3, lsConStrProject) Then
       While Not rststemp3.EOF
            Comb_shift.AddItem (rststemp3.Fields("strValues"))
            Comb_shift.ItemData(Comb_shift.ListCount - 1) = rststemp3.Fields("ID")
            rststemp3.MoveNext
       Wend
   End If
   rststemp3.Close
   Comb_side.Clear
   Dim rststemp4 As New ADODB.Recordset
   If DB_Project_RecordSet_WithConStr("exec gCode_Get 'ESICSide'", rststemp4, lsConStrProject) Then
       While Not rststemp4.EOF
            Comb_side.AddItem (rststemp4.Fields("strValues"))
            Comb_side.ItemData(Comb_side.ListCount - 1) = rststemp4.Fields("ID")
            rststemp4.MoveNext
       Wend
   End If
   rststemp4.Close
   
   Dim rststemp5 As New ADODB.Recordset
   If DB_Project_RecordSet_WithConStr("exec Schdule_Search '" & CStr(combProject.ItemData(combProject.ListIndex)) & "','',''", rststemp5, lsConStrProject) Then
       If Not rststemp5.EOF Then
          MSFlexGrid1.FormatString = "< SchduleID |<ProjectName |<Shift                     |<Modelname                                                          |< linename               |<Side     |<IsConfirm   "
          FlexGrid_FillWithRecordset MSFlexGrid1, rststemp5, "SchduleID,ProjectName,Shift,Modelname,linename,Side,IsConfirm"
       Else
           MSFlexGrid1.Clear
           MSFlexGrid1.FormatString = "<SchduleID          |<ProjectName          |<Shift     |<Modelname        |<linename     |<Side     |<IsConfirm   "
       End If
   End If
   rststemp5.Close
   Text1.Text = ""
    Text1.Visible = True
End Sub

Private Sub Command1_Click()
'  If Text1.Text = "" Then
'     MsgBox "Input model be like"
'     Exit Sub
'  End If
  Comb_model.Clear
   Dim rststemp2 As New ADODB.Recordset
   If DB_Project_RecordSet_WithConStr("exec Model_Get_Enabled_new '" & CStr(combProject.ItemData(combProject.ListIndex)) & "',0,'" & Text1.Text & "'", rststemp2, lsConStrProject) Then
       While Not rststemp2.EOF
            Comb_model.AddItem (rststemp2.Fields("ModelName"))
            Comb_model.ItemData(Comb_model.ListCount - 1) = rststemp2.Fields("ModelID")
            rststemp2.MoveNext
       Wend
       Text1.Visible = False
   End If
   rststemp2.Close
End Sub

Private Sub Commd_Query_Click()
   Dim shiftid As String
   Dim lineid As String
   Dim modelid As String
   Dim sideid As String
  If gSchedule = 0 Then
      FrmArrangeSchedule.Enabled = False
      frmLoginSchedule.Show
   Else
     If combProject.Text = "" Then
        MsgBox "Choose Project", vbExclamation
        Exit Sub
     End If
     If Comb_line.Text = "" Then
        MsgBox "Choose line", vbExclamation
        Exit Sub
     End If
     If Comb_shift.Text = "" Then
        MsgBox "Choose shift", vbExclamation
        Exit Sub
     End If
     If Comb_side.Text = "" Then
        MsgBox "Choose side", vbExclamation
        Exit Sub
     End If
     If Comb_model.Text = "" Then
        MsgBox "Choose model", vbExclamation
        Exit Sub
     End If
     
     Dim rstgetsche As New ADODB.Recordset
     If DB_Project_RecordSet_WithConStr("exec Schdule_Get '" & CStr(combProject.ItemData(combProject.ListIndex)) & "','" & CStr(scheduleid) & "'", rstgetsche, lsConStrProject) Then
        If rstgetsche.EOF Then
            MsgBox "ScheduleID Cannot be found", vbExclamation
            rstgetsche.Close
            Exit Sub
        Else
           If Comb_shift.ListIndex = -1 Then
              shiftid = Trim(rstgetsche!shiftid)
           Else
             shiftid = CStr(Comb_shift.ItemData(Comb_shift.Index))
           End If
           If Comb_line.ListIndex = -1 Then
               lineid = Trim(rstgetsche!lineid)
           Else
               lineid = CStr(Comb_line.ItemData(Comb_line.ListIndex))
           End If
           If Comb_model.ListIndex = -1 Then
               modelid = Trim(rstgetsche!modelid)
           Else
               modelid = CStr(Comb_model.ItemData(Comb_model.ListIndex))
           End If
           If Comb_side.ListIndex = -1 Then
               sideid = Trim(rstgetsche!sideid)
           Else
               sideid = CStr(Comb_side.ItemData(Comb_side.ListIndex))
           End If
        End If
     Else
        MsgBox "Error Code: 3001, Connect to database failed, Schdule_Update!!!", vbExclamation
                rstgetsche.Close
                Exit Sub
     End If
     Dim rststemp As New ADODB.Recordset
     If DB_Project_RecordSet_WithConStr("exec Schdule_Update '" & CStr(combProject.ItemData(combProject.ListIndex)) & "','" & CStr(scheduleid) & "','" & shiftid & "','" & lineid & "','" & modelid & "','" & Txtwo.Text & "','" & Txtqty.Text & "','" & DTPicker1.Value & " " & DTPicker3.Value & "','" & DTPicker2.Value & " " & DTPicker4.Value & "','" & Txt_remark.Text & "','" & gEmpID & "','" & sideid & "'", rststemp, lsConStrProject) Then
          If rststemp(0) = "1" Then
                    MsgBox rststemp(1), vbExclamation
                    rststemp.Close
                    Exit Sub
                End If
       Else
                MsgBox "Error Code: 3001, Connect to database failed, Schdule_Update!!!", vbExclamation
                rststemp.Close
                Exit Sub
      End If
      rststemp.Close
      
      Dim rststemp6 As New ADODB.Recordset
      If DB_Project_RecordSet_WithConStr("exec Schdule_Search '" & CStr(combProject.ItemData(combProject.ListIndex)) & "','',''", rststemp6, lsConStrProject) Then
          If Not rststemp6.EOF Then
              MSFlexGrid1.FormatString = "< SchduleID |<ProjectName |<Shift                     |<Modelname                                                          |< linename               |<Side     |<IsConfirm   "
              FlexGrid_FillWithRecordset MSFlexGrid1, rststemp6, "SchduleID,ProjectName,Shift,Modelname,linename,Side,IsConfirm"
          Else
              MSFlexGrid1.Clear
              MSFlexGrid1.FormatString = "<SchduleID          |<ProjectName          |<Shift     |<Modelname        |<linename     |<Side     |<IsConfirm   "
          End If
      End If
      rststemp6.Close
   End If
End Sub

Private Sub Commd_Save_Click()
   If gSchedule = 0 Then
      FrmArrangeSchedule.Enabled = False
      frmLoginSchedule.Show
   Else
     If combProject.Text = "" Then
        MsgBox "Choose Project", vbExclamation
        Exit Sub
     End If
     If Comb_line.Text = "" Then
        MsgBox "Choose line", vbExclamation
        Exit Sub
     End If
     If Comb_shift.Text = "" Then
        MsgBox "Choose shift", vbExclamation
        Exit Sub
     End If
     If Comb_side.Text = "" Then
        MsgBox "Choose side", vbExclamation
        Exit Sub
     End If
     If Comb_model.Text = "" Then
        MsgBox "Choose model", vbExclamation
        Exit Sub
     End If
      scheduleid = ""
     Dim rststemp As New ADODB.Recordset
     If DB_Project_RecordSet_WithConStr("exec Schdule_Update '" & CStr(combProject.ItemData(combProject.ListIndex)) & "','" & CStr(scheduleid) & "','" & CStr(Comb_shift.ItemData(Comb_shift.ListIndex)) & "','" & CStr(Comb_line.ItemData(Comb_line.ListIndex)) & "','" & CStr(Comb_model.ItemData(Comb_model.ListIndex)) & "','" & Txtwo.Text & "','" & Txtqty.Text & "','" & DTPicker1.Value & " " & DTPicker3.Value & "','" & DTPicker2.Value & " " & DTPicker4.Value & "','" & Txt_remark.Text & "','" & gEmpID & "','" & CStr(Comb_side.ItemData(Comb_side.ListIndex)) & "'", rststemp, lsConStrProject) Then
          If rststemp(0) = "1" Then
                    MsgBox rststemp(1), vbExclamation
                    rststemp.Close
                    Exit Sub
                End If
       Else
                MsgBox "Error Code: 3001, Connect to database failed, Schdule_Update!!!", vbExclamation
                rststemp.Close
                Exit Sub
      End If
      rststemp.Close
      
      Dim rststemp6 As New ADODB.Recordset
      If DB_Project_RecordSet_WithConStr("exec Schdule_Search '" & CStr(combProject.ItemData(combProject.ListIndex)) & "','',''", rststemp6, lsConStrProject) Then
          If Not rststemp6.EOF Then
              MSFlexGrid1.FormatString = "< SchduleID |<ProjectName |<Shift                     |<Modelname                                                          |< linename               |<Side     |<IsConfirm   "
              FlexGrid_FillWithRecordset MSFlexGrid1, rststemp6, "SchduleID,ProjectName,Shift,Modelname,linename,Side,IsConfirm"
          Else
              MSFlexGrid1.Clear
              MSFlexGrid1.FormatString = "<SchduleID          |<ProjectName          |<Shift     |<Modelname        |<linename     |<Side     |<IsConfirm   "
          End If
      End If
      rststemp6.Close
   End If
End Sub

Private Sub Form_Load()
   Dim rststemp As New ADODB.Recordset
   combProject.Clear
  If DB_Master_RecordSet("exec usp_GetProject_ByUser_Client '" & gEmpID & "'", rststemp) Then
       While Not rststemp.EOF
            combProject.AddItem (rststemp.Fields("ProjectCode"))
            combProject.ItemData(combProject.ListCount - 1) = rststemp.Fields("ProjectID")
            rststemp.MoveNext
       Wend
   End If
   rststemp.Close
   gSchedule = 0
   DTPicker1.Value = Format(Now, "MM/DD/YYYY ")
   DTPicker2.Value = Format(Now, "MM/DD/YYYY ")
   DTPicker3.Value = Format(Time, "hh:nn")
   DTPicker4.Value = Format(Time, "hh:nn")
End Sub

Private Sub MSFlexGrid1_Click()
'  With MSFlexGrid1
'      Comb_shift.Text = .TextMatrix(.Row, 2)
'      Comb_line.Text = .TextMatrix(.Row, 4)
'      Comb_model.Text = .TextMatrix(.Row, 3)
'      Comb_side.Text = .TextMatrix(.Row, 5)
'      scheduleid = .TextMatrix(.Row, 0)
'  End With
End Sub
