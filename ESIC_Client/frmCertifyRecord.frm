VERSION 5.00
Begin VB.Form frmCertifyRecord 
   BorderStyle     =   0  'None
   Caption         =   "frmCertifyRecord"
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7680
      TabIndex        =   20
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CheckBox Check8 
      Caption         =   "No"
      Height          =   255
      Left            =   7920
      TabIndex        =   19
      Top             =   3720
      Width           =   615
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Yes"
      Height          =   255
      Left            =   6480
      TabIndex        =   18
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000000&
      Caption         =   "OK"
      Height          =   495
      Left            =   6360
      TabIndex        =   17
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox TxtEMPID 
      Height          =   375
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   240
      Width           =   1935
   End
   Begin VB.CheckBox Check6 
      Caption         =   "No"
      Height          =   255
      Left            =   7920
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Yes"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.CheckBox Check4 
      Caption         =   "No"
      Height          =   495
      Left            =   7920
      TabIndex        =   6
      Top             =   1800
      Width           =   855
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Yes"
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   1800
      Width           =   855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "No"
      Height          =   495
      Left            =   7920
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Yes"
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "EMPID:"
      Height          =   495
      Left            =   480
      TabIndex        =   16
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   375
      Left            =   5040
      TabIndex        =   15
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   495
      Left            =   3120
      TabIndex        =   13
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   1320
      TabIndex        =   11
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   3600
      Width           =   4995
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   4995
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   4995
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   4995
   End
End
Attribute VB_Name = "frmCertifyRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnCancel_Click()
    'Make the form not on top
    SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOZORDER + SWP_NOMOVE + SWP_NOSIZE
    'Make the form not on top
    
    Unload Me

'    'Make the form re-on top
'    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOZORDER + SWP_NOMOVE + SWP_NOSIZE
'    'Make the form re-on top
    
    
End Sub

Private Sub Check1_Click()
    Check2.Value = 0
End Sub

Private Sub Check2_Click()
    Check1.Value = 0
End Sub

Private Sub Check3_Click()
    Check4.Value = 0
End Sub

Private Sub Check4_Click()
    Check3.Value = 0
End Sub

Private Sub Check5_Click()
    Check6.Value = 0
End Sub

Private Sub Check6_Click()
    Check5.Value = 0
End Sub

Private Sub Check7_Click()
    Check8.Value = 0
    
'    If Check2.Value = 1 Or Check4.Value = 1 Or Check6.Value = 1 Then
'        SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOZORDER + SWP_NOMOVE + SWP_NOSIZE
'        MsgBox ("Some Result is No, can not Pass Certify!")
'        Check7.Value = 0
'        Check8.Value = 0
'        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOZORDER + SWP_NOMOVE + SWP_NOSIZE
'        Exit Sub
'    End If
End Sub

Private Sub Check8_Click()
    Check7.Value = 0
End Sub

Private Sub cmdOK_Click()
    
    Dim i As Integer
    Dim j1, j2, j3, j4 As Integer
    i = 1
    Dim Resultid As Integer
    Dim rstTemp As New ADODB.Recordset
    Dim j As Integer
    
    If Check1.Value = 1 Then j1 = 1
    If Check2.Value = 1 Then j1 = 2
    If Check3.Value = 1 Then j2 = 1
    If Check4.Value = 1 Then j2 = 2
    If Check5.Value = 1 Then j3 = 1
    If Check6.Value = 1 Then j3 = 2
    If Check7.Value = 1 Then j4 = 1
    If Check8.Value = 1 Then j4 = 2
    
    'if user not choose any result
    If (j1 = 0 Or j2 = 0 Or j3 = 0 Or j4 = 0) Then
        'Make the form not on top
        SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOZORDER + SWP_NOMOVE + SWP_NOSIZE
        'Make the form not on top
        
        MsgBox ("Certify Fail! Please choose all Result! / 请选择所有的Result")
        
        'Make the form re-on top
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOZORDER + SWP_NOMOVE + SWP_NOSIZE
        'Make the form re-on top
        Exit Sub
    End If
    'if user not choose any result
    
            

    For i = 1 To 4
        If i = 1 Then
            If Check1.Value = 1 Then j1 = 1
            If Check2.Value = 1 Then j1 = 2
            Resultid = j1
        End If
        If i = 2 Then
            If Check3.Value = 1 Then j2 = 1
            If Check4.Value = 1 Then j2 = 2
            Resultid = j2
        End If
        If i = 3 Then
            If Check5.Value = 1 Then j3 = 1
            If Check6.Value = 1 Then j3 = 2
            Resultid = j3
        End If
        If i = 4 Then
            If Check7.Value = 1 Then j4 = 1
            If Check8.Value = 1 Then j5 = 2
            Resultid = j4
        End If
        
        
        'if there is some "No" item in result
        If (Check2.Value = 1 Or Check4.Value = 1 Or Check6.Value = 1) And Check7.Value = 1 Then
            'Make the form not on top
            SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOZORDER + SWP_NOMOVE + SWP_NOSIZE
            'Make the form not on top

            MsgBox ("Can not Pass Certify! As some result is No ! / 所有结果是Yes时才能通过认证")

            'Make the form re-on top
            SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOZORDER + SWP_NOMOVE + SWP_NOSIZE
            'Make the form re-on top
            Exit Sub
        End If
        'if there is some "No" item in result
        
        
        'if first 3 is Yes, but result is "No"
        If (Check1.Value = 1 And Check3.Value = 1 And Check5.Value = 1) And Check8.Value = 1 Then
            'Make the form not on top
            SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOZORDER + SWP_NOMOVE + SWP_NOSIZE
            'Make the form not on top

            MsgBox ("All the questions are Yes, The result should be Yes ! /  所有问题都是Yes, 最后结果应该也是Yes")

            'Make the form re-on top
            SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOZORDER + SWP_NOMOVE + SWP_NOSIZE
            'Make the form re-on top
            Exit Sub
        End If
        'if first 3 is Yes, but result is "No"


            For j = 0 To frmMain.lstCertOpsList.ListCount - 1
                    'steven 2013 01 17, get all info from the list, and then show form '
                    ESICID = gSICID
                    ClientName = gClientName
                    EmpID = Trim(Left(frmMain.lstCertOpsList.List(j), InStr(frmMain.lstCertOpsList.List(j), "-") - 1))
                    CertifyBy = gEmpID
                    
                    txtempid.Text = ""
                    txtempid.Text = EmpID
                    txtempid.Enabled = False
                    
                    Set rstTemp = New ADODB.Recordset
                    gSQL = "exec usp_Certification_Record " & ESICID & ",'" & ClientName & "',"
                    gSQL = gSQL & "'" & EmpID & "',"
                    gSQL = gSQL & "'" & CertifyBy & "'," & i & "," & Resultid & ""
                    If DB_Project_RecordSet(gSQL, rstTemp) Then
                        If rstTemp.Fields("Result") <> "0" Then
                            MsgBox rstTemp.Fields("Des")
                            Exit Sub
                        End If
                    End If
            Next j
        
    Next
    'frmOpenFlag = False
    Unload Me

    
 If Resultid = 1 Then
    frmMain.frmCertNot.Visible = False
    frmMain.frmCertified.Visible = True
    frmMain.frmCertifing.Visible = False
End If

    
 If Resultid = 2 Then
    'frmMain.frmCertifing.Visible = False
    'frmMain.frmCertNot.Visible = True
    
    frmMain.frmCertified.Visible = False
    frmMain.frmCertifing.Visible = False
    frmMain.frmCertifiedFail.Visible = True
    
    
'    frmMain.Label5.Caption = "Certified (认证失败)"
'    frmMain.Label5.ForeColor = &HFF0000
'    frmMain.frmCertified.Visible = True
    
    'frmMain.frmCertified.Visible = True
    
End If


    
'    Call SICCertList_Update(lsSICID, "1")
'
'    lblCertStatus = "Certified at " + Format(Now(), "yyyy/mm/dd hh:mm")
'    lblCertStatus.ForeColor = vbGreen
'
'    gbCertStarted = False

    
    
End Sub

Private Sub Form_Load()

'        'initial the from certify status
'        frmOpenFlag = True
    
        'Make the form always on top
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOZORDER + SWP_NOMOVE + SWP_NOSIZE
        'Make the form always on top

        
        txtempid.Text = ""
        txtempid.Text = EmpID
        txtempid.Enabled = False
        
        Dim rstTemp As New ADODB.Recordset
        Set rstTemp = New ADODB.Recordset
        gSQL = "exec usp_getCertifyQuestion "
        If DB_Project_RecordSet(gSQL, rstTemp) Then

            If rstTemp.Fields("ID") = 1 Then Label1.Caption = rstTemp.Fields("Question")
            rstTemp.MoveNext
            If rstTemp.Fields("ID") = 2 Then Label2.Caption = rstTemp.Fields("Question")
            rstTemp.MoveNext
            If rstTemp.Fields("ID") = 3 Then Label3.Caption = rstTemp.Fields("Question")
            rstTemp.MoveNext
            If rstTemp.Fields("ID") = 4 Then Label4.Caption = rstTemp.Fields("Question")
        End If

     
        frmMain.frmCertNot.Visible = False
        frmMain.frmCertified.Visible = False
        frmMain.frmCertifing.Visible = True
        
End Sub

