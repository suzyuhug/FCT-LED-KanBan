VERSION 5.00
Begin VB.Form frmLoginSchedule 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1650
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   974.874
   ScaleMode       =   0  'User
   ScaleWidth      =   3619.636
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2040
      TabIndex        =   5
      Top             =   1080
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1080
   End
End
Attribute VB_Name = "frmLoginSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    Unload Me
    gSchedule = 0
    FrmArrangeSchedule.Enabled = True
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    gEmpID = txtUserName.Text
    Dim rstUser As New ADODB.Recordset
    Dim rstTemp As New ADODB.Recordset
    Dim lsConStrProject As String
    If DB_Master_RecordSet("exec usp_Login_Validate '" & gEmpID & "','" & txtPassword.Text & "'", rstUser) Then
        Me.MousePointer = vbNormal
        If rstUser.EOF Then
            rstUser.Close
            MsgBox "Invalid user/password!!!"
            MsgBox "Error Code: 2001, Invalid user/password!", vbExclamation
            txtPassword.Text = ""
            txtPassword.SetFocus
            
        Else
            rstUser.Close
             If IsNull(FrmArrangeSchedule.combProject.Text) Or FrmArrangeSchedule.combProject.Text = "" Then
                 MsgBox "Please select project!", vbExclamation
                 Exit Sub
             End If

            If DB_Master_RecordSet("exec usp_GetProject_Info '" & FrmArrangeSchedule.combProject.Text & "'", rstTemp) Then
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
    Dim rstCheckStatus As New ADODB.Recordset
    If DB_Project_RecordSet_WithConStr("exec ESIC_CheckAuth_Schedule '" & gEmpID & "','" & FrmArrangeSchedule.combProject.Text & "'", rstCheckStatus, lsConStrProject) Then
          If rstCheckStatus(0) = "1" Then
                    MsgBox rstCheckStatus(1), vbExclamation
                    Exit Sub
                End If
            Else
                MsgBox "Error Code: 3001, Connect to database failed, ESIC_CheckAuth_Schedule!!!", vbExclamation
                Exit Sub
            End If
            
            'Unload Me
            Me.Hide
            gSchedule = 1
            txtUserName.Text = ""
             txtPassword.Text = ""
            FrmArrangeSchedule.Enabled = True
            'FrmArrangeSchedule.Show

        End If
    Else
        Me.MousePointer = vbNormal
        MsgBox "Error Code: 3001, Connect to Master database failed, usp_Login_Validate!!!", vbExclamation
        'lblStatus.Caption = "Connect to Master database failed!!!"
        'lblStatus.ForeColor = vbRed
    End If
End Sub
Private Sub txtUserName_KeyPress(KeyAscii As Integer)
'   Dim rstproject As New ADODB.Recordset
'   If KeyAscii = 13 Then
'      If DB_Master_RecordSet("exec usp_GetProject_ByUser '" & gEmpID & "'", rstproject) Then
'         While Not rstproject.EOF
'            cbxProject.AddItem rstproject.Fields("ProjectName")
'            rstproject.MoveNext
'         Wend
'      End If
'      rstproject.Close
'   End If
End Sub
