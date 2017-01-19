VERSION 5.00
Begin VB.Form frmInputBox 
   BorderStyle     =   0  'None
   Caption         =   "s"
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtData 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Please input password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    gInputBoxResult = "INPUTBOX_CANCEL"
    Unload Me
End Sub

Private Sub cmdOK_Click()
    txtData.Text = Trim(txtData.Text)
    If IsNull(txtData.Text) Or txtData.Text = "" Then
        MsgBox "Please input data!", vbExclamation
        
        txtData.SetFocus
    Else
        gInputBoxResult = txtData.Text
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    lbltitle = gInputBoxTitle
    If gInputBoxPwdChar <> "" And Not IsNull(gInputBoxPwdChar) Then
        txtData.PasswordChar = gInputBoxPwdChar
    End If

End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOK_Click
    End If
End Sub
