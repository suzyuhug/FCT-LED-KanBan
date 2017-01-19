VERSION 5.00
Begin VB.Form frmClientRegister 
   Caption         =   "Register Estation Client"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4725
   Icon            =   "frmClientRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHostDesc 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   840
      Width           =   2295
   End
   Begin VB.ComboBox cbxLine 
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin VB.ComboBox cbxBU 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtHostName 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Description:"
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
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Line:"
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
      Left            =   480
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Building:"
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
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "PC Name:"
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
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmClientRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbxBU_Change()
    Dim rstTemp As New ADODB.Recordset
    
    cbxLine.Clear
    
    If DB_Master_RecordSet("exec usp_GetLine_ByBU '" & cbxBU.Text & "'", rstTemp) Then
        While Not rstTemp.EOF
            cbxLine.AddItem (rstTemp.Fields("LineName"))
            cbxLine.ItemData(cbxLine.ListCount - 1) = rstTemp.Fields("LineID")
            rstTemp.MoveNext
        Wend
    
        rstTemp.Close
    Else
        'Me.MousePointer = vbNormal
        MsgBox "Error Code 3007: Connect to Master database failed, usp_GetLine_ByBU!!!"
        Unload Me
    End If
End Sub

Private Sub cbxBU_Click()
    Call cbxBU_Change
End Sub

Private Sub cmdRegister_Click()

    If (cbxBU.ListIndex < 0) Then
        MsgBox "Please select Building!", vbExclamation
        
        Exit Sub
    End If
    
    If (cbxLine.ListIndex < 0) Then
        MsgBox "Please select Line!", vbExclamation
        
        Exit Sub
    End If
    
    If IsNull(txtHostName.Text) Or txtHostName.Text = "" Then
        MsgBox "Can NOT get Host Name!", vbExclamation
        
        Exit Sub
    End If
    
    If IsNull(txtHostDesc.Text) Or txtHostDesc.Text = "" Then
        MsgBox "Please input description!", vbExclamation
        
        Exit Sub
    End If
    
    gSQL = "EXEC usp_Client_Register '" & txtHostName.Text & "'," & CStr(cbxLine.ItemData(cbxLine.ListIndex)) & ",'" & txtHostDesc.Text & "'"
   
    If DB_Master_ExecSQL(gSQL) Then
        MsgBox "Register client sucessfully. ", vbInformation
        
        End
    Else
        MsgBox "Error Code 3008: Connect to master database failed, usp_Client_Register!!!", vbExclamation
    End If
   
End Sub

Private Sub Form_Load()
    txtHostName = gClientName
    
    Dim rstTemp As New ADODB.Recordset
    
    If DB_Master_RecordSet("exec usp_GetBU", rstTemp) Then
        While Not rstTemp.EOF
            cbxBU.AddItem (rstTemp.Fields("location"))
            
            rstTemp.MoveNext
        Wend
    
        rstTemp.Close
    Else
        'Me.MousePointer = vbNormal
        MsgBox "Error Code 3006: Connect to Master database failed, usp_GetBU!!!", vbExclamation
        Unload Me
    End If
    
End Sub
