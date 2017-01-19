VERSION 5.00
Begin VB.Form frmVersion 
   Caption         =   "About ESIC System"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3570
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim version As String
    Dim rstTemp As New ADODB.Recordset
    
    If DB_Master_RecordSet("exec usp_GetClientVersion", rstTemp) Then
        If Not rstTemp.EOF Then

              version = rstTemp.Fields("AppRemark")
              For i = 1 To Len(version)
                If Mid(version, i, 1) = "|" Then
                    version = Mid(version, 1, i - 1) + vbCr + Mid(version, i + 1)
                End If
                
              Next i
              'version = Replace(version, "CR", "vbCr")
              Label1.Caption = version
        End If
    Else
       
        MsgBox "Error Code 3002: Connect to Master database failed, usp_GetClientVersion!!!"
        Unload Me
    End If
End Sub
