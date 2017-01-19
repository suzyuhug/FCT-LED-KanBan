VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "E-Station SIC"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4455
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4455
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login(登录)"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtPwd 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtUser 
      Height          =   405
      Left            =   2040
      TabIndex        =   2
      Top             =   690
      Width           =   1575
   End
   Begin VB.Label lblStatus 
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
      Left            =   600
      TabIndex        =   5
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Password(密码):"
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
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "User(用户名):"
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
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sPath     As String     '用于存储服务器上的共享目录
Private sName     As String     '用于存储应用程序名
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128        'Maintenance String For PSS usage
    OsName As String                    '操作系统的名称
End Type



Private Sub cmdLogin_Click()

On Error GoTo Local_Error

    Dim rstUser As New ADODB.Recordset
    Dim lssql As String
    
    lblStatus.Caption = "Connecting to database... Please Wait"
    lblStatus.ForeColor = vbBlue
    
    Me.MousePointer = vbHourglass
    
    DoEvents:    DoEvents: DoEvents
    'App.MousePointer
    'gConStr = "Driver={SQL Server};server=suz-spcs-01;database=Northwind;uid=sa;pwd=spcs01"
    'rstUser.Open "select 1 from Region where RegionDescription='Eastern'", gConStr, adOpenDynamic
    gEmpID = Trim(txtUser.Text)
    
    If DB_Master_RecordSet("exec usp_Login_Validate '" & gEmpID & "','" & txtPwd.Text & "'", rstUser) Then
        Me.MousePointer = vbNormal
        If rstUser.EOF Then
            rstUser.Close
            lblStatus.Caption = "Invalid user/password!!!"
            lblStatus.ForeColor = vbRed
            MsgBox "Error Code: 2001, Invalid user/password!", vbExclamation
            txtPwd.Text = ""
            txtPwd.SetFocus
            
        Else
            rstUser.Close
               
             'Check is activity
             Dim rstCheckStatus As New ADODB.Recordset
             
            If DB_Master_RecordSet("exec usp_Login_CheckStatus '" & gEmpID & "','" & gClientName & "',0", rstCheckStatus) Then
                If rstCheckStatus.Fields("Result") = "1" Then
                    MsgBox "The user has logged in,can not open again!" & vbCr & "员工已经登陆,不可以再打开!"
                    Exit Sub
                End If
            Else
                MsgBox "Error Code: 3001, Connect to Master database failed, usp_Login_CheckStatus!!!", vbExclamation
                Exit Sub
            End If
            
            Unload Me
            frmMain.Show

        End If
    Else
        Me.MousePointer = vbNormal
        MsgBox "Error Code: 3001, Connect to Master database failed, usp_Login_Validate!!!", vbExclamation
        lblStatus.Caption = "Connect to Master database failed!!!"
        lblStatus.ForeColor = vbRed
    End If
    
    Exit Sub

Local_Error:
    If InStr(Err.Description, "MSWINSCK.OCX") > 0 Then
        MsgBox "Error Code 1003: MSWINSCK.OCX Not Installed!!!", vbExclamation
    Else
        MsgBox "Error Code: " & CStr(Err.Number) & " , " & Err.Description, vbCritical
    End If
    
    End
End Sub

Private Sub Form_Load()
If App.PrevInstance Then
Call MsgBox("对不起本程序已在运行中, 不得重复加载!!", vbCritical)
End
End If


On Error GoTo Local_Error
Dim rstTemp As New ADODB.Recordset
Dim lssql As String
Dim s1    As String
 ' Dim Ver As OSVERSIONINFO
 '   Ver = GetWindowsVersion()
    '操作系统名
'    MsgBox Ver.OsName
'    Exit Sub

    lssql = "exec usp_ClientApp_Validate '" & gClientName & "','" & gAppName & "','" & gAppVersion & "','" & gEmpID & "','" & App.Path & "','" & App.EXEName & "'"
    If DB_Master_RecordSet(lssql, rstTemp) Then
            If rstTemp.Fields("Expired") = "Y" Then
                If rstTemp.State = adStateOpen Then
                    rstTemp.Close
                End If
                    
                    'Unload Me
                    'frmUpgrade.Show
                    'added by terry to upgrade
                    sName = "Update.exe"
                    sPath = "\\10.194.1.181\ESICInstall\ESIC\"
                    If Len(App.Path) > 3 Then
                        s1 = App.Path + "\" + sName
                    Else
                        s1 = App.Path + sName
                    End If
                    
                    'Shell ("net use \\10.194.1.181\ESICInstall\ESIC esicread/user:esicread")
                    'On Error Resume Next
                
                    'FileCopy sPath + sName, s1
                    Dim iStm     As ADODB.Stream
                    Dim iRe     As ADODB.Recordset
                    Dim iConc     As String
            
                '数据库连接字符串
                    iConc = gConStrMaster(0) '"Provider=SQLOLEDB.1;Persist Security Info=True;User ID=esic;Password=esic;Initial Catalog=ESIC_Master;Data Source=10.194.1.127"
            
                    '打开表
                    Set iRe = New ADODB.Recordset
                    iRe.Open "SELECT TOP 1 * FROM UpgradeApplication ORDER BY Insert_DateTime DESC ", iConc, adOpenKeyset, adLockReadOnly
    
            
                    '保存到文件
                    Set iStm = New ADODB.Stream
                    With iStm
                        .Mode = adModeReadWrite
                        .Type = adTypeBinary
                        .Open
                        .Write iRe("AppExeFile")
                        .SaveToFile s1, adSaveCreateOverWrite
                
                    End With
            
          '关闭对象
          iRe.Close
          iStm.Close
          
    
                    Dim a     As Long
    
                    a = Shell(s1, vbNormalFocus)
                    End
                    Exit Sub
      
            Else
                    If rstTemp.State = adStateOpen Then
                        rstTemp.Close
                    End If
                    
            End If
        End If
        Exit Sub
Local_Error:
    MsgBox "Updating Fail,Please contact system admin" & vbCr & "系统升级失败，请联系管理员!"
    End
End Sub

Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdLogin.SetFocus
        cmdLogin_Click
    End If
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPwd.SetFocus
    End If
End Sub
Private Function GetWindowsVersion() As OSVERSIONINFO
    Dim Ver As OSVERSIONINFO
    Ver.dwOSVersionInfoSize = 148
    GetVersionEx Ver
  
        
    With Ver
        Select Case .dwPlatformId
            Case 1
                Select Case .dwMinorVersion
                    Case 0
                        .OsName = "Windows 95"
                    Case 10
                        .OsName = "Windows 98"
                    Case 90
                        .OsName = "Windows Mellinnium"
                End Select
            Case 2
                Select Case .dwMajorVersion
                    Case 3
                        .OsName = "Windows NT 3.51"
                    Case 4
                        .OsName = "Windows NT 4.0"
                    Case 5
                        If .dwMinorVersion = 0 Then
                            .OsName = "Windows 2000"
                        Else
                           .OsName = "Windows XP"
                        End If
                End Select
                
        End Select
    End With
    GetWindowsVersion = Ver
End Function
