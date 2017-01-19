Attribute VB_Name = "Module1"
Option Explicit

Public gAppName, gAppVersion As String
Public gstartid As Boolean
Public gSQL As String
Public gURL, gSIC, gSICID, gClientName, gLineName As String
Public gEmpID As String
Public gSchedule As Integer
Public gLineID As String
Public gMasterServer As String

Public gProject, gProjectWebURL, gMasterDefaultURL, gClientReportDefaultURL, gProjectDefaultURL As String
Public gCertificateTime As Integer
Public gClientQueryButtonEnable As String
Public gSelect As Integer

Public gRefreshTime As Integer

Public gSICViewer, gSICViewerHistory As Integer

Public gDebug As Boolean
Public pSICCertList(10, 12) As String
Public gInputBoxTitle, gInputBoxPwdChar, gInputBoxResult As String
 'Constant Declaration
 Private Const ODBC_ADD_DSN = 1        ' Add data source
      Private Const ODBC_CONFIG_DSN = 2     ' Configure (edit) data source
      Private Const ODBC_REMOVE_DSN = 3     ' Remove data source
      Private Const vbAPINull As Long = 0 ' NULL Pointer

      'Function Declare
      #If Win32 Then

          Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
          (ByVal hwndParent As Long, ByVal fRequest As Long, _
          ByVal lpszDriver As String, ByVal lpszAttributes As String) _
          As Long
      #Else
          Private Declare Function SQLConfigDataSource Lib "ODBCINST.DLL" _
          (ByVal hwndParent As Integer, ByVal fRequest As Integer, ByVal _
          lpszDriver As String, ByVal lpszAttributes As String) As Integer
      #End If

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

'　　 函数的使用很简单，我们只须在Form_Load中加入如下语句即可：
'retValue = SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.CurrentX, Me.CurrentY, 300, 300, SWP_SHOWWINDOW)


Public ESICID As String
Public EmpID As String
Public ClientName As String
Public CertifyBy As String
Public frmOpenFlag     As Boolean


Sub Main()
On Error GoTo Local_Error

    Dim rstClient As ADODB.Recordset
    
    Dim CmdData() As String
    Dim i As Integer
    
      #If Win32 Then
          Dim intRet As Long
      #Else
          Dim intRet As Integer
      #End If
      Dim strDriver As String
      Dim strAttributes As String

      'Set the driver to SQL Server because it is most common.
      strDriver = "SQL Server"
      'Set the attributes delimited by null.
      'See driver documentation for a complete
      'list of supported attributes.
      strAttributes = "SERVER=10.194.1.181" & Chr$(0)
      strAttributes = strAttributes & "DESCRIPTION=ESIC DSN" & Chr$(0)
      strAttributes = strAttributes & "DSN=ESIC" & Chr$(0)
      strAttributes = strAttributes & "DATABASE=ESIC" & Chr$(0)
      'strAttributes = strAttributes & "USERID=esic" & Chr$(0)
      'strAttributes = strAttributes & "PWD=esicprod" & Chr$(0)
      strAttributes = strAttributes & "Trusted_Connection=no" & Chr$(0)
      strAttributes = strAttributes & "Network=DBMSSOCN" & Chr$(0)
      strAttributes = strAttributes & "Address=10.194.1.181,1433" & Chr$(0)

'sprintf(attributes + iLen, "Network=dbmssocn");
'
'sprintf(attributes + iLen, "Address=192.168.0.44,1148");

      'To show dialog, use Form1.Hwnd instead of vbAPINull.
      intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, _
      strDriver, strAttributes)
'      If intRet Then
'          MsgBox "DSN Created"
'
'      Else
'          MsgBox "DSN Create Failed"
'      End If
      
    gAppName = "EstationClient"
    gAppVersion = "V2.1.0"
    
    gTest = False
    gDebug = False
    
    gSelect = 0  'choose model first
    
    
    gClientQueryButtonEnable = "NO"
    
    CmdData = Split(Command, " ")
    
    If UBound(CmdData()) >= 1 Then
        If UCase(Trim(CmdData(1))) = "-D" Then
            gDebug = True
        End If
    End If
    
    If UBound(CmdData()) >= 0 Then
        If UCase(Trim(CmdData(0))) = "-T" Then
            gTest = True
        End If

    End If
    
    For i = 1 To UBound(CmdData())
        If UCase(Left(CmdData(i), 7)) = "CLIENT=" Then
            'If InStr(CmdData(i), "=") > 0 Then
                gClientName = Trim(Replace(UCase(CmdData(i)), "CLIENT=", ""))
            'End If
        End If
    Next i
    
'    If gDebug = False Then
'        If App.PrevInstance = True Then
'            MsgBox "Application already startup!", vbExclamation
'
'            Exit Sub
'        End If
'    End If
    
    gURL = ""
    gSIC = ""
    
    If gDebug = True Then
        If IsNull(gClientName) Or gClientName = "" Then
            gClientName = GetPcName()
        End If
    Else
        gClientName = GetPcName()
    End If
    
    If IsNull(gClientName) Or gClientName = "" Or Len(gClientName) = 0 Then
        MsgBox "Error Code: 1001, Can NOT get client Host Name!" & vbCr & "得不到客户端主机名!", vbExclamation
        
        Exit Sub
    End If
    
    Call DB_Initiate
    
    Set rstClient = New ADODB.Recordset
    If DB_Master_RecordSet("exec usp_Client_Validate '" & gClientName & "'", rstClient) Then
        If rstClient.EOF Then
            rstClient.Close
            MsgBox "Error Code: 1002, Client " + gClientName + " NOT configured in system!" & vbCr & "客户端 " + gClientName + " 没有在系统中配置好!", vbExclamation
            
            If (MsgBox("Register this client into system?" & vbCr & "注册此客户端到系统中?", vbYesNo) = vbYes) Then
                frmClientRegister.Show
            End If
            
            Exit Sub
        Else
            gLineName = Trim(IIf(IsNull(rstClient.Fields("LineName")), "", rstClient.Fields("LineName")))
            gLineID = CStr(rstClient.Fields("LineID"))
            
            rstClient.Close
            
            If gLineName = "" Or IsNull(gLineName) Then
                MsgBox "Error Code: 1002, Line NOT configured for Client " + gClientName + " in system!" & vbCr & "客户端 " + gClientName + " 没有配置好对应的生产线!", vbExclamation
                
                Exit Sub
            End If
            
        End If
    End If
    
    If rstClient.State = adStateOpen Then
        rstClient.Close
    End If
    
    Dim rstTemp As New ADODB.Recordset

    If DB_Master_RecordSet("exec usp_Get_MasterDefaultURL '" & gClientName & "'", rstTemp) Then
        gMasterDefaultURL = rstTemp.Fields("MasterDefaultURL")
        gClientReportDefaultURL = rstTemp.Fields("ClientReportDefaultURL")
    End If
    If rstTemp.State = adStateOpen Then
        rstTemp.Close
    End If

    frmLogin.Show
    
    Exit Sub
    
Local_Error:
    MsgBox "Error Code: " & CStr(Err.Number) & " , " & Err.Description, vbCritical

End Sub

Public Function GetPcName() As String
    Dim compname As String, retval As Long
    compname = Space(255)
    retval = GetComputerName(compname, 255)
    compname = Left(compname, InStr(compname, vbNullChar) - 1)
    GetPcName = compname
End Function

