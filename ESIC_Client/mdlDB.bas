Attribute VB_Name = "mdlDB"
Option Explicit

Public gConStrMaster(2), gServer(2) As String
Public gMasterOK As Integer

Public gConStrProject As String
Public gTest As Boolean
Dim lsMsg As String

Public Sub DB_Initiate()

    gMasterOK = -1
  
   
    gTest = False
    If gTest = True Then
        gConStrMaster(0) = "Driver={SQL Server};server=10.194.1.181;database=ESIC_Master;uid=ESIC;pwd=ESICPROD;Network Library=DBMSSOCN"
        gServer(0) = "10.194.1.181"
        gConStrMaster(1) = "Driver={SQL Server};server=10.194.1.181;database=ESIC_Master;uid=ESIC;pwd=ESICPROD;Network Library=DBMSSOCN"
        gServer(1) = "10.194.1.181"
        gConStrMaster(2) = "Driver={SQL Server};server=10.194.1.181;database=ESIC_Master;uid=ESIC;pwd=ESICPROD;Network Library=DBMSSOCN"
        gServer(2) = "10.194.1.181"
    Else
        gConStrMaster(0) = "Driver={SQL Server};server=10.194.1.181;database=ESIC_Master;uid=esic;pwd=esicprod;Network Library=DBMSSOCN"
        gServer(0) = "10.194.1.181"
        gConStrMaster(1) = "Driver={SQL Server};server=10.194.1.181;database=ESIC_Master;uid=esic;pwd=esicprod;Network Library=DBMSSOCN"
        gServer(1) = "10.194.1.181"
        gConStrMaster(2) = "Driver={SQL Server};server=10.194.1.181;database=ESIC_Master;uid=esic;pwd=esicprod;Network Library=DBMSSOCN"
        gServer(2) = "10.194.1.181"
    End If
    
End Sub

Public Function DB_Master_RecordSet(ByVal strSQL As String, ByRef rstTemp As ADODB.Recordset) As Boolean
On Error GoTo Local_Error

    Dim i As Integer
    
    i = 0
'    Dim rstTemp As New ADODB.Recordset
    
    'rstTemp = New ADODB.Recordset
    DB_Master_RecordSet = False
    
    If gMasterOK >= 0 And gMasterOK <= 2 Then
        i = gMasterOK
        gMasterOK = -1
        rstTemp.Open strSQL, gConStrMaster(i)
            
        If rstTemp.State = adStateOpen Then
            gMasterServer = gServer(i)
            DB_Master_RecordSet = True
            gMasterOK = i
            i = 3
        End If
    End If
    
    If gMasterOK < 0 Or gMasterOK > 2 Then
        i = 0
        While (i < 3)
            rstTemp.Open strSQL, gConStrMaster(i)
            
            If rstTemp.State = adStateOpen Then
                gMasterServer = gServer(i)
                DB_Master_RecordSet = True
                gMasterOK = i
                i = 3
            End If
        Wend
    End If
    
    Exit Function
    
Local_Error:
 
    i = i + 1
    
    If i < 3 Then
        Resume Next
    Else
        lsMsg = "Can NOT connect to Master database!"
        
        If gDebug = True Then
            lsMsg = lsMsg & vbCr & "Error:" & Err.Description & vbCr & "Connection: " & gConStrMaster(i - 1) & vbCr & "SQL: " & strSQL
        End If
        
        MsgBox lsMsg, vbCritical
    End If

End Function

Public Function DB_Master_ExecSQL(strSQL As String) As Boolean
On Error GoTo Local_Error

    Dim i, iError As Integer
    
    DB_Master_ExecSQL = False
    
    i = 0
    iError = 0
    
    Dim rstTemp As New ADODB.Recordset
    
    If gMasterOK >= 0 And gMasterOK <= 2 Then
        i = gMasterOK
        iError = i
        
        gMasterOK = -1
        rstTemp.Open strSQL, gConStrMaster(i), adOpenForwardOnly
            
        i = i + 1
        If iError < i Then
            DB_Master_ExecSQL = True
            gMasterOK = i
        End If
    End If
    
    If gMasterOK < 0 Or gMasterOK > 2 Then
        i = 0
        iError = 0
        
        While (i < 3)
            rstTemp.Open strSQL, gConStrMaster(i), adOpenForwardOnly
            
            i = i + 1
            If iError < i Then
                DB_Master_ExecSQL = True
                i = 3
            End If
        Wend
    End If
    
    Exit Function
    
Local_Error:
    iError = iError + 1

    
    If i < 3 Then
        Resume Next
    Else
        lsMsg = "Can NOT connect to Master database!"
        
        If gDebug = True Then
            lsMsg = lsMsg & vbCr & "Error:" & Err.Description & vbCr & "Connection: " & gConStrMaster(i - 1) & vbCr & "SQL: " & strSQL
        End If
        
        MsgBox lsMsg, vbCritical
        'MsgBox "Can NOT connect to Master database!", vbCritical
    End If

End Function

Public Function DB_Project_RecordSet(strSQL As String, ByRef rstTemp As ADODB.Recordset) As Boolean
On Error GoTo Local_Error

'    Dim rstTemp As New ADODB.Recordset
    
    'rstTemp = New ADODB.Recordset
    DB_Project_RecordSet = False
    
    rstTemp.Open strSQL, gConStrProject
        
    If rstTemp.State = adStateOpen Then
        DB_Project_RecordSet = True
    End If

    Exit Function
    
Local_Error:

    DB_Project_RecordSet = False
    
    lsMsg = "Can NOT connect to project database!"
    
    If gDebug = True Then
        lsMsg = lsMsg & vbCr & "Error:" & Err.Description & vbCr & "Connection: " & gConStrProject & vbCr & "SQL: " & strSQL
    End If
    
    MsgBox lsMsg, vbCritical
        
End Function

Public Function DB_Project_RecordSet_WithConStr(strSQL As String, ByRef rstTemp As ADODB.Recordset, strConStr As String) As Boolean
On Error GoTo Local_Error

'    Dim rstTemp As New ADODB.Recordset
    
    'rstTemp = New ADODB.Recordset
    DB_Project_RecordSet_WithConStr = False
    
    rstTemp.Open strSQL, strConStr
        
    If rstTemp.State = adStateOpen Then
        DB_Project_RecordSet_WithConStr = True
    End If

    Exit Function
    
Local_Error:

    DB_Project_RecordSet_WithConStr = False

    lsMsg = "Can NOT connect to project database!"
    
    If gDebug = True Then
        lsMsg = lsMsg & vbCr & "Error:" & Err.Description & vbCr & "Connection: " & strConStr & vbCr & "SQL: " & strSQL
    End If
    
    MsgBox lsMsg, vbCritical
    
End Function

Public Function DB_Project_ExecSQL(strSQL As String) As Boolean
On Error GoTo Local_Error

    Dim i, iError As Integer
    
    DB_Project_ExecSQL = False
    
    i = 0
    iError = 0
    
    Dim rstTemp As New ADODB.Recordset
    
    rstTemp.Open strSQL, gConStrProject, adOpenForwardOnly
    
    If rstTemp.State = adStateOpen Then rstTemp.Close
    
    DB_Project_ExecSQL = True

    Exit Function
    
Local_Error:
    DB_Project_ExecSQL = False
    'MsgBox "Can NOT connect to project database!", vbCritical
    lsMsg = "Can NOT connect to project database!"
    
    If gDebug = True Then
        lsMsg = lsMsg & vbCr & "Error:" & Err.Description & vbCr & "Connection: " & gConStrProject & vbCr & "SQL: " & strSQL
    End If
    
    MsgBox lsMsg, vbCritical
End Function

Public Function DB_Project_ExecSQL_WithConStr(strSQL As String, strConStr As String) As Boolean
On Error GoTo Local_Error

    Dim i, iError As Integer
    
    DB_Project_ExecSQL_WithConStr = False
    
    i = 0
    iError = 0
    
    Dim rstTemp As New ADODB.Recordset
    
    rstTemp.Open strSQL, strConStr, adOpenForwardOnly
    
    If rstTemp.State = adStateOpen Then rstTemp.Close
    
    DB_Project_ExecSQL_WithConStr = True

    Exit Function
    
Local_Error:
    'MsgBox "Can NOT connect to project database!", vbCritical
    DB_Project_ExecSQL_WithConStr = False
    
    lsMsg = "Can NOT connect to project database!"
    
    If gDebug = True Then
        lsMsg = lsMsg & vbCr & "Error:" & Err.Description & vbCr & "Connection: " & strConStr & vbCr & "SQL: " & strSQL
    End If
    
    MsgBox lsMsg, vbCritical
End Function

Public Sub ResizeInit(cForm As Form)
'函数说明:用于窗体加载时初始化
'传入参数:cForm-窗体对象

Dim fWidth As Single    '记录当前控件所处于的容器的宽度
Dim fHeight As Single   '记录当前控件所处于的容器的高度
Dim fObj As Control

On Error Resume Next

For Each fObj In cForm
    fWidth = fObj.Container.Width
    fHeight = fObj.Container.Height

    fObj.Tag = fWidth & " " & fHeight & " " & fObj.Left & " " & fObj.Top & " " & fObj.Width & " " & fObj.Height
Next

End Sub
Public Sub ResizeForm(cForm As Form)
'函数说明:更改控件位置
'传入参数:cForm-窗体对象

Dim xSingle As Single
Dim ySingle As Single
Dim fObj As Control
Dim Pos As Variant
'此句是防止窗体在最大化最小时,出现刷新操作浪费时间(更重要的是我在组态中
'当最大化,最小化时,刷新组态里面的控件会出错,所以加上此句)
If cForm.WindowState <> 0 Then Exit Sub

On Error Resume Next

For Each fObj In cForm
    Pos = Split(fObj.Tag, " ")
    xSingle = fObj.Container.Width / Pos(0)
    ySingle = fObj.Container.Height / Pos(1)

    fObj.FontSize = fObj.FontSize / ySingle

    If TypeOf fObj Is ComboBox Then '当控件为ComboBox,此控件不能改变高度,单独处理
      fObj.Left = Pos(2) * xSingle
      fObj.Top = Pos(3) * ySingle
      fObj.Width = Pos(4) * xSingle
    Else
      fObj.Move Pos(2) * xSingle, Pos(3) * ySingle, Pos(4) * xSingle, Pos(5) * ySingle
    End If
Next

End Sub
Sub FlexGrid_FillWithRecordset(ByVal g As MSFlexGrid, ByVal rs As ADODB.Recordset, ByVal FieldList As String)
    On Error Resume Next
    Dim f
    Dim i As Integer
    g.Rows = 1
    If rs Is Nothing Then Exit Sub
    f = Split(FieldList, ",")
    If g.Cols < (UBound(f) + 1) Then
        g.Cols = UBound(f) + 1
    End If
    Do While Not rs.EOF
        g.Rows = g.Rows + 1
        g.Row = g.Rows - 1
        For i = 0 To UBound(f)
            g.TextMatrix(g.Row, i) = ClearNull(rs(Trim(f(i))))
        Next
        rs.MoveNext
    Loop
    If g.Rows > 1 Then g.Row = 1
End Sub
Function ClearNull(s)
    On Error Resume Next
    ClearNull = IIf(IsNull(s), "", s)
End Function


