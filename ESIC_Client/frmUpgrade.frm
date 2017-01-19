VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmUpgrade 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Upgrade Information"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   Icon            =   "frmUpgrade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10215
      ExtentX         =   18018
      ExtentY         =   15478
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim lssql As String
    Dim lsUpgradeURL As String
    
            Dim rstTemp As New ADODB.Recordset
            lssql = "exec usp_ClientApp_Upgrade '" & gClientName & "','" & gAppName & "','" & gAppVersion & "','" & gEmpID & "','" & App.Path & "','" & App.EXEName & "'"
            If DB_Master_RecordSet(lssql, rstTemp) Then
                lsUpgradeURL = IIf(IsNull(rstTemp.Fields("UpgradeURL")), "http://10.194.1.181/esic/forms/ESICUpgrade.html", rstTemp.Fields("UpgradeURL"))
                
                If rstTemp.State = adStateOpen Then
                    rstTemp.Close
                End If
                
                WebBrowser1.Navigate (lsUpgradeURL)
                
            End If
End Sub
