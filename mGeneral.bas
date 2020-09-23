Attribute VB_Name = "mGeneral"
Option Explicit

Public sDataPath As String
Public sSystemDatabase As String

Public db As New Catalog
Public jro As JetEngine
Public cn As ADODB.Connection

Public tbl As Table
Public Col As Column
Public idx As Index
Public key As key
Public qry As Procedure
Public View As View
Public grp As Group
Public usr As User
Public prop As Property

Public Const Provider40 As String = "Microsoft.jet.oledb.4.0"
Public Const Provider351 As String = "Microsoft.jet.oledb.3.51"

Public Sub Main()

'init value

Set cn = New ADODB.Connection
frmMain.Show


End Sub

Public Function fnConnectionString()

With cn
   If cn.State = adStateOpen Then
      cn.Close
   End If
   
   .Provider = Provider40
   'this is necceserry for user and groups list
   'you can choose your system.mdw file
'   sSystemDatabase = "D:\Documents and Settings\administrator\Application Data\Microsoft\Access\System.mdw"
'   .Properties("jet oledb:system database").Value = sSystemDatabase
   .ConnectionString = sDataPath
   .Open
End With
End Function
