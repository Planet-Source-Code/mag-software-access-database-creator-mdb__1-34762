VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCommand 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   4860
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmMain.frx":0000
      Top             =   2880
      Visible         =   0   'False
      Width           =   4305
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6075
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   8730
      Top             =   3420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lv 
      Height          =   5535
      Left            =   3090
      TabIndex        =   1
      Top             =   300
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   9763
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   3387
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   5535
      Left            =   180
      TabIndex        =   0
      Top             =   300
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   9763
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   88
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpenDatabase 
         Caption         =   "Open database"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuCompact 
         Caption         =   "Compact database"
      End
      Begin VB.Menu mnuCreatedb 
         Caption         =   "Create BAS Module"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------Module header sample----------------------------------
'
' Form frmMain
' File: frmMain.frm
' Author: dUGI
' Date: 13.5.2002
' Purpose: View database structure
'
'------------------------------------------------------------------------

Option Explicit
Dim i As Integer

Private Function fnGetDatabaseInfo()
Dim node As node

On Error GoTo PROC_ERR

Screen.MousePointer = vbHourglass

'open database
Set db = Nothing
Set db = New Catalog
Set db.ActiveConnection = cn

'enumerate tables
For Each tbl In db.Tables
   'tip tablice
   'View - query
   
 
   If tbl.Type = "TABLE" Then
      'add table
      Set node = tv.Nodes.Add("Table", tvwChild, tbl.Name, tbl.Name)
      node.Tag = "TBL"
      
      'add node - COLUMNS
      Set node = tv.Nodes.Add(tbl.Name, tvwChild, tbl.Name & "_" & "Column", "Columns")
      
      'add node - Indexes
      Set node = tv.Nodes.Add(tbl.Name, tvwChild, tbl.Name & "_" & "Index", "Indexes")
      
      'add node - Keys
      Set node = tv.Nodes.Add(tbl.Name, tvwChild, tbl.Name & "_" & "Key", "Keys")
      
   'columns
      For Each Col In tbl.Columns
         Set node = tv.Nodes.Add(tbl.Name & "_" & "Column", tvwChild, , Col.Name)
         node.Tag = "COL"
      Next Col
      
   'indexes
      For Each idx In tbl.Indexes
         Set node = tv.Nodes.Add(tbl.Name & "_" & "Index", tvwChild, , idx.Name)
         node.Tag = "IDX"
      Next idx
      
      For Each key In tbl.Keys
         Set node = tv.Nodes.Add(tbl.Name & "_" & "Key", tvwChild, , key.Name)
         node.Tag = "KEY"
      Next key
   End If
Next tbl

'enum queries

Dim qry As Procedure
For Each qry In db.Procedures
   'tip tablice
   Set node = tv.Nodes.Add("Query", tvwChild, , qry.Name)
   node.Tag = "QRY"
Next qry


'views
Dim oView As View
For Each oView In db.Views
   Set node = tv.Nodes.Add("View", tvwChild, , oView.Name)
   node.Tag = "VIEW"
Next oView

If sSystemDatabase = "" Then GoTo PROC_EXIT

'groups
Dim grp As Group
Dim usr As User

For Each grp In db.Groups
   Set node = tv.Nodes.Add("Group", tvwChild, grp.Name, grp.Name)
   node.Tag = "Group"
   
   For Each usr In grp.Users
      Set node = tv.Nodes.Add(grp.Name, tvwChild, , usr.Name)
      node.Tag = "User"
   Next usr
Next grp

'users
For Each usr In db.Users
   Set node = tv.Nodes.Add("User", tvwChild, , usr.Name)
   node.Tag = "User"
Next usr
Screen.MousePointer = vbDefault

Exit Function

PROC_EXIT:
   fnConnectionInfo
   Screen.MousePointer = vbDefault
   Exit Function
   
PROC_ERR:
   If Err.Number = 3251 Then
      'if there is not system database... no groups and users
      Resume PROC_EXIT
      
   Else
      MsgBox Err.Number & Err.Description
      Resume PROC_EXIT
   End If
   
End Function

Private Sub Form_Activate()
Form_Resize
End Sub

Private Sub Form_Load()

'pripremi treeview
fnPrepareTree

Me.Caption = "Create BAS-MDB " & "v" & App.Major & "." & App.Minor & "." & App.Revision

End Sub

Private Sub Form_Resize()
On Error Resume Next
tv.Move 50, 50, 2800, Me.ScaleHeight - 50 - sbStatus.Height
If txtCommand.Visible Then
   lv.Move 2900, 50, Me.ScaleWidth - lv.Left - 500, 2000
Else
   lv.Move 2900, 50, Me.ScaleWidth - lv.Left - 500, Me.ScaleHeight - 50 - sbStatus.Height
End If
txtCommand.Move 2900, 2100, Me.ScaleWidth - lv.Left - 50, Me.ScaleHeight - 50 - sbStatus.Height - txtCommand.Top
End Sub

Private Sub mnuCreatedb_Click()

On Error GoTo PROC_ERR

With cDlg
   .DefaultExt = ".bas"
   .FileName = ""
   .CancelError = True
   .Flags = cdlOFNExplorer + cdlOFNFileMustExist
   .DialogTitle = "Enter name for bas file"
   .Filter = "Visual basic module (*.bas)|*.bas|"
   .FilterIndex = 1
   .InitDir = "F:\_\"
   .ShowSave
   
   CreateBAS .FileName
      
End With
Exit Sub

PROC_ERR:
   If Err.Number = 32755 Then
      Exit Sub
   Else
      MsgBox Err.Number & vbNewLine & Err.Description
   End If

End Sub

Private Sub mnuOpenDatabase_Click()
On Error GoTo PROC_ERR

With cDlg
   .CancelError = True
   .Flags = cdlOFNExplorer + cdlOFNFileMustExist
   .Filter = "Access databases (*.mdb)|*.mdb|"
   .DialogTitle = "Choose access database"
   .FilterIndex = 1
   .ShowOpen
   sDataPath = .FileName
   fnConnectionString
   fnPrepareTree
   
   fnGetDatabaseInfo
   
End With
Exit Sub

PROC_ERR:
   If Err.Number = 32755 Then
      Exit Sub
   Else
      MsgBox Err.Number & vbNewLine & Err.Description
   End If
End Sub

Private Sub tv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'Dim col As Column

Dim node As node

Set node = tv.HitTest(x, y)

If node Is Nothing Then Exit Sub

fnShowTextBox node

Select Case node.Tag
   Case "Root"
      fnConnectionInfo

   Case "TBL"
      Set tbl = db.Tables(node.Text)
      fnTableInfo
   
   Case "COL"
      Set tbl = db.Tables(node.Parent.Parent.Text)
      Set Col = tbl.Columns(node.Text)
      fnColumnInfo
      
   Case "IDX"
      Set tbl = db.Tables(node.Parent.Parent.Text)
      Set idx = tbl.Indexes(node.Text)
      fnIndexInfo
      
   Case "KEY"
      Set tbl = db.Tables(node.Parent.Parent.Text)
      Set key = tbl.Keys(node.Text)
      fnKeyInfo
   
   Case "QRY"
      Set qry = db.Procedures(node.Text)
      fnQueryInfo
      
   Case "VIEW"
      Set View = db.Views(node.Text)
      fnViewInfo
   
   Case "Group"
      Set grp = db.Groups(node.Text)
      fnGroupInfo
   
   Case "User"
      Set usr = db.Users(node.Text)
      fnUserInfo
   
End Select


Debug.Print "frmMain::MouseUp ==> " '& tbl.Name
End Sub

Private Sub fnConnectionInfo()
Dim LI As ListItem

On Error GoTo PROC_ERR

lv.ListItems.Clear
With lv
   Set LI = lv.ListItems.Add(, , "Table name")
   LI.SubItems(1) = cn.Attributes
   Set LI = lv.ListItems.Add(, , "Connection string")
   LI.SubItems(1) = cn.ConnectionString
   
   Set LI = lv.ListItems.Add(, , "Default database")
   LI.SubItems(1) = cn.DefaultDatabase
   
   For Each prop In cn.Properties
      Set LI = lv.ListItems.Add(, , prop.Name)
      LI.SubItems(1) = prop.Value
   Next prop

End With

PROC_ERR:
   If Err.Number = -2147217865 Then
      Resume Next
   End If

End Sub

Private Sub fnTableInfo()
Dim LI As ListItem

On Error GoTo PROC_ERR

lv.ListItems.Clear
With lv
   Set LI = lv.ListItems.Add(, , "Table name")
   LI.SubItems(1) = tbl.Name
   Set LI = lv.ListItems.Add(, , "Table type")
   LI.SubItems(1) = tbl.Type
   
   Set LI = lv.ListItems.Add(, , "")
   LI.SubItems(1) = ""
   
   Set LI = lv.ListItems.Add(, , "Date Created")
   LI.SubItems(1) = tbl.DateCreated
   Set LI = lv.ListItems.Add(, , "Date modified")
   LI.SubItems(1) = tbl.DateModified
   
   Set LI = lv.ListItems.Add(, , "")
   LI.SubItems(1) = ""
   
   Set LI = lv.ListItems.Add(, , "Primary keys")
   For Each key In tbl.Keys
      If key.Type = adKeyPrimary Then
         For i = 0 To key.Columns.Count - 1
            LI.SubItems(1) = LI.SubItems(1) & key.Columns(i).Name & ";"
         Next i
      End If
   Next key
   
   Set LI = lv.ListItems.Add(, , "Foerign keys")
   For Each key In tbl.Keys
      If key.Type = adKeyForeign Then
         For i = 0 To key.Columns.Count - 1
            LI.SubItems(1) = LI.SubItems(1) & key.Columns(i).Name & ";"
         Next i
      End If
   Next key

   
   Set LI = lv.ListItems.Add(, , "")
   LI.SubItems(1) = ""
   
   For Each prop In tbl.Properties
      Set LI = lv.ListItems.Add(, , prop.Name)
      LI.SubItems(1) = prop.Value
   Next prop
End With
Exit Sub

PROC_ERR:
   If Err.Number = -2147217865 Then
      Resume Next
   End If
End Sub

Private Sub fnPrepareTree()
Dim node As node

lv.ListItems.Clear

With tv
   .Nodes.Clear
   Set node = .Nodes.Add(, , "Root", "Database:")
   node.Expanded = True
   node.Tag = "Root"
   
   'add tables
   .Nodes.Add "Root", tvwChild, "Table", "Tables"
   
   'add queries
   .Nodes.Add "Root", tvwChild, "Query", "Queries"
  
   'add views
   .Nodes.Add "Root", tvwChild, "View", "Views"
  
   'add groups
   .Nodes.Add "Root", tvwChild, "Group", "Groups"
   
   'add users
   .Nodes.Add "Root", tvwChild, "User", "Users"
End With
End Sub

Private Sub fnColumnInfo()
Dim LI As ListItem

lv.ListItems.Clear
With lv
   Set LI = lv.ListItems.Add(, , "Attributes")
   LI.SubItems(1) = Col.Attributes
   Set LI = lv.ListItems.Add(, , "Defined size")
   LI.SubItems(1) = Col.DefinedSize
   
   Set LI = lv.ListItems.Add(, , "")
   LI.SubItems(1) = Col.Name
   
   Set LI = lv.ListItems.Add(, , "Numeric scale")
   LI.SubItems(1) = Col.NumericScale
   
   Set LI = lv.ListItems.Add(, , "Precision")
   LI.SubItems(1) = Col.Precision
   
   Set LI = lv.ListItems.Add(, , "")
   'LI.SubItems(1) = col.RelatedColumn
   
   'LI.SubItems(1) = col.SortOrder
   
   Set LI = lv.ListItems.Add(, , "Column type")
   LI.SubItems(1) = fnDataType(Col.Type)
   
   Set LI = lv.ListItems.Add(, , "")
   
   For Each prop In Col.Properties
      Set LI = lv.ListItems.Add(, , prop.Name)
      LI.SubItems(1) = prop.Value
   Next prop
End With
End Sub

Private Sub fnIndexInfo()
Dim LI As ListItem

lv.ListItems.Clear
With lv
   Set LI = lv.ListItems.Add(, , "Index name")
   LI.SubItems(1) = idx.Name
   
   Set LI = lv.ListItems.Add(, , "Clustered")
   LI.SubItems(1) = idx.Clustered
   
   Set LI = lv.ListItems.Add(, , "Index nulls")
   LI.SubItems(1) = idx.IndexNulls
   
   Set LI = lv.ListItems.Add(, , "Primary key")
   LI.SubItems(1) = idx.PrimaryKey
   
   Set LI = lv.ListItems.Add(, , "Unique")
   LI.SubItems(1) = idx.Unique
   
   For Each prop In idx.Properties
      Set LI = lv.ListItems.Add(, , prop.Name)
      LI.SubItems(1) = prop.Value
   Next prop
End With
End Sub

Private Sub fnKeyInfo()
Dim LI As ListItem

lv.ListItems.Clear
With lv
   Set LI = lv.ListItems.Add(, , "Key name")
   LI.SubItems(1) = key.Name
   
   Set LI = lv.ListItems.Add(, , "Related table")
   LI.SubItems(1) = key.RelatedTable
   
   Set LI = lv.ListItems.Add(, , "Key type")
   LI.SubItems(1) = fnKeys(key.Type)
   
   Set LI = lv.ListItems.Add(, , "Update rule")
   LI.SubItems(1) = key.UpdateRule
   
End With
End Sub

Private Sub fnQueryInfo()
Dim LI As ListItem

With lv
   .ListItems.Clear
   Set LI = lv.ListItems.Add(, , "Query name")
   LI.SubItems(1) = qry.Name
   
   Set LI = lv.ListItems.Add(, , "Date Created")
   LI.SubItems(1) = qry.DateCreated
   Set LI = lv.ListItems.Add(, , "Date modified")
   LI.SubItems(1) = qry.DateModified
   
   Set LI = lv.ListItems.Add(, , "Command")
   LI.SubItems(1) = qry.Command.CommandText
   
   txtCommand.Text = qry.Command.CommandText
End With



End Sub

Private Sub fnViewInfo()
Dim LI As ListItem

With lv
   .ListItems.Clear
   Set LI = lv.ListItems.Add(, , "Query name")
   LI.SubItems(1) = View.Name
   
   Set LI = lv.ListItems.Add(, , "Date Created")
   LI.SubItems(1) = View.DateCreated
   Set LI = lv.ListItems.Add(, , "Date modified")
   LI.SubItems(1) = View.DateModified
   
   Set LI = lv.ListItems.Add(, , "Command")
   LI.SubItems(1) = View.Command.CommandText
   
   txtCommand.Text = View.Command.CommandText
End With
End Sub


Private Sub fnGroupInfo()
Dim LI As ListItem

With lv
   .ListItems.Clear
   Set LI = lv.ListItems.Add(, , "Group name")
   LI.SubItems(1) = grp.Name
   
   For Each prop In grp.Properties
      Set LI = lv.ListItems.Add(, , prop.Name)
      LI.SubItems(1) = prop.Value
   Next prop
   
End With
End Sub

Private Sub fnUserInfo()
Dim LI As ListItem

With lv
   .ListItems.Clear
   Set LI = lv.ListItems.Add(, , "User name")
   LI.SubItems(1) = usr.Name
   
   For Each prop In usr.Properties
      Set LI = lv.ListItems.Add(, , prop.Name)
      LI.SubItems(1) = prop.Value
   Next prop
   
End With
End Sub

Private Function fnShowTextBox(Nod As node)

Select Case Nod.Tag
   Case "VIEW", "QRY"
      txtCommand.Visible = True
      lv.Move 2900, 50, Me.ScaleWidth - lv.Left, 2000
      txtCommand.Move 2900, 2100, Me.ScaleWidth - lv.Left - 50, Me.ScaleHeight - 50 - sbStatus.Height - txtCommand.Top
      
      
   
   Case Else
      txtCommand.Visible = False
      lv.Move 2900, 50, Me.ScaleWidth - lv.Left, Me.ScaleHeight - sbStatus.Height - 50
   End Select
End Function

