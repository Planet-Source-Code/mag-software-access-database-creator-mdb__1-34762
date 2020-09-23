Attribute VB_Name = "mDatabase"
'------------------Module header sample----------------------------------
'
' Module mDatabase
' File: mDatabase.bas
' Author: Miroslav Milak, mmilak@net4u.hr
' Date: 13.5.2002
' Purpose: Create BAS module you can include in applications to create access
'           database from code
'
' History:
'     13.5.2002. Created
'
' ToDo list:
'        - add proper error handling (both project and module)
'        - split command text for queries and views ?!?!
'        - proper enumerations
'        - open databases with password
'        - automaticly open created bas module
'
'  Known problems:
'        -huge databases produce too large procedure (over 64K)
'        -column description with quotes
'
'  Wish list:
'        -you choose :)
'------------------------------------------------------------------------

Option Explicit

Dim cat As Catalog
Dim i As Integer
Dim prop As Property
Dim bas As TextStream
Dim proc As Procedure

Private Const vbTab2 = vbTab & vbTab
Private Const vbTab3 = vbTab & vbTab & vbTab

Public Function CreateBAS(sPath As String)
Dim iRetVal As Integer

Dim oFso As FileSystemObject
Dim sFileName As String
Dim sDbName As String

Dim oLine As String

Set oFso = New FileSystemObject
Screen.MousePointer = vbHourglass
DoEvents

sDbName = ".mdb"
Set bas = oFso.CreateTextFile(sPath, True)


'write headers
bas.WriteLine "Attribute VB_Name = """ & sPath & """"
bas.WriteLine "Option Explicit"
bas.WriteBlankLines 1

bas.WriteLine "'***************************"
bas.WriteLine "'Database BAS Creator module generator"
bas.WriteLine "'Made by Miroslav Milak, mmilak@net4u.hr"
bas.WriteBlankLines 1
bas.WriteLine "'Module created: " & Now
bas.WriteLine "'Note:"
bas.WriteBlankLines 1
bas.WriteLine "'References to include in your product:"
bas.WriteLine vbTab & "'Microsoft scripting runtime"
bas.WriteLine vbTab & "'Microsoft ADO Extensions 2.x for DDL and Security Object Model"
bas.WriteLine vbTab & "'Microsoft ActiveX Data Object Library 2.x"
bas.WriteLine "'***************************"
bas.WriteBlankLines 1

'write declarations
bas.WriteLine "Dim Cat as New ADOX.Catalog"
bas.WriteLine "Dim Col as Column"
bas.WriteLine "Dim Tbl as Table"
bas.WriteLine "Dim Key as Key"
bas.WriteLine "Dim Idx as Index"
bas.WriteBlankLines 1

'write main function
oLine = "Public Sub Main" & vbNewLine
oLine = oLine & vbNewLine
oLine = oLine & vbTab & "If not CreateDatabase Then Exit sub" & vbNewLine
oLine = oLine & vbTab & "CreateTables" & vbNewLine
oLine = oLine & vbTab & "CreateIndexes" & vbNewLine
oLine = oLine & vbTab & "CreateKeys" & vbNewLine
oLine = oLine & vbTab & "CreateViews" & vbNewLine
oLine = oLine & vbTab & "CreateProcedures" & vbNewLine
oLine = oLine & "Set cat = Nothing"
oLine = oLine & vbNewLine
oLine = oLine & "End sub"
bas.WriteLine oLine

'write create database
oLine = "public function CreateDatabase()" & vbNewLine
oLine = oLine & "On Error GoTo PROC_ERR"
oLine = oLine & vbNewLine
oLine = oLine & "" & vbNewLine
oLine = oLine & "Cat.Create "
oLine = oLine & Chr(34) & "Provider=Microsoft.jet.oledb.4.0; Data Source = " & sPath & sDbName & Chr(34) & vbNewLine
oLine = oLine & vbNewLine
bas.WriteLine oLine
bas.WriteLine "CreateDatabase = True"
bas.WriteLine "Exit Function"
bas.WriteLine "PROC_ERR:"
bas.WriteLine "CreateDatabase = False"
bas.WriteLine "IF Err.Number = -2147217897 THEN"
bas.WriteLine vbTab & "'Database already exists"
bas.WriteLine vbTab & "MsgBox" & """Database already exists"""
bas.WriteLine vbTab & "Exit Function"
bas.WriteLine "Else"
bas.WriteLine vbTab & "MsgBox Err.Number & vbNewLine & Err.Description"
bas.WriteLine "End If"
bas.WriteBlankLines 1
bas.WriteLine "End Function"



CreateTables
CreateIndexes
CreateKeys
CreateViews
CreateProcedures

bas.Close

Screen.MousePointer = vbDefault
iRetVal = MsgBox("Module succesfully created" & vbNewLine & "Would you like to open bas file?", vbQuestion + vbYesNo, App.Title)

If iRetVal = vbYes Then
'   MsgBox Shell(sPath, vbMaximizedFocus)
End If

End Function

Public Function CreateTables()

'create tables
bas.WriteLine ""
bas.WriteLine "Private Function CreateTables()"
bas.WriteLine "On Error GoTo PROC_ERR"

For Each tbl In db.Tables
   'kreiraj tablicu
   Select Case tbl.Type
      Case "TABLE"
         bas.WriteLine "'==TABLE " & tbl.Name
         bas.WriteLine vbTab & "Set tbl = New ADOX.Table"
         bas.WriteLine vbTab & "With tbl"
         bas.WriteLine vbTab2 & ".Name=""" & tbl.Name & """"
         bas.WriteLine vbTab2 & "Set .ParentCatalog = cat"
         
'stupci
         bas.WriteLine vbTab2 & "with .columns"
         
         For Each Col In tbl.Columns
            bas.WriteLine vbTab3 & "'Columns: " & Col.Name
            bas.WriteLine vbTab3 & ".Append " & Chr(34) & Col.Name & Chr(34) & ", " & fnDataType(Col.Type) & ", " & Col.DefinedSize
            
            If Col.Precision <> 0 Then bas.WriteLine vbTab3 & ".item(""" & Col.Name & """).precision= " & Chr(34) & Col.Precision & Chr(34)
'atributi rade probleme
'            If Col.Attributes <> 0 Then bas.WriteLine vbTab3 & ".item(""" & Col.Name & """).Attributes= " & Chr(34) & Col.Attributes & Chr(34)
            If Col.NumericScale <> 0 Then bas.WriteLine vbTab3 & ".item(""" & Col.Name & """).NumericScale= " & Chr(34) & Col.NumericScale & Chr(34)
            
            'properties
            For Each prop In Col.Properties
               Select Case prop.Name
                  Case "Autoincrement"
                     If prop.Value <> "False" Then
                        bas.WriteLine vbTab3 & ".item(""" & Col.Name & """).properties(""Autoincrement"").value = " & prop.Value
                     End If

                  Case "Default"
                     If prop.Value <> "" Then
                        bas.WriteLine vbTab3 & ".item(""" & Col.Name & """).properties(""Default"").value = " & Chr(34) & prop.Value & Chr(34)
                     End If
                  
                  Case "Description"
                     If prop.Value <> "" Then
                        bas.WriteLine vbTab3 & ".item(""" & Col.Name & """).properties(""Description"").value = " & Chr(34) & prop.Value & Chr(34)
                     End If
                                        
                  Case "Seed"
                  
                  Case "Increment"
                  Case "Fixed Length"
                  Case "Nullable"
                  Case "Jet OLEDB:Column Validation Text"
                  Case "Jet OLEDB:Column Validation Rule"
                  Case "Jet OLEDB:Allow Zero Length"
                  
                                       
                  Case Else
'                     Debug.Print prop.Name & " " & prop.Value
                     
               End Select
            
            Next prop
            
            bas.WriteBlankLines 1
         Next Col
         
         bas.WriteLine vbTab2 & "End with"
         bas.WriteLine vbTab & "End with"
         
         'dodaj tablicu u bazu
         bas.WriteLine vbTab & "cat.Tables.Append tbl"
         bas.WriteLine vbTab & "Set tbl = nothing"
         bas.WriteBlankLines 1
   End Select

Next tbl

bas.WriteLine "Exit Function"
bas.WriteLine "PROC_ERR:"
bas.WriteLine vbTab & "MsgBox Err.Number & vbNewLine & Err.Description"
bas.WriteLine "End Function"


End Function
'
'Public Function CreatePrimaryKeys()
'
'bas.WriteLine "Private Function CreatePrimaryKeys"
'bas.WriteBlankLines 1
''za svaku tablicu
'bas.WriteLine "For Each tbl In db.Tables"
'For Each tbl In db.Tables
'
'   If tbl.Type = "TABLE" And tbl.Keys.Count > 0 Then
'      'upis u bas
'
'
'      For Each key In tbl.Keys
'
'         If key.Type = 1 Then 'primary key
'            bas.WriteLine vbTab & "tbl.keys.append " & Chr(34) & key.Name & Chr(34)
'            Debug.Print tbl.Name, key.Name, key.Type
'         End If
'      Next key
'
'   End If
'
'Next tbl
'bas.WriteLine "Next tbl"
'bas.WriteBlankLines 1
'bas.WriteLine "End Function"
'
'End Function

Public Function CreateIndexes()

'write bas header
bas.WriteLine "Private Function CreateIndexes()"
'error handler
bas.WriteLine "On Error GoTo PROC_ERR"

For Each tbl In db.Tables
   
   If tbl.Indexes.Count > 0 And tbl.Type = "TABLE" Then
      For Each idx In tbl.Indexes
         If idx.PrimaryKey Then
            WriteIndex tbl.Name, idx
         End If
'         Debug.Print tbl.Name, idx.Name, idx.PrimaryKey, idx.Clustered, idx.IndexNulls, idx.Unique
      Next idx
   End If

Next tbl

bas.WriteLine "Set idx = Nothing"
bas.WriteBlankLines 1
bas.WriteLine "Exit Function"
bas.WriteLine "PROC_ERR:"
bas.WriteLine vbTab & "MsgBox Err.Number & vbNewLine & Err.Description"
bas.WriteBlankLines 1
bas.WriteLine "End Function"

End Function

Private Function WriteIndex(tblName As String, idx As Index)

bas.WriteLine "'Create index: " & idx.Name
bas.WriteLine "SET idx = new ADOX.Index"
bas.WriteLine "IDX.Name = """ & idx.Name & """"

For i = 0 To idx.Columns.Count - 1
   bas.WriteLine vbTab & "IDX.Columns.Append """ & idx.Columns(i).Name & """"
Next i

bas.WriteLine vbTab & "IDX.PrimaryKey = " & idx.PrimaryKey
bas.WriteLine vbTab & "IDX.Unique = " & idx.Unique
bas.WriteLine vbTab & "IDX.Clustered = " & idx.Clustered
bas.WriteLine vbTab & "IDX.IndexNulls = " & idx.IndexNulls
bas.WriteLine vbTab & "cat.Tables(""" & tblName & """).Indexes.Append IDX"
bas.WriteBlankLines 1

End Function

Public Function CreateKeys()

'write bas header
bas.WriteLine "Private Function CreateKeys()"
'error handler
bas.WriteLine "On Error GoTo PROC_ERR"

For Each tbl In db.Tables
   
   If tbl.Indexes.Count > 0 And tbl.Type = "TABLE" Then
      For Each key In tbl.Keys
         If key.Type <> adKeyPrimary Then
            WriteKey tbl.Name, key
            'Debug.Print tbl.Name, key.Name, key.RelatedTable, key.Type, key.UpdateRule
         End If
      Next key
   End If

Next tbl

bas.WriteBlankLines 1
bas.WriteLine "Exit Function"
bas.WriteLine "PROC_ERR:"
bas.WriteLine vbTab & "MsgBox Err.Number & vbNewLine & Err.Description"
bas.WriteBlankLines 1

bas.WriteLine "End Function"
bas.WriteBlankLines 1

End Function

Private Function WriteKey(tblName As String, key As key)

bas.WriteLine "'Create KEY: " & key.Name
bas.WriteLine "SET key= new ADOX.key"
bas.WriteLine vbTab & "key.Name = """ & key.Name & """"
bas.WriteLine vbTab & "key.Type = " & key.Type
bas.WriteLine vbTab & "key.RelatedTable= " & Chr(34) & key.RelatedTable & Chr(34)

For i = 0 To key.Columns.Count - 1
   bas.WriteLine vbTab & "key.Columns.Append """ & key.Columns(i).Name & """"
   bas.WriteLine vbTab & "key.Columns(""" & key.Columns(i).Name & """).RelatedColumn = """ & key.Columns(i).RelatedColumn & """"
Next i

bas.WriteLine vbTab & "key.UpdateRule = " & key.UpdateRule
bas.WriteLine vbTab & "key.DeleteRule= " & key.DeleteRule

'bas.WriteLine vbTab & "key.PrimaryKey = " & key.PrimaryKey


bas.WriteLine vbTab & "cat.Tables(""" & tblName & """).Keys.Append key"
bas.WriteLine vbTab & "Set Key = Nothing"
bas.WriteBlankLines 1

End Function

Public Function CreateViews()
Dim cm As ADODB.Command

'write bas header
bas.WriteLine "Private Function CreateViews()"
bas.WriteBlankLines 1
bas.WriteLine "Dim cm as ADODB.Command"

'error handler
bas.WriteLine "On Error GoTo PROC_ERR"

For Each View In db.Views
   Set cm = View.Command
   
   WriteView View, cm

Next View

bas.WriteBlankLines 1
bas.WriteLine "Exit Function"
bas.WriteLine "PROC_ERR:"
bas.WriteLine vbTab & "MsgBox Err.Number & vbNewLine & Err.Description"
bas.WriteBlankLines 1
bas.WriteLine "End Function"

End Function

Private Function WriteView(View As View, cm As ADODB.Command)

bas.WriteLine "'Create View: " & View.Name
bas.WriteLine "SET cm = New ADODB.Command"

'remove crlf and ' from cm.commandtext
cm.CommandText = Replace(cm.CommandText, vbCrLf, " ", , , vbBinaryCompare)
cm.CommandText = Replace(cm.CommandText, """", "'", , , vbBinaryCompare)

bas.WriteLine "cm.commandText = """ & cm.CommandText & """"

bas.WriteLine vbTab & "cat.Views.Append " & Chr(34) & View.Name & Chr(34) & ", cm"
bas.WriteLine vbTab & "Set cm = Nothing"
bas.WriteBlankLines 1

End Function


Public Function CreateProcedures()
Dim cm As ADODB.Command

'write bas header
bas.WriteLine "Private Function CreateProcedures()"
bas.WriteBlankLines 1
bas.WriteLine "Dim cm as ADODB.Command"

'error handler
bas.WriteLine "On Error GoTo PROC_ERR"


For Each proc In db.Procedures
   Set cm = proc.Command
   
   WriteProcedure proc, cm

Next proc

bas.WriteBlankLines 1
bas.WriteLine "Exit Function"
bas.WriteLine "PROC_ERR:"
bas.WriteLine vbTab & "MsgBox Err.Number & vbNewLine & Err.Description"
bas.WriteBlankLines 1
bas.WriteLine "End Function"

End Function

Private Function WriteProcedure(proc As Procedure, cm As ADODB.Command)

bas.WriteLine "'Create Procedure: " & proc.Name
'bas.WriteLine "SET View = New ADOX.View"
bas.WriteLine "SET cm = New ADODB.Command"

'remove crlf and ' from cm.commandtext
cm.CommandText = Replace(cm.CommandText, vbCrLf, " ", , , vbBinaryCompare)
cm.CommandText = Replace(cm.CommandText, """", "'", , , vbBinaryCompare)

bas.WriteLine "cm.CommandText = """ & cm.CommandText & """"

bas.WriteLine vbTab & "cat.Procedures.Append " & Chr(34) & proc.Name & Chr(34) & ", cm"
bas.WriteLine vbTab & "Set cm = Nothing"
bas.WriteBlankLines 1

End Function
