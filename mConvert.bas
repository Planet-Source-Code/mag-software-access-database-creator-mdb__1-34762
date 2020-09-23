Attribute VB_Name = "mConvert"
Option Explicit

Public Function fnKeys(lKey As ADOX.KeyTypeEnum) As String

Select Case lKey
   Case 1:      fnKeys = "adKeyPrimary"
   Case 2:      fnKeys = "adKeyForeign"
   Case 3:      fnKeys = "adKeyUnique"
   
   Case Else
      MsgBox "new key!?!?!"
End Select

End Function

Public Function fnDataType(lType As ADOX.DataTypeEnum) As String

Select Case lType
    Case adVarWChar
        fnDataType = "adVarWChar"
    Case adCurrency
        fnDataType = "adCurrency"
    Case adInteger
        fnDataType = "adInteger"
    Case adDate
        fnDataType = "adDate"
    Case adWChar
        fnDataType = "adWChar"
    Case adLongVarWChar
        fnDataType = "adLongVarWChar"
   Case adDouble
      fnDataType = "adDouble"
    Case adLongVarBinary
        fnDataType = "adLongVarBinary"
    Case adBoolean
        fnDataType = "adBoolean"
    Case adSmallInt
        fnDataType = "adSmallInt"
    Case Else
        fnDataType = CStr(lType)
        Debug.Print fnDataType
End Select

End Function

Public Function fnColumnAttribute(lType As ADOX.ColumnAttributesEnum) As String

Select Case lType
   Case 1: fnColumnAttribute = "adColFixed"
   Case 2: fnColumnAttribute = "adColNullable"
   Case Else
      MsgBox "new key!?!?!"
End Select

End Function

Public Function fnAllowNulls(lAllow As ADOX.AllowNullsEnum) As String

End Function

Public Function fnRules(lAllow As ADOX.RuleEnum) As String

End Function

Public Function fnRights(lAllow As ADOX.RightsEnum) As String

End Function

Public Function fnSortOrder(lAllow As ADOX.SortOrderEnum) As String

End Function


