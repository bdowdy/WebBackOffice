VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Function ReturnNoSales(TerminalPath As String, ReturnAll As Boolean, CashierName As String, StartTime As String, EndTime As String) As Integer
On Error GoTo ERH:
    ReturnNoSales = 0
    Dim Db As ADODB.Connection
    Dim RS As ADODB.Recordset
    If ConnectADODB(TerminalPath & CheckStatDb, Db) = False Then Exit Function
       If ReturnAll Then
         Sql = "Select * from [No Sale]"
       Else
         Sql = "Select * from [No Sale] Where "
         Sql = Sql & " TRY_CAST([time] AS datetime) between '" & CDate(StartTime) & "' and '" & CDate(EndTime) & "'"
         If CashierName = "" Then
            Sql = Sql & " and [Register Name] = '" & MyRegisterName & "'"
         Else
            Sql = Sql & " and [Name] = '" & CashierName & "'"
         End If
       End If
       If ConnectADORS(Sql, Db, RS) = False Then Exit Function
        If Not RS.EOF Then
         ReturnNoSales = RecordsCount(RS)
        End If
DbClose Db, RS
Exit Function
ERH:
Err.Clear
End Function
Public Sub NoSales(FileLoc As String, NoSaledata As NoSaleInfo, Svr_Name As String)
On Error GoTo ERH:
Dim Db As ADODB.Connection
Dim RS As ADODB.Recordset
If ConnectADODB(FileLoc & CheckStatDb, Db) = False Then Exit Sub
     If NoSaledata = AddNoSale Then
        Sql = "Insert Into [No Sale] ([Name],[Time],[Date],[Register Name])"
        Sql = Sql & "Values('" & Svr_Name & "','" & Now & "','" & Now & "','" & MyRegisterName & "')"
        Db.Execute Sql
       Else
         Sql = "Delete from [No Sale]"
         Db.Execute Sql
       End If
    DbClose Db, RS
Exit Sub
ERH:
Err.Clear
End Sub
Public Function GetAccess(Dept_Name As String) As Long
On Error Resume Next
Dim Db As ADODB.Connection, Tb As ADODB.Recordset
OpenDb SvrPath.LocalTerminalPath & SystemOptionsDB, Db
 Sql = "Select * from [Access Levels] where [Access Name] = '" & Dept_Name & "'"
ConnectADORS Sql, Db, Tb
   If Tb.EOF Then
      GetAccess = 0
   Else
      GetAccess = Tb![Access Level]
   End If
DbClose Db, Tb
End Function
