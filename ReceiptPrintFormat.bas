Attribute VB_Name = "ReceiptPrintFormat"
' ============================================================
'  Receipt print size - reads [Receipt Print Format] and wraps
'  menu-item / modifier lines with the Big or Small ESC code.
'
'  Populated by the web back office (Printers & Hardware ->
'  Receipt Print Size). One row per printer:
'     [Item Size] / [Modifier Size] = "Large" | "Small"
'     [Big Esc]   / [Small Esc]     = CSV decimal bytes ("29,33,17")
'
'  Usage in your receipt routine:
'     LoadReceiptPrintFormat "Receipt"          ' or the printer name
'     AllData$ = AllData$ & FormatItemLine(ItemText)
'     AllData$ = AllData$ & FormatModifierLine(ModText)
' ============================================================
Option Explicit

' Module-level codes for the printer last loaded.
Public RPF_BigEsc      As String   ' e.g. Chr(29) & Chr(33) & Chr(17)
Public RPF_SmallEsc    As String   ' e.g. Chr(29) & Chr(33) & Chr(0)
Public RPF_ItemEsc     As String   ' BigEsc or SmallEsc, per [Item Size]
Public RPF_ModEsc      As String   ' BigEsc or SmallEsc, per [Modifier Size]

' Convert a CSV of decimal byte values ("29,33,17") into a raw
' ESC string Chr(29) & Chr(33) & Chr(17). Ignores blanks/invalid.
Public Function BytesFromCsv(ByVal Csv As String) As String
    Dim Parts() As String
    Dim I       As Integer
    Dim N       As Long
    Dim Result  As String
    If Trim$(Csv) = "" Then Exit Function
    Parts = Split(Csv, ",")
    For I = LBound(Parts) To UBound(Parts)
        If IsNumeric(Trim$(Parts(I))) Then
            N = CLng(Trim$(Parts(I)))
            If N >= 0 And N <= 255 Then Result = Result & Chr$(N)
        End If
    Next I
    BytesFromCsv = Result
End Function

' Load the format row for a printer into the module-level codes.
' Falls back to standard ESC/POS (double vs normal) if no row exists.
Public Sub LoadReceiptPrintFormat(ByVal PrinterName As String)
    On Error GoTo ERH
    Dim Db      As ADODB.Connection
    Dim Rs      As ADODB.Recordset
    Dim ItemSz  As String
    Dim ModSz   As String

    ' Defaults (used if the printer has no saved row).
    RPF_BigEsc = Chr$(29) & Chr$(33) & Chr$(17)   ' GS ! 0x11 double H+W
    RPF_SmallEsc = Chr$(29) & Chr$(33) & Chr$(0)  ' GS ! 0x00 normal
    ItemSz = "Small"
    ModSz = "Small"

    If ConnectADODB("MSaleData", Db) = False Then GoTo Finish

    Sql = "Select * from [Receipt Print Format] where [Printer Name] = '" _
          & Replace(PrinterName, "'", "''") & "'"
    ConnectADORS Sql, Db, Rs

    If Not Rs Is Nothing Then
        If Not Rs.EOF Then
            If Trim$(Rs![Big Esc] & "") <> "" Then RPF_BigEsc = BytesFromCsv(Rs![Big Esc] & "")
            If Trim$(Rs![Small Esc] & "") <> "" Then RPF_SmallEsc = BytesFromCsv(Rs![Small Esc] & "")
            ItemSz = Rs![Item Size] & ""
            ModSz = Rs![Modifier Size] & ""
        End If
    End If

Finish:
    dbClose Db, Rs

    If StrComp(ItemSz, "Large", vbTextCompare) = 0 Then
        RPF_ItemEsc = RPF_BigEsc
    Else
        RPF_ItemEsc = RPF_SmallEsc
    End If

    If StrComp(ModSz, "Large", vbTextCompare) = 0 Then
        RPF_ModEsc = RPF_BigEsc
    Else
        RPF_ModEsc = RPF_SmallEsc
    End If
    Exit Sub
ERH:
    ' On any error keep the defaults already set above.
    RPF_ItemEsc = RPF_SmallEsc
    RPF_ModEsc = RPF_SmallEsc
End Sub

' Wrap a menu-item line: apply the item size, print, reset to normal.
Public Function FormatItemLine(ByVal Text As String) As String
    FormatItemLine = RPF_ItemEsc & Text & RPF_SmallEsc & vbCrLf
End Function

' Wrap a modifier line: apply the modifier size, print, reset to normal.
Public Function FormatModifierLine(ByVal Text As String) As String
    FormatModifierLine = RPF_ModEsc & Text & RPF_SmallEsc & vbCrLf
End Function
