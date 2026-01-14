Attribute VB_Name = "ImportGUR_NAL_NEW"
Option Compare Database
Option Explicit

' ============================================================
' Access 2010 (VBA) - Import из Excel в таблицу GUR_NAL_NEW
'
' Как использовать:
' 1) Откройте Access -> Alt+F11 (VBA editor)
' 2) File -> Import File... -> выберите этот .bas
' 3) Запустите процедуру Import_GUR_NAL_NEW_FromExcel
'
' Логика:
' - Открывается FileDialog, пользователь выбирает Excel файл
' - Данные читаются через ADO в память (Collection строк)
' - Затем данные вставляются в таблицу GUR_NAL_NEW (DAO) в транзакции
'
' Требования к Excel:
' - 1-й ряд: заголовки колонок (HDR=YES)
' - Названия заголовков должны совпадать с именами полей таблицы GUR_NAL_NEW
'   (лишние колонки игнорируются, недостающие поля не заполняются)
' ============================================================

Public Sub Import_GUR_NAL_NEW_FromExcel()
    Dim filePath As String
    filePath = PickExcelFile()
    If Len(filePath) = 0 Then Exit Sub

    Dim rows As Collection
    Set rows = ReadExcelToMemory(filePath) 'читает 1-й лист автоматически

    WriteMemoryToTable rows, "GUR_NAL_NEW"

    MsgBox "Импорт завершён. Строк загружено: " & CStr(Application.Max(0, rows.Count - 1)), vbInformation
End Sub

Private Function PickExcelFile() As String
    On Error GoTo EH

    ' msoFileDialogFilePicker = 3 (чтобы не зависеть от констант Office)
    Dim fd As Object
    Set fd = Application.FileDialog(3)

    With fd
        .Title = "Выберите Excel-файл для импорта"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm;*.xlsb", 1
        If .Show <> -1 Then
            PickExcelFile = vbNullString
        Else
            PickExcelFile = .SelectedItems(1)
        End If
    End With

    Exit Function
EH:
    PickExcelFile = vbNullString
End Function

Private Function ReadExcelToMemory(ByVal excelPath As String, Optional ByVal sheetName As String = vbNullString) As Collection
    On Error GoTo EH

    Dim cn As Object 'ADODB.Connection
    Set cn = CreateObject("ADODB.Connection")
    cn.Open ExcelConnectionString(excelPath, True)

    If Len(sheetName) = 0 Then
        sheetName = GetFirstWorksheetName(cn)
    End If
    If Len(sheetName) = 0 Then Err.Raise vbObjectError + 1000, "ReadExcelToMemory", "Не найден лист в Excel-файле."

    Dim sql As String
    sql = "SELECT * FROM [" & sheetName & "]"

    Dim rs As Object 'ADODB.Recordset
    Set rs = CreateObject("ADODB.Recordset")

    ' adOpenForwardOnly=0, adLockReadOnly=1
    rs.Open sql, cn, 0, 1

    Dim rows As New Collection

    ' Храним индексы полей (по имени) в элементе с ключом "FIELDS"
    Dim fieldIndex As Object 'Scripting.Dictionary
    Set fieldIndex = CreateObject("Scripting.Dictionary")
    fieldIndex.CompareMode = 1 'TextCompare

    Dim i As Long
    For i = 0 To rs.Fields.Count - 1
        fieldIndex(rs.Fields(i).Name) = i
    Next i
    rows.Add fieldIndex, "FIELDS"

    Do While Not rs.EOF
        Dim rowArr() As Variant
        ReDim rowArr(0 To rs.Fields.Count - 1)
        For i = 0 To rs.Fields.Count - 1
            rowArr(i) = rs.Fields(i).Value
        Next i
        rows.Add rowArr
        rs.MoveNext
    Loop

    rs.Close
    cn.Close

    Set ReadExcelToMemory = rows
    Exit Function

EH:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State <> 0 Then rs.Close
    If Not cn Is Nothing Then If cn.State <> 0 Then cn.Close
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetFirstWorksheetName(ByVal cn As Object) As String
    On Error GoTo EH

    ' adSchemaTables = 20
    Dim rsSchema As Object
    Set rsSchema = cn.OpenSchema(20)

    Do While Not rsSchema.EOF
        Dim tbl As String
        tbl = Nz(rsSchema.Fields("TABLE_NAME").Value, vbNullString)

        ' Обычно листы выглядят как "Sheet1$" или "Лист1$"
        ' Диапазоны/именованные области могут быть "SomeRange" — их пропускаем
        If InStr(1, tbl, "$", vbTextCompare) > 0 Then
            ' Убираем возможные апострофы в конце, ADO иногда отдаёт "Sheet1$'"
            If Right$(tbl, 1) = "'" Then tbl = Left$(tbl, Len(tbl) - 1)
            GetFirstWorksheetName = tbl
            Exit Do
        End If
        rsSchema.MoveNext
    Loop

    rsSchema.Close
    Exit Function

EH:
    GetFirstWorksheetName = vbNullString
End Function

Private Function ExcelConnectionString(ByVal excelPath As String, ByVal hasHeaders As Boolean) As String
    Dim ext As String
    ext = LCase$(Mid$(excelPath, InStrRev(excelPath, ".") + 1))

    Dim hdr As String
    hdr = IIf(hasHeaders, "YES", "NO")

    Dim props As String
    Select Case ext
        Case "xls"
            props = "Excel 8.0;HDR=" & hdr & ";IMEX=1"
        Case Else
            ' xlsx/xlsm/xlsb — чаще всего корректно работает с "Excel 12.0 Xml"
            props = "Excel 12.0 Xml;HDR=" & hdr & ";IMEX=1"
    End Select

    ExcelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & excelPath & ";Extended Properties=""" & props & """;"
End Function

Private Sub WriteMemoryToTable(ByVal rows As Collection, ByVal targetTable As String)
    If rows Is Nothing Then Exit Sub
    If rows.Count = 0 Then Exit Sub

    On Error GoTo EH

    Dim fieldIndex As Object
    Set fieldIndex = rows("FIELDS")

    Dim db As DAO.Database
    Set db = CurrentDb

    ' Множество полей таблицы (чтобы игнорировать лишние колонки Excel)
    Dim tableFields As Object 'Scripting.Dictionary
    Set tableFields = CreateObject("Scripting.Dictionary")
    tableFields.CompareMode = 1 'TextCompare

    Dim td As DAO.TableDef
    Set td = db.TableDefs(targetTable)

    Dim f As DAO.Field
    For Each f In td.Fields
        tableFields(f.Name) = True
    Next f

    db.BeginTrans

    Dim rsT As DAO.Recordset
    Set rsT = db.OpenRecordset(targetTable, dbOpenDynaset, dbAppendOnly)

    Dim rowNum As Long
    For rowNum = 2 To rows.Count '1-й элемент — "FIELDS"
        Dim rowArr As Variant
        rowArr = rows(rowNum)

        rsT.AddNew

        Dim colName As Variant
        For Each colName In fieldIndex.Keys
            If tableFields.Exists(CStr(colName)) Then
                Dim v As Variant
                v = rowArr(fieldIndex(CStr(colName)))

                ' Унифицируем пустые значения
                If IsEmpty(v) Then
                    v = Null
                ElseIf VarType(v) = vbString Then
                    If Len(Trim$(CStr(v))) = 0 Then v = Null
                End If

                rsT.Fields(CStr(colName)).Value = v
            End If
        Next colName

        rsT.Update
    Next rowNum

    rsT.Close
    db.CommitTrans
    Exit Sub

EH:
    On Error Resume Next
    db.Rollback
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

