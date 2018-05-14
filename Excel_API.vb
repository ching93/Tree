Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Public Class Excel_API
    Public xlApp As Excel.Application ' = New Microsoft.Office.Interop.Excel.Application()
    Public WB As Excel.Workbook
    Public WS As Excel.Worksheet
    Public WS_range As Excel.Range

    Dim uchet As Integer
    Sub Create_EX_Doc(Visible As Double)
        xlApp = New Excel.Application()
        xlApp.Visible = Visible
        WB = xlApp.Workbooks.Add(1)
        WS = WB.Sheets(1)
    End Sub
    Sub Open_EX_Doc(Visible As Double, excel_file_path As String)
        xlApp = New Excel.Application()
        xlApp.Visible = Visible
        WB = xlApp.Workbooks.Open(excel_file_path)
        WS = WB.Sheets.Item(1)
    End Sub
    Sub exc_close()
        'Me.Application.Workbooks("NewWorkbook.xlsx").Close(SaveChanges:=False)
        'WB.Close()
        'releaseObject(WS)
        'releaseObject(WB)
        'releaseObject(xlApp)
        'WB.ActiveWindow.Close(SaveChanges:=False)
        'WS.Quit

        'WB = Nothinge
        Try
            xlApp.Quit()
            xlApp = Nothing
        Catch ex As Exception

        End Try
    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Sub Create_WBk(WB_Name As String)
        If WB_Name = "" Then WB_Name = "Книга1"
        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Add(WB_Name)
    End Sub
    Sub Create_WSh(WSh_Name As String)
        If WSh_Name = "" Then WSh_Name = "Лист1"
        Dim xlWorkBook As Excel.Workbook = WB.Add(WSh_Name)
    End Sub
    Function Get_LastRowInOneColumn(SheetName As String, Column_Index As Integer) As Long
        'возвращает последнюю использованную строку в столбце: требуется задать имя листа(SheetName) и имя столбца(Column_Name)
        Dim LastRow As Long
        Dim ActiveSheet As Worksheet = WB.Sheets.Item(SheetName)
        With ActiveSheet
            'LastRow = .Cells(.Rows.Count, 1).End(Excel.XlDirection.xlUp).Row
            LastRow = .Cells(.Rows.Count, Column_Index).End(Excel.XlDirection.xlUp).Row
        End With
        Return LastRow
    End Function
    Function Get_Last_Column(SheetName As String) As Long
        'возвращает последнюю использованную строку в столбце: требуется задать имя листа(SheetName) и имя столбца(Column_Name)
        Dim LastColumn As Long
        Dim ActiveSheet As Worksheet = WB.Sheets.Item(SheetName)
        With ActiveSheet
            LastColumn = .Cells(1, .Columns.Count).End(Excel.XlDirection.xlToLeft).Column
        End With
        Return LastColumn
    End Function
    Function get_Value_From_Cell(SheetName As String, columnIndex As Integer, row_index As Integer)
        Dim ActiveSheet As Worksheet = WB.Sheets.Item(SheetName)
        Dim cell_value As String
        Try
            With ActiveSheet
                cell_value = .Cells(row_index, columnIndex).Value()
                Return cell_value
            End With
        Catch ex As Exception

        End Try
    End Function
    Sub set_Value_From_Cell(SheetName As String, columnIndex As Integer, row_index As Integer, Text_in_Cells As String)
        Dim ActiveSheet As Worksheet = WB.Sheets.Item(SheetName)
        Try
            With ActiveSheet
                .Cells(row_index, columnIndex) = Text_in_Cells
            End With
        Catch ex As Exception

        End Try
    End Sub

    Function get_value_bay_FindText(SheetName As String, columnName As String, row_Name As String, FindText As String)
        Dim ActiveSheet As Worksheet = WB.Sheets.Item(SheetName)
        Dim cell_value As String
        With ActiveSheet

        End With

        Dim currentFind As Excel.Range = Nothing
        Dim firstFind As Excel.Range = Nothing
        Try
            Dim Fruits As Excel.Range = ActiveSheet.Range(columnName, row_Name)
            currentFind = Fruits.Find(FindText, ,
            Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
            Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False)
            While Not currentFind Is Nothing

                ' Keep track of the first range you find.
                If firstFind Is Nothing Then
                    firstFind = currentFind
                    ' If you didn't move to a new range, you are done.
                ElseIf currentFind.Address = firstFind.Address Then
                    Exit While
                End If
                If UCase(ActiveSheet.Cells(currentFind.Row, currentFind.Column).Value) = UCase(FindText) Then
                    Return currentFind.Row
                Else
                    uchet = uchet + 1
                    If uchet < 3 Then
                        Call get_value_bay_FindText(SheetName, "B" & currentFind.Row, row_Name, FindText)
                    End If
                End If
                    'With currentFind.Font
                    '    .Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
                    '    .Bold = True
                    'End With

                    currentFind = Fruits.FindNext(currentFind)
            End While
        Catch ex As Exception

        End Try
    End Function

    Function get_value_bay_FindText_Strong(SheetName As String, columnIndex As Integer, row_index As Integer, FindText As String)
        Dim ActiveSheet As Worksheet = WB.Sheets.Item(SheetName)
        Dim lastrow As Integer = Get_LastRowInOneColumn(SheetName, columnIndex)
        Try
            With ActiveSheet
                While row_index < lastrow
                    If .Cells(row_index, columnIndex).Value = FindText Then
                        Exit While
                    End If
                    row_index = row_index + 1
                End While
            End With
            Return row_index
        Catch ex As Exception

        End Try
    End Function
    Public Sub Сreate_Headers(row As Integer, col As Integer, htext As Integer, cell1 As String, cell2 As String,
        mergeColumns As Integer, b As String, font As Boolean, size As Integer, fcolor As String)
        WS.Cells(row, col) = htext
        WS_range = WS.get_Range(cell1, cell2)
        WS_range.Merge(mergeColumns)
        'процедура создает новый заголовок!
        'Входные параметры:
        'row, col -         индексы строк и столбцов, htext - текст в заголовке
        'cell1, cell2 -     Это будет использоваться, чтобы указать, какие ячейки мы будем использовать, например.A1: B1
        'mergeColumns -     содержит количество ячеек, которые мы хотим объединить в ячейке
        'b -                цвет фона для выбранной ячейки
        'font -             True или False для шрифта текста в выбранной ячейке
        'size -             указать размер ячейки
        'fcolor-            указать цветной шрифт
        Select Case b
            Case "желтый"
                WS_range.Interior.Color = System.Drawing.Color.Yellow.ToArgb()
                Exit Select
            Case "серый"
                WS_range.Interior.Color = System.Drawing.Color.Gray.ToArgb()
                Exit Select
            Case "зеленый"
                WS_range.Interior.Color = System.Drawing.Color.Green.ToArgb()
                Exit Select
            Case "красный"
                WS_range.Interior.Color = System.Drawing.Color.Red.ToArgb()
                Exit Select
            Case "PeachPuff"
                WS_range.Interior.Color = System.Drawing.Color.PeachPuff.ToArgb()
                Exit Select
            Case Else
                WS_range.Interior.Color = System.Drawing.Color.White.ToArgb()
                Exit Select
        End Select
        WS_range.Borders.Color = System.Drawing.Color.Black.ToArgb()
        WS_range.Font.Bold = font
        WS_range.ColumnWidth = size
        If CObj(fcolor).Equals("") Then
            WS_range.Font.Color = System.Drawing.Color.White.ToArgb()
        Else
            WS_range.Font.Color = System.Drawing.Color.Black.ToArgb
        End If


    End Sub




























    'Module ExcelCode

    ''' <summary>
    ''' Get last used row in sheetname
    ''' </summary>
    ''' <param name="FileName">path and filename to excel file to work with</param>
    ''' <param name="SheetName">Worksheet name to get information</param>
    ''' <returns>-1 if issues else lasted used row</returns>
    ''' <remarks></remarks>
    Public Function UsedRows(ByVal FileName As String, ByVal SheetName As String) As Integer

            Dim RowsUsed As Integer = -1

            If IO.File.Exists(FileName) Then
                Dim xlApp As Excel.Application = Nothing
                Dim xlWorkBooks As Excel.Workbooks = Nothing
                Dim xlWorkBook As Excel.Workbook = Nothing
                Dim xlWorkSheet As Excel.Worksheet = Nothing
                Dim xlWorkSheets As Excel.Sheets = Nothing

                xlApp = New Excel.Application
                xlApp.DisplayAlerts = False
                xlWorkBooks = xlApp.Workbooks
                xlWorkBook = xlWorkBooks.Open(FileName)

                xlApp.Visible = False

                xlWorkSheets = xlWorkBook.Sheets

                For x As Integer = 1 To xlWorkSheets.Count

                    xlWorkSheet = CType(xlWorkSheets(x), Excel.Worksheet)

                    If xlWorkSheet.Name = SheetName Then
                        Dim xlCells As Excel.Range = Nothing
                        xlCells = xlWorkSheet.Cells

                        Dim thisRange As Excel.Range = xlCells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell)

                        RowsUsed = thisRange.Row
                        Marshal.FinalReleaseComObject(thisRange)
                        thisRange = Nothing

                        Marshal.FinalReleaseComObject(xlCells)
                        xlCells = Nothing

                        Exit For
                    End If

                    Marshal.FinalReleaseComObject(xlWorkSheet)
                    xlWorkSheet = Nothing

                Next

                xlWorkBook.Close()
                xlApp.UserControl = True
                xlApp.Quit()

                ReleaseComObject(xlWorkSheets)
                ReleaseComObject(xlWorkSheet)
                ReleaseComObject(xlWorkBook)
                ReleaseComObject(xlWorkBooks)
                ReleaseComObject(xlApp)
            Else
                Throw New Exception("'" & FileName & "' not found.")
            End If

            Return RowsUsed

        End Function
        ''' <summary>
        ''' Get last used row for a single column
        ''' </summary>
        ''' <param name="FileName"></param>
        ''' <param name="SheetName"></param>
        ''' <param name="Column"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function UseRowsdByColumn(ByVal FileName As String, ByVal SheetName As String, ByVal Column As String) As Integer
            Dim LastRowCount As Integer = 1

            Dim xlApp As Excel.Application = Nothing
            Dim xlWorkBooks As Excel.Workbooks = Nothing
            Dim xlWorkBook As Excel.Workbook = Nothing
            Dim xlWorkSheet As Excel.Worksheet = Nothing
            Dim xlWorkSheets As Excel.Sheets = Nothing

            xlApp = New Excel.Application
            xlApp.DisplayAlerts = False
            xlWorkBooks = xlApp.Workbooks
            xlWorkBook = xlWorkBooks.Open(FileName)

            xlApp.Visible = False

            xlWorkSheets = xlWorkBook.Sheets

            For x As Integer = 1 To xlWorkSheets.Count

                xlWorkSheet = CType(xlWorkSheets(x), Excel.Worksheet)

                If xlWorkSheet.Name = SheetName Then

                    Dim xlCells As Excel.Range = xlWorkSheet.Cells()
                    Dim xlTempRange1 As Excel.Range = xlCells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell)
                    Dim xlTempRange2 As Excel.Range = xlWorkSheet.Rows


                    Dim xlTempRange3 As Excel.Range = xlWorkSheet.Range(Column.ToUpper & xlTempRange2.Count)
                    Dim xlTempRange4 As Excel.Range = xlTempRange3.End(Excel.XlDirection.xlUp)

                    LastRowCount = xlTempRange4.Row

                    Marshal.FinalReleaseComObject(xlTempRange4)
                    xlTempRange4 = Nothing

                    Marshal.FinalReleaseComObject(xlTempRange3)
                    xlTempRange3 = Nothing

                    Marshal.FinalReleaseComObject(xlTempRange2)
                    xlTempRange2 = Nothing

                    Marshal.FinalReleaseComObject(xlTempRange1)
                    xlTempRange1 = Nothing

                    Marshal.FinalReleaseComObject(xlCells)
                    xlCells = Nothing

                End If

                Marshal.FinalReleaseComObject(xlWorkSheet)
                xlWorkSheet = Nothing

            Next

            xlWorkBook.Close()
            xlApp.UserControl = True
            xlApp.Quit()

            ReleaseComObject(xlWorkSheets)
            ReleaseComObject(xlWorkSheet)
            ReleaseComObject(xlWorkBook)
            ReleaseComObject(xlWorkBooks)
            ReleaseComObject(xlApp)

            Return LastRowCount

        End Function

        Public Sub ReleaseComObject(ByVal obj As Object)
            Try
                If obj IsNot Nothing Then
                    Marshal.ReleaseComObject(obj)
                    obj = Nothing
                End If
            Catch ex As Exception
                obj = Nothing
            End Try
        End Sub

    'End Module
End Class
