Imports System.ComponentModel
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports System.Xml
Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Public Class Form1
    Public myExcel As Excel.Application
    Public myBgColorList As List(Of Integer)
    Public myFtColorList As List(Of Integer)
    'Public myColorList As BindingList(Of HLColor)
    Public myXlsxFileList As BindingList(Of String)
    Public myXmlFileList As BindingList(Of String)

    Public myBgColorFlag As Boolean
    Public myFtColorFlag As Boolean
    Public myBgColorOut As Boolean
    Public myFtColorOut As Boolean
    Public myShapeEnable As Boolean

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = ProductName & " Ver." & ProductVersion
        'myColorList = New BindingList(Of HLColor)
        'ColorSet(myColorList)
        'Cmb_Color.DataSource = myColorList
        'Cmb_Color.DisplayMember = "Name"
        'Cmb_Color.SelectedIndex = My.Settings.colorindex
        'Rbn_Highlight.Checked = My.Settings.highlight
        'Pcb_BgColor.BackColor = ColorTranslator.FromHtml("#" & myColorList(Cmb_Color.SelectedIndex).ColorHex)
        If My.Settings.proctype = "bg" Then
            Rbn_BgColor.Checked = True
        ElseIf My.Settings.proctype = "ft" Then
            Rbn_FntColor.Checked = True
        Else
            Rbn_All.Checked = True
        End If
        Nud_BgColorR.Value = My.Settings.bgcolor_r
        Nud_BgColorG.Value = My.Settings.bgcolor_g
        Nud_BgColorB.Value = My.Settings.bgcolor_b
        Nud_FtColorR.Value = My.Settings.ftcolor_r
        Nud_FtColorG.Value = My.Settings.ftcolor_g
        Nud_FtColorB.Value = My.Settings.ftcolor_b

        Pcb_BgColor.BackColor = Color.FromArgb(Nud_BgColorR.Value, Nud_BgColorG.Value, Nud_BgColorB.Value)
        Pcb_FtColor.BackColor = Color.FromArgb(Nud_FtColorR.Value, Nud_FtColorG.Value, Nud_FtColorB.Value)

        'If Rbn_BgColor.Checked = True Then
        '    Lbl_Color.Enabled = True
        '    'Cmb_Color.Enabled = True
        '    Pcb_BgColor.Enabled = True
        'Else
        '    Lbl_Color.Enabled = False
        '    'Cmb_Color.Enabled = False
        '    Pcb_BgColor.Enabled = False
        'End If
        myBgColorList = JsonConvert.DeserializeObject(Of List(Of Integer))(My.Settings.bgcolorlist)
        myFtColorList = JsonConvert.DeserializeObject(Of List(Of Integer))(My.Settings.ftcolorlist)

        If myBgColorList Is Nothing Then
            myBgColorList = New List(Of Integer) From {&H33, &H66, &H99, &HCC, &H3300, &H3333, &H3366, &H3399, &H33CC, &H6600, &H6633, &H6666, &H6699, &H66CC, &H9900, &H9933}
        End If
        If myFtColorList Is Nothing Then
            myFtColorList = New List(Of Integer) From {&H33, &H66, &H99, &HCC, &H3300, &H3333, &H3366, &H3399, &H33CC, &H6600, &H6633, &H6666, &H6699, &H66CC, &H9900, &H9933}
        End If

        If My.Settings.ftout = True Then
            Rbn_FtExport.Checked = True
        Else
            Rbn_FtExclusion.Checked = True
        End If
        If My.Settings.bgout = True Then
            Rbn_BgExport.Checked = True
        Else
            Rbn_BgExclusion.Checked = True
        End If

        Chk_DollarChar.Checked = My.Settings.dollar
        Chk_HiddenSheet.Checked = My.Settings.hiddensheet
        Chk_IncludeShape.Checked = My.Settings.includeshapes
        'Dim list1 = New List(Of Integer) From {1, 2}
        'Dim json1 = JsonConvert.SerializeObject(list1)
        'Dim list2 As List(Of Integer) = JsonConvert.DeserializeObject(Of List(Of Integer))(json1)
        'Debug.WriteLine(json1)
        StandardColorSet()

        myXlsxFileList = New BindingList(Of String)
        Lbx_Xlsx.DataSource = myXlsxFileList
        myXmlFileList = New BindingList(Of String)
        Lbx_Xml.DataSource = myXmlFileList

    End Sub
    Private Sub StandardColorSet()
        Pcb_BgCSet1.BackColor = Color.FromArgb(192, 0, 0)
        Pcb_BgCSet2.BackColor = Color.FromArgb(255, 0, 0)
        Pcb_BgCSet3.BackColor = Color.FromArgb(255, 192, 0)
        Pcb_BgCSet4.BackColor = Color.FromArgb(255, 255, 0)
        Pcb_BgCSet5.BackColor = Color.FromArgb(146, 208, 80)
        Pcb_BgCSet6.BackColor = Color.FromArgb(0, 176, 80)
        Pcb_BgCSet7.BackColor = Color.FromArgb(0, 176, 240)
        Pcb_BgCSet8.BackColor = Color.FromArgb(0, 112, 192)
        Pcb_BgCSet9.BackColor = Color.FromArgb(0, 32, 96)
        Pcb_BgCSet10.BackColor = Color.FromArgb(112, 48, 160)
        Pcb_FtCSet1.BackColor = Color.FromArgb(192, 0, 0)
        Pcb_FtCSet2.BackColor = Color.FromArgb(255, 0, 0)
        Pcb_FtCSet3.BackColor = Color.FromArgb(255, 192, 0)
        Pcb_FtCSet4.BackColor = Color.FromArgb(255, 255, 0)
        Pcb_FtCSet5.BackColor = Color.FromArgb(146, 208, 80)
        Pcb_FtCSet6.BackColor = Color.FromArgb(0, 176, 80)
        Pcb_FtCSet7.BackColor = Color.FromArgb(0, 176, 240)
        Pcb_FtCSet8.BackColor = Color.FromArgb(0, 112, 192)
        Pcb_FtCSet9.BackColor = Color.FromArgb(0, 32, 96)
        Pcb_FtCSet10.BackColor = Color.FromArgb(112, 48, 160)
    End Sub
    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        'My.Settings.colorindex = Cmb_Color.SelectedIndex
        If Rbn_BgColor.Checked = True Then
            My.Settings.proctype = "bg"
        ElseIf Rbn_FntColor.Checked = True Then
            My.Settings.proctype = "ft"
        Else
            My.Settings.proctype = "all"
        End If
        My.Settings.bgcolor_r = Nud_BgColorR.Value
        My.Settings.bgcolor_g = Nud_BgColorG.Value
        My.Settings.bgcolor_b = Nud_BgColorB.Value
        My.Settings.ftcolor_r = Nud_FtColorR.Value
        My.Settings.ftcolor_g = Nud_FtColorG.Value
        My.Settings.ftcolor_b = Nud_FtColorB.Value
        My.Settings.bgcolorlist = JsonConvert.SerializeObject(myBgColorList)
        My.Settings.ftcolorlist = JsonConvert.SerializeObject(myFtColorList)
        If Rbn_BgExport.Checked = True Then
            My.Settings.bgout = True
        Else
            My.Settings.bgout = False
        End If
        If Rbn_FtExport.Checked = True Then
            My.Settings.ftout = True
        Else
            My.Settings.ftout = False
        End If
        My.Settings.dollar = Chk_DollarChar.Checked
        My.Settings.hiddensheet = Chk_HiddenSheet.Checked
        My.Settings.includeshapes = Chk_IncludeShape.Checked
    End Sub




    'Private Sub ColorSet(ByRef myColorList As BindingList(Of HLColor))
    '    myColorList.Add(New HLColor With {.Name = "濃い赤", .ColorHex = "C00000"})
    '    myColorList.Add(New HLColor With {.Name = "赤", .ColorHex = "FF0000"})
    '    myColorList.Add(New HLColor With {.Name = "オレンジ", .ColorHex = "FFC000"})
    '    myColorList.Add(New HLColor With {.Name = "黄", .ColorHex = "FFFF00"})
    '    myColorList.Add(New HLColor With {.Name = "薄い緑", .ColorHex = "92D050"})
    '    myColorList.Add(New HLColor With {.Name = "緑", .ColorHex = "00B050"})
    '    myColorList.Add(New HLColor With {.Name = "薄い青", .ColorHex = "00B0F0"})
    '    myColorList.Add(New HLColor With {.Name = "青", .ColorHex = "0070C0"})
    '    myColorList.Add(New HLColor With {.Name = "濃い青", .ColorHex = "002060"})
    '    myColorList.Add(New HLColor With {.Name = "紫", .ColorHex = "7030A0"})
    'End Sub
    'Private Sub Cmb_Color_SelectedIndexChanged(sender As Object, e As EventArgs)
    '    Pcb_BgColor.BackColor = ColorTranslator.FromHtml("#" & myColorList(Cmb_Color.SelectedIndex).ColorHex)
    'End Sub

    Private Sub Lbx_Xlsx_DragEnter(sender As Object, e As DragEventArgs) Handles Lbx_Xlsx.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub Lbx_Xlsx_DragDrop(sender As Object, e As DragEventArgs) Handles Lbx_Xlsx.DragDrop
        Dim fileNames As String() = CType(e.Data.GetData(DataFormats.FileDrop, False), String())
        Dim fn As String
        For Each fn In fileNames
            If Path.GetExtension(fn).ToLower = ".xlsx" Then
                If myXlsxFileList.Contains(fn) = False Then myXlsxFileList.Add(fn)
            End If
        Next
    End Sub

    Private Sub Tbx_OutputFolder_DragEnter(sender As Object, e As DragEventArgs) Handles Tbx_OutputFolder.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Tbx_OutputFolder_DragDrop(sender As Object, e As DragEventArgs) Handles Tbx_OutputFolder.DragDrop
        Tbx_OutputFolder.Text = e.Data.GetData(DataFormats.FileDrop)(0)
        If Directory.Exists(Tbx_OutputFolder.Text) = False Then
            MsgBox("正しいフォルダを選択してください！", MsgBoxStyle.Exclamation)
            Tbx_OutputFolder.Text = ""
        End If
    End Sub

    Private Sub Lbx_Xml_DragEnter(sender As Object, e As DragEventArgs) Handles Lbx_Xml.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub Lbx_Xml_DragDrop(sender As Object, e As DragEventArgs) Handles Lbx_Xml.DragDrop
        Dim fileNames As String() = CType(e.Data.GetData(DataFormats.FileDrop, False), String())
        Dim fn As String
        For Each fn In fileNames
            If Path.GetExtension(fn).ToLower = ".xml" Then
                If myXmlFileList.Contains(fn) = False Then myXmlFileList.Add(fn)
            End If
        Next
    End Sub

    Private Sub Tbx_OutputFolder2_DragEnter(sender As Object, e As DragEventArgs) Handles Tbx_OutputFolder2.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub Tbx_OutputFolder2_DragDrop(sender As Object, e As DragEventArgs) Handles Tbx_OutputFolder2.DragDrop
        Tbx_OutputFolder2.Text = e.Data.GetData(DataFormats.FileDrop)(0)
        If Directory.Exists(Tbx_OutputFolder2.Text) = False Then
            MsgBox("正しいフォルダを選択してください！", MsgBoxStyle.Exclamation)
            Tbx_OutputFolder2.Text = ""
        End If
    End Sub


    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub Btn_Export_Click(sender As Object, e As EventArgs) Handles Btn_Export.Click
        Me.Enabled = False
        Dim myOutputPath As String = Tbx_OutputFolder.Text
        If Directory.Exists(myOutputPath) Then
            If myXlsxFileList.Count > 0 Then
                myExcel = New Excel.Application With {.Visible = False}
                Dim myExBookList As New List(Of ExBook)
                For i = 0 To myXlsxFileList.Count - 1
                    Lbx_Xlsx.SelectedIndex = i
                    Dim myXlsxFilePath As String = myXlsxFileList(i).Trim()

                    If File.Exists(myXlsxFilePath) AndAlso Path.GetExtension(myXlsxFilePath).ToLower = ".xlsx" Then
                        myBgColorFlag = Rbn_BgColor.Checked
                        myFtColorFlag = Rbn_FntColor.Checked
                        myBgColorOut = Rbn_BgExport.Checked
                        myFtColorOut = Rbn_FtExport.Checked
                        myShapeEnable = Chk_IncludeShape.Checked

                        Dim myBook As Excel.Workbook = myExcel.Workbooks.Open(myXlsxFilePath)
                        Dim myExBook As New ExBook With {.Name = myBook.Name}
                        'myExBook.Name = myBook.Name
                        Dim myExSheets As New List(Of ExSheet)
                        For Each mySheet As Excel.Worksheet In myBook.Worksheets
                            If Chk_HiddenSheet.Checked = True OrElse (mySheet.Visible <> XlSheetVisibility.xlSheetHidden AndAlso mySheet.Visible <> XlSheetVisibility.xlSheetVeryHidden) Then
                                'mySheet.Select()
                                Dim myExSheet As New ExSheet With {.Name = mySheet.Name, .Index = mySheet.Index}
                                'myExSheet.Name = mySheet.Name
                                'myExSheet.Index = mySheet.Index
                                Dim myExCells As New List(Of ExCell)

                                'Dim myRange As Range = mySheet.UsedRange
                                'Dim startCol As Integer = myRange.Column
                                'Dim lastCol As Integer = myRange.Columns.Count + startCol - 1

                                'Dim startRow As Integer = myRange.Row
                                'Dim lastRow As Integer = startRow

                                'For y = startCol To lastCol
                                '    Dim colLastRow = mySheet.Cells(mySheet.Rows.Count, y).End(XlDirection.xlUp).Row
                                '    If lastRow < colLastRow Then lastRow = colLastRow
                                'Next

                                Dim myRange As Excel.Range = mySheet.UsedRange
                                Dim myStartRow As Integer = myRange.Row
                                Dim myStartCol As Integer = myRange.Column
                                Dim myLastRow As Integer = myStartRow + myRange.Rows.Count - 1
                                Dim myLastCol As Integer = myStartCol + myRange.Columns.Count - 1

                                Dim startRow As Integer = myStartRow
                                Dim startcol As Integer = myStartCol
                                Dim lastRow As Integer = 1
                                Dim lastCol As Integer = 1
                                ' 実際にデータが入力されている範囲を特定
                                Dim emptyColCount As Integer = 0
                                For y = myStartCol To myLastCol
                                    Dim colLastRow = mySheet.Cells(mySheet.Rows.Count, y).End(XlDirection.xlUp).Row
                                    If lastRow < colLastRow Then lastRow = colLastRow
                                    If colLastRow = 1 Then
                                        emptyColCount += 1
                                    Else
                                        emptyColCount = 0
                                    End If
                                    If emptyColCount > 100 Then
                                        Exit For
                                    End If
                                Next

                                Dim emptyRowCount As Integer = 0
                                For x = myStartRow To myLastRow
                                    Dim rowLastCol = mySheet.Cells(x, mySheet.Columns.Count).End(XlDirection.xlToLeft).column
                                    If lastCol < rowLastCol Then lastCol = rowLastCol
                                    If rowLastCol = 1 Then
                                        emptyRowCount += 1
                                    Else
                                        emptyRowCount = 0
                                    End If
                                    If emptyRowCount > 100 Then
                                        Exit For
                                    End If
                                Next

                                For Each myCell As Range In mySheet.Range(mySheet.Cells(startRow, startCol), mySheet.Cells(lastRow, lastCol))
                                    Dim myBgColor As Long = Nud_BgColorB.Value * 16 ^ 4 + Nud_BgColorG.Value * 16 ^ 2 + Nud_BgColorR.Value
                                    Dim myFtColor As Long = Nud_FtColorB.Value * 16 ^ 4 + Nud_FtColorG.Value * 16 ^ 2 + Nud_FtColorR.Value
                                    'Dim wkColorHex As String = "" 'myColorList(Cmb_Color.SelectedIndex).ColorHex
                                    'wkColorHex = wkColorHex.Substring(4, 2) & wkColorHex.Substring(2, 2) & wkColorHex.Substring(0, 2)

                                    Try
                                        If myBgColorFlag = True Then
                                            If myCell.Interior.Color = myBgColor AndAlso myBgColorOut = True Then
                                                If TypeOf myCell.Value Is String AndAlso myCell.Value <> "" AndAlso Not myCell.HasFormula Then
                                                    Dim myExCell As New ExCell With {.X = myCell.Row, .Y = myCell.Column, .Text = myCell.Value.Replace(vbLf, "\n").Replace(vbVerticalTab, "\n"), .IsFormula = myCell.HasFormula}
                                                    myExCells.Add(myExCell)
                                                ElseIf myCell.HasFormula Then
                                                    Dim myExCell As New ExCell With {.X = myCell.Row, .Y = myCell.Column, .Text = myCell.Formula, .IsFormula = myCell.HasFormula}
                                                    myExCells.Add(myExCell)
                                                End If
                                            ElseIf myCell.Interior.Color <> myBgColor AndAlso myBgColorOut = False Then
                                                If TypeOf myCell.Value Is String AndAlso myCell.Value <> "" AndAlso Not myCell.HasFormula Then
                                                    Dim myExCell As New ExCell With {.X = myCell.Row, .Y = myCell.Column, .Text = myCell.Value.Replace(vbLf, "\n").Replace(vbVerticalTab, "\n"), .IsFormula = myCell.HasFormula}
                                                    myExCells.Add(myExCell)
                                                ElseIf myCell.HasFormula Then
                                                    Dim myExCell As New ExCell With {.X = myCell.Row, .Y = myCell.Column, .Text = myCell.Formula, .IsFormula = myCell.HasFormula}
                                                    myExCells.Add(myExCell)
                                                End If
                                            End If
                                        ElseIf myFtColorFlag = True Then
                                            If myCell.DisplayFormat.Font.Color = myFtColor AndAlso myFtColorOut = True Then
                                                If TypeOf myCell.Value Is String AndAlso myCell.Value <> "" AndAlso Not myCell.HasFormula Then
                                                    Dim myExCell As New ExCell With {.X = myCell.Row, .Y = myCell.Column, .Text = myCell.Value.Replace(vbLf, "\n").Replace(vbVerticalTab, "\n"), .IsFormula = myCell.HasFormula}
                                                    myExCells.Add(myExCell)
                                                ElseIf myCell.HasFormula Then
                                                    Dim myExCell As New ExCell With {.X = myCell.Row, .Y = myCell.Column, .Text = myCell.Formula, .IsFormula = myCell.HasFormula}
                                                    myExCells.Add(myExCell)
                                                End If
                                            ElseIf myCell.DisplayFormat.Font.Color <> myFtColor AndAlso myFtColorOut = False Then
                                                If TypeOf myCell.Value Is String AndAlso myCell.Value <> "" AndAlso Not myCell.HasFormula Then
                                                    Dim myExCell As New ExCell With {.X = myCell.Row, .Y = myCell.Column, .Text = myCell.Value.Replace(vbLf, "\n").Replace(vbVerticalTab, "\n"), .IsFormula = myCell.HasFormula}
                                                    myExCells.Add(myExCell)
                                                ElseIf myCell.HasFormula Then
                                                    Dim myExCell As New ExCell With {.X = myCell.Row, .Y = myCell.Column, .Text = myCell.Formula, .IsFormula = myCell.HasFormula}
                                                    myExCells.Add(myExCell)
                                                End If
                                            End If
                                        Else
                                            If TypeOf myCell.Value Is String AndAlso myCell.Value <> "" AndAlso Not myCell.HasFormula Then
                                                Dim myExCell As New ExCell With {.X = myCell.Row, .Y = myCell.Column, .Text = myCell.Value.Replace(vbLf, "\n").Replace(vbVerticalTab, "\n"), .IsFormula = myCell.HasFormula}
                                                myExCells.Add(myExCell)
                                            ElseIf myCell.HasFormula Then
                                                Dim myExCell As New ExCell With {.X = myCell.Row, .Y = myCell.Column, .Text = myCell.Formula, .IsFormula = myCell.HasFormula}
                                                myExCells.Add(myExCell)
                                            End If
                                        End If
                                    Catch ex As Exception

                                    End Try
                                    Lbl_Progress.Text = myExSheet.Name & ", " & myCell.Row
                                    Debug.WriteLine("")
                                Next
                                myExSheet.Cells = myExCells

                                If myShapeEnable = True Then    ' 図対象？
                                    Dim myExShapes As New List(Of ExShape)
                                    Dim ShapeNum As Integer = 1
                                    For Each myShape As Excel.Shape In mySheet.Shapes
                                        Dim myExShape As New ExShape
                                        myExShape.No = ShapeNum
                                        ShapeNum += 1
                                        ShapeNum = ShapeProc(myShape, myExShape, ShapeNum)
                                        myExShapes.Add(myExShape)
                                    Next
                                    myExSheet.Shapes = myExShapes
                                End If
                                myExSheets.Add(myExSheet)
                                ReleaseObject(mySheet)
                            End If
                        Next
                        myExBook.Sheets = myExSheets
                        myExBookList.Add(myExBook)
                        ExportXML(myExBook, myXlsxFilePath, myOutputPath)
                        myBook.Close()
                        ReleaseObject(myBook)
                    Else
                        MessageBox.Show("正しいファイルを選択してください！")
                    End If
                Next

                For Each myExBook As ExBook In myExBookList
                    For Each myExSheet In myExBook.Sheets
                        myExBook.LogChart = myExBook.LogChart & "シート名：" & myExSheet.Name & vbCrLf
                        myExBook.LogFormula = myExBook.LogFormula & "シート名：" & myExSheet.Name & vbCrLf
                        myExBook.LogComment = myExBook.LogComment & "シート名：" & myExSheet.Name & vbCrLf

                        Dim myFormulaCellList As List(Of ExCell) = myExSheet.Cells.Where(Function(n) n.IsFormula = True).ToList
                        Dim myFormulaCnt As Integer = myFormulaCellList.Count
                        For Each myExCell In myFormulaCellList
                            myExBook.LogFormula = myExBook.LogFormula & myExCell.X & vbTab & myExCell.Y & vbTab & myExCell.Text & vbCrLf
                        Next
                        myExBook.CntFormula = myExBook.CntFormula + myFormulaCnt

                        If myExSheet.Shapes IsNot Nothing AndAlso myExSheet.Shapes.Count > 0 Then
                            Dim myChartList As List(Of ExShape) = myExSheet.Shapes.Where(Function(n) n.Type = 3).ToList
                            Dim myChartCnt As Integer = myChartList.Count
                            For Each myExChart In myChartList
                                If myExChart.Paragraphs.Count > 0 Then
                                    myExBook.LogChart = myExBook.LogChart & vbTab & "チャート：" & vbTab & myExChart.Paragraphs(0).Text & vbCrLf
                                End If
                            Next
                            myExBook.CntChart = myExBook.CntChart + myChartCnt

                            Dim myCommentList As List(Of ExShape) = myExSheet.Shapes.Where(Function(n) n.Type = 4).ToList
                            Dim myCommentCnt As Integer = myCommentList.Count
                            For Each myExComment In myCommentList
                                If myExComment.Paragraphs.Count > 0 Then
                                    myExBook.LogComment = myExBook.LogComment & vbTab & "コメント：" & vbTab & myExComment.Paragraphs(0).Text & vbCrLf
                                End If
                            Next
                            myExBook.CntComment = myExBook.CntComment + myCommentCnt
                        End If
                    Next
                Next

                If myExBookList.Any(Function(n) n.CntChart > 0 OrElse n.CntFormula > 0 OrElse n.CntComment > 0) Then
                    MessageBox.Show("計算式、チャート、コメントが含まれているファイルが存在しています。" & vbCrLf &
                                    "計算式が含まれるセル数合計：" & myExBookList.Sum(Function(n) n.CntFormula) & vbCrLf &
                                    "グラフ合計：" & myExBookList.Sum(Function(n) n.CntChart) & vbCrLf &
                                    "コメント合計：" & myExBookList.Sum(Function(n) n.CntComment) & vbCrLf &
                                    "詳細はログを参照してください。")
                    For Each myExBook As ExBook In myExBookList
                        Dim myLogPath As String = myOutputPath & "\" & DateTime.Now.ToString("yyMMddHHmmss") & Path.GetFileNameWithoutExtension(myExBook.Name)
                        If myExBook.CntChart > 0 Then
                            File.WriteAllText(myLogPath & "_ChartLog.txt", myExBook.LogChart, System.Text.Encoding.UTF8)
                        End If
                        If myExBook.CntFormula > 0 Then
                            File.WriteAllText(myLogPath & "_FormulaLog.txt", myExBook.LogFormula, System.Text.Encoding.UTF8)
                        End If
                        If myExBook.CntComment > 0 Then
                            File.WriteAllText(myLogPath & "_CommentLog.txt", myExBook.LogComment, System.Text.Encoding.UTF8)
                        End If
                    Next
                    'Dim myLogPath As String = myOutputPath & "\" & DateTime.Now.ToString("yyMMddHHmmss") & "_log.txt"
                    'File.WriteAllText(myLogPath, myLog, System.Text.Encoding.UTF8)
                End If

                Lbl_Progress.Text = "終了！"
                MessageBox.Show("終了！")
                myExcel.Quit()
                ReleaseObject(myExcel)
            Else
                MessageBox.Show("処理するファイルがありません！")
            End If
        Else
            MessageBox.Show("出力先フォルダーを設定してください！")
        End If
        Me.Enabled = True
    End Sub
    Private Function ShapeProc(myShape As Excel.Shape, myExShape As ExShape, ShapeNum As Integer)
        'myShape.Select()
        Try
            myExShape.ID = myShape.ID
        Catch ex As Exception
            myExShape.ID = 0
        End Try
        myExShape.Type = myShape.Type
        myExShape.HasText = False
        If myShape.Type = MsoShapeType.msoGroup Then
            Dim myExShapesGI As New List(Of ExShape)
            For Each myShapeGI As Excel.Shape In myShape.GroupItems
                Dim myExShapeGI As New ExShape
                myExShapeGI.No = ShapeNum
                ShapeNum += 1
                ShapeNum = ShapeProc(myShapeGI, myExShapeGI, ShapeNum)
                myExShapesGI.Add(myExShapeGI)
            Next
            myExShape.Shapes = myExShapesGI
        ElseIf myShape.Type = MsoShapeType.msoSmartArt Then
            Dim myExShapesSI As New List(Of ExShape)
            For Each myShapeSI As Excel.Shape In myShape.GroupItems
                Dim myExShapeSI As New ExShape
                myExShapeSI.No = ShapeNum
                ShapeNum += 1
                ShapeNum = ShapeProc(myShapeSI, myExShapeSI, ShapeNum)
                myExShapesSI.Add(myExShapeSI)
            Next
            myExShape.Shapes = myExShapesSI
        ElseIf myShape.Type = MsoShapeType.msoCanvas Then
            Dim myExShapesCI As New List(Of ExShape)
            For Each myShapeCI As Excel.Shape In myShape.CanvasItems
                Dim myExShapeCI As New ExShape
                myExShapeCI.No = ShapeNum
                ShapeNum += 1
                ShapeNum = ShapeProc(myShapeCI, myExShapeCI, ShapeNum)
                myExShapesCI.Add(myExShapeCI)
            Next
            myExShape.Shapes = myExShapesCI
        ElseIf myShape.Type = MsoShapeType.msoComment Then
            Dim myExparagraphList As New List(Of ExParagraph)
            Dim myExParagraph As New ExParagraph
            myExParagraph.Text = myShape.DrawingObject.Text
            myExparagraphList.Add(myExParagraph)
            myExShape.Paragraphs = myExparagraphList
        ElseIf myShape.Type = MsoShapeType.msoChart Then
            Dim myExparagraphList As New List(Of ExParagraph)
            Dim myExParagraph As New ExParagraph
            myExParagraph.Text = myShape.Chart.ChartTitle.Text
            myExparagraphList.Add(myExParagraph)
            myExShape.Paragraphs = myExparagraphList
        Else
            If myShape.TextFrame2.HasText = True Then
                'If myShape.TextFrame.HasText = MsoTriState.msoTrue Then
                myExShape.HasText = True
                myExShape.ReadError = False
                'myExShape.Text = myShape.TextFrame.TextRange.Text.Replace(ChrW(&HB), vbCrLf)
                Dim myBgColor As Long = Nud_BgColorB.Value * 16 ^ 4 + Nud_BgColorG.Value * 16 ^ 2 + Nud_BgColorR.Value
                Dim myFtColor As Long = Nud_FtColorB.Value * 16 ^ 4 + Nud_FtColorG.Value * 16 ^ 2 + Nud_FtColorR.Value

                Dim myExParagraphList As New List(Of ExParagraph)

                If myBgColorFlag = True Then
                    If myShape.Fill.ForeColor.RGB = myBgColor AndAlso myBgColorOut = True Then
                        Try
                            For Each myPara In myShape.TextFrame2.TextRange.Paragraphs
                                Dim myExParagraph As New ExParagraph
                                myExParagraph.Text = myPara.Text.Replace(ChrW(&HB), vbCrLf).Trim(vbCrLf)
                                myExParagraphList.Add(myExParagraph)
                            Next
                        Catch ex As Exception
                            myExShape.ReadError = True
                        End Try
                        myExShape.Paragraphs = myExParagraphList
                    ElseIf myShape.Fill.ForeColor.RGB <> myBgColor AndAlso myBgColorOut = False Then
                        Try
                            For Each myPara In myShape.TextFrame2.TextRange.Paragraphs
                                Dim myExParagraph As New ExParagraph
                                myExParagraph.Text = myPara.Text.Replace(ChrW(&HB), vbCrLf).Trim(vbCrLf)
                                myExParagraphList.Add(myExParagraph)
                            Next
                        Catch ex As Exception
                            myExShape.ReadError = True
                        End Try
                        myExShape.Paragraphs = myExParagraphList
                    End If
                ElseIf myFtColorFlag = True Then
                    If myShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = myFtColor AndAlso myFtColorOut = True Then
                        Try
                            For Each myPara In myShape.TextFrame2.TextRange.Paragraphs
                                Dim myExParagraph As New ExParagraph
                                myExParagraph.Text = myPara.Text.Replace(ChrW(&HB), vbCrLf).Trim(vbCrLf)
                                myExParagraphList.Add(myExParagraph)
                            Next
                        Catch ex As Exception
                            myExShape.ReadError = True
                        End Try
                        myExShape.Paragraphs = myExParagraphList
                    ElseIf myShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB <> myFtColor AndAlso myFtColorOut = False Then
                        Try
                            For Each myPara In myShape.TextFrame2.TextRange.Paragraphs
                                Dim myExParagraph As New ExParagraph
                                myExParagraph.Text = myPara.Text.Replace(ChrW(&HB), vbCrLf).Trim(vbCrLf)
                                myExParagraphList.Add(myExParagraph)
                            Next
                        Catch ex As Exception
                            myExShape.ReadError = True
                        End Try
                        myExShape.Paragraphs = myExParagraphList
                    End If
                Else
                    Try
                        For Each myPara In myShape.TextFrame2.TextRange.Paragraphs
                            Dim myExParagraph As New ExParagraph
                            myExParagraph.Text = myPara.Text.Replace(ChrW(&HB), vbCrLf).Trim(vbCrLf)
                            myExParagraphList.Add(myExParagraph)
                        Next
                    Catch ex As Exception
                        myExShape.ReadError = True
                    End Try
                    myExShape.Paragraphs = myExParagraphList
                End If

            End If
            'End If
        End If
        Return ShapeNum
    End Function

    Private Sub ExportXML(myExBook As ExBook, myXlsxFilePath As String, myOutputPath As String)
        Dim xDoc As New XDocument(New XElement("book", New XAttribute("name", myExBook.Name)))
        Dim bookElm As XElement = xDoc.Root
        Dim sheetsElm As New XElement("sheets")
        bookElm.Add(sheetsElm)

        For Each myExSheet In myExBook.Sheets
            Dim sheetElm As New XElement("sheet", New XAttribute("name", myExSheet.Name), New XAttribute("index", myExSheet.Index))
            Dim sheetNameElm As New XElement("sheetname")
            sheetNameElm.Add(New XText(myExSheet.Name))
            sheetElm.Add(sheetNameElm)
            Dim cellsElm As New XElement("cells")
            For Each myExCell In myExSheet.Cells
                If myExCell.IsFormula = False Then
                    Dim cellElm As New XElement("cell", New XAttribute("X", myExCell.X), New XAttribute("Y", myExCell.Y))
                    Dim myText = myExCell.Text
                    Dim startPos As Integer = 0
                    Dim endPos As Integer '= myText.Length
                    For Each match As Match In Regex.Matches(myText, "\\n")
                        endPos = match.Index
                        cellElm.Add(New XText(myText.Substring(startPos, endPos - startPos)))
                        Debug.WriteLine(myText.Substring(startPos, endPos - startPos))
                        Dim brElm As New XElement("br")
                        cellElm.Add(brElm)
                        startPos = endPos + 2
                    Next
                    cellElm.Add(New XText(myText.Substring(startPos, myText.Length - startPos)))
                    cellsElm.Add(cellElm)
                End If
            Next
            sheetElm.Add(cellsElm)
            '''''''
            Dim shapesElm As New XElement("shapes")
            ExportShapesXML(myExSheet.Shapes, shapesElm)
            sheetElm.Add(shapesElm)
            '''''''
            sheetsElm.Add(sheetElm)
        Next
        Debug.WriteLine(bookElm.Name)

        'Dim myDirectory As String = Path.GetDirectoryName(myXlsxFilePath) & "\"
        Dim myXmlFileName As String = Path.GetFileNameWithoutExtension(myXlsxFilePath) & ".xml"
        xDoc.Save(myOutputPath & "\" & myXmlFileName)
    End Sub
    Private Sub ExportShapesXML(myExShapes As List(Of ExShape), shapesElm As XElement)
        If myExShapes IsNot Nothing Then
            For Each myExShape In myExShapes
                Dim shapeElm As New XElement("shape", New XAttribute("no", myExShape.No), New XAttribute("id", myExShape.ID), New XAttribute("type", myExShape.Type), New XAttribute("hastext", myExShape.HasText))
                If myExShape.HasText = True Then
                    Dim paragraphsElm As New XElement("paragraphs")
                    paragraphsElm.SetAttributeValue("error", False)
                    If myExShape.ReadError = False Then
                        If myExShape.Paragraphs IsNot Nothing Then
                            For Each myExParagraph As ExParagraph In myExShape.Paragraphs
                                Dim paragraphElm As New XElement("paragraph")
                                paragraphElm.Value = myExParagraph.Text.Replace(ChrW(&HB), vbCrLf).Trim
                                paragraphsElm.Add(paragraphElm)
                            Next
                        End If
                    Else
                        paragraphsElm.SetAttributeValue("error", True)
                    End If
                    shapeElm.Add(paragraphsElm)
                End If

                If myExShape.Shapes IsNot Nothing Then
                    Dim shapesCElm As New XElement("shapes")
                    ExportShapesXML(myExShape.Shapes, shapesCElm)
                    shapeElm.Add(shapesCElm)
                End If
                shapesElm.Add(shapeElm)
            Next
        End If
    End Sub
    Private Sub Btn_Import_Click(sender As Object, e As EventArgs) Handles Btn_Import.Click
        Me.Enabled = False
        Dim myDirectory As String = Tbx_OutputFolder2.Text
        If Directory.Exists(myDirectory) Then
            myExcel = New Excel.Application With {.Visible = False}
            For i = 0 To myXmlFileList.Count - 1
                Dim myXmlFilePath As String = myXmlFileList(i).Trim()
                Lbx_Xml.SelectedIndex = i

                If File.Exists(myXmlFilePath) AndAlso Path.GetExtension(myXmlFilePath).ToLower = ".xml" Then
                    Dim myXMLBaseName = Regex.Replace(Path.GetFileNameWithoutExtension(myXmlFilePath), "_trn$", String.Empty, RegexOptions.IgnoreCase)
                    Dim myXlsxFile As String = Path.GetDirectoryName(myXmlFilePath) & "\" & myXMLBaseName & ".xlsx"
                    If File.Exists(myXlsxFile) Then
                        Dim xDoc As XDocument = XDocument.Load(myXmlFilePath)
                        Dim bookElm As XElement = xDoc.Root

                        Dim myBook As Excel.Workbook = myExcel.Workbooks.Open(myXlsxFile)
                        Dim sheetsElm As XElement = bookElm.Element("sheets")
                        Dim sheetElms = sheetsElm.Elements("sheet")
                        For Each sheetElm In sheetElms
                            Dim sheetIndex As Integer = sheetElm.Attribute("index").Value
                            Dim mySheet As Worksheet = myBook.Worksheets(sheetIndex)
                            Dim sheetnameElm As XElement = sheetElm.Element("sheetname")
                            mySheet.Name = sheetnameElm.Value
                            Dim cellsElm As XElement = sheetElm.Element("cells")
                            Dim cellElms = cellsElm.Elements("cell")
                            Dim ttl = cellElms.Count
                            Dim cnt = 0
                            For Each cellElm In cellElms
                                Dim x As Integer = cellElm.Attribute("X").Value
                                Dim y As Integer = cellElm.Attribute("Y").Value
                                Dim myCell As Excel.Range = myBook.Worksheets(sheetIndex).cells(x, y)
                                If TypeOf myCell.Value Is String AndAlso myCell.Value <> "" AndAlso Not myCell.HasFormula Then
                                    Dim myValue As String = ""
                                    For Each myNode In cellElm.Nodes
                                        If myNode.NodeType = XmlNodeType.Text Then
                                            Dim myText As XText = myNode
                                            myValue &= myText.Value
                                        ElseIf myNode.NodeType = XmlNodeType.Element Then
                                            Dim myElm As XElement = myNode
                                            If myElm.Name = "br" Then
                                                myValue &= vbLf
                                            End If
                                        End If
                                    Next
                                    If Chk_DollarChar.Checked = True Then
                                        myCell.Value = Regex.Replace(myValue, "\w", "$")
                                    Else
                                        myCell.Value = myValue
                                    End If
                                    Debug.WriteLine(myValue)
                                End If
                                cnt += 1
                                Lbl_Progress.Text = sheetElm.Attribute("name").Value & ", " & cnt & " / " & ttl
                            Next
                            Dim shapesElm As XElement = sheetElm.Element("shapes")
                            If shapesElm IsNot Nothing Then
                                Dim ShapeNum As Integer = 1
                                For Each myShape As Excel.Shape In myBook.Worksheets(sheetIndex).Shapes
                                    Dim shapeElm As XElement = shapesElm.Elements.Where(Function(n) n.Attribute("no") = ShapeNum.ToString).FirstOrDefault
                                    ShapeNum += 1
                                    ShapeNum = ShapeProcImport(myShape, shapeElm, ShapeNum)
                                Next
                            End If
                        Next
                        'Dim myDir = Path.GetDirectoryName(myXlsxFilePath)
                        'Dim myFileName = DateTime.Now.ToString("yyMMddHHmmss_") & Path.GetFileName(myXlsxFilePath)
                        myBook.SaveAs(myDirectory & "\" & Path.GetFileName(myXlsxFile))
                        myBook.Close()
                        ReleaseObject(myBook)
                    Else
                        MessageBox.Show(”対応するExcelファイルが存在しません。" & vbCrLf & myXlsxFile)
                    End If
                Else
                    MessageBox.Show("正しいファイルを選択してください！")
                End If
            Next
            myExcel.Quit()
            ReleaseObject(myExcel)
            Lbl_Progress.Text = "終了！"
            MessageBox.Show("終了！")
        Else
            MessageBox.Show("出力先フォルダーを設定してください！")
        End If
        Me.Enabled = True
    End Sub
    Private Function ShapeProcImport(myShape As Excel.Shape, shapeElm As XElement, shapeNum As Integer)
        If shapeElm IsNot Nothing Then
            If myShape.Type = MsoShapeType.msoGroup Then
                Dim shapesGIElm As XElement = shapeElm.Element("shapes")
                For Each myShapeGI As Excel.Shape In myShape.GroupItems
                    Dim shapeGIElm As XElement = shapesGIElm.Elements.Where(Function(n) n.Attribute("no") = shapeNum.ToString).FirstOrDefault
                    shapeNum += 1
                    shapeNum = ShapeProcImport(myShapeGI, shapeGIElm, shapeNum)
                Next
            ElseIf myShape.Type = MsoShapeType.msoSmartArt Then
                Dim shapesSIElm As XElement = shapeElm.Element("shapes")
                For Each myShapeSI As Excel.Shape In myShape.GroupItems
                    Dim shapeSIElm As XElement = shapesSIElm.Elements.Where(Function(n) n.Attribute("no") = shapeNum.ToString).FirstOrDefault
                    shapeNum += 1
                    shapeNum = ShapeProcImport(myShapeSI, shapeSIElm, shapeNum)
                Next
            ElseIf myShape.Type = MsoShapeType.msoCanvas Then
                Dim shapesCIElm As XElement = shapeElm.Element("shapes")
                For Each myShapeCI As Excel.Shape In myShape.CanvasItems
                    Dim shapeCIElm As XElement = shapesCIElm.Elements.Where(Function(n) n.Attribute("no") = shapeNum.ToString).FirstOrDefault
                    shapeNum += 1
                    shapeNum = ShapeProcImport(myShapeCI, shapeCIElm, shapeNum)
                Next
            ElseIf myShape.Type = MsoShapeType.msoComment Then
                '    Dim shapesCIElm As XElement = shapeElm.Element("shapes")
                '    Dim shapeCIElm As XElement = shapesCIElm.Elements.Where(Function(n) n.Attribute("no") = shapeNum.ToString).FirstOrDefault
                '    shapeNum += 1
            ElseIf myShape.Type = MsoShapeType.msoChart Then
                '    Dim shapesCIElm As XElement = shapeElm.Element("shapes")
                '    Dim shapeCIElm As XElement = shapesCIElm.Elements.Where(Function(n) n.Attribute("no") = shapeNum.ToString).FirstOrDefault
                '    shapeNum += 1
            Else
                If myShape.TextFrame2.HasText = True Then
                    'myShape.TextFrame.TextRange.Text = New String("x"c, shapeElm.Value.Length)
                    'myShape.TextFrame.TextRange.Text = Regex.Replace(shapeElm.Value, "[A-Za-z0-9]", "x")
                    Dim paragraphsElm As XElement = shapeElm.Element("paragraphs")

                    If paragraphsElm.Attribute("error") = "false" Then
                        Dim paragraphElmList As List(Of XElement) = paragraphsElm.Elements("paragraph").ToList
                        For i As Integer = 1 To myShape.TextFrame2.TextRange.Paragraphs.Count
                            Dim myParagraph = myShape.TextFrame2.TextRange.Paragraphs(i)
                            Dim paragraphElm As XElement = paragraphElmList(i - 1)
                            If Chk_DollarChar.Checked = True Then
                                myParagraph.Text = Regex.Replace(paragraphElm.Value, "\w", "$")
                            Else
                                myParagraph.Text = paragraphElm.Value
                            End If
                        Next
                    End If
                End If
            End If
            Return shapeNum
        End If
    End Function

    'Private Sub Chk_HighlightExport_CheckedChanged(sender As Object, e As EventArgs)
    '    If Rbn_BgColor.Checked = True Then
    '        Lbl_Color.Enabled = True
    '        'Cmb_Color.Enabled = True
    '        Pcb_BgColor.Enabled = True
    '    Else
    '        Lbl_Color.Enabled = False
    '        'Cmb_Color.Enabled = False
    '        Pcb_BgColor.Enabled = False
    '    End If
    'End Sub

    'Private Sub TextBox1_Click(sender As Object, e As EventArgs)
    '    'Dim cd As New ColorDialog()

    '    ''はじめに選択されている色を設定
    '    'cd.Color = TextBox1.BackColor
    '    ''色の作成部分を表示可能にする
    '    ''デフォルトがTrueのため必要はない
    '    'cd.AllowFullOpen = True
    '    ''純色だけに制限しない
    '    ''デフォルトがFalseのため必要はない
    '    'cd.SolidColorOnly = False
    '    ''[作成した色]に指定した色（RGB値）を表示する
    '    'cd.CustomColors = New Integer() {&H33, &H66, &H99,
    '    '    &HCC, &H3300, &H3333, &H3366, &H3399, &H33CC,
    '    '    &H6600, &H6633, &H6666, &H6699, &H66CC,
    '    '    &H9900, &H9933}

    '    ''ダイアログを表示する
    '    'If cd.ShowDialog() = DialogResult.OK Then
    '    '    '選択された色の取得
    '    '    TextBox1.BackColor = cd.Color
    '    'End If
    'End Sub

    Private Sub Pcb_BgColor_Click(sender As Object, e As EventArgs) Handles Pcb_BgColor.Click
        Dim cd As New ColorDialog() With {
        .Color = Pcb_BgColor.BackColor,
        .AllowFullOpen = True,
        .SolidColorOnly = False,
        .CustomColors = myBgColorList.ToArray
        }
        'ダイアログを表示する
        If cd.ShowDialog() = DialogResult.OK Then
            Pcb_BgColor.BackColor = cd.Color
            Nud_BgColorR.Value = cd.Color.R
            Nud_BgColorG.Value = cd.Color.G
            Nud_BgColorB.Value = cd.Color.B
        End If
        myBgColorList = cd.CustomColors.ToList
        Debug.WriteLine("")
    End Sub

    Private Sub Pcb_FtColor_Click(sender As Object, e As EventArgs) Handles Pcb_FtColor.Click
        Dim cd As New ColorDialog() With {
        .Color = Pcb_FtColor.BackColor,
        .AllowFullOpen = True,
        .SolidColorOnly = False,
        .CustomColors = myFtColorList.ToArray
        }
        'ダイアログを表示する
        If cd.ShowDialog() = DialogResult.OK Then
            Pcb_FtColor.BackColor = cd.Color
            Nud_FtColorR.Value = cd.Color.R
            Nud_FtColorG.Value = cd.Color.G
            Nud_FtColorB.Value = cd.Color.B
        End If
        myFtColorList = cd.CustomColors.ToList
        Debug.WriteLine("")
    End Sub

    Private Sub Rbn_BgColor_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_BgColor.CheckedChanged
        If Rbn_BgColor.Checked = True Then
            Pnl_BgColor.Enabled = True
            Pnl_FntColor.Enabled = False
        End If
    End Sub

    Private Sub Rbn_FntColor_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_FntColor.CheckedChanged
        If Rbn_FntColor.Checked = True Then
            Pnl_BgColor.Enabled = False
            Pnl_FntColor.Enabled = True
        End If
    End Sub

    Private Sub Rbn_All_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_All.CheckedChanged
        If Rbn_All.Checked = True Then
            Pnl_BgColor.Enabled = False
            Pnl_FntColor.Enabled = False
        End If
    End Sub

    Private Sub Nud_BgColorR_ValueChanged(sender As Object, e As EventArgs) Handles Nud_BgColorR.ValueChanged
        Pcb_BgColor.BackColor = Color.FromArgb(Nud_BgColorR.Value, Nud_BgColorG.Value, Nud_BgColorB.Value)
    End Sub

    Private Sub Nud_BgColorG_ValueChanged(sender As Object, e As EventArgs) Handles Nud_BgColorG.ValueChanged
        Pcb_BgColor.BackColor = Color.FromArgb(Nud_BgColorR.Value, Nud_BgColorG.Value, Nud_BgColorB.Value)
    End Sub

    Private Sub Nud_BgColorB_ValueChanged(sender As Object, e As EventArgs) Handles Nud_BgColorB.ValueChanged
        Pcb_BgColor.BackColor = Color.FromArgb(Nud_BgColorR.Value, Nud_BgColorG.Value, Nud_BgColorB.Value)
    End Sub

    Private Sub Nud_FtColorR_ValueChanged(sender As Object, e As EventArgs) Handles Nud_FtColorR.ValueChanged
        Pcb_FtColor.BackColor = Color.FromArgb(Nud_FtColorR.Value, Nud_FtColorG.Value, Nud_FtColorB.Value)
    End Sub

    Private Sub Nud_FtColorG_ValueChanged(sender As Object, e As EventArgs) Handles Nud_FtColorG.ValueChanged
        Pcb_FtColor.BackColor = Color.FromArgb(Nud_FtColorR.Value, Nud_FtColorG.Value, Nud_FtColorB.Value)
    End Sub

    Private Sub Nud_FtColorB_ValueChanged(sender As Object, e As EventArgs) Handles Nud_FtColorB.ValueChanged
        Pcb_FtColor.BackColor = Color.FromArgb(Nud_FtColorR.Value, Nud_FtColorG.Value, Nud_FtColorB.Value)
    End Sub

    Private Sub Pcb_BgCSet1_Click(sender As Object, e As EventArgs) Handles Pcb_BgCSet1.Click
        PcbBgColorSet(sender)
    End Sub
    Private Sub Pcb_BgCSet2_Click(sender As Object, e As EventArgs) Handles Pcb_BgCSet2.Click
        PcbBgColorSet(sender)
    End Sub
    Private Sub Pcb_BgCSet3_Click(sender As Object, e As EventArgs) Handles Pcb_BgCSet3.Click
        PcbBgColorSet(sender)
    End Sub
    Private Sub Pcb_BgCSet4_Click(sender As Object, e As EventArgs) Handles Pcb_BgCSet4.Click
        PcbBgColorSet(sender)
    End Sub
    Private Sub Pcb_BgCSet5_Click(sender As Object, e As EventArgs) Handles Pcb_BgCSet5.Click
        PcbBgColorSet(sender)
    End Sub
    Private Sub Pcb_BgCSet6_Click(sender As Object, e As EventArgs) Handles Pcb_BgCSet6.Click
        PcbBgColorSet(sender)
    End Sub
    Private Sub Pcb_BgCSet7_Click(sender As Object, e As EventArgs) Handles Pcb_BgCSet7.Click
        PcbBgColorSet(sender)
    End Sub
    Private Sub Pcb_BgCSet8_Click(sender As Object, e As EventArgs) Handles Pcb_BgCSet8.Click
        PcbBgColorSet(sender)
    End Sub
    Private Sub Pcb_BgCSet9_Click(sender As Object, e As EventArgs) Handles Pcb_BgCSet9.Click
        PcbBgColorSet(sender)
    End Sub
    Private Sub Pcb_BgCSet10_Click(sender As Object, e As EventArgs) Handles Pcb_BgCSet10.Click
        PcbBgColorSet(sender)
    End Sub
    Private Sub PcbBgColorSet(myPictureBox As PictureBox)
        Nud_BgColorR.Value = myPictureBox.BackColor.R
        Nud_BgColorG.Value = myPictureBox.BackColor.G
        Nud_BgColorB.Value = myPictureBox.BackColor.B
    End Sub

    Private Sub Pcb_FtCSet1_Click(sender As Object, e As EventArgs) Handles Pcb_FtCSet1.Click
        PcbFtColorSet(sender)
    End Sub
    Private Sub Pcb_FtCSet2_Click(sender As Object, e As EventArgs) Handles Pcb_FtCSet2.Click
        PcbFtColorSet(sender)
    End Sub
    Private Sub Pcb_FtCSet3_Click(sender As Object, e As EventArgs) Handles Pcb_FtCSet3.Click
        PcbFtColorSet(sender)
    End Sub
    Private Sub Pcb_FtCSet4_Click(sender As Object, e As EventArgs) Handles Pcb_FtCSet4.Click
        PcbFtColorSet(sender)
    End Sub
    Private Sub Pcb_FtCSet5_Click(sender As Object, e As EventArgs) Handles Pcb_FtCSet5.Click
        PcbFtColorSet(sender)
    End Sub
    Private Sub Pcb_FtCSet6_Click(sender As Object, e As EventArgs) Handles Pcb_FtCSet6.Click
        PcbFtColorSet(sender)
    End Sub
    Private Sub Pcb_FtCSet7_Click(sender As Object, e As EventArgs) Handles Pcb_FtCSet7.Click
        PcbFtColorSet(sender)
    End Sub
    Private Sub Pcb_FtCSet8_Click(sender As Object, e As EventArgs) Handles Pcb_FtCSet8.Click
        PcbFtColorSet(sender)
    End Sub
    Private Sub Pcb_FtCSet9_Click(sender As Object, e As EventArgs) Handles Pcb_FtCSet9.Click
        PcbFtColorSet(sender)
    End Sub
    Private Sub Pcb_FtCSet10_Click(sender As Object, e As EventArgs) Handles Pcb_FtCSet10.Click
        PcbFtColorSet(sender)
    End Sub
    Private Sub PcbFtColorSet(myPictureBox As PictureBox)
        Nud_FtColorR.Value = myPictureBox.BackColor.R
        Nud_FtColorG.Value = myPictureBox.BackColor.G
        Nud_FtColorB.Value = myPictureBox.BackColor.B
    End Sub

    Private Sub Lbx_Xlsx_KeyDown(sender As Object, e As KeyEventArgs) Handles Lbx_Xlsx.KeyDown
        If e.KeyData = Keys.Delete Then
            If Not myXlsxFileList Is Nothing AndAlso myXlsxFileList.Count > 0 Then
                myXlsxFileList.RemoveAt(Lbx_Xlsx.SelectedIndex)
            End If
        End If
    End Sub

    Private Sub Lbx_Xml_KeyDown(sender As Object, e As KeyEventArgs) Handles Lbx_Xml.KeyDown
        If e.KeyData = Keys.Delete Then
            If Not myXmlFileList Is Nothing AndAlso myXmlFileList.Count > 0 Then
                myXmlFileList.RemoveAt(Lbx_Xml.SelectedIndex)
            End If
        End If
    End Sub
End Class

Public Class HLColor
    Public Property Name As String
    Public Property ColorHex As String
End Class

Public Class ExBook
    Public Property Name As String
    Public Property Sheets As List(Of ExSheet)
    Public Property LogChart As String
    Public Property LogFormula As String
    Public Property LogComment As String
    Public Property CntChart As Integer
    Public Property CntFormula As Integer
    Public Property CntComment As Integer
    Public Sub New()
        LogChart = ""
        LogFormula = ""
        LogComment = ""
        CntChart = 0
        CntFormula = 0
        CntComment = 0
    End Sub
End Class
Public Class ExSheet
    Public Property Name As String
    Public Property Index As Integer
    Public Property Cells As List(Of ExCell)
    Public Property Shapes As List(Of ExShape)
End Class
Public Class ExCell
    Public Property X As Integer
    Public Property Y As Integer
    Public Property Text As String
    Public Property IsFormula As Boolean
End Class

Public Class ExShape
    Public Property ID As Integer
    Public Property Type As Integer
    Public Property HasText As Boolean
    Public Property ReadError As Boolean
    Public Property Paragraphs As List(Of ExParagraph)
    Public Property Shapes As List(Of ExShape)
    Public Property No As Integer
End Class
Public Class ExParagraph
    Public Property Text As String
End Class