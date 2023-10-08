'#Region "#usings"
Imports System.Drawing
Imports System.Windows.Forms
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraSpreadsheet

'#End Region  ' #usings
Namespace CustomDrawExample

    Public Partial Class Form1
        Inherits Form

        Public Sub New()
            InitializeComponent()
            spreadsheetControl1.LoadDocument("CustomDrawSample.xlsx")
            spreadsheetControl1.Document.DocumentSettings.R1C1ReferenceStyle = True
            spreadsheetControl1.Options.Behavior.Column.Delete = DocumentCapability.Disabled
            spreadsheetControl1.Options.Behavior.Column.Insert = DocumentCapability.Disabled
            AddHandler spreadsheetControl1.CustomDrawColumnHeader, AddressOf spreadsheetControl1_CustomDrawColumnHeader
            AddHandler spreadsheetControl1.CustomDrawColumnHeaderBackground, AddressOf spreadsheetControl1_CustomDrawColumnHeaderBackground
            AddHandler spreadsheetControl1.CustomDrawRowHeader, AddressOf spreadsheetControl1_CustomDrawRowHeader
            AddHandler spreadsheetControl1.CustomDrawRowHeaderBackground, AddressOf spreadsheetControl1_CustomDrawRowHeaderBackground
            AddHandler spreadsheetControl1.CustomDrawCell, AddressOf spreadsheetControl1_CustomDrawCell
            AddHandler spreadsheetControl1.CustomDrawCellBackground, AddressOf spreadsheetControl1_CustomDrawCellBackground
        End Sub

'#Region "#CustomDrawColumnHeader"
        Private Sub spreadsheetControl1_CustomDrawColumnHeader(ByVal sender As Object, ByVal e As CustomDrawColumnHeaderEventArgs)
            e.Handled = True
            Dim foreColor As Color = Color.Blue
            Dim textBounds As Rectangle = e.Bounds
            Dim headingFont As Font = New Font(e.Font, FontStyle.Italic)
            Dim settingsSheet As Worksheet = spreadsheetControl1.Document.Worksheets("SheetSettings")
            Dim text As String = settingsSheet.Cells(0, e.ColumnIndex).DisplayText
            If Not Equals(text, String.Empty) Then
                Dim formatHeaderText As StringFormat = New StringFormat()
                formatHeaderText.LineAlignment = StringAlignment.Center
                formatHeaderText.Alignment = StringAlignment.Center
                formatHeaderText.Trimming = StringTrimming.EllipsisCharacter
                e.Graphics.DrawString(text, headingFont, e.Cache.GetSolidBrush(foreColor), textBounds, formatHeaderText)
            End If
        End Sub

'#End Region  ' #CustomDrawColumnHeader
'#Region "#CustomDrawColumnHeaderBackground"
        Private Sub spreadsheetControl1_CustomDrawColumnHeaderBackground(ByVal sender As Object, ByVal e As CustomDrawColumnHeaderBackgroundEventArgs)
            e.Handled = True
            Dim is_selected As Boolean = e.IsHovered OrElse e.ColumnIndex = spreadsheetControl1.ActiveCell.ColumnIndex
            Dim backColor As Color = If(is_selected, Color.Yellow, Color.White)
            e.Cache.FillRectangle(e.Cache.GetSolidBrush(backColor), e.Bounds)
        End Sub

'#End Region  ' #CustomDrawColumnHeaderBackground
'#Region "#CustomDrawRowHeader"
        Private Sub spreadsheetControl1_CustomDrawRowHeader(ByVal sender As Object, ByVal e As CustomDrawRowHeaderEventArgs)
            e.Handled = True
            If(e.RowIndex + 1) Mod 5 = 0 Then
                Dim font As System.Drawing.Font = New Font(e.Font.FontFamily, e.Font.Size, FontStyle.Bold)
                Dim textBounds As Rectangle = e.Bounds
                Dim text As String =(e.RowIndex + 1).ToString()
                Dim formatHeaderText As StringFormat = New StringFormat()
                formatHeaderText.LineAlignment = StringAlignment.Center
                formatHeaderText.Alignment = StringAlignment.Center
                e.Graphics.DrawString(text, font, e.Cache.GetSolidBrush(Color.Red), textBounds, formatHeaderText)
            Else
            'e.DrawDefault();
            End If
        End Sub

'#End Region  ' #CustomDrawRowHeader
'#Region "#CustomDrawRowHeaderBackground"
        Private Sub spreadsheetControl1_CustomDrawRowHeaderBackground(ByVal sender As Object, ByVal e As CustomDrawRowHeaderBackgroundEventArgs)
            e.Handled = True
            Dim is_selected As Boolean = e.IsHovered OrElse e.RowIndex = spreadsheetControl1.ActiveCell.RowIndex
            Dim backColor As Color = If(is_selected, Color.Yellow, Color.White)
            e.Cache.FillRectangle(e.Cache.GetSolidBrush(backColor), e.Bounds)
        End Sub

'#End Region  ' #CustomDrawRowHeaderBackground
'#Region "#CustomDrawCell"
        Private Sub spreadsheetControl1_CustomDrawCell(ByVal sender As Object, ByVal e As CustomDrawCellEventArgs)
            If e.Cell.RowIndex = 0 OrElse e.Cell.RowIndex >= 3 Then
                Dim headingFont As Font = New Font("Times New Roman", e.Font.Size)
                Dim cellRef As String = e.Cell.GetReferenceR1C1(ReferenceElement.RowAbsolute Or ReferenceElement.ColumnAbsolute, Nothing)
                Dim formula As String = String.Format("=RANK.AVG({0},R{1}C{2}:R{3}C{4})", cellRef, e.Cell.RowIndex + 1, 2, e.Cell.RowIndex + 1, 10)
                Dim rank As Integer = CInt(spreadsheetControl1.Document.Evaluate(formula).NumericValue)
                ' The DevExpress.Docs.Text.NumberInWords class requires a reference to the DevExpress.Docs assembly. 
                ' To redistribute the DevExpress.Docs assembly the DevExpress Universal subscription or the Document Server license is required.
                Dim rankText As String = DevExpress.Docs.Text.NumberInWords.Cardinal.ConvertToText(rank, DevExpress.Docs.Text.NumberCulture.Roman)
                If rank > 0 AndAlso rank < 4 Then e.Graphics.DrawString(rankText, headingFont, e.Cache.GetSolidBrush(Color.Red), e.Bounds.Left, e.Bounds.Top)
            End If
        End Sub

'#End Region  ' #CustomDrawCell
'#Region "#CustomDrawCellBackground"
        Private Sub spreadsheetControl1_CustomDrawCellBackground(ByVal sender As Object, ByVal e As CustomDrawCellBackgroundEventArgs)
            If e.Cell.HasFormula Then
                e.Handled = True
                Dim hBrush As System.Drawing.Drawing2D.HatchBrush = New System.Drawing.Drawing2D.HatchBrush(System.Drawing.Drawing2D.HatchStyle.BackwardDiagonal, Color.LightGray, Color.White)
                e.Graphics.FillRectangle(hBrush, e.FillBounds)
            End If
        End Sub
'#End Region  ' #CustomDrawCellBackground
    End Class
End Namespace
