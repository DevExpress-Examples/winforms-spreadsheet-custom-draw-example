#region #usings
using DevExpress.Spreadsheet;
using DevExpress.XtraSpreadsheet;
using System;
using System.Drawing;
using System.Windows.Forms;
#endregion #usings

namespace CustomDrawExample
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            spreadsheetControl1.LoadDocument("CustomDrawSample.xlsx");

            spreadsheetControl1.Document.DocumentSettings.R1C1ReferenceStyle = true;
            spreadsheetControl1.Options.Behavior.Column.Delete = DocumentCapability.Disabled;
            spreadsheetControl1.Options.Behavior.Column.Insert = DocumentCapability.Disabled;

            spreadsheetControl1.CustomDrawColumnHeader += spreadsheetControl1_CustomDrawColumnHeader;
            spreadsheetControl1.CustomDrawColumnHeaderBackground += spreadsheetControl1_CustomDrawColumnHeaderBackground;
            spreadsheetControl1.CustomDrawRowHeader += spreadsheetControl1_CustomDrawRowHeader;
            spreadsheetControl1.CustomDrawRowHeaderBackground += spreadsheetControl1_CustomDrawRowHeaderBackground;
            spreadsheetControl1.CustomDrawCell += spreadsheetControl1_CustomDrawCell;
            spreadsheetControl1.CustomDrawCellBackground += spreadsheetControl1_CustomDrawCellBackground;
        }
        #region #CustomDrawColumnHeader
        void spreadsheetControl1_CustomDrawColumnHeader(object sender, CustomDrawColumnHeaderEventArgs e)
        {
            e.Handled = true;
            Color foreColor = Color.Blue;
            Rectangle textBounds = e.Bounds;
            e.Appearance.FontStyleDelta = FontStyle.Italic;
            Worksheet settingsSheet = spreadsheetControl1.Document.Worksheets["SheetSettings"];
            string text = settingsSheet.Cells[0, e.ColumnIndex].DisplayText;
            if (text != String.Empty)
            {
                StringFormat formatHeaderText = new StringFormat();
                formatHeaderText.LineAlignment = StringAlignment.Center;
                formatHeaderText.Alignment = StringAlignment.Center;
                formatHeaderText.Trimming = StringTrimming.EllipsisCharacter;
                e.Graphics.DrawString(text, e.Font, e.Cache.GetSolidBrush(foreColor), textBounds, formatHeaderText);
            }
        }
        #endregion #CustomDrawColumnHeader

        #region #CustomDrawColumnHeaderBackground
        void spreadsheetControl1_CustomDrawColumnHeaderBackground(object sender, CustomDrawColumnHeaderBackgroundEventArgs e)
        {
            e.Handled = true;
            bool is_selected = e.IsHovered || (e.ColumnIndex == spreadsheetControl1.ActiveCell.ColumnIndex);
            Color backColor = is_selected ? Color.Yellow : Color.White;
            e.Cache.FillRectangle(e.Cache.GetSolidBrush(backColor), e.Bounds);
        }
        #endregion #CustomDrawColumnHeaderBackground

        #region #CustomDrawRowHeader
        void spreadsheetControl1_CustomDrawRowHeader(object sender, CustomDrawRowHeaderEventArgs e)
        {
            e.Handled = true;

            if ((e.RowIndex + 1) % 5 == 0)
            {
                e.Appearance.FontStyleDelta = FontStyle.Bold;
                Rectangle textBounds = e.Bounds;
                string text = (e.RowIndex + 1).ToString();
                StringFormat formatHeaderText = new StringFormat();
                formatHeaderText.LineAlignment = StringAlignment.Center;
                formatHeaderText.Alignment = StringAlignment.Center;
                e.Graphics.DrawString(text, e.Font, e.Cache.GetSolidBrush(Color.Red), textBounds, formatHeaderText);
            }
            else
            {
                //e.DrawDefault();
            }
        }
        #endregion #CustomDrawRowHeader

        #region #CustomDrawRowHeaderBackground
        void spreadsheetControl1_CustomDrawRowHeaderBackground(object sender, CustomDrawRowHeaderBackgroundEventArgs e)
        {
            e.Handled = true;
            bool is_selected = e.IsHovered || (e.RowIndex == spreadsheetControl1.ActiveCell.RowIndex);
            Color backColor = is_selected ? Color.Yellow : Color.White;
            e.Cache.FillRectangle(e.Cache.GetSolidBrush(backColor), e.Bounds);
        }
        #endregion #CustomDrawRowHeaderBackground

        #region #CustomDrawCell
        void spreadsheetControl1_CustomDrawCell(object sender, CustomDrawCellEventArgs e)
        {
            if (e.Cell.RowIndex == 0 || e.Cell.RowIndex >= 3)
            {
                using (Font headingFont = new Font("Times New Roman", e.Font.Size))
                {
                    string cellRef = e.Cell.GetReferenceR1C1(ReferenceElement.RowAbsolute | ReferenceElement.ColumnAbsolute, null);
                    string formula = String.Format("=RANK.AVG({0},R{1}C{2}:R{3}C{4})", cellRef, e.Cell.RowIndex + 1, 2, e.Cell.RowIndex + 1, 10);
                    int rank = (int)spreadsheetControl1.Document.Evaluate(formula).NumericValue;
                    // The DevExpress.Docs.Text.NumberInWords class requires a reference to the DevExpress.Docs assembly. 
                    // To redistribute the DevExpress.Docs assembly the DevExpress Universal subscription or the Document Server license is required.
                    string rankText = DevExpress.Docs.Text.NumberInWords.Cardinal.ConvertToText(rank, DevExpress.Docs.Text.NumberCulture.Roman);
                    if (rank > 0 && rank < 4) 
                    { 
                        e.Graphics.DrawString(rankText, headingFont, e.Cache.GetSolidBrush(Color.Red), e.Bounds.Left, e.Bounds.Top); 
                    }
                }
            }
        }
        #endregion #CustomDrawCell

        #region #CustomDrawCellBackground
        void spreadsheetControl1_CustomDrawCellBackground(object sender, CustomDrawCellBackgroundEventArgs e)
        {
            if (e.Cell.HasFormula)
            {
                e.Handled = true;
                System.Drawing.Drawing2D.HatchBrush hBrush = new System.Drawing.Drawing2D.HatchBrush(
                    System.Drawing.Drawing2D.HatchStyle.BackwardDiagonal,
                    Color.LightGray,
                    Color.White);
                e.Graphics.FillRectangle(hBrush, e.FillBounds);
            }
        }
        #endregion #CustomDrawCellBackground
    }
}
