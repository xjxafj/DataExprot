using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Charts;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;
using System.Windows.Media;

namespace DataExport
{
    public class ExcelHelper
    {
        const int PT = 20;
        const int LENGTH = 200;
        const string TIMESNEWROMAN = "Times New Roman";
        const string TITLE = "Spread Sheet Chart Demo";

        public static void Export(string filePath, double[] data) 
        {

            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile excel = new ExcelFile();
            //Excel默认字体
            excel.DefaultFontName = TIMESNEWROMAN;
            //Excel文档属性设置
            excel.DocumentProperties.BuiltIn.Add(new KeyValuePair<BuiltInDocumentProperties, string>(BuiltInDocumentProperties.Title, TITLE));
            excel.DocumentProperties.BuiltIn.Add(new KeyValuePair<BuiltInDocumentProperties, string>(BuiltInDocumentProperties.Author, "CNXY"));
            excel.DocumentProperties.BuiltIn.Add(new KeyValuePair<BuiltInDocumentProperties, string>(BuiltInDocumentProperties.Company, "CNXY"));
            excel.DocumentProperties.BuiltIn.Add(new KeyValuePair<BuiltInDocumentProperties, string>(BuiltInDocumentProperties.Comments, "By CNXY.Website: http://www.cnc6.cn"));
            //新建一个Sheet表格
            ExcelWorksheet sheet = excel.Worksheets.Add(TITLE);
            //设置表格保护
            sheet.ProtectionSettings.SetPassword("cnxy");
            sheet.Protected = true;
            //设置网格线不可见
            sheet.ViewOptions.ShowGridLines = false;
            //定义一个B2-G3的单元格范围
            CellRange range = sheet.Cells.GetSubrange("B2", "J3");
            range.Value = "Chart";
            range.Merged = true;
            //定义一个单元格样式
            CellStyle style = new CellStyle();
            //设置边框
            style.Borders.SetBorders(MultipleBorders.Outside, SpreadsheetColor.FromName(ColorName.Red), LineStyle.Thin);
            //设置水平对齐模式
            style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            //设置垂直对齐模式
            style.VerticalAlignment = VerticalAlignmentStyle.Center;
            //设置字体
            style.Font.Size = 22 * PT;
            style.Font.Weight = ExcelFont.BoldWeight;
            style.Font.Color =  SpreadsheetColor.FromName(ColorName.Red);
            range.Style = style;
            //增加Chart
            LineChart chart = (LineChart)sheet.Charts.Add(ChartType.Line, "B4", "J22");
            chart.Title.IsVisible = false;
            chart.Axes.Horizontal.Title.Text = "Time";
            chart.Axes.Vertical.Title.Text = "Voltage";
            ValueAxis axisY = chart.Axes.VerticalValue;
            //Y轴最大刻度与最小刻度
            axisY.Minimum = -100;
            axisY.Maximum = 100;
            //Y轴主要与次要单位大小
            axisY.MajorUnit = 20;
            axisY.MinorUnit = 10;
            //Y轴主要与次要网格是否可见
            axisY.MajorGridlines.IsVisible = true;
            axisY.MinorGridlines.IsVisible = true;
            //Y轴刻度线类型
            axisY.MajorTickMarkType = TickMarkType.Cross;
            axisY.MinorTickMarkType = TickMarkType.Inside;
            Random random = new Random();
            data = new double[LENGTH];
            for (int i = 0; i < LENGTH; i++)
            {
                if (random.Next(0, 100) > 50)
                    data[i] = random.NextDouble() * 100;
                else
                    data[i] = -random.NextDouble() * 100;
            }
            chart.Series.Add("Random", data);

            //尾部信息
            range = sheet.Cells.GetSubrange("B23", "J24");
            range.Value = $"Write Time:{DateTime.Now:yyyy-MM-dd HH:mm:ss} By CNXY";
            range.Merged = true;
            //B25(三种单元格模式)
            sheet.Cells["B25"].Value = "http://www.cnc6.cn";
            sheet.Cells[24, 1].Style.FillPattern.PatternStyle = FillPatternStyle.Solid;
            sheet.Rows[24].Cells[1].Style.FillPattern.PatternForegroundColor = SpreadsheetColor.FromName(ColorName.Red);
            //B25,J25
            sheet.Cells.GetSubrangeAbsolute(24, 1, 24, 9).Merged = true;
            try
            {
                excel.Save(filePath);
                Process.Start(filePath);
                MessageBox.Show("Write successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            MessageBox.Show("Press any key to continue.");
        }

    }


}
