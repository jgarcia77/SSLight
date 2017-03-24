using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;
using SpreadsheetLight.Charts;

namespace Client
{
    class Program
    {
        static void Main(string[] args)
        {
            //Residuals();

            //ChartDataPoints();

            FillPatterns();

            Console.WriteLine("End of program");
            Console.ReadLine();
        }

        static void FillPatterns()
        {
            SLDocument sl = new SLDocument();

            sl.SetCellValue(5, 5, "Make a pattern");

            SLStyle style = sl.CreateStyle();
            // solid pattern, foreground is red, background is blue
            // But it's more complicated than this. See explanation below...
            style.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.Red, System.Drawing.Color.Blue);
            sl.SetCellStyle(6, 5, style);

            // The typical use of fill is a solid fill with a colour.
            // Internally, Excel uses a solid fill and wait for it... the foreground colour
            // property to store it.
            // The background colour property only comes into play when the fill type
            // is not a solid fill, say a hatch pattern.
            // So for typical use of solid fills, you can just set the pattern type
            // (with the solid enum value) and the foreground colour.
            // Like so:
            //style.Fill.SetPatternType(PatternValues.Solid);
            //style.Fill.SetPatternForegroundColor(SLThemeColorIndexValues.Accent1Color);

            // There are shortcut functions for setting normal pattern and gradient fills.

            style = sl.CreateStyle();
            // dark trellis pattern, foreground is accent 2 colour, background is accent 5 colour
            style.SetPatternFill(PatternValues.DarkTrellis, SLThemeColorIndexValues.Accent2Color, SLThemeColorIndexValues.Accent5Color);
            // Alternatively, use the "long-form" version:
            // style.Fill.SetPattern(PatternValues.DarkTrellis, SLThemeColorIndexValues.Accent2Color, SLThemeColorIndexValues.Accent5Color);
            sl.SetCellStyle(8, 5, style);

            style = sl.CreateStyle();
            // The SLGradientShadingStyleValues enumeration follows Excel gradient options.
            // DiagonalUp3 means interpolate from top-left to bottom-right,
            // using the 1st colour at the top-left, the 2nd color in the middle, and
            // then use the 1st colour at the bottom-right.
            // In this case, the 1st colour is accent 1 colour, and the
            // 2nd colour is the accent 2 colour.
            style.SetGradientFill(SLGradientShadingStyleValues.DiagonalUp3, SLThemeColorIndexValues.Accent1Color, SLThemeColorIndexValues.Accent2Color);
            // Alternatively, use the "long-form" version:
            // style.Fill.SetGradient(SLGradientShadingStyleValues.DiagonalUp3, SLThemeColorIndexValues.Accent1Color, SLThemeColorIndexValues.Accent2Color);
            sl.SetCellStyle(10, 5, style);

            // now for some really fancy gradients...

            style = sl.CreateStyle();
            // set linear interpolation, horizontal interpolation (0),
            // start from left (0) to right (1),
            // don't care about top (null) and bottom (null)
            style.Fill.SetCustomGradient(GradientValues.Linear, 0, 0, 1, null, null);
            // interpolation starts with accent 1 colour
            style.Fill.AppendGradientStop(0.0, SLThemeColorIndexValues.Accent1Color);
            // at midpoint (0.5), set CadetBlue as colour
            style.Fill.AppendGradientStop(0.5, System.Drawing.Color.CadetBlue);
            // at 80% of the way, set accent 5 colour
            style.Fill.AppendGradientStop(0.8, SLThemeColorIndexValues.Accent5Color);
            // at the end, set accent 2 colour, darkened 40%
            style.Fill.AppendGradientStop(1.0, SLThemeColorIndexValues.Accent2Color, -0.4);
            sl.SetCellStyle(12, 5, style);

            sl.SaveAs("Patterns.xlsx");
        }

        static void ChartDataPoints()
        {
            SLDocument sl = new SLDocument();

            sl.SetCellValue("C2", "Apple");
            sl.SetCellValue("D2", "Banana");
            sl.SetCellValue("E2", "Cherry");
            sl.SetCellValue("F2", "Durian");
            sl.SetCellValue("G2", "Elderberry");
            sl.SetCellValue("B3", "North");
            sl.SetCellValue("B4", "South");
            sl.SetCellValue("B5", "East");
            sl.SetCellValue("B6", "West");

            Random rand = new Random();
            for (int i = 3; i <= 6; ++i)
            {
                for (int j = 3; j <= 7; ++j)
                {
                    sl.SetCellValue(i, j, 9000 * rand.NextDouble() + 1000);
                }
            }

            double fChartHeight = 15.0;
            double fChartWidth = 7.5;

            SLChart chart;
            SLDataPointOptions dpoptions;

            //chart = sl.CreateChart("B2", "G6");
            //chart.SetChartType(SLColumnChartType.ClusteredColumn);
            //chart.SetChartPosition(1, 9, 1 + fChartHeight, 9 + fChartWidth);

            //dpoptions = chart.CreateDataPointOptions();
            //// 45 degrees, so it's top-left corner to bottom-right corner
            //dpoptions.Fill.SetLinearGradient(SpreadsheetLight.Drawing.SLGradientPresetValues.Fire, 45);
            //// 0 tint, 0 transparency
            //dpoptions.Line.SetSolidLine(SLThemeColorIndexValues.Accent5Color, 0, 0);
            //// 3.5 point
            //dpoptions.Line.Width = 3.5m;
            //// 3rd data series, 4th data point
            //chart.SetDataPointOptions(3, 4, dpoptions);

            //sl.InsertChart(chart);

            chart = sl.CreateChart("B2", "G6");
            chart.SetChartType(SLLineChartType.StackedLine);
            chart.SetChartPosition(7, 1, 7 + fChartHeight, 1 + fChartWidth);

            dpoptions = chart.CreateDataPointOptions();
            dpoptions.Marker.Symbol = DocumentFormat.OpenXml.Drawing.Charts.MarkerStyleValues.Triangle;
            dpoptions.Marker.Size = 10;
            // 0 tint, 0 transparency
            dpoptions.Marker.Fill.SetSolidFill(SLThemeColorIndexValues.Accent6Color, 0, 0);
            // 1st data series, 5th data point
            chart.SetDataPointOptions(1, 5, dpoptions);

            sl.InsertChart(chart);

            //chart = sl.CreateChart("B2", "G6");
            //chart.SetChartType(SLPieChartType.Pie);
            //chart.SetChartPosition(16, 9, 16 + fChartHeight, 9 + fChartWidth);

            //dpoptions = chart.CreateDataPointOptions();
            //dpoptions.Explosion = 250;
            //dpoptions.Fill.SetRadialGradient(SpreadsheetLight.Drawing.SLGradientPresetValues.Rainbow2, SpreadsheetLight.Drawing.SLGradientDirectionValues.CenterToTopLeftCorner);
            //// it's a pie chart, so only the 1st data series is used.
            //// Then we set it on the 3rd data point.
            //chart.SetDataPointOptions(1, 3, dpoptions);

            //sl.InsertChart(chart);

            sl.SaveAs("ChartsDataPoints.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }

        static void Residuals()
        {
            SLDocument sl = new SLDocument();

            sl.SetCellValue("A1", "Position");
            sl.SetCellValue("B1", "Observation");

            sl.SetCellValue(2, 1, 1.952); sl.SetCellValue(2, 2, 1);
            sl.SetCellValue(3, 1, -5.395); sl.SetCellValue(3, 2, 2);
            sl.SetCellValue(4, 1, -0.374); sl.SetCellValue(4, 2, 3);
            sl.SetCellValue(5, 1, 0.419); sl.SetCellValue(5, 2, 4);
            sl.SetCellValue(6, 1, -3.819); sl.SetCellValue(6, 2, 5);
            sl.SetCellValue(7, 1, -4.342); sl.SetCellValue(7, 2, 6);
            sl.SetCellValue(8, 1, 2.303); sl.SetCellValue(8, 2, 7);
            sl.SetCellValue(9, 1, -7.879); sl.SetCellValue(9, 2, 8);
            sl.SetCellValue(10, 1, -2.027); sl.SetCellValue(10, 2, 9);
            sl.SetCellValue(11, 1, 1.816); sl.SetCellValue(11, 2, 10);
            sl.SetCellValue(12, 1, -1.238); sl.SetCellValue(12, 2, 11);
            sl.SetCellValue(13, 1, 0.722); sl.SetCellValue(13, 2, 12);
            sl.SetCellValue(14, 1, -2.997); sl.SetCellValue(14, 2, 13);
            sl.SetCellValue(15, 1, 0.028); sl.SetCellValue(15, 2, 14);
            sl.SetCellValue(16, 1, 1.036); sl.SetCellValue(16, 2, 15);
            sl.SetCellValue(17, 1, 1.54); sl.SetCellValue(17, 2, 16);
            sl.SetCellValue(18, 1, 1.843); sl.SetCellValue(18, 2, 17);
            sl.SetCellValue(19, 1, 2.045); sl.SetCellValue(19, 2, 18);
            sl.SetCellValue(20, 1, 2.189); sl.SetCellValue(20, 2, 19);
            sl.SetCellValue(21, 1, 2.297); sl.SetCellValue(21, 2, 20);
            sl.SetCellValue(22, 1, 2.381); sl.SetCellValue(22, 2, 21);
            sl.SetCellValue(23, 1, 2.448); sl.SetCellValue(23, 2, 22);
            sl.SetCellValue(24, 1, 2.503); sl.SetCellValue(24, 2, 23);
            sl.SetCellValue(25, 1, 2.549); sl.SetCellValue(25, 2, 24);
            sl.SetCellValue(26, 1, 1.249); sl.SetCellValue(26, 2, 25);
            sl.SetCellValue(27, 1, 1.985); sl.SetCellValue(27, 2, 26);
            sl.SetCellValue(28, 1, 1.411); sl.SetCellValue(28, 2, 27);
            sl.SetCellValue(29, 1, -0.321); sl.SetCellValue(29, 2, 28);
            sl.SetCellValue(30, 1, -0.221); sl.SetCellValue(30, 2, 29);
            sl.SetCellValue(31, 1, -2.208); sl.SetCellValue(31, 2, 30);
            sl.SetCellValue(32, 1, -1.247); sl.SetCellValue(32, 2, 31);
            sl.SetCellValue(33, 1, -3.692); sl.SetCellValue(33, 2, 32);
            sl.SetCellValue(34, 1, 2.962); sl.SetCellValue(34, 2, 33);
            sl.SetCellValue(35, 1, 1.866); sl.SetCellValue(35, 2, 34);
            sl.SetCellValue(36, 1, 5.463); sl.SetCellValue(36, 2, 35);
            sl.SetCellValue(37, 1, 0.394); sl.SetCellValue(37, 2, 36);

            SLChart chart;

            chart = sl.CreateChart("A1", "B37");
            chart.SetChartType(SLScatterChartType.ScatterWithSmoothLinesAndMarkers);
            chart.SetChartStyle(SLChartStyle.Style45);
            chart.SetChartPosition(1, 3, 37, 10);
            sl.InsertChart(chart);

            sl.SaveAs("Residuals.xlsx");
        }

        static void Create1()
        {
            SLDocument sl = new SLDocument();

            sl.SetCellValue("C2", "X");
            sl.SetCellValue("D2", "Y");

            Random rand = new Random();
            for (int i = 3; i <= 10; ++i)
            {
                sl.SetCellValue(i, 3, 5 * rand.NextDouble());
                sl.SetCellValue(i, 4, 5 * rand.NextDouble());
            }

            double fChartHeight = 15.0;
            double fChartWidth = 7.5;

            SLChart chart;

            //chart = sl.CreateChart("C2", "D10");
            //chart.SetChartType(SLScatterChartType.ScatterWithOnlyMarkers);
            //chart.SetChartPosition(1, 9, 1 + fChartHeight, 9 + fChartWidth);
            //sl.InsertChart(chart);

            //chart = sl.CreateChart("C2", "D10");
            //chart.SetChartType(SLScatterChartType.ScatterWithStraightLines);
            //chart.SetChartStyle(SLChartStyle.Style15);
            //chart.SetChartPosition(11, 1, 11 + fChartHeight, 1 + fChartWidth);
            //sl.InsertChart(chart);

            chart = sl.CreateChart("C2", "D10");
            chart.SetChartType(SLScatterChartType.ScatterWithSmoothLinesAndMarkers);
            chart.SetChartStyle(SLChartStyle.Style45);
            chart.SetChartPosition(16, 9, 16 + fChartHeight, 9 + fChartWidth);
            sl.InsertChart(chart);

            sl.SaveAs("ChartsScatter.xlsx");
        }
    }
}
