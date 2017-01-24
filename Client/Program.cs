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

            chart = sl.CreateChart("C2", "D10");
            chart.SetChartType(SLScatterChartType.ScatterWithOnlyMarkers);
            chart.SetChartPosition(1, 9, 1 + fChartHeight, 9 + fChartWidth);
            sl.InsertChart(chart);

            chart = sl.CreateChart("C2", "D10");
            chart.SetChartType(SLScatterChartType.ScatterWithStraightLines);
            chart.SetChartStyle(SLChartStyle.Style15);
            chart.SetChartPosition(11, 1, 11 + fChartHeight, 1 + fChartWidth);
            sl.InsertChart(chart);

            chart = sl.CreateChart("C2", "D10");
            chart.SetChartType(SLScatterChartType.ScatterWithSmoothLinesAndMarkers);
            chart.SetChartStyle(SLChartStyle.Style45);
            chart.SetChartPosition(16, 9, 16 + fChartHeight, 9 + fChartWidth);
            sl.InsertChart(chart);

            sl.SaveAs("ChartsScatter.xlsx");

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
