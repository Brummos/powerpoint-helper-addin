using Microsoft.Office.Tools.Ribbon;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using XlAxisType = Microsoft.Office.Interop.PowerPoint.XlAxisType;
using XlAxisGroup = Microsoft.Office.Interop.PowerPoint.XlAxisGroup;
using SeriesCollection = Microsoft.Office.Interop.PowerPoint.SeriesCollection;

namespace PowerPointAddIn1 {
    public partial class Ribbon1 {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) {}

        private void MakeSameSizeBtn_Click(object sender, RibbonControlEventArgs e) {
            // Get selected objects
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

            // We start at 2 so we can immediatly begin with the conversion
            // The starting at 2 also provides a check to see if we have any shapes to convert
            for (int i = 2; i <= selection.ShapeRange.Count; i++) {
                selection.ShapeRange[i].Width = selection.ShapeRange[1].Width;
                selection.ShapeRange[i].Height = selection.ShapeRange[1].Height;
            }
        }
        
        // TODO fix function name
        private void ChartQuickAlignAndSizeBtn_Click(object sender, RibbonControlEventArgs e) {
            // Get selected objects
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

            // Original is not a chart so we should abort
            if (selection.ShapeRange[1].Type != MsoShapeType.msoChart) {
                return;
            }

            Chart originalChart = selection.ShapeRange[1].Chart;

            // We start at 2 so we can immediatly begin with the conversion
            // The starting at 2 also provides a check to see if we have any shapes to convert
            for (int i = 2; i <= selection.ShapeRange.Count; i++) {
                // Selected shape is not a chart, continue to next
                if (selection.ShapeRange[i].Type != MsoShapeType.msoChart) {
                    continue;
                }

                // Set bounding box
                selection.ShapeRange[i].Width = selection.ShapeRange[1].Width;
                selection.ShapeRange[i].Height = selection.ShapeRange[1].Height;

                // Set plot area
                Chart conversionChart = selection.ShapeRange[i].Chart;
                // TODO see if all of these are necessary
                conversionChart.PlotArea.Height = originalChart.PlotArea.Height;
                conversionChart.PlotArea.Width = originalChart.PlotArea.Width;
                conversionChart.PlotArea.Left = originalChart.PlotArea.Left;
                conversionChart.PlotArea.Top = originalChart.PlotArea.Top;
                conversionChart.PlotArea.InsideLeft = originalChart.PlotArea.InsideLeft;
                conversionChart.PlotArea.InsideTop = originalChart.PlotArea.InsideTop;
                conversionChart.PlotArea.InsideWidth = originalChart.PlotArea.InsideWidth;
                conversionChart.PlotArea.InsideHeight = originalChart.PlotArea.InsideHeight;
            }
        }

        private void TableFormatWithLayoutBtn_Click(object sender, RibbonControlEventArgs e) {
            // Get selected objects
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

            for (int i = 1; i <= selection.ShapeRange.Count; i++) {
                // Check if shape is a table
                if (selection.ShapeRange[i].Type == MsoShapeType.msoTable) {
                    Table table = selection.ShapeRange[i].Table;
                    // Set style of table to custom style
                    table.ApplyStyle("D03447BB-5D67-496B-8E87-E561075AD55C");
                }
            }
        }

        private void ChartQuickFormatColorsOnlyBtn_Click(object sender, RibbonControlEventArgs e) {
            // Get selected objects
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

            for (int i = 1; i <= selection.ShapeRange.Count; i++) {
                // Check if shape is a chart
                if (selection.ShapeRange[i].Type == MsoShapeType.msoChart) {
                    Chart chart = selection.ShapeRange[i].Chart;
                    // Set style of table to custom style
                    chart.ChartStyle = 1;
                    chart.ChartColor = 6;
                    chart.Refresh();

                 
                }
            }

        }

        private void ChartQuickFormatFontsOnlyBtn_Click(object sender, RibbonControlEventArgs e) {
            // Get selected objects
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

            for (int i = 1; i <= selection.ShapeRange.Count; i++) {
                // Check if shape is a chart
                if (selection.ShapeRange[i].Type == MsoShapeType.msoChart) {
                    Chart chart = selection.ShapeRange[i].Chart;
                    if (chart.HasLegend) {
                        //TODO make this a config option
                        chart.Legend.Font.Size = 8;
                    }

                    if (chart.HasTitle) {
                        //TODO make this a config option
                        chart.ChartTitle.Font.Size = 9;
                    }

                    //                  Fields
                    //FIELDS
                    //xlCategory  1
                    //Axis displays categories.

                    //xlSeriesAxis    3
                    //Axis displays data series.

                    //xlValue 2
                    //Axis displays values.

                    // a chart can ussually has 2 axis. value axis Y, catagory axis

                    // TODO better null check
                    Axis valueAxis = chart.Axes(XlAxisType.xlValue);
                    if (valueAxis != null) {
                        //TODO make this a config option
                        valueAxis.TickLabels.Font.Size = 8;
                    }

                    // TODO better null check
                    Axis catagoryAxis = chart.Axes(XlAxisType.xlCategory);
                    if (catagoryAxis != null) {
                        //TODO make this a config option
                        catagoryAxis.TickLabels.Font.Size = 8;
                    }


       

                    var series = chart.SeriesCollection() as SeriesCollection;
                    foreach (var ser in series) {
                        var DataLabels = ((Series) ser).DataLabels();
                        DataLabels.Format.TextFrame2.TextRange.Font.Size = 7;
                    }

                    //TODO make this a config option
                    //chart.ChartArea.Font.Size = 8;
                    //chart.PlotArea.Format.Font.Size = 8;

                    // chart..Font.Size = 50; ;

                }
            }

        }

        private void ChartQuickFormatNumberFormatsOnlyBtn_Click(object sender, RibbonControlEventArgs e) {
            // Get selected objects
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

        }

        private void ChartQuickFormatPositionsOnlyBtn_Click(object sender, RibbonControlEventArgs e) {
            // Get selected objects
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

        }
    }

}
