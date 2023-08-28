
// Chart2Views.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, October 2015

/* Generate many different charts using data from
   chartsData.ods
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;
import com.sun.star.chart2.*;
import com.sun.star.drawing.*;
import com.sun.star.chart2.data.*;

// import com.sun.star.chart.*;

import com.sun.star.chart.TimeIncrement;
import com.sun.star.chart.TimeInterval;
import com.sun.star.chart.TimeUnit;




public class Chart2Views
{
  private static final String CHARTS_DATA = "chartsData.ods";


  public static void main(String args[])
  {
    XComponentLoader loader = Lo.loadOffice();

    XSpreadsheetDocument doc = Calc.openDoc(CHARTS_DATA, loader);
    if (doc == null) {
      System.out.println("Could not open " + CHARTS_DATA);
      Lo.closeOffice();
      return;
    }
    GUI.setVisible(doc, true);
    XSpreadsheet sheet = Calc.getSheet(doc, 0);
/*
    String[] fontNames = Info.getFontNames();
    Lo.printNames(fontNames);
*/

    // ---- draw different types of chart ----
    // uncomment the one you want to use

    // XChartDocument chartDoc = colChart(doc, sheet);
    // String[] templateNames = Chart2.getChartTemplates(chartDoc);
    // Lo.printNames(templateNames, 1);

    // multColChart(doc, sheet);
    // colLineChart(doc, sheet);

    // barChart(doc, sheet);

    // pieChart(doc, sheet);
    // pie3DChart(doc, sheet);
    // donutChart(doc, sheet);

    // areaChart(doc, sheet);

    // lineChart(doc, sheet);
    // linesChart(doc, sheet);

    // scatterChart(doc, sheet);
    // scatterLineLogChart(doc, sheet);
    // scatterLineErrorChart(doc, sheet);

    // labeledBubbleChart(doc, sheet);

    // netChart(doc, sheet);

    happyStockChart(doc, sheet);
    // stockPricesChart(doc, sheet);

    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  private static XChartDocument colChart(XSpreadsheetDocument doc,
                                XSpreadsheet sheet)
  // draw a column chart;
  // uses "Sneakers Sold this Month" table
  {
    CellRangeAddress rangeAddr = Calc.getAddress(sheet, "A2:B8");
    XChartDocument chartDoc =
            Chart2.insertChart(sheet, rangeAddr, "C3", 15, 11, "Column");
    Calc.gotoCell(doc, "A1");

    Chart2.setTitle(chartDoc, Calc.getString(sheet, "A1"));
    Chart2.setXAxisTitle(chartDoc, Calc.getString(sheet, "A2"));
    Chart2.setYAxisTitle(chartDoc, Calc.getString(sheet, "B2"));
    Chart2.rotateYAxisTitle(chartDoc, 90);
    return chartDoc;
  }  // end of colChart()




  private static void multColChart(XSpreadsheetDocument doc,
                                   XSpreadsheet sheet)
  // draws a multiple column chart: 2D and 3D;
  // uses "States with the Most Colleges and Universities by Type"
  {
    CellRangeAddress rangeAddr = Calc.getAddress(sheet, "E15:G21");
    XChartDocument chartDoc =
           Chart2.insertChart(sheet, rangeAddr, "A22", 20, 11, "ThreeDColumnDeep");
                                // Column ThreeDColumnDeep, ThreeDColumnFlat
                                // StackedColumn, PercentStackedColumn
    Calc.gotoCell(doc, "A13");

    Chart2.setTitle(chartDoc, Calc.getString(sheet, "E13"));
    Chart2.setXAxisTitle(chartDoc, Calc.getString(sheet, "E15"));
    Chart2.setYAxisTitle(chartDoc, Calc.getString(sheet, "F14"));

    Chart2.rotateYAxisTitle(chartDoc, 90);
    Chart2.viewLegend(chartDoc, true);

    // for the 3D versions
    Chart2.showAxisLabel(chartDoc, Chart2.Z_AXIS, 0, false);  // hide labels
    Chart2.setChartShape3D(chartDoc, "cylinder");  // box, cylinder, cone, pyramid
  }  // end of multColChart()



  private static void colLineChart(XSpreadsheetDocument doc,
                                   XSpreadsheet sheet)
  // draws a column+line chart;
  // uses "States with the Most Colleges and Universities by Type"
  {
    CellRangeAddress rangeAddr = Calc.getAddress(sheet, "E15:G21");
    XChartDocument chartDoc =
           Chart2.insertChart(sheet, rangeAddr, "A22", 20, 11, "ColumnWithLine");
    Calc.gotoCell(doc, "A13");

    Chart2.setTitle(chartDoc, Calc.getString(sheet, "E13"));
    Chart2.setXAxisTitle(chartDoc, Calc.getString(sheet, "E15"));   // column
    Chart2.setYAxisTitle(chartDoc, Calc.getString(sheet, "F14"));   // line

    Chart2.rotateYAxisTitle(chartDoc, 90);
    Chart2.viewLegend(chartDoc, true);
  }  // end of colLineChart()



  private static void barChart(XSpreadsheetDocument doc,
                                XSpreadsheet sheet)
  // uses "Sneakers Sold this Month" table
  {
    CellRangeAddress rangeAddr = Calc.getAddress(sheet, "A2:B8");
    XChartDocument chartDoc =
                  Chart2.insertChart(sheet, rangeAddr, "B3", 15, 11, "Bar");
    Calc.gotoCell(doc, "A1");

    Chart2.setTitle(chartDoc, Calc.getString(sheet, "A1"));
    Chart2.setXAxisTitle(chartDoc, Calc.getString(sheet, "A2"));
    Chart2.setYAxisTitle(chartDoc, Calc.getString(sheet, "B2"));
    Chart2.rotateXAxisTitle(chartDoc, 90);
                          // rotate x-axis which is now the vertical axis
  }  // end of barChart()




  private static void pieChart(XSpreadsheetDocument doc,
                                XSpreadsheet sheet)
  // draw a pie chart, with legend and subtitle;
  // uses "Top 5 States with the Most Elementary and Secondary Schools"
  {
    CellRangeAddress rangeAddr = Calc.getAddress(sheet, "E2:F8");
    XChartDocument chartDoc =
           Chart2.insertChart(sheet, rangeAddr, "B10", 12, 11, "Pie");
                   // Pie, ThreeDPie, PieAllExploded (ugly), ThreeDPieAllExploded
                   // Column
    Calc.gotoCell(doc, "A1");

    Chart2.setTitle(chartDoc, Calc.getString(sheet, "E1"));
    Chart2.setSubtitle(chartDoc, Calc.getString(sheet, "F2"));
    Chart2.viewLegend(chartDoc, true);
  }  // end of pieChart()




  private static void pie3DChart(XSpreadsheetDocument doc,
                                XSpreadsheet sheet)
  // draws a 3D pie chart with rotation, label change
  // uses "Top 5 States with the Most Elementary and Secondary Schools"
  {
    CellRangeAddress rangeAddr = Calc.getAddress(sheet, "E2:F8");
    XChartDocument chartDoc =
                Chart2.insertChart(sheet, rangeAddr, "B10", 12, 11, "ThreeDPie");
                   // ThreeDPie, ThreeDPieAllExploded
    Calc.gotoCell(doc, "A1");

    Chart2.setTitle(chartDoc, Calc.getString(sheet, "E1"));
    Chart2.setSubtitle(chartDoc, Calc.getString(sheet, "F2"));
    Chart2.viewLegend(chartDoc, true);


    // rotate around horizontal (x-axis) and vertical (y-axis)
    XDiagram diagram = chartDoc.getFirstDiagram();
    Props.setProperty(diagram, "RotationHorizontal", 0);
           // -ve rotates bottom edge out of page; default is -60
    Props.setProperty(diagram, "RotationVertical", -45);
           //  -ve rotates left edge out of page; default is 0 (i.e. no rotation)
    // Props.showObjProps("Diagram", diagram);

/*
    // change all the data points to be bold white 14pt
    XDataSeries[] ds = Chart2.getDataSeries(chartDoc);
    System.out.println("No. of data series: " + ds.length);
    Props.showObjProps("Data Series 0", ds[0]);
    Props.setProperty(ds[0], "CharHeight", 14.0);
    Props.setProperty(ds[0], "CharColor", Calc.WHITE);
    Props.setProperty(ds[0], "CharWeight", com.sun.star.awt.FontWeight.BOLD);
*/

    XPropertySet[] propsArr = Chart2.getDataPointsProps(chartDoc, 0);
    System.out.println("Number of data props in first data series: " + propsArr.length);

    // change only the last data point to be bold white 14pt
    XPropertySet props = Chart2.getDataPointProps(chartDoc, 0, 5);
    // Props.showProps("Data point", props);
    if (props != null) {
      Props.setProperty(props, "CharHeight", 14.0);
      Props.setProperty(props, "CharColor", Calc.WHITE);
      Props.setProperty(props, "CharWeight", com.sun.star.awt.FontWeight.BOLD);
    }
  }  // end of pie3DChart()





  private static void donutChart(XSpreadsheetDocument doc,
                                  XSpreadsheet sheet)
  // draws a 3D donut chart with 2 rings
  // uses the "Annual Expenditure on Institutions" table
  {
    CellRangeAddress rangeAddr = Calc.getAddress(sheet, "A44:C50");
    XChartDocument chartDoc =
           Chart2.insertChart(sheet, rangeAddr, "D43", 15, 11, "Donut");
                                       // Donut, DonutAllExploded, ThreeDDonut
    Calc.gotoCell(doc, "A48");

    Chart2.setTitle(chartDoc, Calc.getString(sheet, "A43"));
    Chart2.viewLegend(chartDoc, true);
    Chart2.setSubtitle(chartDoc,  "Outer: " + Calc.getString(sheet, "B44") + "\n" +
                                  "Inner: " + Calc.getString(sheet, "C44"));

    // Chart2.setDataPointLabels(chartDoc, Chart2.DP_CATEGORY);
  }  // end of donutChart()




  private static void areaChart(XSpreadsheetDocument doc,
                                   XSpreadsheet sheet)
  // draws an area (stacked) chart;
  // uses "Trends in Enrollment in Public Schools in the US" table
  {
    CellRangeAddress rangeAddr = Calc.getAddress(sheet, "E45:G50");
    XChartDocument chartDoc =
        Chart2.insertChart(sheet, rangeAddr, "A52", 16, 11, "Area");
                                       // Area, StackedArea, PercentStackedArea
    Calc.gotoCell(doc, "A43");

    Chart2.setTitle(chartDoc, Calc.getString(sheet, "E43"));
    Chart2.viewLegend(chartDoc, true);
    Chart2.setXAxisTitle(chartDoc, Calc.getString(sheet, "E45"));
    Chart2.setYAxisTitle(chartDoc, Calc.getString(sheet, "F44"));
    Chart2.rotateYAxisTitle(chartDoc, 90);
  }  // end of areaChart()




  private static void lineChart(XSpreadsheetDocument doc,
                                XSpreadsheet sheet)
  // draw a line chart with data points, no legend;
  // uses "Humidity Levels in NY" table
  {
    CellRangeAddress rangeAddr = Calc.getAddress(sheet, "A14:B21");
    XChartDocument chartDoc =
           Chart2.insertChart(sheet, rangeAddr, "D13", 16, 9, "LineSymbol");
    Calc.gotoCell(doc, "A1");

    Chart2.setTitle(chartDoc, Calc.getString(sheet, "A13"));
    Chart2.setXAxisTitle(chartDoc, Calc.getString(sheet, "A14"));
    Chart2.setYAxisTitle(chartDoc, Calc.getString(sheet, "B14"));
  }  // end of lineChart()



  private static void linesChart(XSpreadsheetDocument doc,
                                XSpreadsheet sheet)
  // draw a line chart with two lines;
  // uses "Trends in Expenditure Per Pupil"
  {
    CellRangeAddress rangeAddr = Calc.getAddress(sheet, "E27:G39");
    XChartDocument chartDoc =
           Chart2.insertChart(sheet, rangeAddr, "A40", 22, 11, "LineSymbol");
             // Line, LineSymbol, StackedLineSymbol

    Calc.gotoCell(doc, "A26");

    Chart2.setTitle(chartDoc, Calc.getString(sheet, "E26"));
    Chart2.viewLegend(chartDoc, true);
    Chart2.setXAxisTitle(chartDoc, Calc.getString(sheet, "E27"));
    Chart2.setYAxisTitle(chartDoc, "Expenditure per Pupil");
    Chart2.rotateYAxisTitle(chartDoc, 90);

    Chart2.setDataPointLabels(chartDoc, Chart2.DP_NONE);   // too crowded for data points
  }  // end of linesChart()




  private static void scatterChart(XSpreadsheetDocument doc,
                                         XSpreadsheet sheet)
  // data from http://www.mathsisfun.com/data/scatter-xy-plots.html;
  // uses the "Ice Cream Sales vs Temperature" table
  {
    CellRangeAddress rangeAddr = Calc.getAddress(sheet, "A110:B122");
    XChartDocument chartDoc =
           Chart2.insertChart(sheet, rangeAddr, "C109", 16, 11, "ScatterSymbol");
                               // ScatterSymbol ScatterLine ScatterLineSymbol
    Calc.gotoCell(doc, "A104");

    Chart2.setTitle(chartDoc, Calc.getString(sheet, "A109"));
    Chart2.setXAxisTitle(chartDoc, Calc.getString(sheet, "A110"));
    Chart2.setYAxisTitle(chartDoc, Calc.getString(sheet, "B110"));
    Chart2.rotateYAxisTitle(chartDoc, 90);

    Chart2.calcRegressions(chartDoc);

    Chart2.drawRegressionCurve(chartDoc, Chart2.LINEAR);
      // LINEAR, LOGARITHMIC, EXPONENTIAL, POWER, POLYNOMIAL, MOVING_AVERAGE

  }  // end of scatterChart()




  private static void scatterLineLogChart(XSpreadsheetDocument doc,
                                      XSpreadsheet sheet)
  // draw a x-y scatter chart using log scaling
  // uses the "Power Function Test" table
  {
    CellRangeAddress rangeAddr = Calc.getAddress(sheet, "E110:F120");
    XChartDocument chartDoc =
           Chart2.insertChart(sheet, rangeAddr, "A121", 20, 11, "ScatterLineSymbol");
    Calc.gotoCell(doc, "A121");

    Chart2.setTitle(chartDoc, Calc.getString(sheet, "E109"));
    Chart2.setXAxisTitle(chartDoc, Calc.getString(sheet, "E110"));
    Chart2.setYAxisTitle(chartDoc, Calc.getString(sheet, "F110"));
    Chart2.rotateYAxisTitle(chartDoc, 90);

    // change x- and y- axes to log scaling
    XAxis xaxis = Chart2.scaleXAxis(chartDoc, Chart2.LOGARITHMIC);
    // Chart2.printScaleData("x-axis", xaxis);
    Chart2.scaleYAxis(chartDoc, Chart2.LOGARITHMIC);
      // LINEAR, LOGARITHMIC, (EXPONENTIAL, POWER)

    Chart2.drawRegressionCurve(chartDoc, Chart2.POWER);
  }  // end of scatterLineLogChart()




  private static void scatterLineErrorChart(XSpreadsheetDocument doc,
                                    XSpreadsheet sheet)
  // draws an x-y scatter chart with lines and y-error bars
  // uses the smaller "Impact Data - 1018 Cold Rolled" table
  // the example comes from "Using Descriptive Statistics.pdf"
  {
    CellRangeAddress rangeAddr = Calc.getAddress(sheet, "A142:B146");
    XChartDocument chartDoc =
           Chart2.insertChart(sheet, rangeAddr, "F115", 14, 16, "ScatterLineSymbol");
    Calc.gotoCell(doc, "A123");

    Chart2.setTitle(chartDoc, Calc.getString(sheet, "A141"));
    Chart2.setXAxisTitle(chartDoc, Calc.getString(sheet, "A142"));
    Chart2.setYAxisTitle(chartDoc, "Impact Energy (Joules)");
    Chart2.rotateYAxisTitle(chartDoc, 90);

    System.out.println("Adding y-axis error bars");
    String sheetName = Calc.getSheetName(sheet);
    String errorLabel = sheetName + "." + "C142";
    String errorRange = sheetName + "." + "C143:C146";
    Chart2.setYErrorBars(chartDoc, errorLabel, errorRange);
  }  // end of scatterLineErrorChart()




  private static void labeledBubbleChart(XSpreadsheetDocument doc,
                                  XSpreadsheet sheet)
  // draws a bubble chart with labels;
  // uses the "World data" table
  {
    CellRangeAddress rangeAddr = Calc.getAddress(sheet, "H63:J93");
    XChartDocument chartDoc =
        Chart2.insertChart(sheet, rangeAddr, "A62", 18, 11, "Bubble");
    Calc.gotoCell(doc, "A62");

    Chart2.setTitle(chartDoc, Calc.getString(sheet, "H62"));
    Chart2.setXAxisTitle(chartDoc, Calc.getString(sheet, "H63"));
    Chart2.setYAxisTitle(chartDoc, Calc.getString(sheet, "I63"));
    Chart2.rotateYAxisTitle(chartDoc, 90);
    Chart2.viewLegend(chartDoc, true);

    // change the data points
    XDataSeries[] ds = Chart2.getDataSeries(chartDoc);
    Props.setProperty(ds[0], "Transparency", 50);   // 100 == transparent
    Props.setProperty(ds[0], "BorderStyle", LineStyle.SOLID);
    Props.setProperty(ds[0], "BorderColor", Calc.RED);
    Props.setProperty(ds[0], "LabelPlacement", Chart2.DP_CENTER);

    // Chart2.setDataPointLabels(chartDoc, Chart2.DP_NUMBER);

    String sheetName = Calc.getSheetName(sheet);
    String label = sheetName + "." + "K63";
    String names = sheetName + "." + "K64:K93";
    Chart2.addCatLabels(chartDoc, label, names);
  }  // end of labeledBubbleChart()



  private static void netChart(XSpreadsheetDocument doc,
                               XSpreadsheet sheet)
  // draws a net chart;
  // uses the "No of Calls per Day" table
  {
    CellRangeAddress rangeAddr = Calc.getAddress(sheet, "A56:D63");
    XChartDocument chartDoc =
        Chart2.insertChart(sheet, rangeAddr, "E55", 16, 11, "NetLine");
                             // Net, NetLine, NetSymbol
                             // StackedNet, PercentStackedNet
    Calc.gotoCell(doc, "E55");

    Chart2.setTitle(chartDoc, Calc.getString(sheet, "A55"));
    Chart2.viewLegend(chartDoc, true);
    Chart2.setDataPointLabels(chartDoc, Chart2.DP_NONE);

    // reverse x-axis so days increase clockwise around net
    XAxis xAxis = Chart2.getXAxis(chartDoc);
    ScaleData sd = xAxis.getScaleData();
    sd.Orientation = AxisOrientation.REVERSE;
    xAxis.setScaleData(sd);
  }  // end of netChart()



  private static void happyStockChart(XSpreadsheetDocument doc,
                                 XSpreadsheet sheet)
  // draws a fancy stock chart
  // uses the "Happy Systems (HASY)" table
  {
    CellRangeAddress rangeAddr = Calc.getAddress(sheet, "A86:F104");
    XChartDocument chartDoc =
        Chart2.insertChart(sheet, rangeAddr, "A105", 25, 14,
                                                "StockVolumeOpenLowHighClose");
        // StockLowHighClose, StockOpenLowHighClose, StockVolumeLowHighClose
        // StockVolumeOpenLowHighClose
    Calc.gotoCell(doc, "A105");

    Chart2.setTitle(chartDoc, Calc.getString(sheet, "A85"));
    Chart2.setXAxisTitle(chartDoc, Calc.getString(sheet, "A86"));
    Chart2.setYAxisTitle(chartDoc, Calc.getString(sheet, "B86"));
    Chart2.rotateYAxisTitle(chartDoc, 90);
    Chart2.setYAxis2Title(chartDoc, "Stock Value");
    Chart2.rotateYAxis2Title(chartDoc, 90);

    Chart2.setDataPointLabels(chartDoc, Chart2.DP_NONE);
    // Chart2.viewLegend(chartDoc, true);

    // change 2nd y-axis min and max; default is poor ($0 - $20)
    XAxis yAxis2 = Chart2.getYAxis2(chartDoc);
    ScaleData sd = yAxis2.getScaleData();
    // Chart2.printScaleData("Secondary Y-axis", sd);
    sd.Minimum = 83;
    sd.Maximum = 103;
    yAxis2.setScaleData(sd);

    // change x-axis type from number to date
    XAxis xAxis = Chart2.getXAxis(chartDoc);
    sd = xAxis.getScaleData();
    sd.AxisType = AxisType.DATE;

    // set major increment to 3 days
    TimeInterval ti = new TimeInterval(3, TimeUnit.DAY);
    TimeIncrement tc = new TimeIncrement();
    tc.MajorTimeInterval = ti;
    sd.TimeIncrement = tc;
    xAxis.setScaleData(sd);
/*
    // rotate the axis labels by 45 degrees
    XAxis xAxis = Chart2.getXAxis(chartDoc);
    Props.setProperty(xAxis, "TextRotation", 45);
*/
    Chart2.printChartTypes(chartDoc);

    // change color of "WhiteDay" and "BlackDay" block colors
    XChartType candleCT = Chart2.findChartType(chartDoc, "CandleStickChartType");
    // Props.showObjProps("Stock Chart", stockCT);
    Chart2.colorStockBars(candleCT, Calc.GREEN, Calc.RED);    // increase, decrease

    // thicken the high-low line; make it yellow
    XDataSeries[] ds = Chart2.getDataSeries(chartDoc, "CandleStickChartType");
    System.out.println("No of data series in candle stick chart: " + ds.length);
    // Props.showObjProps("Candle Stick", ds[0]);
    Props.setProperty(ds[0], "LineWidth", 120);   // in 1/100 mm 
    Props.setProperty(ds[0], "Color", Calc.YELLOW);

  }  // end of happyStockChart()





  private static void stockPricesChart(XSpreadsheetDocument doc,
                                        XSpreadsheet sheet)
  // draws a stock chart, with an extra pork bellies line
  // uses the "Calc Guide Stock Prices" table
  {
    CellRangeAddress rangeAddr = Calc.getAddress(sheet, "E141:I146");
    XChartDocument chartDoc =
        Chart2.insertChart(sheet, rangeAddr, "E148", 12, 11,
                                                "StockOpenLowHighClose");
    Calc.gotoCell(doc, "A139");

    Chart2.setTitle(chartDoc, Calc.getString(sheet, "E140"));
    Chart2.setDataPointLabels(chartDoc, Chart2.DP_NONE);
    Chart2.setXAxisTitle(chartDoc, Calc.getString(sheet, "E141"));
    Chart2.setYAxisTitle(chartDoc, "Dollars");
    Chart2.rotateYAxisTitle(chartDoc, 90);

    System.out.println("Adding Pork Bellies line");
    String sheetName = Calc.getSheetName(sheet);
    String porkLabel = sheetName + "." + "J141";
    String porkPoints = sheetName + "." + "J142:J146";
    Chart2.addStockLine(chartDoc, porkLabel, porkPoints);

    Chart2.viewLegend(chartDoc, true);
  }  // end of stockPricesChart()



}  // end of Chart2Views class
