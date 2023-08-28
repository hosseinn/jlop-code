
// TextChart.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2015

/* Features:
      - create writer document
      - create a chart and add it to the first page
*/


import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.text.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;
import com.sun.star.chart2.*;





public class TextChart
{

  public static void main(String[] args)
  {
    XComponentLoader loader = Lo.loadOffice();

    boolean hasChart = makeColChart(loader, "chartsData.ods");

    XTextDocument doc = Write.createDoc(loader);
    if (doc == null) {
      System.out.println("Writer doc creation failed");
      Lo.closeOffice();
      return;
    }
    GUI.setVisible(doc, true);    // to make the construction visible

    XTextCursor cursor = Write.getCursor(doc);
    cursor.gotoEnd(false);   // make sure at end of doc before appending

    Write.appendPara(cursor, "Hello LibreOffice.\n");

    if (hasChart) {
      Lo.delay(1000);
      Lo.dispatchCmd("Paste"); 
    }

    Write.appendPara(cursor, "Figure 1. Sneakers Column Chart.\n");
    Write.stylePrevParagraph(cursor, "ParaAdjust", 
                         com.sun.star.style.ParagraphAdjust.CENTER);

    Write.appendPara(cursor, "Some more text...\n");

    Lo.saveDoc(doc, "hello.odt");

    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()




  private static boolean makeColChart(XComponentLoader loader, String fnm)
  // draw a column chart;
  // uses "Sneakers Sold this Month" table
  {
    XSpreadsheetDocument ssdoc = Calc.openDoc(fnm, loader);
    if (ssdoc == null) {
      System.out.println("Could not open " + fnm);
      return false;
    }
    GUI.setVisible(ssdoc, true);  // or selection not copied
    XSpreadsheet sheet = Calc.getSheet(ssdoc, 0);

    CellRangeAddress rangeAddr = Calc.getAddress(sheet, "A2:B8");
    XChartDocument chartDoc =
            Chart2.insertChart(sheet, rangeAddr, "C3", 15, 11, "Column");
    // Calc.gotoCell(ssdoc, "A1");

    Chart2.setTitle(chartDoc, Calc.getString(sheet, "A1"));
    Chart2.setXAxisTitle(chartDoc, Calc.getString(sheet, "A2"));
    Chart2.setYAxisTitle(chartDoc, Calc.getString(sheet, "B2"));
    Chart2.rotateYAxisTitle(chartDoc, 90);

    Lo.delay(1000);
    Chart2.copyChart(ssdoc, sheet);
    Lo.closeDoc(ssdoc);
    return true;
  }  // end of makeColChart()



}  // end of TextChart class

