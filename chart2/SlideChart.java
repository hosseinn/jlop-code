
// SlideChart.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, October 2015

/* Features:
      - create impress document
      - create a chart and add it to the first slide
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.drawing.*;
import com.sun.star.text.*;

import com.sun.star.beans.*;
import com.sun.star.awt.*;
import com.sun.star.container.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;
import com.sun.star.chart2.*;




public class SlideChart
{


  public static void main (String args[])
  {
    XComponentLoader loader = Lo.loadOffice();

    boolean hasChart = makeColChart(loader, "chartsData.ods");

    XComponent doc = Draw.createImpressDoc(loader);
    if (doc == null) {
      System.out.println("Draw doc creation failed");
      Lo.closeOffice();
      return;
    }
    GUI.setVisible(doc, true);

    XDrawPage currSlide = Draw.getSlide(doc, 0);   // access first page
    XText body = Draw.bulletsSlide(currSlide, "Sneakers Are Selling!");
    Draw.addBullet(body, 0, "Sneaker profits have increased");

    if (hasChart) {
      Lo.delay(1000);
      Lo.dispatchCmd("Paste"); 
    }

    // Draw.showShapesInfo(currSlide);
    XShape oleShape = 
           Draw.findShapeByType(currSlide, "com.sun.star.drawing.OLE2Shape");
    if (oleShape != null) {
      Size slideSize = Draw.getSlideSize(currSlide);
      Size shapeSize = Draw.getSize(oleShape);
      Point shapePos = Draw.getPosition(oleShape);

      int y = slideSize.Height - shapeSize.Height - 20;
      Draw.setPosition(oleShape, shapePos.X, y);    // move pic down
    }

    Lo.saveDoc(doc, "makeslide.odp");
    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()



  private static boolean makeColChart(XComponentLoader loader, String fnm)
  // draw a column chart; uses "Sneakers Sold this Month" table
  // same method as in TextChart.java apart from saveImage() at end
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

    Images.saveImage(Chart2.getChartImage(sheet), "chartImage.png");

    Lo.closeDoc(ssdoc);
    return true;
  }  // end of makeColChart()


} // end of SlideChart class

