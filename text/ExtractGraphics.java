
// ExtractGraphics.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2015

/* Extract images from a text file, either as XGraphic
   objects or Java BufferedImage objects.

   Usage:
     run ExtractGraphics build.odt
     run ExtractGraphics build.doc

*/


import java.util.*;
import java.awt.image.*;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.awt.*;

import com.sun.star.graphic.*;
import com.sun.star.text.*;
import com.sun.star.drawing.*;
import com.sun.star.container.*;




public class ExtractGraphics 
{

  public static void main(String[] args) 
  {
    XComponentLoader loader = Lo.loadOffice();

    XTextDocument textDoc = Write.openDoc(args[0], loader);
    if (textDoc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }

    // save text graphics to files
    ArrayList<XGraphic> pics = Write.getTextGraphics(textDoc);
    if (pics == null)
      return;
    System.out.println("\nNum. of text graphics: " + pics.size());
    int i = 1;
    for(XGraphic pic : pics) {
      Images.saveGraphic(pic, "graphics" + i + ".png", "png");   //".gif", "gif");
      Size sz = (Size) Props.getProperty(pic, "SizePixel");
      System.out.println("Image size in pixels: " + sz.Width + " x " + sz.Height);
      i++;
    }
    System.out.println();


    // this supplier is not created; queryInterface() returns null
    XTextShapesSupplier shpsSupplier = 
                       UnoRuntime.queryInterface(XTextShapesSupplier.class, textDoc);
    if (shpsSupplier == null)
      System.out.println("Could not obtain text shapes supplier");
    else
      System.out.println("Num. of text shapes: " + shpsSupplier.getShapes().getCount());


    // report on shapes in the doc
    XDrawPage drawPage = Write.getShapes(textDoc);
    ArrayList<XShape> shapes = Draw.getShapes(drawPage);
    if (shapes != null)
      System.out.println("\nNum. of draw shapes: " + shapes.size());
    for(XShape shape : shapes)
      Draw.reportPosSize(shape);
    System.out.println();

    Lo.closeDoc(textDoc);
    Lo.closeOffice();
  }  // end of main()



}  // end of ExtractGraphics class