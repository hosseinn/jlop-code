
// MakeSlides.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/* Generate a 5-slide presentation:
     1. title + subtitle
     2. bullets and image (skinner.png)
     3. title and video  (clock.avi or wildlife.wmv)
     4. title and rectangle (button) for playing a video (wildlife.wmv)
        and a rounded button back to start
     5. fifth slide: title and various dispatched shapes -- hacky
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.awt.*;

import com.sun.star.text.*;
import com.sun.star.drawing.*;
import com.sun.star.container.*;

import com.sun.star.presentation.*;



public class MakeSlides
{

  public static void main (String args[])
  {
    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Draw.createImpressDoc(loader);  // Impress doc
    if (doc == null) {
      System.out.println("Impress doc creation failed");
      Lo.closeOffice();
      return;
    }


    // first slide: title + subtitle
    XDrawPage currSlide = Draw.getSlide(doc, 0);
    Draw.titleSlide(currSlide, "Java-Generated Slides", "Using LibreOffice");


    currSlide = Draw.addSlide(doc);    // second slide
    doBullets(currSlide);

    // third slide: title and video
    currSlide = Draw.addSlide(doc);
    Draw.titleOnlySlide(currSlide, "Clock Video");
    Draw.drawMedia(currSlide, "clock.avi", 20, 70, 50, 50);
                           // "wildlife.wmv",
                  // no snapshot of wmv shown, only question mark, but plays;
                  // (x, y) is sometimes wrong in slide show;
                  // causes Office to crash on exiting


    currSlide = Draw.addSlide(doc);   // fourth slide
    buttonShapes(currSlide);

    // dispatchShapes(doc);              // fifth slide 
      // slow to render, hacky, so commented out

    System.out.println("Total no. of slides: " + Draw.getSlidesCount(doc));

    Lo.saveDoc(doc, "makeslides.odp");
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()




  private static void doBullets(XDrawPage currSlide)
  // second slide: bullets and image
  {
    XText body = Draw.bulletsSlide(currSlide, "What is an Algorithm?");

    // bullet levels are 0, 1, 2,...
    Draw.addBullet(body, 0, "An algorithm is a finite set of unambiguous instructions for solving a problem.");
    Draw.addBullet(body, 1, "An algorithm is correct if on all legitimate inputs, it outputs the right answer in a finite amount of time");

    Draw.addBullet(body, 0, "Can be expressed as");
    Draw.addBullet(body, 1, "pseudocode");
    Draw.addBullet(body, 1, "flow charts");
    Draw.addBullet(body, 1, "text in a natural language (e.g. English)");
    Draw.addBullet(body, 1, "computer code");

    // add the image
    XShape im = Draw.drawImageOffset(currSlide, "skinner.png", 0.6, 0.5);
                       // in bottom right corner, and scaled if necessary
    Draw.moveToBottom(currSlide, im);    // below the slide text

  }  // end of doBullets()



  private static void buttonShapes(XDrawPage currSlide)
  /* fourth slide: title and rectangle (button) for playing a video
     and a rounded button back to start
  */
  {
    Draw.titleOnlySlide(currSlide, "Wildlife Video Via Button");

    // button in the center of the slide
    com.sun.star.awt.Size sz = Draw.getSlideSize(currSlide);
    int width = 80;
    int height = 40;
    XShape ellipse = Draw.drawEllipse(currSlide, (sz.Width-width)/2, 
                                        (sz.Height-height)/2, width, height);

    Draw.addText(ellipse, "Start Video", 30);
    Props.setProperty(ellipse, "OnClick", ClickAction.DOCUMENT);
    Props.setProperty(ellipse, "Bookmark", FileIO.fnmToURL("wildlife.wmv"));
                            // mp4's not shown and don't play

    // Props.setProperty(ellipse, "Effect", AnimationEffect.FADE_FROM_BOTTOM); // fade in
                                      // AnimationEffect.HIDE);
    //Props.setProperty(ellipse, "Speed", AnimationSpeed.SLOW);

    // draw a rounded rectangle with text
    XShape button = Draw.drawRectangle(currSlide, 
                      sz.Width-width-5, sz.Height-height-5, width, height);
    Draw.addText(button, "Click to go\nto Slide 1");
    Draw.setGradientColor(button, "Radial red/yellow");
    Props.setProperty(button, "CornerRadius", 300);  // in 1/100 mm units

    Props.setProperty(button, "OnClick", ClickAction.FIRSTPAGE);
       // clicking makes the presentation jump to first slide
/*
    Props.setProperty(button, "OnClick", ClickAction.PROGRAM);
    Props.setProperty(button, "Bookmark", 
                 FileIO.fnmToURL(System.getenv("SystemRoot") + "\\System32\\calc.exe"));
*/
  }  // end of buttonShapes()



  private static void dispatchShapes(XComponent doc)
  /* fifth slide: title and two rows of various dispatched shapes 
        -- hacky and slow drawing
  */
  {
    XDrawPage currSlide = Draw.addSlide(doc);
    Draw.titleOnlySlide(currSlide, "Dispatched Shapes");

    GUI.setVisible(doc, true);
    Lo.wait(1000);

    Draw.gotoPage(doc, currSlide);
    // Lo.wait(1000);
    System.out.println("Viewing Slide number: " + Draw.getSlideNumber(Draw.getViewedPage(doc)));

    // first row
    XShape dShape = Draw.addDispatchShape(currSlide, "BasicShapes.diamond", 20, 60, 50, 30);

    Draw.addDispatchShape(currSlide, "HalfSphere", 80, 60, 50, 30);  // 3D

    dShape = Draw.addDispatchShape(currSlide, "CalloutShapes.cloud-callout", 140, 60, 50, 30);
    Draw.setBitmapColor(dShape, "Sky");

    dShape = Draw.addDispatchShape(currSlide, "FlowChartShapes.flowchart-card", 200, 60, 50, 30);
    Draw.setHatchingColor(dShape, "Black -45 degrees");

    // second row
    dShape = Draw.addDispatchShape(currSlide, "StarShapes.star12", 20, 140, 40, 40);
    Draw.setGradientColor(dShape, "Radial red/yellow");
    Props.setProperty(dShape, "LineStyle", LineStyle.NONE);   // no outline

    dShape = Draw.addDispatchShape(currSlide, "SymbolShapes.heart", 80, 140, 40, 40);
    Props.setProperty(dShape, "FillColor", 0xFF0000);

    Draw.addDispatchShape(currSlide, "ArrowShapes.left-right-arrow", 140, 140, 50, 30);   // Block Arrow menu

    dShape = Draw.addDispatchShape(currSlide, "Cyramid", 200, 120, 50, 50);    // 3D pyramid, misspelt
    Draw.setBitmapColor(dShape, "Stone");

/*
  com.sun.star.drawing.Shape3DSceneObject 
  com.sun.star.drawing.Shape3DCubeObject 
  com.sun.star.drawing.Shape3DSphereObject 
  com.sun.star.drawing.Shape3DLatheObject 
  com.sun.star.drawing.Shape3DExtrudeObject 
  com.sun.star.drawing.Shape3DPolygonObject 
*/
/*
    XShape shape3D = Draw.addShape(currSlide, "Shape3DCubeObject", 120, 120, 60, 60);
                           // nothing appears on the slide, but object is there;
                           // position, width and height can not be changed;
                           // causes Office to crash when exiting
    Rectangle rect = new Rectangle();
    rect.X = 0;
    rect.Y = 0;
    rect.Width = 6000;
    rect.Height = 6000;
    // Props.setProperty(shape3D, "BoundRect", rect);
    Draw.setSize(shape3D, 10, 10);
    Props.showObjProps("3D Shape", shape3D);
*/
    Draw.showShapesInfo(currSlide);

  }  // end of dispatchShapes()


} // end of MakeSlides class

