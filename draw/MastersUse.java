
// MastersUse.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/* - list the shapes on the default master page.
   - add shape and text to the master page
   - set the footer text
   - have normal slides use the slide number and footer on the master page

   - create a second master page
   - link one of the slides to the second master page
*/


import java.awt.Point;
import java.awt.Rectangle;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.drawing.*;

import com.sun.star.beans.*;
import com.sun.star.awt.*;
import com.sun.star.container.*;

import com.sun.star.view.*;



public class MastersUse
{

  public static void main(String args[])
  {
    XComponentLoader loader = Lo.loadOffice();

    XComponent doc = Draw.createImpressDoc(loader);
    if (doc == null) {
      System.out.println("Impress doc creation failed");
      Lo.closeOffice();
      return;
    }
 
    // report on the shapes on the default master page
    XDrawPage masterPage = Draw.getMasterPage(doc, 0);
    System.out.println("Default Master Page");
    Draw.showShapesInfo(masterPage);

    // set the master page's footer text
    Draw.setMasterFooter(masterPage, "Master Use Slides");

    // add a rectangle and text to the default master page
    // at the top-left of the slide
    Size sz = Draw.getSlideSize(masterPage);
    Draw.drawRectangle(masterPage, 5, 7, sz.Width/6, sz.Height/6);
    Draw.drawText(masterPage, "Default Master Page", 10, 15, 100, 10, 24);

    /* set slide 1 to use the master page's slide number
       but its own footer text */
    XDrawPage slide1 = Draw.getSlide(doc, 0);
    Draw.titleSlide(slide1, "Slide 1", "");
    Props.setProperty(slide1, "IsPageNumberVisible", true);
      // use the master page's slide number

    Props.setProperty(slide1, "IsFooterVisible", true);
    Props.setProperty(slide1, "FooterText", "MU Slides");
       // change the master page's footer for first slide; 
       // does not work if the master already has a footer

   /* add three more slides, which use the master page's 
       slide number and footer */
   for(int i=1; i < 4; i++) {
      XDrawPage slide = Draw.insertSlide(doc, i);
      Draw.bulletsSlide(slide, "Slide " + (i+1));
      Props.setProperty(slide, "IsPageNumberVisible", true);
      Props.setProperty(slide, "IsFooterVisible", true);
    }

    // create master page 2
    XDrawPage masterPage2 = Draw.insertMasterPage(doc, 1);
    Draw.addSlideNumber(masterPage2);

    System.out.println("Master Page 2");
    Draw.showShapesInfo(masterPage2);

    // link master page 2 to third slide
    Draw.setMasterPage(Draw.getSlide(doc, 2), masterPage2);

    // put ellipse and text on master page 2
    XShape ellipse = Draw.drawEllipse(masterPage2, 5, 7, sz.Width/6, sz.Height/6);
    Props.setProperty(ellipse, "FillColor", 0xCCFF33);   // neon green
    Draw.drawText(masterPage2, "Master Page 2", 10, 15, 100, 10, 24);

    // GUI.setVisible(doc, true);

    Lo.saveDoc(doc, "muSlides.odp");
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of MastersUse class
