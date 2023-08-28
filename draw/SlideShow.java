
// SlideShow.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/* Create a slide show.
   Three slides are created which use slide transitions and animation
   effects on some of the shapes.

   Slide show starts and user can click on the third slide shapes to jump around 
   the show.
*/

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.text.*;
import com.sun.star.drawing.*;

import com.sun.star.awt.*;
import com.sun.star.beans.*;
import com.sun.star.container.*;
import com.sun.star.presentation.*;



public class SlideShow
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

    // create 3 pages
    while (Draw.getSlidesCount(doc) < 3)
      Draw.addSlide(doc);

    // ----- the first page
    XDrawPage currSlide = Draw.getSlide(doc, 0);
    Draw.setTransition(currSlide, FadeEffect.FADE_FROM_RIGHT, AnimationSpeed.FAST, Draw.AUTO_CHANGE, 1); 
                                     // slide uses fast fade from right
    // draw a square at the top left of the page; and text
    XShape sq1 = Draw.drawRectangle(currSlide, 10, 10, 50, 50);
    Props.setProperty(sq1, "Effect", AnimationEffect.WAVYLINE_FROM_BOTTOM);  
                             // square appears in 'wave' (pixels by pixels)
    Draw.drawText(currSlide, "Page 1", 70, 20, 60, 30, 24);


    // ----- the second page
    currSlide = Draw.getSlide(doc, 1);
    Draw.setTransition(currSlide, FadeEffect.FADE_FROM_RIGHT, AnimationSpeed.FAST, Draw.AUTO_CHANGE, 1); 
                                     // slide uses fast fade from right
    // draw a circle at the bottom right of second page; and text
    XShape circle1 = Draw.drawEllipse(currSlide, 210, 150, 50, 50);
    Props.setProperty(circle1, "Effect", AnimationEffect.HIDE);  
                                                // hide circle after drawing
    Draw.drawText(currSlide, "Page 2", 170, 170, 60, 30, 24);

    String nameSlide = "page two";
    Draw.setName(currSlide, nameSlide);


    // ----- the third page
    currSlide = Draw.getSlide(doc, 2);
    Draw.setTransition(currSlide, FadeEffect.ROLL_FROM_LEFT, AnimationSpeed.MEDIUM, Draw.AUTO_CHANGE, 2);
                                        // slide uses medium roll from left
    Draw.drawText(currSlide, "Page 3", 120, 75, 60, 30, 24);

    // draw a circle containing text
    XShape circle2 = Draw.drawEllipse(currSlide, 10, 8, 50, 50);
    Draw.addText(circle2, "Click to go\nto Page 1");
    Draw.setGradientColor(circle2, "Radial red/yellow");
    Props.setProperty(circle2, "Effect", AnimationEffect.FADE_FROM_BOTTOM); // fade in
    Props.setProperty(circle2, "OnClick", ClickAction.FIRSTPAGE);
       // clicking makes the presentation jump to page one

    // draw a square with text
    XShape sq2 = Draw.drawRectangle(currSlide, 220, 8, 50, 50);
    Draw.addText(sq2, "Click to go\nto Page 2");
    Draw.setGradientColor(sq2, "Radial red/yellow");
    Props.setProperty(sq2, "Effect", AnimationEffect.FADE_FROM_BOTTOM);  // fade in
    Props.setProperty(sq2, "OnClick", ClickAction.BOOKMARK);
    Props.setProperty(sq2, "Bookmark", nameSlide);
       // clicking makes the presentation jump to page two by using a bookmark


    GUI.setVisible(doc, true);
       // slideshow start() crashes if the doc is not visible

    XPresentation2 show = Draw.getShow(doc);
    Props.showObjProps("Slide show", show);
        // an endless full-screen slide show
    show.start();


    // Lo.delay(20000);
    // slideShow.end();

    //Lo.closeDoc(doc);
    //Lo.closeOffice();
  }  // end of main()


}  // end of SlideShow class
