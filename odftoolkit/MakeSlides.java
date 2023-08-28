
// MakeSlides.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Jan. 2017

/* Make a new presentation made up of two slides.
*/

import java.net.*;

import org.odftoolkit.simple.*;
import org.odftoolkit.simple.presentation.*;
import org.odftoolkit.simple.draw.*;
import org.odftoolkit.simple.text.list.*;

import org.odftoolkit.simple.PresentationDocument.PresentationClass;
import org.odftoolkit.simple.presentation.Slide.SlideLayout;


public class MakeSlides
{
  public static void main(String[] args)
  {
    try {
      PresentationDocument doc = PresentationDocument.newPresentationDocument();

      // a title slide
      Slide slide1 = doc.newSlide(0, "slide1", SlideLayout.TITLE_ONLY);
      Textbox titleBox = slide1.getTextboxByUsage(PresentationClass.TITLE).get(0);
      titleBox.setTextContent("Important Slide Presentation");

      // a slide with text bullets and a picture
      Slide slide2 = doc.newSlide(1, "slide2", SlideLayout.TITLE_OUTLINE);
      titleBox = slide2.getTextboxByUsage(PresentationClass.TITLE).get(0);
      titleBox.setTextContent("Overview");

      Textbox outline = slide2.getTextboxByUsage(PresentationClass.OUTLINE).get(0);
      List txtList = outline.addList();   // two bullets
      txtList.addItem("Item 1");
      txtList.addItem("Item 2");

      Image image = Image.newImage(slide2, new URI("skinner.png"));
      FrameRectangle rect = image.getRectangle();
      rect.setX(8);     // position the image
      rect.setY(4);
      image.setRectangle(rect);

      System.out.println("Saving document to makeSlides.odp");
      doc.save("makeSlides.odp");
    }
    catch (Exception e)
    { System.out.println(e);  }
  }  // end of main()

}  // end of MakeSlides class
