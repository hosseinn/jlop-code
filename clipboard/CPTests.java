
// CPTests.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Demonstrates the usage of Office's clipboard service.
   Using the Clip.java class, add and retrieve:
      * text
      * a BufferedImage
      * add an image file; retrieve the image

   See JCPTexts.java for examples using the Java clipboard API.

   A good clipboard tool: ClCl
      http://www.nakka.com/soft/clcl/index_eng.html
*/

import java.io.*;
import java.awt.image.*;
import javax.imageio.*;

import com.sun.star.uno.*;
import com.sun.star.frame.*;
import com.sun.star.text.*;


public class CPTests
{

  public static void main(String[] args)
  {
    Lo.loadOffice();

    // add/retrieve text
    Clip.setText(Lo.getTimeStamp());
    System.out.println("Added text to clipboard");
    System.out.println("Read clipboard: " + Clip.getText());
    System.out.println();


    // add/retrieve a BufferedImage
    BufferedImage im = Images.loadImage("skinner.png");
    System.out.println("Image (w,h): " + im.getWidth() + ", " +
                                         im.getHeight());
    Clip.setImage(im);
    System.out.println("Added image to clipboard");

    BufferedImage imCopy = Clip.getImage();
    if (imCopy != null)
      System.out.println("Image (w,h): " + imCopy.getWidth() + ", " +
                                         imCopy.getHeight());
    System.out.println();


    // add an image file; retrieve the image
    pasteCopyImFile("skinner.jpg");

    Lo.closeOffice();
  }  // end of main()



  private static void pasteCopyImFile(String fnm)
  // add an image file; retrieve the image
  {
    Clip.setFile(fnm);
    System.out.println("Added file to clipboard");

    BufferedImage imCopy = Clip.getImage();
    if (imCopy != null)
      System.out.println("Image (w,h): " + imCopy.getWidth() + ", " +
                                         imCopy.getHeight());
    System.out.println();

    byte[] imData = Clip.getFile(fnm);
    System.out.println("Image byte length: " + imData.length);
  }  // end of pasteCopyImFile()



}  // end of CPTests class
