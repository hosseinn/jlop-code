
// JCPTests.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Demonstrates the usage of the Java clipboard API.
   Using the JClip class, add and retrieve: 
       * text
       * BufferedImage
       * 2D array
       * an ArrayList

   See CPTexts.java for examples using the Office clipboard API.

   A good clipboard tool: ClCl
      http://www.nakka.com/soft/clcl/index_eng.html
*/

import java.io.*;
import java.util.*;
import java.awt.image.*;
import javax.imageio.*;


public class JCPTests
{

  public static void main(String[] args)
  {
    // add/retrieve text
    System.out.println();
    JClip.setText(Lo.getTimeStamp());
    System.out.println("Added text to clipboard");
    System.out.println("Read clipboard: " + JClip.getText());
    Lo.wait(1000);

    // add/retrieve a BufferedImage
    BufferedImage im = Images.loadImage("skinner.png");
    System.out.println("Image (w,h): " + im.getWidth() + ", " +
                                         im.getHeight());
    JClip.setImage(im);
    System.out.println("Added image to clipboard");

    BufferedImage imCopy = JClip.getImage();
    if (imCopy != null)
      System.out.println("Image (w,h): " + imCopy.getWidth() + ", " +
                                         imCopy.getHeight());
    Lo.wait(1000);

    // add/retrieve 2D array
    Object[][] marks = {{50,60,55,67,70},
                        {62,65,70,70,81},
                        {72,66,77,80,69} };    // no generics used
    JClip.setArray(marks);
    System.out.println("Added 2D array to clipboard");
    Object[][] arr = JClip.getArray();

    System.out.println("Read clipboard: ");
    for(int i = 0; i < arr.length; i++) {
      for(int j = 0; j < arr[i].length; j++)
        System.out.print(arr[i][j] + "  ");
      System.out.println();
    }
    Lo.wait(1000);

    // add/retrieve an ArrayList
    ArrayList<String> places = new ArrayList<>();
    places.add("London");
    places.add("Paris");
    places.add("Berlin");

    JClip.setList(places);
    System.out.println("Added list to clipboard");

    System.out.println("Read clipboard: ");
    ArrayList lst = JClip.getList();
    for (Object o : lst)
      System.out.println("  " + o);

  }  // end of main()


}  // end of JCPTests class
