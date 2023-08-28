
// ImageHex.java
// Andrew Davison, Oct 2016, ad@fivedots.coe.psu.ac.th

/* Convert the supplied image into hexadecimal text, which is
   printed to stdout. 

   This may be useful if Addon.xcs is storing icons as hexadecimal
   text.
*/

import java.io.*;
import java.util.*;
import java.awt.image.*;
import javax.imageio.*;



public class ImageHex
{
  static final private String HEXES = "0123456789ABCDEF";


  public static void main(String[] args)
  {
    if (args.length != 1) {
      System.out.println("Usage: java ImageHex <image filename>");
      return;
    }

    String s = getHex(im2bytes(loadImage(args[0])));
    System.out.println(s);
  }  // end of main()



  private static BufferedImage loadImage(String fnm)
  {
    BufferedImage image = null;
    try {
      image = ImageIO.read(new File(fnm));
      // System.out.println("Loaded " + fnm);
      // System.out.println("Width x Height: " + 
      //                 image.getWidth() + " x " + image.getHeight());
    }
    catch (java.io.IOException e)
    {  System.out.println("Unable to load " + fnm);  }
    return image;
  }  // end of loadImage()



  private static byte[] im2bytes(BufferedImage im)
  {
    byte[] bytes =  null;
    try {
      ByteArrayOutputStream baos = new ByteArrayOutputStream();
      ImageIO.write(im, "png", baos);
      baos.flush();
      bytes = baos.toByteArray();
      baos.close();
    }
    catch(java.io.IOException e) 
    {  System.out.println("Could not convert image to bytes");  }
    return bytes;
  }  // end of im2bytes()



  private static String getHex(byte[] raw) 
  //  http://www.rgagnon.com/javadetails/java-0596.html
  {
    if (raw == null)
      return null;
    final StringBuilder hex = new StringBuilder(2 * raw.length);
    for (final byte b : raw)
      hex.append(HEXES.charAt((b & 0xF0) >> 4))
         .append(HEXES.charAt((b & 0x0F)));
    return hex.toString();
  }  // end of getHex()

}  // end of ImageHex class
