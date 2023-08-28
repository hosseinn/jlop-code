
// CheckLinks.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2015

/* Extract and check all the URLs in the slides
   Uses *Java's* pattern matching on extracted slide text.

   Usage
     run CheckLinks algs.ppt
*/

import java.io.*;
import java.util.*;
import java.util.regex.*;
import java.net.*;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.drawing.*;



public class CheckLinks
{
  private static final String URL_REGEX = 
          "http://[-a-zA-Z0-9+&@#/%?=~_|!:,.;]*[-a-zA-Z0-9+&@#/%=~_|]";
    // only simple http:// strings

  private static final int TIMEOUT = 5000;    // 5 secs


  public static void main(String args[])
  {
    if (args.length != 1) {
      System.out.println("Usage: CheckLinks fnm");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();
    XComponent doc = Lo.openDoc(args[0], loader);
    if (doc == null) {
      System.out.println("Could not open " + args[0]);
      Lo.closeOffice();
      return;
    }

    XDrawPages slides = Draw.getSlides(doc);
    if (slides == null) {
      System.out.println("No slides found");
      return;
    }

    Pattern pattern = Pattern.compile(URL_REGEX);

    for (int i=0; i < slides.getCount(); i++) {
      XDrawPage slide = Draw.getSlide(doc, i);
      if (slide == null) {
        System.out.println("Cound not access slide " + i);
        break;
      }
      ArrayList<XShape> shapes = Draw.getShapes(slide);
      for(XShape shape : shapes) {
        String text = Draw.getShapeText(shape);
        if (text != null) {
          // System.out.println("Found text shape");
          Matcher matcher = pattern.matcher(text);
          String urlStr;
          while (matcher.find()) {
            urlStr = matcher.group();
            System.out.println("Found \"" + urlStr + "\" on slide " + (i+1));
            // System.out.println("  Status: " + getStatus(urlStr) + "\n");
          }
        }
      }
    }

    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()




  public static boolean getStatus(String urlStr)
  /* Pings a URL by sending a HEAD request to it and returning true
     if the response is in the 200-399 range.
     The total timeout is effectively two times the TIMEOUT value.
  */
  {
    // System.out.println("Checking " + urlStr + " ...");

    long startTime = 0;
    boolean isFine = false;
    HttpURLConnection urlConn = null;
    try {
      URL url = new URL(urlStr);
      urlConn = (HttpURLConnection) url.openConnection();
      urlConn.setConnectTimeout(TIMEOUT);
      urlConn.setReadTimeout(TIMEOUT);
      startTime = System.currentTimeMillis();

      urlConn.setRequestMethod("HEAD");
      int responseCode = urlConn.getResponseCode();
      String responseMsg = urlConn.getResponseMessage();

      if (responseMsg == null)
        System.out.println("Response code: " + responseCode);
      else
        System.out.println("Response message: " + responseMsg);

      isFine =  (responseCode == HttpURLConnection.HTTP_OK);
    }
    catch (MalformedURLException e) 
    {  System.out.println("URL is incorrect: " + urlStr); }
    catch (IOException e) 
    {  System.out.println("Could not contact the " + urlStr + " website: " + e);  } 
    finally {
      if (urlConn != null)
        urlConn.disconnect();
    }

    if (isFine)
      System.out.println("Ping time: " + (System.currentTimeMillis() - startTime) + " ms");
    return isFine;
  }  // end of getStatus()
 



}  // end of CheckLinks class
