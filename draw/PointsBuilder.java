
// PointsBuilder.java
// Andrew Davison, July 2015, ad@fivedots.coe.psu.ac.th

/* Convert a text file of points into a series of slides.
   Uses a template from Office.

  Usage:
    run PointsBuilder pointsInfo.txt
*/

import java.io.*;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.text.*;
import com.sun.star.drawing.*;
import com.sun.star.container.*;



public class PointsBuilder
{

  public static void main (String args[])
  {
    if (args.length != 1) {
      System.out.println("Usage: PointsBuilder <points fnm>");
      return;
    }

    XComponentLoader loader = Lo.loadOffice();

    reportTemplates();

    // create an impress document using a template
    String templateFnm = Draw.getSlideTemplatePath() + "Inspiration.otp";
    XComponent doc = Lo.createDocFromTemplate(templateFnm, loader);
    // XComponent doc = Lo.createImpressDoc(loader);
    if (doc == null) {
      System.out.println("Impress doc creation failed");
      Lo.closeOffice();
      return;
    }

    readPoints(args[0], doc);

    System.out.println("Total no. of slides: " + Draw.getSlidesCount(doc));
    Lo.saveDoc(doc, "points.odp"); 
        // odp, ppt, pptx
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()




  private static void reportTemplates()
  {
    String[] templateDirs = Info.getDirs("Template");
    System.out.println("Templates dirs:");
    for(String dir : templateDirs)
      System.out.println("  " + dir);

    String templateDir = Draw.getSlideTemplatePath();
    System.out.println("\nTemplates files in \"" + templateDir + "\"");
    String[] templatesFnms = FileIO.getFileNames(templateDir);
    for(String fnm : templatesFnms)
      System.out.println("  " + fnm);
  }  // end of reportTemplates()



  private static void readPoints(String fnm, XComponent doc)
  /* Read in a text file of points which are converted to slides.
     Formatting rules:
       * ">", ">>", etc are points and their levels
       * blank lines and lines starting with "//" are ignored
       * any other lines are the title text of a new slide
  */
  { 
    XDrawPage currSlide  = Draw.getSlide(doc, 0);
    Draw.titleSlide(currSlide, "Java-Generated Slides", "Using LibreOffice");
    XText body = null;
    try {
      BufferedReader br = new BufferedReader( new FileReader(fnm));
      String line;
      char ch;
      while((line = br.readLine()) != null) {
        if (line.length() == 0)  // blank line
          continue;
        if (line.startsWith("//"))   // comment
          continue;
        ch = line.charAt(0);
        if (ch == '>')  // a bullet with some indentation
          processBullet(line, body);
        else {  // a title for a new slide
          currSlide = Draw.addSlide(doc);
          body = Draw.bulletsSlide(currSlide, line);
        }
      }
      br.close();
      System.out.println("Read in points file: " + fnm);
    } 
    catch (IOException e) 
    { System.out.println("Error reading points file: " + fnm); }
  }  // end of readPoints()



  private static void processBullet(String line, XText body)
  // count the number of '>'s to determine the bullet level
  {
    if (body == null)
      System.out.println("No slide body for " + line);
    else {
      int pos = 1;   // there is at least 1 '>'
      char ch = line.charAt(pos);
      while (ch == '>') {
        pos++;
        ch = line.charAt(pos);
      }
      line = line.substring(pos).trim();
      Draw.addBullet(body, pos-1, line);
    }
  } // end of processBullet()


}  // end of PointsBuilder class
