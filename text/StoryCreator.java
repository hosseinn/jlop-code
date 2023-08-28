
// StoryCreator.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2015

/* Read in a text file, apply a new paragraph style, header, page
   numbers in footer, A4 page style, title, and subtitle, and
   save as "bigStory.doc"

   Usage:
     run StoryCreator scandal.txt

*/

import java.io.*;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;

import com.sun.star.text.*;
import com.sun.star.style.*;
import com.sun.star.container.*;
import com.sun.star.beans.*;


public class StoryCreator
{
  public static void main(String args[])
  {
    if (args.length != 1) {
      System.out.println("Usage: java TextReader <fnm>");
      return;
    }


    XComponentLoader loader = Lo.loadOffice();
    XTextDocument doc = Write.createDoc(loader);
    if (doc == null) {
      System.out.println("Writer doc creation failed");
      Lo.closeOffice();
      return;
    }


    String[] styles = Info.getStyleNames(doc, "ParagraphStyles");
    System.out.println("Paragraph Styles");
    Lo.printNames(styles);


    if (!createParaStyle(doc, "adParagraph")){
      System.out.println("Could not create new paragraph style");
      Lo.closeOffice();
      return;
    }

    XTextRange xTextRange = doc.getText().getStart();
    Props.setProperty(xTextRange, "ParaStyleName", "adParagraph");

    Write.setHeader(doc, "From: " + args[0]);
    Write.setA4PageFormat(doc);
    Write.setPageNumbers(doc);

    XTextCursor cursor = Write.getCursor(doc);

    readText(args[0], cursor);
    Write.endParagraph(cursor);

    Write.appendPara(cursor, "Timestamp: " + Lo.getTimeStamp());

    Lo.saveDoc(doc, "bigStory.doc");
                   // doc, docx, rtf, odt, pdf, txt
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()




  public static void readText(String fnm, XTextCursor cursor)
  /* Write the text in fnm into a Office document.
     Use standard paragraph styling unless the line starts
     with "Title: ", "Author: ", or "Part ".
  */
  {
    StringBuilder sb = new StringBuilder(0);
    BufferedReader br = null;
    try {
      br = new BufferedReader(new FileReader(fnm));
      System.out.println("Reading text from: " + fnm);

      String line;
      while ((line = br.readLine()) != null) {
        // System.out.println("<" + line + ">");
        if (line.length() == 0) {
          if (sb.length() > 0)
            Write.appendPara(cursor, sb.toString());
          sb.setLength(0);
        }
        else if (line.startsWith("Title: ")) {
          Write.appendPara(cursor, line.substring(7));
          Write.stylePrevParagraph(cursor, "Title");
        }
        else if (line.startsWith("Author: ")) {
          Write.appendPara(cursor, line.substring(8));
          Write.stylePrevParagraph(cursor, "Subtitle");
        }
        else if (line.startsWith("Part ")) {
          Write.appendPara(cursor, line);
          Write.stylePrevParagraph(cursor, "Heading");
        }
        else {
          sb.append(line + " ");
        }
      }
      if (sb.length() > 0)
        Write.appendPara(cursor, sb.toString());
    }
    catch (FileNotFoundException ex)
    {  System.out.println("Could not open: " + fnm); }
    catch (IOException ex) 
    {  System.out.println("Read error: " + ex); }
    finally {
      try {
        if (br != null)
          br.close();
      }
      catch (IOException ex) 
      {  System.out.println("Problem closing " + fnm); }
    }
  }  // end of readText()





  public static boolean createParaStyle(XTextDocument textDoc, String styleName)
  // create a new paragraph container/style called styleName
  {
    XNameContainer paraStyles = Info.getStyleContainer(textDoc, "ParagraphStyles");
    if (paraStyles == null)
      return false;

    try {
      // create new paragraph style properties set
      XStyle paraStyle = Lo.createInstanceMSF(XStyle.class, 
                                                 "com.sun.star.style.ParagraphStyle");
      XPropertySet props =  UnoRuntime.queryInterface(XPropertySet.class, paraStyle);

      // set some properties
      props.setPropertyValue("CharFontName", "Times New Roman");
      props.setPropertyValue("CharHeight", 12.0f);
      props.setPropertyValue("ParaBottomMargin", 400);	  //  4mm, in 100th mm

      // set paragraph line spacing to 6mm 
		  LineSpacing lineSpacing = new LineSpacing();
		  lineSpacing.Mode = LineSpacingMode.FIX;
		  lineSpacing.Height = 600;
      props.setPropertyValue("ParaLineSpacing", lineSpacing);

      // props.setPropertyValue("CharWeight", com.sun.star.awt.FontWeight.BOLD);
      // props.setPropertyValue("CharAutoKerning", true);
      // props.setPropertyValue("ParaAdjust", ParagraphAdjust.CENTER_value);
      // props.setPropertyValue("ParaFirstLineIndent", 0);
      // props.setPropertyValue("BreakType", BreakType.PAGE_AFTER);

      // store those properties in a container called styleName
      paraStyles.insertByName(styleName, props);
      return true;
    }
    catch(com.sun.star.uno.Exception e) 
    { System.out.println("Could not set paragraph style");  
      return false;
    }
  }  // end of createParaStyle()


}  // end of StoryCreator class

