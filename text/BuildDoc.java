
// BuildDoc.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2015

/* Create a new swriter document, add a few lines, styles,
   images, text frame, bookmark, and save it
   as the file "build" and the specified extension.

   Tested extensions are: doc, docx, rtf, odt, pdf, txt,
   and other extensions are treated like text by default (e.g. "java")
*/

import java.awt.*;
import java.util.*;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.text.*;



public class BuildDoc
{

  public static void main(String[] args)
  {
    XComponentLoader loader = Lo.loadOffice();
    XTextDocument doc = Write.createDoc(loader);
    // XTextDocument doc = Write.createDocFromTemplate("chapter.ott", loader);

    if (doc == null) {
      System.out.println("Writer doc creation failed");
      Lo.closeOffice();
      return;
    }

    GUI.setVisible(doc, true);

    XTextCursor cursor = Write.getCursor(doc);

    Props.showObjProps("Cursor", cursor);

    int pos = Write.append(cursor, "Some examples of simple text ");
    Write.append(cursor, "styles.\n");
    Write.styleLeftBold(cursor, pos);

    pos = Write.getPosition(cursor);
    Write.appendPara(cursor, "This line is written in red italics.");
    Write.styleLeftColor(cursor, pos, Color.RED);
    Write.styleLeftItalic(cursor, pos);

    Write.appendPara(cursor, "Back to old style\n");

    Write.appendPara(cursor, "A Nice Big Heading");
    Write.stylePrevParagraph(cursor, "Heading 1");


    Write.appendPara(cursor, "The following points are important:");
    pos = Write.appendPara(cursor, "Have a good breakfast");
    // System.out.println("pos " + pos);
    Write.appendPara(cursor, "Have a good lunch");
    Write.appendPara(cursor, "Have a good dinner\n");
    Write.styleLeft(cursor, pos, "NumberingStyleName", "Numbering 1");
    // Write.styleLeft(cursor, pos, "NumberingStyleName", "List 1");

    Write.appendPara(cursor, "Breakfast should include:");
    pos = Write.appendPara(cursor, "Porridge");
    // System.out.println("pos " + pos);
    Write.appendPara(cursor, "Orange Juice");
    Write.appendPara(cursor, "A Cup of Tea\n");
    Write.styleLeft(cursor, pos, "NumberingStyleName", "Numbering 1");
    Write.styleLeft(cursor, pos, "ParaIsNumberingRestart", true);
    // Write.styleLeft(cursor, pos, "NumberingStyleName", "List 1");

    Write.append(cursor, "This line ends with a bookmark.");
    Write.addBookmark(cursor, "ad-bookmark");
    Write.appendPara(cursor, "\n");

    Write.appendPara(cursor, "Here's some code:");

    XTextViewCursor tvc = Write.getViewCursor(doc);
    System.out.println("Code starts at coordinate: " + Write.getCoordStr(tvc));
    int yPos = tvc.getPosition().Y;

    pos = Write.getPosition(cursor);
    Write.append(cursor, "\npublic class Hello\n");
    Write.append(cursor, "{\n");
    Write.append(cursor, "  public static void main(String args[]\n");
    Write.append(cursor, "  {  System.out.println(\"Hello Andrew\");  }\n");
    Write.appendPara(cursor, "}  // end of Hello class\n");
    Write.styleLeftCode(cursor, pos);

    Write.appendPara(cursor, "A text frame");
    Write.addTextFrame(cursor, yPos, 
                   "This is a newly created text frame.\nWhich is over on the right of the page, next to the code.", 4000, 1500);

    Write.appendPara(cursor, "A link to my JLOP website:");
    String urlStr = "http://fivedots.coe.psu.ac.th/~ad/jlop/";
    pos = Write.getPosition(cursor);
    Write.append(cursor, urlStr);

    // Write.styleLeft(cursor, pos, "HyperLinkName", "Killer Game Programming in Java");
    // Write.styleLeft(cursor, pos, "HyperLinkTarget", "_blank");
    Write.styleLeft(cursor, pos, "HyperLinkURL", urlStr);
    Write.endParagraph(cursor);

    Write.pageBreak(cursor);

    Write.appendPara(cursor, "Image Example");
    Write.stylePrevParagraph(cursor, "Heading 2");

    String imFnm = "skinner.png";
    Write.appendPara(cursor, "The following image comes from \"" + imFnm + "\":\n");

    Write.append(cursor, "Image as a link: ");
    Write.addImageLink(doc, cursor, imFnm);

    //com.sun.star.awt.Size sz = Images.getSizePixels(imFnm);
    //System.out.println("Image Pixel size: " + sz.Width + " x " + sz.Height);

    com.sun.star.awt.Size imSize = Images.getSize100mm(imFnm);
    System.out.println("Image 100mm size: " + imSize.Width + " x " + imSize.Height);
    int w = (int)Math.round(imSize.Width*1.5);   // enlarge by 1.5x
    int h = (int)Math.round(imSize.Height*1.5);
    Write.addImageLink(doc, cursor, imFnm, w, h);    

    Write.endParagraph(cursor);
    Write.stylePrevParagraph(cursor, "ParaAdjust", com.sun.star.style.ParagraphAdjust.CENTER);
                                             // or LEFT, RIGHT, STRETCH

    Write.append(cursor, "Image as a shape: ");
    Write.addImageShape(doc, cursor, imFnm); 
    Write.endParagraph(cursor);

    // Write.pageBreak(cursor);

    int textWidth = Write.getPageTextWidth(doc);   // in 1/100 mm units
    System.out.println("Text width of page: " + textWidth);
    Write.addLineDivider(cursor, (int)Math.round(textWidth*0.5) );

    Write.appendPara(cursor, "\nTimestamp: " + Lo.getTimeStamp() + "\n");

    Write.append(cursor, "Time (according to office): "); 
    Write.appendDateTime(cursor);
    Write.endParagraph(cursor);

    Info.setDocProps(doc, "Writer Text Example", "Examples", "Andrew Davison");
              // set subject, title, and author props

    Lo.wait(3000);  // wait a bit, so user can see results
    
    // move view cursor to bookmark position
    XTextContent bookmark = Write.findBookmark(doc, "ad-bookmark");
    XTextRange bmRange = bookmark.getAnchor();

    XTextViewCursor viewCursor = Write.getViewCursor(doc);
    viewCursor.gotoRange(bmRange, false);

    Lo.wait(3000);

    Lo.saveDoc(doc, "build.odt");
                   // doc, docx, rtf, odt, pdf, txt
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()



}  // end of BuildDoc class

