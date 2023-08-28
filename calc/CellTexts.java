
// CellTexts.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/*  Based on code in Dev Guide's SpreadSheetSample.java example.

    Add paragraphs, hyperlink to a cell.
    Change the style of the cells.
    Access the text.
    Add an annotation.
*/

import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;
import com.sun.star.text.*;



public class CellTexts
{

  public static void main(String args[])
  {
    XComponentLoader loader = Lo.loadOffice();
    XSpreadsheetDocument doc = Calc.createDoc(loader);
    if (doc == null) {
      System.out.println("Document creation failed");
      Lo.closeOffice();
      return;
    }
    GUI.setVisible(doc, true);

    XSpreadsheet sheet = Calc.getSheet(doc, 0);

    Calc.highlightRange(sheet, "A2:C7", "Cells and Cell Ranges");

    XCell xCell = Calc.getCell(sheet, "B4");

    // Insert two text paragraphs and a hyperlink into the cell
    XText xText = Lo.qi(XText.class, xCell);
    XTextCursor cursor = xText.createTextCursor();
    Write.appendPara(cursor, "Text in first line.");
    Write.append(cursor, "And a ");
    Write.addHyperlink(cursor, "hyperlink",
                    "http://fivedots.coe.psu.ac.th/~ad/jlop/");

    // beautify the cell
    // properties from styles.CharacterProperties
    Props.setProperty(xCell, "CharColor", Calc.DARK_BLUE);
    Props.setProperty(xCell, "CharHeight", 20.0);

    // property from styles.ParagraphProperties
    Props.setProperty(xCell, "ParaLeftMargin", 500);

    printCellText(xCell);

    Calc.addAnnotation(sheet, "B4", "This annotation is located at B4");

    Lo.saveDoc(doc, "linkedText.ods");
    Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()



  private static void printCellText(XCell xCell)
  {
    XText xText = Lo.qi(XText.class, xCell);
    System.out.println("Cell Text: \"" + xText.getString() + "\"");

    XTextCursor cursor = xText.createTextCursor();
    if (cursor == null)
      System.out.println("Text cursor is null");

    // cannot get sentence or paragraph cursors for some reason ??
    XSentenceCursor sentCursor = Lo.qi(XSentenceCursor.class, cursor);
    if (sentCursor == null)
      System.out.println("Sentence cursor is null");

    XParagraphCursor paraCursor = Lo.qi(XParagraphCursor.class, cursor);
    if (paraCursor == null)
      System.out.println("Paragraph cursor is null");
    else {
      paraCursor.gotoStart(false);     // go to start of text; no selection
      do {
        paraCursor.gotoEndOfParagraph(true);   // select all of paragraph
        System.out.println( paraCursor.getString() );
      } while (paraCursor.gotoNextParagraph(false));
    }
  }  // end of printCellText()



}  // end of CellTexts class
