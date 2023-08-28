
// MathQuestions.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, June 2015

/* Create a new swriter documents, add random formulae involving "x",
   a fraction using x, or sqrt of x,
   and save it as "mathQuestions.pdf".
*/

import java.util.*;

import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.text.*;


public class MathQuestions
{

  public static void main(String[] args)
  {

    XComponentLoader loader = Lo.loadOffice();
    XTextDocument doc = Write.createDoc(loader);
    if (doc == null) {
      System.out.println("Writer doc creation failed");
      Lo.closeOffice();
      return;
    }

    XTextCursor cursor = Write.getCursor(doc);

    Write.appendPara(cursor, "Math Questions");
    Write.stylePrevParagraph(cursor, "Heading 1");


    Write.appendPara(cursor, "Solve the following formulae for x:\n");

    Random r = new Random();
    String formula;
    for(int i=0; i < 10; i++) {    // generate 10 formulae
      int iA = r.nextInt(8)+2;  // generate some integers
      int iB = r.nextInt(8)+2;
      int iC = r.nextInt(9)+1;
      int iD = r.nextInt(8)+2;
      int iE = r.nextInt(9)+1;
      int iF1 = r.nextInt(8)+2;

      int choice = r.nextInt(3);    // decide between 3 kinds of formulae
      if (choice == 0)
        formula = "{{{sqrt{" + iA + "x}} over " + iB + "} + {" + iC + " over " + iD +
                     "}={" + iE + " over " + iF1 + "}}";
      else if (choice == 1)
        formula = "{{{" + iA + "x} over " + iB + "} + {" + iC + " over " + iD +
                    "}={" + iE + " over " + iF1 + "}}";
      else
        formula = "{" + iA + "x + " + iB + " =" + iC + "}";

      Write.addFormula(cursor, formula);
      Write.endParagraph(cursor);
    }

    Write.append(cursor, "Timestamp: " + Lo.getTimeStamp()  + "\n");

    Lo.saveDoc(doc, "mathQuestions.pdf");
    Lo.closeDoc(doc);
    Lo.closeOffice();
  } // end of main()



}  // end of MathQuestions class

