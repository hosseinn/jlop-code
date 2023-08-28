
// GoalSeek.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/*  Four goal seek examples.
 
    First example is a programming version of the one in the Calc user
    guide, ch 9, Data Analysis, p.265
*/

import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;



public class GoalSeek
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
    //GUI.setVisible(doc, true);

    XSpreadsheet sheet = Calc.getSheet(doc, 0);
    XGoalSeek gs = Lo.qi(XGoalSeek.class, doc);


    Calc.setVal(sheet, "C1", 9);    // x-variable and starting value 
    Calc.setVal(sheet, "C2", "=SQRT(C1)");   // formula 
    double x = Calc.goalSeek(gs, sheet, "C1", "C2", 4);
    System.out.println("x == " + x + " when sqrt(x) == 4\n");   // x is 16


    // ==============================================

    x = Calc.goalSeek(gs, sheet, "C1", "C2", -4);
    /* The formula is still y = sqrt(x)
       Find x when sqrt(x) == -4, which is impossible. */
    System.out.println("x == " + x + " when sqrt(x) == -4\n");
       /* fails (i.e. divergence is big) */

    // ==============================================

    Calc.setVal(sheet, "D1", 0.8);    // x-variable and starting value
    Calc.setVal(sheet, "D2", "=(D1^2 - 1)/(D1 - 1)");   // formula 
    /* The formula is y = (x^2 -1)/(x-1)
       After factoring, this is just y = x+1
    */
    x = Calc.goalSeek(gs, sheet, "D1", "D2", 2);
    System.out.println("x == " + x + " when x+1 == 2\n");
                                      // x is 1.0000000000000053

    // ==============================================


    Calc.setVal(sheet, "B1", 100000);   // x-variable; starting capital
    Calc.setVal(sheet, "B2", 1);       // n, no. of years
    Calc.setVal(sheet, "B3", 0.075);   // i, interest rate (7.5%)

    Calc.setVal(sheet, "B4", "=B1*B2*B3");   // formula 
    /* The formula is Annual interest = x*n*r
       where capital (x), number of years (n), and interest rate (r).
         Find the capital, if the other values are given.
    */
    x= Calc.goalSeek(gs, sheet, "B1", "B4", 15000);
    System.out.println("x == " + x + " when x*" +
              Calc.getVal(sheet, "B2") + "*" + Calc.getVal(sheet, "B3") +
                       "  == 15000\n");  // x is 200,000


    // ==============================================

    Calc.setVal(sheet, "E1", 0);    // x-variable and starting value
    Calc.setVal(sheet, "E2", "=(E1^3 - 2*E1 + 2");   // formula 
    x = Calc.goalSeek(gs, sheet, "E1", "E2", 0);
    System.out.println("x == " + x + " when formula == 0\n");
       // x is -1.7692923428381226
       // so not using Newton's method which oscillates between 0 and 1


    // Lo.waitEnter();
    Lo.closeDoc(doc);
    Lo.closeOffice();
  }  // end of main()


}  // end of GoalSeek class
