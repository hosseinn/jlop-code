
// FunctionsTest.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, September 2015

/*  Use Calc functions without creating a spreadsheet.

    Lists of functions divided into categories:
      https://help.libreoffice.org/Calc/Functions_by_Category

      https://wiki.openoffice.org/wiki/Documentation/
                How_Tos/Calc:_Functions_listed_by_category

    LO Calc function categories:
      1 Database
        - extract info from Calc data tables, where data is organised in rows

      2 Date & Time  e.g. EASTERSUNDAY

      3 Financial
        - for common business calculations
      
      4 Information
        - info about cells, such as whether they contain text or a formula

      5 Logical  - boolean logic

      6 Mathematical
        - include trigonometric, hyperbolic, logarithmic and summation functions
          e.g. ROUND, SIN, RADIANS

      7 Array
        - operate on and return entire tables of data; e.g. TRANSPOSE

      8 Statistical
        - statistical and probability calculations; e.g. ZTEST

      9 Spreadsheet
        - find values in tables, or cell references

      10 Text

      11 Add-in
         - https://help.libreoffice.org/Calc/Add-in_Functions
           - more date ops, rot13

         - https://help.libreoffice.org/Calc/
                 Add-in_Functions,_List_of_Analysis_Functions_Part_One
           - base conversion, more stats, 

         - https://help.libreoffice.org/Calc/
                  Add-in_Functions,_List_of_Analysis_Functions_Part_Two
           - base conversion, complex arithmetic,
*/


import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;

import com.sun.star.sheet.*;
import com.sun.star.table.*;


public class FunctionsTest
{

  public static void main(String[] argus)
  {
    XComponentLoader loader = Lo.loadOffice();

    // round
    System.out.println("ROUND result fot 1.999 is: " + 
                                 Calc.callFun("ROUND", 1.999) + "\n");

    // sine of 30 degrees
    System.out.printf("SIN result fot 30 degrees is: %.3f\n\n", 
                           Calc.callFun("SIN",  Calc.callFun("RADIANS", 30)));
    
    // average function
    Object[] args = new Object[]{1, 2, 3, 4, 5};
    double avg = (Double) Calc.callFun("AVERAGE", args);
    System.out.println("Average of the numbers is: " +avg + "\n");


    // slope function
    double[][] xData = {{1.0, 2.0, 3.0}};   // must be a matrix
    double[][] yData = {{3.0, 6.0, 9.0}};   // must be a matrix
    args = new Object[]{yData, xData};
    double slope = (Double) Calc.callFun("SLOPE", args);
    System.out.println("SLOPE of the line: " + slope + "\n");


    // zTest function
    double[][] data = {{1.0, 2.0, 3.0}};
    args = new Object[]{ data, 2.0};
    double res = (Double) Calc.callFun("ZTEST", args);
    System.out.println("ZTEST result for data {{1,2,3}} and 2 is: " + res + "\n");

    // easter sunday function
    double easterSun = (Double) Calc.callFun("EASTERSUNDAY", 2015);
    int day = (int) Math.round((Double) Calc.callFun("DAY", easterSun));
    int month = (int) Math.round((Double) Calc.callFun("MONTH", easterSun));
    int year = (int) Math.round((Double) Calc.callFun("YEAR", easterSun));
    System.out.println("Easter Sunday (d/m/y): " + day + "/" + 
                                           month + "/" + year + "\n");


    // transpose a matrix
    double[][] arr = {{1, 2, 3},{4, 5, 6}};
    args = new Object[]{arr};
    Object[][] transMat = (Object[][]) Calc.callFun("TRANSPOSE", args);
    // Object[][] transMat = (Object[][]) Calc.callFun("TRANSPOSE", (Object) arr);
    Calc.printArray(transMat);


    // sum two imaginary numbers: "13+4j" + "5+3j" returns 18+7j.
    String[][] nums = {{"13+4j"},{"5+3j"}};
    args = new Object[]{nums};
    String sum = (String) Calc.callFun("IMSUM", args);
    // String sum = (String) Calc.callFun("IMSUM", (Object)nums);
    System.out.println("13+4j + 5+3j: "+ sum + "\n");


    // decimal to hex
    // DEC2HEX: DEC2HEX(100;4) returns "0064"
    args = new Object[]{100, 4};
    String hex4 = (String) Calc.callFun("DEC2HEX", args);
    System.out.println("100 to hex: "+ hex4 + "\n");

    // ROT13(Text)
    String rot13 = (String) Calc.callFun("ROT13", "hello");
    System.out.println("ROT13 of \"hello\": "+ rot13 + "\n");

    // Roman numbers
    // http://cs.stackexchange.com/questions/7777/is-the-language-of-roman-numerals-ambiguous
    String roman = (String) Calc.callFun("ROMAN", 999);
    args = new Object[]{999, 4};
    String roman4 = (String) Calc.callFun("ROMAN", args);
    System.out.println("999 in Roman numerals: "+ roman + " or " + roman4 + "\n");


    // using ADDRESS
    args = new Object[]{2, 5, 4};   // row, column, abs
    String cellName = (String) Calc.callFun("ADDRESS", args);
    System.out.println("Relative address for (5,2): " + cellName + "\n");


    //System.out.println("Function names");
    //Lo.printNames( Calc.getFunctionNames(), 6);

    Calc.printFunctionInfo("EASTERSUNDAY");
    Calc.printFunctionInfo("ROMAN");

    showRecentFunctions();

    Lo.closeOffice();
  }  // end of main()



  private static void showRecentFunctions()
  {
    int[] recentIDs = Calc.getRecentFunctions();
    if (recentIDs == null)
      return;

    System.out.println("Recently used functions (" + recentIDs.length + "): ");
    for (int i = 0; i < recentIDs.length; i++) {
      PropertyValue[] props = Calc.findFunction(recentIDs[i]);
      System.out.println("  " + Props.getValue("Name", props));
    }
    System.out.println("\n");
  }  // end of showRecentFunctions()


}  // end of FunctionsTest class
