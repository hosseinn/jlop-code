
JLOP Chapter 4. Spreadsheet Processing

From the website:

  Java LibreOffice Programming
  http://fivedots.coe.psu.ac.th/~ad/jlop

  Dr. Andrew Davison
  Dept. of Computer Engineering
  Prince of Songkla University
  Hat yai, Songkhla 90112, Thailand
  E-mail: ad@fivedots.coe.psu.ac.th


If you use this code, please mention my name, and include a link
to the website.

Thanks,
  Andrew

============================

This directory contains 26 Java files:
  * BuildTable.java, CalcSort.java, CellRanges.java, CellTexts.java,
    CopyTests.java, DataSort.java, ExtractNums.java, Filler.java,
    FunctionsTest.java, GarlicSecrets.java, GoalSeek.java, LinearSolverTest.java,
    ModifyListener.java, Multiples.java, PivotSheet1.java, PivotSheet2.java,
    PivotTable.java, Ranger.java, ReplaceAll.java, Scenarios.java,
    SelectListener.java, ShowSheet.java, SolverTest.java, SolverTest2.java,
    StylesAllInfo.java, Validator.java


There are 7 batch files:
  * compile.bat
  * run.bat
     - make sure they refer to the correct locations for your
       Java, JNA, and my Utils classes.
       For details, read:
         "Installing the code for "Java LibreOffice Programming"
          http://fivedots.coe.psu.ac.th/~ad/jlop/install.html

    compile.bat and run.bat file both call:

  * lofind.bat
     - this tries to find LibreOffice (v4 or v5) on your machine,
       and checks its "bitness" with that of Java. They should
       both either be 32-bit or 64-bit. A warning is printed if 
       they are not.

  * loKill.bat        // kills the Office process
  * loList.bat        // lists the Office process (if it is running)

  * loDoc.bat         // searches the online LibreOffice documentation
  * loGuide.bat       // searches the online OpenOffice Dev Guide


Assorted test files for the Java code:
  * pivottable1.ods, pivottable2.ods
    produceSales.xlsx, results.xlsx, skinner.png
    "small totals.ods", sorted.csv
    totals.ods, unsorted.txt


----------------------------
Compilation

All the Java files will compile with the compile.bat batch file:

> compile *.java
    // you must have LibreOffice, Java, JNA, and my Utils classes installed


===================================
Execution:

// You must have LibreOffice, Java, JNA, and my Utils classes installed

The following programs do not require command line arguments:
  * BuildTable.java, CellTexts.java, DataSort.java,
    Filler.java, FunctionsTest.java, GarlicSecrets.java,
    GoalSeek.java, LinearSolverTest.java, ModifyListener.java,
    PivotSheet1.java, PivotSheet2.java, PivotTable.java,
    SelectListener.java, SolverTest.java, SolverTest2.java,
    Validator.java

e.g. call them with run:
> run GarlicSecrets

---------------------------
The following programs require command line arguments:
 * CalcSort.java, ExtractNums.java,
   ReplaceAll.java, ShowSheet.java,
   StylesAllInfo.java

e.g.
> run CalcSort unsorted.txt

> run ExtractNums "small totals.ods"

> run ReplaceAll dog

> run ShowSheet totals.ods show.pdf

> run StylesAllInfo totals.ods


---------------------------
The following programs are not explained in the chapter. 
None of them require command line arguments:
  * CellRanges.java, CopyTests.java, Multiples.java,
    Ranger.java, Scenarios.java

e.g. call them with run:
> run Scenarios

=================================
Last updated: 2nd October 2015
