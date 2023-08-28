
JLOP Chapter 6. Databases

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

This directory contains 17 Java files:
  * AccessQuery.java, CalcQuery.java, CSVQuery.java,
    DataSourcer.java, DBCmdsCreate.java, DBCmdsQuery.java,
    DBCreate.java, DBQuery.java, DBRels.java,
    EmbeddedQuery.java, FancyRS.java, PreparedSales.java,
    SaveToCSV.java, SimpleQuery.java, StoreInCalc.java,
    ThunderbirdQuery.java, UseRowSets.java


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
  * Books.accdb, Books.mdb, Books.odb,
     liangTables.odb, liangTables.txt, liangTests.txt,
     NewBooks.odb, sales.odb, totals.ods, us-500.csv


One folder containing a JDBC program that uses the
Jaybird JDBC and Firebird:
 * Firebird_Tests/


----------------------------
Compilation

All the Java files will compile with the compile.bat batch file:

> compile *.java
    // you must have LibreOffice, Java, JNA, and my Utils classes installed


===================================
Execution:

// You must have LibreOffice, Java, JNA, and my Utils classes installed

The following programs do not require command line arguments:
  * AccessQuery.java, CalcQuery.java, CSVQuery.java,
    DataSourcer.java, DBCreate.java, FancyRS.java,
    PreparedSales.java, SimpleQuery.java, StoreInCalc.java,
    ThunderbirdQuery.java, UseRowSets.java


e.g. call them with run:
> run AccessQuery


---------------------------
The following programs require command line arguments:
 * DBCmdsCreate.java, DBCmdsQuery.java, DBQuery.java,
   DBRels.java, EmbeddedQuery.java, SaveToCSV.java


e.g.
> run DBCmdsCreate liangTables.txt
      // program doesn't exit properly because of open tables
      // in Base's Table view

> run DBCmdsQuery liangTables.odb liangTests.txt

> run DBQuery liangTables.odb
     // close all the JFrames and then the program can exit

> run DBRels NewBooks.odb

> run EmbeddedQuery sales.odb
      // if the embedded database is Firebird then Firebird and
      // Jaybird must be installed -- see readme.txt in 
      // the Firebird_Tests/ folder for some details on installing
      // them

> run SaveToCSV liangTables.odb


// There are more examples in the comments sections of
   the Java files.

=================================
Last updated: 13th April 2016
