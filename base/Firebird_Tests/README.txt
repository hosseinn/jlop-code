
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

This directory contains 1 Java file:
  * SimpleJDBC.java

There are 2 batch files:
  * compile.bat
  * run.bat
     - make sure they refer to the correct location for Jaybird

One Firebird test file for the Java code:
  * CROSSRATE.FDB

===================================
Compilation and Execution

> compile SimpleJDBC.java
> run SimpleJDBC
    // you must have Firebird and Jaybird installed


===================================
Firebird and Jaybird Installation Steps


Download the embedded version of Firebird from 
http://www.firebirdsql.org/en/downloads/, making sure to grab either the 
32-bit or 64-bit version for your OS. 

This should be unzipped to a convenient location (e.g. d:\firebird), and 
the path added to Window's PATH environment variable. 

Bug Note: When I added Firebird 2.5.5 to PATH this caused a problem 
whenever I subsequently opened embedded Firebird databases in Base! 
Office issued a Runtime error r6034, related to the Microsoft Visual C++ 
runtime library. This is triggered by the presence of a mscvcr80.dll 
file in the downloaded firebird/ directory which duplicates one in 
Windows. 

One solution is to move Microsoft.VC80.CRT.manifest, msvcp80.dll, and 
msvcr80.dll (i.e. three files) from the firebird/ directory to some 
other location (e.g. into an UNUSED/ directory in firebird/). 

The JDBC driver, called Jaybird, is a separate download from 
http://www.firebirdsql.org/en/jdbc-driver/. This should be unzipped to 
d:\jaybird. 


=================================
Last updated: 13th April 2016
