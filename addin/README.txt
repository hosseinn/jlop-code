
JLOP Chapter 14. Calc Add-ins

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

Files used in the Doubler example:

 * Doubler.idl
      -- the UNO IDL definitions for Doubler

 * DoublerImpl - complete.java
      -- a completed implementation of DoublerImpl.java

 * ManifestDoubler.txt
      -- the manifest for the Doubler JAR file

 * calcTest.ods
      -- a Calc spreadsheet that uses the Doubler add-ins
         after they've been installed into Office


An empty package directory (org.openoffice)
 * org/openoffice


A Doubler/ folder used to build the Doubler extension (OXT file).
It contains:
  * META-INF/manifest.xml
       -- the names of the RDB and JAR files used in the extension

  * CalcAddin.xcu
       -- the XML description of the add-in functions used by
          Calc's Function Wizard

  * description.xml
       -- information displayed by the Extension Manager when listing
          the installed extension

  * double.png    -- extension's icon

  * license.txt   -- extension's MIT license

  * package-description.txt

  * Utils.jar
        -- a JAR file holding my utilities 


----------------------------
There are 12 batch files for the UNO component 2-part toolchain:

Part 1:
-------
  * idlc.bat, regmerge.bat, javamaker.bat, skelAddin.bat
  * genCode.bat   // this calls the above batch files in order
  * regview.bat

Part 2:
-------
  * compileOrg.bat, toJar.bat, makeOxt.bat, pkg.bat
  * addToOffice.bat    // this calls the above batch files in order
  * extManager.bat 

A picture of how these batch files link together to form the
toolchain is in:
 * toolchain.png

See below for how to call these batch files with Doubler.

----------------------------
Additional tools used in the toolchain:

  * cfr_0_118.jar
      -- from http://www.benf.org/other/cfr/
      -- Java class decompiler

  * elevate-1.3.0-redist.7z
      -- from http://code.kliu.org/misc/elevate/
      -- unzip the 32/64- bit version of elevate.exe for your Windows OS
      -- runs a script with admin privileges

  * tail.exe, tr.exe, zip.exe
      -- from https://github.com/bmatzelle/gow/wiki
      -- Windows versions of the UNIX tail, tr and zip tools


----------------------------
There are 7 Office-related batch files:

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


=================================
Building the Doubler Component


Before you begin you must have LibreOffice, Java, JNA, and my Utils 
classes installed.

Follow the steps shown in the toolchain.png file, and
described in the chapter. The batch files also contain
documentation.

Most of the batch files take a single argument, the service name
of the component (e.g. Doubler).


Part 1 of Toolchain
-------------------
> idlc Doubler
       -- creates Doubler.urd from Doubler.idl
---
You may need to elevate the call to idlc.bat:
> elevate.exe -k -w idlc Doubler
---

> regmerge Doubler
       -- creates Doubler.rdb from Doubler.urd

> regview Doubler
       -- OPTIONAL: display contents of Doubler.rdb

> javamaker Doubler
       -- adds doubler package to org.openoffice
       -- stores XDoubler.class and XDoubler.java in
          doubler package

> skelAddin Doubler
      -- creates skeleton implementation of Doubler in
         DoublerImpl.java
---
You may need to elevate the call to skelAddin.bat:
> elevate.exe -k -w skelAddin Doubler
---

---
The calls to idlc.bat, regmerge.bat, javamaker.bat, skelAddin.bat
are combined in genCode.bat:

> genCode Doubler
      -- only use this when you're sure the separate steps are ok
---

Now you must implement the add-in methods in DoublerImpl.java:
  *   double doubler(double value)
  *   double doublerSum(double[][] vals)
  *   double[][] sortByFirstCol(double[][] vals)


I've already done this, and saved the code in "DoublerImpl - complete.java".
**COPY** the  file, replacing the skeleton DoublerImpl.java with
the complete version.


Part 2 of Toolchain
-------------------

> compileOrg Doubler DoublerImpl.java
    -- creates DoublerImpl.class by compiling
       DoublerImpl.java using Office and the 
       org.openoffice.doubler package

> toJar Doubler
    -- creates DoublerImpl.jar
    -- uses ManifestDoubler.txt for the manifest information

> makeOXT Doubler
    -- creates Doubler.oxt (the Office extension)
    -- uses the previously created Doubler.rdb, Doubler.jar,
       and the information in the Doubler/ directory

> pkg Doubler
    -- installs Doubler.oxt (the Office extension) into Office
    -- the license is printed, which the user must accept or reject
    -- the Extension Manager is started, so the Doubler extension
       can be seen

---
The calls to compileOrg.bat, toJar.bat, makeOxt.bat, pkg.bat
are combined in addToOffice.bat:

> addToOffice Doubler
      -- only use this when you're sure the separate steps are ok
---

> extManager
    -- OPTIONAL: this starts the Extension Manager


=================================
Using the Doubler Add-in

CakcTest.ods uses the three Doubler add-in functions, and 
should look like Figure 2 in Chapter 14.


=================================
Last updated: 4th November 2016
