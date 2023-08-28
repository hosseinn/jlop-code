
JLOP Chapter 12. Coding UNO Components

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

Files used in the RandomSents example:

 * RandomSents.idl
      -- the UNO IDL definitions for RandomSents

 * RandomSentsImpl - complete.java
      -- a completed implementation of RandomSentsImpl.java

 * ManifestRandomSents.txt
      -- the manifest for the RandomSents JAR file

 * ViewRegistry.java
      -- an example of how to use the com.sun.star.registry module

 * PoemCreator.java
      -- and example that uses Office *and* the RandomSents UNO component


An empty package directory (org.openoffice)
 * org/openoffice


A RandomSents/ folder used to build the RandomSents extension (OXT file).
It contains:
  * META-INF/manifest.xml
       -- the names of the RDB and JAR files used in the extension
  * description.xml
       -- information displayed by the Extension Manager when listing
          the installed extension

  * license.txt   -- extension's MIT license

  * package-description.txt

  * stack.png     -- extension's icon


----------------------------
There are 12 batch files for the UNO component 2-part toolchain:

Part 1:
-------
  * idlc.bat, regmerge.bat, javamaker.bat, skelComp.bat
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

See below for how to call these batch files with RandomSents.

----------------------------
Additional tools used in the toolchain:

  * cfr_0_118.jar
      -- from http://www.benf.org/other/cfr/
      -- Java class decompiler

  * elevate-1.3.0-redist.7z
      -- from http://code.kliu.org/misc/elevate/
      -- unzip the 32/64- bit version of elevate.exe for your Windows OS
      -- runs a script with admin privileges

  * tr.exe, zip.exe
      -- from https://github.com/bmatzelle/gow/wiki
      -- Windows versions of the UNIX tr and zip tools


----------------------------
There are 9 Office-related batch files:

  * compileExt.bat
  * runExt.bat
     -- these are used to compile and run programs that use Office 
        *and* the installed UNO component
    -- they use compile.bat and run.bat features, and FindExtJar.java

    - see below for an example

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
Building the RandomSents Component


Before you begin you must have LibreOffice, Java, JNA, and my Utils 
classes installed.

Follow the steps shown in the toolchain.png file, and
described in the chapter. The batch files also contain
documentation.

Most of the batch files take a single argument, the service name
of the component (e.g. RandomSents).


Part 1 of Toolchain
-------------------
> idlc RandomSents
       -- creates RandomSents.urd from RandomSents.idl
---
You may need to elevate the call to idlc.bat:
> elevate.exe -k -w idlc RandomSents
---

> regmerge RandomSents
       -- creates RandomSents.rdb from RandomSents.urd

> regview RandomSents
       -- OPTIONAL: display contents of RandomSents.rdb
---
Instead of regview you can use my ViewRegistry.java:
> compile ViewRegistry.java
> run ViewRegistry RandomSents.rdb
---

> javamaker RandomSents
       -- adds randomsents package to org.openoffice
       -- stores XRandomSents.class and XRandomSents.java in
          randomsents package

> skelComp RandomSents
      -- creates skeleton implementation of RandomSents in
         RandomSentsImpl.java
---
You may need to elevate the call to skelComp.bat:
> elevate.exe -k -w skelComp RandomSents
---

---
The calls to idlc.bat, regmerge.bat, javamaker.bat, skelComp.bat
are combined in genCode.bat:

> genCode RandomSents
      -- only use this when you're sure the separate steps are ok
---

Now you must implement the component-specific methods in RandomSentsImpl.java:
  * getisAllCaps()
  * setisAllCaps()
  * getParagraph()
  * getSentences()

I've already done this, and saved the code in "RandomSentsImpl - complete.java".
**COPY** the  file, replacing the skeleton RandomSentsImpl.java with
the complete version.


Part 2 of Toolchain
-------------------

> compileOrg RandomSents RandomSentsImpl.java
    -- creates RandomSentsImpl.class by compiling
       RandomSentsImpl.java using Office and the 
       org.openoffice.randomsents package

> toJar RandomSents
    -- creates RandomSentsImpl.jar
    -- uses ManifestRandomSents.txt for the manifest information

> makeOXT RandomSents
    -- creates RandomSents.oxt (the Office extension)
    -- uses the previously created RandomSents.rdb, RandomSents.jar,
       and the information in the RandomSents/ directory

> pkg RandomSents
    -- installs RandomSents.oxt (the Office extension) into Office
    -- the license is printed, which the user must accept or reject
    -- the Extension Manager is started, so the RandomSents extension
       can be seen

---
The calls to compileOrg.bat, toJar.bat, makeOxt.bat, pkg.bat
are combined in addToOffice.bat:

> addToOffice RandomSents RandomSentsImpl.java
      -- only use this when you're sure the separate steps are ok
---

> extManager
    -- OPTIONAL: this starts the Extension Manager


=================================
Building the RandomSents Component

PoemCreator.java uses the Office API *and* the installed RandomSents
extension. It finds the extension using FindExtJar.java, which is called
from the compileExt.bat and runExt.bat scripts.

> compile FindExtJar.java
    -- compile before using the batch scripts

> compileExt RandomSents PoemCreator.java
    -- creates PoemCreator.class

>  runExt RandomSents PoemCreator
    -- uses the RandomSents component to create poetry, which is
       saved to "poem.doc" 

WELL DONE!

=================================
Last updated: 26th September 2016
