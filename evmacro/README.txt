
JLOP Chapter 15. Event Macros

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
10 Java files:

* ListMacros.java, FindMacros.java
    - lists/finds macro names

* TextMacro.java, UseLogo.java
    - use Office macros

* BuildForm.java
   - creates "build.odt" form

* ShowEvent.java
   - share macro

* GetNumber.java, GetText.java,
  NumActionListener.java
   - used in user extension macro and document macro

* DocEvents.java
   - sets up Office and document events

----------------------------
Other files and folders: 

* FormMacros\
* FormMacrosTest.odt 
* ShowEvent\
* FDM\
* Utils.jar

----------------------------
5 batch files for event macros:

* installMacros.bat
    - for zipping up a folder as an OXT, and installing it into Office
    - e.g. installMacros FormMacros

* extManager.bat
    - shows Office extension manager 

* unzipDoc.bat, zipDoc.bat
    - for unzipping/zipping FormDocMacros.odt
    - e.g. unzipDoc FormDocMacros.odt
           zipDoc FormDocMacros.odt

* execMacro.bat 
    - for executing macros from the command line
    - e.g. execMacro s build.odt ShowEvent.ShowEvent.show


----------------------------
Additional tools:

  * 7-zip  (not included here)
      - from http://www.7-zip.org/
      - used by my unzipDoc.bat & zipDoc.bat scripts

  * tr.exe, zip.exe
      - from https://github.com/bmatzelle/gow/wiki
      - Windows versions of the UNIX tr and zip tools
      - used by my installMacros.bat script


----------------------------
7 Office-related batch files:

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
Compilation:

> compile *.java
     - make sure it refers to the correct locations for your
       Java, JNA, and my Utils classes.


=================================
Usage:

1) Listing/finding macros

> run ListMacros Java

> run FindMacros "hello"


----------------
2) Running existing Office macros

> run TextMacro

> run UseLogo


----------------
3) Install ShowEvent.show share macro

> jar cvf ShowEvent.jar ShowEvent.class

> copy ShowEvent.jar into ShowEvent\

> copy ShowEvent\ to <OFFICE>\share\Scripts\java\
  (e.g. C:\Program Files\LibreOffice 5\share\Scripts\java\)
     -- you will need admin permissions to do this


----------------
4) Use ShowEvent.show macro in new "build.odt"

> run BuildForm
    -- creates build.odt form and its ShowEvent.show macro links

> open "build.odt"
    - type into FIRSTNAME textfield; press some buttons
    - look for the ShowEvent dialogs at the top-left of the desktop


-----------------
5) Install Form Macros as an extension

> copy GetNumber.class, GetText.class, NumActionListener.class
  to FormMacros\Utils

> installMacros.bat FormMacros
     -- creates FormMacros.oxt from FormMacros\, and installs it
        in Office
     -- type "yes" when copyright notice appears
     -- look for "Form Macros 0.1" in extension manager

> open "FormMacrosTest.odt" and configure events to use 
  the GetNumber.get and GetText.show macros


-----------------
6) Install Form Macros in a document

> **copy** and **rename** FormMacrosTest.odt to FormDocMacros.odt

> unzipDoc.bat FormDocMacros.odt
    -- creates FormDocMacros_odt\

> copy GetNumber.class, GetText.class, NumActionListener.class
  into Utils.jar; **rename** to Macros.jar
     -- use 7-zip to open Utils.jar and drag the class files into it

> copy FDM\dialogLibrary\ into FormDocMacros_odt\

> copy FDM\Scripts\ into FormDocMacros_odt\

> copy Macros.jar into FormDocMacros_odt\Scripts\java\Utils\

> edit FormDocMacros_odt\META-INF\manifest.xml
    -- add contents of FDM\ManifestExtras.txt to end, before </manifest:manifest>

> zipDoc.bat FormDocMacros.odt
   -- zips up FormDocMacros_odt\ as new FormDocMacros.odt

> open FormDocMacros.odt
   -- click on "Enable Macros" button
   -- change events to use document versions of GetNumber.get 
      and GetText.show macros
   -- save changes when you close the file


-----------------
7) Set up Office and Document versions of events

> run DocEvents
    
> save build.odt so it remembers the document event
> close build.odt

> open build.odt
    -- look for startApp and onLoad dialogs at top-left of desktop



=================================
Last updated: 16th December 2016
