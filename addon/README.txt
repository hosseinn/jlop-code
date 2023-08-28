
JLOP Chapter 13. Add-ons

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
Directory structure:

AddOn Tests
|   AddonImpl.ftl
|   compile.bat
|   CreateAddonImpl.java
|   extList.bat
|   extManager.bat
|   EzHighlightAddonImpl - Complete.java
|   ImageHex.java
|   installAddon.bat
|   loDoc.bat
|   lofind.bat
|   lofindTemp.txt
|   loGuide.bat
|   loKill.bat
|   loList.bat
|   manifest.txt
|   run.bat
|   story.odt
|   tr.exe
|   zip.exe
|   
\---EzHighlight
    |   Addons.xcu
    |   description.xml
    |   license.txt
    |   package-description.txt
    |   ProtocolHandler.xcu
    |   Utils.jar
    |   
    +---dialogLibrary
    |       Basic.xdl
    |       EzHighlight.xdl
    |       
    +---images
    |       ezhighlight.png
    |       ezhighlight16.png
    |       ezhighlight26.png
    |       
    \---META-INF
            manifest.xml
            

=================================
Building the EzHighlight Add-on

> compile CreateAddonImpl.java


> run CreateAddonImpl EzHighlight
     -- creates a partial EzHighlightAddonImpl.java using the
        AddonImpl.ftl template

     -- make sure compile.bat & run.bat refer to the 
        correct locations for Java, JNA, FreeMarker, 
        and my Utils classes


// finish off implementing EzHighlightAddonImpl.java, or rename
   EzHighlightAddonImpl - Complete.java


> installAddon EzHighlight
     -- compiles EzHighlightAddonImpl.java, creates a JAR, creates an OXT file,
        installs the extension in Office

> extManager
    -- OPTIONAL: this starts the Extension Manager
> extList
    -- OPTIONAL: this gives more information about the 
       installation status of the extensions in Office


=================================
Using the EzHighlight Add-on

Start Office using story.odt, and access the 
EzHighlight menu or toolbar items


----------------------------
FreeMarker

The FreeMarker template library is available at: http://freemarker.org/

I installed it in D:\freemarker


----------------------------
Additional tools used by installAddon.bat:

  * tr.exe, zip.exe
      -- from https://github.com/bmatzelle/gow/wiki
      -- Windows versions of the UNIX tr and zip tools


----------------------------
There are 7 Office-related batch files:

  * compile.bat
  * run.bat
     - make sure they refer to the correct locations for
       Java, JNA, FreeMarker, and my Utils classes.
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

Creating your own Add-on
------------------------
Decide on a name for your add-on e.g. FooBar


> run CreateAddonImpl FooBar
     -- creates a partial FooBarAddonImpl.java


// finish off implementing FooBarAddonImpl.java


Rename the EzHighlight/ folder to FooBar/


Modify the configuration files inside FooBar/ and manifest.txt
  -- search for "EzHighlight" and change all occurrences to "FooBar"
  -- search for "ezhighlight" and change all occurrences to "foobar"
  -- don't forget to search the dialogLibrary/, images/, and META_INF/ folders
  -- rename files in the same way (e.g. EzHighlight.xdl --> FooBar.xdl)


> installAddon FooBar


Comment out the calls to Console in FooBarAddonImpl.java when the add-on
is working correctly. 


=================================
Last updated: 14th October 2016
