
JLOP Chapter 16. Importing XML

From the website:

  Java LibreOffice Programming
  http://fivedots.coe.psu.ac.th/~ad/jlop

  Dr. Andrew Davison
  Dept. of Computer Engineering
  Prince of Songkla University
  Hat Yai, Songkhla 90112, Thailand
  E-mail: ad@fivedots.coe.psu.ac.th


If you use this code, please mention my name, and include a link
to the website.

Thanks,
  Andrew

============================
Java files:

  * FiltersInfo.java

  * ApplyInFilter.java, ApplyOutFilter.java

  * ExamineCompany.java

  * CreatePay.java, CreateAssoc.java

  * ExtractXMLInfo.java, BuildXMLSheet.java

  * UnmarshallPay.java, UnmarshallClubs.java, 
    UnmarshallWeather.java

----------------------------
XML-related files: 

  * pay.xml,
    payImport.xsl, payExport.xsl
    pay.xsd

  * clubs.xml,
    clubsImport.xsl, clubsExport.xsl, clubsTemplate.ott
    clubs.xsd

  * company.xml

  * weather.xml,
    weatherMod.xsd

----------------------------
Batch files for XML filtering:

  * infilter.bat
  * convert.bat

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

> compile *.java        // **EXCEPT** the 3 UnmarshallXXX.java examples
     - make sure compile.bat refers to the correct locations for your
       Java, JNA, and my Utils classes.


=================================
Usage:

  * Install Office Filters
      -- add a "Pay" filter using payImport.xsl, payExport.xsl
         (see section 1 of chapter, Fig. 3)

      -- add a "Clubs" filter using
         clubsImport.xsl, clubsExport.xsl, clubsTemplate.ott
         (see section 1.1 of chapter, Fig. 8)

----
  * Command line filtering
    e.g.
    > infilter pay.xml "Pay"    // only if "Pay" filter added to Office

    > convert payment.ods "xml:Pay"    
          // only if "Pay" filter added to Office, and there's a payment.ods file

----
  * Get Filter Information
    > run FiltersInfo.java

----
  * Use filters from Java
    e.g.
    > run ApplyInFilter pay.xml payImport.xsl payment.ods
        // creates payment.ods
    
    > run ApplyOutFilter payment.ods payExport.xsl payEx.xml
        // creates payEx.xml

----
  * Use DOM Parsing
    e.g.
    > run ExamineCompany
    > run CreatePay
    > run CreateAssoc

----
  * Extract labeled data from XML
    e.g.
    > run ExtractXMLInfo pay.xml
        // creates payXML.txt

    > run BuildXMLSheet payXML.txt

-----
  * Use XSD.
    The stages are explained in the comments of
    UnmarshallPay.java, UnmarshallClubs.java, and 
    UnmarshallWeather.java

    e.g.
    1. Convert XML to XSD using:
        http://www.freeformatter.com/xsd-generator.html
         -- used pay.xml
         -- used "Salami Slice" to create lots of simple Java classes later
         --> saved to pay.xsd
        OR use pay.xsd included in this folder.

    2. Converted XSD to Java classes:
         > xjc -p Pay pay.xsd

    3. Compiled contents of directory:
         > javac Pay/*.java 

    4. Usage:
        > javac UnmarshallPay.java
        > java UnmarshallPay


=================================
Last updated: 6th January 2017
