
// MailMerge.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Before this program can run:

   * The Addresses.ods spreadsheet must be connected to an Address.odb
     data source whose name is Addresses.
  
   * The Addresses.ods sheet must be called "Addresses".
        
   * Fields must be copied from the Addresses.odb table into the Writer
     template file, formletter.ott.

   * The Addresses.ods spreadsheet must have a column called "E-mail"
     which holds e-mail addresses.


   For info on how to do this read Chapter 11 of the Writer Guide,
   and my chapter.

   ---
   This example generates 6 letters based on 
   merging the 6 rows of spreadsheet data and the form letter. 
   It either saves, prints, or e-mails them depending on which "merge"
   function is uncommented.

   Office reports a crash at the end, but terminates correctly.

   Usage:
     > compile MailMerge.java
     > run MailMerge
*/


import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.beans.*; 




public class MailMerge
{ 
  private static final String DATA_SOURCE_NAME = "Addresses"; 
               // data source name for Addresses.odb

  private static final String TABLE_NAME = "Addresses"; 
               // name of sheet in the spreadsheet

  private static final String TEMPLATE_FNM = "formLetter.ott"; 



  public static void main(String[] args)
  { 
    String password = null;
    if (args.length == 1)
      password = args[0];
    else
      System.out.println("Using a null password...");

    Lo.loadOffice();

    // Mail.mergeLetter(DATA_SOURCE_NAME, TABLE_NAME, TEMPLATE_FNM, true);
                                                      // single file output


 //   Mail.mergePrint(DATA_SOURCE_NAME, TABLE_NAME, TEMPLATE_FNM, "FinePrint", false);
                                   // one print job holding all letters


    // Sends letters as PDF attachments to the 6 e-mails in the spreadsheet.
    Mail.mergeEmail(DATA_SOURCE_NAME, TABLE_NAME, TEMPLATE_FNM,
                   password, "Hello from Andrew at General Supplies", 
                   "Please read the attached message.");

    Lo.closeOffice();
  }  // end of main()


}  // end of MailMerge class
