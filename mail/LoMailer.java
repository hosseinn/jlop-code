
// LoMailer.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/*  Two approaches to sending e-mail using the Office APIs:

       1) Mail.sendEmail() uses the MailServiceProvider service
           -- for SSL communication to work I had to move DLLs
              in Office; see the chapter and the Mailer.java comments

       2) Mail.sendEmailByClient() uses the SimpleSystemMail or 
          SimpleCommandMail service to communicate with the 
          OSes default e-mail client.
          A "Confirm" Dialog appears before the e-mail is sent.

     See JMailer.java for examples using only Java's e-mail APIs.
*/


import com.sun.star.uno.*;
import com.sun.star.lang.*;
import com.sun.star.frame.*;


public class LoMailer
{

  public static void main(String args[])
  {
    if (args.length != 1) {
      System.out.println("Usage: run LoMailer <mail-account-password>");
      return;
    }
    String password = args[0];

    XComponentLoader loader = Lo.loadOffice();

/*
    Mail.sendEmail("fivedots.coe.psu.ac.th", 25,
          "ad", password,
          "dd@foobarZZ.com",
          "LOMail Test 1", "LOMail Body 1", "skinner.png");
*/
/*
    Mail.sendEmail("fivedots.coe.psu.ac.th", 25,
          "ad", password,
          "dd@foobarZZ.com",
          "LOMail Test 2", "LOMail Body 2", "addresses.ods");
*/


    Mail.sendEmail("smtp.gmail.com", 587, 
          "Andrew.Davison50@gmail.com", password,
          "dd@foobarZZ.com",
          "LOMail Gmail Test 1", "LOMail Gmail Body Test 1", 
                  "skinner.png");
                 //"addresses.ods");


/*
    Mail.sendEmailByClient("dd@foobarZZ.com", "TB Mail Test", 
                            "TB Mail Body", "skinner.png");
*/

    Lo.closeOffice();
  }  // end of main()


}  // end of LoMailer class
