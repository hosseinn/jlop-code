
// JMailer.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, July 2016

/* Three ways to send e-mail using Java, without using the
   Office API.

     1) JMail.sendEmail()
           - uses JavaMail (javax.mail.jar)
           - requires mcompile.bat, mrun.bat

     2) JMail.sendEmailByClient()
          - uses Java's Desktop API to access OSes default e-mail client
          - attachments cannot be added

     3) JMail.sendEmailByTB()
          - uses a Java's Desktop API (TBExec.bat) to use the Thunderbird client
            directly

     Both (2) and (3) bring up the e-mail client application, and the
     user must press the "send" button in the app to complete the send.

     See LoMailer.java for examples using the Office e-mail APIs.

   ---
   Usage: compile JMailer.java
          run JMailer

   Both batch files have been extended to access javax.mail.jar
*/


public class JMailer
{

  public static void main(String[] args)
  {
    if (args.length != 1) {
      System.out.println("Usage: run JMailer <mail-account-password>");
      return;
    }
    String password = args[0];

/*
    JMail.sendEmail("fivedots.coe.psu.ac.th", 25, 
                "ad", password,
                "dd@foobarZZ.com",
                "JMail Email Test 1", "JMail Email Body Test 1", "skinner.png");
*/

/*
    JMail.sendEmail("fivedots.coe.psu.ac.th", 25, 
                "ad", password,
                "dd@foobarZZ.com",
                "JMail Email Test 2", "JMail Email Body Test 2", "addresses.ods");
*/


     JMail.sendEmail("smtp.gmail.com", 587, 
                "Andrew.Davison50@gmail.com", password,
                "dd@foobarZZ.com",
                "JMail Gmail Test 1", "JMail Gmail Body Test 1", 
                // "skinner.png");
                "addresses.ods");

/*
    JMail.sendEmailByClient("dd@foobarZZ.com", "Client Mail Test 1",
                             "Client Mail Body 1");   // no attachment can be included
*/

/*
    JMail.sendEmailByTB("dd@foobarZZ.com", "TB Mail Test 2", 
                                   "TB Mail Body 2", "addresses.ods");
*/
  }  // end of main()



}  // end of JMailer class