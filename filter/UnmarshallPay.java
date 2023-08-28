
// UnmarshallPay.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Dec 2016

/* Manipulate XML data as Java objects. 
   No use made of Office.

    1. Convert XML to XSD using:
        http://www.freeformatter.com/xsd-generator.html
         -- used pay.xml
         -- used "Salami Slice" to create lots of simple Java classes later
         --> saved to pay.xsd

    2. Converted XSD to Java classes:
         > xjc -p Pay pay.xsd

    3. Compiled contents of directory:
         > javac Pay/*.java 

    4. Usage:
        > javac UnmarshallPay.java
        > java UnmarshallPay

*/

import java.io.*;
import java.util.*;
import javax.xml.bind.*;

import Pay.*;


public class UnmarshallPay 
{
  public static void main(String[] args) 
  {
    try {
      JAXBContext jaxbContext = JAXBContext.newInstance(Payments.class);
      Unmarshaller jaxbUnmarshaller = jaxbContext.createUnmarshaller();
      Payments pays = (Payments) jaxbUnmarshaller.unmarshal( new File("pay.xml"));

      List<Payment> payList = pays.getPayment();

      // print all payment names and amounts
      System.out.println("Payments");
      for(Payment p : payList)
        System.out.println("  " + p.getPurpose() + ": " + p.getAmount());
    } 
    catch (JAXBException e) {
      e.printStackTrace();
    }
  }  // end of main()

}  // end of UnmarshallPay class