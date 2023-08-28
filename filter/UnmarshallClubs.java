
// UnmarshallClubs.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Dec 2016

/* Manipulate XML data as Java objects. 
   No use made of Office.

    1. Convert XML to XSD using:
        http://www.freeformatter.com/xsd-generator.html
         -- used clubs.xml
         -- used "Salami Slice" to create lots of simple Java classes later
         --> saved to clubs.xsd

    2. Converted XSD to Java classes:
         > xjc -p Clubs clubs.xsd

    3. Compiled contents of directory:
         > javac Clubs/*.java 

    4. Usage:
        > javac UnmarshallClubs.java
        > java UnmarshallClubs

*/

import java.io.*;
import java.util.*;
import javax.xml.bind.*;
import Clubs.*;


public class UnmarshallClubs 
{
  public static void main(String[] args) 
  {
    try {
      JAXBContext jaxbContext = JAXBContext.newInstance(ClubDatabase.class);
      Unmarshaller jaxbUnmarshaller = jaxbContext.createUnmarshaller();
      ClubDatabase cd = (ClubDatabase) jaxbUnmarshaller.unmarshal(new File("clubs.xml"));

      List<Association> assocList = cd.getAssociation();

      // print club names in all the associations
      System.out.println("Associations");
      for(Association assoc : assocList) {
        System.out.println("  " + assoc.getId());
        List<Club> clubs = assoc.getClub();
        for(Club club : clubs)
          System.out.println("    " + club.getName());
      }
    } 
    catch (JAXBException e) {
      e.printStackTrace();
    }
  }  // end of main()

}  // end of UnmarshallClubs class
