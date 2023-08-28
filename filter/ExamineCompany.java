
// ExamineCompany.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Dec. 2016

/* Print out details extracted from company.xml
*/


import java.io.*;
import java.util.*;

import javax.xml.parsers.*;
import org.w3c.dom.*;



public class ExamineCompany
{

  public static void main(String[] args) throws Exception
  {
    Document doc = XML.loadDoc("company.xml");
    
    // Get the document's root
    NodeList root = doc.getChildNodes();
    
    // move down the tree to the executive in the first company node
    Node comps = XML.getNode("Companies", root);
    Node comp = XML.getNode("Company", comps.getChildNodes());
    Node exec = XML.getNode("Executive", comp.getChildNodes());
    
    // print the executive's data
    String execType = XML.getNodeAttr("type", exec);
    NodeList exNodes = exec.getChildNodes();
    String lastName = XML.getNodeValue("LastName", exNodes);
    String firstName = XML.getNodeValue("FirstName", exNodes);
    String street = XML.getNodeValue("street", exNodes);
    String city = XML.getNodeValue("city", exNodes);
    String state = XML.getNodeValue("state", exNodes);
    String zip = XML.getNodeValue("zip", exNodes);
    
    System.out.println(execType);
    System.out.println(lastName + ", " + firstName);
    System.out.println(street);
    System.out.println(city + ", " + state + " " + zip);

    // get all the data in the tree for a given node/tag name
    NodeList lnNodes = doc.getElementsByTagName("LastName");
    ArrayList<String> lastnames = XML.getNodeValues(lnNodes);
    System.out.println("All lastnames:");
    for(String lastname: lastnames)
      System.out.println("  " + lastname);
  }  // end of main()

}  // end of ExamineCompany class
