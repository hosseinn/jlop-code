
// ExtractXMLInfo.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Dec. 2016

/* Extract all element and attribute names and values
   from <fnm>.xml as indented labelled text without XML syntax.
   Data is stored in <fnm>XML.txt

   No use made of Office, but using utilities in XML.java

   The generated files are read in by BuildXMLSheet.java

   Usage:
      > compile ExtractXMLInfo.java

      > run ExtractXMLInfo pay.xml
      > run ExtractXMLInfo clubs.xml
      > run ExtractXMLInfo weather.xml     // plenty of attributes
      > run ExtractXMLInfo company.xml
*/


import java.io.*;

import javax.xml.parsers.*;
import org.w3c.dom.*;



public class ExtractXMLInfo
{

  public static void main(String[] args) throws Exception
  /* Open the XML file as a DOM tree. and print its data
     and attributes as labelled text to a new file.
  */
  {
    if (args.length != 1) {
      System.out.println("Usage: run ExtractXMLInfo <XML file>");
      return;
    }

    Document doc = XML.loadDoc(args[0]);
    if (doc == null)
      return;

    String fname = Info.getName(args[0]);
    String outFnm = fname + "XML.txt";
    System.out.println("Writing XML data from " + args[0] + " to " + outFnm);
    PrintWriter pw = new PrintWriter(new FileWriter(outFnm));

    NodeList root = doc.getChildNodes();
    // there may be multiple trees; visit each one
    for (int i = 0; i < root.getLength(); i++) {
      visitNode(pw, root.item(i), "");
      pw.write("\n");
    }
    pw.close();
  }  // end of main()


  private static void visitNode(PrintWriter pw, Node node, String ind)
  /* Visit a node by printing its name, any attribute data,
     any text node data, and then recursively visiting 
     all the node's children.
  */
  {
    pw.write(ind + node.getNodeName());
    visitAttrs(pw, node);

    // examine all the child nodes
    NodeList nodeList = node.getChildNodes();
    for (int i = 0; i < nodeList.getLength(); i++) {
      Node child = nodeList.item(i);
      if (child.getNodeType() == Node.TEXT_NODE) {
        String trimmedVal = child.getNodeValue().trim();
        if (trimmedVal.length() == 0)
          pw.write("\n");
        else
          pw.write(": \"" + trimmedVal + "\"");
             // element names with values end with ':'
      }
      else if (child.getNodeType() == Node.ELEMENT_NODE)
        visitNode(pw, child, ind+"  ");
    }
  } // end of visitNode()



  private static void visitAttrs(PrintWriter pw, Node node) 
  // print all the attributes -- name and data
  {
    NamedNodeMap attrs = node.getAttributes();
    if (attrs != null) {
      for (int i = 0; i < attrs.getLength(); i++) {
        Node attr = attrs.item(i);
        pw.write("  " + attr.getNodeName() + "= \"" + 
                               attr.getNodeValue() + "\"");
             // attribute names end with '='
      }
    }
  }  // end of visitAttrs()


}  // end of ExtractXMLInfo class
