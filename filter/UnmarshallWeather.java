
// UnmarshallWeather.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Dec 2016

/* Manipulate XML data as Java objects. 
   No use made of Office.

    1. Convert XML to XSD using:
        http://www.freeformatter.com/xsd-generator.html
         -- used weather.xml
         -- used "Salami Slice" to create lots of simple Java classes later
         --> saved to weather.xsd

    1.5. Fix weather.xsd to distinguish between different "value" attributes.
         Also changed weather "icon" type to string to deal with "d" in value.
         Or simply rename weatherMod.xsd to be weather.xsd

    2. Converted XSD to Java classes:
         > xjc -p Weather weather.xsd

    3. Compiled contents of directory:
         > javac Weather/*.java 

    4. Usage:
        > javac UnmarshallWeather.java
        > java UnmarshallWeather

*/

import java.io.*;
import java.util.*;
import java.text.*;

import javax.xml.bind.*;
import javax.xml.datatype.*;

import Weather.*;



public class UnmarshallWeather 
{
  public static void main(String[] args) 
  {
    try {
      JAXBContext jaxbContext = JAXBContext.newInstance(Current.class);
      Unmarshaller jaxbUnmarshaller = jaxbContext.createUnmarshaller();
      Current currWeather = (Current) jaxbUnmarshaller.unmarshal( new File("weather.xml"));

      // get precipitation as a boolean
      String rainingStatus = currWeather.getPrecipitation().getValue();
      boolean isRaining = rainingStatus.equals("yes");

      // get date in day/month/year format
      XMLGregorianCalendar gCal = currWeather.getLastupdate().getValueAttribute();
      Calendar cal = gCal.toGregorianCalendar();
      SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
      formatter.setTimeZone(cal.getTimeZone());
      String dateStr = formatter.format(cal.getTime());

      if (isRaining)
        System.out.println("It was raining on " + dateStr);
      else
        System.out.println("It was NOT raining on " + dateStr);
    } 
    catch (JAXBException e) {
      e.printStackTrace();
    }
  }  // end of main()

}  // end of UnmarshallWeather class