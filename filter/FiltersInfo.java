
// FiltersInfo.java
// Andrew Davison, ad@fivedots.coe.psu.ac.th, Dec. 2016

/* 
*/

import java.io.*;

import com.sun.star.lang.*;
import com.sun.star.frame.*;
import com.sun.star.beans.*;



public class FiltersInfo
{

  public static void main(String[] args) 
  {
    XComponentLoader loader = Lo.loadOffice();

    String[] filterNms = Info.getFilterNames();
    System.out.println("Filter Names");
    Lo.printNames(filterNms, 3);

    PropertyValue[] props = Info.getFilterProps("AbiWord");
    Props.showProps("AbiWord Filter", props);

    props = Info.getFilterProps("Pay");
    Props.showProps("Pay Filter", props);

    props = Info.getFilterProps("Clubs");
    Props.showProps("Clubs Filter", props);
    if (props != null) {
      int flags = (int) Props.getValue("Flags", props);

      System.out.println("Filter flags: " + Integer.toHexString(flags));
      System.out.println(" Import: " + Info.isImport(flags));
      System.out.println(" Export: " + Info.isExport(flags));
    }
    Lo.closeOffice();
  }   // end of main()



}  // end of FiltersInfo class
