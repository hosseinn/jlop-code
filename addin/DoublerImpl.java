
// DoublerImpl.java
// Andrew Davison, October 2016, ad@fivedots.coe.psu.sc.th

/* Code implementing the functions defined in Doubler.idl
   for the XDoubler interface.
*/

import java.util.*;  // ADDED


import com.sun.star.uno.XComponentContext;
import com.sun.star.lang.*;   // CHANGED; made more general
import com.sun.star.registry.XRegistryKey;

// helper documentation is part of "Java UNO Reference"
import com.sun.star.lib.uno.helper.WeakBase;
import com.sun.star.lib.uno.helper.Factory;


public final class DoublerImpl extends WeakBase
             implements com.sun.star.lang.XLocalizable,
                        com.sun.star.lang.XServiceInfo,
                        org.openoffice.doubler.XDoubler
{
  private static final String LOG_FNM = "c:\\arrayInfo.txt";  // ADDED
              // log file for storing debugging output


  private final XComponentContext xcc;

  private static final String implName = DoublerImpl.class.getName();

  private static final String[] serviceNames = 
                                { "org.openoffice.doubler.Doubler" };

  private com.sun.star.lang.Locale locale = new com.sun.star.lang.Locale();

  // private XComponent doc;    // ADDED


  public DoublerImpl(XComponentContext context)
  {  xcc = context;
     // doc = Lo.addonInitialize(xcc);   // ADDED so all my utils can be used safely
  }


  public static XSingleComponentFactory __getComponentFactory(String name) 
  {
    XSingleComponentFactory xFactory = null;
    if (name.equals(implName))
      xFactory = Factory.createComponentFactory(DoublerImpl.class, serviceNames);
    return xFactory;
  }

 
  public static boolean __writeRegistryServiceInfo(XRegistryKey regKey) 
  {  return Factory.writeRegistryServiceInfo(implName, serviceNames, regKey); }


  // -------- XLocalizable methods ------------

  public void setLocale(com.sun.star.lang.Locale l)
  {  locale = l; }


  public com.sun.star.lang.Locale getLocale()
  {  return locale;  }


  // -------- XServiceInfo methods ------------

  public String getImplementationName() 
  {  return implName; }


  public boolean supportsService(String sService) 
  {
    for(int i=0; i < serviceNames.length; i++)
      if (sService.equals(serviceNames[i]))
        return true;
    return false;
  }


  public String[] getSupportedServiceNames() 
  {  return serviceNames; }


  // ------------------- XDoubler function(s) -----------------

  public double doubler(double value)
  {  
/*
    FileIO.appendTo(LOG_FNM, "Window title: " + GUI.getTitleBar());
    FileIO.appendTo(LOG_FNM, "Services for this document:");
    for(String service : Info.getServices(doc))
      FileIO.appendTo(LOG_FNM, "  " + service);
*/
    return value*2;  
  } 


  public double doublerSum(double[][] vals)
  { 
    double sum = 0;
    for (int i = 0; i < vals.length; i++)
      for (int j = 0; j < vals[i].length; j++)
        sum += vals[i][j]*2;
    return sum;
  } // end of doublerSum()


  public double[][] sortByFirstCol(double[][] vals)
  { 
    // FileIO.appendTo(LOG_FNM, Lo.getTimeStamp() + ": sortByFirstCol()");
    selectionSort2(vals);

    //for (int i = 0; i < vals.length; i++)
    //  FileIO.appendTo(LOG_FNM, "  " + Arrays.toString(vals[i]));

    return vals;
  }  // end of sortByFirstCol()


  private void selectionSort(double[][] vals)
  // ascending order based on first column of vals; FAILS ??
  {
    Arrays.sort(vals, new Comparator<double[]>() {
      public int compare(double[] row1, double[] row2) 
      // compare first column of each row
      {  
        // FileIO.appendTo(LOG_FNM, "compared");   // never reached ??
        return Double.compare(row1[0], row2[0]);
      }
    });
  }  // end of selectionSort()


  private void selectionSort2(double[][] vals)
  // ascending order based on first column of vals; WORKS
  {
    double[] temp; 
    for(int i = vals.length-1; i > 0; i--) {
       int first = 0; 
       for(int j = 1; j <= i; j ++) { 
         if(vals[j][0] > vals[first][0]) // compare first col values
           first = j;
       }
       temp = vals[first]; // swap rows
       vals[first] = vals[i];
       vals[i] = temp; 
     }           
   }  // end of selectionSort2()


}  // end of DoublerImpl class
