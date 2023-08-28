
// CreateAddonImpl.java
// Andrew Davison, Oct 2016, ad@fivedots.coe.psu.ac.th

/* Uses FreeMarker to create an Addon implementation file by instantiating
   the ADDON_FTL template with the addonName (e.g. "Highlight").

   FreeMarker replaces three data fields in the template:

      "className"   : addonName + "AddonImpl"
      "extensionID  :  PACKAGE_NAME + addonName.toLowercase() + "Addon"
      "cmdNames"    : [ addonName, "help" ]

   The resulting file is: <addonName>AddonImpl.java
   if a file of this name already exists, then a unique number is added
   to the filename (e.g. <addonName>AddonImpl1.java). The existing file
   is NOT overwritten.

   Usage:
     compile CreateAddonImpl.java
     run CreateAddonImpl EzHighlight
            -- note "F" so FreeMarker is in the classpath
*/


import java.io.*;
import java.util.*;

import freemarker.template.*;
import freemarker.cache.*;


public class CreateAddonImpl
{

  private static final String PACKAGE_NAME = "org.openoffice.";

  private static final String ADDON_FTL = "AddonImpl.ftl";


  public static void main(String[] args)
  {
    if (args.length != 1) {
      System.out.println("Usage: java AddonDataReader <addon-name>");
      return;
    }
    String addonName = args[0];


    // Freemarker configuration object
    Configuration cfg = new Configuration(new Version(2, 3, 21));
    try {
      cfg.setTemplateLoader( new FileTemplateLoader( new File(".")) );

      System.out.println("Loading the template " + ADDON_FTL);
      Template template = cfg.getTemplate(ADDON_FTL);

      // build the data-model
      Map<String, Object> data = new HashMap<String, Object>();

      data.put("className", addonName + "AddonImpl");
      data.put("extensionID", PACKAGE_NAME + addonName.toLowerCase() + "Addon");
      data.put("cmdNames", Arrays.asList(addonName, "help"));

      // File output
      String fnm = Info.getUniqueFnm(addonName+ "AddonImpl.java");
      System.out.println("Generating " + fnm);
      Writer file = new FileWriter(new File(fnm));
      template.process(data, file);
      file.flush();
      file.close();
    }
    catch (IOException e) {
      System.out.println(e);
    }
    catch (TemplateException e) {
      System.out.println(e);
    }
  }  // end of main()




}  // end of CreateAddonImpl class
