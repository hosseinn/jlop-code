
// Speaker.java
// Andrew Davison, July 2015, ad@fivedots.psu.ac.th

/* Instructions for installing and using FreeTTS, Mbrola, and Voices

   1. Get FreeTTS
   - Go to  http://freetts.sourceforge.net/docs/index.php

   - Download freetts-1.2.2-bin.zip 
     from the "Downloading and Installing" section

   - Unzip the "freetts-1.2" folder to D:\freetts-1.2

   ------
   2. Get Mbrola and 3 US Voices
   Go to  http://tcts.fpms.ac.be/synthesis/mbrola.html

   This may direct you to "The MBROLA Project" page instead of the download page.
   In that case, 
     - Click on the "Download" text to the left of the images.
     - Look for the section called "What you will have to copy", 
       and click on the "MBROLA binary and voices" text.

   You should now be on a page headed "Copying the MBROLA Bin and Databases"
   Select the following:
     - PC-DOS --> the mbr301d.zip file will be downloaded
     - the three voices "us1", "us2", "us3"
          the three 'voice' files that are downloaded are called
          us1-980512.zip, us2-980812.zip, us3-990208.zip

   Extract the Mbrola binary and voices:
      - extract "mbrola.exe" from mbr301d.zip
      - extract the folders "us1", "us2", "us3" from the zipped 'voice' files

   Install the Mbrola binary and voices in FreeTTS:
     - move "mbrola.exe", "us1", "us2", "us3" to "D:\freetts-1.2\mbrola\"
       (there is already a mbrola.jar file in that folder)


   ----
   3. Use compileTalk.bat and runTalk.bat to compile and run this file:

      > compileTalk Speaker.java

      > runTalk Speaker
         // You should hear:
            "hi everybody, my name is andrew, my voice is somewhat robotic"

   Although those batch files refer to LibreOffice, this example only uses
   FreeTTS.
*/


import com.sun.speech.freetts.*;


public class Speaker
{
  private static final String VOICE_NAME = "mbrola_us2";
                     // "mbrola_us3";   // male
                     // "mbrola_us2";   // male, quite clear
                     // "mbrola_us1";   // female
                     // "kevin";        // low quality
                     // "kevin16";      // medium
                     // "alan";         
                /* high quality but limited to speaking the time
                   e.g. "The time is now, a little after six in the morning." 
                */

  private Voice voice = null;


  public Speaker()
  { this(VOICE_NAME);  }


  public Speaker(String name)
  { voice = VoiceManager.getInstance().getVoice(name);
    if (voice == null) {
      System.out.println("Could not find the voice: " + name);
      System.exit(1);
    }

    System.out.println("Using the voice: " + name);
    voice.allocate();
  }  // end of Speaker()


  public void say(String[] thingsToSay)
  { for (String s : thingsToSay)
      say(s);
  }


  public void say(String s)
  { if (voice != null)
      voice.speak(s);
  }


  public void dispose()
  { if (voice != null)
      voice.deallocate();
  }


  // --------------------------- test rig -------------

  public static void main(String[] args)
  {
    Speaker speaker;
    if (args.length != 0)
      speaker = new Speaker(args[0]);
    else
      speaker = new Speaker();

    String[] thingsToSay = new String[]
      {   "hi everybody", "my name is andrew", "my voice is somewhat robotic",
      };

    speaker.say(thingsToSay);
    speaker.dispose();
  }  // end of main()

}  // end of the Speaker class

