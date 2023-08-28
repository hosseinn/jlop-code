# Source Code for "Java LibreOffice Programming" (JLOP) Book
This repository contains source code for the "Java LibreOffice Programming" (JLOP) book. The book is written by Dr. Andrew Davison and is available from [here](https://web.archive.org/web/20220807000947/http://fivedots.coe.psu.ac.th/~ad/jlop/). A slightly updated version of the book is also available from [Github pages](https://flywire.github.io/lo-p/), and its source text is also [available](https://github.com/flywire/lo-p).

As the name of the book implies, the source code is written in Java. To compile and run, you should have Java development kit (JDK) installeed on your system.

## Accompanied Scripts
The `scripts/` folder contains scripts to build and run the code. These new scripts are usable on both Linux/macOS (.sh files) and Windows (.bat files).

## Build
To build the code, first you should get the required external libraries. To do that, go to the `lib` folder and use the `download` script. You may have to unzip some of the files manually. The script uses `wget` on Linux/macOS, and `curl` on Windows.

On Linux/macOS:

```
cd lib
./download.sh
```

On Windows:

```
cd lib
download.bat
```

Note: If you use Windows older than Windows 10, you may have to download `curl.exe` manually.

After download is complete, you need to build Utils.jar. To do that:

On Linux/macOS:

```
cd Utils
../scripts/compile.sh
./toJar.sh
```
And on Windows:

```
cd Utils
..\scripts\compile.bat
toJar.bat
```

Then you should have Utils.jar in `Utils/` folder.

To run the examples, you should got to a directory, and use the `run` script.

On Linux/macOS, one can run the `HelloText` example from the `text/` folder this way:

```
cd text
../scripts/compile.sh
../scripts/run.sh HelloText
```

And on Windows:
```
cd text
..\scripts\compile.bat
..\scripts\run.bat HelloText
```

Other examples can be run with almost the same syntax. Many of the `.class` files can be run this way. Please note that running `compile.sh` only once in each folder is enough. Sometimes, you have to run additional commands, which are usually described in the book, or in the `README` files.

## Known Issues
Not all of the examples currently run perfectly. Some of the problems are as follows:

1. Examples related to text to speach (freetts) has problems because of the problems in freetts itself.

2. The `idlc` tool is removed from recent versions of LibreOffice, and now one is supposed to use `unoidl-read`/`unoidl-write` instead. `javamaker` is used in the examples with `.idl` files to generate code. For example, to compile files in `addin` folder, one has to run this command to generate appropriate classes for `Doubler`:

```
/opt/libreoffice7.5/sdk/bin/javamaker /opt/libreoffice7.5/program/types.rdb /opt/libreoffice7.5/program/types/offapi.rdb Doubler.idl
```

In Windows, that would be:

```
"C:\Program Files\LibreOffice\sdk\bin\javamaker.exe" "C:\Program Files\LibreOffice\program\types.rdb" "C:\Program Files\LibreOffice\program\types\offapi.rdb" Doubler.idl
```

3. `jaxb` has a small problem with `xjc`. You have to change `jaxb-ri\bin\xjc.sh` (or `jaxb-ri/bin/xjc.bat`) to use `javax.activation-api.jar` instead of `javax.activation.jar`. In Windows, one should also change `lib/relaxngDatatype.jar` to `mod/relaxng-datatype.jar` in the same file. This is documented in [issue 1273](https://github.com/eclipse-ee4j/jaxb-ri/issues/1273).

The command in `filter` folder to generate code from `pay.xsd` is:

```
../lib/jaxb-ri/bin/xjc.sh -p Pay pay.xsd
```
In Windows, it would be:

```
..\lib\jaxb-ri\bin\xjc.bat -p Pay pay.xsd
```

4. Processing the two files `clubs.xsd` and `weatherMod.xsd` currently does not pass.

5. Some of the scripts are Window-only at the moment. This also reflects in part of the utility code in `Base.java`, `Lo.java` and `JMail.java`.

6. Several `.bat` files are still in the folders. These are temporarily preserved, and the plan is to remove them in the future. Most of the scripts will remain in `scripts` folder.
