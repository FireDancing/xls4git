# xls4git
enables source-code revisioning of XML-based Microsoft&reg; Excel-files.

## Motivation
This is an attempt to enable source-code versioning to VBA-code in Microsoft&reg; Excel-files, as suitable open-source solutions could not be found.

In addition to this my main motivation was to give something to the community. Afte using various open-source-software for more than 20 years now, I think it is time also to contribute something back.

## How it works
xls4git uses a pyhton-script to decompose the (zipped) Excel-sheet in pieces which can be imported into a source-code-revision tool, such as git. During decomposition basically two parts - each in a separate subfolder - are created:
* xml: the complete XML-data of the Excel-file itself
* vba: the complete source-code of the VBA-project, including forms

After decomposition, the user is free to work with the single particular files.

Finally the decomposed files can be used to re-build the Excel-file from the sources.

**Remark**:
Some of the required work was not possible to perform with the python-script. Hence parts of the functions were appointed to a particular Excel-makro (x4g_VBAHandler.xlsm). The user does not need to do anything with this by his own. However, for convenience reaons, this VBAHandler allows manual exection as well (see according buttons on sheet).
The VBAHandler performs the actual export of the VBA-code from the particular Excel-file.

## Usage
Call the python script with the location of the config-file (argument -c) and the according action (argument -a)
* -a (or --action) export: decomposes the according Excel-file
* -a (or --action) build: rebuild the Excel-sheet from the earlier decomposed source.

The configuration is done in the x4g.ini-file

## Limitations
* This project was done mainly as "practical research" (i.e. trial and error) and was tested on one (quite big) internal Excel-project until xls4git worked. However, it is likely that in the depths of the internal XML-structure of Excel cases might occur, which are not yet considered.
* Works only at Microsoft&reg; Windows platforms

## What I would like to do, if I had much for time for this
* get rid of the VBAHandler, if somehow possible
* better understand the internal structure of the XML-based Excel-sheets.

## Feedback
Any feedback and suggestions for improvement are greatly welcome.
