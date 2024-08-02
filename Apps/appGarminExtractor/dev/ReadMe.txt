Title: appGarminExtractor

Purpose:
	This experimental tool is meant to allow field users to extract GPS data from Garmin units.

	The Garmin.hta tool will prompt you to connect the Garmin unit to the computer via USB.  
	When you confirm, it will read the data from the unit and save it as a GPX (XML) file.
	The GPX2HTA stylesheet will process the GPX file and display a table of available track segments
	with summary stats (distance, average speed, # points, etc.).  You select the track segments
	you want to import by setting a check box next to the segment.  You can also include metadata
	about the platform (Platform type, allegiance, Garmin unit number, and platform identifier).
	Once you confirm the track selection, the GPX2CSV stylesheet will convert the selected track
	segments to a CSV file and calculate some first derivative values for QC purposes.
	
Status:
	- This utility is intended to be a gap filler only to meet an immediate need
	- For the more complete solution, see <appNCR-IADS.GPS Importer Plugin>
	- Limited testing has been performed at the development site and at VIGILANT SHIELD (2011)

History:
	20110812 - TAE - Modified to handle the case where saved track logs are present
	20110708 - TAE - Original code

Author:
	Tom Elkins (tom.elkins@mercurysolutions.com)

Compatibility:
	- [X] Windows XP : No issues noted
	- [X] Windows 7 32-bit : No issues noted
	- [X] Windows 7 64-bit : No issues noted

Files:
	Garmin USB Driver.exe - OPTIONAL.  If the host machine does not already have the Garmin USB driver installed, use this file to install them.
		Alternatively, you can download the latest drivers from the Garmin website www.garmin.com
	Garmin.hta - The application.  Running this HTA walks you through the process and serves as the user interface
	gpsbabel.exe - REQUIRED.  GPSBabel is a third-party application that supports a wide variety of GPS devices and data formats.
		It uses a command-line interface with a lot of options.  The HTA formats a specific set of options for extracting Garmin-formatted data from the USB port and writing to a GPX file.
		If this file is missing here, you can download it from the internet (www.gpsbabel.org), or if you have Google Earth installed, you can find a copy of it there.
	GPX2CSV.xsl - REQUIRED. A XML Stylesheet Transform (XSLT) that converts GPX files to CSV files, which can then be imported.  This sheet also adds QC fields for pre-analysis verification.
	GPX2HTA.xsl - REQUIRED. A XML Stylesheet Transform (XSLT) that converts the GPX data read from the Garmin and builds a HTML table summarizing what is available.
		The table is then added to the HTA so the user may select which tracks to import.  This XSLT also applies some logic to highlight viable tracks.
	libexpat.dll - REQUIRED. A third-party library used by gpsbabel.
	ReadMe.txt - This document.

Format:
	HTA (HTML Application) - HTAs use the power of HTML and CSS as the user interface on top of VBScript or Javascript code.  Although the rendering engine is InternetExplorer, HTAs bypass the system's internet
		security restrictions, thus allowing the script code to access and manipulate system resources.

Usage:
	- Double-click on the Garmin.hta file (or a shortcut to the HTA)

Dependencies:
	- None.  The required files are in the same directory as the host

Process:
	- 1. Run Garmin.hta
	- 2. A prompt will appear telling you to connect and activate the Garmin unit
	- 3. The HTA will then pull the data from the Garmin unit and build a GPX file
	- 4. The GPX2HTA stylesheet is applied to the GPX data to build a table of available tracks
	- 5. The table is displayed in the HTA
	- 6. Use the HTA form to enter amplifying information about the player.
	- 6.1  Enter the GPS Unit ID
	- 6.2  Select the player's allegiance (Friend, Hostile)
	- 6.3  Select the player's class (Aircraft, Ground vehicle, Watercraft)
	- 6.4  Enter a brief description of the player
	- 7. Examine the track information extracted from the Garmin unit.  
		You will see the segment number, a checkbox, 
		the number of points in the track, 
		the starting time for the track, 
		the duration of the track, 
		the approximate distance traveled, 
		and the average speed
	- 8. If certain criteria are met, the HTA will automatically select viable tracks for inclusion.  You may override the selection by un-selecting the checkbox or selecting an unselected track.
	- 9. Make sure all of the desired tracks are selected for inclusion and click the "Process Selected Segments" button.
	- 10. The HTA will apply the GPX2CSV stylesheet to the selected tracks in the GPX file and write a CSV file with the date, time, and description provided in the HTA (20101018_161247_AC_C182.csv, for example)
	- 11. You can now examine the file in Excel to look for anomalies
	- 12. Once you are satisfied with the file, you can import it into the analysis database.

Section: GPX2HTA (XSLT stylesheet)

Purpose:
	Process the track and position data in a GPX file to generate a summary table in HTML

Author:
	Tom Elkins (tom.elkins@mercurysolutions.com)

Status:
	- Tested with a wide variety of GPX files, but there may still be some anomalies not yet encountered that will break the process

History:
	20110812 - TAE - Modified to handle the case where saved track logs are present
	20110708 - TAE - Original code

Issues:
	- This stylesheet would not be a good standalone solution.  The HTML it produces is incomplete as it was intended to build only a table, not an entire HTML document.
	- The code necessary to perform calculations on the data is written in VBScript, which restricts the use of this XSLT to Windows machines only.

Notes:
	- The code necessary to perform calculations on the data is embedded in the stylesheet

Format:
	XSLT (XML Stylesheet Transform) - XSLTs are XML files that specify rules for taking data from one file and re-formatting it to transform it into the shape of another file.

Usage:
	- You need an XSLT processor (or write code).  Microsoft provides a free XSLT processor (MSXSL.exe) which can be downloaded from the Microsoft website
	- MSXSL {source file} {stylesheet} -o {output file}
	- MSXSL C:\Test.gpx GPX2HTA.xsl -o Test.htm
	- From Code: 
	
	(start code)
		Set oGPX = CreateObject("Microsoft.XMLDOM")
		Set oXSL = CreateObject("Microsoft.XMLDOM")
		oGPX.Load "C:\Test.gpx"
		oXSL.Load "GPX2HTA.xsl"
		
		sTable = oGPX.TransformNode(oXSL)
		divTable.InnerHTML = sTable
	(end code)

Dependencies:
	- The stylesheet currently works with the GPX 1.1 schema [xmlns:gpx="http://www.topografix.com/GPX/1/1"]

Section: GPX2CSV (XSLT stylesheet)

Purpose:
	Convert a GPX track file to a CSV file and derive some information for checking the quality of the data

Author:
	Tom Elkins (tom.elkins@mercurysolutions.com)

Status:
	- Tested with a wide variety of GPX files, but there may still be some anomalies not yet encountered that will break the process

History:
	20110810 - TAE - Per TRAC #240: Kept altitude as meters, added column for altitude in feet at the end
	20110708 - TAE - Original code

Issues:
	- This stylehsheet would not be a good standalone solution.  It depends on some attributes being added to the GPX file to aid in preparation for analysis; thus, this cannot be used on a native GPX file.
	- The code necessary to perform calculations on the data is written in VBScript, which restricts the use of this XSLT to Windows machines only.

Notes:
	- The code necessary to perform calculations on the data is embedded in the stylesheet

Format:
	XSLT (XML Stylesheet Transform) - XSLTs are XML files that specify rules for taking data from one file and re-formatting it to transform it into the shape of another file.

Usage:
	- You need an XML processor (or write code).  Microsoft provides a free XML processor (MSXSL.exe) which can be downloaded from the Microsoft website
	- MSXSL {source file} {stylesheet} -o {output file}
	- MSXSL C:\Test.gpx GPX2CSV.xsl -o Test.csv
	- From Code: 
	
	(start code)
		Set oGPX = CreateObject("Microsoft.XMLDOM")
		Set oXSL = CreateObject("Microsoft.XMLDOM")
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		
		oGPX.Load "C:\Test.gpx"
		oXSL.Load "GPX2CSV.xsl"
		
		sData = oGPX.TransformNode(oXSL)
		paFiles = Split(sData,"||")
		For iFile = 0 To UBound(paFiles) Step 2
			Set oFile = goFSO.CreateTextFile(paFiles(iFile),True)
			oFile.Write paFiles(iFile+1)
			oFile.Close
		Next
	(end code)

Dependencies:
	- The stylesheet currently works with the GPX 1.1 schema [xmlns:gpx="http://www.topografix.com/GPX/1/1"]
