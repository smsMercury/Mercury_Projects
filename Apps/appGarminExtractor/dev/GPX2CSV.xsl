<?xml version="1.0" encoding="UTF-8"?>
<!-- XSLT:		GPX2CSV
	 AUTHOR:	Tom Elkins
	 PURPOSE:	Convert a GPX track file to a CSV file and derive some information for checking the quality of the data
	 DATE:		20110810
	 HISTORY:	20110810 - TAE - Kept altitude as meters, added column for altitude in feet at the end
				20110708 - TAE - Original code
	 CAVEATS:	Several attributes are added to the source GPX file to aid in Mercury analysis processes; so this stylesheet cannot be used on a native GPX file
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:vb="#fx-functions" exclude-result-prefixes="msxsl" xmlns:gpx="http://www.topografix.com/GPX/1/1">
	
	<!-- We are going to create a CSV file, so tell the processor we are writing text -->
	<xsl:output method="text"/>

	<!-- Build the content wrapper before dealing with the track points -->
	<xsl:template match="/">

		<!-- Build a CSV file for each selected track segment in the GPX file -->
		<xsl:for-each select="gpx:gpx/gpx:trk/gpx:trkseg[@import=1]">
		
			<xsl:variable name="UnitID" select="@unit"/>
			<xsl:variable name="Allegiance" select="@iff"/>
			
			<!-- Create a filename as YYYYMMDD_HHMMSS_CC_TTT.csv -->
			<xsl:value-of select="vb:Timestamp2FileName(string(gpx:trkpt[1]/gpx:time))"/><xsl:text>_</xsl:text>
			<xsl:value-of select="@cat"/><xsl:text>_</xsl:text><xsl:value-of select="@type"/><xsl:text>.csv</xsl:text>
			<xsl:text>||RecordID,UnitID,Time,Latitude,Longitude,Altitude,Allegiance,,Delta Time (sec),Delta Lat (m),Delta Lon (m),Distance (m),Speed (MPH),Delta Alt (Ft),Climb (FPS),Alt (Ft)&#xA;</xsl:text>

			<!-- Loop through the track points -->
			<xsl:for-each select="gpx:trkpt">
				<xsl:variable name="RecID" select="position()"/>
				<xsl:value-of select="$RecID"/><xsl:text>,</xsl:text>
				<xsl:value-of select="$UnitID"/><xsl:text>,</xsl:text>
				<xsl:value-of select="vb:Timestamp2MSTime(string(gpx:time))"/><xsl:text>,</xsl:text>
				<xsl:value-of select="@lat"/><xsl:text>,</xsl:text>
				<xsl:value-of select="@lon"/><xsl:text>,</xsl:text>
				<xsl:value-of select="gpx:ele"/><xsl:text>,</xsl:text>
				<xsl:value-of select="$Allegiance"/><xsl:text>,,</xsl:text>
				<xsl:if test="$RecID &gt; 1">
					<xsl:variable name="DeltaT" select="vb:DeltaTime(string(gpx:time),string(../gpx:trkpt[$RecID - 1]/gpx:time))"/>
					<xsl:value-of select="$DeltaT"/><xsl:text>,</xsl:text>
					<xsl:variable name="LastLat" select="../gpx:trkpt[$RecID - 1]/@lat"/>
					<xsl:variable name="LastLon" select="../gpx:trkpt[$RecID - 1]/@lon"/>
					<xsl:value-of select="(@lat - $LastLat) * 60 * 1852"/><xsl:text>,</xsl:text>
					<xsl:value-of select="(@lon - $LastLon) * 60 * 1852"/><xsl:text>,</xsl:text>
					<xsl:variable name="Dist" select="vb:DistanceTraveled(number(@lat),number(@lon),number($LastLat),number($LastLon))"/>
					<xsl:value-of select="$Dist"/><xsl:text>,</xsl:text>
					<xsl:variable name="Spd" select="$Dist div $DeltaT"/>
					<xsl:value-of select="$Spd * 3.28084 div 5280 * 3600"/><xsl:text>,</xsl:text>
					<xsl:variable name="LastAlt" select="../gpx:trkpt[$RecID - 1]/gpx:ele"/>
					<xsl:value-of select="(gpx:ele - $LastAlt) * 3.28084"/><xsl:text>,</xsl:text>
					<xsl:value-of select="(gpx:ele - $LastAlt) * 3.28084 div $DeltaT"/><xsl:text>,</xsl:text>
					<xsl:value-of select="gpx:ele * 3.28084"/>
				</xsl:if>
				<xsl:text>&#xA;</xsl:text>
			</xsl:for-each>
		</xsl:for-each>
	</xsl:template>
	
	<!-- XSL allows us to extend its capabilities using any language we want -->
	<msxsl:script language="VBScript" implements-prefix="vb" xmlns:msxsl="urn:schemas-microsoft-com:xslt">
		<![CDATA[
			Function Timestamp2MSTime(sTimestamp)
				Timestamp2MSTime = CStr(CDate(Split(Replace(Replace(sTimestamp,"T"," "),"Z",""),".")(0)))
			End Function
			
			Function Timestamp2FileName(sTimestamp)
				dtTime = CDate(Timestamp2MSTime(sTimestamp))
				Timestamp2FileName = DatePart("yyyy",dtTime) & Right("0" & DatePart("m",dtTime),2) & Right("0" & DatePart("d",dtTime),2) & "_" & Right("0" & DatePart("h",dtTime),2) & Right("0" & DatePart("n",dtTime),2) & Right("0" & DatePart("s",dtTime),2)
			End Function
			
			Function DeltaTime(sThisTime, sLastTime)
				dtThis = CDbl(CDate(Timestamp2MSTime(sThisTime)))
				dtLast = CDbl(CDate(Timestamp2MSTime(sLastTime)))
				DeltaTime = (dtThis - dtLast) * 86400
			End Function
			
			Function Sine(dAngle)
				dRad = dAngle * 3.141592654 / 180
				Sine = Sin(dRad)
			End Function
			
			Function Cosine(dAngle)
				dRad = dAngle * 3.141592654 / 180
				Cosine = Cos(dRad)
			End Function
			
			Function ArcSine(dVal)
				If Abs(dVal) = 1.0 Then
					dAngle = (90.0 * Sgn(dVal))
				Else
					dAngle = Atn(dVal / Sqr(-dVal * dVal + 1.0)) * 180 / 3.141592654
				End If
				ArcSine = dAngle
			End Function
			
			Function DistanceTraveled(Lat1,Lon1,Lat2,Lon2)
				DistanceTraveled = (2.0 * ArcSine(Sqr((Sine((Lat1-Lat2)/2))^2 + Cosine(Lat1)*Cosine(Lat2)*(Sine((Lon1-Lon2)/2))^2))) * 60 * 1852
			End Function
				
		]]>
	</msxsl:script>
</xsl:stylesheet>