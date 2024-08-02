<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:vb="#fx-functions" exclude-result-prefixes="msxsl" xmlns:gpx="http://www.topografix.com/GPX/1/1">
	
	<!-- We are going to create HTML content, so tell the processor we are writing HTML -->
	<xsl:output method="html" version="4.0" indent="yes"/>

	<!-- Build the content wrapper before dealing with the track points -->
	<xsl:template match="/">
		<span id="spnGPXNumSeg" style="display:none;"><xsl:value-of select="count(gpx:gpx/gpx:trk/gpx:trkseg)"/></span>
		<span id="spnGPXNorth" style="display:none;"><xsl:value-of select="gpx:gpx/gpx:metadata/gpx:bounds/@maxlat"/></span>
		<span id="spnGPXSouth" style="display:none;"><xsl:value-of select="gpx:gpx/gpx:metadata/gpx:bounds/@minlat"/></span>
		<span id="spnGPXWest" style="display:none;"><xsl:value-of select="gpx:gpx/gpx:metadata/gpx:bounds/@minlon"/></span>
		<span id="spnGPXEast" style="display:none;"><xsl:value-of select="gpx:gpx/gpx:metadata/gpx:bounds/@maxlon"/></span>
		<table border="1" cellpadding="0" style="width:100%;font-size:11pt;">
			<tr><th>Segment</th><th>Include</th><th>Track Points</th><th>Starting Time</th><th>Duration</th><th>Approx. Distance</th><th>Average Speed</th></tr>

			<!-- Build a row for each track segment in the GPX file -->
			<xsl:for-each select="gpx:gpx/gpx:trk/gpx:trkseg">
				<!-- Attempt to ignore segments that have no time -->
				<xsl:if test="count(gpx:trkpt/gpx:time) &gt; 0">
					<xsl:variable name="segDist" select="vb:HowFar(gpx:trkpt)"/>
					<xsl:variable name="segDur" select="vb:HowLong(string(gpx:trkpt[1]/gpx:time),string(gpx:trkpt[last()]/gpx:time))"/>
					<xsl:variable name="segSpd" select="format-number(($segDist div $segDur) * 1852 * 3.28084 * 60 div 5280,'#0.00')"/>
					<tr>
						<!-- Color the row based on the distance -->
						<xsl:choose>
							<xsl:when test="$segDist &gt; 5">
								<xsl:attribute name="style">background-color:#90EE90;</xsl:attribute>
							</xsl:when>
							<xsl:when test="($segSpd) &gt; 2">
								<xsl:attribute name="style">background-color:#F0E68C;</xsl:attribute>
							</xsl:when>
							<xsl:otherwise>
								<xsl:attribute name="style">background-color:#FA8072;</xsl:attribute>
							</xsl:otherwise>
						</xsl:choose>
						
						<!-- Display the segment # -->
						<td style="text-align:center">
							<xsl:value-of select="position()"/>
						</td>
						<td style="text-align:center">
							<input type="checkbox">
								<!-- <xsl:attribute name="id">chkUseSeg<xsl:number count="/gpx:gpx/gpx:trk/gpx:trkseg[count(gpx:trkpt/gpx:time) &gt; 0]"/></xsl:attribute> -->
								<xsl:attribute name="id">chkUseSeg<xsl:value-of select="position()"/></xsl:attribute>
								<xsl:if test="$segDist &gt;= 1"><xsl:if test="$segSpd &gt;= 2"><xsl:attribute name="checked"/></xsl:if></xsl:if>
							</input>
						</td>
						<!-- Display the # pts -->
						<td style="text-align:center">
							<xsl:value-of select="count(gpx:trkpt)"/>
						</td>
						<!-- Display the start date/time -->
						<td style="text-align:center">
							<xsl:value-of select="vb:FormatDateTime(string(gpx:trkpt[1]/gpx:time))"/>
						</td>
						<!-- Display the duration -->
						<td style="text-align:center">
							<xsl:value-of select="$segDur"/> min
						</td>
						<!-- Display the distance -->
						<td style="text-align:center">
							<xsl:value-of select="$segDist"/> NM
						</td>
						<!-- Display the average speed -->
						<td style="text-align:center">
							<xsl:value-of select="format-number(($segDist div $segDur) * 1852 * 3.28084 * 60 div 5280,'#0.00')"/> MPH
						</td>
					</tr>
				</xsl:if>
			</xsl:for-each>
		</table>
	</xsl:template>

	<!-- XSL allows us to extend its capabilities using any language we want -->
	<msxsl:script language="VBScript" implements-prefix="vb" xmlns:msxsl="urn:schemas-microsoft-com:xslt">
		<![CDATA[
			Function FormatDateTime(sValue)
				'
				'	Removes the "T" separator from the time stamp and replaces it with a blank
				FormatDateTime = Replace(sValue,"T"," ")
			End Function
			
			Function HowLong(sStart, sEnd)
				'
				'	Returns the number of minutes elapsed between the two time values
				pdStart = CDbl(CDate(Replace(FormatDateTime(sStart),"Z","")))
				pdEnd = CDbl(CDate(Replace(FormatDateTime(sEnd),"Z","")))
				HowLong = Round((pdEnd - pdStart) * 86400 / 60,2)
			End Function
			
			Function HowFar(oNodeset)
				'
				'	Does a rough calculation of cumulative Nautical Miles traveled between points in the provided node set
				pnDistance = 0
				pdLastLat = 0
				pdLastLon = 0
				pbUseLast = False
				For Each oNode In oNodeset
					If pbUseLast Then
						pnDistance = pnDistance + sqr(((oNode.GetAttribute("lat") - pdLastLat)*60)^2 + ((oNode.GetAttribute("lon") - pdLastLon)*60)^2)
					End If
					pdLastLat = CDbl(oNode.GetAttribute("lat"))
					pdLastLon = CDbl(oNode.GetAttribute("lon"))
					pbUseLast = True
				Next
				HowFar = Round(pnDistance,2)
			End Function
		]]>
	</msxsl:script>
</xsl:stylesheet>