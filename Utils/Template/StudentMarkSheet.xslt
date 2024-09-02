<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

	<xsl:template match="/">
		<html>
			<body>
				<h2>Student Report</h2>

				<xsl:apply-templates select="StudentMarkSheet/StudentInformation"/>
				<xsl:apply-templates select="StudentMarkSheet/AcademicDetails"/>
				<xsl:apply-templates select="StudentMarkSheet/OverallResult"/>

			</body>
		</html>
	</xsl:template>

	<xsl:template match="StudentInformation">
		<h3>Student Information:</h3>
		<ul>
			<li>
				<strong>Student ID:</strong>
				<xsl:value-of select="StudentID"/>
			</li>
			<li>
				<strong>Name:</strong>
				<xsl:value-of select="Name"/>
			</li>
			<li>
				<strong>Address:</strong>
				<xsl:value-of select="Address"/>
			</li>
			<li>
				<strong>Grade:</strong>
				<xsl:value-of select="Grade"/>
			</li>
			<li>
				<strong>Contact:</strong>
				<xsl:value-of select="Contact"/>
			</li>
			<li>
				<strong>Father's Name:</strong>
				<xsl:value-of select="FatherName"/>
			</li>
			<li>
				<strong>Mother's Name:</strong>
				<xsl:value-of select="MotherName"/>
			</li>
		</ul>
	</xsl:template>

	<xsl:template match="AcademicDetails">
		<h3>Academic Details:</h3>
		<table border="1">
			<tr bgcolor="#9acd32">
				<th>Subject</th>
				<th>Marks Obtained</th>
				<th>Total Marks</th>
			</tr>
			<xsl:for-each select="SubjectMarks">
				<tr>
					<td>
						<xsl:value-of select="Subject"/>
					</td>
					<td>
						<xsl:value-of select="MarksObtained"/>
					</td>
					<td>
						<xsl:value-of select="TotalMarks"/>
					</td>
				</tr>
			</xsl:for-each>
		</table>
	</xsl:template>

	<xsl:template match="OverallResult">
		<h3>Overall Result:</h3>
		<ul>
			<li>
				<strong>Total Marks Obtained:</strong>
				<xsl:value-of select="TotalMarksObtained"/>
			</li>
			<li>
				<strong>Percentage:</strong>
				<xsl:value-of select="Percentage"/>
			</li>
			<li>
				<strong>Result:</strong>
				<xsl:value-of select="Result"/>
			</li>
		</ul>
	</xsl:template>

</xsl:stylesheet>
