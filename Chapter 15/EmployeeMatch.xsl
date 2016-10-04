<?xml version="1.0" encoding="UTF-8"?>          
<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">
	<xsl:template match="/">
		<HTML>
		<HEAD>                                                
		<STYLE>
		TH 
		{ 
			background-color: #CCCCCC;
		}
		</STYLE>    
		</HEAD>                                               
		<BODY>                                                
		<TABLE Border="1">                
			<TR>
				<TH ColSpan="2">Hardware Tracking Employees</TH>
			</TR>            
			<TR>
				<TH>First Name</TH>
				<TH>Last Name</TH>
			</TR>    
			<xsl:for-each select="Employees/Employee_T"
				order-by="@Last_Name_VC; @First_Name_VC">
				<TR>
					<xsl:apply-templates select="@First_Name_VC"/>
					<xsl:apply-templates select="@Last_Name_VC"/>
				</TR>
			</xsl:for-each>
		</TABLE>
		</BODY>
		</HTML>
	</xsl:template>
	
	<xsl:template match="@First_Name_VC">
		<TD><FONT Color="#000080"><xsl:value-of/></FONT></TD>
	</xsl:template>
	<xsl:template match="@Last_Name_VC">
		<TD><B><xsl:value-of/></B></TD>
	</xsl:template>
</xsl:stylesheet>