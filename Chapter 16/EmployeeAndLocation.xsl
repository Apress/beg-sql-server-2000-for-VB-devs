<?xml version="1.0" encoding="UTF-8"?>          
<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">    
	<xsl:template match = "/">                               
		<HTML>                                                  
		<HEAD>                                                
		<TITLE>Beginning SQL Server 2000 for VB Developers</TITLE>
		<LINK Rel="StyleSheet" Type="Text/CSS" Href="http://WSTravel/HardwareTracking/htPageStyles.css"/>
		</HEAD>                                               
		<BODY>                                                
		<TABLE Border="0" Width="100%">
			<TR>
				<TH ColSpan="4">                                
					<CENTER>Employees and Locations</CENTER>
				</TH>
			</TR>
			<TR Class="NormalText">
				<TD>
					<B>First Name</B>
				</TD>
				<TD>
					<B>Last Name</B>
				</TD>
				<TD>
					<B>Phone Number</B>
				</TD>
				<TD>
					<B>Location</B>
				</TD>
			</TR>    
			<xsl:for-each select="Employees/Employee_T">
				<TR onMouseOver="this.style.color='#FF0000'" 
					onMouseOut="this.style.color='#000000'">
					
					<xsl:choose>
						<xsl:when expr="even(this)">
							<xsl:attribute name="Class">EvenRow</xsl:attribute>
						</xsl:when>
						<xsl:otherwise>
							<xsl:attribute name="Class">OddRow</xsl:attribute>
						</xsl:otherwise>
					</xsl:choose>
					
					<TD>
						<xsl:value-of select="@First_Name_VC"/>
					</TD>      
					<TD>
						<xsl:value-of select="@Last_Name_VC"/>
					</TD>
					<TD>
						<xsl:value-of select="@Phone_Number_VC"/>
					</TD>
					
					<xsl:for-each select="Location_T">
						<TD>
							<xsl:value-of select="@Location_Name_VC"/>
						</TD>
     				</xsl:for-each>
				</TR>            
			</xsl:for-each>                                      
		</TABLE>                                             
		<P>
			<B Class="BlueText" onMouseOver="this.className='RedText'" 
				onMouseOut="this.className='BlueText'" 
				onclick="window.history.back();">Return To Menu</B>
		</P>
		</BODY>                                               
		</HTML>                                                 
	</xsl:template>                                           
	<xsl:script language="JavaScript">
	<![CDATA[
		function even(oRow) 
		{
			return ChildNumber(oRow)%2 == 0;
		}
	]]>
	</xsl:script>
</xsl:stylesheet>