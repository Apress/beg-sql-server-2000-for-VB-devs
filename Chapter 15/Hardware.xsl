<?xml version="1.0" encoding="UTF-8"?>          
<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">    
	<xsl:template match = "/">                               
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
		<TABLE border="1">                
			<TR>
				<TH ColSpan="2">Hardware</TH>
			</TR>            
			<TR>
				<TH>Manufacturer</TH>
				<TH>Model</TH>
			</TR>    
			<xsl:for-each select="Hardware/Hardware_T">
				<TR>                                                   
					<TD><xsl:value-of 
						select="@Manufacturer_VC"/></TD>      
					<TD><xsl:value-of select="@Model_VC"/></TD>
				</TR>                                                  
			</xsl:for-each>                                      
		</TABLE>                                             
		</BODY>                                               
		</HTML>                                                 
	</xsl:template>                                           
</xsl:stylesheet>