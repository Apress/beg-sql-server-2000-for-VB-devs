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
		<TABLE Border="1">                
			<TR>
				<TH ColSpan="2">Hardware Tracking Employees</TH>
			</TR>            
			<TR>
				<TH>First Name</TH>
				<TH>Last Name</TH>
			</TR>    
			<xsl:for-each select="Employees/Employee_T">
				<TR>                                                   
					<TD><xsl:value-of select="@First_Name_VC"/></TD>      
					<TD><xsl:value-of select="@Last_Name_VC"/></TD>
				</TR>            
			</xsl:for-each>                                      
		</TABLE>                                             
		</BODY>                                               
		</HTML>                                                 
	</xsl:template>                                           
</xsl:stylesheet>