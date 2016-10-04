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
				<TH ColSpan="3">                                
					<CENTER>Software Categories and Titles</CENTER>
				</TH>
			</TR>
			<TR Class="NormalText">
				<TD>
					<B>Categories</B>
				</TD>
				<TD>
					<B>Software Titles</B>
				</TD>
			</TR>
			<TR>
				<TD VAlign="Top">
					<SELECT Name="cboCategories" onChange="HideShow()">
					<xsl:for-each select="Software/Software_Category_T">
						<OPTION>
							<xsl:attribute name="Value">
								<xsl:value-of select="@Software_Category_ID"/>
							</xsl:attribute>
							<xsl:value-of select="@Software_Category_VC"/>
						</OPTION>
					</xsl:for-each>
					</SELECT>
				</TD>
				<TD>
					<xsl:for-each select="Software/Software_Category_T">
	     				<SPAN Class="NormalText" Style="display:none;">
		 					<DIV NoWrap="">	
								<xsl:for-each select="Software_T">
									<xsl:value-of select="@Software_Name_VC"/><BR/> 
			     				</xsl:for-each>
			     			</DIV>
			     		</SPAN>
					</xsl:for-each>
				</TD>
				<TD Width="100%"><BR/></TD>
			</TR>
		</TABLE>
		<P>
			<B Class="BlueText" onMouseOver="this.className='RedText'" 
				onMouseOut="this.className='BlueText'" 
				onclick="window.history.back();">Return To Menu</B>
		</P>
		<SCRIPT Language="JavaScript" TYPE="Text/JavaScript">
			<xsl:comment><![CDATA[
			function HideShow()
			{
				// Declare an object for the collection of Span
				// elements and an object for the cboCategories
				// element
				var oSpan = document.all.tags("SPAN");
				var oSelect = document.all.item("cboCategories");
				
				// Loop through the span elements collection
				// and hide all span elements
				for (i=0; i<oSpan.length; i++)
				{
					oSpan[i].style.display = "none";
				}
				
				// Display the correct span element
				oSpan[oSelect.selectedIndex].style.display = "";
			}
			]]></xsl:comment>
		</SCRIPT>
		<SCRIPT Language="JavaScript" TYPE="Text/JavaScript"
			For="window" Event="onload">
			<xsl:comment><![CDATA[
			// Show the correct Span element for the first Option
			// in the cboCategories Select element
			HideShow();
			]]></xsl:comment>
		</SCRIPT>
		</BODY>                                               
		</HTML>   
	</xsl:template>   
</xsl:stylesheet>
