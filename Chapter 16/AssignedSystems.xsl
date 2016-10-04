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
				<TH ColSpan="2">                                
					<CENTER>Assigned Systems</CENTER>
				</TH>
			</TR>
			<xsl:for-each select="Systems/Employee_T">
				<TR>
					<TD Nowrap="" VAlign="Top" Class="NormalText">
						<xsl:value-of select="@First_Name_VC"/>
						<xsl:value-of select="@Last_Name_VC"/> 
					</TD>
					<TD>   
						<LABEL Class="BlueText" 
							onMouseOver="this.className='RedText'" 
							onMouseOut="this.className='BlueText'" 
							onClick="HideShow()">
							
						<xsl:for-each select="Hardware_T">
							<xsl:value-of select="@Manufacturer_VC"/>
							<xsl:value-of select="@Model_VC"/>, 
							<xsl:value-of select="@Processor_Speed_VC"/>
							with
							<xsl:value-of select="@Memory_VC"/> of memory
						</xsl:for-each>
						
						</LABEL>
						
     					<SPAN Class="NormalText" Style="display:none;">
     						
							<xsl:for-each select="Hardware_T">
							
	     						<DIV Style="margin-left:25px">	
									<xsl:value-of select="@HardDrive_VC"/> 
									hardrive and 
									<xsl:for-each select="CD_T">
										<xsl:value-of select="@CD_Type_CH"/>
     								</xsl:for-each>
     							</DIV>
     							
     							<DIV Style="margin-left:25px">	 
		     						<xsl:value-of select = "@Video_Card_VC"/> 
		     						video card and <xsl:value-of select="@Monitor_VC"/>
     								monitor
     							</DIV>
     							
     							<DIV Style="margin-left:25px">	 
     								<xsl:value-of select="@Sound_Card_VC"/> 
     								sound card and <xsl:value-of select="@Speakers_VC"/> 
     								speakers
     							</DIV>
     							
     							<DIV Style="margin-left:25px">	 
									Serial number  
									<xsl:value-of select="@Serial_Number_VC"/>, 
									Lease expiration 
									<xsl:value-of select="@Lease_Expiration_DT"/>
								</DIV>
								
								<LABEL Class="BlueText" 
									onMouseOver="this.className='RedText'" 
									onMouseOut="this.className='BlueText'" 
									onClick="HideShow()">
								Software</LABEL><BR/>
								
     							<SPAN Class="NormalText" Style="display:none;">
     								
									<DIV Style="margin-left:25px">
										<xsl:for-each select="CD_T/Software_T">
											<xsl:value-of select="@Software_Name_VC"/>, 
										</xsl:for-each>
									</DIV>
									
								</SPAN>
							</xsl:for-each>
						</SPAN>
					</TD>
				</TR>
			</xsl:for-each>
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
				// and Label elements
				var oSpan = document.all.tags("SPAN");
				var oLabel = document.all.tags("LABEL");
				
				// Loop through the label elements collection
				for (i=0; i<oLabel.length; i++)
				{
					// If the className attribute is RedText then
					// process the Span object
					if (oLabel[i].className == "RedText")
					{
						// If the display style is on
						if (oSpan[i].style.display == "")
						{
							// Hide it
							oSpan[i].style.display = "none";
						}
						else
						{
							// Show it
							oSpan[i].style.display = "";
						}
					}
				}
			}
			]]></xsl:comment>
		</SCRIPT>
		</BODY>                                               
		</HTML>   
	</xsl:template>   
</xsl:stylesheet>
