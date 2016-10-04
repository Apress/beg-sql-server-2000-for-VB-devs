<?xml version="1.0" encoding="UTF-8"?>          
<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">   
	<xsl:template match = "/">                                
		<HTML>                                                  
		<HEAD>                                                
		<STYLE>
		.Title 
		{ 
			background-color: #CCCCCC;  
		}
		.NormalText
		{
			font-family: Arail;
			font-size: 10pt;
		}
		</STYLE>    
		</HEAD>                                               
		<BODY>                                                
			<P Width="100%">
				<CENTER Class="Title">Hardware Specifications</CENTER>
			</P>
			<xsl:for-each select="Hardware/Hardware_T">
				<P Class="NormalText">
					<B Style="color: #000080;">
					<xsl:value-of select="@Manufacturer_VC"/>
					<xsl:value-of select="@Model_VC"/> with
					<xsl:value-of select="@Memory_VC"/> of memory.
     				</B>
     			</P>
				<P Class="NormalText">
					Comes with a <xsl:value-of select="@HardDrive_VC"/> 
					hardrive and 
					<xsl:for-each select="CD_T">
						<xsl:value-of select="@CD_Type_CH"/>.
     				</xsl:for-each>
     			</P>
				<P Class="NormalText">
     				<xsl:value-of select = "@Video_Card_VC"/> 
     				video card and a <xsl:value-of select="@Monitor_VC"/>
     				monitor come as standard equipment.
     			</P>
				<P Class="NormalText">
     				<xsl:value-of select="@Sound_Card_VC"/> 
     				sound card and <xsl:value-of select="@Speakers_VC"/> 
     				speakers for true stereo sound.
     			</P>
				<P Class="NormalText">
					Serial number for this model is 
					<xsl:value-of select="@Serial_Number_VC"/>
					and the lease expires on 
					<xsl:value-of select="@Lease_Expiration_DT"/>.
     			</P>
			</xsl:for-each>
		</BODY>                                               
		</HTML>                                                 
	</xsl:template>   
</xsl:stylesheet>
