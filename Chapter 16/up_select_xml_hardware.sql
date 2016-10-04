SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




ALTER PROCEDURE dbo.up_select_xml_hardware AS

SELECT Hardware_ID, Manufacturer_VC, Model_VC, Processor_Speed_VC,
	Memory_VC, HardDrive_VC, Sound_Card_VC, Speakers_VC, 
	Video_Card_VC, Monitor_VC, Serial_Number_VC, 
	DATENAME(MONTH,Lease_Expiration_DT) + ' ' +
	DATENAME(DAY,Lease_Expiration_DT) + ', ' +
	DATENAME(YEAR,Lease_Expiration_DT) AS 'Lease_Expiration_DT',
	CD_Type_CH

	FROM Hardware_T

	JOIN CD_T ON Hardware_T.CD_ID = CD_T.CD_ID

	ORDER BY Manufacturer_VC, Model_VC

	FOR XML AUTO





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

