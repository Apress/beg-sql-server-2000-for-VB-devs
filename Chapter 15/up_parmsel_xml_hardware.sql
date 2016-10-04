CREATE PROCEDURE dbo.up_parmsel_xml_hardware 
	@Hardware_ID INT AS

SELECT Manufacturer_VC, Model_VC, Processor_Speed_VC,
	Memory_VC, HardDrive_VC, Sound_Card_VC,
	Speakers_VC, Video_Card_VC, Monitor_VC, 
	Serial_Number_VC, Lease_Expiration_DT,
	CD_Type_CH

	FROM Hardware_T
	JOIN CD_T ON Hardware_T.CD_ID = CD_T.CD_ID

	WHERE Hardware_ID = @Hardware_ID

	FOR XML AUTO

GO

GRANT EXECUTE ON up_parmsel_xml_hardware TO [Hardware Users]	
