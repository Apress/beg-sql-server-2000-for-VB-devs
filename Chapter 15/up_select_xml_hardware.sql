CREATE PROCEDURE dbo.up_select_xml_hardware AS

SELECT Manufacturer_VC, Model_VC
	FROM Hardware_T
	FOR XML AUTO

GO

GRANT EXECUTE ON up_select_xml_hardware TO [Hardware Users]
