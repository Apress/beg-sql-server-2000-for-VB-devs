CREATE PROCEDURE dbo.up_select_xml_software AS

SELECT Software_Category_T.Software_Category_ID, Software_Category_VC,
	Software_Name_VC

	FROM Software_Category_T

	JOIN Software_T ON Software_Category_T.Software_Category_ID
		= Software_T.Software_Category_ID

	ORDER BY Software_Category_VC, Software_Name_VC

	FOR XML AUTO

GO

GRANT EXECUTE ON up_select_xml_software TO [Hardware Users]
