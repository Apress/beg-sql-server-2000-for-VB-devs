CREATE PROCEDURE dbo.up_select_software AS

SELECT Software_ID, Software_Name_VC, Software_T.Software_Category_ID, 
	Software_Category_VC

	FROM Software_T
	
	JOIN Software_Category_T on Software_T.Software_Category_ID = 
		Software_Category_T.Software_Category_ID

	ORDER BY Software_Name_VC