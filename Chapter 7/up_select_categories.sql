CREATE PROCEDURE dbo.up_select_categories AS

SELECT Software_Category_ID, Software_Category_VC

	FROM Software_Category_T
	
	ORDER BY Software_Category_VC