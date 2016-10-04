CREATE PROCEDURE dbo.up_select_locations AS

SELECT Location_ID, Location_Name_VC
	FROM Location_T
	ORDER BY Location_Name_VC

