CREATE PROCEDURE dbo.up_select_xml_employees_and_locations AS

SELECT First_Name_VC, Last_Name_VC, Phone_Number_VC,
	Location_Name_VC

	FROM Employee_T

	JOIN Location_T ON Employee_T.Location_ID = Location_T.Location_ID

	ORDER BY Location_Name_VC, Last_Name_VC, First_Name_VC

	FOR XML AUTO

GO

GRANT EXECUTE ON up_select_xml_employees_and_locations TO [Hardware Users]	





