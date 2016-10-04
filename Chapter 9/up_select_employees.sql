CREATE PROCEDURE dbo.up_select_employees AS

--
--	Select data from the Employee_T table
--
SELECT Employee_T.Employee_ID, Employee_T.First_Name_VC, Employee_T.Last_Name_VC, 
	Employee_T.Phone_Number_VC,
--
--	Select data from the Location_T table
--
	Location_T.Location_ID, Location_T.Location_Name_VC,
--
--	Select data from the System_Assignment_T table
--
	System_Assignment_T.System_Assignment_ID
--
--	From
--
	FROM Employee_T
--
--	Join the Location_T table
--
	LEFT OUTER JOIN Location_T ON Employee_T.Location_ID = 
		Location_T.Location_ID
--
--	Join the System_Assignment_T table
--
	LEFT OUTER JOIN System_Assignment_T ON Employee_T.Employee_ID = 
		System_Assignment_T.Employee_ID
--
--	Sort the results
--
	ORDER BY Last_Name_VC, First_Name_VC