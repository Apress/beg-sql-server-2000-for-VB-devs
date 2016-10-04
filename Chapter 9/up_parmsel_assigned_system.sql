CREATE PROCEDURE dbo.up_parmsel_assigned_system
	@Employee_ID INT AS

--
-- Select data from the System_Assignment_T table
--
SELECT System_Assignment_T.Hardware_ID, 
--
--	Select data from the System_Software_Relationship_T table
--
	System_Software_Relationship_T.Software_ID
--
--	From
--	
	FROM System_Assignment_T
--
--	JOIN the System_Software_Relationship_T table
--
	JOIN System_Software_Relationship_T ON System_Assignment_T.System_Assignment_ID = 
		System_Software_Relationship_T.System_Assignment_ID
--
--	Where
--
	WHERE System_Assignment_T.Employee_ID = @Employee_ID
