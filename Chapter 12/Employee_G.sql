CREATE TRIGGER Employee_G ON Employee_T
	AFTER DELETE AS

-- **************************************************************************
-- Declare variables
-- **************************************************************************
DECLARE @Employee_ID 	INT

-- **************************************************************************
-- Select the Employee_ID to get a count of employees at this location
-- **************************************************************************
SELECT @Employee_ID = Employee_ID
	FROM Employee_T
	WHERE Location_ID = (SELECT Location_ID FROM Deleted)
--
-- Check the row count, if this was the last employee at this location
-- then delete the location
--
IF @@ROWCOUNT = 0
	--
	-- Delete the location
	--
	DELETE FROM Location_T
		WHERE Location_ID = (SELECT Location_ID FROM Deleted)
