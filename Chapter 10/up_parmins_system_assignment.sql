CREATE PROCEDURE dbo.up_parmins_system_assignment
	@Employee_ID	INT,
	@Hardware_ID	INT,
	@System_ID	INT OUTPUT AS

-- **************************************************************************
-- Insert the system assignment
-- **************************************************************************
INSERT INTO System_Assignment_T
	(Employee_ID, Hardware_ID, Last_Update_DT)
	VALUES(@Employee_ID, @Hardware_ID, GETDATE())
--
-- Check for errors
--
IF @@ERROR > 0
	BEGIN
	RAISERROR('Insert of system assignment failed',16,1)
	RETURN 99
	END
--
-- Get the IDENTITY value inserted to return to the caller
--
SET @System_ID = @@IDENTITY
--
-- Return to the caller
--
RETURN 0


