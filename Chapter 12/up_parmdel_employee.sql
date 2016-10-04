CREATE PROCEDURE dbo.up_parmdel_employee
	@Employee_ID		INT AS

-- **************************************************************************
-- Delete the employee
-- **************************************************************************
DELETE FROM Employee_T
	WHERE Employee_ID = @Employee_ID
--
-- Check for errors
--
IF @@ERROR > 0
	BEGIN
	RAISERROR('Delete employee failed',16,1)
	RETURN 99
	END
--
-- Return to the caller
--
RETURN 0

GO

GRANT EXECUTE ON up_parmdel_employee TO [Hardware Users]