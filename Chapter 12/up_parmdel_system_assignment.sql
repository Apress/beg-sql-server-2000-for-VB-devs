CREATE PROCEDURE dbo.up_parmdel_system_assignment
	@System_Assignment_ID	INT AS

-- **************************************************************************
-- Delete the system assignment
-- **************************************************************************
DELETE FROM System_Assignment_T
	WHERE System_Assignment_ID = @System_Assignment_ID
--
-- Check for errors
--
IF @@ERROR > 0
	BEGIN
	RAISERROR('Delete system assignment failed',16,1)
	RETURN 99
	END
--
-- Return to the caller
--
RETURN 0

GO

GRANT EXECUTE ON up_parmdel_system_assignment TO [Hardware Users]	


