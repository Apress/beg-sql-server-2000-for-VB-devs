CREATE PROCEDURE dbo.up_parmdel_software 
	@Software_ID		INT AS

-- **************************************************************************
-- Delete the software
-- **************************************************************************
DELETE FROM Software_T
	WHERE Software_ID = @Software_ID
--
-- Check for errors
--
IF @@ERROR > 0
	BEGIN
	RAISERROR('Delete software failed',16,1)
	RETURN 99
	END
--
-- Return to the caller
--
RETURN 0

GO

GRANT EXECUTE ON up_parmdel_software TO [Hardware Users]	
	