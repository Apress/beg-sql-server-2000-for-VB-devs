CREATE PROCEDURE dbo.up_parmdel_hardware
	@Hardware_ID		INT AS

-- **************************************************************************
-- Delete the hardware
-- **************************************************************************
DELETE FROM Hardware_T
	WHERE Hardware_ID = @Hardware_ID
--
-- Check for errors
--
IF @@ERROR > 0
	BEGIN
	RAISERROR('Delete hardware failed',16,1)
	RETURN 99
	END
--
-- Return to the caller
--
RETURN 0

GO

GRANT EXECUTE ON up_parmdel_hardware TO [Hardware Users]		