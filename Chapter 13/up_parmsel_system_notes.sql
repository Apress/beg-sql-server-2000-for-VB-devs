CREATE PROCEDURE dbo.up_parmsel_system_notes 
	@Hardware_ID INT AS

-- **************************************************************************
-- Select the notes for a specific system
-- **************************************************************************
SELECT Hardware_Notes_ID, Hardware_Notes_TX
	FROM Hardware_Notes_T
	WHERE Hardware_ID = @Hardware_ID
--
-- Check for errors
--
IF @@ERROR > 0
	BEGIN
	RAISERROR('Select system notes failed',16,1)
	RETURN 99
	END
--
-- Return to the caller
--
RETURN 0

GO

GRANT EXECUTE ON up_parmsel_system_notes TO [Hardware Users]	

