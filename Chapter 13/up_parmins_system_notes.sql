CREATE PROCEDURE dbo.up_parmins_system_notes 
	@Hardware_ID		INT, 
	@System_Notes_TX 	TEXT AS

-- **************************************************************************
-- Insert system notes
-- **************************************************************************
INSERT INTO Hardware_Notes_T
	(Hardware_ID, Hardware_Notes_TX, Last_Update_DT)
	VALUES(@Hardware_ID, @System_Notes_TX, GETDATE())
--
-- Check for errors
--
IF @@ERROR > 0
	BEGIN
	RAISERROR('Insert of system notes failed',16,1)
	RETURN 99
	END
--
-- Return to the caller
--
RETURN 0

GO

GRANT EXECUTE ON up_parmins_system_notes TO [Hardware Users]	
