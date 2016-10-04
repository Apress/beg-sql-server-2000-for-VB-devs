CREATE PROCEDURE dbo.up_parmupd_replace_system_notes
	@Notes_ID		INT, 
	@System_Notes_TX	TEXT AS

-- **************************************************************************
-- Declare variable for the text pointer
-- **************************************************************************
DECLARE @Text_Pointer VARBINARY(16)

-- **************************************************************************
-- Initialize the text pointer
-- **************************************************************************
SELECT @Text_Pointer = TEXTPTR(Hardware_Notes_TX)
	FROM Hardware_Notes_T
	WHERE Hardware_Notes_ID = @Notes_ID
-- 
-- Check for a valid text pointer
-- 
IF TEXTVALID('Hardware_Notes_T.Hardware_Notes_TX',@Text_Pointer) <> 1
	BEGIN
	RAISERROR('Text pointer is invalid',16,1)
	RETURN 99
	END

-- **************************************************************************
-- Write the text, replacing any existing text
-- **************************************************************************
WRITETEXT Hardware_Notes_T.Hardware_Notes_TX @Text_Pointer @System_Notes_TX
--
-- Check for errors
--
IF @@ERROR > 0
	BEGIN
	RAISERROR('System notes replacement failed',16,1)
	RETURN 99
	END
--
-- Return to the caller
--
RETURN 0

GO

GRANT EXECUTE ON up_parmupd_replace_system_notes TO [Hardware Users]	
