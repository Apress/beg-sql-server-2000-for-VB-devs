CREATE PROCEDURE dbo.up_parmsel_read_partial_system_notes 
	@Notes_ID	INT, 
	@Offset 	INT, 
	@Length 	INT AS

-- **************************************************************************
-- Declare variable for the text pointer
-- **************************************************************************
DECLARE @Text_Pointer VARBINARY(16)

-- **************************************************************************
-- Initialize the text pointer
-- **************************************************************************
SELECT @Text_pointer = TEXTPTR(Hardware_Notes_TX)
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
-- Read the text data
-- **************************************************************************
READTEXT Hardware_Notes_T.Hardware_Notes_TX @Text_Pointer @Offset @Length
--
-- Check for errors
--
IF @@ERROR > 0
	BEGIN
	RAISERROR('Reading system notes failed',16,1)
	RETURN 99
	END

--
-- Return to the caller
--
RETURN 0

GO

GRANT EXECUTE ON up_parmsel_read_partial_system_notes TO [Hardware Users]	
