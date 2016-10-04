CREATE PROCEDURE dbo.up_parmupd_software 
	@Software_ID		INT,
	@Software_Title_VC 	VARCHAR(30),
	@Software_Category_ID 	INT AS

-- **************************************************************************
-- Delcare variables
-- **************************************************************************
DECLARE @Validated	BIT
--
-- Set default values
--
SET @Validated = 1

-- **************************************************************************
-- Validate data
-- **************************************************************************
--
-- Validate Software ID
--
IF @Software_ID = 0 OR @Software_ID IS NULL
	--
	-- Set the @Validated variable to false
	--
	SET @Validated = 0
--
-- Validate Software Title
--
IF LEN(@Software_Title_VC) = 0 OR @Software_Title_VC IS NULL
	--
	-- Set the @Validated variable to false
	--
	SET @Validated = 0
--
-- Validate Software Category ID
--
IF @Software_Category_ID = 0 OR @Software_Category_ID IS NULL
	--
	-- Set the @Validated variable to false
	--
	SET @Validated = 0
	
-- **************************************************************************
-- Check validation variable
-- **************************************************************************
IF @Validated = 0
	BEGIN
	RAISERROR('Data validation of software failed',16,1)
	RETURN 99
	END

-- **************************************************************************
-- Update the software
-- **************************************************************************
UPDATE Software_T
	SET Software_Name_VC = @Software_Title_VC, 
	Software_Category_ID = @Software_Category_ID, 
	Last_Update_DT = GETDATE()
	WHERE Software_ID = @Software_ID
--
-- Check for errors
--
IF @@ERROR > 0
	BEGIN
	RAISERROR('Update of software failed',16,1)
	RETURN 99
	END

--
-- Return to the caller
--
RETURN 0

GO

GRANT EXECUTE ON up_parmupd_software TO [Hardware Users]	
	