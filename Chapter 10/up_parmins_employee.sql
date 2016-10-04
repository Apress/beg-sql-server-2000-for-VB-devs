CREATE PROCEDURE dbo.up_parmins_employee
	@First_Name_VC		VARCHAR(15),
	@Last_Name_VC		VARCHAR(15),
	@Phone_Number_VC	VARCHAR(20),
	@Location_VC		VARCHAR(30) AS

-- **************************************************************************
-- Declare variables
-- **************************************************************************
DECLARE @Location_ID		INT

-- **************************************************************************
-- See if the location name already exists by selecting the Location_ID
-- **************************************************************************
SELECT @Location_ID = Location_ID
	FROM Location_T
	WHERE Location_Name_VC = @Location_VC

IF @Location_ID IS NULL
	--
	-- This location does not exists so insert it
	--
	BEGIN
	INSERT INTO Location_T
		(Location_Name_VC, Last_Update_DT)
		VALUES(@Location_VC, GETDATE())
	--
	-- Save the IDENTITY value
	--
	SET @Location_ID = @@IDENTITY
	END

-- **************************************************************************
-- Insert the employee
-- **************************************************************************
INSERT INTO Employee_T
	(Location_ID, First_Name_VC, Last_Name_VC, Phone_Number_VC,
		Last_Update_DT)
	VALUES(@Location_ID, @First_Name_VC, @Last_Name_VC, @Phone_Number_VC,
		GETDATE())
--
-- Return to the caller
--
RETURN 0
