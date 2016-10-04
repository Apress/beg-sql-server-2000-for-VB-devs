CREATE PROCEDURE dbo.up_parmupd_employee
	@Employee_ID		INT,
	@First_Name_VC		VARCHAR(15),
	@Last_Name_VC		VARCHAR(15),
	@Phone_Number_VC	VARCHAR(20),
	@Location_ID		INT,
	@Location_VC		VARCHAR(30) AS

-- **************************************************************************
-- Begin Transaction
-- **************************************************************************
BEGIN TRANSACTION Update_Employee

-- **************************************************************************
-- Update the location
-- **************************************************************************
UPDATE Location_T
	SET Location_Name_VC = @Location_VC,
	Last_Update_DT = GETDATE()
	WHERE Location_ID = @Location_ID
--
-- Check for errors
--
IF @@ERROR > 0
	BEGIN
	RAISERROR('Update of location failed',16,1)
	ROLLBACK TRANSACTION Update_Employee
	RETURN 99
	END

-- **************************************************************************
-- Update the employee
-- **************************************************************************
UPDATE Employee_T
	SET Location_ID = @Location_ID, 
	First_Name_VC = @First_Name_VC, 
	Last_Name_VC = @Last_Name_VC, 
	Phone_Number_VC = @Phone_Number_VC,
	Last_Update_DT = GETDATE()
	WHERE Employee_ID = @Employee_ID
--
-- Check for errors
--
IF @@ERROR > 0
	BEGIN
	RAISERROR('Update of employee failed',16,1)
	ROLLBACK TRANSACTION Update_Employee
	RETURN 99
	END

-- **************************************************************************
-- Commit Transaction
-- **************************************************************************
COMMIT TRANSACTION Update_Employee

--
-- Return to the caller
--
RETURN 0

GO

GRANT EXECUTE ON up_parmupd_employee TO [Hardware Users]