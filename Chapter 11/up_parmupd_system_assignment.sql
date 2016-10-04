CREATE PROCEDURE dbo.up_parmupd_system_assignment
	@System_Assignment_ID	INT,
	@Employee_ID		INT,
	@Hardware_ID		INT AS

-- **************************************************************************
-- Begin Transaction
-- **************************************************************************
BEGIN TRANSACTION Update_System_Assignment

-- **************************************************************************
-- Update the system assignment
-- **************************************************************************
UPDATE System_Assignment_T
	SET Employee_ID = @Employee_ID, 
	Hardware_ID = @Hardware_ID, 
	Last_Update_DT = GETDATE()
	WHERE System_Assignment_ID = @System_Assignment_ID
--
-- Check for errors
--
IF @@ERROR > 0
	BEGIN
	RAISERROR('Update of system assignment failed',16,1)
	ROLLBACK TRANSACTION Update_System_Assignment
	RETURN 99
	END

-- **************************************************************************
-- Delete all associated software
-- **************************************************************************
DELETE FROM System_Software_Relationship_T
	WHERE System_Assignment_ID = @System_Assignment_ID
--
-- Check for errors
--
IF @@ERROR > 0
	BEGIN
	RAISERROR('Delete of associated software failed',16,1)
	ROLLBACK TRANSACTION Update_System_Assignment
	RETURN 99
	END

-- **************************************************************************
-- Commit Transaction
-- **************************************************************************
COMMIT TRANSACTION Update_System_Assignment

--
-- Return to the caller
--
RETURN 0

GO

GRANT EXECUTE ON up_parmupd_system_assignment TO [Hardware Users]	


