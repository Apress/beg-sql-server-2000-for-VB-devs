CREATE PROCEDURE dbo.up_parmins_system_software
	@System_ID	INT,
	@Software_ID	INT AS

-- **************************************************************************
-- Insert the system software
-- **************************************************************************
INSERT INTO System_Software_Relationship_T
	(System_Assignment_ID, Software_ID)
	VALUES(@System_ID, @Software_ID)
--
-- Check for errors
--
IF @@ERROR > 0
	BEGIN
	RAISERROR('Insert of system software failed',16,1)
	RETURN 99
	END
--
-- Return to the caller
--
RETURN 0
