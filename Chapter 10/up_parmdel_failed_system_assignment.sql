CREATE PROCEDURE dbo.up_parmdel_failed_system_assignment
	@System_ID	INT AS

-- **************************************************************************
-- Delete any existing software
-- **************************************************************************
DELETE FROM System_Software_Relationship_T 
	WHERE System_Assignment_ID = @System_ID

-- **************************************************************************
-- Delete the system assignment
-- **************************************************************************
DELETE FROM System_Assignment_T 
	WHERE System_Assignment_ID = @System_ID