CREATE PROCEDURE dbo.up_parmdel_software
	@Software_ID INT AS

DELETE FROM Software_T
	WHERE Software_ID = @Software_ID