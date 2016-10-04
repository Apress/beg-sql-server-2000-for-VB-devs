CREATE PROCEDURE dbo.up_parmupd_software
	@Software_ID INT,
	@Software_Title_VC VARCHAR(30),
	@Software_Category_ID INT AS

UPDATE Software_T
	SET Software_Name_VC = @Software_Title_VC,
	Software_Category_ID = @Software_Category_ID,
	Last_Update_DT = GETDATE()
	WHERE Software_ID = @Software_ID
