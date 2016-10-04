CREATE PROCEDURE dbo.up_parmupd_cd_type
	@CD_ID INT, 
	@CD_Type CHAR(4), 
	@Return_Code INT OUTPUT AS

UPDATE CD_T
	SET CD_Type_CH = @CD_Type, 
	Last_Update_DT = GETDATE()
	WHERE CD_ID = @CD_ID

SET @Return_Code = @@ERROR