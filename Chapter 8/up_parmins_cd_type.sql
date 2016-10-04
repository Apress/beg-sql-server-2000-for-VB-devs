CREATE PROCEDURE dbo.up_parmins_cd_type 
	@CD_Type CHAR(4) AS

INSERT INTO CD_T
	(CD_Type_CH, Last_Update_DT)
	VALUES(@CD_Type, GETDATE())

RETURN @@IDENTITY