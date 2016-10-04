CREATE PROCEDURE dbo.up_parmins_software 
	@Software_Title_VC VARCHAR(30),
	@Software_Category_ID INT AS

INSERT INTO Software_T
	(Software_Name_VC, Software_Category_ID, Last_Update_DT)
	VALUES(@Software_Title_VC, @Software_Category_ID, GETDATE())