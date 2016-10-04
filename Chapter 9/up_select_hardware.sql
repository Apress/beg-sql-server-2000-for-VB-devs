CREATE PROCEDURE dbo.up_select_hardware AS

--
--	Select data from the Hardware_T table
--
SELECT Hardware_ID,Manufacturer_VC,Model_VC, Processor_Speed_VC,
	Memory_VC, HardDrive_VC, Sound_Card_VC, Speakers_VC,
	Video_Card_VC, Monitor_VC, Serial_Number_VC, Lease_Expiration_DT,
--
--	Select data from the CD_T table which is aliased as CD_Table
--
	CD_Table.CD_ID, CD_Type_CH AS CD_DRIVE
--
--	From
--
	FROM Hardware_T
--
--	Join the CD_T table and alias it as CD_Table
--
	JOIN CD_T AS CD_Table ON Hardware_T.CD_ID = CD_Table.CD_ID
--
--	Sort the results
--
	ORDER BY Manufacturer_VC, Model_VC