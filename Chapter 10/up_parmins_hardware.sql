CREATE PROCEDURE dbo.up_parmins_hardware
	@Manufacturer_VC	VARCHAR(30),
	@Model_VC		VARCHAR(30),
	@Processor_Speed_VC	VARCHAR(20),
	@Memory_VC		VARCHAR(10),
	@HardDrive_VC		VARCHAR(15),
	@Sound_Card_VC		VARCHAR(30),
	@Speakers_VC		VARCHAR(30),
	@Video_Card_VC		VARCHAR(30),
	@Monitor_VC		VARCHAR(30),
	@Serial_Number_VC	VARCHAR(30),
	@Lease_Expiration_DT	VARCHAR(22),
	@CD_ID			INT	AS

-- **************************************************************************
-- Insert the hardware
-- **************************************************************************
INSERT INTO Hardware_T
	(Manufacturer_VC, Model_VC, Processor_Speed_VC, Memory_VC,
		HardDrive_VC, Sound_Card_VC, Speakers_VC, Video_Card_VC, 
		Monitor_VC, Serial_Number_VC, Lease_Expiration_DT,
		CD_ID, Last_Update_DT)
	VALUES(@Manufacturer_VC, @Model_VC, @Processor_Speed_VC, @Memory_VC,
		@HardDrive_VC, @Sound_Card_VC, @Speakers_VC, @Video_Card_VC,
		@Monitor_VC, @Serial_Number_VC, @Lease_Expiration_DT, 
		@CD_ID, GETDATE())
--
-- Check for errors
--
IF @@ERROR > 0
	BEGIN
	RAISERROR('Insert of hardware failed',16,1)
	RETURN 99
	END
--
-- Return to the caller
--
RETURN 0
		