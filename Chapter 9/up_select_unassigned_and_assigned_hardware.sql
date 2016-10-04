CREATE PROCEDURE dbo.up_select_unassigned_and_assigned_hardware AS

--
-- Create temporary hardware table
--
CREATE TABLE #Tmp_Hardware
	(
	Hardware_ID		INT		NOT NULL,
	Manufacturer_VC		VARCHAR(60)	NOT NULL,
	Serial_Number_VC	VARCHAR(30)	NOT NULL
	)

--
-- Insert all hardware that is not assigned
--
INSERT INTO #Tmp_Hardware
	SELECT DISTINCT Hardware_T.Hardware_ID, 
		Manufacturer_VC + ' ' + Model_VC AS Manufacturer_VC,
		Serial_Number_VC
		FROM Hardware_T 
		JOIN System_Assignment_T ON Hardware_T.Hardware_ID <> 
			System_Assignment_T.Hardware_ID
		WHERE Hardware_T.Hardware_ID NOT IN 
			(SELECT Hardware_ID FROM System_Assignment_T)
--
-- Insert all hardware that is assigned and mark it as not available (NA)
--
INSERT INTO #Tmp_Hardware
	(Hardware_T.hardware_ID,Manufacturer_VC,Serial_Number_VC)
	SELECT Hardware_T.Hardware_ID, 
		'NA - ' + Manufacturer_VC + ' ' + Model_VC AS Manufacturer_VC,
		 Serial_Number_VC
		FROM Hardware_T 
		JOIN System_Assignment_T ON Hardware_T.Hardware_ID = 
			System_Assignment_T.Hardware_ID
--
-- Select all data from temporary hardware table
--
SELECT Hardware_ID, Manufacturer_VC, Serial_Number_VC 
	FROM #Tmp_Hardware
	ORDER BY Manufacturer_VC
--
-- Drop temporary hardware table
--
DROP TABLE #Tmp_Hardware