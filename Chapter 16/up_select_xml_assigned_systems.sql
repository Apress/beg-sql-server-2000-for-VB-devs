CREATE PROCEDURE dbo.up_select_xml_assigned_systems AS

--
-- Select columns from the Employee_T table
--
SELECT First_Name_VC, Last_Name_VC, 
--
-- Select columns from the Hardware_T table
--
	Manufacturer_VC, Model_VC, Processor_Speed_VC, Memory_VC,
	HardDrive_VC, Sound_Card_VC, Speakers_VC, Video_Card_VC,
	Monitor_VC, Serial_Number_VC, 
--
-- Format the Lease_Expiration_DT column as month dd, yyyy
--
	DATENAME(MONTH,Lease_Expiration_DT) + ' ' +
	DATENAME(DAY,Lease_Expiration_DT) + ', ' +
	DATENAME(YEAR,Lease_Expiration_DT) AS 'Lease_Expiration_DT',
--
-- Select columns from the CD_T table
--
	CD_Type_CH, 
--
-- Select columns from the Software_T table
--
	Software_Name_VC
--
-- From the System_Assignment_T table
--
	FROM System_Assignment_T
--
-- Join the supporting tables
--
	JOIN Employee_T ON System_Assignment_T.Employee_ID = Employee_T.Employee_ID
	JOIN Hardware_T ON System_Assignment_T.Hardware_ID = Hardware_T.Hardware_ID
	JOIN CD_T ON Hardware_T.CD_ID = CD_T.CD_ID
	JOIN System_Software_Relationship_T ON System_Assignment_T.System_Assignment_ID = System_Software_Relationship_T.System_Assignment_ID
	JOIN Software_T ON System_Software_Relationship_T.Software_ID = Software_T.Software_ID
--
-- Sort the results 
--
	ORDER BY Last_Name_VC, First_Name_VC
--
-- Return the results as XML data
--
	FOR XML AUTO

GO

GRANT EXECUTE ON up_select_xml_assigned_systems TO [Hardware Users]
