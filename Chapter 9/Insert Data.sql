INSERT INTO Location_T
	(Location_Name_VC, Last_Update_DT)
	VALUES('HOME OFFICE',GETDATE())

INSERT INTO Employee_T
	(Location_ID, First_Name_VC, Last_Name_VC, Phone_Number_VC, Last_Update_DT)
	VALUES(1,'Thearon','Willis','123-456-7890',GETDATE())

INSERT INTO Hardware_T
	(CD_ID, Manufacturer_VC, Model_VC, Processor_Speed_VC, Memory_VC,
		HardDrive_VC, Sound_Card_VC, Speakers_VC, Video_Card_VC,
		Monitor_VC, Serial_Number_VC, Lease_Expiration_DT,
		Last_Update_DT)
	VALUES(3,'Dell','Dimension XPS B800','800 MHZ','256 MB',
		'40 GB Ultra ATA','Turtle Beach Montego II A3D',
		'Altec Lansing ACS-340','32 MB nVIDIA AGP','17" P780 Triniton',
		'123-980A','12/30/00',GETDATE())

INSERT INTO System_Assignment_T
	(Employee_ID, Hardware_ID, Last_Update_DT)
	VALUES(1,1,GETDATE())

INSERT INTO System_Software_Relationship_T
	(System_Assignment_ID, Software_ID)
	VALUES(1,1)

INSERT INTO System_Software_Relationship_T
	(System_Assignment_ID, Software_ID)
	VALUES(1,12)

INSERT INTO System_Software_Relationship_T
	(System_Assignment_ID, Software_ID)
	VALUES(1,18)

INSERT INTO System_Software_Relationship_T
	(System_Assignment_ID, Software_ID)
	VALUES(1,19)

INSERT INTO System_Software_Relationship_T
	(System_Assignment_ID, Software_ID)
	VALUES(1,24)

INSERT INTO System_Software_Relationship_T
	(System_Assignment_ID, Software_ID)
	VALUES(1,25)

INSERT INTO System_Software_Relationship_T
	(System_Assignment_ID, Software_ID)
	VALUES(1,23)


