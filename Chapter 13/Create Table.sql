CREATE TABLE Hardware_Notes_T
	(
	Hardware_Notes_ID	INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,
	Hardware_ID		INT 	NOT NULL,
	Hardware_Notes_TX	TEXT 	NOT NULL,
	Last_Update_DT		DATETIME NOT NULL
	)

GO

ALTER TABLE Hardware_Notes_T 
	ADD CONSTRAINT FK_Hardware_Notes_T
	FOREIGN KEY (Hardware_ID)
	REFERENCES Hardware_T(Hardware_ID)
	ON DELETE CASCADE

GO

GRANT SELECT, UPDATE, INSERT, DELETE ON Hardware_Notes_T TO [Hardware Users]

