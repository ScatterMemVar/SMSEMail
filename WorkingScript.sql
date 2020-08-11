
DECLARE
	@UserID INT
	, @EMailTypeID INT

SET XACT_ABORT ON;

BEGIN TRY
	BEGIN TRANSACTION;
	BEGIN
		DECLARE @EMailElementID int
		SET @EMailElementID = (SELECT ee.EMail_Element_ID FROM dbo.EMail_Element ee WHERE ee.EMail_Type_PK1 = @EMailTypeID)

		INSERT INTO	dbo.EMail_Log_Entry
		(
		    Date_Sent
			, User_PK1
			, EMail_Element_PK1
			, Available
		)
		VALUES
		(
		    GETDATE() -- Date_Sent - datetime
		    , @UserID -- User_PK1 - int
		    , @EMailElementID -- EMail_Element_PK1 - int
		    , 1 -- Available - bit
		)		
	END;
	COMMIT TRANSACTION;
END TRY
BEGIN CATCH
	BEGIN
		IF
		   (XACT_STATE()) = -1
			BEGIN
				ROLLBACK TRANSACTION;
			END;

		IF
		   (XACT_STATE()) = 1
			BEGIN
				COMMIT TRANSACTION;
			END;
	END;
END CATCH;
