USE [CRM]
GO
/****** Object:  StoredProcedure [dbo].[proc_Credit_Debit_Debt_Insert]    Script Date: 2021/08/04 3:58:32 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		Shaneil Ramnath
-- Create date: 20 July 2020
-- Description:	Credit/Debit/Bad Debt insert
-- =============================================
ALTER PROCEDURE [dbo].[proc_Credit_Debit_Debt_Insert]
	@Month				INT,
	@Year				INT,
	@CreditSheetName	NVARCHAR(MAX) = 'Credits 15% Vat',
	@CreditSheetColumns NVARCHAR(MAX) = '$A7:Z',
	@DebitSheetName		NVARCHAR(MAX) = 'Debits 15% Vat',
	@DebitSheetColumns	NVARCHAR(MAX) = '$A7:Z',
	@BadSheetName		NVARCHAR(MAX) = 'Bad Debt',
	@BadSheetColumns	NVARCHAR(MAX) = '$A6:Z',
	@FilePath			NVARCHAR(MAX)

	--@CreditSheetName NVARCHAR(MAX) = 'Credits 15% Vat',
	--@CreditSheetColumns NVARCHAR(MAX) = '$A7:Z',
	--@DebitSheetName NVARCHAR(MAX) = 'Debits 15% Vat',
	--@DebitSheetColumns NVARCHAR(MAX) = '$A7:Z',
	--@BadSheetName NVARCHAR(MAX) = 'Bad Debt',
	--@BadSheetColumns NVARCHAR(MAX) = '$A6:Z',
	--@FilePath  NVARCHAR(MAX) = 'c:\Temp\Credits May2020.xlsx'
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;	

	DECLARE @Sql       NVARCHAR(MAX);

	-- Remove all the information from the Temp Table
	TRUNCATE TABLE CRNoteTemp
	TRUNCATE TABLE DRNoteTemp

	IF OBJECT_ID('tempdb..##tmp_BadDebt') IS NOT NULL DROP TABLE ##tmp_BadDebt

	------------------------------------------------ Credit Notes ------------------------------------------------

	SET @Sql =  ' INSERT INTO CRNoteTemp '									+
				' SELECT * '												+
				' FROM OPENROWSET(''Microsoft.ACE.OLEDB.12.0'', '			+
				' ''Excel 12.0; HDR=NO; Database='+ @FilePath + ';'' ,'   +
				' ''SELECT* FROM ['+ @CreditSheetName + @CreditSheetColumns +']'')' 

	EXECUTE sp_executesql @Sql

	-- Remove the NULL columns based on the Customer Code
	DELETE FROM CRNoteTemp WHERE [Cust Code] IS NULL

	-- Insert the Credit Notes
	INSERT INTO CreditNotes (CustomerID, WaybillNumber, AmountEx, CreditNoteRef, Remarks, DebitNoteRef, CRDRReasonCodeID, FinancialYear, FinancialMonth, WaybillID)
		SELECT C.CustomerID, CNT.[W/b No#], CNT.[Excl Amt], CNT.[C/N No#], CNT.Remarks, CNT.[Debit note (if app)], NRC.CRDRNoteReasonCodeID, @Year, @Month, W.WaybillID
		FROM CRNoteTemp  CNT
		LEFT OUTER JOIN CRDRNoteReasonCodes NRC ON CNT.[R/code] = NRC.Code 
		INNER JOIN Customers C ON CNT.[Cust Code] = C.AccountCode
		LEFT OUTER JOIN Waybills W ON CNT.[W/b No#] = W.WaybillNo 
		WHERE C.CustomerID IS NOT NULL



	------------------------------------------------ Debit Notes ------------------------------------------------

	SET @Sql =  ' INSERT INTO DRNoteTemp '									+
				' SELECT * '												+
				' FROM OPENROWSET(''Microsoft.ACE.OLEDB.12.0'', '			+
				' ''Excel 12.0; HDR=NO; Database='+ @FilePath + ';'' ,'   +
				' ''SELECT* FROM ['+ @DebitSheetName + @DebitSheetColumns +']'')' 

	EXECUTE sp_executesql @Sql

	-- Remove the NULL columns based on the Customer Code
	DELETE FROM DRNoteTemp WHERE [Cust Code] IS NULL

	-- Insert the Debit Notes
	INSERT INTO DebitNotes(CustomerID, WaybillNumber, AmountEx, DebitNoteRef, CRDRNoteReasonCodeID, Remarks, FinancialMonth, FinancialYear)
		SELECT C.CustomerID, DNT.[W/b No#], DNT.[Excl Amt], DNT.[D/N No#], CNR.CRDRNoteReasonCodeID, DNT.Remarks, @Month, @Year 
		FROM DRNoteTemp  DNT
		LEFT OUTER JOIN CRDRNoteReasonCodes CNR ON DNT.[R/code] = CNR.CRDRNoteReasonCodeID 
		LEFT OUTER JOIN Customers C ON DNT.[Cust Code] = C.AccountCode
		WHERE C.CustomerID IS NOT NULL



	------------------------------------------------ Bad Debt ------------------------------------------------

	CREATE TABLE ##tmp_BadDebt (
		Remarks			VARCHAR(MAX),
		CustCode		VARCHAR(10),
		CustomerName	VARCHAR(500),
		Amount			FLOAT
	);
	-- Debit Notes
	SET @Sql =  ' INSERT INTO ##tmp_BadDebt '								+
				' SELECT * '												+
				' FROM OPENROWSET(''Microsoft.ACE.OLEDB.12.0'', '			+
				' ''Excel 12.0; HDR=NO; Database='+ @FilePath + ';'' ,'   +
				' ''SELECT * FROM ['+ @BadSheetName + @BadSheetColumns +']'')'

	EXECUTE sp_executesql @Sql

	-- Remove the NULL columns based on the Customer Code
	DELETE FROM ##tmp_BadDebt WHERE CustCode IS NULL

	-- Add the CustomerID column
	ALTER TABLE ##tmp_BadDebt ADD CustomerID INT;

	-- Update the CustomerID Column
	UPDATE ##tmp_BadDebt SET CustomerID = C.CustomerID
	FROM ##tmp_BadDebt T
	LEFT JOIN Customers C ON C.AccountCode = T.CustCode

	-- Insert the Bad Debts into the CreditNotes table
	INSERT INTO CreditNotes (CustomerID, WaybillNumber, AmountEx, CreditNoteRef, Remarks, DebitNoteRef, CRDRReasonCodeID, FinancialYear, FinancialMonth, WaybillID)
		SELECT CustomerID, NULL, Amount, NULL, 'Bad Debt', NULL, 23, @Year, @Month, NULL
		FROM ##tmp_BadDebt

	-- Remove the Temp Table
	DROP TABLE ##tmp_BadDebt
END
