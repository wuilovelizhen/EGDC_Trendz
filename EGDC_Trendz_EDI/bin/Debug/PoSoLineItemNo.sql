Alter Table Impo2 Add PoLineItemNo int
 GO 


Alter VIEW vw_Impo2
AS
SELECT
	TrxNo AS [Trx No],
	LineItemNo AS [Line Item No],
	Amt,
	BalanceQty AS [Balance Qty],
	CountryCode AS [Country Code],
	CurrCode AS [Curr Code],
	CurrRate AS [Curr Rate],
	Description, 
	ItemType AS [Item Type],
	LocalAmt AS [Local Amt],
	PoLineItemNo AS [Po Line Item No],
	ProductCode AS [Product Code],
	Qty,
	Remark,
	UnitPrice AS [Unit Price],
	UomCode AS [Uom Code],
	VendorProductCode AS [Vendor Product Code], 
	Volume,
	Weight
FROM	Impo2

 GO 
 
Alter Table Imso2 Add SoLineItemNo int
 GO 


Alter VIEW vw_Imso2
AS
SELECT
	TrxNo AS [Trx No],
	LineItemNo AS [Line Item No],
	Amt,
	BalanceQty AS [Balance Qty],
	CurrCode AS [Curr Code],
	CurrRate AS [Curr Rate],
	Description,
	ItemType AS [Item Type],
	LocalAmt AS [Local Amt],
	ProductCode AS [Product Code],
	Qty,
	Remark,
	SoLineItemNo AS [So Line Item No],
	UnitPrice AS [Unit Price],
	UomCode AS [Uom Code],
	Volume,
	Weight
FROM Imso2

 GO 




