PARAMETERS lcReprocessDate          && FORMATO 'AAAA-MM-DD'
SET TALK OFF
SET ESCAPE ON
SET STEP ON

SET PROCEDURE TO _APPBIPGDIR+'Develop\MySqlLibs\MySqlLib' ADDITIVE

MySqlConnectAt = _AppbiDBSchema
=MySqlConn(MySqlConnectAt)

=Log_Process(SYS(16),'START')

IF EMPTY(lcReprocessDate)
   lcPROC = dtoc(date()-3) 
   ldprocessdate = substr(lcPROC,7,4)+'-'+substr(lcPROC,1,2)+'-'+substr(lcPROC,4,2)
ELSE
   ldprocessdate = lcReprocessDate
ENDIF


xSQL = "SELECT OrderId, OrderItemID, MrktplaceOrderID, ProductID, ShadowOf as ParentSKU, SerialNumber, ProductName, " 
xSQL = xSQL + "  SourceMarket, CompanyName, VendorName, ItemQtyOrdered as QtyOrdered, QtyShipped, rmaQtyReceived as QtyRmaReceived, "
xSQL = xSQL + "  Rma_QtyReturned as QtyRmaReturned, QtyReturnedOrder, "
xSQL = xSQL + "      (AdjustedSitePrice*QtyShipped) as RevenuebyProducts, (AdjustedSitePrice*ItemQtyOrdered ) as RevenuebyProducts_QtyOrdered, "
xSQL = xSQL + "  LineTaxTotal, ShippingCost, ShippingTax, OrderDiscountsTotal, ShippingDiscountsTotal, " 
xSQL = xSQL + "  ((AdjustedSitePrice*QtyShipped)+ShippingCost+ShippingTax+LineTaxTotal-OrderDiscountsTotal-ShippingDiscountsTotal) as GrossSales, " 
xSQL = xSQL + "  PaymentCharged, PaymentRefund, (PaymentCharged-PaymentRefund) as NetSales, "
 
xSQL = xSQL + "  IF (RMAid = 0, (Serialnumber_Cost*QtyShipped) , " 
xSQL = xSQL + "    IF (rmaQtyReceived <> 0 , (Serialnumber_Cost*QtyShipped)-(rmaQtyReceived*Serialnumber_Cost) ,  "
xSQL = xSQL + "              (Serialnumber_Cost*QtyShipped) ) ) as COGS_POs, " 

xSQL = xSQL + "  IF (RMAid = 0, (AverageCost*QtyShipped) , " 
xSQL = xSQL + "    IF (rmaQtyReceived <> 0 , (AverageCost*QtyShipped)-(rmaQtyReceived*AverageCost) , " 
xSQL = xSQL + "              (AverageCost*QtyShipped) ) ) as COGS_Avg, " 

xSQL = xSQL + "  FinalShippingFee, ComissionTotalxItem,  "
xSQL = xSQL + "  (ComissionTotalxItem) / ((AdjustedSitePrice*QtyShipped) - OrderDiscountsTotal) as ComissionPercentValue, "
xSQL = xSQL + "  PaypalFeeTotal, pladjustment, CoOpFee, " 

xSQL = xSQL + "  GrossProfitCash_PO, GrossProfitCash_Avg, " 

xSQL = xSQL + "  SettlementID, Settlement_Date, Settlement_Filter, " 
xSQL = xSQL + "  RmaID, RmaSituation, RmaStatus, " 
xSQL = xSQL + "  RmaReasom, RmaResolution, rmaCategory, rmaSubCategory, " 
xSQL = xSQL + "  Createdate, TimeOfOrder, Shipdate, Paymentdate, " 
xSQL = xSQL + "  POpurchaseid, VendorInvoiceNumber, POCreatedON, POvendorid, Marginal_Vat, ConditionPO, " 
xSQL = xSQL + "  poReceivingStatus, poPaymentStatus, Order_Status, Shipping_Status, Payment_Status, " 
xSQL = xSQL + "  (AdjustedSitePrice*QtyShipped) /  "
xSQL = xSQL + "       (IF (RMAid = 0, (Serialnumber_Cost*QtyShipped) , " 
xSQL = xSQL + "    IF (rmaQtyReceived <> 0 , (Serialnumber_Cost*QtyShipped)-(rmaQtyReceived*Serialnumber_Cost) ,  "
xSQL = xSQL + "              (Serialnumber_Cost*QtyShipped) ) )  ) as MarkUp, "

xSQL = xSQL + "  (PaymentCharged - PaymentRefund - "
xSQL = xSQL + "  (IF (RMAid = 0, (Serialnumber_Cost*QtyShipped) , " 
xSQL = xSQL + "    IF (rmaQtyReceived <> 0 , (Serialnumber_Cost*QtyShipped)-(rmaQtyReceived*Serialnumber_Cost) , " 
xSQL = xSQL + "              (Serialnumber_Cost*QtyShipped) ))) - FinalShippingFee - ComissionTotalxItem -  "
xSQL = xSQL + "          CoOpFee - PaypalFeeTotal -  "
xSQL = xSQL + "          (IF (PaymentCharged - PaymentRefund = 0, 0, LineTaxTotal)) + pladjustment ) / (AdjustedSitePrice*QtyShipped)  as GrossProfitMargin_POs, " 
xSQL = xSQL + "  Prod_Category, ConditionName, BrandName, Url_SCloud, ShipFromWarehouse, ShippingCountry, ShippingRegion "
 
xSQL = xSQL + "  FROM dw_salesorders WHERE Createdate >= '"+ldprocessdate+"'  " 
*******xSQL = xSQL + "      and   (PaymentCharged - PaymentRefund - (AverageCost*QtyShipped) - FinalShippingFee - ComissionTotalxItem - CoOpFee - PaypalFeeTotal - LineTaxTotal + pladjustment) < 0  "
=MySqlSelect(xSQL,'SQCur')

SELE SQCur
IF RECCOUNT() <> 0
  ?'====== Transfer Google Drive GoogleSheet .... '+TTOC(datetime())
  lcFileName = 'H:\.shortcut-targets-by-id\1JqMceiWKLnNFr92NDw3uIleU3Sa1099l\Report\Orders_Negative_Profit_'
  lcFileName = lcFileName +_AppbiSchemaID+'_'+ substr(strtran(ttoc(datetime()),':',''),7,4)+'_'+substr(strtran(ttoc(datetime()),':',''),1,2)+'_'+substr(strtran(ttoc(datetime()),':',''),4,2)+'_'+substr(strtran(ttoc(datetime()),':',''),12)
  lcFileName = lcFileName + '.csv '
  IF FILE(lcFileName)
     DELETE FILE (lcFileName) 
  ENDIF

  SELE SQcur  
  COPY TO (lcFileName) type csv 
ENDIF


*====================================  SENDING MAIL
*******DO  _SendMail


=MySqlDisconn()

SET TALK ON
SET MESSAGE TO 

CLEAR
RETURN




*______________________________________________________________________________________________________
PROCEDURE  _SendMail
?'--------------------------- Send Mail '+time()

lcSUBJECT = 'NegativeProfit '+ldprocessdate 
lcADDTO   = 'salo.ertag@gmail.com'
**lcADDTO   = 'lanr@buyspry.com'

lcSAVEDIR = SYS(5)+CURDIR()
lcDIR = _APPBIPgDir+'Reports\'+_APPBISchemaID+'\' 
SET DEFA TO &lcDIR

lcATTACHMENTS = ''
lnTOPE = ADIR(VFILES,'NegativeProfit.xls')
IF lnTOPE = 0
   RETURN
ENDIF

SET DEFA TO &lcSAVEDIR 
FOR I = 1 TO lnTOPE
    lcATT = lcDIR+'\'+ALLTRIM(VFILES(I,1))
    ?'-----> '+ALLTRIM(lcATT)+'....'+time()
    lcATTACHMENTS = lcATTACHMENTS + lcATT + ', '
ENDFOR

lcBODY = ' '
lcBODY = lcBODY+' <p style="font-family:verdana;font-size:14px;">'
lcBODY = lcBODY + ' <br> <br> '
lcBODY = lcBODY + '____________________________________________________________________________________________________________<br>'  
lcBODY = lcBODY + ' <br> <br> '
lcBODY = lcBODY + ' This is an automatically generated email.  '
lcBODY = lcBODY + ' <br> <br> '
lcBODY = lcBODY + '____________________________________________________________________________________________________________<br>'  
lcBODY = lcBODY + ' <br> <br> '

DO _APPBIPGDIR+'Develop\PG\SndMl_CsFoxySmtp.prg'  with  lcSUBJECT, lcADDTO, lcATTACHMENTS, lcBODY    

RETURN
