use AdaTrain
go
CREATE VIEW VTRTSleAll
as
SELECT     dbo.TTRTSleHD.FTSleHDDocNo, dbo.TTRTSleHD.FDSleHDDocDate, dbo.TTRTSleHD.FTSleHDCstCode, dbo.TTRTSleHD.FTSleHDCstName, 
                      dbo.TTRMCst.FTCstAddress, dbo.TTRMCst.FTCstPriceLv, dbo.TTRMCst.FTCstStatus, dbo.TTRMCst.FTCstTel, dbo.TTRMCst.FTCstFax, dbo.TTRMCst.FCCstARBal, 
                      dbo.TTRMCst.FCCstChqBal, dbo.TTRTSleHD.FTSleHDDocType, dbo.TTRTSleHD.FCSleHDVatRate, dbo.TTRTSleHD.FTSleHDVatType, 
                      dbo.TTRTSleHD.FCSleHDDocAmt, dbo.TTRTSleHD.FCSleHDDiscAmt, dbo.TTRTSleHD.FCSleHDBeforeVat, dbo.TTRTSleHD.FCSleHDVatAmt, 
                      dbo.TTRTSleHD.FCSleHDDocTotal, dbo.TTRTSleHD.FTSleHDSpnCode, dbo.TTRTSleHD.FTSleHDSpnName, dbo.TTRTSleDT.FNSleDTSeq, 
                      dbo.TTRTSleDT.FDSleDTDocDate, dbo.TTRTSleDT.FTSleDTPdtCode, dbo.TTRTSleDT.FTSleDTPdtName, dbo.TTRTSleDT.FTSleDTPdtUnit, 
                      ISNULL(dbo.TTRMPdtSet.FCPdtQtySet, 1) AS FTPdtQtySet, dbo.TTRMPdt.FTPdtUnit, dbo.TTRMPdtGrp.FTPdtGrpCode, dbo.TTRMPdtGrp.FTPdtGrpName, 
                      dbo.TTRMPdt.FCPriceSale1, dbo.TTRMPdt.FCPriceSale2, dbo.TTRMPdt.FCPriceSale3, dbo.TTRMPdt.FCPriceSale4, dbo.TTRTSleDT.FCSleDTQty, 
                      dbo.TTRMUnit.FTUntName, dbo.TTRMUnit.FTRemark, dbo.TTRTSleDT.FCSleDTPrice, dbo.TTRTSleDT.FCSleDTDisc, dbo.TTRTSleDT.FCSleDTAmt
FROM         dbo.TTRTSleHD INNER JOIN
                      dbo.TTRMCst ON dbo.TTRTSleHD.FTSleHDCstCode = dbo.TTRMCst.FTCstCode INNER JOIN
                      dbo.TTRMSpn ON dbo.TTRTSleHD.FTSleHDSpnCode = dbo.TTRMSpn.FTSpnCode INNER JOIN
                      dbo.TTRTSleDT ON dbo.TTRTSleHD.FTSleHDDocNo = dbo.TTRTSleDT.FTSleHDDocNo INNER JOIN
                      dbo.TTRMPdt ON dbo.TTRTSleDT.FTSleDTPdtCode = dbo.TTRMPdt.FTPdtCode INNER JOIN
                      dbo.TTRMPdtGrp ON dbo.TTRMPdt.FTPdtGroup = dbo.TTRMPdtGrp.FTPdtGrpCode INNER JOIN
                      dbo.TTRMUnit ON dbo.TTRTSleDT.FTSleDTPdtUnit = dbo.TTRMUnit.FTUntCode LEFT OUTER JOIN
                      dbo.TTRMPdtSet ON dbo.TTRTSleDT.FTSleDTPdtCode = dbo.TTRMPdtSet.FTPdtCode AND dbo.TTRTSleDT.FTSleDTPdtUnit = dbo.TTRMPdtSet.FTPdtUntSet