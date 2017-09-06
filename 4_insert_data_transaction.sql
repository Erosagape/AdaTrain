use adatrain
go

delete from TTRTSleHD;
delete from TTRTSleDT;

declare @dDocDate date;
declare @tDocNo nvarchar(15);
---------------------------
set @dDocDate='2014-06-25';
---------------------------
set @tDocNo=dbo.DFN_TRtGETNewSleDocNo(@dDocDate);
EXEC dbo.SPD_TRxSalesDTIns @tDocNo,@dDocDate,'145874',10,0,'PA','','test'
EXEC dbo.SPD_TRxSalesHDIns @tDocNo,'0',7,'0','C-0003','E-0002','','test'

set @tDocNo=dbo.DFN_TRtGETNewSleDocNo(@dDocDate);
EXEC dbo.SPD_TRxSalesDTIns @tDocNo,@dDocDate,'112853',1,0,'PC','','test'
EXEC dbo.SPD_TRxSalesDTIns @tDocNo,@dDocDate,'145587',2,0,'PA','','test'
EXEC dbo.SPD_TRxSalesHDIns @tDocNo,'0',7,'0','C-0001','E-0002','','test'
-----------------------
set @dDocDate='2014-06-26';
----------------------
set @tDocNo=dbo.DFN_TRtGETNewSleDocNo(@dDocDate);
EXEC dbo.SPD_TRxSalesDTIns @tDocNo,@dDocDate,'827070',2,0,'BG','','test'
EXEC dbo.SPD_TRxSalesHDIns @tDocNo,'0',7,'0','C-0002','E-0001','','test'

set @tDocNo=dbo.DFN_TRtGETNewSleDocNo(@dDocDate);
EXEC dbo.SPD_TRxSalesDTIns @tDocNo,@dDocDate,'827070',1,0,'BG','','test'
EXEC dbo.SPD_TRxSalesDTIns @tDocNo,@dDocDate,'145874',10,0,'PA','','test'
EXEC dbo.SPD_TRxSalesDTIns @tDocNo,@dDocDate,'160492',2,0,'PA','','test'
EXEC dbo.SPD_TRxSalesHDIns @tDocNo,'0',7,'0','C-0001','E-0003','','test'

set @tDocNo=dbo.DFN_TRtGETNewSleDocNo(@dDocDate);
EXEC dbo.SPD_TRxSalesDTIns @tDocNo,@dDocDate,'827070',5,0,'BG','','test'
EXEC dbo.SPD_TRxSalesHDIns @tDocNo,'0',7,'0','C-0003','E-0003','','test'

------------------------
set @dDocDate='2014-07-03';
------------------------

set @tDocNo=dbo.DFN_TRtGETNewSleDocNo(@dDocDate);
EXEC dbo.SPD_TRxSalesDTIns @tDocNo,@dDocDate,'248352',2,0,'TU','','test'
EXEC dbo.SPD_TRxSalesDTIns @tDocNo,@dDocDate,'145874',2,0,'EA','','test'
EXEC dbo.SPD_TRxSalesHDIns @tDocNo,'0',7,'0','C-0002','E-0003','','test'

set @tDocNo=dbo.DFN_TRtGETNewSleDocNo(@dDocDate);
EXEC dbo.SPD_TRxSalesDTIns @tDocNo,@dDocDate,'248352',15,0,'TU','','test'
EXEC dbo.SPD_TRxSalesHDIns @tDocNo,'0',7,'0','C-0002','E-0003','','test'
