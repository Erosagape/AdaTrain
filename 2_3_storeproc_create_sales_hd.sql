USE [AdaTrain]
GO

/****** Object:  StoredProcedure [dbo].[SPD_TRxINSERTSalesHD]    Script Date: 04/09/2017 10:17:56 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

create procedure [dbo].[SPD_TRxSalesDTCalSum]
(
	@pFTSleHDDocNo nvarchar(15),
	@pFTSleHDCstCode nvarchar(15)
)
as
begin
update d
set d.FTSleDTPdtName =p.FTPdtName,
d.FCSleDTPrice=(
	case when c.FTCstPriceLv='4' then p.FCPriceSale4 else (
	case when c.FTCstPriceLv='3' then p.FCPriceSale3 else (
	case when c.FTCstPriceLv='2' then p.FCPriceSale2 else 
	p.FCPriceSale1 
	end)
	end) 
	end)*ISNULL(s.FCPdtQtySet,1),
d.FCSleDTAmt=((d.FCSleDTQty*ISNULL(s.FCPdtQtySet,1))*(
	case when c.FTCstPriceLv='4' then p.FCPriceSale4 else (
	case when c.FTCstPriceLv='3' then p.FCPriceSale3 else (
	case when c.FTCstPriceLv='2' then p.FCPriceSale2 else 
	p.FCPriceSale1 
	end)
	end) 
	end)
)-(d.FCSleDTDisc*(d.FCSleDTQty*ISNULL(s.FCPdtQtySet,1)))
from TTRTSleDT d inner join TTRMPdt p
on d.FTSleDTPdtCode=p.FTPdtCode
left outer join TTRMPdtSet s
on d.FTSleDTPdtCode=s.FTPdtCode
and d.FTSleDTPdtUnit=s.FTPdtUntSet
,TTRMCst c 
where d.FTSleHDDocNo=@pFTSleHDDocNo
and c.FTCstCode=@pFTSleHDCstCode;

end
GO

create procedure [dbo].[SPD_TRxSalesHDIns]
(
	@pFTSleHDDocNo nvarchar(15),
	@pFTSleHDDocType nvarchar(1),
	@pFCSleHDVatRate float,
	@pFTSleHDVatType nvarchar(1),
	@pFTSleHDCstCode nvarchar(15),
	@pFTSleHDSpnCode nvarchar(15),
	@pFTRemark nvarchar(100),
	@pFTWhoIns varchar(50)
)
as
begin
---calculate price from customer and update data related---
EXEC SPD_TRxSalesDTCalSum @pFTSleHDDocNo,@pFTSleHDCstCode;

---insert header and calculate vat by type of document---
declare @cVatCal float;
declare @cVatExtract float;
set @cVatCal=(CASE WHEN @pFTSleHDVatType='0' then (@pFCSleHDVatRate*0.01) else (@pFCSleHDVatRate/(100+@pFCSleHDVatRate)) end);
set @cVatExtract=(case when @pFTSleHDVatType='1' then (100/(100+@pFCSleHDVatRate)) else 1 end);

insert into TTRTSleHD
(FTSleHDDocNo,FDSleHDDocDate,FTSleHDCstCode,FTSleHDCstName,FTSleHDDocType,FCSleHDVatRate,FTSleHDVatType,
FCSleHDDocAmt,FCSleHDDiscAmt,FCSleHDBeforeVat,FCSleHDVatAmt,FCSleHDDocTotal,FTSleHDSpnCode,FTSleHDSpnName,
FTRemark,FDDateIns,FTTimeIns,FTWhoIns,FDDateUpd,FTTimeUpd,FTWhoUpd)
select d.FTSleHDDocNo,d.FDSleDTDocDate,c.FTCstCode,c.FTCstName,@pFTSleHDDocType,@pFCSleHDVatRate,@pFTSleHDVatType,
sum(d.FCSleDTQty*d.FCSleDTPrice),sum(d.FCSleDTQty*d.FCSleDTDisc),
sum(d.FCSleDTAmt)*@cVatExtract,
(sum(d.FCSleDTQty*d.FCSleDTPrice)-sum(d.FCSleDTQty*d.FCSleDTDisc))*@cVatCal,
(sum(d.FCSleDTQty*d.FCSleDTPrice)-sum(d.FCSleDTQty*d.FCSleDTDisc))*(1+((case when @pFTSleHDVatType='0' then @pFCSleHDVatRate*0.01 else 0 end)))
,s.FTSpnCode,s.FTSpnName,@pFTRemark,convert(date,GETDATE()),Convert(Time,GETDATE()),@pFTWhoIns,convert(date,GETDATE()),Convert(Time,GETDATE()),@pFTWhoIns
from TTRTSleDT d,TTRMCst c,TTRMSpn s
where c.FTCstCode=@pFTSleHDCstCode and s.FTSpnCode=@pFTSleHDSpnCode and d.FTSleHDDocNo=@pFTSleHDDocNo
group by d.FTSleHDDocNo,d.FDSleDTDocDate,c.FTCstCode,c.FTCstName,s.FTSpnCode,s.FTSpnName
end
GO


create procedure [dbo].[SPD_TRxSalesHDUpd]
(
	@pFTSleHDDocNo nvarchar(15),
	@pFDSleHDDocDate date,
	@pFTSleHDDocType nvarchar(1),
	@pFCSleHDVatRate float,
	@pFTSleHDVatType nvarchar(1),
	@pFTSleHDCstCode nvarchar(15),
	@pFTSleHDSpnCode nvarchar(15),
	@pFTRemark nvarchar(100),
	@pFTWhoUpd varchar(50)
)
as
begin

EXEC dbo.SPD_TRxSalesDTCalSum @pFTSleHDDocNo,@pFTSleHDCstCode;

declare @cVatCal float;
declare @cVatExtract float;
declare @cSleAmt float;
declare @cSleDisc float;
declare @cSleBeforeVat float;
declare @cSleVat float;
declare @cSleAfterVat float;

set @cVatCal=(CASE WHEN @pFTSleHDVatType='0' then (@pFCSleHDVatRate*0.01) else (@pFCSleHDVatRate/(100+@pFCSleHDVatRate)) end);
set @cVatExtract=(case when @pFTSleHDVatType='1' then (100/(100+@pFCSleHDVatRate)) else 1 end);

select @cSleAmt=sum(d.FCSleDTQty*d.FCSleDTPrice),
@cSleDisc=sum(d.FCSleDTQty*d.FCSleDTDisc),
@cSleBeforeVat=sum(d.FCSleDTAmt)*@cVatExtract,
@cSleVat=(sum(d.FCSleDTQty*d.FCSleDTPrice)-sum(d.FCSleDTQty*d.FCSleDTDisc))*@cVatCal,
@cSleAfterVat=(sum(d.FCSleDTQty*d.FCSleDTPrice)-sum(d.FCSleDTQty*d.FCSleDTDisc))*(1+((case when @pFTSleHDVatType='0' then @pFCSleHDVatRate*0.01 else 0 end)))
from TTRTSleDT d,TTRMCst c
where c.FTCstCode=@pFTSleHDCstCode and d.FTSleHDDocNo=@pFTSleHDDocNo

declare @tCstName nvarchar(100);
select @tCstName=FTCstName from TTRMCst where FTCstCode=@pFTSleHDCstCode;

declare @tSpnName nvarchar(100);
select @tSpnName=FTSpnName from TTRMSpn where FTSpnCode=@pFTSleHDSpnCode;

update TTRTSleH  SET 
FDSleHDDocDate=@pFDSleHDDocDate,
FTSleHDDocType=@pFTSleHDDocType,
FCSleHDVatRate=@pFCSleHDVatRate,
FTSleHDVatType=@pFTSleHDVatType,
FTSleHDCstCode=@pFTSleHDCstCode,
FTSleHDCstName=@tCstName,
FTSleHDSpnCode=@pFTSleHDSpnCode,
FTSleHDSpnName=@tSpnName,
FCSleHDDocAmt=@cSleAmt,
FCSleHDDiscAmt=@cSleDisc,
FCSleHDBeforeVat=@cSleBeforeVat,
FCSleHDVatAmt=@cSleVat,
FCSleHDDocTotal=@cSleAfterVat,
FTRemark=@pFTRemark,
FDDateUpd=convert(date,getdate()),
FTTimeUpd=convert(time,getdate()),
FTWhoUpd=@pFTWhoUpd
where FTSleHDDocNo=@pFTSleHDDocNo;
end
GO



