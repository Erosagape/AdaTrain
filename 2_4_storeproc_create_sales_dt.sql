USE [AdaTrain]
GO

/****** Object:  StoredProcedure [dbo].[sp_create_sales_dt]    Script Date: 01/09/2017 4:50:07 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[SPD-TRxSalesDTUpd]
(
@pFTSleHDDocNo nvarchar(15),
@pFNSleDTSeq int,
@pFTSleDTDocDate date,
@pFTSleDTPdtCode nvarchar(15),
@pFTSleDTPdtName nvarchar(100),
@pFTSleDTPdtUnit nvarchar(10),
@pFCSleDTQty float,
@pFCSleDTDisc float,
@pFCSleDTPrice float,
@pFCSleDTAmt float,
@pFTRemark varchar(100),
@pFTWhoUpd varchar(50)
)
as
begin

update TTRTSleDT set
FDSleDTDocDate=@pFTSleDTDocDate,
FTSleDTPdtCode=@pFTSleDTPdtCode,
FTSleDTPdtName=@pFTSleDTPdtName,
FTSleDTPdtUnit=@pFTSleDTPdtUnit,
FCSleDTQty=@pFCSleDTQty,
FCSleDTPrice=@pFCSleDTPrice,
FCSleDTDisc=@pFCSleDTDisc,
FCSleDTAmt=@pFCSleDTAmt,
FTRemark=@pFTRemark,
FDDateUpd=convert(date,getdate()),
FTTimeUpd=convert(time,getdate()),
FTWhoUpd=@pFTWhoUpd
where FTSleHDDocNo=@pFTSleHDDocNo
and FNSleDTSeq=@pFNSleDTSeq;
end
go

create procedure [dbo].[SPD_TRxSalesDTIns]
(
@pFTSleHDDocNo nvarchar(15),
@pFTSleDTDocDate date,
@pFTSleDTPdtCode nvarchar(15),
@pFCSleDTQty float,
@pFCSleDTDisc float,
@pFTSleDTPdtUnit nvarchar(10),
@pFTRemark varchar(100),
@pFTWhoIns varchar(50)
)
as
begin
declare @nFNSleDTSeq int;
select @nFNSleDTSeq=ISNULL(max(FNSleDTSeq)+1,1) from TTRTSleDT where FTSleHDDocNo=@pFTSleHDDocNo;

insert into TTRTSleDT
(FTSleHDDocNo,FNSleDTSeq,FDSleDTDocDate,FTSleDTPdtCode,FTSleDTPdtName,FTSleDTPdtUnit,
FCSleDTQty,FCSleDTPrice,FCSleDTDisc,FCSleDTAmt,
FTRemark,FDDateIns,FTTimeIns,FTWhoIns,FDDateUpd,FTTimeUpd,FTWhoUpd)
values
(@pFTSleHDDocNo,@nFNSleDTSeq,@pFTSleDTDocDate,@pFTSleDTPdtCode,'',@pFTSleDTPdtUnit,@pFCSleDTQty,0,@pFCSleDTDisc,0,
@pFTRemark,convert(date,GETDATE()),Convert(time,getdate()),@pFTWhoIns,convert(date,GETDATE()),convert(time,getdate()),@pFTWhoIns);

end


GO


