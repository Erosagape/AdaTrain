USE [AdaTrain]
GO

/****** Object:  UserDefinedFunction [dbo].[GenNewSalesNo]    Script Date: 01/09/2017 1:53:50 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO
--ส่งค่ามาเป็นวันที่ 'yyyy-MM-dd'
create function [dbo].[DFN_TRtGETNewSleDocNo]
(
 @pFTDocDate nvarchar(10) 
)
returns nvarchar(15)
as
begin
declare @tSleHDDocNo nvarchar(15);
declare @tSleHDDocFormat nvarchar(9);
set @tSleHDDocFormat='DO' + replace(substring(@pFTDocDate,3,8),'-','')+ '-';

select @tSleHDDocNo=@tSleHDDocFormat+ FORMAT(ISNULL(max(right(FTSleHDDocNo,5))+1,1),'00000') from TTRTSleDT where FDSleDTDocDate=@pFTDocDate

return @tSleHDDocNo;
end
GO


