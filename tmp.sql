use adatrain
go
---- Insert details ---
insert into TTRTSleSlipDT
(FTSleDDocNo,FNSleDSeq,FDSleDDocDate,FTSleDPdtCode,FTSleDPdtName,FTSleDPdtUnit,
FCSleDQty,FCSleDPrice,FCSleDDisc,FCSleDAmt)
values
('DO140625-00001',1,'2014-06-25','145874','','',10,0,0,0),
('DO140625-00002',1,'2014-06-25','112853','','',1,0,0,0),
('DO140625-00002',2,'2014-06-25','145587','','',2,0,0,0),
('DO140626-00001',1,'2014-06-26','827070','','',2,0,0,0),
('DO140626-00002',1,'2014-06-26','827070','','',1,0,0,0),
('DO140626-00002',2,'2014-06-26','145874','','',10,0,0,0),
('DO140626-00002',3,'2014-06-26','160492','','',2,0,0,0),
('DO140626-00003',1,'2014-06-26','827070','','',5,0,0,0),
('DO140703-00001',1,'2014-07-03','248352','','',2,0,0,0),
('DO140703-00001',2,'2014-07-03','145874','','',2,0,0,0),
('DO140703-00002',1,'2014-07-03','248352','','',15,0,0,0)
go
---calculate amount ---
update d
set d.FTSleDPdtName =p.FTPdtName,
d.FTSleDPdtUnit =p.FTPdtUnit,
d.FCSleDPrice=p.FCPriceSale1,
d.FCSleDAmt=d.FCSleDQty*p.FCPriceSale1
from TTRTSleSlipDT d inner join TTRMPdt p
on d.FTSleDPdtCode=p.FTPdtCode
go
--- insert header ---
insert into TTRTSleSlipHD
select d.FTSleDDocNo,d.FDSleDDocDate,c.FTCstCode,c.FTCstName,'0',7,'1',
sum(d.FCSleDQty*d.FCSleDPrice),sum(d.FCSleDQty*d.FCSleDDisc),
sum(d.FCSleDAmt),
(sum(d.FCSleDQty*d.FCSleDPrice)-sum(d.FCSleDQty*d.FCSleDDisc))*0.07,
(sum(d.FCSleDQty*d.FCSleDPrice)-sum(d.FCSleDQty*d.FCSleDDisc))*1.07
,s.FTSpnCode,s.FTSpnName
from TTRTSleSlipDT d,TTRMCst c,TTRMSlePsn s
where c.FTCstCode='C-0003' and s.FTSpnCode='E-0002' and d.FTSleDDocNo='DO140625-00001'
group by d.FTSleDDocNo,d.FDSleDDocDate,c.FTCstCode,c.FTCstName,s.FTSpnCode,s.FTSpnName
go
insert into TTRTSleSlipHD
select d.FTSleDDocNo,d.FDSleDDocDate,c.FTCstCode,c.FTCstName,'0',7,'1',
sum(d.FCSleDQty*d.FCSleDPrice),sum(d.FCSleDQty*d.FCSleDDisc),
sum(d.FCSleDAmt),
(sum(d.FCSleDQty*d.FCSleDPrice)-sum(d.FCSleDQty*d.FCSleDDisc))*0.07,
(sum(d.FCSleDQty*d.FCSleDPrice)-sum(d.FCSleDQty*d.FCSleDDisc))*1.07
,s.FTSpnCode,s.FTSpnName
from TTRTSleSlipDT d,TTRMCst c,TTRMSlePsn s
where c.FTCstCode='C-0001' and s.FTSpnCode='E-0002' and d.FTSleDDocNo='DO140625-00002'
group by d.FTSleDDocNo,d.FDSleDDocDate,c.FTCstCode,c.FTCstName,s.FTSpnCode,s.FTSpnName
go
insert into TTRTSleSlipHD
select d.FTSleDDocNo,d.FDSleDDocDate,c.FTCstCode,c.FTCstName,'0',7,'1',
sum(d.FCSleDQty*d.FCSleDPrice),sum(d.FCSleDQty*d.FCSleDDisc),
sum(d.FCSleDAmt),
(sum(d.FCSleDQty*d.FCSleDPrice)-sum(d.FCSleDQty*d.FCSleDDisc))*0.07,
(sum(d.FCSleDQty*d.FCSleDPrice)-sum(d.FCSleDQty*d.FCSleDDisc))*1.07
,s.FTSpnCode,s.FTSpnName
from TTRTSleSlipDT d,TTRMCst c,TTRMSlePsn s
where c.FTCstCode='C-0002' and s.FTSpnCode='E-0001' and d.FTSleDDocNo='DO140626-00001'
group by d.FTSleDDocNo,d.FDSleDDocDate,c.FTCstCode,c.FTCstName,s.FTSpnCode,s.FTSpnName
go
insert into TTRTSleSlipHD
select d.FTSleDDocNo,d.FDSleDDocDate,c.FTCstCode,c.FTCstName,'0',7,'1',
sum(d.FCSleDQty*d.FCSleDPrice),sum(d.FCSleDQty*d.FCSleDDisc),
sum(d.FCSleDAmt),
(sum(d.FCSleDQty*d.FCSleDPrice)-sum(d.FCSleDQty*d.FCSleDDisc))*0.07,
(sum(d.FCSleDQty*d.FCSleDPrice)-sum(d.FCSleDQty*d.FCSleDDisc))*1.07
,s.FTSpnCode,s.FTSpnName
from TTRTSleSlipDT d,TTRMCst c,TTRMSlePsn s
where c.FTCstCode='C-0001' and s.FTSpnCode='E-0003' and d.FTSleDDocNo='DO140626-00002'
group by d.FTSleDDocNo,d.FDSleDDocDate,c.FTCstCode,c.FTCstName,s.FTSpnCode,s.FTSpnName
go
insert into TTRTSleSlipHD
select d.FTSleDDocNo,d.FDSleDDocDate,c.FTCstCode,c.FTCstName,'0',7,'1',
sum(d.FCSleDQty*d.FCSleDPrice),sum(d.FCSleDQty*d.FCSleDDisc),
sum(d.FCSleDAmt),
(sum(d.FCSleDQty*d.FCSleDPrice)-sum(d.FCSleDQty*d.FCSleDDisc))*0.07,
(sum(d.FCSleDQty*d.FCSleDPrice)-sum(d.FCSleDQty*d.FCSleDDisc))*1.07
,s.FTSpnCode,s.FTSpnName
from TTRTSleSlipDT d,TTRMCst c,TTRMSlePsn s
where c.FTCstCode='C-0003' and s.FTSpnCode='E-0003' and d.FTSleDDocNo='DO140626-00003'
group by d.FTSleDDocNo,d.FDSleDDocDate,c.FTCstCode,c.FTCstName,s.FTSpnCode,s.FTSpnName
go
insert into TTRTSleSlipHD
select d.FTSleDDocNo,d.FDSleDDocDate,c.FTCstCode,c.FTCstName,'0',7,'1',
sum(d.FCSleDQty*d.FCSleDPrice),sum(d.FCSleDQty*d.FCSleDDisc),
sum(d.FCSleDAmt),
(sum(d.FCSleDQty*d.FCSleDPrice)-sum(d.FCSleDQty*d.FCSleDDisc))*0.07,
(sum(d.FCSleDQty*d.FCSleDPrice)-sum(d.FCSleDQty*d.FCSleDDisc))*1.07
,s.FTSpnCode,s.FTSpnName
from TTRTSleSlipDT d,TTRMCst c,TTRMSlePsn s
where c.FTCstCode='C-0002' and s.FTSpnCode='E-0003' and d.FTSleDDocNo='DO140703-00001'
group by d.FTSleDDocNo,d.FDSleDDocDate,c.FTCstCode,c.FTCstName,s.FTSpnCode,s.FTSpnName
go
insert into TTRTSleSlipHD
select d.FTSleDDocNo,d.FDSleDDocDate,c.FTCstCode,c.FTCstName,'0',7,'1',
sum(d.FCSleDQty*d.FCSleDPrice),sum(d.FCSleDQty*d.FCSleDDisc),
sum(d.FCSleDAmt),
(sum(d.FCSleDQty*d.FCSleDPrice)-sum(d.FCSleDQty*d.FCSleDDisc))*0.07,
(sum(d.FCSleDQty*d.FCSleDPrice)-sum(d.FCSleDQty*d.FCSleDDisc))*1.07
,s.FTSpnCode,s.FTSpnName
from TTRTSleSlipDT d,TTRMCst c,TTRMSlePsn s
where c.FTCstCode='C-0002' and s.FTSpnCode='E-0003' and d.FTSleDDocNo='DO140703-00002'
group by d.FTSleDDocNo,d.FDSleDDocDate,c.FTCstCode,c.FTCstName,s.FTSpnCode,s.FTSpnName
go

