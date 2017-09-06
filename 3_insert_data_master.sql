use AdaTrain
go

-- Customer Default --
insert into TTRMCst
(FTCstCode,FTCstName,FTCstAddress,FTCstPriceLv,FTCstStatus,FTCstTel,FTCstFax,FCCstARBal,FCCstChqBal,
FTRemark,FDDateIns,FTTimeIns,FTWhoIns,FDDateUpd,FTTimeUpd,FTWhoUpd)
values
('C-0001','�.�����Թ���� �ӡѴ','25/2 �.����ʹ','1','0','026544587','026544588',0,0,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('C-0002','�.�ͫ� �ӡѴ','11 �ҧ��С͡','2','0','026654874','026654872',0,0,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('C-0003','�������� ���ҡ','123 ������˧ 65','1','0','024411545','',0,0,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system')
go

--- Sales Person Default --
insert into TTRMSpn
(FTSpnCode,FTSpnName,
FTRemark,FDDateIns,FTTimeIns,FTWhoIns,FDDateUpd,FTTimeUpd,FTWhoUpd)
values
('E-0001','��¨����ط �آ�','',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('E-0002','�.�.�ط�� 㨴�','',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('E-0003','����ҸԵ ˹ͧ���','',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system')
go

---Product Default---
insert into TTRMPdt
(FTPdtCode,FTPdtBarCode,FTPdtName,FTPdtUnit,FTPdtGroup,FCPriceSale1,FCPriceSale2,FCPriceSale3,FCPriceSale4,
FTRemark,FDDateIns,FTTimeIns,FTWhoIns,FDDateUpd,FTTimeUpd,FTWhoUpd)
values
('145874','145874','����ҵ����','EA','001',5,4,0,0,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('145587','145587','����ҵ�����','EA','001',5,4,0,0,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('248352','248352','�����ѹᴴ�ٵù�� ���ͺպ� ����Ѻ���˹�� �繡ѹᴴ 2 in 1','EA','002',125.5,120,0,0,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('112853','112853','�ٿѧ Beats ��� Studio Over Ear Headphone �� Silver','EA','003',11700,10700,0,0,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('827070','827070','Fineline ��ӵһ�Ѻ��ҹ��� 俹��Ź� ���蹫ҡ��� �ا��� 550 ��.','EA','002',45,30,0,0,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('160492','160492','�����鹼Ѵ��ԡ���о� (�) ��� �ͪ�ҧ','EA','001',220,210,0,0,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system')
go

---Product group Default ---
insert into TTRMPdtGrp
(FTPdtGrpCode,FTPdtGrpName,
FTRemark,FDDateIns,FTTimeIns,FTWhoIns,FDDateUpd,FTTimeUpd,FTWhoUpd)
values
('001','�ػ���-������','',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('002','����','',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('003','���Ť�ùԤ','',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system')
go

---Product unit Default ---
insert into TTRMUnit
(FTUntCode,FTUntName,
FTRemark,FDDateIns,FTTimeIns,FTWhoIns,FDDateUpd,FTTimeUpd,FTWhoUpd)
values
('EA','EACH','�ͧ',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('PA','PACK','���',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('PC','PIECE','���',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('BA','BAG','�ا',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('TU','TUBE','��ʹ',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system')
go

---Product Set Default---
insert into TTRMPdtSet
(FTPdtCode,FTPdtUntSet,FCPdtQtySet,
FTRemark,FDDateIns,FTTimeIns,FTWhoIns,FDDateUpd,FTTimeUpd,FTWhoUpd)
values
('145587','PA',10,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('145874','PA',10,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system')
go