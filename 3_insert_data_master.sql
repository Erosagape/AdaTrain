use AdaTrain
go

-- Customer Default --
insert into TTRMCst
(FTCstCode,FTCstName,FTCstAddress,FTCstPriceLv,FTCstStatus,FTCstTel,FTCstFax,FCCstARBal,FCCstChqBal,
FTRemark,FDDateIns,FTTimeIns,FTWhoIns,FDDateUpd,FTTimeUpd,FTWhoUpd)
values
('C-0001','บ.สยามอินเตอร์ จำกัด','25/2 ถ.สามเสน','1','0','026544587','026544588',0,0,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('C-0002','บ.เอซี จำกัด','11 บางประกอก','2','0','026654874','026654872',0,0,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('C-0003','นายสมชาย ดีมาก','123 รามคำแหง 65','1','0','024411545','',0,0,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system')
go

--- Sales Person Default --
insert into TTRMSpn
(FTSpnCode,FTSpnName,
FTRemark,FDDateIns,FTTimeIns,FTWhoIns,FDDateUpd,FTTimeUpd,FTWhoUpd)
values
('E-0001','นายจิรายุท สุขใจ','',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('E-0002','น.ส.อุทัย ใจดี','',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('E-0003','นายสาธิต หนองเต่า','',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system')
go

---Product Default---
insert into TTRMPdt
(FTPdtCode,FTPdtBarCode,FTPdtName,FTPdtUnit,FTPdtGroup,FCPriceSale1,FCPriceSale2,FCPriceSale3,FCPriceSale4,
FTRemark,FDDateIns,FTTimeIns,FTWhoIns,FDDateUpd,FTTimeUpd,FTWhoUpd)
values
('145874','145874','มาม่าต้มยำ','EA','001',5,4,0,0,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('145587','145587','มาม่าต้มโคล้ง','EA','001',5,4,0,0,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('248352','248352','ครีมกันแดดสูตรน้ำ เนื้อบีบี สำหรับผิวหน้า เป็นกันแดด 2 in 1','EA','002',125.5,120,0,0,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('112853','112853','หูฟัง Beats รุ่น Studio Over Ear Headphone สี Silver','EA','003',11700,10700,0,0,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('827070','827070','Fineline น้ำตาปรับผ้านุ่ม ไฟน์ไลน์ กสิ่นซากุระ ถุงเติม 550 มล.','EA','002',45,30,0,0,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('160492','160492','วุ้นเส้นผัดพริกโหระพา (เจ) ตรา ชอช้าง','EA','001',220,210,0,0,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system')
go

---Product group Default ---
insert into TTRMPdtGrp
(FTPdtGrpCode,FTPdtGrpName,
FTRemark,FDDateIns,FTTimeIns,FTWhoIns,FDDateUpd,FTTimeUpd,FTWhoUpd)
values
('001','อุปโภค-บริโภค','',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('002','ครีม','',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('003','อิเลคโทรนิค','',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system')
go

---Product unit Default ---
insert into TTRMUnit
(FTUntCode,FTUntName,
FTRemark,FDDateIns,FTTimeIns,FTWhoIns,FDDateUpd,FTTimeUpd,FTWhoUpd)
values
('EA','EACH','ซอง',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('PA','PACK','ห่อ',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('PC','PIECE','ชิ้น',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('BA','BAG','ถุง',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('TU','TUBE','หลอด',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system')
go

---Product Set Default---
insert into TTRMPdtSet
(FTPdtCode,FTPdtUntSet,FCPdtQtySet,
FTRemark,FDDateIns,FTTimeIns,FTWhoIns,FDDateUpd,FTTimeUpd,FTWhoUpd)
values
('145587','PA',10,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system'),
('145874','PA',10,'',convert(date,getdate()),convert(time,getdate()),'system',convert(date,getdate()),convert(time,getdate()),'system')
go