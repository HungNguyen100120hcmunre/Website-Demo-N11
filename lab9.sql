create table Lop(
	MaLop char(5) not null primary key,
	TenLop nvarchar(20),
	SiSo int
)
create table Sinhvien(
	MaSV char(5) not null primary key,
	Hoten nvarchar(20),
	Ngaysinh date,
	MaLop char(5) constraint fk_malop references lop(malop)
)

create table MonHoc(
	MaMH char(5) not null primary key,
	TenMH nvarchar(20)
)

create table KetQua(
	MaSV char(5) not null,
	MaMH char(5) not null,
	Diemthi float,
	constraint fk_Masv foreign key(MaSV) references sinhvien(MaSV),
	constraint fk_Mamh foreign key(MaMH) references Monhoc(MaMH),
	constraint pk_Masv_Mamh primary key(Masv, mamh)
)

insert lop values
('a','lop a',0),
('b','lop b',0),
('c','lop c', 0)

insert sinhvien values
('01','Thanh Lam','2002-1-1','a'),
('02','Tran Hung','2002-12-1','b'),
('03','minh Tuan','2002-05-12','c')

insert monhoc values
('PPLT','Phuong phap LT'),
('CSDL','Co so du lieu'),
('SQL','He quan tri CSDL'),
('NMJ','Nhap mon Java')

insert KetQua values
('01','PPLT',8),
('01','SQL',7),
('02','PPLT',8),
('01','CSDL',5),
('02','NMJ',5)


--- Câu 1 -----
go
create function diemtb (@msv char(5))
returns float
as
begin
 declare @tb float
 set @tb = (select avg(Diemthi)
 from KetQua
where MaSV=@msv)
 return @tb
end
go
select dbo.diemtb ('01')


----Câu 2 ----
--- Cách 1 ---

create function trbinhlop(@malop char(5))
returns table
as
return
 select s.masv, Hoten, trungbinh=dbo.diemtb(s.MaSV)
 from Sinhvien s join KetQua k on s.MaSV=k.MaSV
 where MaLop=@malop
 group by s.masv, Hoten


 --- Cách 2 ---

go
create function trbinhlop1(@malop char(5))
returns @dsdiemtb table (masv char(5), tensv nvarchar(20), dtb float)
as
begin
 insert @dsdiemtb
 select s.masv, Hoten, trungbinh=dbo.diemtb(s.MaSV)
 from Sinhvien s join KetQua k on s.MaSV=k.MaSV
 where MaLop=@malop
 group by s.masv, Hoten
 return
end
go
select*from trbinhlop1('a')

---- Câu 3 -----

go
create proc ktra @msv char(5)
as
begin
 declare @n int
 set @n=(select count(*) from ketqua where Masv=@msv)
 if @n=0
 print 'sinh vien '+@msv + 'khong thi mon nao'
 else
 print 'sinh vien '+ @msv+ 'thi '+cast(@n as char(2))+ 'mon'
end
go
exec ktra '01'--- Câu 4 ----go
create trigger updatesslop
on sinhvien
for insert
as
begin
 declare @ss int
 set @ss=(select count(*) from sinhvien s
 where malop in(select malop from inserted))
 if @ss>10
 begin
 print 'Lop day'
 rollback tran
 end
 else
 begin
 update lop
 set SiSo=@ss
 where malop in (select malop from inserted)
 end--- Câu 5 ------- Tạo login ---create login user1 with password = '123'
create login user2 with password = '456'
--hoac
sp_addlogin 'user1','123'
--tao user
create user user1 for login user1
create user user2 for login user2
--hoac
sp_adduser 'user1','user1'
go