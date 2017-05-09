use master
go

drop database aplikasiPresensi
go

create database aplikasiPresensi
go

use aplikasiPresensi
go

CREATE table guru (
	id int not null PRIMARY KEY IDENTITY(1,1),
	nama varchar(50) NOT NULL,
	jeniskelamin varchar(1) NOT NULL DEFAULT 'L',
	nip varchar(50),
	status varchar(15) NOT NULL,
	agama varchar(10) NOT NULL,
	tempat varchar(25),
	tanggallahir datetime,
	jabatan varchar(15),
	deleted int default 0
)
go
create table pelajaran(
	id int identity(1,1) primary key,
	nama varchar(50) not null unique,
	deleted int default 0
)
go
create table kelas (
	id int identity(1,1) primary key,
	nama varchar(50)not null unique,
	deleted int default 0
)
go
create table semester(
	id int identity(1,1) primary key,
	nama varchar(50) not null unique,
	deleted int default 0
)
go 
create table jadwal(
	id int identity(1,1) primary key,
	semester int foreign key references semester(id),
	tahun int,
	waktumulai datetime not null,
	waktuselesai datetime not null,
	hari int,
	idkelas int foreign key references kelas(id),
	idpelajaran int foreign key references pelajaran(id),
	idguru int foreign key references guru(id),
	deleted int default 0
)
go
create table rekapjadwal(
	id int identity (1,1) primary key,
	idjadwal int foreign key references jadwal(id),
tanggal datetime default getdate(),
	waktumulai datetime null,
	waktuselesai datetime null,
	keterangan varchar(50) default 'Tidak Hadir',
	deleted int default 0
)
go 
create table operator(
	username varchar(25) primary key,
	pass varchar(25)
)
go 

insert operator values ('admin', 'admin')


SELECT @@Identity

insert into guru values ('Yoni', 'L', '123', 'Belum nikah', 'islam', 'karanganyar', '12-11-1993' ,'guru', '123', 0)
insert into pelajaran values ('Matematika', 0)
insert into pelajaran values ('Biologi', 0)
insert into kelas values ('Kelas 1', 0)
insert into kelas values ('Kelas 2', 0)
insert into semester values ('Semester 1', 0)
insert into semester values ('Semester 2', 0)
insert into jadwal values (1,2016, '2016-07-15 01:00:00.000','2016-07-15 13:00:00.000',2,1,1,1,0)
select * from jadwal where id = 4
select * from jadwal 
where 
(datepart(hour, waktumulai) <= 8 or datepart(hour, waktuselesai) >= 8) and
datepart(minute, waktumulai) <= 8 and datepart(minute, waktuselesai) >= 8 and
datepart(hour, waktumulai) <= 8 and datepart(hour, waktuselesai) >= 8
delete from rekapjadwal
select * from rekapjadwal
select * from guru
select rtrim(ltrim('asdf  '))
