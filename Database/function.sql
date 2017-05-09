alter function GetJadwalFromWaktu(@date  datetime)
RETURNS @returntable TABLE (
	id int,
	semester int,
	tahun int,
	waktumulai datetime not null,
	waktuselesai datetime not null,
	hari int,
	idkelas int,
	idpelajaran int,
	idguru int,
	deleted int
)
AS
BEGIN
	declare @wktmulai datetime
	declare @waktusekarang datetime
	declare @wktsekarang datetime

	select @waktusekarang = @date

	set @wktsekarang =  convert(datetime, '2000-01-01 ' + convert(varchar(2), datepart(hour, @waktusekarang)) + ':'+ convert(varchar(2), datepart(minute, @waktusekarang)) + ':'+ convert(varchar(2), datepart(second, @waktusekarang)) +'.000' )

	insert into @returntable select * from jadwal 
	where convert(datetime, '2000-01-01 ' + convert(varchar(2), datepart(hour, waktumulai)) + ':'+ convert(varchar(2), datepart(minute, waktumulai)) + ':'+ convert(varchar(2), datepart(second, waktumulai)) +'.000' ) < @wktsekarang 
	and convert(datetime, '2000-01-01 ' + convert(varchar(2), datepart(hour, waktuselesai)) + ':'+ convert(varchar(2), datepart(minute, waktuselesai)) + ':'+ convert(varchar(2), datepart(second, waktuselesai)) +'.000' ) > @wktsekarang

	return
end

create function test()
returns varchar(10)
as 
begin
return 'asdfasdfas'
select * from guru
select * from jadwal WHERE waktumulai - CAST(FLOOR(CAST(waktumulai AS float)) AS datetime) < '01:00' and waktuselesai - CAST(FLOOR(CAST(waktuselesai AS float)) AS datetime) > '01:00'  
select * from GetJadwalFromWaktu('2016-08-26 2:54:26.480')
Select * from rekapjadwal 
where CAST(FLOOR(CAST(waktumulai AS float)) AS datetime) = '2016-07-15'

select convert(datetime, '2000-01-01 ' + convert(varchar(2), datepart(hour, waktumulai)) + ':'+ convert(varchar(2), datepart(minute, waktumulai)) + ':'+ convert(varchar(2), datepart(second, waktumulai)) +'.000' ) from jadwal

select * from GetJadwalFromWaktu('2016-07-15 12:00:00.000')