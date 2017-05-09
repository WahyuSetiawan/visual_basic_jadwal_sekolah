Select 
jadwal.idguru, 
jadwal.waktumulai,
jadwal.waktuselesai, 
semester.nama as namasemester, 
jadwal.tahun, 
convert(varchar(10), tanggal, 105) as tanggal,
(CASE WHEN hari=1 then 'Minggu'
WHEN hari=2 THEN 'Senin'
WHEN hari=3 THEN 'Selasa'
WHEN hari=4 THEN 'Rabu'
WHEN hari=5 THEN 'Kamis'
WHEN hari=6 THEN 'Jumat' ELSE 'Sabtu' END ) as namahari,
guru.nama as namaguru, 
guru.nip,
pelajaran.nama as namapelajaran, 
kelas.nama as namakelas 
from jadwal 
inner join guru on guru.id = jadwal.idguru 
inner join pelajaran on pelajaran.id = jadwal.idpelajaran 
inner join kelas on kelas.id = jadwal.idkelas
inner join semester on semester.id = jadwal.semester
inner join rekapjadwal on rekapjadwal.idjadwal = jadwal.id

select * from guru
select * from rekapjadwal

select guru.* from guru 
inner join jadwal on jadwal.idguru = guru.id 
inner join rekapjadwal on jadwal.id = rekapjadwal.idjadwal 
where guru.id = 2 and nip = '005' and rekapjadwal.id = 6 and guru.deleted = 0


select Guru.id, guru.nama, guru.jabatan, guru.nip, pelajaran.nama as namapelajaran,  jadwal.tahun, semester.nama as namasemester  from jadwal 
inner join guru on guru.id = jadwal.idguru 
inner join semester on semester.id = jadwal.semester
inner join pelajaran on pelajaran.id = jadwal.idpelajaran
group by Guru.id, guru.nama, guru.jabatan, guru.nip, pelajaran.nama,  jadwal.tahun, semester.nama