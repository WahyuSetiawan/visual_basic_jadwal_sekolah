Select jadwal.id, guru.nama as namaguru, pelajaran.nama as namapelajaran, semester.nama as namasemester, jadwal.waktumulai, jadwal.waktuselesai
,count(rekapjadwal.id) as jumlah, rekapjadwal.keterangan
from jadwal
inner join rekapjadwal on rekapjadwal.idjadwal = jadwal.id
inner join guru on guru.id  = jadwal.idguru
inner join pelajaran on pelajaran.id = jadwal.idpelajaran
inner join semester on semester.id = jadwal.semester
where jadwal.deleted = 0 and  guru.deleted = 0 and pelajaran.deleted = 0
and guru.id = 1
and pelajaran.id = 1
and jadwal.tahun = 2016
and jadwal.semester = 1
and jadwal.id = 1
group by rekapjadwal.keterangan, jadwal.id, guru.nama, pelajaran.nama , semester.nama , jadwal.waktumulai, jadwal.waktuselesai

select  * from rekapjadwal