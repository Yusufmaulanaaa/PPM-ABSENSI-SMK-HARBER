import pool from '../config/db.js';
import {
  countPresensiSiswa,
  getAttendanceTrendSiswa,
  getDateRangeLabels,
  todayDateString,
} from '../utils/attendanceStats.js';

async function getKelasByWali(idGuru) {
  const [[kelas]] = await pool.query(
    `SELECT k.*, j.jurusan, CONCAT(k.tingkat, ' ', j.jurusan, IF(k.index_kelas != '', CONCAT(' ', k.index_kelas), '')) AS nama_kelas
     FROM tb_kelas k LEFT JOIN tb_jurusan j ON j.id = k.id_jurusan
     WHERE k.id_wali_kelas = ? LIMIT 1`,
    [idGuru]
  );
  return kelas || null;
}

export async function getTeacherDashboard(req, res) {
  try {
    const idGuru = req.user.id_guru;
    if (!idGuru) {
      return res.status(403).json({ success: false, message: 'Anda bukan Guru.' });
    }

    const kelas = await getKelasByWali(idGuru);
    if (!kelas) {
      return res.json({ success: true, data: { noClass: true } });
    }

    const today = todayDateString();
    const [[totalSiswaRow]] = await pool.query('SELECT COUNT(*) AS total FROM tb_siswa WHERE id_kelas = ?', [kelas.id_kelas]);

    const [hadir, sakit, izin, alfa] = await Promise.all([
      countPresensiSiswa(1, today, kelas.id_kelas),
      countPresensiSiswa(2, today, kelas.id_kelas),
      countPresensiSiswa(3, today, kelas.id_kelas),
      countPresensiSiswa(4, today, kelas.id_kelas),
    ]);

    const grafikKehadiran = await getAttendanceTrendSiswa(7, kelas.id_kelas);

    res.json({
      success: true,
      data: {
        noClass: false,
        kelas,
        summary: {
          total_siswa: Number(totalSiswaRow.total),
          hadir_hari_ini: Number(hadir),
          sakit_hari_ini: Number(sakit),
          izin_hari_ini: Number(izin),
          alfa_hari_ini: Number(alfa),
        },
        dateRange: getDateRangeLabels(7),
        grafikKehadiran,
      },
    });
  } catch (err) {
    console.error('getTeacherDashboard error:', err);
    res.status(500).json({ success: false, message: 'Gagal memuat dashboard.' });
  }
}

export async function getTeacherAttendancePage(req, res) {
  try {
    const idGuru = req.user.id_guru;
    const kelas = await getKelasByWali(idGuru);
    if (!kelas) {
      return res.status(403).json({ success: false, message: 'Anda belum ditugaskan sebagai Wali Kelas.' });
    }
    res.json({ success: true, data: { kelas, date: todayDateString() } });
  } catch (err) {
    console.error('getTeacherAttendancePage error:', err);
    res.status(500).json({ success: false, message: 'Gagal memuat halaman.' });
  }
}

export async function getTeacherAttendanceList(req, res) {
  try {
    const idGuru = req.user.id_guru;
    const kelas = await getKelasByWali(idGuru);
    if (!kelas) return res.status(403).json({ success: false, message: 'Anda bukan wali kelas.' });

    const { tanggal } = req.body;
    const [rows] = await pool.query(
      `SELECT s.*, p.id_presensi, p.tanggal, p.jam_masuk, p.jam_keluar, p.id_kehadiran, p.keterangan
       FROM tb_siswa s
       LEFT JOIN tb_presensi_siswa p ON p.id_siswa = s.id_siswa AND p.tanggal = ?
       WHERE s.id_kelas = ? ORDER BY s.nama_siswa`,
      [tanggal, kelas.id_kelas]
    );

    const lewat = new Date(tanggal) > new Date(new Date().toDateString());
    res.json({ success: true, data: rows, kelas: kelas.nama_kelas, lewat, empty: rows.length === 0 });
  } catch (err) {
    console.error('getTeacherAttendanceList error:', err);
    res.status(500).json({ success: false, message: 'Gagal mengambil data.' });
  }
}

export async function updateTeacherAttendance(req, res) {
  try {
    const idGuru = req.user.id_guru;
    const kelas = await getKelasByWali(idGuru);
    if (!kelas) return res.status(403).json({ success: false, message: 'Anda bukan wali kelas.' });

    const { id_siswa, tanggal, id_kehadiran, jam_masuk, jam_keluar, keterangan } = req.body;

    const [[existing]] = await pool.query(
      'SELECT id_presensi FROM tb_presensi_siswa WHERE id_siswa = ? AND tanggal = ?',
      [id_siswa, tanggal]
    );

    if (existing) {
      const fields = ['id_kehadiran = ?', 'keterangan = ?'];
      const params = [id_kehadiran, keterangan || ''];
      if (jam_masuk) { fields.push('jam_masuk = ?'); params.push(jam_masuk); }
      if (jam_keluar) { fields.push('jam_keluar = ?'); params.push(jam_keluar); }
      params.push(existing.id_presensi);
      await pool.query(`UPDATE tb_presensi_siswa SET ${fields.join(', ')} WHERE id_presensi = ?`, params);
    } else {
      await pool.query(
        `INSERT INTO tb_presensi_siswa (id_siswa, id_kelas, tanggal, jam_masuk, jam_keluar, id_kehadiran, keterangan)
         VALUES (?, ?, ?, ?, ?, ?, ?)`,
        [id_siswa, kelas.id_kelas, tanggal, jam_masuk || null, jam_keluar || null, id_kehadiran, keterangan || '']
      );
    }

    const [[siswa]] = await pool.query('SELECT nama_siswa FROM tb_siswa WHERE id_siswa = ?', [id_siswa]);
    res.json({ success: true, status: true, nama_siswa: siswa?.nama_siswa || '' });
  } catch (err) {
    console.error('updateTeacherAttendance error:', err);
    res.status(500).json({ success: false, status: false, message: 'Gagal mengubah data.' });
  }
}

export async function getTeacherQrPage(req, res) {
  try {
    const idGuru = req.user.id_guru;
    const kelas = await getKelasByWali(idGuru);
    if (!kelas) return res.status(403).json({ success: false, message: 'Anda belum ditugaskan sebagai Wali Kelas.' });

    const [siswa] = await pool.query('SELECT * FROM tb_siswa WHERE id_kelas = ? ORDER BY nama_siswa', [kelas.id_kelas]);
    res.json({ success: true, data: { kelas, siswa } });
  } catch (err) {
    console.error('getTeacherQrPage error:', err);
    res.status(500).json({ success: false, message: 'Gagal memuat halaman.' });
  }
}

export async function getTeacherReportsPage(req, res) {
  try {
    const idGuru = req.user.id_guru;
    const kelas = await getKelasByWali(idGuru);
    res.json({ success: true, data: { kelas } });
  } catch (err) {
    console.error('getTeacherReportsPage error:', err);
    res.status(500).json({ success: false, message: 'Gagal memuat halaman.' });
  }
}
