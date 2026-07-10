import pool from '../config/db.js';
import { Kehadiran, TipeUser } from '../utils/constants.js';
import { sendWhatsAppNotification } from '../utils/whatsapp.js';

function nowTimeString() {
  const d = new Date();
  return d.toLocaleTimeString('sv-SE', { timeZone: 'Asia/Jakarta' }); // HH:MM:SS WIB
}

function todayDateString() {
  const d = new Date();
  return d.toLocaleDateString('sv-SE', { timeZone: 'Asia/Jakarta' }); // YYYY-MM-DD WIB
}

async function findSiswaByCode(code) {
  const [rows] = await pool.query(
    `SELECT s.*, k.tingkat, k.index_kelas, j.jurusan,
            CONCAT(k.tingkat, ' ', j.jurusan, ' ', k.index_kelas) AS kelas
     FROM tb_siswa s
     LEFT JOIN tb_kelas k ON k.id_kelas = s.id_kelas
     LEFT JOIN tb_jurusan j ON j.id = k.id_jurusan
     WHERE s.unique_code = ? OR s.rfid_code = ? LIMIT 1`,
    [code, code]
  );
  return rows[0] || null;
}

async function findGuruByCode(code) {
  const [rows] = await pool.query(
    `SELECT * FROM tb_guru WHERE unique_code = ? OR rfid_code = ? LIMIT 1`,
    [code, code]
  );
  return rows[0] || null;
}

/**
 * POST /api/scan/cek
 * body: { unique_code, waktu: 'masuk' | 'pulang' }
 */
export async function cekKode(req, res) {
  try {
    const { unique_code, waktu } = req.body;

    if (!unique_code) {
      return res.status(400).json({ success: false, message: 'Kode tidak boleh kosong.' });
    }

    let type = TipeUser.Siswa;
    let result = await findSiswaByCode(unique_code);

    if (!result) {
      result = await findGuruByCode(unique_code);
      if (result) {
        type = TipeUser.Guru;
      }
    }

    if (!result) {
      return res.status(404).json({ success: false, message: 'Data tidak ditemukan' });
    }

    if (waktu === 'masuk') {
      return await absenMasuk(type, result, res);
    } else if (waktu === 'pulang') {
      return await absenPulang(type, result, res);
    } else {
      return res.status(400).json({ success: false, message: 'Data tidak valid' });
    }
  } catch (err) {
    console.error('cekKode error:', err);
    res.status(500).json({ success: false, message: 'Terjadi kesalahan pada server.' });
  }
}

async function absenMasuk(type, result, res) {
  const date = todayDateString();
  const time = nowTimeString();
  let messageString = ` sudah absen masuk pada tanggal ${date} jam ${time}`;
  let presensi;

  if (type === TipeUser.Guru) {
    const [[sudahAbsen]] = await pool.query(
      'SELECT * FROM tb_presensi_guru WHERE id_guru = ? AND tanggal = ?',
      [result.id_guru, date]
    );

    if (sudahAbsen) {
      return res.json({
        success: false,
        alreadyMarked: true,
        message: 'Anda sudah absen hari ini',
        type,
        data: result,
        presensi: sudahAbsen,
        waktu: 'masuk',
      });
    }

    await pool.query(
      `INSERT INTO tb_presensi_guru (id_guru, tanggal, jam_masuk, id_kehadiran, keterangan) VALUES (?, ?, ?, ?, '')`,
      [result.id_guru, date, time, Kehadiran.Hadir]
    );
    messageString = `${result.nama_guru} dengan NIP ${result.nuptk}${messageString}`;

    const [[p]] = await pool.query('SELECT * FROM tb_presensi_guru WHERE id_guru = ? AND tanggal = ?', [result.id_guru, date]);
    presensi = p;
  } else {
    const [[sudahAbsen]] = await pool.query(
      'SELECT * FROM tb_presensi_siswa WHERE id_siswa = ? AND tanggal = ?',
      [result.id_siswa, date]
    );

    if (sudahAbsen) {
      return res.json({
        success: false,
        alreadyMarked: true,
        message: 'Anda sudah absen hari ini',
        type,
        data: result,
        presensi: sudahAbsen,
        waktu: 'masuk',
      });
    }

    await pool.query(
      `INSERT INTO tb_presensi_siswa (id_siswa, id_kelas, tanggal, jam_masuk, id_kehadiran, keterangan) VALUES (?, ?, ?, ?, ?, '')`,
      [result.id_siswa, result.id_kelas, date, time, Kehadiran.Hadir]
    );
    messageString = `Siswa ${result.nama_siswa} dengan NIS ${result.nis}${messageString}`;

    const [[p]] = await pool.query('SELECT * FROM tb_presensi_siswa WHERE id_siswa = ? AND tanggal = ?', [result.id_siswa, date]);
    presensi = p;
  }

  // Kirim notifikasi WhatsApp ke orang tua/guru secara real-time
  if (result.no_hp) {
    sendWhatsAppNotification(result.no_hp, messageString).catch((e) =>
      console.error('Gagal mengirim notifikasi WA:', e.message)
    );
  }

  res.json({ success: true, type, data: result, presensi, waktu: 'masuk' });
}

async function absenPulang(type, result, res) {
  const date = todayDateString();
  const time = nowTimeString();
  let messageString = ` sudah absen pulang pada tanggal ${date} jam ${time}`;
  let presensi;

  if (type === TipeUser.Guru) {
    const [[sudahAbsen]] = await pool.query(
      'SELECT * FROM tb_presensi_guru WHERE id_guru = ? AND tanggal = ?',
      [result.id_guru, date]
    );

    if (!sudahAbsen) {
      return res.json({ success: false, message: 'Anda belum absen hari ini', type, data: result, waktu: 'pulang' });
    }

    await pool.query('UPDATE tb_presensi_guru SET jam_keluar = ? WHERE id_presensi = ?', [time, sudahAbsen.id_presensi]);
    messageString = `${result.nama_guru} dengan NIP ${result.nuptk}${messageString}`;
    presensi = { ...sudahAbsen, jam_keluar: time };
  } else {
    const [[sudahAbsen]] = await pool.query(
      'SELECT * FROM tb_presensi_siswa WHERE id_siswa = ? AND tanggal = ?',
      [result.id_siswa, date]
    );

    if (!sudahAbsen) {
      return res.json({ success: false, message: 'Anda belum absen hari ini', type, data: result, waktu: 'pulang' });
    }

    await pool.query('UPDATE tb_presensi_siswa SET jam_keluar = ? WHERE id_presensi = ?', [time, sudahAbsen.id_presensi]);
    messageString = `Siswa ${result.nama_siswa} dengan NIS ${result.nis}${messageString}`;
    presensi = { ...sudahAbsen, jam_keluar: time };
  }

  if (result.no_hp) {
    sendWhatsAppNotification(result.no_hp, messageString).catch((e) =>
      console.error('Gagal mengirim notifikasi WA:', e.message)
    );
  }

  res.json({ success: true, type, data: result, presensi, waktu: 'pulang' });
}
