import pool from '../config/db.js';
import { generateToken } from '../utils/token.js';
import fs from 'fs';
import path from 'path';
import { uploadsRoot } from '../middleware/upload.js';

const KELAS_JOIN = `
  SELECT s.*, k.tingkat, k.index_kelas, j.jurusan,
    CONCAT(k.tingkat, ' ', j.jurusan, IF(k.index_kelas != '', CONCAT(' ', k.index_kelas), '')) AS kelas
  FROM tb_siswa s
  LEFT JOIN tb_kelas k ON k.id_kelas = s.id_kelas
  LEFT JOIN tb_jurusan j ON j.id = k.id_jurusan
`;

export async function listSiswa(req, res) {
  try {
    const { kelas, jurusan, index } = req.query;
    let query = KELAS_JOIN + ' WHERE 1=1';
    const params = [];

    if (kelas) { query += ' AND k.tingkat = ?'; params.push(kelas); }
    if (jurusan) { query += ' AND j.jurusan = ?'; params.push(jurusan); }
    if (index) { query += ' AND k.index_kelas = ?'; params.push(index); }

    query += ' ORDER BY s.nama_siswa';

    const [rows] = await pool.query(query, params);
    res.json({ success: true, data: rows, empty: rows.length === 0 });
  } catch (err) {
    console.error('listSiswa error:', err);
    res.status(500).json({ success: false, message: 'Gagal mengambil data siswa.' });
  }
}

export async function getSiswaById(req, res) {
  try {
    const [rows] = await pool.query(`${KELAS_JOIN} WHERE s.id_siswa = ?`, [req.params.id]);
    if (!rows[0]) return res.status(404).json({ success: false, message: 'Data siswa tidak ditemukan.' });
    res.json({ success: true, data: rows[0] });
  } catch (err) {
    console.error('getSiswaById error:', err);
    res.status(500).json({ success: false, message: 'Terjadi kesalahan server.' });
  }
}

function validateSiswaInput(body, isUpdate = false) {
  const errors = {};
  if (!body.nis || body.nis.length < 4 || body.nis.length > 20) {
    errors.nis = 'NIS harus diisi (4-20 karakter).';
  }
  if (!body.nama || body.nama.trim().length < 3) {
    errors.nama = 'Nama harus diisi minimal 3 karakter.';
  }
  if (!body.id_kelas) {
    errors.id_kelas = 'Kelas harus diisi.';
  }
  if (!body.jk) {
    errors.jk = 'Jenis kelamin wajib diisi.';
  }
  if (!body.no_hp || !/^\d{5,20}$/.test(body.no_hp)) {
    errors.no_hp = 'No HP harus diisi, hanya angka (5-20 digit).';
  }
  return errors;
}

export async function createSiswa(req, res) {
  try {
    const errors = validateSiswaInput(req.body);

    const [[dup]] = await pool.query('SELECT id_siswa FROM tb_siswa WHERE nis = ?', [req.body.nis]);
    if (dup) errors.nis = 'NIS ini telah terdaftar.';

    if (req.body.rfid) {
      const [[dupRfid]] = await pool.query('SELECT id_siswa FROM tb_siswa WHERE rfid_code = ?', [req.body.rfid]);
      if (dupRfid) errors.rfid = 'RFID code sudah digunakan.';
    }

    if (Object.keys(errors).length > 0) {
      return res.status(422).json({ success: false, message: 'Validasi gagal.', errors });
    }

    const foto = req.file ? req.file.filename : null;
    const uniqueCode = generateToken();

    const [result] = await pool.query(
      `INSERT INTO tb_siswa (nis, nama_siswa, id_kelas, jenis_kelamin, no_hp, unique_code, rfid_code, foto)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
      [req.body.nis, req.body.nama, req.body.id_kelas, req.body.jk, req.body.no_hp, uniqueCode, req.body.rfid || null, foto]
    );

    res.json({ success: true, message: 'Tambah data berhasil.', id: result.insertId });
  } catch (err) {
    console.error('createSiswa error:', err);
    res.status(500).json({ success: false, message: 'Gagal menambah data.' });
  }
}

export async function updateSiswa(req, res) {
  try {
    const id = req.body.id || req.params.id;
    const errors = validateSiswaInput(req.body, true);

    const [[existing]] = await pool.query('SELECT * FROM tb_siswa WHERE id_siswa = ?', [id]);
    if (!existing) return res.status(404).json({ success: false, message: 'Data siswa tidak ditemukan.' });

    if (existing.nis !== req.body.nis) {
      const [[dup]] = await pool.query('SELECT id_siswa FROM tb_siswa WHERE nis = ? AND id_siswa != ?', [req.body.nis, id]);
      if (dup) errors.nis = 'NIS ini telah terdaftar.';
    }

    if (req.body.rfid) {
      const [[dupRfid]] = await pool.query('SELECT id_siswa FROM tb_siswa WHERE rfid_code = ? AND id_siswa != ?', [req.body.rfid, id]);
      if (dupRfid) errors.rfid = 'RFID code sudah digunakan.';
    }

    if (Object.keys(errors).length > 0) {
      return res.status(422).json({ success: false, message: 'Validasi gagal.', errors });
    }

    let foto = existing.foto;
    if (req.file) {
      foto = req.file.filename;
      if (existing.foto) {
        const oldPath = path.join(uploadsRoot, 'siswa', existing.foto);
        if (fs.existsSync(oldPath)) fs.unlinkSync(oldPath);
      }
    }

    await pool.query(
      `UPDATE tb_siswa SET nis=?, nama_siswa=?, id_kelas=?, jenis_kelamin=?, no_hp=?, rfid_code=?, foto=? WHERE id_siswa=?`,
      [req.body.nis, req.body.nama, req.body.id_kelas, req.body.jk, req.body.no_hp, req.body.rfid || null, foto, id]
    );

    res.json({ success: true, message: 'Edit data berhasil.' });
  } catch (err) {
    console.error('updateSiswa error:', err);
    res.status(500).json({ success: false, message: 'Gagal mengubah data.' });
  }
}

export async function deleteSiswa(req, res) {
  try {
    const [[existing]] = await pool.query('SELECT foto FROM tb_siswa WHERE id_siswa = ?', [req.params.id]);
    await pool.query('DELETE FROM tb_siswa WHERE id_siswa = ?', [req.params.id]);
    if (existing?.foto) {
      const p = path.join(uploadsRoot, 'siswa', existing.foto);
      if (fs.existsSync(p)) fs.unlinkSync(p);
    }
    res.json({ success: true, message: 'Data berhasil dihapus.' });
  } catch (err) {
    console.error('deleteSiswa error:', err);
    res.status(500).json({ success: false, message: 'Gagal menghapus data.' });
  }
}

export async function deleteSelectedSiswa(req, res) {
  try {
    const { siswa_ids } = req.body;
    if (!Array.isArray(siswa_ids) || siswa_ids.length === 0) {
      return res.status(400).json({ success: false, message: 'Tidak ada data yang dipilih.' });
    }
    await pool.query('DELETE FROM tb_siswa WHERE id_siswa IN (?)', [siswa_ids]);
    res.json({ success: true, message: `${siswa_ids.length} data berhasil dihapus.` });
  } catch (err) {
    console.error('deleteSelectedSiswa error:', err);
    res.status(500).json({ success: false, message: 'Gagal menghapus data terpilih.' });
  }
}

export async function importSiswaCSV(req, res) {
  try {
    const { rows } = req.body; // array of { nis, nama, id_kelas, jk, no_hp, rfid }
    if (!Array.isArray(rows) || rows.length === 0) {
      return res.status(400).json({ success: false, message: 'Tidak ada data untuk diimpor.' });
    }

    let success = 0;
    const failed = [];

    for (const row of rows) {
      try {
        const errors = validateSiswaInput(row);
        const [[dup]] = await pool.query('SELECT id_siswa FROM tb_siswa WHERE nis = ?', [row.nis]);
        if (dup) errors.nis = 'NIS sudah terdaftar.';

        if (Object.keys(errors).length > 0) {
          failed.push({ row, errors });
          continue;
        }

        await pool.query(
          `INSERT INTO tb_siswa (nis, nama_siswa, id_kelas, jenis_kelamin, no_hp, unique_code, rfid_code)
           VALUES (?, ?, ?, ?, ?, ?, ?)`,
          [row.nis, row.nama, row.id_kelas, row.jk, row.no_hp, generateToken(), row.rfid || null]
        );
        success++;
      } catch (e) {
        failed.push({ row, errors: { general: e.message } });
      }
    }

    res.json({ success: true, message: `${success} data berhasil diimpor, ${failed.length} gagal.`, imported: success, failed });
  } catch (err) {
    console.error('importSiswaCSV error:', err);
    res.status(500).json({ success: false, message: 'Gagal mengimpor data.' });
  }
}
