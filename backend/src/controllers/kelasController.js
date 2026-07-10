import pool from '../config/db.js';

const KELAS_QUERY = `
  SELECT k.*, j.jurusan,
    g.nama_guru AS nama_wali_kelas,
    CONCAT(k.tingkat, ' ', j.jurusan, IF(k.index_kelas != '', CONCAT(' ', k.index_kelas), '')) AS nama_kelas,
    (SELECT COUNT(*) FROM tb_siswa s WHERE s.id_kelas = k.id_kelas) AS jumlah_siswa
  FROM tb_kelas k
  LEFT JOIN tb_jurusan j ON j.id = k.id_jurusan
  LEFT JOIN tb_guru g ON g.id_guru = k.id_wali_kelas
`;

export async function listKelas(req, res) {
  try {
    const [rows] = await pool.query(`${KELAS_QUERY} ORDER BY k.tingkat, j.jurusan, k.index_kelas`);
    res.json({ success: true, data: rows, empty: rows.length === 0 });
  } catch (err) {
    console.error('listKelas error:', err);
    res.status(500).json({ success: false, message: 'Gagal mengambil data kelas.' });
  }
}

export async function getKelasById(req, res) {
  try {
    const [rows] = await pool.query(`${KELAS_QUERY} WHERE k.id_kelas = ?`, [req.params.id]);
    if (!rows[0]) return res.status(404).json({ success: false, message: 'Data kelas tidak ditemukan.' });
    res.json({ success: true, data: rows[0] });
  } catch (err) {
    console.error('getKelasById error:', err);
    res.status(500).json({ success: false, message: 'Terjadi kesalahan server.' });
  }
}

function validateKelasInput(body) {
  const errors = {};
  if (!body.tingkat || body.tingkat.length > 10) errors.tingkat = 'Tingkat harus diisi (maks. 10 karakter).';
  if (!body.id_jurusan) errors.id_jurusan = 'Jurusan harus diisi.';
  if (!body.index_kelas || body.index_kelas.length > 5) errors.index_kelas = 'Index kelas harus diisi (maks. 5 karakter).';
  return errors;
}

export async function createKelas(req, res) {
  try {
    const errors = validateKelasInput(req.body);
    if (Object.keys(errors).length > 0) {
      return res.status(422).json({ success: false, message: 'Validasi gagal.', errors });
    }

    const [result] = await pool.query(
      `INSERT INTO tb_kelas (tingkat, id_jurusan, index_kelas, id_wali_kelas, created_at, updated_at)
       VALUES (?, ?, ?, ?, NOW(), NOW())`,
      [req.body.tingkat, req.body.id_jurusan, req.body.index_kelas, req.body.id_wali_kelas || null]
    );
    res.json({ success: true, message: 'Tambah data berhasil.', id: result.insertId });
  } catch (err) {
    console.error('createKelas error:', err);
    res.status(500).json({ success: false, message: 'Gagal menambah data.' });
  }
}

export async function updateKelas(req, res) {
  try {
    const id = req.body.id || req.params.id;
    const errors = validateKelasInput(req.body);
    if (Object.keys(errors).length > 0) {
      return res.status(422).json({ success: false, message: 'Validasi gagal.', errors });
    }

    await pool.query(
      `UPDATE tb_kelas SET tingkat=?, id_jurusan=?, index_kelas=?, id_wali_kelas=?, updated_at=NOW() WHERE id_kelas=?`,
      [req.body.tingkat, req.body.id_jurusan, req.body.index_kelas, req.body.id_wali_kelas || null, id]
    );
    res.json({ success: true, message: 'Edit data berhasil.' });
  } catch (err) {
    console.error('updateKelas error:', err);
    res.status(500).json({ success: false, message: 'Gagal mengubah data.' });
  }
}

export async function deleteKelas(req, res) {
  try {
    const id = req.body.id || req.params.id;
    const [[countRow]] = await pool.query('SELECT COUNT(*) AS total FROM tb_siswa WHERE id_kelas = ?', [id]);
    if (Number(countRow.total) > 0) {
      return res.status(409).json({ success: false, message: 'Kelas masih memiliki siswa aktif, tidak dapat dihapus.' });
    }
    await pool.query('DELETE FROM tb_kelas WHERE id_kelas = ?', [id]);
    res.json({ success: true, message: 'Data berhasil dihapus.' });
  } catch (err) {
    console.error('deleteKelas error:', err);
    res.status(500).json({ success: false, message: 'Gagal menghapus data.' });
  }
}

export async function importKelasCSV(req, res) {
  try {
    const { rows } = req.body;
    if (!Array.isArray(rows) || rows.length === 0) {
      return res.status(400).json({ success: false, message: 'Tidak ada data untuk diimpor.' });
    }
    let success = 0;
    const failed = [];

    for (const row of rows) {
      try {
        const errors = validateKelasInput(row);
        if (Object.keys(errors).length > 0) {
          failed.push({ row, errors });
          continue;
        }
        await pool.query(
          `INSERT INTO tb_kelas (tingkat, id_jurusan, index_kelas, id_wali_kelas, created_at, updated_at) VALUES (?, ?, ?, ?, NOW(), NOW())`,
          [row.tingkat, row.id_jurusan, row.index_kelas, row.id_wali_kelas || null]
        );
        success++;
      } catch (e) {
        failed.push({ row, errors: { general: e.message } });
      }
    }
    res.json({ success: true, message: `${success} data berhasil diimpor, ${failed.length} gagal.`, imported: success, failed });
  } catch (err) {
    console.error('importKelasCSV error:', err);
    res.status(500).json({ success: false, message: 'Gagal mengimpor data.' });
  }
}

// ── Jurusan ──

export async function listJurusan(req, res) {
  try {
    const [rows] = await pool.query('SELECT * FROM tb_jurusan ORDER BY jurusan');
    res.json({ success: true, data: rows, empty: rows.length === 0 });
  } catch (err) {
    console.error('listJurusan error:', err);
    res.status(500).json({ success: false, message: 'Gagal mengambil data jurusan.' });
  }
}

export async function getJurusanById(req, res) {
  try {
    const [rows] = await pool.query('SELECT * FROM tb_jurusan WHERE id = ?', [req.params.id]);
    if (!rows[0]) return res.status(404).json({ success: false, message: 'Data jurusan tidak ditemukan.' });
    res.json({ success: true, data: rows[0] });
  } catch (err) {
    console.error('getJurusanById error:', err);
    res.status(500).json({ success: false, message: 'Terjadi kesalahan server.' });
  }
}

export async function createJurusan(req, res) {
  try {
    if (!req.body.jurusan || req.body.jurusan.length > 32) {
      return res.status(422).json({ success: false, message: 'Validasi gagal.', errors: { jurusan: 'Nama jurusan harus diisi (maks. 32 karakter).' } });
    }
    const [[dup]] = await pool.query('SELECT id FROM tb_jurusan WHERE jurusan = ?', [req.body.jurusan]);
    if (dup) {
      return res.status(422).json({ success: false, message: 'Validasi gagal.', errors: { jurusan: 'Jurusan ini sudah ada.' } });
    }
    const [result] = await pool.query('INSERT INTO tb_jurusan (jurusan, created_at, updated_at) VALUES (?, NOW(), NOW())', [req.body.jurusan]);
    res.json({ success: true, message: 'Tambah data berhasil.', id: result.insertId });
  } catch (err) {
    console.error('createJurusan error:', err);
    res.status(500).json({ success: false, message: 'Gagal menambah data.' });
  }
}

export async function updateJurusan(req, res) {
  try {
    const id = req.body.id || req.params.id;
    if (!req.body.jurusan || req.body.jurusan.length > 32) {
      return res.status(422).json({ success: false, message: 'Validasi gagal.', errors: { jurusan: 'Nama jurusan harus diisi (maks. 32 karakter).' } });
    }
    const [[dup]] = await pool.query('SELECT id FROM tb_jurusan WHERE jurusan = ? AND id != ?', [req.body.jurusan, id]);
    if (dup) {
      return res.status(422).json({ success: false, message: 'Validasi gagal.', errors: { jurusan: 'Jurusan ini sudah ada.' } });
    }
    await pool.query('UPDATE tb_jurusan SET jurusan=?, updated_at=NOW() WHERE id=?', [req.body.jurusan, id]);
    res.json({ success: true, message: 'Edit data berhasil.' });
  } catch (err) {
    console.error('updateJurusan error:', err);
    res.status(500).json({ success: false, message: 'Gagal mengubah data.' });
  }
}

export async function deleteJurusan(req, res) {
  try {
    const id = req.body.id || req.params.id;
    const [[countRow]] = await pool.query('SELECT COUNT(*) AS total FROM tb_kelas WHERE id_jurusan = ?', [id]);
    if (Number(countRow.total) > 0) {
      return res.status(409).json({ success: false, message: 'Jurusan masih digunakan oleh kelas, tidak dapat dihapus.' });
    }
    await pool.query('DELETE FROM tb_jurusan WHERE id = ?', [id]);
    res.json({ success: true, message: 'Data berhasil dihapus.' });
  } catch (err) {
    console.error('deleteJurusan error:', err);
    res.status(500).json({ success: false, message: 'Gagal menghapus data.' });
  }
}

export async function importJurusanCSV(req, res) {
  try {
    const { rows } = req.body;
    if (!Array.isArray(rows) || rows.length === 0) {
      return res.status(400).json({ success: false, message: 'Tidak ada data untuk diimpor.' });
    }
    let success = 0;
    const failed = [];
    for (const row of rows) {
      try {
        if (!row.jurusan) { failed.push({ row, errors: { jurusan: 'Wajib diisi' } }); continue; }
        const [[dup]] = await pool.query('SELECT id FROM tb_jurusan WHERE jurusan = ?', [row.jurusan]);
        if (dup) { failed.push({ row, errors: { jurusan: 'Sudah ada' } }); continue; }
        await pool.query('INSERT INTO tb_jurusan (jurusan, created_at, updated_at) VALUES (?, NOW(), NOW())', [row.jurusan]);
        success++;
      } catch (e) {
        failed.push({ row, errors: { general: e.message } });
      }
    }
    res.json({ success: true, message: `${success} data berhasil diimpor, ${failed.length} gagal.`, imported: success, failed });
  } catch (err) {
    console.error('importJurusanCSV error:', err);
    res.status(500).json({ success: false, message: 'Gagal mengimpor data.' });
  }
}
