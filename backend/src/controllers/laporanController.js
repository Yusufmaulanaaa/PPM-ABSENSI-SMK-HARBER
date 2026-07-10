import pool from '../config/db.js';
import PDFDocument from 'pdfkit';
import ExcelJS from 'exceljs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const LOGO_PATH = path.join(__dirname, '../../uploads/logo/edb9821bc24b5b092fb26a7d965b81f3.jpg');

const HARI_SINGKAT = ['Min', 'Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab'];
const BULAN_NAMA = [
  'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember',
];

function pad2(n) { return String(n).padStart(2, '0'); }
function toDateStr(d) { return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; }

/** Bangun array tanggal kerja (skip Sabtu/Minggu) untuk bulan "YYYY-MM" */
function buildWorkingDays(bulanStr) {
  const [year, month] = bulanStr.split('-').map(Number);
  const lastDate = new Date(year, month, 0).getDate();
  const days = [];
  for (let d = 1; d <= lastDate; d++) {
    const date = new Date(year, month - 1, d);
    const dow = date.getDay(); // 0=Min, 6=Sab
    if (dow !== 0 && dow !== 6) {
      days.push(date);
    }
  }
  return days;
}

async function getGeneralSettings() {
  const [[row]] = await pool.query('SELECT * FROM general_settings LIMIT 1');
  return row || { school_name: 'SMK Harapan Bersama Tegal', school_year: '2024/2025' };
}

// ============================================================
// LAPORAN SISWA
// ============================================================

export async function generateLaporanSiswaData(req, res) {
  try {
    const { kelas: idKelas, bulan } = req.body;

    const [siswaList] = await pool.query('SELECT * FROM tb_siswa WHERE id_kelas = ? ORDER BY nama_siswa', [idKelas]);
    if (siswaList.length === 0) {
      return res.status(404).json({ success: false, message: 'Data siswa kosong!' });
    }

    const [[kelasRow]] = await pool.query(
      `SELECT k.*, j.jurusan, CONCAT(k.tingkat, ' ', j.jurusan, IF(k.index_kelas != '', CONCAT(' ', k.index_kelas), '')) AS kelas
       FROM tb_kelas k LEFT JOIN tb_jurusan j ON j.id = k.id_jurusan WHERE k.id_kelas = ?`,
      [idKelas]
    );

    const workingDays = buildWorkingDays(bulan);
    const today = new Date(); today.setHours(0, 0, 0, 0);

    // ambil semua presensi sebulan untuk kelas ini sekaligus (lebih efisien)
    const [presensiBulan] = await pool.query(
      `SELECT * FROM tb_presensi_siswa WHERE id_kelas = ? AND tanggal BETWEEN ? AND ?`,
      [idKelas, toDateStr(workingDays[0]), toDateStr(workingDays[workingDays.length - 1])]
    );

    const presensiMap = {}; // key: `${id_siswa}_${tanggal}`
    for (const p of presensiBulan) {
      presensiMap[`${p.id_siswa}_${p.tanggal}`] = p;
    }

    const tanggalList = workingDays.map((d) => ({
      tanggal: toDateStr(d),
      hariSingkat: HARI_SINGKAT[d.getDay()],
      tgl: d.getDate(),
      lewat: d > today,
    }));

    const rows = siswaList.map((siswa) => {
      let hadir = 0, sakit = 0, izin = 0, alfa = 0;
      const harian = tanggalList.map((t) => {
        if (t.lewat) return { id_kehadiran: null, lewat: true };
        const p = presensiMap[`${siswa.id_siswa}_${t.tanggal}`];
        const idK = p ? p.id_kehadiran : 4;
        if (idK === 1) hadir++;
        else if (idK === 2) sakit++;
        else if (idK === 3) izin++;
        else alfa++;
        return { id_kehadiran: idK, lewat: false };
      });
      return { siswa, harian, hadir, sakit, izin, alfa };
    });

    const laki = siswaList.filter((s) => s.jenis_kelamin !== 'Perempuan').length;
    const generalSettings = await getGeneralSettings();

    res.json({
      success: true,
      data: {
        tanggalList,
        bulanLabel: `${BULAN_NAMA[Number(bulan.split('-')[1]) - 1]} ${bulan.split('-')[0]}`,
        kelas: kelasRow,
        rows,
        rekap: { laki, perempuan: siswaList.length - laki, total: siswaList.length },
        generalSettings,
      },
    });
  } catch (err) {
    console.error('generateLaporanSiswaData error:', err);
    res.status(500).json({ success: false, message: 'Gagal membuat laporan.' });
  }
}

export async function exportLaporanSiswaExcel(req, res) {
  try {
    const { kelas: idKelas, bulan } = req.body;
    let payload;
    await generateLaporanSiswaData({ body: { kelas: idKelas, bulan } }, { json: (p) => { payload = p; }, status: () => ({ json: (p) => { payload = p; } }) });

    if (!payload?.success) return res.status(404).json({ success: false, message: 'Data tidak ditemukan.' });

    const { tanggalList, bulanLabel, kelas, rows, rekap, generalSettings } = payload.data;
    const filename = `laporan_absen_${kelas.kelas.replace(/\s+/g, '_')}_${bulanLabel.replace(/\s+/g, '-')}.xlsx`;

    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Laporan Siswa');

    const HEADER_BG = 'FF4472C4';
    const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_BG } };
    const thinBorder = {
      top: { style: 'thin' }, left: { style: 'thin' },
      bottom: { style: 'thin' }, right: { style: 'thin' },
    };
    const centerAlign = { horizontal: 'center', vertical: 'middle' };

    const colorMap = {
      1: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF90EE90' } },
      2: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
      3: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
      4: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF6B6B' } },
    };
    const labelMap = { 1: 'H', 2: 'S', 3: 'I', 4: 'A' };

    // Header info rows
    ws.mergeCells(1, 1, 1, 4 + tanggalList.length);
    const titleCell = ws.getCell('A1');
    titleCell.value = `DAFTAR HADIR SISWA - ${generalSettings.school_name}`;
    titleCell.font = { bold: true, size: 14 };
    titleCell.alignment = { horizontal: 'center' };

    ws.mergeCells(2, 1, 2, 4 + tanggalList.length);
    ws.getCell('A2').value = `TAHUN PELAJARAN ${generalSettings.school_year}`;
    ws.getCell('A2').font = { bold: true, size: 11 };
    ws.getCell('A2').alignment = { horizontal: 'center' };

    ws.mergeCells(3, 1, 3, 4 + tanggalList.length);
    ws.getCell('A3').value = `Bulan: ${bulanLabel}    Kelas: ${kelas.kelas}`;
    ws.getCell('A3').font = { size: 10 };
    ws.getCell('A3').alignment = { horizontal: 'center' };

    // Table header row (row 5)
    const headerRow1 = ws.getRow(5);
    const headerRow2 = ws.getRow(6);
    const headerRow3 = ws.getRow(7);

    // Row 5: blank, blank, "Hari/Tanggal" merged, blank, blank, blank, blank
    ws.mergeCells(5, 3, 5, 2 + tanggalList.length);
    const hariTanggalCell = ws.getCell(5, 3);
    hariTanggalCell.value = 'Hari/Tanggal';
    hariTanggalCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9 };
    hariTanggalCell.fill = headerFill;
    hariTanggalCell.alignment = centerAlign;
    hariTanggalCell.border = thinBorder;

    ws.mergeCells(5, 3 + tanggalList.length, 5, 6 + tanggalList.length);
    const totalCell = ws.getCell(5, 3 + tanggalList.length);
    totalCell.value = 'Total';
    totalCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9 };
    totalCell.fill = headerFill;
    totalCell.alignment = centerAlign;
    totalCell.border = thinBorder;

    // Fill borders on row 5 for merged cells
    for (let c = 1; c <= 6 + tanggalList.length; c++) {
      const cell = ws.getCell(5, c);
      cell.border = thinBorder;
      if (!cell.fill || !cell.fill.fgColor) {
        cell.fill = headerFill;
      }
    }
    ws.getCell(5, 1).value = '';
    ws.getCell(5, 2).value = '';

    // Row 6: day names
    headerRow2.getCell(1).value = '';
    headerRow2.getCell(2).value = '';
    tanggalList.forEach((t, idx) => {
      const cell = headerRow2.getCell(3 + idx);
      cell.value = t.hariSingkat;
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 8 };
      cell.fill = headerFill;
      cell.alignment = centerAlign;
      cell.border = thinBorder;
    });
    for (let c = 3 + tanggalList.length; c <= 6 + tanggalList.length; c++) {
      const cell = headerRow2.getCell(c);
      cell.value = '';
      cell.fill = headerFill;
      cell.border = thinBorder;
    }
    headerRow2.getCell(1).fill = headerFill;
    headerRow2.getCell(1).border = thinBorder;
    headerRow2.getCell(2).fill = headerFill;
    headerRow2.getCell(2).border = thinBorder;

    // Row 7: No, Nama, dates, H, S, I, A
    headerRow3.getCell(1).value = 'No';
    headerRow3.getCell(2).value = 'Nama';
    tanggalList.forEach((t, idx) => {
      headerRow3.getCell(3 + idx).value = t.tgl;
    });
    headerRow3.getCell(3 + tanggalList.length).value = 'H';
    headerRow3.getCell(4 + tanggalList.length).value = 'S';
    headerRow3.getCell(5 + tanggalList.length).value = 'I';
    headerRow3.getCell(6 + tanggalList.length).value = 'A';

    for (let c = 1; c <= 6 + tanggalList.length; c++) {
      const cell = headerRow3.getCell(c);
      cell.font = { bold: true, size: 9 };
      cell.alignment = centerAlign;
      cell.border = thinBorder;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
    }
    // Color the H, S, I, A header cells
    headerRow3.getCell(3 + tanggalList.length).fill = colorMap[1];
    headerRow3.getCell(4 + tanggalList.length).fill = colorMap[2];
    headerRow3.getCell(5 + tanggalList.length).fill = colorMap[3];
    headerRow3.getCell(6 + tanggalList.length).fill = colorMap[4];

    // Data rows
    rows.forEach((r, i) => {
      const rowNum = 8 + i;
      const row = ws.getRow(rowNum);
      row.getCell(1).value = i + 1;
      row.getCell(1).alignment = centerAlign;
      row.getCell(2).value = r.siswa.nama_siswa;

      r.harian.forEach((h, idx) => {
        const cell = row.getCell(3 + idx);
        if (!h.lewat && h.id_kehadiran) {
          cell.value = labelMap[h.id_kehadiran];
          cell.fill = colorMap[h.id_kehadiran];
        }
        cell.alignment = centerAlign;
      });

      row.getCell(3 + tanggalList.length).value = r.hadir || 0;
      row.getCell(4 + tanggalList.length).value = r.sakit || 0;
      row.getCell(5 + tanggalList.length).value = r.izin || 0;
      row.getCell(6 + tanggalList.length).value = r.alfa || 0;

      for (let c = 1; c <= 6 + tanggalList.length; c++) {
        row.getCell(c).border = thinBorder;
        if (c >= 3 + tanggalList.length) {
          row.getCell(c).alignment = centerAlign;
        }
      }
    });

    // Summary rows
    const summaryStart = 8 + rows.length + 1;
    const summaryData = [
      ['Jumlah siswa', rekap.total],
      ['Laki-laki', rekap.laki],
      ['Perempuan', rekap.perempuan],
    ];
    summaryData.forEach(([label, val], i) => {
      const row = ws.getRow(summaryStart + i);
      row.getCell(1).value = label;
      row.getCell(2).value = val;
      row.getCell(1).font = { bold: true };
    });

    // Auto column widths
    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 22;
    tanggalList.forEach((_, idx) => {
      ws.getColumn(3 + idx).width = 4;
    });
    for (let c = 3 + tanggalList.length; c <= 6 + tanggalList.length; c++) {
      ws.getColumn(c).width = 5;
    }

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error('exportLaporanSiswaExcel error:', err);
    res.status(500).json({ success: false, message: 'Gagal mengekspor laporan Excel.' });
  }
}

export async function exportLaporanSiswaPdf(req, res) {
  try {
    const { kelas: idKelas, bulan } = req.body;
    let payload;
    await generateLaporanSiswaData({ body: { kelas: idKelas, bulan } }, { json: (p) => { payload = p; }, status: () => ({ json: (p) => { payload = p; } }) });

    if (!payload?.success) return res.status(404).json({ success: false, message: 'Data tidak ditemukan.' });

    const { tanggalList, bulanLabel, kelas, rows, rekap, generalSettings } = payload.data;
    const filename = `laporan_absen_${kelas.kelas.replace(/\s+/g, '_')}_${bulanLabel.replace(/\s+/g, '-')}.pdf`;

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);

    const doc = new PDFDocument({ margin: 30, size: 'A4', layout: 'landscape' });
    doc.pipe(res);

    // Header with logo
    try {
      doc.image(LOGO_PATH, 30, 30, { width: 50 });
    } catch (_) { /* logo not found, skip */ }
    doc.fontSize(14).text('DAFTAR HADIR SISWA', 90, 32, { align: 'left' });
    doc.fontSize(10).text(generalSettings.school_name, 90, 50, { align: 'left' });
    doc.text(`TAHUN PELAJARAN ${generalSettings.school_year}`, 90, 64, { align: 'left' });
    doc.moveDown(1.5);
    doc.fontSize(9).text(`Bulan: ${bulanLabel}     Kelas: ${kelas.kelas}`);
    doc.moveDown(0.5);

    const startX = 30;
    let y = doc.y;
    const colNo = 25, colNama = 100, colDay = 20, colSum = 22;
    const rowH = 22;

    function drawRow(cells, widths, opts = {}) {
      let x = startX;
      cells.forEach((c, i) => {
        if (opts.bg && opts.bg[i]) {
          doc.rect(x, y, widths[i], rowH).fill(opts.bg[i]);
          doc.fillColor('black');
        }
        doc.fontSize(7).text(String(c), x + 1, y + 7, { width: widths[i] - 2, align: 'center' });
        doc.lineWidth(0.5).rect(x, y, widths[i], rowH).stroke();
        x += widths[i];
      });
      y += rowH;
    }

    const widths = [colNo, colNama, ...tanggalList.map(() => colDay), colSum, colSum, colSum, colSum];

    // Combined header: day name + date number in one cell
    let hx = startX;
    doc.rect(startX, y, colNo + colNama, rowH).fill('#4472C4');
    doc.fillColor('white').fontSize(7).text('No / Nama', startX + 1, y + 7, { width: colNo + colNama - 2, align: 'center' });
    doc.fillColor('black');
    doc.lineWidth(0.5).rect(startX, y, colNo + colNama, rowH).stroke();
    hx = startX + colNo + colNama;
    tanggalList.forEach((t) => {
      doc.rect(hx, y, colDay, rowH).fill('#4472C4');
      doc.fillColor('white').fontSize(6).text(t.hariSingkat, hx + 1, y + 3, { width: colDay - 2, align: 'center' });
      doc.fontSize(7).text(String(t.tgl), hx + 1, y + 12, { width: colDay - 2, align: 'center' });
      doc.fillColor('black');
      doc.lineWidth(0.5).rect(hx, y, colDay, rowH).stroke();
      hx += colDay;
    });
    ['H', 'S', 'I', 'A'].forEach((label, idx) => {
      const bg = ['#90ee90', '#ffff00', '#ffff00', '#ff6b6b'][idx];
      doc.rect(hx, y, colSum, rowH).fill(bg);
      doc.fillColor('black').fontSize(7).text(label, hx + 1, y + 7, { width: colSum - 2, align: 'center' });
      doc.lineWidth(0.5).rect(hx, y, colSum, rowH).stroke();
      hx += colSum;
    });
    y += rowH;

    for (const [i, r] of rows.entries()) {
      if (y > doc.page.height - 60) {
        doc.addPage();
        y = doc.y;
      }
      const map = { 1: ['H', '#90ee90'], 2: ['S', '#ffff00'], 3: ['I', '#ffff00'], 4: ['A', '#ff6b6b'] };
      const harianCells = r.harian.map((h) => (h.lewat ? '' : (map[h.id_kehadiran]?.[0] || 'A')));
      const harianBg = r.harian.map((h) => (h.lewat ? null : (map[h.id_kehadiran]?.[1] || '#ff6b6b')));
      drawRow(
        [i + 1, r.siswa.nama_siswa, ...harianCells, r.hadir || '-', r.sakit || '-', r.izin || '-', r.alfa || '-'],
        widths,
        { bg: [null, null, ...harianBg, null, null, null, null] }
      );
    }

    doc.moveDown(1);
    doc.fontSize(9).text(`Jumlah siswa: ${rekap.total}    Laki-laki: ${rekap.laki}    Perempuan: ${rekap.perempuan}`, startX, y + 10);

    // Page number
    const pageCount = doc.bufferedPageRange().count;
    for (let p = 0; p < pageCount; p++) {
      doc.switchToPage(p);
      doc.fontSize(8).text(`Halaman ${p + 1} dari ${pageCount}`, 30, doc.page.height - 30, { width: doc.page.width - 60, align: 'center' });
    }

    doc.end();
  } catch (err) {
    console.error('exportLaporanSiswaPdf error:', err);
    res.status(500).json({ success: false, message: 'Gagal mengekspor laporan PDF.' });
  }
}

// ============================================================
// LAPORAN GURU
// ============================================================

export async function generateLaporanGuruData(req, res) {
  try {
    const { bulan } = req.body;
    const [guruList] = await pool.query('SELECT * FROM tb_guru ORDER BY nama_guru');
    if (guruList.length === 0) {
      return res.status(404).json({ success: false, message: 'Data guru kosong!' });
    }

    const workingDays = buildWorkingDays(bulan);
    const today = new Date(); today.setHours(0, 0, 0, 0);

    const [presensiBulan] = await pool.query(
      `SELECT * FROM tb_presensi_guru WHERE tanggal BETWEEN ? AND ?`,
      [toDateStr(workingDays[0]), toDateStr(workingDays[workingDays.length - 1])]
    );
    const presensiMap = {};
    for (const p of presensiBulan) presensiMap[`${p.id_guru}_${p.tanggal}`] = p;

    const tanggalList = workingDays.map((d) => ({
      tanggal: toDateStr(d), hariSingkat: HARI_SINGKAT[d.getDay()], tgl: d.getDate(), lewat: d > today,
    }));

    const rows = guruList.map((guru) => {
      let hadir = 0, sakit = 0, izin = 0, alfa = 0;
      const harian = tanggalList.map((t) => {
        if (t.lewat) return { id_kehadiran: null, lewat: true };
        const p = presensiMap[`${guru.id_guru}_${t.tanggal}`];
        const idK = p ? p.id_kehadiran : 4;
        if (idK === 1) hadir++; else if (idK === 2) sakit++; else if (idK === 3) izin++; else alfa++;
        return { id_kehadiran: idK, lewat: false };
      });
      return { guru, harian, hadir, sakit, izin, alfa };
    });

    const laki = guruList.filter((g) => g.jenis_kelamin !== 'Perempuan').length;
    const generalSettings = await getGeneralSettings();

    res.json({
      success: true,
      data: {
        tanggalList,
        bulanLabel: `${BULAN_NAMA[Number(bulan.split('-')[1]) - 1]} ${bulan.split('-')[0]}`,
        rows,
        rekap: { laki, perempuan: guruList.length - laki, total: guruList.length },
        generalSettings,
      },
    });
  } catch (err) {
    console.error('generateLaporanGuruData error:', err);
    res.status(500).json({ success: false, message: 'Gagal membuat laporan.' });
  }
}

export async function exportLaporanGuruExcel(req, res) {
  try {
    const { bulan } = req.body;
    let payload;
    await generateLaporanGuruData({ body: { bulan } }, { json: (p) => { payload = p; }, status: () => ({ json: (p) => { payload = p; } }) });
    if (!payload?.success) return res.status(404).json({ success: false, message: 'Data tidak ditemukan.' });

    const { tanggalList, bulanLabel, rows, rekap, generalSettings } = payload.data;
    const filename = `laporan_absen_guru_${bulanLabel.replace(/\s+/g, '-')}.xlsx`;

    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Laporan Guru');

    const HEADER_BG = 'FF4472C4';
    const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_BG } };
    const thinBorder = {
      top: { style: 'thin' }, left: { style: 'thin' },
      bottom: { style: 'thin' }, right: { style: 'thin' },
    };
    const centerAlign = { horizontal: 'center', vertical: 'middle' };

    const colorMap = {
      1: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF90EE90' } },
      2: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
      3: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
      4: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF6B6B' } },
    };
    const labelMap = { 1: 'H', 2: 'S', 3: 'I', 4: 'A' };

    // Header info rows
    ws.mergeCells(1, 1, 1, 4 + tanggalList.length);
    const titleCell = ws.getCell('A1');
    titleCell.value = `DAFTAR HADIR GURU - ${generalSettings.school_name}`;
    titleCell.font = { bold: true, size: 14 };
    titleCell.alignment = { horizontal: 'center' };

    ws.mergeCells(2, 1, 2, 4 + tanggalList.length);
    ws.getCell('A2').value = `TAHUN PELAJARAN ${generalSettings.school_year}`;
    ws.getCell('A2').font = { bold: true, size: 11 };
    ws.getCell('A2').alignment = { horizontal: 'center' };

    ws.mergeCells(3, 1, 3, 4 + tanggalList.length);
    ws.getCell('A3').value = `Bulan: ${bulanLabel}`;
    ws.getCell('A3').font = { size: 10 };
    ws.getCell('A3').alignment = { horizontal: 'center' };

    // Table header row (row 5)
    const headerRow2 = ws.getRow(6);
    const headerRow3 = ws.getRow(7);

    ws.mergeCells(5, 3, 5, 2 + tanggalList.length);
    const hariTanggalCell = ws.getCell(5, 3);
    hariTanggalCell.value = 'Hari/Tanggal';
    hariTanggalCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9 };
    hariTanggalCell.fill = headerFill;
    hariTanggalCell.alignment = centerAlign;
    hariTanggalCell.border = thinBorder;

    ws.mergeCells(5, 3 + tanggalList.length, 5, 6 + tanggalList.length);
    const totalCell = ws.getCell(5, 3 + tanggalList.length);
    totalCell.value = 'Total';
    totalCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9 };
    totalCell.fill = headerFill;
    totalCell.alignment = centerAlign;
    totalCell.border = thinBorder;

    for (let c = 1; c <= 6 + tanggalList.length; c++) {
      const cell = ws.getCell(5, c);
      cell.border = thinBorder;
      if (!cell.fill || !cell.fill.fgColor) cell.fill = headerFill;
    }
    ws.getCell(5, 1).value = '';
    ws.getCell(5, 2).value = '';

    // Row 6: day names
    headerRow2.getCell(1).value = '';
    headerRow2.getCell(2).value = '';
    tanggalList.forEach((t, idx) => {
      const cell = headerRow2.getCell(3 + idx);
      cell.value = t.hariSingkat;
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 8 };
      cell.fill = headerFill;
      cell.alignment = centerAlign;
      cell.border = thinBorder;
    });
    for (let c = 3 + tanggalList.length; c <= 6 + tanggalList.length; c++) {
      const cell = headerRow2.getCell(c);
      cell.value = '';
      cell.fill = headerFill;
      cell.border = thinBorder;
    }
    headerRow2.getCell(1).fill = headerFill;
    headerRow2.getCell(1).border = thinBorder;
    headerRow2.getCell(2).fill = headerFill;
    headerRow2.getCell(2).border = thinBorder;

    // Row 7: No, Nama, dates, H, S, I, A
    headerRow3.getCell(1).value = 'No';
    headerRow3.getCell(2).value = 'Nama';
    tanggalList.forEach((t, idx) => {
      headerRow3.getCell(3 + idx).value = t.tgl;
    });
    headerRow3.getCell(3 + tanggalList.length).value = 'H';
    headerRow3.getCell(4 + tanggalList.length).value = 'S';
    headerRow3.getCell(5 + tanggalList.length).value = 'I';
    headerRow3.getCell(6 + tanggalList.length).value = 'A';

    for (let c = 1; c <= 6 + tanggalList.length; c++) {
      const cell = headerRow3.getCell(c);
      cell.font = { bold: true, size: 9 };
      cell.alignment = centerAlign;
      cell.border = thinBorder;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
    }
    headerRow3.getCell(3 + tanggalList.length).fill = colorMap[1];
    headerRow3.getCell(4 + tanggalList.length).fill = colorMap[2];
    headerRow3.getCell(5 + tanggalList.length).fill = colorMap[3];
    headerRow3.getCell(6 + tanggalList.length).fill = colorMap[4];

    // Data rows
    rows.forEach((r, i) => {
      const rowNum = 8 + i;
      const row = ws.getRow(rowNum);
      row.getCell(1).value = i + 1;
      row.getCell(1).alignment = centerAlign;
      row.getCell(2).value = r.guru.nama_guru;

      r.harian.forEach((h, idx) => {
        const cell = row.getCell(3 + idx);
        if (!h.lewat && h.id_kehadiran) {
          cell.value = labelMap[h.id_kehadiran];
          cell.fill = colorMap[h.id_kehadiran];
        }
        cell.alignment = centerAlign;
      });

      row.getCell(3 + tanggalList.length).value = r.hadir || 0;
      row.getCell(4 + tanggalList.length).value = r.sakit || 0;
      row.getCell(5 + tanggalList.length).value = r.izin || 0;
      row.getCell(6 + tanggalList.length).value = r.alfa || 0;

      for (let c = 1; c <= 6 + tanggalList.length; c++) {
        row.getCell(c).border = thinBorder;
        if (c >= 3 + tanggalList.length) row.getCell(c).alignment = centerAlign;
      }
    });

    // Summary rows
    const summaryStart = 8 + rows.length + 1;
    const summaryData = [
      ['Jumlah guru', rekap.total],
      ['Laki-laki', rekap.laki],
      ['Perempuan', rekap.perempuan],
    ];
    summaryData.forEach(([label, val], i) => {
      const row = ws.getRow(summaryStart + i);
      row.getCell(1).value = label;
      row.getCell(2).value = val;
      row.getCell(1).font = { bold: true };
    });

    // Auto column widths
    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 22;
    tanggalList.forEach((_, idx) => { ws.getColumn(3 + idx).width = 4; });
    for (let c = 3 + tanggalList.length; c <= 6 + tanggalList.length; c++) {
      ws.getColumn(c).width = 5;
    }

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error('exportLaporanGuruExcel error:', err);
    res.status(500).json({ success: false, message: 'Gagal mengekspor laporan Excel.' });
  }
}

export async function exportLaporanGuruPdf(req, res) {
  try {
    const { bulan } = req.body;
    let payload;
    await generateLaporanGuruData({ body: { bulan } }, { json: (p) => { payload = p; }, status: () => ({ json: (p) => { payload = p; } }) });
    if (!payload?.success) return res.status(404).json({ success: false, message: 'Data tidak ditemukan.' });

    const { tanggalList, bulanLabel, rows, rekap, generalSettings } = payload.data;
    const filename = `laporan_absen_guru_${bulanLabel.replace(/\s+/g, '-')}.pdf`;

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);

    const doc = new PDFDocument({ margin: 30, size: 'A4', layout: 'landscape' });
    doc.pipe(res);

    // Header with logo
    try {
      doc.image(LOGO_PATH, 30, 30, { width: 50 });
    } catch (_) { /* logo not found, skip */ }
    doc.fontSize(14).text('DAFTAR HADIR GURU', 90, 32, { align: 'left' });
    doc.fontSize(10).text(generalSettings.school_name, 90, 50, { align: 'left' });
    doc.text(`TAHUN PELAJARAN ${generalSettings.school_year}`, 90, 64, { align: 'left' });
    doc.moveDown(1.5);
    doc.fontSize(9).text(`Bulan: ${bulanLabel}`);
    doc.moveDown(0.5);

    const startX = 30;
    let y = doc.y;
    const colNo = 25, colNama = 120, colDay = 20, colSum = 22;
    const rowH = 22;

    function drawRow(cells, widths, opts = {}) {
      let x = startX;
      cells.forEach((c, i) => {
        if (opts.bg && opts.bg[i]) {
          doc.rect(x, y, widths[i], rowH).fill(opts.bg[i]);
          doc.fillColor('black');
        }
        doc.fontSize(7).text(String(c), x + 1, y + 7, { width: widths[i] - 2, align: 'center' });
        doc.lineWidth(0.5).rect(x, y, widths[i], rowH).stroke();
        x += widths[i];
      });
      y += rowH;
    }

    const widths = [colNo, colNama, ...tanggalList.map(() => colDay), colSum, colSum, colSum, colSum];

    // Combined header: day name + date number in one cell
    let hx = startX;
    doc.rect(startX, y, colNo + colNama, rowH).fill('#4472C4');
    doc.fillColor('white').fontSize(7).text('No / Nama', startX + 1, y + 7, { width: colNo + colNama - 2, align: 'center' });
    doc.fillColor('black');
    doc.lineWidth(0.5).rect(startX, y, colNo + colNama, rowH).stroke();
    hx = startX + colNo + colNama;
    tanggalList.forEach((t) => {
      doc.rect(hx, y, colDay, rowH).fill('#4472C4');
      doc.fillColor('white').fontSize(6).text(t.hariSingkat, hx + 1, y + 3, { width: colDay - 2, align: 'center' });
      doc.fontSize(7).text(String(t.tgl), hx + 1, y + 12, { width: colDay - 2, align: 'center' });
      doc.fillColor('black');
      doc.lineWidth(0.5).rect(hx, y, colDay, rowH).stroke();
      hx += colDay;
    });
    ['H', 'S', 'I', 'A'].forEach((label, idx) => {
      const bg = ['#90ee90', '#ffff00', '#ffff00', '#ff6b6b'][idx];
      doc.rect(hx, y, colSum, rowH).fill(bg);
      doc.fillColor('black').fontSize(7).text(label, hx + 1, y + 7, { width: colSum - 2, align: 'center' });
      doc.lineWidth(0.5).rect(hx, y, colSum, rowH).stroke();
      hx += colSum;
    });
    y += rowH;

    for (const [i, r] of rows.entries()) {
      if (y > doc.page.height - 60) { doc.addPage(); y = doc.y; }
      const map = { 1: ['H', '#90ee90'], 2: ['S', '#ffff00'], 3: ['I', '#ffff00'], 4: ['A', '#ff6b6b'] };
      const harianCells = r.harian.map((h) => (h.lewat ? '' : (map[h.id_kehadiran]?.[0] || 'A')));
      const harianBg = r.harian.map((h) => (h.lewat ? null : (map[h.id_kehadiran]?.[1] || '#ff6b6b')));
      drawRow(
        [i + 1, r.guru.nama_guru, ...harianCells, r.hadir || '-', r.sakit || '-', r.izin || '-', r.alfa || '-'],
        widths,
        { bg: [null, null, ...harianBg, null, null, null, null] }
      );
    }

    doc.moveDown(1);
    doc.fontSize(9).text(`Jumlah guru: ${rekap.total}    Laki-laki: ${rekap.laki}    Perempuan: ${rekap.perempuan}`, startX, y + 10);

    // Page number
    const pageCount = doc.bufferedPageRange().count;
    for (let p = 0; p < pageCount; p++) {
      doc.switchToPage(p);
      doc.fontSize(8).text(`Halaman ${p + 1} dari ${pageCount}`, 30, doc.page.height - 30, { width: doc.page.width - 60, align: 'center' });
    }

    doc.end();
  } catch (err) {
    console.error('exportLaporanGuruPdf error:', err);
    res.status(500).json({ success: false, message: 'Gagal mengekspor laporan PDF.' });
  }
}
