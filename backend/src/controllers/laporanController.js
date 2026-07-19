import pool from '../config/db.js';
import PDFDocument from 'pdfkit';
import ExcelJS from 'exceljs';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const HARI_SINGKAT = ['Min', 'Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab'];
const BULAN_NAMA = [
  'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember',
];

function pad2(n) { return String(n).padStart(2, '0'); }
function toDateStr(d) { return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; }

function buildWorkingDays(bulanStr) {
  const [year, month] = bulanStr.split('-').map(Number);
  const lastDate = new Date(year, month, 0).getDate();
  const days = [];
  for (let d = 1; d <= lastDate; d++) {
    const date = new Date(year, month - 1, d);
    const dow = date.getDay();
    if (dow !== 0 && dow !== 6) {
      days.push(date);
    }
  }
  return days;
}

async function getGeneralSettings() {
  const [[row]] = await pool.query('SELECT * FROM general_settings LIMIT 1');
  return row || { school_name: 'SMK Harapan Bersama Tegal', school_year: '2024/2025', logo: null };
}

function getLogoPath(logoFromDb) {
  if (!logoFromDb) return null;
  const abs = path.isAbsolute(logoFromDb) ? logoFromDb : path.join(process.cwd(), logoFromDb);
  return fs.existsSync(abs) ? abs : null;
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
  let headersSent = false;
  try {
    const { kelas: idKelas, bulan } = req.body;
    let payload;
    await generateLaporanSiswaData({ body: { kelas: idKelas, bulan } }, { json: (p) => { payload = p; }, status: () => ({ json: (p) => { payload = p; } }) });

    if (!payload?.success) return res.status(404).json({ success: false, message: 'Data tidak ditemukan.' });

    const { tanggalList, bulanLabel, kelas, rows, rekap, generalSettings } = payload.data;
    const filename = `laporan_absen_${kelas.kelas.replace(/\s+/g, '_')}_${bulanLabel.replace(/\s+/g, '-')}.pdf`;
    const logoPath = getLogoPath(generalSettings.logo);

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);

    const doc = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 40 });
    doc.on('error', (err) => console.error('PDF stream error (siswa):', err));
    doc.pipe(res);
    headersSent = true;

    const pageW = doc.page.width;
    const pageH = doc.page.height;
    const margin = 40;
    const contentW = pageW - margin * 2;

    const BLUE = '#1a3a6b';
    const LIGHT_BLUE = '#e8eef7';
    const WHITE = '#ffffff';
    const GRAY = '#666666';
    const LIGHT_GRAY = '#f5f5f5';

    // ── HEADER: Logo + School Info (centered) ──
    const headerY = margin;
    const logoSize = 55;
    let textX = margin;
    let textW = contentW;

    if (logoPath) {
      try {
        doc.image(logoPath, margin, headerY, { width: logoSize });
      } catch (_) {}
      textX = margin + logoSize + 12;
      textW = contentW - logoSize - 12;
    }

    doc.font('Helvetica-Bold').fontSize(16).fillColor(BLUE)
      .text(generalSettings.school_name.toUpperCase(), textX, headerY + 2, { width: textW, align: 'center' });
    doc.font('Helvetica').fontSize(9).fillColor(GRAY)
      .text('Daftar Hadir Siswa', textX, headerY + 22, { width: textW, align: 'center' });
    doc.fontSize(9)
      .text(`Tahun Pelajaran ${generalSettings.school_year}`, textX, headerY + 34, { width: textW, align: 'center' });

    // ── Decorative line ──
    const lineY = headerY + logoSize + 10;
    doc.moveTo(margin, lineY).lineTo(pageW - margin, lineY).lineWidth(2).strokeColor(BLUE).stroke();
    doc.moveTo(margin, lineY + 1.5).lineTo(pageW - margin, lineY + 1.5).lineWidth(0.5).strokeColor(BLUE).stroke();

    // ── Info section ──
    const infoY = lineY + 12;
    doc.font('Helvetica').fontSize(9).fillColor('#333333');
    doc.text(`Bulan : ${bulanLabel}`, margin, infoY);
    doc.text(`Kelas : ${kelas.kelas}`, margin, infoY + 14);

    // ── Build table ──
    const colNo = 28;
    const colNama = 110;
    const colDay = 19;
    const colSum = 24;
    const totalCols = 2 + tanggalList.length + 4;
    const tableW = colNo + colNama + tanggalList.length * colDay + 4 * colSum;

    const startX = margin + (contentW - tableW) / 2;
    let y = infoY + 38;
    const rowH = 20;
    const headerH = 30;

    function checkPage(needH) {
      if (y + needH > pageH - margin - 50) {
        doc.addPage();
        y = margin;
      }
    }

    // ── Table header row 1: "Hari/Tanggal" merged + "Rekap" merged ──
    checkPage(headerH + rowH);
    const hdr1Y = y;

    // No + Nama header (spans 2 rows)
    doc.rect(startX, hdr1Y, colNo + colNama, headerH).fill(BLUE);
    doc.rect(startX, hdr1Y, colNo + colNama, headerH).lineWidth(0.5).strokeColor('#ffffff').stroke();
    doc.font('Helvetica-Bold').fontSize(8).fillColor(WHITE)
      .text('No / Nama', startX, hdr1Y + 9, { width: colNo + colNama, align: 'center' });

    // Hari/Tanggal header
    const htX = startX + colNo + colNama;
    const htW = tanggalList.length * colDay;
    doc.rect(htX, hdr1Y, htW, headerH / 2).fill(BLUE);
    doc.rect(htX, hdr1Y, htW, headerH / 2).lineWidth(0.5).strokeColor('#ffffff').stroke();
    doc.font('Helvetica-Bold').fontSize(7).fillColor(WHITE)
      .text('Hari / Tanggal', htX, hdr1Y + 3, { width: htW, align: 'center' });

    // Rekap header
    const rekapX = htX + htW;
    const rekapW = 4 * colSum;
    doc.rect(rekapX, hdr1Y, rekapW, headerH / 2).fill(BLUE);
    doc.rect(rekapX, hdr1Y, rekapW, headerH / 2).lineWidth(0.5).strokeColor('#ffffff').stroke();
    doc.font('Helvetica-Bold').fontSize(7).fillColor(WHITE)
      .text('Rekap', rekapX, hdr1Y + 3, { width: rekapW, align: 'center' });

    // ── Table header row 2: day names + date numbers + H S I A ──
    const hdr2Y = hdr1Y + headerH / 2;

    tanggalList.forEach((t, idx) => {
      const cx = htX + idx * colDay;
      doc.rect(cx, hdr2Y, colDay, headerH / 2).fill(LIGHT_BLUE);
      doc.rect(cx, hdr2Y, colDay, headerH / 2).lineWidth(0.5).strokeColor(BLUE).stroke();
      doc.font('Helvetica-Bold').fontSize(5.5).fillColor(BLUE)
        .text(t.hariSingkat, cx, hdr2Y + 2, { width: colDay, align: 'center' });
      doc.font('Helvetica-Bold').fontSize(7).fillColor(BLUE)
        .text(String(t.tgl), cx, hdr2Y + 10, { width: colDay, align: 'center' });
    });

    const sumColors = ['#2e7d32', '#f9a825', '#f9a825', '#c62828'];
    const sumLabels = ['H', 'S', 'I', 'A'];
    sumLabels.forEach((label, idx) => {
      const cx = rekapX + idx * colSum;
      doc.rect(cx, hdr2Y, colSum, headerH / 2).fill(sumColors[idx]);
      doc.rect(cx, hdr2Y, colSum, headerH / 2).lineWidth(0.5).strokeColor(WHITE).stroke();
      doc.font('Helvetica-Bold').fontSize(8).fillColor(WHITE)
        .text(label, cx, hdr2Y + 4, { width: colSum, align: 'center' });
    });

    y = hdr1Y + headerH;

    // ── Data rows ──
    const attColors = { 1: '#e8f5e9', 2: '#fff9c4', 3: '#fff9c4', 4: '#ffebee' };
    const attLabels = { 1: 'H', 2: 'S', 3: 'I', 4: 'A' };

    for (const [i, r] of rows.entries()) {
      checkPage(rowH);
      const ry = y;
      const isEven = i % 2 === 0;
      const rowBg = isEven ? WHITE : LIGHT_GRAY;

      // No
      doc.rect(startX, ry, colNo, rowH).fill(rowBg);
      doc.rect(startX, ry, colNo, rowH).lineWidth(0.3).strokeColor('#cccccc').stroke();
      doc.font('Helvetica').fontSize(7).fillColor('#333333')
        .text(String(i + 1), startX, ry + 5, { width: colNo, align: 'center' });

      // Nama
      doc.rect(startX + colNo, ry, colNama, rowH).fill(rowBg);
      doc.rect(startX + colNo, ry, colNama, rowH).lineWidth(0.3).strokeColor('#cccccc').stroke();
      doc.font('Helvetica').fontSize(7).fillColor('#333333')
        .text(r.siswa.nama_siswa, startX + colNo + 3, ry + 5, { width: colNama - 6, align: 'left' });

      // Attendance cells
      r.harian.forEach((h, idx) => {
        const cx = htX + idx * colDay;
        const cellBg = h.lewat ? rowBg : (attColors[h.id_kehadiran] || rowBg);
        doc.rect(cx, ry, colDay, rowH).fill(cellBg);
        doc.rect(cx, ry, colDay, rowH).lineWidth(0.3).strokeColor('#cccccc').stroke();
        if (!h.lewat && h.id_kehadiran) {
          doc.font('Helvetica').fontSize(6).fillColor('#333333')
            .text(attLabels[h.id_kehadiran], cx, ry + 5, { width: colDay, align: 'center' });
        }
      });

      // Summary columns
      const sumData = [r.hadir || 0, r.sakit || 0, r.izin || 0, r.alfa || 0];
      sumData.forEach((val, idx) => {
        const cx = rekapX + idx * colSum;
        doc.rect(cx, ry, colSum, rowH).fill(rowBg);
        doc.rect(cx, ry, colSum, rowH).lineWidth(0.3).strokeColor('#cccccc').stroke();
        doc.font('Helvetica').fontSize(7).fillColor('#333333')
          .text(String(val), cx, ry + 5, { width: colSum, align: 'center' });
      });

      y += rowH;
    }

    // ── Summary footer ──
    y += 10;
    checkPage(80);
    doc.font('Helvetica-Bold').fontSize(9).fillColor(BLUE)
      .text(`Jumlah Siswa : ${rekap.total}`, margin, y);
    doc.font('Helvetica').fontSize(9).fillColor('#333333')
      .text(`Laki-laki : ${rekap.laki}     Perempuan : ${rekap.perempuan}`, margin, y + 14);

    // ── Signature section ──
    y += 50;
    checkPage(80);
    const sigW = 180;
    const leftSigX = margin + 20;
    const rightSigX = pageW - margin - sigW - 20;

    doc.font('Helvetica').fontSize(9).fillColor('#333333');
    doc.text('Mengetahui,', leftSigX, y, { width: sigW, align: 'center' });
    doc.text('Kepala Sekolah', leftSigX, y + 12, { width: sigW, align: 'center' });
    doc.moveDown(4);
    doc.text('_________________________', leftSigX, y + 56, { width: sigW, align: 'center' });
    doc.text('NIP.', leftSigX, y + 70, { width: sigW, align: 'center' });

    doc.text(`${bulanLabel}`, rightSigX, y, { width: sigW, align: 'center' });
    doc.text('Guru Kelas', rightSigX, y + 12, { width: sigW, align: 'center' });
    doc.text('_________________________', rightSigX, y + 56, { width: sigW, align: 'center' });
    doc.text('NIP.', rightSigX, y + 70, { width: sigW, align: 'center' });

    // ── Page numbers ──
    const pageCount = doc.bufferedPageRange().count;
    for (let p = 0; p < pageCount; p++) {
      doc.switchToPage(p);
      doc.font('Helvetica').fontSize(7).fillColor(GRAY)
        .text(`Halaman ${p + 1} dari ${pageCount}`, margin, pageH - 25, { width: contentW, align: 'center' });
    }

    doc.end();
  } catch (err) {
    console.error('exportLaporanSiswaPdf error:', err);
    if (!headersSent && !res.headersSent) {
      res.status(500).json({ success: false, message: 'Gagal mengekspor laporan PDF.' });
    } else if (!res.writableEnded) {
      res.end();
    }
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
  let headersSent = false;
  try {
    const { bulan } = req.body;
    let payload;
    await generateLaporanGuruData({ body: { bulan } }, { json: (p) => { payload = p; }, status: () => ({ json: (p) => { payload = p; } }) });
    if (!payload?.success) return res.status(404).json({ success: false, message: 'Data tidak ditemukan.' });

    const { tanggalList, bulanLabel, rows, rekap, generalSettings } = payload.data;
    const filename = `laporan_absen_guru_${bulanLabel.replace(/\s+/g, '-')}.pdf`;
    const logoPath = getLogoPath(generalSettings.logo);

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);

    const doc = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 40 });
    doc.on('error', (err) => console.error('PDF stream error (guru):', err));
    doc.pipe(res);
    headersSent = true;

    const pageW = doc.page.width;
    const pageH = doc.page.height;
    const margin = 40;
    const contentW = pageW - margin * 2;

    const BLUE = '#1a3a6b';
    const LIGHT_BLUE = '#e8eef7';
    const WHITE = '#ffffff';
    const GRAY = '#666666';
    const LIGHT_GRAY = '#f5f5f5';

    // ── HEADER: Logo + School Info (centered) ──
    const headerY = margin;
    const logoSize = 55;
    let textX = margin;
    let textW = contentW;

    if (logoPath) {
      try {
        doc.image(logoPath, margin, headerY, { width: logoSize });
      } catch (_) {}
      textX = margin + logoSize + 12;
      textW = contentW - logoSize - 12;
    }

    doc.font('Helvetica-Bold').fontSize(16).fillColor(BLUE)
      .text(generalSettings.school_name.toUpperCase(), textX, headerY + 2, { width: textW, align: 'center' });
    doc.font('Helvetica').fontSize(9).fillColor(GRAY)
      .text('Daftar Hadir Guru', textX, headerY + 22, { width: textW, align: 'center' });
    doc.fontSize(9)
      .text(`Tahun Pelajaran ${generalSettings.school_year}`, textX, headerY + 34, { width: textW, align: 'center' });

    // ── Decorative line ──
    const lineY = headerY + logoSize + 10;
    doc.moveTo(margin, lineY).lineTo(pageW - margin, lineY).lineWidth(2).strokeColor(BLUE).stroke();
    doc.moveTo(margin, lineY + 1.5).lineTo(pageW - margin, lineY + 1.5).lineWidth(0.5).strokeColor(BLUE).stroke();

    // ── Info section ──
    const infoY = lineY + 12;
    doc.font('Helvetica').fontSize(9).fillColor('#333333');
    doc.text(`Bulan : ${bulanLabel}`, margin, infoY);

    // ── Build table ──
    const colNo = 28;
    const colNama = 130;
    const colDay = 19;
    const colSum = 24;
    const tableW = colNo + colNama + tanggalList.length * colDay + 4 * colSum;

    const startX = margin + (contentW - tableW) / 2;
    let y = infoY + 28;
    const rowH = 20;
    const headerH = 30;

    function checkPage(needH) {
      if (y + needH > pageH - margin - 50) {
        doc.addPage();
        y = margin;
      }
    }

    // ── Table header row 1 ──
    checkPage(headerH + rowH);
    const hdr1Y = y;

    doc.rect(startX, hdr1Y, colNo + colNama, headerH).fill(BLUE);
    doc.rect(startX, hdr1Y, colNo + colNama, headerH).lineWidth(0.5).strokeColor('#ffffff').stroke();
    doc.font('Helvetica-Bold').fontSize(8).fillColor(WHITE)
      .text('No / Nama', startX, hdr1Y + 9, { width: colNo + colNama, align: 'center' });

    const htX = startX + colNo + colNama;
    const htW = tanggalList.length * colDay;
    doc.rect(htX, hdr1Y, htW, headerH / 2).fill(BLUE);
    doc.rect(htX, hdr1Y, htW, headerH / 2).lineWidth(0.5).strokeColor('#ffffff').stroke();
    doc.font('Helvetica-Bold').fontSize(7).fillColor(WHITE)
      .text('Hari / Tanggal', htX, hdr1Y + 3, { width: htW, align: 'center' });

    const rekapX = htX + htW;
    const rekapW = 4 * colSum;
    doc.rect(rekapX, hdr1Y, rekapW, headerH / 2).fill(BLUE);
    doc.rect(rekapX, hdr1Y, rekapW, headerH / 2).lineWidth(0.5).strokeColor('#ffffff').stroke();
    doc.font('Helvetica-Bold').fontSize(7).fillColor(WHITE)
      .text('Rekap', rekapX, hdr1Y + 3, { width: rekapW, align: 'center' });

    // ── Table header row 2 ──
    const hdr2Y = hdr1Y + headerH / 2;

    tanggalList.forEach((t, idx) => {
      const cx = htX + idx * colDay;
      doc.rect(cx, hdr2Y, colDay, headerH / 2).fill(LIGHT_BLUE);
      doc.rect(cx, hdr2Y, colDay, headerH / 2).lineWidth(0.5).strokeColor(BLUE).stroke();
      doc.font('Helvetica-Bold').fontSize(5.5).fillColor(BLUE)
        .text(t.hariSingkat, cx, hdr2Y + 2, { width: colDay, align: 'center' });
      doc.font('Helvetica-Bold').fontSize(7).fillColor(BLUE)
        .text(String(t.tgl), cx, hdr2Y + 10, { width: colDay, align: 'center' });
    });

    const sumColors = ['#2e7d32', '#f9a825', '#f9a825', '#c62828'];
    const sumLabels = ['H', 'S', 'I', 'A'];
    sumLabels.forEach((label, idx) => {
      const cx = rekapX + idx * colSum;
      doc.rect(cx, hdr2Y, colSum, headerH / 2).fill(sumColors[idx]);
      doc.rect(cx, hdr2Y, colSum, headerH / 2).lineWidth(0.5).strokeColor(WHITE).stroke();
      doc.font('Helvetica-Bold').fontSize(8).fillColor(WHITE)
        .text(label, cx, hdr2Y + 4, { width: colSum, align: 'center' });
    });

    y = hdr1Y + headerH;

    // ── Data rows ──
    const attColors = { 1: '#e8f5e9', 2: '#fff9c4', 3: '#fff9c4', 4: '#ffebee' };
    const attLabels = { 1: 'H', 2: 'S', 3: 'I', 4: 'A' };

    for (const [i, r] of rows.entries()) {
      checkPage(rowH);
      const ry = y;
      const isEven = i % 2 === 0;
      const rowBg = isEven ? WHITE : LIGHT_GRAY;

      doc.rect(startX, ry, colNo, rowH).fill(rowBg);
      doc.rect(startX, ry, colNo, rowH).lineWidth(0.3).strokeColor('#cccccc').stroke();
      doc.font('Helvetica').fontSize(7).fillColor('#333333')
        .text(String(i + 1), startX, ry + 5, { width: colNo, align: 'center' });

      doc.rect(startX + colNo, ry, colNama, rowH).fill(rowBg);
      doc.rect(startX + colNo, ry, colNama, rowH).lineWidth(0.3).strokeColor('#cccccc').stroke();
      doc.font('Helvetica').fontSize(7).fillColor('#333333')
        .text(r.guru.nama_guru, startX + colNo + 3, ry + 5, { width: colNama - 6, align: 'left' });

      r.harian.forEach((h, idx) => {
        const cx = htX + idx * colDay;
        const cellBg = h.lewat ? rowBg : (attColors[h.id_kehadiran] || rowBg);
        doc.rect(cx, ry, colDay, rowH).fill(cellBg);
        doc.rect(cx, ry, colDay, rowH).lineWidth(0.3).strokeColor('#cccccc').stroke();
        if (!h.lewat && h.id_kehadiran) {
          doc.font('Helvetica').fontSize(6).fillColor('#333333')
            .text(attLabels[h.id_kehadiran], cx, ry + 5, { width: colDay, align: 'center' });
        }
      });

      const sumData = [r.hadir || 0, r.sakit || 0, r.izin || 0, r.alfa || 0];
      sumData.forEach((val, idx) => {
        const cx = rekapX + idx * colSum;
        doc.rect(cx, ry, colSum, rowH).fill(rowBg);
        doc.rect(cx, ry, colSum, rowH).lineWidth(0.3).strokeColor('#cccccc').stroke();
        doc.font('Helvetica').fontSize(7).fillColor('#333333')
          .text(String(val), cx, ry + 5, { width: colSum, align: 'center' });
      });

      y += rowH;
    }

    // ── Summary footer ──
    y += 10;
    checkPage(80);
    doc.font('Helvetica-Bold').fontSize(9).fillColor(BLUE)
      .text(`Jumlah Guru : ${rekap.total}`, margin, y);
    doc.font('Helvetica').fontSize(9).fillColor('#333333')
      .text(`Laki-laki : ${rekap.laki}     Perempuan : ${rekap.perempuan}`, margin, y + 14);

    // ── Signature section ──
    y += 50;
    checkPage(80);
    const sigW = 180;
    const leftSigX = margin + 20;
    const rightSigX = pageW - margin - sigW - 20;

    doc.font('Helvetica').fontSize(9).fillColor('#333333');
    doc.text('Mengetahui,', leftSigX, y, { width: sigW, align: 'center' });
    doc.text('Kepala Sekolah', leftSigX, y + 12, { width: sigW, align: 'center' });
    doc.text('_________________________', leftSigX, y + 56, { width: sigW, align: 'center' });
    doc.text('NIP.', leftSigX, y + 70, { width: sigW, align: 'center' });

    doc.text(`${bulanLabel}`, rightSigX, y, { width: sigW, align: 'center' });
    doc.text('Guru', rightSigX, y + 12, { width: sigW, align: 'center' });
    doc.text('_________________________', rightSigX, y + 56, { width: sigW, align: 'center' });
    doc.text('NIP.', rightSigX, y + 70, { width: sigW, align: 'center' });

    // ── Page numbers ──
    const pageCount = doc.bufferedPageRange().count;
    for (let p = 0; p < pageCount; p++) {
      doc.switchToPage(p);
      doc.font('Helvetica').fontSize(7).fillColor(GRAY)
        .text(`Halaman ${p + 1} dari ${pageCount}`, margin, pageH - 25, { width: contentW, align: 'center' });
    }

    doc.end();
  } catch (err) {
    console.error('exportLaporanGuruPdf error:', err);
    if (!headersSent && !res.headersSent) {
      res.status(500).json({ success: false, message: 'Gagal mengekspor laporan PDF.' });
    } else if (!res.writableEnded) {
      res.end();
    }
  }
}
