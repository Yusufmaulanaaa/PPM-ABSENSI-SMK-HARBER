import pool from '../config/db.js';
import PDFDocument from 'pdfkit';
import ExcelJS from 'exceljs';
import { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun, WidthType, AlignmentType, BorderStyle, ImageRun, VerticalAlign, PageOrientation } from 'docx';
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
  if (fs.existsSync(abs)) return abs;
  const uploadsDir = path.join(process.cwd(), 'uploads', 'logo');
  if (fs.existsSync(uploadsDir)) {
    const files = fs.readdirSync(uploadsDir);
    if (files.length > 0) return path.join(uploadsDir, files[0]);
  }
  return null;
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

    const logoPathXlsx = getLogoPath(generalSettings.logo);
    let logoColEnd = 1;
    if (logoPathXlsx) {
      try {
        const buf = fs.readFileSync(logoPathXlsx);
        const ext = logoPathXlsx.split('.').pop().toLowerCase();
        const imgExt = ext === 'jpg' ? 'jpeg' : ext;
        if (imgExt === 'jpeg' || imgExt === 'png' || imgExt === 'gif') {
          const imgId = workbook.addImage({ buffer: buf, extension: imgExt });
          ws.addImage(imgId, { tl: { col: 0, row: 0 }, ext: { width: 80, height: 80 } });
          logoColEnd = 2;
        }
      } catch (err) {
        console.error('Logo Excel read error:', err.message);
      }
    }

    const HEADER_BG = 'FF1A3A6B';
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

    const totalCols = 6 + tanggalList.length;
    const headerStartCol = logoColEnd;
    ws.mergeCells(1, headerStartCol, 1, totalCols);
    const titleCell = ws.getCell(1, headerStartCol);
    titleCell.value = generalSettings.school_name.toUpperCase();
    titleCell.font = { bold: true, size: 14, color: { argb: 'FF1A3A6B' } };
    titleCell.alignment = { horizontal: 'center' };

    ws.mergeCells(2, headerStartCol, 2, totalCols);
    const subCell = ws.getCell(2, headerStartCol);
    subCell.value = 'DAFTAR HADIR SISWA';
    subCell.font = { bold: true, size: 12 };
    subCell.alignment = { horizontal: 'center' };

    ws.mergeCells(3, headerStartCol, 3, totalCols);
    ws.getCell(3, headerStartCol).value = `Tahun Pelajaran ${generalSettings.school_year}`;
    ws.getCell(3, headerStartCol).font = { bold: true, size: 11 };
    ws.getCell(3, headerStartCol).alignment = { horizontal: 'center' };

    ws.mergeCells(4, headerStartCol, 4, totalCols);
    ws.getCell(4, headerStartCol).value = `Bulan: ${bulanLabel}    Kelas: ${kelas.kelas}`;
    ws.getCell(4, headerStartCol).font = { size: 10, italic: true };
    ws.getCell(4, headerStartCol).alignment = { horizontal: 'center' };

    // Table header row (row 6)
    const headerRow1 = ws.getRow(6);
    const headerRow2 = ws.getRow(7);
    const headerRow3 = ws.getRow(8);

    // Row 6: blank, blank, "Hari/Tanggal" merged, blank, blank, blank, blank
    ws.mergeCells(6, 3, 6, 2 + tanggalList.length);
    const hariTanggalCell = ws.getCell(6, 3);
    hariTanggalCell.value = 'Hari/Tanggal';
    hariTanggalCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9 };
    hariTanggalCell.fill = headerFill;
    hariTanggalCell.alignment = centerAlign;
    hariTanggalCell.border = thinBorder;

    ws.mergeCells(6, 3 + tanggalList.length, 6, 6 + tanggalList.length);
    const totalCell = ws.getCell(6, 3 + tanggalList.length);
    totalCell.value = 'Total';
    totalCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9 };
    totalCell.fill = headerFill;
    totalCell.alignment = centerAlign;
    totalCell.border = thinBorder;

    // Fill borders on row 6 for merged cells
    for (let c = 1; c <= 6 + tanggalList.length; c++) {
      const cell = ws.getCell(6, c);
      cell.border = thinBorder;
      if (!cell.fill || !cell.fill.fgColor) {
        cell.fill = headerFill;
      }
    }
    ws.getCell(6, 1).value = '';
    ws.getCell(6, 2).value = '';

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
      const rowNum = 9 + i;
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
    const summaryStart = 9 + rows.length + 1;
    const summaryData = [
      ['Jumlah siswa', rekap.total],
      ['Laki-laki', rekap.laki],
      ['Perempuan', rekap.perempuan],
    ];
    summaryData.forEach(([label, val], i) => {
      const row = ws.getRow(summaryStart + i);
      row.getCell(1).value = label;
      row.getCell(2).value = val;
      row.getCell(1).font = { bold: true, size: 11 };
      row.getCell(2).font = { bold: true, size: 11 };
    });

    const bpSign = bulanLabel.split(' ');
    const tglCetak = new Date().getDate();
    const sigStart = summaryStart + 4;
    ws.getRow(sigStart).getCell(1).value = 'Mengetahui,';
    ws.getRow(sigStart).getCell(1).font = { bold: true, size: 11 };
    ws.getRow(sigStart + 1).getCell(1).value = 'Kepala Sekolah';
    ws.getRow(sigStart + 1).getCell(1).font = { bold: true, size: 11 };
    ws.getRow(sigStart + 5).getCell(1).value = '_________________________';
    ws.getRow(sigStart + 6).getCell(1).value = 'NIP.';
    ws.getRow(sigStart + 6).getCell(1).font = { size: 10 };

    ws.getRow(sigStart).getCell(4).value = `Tegal, ${tglCetak} ${bpSign[0] || ''} ${bpSign[1] || ''}`;
    ws.getRow(sigStart).getCell(4).font = { bold: true, size: 11 };
    ws.getRow(sigStart + 1).getCell(4).value = 'Guru Kelas';
    ws.getRow(sigStart + 1).getCell(4).font = { bold: true, size: 11 };
    ws.getRow(sigStart + 5).getCell(4).value = '_________________________';
    ws.getRow(sigStart + 6).getCell(4).value = 'NIP.';
    ws.getRow(sigStart + 6).getCell(4).font = { size: 10 };

    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 22;
    tanggalList.forEach((_, idx) => {
      ws.getColumn(3 + idx).width = 4;
    });
    for (let c = 3 + tanggalList.length; c <= 6 + tanggalList.length; c++) {
      ws.getColumn(c).width = 5;
    }

    ws.views = [{ state: 'frozen', ySplit: 8, activeCell: 'A9' }];
    ws.pageSetup = { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 };

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

    if (!payload?.success) return res.status(404).json({ success: false, message: payload?.message || 'Data tidak ditemukan.' });

    const { tanggalList, bulanLabel, kelas, rows, rekap, generalSettings } = payload.data;
    const filename = `laporan_absen_siswa_${kelas.kelas.replace(/\s+/g, '_')}_${bulanLabel.replace(/\s+/g, '-')}.pdf`;

    let logoBuffer = null;
    try {
      const lp = getLogoPath(generalSettings.logo);
      if (lp) logoBuffer = fs.readFileSync(lp);
    } catch (_) {}

    const pdfBuffer = await new Promise((resolve, reject) => {
      const doc = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 35, bufferPages: true });
      const chunks = [];
      doc.on('data', (chunk) => chunks.push(chunk));
      doc.on('end', () => resolve(Buffer.concat(chunks)));
      doc.on('error', (err) => reject(err));

      try {
        const pageW = doc.page.width;
        const pageH = doc.page.height;
        const margin = 35;
        const contentW = pageW - margin * 2;

        const BLUE = '#1a3a6b';
        const DARK = '#222222';
        const LIGHT_BLUE = '#dde4f0';
        const WHITE = '#ffffff';
        const GRAY = '#888888';
        const LIGHT_GRAY = '#f4f4f4';

        let currentPage = 1;

        function drawFooter() {
          doc.save();
          doc.font('Helvetica').fontSize(7).fillColor(GRAY);
          doc.text(`Halaman ${currentPage}`, margin, pageH - 22, { width: contentW, align: 'center' });
          doc.restore();
        }

        function drawHeader(isFirstPage) {
          const headerY = margin;
          const logoSize = 60;
          let textX = margin;
          let textW = contentW;

          if (logoBuffer) {
            try {
              doc.image(logoBuffer, margin, headerY, { width: logoSize, height: logoSize });
              textX = margin + logoSize + 12;
              textW = contentW - logoSize - 12;
            } catch (_) {}
          }

          doc.font('Helvetica-Bold').fontSize(15).fillColor(BLUE)
            .text(generalSettings.school_name.toUpperCase(), textX, headerY, { width: textW, align: 'center' });
          doc.font('Helvetica').fontSize(8).fillColor(DARK)
            .text('Sistem Absensi Sekolah', textX, headerY + 18, { width: textW, align: 'center' });

          const lineY = headerY + logoSize + 6;
          doc.moveTo(margin, lineY).lineTo(pageW - margin, lineY).lineWidth(2).strokeColor(BLUE).stroke();
          doc.moveTo(margin, lineY + 1.5).lineTo(pageW - margin, lineY + 1.5).lineWidth(0.5).strokeColor(BLUE).stroke();

          return lineY + 8;
        }

        let y = drawHeader(true);

        doc.font('Helvetica-Bold').fontSize(12).fillColor(DARK)
          .text('LAPORAN DAFTAR HADIR SISWA', margin, y, { width: contentW, align: 'center' });
        y += 18;
        doc.font('Helvetica').fontSize(10).fillColor(GRAY)
          .text(`Tahun Pelajaran ${generalSettings.school_year}`, margin, y, { width: contentW, align: 'center' });
        y += 20;

        doc.font('Helvetica-Bold').fontSize(9).fillColor(DARK);
        doc.text(`Bulan  : ${bulanLabel}`, margin, y);
        doc.text(`Kelas  : ${kelas.kelas}`, margin, y + 14);
        y += 32;

        const colNo = 26, colNama = 105, colDay = 18, colSum = 22;
        const tableW = colNo + colNama + tanggalList.length * colDay + 4 * colSum;
        const startX = margin + (contentW - tableW) / 2;
        const rowH = 17, headerH = 26;

        function checkPage(needH) {
          if (y + needH > pageH - margin - 40) {
            drawFooter();
            doc.addPage();
            currentPage++;
            y = drawHeader(false);
          }
        }

        checkPage(headerH + rowH);
        const hdr1Y = y;

        doc.rect(startX, hdr1Y, colNo + colNama, headerH).fill(BLUE);
        doc.rect(startX, hdr1Y, colNo + colNama, headerH).lineWidth(0.5).strokeColor(BLUE).stroke();
        doc.font('Helvetica-Bold').fontSize(7).fillColor(WHITE)
          .text('No / Nama', startX, hdr1Y + 8, { width: colNo + colNama, align: 'center' });

        const htX = startX + colNo + colNama;
        const htW = tanggalList.length * colDay;
        doc.rect(htX, hdr1Y, htW, headerH / 2).fill(BLUE);
        doc.rect(htX, hdr1Y, htW, headerH / 2).lineWidth(0.5).strokeColor(WHITE).stroke();
        doc.font('Helvetica-Bold').fontSize(6).fillColor(WHITE)
          .text('Hari / Tanggal', htX, hdr1Y + 2, { width: htW, align: 'center' });

        const rekapX = htX + htW;
        const rekapW = 4 * colSum;
        doc.rect(rekapX, hdr1Y, rekapW, headerH / 2).fill(BLUE);
        doc.rect(rekapX, hdr1Y, rekapW, headerH / 2).lineWidth(0.5).strokeColor(WHITE).stroke();
        doc.font('Helvetica-Bold').fontSize(6).fillColor(WHITE)
          .text('Rekap', rekapX, hdr1Y + 2, { width: rekapW, align: 'center' });

        const hdr2Y = hdr1Y + headerH / 2;
        tanggalList.forEach((t, idx) => {
          const cx = htX + idx * colDay;
          doc.rect(cx, hdr2Y, colDay, headerH / 2).fill(LIGHT_BLUE);
          doc.rect(cx, hdr2Y, colDay, headerH / 2).lineWidth(0.5).strokeColor(BLUE).stroke();
          doc.font('Helvetica-Bold').fontSize(5).fillColor(BLUE)
            .text(t.hariSingkat, cx, hdr2Y + 1, { width: colDay, align: 'center' });
          doc.font('Helvetica-Bold').fontSize(6).fillColor(BLUE)
            .text(String(t.tgl), cx, hdr2Y + 8, { width: colDay, align: 'center' });
        });

        const sumColors = ['#2e7d32', '#e8a800', '#e8a800', '#c62828'];
        ['H', 'S', 'I', 'A'].forEach((label, idx) => {
          const cx = rekapX + idx * colSum;
          doc.rect(cx, hdr2Y, colSum, headerH / 2).fill(sumColors[idx]);
          doc.rect(cx, hdr2Y, colSum, headerH / 2).lineWidth(0.5).strokeColor(WHITE).stroke();
          doc.font('Helvetica-Bold').fontSize(7).fillColor(WHITE)
            .text(label, cx, hdr2Y + 3, { width: colSum, align: 'center' });
        });

        y = hdr1Y + headerH;

        const attColors = { 1: '#e8f5e9', 2: '#fff9c4', 3: '#fff9c4', 4: '#ffebee' };
        const attLabels = { 1: 'H', 2: 'S', 3: 'I', 4: 'A' };

        for (const [i, r] of rows.entries()) {
          checkPage(rowH);
          const ry = y;
          const rowBg = i % 2 === 0 ? WHITE : LIGHT_GRAY;

          doc.rect(startX, ry, colNo, rowH).fill(rowBg);
          doc.rect(startX, ry, colNo, rowH).lineWidth(0.3).strokeColor('#d0d0d0').stroke();
          doc.font('Helvetica').fontSize(6).fillColor(DARK)
            .text(String(i + 1), startX, ry + 4, { width: colNo, align: 'center' });

          doc.rect(startX + colNo, ry, colNama, rowH).fill(rowBg);
          doc.rect(startX + colNo, ry, colNama, rowH).lineWidth(0.3).strokeColor('#d0d0d0').stroke();
          doc.font('Helvetica').fontSize(6).fillColor(DARK)
            .text(r.siswa.nama_siswa, startX + colNo + 3, ry + 4, { width: colNama - 6, align: 'left' });

          r.harian.forEach((h, idx) => {
            const cx = htX + idx * colDay;
            const cellBg = h.lewat ? rowBg : (attColors[h.id_kehadiran] || rowBg);
            doc.rect(cx, ry, colDay, rowH).fill(cellBg);
            doc.rect(cx, ry, colDay, rowH).lineWidth(0.3).strokeColor('#d0d0d0').stroke();
            if (!h.lewat && h.id_kehadiran) {
              doc.font('Helvetica').fontSize(5.5).fillColor(DARK)
                .text(attLabels[h.id_kehadiran], cx, ry + 4, { width: colDay, align: 'center' });
            }
          });

          [r.hadir || 0, r.sakit || 0, r.izin || 0, r.alfa || 0].forEach((val, idx) => {
            const cx = rekapX + idx * colSum;
            doc.rect(cx, ry, colSum, rowH).fill(rowBg);
            doc.rect(cx, ry, colSum, rowH).lineWidth(0.3).strokeColor('#d0d0d0').stroke();
            doc.font('Helvetica').fontSize(6).fillColor(DARK)
              .text(String(val), cx, ry + 4, { width: colSum, align: 'center' });
          });

          y += rowH;
        }

        y += 12;
        checkPage(80);
        doc.font('Helvetica-Bold').fontSize(9).fillColor(BLUE)
          .text(`Jumlah Siswa : ${rekap.total}`, margin, y);
        doc.font('Helvetica').fontSize(9).fillColor(DARK)
          .text(`Laki-laki : ${rekap.laki}     Perempuan : ${rekap.perempuan}`, margin, y + 14);

        y += 45;
        checkPage(100);
        const sigW = 220;
        const leftSigX = margin + 20;
        const rightSigX = pageW - margin - sigW - 20;

        const bulanParts = bulanLabel.split(' ');
        const namaBulan = bulanParts[0] || '';
        const tahunBulan = bulanParts[1] || '';
        const tanggalCetak = new Date().getDate();

        doc.font('Helvetica').fontSize(9).fillColor(DARK);
        doc.text('Mengetahui,', leftSigX, y, { width: sigW, align: 'center' });
        doc.text('Kepala Sekolah', leftSigX, y + 13, { width: sigW, align: 'center' });
        doc.text('', leftSigX, y + 26, { width: sigW, align: 'center' });
        doc.text('', leftSigX, y + 39, { width: sigW, align: 'center' });
        doc.text('', leftSigX, y + 52, { width: sigW, align: 'center' });
        doc.font('Helvetica').fontSize(9).fillColor(DARK)
          .text('_________________________', leftSigX, y + 65, { width: sigW, align: 'center' });
        doc.fontSize(8).text('NIP.', leftSigX, y + 78, { width: sigW, align: 'center' });

        doc.font('Helvetica').fontSize(9).fillColor(DARK);
        doc.text(`Tegal, ${tanggalCetak} ${namaBulan} ${tahunBulan}`, rightSigX, y, { width: sigW, align: 'center' });
        doc.text('Guru Kelas', rightSigX, y + 13, { width: sigW, align: 'center' });
        doc.text('', rightSigX, y + 26, { width: sigW, align: 'center' });
        doc.text('', rightSigX, y + 39, { width: sigW, align: 'center' });
        doc.text('', rightSigX, y + 52, { width: sigW, align: 'center' });
        doc.font('Helvetica').fontSize(9).fillColor(DARK)
          .text('_________________________', rightSigX, y + 65, { width: sigW, align: 'center' });
        doc.fontSize(8).text('NIP.', rightSigX, y + 78, { width: sigW, align: 'center' });

        drawFooter();
        doc.end();
      } catch (err) {
        reject(err);
      }
    });

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', pdfBuffer.length);
    res.end(pdfBuffer);
  } catch (err) {
    console.error('exportLaporanSiswaPdf error:', err);
    if (!res.headersSent) {
      res.status(500).json({ success: false, message: `Gagal mengekspor PDF: ${err.message}` });
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

    const logoPathXlsx = getLogoPath(generalSettings.logo);
    let logoColEnd = 1;
    if (logoPathXlsx) {
      try {
        const buf = fs.readFileSync(logoPathXlsx);
        const ext = logoPathXlsx.split('.').pop().toLowerCase();
        const imgExt = ext === 'jpg' ? 'jpeg' : ext;
        if (imgExt === 'jpeg' || imgExt === 'png' || imgExt === 'gif') {
          const imgId = workbook.addImage({ buffer: buf, extension: imgExt });
          ws.addImage(imgId, { tl: { col: 0, row: 0 }, ext: { width: 80, height: 80 } });
          logoColEnd = 2;
        }
      } catch (err) {
        console.error('Logo Excel read error:', err.message);
      }
    }

    const HEADER_BG = 'FF1A3A6B';
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

    const totalCols = 6 + tanggalList.length;
    const headerStartCol = logoColEnd;
    ws.mergeCells(1, headerStartCol, 1, totalCols);
    const titleCell = ws.getCell(1, headerStartCol);
    titleCell.value = generalSettings.school_name.toUpperCase();
    titleCell.font = { bold: true, size: 14, color: { argb: 'FF1A3A6B' } };
    titleCell.alignment = { horizontal: 'center' };

    ws.mergeCells(2, headerStartCol, 2, totalCols);
    const subCell = ws.getCell(2, headerStartCol);
    subCell.value = 'DAFTAR HADIR GURU';
    subCell.font = { bold: true, size: 12 };
    subCell.alignment = { horizontal: 'center' };

    ws.mergeCells(3, headerStartCol, 3, totalCols);
    ws.getCell(3, headerStartCol).value = `Tahun Pelajaran ${generalSettings.school_year}`;
    ws.getCell(3, headerStartCol).font = { bold: true, size: 11 };
    ws.getCell(3, headerStartCol).alignment = { horizontal: 'center' };

    ws.mergeCells(4, headerStartCol, 4, totalCols);
    ws.getCell(4, headerStartCol).value = `Bulan: ${bulanLabel}`;
    ws.getCell(4, headerStartCol).font = { size: 10, italic: true };
    ws.getCell(4, headerStartCol).alignment = { horizontal: 'center' };

    // Table header row (row 6)
    const headerRow2 = ws.getRow(7);
    const headerRow3 = ws.getRow(8);

    ws.mergeCells(6, 3, 6, 2 + tanggalList.length);
    const hariTanggalCell = ws.getCell(6, 3);
    hariTanggalCell.value = 'Hari/Tanggal';
    hariTanggalCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9 };
    hariTanggalCell.fill = headerFill;
    hariTanggalCell.alignment = centerAlign;
    hariTanggalCell.border = thinBorder;

    ws.mergeCells(6, 3 + tanggalList.length, 6, 6 + tanggalList.length);
    const totalCell = ws.getCell(6, 3 + tanggalList.length);
    totalCell.value = 'Total';
    totalCell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9 };
    totalCell.fill = headerFill;
    totalCell.alignment = centerAlign;
    totalCell.border = thinBorder;

    for (let c = 1; c <= 6 + tanggalList.length; c++) {
      const cell = ws.getCell(6, c);
      cell.border = thinBorder;
      if (!cell.fill || !cell.fill.fgColor) cell.fill = headerFill;
    }
    ws.getCell(6, 1).value = '';
    ws.getCell(6, 2).value = '';

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
      const rowNum = 9 + i;
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
    const summaryStart = 9 + rows.length + 1;
    const summaryData = [
      ['Jumlah guru', rekap.total],
      ['Laki-laki', rekap.laki],
      ['Perempuan', rekap.perempuan],
    ];
    summaryData.forEach(([label, val], i) => {
      const row = ws.getRow(summaryStart + i);
      row.getCell(1).value = label;
      row.getCell(2).value = val;
      row.getCell(1).font = { bold: true, size: 11 };
      row.getCell(2).font = { bold: true, size: 11 };
    });

    const bpSign = bulanLabel.split(' ');
    const tglCetak = new Date().getDate();
    const sigStart = summaryStart + 4;
    ws.getRow(sigStart).getCell(1).value = 'Mengetahui,';
    ws.getRow(sigStart).getCell(1).font = { bold: true, size: 11 };
    ws.getRow(sigStart + 1).getCell(1).value = 'Kepala Sekolah';
    ws.getRow(sigStart + 1).getCell(1).font = { bold: true, size: 11 };
    ws.getRow(sigStart + 5).getCell(1).value = '_________________________';
    ws.getRow(sigStart + 6).getCell(1).value = 'NIP.';
    ws.getRow(sigStart + 6).getCell(1).font = { size: 10 };

    ws.getRow(sigStart).getCell(4).value = `Tegal, ${tglCetak} ${bpSign[0] || ''} ${bpSign[1] || ''}`;
    ws.getRow(sigStart).getCell(4).font = { bold: true, size: 11 };
    ws.getRow(sigStart + 1).getCell(4).value = 'Guru';
    ws.getRow(sigStart + 1).getCell(4).font = { bold: true, size: 11 };
    ws.getRow(sigStart + 5).getCell(4).value = '_________________________';
    ws.getRow(sigStart + 6).getCell(4).value = 'NIP.';
    ws.getRow(sigStart + 6).getCell(4).font = { size: 10 };

    ws.getColumn(1).width = 6;
    ws.getColumn(2).width = 22;
    tanggalList.forEach((_, idx) => { ws.getColumn(3 + idx).width = 4; });
    for (let c = 3 + tanggalList.length; c <= 6 + tanggalList.length; c++) {
      ws.getColumn(c).width = 5;
    }

    ws.views = [{ state: 'frozen', ySplit: 8, activeCell: 'A9' }];
    ws.pageSetup = { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 };

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
    if (!payload?.success) return res.status(404).json({ success: false, message: payload?.message || 'Data tidak ditemukan.' });

    const { tanggalList, bulanLabel, rows, rekap, generalSettings } = payload.data;
    const filename = `laporan_absen_guru_${bulanLabel.replace(/\s+/g, '-')}.pdf`;

    let logoBuffer = null;
    try {
      const lp = getLogoPath(generalSettings.logo);
      if (lp) logoBuffer = fs.readFileSync(lp);
    } catch (_) {}

    const pdfBuffer = await new Promise((resolve, reject) => {
      const doc = new PDFDocument({ size: 'A4', layout: 'landscape', margin: 35, bufferPages: true });
      const chunks = [];
      doc.on('data', (chunk) => chunks.push(chunk));
      doc.on('end', () => resolve(Buffer.concat(chunks)));
      doc.on('error', (err) => reject(err));

      try {
        const pageW = doc.page.width;
        const pageH = doc.page.height;
        const margin = 35;
        const contentW = pageW - margin * 2;

        const BLUE = '#1a3a6b';
        const DARK = '#222222';
        const LIGHT_BLUE = '#dde4f0';
        const WHITE = '#ffffff';
        const GRAY = '#888888';
        const LIGHT_GRAY = '#f4f4f4';

        let currentPage = 1;

        function drawFooter() {
          doc.save();
          doc.font('Helvetica').fontSize(7).fillColor(GRAY);
          doc.text(`Halaman ${currentPage}`, margin, pageH - 22, { width: contentW, align: 'center' });
          doc.restore();
        }

        function drawHeader() {
          const headerY = margin;
          const logoSize = 60;
          let textX = margin;
          let textW = contentW;

          if (logoBuffer) {
            try {
              doc.image(logoBuffer, margin, headerY, { width: logoSize, height: logoSize });
              textX = margin + logoSize + 12;
              textW = contentW - logoSize - 12;
            } catch (_) {}
          }

          doc.font('Helvetica-Bold').fontSize(15).fillColor(BLUE)
            .text(generalSettings.school_name.toUpperCase(), textX, headerY, { width: textW, align: 'center' });
          doc.font('Helvetica').fontSize(8).fillColor(DARK)
            .text('Sistem Absensi Sekolah', textX, headerY + 18, { width: textW, align: 'center' });

          const lineY = headerY + logoSize + 6;
          doc.moveTo(margin, lineY).lineTo(pageW - margin, lineY).lineWidth(2).strokeColor(BLUE).stroke();
          doc.moveTo(margin, lineY + 1.5).lineTo(pageW - margin, lineY + 1.5).lineWidth(0.5).strokeColor(BLUE).stroke();

          return lineY + 8;
        }

        let y = drawHeader();

        doc.font('Helvetica-Bold').fontSize(12).fillColor(DARK)
          .text('LAPORAN DAFTAR HADIR GURU', margin, y, { width: contentW, align: 'center' });
        y += 18;
        doc.font('Helvetica').fontSize(10).fillColor(GRAY)
          .text(`Tahun Pelajaran ${generalSettings.school_year}`, margin, y, { width: contentW, align: 'center' });
        y += 20;

        doc.font('Helvetica-Bold').fontSize(9).fillColor(DARK);
        doc.text(`Bulan  : ${bulanLabel}`, margin, y);
        y += 20;

        const colNo = 26, colNama = 120, colDay = 18, colSum = 22;
        const tableW = colNo + colNama + tanggalList.length * colDay + 4 * colSum;
        const startX = margin + (contentW - tableW) / 2;
        const rowH = 17, headerH = 26;

        function checkPage(needH) {
          if (y + needH > pageH - margin - 40) {
            drawFooter();
            doc.addPage();
            currentPage++;
            y = drawHeader();
          }
        }

        checkPage(headerH + rowH);
        const hdr1Y = y;

        doc.rect(startX, hdr1Y, colNo + colNama, headerH).fill(BLUE);
        doc.rect(startX, hdr1Y, colNo + colNama, headerH).lineWidth(0.5).strokeColor(BLUE).stroke();
        doc.font('Helvetica-Bold').fontSize(7).fillColor(WHITE)
          .text('No / Nama', startX, hdr1Y + 8, { width: colNo + colNama, align: 'center' });

        const htX = startX + colNo + colNama;
        const htW = tanggalList.length * colDay;
        doc.rect(htX, hdr1Y, htW, headerH / 2).fill(BLUE);
        doc.rect(htX, hdr1Y, htW, headerH / 2).lineWidth(0.5).strokeColor(WHITE).stroke();
        doc.font('Helvetica-Bold').fontSize(6).fillColor(WHITE)
          .text('Hari / Tanggal', htX, hdr1Y + 2, { width: htW, align: 'center' });

        const rekapX = htX + htW;
        const rekapW = 4 * colSum;
        doc.rect(rekapX, hdr1Y, rekapW, headerH / 2).fill(BLUE);
        doc.rect(rekapX, hdr1Y, rekapW, headerH / 2).lineWidth(0.5).strokeColor(WHITE).stroke();
        doc.font('Helvetica-Bold').fontSize(6).fillColor(WHITE)
          .text('Rekap', rekapX, hdr1Y + 2, { width: rekapW, align: 'center' });

        const hdr2Y = hdr1Y + headerH / 2;
        tanggalList.forEach((t, idx) => {
          const cx = htX + idx * colDay;
          doc.rect(cx, hdr2Y, colDay, headerH / 2).fill(LIGHT_BLUE);
          doc.rect(cx, hdr2Y, colDay, headerH / 2).lineWidth(0.5).strokeColor(BLUE).stroke();
          doc.font('Helvetica-Bold').fontSize(5).fillColor(BLUE)
            .text(t.hariSingkat, cx, hdr2Y + 1, { width: colDay, align: 'center' });
          doc.font('Helvetica-Bold').fontSize(6).fillColor(BLUE)
            .text(String(t.tgl), cx, hdr2Y + 8, { width: colDay, align: 'center' });
        });

        const sumColors = ['#2e7d32', '#e8a800', '#e8a800', '#c62828'];
        ['H', 'S', 'I', 'A'].forEach((label, idx) => {
          const cx = rekapX + idx * colSum;
          doc.rect(cx, hdr2Y, colSum, headerH / 2).fill(sumColors[idx]);
          doc.rect(cx, hdr2Y, colSum, headerH / 2).lineWidth(0.5).strokeColor(WHITE).stroke();
          doc.font('Helvetica-Bold').fontSize(7).fillColor(WHITE)
            .text(label, cx, hdr2Y + 3, { width: colSum, align: 'center' });
        });

        y = hdr1Y + headerH;

        const attColors = { 1: '#e8f5e9', 2: '#fff9c4', 3: '#fff9c4', 4: '#ffebee' };
        const attLabels = { 1: 'H', 2: 'S', 3: 'I', 4: 'A' };

        for (const [i, r] of rows.entries()) {
          checkPage(rowH);
          const ry = y;
          const rowBg = i % 2 === 0 ? WHITE : LIGHT_GRAY;

          doc.rect(startX, ry, colNo, rowH).fill(rowBg);
          doc.rect(startX, ry, colNo, rowH).lineWidth(0.3).strokeColor('#d0d0d0').stroke();
          doc.font('Helvetica').fontSize(6).fillColor(DARK)
            .text(String(i + 1), startX, ry + 4, { width: colNo, align: 'center' });

          doc.rect(startX + colNo, ry, colNama, rowH).fill(rowBg);
          doc.rect(startX + colNo, ry, colNama, rowH).lineWidth(0.3).strokeColor('#d0d0d0').stroke();
          doc.font('Helvetica').fontSize(6).fillColor(DARK)
            .text(r.guru.nama_guru, startX + colNo + 3, ry + 4, { width: colNama - 6, align: 'left' });

          r.harian.forEach((h, idx) => {
            const cx = htX + idx * colDay;
            const cellBg = h.lewat ? rowBg : (attColors[h.id_kehadiran] || rowBg);
            doc.rect(cx, ry, colDay, rowH).fill(cellBg);
            doc.rect(cx, ry, colDay, rowH).lineWidth(0.3).strokeColor('#d0d0d0').stroke();
            if (!h.lewat && h.id_kehadiran) {
              doc.font('Helvetica').fontSize(5.5).fillColor(DARK)
                .text(attLabels[h.id_kehadiran], cx, ry + 4, { width: colDay, align: 'center' });
            }
          });

          [r.hadir || 0, r.sakit || 0, r.izin || 0, r.alfa || 0].forEach((val, idx) => {
            const cx = rekapX + idx * colSum;
            doc.rect(cx, ry, colSum, rowH).fill(rowBg);
            doc.rect(cx, ry, colSum, rowH).lineWidth(0.3).strokeColor('#d0d0d0').stroke();
            doc.font('Helvetica').fontSize(6).fillColor(DARK)
              .text(String(val), cx, ry + 4, { width: colSum, align: 'center' });
          });

          y += rowH;
        }

        y += 12;
        checkPage(80);
        doc.font('Helvetica-Bold').fontSize(9).fillColor(BLUE)
          .text(`Jumlah Guru : ${rekap.total}`, margin, y);
        doc.font('Helvetica').fontSize(9).fillColor(DARK)
          .text(`Laki-laki : ${rekap.laki}     Perempuan : ${rekap.perempuan}`, margin, y + 14);

        y += 45;
        checkPage(100);
        const sigW = 220;
        const leftSigX = margin + 20;
        const rightSigX = pageW - margin - sigW - 20;

        const bulanParts = bulanLabel.split(' ');
        const namaBulan = bulanParts[0] || '';
        const tahunBulan = bulanParts[1] || '';
        const tanggalCetak = new Date().getDate();

        doc.font('Helvetica').fontSize(9).fillColor(DARK);
        doc.text('Mengetahui,', leftSigX, y, { width: sigW, align: 'center' });
        doc.text('Kepala Sekolah', leftSigX, y + 13, { width: sigW, align: 'center' });
        doc.text('', leftSigX, y + 26, { width: sigW, align: 'center' });
        doc.text('', leftSigX, y + 39, { width: sigW, align: 'center' });
        doc.text('', leftSigX, y + 52, { width: sigW, align: 'center' });
        doc.font('Helvetica').fontSize(9).fillColor(DARK)
          .text('_________________________', leftSigX, y + 65, { width: sigW, align: 'center' });
        doc.fontSize(8).text('NIP.', leftSigX, y + 78, { width: sigW, align: 'center' });

        doc.font('Helvetica').fontSize(9).fillColor(DARK);
        doc.text(`Tegal, ${tanggalCetak} ${namaBulan} ${tahunBulan}`, rightSigX, y, { width: sigW, align: 'center' });
        doc.text('Guru', rightSigX, y + 13, { width: sigW, align: 'center' });
        doc.text('', rightSigX, y + 26, { width: sigW, align: 'center' });
        doc.text('', rightSigX, y + 39, { width: sigW, align: 'center' });
        doc.text('', rightSigX, y + 52, { width: sigW, align: 'center' });
        doc.font('Helvetica').fontSize(9).fillColor(DARK)
          .text('_________________________', rightSigX, y + 65, { width: sigW, align: 'center' });
        doc.fontSize(8).text('NIP.', rightSigX, y + 78, { width: sigW, align: 'center' });

        drawFooter();
        doc.end();
      } catch (err) {
        reject(err);
      }
    });

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', pdfBuffer.length);
    res.end(pdfBuffer);
  } catch (err) {
    console.error('exportLaporanGuruPdf error:', err);
    if (!res.headersSent) {
      res.status(500).json({ success: false, message: `Gagal mengekspor PDF: ${err.message}` });
    }
  }
}

// ============================================================
// WORD/DOCX HELPERS
// ============================================================

const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDERS = { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER };
const THIN_BORDER = { style: BorderStyle.SINGLE, size: 1, color: 'AAAAAA' };
const CELL_BORDERS = { top: THIN_BORDER, bottom: THIN_BORDER, left: THIN_BORDER, right: THIN_BORDER };

function wordCell(text, opts = {}) {
  const { bold, size, color, font, align, fill, colSpan, borders } = opts;
  const cellOpts = {
    borders: borders ?? CELL_BORDERS,
    verticalAlign: VerticalAlign.CENTER,
    children: [
      new Paragraph({
        alignment: align ?? AlignmentType.CENTER,
        spacing: { before: 20, after: 20 },
        children: [
          new TextRun({
            text: text ?? '',
            bold: bold ?? false,
            size: size ?? 18,
            color: color ?? '333333',
            font: font ?? 'Times New Roman',
          }),
        ],
      }),
    ],
  };
  if (fill) cellOpts.shading = { type: 'clear', fill };
  if (colSpan) cellOpts.columnSpan = colSpan;
  return new TableCell(cellOpts);
}

function wordCellMultiParagraph(paragraphs, opts = {}) {
  const { fill, colSpan, borders } = opts;
  const cellOpts = {
    borders: borders ?? CELL_BORDERS,
    verticalAlign: VerticalAlign.CENTER,
    children: paragraphs,
  };
  if (fill) cellOpts.shading = { type: 'clear', fill };
  if (colSpan) cellOpts.columnSpan = colSpan;
  return new TableCell(cellOpts);
}

function buildWordDoc({ logoBuffer, schoolName, schoolYear, title, bulanLabel, kelasName, tanggalList, rows, rekap, personKey, nameKey }) {
  const children = [];

  const colNo = 500, colNama = 2400, colDay = 380, colSum = 400;
  const colWidths = [colNo, colNama, ...tanggalList.map(() => colDay), colSum, colSum, colSum, colSum];
  const totalW = colWidths.reduce((a, b) => a + b, 0);

  const headerChildren = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 0 },
      children: [new TextRun({ text: schoolName.toUpperCase(), bold: true, size: 36, color: '1a3a6b', font: 'Times New Roman' })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 0 },
      children: [new TextRun({ text: title, bold: true, size: 28, color: '333333', font: 'Times New Roman' })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 0 },
      children: [new TextRun({ text: `Tahun Pelajaran ${schoolYear}`, size: 24, font: 'Times New Roman' })],
    }),
  ];

  if (logoBuffer) {
    children.push(new Table({
      columnWidths: [2000, totalW - 2000],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 2000, type: WidthType.DXA },
              borders: NO_BORDERS,
              verticalAlign: VerticalAlign.CENTER,
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [new ImageRun({ data: logoBuffer, transformation: { width: 80, height: 80 }, type: 'jpg' })],
                }),
              ],
            }),
            new TableCell({
              width: { size: totalW - 2000, type: WidthType.DXA },
              borders: NO_BORDERS,
              verticalAlign: VerticalAlign.CENTER,
              children: headerChildren,
            }),
          ],
        }),
      ],
    }));
  } else {
    children.push(...headerChildren);
  }

  children.push(
    new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: '1a3a6b' } }, spacing: { after: 0 }, children: [] }),
    new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '1a3a6b' } }, spacing: { after: 100 }, children: [] }),
  );

  children.push(new Paragraph({ spacing: { before: 100, after: 0 }, children: [new TextRun({ text: `Bulan : ${bulanLabel}`, bold: true, size: 22, font: 'Times New Roman' })] }));
  if (kelasName) children.push(new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: `Kelas : ${kelasName}`, bold: true, size: 22, font: 'Times New Roman' })] }));
  children.push(new Paragraph({ spacing: { after: 200 }, children: [] }));

  const BLUE = '1a3a6b', LIGHT_BLUE = 'e8eef7';
  const attFill = { 1: 'e8f5e9', 2: 'fff9c4', 3: 'fff9c4', 4: 'ffebee' };
  const attLabel = { 1: 'H', 2: 'S', 3: 'I', 4: 'A' };
  const sumColors = ['2e7d32', 'f9a825', 'f9a825', 'c62828'];

  const hdr1 = new TableRow({
    children: [
      wordCell('No / Nama', { bold: true, color: 'FFFFFF', fill: BLUE, colSpan: 2 }),
      wordCell('Hari / Tanggal', { bold: true, color: 'FFFFFF', fill: BLUE, colSpan: tanggalList.length }),
      wordCell('Rekap', { bold: true, color: 'FFFFFF', fill: BLUE, colSpan: 4 }),
    ],
  });

  const hdr2Cells = [
    wordCell('', { fill: BLUE, colSpan: 2 }),
    ...tanggalList.map((t) =>
      wordCellMultiParagraph([
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 10, after: 0 }, children: [new TextRun({ text: t.hariSingkat, bold: true, size: 14, color: BLUE, font: 'Times New Roman' })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 10 }, children: [new TextRun({ text: String(t.tgl), bold: true, size: 16, color: BLUE, font: 'Times New Roman' })] }),
      ], { fill: LIGHT_BLUE })
    ),
    ...Array(4).fill(wordCell('', { fill: BLUE })),
  ];
  const hdr2 = new TableRow({ children: hdr2Cells });

  const hdr3Cells = [
    wordCell('No', { bold: true, fill: 'D9E1F2' }),
    wordCell('Nama', { bold: true, fill: 'D9E1F2' }),
    ...tanggalList.map((t) => wordCell(String(t.tgl), { bold: true, size: 16, fill: 'D9E1F2' })),
    ...['H', 'S', 'I', 'A'].map((l, i) => wordCell(l, { bold: true, color: l === 'H' ? 'FFFFFF' : '333333', fill: sumColors[i] })),
  ];
  const hdr3 = new TableRow({ children: hdr3Cells });

  const dataRows = rows.map((r, i) => {
    const rowBg = i % 2 === 0 ? 'FFFFFF' : 'F8F8F8';
    const person = r.siswa || r.guru;
    const cells = [
      wordCell(String(i + 1), { size: 16, fill: rowBg }),
      wordCell(person[nameKey], { size: 16, fill: rowBg, align: AlignmentType.LEFT }),
      ...r.harian.map((h) => {
        const bg = h.lewat ? rowBg : (attFill[h.id_kehadiran] || rowBg);
        const lbl = (!h.lewat && h.id_kehadiran) ? attLabel[h.id_kehadiran] : '';
        return wordCell(lbl, { size: 14, fill: bg });
      }),
      ...[r.hadir || 0, r.sakit || 0, r.izin || 0, r.alfa || 0].map((v) =>
        wordCell(String(v), { size: 16, fill: rowBg })
      ),
    ];
    return new TableRow({ children: cells });
  });

  children.push(new Table({
    columnWidths: colWidths,
    rows: [hdr1, hdr2, hdr3, ...dataRows],
  }));

  children.push(
    new Paragraph({ spacing: { before: 200 }, children: [] }),
    new Paragraph({ children: [new TextRun({ text: `Jumlah ${personKey === 'siswa' ? 'Siswa' : 'Guru'} : ${rekap.total}`, bold: true, size: 22, color: '1a3a6b', font: 'Times New Roman' })] }),
    new Paragraph({ children: [new TextRun({ text: `Laki-laki : ${rekap.laki}     Perempuan : ${rekap.perempuan}`, size: 22, font: 'Times New Roman' })] }),
  );

  const sigW = Math.floor(totalW / 2);
  children.push(
    new Paragraph({ spacing: { before: 400 }, children: [] }),
    new Table({
      columnWidths: [sigW, sigW],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: NO_BORDERS,
              children: [
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [new TextRun({ text: 'Mengetahui,', bold: true, size: 22, font: 'Times New Roman' })] }),
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [new TextRun({ text: 'Kepala Sekolah', bold: true, size: 22, font: 'Times New Roman' })] }),
                new Paragraph({ spacing: { before: 700 }, children: [] }),
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [new TextRun({ text: '_________________________', size: 22, font: 'Times New Roman' })] }),
                new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'NIP.', size: 22, font: 'Times New Roman' })] }),
              ],
            }),
            new TableCell({
              borders: NO_BORDERS,
              children: [
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [new TextRun({ text: `Tegal, ${bulanLabel}`, bold: true, size: 22, font: 'Times New Roman' })] }),
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [new TextRun({ text: personKey === 'siswa' ? 'Guru Kelas' : 'Guru', bold: true, size: 22, font: 'Times New Roman' })] }),
                new Paragraph({ spacing: { before: 700 }, children: [] }),
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [new TextRun({ text: '_________________________', size: 22, font: 'Times New Roman' })] }),
                new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'NIP.', size: 22, font: 'Times New Roman' })] }),
              ],
            }),
          ],
        }),
      ],
    }),
  );

  return new Document({
    sections: [{
      properties: {
        page: {
          size: { orientation: PageOrientation.LANDSCAPE },
          margin: { top: 567, right: 567, bottom: 567, left: 567 },
        },
      },
      children,
    }],
  });
}

export async function exportLaporanSiswaWord(req, res) {
  try {
    const { kelas: idKelas, bulan } = req.body;
    let payload;
    await generateLaporanSiswaData({ body: { kelas: idKelas, bulan } }, { json: (p) => { payload = p; }, status: () => ({ json: (p) => { payload = p; } }) });
    if (!payload?.success) return res.status(404).json({ success: false, message: 'Data tidak ditemukan.' });

    const { tanggalList, bulanLabel, kelas, rows, rekap, generalSettings } = payload.data;
    const filename = `laporan_absen_${kelas.kelas.replace(/\s+/g, '_')}_${bulanLabel.replace(/\s+/g, '-')}.docx`;

    const logoPath = getLogoPath(generalSettings.logo);
    let logoBuffer = null;
    if (logoPath) try { logoBuffer = fs.readFileSync(logoPath); } catch (_) {}

    const doc = buildWordDoc({
      logoBuffer, schoolName: generalSettings.school_name, schoolYear: generalSettings.school_year,
      title: 'DAFTAR HADIR SISWA', bulanLabel, kelasName: kelas.kelas,
      tanggalList, rows, rekap, personKey: 'siswa', nameKey: 'nama_siswa',
    });

    const buffer = await Packer.toBuffer(doc);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', buffer.length);
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error('exportLaporanSiswaWord error:', err);
    if (!res.headersSent) res.status(500).json({ success: false, message: 'Gagal mengekspor laporan Word.' });
  }
}

export async function exportLaporanGuruWord(req, res) {
  try {
    const { bulan } = req.body;
    let payload;
    await generateLaporanGuruData({ body: { bulan } }, { json: (p) => { payload = p; }, status: () => ({ json: (p) => { payload = p; } }) });
    if (!payload?.success) return res.status(404).json({ success: false, message: 'Data tidak ditemukan.' });

    const { tanggalList, bulanLabel, rows, rekap, generalSettings } = payload.data;
    const filename = `laporan_absen_guru_${bulanLabel.replace(/\s+/g, '-')}.docx`;

    const logoPath = getLogoPath(generalSettings.logo);
    let logoBuffer = null;
    if (logoPath) try { logoBuffer = fs.readFileSync(logoPath); } catch (_) {}

    const doc = buildWordDoc({
      logoBuffer, schoolName: generalSettings.school_name, schoolYear: generalSettings.school_year,
      title: 'DAFTAR HADIR GURU', bulanLabel, kelasName: null,
      tanggalList, rows, rekap, personKey: 'guru', nameKey: 'nama_guru',
    });

    const buffer = await Packer.toBuffer(doc);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', buffer.length);
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error('exportLaporanGuruWord error:', err);
    if (!res.headersSent) res.status(500).json({ success: false, message: 'Gagal mengekspor laporan Word.' });
  }
}
