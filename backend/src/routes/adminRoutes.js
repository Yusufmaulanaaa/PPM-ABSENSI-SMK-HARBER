import express from 'express';
import { authenticate, requirePermission } from '../middleware/auth.js';
import { uploadSiswaFoto, uploadLogo, uploadBackup } from '../middleware/upload.js';

import { getDashboard, filterDashboardData } from '../controllers/dashboardController.js';
import * as siswaCtrl from '../controllers/siswaController.js';
import * as guruCtrl from '../controllers/guruController.js';
import * as kelasCtrl from '../controllers/kelasController.js';
import * as petugasCtrl from '../controllers/petugasController.js';
import * as absensiCtrl from '../controllers/absensiController.js';
import * as qrCtrl from '../controllers/qrController.js';
import * as laporanCtrl from '../controllers/laporanController.js';
import * as settingsCtrl from '../controllers/settingsController.js';
import * as backupCtrl from '../controllers/backupController.js';

const router = express.Router();

router.use(authenticate);

// ── Dashboard ──
router.get('/dashboard', requirePermission('dashboard.view-admin'), getDashboard);
router.post('/dashboard/filter-data', requirePermission('dashboard.view-admin'), filterDashboardData);

// ── Absensi Siswa & Guru ──
router.get('/absen-siswa', requirePermission('attendance.edit'), absensiCtrl.listAbsenSiswa);
router.post('/absen-siswa/edit', requirePermission('attendance.edit'), absensiCtrl.updateKehadiranSiswa);
router.get('/absen-siswa/detail', requirePermission('attendance.edit'), absensiCtrl.getDetailKehadiranSiswa);

router.get('/absen-guru', requirePermission('attendance.edit'), absensiCtrl.listAbsenGuru);
router.post('/absen-guru/edit', requirePermission('attendance.edit'), absensiCtrl.updateKehadiranGuru);
router.get('/absen-guru/detail', requirePermission('attendance.edit'), absensiCtrl.getDetailKehadiranGuru);

// ── Data Siswa ──
router.get('/siswa', requirePermission('students.manage'), siswaCtrl.listSiswa);
router.get('/siswa/:id', requirePermission('students.manage'), siswaCtrl.getSiswaById);
router.post('/siswa', requirePermission('students.manage'), uploadSiswaFoto.single('foto'), siswaCtrl.createSiswa);
router.put('/siswa/:id', requirePermission('students.manage'), uploadSiswaFoto.single('foto'), siswaCtrl.updateSiswa);
router.delete('/siswa/:id', requirePermission('students.manage'), siswaCtrl.deleteSiswa);
router.post('/siswa/delete-selected', requirePermission('students.manage'), siswaCtrl.deleteSelectedSiswa);
router.post('/siswa/import', requirePermission('students.manage'), siswaCtrl.importSiswaCSV);

// ── Data Guru ──
router.get('/guru', requirePermission('teachers.manage'), guruCtrl.listGuru);
router.get('/guru/:id', requirePermission('teachers.manage'), guruCtrl.getGuruById);
router.post('/guru', requirePermission('teachers.manage'), guruCtrl.createGuru);
router.put('/guru/:id', requirePermission('teachers.manage'), guruCtrl.updateGuru);
router.delete('/guru/:id', requirePermission('teachers.manage'), guruCtrl.deleteGuru);
router.post('/guru/import', requirePermission('teachers.manage'), guruCtrl.importGuruCSV);

// ── Kelas ──
router.get('/kelas', requirePermission('classes.manage'), kelasCtrl.listKelas);
router.get('/kelas/:id', requirePermission('classes.manage'), kelasCtrl.getKelasById);
router.post('/kelas', requirePermission('classes.manage'), kelasCtrl.createKelas);
router.put('/kelas/:id', requirePermission('classes.manage'), kelasCtrl.updateKelas);
router.delete('/kelas/:id', requirePermission('classes.manage'), kelasCtrl.deleteKelas);
router.post('/kelas/import', requirePermission('classes.manage'), kelasCtrl.importKelasCSV);

// ── Jurusan ──
router.get('/jurusan', requirePermission('classes.manage'), kelasCtrl.listJurusan);
router.get('/jurusan/:id', requirePermission('classes.manage'), kelasCtrl.getJurusanById);
router.post('/jurusan', requirePermission('classes.manage'), kelasCtrl.createJurusan);
router.put('/jurusan/:id', requirePermission('classes.manage'), kelasCtrl.updateJurusan);
router.delete('/jurusan/:id', requirePermission('classes.manage'), kelasCtrl.deleteJurusan);
router.post('/jurusan/import', requirePermission('classes.manage'), kelasCtrl.importJurusanCSV);

// ── Generate QR ──
router.post('/generate/siswa-by-kelas', requirePermission('qr.generate'), qrCtrl.getSiswaByKelas);
router.get('/qr/siswa/:id/view', requirePermission('qr.generate'), qrCtrl.viewQrSiswa);
router.get('/qr/siswa/:id/download', requirePermission('qr.generate'), qrCtrl.downloadQrSiswa);
router.get('/qr/siswa/download', requirePermission('qr.generate'), qrCtrl.downloadAllQrSiswa);
router.get('/qr/guru/:id/view', requirePermission('qr.generate'), qrCtrl.viewQrGuru);
router.get('/qr/guru/:id/download', requirePermission('qr.generate'), qrCtrl.downloadQrGuru);
router.get('/qr/guru/download', requirePermission('qr.generate'), qrCtrl.downloadAllQrGuru);

// ── Laporan ──
router.get('/laporan/test-pdf', laporanCtrl.testPdfEndpoint);
router.post('/laporan/siswa', requirePermission('attendance.view'), laporanCtrl.generateLaporanSiswaData);
router.post('/laporan/siswa/xlsx', requirePermission('attendance.view'), laporanCtrl.exportLaporanSiswaExcel);
router.post('/laporan/siswa/pdf', requirePermission('attendance.view'), laporanCtrl.exportLaporanSiswaPdf);
router.post('/laporan/siswa/docx', requirePermission('attendance.view'), laporanCtrl.exportLaporanSiswaWord);
router.post('/laporan/guru', requirePermission('attendance.view'), laporanCtrl.generateLaporanGuruData);
router.post('/laporan/guru/xlsx', requirePermission('attendance.view'), laporanCtrl.exportLaporanGuruExcel);
router.post('/laporan/guru/pdf', requirePermission('attendance.view'), laporanCtrl.exportLaporanGuruPdf);
router.post('/laporan/guru/docx', requirePermission('attendance.view'), laporanCtrl.exportLaporanGuruWord);

// ── Petugas ──
router.get('/petugas', requirePermission('petugas.manage'), petugasCtrl.listPetugas);
router.get('/petugas/:id', requirePermission('petugas.manage'), petugasCtrl.getPetugasById);
router.post('/petugas/register', requirePermission('petugas.manage'), petugasCtrl.registerPetugas);
router.put('/petugas/:id', requirePermission('petugas.manage'), petugasCtrl.updatePetugas);
router.delete('/petugas/:id', requirePermission('petugas.manage'), petugasCtrl.deletePetugas);
router.post('/petugas/:id/toggle-activation', requirePermission('petugas.manage'), petugasCtrl.toggleActivation);

// ── Pengaturan ──
router.get('/general-settings', requirePermission('settings.manage'), settingsCtrl.getGeneralSettings);
router.post('/general-settings', requirePermission('settings.manage'), uploadLogo.single('logo'), settingsCtrl.updateGeneralSettings);

// ── Backup & Restore ──
router.get('/backup/db', requirePermission('backup.manage'), backupCtrl.dbBackup);
router.post('/backup/db/restore', requirePermission('backup.manage'), uploadBackup.single('file'), backupCtrl.dbRestore);
router.get('/backup/photos', requirePermission('backup.manage'), backupCtrl.photosBackup);
router.post('/backup/photos/restore', requirePermission('backup.manage'), uploadBackup.single('file'), backupCtrl.photosRestore);

export default router;
