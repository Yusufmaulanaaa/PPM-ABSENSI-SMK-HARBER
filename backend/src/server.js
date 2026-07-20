import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import path from 'path';
import fs from 'fs';

import authRoutes from './routes/authRoutes.js';
import adminRoutes from './routes/adminRoutes.js';
import teacherRoutes from './routes/teacherRoutes.js';
import scanRoutes from './routes/scanRoutes.js';

dotenv.config();

const app = express();
const PORT = process.env.PORT || 4000;

// ── CORS — allow all origins (auth handled by JWT) ──
app.use(cors({
  origin: true,
  credentials: true,
}));

app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));

// ── Static files (foto siswa, logo) ──
const uploadsRoot = path.join(process.cwd(), 'uploads');
if (!fs.existsSync(uploadsRoot)) fs.mkdirSync(uploadsRoot, { recursive: true });
app.use('/uploads', express.static(uploadsRoot));

// ── Health check ──
app.get('/api/health', (req, res) => {
  res.json({ success: true, message: 'Absensi Sekolah API berjalan dengan baik.', timestamp: new Date().toISOString() });
});

// ── Routes ──
app.use('/api/auth', authRoutes);
app.use('/api/admin', adminRoutes);
app.use('/api/teacher', teacherRoutes);
app.use('/api/scan', scanRoutes);

// ── 404 handler ──
app.use((req, res) => {
  res.status(404).json({ success: false, message: 'Endpoint tidak ditemukan.' });
});

// ── Global error handler ──
app.use((err, req, res, next) => {
  console.error('Unhandled error:', err);
  res.status(500).json({ success: false, message: 'Terjadi kesalahan internal pada server.' });
});

app.listen(PORT, () => {
  console.log(`✓ Server backend Absensi Sekolah berjalan di http://localhost:${PORT}`);
});