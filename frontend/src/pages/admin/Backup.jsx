import { useState, useRef } from 'react';
import toast from 'react-hot-toast';
import { Download } from 'lucide-react';
import api from '../../api/client';
import Card from '../../components/Card';
import Button from '../../components/Button';
import Modal from '../../components/Modal';

export default function Backup() {
  const [downloading, setDownloading] = useState('');
  const [restoring, setRestoring] = useState(false);
  const [confirmRestore, setConfirmRestore] = useState(null);
  const dbFileRef = useRef(null);
  const photosFileRef = useRef(null);

  async function handleDownload(type) {
    setDownloading(type);
    try {
      const endpoint = type === 'db' ? '/admin/backup/db' : '/admin/backup/photos';
      const res = await api.get(endpoint, { responseType: 'blob' });
      const blob = new Blob([res.data]);
      const filename = type === 'db' ? `backup_db_${Date.now()}.json` : `backup_photos_${Date.now()}.zip`;
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = filename;
      a.click();
      URL.revokeObjectURL(a.href);
      toast.success('Backup berhasil diunduh.');
    } catch (err) {
      toast.error(err.response?.data?.message || 'Gagal membuat backup.');
    } finally {
      setDownloading('');
    }
  }

  async function handleRestore(type) {
    const fileInput = type === 'db' ? dbFileRef.current : photosFileRef.current;
    const file = fileInput.files[0];
    if (!file) {
      toast.error('Pilih file backup terlebih dahulu.');
      return;
    }

    setRestoring(true);
    const fd = new FormData();
    fd.append('file', file);

    try {
      const endpoint = type === 'db' ? '/admin/backup/db/restore' : '/admin/backup/photos/restore';
      const res = await api.post(endpoint, fd, { headers: { 'Content-Type': 'multipart/form-data' } });
      toast.success(res.data.message);
      fileInput.value = '';
    } catch (err) {
      toast.error(err.response?.data?.message || 'Gagal melakukan restore.');
    } finally {
      setRestoring(false);
      setConfirmRestore(null);
    }
  }

  return (
    <div>
      <div className="page-header">
        <h2>Backup & Restore</h2>
        <p>Amankan data sistem secara berkala atau pulihkan dari backup sebelumnya.</p>
      </div>

      <Card title="Backup & Restore Database" headerColor="danger">
        <p className="text-muted">Database mencakup data siswa, guru, kelas, jurusan, dan riwayat presensi (format JSON).</p>
        <div className="backup-actions">
          <Button variant="danger" loading={downloading === 'db'} onClick={() => handleDownload('db')}><Download size={15} style={{ marginRight: 6 }} />Download Backup Database</Button>
          <div className="backup-restore-row">
            <input type="file" accept=".json" ref={dbFileRef} className="form-control" />
            <Button variant="outline" onClick={() => setConfirmRestore('db')}>Restore Database</Button>
          </div>
        </div>
      </Card>

      <Card title="Backup & Restore Foto Siswa" headerColor="danger">
        <p className="text-muted">Mencakup seluruh foto profil siswa dan logo sekolah (format ZIP).</p>
        <div className="backup-actions">
          <Button variant="danger" loading={downloading === 'photos'} onClick={() => handleDownload('photos')}><Download size={15} style={{ marginRight: 6 }} />Download Backup Foto</Button>
          <div className="backup-restore-row">
            <input type="file" accept=".zip" ref={photosFileRef} className="form-control" />
            <Button variant="outline" onClick={() => setConfirmRestore('photos')}>Restore Foto</Button>
          </div>
        </div>
      </Card>

      <Modal
        open={!!confirmRestore} onClose={() => setConfirmRestore(null)} title="Konfirmasi Restore"
        footer={<><Button variant="outline" onClick={() => setConfirmRestore(null)}>Batal</Button><Button variant="danger" loading={restoring} onClick={() => handleRestore(confirmRestore)}>Ya, Restore Sekarang</Button></>}
      >
        <p>
          Tindakan ini akan <strong>mengganti seluruh data {confirmRestore === 'db' ? 'database' : 'foto'} saat ini</strong> dengan isi file backup. Tindakan ini tidak dapat dibatalkan. Lanjutkan?
        </p>
      </Modal>
    </div>
  );
}
