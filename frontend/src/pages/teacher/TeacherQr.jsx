import { useState, useEffect } from 'react';
import api, { UPLOADS_BASE_URL } from '../../api/client';
import Card from '../../components/Card';
import Button from '../../components/Button';

export default function TeacherQr() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    api.get('/teacher/qr').then((res) => setData(res.data.data)).finally(() => setLoading(false));
  }, []);

  function downloadUrl(path) {
    const token = localStorage.getItem('absensi_token');
    return `${UPLOADS_BASE_URL}/api${path}?token=${token}`;
  }

  async function handleDownload(siswaId, filename) {
    const res = await api.get(`/admin/qr/siswa/${siswaId}/download`, { responseType: 'blob' });
    const blob = new Blob([res.data]);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  if (loading) return <div className="fullpage-loader"><div className="spinner" /></div>;

  if (!data?.kelas) {
    return (
      <div className="page-header">
        <h2>QR Code Siswa</h2>
        <div className="alert alert-danger" style={{ marginTop: 16 }}>Anda belum ditugaskan sebagai wali kelas.</div>
      </div>
    );
  }

  return (
    <div>
      <div className="page-header">
        <h2>QR Code Siswa</h2>
        <p>QR Code presensi untuk siswa kelas {data.kelas.nama_kelas}.</p>
      </div>

      <Card title={`Daftar QR Siswa — ${data.kelas.nama_kelas}`} headerColor="danger">
        <div className="qr-grid">
          {data.siswa.length === 0 && <div className="table-empty">Belum ada data siswa di kelas ini.</div>}
          {data.siswa.map((s) => (
            <div className="qr-card" key={s.id_siswa}>
              <img src={downloadUrl(`/admin/qr/siswa/${s.id_siswa}/view`)} alt={s.nama_siswa} />
              <p className="qr-card-name">{s.nama_siswa}</p>
              <p className="qr-card-meta">NIS: {s.nis}</p>
              <Button variant="outline" size="sm" onClick={() => handleDownload(s.id_siswa, `${s.nama_siswa}.svg`)}>Download</Button>
            </div>
          ))}
        </div>
      </Card>
    </div>
  );
}
