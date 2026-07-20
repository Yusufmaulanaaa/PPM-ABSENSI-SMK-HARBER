import { useState, useEffect } from 'react';
import { Download } from 'lucide-react';
import api, { UPLOADS_BASE_URL } from '../../api/client';
import Card from '../../components/Card';
import Button from '../../components/Button';

export default function GenerateQR() {
  const [tab, setTab] = useState('siswa');
  const [kelasList, setKelasList] = useState([]);
  const [idKelas, setIdKelas] = useState('');
  const [siswaList, setSiswaList] = useState([]);
  const [guruList, setGuruList] = useState([]);
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    api.get('/admin/kelas').then((res) => setKelasList(res.data.data));
    api.get('/admin/guru').then((res) => setGuruList(res.data.data));
  }, []);

  useEffect(() => {
    if (tab !== 'siswa') return;
    setLoading(true);
    api.post('/admin/generate/siswa-by-kelas', { idKelas: idKelas || undefined })
      .then((res) => setSiswaList(res.data))
      .catch(() => {})
      .finally(() => setLoading(false));
  }, [tab, idKelas]);

  function downloadUrl(path) {
    const token = localStorage.getItem('absensi_token');
    return `${UPLOADS_BASE_URL}/api${path}${path.includes('?') ? '&' : '?'}token=${token}`;
  }

  async function handleDownload(url, filename) {
    const res = await api.get(url, { responseType: 'blob' });
    const blob = new Blob([res.data]);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  return (
    <div>
      <div className="page-header">
        <h2>Generate QR Code</h2>
        <p>Buat dan unduh QR Code untuk presensi siswa dan guru.</p>
      </div>

      <div className="tab-bar">
        <button className={`tab-btn ${tab === 'siswa' ? 'active' : ''}`} onClick={() => setTab('siswa')}>QR Siswa</button>
        <button className={`tab-btn ${tab === 'guru' ? 'active' : ''}`} onClick={() => setTab('guru')}>QR Guru</button>
      </div>

      {tab === 'siswa' && (
        <Card
          headerColor="danger"
          title="QR Code Siswa"
          actions={
            <>
              <select className="form-control" style={{ width: 200 }} value={idKelas} onChange={(e) => setIdKelas(e.target.value)}>
                <option value="">Semua Kelas</option>
                {kelasList.map((k) => <option key={k.id_kelas} value={k.id_kelas}>{k.nama_kelas}</option>)}
              </select>
              <Button variant="danger" onClick={() => handleDownload(`/admin/qr/siswa/download${idKelas ? `?id_kelas=${idKelas}` : ''}`, 'qr-siswa.zip')}>
                <Download size={15} style={{ marginRight: 6 }} />Download Semua (ZIP)
              </Button>
            </>
          }
        >
          {loading ? (
            <div className="fullpage-loader" style={{ height: 160 }}><div className="spinner" /></div>
          ) : (
            <div className="qr-grid">
              {siswaList.length === 0 && <div className="table-empty">Tidak ada data siswa.</div>}
              {siswaList.map((s) => (
                <div className="qr-card" key={s.id_siswa}>
                  <img src={downloadUrl(`/admin/qr/siswa/${s.id_siswa}/view`)} alt={s.nama_siswa} />
                  <p className="qr-card-name">{s.nama_siswa}</p>
                  <p className="qr-card-meta">NIS: {s.nis}</p>
                  <Button variant="outline" size="sm" onClick={() => handleDownload(`/admin/qr/siswa/${s.id_siswa}/download`, `${s.nama_siswa}.svg`)}>Download</Button>
                </div>
              ))}
            </div>
          )}
        </Card>
      )}

      {tab === 'guru' && (
        <Card
          headerColor="danger"
          title="QR Code Guru"
          actions={<Button variant="danger" onClick={() => handleDownload('/admin/qr/guru/download', 'qr-guru.zip')}><Download size={15} style={{ marginRight: 6 }} />Download Semua (ZIP)</Button>}
        >
          <div className="qr-grid">
            {guruList.length === 0 && <div className="table-empty">Tidak ada data guru.</div>}
            {guruList.map((g) => (
              <div className="qr-card" key={g.id_guru}>
                <img src={downloadUrl(`/admin/qr/guru/${g.id_guru}/view`)} alt={g.nama_guru} />
                <p className="qr-card-name">{g.nama_guru}</p>
                <p className="qr-card-meta">NUPTK: {g.nuptk}</p>
                <Button variant="outline" size="sm" onClick={() => handleDownload(`/admin/qr/guru/${g.id_guru}/download`, `${g.nama_guru}.svg`)}>Download</Button>
              </div>
            ))}
          </div>
        </Card>
      )}
    </div>
  );
}
