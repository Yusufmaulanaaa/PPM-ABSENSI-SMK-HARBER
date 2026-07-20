import { useState, useEffect } from 'react';
import toast from 'react-hot-toast';
import { FileSpreadsheet, FileDown, FileText } from 'lucide-react';
import api, { UPLOADS_BASE_URL } from '../../api/client';
import Card from '../../components/Card';
import Button from '../../components/Button';

function currentMonthStr() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

const BADGE_BG = { 1: '#90ee90', 2: '#ffff00', 3: '#ffff00', 4: '#ff6b6b' };
const BADGE_TEXT = { 1: 'H', 2: 'S', 3: 'I', 4: 'A' };

export default function GenerateLaporan() {
  const [tab, setTab] = useState('siswa');
  const [kelasList, setKelasList] = useState([]);
  const [idKelas, setIdKelas] = useState('');
  const [bulan, setBulan] = useState(currentMonthStr());
  const [report, setReport] = useState(null);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);

  useEffect(() => {
    api.get('/admin/kelas').then((res) => {
      setKelasList(res.data.data);
      if (res.data.data.length > 0) setIdKelas(res.data.data[0].id_kelas);
    });
  }, []);

  async function generatePreview() {
    setLoading(true);
    setReport(null);
    try {
      if (tab === 'siswa') {
        const res = await api.post('/admin/laporan/siswa', { kelas: idKelas, bulan });
        setReport(res.data.data);
      } else {
        const res = await api.post('/admin/laporan/guru', { bulan });
        setReport(res.data.data);
      }
    } catch (err) {
      toast.error(err.response?.data?.message || 'Gagal membuat laporan.');
    } finally {
      setLoading(false);
    }
  }

  async function handleExport(format) {
    setExporting(true);
    try {
      const token = localStorage.getItem('absensi_token');
      const endpoint = tab === 'siswa' ? `/admin/laporan/siswa/${format}` : `/admin/laporan/guru/${format}`;
      const body = tab === 'siswa' ? { kelas: idKelas, bulan } : { bulan };

      const res = await fetch(`${UPLOADS_BASE_URL}/api${endpoint}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify(body),
      });

      if (!res.ok) throw new Error('Gagal mengekspor laporan.');

      const blob = await res.blob();
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = `laporan-${tab}-${bulan}.${format === 'xlsx' ? 'xlsx' : format === 'docx' ? 'docx' : 'pdf'}`;
      a.click();
      URL.revokeObjectURL(a.href);
    } catch (err) {
      toast.error(err.message || 'Gagal mengekspor laporan.');
    } finally {
      setExporting(false);
    }
  }

  const rows = report?.rows || [];

  return (
    <div>
      <div className="page-header">
        <h2>Generate Laporan</h2>
        <p>Buat rekap kehadiran bulanan siswa atau guru, lalu unduh sebagai PDF/Word/Excel.</p>
      </div>

      <div className="tab-bar">
        <button className={`tab-btn ${tab === 'siswa' ? 'active' : ''}`} onClick={() => { setTab('siswa'); setReport(null); }}>Laporan Siswa</button>
        <button className={`tab-btn ${tab === 'guru' ? 'active' : ''}`} onClick={() => { setTab('guru'); setReport(null); }}>Laporan Guru</button>
      </div>

      <Card headerColor="azure" title="Filter Laporan">
        <div className="form-row">
          {tab === 'siswa' && (
            <div className="form-group">
              <label className="form-label">Kelas</label>
              <select className="form-control" value={idKelas} onChange={(e) => setIdKelas(e.target.value)}>
                {kelasList.map((k) => <option key={k.id_kelas} value={k.id_kelas}>{k.nama_kelas}</option>)}
              </select>
            </div>
          )}
          <div className="form-group">
            <label className="form-label">Bulan</label>
            <input type="month" className="form-control" value={bulan} onChange={(e) => setBulan(e.target.value)} />
          </div>
          <div className="form-group" style={{ display: 'flex', alignItems: 'flex-end' }}>
            <Button variant="info" loading={loading} onClick={generatePreview}>Tampilkan Laporan</Button>
          </div>
        </div>
      </Card>

      {report && (
        <Card
          title={`Daftar Hadir ${tab === 'siswa' ? 'Siswa' : 'Guru'} — ${report.bulanLabel}`}
          headerColor="azure"
          actions={
            <>
              <Button variant="outline" loading={exporting} onClick={() => handleExport('xlsx')}><FileSpreadsheet size={15} style={{ marginRight: 6 }} />Export Excel</Button>
              <Button variant="outline" loading={exporting} onClick={() => handleExport('docx')}><FileText size={15} style={{ marginRight: 6 }} />Export Word</Button>
              <Button variant="info" loading={exporting} onClick={() => handleExport('pdf')}><FileDown size={15} style={{ marginRight: 6 }} />Export PDF</Button>
            </>
          }
        >
          <div className="table-wrapper">
            <table className="ui-table laporan-table">
              <thead>
                <tr>
                  <th>No</th><th>Nama</th>
                  {report.tanggalList.map((t) => <th key={t.tanggal} className="text-center">{t.tgl}</th>)}
                  <th className="text-center" style={{ background: '#90ee90' }}>H</th>
                  <th className="text-center" style={{ background: '#ffff00' }}>S</th>
                  <th className="text-center" style={{ background: '#ffff00' }}>I</th>
                  <th className="text-center" style={{ background: '#ff6b6b' }}>A</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => {
                  const person = r.siswa || r.guru;
                  return (
                    <tr key={person.id_siswa || person.id_guru}>
                      <td>{i + 1}</td>
                      <td>{person.nama_siswa || person.nama_guru}</td>
                      {r.harian.map((h, idx) => (
                        <td key={idx} className="text-center" style={{ background: h.lewat ? undefined : BADGE_BG[h.id_kehadiran] }}>
                          {h.lewat ? '' : BADGE_TEXT[h.id_kehadiran]}
                        </td>
                      ))}
                      <td className="text-center">{r.hadir || '-'}</td>
                      <td className="text-center">{r.sakit || '-'}</td>
                      <td className="text-center">{r.izin || '-'}</td>
                      <td className="text-center">{r.alfa || '-'}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <div className="laporan-summary">
            <p>Jumlah {tab === 'siswa' ? 'siswa' : 'guru'}: <strong>{report.rekap.total}</strong></p>
            <p>Laki-laki: <strong>{report.rekap.laki}</strong></p>
            <p>Perempuan: <strong>{report.rekap.perempuan}</strong></p>
          </div>
        </Card>
      )}
    </div>
  );
}
