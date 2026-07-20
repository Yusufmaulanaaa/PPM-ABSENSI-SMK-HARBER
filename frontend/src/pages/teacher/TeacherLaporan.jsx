import { useState, useEffect } from 'react';
import toast from 'react-hot-toast';
import { FileSpreadsheet, FileDown, FileText } from 'lucide-react';
import api from '../../api/client';
import Card from '../../components/Card';
import Button from '../../components/Button';

function currentMonthStr() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

const BADGE_BG = { 1: '#90ee90', 2: '#ffff00', 3: '#ffff00', 4: '#ff6b6b' };
const BADGE_TEXT = { 1: 'H', 2: 'S', 3: 'I', 4: 'A' };

export default function TeacherLaporan() {
  const [kelas, setKelas] = useState(null);
  const [bulan, setBulan] = useState(currentMonthStr());
  const [report, setReport] = useState(null);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [pageLoading, setPageLoading] = useState(true);

  useEffect(() => {
    api.get('/teacher/laporan')
      .then((res) => setKelas(res.data.data.kelas))
      .finally(() => setPageLoading(false));
  }, []);

  async function generatePreview() {
    setLoading(true);
    setReport(null);
    try {
      const res = await api.post('/admin/laporan/siswa', { kelas: kelas.id_kelas, bulan });
      setReport(res.data.data);
    } catch (err) {
      toast.error(err.response?.data?.message || 'Gagal membuat laporan.');
    } finally {
      setLoading(false);
    }
  }

  async function handleExport(format) {
    setExporting(true);
    try {
      const res = await api.post(`/admin/laporan/siswa/${format}`, { kelas: kelas.id_kelas, bulan }, { responseType: 'blob' });

      const blob = new Blob([res.data]);
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = `laporan-${kelas.nama_kelas}-${bulan}.${format}`;
      a.click();
      URL.revokeObjectURL(a.href);
    } catch (err) {
      if (err.response?.data) {
        try {
          const text = await err.response.data.text();
          const json = JSON.parse(text);
          toast.error(json.message || 'Gagal mengekspor laporan.');
        } catch {
          toast.error('Gagal mengekspor laporan.');
        }
      } else {
        toast.error(err.message || 'Gagal mengekspor laporan.');
      }
    } finally {
      setExporting(false);
    }
  }

  if (pageLoading) return <div className="fullpage-loader"><div className="spinner" /></div>;

  if (!kelas) {
    return (
      <div className="page-header">
        <h2>Laporan Kelas</h2>
        <div className="alert alert-danger" style={{ marginTop: 16 }}>Anda belum ditugaskan sebagai wali kelas.</div>
      </div>
    );
  }

  const rows = report?.rows || [];

  return (
    <div>
      <div className="page-header">
        <h2>Laporan Kelas</h2>
        <p>Rekap kehadiran bulanan untuk kelas {kelas.nama_kelas}.</p>
      </div>

      <Card headerColor="azure" title="Filter Laporan">
        <div className="form-row">
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
          title={`Daftar Hadir Siswa — ${report.bulanLabel}`}
          headerColor="azure"
          actions={<><Button variant="outline" loading={exporting} onClick={() => handleExport('xlsx')}><FileSpreadsheet size={15} style={{ marginRight: 6 }} />Export Excel</Button><Button variant="outline" loading={exporting} onClick={() => handleExport('docx')}><FileText size={15} style={{ marginRight: 6 }} />Export Word</Button><Button variant="info" loading={exporting} onClick={() => handleExport('pdf')}><FileDown size={15} style={{ marginRight: 6 }} />Export PDF</Button></>}
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
                {rows.map((r, i) => (
                  <tr key={r.siswa.id_siswa}>
                    <td>{i + 1}</td>
                    <td>{r.siswa.nama_siswa}</td>
                    {r.harian.map((h, idx) => (
                      <td key={idx} className="text-center" style={{ background: h.lewat ? undefined : BADGE_BG[h.id_kehadiran] }}>{h.lewat ? '' : BADGE_TEXT[h.id_kehadiran]}</td>
                    ))}
                    <td className="text-center">{r.hadir || '-'}</td>
                    <td className="text-center">{r.sakit || '-'}</td>
                    <td className="text-center">{r.izin || '-'}</td>
                    <td className="text-center">{r.alfa || '-'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="laporan-summary">
            <p>Jumlah siswa: <strong>{report.rekap.total}</strong></p>
            <p>Laki-laki: <strong>{report.rekap.laki}</strong></p>
            <p>Perempuan: <strong>{report.rekap.perempuan}</strong></p>
          </div>
        </Card>
      )}
    </div>
  );
}
