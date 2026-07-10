import { useState, useEffect, useCallback } from 'react';
import toast from 'react-hot-toast';
import api, { UPLOADS_BASE_URL } from '../../api/client';
import Card from '../../components/Card';
import Button from '../../components/Button';
import Modal from '../../components/Modal';

const emptyForm = { nis: '', nama: '', id_kelas: '', jk: 'Laki-laki', no_hp: '', rfid: '', foto: null };

export default function DataSiswa() {
  const [siswaList, setSiswaList] = useState([]);
  const [kelasList, setKelasList] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState('');
  const [filterKelas, setFilterKelas] = useState('');

  const [modalOpen, setModalOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState(emptyForm);
  const [errors, setErrors] = useState({});
  const [saving, setSaving] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState(null);

  const loadData = useCallback(() => {
    setLoading(true);
    const params = {};
    Promise.all([
      api.get('/admin/siswa', { params }),
      api.get('/admin/kelas'),
    ])
      .then(([siswaRes, kelasRes]) => {
        setSiswaList(siswaRes.data.data);
        setKelasList(kelasRes.data.data);
      })
      .catch(() => toast.error('Gagal memuat data siswa.'))
      .finally(() => setLoading(false));
  }, []);

  useEffect(() => { loadData(); }, [loadData]);

  function openCreateModal() {
    setEditingId(null);
    setForm(emptyForm);
    setErrors({});
    setModalOpen(true);
  }

  function openEditModal(siswa) {
    setEditingId(siswa.id_siswa);
    setForm({
      nis: siswa.nis, nama: siswa.nama_siswa, id_kelas: siswa.id_kelas,
      jk: siswa.jenis_kelamin, no_hp: siswa.no_hp, rfid: siswa.rfid_code || '', foto: null,
    });
    setErrors({});
    setModalOpen(true);
  }

  async function handleSubmit(e) {
    e.preventDefault();
    setSaving(true);
    setErrors({});

    const fd = new FormData();
    fd.append('nis', form.nis);
    fd.append('nama', form.nama);
    fd.append('id_kelas', form.id_kelas);
    fd.append('jk', form.jk);
    fd.append('no_hp', form.no_hp);
    fd.append('rfid', form.rfid);
    if (form.foto) fd.append('foto', form.foto);

    try {
      if (editingId) {
        await api.put(`/admin/siswa/${editingId}`, fd, { headers: { 'Content-Type': 'multipart/form-data' } });
        toast.success('Data siswa berhasil diubah.');
      } else {
        await api.post('/admin/siswa', fd, { headers: { 'Content-Type': 'multipart/form-data' } });
        toast.success('Siswa baru berhasil ditambahkan.');
      }
      setModalOpen(false);
      loadData();
    } catch (err) {
      const res = err.response?.data;
      if (res?.errors) setErrors(res.errors);
      else toast.error(res?.message || 'Gagal menyimpan data.');
    } finally {
      setSaving(false);
    }
  }

  async function handleDelete() {
    try {
      await api.delete(`/admin/siswa/${deleteTarget.id_siswa}`);
      toast.success('Data siswa berhasil dihapus.');
      setDeleteTarget(null);
      loadData();
    } catch (err) {
      toast.error(err.response?.data?.message || 'Gagal menghapus data.');
    }
  }

  const filtered = siswaList.filter((s) => {
    const matchSearch = s.nama_siswa.toLowerCase().includes(search.toLowerCase()) || s.nis.includes(search);
    const matchKelas = !filterKelas || String(s.id_kelas) === String(filterKelas);
    return matchSearch && matchKelas;
  });

  return (
    <div>
      <div className="page-header">
        <h2>Data Siswa</h2>
        <p>Kelola data induk siswa, kelas, dan kode unik untuk QR/RFID.</p>
      </div>

      <Card
        title="Daftar Siswa"
        headerColor="purple"
        actions={<Button variant="primary" onClick={openCreateModal}>+ Tambah Siswa</Button>}
      >
        <div className="table-toolbar">
          <input className="form-control" style={{ maxWidth: 280 }} placeholder="Cari nama atau NIS..." value={search} onChange={(e) => setSearch(e.target.value)} />
          <select className="form-control" style={{ maxWidth: 220 }} value={filterKelas} onChange={(e) => setFilterKelas(e.target.value)}>
            <option value="">Semua Kelas</option>
            {kelasList.map((k) => <option key={k.id_kelas} value={k.id_kelas}>{k.nama_kelas}</option>)}
          </select>
        </div>

        {loading ? (
          <div className="fullpage-loader" style={{ height: 200 }}><div className="spinner" /></div>
        ) : (
          <div className="table-wrapper">
            <table className="ui-table">
              <thead>
                <tr>
                  <th>Foto</th><th>NIS</th><th>Nama</th><th>Kelas</th><th>Wali Kelas</th><th>Jenis Kelamin</th><th>No. HP Ortu</th><th>RFID</th><th className="text-center">Aksi</th>
                </tr>
              </thead>
              <tbody>
                {filtered.length === 0 && (
                  <tr><td colSpan={9} className="table-empty">Tidak ada data siswa ditemukan.</td></tr>
                )}
                {filtered.map((s) => {
                  const waliKelas = kelasList.find(k => String(k.id_kelas) === String(s.id_kelas))?.nama_wali_kelas;
                  return (
                    <tr key={s.id_siswa}>
                      <td>
                        {s.foto ? (
                          <img className="avatar" src={`${UPLOADS_BASE_URL}/uploads/siswa/${s.foto}`} alt={s.nama_siswa} />
                        ) : (
                          <div className="avatar" />
                        )}
                      </td>
                      <td>{s.nis}</td>
                      <td>{s.nama_siswa}</td>
                      <td>{s.kelas}</td>
                      <td>{waliKelas || <span className="text-muted">—</span>}</td>
                      <td>{s.jenis_kelamin}</td>
                      <td>{s.no_hp}</td>
                      <td>{s.rfid_code || <span className="text-muted">—</span>}</td>
                      <td className="text-center">
                        <Button variant="outline" size="sm" onClick={() => openEditModal(s)}>Edit</Button>{' '}
                        <Button variant="danger" size="sm" onClick={() => setDeleteTarget(s)}>Hapus</Button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </Card>

      <Modal
        open={modalOpen}
        onClose={() => setModalOpen(false)}
        title={editingId ? 'Edit Data Siswa' : 'Tambah Siswa Baru'}
        footer={
          <>
            <Button variant="outline" onClick={() => setModalOpen(false)}>Batal</Button>
            <Button variant="primary" loading={saving} onClick={handleSubmit}>Simpan</Button>
          </>
        }
      >
        <form onSubmit={handleSubmit}>
          <div className="form-row">
            <div className="form-group">
              <label className="form-label">NIS</label>
              <input className="form-control" value={form.nis} onChange={(e) => setForm({ ...form, nis: e.target.value })} />
              {errors.nis && <div className="form-error">{errors.nis}</div>}
            </div>
            <div className="form-group">
              <label className="form-label">Nama Lengkap</label>
              <input className="form-control" value={form.nama} onChange={(e) => setForm({ ...form, nama: e.target.value })} />
              {errors.nama && <div className="form-error">{errors.nama}</div>}
            </div>
          </div>
          <div className="form-row">
            <div className="form-group">
              <label className="form-label">Kelas</label>
              <select className="form-control" value={form.id_kelas} onChange={(e) => setForm({ ...form, id_kelas: e.target.value })}>
                <option value="">Pilih kelas</option>
                {kelasList.map((k) => <option key={k.id_kelas} value={k.id_kelas}>{k.nama_kelas}</option>)}
              </select>
              {errors.id_kelas && <div className="form-error">{errors.id_kelas}</div>}
            </div>
            <div className="form-group">
              <label className="form-label">Jenis Kelamin</label>
              <select className="form-control" value={form.jk} onChange={(e) => setForm({ ...form, jk: e.target.value })}>
                <option value="Laki-laki">Laki-laki</option>
                <option value="Perempuan">Perempuan</option>
              </select>
            </div>
          </div>
          <div className="form-row">
            <div className="form-group">
              <label className="form-label">No. HP Orang Tua (untuk notifikasi WA)</label>
              <input className="form-control" value={form.no_hp} onChange={(e) => setForm({ ...form, no_hp: e.target.value })} placeholder="6281234567890" />
              {errors.no_hp && <div className="form-error">{errors.no_hp}</div>}
            </div>
            <div className="form-group">
              <label className="form-label">Kode RFID (opsional)</label>
              <input className="form-control" value={form.rfid} onChange={(e) => setForm({ ...form, rfid: e.target.value })} placeholder="Tap kartu untuk isi otomatis" />
              {errors.rfid && <div className="form-error">{errors.rfid}</div>}
            </div>
          </div>
          <div className="form-group">
            <label className="form-label">Foto (opsional)</label>
            <input type="file" accept="image/*" className="form-control" onChange={(e) => setForm({ ...form, foto: e.target.files[0] })} />
          </div>
        </form>
      </Modal>

      <Modal
        open={!!deleteTarget}
        onClose={() => setDeleteTarget(null)}
        title="Hapus Data Siswa"
        footer={
          <>
            <Button variant="outline" onClick={() => setDeleteTarget(null)}>Batal</Button>
            <Button variant="danger" onClick={handleDelete}>Ya, Hapus</Button>
          </>
        }
      >
        <p>Yakin ingin menghapus data <strong>{deleteTarget?.nama_siswa}</strong>? Tindakan ini tidak dapat dibatalkan.</p>
      </Modal>
    </div>
  );
}
