/* common.js — data bersama, penyimpanan, dan komponen kecil untuk semua halaman */
(function (root) {
  'use strict';
  var G = root.DocGen;
  var K_SAVED = 'umlift.inv.bawaan.v1', K_UMUM = 'umlift.inv.umum.v1', K_HIST = 'umlift.inv.hist.v1', K_DRAFT = 'umlift.draft.v4', K_SPP = 'umlift.spp.umum.v1', K_TKDN = 'umlift.tkdn.umum.v1';

  function $(id) { return document.getElementById(id); }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }
  function lsGet(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} }
  function ssGet(k) { try { return JSON.parse(sessionStorage.getItem(k)); } catch (e) { return null; } }
  function ssSet(k, v) { try { sessionStorage.setItem(k, JSON.stringify(v)); } catch (e) {} }
  function today() { var d = new Date(), p = function (n) { return ('0' + n).slice(-2); }; return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()); }
  function mk(tag, cls, attrs) { var e = document.createElement(tag); if (cls) e.className = cls; if (attrs) for (var k in attrs) e.setAttribute(k, attrs[k]); return e; }
  function lab(t, input) { var d = mk('div'), l = mk('label'); l.textContent = t; d.appendChild(l); d.appendChild(input); return d; }
  function money(el) { var d = el.value.replace(/\D/g, ''); el.value = d ? G.rp(+d) : ''; return d; }

  var toastT;
  function toast(m, err) {
    var t = $('toast'); if (!t) { t = mk('div'); t.id = 'toast'; t.setAttribute('role', 'status'); document.body.appendChild(t); }
    t.textContent = m; t.className = 'on' + (err ? ' err' : '');
    clearTimeout(toastT); toastT = setTimeout(function () { t.className = ''; }, err ? 4500 : 2500);
  }
  function saveBlob(blob, name) {
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  }

  /* ---------- daftar dokumen ---------- */
  var DOCS = [
    { key: 'invoice', title: 'Invoice', file: 'invoice.html', ready: true, desc: 'Pilih kode 1–4, periksa barang & harga, unduh Word.' },
    { key: 'kuitansi', title: 'Kuitansi', file: 'kuitansi.html', ready: true, desc: 'Otomatis mengikuti invoice yang sama.' },
    { key: 'permohonan', title: 'Surat Permohonan Pembayaran', file: 'permohonan.html', ready: true, desc: 'Nominal & paket otomatis dari invoice.' },
    { key: 'tkdn', title: 'Surat TKDN', file: 'tkdn.html', ready: true, desc: 'Paket & PO otomatis dari invoice.' },
    { key: 'ba', title: 'Berita Acara Pemeriksaan Pekerjaan', file: '#', ready: false, desc: 'Segera hadir.' },
    { key: 'sr', title: 'Service Report', file: '#', ready: false, desc: 'Segera hadir.' }
  ];
  function nav(active) {
    var n = $('nav'); if (!n) return; n.innerHTML = '';
    DOCS.forEach(function (d) {
      var short = { permohonan: 'Permohonan Pembayaran', ba: 'Berita Acara' }[d.key] || d.title;
      var c = d.ready ? mk('a', 'chip' + (d.key === active ? ' on' : ''), { href: d.file }) : mk('span', 'chip off');
      c.textContent = short + (d.ready ? '' : ' · segera'); n.appendChild(c);
    });
    var on = n.querySelector('.on'); if (on && on.scrollIntoView) { try { var x = on.offsetLeft - n.offsetLeft; if (x + on.offsetWidth > n.clientWidth) n.scrollLeft = Math.max(0, x - 16); } catch (e) {} }
  }

  /* ---------- bawaan awal: dari invoice & kuitansi asli ---------- */
  function it(nama, qty, harga) { return { nama: nama, satuan: 'Pcs', qty: String(qty), harga: String(harga) }; }
  var FACTORY = {
    1: { noPO: '185/PPK.10/DIRI-BRIN/PO/6/2026', tglPO: '2026-06-02', jatuhTempo: '2026-09-11', jtManual: true, potongan: '', uraian: 'Pemeliharaan rutin lift Pillar pada gedung GWB KST BJ Habibie',
         paket: 'Pemeliharaan rutin lift Pillar pada gedung GWB KST BJ Habibie', noEP: 'EP-01KZNHS5B663XY6NEDT4SCX6R6', tglEP: '',
         items: [it('Service berkala lift Pillar 1 Agustus', 1, 2500000), it('Service berkala lift Pillar 2 Agustus', 1, 2500000)] },
    2: { noPO: '310/PPK.10/DIRI-BRIN/PO/7/2026', tglPO: '2026-07-14', jatuhTempo: '2026-10-14', jtManual: true, potongan: '', uraian: 'Pemeliharaan rutin lift (Non Otis) di KST Soekarno dan KST Rancabungur',
         paket: 'Pemeliharaan rutin lift (Non Otis) di KST Soekarno dan KST Rancabungur', noEP: 'EP-01KZNHS5B663XY6NEDT4SCX6R6', tglEP: '',
         items: [it('Service berkala lift non otis 2 lantai', 3, 1000000), it('Service berkala lift non otis 2 lantai 3000kg', 1, 1200000), it('Service berkala lift non otis 3 lantai', 2, 1500000)] },
    3: { noPO: '467/PPK.10/DIRI-BRIN/PO/9/2026', tglPO: '2026-09-14', jatuhTempo: '2026-10-14', jtManual: true, potongan: '', uraian: 'Perbaikan lift Hyunday pada Gedung GWB KST NJ Habibie',
         paket: 'Jasa Perawatan dan Service Berkala lift Hyunday', noEP: 'EP-01KZNHS5B663XY6NEDT4SCX6R6', tglEP: '',
         items: [it('Service berkala lift Hyunday September', 1, 2500000)] },
    4: { noPO: '', tglPO: '', jatuhTempo: '', jtManual: false, potongan: '', uraian: '', paket: '', noEP: '', tglEP: '', items: [] }
  };
  var UMUM0 = { kepada: 'DEPUTI BIDANG STRUKTUR RISET DAN INOVASI BRIN',
    alamat1: 'Gd. B.J Habibie Lt.14. Jl. Kp. Bali XXXIV No. 15 8, RT 10/RW 10 Kp Bali',
    alamat2: 'Kecamatan Tanah Abang, Kota Jakarta Pusat', keterangan: 'Service berkala', ppn: '11' };

  function norm(d) {
    d = clone(d); d.noManual = d.noManual || ''; d.tglKw = d.tglKw || ''; d.uraian = d.uraian || ''; d.paket = d.paket || ''; d.noEP = d.noEP || ''; d.tglEP = d.tglEP || ''; d.sppNo = d.sppNo || ''; d.sppTgl = d.sppTgl || ''; d.tkdnTgl = d.tkdnTgl || ''; d.items = d.items || [];
    while (d.items.length < 5) d.items.push({ nama: '', satuan: 'Pcs', qty: '', harga: '' });
    d.items = d.items.slice(0, 5); return d;
  }

  var saved = lsGet(K_SAVED) || {};
  var umum = Object.assign({}, UMUM0, lsGet(K_UMUM) || {});
  var hist = lsGet(K_HIST) || {};
  var spp = Object.assign({ urut: '01' }, G.SPP0, lsGet(K_SPP) || {});
  var tkdnSet = Object.assign({ jabatan: G.TKDN0.jabatan, alamat: G.TKDN0.alamat, pekerjaan: '' }, lsGet(K_TKDN) || {});
  function base(k) { return saved[k] || FACTORY[k]; }
  function fresh(k) { var b = clone(base(k)); b.noManual = ''; b.tglKw = ''; b.sppNo = ''; b.sppTgl = ''; b.tkdnTgl = ''; return norm(b); }

  var draft = ssGet(K_DRAFT) || {};
  var S = { kode: draft.kode || 1, tanggal: draft.tanggal || today(), work: {} };
  [1, 2, 3, 4].forEach(function (k) { S.work[k] = (draft.work && draft.work[k]) ? norm(draft.work[k]) : fresh(k); });
  function persist() { ssSet(K_DRAFT, { kode: S.kode, tanggal: S.tanggal, work: S.work }); }
  function persistUmum() { lsSet(K_UMUM, umum); }
  function persistSpp() { lsSet(K_SPP, spp); }
  function persistTkdn() { lsSet(K_TKDN, tkdnSet); }
  // data untuk Surat TKDN: nama penandatangan dipakai bersama dengan surat permohonan
  function dataTkdn(k) {
    var w = S.work[k], d = dataFor(k);
    d.tanggalSurat = w.tkdnTgl || S.tanggal;
    d.tkdn = { nama: spp.penandatangan, jabatan: tkdnSet.jabatan, alamat: tkdnSet.alamat, pekerjaan: tkdnSet.pekerjaan };
    return d;
  }

  function auto(k) { return G.nomorInvoice(S.tanggal, k); }
  function dataFor(k) {
    var w = S.work[k];
    return { noInvoice: w.noManual || auto(k), tanggal: S.tanggal, kepada: umum.kepada, alamat1: umum.alamat1, alamat2: umum.alamat2,
      keterangan: umum.keterangan, ppn: umum.ppn, noPO: w.noPO, tglPO: w.tglPO, jatuhTempo: w.jatuhTempo, items: w.items, potongan: w.potongan,
      uraian: w.uraian, tanggalKw: w.tglKw || S.tanggal,
      paket: w.paket || w.uraian, noEP: w.noEP, tglEP: w.tglEP || w.sppTgl || S.tanggal, tanggalSurat: w.sppTgl || S.tanggal,
      noSurat: w.sppNo || G.nomorSurat(spp.urut, w.sppTgl || S.tanggal), spp: spp };
  }
  function cek(k) { var d = dataFor(k); if (!d.tanggal) throw new Error('Isi tanggal invoice.'); if (!d.noInvoice) throw new Error('Nomor invoice kosong.'); return d; }

  function remember(k) {
    S.work[k].items.forEach(function (r) { var n = (r.nama || '').trim(), h = G.num(r.harga); if (n && h) hist[n] = { harga: h }; });
    lsSet(K_HIST, hist);
  }
  function names() {
    var o = {}; Object.keys(hist).forEach(function (n) { o[n] = 1; });
    [1, 2, 3, 4].forEach(function (k) { base(k).items.forEach(function (r) { if (r.nama) o[r.nama] = 1; }); });
    return Object.keys(o);
  }
  function histPrice(n) { var h = hist[n]; if (h) return h.harga; for (var k = 1; k <= 4; k++) { var f = base(k).items.filter(function (r) { return r.nama === n; })[0]; if (f) return G.num(f.harga); } return 0; }

  /* ---------- aksi bawaan ---------- */
  function saveBawaan(k) { var w = clone(S.work[k]); w.noManual = ''; w.tglKw = ''; w.sppNo = ''; w.sppTgl = ''; w.tkdnTgl = ''; saved[k] = w; lsSet(K_SAVED, saved); }
  function resetKode(k) { delete saved[k]; lsSet(K_SAVED, saved); S.work[k] = norm(Object.assign(clone(FACTORY[k]), { noManual: '', tglKw: '', sppNo: '', sppTgl: '', tkdnTgl: '' })); persist(); }

  /* ---------- pilihan kode (dipakai invoice & kuitansi) ---------- */
  function kodeTiles(box, onChange) {
    function draw() {
      box.innerHTML = '';
      [1, 2, 3, 4].forEach(function (k) {
        var c = G.hitung(dataFor(k)), b = mk('button', k === S.kode ? 'on' : '', { type: 'button' });
        b.innerHTML = '<b>Kode ' + k + '</b><span>' + (c.rows.length ? 'Rp ' + G.rp(c.total) : 'Kosong') + '</span>';
        b.addEventListener('click', function () { S.kode = k; persist(); draw(); onChange(); });
        box.appendChild(b);
      });
    }
    draw();
    return { redraw: draw };
  }

  /* ---------- bagikan / unduh ---------- */
  function guard(fn) { return async function () { try { await fn(); } catch (e) { if (e && e.name !== 'AbortError') toast(e.message || String(e), true); } }; }
  function canShare() { return !!navigator.canShare; }

  root.App = {
    G: G, $: $, clone: clone, mk: mk, lab: lab, money: money, toast: toast, saveBlob: saveBlob, today: today, nav: nav, DOCS: DOCS,
    S: S, umum: umum, spp: spp, tkdnSet: tkdnSet, persistTkdn: persistTkdn, dataTkdn: dataTkdn, persist: persist, persistUmum: persistUmum, persistSpp: persistSpp, dataFor: dataFor, cek: cek, auto: auto,
    remember: remember, names: names, histPrice: histPrice, saveBawaan: saveBawaan, resetKode: resetKode, kodeTiles: kodeTiles, guard: guard, canShare: canShare,
    nama: function (k) { return 'Invoice ' + dataFor(k).noInvoice; }
  };
})(window);
