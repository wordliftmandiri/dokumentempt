/* docgen.js — logika pengisi template Word (Invoice)
 * Template .docx asli dibuka (zip), bagian yang berubah diganti, lalu di-zip lagi.
 * Semua berjalan di browser; tidak ada data yang dikirim ke server.
 */
(function (root) {
  'use strict';

  var W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
  var XMLNS = 'http://www.w3.org/XML/1998/namespace';
  var NSDECL = 'xmlns:w="' + W + '"';
  var MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  var BULAN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

  /* ---------- utilitas angka & tanggal ---------- */
  function num(s) { return parseInt(String(s == null ? '' : s).replace(/\D/g, ''), 10) || 0; }
  function rp(n) { return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
  function pad(n) { return (n < 10 ? '0' : '') + n; }

  function tglID(iso) {
    if (!iso) return '';
    var p = iso.split('-').map(Number);
    return p[2] + ' ' + BULAN[p[1] - 1] + ' ' + p[0];
  }
  function addMonths(iso, n) {
    if (!iso) return '';
    var p = iso.split('-').map(Number);
    var y = p[0], m = p[1] - 1 + n, d = p[2];
    y += Math.floor(m / 12); m = ((m % 12) + 12) % 12;
    var last = new Date(y, m + 1, 0).getDate();
    return y + '-' + pad(m + 1) + '-' + pad(Math.min(d, last));
  }
  /* 2610271 = tahun(26) bulan(10) tanggal(27) + nomor tempat(1) */
  function nomorInvoice(iso, urut) {
    if (!iso) return '';
    var p = iso.split('-');
    return p[0].slice(2) + p[1] + p[2] + urut;
  }

  var SATUAN = ['', 'Satu', 'Dua', 'Tiga', 'Empat', 'Lima', 'Enam', 'Tujuh', 'Delapan', 'Sembilan', 'Sepuluh', 'Sebelas'];
  function tb(n) {
    if (n < 12) return SATUAN[n];
    if (n < 20) return tb(n - 10) + ' Belas';
    if (n < 100) return tb(Math.floor(n / 10)) + ' Puluh' + (n % 10 ? ' ' + tb(n % 10) : '');
    if (n < 200) return 'Seratus' + (n > 100 ? ' ' + tb(n - 100) : '');
    if (n < 1000) return tb(Math.floor(n / 100)) + ' Ratus' + (n % 100 ? ' ' + tb(n % 100) : '');
    if (n < 2000) return 'Seribu' + (n > 1000 ? ' ' + tb(n - 1000) : '');
    var unit = [[1e12, 'Triliun'], [1e9, 'Miliar'], [1e6, 'Juta'], [1e3, 'Ribu']];
    for (var i = 0; i < unit.length; i++) {
      if (n >= unit[i][0]) {
        var q = Math.floor(n / unit[i][0]), r = n % unit[i][0];
        return tb(q) + ' ' + unit[i][1] + (r ? ' ' + tb(r) : '');
      }
    }
    return '';
  }
  function terbilang(n) {
    n = Math.round(n);
    if (n <= 0) return 'Nol Rupiah';
    return tb(n) + ' Rupiah';
  }

  /* ---------- perhitungan ---------- */
  function hitung(d) {
    var rows = (d.items || []).map(function (it) {
      var q = num(it.qty), h = num(it.harga);
      return { nama: String(it.nama || '').trim(), satuan: String(it.satuan || '').trim() || 'Pcs', q: q, h: h, tot: q * h };
    }).filter(function (r) { return !!r.nama; });
    var jumlah = rows.reduce(function (s, r) { return s + r.tot; }, 0);
    var pot = num(d.potongan);
    var sub = Math.max(0, jumlah - pot);
    var pct = parseFloat(String(d.ppn == null ? 11 : d.ppn).replace(',', '.')) || 0;
    var ppn = Math.round(sub * pct / 100);
    return { rows: rows, jumlah: jumlah, pot: pot, sub: sub, pct: pct, ppn: ppn, total: sub + ppn };
  }

  /* ---------- utilitas XML ---------- */
  function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  function kids(el, name) {
    return Array.prototype.filter.call(el.childNodes, function (n) {
      return n.nodeType === 1 && n.namespaceURI === W && (!name || n.localName === name);
    });
  }
  function ptext(p) {
    return Array.prototype.map.call(p.getElementsByTagNameNS(W, 't'), function (t) { return t.textContent; }).join('');
  }
  function frag(doc, xml) {
    var d = new DOMParser().parseFromString('<root ' + NSDECL + '>' + xml + '</root>', 'application/xml');
    if (d.getElementsByTagName('parsererror').length) throw new Error('XML tidak valid');
    return Array.prototype.map.call(d.documentElement.childNodes, function (n) { return doc.importNode(n, true); });
  }
  function run(t, o) {
    o = o || {};
    var f = o.f || 'Calibri';
    return '<w:r><w:rPr><w:rFonts w:ascii="' + f + '" w:hAnsi="' + f + '"/>' +
      (o.b ? '<w:b/><w:bCs/>' : '') + (o.i ? '<w:i/><w:iCs/>' : '') +
      '<w:sz w:val="' + (o.sz || 19) + '"/></w:rPr><w:t xml:space="preserve">' + esc(t) + '</w:t></w:r>';
  }
  function tabRun() { return '<w:r><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="19"/></w:rPr><w:tab/></w:r>'; }

  function clearAttrIds(p) {
    p.removeAttributeNS(W14, 'paraId');
    p.removeAttributeNS(W14, 'textId');
  }
  /* Ganti seluruh isi paragraf dengan satu run berformat sama dengan run teks pertama */
  function setParaText(p, text, opt) {
    opt = opt || {};
    var doc = p.ownerDocument;
    var r0 = Array.prototype.filter.call(p.getElementsByTagNameNS(W, 'r'), function (r) {
      return r.getElementsByTagNameNS(W, 't').length > 0;
    })[0];
    var rPr = r0 ? kids(r0, 'rPr')[0] : null;
    rPr = rPr ? rPr.cloneNode(true) : null;
    if (rPr) kids(rPr).forEach(function (e) {
      if (e.localName === 'spacing' || e.localName === 'lang') rPr.removeChild(e);
    });
    if (rPr && opt.font) {
      var rf = kids(rPr, 'rFonts')[0];
      if (rf) { rf.setAttributeNS(W, 'w:ascii', opt.font); rf.setAttributeNS(W, 'w:hAnsi', opt.font); }
    }
    if (rPr && opt.sz) {
      kids(rPr, 'sz').concat(kids(rPr, 'szCs')).forEach(function (e) { e.setAttributeNS(W, 'w:val', String(opt.sz)); });
    }
    kids(p).forEach(function (c) { if (c.localName !== 'pPr') p.removeChild(c); });
    var r = doc.createElementNS(W, 'w:r');
    if (rPr) r.appendChild(rPr);
    var t = doc.createElementNS(W, 'w:t');
    t.setAttributeNS(XMLNS, 'xml:space', 'preserve');
    t.textContent = text;
    r.appendChild(t);
    p.appendChild(r);
  }

  /* ---------- pengisi Invoice ---------- */
  /* Blok kanan-atas dibuat sebagai tabel tanpa garis (3 kolom) agar rapi di Word, LibreOffice, maupun pratinjau web */
  function cellXml(w, content) {
    return '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/></w:tcPr>' + content + '</w:tc>';
  }
  function cellP(runs) {
    return '<w:p><w:pPr><w:spacing w:before="20" w:after="0"/></w:pPr>' + runs + '</w:p>';
  }
  function hdrTable(rows) {
    var W1 = 7000, W2 = 1500, W3 = 3100;
    var xml = '<w:tbl><w:tblPr><w:tblW w:w="' + (W1 + W2 + W3) + '" w:type="dxa"/><w:tblInd w:w="60" w:type="dxa"/>' +
      '<w:tblBorders><w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tblBorders>' +
      '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="' + W1 + '"/><w:gridCol w:w="' + W2 + '"/><w:gridCol w:w="' + W3 + '"/></w:tblGrid>';
    rows.forEach(function (r) {
      var o = { f: 'Calibri Light', b: r.b };
      xml += '<w:tr>' + cellXml(W1, cellP(r.left || '')) + cellXml(W2, cellP(run(r.label, o))) +
        cellXml(W3, cellP(run(': ' + r.value, { f: 'Calibri Light', b: r.b, sz: r.sz || 19 }))) + '</w:tr>';
    });
    // paragraf pemisah sangat tipis: tanpa ini Word/LibreOffice menggabungkan tabel ini dengan tabel barang di bawahnya
    return xml + '</w:tbl><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="40" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>';
  }

  /* Hilangkan halaman kosong di belakang (section terakhir hanya berisi paragraf kosong) */
  function dropBlankTail(doc) {
    var body = doc.getElementsByTagNameNS(W, 'body')[0];
    var finalSect = kids(body, 'sectPr')[0];
    var withSect = kids(body, 'p').filter(function (p) {
      var pp = kids(p, 'pPr')[0]; return pp && kids(pp, 'sectPr').length;
    });
    var lastP = withSect[withSect.length - 1];
    if (lastP && finalSect) {
      var trailing = [], n = lastP.nextSibling;
      while (n && n !== finalSect) { trailing.push(n); n = n.nextSibling; }
      var kosong = trailing.every(function (t) {
        return t.nodeType !== 1 || (t.localName === 'p' && !ptext(t).trim() && !t.getElementsByTagNameNS(W, 'drawing').length);
      });
      if (kosong) {
        var ls = kids(kids(lastP, 'pPr')[0], 'sectPr')[0];
        trailing.forEach(function (t) { body.removeChild(t); });
        body.replaceChild(ls, finalSect);
      }
    }
  }

  function fillInvoice(doc, d) {
    var c = hitung(d);
    if (c.rows.length > 5) throw new Error('Template invoice hanya menampung 5 baris barang.');
    if (c.rows.length === 0) throw new Error('Isi minimal satu barang.');

    var body = doc.getElementsByTagNameNS(W, 'body')[0];
    var top = kids(body, 'p');
    function findTop(prefix) {
      var p = top.filter(function (x) { return ptext(x).trim().indexOf(prefix) === 0; })[0];
      if (!p) throw new Error('Template tidak cocok (tidak ketemu: ' + prefix + ')');
      return p;
    }

    /* 1. Nomor invoice */
    var pInv = findTop('INVOICE');
    var seenTab = false, first = null;
    kids(pInv, 'r').forEach(function (r) {
      if (seenTab) { if (!first) first = r; pInv.removeChild(r); }
      else if (r.getElementsByTagNameNS(W, 'tab').length) seenTab = true;
    });
    var nr = doc.createElementNS(W, 'w:r');
    var rp0 = first && kids(first, 'rPr')[0];
    if (rp0) {
      rp0 = rp0.cloneNode(true);
      kids(rp0).forEach(function (e) { if (e.localName === 'lang') rp0.removeChild(e); });
      nr.appendChild(rp0);
    }
    var nt = doc.createElementNS(W, 'w:t'); nt.textContent = d.noInvoice; nr.appendChild(nt);
    pInv.appendChild(nr);

    /* 1b. Kolom "Kepada :" terlalu sempit (kata terpotong jadi "Kepad / a :"): lebarkan, total lebar tetap */
    Array.prototype.forEach.call(doc.getElementsByTagNameNS(W, 'cols'), function (cs) {
      var cl = kids(cs, 'col');
      if (cl.length === 2 && cl[0].getAttributeNS(W, 'w') === '811') {
        cl[0].setAttributeNS(W, 'w:w', '1500');
        cl[0].setAttributeNS(W, 'w:space', '4268');
      }
    });

    /* 2. Blok pelanggan + tanggal/PO/keterangan (ditulis ulang agar rapi & tidak bergantung pada wrap) */
    var p1 = findTop('DEPUTI'), p2 = findTop('Gd.'), p3 = findTop('Kecamatan');
    var drawingRun = kids(p3, 'r').filter(function (r) { return r.getElementsByTagNameNS(W, 'drawing').length; })[0];
    var po = String(d.noPO || '');
    var nodes = frag(doc, hdrTable([
      { left: run(d.kepada || '', { b: true }), label: 'Tanggal', value: tglID(d.tanggal) },
      { left: run(d.alamat1 || ''), label: 'PO', value: po, sz: po.length > 26 ? 15 : 19 },
      { left: run(d.alamat2 || ''), label: 'Tanggal PO', value: tglID(d.tglPO) },
      { label: 'Jatuh Tempo', value: tglID(d.jatuhTempo) },
      { label: 'KETERANGAN', value: d.keterangan || '', b: true }
    ]));
    if (drawingRun) { // pertahankan watermark logo (ditaruh di baris ke-3 sisi kiri)
      var anchor = drawingRun.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'anchor')[0];
      if (anchor) anchor.setAttribute('layoutInCell', '0');
      var cellsP = nodes[0].getElementsByTagNameNS(W, 'p'); // nodes[0] = tabel
      var pd = cellsP[6]; // baris 3, kolom kiri
      pd.insertBefore(drawingRun, kids(pd, 'pPr')[0].nextSibling);
    }
    nodes.forEach(function (n) { body.insertBefore(n, p1); });
    [p1, p2, p3].forEach(function (p) { body.removeChild(p); });

    /* 3. Tabel barang */
    var tbls = Array.prototype.slice.call(doc.getElementsByTagNameNS(W, 'tbl'));
    var main = tbls.filter(function (t) { return /NAMA/.test(t.textContent) && /SATUAN/.test(t.textContent); })[0];
    var sum = tbls.filter(function (t) { return /SUB/.test(t.textContent) && /TOTAL/.test(t.textContent); })[0];
    if (!main || !sum) throw new Error('Template tidak cocok (tabel)');
    var trs = kids(main, 'tr');
    if (trs.length !== 7) throw new Error('Template tidak cocok (jumlah baris tabel = ' + trs.length + ')');
    function cellP(tr, i) { return kids(kids(tr, 'tc')[i], 'p')[0]; }
    function cellOf(tr, i) { return kids(tr, 'tc')[i]; }

    var protoUsed = [];
    for (var i = 0; i < 7; i++) protoUsed.push(cellP(trs[2], i).cloneNode(true));
    var protoBlankName = cellP(trs[4], 0).cloneNode(true); // "0" rata kanan seperti aslinya
    for (var r = 0; r < 5; r++) {
      var tr = trs[r + 1], row = c.rows[r];
      // semua baris barang: tinggi sama & format paragraf sama (baris kosong tetap menampilkan 0 / Pcs / 0 / Rp - / Rp - seperti aslinya)
      var th = kids(kids(tr, 'trPr')[0], 'trHeight')[0];
      th.setAttributeNS(W, 'w:val', '348');
      var texts = row ? [row.nama, row.satuan, String(row.q), 'Rp', row.h ? rp(row.h) : '-', 'Rp', row.tot ? rp(row.tot) : '-']
                      : ['0', 'Pcs', '0', 'Rp', '-', 'Rp', '-'];
      for (var k = 0; k < 7; k++) {
        var tc = cellOf(tr, k), old = cellP(tr, k);
        var np = (!row && k === 0 ? protoBlankName : protoUsed[k]).cloneNode(true);
        clearAttrIds(np);
        // Calibri biasa (bukan Light): ada padanan metrik di LibreOffice/pratinjau, jadi lebar teks sama dengan di Word.
        // Nama panjang diperkecil bertahap supaya tetap satu baris.
        var L = texts[k].length, nameSz = (k === 0 && row) ? (L > 60 ? 15 : L > 54 ? 17 : L > 48 ? 19 : 0) : 0;
        setParaText(np, texts[k], { font: 'Calibri', sz: nameSz });
        tc.replaceChild(np, old);
      }
    }
    // jarak satu baris antara daftar barang dan "Potongan Harga"
    kids(trs[6], 'tc').forEach(function (tc) {
      kids(tc, 'p').forEach(function (pp) {
        var pr = kids(pp, 'pPr')[0], sp0 = pr && kids(pr, 'spacing')[0];
        if (sp0 && sp0.getAttributeNS(W, 'before')) sp0.setAttributeNS(W, 'w:before', '515');
      });
    });
    // potongan harga
    setParaText(cellP(trs[6], 6), c.pot ? rp(c.pot) : '-');

    /* 4. Subtotal / PPN / Total */
    var sp = Array.prototype.slice.call(sum.getElementsByTagNameNS(W, 'p'));
    function idxOf(fn) {
      var i = sp.findIndex(function (p) { return fn(ptext(p).replace(/\s+/g, ' ').trim()); });
      if (i < 0) throw new Error('Template tidak cocok (ringkasan)');
      return i;
    }
    var iSub = idxOf(function (t) { return t === 'SUB TOTAL'; });
    var iPpn = idxOf(function (t) { return t.indexOf('PPN') === 0; });
    var iTot = idxOf(function (t) { return t === 'TOTAL'; });
    setParaText(sp[iSub + 2], rp(c.sub));
    setParaText(sp[iPpn], 'PPN ' + String(c.pct).replace('.', ',') + '%');
    setParaText(sp[iPpn + 2], rp(c.ppn));
    setParaText(sp[iTot + 2], rp(c.total));

    /* 4b. Kotak rekening bank: garis tabel asli (gambar kotak aslinya sering terpotong) */
    var sTrs = kids(sum, 'tr');
    Array.prototype.forEach.call(sum.getElementsByTagNameNS(W, 'drawing'), function (dr) {
      var rn = dr; // naik sampai <w:r> (gambar dibungkus mc:AlternateContent dengan cadangan VML)
      while (rn && !(rn.namespaceURI === W && rn.localName === 'r')) rn = rn.parentNode;
      if (rn && rn.parentNode) rn.parentNode.removeChild(rn);
    });
    function bd(side) { return '<w:' + side + ' w:val="single" w:sz="8" w:space="0" w:color="000000"/>'; }
    function setBorders(tr, cellIdx, sides) {
      var tcPr = kids(kids(tr, 'tc')[cellIdx], 'tcPr')[0];
      frag(doc, '<w:tcBorders>' + sides.map(bd).join('') + '</w:tcBorders>').forEach(function (n) {
        tcPr.insertBefore(n, kids(tcPr, 'tcW')[0].nextSibling);
      });
    }
    setBorders(sTrs[0], 0, ['top', 'left']);   setBorders(sTrs[0], 1, ['top', 'right']);
    setBorders(sTrs[1], 0, ['left', 'bottom']); setBorders(sTrs[1], 1, ['bottom', 'right']);
    kids(kids(sTrs[0], 'trPr')[0], 'trHeight')[0].setAttributeNS(W, 'w:val', '360');
    kids(kids(sTrs[1], 'trPr')[0], 'trHeight')[0].setAttributeNS(W, 'w:val', '340');

    /* 5. Terbilang */
    var pTb = findTop('Terbilang');
    kids(pTb).forEach(function (e) { if (e.localName !== 'pPr') pTb.removeChild(e); });
    frag(doc, run('Terbilang : ', { f: 'Calibri Light', sz: 21 }) +
      run(terbilang(c.total), { f: 'Calibri Light', sz: 21, b: true, i: true })).forEach(function (n) { pTb.appendChild(n); });

    dropBlankTail(doc);
  }

  /* ---------- pengisi Kuitansi ---------- */
  function kuitansiNo(noInvoice, iso) { return noInvoice + '/KWT/UML-BRIN/' + (iso ? iso.slice(0, 4) : ''); }

  function setLastRunText(p, text) {
    var rs = Array.prototype.filter.call(p.getElementsByTagNameNS(W, 'r'), function (r) { return r.getElementsByTagNameNS(W, 't').length; });
    var last = rs[rs.length - 1];
    if (!last) throw new Error('Template kuitansi tidak cocok (angka)');
    var ts = last.getElementsByTagNameNS(W, 't');
    ts[0].textContent = text;
    for (var i = ts.length - 1; i > 0; i--) last.removeChild(ts[i]);
  }

  function fillKuitansi(doc, d) {
    var c = hitung(d);
    if (c.rows.length === 0) throw new Error('Isi minimal satu barang di invoice dulu.');
    var all = Array.prototype.slice.call(doc.getElementsByTagNameNS(W, 'p')); // termasuk kotak teks (Choice & Fallback)
    var cnt = { no: 0, kepada: 0, tb: 0, untuk: 0, tgl: 0 }, amounts = [];
    var noKw = kuitansiNo(d.noInvoice, d.tanggal);
    all.forEach(function (p) {
      var t = ptext(p).trim();
      if (t === 'USAHA MANDIRI LIFT') { setParaText(p, 'USAHA MANDIRI LIFT'); } // satu run: judul tidak lagi "naik-turun"
      else if (/^\d{7}\/KWT\//.test(t)) { setParaText(p, noKw); cnt.no++; }
      else if (t === 'DEPUTI BIDANG STRUKTUR RISET DAN INOVASI BRIN') { setParaText(p, d.kepada || ''); cnt.kepada++; }
      else if (/^[A-Z][A-Za-z ]* Rupiah$/.test(t)) { setParaText(p, terbilang(c.total)); cnt.tb++; }
      else if (/^Untuk Pembayaran/.test(t)) {
        var tabs = 0, proto = null;
        kids(p, 'r').forEach(function (r) {
          if (tabs >= 2) { if (!proto && r.getElementsByTagNameNS(W, 't').length) proto = r; p.removeChild(r); }
          else if (r.getElementsByTagNameNS(W, 'tab').length) tabs++;
        });
        var rPr = proto && kids(proto, 'rPr')[0];
        rPr = rPr ? rPr.cloneNode(true) : null;
        if (rPr) kids(rPr).forEach(function (e) { if (e.localName === 'spacing' || e.localName === 'lang') rPr.removeChild(e); });
        var r = doc.createElementNS(W, 'w:r');
        if (rPr) r.appendChild(rPr);
        var tt = doc.createElementNS(W, 'w:t');
        tt.setAttributeNS(XMLNS, 'xml:space', 'preserve');
        tt.textContent = (d.uraian ? String(d.uraian).trim() + ' ' : '') + 'sesuai dengan Nomor PO - ' + (d.noPO || '');
        r.appendChild(tt); p.appendChild(r); cnt.untuk++;
      }
      else if (/^Depok,/.test(t)) { setParaText(p, 'Depok, ' + tglID(d.tanggalKw || d.tanggal)); cnt.tgl++; }
      else if (/^Rp(-|[\d.]+)$/.test(t)) amounts.push(p);
    });
    if (!cnt.no || !cnt.kepada || !cnt.tb || !cnt.untuk || !cnt.tgl) throw new Error('Template kuitansi tidak cocok');
    if (amounts.length !== 8) throw new Error('Template kuitansi tidak cocok (angka: ' + amounts.length + ')');
    // urutan di dokumen: Harga, Discount, Etax (dua salinan) lalu Harga Total (dua salinan)
    var v = [rp(c.jumlah), c.pot ? rp(c.pot) : '-', rp(c.ppn)];
    var vals = v.concat(v, [rp(c.total), rp(c.total)]);
    function tabEl(p) {
      var pr = kids(p, 'pPr')[0], tb0 = pr && kids(pr, 'tabs')[0];
      return tb0 && kids(tb0, 'tab')[0];
    }
    amounts.forEach(function (p, i) {
      setLastRunText(p, vals[i]);
      // baris Discount: tab bawaan dibuat untuk tanda "-" (sempit); bila ada angka, pakai tab seperti baris Harga
      if (c.pot && (i === 1 || i === 4)) {
        var src = tabEl(amounts[i - 1]), dst = tabEl(p);
        if (src && dst) dst.setAttributeNS(W, 'w:pos', src.getAttributeNS(W, 'pos'));
      }
    });
    dropBlankTail(doc);
  }


  /* ---------- pengisi Surat Permohonan Pembayaran ---------- */
  var ROMAWI = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII'];
  function nomorSurat(urut, iso) {
    var p = (iso || '').split('-').map(Number);
    return (urut || '01') + '/AMK/SPP/' + (p[1] ? ROMAWI[p[1] - 1] : '') + '/' + (p[0] || '');
  }
  function tglSlash(iso) { if (!iso) return ''; var p = iso.split('-'); return p[2] + '/' + p[1] + '/' + p[0]; }

  function textRuns(p) { return kids(p, 'r').filter(function (r) { return r.getElementsByTagNameNS(W, 't').length > 0; }); }
  function cleanRPr(r) {
    var rPr = r && kids(r, 'rPr')[0];
    rPr = rPr ? rPr.cloneNode(true) : null;
    if (rPr) kids(rPr).forEach(function (e) { if (e.localName === 'spacing' || e.localName === 'lang') rPr.removeChild(e); });
    return rPr;
  }
  function mkRun(doc, rPr, text, o) {
    o = o || {};
    var r = doc.createElementNS(W, 'w:r'), pr = rPr ? rPr.cloneNode(true) : doc.createElementNS(W, 'w:rPr');
    if (o.i) pr.insertBefore(doc.createElementNS(W, 'w:iCs'), pr.firstChild), pr.insertBefore(doc.createElementNS(W, 'w:i'), pr.firstChild);
    if (o.b) pr.insertBefore(doc.createElementNS(W, 'w:b'), pr.firstChild);
    if (pr.firstChild) r.appendChild(pr);
    var t = doc.createElementNS(W, 'w:t'); t.setAttributeNS(XMLNS, 'xml:space', 'preserve'); t.textContent = text;
    r.appendChild(t); return r;
  }
  /* Ganti hanya nilai setelah tanda ":" (label & tab tetap utuh) */
  function setAfterColon(p, value) {
    var doc = p.ownerDocument, rs = textRuns(p), idx = -1, i;
    for (i = 0; i < rs.length; i++) if (ptext(rs[i]).indexOf(':') >= 0) { idx = i; break; }
    if (idx < 0) throw new Error('Template surat tidak cocok (' + ptext(p).slice(0, 20) + ')');
    var cur = rs[idx], ct = ptext(cur), tail = ct.slice(ct.indexOf(':') + 1), rm = rs.slice(idx + 1);
    var rPr = cleanRPr(rm.length ? rm[rm.length - 1] : cur);
    rm.forEach(function (r) { p.removeChild(r); });
    var nx = cur.nextSibling; // sisa run di belakang ":" (mis. run berisi tab saja) ikut dibuang
    while (nx) { var nn = nx.nextSibling; if (nx.nodeType === 1 && nx.localName === 'r') p.removeChild(nx); nx = nn; }
    var sep = ' ';
    if (tail.replace(/\s/g, '') === '') { if (/\s$/.test(tail)) sep = ''; }
    else { // nilai ikut berada dalam run yang memuat ":" -> potong sampai ":"
      var ts = cur.getElementsByTagNameNS(W, 't');
      ts[0].textContent = ct.slice(0, ct.indexOf(':') + 1);
      for (var k = ts.length - 1; k > 0; k--) cur.removeChild(ts[k]);
    }
    p.appendChild(mkRun(doc, rPr, sep + value));
  }

  var SPP0 = { tujuan: 'Deputi Infrastruktur Riset dan Inovasi Badan Riset dan Inovasi Nasional (BRIN) Jakarta', lampiran: '1 (satu) berkas',
    bank: 'Bank BNI', atasNama: 'CV. Usaha Mandiri Lift', noRek: '1.688.568.246', penandatangan: 'Heru Santoso', jabatan: 'Direktur Utama' };

  function fillPermohonan(doc, d) {
    var c = hitung(d);
    if (c.rows.length === 0) throw new Error('Isi minimal satu barang di invoice dulu.');
    var s = Object.assign({}, SPP0, d.spp || {});
    var cnt = { tgl: 0, no: 0, isi: 0 };
    Array.prototype.slice.call(doc.getElementsByTagNameNS(W, 'p')).forEach(function (p) {
      var t = ptext(p).trim();
      if (t === 'USAHA MANDIRI LIFT') { setParaText(p, 'USAHA MANDIRI LIFT'); } // judul kop: satu run agar rata
      else if (/^Jakarta,/.test(t)) { setParaText(p, 'Jakarta, ' + tglID(d.tanggalSurat)); cnt.tgl++; }
      else if (/^Nomor:/.test(t)) { setAfterColon(p, d.noSurat); cnt.no++; }
      else if (/^Lampiran:/.test(t)) { if (s.lampiran !== SPP0.lampiran) setAfterColon(p, s.lampiran); }
      else if (/^Deputi Infrastruktur/.test(t)) { if (s.tujuan !== SPP0.tujuan) setParaText(p, s.tujuan); }
      else if (/^Nama Bank/.test(t)) { if (s.bank !== SPP0.bank) setAfterColon(p, s.bank); }
      else if (/^Atas Nama/.test(t)) { if (s.atasNama !== SPP0.atasNama) setAfterColon(p, s.atasNama); }
      else if (/^Nomor Rekening/.test(t)) { if (s.noRek !== SPP0.noRek) setAfterColon(p, s.noRek); }
      else if (t === SPP0.penandatangan) { if (s.penandatangan !== SPP0.penandatangan) setParaText(p, s.penandatangan); }
      else if (t === SPP0.jabatan) { if (s.jabatan !== SPP0.jabatan) setParaText(p, s.jabatan); }
      else if (/^Sehubungan dengan/.test(t)) {
        var base = cleanRPr(textRuns(p)[0]);
        kids(p).forEach(function (e) { if (e.localName !== 'pPr') p.removeChild(e); });
        var paket = String(d.paket || '').trim().replace(/\.$/, '');
        p.appendChild(mkRun(doc, base, 'Sehubungan dengan telah selesai nya Paket Pekerjaan : ' + paket + '. Dengan Nomor: ' + (d.noEP || '') +
          ' Tanggal ' + tglSlash(d.tglEP) + ', telah selesai 100% pekerjaanya. Maka dengan ini kami mengajukan pembayaran atas pekerjaan tersebut sebesar '));
        p.appendChild(mkRun(doc, base, 'Rp. ' + rp(c.total) + ', ( ', { b: true }));
        p.appendChild(mkRun(doc, base, terbilang(c.total), { b: true, i: true }));
        p.appendChild(mkRun(doc, base, ',-)', { b: true }));
        cnt.isi++;
      }
    });
    if (!cnt.tgl || !cnt.no || !cnt.isi) throw new Error('Template surat tidak cocok');
  }


  /* ---------- pengisi Surat TKDN ---------- */
  var TKDN0 = { nama: 'Heru Santoso', jabatan: 'Direktur', alamat: 'Cabang : Jln. H. Maksum Rt 04 Rw 02 No 38D Sawangan Baru Depok – Jabar.16511' };
  function norm1(t) { return String(t).replace(/\s+/g, ' ').trim(); }

  function fillTkdn(doc, d) {
    var s = Object.assign({}, TKDN0, d.tkdn || {});
    var cnt = { paket: 0, po: 0, tpo: 0, tgl: 0, kal: 0 };
    Array.prototype.slice.call(doc.getElementsByTagNameNS(W, 'p')).forEach(function (p) {
      var t = norm1(ptext(p));
      if (t === 'USAHA MANDIRI LIFT') { setParaText(p, 'USAHA MANDIRI LIFT'); }
      else if (/^Nama:/.test(t)) { if (s.nama !== TKDN0.nama) setAfterColon(p, s.nama); }
      else if (/^Jabatan:/.test(t)) { if (s.jabatan !== TKDN0.jabatan) setAfterColon(p, s.jabatan); }
      else if (/^Alamat\s*:/.test(t)) { if (norm1(s.alamat) !== norm1(TKDN0.alamat)) setAfterColon(p, s.alamat); }
      else if (/^Paket Pekerjaan:/.test(t)) { setAfterColon(p, String(d.paket || '').trim()); cnt.paket++; }
      else if (/^Nomor PO:/.test(t)) { setAfterColon(p, d.noPO || ''); cnt.po++; }
      else if (/^Tanggal PO:/.test(t)) { setAfterColon(p, tglID(d.tglPO)); cnt.tpo++; }
      else if (/^Sehubungan dengan .*, maka dengan ini kami menyatakan/.test(t)) {
        // pekerjaan di kalimat ini otomatis = paket pekerjaan (template asli berisi sisa surat lama: "Perbaikan Sistem Pompa Submersible Intake")
        var kerja = String(s.pekerjaan || d.paket || '').trim().replace(/\.$/, '');
        setParaText(p, 'Sehubungan dengan ' + kerja + t.slice(t.indexOf(', maka dengan ini')));
        cnt.kal++;
      }
      else if (/^Jakarta,/.test(t)) { setParaText(p, 'Jakarta, ' + tglID(d.tanggalSurat)); cnt.tgl++; }
      else if (t === TKDN0.nama) { if (s.nama !== TKDN0.nama) setParaText(p, s.nama); }
      else if (t === TKDN0.jabatan) { if (s.jabatan !== TKDN0.jabatan) setParaText(p, s.jabatan); }
    });
    if (!cnt.paket || !cnt.po || !cnt.tpo || !cnt.tgl || !cnt.kal) throw new Error('Template TKDN tidak cocok');
    // paragraf kosong di ujung dokumen bisa menyeberang ke halaman kedua (kosong) saat isian lebih panjang -> buang
    var body = doc.getElementsByTagNameNS(W, 'body')[0], last = body.lastChild;
    while (last && !(last.nodeType === 1 && last.localName === 'sectPr')) last = last.previousSibling;
    var q = last && last.previousSibling;
    while (q && (q.nodeType !== 1 || (q.localName === 'p' && !ptext(q).trim() && !q.getElementsByTagNameNS(W, 'drawing').length && !q.getElementsByTagNameNS(W, 'pict').length && !kids(q, 'pPr').some(function (pp) { return kids(pp, 'sectPr').length; })))) {
      var pv = q.previousSibling; body.removeChild(q); q = pv;
    }
    dropBlankTail(doc);
  }

  async function buildDoc(fill, d, tplB64) {
    var zip = await JSZip.loadAsync(tplB64, { base64: true });
    var xml = await zip.file('word/document.xml').async('string');
    var doc = new DOMParser().parseFromString(xml, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('Template rusak');
    fill(doc, d);
    var out = new XMLSerializer().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, '');
    zip.file('word/document.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + out);
    return zip.generateAsync({ type: 'blob', mimeType: MIME, compression: 'DEFLATE' });
  }
  function buildInvoice(d, tplB64) { return buildDoc(fillInvoice, d, tplB64 || root.TEMPLATES.invoice); }
  function buildKuitansi(d, tplB64) { return buildDoc(fillKuitansi, d, tplB64 || root.TEMPLATES.kuitansi); }
  function buildTkdn(d, tplB64) { return buildDoc(fillTkdn, d, tplB64 || root.TEMPLATES.tkdn); }
  function buildPermohonan(d, tplB64) { return buildDoc(fillPermohonan, d, tplB64 || root.TEMPLATES.permohonan); }

  root.DocGen = {
    MIME: MIME, num: num, rp: rp, tglID: tglID, addMonths: addMonths, nomorInvoice: nomorInvoice,
    terbilang: terbilang, hitung: hitung, buildInvoice: buildInvoice, buildKuitansi: buildKuitansi, kuitansiNo: kuitansiNo,
    buildPermohonan: buildPermohonan, buildTkdn: buildTkdn, TKDN0: TKDN0, nomorSurat: nomorSurat, SPP0: SPP0
  };
})(typeof window !== 'undefined' ? window : globalThis);
