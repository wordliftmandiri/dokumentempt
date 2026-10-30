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
    }).filter(function (r) { return r.nama || r.h > 0; });
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
  function setParaText(p, text) {
    var doc = p.ownerDocument;
    var r0 = Array.prototype.filter.call(p.getElementsByTagNameNS(W, 'r'), function (r) {
      return r.getElementsByTagNameNS(W, 't').length > 0;
    })[0];
    var rPr = r0 ? kids(r0, 'rPr')[0] : null;
    rPr = rPr ? rPr.cloneNode(true) : null;
    if (rPr) kids(rPr).forEach(function (e) {
      if (e.localName === 'spacing' || e.localName === 'lang') rPr.removeChild(e);
    });
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
        setParaText(np, texts[k]);
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

    /* 6. Hilangkan halaman kosong di belakang (section terakhir hanya berisi paragraf kosong) */
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

  async function buildInvoice(d, tplB64) {
    var zip = await JSZip.loadAsync(tplB64 || root.TEMPLATES.invoice, { base64: true });
    var xml = await zip.file('word/document.xml').async('string');
    var doc = new DOMParser().parseFromString(xml, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('Template rusak');
    fillInvoice(doc, d);
    var out = new XMLSerializer().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, '');
    zip.file('word/document.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + out);
    return zip.generateAsync({ type: 'blob', mimeType: MIME, compression: 'DEFLATE' });
  }

  root.DocGen = {
    MIME: MIME, num: num, rp: rp, tglID: tglID, addMonths: addMonths, nomorInvoice: nomorInvoice,
    terbilang: terbilang, hitung: hitung, buildInvoice: buildInvoice
  };
})(typeof window !== 'undefined' ? window : globalThis);
