/* ── Estado global ────────────────────────────────────────────── */
let fa = 'todos';
let dtable;

const STORE_KEY = 'reaf_cambios_' + document.title.replace(/\s+/g,'_');

/* ── getUUID(): obtiene UUID del <tr> robusto con scrollX ─────── */
function getUUID(tr) {
  if (!tr) return '';
  let uuid = tr.getAttribute('data-uuid') || '';
  if (!uuid) {
    const tdUu = tr.querySelector('td.uu[title]');
    uuid = tdUu ? tdUu.getAttribute('title') : '';
  }
  if (!uuid && tr.cells[0]) uuid = tr.cells[0].getAttribute('title') || '';
  return uuid.trim().toUpperCase();
}

/* ── localStorage: guardar / borrar / leer cambios ───────────── */
function guardarCambio(uuid, estatus) {
  try {
    const d = JSON.parse(localStorage.getItem(STORE_KEY) || '{}');
    d[uuid] = estatus;
    localStorage.setItem(STORE_KEY, JSON.stringify(d));
  } catch(e) { console.warn('localStorage no disponible', e); }
}
function borrarCambio(uuid) {
  try {
    const d = JSON.parse(localStorage.getItem(STORE_KEY) || '{}');
    delete d[uuid];
    localStorage.setItem(STORE_KEY, JSON.stringify(d));
  } catch(e) {}
}
function getCambios() {
  try { return JSON.parse(localStorage.getItem(STORE_KEY) || '{}'); }
  catch(e) { return {}; }
}

/* ── guardarFila(): persiste sub2+iva+sub0+estatus ───────────── */
function guardarFila(tr) {
  try {
    const uuid = getUUID(tr);
    if (!uuid) return;
    const inputs = tr.querySelectorAll('.ip');
    const sel    = tr.querySelector('.se');
    const s2  = inputs[0] ? inputs[0].value : '';
    const iva = inputs[1] ? inputs[1].value : '';
    const s0  = inputs[2] ? inputs[2].value : '';
    const estatus = sel ? sel.value : '';
    const data = JSON.parse(localStorage.getItem(STORE_KEY) || '{}');
    data[uuid] = { sub2: s2, iva: iva, sub0: s0, estatus: estatus };
    localStorage.setItem(STORE_KEY, JSON.stringify(data));
  } catch(e) { console.warn('localStorage error', e); }
}

/* ── restaurarCambios(): restaura sub2+iva+sub0+estatus al cargar */
function restaurarCambios() {
  const cambios = getCambios();
  if (!Object.keys(cambios).length) return;
  $(dtable.rows().nodes()).each(function() {
    const tr   = this;
    const uuid = getUUID(tr);
    if (!uuid || !cambios[uuid]) return;
    const fila   = cambios[uuid];
    const inputs = tr.querySelectorAll('.ip');
    if (inputs.length >= 3) {
      if (fila.sub2 !== undefined) inputs[0].value = fila.sub2;
      if (fila.iva  !== undefined) inputs[1].value = fila.iva;
      if (fila.sub0 !== undefined) inputs[2].value = fila.sub0;
      recalcularTotal(tr,
        parseFloat(inputs[0].value) || 0,
        parseFloat(inputs[1].value) || 0,
        parseFloat(inputs[2].value) || 0);
    }
    const sel = tr.querySelector('.se');
    if (sel && fila.estatus) { sel.value = fila.estatus; ce(sel, false); }
  });
  dtable.draw(false);
  actualizarDiot();
  actualizarContadores();
  loadExt();
}

/* ── exportarCambiosCSV(): descarga cambios como CSV de respaldo ─ */
function exportarCambiosCSV() {
  const cambios = getCambios();
  const keys = Object.keys(cambios);
  if (!keys.length) { alert('No hay cambios guardados.'); return; }
  let csv = 'UUID,ESTATUS_MANUAL\n';
  keys.forEach(uuid => { csv += uuid + ',' + cambios[uuid] + '\n'; });
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url; a.download = 'cambios_diot_' + STORE_KEY.slice(-8) + '.csv';
  a.click(); URL.revokeObjectURL(url);
}

/* ── actualizarDiot(): recalcula DIOT con estatus actuales ──────
   FIX: usa dtable.rows().nodes() en lugar de $('#tbl tbody tr')
        para incluir filas de TODAS las páginas, no solo la visible.
   ─────────────────────────────────────────────────────────────── */
function actualizarDiot() {
  const sec = document.getElementById('diot-section');
  if (!sec || sec.style.display === 'none') return;

  const provs = {};
  $(dtable.rows().nodes()).each(function() {
    const est = this.getAttribute('data-est') || '';
    if (!est.startsWith('ded') && !est.startsWith('efe')) return;
    const cells = this.cells;
    const razon = cells[4] ? cells[4].innerText.trim() : '';
    const rfc   = this.getAttribute('data-rfc') || '';
    if (!rfc) return;
    const s2  = parseFloat((cells[5]?.querySelector('input')?.value) || 0) || 0;
    const i16 = parseFloat((cells[6]?.querySelector('input')?.value) || 0) || 0;
    const s0  = parseFloat((cells[7]?.querySelector('input')?.value) || 0) || 0;
    const tot = parseFloat(((cells[8]?.innerText || '0').replace(/[^0-9.]/g,''))) || 0;
    if (!provs[rfc]) provs[rfc] = { razon, rfc, s2:0, i16:0, s0:0, tot:0 };
    provs[rfc].s2  += s2;  provs[rfc].i16 += i16;
    provs[rfc].s0  += s0;  provs[rfc].tot += tot;
  });

  const lista = Object.values(provs).sort((a,b) => a.razon.localeCompare(b.razon));
  const fmt   = v => v > 0 ? '$' + v.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g,',') : '-';
  let tS2=0, tI16=0, tS0=0, tTot=0;
  lista.forEach(p => { tS2+=p.s2; tI16+=p.i16; tS0+=p.s0; tTot+=p.tot; });

  let rows = lista.map((p,i) =>
    `<tr class="diot-row${i%2===0?' diot-alt':''}">
      <td class="diot-rs">${p.razon}</td><td class="diot-rfc">${p.rfc}</td>
      <td class="diot-num">${fmt(p.s2)}</td><td class="diot-num">${fmt(p.i16)}</td>
      <td class="diot-num">${fmt(p.s0)}</td><td class="diot-tot">${fmt(p.tot)}</td>
    </tr>`
  ).join('');
  rows += `<tr class="diot-totrow">
    <td colspan="2" class="diot-tl">TOTALES</td>
    <td class="diot-tot">${fmt(tS2)}</td><td class="diot-tot">${fmt(tI16)}</td>
    <td class="diot-tot">${fmt(tS0)}</td><td class="diot-tot">${fmt(tTot)}</td>
  </tr>`;
  const tbody = sec.querySelector('.diot-tbl tbody');
  if (tbody) tbody.innerHTML = rows;
}

/* ── Filtro por deducibilidad (data-est del <tr>) ─────────────── */
$.fn.dataTable.ext.search.push(function(settings, data, dataIndex) {
  if (fa === 'todos') return true;
  const row = dtable.row(dataIndex).node();
  const est = (row ? row.getAttribute('data-est') : '') || '';
  if (fa === 'ded16')       return est === 'ded16' || est === 'ded';
  if (fa === 'ded0')        return est === 'ded0';
  if (fa === 'ded160')      return est === 'ded160';
  if (fa === 'efe')         return est.startsWith('efe');
  if (fa === 'no-ded')      return est === 'no-ded';
  if (fa === 'pendiente')   return est === 'pendiente';
  if (fa === 'egreso')      return est === 'egreso';
  if (fa === 'complemento') return est === 'complemento';
  return true;
});

/* ── Inicialización DataTable ─────────────────────────────────── */
$(document).ready(function() {
  dtable = $('#tbl').DataTable({
    pageLength : 100,
    lengthMenu : [25, 50, 100, 200, 500],
    deferRender: true,
    scrollX    : true,
    autoWidth  : true,
    order      : [[3, 'asc']],
    stateSave  : true,
    stateDuration: 0,
    language   : {
      search     : 'Buscar:',
      lengthMenu : 'Mostrar _MENU_',
      info       : '_START_ a _END_ de _TOTAL_',
      paginate   : { previous: 'Ant', next: 'Sig' },
      zeroRecords: 'Sin resultados'
    },
    columnDefs: [
      { targets: [0],        width: '25px'  },
      { targets: [1, 2],     width: '160px' },
      { targets: [3],        width: '100px' },
      { targets: [4],        width: '160px' },
      { targets: [5, 6, 7],  width: '80px'  },
      { targets: [8],        width: '90px'  },
      { targets: [9, 10, 11],width: '140px' },
      { targets: [12],       width: '140px' },
      { targets: [13],       width: '220px' }
    ]
  });

  $('#bq-custom').on('keyup', function() { dtable.search(this.value).draw(); });

  // Inicializar fórmulas visibles en todas las filas al cargar
  dtable.rows().every(function() {
    const tr  = this.node();
    const ins = tr ? tr.querySelectorAll('.ip') : [];
    if (ins.length >= 3) {
      recalcularTotal(tr,
        parseFloat(ins[0].value)||0,
        parseFloat(ins[1].value)||0,
        parseFloat(ins[2].value)||0
      );
    }
  });

  restaurarCambios();
  actualizarContadores();
});

/* ── toggleDiot(): mostrar/ocultar DIOT ──────────────────────── */
function toggleDiot(card) {
  const sec = document.getElementById('diot-section');
  if (!sec) return;
  const visible = sec.style.display !== 'none';
  sec.style.display = visible ? 'none' : 'block';
  card.classList.toggle('diot-open', !visible);
  if (!visible) actualizarDiot();
}

/* ── actualizarContadores(): recuenta data-est en TODAS las filas */
function actualizarContadores() {
  let cnt16=0, cnt0=0, cnt160=0, e16=0, e0=0, e160=0, nod=0, pen=0, eg=0, cp=0, tom=0, tma=0, tot=0;

  // ── Acumuladores separados por componente ──────────────────────────────────
  // SUB16%, IVA16%, SUB0% se guardan por separado (no mezclados)
  // Prefijo d_ = deducible (ded*), e_ = efectivo (efe*)
  let dSub16=0, dIva16=0, dSub0=0;   // transferencia/tarjeta deducible
  let eSub16=0, eIva16=0, eSub0=0;   // efectivo deducible
  let mNod=0, mPen=0;                 // no deducibles y pendientes (sobre TOTAL)

  $(dtable.rows().nodes()).each(function() {
    const est  = this.getAttribute('data-est') || '';
    const s2   = parseFloat(String($(this).find('.ip').eq(0).val()).replace(/[^0-9.-]/g, '')) || 0;
    const i16  = parseFloat(String($(this).find('.ip').eq(1).val()).replace(/[^0-9.-]/g, '')) || 0;
    const s0   = parseFloat(String($(this).find('.ip').eq(2).val()).replace(/[^0-9.-]/g, '')) || 0;
    const tds  = $(this).find('td:eq(8)').text() || '0';
    const tots = parseFloat(tds.replace(/[^0-9.-]/g,'')) || 0;
    tot++;

    if      (est === 'ded16' || est === 'ded') { cnt16++;  dSub16 += s2; dIva16 += i16; }
    else if (est === 'ded0')                   { cnt0++;   dSub0  += s0; }
    else if (est === 'ded160')                 { cnt160++; dSub16 += s2; dIva16 += i16; dSub0 += s0; }
    else if (est === 'efe16')                  { e16++;    eSub16 += s2; eIva16 += i16; }
    else if (est === 'efe0')                   { e0++;     eSub0  += s0; }
    else if (est === 'efe160')                 { e160++;   eSub16 += s2; eIva16 += i16; eSub0 += s0; }
    else if (est === 'no-ded')                 { nod++;    mNod   += tots; }
    else if (est === 'pendiente')              { pen++;    mPen   += tots; }
    else if (est === 'egreso')                 { eg++; }
    else if (est === 'complemento')            { cp++; }
    else if (est === 'otro-mes')               { tom++; }
    else if (est === 'mes-ant')                { tma++; }
  });

  const totalSub16 = dSub16 + eSub16;          // SUB16% total (ded + efe)
  const totalIva16 = dIva16 + eIva16;          // IVA acreditable total (ded + efe)
  const totalSub0  = dSub0  + eSub0;           // SUB0% total (ded + efe)
  const granTotal  = totalSub16 + totalIva16 + totalSub0; // TOTAL deducible fiscal

  const fmt = v => '$'+v.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g,',');
  const el  = id => document.getElementById(id);

  const bf = document.querySelectorAll('.fls .bf');
  if (bf[0]) bf[0].textContent = 'TODOS (' +tot+ ')';
  if (bf[1]) bf[1].textContent = 'DED 16% (' +cnt16+ ')';
  if (bf[2]) bf[2].textContent = 'DED 0% (' +cnt0+ ')';
  if (bf[3]) bf[3].textContent = '16 Y 0% (' +cnt160+ ')';
  if (bf[4]) bf[4].textContent = 'EFE (' +(e16+e0+e160)+ ')';
  if (bf[5]) bf[5].textContent = 'NO DED (' +nod+ ')';
  if (bf[6]) bf[6].textContent = 'PEND ('  +pen+ ')';
  if (bf[7]) bf[7].textContent = 'EGRESO ('+eg+  ')';
  if (bf[8]) bf[8].textContent = 'CP01 ('  +cp+  ')';

  if (el('cnt-tot16'))      el('cnt-tot16').textContent      = (cnt16 + cnt160 + e16 + e160);
  if (el('mnt-tot16'))      el('mnt-tot16').textContent      = fmt(totalSub16);

  if (el('cnt-tot0'))       el('cnt-tot0').textContent       = (cnt0 + cnt160 + e0 + e160);
  if (el('mnt-tot0'))       el('mnt-tot0').textContent       = fmt(totalSub0);

  if (el('mnt-iva-acred'))  el('mnt-iva-acred').textContent  = fmt(totalIva16);
  if (el('mnt-gran-total')) el('mnt-gran-total').textContent = fmt(granTotal);

  if (el('cnt-efe'))        el('cnt-efe').textContent        = (e16 + e0 + e160);
  if (el('mnt-efe'))        el('mnt-efe').textContent        = fmt(eSub16 + eIva16 + eSub0);

  if (el('cnt-nod'))        el('cnt-nod').textContent        = nod;
  if (el('mnt-nod'))        el('mnt-nod').textContent        = fmt(mNod);

  if (el('cnt-pen'))        el('cnt-pen').textContent        = pen;
  if (el('mnt-pen'))        el('mnt-pen').textContent        = fmt(mPen);

  if (el('cnt-tot'))        el('cnt-tot').textContent        = tot;

  if (el('mnt-efe16'))      el('mnt-efe16').textContent      = fmt(eSub16);
  if (el('mnt-efe16-iva'))  el('mnt-efe16-iva').textContent  = fmt(eIva16);
  if (el('mnt-efe0'))       el('mnt-efe0').textContent       = fmt(eSub0);
  if (el('mnt-ded16'))      el('mnt-ded16').textContent      = fmt(dSub16);
  if (el('mnt-ded16-iva'))  el('mnt-ded16-iva').textContent  = fmt(dIva16);
  if (el('mnt-ded0'))       el('mnt-ded0').textContent       = fmt(dSub0);

  actualizarDiot();
}

/* ── ft(): activar filtro por botón ──────────────────────────── */
function ft(t, btn) {
  fa = t;
  document.querySelectorAll('.fls .bf').forEach(b => b.classList.remove('ac'));
  if (btn) btn.classList.add('ac');
  dtable.draw();
}

/* ── ce(): color + data-est + localStorage ──────────────────────
   FIX PRINCIPAL: función única con parámetro guardar=true.
   La segunda definición anterior (sin guardarCambio) causaba que
   los cambios del usuario no se persistieran en localStorage.
   ─────────────────────────────────────────────────────────────── */
function ce(sel, guardar = true) {
  const v = sel.value;
  const o = sel.dataset.original;
  sel.className = 'se';

  if      (v.includes('COMPLE'))                        sel.classList.add('complemento');
  else if (v.includes('PEND'))                          sel.classList.add('pendiente');
  else if (v.includes('OTRO MES'))                      sel.classList.add('otro-mes');
  else if (v.includes('MES ANTER'))                     sel.classList.add('mes-ant');
  else if (v.includes('NO DED') || v.includes('ERROR')) sel.classList.add('no-ded');
  else if (v.includes('EGRESO'))                        sel.classList.add('egreso');
  else if (v.includes('16 Y 0'))                        sel.classList.add('mix');
  else if (v.includes('0%') && !v.includes('16'))       sel.classList.add('ded0');
  else if (v.includes('16%'))                           sel.classList.add('ded16');
  else                                                  sel.classList.add('no-ded');

  if (v !== o) sel.classList.add('changed');
  else         sel.classList.remove('changed');

  const tr = sel.closest('tr');
  if (tr) {
    let est = 'no-ded';
    if      (v.includes('COMPLE'))                        est = 'complemento';
    else if (v.includes('PEND'))                          est = 'pendiente';
    else if (v.includes('OTRO MES'))                      est = 'otro-mes';
    else if (v.includes('MES ANTER'))                     est = 'mes-ant';
    else if (v.includes('NO DED') || v.includes('ERROR')) est = 'no-ded';
    else if (v.includes('EGRESO'))                        est = 'egreso';
    else if (v.includes('EFE') && v.includes('0%'))       est = 'efe0';
    else if (v.includes('EFE'))                           est = 'efe16';
    else if (v.includes('16 Y 0'))                        est = 'ded160';
    else if (v.includes('0%'))                            est = 'ded0';
    else if (v.includes('DED'))                           est = 'ded16';
    tr.setAttribute('data-est', est);

    if (guardar) {
      const uuid = (tr.cells[1] ? tr.cells[1].getAttribute('title') || '' : '').trim().toUpperCase();
      if (uuid) {
        if (v !== o) guardarCambio(uuid, v);
        else         borrarCambio(uuid);
      }
    }
  }
  actualizarContadores();
}

/* ── limpiarCambios(): resetea al estado original ─────────────── */
function limpiarCambios() {
  if (!confirm('¿Borrar todos los cambios manuales y volver al estado original?')) return;
  try { localStorage.removeItem(STORE_KEY); } catch(e) {}
  location.reload();
}

/* ── beforeunload: guardar todas las filas al cerrar/recargar ─── */
window.addEventListener('beforeunload', function() {
  try {
    $(dtable.rows().nodes()).each(function() {
      guardarFila(this);
    });
  } catch(e) {}
});

/* ── recalcularTotal(): SUB2 + IVA16 + SUB0 = TOTAL ───────────── */
function recalcularTotal(tr, s2, i16, s0) {
  const total  = Math.round((s2 + i16 + s0) * 100) / 100;
  const totVal = tr.querySelector('.tot-val');
  const totFrm = tr.querySelector('.tot-formula');
  if (totVal) {
    totVal.textContent = '$' + total.toLocaleString('es-MX', {
      minimumFractionDigits: 2, maximumFractionDigits: 2
    });
  }
  if (totFrm) {
    const fmt = v => v % 1 === 0 ? v.toFixed(0) : v.toFixed(2);
    totFrm.textContent = (s2||i16||s0) ? fmt(s2)+' + '+fmt(i16)+' + '+fmt(s0) : '';
  }
  return total;
}

/* ── rc(): recalcular estatus al editar montos ────────────────── */
function rc(inp) {
  const tr = inp.closest('tr');
  if (!tr) return;
  const ins = tr.querySelectorAll('.ip');
  if (ins.length < 3) return;
  const s2  = parseFloat(ins[0].value) || 0;
  const i16 = parseFloat(ins[1].value) || 0;
  const s0  = parseFloat(ins[2].value) || 0;
  const tot = recalcularTotal(tr, s2, i16, s0);
  guardarFila(tr);
  const sel = tr.querySelector('.se');
  if (!sel) return;
  const forma = (tr.cells[9]?.innerText  || '').trim();
  const uso   = (tr.cells[10]?.innerText || '').trim();
  const met   = (tr.cells[8]?.innerText  || '').trim();
  const fp = forma.substring(0,2);
  const uc = uso.substring(0,3).toUpperCase();
  const mc = met.substring(0,3).toUpperCase();
  if (uc === 'S01') { sel.value = 'NO DEDUCIBLE'; ce(sel); return; }
  if (!['G01','G02','G03'].includes(uc) ||
      !['PUE','PPD'].includes(mc)       ||
      !['01','02','03','04','28'].includes(fp)) return;
  if (fp==='01' && tot>=2000) { sel.value='NO DEDUCIBLE: Efectivo >= $2,000'; ce(sel); return; }
  if (uc==='G02') { sel.value='EGRESO'; ce(sel); return; }
  let suf = 'NO DEDUCIBLE';
  if      (s2>0 && i16>0 && s0===0) suf='16%';
  else if (s2>0 && i16>0 && s0>0)  suf='16 Y 0%';
  else if (s2===0 && i16===0 && s0>0) suf='0%';
  sel.value = (fp==='01' ? 'EFE ' : 'DED ') + suf;
  ce(sel);
}

/* ── Extras Módulos (Nómina y Depreciación) ───────────────────── */
function addExtRow(tid) {
  const tbl = document.getElementById(tid);
  if (!tbl) return;
  const tbody = tbl.querySelector('tbody');
  const row = tbody.rows[0].cloneNode(true);
  row.querySelectorAll('.ext-ip').forEach(inp => inp.value = '');
  tbody.appendChild(row);
  bindExtEvents();
  saveExt();
}

function delExtRow(btn, tid) {
  const tbody = document.getElementById(tid).querySelector('tbody');
  if (tbody.rows.length > 1) {
    btn.closest('tr').remove();
    saveExt();
  } else {
    btn.closest('tr').querySelectorAll('.ext-ip').forEach(inp => inp.value = '');
    saveExt();
  }
}

function bindExtEvents() {
  document.querySelectorAll('.ext-ip').forEach(inp => {
    inp.removeEventListener('input', saveExt);
    inp.addEventListener('input', saveExt);
  });
}

function calcExtTotals() {
  const fmt = v => '$'+v.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g,',');
  let nSueldo=0, nIsr=0, nImss=0, nTot=0;
  document.querySelectorAll('#tbl-nom tbody tr').forEach(tr => {
    const ips = tr.querySelectorAll('.num-nom');
    if(ips.length >= 4) {
      nSueldo += parseFloat(ips[0].value.replace(/[^0-9.-]/g,'')) || 0;
      nIsr    += parseFloat(ips[1].value.replace(/[^0-9.-]/g,'')) || 0;
      nImss   += parseFloat(ips[2].value.replace(/[^0-9.-]/g,'')) || 0;
      nTot    += parseFloat(ips[3].value.replace(/[^0-9.-]/g,'')) || 0;
    }
  });
  if(document.getElementById('tot-nom-sueldo')) document.getElementById('tot-nom-sueldo').textContent = fmt(nSueldo);
  if(document.getElementById('tot-nom-isr')) document.getElementById('tot-nom-isr').textContent = fmt(nIsr);
  if(document.getElementById('tot-nom-imss')) document.getElementById('tot-nom-imss').textContent = fmt(nImss);
  if(document.getElementById('tot-nom-total')) document.getElementById('tot-nom-total').textContent = fmt(nTot);

  let dTot=0;
  document.querySelectorAll('#tbl-dep tbody tr').forEach(tr => {
    const ips = tr.querySelectorAll('.num-dep');
    if(ips.length >= 1) {
      dTot  += parseFloat(ips[0].value.replace(/[^0-9.-]/g,'')) || 0;
    }
  });
  if(document.getElementById('tot-dep-total')) document.getElementById('tot-dep-total').textContent = fmt(dTot);
}

let pasteHistory = [];
document.addEventListener('keydown', function(e) {
  if (e.ctrlKey && e.key === 'z' && pasteHistory.length > 0) {
    const isIp = e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA';
    if (!isIp) {
      e.preventDefault();
      const lastState = pasteHistory.pop();
      lastState.forEach(item => {
        if (item.el) {
          item.el.value = item.oldVal;
          item.el.dispatchEvent(new Event('input', { bubbles: true }));
          item.el.dispatchEvent(new Event('change', { bubbles: true }));
        }
      });
    }
  }
});

document.addEventListener('paste', function(e) {
  const target = e.target;
  const isExtTbl = target.closest('.ext-tbl');
  const isMainTbl = target.closest('#tbl');
  if (!isExtTbl && !isMainTbl) return;
  if (target.tagName !== 'INPUT' && target.tagName !== 'SELECT') return;

  const clipboardData = e.clipboardData || window.clipboardData;
  const pastedData = clipboardData.getData('Text');
  if (!pastedData) return;

  const rows = pastedData.replace(/\r/g, '').split('\n').map(r => r.split('\t'));
  if (rows.length <= 1 && rows[0].length <= 1) return;

  e.preventDefault();
  const tbody = target.closest('tbody');
  const tr = target.closest('tr');
  const td = target.closest('td');
  if (!tbody || !tr || !td) return;

  const startRowIdx = Array.from(tbody.children).indexOf(tr);
  const startColIdx = Array.from(tr.children).indexOf(td);
  
  let tblType = 'main';
  if (tbody.id === 'tbody-nom' || (isExtTbl && isExtTbl.id === 'tbl-nom')) tblType = 'nom';
  else if (tbody.id === 'tbody-dep' || (isExtTbl && isExtTbl.id === 'tbl-dep')) tblType = 'dep';

  const historyItem = [];

  rows.forEach((rowCells, rOffset) => {
    if (rowCells.length === 1 && rowCells[0].trim() === '') return;
    
    if (tblType === 'nom' && rowCells.length >= 11 && rowCells[7].trim() === '' && rowCells[8].trim() === '') {
        rowCells.splice(7, 2);
    }
    
    let targetRow = tbody.children[startRowIdx + rOffset];
    if (!targetRow) {
      if (tblType === 'main') return;
      const clone = tbody.rows[0].cloneNode(true);
      clone.querySelectorAll('input, select').forEach(inp => inp.value = '');
      clone.querySelectorAll('.ext-sum, .ext-tot').forEach(inp => inp.textContent = '$0.00');
      tbody.appendChild(clone);
      targetRow = clone;
    }

    rowCells.forEach((cellData, cOffset) => {
      const targetTd = targetRow.children[startColIdx + cOffset];
      if (!targetTd || targetTd.classList.contains('ext-sum') || targetTd.classList.contains('ext-tot') || targetTd.classList.contains('ext-act')) return;
      let input = targetTd.querySelector('input, select');
      if (input) {
        historyItem.push({ el: input, oldVal: input.value });
        let val = cellData.trim();
        if (input.tagName === 'INPUT') {
            val = val.replace(/[^0-9A-Za-z. ,\/ñÑáéíóúÁÉÍÓÚ-]/g, '');
        }
        input.value = val;
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
      }
    });
  });
  if (historyItem.length > 0) pasteHistory.push(historyItem);
});


function saveExt() {
  calcExtTotals();
  const data = { nom: [], dep: [] };
  const tblNom = document.getElementById('tbl-nom');
  const tblDep = document.getElementById('tbl-dep');
  
  if (tblNom) {
    const rows = tblNom.querySelector('tbody').rows;
    for(let i=0; i<rows.length; i++) {
      const vals = Array.from(rows[i].querySelectorAll('.ext-ip')).map(x => x.value);
      if(vals.join('').trim() !== '') data.nom.push(vals);
    }
  }
  if (tblDep) {
    const rows = tblDep.querySelector('tbody').rows;
    for(let i=0; i<rows.length; i++) {
      const vals = Array.from(rows[i].querySelectorAll('.ext-ip')).map(x => x.value);
      if(vals.join('').trim() !== '') data.dep.push(vals);
    }
  }
  try { localStorage.setItem(STORE_KEY + '_ext', JSON.stringify(data)); } catch(e){}
}

function loadExt() {
  try {
    const data = JSON.parse(localStorage.getItem(STORE_KEY + '_ext'));
    if(!data) { calcExtTotals(); return; }
    
    const tblNom = document.getElementById('tbl-nom');
    if(data.nom && data.nom.length > 0 && tblNom) {
      const tbody = tblNom.querySelector('tbody');
      while(tbody.rows.length < data.nom.length) {
        tbody.appendChild(tbody.rows[0].cloneNode(true));
      }
      data.nom.forEach((vals, i) => {
        const inps = tbody.rows[i].querySelectorAll('.ext-ip');
        vals.forEach((v, j) => { if(inps[j]) inps[j].value = v; });
      });
    }
    const tblDep = document.getElementById('tbl-dep');
    if(data.dep && data.dep.length > 0 && tblDep) {
      const tbody = tblDep.querySelector('tbody');
      while(tbody.rows.length < data.dep.length) {
        tbody.appendChild(tbody.rows[0].cloneNode(true));
      }
      data.dep.forEach((vals, i) => {
        const inps = tbody.rows[i].querySelectorAll('.ext-ip');
        vals.forEach((v, j) => { if(inps[j]) inps[j].value = v; });
      });
    }
  } catch(e) {}
  bindExtEvents();
  calcExtTotals();
}

/* ── toggleDesglose(): mostrar/ocultar desglose EFE/DED ─────────── */
function toggleDesglose() {
  const wrap  = document.getElementById('desglose-wrap');
  const arrow = document.getElementById('desglose-arrow');
  if (!wrap) return;
  const open = wrap.classList.toggle('open');
  if (arrow) arrow.style.transform = open ? 'rotate(180deg)' : 'rotate(0deg)';
}
