const { useState, useEffect, useCallback } = React;

/* ============ UTILIDADES PARSER ============ */
const MATERIALS = ['PLA','PETG','ABS','ASA','TPU','TPE','Nylon','PA6','PA12','PC','HIPS','PVA','BVOH','PET','PP','PEEK','PEI'];
const COLORS = ['Black','White','Red','Blue','Green','Yellow','Orange','Purple','Pink','Grey','Gray','Silver','Gold','Transparent','Natural','Negro','Blanco','Rojo','Azul','Verde','Amarillo','Naranja','Morado','Rosa','Gris','Plateado','Dorado','Transparente','Natural'];

function parseNFCData(rawText, tagInfo='') {
  if (!rawText || typeof rawText !== 'string') return createEmptyRecord(tagInfo);
  const text = rawText.trim();
  const record = createEmptyRecord(tagInfo);
  record.rawData = text;

  for (const mat of MATERIALS) {
    if (new RegExp(`\\b${mat}\\b`, 'i').test(text)) { record.material = mat.toUpperCase(); break; }
  }
  for (const color of COLORS) {
    if (new RegExp(`\\b${color}\\b`, 'i').test(text)) { record.color = color.charAt(0).toUpperCase() + color.slice(1).toLowerCase(); break; }
  }
  const dm = text.match(/(\d+[.,]?\d*)\s*mm/i);
  if (dm) { const v = parseFloat(dm[1].replace(',','.')); if (v>=1 && v<=3.5) record.diameter = v.toFixed(2)+' mm'; }
  const wm = text.match(/(\d+[.,]?\d*)\s*(kg|g|gr)\b/i);
  if (wm) { const v=parseFloat(wm[1].replace(',','.')); const u=wm[2].toLowerCase(); record.weight = u==='kg' ? (v>=0.1&&v<=5 ? v+' kg' : '') : (v>=100&&v<=5000 ? (v/1000)+' kg' : ''); }
  const nm = text.match(/(?:nozzle|extrusion|extruder|hotend|temp\.?\s*nozzle|temp\.?\s*extrusion)[\s:]*(\d+)[°\s]*C?/i) || text.match(/(\d+)\s*°?\s*C?\s*(?:nozzle|extrusion|extruder)/i) || text.match(/(?:print|printing)\s*temp[\s:]*(\d+)/i);
  if (nm) { const t=parseInt(nm[1]); if(t>=150&&t<=450) record.nozzleTemp = t+' °C'; }
  const bm = text.match(/(?:bed|cama|plataforma|platform)[\s:]*(\d+)[°\s]*C?/i) || text.match(/(\d+)\s*°?\s*C?\s*(?:bed|cama)/i);
  if (bm) { const t=parseInt(bm[1]); if(t>=0&&t<=150) record.bedTemp = t+' °C'; }
  const bps = [/(?:brand|marca|manufacturer|fabricante)[\s:]*([^\n,;]+)/i, /\b(Bambu\s*Lab|Prusa|Anycubic|Creality|Ender|Esun|PolyTerra|PolyLite|Sunlu|Overture|Hatchbox|Prusament|Fiberlogy|Sakata|3D\s*Jake|ColorFabb|Fillamentum|ProtoPasta)\b/i];
  for (const p of bps) { const m=text.match(p); if(m){ record.brand=m[1].trim(); break; } }
  const hm = text.match(/#?([0-9A-Fa-f]{6})\b/);
  if (hm && !record.color) record.color = '#'+hm[1].toUpperCase();
  if (!record.material && !record.color && !record.brand) record.notes = text.substring(0,200);
  record.readDate = new Date().toLocaleString('es-ES');
  return record;
}
function createEmptyRecord(tagInfo='') {
  return { material:'', color:'', diameter:'', weight:'', brand:'', nozzleTemp:'', bedTemp:'', notes:'', rawData:'', readDate:'', tagInfo };
}
const HEADERS = ['Fecha Lectura','Marca','Material','Color','Diámetro','Peso','Temp. Nozzle','Temp. Cama','Notas','Datos Crudos','Info Tag'];
function recordToArray(r){ return [r.readDate,r.brand,r.material,r.color,r.diameter,r.weight,r.nozzleTemp,r.bedTemp,r.notes,r.rawData,r.tagInfo]; }

function bufToHex(buffer) {
  return Array.from(new Uint8Array(buffer)).map(b => b.toString(16).padStart(2,'0')).join(' ').toUpperCase();
}
function tryDecodeText(buffer) {
  try { return new TextDecoder('utf-8').decode(buffer); } catch { return null; }
}

/* ============ EXCEL ============ */
function exportToExcel(records) {
  if (!records.length) { alert('No hay registros para exportar'); return; }
  const ws = XLSX.utils.aoa_to_sheet([HEADERS, ...records.map(recordToArray)]);
  ws['!cols'] = HEADERS.map(h => ({ wch: Math.max(h.length+2, 12) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Filamentos');
  XLSX.writeFile(wb, 'filamentos_nfc.xlsx');
}

/* ============ STORAGE HOOK ============ */
function useStorage() {
  const [records, setRecords] = useState(() => { try { const s=localStorage.getItem('nfc_filamento_records'); return s?JSON.parse(s):[]; }catch{return [];} });
  useEffect(() => { localStorage.setItem('nfc_filamento_records', JSON.stringify(records)); }, [records]);
  const addRecord = useCallback(r => setRecords(p=>[r,...p]), []);
  const updateRecord = useCallback((i,r) => setRecords(p=>{ const n=[...p]; n[i]=r; return n; }), []);
  const deleteRecord = useCallback(i => setRecords(p=>p.filter((_,idx)=>idx!==i)), []);
  const clearAll = useCallback(() => { if(confirm('¿Borrar TODOS los registros?')) setRecords([]); }, []);
  return { records, addRecord, updateRecord, deleteRecord, clearAll };
}

/* ============ NFC HOOK ============ */
function useNFC() {
  const [scanning, setScanning] = useState(false);
  const [error, setError] = useState(null);
  const [supported] = useState('NDEFReader' in window);
  const ctrlRef = React.useRef(null);

  const startScan = useCallback(async (onRecord) => {
    if (!supported) { setError('NFC no soportado. Usa Chrome en Android con NFC activado.'); return; }
    try {
      setError(null); setScanning(true);
      const ndef = new NDEFReader();
      ctrlRef.current = new AbortController();
      await ndef.scan({ signal: ctrlRef.current.signal });

      ndef.addEventListener('reading', e => {
        let txt='';
        let hexParts=[];
        let info = `Serial: ${e.serialNumber || 'N/A'}`;
        for (const rec of e.message.records) {
          info += ` | Tipo: ${rec.recordType}`;
          if (rec.recordType==='text') {
            txt += new TextDecoder(rec.encoding||'utf-8').decode(rec.data)+' ';
          } else if (rec.recordType==='url') {
            txt += new TextDecoder().decode(rec.data)+' ';
          } else if (rec.recordType==='mime') {
            info += ` (${rec.mediaType})`;
            const decoded = tryDecodeText(rec.data);
            if (decoded) txt += decoded + ' ';
            else hexParts.push(bufToHex(rec.data));
          } else if (rec.recordType==='unknown') {
            hexParts.push(bufToHex(rec.data));
          } else {
            const decoded = tryDecodeText(rec.data);
            if (decoded) txt += decoded + ' ';
            else hexParts.push(bufToHex(rec.data));
          }
        }
        if (hexParts.length && !txt.trim()) txt = hexParts.join(' | ');
        onRecord(txt.trim(), info);
      });

      ndef.addEventListener('readingerror', () => {
        setError('Error leyendo etiqueta. El tag no contiene datos NDEF válidos. Es posible que use un formato binario propietario (común en Anycubic, Sunlu...). Prueba el "Modo pegar datos" con NFC Tools.');
      });
    } catch(err) { if(err.name!=='AbortError') setError(err.message||'Error al escanear'); setScanning(false); }
  }, [supported]);

  const stopScan = useCallback(() => { if(ctrlRef.current){ ctrlRef.current.abort(); ctrlRef.current=null; } setScanning(false); }, []);
  return { scanning, error, supported, startScan, stopScan };
}

/* ============ COMPONENTES ============ */
function App() {
  const { scanning, error, supported, startScan, stopScan } = useNFC();
  const { records, addRecord, updateRecord, deleteRecord, clearAll } = useStorage();
  const [editing, setEditing] = useState(null);
  const [form, setForm] = useState({});
  const [showRaw, setShowRaw] = useState({});
  const [showManual, setShowManual] = useState(false);
  const [showPaste, setShowPaste] = useState(false);
  const [manualForm, setManualForm] = useState(createEmptyRecord());
  const [pasteText, setPasteText] = useState('');

  const onRead = useCallback((txt, tagInfo) => {
    const rec = parseNFCData(txt, tagInfo);
    addRecord(rec);
    stopScan();
  }, [addRecord, stopScan]);

  const beginEdit = i => { setEditing(i); setForm({...records[i]}); };
  const saveEdit = () => { if(editing!==null){ updateRecord(editing, form); setEditing(null); } };
  const cancelEdit = () => setEditing(null);
  const toggleRaw = i => setShowRaw(p=>({...p,[i]:!p[i]}));

  const saveManual = () => {
    const rec = { ...manualForm, readDate: new Date().toLocaleString('es-ES') };
    addRecord(rec);
    setManualForm(createEmptyRecord());
    setShowManual(false);
  };

  const savePaste = () => {
    if (!pasteText.trim()) return;
    const rec = parseNFCData(pasteText, 'Pegado manual desde NFC Tools');
    addRecord(rec);
    setPasteText('');
    setShowPaste(false);
  };

  const manualInput = (label, field, placeholder='') => (
    <div style={S.fieldRow}>
      <label style={S.label}>{label}</label>
      <input style={S.manualInp} placeholder={placeholder} value={manualForm[field]} onChange={e=>setManualForm({...manualForm,[field]:e.target.value})}/>
    </div>
  );

  return (
    <div style={S.container}>
      <header style={S.header}>
        <h1 style={S.title}>📡 NFC Filamento</h1>
        <p style={S.subtitle}>Lector de etiquetas NFC para filamentos 3D</p>
      </header>

      <div style={S.card}>
        {!supported && <div style={S.errBox}>⚠️ NFC no disponible. Usa <strong>Chrome en Android</strong> con NFC activado.</div>}
        {error && <div style={S.errBox}>❌ {error}</div>}
        <button onClick={scanning?stopScan:()=>startScan(onRead)} disabled={!supported} style={{...S.btn, background: scanning?'#ef4444':'#0ea5e9', opacity: supported?1:0.5}}>
          {scanning?'⏹️ Detener escaneo':'📲 Escanear etiqueta NFC'}
        </button>
        {scanning && <div style={S.scanning}><div style={S.pulse}></div><p>Acerca el móvil a la etiqueta NFC...</p></div>}

        <button onClick={()=>{setShowManual(!showManual);setShowPaste(false)}} style={{...S.btn, background:'#475569', marginTop:'10px'}}>
          ✏️ {showManual?'Ocultar entrada manual':'Añadir entrada manual'}
        </button>
        <button onClick={()=>{setShowPaste(!showPaste);setShowManual(false)}} style={{...S.btn, background:'#7c3aed', marginTop:'10px'}}>
          📋 {showPaste?'Ocultar modo pegar datos':'Pegar datos NFC (NFC Tools)'}
        </button>

        {showManual && (
          <div style={{marginTop:'16px', padding:'16px', background:'#0f172a', borderRadius:'10px'}}>
            <h3 style={{margin:'0 0 12px', color:'#94a3b8', fontSize:'14px'}}>Introduce los datos del filamento:</h3>
            {manualInput('Marca','brand','Ej: Anycubic')}
            {manualInput('Material','material','Ej: PLA')}
            {manualInput('Color','color','Ej: Rojo')}
            {manualInput('Diámetro','diameter','Ej: 1.75 mm')}
            {manualInput('Peso','weight','Ej: 1 kg')}
            {manualInput('Temp. Nozzle','nozzleTemp','Ej: 200 °C')}
            {manualInput('Temp. Cama','bedTemp','Ej: 60 °C')}
            {manualInput('Notas','notes','Cualquier dato extra')}
            <button onClick={saveManual} style={{...S.btn, background:'#10b981', marginTop:'10px'}}>💾 Guardar registro</button>
          </div>
        )}

        {showPaste && (
          <div style={{marginTop:'16px', padding:'16px', background:'#0f172a', borderRadius:'10px'}}>
            <h3 style={{margin:'0 0 8px', color:'#94a3b8', fontSize:'14px'}}>Pega aquí los datos hexadecimales de NFC Tools:</h3>
            <p style={{margin:'0 0 10px', fontSize:'12px', color:'#64748b'}}>Abre NFC Tools → "Other" → "Advanced NFC commands" → lee la etiqueta → copia el hex → pégalo aquí</p>
            <textarea style={{...S.manualInp, width:'100%', minHeight:'80px', fontFamily:'monospace', fontSize:'12px'}} value={pasteText} onChange={e=>setPasteText(e.target.value)} placeholder="04 A5 2B..."/>
            <button onClick={savePaste} style={{...S.btn, background:'#10b981', marginTop:'10px'}}>💾 Guardar datos pegados</button>
          </div>
        )}
      </div>

      {records.length>0 && (
        <div style={S.card}>
          <div style={S.secHead}>
            <h2 style={S.secTitle}>📋 Registros ({records.length})</h2>
            <div style={S.actions}>
              <button onClick={()=>exportToExcel(records)} style={S.expBtn}>📥 Exportar Excel</button>
              <button onClick={clearAll} style={S.delBtn}>🗑️ Borrar todo</button>
            </div>
          </div>
          <div style={S.tableWrap}>
            <table style={S.table}>
              <thead><tr style={S.thead}>
                <th style={S.th}>Fecha</th><th style={S.th}>Marca</th><th style={S.th}>Material</th><th style={S.th}>Color</th>
                <th style={S.th}>Diámetro</th><th style={S.th}>Peso</th><th style={S.th}>Nozzle</th><th style={S.th}>Cama</th><th style={S.th}>Acciones</th>
              </tr></thead>
              <tbody>
                {records.map((r,i)=> (
                  <React.Fragment key={i}>
                    <tr style={S.tr}>
                      {editing===i ? (<>
                        <td style={S.td}>{r.readDate}</td>
                        {['brand','material','color','diameter','weight','nozzleTemp','bedTemp'].map(f => (
                          <td style={S.td} key={f}><input style={S.inp} value={form[f]} onChange={e=>setForm({...form,[f]:e.target.value})}/></td>
                        ))}
                        <td style={S.td}>
                          <button style={S.icon} onClick={saveEdit}>💾</button>
                          <button style={S.icon} onClick={cancelEdit}>❌</button>
                        </td>
                      </>) : (<>
                        <td style={S.td}>{r.readDate}</td>
                        <td style={S.td}>{r.brand}</td>
                        <td style={S.td}><strong>{r.material}</strong></td>
                        <td style={S.td}>{r.color}</td>
                        <td style={S.td}>{r.diameter}</td>
                        <td style={S.td}>{r.weight}</td>
                        <td style={S.td}>{r.nozzleTemp}</td>
                        <td style={S.td}>{r.bedTemp}</td>
                        <td style={S.td}>
                          <button style={S.icon} onClick={()=>beginEdit(i)}>✏️</button>
                          <button style={S.icon} onClick={()=>deleteRecord(i)}>🗑️</button>
                          {(r.rawData || r.tagInfo) && <button style={S.icon} onClick={()=>toggleRaw(i)}>{showRaw[i]?'🔼':'🔽'}</button>}
                        </td>
                      </>)}
                    </tr>
                    {showRaw[i] && (
                      <tr><td colSpan={9} style={S.rawCell}>
                        <div style={S.rawBox}>
                          {r.tagInfo && <div><strong>Info Tag:</strong> {r.tagInfo}</div>}
                          <div><strong>Datos crudos:</strong></div>
                          <pre style={S.pre}>{r.rawData || '(sin datos)'}</pre>
                        </div>
                      </td></tr>
                    )}
                  </React.Fragment>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {records.length===0 && !scanning && (
        <div style={S.empty}>
          <p>📭 Aún no hay registros</p>
          <p style={S.emptySub}>Pulsa "Escanear etiqueta NFC", "Entrada manual" o "Pegar datos NFC"</p>
        </div>
      )}
    </div>
  );
}

/* ============ ESTILOS ============ */
const S = {
  container: { maxWidth:'1200px', margin:'0 auto', padding:'16px', background:'#0f172a', minHeight:'100vh' },
  header: { textAlign:'center', padding:'20px 0', borderBottom:'1px solid #334155', marginBottom:'20px' },
  title: { margin:0, fontSize:'28px', color:'#38bdf8', fontWeight:700 },
  subtitle: { margin:'8px 0 0 0', color:'#94a3b8', fontSize:'14px' },
  card: { background:'#1e293b', borderRadius:'12px', padding:'20px', marginBottom:'16px', boxShadow:'0 4px 6px -1px rgba(0,0,0,0.3)' },
  errBox: { background:'#7f1d1d', color:'#fecaca', padding:'12px', borderRadius:'8px', marginBottom:'12px', fontSize:'14px' },
  btn: { width:'100%', padding:'18px', fontSize:'18px', fontWeight:600, color:'white', border:'none', borderRadius:'10px', cursor:'pointer' },
  scanning: { textAlign:'center', marginTop:'16px', color:'#38bdf8' },
  pulse: { width:'60px', height:'60px', margin:'0 auto 10px', borderRadius:'50%', background:'#0ea5e9', animation:'pulse 1.5s infinite', opacity:0.7 },
  secHead: { display:'flex', justifyContent:'space-between', alignItems:'center', flexWrap:'wrap', gap:'10px', marginBottom:'16px' },
  secTitle: { margin:0, fontSize:'18px' },
  actions: { display:'flex', gap:'8px' },
  expBtn: { background:'#10b981', color:'white', border:'none', padding:'8px 14px', borderRadius:'6px', cursor:'pointer', fontSize:'13px', fontWeight:600 },
  delBtn: { background:'#ef4444', color:'white', border:'none', padding:'8px 14px', borderRadius:'6px', cursor:'pointer', fontSize:'13px', fontWeight:600 },
  tableWrap: { overflowX:'auto', WebkitOverflowScrolling:'touch' },
  table: { width:'100%', borderCollapse:'collapse', fontSize:'13px' },
  thead: { background:'#0f172a' },
  th: { padding:'10px 8px', textAlign:'left', color:'#94a3b8', fontWeight:600, borderBottom:'2px solid #334155', whiteSpace:'nowrap' },
  tr: { borderBottom:'1px solid #334155' },
  td: { padding:'8px', color:'#e2e8f0', verticalAlign:'middle' },
  inp: { width:'80px', padding:'4px 6px', borderRadius:'4px', border:'1px solid #475569', background:'#0f172a', color:'#e2e8f0', fontSize:'12px' },
  manualInp: { flex:1, padding:'8px 10px', borderRadius:'6px', border:'1px solid #475569', background:'#1e293b', color:'#e2e8f0', fontSize:'14px' },
  fieldRow: { display:'flex', alignItems:'center', gap:'10px', marginBottom:'8px' },
  label: { width:'110px', fontSize:'13px', color:'#94a3b8', fontWeight:500 },
  icon: { background:'none', border:'none', cursor:'pointer', fontSize:'16px', padding:'2px 4px', opacity:0.8 },
  rawCell: { padding:'0 8px 8px' },
  rawBox: { padding:'10px', background:'#0f172a', borderRadius:'6px', fontSize:'12px', color:'#94a3b8', wordBreak:'break-all' },
  pre: { margin:'4px 0 0 0', whiteSpace:'pre-wrap', fontFamily:'monospace', fontSize:'11px', color:'#cbd5e1' },
  empty: { textAlign:'center', padding:'40px 20px', color:'#64748b' },
  emptySub: { fontSize:'13px', marginTop:'8px' }
};

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<App />);

/* Service Worker */
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('public/sw.js').catch(()=>{});
}
