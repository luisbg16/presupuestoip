import React, { useState, useEffect, useMemo } from 'react';
import { createClient } from '@supabase/supabase-js';
import * as XLSX from 'xlsx';
import { 
  LayoutDashboard, Receipt, UploadCloud, History, Search,
  Camera, LogOut, FileSpreadsheet, UserCheck, ShieldCheck, Eye, ChevronDown, ChevronUp, Download, CheckCircle, AlertCircle, Check, List
} from 'lucide-react';

const supabase = createClient(import.meta.env.VITE_SUPABASE_URL, import.meta.env.VITE_SUPABASE_ANON_KEY);
const MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
const COLOR_IP_PRIMARY = "#005aba"; 
const COLOR_ACCENT = "#ffd100";
const HOY = new Date().toISOString().split('T')[0];

const ADMIN_EMAILS = ["cavendano@chorotega.hn", "test@admin.com", "mrodriguez@chorotega.h"];

function App() {
  const [session, setSession] = useState(null);
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [seccion, setSeccion] = useState('reportes');
  const [tabReporte, setTabReporte] = useState('mensual');
  const [lineas, setLineas] = useState([]);
  const [historial, setHistorial] = useState([]);
  const [verDetalle, setVerDetalle] = useState(false);
  const [archivoExcel, setArchivoExcel] = useState(null);
  const [loading, setLoading] = useState(false);
  const [busquedaTabla, setBusquedaTabla] = useState('');
  const [busquedaHistorial, setBusquedaHistorial] = useState('');
  const [lineaExpandida, setLineaExpandida] = useState(null);
  
  const [mesResumen, setMesResumen] = useState(MESES[new Date().getMonth()]);
  const [compra, setCompra] = useState({ lineaId: '', monto: '', desc: '', foto: null, fecha: HOY });

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => setSession(session));
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => setSession(session));
    return () => subscription.unsubscribe();
  }, []);

  useEffect(() => { if (session) obtenerDatos(); }, [session]);

  const obtenerDatos = async () => {
    const { data: p } = await supabase.from('presupuestos').select('*').order('linea_nombre', { ascending: true });
    const { data: c } = await supabase.from('compras').select('*, presupuestos(*)').order('fecha', { ascending: false });
    setLineas(p || []);
    setHistorial(c || []);
  };

  const registrarGasto = async () => {
    if (!compra.fecha || !compra.lineaId || !compra.monto || !compra.foto) return alert("🚫 Faltan datos obligatorios.");
    if (compra.fecha > HOY) return alert("🚫 No se pueden registrar gastos con fecha futura.");
    
    const montoTotalFactura = parseFloat(compra.monto);
    const lineaSel = lineas.find(l => l.id.toString() === compra.lineaId.toString());

    if (montoTotalFactura > lineaSel.monto_actual) {
      return alert(`🚫 GASTO RECHAZADO: Saldo insuficiente en ${lineaSel.mes}.\nDisponible: L${lineaSel.monto_actual.toLocaleString()}`);
    }

    setLoading(true);
    try {
      const fileName = `${Date.now()}.${compra.foto.name.split('.').pop()}`;
      await supabase.storage.from('facturas').upload(fileName, compra.foto);
      const nuevoSaldo = lineaSel.monto_actual - montoTotalFactura;

      await supabase.from('presupuestos').update({ monto_actual: nuevoSaldo }).eq('id', lineaSel.id);
      await supabase.from('compras').insert([{ 
        presupuesto_id: lineaSel.id, 
        monto_lps: montoTotalFactura,
        monto_total_factura: montoTotalFactura, 
        descripcion: compra.desc, 
        fecha: compra.fecha, 
        url_factura: fileName, 
        creado_por: session.user.email, 
        aprobado: true,
        es_sobregiro: false 
      }]);

      alert("✅ Gasto registrado con éxito.");
      setCompra({ ...compra, monto: '', desc: '', foto: null, fecha: HOY, lineaId: '' });
      obtenerDatos();
    } catch (err) { alert(err.message); } finally { setLoading(false); }
  };

  const importarExcelIP = async () => {
    if (!archivoExcel) return alert("Selecciona un archivo.");
    setLoading(true);
    const reader = new FileReader();
    reader.readAsArrayBuffer(archivoExcel);
    reader.onload = async (e) => {
      try {
        await supabase.storage.from('facturas').upload('ultimo_presupuesto.xlsx', archivoExcel, { upsert: true });
        const data = new Uint8Array(e.target.result);
        const json = XLSX.utils.sheet_to_json(XLSX.read(data, { type: 'array' }).Sheets[XLSX.read(data, { type: 'array' }).SheetNames[0]], { defval: 0 });
        const mapaMeses = { "ene": "Ene", "feb": "Feb", "mar": "Mar", "abr": "Abr", "may": "May", "jun": "Jun", "jul": "Jul", "ago": "Ago", "sep": "Sep", "oct": "Oct", "nov": "Nov", "dic": "Dic" };
        const filas = [];
        json.forEach(fRaw => {
          const f = Object.keys(fRaw).reduce((acc, k) => { acc[k.toLowerCase().trim()] = fRaw[k]; return acc; }, {});
          if (f["línea"] || f["linea"]) {
            Object.keys(mapaMeses).forEach(m => {
              if (f[m] !== undefined) {
                const val = parseFloat(f[m].toString().replace(/[^\d.]/g, "")) || 0;
                filas.push({ linea_nombre: (f["línea"] || f["linea"]).trim(), responsable: (f["responsable"] || "").trim(), mes: mapaMeses[m], monto_inicial: val, monto_actual: val });
              }
            });
          }
        });
        await supabase.from('presupuestos').delete().neq('id', 0);
        await supabase.from('presupuestos').insert(filas);
        alert("✅ Presupuesto maestro actualizado."); obtenerDatos();
      } catch (err) { alert(err.message); } finally { setLoading(false); }
    };
  };

  const descargarReporteExcel = () => {
    const dataReporte = tabReporte === 'mensual' 
      ? lineas.filter(l => l.mes === mesResumen && l.monto_inicial > 0).map(l => ({
          'Línea': l.linea_nombre, 'Responsable': l.responsable, 'Presp. Inicial': l.monto_inicial, 'Gastado': l.monto_inicial - l.monto_actual, 'Disponible': l.monto_actual
        }))
      : lineasTablaFiltradas.map(l => ({
          'Línea': l.linea_nombre, 'Responsable': l.responsable, 'Presp. Anual': l.monto_inicial, 'Gastado Anual': l.monto_inicial - l.monto_actual, 'Disponible Anual': l.monto_actual
        }));
    const ws = XLSX.utils.json_to_sheet(dataReporte);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Reporte");
    XLSX.writeFile(wb, `Reporte_IP_${tabReporte}_${Date.now()}.xlsx`);
  };

  const historialFinal = useMemo(() => {
    return historial.filter(h => tabReporte === 'anual' || h.presupuestos?.mes === mesResumen)
      .filter(h => h.descripcion?.toLowerCase().includes(busquedaHistorial.toLowerCase()) || 
                   h.presupuestos?.linea_nombre?.toLowerCase().includes(busquedaHistorial.toLowerCase()));
  }, [historial, tabReporte, mesResumen, busquedaHistorial]);

  const stats = useMemo(() => {
    const filtradas = tabReporte === 'mensual' ? lineas.filter(l => l.mes === mesResumen) : lineas;
    const tP = filtradas.reduce((a, b) => a + (b.monto_inicial || 0), 0);
    const tD = filtradas.reduce((a, b) => a + (b.monto_actual || 0), 0);
    return { tP, tG: tP - tD, tD };
  }, [lineas, tabReporte, mesResumen]);

  const lineasTablaFiltradas = useMemo(() => {
    let base = tabReporte === 'mensual' ? lineas.filter(l => l.mes === mesResumen && l.monto_inicial > 0) : 
    [...new Set(lineas.map(l => l.linea_nombre))].map(n => {
        const g = lineas.filter(l => l.linea_nombre === n);
        return { id: n, linea_nombre: n, responsable: g[0]?.responsable, monto_inicial: g.reduce((a, b) => a + b.monto_inicial, 0), monto_actual: g.reduce((a, b) => a + b.monto_actual, 0) };
      }).filter(i => i.monto_inicial > 0);
    return base.filter(l => l.linea_nombre.toLowerCase().includes(busquedaTabla.toLowerCase()));
  }, [lineas, tabReporte, mesResumen, busquedaTabla]);

  if (!session) return (
    <div style={loginWrapper}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;800&display=swap'); body { font-family: 'Plus Jakarta Sans', sans-serif; }`}</style>
      <div style={loginCard}><h1 style={{color: COLOR_IP_PRIMARY, fontWeight: 800}}>IP CONTROL</h1><form onSubmit={async (e)=>{e.preventDefault(); const {error}=await supabase.auth.signInWithPassword({email, password}); if(error) alert("Error");}}><input type="email" placeholder="Usuario" style={inputStyle} onChange={e=>setEmail(e.target.value)} required /><input type="password" placeholder="Pass" style={inputStyle} onChange={e=>setPassword(e.target.value)} required /><button style={{...btnPro, background: COLOR_IP_PRIMARY, color:'white'}}>INGRESAR</button></form></div>
    </div>
  );

  return (
    <div style={appContainer}>
      <header style={{...headerStyle, background: COLOR_IP_PRIMARY}}>
        <div style={{display:'flex', alignItems:'center', gap:'12px'}}><ShieldCheck size={22} color={COLOR_ACCENT}/><span style={{fontWeight: 800}}>IP - CONTROL</span></div>
        <button onClick={()=>supabase.auth.signOut()} style={logoutBtn}><LogOut size={18}/></button>
      </header>

      <main style={mainStyle}>
        <nav style={navBar}>
          <button onClick={()=>setSeccion('compras')} style={seccion==='compras'?navBtnActive:navBtn}><Receipt size={24}/><span>Gasto</span></button>
          <button onClick={()=>setSeccion('reportes')} style={seccion==='reportes'?navBtnActive:navBtn}><LayoutDashboard size={24}/><span>Panel</span></button>
          <button onClick={()=>setSeccion('perfil')} style={seccion==='perfil'?navBtnActive:navBtn}><UserCheck size={24}/><span>Mi IP</span></button>
          <button onClick={()=>setSeccion('config')} style={seccion==='config'?navBtnActive:navBtn}><UploadCloud size={24}/><span>Excel</span></button>
        </nav>

        {seccion === 'compras' && (
          <div style={card}>
            <h3 style={cardTitle}><Receipt size={18}/> Registro de Factura</h3>
            <input type="date" style={inputStyle} value={compra.fecha} max={HOY} onChange={(e)=>setCompra({...compra, fecha:e.target.value})} />
            <select style={inputStyle} value={compra.lineaId} onChange={(e)=>setCompra({...compra, lineaId:e.target.value})}>
              <option value="">Línea...</option>
              {lineas.filter(l => l.mes === MESES[new Date(compra.fecha + 'T12:00:00').getMonth()] && l.monto_inicial > 0).map(l => (
                <option key={l.id} value={l.id}>{l.linea_nombre} (Saldo: L{l.monto_actual.toLocaleString()})</option>
              ))}
            </select>
            <input type="number" placeholder="Monto Lps" style={inputStyle} value={compra.monto} onChange={(e)=>setCompra({...compra, monto:e.target.value})} />
            <input type="text" placeholder="Proveedor / Detalle" style={inputStyle} value={compra.desc} onChange={(e)=>setCompra({...compra, desc:e.target.value})} />
            <label style={{...cameraBtn, background: compra.foto ? '#dcfce7' : '#f1f5f9', color: compra.foto ? '#16a34a' : '#475569'}}>
              {compra.foto ? <Check size={18}/> : <Camera size={18}/>} {compra.foto ? "FOTO LISTA" : "SELECCIONAR FOTO"} 
              <input type="file" hidden onChange={(e)=>setCompra({...compra, foto:e.target.files[0]})} />
            </label>
            <button onClick={registrarGasto} style={{...btnPro, background: COLOR_IP_PRIMARY, color:'white'}} disabled={loading}>{loading ? "..." : "REGISTRAR GASTO"}</button>
          </div>
        )}

        {seccion === 'reportes' && (
            <div style={{marginTop:'15px'}}>
                <div style={toggleContainer}><button onClick={()=>setTabReporte('mensual')} style={tabReporte==='mensual'?toggleActive:toggleInactive}>MENSUAL</button><button onClick={()=>setTabReporte('anual')} style={tabReporte==='anual'?toggleActive:toggleInactive}>ANUAL</button></div>
                
                <div style={{...dashboardGrid, marginTop:'15px'}}>
                    <div style={dashItem}><span style={dashLabel}>PRESUPUESTO</span><br/><b>L{stats.tP.toLocaleString()}</b></div>
                    <div style={{...dashItem, borderLeft:'1px solid #f1f5f9', borderRight:'1px solid #f1f5f9'}}><span style={dashLabel}>GASTADO</span><br/><b style={{color:'#dc2626'}}>L{stats.tG.toLocaleString()}</b></div>
                    <div style={dashItem}><span style={dashLabel}>DISPONIBLE</span><br/><b style={{color:COLOR_IP_PRIMARY}}>L{stats.tD.toLocaleString()}</b></div>
                </div>

                <div style={{...card, marginTop:'20px'}}>
                    <div style={{display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:'15px'}}>
                        <div style={{display:'flex', alignItems:'center', gap:'10px'}}>
                          <h3 style={cardTitle}><History size={18}/> HISTORIAL</h3>
                          <button onClick={descargarReporteExcel} style={{background:'none', border:'none', cursor:'pointer', color:'#94a3b8'}}><FileSpreadsheet size={18}/></button>
                        </div>
                        {tabReporte === 'mensual' && <select style={{...inputStyle, width:'auto', marginBottom:0, padding:'5px'}} value={mesResumen} onChange={(e)=>setMesResumen(e.target.value)}>{MESES.map(m=><option key={m} value={m}>{m}</option>)}</select>}
                    </div>
                    <div style={{display:'flex', alignItems:'center', gap:'10px', background:'#f8fafc', padding:'8px 12px', borderRadius:'10px', marginBottom:'15px'}}>
                        <Search size={14} color="#94a3b8"/><input type="text" placeholder="Buscar..." style={{border:'none', background:'none', fontSize:'12px', outline:'none', width:'100%'}} value={busquedaHistorial} onChange={(e)=>setBusquedaHistorial(e.target.value)}/>
                    </div>

                    <div style={{maxHeight:'400px', overflowY:'auto'}}>
                      {historialFinal.map(h => (
                          <div key={h.id} style={historyItem}>
                              <div style={{flex:1}}>
                                  <div style={{fontWeight:800, fontSize:'12px', color: COLOR_IP_PRIMARY}}>{h.presupuestos?.linea_nombre}</div>
                                  <div style={{fontSize:'11px', color:'#1e293b', fontWeight:600, marginTop:'2px'}}>{h.descripcion}</div>
                                  <div style={{fontSize:'10px', color:'#64748b', marginTop:'2px'}}>{h.fecha} • {h.creado_por?.split('@')[0]}</div>
                              </div>
                              <div style={{textAlign:'right', display:'flex', alignItems:'center', gap:'10px'}}>
                                  <b style={{fontSize:'13px', color: '#dc2626'}}>-L{h.monto_lps.toLocaleString()}</b>
                                  <button onClick={() => window.open(`${import.meta.env.VITE_SUPABASE_URL}/storage/v1/object/public/facturas/${h.url_factura}`, '_blank')} style={eyeBtn}><Eye size={16} color="#64748b"/></button>
                              </div>
                          </div>
                      ))}
                    </div>
                </div>
                
                <button onClick={()=>setVerDetalle(!verDetalle)} style={btnAccordion}>{verDetalle ? <ChevronUp size={16}/> : <ChevronDown size={16}/>} TABLA DETALLADA</button>
                {verDetalle && (
                  <div style={{...card, marginTop:'10px', padding:'15px'}}>
                    <div style={{position:'relative', marginBottom:'15px'}}><Search size={14} style={{position:'absolute', left:'10px', top:'10px', color:'#94a3b8'}}/><input type="text" placeholder="Filtrar..." style={{...inputStyle, paddingLeft:'32px', marginBottom:0, height:'35px', background:'#f8fafc', fontSize:'12px'}} value={busquedaTabla} onChange={(e)=>setBusquedaTabla(e.target.value)} /></div>
                    <div style={{overflowX:'auto'}}><table style={tableStyle}><thead><tr><th style={thStyle}>LÍNEA / RESP</th><th style={thStyle}>PRESP.</th><th style={thStyle}>GASTO</th><th style={thStyle}>SALDO</th></tr></thead><tbody>
                        {lineasTablaFiltradas.map(l => (
                            <tr key={l.id}>
                                <td style={tdStyle}><b>{l.linea_nombre}</b><br/><small>{l.responsable}</small></td>
                                <td style={tdStyle}>L{l.monto_inicial.toLocaleString()}</td><td style={tdStyle}>L{(l.monto_inicial - l.monto_actual).toLocaleString()}</td><td style={{...tdStyle, color: COLOR_IP_PRIMARY, fontWeight:800}}>L{l.monto_actual.toLocaleString()}</td>
                            </tr>
                        ))}
                    </tbody></table></div></div>
                )}
            </div>
        )}

        {seccion === 'perfil' && (
            <div style={{marginTop:'15px'}}>
                <div style={{...card, background: COLOR_IP_PRIMARY, color:'white', marginBottom:'20px'}}><h3 style={{color:'white', margin:0}}><UserCheck size={18}/> Mi Resumen Anual</h3></div>
                {[...new Set(lineas.filter(l => l.responsable?.trim().toLowerCase() === session?.user?.email?.toLowerCase()).map(l => l.linea_nombre))].map(nombreLinea => {
                    const mesesLinea = lineas.filter(l => l.linea_nombre === nombreLinea);
                    const totalAnualActual = mesesLinea.reduce((a, b) => a + b.monto_actual, 0);
                    const totalAnualInicial = mesesLinea.reduce((a, b) => a + b.monto_inicial, 0);
                    const isExpanded = lineaExpandida === nombreLinea;
                    return (
                        <div key={nombreLinea} style={{...card, marginBottom:'15px', borderLeft:`6px solid ${COLOR_ACCENT}`, cursor:'pointer'}} onClick={() => setLineaExpandida(isExpanded ? null : nombreLinea)}>
                            <div style={{display:'flex', justifyContent:'space-between', alignItems:'center'}}>
                                <div><b style={{fontSize:'13px', color: COLOR_IP_PRIMARY}}>{nombreLinea.toUpperCase()}</b><br/><span style={{...dashLabel, fontSize:'11px'}}>DISPONIBLE: </span><b style={{fontSize:'14px', color: '#16a34a'}}>L{totalAnualActual.toLocaleString()}</b><br/><span style={{fontSize:'10px', color: '#94a3b8'}}>Presp. Total: L{totalAnualInicial.toLocaleString()}</span></div>
                                <button style={{background:'none', border:'none', color:COLOR_IP_PRIMARY}}>{isExpanded ? <ChevronUp size={20}/> : <ChevronDown size={20}/>}</button>
                            </div>
                            {isExpanded && (
                                <div style={{display:'grid', gridTemplateColumns:'repeat(4, 1fr)', gap:'8px', marginTop:'15px', borderTop:'1px solid #f1f5f9', paddingTop:'10px'}}>
                                    {mesesLinea.sort((a,b) => MESES.indexOf(a.mes) - MESES.indexOf(b.mes)).map(m => (
                                        <div key={m.id} style={{textAlign:'center', background:'#f8fafc', padding:'5px', borderRadius:'6px'}}><div style={{fontSize:'9px', fontWeight:800, color:'#64748b'}}>{m.mes}</div><div style={{fontSize:'10px', fontWeight:700, color: '#475569'}}>L{m.monto_actual.toLocaleString()}</div></div>
                                    ))}
                                </div>
                            )}
                        </div>
                    );
                })}
            </div>
        )}

        {seccion === 'config' && (
          <div style={card}><h3 style={cardTitle}><UploadCloud size={18}/> Cargar presupuesto</h3>
            <p style={{fontSize:'11px', color:'#64748b', marginBottom:'15px'}}>Sube el Excel para resetear el presupuesto anual de todas las líneas.</p>
            <input type="file" onChange={(e) => setArchivoExcel(e.target.files[0])} style={{margin:'15px 0', fontSize:'12px'}} />
            <button onClick={importarExcelIP} style={{...btnPro, background: COLOR_IP_PRIMARY, color:'white'}} disabled={loading}>{loading ? "Cargando..." : "SUBIR EXCEL"}</button>
            <button onClick={() => window.open(`${import.meta.env.VITE_SUPABASE_URL}/storage/v1/object/public/facturas/ultimo_presupuesto.xlsx`, '_blank')} style={{...btnPro, background: '#f8fafc', color: '#475569', border: '1px solid #e2e8f0', marginTop:'10px', display:'flex', alignItems:'center', justifyContent:'center', gap:'8px'}}><Download size={18}/> DESCARGAR ORIGINAL</button>
          </div>
        )}
      </main>
    </div>
  );
}

// ESTILOS
const appContainer = { minHeight:'100vh', background:'#f8fafc', paddingBottom:'110px', fontFamily: "'Plus Jakarta Sans', sans-serif" };
const loginWrapper = { display:'flex', height:'100vh', alignItems:'center', justifyContent:'center', background: COLOR_IP_PRIMARY };
const loginCard = { background:'white', padding:'40px', borderRadius:'24px', textAlign:'center', width:'320px', boxShadow: '0 20px 25px -5px rgba(0,0,0,0.1)' };
const headerStyle = { color:'white', padding:'20px', display:'flex', justifyContent:'space-between', alignItems:'center', position:'sticky', top:0, zIndex:100 };
const mainStyle = { padding:'15px', maxWidth:'500px', margin:'0 auto' };
const card = { background:'white', padding:'20px', borderRadius:'16px', boxShadow:'0 4px 6px rgba(0,0,0,0.05)', boxSizing:'border-box' };
const cardTitle = { fontSize:'14px', fontWeight:800, display:'flex', alignItems:'center', gap:'8px' };
const inputStyle = { width:'100%', padding:'12px', borderRadius:'12px', border:'1px solid #e2e8f0', marginBottom:'12px', fontSize:'14px', boxSizing:'border-box' };
const btnPro = { width:'100%', padding:'15px', border:'none', borderRadius:'12px', fontWeight:800, cursor:'pointer' };
const cameraBtn = { display:'flex', alignItems:'center', justifyContent:'center', gap:'10px', padding:'12px', borderRadius:'12px', marginBottom:'15px', fontSize:'13px', cursor:'pointer', fontWeight:600, background: '#f1f5f9' };
const dashboardGrid = { display:'flex', background:'white', borderRadius:'16px', overflow:'hidden' };
const dashItem = { flex:1, padding:'15px', textAlign:'center' };
const dashLabel = { fontSize:'10px', color:'#64748b', fontWeight:800 };
const navBar = { position:'fixed', bottom:0, width:'100%', background:'white', height:'90px', left:0, display:'flex', justifyContent:'space-around', alignItems:'center', borderTop:'1px solid #f1f5f9', zIndex:1000 };
const navBtn = { border:'none', background:'none', display:'flex', flexDirection:'column', alignItems:'center', color:'#94a3b8', fontSize:'11px', fontWeight:700, gap:'5px', cursor: 'pointer' };
const navBtnActive = { ...navBtn, color: COLOR_IP_PRIMARY };
const historyItem = { padding:'15px 0', borderBottom:'1px solid #f1f5f9', display:'flex', justifyContent:'space-between', alignItems:'center' };
const toggleContainer = { display:'flex', background:'#f1f5f9', borderRadius:'14px', padding:'4px' };
const toggleActive = { flex:1, border:'none', background:'white', padding:'10px', borderRadius:'12px', fontSize:'11px', fontWeight:800, color: COLOR_IP_PRIMARY, cursor: 'pointer' };
const toggleInactive = { flex:1, border:'none', background:'none', padding:'10px', color:'#94a3b8', fontSize:'11px', fontWeight:700, cursor: 'pointer' };
const logoutBtn = { background:'rgba(255,255,255,0.15)', border:'none', color:'white', padding:'8px', borderRadius:'10px', cursor: 'pointer' };
const eyeBtn = { background:'#f1f5f9', border:'none', padding:'8px', borderRadius:'10px', cursor:'pointer' };
const tableStyle = { width:'100%', borderCollapse:'collapse', fontSize:'12px' };
const thStyle = { padding:'12px 10px', textAlign:'left', borderBottom:'2px solid #f1f5f9', color:'#64748b', fontWeight:800 };
const tdStyle = { padding:'12px 10px', borderBottom:'1px solid #f8fafc', color:'#1e293b' };
const btnAccordion = { width:'100%', padding:'14px', background:'white', border:'1px solid #e2e8f0', borderRadius:'14px', fontWeight:800, color: COLOR_IP_PRIMARY, marginTop:'15px', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:'8px' };

export default App;