import React, { useState, useEffect, useMemo } from 'react';
import { createClient } from '@supabase/supabase-js';
import * as XLSX from 'xlsx';
import { 
  LayoutDashboard, Receipt, UploadCloud, History, 
  Camera, LogOut, FileSpreadsheet, BarChart3, UserCheck, ShieldCheck, Eye, ChevronDown, ChevronUp
} from 'lucide-react';

const supabase = createClient(import.meta.env.VITE_SUPABASE_URL, import.meta.env.VITE_SUPABASE_ANON_KEY);
const MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
const COLOR_IP_PRIMARY = "#005aba"; 
const COLOR_ACCENT = "#ffd100";
const HOY = new Date().toISOString().split('T')[0];

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
  
  const [mesResumen, setMesResumen] = useState(MESES[new Date().getMonth()]);
  const [mesPersonal, setMesPersonal] = useState(MESES[new Date().getMonth()]);
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

  const importarExcelIP = async () => {
    if (!archivoExcel) return alert("Por favor, selecciona un archivo primero.");
    setLoading(true);
    const reader = new FileReader();
    reader.readAsArrayBuffer(archivoExcel);
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: 0 });

        const mapaMeses = {
          "enero": "Ene", "ene": "Ene", "febrero": "Feb", "feb": "Feb",
          "marzo": "Mar", "mar": "Mar", "abril": "Abr", "abr": "Abr",
          "mayo": "May", "may": "May", "junio": "Jun", "jun": "Jun",
          "julio": "Jul", "jul": "Jul", "agosto": "Ago", "ago": "Ago",
          "septiembre": "Sep", "sep": "Sep", "octubre": "Oct", "oct": "Oct",
          "noviembre": "Nov", "nov": "Nov", "diciembre": "Dic", "dic": "Dic"
        };

        const filasParaSubir = [];
        json.forEach((filaRaw) => {
          const fila = Object.keys(filaRaw).reduce((acc, key) => {
            acc[key.toLowerCase().trim()] = filaRaw[key];
            return acc;
          }, {});
          const nombreLinea = fila["línea"] || fila["linea"];
          const responsable = fila["responsable"];
          if (nombreLinea && responsable && 
              !nombreLinea.toString().toLowerCase().includes("total") &&
              !nombreLinea.toString().toLowerCase().includes("resumen")) {
            Object.keys(mapaMeses).forEach(mesExcel => {
              if (fila[mesExcel] !== undefined) {
                const mesSistema = mapaMeses[mesExcel];
                const valorRaw = fila[mesExcel].toString().replace(/[^\d.]/g, "");
                const monto = parseFloat(valorRaw) || 0;
                filasParaSubir.push({
                  linea_nombre: nombreLinea.toString().trim(),
                  responsable: responsable.toString().trim(),
                  mes: mesSistema,
                  monto_inicial: monto,
                  monto_actual: monto
                });
              }
            });
          }
        });
        if (filasParaSubir.length === 0) throw new Error("No se detectaron datos válidos.");
        await supabase.from('presupuestos').delete().neq('id', 0);
        const { error } = await supabase.from('presupuestos').insert(filasParaSubir);
        if(error) throw error;
        alert("✅ Importación limpia con éxito.");
        setArchivoExcel(null);
        obtenerDatos();
      } catch (err) { alert("Error: " + err.message); } finally { setLoading(false); }
    };
  };

  const registrarGasto = async () => {
    if (!compra.lineaId || !compra.monto || !compra.foto) return alert("🚫 Campos obligatorios vacíos.");
    const montoGasto = parseFloat(compra.monto);
    const fechaObj = new Date(compra.fecha + 'T12:00:00');
    const mesGastoIdx = fechaObj.getMonth();
    
    const lineaSel = lineas.find(l => l.id.toString() === compra.lineaId.toString());
    const todasIguales = lineas.filter(l => l.linea_nombre === lineaSel.linea_nombre);
    
    const lineasDisponibles = todasIguales
      .filter(l => MESES.indexOf(l.mes) <= mesGastoIdx)
      .sort((a, b) => MESES.indexOf(a.mes) - MESES.indexOf(b.mes));

    const totalDisponibleAcumulado = lineasDisponibles.reduce((a, b) => a + b.monto_actual, 0);

    if (montoGasto > totalDisponibleAcumulado) {
      alert(`🚫 FONDOS INSUFICIENTES: No se permiten sobregiros.\n\nDisponible acumulado (Ene - ${lineaSel.mes}): L${totalDisponibleAcumulado.toLocaleString()}\nIntentado: L${montoGasto.toLocaleString()}`);
      return;
    }

    setLoading(true);
    try {
      const ext = compra.foto.name.split('.').pop();
      const nombreFoto = `${Date.now()}.${ext}`;
      await supabase.storage.from('facturas').upload(nombreFoto, compra.foto);
      
      let restante = montoGasto;
      for (let l of lineasDisponibles) {
        if (restante <= 0) break;
        let disponibleEnMes = l.monto_actual;
        if (disponibleEnMes <= 0) continue;
        let aQuitar = Math.min(disponibleEnMes, restante);
        await supabase.from('presupuestos').update({ monto_actual: disponibleEnMes - aQuitar }).eq('id', l.id);
        restante -= aQuitar;
      }

      await supabase.from('compras').insert([{ 
        presupuesto_id: lineaSel.id, monto_lps: montoGasto, descripcion: compra.desc, fecha: compra.fecha, url_factura: nombreFoto, creado_por: session.user.email 
      }]);

      alert("✅ Gasto registrado con éxito.");
      setCompra({ ...compra, monto: '', desc: '', foto: null, lineaId: '' });
      obtenerDatos();
    } catch (err) { alert(err.message); } finally { setLoading(false); }
  };

  const descargarExcel = () => {
    let data = [];
    if (tabReporte === 'mensual') {
      data = lineas.filter(l => l.mes === mesResumen && l.monto_inicial > 0).map(l => ({
        "Línea": l.linea_nombre, "Responsable": l.responsable, "Presupuesto Inicial": l.monto_inicial, "Gasto Mes": l.monto_inicial - l.monto_actual, "Saldo": l.monto_actual
      }));
    } else {
      const unicos = [...new Set(lineas.map(l => l.linea_nombre))];
      data = unicos.map(n => {
        const lN = lineas.filter(l => l.linea_nombre === n);
        const ini = lN.reduce((a, b) => a + b.monto_inicial, 0);
        const act = lN.reduce((a, b) => a + b.monto_actual, 0);
        return { "Línea": n, "Responsable": lN[0].responsable, "Presp. Anual": ini, "Gasto Anual": ini - act, "Saldo Anual": act };
      }).filter(item => item["Presp. Anual"] > 0);
    }
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Reporte_IP");
    XLSX.writeFile(wb, `Reporte_IP_${tabReporte}.xlsx`);
  };

  const stats = useMemo(() => {
    const filtradas = tabReporte === 'mensual' ? lineas.filter(l => l.mes === mesResumen) : lineas;
    const tP = filtradas.reduce((a, b) => a + (b.monto_inicial || 0), 0);
    const tD = filtradas.reduce((a, b) => a + (b.monto_actual || 0), 0);
    const tG = tabReporte === 'mensual' ? (tP - tD) : historial.reduce((a, b) => a + (b.monto_lps || 0), 0);
    return { tP, tG, tD: tP - tG };
  }, [lineas, tabReporte, mesResumen, historial]);

  const lineasTabla = useMemo(() => {
    if (tabReporte === 'mensual') {
        return lineas.filter(l => l.mes === mesResumen && l.monto_inicial > 0);
    }
    const unicos = [...new Set(lineas.map(l => l.linea_nombre))];
    return unicos.map(n => {
      const grupo = lineas.filter(l => l.linea_nombre === n);
      return {
        id: n, linea_nombre: n, responsable: grupo[0].responsable,
        monto_inicial: grupo.reduce((a, b) => a + b.monto_inicial, 0),
        monto_actual: grupo.reduce((a, b) => a + b.monto_actual, 0)
      };
    }).filter(item => item.monto_inicial > 0);
  }, [lineas, tabReporte, mesResumen]);

  if (!session) return (
    <div style={loginWrapper}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;800&display=swap'); body { font-family: 'Plus Jakarta Sans', sans-serif; }`}</style>
      <div style={loginCard}>
        <h1 style={{color: COLOR_IP_PRIMARY, fontWeight: 800}}>CONTROL DE PRESUPUESTO</h1>
        <form onSubmit={async (e)=>{e.preventDefault(); const {error}=await supabase.auth.signInWithPassword({email, password}); if(error) alert("Error");}}>
          <input type="email" placeholder="Usuario" style={inputStyle} onChange={e=>setEmail(e.target.value)} required />
          <input type="password" placeholder="Contraseña" style={inputStyle} onChange={e=>setPassword(e.target.value)} required />
          <button style={{...btnPro, background: COLOR_IP_PRIMARY}}>INGRESAR</button>
        </form>
      </div>
    </div>
  );

  return (
    <div style={appContainer}>
      <header style={{...headerStyle, background: COLOR_IP_PRIMARY}}>
        <div style={{display:'flex', alignItems:'center', gap:'12px'}}><ShieldCheck size={22} color={COLOR_ACCENT}/><span style={{fontWeight: 800}}>CONTROL PRESUPUESTARIO - IP</span></div>
        <button onClick={()=>supabase.auth.signOut()} style={logoutBtn}><LogOut size={18}/></button>
      </header>

      <main style={mainStyle}>
        {seccion === 'reportes' && (
          <div style={toggleContainer}>
            <button onClick={()=>setTabReporte('mensual')} style={tabReporte==='mensual'?toggleActive:toggleInactive}>MENSUAL</button>
            <button onClick={()=>setTabReporte('anual')} style={tabReporte==='anual'?toggleActive:toggleInactive}>ANUAL</button>
          </div>
        )}

        {seccion === 'compras' && (
          <div style={card}>
            <h3 style={cardTitle}><Receipt size={18} color={COLOR_IP_PRIMARY}/> Nuevo Registro</h3>
            <input type="date" style={inputStyle} value={compra.fecha} max={HOY} onChange={(e)=>setCompra({...compra, fecha:e.target.value})} />
            <select style={inputStyle} value={compra.lineaId} onChange={(e)=>setCompra({...compra, lineaId:e.target.value})}>
              <option value="">Línea...</option>
              {lineas.filter(l => l.mes === MESES[new Date(compra.fecha + 'T12:00:00').getMonth()] && l.monto_inicial > 0).map(l => (
                <option key={l.id} value={l.id}>{l.linea_nombre} (Saldo: L{l.monto_actual.toLocaleString()})</option>
              ))}
            </select>
            <input type="number" placeholder="Monto Lps" style={inputStyle} value={compra.monto} onChange={(e)=>setCompra({...compra, monto:e.target.value})} />
            <input type="text" placeholder="Descripción" style={inputStyle} value={compra.desc} onChange={(e)=>setCompra({...compra, desc:e.target.value})} />
<label style={{...cameraBtn, background: compra.foto ? COLOR_ACCENT : '#f1f5f9'}}><Camera size={18}/> <span>{compra.foto ? "ARCHIVO LISTO ✅" : "ADJUNTAR FACTURA"}</span><input type="file" accept="image/*,.pdf" hidden onChange={(e)=>setCompra({...compra, foto:e.target.files[0]})} /></label>            <button onClick={registrarGasto} style={{...btnPro, background: loading ? '#cbd5e1' : COLOR_IP_PRIMARY}} disabled={loading}>{loading ? "PROCESANDO..." : "REGISTRAR GASTO"}</button>
          </div>
        )}

        {seccion === 'reportes' && (
            <div style={{marginTop:'15px'}}>
                <div style={dashboardGrid}>
                    <div style={dashItem}><span style={dashLabel}>PRESUPUESTO</span><br/><b>L{stats.tP.toLocaleString()}</b></div>
                    <div style={{...dashItem, borderLeft:'1px solid #f1f5f9', borderRight:'1px solid #f1f5f9'}}><span style={dashLabel}>GASTADO</span><br/><b style={{color:'#dc2626'}}>L{stats.tG.toLocaleString()}</b></div>
                    <div style={dashItem}><span style={dashLabel}>DISPONIBLE</span><br/><b style={{color:'#005aba'}}>L{stats.tD.toLocaleString()}</b></div>
                </div>
                <div style={{...card, marginTop:'20px'}}>
                  <div style={{display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:'15px'}}>
                    <h3 style={cardTitle}><BarChart3 size={18} color={COLOR_IP_PRIMARY}/> HISTORIAL</h3>
                    <div style={{display:'flex', gap:'8px'}}>
                      {tabReporte === 'mensual' && <select style={{...inputStyle, width:'auto', marginBottom:0, padding:'5px 10px'}} value={mesResumen} onChange={(e)=>setMesResumen(e.target.value)}>{MESES.map(m=><option key={m} value={m}>{m}</option>)}</select>}
                      <button onClick={descargarExcel} style={eyeBtn}><FileSpreadsheet size={16} color="#16a34a"/></button>
                    </div>
                  </div>
                  {historial.filter(h => tabReporte === 'anual' || h.presupuestos?.mes === mesResumen).slice(0, 15).map(h => (
                    <div key={h.id} style={historyItem}>
                      <div style={{flex:1}}>
                        <div style={{fontWeight:700, fontSize:'12px'}}>{h.presupuestos?.linea_nombre}</div>
                        <div style={{fontSize:'10px', color:'#94a3b8'}}>{h.fecha} • {h.creado_por || 'Sistema'}</div>
                      </div>
                      <div style={{textAlign:'right', display:'flex', alignItems:'center', gap:'8px'}}>
                        <b style={{fontSize:'12px', color:'#dc2626'}}>-L{h.monto_lps.toLocaleString()}</b>
                        <button onClick={() => window.open(`${import.meta.env.VITE_SUPABASE_URL}/storage/v1/object/public/facturas/${h.url_factura}`, '_blank')} style={eyeBtn}><Eye size={14}/></button>
                      </div>
                    </div>
                  ))}
                </div>
                <button onClick={()=>setVerDetalle(!verDetalle)} style={btnAccordion}>{verDetalle ? <ChevronUp size={16}/> : <ChevronDown size={16}/>} TABLA DETALLADA</button>
                {verDetalle && (
                  <div style={{...card, marginTop:'10px', padding:'10px', overflowX:'auto'}}>
                    <table style={tableStyle}>
                      <thead><tr><th style={thStyle}>LÍNEA / RESP</th><th style={thStyle}>PRESP.</th><th style={thStyle}>GASTO</th><th style={thStyle}>SALDO</th></tr></thead>
                      <tbody>
                        {lineasTabla.map(l => (
                          <tr key={l.id}>
                            <td style={tdStyle}><b>{l.linea_nombre}</b><br/><small>{l.responsable}</small></td>
                            <td style={tdStyle}>L{l.monto_inicial.toLocaleString()}</td>
                            <td style={tdStyle}>L{(l.monto_inicial - l.monto_actual).toLocaleString()}</td>
                            <td style={{...tdStyle, color: l.monto_actual < 0 ? '#dc2626' : COLOR_IP_PRIMARY, fontWeight:800}}>L{l.monto_actual.toLocaleString()}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
            </div>
        )}

        {seccion === 'perfil' && (
          <div>
            <div style={{...card, background: COLOR_IP_PRIMARY, color:'white', marginBottom:'20px'}}>
              <h3 style={{...cardTitle, color:'white'}}><UserCheck size={18}/> Mi Panel</h3>
              <select style={{...inputStyle, background:'rgba(255,255,255,0.1)', color:'white', border:'none'}} value={mesPersonal} onChange={(e)=>setMesPersonal(e.target.value)}>
                {MESES.map(m => <option key={m} value={m} style={{color:'black'}}>{m}</option>)}
              </select>
            </div>
            {lineas.filter(l => l.responsable.trim().toLowerCase() === session?.user?.email?.toLowerCase() && l.mes === mesPersonal && l.monto_inicial > 0).map(l => {
                const todasLineasMismoNombre = lineas.filter(item => item.linea_nombre === l.linea_nombre);
                const mesIdx = MESES.indexOf(mesPersonal);
                const tDispAcumulado = todasLineasMismoNombre
                  .filter(item => MESES.indexOf(item.mes) <= mesIdx)
                  .reduce((a, b) => a + b.monto_actual, 0);

                return (
                  <div key={l.id} style={{...card, marginBottom:'12px', borderLeft:`6px solid ${COLOR_ACCENT}`, paddingBottom: '10px'}}>
                    <div style={{fontWeight:800, fontSize:'13px', color: COLOR_IP_PRIMARY}}>{l.linea_nombre.toUpperCase()}</div>
                    <div style={{display:'flex', justifyContent:'space-between', marginTop:'12px'}}>
                      <div><span style={dashLabel}>PRESUPUESTO MES:</span><br/><b>L{l.monto_inicial.toLocaleString()}</b></div>
                      <div style={{textAlign:'right'}}><span style={dashLabel}>DISPONIBLE ACUM.:</span><br/><b style={{color: tDispAcumulado <= 0 ? '#dc2626' : '#16a34a'}}>L{tDispAcumulado.toLocaleString()}</b></div>
                    </div>
                  </div>
                );
            })}
          </div>
        )}

        {seccion === 'config' && (
          <div style={card}>
            <h3 style={cardTitle}><UploadCloud size={18}/> Excel</h3>
            <input type="file" accept=".xlsx, .xls" style={{margin:'20px 0', fontSize:'12px'}} onChange={(e) => setArchivoExcel(e.target.files[0])} />
            <button onClick={importarExcelIP} style={{...btnPro, background: loading ? '#94a3b8' : '#16a34a'}} disabled={loading}>{loading ? "PROCESANDO..." : "SUBIR PRESUPUESTO"}</button>
          </div>
        )}
      </main>

      <nav style={navBar}>
        <button onClick={()=>setSeccion('compras')} style={seccion==='compras'?navBtnActive:navBtn}><Receipt size={24}/><span>Gasto</span></button>
        <button onClick={()=>setSeccion('reportes')} style={seccion==='reportes'?navBtnActive:navBtn}><LayoutDashboard size={24}/><span>Panel</span></button>
        <button onClick={()=>setSeccion('perfil')} style={seccion==='perfil'?navBtnActive:navBtn}><UserCheck size={24}/><span>Mi IP</span></button>
        <button onClick={()=>setSeccion('config')} style={seccion==='config'?navBtnActive:navBtn}><UploadCloud size={24}/><span>Excel</span></button>
      </nav>
    </div>
  );
}

const appContainer = { minHeight:'100vh', background:'#f8fafc', paddingBottom:'110px', fontFamily:"'Plus Jakarta Sans', sans-serif" };
const loginWrapper = { display:'flex', height:'100vh', alignItems:'center', justifyContent:'center', background: COLOR_IP_PRIMARY };
const loginCard = { background:'white', padding:'40px', borderRadius:'24px', textAlign:'center', width:'320px', boxShadow:'0 20px 25px -5px rgba(0,0,0,0.1)' };
const headerStyle = { color:'white', padding:'20px', display:'flex', justifyContent:'space-between', alignItems:'center', position:'sticky', top:0, zIndex:100, boxShadow:'0 4px 12px rgba(0,0,0,0.1)', borderBottom: `3px solid ${COLOR_ACCENT}` };
const mainStyle = { padding:'15px', maxWidth:'500px', margin:'0 auto' };
const card = { background:'white', padding:'20px', borderRadius:'16px', boxShadow:'0 10px 15px -3px rgba(0,0,0,0.05)', boxSizing:'border-box' };
const cardTitle = { fontSize:'13px', fontWeight:800, display:'flex', alignItems:'center', gap:'8px', margin:0 };
const inputStyle = { width:'100%', padding:'12px', borderRadius:'10px', border:'1px solid #e2e8f0', marginBottom:'12px', fontSize:'13px', boxSizing:'border-box', fontWeight:600, background:'#f8fafc', outline:'none' };
const btnPro = { width:'100%', padding:'15px', color:'white', border:'none', borderRadius:'12px', fontWeight:800, cursor:'pointer' };
const cameraBtn = { display:'flex', alignItems:'center', justifyContent:'center', gap:'10px', padding:'12px', background:'#f1f5f9', borderRadius:'10px', marginBottom:'15px', fontSize:'12px', cursor:'pointer', fontWeight:700 };
const dashboardGrid = { display:'flex', background:'white', borderRadius:'16px', boxShadow:'0 10px 15px -3px rgba(0,0,0,0.05)', overflow:'hidden' };
const dashItem = { flex:1, padding:'20px', textAlign:'center' };
const dashLabel = { fontSize:'9px', color:'#64748b', fontWeight:800 };
const navBar = { position:'fixed', bottom:0, width:'100%', background:'white', borderTop:'1px solid #f1f5f9', height:'90px', left:0, display:'flex', justifyContent:'space-around', alignItems:'center' };
const navBtn = { border:'none', background:'none', display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', color:'#94a3b8', fontSize:'10px', fontWeight:700, gap:'5px' };
const navBtnActive = { ...navBtn, color: COLOR_IP_PRIMARY };
const historyItem = { padding:'12px 0', borderBottom:'1px solid #f8fafc', display:'flex', justifyContent:'space-between', alignItems:'center' };
const toggleContainer = { display:'flex', background:'#f1f5f9', borderRadius:'12px', padding:'4px' };
const toggleActive = { flex:1, border:'none', background:'white', padding:'10px', borderRadius:'10px', fontSize:'11px', fontWeight:800, color: COLOR_IP_PRIMARY, boxShadow:'0 2px 4px rgba(0,0,0,0.1)' };
const toggleInactive = { flex:1, border:'none', background:'none', padding:'10px', borderRadius:'10px', fontSize:'11px', fontWeight:700, color:'#94a3b8', cursor:'pointer' };
const logoutBtn = { background:'rgba(255,255,255,0.1)', border:'none', color:'white', padding:'8px', borderRadius:'10px' };
const eyeBtn = { background:'#f1f5f9', border:'none', padding:'8px', borderRadius:'8px', cursor:'pointer' };
const tableStyle = { width:'100%', borderCollapse:'collapse', fontSize:'11px' };
const thStyle = { padding:'10px', textAlign:'left', borderBottom:'2px solid #f1f5f9', color:'#64748b' };
const tdStyle = { padding:'10px', borderBottom:'1px solid #f8fafc' };
const btnAccordion = { width:'100%', padding:'12px', background:'white', border:'1px solid #e2e8f0', borderRadius:'12px', fontWeight:800, color: COLOR_IP_PRIMARY, marginTop:'15px', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:'8px' };

export default App;