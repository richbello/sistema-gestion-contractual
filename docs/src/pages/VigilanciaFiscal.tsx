import { useState } from 'react';
import { PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from 'recharts';
import { ENTIDADES } from './entidades';

interface Alerta {
  fuente: string; nivel: string; motivo: string; documento: string;
  proveedor_secop: string; nombre_fuente: string; tipo: string; sancion: string;
  fecha_sancion: string; fin_inhabilidad: string; vigente: string; entidad: string;
  valor: string; estado: string; fecha_firma: string; objeto: string; autoridad: string;
}
interface Resultado {
  entidad: string; fecha_analisis: string; total_contratos: number; total_alertas: number;
  criticas: number; altas: number; medias: number; info: number; registros_siri?: number; registros_boletin?: number; alertas: Alerta[];
}

const COLORES: Record<string, string> = { CRITICA: '#dc2626', ALTA: '#ea580c', MEDIA: '#ca8a04', INFO: '#64748b' };
const fmtCOP = (v: string) => { const n = parseInt(v || '0', 10); return isNaN(n) ? '$0' : '$' + n.toLocaleString('es-CO'); };

export default function VigilanciaFiscal() {
  const [entidad, setEntidad] = useState('USME');
  const [resultado, setResultado] = useState<Resultado | null>(null);
  const [loading, setLoading] = useState(false);
  const [filtro, setFiltro] = useState('TODAS');
  const [busqueda, setBusqueda] = useState('');
  const [abierta, setAbierta] = useState<number | null>(null);
  const [docBusca, setDocBusca] = useState('');
  const [docResult, setDocResult] = useState<any>(null);
  const [docLoading, setDocLoading] = useState(false);
  const [progreso, setProgreso] = useState<string[]>([]);
  const [cronometro, setCronometro] = useState(0);

  const analizar = () => {
    setLoading(true); setAbierta(null); setResultado(null); setProgreso([]); setCronometro(0);
    const t0 = Date.now();
    const timer = setInterval(() => setCronometro(Math.floor((Date.now() - t0) / 1000)), 1000);
    const es = new EventSource(`https://sistema-gestion-contractual.onrender.com/api/vfi/analizar-stream?entidad=${encodeURIComponent(entidad)}`);
    es.addEventListener('progress', (e: any) => {
      setProgreso(prev => [...prev, JSON.parse(e.data)]);
    });
    es.addEventListener('done', (e: any) => {
      setResultado(JSON.parse(e.data));
      setLoading(false); clearInterval(timer); es.close();
    });
    es.addEventListener('error', () => {
      alert('Error en el análisis. Verifica que el backend esté corriendo.');
      setLoading(false); clearInterval(timer); es.close();
    });
  };

  const buscarDoc = async () => {
    if (!docBusca.trim()) return;
    setDocLoading(true); setDocResult(null);
    try {
      const r = await fetch(`https://sistema-gestion-contractual.onrender.com/api/vfi/buscar?doc=${encodeURIComponent(docBusca)}`, { method: 'POST' });
      setDocResult(await r.json());
    } catch { alert('Error conectando al backend.'); }
    setDocLoading(false);
  };

  const limpiar = () => {
    setResultado(null);
    setDocResult(null);
    setDocBusca('');
    setBusqueda('');
    setFiltro('TODAS');
    setAbierta(null);
  };

  const alertas = (resultado?.alertas || []).filter(a => {
    const okNivel = filtro === 'TODAS' || a.nivel === filtro;
    const q = busqueda.toLowerCase();
    const okBusca = !q || a.proveedor_secop.toLowerCase().includes(q) || a.documento.includes(q) ||
      a.sancion.toLowerCase().includes(q) || a.entidad.toLowerCase().includes(q) || (a.objeto || '').toLowerCase().includes(q);
    return okNivel && okBusca;
  });

  const dataNivel = resultado ? [
    { name: 'Críticas', value: resultado.criticas, color: COLORES.CRITICA },
    { name: 'Altas', value: resultado.altas, color: COLORES.ALTA },
    { name: 'Medias', value: resultado.medias, color: COLORES.MEDIA },
    { name: 'Info', value: resultado.info, color: COLORES.INFO },
  ].filter(d => d.value > 0) : [];

  const dataFuente = Object.values((resultado?.alertas || []).reduce((acc: Record<string, { name: string; value: number }>, a) => {
    const key = a.fuente.includes('SIRI') ? 'SIRI · Procuraduría' : 'Boletín · Contraloría';
    acc[key] = acc[key] || { name: key, value: 0 }; acc[key].value += 1; return acc;
  }, {}));

  return (
    <div style={{ fontFamily: "'Segoe UI', sans-serif", color: '#0f172a', maxWidth: 1200, margin: '0 auto' }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 14, marginBottom: 8 }}>
        <div style={{ fontSize: 34 }}>⚖️</div>
        <div>
          <h1 style={{ margin: 0, fontSize: 26, fontWeight: 800 }}>Vigilancia Fiscal IA</h1>
          <p style={{ margin: 0, color: '#64748b', fontSize: 13 }}>Cruce en tiempo real · SECOP II × SIRI (Procuraduría) × Boletín (Contraloría)</p>
        </div>
      </div>

      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', marginBottom: 18 }}>
        <FuenteChip color="#2563eb" titulo="SECOP II" detalle="Contratación pública · tiempo real" />
        <FuenteChip color="#7c3aed" titulo="SIRI · Procuraduría" detalle={resultado?.registros_siri ? `${resultado.registros_siri.toLocaleString('es-CO')} registros disciplinarios` : 'Antecedentes disciplinarios · refresco periódico'} />
        <FuenteChip color="#ca8a04" titulo="Boletín · Contraloría" detalle={resultado?.registros_boletin ? `${resultado.registros_boletin.toLocaleString('es-CO')} responsables fiscales · trimestral` : 'Responsables fiscales · trimestral'} />
      </div>

      <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 14, padding: 18, marginBottom: 18, boxShadow: '0 1px 3px rgba(0,0,0,.06)' }}>
        <label style={{ display: 'block', fontSize: 11, fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 6 }}>Entidad a auditar</label>
        <div style={{ display: 'flex', gap: 12, alignItems: 'flex-end', flexWrap: 'wrap' }}>
          <select value={entidad} onChange={e => setEntidad(e.target.value)}
            style={{ flex: 1, minWidth: 220, padding: '11px 14px', border: '1.5px solid #cbd5e1', borderRadius: 9, fontSize: 14, boxSizing: 'border-box', background: '#fff' }}>
            {ENTIDADES.map(g => (
              <optgroup key={g.grupo} label={g.grupo}>
                {g.items.map(it => <option key={it} value={it}>{it}</option>)}
              </optgroup>
            ))}
          </select>
          <button onClick={analizar} disabled={loading}
            style={{ background: loading ? '#94a3b8' : 'linear-gradient(135deg,#2563eb,#1d4ed8)', color: '#fff', border: 'none', borderRadius: 9, padding: '12px 26px', fontWeight: 700, fontSize: 14, cursor: loading ? 'default' : 'pointer', whiteSpace: 'nowrap' }}>
            {loading ? 'Analizando…' : 'Analizar'}
          </button>
        </div>
        {loading && (
          <div style={{ marginTop: 14, background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 10, padding: 14 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
              <span style={{ fontSize: 12, fontWeight: 700, color: '#2563eb' }}>Procesando análisis…</span>
              <span style={{ fontSize: 13, fontWeight: 800, color: '#0f172a', fontFamily: 'monospace' }}>
                {Math.floor(cronometro / 60)}:{String(cronometro % 60).padStart(2, '0')}
              </span>
            </div>
            <div style={{ maxHeight: 140, overflowY: 'auto', fontSize: 12, color: '#475569', fontFamily: 'monospace', lineHeight: 1.6 }}>
              {progreso.map((m, i) => {
                let texto = m;
                let color = '#64748b';
                if (m.includes('SIRI')) { texto = '📋 Procuraduría (SIRI): ' + m.replace(/\[.*?\]/, '').trim(); color = '#7c3aed'; }
                else if (m.includes('Boletín') || m.includes('Boletin')) { texto = '💰 Contraloría (Boletín): ' + m.replace(/\[.*?\]/, '').trim(); color = '#ca8a04'; }
                else if (m.includes('contratos') || m.includes('SECOP')) { texto = '📄 SECOP II: ' + m.replace(/\[.*?\]/, '').trim(); color = '#2563eb'; }
                return <div key={i} style={{ color }}>{texto}</div>;
              })}
              {progreso.length === 0 && <div>Iniciando…</div>}
            </div>
          </div>
        )}
      </div>

      <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 14, padding: 18, marginBottom: 18, boxShadow: '0 1px 3px rgba(0,0,0,.06)' }}>
        <label style={{ display: 'block', fontSize: 11, fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 6 }}>Buscar persona o empresa por documento</label>
        <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap' }}>
          <input value={docBusca} onChange={e => setDocBusca(e.target.value)} onKeyDown={e => e.key === 'Enter' && buscarDoc()}
            placeholder="Ej: 51699228"
            style={{ flex: 1, minWidth: 220, padding: '11px 14px', border: '1.5px solid #cbd5e1', borderRadius: 9, fontSize: 14, boxSizing: 'border-box' }} />
          <button onClick={buscarDoc} disabled={docLoading}
            style={{ background: docLoading ? '#94a3b8' : 'linear-gradient(135deg,#059669,#047857)', color: '#fff', border: 'none', borderRadius: 9, padding: '12px 26px', fontWeight: 700, fontSize: 14, cursor: docLoading ? 'default' : 'pointer', whiteSpace: 'nowrap' }}>
            {docLoading ? 'Buscando…' : 'Buscar'}
          </button>
          <button onClick={limpiar}
            style={{ background: '#f1f5f9', color: '#475569', border: '1.5px solid #cbd5e1', borderRadius: 9, padding: '12px 22px', fontWeight: 700, fontSize: 14, cursor: 'pointer', whiteSpace: 'nowrap' }}>
            Limpiar
          </button>
        </div>
      </div>

      {resultado && (
        <>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit,minmax(180px,1fr))', gap: 14, marginBottom: 18 }}>
            <Tarjeta label="Contratos analizados" valor={resultado.total_contratos} color="#2563eb" sub={`Entidad: ${resultado.entidad}`} />
            <Tarjeta label="Alertas totales" valor={resultado.total_alertas} color="#0f172a" sub={`Análisis: ${resultado.fecha_analisis}`} />
            <Tarjeta label="Críticas" valor={resultado.criticas} color={COLORES.CRITICA} sub="Contrato posterior a inhabilidad" />
            <Tarjeta label="Altas" valor={resultado.altas} color={COLORES.ALTA} sub="Inhabilidad vigente" />
          </div>

          {resultado.total_alertas > 0 && (
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit,minmax(320px,1fr))', gap: 18, marginBottom: 18 }}>
              <Panel titulo="Distribución por nivel de riesgo">
                <ResponsiveContainer width="100%" height={240}>
                  <PieChart>
                    <Pie data={dataNivel} cx="50%" cy="50%" innerRadius={55} outerRadius={95} paddingAngle={2} dataKey="value" label={(e: any) => `${e.name}: ${e.value}`}>
                      {dataNivel.map((d, i) => <Cell key={i} fill={d.color} />)}
                    </Pie>
                    <Tooltip />
                  </PieChart>
                </ResponsiveContainer>
              </Panel>
              <Panel titulo="Alertas por fuente oficial">
                <ResponsiveContainer width="100%" height={240}>
                  <BarChart data={dataFuente}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                    <XAxis dataKey="name" stroke="#64748b" fontSize={11} />
                    <YAxis stroke="#64748b" fontSize={11} allowDecimals={false} />
                    <Tooltip />
                    <Bar dataKey="value" fill="#2563eb" radius={[6, 6, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </Panel>
            </div>
          )}

          <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 14, padding: 18, boxShadow: '0 1px 3px rgba(0,0,0,.06)' }}>
            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 14 }}>
              {['TODAS', 'CRITICA', 'ALTA', 'MEDIA', 'INFO'].map(n => (
                <button key={n} onClick={() => setFiltro(n)}
                  style={{ padding: '7px 16px', borderRadius: 20, fontSize: 12, fontWeight: 700, cursor: 'pointer', border: 'none', background: filtro === n ? (COLORES[n] || '#2563eb') : '#f1f5f9', color: filtro === n ? '#fff' : '#475569' }}>
                  {n === 'TODAS' ? 'Todas' : n}
                </button>
              ))}
            </div>
            <input value={busqueda} onChange={e => setBusqueda(e.target.value)} placeholder="Buscar por proveedor, documento, sanción, entidad u objeto…"
              style={{ width: '100%', padding: '10px 14px', border: '1.5px solid #cbd5e1', borderRadius: 9, fontSize: 13, boxSizing: 'border-box', marginBottom: 14 }} />
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                <thead>
                  <tr style={{ borderBottom: '2px solid #e2e8f0', textAlign: 'left', color: '#64748b' }}>
                    <th style={{ padding: '10px 8px', width: 30 }}></th>
                    <th style={{ padding: '10px 8px' }}>Nivel</th>
                    <th style={{ padding: '10px 8px' }}>Proveedor</th>
                    <th style={{ padding: '10px 8px' }}>Antecedente</th>
                    <th style={{ padding: '10px 8px' }}>Fuente</th>
                    <th style={{ padding: '10px 8px', textAlign: 'right' }}>Valor contrato</th>
                  </tr>
                </thead>
                <tbody>
                  {alertas.length === 0 && (<tr><td colSpan={6} style={{ padding: 28, textAlign: 'center', color: '#94a3b8' }}>Sin alertas para este filtro.</td></tr>)}
                  {alertas.map((a, i) => (
                    <tr key={i} onClick={() => setAbierta(abierta === i ? null : i)} style={{ borderBottom: '1px solid #f1f5f9', cursor: 'pointer', background: abierta === i ? '#f8fafc' : '#fff' }}>
                      <td style={{ padding: '11px 8px', color: '#94a3b8' }}>{abierta === i ? '▼' : '▶'}</td>
                      <td style={{ padding: '11px 8px' }}><span style={{ background: COLORES[a.nivel] || '#64748b', color: '#fff', padding: '3px 10px', borderRadius: 12, fontSize: 11, fontWeight: 800 }}>{a.nivel}</span></td>
                      <td style={{ padding: '11px 8px' }}><div style={{ fontWeight: 700 }}>{a.proveedor_secop}</div><div style={{ fontSize: 11, color: '#94a3b8' }}>CC {a.documento}</div></td>
                      <td style={{ padding: '11px 8px' }}><div>{a.sancion}</div><div style={{ fontSize: 11, color: '#94a3b8' }}>{a.tipo}</div></td>
                      <td style={{ padding: '11px 8px', fontSize: 12 }}>{a.fuente}</td>
                      <td style={{ padding: '11px 8px', textAlign: 'right', fontFamily: 'monospace', fontWeight: 700 }}>{fmtCOP(a.valor)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
    </div>
  );
}

function Tarjeta({ label, valor, color, sub }: { label: string; valor: number; color: string; sub?: string }) {
  return (
    <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderLeft: `4px solid ${color}`, borderRadius: 12, padding: 16, boxShadow: '0 1px 3px rgba(0,0,0,.06)' }}>
      <div style={{ fontSize: 32, fontWeight: 800, color }}>{valor}</div>
      <div style={{ fontSize: 12, fontWeight: 700, color: '#334155' }}>{label}</div>
      {sub && <div style={{ fontSize: 11, color: '#94a3b8', marginTop: 3 }}>{sub}</div>}
    </div>
  );
}
function Panel({ titulo, children }: { titulo: string; children: React.ReactNode }) {
  return (
    <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderRadius: 14, padding: 18, boxShadow: '0 1px 3px rgba(0,0,0,.06)' }}>
      <div style={{ fontSize: 12, fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 10 }}>{titulo}</div>
      {children}
    </div>
  );
}
function Campo({ label, valor, destaca }: { label: string; valor: string; destaca?: boolean }) {
  return (
    <div>
      <div style={{ fontSize: 11, fontWeight: 700, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: 0.5, marginBottom: 3 }}>{label}</div>
      <div style={{ fontSize: 13, fontWeight: destaca ? 800 : 500, color: destaca ? '#dc2626' : '#334155' }}>{valor}</div>
    </div>
  );
}

function FuenteChip({ color, titulo, detalle }: { color: string; titulo: string; detalle: string }) {
  return (
    <div style={{ background: '#fff', border: '1px solid #e2e8f0', borderLeft: `3px solid ${color}`, borderRadius: 10, padding: '10px 14px', flex: 1, minWidth: 200 }}>
      <div style={{ fontSize: 12, fontWeight: 800, color }}>{titulo}</div>
      <div style={{ fontSize: 11, color: '#64748b', marginTop: 2 }}>{detalle}</div>
    </div>
  );
}
