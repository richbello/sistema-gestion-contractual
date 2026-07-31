// ============================================================================
// ESTADO DE CUENTA — MÓDULO MASIVO (integración Módulo 04)
// - Maneja el submenú (Individual | Masivo) sin tocar el individual.
// - Carga Pyodide + JSZip de forma PEREZOSA (solo al entrar a "Masivo").
// - Genera los estados de cuenta COMPLETOS (pagos + cabecera SECOP2) en el
//   navegador con openpyxl, preservando el formato. No usa el backend.
// Requiere que XLSX (SheetJS) ya esté cargado en la plataforma.
// ============================================================================

(function () {
  'use strict';

  // ---------- Submenú (pestañas) ----------
  document.addEventListener('DOMContentLoaded', function () {
    const tabs = document.querySelectorAll('.ec-tab');
    if (!tabs.length) return;

    tabs.forEach(tab => {
      tab.addEventListener('click', () => {
        const destino = tab.getAttribute('data-panel');
        document.querySelectorAll('.ec-tab').forEach(t => t.classList.remove('activo'));
        document.querySelectorAll('.ec-panel').forEach(p => p.classList.remove('activo'));
        tab.classList.add('activo');
        const panel = document.getElementById(destino);
        if (panel) panel.classList.add('activo');
        if (destino === 'ec-masivo') prepararMotor();
      });
    });

    const btn = document.getElementById('btn-ec-masivo');
    if (btn) btn.addEventListener('click', generarMasivo);
  });

  // ---------- Carga perezosa de dependencias ----------
  let pyodideReady = null;

  function cargarScript(src) {
    return new Promise((res, rej) => {
      if ([...document.scripts].some(s => s.src === src)) return res();
      const s = document.createElement('script');
      s.src = src; s.onload = res; s.onerror = () => rej(new Error('No cargó ' + src));
      document.head.appendChild(s);
    });
  }

  function estadoMotor(txt, ok) {
    const el = document.getElementById('ec-masivo-motor');
    if (!el) return;
    el.textContent = txt;
    el.style.background = ok ? '#f0fff4' : '#fffaf0';
    el.style.borderColor = ok ? '#68d391' : '#f6ad55';
  }

  function prepararMotor() {
    if (pyodideReady) return pyodideReady;
    estadoMotor('⏳ Preparando el motor (Python en el navegador)… unos segundos la primera vez.', false);

    pyodideReady = (async () => {
      await cargarScript('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js');
      await cargarScript('https://cdn.jsdelivr.net/pyodide/v0.26.4/full/pyodide.js');
      const py = await loadPyodide();
      await py.loadPackage('micropip');
      await py.runPythonAsync('import micropip\nawait micropip.install("openpyxl")');
      await py.runPythonAsync(PYTHON_FILL);
      estadoMotor('✅ Motor listo. Ya puedes generar.', true);
      return py;
    })().catch(err => {
      estadoMotor('❌ Error cargando el motor: ' + err, false);
      pyodideReady = null;
      throw err;
    });

    return pyodideReady;
  }

  // ---------- Utilidades ----------
  function ab2b64(buf) {
    let bin = '', bytes = new Uint8Array(buf), chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    return btoa(bin);
  }
  function b642u8(b64) {
    const bin = atob(b64), len = bin.length, bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }
  function prog(t) { const el = document.getElementById('ec-masivo-progreso'); if (el) el.textContent = t; }
  function alerta(t) { const el = document.getElementById('ec-masivo-alerta'); if (el) el.textContent = t || ''; }
  function metrica(num, lbl, clase) {
    return '<div class="ec-metrica ' + (clase || '') + '"><div class="num">' + num + '</div><div class="lbl">' + lbl + '</div></div>';
  }

  const _normJS = (s) => String(s).trim().toLowerCase()
    .normalize('NFKD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[º°]/g, '').replace(/\s+/g, ' ');
  const getter = (row) => {
    const map = {};
    for (const k in row) map[_normJS(k)] = row[k];
    return (...names) => {
      for (const n of names) { const v = map[_normJS(n)]; if (v !== undefined && v !== '') return v; }
      return '';
    };
  };

  async function leerExcel(file) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true });
    const hoja = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(hoja, { raw: true, defval: '' });
  }

  // ---------- Proceso principal ----------
  async function generarMasivo() {
    alerta('');
    const res = document.getElementById('ec-masivo-resultados');
    if (res) res.style.display = 'none';

    const fPlantilla = document.getElementById('input-plantilla-ec-masivo').files[0];
    const fHistorico = document.getElementById('input-historico-ec-masivo').files[0];
    const fSecop = document.getElementById('input-secop-ec-masivo').files[0];
    const contratos = document.getElementById('input-contratos-ec-masivo').value
      .split(/[\n,;]+/).map(c => c.trim()).filter(Boolean);

    if (!fPlantilla || !fHistorico || !fSecop) { alerta('Faltan archivos: plantilla, histórico y SECOP2.'); return; }
    if (!contratos.length) { alerta('Escribe al menos un contrato (uno por línea).'); return; }

    const btn = document.getElementById('btn-ec-masivo');
    btn.disabled = true;

    try {
      const py = await prepararMotor();
      const llenar = py.globals.get('llenar');

      prog('Leyendo plantilla…');
      const plantillaB64 = ab2b64(await fPlantilla.arrayBuffer());

      prog('Leyendo histórico (puede tardar)…');
      const hist = await leerExcel(fHistorico);

      prog('Leyendo SECOP2…');
      const secop = await leerExcel(fSecop);

      const idx = {};
      for (const r of secop) {
        const c = String(getter(r)('CONTRATO') || '').trim();
        if (!c) continue;
        (idx[c] = idx[c] || []).push(r);
      }
      const buscarSecop = (contrato, cedula) => {
        const filas = idx[String(contrato).trim()] || [];
        if (!filas.length) return null;
        if (filas.length === 1) return filas[0];
        if (cedula) {
          const cp = String(cedula).split('.')[0];
          const hit = filas.find(f => String(getter(f)('NUMERO DE IDENTIFICACION')).split('.')[0] === cp);
          if (hit) return hit;
        }
        return filas[0];
      };
      const datosDe = (r) => {
        const g = getter(r);
        return {
          nombre: g('NOMBRE DEL CONTRATISTA'),
          cedula: g('NUMERO DE IDENTIFICACION'),
          valor_inicial: g('VALOR INICIAL DEL CONTRATO', 'VALOR INICIAL'),
          valor_final: g('VALOR FINAL DEL CONTRATO', 'VALOR FINAL'),
          fecha_inicial: g('FECHA INICIAL DE CONTRATO', 'FECHA INICIAL'),
          fecha_term_inicial: g('FECHA DE TERMINACION INICIAL', 'FECHA TERMINACION INICIAL'),
          fecha_term_final: g('FECHA DE TERMINACION FINAL', 'FECHA TERMINACION FINAL'),
        };
      };

      const zip = new JSZip();
      const ok = [], sinPagos = [], sinSecop = [];

      for (let n = 0; n < contratos.length; n++) {
        const c = contratos[n];
        prog(`Procesando ${n + 1} de ${contratos.length}: ${c}`);

        const pagos = hist.filter(r => String(r['Referencia'] || '').includes(c));
        if (!pagos.length) { sinPagos.push(c); continue; }

        const srow = buscarSecop(c, getter(pagos[0])('Nº identificación', 'Numero identificacion', 'Cedula'));
        const datos = srow ? datosDe(srow) : null;
        if (!datos) sinSecop.push(c);

        const b64 = llenar(plantillaB64, JSON.stringify(pagos), String(c), datos ? JSON.stringify(datos) : '');
        const seguro = String(c).replace(/[\\/:*?"<>|]/g, '_');
        zip.file('Estado_de_Cuenta_' + seguro + '.xlsx', b642u8(b64));
        ok.push(c);
      }

      if (!ok.length) { alerta('No se generó ningún archivo. Revisa contratos e histórico.'); prog(''); return; }

      prog('Comprimiendo ' + ok.length + ' archivos…');
      const blob = await zip.generateAsync({ type: 'blob' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = 'Estados_de_Cuenta.zip'; a.click();
      setTimeout(() => URL.revokeObjectURL(url), 120000);

      const met = document.getElementById('ec-masivo-metricas');
      if (met) met.innerHTML =
        metrica(contratos.length, 'Solicitados') +
        metrica(ok.length, 'Generados', 'verde') +
        metrica(sinPagos.length, 'Sin pagos', 'ocre') +
        metrica(sinSecop.length, 'Sin SECOP2', sinSecop.length ? 'ocre' : 'verde');

      const det = document.getElementById('ec-masivo-detalle');
      if (det) {
        let html = '';
        if (sinPagos.length) html += '<p><strong>Sin pagos:</strong> ' + sinPagos.join(', ') + '</p>';
        if (sinSecop.length) html += '<p><strong>Sin match en SECOP2 (cabecera parcial):</strong> ' + sinSecop.join(', ') + '</p>';
        det.innerHTML = html;
      }
      if (res) res.style.display = 'block';
      prog('✅ Listo. Se descargó Estados_de_Cuenta.zip');

    } catch (err) {
      alerta('Error: ' + err);
      console.error(err);
    } finally {
      btn.disabled = false;
    }
  }

  // ---------- Código Python (openpyxl) que corre en Pyodide ----------
  const PYTHON_FILL = `
import io, json, base64, unicodedata
from datetime import datetime, timedelta
import openpyxl
from openpyxl.utils import range_boundaries, column_index_from_string

def _norm(s):
    s=str(s).strip().lower().replace("º","").replace("°","")
    s=unicodedata.normalize("NFKD",s)
    s="".join(c for c in s if not unicodedata.combining(c))
    return " ".join(s.split())
def _pn(fila):
    return {_norm(k):v for k,v in fila.items()}
def _g(pn,*nombres):
    for n in nombres:
        v=pn.get(_norm(n))
        if v is not None: return v
    return ""
def _limpiar(d):
    x=str(d).split(".")[0]
    return "" if x in ("nan","None","NaT") else x
def _num(x):
    try: return float(x)
    except: return 0.0
def _txt(v):
    return "" if v is None or str(v).strip().lower() in ("","nan","none","nat") else str(v)
def _fecha(v):
    if v is None: return None
    if isinstance(v,datetime): return v
    s=str(v).strip()
    if s in ("","nan","None","NaT"): return None
    s=s.replace("Z","").split(".")[0]
    for f in ("%Y-%m-%dT%H:%M:%S","%Y-%m-%d %H:%M:%S","%Y-%m-%d","%d/%m/%Y"):
        try: return datetime.strptime(s,f)
        except: continue
    try: return datetime.fromisoformat(s)
    except: return None
def _mm(ws):
    m={}
    for r in ws.merged_cells.ranges:
        a,b,c,d=range_boundaries(r.coord)
        for rr in range(b,d+1):
            for cc in range(a,c+1): m[(rr,cc)]=(b,a)
    return m
def _wrc(ws,r,c,v,money=False,m=None):
    dd=m.get((r,c),(r,c)); cel=ws.cell(row=dd[0],column=dd[1]); cel.value=v
    if money: cel.number_format='"$"#,##0'
def _wc(ws,coord,v,money=False,m=None):
    cl="".join(filter(str.isalpha,coord)); fn=int("".join(filter(str.isdigit,coord)))
    _wrc(ws,fn,column_index_from_string(cl),v,money,m)
def _dm(ws,fila,ci=2,cf=10):
    q=[]
    for r in list(ws.merged_cells.ranges):
        a,b,c,d=range_boundaries(r.coord)
        if b<=fila<=d and a<=cf and c>=ci: q.append(r.coord)
    for x in q: ws.unmerge_cells(x)
def _cab(ws,m,dc,nombre_hist=""):
    vi=_num(dc.get("valor_inicial")); vf=_num(dc.get("valor_final")); ad=vf-vi
    fi=_fecha(dc.get("fecha_inicial")); fti=_fecha(dc.get("fecha_term_inicial")); ftf=_fecha(dc.get("fecha_term_final"))
    nom=_txt(dc.get("nombre")) or _txt(nombre_hist)
    if nom: _wc(ws,"D6",nom,m=m)
    if _limpiar(dc.get("cedula","")): _wc(ws,"H6",_limpiar(dc.get("cedula","")),m=m)
    _wc(ws,"D7",vi,money=True,m=m)
    if fi: _wc(ws,"D8",fi,m=m)
    if fti:_wc(ws,"H8",fti,m=m)
    if ad>0:
        _wc(ws,"D9",ad,money=True,m=m)
        if fti:_wc(ws,"D10",fti+timedelta(days=1),m=m)
        if ftf:_wc(ws,"H10",ftf,m=m)
    return vf if vf>0 else vi

def _sinac(s):
    return "".join(c for c in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(c)).upper()
def _buscar(ws, texto, col=3, desde=1, hasta=None):
    t=texto.strip().upper(); hasta=hasta or ws.max_row
    for r in range(desde,hasta+1):
        v=ws.cell(row=r,column=col).value
        if v and t in str(v).strip().upper(): return r
    return None
def _detectar_cesion_masivo(pagos, dc):
    pagos=sorted(pagos, key=lambda p:_fecha(_g(_pn(p),"Fecha de pago")) or datetime.max)
    por_bp={}; orden=[]
    for p in pagos:
        pn=_pn(p); bp=_limpiar(_g(pn,"Proveedor"))
        if not bp: continue
        if bp not in por_bp: por_bp[bp]=[]; orden.append(bp)
        por_bp[bp].append(p)
    if len(orden)<2: return None
    bpc=orden[0]
    vi=_num(dc.get("valor_inicial")) if dc else 0.0
    if not vi:
        for p in pagos:
            v=_num(_g(_pn(p),"VALOR FINAL DEL CONTRATO","VALOR INICIAL DEL CONTRATO"))
            if v>0: vi=v; break
    pc=por_bp[bpc]; suma=sum(_num(_g(_pn(p),"Valor Bruto")) for p in pc)
    vces=vi-suma; ces=[]
    for bp in orden[1:]:
        pgs=por_bp[bp]; pn0=_pn(pgs[0])
        ces.append({"bp":bp,"nombre":_txt(_g(pn0,"Nombre")),"pagos":pgs,"valor_cesion":vces})
    return {"bp_cedente":bpc,"pagos_cedente":pc,"valor_cesion":vces,"cesionarios":ces}
def _llenar_bloque_cesion(ws, c, contrato, dc, fila_titulo):
    m=_mm(ws)
    f_cto=_buscar(ws,"CTO Y VIG",col=3,desde=fila_titulo)
    if not f_cto: return False
    _wrc(ws,f_cto,4,contrato,m=m); _wrc(ws,f_cto,8,c["bp"],m=m)
    _wrc(ws,f_cto+1,4,c["nombre"],m=m)
    _wrc(ws,f_cto+2,4,c["valor_cesion"],money=True,m=m)
    if c["pagos"]:
        pn0=_pn(c["pagos"][0]); _wrc(ws,f_cto+2,8,_limpiar(_g(pn0,"Numero RP")),m=m)
    if dc:
        fti=_fecha(dc.get("fecha_term_inicial")); ftf=_fecha(dc.get("fecha_term_final"))
        if fti: _wrc(ws,f_cto+3,4,fti+timedelta(days=1),m=m)
        if ftf: _wrc(ws,f_cto+3,8,ftf,m=m)
    f_hdr=_buscar(ws,"PERIODO",col=3,desde=f_cto)
    if not f_hdr: return False
    f_pago=f_hdr+1; f_total=_buscar(ws,"TOTAL",col=2,desde=f_pago)
    slots=(f_total-f_pago) if f_total else 6
    for i in range(min(len(c["pagos"]),slots)): _dm(ws,f_pago+i)
    m=_mm(ws); saldo=c["valor_cesion"]; suma=0
    for i,p in enumerate(c["pagos"][:slots]):
        pn=_pn(p); mo=_num(_g(pn,"Valor Bruto")); saldo-=mo; suma+=mo; r=f_pago+i
        fp=_fecha(_g(pn,"Fecha de pago")) or _txt(_g(pn,"Fecha de pago"))
        _wrc(ws,r,2,i+1,m=m); _wrc(ws,r,3,_txt(_g(pn,"Texto cabecera documento")),m=m)
        _wrc(ws,r,4,mo,money=True,m=m); _wrc(ws,r,5,saldo,money=True,m=m)
        _wrc(ws,r,6,_limpiar(_g(pn,"Doc.compensación","Doc compensacion")),m=m)
        _wrc(ws,r,7,fp,m=m); _wrc(ws,r,8,_limpiar(_g(pn,"Numero RP")),m=m)
        _wrc(ws,r,9,_limpiar(_g(pn,"CDP Externo")),m=m); _wrc(ws,r,10,_limpiar(_g(pn,"CRP Externo")),m=m)
    if f_total: _wrc(ws,f_total,4,suma,money=True,m=m); _wrc(ws,f_total,5,saldo,money=True,m=m)
    return True

def llenar(plantilla_b64, pagos_json, contrato, datos_json):
    pl=base64.b64decode(plantilla_b64)
    pagos=json.loads(pagos_json)
    dc=json.loads(datos_json) if datos_json else None
    wb=openpyxl.load_workbook(io.BytesIO(pl)); ws=wb.active
    ini=17
    for f in range(ini,ini+len(pagos)): _dm(ws,f)
    m=_mm(ws)
    pn0=_pn(pagos[0])
    _wc(ws,"D5",contrato,m=m)
    _wc(ws,"H5",_limpiar(_g(pn0,"Proveedor")),m=m)
    _wc(ws,"H7",_limpiar(_g(pn0,"Numero RP")),m=m)
    if dc: val=_cab(ws,m,dc,_g(pn0,"Nombre"))
    else:
        val=_num(_g(pn0,"VALOR FINAL DEL CONTRATO"))
        _wc(ws,"D6",_txt(_g(pn0,"Nombre")),m=m)
        _wc(ws,"D7",val,money=True,m=m)
        _wc(ws,"H6",_limpiar(_g(pn0,"Nº identificación")),m=m)
    fa=ini; saldo=val
    for i,p in enumerate(pagos,1):
        pn=_pn(p)
        mo=_num(_g(pn,"Valor Bruto")); saldo-=mo
        fp=_fecha(_g(pn,"Fecha de pago")) or _txt(_g(pn,"Fecha de pago"))
        _wrc(ws,fa,2,i,m=m); _wrc(ws,fa,3,_txt(_g(pn,"Texto cabecera documento")),m=m)
        _wrc(ws,fa,4,mo,money=True,m=m); _wrc(ws,fa,5,saldo,money=True,m=m)
        _wrc(ws,fa,6,_limpiar(_g(pn,"Doc.compensación","Doc compensacion")),m=m)
        _wrc(ws,fa,7,fp,m=m); _wrc(ws,fa,8,_limpiar(_g(pn,"Numero RP")),m=m)
        _wrc(ws,fa,9,_limpiar(_g(pn,"CDP Externo")),m=m); _wrc(ws,fa,10,_limpiar(_g(pn,"CRP Externo")),m=m)
        fa+=1
    # --- CESIONES (deteccion automatica por BP en el historico) ---
    try:
        _ces=_detectar_cesion_masivo(pagos, dc)
        if _ces and _ces["cesionarios"]:
            _fc=_buscar(ws,"CESION",col=2,desde=1) or _buscar(ws,chr(67)+chr(69)+chr(83)+chr(73)+chr(211)+chr(78),col=2,desde=1)
            if _fc:
                _llenar_bloque_cesion(ws, _ces["cesionarios"][0], contrato, dc, _fc)
    except Exception as _e:
        import traceback as _tb
        print("ERROR CESION:", _e)
        _tb.print_exc()
    out=io.BytesIO(); wb.save(out)
    return base64.b64encode(out.getvalue()).decode()
`;

})();