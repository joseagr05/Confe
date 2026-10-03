#!/usr/bin/env node
/*
 * servidor_tiendas.js — Servidor de "Asignación de tiendas" (Conferencia JAS)
 *
 * Sirve la misma app (asignacion_tiendas.html) a varios dispositivos a la vez:
 *   - Edición en tiempo real: el servidor guarda el estado (estado_tiendas.json),
 *     valida cada operación y la difunde al instante por Server-Sent Events.
 *   - Enlace de SOLO LECTURA: ve todo en vivo pero no puede modificar nada
 *     (el servidor rechaza las escrituras, no es solo ocultar botones).
 *   - Sin dependencias: solo Node (>= 16). No necesita internet para la red local.
 *
 * Uso (red local del evento):
 *     node servidor_tiendas.js
 *     -> imprime dos enlaces: uno de editor y uno de solo lectura
 *
 * Uso (nube): sube servidor_tiendas.js + asignacion_tiendas.html y define las
 *     variables de entorno TOKEN_EDITOR y TOKEN_VER (opcional; si no existen se
 *     generan y quedan guardados en servidor_config.json). El puerto lo asigna
 *     la plataforma en PORT.
 *
 * Opciones por línea de comandos (o variables de entorno):
 *     --puerto N      | PORT             (por defecto 8787)
 *     --html RUTA     | HTML             (app a servir)
 *     --estado RUTA   | ESTADO           (archivo de estado)
 *     --roster RUTA   | ROSTER           (archivo con la lista de participantes)
 *     --config RUTA   | CONFIG           (archivo de configuración/tokens)
 *     --editor TOKEN  | TOKEN_EDITOR     (código de editor)
 *     --ver TOKEN     | TOKEN_VER        (código de solo lectura)
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');

const DIR = __dirname;

function arg(nombre, pred){
  const i = process.argv.indexOf(nombre);
  if(i !== -1 && process.argv[i + 1] !== undefined) return process.argv[i + 1];
  return pred;
}

const puertoRaw = parseInt(arg('--puerto', process.env.PORT || '8787'), 10);
const PUERTO = isNaN(puertoRaw) ? 8787 : puertoRaw; // 0 = que el sistema elija uno libre
const RUTA_HTML = path.resolve(arg('--html', process.env.HTML || path.join(DIR, 'asignacion_tiendas.html')));
const RUTA_ESTADO = path.resolve(arg('--estado', process.env.ESTADO || path.join(DIR, 'estado_tiendas.json')));
const RUTA_ROSTER = path.resolve(arg('--roster', process.env.ROSTER || path.join(DIR, 'roster.json')));
const RUTA_CONFIG = path.resolve(arg('--config', process.env.CONFIG || path.join(DIR, 'servidor_config.json')));

// ---------------------------------------------------------------- tokens
let cfg = {};
try{ cfg = JSON.parse(fs.readFileSync(RUTA_CONFIG, 'utf8')) || {}; }catch(e){ cfg = {}; }
const TOK_EDITOR = process.env.TOKEN_EDITOR || arg('--editor', cfg.tokenEditor || '') || crypto.randomBytes(12).toString('hex');
const TOK_VER = process.env.TOKEN_VER || arg('--ver', cfg.tokenVer || '') || crypto.randomBytes(12).toString('hex');
if(!cfg.tokenEditor || !cfg.tokenVer){
  cfg = { tokenEditor: TOK_EDITOR, tokenVer: TOK_VER };
  try{ fs.writeFileSync(RUTA_CONFIG, JSON.stringify(cfg, null, 2)); }catch(e){}
}

// ---------------------------------------------------------------- lógica de la app
// El servidor ejecuta EL MISMO código que el navegador (extraído del HTML) con
// un DOM falso: así no hay dos implementaciones de las reglas de asignación.
function cargarLogica(){
  if(!fs.existsSync(RUTA_HTML)) throw new Error('No se encontró la app: ' + RUTA_HTML);
  const html = fs.readFileSync(RUTA_HTML, 'utf8');
  const bloques = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => m[1]);
  if(!bloques.length) throw new Error('La app no tiene script inline.');
  const src = bloques[bloques.length - 1];

  const elements = {};
  function makeEl(id){
    return {
      id, textContent:'', innerHTML:'', value:'', disabled:false,
      style:{}, dataset:{}, offsetWidth:0, src:'', href:'', download:'',
      classList:{ add(){}, remove(){}, toggle(){ return false; }, contains(){ return false; } },
      addEventListener(){}, dispatchEvent(){}, click(){},
      querySelectorAll(){ return []; }, querySelector(){ return null; },
      appendChild(){}, removeChild(){},
      getAttribute(){ return null; }, setAttribute(){}
    };
  }
  global.document = {
    getElementById(id){ if(!elements[id]) elements[id] = makeEl(id); return elements[id]; },
    createElement(){ return makeEl('nuevo'); },
    querySelectorAll(){ return []; },
    querySelector(){ return null; },
    head:{ appendChild(){} },
    body:{ appendChild(){}, removeChild(){}, classList:{ add(){}, remove(){} } },
    addEventListener(){}
  };
  const store = {};
  global.localStorage = {
    getItem(k){ return (k in store) ? store[k] : null; },
    setItem(k, v){ store[k] = String(v); },
    removeItem(k){ delete store[k]; }
  };
  global.window = { print(){} };
  global.XLSX = { read(){}, utils:{} };
  // OJO: no reemplazar global.URL — el servidor necesita el constructor para leer
  // los tokens. Solo se agregan los métodos que la app usa al exportar archivos.
  if(typeof global.URL === 'undefined'){
    global.URL = { createObjectURL(){ return 'blob:x'; }, revokeObjectURL(){} };
  }else{
    if(!global.URL.createObjectURL) global.URL.createObjectURL = () => 'blob:x';
    if(!global.URL.revokeObjectURL) global.URL.revokeObjectURL = () => {};
  }
  global.Blob = function(){};
  global.Event = class Event { constructor(t){ this.type = t; } };
  global.FileReader = function(){};
  global.confirm = () => true; // el navegador ya confirmó antes de enviar la operación

  // El bloque final del HTML expone la API interna al servidor.
  const API_INTERNA = `
;global.__TIENDAS__ = {
  get estado(){ return estado; },
  set estado(v){ estado = v; },
  get personas(){ return personas; },
  get sesiones(){ return sesiones; },
  get sesionActual(){ return sesionActual; },
  set sesionActual(v){ sesionActual = v; },
  setRoster(r){
    personas = r.personas; sesiones = r.sesiones.slice();
    sesionActual = sesiones[0] || sesionActual; sel = null;
    sesiones.forEach(asegurarSesion);
    reconstruirIndice();
  },
  aplicarFiltros(f){ if(!f) return; filtroGenero = f.g || 'H'; filtroEstaca = f.e || ''; filtroTexto = f.t || ''; },
  reconstruirIndice, asegurarSesion,
  asignar, liberar, tachar, destachar,
  nuevaTienda, eliminarTienda, rellenar, autoAsignar,
  restablecerNoche, configurarSitio
};`;

  eval(src + '\n' + API_INTERNA);
  const T = global.__TIENDAS__;
  if(!T) throw new Error('No se pudo cargar la API interna de la app.');
  return {
    T,
    limpiarToast(){ global.document.getElementById('toast').textContent = ''; },
    textoToast(){ return global.document.getElementById('toast').textContent || ''; }
  };
}

const { T, limpiarToast, textoToast } = cargarLogica();

// ---------------------------------------------------------------- estado y lista
let version = 0;
const runId = crypto.randomBytes(4).toString('hex');
let roster = null;

function guardarAtomico(ruta, obj){
  try{
    const tmp = ruta + '.tmp';
    fs.writeFileSync(tmp, JSON.stringify(obj));
    fs.renameSync(tmp, ruta);
  }catch(e){ console.error('[aviso] no se pudo guardar ' + ruta + ': ' + e.message); }
}
function guardarEstado(){ guardarAtomico(RUTA_ESTADO, { v: 1, version: version, estado: T.estado }); }

function estadoValido(d){
  return d && d.v === 1 && Array.isArray(d.lugares) && d.sesiones && typeof d.sesiones === 'object';
}
function cargarInicial(){
  try{
    if(fs.existsSync(RUTA_ESTADO)){
      const d = JSON.parse(fs.readFileSync(RUTA_ESTADO, 'utf8'));
      if(d && estadoValido(d.estado)){
        T.estado = d.estado;
        version = parseInt(d.version, 10) || 0;
      }
    }
  }catch(e){ console.error('[aviso] estado ilegible, se empieza de cero: ' + e.message); }
  try{
    if(fs.existsSync(RUTA_ROSTER)){
      const r = JSON.parse(fs.readFileSync(RUTA_ROSTER, 'utf8'));
      const err = validarRoster(r);
      if(!err){ roster = r; T.setRoster(r); }
      else console.error('[aviso] roster.json inválido (' + err + '), se ignora.');
    }
  }catch(e){ console.error('[aviso] roster ilegible, se ignora: ' + e.message); }
  T.sesionActual = (roster && roster.sesiones[0]) || T.sesionActual;
  T.reconstruirIndice();
}
function validarRoster(r){
  if(!r || typeof r !== 'object') return 'La lista no es un objeto válido.';
  if(!Array.isArray(r.sesiones) || !r.sesiones.length) return 'La lista no tiene sesiones.';
  if(!r.personas || typeof r.personas !== 'object') return 'La lista no tiene personas.';
  let total = 0;
  for(const s of r.sesiones){
    const arr = r.personas[s];
    if(!Array.isArray(arr)) return 'La sesión ' + s + ' no tiene personas.';
    for(const p of arr){
      if(!p || typeof p.k !== 'string' || typeof p.n !== 'string') return 'Persona mal formada en ' + s + '.';
      total++;
    }
  }
  if(total > 100000) return 'La lista es demasiado grande.';
  return null;
}

// ---------------------------------------------------------------- operaciones
function ejecutarOp(p){
  if(!p || typeof p !== 'object') return { ok:false, error:'Petición inválida.' };
  const tipos = ['asignar','liberar','tachar','destachar','nuevaTienda','eliminarTienda','rellenar','autoAsignar','restablecer','sitio'];
  if(tipos.indexOf(p.tipo) === -1) return { ok:false, error:'Operación desconocida.' };
  if(!Array.isArray(p.args)) return { ok:false, error:'Argumentos inválidos.' };
  if(!T.sesiones.length) return { ok:false, error:'Aún no se ha cargado la lista de participantes.' };
  if(p.sesion && T.sesiones.indexOf(p.sesion) === -1) return { ok:false, error:'Sesión no válida.' };

  T.sesionActual = p.sesion || T.sesiones[0];
  T.aplicarFiltros(p.filtros);
  T.reconstruirIndice();

  const a = p.args;
  if(p.tipo === 'sitio' && a[1] && typeof a[1] === 'object'){
    const c = a[1];
    if(typeof c.nombre === 'string') c.nombre = c.nombre.slice(0, 60);
    if(c.genero != null && c.genero !== 'H' && c.genero !== 'M') delete c.genero;
    if(c.maxTiendas != null) c.maxTiendas = Math.max(0, Math.min(1000, +c.maxTiendas || 0));
  }

  limpiarToast();
  let ok = true, error = '';
  try{
    switch(p.tipo){
      case 'asignar':        ok = !!T.asignar(a[0], a[1], a[2]); break;
      case 'liberar':        ok = !!T.liberar(a[0], a[1], a[2]); break;
      case 'tachar':         ok = !!T.tachar(a[0]); break;
      case 'destachar':      ok = !!T.destachar(a[0]); break;
      case 'nuevaTienda':    ok = T.nuevaTienda(a[0]) !== null; break;
      case 'eliminarTienda': T.eliminarTienda(a[0], a[1]); ok = true; break;
      case 'rellenar':       T.rellenar(a[0], a[1]); ok = true; break;
      case 'autoAsignar':    T.autoAsignar(); ok = true; break;
      case 'restablecer':    T.restablecerNoche(); ok = true; break;
      case 'sitio':          ok = !!T.configurarSitio(a[0], a[1] || {}); break;
      default:               ok = false; error = 'Operación desconocida.';
    }
  }catch(err){
    ok = false;
    error = 'Error al aplicar la operación: ' + (err && err.message ? err.message : String(err));
  }
  if(!ok && !error) error = textoToast() || 'La operación no se pudo aplicar en el servidor.';
  return { ok, error };
}

// ---------------------------------------------------------------- SSE (tiempo real)
const streams = new Set();
function broadcast(obj){
  const data = 'data: ' + JSON.stringify(obj) + '\n\n';
  for(const res of streams){ try{ res.write(data); }catch(e){} }
}
function broadcastPresencia(){ broadcast({ tipo:'presencia', conectados: streams.size, runId: runId }); }
function abrirStream(res){
  res.writeHead(200, {
    'Content-Type': 'text/event-stream; charset=utf-8',
    'Cache-Control': 'no-cache, no-transform',
    'Connection': 'keep-alive',
    'Access-Control-Allow-Origin': '*'
  });
  res.write(': conectado\n\n');
  streams.add(res);
  res.on('close', () => { streams.delete(res); broadcastPresencia(); });
  broadcastPresencia();
}
const latido = setInterval(() => {
  for(const res of streams){ try{ res.write(': hb\n\n'); }catch(e){ streams.delete(res); } }
}, 25000);
if(latido.unref) latido.unref();

// ---------------------------------------------------------------- http
function enviarJSON(res, codigo, obj){
  res.writeHead(codigo, {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Cache-Control': 'no-store'
  });
  res.end(JSON.stringify(obj));
}
function leerToken(u){
  try{ return new URL(u, 'http://localhost').searchParams.get('acceso') || ''; }catch(e){ return ''; }
}
function rolDe(u){
  const t = leerToken(u);
  if(t && t === TOK_EDITOR) return 'editor';
  if(t && t === TOK_VER) return 'lector';
  return null;
}
function leerCuerpo(req, max, cb){
  const trozos = [];
  let n = 0;
  req.on('data', c => {
    n += c.length;
    if(n > max){ req.destroy(); return; }
    trozos.push(c);
  });
  req.on('end', () => cb(Buffer.concat(trozos).toString('utf8')));
  req.on('error', () => {});
}
function baseEstado(){
  return { version: version, runId: runId, conectados: streams.size, estado: T.estado };
}

const servidor = http.createServer((req, res) => {
  const u = req.url || '/';
  const ruta = u.split('?')[0];

  if(req.method === 'OPTIONS'){
    res.writeHead(204, {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
    return res.end();
  }

  // salud: la app lo usa para descubrir si está en modo red
  if(ruta === '/api/salud') return enviarJSON(res, 200, { ok:true, app:'asignacion_tiendas' });

  const rol = rolDe(u);

  if(ruta === '/api/estado'){
    if(!rol) return enviarJSON(res, 401, { ok:false, error:'Código de acceso no válido.' });
    return enviarJSON(res, 200, Object.assign({ ok:true, rol: rol, roster: roster }, baseEstado()));
  }

  if(ruta === '/api/stream'){
    if(!rol) return enviarJSON(res, 401, { ok:false, error:'Código de acceso no válido.' });
    return abrirStream(res);
  }

  if(ruta === '/api/operacion'){
    if(req.method !== 'POST') return enviarJSON(res, 405, { ok:false, error:'Método no permitido.' });
    if(rol !== 'editor'){
      // los lectores (o quien intente hackear desde la consola) no escriben jamás
      return enviarJSON(res, 403, Object.assign({ ok:false, error:'Este enlace es de solo lectura.' }, baseEstado()));
    }
    return leerCuerpo(req, 1024 * 1024, raw => {
      let p;
      try{ p = JSON.parse(raw); }catch(e){ return enviarJSON(res, 400, { ok:false, error:'JSON inválido.' }); }
      const r = ejecutarOp(p);
      if(r.ok){
        version++;
        guardarEstado();
        broadcast(Object.assign({ tipo:'estado' }, baseEstado()));
      }
      enviarJSON(res, 200, Object.assign({ ok: r.ok, error: r.error }, baseEstado()));
    });
  }

  if(ruta === '/api/roster'){
    if(req.method !== 'POST') return enviarJSON(res, 405, { ok:false, error:'Método no permitido.' });
    if(rol !== 'editor') return enviarJSON(res, 403, { ok:false, error:'Este enlace es de solo lectura.' });
    return leerCuerpo(req, 8 * 1024 * 1024, raw => {
      let r;
      try{ r = JSON.parse(raw); }catch(e){ return enviarJSON(res, 400, { ok:false, error:'JSON inválido.' }); }
      const err = validarRoster(r);
      if(err) return enviarJSON(res, 400, { ok:false, error: err });
      roster = r;
      T.setRoster(r);
      guardarAtomico(RUTA_ROSTER, r);
      broadcast({ tipo:'roster', roster: roster, conectados: streams.size, runId: runId });
      enviarJSON(res, 200, { ok:true, sesiones: roster.sesiones });
    });
  }

  if(ruta === '/api/copia'){
    if(rol !== 'editor') return enviarJSON(res, 403, { ok:false, error:'Requiere enlace de editor.' });
    if(req.method === 'GET'){
      return enviarJSON(res, 200, { ok:true, version: version, runId: runId, estado: T.estado, roster: roster });
    }
    if(req.method === 'POST'){
      return leerCuerpo(req, 8 * 1024 * 1024, raw => {
        let d;
        try{ d = JSON.parse(raw); }catch(e){ return enviarJSON(res, 400, { ok:false, error:'JSON inválido.' }); }
        if(!estadoValido(d.estado)) return enviarJSON(res, 400, { ok:false, error:'Estado inválido.' });
        T.estado = d.estado;
        let rosterRestaurado = false;
        if(d.roster && !validarRoster(d.roster)){ // la copia también repone la lista
          roster = d.roster;
          T.setRoster(roster);
          guardarAtomico(RUTA_ROSTER, roster);
          rosterRestaurado = true;
        }
        if(roster) T.sesiones.forEach(s => T.asegurarSesion(s));
        T.reconstruirIndice();
        version++;
        guardarEstado();
        if(rosterRestaurado) broadcast({ tipo:'roster', roster: roster, conectados: streams.size, runId: runId });
        broadcast(Object.assign({ tipo:'estado' }, baseEstado()));
        enviarJSON(res, 200, Object.assign({ ok:true }, baseEstado()));
      });
    }
    return enviarJSON(res, 405, { ok:false, error:'Método no permitido.' });
  }

  // la app
  if(req.method === 'GET' && (ruta === '/' || ruta === '/asignacion_tiendas.html')){
    try{
      const cuerpo = fs.readFileSync(RUTA_HTML);
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-cache' });
      return res.end(cuerpo);
    }catch(e){
      return enviarJSON(res, 500, { ok:false, error:'No se pudo leer la app: ' + e.message });
    }
  }
  if(ruta === '/favicon.ico'){ res.writeHead(204); return res.end(); }

  enviarJSON(res, 404, { ok:false, error:'No encontrado.' });
});

function ipsLAN(){
  const out = [];
  const ifs = os.networkInterfaces();
  for(const nombre of Object.keys(ifs)){
    for(const i of (ifs[nombre] || [])){
      if((i.family === 'IPv4' || i.family === 4) && !i.internal) out.push(i.address);
    }
  }
  return out;
}

cargarInicial();

servidor.listen(PUERTO, '0.0.0.0', () => {
  const puertoReal = servidor.address().port;
  console.log('==================================================');
  console.log(' Asignación de tiendas · servidor de red listo');
  console.log('==================================================');
  console.log('PUERTO=' + puertoReal);
  console.log(' App   : ' + RUTA_HTML);
  console.log(' Estado: ' + RUTA_ESTADO);
  console.log(' Lista : ' + (roster ? RUTA_ROSTER + ' (cargada)' : RUTA_ROSTER + ' (aún sin cargar; el primer editor sube el Excel)'));
  const ips = ipsLAN();
  if(ips.length){
    console.log('');
    console.log(' Enlaces para los dispositivos (misma Wi-Fi, sin internet):');
    for(const ip of ips){
      console.log('   Editor     http://' + ip + ':' + puertoReal + '/?acceso=' + TOK_EDITOR);
      console.log('   Solo lect. http://' + ip + ':' + puertoReal + '/?acceso=' + TOK_VER);
    }
  }
  console.log('');
  console.log(' En este equipo: http://localhost:' + puertoReal + '/?acceso=' + TOK_EDITOR);
  console.log(' Códigos -> editor: ' + TOK_EDITOR + ' | solo lectura: ' + TOK_VER);
  console.log(' (Ctrl+C para detener)');
});

function apagar(){
  console.log('\nDeteniendo servidor…');
  for(const res of streams){ try{ res.end(); }catch(e){} }
  servidor.close(() => process.exit(0));
  setTimeout(() => process.exit(0), 500).unref();
}
process.on('SIGINT', apagar);
process.on('SIGTERM', apagar);
