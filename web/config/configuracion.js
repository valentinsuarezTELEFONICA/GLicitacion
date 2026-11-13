/* Panel de configuración de Licitaciones.
 * Lee y escribe palabras_clave.xlsx directamente (no hay copia paralela de la configuración):
 *   Grupos      -> familias de servicio y sus puntos (la "prioridad")
 *   Palabras    -> palabras clave, cada una dentro de un grupo
 *   CPV         -> prefijos de código CPV y sus puntos
 *   Exclusiones -> palabras que restan puntos
 *   Ajustes     -> umbrales (solo se cambia el valor; el programa busca cada ajuste por su nombre)
 *                  y, en la pestaña «Empresa», el nombre de la empresa y los logos (filas de texto de Ajustes)
 * Solo se reescribe el contenido de esas cinco hojas; el resto del archivo no se toca.
 */
(function(){
"use strict";

var LIB_ZIP="config/jszip.min.js";
var HOJAS={
  grupos:{hoja:"Grupos",cab:["Grupo","Puntos","Descripción"]},
  palabras:{hoja:"Palabras",cab:["Grupo","Palabra o expresión","Tipo","Activa"]},
  cpv:{hoja:"CPV",cab:["Prefijo del código CPV","Puntos","Descripción"]},
  excl:{hoja:"Exclusiones",cab:["Palabra o expresión","Tipo","Activa"]},
  ajustes:{hoja:"Ajustes",cab:["Ajuste","Valor","Explicación"]}
};
var TABS=[{id:"grupos",t:"Grupos"},{id:"palabras",t:"Palabras clave"},{id:"cpv",t:"Códigos CPV"},{id:"exclusiones",t:"Exclusiones"},{id:"ajustes",t:"Ajustes"},{id:"empresa",t:"Empresa"}];
/* filas de texto de la hoja Ajustes (las lee el programa por el principio del nombre). Si faltan en el Excel,
   se añaden al leerlo para que se guarden con el resto. */
var AJ_TEXTO=[
  ["Nombre de la empresa","Aparece en la cabecera de la web y en el PDF. Vacío = el de configuracion.py."],
  ["Logo para el tema claro","Imagen de la carpeta web/config. Vacío = el icono por defecto."],
  ["Logo para el tema oscuro","Imagen de la carpeta web/config. Vacío = el mismo del tema claro."]
];

/* ====================================================================
   utilidades
   ==================================================================== */
function esc(s){return String(s==null?"":s).replace(/[&<>"']/g,function(c){return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c];});}
function escXml(s){return String(s).replace(/[&<>]/g,function(c){return {"&":"&amp;","<":"&lt;",">":"&gt;"}[c];}).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,"");}
function normaliza(s){return String(s==null?"":s).normalize("NFD").replace(/[̀-ͯ]/g,"").toLowerCase();}
function clave(s){return normaliza(s).replace(/\s+/g," ").trim();}
function pl(n,uno,varios){return n+" "+(n===1?uno:varios);}
function colNum(ref){var m=/^([A-Z]+)/.exec(ref||"");if(!m)return 0;var n=0;for(var i=0;i<m[1].length;i++)n=n*26+(m[1].charCodeAt(i)-64);return n;}
function colLetra(n){var s="";while(n>0){var r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26);}return s;}
function num(v){v=String(v==null?"":v).trim().replace(",",".");if(v==="")return null;var n=Number(v);return isNaN(n)?v:n;}
function iguales(a,b){if(a.byteLength!==b.byteLength)return false;var x=new Uint8Array(a),y=new Uint8Array(b);for(var i=0;i<x.length;i++)if(x[i]!==y[i])return false;return true;}
function hora(d){return d.toLocaleTimeString("es-ES",{hour:"2-digit",minute:"2-digit"});}

/* ====================================================================
   capa del Excel (xlsx = zip + XML). Solo se regenera <sheetData> de las cinco hojas.
   ==================================================================== */
var libP=null;
function cargarLib(){
  if(window.JSZip)return Promise.resolve();
  if(libP)return libP;
  libP=new Promise(function(ok,mal){
    var s=document.createElement("script");s.src=LIB_ZIP;
    s.onload=ok;s.onerror=function(){libP=null;mal(new Error("Falta el archivo config/jszip.min.js. Copia la carpeta config completa junto a index.html."));};
    document.head.appendChild(s);
  });
  return libP;
}
function xmlDoc(txt){
  var d=new DOMParser().parseFromString(txt,"application/xml");
  if(d.getElementsByTagName("parsererror").length)throw new Error("El archivo no es un Excel (.xlsx) válido.");
  return d;
}
function todos(el,nombre){return Array.prototype.slice.call(el.getElementsByTagNameNS("*",nombre));}
function hijoDe(el,nombre){for(var n=el.firstChild;n;n=n.nextSibling)if(n.localName===nombre)return n;return null;}
function textoSi(si){
  var t="";
  for(var n=si.firstChild;n;n=n.nextSibling){
    if(n.localName==="t")t+=n.textContent;
    else if(n.localName==="r"){var tt=hijoDe(n,"t");if(tt)t+=tt.textContent;}
  }
  return t;
}
async function abrirZip(buf){await cargarLib();try{return await JSZip.loadAsync(buf);}catch(e){throw new Error("El archivo no es un Excel (.xlsx) válido.");}}
async function texto(zip,ruta){var f=zip.file(ruta);if(!f)throw new Error("Al archivo le falta «"+ruta+"»: no parece un Excel normal.");return f.async("string");}

async function rutasHojas(zip){
  var wb=xmlDoc(await texto(zip,"xl/workbook.xml")),rel=xmlDoc(await texto(zip,"xl/_rels/workbook.xml.rels"));
  var dest={};
  todos(rel,"Relationship").forEach(function(r){dest[r.getAttribute("Id")]=r.getAttribute("Target");});
  var out={};
  todos(wb,"sheet").forEach(function(s){
    var id=s.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships","id")||s.getAttribute("r:id");
    var t=dest[id];if(!t)return;
    out[s.getAttribute("name")]=t.charAt(0)==="/"?t.slice(1):"xl/"+t.replace(/^\.\//,"");
  });
  return out;
}
async function cadenasCompartidas(zip){
  var f=zip.file("xl/sharedStrings.xml");if(!f)return [];
  return todos(xmlDoc(await f.async("string")),"si").map(textoSi);
}
function valorCelda(c,comp){
  var t=c.getAttribute("t"),v=hijoDe(c,"v"),tx=v?v.textContent:null;
  if(t==="inlineStr"){var is=hijoDe(c,"is");return is?textoSi(is):"";}
  if(tx==null||tx==="")return null;
  if(t==="s")return comp[parseInt(tx,10)];
  if(t==="str")return tx;
  if(t==="b")return tx==="1"?"VERDADERO":"FALSO";
  if(t==="e")return null;
  var n=Number(tx);return isNaN(n)?tx:n;
}
/* devuelve las filas de datos (sin la cabecera) y el estilo de celda que usa cada columna */
function leerHoja(xml,comp,ncols,nombre){
  var doc=xmlDoc(xml),filas=[],estilos={},extra=false;
  if(/<tableParts\b/.test(xml))throw new Error("La hoja «"+nombre+"» tiene una tabla de Excel. Conviértela en rango normal (Diseño de tabla > Convertir en rango) y vuelve a abrirla.");
  todos(doc,"row").forEach(function(row){
    var r=parseInt(row.getAttribute("r"),10);if(!(r>=2))return;
    var vals=[],algo=false;for(var i=0;i<ncols;i++)vals.push(null);
    todos(row,"c").forEach(function(c){
      var ci=colNum(c.getAttribute("r"));if(!ci)return;
      var v=valorCelda(c,comp);
      if(ci<=ncols){vals[ci-1]=v;if(v!==null&&v!=="")algo=true;if(!(ci in estilos)&&c.getAttribute("s")!=null)estilos[ci]=c.getAttribute("s");}
      else if(v!==null&&String(v).trim()!=="")extra=true;
    });
    if(algo)filas.push(vals);
  });
  return {filas:filas,estilos:estilos,extra:extra};
}
function cad(v){return v==null?"":String(v).trim();}
function numOrRaw(v){if(v==null||v==="")return null;if(typeof v==="number")return v;var n=Number(String(v).replace(",","."));return isNaN(n)?String(v):n;}
function esTexto(label){var k=clave(label);return k.indexOf("nombre de la empresa")===0||k.indexOf("logo")===0;}
function siNo(v){var k=clave(v);return k===""||k==="si"||k==="sí"||k==="s"||k==="1"||k==="verdadero"||k==="true"?"SÍ":"NO";}

/* lee las cinco hojas y devuelve el "plano": arrays simples, sin ids */
async function leerExcel(buf){
  var zip=await abrirZip(buf),rutas=await rutasHojas(zip),comp=await cadenasCompartidas(zip);
  var plano={},hojas={},avisos=[];
  for(var k in HOJAS){
    var def=HOJAS[k],ruta=rutas[def.hoja];
    if(!ruta)throw new Error("Este Excel no tiene la hoja «"+def.hoja+"». ¿Es el archivo de palabras clave?");
    var xml=await texto(zip,ruta),h=leerHoja(xml,comp,def.cab.length,def.hoja);
    // cabecera: debe coincidir para estar seguros de qué columna es cada cosa
    var doc=xmlDoc(xml),cab=[];
    todos(doc,"row").forEach(function(row){if(parseInt(row.getAttribute("r"),10)===1)todos(row,"c").forEach(function(c){var ci=colNum(c.getAttribute("r"));if(ci&&ci<=def.cab.length)cab[ci-1]=valorCelda(c,comp);});});
    for(var i=0;i<def.cab.length;i++)if(clave(cab[i])!==clave(def.cab[i]))throw new Error("En la hoja «"+def.hoja+"» la columna "+colLetra(i+1)+" debería llamarse «"+def.cab[i]+"» y se llama «"+cad(cab[i])+"».");
    if(h.extra)avisos.push({tab:k==="excl"?"exclusiones":k,txt:"La hoja «"+def.hoja+"» tiene datos en columnas que este panel no gestiona. Si guardas, se perderán."});
    hojas[k]={ruta:ruta,xml:xml,estilos:h.estilos};
    plano[k]=h.filas;
  }
  var p={
    grupos:plano.grupos.map(function(f){return [cad(f[0]),numOrRaw(f[1]),cad(f[2])];}),
    palabras:plano.palabras.map(function(f){return [cad(f[0]),cad(f[1]),clave(f[2])==="exacta"?"Exacta":"Normal",siNo(f[3])];}),
    cpv:plano.cpv.map(function(f){return [cad(f[0]),numOrRaw(f[1]),cad(f[2])];}),
    excl:plano.excl.map(function(f){return [cad(f[0]),clave(f[1])==="exacta"?"Exacta":"Normal",siNo(f[2])];}),
    ajustes:plano.ajustes.map(function(f){return [cad(f[0]),esTexto(f[0])?cad(f[1]):numOrRaw(f[1]),cad(f[2])];})
  };
  AJ_TEXTO.forEach(function(t){
    if(!p.ajustes.some(function(f){return clave(f[0]).indexOf(clave(t[0]))===0;}))p.ajustes.push([t[0],"",t[1]]);
  });
  plano.palabras.forEach(function(f){var t=clave(f[2]);if(t&&t!=="normal"&&t!=="exacta"){avisos.push({tab:"palabras",txt:"«"+cad(f[1])+"» tenía el tipo «"+cad(f[2])+"»; se ha tratado como Normal."});}});
  return {plano:p,hojas:hojas,avisos:avisos};
}

function celdaXml(ref,s,v){
  var st=s!=null?' s="'+s+'"':"";
  if(v===null||v===undefined||v==="")return s!=null?'<c r="'+ref+'"'+st+'/>':"";
  if(typeof v==="number")return '<c r="'+ref+'"'+st+'><v>'+v+'</v></c>';
  return '<c r="'+ref+'"'+st+' t="inlineStr"><is><t xml:space="preserve">'+escXml(v)+'</t></is></c>';
}
function reescribirHoja(xml,filas,ncols,estilos,extraGrupos){
  var m=/<sheetData\s*\/>|<sheetData[^>]*>[\s\S]*<\/sheetData>/.exec(xml);
  if(!m)throw new Error("No se reconoce el formato interno de una hoja del Excel.");
  var cab=/<row\b[^>]*\sr="1"[^>]*?(?:\/>|>[\s\S]*?<\/row>)/.exec(m[0]);
  var out=cab?cab[0]:"";
  filas.forEach(function(f,i){
    var r=i+2,c="";
    for(var j=0;j<ncols;j++)c+=celdaXml(colLetra(j+1)+r,estilos[j+1],f[j]);
    out+='<row r="'+r+'">'+c+'</row>';
  });
  xml=xml.slice(0,m.index)+"<sheetData>"+out+"</sheetData>"+xml.slice(m.index+m[0].length);
  xml=xml.replace(/<dimension\s+ref="[^"]*"\s*\/>/,'<dimension ref="A1:'+colLetra(ncols)+(filas.length+1)+'"/>');
  if(extraGrupos&&extraGrupos+1>60)xml=xml.replace(/(Grupos!\$A\$2:\$A\$)\d+/g,"$1"+(extraGrupos+1));
  return xml;
}
async function generarExcel(buf,hojas,plano){
  var zip=await abrirZip(buf);
  for(var k in HOJAS){
    var h=hojas[k];
    zip.file(h.ruta,reescribirHoja(h.xml,plano[k],HOJAS[k].cab.length,h.estilos,k==="palabras"?plano.grupos.length:0));
  }
  return zip.generateAsync({type:"arraybuffer",compression:"DEFLATE"});
}

/* ====================================================================
   modelo en memoria: grupos con id; las palabras apuntan al id (no al nombre),
   así renombrar un grupo se propaga solo y no puede quedar ninguna referencia rota.
   ==================================================================== */
var S={
  estado:"vacio",error:"",modelo:null,base:"",buf:null,hojas:null,nombre:"",handle:null,modificado:0,
  deshacer:[],sigId:1,tab:"grupos",filtro:{gid:"",q:"",tipo:"",act:""},sel:{},avisosCarga:[],
  val:{errores:[],avisos:[],marcas:{}},sucio:false,guardando:false,guardadoEn:null,cont:null,
  recordado:null,iniciado:false,cssOk:false
};
function nid(){return S.sigId++;}
function aModelo(p){
  var m={grupos:[],palabras:[],cpv:[],excl:[],ajustes:[]},porClave={};
  p.grupos.forEach(function(g){var o={id:nid(),nombre:g[0],puntos:g[1],desc:g[2]};m.grupos.push(o);var k=clave(o.nombre);if(!(k in porClave))porClave[k]=o.id;});
  p.palabras.forEach(function(f){var gid=porClave[clave(f[0])];m.palabras.push({id:nid(),gid:gid===undefined?null:gid,grupoTxt:f[0],palabra:f[1],tipo:f[2],activa:f[3]==="SÍ"});});
  p.cpv.forEach(function(f){m.cpv.push({id:nid(),prefijo:f[0],puntos:f[1],desc:f[2]});});
  p.excl.forEach(function(f){m.excl.push({id:nid(),palabra:f[0],tipo:f[1],activa:f[2]==="SÍ"});});
  p.ajustes.forEach(function(f){m.ajustes.push({id:nid(),label:f[0],valor:f[1],expl:f[2]});});
  return m;
}
function aPlano(m){
  var nom={};m.grupos.forEach(function(g){nom[g.id]=g.nombre;});
  return {
    grupos:m.grupos.map(function(g){return [g.nombre.trim(),g.puntos,(g.desc||"").trim()];}),
    palabras:m.palabras.map(function(p){return [p.gid!=null&&nom[p.gid]!==undefined?nom[p.gid].trim():(p.grupoTxt||""),p.palabra.trim(),p.tipo,p.activa?"SÍ":"NO"];}),
    cpv:m.cpv.map(function(c){return [String(c.prefijo).trim(),c.puntos,(c.desc||"").trim()];}),
    excl:m.excl.map(function(x){return [x.palabra.trim(),x.tipo,x.activa?"SÍ":"NO"];}),
    ajustes:m.ajustes.map(function(a){return [a.label,a.valor,a.expl];})
  };
}
function porId(lista,id){for(var i=0;i<lista.length;i++)if(lista[i].id===id)return lista[i];return null;}
function grupoDe(id){return porId(S.modelo.grupos,id);}
function cuentaPalabras(gid){var t=0,a=0;S.modelo.palabras.forEach(function(p){if(p.gid===gid){t++;if(p.activa)a++;}});return {total:t,activas:a};}
function snap(){S.deshacer.push(JSON.stringify(S.modelo));if(S.deshacer.length>80)S.deshacer.shift();}
function deshacer(){
  if(!S.deshacer.length)return;
  S.modelo=JSON.parse(S.deshacer.pop());S.sel={};render();
}

/* ---------- operaciones con cascada ---------- */
function existeEnGrupo(gid,palabra,tipo,exceptoId){
  var k=clave(palabra);
  return S.modelo.palabras.some(function(p){return p.gid===gid&&p.id!==exceptoId&&p.tipo===tipo&&clave(p.palabra)===k;});
}
/* mueve palabras a otro grupo; si ya existían allí se fusionan (no se duplican) */
function moverPalabras(ids,destino){
  var movidas=0,fusionadas=0,quitar={};
  ids.forEach(function(id){
    var p=porId(S.modelo.palabras,id);if(!p||p.gid===destino)return;
    if(existeEnGrupo(destino,p.palabra,p.tipo,p.id)){quitar[p.id]=1;fusionadas++;}
    else{p.gid=destino;p.grupoTxt="";movidas++;}
  });
  if(fusionadas)S.modelo.palabras=S.modelo.palabras.filter(function(p){return !quitar[p.id];});
  return {movidas:movidas,fusionadas:fusionadas};
}
function eliminarGrupo(id,modo,destino){
  var g=grupoDe(id),ids=S.modelo.palabras.filter(function(p){return p.gid===id;}).map(function(p){return p.id;}),r={movidas:0,fusionadas:0,borradas:0};
  snap();
  if(modo==="mover"&&destino!=null&&ids.length)r=Object.assign({borradas:0},moverPalabras(ids,destino));
  else if(ids.length){S.modelo.palabras=S.modelo.palabras.filter(function(p){return p.gid!==id;});r.borradas=ids.length;}
  S.modelo.grupos=S.modelo.grupos.filter(function(x){return x.id!==id;});
  return {g:g,r:r};
}

/* ---------- validación ---------- */
function validar(m){
  var R={errores:[],avisos:[],marcas:{}};
  function E(tab,txt,k){R.errores.push({tab:tab,txt:txt});if(k)R.marcas[k]=1;}
  function A(tab,txt,k){R.avisos.push({tab:tab,txt:txt});if(k&&!R.marcas[k])R.marcas[k]=2;}
  var nombres={},vistos={};
  m.grupos.forEach(function(g){
    var n=g.nombre.trim();nombres[g.id]=n;
    if(!n)E("grupos","Hay un grupo sin nombre.","g.nombre."+g.id);
    else{var k=clave(n);if(vistos[k]!==undefined){E("grupos","Hay dos grupos que se llaman «"+n+"».","g.nombre."+g.id);R.marcas["g.nombre."+vistos[k]]=1;}else vistos[k]=g.id;}
    if(typeof g.puntos!=="number"||!isFinite(g.puntos)||g.puntos<0)E("grupos","Los puntos de «"+(n||"grupo sin nombre")+"» deben ser un número igual o mayor que 0.","g.puntos."+g.id);
    if(n&&!m.palabras.some(function(p){return p.gid===g.id;}))A("grupos","El grupo «"+n+"» no tiene ninguna palabra clave: nunca sumará puntos.","g.nombre."+g.id);
  });
  var pk={},porPal={},activasExcl={};
  m.excl.forEach(function(x){if(x.activa&&x.palabra.trim())activasExcl[clave(x.palabra)]=x.palabra.trim();});
  m.palabras.forEach(function(p){
    var w=p.palabra.trim(),kk="p.palabra."+p.id;
    if(!w){E("palabras","Hay una palabra clave vacía.",kk);return;}
    if(p.gid==null||nombres[p.gid]===undefined)E("palabras","«"+w+"» no tiene grupo: elige uno.","p.gid."+p.id);
    else{
      var k=p.gid+"|"+clave(w)+"|"+p.tipo;
      if(pk[k])E("palabras","«"+w+"» está repetida en el grupo «"+nombres[p.gid]+"».",kk);else pk[k]=1;
    }
    var c=clave(w);if(!porPal[c])porPal[c]={txt:w,g:{}};porPal[c].g[p.gid]=1;
    if(p.tipo==="Normal"&&c.length<4)A("palabras","«"+w+"» es muy corta para el tipo Normal: puede aparecer dentro de otras palabras. Valora ponerla como Exacta.",kk);
    if(p.activa&&activasExcl[c])A("palabras","«"+w+"» está a la vez como palabra clave y como exclusión.",kk);
  });
  Object.keys(porPal).forEach(function(c){if(Object.keys(porPal[c].g).length>1)A("palabras","«"+porPal[c].txt+"» está en varios grupos y sumará los puntos de todos.");});
  var cp={};
  m.cpv.forEach(function(c){
    var p=String(c.prefijo).trim();
    if(!/^\d{1,8}$/.test(p))E("cpv","El prefijo CPV «"+p+"» debe tener solo cifras (de 1 a 8).","c.prefijo."+c.id);
    else if(cp[p])E("cpv","El prefijo CPV «"+p+"» está repetido.","c.prefijo."+c.id);else cp[p]=1;
    if(typeof c.puntos!=="number"||!isFinite(c.puntos)||c.puntos<0)E("cpv","Los puntos del CPV «"+p+"» deben ser un número igual o mayor que 0.","c.puntos."+c.id);
  });
  var xk={};
  m.excl.forEach(function(x){
    var w=x.palabra.trim();if(!w){E("exclusiones","Hay una exclusión vacía.","x.palabra."+x.id);return;}
    var k=clave(w)+"|"+x.tipo;if(xk[k])E("exclusiones","La exclusión «"+w+"» está repetida.","x.palabra."+x.id);else xk[k]=1;
  });
  var aj={};
  var logos=logosDisponibles();
  m.ajustes.forEach(function(a){
    if(esTexto(a.label)){
      var v=cad(a.valor);
      if(clave(a.label).indexOf("nombre")===0){if(v.length>60)E("empresa","El nombre de la empresa es muy largo (máximo 60 caracteres).","a.valor."+a.id);}
      else if(v&&logos&&logos.indexOf(v)<0)A("empresa","El logo «"+v+"» no está en la carpeta web/config: se usará el icono por defecto.","a.valor."+a.id);
      return;
    }
    var k=clave(a.label),opcional=k.indexOf("importe")===0;
    if(a.valor===null&&!opcional)E("ajustes","«"+a.label+"» necesita un valor.","a.valor."+a.id);
    else if(a.valor!==null&&(typeof a.valor!=="number"||!isFinite(a.valor)||a.valor<0))E("ajustes","«"+a.label+"» debe ser un número igual o mayor que 0.","a.valor."+a.id);
    if(k.indexOf("minima para que aparezca")>=0)aj.min=a;
    else if(k.indexOf("encaje es alto")>=0)aj.alto=a;
    else if(k.indexOf("importe minimo")===0)aj.imin=a;
    else if(k.indexOf("importe maximo")===0)aj.imax=a;
  });
  if(aj.min&&aj.alto&&typeof aj.min.valor==="number"&&typeof aj.alto.valor==="number"&&aj.alto.valor<aj.min.valor)E("ajustes","La puntuación para encaje ALTO ("+aj.alto.valor+") no puede ser menor que la mínima ("+aj.min.valor+").","a.valor."+aj.alto.id);
  if(aj.imin&&aj.imax&&typeof aj.imin.valor==="number"&&typeof aj.imax.valor==="number"&&aj.imin.valor>aj.imax.valor)E("ajustes","El importe mínimo es mayor que el máximo.","a.valor."+aj.imin.id);
  return R;
}

/* ====================================================================
   estilos (usa las variables de color de index.html, así sigue el tema claro/oscuro)
   ==================================================================== */
var CSS=[
".cfg .barra{position:sticky;top:57px;z-index:4;display:flex;gap:10px 14px;align-items:center;flex-wrap:wrap;background:var(--tarjeta);border:1px solid var(--linea);border-radius:10px;padding:10px 14px;margin-bottom:14px}",
"@media(max-width:860px){.cfg .barra{position:static}}",
".cfg .barra .estado{flex:1;min-width:220px}",
".cfg .barra .estado small{display:block;color:var(--suave)}",
".cfg .barra .acciones{display:flex;gap:8px;flex-wrap:wrap}",
".cfg .sucio{color:var(--aviso);font-weight:700}.cfg .limpio{color:var(--bien);font-weight:700}",
".cfg-pest{display:flex;gap:4px;flex-wrap:wrap;border-bottom:1px solid var(--linea);margin:0 0 14px}",
".cfg-pest a{padding:8px 14px;border-radius:6px 6px 0 0;text-decoration:none;color:var(--suave);font-weight:600;border:1px solid transparent;border-bottom:0;margin-bottom:-1px}",
".cfg-pest a:hover{color:var(--tinta)}",
".cfg-pest a[aria-current=page]{background:var(--tarjeta);border-color:var(--linea);color:var(--marca)}",
".cfg-pest .n{font-weight:400;color:var(--suave);margin-left:5px}",
".emp-nombre{display:flex;flex-direction:column;gap:4px;max-width:520px;font-weight:600;margin:0 0 18px}",
".emp-nombre input{font-weight:400}",
".logos{border:0;padding:0;margin:0 0 18px}",
".logos legend{font-weight:600;margin-bottom:8px}",
".logos-op{display:flex;gap:10px;flex-wrap:wrap}",
".logo-op{position:relative;display:flex;flex-direction:column;align-items:center;gap:6px;width:150px;padding:10px;border:2px solid var(--linea);border-radius:10px;cursor:pointer;text-align:center;font-size:12px;color:var(--suave);background:var(--tarjeta)}",
".logo-op:has(input:checked){border-color:var(--marca);color:var(--tinta)}",
".logo-op input{position:absolute;opacity:0;pointer-events:none}",
".logo-op:has(input:focus-visible){outline:3px solid var(--marca);outline-offset:2px}",
".logo-op .muestra{display:flex;align-items:center;justify-content:center;width:64px;height:64px;border-radius:8px;background:#ffffff}",
".logo-op.oscuro .muestra{background:#031a34}",
".logo-op .muestra img,.logo-op .muestra svg{width:48px;height:48px;object-fit:contain}",
".logo-op .nom{overflow-wrap:anywhere;line-height:1.25}",
".cfg-pest .e{background:var(--alerta);color:#fff;border-radius:9px;padding:0 6px;font-size:11px;margin-left:6px}",
".cfg .tabla input[type=text],.cfg .tabla input[type=number],.cfg .alta input[type=text],.cfg .alta input[type=number],.cfg select,.cfg textarea{border:1px solid var(--linea);background:var(--tarjeta);border-radius:6px;padding:6px 8px;width:100%;min-width:0}",
".cfg .tabla td{vertical-align:middle}",
".cfg .tabla input.mal,.cfg .tabla select.mal,.cfg .alta input.mal{border-color:var(--alerta);background:var(--alerta-fondo)}",
".cfg .tabla input.ojo,.cfg .tabla select.ojo{border-color:var(--aviso)}",
".cfg .tabla td.c-n{width:96px}.cfg .tabla td.c-t{width:120px}.cfg .tabla td.c-x{width:1%;white-space:nowrap}.cfg .tabla td.c-k{width:34px}",
".cfg .tabla th.c-k{width:34px}",
".cfg .envol{overflow-x:auto}",
".cfg .tabla{min-width:640px}",
".cfg .boton.peq{padding:4px 10px;font-weight:600;font-size:13px}",
".cfg .boton.peligro{background:none;color:var(--alerta);border-color:var(--alerta)}",
".cfg .boton[disabled]{opacity:.45;cursor:not-allowed}",
".cfg .alta{display:flex;gap:8px 10px;flex-wrap:wrap;align-items:end;margin-top:14px;padding-top:14px;border-top:1px dashed var(--linea)}",
".cfg .alta label{display:flex;flex-direction:column;font-size:12px;color:var(--suave);gap:2px}",
".cfg .alta .g{flex:2 1 220px}.cfg .alta .m{flex:1 1 150px}.cfg .alta .p{flex:0 0 90px}",
".cfg .err{color:var(--alerta);font-weight:600;min-height:1.4em;margin:6px 0 0}",
".cfg .barpr{display:block;height:8px;background:var(--mapa0);border-radius:4px;overflow:hidden;min-width:70px}",
".cfg .barpr i{display:block;height:100%;background:var(--marca)}",
".cfg .rango{font-size:12px;color:var(--suave);white-space:nowrap}",
".cfg .rango b{color:var(--tinta);font-size:15px}",
".cfg .sel{display:flex;gap:8px;flex-wrap:wrap;align-items:center;background:var(--marca-fondo);border-radius:8px;padding:8px 12px;margin:0 0 10px}",
".cfg .sel select{width:auto;max-width:260px}",
".cfg .ayuda{color:var(--suave);font-size:13px;margin:0 0 10px}",
".cfg .vacio-tabla{color:var(--suave);padding:14px 6px}",
".cfg ul.prob{margin:6px 0 0;padding-left:20px}",
".cfg .aviso details summary{cursor:pointer;font-weight:700}",
".cfg-bienvenida{max-width:640px}",
".cfg-bienvenida .acciones{display:flex;gap:10px;flex-wrap:wrap;margin:14px 0 4px}",
"nav.menu a[data-sucio]::after{content:\"\";display:inline-block;width:8px;height:8px;border-radius:50%;background:var(--aviso);margin-left:6px;vertical-align:middle}",
"dialog.cfg-dlg{border:1px solid var(--linea);border-radius:12px;background:var(--tarjeta);color:var(--tinta);padding:20px 22px;max-width:560px;width:calc(100% - 32px);font:inherit}",
"dialog.cfg-dlg::backdrop{background:rgba(10,20,26,.55)}",
"dialog.cfg-dlg h2{margin-top:0}",
"dialog.cfg-dlg fieldset{border:1px solid var(--linea);border-radius:8px;margin:12px 0;padding:10px 12px}",
"dialog.cfg-dlg legend{font-weight:700;padding:0 6px}",
"dialog.cfg-dlg label.op{display:flex;gap:8px;align-items:center;padding:5px 0;flex-wrap:wrap}",
"dialog.cfg-dlg select{border:1px solid var(--linea);background:var(--tarjeta);border-radius:6px;padding:6px 8px;max-width:100%}",
"dialog.cfg-dlg .pie{display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap;margin-top:16px}",
"dialog.cfg-dlg .err{color:var(--alerta);font-weight:600;min-height:1.4em;margin:6px 0 0}",
".cfg .filtros select,.cfg .filtros input{padding:7px 9px}",
"dialog.cfg-dlg .pals{display:flex;gap:4px;flex-wrap:wrap;margin-top:8px;max-height:140px;overflow:auto}",
"dialog.cfg-dlg .pals span{background:var(--marca-fondo);color:var(--marca);border-radius:4px;padding:0 6px;font-size:12px}",
".cfg-toast{position:fixed;left:16px;bottom:16px;z-index:20;background:var(--tinta);color:var(--papel);border-radius:10px;padding:10px 14px;max-width:min(520px,calc(100% - 32px));box-shadow:0 6px 24px rgba(0,0,0,.3);display:flex;gap:12px;align-items:center}",
".cfg-toast button{background:none;border:1px solid currentColor;color:inherit;border-radius:6px;padding:3px 10px;cursor:pointer;font-weight:700}",
".cfg-toast.mal{background:var(--alerta);color:#fff}"
].join("\n");
function css(){
  if(S.cssOk)return;S.cssOk=true;
  var st=document.createElement("style");st.id="cfg-css";st.textContent=CSS;document.head.appendChild(st);
}

/* ====================================================================
   avisos emergentes y ventanas de confirmación
   ==================================================================== */
var toastT=null;
function toast(msg,o){
  o=o||{};
  var t=document.getElementById("cfg-toast");if(t)t.remove();
  t=document.createElement("div");t.id="cfg-toast";t.className="cfg-toast"+(o.mal?" mal":"");t.setAttribute("role",o.mal?"alert":"status");
  t.innerHTML="<span>"+esc(msg)+"</span>"+(o.deshacer?'<button type="button">Deshacer</button>':"");
  document.body.appendChild(t);
  if(o.deshacer)t.querySelector("button").addEventListener("click",function(){t.remove();deshacer();});
  clearTimeout(toastT);toastT=setTimeout(function(){t.remove();},o.mal?12000:9000);
}
function modal(titulo,cuerpo,botones,alAbrir){
  return new Promise(function(res){
    var d=document.createElement("dialog");d.className="cfg-dlg";
    d.innerHTML='<form method="dialog"><h2>'+esc(titulo)+'</h2><div>'+cuerpo+'</div><div class="pie">'
      +botones.map(function(b){return '<button class="boton'+(b.cls?" "+b.cls:"")+'" value="'+esc(b.val)+'"'+(b.val==="no"?' formnovalidate':'')+'>'+esc(b.txt)+'</button>';}).join("")+'</div></form>';
    document.body.appendChild(d);
    d.addEventListener("close",function(){
      var datos=new FormData(d.querySelector("form"));
      var r={val:d.returnValue||"no",datos:{}};datos.forEach(function(v,k){r.datos[k]=v;});
      d.remove();res(r);
    });
    if(alAbrir)alAbrir(d);
    d.showModal();
  });
}

/* ====================================================================
   conexión con el archivo
   ==================================================================== */
function idb(){return new Promise(function(ok,mal){var r=indexedDB.open("licitaciones-config",1);r.onupgradeneeded=function(){r.result.createObjectStore("h");};r.onsuccess=function(){ok(r.result);};r.onerror=function(){mal(r.error);};});}
async function idbGet(k){try{var db=await idb();return await new Promise(function(ok){var q=db.transaction("h").objectStore("h").get(k);q.onsuccess=function(){ok(q.result);};q.onerror=function(){ok(null);};});}catch(e){return null;}}
async function idbSet(k,v){try{var db=await idb();await new Promise(function(ok){var t=db.transaction("h","readwrite");t.objectStore("h").put(v,k);t.oncomplete=ok;t.onerror=ok;});}catch(e){}}

async function cargar(buf,nombre,handle,mod){
  S.estado="cargando";S.error="";render();
  var r;
  try{r=await leerExcel(buf);}
  catch(e){S.estado="vacio";S.error=e.message||String(e);S.modelo=null;render();return;}
  S.modelo=aModelo(r.plano);S.base=JSON.stringify(r.plano);S.buf=buf;S.hojas=r.hojas;S.nombre=nombre;S.handle=handle;S.modificado=mod;
  S.deshacer=[];S.sel={};S.avisosCarga=r.avisos;S.estado="listo";S.guardadoEn=null;
  S.filtro={gid:"",q:"",tipo:"",act:""};
  render();
}
async function usarHandle(h){
  try{
    var f=await h.getFile(),buf=await f.arrayBuffer();
    await cargar(buf,f.name,h,f.lastModified);
    S.recordado=h;idbSet("xlsx",h);
  }catch(e){S.estado="vacio";S.error=e&&e.name==="NotFoundError"?"El archivo ya no está en su sitio. Elígelo de nuevo.":"No se pudo leer el archivo: "+(e.message||e);render();}
}
async function confirmarDescarte(){
  if(!S.sucio)return true;
  var r=await modal("Hay cambios sin guardar","<p>Si continúas, se perderán los cambios que aún no has guardado en el Excel.</p>",[{txt:"Seguir editando",val:"no",cls:"sec"},{txt:"Descartar cambios",val:"si",cls:"peligro"}]);
  return r.val==="si";
}
async function elegirArchivo(){
  if(window.showOpenFilePicker){
    if(!(await confirmarDescarte()))return;
    var hs;
    try{hs=await window.showOpenFilePicker({multiple:false,types:[{description:"Excel",accept:{"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":[".xlsx"]}}]});}
    catch(e){if(e&&e.name==="AbortError")return;S.error="No se pudo abrir el selector de archivos: "+(e.message||e);render();return;}
    await usarHandle(hs[0]);
  }else{
    var inp=S.cont.querySelector("#cfg-file");if(inp)inp.click();
  }
}
async function reconectar(){
  var h=S.recordado;if(!h)return;
  try{
    var q=await h.requestPermission({mode:"readwrite"});
    if(q!=="granted"){S.error="Sin permiso para el archivo. Pulsa de nuevo y acepta.";render();return;}
  }catch(e){}
  await usarHandle(h);
}
async function releer(){
  if(!S.handle)return;
  if(!(await confirmarDescarte()))return;
  await usarHandle(S.handle);
  toast("Datos leídos de nuevo desde el Excel.");
}
async function iniciar(){
  if(S.iniciado)return;S.iniciado=true;
  var h=await idbGet("xlsx");if(!h)return;
  S.recordado=h;
  try{
    var q=await h.queryPermission({mode:"readwrite"});
    if(q==="granted"&&S.estado==="vacio"){await usarHandle(h);return;}
  }catch(e){}
  if(S.estado==="vacio")render();
}

/* ====================================================================
   guardado seguro:
   1) valida  2) genera el .xlsx nuevo en memoria  3) lo vuelve a leer y comprueba que coincide
   4) escribe (el navegador escribe en un archivo temporal y solo lo sustituye al cerrar: si falla, el original queda intacto)
   5) lo lee del disco otra vez y comprueba que coincide
   ==================================================================== */
function descargar(buf,nombre){
  var b=new Blob([buf],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),a=document.createElement("a");
  a.href=URL.createObjectURL(b);a.download=nombre||"palabras_clave.xlsx";document.body.appendChild(a);a.click();a.remove();
  setTimeout(function(){URL.revokeObjectURL(a.href);},4000);
}
async function guardar(){
  if(S.guardando||S.estado!=="listo")return;
  S.val=validar(S.modelo);
  if(S.val.errores.length){
    refrescar();var p=S.cont.querySelector("#cfg-prob");if(p&&p.scrollIntoView)p.scrollIntoView({block:"center"});
    toast("Hay "+pl(S.val.errores.length,"cosa","cosas")+" que corregir antes de guardar.",{mal:true});return;
  }
  S.guardando=true;refrescar();
  try{
    var p=aPlano(S.modelo),nuevo=await generarExcel(S.buf,S.hojas,p);
    var chk=await leerExcel(nuevo);
    if(JSON.stringify(chk.plano)!==JSON.stringify(p))throw new Error("La comprobación previa no coincide con lo que querías guardar. No se ha escrito nada.");
    var mod=S.modificado;
    if(S.handle){
      var h=S.handle;
      try{if((await h.queryPermission({mode:"readwrite"}))!=="granted"&&(await h.requestPermission({mode:"readwrite"}))!=="granted")throw Object.assign(new Error("permiso"),{name:"NotAllowedError"});}
      catch(e){if(e.name==="NotAllowedError")throw e;}
      var f0=await h.getFile();
      if(f0.lastModified!==S.modificado&&!iguales(await f0.arrayBuffer(),S.buf)){
        var r=await modal("El Excel ha cambiado desde que lo abriste","<p>Alguien (o Excel) ha modificado <b>"+esc(S.nombre)+"</b> después de que lo cargaras aquí. Si guardas ahora, esos cambios se sobrescribirán.</p><p>Para conservarlos, cancela, pulsa «Releer del Excel» y vuelve a hacer tus cambios.</p>",[{txt:"Cancelar",val:"no",cls:"sec"},{txt:"Sobrescribir de todos modos",val:"si",cls:"peligro"}]);
        if(r.val!=="si"){S.guardando=false;refrescar();return;}
      }
      var w=await h.createWritable();
      try{await w.write(nuevo);await w.close();}catch(e){try{await w.abort();}catch(x){}throw e;}
      var f1=await h.getFile(),disco=await f1.arrayBuffer(),rel=await leerExcel(disco);
      if(JSON.stringify(rel.plano)!==JSON.stringify(p))throw new Error("El archivo guardado no coincide con lo esperado. Revisa el Excel antes de usarlo.");
      S.buf=disco;S.hojas=rel.hojas;mod=f1.lastModified;S.modificado=mod;
    }else{
      descargar(nuevo,S.nombre);S.buf=nuevo;S.hojas=chk.hojas;
    }
    S.base=JSON.stringify(p);S.guardadoEn=new Date();
    aplicarMarca();
    toast(S.handle?"Guardado en "+S.nombre+" y comprobado.":"Excel descargado. Sustituye el original por este archivo.");
  }catch(e){
    var msg=e&&e.message?e.message:String(e);
    if(e&&e.name==="NotAllowedError")msg="No hay permiso para escribir en el archivo. Pulsa Guardar de nuevo y acepta el permiso.";
    else if(e&&(e.name==="NoModificationAllowedError"||/lock|bloque|in use|being used/i.test(msg)))msg="El archivo está bloqueado. Si lo tienes abierto en Excel, ciérralo y vuelve a pulsar Guardar.";
    toast("No se ha guardado: "+msg+" Tus cambios siguen aquí.",{mal:true});
  }finally{S.guardando=false;refrescar();}
}
function descartar(){
  if(!S.sucio)return;
  modal("Descartar cambios","<p>Se volverá a la configuración que hay ahora mismo en el Excel y se perderán los cambios sin guardar.</p>",[{txt:"Cancelar",val:"no",cls:"sec"},{txt:"Descartar cambios",val:"si",cls:"peligro"}]).then(function(r){
    if(r.val!=="si")return;
    S.modelo=aModelo(JSON.parse(S.base));S.deshacer=[];S.sel={};render();toast("Cambios descartados.");
  });
}

/* ====================================================================
   vistas
   ==================================================================== */
function vistaConexion(){
  var h='<section class="panel cfg-bienvenida"><h2>Conecta el Excel de palabras clave</h2>'
   +'<p>Este panel lee y guarda directamente en <b>palabras_clave.xlsx</b>, el mismo archivo que usa el programa. No hay ninguna copia aparte de la configuración.</p>';
  if(S.error)h+='<div class="aviso mal" role="alert"><strong>No se pudo abrir.</strong>'+esc(S.error)+'</div>';
  if(S.estado==="cargando")h+='<p aria-live="polite">Leyendo el archivo…</p>';
  else{
    h+='<div class="acciones">';
    if(S.recordado&&window.showOpenFilePicker)h+='<button class="boton" type="button" data-act="reconectar">Reconectar con «'+esc(S.recordado.name||"palabras_clave.xlsx")+'»</button>';
    h+='<button class="boton'+(S.recordado&&window.showOpenFilePicker?" sec":"")+'" type="button" data-act="elegir">Elegir palabras_clave.xlsx</button></div>';
    h+=window.showOpenFilePicker
      ?'<p class="ayuda">La primera vez el navegador pedirá permiso para modificar el archivo. Si lo tienes abierto en Excel, ciérralo antes de guardar.</p>'
      :'<div class="aviso ojo"><strong>Este navegador no puede escribir en el archivo.</strong>Podrás editar todo, y al guardar se descargará el Excel actualizado para que sustituyas el original. Con Chrome o Edge se guarda directamente.</div>';
  }
  return h+'</section>';
}
function barraHTML(){
  var errs=S.val.errores.length,estado;
  if(S.guardando)estado='<span>Guardando…</span>';
  else if(S.sucio)estado='<span class="sucio">Cambios sin guardar</span>';
  else estado='<span class="limpio">'+(S.guardadoEn?"Guardado y comprobado a las "+hora(S.guardadoEn):"Sin cambios")+'</span>';
  return '<div class="estado"><b>'+esc(S.nombre)+'</b> · '+estado
    +'<small>'+(S.handle?'Se guarda directamente en este archivo.':'Este navegador no escribe en el archivo: al guardar se descarga el Excel actualizado.')+'</small></div>'
    +'<div class="acciones">'
    +'<button class="boton sec peq" type="button" data-act="deshacer"'+(S.deshacer.length?'':' disabled')+'>Deshacer'+(S.deshacer.length?' ('+S.deshacer.length+')':'')+'</button>'
    +'<button class="boton sec peq" type="button" data-act="descartar"'+(S.sucio?'':' disabled')+'>Descartar cambios</button>'
    +(S.handle?'<button class="boton sec peq" type="button" data-act="releer">Releer del Excel</button>':'')
    +'<button class="boton sec peq" type="button" data-act="elegir">Cambiar archivo</button>'
    +'<button class="boton" type="button" data-act="guardar"'+((S.sucio&&!S.guardando)?'':' disabled')+'>'+(S.handle?'Guardar en el Excel':'Descargar Excel actualizado')+(errs&&S.sucio?' ('+errs+' por corregir)':'')+'</button></div>';
}
function listaProb(arr,max){
  return '<ul class="prob">'+arr.slice(0,max).map(function(x){return '<li><a href="#/config/'+x.tab+'">'+esc(x.txt)+'</a></li>';}).join("")+(arr.length>max?'<li>… y '+(arr.length-max)+' más.</li>':'')+'</ul>';
}
function problemasHTML(){
  var h="",e=S.val.errores,a=S.val.avisos.concat(S.avisosCarga);
  if(e.length)h+='<div class="aviso mal" role="alert"><strong>Corrige '+pl(e.length,"cosa","cosas")+' antes de guardar</strong>'+listaProb(e,8)+'</div>';
  if(a.length)h+='<div class="aviso ojo"><details><summary>'+pl(a.length,"aviso que conviene revisar","avisos que conviene revisar")+' (no impiden guardar)</summary>'+listaProb(a,12)+'</details></div>';
  return h;
}
function pestanasHTML(){
  var m=S.modelo,cnt={grupos:m.grupos.length,palabras:m.palabras.length,cpv:m.cpv.length,exclusiones:m.excl.length,ajustes:"",empresa:""};
  return TABS.map(function(t){
    var ne=S.val.errores.filter(function(x){return x.tab===t.id;}).length;
    return '<a href="#/config/'+t.id+'"'+(S.tab===t.id?' aria-current="page"':'')+'>'+t.t+(cnt[t.id]!==""?'<span class="n">'+cnt[t.id]+'</span>':'')+(ne?'<span class="e" title="'+pl(ne,"error","errores")+'">'+ne+'</span>':'')+'</a>';
  }).join("");
}
function mk(k){var v=S.val.marcas[k];return v===1?" mal":v===2?" ojo":"";}
function opcionesGrupos(sel,conVacio){
  var h=conVacio?'<option value="">'+conVacio+'</option>':"";
  return h+S.modelo.grupos.map(function(g){return '<option value="'+g.id+'"'+(g.id===sel?" selected":"")+'>'+esc(g.nombre||"(sin nombre)")+'</option>';}).join("");
}
function opcionesTipo(sel){return ["Normal","Exacta"].map(function(t){return '<option'+(t===sel?" selected":"")+'>'+t+'</option>';}).join("");}
function btnEliminar(act,id,que){return '<button class="boton sec peq peligro" type="button" data-act="'+act+'" data-id="'+id+'" aria-label="Eliminar '+esc(que)+'">Eliminar</button>';}

/* ---------- Grupos ---------- */
function vistaGrupos(){
  var m=S.modelo,max=1;
  m.grupos.forEach(function(g){if(typeof g.puntos==="number"&&g.puntos>max)max=g.puntos;});
  var h='<p class="ayuda">Cada grupo es una familia de servicio. Si en una licitación aparece cualquiera de sus palabras, suma los puntos del grupo (una sola vez por grupo). <b>Los puntos son la prioridad</b>: cuantos más, más pesa el grupo. Al renombrar un grupo, todas sus palabras lo siguen; al borrarlo, eliges qué hacer con ellas.</p>'
   +'<div class="envol"><table class="tabla" id="t-grupos"><thead><tr><th>Prioridad</th><th>Grupo</th><th>Puntos</th><th>Descripción</th><th>Palabras</th><th></th></tr></thead><tbody>';
  if(!m.grupos.length)h+='<tr><td colspan="6" class="vacio-tabla">No hay grupos. Crea el primero abajo.</td></tr>';
  m.grupos.forEach(function(g){
    var c=cuentaPalabras(g.id);
    h+='<tr data-id="'+g.id+'"><td class="c-n"><span class="rango" data-rango="'+g.id+'"></span><span class="barpr"><i data-barra="'+g.id+'" style="width:0"></i></span></td>'
     +'<td><input type="text" data-f="g.nombre" data-id="'+g.id+'" data-k="g.nombre.'+g.id+'" class="'+mk("g.nombre."+g.id).trim()+'" value="'+esc(g.nombre)+'" aria-label="Nombre del grupo"></td>'
     +'<td class="c-n"><input type="number" min="0" step="1" data-f="g.puntos" data-id="'+g.id+'" data-k="g.puntos.'+g.id+'" class="'+mk("g.puntos."+g.id).trim()+'" value="'+(g.puntos==null?"":esc(g.puntos))+'" aria-label="Puntos de '+esc(g.nombre)+'"></td>'
     +'<td><input type="text" data-f="g.desc" data-id="'+g.id+'" value="'+esc(g.desc)+'" aria-label="Descripción de '+esc(g.nombre)+'"></td>'
     +'<td class="c-x"><a href="#/config/palabras" data-act="ver-palabras" data-id="'+g.id+'" data-cuenta="'+g.id+'">'+pl(c.total,"palabra","palabras")+'</a></td>'
     +'<td class="c-x">'+btnEliminar("del-grupo",g.id,g.nombre)+'</td></tr>';
  });
  h+='</tbody></table></div>'
   +'<div style="margin-top:10px"><button class="boton sec peq" type="button" data-act="ordenar-grupos">Ordenar de mayor a menor prioridad</button></div>'
   +'<div class="alta"><label class="g">Nuevo grupo<input type="text" id="ng-nombre" placeholder="Nombre del grupo"></label>'
   +'<label class="p">Puntos<input type="number" min="0" step="1" id="ng-puntos" value="30"></label>'
   +'<label class="g">Descripción (opcional)<input type="text" id="ng-desc"></label>'
   +'<button class="boton" type="button" data-act="add-grupo">Añadir grupo</button></div><p class="err" id="ng-err" role="alert"></p>';
  return h;
}

/* ---------- Palabras ---------- */
function coincidePalabra(p){
  var f=S.filtro;
  if(f.gid==="none"){if(p.gid!==null&&grupoDe(p.gid))return false;}
  else if(f.gid&&String(p.gid)!==f.gid)return false;
  if(f.tipo&&p.tipo!==f.tipo)return false;
  if(f.act==="si"&&!p.activa)return false;
  if(f.act==="no"&&p.activa)return false;
  if(f.q&&clave(p.palabra).indexOf(clave(f.q))<0)return false;
  return true;
}
function filaPalabra(p){
  var huerf=p.gid==null||!grupoDe(p.gid);
  return '<tr data-id="'+p.id+'"'+(coincidePalabra(p)?"":" hidden")+'>'
   +'<td class="c-k"><input type="checkbox" data-sel="'+p.id+'"'+(S.sel[p.id]?" checked":"")+' aria-label="Seleccionar «'+esc(p.palabra)+'»"></td>'
   +'<td><input type="text" data-f="p.palabra" data-id="'+p.id+'" data-k="p.palabra.'+p.id+'" class="'+mk("p.palabra."+p.id).trim()+'" value="'+esc(p.palabra)+'" aria-label="Palabra clave"></td>'
   +'<td><select data-f="p.gid" data-id="'+p.id+'" data-k="p.gid.'+p.id+'" class="'+mk("p.gid."+p.id).trim()+'" aria-label="Grupo de «'+esc(p.palabra)+'»">'
     +(huerf?'<option value="" selected>— elige un grupo'+(p.grupoTxt?' (era «'+esc(p.grupoTxt)+'»)':'')+' —</option>':"")+opcionesGrupos(p.gid)+'</select></td>'
   +'<td class="c-t"><select data-f="p.tipo" data-id="'+p.id+'" aria-label="Tipo de «'+esc(p.palabra)+'»">'+opcionesTipo(p.tipo)+'</select></td>'
   +'<td class="c-x"><label><input type="checkbox" data-f="p.activa" data-id="'+p.id+'"'+(p.activa?" checked":"")+'> Activa</label></td>'
   +'<td class="c-x">'+btnEliminar("del-palabra",p.id,p.palabra)+'</td></tr>';
}
function vistaPalabras(){
  var m=S.modelo,f=S.filtro,hayHuerf=m.palabras.some(function(p){return p.gid==null||!grupoDe(p.gid);});
  var h='<p class="ayuda"><b>Normal</b>: acepta terminaciones («museo» encuentra museos, museográfico); escribe la raíz. <b>Exacta</b>: solo la palabra completa, para siglas como VR, CMS o app. No importan mayúsculas ni tildes. Una palabra pertenece a un grupo: cámbiala de grupo desde su fila.</p>'
   +'<form class="filtros" id="f-pal" onsubmit="return false"><label>Grupo<select id="fl-gid">'+opcionesGrupos(/^\d+$/.test(f.gid)?Number(f.gid):-1,"Todos")+(hayHuerf?'<option value="none"'+(f.gid==="none"?" selected":"")+'>Sin grupo asignado</option>':"")+'</select></label>'
   +'<label>Buscar<input type="search" id="fl-q" value="'+esc(f.q)+'" placeholder="Palabra"></label>'
   +'<label>Tipo<select id="fl-tipo"><option value="">Todos</option><option'+(f.tipo==="Normal"?" selected":"")+'>Normal</option><option'+(f.tipo==="Exacta"?" selected":"")+'>Exacta</option></select></label>'
   +'<label>Estado<select id="fl-act"><option value="">Todas</option><option value="si"'+(f.act==="si"?" selected":"")+'>Activas</option><option value="no"'+(f.act==="no"?" selected":"")+'>Desactivadas</option></select></label>'
   +'<button class="boton sec peq" type="button" data-act="quitar-filtros">Quitar filtros</button></form>'
   +'<p class="cuenta" id="pal-cuenta" aria-live="polite"></p>'
   +'<div id="sel-barra"></div>'
   +'<div class="envol"><table class="tabla" id="t-palabras"><thead><tr><th class="c-k"><input type="checkbox" id="sel-todas" aria-label="Seleccionar las visibles"></th><th>Palabra o expresión</th><th>Grupo</th><th>Tipo</th><th>Estado</th><th></th></tr></thead><tbody>'
   +m.palabras.map(filaPalabra).join("")+'</tbody></table></div>'
   +'<p class="vacio-tabla" id="pal-vacio" hidden>Ninguna palabra con estos filtros.</p>';
  if(!m.grupos.length)return h+'<p class="aviso ojo">Para añadir palabras primero crea un grupo en la pestaña Grupos.</p>';
  var gsel=/^\d+$/.test(f.gid)?Number(f.gid):(m.grupos[0]&&m.grupos[0].id);
  h+='<div class="alta"><label class="g">Nueva palabra clave<input type="text" id="np-pal" placeholder="Palabra o expresión"></label>'
   +'<label class="g">Grupo<select id="np-gid">'+opcionesGrupos(gsel)+'</select></label>'
   +'<label class="m">Tipo<select id="np-tipo">'+opcionesTipo("Normal")+'</select></label>'
   +'<button class="boton" type="button" data-act="add-palabra">Añadir palabra</button></div><p class="err" id="np-err" role="alert"></p>'
   +'<details style="margin-top:6px"><summary>Añadir varias palabras de golpe</summary><div class="alta" style="border-top:0;margin-top:6px;padding-top:0">'
   +'<label class="g">Una por línea<textarea id="nv-txt" rows="5"></textarea></label>'
   +'<label class="g">Grupo<select id="nv-gid">'+opcionesGrupos(gsel)+'</select></label>'
   +'<label class="m">Tipo<select id="nv-tipo">'+opcionesTipo("Normal")+'</select></label>'
   +'<button class="boton" type="button" data-act="add-varias">Añadir todas</button></div><p class="err" id="nv-err" role="alert"></p></details>';
  return h;
}

/* ---------- CPV ---------- */
function vistaCpv(){
  var m=S.modelo;
  var h='<p class="ayuda">Prefijos del código CPV de la licitación (el prefijo «72» cubre todos los 72xxxxxx). Cada uno da los puntos indicados; si coinciden varios, cuenta el mayor. Son independientes de los grupos.</p>'
   +'<div class="envol"><table class="tabla"><thead><tr><th>Prefijo CPV</th><th>Puntos</th><th>Descripción</th><th></th></tr></thead><tbody>';
  if(!m.cpv.length)h+='<tr><td colspan="4" class="vacio-tabla">No hay códigos CPV.</td></tr>';
  m.cpv.forEach(function(c){
    h+='<tr data-id="'+c.id+'"><td class="c-n"><input type="text" inputmode="numeric" data-f="c.prefijo" data-id="'+c.id+'" data-k="c.prefijo.'+c.id+'" class="'+mk("c.prefijo."+c.id).trim()+'" value="'+esc(c.prefijo)+'" aria-label="Prefijo CPV"></td>'
     +'<td class="c-n"><input type="number" min="0" step="1" data-f="c.puntos" data-id="'+c.id+'" data-k="c.puntos.'+c.id+'" class="'+mk("c.puntos."+c.id).trim()+'" value="'+(c.puntos==null?"":esc(c.puntos))+'" aria-label="Puntos del CPV '+esc(c.prefijo)+'"></td>'
     +'<td><input type="text" data-f="c.desc" data-id="'+c.id+'" value="'+esc(c.desc)+'" aria-label="Descripción del CPV '+esc(c.prefijo)+'"></td>'
     +'<td class="c-x">'+btnEliminar("del-cpv",c.id,"CPV "+c.prefijo)+'</td></tr>';
  });
  return h+'</tbody></table></div><div class="alta"><label class="m">Nuevo prefijo<input type="text" inputmode="numeric" id="nc-pre" placeholder="Ej. 7222"></label>'
   +'<label class="p">Puntos<input type="number" min="0" step="1" id="nc-puntos" value="25"></label>'
   +'<label class="g">Descripción (opcional)<input type="text" id="nc-desc"></label>'
   +'<button class="boton" type="button" data-act="add-cpv">Añadir CPV</button></div><p class="err" id="nc-err" role="alert"></p>';
}

/* ---------- Exclusiones ---------- */
function vistaExcl(){
  var m=S.modelo;
  var h='<p class="ayuda">Palabras que restan puntos cuando aparecen (obras, limpieza…). Funcionan igual que las palabras clave. La penalización se ajusta en la pestaña Ajustes.</p>'
   +'<div class="envol"><table class="tabla"><thead><tr><th>Palabra o expresión</th><th>Tipo</th><th>Estado</th><th></th></tr></thead><tbody>';
  if(!m.excl.length)h+='<tr><td colspan="4" class="vacio-tabla">No hay exclusiones.</td></tr>';
  m.excl.forEach(function(x){
    h+='<tr data-id="'+x.id+'"><td><input type="text" data-f="x.palabra" data-id="'+x.id+'" data-k="x.palabra.'+x.id+'" class="'+mk("x.palabra."+x.id).trim()+'" value="'+esc(x.palabra)+'" aria-label="Exclusión"></td>'
     +'<td class="c-t"><select data-f="x.tipo" data-id="'+x.id+'" aria-label="Tipo de «'+esc(x.palabra)+'»">'+opcionesTipo(x.tipo)+'</select></td>'
     +'<td class="c-x"><label><input type="checkbox" data-f="x.activa" data-id="'+x.id+'"'+(x.activa?" checked":"")+'> Activa</label></td>'
     +'<td class="c-x">'+btnEliminar("del-excl",x.id,x.palabra)+'</td></tr>';
  });
  return h+'</tbody></table></div><div class="alta"><label class="g">Nueva exclusión<input type="text" id="nx-pal" placeholder="Palabra o expresión"></label>'
   +'<label class="m">Tipo<select id="nx-tipo">'+opcionesTipo("Normal")+'</select></label>'
   +'<button class="boton" type="button" data-act="add-excl">Añadir exclusión</button></div><p class="err" id="nx-err" role="alert"></p>';
}

/* ---------- Ajustes ---------- */
function vistaAjustes(){
  var h='<p class="ayuda">Umbrales de puntuación e importes. Estos ajustes son fijos: el programa busca cada uno por su nombre, así que aquí solo se cambia el valor. En los importes, vacío significa sin límite.</p>'
   +'<div class="envol"><table class="tabla"><thead><tr><th>Ajuste</th><th>Valor</th><th>Explicación</th></tr></thead><tbody>';
  S.modelo.ajustes.forEach(function(a){
    if(esTexto(a.label))return;                     // nombre y logos: pestaña Empresa
    h+='<tr><td>'+esc(a.label)+'</td><td class="c-n"><input type="number" min="0" step="any" data-f="a.valor" data-id="'+a.id+'" data-k="a.valor.'+a.id+'" class="'+mk("a.valor."+a.id).trim()+'" value="'+(a.valor==null?"":esc(a.valor))+'" aria-label="'+esc(a.label)+'"></td><td class="sub" style="margin:0">'+esc(a.expl)+'</td></tr>';
  });
  return h+'</tbody></table></div>';
}

/* ---------- Empresa: nombre y logos de la cabecera ---------- */
/* imágenes de web/config que el programa ha publicado en datos/indice.js (null si aún no hay datos) */
function logosDisponibles(){var d=window.DATOS;return d&&Array.isArray(d.logos)?d.logos:null;}
function ajTexto(pref){return S.modelo.ajustes.filter(function(a){return clave(a.label).indexOf(clave(pref))===0;})[0];}
var ICONO_DEFECTO='<svg viewBox="0 0 34 34" aria-hidden="true"><rect width="34" height="34" rx="9" fill="#0066ff"/><circle cx="16" cy="18" r="7.5" fill="none" stroke="#fff" stroke-width="3"/><circle cx="25" cy="9" r="3.6" fill="#fff"/></svg>';
function selectorLogo(a,titulo,oscuro,logos){
  if(!a)return "";
  var v=cad(a.valor),ops=[["",oscuro?"El mismo del tema claro":"Icono por defecto"]].concat((logos||[]).map(function(n){return [n,n];}));
  if(v&&!ops.some(function(o){return o[0]===v;}))ops.push([v,v+" (no encontrado)"]);
  return '<fieldset class="logos"><legend>'+esc(titulo)+'</legend><div class="logos-op">'+ops.map(function(o){
    return '<label class="logo-op'+(oscuro?' oscuro':'')+'"><input type="radio" name="logo-'+a.id+'" data-f="a.texto" data-id="'+a.id+'" data-k="a.valor.'+a.id+'" value="'+esc(o[0])+'"'+(v===o[0]?' checked':'')+'>'
      +'<span class="muestra">'+(o[0]?'<img src="config/'+encodeURIComponent(o[0])+'" alt="">':ICONO_DEFECTO)+'</span><span class="nom">'+esc(o[1])+'</span></label>';
  }).join("")+'</div></fieldset>';
}
function vistaEmpresa(){
  var n=ajTexto("Nombre de la empresa"),lc=ajTexto("Logo para el tema claro"),lo=ajTexto("Logo para el tema oscuro"),logos=logosDisponibles();
  var h='<p class="ayuda">Nombre y logo de la cabecera de la web (el nombre sale también en el PDF del día). Al guardar se ven al momento en este navegador; '
   +'para todos, a partir de la próxima ejecución del programa, o al momento con <code>py licitaciones.py --solo-web</code>.</p>';
  if(n)h+='<label class="emp-nombre">Nombre de la empresa<input type="text" maxlength="80" data-f="a.texto" data-id="'+n.id+'" data-k="a.valor.'+n.id+'" class="'+mk("a.valor."+n.id).trim()+'" value="'+esc(cad(n.valor))+'" placeholder="Vacío = el de configuracion.py"></label>';
  if(!logos)h+='<p class="aviso ojo">Todavía no se ve la lista de logos: ejecuta el programa una vez para que la publique.</p>';
  h+=selectorLogo(lc,"Logo para el tema claro",false,logos)+selectorLogo(lo,"Logo para el tema oscuro (se muestra sobre fondo oscuro)",true,logos);
  return h+'<p class="sub" style="margin:0">Para añadir otro logo, copia la imagen (PNG, JPG o SVG; mejor cuadrada) en la carpeta <b>web/config</b> del programa y ejecútalo: aparecerá aquí.</p>';
}
/* aplica el nombre y el logo guardados a la cabecera de esta página */
function aplicarMarca(){
  if(typeof window.pintarMarca!=="function")return;
  var n=ajTexto("Nombre de la empresa"),lc=ajTexto("Logo para el tema claro"),lo=ajTexto("Logo para el tema oscuro"),logos=logosDisponibles();
  function ok(a){var v=a?cad(a.valor):"";return v&&(!logos||logos.indexOf(v)>=0)?v:"";}
  var d=window.DATOS||{};
  window.pintarMarca(n&&cad(n.valor)?cad(n.valor):(d.empresaPorDefecto||d.empresa||""),ok(lc),ok(lo));
}

/* ====================================================================
   pintado
   ==================================================================== */
function render(){
  if(!S.cont)return;
  css();
  var h='<div class="cfg"><h1>Configuración</h1><p class="sub">Qué licitaciones te interesan y cuánto pesa cada cosa. Todo se guarda en el Excel de palabras clave.</p>';
  if(S.estado!=="listo")h+=vistaConexion();
  else{
    S.val=validar(S.modelo);S.sucio=JSON.stringify(aPlano(S.modelo))!==S.base;
    h+='<div class="barra" id="cfg-barra">'+barraHTML()+'</div><div id="cfg-prob">'+problemasHTML()+'</div>'
     +'<nav class="cfg-pest" id="cfg-pest" aria-label="Apartados de la configuración">'+pestanasHTML()+'</nav>'
     +'<section class="panel" id="cfg-cuerpo">'+({grupos:vistaGrupos,palabras:vistaPalabras,cpv:vistaCpv,exclusiones:vistaExcl,ajustes:vistaAjustes,empresa:vistaEmpresa}[S.tab]||vistaGrupos)()+'</section>';
  }
  S.cont.innerHTML=h+'<input type="file" id="cfg-file" accept=".xlsx" hidden aria-hidden="true" tabindex="-1"></div>';
  if(S.estado==="listo"){pintarPrioridades();aplicarFiltros(true);marcarMenu();}
}
function marcarMenu(){var a=document.querySelector('nav.menu a[data-v="config"]');if(a){if(S.sucio)a.setAttribute("data-sucio","");else a.removeAttribute("data-sucio");}}
/* actualización parcial: no toca los campos donde se está escribiendo */
function refrescar(){
  if(!S.cont||S.estado!=="listo")return;
  S.val=validar(S.modelo);S.sucio=JSON.stringify(aPlano(S.modelo))!==S.base;
  var b=S.cont.querySelector("#cfg-barra"),p=S.cont.querySelector("#cfg-prob"),t=S.cont.querySelector("#cfg-pest");
  if(b)b.innerHTML=barraHTML();if(p)p.innerHTML=problemasHTML();if(t)t.innerHTML=pestanasHTML();
  [].forEach.call(S.cont.querySelectorAll("[data-k]"),function(el){el.classList.remove("mal","ojo");var v=S.val.marcas[el.getAttribute("data-k")];if(v===1)el.classList.add("mal");else if(v===2)el.classList.add("ojo");});
  [].forEach.call(S.cont.querySelectorAll("[data-cuenta]"),function(el){el.textContent=pl(cuentaPalabras(Number(el.getAttribute("data-cuenta"))).total,"palabra","palabras");});
  pintarPrioridades();marcarMenu();
}
function pintarPrioridades(){
  var gs=S.modelo.grupos,max=1,vals=[];
  gs.forEach(function(g){if(typeof g.puntos==="number"&&isFinite(g.puntos)){if(g.puntos>max)max=g.puntos;vals.push(g.puntos);}});
  vals.sort(function(a,b){return b-a;});
  gs.forEach(function(g){
    var i=S.cont.querySelector('[data-barra="'+g.id+'"]'),r=S.cont.querySelector('[data-rango="'+g.id+'"]');
    if(!i)return;
    var ok=typeof g.puntos==="number"&&isFinite(g.puntos);
    i.style.width=ok?Math.max(2,Math.round(g.puntos/max*100))+"%":"0";
    r.innerHTML=ok?"<b>#"+(vals.indexOf(g.puntos)+1)+"</b> de "+gs.length:"—";
  });
}
function aplicarFiltros(silencioso){
  var t=S.cont.querySelector("#t-palabras");if(!t)return;
  var vis=0,tot=0;
  [].forEach.call(t.querySelectorAll("tbody tr[data-id]"),function(tr){
    var p=porId(S.modelo.palabras,Number(tr.getAttribute("data-id")));if(!p)return;tot++;
    var ok=coincidePalabra(p);tr.hidden=!ok;if(ok)vis++;
    else if(S.sel[p.id]){delete S.sel[p.id];var c=tr.querySelector("[data-sel]");if(c)c.checked=false;}
  });
  var c=S.cont.querySelector("#pal-cuenta");if(c)c.textContent=vis+" de "+tot+" palabras";
  var v=S.cont.querySelector("#pal-vacio");if(v)v.hidden=vis>0||tot===0;
  pintarSeleccion();
}
function pintarSeleccion(){
  var ids=Object.keys(S.sel).map(Number),b=S.cont.querySelector("#sel-barra");if(!b)return;
  if(!ids.length){b.innerHTML="";var a=S.cont.querySelector("#sel-todas");if(a)a.checked=false;return;}
  b.innerHTML='<div class="sel" role="group" aria-label="Acciones sobre la selección"><b>'+pl(ids.length,"seleccionada","seleccionadas")+'</b>'
   +'<select id="sel-dest" aria-label="Grupo de destino">'+opcionesGrupos(null,"Mover a…")+'</select><button class="boton sec peq" type="button" data-act="sel-mover">Mover</button>'
   +'<button class="boton sec peq" type="button" data-act="sel-activar">Activar</button><button class="boton sec peq" type="button" data-act="sel-desactivar">Desactivar</button>'
   +'<button class="boton sec peq peligro" type="button" data-act="sel-borrar">Eliminar</button><button class="boton sec peq" type="button" data-act="sel-quitar">Quitar selección</button></div>';
}

/* ====================================================================
   acciones
   ==================================================================== */
function err(id,txt){var e=S.cont.querySelector("#"+id);if(e)e.textContent=txt||"";}
function addGrupo(){
  var nombre=S.cont.querySelector("#ng-nombre").value.trim(),pts=num(S.cont.querySelector("#ng-puntos").value),desc=S.cont.querySelector("#ng-desc").value.trim();
  if(!nombre)return err("ng-err","Escribe el nombre del grupo.");
  if(S.modelo.grupos.some(function(g){return clave(g.nombre)===clave(nombre);}))return err("ng-err","Ya existe un grupo con ese nombre.");
  if(typeof pts!=="number"||pts<0)return err("ng-err","Los puntos deben ser un número igual o mayor que 0.");
  snap();S.modelo.grupos.push({id:nid(),nombre:nombre,puntos:pts,desc:desc});
  render();toast("Grupo «"+nombre+"» añadido. Ahora puedes darle palabras clave.",{deshacer:true});
  var i=S.cont.querySelector("#ng-nombre");if(i)i.focus();
}
function addPalabra(){
  var w=S.cont.querySelector("#np-pal").value.trim(),gid=Number(S.cont.querySelector("#np-gid").value),tipo=S.cont.querySelector("#np-tipo").value;
  if(!w)return err("np-err","Escribe la palabra o expresión.");
  if(existeEnGrupo(gid,w,tipo,-1))return err("np-err","«"+w+"» ya está en ese grupo.");
  snap();S.modelo.palabras.push({id:nid(),gid:gid,grupoTxt:"",palabra:w,tipo:tipo,activa:true});
  render();var i=S.cont.querySelector("#np-pal");if(i)i.focus();
  S.cont.querySelector("#np-gid").value=String(gid);
}
function addVarias(){
  var gid=Number(S.cont.querySelector("#nv-gid").value),tipo=S.cont.querySelector("#nv-tipo").value;
  var lineas=S.cont.querySelector("#nv-txt").value.split(/\r?\n/).map(function(x){return x.trim();}).filter(Boolean);
  if(!lineas.length)return err("nv-err","Escribe al menos una palabra.");
  snap();var nuevas=0,rep=0;
  lineas.forEach(function(w){if(existeEnGrupo(gid,w,tipo,-1)){rep++;return;}S.modelo.palabras.push({id:nid(),gid:gid,grupoTxt:"",palabra:w,tipo:tipo,activa:true});nuevas++;});
  if(!nuevas){S.deshacer.pop();return err("nv-err","Todas esas palabras ya estaban en el grupo.");}
  render();toast(pl(nuevas,"palabra añadida","palabras añadidas")+(rep?" ("+pl(rep,"ya estaba","ya estaban")+" y se ha omitido)":"")+".",{deshacer:true});
}
function addCpv(){
  var pre=S.cont.querySelector("#nc-pre").value.trim(),pts=num(S.cont.querySelector("#nc-puntos").value),desc=S.cont.querySelector("#nc-desc").value.trim();
  if(!/^\d{1,8}$/.test(pre))return err("nc-err","El prefijo debe tener solo cifras (de 1 a 8).");
  if(S.modelo.cpv.some(function(c){return String(c.prefijo).trim()===pre;}))return err("nc-err","Ese prefijo ya existe.");
  if(typeof pts!=="number"||pts<0)return err("nc-err","Los puntos deben ser un número igual o mayor que 0.");
  snap();S.modelo.cpv.push({id:nid(),prefijo:pre,puntos:pts,desc:desc});render();
  var i=S.cont.querySelector("#nc-pre");if(i)i.focus();
}
function addExcl(){
  var w=S.cont.querySelector("#nx-pal").value.trim(),tipo=S.cont.querySelector("#nx-tipo").value;
  if(!w)return err("nx-err","Escribe la palabra o expresión.");
  if(S.modelo.excl.some(function(x){return clave(x.palabra)===clave(w)&&x.tipo===tipo;}))return err("nx-err","Esa exclusión ya existe.");
  snap();S.modelo.excl.push({id:nid(),palabra:w,tipo:tipo,activa:true});render();
  var i=S.cont.querySelector("#nx-pal");if(i)i.focus();
}
async function pedirBorrarGrupo(id){
  var g=grupoDe(id);if(!g)return;
  var ps=S.modelo.palabras.filter(function(p){return p.gid===id;}),otros=S.modelo.grupos.filter(function(x){return x.id!==id;});
  var cuerpo='<p>Vas a eliminar el grupo <b>«'+esc(g.nombre)+'»</b> ('+(typeof g.puntos==="number"?g.puntos+" puntos":"sin puntos")+').</p>';
  if(!ps.length){
    cuerpo+='<p>No tiene palabras clave asociadas, así que no hay nada más que decidir.</p>';
    var r0=await modal("Eliminar grupo",cuerpo,[{txt:"Cancelar",val:"no",cls:"sec"},{txt:"Eliminar grupo",val:"si",cls:"peligro"}]);
    if(r0.val!=="si")return;
    var x0=eliminarGrupo(id,"borrar",null);render();toast("Grupo «"+x0.g.nombre+"» eliminado.",{deshacer:true});return;
  }
  cuerpo+='<p>Tiene <b>'+pl(ps.length,"palabra clave","palabras clave")+'</b> ('+pl(ps.filter(function(p){return p.activa;}).length,"activa","activas")+'). Los códigos CPV y las exclusiones no dependen de ningún grupo y no se tocan.</p>'
   +'<fieldset><legend>¿Qué hago con sus palabras?</legend>'
   +(otros.length?'<label class="op"><input type="radio" name="modo" value="mover" checked> Pasarlas a otro grupo:<select name="destino" aria-label="Grupo de destino"><option value="">Elige un grupo…</option>'+otros.map(function(o){return '<option value="'+o.id+'">'+esc(o.nombre)+'</option>';}).join("")+'</select></label>':'')
   +'<label class="op"><input type="radio" name="modo" value="borrar"'+(otros.length?'':' checked')+'> Eliminarlas también</label></fieldset>'
   +'<p class="err" role="alert"></p>'
   +'<details><summary>Ver las palabras</summary><div class="pals">'+ps.map(function(p){return '<span>'+esc(p.palabra)+'</span>';}).join("")+'</div></details>';
  var r=await modal("Eliminar grupo",cuerpo,[{txt:"Cancelar",val:"no",cls:"sec"},{txt:"Eliminar grupo",val:"si",cls:"peligro"}],function(d){
    var sel=d.querySelector('select[name="destino"]'),mover=d.querySelector('input[value="mover"]');
    if(sel)sel.addEventListener("change",function(){mover.checked=true;d.querySelector(".err").textContent="";});
    /* el destino se elige a propósito: sin destino no se cierra el diálogo */
    d.querySelector('button[value="si"]').addEventListener("click",function(e){
      if(mover&&mover.checked&&sel&&!sel.value){e.preventDefault();d.querySelector(".err").textContent="Elige a qué grupo pasan las palabras, o marca «Eliminarlas también».";sel.focus();}
    });
  });
  if(r.val!=="si")return;
  var modo=r.datos.modo==="borrar"?"borrar":"mover",dest=modo==="mover"?Number(r.datos.destino):null;
  var x=eliminarGrupo(id,modo,dest),msg="Grupo «"+x.g.nombre+"» eliminado";
  if(modo==="mover"){var d=grupoDe(dest);msg+="; "+pl(x.r.movidas,"palabra pasada","palabras pasadas")+" a «"+(d?d.nombre:"")+"»"+(x.r.fusionadas?" ("+pl(x.r.fusionadas,"ya existía","ya existían")+" allí y se han fusionado)":"");}
  else msg+=" junto con "+pl(x.r.borradas,"palabra","palabras");
  S.sel={};render();toast(msg+".",{deshacer:true});
}
function idsSeleccion(){return Object.keys(S.sel).map(Number).filter(function(id){return porId(S.modelo.palabras,id);});}

function accion(a,id,el){
  var m=S.modelo;
  switch(a){
    case "elegir":elegirArchivo();break;
    case "reconectar":reconectar();break;
    case "guardar":guardar();break;
    case "deshacer":deshacer();break;
    case "descartar":descartar();break;
    case "releer":releer();break;
    case "add-grupo":addGrupo();break;
    case "add-palabra":addPalabra();break;
    case "add-varias":addVarias();break;
    case "add-cpv":addCpv();break;
    case "add-excl":addExcl();break;
    case "del-grupo":pedirBorrarGrupo(id);break;
    case "del-palabra":{var p=porId(m.palabras,id);if(!p)break;snap();m.palabras=m.palabras.filter(function(x){return x.id!==id;});delete S.sel[id];render();toast("Palabra «"+p.palabra+"» eliminada.",{deshacer:true});break;}
    case "del-cpv":{var c=porId(m.cpv,id);if(!c)break;snap();m.cpv=m.cpv.filter(function(x){return x.id!==id;});render();toast("CPV «"+c.prefijo+"» eliminado.",{deshacer:true});break;}
    case "del-excl":{var x=porId(m.excl,id);if(!x)break;snap();m.excl=m.excl.filter(function(y){return y.id!==id;});render();toast("Exclusión «"+x.palabra+"» eliminada.",{deshacer:true});break;}
    case "ordenar-grupos":snap();m.grupos.sort(function(a,b){return (typeof b.puntos==="number"?b.puntos:-1)-(typeof a.puntos==="number"?a.puntos:-1);});render();toast("Grupos ordenados por prioridad.",{deshacer:true});break;
    case "ver-palabras":S.filtro={gid:String(id),q:"",tipo:"",act:""};break; /* el enlace cambia de pestaña */
    case "quitar-filtros":S.filtro={gid:"",q:"",tipo:"",act:""};render();break;
    case "sel-quitar":S.sel={};[].forEach.call(S.cont.querySelectorAll("[data-sel]"),function(c){c.checked=false;});pintarSeleccion();break;
    case "sel-mover":{
      var ids=idsSeleccion(),dest=Number(S.cont.querySelector("#sel-dest").value);
      if(!dest){toast("Elige primero el grupo de destino.",{mal:true});break;}
      snap();var r=moverPalabras(ids,dest),g=grupoDe(dest);S.sel={};render();
      toast(pl(r.movidas,"palabra movida","palabras movidas")+" a «"+g.nombre+"»"+(r.fusionadas?" ("+pl(r.fusionadas,"ya existía","ya existían")+" y se han fusionado)":"")+".",{deshacer:true});break;}
    case "sel-activar":case "sel-desactivar":{
      var v=a==="sel-activar";snap();idsSeleccion().forEach(function(i){porId(m.palabras,i).activa=v;});render();break;}
    case "sel-borrar":{
      var bs=idsSeleccion();
      modal("Eliminar palabras","<p>Se eliminarán <b>"+pl(bs.length,"palabra clave","palabras clave")+"</b>.</p>",[{txt:"Cancelar",val:"no",cls:"sec"},{txt:"Eliminar",val:"si",cls:"peligro"}]).then(function(r){
        if(r.val!=="si")return;snap();m.palabras=m.palabras.filter(function(p){return bs.indexOf(p.id)<0;});S.sel={};render();toast(pl(bs.length,"palabra eliminada","palabras eliminadas")+".",{deshacer:true});});
      break;}
  }
}
/* edición en línea: cada cambio se aplica al modelo y la cascada es automática por id */
function cambio(el){
  var f=el.getAttribute("data-f"),id=Number(el.getAttribute("data-id")),m=S.modelo,o,v;
  function set(obj,campo,val){if(obj[campo]!==val){snap();obj[campo]=val;}}
  switch(f){
    case "g.nombre":o=porId(m.grupos,id);set(o,"nombre",el.value.trim());break;
    case "g.puntos":o=porId(m.grupos,id);set(o,"puntos",num(el.value));break;
    case "g.desc":o=porId(m.grupos,id);set(o,"desc",el.value.trim());break;
    case "p.palabra":o=porId(m.palabras,id);set(o,"palabra",el.value.trim());break;
    case "p.gid":o=porId(m.palabras,id);v=el.value===""?null:Number(el.value);if(v!==null){var antes=o.gid;if(antes!==v){snap();o.gid=v;o.grupoTxt="";}}break;
    case "p.tipo":o=porId(m.palabras,id);set(o,"tipo",el.value);break;
    case "p.activa":o=porId(m.palabras,id);set(o,"activa",el.checked);break;
    case "c.prefijo":o=porId(m.cpv,id);set(o,"prefijo",el.value.trim());break;
    case "c.puntos":o=porId(m.cpv,id);set(o,"puntos",num(el.value));break;
    case "c.desc":o=porId(m.cpv,id);set(o,"desc",el.value.trim());break;
    case "x.palabra":o=porId(m.excl,id);set(o,"palabra",el.value.trim());break;
    case "x.tipo":o=porId(m.excl,id);set(o,"tipo",el.value);break;
    case "x.activa":o=porId(m.excl,id);set(o,"activa",el.checked);break;
    case "a.valor":o=porId(m.ajustes,id);set(o,"valor",num(el.value));break;
    case "a.texto":o=porId(m.ajustes,id);set(o,"valor",el.value.trim());break;
  }
  if(f==="p.gid")aplicarFiltros();
  refrescar();
}
function enlazar(cont){
  if(cont._cfgOn)return;cont._cfgOn=true;
  function dentro(e){return e.target.closest&&e.target.closest(".cfg");}
  cont.addEventListener("click",function(e){
    if(!dentro(e))return;
    var b=e.target.closest("[data-act]");if(!b)return;
    var a=b.getAttribute("data-act");
    if(a!=="ver-palabras")e.preventDefault();
    accion(a,Number(b.getAttribute("data-id")),b);
  });
  cont.addEventListener("change",function(e){
    if(!dentro(e))return;
    var t=e.target;
    if(t.id==="cfg-file"){var f=t.files&&t.files[0];if(!f)return;confirmarDescarte().then(function(ok){if(!ok){t.value="";return;}f.arrayBuffer().then(function(buf){cargar(buf,f.name,null,f.lastModified);});});return;}
    if(t.hasAttribute("data-sel")){var id=Number(t.getAttribute("data-sel"));if(t.checked)S.sel[id]=1;else delete S.sel[id];pintarSeleccion();return;}
    if(t.id==="sel-todas"){[].forEach.call(S.cont.querySelectorAll("#t-palabras tbody tr:not([hidden]) [data-sel]"),function(c){c.checked=t.checked;var id=Number(c.getAttribute("data-sel"));if(t.checked)S.sel[id]=1;else delete S.sel[id];});pintarSeleccion();return;}
    if(t.id==="fl-gid"){S.filtro.gid=t.value;aplicarFiltros();var g=S.cont.querySelector("#np-gid"),g2=S.cont.querySelector("#nv-gid");if(/^\d+$/.test(t.value)){if(g)g.value=t.value;if(g2)g2.value=t.value;}return;}
    if(t.id==="fl-tipo"){S.filtro.tipo=t.value;aplicarFiltros();return;}
    if(t.id==="fl-act"){S.filtro.act=t.value;aplicarFiltros();return;}
    if(t.hasAttribute("data-f"))cambio(t);
  });
  cont.addEventListener("input",function(e){
    if(!dentro(e))return;
    if(e.target.id==="fl-q"){S.filtro.q=e.target.value;aplicarFiltros();}
  });
  cont.addEventListener("keydown",function(e){
    if(!dentro(e)||e.key!=="Enter")return;
    var t=e.target,mapa={"ng-nombre":"add-grupo","ng-puntos":"add-grupo","ng-desc":"add-grupo","np-pal":"add-palabra","nc-pre":"add-cpv","nc-puntos":"add-cpv","nc-desc":"add-cpv","nx-pal":"add-excl"};
    if(mapa[t.id]){e.preventDefault();accion(mapa[t.id]);}
    else if(t.hasAttribute&&t.hasAttribute("data-f")&&t.tagName==="INPUT"&&t.type!=="checkbox"){e.preventDefault();t.blur();}
  });
}

/* ====================================================================
   punto de entrada
   ==================================================================== */
window.addEventListener("beforeunload",function(e){if(S.sucio){e.preventDefault();e.returnValue="";}});
window.PanelConfig={
  mostrar:function(cont,tab){
    S.cont=cont;S.tab=TABS.some(function(t){return t.id===tab;})?tab:"grupos";
    enlazar(cont);render();iniciar();
  },
  _estado:S /* solo para pruebas */
};
})();
