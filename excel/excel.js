/* Excel editor (gaya Excel 2013). Mengekspos Apps.excel.mount(root,{file,onHome}) */
(function(){
'use strict';
const CDN={exceljs:'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js'};
const libs={};
function lib(n,test){
  if(test()) return Promise.resolve();
  return libs[n]||(libs[n]=new Promise((ok,err)=>{const s=document.createElement('script');s.src=CDN[n];s.onload=ok;s.onerror=()=>{delete libs[n];err(new Error('Gagal memuat '+n));};document.head.appendChild(s);}));
}

const MAXR=1048576,MAXC=16384,DW=64,DH=20,HW=40,HH=20;
const FONTS=['Calibri','Arial','Times New Roman','Cambria','Georgia','Verdana','Courier New','Segoe UI','Tahoma','Trebuchet MS'];
const SIZES=[8,9,10,11,12,14,16,18,20,22,24,26,28,36,48,72];

/* ---------- alamat ---------- */
const colName=c=>{let s='';c++;while(c>0){const m=(c-1)%26;s=String.fromCharCode(65+m)+s;c=(c-1-m)/26;}return s;};
const colIdx=s=>{let n=0;for(const ch of s.toUpperCase())n=n*26+ch.charCodeAt(0)-64;return n-1;};
const addr=(r,c)=>colName(c)+(r+1);
const parseAddr=s=>{const m=/^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(s.trim());if(!m)return null;const c=colIdx(m[1]),r=+m[2]-1;return r>=0&&r<MAXR&&c>=0&&c<MAXC?{r,c}:null;};
const key=(r,c)=>r*MAXC+c;
const esc=t=>String(t).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

/* ---------- galat ---------- */
class XErr{constructor(c){this.code=c;}toString(){return this.code;}}
const E={DIV:new XErr('#DIV/0!'),VAL:new XErr('#VALUE!'),REF:new XErr('#REF!'),NAME:new XErr('#NAME?'),NA:new XErr('#N/A'),NUM:new XErr('#NUM!'),CIRC:new XErr('#REF!')};
const isErr=v=>v instanceof XErr;

/* ---------- tanggal ---------- */
const EPOCH=Date.UTC(1899,11,30);
const dateToSerial=d=>(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate(),d.getHours(),d.getMinutes(),d.getSeconds())-EPOCH)/864e5;
const serialToDate=n=>{const ms=Math.round(n*864e5)+EPOCH,d=new Date(ms);return new Date(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate(),d.getUTCHours(),d.getUTCMinutes(),d.getUTCSeconds());};
const MON=['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
const DAY=['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];

/* ---------- format angka ---------- */
const FMT={
  General:'General',Number:'0.00',Currency:'"Rp"#,##0',Accounting:'_("Rp"* #,##0_);_("Rp"* (#,##0);_("Rp"* "-"_);_(@_)',
  ShortDate:'dd/mm/yyyy',LongDate:'dddd, d mmmm yyyy',Time:'hh:mm:ss',Percent:'0%',Fraction:'# ?/?',Sci:'0.00E+00',Text:'@',Comma:'#,##0.00'
};
const FMTLIST=[['General','Umum'],['Number','Angka'],['Currency','Mata uang'],['Accounting','Akuntansi'],['ShortDate','Tanggal singkat'],['LongDate','Tanggal panjang'],['Time','Waktu'],['Percent','Persentase'],['Fraction','Pecahan'],['Sci','Ilmiah'],['Text','Teks']];
const fmtKey=code=>{for(const k in FMT)if(FMT[k]===code)return k;return null;};
function splitSections(code){const out=[];let cur='',q=false;for(let i=0;i<code.length;i++){const ch=code[i];if(ch==='"')q=!q;if(ch===';'&&!q){out.push(cur);cur='';}else cur+=ch;}out.push(cur);return out;}
function addCommas(s){return s.replace(/\B(?=(\d{3})+(?!\d))/g,'.');}  /* pemisah ribuan lokal Indonesia: titik */
function fmtNumber(v,sec){
  let code=sec.replace(/\[(?:Red|Merah|Black|Blue|Green|Color\d+)\]/gi,'').replace(/_./g,' ').replace(/\*./g,'');
  /* pisahkan literal */
  const toks=[];let i=0;
  while(i<code.length){const ch=code[i];
    if(ch==='"'){const j=code.indexOf('"',i+1);toks.push({t:'lit',v:code.slice(i+1,j<0?code.length:j)});i=j<0?code.length:j+1;}
    else if(ch==='\\'){toks.push({t:'lit',v:code[i+1]||''});i+=2;}
    else{toks.push({t:'ch',v:ch});i++;}}
  const body=toks.map(t=>t.t==='lit'?'\u0001':t.v).join('');
  const lits=toks.filter(t=>t.t==='lit').map(t=>t.v);let li=0;
  const pct=(body.match(/%/g)||[]).length;if(pct)v*=Math.pow(100,pct);
  const sci=/0\.?0*E[+-]0+/i.exec(body);
  let res;
  if(sci){
    const m=/^(.*?)(#*0+)(?:\.(0+|#+))?E([+-])(0+)(.*)$/i.exec(body.replace(/,/g,''));
    const dec=m&&m[3]?m[3].length:0,ed=m?m[5].length:2;
    let e=v===0?0:Math.floor(Math.log10(Math.abs(v))),mant=v/Math.pow(10,e);
    let ms=Math.abs(mant).toFixed(dec);if(+ms>=10){e++;ms=(Math.abs(v)/Math.pow(10,e)).toFixed(dec);}
    res=(v<0?'-':'')+ms.replace('.',',')+'E'+(e<0?'-':'+')+String(Math.abs(e)).padStart(ed,'0');
    return res;
  }
  if(/\?\/[?\d]+/.test(body)){
    const w=Math.trunc(Math.abs(v)),fr=Math.abs(v)-w;let bn=0,bd=1,be=1;
    for(let d=1;d<=99;d++){const n=Math.round(fr*d),er=Math.abs(fr-n/d);if(er<be-1e-12){be=er;bn=n;bd=d;}if(er<1e-9)break;}
    const sign=v<0?'-':'';
    if(bn===0)return sign+w;
    return sign+(w?w+' ':'')+bn+'/'+bd;
  }
  /* bagian bilangan */
  const mm=/[#0?,.]+/.exec(body);
  if(!mm)return body.replace(/\u0001/g,()=>lits[li++]||'');
  const pat=mm[0],hasComma=/,/.test(pat.replace(/,+$/,'')),trail=(/,+$/.exec(pat)||[''])[0].length;
  const clean=pat.replace(/,/g,''),dot=clean.indexOf('.');
  const ip=dot<0?clean:clean.slice(0,dot),fp=dot<0?'':clean.slice(dot+1);
  let n=Math.abs(v)/Math.pow(1000,trail);
  let s=n.toFixed(fp.length);
  let [is,fs='']=s.split('.');
  const minInt=(ip.match(/0/g)||[]).length;
  if(is==='0'&&minInt===0&&!/0/.test(ip))is='';
  is=is.padStart(minInt,'0');
  /* hapus nol akhir untuk # */
  let f2=fs;for(let k=fp.length-1;k>=0&&f2.length>k;k--){if(fp[k]!=='0'&&f2[k]==='0'&&k===f2.length-1)f2=f2.slice(0,-1);else break;}
  fs=f2;
  if(hasComma)is=addCommas(is);
  let num=is+(fs?','+fs:(fp.includes('0')?'':''));
  if(!is&&!fs)num='';
  let out=body.slice(0,mm.index)+num+body.slice(mm.index+pat.length);
  out=out.replace(/%/g,'%').replace(/\u0001/g,()=>lits[li++]||'');
  return (v<0&&+n.toFixed(fp.length)!==0?'-':'')+out;
}
function fmtDate(v,code){
  const d=serialToDate(v),Y=d.getFullYear(),M=d.getMonth(),D=d.getDate(),h=d.getHours(),mi=d.getMinutes(),s=d.getSeconds(),wd=d.getDay();
  const ampm=/AM\/PM|A\/P/i.test(code);
  let out='',i=0,lastH=false;
  const p2=n=>String(n).padStart(2,'0');
  while(i<code.length){
    const rest=code.slice(i);let m;
    if(code[i]==='"'){const j=code.indexOf('"',i+1);out+=code.slice(i+1,j<0?code.length:j);i=j<0?code.length:j+1;continue;}
    if(code[i]==='\\'){out+=code[i+1]||'';i+=2;continue;}
    if(code[i]==='_'||code[i]==='*'){i+=2;continue;}
    if(m=/^\[(?:Red|Black|Blue|Green)\]/i.exec(rest)){i+=m[0].length;continue;}
    if(m=/^yyyy/i.exec(rest)){out+=Y;i+=4;lastH=false;}
    else if(m=/^yy/i.exec(rest)){out+=p2(Y%100);i+=2;lastH=false;}
    else if(m=/^mmmmm/i.exec(rest)){out+=MON[M][0];i+=5;}
    else if(m=/^mmmm/i.exec(rest)){out+=MON[M];i+=4;}
    else if(m=/^mmm/i.exec(rest)){out+=MON[M].slice(0,3);i+=3;}
    else if(m=/^mm/i.exec(rest)){
      const prev=out.length&&lastH,nextS=/^mm[^a-z]*s/i.test(rest);
      out+=(prev||nextS)?p2(mi):p2(M+1);i+=2;lastH=false;}
    else if(m=/^m/i.exec(rest)){const prev=lastH,nextS=/^m[^a-z]*s/i.test(rest);out+=(prev||nextS)?mi:(M+1);i+=1;lastH=false;}
    else if(m=/^dddd/i.exec(rest)){out+=DAY[wd];i+=4;}
    else if(m=/^ddd/i.exec(rest)){out+=DAY[wd].slice(0,3);i+=3;}
    else if(m=/^dd/i.exec(rest)){out+=p2(D);i+=2;}
    else if(m=/^d/i.exec(rest)){out+=D;i+=1;}
    else if(m=/^hh/i.exec(rest)){out+=p2(ampm?(h%12||12):h);i+=2;lastH=true;}
    else if(m=/^h/i.exec(rest)){out+=(ampm?(h%12||12):h);i+=1;lastH=true;}
    else if(m=/^ss/i.exec(rest)){out+=p2(s);i+=2;lastH=false;}
    else if(m=/^s/i.exec(rest)){out+=s;i+=1;lastH=false;}
    else if(m=/^AM\/PM/i.exec(rest)){out+=h<12?'AM':'PM';i+=5;}
    else if(m=/^A\/P/i.exec(rest)){out+=h<12?'A':'P';i+=3;}
    else{out+=code[i];i++;}
  }
  return out;
}
const isDateCode=c=>/(^|[^\\"a-z])(y|d|h|s|m)/i.test(c.replace(/"[^"]*"/g,'').replace(/\[[^\]]*\]/g,'').replace(/\\./g,''))&&!/^General$/i.test(c)&&!/[#0?]/.test(c.replace(/"[^"]*"/g,'').replace(/\\./g,'').replace(/\[[^\]]*\]/g,''));
function fmtValue(v,code){
  if(isErr(v))return {t:v.code,al:'center'};
  if(v===null||v===undefined||v==='')return {t:''};
  if(typeof v==='boolean')return {t:v?'TRUE':'FALSE',al:'center'};
  code=code||'General';
  if(typeof v==='string'){
    const secs=splitSections(code),sec=secs[secs.length>=4?3:secs.length>1&&/@/.test(secs[secs.length-1])?secs.length-1:0];
    if(/@/.test(code)&&code!=='General'){return {t:(secs.find(s=>/@/.test(s))||'@').replace(/"([^"]*)"/g,'$1').replace(/_./g,' ').replace(/\*./g,'').replace(/@/g,v)};}
    return {t:v};
  }
  if(code==='General'){
    let t;
    if(Math.abs(v)>=1e11||(Math.abs(v)<1e-9&&v!==0)){t=v.toExponential(5).replace(/\.?0+e/,'E').replace('e','E').replace(/E([+-])(\d)$/,'E$10$2').replace('.',',');}
    else{t=String(+v.toPrecision(11));if(t.includes('e'))t=v.toExponential(5).replace(/\.?0+e/,'E').replace('e','E');t=t.replace('.',',');}
    return {t,al:'right'};
  }
  const secs=splitSections(code);
  let sec=secs[0];
  if(secs.length>1&&v<0)sec=secs[1];
  if(secs.length>2&&v===0)sec=secs[2];
  const useAbs=secs.length>1&&v<0;
  let t;
  if(isDateCode(sec))t=fmtDate(v,sec);
  else t=fmtNumber(useAbs?Math.abs(v):v,sec);
  return {t,al:'right'};
}

/* ---------- koersi ---------- */
const num=v=>{
  if(isErr(v))return v;
  if(v===null||v===undefined||v==='')return 0;
  if(typeof v==='number')return v;
  if(typeof v==='boolean')return v?1:0;
  const s=String(v).trim();
  if(/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(s))return parseFloat(s);
  if(/^[-+]?\d+,\d+$/.test(s))return parseFloat(s.replace(',','.'));
  if(/^[-+]?\d+(\.\d{3})*(,\d+)?%$/.test(s))return parseFloat(s.replace(/\./g,'').replace(',','.'))/100;
  return E.VAL;
};
const str=v=>{
  if(isErr(v))return v;
  if(v===null||v===undefined)return '';
  if(typeof v==='boolean')return v?'TRUE':'FALSE';
  if(typeof v==='number')return fmtValue(v,'General').t;
  return String(v);
};
const truthy=v=>{if(isErr(v))return v;if(typeof v==='string'){const u=v.toUpperCase();if(u==='TRUE')return true;if(u==='FALSE')return false;return E.VAL;}return !!num(v);};
/* tebak tipe dari masukan pengguna */
function parseInput(t){
  if(t==='')return {v:null};
  const s=t.trim();
  if(/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(s))return {v:parseFloat(s)};
  if(/^[-+]?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s))return {v:parseFloat(s.replace(/\./g,'').replace(',','.')),fmt:'#,##0'+(s.includes(',')?'.'+'0'.repeat(s.split(',')[1].length):'')};
  if(/^[-+]?\d+,\d+$/.test(s))return {v:parseFloat(s.replace(',','.'))};
  let m;
  if(m=/^([-+]?\d+(?:[.,]\d+)?)\s*%$/.exec(s))return {v:parseFloat(m[1].replace(',','.'))/100,fmt:/[.,]/.test(m[1])?'0.00%':'0%'};
  if(m=/^(?:Rp\.?\s*)([-]?\d{1,3}(?:\.\d{3})*(?:,\d+)?|\d+)$/i.exec(s))return {v:parseFloat(m[1].replace(/\./g,'').replace(',','.')),fmt:FMT.Currency};
  if(/^(true|false)$/i.test(s))return {v:s.toUpperCase()==='TRUE'};
  if(m=/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/.exec(s)){
    let d=+m[1],mo=+m[2],y=+m[3];if(y<100)y+=y<30?2000:1900;
    const dt=new Date(y,mo-1,d);if(dt.getMonth()===mo-1&&dt.getDate()===d)return {v:dateToSerial(dt),fmt:FMT.ShortDate};
  }
  if(m=/^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s)){const dt=new Date(+m[1],+m[2]-1,+m[3]);if(dt.getMonth()===+m[2]-1)return {v:dateToSerial(dt),fmt:FMT.ShortDate};}
  if(m=/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(s)){const h=+m[1],mi=+m[2],se=+(m[3]||0);if(h<24&&mi<60&&se<60)return {v:(h*3600+mi*60+se)/86400,fmt:m[3]?'hh:mm:ss':'hh:mm'};}
  return {v:t};
}

/* ---------- mesin rumus ---------- */
const FN_NAMES=new Set();
function tokenize(src){
  const out=[];let i=0;const n=src.length;
  while(i<n){
    const ch=src[i];
    if(/\s/.test(ch)){i++;continue;}
    if(ch==='"'){let j=i+1,s='';while(j<n){if(src[j]==='"'){if(src[j+1]==='"'){s+='"';j+=2;continue;}break;}s+=src[j++];}out.push({t:'str',v:s});i=j+1;continue;}
    let m=/^(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!/.exec(src.slice(i));
    let sheet=null;
    if(m){sheet=(m[1]!==undefined?m[1].replace(/''/g,"'"):m[2]);i+=m[0].length;}
    const rest=src.slice(i);
    if(m=/^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?(?![\w(])/.exec(rest)){out.push({t:'ref',v:m[0],sheet});i+=m[0].length;continue;}
    if(m=/^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![\w(])/.exec(rest)){out.push({t:'ref',v:m[0],sheet});i+=m[0].length;continue;}
    if(m=/^\$?\d+:\$?\d+(?![\w(])/.exec(rest)){out.push({t:'ref',v:m[0],sheet});i+=m[0].length;continue;}
    if(sheet!==null){out.push({t:'err',v:'#REF!'});continue;}
    if(m=/^(\d+\.?\d*|\.\d+)(e[-+]?\d+)?/i.exec(rest)){out.push({t:'num',v:parseFloat(m[0])});i+=m[0].length;continue;}
    if(m=/^#(?:REF!|N\/A|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!)/i.exec(rest)){out.push({t:'err',v:m[0].toUpperCase()});i+=m[0].length;continue;}
    if(m=/^[A-Za-z_\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*/.exec(rest)){
      const id=m[0];i+=id.length;
      if(src[i]==='(')out.push({t:'fn',v:id.toUpperCase()});
      else if(/^(TRUE|FALSE)$/i.test(id))out.push({t:'bool',v:id.toUpperCase()==='TRUE'});
      else out.push({t:'name',v:id});
      continue;}
    if(rest.startsWith('<>')||rest.startsWith('>=')||rest.startsWith('<=')){out.push({t:'op',v:rest.slice(0,2)});i+=2;continue;}
    if('+-*/^&=<>%'.includes(ch)){out.push({t:'op',v:ch});i++;continue;}
    if(ch==='('||ch===')'){out.push({t:ch});i++;continue;}
    if(ch===','||ch===';'){out.push({t:','});i++;continue;}
    throw new Error('sintaks');
  }
  return out;
}
function parseRefTok(t){
  const p=t.v.replace(/\$/g,'');
  const parts=p.split(':');
  const one=s=>{let m;
    if(m=/^([A-Za-z]+)(\d+)$/.exec(s))return {r:+m[2]-1,c:colIdx(m[1])};
    if(m=/^([A-Za-z]+)$/.exec(s))return {r:null,c:colIdx(m[1])};
    if(m=/^(\d+)$/.exec(s))return {r:+m[1]-1,c:null};};
  const a=one(parts[0]),b=parts[1]?one(parts[1]):a;
  const r1=a.r===null?0:a.r,r2=b.r===null?MAXR-1:b.r,c1=a.c===null?0:a.c,c2=b.c===null?MAXC-1:b.c;
  return {sheet:t.sheet,r1:Math.min(r1,r2),r2:Math.max(r1,r2),c1:Math.min(c1,c2),c2:Math.max(c1,c2),single:!parts[1]};
}
function parseFormula(src){
  const tk=tokenize(src);let p=0;
  const peek=()=>tk[p],next=()=>tk[p++];
  const isOp=(v)=>peek()&&peek().t==='op'&&(Array.isArray(v)?v.includes(peek().v):peek().v===v);
  function cmp(){let l=cat();while(isOp(['=','<>','<','>','<=','>='])){const o=next().v;l={k:'bin',o,l,r:cat()};}return l;}
  function cat(){let l=add();while(isOp('&')){next();l={k:'bin',o:'&',l,r:add()};}return l;}
  function add(){let l=mul();while(isOp(['+','-'])){const o=next().v;l={k:'bin',o,l,r:mul()};}return l;}
  function mul(){let l=pow();while(isOp(['*','/'])){const o=next().v;l={k:'bin',o,l,r:pow()};}return l;}
  function pow(){let l=unary();while(isOp('^')){next();l={k:'bin',o:'^',l,r:unary()};}return l;}
  function unary(){if(isOp(['-','+'])){const o=next().v;return {k:'un',o,e:unary()};}return post();}
  function post(){let e=prim();while(isOp('%')){next();e={k:'pct',e};}return e;}
  function prim(){
    const t=next();if(!t)throw new Error('sintaks');
    if(t.t==='num')return {k:'num',v:t.v};
    if(t.t==='str')return {k:'str',v:t.v};
    if(t.t==='bool')return {k:'bool',v:t.v};
    if(t.t==='err')return {k:'err',v:t.v};
    if(t.t==='ref')return {k:'ref',...parseRefTok(t)};
    if(t.t==='name')return {k:'name',v:t.v};
    if(t.t==='('){const e=cmp();if(!peek()||next().t!==')')throw new Error('sintaks');return e;}
    if(t.t==='fn'){
      next();const args=[];
      if(peek()&&peek().t===')'){next();return {k:'fn',n:t.v,a:args};}
      for(;;){
        if(peek()&&(peek().t===','||peek().t===')'))args.push({k:'empty'});else args.push(cmp());
        const x=next();if(!x)throw new Error('sintaks');
        if(x.t===')')break;if(x.t!==',')throw new Error('sintaks');
      }
      return {k:'fn',n:t.v,a:args};
    }
    throw new Error('sintaks');
  }
  const ast=cmp();if(p<tk.length)throw new Error('sintaks');return ast;
}

/* ---------- penulisan ulang referensi (tempel, sisip, hapus) ---------- */
function rewriteRefs(src,fn){
  let out='',i=0;const n=src.length;
  while(i<n){
    const ch=src[i];
    if(ch==='"'){let j=i+1;while(j<n){if(src[j]==='"'){if(src[j+1]==='"'){j+=2;continue;}break;}j++;}out+=src.slice(i,j+1);i=j+1;continue;}
    const rest=src.slice(i);let m;
    const prevWord=out.length&&/[\w.]$/.test(out);
    if(!prevWord&&(m=/^((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?(?![\w(])/.exec(rest))){
      const sh=m[1]?m[1].slice(0,-1).replace(/^'|'$/g,'').replace(/''/g,"'"):null;
      const a={r:+m[5]-1,c:colIdx(m[3]),ar:!!m[4],ac:!!m[2]};
      const b=m[7]?{r:+m[9]-1,c:colIdx(m[7]),ar:!!m[8],ac:!!m[6]}:null;
      const res=fn(a,b,sh);
      if(res===null)out+=(m[1]||'')+'#REF!';
      else{const f=(x)=>(x.ac?'$':'')+colName(x.c)+(x.ar?'$':'')+(x.r+1);out+=(m[1]||'')+f(res[0])+(res[1]?':'+f(res[1]):'');}
      i+=m[0].length;continue;
    }
    if(/[A-Za-z_]/.test(ch)){const w=/^[A-Za-z_][\w.]*/.exec(rest)[0];out+=w;i+=w.length;continue;}
    if(/\d/.test(ch)){const w=/^[\d.]+(e[-+]?\d+)?/i.exec(rest)[0];out+=w;i+=w.length;continue;}
    out+=ch;i++;
  }
  return out;
}
function shiftFormula(src,dr,dc){
  return rewriteRefs(src,(a,b)=>{
    const mv=x=>({r:x.ar?x.r:x.r+dr,c:x.ac?x.c:x.c+dc,ar:x.ar,ac:x.ac});
    const A=mv(a),B=b?mv(b):null;
    if(A.r<0||A.c<0||A.r>=MAXR||A.c>=MAXC||(B&&(B.r<0||B.c<0||B.r>=MAXR||B.c>=MAXC)))return null;
    return [A,B];});
}

/* ---------- evaluator ---------- */
class Rng{constructor(sh,r1,c1,r2,c2){this.sh=sh;this.r1=r1;this.c1=c1;this.r2=r2;this.c2=c2;}
  get rows(){return this.r2-this.r1+1;}get cols(){return this.c2-this.c1+1;}}
const PASS={n:0};
function rawVal(wb,sh,r,c){
  const cell=sh.cells.get(key(r,c));if(!cell)return null;
  if(cell.f!==undefined&&cell.f!==null)return calcCell(wb,sh,cell,r,c);
  return cell.v===undefined?null:cell.v;
}
function calcCell(wb,sh,cell,r,c){
  if(cell.pass===PASS.n)return cell.cv;
  if(cell.busy){wb.circ=true;return 0;}
  cell.busy=true;
  let v;
  try{
    if(cell.ast===undefined){try{cell.ast=parseFormula(cell.f);}catch(e){cell.ast=null;}}
    v=cell.ast===null?E.NAME:evalAst(cell.ast,{wb,sh,cell:{r,c}});
    if(v instanceof Rng)v=v.rows*v.cols===1?rawVal(wb,v.sh,v.r1,v.c1):E.VAL;
    if(v===null||v===undefined)v=0;
  }catch(e){v=E.VAL;}
  cell.busy=false;cell.cv=v;cell.pass=PASS.n;
  return v;
}
function* cellsIn(rg){
  const sh=rg.sh,area=rg.rows*rg.cols;
  if(area>sh.cells.size*2){
    for(const [k,cell] of sh.cells){const r=Math.floor(k/MAXC),c=k%MAXC;if(r>=rg.r1&&r<=rg.r2&&c>=rg.c1&&c<=rg.c2)yield [r,c];}
  }else for(let r=rg.r1;r<=rg.r2;r++)for(let c=rg.c1;c<=rg.c2;c++)yield [r,c];
}
function evalAst(a,cx){
  switch(a.k){
    case 'num':case 'str':case 'bool':return a.v;
    case 'empty':return null;
    case 'err':return new XErr(a.v);
    case 'name':return E.NAME;
    case 'ref':{
      let sh=cx.sh;
      if(a.sheet){sh=cx.wb.sheets.find(s=>s.name.toLowerCase()===a.sheet.toLowerCase());if(!sh)return E.REF;}
      return new Rng(sh,a.r1,a.c1,a.r2,a.c2);}
    case 'pct':{const v=scalar(evalAst(a.e,cx),cx);const n=num(v);return isErr(n)?n:n/100;}
    case 'un':{const n=num(scalar(evalAst(a.e,cx),cx));if(isErr(n))return n;return a.o==='-'?-n:n;}
    case 'bin':{
      const l=scalar(evalAst(a.l,cx),cx),r=scalar(evalAst(a.r,cx),cx);
      if(isErr(l))return l;if(isErr(r))return r;
      if(a.o==='&')return str(l)+str(r);
      if(['=','<>','<','>','<=','>='].includes(a.o)){
        const c=compare(l,r);
        return a.o==='='?c===0:a.o==='<>'?c!==0:a.o==='<'?c<0:a.o==='>'?c>0:a.o==='<='?c<=0:c>=0;}
      const x=num(l),y=num(r);if(isErr(x))return x;if(isErr(y))return y;
      switch(a.o){case '+':return x+y;case '-':return x-y;case '*':return x*y;
        case '/':return y===0?E.DIV:x/y;
        case '^':{const v=Math.pow(x,y);return isFinite(v)?v:E.NUM;}}
      return E.VAL;}
    case 'fn':return callFn(a,cx);
  }
  return E.VAL;
}
function scalar(v,cx){
  if(v instanceof Rng){
    if(v.rows*v.cols===1)return rawVal(cx.wb,v.sh,v.r1,v.c1);
    const sh=cx.cell;return E.VAL;}
  return v;
}
function compare(a,b){
  const rank=v=>v===null?0:typeof v==='number'?1:typeof v==='string'?2:3;
  if(a===null)a=typeof b==='string'?'':typeof b==='boolean'?false:0;
  if(b===null)b=typeof a==='string'?'':typeof a==='boolean'?false:0;
  const ra=rank(a),rb=rank(b);if(ra!==rb)return ra-rb;
  if(typeof a==='string')return a.toLowerCase()<b.toLowerCase()?-1:a.toLowerCase()>b.toLowerCase()?1:0;
  return a<b?-1:a>b?1:0;
}
/* kriteria COUNTIF dst. */
function critFn(c){
  if(typeof c==='number'||typeof c==='boolean')return v=>typeof v===typeof c&&v===c||(typeof c==='number'&&typeof v==='number'&&v===c);
  let s=String(c),op='=';
  const m=/^(<=|>=|<>|=|<|>)(.*)$/.exec(s);if(m){op=m[1];s=m[2];}
  const n=num(s),isN=!isErr(n)&&s.trim()!=='';
  const wild=/[*?]/.test(s);
  const re=wild?new RegExp('^'+s.replace(/[.+^${}()|[\]\\]/g,'\\$&').replace(/\*/g,'.*').replace(/\?/g,'.')+'$','i'):null;
  return v=>{
    if(isN){if(typeof v!=='number'){if(op==='<>')return true;return false;}
      return op==='='?v===n:op==='<>'?v!==n:op==='<'?v<n:op==='>'?v>n:op==='<='?v<=n:v>=n;}
    if(v===null)v='';if(typeof v!=='string'&&typeof v!=='boolean'&&typeof v!=='number')return false;
    const t=String(v);
    if(re){const ok=re.test(t);return op==='<>'?!ok:ok;}
    const c=t.toLowerCase()===s.toLowerCase();
    if(op==='=')return c;if(op==='<>')return !c;
    if(typeof v==='number')return false;
    const k=t.toLowerCase().localeCompare(s.toLowerCase());
    return op==='<'?k<0:op==='>'?k>0:op==='<='?k<=0:k>=0;};
}
const FN={};
function callFn(a,cx){
  const f=FN[a.n];
  if(!f)return E.NAME;
  const lazy=f.lazy;
  if(lazy)return f(a.a,cx);
  const args=a.a.map(x=>evalAst(x,cx));
  for(const x of args)if(isErr(x))return x;
  return f(args,cx);
}
/* ambil nilai-nilai numerik dari argumen (rentang melewati teks/boolean/kosong) */
function nums(args,cx,strict){
  const out=[];
  for(const a of args){
    if(a instanceof Rng){
      for(const [r,c] of cellsIn(a)){const v=rawVal(cx.wb,a.sh,r,c);if(isErr(v))return v;if(typeof v==='number')out.push(v);}
    }else if(a===null){}
    else{const n=num(a);if(isErr(n)){if(typeof a==='string'&&!strict)return E.VAL;return n;}out.push(n);}
  }
  return out;
}
function flatVals(a,cx){
  if(!(a instanceof Rng))return [a];
  const out=[];for(let r=a.r1;r<=a.r2;r++)for(let c=a.c1;c<=a.c2;c++)out.push(rawVal(cx.wb,a.sh,r,c));return out;
}
const S1=(a,cx)=>scalar(a[0],cx);
const def=(n,f,lazy)=>{if(lazy)f.lazy=true;FN[n]=f;FN_NAMES.add(n);};
const agg=(n,f,empty)=>def(n,(a,cx)=>{const v=nums(a,cx);if(isErr(v))return v;return v.length||empty!==undefined?f(v):E.DIV;});
agg('SUM',v=>v.reduce((x,y)=>x+y,0),0);
agg('AVERAGE',v=>v.reduce((x,y)=>x+y,0)/v.length);
agg('MIN',v=>v.length?Math.min(...v):0,0);
agg('MAX',v=>v.length?Math.max(...v):0,0);
agg('PRODUCT',v=>v.reduce((x,y)=>x*y,1),0);
agg('MEDIAN',v=>{const s=[...v].sort((x,y)=>x-y),m=s.length>>1;return s.length%2?s[m]:(s[m-1]+s[m])/2;});
const variance=(v,pop)=>{const m=v.reduce((x,y)=>x+y,0)/v.length;return v.reduce((x,y)=>x+(y-m)**2,0)/(v.length-(pop?0:1));};
def('VAR',(a,cx)=>{const v=nums(a,cx);return isErr(v)?v:v.length<2?E.DIV:variance(v);});
def('VAR.S',FN.VAR);def('VAR.P',(a,cx)=>{const v=nums(a,cx);return isErr(v)?v:v.length<1?E.DIV:variance(v,true);});
def('STDEV',(a,cx)=>{const v=nums(a,cx);return isErr(v)?v:v.length<2?E.DIV:Math.sqrt(variance(v));});
def('STDEV.S',FN.STDEV);def('STDEV.P',(a,cx)=>{const v=nums(a,cx);return isErr(v)?v:v.length<1?E.DIV:Math.sqrt(variance(v,true));});
def('COUNT',(a,cx)=>{let n=0;for(const x of a){if(x instanceof Rng){for(const [r,c] of cellsIn(x))if(typeof rawVal(cx.wb,x.sh,r,c)==='number')n++;}else if(typeof x==='number'||(typeof x==='string'&&!isErr(num(x))&&x.trim()!==''))n++;}return n;});
def('COUNTA',(a,cx)=>{let n=0;for(const x of a){if(x instanceof Rng){for(const [r,c] of cellsIn(x)){const v=rawVal(cx.wb,x.sh,r,c);if(v!==null&&v!=='')n++;}}else if(x!==null)n++;}return n;});
def('COUNTBLANK',(a,cx)=>{const x=a[0];if(!(x instanceof Rng))return x===null||x===''?1:0;let n=x.rows*x.cols;for(const [r,c] of cellsIn(x)){const v=rawVal(cx.wb,x.sh,r,c);if(v!==null&&v!=='')n--;}return n;});
def('LARGE',(a,cx)=>{const v=nums([a[0]],cx);if(isErr(v))return v;const k=num(scalar(a[1],cx));if(isErr(k))return k;const s=v.sort((x,y)=>y-x);return k>=1&&k<=s.length?s[Math.floor(k)-1]:E.NUM;});
def('SMALL',(a,cx)=>{const v=nums([a[0]],cx);if(isErr(v))return v;const k=num(scalar(a[1],cx));if(isErr(k))return k;const s=v.sort((x,y)=>x-y);return k>=1&&k<=s.length?s[Math.floor(k)-1]:E.NUM;});
def('RANK',(a,cx)=>{const x=num(S1(a,cx));const v=nums([a[1]],cx);if(isErr(x)||isErr(v))return isErr(x)?x:v;const asc=a[2]!==undefined&&num(scalar(a[2],cx))!==0;const i=v.filter(y=>asc?y<x:y>x).length;return v.includes(x)?i+1:E.NA;});
const m1=(n,f)=>def(n,(a,cx)=>{const x=num(S1(a,cx));if(isErr(x))return x;const r=f(x);return typeof r==='number'&&!isFinite(r)?E.NUM:r;});
m1('ABS',Math.abs);m1('SQRT',x=>x<0?E.NUM:Math.sqrt(x));m1('EXP',Math.exp);m1('LN',x=>x<=0?E.NUM:Math.log(x));m1('LOG10',x=>x<=0?E.NUM:Math.log10(x));
m1('SIN',Math.sin);m1('COS',Math.cos);m1('TAN',Math.tan);m1('ATAN',Math.atan);m1('ASIN',x=>Math.abs(x)>1?E.NUM:Math.asin(x));m1('ACOS',x=>Math.abs(x)>1?E.NUM:Math.acos(x));
m1('INT',Math.floor);m1('SIGN',Math.sign);m1('DEGREES',x=>x*180/Math.PI);m1('RADIANS',x=>x*Math.PI/180);
m1('FACT',x=>{if(x<0)return E.NUM;let r=1;for(let i=2;i<=Math.floor(x);i++)r*=i;return r;});
m1('EVEN',x=>{const s=Math.sign(x)||1,v=Math.ceil(Math.abs(x));return s*(v%2?v+1:v);});
m1('ODD',x=>{const s=Math.sign(x)||1,v=Math.ceil(Math.abs(x));return s*(v%2?v:v+1);});
def('PI',()=>Math.PI);
def('LOG',(a,cx)=>{const x=num(S1(a,cx)),b=a[1]===undefined?10:num(scalar(a[1],cx));if(isErr(x))return x;if(isErr(b))return b;return x<=0||b<=0||b===1?E.NUM:Math.log(x)/Math.log(b);});
def('POWER',(a,cx)=>{const x=num(S1(a,cx)),y=num(scalar(a[1],cx));if(isErr(x))return x;if(isErr(y))return y;const v=Math.pow(x,y);return isFinite(v)?v:E.NUM;});
def('MOD',(a,cx)=>{const x=num(S1(a,cx)),y=num(scalar(a[1],cx));if(isErr(x))return x;if(isErr(y))return y;if(y===0)return E.DIV;return x-y*Math.floor(x/y);});
def('QUOTIENT',(a,cx)=>{const x=num(S1(a,cx)),y=num(scalar(a[1],cx));if(isErr(x)||isErr(y))return isErr(x)?x:y;return y===0?E.DIV:Math.trunc(x/y);});
const rnd=(f)=>(a,cx)=>{const x=num(S1(a,cx)),d=a[1]===undefined?0:num(scalar(a[1],cx));if(isErr(x))return x;if(isErr(d))return d;const p=Math.pow(10,Math.trunc(d));return f(x*p+0)/p;};
const rr=x=>Math.sign(x)*Math.round(Math.abs(x)+1e-12);
def('ROUND',rnd(rr));def('ROUNDUP',rnd(x=>Math.sign(x)*Math.ceil(Math.abs(x)-1e-12)));def('ROUNDDOWN',rnd(x=>Math.trunc(x)));
def('TRUNC',rnd(x=>Math.trunc(x)));
def('CEILING',(a,cx)=>{const x=num(S1(a,cx)),s=a[1]===undefined?1:num(scalar(a[1],cx));if(isErr(x)||isErr(s))return isErr(x)?x:s;return s===0?0:Math.ceil(x/s)*s;});
def('FLOOR',(a,cx)=>{const x=num(S1(a,cx)),s=a[1]===undefined?1:num(scalar(a[1],cx));if(isErr(x)||isErr(s))return isErr(x)?x:s;return s===0?E.DIV:Math.floor(x/s)*s;});
def('MROUND',(a,cx)=>{const x=num(S1(a,cx)),s=num(scalar(a[1],cx));if(isErr(x)||isErr(s))return isErr(x)?x:s;return s===0?0:Math.round(x/s)*s;});
def('RAND',()=>Math.random());
def('RANDBETWEEN',(a,cx)=>{const lo=num(S1(a,cx)),hi=num(scalar(a[1],cx));if(isErr(lo)||isErr(hi))return E.VAL;return Math.floor(lo+Math.random()*(hi-lo+1));});
def('SUMPRODUCT',(a,cx)=>{const arrs=a.map(x=>flatVals(x,cx));const n=arrs[0].length;if(arrs.some(x=>x.length!==n))return E.VAL;let s=0;for(let i=0;i<n;i++){let p=1;for(const x of arrs){const v=x[i];p*=typeof v==='number'?v:0;}s+=p;}return s;});
/* logika */
def('IF',(a,cx)=>{
  const c=truthy(scalar(evalAst(a[0],cx),cx));if(isErr(c))return c;
  if(c)return a[1]?evalAst(a[1],cx):true;
  return a[2]?evalAst(a[2],cx):false;},true);
def('IFERROR',(a,cx)=>{const v=scalar(evalAst(a[0],cx),cx);return isErr(v)?evalAst(a[1],cx):v;},true);
def('IFNA',(a,cx)=>{const v=scalar(evalAst(a[0],cx),cx);return v===E.NA||(isErr(v)&&v.code==='#N/A')?evalAst(a[1],cx):v;},true);
def('AND',(a,cx)=>{let r=true,any=false;for(const x of a)for(const v of flatVals(x,cx)){if(v===null||(typeof v==='string'&&x instanceof Rng))continue;const t=truthy(v);if(isErr(t))return t;any=true;if(!t)r=false;}return any?r:E.VAL;});
def('OR',(a,cx)=>{let r=false,any=false;for(const x of a)for(const v of flatVals(x,cx)){if(v===null||(typeof v==='string'&&x instanceof Rng))continue;const t=truthy(v);if(isErr(t))return t;any=true;if(t)r=true;}return any?r:E.VAL;});
def('NOT',(a,cx)=>{const t=truthy(S1(a,cx));return isErr(t)?t:!t;});
def('TRUE',()=>true);def('FALSE',()=>false);
def('XOR',(a,cx)=>{let n=0;for(const x of a)for(const v of flatVals(x,cx)){const t=truthy(v);if(isErr(t))return t;if(t)n++;}return n%2===1;});
def('IFS',(a,cx)=>{for(let i=0;i+1<a.length;i+=2){const c=truthy(scalar(evalAst(a[i],cx),cx));if(isErr(c))return c;if(c)return evalAst(a[i+1],cx);}return E.NA;},true);
def('SWITCH',(a,cx)=>{const v=scalar(evalAst(a[0],cx),cx);for(let i=1;i+1<a.length;i+=2){if(compare(v,scalar(evalAst(a[i],cx),cx))===0)return evalAst(a[i+1],cx);}return a.length%2===0?evalAst(a[a.length-1],cx):E.NA;},true);
def('CHOOSE',(a,cx)=>{const i=num(scalar(evalAst(a[0],cx),cx));if(isErr(i))return i;const k=Math.floor(i);return k>=1&&k<a.length?evalAst(a[k],cx):E.VAL;},true);
/* info */
def('ISNUMBER',(a,cx)=>typeof S1(a,cx)==='number');
def('ISTEXT',(a,cx)=>typeof S1(a,cx)==='string');
def('ISBLANK',(a,cx)=>{const v=a[0] instanceof Rng?rawVal(cx.wb,a[0].sh,a[0].r1,a[0].c1):a[0];return v===null;});
def('ISLOGICAL',(a,cx)=>typeof S1(a,cx)==='boolean');
def('ISERROR',(a,cx)=>isErr(S1(a,cx)),false);
def('ISERR',(a,cx)=>{const v=S1(a,cx);return isErr(v)&&v.code!=='#N/A';});
def('ISNA',(a,cx)=>{const v=S1(a,cx);return isErr(v)&&v.code==='#N/A';});
def('NA',()=>E.NA);
/* ISERROR harus menerima galat -> daftarkan sebagai lazy */
for(const n of ['ISERROR','ISERR','ISNA','ISNUMBER','ISTEXT','ISLOGICAL','ISBLANK']){const g=FN[n];FN[n]=(a,cx)=>{const v=a.length&&a[0] instanceof Rng?rawVal(cx.wb,a[0].sh,a[0].r1,a[0].c1):evalAst(a[0],cx);
  const v2=v instanceof Rng?rawVal(cx.wb,v.sh,v.r1,v.c1):v;
  if(n==='ISERROR')return isErr(v2);if(n==='ISERR')return isErr(v2)&&v2.code!=='#N/A';if(n==='ISNA')return isErr(v2)&&v2.code==='#N/A';
  if(n==='ISNUMBER')return typeof v2==='number';if(n==='ISTEXT')return typeof v2==='string';if(n==='ISLOGICAL')return typeof v2==='boolean';return v2===null;};FN[n].lazy=true;}
def('ROW',(a,cx)=>a.length?(a[0].k==='ref'?a[0].r1+1:E.VAL):(cx.cell?cx.cell.r+1:1),true);
def('COLUMN',(a,cx)=>a.length?(a[0].k==='ref'?a[0].c1+1:E.VAL):(cx.cell?cx.cell.c+1:1),true);
def('ROWS',(a,cx)=>a[0] instanceof Rng?a[0].rows:1);def('COLUMNS',(a,cx)=>a[0] instanceof Rng?a[0].cols:1);
/* teks */
const T1=(n,f)=>def(n,(a,cx)=>{const s=str(S1(a,cx));return isErr(s)?s:f(s);});
T1('LEN',s=>s.length);T1('UPPER',s=>s.toUpperCase());T1('LOWER',s=>s.toLowerCase());
T1('PROPER',s=>s.toLowerCase().replace(/(^|[^a-z\u00C0-\uFFFF])([a-z\u00C0-\uFFFF])/g,(m,p,c)=>p+c.toUpperCase()));
T1('TRIM',s=>s.trim().replace(/\s+/g,' '));T1('CLEAN',s=>s.replace(/[\x00-\x1f]/g,''));
def('CODE',(a,cx)=>{const s=str(S1(a,cx));return isErr(s)?s:s?s.charCodeAt(0):E.VAL;});
def('CHAR',(a,cx)=>{const n=num(S1(a,cx));return isErr(n)?n:n<1||n>65535?E.VAL:String.fromCharCode(n);});
def('LEFT',(a,cx)=>{const s=str(S1(a,cx)),n=a[1]===undefined?1:num(scalar(a[1],cx));if(isErr(s))return s;if(isErr(n)||n<0)return E.VAL;return s.slice(0,n);});
def('RIGHT',(a,cx)=>{const s=str(S1(a,cx)),n=a[1]===undefined?1:num(scalar(a[1],cx));if(isErr(s))return s;if(isErr(n)||n<0)return E.VAL;return n?s.slice(-n):'';});
def('MID',(a,cx)=>{const s=str(S1(a,cx)),st=num(scalar(a[1],cx)),n=num(scalar(a[2],cx));if(isErr(s))return s;if(isErr(st)||isErr(n)||st<1||n<0)return E.VAL;return s.substr(st-1,n);});
def('REPT',(a,cx)=>{const s=str(S1(a,cx)),n=num(scalar(a[1],cx));if(isErr(s))return s;if(isErr(n)||n<0)return E.VAL;return s.repeat(n);});
const joinFn=(a,cx)=>{let o='';for(const x of a)for(const v of flatVals(x,cx)){const s=str(v);if(isErr(s))return s;o+=s;}return o;};
def('CONCATENATE',joinFn);def('CONCAT',joinFn);
def('TEXTJOIN',(a,cx)=>{const d=str(scalar(a[0],cx)),ig=truthy(scalar(a[1],cx));const o=[];for(const x of a.slice(2))for(const v of flatVals(x,cx)){const s=str(v);if(isErr(s))return s;if(ig&&s==='')continue;o.push(s);}return o.join(d);});
def('EXACT',(a,cx)=>str(S1(a,cx))===str(scalar(a[1],cx)));
def('SUBSTITUTE',(a,cx)=>{const s=str(S1(a,cx)),o=str(scalar(a[1],cx)),n=str(scalar(a[2],cx));if(isErr(s))return s;if(!o)return s;
  if(a[3]!==undefined){const k=num(scalar(a[3],cx));let i=-1,c=0;while(c<k){i=s.indexOf(o,i+1);if(i<0)return s;c++;}return s.slice(0,i)+n+s.slice(i+o.length);}
  return s.split(o).join(n);});
def('REPLACE',(a,cx)=>{const s=str(S1(a,cx)),st=num(scalar(a[1],cx)),n=num(scalar(a[2],cx)),t=str(scalar(a[3],cx));return s.slice(0,st-1)+t+s.slice(st-1+n);});
const findF=ci=>(a,cx)=>{let f=str(S1(a,cx)),w=str(scalar(a[1],cx));const st=a[2]===undefined?1:num(scalar(a[2],cx));if(isErr(f)||isErr(w))return E.VAL;
  if(ci){f=f.toLowerCase();w=w.toLowerCase();}
  const i=w.indexOf(f,st-1);return i<0?E.VAL:i+1;};
def('FIND',findF(false));def('SEARCH',findF(true));
def('VALUE',(a,cx)=>{const v=S1(a,cx);const n=num(v);if(!isErr(n))return n;const p=parseInput(String(v));return typeof p.v==='number'?p.v:E.VAL;});
def('N',(a,cx)=>{const v=S1(a,cx);return typeof v==='number'?v:typeof v==='boolean'?+v:0;});
def('T',(a,cx)=>{const v=S1(a,cx);return typeof v==='string'?v:'';});
def('TEXT',(a,cx)=>{const v=S1(a,cx),f=str(scalar(a[1],cx));if(isErr(f))return f;
  const n=typeof v==='string'?num(v):v;
  if(typeof v==='string'&&isErr(n))return v;
  const alias={'Rp #,##0':'"Rp" #,##0'};return fmtValue(typeof n==='number'?n:v,alias[f]||f).t;});
/* tanggal */
const today=()=>{const d=new Date();return Math.floor(dateToSerial(new Date(d.getFullYear(),d.getMonth(),d.getDate())));};
def('TODAY',()=>today());def('NOW',()=>dateToSerial(new Date()));
def('DATE',(a,cx)=>{const y=num(S1(a,cx)),m=num(scalar(a[1],cx)),d=num(scalar(a[2],cx));if(isErr(y)||isErr(m)||isErr(d))return E.VAL;return dateToSerial(new Date(y<100?1900+y:y,m-1,d));});
def('TIME',(a,cx)=>{const h=num(S1(a,cx)),m=num(scalar(a[1],cx)),s=num(scalar(a[2],cx));return (h*3600+m*60+s)/86400%1;});
const dp=(n,f)=>def(n,(a,cx)=>{const v=num(S1(a,cx));if(isErr(v))return v;return f(serialToDate(v));});
dp('YEAR',d=>d.getFullYear());dp('MONTH',d=>d.getMonth()+1);dp('DAY',d=>d.getDate());dp('HOUR',d=>d.getHours());dp('MINUTE',d=>d.getMinutes());dp('SECOND',d=>d.getSeconds());
def('WEEKDAY',(a,cx)=>{const v=num(S1(a,cx)),t=a[1]===undefined?1:num(scalar(a[1],cx));if(isErr(v))return v;const w=serialToDate(v).getDay();return t===2?(w+6)%7+1:t===3?(w+6)%7:w+1;});
def('DAYS',(a,cx)=>{const x=num(S1(a,cx)),y=num(scalar(a[1],cx));return isErr(x)?x:isErr(y)?y:Math.trunc(x)-Math.trunc(y);});
def('EDATE',(a,cx)=>{const v=num(S1(a,cx)),m=num(scalar(a[1],cx));if(isErr(v)||isErr(m))return E.VAL;const d=serialToDate(v);return dateToSerial(new Date(d.getFullYear(),d.getMonth()+m,d.getDate()));});
def('EOMONTH',(a,cx)=>{const v=num(S1(a,cx)),m=num(scalar(a[1],cx));if(isErr(v)||isErr(m))return E.VAL;const d=serialToDate(v);return dateToSerial(new Date(d.getFullYear(),d.getMonth()+m+1,0));});
def('DATEDIF',(a,cx)=>{const s=num(S1(a,cx)),e=num(scalar(a[1],cx)),u=str(scalar(a[2],cx)).toUpperCase();if(isErr(s)||isErr(e)||s>e)return E.NUM;
  const d1=serialToDate(s),d2=serialToDate(e);let m=(d2.getFullYear()-d1.getFullYear())*12+d2.getMonth()-d1.getMonth();if(d2.getDate()<d1.getDate())m--;
  return u==='D'?Math.trunc(e)-Math.trunc(s):u==='M'?m:u==='Y'?Math.floor(m/12):E.NUM;});
/* bersyarat */
function critRange(a,cx,pairs,sumRng){
  const rgs=pairs.map(([r])=>r);
  if(!rgs.every(r=>r instanceof Rng))return E.VAL;
  const R=rgs[0].rows,C=rgs[0].cols;
  if(rgs.some(r=>r.rows!==R||r.cols!==C)||(sumRng&&(sumRng.rows!==R||sumRng.cols!==C)))return E.VAL;
  const fs=pairs.map(([,c])=>critFn(c));
  const out=[];
  for(let i=0;i<R;i++)for(let j=0;j<C;j++){
    let ok=true;for(let k=0;k<rgs.length&&ok;k++)ok=fs[k](rawVal(cx.wb,rgs[k].sh,rgs[k].r1+i,rgs[k].c1+j));
    if(ok)out.push(sumRng?rawVal(cx.wb,sumRng.sh,sumRng.r1+i,sumRng.c1+j):1);
  }
  return out;
}
const cv=(x,cx)=>{const v=scalar(x,cx);return v===null?'':v;};
def('COUNTIF',(a,cx)=>{const o=critRange(a,cx,[[a[0],cv(a[1],cx)]]);return isErr(o)?o:o.length;});
def('SUMIF',(a,cx)=>{const sr=a[2]||a[0];const o=critRange(a,cx,[[a[0],cv(a[1],cx)]],sr instanceof Rng?new Rng(sr.sh,sr.r1,sr.c1,sr.r1+a[0].rows-1,sr.c1+a[0].cols-1):sr);return isErr(o)?o:o.reduce((s,v)=>s+(typeof v==='number'?v:0),0);});
def('AVERAGEIF',(a,cx)=>{const sr=a[2]||a[0];const o=critRange(a,cx,[[a[0],cv(a[1],cx)]],new Rng(sr.sh,sr.r1,sr.c1,sr.r1+a[0].rows-1,sr.c1+a[0].cols-1));if(isErr(o))return o;const n=o.filter(v=>typeof v==='number');return n.length?n.reduce((s,v)=>s+v,0)/n.length:E.DIV;});
def('COUNTIFS',(a,cx)=>{const p=[];for(let i=0;i+1<a.length;i+=2)p.push([a[i],cv(a[i+1],cx)]);const o=critRange(a,cx,p);return isErr(o)?o:o.length;});
def('SUMIFS',(a,cx)=>{const p=[];for(let i=1;i+1<a.length;i+=2)p.push([a[i],cv(a[i+1],cx)]);const o=critRange(a,cx,p,a[0]);return isErr(o)?o:o.reduce((s,v)=>s+(typeof v==='number'?v:0),0);});
def('AVERAGEIFS',(a,cx)=>{const p=[];for(let i=1;i+1<a.length;i+=2)p.push([a[i],cv(a[i+1],cx)]);const o=critRange(a,cx,p,a[0]);if(isErr(o))return o;const n=o.filter(v=>typeof v==='number');return n.length?n.reduce((s,v)=>s+v,0)/n.length:E.DIV;});
def('MAXIFS',(a,cx)=>{const p=[];for(let i=1;i+1<a.length;i+=2)p.push([a[i],cv(a[i+1],cx)]);const o=critRange(a,cx,p,a[0]);if(isErr(o))return o;const n=o.filter(v=>typeof v==='number');return n.length?Math.max(...n):0;});
def('MINIFS',(a,cx)=>{const p=[];for(let i=1;i+1<a.length;i+=2)p.push([a[i],cv(a[i+1],cx)]);const o=critRange(a,cx,p,a[0]);if(isErr(o))return o;const n=o.filter(v=>typeof v==='number');return n.length?Math.min(...n):0;});
/* pencarian */
const at=(a,cx,i,j)=>rawVal(cx.wb,a.sh,a.r1+i,a.c1+j);
def('VLOOKUP',(a,cx)=>{const k=S1(a,cx),t=a[1],ci=num(scalar(a[2],cx)),ap=a[3]===undefined?true:truthy(scalar(a[3],cx));
  if(!(t instanceof Rng))return E.VAL;if(isErr(ci)||ci<1||ci>t.cols)return E.REF;
  let best=-1;
  for(let i=0;i<t.rows;i++){const v=at(t,cx,i,0);
    if(!ap){if(compare(v,k)===0&&v!==null){best=i;break;}}
    else{if(v===null)continue;if(typeof v!==typeof k)continue;if(compare(v,k)<=0)best=i;else break;}}
  if(best<0)return E.NA;return at(t,cx,best,Math.floor(ci)-1);});
def('HLOOKUP',(a,cx)=>{const k=S1(a,cx),t=a[1],ri=num(scalar(a[2],cx)),ap=a[3]===undefined?true:truthy(scalar(a[3],cx));
  if(!(t instanceof Rng))return E.VAL;if(isErr(ri)||ri<1||ri>t.rows)return E.REF;
  let best=-1;
  for(let j=0;j<t.cols;j++){const v=at(t,cx,0,j);
    if(!ap){if(compare(v,k)===0&&v!==null){best=j;break;}}
    else{if(v===null||typeof v!==typeof k)continue;if(compare(v,k)<=0)best=j;else break;}}
  if(best<0)return E.NA;return at(t,cx,Math.floor(ri)-1,best);});
def('INDEX',(a,cx)=>{const t=a[0];if(!(t instanceof Rng))return t;let r=a[1]===undefined?1:num(scalar(a[1],cx)),c=a[2]===undefined?1:num(scalar(a[2],cx));
  if(isErr(r)||isErr(c))return E.VAL;if(t.rows===1&&a[2]===undefined){c=r;r=1;}
  if(r<0||c<0||r>t.rows||c>t.cols)return E.REF;
  return at(t,cx,Math.max(r,1)-1,Math.max(c,1)-1);});
def('MATCH',(a,cx)=>{const k=S1(a,cx),t=a[1],mt=a[2]===undefined?1:num(scalar(a[2],cx));if(!(t instanceof Rng))return E.VAL;
  const vals=flatVals(t,cx);
  if(mt===0){const i=vals.findIndex(v=>v!==null&&compare(v,k)===0&&typeof v===typeof k);return i<0?E.NA:i+1;}
  let best=-1;
  for(let i=0;i<vals.length;i++){const v=vals[i];if(v===null||typeof v!==typeof k)continue;const c=compare(v,k);if(mt>0?c<=0:c>=0)best=i;else break;}
  return best<0?E.NA:best+1;});
def('XLOOKUP',(a,cx)=>{const k=S1(a,cx),l=a[1],r=a[2];if(!(l instanceof Rng)||!(r instanceof Rng))return E.VAL;const lv=flatVals(l,cx),i=lv.findIndex(v=>v!==null&&compare(v,k)===0);
  if(i<0)return a[3]!==undefined?scalar(a[3],cx):E.NA;return flatVals(r,cx)[i];});
def('LOOKUP',(a,cx)=>{const k=S1(a,cx),l=a[1],r=a[2]||l;const lv=flatVals(l,cx);let best=-1;for(let i=0;i<lv.length;i++){const v=lv[i];if(v===null||typeof v!==typeof k)continue;if(compare(v,k)<=0)best=i;else break;}return best<0?E.NA:flatVals(r,cx)[best];});
def('TRANSPOSE',(a,cx)=>E.VAL);

/* semua fungsi yang menerima nilai tunggal dengan kosong -> 0 */

/* ---------- evaluasi seluruh buku kerja ---------- */
function recalc(wb){
  PASS.n++;wb.circ=false;
  for(const sh of wb.sheets)for(const [k,cell] of sh.cells){
    if(cell.f!==undefined&&cell.f!==null)calcCell(wb,sh,cell,Math.floor(k/MAXC),k%MAXC);
  }
}

/* ---------- geometri sumbu (lebar kolom / tinggi baris, dengan baris tersembunyi) ---------- */
class Axis{
  constructor(def,max){this.def=def;this.max=max;this.ov=new Map();this.hid=new Set();this.fil=new Set();this.dirty=true;}
  size(i){if(this.hid.has(i)||this.fil.has(i))return 0;const v=this.ov.get(i);return v===undefined?this.def:v;}
  build(){
    const ks=new Set([...this.ov.keys(),...this.hid,...this.fil]);
    this.idx=[...ks].sort((a,b)=>a-b);let acc=0;
    this.cum=this.idx.map(i=>{acc+=this.size(i)-this.def;return acc;});this.dirty=false;
  }
  pos(i){
    if(this.dirty)this.build();
    let lo=0,hi=this.idx.length;while(lo<hi){const m=(lo+hi)>>1;if(this.idx[m]<i)lo=m+1;else hi=m;}
    return i*this.def+(lo?this.cum[lo-1]:0);
  }
  at(px){  /* indeks yang memuat px */
    let lo=0,hi=this.max-1;
    while(lo<hi){const m=(lo+hi+1)>>1;if(this.pos(m)<=px)lo=m;else hi=m-1;}
    while(lo>0&&this.size(lo)===0)lo--;
    return lo;
  }
  total(){return this.pos(this.max);}
  clone(){const a=new Axis(this.def,this.max);a.ov=new Map(this.ov);a.hid=new Set(this.hid);a.fil=new Set(this.fil);return a;}
}

/* ---------- gaya sel bawaan ---------- */
const bd=(c,s='thin')=>({t:[s,c],r:[s,c],b:[s,c],l:[s,c]});
const CSTY=[
 ['Normal','Normal',{}],
 ['Buruk','Buruk',{fill:'#ffc7ce',color:'#9c0006'}],
 ['Bagus','Bagus',{fill:'#c6efce',color:'#006100'}],
 ['Netral','Netral',{fill:'#ffeb9c',color:'#9c5700'}],
 ['Perhitungan','Perhitungan',{fill:'#f2f2f2',color:'#fa7d00',b:1,bd:bd('#7f7f7f')}],
 ['Sel Pemeriksaan','Sel Pemeriksaan',{fill:'#a5a5a5',color:'#ffffff',b:1,bd:{t:['double','#3f3f3f'],r:['double','#3f3f3f'],b:['double','#3f3f3f'],l:['double','#3f3f3f']}}],
 ['Teks Penjelasan','Teks Penjelasan',{color:'#7f7f7f',i:1}],
 ['Input','Input',{fill:'#ffcc99',color:'#3f3f76',bd:bd('#7f7f7f')}],
 ['Sel Tertaut','Sel Tertaut',{color:'#fa7d00',bd:{b:['double','#ff8001']}}],
 ['Catatan','Catatan',{fill:'#ffffcc',bd:bd('#b2b2b2')}],
 ['Output','Output',{fill:'#f2f2f2',color:'#3f3f3f',b:1,bd:bd('#3f3f3f')}],
 ['Teks Peringatan','Teks Peringatan',{color:'#ff0000'}],
 ['Judul 1','Judul 1',{b:1,sz:15,color:'#44546a',bd:{b:['thick','#4472c4']}}],
 ['Judul 2','Judul 2',{b:1,sz:13,color:'#44546a',bd:{b:['thick','#acc8f1']}}],
 ['Judul 3','Judul 3',{b:1,sz:11,color:'#44546a',bd:{b:['medium','#8eaadb']}}],
 ['Judul 4','Judul 4',{b:1,sz:11,color:'#44546a'}],
 ['Judul','Judul',{sz:18,color:'#44546a',font:'Cambria'}],
 ['Total','Total',{b:1,bd:{t:['thin','#4472c4'],b:['double','#4472c4']}}],
 ['Aksen1','Aksen1',{fill:'#4472c4',color:'#ffffff'}],['Aksen2','Aksen2',{fill:'#ed7d31',color:'#ffffff'}],['Aksen3','Aksen3',{fill:'#a5a5a5',color:'#ffffff'}],
 ['Aksen4','Aksen4',{fill:'#ffc000',color:'#ffffff'}],['Aksen5','Aksen5',{fill:'#5b9bd5',color:'#ffffff'}],['Aksen6','Aksen6',{fill:'#70ad47',color:'#ffffff'}],
 ['Koma','Koma',{fmt:'#,##0.00'}],['Mata Uang','Mata Uang',{fmt:FMT.Currency}],['Persen','Persen',{fmt:'0%'}]
];
const CSTY_KEYS=['fill','color','b','i','sz','font','bd','fmt','u','s'];

/* ---------- gaya tabel ---------- */
const TBL=[['Terang','#4472c4'],['Terang','#ed7d31'],['Terang','#7f7f7f'],['Terang','#ffc000'],['Terang','#5b9bd5'],['Terang','#70ad47'],
 ['Sedang','#4472c4'],['Sedang','#ed7d31'],['Sedang','#7f7f7f'],['Sedang','#ffc000'],['Sedang','#5b9bd5'],['Sedang','#70ad47'],
 ['Gelap','#4472c4'],['Gelap','#ed7d31'],['Gelap','#7f7f7f'],['Gelap','#70ad47']];
const mix=(h,t)=>{const n=parseInt(h.slice(1),16),f=x=>Math.round(x+(255-x)*t).toString(16).padStart(2,'0');return '#'+f(n>>16)+f((n>>8)&255)+f(n&255);};
function tableStyle(i){
  const [kind,c]=TBL[i];
  if(kind==='Terang')return {head:{b:1,color:'#000000',bd:{b:['medium',c]}},band:[null,mix(c,.8)],lines:{t:['thin',c],b:['thin',c]},bd:{t:['thin',c],b:['thin',c]}};
  if(kind==='Sedang')return {head:{b:1,color:'#ffffff',fill:c},band:[mix(c,.8),mix(c,.9)],bd:{t:['thin',c],b:['thin',c],l:['thin','#ffffff'],r:['thin','#ffffff']}};
  return {head:{b:1,color:'#ffffff',fill:'#000000'},band:[mix(c,.45),mix(c,.7)],bd:null,dark:c};
}

/* ---------- palet warna ---------- */
const PAL=[
 ['#ffffff','#000000','#e7e6e6','#44546a','#4472c4','#ed7d31','#a5a5a5','#ffc000','#5b9bd5','#70ad47'],
 ['#f2f2f2','#7f7f7f','#d0cece','#d6dce4','#d9e2f3','#fbe5d5','#ededed','#fff2cc','#deebf6','#e2efd9'],
 ['#d8d8d8','#595959','#aeaaaa','#adb9ca','#b4c6e7','#f7cbac','#dbdbdb','#ffe599','#bdd7ee','#c5e0b3'],
 ['#bfbfbf','#3f3f3f','#757070','#8496b0','#8eaadb','#f4b083','#c9c9c9','#ffd966','#9cc3e5','#a8d08d'],
 ['#a5a5a5','#262626','#3a3838','#323f4f','#2f5496','#c55a11','#7b7b7b','#bf9000','#2e75b5','#538135'],
 ['#7f7f7f','#0c0c0c','#171616','#222a35','#1f3864','#833c0b','#525252','#7f6000','#1f4e78','#375623']
];
const PALSTD=['#c00000','#ff0000','#ffc000','#ffff00','#92d050','#00b050','#00b0f0','#0070c0','#002060','#7030a0'];

/* ---------- CSS ---------- */
const CSS=`
.xl{--g:#217346;--gd:#1a5c38;--ln:#d4d4d4;--gl:#d4d4d4;height:100%;display:flex;flex-direction:column;font:12px "Segoe UI",system-ui,sans-serif;color:#444;background:#fff;position:relative;user-select:none;-webkit-user-select:none}
.xl *{box-sizing:border-box}
.xl button,.xl select,.xl input,.xl textarea{font:inherit;color:inherit}
.xl [hidden]{display:none!important}
.xl em{font-style:normal}
.xl .qat{height:26px;display:flex;align-items:center;justify-content:center;position:relative;font-size:13px;flex:none}
.xl .qat .l{position:absolute;left:6px;display:flex;gap:2px;align-items:center}
.xl .qat .w{background:var(--g);color:#fff;font-weight:700;width:18px;height:16px;display:grid;place-items:center;font-size:11px;margin-right:6px}
.xl .qat button{background:none;border:0;cursor:pointer;width:24px;height:22px;color:#555}
.xl .qat button:hover{background:#e1e1e1}
.xl .qat button:disabled{opacity:.35;cursor:default;background:none}
.xl .tabs{display:flex;height:26px;border-bottom:1px solid var(--ln);background:#fff;overflow-x:auto;scrollbar-width:none;flex:none}
.xl .tabs::-webkit-scrollbar{display:none}
.xl .tab{background:none;border:1px solid transparent;border-bottom:0;padding:0 14px;cursor:pointer;height:26px;font-size:12px;flex:none;white-space:nowrap}
.xl .tab:hover{color:var(--g)}
.xl .tab.on{border-color:var(--ln);background:#fff;color:var(--g);height:27px}
.xl .tab.file{background:var(--g);color:#fff;padding:0 22px;margin-right:6px;border:0}
.xl .tab.file:hover{background:var(--gd);color:#fff}
.xl .ribbon{height:92px;background:#fff;border-bottom:1px solid var(--ln);display:flex;padding:3px 4px;overflow-x:auto;overflow-y:hidden;white-space:nowrap;flex:none;-webkit-overflow-scrolling:touch}
.xl .grp{display:flex;flex-direction:column;border-right:1px solid #e3e3e3;padding:0 6px;position:relative;flex:none}
.xl .gb{flex:1;display:flex;gap:4px;align-items:flex-start}
.xl .gn{text-align:center;font-size:11px;color:#666;height:16px;padding:0 12px}
.xl .gl{position:absolute;right:1px;bottom:1px;width:12px;height:12px;border:0;background:none;cursor:pointer;font-size:9px;line-height:1;color:#666;padding:0}
.xl .gl:hover{background:#e6e6e6}
.xl .rows{display:flex;flex-direction:column;gap:2px}
.xl .row{display:flex;gap:2px;align-items:center;height:26px}
.xl .rows.t3{gap:0}.xl .rows.t3 .row{height:22px}
.xl .b{min-width:24px;height:24px;border:1px solid transparent;background:none;display:inline-flex;align-items:center;justify-content:center;gap:4px;padding:0 4px;cursor:pointer;position:relative}
.xl .b:hover{background:#e6e6e6;border-color:#cfcfcf}
.xl .b.on{background:#d3e8d9;border-color:#8fc4a1}
.xl .b:disabled{opacity:.4;pointer-events:none}
.xl .b.big{flex-direction:column;height:64px;min-width:46px;max-width:92px;gap:2px;font-size:11px;white-space:normal;text-align:center;line-height:1.15;justify-content:flex-start;padding-top:3px}
.xl .b.big em{font-size:24px;line-height:1.2;height:30px;display:flex;align-items:center}
.xl .b.rw{height:21px;width:100%;justify-content:flex-start;font-size:11.5px}
.xl .b.rw em{width:18px;text-align:center}
.xl .b .car{font-size:8px;font-style:normal;margin-left:1px;color:#555}
.xl .b.sm{font-size:11px;min-width:22px;padding:0 3px}
.xl select{height:23px;border:1px solid #ababab;background:#fff}
.xl .clr{position:relative;display:inline-flex;align-items:center}
.xl .clr .b{padding:0 3px}
.xl .clr .sw{display:inline-block;line-height:1.1;padding:0 2px;border-bottom:3px solid var(--sw,#ff0000);font-weight:600}
.xl .gal{display:grid;grid-template-columns:repeat(3,70px);grid-template-rows:repeat(2,28px);gap:3px;padding:2px;border:1px solid var(--ln);height:64px;align-self:flex-start}
.xl .gal button{border:1px solid #d0d0d0;cursor:pointer;padding:0 4px;font-size:12px;white-space:nowrap;overflow:hidden;text-align:left;background:#fff;font-family:Calibri,Carlito,sans-serif}
.xl .gal button:hover{outline:2px solid #f5c35d;outline-offset:-1px}
.xl .gmore{display:flex;flex-direction:column;width:16px;height:64px;border:1px solid var(--ln);border-left:0;margin-left:-4px}
.xl .gmore button{flex:1;border:0;background:#fff;cursor:pointer;font-size:8px}.xl .gmore button:hover{background:#e6e6e6}
.xl .soon{padding:18px 20px;color:#777;font-size:12px}
.xl .fbar{display:flex;align-items:center;height:26px;border-bottom:1px solid var(--ln);background:#fff;padding:2px 4px;gap:2px;flex:none}
.xl .nb{width:100px;height:22px;border:1px solid #c6c6c6;padding:0 4px;background:#fff;flex:none}
.xl .nbw{position:relative;display:flex;flex:none}
.xl .fbar .fb{width:24px;height:22px;border:0;background:none;cursor:pointer;color:#666;font-size:13px}
.xl .fbar .fb:hover{background:#e6e6e6}
.xl .fbar .fb.ok{color:#217346}.xl .fbar .fb.no{color:#c00}
.xl .fbar .fb.dis{opacity:.35;pointer-events:none}
.xl .fx{flex:1;height:22px;border:1px solid #c6c6c6;padding:0 6px;background:#fff;min-width:40px;user-select:text;-webkit-user-select:text}
.xl .fx:focus,.xl .nb:focus{outline:1px solid var(--g);outline-offset:-1px}
.xl .gwrap{overflow:hidden;flex:1;display:grid;grid-template-columns:auto 1fr;grid-template-rows:auto 1fr;min-height:0;background:#fff;position:relative}
.xl .corner{width:40px;height:20px;background:#fff;border-right:1px solid #ababab;border-bottom:1px solid #ababab;position:relative;cursor:pointer}
.xl .corner::after{content:"";position:absolute;right:3px;bottom:3px;border:6px solid transparent;border-right-color:#c0c0c0;border-bottom-color:#c0c0c0;border-top:0;border-left:0}
.xl .colh,.xl .rowh{position:relative;overflow:hidden;background:#fff}
.xl .colh{height:20px;border-bottom:1px solid #ababab}
.xl .rowh{width:40px;border-right:1px solid #ababab}
.xl .hc{position:absolute;display:flex;align-items:center;justify-content:center;color:#444;border-right:1px solid #d4d4d4;border-bottom:1px solid #d4d4d4;background:#fff;overflow:hidden;font-size:12px}
.xl .hc.s{background:#e1e1e1;color:var(--g)}
.xl .hc.a{background:var(--g);color:#fff;border-color:var(--g)}
.xl .colh .hc.s{border-bottom:2px solid var(--g)}
.xl .rowh .hc.s{border-right:2px solid var(--g)}
.xl .hc.hd{background:#d8d8d8}
.xl .vp{position:relative;overflow:auto;outline:none;background:#fff;-webkit-overflow-scrolling:touch}
.xl .sizer{width:1px;height:1px;pointer-events:none}
.xl .stage{position:sticky;left:0;top:0;overflow:hidden;margin-bottom:0}
.xl .stage>*{position:absolute}
.xl .gv,.xl .gh{background:var(--gl);pointer-events:none}
.xl .gv{width:1px}.xl .gh{height:1px}
.xl .cb{overflow:hidden;pointer-events:none}
.xl .ct{display:flex;flex-direction:column;overflow:hidden;pointer-events:none;font:11pt Calibri,Carlito,"Segoe UI",sans-serif;color:#000;line-height:1.2;padding:0 2px}
.xl .ct span{display:block;width:100%;white-space:pre}
.xl .ct.w span{white-space:pre-wrap;word-break:break-word}
.xl .dbar{pointer-events:none}
.xl .ico{pointer-events:none;font-size:11px;display:flex;align-items:center;padding-left:3px}
.xl .tint{background:rgba(33,115,70,.12);pointer-events:none}
.xl .selb{border:2px solid var(--g);pointer-events:none}
.xl .fh{width:7px;height:7px;background:var(--g);border:1px solid #fff;cursor:crosshair;z-index:3}
.xl .cpy{border:2px dashed var(--g);pointer-events:none}
.xl .pnt{border:2px solid #4472c4;pointer-events:none;background:rgba(68,114,196,.08)}
.xl .fdd{width:16px;height:16px;border:1px solid #999;background:#f5f5f5;font-size:8px;display:grid;place-items:center;cursor:pointer;z-index:2;padding:0}
.xl .fdd.on{background:#d3e8d9;border-color:var(--g)}
.xl .ed{z-index:4;border:2px solid var(--g);background:#fff;outline:none;resize:none;overflow:hidden;font:11pt Calibri,Carlito,"Segoe UI",sans-serif;color:#000;padding:0 2px;line-height:1.2;user-select:text;-webkit-user-select:text;white-space:pre;margin:0}
.xl .stbar{background:var(--g);color:#fff;height:22px;display:flex;align-items:center;justify-content:space-between;padding:0 8px;font-size:11px;flex:none}
.xl .stbar .s{display:flex;gap:18px;align-items:center}
.xl .stbar button{background:none;border:0;color:#fff;cursor:pointer;width:22px;height:20px;font-size:13px}
.xl .stbar button.on{background:rgba(255,255,255,.28)}
.xl .stbar input{width:120px}
.xl .stbar .mode{text-transform:uppercase;min-width:42px}
.xl .shbar{display:flex;align-items:stretch;height:24px;background:#f3f3f3;border-top:1px solid #c6c6c6;flex:none}
.xl .shnav{display:flex;align-items:center;padding:0 4px}
.xl .shnav button{background:none;border:0;width:20px;color:#777;cursor:pointer}
.xl .shs{display:flex;overflow-x:auto;scrollbar-width:none;align-items:stretch;flex:0 1 auto;max-width:calc(100% - 140px)}
.xl .shs::-webkit-scrollbar{display:none}
.xl .sht{padding:0 14px;display:flex;align-items:center;cursor:pointer;white-space:nowrap;border-right:1px solid #d4d4d4;color:#555;position:relative;flex:none;min-width:60px;justify-content:center}
.xl .sht:hover{background:#e8e8e8}
.xl .sht.on{background:#fff;color:var(--g);font-weight:700;box-shadow:inset 0 -2px 0 var(--g);border-top:1px solid #fff;margin-top:-1px}
.xl .sht input{width:90px;height:18px;border:1px solid var(--g);padding:0 3px;user-select:text}
.xl .shadd{width:28px;background:none;border:0;cursor:pointer;font-size:16px;color:#555;border-radius:50%;margin:1px 4px;width:22px;height:22px;align-self:center;border:1px solid #aaa;line-height:1;padding:0}
.xl .shadd:hover{background:#e0e0e0}
.xl .menu{position:absolute;z-index:10;top:52px;left:0;bottom:22px;width:260px;background:var(--g);color:#fff;padding:8px 0}
.xl .menu button{display:block;width:100%;background:none;border:0;color:#fff;text-align:left;padding:12px 24px;cursor:pointer;font-size:14px}
.xl .menu button:hover{background:rgba(255,255,255,.18)}
.xl .menu hr{border:0;border-top:1px solid rgba(255,255,255,.3);margin:6px 0}
.xl .toast{position:absolute;right:16px;bottom:50px;background:#323130;color:#fff;padding:8px 14px;z-index:60;max-width:min(360px,calc(100% - 32px))}
.xl .pop{position:fixed;z-index:50;background:#fff;border:1px solid #bbb;box-shadow:0 2px 8px rgba(0,0,0,.28);padding:2px 0;min-width:160px;font-size:12px;max-height:80vh;overflow-y:auto}
.xl .pop .mi{display:flex;align-items:center;gap:8px;padding:5px 24px 5px 8px;cursor:pointer;white-space:nowrap;position:relative}
.xl .pop .mi:hover,.xl .pop .mi.hv{background:#d3e8d9}
.xl .pop .mi.dis{color:#aaa;pointer-events:none}
.xl .pop .mi i{font-style:normal;width:18px;text-align:center;flex:none}
.xl .pop .mi .ar{position:absolute;right:8px;font-size:9px}
.xl .pop .mi .ck{margin-left:auto;color:var(--g)}
.xl .pop hr{border:0;border-top:1px solid #e1e1e1;margin:2px 0}
.xl .pop .mh{padding:4px 8px;font-weight:600;color:var(--g);font-size:11px;background:#f3f3f3}
.xl .pal{padding:6px 8px}
.xl .pal .pr{display:flex;gap:2px;margin-bottom:2px}
.xl .pal .pr button{width:18px;height:18px;border:1px solid #c8c8c8;padding:0;cursor:pointer}
.xl .pal .pr button:hover{outline:2px solid #f5c35d;z-index:1}
.xl .pal .pl{font-size:11px;color:#555;margin:4px 0 3px}
.xl .pal .po{display:block;width:100%;text-align:left;border:0;background:none;padding:4px 2px;cursor:pointer}
.xl .pal .po:hover{background:#d3e8d9}
.xl .tgal{display:grid;grid-template-columns:repeat(4,64px);gap:6px;padding:8px}
.xl .tgal button{width:64px;height:50px;border:1px solid #bbb;background:#fff;cursor:pointer;padding:2px;display:flex;flex-direction:column;gap:1px}
.xl .tgal button:hover{outline:2px solid #f5c35d}
.xl .tgal button i{display:block;flex:1}
.xl .sgal{display:grid;grid-template-columns:repeat(3,100px);gap:4px;padding:6px 8px}
.xl .sgal button{border:1px solid #d0d0d0;background:#fff;padding:3px 6px;cursor:pointer;text-align:left;font:12px Calibri,Carlito,sans-serif;white-space:nowrap;overflow:hidden}
.xl .sgal button:hover{outline:2px solid #f5c35d;outline-offset:-1px}
.xl .ov{position:absolute;inset:0;background:rgba(0,0,0,.25);z-index:70;display:grid;place-items:center}
.xl .dg{background:#fff;border:1px solid #888;box-shadow:0 4px 18px rgba(0,0,0,.3);padding:14px 16px;min-width:min(340px,92vw);max-width:94%;max-height:92%;display:flex;flex-direction:column;user-select:text;-webkit-user-select:text}
.xl .dg h3{margin:0 0 10px;font-size:14px;font-weight:600;color:var(--g)}
.xl .db{display:flex;flex-direction:column;gap:8px;overflow:auto;min-height:0}
.xl .db label{display:flex;align-items:center;gap:6px}
.xl .db input[type=text],.xl .db input[type=number],.xl .db select{height:26px;border:1px solid #aaa;padding:0 4px;max-width:100%}
.xl .db .r2{display:grid;grid-template-columns:auto 1fr;gap:6px 10px;align-items:center}
.xl .dbt{display:flex;gap:6px;justify-content:flex-end;margin-top:12px}
.xl .dbt button,.xl .db button.bt{height:26px;border:1px solid #aaa;background:#f3f3f3;padding:0 14px;cursor:pointer}
.xl .dbt .ok{background:var(--g);color:#fff;border-color:var(--g)}
.xl .dtabs{display:flex;border-bottom:1px solid #ccc;margin-bottom:8px;flex-wrap:wrap}
.xl .dtabs button{border:1px solid transparent;border-bottom:0;background:none;padding:4px 12px;cursor:pointer}
.xl .dtabs button.on{border-color:#ccc;background:#fff;color:var(--g);margin-bottom:-1px;border-bottom:1px solid #fff}
.xl .fcl{border:1px solid #aaa;height:140px;overflow:auto;background:#fff}
.xl .fcl div{padding:2px 6px;cursor:pointer}.xl .fcl div.on{background:var(--g);color:#fff}
.xl .bdp{display:grid;grid-template-columns:repeat(3,auto);gap:4px;justify-content:start}
.xl .bdp button{width:34px;height:28px;border:1px solid #bbb;background:#fff;cursor:pointer;padding:0}
.xl .bdp button.on{background:#d3e8d9;border-color:var(--g)}
.xl .prev{border:1px solid #aaa;height:54px;display:grid;place-items:center;font:11pt Calibri,Carlito,sans-serif;background:#fff;color:#000;min-width:200px}
.xl .fl{max-height:180px;overflow:auto;border:1px solid #bbb;padding:4px;background:#fff}
.xl .fl label{display:flex;gap:6px;padding:1px 0}
.xl .find{position:absolute;right:24px;top:118px;z-index:20;background:#fff;border:1px solid #aaa;box-shadow:0 2px 8px rgba(0,0,0,.25);padding:10px;display:grid;grid-template-columns:auto 190px;gap:6px;align-items:center;user-select:text}
.xl .find input[type=text]{height:24px;border:1px solid #aaa;padding:0 4px}
.xl .find .fo{grid-column:1/3;display:flex;gap:10px;font-size:11px;flex-wrap:wrap}
.xl .find .fb{grid-column:1/3;display:flex;gap:4px;justify-content:flex-end;flex-wrap:wrap}
.xl .find button{height:24px;border:1px solid #aaa;background:#f3f3f3;cursor:pointer;padding:0 8px}
.xl.rc .ribbon{display:none}
@media (pointer:coarse){
 .xl .qat{height:36px}.xl .qat button{width:34px;height:32px;font-size:16px}
 .xl .tabs{height:36px}.xl .tab{height:36px;padding:0 16px;font-size:13px}.xl .tab.on{height:37px}
 .xl .ribbon{height:112px}.xl .row{height:34px}.xl .b{min-width:34px;height:34px}.xl .b.rw{height:30px}.xl .b.big{height:84px}
 .xl select{height:34px;font-size:16px}.xl .fbar{height:36px}.xl .nb,.xl .fx{height:30px;font-size:16px}
 .xl .stbar{height:30px}.xl .shbar{height:34px}.xl .pop .mi{padding:10px 24px 10px 10px}
 .xl .menu{top:72px;bottom:30px}.xl .menu button{padding:16px 24px}
}
@media (max-width:700px){
 .xl .qat .fn{max-width:calc(100% - 150px);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
 .xl .menu{width:min(260px,86vw)}
 .xl .stbar .s{gap:10px}.xl .stbar .sum{display:none}.xl .stbar input{width:70px}
 .xl .find{left:8px;right:8px;top:auto;bottom:70px;grid-template-columns:auto 1fr}
 .xl .nb{width:64px}
}
@media print{body *{visibility:hidden}#xl-print,#xl-print *{visibility:visible}#xl-print{position:absolute;left:0;top:0}}
`;

/* ---------- pita ---------- */
const TABS=[['home','BERANDA'],['insert','SISIPKAN'],['layout','TATA LETAK HALAMAN'],['formulas','RUMUS'],['data','DATA'],['review','TINJAU'],['view','TAMPILAN']];
const B=(ic,tip,a,m,lab,dd)=>`<button class="b ${m||''}" title="${tip}" data-a="${a}"><em>${ic}</em>${m?`<span>${lab||tip}${dd?(m==='big'?'<br>▾':' ▾'):''}</span>`:''}${dd&&!m?'<i class="car">▾</i>':''}</button>`;
const S=(n,o,w,tip)=>`<select data-s="${n}" title="${tip||''}" style="width:${w}px">${o.map(x=>Array.isArray(x)?`<option value="${x[0]}">${x[1]}</option>`:`<option>${x}</option>`).join('')}</select>`;
const G=(n,h,l)=>`<div class="grp"><div class="gb">${h}</div><div class="gn">${n}</div>${l?`<button class="gl" data-a="fmtcells:${l}" title="Format Sel">↘</button>`:''}</div>`;
const R2=(a,b)=>`<div class="rows"><div class="row">${a}</div><div class="row">${b}</div></div>`;
const SPL=(ic,tip,a,ma,lab)=>`<div class="row"><button class="b rw" title="${tip}" data-a="${a}" style="flex:1"><em>${ic}</em><span>${lab||tip}</span></button><button class="b" title="${tip}" data-a="${ma}" style="min-width:16px;padding:0 2px"><i class="car">▾</i></button></div>`;
const CLRB=(w,glyph,c,tip,a,ma)=>`<span class="clr"><button class="b" title="${tip}" data-a="${a}"><em class="sw" data-w="${w}" style="--sw:${c}"><span class="sw" style="--sw:${c}">${glyph}</span></em></button><button class="b" data-a="${ma}" title="${tip}" style="min-width:14px;padding:0 1px"><i class="car">▾</i></button></span>`;
const AL=t=>{const w=[14,9,14,9],x=i=>t==='L'||t==='J'?0:t==='C'?(14-w[i])/2:14-w[i];if(t==='J')w.fill(14);return `<svg width="16" height="14">${w.map((v,i)=>`<rect x="${1+x(i)}" y="${1+i*3.5}" width="${v}" height="1.6" fill="#444"/>`).join('')}</svg>`;};
const AV=p=>{const y=p==='t'?[2,5.5,9]:p==='m'?[4,7.5,11]:[6,9.5,13];return `<svg width="16" height="16"><rect x=".5" y=".5" width="15" height="15" fill="none" stroke="#bbb"/>${y.map((v,i)=>`<rect x="${i===1?4:2}" y="${v}" width="${i===1?8:12}" height="1.6" fill="#444"/>`).join('')}</svg>`;};
const FILLI='<svg width="16" height="14" viewBox="0 0 16 14"><path d="M3 7 8 2l5 5-4.5 4.5z" fill="#fff" stroke="#444"/><path d="M13.5 8.5c1 1.4 1.5 2.2 1.5 3a1.5 1.5 0 0 1-3 0c0-.8.5-1.6 1.5-3z" fill="#444"/></svg>';
const BRDI='<svg width="16" height="16"><rect x="1.5" y="1.5" width="13" height="13" fill="none" stroke="#444" stroke-dasharray="1.5 1.5"/><path d="M1.5 14.5h13" stroke="#444" stroke-width="2"/></svg>';
const INC='<small style="font-size:9px;letter-spacing:-.5px">←,0</small>',DEC='<small style="font-size:9px;letter-spacing:-.5px">,00→</small>';
const csGal=CSTY.slice(0,6).map((s,i)=>{const o=s[2];return `<button data-a="cstyle:${i}" title="${s[1]}" style="${o.fill?'background:'+o.fill+';':''}${o.color?'color:'+o.color+';':''}${o.b?'font-weight:700;':''}${o.bd?'border:2px solid '+(o.bd.t?o.bd.t[1]:'#7f7f7f')+';':''}">${s[1]}</button>`;}).join('');
const RIB={
home:
 G('Papan klip',`<div style="display:flex;flex-direction:column;align-items:center"><button class="b big" data-a="paste" title="Tempel (Ctrl+V)" style="height:46px;min-width:46px"><em>📋</em></button><button class="b" data-a="m:paste" title="Opsi tempel" style="height:18px;font-size:11px;padding:0 4px">Tempel ▾</button></div>`+
   `<div class="rows t3">${B('✂','Potong (Ctrl+X)','cut','rw','Potong')}${B('⎘','Salin (Ctrl+C)','copy','rw','Salin')}${B('🖌','Kuas Format','painter','rw','Kuas Format')}</div>`)+
 G('Font',R2(
   S('font',FONTS,128,'Font')+S('size',SIZES,48,'Ukuran font')+B('A⁺','Perbesar font','grow')+B('A⁻','Perkecil font','shrink'),
   B('<b>B</b>','Tebal (Ctrl+B)','bold')+B('<i>I</i>','Miring (Ctrl+I)','italic')+B('<u>U</u>','Garis bawah (Ctrl+U)','underline')+
   `<button class="b" data-a="m:ul" title="Jenis garis bawah" style="min-width:12px;padding:0 1px;margin-left:-3px"><i class="car">▾</i></button>`+
   `<span style="width:1px;height:18px;background:#ddd;margin:0 2px"></span>`+
   `<span class="clr"><button class="b" title="Garis tepi" data-a="border:last"><em>${BRDI}</em></button><button class="b" data-a="m:border" title="Garis tepi" style="min-width:14px;padding:0 1px"><i class="car">▾</i></button></span>`+
   CLRB('fill',FILLI,'#ffff00','Warna isi','fillnow','m:fill')+CLRB('color','A','#ff0000','Warna font','colornow','m:color')),'font')+
 G('Perataan',R2(
   B(AV('t'),'Rata atas','va:top')+B(AV('m'),'Rata tengah','va:middle')+B(AV('b'),'Rata bawah','va:bottom')+B('<span style="display:inline-block;transform:rotate(-40deg);font-size:11px">ab</span>','Orientasi','m:orient',null,null,1)+
   B('↵','Bungkus teks','wrap','sm','Bungkus Teks'),
   B(AL('L'),'Rata kiri','ha:left')+B(AL('C'),'Tengah','ha:center')+B(AL('R'),'Rata kanan','ha:right')+B('⇤','Kurangi indentasi','indent:-1')+B('⇥','Tambah indentasi','indent:1')+
   B('⇔','Gabung & Tengah','merge:center','sm','Gabung & Tengah')+`<button class="b" data-a="m:merge" title="Opsi gabung" style="min-width:14px;padding:0 1px"><i class="car">▾</i></button>`),'align')+
 G('Angka',R2(
   S('nfmt',[...FMTLIST,['custom','Kustom']],130,'Format angka'),
   `<button class="b sm" data-a="m:acct" title="Format akuntansi"><em>Rp</em><i class="car">▾</i></button>`+B('%','Gaya persen','nf:Percent')+B('<b>,</b>','Gaya koma','nf:Comma')+B(INC,'Tambah desimal','dec:1')+B(DEC,'Kurangi desimal','dec:-1')),'number')+
 G('Gaya',B('◧','Pemformatan Bersyarat','m:cf','big','Pemformatan Bersyarat',1)+B('▦','Format sebagai Tabel','m:tbl','big','Format sebagai Tabel',1)+
   `<div class="gal">${csGal}</div><div class="gmore"><button data-a="m:csty" title="Gaya sel lainnya">▾</button></div>`+
   '<div style="width:0"></div>')+
 G('Sel',B('⊞','Sisipkan','m:insert','big','Sisipkan',1)+B('⊟','Hapus','m:delete','big','Hapus',1)+B('▤','Format','m:format','big','Format',1))+
 G('Pengeditan',`<div class="rows t3">${SPL('Σ','Jumlah Otomatis','autosum:SUM','m:sum','Jumlah Otomatis')}${SPL('⬇','Isi','filldown','m:fillm','Isi')}${SPL('⌧','Hapus','clearc','m:clear','Hapus')}</div>`+
   B('⇅','Urutkan & Filter','m:sort','big','Urutkan & Filter',1)+B('🔍','Cari & Pilih','m:find','big','Cari & Pilih',1))
};
const SOON='<div class="soon">Menu ini akan dibuat pada tahap berikutnya.</div>';

const BTYPES={thin:['1px','solid'],medium:['2px','solid'],thick:['3px','solid'],dashed:['1px','dashed'],dotted:['1px','dotted'],double:['3px','double'],hair:['1px','dotted']};
const SIDES=['t','r','b','l'];
const hexOk=c=>/^#[0-9a-f]{6}$/i.test(c||'');

function mount(root,opts={}){
  if(!document.getElementById('xl-css')){const s=document.createElement('style');s.id='xl-css';s.textContent=CSS;document.head.appendChild(s);}
  root.innerHTML=`<div class="xl">
    <div class="qat"><div class="l"><span class="w">X</span><button data-a="save" title="Simpan (Ctrl+S)">💾</button><button data-a="undo" title="Urungkan (Ctrl+Z)">↶</button><button data-a="redo" title="Ulangi (Ctrl+Y)">↷</button></div><span class="fn"></span></div>
    <div class="tabs"><button class="tab file" data-a="menu">BERKAS</button>${TABS.map((t,i)=>`<button class="tab${i?'':' on'}" data-tab="${t[0]}">${t[1]}</button>`).join('')}</div>
    ${TABS.map((t,i)=>`<div class="ribbon" data-pane="${t[0]}"${i?' hidden':''}>${RIB[t[0]]||SOON}</div>`).join('')}
    <div class="fbar"><div class="nbw"><input class="nb" value="A1" spellcheck="false" title="Kotak Nama"></div>
      <button class="fb no dis" data-a="editcancel" title="Batal">✕</button><button class="fb ok dis" data-a="editok" title="Masukkan">✓</button><button class="fb" data-a="fxdlg" title="Sisipkan fungsi"><i style="font-family:Georgia,serif">fx</i></button>
      <input class="fx" spellcheck="false" autocomplete="off"></div>
    <div class="gwrap"><div class="corner" title="Pilih semua"></div><div class="colh"></div><div class="rowh"></div>
      <div class="vp" tabindex="0"><div class="stage"></div><div class="sizer"></div></div><textarea class="ed" hidden spellcheck="false" style="position:absolute"></textarea></div>
    <div class="shbar"><div class="shnav"><button data-a="shprev">◀</button><button data-a="shnext">▶</button></div><div class="shs"></div><button class="shadd" data-a="addsheet" title="Lembar baru">+</button></div>
    <div class="stbar"><div class="s"><span class="mode">SIAP</span></div><div class="s"><span class="sum"></span></div>
      <div class="s"><span><button class="on" title="Normal">▦</button><button data-a="soonv" title="Tata Letak Halaman">▭</button><button data-a="soonv" title="Pratinjau Sela Halaman">▥</button></span>
      <span><button data-a="zoomd:-10">−</button><input type="range" min="20" max="200" step="10" value="100" class="zr"><button data-a="zoomd:10">+</button> <span class="zv" style="display:inline-block;min-width:34px;text-align:right">100%</span></span></div></div>
    <div class="find" hidden><span>Cari:</span><input type="text" class="f1"><span>Ganti:</span><input type="text" class="f2">
      <div class="fo"><label><input type="checkbox" class="fcase"> Cocokkan huruf besar/kecil</label><label><input type="checkbox" class="fwhole"> Seluruh isi sel</label></div>
      <div class="fb"><button data-a="fnext">Cari Berikutnya</button><button data-a="frep">Ganti</button><button data-a="frall">Ganti Semua</button><button data-a="fclose">Tutup</button></div></div>
    <div class="menu" hidden>
      <button data-a="new">📄  Buku kerja kosong baru</button><button data-a="open">📂  Buka…</button>
      <button data-a="save">💾  Simpan sebagai .xlsx</button><button data-a="savecsv">📑  Simpan lembar sebagai .csv</button><button data-a="print">🖨  Cetak / simpan PDF</button>
      <hr><button data-a="home">⬅  Kembali ke menu utama</button></div>
    <input type="file" class="fin" accept=".xlsx,.csv,.tsv,.txt" hidden>
  </div>`;
  const $=s=>root.querySelector(s),$$=s=>[...root.querySelectorAll(s)];
  const X=$('.xl'),vp=$('.vp'),stage=$('.stage'),sizer=$('.sizer'),colh=$('.colh'),rowh=$('.rowh'),menu=$('.menu'),fin=$('.fin');
  const nb=$('.nb'),fxi=$('.fx'),fd=$('.find');

  /* ===== model ===== */
  const wb={sheets:[],active:0,ver:0,circ:false};
  let cur=null,Z=1,name='Buku1.xlsx',showGrid=true,showHead=true;
  let ed=null,clip=null,pointing=null,painting=null,tabStart=null;
  let undoS=[],redoS=[],dirty=false;
  const newSheet=n=>({name:n,cells:new Map(),cw:new Axis(DW,MAXC),rh:new Axis(DH,MAXR),merges:[],cf:[],filter:null,tab:null,
    sel:{ar:0,ac:0,fr:0,fc:0},sx:0,sy:0});
  const uniqueName=(base,except)=>{let n=base,i=1;while(wb.sheets.some(s=>s!==except&&s.name.toLowerCase()===n.toLowerCase()))n=base.replace(/\d+$/,'')+(++i);return n;};
  const toast=m=>{const t=document.createElement('div');t.className='toast';t.textContent=m;X.appendChild(t);setTimeout(()=>t.remove(),3200);};
  const setName=n=>{name=n;$('.fn').textContent=n.replace(/\.(xlsx|csv)$/i,'')+' - Excel';document.title=$('.fn').textContent;};

  const cellAt=(sh,r,c)=>sh.cells.get(key(r,c));
  const cellVal=(sh,r,c)=>rawVal(wb,sh,r,c);
  function putCell(sh,r,c,cell){
    const k=key(r,c);
    if(!cell||((cell.v===null||cell.v===undefined)&&(cell.f===undefined||cell.f===null)&&(!cell.st||!Object.keys(cell.st).length)))sh.cells.delete(k);
    else sh.cells.set(k,cell);
  }
  const mkInput=(text,old)=>{ /* membuat sel dari masukan pengguna, mempertahankan gaya */
    const st=old&&old.st;
    if(text===''||text==null)return st?{v:null,st}:null;
    if(text[0]==='='&&text.length>1)return {f:text.slice(1),raw:text,st};
    if(text[0]==="'")return {v:text.slice(1),st};
    const p=parseInput(text);
    let ns=st;
    if(p.fmt&&(!st||!st.fmt||st.fmt==='General'))ns={...(st||{}),fmt:p.fmt};
    return {v:p.v,st:ns};
  };
  const rawText=cell=>!cell?'':cell.f!==undefined&&cell.f!==null?'='+cell.f:cell.v===null||cell.v===undefined?'':typeof cell.v==='boolean'?(cell.v?'TRUE':'FALSE'):typeof cell.v==='number'?(isDateCode(cell.st&&cell.st.fmt||'')?fmtValue(cell.v,cell.st.fmt).t:String(cell.v).replace('.',',')):String(cell.v);
  const recalcAll=()=>{wb.ver++;recalc(wb);};
  function shotSheet(s){return {name:s.name,cells:new Map(s.cells),cw:s.cw.clone(),rh:s.rh.clone(),merges:s.merges.map(m=>({...m})),cf:[...s.cf],filter:s.filter?{...s.filter,crit:{...s.filter.crit}}:null,tab:s.tab,sel:{...s.sel}};}
  const snapshot=()=>({sheets:wb.sheets.map(shotSheet),active:wb.active});
  function restore(sn){
    wb.sheets=sn.sheets.map(shotSheet);wb.active=Math.min(sn.active,wb.sheets.length-1);
    wb.sheets.forEach(s=>{s.sx=0;s.sy=0;});
    cur=wb.sheets[wb.active];recalcAll();drawTabs();sync();
  }
  function tx(fn,quiet){ /* transaksi: simpan keadaan sebelum perubahan untuk undo */
    const s=snapshot();
    try{fn();}catch(e){console.error(e);toast('Terjadi kesalahan: '+e.message);return false;}
    undoS.push(s);if(undoS.length>60)undoS.shift();redoS=[];dirty=true;
    recalcAll();if(!quiet)sync();updUndo();return true;
  }
  function undo(){if(!undoS.length)return;redoS.push(snapshot());restore(undoS.pop());updUndo();}
  function redo(){if(!redoS.length)return;undoS.push(snapshot());restore(redoS.pop());updUndo();}
  const updUndo=()=>{$('[data-a=undo]').disabled=!undoS.length;$('[data-a=redo]').disabled=!redoS.length;};

  /* ===== seleksi ===== */
  const sel=()=>cur.sel;
  const rg=()=>{const s=cur.sel;return {r1:Math.min(s.ar,s.fr),r2:Math.max(s.ar,s.fr),c1:Math.min(s.ac,s.fc),c2:Math.max(s.ac,s.fc)};};
  const mergeAtSh=(sh,r,c)=>sh.merges.find(m=>r>=m.r1&&r<=m.r2&&c>=m.c1&&c<=m.c2);
  const mergeAt=(r,c)=>mergeAtSh(cur,r,c);
  function expandMerge(g){ /* perluas rentang agar memuat sel gabungan yang terpotong */
    let ch=true;g={...g};
    while(ch){ch=false;for(const m of cur.merges){if(m.r2<g.r1||m.r1>g.r2||m.c2<g.c1||m.c1>g.c2)continue;
      if(m.r1<g.r1){g.r1=m.r1;ch=true;}if(m.r2>g.r2){g.r2=m.r2;ch=true;}if(m.c1<g.c1){g.c1=m.c1;ch=true;}if(m.c2>g.c2){g.c2=m.c2;ch=true;}}}
    return g;
  }
  function select(ar,ac,fr,fc,keepScroll){
    tabStart=null;const s=cur.sel;s.ar=ar;s.ac=ac;s.fr=fr===undefined?ar:fr;s.fc=fc===undefined?ac:fc;
    if(!keepScroll)reveal(fr===undefined?ar:fr,fc===undefined?ac:fc);
    sync();
  }
  function moveTo(r,c,ext){
    r=Math.max(0,Math.min(MAXR-1,r));c=Math.max(0,Math.min(MAXC-1,c));
    const s=cur.sel;
    if(ext)select(s.ar,s.ac,r,c);else select(r,c);
  }
  function reveal(r,c){
    const m=mergeAt(r,c);if(m){r=m.r1;c=m.c1;}
    const W=vp.clientWidth,H=vp.clientHeight;
    const x1=cx(c),x2=cx(c+1),y1=ry(r),y2=ry(r+1);
    if(x1<vp.scrollLeft)vp.scrollLeft=x1;else if(x2>vp.scrollLeft+W-17)vp.scrollLeft=Math.min(x1,x2-W+17);
    if(y1<vp.scrollTop)vp.scrollTop=y1;else if(y2>vp.scrollTop+H-17)vp.scrollTop=Math.min(y1,y2-H+17);
  }

  /* ===== geometri ===== */
  const cx=c=>Math.round(cur.cw.pos(c)*Z),ry=r=>Math.round(cur.rh.pos(r)*Z);
  const cwid=c=>cx(c+1)-cx(c),rhei=r=>ry(r+1)-ry(r);
  const rectOf=(r1,c1,r2,c2)=>({x:cx(c1),y:ry(r1),w:cx(c2+1)-cx(c1),h:ry(r2+1)-ry(r1)});
  const mctx=document.createElement('canvas').getContext('2d');
  const tw=(t,st)=>{mctx.font=`${st.i?'italic ':''}${st.b?'bold ':''}${(st.sz||11)*4/3*Z}px ${st.font||'Calibri'},Carlito,sans-serif`;return mctx.measureText(t).width;};
  const isEmptyCell=(sh,r,c)=>{const cl=sh.cells.get(key(r,c));if(!cl)return true;return (cl.v===null||cl.v===undefined)&&(cl.f===undefined||cl.f===null);};
  const hasVal=(sh,r,c)=>{const v=cellVal(sh,r,c);return v!==null&&v!==undefined&&v!=='';};
  function fitRows(rows){ /* tinggi baris otomatis untuk teks yang dibungkus */
    const m=document.createElement('div');m.style.cssText='position:absolute;visibility:hidden;left:-9999px;top:0;white-space:pre-wrap;word-break:break-word;line-height:1.2;padding:0 2px';X.appendChild(m);
    for(const r of rows){
      if(cur.rh.ov.has(r)&&cur.rh.manual&&cur.rh.manual.has(r))continue;
      let h=DH;
      for(const [k,cl] of cur.cells){if(Math.floor(k/MAXC)!==r)continue;const st=cl.st||{};
        const v=cellVal(cur,r,k%MAXC);const t=fmtValue(v,st.fmt).t;
        if(!t)continue;const mg=mergeAtSh(cur,r,k%MAXC);if(mg&&mg.r1!==mg.r2)continue;
        const sz=(st.sz||11)*4/3;m.style.font=`${st.i?'italic ':''}${st.b?'bold ':''}${sz}px ${st.font||'Calibri'},Carlito,sans-serif`;m.style.lineHeight='1.2';
        let w=cur.cw.size(k%MAXC);if(mg)for(let c=mg.c1+1;c<=mg.c2;c++)w+=cur.cw.size(c);
        m.style.width=(st.wrap?w:100000)+'px';m.textContent=t;h=Math.max(h,Math.ceil(m.scrollHeight)+(st.sz&&st.sz>11?3:2));}
      if(h!==DH)cur.rh.ov.set(r,h);else cur.rh.ov.delete(r);
    }
    m.remove();cur.rh.dirty=true;
  }

  /* ===== pemformatan bersyarat (evaluasi) ===== */
  function cfStat(rule){
    if(rule._ver===wb.ver&&rule._st)return rule._st;
    const sh=rule.sh||cur;const nums=[],all=new Map();let sum=0;
    for(const [r,c] of cellsIn(new Rng(sh,rule.r1,rule.c1,rule.r2,rule.c2))){const v=cellVal(sh,r,c);if(typeof v==='number'){nums.push(v);sum+=v;}if(v!==null&&v!=='')all.set(String(v).toLowerCase(),(all.get(String(v).toLowerCase())||0)+1);}
    const sorted=[...nums].sort((a,b)=>a-b);
    rule._st={min:sorted[0],max:sorted[sorted.length-1],avg:nums.length?sum/nums.length:0,sorted,n:nums.length,dup:all};rule._ver=wb.ver;return rule._st;
  }
  const lerp=(a,b,t)=>{const A=parseInt(a.slice(1),16),Bc=parseInt(b.slice(1),16),f=s=>Math.round(((A>>s)&255)*(1-t)+((Bc>>s)&255)*t).toString(16).padStart(2,'0');return '#'+f(16)+f(8)+f(0);};
  function cfFor(r,c,v){
    let out=null;
    for(const rule of cur.cf){
      if(r<rule.r1||r>rule.r2||c<rule.c1||c>rule.c2)continue;
      const o=out||(out={});
      const t=rule.type;
      if(t==='cell'){
        const a=typeof rule.a==='string'&&rule.a[0]==='='?cellVal(cur,...(()=>{const p=parseAddr(rule.a.slice(1));return p?[p.r,p.c]:[0,0];})()):rule.a;
        const b=rule.b;let ok=false;const nv=typeof v==='number';
        const cmpv=(x)=>typeof x==='number'?x:(typeof x==='string'&&!isErr(num(x))&&x.trim()!==''?num(x):x);
        const A=cmpv(a),Bv=cmpv(b);
        if(v===null||v==='')ok=false;
        else if(rule.op==='gt')ok=nv&&typeof A==='number'&&v>A;
        else if(rule.op==='lt')ok=nv&&typeof A==='number'&&v<A;
        else if(rule.op==='ge')ok=nv&&typeof A==='number'&&v>=A;
        else if(rule.op==='le')ok=nv&&typeof A==='number'&&v<=A;
        else if(rule.op==='eq')ok=typeof A==='number'?v===A:String(v).toLowerCase()===String(A).toLowerCase();
        else if(rule.op==='btw')ok=nv&&v>=Math.min(A,Bv)&&v<=Math.max(A,Bv);
        if(ok)Object.assign(o,rule.fmt);
      }else if(t==='text'){if(typeof v==='string'&&v.toLowerCase().includes(String(rule.a).toLowerCase()))Object.assign(o,rule.fmt);}
      else if(t==='dup'||t==='uniq'){if(v!==null&&v!==''){const n=cfStat(rule).dup.get(String(v).toLowerCase())||0;if((t==='dup'&&n>1)||(t==='uniq'&&n===1))Object.assign(o,rule.fmt);}}
      else if(typeof v==='number'){
        const st=cfStat(rule);
        if(t==='top'||t==='bottom'){
          let k=rule.pct?Math.max(1,Math.floor(st.n*rule.n/100)):rule.n;k=Math.min(k,st.n);
          const th=t==='top'?st.sorted[st.n-k]:st.sorted[k-1];
          if(t==='top'?v>=th:v<=th)Object.assign(o,rule.fmt);
        }else if(t==='above'){if(v>st.avg)Object.assign(o,rule.fmt);}
        else if(t==='below'){if(v<st.avg)Object.assign(o,rule.fmt);}
        else if(t==='bar'){const d=st.max-st.min;o.bar={p:d?Math.max(.05,(v-Math.min(0,st.min))/(Math.max(0,st.max)-Math.min(0,st.min)||1)):1,c:rule.color};}
        else if(t==='scale'){
          const d=st.max-st.min,p=d?(v-st.min)/d:.5,cs=rule.colors;
          o.fill=cs.length===2?lerp(cs[0],cs[1],p):(p<.5?lerp(cs[0],cs[1],p*2):lerp(cs[1],cs[2],(p-.5)*2));
        }else if(t==='icons'){
          const d=st.max-st.min,p=d?(v-st.min)/d:.5,i=p>=2/3?0:p>=1/3?1:2;o.icon=ICONS[rule.set][i];
        }
      }
    }
    return out;
  }
  const ICONS={arrows:[['▲','#00b050'],['►','#ffc000'],['▼','#ff0000']],lights:[['●','#00b050'],['●','#ffc000'],['●','#ff0000']],symbols:[['✔','#00b050'],['!','#ffc000'],['✖','#ff0000']]};

  /* ===== penggambaran ===== */
  let rafId=0;
  const schedule=()=>{if(!rafId)rafId=requestAnimationFrame(()=>{rafId=0;draw();});};
  const bdCss=(sd,pos)=>{if(!sd)return '';const [w,s]=BTYPES[sd[0]]||BTYPES.thin;return `border-${pos}:${w} ${s} ${sd[1]||'#000'};`;};
  function effBorders(sh,r,c,r2,c2,c1){
    /* tepi kanan/bawah memuat gabungan dengan tepi kiri/atas sel tetangga */
    const g=(rr,cc)=>{const cl=sh.cells.get(key(rr,cc));return cl&&cl.st&&cl.st.bd||null;};
    const o=g(r,c)||{};const eR=g(r,c2)||{},eB=g(r2,c)||{};
    const rn=g(r,c2+1),bn=g(r2+1,c);
    return {R:eR.r||(rn&&rn.l)||null,B:eB.b||(bn&&bn.t)||null,L:c===0?o.l:null,T:r===0?o.t:null};
  }
  function draw(){
    if(!cur)return;
    const W=vp.clientWidth,H=vp.clientHeight,sl=vp.scrollLeft,st0=vp.scrollTop;
    stage.style.width=W+'px';stage.style.height=H+'px';
    const hw=Math.round(HW*Z*(Z<1?1:1)),hh=HH;
    let c0=cur.cw.at(sl/Z),r0=cur.rh.at(st0/Z);
    let c1=c0,r1=r0;while(c1<MAXC-1&&cx(c1)-sl<W)c1++;while(r1<MAXR-1&&ry(r1)-st0<H)r1++;
    const cs=Math.max(0,c0-1),rs=Math.max(0,r0-1);
    let html='';
    /* garis kisi */
    if(showGrid){
      for(let c=c0;c<=c1;c++){const x=cx(c+1)-sl-1;if(cwid(c)>0)html+=`<div class="gv" style="left:${x}px;top:0;height:${H}px"></div>`;}
      for(let r=r0;r<=r1;r++){const y=ry(r+1)-st0-1;if(rhei(r)>0)html+=`<div class="gh" style="top:${y}px;left:0;width:${W}px"></div>`;}
    }
    /* sel */
    const merged=new Set();let bgH='',txH='',exH='';
    const mm=cur.merges;
    for(const m of mm)for(let r=m.r1;r<=m.r2;r++){if(r<rs||r>r1)continue;for(let c=m.c1;c<=m.c2;c++)if(c>=cs&&c<=c1)merged.add(key(r,c));}
    const cellOut=(r,c,rr2,cc2)=>{
      const cell=cur.cells.get(key(r,c));
      const stl=(cell&&cell.st)||{};
      const x=cx(c)-sl,y=ry(r)-st0,w=cx(cc2+1)-cx(c),h=ry(rr2+1)-ry(r);
      if(w<=0||h<=0)return;
      let v=cell?cellVal(cur,r,c):null;
      const cf=cur.cf.length?cfFor(r,c,v):null;
      const fill=(cf&&cf.fill)||stl.fill,color=(cf&&cf.color)||stl.color;
      const b=cf&&cf.b?1:stl.b;
      /* latar & tepi */
      const eb=effBorders(cur,r,c,rr2,cc2);
      let bs='';
      if(mergeAtSh(cur,r,c)&&(rr2>r||cc2>c)){const o=cell&&cell.st&&cell.st.bd||{};const cl=(rr,cc)=>{const q=cur.cells.get(key(rr,cc));return q&&q.st&&q.st.bd||{};};
        eb.R=cl(r,cc2).r||eb.R;eb.B=cl(rr2,c).b||eb.B;eb.L=o.l||eb.L;eb.T=o.t||eb.T;}
      bs=bdCss(eb.R,'right')+bdCss(eb.B,'bottom')+bdCss(eb.L,'left')+bdCss(eb.T,'top');
      if(fill||bs)bgH+=`<div class="cb" style="left:${x}px;top:${y}px;width:${w}px;height:${h}px;${fill?'background:'+fill+';':''}${bs}"></div>`;
      if(cf&&cf.bar){const bw=Math.max(2,Math.round((w-4)*cf.bar.p));bgH+=`<div class="dbar" style="left:${x+2}px;top:${y+2}px;width:${bw}px;height:${h-4}px;background:linear-gradient(90deg,${cf.bar.c},${mix(cf.bar.c,.65)})"></div>`;}
      if(v===null||v===undefined||v==='')return;
      /* teks */
      const f=fmtValue(v,stl.fmt);let t=f.t;
      if(t==='')return;
      const sz=(stl.sz||11)*4/3*Z;
      const fs={i:stl.i,b,sz:stl.sz,font:stl.font};
      let ha=stl.ha&&stl.ha!=='general'?stl.ha:(f.al||'left');
      const isNum=typeof v==='number';
      const pad=4,avail=w-pad-(stl.ind&&ha!=='center'?stl.ind*9*Z:0);
      if(isNum&&!stl.wrap){
        let wdt=tw(t,fs);
        if(wdt>avail){
          if(!stl.fmt||stl.fmt==='General'){
            for(let p=9;p>=1&&wdt>avail;p--){let u=Math.abs(v)>=1e10||(Math.abs(v)<1e-4&&v!==0)?v.toExponential(Math.max(0,p-1)).replace('e','E').replace('.',','):String(+v.toPrecision(p)).replace('.',',');if(u.includes('e'))u=v.toExponential(Math.max(0,p-1)).replace('e','E').replace('.',',');t=u;wdt=tw(t,fs);}
            if(wdt>avail)t='#'.repeat(Math.max(1,Math.floor(avail/(tw('#',fs)||6))));
          }else t='#'.repeat(Math.max(1,Math.floor(avail/(tw('#',fs)||6))));
        }
      }
      let tx0=x,tw0=w;
      if(!stl.wrap&&!isNum&&typeof v!=='boolean'&&!isErr(v)&&!stl.rot){
        const wdt=tw(t,fs)+pad;
        if(wdt>w){
          let ext=0;
          if(ha==='left'){let c=cc2+1,sp=0;while(sp<wdt-w&&c<MAXC&&c<=c1+2&&isEmptyCell(cur,r,c)&&!mergeAtSh(cur,r,c)){sp+=cwid(c);c++;}ext=sp;tw0=w+Math.min(sp,wdt-w+2);}
          else if(ha==='right'){let sp=0,k=c-1;while(sp<wdt-w&&k>=0&&k>=c0-2&&isEmptyCell(cur,r,k)&&!mergeAtSh(cur,r,k)){sp+=cwid(k);k--;}tx0=x-Math.min(sp,wdt-w+2);tw0=w+Math.min(sp,wdt-w+2);}
          else{let spL=0,spR=0,k=c-1,j=cc2+1;const need=(wdt-w)/2;while(spL<need&&k>=0&&k>=c0-2&&isEmptyCell(cur,r,k)){spL+=cwid(k);k--;}while(spR<need&&j<MAXC&&j<=c1+2&&isEmptyCell(cur,r,j)){spR+=cwid(j);j++;}
            const e=Math.min(spL,spR);tx0=x-e;tw0=w+2*e;}
        }
      }
      let jc=stl.va==='top'?'flex-start':stl.va==='middle'?'center':'flex-end';
      let spanCss='text-align:'+(ha==='fill'?'left':ha)+';';
      if(stl.rot){
        const R=stl.rot;
        if(R===255)spanCss+='writing-mode:vertical-rl;text-orientation:upright;';
        else if(R===90||R===-90){spanCss+='writing-mode:vertical-rl;'+(R===90?'transform:rotate(180deg);':'');jc=stl.va==='top'?'flex-start':stl.va==='bottom'?'flex-end':'center';}
        else spanCss+=`transform:rotate(${-R}deg);transform-origin:${R>0?'0 100%':'100% 100%'};white-space:nowrap;`;
      }
      const ind=stl.ind?(ha==='right'?`padding-right:${stl.ind*9*Z+2}px;`:`padding-left:${stl.ind*9*Z+2}px;`):'';
      let ico='';
      if(cf&&cf.icon)ico=`<div class="ico" style="left:${x}px;top:${y}px;height:${h}px;color:${cf.icon[1]};font-size:${Math.max(8,sz*.85)}px">${cf.icon[0]}</div>`;
      const dec=(stl.u?(stl.u===2?'underline double':'underline'):'')+(stl.s?' line-through':'');
      txH+=`<div class="ct${stl.wrap?' w':''}" style="left:${tx0}px;top:${y}px;width:${tw0}px;height:${h}px;justify-content:${jc};font-size:${sz}px;${color?'color:'+color+';':''}${stl.font?'font-family:'+stl.font+',Calibri,Carlito,sans-serif;':''}${b?'font-weight:700;':''}${stl.i?'font-style:italic;':''}${dec?'text-decoration:'+dec+';':''}${ind}${cf&&cf.icon?'padding-left:'+(sz+6)+'px;':''}"><span style="${spanCss}">${esc(t)}</span></div>`+ico;
    };
    for(let r=rs;r<=r1;r++){
      if(rhei(r)<=0)continue;
      for(let c=cs;c<=c1;c++){
        if(cwid(c)<=0)continue;
        const k=key(r,c);if(merged.has(k))continue;
        const cl=cur.cells.get(k);
        if(!cl){ /* sel kosong: hanya gambar bila tetangga membawa tepi */ 
          const rn=cur.cells.get(key(r,c+1)),bn=cur.cells.get(key(r+1,c));
          if(!((rn&&rn.st&&rn.st.bd&&rn.st.bd.l)||(bn&&bn.st&&bn.st.bd&&bn.st.bd.t)))continue;
        }
        cellOut(r,c,r,c);
      }
    }
    for(const m of mm){if(m.r2<rs||m.r1>r1||m.c2<cs||m.c1>c1)continue;cellOut(m.r1,m.c1,m.r2,m.c2);}
    html+=bgH+txH;
    /* filter */
    if(cur.filter){const f=cur.filter;
      for(let c=f.c1;c<=f.c2;c++){if(c<c0||c>c1||ry(f.r1)-st0>H||ry(f.r1+1)<st0)continue;
        const x=cx(c+1)-sl-17,y=ry(f.r1+1)-st0-17;
        html+=`<button class="fdd${f.crit[c]?' on':''}" data-fc="${c}" style="left:${x}px;top:${y}px">${f.crit[c]?'⧩':'▾'}</button>`;}}
    /* seleksi */
    const g0=rg(),g=expandMerge(g0),s=cur.sel;
    const ac=mergeAt(s.ar,s.ac),ar=ac?{r1:ac.r1,c1:ac.c1,r2:ac.r2,c2:ac.c2}:{r1:s.ar,c1:s.ac,r2:s.ar,c2:s.ac};
    const R=rectOf(g.r1,g.c1,g.r2,g.c2),A=rectOf(ar.r1,ar.c1,ar.r2,ar.c2);
    const rx=R.x-sl,ryy=R.y-st0,ax=A.x-sl,ay=A.y-st0;
    const multi=g.r1!==g.r2||g.c1!==g.c2;
    const multiNotMerge=multi&&!(ac&&ac.r1===g.r1&&ac.r2===g.r2&&ac.c1===g.c1&&ac.c2===g.c2);
    if(multiNotMerge){ /* arsiran di sekitar sel aktif */
      const T=(x,y,w,h)=>w>0&&h>0?`<div class="tint" style="left:${x}px;top:${y}px;width:${w}px;height:${h}px"></div>`:'';
      html+=T(rx,ryy,R.w,ay-ryy)+T(rx,ay+A.h,R.w,ryy+R.h-ay-A.h)+T(rx,ay,ax-rx,A.h)+T(ax+A.w,ay,rx+R.w-ax-A.w,A.h);
    }
    html+=`<div class="selb" style="left:${rx-1}px;top:${ryy-1}px;width:${R.w+1}px;height:${R.h+1}px"></div>`;
    if(clip&&clip.sh===cur){const q=rectOf(clip.r1,clip.c1,clip.r2,clip.c2);html+=`<div class="cpy" style="left:${q.x-sl-1}px;top:${q.y-st0-1}px;width:${q.w+1}px;height:${q.h+1}px"></div>`;}
    if(pointing){const q=rectOf(pointing.r1,pointing.c1,pointing.r2,pointing.c2);html+=`<div class="pnt" style="left:${q.x-sl-1}px;top:${q.y-st0-1}px;width:${q.w+1}px;height:${q.h+1}px"></div>`;}
    if(!ed)html+=`<div class="fh" style="left:${rx+R.w-5}px;top:${ryy+R.h-5}px" data-fill="1"></div>`;
    stage.innerHTML=html;
    if(ed)placeEditor();
    /* tajuk */
    let ch='',rh2='';
    const all=g.r1===0&&g.r2===MAXR-1,allc=g.c1===0&&g.c2===MAXC-1;
    for(let c=c0;c<=c1;c++){const w=cwid(c);const on=c>=g.c1&&c<=g.c2;if(w<=0){ch+=`<div class="hc hd" style="left:${cx(c)-sl-1}px;width:3px;top:0;height:${hh}px"></div>`;continue;}
      ch+=`<div class="hc${on?(all?' a':' s'):''}" data-h="c${c}" style="left:${cx(c)-sl}px;width:${w}px;top:0;height:${hh}px">${colName(c)}</div>`;}
    for(let r=r0;r<=r1;r++){const h=rhei(r);const on=r>=g.r1&&r<=g.r2;if(h<=0){rh2+=`<div class="hc hd" style="top:${ry(r)-st0-1}px;height:3px;left:0;width:${HW}px"></div>`;continue;}
      rh2+=`<div class="hc${on?(allc?' a':' s'):''}" data-h="r${r}" style="top:${ry(r)-st0}px;height:${h}px;left:0;width:${HW}px;font-size:${Math.max(8,12*Math.min(1,Z))}px">${r+1}</div>`;}
    colh.innerHTML=ch;rowh.innerHTML=rh2;
    sizer.style.width=Math.min(33e6,cx(MAXC))+'px';sizer.style.height=Math.min(33e6,ry(MAXR))+'px';
    stage.style.marginBottom=-H+'px';
  }

  /* ===== sinkronisasi UI ===== */
  const edEl=$('.ed');
  function updStatus(){
    const g=rg();let n=0,sum=0,cnt=0;
    if((g.r2-g.r1+1)*(g.c2-g.c1+1)>1){
      for(const [r,c] of cellsIn(new Rng(cur,g.r1,g.c1,g.r2,g.c2))){const v=cellVal(cur,r,c);if(v!==null&&v!==''){cnt++;if(typeof v==='number'){n++;sum+=v;}}}
    }
    const f=v=>fmtValue(+v.toPrecision(12),'General').t;
    $('.sum').textContent=cnt>1&&n>0?`Rata-rata: ${f(sum/n)}     Hitung: ${cnt}     Jumlah: ${f(sum)}`:cnt>1?`Hitung: ${cnt}`:'';
  }
  const stOf=(r,c)=>{const cl=cellAt(cur,r,c);return (cl&&cl.st)||{};};
  function setSel2(n,v){const s=$(`[data-s=${n}]`);if(!s)return;if(![...s.options].some(o=>o.value==String(v))){const o=document.createElement('option');o.value=String(v);o.textContent=String(v);s.appendChild(o);}s.value=String(v);}
  function refreshRibbon(){
    const s=cur.sel,st=stOf(s.ar,s.ac);
    const T=(a,on)=>{const b=root.querySelector(`[data-a="${a}"]`);if(b)b.classList.toggle('on',!!on);};
    T('bold',st.b);T('italic',st.i);T('underline',st.u);T('wrap',st.wrap);
    T('va:top',st.va==='top');T('va:middle',st.va==='middle');T('va:bottom',!st.va||st.va==='bottom');
    T('ha:left',st.ha==='left');T('ha:center',st.ha==='center'||st.ha==='centerAcross');T('ha:right',st.ha==='right');
    T('merge:center',!!mergeAt(s.ar,s.ac));
    setSel2('font',st.font||'Calibri');setSel2('size',st.sz||11);
    const k=fmtKey(st.fmt||'General');setSel2('nfmt',k||'custom');
  }
  function sync(){
    if(!cur)return;
    const s=cur.sel,g=rg();
    if(document.activeElement!==nb){
      const m=mergeAt(s.ar,s.ac);
      nb.value=(g.r1===g.r2&&g.c1===g.c2)||m&&m.r1===g.r1&&m.r2===g.r2&&m.c1===g.c1&&m.c2===g.c2?addr(m?m.r1:s.ar,m?m.c1:s.ac)
        :(g.c1===0&&g.c2===MAXC-1?(g.r1+1)+':'+(g.r2+1):g.r1===0&&g.r2===MAXR-1?colName(g.c1)+':'+colName(g.c2):`${addr(g.r1,g.c1)}:${addr(g.r2,g.c2)}`);
      if(g.c1===0&&g.c2===MAXC-1&&g.r1===g.r2)nb.value=`${g.r1+1}R`;
      else if(g.r1===0&&g.r2===MAXR-1&&g.c1===g.c2)nb.value=`${colName(g.c1)}K`;
      if(g.c1===0&&g.c2===MAXC-1&&g.r1===g.r2)nb.value=addr(s.ar,s.ac);
    }
    if(!ed){const m=mergeAt(s.ar,s.ac);fxi.value=rawText(cellAt(cur,m?m.r1:s.ar,m?m.c1:s.ac));}
    refreshRibbon();updStatus();schedule();
  }
  /* ===== editor sel ===== */
  const setMode=m=>{$('.mode').textContent=m;};
  const isFormulaText=t=>t.length>1&&t[0]==='=';
  function startEdit(r,c,init,fromBar){
    if(ed)return;
    const m=mergeAt(r,c);if(m){r=m.r1;c=m.c1;}
    const cell=cellAt(cur,r,c);
    reveal(r,c);
    ed={r,c,text:init!==undefined?init:rawText(cell),mode:init!==undefined?'enter':'edit',bar:!!fromBar};
    edEl.hidden=false;edEl.value=ed.text;fxi.value=ed.text;
    $('.fb.no').classList.remove('dis');$('.fb.ok').classList.remove('dis');
    setMode(ed.mode==='enter'?'Masuk':'Edit');
    placeEditor();
    if(fromBar){fxi.focus();}else{edEl.focus();const n=edEl.value.length;edEl.setSelectionRange(n,n);}
    schedule();
  }
  function placeEditor(){
    if(!ed)return;
    const st=stOf(ed.r,ed.c),sl=vp.scrollLeft,sy=vp.scrollTop;
    const m=mergeAt(ed.r,ed.c);
    const rr=m?m.r2:ed.r,cc=m?m.c2:ed.c;
    const x=HW+cx(ed.c)-sl,y=HH+ry(ed.r)-sy,w0=cx(cc+1)-cx(ed.c),h0=ry(rr+1)-ry(ed.r);
    const sz=(st.sz||11)*4/3*Z;
    edEl.style.fontSize=sz+'px';edEl.style.fontFamily=(st.font||'Calibri')+',Carlito,sans-serif';
    edEl.style.fontWeight=st.b?'700':'400';edEl.style.fontStyle=st.i?'italic':'normal';edEl.style.color=st.color||'#000';
    const lines=edEl.value.split('\n');
    let mw=0;for(const l of lines)mw=Math.max(mw,tw(l,{i:st.i,b:st.b,sz:st.sz,font:st.font}));
    const maxW=vp.clientWidth+HW-x;
    const w=Math.max(w0,Math.min(maxW,mw+16)),h=Math.max(h0,lines.length*sz*1.2+4);
    Object.assign(edEl.style,{left:x-1+'px',top:y-1+'px',width:w+2+'px',height:h+2+'px'});
    edEl.style.whiteSpace=w>=maxW?'pre-wrap':'pre';
  }
  function endEdit(){
    ed=null;pointing=null;edEl.hidden=true;edEl.value='';
    $('.fb.no').classList.add('dis');$('.fb.ok').classList.add('dis');setMode('Siap');
    if(document.activeElement!==nb)vp.focus({preventScroll:true});
  }
  function commitEdit(dr,dc){
    if(!ed)return true;
    const {r,c}=ed;let text=ed.text;
    if(isFormulaText(text)){let dp=0,qq=false;for(const ch of text){if(ch==='"')qq=!qq;else if(!qq){if(ch==='(')dp++;else if(ch===')')dp--;}}if(dp>0&&!qq)text+=')'.repeat(dp);}
    const startC=tabStart===null?c:tabStart;
    if(isFormulaText(text)){
      try{parseFormula(text.slice(1));}catch(e){toast('Ada masalah pada rumus ini. Periksa tanda kurung, argumen, dan alamat sel.');return false;}
    }
    const old=cellAt(cur,r,c);
    const same=rawText(old)===text;
    endEdit();
    if(!same){
      tx(()=>{putCell(cur,r,c,mkInput(text,old));const st=stOf(r,c);if(st.wrap)fitRows([r]);
        const nc=cellAt(cur,r,c);
        if(nc&&typeof nc.v==='number'&&nc.st&&nc.st.fmt&&isDateCode(nc.st.fmt)&&!cur.cw.ov.has(c)){const sz=Z;Z=1;const w=tw(fmtValue(nc.v,nc.st.fmt).t,{sz:nc.st.sz,font:nc.st.font})+10;Z=sz;if(w>cur.cw.def){cur.cw.ov.set(c,Math.ceil(w));cur.cw.dirty=true;}}},true);
    }
    if(dc){moveSel(dr,dc);tabStart=startC;}
    else if(dr&&tabStart!==null){moveTo(r+dr,startC);}
    else if(dr)moveSel(dr,dc);else sync();
    return true;
  }
  function cancelEdit(){if(!ed)return;endEdit();sync();}
  function moveSel(dr,dc){ /* Enter/Tab di dalam rentang terpilih */
    const g=rg(),s=cur.sel;
    if(g.r1!==g.r2||g.c1!==g.c2){
      let r=s.ar+dr,c=s.ac+dc;
      if(dr){if(r>g.r2){r=g.r1;c=s.ac+1;if(c>g.c2)c=g.c1;}else if(r<g.r1){r=g.r2;c=s.ac-1;if(c<g.c1)c=g.c2;}}
      if(dc){if(c>g.c2){c=g.c1;r=s.ar+1;if(r>g.r2)r=g.r1;}else if(c<g.c1){c=g.c2;r=s.ar-1;if(r<g.r1)r=g.r2;}}
      s.ar=r;s.ac=c;reveal(r,c);sync();
    }else{
      const m=mergeAt(s.ar,s.ac);let r=s.ar,c=s.ac;
      if(m){r=dr>0?m.r2:m.r1;c=dc>0?m.c2:m.c1;}
      moveTo(r+dr,c+dc,false);
    }
  }
  /* mode TUNJUK (point) pada rumus */
  const canPoint=()=>{
    if(!ed||!isFormulaText(ed.text))return false;
    const el=ed.bar?fxi:edEl,p=el.selectionStart===undefined?ed.text.length:el.selectionStart;
    if(pointing)return true;
    const pre=ed.text.slice(0,p).replace(/\s+$/,'');
    return pre.length>0&&'=+-*/^&(,;<>'.includes(pre[pre.length-1]);
  };
  function startPoint(r,c){
    const el=ed.bar?fxi:edEl,p=el.selectionStart===undefined?ed.text.length:el.selectionStart;
    pointing={r1:r,c1:c,r2:r,c2:c,ar:r,ac:c,start:pointing?pointing.start:p,len:pointing?pointing.len:0};
    setMode('Tunjuk');updPoint();
  }
  function updPoint(){
    const p=pointing;
    const a=addr(Math.min(p.r1,p.r2),Math.min(p.c1,p.c2)),b=addr(Math.max(p.r1,p.r2),Math.max(p.c1,p.c2));
    const ref=a===b?a:a+':'+b;
    const t=ed.text.slice(0,p.start)+ref+ed.text.slice(p.start+p.len);
    p.len=ref.length;ed.text=t;edEl.value=t;fxi.value=t;
    const el=ed.bar?fxi:edEl;const pos=p.start+ref.length;try{el.setSelectionRange(pos,pos);}catch(e){}
    placeEditor();schedule();
  }
  function onEdInput(src){
    if(!ed)return;
    ed.text=src.value;if(src===edEl)fxi.value=src.value;else edEl.value=src.value;
    pointing=null;if(ed.mode==='enter'&&false)ed.mode='edit';
    setMode(ed.mode==='enter'?'Masuk':'Edit');placeEditor();schedule();
  }
  edEl.addEventListener('input',()=>onEdInput(edEl));
  fxi.addEventListener('input',()=>onEdInput(fxi));
  fxi.addEventListener('focus',()=>{if(!ed){startEdit(cur.sel.ar,cur.sel.ac,undefined,true);}});
  function editKey(e,src){
    if(!ed)return;
    const k=e.key;
    if(k==='Enter'&&!e.altKey){e.preventDefault();commitEdit(e.shiftKey?-1:1,0);}
    else if(k==='Enter'&&e.altKey&&src===edEl){e.preventDefault();const p=edEl.selectionStart;edEl.setRangeText('\n',p,edEl.selectionEnd,'end');onEdInput(edEl);}
    else if(k==='Tab'){e.preventDefault();commitEdit(0,e.shiftKey?-1:1);}
    else if(k==='Escape'){e.preventDefault();cancelEdit();}
    else if(ed.mode==='enter'&&!isFormulaText(ed.text)&&['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'].includes(k)&&src===edEl){
      e.preventDefault();commitEdit(k==='ArrowDown'?1:k==='ArrowUp'?-1:0,k==='ArrowRight'?1:k==='ArrowLeft'?-1:0);}
    else if(k==='F2'){e.preventDefault();ed.mode=ed.mode==='enter'?'edit':'enter';setMode(ed.mode==='enter'?'Masuk':'Edit');}
  }
  edEl.addEventListener('keydown',e=>editKey(e,edEl));
  fxi.addEventListener('keydown',e=>editKey(e,fxi));

  /* ===== penyalinan ===== */
  const cloneCell=c=>c?{...c,ast:undefined,pass:undefined,busy:undefined,cv:undefined}:null;
  function captureClip(cut){
    const g=expandMerge(limitRange(rg()));
    const cells=[];
    for(let r=g.r1;r<=g.r2;r++){const row=[];for(let c=g.c1;c<=g.c2;c++)row.push(cloneCell(cellAt(cur,r,c)));cells.push(row);}
    const merges=cur.merges.filter(m=>m.r1>=g.r1&&m.r2<=g.r2&&m.c1>=g.c1&&m.c2<=g.c2).map(m=>({...m}));
    const text=cells.map((row,i)=>row.map((cl,j)=>{const v=cl?cellVal(cur,g.r1+i,g.c1+j):null;const t=fmtValue(v,cl&&cl.st&&cl.st.fmt).t;return /[\t\n"]/.test(t)?'"'+t.replace(/"/g,'""')+'"':t;}).join('\t')).join('\n');
    clip={sh:cur,...g,cut:!!cut,cells,merges,text};
    return clip;
  }
  function limitRange(g){
    const used=usedRange(cur);
    if(g.r2-g.r1>5000)g.r2=Math.min(g.r2,Math.max(used.r2,g.r1+99));
    if(g.c2-g.c1>200)g.c2=Math.min(g.c2,Math.max(used.c2,g.c1+25));
    return g;
  }
  function usedRange(sh){
    let r2=0,c2=0,r1=MAXR,c1=MAXC,any=false;
    for(const k of sh.cells.keys()){const r=Math.floor(k/MAXC),c=k%MAXC;any=true;if(r>r2)r2=r;if(c>c2)c2=c;if(r<r1)r1=r;if(c<c1)c1=c;}
    return any?{r1,c1,r2,c2,any}:{r1:0,c1:0,r2:0,c2:0,any:false};
  }
  async function copyAct(cut){
    const c=captureClip(cut);
    try{await navigator.clipboard.writeText(c.text);}catch(e){}
    schedule();
  }
  function pasteMode(mode,tr,tc,clp,txt){
    clp=clp||clip;
    const s=cur.sel;
    let g=rg();tr=tr===undefined?g.r1:tr;tc=tc===undefined?g.c1:tc;
    if(!clp){ // teks luar
      const rows=parseTSV(txt||'');if(!rows.length)return;
      tx(()=>{rows.forEach((row,i)=>row.forEach((t,j)=>{if(tr+i>=MAXR||tc+j>=MAXC)return;const old=cellAt(cur,tr+i,tc+j);putCell(cur,tr+i,tc+j,mkInput(t,old));}));});
      select(tr,tc,tr+rows.length-1,tc+rows[0].length-1);return;
    }
    const h=clp.r2-clp.r1+1,w=clp.c2-clp.c1+1;
    const T=mode==='transpose';
    const H=T?w:h,Wd=T?h:w;
    let rr=1,cr=1;
    if(h===1&&w===1&&(g.r2>g.r1||g.c2>g.c1)&&mode!=='transpose'){rr=g.r2-g.r1+1;cr=g.c2-g.c1+1;}
    else if(!T&&(g.r2-g.r1+1)%h===0&&(g.c2-g.c1+1)%w===0&&((g.r2-g.r1+1)>h||(g.c2-g.c1+1)>w)){rr=(g.r2-g.r1+1)/h;cr=(g.c2-g.c1+1)/w;}
    const srcSh=clp.sh;
    tx(()=>{
      for(let a=0;a<rr;a++)for(let b=0;b<cr;b++){
        const br=tr+a*H,bc=tc+b*Wd;
        for(let i=0;i<h;i++)for(let j=0;j<w;j++){
          const dr=T?j:i,dc=T?i:j,R=br+dr,C=bc+dc;
          if(R>=MAXR||C>=MAXC)continue;
          const sc=clp.cells[i][j],old=cellAt(cur,R,C);
          let nc=null;
          if(mode==='formats'){nc=sc&&sc.st?{...(old||{v:null}),st:sc.st}:(old?{...old,st:undefined}:null);}
          else{
            if(sc){
              let base={};
              if(sc.f!==undefined&&sc.f!==null&&mode!=='values'){base.f=sc.f.length?shiftFormula(sc.f,R-(clp.r1+i),C-(clp.c1+j)):sc.f;if(srcSh!==cur&&false)base.f=sc.f;}
              else if(sc.f!==undefined&&sc.f!==null){base.v=rawVal(wb,srcSh,clp.r1+i,clp.c1+j);if(isErr(base.v))base.v=base.v.code;}
              else base.v=sc.v;
              if(mode==='all'||mode==='transpose'||mode==='values'&&false)base.st=sc.st;
              else if(mode==='values'||mode==='formulas')base.st=old&&old.st;
              nc=base;
            }else if(mode==='all'||mode==='transpose')nc=null;
            else nc=old||null;
            if(mode==='values'&&sc&&sc.st&&sc.st.fmt&&(!nc.st||!nc.st.fmt))nc.st={...(nc.st||{}),fmt:sc.st.fmt};
          }
          putCell(cur,R,C,nc);
        }
        if(mode==='all'||mode==='transpose'){
          cur.merges=cur.merges.filter(m=>!(m.r1>=br&&m.r2<br+H&&m.c1>=bc&&m.c2<bc+Wd));
          for(const m of clp.merges){const n=T?{r1:br+(m.c1-clp.c1),r2:br+(m.c2-clp.c1),c1:bc+(m.r1-clp.r1),c2:bc+(m.r2-clp.r1)}:{r1:br+m.r1-clp.r1,r2:br+m.r2-clp.r1,c1:bc+m.c1-clp.c1,c2:bc+m.c2-clp.c1};cur.merges.push(n);}
        }
      }
      if(clp.cut&&clp.sh===cur&&mode==='all'){
        for(let i=0;i<h;i++)for(let j=0;j<w;j++){const R=clp.r1+i,C=clp.c1+j;if(R>=tr&&R<tr+H*rr&&C>=tc&&C<tc+Wd*cr)continue;putCell(cur,R,C,null);}
        cur.merges=cur.merges.filter(m=>!(m.r1>=clp.r1&&m.r2<=clp.r2&&m.c1>=clp.c1&&m.c2<=clp.c2&&!(m.r1>=tr&&m.r1<tr+H*rr&&m.c1>=tc&&m.c1<tc+Wd*cr)));
        clip=null;
      }
    });
    select(tr,tc,tr+H*rr-1,tc+Wd*cr-1);
  }
  function parseTSV(t){
    const rows=[];let row=[],cur2='',q=false;t=t.replace(/\r\n?/g,'\n');
    for(let i=0;i<t.length;i++){const ch=t[i];
      if(q){if(ch==='"'){if(t[i+1]==='"'){cur2+='"';i++;}else q=false;}else cur2+=ch;}
      else if(ch==='"'&&cur2==='')q=true;
      else if(ch==='\t'){row.push(cur2);cur2='';}
      else if(ch==='\n'){row.push(cur2);rows.push(row);row=[];cur2='';}
      else cur2+=ch;}
    if(cur2!==''||row.length){row.push(cur2);rows.push(row);}
    const w=Math.max(0,...rows.map(r=>r.length));rows.forEach(r=>{while(r.length<w)r.push('');});
    return rows;
  }
  root.addEventListener('copy',e=>{
    if(ed||document.activeElement===nb||/INPUT|TEXTAREA/.test(document.activeElement.tagName)&&document.activeElement.tagName!=='TEXTAREA'&&!vp.contains(document.activeElement))return;
    if(ed)return;
    e.preventDefault();const c=captureClip(false);e.clipboardData.setData('text/plain',c.text);schedule();
  });
  root.addEventListener('cut',e=>{
    if(ed||/INPUT|TEXTAREA/.test(document.activeElement.tagName))return;
    e.preventDefault();const c=captureClip(true);e.clipboardData.setData('text/plain',c.text);schedule();
  });
  root.addEventListener('paste',e=>{
    if(ed||/INPUT|TEXTAREA/.test(document.activeElement.tagName))return;
    e.preventDefault();const t=e.clipboardData.getData('text/plain');
    if(clip&&(t===clip.text||t===''))pasteMode('all');else pasteMode('all',undefined,undefined,null,t);
  });

  /* ===== penunjuk (mouse/sentuh) di lembar ===== */
  const hit=(ev)=>{const rc=vp.getBoundingClientRect();const x=ev.clientX-rc.left,y=ev.clientY-rc.top;
    return {x,y,c:cur.cw.at((x+vp.scrollLeft)/Z),r:cur.rh.at((y+vp.scrollTop)/Z)};};
  let drag=null,dragTimer=0,lastEv=null;
  function extendTo(ev){
    const h=hit(ev);
    const W=vp.clientWidth,H=vp.clientHeight;
    if(drag.kind==='cell'){
      let {r,c}=h;moveToDrag(r,c);
    }else if(drag.kind==='col'){moveToDrag(MAXR-1,h.c,true,'col');}
    else if(drag.kind==='row'){moveToDrag(h.r,MAXC-1,true,'row');}
    else if(drag.kind==='fill'){fillDrag(h);}
    else if(drag.kind==='point'){
      pointing.r2=h.r;pointing.c2=h.c;updPoint();}
    /* gulir otomatis */
    const dx=h.x<0?h.x:h.x>W?h.x-W:0,dy=h.y<0?h.y:h.y>H?h.y-H:0;
    if(dx||dy){vp.scrollLeft+=dx/3;vp.scrollTop+=dy/3;}
  }
  function moveToDrag(r,c,whole,kind){
    const s=cur.sel;
    if(kind==='col'){s.ar=0;s.fr=MAXR-1;s.ac=drag.c0;s.fc=c;}
    else if(kind==='row'){s.ac=0;s.fc=MAXC-1;s.ar=drag.r0;s.fr=r;}
    else{s.fr=r;s.fc=c;}
    sync();
  }
  vp.addEventListener('pointerdown',e=>{
    if(e.button!==0&&e.button!==2)return;
    const rc=vp.getBoundingClientRect();if(e.clientX-rc.left>vp.clientWidth||e.clientY-rc.top>vp.clientHeight)return;
    closePop();
    const fdd=e.target.closest&&e.target.closest('[data-fc]');
    if(fdd){e.preventDefault();filterMenu(+fdd.dataset.fc,fdd);return;}
    if(e.target.dataset&&e.target.dataset.fill){
      e.preventDefault();const g=expandMerge(rg());drag={kind:'fill',src:g,tgt:null};vp.setPointerCapture(e.pointerId);startDragLoop();return;}
    const h=hit(e);
    if(e.pointerType!=='mouse'&&e.button===0){return;} /* sentuhan: ditangani oleh click */
    if(e.button===2){const g=expandMerge(rg());if(!(h.r>=g.r1&&h.r<=g.r2&&h.c>=g.c1&&h.c<=g.c2))select(h.r,h.c);return;}
    /* mode tunjuk rumus */
    if(ed){
      if(canPoint()){e.preventDefault();startPoint(h.r,h.c);drag={kind:'point'};vp.setPointerCapture(e.pointerId);startDragLoop();return;}
      if(!commitEdit())return;
    }
    e.preventDefault();vp.focus({preventScroll:true});
    if(painting){const m=mergeAt(h.r,h.c);select(h.r,h.c);applyPaint();return;}
    const m=mergeAt(h.r,h.c);
    if(e.shiftKey){const s=cur.sel;s.fr=h.r;s.fc=h.c;sync();}
    else select(m?m.r1:h.r,m?m.c1:h.c,m?m.r2:h.r,m?m.c2:h.c,true),(m&&(cur.sel.ar=m.r1,cur.sel.ac=m.c1));
    drag={kind:'cell'};vp.setPointerCapture(e.pointerId);startDragLoop();
  });
  function startDragLoop(){clearInterval(dragTimer);dragTimer=setInterval(()=>{if(drag&&lastEv)extendTo(lastEv);},60);}
  vp.addEventListener('pointermove',e=>{
    lastEv=e;
    if(drag){extendTo(e);return;}
    if(e.target.dataset&&e.target.dataset.fill){vp.style.cursor='crosshair';return;}
    vp.style.cursor=painting?'copy':'cell';
  });
  const endDrag=e=>{
    if(!drag)return;clearInterval(dragTimer);
    const d=drag;drag=null;
    if(d.kind==='fill'&&d.tgt)doFill(d.src,d.tgt);
    if(d.kind==='point')setMode('Tunjuk');
    if(d.kind==='cell'&&painting&&false){}
    schedule();
  };
  vp.addEventListener('pointerup',endDrag);vp.addEventListener('pointercancel',endDrag);
  vp.addEventListener('click',e=>{
    if(e.pointerType==='mouse'||e.detail>0&&e.pointerType==='')return;
    if(e.target.closest&&e.target.closest('[data-fc]'))return;
    const rc=vp.getBoundingClientRect();if(e.clientX-rc.left>vp.clientWidth||e.clientY-rc.top>vp.clientHeight)return;
    const h=hit(e);
    if(ed){if(canPoint()){startPoint(h.r,h.c);return;}if(!commitEdit())return;}
    const s=cur.sel,m=mergeAt(h.r,h.c);
    const again=s.ar===(m?m.r1:h.r)&&s.ac===(m?m.c1:h.c);
    if(again&&Date.now()-lastTap<600){startEdit(s.ar,s.ac);lastTap=0;return;}
    lastTap=Date.now();
    if(painting){select(h.r,h.c);applyPaint();return;}
    select(m?m.r1:h.r,m?m.c1:h.c,m?m.r2:h.r,m?m.c2:h.c,true);
  });
  let lastTap=0;
  vp.addEventListener('dblclick',e=>{
    if(e.target.dataset&&e.target.dataset.fill)return;
    const rc=vp.getBoundingClientRect();if(e.clientX-rc.left>vp.clientWidth||e.clientY-rc.top>vp.clientHeight)return;
    const h=hit(e);if(ed)return;
    startEdit(h.r,h.c);
  });
  vp.addEventListener('contextmenu',e=>{
    e.preventDefault();const h=hit(e);
    const g=expandMerge(rg());if(!(h.r>=g.r1&&h.r<=g.r2&&h.c>=g.c1&&h.c<=g.c2))select(h.r,h.c);
    openMenuAt(e.clientX,e.clientY,menus.ctx());
  });
  vp.addEventListener('scroll',()=>{schedule();if(ed)placeEditor();});
  vp.addEventListener('wheel',e=>{if(e.ctrlKey){e.preventDefault();setZoom(Math.min(200,Math.max(20,Math.round((Z*100-Math.sign(e.deltaY)*10)/10)*10)));}},{passive:false});

  /* ===== isi otomatis (fill handle) ===== */
  function fillDrag(h){
    const s=drag.src;let tg={...s};
    const dr=h.r>s.r2?h.r-s.r2:h.r<s.r1?h.r-s.r1:0,dc=h.c>s.c2?h.c-s.c2:h.c<s.c1?h.c-s.c1:0;
    if(!dr&&!dc){drag.tgt=null;cur.sel.ar=s.r1;cur.sel.ac=s.c1;cur.sel.fr=s.r2;cur.sel.fc=s.c2;sync();return;}
    if(Math.abs(dr)*(1)>=Math.abs(dc)*1&&Math.abs(dr)>0&&(Math.abs(dr)>=Math.abs(dc))){ if(dr>0)tg.r2=s.r2+dr;else tg.r1=s.r1+dr;}
    else{if(dc>0)tg.c2=s.c2+dc;else tg.c1=s.c1+dc;}
    tg.r1=Math.max(0,tg.r1);tg.c1=Math.max(0,tg.c1);tg.r2=Math.min(MAXR-1,tg.r2);tg.c2=Math.min(MAXC-1,tg.c2);
    drag.tgt=tg;const q=cur.sel;q.ar=tg.r1;q.ac=tg.c1;q.fr=tg.r2;q.fc=tg.c2;sync();
  }
  const LISTS=[DAY,DAY.map(d=>d.slice(0,3)),MON,MON.map(m=>m.slice(0,3))];
  function seriesInfo(vals){ /* vals: array of cells values (raw) from the source line */
    const n=vals.length;
    const numsOk=vals.every(v=>typeof v==='number');
    if(numsOk&&n>=2){const step=(vals[n-1]-vals[0])/(n-1);if(vals.every((v,i)=>Math.abs(v-(vals[0]+step*i))<1e-9))return {k:'num',step,last:vals[n-1],first:vals[0]};return {k:'copy'};}
    if(numsOk&&n===1)return {k:'copy1'};
    if(vals.every(v=>typeof v==='string')){
      const li=LISTS.find(L=>vals.every(v=>L.some(x=>x.toLowerCase()===v.toLowerCase())));
      if(li){const idx=vals.map(v=>li.findIndex(x=>x.toLowerCase()===v.toLowerCase()));const step=n>1?(idx[n-1]-idx[0])||1:1;return {k:'list',L:li,last:idx[n-1],first:idx[0],step,cap:vals[n-1]};}
      const m=/^(.*?)(\d+)$/.exec(vals[n-1]);
      if(m&&vals.every(v=>/^(.*?)(\d+)$/.test(v)&&/^(.*?)(\d+)$/.exec(v)[1]===m[1])){
        const ns=vals.map(v=>+/(\d+)$/.exec(v)[1]);const step=n>1?ns[n-1]-ns[n-2]:1;return {k:'txt',pre:m[1],last:+m[2],first:+/(\d+)$/.exec(vals[0])[1],step,w:m[2].length};}
    }
    return {k:'copy'};
  }
  function doFill(src,tg){
    const vert=tg.r1!==src.r1||tg.r2!==src.r2;
    tx(()=>{
      if(vert){
        for(let c=src.c1;c<=src.c2;c++){
          const col=[];for(let r=src.r1;r<=src.r2;r++)col.push(cellAt(cur,r,c));
          const forward=tg.r2>src.r2;
          const info=seriesColumn(col,src.r1,c);
          const cnt=forward?tg.r2-src.r2:src.r1-tg.r1;
          for(let i=1;i<=cnt;i++){
            const R=forward?src.r2+i:src.r1-i,si=forward?(i-1)%col.length:col.length-1-((i-1)%col.length);
            putCell(cur,R,c,fillCell(col[si],info,forward?i:-i,R-(src.r1+si),0,si,forward?col.length:col.length));
          }
        }
      }else{
        for(let r=src.r1;r<=src.r2;r++){
          const row=[];for(let c=src.c1;c<=src.c2;c++)row.push(cellAt(cur,r,c));
          const forward=tg.c2>src.c2;
          const info=seriesColumn(row,r,src.c1);
          const cnt=forward?tg.c2-src.c2:src.c1-tg.c1;
          for(let i=1;i<=cnt;i++){
            const C=forward?src.c2+i:src.c1-i,si=forward?(i-1)%row.length:row.length-1-((i-1)%row.length);
            putCell(cur,r,C,fillCell(row[si],info,forward?i:-i,0,C-(src.c1+si),si,row.length));
          }
        }
      }
    });
    select(tg.r1,tg.c1,tg.r2,tg.c2);
  }
  function seriesColumn(line,r0,c0){
    if(line.some(c=>c&&c.f!==undefined&&c.f!==null))return {k:'copy'};
    const vals=line.map(c=>c?c.v:null);
    if(vals.some(v=>v===null||v===undefined))return {k:'copy'};
    const info=seriesInfo(vals);
    if(info.k==='copy1'){const f=line[0]&&line[0].st&&line[0].st.fmt;if(f&&isDateCode(f))return {k:'num',step:1,last:vals[0],first:vals[0]};return {k:'copy'};}
    return info;
  }
  function fillCell(src,info,i,dr,dc,si,len){
    if(!src)return null;
    const out={st:src.st};
    if(src.f!==undefined&&src.f!==null){out.f=shiftFormula(src.f,dr,dc);return out;}
    const base=(k)=>i>0?info.last+info.step*i:info.first+info.step*i;
    if(info.k==='num'){out.v=base();return out;}
    if(info.k==='list'){const n=info.L.length;out.v=info.L[(((base())%n)+n)%n];if(info.cap&&info.cap===info.cap.toUpperCase())out.v=out.v.toUpperCase();return out;}
    if(info.k==='txt'){const v=base();out.v=info.pre+String(Math.max(0,v)).padStart(info.w,'0');return out;}
    out.v=src.v;return out;
  }

  /* ===== tajuk: pilih / ubah ukuran ===== */
  let rsz=null;
  function hdrHit(ev,kind){
    const rc=(kind==='c'?colh:rowh).getBoundingClientRect();
    const p=kind==='c'?ev.clientX-rc.left+vp.scrollLeft:ev.clientY-rc.top+vp.scrollTop;
    const ax=kind==='c'?cur.cw:cur.rh;
    const i=ax.at(p/Z),edge=Math.round(ax.pos(i+1)*Z);
    let near=null;if(Math.abs(p-edge)<=4)near=i;else{const left=Math.round(ax.pos(i)*Z);if(i>0&&Math.abs(p-left)<=4)near=i-1;}
    return {i,edge:near};
  }
  [['c',colh],['r',rowh]].forEach(([kind,el])=>{
    el.addEventListener('pointermove',e=>{if(rsz)return;const h=hdrHit(e,kind);el.style.cursor=h.edge!==null?(kind==='c'?'col-resize':'row-resize'):'default';});
    el.addEventListener('pointerdown',e=>{
      if(e.button!==0)return;if(ed&&!commitEdit())return;
      e.preventDefault();closePop();vp.focus({preventScroll:true});
      const h=hdrHit(e,kind);
      if(h.edge!==null){rsz={kind,i:h.edge,start:kind==='c'?e.clientX:e.clientY,orig:(kind==='c'?cur.cw:cur.rh).size(h.edge),snap:snapshot(),moved:false};el.setPointerCapture(e.pointerId);return;}
      const s=cur.sel;
      if(kind==='c'){drag={kind:'col',c0:e.shiftKey?s.ac:h.i};if(e.shiftKey){s.fc=h.i;s.fr=MAXR-1;}else{s.ar=0;s.ac=h.i;s.fr=MAXR-1;s.fc=h.i;}}
      else{drag={kind:'row',r0:e.shiftKey?s.ar:h.i};if(e.shiftKey){s.fr=h.i;s.fc=MAXC-1;}else{s.ac=0;s.ar=h.i;s.fc=MAXC-1;s.fr=h.i;}}
      sync();el.setPointerCapture(e.pointerId);startDragLoop();lastEv=e;
    });
    el.addEventListener('pointermove',e=>{
      if(rsz&&rsz.kind===kind){const d=((kind==='c'?e.clientX:e.clientY)-rsz.start)/Z;const ax=kind==='c'?cur.cw:cur.rh;
        const v=Math.max(0,Math.round(rsz.orig+d));rsz.moved=true;if(v===0)ax.hid.add(rsz.i);else ax.hid.delete(rsz.i);ax.ov.set(rsz.i,v);ax.dirty=true;
        if(kind==='r'){if(!ax.manual)ax.manual=new Set();ax.manual.add(rsz.i);}
        schedule();return;}
      if(drag&&(drag.kind==='col'||drag.kind==='row')){lastEv=e;extendTo(e);}
    });
    const up=e=>{
      if(rsz&&rsz.kind===kind){if(rsz.moved){undoS.push(rsz.snap);redoS=[];dirty=true;updUndo();recalcAll();}rsz=null;sync();return;}
      if(drag&&(drag.kind==='col'||drag.kind==='row')){clearInterval(dragTimer);drag=null;}
    };
    el.addEventListener('pointerup',up);el.addEventListener('pointercancel',up);
    el.addEventListener('dblclick',e=>{const h=hdrHit(e,kind);if(h.edge!==null)(kind==='c'?autoFitCols([h.edge]):autoFitRowsH([h.edge]));});
  });
  $('.corner').addEventListener('pointerdown',e=>{e.preventDefault();if(ed&&!commitEdit())return;select(0,0,MAXR-1,MAXC-1,true);cur.sel.ar=cur.sel.ar;vp.focus();});

  function autoFitCols(cols){
    tx(()=>{for(const c of cols){let w=0;
      for(const [k,cl] of cur.cells){if(k%MAXC!==c)continue;const r=Math.floor(k/MAXC);if(mergeAtSh(cur,r,c))continue;const v=cellVal(cur,r,c);if(v===null||v==='')continue;
        const st=cl.st||{};const t=fmtValue(v,st.fmt).t;const sz=st.sz||11;
        const saveZ=Z;Z=1;const wd=tw(t.split('\n').sort((a,b)=>b.length-a.length)[0],{i:st.i,b:st.b,sz,font:st.font});Z=saveZ;w=Math.max(w,wd+8+(st.ind?st.ind*9:0));}
      if(w>0)cur.cw.ov.set(c,Math.min(Math.max(Math.ceil(w),12),600));else cur.cw.ov.delete(c);cur.cw.dirty=true;}});
  }
  function autoFitRowsH(rows){
    tx(()=>{for(const r of rows){if(cur.rh.manual)cur.rh.manual.delete(r);cur.rh.ov.delete(r);}fitRows(rows);});
  }

  /* ===== papan ketik (lembar) ===== */
  const nextEdge=(r,c,dr,dc)=>{
    const filled=(rr,cc)=>hasVal(cur,rr,cc);
    let a=r+dr,b=c+dc;
    if(a<0||b<0||a>=MAXR||b>=MAXC)return [r,c];
    if(filled(r,c)&&filled(a,b)){
      while(a+dr>=0&&b+dc>=0&&a+dr<MAXR&&b+dc<MAXC&&filled(a+dr,b+dc)){a+=dr;b+=dc;}
      return [a,b];
    }
    /* cari sel berisi berikutnya */
    const ur=usedRange(cur);
    while(a>=0&&b>=0&&a<MAXR&&b<MAXC&&!filled(a,b)){
      if((dr>0&&a>ur.r2)||(dr<0&&a<ur.r1)||(dc>0&&b>ur.c2)||(dc<0&&b<ur.c1))return dr>0?[MAXR-1,c]:dr<0?[0,c]:dc>0?[r,MAXC-1]:[r,0];
      a+=dr;b+=dc;}
    return [Math.max(0,Math.min(MAXR-1,a)),Math.max(0,Math.min(MAXC-1,b))];
  };
  function gridKey(e){
    const s=cur.sel,ctrl=e.ctrlKey||e.metaKey,k=e.key;
    const m=mergeAt(s.fr,s.fc);
    const foc={r:s.fr,c:s.fc};
    const step=(dr,dc)=>{
      let r=s.fr,c=s.fc;
      if(ctrl){const q=nextEdge(r,c,dr,dc);r=q[0];c=q[1];}
      else{const mg=mergeAt(r,c);if(mg){if(dr>0)r=mg.r2;if(dr<0)r=mg.r1;if(dc>0)c=mg.c2;if(dc<0)c=mg.c1;}
        r+=dr;c+=dc;
        for(let n=0;n<2000&&cur.rh.size(r)===0&&r>0&&r<MAXR-1;n++)r+=dr||0;
        for(let n=0;n<2000&&cur.cw.size(c)===0&&c>0&&c<MAXC-1;n++)c+=dc||0;}
      moveTo(r,c,e.shiftKey);
    };
    const pageR=Math.max(1,Math.floor(vp.clientHeight/(DH*Z))-1);
    switch(k){
      case 'ArrowUp':e.preventDefault();step(-1,0);return true;
      case 'ArrowDown':e.preventDefault();step(1,0);return true;
      case 'ArrowLeft':e.preventDefault();step(0,-1);return true;
      case 'ArrowRight':e.preventDefault();step(0,1);return true;
      case 'PageDown':e.preventDefault();moveTo(s.fr+pageR,s.fc,e.shiftKey);return true;
      case 'PageUp':e.preventDefault();moveTo(s.fr-pageR,s.fc,e.shiftKey);return true;
      case 'Home':e.preventDefault();ctrl?moveTo(0,0,e.shiftKey):moveTo(s.fr,0,e.shiftKey);return true;
      case 'End':e.preventDefault();if(ctrl){const u=usedRange(cur);moveTo(u.r2,u.c2,e.shiftKey);}else moveTo(s.fr,usedRange(cur).c2,e.shiftKey);return true;
      case 'Enter':e.preventDefault();moveSel(e.shiftKey?-1:1,0);return true;
      case 'Tab':e.preventDefault();moveSel(0,e.shiftKey?-1:1);return true;
      case 'F2':e.preventDefault();startEdit(s.ar,s.ac);return true;
      case 'Delete':e.preventDefault();acts.clearc();return true;
      case 'Backspace':e.preventDefault();startEdit(s.ar,s.ac,'');return true;
      case 'Escape':if(clip){clip=null;schedule();}if(painting){painting=null;vp.style.cursor='';}return true;
      case ' ':if(ctrl){e.preventDefault();const g=rg();select(0,g.c1,MAXR-1,g.c2,true);cur.sel.ar=0;cur.sel.ac=g.c1;return true;}
               if(e.shiftKey){e.preventDefault();const g=rg();select(g.r1,0,g.r2,MAXC-1,true);cur.sel.ar=g.r1;cur.sel.ac=0;return true;}break;
    }
    if(ctrl){
      const l=k.toLowerCase();
      if(l==='a'){e.preventDefault();select(0,0,MAXR-1,MAXC-1,true);return true;}
      if(l==='b'){e.preventDefault();acts.bold();return true;}
      if(l==='i'){e.preventDefault();acts.italic();return true;}
      if(l==='u'){e.preventDefault();acts.underline();return true;}
      if(l==='5'){e.preventDefault();acts.strike();return true;}
      if(l==='z'){e.preventDefault();e.shiftKey?redo():undo();return true;}
      if(l==='y'){e.preventDefault();redo();return true;}
      if(l==='d'){e.preventDefault();fillDir('down');return true;}
      if(l==='r'){e.preventDefault();fillDir('right');return true;}
      if(l==='f'||l==='h'){e.preventDefault();showFind();return true;}
      if(l==='g'){e.preventDefault();acts.goto();return true;}
      if(l==='1'){e.preventDefault();acts['fmtcells']('number');return true;}
      if(l==='s'){e.preventDefault();acts.save();return true;}
      if(l==='x'&&!e.shiftKey){return false;}
      if(k===';'){e.preventDefault();startEdit(s.ar,s.ac,fmtValue(today(),'dd/mm/yyyy').t);return true;}
      if(k==='+'||k==='='){if(e.shiftKey||k==='+'){e.preventDefault();acts.insdlg();return true;}}
      if(k==='-'){e.preventDefault();acts.deldlg();return true;}
      return false;
    }
    if(k.length===1&&!e.altKey){e.preventDefault();startEdit(s.ar,s.ac,k);return true;}
    return false;
  }
  vp.addEventListener('keydown',e=>{if(ed)return;gridKey(e);});

  /* ===== menu popup ===== */
  let pops=[];
  function closePop(){pops.forEach(p=>p.remove());pops=[];}
  function showMenu(items,x,y,level,owner){
    pops.slice(level).forEach(p=>p.remove());pops.length=level;
    const el=document.createElement('div');el.className='pop';
    el.innerHTML=items.map((it,i)=>it.sep?'<hr>':it.h?`<div class="mh">${it.h}</div>`:it.html!==undefined?`<div data-i="${i}">${it.html}</div>`:
      `<div class="mi${it.dis?' dis':''}" data-i="${i}"><i>${it.i||''}</i><span>${it.l}</span>${it.sub?'<span class="ar">▶</span>':''}${it.ck?'<span class="ck">✓</span>':''}</div>`).join('');
    X.appendChild(el);el._owner=owner;
    const xr=X.getBoundingClientRect(),w=el.offsetWidth,h=el.offsetHeight;
    let L=x,T=y;
    if(L+w>innerWidth-4)L=Math.max(4,(level?x-w-(w>0?el.offsetWidth*0:0)-150:innerWidth-w-4));
    if(T+h>innerHeight-4)T=Math.max(4,innerHeight-h-4);
    el.style.left=L+'px';el.style.top=T+'px';pops[level]=el;
    items.forEach((it,i)=>{if(it.on){const d=el.querySelector(`[data-i="${i}"]`);if(d)it.on(d);}});
    el.addEventListener('mouseover',e=>{
      const mi=e.target.closest('.mi');if(!mi)return;
      pops.slice(level+1).forEach(p=>p.remove());pops.length=level+1;
      const it=items[+mi.dataset.i];
      if(it&&it.sub){const r=mi.getBoundingClientRect();showMenu(typeof it.sub==='function'?it.sub():it.sub,r.right-2,r.top-3,level+1,owner);}
    });
    el.addEventListener('click',e=>{
      const mi=e.target.closest('.mi');if(!mi)return;const it=items[+mi.dataset.i];
      if(!it||it.sub||it.dis)return;closePop();if(it.a)it.a();
    });
    return el;
  }
  function openMenuAt(x,y,items,owner){closePop();showMenu(items,x,y,0,owner);}
  function openMenu(anchor,items){
    if(pops[0]&&pops[0]._owner===anchor){closePop();return;}
    const r=anchor.getBoundingClientRect();openMenuAt(r.left,r.bottom,items,anchor);
  }
  const docDown=e=>{
    if(!root.isConnected){document.removeEventListener('pointerdown',docDown,true);return;}
    if(!pops.length)return;
    if(e.target.closest&&e.target.closest('.pop'))return;
    if(pops[0]._owner&&pops[0]._owner.contains&&pops[0]._owner.contains(e.target))return;
    closePop();
  };
  document.addEventListener('pointerdown',docDown,true);

  /* palet warna */
  function palette(cb,opt){
    const rows=PAL.map(r=>`<div class="pr">${r.map(c=>`<button style="background:${c}" data-c="${c}" title="${c}"></button>`).join('')}</div>`).join('');
    const html=`<div class="pal">${opt.none?`<button class="po" data-n="1">${opt.none}</button>`:''}<div class="pl">Warna Tema</div>${rows}<div class="pl">Warna Standar</div><div class="pr">${PALSTD.map(c=>`<button style="background:${c}" data-c="${c}" title="${c}"></button>`).join('')}</div><button class="po" data-m="1">🎨 Warna Lainnya…</button><input type="color" style="position:absolute;opacity:0;pointer-events:none;width:1px;height:1px"></div>`;
    return [{html,on:d=>{
      d.addEventListener('click',e=>{
        const b=e.target.closest('button');if(!b)return;
        if(b.dataset.c){closePop();cb(b.dataset.c);}
        else if(b.dataset.n){closePop();cb(null);}
        else if(b.dataset.m){const i=d.querySelector('input[type=color]');i.onchange=()=>{closePop();cb(i.value);};i.click();}
      });}}];
  }

  /* ===== penerapan gaya ===== */
  const cleanSt=st=>{
    const o={};
    for(const k in st){const v=st[k];if(v===undefined||v===null||v===false)continue;
      if(k==='bd'){const b={};for(const s of SIDES)if(v[s])b[s]=v[s];if(Object.keys(b).length)o.bd=b;continue;}
      if(k==='fmt'&&v==='General')continue;if(k==='ha'&&v==='general')continue;if(k==='va'&&v==='bottom'&&false)continue;if(k==='ind'&&!v)continue;if(k==='rot'&&!v)continue;
      o[k]=v;}
    return Object.keys(o).length?o:undefined;
  };
  function styleSel(fn,rowsFit){
    const g=expandMerge(limitRange(rg()));
    return tx(()=>{
      for(let r=g.r1;r<=g.r2;r++)for(let c=g.c1;c<=g.c2;c++){
        const cl=cellAt(cur,r,c),st=fn((cl&&cl.st)||{},r,c,g);
        const ns=cleanSt(st);
        if(!cl&&!ns)continue;
        putCell(cur,r,c,{...(cl||{v:null}),st:ns});
      }
      if(rowsFit){const rs=[];for(let r=g.r1;r<=g.r2;r++)rs.push(r);if(rs.length<5000)fitRows(rs);}
    });
  }
  const applyStyle=(patch,rowsFit)=>styleSel(st=>({...st,...patch}),rowsFit);
  const colorState={fill:'#ffff00',color:'#ff0000'};
  const setSw=(w,c)=>$$(`.sw[data-w=${w}]`).forEach(e=>{e.style.setProperty('--sw',c);e.querySelectorAll('.sw').forEach(x=>x.style.setProperty('--sw',c));});
  let bdColor='#000000',bdStyle='thin';

  /* ===== batas ===== */
  function applyBorderSpec(sp){
    const g=expandMerge(limitRange(rg()));
    const perim=(r,c)=>({t:r===g.r1,b:r===g.r2,l:c===g.c1,r:c===g.c2});
    tx(()=>{
      for(let r=g.r1;r<=g.r2;r++)for(let c=g.c1;c<=g.c2;c++){
        const p=perim(r,c);const cl=cellAt(cur,r,c),st=(cl&&cl.st)||{};const bd={...(st.bd||{})};let ch=false;
        const set=(s,v)=>{if(v===null)delete bd[s];else bd[s]=v;ch=true;};
        for(const s of SIDES)if(sp[s]!==undefined&&p[s])set(s,sp[s]);
        if(sp.h!==undefined){if(r<g.r2)set('b',sp.h);if(r>g.r1)set('t',sp.h);}
        if(sp.v!==undefined){if(c<g.c2)set('r',sp.v);if(c>g.c1)set('l',sp.v);}
        if(ch){const ns=cleanSt({...st,bd});if(ns||cl)putCell(cur,r,c,{...(cl||{v:null}),st:ns});}
      }
      /* sisi tetangga yang menghadap rentang ikut dihapus bila memilih "tanpa tepi" */
      const clearN=(r,c,s)=>{if(r<0||c<0||r>=MAXR||c>=MAXC)return;const cl=cellAt(cur,r,c);if(!cl||!cl.st||!cl.st.bd||!cl.st.bd[s])return;const bd={...cl.st.bd};delete bd[s];putCell(cur,r,c,{...cl,st:cleanSt({...cl.st,bd})});};
      for(let r=g.r1;r<=g.r2;r++){if(sp.l===null)clearN(r,g.c1-1,'r');if(sp.r===null)clearN(r,g.c2+1,'l');}
      for(let c=g.c1;c<=g.c2;c++){if(sp.t===null)clearN(g.r1-1,c,'b');if(sp.b===null)clearN(g.r2+1,c,'t');}
    });
  }
  function borderKind(k){
    const v=[bdStyle,bdColor],th=['thick',bdColor],dbl=['double',bdColor],md=['medium',bdColor];
    const M={bottom:{b:v},top:{t:v},left:{l:v},right:{r:v},none:{t:null,b:null,l:null,r:null,h:null,v:null},all:{t:v,b:v,l:v,r:v,h:v,v:v},outer:{t:v,b:v,l:v,r:v},
      thickOuter:{t:th,b:th,l:th,r:th},dblBottom:{b:dbl},thickBottom:{b:th},medBottom:{b:md},topBottom:{t:v,b:v},topThickBottom:{t:v,b:th},topDblBottom:{t:v,b:dbl}};
    if(k==='last')k=borderKind.last||'bottom';
    if(M[k]){borderKind.last=k;applyBorderSpec(M[k]);}
  }

  /* ===== gabung sel ===== */
  function mergeSel(kind){
    const g=expandMerge(rg());
    const same=cur.merges.find(m=>m.r1===g.r1&&m.r2===g.r2&&m.c1===g.c1&&m.c2===g.c2);
    if(kind==='un'||(same&&kind==='center')){
      tx(()=>{cur.merges=cur.merges.filter(m=>!(m.r1>=g.r1&&m.r2<=g.r2&&m.c1>=g.c1&&m.c2<=g.c2));if(same&&kind==='center'){const cl=cellAt(cur,g.r1,g.c1);if(cl&&cl.st&&cl.st.ha==='center')putCell(cur,g.r1,g.c1,{...cl,st:cleanSt({...cl.st,ha:null})});}});
      return;
    }
    if(g.r1===g.r2&&g.c1===g.c2)return;
    let lost=false;
    for(let r=g.r1;r<=g.r2&&!lost;r++)for(let c=g.c1;c<=g.c2;c++)if((r!==g.r1||c!==g.c1)&&hasVal(cur,r,c)){lost=true;break;}
    const run=()=>tx(()=>{
      cur.merges=cur.merges.filter(m=>!(m.r2<g.r1||m.r1>g.r2||m.c2<g.c1||m.c1>g.c2));
      const parts=kind==='across'?Array.from({length:g.r2-g.r1+1},(_,i)=>({r1:g.r1+i,r2:g.r1+i,c1:g.c1,c2:g.c2})):[g];
      for(const p of parts){
        for(let r=p.r1;r<=p.r2;r++)for(let c=p.c1;c<=p.c2;c++){if(r===p.r1&&c===p.c1)continue;const cl=cellAt(cur,r,c);putCell(cur,r,c,cl&&cl.st?{v:null,st:cl.st}:null);}
        if(p.c1!==p.c2||p.r1!==p.r2)cur.merges.push({...p});
        if(kind==='center'){const cl=cellAt(cur,p.r1,p.c1);putCell(cur,p.r1,p.c1,{...(cl||{v:null}),st:cleanSt({...(cl&&cl.st||{}),ha:'center'})});}
      }
    });
    if(lost)confirmDlg('Menggabungkan sel hanya mempertahankan nilai dari sel kiri atas. Nilai lainnya akan dibuang.',run);else run();
  }

  /* ===== struktur: sisip/hapus ===== */
  function adjRefsAll(kind,at,n,del,target){
    for(const sh of wb.sheets){
      for(const [k,cell] of sh.cells){
        if(cell.f===undefined||cell.f===null)continue;
        const nf=rewriteRefs(cell.f,(a,b,sn)=>{
          const on=sn?sn.toLowerCase()===target.name.toLowerCase():sh===target;
          if(!on)return [a,b];
          const idx=kind==='r'?'r':'c';
          const fx=(x)=>{const v=x[idx];
            if(!del)return v>=at?v+n:v;
            if(v<at)return v;if(v<at+n)return -1;return v-n;};
          if(!b){const v=fx(a);if(v<0)return null;return [{...a,[idx]:v},null];}
          let lo=a[idx],hi=b[idx];
          if(!del){if(lo>=at)lo+=n;if(hi>=at)hi+=n;}
          else{const e=at+n-1;lo=lo<at?lo:lo<=e?at:lo-n;hi=hi<at?hi:hi<=e?at-1:hi-n;if(hi<lo)return null;}
          return [{...a,[idx]:lo},{...b,[idx]:hi}];
        });
        if(nf!==cell.f)sh.cells.set(k,{...cell,f:nf,ast:undefined});
      }
    }
  }
  function remapAxis(ax,at,n,del){
    const f=i=>!del?(i>=at?i+n:i):(i<at?i:i<at+n?-1:i-n);
    const o=new Map();for(const [i,v] of ax.ov){const j=f(i);if(j>=0)o.set(j,v);}ax.ov=o;
    ax.hid=new Set([...ax.hid].map(f).filter(j=>j>=0));ax.fil=new Set([...ax.fil].map(f).filter(j=>j>=0));
    if(ax.manual)ax.manual=new Set([...ax.manual].map(f).filter(j=>j>=0));
    ax.dirty=true;
  }
  function structOp(kind,at,n,del){
    tx(()=>{
      adjRefsAll(kind,at,n,del,cur);
      const nm=new Map();
      for(const [k,cell] of cur.cells){
        let r=Math.floor(k/MAXC),c=k%MAXC;const v=kind==='r'?r:c;
        let nv=!del?(v>=at?v+n:v):(v<at?v:v<at+n?-1:v-n);
        if(nv<0)continue;
        if(kind==='r')r=nv;else c=nv;
        if(r>=MAXR||c>=MAXC)continue;
        nm.set(key(r,c),cell);
      }
      cur.cells=nm;
      remapAxis(kind==='r'?cur.rh:cur.cw,at,n,del);
      const adj=(o,lo,hi)=>{
        const f=i=>!del?(i>=at?i+n:i):(i<at?i:i<at+n?-1:i-n);
        let a=o[lo],b=o[hi];
        if(!del){if(a>=at)a+=n;if(b>=at)b+=n;}
        else{const e=at+n-1;a=a<at?a:a<=e?at:a-n;b=b<at?b:b<=e?at-1:b-n;if(b<a)return null;}
        return {...o,[lo]:a,[hi]:b};
      };
      const lo=kind==='r'?'r1':'c1',hi=kind==='r'?'r2':'c2';
      cur.merges=cur.merges.map(m=>adj(m,lo,hi)).filter(Boolean).filter(m=>m.r1!==m.r2||m.c1!==m.c2);
      cur.cf=cur.cf.map(rl=>adj(rl,lo,hi)).filter(Boolean);
      if(cur.filter){const f=adj(cur.filter,lo,hi);cur.filter=f;}
      const s=cur.sel;
      if(!del){if(kind==='r'){s.ar=at;s.fr=at+n-1;s.ac=0;s.fc=MAXC-1;}else{s.ac=at;s.fc=at+n-1;s.ar=0;s.fr=MAXR-1;}}
      else{if(kind==='r'){s.fr=s.ar=Math.min(at,MAXR-1);}else{s.fc=s.ac=Math.min(at,MAXC-1);}}
      if(cur.filter)applyFilter(true);
    });
  }
  function shiftCells(dir,ins){ /* geser sebagian sel (tanpa menyesuaikan rumus lain) */
    const g=expandMerge(rg());const h=g.r2-g.r1+1,w=g.c2-g.c1+1;
    tx(()=>{
      const nm=new Map(cur.cells);
      const moved=[];
      for(const [k,cell] of cur.cells){const r=Math.floor(k/MAXC),c=k%MAXC;
        if(dir==='down'&&c>=g.c1&&c<=g.c2&&(ins?r>=g.r1:r>g.r2)){nm.delete(k);moved.push([ins?r+h:r-h,c,cell]);}
        else if(dir==='right'&&r>=g.r1&&r<=g.r2&&(ins?c>=g.c1:c>g.c2)){nm.delete(k);moved.push([r,ins?c+w:c-w,cell]);}}
      if(!ins){for(let r=g.r1;r<=g.r2;r++)for(let c=g.c1;c<=g.c2;c++)nm.delete(key(r,c));}
      for(const [r,c,cell] of moved)if(r<MAXR&&c<MAXC&&r>=0&&c>=0)nm.set(key(r,c),cell);
      cur.cells=nm;
    });
  }
  function setHidden(kind,on){
    const g=rg();tx(()=>{
      const ax=kind==='r'?cur.rh:cur.cw,a=kind==='r'?g.r1:g.c1,b=kind==='r'?g.r2:g.c2;
      if(on){for(let i=a;i<=b&&i-a<5000;i++)ax.hid.add(i);}
      else{for(let i=Math.max(0,a-1);i<=Math.min((kind==='r'?MAXR:MAXC)-1,b+1)&&i-a<5000;i++)ax.hid.delete(i);}
      ax.dirty=true;});
  }

  /* ===== urut & filter ===== */
  function region(){ /* wilayah data di sekitar sel aktif */
    const g=rg();if(g.r1!==g.r2||g.c1!==g.c2)return g;
    let r1=g.r1,r2=g.r1,c1=g.c1,c2=g.c1;
    if(!hasVal(cur,r1,c1)){return g;}
    let ch=true;
    const colHas=(c,a,b)=>{for(let r=a;r<=b;r++)if(hasVal(cur,r,c))return true;return false;};
    const rowHas=(r,a,b)=>{for(let c=a;c<=b;c++)if(hasVal(cur,r,c))return true;return false;};
    while(ch){ch=false;
      if(c1>0&&colHas(c1-1,r1,r2)){c1--;ch=true;}if(c2<MAXC-1&&colHas(c2+1,r1,r2)){c2++;ch=true;}
      if(r1>0&&rowHas(r1-1,c1,c2)){r1--;ch=true;}if(r2<MAXR-1&&rowHas(r2+1,c1,c2)){r2++;ch=true;}
      if(r2-r1>200000)break;}
    return {r1,r2,c1,c2};
  }
  function sortRange(g,col,asc,header){
    const r1=g.r1+(header?1:0);if(r1>=g.r2)return;
    tx(()=>{
      const rows=[];
      for(let r=r1;r<=g.r2;r++){const line=[];for(let c=g.c1;c<=g.c2;c++)line.push(cellAt(cur,r,c));rows.push({r,line,k:cellVal(cur,r,col)});}
      const rank=v=>v===null||v===''?3:typeof v==='number'?0:typeof v==='string'?1:2;
      rows.sort((a,b)=>{const ra=rank(a.k),rb=rank(b.k);if(ra!==rb)return ra-rb;if(ra===3)return a.r-b.r;
        let c=ra===0?a.k-b.k:ra===1?String(a.k).localeCompare(String(b.k),'id',{sensitivity:'base'}):(+a.k)-(+b.k);if(!asc&&ra!==3)c=-c;return c||a.r-b.r;});
      rows.forEach((row,i)=>{const R=r1+i;row.line.forEach((cl,j)=>{
        let nc=cl;if(cl&&cl.f!==undefined&&cl.f!==null&&R!==row.r)nc={...cl,f:shiftFormula(cl.f,R-row.r,0),ast:undefined};
        putCell(cur,R,g.c1+j,nc||null);});});
    });
    if(cur.filter)applyFilter(true);
  }
  function quickSort(asc){
    const s=cur.sel,g=region();
    if(g.r1===g.r2&&g.c1===g.c2){toast('Pilih sel di dalam data yang akan diurutkan.');return;}
    const hdr=guessHeader(g);sortRange(g,Math.min(Math.max(s.ac,g.c1),g.c2),asc,hdr);
  }
  function guessHeader(g){
    if(g.r2<=g.r1)return false;
    for(let c=g.c1;c<=g.c2;c++){const a=cellVal(cur,g.r1,c),b=cellVal(cur,g.r1+1,c);if(typeof a==='string'&&typeof b==='number')return true;}
    let allText=true;for(let c=g.c1;c<=g.c2;c++)if(typeof cellVal(cur,g.r1,c)!=='string')allText=false;
    return allText;
  }
  function toggleFilter(){
    if(cur.filter){tx(()=>{cur.filter=null;cur.rh.fil=new Set();cur.rh.dirty=true;});return;}
    const g=region();if(g.r1===g.r2&&g.c1===g.c2){toast('Pilih sel di dalam data untuk menerapkan filter.');return;}
    tx(()=>{cur.filter={r1:g.r1,c1:g.c1,r2:g.r2,c2:g.c2,crit:{}};});
  }
  function applyFilter(inTx){
    const f=cur.filter;const hid=new Set();
    if(f){for(let r=f.r1+1;r<=f.r2;r++){
      for(const c in f.crit){const set=f.crit[c];const v=cellVal(cur,r,+c);const t=v===null||v===''?'(Kosong)':fmtValue(v,(stOf(r,+c)||{}).fmt).t;if(!set.has(t)){hid.add(r);break;}}}}
    cur.rh.fil=hid;cur.rh.dirty=true;if(!inTx)sync();
  }
  function filterMenu(c,btn){
    const f=cur.filter;if(!f)return;
    const vals=new Map();
    for(let r=f.r1+1;r<=f.r2;r++){const v=cellVal(cur,r,c);const t=v===null||v===''?'(Kosong)':fmtValue(v,(stOf(r,c)||{}).fmt).t;vals.set(t,v);}
    const list=[...vals.keys()].sort((a,b)=>{const x=vals.get(a),y=vals.get(b);if(typeof x==='number'&&typeof y==='number')return x-y;return a.localeCompare(b,'id');});
    const cur0=f.crit[c];
    const html=`<div style="padding:6px 8px;width:220px"><input type="text" class="fs" placeholder="Cari" style="width:100%;height:24px;border:1px solid #aaa;margin-bottom:4px"><label style="display:flex;gap:6px"><input type="checkbox" class="fa"> (Pilih Semua)</label><div class="fl">${list.map((t,i)=>`<label><input type="checkbox" data-t="${i}" ${!cur0||cur0.has(t)?'checked':''}> ${esc(t)}</label>`).join('')}</div><div style="display:flex;gap:6px;justify-content:flex-end;margin-top:6px"><button class="bt ok" style="background:#217346;color:#fff;border:1px solid #217346;height:24px;padding:0 12px;cursor:pointer">OK</button><button class="bt no" style="height:24px;padding:0 10px;border:1px solid #aaa;background:#f3f3f3;cursor:pointer">Batal</button></div></div>`;
    const r=btn.getBoundingClientRect();
    const items=[{l:'Urutkan A ke Z',i:'↓',a:()=>sortRange({r1:f.r1,r2:f.r2,c1:f.c1,c2:f.c2},c,true,true)},{l:'Urutkan Z ke A',i:'↑',a:()=>sortRange({r1:f.r1,r2:f.r2,c1:f.c1,c2:f.c2},c,false,true)},
      {l:'Hapus Filter dari kolom ini',i:'⧩',dis:!cur0,a:()=>{tx(()=>{delete f.crit[c];applyFilter(true);});}},{sep:1},
      {html,on:d=>{
        const boxes=[...d.querySelectorAll('[data-t]')],all=d.querySelector('.fa');
        const upd=()=>{all.checked=boxes.every(b=>b.checked);};upd();
        all.onchange=()=>boxes.forEach(b=>{if(!b.parentElement.hidden)b.checked=all.checked;});
        boxes.forEach(b=>b.onchange=upd);
        d.querySelector('.fs').oninput=e=>{const q=e.target.value.toLowerCase();boxes.forEach(b=>b.parentElement.hidden=!list[+b.dataset.t].toLowerCase().includes(q));};
        d.querySelector('.no').onclick=closePop;
        d.querySelector('.ok').onclick=()=>{const chosen=new Set(boxes.filter(b=>b.checked).map(b=>list[+b.dataset.t]));closePop();
          tx(()=>{if(chosen.size===list.length)delete f.crit[c];else f.crit[c]=chosen;applyFilter(true);});};
      }}];
    openMenuAt(r.left,r.bottom,items,btn);
  }

  /* ===== cari & ganti ===== */
  const f1=$('.f1'),f2=$('.f2');
  const showFind=()=>{fd.hidden=false;f1.focus();f1.select();};
  function matcher(){
    const t=f1.value;if(!t)return null;const cs=$('.fcase').checked,wh=$('.fwhole').checked;
    const q=cs?t:t.toLowerCase();
    return s=>{const x=cs?s:s.toLowerCase();return wh?x===q:x.includes(q);};
  }
  function findNext(){
    const m=matcher();if(!m)return false;
    const ks=[...cur.cells.keys()].sort((a,b)=>a-b);if(!ks.length){toast('Tidak ditemukan');return false;}
    const s=cur.sel,here=key(s.ar,s.ac);
    const test=k=>{const r=Math.floor(k/MAXC),c=k%MAXC,cl=cur.cells.get(k);const v=cellVal(cur,r,c);const t=fmtValue(v,cl.st&&cl.st.fmt).t;return m(t)||m(rawText(cl));};
    let start=ks.findIndex(k=>k>here);if(start<0)start=0;
    for(let i=0;i<ks.length;i++){const k=ks[(start+i)%ks.length];if(test(k)){select(Math.floor(k/MAXC),k%MAXC);return true;}}
    toast('Tidak ditemukan');return false;
  }
  function replaceOne(){
    const m=matcher();if(!m)return;const s=cur.sel,cl=cellAt(cur,s.ar,s.ac);
    if(cl&&m(rawText(cl))){const re=new RegExp(f1.value.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),$('.fcase').checked?'g':'gi');
      const nt=$('.fwhole').checked?f2.value:rawText(cl).replace(re,()=>f2.value);
      tx(()=>putCell(cur,s.ar,s.ac,mkInput(nt,cl)));}
    findNext();
  }
  function replaceAll(){
    const m=matcher();if(!m)return;let n=0;
    const re=new RegExp(f1.value.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),$('.fcase').checked?'g':'gi');
    tx(()=>{for(const [k,cl] of [...cur.cells]){const t=rawText(cl);if(!m(t))continue;const nt=$('.fwhole').checked?f2.value:t.replace(re,()=>f2.value);n++;putCell(cur,Math.floor(k/MAXC),k%MAXC,mkInput(nt,cl));}});
    toast(n+' penggantian dibuat');
  }
  f1.addEventListener('keydown',e=>{if(e.key==='Enter')findNext();if(e.key==='Escape')fd.hidden=true;});
  f2.addEventListener('keydown',e=>{if(e.key==='Enter')replaceOne();if(e.key==='Escape')fd.hidden=true;});

  /* ===== isi arah ===== */
  function fillDir(dir){
    const g=expandMerge(rg());
    if(g.r1===g.r2&&(dir==='down'||dir==='up')||g.c1===g.c2&&(dir==='right'||dir==='left')){
      /* satu baris/kolom: salin dari sel di atas/kiri */
      if(dir==='down'&&g.r1>0){g.r1--;}else if(dir==='right'&&g.c1>0){g.c1--;}else if(dir==='up'&&g.r2<MAXR-1){g.r2++;}else if(dir==='left'&&g.c2<MAXC-1){g.c2++;}
      else return;
    }
    tx(()=>{
      if(dir==='down'||dir==='up'){
        for(let c=g.c1;c<=g.c2;c++){const src=cellAt(cur,dir==='down'?g.r1:g.r2,c),sr=dir==='down'?g.r1:g.r2;
          for(let r=g.r1;r<=g.r2;r++){if(r===sr)continue;putCell(cur,r,c,src?{...cloneCell(src),f:src.f!==undefined&&src.f!==null?shiftFormula(src.f,r-sr,0):undefined}:null);if(src&&(src.f===undefined||src.f===null)){const x=cellAt(cur,r,c);delete x.f;}}}
      }else{
        for(let r=g.r1;r<=g.r2;r++){const sc=dir==='right'?g.c1:g.c2,src=cellAt(cur,r,sc);
          for(let c=g.c1;c<=g.c2;c++){if(c===sc)continue;putCell(cur,r,c,src?{...cloneCell(src),f:src.f!==undefined&&src.f!==null?shiftFormula(src.f,0,c-sc):undefined}:null);if(src&&(src.f===undefined||src.f===null)){const x=cellAt(cur,r,c);delete x.f;}}}
      }
    });
    select(g.r1,g.c1,g.r2,g.c2);
  }

  /* ===== format painter ===== */
  function applyPaint(){
    const p=painting;painting=null;vp.style.cursor='';if(!p)return;
    styleSel(()=>({...p}));
  }
  /* ===== zoom ===== */
  function setZoom(v){
    const old=Z;Z=Math.min(Math.max(v,20),200)/100;
    const sl=vp.scrollLeft/old,st=vp.scrollTop/old;
    $('.zr').value=Math.round(Z*100);$('.zv').textContent=Math.round(Z*100)+'%';
    draw();vp.scrollLeft=sl*Z;vp.scrollTop=st*Z;schedule();
  }

  /* ===== lembar ===== */
  function drawTabs(){
    const box=$('.shs');
    box.innerHTML=wb.sheets.map((s,i)=>`<div class="sht${i===wb.active?' on':''}" data-sh="${i}"${s.tab?` style="border-bottom:3px solid ${s.tab}"`:''}>${esc(s.name)}</div>`).join('');
    const on=box.querySelector('.on');if(on)on.scrollIntoView({block:'nearest',inline:'nearest'});
  }
  function switchSheet(i){
    if(i===wb.active&&cur)return;if(ed&&!commitEdit())return;
    cur.sx=vp.scrollLeft;cur.sy=vp.scrollTop;
    wb.active=i;cur=wb.sheets[i];clip=clip&&clip.sh===cur?clip:clip;
    drawTabs();sync();
    vp.scrollLeft=cur.sx;vp.scrollTop=cur.sy;schedule();
  }
  function addSheet(){
    tx(()=>{const s=newSheet(uniqueName('Sheet'+(wb.sheets.length+1)));wb.sheets.splice(wb.active+1,0,s);wb.active++;cur=s;},true);
    drawTabs();sync();
  }
  function renameSheet(i){
    const el=$(`.sht[data-sh="${i}"]`);if(!el)return;const s=wb.sheets[i];
    el.innerHTML=`<input value="${esc(s.name)}" maxlength="31">`;const inp=el.querySelector('input');inp.focus();inp.select();
    let done=false;
    const fin=ok=>{if(done)return;done=true;const v=inp.value.trim();
      if(ok&&v&&v!==s.name){if(wb.sheets.some(x=>x!==s&&x.name.toLowerCase()===v.toLowerCase())||/[\\\/?*\[\]:]/.test(v)){toast('Nama lembar tidak valid atau sudah dipakai.');}
        else tx(()=>{const old=s.name;s.name=v;
          for(const sh of wb.sheets)for(const [k,cl] of sh.cells){if(cl.f&&cl.f.includes('!')){const nf=cl.f.replace(new RegExp("(^|[^\\w.'])(?:'"+old.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+"'|"+old.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+")!",'g'),(m,p)=>p+(/^[A-Za-z_][\w.]*$/.test(v)?v:"'"+v.replace(/'/g,"''")+"'")+'!');if(nf!==cl.f)sh.cells.set(k,{...cl,f:nf,ast:undefined});}}},true);}
      drawTabs();sync();};
    inp.onkeydown=e=>{e.stopPropagation();if(e.key==='Enter')fin(true);else if(e.key==='Escape')fin(false);};
    inp.onblur=()=>fin(true);
  }
  function deleteSheet(i){
    if(wb.sheets.length<2){toast('Buku kerja harus memiliki minimal satu lembar.');return;}
    const run=()=>{tx(()=>{wb.sheets.splice(i,1);wb.active=Math.min(wb.active,wb.sheets.length-1);if(i<wb.active||i===wb.active)wb.active=Math.max(0,Math.min(i,wb.sheets.length-1));cur=wb.sheets[wb.active];},true);drawTabs();sync();};
    const s=wb.sheets[i];if(s.cells.size)confirmDlg(`Hapus lembar "${s.name}" secara permanen? Tindakan ini bisa dibatalkan dengan Urungkan.`,run);else run();
  }
  function moveSheet(i,d){
    const j=i+d;if(j<0||j>=wb.sheets.length)return;
    tx(()=>{const s=wb.sheets.splice(i,1)[0];wb.sheets.splice(j,0,s);wb.active=j;cur=s;},true);drawTabs();sync();
  }

  /* ===== dialog ===== */
  function dlg(title,body,ok,okText,noCancel){
    const d=document.createElement('div');d.className='ov';
    d.innerHTML=`<div class="dg"><h3>${title}</h3><div class="db">${body}</div><div class="dbt">${noCancel?'':'<button class="x">Batal</button>'}<button class="ok">${okText||'OK'}</button></div></div>`;
    X.appendChild(d);const xb=d.querySelector('.x');if(xb)xb.onclick=()=>d.remove();
    d.querySelector('.ok').onclick=()=>{if(!ok||ok(d)!==false)d.remove();};
    d.addEventListener('keydown',e=>{e.stopPropagation();if(e.key==='Escape')d.remove();else if(e.key==='Enter'&&e.target.tagName!=='TEXTAREA'&&e.target.tagName!=='SELECT'){d.querySelector('.ok').click();}});
    const f=d.querySelector('input,select');if(f)f.focus();
    return d;
  }
  const confirmDlg=(msg,fn)=>dlg('Microsoft Excel',`<div style="max-width:320px">${esc(msg)}</div>`,()=>{fn();});
  const num0=(d,s)=>+d.querySelector(s).value;
  function askNumber(title,label,val,fn){
    dlg(title,`<label>${label} <input type="number" class="v" value="${val}" step="any" style="width:90px"></label>`,d=>{const v=num0(d,'.v');if(!(v>=0)||isNaN(v))return false;fn(v);});
  }

  /* ===== aksi ===== */
  const SZ=()=>stOf(cur.sel.ar,cur.sel.ac).sz||11;
  const togg=(k,v)=>{const on=!stOf(cur.sel.ar,cur.sel.ac)[k];applyStyle({[k]:on?(v===undefined?1:v):null},k==='sz');};
  function decimals(d){
    const s=cur.sel;const st=stOf(s.ar,s.ac);let code=st.fmt||'General';
    if(isDateCode(code)||code==='@')return;
    if(code==='General'){const v=cellVal(cur,s.ar,s.ac);if(typeof v!=='number')return;const t=fmtValue(v,'General').t;const n=t.includes(',')&&!/E/.test(t)?t.split(',')[1].length:0;code=n?'0.'+'0'.repeat(n):'0';}
    const m=/0(\.(0+))?/.exec(code);if(!m){return;}
    const cur0=m[2]?m[2].length:0,n=Math.max(0,Math.min(30,cur0+d));
    applyStyle({fmt:code.replace(/0(\.0+)?/,n?'0.'+'0'.repeat(n):'0')});
  }
  function autoSum(fn){
    const g=rg(),s=cur.sel;
    tx(()=>{
      const put=(r,c,f)=>{const o=cellAt(cur,r,c);putCell(cur,r,c,{f,st:o&&o.st});};
      if(g.r1===g.r2&&g.c1===g.c2){
        let r=s.ar-1;while(r>=0&&typeof cellVal(cur,r,s.ac)==='number')r--;
        let a=r+1,b=s.ar-1,col=true;
        if(b<a){let c=s.ac-1;while(c>=0&&typeof cellVal(cur,s.ar,c)==='number')c--;a=c+1;b=s.ac-1;col=false;}
        if(b<a){put(s.ar,s.ac,`${fn}()`);return;}
        put(s.ar,s.ac,`${fn}(${col?addr(a,s.ac)+':'+addr(b,s.ac):addr(s.ar,a)+':'+addr(s.ar,b)})`);
      }else if(g.r2>g.r1){for(let c=g.c1;c<=g.c2;c++)if(g.r2+1<MAXR)put(g.r2+1,c,`${fn}(${addr(g.r1,c)}:${addr(g.r2,c)})`);}
      else if(g.c2+1<MAXC)put(g.r1,g.c2+1,`${fn}(${addr(g.r1,g.c1)}:${addr(g.r1,g.c2)})`);
    });
  }
  function clearWhat(w){
    const g=expandMerge(limitRange(rg()));
    tx(()=>{
      if(w==='all'||w==='formats')cur.merges=cur.merges.filter(m=>w==='formats'&&false||!(m.r1>=g.r1&&m.r2<=g.r2&&m.c1>=g.c1&&m.c2<=g.c2));
      if(g.r2-g.r1>2000&&w!=='formats'){for(const k of [...cur.cells.keys()]){const r=Math.floor(k/MAXC),c=k%MAXC;if(r>=g.r1&&r<=g.r2&&c>=g.c1&&c<=g.c2)clr(r,c,w);}return;}
      for(let r=g.r1;r<=g.r2;r++)for(let c=g.c1;c<=g.c2;c++)if(cur.cells.has(key(r,c)))clr(r,c,w);
    });
  }
  function clr(r,c,w){
    const cl=cellAt(cur,r,c);
    if(w==='all')putCell(cur,r,c,null);
    else if(w==='formats')putCell(cur,r,c,{...cl,st:undefined});
    else putCell(cur,r,c,cl.st?{v:null,st:cl.st}:null);
  }
  function selectAddr(t){
    t=t.trim().toUpperCase().replace(/\s/g,'');let m;
    if(m=/^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(t)){
      const a=parseAddr(m[1]+m[2]),b=m[3]?parseAddr(m[3]+m[4]):a;if(a&&b){select(a.r,a.c,b.r,b.c);return true;}}
    if(m=/^\$?([A-Z]+):\$?([A-Z]+)$/.exec(t)){const a=colIdx(m[1]),b=colIdx(m[2]);if(a<MAXC&&b<MAXC){select(0,a,MAXR-1,b);cur.sel.ar=0;cur.sel.ac=a;return true;}}
    if(m=/^(\d+):(\d+)$/.exec(t)){const a=+m[1]-1,b=+m[2]-1;if(a>=0&&b<MAXR){select(a,0,b,MAXC-1);return true;}}
    return false;
  }
  nb.addEventListener('keydown',e=>{e.stopPropagation();
    if(e.key==='Enter'){if(selectAddr(nb.value)){vp.focus();}else{toast('Referensi tidak valid.');}}
    else if(e.key==='Escape'){sync();vp.focus();}});
  nb.addEventListener('focus',()=>nb.select());
  nb.addEventListener('blur',()=>sync());

  const acts={
    menu(){menu.hidden=!menu.hidden;},
    new(){confirmNew();},open(){fin.click();},save:()=>saveXlsx(),savecsv:()=>saveCsv(),print:()=>printSheet(),home(){opts.onHome&&opts.onHome();},
    undo,redo,
    async paste(){
      if(clip){pasteMode('all');return;}
      try{pasteMode('all',undefined,undefined,null,await navigator.clipboard.readText());}catch{toast('Gunakan Ctrl+V untuk menempel');}},
    cut(){copyAct(true);},copy(){copyAct(false);},
    painter(){const s=cur.sel;painting={...stOf(s.ar,s.ac)};vp.style.cursor='copy';toast('Pilih sel tujuan untuk menerapkan format');},
    grow(){const n=SIZES.find(s=>s>SZ());if(n)applyStyle({sz:n},true);},
    shrink(){const n=[...SIZES].reverse().find(s=>s<SZ());if(n)applyStyle({sz:n},true);},
    bold(){togg('b');},italic(){togg('i');},underline(){togg('u',1);},strike(){togg('s');},
    fillnow(){applyStyle({fill:colorState.fill});},colornow(){applyStyle({color:colorState.color});},
    va(v){applyStyle({va:v});},ha(v){applyStyle({ha:v});},
    wrap(){togg('wrap');},
    indent(v){const cur0=stOf(cur.sel.ar,cur.sel.ac);const n=Math.max(0,Math.min(15,(cur0.ind||0)+ +v));applyStyle({ind:n||null,ha:n&&(!cur0.ha||cur0.ha==='general')?'left':cur0.ha});},
    merge(v){mergeSel(v);},
    nf(v){const st=stOf(cur.sel.ar,cur.sel.ac);applyStyle({fmt:st.fmt===FMT[v]?null:FMT[v]});},
    dec(v){decimals(+v);},
    cstyle(i){applyCellStyle(+i);},
    autosum(fn){autoSum(fn);},
    filldown(){fillDir('down');},clearc(){clearWhat('contents');},
    border(k){borderKind(k);},
    goto(){dlg('Pergi ke',`<label>Referensi: <input type="text" class="v" placeholder="mis. B12 atau A1:C10"></label>`,d=>{if(!selectAddr(d.querySelector('.v').value)){toast('Referensi tidak valid.');return false;}});},
    editok(){commitEdit(0,0);},editcancel(){cancelEdit();},
    fxdlg(){fnDialog();},
    soonv(){toast('Tampilan ini belum tersedia.');},
    zoomd(d){setZoom(Math.round(Z*100)+ +d);},
    shprev(){$('.shs').scrollLeft-=120;},shnext(){$('.shs').scrollLeft+=120;},
    addsheet:addSheet,
    fnext:findNext,frep:replaceOne,frall:replaceAll,fclose(){fd.hidden=true;},
    insdlg(){insertDialog();},deldlg(){deleteDialog();},
    fmtcells(tab){formatCells(tab);},
    rowh(){const r=rg().r1;askNumber('Tinggi Baris','Tinggi baris (pt):',+(cur.rh.size(r)*.75).toFixed(2),v=>tx(()=>{const g=rg();for(let i=g.r1;i<=Math.min(g.r2,g.r1+5000);i++){cur.rh.ov.set(i,Math.round(v/.75));cur.rh.hid.delete(i);(cur.rh.manual=cur.rh.manual||new Set()).add(i);}cur.rh.dirty=true;}));},
    colw(){const c=rg().c1;askNumber('Lebar Kolom','Lebar kolom:',+((cur.cw.size(c)-5)/7).toFixed(2),v=>tx(()=>{const g=rg();for(let i=g.c1;i<=Math.min(g.c2,g.c1+2000);i++){cur.cw.ov.set(i,Math.round(v*7+5));cur.cw.hid.delete(i);}cur.cw.dirty=true;}));},
    defw(){askNumber('Lebar Default','Lebar kolom standar:',+((cur.cw.def-5)/7).toFixed(2),v=>tx(()=>{cur.cw.def=Math.round(v*7+5);cur.cw.dirty=true;}));},
    fitr(){const g=rg();const rs=[];for(let r=g.r1;r<=Math.min(g.r2,g.r1+5000);r++)rs.push(r);autoFitRowsH(rs);},
    fitc(){const g=rg();const cs=[];for(let c=g.c1;c<=Math.min(g.c2,g.c1+500);c++)cs.push(c);autoFitCols(cs);},
    hider(){setHidden('r',true);},unhider(){setHidden('r',false);},hidec(){setHidden('c',true);},unhidec(){setHidden('c',false);},
    rensheet(){renameSheet(wb.active);},
    delsheet(){deleteSheet(wb.active);},
    inssheet(){addSheet();},
    mvl(){moveSheet(wb.active,-1);},mvr(){moveSheet(wb.active,1);}
  };
  const sels={
    font:v=>applyStyle({font:v==='Calibri'?null:v}),
    size:v=>applyStyle({sz:+v===11?null:+v},true),
    nfmt(v){if(v==='custom'){formatCells('number');return;}applyStyle({fmt:FMT[v]==='General'?null:FMT[v]});}
  };
  function applyCellStyle(i){
    const o=CSTY[i][2];
    styleSel(st=>{const n={...st};for(const k of CSTY_KEYS)delete n[k];
      for(const k in o)n[k]=o[k];return n;},true);
  }

  /* ===== sisip / hapus ===== */
  function insertDialog(){
    dlg('Sisipkan',`<label><input type="radio" name="m" value="right" checked> Geser sel ke kanan</label><label><input type="radio" name="m" value="down"> Geser sel ke bawah</label><label><input type="radio" name="m" value="row"> Seluruh baris</label><label><input type="radio" name="m" value="col"> Seluruh kolom</label>`,d=>{
      const v=d.querySelector('input:checked').value;const g=rg();
      if(v==='row')structOp('r',g.r1,g.r2-g.r1+1,false);else if(v==='col')structOp('c',g.c1,g.c2-g.c1+1,false);else shiftCells(v,true);});
  }
  function deleteDialog(){
    dlg('Hapus',`<label><input type="radio" name="m" value="left" checked> Geser sel ke kiri</label><label><input type="radio" name="m" value="up"> Geser sel ke atas</label><label><input type="radio" name="m" value="row"> Seluruh baris</label><label><input type="radio" name="m" value="col"> Seluruh kolom</label>`,d=>{
      const v=d.querySelector('input:checked').value;const g=rg();
      if(v==='row')structOp('r',g.r1,g.r2-g.r1+1,true);else if(v==='col')structOp('c',g.c1,g.c2-g.c1+1,true);else shiftCells(v==='left'?'right':'down',false);});
  }
  function fnDialog(){
    const names=[...FN_NAMES].sort();
    dlg('Sisipkan Fungsi',`<input type="text" class="q" placeholder="Cari fungsi"><div class="fcl">${names.map(n=>`<div data-n="${n}">${n}</div>`).join('')}</div>`,d=>{
      const on=d.querySelector('.fcl .on');if(!on)return false;
      const s=cur.sel;const t=(ed?ed.text:'=')+'';
      if(!ed)startEdit(s.ar,s.ac,'='+on.dataset.n+'(');else{const p=ed.bar?fxi:edEl;p.setRangeText(on.dataset.n+'(',p.selectionStart,p.selectionEnd,'end');onEdInput(p);}
    },'Sisipkan');
    const d=X.querySelector('.ov'),list=d.querySelector('.fcl');
    list.onclick=e=>{const t=e.target.closest('[data-n]');if(!t)return;list.querySelectorAll('.on').forEach(x=>x.classList.remove('on'));t.classList.add('on');};
    list.ondblclick=()=>d.querySelector('.ok').click();
    d.querySelector('.q').oninput=e=>{const q=e.target.value.toUpperCase();list.querySelectorAll('[data-n]').forEach(x=>x.hidden=!x.dataset.n.includes(q));};
  }

  /* ===== format tabel ===== */
  function applyTable(i){
    const g0=region();let g={...g0};
    if(g.r1===g.r2&&g.c1===g.c2){toast('Pilih sel di dalam data untuk dijadikan tabel.');return;}
    dlg('Format sebagai Tabel',`<div>Rentang data: <b>${addr(g.r1,g.c1)}:${addr(g.r2,g.c2)}</b></div><label><input type="checkbox" class="hd" checked> Tabel saya memiliki header</label>`,d=>{
      const hdr=d.querySelector('.hd').checked,ts=tableStyle(i);
      tx(()=>{
        for(let r=g.r1;r<=g.r2;r++)for(let c=g.c1;c<=g.c2;c++){
          const cl=cellAt(cur,r,c),base={...(cl&&cl.st||{})};
          for(const k of ['fill','color','b','bd'])delete base[k];
          const isH=hdr&&r===g.r1,bi=(r-g.r1-(hdr?1:0));
          let st={...base};
          if(isH)Object.assign(st,ts.head);
          else{const bc=ts.band[((bi%2)+2)%2];if(bc)st.fill=bc;if(ts.dark)st.color='#000000';}
          if(ts.bd&&!isH)st.bd={...ts.bd};
          if(ts.dark&&!isH)st.bd={t:['thin','#ffffff'],b:['thin','#ffffff'],l:['thin','#ffffff'],r:['thin','#ffffff']};
          putCell(cur,r,c,{...(cl||{v:null}),st:cleanSt(st)});
        }
        if(hdr&&!cur.filter)cur.filter={...g,crit:{}};
      });
    });
  }

  /* ===== pemformatan bersyarat ===== */
  const CFP=[['Isi merah muda dengan teks merah tua',{fill:'#ffc7ce',color:'#9c0006'}],['Isi kuning dengan teks kuning tua',{fill:'#ffeb9c',color:'#9c5700'}],['Isi hijau dengan teks hijau tua',{fill:'#c6efce',color:'#006100'}],['Isi merah muda saja',{fill:'#ffc7ce'}],['Teks merah tua saja',{color:'#9c0006'}],['Teks hijau tua tebal',{color:'#006100',b:1}]];
  const cfRange=()=>{const g=expandMerge(limitRange(rg()));return {r1:g.r1,c1:g.c1,r2:g.r2,c2:g.c2,sh:cur};};
  const addRule=rl=>tx(()=>{cur.cf=[...cur.cf,{...cfRange(),...rl}];});
  function cfDialog(title,label,type,op,twoVals){
    dlg(title,`<div>${label}</div><div class="r2"><span>Nilai:</span><span><input type="text" class="a" style="width:90px">${twoVals?' dan <input type="text" class="b" style="width:90px">':''}</span><span>Format:</span><select class="p">${CFP.map((p,i)=>`<option value="${i}">${p[0]}</option>`).join('')}</select></div>`,d=>{
      const a=d.querySelector('.a').value.trim(),b=twoVals?d.querySelector('.b').value.trim():undefined;
      if(a===''&&type!=='dup'&&type!=='uniq'){return false;}
      const cv=x=>x!==undefined&&x!==''&&!isErr(num(x))&&type==='cell'&&op!=='eq'?num(x):x;
      addRule({type,op,a:type==='text'?a:(!isNaN(+a.replace(',','.'))&&a!==''?+a.replace(',','.'):a),b:b!==undefined&&!isNaN(+b.replace(',','.'))&&b!==''?+b.replace(',','.'):b,fmt:CFP[+d.querySelector('.p').value][1]});});
  }
  function cfDup(type){
    dlg(type==='dup'?'Nilai Duplikat':'Nilai Unik',`<label>Format: <select class="p">${CFP.map((p,i)=>`<option value="${i}">${p[0]}</option>`).join('')}</select></label>`,d=>{addRule({type,fmt:CFP[+d.querySelector('.p').value][1]});});
  }
  function cfTop(type){
    const t=type==='top'?'Teratas':'Terbawah';
    dlg('10 '+t,`<label>Format sel yang berada di ${t.toLowerCase()} <input type="number" class="n" value="10" min="1" style="width:60px"> <label><input type="checkbox" class="pc"> persen</label></label><label>Format: <select class="p">${CFP.map((p,i)=>`<option value="${i}">${p[0]}</option>`).join('')}</select></label>`,d=>{addRule({type,n:Math.max(1,num0(d,'.n')),pct:d.querySelector('.pc').checked,fmt:CFP[+d.querySelector('.p').value][1]});});
  }
  function cfClear(all){
    const g=rg();tx(()=>{cur.cf=all?[]:cur.cf.filter(r=>r.r2<g.r1||r.r1>g.r2||r.c2<g.c1||r.c1>g.c2);});
  }
  const cfDesc=r=>{const rg2=addr(r.r1,r.c1)+':'+addr(r.r2,r.c2);
    const T={cell:{gt:'Nilai sel > ',lt:'Nilai sel < ',ge:'Nilai sel ≥ ',le:'Nilai sel ≤ ',eq:'Nilai sel = ',btw:'Nilai sel antara '}[r.op]+(r.a!==undefined?r.a:'')+(r.op==='btw'?' dan '+r.b:''),text:'Teks berisi "'+r.a+'"',dup:'Nilai duplikat',uniq:'Nilai unik',
      top:'Teratas '+r.n+(r.pct?'%':''),bottom:'Terbawah '+r.n+(r.pct?'%':''),above:'Di atas rata-rata',below:'Di bawah rata-rata',bar:'Bilah data',scale:'Skala warna',icons:'Set ikon'};
    return `${T[r.type]||r.type} — ${rg2}`;};
  function cfManage(){
    const body=()=>cur.cf.length?cur.cf.map((r,i)=>`<div class="sli" style="display:flex;justify-content:space-between;gap:12px;border-bottom:1px solid #eee;padding:3px 0"><span>${esc(cfDesc(r))}</span><button class="bt" data-d="${i}">✕</button></div>`).join(''):'<i>Belum ada aturan</i>';
    const d=dlg('Pengelola Aturan Pemformatan Bersyarat',`<div class="cl" style="min-width:300px">${body()}</div>`,null,'Selesai',true);
    d.querySelector('.cl').onclick=e=>{const i=e.target.dataset.d;if(i===undefined)return;tx(()=>{cur.cf=cur.cf.filter((_,k)=>k!==+i);});d.querySelector('.cl').innerHTML=body();};
  }

  /* ===== dialog Format Sel ===== */
  function formatCells(tab){
    const s=cur.sel,st0=stOf(s.ar,s.ac),v0=cellVal(cur,s.ar,s.ac);
    const S={fmt:st0.fmt||'General',ha:st0.ha||'general',va:st0.va||'bottom',ind:st0.ind||0,wrap:!!st0.wrap,merge:!!mergeAt(s.ar,s.ac),rot:st0.rot||0,
      font:st0.font||'Calibri',sz:st0.sz||11,b:!!st0.b,i:!!st0.i,u:st0.u||0,s:!!st0.s,color:st0.color||'#000000',fill:st0.fill||null,bd:{}};
    const bdSet={t:null,b:null,l:null,r:null,h:undefined,v:undefined};let bdStyleSel='thin',bdColorSel='#000000';
    const TABN=[['number','Angka'],['align','Perataan'],['font','Font'],['border','Garis Tepi'],['fill','Isi']];
    const CATS=[['General','Umum'],['Number','Angka'],['Currency','Mata uang'],['Accounting','Akuntansi'],['Date','Tanggal'],['Time','Waktu'],['Percent','Persentase'],['Fraction','Pecahan'],['Sci','Ilmiah'],['Text','Teks'],['Custom','Kustom']];
    const DATES=['dd/mm/yyyy','d mmmm yyyy','dddd, d mmmm yyyy','d-mmm-yy','mmm-yy','mmmm yyyy','yyyy-mm-dd'];
    const TIMES=['hh:mm','hh:mm:ss','hh:mm AM/PM','dd/mm/yyyy hh:mm'];
    let cat='Custom';const k0=fmtKey(S.fmt);
    if(k0)cat={ShortDate:'Date',LongDate:'Date',Time:'Time',Comma:'Number'}[k0]||k0;
    else if(isDateCode(S.fmt))cat='Date';
    const catSt={dec:2,thou:false,sym:'"Rp"'};
    const m0=/0\.(0+)/.exec(S.fmt);if(m0)catSt.dec=m0[1].length;else if(/0/.test(S.fmt)&&!/\./.test(S.fmt))catSt.dec=0;
    catSt.thou=/#,##0/.test(S.fmt);
    const d=dlg('Format Sel',`<div class="dtabs">${TABN.map(t=>`<button data-t="${t[0]}">${t[1]}</button>`).join('')}</div><div class="dp"></div>`,()=>{
      const patch={fmt:S.fmt==='General'?null:S.fmt,ha:S.ha==='general'?null:S.ha,va:S.va==='bottom'?null:S.va,ind:S.ind||null,wrap:S.wrap||null,rot:S.rot||null,
        font:S.font==='Calibri'?null:S.font,sz:S.sz===11?null:S.sz,b:S.b||null,i:S.i||null,u:S.u||null,s:S.s||null,color:S.color==='#000000'?null:S.color,fill:S.fill};
      const hasB=Object.values(bdSet).some(x=>x!==undefined&&x!==null)||bdTouched;
      styleSel(st=>({...st,...patch}),true);
      if(hasB||bdTouched)applyBorderSpec(Object.fromEntries(Object.entries(bdSet).filter(([,v])=>v!==undefined)));
      const g=expandMerge(rg()),isM=cur.merges.some(m=>m.r1===g.r1&&m.r2===g.r2&&m.c1===g.c1&&m.c2===g.c2);
      if(S.merge&&!isM)mergeSel('cells');else if(!S.merge&&isM)mergeSel('un');
    },'OK');
    let bdTouched=false;
    const P=d.querySelector('.dp');
    const sample=()=>{const el=P.querySelector('.smp');if(el)el.textContent=v0===null?'(vide)':fmtValue(v0,S.fmt).t;};
    function buildCode(){
      const dec=catSt.dec,z=dec?'.'+'0'.repeat(dec):'';
      switch(cat){
        case 'General':return 'General';
        case 'Number':return (catSt.thou?'#,##0':'0')+z;
        case 'Currency':return (catSt.sym?catSt.sym:'')+'#,##0'+z;
        case 'Accounting':return '_('+catSt.sym+'* #,##0'+z+'_);_('+catSt.sym+'* (#,##0'+z+');_('+catSt.sym+'* "-"_);_(@_)';
        case 'Percent':return '0'+z+'%';
        case 'Fraction':return '# ?/?';case 'Sci':return '0'+z+'E+00';case 'Text':return '@';
      }
      return S.fmt;
    }
    function show(t){
      d.querySelectorAll('.dtabs button').forEach(b=>b.classList.toggle('on',b.dataset.t===t));
      if(t==='number'){
        P.innerHTML=`<div style="display:flex;gap:10px;flex-wrap:wrap"><div><div>Kategori:</div><div class="fcl" style="width:150px">${CATS.map(c=>`<div data-c="${c[0]}" class="${c[0]===cat?'on':''}">${c[1]}</div>`).join('')}</div></div><div style="min-width:200px;display:flex;flex-direction:column;gap:8px"><div>Contoh:</div><div class="prev smp" style="min-width:200px"></div><div class="opt"></div></div></div>`;
        const opt=P.querySelector('.opt');
        const drawOpt=()=>{
          let h='';
          if(['Number','Currency','Accounting','Percent','Sci'].includes(cat))h+=`<label>Tempat desimal: <input type="number" class="dc" min="0" max="15" value="${catSt.dec}" style="width:56px"></label>`;
          if(cat==='Number')h+=`<label><input type="checkbox" class="th" ${catSt.thou?'checked':''}> Gunakan pemisah ribuan</label>`;
          if(cat==='Currency'||cat==='Accounting')h+=`<label>Simbol: <select class="sy">${[['"Rp"','Rp'],['$','$'],['"€"','€'],['"£"','£'],['"¥"','¥'],['','(tanpa)']].map(x=>`<option value="${esc(x[0])}" ${x[0]===catSt.sym?'selected':''}>${x[1]}</option>`).join('')}</select></label>`;
          if(cat==='Date')h+=`<div>Tipe:</div><div class="fcl tp">${DATES.map(x=>`<div data-f="${x}" class="${x===S.fmt?'on':''}">${fmtDate(v0===null||typeof v0!=='number'?45000:v0,x)}</div>`).join('')}</div>`;
          if(cat==='Time')h+=`<div>Tipe:</div><div class="fcl tp">${TIMES.map(x=>`<div data-f="${x}" class="${x===S.fmt?'on':''}">${fmtDate(typeof v0==='number'?v0:.75,x)}</div>`).join('')}</div>`;
          if(cat==='Custom')h+=`<label>Kode format:<input type="text" class="cu" value="${esc(S.fmt)}" style="width:100%"></label><div style="font-size:11px;color:#666">Contoh: #,##0.00 · 0% · dd/mm/yyyy · "Rp"#,##0</div>`;
          if(cat==='General')h+='<div style="font-size:11px;color:#666">Sel berformat Umum tidak memiliki format angka tertentu.</div>';
          opt.innerHTML=h;
          const q=s=>opt.querySelector(s);
          if(q('.dc'))q('.dc').oninput=e=>{catSt.dec=Math.max(0,Math.min(15,+e.target.value||0));S.fmt=buildCode();sample();};
          if(q('.th'))q('.th').onchange=e=>{catSt.thou=e.target.checked;S.fmt=buildCode();sample();};
          if(q('.sy'))q('.sy').onchange=e=>{catSt.sym=e.target.value;S.fmt=buildCode();sample();};
          if(q('.cu'))q('.cu').oninput=e=>{S.fmt=e.target.value||'General';sample();};
          const tp=q('.tp');if(tp)tp.onclick=e=>{const t=e.target.closest('[data-f]');if(!t)return;S.fmt=t.dataset.f;tp.querySelectorAll('.on').forEach(x=>x.classList.remove('on'));t.classList.add('on');sample();};
        };
        P.querySelector('.fcl').onclick=e=>{const t=e.target.closest('[data-c]');if(!t)return;cat=t.dataset.c;P.querySelectorAll('.fcl>div[data-c]').forEach(x=>x.classList.toggle('on',x===t));
          if(cat==='Date'&&!isDateCode(S.fmt))S.fmt=DATES[0];else if(cat==='Time'&&!isDateCode(S.fmt))S.fmt=TIMES[0];else if(cat!=='Custom')S.fmt=buildCode();drawOpt();sample();};
        drawOpt();sample();
      }else if(t==='align'){
        P.innerHTML=`<div class="r2"><span>Horizontal:</span><select class="ha">${[['general','Umum'],['left','Kiri (Indentasi)'],['center','Tengah'],['right','Kanan (Indentasi)'],['fill','Isi'],['justify','Rata kiri-kanan']].map(x=>`<option value="${x[0]}" ${x[0]===S.ha?'selected':''}>${x[1]}</option>`).join('')}</select>
          <span>Indentasi:</span><input type="number" class="in" min="0" max="15" value="${S.ind}" style="width:60px">
          <span>Vertikal:</span><select class="va">${[['top','Atas'],['middle','Tengah'],['bottom','Bawah']].map(x=>`<option value="${x[0]}" ${x[0]===S.va?'selected':''}>${x[1]}</option>`).join('')}</select>
          <span>Orientasi (derajat):</span><input type="number" class="ro" min="-90" max="90" value="${S.rot===255?0:S.rot}" style="width:60px"></div>
          <label><input type="checkbox" class="wr" ${S.wrap?'checked':''}> Bungkus teks</label><label><input type="checkbox" class="mg" ${S.merge?'checked':''}> Gabungkan sel</label>`;
        P.querySelector('.ha').onchange=e=>S.ha=e.target.value;P.querySelector('.va').onchange=e=>S.va=e.target.value;
        P.querySelector('.in').oninput=e=>S.ind=Math.max(0,+e.target.value||0);P.querySelector('.ro').oninput=e=>S.rot=Math.max(-90,Math.min(90,+e.target.value||0));
        P.querySelector('.wr').onchange=e=>S.wrap=e.target.checked;P.querySelector('.mg').onchange=e=>S.merge=e.target.checked;
      }else if(t==='font'){
        P.innerHTML=`<div class="r2"><span>Font:</span><select class="fo">${FONTS.map(f=>`<option ${f===S.font?'selected':''}>${f}</option>`).join('')}</select>
          <span>Gaya:</span><select class="gs">${['Reguler','Miring','Tebal','Tebal Miring'].map((x,i)=>`<option value="${i}" ${(S.b?2:0)+(S.i?1:0)===i?'selected':''}>${x}</option>`).join('')}</select>
          <span>Ukuran:</span><input type="number" class="sz" min="1" max="409" value="${S.sz}" style="width:70px">
          <span>Garis bawah:</span><select class="ul">${[['0','Tidak ada'],['1','Tunggal'],['2','Ganda']].map(x=>`<option value="${x[0]}" ${+x[0]===S.u?'selected':''}>${x[1]}</option>`).join('')}</select>
          <span>Warna:</span><input type="color" class="co" value="${S.color}"></div><label><input type="checkbox" class="st" ${S.s?'checked':''}> Coret</label><div class="prev" style="font-size:14px">AaBbCcYyZz</div>`;
        const pv=()=>{const e=P.querySelector('.prev');e.style.cssText=`font-family:${S.font};font-size:${Math.min(S.sz,28)}pt;font-weight:${S.b?700:400};font-style:${S.i?'italic':'normal'};color:${S.color};text-decoration:${S.u?'underline':''} ${S.s?'line-through':''}`;};
        P.querySelector('.fo').onchange=e=>{S.font=e.target.value;pv();};
        P.querySelector('.gs').onchange=e=>{const v=+e.target.value;S.b=v>=2;S.i=v%2===1;pv();};
        P.querySelector('.sz').oninput=e=>{S.sz=+e.target.value||11;pv();};P.querySelector('.ul').onchange=e=>{S.u=+e.target.value;pv();};
        P.querySelector('.co').oninput=e=>{S.color=e.target.value;pv();};P.querySelector('.st').onchange=e=>{S.s=e.target.checked;pv();};pv();
      }else if(t==='border'){
        const sides=[['t','Atas'],['b','Bawah'],['l','Kiri'],['r','Kanan'],['h','Dalam horizontal'],['v','Dalam vertikal']];
        P.innerHTML=`<div class="r2"><span>Gaya:</span><select class="bs">${Object.keys(BTYPES).filter(x=>x!=='hair').map(x=>`<option ${x===bdStyleSel?'selected':''}>${x}</option>`).join('')}</select><span>Warna:</span><input type="color" class="bc" value="${bdColorSel}"></div>
          <div class="bdl">${sides.map(x=>`<label><input type="checkbox" data-s="${x[0]}"> ${x[1]}</label>`).join('')}</div>
          <div style="display:flex;gap:6px"><button class="bt" data-p="none">Tidak ada</button><button class="bt" data-p="outer">Garis luar</button><button class="bt" data-p="all">Semua</button></div>`;
        const bx=[...P.querySelectorAll('[data-s]')];
        bx.forEach(b=>{b.checked=bdSet[b.dataset.s]!==undefined&&bdSet[b.dataset.s]!==null;
          b.onchange=()=>{bdTouched=true;const k=b.dataset.s;bdSet[k]=b.checked?[bdStyleSel,bdColorSel]:(['h','v'].includes(k)?null:null);};});
        P.querySelector('.bs').onchange=e=>{bdStyleSel=e.target.value;bx.forEach(b=>{if(b.checked)bdSet[b.dataset.s]=[bdStyleSel,bdColorSel];});};
        P.querySelector('.bc').oninput=e=>{bdColorSel=e.target.value;bx.forEach(b=>{if(b.checked)bdSet[b.dataset.s]=[bdStyleSel,bdColorSel];});};
        P.querySelectorAll('[data-p]').forEach(b=>b.onclick=()=>{bdTouched=true;const p=b.dataset.p;
          bx.forEach(x=>{const k=x.dataset.s;const on=p==='all'||(p==='outer'&&'tblr'.includes(k));x.checked=on;bdSet[k]=on?[bdStyleSel,bdColorSel]:null;});});
      }else{
        P.innerHTML=`<div>Warna latar:</div><div class="pal">${PAL.map(r=>`<div class="pr">${r.map(c=>`<button style="background:${c}" data-c="${c}"></button>`).join('')}</div>`).join('')}<div class="pr" style="margin-top:6px">${PALSTD.map(c=>`<button style="background:${c}" data-c="${c}"></button>`).join('')}</div></div><div style="display:flex;gap:10px;align-items:center"><button class="bt" data-n="1">Tanpa Warna</button><label>Lainnya: <input type="color" class="fc" value="${S.fill||'#ffffff'}"></label></div><div class="prev fp" style="background:${S.fill||'#fff'}">Contoh</div>`;
        const pv=()=>P.querySelector('.fp').style.background=S.fill||'#fff';
        P.onclick=e=>{const b=e.target.closest('button');if(!b||!P.querySelector('.fp'))return;if(b.dataset.c){S.fill=b.dataset.c;pv();}else if(b.dataset.n){S.fill=null;pv();}};
        P.querySelector('.fc').oninput=e=>{S.fill=e.target.value;pv();};
      }
    }
    d.querySelector('.dtabs').onclick=e=>{const b=e.target.closest('button');if(b)show(b.dataset.t);};
    show(TABN.some(x=>x[0]===tab)?tab:'number');
  }

  /* ===== menu isi ===== */
  const I=(l,a,i,extra)=>({l,a,i,...extra});
  const BI=(l,k,i)=>I(l,()=>borderKind(k),i);
  const rowOrCol=()=>{const g=rg();return {r:g.r2-g.r1+1,c:g.c2-g.c1+1};};
  const menus={
    paste:()=>[I('Tempel',()=>pasteMode('all'),'📋'),I('Nilai',()=>pasteMode('values'),'123'),I('Rumus',()=>pasteMode('formulas'),'ƒ'),I('Pemformatan',()=>pasteMode('formats'),'🖌'),I('Transpose',()=>pasteMode('transpose'),'⇄')].map(x=>({...x,dis:!clip&&x.l!=='Tempel'})),
    ul:()=>[I('Garis Bawah',()=>applyStyle({u:1}),'U'),I('Garis Bawah Ganda',()=>applyStyle({u:2}),'U̳'),I('Tanpa Garis Bawah',()=>applyStyle({u:null}),'')],
    border:()=>[{h:'Tepi'},BI('Tepi Bawah','bottom','▁'),BI('Tepi Atas','top','▔'),BI('Tepi Kiri','left','▏'),BI('Tepi Kanan','right','▕'),{sep:1},BI('Tanpa Tepi','none','☐'),BI('Semua Tepi','all','▦'),BI('Tepi Luar','outer','▢'),BI('Tepi Luar Tebal','thickOuter','■'),{sep:1},
      BI('Tepi Bawah Ganda','dblBottom','═'),BI('Tepi Bawah Tebal','thickBottom','▃'),BI('Tepi Atas dan Bawah','topBottom','☰'),BI('Atas dan Bawah Tebal','topThickBottom','☰'),BI('Atas dan Bawah Ganda','topDblBottom','☰'),{sep:1},
      {l:'Warna Garis',i:'🎨',sub:()=>palette(c=>{if(c)bdColor=c;},{})},
      {l:'Gaya Garis',i:'─',sub:Object.keys(BTYPES).filter(x=>x!=='hair').map(k=>({l:k==='thin'?'Tipis':k==='medium'?'Sedang':k==='thick'?'Tebal':k==='dashed'?'Putus-putus':k==='dotted'?'Titik-titik':'Ganda',ck:bdStyle===k,a:()=>{bdStyle=k;}}))},
      {sep:1},I('Garis Tepi Lainnya…',()=>formatCells('border'),'▤')],
    fill:()=>palette(c=>{if(c===null)applyStyle({fill:null});else{colorState.fill=c;setSw('fill',c);applyStyle({fill:c});}},{none:'∅ Tanpa Isi'}),
    color:()=>palette(c=>{if(c===null)applyStyle({color:null});else{colorState.color=c;setSw('color',c);applyStyle({color:c});}},{none:'A Otomatis'}),
    orient:()=>[['Sudut Berlawanan Jarum Jam',45,'↗'],['Sudut Searah Jarum Jam',-45,'↘'],['Teks Vertikal',255,'⇣'],['Putar Teks ke Atas',90,'↑'],['Putar Teks ke Bawah',-90,'↓'],['Tanpa Rotasi',0,'→']].map(([l,v,i])=>I(l,()=>applyStyle({rot:v||null}),i)),
    merge:()=>[I('Gabung & Tengah',()=>mergeSel('center'),'⇔'),I('Gabung Menyilang',()=>mergeSel('across'),'⇔'),I('Gabungkan Sel',()=>mergeSel('cells'),'⇔'),I('Pisahkan Sel',()=>mergeSel('un'),'⇱')],
    acct:()=>[['Rp Indonesia','"Rp"#,##0'],['$ Inggris (AS)','$#,##0.00'],['€ Euro','"€"#,##0.00'],['£ Inggris (Inggris Raya)','"£"#,##0.00'],['¥ Yen','"¥"#,##0']].map(([l,f])=>I(l,()=>applyStyle({fmt:f}),'')),
    csty:()=>[{html:`<div class="sgal">${CSTY.map((s,i)=>{const o=s[2];return `<button data-i="${i}" style="${o.fill?'background:'+o.fill+';':''}${o.color?'color:'+o.color+';':''}${o.b?'font-weight:700;':''}${o.i?'font-style:italic;':''}${o.sz?'font-size:'+Math.min(o.sz,15)+'px;':''}${o.bd&&o.bd.b?'border-bottom:2px solid '+o.bd.b[1]+';':''}">${s[1]}</button>`;}).join('')}</div>`,
      on:d=>d.addEventListener('click',e=>{const b=e.target.closest('button');if(b){closePop();applyCellStyle(+b.dataset.i);}})}],
    tbl:()=>[{html:`<div class="tgal">${TBL.map((t,i)=>{const s=tableStyle(i),c=t[1];return `<button data-i="${i}" title="${t[0]}"><i style="background:${s.head.fill||'#fff'};border-bottom:2px solid ${c}"></i><i style="background:${s.band[0]||'#fff'}"></i><i style="background:${s.band[1]||'#fff'}"></i><i style="background:${s.band[0]||'#fff'}"></i></button>`;}).join('')}</div>`,
      on:d=>d.addEventListener('click',e=>{const b=e.target.closest('button');if(b){closePop();applyTable(+b.dataset.i);}})}],
    cf:()=>[
      {l:'Aturan Sorot Sel',i:'▤',sub:[I('Lebih Besar Dari…',()=>cfDialog('Lebih Besar Dari','Format sel yang LEBIH BESAR DARI:','cell','gt')),I('Lebih Kecil Dari…',()=>cfDialog('Lebih Kecil Dari','Format sel yang LEBIH KECIL DARI:','cell','lt')),
        I('Antara…',()=>cfDialog('Antara','Format sel yang berada ANTARA:','cell','btw',true)),I('Sama Dengan…',()=>cfDialog('Sama Dengan','Format sel yang SAMA DENGAN:','cell','eq')),I('Teks yang Berisi…',()=>cfDialog('Teks yang Berisi','Format sel yang berisi teks:','text')),I('Nilai Duplikat…',()=>cfDup('dup'))]},
      {l:'Aturan Atas/Bawah',i:'⇵',sub:[I('10 Item Teratas…',()=>cfTop('top')),I('10 Item Terbawah…',()=>cfTop('bottom')),I('Di Atas Rata-rata',()=>addRule({type:'above',fmt:CFP[0][1]})),I('Di Bawah Rata-rata',()=>addRule({type:'below',fmt:CFP[2][1]}))]},
      {l:'Bilah Data',i:'▥',sub:[['Biru','#638ec6'],['Hijau','#63be7b'],['Merah','#ff555a'],['Oranye','#ffb628'],['Ungu','#a569bd']].map(([l,c])=>I(l,()=>addRule({type:'bar',color:c}),'▬'))},
      {l:'Skala Warna',i:'🎨',sub:[['Hijau - Kuning - Merah',['#63be7b','#ffeb84','#f8696b']],['Merah - Kuning - Hijau',['#f8696b','#ffeb84','#63be7b']],['Biru - Putih - Merah',['#5a8ac6','#fcfcff','#f8696b']],['Putih - Merah',['#fcfcff','#f8696b']],['Putih - Hijau',['#fcfcff','#63be7b']],['Kuning - Hijau',['#ffeb84','#63be7b']]].map(([l,cs])=>I(l,()=>addRule({type:'scale',colors:cs}),'▦'))},
      {l:'Set Ikon',i:'▲',sub:[['Panah','arrows'],['Lampu Lalu Lintas','lights'],['Simbol','symbols']].map(([l,s])=>I(l,()=>addRule({type:'icons',set:s}),ICONS[s][0][0]))},
      {sep:1},{l:'Hapus Aturan',i:'✖',sub:[I('Dari Sel yang Dipilih',()=>cfClear(false)),I('Dari Seluruh Lembar',()=>cfClear(true))]},I('Kelola Aturan…',cfManage,'⚙')],
    insert:()=>[I('Sisipkan Sel…',insertDialog,'⊞'),I('Sisipkan Baris Lembar',()=>{const g=rg();structOp('r',g.r1,g.r2-g.r1+1,false);},'☰'),I('Sisipkan Kolom Lembar',()=>{const g=rg();structOp('c',g.c1,g.c2-g.c1+1,false);},'▥'),I('Sisipkan Lembar',addSheet,'▭')],
    delete:()=>[I('Hapus Sel…',deleteDialog,'⊟'),I('Hapus Baris Lembar',()=>{const g=rg();structOp('r',g.r1,g.r2-g.r1+1,true);},'☰'),I('Hapus Kolom Lembar',()=>{const g=rg();structOp('c',g.c1,g.c2-g.c1+1,true);},'▥'),I('Hapus Lembar',()=>deleteSheet(wb.active),'▭')],
    format:()=>[{h:'Ukuran Sel'},I('Tinggi Baris…',acts.rowh,'↕'),I('AutoFit Tinggi Baris',acts.fitr,'↕'),I('Lebar Kolom…',acts.colw,'↔'),I('AutoFit Lebar Kolom',acts.fitc,'↔'),I('Lebar Default…',acts.defw,'↔'),{sep:1},
      {l:'Sembunyikan & Tampilkan',i:'👁',sub:[I('Sembunyikan Baris',acts.hider),I('Sembunyikan Kolom',acts.hidec),I('Tampilkan Baris',acts.unhider),I('Tampilkan Kolom',acts.unhidec)]},{sep:1},
      {h:'Atur Lembar'},I('Ganti Nama Lembar',acts.rensheet,'✎'),I('Pindahkan Lembar ke Kiri',acts.mvl,'◀'),I('Pindahkan Lembar ke Kanan',acts.mvr,'▶'),
      {l:'Warna Tab',i:'🎨',sub:()=>palette(c=>{tx(()=>{cur.tab=c;},true);drawTabs();},{none:'Tanpa Warna'})},{sep:1},I('Format Sel…',()=>formatCells('number'),'▤')],
    sum:()=>[['Jumlah','SUM'],['Rata-rata','AVERAGE'],['Hitung Angka','COUNT'],['Maks','MAX'],['Min','MIN']].map(([l,f])=>I(l,()=>autoSum(f),'Σ')),
    fillm:()=>[I('Bawah',()=>fillDir('down'),'⬇'),I('Kanan',()=>fillDir('right'),'➡'),I('Atas',()=>fillDir('up'),'⬆'),I('Kiri',()=>fillDir('left'),'⬅'),{sep:1},I('Seri…',seriesDialog,'123')],
    clear:()=>[I('Hapus Semua',()=>clearWhat('all'),'⌧'),I('Hapus Format',()=>clearWhat('formats'),'🖌'),I('Hapus Isi',()=>clearWhat('contents'),'✂')],
    sort:()=>[I('Urutkan A ke Z',()=>quickSort(true),'↓'),I('Urutkan Z ke A',()=>quickSort(false),'↑'),I('Urutan Kustom…',sortDialog,'⇅'),{sep:1},I('Filter',toggleFilter,'⧩',{ck:!!cur.filter}),I('Bersihkan Filter',()=>{if(cur.filter)tx(()=>{cur.filter.crit={};applyFilter(true);});},'✖',{dis:!cur.filter})],
    find:()=>[I('Cari…',showFind,'🔍'),I('Ganti…',showFind,'ab'),I('Pergi ke…',acts.goto,'→'),{sep:1},I('Pilih Semua',()=>select(0,0,MAXR-1,MAXC-1,true),'⬚')],
    ctx:()=>[I('Potong',()=>copyAct(true),'✂'),I('Salin',()=>copyAct(false),'⎘'),I('Tempel',()=>acts.paste(),'📋'),{l:'Tempel Spesial',i:'',sub:menus.paste()},{sep:1},I('Sisipkan…',insertDialog,'⊞'),I('Hapus…',deleteDialog,'⊟'),I('Hapus Isi',()=>clearWhat('contents'),'⌫'),{sep:1},
      I('Urutkan A ke Z',()=>quickSort(true),'↓'),I('Lebar Kolom…',acts.colw,'↔'),I('Tinggi Baris…',acts.rowh,'↕'),I('Sembunyikan Baris',acts.hider),I('Sembunyikan Kolom',acts.hidec),{sep:1},I('Format Sel…',()=>formatCells('number'),'▤')],
    sheetctx:()=>[I('Sisipkan Lembar',addSheet,'▭'),I('Hapus',()=>deleteSheet(wb.active),'✖'),I('Ganti Nama',acts.rensheet,'✎'),I('Pindahkan ke Kiri',acts.mvl,'◀'),I('Pindahkan ke Kanan',acts.mvr,'▶'),{l:'Warna Tab',i:'🎨',sub:()=>palette(c=>{tx(()=>{cur.tab=c;},true);drawTabs();},{none:'Tanpa Warna'})}]
  };
  function sortDialog(){
    const g=region();if(g.r1===g.r2&&g.c1===g.c2){toast('Pilih sel di dalam data yang akan diurutkan.');return;}
    const hdr=guessHeader(g);
    const colLabel=(c,h)=>h?(String(cellVal(cur,g.r1,c)??'')||('Kolom '+colName(c))):'Kolom '+colName(c);
    const d=dlg('Urutkan',`<label><input type="checkbox" class="hd" ${hdr?'checked':''}> Data saya memiliki header</label><div class="r2"><span>Urutkan menurut:</span><select class="cl"></select><span>Urutan:</span><select class="od"><option value="1">A ke Z / Terkecil ke Terbesar</option><option value="0">Z ke A / Terbesar ke Terkecil</option></select></div>`,d=>{
      sortRange(g,+d.querySelector('.cl').value,d.querySelector('.od').value==='1',d.querySelector('.hd').checked);});
    const fill=()=>{const h=d.querySelector('.hd').checked;d.querySelector('.cl').innerHTML=Array.from({length:g.c2-g.c1+1},(_,i)=>`<option value="${g.c1+i}">${esc(colLabel(g.c1+i,h))}</option>`).join('');};
    d.querySelector('.hd').onchange=fill;fill();
  }
  function seriesDialog(){
    dlg('Seri',`<div class="r2"><span>Seri dalam:</span><select class="dr"><option value="c">Kolom</option><option value="r">Baris</option></select><span>Tipe:</span><select class="ty"><option value="lin">Linear</option><option value="gro">Pertumbuhan</option></select><span>Nilai langkah:</span><input type="number" class="st" value="1" step="any"><span>Nilai akhir:</span><input type="text" class="en"></div>`,d=>{
      const g=rg(),col=d.querySelector('.dr').value==='c',step=num0(d,'.st'),gro=d.querySelector('.ty').value==='gro',en=d.querySelector('.en').value.trim()===''?null:+d.querySelector('.en').value;
      tx(()=>{
        const lines=col?g.c2-g.c1+1:g.r2-g.r1+1,len=col?g.r2-g.r1+1:g.c2-g.c1+1;
        for(let l=0;l<lines;l++){
          const at=i=>col?[g.r1+i,g.c1+l]:[g.r1+l,g.c1+i];
          const [r0,c0]=at(0);let v=cellVal(cur,r0,c0);if(typeof v!=='number')v=0;
          const stOld=stOf(r0,c0);
          const n=en!==null&&len===1?1000:len;
          for(let i=0;i<n;i++){const val=gro?v*Math.pow(step,i):v+step*i;if(en!==null&&(step>=0?val>en:val<en))break;const [r,c]=at(i);if(r>=MAXR||c>=MAXC)break;const o=cellAt(cur,r,c);putCell(cur,r,c,{v:val,st:o&&o.st||stOld});}
        }
      });
    });
  }

  /* ===== berkas ===== */
  const hex6=c=>c?c.replace('#','').toUpperCase():null;
  const argb=c=>'FF'+hex6(c);
  const THEME=['FFFFFF','000000','E7E6E6','44546A','4472C4','ED7D31','A5A5A5','FFC000','5B9BD5','70AD47'];
  function colorOf(c){
    if(!c)return null;
    if(c.argb){const a=c.argb.length===8?c.argb.slice(2):c.argb;return '#'+a.toLowerCase();}
    if(c.theme!==undefined){let h=THEME[c.theme]||'000000';let n=parseInt(h,16),R=n>>16,Gc=(n>>8)&255,Bc=n&255;const t=c.tint||0;
      const f=x=>Math.round(t>0?x+(255-x)*t:x*(1+t));return '#'+[f(R),f(Gc),f(Bc)].map(x=>Math.max(0,Math.min(255,x)).toString(16).padStart(2,'0')).join('');}
    return null;
  }
  const SERIALNUM=v=>v instanceof Date?dateToSerial(new Date(v.getUTCFullYear(),v.getUTCMonth(),v.getUTCDate(),v.getUTCHours(),v.getUTCMinutes(),v.getUTCSeconds())):v;
  async function importXlsx(buf){
    await lib('exceljs',()=>window.ExcelJS);
    const w=new ExcelJS.Workbook();await w.xlsx.load(buf);
    const sheets=[];
    for(const ws of w.worksheets){
      const s=newSheet(ws.name);
      const dc=ws.properties&&ws.properties.defaultColWidth;if(dc)s.cw.def=Math.round(dc*7+5);
      const cc=ws.columnCount||0;
      for(let c=1;c<=cc;c++){const col=ws.getColumn(c);if(col.width)s.cw.ov.set(c-1,Math.round(col.width*7+5));if(col.hidden)s.cw.hid.add(c-1);}
      ws.eachRow({includeEmpty:true},(row,rn)=>{
        if(row.height)s.rh.ov.set(rn-1,Math.round(row.height*4/3)),(s.rh.manual=s.rh.manual||new Set()).add(rn-1);
        if(row.hidden)s.rh.hid.add(rn-1);
        row.eachCell({includeEmpty:true},(cell,cn)=>{
          if(cell.type===ExcelJS.ValueType.Merge)return;
          const o={};let val=cell.value;
          if(val&&typeof val==='object'&&!(val instanceof Date)){
            if(val.richText)val=val.richText.map(x=>x.text).join('');
            else if('formula' in val||'sharedFormula' in val){let f=cell.formula;
              if(f){o.f=f.replace(/^=/,'');}else{val=val.result!==undefined?val.result:null;if(val&&val.error)val=val.error;}
              if(o.f!==undefined)val=null;}
            else if(val.hyperlink)val=val.text||val.hyperlink;
            else if(val.error){val=val.error;}
          }
          if(o.f===undefined){if(val===null||val===undefined||val==='')o.v=null;else o.v=SERIALNUM(val);}
          const st={};
          const f=cell.font;
          if(f){if(f.name&&f.name!=='Calibri')st.font=f.name;if(f.size&&f.size!==11)st.sz=f.size;if(f.bold)st.b=1;if(f.italic)st.i=1;if(f.underline)st.u=f.underline==='double'?2:1;if(f.strike)st.s=1;const fc=colorOf(f.color);if(fc&&fc!=='#000000')st.color=fc;}
          const fl=cell.fill;if(fl&&fl.type==='pattern'&&fl.pattern==='solid'){const fg=colorOf(fl.fgColor);if(fg&&fg!=='#ffffff'||fg==='#ffffff'&&false)st.fill=fg;}
          const al=cell.alignment;if(al){if(al.horizontal&&al.horizontal!=='general')st.ha=al.horizontal==='centerContinuous'?'center':al.horizontal;if(al.vertical&&al.vertical!=='bottom')st.va=al.vertical==='center'?'middle':al.vertical;if(al.wrapText)st.wrap=1;if(al.indent)st.ind=al.indent;if(al.textRotation&&al.textRotation!=='vertical')st.rot=al.textRotation>90?-(al.textRotation-90):al.textRotation;}
          const b=cell.border;if(b){const bd={};for(const [k,n] of [['t','top'],['r','right'],['b','bottom'],['l','left']])if(b[n]&&b[n].style)bd[k]=[BTYPES[b[n].style]?b[n].style:'thin',colorOf(b[n].color)||'#000000'];if(Object.keys(bd).length)st.bd=bd;}
          if(cell.numFmt&&cell.numFmt!=='General')st.fmt=cell.numFmt;
          if(Object.keys(st).length)o.st=st;
          if(o.v!==null||o.f!==undefined||o.st)s.cells.set(key(rn-1,cn-1),o);
        });
      });
      for(const m of (ws.model.merges||[])){const p=m.split(':');const a=parseAddr(p[0]),b=parseAddr(p[1]||p[0]);if(a&&b)s.merges.push({r1:a.r,c1:a.c,r2:b.r,c2:b.c});}
      const af=ws.autoFilter;if(af){let a,b;if(typeof af==='string'){const p=af.split(':');a=parseAddr(p[0]);b=parseAddr(p[1]||p[0]);}else if(af.from&&af.to){a={r:af.from.row-1,c:af.from.column-1};b={r:af.to.row-1,c:af.to.column-1};}if(a&&b)s.filter={r1:a.r,c1:a.c,r2:b.r,c2:b.c,crit:{}};}
      const tc=ws.properties&&ws.properties.tabColor;if(tc)s.tab=colorOf(tc);
      try{for(const cf of (ws.conditionalFormattings||[])){const p=cf.ref.split(':');const a=parseAddr(p[0]),b=parseAddr(p[1]||p[0]);if(!a||!b)continue;
        for(const rl of cf.rules||[]){const fm={};const sty=rl.style||{};if(sty.fill){const c=colorOf(sty.fill.bgColor)||colorOf(sty.fill.fgColor);if(c)fm.fill=c;}if(sty.font){const c=colorOf(sty.font.color);if(c)fm.color=c;if(sty.font.bold)fm.b=1;}
          const base={r1:a.r,c1:a.c,r2:b.r,c2:b.c,sh:s};
          if(rl.type==='cellIs'){const op={greaterThan:'gt',lessThan:'lt',greaterThanOrEqual:'ge',lessThanOrEqual:'le',equal:'eq',between:'btw'}[rl.operator];if(op)s.cf.push({...base,type:'cell',op,a:isNaN(+rl.formulae[0])?String(rl.formulae[0]).replace(/^"|"$/g,''):+rl.formulae[0],b:rl.formulae[1]!==undefined?+rl.formulae[1]:undefined,fmt:fm});}
          else if(rl.type==='containsText')s.cf.push({...base,type:'text',a:rl.text,fmt:fm});
          else if(rl.type==='top10')s.cf.push({...base,type:rl.bottom?'bottom':'top',n:rl.rank||10,pct:!!rl.percent,fmt:fm});
          else if(rl.type==='aboveAverage')s.cf.push({...base,type:rl.aboveAverage===false?'below':'above',fmt:fm});
          else if(rl.type==='colorScale'){const cs=(rl.color||[]).map(colorOf).filter(Boolean);if(cs.length>=2)s.cf.push({...base,type:'scale',colors:cs});}
          else if(rl.type==='dataBar'){s.cf.push({...base,type:'bar',color:colorOf(rl.color)||'#638ec6'});}
        }}}catch(e){}
      sheets.push(s);
    }
    return sheets;
  }
  function parseCsv(t){
    t=t.replace(/^\uFEFF/,'');const first=t.split(/\r?\n/)[0]||'';
    const cnt=ch=>(first.match(new RegExp(ch==='\t'?'\t':'\\'+ch,'g'))||[]).length;
    const del=[[';',cnt(';')],[',',cnt(',')],['\t',cnt('\t')]].sort((a,b)=>b[1]-a[1])[0][0];
    const rows=[];let row=[],f='',q=false;
    for(let i=0;i<t.length;i++){const ch=t[i];
      if(q){if(ch==='"'){if(t[i+1]==='"'){f+='"';i++;}else q=false;}else f+=ch;}
      else if(ch==='"'&&f==='')q=true;else if(ch===del){row.push(f);f='';}
      else if(ch==='\n'||ch==='\r'){if(ch==='\r'&&t[i+1]==='\n')i++;row.push(f);rows.push(row);row=[];f='';}else f+=ch;}
    if(f!==''||row.length){row.push(f);rows.push(row);}
    return rows;
  }
  async function openFile(file){
    try{
      const ext=file.name.split('.').pop().toLowerCase();let sheets;
      if(ext==='xlsx'||ext==='xlsm')sheets=await importXlsx(await file.arrayBuffer());
      else{const rows=parseCsv(await file.text());const s=newSheet(file.name.replace(/\.[^.]+$/,'').slice(0,31)||'Sheet1');
        rows.forEach((row,r)=>row.forEach((t,c)=>{if(t==='')return;const o=mkInput(t.trim()===t?t:t,null);if(o)s.cells.set(key(r,c),o);}));sheets=[s];}
      if(!sheets.length)sheets=[newSheet('Sheet1')];
      wb.sheets=sheets;wb.active=0;cur=sheets[0];undoS=[];redoS=[];clip=null;dirty=false;
      sheets.forEach(s=>{const rs=[];for(const [k,cl] of s.cells)if(cl.st&&cl.st.wrap)rs.push(Math.floor(k/MAXC));if(rs.length){const save=cur;cur=s;fitRows([...new Set(rs)].filter(r=>!(s.rh.ov.has(r))));cur=save;}if(s.filter){const save=cur;cur=s;applyFilter(true);cur=save;}});
      recalcAll();setName(file.name.replace(/\.(csv|tsv|txt|xlsm)$/i,'.xlsx'));updUndo();drawTabs();
      vp.scrollLeft=0;vp.scrollTop=0;select(0,0);
    }catch(e){console.error(e);toast('Tidak bisa membuka berkas: '+e.message);}
  }
  function confirmNew(){
    const run=()=>{wb.sheets=[newSheet('Sheet1')];wb.active=0;cur=wb.sheets[0];undoS=[];redoS=[];clip=null;dirty=false;setName('Buku1.xlsx');recalcAll();updUndo();drawTabs();vp.scrollLeft=0;vp.scrollTop=0;select(0,0);};
    if(dirty)confirmDlg('Perubahan belum disimpan akan hilang. Lanjutkan membuat buku kerja baru?',run);else run();
  }
  const dl=(blob,fn)=>{const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=fn;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);toast('Tersimpan: '+fn);};
  async function saveXlsx(){
    try{
      if(ed&&!commitEdit())return;
      await lib('exceljs',()=>window.ExcelJS);
      const w=new ExcelJS.Workbook();w.creator='Office Web';
      for(const s of wb.sheets){
        const ws=w.addWorksheet(s.name,{properties:s.tab?{tabColor:{argb:argb(s.tab)}}:{},views:[{showGridLines:showGrid}]});
        if(s.cw.def!==DW)ws.properties.defaultColWidth=(s.cw.def-5)/7;
        for(const [c,v] of s.cw.ov)ws.getColumn(c+1).width=Math.max(0,(v-5)/7);
        for(const c of s.cw.hid)ws.getColumn(c+1).hidden=true;
        for(const [r,v] of s.rh.ov)ws.getRow(r+1).height=v*.75;
        for(const r of s.rh.hid)ws.getRow(r+1).hidden=true;
        for(const [k,cl] of s.cells){
          const r=Math.floor(k/MAXC),c=k%MAXC,x=ws.getCell(r+1,c+1);
          if(cl.f!==undefined&&cl.f!==null){const v=rawVal(wb,s,r,c);x.value={formula:cl.f,result:isErr(v)?{error:v.code}:v};}
          else if(cl.v!==null&&cl.v!==undefined)x.value=cl.v;
          const st=cl.st;if(!st)continue;
          if(st.font||st.sz||st.b||st.i||st.u||st.s||st.color){x.font={name:st.font||'Calibri',size:st.sz||11,bold:!!st.b,italic:!!st.i,underline:st.u?(st.u===2?'double':true):false,strike:!!st.s,...(st.color?{color:{argb:argb(st.color)}}:{})};}
          if(st.fill)x.fill={type:'pattern',pattern:'solid',fgColor:{argb:argb(st.fill)}};
          if(st.ha||st.va||st.wrap||st.ind||st.rot){x.alignment={horizontal:st.ha,vertical:st.va==='middle'?'middle':st.va,wrapText:!!st.wrap,indent:st.ind||0,textRotation:st.rot?(st.rot===255?255:st.rot>0?st.rot:90-st.rot):0};}
          if(st.bd){const b={};for(const [k2,n] of [['t','top'],['r','right'],['b','bottom'],['l','left']])if(st.bd[k2])b[n]={style:st.bd[k2][0],color:{argb:argb(st.bd[k2][1])}};x.border=b;}
          if(st.fmt)x.numFmt=st.fmt;
        }
        for(const m of s.merges)try{ws.mergeCells(m.r1+1,m.c1+1,m.r2+1,m.c2+1);}catch(e){}
        if(s.filter)ws.autoFilter={from:{row:s.filter.r1+1,column:s.filter.c1+1},to:{row:s.filter.r2+1,column:s.filter.c2+1}};
        for(const rl of s.cf){
          const ref=addr(rl.r1,rl.c1)+':'+addr(rl.r2,rl.c2);
          const sty=rl.fmt?{...(rl.fmt.fill?{fill:{type:'pattern',pattern:'solid',bgColor:{argb:argb(rl.fmt.fill)}}}:{}),font:{...(rl.fmt.color?{color:{argb:argb(rl.fmt.color)}}:{}),...(rl.fmt.b?{bold:true}:{})}}:undefined;
          let rule=null;
          if(rl.type==='cell'&&typeof rl.a!=='string')rule={type:'cellIs',operator:{gt:'greaterThan',lt:'lessThan',ge:'greaterThanOrEqual',le:'lessThanOrEqual',eq:'equal',btw:'between'}[rl.op],formulae:rl.op==='btw'?[rl.a,rl.b]:[rl.a],style:sty};
          else if(rl.type==='text')rule={type:'containsText',operator:'containsText',text:rl.a,style:sty};
          else if(rl.type==='top'||rl.type==='bottom')rule={type:'top10',rank:rl.n,percent:!!rl.pct,bottom:rl.type==='bottom',style:sty};
          else if(rl.type==='above'||rl.type==='below')rule={type:'aboveAverage',aboveAverage:rl.type==='above',style:sty};
          else if(rl.type==='scale')rule={type:'colorScale',cfvo:rl.colors.length===2?[{type:'min'},{type:'max'}]:[{type:'min'},{type:'percentile',value:50},{type:'max'}],color:rl.colors.map(c=>({argb:argb(c)}))};
          else if(rl.type==='bar')rule={type:'dataBar',cfvo:[{type:'min'},{type:'max'}],color:{argb:argb(rl.color)}};
          else if(rl.type==='dup'||rl.type==='uniq')rule={type:'expression',formulae:[`COUNTIF(${ref.replace(/([A-Z]+)(\d+)/g,'$$$1$$$2')},${addr(rl.r1,rl.c1)})${rl.type==='dup'?'>1':'=1'}`],style:sty};
          if(rule)try{ws.addConditionalFormatting({ref,rules:[{priority:1,...rule}]});}catch(e){}
        }
      }
      const buf=await w.xlsx.writeBuffer();
      dl(new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),/\.xlsx$/i.test(name)?name:name.replace(/\.[^.]+$/,'')+'.xlsx');
      dirty=false;
    }catch(e){console.error(e);toast('Gagal menyimpan: '+e.message);}
  }
  function saveCsv(){
    const u=usedRange(cur);const lines=[];
    if(u.any)for(let r=0;r<=u.r2;r++){const row=[];for(let c=0;c<=u.c2;c++){const cl=cellAt(cur,r,c);const v=cl?cellVal(cur,r,c):null;let t=v===null?'':typeof v==='number'?String(v):fmtValue(v,cl&&cl.st&&cl.st.fmt).t;row.push(/[",\n;]/.test(t)?'"'+t.replace(/"/g,'""')+'"':t);}lines.push(row.join(','));}
    dl(new Blob(['\uFEFF'+lines.join('\r\n')],{type:'text/csv;charset=utf-8'}),name.replace(/\.[^.]+$/,'')+'.csv');
  }
  function printSheet(){
    const u=usedRange(cur);let h='<table style="border-collapse:collapse;font:11pt Calibri,sans-serif">';
    if(u.any)for(let r=0;r<=u.r2;r++){if(cur.rh.size(r)===0)continue;h+='<tr>';for(let c=0;c<=u.c2;c++){if(cur.cw.size(c)===0)continue;const m=mergeAtSh(cur,r,c);if(m&&(m.r1!==r||m.c1!==c))continue;
      const cl=cellAt(cur,r,c),st=(cl&&cl.st)||{},f=fmtValue(cl?cellVal(cur,r,c):null,st.fmt);
      h+=`<td ${m?`rowspan="${m.r2-m.r1+1}" colspan="${m.c2-m.c1+1}"`:''} style="padding:1px 4px;height:${cur.rh.size(r)*.75}pt;min-width:${cur.cw.size(c)}px;text-align:${st.ha&&st.ha!=='general'?st.ha:(f.al||'left')};${st.fill?'background:'+st.fill+';':''}${st.color?'color:'+st.color+';':''}${st.b?'font-weight:700;':''}${st.i?'font-style:italic;':''}${st.sz?'font-size:'+st.sz+'pt;':''}${st.bd?SIDES.map(s=>st.bd[s]?`border-${{t:'top',r:'right',b:'bottom',l:'left'}[s]}:${BTYPES[st.bd[s][0]].join(' ')} ${st.bd[s][1]};`:'').join(''):''}">${esc(f.t)}</td>`;}h+='</tr>';}
    h+='</table>';
    const p=document.createElement('div');p.id='xl-print';p.innerHTML=h;document.body.appendChild(p);
    const done=()=>{removeEventListener('afterprint',done);p.remove();};addEventListener('afterprint',done);window.print();
  }

  /* ===== event ===== */
  root.addEventListener('mousedown',e=>{if(e.target.closest('.b,.tab,.stbar button,.gal button,.gmore button,.fb,.shbar button,.shadd')){e.preventDefault();}});
  root.addEventListener('click',e=>{
    const t=e.target.closest('[data-a],[data-tab]');
    if(!t){const sh=e.target.closest('[data-sh]');if(sh&&!e.target.closest('input'))switchSheet(+sh.dataset.sh);return;}
    if(t.dataset.tab){
      const small=matchMedia('(max-width:700px),(max-height:500px)').matches;
      if(small&&t.classList.contains('on')&&!X.classList.contains('rc')){X.classList.add('rc');t.classList.remove('on');return;}
      X.classList.remove('rc');
      $$('.tab[data-tab]').forEach(x=>x.classList.toggle('on',x===t));$$('[data-pane]').forEach(p=>p.hidden=p.dataset.pane!==t.dataset.tab);return;}
    const a=t.dataset.a;
    if(a.startsWith('m:')){const n=a.slice(2);if(menus[n])openMenu(t,menus[n]());return;}
    const i=a.indexOf(':'),k=i<0?a:a.slice(0,i),v=i<0?undefined:a.slice(i+1);
    if(acts[k])acts[k](v);
    if(k!=='menu')menu.hidden=true;
    if(['bold','italic','underline','paste','cut','copy','painter'].includes(k)&&!ed)vp.focus({preventScroll:true});
  });
  root.addEventListener('dblclick',e=>{const sh=e.target.closest('[data-sh]');if(sh){switchSheet(+sh.dataset.sh);renameSheet(+sh.dataset.sh);}});
  root.addEventListener('contextmenu',e=>{const sh=e.target.closest('[data-sh]');if(sh){e.preventDefault();switchSheet(+sh.dataset.sh);openMenuAt(e.clientX,e.clientY-120,menus.sheetctx());}});
  root.addEventListener('change',e=>{const s=e.target.dataset&&e.target.dataset.s;if(!s||!sels[s])return;sels[s](e.target.value);if(!ed)vp.focus({preventScroll:true});});
  $('.zr').oninput=e=>setZoom(+e.target.value);
  vp.addEventListener('mousedown',()=>{menu.hidden=true;});
  root.addEventListener('keydown',e=>{
    if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='s'){e.preventDefault();saveXlsx();}
  });
  fin.onchange=()=>{if(fin.files[0])openFile(fin.files[0]);fin.value='';};
  if(window.ResizeObserver)new ResizeObserver(()=>{if(root.isConnected)schedule();}).observe(vp);
  if(matchMedia('(max-width:700px),(max-height:500px)').matches){X.classList.add('rc');$$('.tab[data-tab]').forEach(x=>x.classList.remove('on'));}

  /* ===== mulai ===== */
  wb.sheets=[newSheet('Sheet1')];cur=wb.sheets[0];
  setName(opts.file?opts.file.name.replace(/\.(csv|tsv|txt|xlsm)$/i,'.xlsx'):'Buku1.xlsx');
  recalcAll();drawTabs();updUndo();sync();
  if(opts.file)openFile(opts.file);else vp.focus({preventScroll:true});
  window.__xl={wb:()=>wb,cur:()=>cur,tx,select,acts,menus,formatCells,applyStyle,putCell,mkInput,recalcAll,startEdit,commitEdit,rawText,key,addr,cellVal,sync};
}

window.Apps=window.Apps||{};
Apps.excel={mount};
})();
