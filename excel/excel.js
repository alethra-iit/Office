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
  if(src.includes(';')){ /* gaya Indonesia: pemisah argumen ';' dan koma desimal */
    let o='',q=false;for(let k=0;k<src.length;k++){const ch=src[k];if(ch==='"')q=!q;
      if(!q&&ch===','&&/\d/.test(src[k-1]||'')&&/\d/.test(src[k+1]||'')&&!/[A-Za-z_$]\d*$/.test(o.replace(/\d+$/,'')+'')){o+='.';}else o+=ch;}
    src=o;}
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
    case 'name':{
      const nm=cx.wb.names&&cx.wb.names.find(n=>n.name.toLowerCase()===a.v.toLowerCase());
      if(!nm)return E.NAME;
      if(nm.ast===undefined){try{nm.ast=parseFormula(nm.ref.replace(/^=/,''));}catch(e){nm.ast=null;}}
      if(nm.ast===null||nm.depth>8)return E.NAME;
      nm.depth=(nm.depth||0)+1;let r;try{r=evalAst(nm.ast,{wb:cx.wb,sh:cx.sh,cell:cx.cell});}finally{nm.depth--;}
      return r;}
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
def('SUBTOTAL',(a,cx)=>{
  const code=num(S1(a,cx));if(isErr(code))return code;
  const base=code>100?code-100:code,skipHid=code>100;
  if(base<1||base>11)return E.VAL;
  const vals=[],all=[];
  for(const x of a.slice(1)){
    if(!(x instanceof Rng)){if(typeof x==='number')vals.push(x);continue;}
    for(const [r,c] of cellsIn(x)){
      const cl=x.sh.cells.get(key(r,c));if(!cl)continue;
      if(cl.f&&/^SUBTOTAL\(/i.test(cl.f))continue;
      if(x.sh.rh.fil.has(r)||(skipHid&&x.sh.rh.hid.has(r)))continue;
      const v=rawVal(cx.wb,x.sh,r,c);if(isErr(v))return v;
      if(v!==null&&v!=='')all.push(v);if(typeof v==='number')vals.push(v);}
  }
  const sum=vals.reduce((p,q)=>p+q,0),n=vals.length;
  const vr=pop=>{if(n<(pop?1:2))return E.DIV;const m=sum/n;return vals.reduce((p,q)=>p+(q-m)**2,0)/(n-(pop?0:1));};
  switch(base){case 1:return n?sum/n:E.DIV;case 2:return n;case 3:return all.length;case 4:return n?Math.max(...vals):0;case 5:return n?Math.min(...vals):0;
    case 6:return n?vals.reduce((p,q)=>p*q,1):0;case 7:{const v=vr(false);return isErr(v)?v:Math.sqrt(v);}case 8:{const v=vr(true);return isErr(v)?v:Math.sqrt(v);}
    case 9:return sum;case 10:return vr(false);default:return vr(true);}
});

/* semua fungsi yang menerima nilai tunggal dengan kosong -> 0 */


/* ---------- fungsi keuangan ---------- */
const FA=(a,i,d,cx)=>a[i]===undefined||a[i]===null?d:num(scalar(a[i],cx));
const pw=(r,n)=>Math.pow(1+r,n);
function fvF(r,n,p,pv,t){return r===0?-(pv+p*n):-(pv*pw(r,n)+p*(1+r*t)*(pw(r,n)-1)/r);}
function pmtF(r,n,pv,fv,t){return r===0?-(pv+fv)/n:-(r*(pv*pw(r,n)+fv))/((1+r*t)*(pw(r,n)-1));}
const fin=(n,f)=>def(n,(a,cx)=>{const v=f(a,cx);return typeof v==='number'&&!isFinite(v)?E.NUM:v;});
const chk=(...v)=>v.find(isErr);
fin('PMT',(a,cx)=>{const r=FA(a,0,0,cx),n=FA(a,1,0,cx),pv=FA(a,2,0,cx),fv=FA(a,3,0,cx),t=FA(a,4,0,cx);return chk(r,n,pv,fv,t)||(n===0?E.NUM:pmtF(r,n,pv,fv,t?1:0));});
fin('FV',(a,cx)=>{const r=FA(a,0,0,cx),n=FA(a,1,0,cx),p=FA(a,2,0,cx),pv=FA(a,3,0,cx),t=FA(a,4,0,cx);return chk(r,n,p,pv,t)||fvF(r,n,p,pv,t?1:0);});
fin('PV',(a,cx)=>{const r=FA(a,0,0,cx),n=FA(a,1,0,cx),p=FA(a,2,0,cx),fv=FA(a,3,0,cx),t0=FA(a,4,0,cx);if(chk(r,n,p,fv,t0))return chk(r,n,p,fv,t0);const t=t0?1:0;return r===0?-(fv+p*n):-(fv+p*(1+r*t)*(pw(r,n)-1)/r)/pw(r,n);});
fin('NPER',(a,cx)=>{const r=FA(a,0,0,cx),p=FA(a,1,0,cx),pv=FA(a,2,0,cx),fv=FA(a,3,0,cx),t=FA(a,4,0,cx);if(chk(r,p,pv,fv,t))return chk(r,p,pv,fv,t);const tt=t?1:0;if(r===0)return p===0?E.NUM:-(pv+fv)/p;const q=(p*(1+r*tt)-fv*r)/(p*(1+r*tt)+pv*r);return q<=0?E.NUM:Math.log(q)/Math.log(1+r);});
fin('NPV',(a,cx)=>{const r=FA(a,0,0,cx);if(isErr(r))return r;const v=nums(a.slice(1),cx);if(isErr(v))return v;return v.reduce((s,x,i)=>s+x/pw(r,i+1),0);});
fin('SLN',(a,cx)=>{const c=FA(a,0,0,cx),s2=FA(a,1,0,cx),l=FA(a,2,0,cx);return chk(c,s2,l)||(l===0?E.DIV:(c-s2)/l);});
fin('IPMT',(a,cx)=>{const r=FA(a,0,0,cx),per=FA(a,1,0,cx),n=FA(a,2,0,cx),pv=FA(a,3,0,cx),fv=FA(a,4,0,cx);if(chk(r,per,n,pv,fv))return chk(r,per,n,pv,fv);if(per<1||per>n)return E.NUM;const p=pmtF(r,n,pv,fv,0);return fvF(r,per-1,p,pv,0)*r;});
fin('PPMT',(a,cx)=>{const r=FA(a,0,0,cx),per=FA(a,1,0,cx),n=FA(a,2,0,cx),pv=FA(a,3,0,cx),fv=FA(a,4,0,cx);if(chk(r,per,n,pv,fv))return chk(r,per,n,pv,fv);if(per<1||per>n)return E.NUM;const p=pmtF(r,n,pv,fv,0);return p-fvF(r,per-1,p,pv,0)*r;});
fin('IRR',(a,cx)=>{const v=nums([a[0]],cx);if(isErr(v))return v;if(!v.some(x=>x>0)||!v.some(x=>x<0))return E.NUM;
  const f=r=>v.reduce((s,x,i)=>s+x/Math.pow(1+r,i),0);let lo=-0.9999,hi=10,flo=f(lo),fhi=f(hi);if(flo*fhi>0)return E.NUM;
  for(let i=0;i<200;i++){const m=(lo+hi)/2,fm=f(m);if(Math.abs(fm)<1e-10)return m;if(fm*flo<0){hi=m;fhi=fm;}else{lo=m;flo=fm;}}return (lo+hi)/2;});
fin('RATE',(a,cx)=>{const n=FA(a,0,0,cx),p=FA(a,1,0,cx),pv=FA(a,2,0,cx),fv=FA(a,3,0,cx),t=FA(a,4,0,cx),g=FA(a,5,.1,cx);if(chk(n,p,pv,fv,t,g))return chk(n,p,pv,fv,t,g);
  const tt=t?1:0,f=r=>r===0?pv+p*n+fv:pv*pw(r,n)+p*(1+r*tt)*(pw(r,n)-1)/r+fv;let r=g;
  for(let i=0;i<100;i++){const y=f(r),h=1e-7,d=(f(r+h)-y)/h;if(!d||!isFinite(d))break;const nr=r-y/d;if(Math.abs(nr-r)<1e-10)return nr;r=nr;if(r<=-1)r=-0.99;}return E.NUM;});
def('TRANSPOSE',(a,cx)=>E.VAL);

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
let PAL=[
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
.xl .ct{display:flex;flex-direction:column;overflow:hidden;pointer-events:none;font:11pt var(--df,Calibri),Carlito,"Segoe UI",sans-serif;color:#000;line-height:1.2;padding:0 2px}
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
.xl .ed{z-index:4;border:2px solid var(--g);background:#fff;outline:none;resize:none;overflow:hidden;font:11pt var(--df,Calibri),Carlito,"Segoe UI",sans-serif;color:#000;padding:0 2px;line-height:1.2;user-select:text;-webkit-user-select:text;white-space:pre;margin:0}
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
.xl .ob{pointer-events:auto;cursor:move;z-index:2}
.xl .ob.sel{outline:2px solid var(--g)}
.xl .ob .orz{position:absolute;right:-5px;bottom:-5px;width:9px;height:9px;background:#fff;border:1px solid var(--g);cursor:nwse-resize}
.xl .ob svg{display:block;pointer-events:none}
.xl .ob img{display:block;width:100%;height:100%;pointer-events:none;user-select:none;-webkit-user-drag:none}
.xl .ob .tx{width:100%;height:100%;display:flex;align-items:center;justify-content:center;text-align:center;font-family:var(--df,Calibri),Carlito,sans-serif;overflow:hidden;padding:4px;white-space:pre-wrap;pointer-events:none;box-sizing:border-box}
.xl .sym{display:grid;grid-template-columns:repeat(10,32px);gap:2px}
.xl .sym button{width:32px;height:30px;border:1px solid #ccc;background:#fff;cursor:pointer;font-size:15px;padding:0}
.xl .sym button:hover{background:#d3e8d9}
.xl .chg{display:grid;grid-template-columns:auto auto;gap:2px 0;align-self:flex-start}
.xl .chg .b{height:21px}
.xl .corner{grid-column:1;grid-row:1}.xl .colh{grid-column:2;grid-row:1}.xl .rowh{grid-column:1;grid-row:2}.xl .vp{grid-column:2;grid-row:2}
.xl .gwrap.nh .corner,.xl .gwrap.nh .colh,.xl .gwrap.nh .rowh{display:none}
.xl .row label{display:flex;gap:4px;align-items:center;font-size:12px;cursor:pointer}
.xl .selp{position:absolute;right:22px;top:128px;width:230px;max-height:50%;overflow:auto;background:#fff;border:1px solid #aaa;box-shadow:0 2px 8px rgba(0,0,0,.25);z-index:20;font-size:12px}
.xl .selp .sh{display:flex;justify-content:space-between;align-items:center;padding:6px 8px;background:#f3f3f3;font-weight:600;color:var(--g)}
.xl .selp .sh button{border:0;background:none;cursor:pointer}
.xl .selp .si{display:flex;gap:6px;align-items:center;padding:4px 8px;cursor:pointer}
.xl .selp .si:hover{background:#eef6f0}.xl .selp .si.on{background:#d3e8d9}
.xl .selp .si button{border:0;background:none;cursor:pointer;width:20px}
.xl .ogl{position:absolute;background:#7f7f7f}
.xl .ogb{position:absolute;width:12px;height:12px;border:1px solid #7f7f7f;background:#fff;font-size:11px;line-height:8px;padding:0;cursor:pointer;display:grid;place-items:center;color:#333}
.xl .ogb:hover{background:#d3e8d9}
.xl .dvtip{background:#ffffe1;border:1px solid #888;padding:4px 7px;font-size:11px;max-width:220px;pointer-events:none;z-index:6;color:#222;box-shadow:1px 1px 3px rgba(0,0,0,.25)}
.xl .cmt{position:absolute;width:0;height:0;border-top:6px solid #e53935;border-left:6px solid transparent;pointer-events:none;z-index:2}
.xl .cbox{position:absolute;background:#ffffe1;border:1px solid #666;padding:4px 6px;font-size:11px;color:#222;z-index:4;pointer-events:none;box-shadow:2px 2px 4px rgba(0,0,0,.25);white-space:pre-wrap;word-break:break-word}
.xl .trk{position:absolute;border:2px solid #4a86e8;pointer-events:none;z-index:2}
.xl .trk::before{content:"";position:absolute;left:0;top:0;border-top:6px solid #4a86e8;border-right:6px solid transparent}
.xl .rgn{position:absolute;overflow:hidden;pointer-events:none}
.xl .rgn>*{position:absolute}
.xl .rgn .ob{pointer-events:auto}
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
const BD=(ic,tip,lab)=>`<button class="b big" disabled title="${tip} (belum tersedia)"><em>${ic}</em><span>${lab||tip}</span></button>`;
const CHI={
 chCol:'<svg width="16" height="16"><rect x="2" y="7" width="3" height="7" fill="#4472c4"/><rect x="6.5" y="3" width="3" height="11" fill="#ed7d31"/><rect x="11" y="9" width="3" height="5" fill="#4472c4"/></svg>',
 chBar:'<svg width="16" height="16"><rect x="2" y="2" width="7" height="3" fill="#4472c4"/><rect x="2" y="6.5" width="11" height="3" fill="#ed7d31"/><rect x="2" y="11" width="5" height="3" fill="#4472c4"/></svg>',
 chLine:'<svg width="16" height="16"><polyline points="2,12 6,6 10,9 14,3" fill="none" stroke="#4472c4" stroke-width="2"/><polyline points="2,14 6,11 10,13 14,8" fill="none" stroke="#ed7d31" stroke-width="2"/></svg>',
 chArea:'<svg width="16" height="16"><polygon points="2,14 2,8 6,5 10,9 14,3 14,14" fill="#5b9bd5"/></svg>',
 chPie:'<svg width="16" height="16"><circle cx="8" cy="8" r="6" fill="#ed7d31"/><path d="M8 8V2A6 6 0 0 1 14 8z" fill="#4472c4"/></svg>',
 chSct:'<svg width="16" height="16"><circle cx="4" cy="11" r="1.6" fill="#4472c4"/><circle cx="8" cy="6" r="1.6" fill="#4472c4"/><circle cx="12" cy="9" r="1.6" fill="#4472c4"/><circle cx="12" cy="3.5" r="1.6" fill="#4472c4"/></svg>'};
const CB=(k,tip)=>`<button class="b" title="${tip}" data-a="m:${k}"><em>${CHI[k]}</em><i class="car">▾</i></button>`;
RIB.insert=
 G('Tabel',BD('⊞','PivotTable')+BD('⊟','Rekomendasi PivotTable','Rekomendasi<br>PivotTable')+B('▦','Tabel','tblnow','big','Tabel'))+
 G('Ilustrasi',B('🖼','Sisipkan gambar dari komputer','pic','big','Gambar')+BD('🌐','Gambar Online','Gambar<br>Online')+B('◇','Bentuk','m:shapes','big','Bentuk',1)+BD('▤','SmartArt')+BD('📷','Tangkapan Layar','Tangkapan<br>Layar'))+
 G('Aplikasi',BD('🛍','Aplikasi untuk Office','Aplikasi untuk<br>Office'))+
 G('Bagan',BD('📈','Bagan yang Disarankan','Bagan yang<br>Disarankan')+`<div class="chg">${CB('chCol','Sisipkan bagan kolom')}${CB('chBar','Sisipkan bagan batang')}${CB('chLine','Sisipkan bagan garis')}${CB('chArea','Sisipkan bagan area')}${CB('chPie','Sisipkan bagan pai')}${CB('chSct','Sisipkan bagan sebar')}</div>`+BD('📊','PivotChart'))+
 G('Laporan',BD('📊','Power View'))+
 G('Sparkline',B('〰','Sparkline Garis','spark:line','big','Garis')+B('▮▯','Sparkline Kolom','spark:col','big','Kolom')+B('▀▄','Sparkline Menang/Kalah','spark:wl','big','Menang/<br>Kalah'))+
 G('Filter',BD('⧩','Slicer')+BD('⏱','Timeline'))+
 G('Tautan',B('🔗','Hyperlink','hyperlink','big','Hyperlink'))+
 G('Teks',B('A','Kotak Teks','textbox','big','Kotak Teks')+BD('▭','Header & Footer','Header<br>& Footer')+BD('𝒜','WordArt')+BD('✍','Baris Tanda Tangan','Baris Tanda<br>Tangan')+BD('▢','Objek'))+
 G('Simbol',BD('π','Persamaan')+B('Ω','Simbol','symbol','big','Simbol'));
const FITO=[['0','Otomatis'],...[1,2,3,4,5,6,7,8,9].map(n=>[String(n),n+' halaman'])];
RIB.layout=
 G('Tema',B('Aa','Tema','m:themes','big','Tema',1)+`<div class="rows t3">${B('◍','Warna','m:tcolors','rw','Warna',1)}${B('A','Font','m:tfonts','rw','Font',1)}<button class="b rw" disabled title="Efek (belum tersedia)"><em>◉</em><span>Efek ▾</span></button></div>`)+
 G('Penyiapan Halaman',B('▭','Margin','m:margins','big','Margin',1)+B('⬒','Orientasi','m:orient2','big','Orientasi',1)+B('▯','Ukuran','m:psize','big','Ukuran',1)+B('▦','Area Cetak','m:parea','big','Area<br>Cetak',1)+BD('⇥','Sela')+B('🖼','Latar Belakang','bg','big','Latar<br>Belakang')+B('⊞','Judul Cetak','ptitles','big','Judul<br>Cetak'),'page')+
 G('Skala agar Muat',`<div class="rows t3"><div class="row"><span style="width:46px">Lebar:</span>${S('fitw',FITO,98)}</div><div class="row"><span style="width:46px">Tinggi:</span>${S('fith',FITO,98)}</div><div class="row"><span style="width:46px">Skala:</span><input type="number" data-s="scale" min="10" max="400" value="100" style="width:60px;height:23px;border:1px solid #ababab"> %</div></div>`,'page')+
 G('Opsi Lembar',`<div class="rows t3" style="min-width:72px"><div class="row"><b>Garis Kisi</b></div><div class="row"><label><input type="checkbox" data-s="gview"> Tampilan</label></div><div class="row"><label><input type="checkbox" data-s="gprint"> Cetak</label></div></div><div class="rows t3" style="min-width:72px"><div class="row"><b>Tajuk</b></div><div class="row"><label><input type="checkbox" data-s="hview"> Tampilan</label></div><div class="row"><label><input type="checkbox" data-s="hprint"> Cetak</label></div></div>`,'page')+
 G('Susun',B('⬆','Majukan','m:fwd','big','Majukan',1)+B('⬇','Mundurkan','m:back','big','Mundurkan',1)+B('▤','Panel Seleksi','selpane','big','Panel<br>Seleksi')+B('⇤','Ratakan','m:align','big','Ratakan',1)+BD('▣','Grup')+B('⟳','Putar','m:rot','big','Putar',1));

const FNL=`SUM|mt|SUM(angka1; [angka2]; …)|Menjumlahkan semua angka dalam rentang sel.
SUMIF|mt|SUMIF(rentang; kriteria; [rentang_jumlah])|Menjumlahkan sel yang memenuhi satu kriteria.
SUMIFS|mt|SUMIFS(rentang_jumlah; rentang_kriteria1; kriteria1; …)|Menjumlahkan sel yang memenuhi beberapa kriteria.
SUMPRODUCT|mt|SUMPRODUCT(larik1; [larik2]; …)|Mengalikan komponen larik lalu menjumlahkan hasilnya.
PRODUCT|mt|PRODUCT(angka1; [angka2]; …)|Mengalikan semua angka pada argumen.
ABS|mt|ABS(angka)|Mengembalikan nilai absolut angka.
SQRT|mt|SQRT(angka)|Mengembalikan akar kuadrat.
POWER|mt|POWER(angka; pangkat)|Memangkatkan angka.
EXP|mt|EXP(angka)|Mengembalikan e dipangkatkan angka.
LN|mt|LN(angka)|Logaritma natural.
LOG|mt|LOG(angka; [basis])|Logaritma dengan basis tertentu.
LOG10|mt|LOG10(angka)|Logaritma basis 10.
MOD|mt|MOD(angka; pembagi)|Sisa hasil bagi.
QUOTIENT|mt|QUOTIENT(pembilang; penyebut)|Bagian bulat dari pembagian.
INT|mt|INT(angka)|Membulatkan ke bawah ke bilangan bulat terdekat.
ROUND|mt|ROUND(angka; jumlah_digit)|Membulatkan angka ke jumlah digit tertentu.
ROUNDUP|mt|ROUNDUP(angka; jumlah_digit)|Membulatkan ke atas, menjauhi nol.
ROUNDDOWN|mt|ROUNDDOWN(angka; jumlah_digit)|Membulatkan ke bawah, mendekati nol.
TRUNC|mt|TRUNC(angka; [jumlah_digit])|Memotong angka menjadi bilangan bulat.
CEILING|mt|CEILING(angka; kelipatan)|Membulatkan ke atas ke kelipatan terdekat.
FLOOR|mt|FLOOR(angka; kelipatan)|Membulatkan ke bawah ke kelipatan terdekat.
MROUND|mt|MROUND(angka; kelipatan)|Membulatkan ke kelipatan yang diinginkan.
EVEN|mt|EVEN(angka)|Membulatkan ke bilangan genap terdekat (menjauhi nol).
ODD|mt|ODD(angka)|Membulatkan ke bilangan ganjil terdekat (menjauhi nol).
SIGN|mt|SIGN(angka)|Tanda angka: 1, 0, atau -1.
FACT|mt|FACT(angka)|Faktorial.
PI|mt|PI()|Nilai pi, 3,14159265358979.
SIN|mt|SIN(angka)|Sinus sudut (radian).
COS|mt|COS(angka)|Kosinus sudut (radian).
TAN|mt|TAN(angka)|Tangen sudut (radian).
ASIN|mt|ASIN(angka)|Arkus sinus.
ACOS|mt|ACOS(angka)|Arkus kosinus.
ATAN|mt|ATAN(angka)|Arkus tangen.
DEGREES|mt|DEGREES(sudut)|Mengubah radian menjadi derajat.
RADIANS|mt|RADIANS(sudut)|Mengubah derajat menjadi radian.
RAND|mt|RAND()|Angka acak antara 0 dan 1.
RANDBETWEEN|mt|RANDBETWEEN(bawah; atas)|Bilangan bulat acak di antara dua angka.
AVERAGE|st|AVERAGE(angka1; [angka2]; …)|Rata-rata dari argumen.
AVERAGEIF|st|AVERAGEIF(rentang; kriteria; [rentang_rata])|Rata-rata sel yang memenuhi kriteria.
AVERAGEIFS|st|AVERAGEIFS(rentang_rata; rentang_kriteria1; kriteria1; …)|Rata-rata sel yang memenuhi beberapa kriteria.
COUNT|st|COUNT(nilai1; [nilai2]; …)|Menghitung sel yang berisi angka.
COUNTA|st|COUNTA(nilai1; [nilai2]; …)|Menghitung sel yang tidak kosong.
COUNTBLANK|st|COUNTBLANK(rentang)|Menghitung sel kosong.
COUNTIF|st|COUNTIF(rentang; kriteria)|Menghitung sel yang memenuhi kriteria.
COUNTIFS|st|COUNTIFS(rentang1; kriteria1; …)|Menghitung sel yang memenuhi beberapa kriteria.
MAX|st|MAX(angka1; [angka2]; …)|Nilai terbesar.
MIN|st|MIN(angka1; [angka2]; …)|Nilai terkecil.
MAXIFS|st|MAXIFS(rentang_maks; rentang_kriteria1; kriteria1; …)|Nilai terbesar dengan kriteria.
MINIFS|st|MINIFS(rentang_min; rentang_kriteria1; kriteria1; …)|Nilai terkecil dengan kriteria.
MEDIAN|st|MEDIAN(angka1; [angka2]; …)|Nilai tengah.
LARGE|st|LARGE(larik; k)|Nilai terbesar ke-k.
SMALL|st|SMALL(larik; k)|Nilai terkecil ke-k.
RANK|st|RANK(angka; ref; [urutan])|Peringkat angka dalam daftar.
STDEV|st|STDEV(angka1; …)|Simpangan baku sampel.
STDEV.S|st|STDEV.S(angka1; …)|Simpangan baku sampel.
STDEV.P|st|STDEV.P(angka1; …)|Simpangan baku populasi.
VAR|st|VAR(angka1; …)|Varians sampel.
VAR.S|st|VAR.S(angka1; …)|Varians sampel.
VAR.P|st|VAR.P(angka1; …)|Varians populasi.
IF|log|IF(uji_logika; [nilai_jika_benar]; [nilai_jika_salah])|Memeriksa kondisi dan memberi satu nilai bila benar, nilai lain bila salah.
IFS|log|IFS(uji1; nilai1; [uji2; nilai2]; …)|Memeriksa beberapa kondisi dan memberi nilai untuk yang pertama benar.
IFERROR|log|IFERROR(nilai; nilai_jika_galat)|Memberi nilai lain bila rumus menghasilkan galat.
IFNA|log|IFNA(nilai; nilai_jika_na)|Memberi nilai lain bila hasilnya #N/A.
AND|log|AND(logika1; [logika2]; …)|BENAR jika semua argumen benar.
OR|log|OR(logika1; [logika2]; …)|BENAR jika salah satu argumen benar.
XOR|log|XOR(logika1; [logika2]; …)|XOR logika dari semua argumen.
NOT|log|NOT(logika)|Membalik BENAR/SALAH.
SWITCH|log|SWITCH(ekspresi; nilai1; hasil1; …; [default])|Mencocokkan ekspresi dengan daftar nilai.
TRUE|log|TRUE()|Mengembalikan BENAR.
FALSE|log|FALSE()|Mengembalikan SALAH.
LEN|txt|LEN(teks)|Jumlah karakter dalam teks.
LEFT|txt|LEFT(teks; [jumlah])|Karakter paling kiri dari teks.
RIGHT|txt|RIGHT(teks; [jumlah])|Karakter paling kanan dari teks.
MID|txt|MID(teks; awal; jumlah)|Karakter dari tengah teks.
UPPER|txt|UPPER(teks)|Mengubah teks menjadi huruf besar.
LOWER|txt|LOWER(teks)|Mengubah teks menjadi huruf kecil.
PROPER|txt|PROPER(teks)|Huruf pertama tiap kata menjadi kapital.
TRIM|txt|TRIM(teks)|Menghapus spasi berlebih.
CLEAN|txt|CLEAN(teks)|Menghapus karakter yang tidak tercetak.
CONCATENATE|txt|CONCATENATE(teks1; [teks2]; …)|Menggabungkan beberapa teks.
CONCAT|txt|CONCAT(teks1; [teks2]; …)|Menggabungkan teks dari beberapa rentang.
TEXTJOIN|txt|TEXTJOIN(pemisah; abaikan_kosong; teks1; …)|Menggabungkan teks dengan pemisah.
SUBSTITUTE|txt|SUBSTITUTE(teks; lama; baru; [instans])|Mengganti teks tertentu.
REPLACE|txt|REPLACE(teks; awal; jumlah; baru)|Mengganti sebagian teks berdasarkan posisi.
FIND|txt|FIND(cari; dalam_teks; [awal])|Posisi teks (peka huruf besar/kecil).
SEARCH|txt|SEARCH(cari; dalam_teks; [awal])|Posisi teks (tidak peka huruf).
EXACT|txt|EXACT(teks1; teks2)|Memeriksa dua teks identik.
REPT|txt|REPT(teks; kali)|Mengulang teks beberapa kali.
TEXT|txt|TEXT(nilai; format)|Mengubah nilai menjadi teks dengan format tertentu.
VALUE|txt|VALUE(teks)|Mengubah teks menjadi angka.
CODE|txt|CODE(teks)|Kode karakter pertama.
CHAR|txt|CHAR(angka)|Karakter dari kode.
T|txt|T(nilai)|Mengembalikan teks bila nilai berupa teks.
TODAY|dt|TODAY()|Tanggal hari ini.
NOW|dt|NOW()|Tanggal dan waktu saat ini.
DATE|dt|DATE(tahun; bulan; hari)|Membuat tanggal dari tahun, bulan, hari.
TIME|dt|TIME(jam; menit; detik)|Membuat waktu.
YEAR|dt|YEAR(tanggal)|Tahun dari tanggal.
MONTH|dt|MONTH(tanggal)|Bulan dari tanggal.
DAY|dt|DAY(tanggal)|Hari dari tanggal.
HOUR|dt|HOUR(waktu)|Jam dari waktu.
MINUTE|dt|MINUTE(waktu)|Menit dari waktu.
SECOND|dt|SECOND(waktu)|Detik dari waktu.
WEEKDAY|dt|WEEKDAY(tanggal; [tipe])|Hari dalam seminggu (angka).
DAYS|dt|DAYS(tanggal_akhir; tanggal_awal)|Selisih hari antar dua tanggal.
EDATE|dt|EDATE(tanggal; bulan)|Tanggal beberapa bulan sebelum/sesudah.
EOMONTH|dt|EOMONTH(tanggal; bulan)|Hari terakhir bulan.
DATEDIF|dt|DATEDIF(awal; akhir; unit)|Selisih dua tanggal dalam "Y", "M", atau "D".
VLOOKUP|lk|VLOOKUP(nilai; tabel; no_kolom; [cocok_perkiraan])|Mencari nilai di kolom pertama tabel dan mengambil dari kolom lain.
HLOOKUP|lk|HLOOKUP(nilai; tabel; no_baris; [cocok_perkiraan])|Mencari nilai di baris pertama tabel dan mengambil dari baris lain.
XLOOKUP|lk|XLOOKUP(nilai; rentang_cari; rentang_hasil; [jika_tak_ada])|Mencari di satu rentang dan mengembalikan hasil dari rentang lain.
LOOKUP|lk|LOOKUP(nilai; vektor_cari; [vektor_hasil])|Mencari nilai dalam satu baris atau kolom.
INDEX|lk|INDEX(larik; no_baris; [no_kolom])|Mengambil nilai pada posisi tertentu.
MATCH|lk|MATCH(nilai; larik; [tipe_cocok])|Posisi relatif nilai dalam larik.
CHOOSE|lk|CHOOSE(no_indeks; nilai1; [nilai2]; …)|Memilih nilai dari daftar.
ROW|lk|ROW([ref])|Nomor baris referensi.
COLUMN|lk|COLUMN([ref])|Nomor kolom referensi.
ROWS|lk|ROWS(larik)|Jumlah baris dalam referensi.
COLUMNS|lk|COLUMNS(larik)|Jumlah kolom dalam referensi.
ISNUMBER|inf|ISNUMBER(nilai)|BENAR jika nilai berupa angka.
ISTEXT|inf|ISTEXT(nilai)|BENAR jika nilai berupa teks.
ISBLANK|inf|ISBLANK(nilai)|BENAR jika sel kosong.
ISLOGICAL|inf|ISLOGICAL(nilai)|BENAR jika nilai logika.
ISERROR|inf|ISERROR(nilai)|BENAR jika nilai adalah galat.
ISERR|inf|ISERR(nilai)|BENAR jika galat selain #N/A.
ISNA|inf|ISNA(nilai)|BENAR jika nilai #N/A.
NA|inf|NA()|Mengembalikan galat #N/A.
N|inf|N(nilai)|Mengubah nilai menjadi angka.
PMT|fin|PMT(bunga; nper; pv; [fv]; [tipe])|Pembayaran berkala pinjaman dengan bunga tetap.
IPMT|fin|IPMT(bunga; per; nper; pv; [fv])|Bagian bunga dari pembayaran pada periode tertentu.
PPMT|fin|PPMT(bunga; per; nper; pv; [fv])|Bagian pokok dari pembayaran pada periode tertentu.
FV|fin|FV(bunga; nper; pmt; [pv]; [tipe])|Nilai masa depan investasi.
PV|fin|PV(bunga; nper; pmt; [fv]; [tipe])|Nilai sekarang investasi.
NPV|fin|NPV(bunga; nilai1; [nilai2]; …)|Nilai bersih sekarang dari arus kas.
IRR|fin|IRR(nilai; [tebakan])|Tingkat pengembalian internal.
NPER|fin|NPER(bunga; pmt; pv; [fv]; [tipe])|Jumlah periode investasi.
RATE|fin|RATE(nper; pmt; pv; [fv]; [tipe]; [tebakan])|Suku bunga per periode.
SLN|fin|SLN(biaya; sisa; umur)|Penyusutan garis lurus per periode.`;
const FNI={};FNL.split('\n').forEach(l=>{const p=l.split('|');FNI[p[0]]={c:p[1],syn:p[2],d:p[3]};});
const FNCAT=[['all','Semua'],['rec','Baru Digunakan'],['fin','Keuangan'],['log','Logika'],['txt','Teks'],['dt','Tanggal & Waktu'],['lk','Pencarian & Referensi'],['mt','Matematika & Trigonometri'],['st','Statistik'],['inf','Informasi']];
const BC=(ic,tip,a,lab,col)=>`<button class="b big" title="${tip}" data-a="${a}"><em style="color:${col};font-weight:700">${ic}</em><span>${lab}<br>▾</span></button>`;
RIB.formulas=
 G('Pustaka Fungsi',B('<i style="font-family:Georgia,serif">fx</i>','Sisipkan Fungsi','fxdlg','big','Sisipkan<br>Fungsi')+
   B('Σ','Jumlah Otomatis','m:sum','big','Jumlah<br>Otomatis',1)+BC('★','Baru Digunakan','m:cRec','Baru<br>Digunakan','#2f5496')+BC('$','Keuangan','m:cfin','Keuangan','#4a8a6e')+BC('?','Logika','m:clog','Logika','#8a4aa6')+
   BC('A','Teks','m:ctxt','Teks','#e69b2c')+BC('▦','Tanggal & Waktu','m:cdt','Tanggal &<br>Waktu','#d2562b')+BC('🔍','Pencarian & Referensi','m:clk','Pencarian &<br>Referensi','#2f5496')+BC('θ','Matematika & Trigonometri','m:cmt','Matematika<br>& Trig','#4a8a6e')+BC('…','Fungsi Lainnya','m:cmore','Fungsi<br>Lainnya','#d2562b'))+
 G('Nama Terdefinisi',B('▤','Pengelola Nama','namemgr','big','Pengelola<br>Nama')+`<div class="rows t3">${B('▭','Tentukan Nama','m:defn','rw','Tentukan Nama',1)}${B('ƒ','Gunakan di Rumus','m:usef','rw','Gunakan di Rumus',1)}${B('▦','Buat dari Pilihan','namesel','rw','Buat dari Pilihan')}</div>`)+
 G('Audit Rumus',`<div class="rows t3">${B('⇢','Telusuri Preseden','trace:prec','rw','Telusuri Preseden')}${B('⇠','Telusuri Dependen','trace:dep','rw','Telusuri Dependen')}${B('✖','Hapus Panah','m:rmarr','rw','Hapus Panah',1)}</div><div class="rows t3">${B('ƒ','Tampilkan Rumus','showf','rw','Tampilkan Rumus')}${B('⚠','Pemeriksaan Galat','m:errchk','rw','Pemeriksaan Galat',1)}${B('ƒ','Evaluasi Rumus','evalf','rw','Evaluasi Rumus')}</div>`+B('👓','Jendela Pantau','watch','big','Jendela<br>Pantau'))+
 G('Perhitungan',B('▦','Opsi Perhitungan','m:calcopt','big','Opsi<br>Perhitungan',1)+`<div class="rows t3">${B('▦','Hitung Sekarang (F9)','calcnow','rw','Hitung Sekarang')}${B('▦','Hitung Lembar (Shift+F9)','calcsheet','rw','Hitung Lembar')}</div>`);
RIB.data=
 G('Ambil Data Eksternal',BD('🗄','Dari Access','Dari<br>Access')+B('🌐','Dari Web','fromweb','big','Dari<br>Web')+B('📄','Dari Teks','fromtext','big','Dari<br>Teks')+BD('⚙','Dari Sumber Lain','Dari Sumber<br>Lain ▾')+BD('🔗','Koneksi yang Ada','Koneksi<br>yang Ada'))+
 G('Koneksi',B('⟳','Segarkan Semua','m:refresh','big','Segarkan<br>Semua',1)+`<div class="rows t3">${B('⛓','Koneksi','conns','rw','Koneksi')}<button class="b rw" disabled title="Properti (belum tersedia)"><em>▤</em><span>Properti</span></button><button class="b rw" disabled title="Edit Tautan (belum tersedia)"><em>🔗</em><span>Edit Tautan</span></button></div>`)+
 G('Urutkan & Filter',`<div class="rows">${B('<small style="line-height:1;font-size:10px">A↓<br>Z</small>','Urutkan A ke Z','sortaz')}${B('<small style="line-height:1;font-size:10px">Z↓<br>A</small>','Urutkan Z ke A','sortza')}</div>`+B('⇅','Urutkan','sortdlg','big','Urutkan')+B('⧩','Filter','filter','big','Filter')+`<div class="rows t3">${B('✖','Bersihkan','fclear','rw','Bersihkan')}${B('⟳','Terapkan Ulang','freapply','rw','Terapkan Ulang')}${B('⧩','Lanjutan','fadv','rw','Lanjutan')}</div>`)+
 G('Alat Data',B('▥','Teks ke Kolom','t2c','big','Teks ke<br>Kolom')+B('⚡','Isi Cepat (Ctrl+E)','flash','big','Isi<br>Cepat')+B('⊟','Hapus Duplikat','rmdup','big','Hapus<br>Duplikat')+B('✔','Validasi Data','m:dval','big','Validasi<br>Data',1)+B('⇥','Konsolidasi','consol','big','Konsolidasi')+B('?','Analisis Bagaimana-Jika','m:whatif','big','Analisis<br>Bagaimana-Jika',1)+BD('⇄','Hubungan','Hubungan'))+
 G('Kerangka',B('▣','Grup','m:grpm','big','Grup',1)+B('▢','Pisahkan Grup','m:ungrpm','big','Pisahkan<br>Grup',1)+B('Σ','Subtotal','subtotal','big','Subtotal')+`<div class="rows t3">${B('＋','Tampilkan Detail','showdet','rw','Tampilkan Detail')}${B('－','Sembunyikan Detail','hidedet','rw','Sembunyikan Detail')}</div>`);
RIB.review=
 G('Pemeriksaan',BD('ABC','Ejaan')+BD('🔎','Riset')+BD('📖','Tesaurus'))+
 G('Bahasa',BD('🌐','Terjemahkan'))+
 G('Komentar',B('🗨','Komentar Baru','cmnew','big','Komentar<br>Baru')+B('✖','Hapus Komentar','cmdel','big','Hapus')+B('◀','Komentar Sebelumnya','cmprev','big','Sebelumnya')+B('▶','Komentar Berikutnya','cmnext','big','Berikutnya')+
   `<div class="rows t3">${B('🗨','Tampilkan/Sembunyikan Komentar','cmshow','rw','Tampilkan/Sembunyikan Komentar')}${B('🗨','Tampilkan Semua Komentar','cmall','rw','Tampilkan Semua Komentar')}<button class="b rw" disabled title="Tampilkan Tinta (belum tersedia)"><em>✎</em><span>Tampilkan Tinta</span></button></div>`)+
 G('Perubahan',B('🔒','Lindungi Lembar','protsheet','big','Lindungi<br>Lembar')+B('🔒','Lindungi Buku Kerja','protwb','big','Lindungi<br>Buku Kerja')+BD('⇄','Bagikan Buku Kerja','Bagikan<br>Buku Kerja')+
   `<div class="rows t3"><button class="b rw" disabled title="Lindungi dan Bagikan Buku Kerja (belum tersedia)"><em>🔒</em><span>Lindungi dan Bagikan Buku Kerja</span></button>${B('▦','Izinkan Pengguna Mengedit Rentang','allowr','rw','Izinkan Pengguna Mengedit Rentang')}${B('✎','Lacak Perubahan','m:track','rw','Lacak Perubahan',1)}</div>`);
RIB.view=
 G('Tampilan Buku Kerja',B('▦','Normal','view:normal','big','Normal')+B('▥','Pratinjau Sela Halaman','view:pbp','big','Pratinjau<br>Sela Halaman')+BD('▭','Tata Letak Halaman','Tata Letak<br>Halaman')+B('👁','Tampilan Kustom','cviews','big','Tampilan<br>Kustom'))+
 G('Tampilkan',`<div class="rows t3" style="min-width:84px"><div class="row"><label style="opacity:.45"><input type="checkbox" disabled> Penggaris</label></div><div class="row"><label><input type="checkbox" data-s="gview"> Garis Kisi</label></div></div><div class="rows t3" style="min-width:96px"><div class="row"><label><input type="checkbox" data-s="fbar"> Bilah Rumus</label></div><div class="row"><label><input type="checkbox" data-s="hview"> Tajuk</label></div></div>`)+
 G('Zoom',B('🔍','Zoom','zoomdlg','big','Zoom')+B('100','Zoom 100%','zoom100','big','100%')+B('⊡','Zoom ke Seleksi','zoomsel','big','Zoom ke<br>Seleksi'))+
 G('Jendela',BD('▤','Jendela Baru','Jendela<br>Baru')+BD('▦','Susun Semua','Susun<br>Semua')+B('❄','Bekukan Panel','m:freeze','big','Bekukan<br>Panel',1)+
   `<div class="rows t3"><button class="b rw" disabled title="Pisahkan (belum tersedia)"><em>▭</em><span>Pisahkan</span></button><button class="b rw" disabled title="Sembunyikan (belum tersedia)"><em>▭</em><span>Sembunyikan</span></button><button class="b rw" disabled title="Tampilkan (belum tersedia)"><em>▭</em><span>Tampilkan</span></button></div><div class="rows t3"><button class="b rw" disabled title="Tampilan Berdampingan (belum tersedia)"><em>▥</em><span>Tampilan Berdampingan</span></button><button class="b rw" disabled title="Gulir Sinkron (belum tersedia)"><em>⇅</em><span>Gulir Sinkron</span></button><button class="b rw" disabled title="Setel Ulang Posisi Jendela (belum tersedia)"><em>▦</em><span>Setel Ulang Posisi Jendela</span></button></div>`+B('▤','Beralih Jendela','m:switchw','big','Beralih<br>Jendela',1))+
 G('Makro',BD('▤','Makro','Makro ▾'));
const SOON='<div class="soon">Menu ini akan dibuat pada tahap berikutnya.</div>';

const PAPER={A4:[210,297],A3:[297,420],A5:[148,210],Letter:[215.9,279.4],Legal:[215.9,355.6]};
const PAPERN={A4:'A4 (21 × 29,7 cm)',A3:'A3 (29,7 × 42 cm)',A5:'A5 (14,8 × 21 cm)',Letter:'Letter (21,6 × 27,9 cm)',Legal:'Legal (21,6 × 35,6 cm)'};
const newPage=()=>({orient:'portrait',paper:'A4',m:{t:.75,b:.75,l:.7,r:.7},area:null,tRows:null,tCols:null,fitW:0,fitH:0,scale:100,pg:false,ph:false});
const THEMES=[
 {n:'Office',c:['#ffffff','#000000','#e7e6e6','#44546a','#4472c4','#ed7d31','#a5a5a5','#ffc000','#5b9bd5','#70ad47'],f:'Calibri'},
 {n:'Skala Abu-abu',c:['#ffffff','#000000','#e7e6e6','#262626','#595959','#7f7f7f','#a5a5a5','#bfbfbf','#d9d9d9','#3f3f3f'],f:'Arial'},
 {n:'Biru',c:['#ffffff','#000000','#dfe9f5','#17406d','#0f6fc6','#009dd9','#0bd0d9','#10cf9b','#7cca62','#a5c249'],f:'Segoe UI'},
 {n:'Hijau',c:['#ffffff','#000000','#e6efd9','#2c4a2a','#54a021','#e6b91e','#e76618','#c42f1a','#918655','#4a6f4e'],f:'Verdana'},
 {n:'Oranye',c:['#ffffff','#000000','#fbe9d5','#5a3a1c','#e48312','#bd582c','#865640','#9b8357','#c2bc80','#94a088'],f:'Georgia'},
 {n:'Klasik',c:['#ffffff','#000000','#eeece1','#1f497d','#4f81bd','#c0504d','#9bbb59','#8064a2','#4bacc6','#f79646'],f:'Times New Roman'}
];
const shade=(h,t)=>{const n=parseInt(h.slice(1),16),f=x=>Math.round(x*(1-t)).toString(16).padStart(2,'0');return '#'+f(n>>16)+f((n>>8)&255)+f(n&255);};
const buildPal=cols=>{const col=h=>h==='#ffffff'?['#f2f2f2','#d9d9d9','#bfbfbf','#a5a5a5','#7f7f7f']:h==='#000000'?['#7f7f7f','#595959','#3f3f3f','#262626','#0c0c0c']:[mix(h,.8),mix(h,.6),mix(h,.4),shade(h,.25),shade(h,.5)];
  const cs=cols.map(col);return [cols,...[0,1,2,3,4].map(i=>cs.map(c=>c[i]))];};
const BTYPES={thin:['1px','solid'],medium:['2px','solid'],thick:['3px','solid'],dashed:['1px','dashed'],dotted:['1px','dotted'],double:['3px','double'],hair:['1px','dotted']};
const SIDES=['t','r','b','l'];
const hexOk=c=>/^#[0-9a-f]{6}$/i.test(c||'');

function mount(root,opts={}){
  if(!document.getElementById('xl-css')){const s=document.createElement('style');s.id='xl-css';s.textContent=CSS;document.head.appendChild(s);}
  root.innerHTML=`<div class="xl">
    <div class="qat"><div class="l"><span class="w">X</span><button data-a="save" title="Simpan (Ctrl+S)">💾</button><button data-a="undo" title="Urungkan (Ctrl+Z)">↶</button><button data-a="redo" title="Ulangi (Ctrl+Y)">↷</button></div><span class="fn"></span></div>
    <div class="tabs"><button class="tab file" data-a="menu">BERKAS</button>${TABS.map((t,i)=>`<button class="tab${i?'':' on'}" data-tab="${t[0]}">${t[1]}</button>`).join('')}</div>
    ${TABS.map((t,i)=>`<div class="ribbon" data-pane="${t[0]}"${i?' hidden':''}>${RIB[t[0]]||SOON}</div>`).join('')}
    <div class="fbar"><div class="nbw"><input class="nb" value="A1" spellcheck="false" title="Kotak Nama"><button class="fb" data-a="m:nbnames" title="Daftar nama" style="width:16px;font-size:8px">▾</button></div>
      <button class="fb no dis" data-a="editcancel" title="Batal">✕</button><button class="fb ok dis" data-a="editok" title="Masukkan">✓</button><button class="fb" data-a="fxdlg" title="Sisipkan fungsi"><i style="font-family:Georgia,serif">fx</i></button>
      <input class="fx" spellcheck="false" autocomplete="off"></div>
    <div class="gwrap"><div class="corner" title="Pilih semua"></div><div class="colh"></div><div class="rowh"></div>
      <div class="vp" tabindex="0"><div class="stage"></div><div class="sizer"></div></div><textarea class="ed" hidden spellcheck="false" style="position:absolute"></textarea></div>
    <div class="shbar"><div class="shnav"><button data-a="shprev">◀</button><button data-a="shnext">▶</button></div><div class="shs"></div><button class="shadd" data-a="addsheet" title="Lembar baru">+</button></div>
    <div class="stbar"><div class="s"><span class="mode">SIAP</span><span class="calcs"></span></div><div class="s"><span class="sum"></span></div>
      <div class="s"><span><button class="on" data-a="view:normal" title="Normal">▦</button><button data-a="soonv" title="Tata Letak Halaman">▭</button><button data-a="view:pbp" title="Pratinjau Sela Halaman">▥</button></span>
      <span><button data-a="zoomd:-10">−</button><input type="range" min="20" max="200" step="10" value="100" class="zr"><button data-a="zoomd:10">+</button> <span class="zv" style="display:inline-block;min-width:34px;text-align:right">100%</span></span></div></div>
    <div class="find" hidden><span>Cari:</span><input type="text" class="f1"><span>Ganti:</span><input type="text" class="f2">
      <div class="fo"><label><input type="checkbox" class="fcase"> Cocokkan huruf besar/kecil</label><label><input type="checkbox" class="fwhole"> Seluruh isi sel</label></div>
      <div class="fb"><button data-a="fnext">Cari Berikutnya</button><button data-a="frep">Ganti</button><button data-a="frall">Ganti Semua</button><button data-a="fclose">Tutup</button></div></div>
    <div class="menu" hidden>
      <button data-a="new">📄  Buku kerja kosong baru</button><button data-a="open">📂  Buka…</button>
      <button data-a="save">💾  Simpan sebagai .xlsx</button><button data-a="savecsv">📑  Simpan lembar sebagai .csv</button><button data-a="print">🖨  Cetak / simpan PDF</button>
      <hr><button data-a="home">⬅  Kembali ke menu utama</button></div>
    <input type="file" class="fin" accept=".xlsx,.csv,.tsv,.txt" hidden><input type="file" class="imgin" accept="image/*" hidden><input type="file" class="txtin" accept=".txt,.csv,.tsv,.prn" hidden><input type="file" class="bgin" accept="image/*" hidden>
  </div>`;
  const $=s=>root.querySelector(s),$$=s=>[...root.querySelectorAll(s)];
  const X=$('.xl'),vp=$('.vp'),stage=$('.stage'),sizer=$('.sizer'),colh=$('.colh'),rowh=$('.rowh'),menu=$('.menu'),fin=$('.fin');
  const cornerEl=$('.corner'),nb=$('.nb'),fxi=$('.fx'),fd=$('.find');

  /* ===== model ===== */
  const wb={sheets:[],active:0,ver:0,circ:false,font:'Calibri',theme:0,names:[],conns:[],track:false,changes:[],cviews:[],viewMode:'normal',fbar:true,trackPause:false,protStruct:null,manual:false,stale:false,recent:['SUM','AVERAGE','IF','COUNT','MAX','MIN','VLOOKUP','TODAY','ROUND','CONCATENATE'],watch:[]};
  let cur=null,Z=1,name='Buku1.xlsx',showGrid=true,showHead=true;
  let ed=null,clip=null,pointing=null,painting=null,tabStart=null,selObj=null,selSet=new Set(),selpEl=null,watchEl=null,watchSel=-1,gutR=0,gutC=0,txStruct=false;
  let undoS=[],redoS=[],dirty=false;
  const newSheet=n=>({name:n,cells:new Map(),cw:new Axis(DW,MAXC),rh:new Axis(DH,MAXR),merges:[],cf:[],filter:null,tab:null,objs:[],spark:[],freeze:null,cm:[],prot:null,allow:[],showAllCm:false,og:{rows:[],cols:[],hr:new Set(),hc:new Set()},dv:[],adv:null,circ:false,grid:true,head:true,bg:null,page:newPage(),
    sel:{ar:0,ac:0,fr:0,fc:0},sx:0,sy:0});
  const uniqueName=(base,except)=>{let n=base,i=1;while(wb.sheets.some(s=>s!==except&&s.name.toLowerCase()===n.toLowerCase()))n=base.replace(/\d+$/,'')+(++i);return n;};
  const DF=()=>wb.font||'Calibri';
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
  const recalcAll=()=>{wb.ver++;if(wb.manual){wb.stale=true;updCalcLabel();return;}recalc(wb);};
  const updCalcLabel=()=>{const e=$('.calcs');if(e)e.textContent=wb.manual&&wb.stale?'Hitung':'';};
  function shotSheet(s){return {name:s.name,cells:new Map(s.cells),cw:s.cw.clone(),rh:s.rh.clone(),merges:s.merges.map(m=>({...m})),cf:[...s.cf],filter:s.filter?{...s.filter,crit:{...s.filter.crit}}:null,tab:s.tab,freeze:s.freeze?{...s.freeze}:null,cm:(s.cm||[]).map(x=>({...x})),prot:s.prot?JSON.parse(JSON.stringify(s.prot)):null,allow:(s.allow||[]).map(x=>({...x})),showAllCm:s.showAllCm,og:s.og?{rows:s.og.rows.map(g=>({...g})),cols:s.og.cols.map(g=>({...g})),hr:new Set(s.og.hr),hc:new Set(s.og.hc)}:{rows:[],cols:[],hr:new Set(),hc:new Set()},dv:(s.dv||[]).map(d=>({...d})),adv:s.adv?new Set(s.adv):null,circ:s.circ,grid:s.grid,showF:s.showF,head:s.head,bg:s.bg,page:JSON.parse(JSON.stringify(s.page||newPage())),objs:(s.objs||[]).map(o=>({...o,src:o.src&&typeof o.src==='object'?{...o.src}:o.src})),spark:(s.spark||[]).map(o=>({...o,src:{...o.src}})),sel:{...s.sel}};}
  const snapshot=()=>({sheets:wb.sheets.map(shotSheet),active:wb.active,names:wb.names.map(n=>({...n,ast:undefined,depth:0}))});
  function restore(sn){
    wb.sheets=sn.sheets.map(shotSheet);if(sn.names)wb.names=sn.names.map(n=>({...n,ast:undefined}));wb.active=Math.min(sn.active,wb.sheets.length-1);
    wb.sheets.forEach(s=>{s.sx=0;s.sy=0;});
    cur=wb.sheets[wb.active];recalcAll();drawTabs();sync();
  }
  function tx(fn,quiet){ /* transaksi: simpan keadaan sebelum perubahan untuk undo */
    const s=snapshot();
    try{fn();}catch(e){console.error(e);toast('Terjadi kesalahan: '+e.message);return false;}
    if(!txGuard(s))return false;
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
    tabStart=null;selObj=null;selSet=new Set();const s=cur.sel;s.ar=ar;s.ac=ac;s.fr=fr===undefined?ar:fr;s.fc=fc===undefined?ac:fc;
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
    const x1=cx(c),x2=cx(c+1),y1=ry(r),y2=ry(r+1),fz=frzOf();
    if(c>=fz.c){if(x1<vp.scrollLeft+fz.W)vp.scrollLeft=Math.max(0,x1-fz.W);else if(x2>vp.scrollLeft+W-17)vp.scrollLeft=Math.min(x1-fz.W,x2-W+17);}
    if(r>=fz.r){if(y1<vp.scrollTop+fz.H)vp.scrollTop=Math.max(0,y1-fz.H);else if(y2>vp.scrollTop+H-17)vp.scrollTop=Math.min(y1-fz.H,y2-H+17);}
  }

  /* ===== geometri ===== */
  const cx=c=>Math.round(cur.cw.pos(c)*Z),ry=r=>Math.round(cur.rh.pos(r)*Z);
  const cwid=c=>cx(c+1)-cx(c),rhei=r=>ry(r+1)-ry(r);
  const rectOf=(r1,c1,r2,c2)=>({x:cx(c1),y:ry(r1),w:cx(c2+1)-cx(c1),h:ry(r2+1)-ry(r1)});
  const mctx=document.createElement('canvas').getContext('2d');
  const tw=(t,st)=>{mctx.font=`${st.i?'italic ':''}${st.b?'bold ':''}${(st.sz||11)*4/3*Z}px ${st.font||DF()},Carlito,sans-serif`;return mctx.measureText(t).width;};
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
        const sz=(st.sz||11)*4/3;m.style.font=`${st.i?'italic ':''}${st.b?'bold ':''}${sz}px ${st.font||DF()},Carlito,sans-serif`;m.style.lineHeight='1.2';
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
    const hh=HH;
    const fz=frzOf(),FW=fz.W,FH=fz.H;
    const mapX=x=>x<FW?x:x-sl,mapY=y=>y<FH?y:y-st0;
    const mapRect=q=>{const x1=mapX(q.x),y1=mapY(q.y);let x2=mapX(q.x+q.w),y2=mapY(q.y+q.h);if(q.x+q.w<=FW)x2=q.x+q.w;if(q.y+q.h<=FH)y2=q.y+q.h;return {x:x1,y:y1,w:Math.max(0,x2-x1),h:Math.max(0,y2-y1)};};
    let c0=cur.cw.at((sl+FW)/Z),r0=cur.rh.at((st0+FH)/Z);
    let c1=c0,r1=r0;while(c1<MAXC-1&&cx(c1)-sl<W)c1++;while(r1<MAXR-1&&ry(r1)-st0<H)r1++;
    const colList=[],rowList=[];
    for(let c=c0;c<=c1;c++)colList.push(c);for(let c=0;c<fz.c;c++)colList.push(c);
    for(let r=r0;r<=r1;r++)rowList.push(r);for(let r=0;r<fz.r;r++)rowList.push(r);
    const content=(sl,st0,W,H,main)=>{
    let c0=cur.cw.at(sl/Z),r0=cur.rh.at(st0/Z);
    let c1=c0,r1=r0;while(c1<MAXC-1&&cx(c1)-sl<W)c1++;while(r1<MAXR-1&&ry(r1)-st0<H)r1++;
    const cs=Math.max(0,c0-1),rs=Math.max(0,r0-1);
    let html='';
    /* garis kisi */
    if(cur.grid!==false){
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
      if(cur.showF&&cell&&cell.f!==undefined&&cell.f!==null)v='='+cell.f;
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
    for(const cm of cur.cm){
      if(cm.r<r0||cm.r>r1||cm.c<c0||cm.c>c1)continue;const w=cwid(cm.c),h=rhei(cm.r);if(w<=0||h<=0)continue;
      const x=cx(cm.c)-sl,y=ry(cm.r)-st0;
      html+=`<div class="cmt" style="left:${x+w-6}px;top:${y}px"></div>`;
      if(cm.show||cur.showAllCm){const bw=Math.round(cm.w*Z),bx=x+w+16,by=y-4;
        html+=`<svg style="position:absolute;left:0;top:0;pointer-events:none;z-index:4" width="${W}" height="${H}"><line x1="${x+w-2}" y1="${y+2}" x2="${bx}" y2="${by+10}" stroke="#d33" stroke-width="1"/></svg><div class="cbox" style="left:${bx}px;top:${by}px;width:${bw}px;min-height:${Math.round(cm.h*Z)}px"><b>${esc(cm.author)}:</b><br>${esc(cm.text)}</div>`;}
    }
    if(wb.changes.length)for(const ch of wb.changes){
      if(ch.sh!==cur.name||ch.r<r0||ch.r>r1||ch.c<c0||ch.c>c1)continue;const w=cwid(ch.c),h=rhei(ch.r);if(w<=0||h<=0)continue;
      html+=`<div class="trk" style="left:${cx(ch.c)-sl}px;top:${ry(ch.r)-st0}px;width:${w-1}px;height:${h-1}px"></div>`;
    }
    for(const sp of cur.spark){
      if(sp.r<r0||sp.r>r1||sp.c<c0||sp.c>c1)continue;
      const w=cwid(sp.c),h=rhei(sp.r);if(w<=0||h<=0)continue;
      html+=`<div class="cb" style="left:${cx(sp.c)-sl}px;top:${ry(sp.r)-st0}px;width:${w}px;height:${h}px">${sparkSvg(sp,w,h)}</div>`;
    }
    if(main)for(const o of cur.objs){
      const x=Math.round(o.x*Z)-sl,y=Math.round(o.y*Z)-st0,w=Math.round(o.w*Z),h=Math.round(o.h*Z);
      if(x+w<-20||y+h<-20||x>W+20||y>H+20)continue;
      html+=objHtml(o,x,y,w,h);
    }
    return html;};
    let html='';
    {const reg=(l,t,w,h,sl2,st2,main,fr)=>w>0&&h>0?`<div class="rgn" style="left:${l}px;top:${t}px;width:${w}px;height:${h}px${fr?';background:#fff':''}">${content(sl2,st2,w,h,main)}</div>`:'';
      if(FW||FH)html+=reg(FW,FH,W-FW,H-FH,sl+FW,st0+FH,true,false)+reg(FW,0,W-FW,FH,sl+FW,0,false,true)+reg(0,FH,FW,H-FH,0,st0+FH,false,true)+reg(0,0,FW,FH,0,0,false,true);
      else html+=reg(0,0,W,H,sl,st0,true,false);}
    /* filter */
    if(cur.filter){const f=cur.filter;
      for(let c=f.c1;c<=f.c2;c++){if(c<c0||c>c1||ry(f.r1)-st0>H||ry(f.r1+1)<st0)continue;
        const x=mapX(cx(c+1))-17,y=mapY(ry(f.r1+1))-17;
        html+=`<button class="fdd${f.crit[c]?' on':''}" data-fc="${c}" style="left:${x}px;top:${y}px">${f.crit[c]?'⧩':'▾'}</button>`;}}
    if(cur.tr){
      if(cur.tr.ver!==wb.ver)computeTrace();
      let ar='';
      const cc=(r,c)=>[mapX(cx(c)+cwid(c)/2),mapY(ry(r)+rhei(r)/2)];
      for(const a of cur.tr.arrows){
        const f=a.from,[x2,y2]=cc(a.to.r,a.to.c);
        const multi=f.r1!==f.r2||f.c1!==f.c2,q=mapRect(rectOf(f.r1,f.c1,f.r2,f.c2));
        if(multi)ar+=`<rect x="${q.x}" y="${q.y}" width="${q.w}" height="${q.h}" fill="none" stroke="#2f5aa6" stroke-width="1.5"/>`;
        const [x1,y1]=multi?[q.x+3,q.y+3]:cc(f.r1,f.c1);
        ar+=`<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="#2f5aa6" stroke-width="1.5" marker-end="url(#xlah)"/><circle cx="${x1}" cy="${y1}" r="3" fill="#2f5aa6"/>`;
      }
      if(ar)html+=`<svg style="position:absolute;left:0;top:0;pointer-events:none;z-index:3" width="${W}" height="${H}"><defs><marker id="xlah" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0 0L8 4L0 8z" fill="#2f5aa6"/></marker></defs>${ar}</svg>`;
    }
    if(wb.viewMode==='pbp')html+=pbpOverlay(W,H,mapX,mapY);
    /* seleksi */
    const g0=rg(),g=expandMerge(g0),s=cur.sel;
    const ac=mergeAt(s.ar,s.ac),ar=ac?{r1:ac.r1,c1:ac.c1,r2:ac.r2,c2:ac.c2}:{r1:s.ar,c1:s.ac,r2:s.ar,c2:s.ac};
    const R=mapRect(rectOf(g.r1,g.c1,g.r2,g.c2)),A=mapRect(rectOf(ar.r1,ar.c1,ar.r2,ar.c2));
    const rx=R.x,ryy=R.y,ax=A.x,ay=A.y;
    const multi=g.r1!==g.r2||g.c1!==g.c2;
    const multiNotMerge=multi&&!(ac&&ac.r1===g.r1&&ac.r2===g.r2&&ac.c1===g.c1&&ac.c2===g.c2);
    if(multiNotMerge){ /* arsiran di sekitar sel aktif */
      const T=(x,y,w,h)=>w>0&&h>0?`<div class="tint" style="left:${x}px;top:${y}px;width:${w}px;height:${h}px"></div>`:'';
      html+=T(rx,ryy,R.w,ay-ryy)+T(rx,ay+A.h,R.w,ryy+R.h-ay-A.h)+T(rx,ay,ax-rx,A.h)+T(ax+A.w,ay,rx+R.w-ax-A.w,A.h);
    }
    html+=`<div class="selb" style="left:${rx-1}px;top:${ryy-1}px;width:${R.w+1}px;height:${R.h+1}px"></div>`;
    if(clip&&clip.sh===cur){const q=mapRect(rectOf(clip.r1,clip.c1,clip.r2,clip.c2));html+=`<div class="cpy" style="left:${q.x-1}px;top:${q.y-1}px;width:${q.w+1}px;height:${q.h+1}px"></div>`;}
    if(pointing){const q=mapRect(rectOf(pointing.r1,pointing.c1,pointing.r2,pointing.c2));html+=`<div class="pnt" style="left:${q.x-1}px;top:${q.y-1}px;width:${q.w+1}px;height:${q.h+1}px"></div>`;}
    {const rule=!ed?dvAt(s.ar,s.ac):null;
      if(rule){
        if(rule.type==='list')html+=`<button class="fdd" data-dvl="1" style="left:${ax+A.w-17}px;top:${ay+A.h-17}px;z-index:5">▾</button>`;
        if(rule.showIn&&(rule.inMsg||rule.inTitle))html+=`<div class="dvtip" style="left:${ax}px;top:${ay+A.h+4}px">${rule.inTitle?'<b>'+esc(rule.inTitle)+'</b><br>':''}${esc(rule.inMsg||'')}</div>`;}
      if(cur.circ&&cur.dv.length){let ci='';
        for(let r=r0;r<=r1;r++)for(let c=c0;c<=c1;c++){const rl=dvAt(r,c);if(!rl)continue;const cl=cellAt(cur,r,c);if(!cl||cl.f!=null)continue;const tx0=rawText(cl);if(tx0===''||!dvCheck(r,c,tx0))continue;
          ci+=`<ellipse cx="${cx(c)+cwid(c)/2-sl}" cy="${ry(r)+rhei(r)/2-st0}" rx="${cwid(c)/2+2}" ry="${rhei(r)/2+1}" fill="none" stroke="#e00" stroke-width="2"/>`;}
        if(ci)html+=`<svg style="position:absolute;left:0;top:0;pointer-events:none;z-index:3" width="${W}" height="${H}">${ci}</svg>`;}}
    if(!ed)html+=`<div class="fh" style="left:${rx+R.w-5}px;top:${ryy+R.h-5}px" data-fill="1"></div>`;
    if(FW)html+=`<div style="position:absolute;left:${FW-1}px;top:0;width:1px;height:${H}px;background:#808080;z-index:3;pointer-events:none"></div>`;
    if(FH)html+=`<div style="position:absolute;top:${FH-1}px;left:0;height:1px;width:${W}px;background:#808080;z-index:3;pointer-events:none"></div>`;
    stage.innerHTML=html;
    if(ed)placeEditor();
    /* tajuk */
    let ch='',rh2='';
    const ogl=cur.og||{rows:[],cols:[]};
    const rgut=14*ogl.rows.reduce((m,x)=>Math.max(m,ogDepth(ogl.rows,x)),0),cgut=14*ogl.cols.reduce((m,x)=>Math.max(m,ogDepth(ogl.cols,x)),0);
    gutR=cur.head===false?0:rgut;gutC=cur.head===false?0:cgut;
    const all=g.r1===0&&g.r2===MAXR-1,allc=g.c1===0&&g.c2===MAXC-1;
    for(const c of colList){const w=cwid(c);const on=c>=g.c1&&c<=g.c2;if(w<=0){ch+=`<div class="hc hd" style="left:${mapX(cx(c))-1}px;width:3px;top:${cgut}px;height:${hh}px"></div>`;continue;}
      ch+=`<div class="hc${on?(all?' a':' s'):''}" data-h="c${c}" style="left:${mapX(cx(c))}px;width:${w}px;top:${cgut}px;height:${hh}px">${colName(c)}</div>`;}
    for(const r of rowList){const h=rhei(r);const on=r>=g.r1&&r<=g.r2;if(h<=0){rh2+=`<div class="hc hd" style="top:${mapY(ry(r))-1}px;height:3px;left:${rgut}px;width:${HW}px"></div>`;continue;}
      rh2+=`<div class="hc${on?(allc?' a':' s'):''}" data-h="r${r}" style="top:${mapY(ry(r))}px;height:${h}px;left:${rgut}px;width:${HW}px;font-size:${Math.max(8,12*Math.min(1,Z))}px">${r+1}</div>`;}
    ogl.rows.forEach((x,i)=>{
      const d=ogDepth(ogl.rows,x),xo=(d-1)*14+3,y1=mapY(ry(x.a)),y2=mapY(ry(x.b+1));if(y2<-30||y1>H+30)return;const sh2=rhei(x.b+1);
      if(y2>y1)rh2+=`<div class="ogl" style="left:${xo+5}px;top:${y1}px;width:1px;height:${y2-y1}px"></div>`;
      rh2+=`<div class="ogl" style="left:${xo+5}px;top:${y2-1}px;width:7px;height:1px"></div>`;
      if(sh2>0)rh2+=`<button class="ogb" data-og="r:${i}" style="left:${xo}px;top:${y2+(sh2-11)/2}px">${x.cl?'+':'−'}</button>`;});
    ogl.cols.forEach((x,i)=>{
      const d=ogDepth(ogl.cols,x),yo=(d-1)*14+3,x1=mapX(cx(x.a)),x2=mapX(cx(x.b+1));if(x2<-30||x1>W+30)return;const sw=cwid(x.b+1);
      if(x2>x1)ch+=`<div class="ogl" style="top:${yo+5}px;left:${x1}px;height:1px;width:${x2-x1}px"></div>`;
      ch+=`<div class="ogl" style="top:${yo+5}px;left:${x2-1}px;height:7px;width:1px"></div>`;
      if(sw>0)ch+=`<button class="ogb" data-og="c:${i}" style="top:${yo}px;left:${x2+(sw-11)/2}px">${x.cl?'+':'−'}</button>`;});
    colh.style.height=(hh+cgut)+'px';rowh.style.width=(HW+rgut)+'px';cornerEl.style.width=(HW+rgut)+'px';cornerEl.style.height=(hh+cgut)+'px';
    {const lv=Math.max(rgut,cgut)/14+1,k=(rgut||cgut)?lv:0;if(cornerEl.dataset.lv!==String(k)){cornerEl.dataset.lv=k;cornerEl.innerHTML=k?Array.from({length:k},(_,i)=>`<button class="ogb" data-og="L:${i+1}" style="left:${i*14+2}px;top:2px">${i+1}</button>`).join(''):'';}}
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
    setSel2('font',st.font||DF());setSel2('size',st.sz||11);
    const k=fmtKey(st.fmt||'General');setSel2('nfmt',k||'custom');
    refreshLayout();
    {const hereCm=cmAt(s.ar,s.ac),sd=(a,v)=>{const b=root.querySelector(`[data-a="${a}"]`);if(b)b.disabled=!!v;};
      sd('cmdel',!hereCm);sd('cmshow',!hereCm);sd('cmprev',!cur.cm.length);sd('cmnext',!cur.cm.length);T('cmshow',hereCm&&hereCm.show);T('cmall',cur.showAllCm);
      const ps=root.querySelector('[data-a=protsheet] span');if(ps)ps.innerHTML=cur.prot?'Cabut Proteksi<br>Lembar':'Lindungi<br>Lembar';
      const pw=root.querySelector('[data-a=protwb] span');if(pw)pw.innerHTML=wb.protStruct?'Cabut Proteksi<br>Buku Kerja':'Lindungi<br>Buku Kerja';}
    const sf=root.querySelector('[data-a=showf]');if(sf)sf.classList.toggle('on',!!cur.showF);
    if(watchEl)renderWatch();
    if(!document.activeElement||document.activeElement!==nb){const nn=nameOfSel();if(nn)nb.value=nn;}
  }
  function refreshLayout(){
    const P=cur.page;
    setSel2('fitw',P.fitW);setSel2('fith',P.fitH);
    const sc=root.querySelector('[data-s=scale]');if(sc){if(document.activeElement!==sc)sc.value=P.scale;sc.disabled=!!(P.fitW||P.fitH);}
    const C=(n,v)=>{root.querySelectorAll(`[data-s=${n}]`).forEach(e=>{e.checked=!!v;});};
    C('gview',cur.grid!==false);C('gprint',P.pg);C('hview',cur.head!==false);C('hprint',P.ph);C('fbar',wb.fbar!==false);
    root.querySelectorAll('[data-a="view:normal"]').forEach(b=>b.classList.toggle('on',wb.viewMode!=='pbp'));
    root.querySelectorAll('[data-a="view:pbp"]').forEach(b=>b.classList.toggle('on',wb.viewMode==='pbp'));
    const bg=root.querySelector('[data-a=bg] span');if(bg)bg.innerHTML=cur.bg?'Hapus Latar<br>Belakang':'Latar<br>Belakang';
    $('.gwrap').classList.toggle('nh',cur.head===false);
    vp.style.backgroundImage=cur.bg?`url(${cur.bg})`:'';vp.style.backgroundAttachment='local';
    if(selpEl)renderSelp();
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
    if(!ed){const m=mergeAt(s.ar,s.ac);const fc=cellAt(cur,m?m.r1:s.ar,m?m.c1:s.ac);fxi.value=cur.prot&&fc&&fc.st&&fc.st.hide&&fc.f!=null?'':rawText(fc);}
    refreshRibbon();updStatus();schedule();
  }
  /* ===== editor sel ===== */
  const setMode=m=>{$('.mode').textContent=m;};
  const isFormulaText=t=>t.length>1&&t[0]==='=';
  function startEdit(r,c,init,fromBar){
    if(ed)return;
    const m=mergeAt(r,c);if(m){r=m.r1;c=m.c1;}
    if(cur.prot&&isLocked(cur,r,c)){toast(protMsg);return;}
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
    const hwid=cur.head===false?0:HW+gutR,hhei=cur.head===false?0:HH+gutC;
    const fe=frzOf();
    const x=hwid+(ed.c<fe.c?cx(ed.c):cx(ed.c)-sl),y=hhei+(ed.r<fe.r?ry(ed.r):ry(ed.r)-sy),w0=cx(cc+1)-cx(ed.c),h0=ry(rr+1)-ry(ed.r);
    const sz=(st.sz||11)*4/3*Z;
    edEl.style.fontSize=sz+'px';edEl.style.fontFamily=(st.font||DF())+',Carlito,sans-serif';
    edEl.style.fontWeight=st.b?'700':'400';edEl.style.fontStyle=st.i?'italic':'normal';edEl.style.color=st.color||'#000';
    const lines=edEl.value.split('\n');
    let mw=0;for(const l of lines)mw=Math.max(mw,tw(l,{i:st.i,b:st.b,sz:st.sz,font:st.font}));
    const maxW=vp.clientWidth+hwid-x;
    const w=Math.max(w0,Math.min(maxW,mw+16)),h=Math.max(h0,lines.length*sz*1.2+4);
    Object.assign(edEl.style,{left:x-1+'px',top:y-1+'px',width:w+2+'px',height:h+2+'px'});
    edEl.style.whiteSpace=w>=maxW?'pre-wrap':'pre';
  }
  function endEdit(){
    ed=null;pointing=null;edEl.hidden=true;edEl.value='';
    $('.fb.no').classList.add('dis');$('.fb.ok').classList.add('dis');setMode('Siap');
    if(document.activeElement!==nb)vp.focus({preventScroll:true});
  }
  function commitEdit(dr,dc,force){
    if(!ed)return true;
    const {r,c}=ed;let text=ed.text;
    if(isFormulaText(text)){let dp=0,qq=false;for(const ch of text){if(ch==='"')qq=!qq;else if(!qq){if(ch==='(')dp++;else if(ch===')')dp--;}}if(dp>0&&!qq)text+=')'.repeat(dp);}
    const startC=tabStart===null?c:tabStart;
    if(!force){const bad=dvCheck(r,c,text);if(bad&&dvReject(bad,dr,dc))return false;}
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
    const f=frzOf();return {x,y,c:cur.cw.at((x<f.W?x:x+vp.scrollLeft)/Z),r:cur.rh.at((y<f.H?y:y+vp.scrollTop)/Z)};};
  let drag=null,dragTimer=0,lastEv=null;
  function extendTo(ev){
    const h=hit(ev);
    const W=vp.clientWidth,H=vp.clientHeight;
    if(drag.kind==='obj'){
      const dx=(ev.clientX-drag.x0)/Z,dy=(ev.clientY-drag.y0)/Z;if(Math.abs(dx)+Math.abs(dy)>2)drag.moved=true;
      const o=drag.o;if(drag.mode==='mv'){for(const q of drag.group){q.o.x=Math.max(0,q.ox+dx);q.o.y=Math.max(0,q.oy+dy);}}else{o.w=Math.max(24,drag.ow+dx);o.h=Math.max(16,drag.oh+dy);}
      schedule();return;}
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
    const dvb=e.target.closest&&e.target.closest('[data-dvl]');
    if(dvb){e.preventDefault();dvDropdown(dvb);return;}
    const fdd=e.target.closest&&e.target.closest('[data-fc]');
    if(fdd){e.preventDefault();filterMenu(+fdd.dataset.fc,fdd);return;}
    if(e.target.dataset&&e.target.dataset.fill){
      e.preventDefault();const g=expandMerge(rg());drag={kind:'fill',src:g,tgt:null};vp.setPointerCapture(e.pointerId);startDragLoop();return;}
    const h=hit(e);
    if(e.pointerType!=='mouse'&&e.button===0){return;} /* sentuhan: ditangani oleh click */
    if(e.button===2){const g=expandMerge(rg());if(!(h.r>=g.r1&&h.r<=g.r2&&h.c>=g.c1&&h.c<=g.c2))select(h.r,h.c);return;}
    const ob=e.target.closest&&e.target.closest('[data-o]');
    if(ob){
      e.preventDefault();if(ed&&!commitEdit())return;vp.focus({preventScroll:true});
      const o=cur.objs.find(q=>q.id===ob.dataset.o);
      if(o){
        if(cur.prot&&!cur.prot.o.objects){selSet=new Set([o.id]);selObj=o.id;sync();return;}
        if(e.shiftKey){if(selSet.has(o.id)){selSet.delete(o.id);selObj=[...selSet].pop()||null;sync();return;}selSet.add(o.id);selObj=o.id;}
        else if(!selSet.has(o.id)){selSet=new Set([o.id]);selObj=o.id;}else selObj=o.id;
        drag={kind:'obj',o,mode:e.target.dataset.orz?'rz':'mv',x0:e.clientX,y0:e.clientY,ox:o.x,oy:o.y,ow:o.w,oh:o.h,group:cur.objs.filter(q=>selSet.has(q.id)).map(q=>({o:q,ox:q.x,oy:q.y})),snap:snapshot(),moved:false};
        vp.setPointerCapture(e.pointerId);startDragLoop();lastEv=e;sync();return;}
    }
    if(e.ctrlKey||e.metaKey){const lc=cellAt(cur,h.r,h.c);if(lc&&lc.st&&lc.st.link){e.preventDefault();openLink(lc.st.link);return;}}
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
    if(d.kind==='obj'&&d.moved){undoS.push(d.snap);if(undoS.length>60)undoS.shift();redoS=[];dirty=true;updUndo();}
    if(d.kind==='point')setMode('Tunjuk');
    if(d.kind==='cell'&&painting&&false){}
    schedule();
  };
  vp.addEventListener('pointerup',endDrag);vp.addEventListener('pointercancel',endDrag);
  vp.addEventListener('click',e=>{
    if(e.pointerType==='mouse'||e.detail>0&&e.pointerType==='')return;
    if(e.target.closest&&e.target.closest('[data-fc]'))return;
    const tob=e.target.closest&&e.target.closest('[data-o]');if(tob){selObj=tob.dataset.o;selSet=new Set([selObj]);sync();return;}
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
    const dob=e.target.closest&&e.target.closest('[data-o]');if(dob){objEdit(dob.dataset.o);return;}
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
    const fzz=frzOf(),base=kind==='c'?ev.clientX-rc.left:ev.clientY-rc.top,fzs=kind==='c'?fzz.W:fzz.H;
    const p=base<fzs?base:base+(kind==='c'?vp.scrollLeft:vp.scrollTop);
    const ax=kind==='c'?cur.cw:cur.rh;
    const i=ax.at(p/Z),edge=Math.round(ax.pos(i+1)*Z);
    let near=null;if(Math.abs(p-edge)<=4)near=i;else{const left=Math.round(ax.pos(i)*Z);if(i>0&&Math.abs(p-left)<=4)near=i-1;}
    return {i,edge:near};
  }
  [['c',colh],['r',rowh]].forEach(([kind,el])=>{
    el.addEventListener('pointermove',e=>{if(rsz)return;const h=hdrHit(e,kind);el.style.cursor=h.edge!==null?(kind==='c'?'col-resize':'row-resize'):'default';});
    el.addEventListener('pointerdown',e=>{
      if(e.button!==0)return;
      if(e.target.closest&&e.target.closest('[data-og]'))return;
      {const rc0=el.getBoundingClientRect();if(kind==='r'&&e.clientX-rc0.left<gutR)return;if(kind==='c'&&e.clientY-rc0.top<gutC)return;}
      if(ed&&!commitEdit())return;
      e.preventDefault();closePop();vp.focus({preventScroll:true});
      const h=hdrHit(e,kind);
      if(h.edge!==null){if(protBlock(kind==='c'?'formatCols':'formatRows'))return;rsz={kind,i:h.edge,start:kind==='c'?e.clientX:e.clientY,orig:(kind==='c'?cur.cw:cur.rh).size(h.edge),snap:snapshot(),moved:false};el.setPointerCapture(e.pointerId);return;}
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
  $('.corner').addEventListener('pointerdown',e=>{if(e.target.closest&&e.target.closest('[data-og]'))return;e.preventDefault();if(ed&&!commitEdit())return;select(0,0,MAXR-1,MAXC-1,true);cur.sel.ar=cur.sel.ar;vp.focus();});

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
      case 'F9':e.preventDefault();e.shiftKey?calcSheet():calcNow();return true;
      case 'Delete':e.preventDefault();if(selObj){deleteObj();return true;}acts.clearc();return true;
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
      if(k==='`'){e.preventDefault();acts.showf();return true;}
      if(l==='e'){e.preventDefault();flashFill();return true;}
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
    for(const k in st){const v=st[k];if(v===undefined||v===null||(v===false&&k!=='lock'))continue;
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
    if(protBlock('formatCells'))return;
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
  function adjNames(kind,at,n,del,target){
    wb.names=wb.names.map(nm=>{
      const nf=rewriteRefs(nm.ref,(a,b,sn)=>{
        if(!sn||sn.toLowerCase()!==target.name.toLowerCase())return [a,b];
        const idx=kind==='r'?'r':'c';
        const fx=x=>{const v=x[idx];if(!del)return v>=at?v+n:v;if(v<at)return v;if(v<at+n)return -1;return v-n;};
        if(!b){const v=fx(a);if(v<0)return null;return [{...a,[idx]:v},null];}
        let lo=a[idx],hi=b[idx];
        if(!del){if(lo>=at)lo+=n;if(hi>=at)hi+=n;}else{const e=at+n-1;lo=lo<at?lo:lo<=e?at:lo-n;hi=hi<at?hi:hi<=e?at-1:hi-n;if(hi<lo)return null;}
        return [{...a,[idx]:lo},{...b,[idx]:hi}];});
      return nf===nm.ref?nm:{...nm,ref:nf,ast:undefined};});
  }
  function adjRefsAll(kind,at,n,del,target){
    adjNames(kind,at,n,del,target);
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
    if(cur.prot){const f=(del?'del':'ins')+(kind==='r'?'Rows':'Cols');if(protBlock(f))return;}
    txStruct=true;try{tx(()=>structRaw(kind,at,n,del));}finally{txStruct=false;}
  }
  function structRaw(kind,at,n,del){
    {
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
      cur.dv=cur.dv.map(d=>adj(d,lo,hi)).filter(Boolean);
      cur.cm=cur.cm.map(x=>kind==='r'?adj(x,'r','r'):adj(x,'c','c')).filter(Boolean);
      {const ol=kind==='r'?cur.og.rows:cur.og.cols;const nl=ol.map(g=>adj(g,'a','b')).filter(Boolean);if(kind==='r')cur.og.rows=nl;else cur.og.cols=nl;applyOutline();}
      if(cur.filter)applyFilter(true);
    }
  }
  function shiftCells(dir,ins){ /* geser sebagian sel (tanpa menyesuaikan rumus lain) */
    if(protBlock(null))return;
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
    if(protBlock(kind==='r'?'formatRows':'formatCols'))return;
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
  function sortRange(g,cols,asc,header){
    const keys=Array.isArray(cols)?cols:[{col:cols,asc}];
    const r1=g.r1+(header?1:0);if(r1>=g.r2)return;
    if(protBlock('sort'))return;
    tx(()=>{
      const rows=[];
      for(let r=r1;r<=g.r2;r++){const line=[];for(let c=g.c1;c<=g.c2;c++)line.push(cellAt(cur,r,c));rows.push({r,line,k:keys.map(q=>cellVal(cur,r,q.col))});}
      const rank=v=>v===null||v===''?3:typeof v==='number'?0:typeof v==='string'?1:2;
      rows.sort((a,b)=>{for(let i=0;i<keys.length;i++){const x=a.k[i],y=b.k[i],rx=rank(x),ry2=rank(y);if(rx!==ry2)return rx-ry2;if(rx===3)continue;
        let c=rx===0?x-y:rx===1?String(x).localeCompare(String(y),'id',{sensitivity:'base'}):(+x)-(+y);if(c)return keys[i].asc?c:-c;}return a.r-b.r;});
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
    if(protBlock('filter'))return;
    if(cur.filter){tx(()=>{cur.filter=null;cur.adv=null;cur.rh.fil=new Set();cur.rh.dirty=true;});return;}
    const g=region();if(g.r1===g.r2&&g.c1===g.c2){toast('Pilih sel di dalam data untuk menerapkan filter.');return;}
    tx(()=>{cur.filter={r1:g.r1,c1:g.c1,r2:g.r2,c2:g.c2,crit:{}};});
  }
  function applyFilter(inTx){
    const f=cur.filter;const hid=new Set(cur.adv||[]);
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
    box.innerHTML=wb.sheets.map((s,i)=>`<div class="sht${i===wb.active?' on':''}" data-sh="${i}"${s.tab?` style="border-bottom:3px solid ${s.tab}"`:''}>${s.prot?'🔒 ':''}${esc(s.name)}</div>`).join('');
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
    if(sbBlock())return;
    tx(()=>{const s=newSheet(uniqueName('Sheet'+(wb.sheets.length+1)));wb.sheets.splice(wb.active+1,0,s);wb.active++;cur=s;},true);
    drawTabs();sync();
  }
  function renameSheet(i){
    if(sbBlock())return;
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
    if(sbBlock())return;
    if(wb.sheets.length<2){toast('Buku kerja harus memiliki minimal satu lembar.');return;}
    const run=()=>{tx(()=>{wb.sheets.splice(i,1);wb.active=Math.min(wb.active,wb.sheets.length-1);if(i<wb.active||i===wb.active)wb.active=Math.max(0,Math.min(i,wb.sheets.length-1));cur=wb.sheets[wb.active];},true);drawTabs();sync();};
    const s=wb.sheets[i];if(s.cells.size)confirmDlg(`Hapus lembar "${s.name}" secara permanen? Tindakan ini bisa dibatalkan dengan Urungkan.`,run);else run();
  }
  function moveSheet(i,d){
    if(sbBlock())return;
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
    if(e.key==='Enter'){if(selectAddr(nb.value)||gotoName(nb.value.trim())){vp.focus();}else{toast('Referensi tidak valid.');}}
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
    fmtcells(tab){if(tab==='page'){pageSetupDialog();return;}formatCells(tab);},
    rowh(){if(protBlock('formatRows'))return;const r=rg().r1;askNumber('Tinggi Baris','Tinggi baris (pt):',+(cur.rh.size(r)*.75).toFixed(2),v=>tx(()=>{const g=rg();for(let i=g.r1;i<=Math.min(g.r2,g.r1+5000);i++){cur.rh.ov.set(i,Math.round(v/.75));cur.rh.hid.delete(i);(cur.rh.manual=cur.rh.manual||new Set()).add(i);}cur.rh.dirty=true;}));},
    colw(){if(protBlock('formatCols'))return;const c=rg().c1;askNumber('Lebar Kolom','Lebar kolom:',+((cur.cw.size(c)-5)/7).toFixed(2),v=>tx(()=>{const g=rg();for(let i=g.c1;i<=Math.min(g.c2,g.c1+2000);i++){cur.cw.ov.set(i,Math.round(v*7+5));cur.cw.hid.delete(i);}cur.cw.dirty=true;}));},
    defw(){if(protBlock('formatCols'))return;askNumber('Lebar Default','Lebar kolom standar:',+((cur.cw.def-5)/7).toFixed(2),v=>tx(()=>{cur.cw.def=Math.round(v*7+5);cur.cw.dirty=true;}));},
    fitr(){if(protBlock('formatRows'))return;const g=rg();const rs=[];for(let r=g.r1;r<=Math.min(g.r2,g.r1+5000);r++)rs.push(r);autoFitRowsH(rs);},
    fitc(){if(protBlock('formatCols'))return;const g=rg();const cs=[];for(let c=g.c1;c<=Math.min(g.c2,g.c1+500);c++)cs.push(c);autoFitCols(cs);},
    hider(){setHidden('r',true);},unhider(){setHidden('r',false);},hidec(){setHidden('c',true);},unhidec(){setHidden('c',false);},
    rensheet(){renameSheet(wb.active);},
    delsheet(){deleteSheet(wb.active);},
    inssheet(){addSheet();},
    mvl(){moveSheet(wb.active,-1);},mvr(){moveSheet(wb.active,1);}
  };
  const sels={
    font:v=>applyStyle({font:v===DF()?null:v}),
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
  function fnDialog(){fxDialog();}

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
      lock:st0.lock!==false,hide:!!st0.hide,font:st0.font||DF(),sz:st0.sz||11,b:!!st0.b,i:!!st0.i,u:st0.u||0,s:!!st0.s,color:st0.color||'#000000',fill:st0.fill||null,bd:{}};
    const bdSet={t:null,b:null,l:null,r:null,h:undefined,v:undefined};let bdStyleSel='thin',bdColorSel='#000000';
    const TABN=[['number','Angka'],['align','Perataan'],['font','Font'],['border','Garis Tepi'],['fill','Isi'],['prot','Proteksi']];
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
        font:S.font===DF()?null:S.font,sz:S.sz===11?null:S.sz,b:S.b||null,i:S.i||null,u:S.u||null,s:S.s||null,color:S.color==='#000000'?null:S.color,fill:S.fill,lock:S.lock?null:false,hide:S.hide||null};
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
      }else if(t==='prot'){
        P.innerHTML=`<label><input type="checkbox" class="lk" ${S.lock?'checked':''}> Terkunci</label><label><input type="checkbox" class="hd" ${S.hide?'checked':''}> Tersembunyi (rumus tidak tampil di bilah rumus)</label><div style="font-size:11px;color:#666;max-width:320px">Penguncian atau penyembunyian sel baru berlaku setelah lembar dilindungi (Tinjau → Lindungi Lembar).</div>`;
        P.querySelector('.lk').onchange=e=>S.lock=e.target.checked;P.querySelector('.hd').onchange=e=>S.hide=e.target.checked;
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
    clear:()=>[I('Hapus Semua',()=>clearWhat('all'),'⌧'),I('Hapus Format',()=>clearWhat('formats'),'🖌'),I('Hapus Isi',()=>clearWhat('contents'),'✂'),I('Hapus Hyperlink',()=>styleSel(st=>{const n={...st};if(n.link){delete n.link;delete n.color;delete n.u;}return n;}),'🔗')],
    sort:()=>[I('Urutkan A ke Z',()=>quickSort(true),'↓'),I('Urutkan Z ke A',()=>quickSort(false),'↑'),I('Urutan Kustom…',sortDialog,'⇅'),{sep:1},I('Filter',toggleFilter,'⧩',{ck:!!cur.filter}),I('Bersihkan Filter',()=>{if(cur.filter)tx(()=>{cur.filter.crit={};applyFilter(true);});},'✖',{dis:!cur.filter})],
    find:()=>[I('Cari…',showFind,'🔍'),I('Ganti…',showFind,'ab'),I('Pergi ke…',acts.goto,'→'),{sep:1},I('Pilih Semua',()=>select(0,0,MAXR-1,MAXC-1,true),'⬚')],
    ctx:()=>[I('Potong',()=>copyAct(true),'✂'),I('Salin',()=>copyAct(false),'⎘'),I('Tempel',()=>acts.paste(),'📋'),{l:'Tempel Spesial',i:'',sub:menus.paste()},{sep:1},I('Sisipkan…',insertDialog,'⊞'),I('Hapus…',deleteDialog,'⊟'),I('Hapus Isi',()=>clearWhat('contents'),'⌫'),{sep:1},I(cmAt(cur.sel.ar,cur.sel.ac)?'Edit Komentar':'Sisipkan Komentar',commentDialog,'🗨'),...(cmAt(cur.sel.ar,cur.sel.ac)?[I('Hapus Komentar',commentDelete,'✖')]:[]),{sep:1},
      I('Urutkan A ke Z',()=>quickSort(true),'↓'),I('Lebar Kolom…',acts.colw,'↔'),I('Tinggi Baris…',acts.rowh,'↕'),I('Sembunyikan Baris',acts.hider),I('Sembunyikan Kolom',acts.hidec),{sep:1},I('Format Sel…',()=>formatCells('number'),'▤')],
    sheetctx:()=>[I('Sisipkan Lembar',addSheet,'▭'),I('Hapus',()=>deleteSheet(wb.active),'✖'),I('Ganti Nama',acts.rensheet,'✎'),I('Pindahkan ke Kiri',acts.mvl,'◀'),I('Pindahkan ke Kanan',acts.mvr,'▶'),{l:'Warna Tab',i:'🎨',sub:()=>palette(c=>{tx(()=>{cur.tab=c;},true);drawTabs();},{none:'Tanpa Warna'})}]
  };
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


  /* ===== SISIPKAN: objek, bagan, sparkline, tautan, simbol ===== */
  const CCOL=['#4472c4','#ed7d31','#a5a5a5','#ffc000','#5b9bd5','#70ad47','#264478','#9e480e'];
  const CHT=[['col','Kolom Berkelompok'],['colS','Kolom Bertumpuk'],['bar','Batang Berkelompok'],['barS','Batang Bertumpuk'],['line','Garis'],['lineM','Garis dengan Penanda'],['pie','Pai'],['donut','Donat'],['area','Area'],['areaS','Area Bertumpuk'],['scatter','Sebar']];
  const SHP=[['rect','Persegi Panjang'],['rrect','Persegi Panjang Tumpul'],['ellipse','Elips'],['tri','Segitiga'],['diamond','Belah Ketupat'],['rarrow','Panah Kanan'],['star','Bintang'],['line','Garis']];
  const niceTicks=(mn,mx)=>{if(mn===mx)mx=mn+1;const raw=(mx-mn)/5,p=Math.pow(10,Math.floor(Math.log10(raw))),f=raw/p,st=(f<=1?1:f<=2?2:f<=5?5:10)*p,lo=Math.floor(mn/st+1e-9)*st,hi=Math.ceil(mx/st-1e-9)*st,t=[];for(let v=lo;v<=hi+st/2;v+=st)t.push(+v.toFixed(10));return t;};
  const fT=v=>fmtValue(+v.toFixed(8),'General').t;
  const parseRng=t=>{const m=/^\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$/.exec((t||'').trim());if(!m)return null;const a=parseAddr(m[1]+m[2]),b=m[3]?parseAddr(m[3]+m[4]):a;if(!a||!b)return null;return {r1:Math.min(a.r,b.r),r2:Math.max(a.r,b.r),c1:Math.min(a.c,b.c),c2:Math.max(a.c,b.c)};};
  const rngStr=g=>g.r1===g.r2&&g.c1===g.c2?addr(g.r1,g.c1):addr(g.r1,g.c1)+':'+addr(g.r2,g.c2);

  function chartData(o){
    const g=o.src,val=(r,c)=>cellVal(cur,r,c),R=g.r2-g.r1+1,C=g.c2-g.c1+1,txt=v=>typeof v==='string'&&v!=='';
    let hdr=R>1;for(let c=g.c1;c<=g.c2&&hdr;c++){const v=val(g.r1,c);if(v!==null&&v!==''&&typeof v!=='string')hdr=false;}
    if(hdr){let any=false;for(let c=g.c1;c<=g.c2;c++)if(txt(val(g.r1,c)))any=true;if(!any)hdr=false;}
    const r0=g.r1+(hdr?1:0),n=Math.max(0,Math.min(g.r2-r0+1,300));
    let cat=C>1&&n>0;for(let r=r0;r<r0+n&&cat;r++)if(!txt(val(r,g.c1)))cat=false;
    let xs=null,c0=g.c1;const cats=[];
    if(o.ct==='scatter'){if(C>1){xs=[];for(let r=r0;r<r0+n;r++){const v=val(r,g.c1);xs.push(typeof v==='number'?v:r-r0+1);}c0=g.c1+1;}else xs=Array.from({length:n},(_,i)=>i+1);}
    else if(cat){for(let r=r0;r<r0+n;r++)cats.push(String(val(r,g.c1)));c0=g.c1+1;}
    else for(let i=0;i<n;i++)cats.push(String(i+1));
    const series=[];
    for(let c=c0;c<=g.c2;c++){const vals=[];for(let r=r0;r<r0+n;r++){const v=val(r,c);vals.push(typeof v==='number'?v:0);}
      series.push({name:hdr&&txt(val(g.r1,c))?val(g.r1,c):'Seri'+(series.length+1),vals});}
    return {cats,series,xs,n};
  }
  function chartSvg(o,w,h){
    const d=chartData(o),ct=o.ct,S=d.series,n=d.n;
    let s=`<svg width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" font-family="Calibri,Carlito,sans-serif"><rect x=".5" y=".5" width="${w-1}" height="${h-1}" fill="#fff" stroke="#d9d9d9"/>`;
    const title=o.title!==null&&o.title!==undefined?o.title:(S.length===1?S[0].name:'Judul Bagan');
    const fs=Math.max(8,Math.min(14,w/34));
    if(title)s+=`<text x="${w/2}" y="${fs+10}" text-anchor="middle" font-size="${fs+2}" fill="#595959">${esc(title)}</text>`;
    if(!S.length||!n)return s+`<text x="${w/2}" y="${h/2}" text-anchor="middle" fill="#999" font-size="${fs}">Tidak ada data</text></svg>`;
    const pie=ct==='pie'||ct==='donut';
    const legend=pie?d.cats.slice(0,12).map((c,i)=>[c,CCOL[i%8]]):(S.length>1?S.map((q,i)=>[q.name,CCOL[i%8]]):[]);
    const top=title?fs+22:12,bot=legend.length?30:10;
    if(legend.length){const fsz=Math.max(8,fs-2),ws=legend.map(l=>14+String(l[0]).length*fsz*.55+12);let x=Math.max(6,(w-ws.reduce((a,b)=>a+b,0))/2);
      legend.forEach((l,i)=>{s+=`<rect x="${x}" y="${h-22}" width="9" height="9" fill="${l[1]}"/><text x="${x+13}" y="${h-13}" font-size="${fsz}" fill="#595959">${esc(l[0])}</text>`;x+=ws[i];});}
    if(pie){
      const vals=S[0].vals.map(v=>Math.max(0,v)),tot=vals.reduce((a,b)=>a+b,0)||1;
      const cxp=w/2,cyp=top+(h-top-bot)/2,r=Math.max(10,Math.min(w-20,h-top-bot-10)/2);let a0=-Math.PI/2;
      vals.forEach((v,i)=>{const a1=a0+v/tot*Math.PI*2;
        if(v>0){if(v/tot>=.9999)s+=`<circle cx="${cxp}" cy="${cyp}" r="${r}" fill="${CCOL[i%8]}" stroke="#fff"/>`;
          else s+=`<path d="M${cxp} ${cyp}L${cxp+r*Math.cos(a0)} ${cyp+r*Math.sin(a0)}A${r} ${r} 0 ${a1-a0>Math.PI?1:0} 1 ${cxp+r*Math.cos(a1)} ${cyp+r*Math.sin(a1)}Z" fill="${CCOL[i%8]}" stroke="#fff" stroke-width="1.5"/>`;
          if(v/tot>.05){const am=(a0+a1)/2,rr=r*(ct==='donut'?.78:.65);s+=`<text x="${cxp+rr*Math.cos(am)}" y="${cyp+rr*Math.sin(am)+4}" text-anchor="middle" font-size="${fs-1}" fill="#fff">${Math.round(v/tot*100)}%</text>`;}}
        a0=a1;});
      if(ct==='donut')s+=`<circle cx="${cxp}" cy="${cyp}" r="${r*.55}" fill="#fff"/>`;
      return s+'</svg>';
    }
    const horiz=ct==='bar'||ct==='barS',stack=ct==='colS'||ct==='barS'||ct==='areaS',sc=ct==='scatter';
    const L=horiz?Math.min(110,Math.max(40,...d.cats.map(c=>String(c).length*5.5+10))):46;
    const px0=L,px1=w-14,py0=top,py1=h-bot-18;
    let mx=-Infinity,mn=Infinity;
    if(stack){for(let i=0;i<n;i++){let p=0,q=0;for(const z of S){const v=z.vals[i];if(v>=0)p+=v;else q+=v;}mx=Math.max(mx,p);mn=Math.min(mn,q);}}
    else for(const z of S)for(const v of z.vals){mx=Math.max(mx,v);mn=Math.min(mn,v);}
    if(!sc){mn=Math.min(0,mn);mx=Math.max(0,mx);}
    const tk=niceTicks(mn,mx),v0=tk[0],v1=tk[tk.length-1];
    const vy=v=>py1-(v-v0)/(v1-v0)*(py1-py0),vx=v=>px0+(v-v0)/(v1-v0)*(px1-px0);
    const fz=Math.max(8,fs-3);
    if(horiz){tk.forEach(t=>{const x=vx(t);s+=`<line x1="${x}" y1="${py0}" x2="${x}" y2="${py1}" stroke="#d9d9d9"/><text x="${x}" y="${py1+13}" text-anchor="middle" font-size="${fz}" fill="#595959">${fT(t)}</text>`;});}
    else tk.forEach(t=>{const y=vy(t);s+=`<line x1="${px0}" y1="${y}" x2="${px1}" y2="${y}" stroke="#d9d9d9"/><text x="${px0-6}" y="${y+3}" text-anchor="end" font-size="${fz}" fill="#595959">${fT(t)}</text>`;});
    if(sc){
      const xt=niceTicks(Math.min(...d.xs),Math.max(...d.xs)),x0=xt[0],x1=xt[xt.length-1],sx=v=>px0+(v-x0)/(x1-x0)*(px1-px0);
      xt.forEach(t=>{const x=sx(t);s+=`<line x1="${x}" y1="${py0}" x2="${x}" y2="${py1}" stroke="#eee"/><text x="${x}" y="${py1+13}" text-anchor="middle" font-size="${fz}" fill="#595959">${fT(t)}</text>`;});
      S.forEach((z,si)=>z.vals.forEach((v,i)=>{s+=`<circle cx="${sx(d.xs[i])}" cy="${vy(v)}" r="4" fill="${CCOL[si%8]}"/>`;}));
      return s+'</svg>';
    }
    const band=(horiz?py1-py0:px1-px0)/n,step=Math.ceil(n/Math.max(1,Math.floor((horiz?py1-py0:px1-px0)/(horiz?16:44))));
    d.cats.forEach((c,i)=>{if(i%step)return;const t=esc(String(c).slice(0,14));
      s+=horiz?`<text x="${px0-6}" y="${py0+band*(i+.5)+3}" text-anchor="end" font-size="${fz}" fill="#595959">${t}</text>`:`<text x="${px0+band*(i+.5)}" y="${py1+14}" text-anchor="middle" font-size="${fz}" fill="#595959">${t}</text>`;});
    s+=horiz?`<line x1="${vx(0)}" y1="${py0}" x2="${vx(0)}" y2="${py1}" stroke="#bfbfbf"/>`:`<line x1="${px0}" y1="${vy(0)}" x2="${px1}" y2="${vy(0)}" stroke="#bfbfbf"/>`;
    if(ct==='col'||ct==='colS'||horiz){
      const gw=band/1.8;
      for(let i=0;i<n;i++){let pos=0,neg=0;
        S.forEach((z,si)=>{const v=z.vals[i];let a,b;
          if(stack){if(v>=0){a=pos;b=pos+v;pos=b;}else{a=neg;b=neg+v;neg=b;}}else{a=0;b=v;}
          const bw=stack?gw:gw/S.length,off=stack?0:si*bw;
          if(!horiz){const y1=vy(Math.max(a,b)),y2=vy(Math.min(a,b));s+=`<rect x="${px0+band*i+(band-gw)/2+off}" y="${y1}" width="${bw}" height="${Math.max(.5,y2-y1)}" fill="${CCOL[si%8]}"/>`;}
          else{const x1=vx(Math.min(a,b)),x2=vx(Math.max(a,b));s+=`<rect x="${x1}" y="${py0+band*i+(band-gw)/2+off}" width="${Math.max(.5,x2-x1)}" height="${bw}" fill="${CCOL[si%8]}"/>`;}
        });}
    }else{
      const px=i=>px0+band*(i+.5);
      const cum=S.map(()=>[]);
      for(let i=0;i<n;i++){let acc=0;S.forEach((z,si)=>{acc=stack?acc+z.vals[i]:z.vals[i];cum[si][i]=acc;});}
      if(ct==='area'||ct==='areaS'){
        const order=stack?[...S.keys()].reverse():[...S.keys()];
        for(const si of order){const top2=cum[si].map((v,i)=>`${px(i)},${vy(v)}`);const bot2=stack&&si>0?cum[si-1].map((v,i)=>`${px(i)},${vy(v)}`).reverse():[`${px(n-1)},${vy(0)}`,`${px(0)},${vy(0)}`];
          s+=`<polygon points="${top2.concat(bot2).join(' ')}" fill="${CCOL[si%8]}" fill-opacity=".8"/>`;}
      }else S.forEach((z,si)=>{
        s+=`<polyline points="${cum[si].map((v,i)=>`${px(i)},${vy(v)}`).join(' ')}" fill="none" stroke="${CCOL[si%8]}" stroke-width="2.25" stroke-linejoin="round"/>`;
        if(ct==='lineM')cum[si].forEach((v,i)=>{s+=`<circle cx="${px(i)}" cy="${vy(v)}" r="3.5" fill="${CCOL[si%8]}"/>`;});
      });
    }
    return s+'</svg>';
  }
  function shapeSvg(o,w,h){
    const f=o.fill||'#4472c4',a=`fill="${f}" stroke="#2f528f" stroke-width="1.5" vector-effect="non-scaling-stroke"`;
    let b='';
    switch(o.kind){
      case 'rrect':b=`<rect x="1" y="1" width="98" height="98" rx="14" ${a}/>`;break;
      case 'ellipse':b=`<ellipse cx="50" cy="50" rx="49" ry="49" ${a}/>`;break;
      case 'tri':b=`<polygon points="50,2 98,98 2,98" ${a}/>`;break;
      case 'diamond':b=`<polygon points="50,2 98,50 50,98 2,50" ${a}/>`;break;
      case 'rarrow':b=`<polygon points="2,30 60,30 60,6 98,50 60,94 60,70 2,70" ${a}/>`;break;
      case 'star':{const p=[];for(let i=0;i<10;i++){const r=i%2?20:49,t=-Math.PI/2+i*Math.PI/5;p.push((50+r*Math.cos(t)).toFixed(1)+','+(52+r*Math.sin(t)).toFixed(1));}b=`<polygon points="${p.join(' ')}" ${a}/>`;break;}
      case 'line':b=`<line x1="2" y1="98" x2="98" y2="2" stroke="${f}" stroke-width="3" vector-effect="non-scaling-stroke"/>`;break;
      default:b=`<rect x="1" y="1" width="98" height="98" ${a}/>`;
    }
    return `<svg width="${w}" height="${h}" viewBox="0 0 100 100" preserveAspectRatio="none">${b}</svg>`;
  }
  function objHtml(o,x,y,w,h){
    if(o.hidden)return '';
    const sel=selSet.has(o.id);let inner='';const fz=11*4/3*Z;
    if(o.type==='img')inner=`<img src="${o.url||o.src}" draggable="false">`;
    else if(o.type==='chart')inner=chartSvg(o,w,h);
    else if(o.type==='shape')inner=shapeSvg(o,w,h)+(o.text?`<div class="tx" style="position:absolute;left:0;top:0;color:#fff;font-size:${fz}px">${esc(o.text)}</div>`:'');
    else inner=`<div class="tx" style="background:#fff;border:1px solid #7f7f7f;justify-content:flex-start;align-items:flex-start;text-align:left;color:#000;font-size:${fz}px">${esc(o.text||'')}</div>`;
    return `<div class="ob${sel?' sel':''}" data-o="${o.id}" style="left:${x}px;top:${y}px;width:${w}px;height:${h}px;${o.rot||o.fx||o.fy?`transform:rotate(${o.rot||0}deg) scale(${o.fx?-1:1},${o.fy?-1:1});`:''}">${inner}${sel&&selSet.size===1?'<i class="orz" data-orz="1"></i>':''}</div>`;
  }
  function sparkSvg(sp,w,h){
    const g=sp.src,v=[];
    for(let r=g.r1;r<=g.r2;r++)for(let c=g.c1;c<=g.c2;c++){const x=cellVal(cur,r,c);v.push(typeof x==='number'?x:null);}
    const nn=v.filter(x=>x!==null);if(!nn.length)return '';
    let mn=Math.min(...nn),mx=Math.max(...nn);if(sp.type!=='line'){mn=Math.min(mn,0);mx=Math.max(mx,0);}if(mn===mx){mn-=1;mx+=1;}
    const pad=4,W=w-2*pad,H=h-2*pad,y=x=>pad+H-(x-mn)/(mx-mn)*H,n=v.length,bw=W/n;
    let b='';
    if(sp.type==='line'){const pts=v.map((x,i)=>x===null?null:`${(pad+(n>1?i*W/(n-1):W/2)).toFixed(1)},${y(x).toFixed(1)}`).filter(Boolean);b=`<polyline points="${pts.join(' ')}" fill="none" stroke="#4472c4" stroke-width="1.5" stroke-linejoin="round"/>`;}
    else if(sp.type==='col')v.forEach((x,i)=>{if(x===null)return;const y0=y(0),y1=y(x);b+=`<rect x="${pad+i*bw+1}" y="${Math.min(y0,y1)}" width="${Math.max(1,bw-2)}" height="${Math.max(1,Math.abs(y0-y1))}" fill="${x<0?'#ed7d31':'#4472c4'}"/>`;});
    else v.forEach((x,i)=>{if(x===null)return;const m=pad+H/2,up=x>0,hh=x===0?1:H/2-1;b+=`<rect x="${pad+i*bw+1}" y="${up?m-hh:m+(x===0?-.5:1)}" width="${Math.max(1,bw-2)}" height="${hh}" fill="${x<0?'#ed7d31':x===0?'#a5a5a5':'#4472c4'}"/>`;});
    return `<svg width="${w}" height="${h}">${b}</svg>`;
  }

  const newId=()=>'o'+Date.now().toString(36)+Math.random().toString(36).slice(2,5);
  const viewPos=()=>{const k=(cur.objs.length%8)*24;return {x:Math.round(vp.scrollLeft/Z)+60+k,y:Math.round(vp.scrollTop/Z)+40+k};};
  function addObj(o){if(protBlock('objects'))return;tx(()=>{o.id=newId();Object.assign(o,viewPos());cur.objs=[...cur.objs,o];selObj=o.id;selSet=new Set([o.id]);});}
  function deleteObj(){if(protBlock('objects'))return;const ids=new Set(selSet);tx(()=>{cur.objs=cur.objs.filter(o=>!ids.has(o.id));selObj=null;selSet=new Set();});}
  function addChart(type){
    let g=rg();if(g.r1===g.r2&&g.c1===g.c2)g=region();
    if(g.r1===g.r2&&g.c1===g.c2&&!hasVal(cur,g.r1,g.c1)){toast('Pilih data yang akan dibuat bagan.');return;}
    if(g.r2-g.r1>2000)g.r2=g.r1+2000;if(g.c2-g.c1>40)g.c2=g.c1+40;
    addObj({type:'chart',ct:type,src:{r1:g.r1,c1:g.c1,r2:g.r2,c2:g.c2},w:480,h:288,title:null});
  }
  function objEdit(id){
    const o=cur.objs.find(q=>q.id===id);if(!o)return;if(protBlock('objects'))return;selObj=id;selSet=new Set([id]);schedule();
    const upd=p=>tx(()=>{cur.objs=cur.objs.map(q=>q.id===id?{...q,...p}:q);});
    if(o.type==='text'||o.type==='shape'){
      const d=dlg(o.type==='text'?'Kotak Teks':'Teks Bentuk',`<textarea class="t" rows="4" style="width:300px">${esc(o.text||'')}</textarea>`,d=>{upd({text:d.querySelector('.t').value});});
    }else if(o.type==='chart'){
      const cur0=o.title!==null&&o.title!==undefined?o.title:'';
      dlg('Format Bagan',`<div class="r2"><span>Judul:</span><input type="text" class="t" value="${esc(cur0)}" placeholder="(otomatis)"><span>Jenis:</span><select class="k">${CHT.map(([k,l])=>`<option value="${k}" ${k===o.ct?'selected':''}>${l}</option>`).join('')}</select><span>Rentang data:</span><input type="text" class="rg" value="${rngStr(o.src)}"></div>`,d=>{
        const r=parseRng(d.querySelector('.rg').value);if(!r){toast('Rentang data tidak valid.');return false;}
        const t=d.querySelector('.t').value;upd({title:t===''?null:t,ct:d.querySelector('.k').value,src:r});});
    }
  }
  function symbolDialog(){
    const SY='© ® ™ ° ± × ÷ ≠ ≤ ≥ ≈ ∞ π Ω α β γ δ θ λ μ σ Σ √ ∫ ∂ € £ ¥ ¢ § ¶ † ‡ • … ‰ ← → ↑ ↓ ↔ ✓ ✗ ★ ☆ ♠ ♣ ♥ ♦ ½ ¼ ¾ ² ³ ¹ ª º « »'.split(' ');
    const d=dlg('Simbol',`<div class="sym">${SY.map(c=>`<button data-c="${c}">${c}</button>`).join('')}</div>`,()=>{if(ed)(ed.bar?fxi:edEl).focus();},'Tutup',true);
    d.querySelector('.sym').onclick=e=>{const b=e.target.closest('button');if(!b)return;
      if(!ed){const s=cur.sel;startEdit(s.ar,s.ac);}
      const el=ed.bar?fxi:edEl,p=el.selectionStart==null?el.value.length:el.selectionStart;el.setRangeText(b.dataset.c,p,el.selectionEnd==null?p:el.selectionEnd,'end');onEdInput(el);};
  }
  function openLink(u){
    if(u[0]==='#'){if(!selectAddr(u.slice(1).replace(/^.*!/,'')))toast('Referensi tidak valid.');return;}
    window.open(/^[a-z][a-z0-9+.-]*:/i.test(u)?u:'https://'+u,'_blank','noopener');
  }
  function hyperlinkDialog(){
    const s=cur.sel,cl=cellAt(cur,s.ar,s.ac),st=(cl&&cl.st)||{};
    const txt=cl?fmtValue(cellVal(cur,s.ar,s.ac),st.fmt).t:'';
    dlg('Sisipkan Hyperlink',`<div class="r2"><span>Teks yang ditampilkan:</span><input type="text" class="t" value="${esc(txt)}"><span>Alamat:</span><input type="text" class="u" value="${esc(st.link||'')}" placeholder="https://… atau #B5"></div><div style="font-size:11px;color:#666">Tekan Ctrl+klik pada sel untuk membuka tautan.</div>`,d=>{
      const u=d.querySelector('.u').value.trim();if(!u)return false;const t=d.querySelector('.t').value||u;
      tx(()=>{putCell(cur,s.ar,s.ac,{v:t,st:cleanSt({...st,link:u,color:'#0563c1',u:1})});});});
  }
  function sparkDialog(type){
    let g=rg();if(g.r1===g.r2&&g.c1===g.c2)g=region();
    const loc=g.r2>g.r1?`${addr(g.r1,g.c2+1)}:${addr(g.r2,g.c2+1)}`:addr(g.r1,Math.min(MAXC-1,g.c2+1));
    dlg('Buat Sparkline',`<div class="r2"><span>Rentang data:</span><input type="text" class="d" value="${rngStr(g)}"><span>Rentang lokasi:</span><input type="text" class="l" value="${loc}"></div>`,d=>{
      const D=parseRng(d.querySelector('.d').value),Lc=parseRng(d.querySelector('.l').value);
      if(!D||!Lc){toast('Rentang tidak valid.');return false;}
      const R=D.r2-D.r1+1,C=D.c2-D.c1+1,nL=(Lc.r2-Lc.r1+1)*(Lc.c2-Lc.c1+1),cells=[];
      for(let r=Lc.r1;r<=Lc.r2;r++)for(let c=Lc.c1;c<=Lc.c2;c++)cells.push([r,c]);
      const items=[];
      if(nL===1)items.push({r:cells[0][0],c:cells[0][1],src:D});
      else if(nL===R)cells.forEach(([r,c],i)=>items.push({r,c,src:{r1:D.r1+i,r2:D.r1+i,c1:D.c1,c2:D.c2}}));
      else if(nL===C)cells.forEach(([r,c],i)=>items.push({r,c,src:{r1:D.r1,r2:D.r2,c1:D.c1+i,c2:D.c1+i}}));
      else{toast('Jumlah sel lokasi harus sama dengan jumlah baris atau kolom data.');return false;}
      tx(()=>{cur.spark=cur.spark.filter(sp=>!(sp.r>=Lc.r1&&sp.r<=Lc.r2&&sp.c>=Lc.c1&&sp.c<=Lc.c2)).concat(items.map(i=>({...i,type})));});
    });
  }
  function addPicture(file){
    const url=URL.createObjectURL(file),fr=new FileReader();
    fr.onload=()=>{const im=new Image();im.onload=()=>{let w=im.naturalWidth||200,h=im.naturalHeight||150;const k=Math.min(1,360/w,300/h);w=Math.round(w*k);h=Math.round(h*k);
      addObj({type:'img',url,src:fr.result,w,h});};im.onerror=()=>toast('Berkas bukan gambar yang valid.');im.src=fr.result;};
    fr.readAsDataURL(file);
  }
  $('.imgin').onchange=e=>{const f=e.target.files[0];if(f)addPicture(f);e.target.value='';};
  const chartMenu=(l)=>()=>l.map(([k,t])=>I(t,()=>addChart(k),''));
  Object.assign(menus,{
    chCol:chartMenu([['col','Kolom Berkelompok'],['colS','Kolom Bertumpuk']]),chBar:chartMenu([['bar','Batang Berkelompok'],['barS','Batang Bertumpuk']]),
    chLine:chartMenu([['line','Garis'],['lineM','Garis dengan Penanda']]),chPie:chartMenu([['pie','Pai'],['donut','Donat']]),
    chArea:chartMenu([['area','Area'],['areaS','Area Bertumpuk']]),chSct:chartMenu([['scatter','Sebar (XY)']]),
    shapes:()=>SHP.map(([k,t])=>I(t,()=>addObj({type:'shape',kind:k,w:k==='line'?140:130,h:k==='line'?60:90,text:'',fill:'#4472c4'}),''))
  });
  Object.assign(acts,{
    tblnow(){applyTable(6);},pic(){$('.imgin').click();},hyperlink:hyperlinkDialog,symbol:symbolDialog,
    textbox(){addObj({type:'text',w:180,h:60,text:'Kotak teks'});},
    spark(v){sparkDialog(v);}
  });


  /* ===== TATA LETAK HALAMAN ===== */
  const setPage=p=>tx(()=>{cur.page={...cur.page,m:{...cur.page.m},...p};});
  function applyTheme(i,part){
    const t=THEMES[i];if(part!=='font'){PAL=buildPal(t.c);wb.themeC=i;}
    if(part!=='color'){wb.font=t.f;wb.themeF=i;}
    X.style.setProperty('--df',DF());wb.ver++;sync();
  }
  const FTL=['Calibri','Cambria','Arial','Times New Roman','Georgia','Verdana','Segoe UI','Tahoma'];
  const near=(a,b)=>Math.abs(a-b)<.005;
  function pageSetupDialog(){
    const P=cur.page,cm=v=>+(v*2.54).toFixed(2);
    const d=dlg('Penyiapan Halaman',`<div class="r2"><span>Orientasi:</span><select class="o"><option value="portrait">Potret</option><option value="landscape">Lanskap</option></select>
      <span>Ukuran kertas:</span><select class="p">${Object.keys(PAPER).map(k=>`<option value="${k}">${PAPERN[k]}</option>`).join('')}</select>
      <span>Margin atas (cm):</span><input type="number" class="mt" step="0.1" min="0" value="${cm(P.m.t)}"><span>Margin bawah (cm):</span><input type="number" class="mb" step="0.1" min="0" value="${cm(P.m.b)}">
      <span>Margin kiri (cm):</span><input type="number" class="ml" step="0.1" min="0" value="${cm(P.m.l)}"><span>Margin kanan (cm):</span><input type="number" class="mr" step="0.1" min="0" value="${cm(P.m.r)}">
      <span>Skala (%):</span><input type="number" class="sc" min="10" max="400" value="${P.scale}"><span>Muat ke lebar:</span><select class="fw">${FITO.map(x=>`<option value="${x[0]}">${x[1]}</option>`).join('')}</select><span>Muat ke tinggi:</span><select class="fh">${FITO.map(x=>`<option value="${x[0]}">${x[1]}</option>`).join('')}</select></div>
      <label><input type="checkbox" class="pg"> Cetak garis kisi</label><label><input type="checkbox" class="ph"> Cetak tajuk baris dan kolom</label>`,d=>{
      const n=c=>Math.max(0,+d.querySelector(c).value||0)/2.54;
      setPage({orient:d.querySelector('.o').value,paper:d.querySelector('.p').value,m:{t:n('.mt'),b:n('.mb'),l:n('.ml'),r:n('.mr')},scale:Math.max(10,Math.min(400,+d.querySelector('.sc').value||100)),fitW:+d.querySelector('.fw').value,fitH:+d.querySelector('.fh').value,pg:d.querySelector('.pg').checked,ph:d.querySelector('.ph').checked});});
    d.querySelector('.o').value=P.orient;d.querySelector('.p').value=P.paper;d.querySelector('.fw').value=P.fitW;d.querySelector('.fh').value=P.fitH;d.querySelector('.pg').checked=P.pg;d.querySelector('.ph').checked=P.ph;
  }
  function titlesDialog(){
    const P=cur.page;
    dlg('Judul Cetak',`<div class="r2"><span>Baris untuk diulang di atas:</span><input type="text" class="r" placeholder="mis. 1:2" value="${P.tRows?`${P.tRows.r1+1}:${P.tRows.r2+1}`:''}"><span>Kolom untuk diulang di kiri:</span><input type="text" class="c" placeholder="mis. A:B" value="${P.tCols?`${colName(P.tCols.c1)}:${colName(P.tCols.c2)}`:''}"></div>`,d=>{
      const rv=d.querySelector('.r').value.trim(),cv=d.querySelector('.c').value.trim();let tr=null,tc=null,m;
      if(rv){m=/^\$?(\d+)(?::\$?(\d+))?$/.exec(rv);if(!m){toast('Baris tidak valid. Contoh: 1:2');return false;}tr={r1:+m[1]-1,r2:+(m[2]||m[1])-1};if(tr.r2<tr.r1||tr.r1<0){toast('Baris tidak valid.');return false;}}
      if(cv){m=/^\$?([A-Za-z]+)(?::\$?([A-Za-z]+))?$/.exec(cv);if(!m){toast('Kolom tidak valid. Contoh: A:B');return false;}tc={c1:colIdx(m[1]),c2:colIdx(m[2]||m[1])};}
      setPage({tRows:tr,tCols:tc});});
  }
  function renderSelp(){
    if(!selpEl)return;
    const cnt={};const name=o=>{const L={img:'Gambar',chart:'Bagan',shape:'Bentuk',text:'Kotak Teks'}[o.type]||'Objek';return o.name||L+' '+o._n;};
    cur.objs.forEach(o=>{const k=o.type;cnt[k]=(cnt[k]||0)+1;o._n=cnt[k];});
    selpEl.innerHTML=`<div class="sh"><span>Panel Seleksi</span><button data-x="1" title="Tutup">✕</button></div>${[...cur.objs].reverse().map(o=>`<div class="si${selSet.has(o.id)?' on':''}" data-id="${o.id}"><button data-eye="${o.id}" title="Tampilkan/sembunyikan">${o.hidden?'🚫':'👁'}</button><span>${esc(name(o))}</span></div>`).join('')||'<div class="si" style="color:#888">Tidak ada objek</div>'}`;
  }
  function toggleSelp(){
    if(selpEl){selpEl.remove();selpEl=null;return;}
    selpEl=document.createElement('div');selpEl.className='selp';$('.gwrap').appendChild(selpEl);renderSelp();
    selpEl.onclick=e=>{
      if(e.target.dataset.x){selpEl.remove();selpEl=null;return;}
      const eye=e.target.dataset.eye;
      if(eye){tx(()=>{cur.objs=cur.objs.map(o=>o.id===eye?{...o,hidden:!o.hidden}:o);});return;}
      const it=e.target.closest('[data-id]');if(it){selObj=it.dataset.id;selSet=e.shiftKey?new Set([...selSet,selObj]):new Set([selObj]);sync();}};
  }
  const selList=()=>cur.objs.filter(o=>selSet.has(o.id));
  function zMove(k){
    if(!selSet.size){toast('Pilih objek terlebih dahulu.');return;}
    tx(()=>{let a=[...cur.objs];const ids=selSet;
      if(k==='front')a=a.filter(o=>!ids.has(o.id)).concat(a.filter(o=>ids.has(o.id)));
      else if(k==='back')a=a.filter(o=>ids.has(o.id)).concat(a.filter(o=>!ids.has(o.id)));
      else if(k==='fwd'){for(let i=a.length-2;i>=0;i--)if(ids.has(a[i].id)&&!ids.has(a[i+1].id))[a[i],a[i+1]]=[a[i+1],a[i]];}
      else{for(let i=1;i<a.length;i++)if(ids.has(a[i].id)&&!ids.has(a[i-1].id))[a[i],a[i-1]]=[a[i-1],a[i]];}
      cur.objs=a;});
  }
  function alignObjs(k){
    const L=selList();if(L.length<2)return;
    tx(()=>{const x0=Math.min(...L.map(o=>o.x)),x1=Math.max(...L.map(o=>o.x+o.w)),y0=Math.min(...L.map(o=>o.y)),y1=Math.max(...L.map(o=>o.y+o.h));
      for(const o of cur.objs){if(!selSet.has(o.id))continue;
        if(k==='l')o.x=x0;else if(k==='r')o.x=x1-o.w;else if(k==='c')o.x=(x0+x1)/2-o.w/2;
        else if(k==='t')o.y=y0;else if(k==='b')o.y=y1-o.h;else o.y=(y0+y1)/2-o.h/2;}});
  }
  function rotObjs(k){
    if(!selSet.size){toast('Pilih objek terlebih dahulu.');return;}
    tx(()=>{for(const o of cur.objs){if(!selSet.has(o.id))continue;
      if(k==='r')o.rot=((o.rot||0)+90)%360;else if(k==='l')o.rot=((o.rot||0)+270)%360;else if(k==='v')o.fy=!o.fy;else o.fx=!o.fx;}});
  }
  $('.bgin').onchange=e=>{const f=e.target.files[0];e.target.value='';if(!f)return;cur.bg=URL.createObjectURL(f);dirty=true;sync();toast('Latar belakang hanya tampil di lembar ini dan belum disimpan ke .xlsx.');};
  Object.assign(sels,{
    fitw:v=>setPage({fitW:+v}),fith:v=>setPage({fitH:+v}),
    scale:v=>setPage({scale:Math.max(10,Math.min(400,Math.round(+v)||100))}),
    gview:v=>{cur.grid=!!v;sync();},gprint:v=>setPage({pg:!!v}),hview:v=>{cur.head=!!v;sync();},hprint:v=>setPage({ph:!!v})
  });
  Object.assign(acts,{
    bg(){if(cur.bg){cur.bg=null;sync();}else $('.bgin').click();},
    ptitles:titlesDialog,selpane:toggleSelp
  });
  const MG={n:[.75,.75,.7,.7],w:[1,1,1,1],s:[.75,.75,.25,.25]};
  const mgItem=(l,k)=>{const v=MG[k],P=cur.page.m;return I(l,()=>setPage({m:{t:v[0],b:v[1],l:v[2],r:v[3]}}),'▭',{ck:near(P.t,v[0])&&near(P.b,v[1])&&near(P.l,v[2])&&near(P.r,v[3])});};
  Object.assign(menus,{
    themes:()=>THEMES.map((t,i)=>I(t.n+'  ('+t.f+')',()=>applyTheme(i),'Aa',{ck:wb.themeC===i&&wb.themeF===i})),
    tcolors:()=>THEMES.map((t,i)=>({html:`<div class="mi" style="gap:6px"><span style="display:flex;gap:1px">${t.c.slice(2).map(c=>`<i style="display:block;width:10px;height:14px;background:${c}"></i>`).join('')}</span><span>${t.n}</span></div>`,on:d=>d.onclick=()=>{closePop();applyTheme(i,'color');}})),
    tfonts:()=>FTL.map(f=>I(f,()=>{wb.font=f;wb.themeF=-1;X.style.setProperty('--df',DF());wb.ver++;sync();},'A',{ck:DF()===f})),
    margins:()=>[mgItem('Normal  (atas/bawah 1,9 cm · kiri/kanan 1,8 cm)','n'),mgItem('Lebar  (2,5 cm di semua sisi)','w'),mgItem('Sempit  (kiri/kanan 0,6 cm)','s'),{sep:1},I('Margin Kustom…',pageSetupDialog,'⚙')],
    orient2:()=>[I('Potret',()=>setPage({orient:'portrait'}),'▯',{ck:cur.page.orient==='portrait'}),I('Lanskap',()=>setPage({orient:'landscape'}),'▭',{ck:cur.page.orient==='landscape'})],
    psize:()=>Object.keys(PAPER).map(k=>I(PAPERN[k],()=>setPage({paper:k}),'▯',{ck:cur.page.paper===k})),
    parea:()=>[I('Atur Area Cetak',()=>{let g=expandMerge(limitRange(rg()));setPage({area:{r1:g.r1,c1:g.c1,r2:g.r2,c2:g.c2}});},'▦'),I('Hapus Area Cetak',()=>setPage({area:null}),'✖',{dis:!cur.page.area})],
    fwd:()=>[I('Majukan',()=>zMove('fwd'),'⬆'),I('Bawa ke Depan',()=>zMove('front'),'⇈')],
    back:()=>[I('Mundurkan',()=>zMove('bk'),'⬇'),I('Kirim ke Belakang',()=>zMove('back'),'⇊')],
    align:()=>{const ok=selSet.size>=2;return [['Rata Kiri','l'],['Rata Tengah','c'],['Rata Kanan','r'],['Rata Atas','t'],['Rata Tengah Vertikal','m'],['Rata Bawah','b']].map(([l,k])=>I(l,()=>alignObjs(k),'⇤',{dis:!ok}));},
    rot:()=>[I('Putar Kanan 90°',()=>rotObjs('r'),'↻'),I('Putar Kiri 90°',()=>rotObjs('l'),'↺'),I('Balik Vertikal',()=>rotObjs('v'),'⇵'),I('Balik Horizontal',()=>rotObjs('h'),'⇄')]
  });
  X.style.setProperty('--df',DF());


  /* ===== RUMUS ===== */
  const q1=n=>/^[A-Za-z_][\w.]*$/.test(n)?n:"'"+n.replace(/'/g,"''")+"'";
  const absRef=(sh,g)=>`${q1(sh.name)}!$${colName(g.c1)}$${g.r1+1}`+(g.r1===g.r2&&g.c1===g.c2?'':`:$${colName(g.c2)}$${g.r2+1}`);
  const validName=n=>/^[A-Za-z_\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*$/.test(n)&&!/^[A-Za-z]{1,3}\d+$/.test(n)&&!/^[RrCc]$/.test(n)&&!/^(TRUE|FALSE)$/i.test(n)&&n.length<=255;
  const nameTarget=nm=>{try{const a=parseFormula(nm.ref.replace(/^=/,''));if(a.k!=='ref')return null;const sh=a.sheet?wb.sheets.find(x=>x.name.toLowerCase()===a.sheet.toLowerCase()):cur;return sh?{sh,r1:a.r1,c1:a.c1,r2:a.r2,c2:a.c2}:null;}catch(e){return null;}};
  function gotoName(t){
    const nm=wb.names.find(n=>n.name.toLowerCase()===t.toLowerCase());if(!nm)return false;
    const tg=nameTarget(nm);if(!tg){toast('Nama ini tidak merujuk ke rentang sel.');return true;}
    const i=wb.sheets.indexOf(tg.sh);if(i!==wb.active)switchSheet(i);
    select(tg.r1,tg.c1,tg.r2,tg.c2);return true;
  }
  function nameOfSel(){
    const g=expandMerge(rg());
    for(const nm of wb.names){const t=nameTarget(nm);if(t&&t.sh===cur&&t.r1===g.r1&&t.c1===g.c1&&t.r2===g.r2&&t.c2===g.c2)return nm.name;}
    return null;
  }
  const insertText=t=>{
    const s=cur.sel;
    if(!ed)startEdit(s.ar,s.ac,undefined);
    const el=ed.bar?fxi:edEl;if(el.value==='')el.value=ed.text='=';
    const p=el.selectionStart==null?el.value.length:el.selectionStart;
    el.setRangeText(t,p,el.selectionEnd==null?p:el.selectionEnd,'end');onEdInput(el);el.focus();
  };
  function insertFn(n){
    wb.recent=[n,...wb.recent.filter(x=>x!==n)].slice(0,10);
    const s=cur.sel;
    if(!ed)startEdit(s.ar,s.ac,'='+n+'(');else insertText((ed.text===''||!isFormulaText(ed.text)&&ed.text==='='?'':'')+n+'(');
    if(FNI[n])toast(FNI[n].syn);
  }
  const fnItems=list=>list.filter(n=>FN[n]).map(n=>({l:n,i:'ƒ',a:()=>insertFn(n),tip:FNI[n]&&FNI[n].d}));
  const catList=c=>Object.keys(FNI).filter(n=>FNI[n].c===c&&FN[n]);
  function fxDialog(){
    const d=dlg('Sisipkan Fungsi',`<div class="r2"><span>Cari fungsi:</span><input type="text" class="q" placeholder="mis. jumlah, SUM"><span>Kategori:</span><select class="cat">${FNCAT.map(c=>`<option value="${c[0]}">${c[1]}</option>`).join('')}</select></div><div class="fcl" style="height:170px"></div><div class="fdesc" style="min-height:54px;font-size:12px"></div>`,d=>{
      const on=d.querySelector('.fcl .on');if(!on)return false;insertFn(on.dataset.n);},'OK');
    const list=d.querySelector('.fcl'),desc=d.querySelector('.fdesc'),q=d.querySelector('.q'),cat=d.querySelector('.cat');
    const fill=()=>{
      const c=cat.value,t=q.value.trim().toUpperCase(),tl=q.value.trim().toLowerCase();
      let names=c==='rec'?wb.recent.filter(n=>FN[n]):c==='all'?Object.keys(FNI).filter(n=>FN[n]):catList(c);
      if(t)names=Object.keys(FNI).filter(n=>FN[n]&&(n.includes(t)||FNI[n].d.toLowerCase().includes(tl)));
      names=[...names].sort((a,b)=>c==='rec'&&!t?0:a.localeCompare(b));
      list.innerHTML=names.map(n=>`<div data-n="${n}">${n}</div>`).join('');desc.innerHTML='';
    };
    q.oninput=fill;cat.onchange=fill;fill();
    list.onclick=e=>{const t=e.target.closest('[data-n]');if(!t)return;list.querySelectorAll('.on').forEach(x=>x.classList.remove('on'));t.classList.add('on');const i=FNI[t.dataset.n];desc.innerHTML=`<b>${esc(i.syn)}</b><br>${esc(i.d)}`;};
    list.ondblclick=()=>d.querySelector('.ok').click();
  }
  /* nama terdefinisi */
  function defineName(existing,ref0){
    const g=expandMerge(rg());
    const nm0=existing||{name:'',ref:'='+absRef(cur,g)};
    dlg(existing?'Edit Nama':'Nama Baru',`<div class="r2"><span>Nama:</span><input type="text" class="n" value="${esc(nm0.name)}"><span>Merujuk ke:</span><input type="text" class="r" value="${esc(ref0||nm0.ref)}" style="width:260px"></div><div style="font-size:11px;color:#666">Nama harus diawali huruf atau _, tanpa spasi, dan bukan alamat sel.</div>`,d=>{
      const n=d.querySelector('.n').value.trim();let r=d.querySelector('.r').value.trim();
      if(!validName(n)){toast('Nama tidak valid.');return false;}
      if(!r){toast('Isi bagian "Merujuk ke".');return false;}
      if(r[0]!=='=')r='='+r;
      try{parseFormula(r.slice(1));}catch(e){toast('Rumus "Merujuk ke" tidak valid.');return false;}
      if(wb.names.some(x=>x!==existing&&x.name.toLowerCase()===n.toLowerCase())){toast('Nama ini sudah ada.');return false;}
      tx(()=>{const o={name:n,ref:r};if(existing)wb.names=wb.names.map(x=>x===existing?o:x);else wb.names=[...wb.names,o];},true);sync();
      if(nameMgrRefresh)nameMgrRefresh();});
  }
  let nameMgrRefresh=null;
  function nameManager(){
    let selI=-1;
    const body=()=>`<table style="border-collapse:collapse;width:100%;font-size:12px"><tr style="background:#f3f3f3"><th align="left" style="padding:3px 6px">Nama</th><th align="left" style="padding:3px 6px">Nilai</th><th align="left" style="padding:3px 6px">Merujuk ke</th></tr>${wb.names.map((n,i)=>{const t=nameTarget(n);let v='';if(t){const x=rawVal(wb,t.sh,t.r1,t.c1);v=t.r1===t.r2&&t.c1===t.c2?fmtValue(x,'General').t:'{…}';}
      return `<tr data-i="${i}" style="cursor:pointer;${i===selI?'background:#d3e8d9':''}"><td style="padding:3px 6px">${esc(n.name)}</td><td style="padding:3px 6px">${esc(v)}</td><td style="padding:3px 6px">${esc(n.ref)}</td></tr>`;}).join('')||'<tr><td colspan="3" style="padding:8px;color:#888">Belum ada nama terdefinisi</td></tr>'}</table>`;
    const d=dlg('Pengelola Nama',`<div class="nm" style="min-width:420px;max-height:220px;overflow:auto;border:1px solid #bbb"></div><div style="display:flex;gap:6px"><button class="bt nw">Baru…</button><button class="bt ed">Edit…</button><button class="bt dl">Hapus</button></div>`,()=>{nameMgrRefresh=null;},'Tutup',true);
    const box=d.querySelector('.nm'),draw=()=>{box.innerHTML=body();};draw();nameMgrRefresh=()=>{selI=-1;draw();};
    box.onclick=e=>{const tr=e.target.closest('[data-i]');if(tr){selI=+tr.dataset.i;draw();}};
    d.querySelector('.nw').onclick=()=>defineName(null);
    d.querySelector('.ed').onclick=()=>{if(selI>=0)defineName(wb.names[selI]);else toast('Pilih nama terlebih dahulu.');};
    d.querySelector('.dl').onclick=()=>{if(selI<0){toast('Pilih nama terlebih dahulu.');return;}const n=wb.names[selI];tx(()=>{wb.names=wb.names.filter(x=>x!==n);},true);sync();selI=-1;draw();};
  }
  function namesFromSel(){
    const g=expandMerge(rg());if(g.r1===g.r2&&g.c1===g.c2){toast('Pilih rentang yang berisi label dan data.');return;}
    const hasT=(r,c)=>typeof cellVal(cur,r,c)==='string'&&cellVal(cur,r,c)!=='';
    let top=true;for(let c=g.c1;c<=g.c2;c++)if(!hasT(g.r1,c))top=false;
    let left=!top;if(!top)for(let r=g.r1;r<=g.r2;r++)if(!hasT(r,g.c1))left=false;
    const d=dlg('Buat Nama dari Pilihan',`<div style="font-size:12px">Buat nama dari nilai di:</div><label><input type="checkbox" class="t" ${top?'checked':''}> Baris atas</label><label><input type="checkbox" class="l" ${left?'checked':''}> Kolom kiri</label><label><input type="checkbox" class="b"> Baris bawah</label><label><input type="checkbox" class="r"> Kolom kanan</label>`,d=>{
      const out=[];const clean=t=>{let x=String(t).trim().replace(/[^\w.\u00C0-\uFFFF]/g,'_');if(/^\d/.test(x))x='_'+x;return x;};
      const add=(label,rg2)=>{const n=clean(label);if(validName(n))out.push({name:n,ref:'='+absRef(cur,rg2)});};
      if(d.querySelector('.t').checked)for(let c=g.c1;c<=g.c2;c++)add(cellVal(cur,g.r1,c),{r1:g.r1+1,r2:g.r2,c1:c,c2:c});
      if(d.querySelector('.b').checked)for(let c=g.c1;c<=g.c2;c++)add(cellVal(cur,g.r2,c),{r1:g.r1,r2:g.r2-1,c1:c,c2:c});
      if(d.querySelector('.l').checked)for(let r=g.r1;r<=g.r2;r++)add(cellVal(cur,r,g.c1),{r1:r,r2:r,c1:g.c1+1,c2:g.c2});
      if(d.querySelector('.r').checked)for(let r=g.r1;r<=g.r2;r++)add(cellVal(cur,r,g.c2),{r1:r,r2:r,c1:g.c1,c2:g.c2-1});
      if(!out.length){toast('Tidak ada label yang valid untuk dijadikan nama.');return false;}
      tx(()=>{const keep=wb.names.filter(x=>!out.some(o=>o.name.toLowerCase()===x.name.toLowerCase()));wb.names=[...keep,...out];},true);sync();toast(out.length+' nama dibuat');});
  }
  /* penelusuran */
  function refsOf(sh,r,c){
    const cl=sh.cells.get(key(r,c));if(!cl||cl.f===undefined||cl.f===null)return [];
    let ast;try{ast=parseFormula(cl.f);}catch(e){return [];}
    const out=[];
    const walk=n=>{
      if(n.k==='ref'){if(!n.sheet||n.sheet.toLowerCase()===sh.name.toLowerCase()){if((n.r2-n.r1+1)*(n.c2-n.c1+1)<=10000)out.push({r1:n.r1,c1:n.c1,r2:n.r2,c2:n.c2});}}
      else if(n.k==='name'){const nm=wb.names.find(x=>x.name.toLowerCase()===n.v.toLowerCase());const t=nm&&nameTarget(nm);if(t&&t.sh===sh)out.push({r1:t.r1,c1:t.c1,r2:t.r2,c2:t.c2});}
      else if(n.k==='bin'){walk(n.l);walk(n.r);}else if(n.k==='un'||n.k==='pct')walk(n.e);else if(n.k==='fn')n.a.forEach(walk);};
    walk(ast);return out;
  }
  function computeTrace(){
    const t=cur.tr,arrows=[];t.ver=wb.ver;
    if(t.prec){let fr=[{r:t.r,c:t.c}];const seen=new Set();
      for(let d=0;d<t.prec&&fr.length;d++){const nx=[];
        for(const f of fr){const k=key(f.r,f.c);if(seen.has(k))continue;seen.add(k);
          for(const rg2 of refsOf(cur,f.r,f.c)){arrows.push({from:rg2,to:f});
            if((rg2.r2-rg2.r1+1)*(rg2.c2-rg2.c1+1)<=100)for(let r=rg2.r1;r<=rg2.r2;r++)for(let c=rg2.c1;c<=rg2.c2;c++)if(cellAt(cur,r,c)&&cellAt(cur,r,c).f!=null)nx.push({r,c});}}
        fr=nx;}}
    if(t.dep){let fr=[{r:t.r,c:t.c}];const seen=new Set();
      const fcells=[];for(const [k,cl] of cur.cells)if(cl.f!==undefined&&cl.f!==null)fcells.push([Math.floor(k/MAXC),k%MAXC]);
      for(let d=0;d<t.dep&&fr.length;d++){const nx=[];
        for(const f of fr){const k=key(f.r,f.c);if(seen.has(k))continue;seen.add(k);
          for(const [r,c] of fcells){if(refsOf(cur,r,c).some(g=>f.r>=g.r1&&f.r<=g.r2&&f.c>=g.c1&&f.c<=g.c2)){arrows.push({from:{r1:f.r,c1:f.c,r2:f.r,c2:f.c},to:{r,c}});nx.push({r,c});}}}
        fr=nx;}}
    t.arrows=arrows;
  }
  function trace(kind){
    const s=cur.sel,m=mergeAt(s.ar,s.ac),r=m?m.r1:s.ar,c=m?m.c1:s.ac;
    if(kind==='prec'&&!refsOf(cur,r,c).length){toast('Sel ini tidak memiliki sel preseden pada lembar ini.');return;}
    if(!cur.tr||cur.tr.r!==r||cur.tr.c!==c)cur.tr={r,c,prec:0,dep:0,arrows:[],ver:-1};
    cur.tr[kind]++;cur.tr.ver=-1;computeTrace();
    if(kind==='dep'&&!cur.tr.arrows.some(a=>a.from.r1===r&&a.from.c1===c)&&cur.tr.dep===1)toast('Tidak ada sel dependen pada lembar ini.');
    schedule();
  }
  const clearArrows=k=>{if(!cur.tr)return;if(k==='all')cur.tr=null;else{cur.tr[k]=0;if(!cur.tr.prec&&!cur.tr.dep)cur.tr=null;else cur.tr.ver=-1;}schedule();};
  /* pemeriksaan galat */
  function errorCheck(){
    const errs=[];for(const [k,cl] of cur.cells){const r=Math.floor(k/MAXC),c=k%MAXC;const v=cellVal(cur,r,c);if(isErr(v))errs.push({r,c,e:v.code,f:cl.f!=null?'='+cl.f:''});}
    errs.sort((a,b)=>a.r-b.r||a.c-b.c);
    const d=dlg('Pemeriksaan Galat',`<div>${errs.length?errs.length+' sel berisi galat:':'Tidak ada galat yang ditemukan pada lembar ini.'}${wb.circ?'<br><b style="color:#c00">Terdeteksi referensi melingkar.</b>':''}</div>${errs.length?`<div class="fcl" style="height:160px;min-width:340px">${errs.map((x,i)=>`<div data-i="${i}">${addr(x.r,x.c)} — ${esc(x.e)} <span style="opacity:.7">${esc(x.f)}</span></div>`).join('')}</div>`:''}`,null,'Tutup',true);
    const l=d.querySelector('.fcl');if(l)l.onclick=e=>{const t=e.target.closest('[data-i]');if(!t)return;l.querySelectorAll('.on').forEach(x=>x.classList.remove('on'));t.classList.add('on');const x=errs[+t.dataset.i];select(x.r,x.c);};
  }
  /* evaluasi rumus */
  function evalDialog(){
    const s=cur.sel,m=mergeAt(s.ar,s.ac),r0=m?m.r1:s.ar,c0=m?m.c1:s.ac,cl=cellAt(cur,r0,c0);
    if(!cl||cl.f===undefined||cl.f===null){toast('Pilih sel yang berisi rumus.');return;}
    let ast;const init=()=>{try{ast=parseFormula(cl.f);}catch(e){ast=null;}};init();
    if(!ast){toast('Rumus tidak valid.');return;}
    const kids=n=>n.k==='bin'?[n.l,n.r]:n.k==='un'||n.k==='pct'?[n.e]:n.k==='fn'?n.a:[];
    const lit=n=>n.k==='num'||n.k==='str'||n.k==='bool'||n.k==='err'||n.k==='empty';
    const rangeRef=n=>n.k==='ref'&&!n.single;
    const next=n=>{if(lit(n)||rangeRef(n))return null;for(const c of kids(n)){const r=next(c);if(r)return r;}return n;};
    const PR={'=':1,'<>':1,'<':1,'>':1,'<=':1,'>=':1,'&':2,'+':3,'-':3,'*':4,'/':4,'^':5};
    const st=(n,par,mark)=>{let t;
      switch(n.k){
        case 'num':t=String(+n.v.toPrecision(12)).replace('.',',');break;
        case 'str':t='"'+n.v.replace(/"/g,'""')+'"';break;
        case 'bool':t=n.v?'TRUE':'FALSE';break;case 'err':t=n.v;break;case 'empty':t='';break;case 'name':t=n.v;break;
        case 'ref':t=(n.sheet?n.sheet+'!':'')+(n.single?addr(n.r1,n.c1):addr(n.r1,n.c1)+':'+addr(n.r2,n.c2));break;
        case 'un':t=n.o+st(n.e,6,mark);break;case 'pct':t=st(n.e,6,mark)+'%';break;
        case 'bin':{const p=PR[n.o];t=st(n.l,p,mark)+(p<=2?' '+n.o+' ':n.o)+st(n.r,p+.5,mark);if(p<par)t='('+t+')';break;}
        case 'fn':t=n.n+'('+n.a.map(x=>st(x,0,mark)).join(';')+')';break;
        default:t='?';}
      return n===mark?'\u0001'+t+'\u0002':t;};
    const toLit=v=>{if(isErr(v))return {k:'err',v:v.code};if(typeof v==='string')return {k:'str',v};if(typeof v==='boolean')return {k:'bool',v};if(v===null||v===undefined)return {k:'num',v:0};return {k:'num',v};};
    const d=dlg('Evaluasi Rumus',`<div class="r2"><span>Referensi:</span><b>${esc(q1(cur.name))}!${addr(r0,c0)}</b><span>Evaluasi:</span><div class="ev" style="border:1px solid #aaa;padding:6px;min-width:320px;min-height:34px;font-family:Consolas,monospace;background:#fff;word-break:break-all"></div></div><div style="display:flex;gap:6px"><button class="bt go">Evaluasi</button><button class="bt rs">Mulai Ulang</button></div><div style="font-size:11px;color:#666">Bagian bergaris bawah akan dievaluasi pada langkah berikutnya.</div>`,null,'Tutup',true);
    const box=d.querySelector('.ev'),go=d.querySelector('.go');
    const show=()=>{const n=next(ast);const h=esc(st(ast,0,n)).replace('\u0001','<u>').replace('\u0002','</u>');box.innerHTML=h||'&nbsp;';go.disabled=!n;};
    go.onclick=()=>{const n=next(ast);if(!n)return;let v;try{v=evalAst(n,{wb,sh:cur,cell:{r:r0,c:c0}});}catch(e){v=E.VAL;}
      if(v instanceof Rng){v=v.rows*v.cols===1?rawVal(wb,v.sh,v.r1,v.c1):E.VAL;}
      const L=toLit(v);for(const k of Object.keys(n))delete n[k];Object.assign(n,L);show();};
    d.querySelector('.rs').onclick=()=>{init();show();};show();
  }
  /* jendela pantau */
  function renderWatch(){
    if(!watchEl)return;
    const rows=wb.watch.map((w,i)=>{const sh=wb.sheets.find(x=>x.name===w.sh);let v='',f='',nm='';
      if(sh){const cl=sh.cells.get(key(w.r,w.c));v=fmtValue(rawVal(wb,sh,w.r,w.c),cl&&cl.st&&cl.st.fmt).t;f=cl&&cl.f!=null?'='+cl.f:'';
        const hit=wb.names.find(n=>{const t=nameTarget(n);return t&&t.sh===sh&&t.r1===w.r&&t.r2===w.r&&t.c1===w.c&&t.c2===w.c;});nm=hit?hit.name:'';}
      return `<tr data-i="${i}" style="cursor:pointer;${i===watchSel?'background:#d3e8d9':''}"><td>${esc(w.sh)}</td><td>${esc(nm)}</td><td>${addr(w.r,w.c)}</td><td>${esc(v)}</td><td>${esc(f)}</td></tr>`;}).join('');
    watchEl.querySelector('.wt').innerHTML=`<tr style="background:#f3f3f3"><th align="left">Lembar</th><th align="left">Nama</th><th align="left">Sel</th><th align="left">Nilai</th><th align="left">Rumus</th></tr>${rows||'<tr><td colspan="5" style="color:#888;padding:8px">Belum ada sel yang dipantau</td></tr>'}`;
  }
  function toggleWatch(){
    if(watchEl){watchEl.remove();watchEl=null;return;}
    watchEl=document.createElement('div');watchEl.className='selp';watchEl.style.cssText='width:min(540px,92%);top:auto;bottom:62px;right:22px';
    watchEl.innerHTML=`<div class="sh"><span>Jendela Pantau</span><span><button data-add="1" title="Tambah pantau">＋ Tambah</button><button data-del="1" title="Hapus pantau">🗑 Hapus</button><button data-x="1">✕</button></span></div><table class="wt" style="width:100%;border-collapse:collapse;font-size:12px"></table>`;
    watchEl.querySelectorAll('td,th');$('.gwrap').appendChild(watchEl);renderWatch();
    watchEl.onclick=e=>{
      if(e.target.dataset.x){watchEl.remove();watchEl=null;return;}
      if(e.target.dataset.add){const g=expandMerge(rg());const arr=[];for(let r=g.r1;r<=Math.min(g.r2,g.r1+99);r++)for(let c=g.c1;c<=Math.min(g.c2,g.c1+99);c++)if(!wb.watch.some(w=>w.sh===cur.name&&w.r===r&&w.c===c))arr.push({sh:cur.name,r,c});wb.watch=[...wb.watch,...arr];renderWatch();return;}
      if(e.target.dataset.del){if(watchSel>=0){wb.watch=wb.watch.filter((_,i)=>i!==watchSel);watchSel=-1;renderWatch();}return;}
      const tr=e.target.closest('[data-i]');if(tr){watchSel=+tr.dataset.i;renderWatch();}};
    watchEl.ondblclick=e=>{const tr=e.target.closest('[data-i]');if(!tr)return;const w=wb.watch[+tr.dataset.i],i=wb.sheets.findIndex(x=>x.name===w.sh);if(i>=0){if(i!==wb.active)switchSheet(i);select(w.r,w.c);}};
  }
  /* perhitungan */
  function calcNow(){wb.stale=false;wb.ver++;recalc(wb);updCalcLabel();sync();}
  function calcSheet(){
    wb.ver++;wb.circ=false;
    for(const [k,cl] of cur.cells)if(cl.f!==undefined&&cl.f!==null)cl.pass=undefined;
    for(const [k,cl] of cur.cells)if(cl.f!==undefined&&cl.f!==null)calcCell(wb,cur,cl,Math.floor(k/MAXC),k%MAXC);
    updCalcLabel();sync();
  }
  Object.assign(acts,{
    namemgr:nameManager,namesel:namesFromSel,
    trace(k){trace(k);},showf(){cur.showF=!cur.showF;sync();},evalf:evalDialog,watch:toggleWatch,
    calcnow:calcNow,calcsheet:calcSheet
  });
  const catMenu=c=>()=>fnItems(catList(c));
  Object.assign(menus,{
    cRec:()=>fnItems(wb.recent),cfin:catMenu('fin'),clog:catMenu('log'),ctxt:catMenu('txt'),cdt:catMenu('dt'),clk:catMenu('lk'),cmt:catMenu('mt'),
    cmore:()=>[{l:'Statistik',i:'∑',sub:()=>fnItems(catList('st'))},{l:'Informasi',i:'ⓘ',sub:()=>fnItems(catList('inf'))}],
    defn:()=>[I('Tentukan Nama…',()=>defineName(null),'▭'),I('Pengelola Nama…',nameManager,'▤')],
    usef:()=>wb.names.length?[...wb.names.map(n=>({l:n.name,i:'▭',a:()=>insertText(n.name)})),{sep:1},I('Pengelola Nama…',nameManager,'▤')]:[{l:'(belum ada nama)',dis:true,i:''},{sep:1},I('Pengelola Nama…',nameManager,'▤')],
    nbnames:()=>wb.names.length?wb.names.map(n=>({l:n.name,i:'▭',a:()=>gotoName(n.name)})):[{l:'(belum ada nama)',dis:true,i:''}],
    rmarr:()=>[I('Hapus Semua Panah',()=>clearArrows('all'),'✖',{dis:!cur.tr}),I('Hapus Panah Preseden',()=>clearArrows('prec'),'⇢',{dis:!cur.tr}),I('Hapus Panah Dependen',()=>clearArrows('dep'),'⇠',{dis:!cur.tr})],
    errchk:()=>[I('Periksa Galat…',errorCheck,'⚠'),I(wb.circ?'Referensi Melingkar terdeteksi':'Referensi Melingkar: tidak ada',()=>toast(wb.circ?'Ada rumus yang merujuk ke dirinya sendiri, langsung maupun tidak langsung.':'Tidak ada referensi melingkar.'),'↻')],
    calcopt:()=>[I('Otomatis',()=>{if(wb.manual){wb.manual=false;calcNow();}},'',{ck:!wb.manual}),I('Manual',()=>{wb.manual=true;sync();},'',{ck:wb.manual}),{sep:1},I('Hitung Sekarang (F9)',calcNow,'▦',{dis:!wb.manual}),I('Hitung Lembar (Shift+F9)',calcSheet,'▦')]
  });

  /* ===== DATA ===== */
  const txtin=$('.txtin');
  function putGrid(rows,r0,c0){ /* tulis array 2D teks ke lembar */
    tx(()=>{rows.forEach((row,i)=>row.forEach((t,j)=>{if(r0+i>=MAXR||c0+j>=MAXC)return;const old=cellAt(cur,r0+i,c0+j);putCell(cur,r0+i,c0+j,t===''?(old&&old.st?{v:null,st:old.st}:null):mkInput(String(t),old));}));});
    select(r0,c0,r0+rows.length-1,c0+Math.max(...rows.map(r=>r.length))-1);
  }
  txtin.onchange=async e=>{const f=e.target.files[0];e.target.value='';if(!f)return;
    try{const rows=parseCsv(await f.text());if(!rows.length){toast('Berkas kosong.');return;}const s=cur.sel;putGrid(rows,s.ar,s.ac);}catch(err){toast('Gagal membaca berkas: '+err.message);}};
  function htmlTables(txt){
    const doc=new DOMParser().parseFromString(txt,'text/html');
    return [...doc.querySelectorAll('table')].map(t=>[...t.rows].map(r=>[...r.cells].map(c=>c.textContent.replace(/\s+/g,' ').trim()))).filter(t=>t.length);
  }
  async function fetchTable(url,idx){
    const res=await fetch(url);if(!res.ok)throw new Error('HTTP '+res.status);
    const txt=await res.text();const ct=res.headers.get('content-type')||'';
    if(/html/i.test(ct)||/<table/i.test(txt)){const T=htmlTables(txt);if(!T.length)throw new Error('Tidak ada tabel pada halaman.');return {tables:T,rows:T[Math.min(idx||0,T.length-1)]};}
    return {tables:null,rows:parseCsv(txt)};
  }
  function fromWebDialog(){
    const d=dlg('Dari Web',`<div class="r2"><span>Alamat (URL):</span><input type="text" class="u" placeholder="https://contoh.com/data.csv" style="width:300px"></div><div class="tl" style="display:none"><span>Tabel:</span> <select class="ti"></select></div><div style="font-size:11px;color:#666;max-width:340px">Mendukung halaman HTML yang berisi tabel serta berkas CSV. Server harus mengizinkan akses lintas-situs (CORS), kalau tidak browser akan memblokirnya.</div>`,async d=>{
      const u=d.querySelector('.u').value.trim();if(!u)return false;
      const idx=+d.querySelector('.ti').value||0;
      try{
        const res=await fetchTable(u,idx);
        if(res.tables&&res.tables.length>1&&d.querySelector('.tl').style.display==='none'){
          d.querySelector('.tl').style.display='';d.querySelector('.ti').innerHTML=res.tables.map((t,i)=>`<option value="${i}">Tabel ${i+1} (${t.length}×${t[0].length})</option>`).join('');
          toast('Ditemukan '+res.tables.length+' tabel. Pilih satu lalu tekan OK lagi.');return false;}
        const s=cur.sel;putGrid(res.rows,s.ar,s.ac);
        wb.conns.push({name:'Koneksi '+(wb.conns.length+1),url:u,idx,sheet:cur.name,r:s.ar,c:s.ac,rows:res.rows.length,cols:Math.max(...res.rows.map(r=>r.length))});
        d.remove();
      }catch(e){toast('Gagal mengambil data: '+e.message+' (kemungkinan diblokir CORS)');}
      return false;});
  }
  async function refreshConn(cn){
    const res=await fetchTable(cn.url,cn.idx);const i=wb.sheets.findIndex(s=>s.name===cn.sheet);if(i<0)return;
    const keep=cur;cur=wb.sheets[i];
    try{putGrid(res.rows,cn.r,cn.c);}finally{cur=keep;}
    cn.rows=res.rows.length;cn.cols=Math.max(...res.rows.map(r=>r.length));
  }
  async function refreshAll(){
    let n=0,bad=0;
    for(const cn of wb.conns){try{await refreshConn(cn);n++;}catch(e){bad++;}}
    calcNow();
    toast(wb.conns.length?`${n} koneksi disegarkan${bad?`, ${bad} gagal`:''}.`:'Perhitungan disegarkan (tidak ada koneksi data).');
  }
  function connsDialog(){
    const body=()=>wb.conns.length?wb.conns.map((c,i)=>`<div class="si" data-i="${i}" style="display:flex;gap:8px;justify-content:space-between;border-bottom:1px solid #eee;padding:4px 0"><span><b>${esc(c.name)}</b><br><small>${esc(c.url)} → ${esc(c.sheet)}!${addr(c.r,c.c)}</small></span><span><button class="bt rf" data-i="${i}">Segarkan</button> <button class="bt dl" data-i="${i}">Hapus</button></span></div>`).join(''):'<i>Belum ada koneksi data.</i>';
    const d=dlg('Koneksi Buku Kerja',`<div class="cl" style="min-width:360px;max-height:240px;overflow:auto">${body()}</div>`,null,'Tutup',true);
    d.querySelector('.cl').onclick=async e=>{const i=e.target.dataset.i;if(i===undefined)return;
      if(e.target.classList.contains('dl')){wb.conns.splice(+i,1);}else{try{await refreshConn(wb.conns[+i]);toast('Disegarkan.');}catch(err){toast('Gagal: '+err.message);}}
      d.querySelector('.cl').innerHTML=body();};
  }
  /* urut bertingkat */
  function sortDialog(){
    const g=region();if(g.r1===g.r2&&g.c1===g.c2){toast('Pilih sel di dalam data yang akan diurutkan.');return;}
    const hdr=guessHeader(g);
    const cols=h=>Array.from({length:g.c2-g.c1+1},(_,i)=>`<option value="${g.c1+i}">${esc(h?(String(cellVal(cur,g.r1,g.c1+i)??'')||'Kolom '+colName(g.c1+i)):'Kolom '+colName(g.c1+i))}</option>`).join('');
    const d=dlg('Urutkan',`<label><input type="checkbox" class="hd" ${hdr?'checked':''}> Data saya memiliki header</label><div class="lv"></div><div style="display:flex;gap:6px"><button class="bt ad">＋ Tambah Tingkat</button><button class="bt rm">✕ Hapus Tingkat</button></div>`,d=>{
      const keys=[...d.querySelectorAll('.lvr')].map(r=>({col:+r.querySelector('.cl').value,asc:r.querySelector('.od').value==='1'}));
      sortRange(g,keys,true,d.querySelector('.hd').checked);});
    const lv=d.querySelector('.lv');
    const row=i=>`<div class="r2 lvr"><span>${i?'Kemudian menurut':'Urutkan menurut'}:</span><span><select class="cl">${cols(d.querySelector('.hd').checked)}</select> <select class="od"><option value="1">A ke Z / Terkecil</option><option value="0">Z ke A / Terbesar</option></select></span></div>`;
    const reset=()=>{const n=Math.max(1,lv.children.length);const sv=[...lv.querySelectorAll('.lvr')].map(r=>[r.querySelector('.cl').value,r.querySelector('.od').value]);lv.innerHTML=Array.from({length:n},(_,i)=>row(i)).join('');[...lv.querySelectorAll('.lvr')].forEach((r,i)=>{if(sv[i]){r.querySelector('.cl').value=sv[i][0];r.querySelector('.od').value=sv[i][1];}});};
    lv.innerHTML=row(0);
    d.querySelector('.ad').onclick=()=>{lv.insertAdjacentHTML('beforeend',row(lv.children.length));};
    d.querySelector('.rm').onclick=()=>{if(lv.children.length>1)lv.lastElementChild.remove();};
    d.querySelector('.hd').onchange=reset;
  }
  /* filter lanjutan */
  function advFilter(){
    const g=region();
    const d=dlg('Filter Lanjutan',`<label><input type="radio" name="m" value="in" checked> Saring daftar di tempat</label><label><input type="radio" name="m" value="cp"> Salin ke lokasi lain</label><div class="r2"><span>Rentang daftar:</span><input type="text" class="l" value="${rngStr(g)}"><span>Rentang kriteria:</span><input type="text" class="c" placeholder="mis. F1:G3"><span>Salin ke:</span><input type="text" class="t" placeholder="mis. J1"></div><label><input type="checkbox" class="u"> Hanya rekaman unik</label><div style="font-size:11px;color:#666;max-width:330px">Baris pertama rentang kriteria berisi nama kolom sama seperti header daftar. Kriteria pada baris yang sama berarti DAN, pada baris berbeda berarti ATAU (mis. &gt;100, =apel, *kata*).</div>`,d=>{
      const L=parseRng(d.querySelector('.l').value),C=d.querySelector('.c').value.trim()?parseRng(d.querySelector('.c').value):null;
      if(!L||L.r2<=L.r1){toast('Rentang daftar tidak valid.');return false;}
      if(d.querySelector('.c').value.trim()&&!C){toast('Rentang kriteria tidak valid.');return false;}
      const heads=[];for(let c=L.c1;c<=L.c2;c++)heads.push(String(cellVal(cur,L.r1,c)??'').trim().toLowerCase());
      const preds=[];
      if(C&&C.r2>C.r1){for(let r=C.r1+1;r<=C.r2;r++){const and=[];for(let c=C.c1;c<=C.c2;c++){const h=String(cellVal(cur,C.r1,c)??'').trim().toLowerCase(),ci=heads.indexOf(h),cv=cellVal(cur,r,c);if(ci<0||cv===null||cv==='')continue;and.push([L.c1+ci,critFn(cv)]);}preds.push(and);}}
      const ok=r=>!preds.length||preds.some(and=>and.every(([c,f])=>f(cellVal(cur,r,c))));
      const uniq=d.querySelector('.u').checked,seen=new Set(),match=[];
      for(let r=L.r1+1;r<=L.r2;r++){if(!ok(r))continue;if(uniq){const k=heads.map((_,i)=>String(cellVal(cur,r,L.c1+i))).join('\u0001');if(seen.has(k))continue;seen.add(k);}match.push(r);}
      if(d.querySelector('input[name=m]:checked').value==='in'){
        const hide=new Set();for(let r=L.r1+1;r<=L.r2;r++)if(!match.includes(r))hide.add(r);
        tx(()=>{cur.adv=hide;applyFilter(true);});toast(match.length+' dari '+(L.r2-L.r1)+' rekaman ditemukan.');
      }else{
        const T=parseAddr(d.querySelector('.t').value);if(!T){toast('Lokasi tujuan tidak valid.');return false;}
        tx(()=>{[L.r1,...match].forEach((r,i)=>{for(let c=L.c1;c<=L.c2;c++)putCell(cur,T.r+i,T.c+(c-L.c1),cloneCell(cellAt(cur,r,c)));});});
        select(T.r,T.c,T.r+match.length,T.c+(L.c2-L.c1));
      }});
  }
  function clearFilters(){tx(()=>{if(cur.filter)cur.filter.crit={};cur.adv=null;applyFilter(true);});}
  /* teks ke kolom */
  function splitText(t,o){
    const D=[];if(o.tab)D.push('\t');if(o.semi)D.push(';');if(o.comma)D.push(',');if(o.space)D.push(' ');if(o.other)D.push(...o.other.split(''));
    if(!D.length)return [t];
    const out=[];let cur2='',q=false,last=false;
    for(let i=0;i<t.length;i++){const ch=t[i];
      if(ch==='"'){if(q&&t[i+1]==='"'){cur2+='"';i++;}else q=!q;continue;}
      if(!q&&D.includes(ch)){if(o.cons&&last&&cur2==='')continue;out.push(cur2);cur2='';last=true;continue;}
      cur2+=ch;last=false;}
    out.push(cur2);return out;
  }
  function textToColumns(){
    const g=expandMerge(limitRange(rg()));
    if(g.c1!==g.c2){toast('Pilih satu kolom data untuk dipecah.');return;}
    const samples=[];for(let r=g.r1;r<=Math.min(g.r2,g.r1+4);r++)samples.push(String(cellVal(cur,r,g.c1)??''));
    const d=dlg('Teks ke Kolom',`<div style="font-weight:600">Pembatas</div><div style="display:flex;gap:12px;flex-wrap:wrap"><label><input type="checkbox" class="o-tab"> Tab</label><label><input type="checkbox" class="o-semi"> Titik koma</label><label><input type="checkbox" class="o-comma" checked> Koma</label><label><input type="checkbox" class="o-space"> Spasi</label><label>Lainnya: <input type="text" class="o-other" maxlength="3" style="width:40px"></label></div><label><input type="checkbox" class="o-cons"> Anggap pembatas berurutan sebagai satu</label><div>Pratinjau:</div><div class="pv" style="border:1px solid #bbb;background:#fff;max-height:130px;overflow:auto;font-size:12px"></div>`,d=>{
      const o=opts();tx(()=>{for(let r=g.r1;r<=g.r2;r++){const t=String(cellVal(cur,r,g.c1)??'');if(t==='')continue;const parts=splitText(t,o);parts.forEach((p,j)=>{if(g.c1+j>=MAXC)return;const old=cellAt(cur,r,g.c1+j);putCell(cur,r,g.c1+j,mkInput(p,old));});}});});
    const opts=()=>({tab:d.querySelector('.o-tab').checked,semi:d.querySelector('.o-semi').checked,comma:d.querySelector('.o-comma').checked,space:d.querySelector('.o-space').checked,other:d.querySelector('.o-other').value,cons:d.querySelector('.o-cons').checked});
    const pv=()=>{const o=opts();d.querySelector('.pv').innerHTML='<table style="border-collapse:collapse">'+samples.map(t=>'<tr>'+splitText(t,o).map(x=>`<td style="border:1px solid #ddd;padding:1px 6px">${esc(x)}</td>`).join('')+'</tr>').join('')+'</table>';};
    d.addEventListener('input',pv);d.addEventListener('change',pv);pv();
  }
  /* isi cepat */
  function flashFill(){
    const s=cur.sel,c=s.ac;
    let sc=-1;for(const cand of [c-1,c+1])if(cand>=0&&cand<MAXC&&hasVal(cur,s.ar,cand)||(cand>=0&&cand<MAXC&&hasVal(cur,s.ar-1,cand))){sc=cand;break;}
    if(sc<0){toast('Isi Cepat membutuhkan kolom data di sebelah kolom ini.');return;}
    const txt=(r,cc)=>{const v=cellVal(cur,r,cc);return v===null||v===undefined?'':typeof v==='number'?fmtValue(v,(stOf(r,cc)||{}).fmt).t:String(v);};
    let top=s.ar;while(top>0&&(hasVal(cur,top-1,sc)))top--;
    const ex=[];let r=top;
    for(;r<MAXR&&hasVal(cur,r,sc);r++){const cl=cellAt(cur,r,c);if(cl&&cl.v!==null&&cl.v!==undefined&&(cl.f===undefined||cl.f===null))ex.push(r);else if(ex.length)break;}
    const exs=ex.filter(rr=>hasVal(cur,rr,sc));
    if(!exs.length){toast('Ketik contoh hasil di sel pertama kolom ini, lalu jalankan Isi Cepat.');return;}
    const sec=sc>0?sc-1:sc+1;
    const words=x=>x.trim().split(/\s+/).filter(Boolean),cap=w=>w?w[0].toUpperCase()+w.slice(1).toLowerCase():w;
    const T=[];
    T.push(x=>x.toUpperCase(),x=>x.toLowerCase(),x=>words(x).map(cap).join(' '),x=>x.replace(/\s+/g,''),x=>(x.match(/\d+/g)||[]).join(''),x=>x.replace(/[^A-Za-z\u00C0-\uFFFF]/g,''));
    for(const k of [0,1,2,3,-1,-2])T.push(x=>{const w=words(x);return k>=0?w[k]:w[w.length+k];});
    T.push(x=>words(x).map(w=>w[0]).join(''),x=>words(x).map(w=>w[0]).join('').toUpperCase());
    for(let n=1;n<=8;n++){T.push(x=>x.slice(0,n),x=>x.slice(-n));}
    for(const sp of ['@',' ','-','/','.',',','_',':','(',')']){
      T.push(x=>x.includes(sp)?x.split(sp)[0]:undefined,x=>x.includes(sp)?x.slice(x.indexOf(sp)+1):undefined,x=>x.includes(sp)?x.split(sp).pop():undefined,x=>{const p=x.split(sp);return p.length>2?p[1]:undefined;});
    }
    T.push(x=>{const w=words(x);return w.length>1?w[w.length-1]+' '+w.slice(0,-1).join(' '):undefined;},x=>{const w=words(x);return w.length>1?w[w.length-1]+', '+w.slice(0,-1).join(' '):undefined;},x=>{const w=words(x);return w.length>1?w.slice(1).join(' ')+' '+w[0]:undefined;});
    const apply=(f,rr)=>{try{return f(txt(rr,sc),rr);}catch(e){return undefined;}};
    let fn=null;
    for(const f of T){if(exs.every(rr=>apply(f,rr)===txt(rr,c))){fn=f;break;}}
    if(!fn&&sec>=0&&sec<MAXC&&hasVal(cur,exs[0],sec)){
      for(const sep of ['',' ','-',', ','_','.','/','@',' - '])for(const ord of [0,1]){const f=(x,rr)=>{const a=txt(rr,sc),b=txt(rr,sec);return ord?b+sep+a:a+sep+b;};if(exs.every(rr=>f(null,rr)===txt(rr,c))){fn=f;break;}}
    }
    if(!fn){toast('Pola tidak dikenali. Ketik beberapa contoh lagi lalu coba kembali.');return;}
    let n=0;
    tx(()=>{for(let rr=top;rr<MAXR&&hasVal(cur,rr,sc);rr++){if(exs.includes(rr))continue;const cl=cellAt(cur,rr,c);if(cl&&cl.v!==null&&cl.v!==undefined)continue;const v=apply(fn,rr);if(v===undefined||v==='')continue;putCell(cur,rr,c,mkInput("'"+v,cl));n++;}});
    toast(n+' sel diisi.');
  }
  /* hapus duplikat */
  function removeDuplicates(){
    const g=region();if(g.r1===g.r2&&g.c1===g.c2){toast('Pilih sel di dalam data.');return;}
    const hdr=guessHeader(g);
    const d=dlg('Hapus Duplikat',`<label><input type="checkbox" class="hd" ${hdr?'checked':''}> Data saya memiliki header</label><div style="font-weight:600">Kolom yang dibandingkan:</div><div class="cl2 fl"></div>`,d=>{
      const h=d.querySelector('.hd').checked,cs=[...d.querySelectorAll('.cb:checked')].map(x=>+x.value);
      if(!cs.length){toast('Pilih minimal satu kolom.');return false;}
      const r1=g.r1+(h?1:0),seen=new Set(),keep=[];let rem=0;
      for(let r=r1;r<=g.r2;r++){const k=cs.map(c=>String(cellVal(cur,r,c)??'').toLowerCase()).join('\u0001');if(seen.has(k)){rem++;continue;}seen.add(k);keep.push(r);}
      if(!rem){toast('Tidak ada nilai duplikat.');return;}
      tx(()=>{const lines=keep.map(r=>{const l=[];for(let c=g.c1;c<=g.c2;c++)l.push([cellAt(cur,r,c),r]);return l;});
        lines.forEach((line,i)=>{const R=r1+i;line.forEach(([cl,from],j)=>{let nc=cl;if(cl&&cl.f!==undefined&&cl.f!==null&&R!==from)nc={...cl,f:shiftFormula(cl.f,R-from,0),ast:undefined};putCell(cur,R,g.c1+j,nc||null);});});
        for(let R=r1+keep.length;R<=g.r2;R++)for(let c=g.c1;c<=g.c2;c++)putCell(cur,R,c,null);});
      toast(`${rem} nilai duplikat ditemukan dan dihapus; ${keep.length} nilai unik tersisa.`);});
    const fill=()=>{const h=d.querySelector('.hd').checked;d.querySelector('.cl2').innerHTML=Array.from({length:g.c2-g.c1+1},(_,i)=>`<label><input type="checkbox" class="cb" value="${g.c1+i}" checked> ${esc(h?(String(cellVal(cur,g.r1,g.c1+i)??'')||'Kolom '+colName(g.c1+i)):'Kolom '+colName(g.c1+i))}</label>`).join('');};
    d.querySelector('.hd').onchange=fill;fill();
  }
  /* validasi data */
  const dvAt=(r,c)=>{for(let i=cur.dv.length-1;i>=0;i--){const x=cur.dv[i];if(r>=x.r1&&r<=x.r2&&c>=x.c1&&c<=x.c2)return x;}return null;};
  const DVOPS=[['between','Di antara'],['notBetween','Tidak di antara'],['equal','Sama dengan'],['notEqual','Tidak sama dengan'],['gt','Lebih besar dari'],['lt','Lebih kecil dari'],['ge','Lebih besar atau sama dengan'],['le','Lebih kecil atau sama dengan']];
  const DVTYPES=[['any','Nilai apa pun'],['whole','Bilangan bulat'],['decimal','Desimal'],['list','Daftar'],['date','Tanggal'],['len','Panjang teks'],['custom','Kustom (rumus)']];
  function dvItems(rule){
    const l=rule.list||'';
    if(l[0]==='='){try{const a=parseFormula(l.slice(1));let v=evalAst(a,{wb,sh:cur,cell:{r:rule.r1,c:rule.c1}});if(v instanceof Rng){const out=[];for(let r=v.r1;r<=v.r2;r++)for(let c=v.c1;c<=v.c2;c++){const x=rawVal(wb,v.sh,r,c);if(x!==null&&x!=='')out.push(typeof x==='number'?fmtValue(x,'General').t:String(x));}return out;}return [];}catch(e){return [];}}
    return l.split(/[;,]/).map(x=>x.trim()).filter(x=>x!=='');
  }
  function dvOperand(t,rule,r,c){
    if(t===''||t==null)return null;if(typeof t==='number')return t;
    const s=String(t);
    if(s[0]==='='){try{let v=evalAst(parseFormula(shiftFormula(s.slice(1),r-rule.r1,c-rule.c1)),{wb,sh:cur,cell:{r,c}});if(v instanceof Rng)v=rawVal(wb,v.sh,v.r1,v.c1);return typeof v==='number'?v:null;}catch(e){return null;}}
    const p=parseInput(s);return typeof p.v==='number'?p.v:null;
  }
  function dvCheck(r,c,text){
    const rule=dvAt(r,c);if(!rule||rule.type==='any')return null;
    if(text===''||text==null)return rule.blank===false?rule:null;
    if(text[0]==='='&&rule.type!=='custom')return null;
    const p=parseInput(text),v=p.v,t=String(text);
    const cmp=x=>{const a=dvOperand(rule.a,rule,r,c),b=dvOperand(rule.b,rule,r,c);if(a===null)return true;
      switch(rule.op){case 'between':return b===null?x>=a:x>=Math.min(a,b)&&x<=Math.max(a,b);case 'notBetween':return b===null?x<a:x<Math.min(a,b)||x>Math.max(a,b);case 'equal':return x===a;case 'notEqual':return x!==a;case 'gt':return x>a;case 'lt':return x<a;case 'ge':return x>=a;case 'le':return x<=a;}return true;};
    let ok=true;
    switch(rule.type){
      case 'list':ok=dvItems(rule).some(i=>i.toLowerCase()===t.trim().toLowerCase());break;
      case 'whole':ok=typeof v==='number'&&Number.isInteger(v)&&cmp(v);break;
      case 'decimal':case 'date':ok=typeof v==='number'&&cmp(v);break;
      case 'len':ok=cmp(t.length);break;
      case 'custom':{const old=cur.cells.get(key(r,c));cur.cells.set(key(r,c),{v:p.v,st:old&&old.st});
        try{let x=evalAst(parseFormula(shiftFormula(String(rule.a||'').replace(/^=/,''),r-rule.r1,c-rule.c1)),{wb,sh:cur,cell:{r,c}});if(x instanceof Rng)x=rawVal(wb,x.sh,x.r1,x.c1);ok=!isErr(x)&&!!(typeof x==='number'?x:x===true);}catch(e){ok=false;}
        if(old)cur.cells.set(key(r,c),old);else cur.cells.delete(key(r,c));break;}
    }
    return ok?null:rule;
  }
  function dvReject(rule,dr,dc){
    if(rule.showErr===false)return false;
    const title=rule.errTitle||'Microsoft Excel',msg=rule.errMsg||'Nilai ini tidak sesuai dengan batasan validasi data untuk sel ini.';
    if(rule.style==='information'){toast(msg);return false;}
    if(rule.style==='warning'){
      const d=dlg(esc(title),`<div style="max-width:320px">${esc(msg)}<br><br>Lanjutkan?</div>`,()=>{commitEdit(dr,dc,true);},'Ya');
      return true;}
    const d=dlg(esc(title),`<div style="max-width:320px">${esc(msg)}</div>`,()=>{if(ed)(ed.bar?fxi:edEl).focus();},'Coba Lagi',true);
    const cb=document.createElement('button');cb.textContent='Batal';cb.onclick=()=>{d.remove();cancelEdit();};d.querySelector('.dbt').prepend(cb);
    return true;
  }
  function dvDialog(){
    const g=expandMerge(limitRange(rg())),ex=dvAt(g.r1,g.c1)||{type:'any',op:'between',a:'',b:'',list:'',blank:true,showIn:true,inTitle:'',inMsg:'',showErr:true,style:'stop',errTitle:'',errMsg:''};
    const d=dlg('Validasi Data',`<div class="dtabs"><button data-t="s" class="on">Pengaturan</button><button data-t="i">Pesan Input</button><button data-t="e">Peringatan Kesalahan</button></div>
     <div class="dp s"><div class="r2"><span>Izinkan:</span><select class="ty">${DVTYPES.map(t=>`<option value="${t[0]}">${t[1]}</option>`).join('')}</select><span class="lo">Data:</span><select class="op">${DVOPS.map(t=>`<option value="${t[0]}">${t[1]}</option>`).join('')}</select><span class="la">Minimum:</span><input type="text" class="a"><span class="lb">Maksimum:</span><input type="text" class="b"><span class="ll">Sumber:</span><input type="text" class="l" placeholder="a;b;c atau =$A$1:$A$5"></div><label><input type="checkbox" class="bl"> Abaikan kosong</label></div>
     <div class="dp i" hidden><label><input type="checkbox" class="si"> Tampilkan pesan input saat sel dipilih</label><div class="r2"><span>Judul:</span><input type="text" class="it"><span>Pesan:</span><textarea class="im" rows="3"></textarea></div></div>
     <div class="dp e" hidden><label><input type="checkbox" class="se"> Tampilkan peringatan kesalahan setelah data tidak valid dimasukkan</label><div class="r2"><span>Gaya:</span><select class="st"><option value="stop">Stop</option><option value="warning">Peringatan</option><option value="information">Informasi</option></select><span>Judul:</span><input type="text" class="et"><span>Pesan:</span><textarea class="em" rows="3"></textarea></div></div>
     <div><button class="bt clr" style="margin-top:4px">Hapus Semua</button></div>`,d=>{
      const ty=d.querySelector('.ty').value;
      if(ty==='list'&&!d.querySelector('.l').value.trim()){toast('Isi sumber daftar.');return false;}
      if(!['any','list','custom'].includes(ty)&&d.querySelector('.a').value.trim()===''){toast('Isi nilai batasan.');return false;}
      const rule={r1:g.r1,c1:g.c1,r2:g.r2,c2:g.c2,type:ty,op:d.querySelector('.op').value,a:d.querySelector('.a').value.trim(),b:d.querySelector('.b').value.trim(),list:d.querySelector('.l').value.trim(),blank:d.querySelector('.bl').checked,showIn:d.querySelector('.si').checked,inTitle:d.querySelector('.it').value,inMsg:d.querySelector('.im').value,showErr:d.querySelector('.se').checked,style:d.querySelector('.st').value,errTitle:d.querySelector('.et').value,errMsg:d.querySelector('.em').value};
      tx(()=>{cur.dv=cur.dv.filter(x=>x.r2<g.r1||x.r1>g.r2||x.c2<g.c1||x.c1>g.c2);if(ty!=='any')cur.dv.push(rule);});});
    const q=s=>d.querySelector(s);
    q('.ty').value=ex.type;q('.op').value=ex.op;q('.a').value=ex.a;q('.b').value=ex.b;q('.l').value=ex.list;q('.bl').checked=ex.blank!==false;q('.si').checked=ex.showIn!==false;q('.it').value=ex.inTitle||'';q('.im').value=ex.inMsg||'';q('.se').checked=ex.showErr!==false;q('.st').value=ex.style||'stop';q('.et').value=ex.errTitle||'';q('.em').value=ex.errMsg||'';
    const upd=()=>{const t=q('.ty').value,two=['between','notBetween'].includes(q('.op').value);
      const show=(sel,on)=>{d.querySelectorAll(sel).forEach(e=>e.hidden=!on);};
      const cmpT=['whole','decimal','date','len'].includes(t);
      show('.lo,.op',cmpT);show('.la,.a',cmpT||t==='custom');show('.lb,.b',cmpT&&two);show('.ll,.l',t==='list');
      d.querySelector('.la').textContent=t==='custom'?'Rumus:':two?'Minimum:':'Nilai:';};
    d.querySelector('.dtabs').onclick=e=>{const b=e.target.closest('button');if(!b)return;d.querySelectorAll('.dtabs button').forEach(x=>x.classList.toggle('on',x===b));d.querySelectorAll('.dp').forEach(p=>p.hidden=!p.classList.contains(b.dataset.t));};
    d.addEventListener('change',upd);upd();
    q('.clr').onclick=()=>{tx(()=>{cur.dv=cur.dv.filter(x=>x.r2<g.r1||x.r1>g.r2||x.c2<g.c1||x.c1>g.c2);});d.remove();};
  }
  function dvDropdown(btn){
    const s=cur.sel,rule=dvAt(s.ar,s.ac);if(!rule)return;
    const items=dvItems(rule),r=btn.getBoundingClientRect();
    openMenuAt(r.left-120,r.bottom,items.length?items.map(it=>({l:it,i:'',a:()=>{const old=cellAt(cur,s.ar,s.ac);tx(()=>putCell(cur,s.ar,s.ac,mkInput(it,old)));}})):[{l:'(daftar kosong)',dis:true,i:''}],btn);
  }
  function dvCircle(on){
    if(on&&!cur.dv.length){toast('Belum ada aturan validasi data pada lembar ini.');return;}
    cur.circ=on;schedule();
  }
  /* konsolidasi */
  function consolidate(){
    const refs=[];
    const d=dlg('Konsolidasi',`<div class="r2"><span>Fungsi:</span><select class="fn"><option value="SUM">Jumlah</option><option value="COUNT">Hitung</option><option value="AVERAGE">Rata-rata</option><option value="MAX">Maks</option><option value="MIN">Min</option><option value="PRODUCT">Produk</option></select><span>Referensi:</span><span><input type="text" class="rf" placeholder="mis. Sheet2!A1:C5" style="width:190px"> <button class="bt ad">Tambah</button></span></div><div class="fl rl" style="min-height:70px"></div><div style="font-weight:600">Gunakan label di:</div><label><input type="checkbox" class="lt"> Baris atas</label><label><input type="checkbox" class="ll"> Kolom kiri</label><div style="font-size:11px;color:#666">Hasil ditulis mulai dari sel aktif (${addr(cur.sel.ar,cur.sel.ac)}).</div>`,d=>{
      if(!refs.length){toast('Tambahkan minimal satu referensi.');return false;}
      const fn=d.querySelector('.fn').value,top=d.querySelector('.lt').checked,left=d.querySelector('.ll').checked;
      const agg={SUM:v=>v.reduce((a,b)=>a+b,0),COUNT:v=>v.length,AVERAGE:v=>v.reduce((a,b)=>a+b,0)/v.length,MAX:v=>Math.max(...v),MIN:v=>Math.min(...v),PRODUCT:v=>v.reduce((a,b)=>a*b,1)}[fn];
      const rowKeys=[],colKeys=[],cells=new Map();
      const addK=(arr,k)=>{if(!arr.includes(k))arr.push(k);};
      for(const rf of refs){
        const sh=rf.sh,R=rf;
        for(let r=R.r1;r<=R.r2;r++)for(let c=R.c1;c<=R.c2;c++){
          if(top&&r===R.r1&&left&&c===R.c1)continue;
          if(top&&r===R.r1){addK(colKeys,String(rawVal(wb,sh,r,c)??''));continue;}
          if(left&&c===R.c1){addK(rowKeys,String(rawVal(wb,sh,r,c)??''));continue;}
          const v=rawVal(wb,sh,r,c);if(typeof v!=='number')continue;
          const rk=left?String(rawVal(wb,sh,r,R.c1)??''):'#'+(r-R.r1-(top?1:0));
          const ck=top?String(rawVal(wb,sh,R.r1,c)??''):'#'+(c-R.c1-(left?1:0));
          if(!left)addK(rowKeys,rk);if(!top)addK(colKeys,ck);
          const kk=rk+'\u0001'+ck;if(!cells.has(kk))cells.set(kk,[]);cells.get(kk).push(v);}
      }
      const s=cur.sel;
      tx(()=>{
        const wr=(r,c,v)=>{const old=cellAt(cur,r,c);putCell(cur,r,c,{v,st:old&&old.st});};
        const r0=s.ar+(top?1:0),c0=s.ac+(left?1:0);
        if(top)colKeys.forEach((k,j)=>wr(s.ar,c0+j,k));
        if(left)rowKeys.forEach((k,i)=>wr(r0+i,s.ac,k));
        rowKeys.forEach((rk,i)=>colKeys.forEach((ck,j)=>{const v=cells.get(rk+'\u0001'+ck);if(v&&v.length)wr(r0+i,c0+j,agg(v));}));
      });});
    const list=d.querySelector('.rl'),draw=()=>{list.innerHTML=refs.map((r,i)=>`<div data-i="${i}" style="display:flex;justify-content:space-between"><span>${esc(r.t)}</span><button class="bt" data-i="${i}">✕</button></div>`).join('')||'<i style="color:#888">Belum ada referensi</i>';};draw();
    d.querySelector('.ad').onclick=()=>{
      const t=d.querySelector('.rf').value.trim();let a;try{a=parseFormula(t);if(a.k!=='ref')throw 0;}catch(e){toast('Referensi tidak valid.');return;}
      const sh=a.sheet?wb.sheets.find(x=>x.name.toLowerCase()===a.sheet.toLowerCase()):cur;if(!sh){toast('Lembar tidak ditemukan.');return;}
      refs.push({t,sh,r1:a.r1,c1:a.c1,r2:a.r2,c2:a.c2});d.querySelector('.rf').value='';draw();};
    list.onclick=e=>{const i=e.target.dataset.i;if(i!==undefined&&e.target.tagName==='BUTTON'){refs.splice(+i,1);draw();}};
    const g=expandMerge(rg());if(!(g.r1===g.r2&&g.c1===g.c2))d.querySelector('.rf').value=rngStr(g);
  }
  /* analisis bagaimana-jika: pencarian sasaran */
  function goalSeek(){
    const s=cur.sel;
    dlg('Pencarian Sasaran',`<div class="r2"><span>Setel sel:</span><input type="text" class="a" value="${addr(s.ar,s.ac)}"><span>Ke nilai:</span><input type="text" class="v"><span>Dengan mengubah sel:</span><input type="text" class="c"></div>`,d=>{
      const A=parseAddr(d.querySelector('.a').value),C=parseAddr(d.querySelector('.c').value),T=num(d.querySelector('.v').value.replace(',','.'));
      if(!A||!C||isErr(T)){toast('Isi sel dan nilai sasaran dengan benar.');return false;}
      const fcl=cellAt(cur,A.r,A.c);if(!fcl||fcl.f===undefined||fcl.f===null){toast('Sel yang disetel harus berisi rumus.');return false;}
      const orig=cellAt(cur,C.r,C.c),k=key(C.r,C.c);
      const f=x=>{cur.cells.set(k,{v:x,st:orig&&orig.st});wb.ver++;recalc(wb);const y=cellVal(cur,A.r,A.c);return typeof y==='number'?y-T:NaN;};
      let x0=orig&&typeof orig.v==='number'?orig.v:0,x1=x0+(Math.abs(x0)>1?Math.abs(x0)*.1:1),f0=f(x0),f1=f(x1),sol=null;
      for(let i=0;i<80&&isFinite(f0)&&isFinite(f1);i++){
        if(Math.abs(f1)<1e-9){sol=x1;break;}
        if(f1===f0)break;
        const x2=x1-f1*(x1-x0)/(f1-f0);x0=x1;f0=f1;x1=x2;f1=f(x1);}
      if(sol===null&&isFinite(f1)&&Math.abs(f1)<1e-6)sol=x1;
      if(orig)cur.cells.set(k,orig);else cur.cells.delete(k);wb.ver++;recalc(wb);sync();
      if(sol===null){dlg('Status Pencarian Sasaran','<div>Pencarian sasaran mungkin tidak menemukan solusi.</div>',null,'OK',true);return;}
      const solR=+sol.toPrecision(10);
      dlg('Status Pencarian Sasaran',`<div>Pencarian sasaran dengan sel ${addr(A.r,A.c)} menemukan solusi.<br>Nilai sasaran: ${esc(String(T).replace('.',','))}<br>Nilai saat ini: ${esc(String(solR).replace('.',','))} (sel ${addr(C.r,C.c)})</div>`,()=>{tx(()=>putCell(cur,C.r,C.c,{v:solR,st:orig&&orig.st}));},'OK');
    },'OK');
  }
  /* outline: grup */
  const ogDepth=(list,g)=>list.filter(h=>h.a<=g.a&&h.b>=g.b).length;
  function applyOutline(){
    for(const [ax,list,k] of [[cur.rh,cur.og.rows,'hr'],[cur.cw,cur.og.cols,'hc']]){
      for(const i of cur.og[k])ax.hid.delete(i);
      const hs=new Set();for(const g of list)if(g.cl)for(let i=g.a;i<=Math.min(g.b,g.a+200000);i++)hs.add(i);
      hs.forEach(i=>ax.hid.add(i));cur.og[k]=hs;ax.dirty=true;
    }
  }
  function groupSel(kind){
    const g=expandMerge(rg());
    const run=k=>{
      const a=k==='r'?g.r1:g.c1,b=k==='r'?g.r2:g.c2,list=k==='r'?cur.og.rows:cur.og.cols;
      if(list.some(x=>x.a===a&&x.b===b)){toast('Rentang ini sudah menjadi grup.');return;}
      if(ogDepth(list,{a,b})>=8){toast('Tingkat outline maksimum adalah 8.');return;}
      tx(()=>{list.push({a,b,cl:false});});
    };
    if(kind){run(kind);return;}
    const wholeR=g.c1===0&&g.c2===MAXC-1,wholeC=g.r1===0&&g.r2===MAXR-1;
    if(wholeR&&!wholeC){run('r');return;}if(wholeC&&!wholeR){run('c');return;}
    dlg('Grup',`<label><input type="radio" name="k" value="r" checked> Baris</label><label><input type="radio" name="k" value="c"> Kolom</label>`,d=>{run(d.querySelector('input:checked').value);});
  }
  function ungroupSel(kind){
    const g=expandMerge(rg());
    const run=k=>{
      const a=k==='r'?g.r1:g.c1,b=k==='r'?g.r2:g.c2,list=k==='r'?cur.og.rows:cur.og.cols;
      let hit=list.filter(x=>x.a>=a&&x.b<=b);
      if(!hit.length){const c2=list.filter(x=>x.a<=a&&x.b>=b);if(c2.length){const m=Math.max(...c2.map(x=>ogDepth(list,x)));hit=c2.filter(x=>ogDepth(list,x)===m);}}
      if(!hit.length){toast('Tidak ada grup pada pilihan ini.');return;}
      const mx=Math.max(...hit.map(x=>ogDepth(list,x)));
      tx(()=>{const rm=hit.filter(x=>ogDepth(list,x)===mx);const keep=list.filter(x=>!rm.includes(x));if(k==='r')cur.og.rows=keep;else cur.og.cols=keep;applyOutline();});
    };
    if(kind){run(kind);return;}
    const wholeR=g.c1===0&&g.c2===MAXC-1,wholeC=g.r1===0&&g.r2===MAXR-1;
    if(wholeR&&!wholeC){run('r');return;}if(wholeC&&!wholeR){run('c');return;}
    dlg('Pisahkan Grup',`<label><input type="radio" name="k" value="r" checked> Baris</label><label><input type="radio" name="k" value="c"> Kolom</label>`,d=>{run(d.querySelector('input:checked').value);});
  }
  function clearOutline(){tx(()=>{cur.og.rows=[];cur.og.cols=[];applyOutline();});}
  function detail(show){
    const g=expandMerge(rg());let done=false;
    for(const [list,p] of [[cur.og.rows,g.r1],[cur.og.cols,g.c1]]){
      const cand=list.filter(x=>p>=x.a&&p<=x.b+1&&(show?x.cl:!x.cl));
      if(cand.length){const m=Math.max(...cand.map(x=>ogDepth(list,x)));const tg=cand.find(x=>ogDepth(list,x)===m);tx(()=>{tg.cl=!show;applyOutline();});done=true;break;}
    }
    if(!done)toast(show?'Tidak ada grup yang diciutkan pada pilihan ini.':'Pilih sel di dalam grup atau baris ringkasannya.');
  }
  function ogClick(v){
    const [k,i]=v.split(':');
    if(k==='L'){const n=+i;tx(()=>{for(const x of cur.og.rows)x.cl=ogDepth(cur.og.rows,x)>=n;for(const x of cur.og.cols)x.cl=ogDepth(cur.og.cols,x)>=n;applyOutline();},true);sync();return;}
    const list=k==='r'?cur.og.rows:cur.og.cols,x=list[+i];if(!x)return;
    tx(()=>{x.cl=!x.cl;applyOutline();},true);sync();
  }
  /* subtotal */
  function subtotalRows(g){const rows=[];for(let r=g.r1;r<=g.r2;r++){for(let c=g.c1;c<=g.c2;c++){const cl=cellAt(cur,r,c);if(cl&&cl.f&&/^SUBTOTAL\(/i.test(cl.f)){rows.push(r);break;}}}return rows;}
  function removeSubtotalsRaw(g){
    const rows=subtotalRows(g);
    for(let i=rows.length-1;i>=0;i--){structRaw('r',rows[i],1,true);g.r2--;}
    cur.og.rows=cur.og.rows.filter(x=>!(x.a>=g.r1&&x.b<=g.r2+1));applyOutline();
  }
  function subtotalDialog(){
    const g=region();if(g.r1===g.r2){toast('Pilih sel di dalam data yang memiliki header.');return;}
    const hn=c=>String(cellVal(cur,g.r1,c)??'')||('Kolom '+colName(c));
    const d=dlg('Subtotal',`<div class="r2"><span>Pada setiap perubahan di:</span><select class="kc">${Array.from({length:g.c2-g.c1+1},(_,i)=>`<option value="${g.c1+i}">${esc(hn(g.c1+i))}</option>`).join('')}</select><span>Gunakan fungsi:</span><select class="fn"><option value="SUM">Jumlah</option><option value="COUNT">Hitung</option><option value="AVERAGE">Rata-rata</option><option value="MAX">Maks</option><option value="MIN">Min</option><option value="PRODUCT">Produk</option></select></div><div style="font-weight:600">Tambahkan subtotal ke:</div><div class="fl cs"></div><label><input type="checkbox" class="rp" checked> Ganti subtotal yang ada</label><div><button class="bt ra">Hapus Semua</button></div><div style="font-size:11px;color:#666;max-width:320px">Data harus sudah diurutkan menurut kolom yang dipilih.</div>`,d=>{
      const kc=+d.querySelector('.kc').value,fn=d.querySelector('.fn').value,cols=[...d.querySelectorAll('.cb:checked')].map(x=>+x.value);
      if(!cols.length){toast('Pilih kolom yang akan disubtotal.');return false;}
      const FNN={SUM:9,COUNT:2,AVERAGE:1,MAX:4,MIN:5,PRODUCT:6},LBL={SUM:'Jumlah',COUNT:'Hitung',AVERAGE:'Rata-rata',MAX:'Maks',MIN:'Min',PRODUCT:'Produk'};
      tx(()=>{
        const G={...g};
        if(d.querySelector('.rp').checked)removeSubtotalsRaw(G);
        const groups=[];let a=G.r1+1;
        for(let r=G.r1+1;r<=G.r2;r++){const k=String(cellVal(cur,r,kc)??'');const nk=r<G.r2?String(cellVal(cur,r+1,kc)??''):null;if(r===G.r2||nk!==k){groups.push({a,b:r,k});a=r+1;}}
        let off=0;const put=(r,c,o)=>{const old=cellAt(cur,r,c);putCell(cur,r,c,{...o,st:(old&&old.st)||{b:1}});};
        for(const gr of groups){
          const pos=gr.b+1+off;structRaw('r',pos,1,false);
          put(pos,kc,{v:`${gr.k} ${LBL[fn]}`});
          for(const c of cols)put(pos,c,{f:`SUBTOTAL(${FNN[fn]},${addr(gr.a+off,c)}:${addr(gr.b+off,c)})`});
          cur.og.rows.push({a:gr.a+off,b:gr.b+off,cl:false});off++;
        }
        const gr=G.r2+1+off;structRaw('r',gr,1,false);
        put(gr,g.c1,{v:`Total ${LBL[fn]}`});
        for(const c of cols)put(gr,c,{f:`SUBTOTAL(${FNN[fn]},${addr(G.r1+1,c)}:${addr(gr-1,c)})`});
        cur.og.rows.push({a:G.r1+1,b:gr-1,cl:false});
        applyOutline();
      });});
    const fill=()=>{d.querySelector('.cs').innerHTML=Array.from({length:g.c2-g.c1+1},(_,i)=>`<label><input type="checkbox" class="cb" value="${g.c1+i}" ${typeof cellVal(cur,g.r1+1,g.c1+i)==='number'?'checked':''}> ${esc(hn(g.c1+i))}</label>`).join('');};fill();
    d.querySelector('.ra').onclick=()=>{tx(()=>{removeSubtotalsRaw({...g});});d.remove();};
  }
  Object.assign(acts,{
    fromtext(){txtin.click();},fromweb:fromWebDialog,conns:connsDialog,
    refreshall:refreshAll,
    sortaz(){quickSort(true);},sortza(){quickSort(false);},sortdlg:sortDialog,
    filter:toggleFilter,fclear:clearFilters,freapply(){tx(()=>applyFilter(true));},fadv:advFilter,
    t2c:textToColumns,flash:flashFill,rmdup:removeDuplicates,consol:consolidate,
    dvdlg:dvDialog,goal:goalSeek,
    grp(){groupSel();},ungrp(){ungroupSel();},subtotal:subtotalDialog,showdet(){detail(true);},hidedet(){detail(false);}
  });
  Object.assign(menus,{
    refresh:()=>[I('Segarkan Semua',refreshAll,'⟳'),I('Segarkan Koneksi',connsDialog,'⛓')],
    dval:()=>[I('Validasi Data…',dvDialog,'✔'),I('Lingkari Data Tidak Valid',()=>dvCircle(true),'◯'),I('Hapus Lingkaran Validasi',()=>dvCircle(false),'✖',{dis:!cur.circ})],
    whatif:()=>[{l:'Manajer Skenario…',i:'',dis:true},I('Pencarian Sasaran…',goalSeek,'🎯'),{l:'Tabel Data…',i:'',dis:true}],
    grpm:()=>[I('Grup…',()=>groupSel(),'▣'),I('Grup Baris',()=>groupSel('r'),'☰'),I('Grup Kolom',()=>groupSel('c'),'▥')],
    ungrpm:()=>[I('Pisahkan Grup…',()=>ungroupSel(),'▢'),I('Hapus Outline',clearOutline,'✖')]
  });

  /* ===== TINJAU: komentar, proteksi, pelacakan ===== */
  const hash=s=>{let h=5381;for(const ch of String(s))h=((h<<5)+h+ch.charCodeAt(0))>>>0;return h;};
  const isLocked=(sh,r,c)=>{if((sh.allow||[]).some(a=>r>=a.r1&&r<=a.r2&&c>=a.c1&&c<=a.c2))return false;const cl=sh.cells.get(key(r,c));return !(cl&&cl.st&&cl.st.lock===false);};
  const protMsg='Sel atau bagan yang Anda coba ubah berada di lembar yang dilindungi. Untuk mengubahnya, cabut proteksi lembar.';
  const protBlock=(flag,msg)=>{if(cur.prot&&!(flag&&cur.prot.o[flag])){toast(msg||'Operasi ini tidak diizinkan karena lembar dilindungi. Cabut proteksi lembar untuk melanjutkan.');return true;}return false;};
  const sbBlock=()=>{if(wb.protStruct){toast('Struktur buku kerja dilindungi. Cabut proteksi buku kerja untuk menambah, menghapus, mengganti nama, atau memindahkan lembar.');return true;}return false;};
  const sameVal=(a,b)=>{const ea=!a||((a.v===null||a.v===undefined)&&(a.f===undefined||a.f===null)),eb=!b||((b.v===null||b.v===undefined)&&(b.f===undefined||b.f===null));if(ea&&eb)return true;if(ea!==eb)return false;return a.v===b.v&&a.f===b.f;};
  const sameSt=(a,b)=>JSON.stringify((a&&a.st)||null)===JSON.stringify((b&&b.st)||null);
  function diffSheet(prev,now){
    const out=[];
    for(const [k,b] of now.cells){const a=prev.cells.get(k);if(a!==b&&(!a||!sameVal(a,b)||!sameSt(a,b)))out.push({k,a,b});}
    for(const [k,a] of prev.cells)if(!now.cells.has(k)&&(!sameVal(a,null)||!sameSt(a,null)))out.push({k,a,b:undefined});
    return out;
  }
  function txGuard(s){
    if(txStruct||!cur||s.sheets.length!==wb.sheets.length||s.active!==wb.active)return true;
    if(!cur.prot&&!wb.track)return true;
    const d=diffSheet(s.sheets[wb.active],cur);
    if(cur.prot){
      const bad=d.find(x=>{const r=Math.floor(x.k/MAXC),c=x.k%MAXC;if(!isLocked({cells:s.sheets[wb.active].cells,allow:cur.allow},r,c))return false;
        return !sameVal(x.a,x.b)||(!sameSt(x.a,x.b)&&!cur.prot.o.formatCells);});
      if(bad){restore(s);toast(protMsg);return false;}
    }
    if(wb.track&&!wb.trackPause)logChanges(d);
    return true;
  }
  function logChanges(d){
    let n=0;
    for(const x of d){
      if(sameVal(x.a,x.b))continue;if(++n>300)break;
      const r=Math.floor(x.k/MAXC),c=x.k%MAXC,nw=x.b?rawText(x.b):'',old=x.a?rawText(x.a):'';
      const ex=wb.changes.findIndex(e=>e.sh===cur.name&&e.r===r&&e.c===c);
      if(ex>=0){const e=wb.changes[ex];if(e.old===nw)wb.changes.splice(ex,1);else{e.nw=nw;e.time=new Date();}}
      else wb.changes.push({sh:cur.name,r,c,old,nw,time:new Date(),user:'Pengguna'});
    }
  }
  /* komentar */
  const cmAt=(r,c)=>cur.cm.find(x=>x.r===r&&x.c===c);
  function commentDialog(){
    const s=cur.sel,m=mergeAt(s.ar,s.ac),r=m?m.r1:s.ar,c=m?m.c1:s.ac,ex=cmAt(r,c);
    if(cur.prot&&isLocked(cur,r,c)&&!cur.prot.o.objects){toast(protMsg);return;}
    const d=dlg((ex?'Edit Komentar ':'Komentar Baru ')+addr(r,c),`<textarea class="t" rows="5" style="width:320px" placeholder="Tulis komentar…">${esc(ex?ex.text:'')}</textarea>`,d=>{
      const t=d.querySelector('.t').value;if(!t.trim()){toast('Komentar tidak boleh kosong.');return false;}
      tx(()=>{cur.cm=ex?cur.cm.map(x=>x===ex?{...x,text:t}:x):[...cur.cm,{r,c,text:t,author:'Pengguna',w:200,h:70,show:false}];});});
    d.querySelector('.t').focus();
  }
  function commentDelete(){
    const s=cur.sel,ex=cmAt(s.ar,s.ac);if(!ex){toast('Tidak ada komentar pada sel ini.');return;}
    tx(()=>{cur.cm=cur.cm.filter(x=>x!==ex);});
  }
  function commentNav(dir){
    if(!cur.cm.length){toast('Tidak ada komentar pada lembar ini.');return;}
    const L=[...cur.cm].sort((a,b)=>a.r-b.r||a.c-b.c),s=cur.sel;
    let i=L.findIndex(x=>x.r===s.ar&&x.c===s.ac);
    if(i<0){i=dir>0?L.findIndex(x=>x.r>s.ar||(x.r===s.ar&&x.c>s.ac)):(()=>{for(let k=L.length-1;k>=0;k--)if(L[k].r<s.ar||(L[k].r===s.ar&&L[k].c<s.ac))return k;return -1;})();if(i<0)i=dir>0?0:L.length-1;}
    else i=(i+dir+L.length)%L.length;
    select(L[i].r,L[i].c);
  }
  function commentToggle(){
    const s=cur.sel,ex=cmAt(s.ar,s.ac);if(!ex){toast('Pilih sel yang memiliki komentar.');return;}
    cur.cm=cur.cm.map(x=>x===ex?{...x,show:!x.show}:x);sync();
  }
  /* proteksi lembar */
  const PROTOPTS=[['formatCells','Format sel'],['formatCols','Format kolom'],['formatRows','Format baris'],['insCols','Sisipkan kolom'],['insRows','Sisipkan baris'],['delCols','Hapus kolom'],['delRows','Hapus baris'],['sort','Urutkan'],['filter','Gunakan AutoFilter'],['objects','Edit objek']];
  function protectSheet(){
    if(cur.prot){
      const pw=cur.prot.pw;
      if(pw){dlg('Cabut Proteksi Lembar',`<label>Kata sandi: <input type="password" class="p"></label>`,d=>{if(hash(d.querySelector('.p').value)!==cur.prot.h){toast('Kata sandi salah.');return false;}tx(()=>{cur.prot=null;},true);sync();});}
      else{tx(()=>{cur.prot=null;},true);sync();}
      return;
    }
    dlg('Lindungi Lembar',`<label>Kata sandi (opsional): <input type="password" class="p"></label><div style="font-weight:600">Izinkan semua pengguna lembar ini untuk:</div><div class="fl">${PROTOPTS.map(([k,l])=>`<label><input type="checkbox" data-k="${k}"> ${l}</label>`).join('')}</div><div style="font-size:11px;color:#666;max-width:320px">Sel dikunci secara bawaan. Buka kunci sel lewat Format Sel → tab Proteksi sebelum melindungi lembar. Proteksi ini bersifat perlindungan ringan dan tidak menggantikan keamanan berkas.</div>`,d=>{
      const o={};d.querySelectorAll('[data-k]').forEach(x=>o[x.dataset.k]=x.checked);
      const pw=d.querySelector('.p').value;
      tx(()=>{cur.prot={pw:pw||null,h:pw?hash(pw):null,o};},true);sync();toast('Lembar dilindungi.');});
  }
  function protectWb(){
    if(wb.protStruct){
      const p=wb.protStruct;
      if(p.h!==null){dlg('Cabut Proteksi Buku Kerja',`<label>Kata sandi: <input type="password" class="p"></label>`,d=>{if(hash(d.querySelector('.p').value)!==p.h){toast('Kata sandi salah.');return false;}wb.protStruct=null;sync();toast('Proteksi buku kerja dicabut.');});}
      else{wb.protStruct=null;sync();toast('Proteksi buku kerja dicabut.');}
      return;
    }
    dlg('Lindungi Struktur dan Jendela',`<label><input type="checkbox" checked disabled> Struktur</label><div style="font-size:11px;color:#666">Mencegah penambahan, penghapusan, penggantian nama, atau pemindahan lembar.</div><label>Kata sandi (opsional): <input type="password" class="p"></label>`,d=>{
      const pw=d.querySelector('.p').value;wb.protStruct={h:pw?hash(pw):null};sync();toast('Struktur buku kerja dilindungi (berlaku selama sesi ini).');});
  }
  function allowRanges(){
    const body=()=>(cur.allow||[]).length?cur.allow.map((a,i)=>`<div class="si" style="display:flex;justify-content:space-between;border-bottom:1px solid #eee;padding:3px 0"><span><b>${esc(a.title)}</b> — ${addr(a.r1,a.c1)}${a.r1===a.r2&&a.c1===a.c2?'':':'+addr(a.r2,a.c2)}</span><button class="bt" data-i="${i}">Hapus</button></div>`).join(''):'<i>Belum ada rentang</i>';
    const g=expandMerge(rg());
    const d=dlg('Izinkan Pengguna Mengedit Rentang',`<div class="cl" style="min-width:320px;max-height:140px;overflow:auto"></div><div class="r2"><span>Judul:</span><input type="text" class="ti"><span>Merujuk ke sel:</span><input type="text" class="rg" value="${rngStr(g)}"></div><button class="bt ad">Tambah</button><div style="font-size:11px;color:#666;max-width:330px">Sel di dalam rentang ini tetap bisa diedit saat lembar dilindungi.</div>`,null,'Tutup',true);
    const cl=d.querySelector('.cl'),draw=()=>{cl.innerHTML=body();};draw();
    d.querySelector('.ad').onclick=()=>{const q=parseRng(d.querySelector('.rg').value);if(!q){toast('Rentang tidak valid.');return;}const t=d.querySelector('.ti').value.trim()||('Rentang'+((cur.allow||[]).length+1));tx(()=>{cur.allow=[...(cur.allow||[]),{title:t,...q}];},true);draw();};
    cl.onclick=e=>{const i=e.target.dataset.i;if(i===undefined)return;tx(()=>{cur.allow=cur.allow.filter((_,k)=>k!==+i);},true);draw();};
  }
  function changesDialog(){
    const mine=()=>wb.changes.map((e,i)=>({e,i}));
    const body=()=>mine().length?`<table style="border-collapse:collapse;font-size:12px;width:100%"><tr style="background:#f3f3f3"><th align="left">Lembar</th><th align="left">Sel</th><th align="left">Lama</th><th align="left">Baru</th><th align="left">Waktu</th></tr>${mine().map(({e,i})=>`<tr data-i="${i}" style="cursor:pointer"><td>${esc(e.sh)}</td><td>${addr(e.r,e.c)}</td><td>${esc(e.old)}</td><td>${esc(e.nw)}</td><td>${e.time.toLocaleTimeString('id-ID')}</td></tr>`).join('')}</table>`:'<i>Tidak ada perubahan yang dilacak.</i>';
    let sel=-1;
    const d=dlg('Terima atau Tolak Perubahan',`<div class="cl" style="min-width:420px;max-height:220px;overflow:auto;border:1px solid #bbb"></div><div style="display:flex;gap:6px;flex-wrap:wrap"><button class="bt ac">Terima</button><button class="bt rj">Tolak</button><button class="bt aa">Terima Semua</button><button class="bt ra">Tolak Semua</button></div>`,null,'Tutup',true);
    const cl=d.querySelector('.cl'),draw=()=>{cl.innerHTML=body();cl.querySelectorAll('tr[data-i]').forEach(tr=>{if(+tr.dataset.i===sel)tr.style.background='#d3e8d9';});};draw();
    cl.onclick=e=>{const tr=e.target.closest('[data-i]');if(tr){sel=+tr.dataset.i;const c=wb.changes[sel];const i=wb.sheets.findIndex(s=>s.name===c.sh);if(i>=0){if(i!==wb.active)switchSheet(i);select(c.r,c.c);}draw();}};
    const reject=list=>{wb.trackPause=true;try{for(const e of list){const sh=wb.sheets.find(s=>s.name===e.sh);if(!sh)continue;const keep=cur;cur=sh;try{tx(()=>{const old=cellAt(cur,e.r,e.c);putCell(cur,e.r,e.c,mkInput(e.old,old));},true);}finally{cur=keep;}}}finally{wb.trackPause=false;}wb.changes=wb.changes.filter(x=>!list.includes(x));sync();};
    d.querySelector('.ac').onclick=()=>{if(sel<0){toast('Pilih perubahan terlebih dahulu.');return;}wb.changes.splice(sel,1);sel=-1;draw();sync();};
    d.querySelector('.rj').onclick=()=>{if(sel<0){toast('Pilih perubahan terlebih dahulu.');return;}reject([wb.changes[sel]]);sel=-1;draw();};
    d.querySelector('.aa').onclick=()=>{wb.changes=[];sel=-1;draw();sync();};
    d.querySelector('.ra').onclick=()=>{reject([...wb.changes]);sel=-1;draw();};
  }
  /* tooltip komentar / perubahan */
  const tipEl=document.createElement('div');tipEl.className='cbox';tipEl.hidden=true;tipEl.style.cssText='position:absolute;z-index:15;pointer-events:none';$('.gwrap').appendChild(tipEl);
  let lastHover='';
  vp.addEventListener('pointermove',e=>{
    if(drag||ed||e.pointerType!=='mouse')return;
    const rc=vp.getBoundingClientRect();if(e.clientX-rc.left>vp.clientWidth||e.clientY-rc.top>vp.clientHeight){tipEl.hidden=true;return;}
    const h=hit(e),k=h.r+','+h.c;if(k===lastHover)return;lastHover=k;
    const m=mergeAt(h.r,h.c),r=m?m.r1:h.r,c=m?m.c1:h.c;
    const cm=cmAt(r,c),ch=wb.changes.find(x=>x.sh===cur.name&&x.r===r&&x.c===c);
    let html='';
    if(cm&&!(cm.show||cur.showAllCm))html+=`<b>${esc(cm.author)}:</b><br>${esc(cm.text)}`;
    if(ch)html+=(html?'<hr style="margin:4px 0">':'')+`<b>${esc(ch.user)}, ${ch.time.toLocaleString('id-ID')}:</b><br>Sel ${addr(r,c)} diubah dari "${esc(ch.old)}" menjadi "${esc(ch.nw)}".`;
    if(!html){tipEl.hidden=true;return;}
    tipEl.innerHTML=html;tipEl.style.width='200px';
    tipEl.style.left=((cur.head===false?0:HW+gutR)+cx(c+1)-vp.scrollLeft+8)+'px';tipEl.style.top=((cur.head===false?0:HH+gutC)+ry(r)-vp.scrollTop)+'px';tipEl.hidden=false;
  });
  vp.addEventListener('pointerleave',()=>{tipEl.hidden=true;lastHover='';});
  vp.addEventListener('scroll',()=>{tipEl.hidden=true;lastHover='';});
  Object.assign(acts,{
    cmnew:commentDialog,cmdel:commentDelete,cmprev(){commentNav(-1);},cmnext(){commentNav(1);},cmshow:commentToggle,
    cmall(){cur.showAllCm=!cur.showAllCm;sync();},
    protsheet:protectSheet,protwb:protectWb,allowr:allowRanges
  });
  Object.assign(menus,{
    track:()=>[I('Lacak Perubahan saat Mengedit',()=>{wb.track=!wb.track;toast(wb.track?'Pelacakan perubahan aktif. Perubahan sel akan disorot.':'Pelacakan perubahan dimatikan.');sync();},'',{ck:wb.track}),I('Terima atau Tolak Perubahan…',changesDialog,'✔'),{sep:1},I('Hapus Riwayat Perubahan',()=>{wb.changes=[];sync();},'✖',{dis:!wb.changes.length})]
  });

  /* ===== TAMPILAN ===== */
  const frzOf=()=>{const f=cur.freeze||{r:0,c:0};return {r:f.r||0,c:f.c||0,W:f.c?cx(f.c):0,H:f.r?ry(f.r):0};};
  function setFreeze(r,c){
    tx(()=>{cur.freeze=(r||c)?{r,c}:null;},true);
    vp.scrollLeft=0;vp.scrollTop=0;sync();
  }
  function pbpOverlay(W,H,mapX,mapY){
    const P=cur.page,u=usedRange(cur);
    const g=P.area?{...P.area}:(u.any?{r1:0,c1:0,r2:u.r2,c2:u.c2}:{r1:0,c1:0,r2:0,c2:0});
    let [pw,ph]=PAPER[P.paper]||PAPER.A4;if(P.orient==='landscape')[pw,ph]=[ph,pw];
    let totW=0,totH=0;for(let c=g.c1;c<=g.c2;c++)totW+=cur.cw.size(c);for(let r=g.r1;r<=g.r2;r++)totH+=cur.rh.size(r);
    const aW0=(pw/25.4-P.m.l-P.m.r)*96,aH0=(ph/25.4-P.m.t-P.m.b)*96;
    let zoom=P.scale/100;if(P.fitW||P.fitH)zoom=Math.min(1,P.fitW?aW0*P.fitW/Math.max(1,totW):Infinity,P.fitH?aH0*P.fitH/Math.max(1,totH):Infinity);
    const aW=aW0/zoom,aH=aH0/zoom,colB=[g.c1],rowB=[g.r1];
    let acc=0;for(let c=g.c1;c<=g.c2;c++){const w=cur.cw.size(c);if(acc+w>aW+.5&&acc>0){colB.push(c);acc=0;}acc+=w;}
    acc=0;for(let r=g.r1;r<=g.r2;r++){const h=cur.rh.size(r);if(acc+h>aH+.5&&acc>0){rowB.push(r);acc=0;}acc+=h;}
    colB.push(g.c2+1);rowB.push(g.r2+1);
    const X=c=>mapX(cx(c)),Y=r=>mapY(ry(r));
    const x1=X(g.c1),x2=X(g.c2+1),y1=Y(g.r1),y2=Y(g.r2+1);
    const R=(x,y,w,h,st)=>w>0&&h>0?`<div style="position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;pointer-events:none;${st}"></div>`:'';
    const gray='background:rgba(110,110,110,.45)';
    let h=R(0,0,W,Math.max(0,y1),gray)+R(0,y2,W,Math.max(0,H-y2),gray)+R(0,Math.max(0,y1),Math.max(0,x1),y2-Math.max(0,y1),gray)+R(x2,Math.max(0,y1),Math.max(0,W-x2),y2-Math.max(0,y1),gray);
    for(let i=1;i<colB.length-1;i++)h+=R(X(colB[i])-1,y1,2,y2-y1,'background:repeating-linear-gradient(to bottom,#1f4e9e 0 7px,transparent 7px 11px)');
    for(let j=1;j<rowB.length-1;j++)h+=R(x1,Y(rowB[j])-1,x2-x1,2,'background:repeating-linear-gradient(to right,#1f4e9e 0 7px,transparent 7px 11px)');
    h+=R(x1-1,y1-1,x2-x1+1,y2-y1+1,'border:2px solid #1f4e9e');
    let n=1;
    for(let i=0;i<colB.length-1;i++)for(let j=0;j<rowB.length-1;j++){
      const a=X(colB[i]),b=X(colB[i+1]),c=Y(rowB[j]),d=Y(rowB[j+1]);
      if(b>0&&a<W&&d>0&&c<H&&b>a&&d>c){const fs=Math.max(12,Math.min(64,(b-a)/9,(d-c)/6));const ty=Math.max(c,0),th=Math.min(d,H)-ty,tx2=Math.max(a,0),tw=Math.min(b,W)-tx2;h+=`<div style="position:absolute;left:${tx2}px;top:${ty}px;width:${tw}px;height:${th}px;display:flex;align-items:center;justify-content:center;font-size:${fs}px;white-space:nowrap;color:rgba(110,110,110,.45);font-weight:700;pointer-events:none;overflow:hidden">Halaman ${n}</div>`;}
      n++;
    }
    return h;
  }
  function zoomDialog(){
    const cur0=Math.round(Z*100);
    dlg('Zoom',`<label><input type="radio" name="z" value="200"> 200%</label><label><input type="radio" name="z" value="100"> 100%</label><label><input type="radio" name="z" value="75"> 75%</label><label><input type="radio" name="z" value="50"> 50%</label><label><input type="radio" name="z" value="25"> 25%</label><label><input type="radio" name="z" value="fit"> Muatkan ke seleksi</label><label><input type="radio" name="z" value="c"> Kustom: <input type="number" class="cz" min="20" max="200" value="${cur0}" style="width:70px"> %</label>`,d=>{
      const v=d.querySelector('input[name=z]:checked');if(!v)return;
      if(v.value==='fit')zoomSel();else if(v.value==='c')setZoom(Math.max(20,Math.min(200,+d.querySelector('.cz').value||100)));else setZoom(+v.value);});
    const pre=document.querySelector(`.xl .ov input[name=z][value="${[200,100,75,50,25].includes(cur0)?cur0:'c'}"]`);if(pre)pre.checked=true;
  }
  function zoomSel(){
    const g=expandMerge(rg());
    const w=cur.cw.pos(g.c2+1)-cur.cw.pos(g.c1),h=cur.rh.pos(g.r2+1)-cur.rh.pos(g.r1);
    const vw=vp.clientWidth-20,vh=vp.clientHeight-20;
    const z=Math.max(20,Math.min(200,Math.floor(Math.min(vw/Math.max(w,1),vh/Math.max(h,1))*100/5)*5));
    setZoom(z);vp.scrollLeft=cx(g.c1);vp.scrollTop=ry(g.r1);schedule();
  }
  function customViews(){
    const body=()=>wb.cviews.length?wb.cviews.map((v,i)=>`<div class="si" data-i="${i}" style="display:flex;justify-content:space-between;gap:8px;border-bottom:1px solid #eee;padding:3px 0;cursor:pointer"><span>${esc(v.name)} <small style="color:#888">(${esc(v.sheet)}, ${v.z}%)</small></span></div>`).join(''):'<i>Belum ada tampilan kustom.</i>';
    let sel=-1;
    const d=dlg('Tampilan Kustom',`<div class="cl" style="min-width:320px;max-height:150px;overflow:auto;border:1px solid #bbb;padding:4px"></div><div class="r2"><span>Nama:</span><input type="text" class="nm"></div><div style="display:flex;gap:6px"><button class="bt ad">Tambah</button><button class="bt sh">Tampilkan</button><button class="bt dl">Hapus</button></div><div style="font-size:11px;color:#666;max-width:330px">Menyimpan zoom, garis kisi, tajuk, bilah rumus, baris/kolom tersembunyi, panel beku, dan posisi sel untuk lembar aktif.</div>`,null,'Tutup',true);
    const cl=d.querySelector('.cl'),draw=()=>{cl.innerHTML=body();cl.querySelectorAll('[data-i]').forEach(x=>{if(+x.dataset.i===sel)x.style.background='#d3e8d9';});};draw();
    cl.onclick=e=>{const t=e.target.closest('[data-i]');if(t){sel=+t.dataset.i;draw();}};
    d.querySelector('.ad').onclick=()=>{
      const n=d.querySelector('.nm').value.trim();if(!n){toast('Isi nama tampilan.');return;}
      const v={name:n,sheet:cur.name,z:Math.round(Z*100),grid:cur.grid!==false,head:cur.head!==false,fbar:wb.fbar!==false,showF:!!cur.showF,hr:[...cur.rh.hid],hc:[...cur.cw.hid],freeze:cur.freeze?{...cur.freeze}:null,sel:{...cur.sel},mode:wb.viewMode||'normal'};
      wb.cviews=[...wb.cviews.filter(x=>x.name.toLowerCase()!==n.toLowerCase()),v];d.querySelector('.nm').value='';draw();toast('Tampilan "'+n+'" disimpan.');};
    d.querySelector('.sh').onclick=()=>{
      if(sel<0){toast('Pilih tampilan terlebih dahulu.');return;}
      const v=wb.cviews[sel],i=wb.sheets.findIndex(s=>s.name===v.sheet);if(i<0){toast('Lembar untuk tampilan ini sudah tidak ada.');return;}
      if(i!==wb.active)switchSheet(i);
      cur.grid=v.grid;cur.head=v.head;cur.showF=v.showF;wb.fbar=v.fbar;wb.viewMode=v.mode;cur.freeze=v.freeze?{...v.freeze}:null;
      cur.rh.hid=new Set(v.hr);cur.cw.hid=new Set(v.hc);cur.rh.dirty=true;cur.cw.dirty=true;
      cur.sel={...v.sel};setZoom(v.z);vp.scrollLeft=0;vp.scrollTop=0;sync();reveal(v.sel.ar,v.sel.ac);d.remove();};
    d.querySelector('.dl').onclick=()=>{if(sel<0)return;wb.cviews.splice(sel,1);sel=-1;draw();};
  }
  const applyFbar=()=>{$('.fbar').hidden=wb.fbar===false;};
  Object.assign(sels,{fbar:v=>{wb.fbar=!!v;applyFbar();sync();}});
  Object.assign(acts,{
    view(v){wb.viewMode=v;sync();},
    zoomdlg:zoomDialog,zoom100(){setZoom(100);},zoomsel:zoomSel,cviews:customViews
  });
  Object.assign(menus,{
    freeze:()=>{const f=frzOf(),on=f.r||f.c;return [on?I('Lepas Pembekuan Panel',()=>setFreeze(0,0),'❄'):I('Bekukan Panel',()=>{const s=cur.sel;setFreeze(s.ar,s.ac);},'❄'),I('Bekukan Baris Teratas',()=>setFreeze(1,0),'☰'),I('Bekukan Kolom Pertama',()=>setFreeze(0,1),'▥')];},
    switchw:()=>[I(name,()=>{},'▦',{ck:true})]
  });
  applyFbar();

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
    try{const nm=[];for(const d of (w.definedNames.model||[])){if(!d.name||!d.ranges||!d.ranges.length||/^_xlnm/.test(d.name))continue;const rr=d.ranges[0];if(/^'?[^!]+'?!\$?[A-Za-z]+\$?\d+(:\$?[A-Za-z]+\$?\d+)?$/.test(rr)&&validName(d.name))nm.push({name:d.name,ref:'='+rr});}wb.names=nm;}catch(e){wb.names=[];}
    for(const ws of w.worksheets){
      const s=newSheet(ws.name);
      try{const v=ws.views&&ws.views[0];if(v){if(v.showGridLines===false)s.grid=false;if(v.showRowColHeaders===false)s.head=false;if(v.state==='frozen'&&(v.xSplit||v.ySplit))s.freeze={r:v.ySplit||0,c:v.xSplit||0};}
        const ps=ws.pageSetup;if(ps){const P=s.page;P.paper={9:'A4',8:'A3',11:'A5',1:'Letter',5:'Legal'}[ps.paperSize]||'A4';P.orient=ps.orientation==='landscape'?'landscape':'portrait';
          if(ps.margins){for(const [k,n] of [['t','top'],['b','bottom'],['l','left'],['r','right']])if(typeof ps.margins[n]==='number')P.m[k]=ps.margins[n];}
          P.pg=!!ps.showGridLines;P.ph=!!ps.showRowColHeaders;
          if(ps.fitToPage){P.fitW=ps.fitToWidth===undefined?1:ps.fitToWidth;P.fitH=ps.fitToHeight===undefined?1:ps.fitToHeight;}else if(ps.scale)P.scale=ps.scale;
          if(typeof ps.printArea==='string'){const q=parseRng(ps.printArea.split(',')[0].replace(/^.*!/,''));if(q)P.area=q;}
          if(typeof ps.printTitlesRow==='string'){const m=/^\$?(\d+):\$?(\d+)$/.exec(ps.printTitlesRow);if(m)P.tRows={r1:+m[1]-1,r2:+m[2]-1};}
          if(typeof ps.printTitlesColumn==='string'){const m=/^\$?([A-Za-z]+):\$?([A-Za-z]+)$/.exec(ps.printTitlesColumn);if(m)P.tCols={c1:colIdx(m[1]),c2:colIdx(m[2])};}}}catch(e){}
      const dc=ws.properties&&ws.properties.defaultColWidth;if(dc)s.cw.def=Math.round(dc*7+5);
      const cc=ws.columnCount||0;
      for(let c=1;c<=cc;c++){const col=ws.getColumn(c);if(col.width)s.cw.ov.set(c-1,Math.round(col.width*7+5));if(col.hidden)s.cw.hid.add(c-1);}
      ws.eachRow({includeEmpty:true},(row,rn)=>{
        if(row.height)s.rh.ov.set(rn-1,Math.round(row.height*4/3)),(s.rh.manual=s.rh.manual||new Set()).add(rn-1);
        if(row.hidden)s.rh.hid.add(rn-1);
        row.eachCell({includeEmpty:true},(cell,cn)=>{
          if(cell.type===ExcelJS.ValueType.Merge)return;
          const o={};let val=cell.value;let link=null;
          if(val&&typeof val==='object'&&!(val instanceof Date)){
            if(val.richText)val=val.richText.map(x=>x.text).join('');
            else if('formula' in val||'sharedFormula' in val){let f=cell.formula;
              if(f){o.f=f.replace(/^=/,'');}else{val=val.result!==undefined?val.result:null;if(val&&val.error)val=val.error;}
              if(o.f!==undefined)val=null;}
            else if(val.hyperlink){link=val.hyperlink;val=typeof val.text==='string'?val.text:val.hyperlink;}
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
          {const pr=cell.protection;if(pr){if(pr.locked===false)st.lock=false;if(pr.hidden)st.hide=1;}
           const nt=cell.note;if(nt){const t=typeof nt==='string'?nt:(nt.texts||[]).map(z=>z.text).join('');const mm=/^([^\n:]{1,40}):\n([\s\S]*)$/.exec(t);s.cm.push({r:rn-1,c:cn-1,author:mm?mm[1]:'Pengguna',text:mm?mm[2]:t,w:200,h:70,show:false});}}
          if(link)st.link=String(link);if(Object.keys(st).length)o.st=st;
          if(o.v!==null||o.f!==undefined||o.st)s.cells.set(key(rn-1,cn-1),o);
        });
      });
      try{for(const im of (ws.getImages()||[])){
        const img=w.getImage(im.imageId);if(!img||!img.buffer)continue;
        const tl=im.range.tl,cc=Math.floor(tl.col),rr=Math.floor(tl.row);
        const x=s.cw.pos(cc)+(tl.col-cc)*s.cw.size(cc),y=s.rh.pos(rr)+(tl.row-rr)*s.rh.size(rr);
        let wd=150,ht=100;
        if(im.range.ext){wd=im.range.ext.width;ht=im.range.ext.height;}
        else if(im.range.br){const br=im.range.br,bc=Math.floor(br.col),brr=Math.floor(br.row);wd=s.cw.pos(bc)+(br.col-bc)*s.cw.size(bc)-x;ht=s.rh.pos(brr)+(br.row-brr)*s.rh.size(brr)-y;}
        const blob=new Blob([img.buffer],{type:'image/'+img.extension});
        const src=await new Promise(res=>{const f=new FileReader();f.onload=()=>res(f.result);f.readAsDataURL(blob);});
        s.objs.push({id:newId(),type:'img',url:URL.createObjectURL(blob),src,x,y,w:Math.max(20,wd),h:Math.max(20,ht)});
      }}catch(e){}
      for(const m of (ws.model.merges||[])){const p=m.split(':');const a=parseAddr(p[0]),b=parseAddr(p[1]||p[0]);if(a&&b)s.merges.push({r1:a.r,c1:a.c,r2:b.r,c2:b.c});}
      const af=ws.autoFilter;if(af){let a,b;if(typeof af==='string'){const p=af.split(':');a=parseAddr(p[0]);b=parseAddr(p[1]||p[0]);}else if(af.from&&af.to){a={r:af.from.row-1,c:af.from.column-1};b={r:af.to.row-1,c:af.to.column-1};}if(a&&b)s.filter={r1:a.r,c1:a.c,r2:b.r,c2:b.c,crit:{}};}
      try{
        const rl=new Map();ws.eachRow({includeEmpty:false},(row,rn)=>{if(row.outlineLevel)rl.set(rn-1,row.outlineLevel);});
        const build=(lv,max,hidden)=>{const out=[];for(let L=1;L<=8;L++){let a=null;for(let i=0;i<=max+1;i++){const on=(lv.get(i)||0)>=L;if(on&&a===null)a=i;if(!on&&a!==null){const b=i-1;const hid=[...Array(b-a+1).keys()].every(k=>hidden(a+k));out.push({a,b,cl:hid});a=null;}}}return out;};
        const mr=Math.max(0,...rl.keys());if(rl.size)s.og.rows=build(rl,mr,i=>s.rh.hid.has(i));
        const cl2=new Map();for(let c=1;c<=cc;c++){const col=ws.getColumn(c);if(col.outlineLevel)cl2.set(c-1,col.outlineLevel);}
        if(cl2.size)s.og.cols=build(cl2,Math.max(...cl2.keys()),i=>s.cw.hid.has(i));
        for(const g of s.og.rows)if(g.cl)for(let i=g.a;i<=g.b;i++)s.og.hr.add(i);
        for(const g of s.og.cols)if(g.cl)for(let i=g.a;i<=g.b;i++)s.og.hc.add(i);
      }catch(e){}
      try{const dvm=ws.dataValidations&&ws.dataValidations.model||{};const arr=[];
        for(const adr in dvm){const v=dvm[adr],pq=parseRng(adr.split(':')[0]),p=pq&&{r:pq.r1,c:pq.c1};if(!p||!v||!v.type)continue;
          const t={whole:'whole',decimal:'decimal',list:'list',date:'date',textLength:'len',custom:'custom'}[v.type];if(!t)continue;
          const op={between:'between',notBetween:'notBetween',equal:'equal',notEqual:'notEqual',greaterThan:'gt',lessThan:'lt',greaterThanOrEqual:'ge',lessThanOrEqual:'le'}[v.operator]||'between';
          const f=v.formulae||[];const rule={type:t,op,a:'',b:'',list:'',blank:v.allowBlank!==false,showIn:!!v.showInputMessage,inTitle:v.promptTitle||'',inMsg:v.prompt||'',showErr:v.showErrorMessage!==false,style:v.errorStyle||'stop',errTitle:v.errorTitle||'',errMsg:v.error||''};
          if(t==='list'){const x=String(f[0]??'');rule.list=/^".*"$/.test(x)?x.slice(1,-1).replace(/,/g,';'):'='+x;}
          else if(t==='custom')rule.a=String(f[0]??'');else{rule.a=f[0]===undefined?'':String(f[0]).replace('.',',');rule.b=f[1]===undefined?'':String(f[1]).replace('.',',');}
          const last=arr[arr.length-1];const k=JSON.stringify(rule);
          if(last&&last.k===k&&last.r.c1===p.c&&last.r.c2===p.c&&last.r.r2===p.r-1)last.r.r2=p.r;else arr.push({k,r:{...rule,r1:p.r,r2:p.r,c1:p.c,c2:p.c}});}
        s.dv=arr.map(x=>x.r);}catch(e){}
      try{const sp=ws.sheetProtection;if(sp&&(sp.sheet===true||sp.sheet==='1'||sp.sheet===1)){const al=k=>sp[k]===true;s.prot={pw:null,h:null,o:{formatCells:al('formatCells'),formatCols:al('formatColumns'),formatRows:al('formatRows'),insCols:al('insertColumns'),insRows:al('insertRows'),delCols:al('deleteColumns'),delRows:al('deleteRows'),sort:al('sort'),filter:al('autoFilter'),objects:sp.objects===true}};}}catch(e){}
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
      if(ext!=='xlsx'&&ext!=='xlsm')wb.names=[];
      wb.conns=ext==='xlsx'||ext==='xlsm'?[]:[];
      wb.sheets=sheets;wb.active=0;cur=sheets[0];undoS=[];redoS=[];clip=null;dirty=false;
      sheets.forEach(s=>{const rs=[];for(const [k,cl] of s.cells)if(cl.st&&cl.st.wrap)rs.push(Math.floor(k/MAXC));if(rs.length){const save=cur;cur=s;fitRows([...new Set(rs)].filter(r=>!(s.rh.ov.has(r))));cur=save;}if(s.filter){const save=cur;cur=s;applyFilter(true);cur=save;}});
      recalcAll();setName(file.name.replace(/\.(csv|tsv|txt|xlsm)$/i,'.xlsx'));updUndo();drawTabs();
      vp.scrollLeft=0;vp.scrollTop=0;select(0,0);
    }catch(e){console.error(e);toast('Tidak bisa membuka berkas: '+e.message);}
  }
  function confirmNew(){
    const run=()=>{wb.conns=[];wb.names=[];wb.watch=[];wb.manual=false;wb.sheets=[newSheet('Sheet1')];wb.active=0;cur=wb.sheets[0];undoS=[];redoS=[];clip=null;dirty=false;setName('Buku1.xlsx');recalcAll();updUndo();drawTabs();vp.scrollLeft=0;vp.scrollTop=0;select(0,0);};
    if(dirty)confirmDlg('Perubahan belum disimpan akan hilang. Lanjutkan membuat buku kerja baru?',run);else run();
  }
  const dl=(blob,fn)=>{const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=fn;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);toast('Tersimpan: '+fn);};
  async function saveXlsx(){
    try{
      if(ed&&!commitEdit())return;
      await lib('exceljs',()=>window.ExcelJS);
      const w=new ExcelJS.Workbook();w.creator='Office Web';let unsup=0;
      for(const s of wb.sheets){
        const ws=w.addWorksheet(s.name,{properties:s.tab?{tabColor:{argb:argb(s.tab)}}:{},views:[{...(s.freeze&&(s.freeze.r||s.freeze.c)?{state:'frozen',xSplit:s.freeze.c,ySplit:s.freeze.r}:{}),showGridLines:s.grid!==false,showRowColHeaders:s.head!==false}]});
        if(s.cw.def!==DW)ws.properties.defaultColWidth=(s.cw.def-5)/7;
        {const P=s.page||newPage();
          ws.pageSetup={paperSize:{A4:9,A3:8,A5:11,Letter:1,Legal:5}[P.paper]||9,orientation:P.orient==='landscape'?'landscape':'portrait',margins:{left:P.m.l,right:P.m.r,top:P.m.t,bottom:P.m.b,header:.3,footer:.3},showGridLines:!!P.pg,showRowColHeaders:!!P.ph,
            ...(P.area?{printArea:rngStr(P.area)}:{}),...(P.tRows?{printTitlesRow:`${P.tRows.r1+1}:${P.tRows.r2+1}`}:{}),...(P.tCols?{printTitlesColumn:`${colName(P.tCols.c1)}:${colName(P.tCols.c2)}`}:{}),
            ...(P.fitW||P.fitH?{fitToPage:true,fitToWidth:P.fitW||0,fitToHeight:P.fitH||0}:{scale:P.scale})};}
        for(const [c,v] of s.cw.ov)ws.getColumn(c+1).width=Math.max(0,(v-5)/7);
        for(const c of s.cw.hid)ws.getColumn(c+1).hidden=true;
        for(const [r,v] of s.rh.ov)ws.getRow(r+1).height=v*.75;
        for(const r of s.rh.hid)ws.getRow(r+1).hidden=true;
        for(const [k,cl] of s.cells){
          const r=Math.floor(k/MAXC),c=k%MAXC,x=ws.getCell(r+1,c+1);
          if(cl.f!==undefined&&cl.f!==null){const v=rawVal(wb,s,r,c);x.value={formula:cl.f,result:isErr(v)?{error:v.code}:v};}
          else if(cl.v!==null&&cl.v!==undefined)x.value=cl.v;
          if(cl.st&&cl.st.link&&(cl.f===undefined||cl.f===null))x.value={text:String(cl.v??''),hyperlink:cl.st.link};
          const st=cl.st;if(!st)continue;
          if(st.lock===false||st.hide)x.protection={locked:st.lock!==false,hidden:!!st.hide};
          if(st.font||st.sz||st.b||st.i||st.u||st.s||st.color){x.font={name:st.font||'Calibri',size:st.sz||11,bold:!!st.b,italic:!!st.i,underline:st.u?(st.u===2?'double':true):false,strike:!!st.s,...(st.color?{color:{argb:argb(st.color)}}:{})};}
          if(st.fill)x.fill={type:'pattern',pattern:'solid',fgColor:{argb:argb(st.fill)}};
          if(st.ha||st.va||st.wrap||st.ind||st.rot){x.alignment={horizontal:st.ha,vertical:st.va==='middle'?'middle':st.va,wrapText:!!st.wrap,indent:st.ind||0,textRotation:st.rot?(st.rot===255?255:st.rot>0?st.rot:90-st.rot):0};}
          if(st.bd){const b={};for(const [k2,n] of [['t','top'],['r','right'],['b','bottom'],['l','left']])if(st.bd[k2])b[n]={style:st.bd[k2][0],color:{argb:argb(st.bd[k2][1])}};x.border=b;}
          if(st.fmt)x.numFmt=st.fmt;
        }
        for(const o of s.objs){
          if(o.type!=='img'||!o.src){unsup++;continue;}
          const mm=/^data:image\/(\w+);base64,(.*)$/.exec(o.src);if(!mm){unsup++;continue;}
          try{const id=w.addImage({base64:mm[2],extension:mm[1]==='jpg'?'jpeg':mm[1]});
            const cc=s.cw.at(o.x),rr=s.rh.at(o.y);
            ws.addImage(id,{tl:{col:cc+(o.x-s.cw.pos(cc))/Math.max(1,s.cw.size(cc)),row:rr+(o.y-s.rh.pos(rr))/Math.max(1,s.rh.size(rr))},ext:{width:o.w,height:o.h}});}catch(e){unsup++;}
        }
        unsup+=s.spark.length;
        for(const d of s.dv){
          const dvR={type:{whole:'whole',decimal:'decimal',list:'list',date:'date',len:'textLength',custom:'custom'}[d.type],allowBlank:d.blank!==false,showErrorMessage:d.showErr!==false,errorStyle:d.style||'stop',errorTitle:d.errTitle||'',error:d.errMsg||'',showInputMessage:d.showIn!==false&&!!(d.inMsg||d.inTitle),promptTitle:d.inTitle||'',prompt:d.inMsg||''};
          if(!dvR.type)continue;
          if(d.type==='list'){dvR.formulae=[d.list[0]==='='?d.list.slice(1):'"'+d.list.replace(/;/g,',')+'"'];}
          else if(d.type==='custom')dvR.formulae=[String(d.a).replace(/^=/,'')];
          else{dvR.operator=({between:'between',notBetween:'notBetween',equal:'equal',notEqual:'notEqual',gt:'greaterThan',lt:'lessThan',ge:'greaterThanOrEqual',le:'lessThanOrEqual'})[d.op]||'between';
            const fv=v=>{const n=parseFloat(String(v).replace(',','.'));return isNaN(n)?String(v).replace(/^=/,''):n;};dvR.formulae=[fv(d.a)];if(dvR.operator==='between'||dvR.operator==='notBetween')dvR.formulae.push(fv(d.b));}
          let n=0;for(let r=d.r1;r<=d.r2&&n<20000;r++)for(let c=d.c1;c<=d.c2&&n<20000;c++,n++)try{ws.getCell(r+1,c+1).dataValidation={...dvR};}catch(e){}
        }
        for(const g of (s.og?s.og.rows:[])){for(let r=g.a;r<=Math.min(g.b,g.a+100000);r++){const row=ws.getRow(r+1);row.outlineLevel=(row.outlineLevel||0)+1;}}
        for(const g of (s.og?s.og.cols:[])){for(let c=g.a;c<=Math.min(g.b,g.a+2000);c++){const col=ws.getColumn(c+1);col.outlineLevel=(col.outlineLevel||0)+1;}}
        for(const cm of s.cm)try{ws.getCell(cm.r+1,cm.c+1).note=cm.author+':\n'+cm.text;}catch(e){}
        if(s.prot){const o=s.prot.o;try{await ws.protect(s.prot.pw||'',{selectLockedCells:true,selectUnlockedCells:true,formatCells:!!o.formatCells,formatColumns:!!o.formatCols,formatRows:!!o.formatRows,insertColumns:!!o.insCols,insertRows:!!o.insRows,deleteColumns:!!o.delCols,deleteRows:!!o.delRows,sort:!!o.sort,autoFilter:!!o.filter,objects:!!o.objects});}catch(e){}}
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
      try{for(const n of wb.names){const t=nameTarget(n);if(t&&/^[^!]+![$A-Za-z0-9:']+$/.test(n.ref.slice(1)))w.definedNames.add(n.ref.slice(1),n.name);}}catch(e){}
      const buf=await w.xlsx.writeBuffer();
      dl(new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),/\.xlsx$/i.test(name)?name:name.replace(/\.[^.]+$/,'')+'.xlsx');
      dirty=false;
      if(unsup)setTimeout(()=>toast(unsup+' objek (bagan, bentuk, kotak teks, atau sparkline) belum bisa disimpan ke .xlsx dan tidak ikut tersimpan.'),600);
    }catch(e){console.error(e);toast('Gagal menyimpan: '+e.message);}
  }
  function saveCsv(){
    const u=usedRange(cur);const lines=[];
    if(u.any)for(let r=0;r<=u.r2;r++){const row=[];for(let c=0;c<=u.c2;c++){const cl=cellAt(cur,r,c);const v=cl?cellVal(cur,r,c):null;let t=v===null?'':typeof v==='number'?String(v):fmtValue(v,cl&&cl.st&&cl.st.fmt).t;row.push(/[",\n;]/.test(t)?'"'+t.replace(/"/g,'""')+'"':t);}lines.push(row.join(','));}
    dl(new Blob(['\uFEFF'+lines.join('\r\n')],{type:'text/csv;charset=utf-8'}),name.replace(/\.[^.]+$/,'')+'.csv');
  }
  function printSheet(){
    if(ed&&!commitEdit())return;
    const P=cur.page,u=usedRange(cur);
    const g=P.area?{...P.area}:{r1:0,c1:0,r2:u.r2,c2:u.c2};
    let [pw,ph]=PAPER[P.paper]||PAPER.A4;if(P.orient==='landscape')[pw,ph]=[ph,pw];
    let totW=P.ph?40:0,totH=P.ph?20:0;
    for(let c=g.c1;c<=g.c2;c++)totW+=cur.cw.size(c);
    for(let r=g.r1;r<=g.r2;r++)totH+=cur.rh.size(r);
    const availW=(pw/25.4-P.m.l-P.m.r)*96,availH=(ph/25.4-P.m.t-P.m.b)*96;
    let zoom=P.scale/100;
    if(P.fitW||P.fitH)zoom=Math.min(1,P.fitW?availW*P.fitW/totW:Infinity,P.fitH?availH*P.fitH/totH:Infinity);
    const gl=P.pg?'border:1px solid #bbb;':'';
    const cell=(r,c)=>{
      if(cur.cw.size(c)===0)return '';
      const m=mergeAtSh(cur,r,c);if(m&&(m.r1!==r||m.c1!==c))return '';
      const cl=cellAt(cur,r,c),st=(cl&&cl.st)||{},f=fmtValue(cl?cellVal(cur,r,c):null,st.fmt);
      const bds=st.bd?SIDES.map(sd=>st.bd[sd]?`border-${{t:'top',r:'right',b:'bottom',l:'left'}[sd]}:${BTYPES[st.bd[sd][0]].join(' ')} ${st.bd[sd][1]};`:'').join(''):'';
      return `<td ${m?`rowspan="${m.r2-m.r1+1}" colspan="${m.c2-m.c1+1}"`:''} style="${gl}padding:1px 4px;height:${cur.rh.size(r)*.75}pt;min-width:${cur.cw.size(c)}px;max-width:${cur.cw.size(c)}px;overflow:hidden;white-space:nowrap;text-align:${st.ha&&st.ha!=='general'?st.ha:(f.al||'left')};${st.fill?'background:'+st.fill+';':''}${st.color?'color:'+st.color+';':''}${st.b?'font-weight:700;':''}${st.i?'font-style:italic;':''}${st.u?'text-decoration:underline;':''}${st.sz?'font-size:'+st.sz+'pt;':''}${st.font?'font-family:'+st.font+';':''}${bds}">${esc(f.t)}</td>`;};
    const th='border:1px solid #bbb;background:#eee;font-weight:400;text-align:center;';
    const row=r=>{let h='<tr>';if(P.ph)h+=`<th style="${th}">${r+1}</th>`;for(let c=g.c1;c<=g.c2;c++)h+=cell(r,c);return h+'</tr>';};
    let head='',body='';
    if(P.ph){head+='<tr><th style="'+th+'"></th>';for(let c=g.c1;c<=g.c2;c++)if(cur.cw.size(c))head+=`<th style="${th}">${colName(c)}</th>`;head+='</tr>';}
    const tr=[];if(P.tRows)for(let r=P.tRows.r1;r<=P.tRows.r2;r++)tr.push(r);
    for(const r of tr)if(cur.rh.size(r))head+=row(r);
    for(let r=g.r1;r<=g.r2;r++){if(tr.includes(r)||cur.rh.size(r)===0)continue;body+=row(r);}
    const p=document.createElement('div');p.id='xl-print';
    p.innerHTML=`<table style="border-collapse:collapse;table-layout:fixed;font:11pt ${DF()},Calibri,Carlito,sans-serif;zoom:${zoom}"><thead>${head}</thead><tbody>${body}</tbody></table>`;
    const pst=document.createElement('style');pst.textContent=`@page{size:${pw}mm ${ph}mm;margin:${P.m.t}in ${P.m.r}in ${P.m.b}in ${P.m.l}in}`;
    document.head.appendChild(pst);document.body.appendChild(p);
    const done=()=>{removeEventListener('afterprint',done);p.remove();pst.remove();};addEventListener('afterprint',done);window.print();
  }

  /* ===== event ===== */
  root.addEventListener('mousedown',e=>{if(e.target.closest('.b,.tab,.stbar button,.gal button,.gmore button,.fb,.shbar button,.shadd')){e.preventDefault();}});
  root.addEventListener('click',e=>{
    const ogb=e.target.closest('[data-og]');if(ogb){ogClick(ogb.dataset.og);return;}
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
  root.addEventListener('change',e=>{const s=e.target.dataset&&e.target.dataset.s;if(!s||!sels[s])return;sels[s](e.target.type==='checkbox'?e.target.checked:e.target.value);if(e.target.type!=='number'&&!ed)vp.focus({preventScroll:true});});
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
