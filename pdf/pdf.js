/* PDF viewer/editor — tampilan seperti Word, menu ala Nitro Pro (saat ini: Home & Convert).
   Mengekspos Apps.pdf.mount(root,{file,onHome}) */
(function(){
const CDN={
  pdfjs:'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js',
  worker:'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js',
  pdflib:'https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js',
  jszip:'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
  docx:'https://cdn.jsdelivr.net/npm/docx@9.5.1/dist/index.iife.js',
  xlsx:'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  pptx:'https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundle.js',
  pdfenc:'https://cdn.jsdelivr.net/npm/pdf-lib-plus-encrypt@1.1.0/dist/pdf-lib-plus-encrypt.min.js'
};
const libs={};
function lib(n,test){
  if(test()) return Promise.resolve();
  return libs[n]||(libs[n]=new Promise((ok,err)=>{const s=document.createElement('script');s.src=CDN[n];s.onload=ok;s.onerror=()=>{delete libs[n];err(new Error('Gagal memuat '+n));};document.head.appendChild(s);}));
}
/* pdf-lib biasa tidak bisa mengenkripsi; fork "plus-encrypt" dimuat terpisah tanpa menimpa window.PDFLib */
let PLE=null;
async function loadEnc(){
  if(PLE) return PLE;
  await lib('pdflib',()=>window.PDFLib);
  const old=window.PDFLib;
  await new Promise((ok,err)=>{const s=document.createElement('script');s.src=CDN.pdfenc;s.onload=ok;s.onerror=()=>err(new Error('Gagal memuat pustaka enkripsi'));document.head.appendChild(s);});
  PLE=window.PDFLib;window.PDFLib=old;
  if(!PLE||!PLE.PDFDocument||!PLE.PDFDocument.prototype.encrypt) throw new Error('Pustaka enkripsi tidak valid');
  return PLE;
}
let pdfjsReady;
function initPdfjs(){
  return pdfjsReady||(pdfjsReady=(async()=>{
    await lib('pdfjs',()=>window.pdfjsLib);
    try{const t=await (await fetch(CDN.worker)).text();pdfjsLib.GlobalWorkerOptions.workerSrc=URL.createObjectURL(new Blob([t],{type:'text/javascript'}));}
    catch{pdfjsLib.GlobalWorkerOptions.workerSrc=CDN.worker;}
  })().catch(e=>{pdfjsReady=null;throw e;}));
}
const DOCS={list:[],active:null,seq:0}; /* daftar berkas terbuka; tiap berkas punya riwayat undo/redo sendiri */
const store={get(k,d){try{const v=localStorage.getItem('pdf.'+k);return v==null?d:JSON.parse(v);}catch{return d;}},set(k,v){try{localStorage.setItem('pdf.'+k,JSON.stringify(v));}catch{}}};

/* ====================== fungsi murni (bisa diuji tanpa DOM) ====================== */
const U={};
U.parseRange=(s,n)=>{
  s=(s||'').trim();
  if(!s||/^(semua|all)$/i.test(s)) return [...Array(n).keys()];
  const out=new Set();
  for(const part of s.split(/[,;\s]+/).filter(Boolean)){
    const m=part.match(/^(\d+)(?:-(\d+))?$/); if(!m) return null;
    let a=+m[1],b=m[2]?+m[2]:a; if(a>b)[a,b]=[b,a];
    if(a<1||b>n) return null;
    for(let i=a;i<=b;i++) out.add(i-1);
  }
  return [...out].sort((x,y)=>x-y);
};
U.safe=s=>[...String(s)].map(c=>/[\x20-\x7E\xA0-\xFF]/.test(c)?c:'?').join('');
/* item teks pdf.js -> baris */
U.groupLines=items=>{
  const its=items.filter(t=>t.s&&t.s.trim()).sort((a,b)=>b.y-a.y||a.x-b.x);
  const lines=[];
  for(const t of its){
    const L=lines[lines.length-1];
    if(L&&Math.abs(L.y-t.y)<=Math.max(2,L.h*0.45)){L.items.push(t);L.h=Math.max(L.h,t.h);}
    else lines.push({y:t.y,h:t.h,items:[t]});
  }
  for(const L of lines){
    L.items.sort((a,b)=>a.x-b.x);
    let txt='',pe=null;
    for(const t of L.items){
      if(pe!=null&&t.x-pe>t.h*0.2&&!/\s$/.test(txt)&&!/^\s/.test(t.s)) txt+=' ';
      txt+=t.s; pe=t.x+t.w;
    }
    L.text=txt.replace(/\s+$/,'');
    L.x0=L.items[0].x;
    const last=L.items[L.items.length-1]; L.x1=last.x+last.w;
  }
  return lines;
};
/* baris -> paragraf (baris yang terpotong otomatis digabung) */
U.paras=lines=>{
  if(!lines.length) return [];
  const minL=Math.min(...lines.map(l=>l.x0)),maxR=Math.max(...lines.map(l=>l.x1)),span=Math.max(1,maxR-minL);
  const hs=lines.map(l=>l.h).sort((a,b)=>a-b),med=hs[hs.length>>1]||10;
  const out=[];let cur=null,prev=null;
  for(const l of lines){
    const gap=prev?prev.y-l.y:0;
    const wrap=prev&&cur&&prev.x1>=minL+span*0.88&&gap<=Math.max(prev.h,l.h)*1.6&&Math.abs(l.h-prev.h)<1&&l.x0<=minL+span*0.2;
    if(wrap) cur.text+=' '+l.text;
    else{cur={text:l.text,size:l.h,bold:l.h>med*1.25};out.push(cur);}
    prev=l;
  }
  return out;
};
/* baris -> sel lembar kerja (kolom disejajarkan menurut posisi x) */
U.sheetRows=lines=>{
  const rows=lines.map(l=>{
    const cells=[];let cur=null;
    for(const t of l.items){
      if(cur&&t.x-cur.end<=Math.max(8,t.h*0.9)){
        cur.text+=(t.x-cur.end>t.h*0.2&&!/\s$/.test(cur.text)&&!/^\s/.test(t.s)?' ':'')+t.s;cur.end=t.x+t.w;
      }else{cur={x:t.x,end:t.x+t.w,text:t.s};cells.push(cur);}
    }
    return cells;
  });
  const xs=[...new Set(rows.flat().map(c=>Math.round(c.x)))].sort((a,b)=>a-b),cols=[];
  for(const x of xs) if(!cols.length||x-cols[cols.length-1]>6) cols.push(x);
  return rows.map(r=>{
    const row=[];
    for(const c of r){
      let k=cols.length-1;while(k>0&&cols[k]>c.x+3)k--;
      while(row[k]!==undefined)k++;
      const v=c.text.trim();
      row[k]=/^-?\d+(\.\d+)?$/.test(v)?Number(v):v;
    }
    for(let i=0;i<row.length;i++) if(row[i]===undefined) row[i]='';
    return row;
  });
};
U.rtfEsc=s=>[...s].map(c=>{
  const n=c.codePointAt(0);
  if(c==='\\'||c==='{'||c==='}') return '\\'+c;
  if(n>126){
    if(n>0xFFFF){const m=n-0x10000,hi=0xD800+(m>>10),lo=0xDC00+(m&1023);return `\\u${hi-65536}?\\u${lo-65536}?`;}
    return `\\u${n>32767?n-65536:n}?`;
  }
  return c;
}).join('');
U.rtf=pagesParas=>{
  let o='{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\fnil Calibri;}}\\f0\n';
  pagesParas.forEach((ps,i)=>{
    if(i>0) o+='\\page\n';
    for(const p of ps){
      const fs=Math.max(16,Math.min(144,Math.round(p.size*2)));
      o+=`\\pard\\sa160\\fs${fs} ${p.bold?'\\b ':''}${U.rtfEsc(p.text)}${p.bold?'\\b0':''}\\par\n`;
    }
  });
  return o+'}';
};
U.wrap=(text,font,size,maxW)=>{
  const out=[],W=s=>font.widthOfTextAtSize(s,size);
  for(const raw of String(text).replace(/\r\n?/g,'\n').split('\n')){
    const line=U.safe(raw.replace(/\t/g,'    '));
    if(!line){out.push('');continue;}
    let cur='';
    for(const w of line.split(/(\s+)/)){
      if(!w) continue;
      if(W(cur+w)>maxW&&cur.trim()){out.push(cur.replace(/\s+$/,''));cur=w.replace(/^\s+/,'');}
      else cur+=w;
      while(W(cur)>maxW&&cur.length>1){let k=cur.length-1;while(k>1&&W(cur.slice(0,k))>maxW)k--;out.push(cur.slice(0,k));cur=cur.slice(k);}
    }
    out.push(cur);
  }
  return out;
};
const TXT_ASC=0.9465; /* posisi baseline dalam kotak teks (line-height 1.2) */
/* ---- deteksi blok paragraf (untuk alat Edit ala Nitro) ---- */
U.MET={serif:[0.891,0.216],sans:[0.905,0.212],mono:[0.833,0.3]};
U.FN={serif:['TimesRoman','TimesRomanBold','TimesRomanItalic','TimesRomanBoldItalic'],sans:['Helvetica','HelveticaBold','HelveticaOblique','HelveticaBoldOblique'],mono:['Courier','CourierBold','CourierOblique','CourierBoldOblique']};
U.OKCH=/^[\x20-\x7E\xA0-\xFF\u2018\u2019\u201A\u201C\u201D\u201E\u2013\u2014\u2022\u2026\u20AC\u2122\u2020\u2021\u2030\u2039\u203A\u0152\u0153\u0160\u0161\u0178\u017D\u017E\u0192\u02C6\u02DC]$/;
U.cleanStr=s=>String(s).replace(/[\u200B-\u200D\uFEFF\u00AD]/g,'').replace(/\u00A0/g,' ');
U.normRuns=runs=>{
  const out=[];let sp=true;
  for(const r of runs){
    let t=r.t.replace(/[ \t\r\n]+/g,' ');
    if(sp&&t[0]===' ')t=t.slice(1);
    if(!t)continue;
    sp=t[t.length-1]===' ';
    const l=out[out.length-1];
    if(l&&l.i===r.i&&l.b===r.b)l.t+=t;else out.push({t,i:!!r.i,b:!!r.b});
  }
  if(out.length){const l=out[out.length-1];l.t=l.t.replace(/ +$/,'');if(!l.t)out.pop();}
  return out;
};
/* items: {s,x,y(baseline, asal kiri-atas),w,h,it,bd,fam} dalam satuan halaman */
U.blocks=items=>{
  const its=items.map(t=>({...t,s:U.cleanStr(t.s)})).filter(t=>t.s!==''&&isFinite(t.x)&&isFinite(t.y)&&isFinite(t.w)&&t.h>0);
  if(!its.length)return [];
  its.sort((a,b)=>a.y-b.y||a.x-b.x);
  const rows=[];
  for(const t of its){
    const R=rows[rows.length-1];
    if(R&&Math.abs(t.y-R.y)<=Math.max(2,R.h*0.45)){R.items.push(t);R.h=Math.max(R.h,t.h);}
    else rows.push({y:t.y,h:t.h,items:[t]});
  }
  const segOf=list=>{
    const nz=list.filter(t=>t.s.trim()!=='');
    if(!nz.length)return null;
    const runs=[];let pe=null,txt='';
    const addR=(s,t)=>{const l=runs[runs.length-1];if(l&&l.i===!!t.it&&l.b===!!t.bd)l.t+=s;else runs.push({t:s,i:!!t.it,b:!!t.bd});txt+=s;};
    for(const t of list){
      if(pe!=null&&t.x-pe>t.h*0.2&&!/\s$/.test(txt)&&!/^\s/.test(t.s))addR(' ',t);
      addR(t.s,t);pe=t.x+t.w;
    }
    const big=nz.reduce((m,t)=>t.s.length>m.s.length?t:m,nz[0]),last=nz[nz.length-1];
    const small=nz.some(t=>t.h<big.h*0.8);
    const ok=!small&&[...txt].every(c=>U.OKCH.test(c));
    return{x0:nz[0].x,x1:last.x+last.w,y:big.y,h:big.h,runs,bold:nz.every(t=>t.bd),fam:big.fam||'serif',ok};
  };
  const segs=[];
  for(const R of rows){
    R.items.sort((a,b)=>a.x-b.x);
    let cur=[],end=null;
    const flush=()=>{const s=segOf(cur);if(s)segs.push(s);cur=[];end=null;};
    for(const t of R.items){
      const sp=t.s.trim()==='';
      if(!sp&&end!=null&&t.x-end>R.h*5)flush();
      cur.push(t);if(!sp)end=t.x+t.w;
    }
    flush();
  }
  if(!segs.length)return [];
  const longs=segs.filter(s=>s.x1-s.x0>8*s.h),pool=longs.length?longs:segs;
  const rs=pool.map(s=>s.x1).sort((a,b)=>a-b),ls=pool.map(s=>s.x0).sort((a,b)=>a-b);
  const pageRight=rs[Math.min(rs.length-1,Math.floor(rs.length*0.85))],pageLeft=ls[Math.floor(ls.length*0.1)],colW=pageRight-pageLeft;
  /* dokumen rata kiri-kanan -> ujung baris ketat; rata kiri -> ujung baris boleh bergerigi */
  const justDoc=pool.filter(s=>Math.abs(s.x1-pageRight)<=1.5).length>=0.5*pool.length;
  segs.sort((a,b)=>a.y-b.y||a.x0-b.x0);
  const blocks=[];
  for(const sg of segs){
    let tgt=null;
    for(let k=blocks.length-1;k>=Math.max(0,blocks.length-40);k--){
      const b=blocks[k],l=b.segs[b.segs.length-1],d=sg.y-l.y,hh=Math.max(sg.h,l.h);
      if(d<0.7*sg.h||d>2.7*hh)continue;
      const r=sg.h/l.h;if(r<0.8||r>1.25)continue;
      if(sg.bold!==l.bold)continue;
      if(Math.min(sg.x1,l.x1)-Math.max(sg.x0,l.x0)<=0)continue;
      if(colW>20*sg.h&&l.x1<pageRight-(justDoc?Math.max(2.5*sg.h,0.08*colW):Math.max(5*sg.h,0.22*colW)))continue; /* baris sebelumnya pendek = paragraf berakhir */
      const ref=b.segs.length>1?Math.min(...b.segs.slice(1).map(s=>s.x0)):l.x0;
      if(sg.x0>ref+1.2*sg.h)continue; /* baris menjorok = paragraf baru */
      if(sg.x0<ref-(b.segs.length>1?1.5:4)*sg.h)continue;
      tgt=b;break;
    }
    if(tgt)tgt.segs.push(sg);else blocks.push({segs:[sg]});
  }
  return blocks.map((b,id)=>{
    const S=b.segs,n=S.length,f=S[0],size=f.h,fam=U.MET[f.fam]?f.fam:'serif';
    const x0=Math.min(...S.map(s=>s.x0)),x1=Math.max(...S.map(s=>s.x1));
    let align='left',bx=x0,bw=x1-x0,indent=0,cx0=x0,cx1=x1;
    const centered=colW>0&&n<=2&&S.every(s=>{const L=s.x0-pageLeft,R=pageRight-s.x1;return L>2*size&&R>2*size&&Math.abs(L-R)<=Math.max(3,size*0.35);});
    if(centered){align='center';bx=pageLeft;bw=colW;}
    else{
      indent=n>1?Math.max(0,f.x0-x0):0;
      if(n>=3){const nl=S.slice(0,-1),cnt=nl.filter(s=>Math.abs(s.x1-x1)<=1.5).length;if(cnt>=0.7*nl.length){align='justify';}}
      if(n===1){const room=Math.min(pageRight-bx,bw+6*size);bw=Math.max(bw+1,room);}
      else bw=x1-bx;
    }
    let pitch=0;
    if(n>1){const df=S.slice(1).map((s,k)=>s.y-S[k].y).sort((a,c)=>a-c);pitch=df[df.length>>1];}
    const lh=n>1?Math.round(Math.max(1,Math.min(3,pitch/size))*1000)/1000:1.2;
    const [A,D]=U.MET[fam],L=lh*size,by=f.y-((L-(A+D)*size)/2+A*size);
    const top=Math.min(...S.map(s=>s.y-s.h*0.95))-1,bot=Math.max(...S.map(s=>s.y+s.h*0.28))+1;
    let rr=[];
    S.forEach((s,k)=>{
      if(k>0&&rr.length){
        const pl=rr[rr.length-1],nt=s.runs[0]?s.runs[0].t:'';
        if(!(/[-\u2010]$/.test(pl.t.trimEnd())&&/^[a-z]/.test(nt.trimStart())))rr.push({t:' ',i:pl.i,b:pl.b});
      }
      s.runs.forEach(r=>rr.push({...r}));
    });
    rr=U.normRuns(rr);
    const ok=S.every(s=>s.ok)&&rr.every(r=>[...r.t].every(c=>U.OKCH.test(c)));
    return{id,x:bx,y:by,w:bw,h:n*L,size,fam,lh,indent,align,runs:rr,ok,n,cover:{x:cx0-2,y:top,w:cx1-cx0+4,h:bot-top}};
  });
};
/* teks pdf.js -> blok (koordinat viewport skala 1) */
U.pageBlocks=(page,tc,vp)=>{
  const st=tc.styles||{},fc={};
  const finfo=fn=>fc[fn]||(fc[fn]=(()=>{
    let it=false,bd=false,nm='';
    try{const f=page.commonObjs.get(fn);nm=String(f.name||f.fallbackName||'');it=!!f.italic||/ital|obli/i.test(nm);bd=!!f.bold||/bold|black|heavy|semibold|demi/i.test(nm);}catch{}
    const ff=String((st[fn]&&st[fn].fontFamily)||'serif');
    let fam=/mono/i.test(ff)?'mono':/sans/i.test(ff)?'sans':'serif';
    if(/courier|mono|consolas/i.test(nm))fam='mono';else if(/arial|helvet|calibri|verdana|tahoma|sans/i.test(nm))fam='sans';else if(/times|roman|georgia|cambria|garamond|serif/i.test(nm))fam='serif';
    return{it,bd,fam};
  })());
  const items=[];
  for(const t of tc.items){
    if(typeof t.str!=='string'||!t.str)continue;
    const tr=t.transform;if(Math.abs(tr[1])>0.1*Math.abs(tr[0]))continue;
    const [x,y]=vp.convertToViewportPoint(tr[4],tr[5]),fi=finfo(t.fontName);
    items.push({s:t.str,x,y,w:t.width||0,h:Math.abs(tr[3])||t.height||10,it:fi.it,bd:fi.bd,fam:fi.fam});
  }
  return U.blocks(items);
};
U.enc=(font,s)=>{
  const c=font.__ok||(font.__ok=new Map());
  return [...String(s)].map(ch=>{
    if(ch==='\n')return ch;if(ch==='\t')return ' ';
    let v=c.get(ch);
    if(v===undefined){try{font.widthOfTextAtSize(ch,1);v=true;}catch{v=false;}c.set(ch,v);}
    return v?ch:'?';
  }).join('');
};
/* pecah run menjadi baris menurut lebar. meas(token,teks)->lebar */
U.layoutRuns=(runs,meas,width,indent)=>{
  const toks=[];
  for(const r of runs)for(const part of r.t.split(/(\n| +)/)){
    if(part==='')continue;
    if(part==='\n')toks.push({nl:1});else toks.push({t:part,k:r.k,sp:part[0]===' '});
  }
  const lines=[];let cur={toks:[],w:0},avail=width-indent;
  const push=hard=>{while(cur.toks.length&&cur.toks[cur.toks.length-1].sp)cur.w-=cur.toks.pop().w;cur.hard=hard;lines.push(cur);cur={toks:[],w:0};avail=width;};
  for(const t of toks){
    if(t.nl){push(true);continue;}
    t.w=meas(t,t.t);
    if(t.sp){if(cur.toks.length){cur.toks.push(t);cur.w+=t.w;}continue;}
    if(cur.w+t.w>avail+0.01&&cur.toks.length)push(false);
    if(t.w>avail){
      let s=t.t;
      while(s){let k=s.length;while(k>1&&meas(t,s.slice(0,k))>avail)k--;
        const p={...t,t:s.slice(0,k)};p.w=meas(t,p.t);cur.toks.push(p);cur.w+=p.w;s=s.slice(k);if(s)push(false);}
      continue;
    }
    cur.toks.push(t);cur.w+=t.w;
  }
  if(cur.toks.length||!lines.length)push(true);
  return lines;
};
/* ---- gaya objek: font, spasi, garis putus, ujung garis ---- */
U.FAM={helv:'Helvetica,Arial,sans-serif',times:'"Times New Roman",Times,serif',cour:'"Courier New",Courier,monospace'};
U.FAMNAME={helv:'Arial',times:'Times New Roman',cour:'Courier New'};
U.stdFont=async(PL,doc,fam,b,i,cache)=>{
  const SF=PL.StandardFonts,N={helv:['Helvetica','HelveticaBold','HelveticaOblique','HelveticaBoldOblique'],times:['TimesRoman','TimesRomanBold','TimesRomanItalic','TimesRomanBoldItalic'],cour:['Courier','CourierBold','CourierOblique','CourierBoldOblique']}[fam]||['Helvetica','HelveticaBold','HelveticaOblique','HelveticaBoldOblique'];
  const idx=b?(i?3:1):(i?2:0),key=(fam||'helv')+idx;
  if(!cache[key]) cache[key]=await doc.embedFont(SF[N[idx]]);
  return cache[key];
};
/* posisi baseline baris pertama dalam kotak untuk line-height L (kelipatan ukuran font) */
U.asc=(L,fam)=>L/2+({helv:.3465,times:.3375,cour:.2665}[fam]||.3465);
U.textW=(font,size,s,cs=0,ws=0)=>font.widthOfTextAtSize(s,size)+cs*s.length+(ws?ws*(s.split(' ').length-1):0);
U.dash=(st,lw)=>st==='dashed'?[lw*4+2,lw*2+2]:st==='dotted'?[0.01,lw*2.2]:null;
U.ends=a=>a.shape==='arrow'?{s:a.s||'none',e:a.e||'open'}:a.shape==='line'?{s:a.s||'none',e:a.e||'none'}:{s:'none',e:'none'};
U.kindLabel=a=>a.type==='tx'?(a.co?'Call Out':a.bg?'Note':a.box?'Text Box':'Text'):a.type==='sh'?({rect:'Rectangle',ell:'Ellipse',line:'Line',arrow:'Arrow',measure:'Measure'}[a.shape]||'Shape'):a.type==='dr'?'Drawing':a.type==='im'?'Image':a.type==='lk'?'Link':a.type==='hl'?({ul:'Underline',st:'Strikeout',rd:'Redaction'}[a.kind]||'Highlight'):a.type==='tb'?'Text Edit':'Object';
/* kotak pembatas perkiraan (dipakai bila ukuran DOM tak tersedia) */
U.bbox=a=>{
  if(a.type==='tx'){const ls=String(a.text||'').split('\n'),w=Math.max(8,...ls.map(l=>l.length))*a.size*0.55+((a.box||a.bg)?8:0),h=Math.max(1,ls.length)*a.size*(a.ls||1.2)+((a.box||a.bg)?8:0);return{x:a.x,y:a.y,w,h};}
  return{x:a.x,y:a.y,w:a.w||0,h:a.h||0};
};
/* ---- helper murni untuk Review / Page Layout / Protect ---- */
U.col=(h,d=[0,0,0])=>{const m=/^#?([0-9a-f]{6})$/i.exec(h||'');return m?[0,2,4].map(i=>parseInt(m[1].slice(i,i+2),16)/255):d;};
U.rot=a=>{a=((Math.round(a||0)%360)+360)%360;return a%90===0?a:0;};
/* koordinat tampilan (asal kiri-bawah, y ke atas) -> koordinat halaman tak-berputar relatif cropbox */
U.v2p=(R,W,H,vx,vy)=>R===90?[W-vy,vx]:R===180?[W-vx,H-vy]:R===270?[vy,H-vx]:[vx,vy];
/* margin menurut tampilan {t,b,l,r} -> margin pada ruang halaman tak-berputar {L,R,T,B} */
U.cropMargins=(R,m)=>R===90?{L:m.t,T:m.r,R:m.b,B:m.l}:R===180?{L:m.r,R:m.l,T:m.b,B:m.t}:R===270?{L:m.b,T:m.l,R:m.t,B:m.r}:{L:m.l,R:m.r,T:m.t,B:m.b};
U.arrowPts=(sx,sy,ex,ey,len)=>{const a=Math.atan2(ey-sy,ex-sx),s=Math.PI/7;return[[ex-len*Math.cos(a-s),ey-len*Math.sin(a-s)],[ex-len*Math.cos(a+s),ey-len*Math.sin(a+s)]];};
U.ticks=(sx,sy,ex,ey,len)=>{const L=Math.hypot(ex-sx,ey-sy)||1,nx=-(ey-sy)/L*len,ny=(ex-sx)/L*len;return[[[sx-nx,sy-ny],[sx+nx,sy+ny]],[[ex-nx,ey-ny],[ex+nx,ey+ny]]];};
U.measureLabel=a=>{const d=Math.hypot(a.w,a.h),u=a.unit||'cm';return u==='cm'?(d/72*2.54).toFixed(2)+' cm':u==='in'?(d/72).toFixed(2)+' in':d.toFixed(1)+' pt';};
U.PENS={black:{color:'#111111',lw:1.5},red:{color:'#d90000',lw:1.5},blue:{color:'#1a56db',lw:1.5},marker:{color:'#ffe600',lw:12,mul:true}};
/* coretan bebas -> anotasi dengan kotak pembatas; titik disimpan relatif terhadap x,y */
U.normDr=(pts,P,page)=>{
  const pad=P.lw/2+1;let x0=1e9,y0=1e9,x1=-1e9,y1=-1e9;
  for(const[x,y]of pts){x0=Math.min(x0,x);y0=Math.min(y0,y);x1=Math.max(x1,x);y1=Math.max(y1,y);}
  x0-=pad;y0-=pad;x1+=pad;y1+=pad;
  return{type:'dr',page,x:x0,y:y0,w:x1-x0,h:y1-y0,pts:pts.map(([x,y])=>[+(x-x0).toFixed(2),+(y-y0).toFixed(2)]),color:P.color,lw:P.lw,mul:!!P.mul,cm:true};
};
U.shapeSvg=a=>{
  const w=a.w,h=a.h,lw=a.lw||2,c=a.color||'#d90000',da=U.dash(a.dash,lw),base=`stroke="${c}" stroke-width="${lw}" stroke-linecap="round" stroke-linejoin="round"`;
  const att=`fill="none" ${base}${da?` stroke-dasharray="${da.join(' ')}"`:''}`,fatt=`fill="${a.fill||'none'}" ${base}${da?` stroke-dasharray="${da.join(' ')}"`:''}`,solid=`fill="none" ${base}`;
  let b='';
  if(a.shape==='rect') b=`<rect x="0" y="0" width="${w}" height="${h}" ${fatt}/>`;
  else if(a.shape==='ell') b=`<ellipse cx="${w/2}" cy="${h/2}" rx="${w/2}" ry="${h/2}" ${fatt}/>`;
  else{
    const sx=a.fx?w:0,sy=a.fy?h:0,ex=a.fx?0:w,ey=a.fy?0:h,E=U.ends(a);
    b=`<line x1="${sx}" y1="${sy}" x2="${ex}" y2="${ey}" ${att}/>`;
    const head=(t,tx,ty,fx,fy)=>{
      if(t==='none') return '';
      const[p,q]=U.arrowPts(fx,fy,tx,ty,8+lw*2);
      return t==='closed'?`<path d="M${tx} ${ty}L${p[0]} ${p[1]}L${q[0]} ${q[1]}Z" fill="${c}" ${base}/>`:`<path d="M${p[0]} ${p[1]}L${tx} ${ty}L${q[0]} ${q[1]}" ${solid}/>`;
    };
    b+=head(E.e,ex,ey,sx,sy)+head(E.s,sx,sy,ex,ey);
    if(a.shape==='measure'){
      for(const[[x1,y1],[x2,y2]]of U.ticks(sx,sy,ex,ey,5)) b+=`<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" ${att}/>`;
      b+=`<text x="${(sx+ex)/2}" y="${(sy+ey)/2-6}" text-anchor="middle" font-size="11" font-family="Arial,Helvetica,sans-serif" font-weight="700" fill="${c}" stroke="#fff" stroke-width="3" paint-order="stroke">${U.measureLabel(a)}</text>`;
    }
  }
  return `<svg viewBox="0 0 ${w} ${h}" preserveAspectRatio="none" style="position:absolute;left:0;top:0;width:100%;height:100%;overflow:visible;pointer-events:none">${b}</svg>`;
};
/* "3" = tiap 3 halaman; "1-3; 4-6; 7" = bagian menurut rentang */
U.splitPlan=(spec,n)=>{
  spec=(spec||'').trim();if(!spec) return null;
  if(/^\d+$/.test(spec)){const k=+spec;if(k<1) return null;const out=[];for(let i=0;i<n;i+=k)out.push(Array.from({length:Math.min(k,n-i)},(_,j)=>i+j));return out;}
  const out=[];
  for(const part of spec.split(';')){if(!part.trim()) continue;const r=U.parseRange(part,n);if(!r||!r.length) return null;out.push(r);}
  return out.length?out:null;
};
U.hfText=(t,page,total)=>String(t).replace(/\{page\}/gi,page).replace(/\{total\}/gi,total).replace(/\{date\}/gi,new Date().toLocaleDateString('id-ID'));
U.bates=(pre,n,dig,suf)=>pre+String(n).padStart(dig,'0')+suf;
U.wordDiff=(a,b)=>{
  const m=new Map();for(const w of a.split(' ').filter(Boolean))m.set(w,(m.get(w)||0)+1);
  let add=0,del=0;
  for(const w of b.split(' ').filter(Boolean)){const c=m.get(w)||0;if(c>0)m.set(w,c-1);else add++;}
  for(const c of m.values())del+=c;
  return{add,del};
};
/* buang objek yang tak terjangkau dari Root/Info (mis. isi halaman lama setelah redaksi, XMP yang dilepas) */
U.gc=(PL,d)=>{
  const{PDFRef,PDFDict,PDFArray,PDFStream}=PL,ctx=d.context,seen=new Set(),stack=[];
  const add=o=>{if(o)stack.push(o);};
  const tr=ctx.trailerInfo;[tr.Root,tr.Info,tr.Encrypt,tr.ID].forEach(add);
  while(stack.length){
    const o=stack.pop();
    if(o instanceof PDFRef){if(seen.has(o.tag))continue;seen.add(o.tag);add(ctx.lookup(o));}
    else if(o instanceof PDFStream)add(o.dict);
    else if(o instanceof PDFDict)o.values().forEach(add);
    else if(o instanceof PDFArray)for(let i=0;i<o.size();i++)add(o.get(i));
  }
  let n=0;
  for(const[ref]of ctx.enumerateIndirectObjects())if(!seen.has(ref.tag)){ctx.delete(ref);n++;}
  return n;
};
/* ganti isi halaman i dengan gambar JPEG (ukuran tampilan w x h) di tempat: ref halaman tetap, semua isi lama dibuang */
U.redactPage=async(PL,d,i,jpg,w,h)=>{
  const{PDFName,PDFPage}=PL,pg=d.getPage(i),node=pg.node,ref=pg.ref;
  for(const k of node.keys()){const n=k.asString();if(n!=='/Type'&&n!=='/Parent')node.delete(k);}
  node.set(PDFName.of('Resources'),d.context.obj({}));
  const np=PDFPage.of(node,ref,d);
  np.setMediaBox(0,0,w,h);
  const im=await d.embedJpg(jpg);
  np.drawImage(im,{x:0,y:0,width:w,height:h});
};
/* tambah satu bookmark (outline) ke halaman pi pada akhir daftar */
U.addBookmark=(PL,d,title,pi)=>{
  const{PDFName,PDFNumber,PDFHexString,PDFDict}=PL,ctx=d.context,cat=d.catalog,N=n=>PDFName.of(n);
  let olRef=cat.get(N('Outlines')),ol=olRef&&ctx.lookup(olRef,PDFDict);
  if(!ol){ol=ctx.obj({Type:'Outlines',Count:0});olRef=ctx.register(ol);cat.set(N('Outlines'),olRef);}
  const item=ctx.obj({Title:PDFHexString.fromText(title),Parent:olRef,Dest:[d.getPage(pi).ref,'Fit']}),ref=ctx.register(item),last=ol.get(N('Last'));
  if(last){ctx.lookup(last,PDFDict).set(N('Next'),ref);item.set(N('Prev'),last);}
  else ol.set(N('First'),ref);
  ol.set(N('Last'),ref);
  const cnt=ol.get(N('Count'));ol.set(N('Count'),PDFNumber.of(Math.abs(cnt instanceof PDFNumber?cnt.asNumber():0)+1));
};
/* tempelkan anotasi ke dokumen pdf-lib. pages[i]={vp,rot} */
U.flatten=async(PL,doc,annots,pages)=>{
  const {StandardFonts,rgb,degrees,BlendMode,LineCapStyle}=PL;
  const font=await doc.embedFont(StandardFonts.Helvetica),pgs=doc.getPages(),fcache={};
  /* lapis 1: tutup teks asli blok yang diedit (putih) */
  for(const a of annots){
    if(a.type!=='tb'||!a.touched||!a.cover) continue;
    const pg=pgs[a.page],info=pages[a.page]; if(!pg||!info) continue;
    const vp=info.vp,r=a.cover,[x1,y1]=vp.convertToPdfPoint(r.x,r.y),[x2,y2]=vp.convertToPdfPoint(r.x+r.w,r.y+r.h);
    pg.drawRectangle({x:Math.min(x1,x2),y:Math.min(y1,y2),width:Math.abs(x2-x1),height:Math.abs(y2-y1),color:rgb(1,1,1),borderWidth:0});
  }
  for(const a of annots){
    const pg=pgs[a.page],info=pages[a.page]; if(!pg||!info) continue;
    const vp=info.vp,rot=info.rot||0;
    const P=(x,y)=>{const[px,py]=vp.convertToPdfPoint(x,y);return{x:px,y:py};};
    const OP=a.op==null?1:Math.max(0.05,Math.min(1,a.op/100)),opO=OP<1?{opacity:OP}:{};
    if(a.type==='tx'){
      const fam=a.font||'helv',tf=await U.stdFont(PL,doc,fam,a.bold,a.italic,fcache),c=U.col(a.color),lines=String(a.text).replace(/\r/g,'').split('\n').map(U.safe),L=a.ls||1.2,lh=a.size*L;
      const hasBox=!!(a.box||a.bg),bw=a.bw||1,BP=hasBox?3+bw:0,cs=a.cs||0,ws=a.ws||0,lws=lines.map(l=>U.textW(tf,a.size,l,cs,ws)),tw=Math.max(0,...lws);
      const W=hasBox?Math.max(tw+2*BP,a.bg?100:80):tw,cw=W-2*BP,H=lines.length*lh+2*BP;
      if(hasBox){
        const p1=P(a.x+bw/2,a.y+bw/2),p2=P(a.x+W-bw/2,a.y+H-bw/2),bc=U.col(a.box||'#444444',[.27,.27,.27]),da=U.dash(a.dash,bw);
        const o={x:Math.min(p1.x,p2.x),y:Math.min(p1.y,p2.y),width:Math.abs(p2.x-p1.x),height:Math.abs(p2.y-p1.y),borderColor:rgb(bc[0],bc[1],bc[2]),borderWidth:bw,...(da?{borderDashArray:da,borderLineCap:LineCapStyle.Round}:{}),...(OP<1?{opacity:OP,borderOpacity:OP}:{})};
        if(a.bg){const f=U.col(a.bg,[1,1,1]);o.color=rgb(f[0],f[1],f[2]);}
        pg.drawRectangle(o);
        if(a.co){const ll=a.ll||40,lo=a.lo||0,q=ll*Math.SQRT1_2;pg.drawLine({start:P(a.x+lo,a.y+H),end:P(a.x+lo-q,a.y+H+q),thickness:bw,color:rgb(bc[0],bc[1],bc[2]),...(da?{dashArray:da,lineCap:LineCapStyle.Round}:{}),...opO});}
      }
      lines.forEach((ln,i)=>{
        if(!ln) return;
        const off=a.align==='c'?(cw-lws[i])/2:a.align==='r'?cw-lws[i]:0,vx=a.x+BP+off,vy=a.y+BP+i*lh+a.size*U.asc(L,fam),col=rgb(c[0],c[1],c[2]);
        if(!cs&&!ws){const[px,py]=vp.convertToPdfPoint(vx,vy);pg.drawText(ln,{x:px,y:py,size:a.size,font:tf,color:col,rotate:degrees(rot),...opO});return;}
        let x=vx;
        for(const ch of ln){const[px,py]=vp.convertToPdfPoint(x,vy);if(ch!==' ')pg.drawText(ch,{x:px,y:py,size:a.size,font:tf,color:col,rotate:degrees(rot),...opO});x+=tf.widthOfTextAtSize(ch,a.size)+cs+(ch===' '?ws:0);}
      });
    }else if(a.type==='hl'){
      if(a.kind==='rd') continue; /* tanda redaksi diterapkan lewat jalur khusus, bukan di sini */
      for(const r of a.rects){
        if(a.kind==='ul'||a.kind==='st'){
          const yy=a.kind==='ul'?r.y+r.h:r.y+r.h*0.55;
          pg.drawLine({start:P(r.x,yy),end:P(r.x+r.w,yy),thickness:Math.max(0.8,r.h*0.07),color:rgb(.85,0,0)});
        }else{
          const [x1,y1]=vp.convertToPdfPoint(r.x,r.y),[x2,y2]=vp.convertToPdfPoint(r.x+r.w,r.y+r.h);
          pg.drawRectangle({x:Math.min(x1,x2),y:Math.min(y1,y2),width:Math.abs(x2-x1),height:Math.abs(y2-y1),color:rgb(1,0.9,0),borderWidth:0,blendMode:BlendMode.Multiply});
        }
      }
    }else if(a.type==='sh'){
      const c=U.col(a.color,[.85,0,0]),col=rgb(c[0],c[1],c[2]),lw=a.lw||2,da=U.dash(a.dash,lw),dA=da?{dashArray:da}:{};
      if(a.shape==='rect'||a.shape==='ell'){
        const p1=P(a.x,a.y),p2=P(a.x+a.w,a.y+a.h),o={borderColor:col,borderWidth:lw,...(da?{borderDashArray:da,borderLineCap:LineCapStyle.Round}:{}),...(OP<1?{opacity:OP,borderOpacity:OP}:{})};
        if(a.fill){const f=U.col(a.fill);o.color=rgb(f[0],f[1],f[2]);}
        if(a.shape==='rect') pg.drawRectangle({x:Math.min(p1.x,p2.x),y:Math.min(p1.y,p2.y),width:Math.abs(p2.x-p1.x),height:Math.abs(p2.y-p1.y),...o});
        else pg.drawEllipse({x:(p1.x+p2.x)/2,y:(p1.y+p2.y)/2,xScale:Math.abs(p2.x-p1.x)/2,yScale:Math.abs(p2.y-p1.y)/2,...o});
      }else{
        const sx=a.x+(a.fx?a.w:0),sy=a.y+(a.fy?a.h:0),ex=a.x+(a.fx?0:a.w),ey=a.y+(a.fy?0:a.h),E=U.ends(a);
        const L=(x1,y1,x2,y2,dd)=>pg.drawLine({start:P(x1,y1),end:P(x2,y2),thickness:lw,color:col,lineCap:LineCapStyle.Round,...(dd?dA:{}),...opO});
        L(sx,sy,ex,ey,true);
        const head=(t,tx,ty,fx,fy)=>{
          if(t==='none') return;
          const[h1,h2]=U.arrowPts(fx,fy,tx,ty,8+lw*2);
          if(t==='open'){L(tx,ty,h1[0],h1[1]);L(tx,ty,h2[0],h2[1]);return;}
          const T=P(tx,ty),A1=P(h1[0],h1[1]),A2=P(h2[0],h2[1]);
          pg.drawSvgPath(`M ${T.x} ${-T.y} L ${A1.x} ${-A1.y} L ${A2.x} ${-A2.y} Z`,{x:0,y:0,color:col,borderColor:col,borderWidth:Math.min(lw,1),...(OP<1?{opacity:OP,borderOpacity:OP}:{})});
        };
        head(E.e,ex,ey,sx,sy);head(E.s,sx,sy,ex,ey);
        if(a.shape==='measure'){
          for(const[[x1,y1],[x2,y2]]of U.ticks(sx,sy,ex,ey,5)) L(x1,y1,x2,y2);
          const lab=U.measureLabel(a),tw=font.widthOfTextAtSize(lab,11),[px,py]=vp.convertToPdfPoint((sx+ex)/2-tw/2,(sy+ey)/2-6);
          pg.drawText(lab,{x:px,y:py,size:11,font,color:col,rotate:degrees(rot),...opO});
        }
      }
    }else if(a.type==='dr'){
      const c=U.col(a.color),col=rgb(c[0],c[1],c[2]),pts=a.pts.map(q=>P(a.x+q[0],a.y+q[1])),bm=a.mul?{blendMode:BlendMode.Multiply}:{};
      if(pts.length===1) pg.drawCircle({x:pts[0].x,y:pts[0].y,size:a.lw/2,color:col,...bm,...opO});
      for(let i=1;i<pts.length;i++) pg.drawLine({start:pts[i-1],end:pts[i],thickness:a.lw,color:col,lineCap:LineCapStyle.Round,...bm,...opO});
    }else if(a.type==='lk'){
      const p1=P(a.x,a.y),p2=P(a.x+a.w,a.y+a.h),ctx=doc.context,base={Type:'Annot',Subtype:'Link',Rect:[Math.min(p1.x,p2.x),Math.min(p1.y,p2.y),Math.max(p1.x,p2.x),Math.max(p1.y,p2.y)],Border:[0,0,0],F:4};
      let dict=null;
      if(a.url) dict=ctx.obj({...base,A:{Type:'Action',S:'URI',URI:PL.PDFString.of(String(a.url).replace(/[^\x20-\x7E]/g,ch=>encodeURIComponent(ch)))}});
      else if(pgs[a.dest]) dict=ctx.obj({...base,Dest:[pgs[a.dest].ref,'Fit']});
      if(dict) pg.node.addAnnot(ctx.register(dict));
    }else if(a.type==='tb'){
      if(!a.touched||!a.runs) continue;
      const fam=U.MET[a.fam]?a.fam:'serif',fs={},size=a.size,runs=[];
      for(const r of a.runs){
        const k=(r.i?1:0)+''+(r.b?1:0);
        if(!fs[k]) fs[k]=await doc.embedFont(StandardFonts[U.FN[fam][r.b?(r.i?3:1):(r.i?2:0)]]);
        runs.push({t:U.enc(fs[k],r.t.replace(/\r/g,'')),k});
      }
      const lines=U.layoutRuns(runs,(t,s)=>fs[t.k].widthOfTextAtSize(s,size),a.w,a.indent||0);
      const [A,D]=U.MET[fam],L=a.lh*size,b0=a.y+(L-(A+D)*size)/2+A*size;
      const m=/^#?([0-9a-f]{6})$/i.exec(a.color||'#000000'),c=m?[0,2,4].map(i=>parseInt(m[1].slice(i,i+2),16)/255):[0,0,0];
      lines.forEach((ln,k)=>{
        const ind=k===0?(a.indent||0):0,avail=a.w-ind,by=b0+k*L;
        let x=a.x+ind;
        if(a.align==='center') x=a.x+(a.w-ln.w)/2;
        const spaces=ln.toks.filter(t=>t.sp).length,just=a.align==='justify'&&!ln.hard&&spaces>0&&avail>ln.w,sw=just?(avail-ln.w)/spaces:0;
        const draw=(s,kk,xx)=>{if(!s.trim())return;const [px,py]=vp.convertToPdfPoint(xx,by);pg.drawText(s,{x:px,y:py,size,font:fs[kk],color:rgb(c[0],c[1],c[2]),rotate:degrees(rot)});};
        if(just){for(const t of ln.toks){if(!t.sp)draw(t.t,t.k,x);x+=t.w+(t.sp?sw:0);}}
        else{
          let buf='',bk=null,bx=x;
          for(const t of ln.toks){
            if(bk!==null&&t.k!==bk){draw(buf,bk,bx);buf='';}
            if(buf==='')bx=x;
            bk=t.k;buf+=t.t;x+=t.w;
          }
          if(bk!==null)draw(buf,bk,bx);
        }
      });
    }else if(a.type==='im'){
      const bin=atob(a.src.split(',')[1]),b=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)b[i]=bin.charCodeAt(i);
      const img=await doc.embedPng(b);
      const [px,py]=vp.convertToPdfPoint(a.x,a.y+a.h);
      pg.drawImage(img,{x:px,y:py,width:a.w,height:a.h,rotate:degrees(rot),...opO});
    }
  }
};

/* ====================== ikon (SVG) ====================== */
const svg=(b,s=32)=>`<svg viewBox="0 0 32 32" width="${s}" height="${s}" fill="none" stroke-linejoin="round" stroke-linecap="round">${b}</svg>`;
const PG='<path d="M7 2h13l7 7v21H7z" fill="#fff" stroke="#8a8a8a"/><path d="M20 2v7h7" fill="#e8e8e8" stroke="#8a8a8a"/>';
const bdg=(c,t,fs=12)=>`<rect x="1" y="14" width="18" height="15" rx="2" fill="${c}" stroke="none"/><text x="10" y="${fs>10?26:25}" font-size="${fs}" font-weight="700" fill="#fff" stroke="none" text-anchor="middle" font-family="Segoe UI,Arial,sans-serif">${t}</text>`;
const dI=(c,t,s,fs)=>svg(PG+bdg(c,t,fs),s);
const IC={
  hand:s=>svg('<path d="M9 18V9a2 2 0 0 1 4 0v-1a2 2 0 0 1 4 0v1a2 2 0 0 1 4 0v7l1-3a2 2 0 0 1 3 1l-2 9c-1 4-4 6-8 6-5 0-9-4-9-8z" fill="#fff" stroke="#555" stroke-width="1.6"/>',s),
  edit:s=>svg('<rect x="5" y="5" width="22" height="22" stroke="#e5622c" stroke-width="2" stroke-dasharray="3 3"/><path d="M3 3h6v6H3zM23 23h6v6h-6z" fill="#e5622c" stroke="none"/>',s),
  zoom:s=>svg('<circle cx="13" cy="13" r="8" fill="#fff" stroke="#555" stroke-width="2.5"/><path d="M19 19l9 9" stroke="#555" stroke-width="4"/><path d="M9 13h8M13 9v8" stroke="#555" stroke-width="2"/>',s),
  select:s=>svg('<path d="M6 5v18M3 5h6M3 23h6" stroke="#d9534f" stroke-width="2.2"/><path d="M12 6v19l5-4.5 3.5 7 3-1.5-3.5-7H27z" fill="#444" stroke="#fff" stroke-width="1.2"/>',s),
  text:s=>svg('<text x="16" y="24" font-size="27" font-family="Georgia,serif" fill="#555" text-anchor="middle" stroke="none">T</text><path d="M4 29h24" stroke="#d9534f" stroke-width="2"/>',s),
  sign:s=>svg('<path d="M21 3l7 7-14 14-8 2 2-8z" fill="#666" stroke="#444"/><path d="M3 29h13" stroke="#d9534f" stroke-width="2" stroke-dasharray="2 2"/>',s),
  pdf:s=>dI('#e5622c','+',s,16),
  combine:s=>svg('<path d="M4 6h13v18H4z" fill="#fff" stroke="#888"/><path d="M12 10h13v18H12z" fill="#fff" stroke="#888"/><path d="M25 3v10M20 8h10" stroke="#e5622c" stroke-width="3"/>',s),
  word:s=>dI('#2b579a','W',s), excel:s=>dI('#217346','X',s), ppt:s=>dI('#d24726','P',s),
  image:s=>svg(PG+'<rect x="1" y="14" width="18" height="15" rx="2" fill="#3b8fd9" stroke="none"/><path d="M3 27l5-6 3 3 3-4 4 7z" fill="#fff" stroke="none"/>',s),
  pdfa:s=>dI('#e5622c','A',s),
  rich:s=>dI('#777','R',s), plain:s=>dI('#777','T',s),
  extimg:s=>svg('<rect x="3" y="4" width="20" height="18" rx="2" fill="#fff" stroke="#888"/><path d="M5 20l5-7 3 4 3-4 5 7z" fill="#3b8fd9" stroke="none"/><path d="M20 20h10M26 15l5 5-5 5" stroke="#e5622c" stroke-width="2.4"/>',s),
  portfolio:s=>svg('<path d="M3 3h11v11H3zM18 3h11v11H18zM3 18h11v11H3zM18 18h11v11H18z" fill="#fff" stroke="#888"/><path d="M16 9v14M9 16h14" stroke="#e5622c" stroke-width="4"/>',s),
  hl:s=>svg('<path d="M3 7h26v13l-4 5H3z" fill="#fdd835" stroke="#c9a800"/><text x="14" y="21" font-size="14" font-weight="700" fill="#333" stroke="none" text-anchor="middle" font-family="Segoe UI,Arial,sans-serif">A</text>',s),
  insert:s=>svg(PG+'<path d="M3 21h13M11 15l6 6-6 6" stroke="#e5622c" stroke-width="2.6"/>',s),
  rot:s=>svg('<path d="M25 11a10 10 0 1 0 1.5 9" stroke="#444" stroke-width="2.6"/><path d="M27 4v8h-8" stroke="#444" stroke-width="2.6"/>',s),
  del:s=>svg('<path d="M8 9h16l-1.5 19h-13z" fill="#d0d0d0" stroke="#777"/><path d="M6 9h20M13 9V5h6v4" stroke="#777" stroke-width="2"/>',s),
  extract:s=>svg(PG+'<path d="M16 21h13M24 15l6 6-6 6" stroke="#e5622c" stroke-width="2.6"/>',s),
  find:s=>svg(PG+'<circle cx="14" cy="17" r="7" fill="#fff" stroke="#e5622c" stroke-width="3"/><path d="M19 22l9 8" stroke="#e5622c" stroke-width="4"/>',s),
  tools:s=>svg('<path d="M4 12h10M9 7v10" stroke="#e5622c" stroke-width="3.5"/><path d="M19 29l8-14a5 5 0 1 0-4-4z" fill="#555" stroke="#555"/>',s),
  textbox:s=>svg('<rect x="3" y="7" width="26" height="18" fill="#fff" stroke="#888" stroke-dasharray="3 2"/><text x="16" y="20" font-size="10" font-weight="700" fill="#555" stroke="none" text-anchor="middle" font-family="Segoe UI,Arial,sans-serif">abc</text>',s),
  callout:s=>svg('<rect x="3" y="3" width="26" height="15" fill="#fff59d" stroke="#c9a800"/><path d="M7 18l-3 10 11-10" fill="#fff59d" stroke="#c9a800"/><path d="M7 8h18M7 13h11" stroke="#8a6d00"/>',s),
  note:s=>svg('<path d="M4 4h24v17l-7 7H4z" fill="#fdd835" stroke="#c9a800"/><path d="M21 28v-7h7" fill="#e6bc00" stroke="#c9a800"/><path d="M8 10h16M8 15h16M8 20h8" stroke="#8a6d00" stroke-width="1.6"/>',s),
  strike:s=>svg('<text x="16" y="24" font-size="22" font-weight="700" fill="#555" stroke="none" text-anchor="middle" font-family="Segoe UI,Arial,sans-serif">A</text><path d="M3 16h26" stroke="#d9322d" stroke-width="2.6"/>',s),
  underline:s=>svg('<text x="16" y="23" font-size="22" font-weight="700" fill="#555" stroke="none" text-anchor="middle" font-family="Segoe UI,Arial,sans-serif">A</text><path d="M4 28h24" stroke="#2e7d32" stroke-width="2.6"/>',s),
  replace:s=>svg('<text x="12" y="24" font-size="20" font-weight="700" fill="#888" stroke="none" text-anchor="middle" font-family="Segoe UI,Arial,sans-serif">A</text><path d="M2 16h20" stroke="#d9322d" stroke-width="2.4"/><path d="M20 6l4-3 4 3M24 3v10" stroke="#2e7d32" stroke-width="2"/>',s),
  stamp:s=>svg('<rect x="4" y="21" width="24" height="7" rx="1" fill="#8a8a8a" stroke="#666"/><path d="M11 21v-6a5 5 0 0 1 10 0v6z" fill="#c4262e" stroke="#9a1c22"/><rect x="13" y="4" width="6" height="6" rx="3" fill="#c4262e" stroke="#9a1c22"/>',s),
  shapes:s=>svg('<rect x="3" y="13" width="15" height="15" fill="#fff" stroke="#555" stroke-width="1.8"/><circle cx="21" cy="12" r="8" fill="#fff" stroke="#e5622c" stroke-width="1.8"/>',s),
  drawing:s=>svg('<path d="M6 27l2-7L22 6l4 4L12 24z" fill="#f4b183" stroke="#555"/><path d="M6 27l2-7 4 4z" fill="#555" stroke="#555"/><path d="M3 12q4-7 8 0" stroke="#e5622c" stroke-width="2"/>',s),
  addimage:s=>svg('<rect x="3" y="4" width="20" height="18" rx="2" fill="#fff" stroke="#888"/><path d="M5 20l5-7 3 4 3-4 5 7z" fill="#3b8fd9" stroke="none"/><path d="M24 17v12M18 23h12" stroke="#e5622c" stroke-width="3"/>',s),
  measure:s=>svg('<path d="M3 24L24 3l5 5L8 29z" fill="#fdd835" stroke="#555"/><path d="M8 19l3 3M12 15l2 2M16 11l3 3M20 7l2 2" stroke="#555"/>',s),
  attach:s=>svg('<path d="M22 9L12 20a3 3 0 0 0 4 4l11-12a5 5 0 0 0-7-7L8 17a7 7 0 0 0 10 10" stroke="#555" stroke-width="2.2"/>',s),
  ocr:s=>svg('<text x="15" y="13" font-size="9" font-weight="700" fill="#777" stroke="none" text-anchor="middle" font-family="Segoe UI,Arial,sans-serif">abc</text><circle cx="15" cy="21" r="6" fill="#fff" stroke="#e5622c" stroke-width="2.4"/><path d="M19 25l7 5" stroke="#e5622c" stroke-width="3"/>',s),
  compare:s=>svg('<path d="M3 4h12v22H3z" fill="#fff" stroke="#888"/><path d="M17 8h12v22H17z" fill="#fff" stroke="#888"/><path d="M6 10h6M6 14h6M20 14h6M20 18h6" stroke="#e5622c" stroke-width="1.6"/>',s),
  summ:s=>svg('<rect x="3" y="4" width="26" height="18" fill="#fdd835" stroke="#c9a800"/><path d="M7 10h18M7 15h18" stroke="#8a6d00"/><path d="M8 22v6l6-6" fill="#fdd835" stroke="#c9a800"/>',s),
  migrate:s=>svg('<path d="M5 12h20M19 6l6 6-6 6M27 22H7M13 16l-6 6 6 6" stroke="#e5622c" stroke-width="2.4"/>',s),
  hide:s=>svg('<path d="M3 16c4-7 9-10 13-10s9 3 13 10c-4 7-9 10-13 10S7 23 3 16z" fill="#fff" stroke="#555" stroke-width="1.8"/><circle cx="16" cy="16" r="4" fill="#555" stroke="none"/><path d="M5 27L27 5" stroke="#d9322d" stroke-width="2.4"/>',s),
  crop:s=>svg('<path d="M9 3v20h20M3 9h20v20" stroke="#444" stroke-width="3"/>',s),
  replpg:s=>svg('<path d="M4 4h12v16H4z" fill="#fff" stroke="#888"/><path d="M14 12h12v16H14z" fill="#fff" stroke="#e5622c"/><path d="M20 4l5 4-5 4" stroke="#e5622c" stroke-width="2"/>',s),
  split:s=>svg('<path d="M5 3h22v10H5z" fill="#fff" stroke="#888"/><path d="M5 19h22v10H5z" fill="#fff" stroke="#888"/><path d="M2 16h28" stroke="#e5622c" stroke-width="2.4" stroke-dasharray="4 3"/>',s),
  hf:s=>svg('<path d="M7 2h18v28H7z" fill="#fff" stroke="#888"/><rect x="7" y="2" width="18" height="6" fill="#e5622c" stroke="none"/><rect x="7" y="24" width="18" height="6" fill="#e5622c" stroke="none"/>',s),
  wm:s=>svg('<path d="M7 2h18v28H7z" fill="#fff" stroke="#888"/><text x="16" y="22" font-size="16" font-weight="700" fill="#e5622c" stroke="none" text-anchor="middle" font-family="Segoe UI,Arial,sans-serif" transform="rotate(-35 16 17)">A</text>',s),
  bates:s=>svg('<path d="M7 2h18v28H7z" fill="#fff" stroke="#888"/><path d="M11 8h10M11 12h10M11 16h10" stroke="#bbb"/><text x="16" y="26" font-size="9" font-weight="700" fill="#e5622c" stroke="none" text-anchor="middle" font-family="Segoe UI,Arial,sans-serif">001</text>',s),
  bm:s=>svg('<path d="M8 3h16v26l-8-7-8 7z" fill="#e5622c" stroke="#b34a1e"/>',s),
  link:s=>svg('<path d="M13 19a5 5 0 0 0 7 0l5-5a5 5 0 0 0-7-7l-2 2M19 13a5 5 0 0 0-7 0l-5 5a5 5 0 0 0 7 7l2-2" stroke="#555" stroke-width="2.6"/>',s),
  deskew:s=>svg('<path d="M8 5h14v20H8z" fill="#fff" stroke="#888" transform="rotate(-12 15 15)"/><path d="M4 29h24" stroke="#e5622c" stroke-width="2.4"/>',s),
  redmark:s=>svg('<path d="M7 2h18v28H7z" fill="#fff" stroke="#888"/><rect x="10" y="8" width="12" height="4" fill="#222" stroke="none"/><rect x="10" y="15" width="9" height="4" fill="#222" stroke="none"/><path d="M10 24h12" stroke="#bbb"/>',s),
  redapply:s=>svg('<rect x="3" y="5" width="20" height="5" fill="#222" stroke="none"/><rect x="3" y="13" width="26" height="5" fill="#222" stroke="none"/><path d="M10 25l4 4 9-9" stroke="#2e7d32" stroke-width="3"/>',s),
  redsearch:s=>svg('<rect x="3" y="5" width="18" height="5" fill="#222" stroke="none"/><circle cx="14" cy="19" r="6" fill="#fff" stroke="#e5622c" stroke-width="2.4"/><path d="M19 24l8 6" stroke="#e5622c" stroke-width="3"/>',s),
  rmmeta:s=>svg('<path d="M7 2h18v28H7z" fill="#fff" stroke="#888"/><path d="M11 8h10M11 13h10" stroke="#bbb"/><path d="M12 19l8 8M20 19l-8 8" stroke="#d9322d" stroke-width="3"/>',s),
  pwd:s=>svg('<rect x="7" y="14" width="18" height="14" rx="2" fill="#e5622c" stroke="#b34a1e"/><path d="M11 14v-4a5 5 0 0 1 10 0v4" stroke="#555" stroke-width="2.4"/><circle cx="16" cy="21" r="2" fill="#fff" stroke="none"/>',s),
  cert:s=>svg('<path d="M5 2h16v22H5z" fill="#fff" stroke="#888"/><circle cx="22" cy="21" r="6" fill="#e5622c" stroke="#b34a1e"/><path d="M19 26l-2 5 5-2 5 2-2-5" fill="#e5622c" stroke="#b34a1e"/>',s),
  prof:s=>svg('<rect x="3" y="14" width="14" height="11" rx="2" fill="#e5622c" stroke="#b34a1e"/><rect x="14" y="7" width="14" height="11" rx="2" fill="#f08a5d" stroke="#b34a1e"/>',s),
  sens:s=>svg('<path d="M6 3h20v26H6z" fill="#e8e8e8" stroke="#aaa"/><circle cx="16" cy="14" r="5" fill="none" stroke="#aaa" stroke-width="2.4"/><path d="M10 25h12" stroke="#aaa" stroke-width="2"/>',s),
  certify:s=>svg('<circle cx="16" cy="13" r="9" fill="#3b8fd9" stroke="#2a6db0"/><path d="M11 21l-3 9 8-4 8 4-3-9" fill="#3b8fd9" stroke="#2a6db0"/><path d="M12 13l3 3 6-6" stroke="#fff" stroke-width="2.6"/>',s),
  digid:s=>svg('<rect x="3" y="7" width="26" height="18" rx="2" fill="#fff" stroke="#888"/><circle cx="11" cy="14" r="3" fill="#e5622c" stroke="none"/><path d="M6 22c0-5 10-5 10 0z" fill="#e5622c" stroke="none"/><path d="M19 12h7M19 17h7" stroke="#bbb"/>',s),
  trust:s=>svg('<circle cx="11" cy="11" r="5" fill="#e5622c" stroke="none"/><path d="M2 27c0-9 18-9 18 0z" fill="#e5622c" stroke="none"/><circle cx="22" cy="12" r="4" fill="#9a4a2c" stroke="none"/><path d="M15 27c0-6 15-6 15 0z" fill="#9a4a2c" stroke="none"/>',s),
  aL:s=>svg('<path d="M5 3v26" stroke="#444" stroke-width="2.4"/><rect x="8" y="7" width="19" height="7" fill="#e5622c" stroke="none"/><rect x="8" y="18" width="11" height="7" fill="#e5622c" stroke="none"/>',s),
  aR:s=>svg('<path d="M27 3v26" stroke="#444" stroke-width="2.4"/><rect x="5" y="7" width="19" height="7" fill="#e5622c" stroke="none"/><rect x="13" y="18" width="11" height="7" fill="#e5622c" stroke="none"/>',s),
  aT:s=>svg('<path d="M3 5h26" stroke="#444" stroke-width="2.4"/><rect x="7" y="8" width="7" height="19" fill="#e5622c" stroke="none"/><rect x="18" y="8" width="7" height="11" fill="#e5622c" stroke="none"/>',s),
  aB:s=>svg('<path d="M3 27h26" stroke="#444" stroke-width="2.4"/><rect x="7" y="5" width="7" height="19" fill="#e5622c" stroke="none"/><rect x="18" y="13" width="7" height="11" fill="#e5622c" stroke="none"/>',s),
  aV:s=>svg('<path d="M3 16h26" stroke="#444" stroke-width="1.6" stroke-dasharray="3 2"/><rect x="6" y="6" width="7" height="20" fill="#e5622c" stroke="none"/><rect x="19" y="10" width="7" height="12" fill="#e5622c" stroke="none"/>',s),
  aH:s=>svg('<path d="M16 3v26" stroke="#444" stroke-width="1.6" stroke-dasharray="3 2"/><rect x="6" y="6" width="20" height="7" fill="#e5622c" stroke="none"/><rect x="10" y="19" width="12" height="7" fill="#e5622c" stroke="none"/>',s),
  cV:s=>svg('<rect x="4" y="3" width="24" height="26" fill="#fff" stroke="#888"/><path d="M4 16h24" stroke="#444" stroke-dasharray="3 2"/><rect x="10" y="11" width="12" height="10" fill="#e5622c" stroke="none"/>',s),
  cH:s=>svg('<rect x="3" y="4" width="26" height="24" fill="#fff" stroke="#888"/><path d="M16 4v24" stroke="#444" stroke-dasharray="3 2"/><rect x="11" y="10" width="10" height="12" fill="#e5622c" stroke="none"/>',s),
  cB:s=>svg('<rect x="4" y="4" width="24" height="24" fill="#fff" stroke="#888"/><path d="M16 4v24M4 16h24" stroke="#444" stroke-dasharray="3 2"/><rect x="11" y="11" width="10" height="10" fill="#e5622c" stroke="none"/>',s),
  dV:s=>svg('<rect x="6" y="3" width="20" height="5" fill="#e5622c" stroke="none"/><rect x="6" y="13" width="20" height="5" fill="#e5622c" stroke="none"/><rect x="6" y="24" width="20" height="5" fill="#e5622c" stroke="none"/>',s),
  dH:s=>svg('<rect x="3" y="6" width="5" height="20" fill="#e5622c" stroke="none"/><rect x="13" y="6" width="5" height="20" fill="#e5622c" stroke="none"/><rect x="24" y="6" width="5" height="20" fill="#e5622c" stroke="none"/>',s),
  sH:s=>svg('<rect x="6" y="9" width="14" height="14" fill="#e5622c" stroke="none"/><path d="M26 4v24M23 7l3-3 3 3M23 25l3 3 3-3" stroke="#444" stroke-width="1.8"/>',s),
  sW:s=>svg('<rect x="9" y="6" width="14" height="14" fill="#e5622c" stroke="none"/><path d="M4 26h24M7 23l-3 3 3 3M25 23l3 3-3 3" stroke="#444" stroke-width="1.8"/>',s),
  sB:s=>svg('<rect x="5" y="5" width="16" height="16" fill="#e5622c" stroke="none"/><path d="M12 28h16V12M25 15l3-3 3 3M15 25l-3 3 3 3" stroke="#444" stroke-width="1.8"/>',s),
  spHp:s=>svg('<rect x="3" y="8" width="6" height="16" fill="#e5622c" stroke="none"/><rect x="23" y="8" width="6" height="16" fill="#e5622c" stroke="none"/><path d="M12 16h8M14 13l-3 3 3 3M18 13l3 3-3 3" stroke="#444" stroke-width="1.6"/>',s),
  spHm:s=>svg('<rect x="3" y="8" width="6" height="16" fill="#e5622c" stroke="none"/><rect x="23" y="8" width="6" height="16" fill="#e5622c" stroke="none"/><path d="M11 16h10M18 13l3 3-3 3M14 13l-3 3 3 3" stroke="#444" stroke-width="1.6"/>',s),
  spVp:s=>svg('<rect x="8" y="3" width="16" height="6" fill="#e5622c" stroke="none"/><rect x="8" y="23" width="16" height="6" fill="#e5622c" stroke="none"/><path d="M16 12v8M13 14l3-3 3 3M13 18l3 3 3-3" stroke="#444" stroke-width="1.6"/>',s),
  spVm:s=>svg('<rect x="8" y="3" width="16" height="6" fill="#e5622c" stroke="none"/><rect x="8" y="23" width="16" height="6" fill="#e5622c" stroke="none"/><path d="M16 11v10M13 18l3 3 3-3M13 14l3-3 3 3" stroke="#444" stroke-width="1.6"/>',s),
  tt:s=>svg('<text x="3" y="23" font-size="20" font-weight="700" fill="#555" stroke="none" font-family="Georgia,serif">T</text><text x="17" y="26" font-size="14" font-weight="700" fill="#555" stroke="none" font-family="Georgia,serif">t</text>',s),
  tal:s=>svg('<path d="M4 6h22M4 12h14M4 18h22M4 24h14" stroke="#555" stroke-width="2.2"/>',s),
  tac:s=>svg('<path d="M5 6h22M9 12h14M5 18h22M9 24h14" stroke="#555" stroke-width="2.2"/>',s),
  tar:s=>svg('<path d="M6 6h22M14 12h14M6 18h22M14 24h14" stroke="#555" stroke-width="2.2"/>',s),
  finish:s=>svg('<text x="3" y="25" font-size="26" font-weight="700" fill="#555" stroke="none" font-family="Georgia,serif">T</text><path d="M14 19l5 6 10-14" stroke="#e5622c" stroke-width="3.4"/>',s),
  pdef:s=>svg('<path d="M10 4h12v18H10z" fill="#fff" stroke="#888"/><path d="M20 12l2 4 4.4.6-3.2 3 .8 4.4-4-2.2-4 2.2.8-4.4-3.2-3 4.4-.6z" fill="#e5622c" stroke="none"/>',s),
  pinfo:s=>svg('<rect x="4" y="4" width="20" height="24" fill="#fff" stroke="#888"/><path d="M8 10h12M8 15h12M8 20h7" stroke="#bbb"/><circle cx="23" cy="23" r="6" fill="#e5622c" stroke="none"/><path d="M23 20v4M23 26v.5" stroke="#fff" stroke-width="2"/>',s)
};

/* ====================== ribbon ====================== */
const G=(n,b)=>`<div class="grp"><div class="gb">${b}</div><div class="gn">${n}</div></div>`;
const fld=(l,c,x='')=>`<label class="fld${x}"><span>${l}</span>${c}</label>`;
const opt=(a)=>a.map(([v,t])=>`<option value="${v}">${t}</option>`).join('');
const lab=(l,dd)=>{const t=l.replace(' ','<br>');return dd?(l.includes(' ')?t+' ▾':t+'<br>▾'):t;};
const big=(a,l,i,o={})=>`<button class="b big${o.x?' dim':''}"${o.dis?' disabled':''} ${o.dd?`data-dd="${a}"`:o.t?`data-t="${a}"`:`data-a="${a}"`} title="${l}">${IC[i](32)}<span>${lab(l,o.dd)}</span></button>`;
const sm=(a,l,i,o={})=>`<button class="b rw${o.x?' dim':''}${o.ic?' ico':''}" ${o.dd?`data-dd="${a}"`:o.t?`data-t="${a}"`:`data-a="${a}"`}${o.v!==undefined?` data-v="${o.v}"`:''}${o.n?` data-n="${o.n}"`:''} title="${l}">${IC[i](16)}${o.ic?'':`<span>${l}</span>`}${o.dd?'<i>▾</i>':''}</button>`;
const MODES=`<div class="mt">${sm('hand','Hand','hand',{t:1})}${sm('edit','Edit','edit',{t:1})}<div class="zr2">${sm('zoom','Zoom','zoom',{t:1})}<button class="b ddc" data-dd="zoom" title="Opsi zoom">▾</button></div></div>`;
const RIB={
home:
 G('Tools',big('select','Select','select',{t:1})+big('text','Type Text','text',{t:1})+big('sign','QuickSign','sign',{dd:1}))+
 G('Create',big('create','PDF','pdf',{dd:1})+big('combine','Combine','combine'))+
 G('Convert',big('toword','To Word','word')+big('toexcel','To Excel','excel'))+
 G('Review',big('highlight','Highlight','hl',{t:1}))+
 G('Page Layout',big('insert','Insert','insert',{dd:1})+`<div class="rows">${sm('rotate','Rotate','rot',{dd:1})}${sm('del','Delete','del')}${sm('extract','Extract','extract')}</div>`)+
 G('Document',big('find','Find','find'))+
 G('Favorite Tools',`<div class="favs"></div>`+big('addtools','Add Tools','tools')),
convert:
 G('Formats',big('toword','To Word','word')+big('toexcel','To Excel','excel')+big('topptx','To PowerPoint','ppt')+big('toimage','To Image','image')+big('pdfa','To PDF/A','pdfa',{dd:1})+
   `<div class="rows">${sm('torich','To Rich Text','rich')}${sm('toplain','To Plain Text','plain')}${sm('extimg','Extract Images','extimg')}</div>`)+
 G('Create',big('portfolio','Portfolio','portfolio')),
review:
 G('Markup',big('select','Select','select',{t:1})+big('highlight','Highlight','hl',{t:1})+big('textbox','Text Box','textbox',{t:1})+big('callout','Call Out','callout',{t:1})+big('note','Add Note','note',{t:1})+`<div class="rows">${sm('strike','Cross Out','strike',{t:1})}${sm('underline','Underline','underline',{t:1})}${sm('replace','Replace','replace',{t:1})}</div>`)+
 G('Review Tools',big('stamp','Stamp','stamp',{dd:1})+big('shapes','Shapes','shapes',{dd:1})+big('drawing','Drawing','drawing',{dd:1})+`<div class="rows">${sm('addimage','Add Image','addimage',{dd:1})}${sm('measure','Measure','measure',{dd:1})}${sm('attach','Attach Files','attach')}</div>`)+
 G('Document',big('ocr','OCR','ocr',{x:1})+big('compare','Compare','compare'))+
 G('Comments',`<div class="rows">${sm('summarize','Summarize Comments','summ',{dd:1})}${sm('migrate','Migrate','migrate',{x:1})}${sm('hidecm','Hide Comments','hide')}</div>`),
layout:
 G('Pages',big('insert','Insert','insert',{dd:1})+big('rotate','Rotate','rot',{dd:1})+big('del','Delete','del')+big('crop','Crop','crop')+`<div class="rows">${sm('extract','Extract','extract')}${sm('replpg','Replace','replpg')}${sm('splitdoc','Split','split')}</div>`)+
 G('Page Marks',big('hf','Header &amp; Footer','hf',{dd:1})+big('wm','Watermark','wm',{dd:1})+big('bates','Bates Number','bates',{dd:1}))+
 G('Navigation',big('bm','Bookmark','bm',{dd:1})+big('lk','Link','link',{dd:1}))+
 G('Document',big('deskew','Auto De-skew','deskew',{x:1})),
protect:
 G('Redact',big('markred','Mark for Redaction','redmark',{dd:1})+`<div class="rows">${sm('redapply','Apply All Redactions','redapply')}${sm('redsearch','Search and Redact','redsearch')}${sm('rmmeta','Remove Metadata','rmmeta')}</div>`)+
 G('Secure Document',big('pwd','Password Security','pwd')+big('certsec','Certificate Security','cert',{x:1})+big('profiles','Manage Profiles','prof',{x:1}))+
 G('Sensitivity',big('sens','Sensitivity','sens',{dis:1}))+
 G('Certify Document',big('sign','Sign','sign',{dd:1})+big('certify','Certify','certify',{x:1}))+
 G('Identification',big('digid','My Digital IDs','digid',{x:1})+big('trusted','Trusted Contacts','trust',{x:1})),
props:
 G('General',`<div class="col">${fld('Author:','<input type="text" data-pp="author" class="w120">')}${fld('Subject:','<input type="text" data-pp="subject" class="w120">')}</div>`)+
 G('Appearance',`<div class="col">${fld('Style:',`<select data-pp="style" class="w80">${opt([['solid','Solid'],['dashed','Dashed'],['dotted','Dotted']])}</select>`)}${fld('Opacity [%]:','<input type="number" min="5" max="100" step="5" data-pp="opacity" class="w56">')}</div>`+
   `<div class="col"><label class="b rw clr2" title="Border color"><input type="color" data-pp="border" class="clrin"><i class="sw" data-sw="border"></i><span>Border color ▾</span></label><button class="b rw" data-dd="fillc" title="Fill color"><i class="sw" data-sw="fill"></i><span>Fill color ▾</span></button><input type="color" data-pp="fill" class="clrin"></div>`+
   `<div class="col">${fld('Start:',`<select data-pp="start" class="w80">${opt([['none','None'],['open','Open'],['closed','Closed']])}</select>`)}${fld('End:',`<select data-pp="end" class="w80">${opt([['none','None'],['open','Open'],['closed','Closed']])}</select>`)}</div>`+
   `<div class="col">${fld('Thickness:','<input type="number" min="0.5" max="20" step="0.5" data-pp="thick" class="w56">')}${fld('Leader Length:','<input type="number" min="5" max="300" step="5" data-pp="ll" class="w56">')}</div>`+
   `<div class="col">${fld('Leader Offset:','<input type="number" min="0" max="1000" step="5" data-pp="lo" class="w56">')}${fld('Leader Extension:','<input type="number" disabled class="w56" title="Belum didukung">',' off')}</div>`)+
 G('More',`<div class="rows">${sm('pdefault','Set As Default','pdef')}${sm('pinfo','Properties','pinfo')}${sm('delsel','Delete','del')}</div>`),
align:
 G('Align',`<div class="cols">${sm('alnL','Left','aL',{n:2})}${sm('alnR','Right','aR',{n:2})}${sm('alnT','Top','aT',{n:2})}${sm('alnB','Bottom','aB',{n:2})}${sm('alnV','Vertically','aV',{n:2})}${sm('alnH','Horizontally','aH',{n:2})}</div>`)+
 G('Center',`<div class="rows">${sm('ctrV','Vertically','cV',{n:1})}${sm('ctrH','Horizontally','cH',{n:1})}${sm('ctrB','Both','cB',{n:1})}</div>`)+
 G('Distribute',`<div class="rows">${sm('distV','Vertically','dV',{n:3})}${sm('distH','Horizontally','dH',{n:3})}</div>`)+
 G('Size',`<div class="rows">${sm('szH','Height','sH',{n:2})}${sm('szW','Width','sW',{n:2})}${sm('szB','Both','sB',{n:2})}</div>`)+
 G('Spacing',`<div class="cols2">${sm('spHp','Tambah jarak horizontal','spHp',{n:2,ic:1})}${sm('spHm','Kurangi jarak horizontal','spHm',{n:2,ic:1})}${sm('spVp','Tambah jarak vertikal','spVp',{n:2,ic:1})}${sm('spVm','Kurangi jarak vertikal','spVm',{n:2,ic:1})}</div>`),
format:
 G('Font',`<div class="col"><select data-pp="font" class="w150">${opt([['helv','Arial'],['times','Times New Roman'],['cour','Courier New']])}</select><div class="rowx"><button class="b" data-dd="ftt" title="Gaya huruf">${IC.tt(20)}<i>▾</i></button><label class="b clr" title="Warna teks"><input type="color" data-pp="tcolor" class="clrin"><span class="Ac">A</span><i class="sw bar" data-sw="tcolor"></i></label><select data-pp="size" class="w56">${opt([8,9,10,11,12,14,16,18,20,24,28,32,36,48,72].map(x=>[x,x]))}</select></div></div>`)+
 G('Alignment',`<div class="rowx pad">${sm('talign','Rata kiri','tal',{v:'l',ic:1})}${sm('talign','Rata tengah','tac',{v:'c',ic:1})}${sm('talign','Rata kanan','tar',{v:'r',ic:1})}</div>`)+
 G('Spacing',`<div class="col">${fld('Character','<input type="number" step="0.5" data-pp="cs" class="w56">')}${fld('Word','<input type="number" step="1" data-pp="ws" class="w56">')}${fld('Line',`<select data-pp="ls" class="w80">${opt([['1.2','Normal'],['1','1.0'],['1.5','1.5'],['2','2.0'],['3','3.0']])}</select>`)}</div>`)+
 G('Typing',big('tfinish','Finish','finish'))
};
const DD={
  create:()=>[['PDF kosong baru','blank'],['PDF dari gambar…','fromimg'],['PDF dari teks (.txt)…','fromtxt']],
  sign:sg=>[['Gambar tanda tangan baru…','signnew'],...(sg?[['Pakai tanda tangan tersimpan','signuse'],['Hapus tanda tangan tersimpan','signdel']]:[])],
  insert:()=>[['Halaman kosong (setelah halaman ini)','insblank'],['Dari berkas PDF…','inspdf'],['Dari gambar…','insimg']],
  rotate:()=>[['Putar kanan 90° (halaman ini)','rot:90'],['Putar kiri 90° (halaman ini)','rot:-90'],['Putar 180° (halaman ini)','rot:180'],'-',['Putar semua halaman ke kanan','rotall:90'],['Putar semua halaman ke kiri','rotall:-90']],
  zoom:()=>[['Perbesar','zin'],['Perkecil','zout'],'-',['50%','z:50'],['75%','z:75'],['100%','z:100'],['150%','z:150'],['200%','z:200'],'-',['Sesuaikan lebar','zfitw'],['Sesuaikan halaman','zfitp']],
  pdfa:()=>[['PDF/A-1b','pdfa'],['PDF/A-2b','pdfa'],['PDF/A-3b','pdfa']],
  stamp:()=>[['DISETUJUI','stamp:DISETUJUI'],['DITOLAK','stamp:DITOLAK'],['DRAF','stamp:DRAF'],['RAHASIA','stamp:RAHASIA'],['FINAL','stamp:FINAL'],['SELESAI','stamp:SELESAI'],['SALINAN','stamp:SALINAN'],'-',['Tanggal hari ini','stamp:TGL']],
  shapes:()=>[['Persegi panjang','shape:rect'],['Elips','shape:ell'],['Garis','shape:line'],['Panah','shape:arrow']],
  drawing:()=>[['Pena hitam','pen:black'],['Pena merah','pen:red'],['Pena biru','pen:blue'],['Marker kuning','pen:marker']],
  addimage:()=>[['Dari berkas gambar…','addimg'],['Tempel dari clipboard','pasteimg']],
  measure:()=>[['Jarak (cm)','measure:cm'],['Jarak (inci)','measure:in'],['Jarak (poin)','measure:pt']],
  summarize:()=>[['Tampilkan ringkasan…','sumshow'],['Ekspor ringkasan (.txt)','sumexp']],
  hf:()=>[['Tambah header & footer…','hf:full'],['Tambah nomor halaman…','hf:num']],
  wm:()=>[['Teks kustom…','wm:custom'],'-',['DRAF','wm:draf'],['RAHASIA','wm:rahasia'],['SALINAN','wm:salinan'],['JANGAN DISALIN','wm:jangan']],
  bates:()=>[['Tambah nomor Bates…','bates']],
  bm:()=>[['Tambah bookmark di halaman ini…','bmadd'],['Lihat daftar bookmark…','bmlist']],
  lk:()=>[['Tautan ke URL (gambar kotak)','link:url'],['Tautan ke halaman (gambar kotak)','link:page']],
  fillc:()=>[['Pilih warna…','fillpick'],['Tanpa isi','fillnone']],
  ftt:()=>[['Normal','ftt:n'],['Tebal','ftt:b'],['Miring','ftt:i'],['Tebal + miring','ftt:bi']],
  markred:()=>[['Tandai area (seret)','redarea'],['Tandai teks (sorot teks)','redtext'],'-',['Hapus semua tanda redaksi','redclear']]
};
const FAVS=[['rot:90','Putar kanan','rot'],['rot:-90','Putar kiri','rot'],['del','Hapus halaman','del'],['extract','Ekstrak halaman','extract'],['find','Cari','find'],['combine','Gabungkan','combine'],['blank','PDF kosong baru','pdf'],
  ['toword','Ke Word','word'],['toexcel','Ke Excel','excel'],['topptx','Ke PowerPoint','ppt'],['toimage','Ke gambar','image'],['torich','Ke Rich Text','rich'],['toplain','Ke Plain Text','plain'],['extimg','Ekstrak gambar','extimg'],['print','Cetak','pdf']];

const CSS=`
.pd{--acc:#b7282e;--ln:#d4d4d4;height:100%;display:flex;flex-direction:column;font:12px "Segoe UI",system-ui,sans-serif;color:#444;background:#fff;position:relative;outline:none}
.pd *{box-sizing:border-box}
.pd button,.pd select,.pd input{font:inherit;color:inherit}
.pd [hidden]{display:none!important}
.pd .qat{height:26px;display:flex;align-items:center;justify-content:center;position:relative;font-size:13px;flex:none}
.pd .qat .l{position:absolute;left:6px;display:flex;gap:2px;align-items:center}
.pd .qat .w{background:var(--acc);color:#fff;font-weight:700;width:18px;height:16px;display:grid;place-items:center;font-size:11px;margin-right:6px}
.pd .qat button{background:none;border:0;cursor:pointer;width:24px;height:22px;color:#555}
.pd .qat button:hover{background:#e1e1e1}
.pd .qat button:disabled{opacity:.35;cursor:default;background:none}
.pd .qat .fn{max-width:calc(100% - 240px);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.pd .tabs{display:flex;height:26px;border-bottom:1px solid var(--ln);background:#fff;overflow-x:auto;scrollbar-width:none;flex:none}
.pd .tabs::-webkit-scrollbar{display:none}
.pd .tab{background:none;border:1px solid transparent;border-bottom:0;padding:0 14px;cursor:pointer;height:26px;font-size:12px;flex:none;white-space:nowrap}
.pd .tab:hover{color:var(--acc)}
.pd .tab.on{border-color:var(--ln);background:#fff;color:var(--acc);height:27px}
.pd .tab.file{background:var(--acc);color:#fff;padding:0 22px;margin-right:6px;border:0}
.pd .rib{height:92px;background:#fff;border-bottom:1px solid var(--ln);display:flex;padding:3px 4px;overflow-x:auto;white-space:nowrap;flex:none;-webkit-overflow-scrolling:touch}
.pd .ribbon{display:flex}
.pd .mt{display:flex;flex-direction:column;gap:2px;padding:2px 8px 0 2px;border-right:1px solid #e3e3e3;width:88px;flex:none}
.pd .mt .b.rw{width:100%}
.pd .zr2{display:flex}.pd .zr2 .b.rw{flex:1;width:auto}
.pd .ddc{width:16px;min-width:16px;height:22px;border:1px solid transparent;background:none;cursor:pointer;font-size:9px;padding:0}
.pd .ddc:hover{background:#e6e6e6;border-color:#cfcfcf}
.pd .grp{display:flex;flex-direction:column;border-right:1px solid #e3e3e3;padding:0 6px}
.pd .gb{flex:1;display:flex;gap:4px;align-items:flex-start}
.pd .gn{text-align:center;font-size:11px;color:#666;height:16px}
.pd .rows{display:flex;flex-direction:column;gap:2px}
.pd .favs{display:flex;flex-direction:column;gap:2px}
.pd .b{min-width:24px;height:24px;border:1px solid transparent;background:none;display:inline-flex;align-items:center;justify-content:center;gap:5px;padding:0 4px;cursor:pointer;color:#444}
.pd .b:hover{background:#e6e6e6;border-color:#cfcfcf}
.pd .b.on{background:#d9d9d9;border-color:#a8a8a8}
.pd .b.big{flex-direction:column;height:66px;min-width:50px;gap:2px;font-size:11px;line-height:12px;justify-content:flex-start;padding-top:3px}
.pd .b.big span{text-align:center}
.pd .b.rw{height:22px;justify-content:flex-start;white-space:nowrap}
.pd .b i{font-style:normal;font-size:9px}
.pd .b svg{flex:none}
.pd .pmenu{position:absolute;z-index:15;background:#fff;border:1px solid #aaa;box-shadow:0 2px 8px rgba(0,0,0,.25);padding:2px 0;min-width:190px}
.pd .pmenu button{display:block;width:100%;text-align:left;background:none;border:0;padding:6px 14px;cursor:pointer;white-space:nowrap}
.pd .pmenu button:hover{background:#f1d7d9}
.pd .pmenu hr{border:0;border-top:1px solid #ddd;margin:2px 0}
.pd .desk{flex:1;overflow:auto;background:#e9e9e9;position:relative;outline:none}
.pd .pgs{width:max-content;min-width:100%;padding:12px 16px 24px;display:flex;flex-direction:column;align-items:center;gap:10px}
.pd .pg{position:relative;background:#fff;box-shadow:0 0 0 1px #ababab,0 2px 8px rgba(0,0,0,.15);flex:none}
.pd .pg canvas{position:absolute;left:0;top:0;width:100%;height:100%;display:block}
.pd .tl{position:absolute;inset:0;overflow:hidden;line-height:1;text-size-adjust:none}
.pd .tl :is(span,br){color:transparent;position:absolute;white-space:pre;cursor:text;transform-origin:0 0}
.pd .tl .markedContent{display:contents;position:static}
.pd .tl ::selection{background:rgba(0,90,200,.35)}
.pd .tl .hit{background:rgba(255,235,0,.55);border-radius:2px}
.pd .tl .hit.cur{background:rgba(255,140,0,.7)}
.pd .annl{position:absolute;left:0;top:0;transform-origin:0 0;pointer-events:none}
.pd .an{position:absolute}
.pd .tx{font:12px/1.2 Helvetica,Arial,"Liberation Sans",sans-serif;white-space:pre;color:#000;min-width:8px;min-height:1.2em;outline:1px dashed transparent}
.pd .tx.sel,.pd .tx:focus{outline:1px dashed var(--acc)}
.pd .im{cursor:move;outline:1px solid transparent}.pd .im.sel{outline:1px solid var(--acc)}
.pd .im img{width:100%;height:100%;display:block;pointer-events:none}
.pd .rs{position:absolute;right:-5px;bottom:-5px;width:10px;height:10px;background:var(--acc);cursor:nwse-resize;display:none}
.pd .im.sel .rs,.pd .sh.sel .rs{display:block}
.pd .hll{position:absolute;inset:0;pointer-events:none}
.pd .hll i{position:absolute;background:#ffe600;mix-blend-mode:multiply;border-radius:2px;pointer-events:none}
.pd .cov{position:absolute;background:#fff}
.pd .annl>.cov,.pd .annl>.hv{pointer-events:none!important}
.pd .hv{position:absolute;outline:1px dashed #e5622c;outline-offset:2px;background:rgba(229,98,44,.08)}
.pd .hv.bad{outline-color:#9a9a9a;background:rgba(120,120,120,.08)}
.pd .tb{font-family:"Times New Roman","Liberation Serif",Tinos,Times,serif;color:#000;min-width:12px;outline:1px solid transparent;outline-offset:1px}
.pd .tb.sans{font-family:Arial,Helvetica,"Liberation Sans",sans-serif}
.pd .tb.mono{font-family:"Courier New","Liberation Mono",monospace}
.pd .tb .tbt{white-space:pre-wrap;overflow-wrap:break-word;min-height:1em;outline:none;cursor:text}
.pd .tb.sel{outline-color:#e5622c}
.pd .tb .hd,.pd .tb .mv{display:none;position:absolute;z-index:2}
.pd .tb.sel .hd{display:block;width:9px;height:9px;background:#e5622c;border:1px solid #fff}
.pd .tb.sel .mv{display:block}
.pd .hd[data-h=nw]{left:-5px;top:-5px;cursor:nwse-resize}
.pd .hd[data-h=n]{left:calc(50% - 4px);top:-5px;cursor:ns-resize}
.pd .hd[data-h=ne]{right:-5px;top:-5px;cursor:nesw-resize}
.pd .hd[data-h=e]{right:-5px;top:calc(50% - 4px);cursor:ew-resize}
.pd .hd[data-h=se]{right:-5px;bottom:-5px;cursor:nwse-resize}
.pd .hd[data-h=s]{left:calc(50% - 4px);bottom:-5px;cursor:ns-resize}
.pd .hd[data-h=sw]{left:-5px;bottom:-5px;cursor:nesw-resize}
.pd .hd[data-h=w]{left:-5px;top:calc(50% - 4px);cursor:ew-resize}
.pd .mv[data-m=t]{left:0;right:0;top:-5px;height:8px;cursor:move}
.pd .mv[data-m=b]{left:0;right:0;bottom:-5px;height:8px;cursor:move}
.pd .mv[data-m=l]{top:0;bottom:0;left:-5px;width:8px;cursor:move}
.pd .mv[data-m=r]{top:0;bottom:0;right:-5px;width:8px;cursor:move}
.pd .t-edit .an.tb{cursor:text}
.pd .t-edit .an:not(.tx){cursor:move}.pd .t-edit .tx{cursor:move}.pd .t-text .tx{cursor:text}
.pd .t-edit .annl>*,.pd .t-text .annl>*{pointer-events:auto}
.pd .t-hand .tl,.pd .t-zoom .tl,.pd .t-edit .tl,.pd .t-text .tl,.pd .t-sign .tl{pointer-events:none}
.pd .t-hand{cursor:grab}.pd .t-hand.drag{cursor:grabbing}.pd .t-zoom{cursor:zoom-in}.pd .t-text,.pd .t-sign{cursor:crosshair}
.pd .t-hand .pg,.pd .t-zoom .pg{user-select:none}
.pd .find{position:absolute;right:24px;top:150px;z-index:5;background:#fff;border:1px solid #aaa;box-shadow:0 2px 8px rgba(0,0,0,.25);padding:8px 10px;display:flex;gap:6px;align-items:center}
.pd .find input{height:24px;border:1px solid #aaa;padding:0 4px;width:170px}
.pd .find button{height:24px;border:1px solid #aaa;background:#f3f3f3;cursor:pointer;padding:0 8px}
.pd .find .fc{min-width:48px;text-align:center;color:#666}
.pd .status{background:var(--acc);color:#fff;height:24px;display:flex;align-items:center;justify-content:space-between;padding:0 8px;font-size:11px;flex:none}
.pd .status .s{display:flex;gap:10px;align-items:center}
.pd .status button{background:none;border:0;color:#fff;cursor:pointer;min-width:22px;height:22px;font-size:12px}
.pd .status button:hover{background:rgba(255,255,255,.2)}
.pd .status input[type=number]{width:44px;height:18px;border:0;text-align:center;-moz-appearance:textfield;background:#fff;color:#333}
.pd .status input[type=number]::-webkit-inner-spin-button{display:none}
.pd .status input[type=range]{width:130px}
.pd .menu{position:absolute;z-index:10;top:52px;left:0;bottom:24px;width:260px;background:var(--acc);color:#fff;padding:8px 0}
.pd .menu button{display:block;width:100%;background:none;border:0;color:#fff;text-align:left;padding:12px 24px;cursor:pointer;font-size:14px}
.pd .menu button:hover{background:rgba(255,255,255,.18)}
.pd .menu hr{border:0;border-top:1px solid rgba(255,255,255,.3);margin:6px 0}
.pd .menu{overflow:auto}
.pd .menu .ofl .lbl{font-size:11px;opacity:.75;padding:4px 24px;text-transform:uppercase;letter-spacing:.04em}
.pd .menu .ofl .of{display:flex;align-items:center}
.pd .menu .ofl .of.on{background:rgba(255,255,255,.22)}
.pd .menu .ofl .ofn{flex:1;width:auto;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;padding:10px 8px 10px 24px;font-size:13px}
.pd .menu .ofl .ofx{flex:none;width:40px;padding:10px 0;text-align:center;font-size:18px;line-height:1}
.pd .menu .ofl .ofx:hover{background:rgba(255,255,255,.3)}
.pd .toast{position:absolute;right:16px;bottom:36px;background:#323130;color:#fff;padding:8px 14px;z-index:40;max-width:min(360px,calc(100% - 32px))}
.pd .ov{position:absolute;inset:0;background:rgba(0,0,0,.25);z-index:30;display:grid;place-items:center}
.pd .dg{background:#fff;border:1px solid #888;box-shadow:0 4px 18px rgba(0,0,0,.3);padding:14px 16px;min-width:min(340px,92vw);max-width:92%}
.pd .dg h3{margin:0 0 10px;font-size:14px;font-weight:600;color:var(--acc)}
.pd .db{display:flex;flex-direction:column;gap:8px;max-height:60vh;overflow:auto}
.pd .db label{display:flex;gap:8px;align-items:center}
.pd .db input[type=text],.pd .db input[type=password],.pd .db select{height:26px;border:1px solid #aaa;padding:0 4px;max-width:100%}
.pd .db canvas{border:1px solid #aaa;background:#fff;touch-action:none;max-width:100%;cursor:crosshair}
.pd .dbt{display:flex;gap:6px;justify-content:flex-end;margin-top:12px}
.pd .dbt button{height:26px;border:1px solid #aaa;background:#f3f3f3;cursor:pointer;padding:0 14px}
.pd .dbt .ok{background:var(--acc);border-color:var(--acc);color:#fff}
@media (pointer:coarse){
 .pd .qat{height:36px}.pd .qat button{width:34px;height:32px;font-size:16px}
 .pd .tabs{height:36px}.pd .tab{height:36px;padding:0 16px;font-size:13px}.pd .tab.on{height:37px}
 .pd .rib{height:112px}.pd .b.rw{height:30px}.pd .b.big{height:84px;min-width:56px}
 .pd .status{height:30px}.pd .status button{min-width:30px;height:28px;font-size:16px}
 .pd .menu{top:72px;bottom:30px}.pd .menu button{padding:16px 24px}
 .pd .find input,.pd .db input,.pd .db select{font-size:16px;height:34px}
}
@media (max-width:700px){
 .pd .qat .fn{max-width:calc(100% - 150px)}
 .pd .menu{width:min(260px,86vw)}
 .pd .status input[type=range]{width:70px}
 .pd .find{left:8px;right:8px;top:auto;bottom:40px}.pd .find input{width:100%;min-width:0}
}
.pd .b.dim{color:#8a8a8a}.pd .b.dim svg{opacity:.55}
.pd .b:disabled{opacity:.4;cursor:default}.pd .b:disabled:hover{background:none;border-color:transparent}
.pd .tx.bx{padding:3px;border:var(--tbw,1px) var(--tbs,solid) var(--tbc,#444);background:var(--tbg,transparent);min-width:80px}
.pd .tx.nt{min-width:100px}
.pd .tx.co::after{content:"";position:absolute;left:calc(var(--tlo,0px) - var(--tbw,1px));top:100%;width:var(--tll,40px);border-top:var(--tbw,1px) var(--tbs,solid) var(--tbc,#444);transform-origin:0 0;transform:rotate(135deg)}
.pd .tab.ctx{background:#fdeee5;color:#b34a1e;border-color:#f3cdb7}.pd .tab.ctx.on{background:#fff;color:var(--acc);border-color:var(--ln)}
.pd .col{display:flex;flex-direction:column;gap:2px;padding:2px 4px;justify-content:flex-start}
.pd .cols{display:grid;grid-auto-flow:column;grid-template-rows:repeat(3,22px);column-gap:6px;padding:2px 4px}
.pd .cols2{display:grid;grid-auto-flow:column;grid-template-rows:repeat(2,26px);column-gap:4px;row-gap:4px;padding:6px 4px 0}
.pd .rowx{display:flex;align-items:center;gap:4px}.pd .rowx.pad{padding-top:8px}
.pd .fld{display:flex;align-items:center;gap:4px;font-size:12px;height:21px;white-space:nowrap}
.pd .fld.off{color:#9a9a9a}
.pd .fld input,.pd .fld select,.pd .col>select,.pd .rowx>select{height:20px;border:1px solid #bdbdbd;background:#fff;font:inherit;font-size:12px;padding:0 3px;min-width:0}
.pd .fld input:disabled,.pd .fld select:disabled,.pd .col>select:disabled,.pd .rowx>select:disabled{background:#eee;color:#999}
.pd .w56{width:56px}.pd .w80{width:80px}.pd .w120{width:120px}.pd .w150{width:150px}
.pd .clrin{position:absolute;width:1px;height:1px;opacity:0;pointer-events:none}
.pd .b.clr{flex-direction:column;gap:0;height:30px;position:relative}.pd .Ac{font-weight:700;font-size:15px;line-height:15px;color:#333}
.pd .sw{display:inline-block;width:14px;height:14px;border:1px solid #888;background:#000}.pd .sw.bar{height:4px;width:18px;border:0}
.pd .sw.none{background:linear-gradient(135deg,#fff 44%,#d00 44%,#d00 56%,#fff 56%)!important}
.pd .b.clr2{position:relative}
.pd .b.ico{width:28px;min-width:28px;justify-content:center}
.pd .b:disabled{opacity:.4;cursor:default}
.pd .hll i.sel{outline:2px solid #1a73e8}
.pd .multi .rs{display:none!important}
.pd .t-select .annl>.an{pointer-events:auto}.pd .t-select .an:not(.tb):not(.tx){cursor:move}.pd .t-select .tx{cursor:move}

.pd .sh,.pd .dr{outline:1px solid transparent;cursor:move}.pd .sh.sel,.pd .dr.sel,.pd .lk.sel{outline:1px dashed var(--acc)}
.pd .dr svg{position:absolute;left:0;top:0;width:100%;height:100%;overflow:visible;pointer-events:none}
.pd .lk{border:1px dashed #1a73e8;background:rgba(26,115,232,.12);cursor:move}
.pd .dprev{position:absolute;border:1px dashed #1a73e8;background:rgba(26,115,232,.1);pointer-events:none}.pd .dprev.redarea{border-color:#d90000;background:rgba(217,0,0,.18)}
.pd .hll i.ul{background:none;mix-blend-mode:normal;border-radius:0;border-bottom:calc(var(--s)*1.2px) solid #d90000}
.pd .hll i.st{background:none;mix-blend-mode:normal;border-radius:0}
.pd .hll i.st::after{content:"";position:absolute;left:0;right:0;top:55%;border-top:calc(var(--s)*1.2px) solid #d90000}
.pd .hll i.rd{background:rgba(217,0,0,.18);mix-blend-mode:normal;border-radius:0;outline:calc(var(--s)*1px) solid #d90000}
.pd .desk.hidecm .cm{display:none}
.pd .t-textbox .tl,.pd .t-callout .tl,.pd .t-note .tl,.pd .t-stamp .tl,.pd .t-shape .tl,.pd .t-pen .tl,.pd .t-measure .tl,.pd .t-link .tl,.pd .t-redarea .tl{pointer-events:none}
.pd .t-textbox,.pd .t-callout,.pd .t-note,.pd .t-stamp,.pd .t-shape,.pd .t-pen,.pd .t-measure,.pd .t-link,.pd .t-redarea{cursor:crosshair}
.pd .t-shape .pg,.pd .t-pen .pg,.pd .t-measure .pg,.pd .t-link .pg,.pd .t-redarea .pg{touch-action:none;user-select:none}
.pd .t-textbox .annl>*,.pd .t-callout .annl>*,.pd .t-note .annl>*{pointer-events:auto}
.pd .t-textbox .tx,.pd .t-callout .tx,.pd .t-note .tx{cursor:text}
.pd .db .bmi{border:0;background:none;cursor:pointer;padding:5px 8px;font:inherit;text-align:left}.pd .db .bmi:hover{background:#f1f1f1}
.pd-print{display:none}
@media print{body>*:not(.pd-print){display:none!important}.pd-print{display:block}.pd-print img{display:block;width:100%;break-after:page}@page{margin:0}}
`;

/* ====================== aplikasi ====================== */
function mount(root,opts={}){
  if(!document.getElementById('pd-css')){const s=document.createElement('style');s.id='pd-css';s.textContent=CSS;document.head.appendChild(s);}
  let name='Untitled1.pdf',bytes=null,pdfDoc=null,pages=[],annots=[],hist=[],hi=-1;
  let zoom=100,S=4/3,autoFit=true,tool='select',dirty=false,pristine=false,cur=0,selId=null,idc=0,locked=false,io=null,ddBtn=null;
  let sign=store.get('sign',null),favs=store.get('favs',[]);
  let fq='',fm=[],fi=-1,ft=0;
  let stampSel='DISETUJUI',shapeSel='rect',penSel='black',measureUnit='cm',linkMode='url',hidecm=false;
  let sel=[],editId=null,curEditDone=null,lastMain='home',ctxOn=false,edShown=false,quiet=false,author=store.get('author',''),defs=store.get('defs',{});
  const SELTOOL={highlight:'hl',underline:'ul',strike:'st',replace:'rp',redtext:'rd'};
  const TXT={text:{size:12,p:{}},textbox:{size:12,box:1,p:{box:'#444444',cm:true}},callout:{size:12,box:1,p:{box:'#444444',co:true,cm:true}},note:{size:11,bg:1,p:{bg:'#fff59d',box:'#c9a800',cm:true}}};
  const DRAWT={shape:1,pen:1,measure:1,link:1,redarea:1};
  const isRd=a=>a.type==='hl'&&a.kind==='rd';
  const textCache=new Map(),ZL=[10,25,50,75,100,125,150,200,300,400,500];

  root.innerHTML=`<div class="pd" tabindex="-1">
    <div class="qat"><div class="l"><span class="w">A</span><button data-a="save" title="Simpan (Ctrl+S)">💾</button><button data-a="undo" title="Urungkan (Ctrl+Z)">↶</button><button data-a="redo" title="Ulangi (Ctrl+Y)">↷</button></div><span class="fn"></span></div>
    <div class="tabs"><button class="tab file" data-a="menu">File</button><button class="tab on" data-tab="home">Home</button><button class="tab" data-tab="convert">Convert</button><button class="tab" data-tab="review">Review</button><button class="tab" data-tab="layout">Page Layout</button><button class="tab" data-tab="protect">Protect</button><button class="tab ctx" data-tab="props" hidden>Properties</button><button class="tab ctx" data-tab="align" hidden>Alignment</button><button class="tab ctx" data-tab="format" hidden>Format</button></div>
    <div class="rib">${MODES}<div class="ribbon" data-pane="home">${RIB.home}</div><div class="ribbon" data-pane="convert" hidden>${RIB.convert}</div><div class="ribbon" data-pane="review" hidden>${RIB.review}</div><div class="ribbon" data-pane="layout" hidden>${RIB.layout}</div><div class="ribbon" data-pane="protect" hidden>${RIB.protect}</div><div class="ribbon" data-pane="props" hidden>${RIB.props}</div><div class="ribbon" data-pane="align" hidden>${RIB.align}</div><div class="ribbon" data-pane="format" hidden>${RIB.format}</div></div>
    <div class="desk t-select" tabindex="-1"><div class="pgs"></div></div>
    <div class="find" hidden><input type="text" class="fq" placeholder="Cari teks…"><span class="fc"></span><button data-a="fprev" title="Sebelumnya">▲</button><button data-a="fnext" title="Berikutnya">▼</button><button data-a="fclose">Tutup</button></div>
    <div class="status"><div class="s"><span><button data-a="first" title="Halaman pertama">|◀</button><button data-a="prev" title="Sebelumnya">◀</button> <input type="number" class="pi" min="1" value="1"> <span class="of">OF 1</span> <button data-a="next" title="Berikutnya">▶</button><button data-a="last" title="Halaman terakhir">▶|</button></span></div>
      <div class="s"><span><button data-a="zout">−</button><input type="range" min="10" max="400" value="100" class="zs"><button data-a="zin">+</button> <span class="zv">100%</span></span></div></div>
    <div class="menu" hidden>
      <button data-a="blank">📄  PDF kosong baru</button><button data-a="open">📂  Buka…</button>
      <button data-a="save">💾  Simpan sebagai .pdf</button><button data-a="print">🖨  Cetak</button>
      <hr><button data-a="home">⬅  Kembali ke menu utama</button><div class="ofl"></div></div>
  </div>`;
  const $=s=>root.querySelector(s),$$=s=>[...root.querySelectorAll(s)];
  const pd=$('.pd'),desk=$('.desk'),pgs=$('.pgs'),menu=$('.menu'),fd=$('.find'),fqIn=$('.fq');
  const toast=m=>{const t=document.createElement('div');t.className='toast';t.textContent=m;pd.appendChild(t);setTimeout(()=>t.remove(),3500);};
  const note=m=>{const t=document.createElement('div');t.className='toast';t.textContent=m;pd.appendChild(t);return{set:x=>t.textContent=x,end:()=>t.remove()};};
  const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
  const setName=n=>{name=n;$('.fn').textContent=n+' - PDF';document.title=$('.fn').textContent;};
  const base=()=>name.replace(/\.pdf$/i,'');
  const dl=(blob,fn)=>{const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=fn;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),4000);};
  const pickFiles=(accept,multi)=>new Promise(res=>{const i=document.createElement('input');i.type='file';i.accept=accept;i.multiple=!!multi;i.onchange=()=>res([...i.files]);i.oncancel=()=>res([]);i.click();});
  const clone=a=>JSON.parse(JSON.stringify(a));

  function dialog(title,html,ok='OK',onOpen){
    return new Promise(res=>{
      const ov=document.createElement('div');ov.className='ov';
      ov.innerHTML=`<div class="dg"><h3></h3><div class="db">${html}</div><div class="dbt"><button class="ok">${ok}</button><button class="cx">Batal</button></div></div>`;
      ov.querySelector('h3').textContent=title;
      const body=ov.querySelector('.db'),end=v=>{ov.remove();res(v);};
      ov.querySelector('.ok').onclick=()=>end(body);ov.querySelector('.cx').onclick=()=>end(null);
      ov.addEventListener('keydown',e=>{if(e.key==='Escape')end(null);else if(e.key==='Enter'&&e.target.tagName==='INPUT'&&e.target.type==='text'){e.preventDefault();end(body);}});
      pd.appendChild(ov);onOpen&&onOpen(body);
      const f=body.querySelector('input,select');f&&f.focus();
    });
  }

  /* ---------- riwayat ---------- */
  function push(){hist=hist.slice(0,hi+1);hist.push({bytes,annots:clone(annots.filter(a=>!(a.type==='tb'&&!a.touched)))});if(hist.length>40)hist.shift();hi=hist.length-1;dirty=true;btns();}
  function btns(){const u=$('[data-a=undo]'),r=$('[data-a=redo]');if(u)u.disabled=hi<=0;if(r)r.disabled=hi>=hist.length-1;}
  async function restore(i){
    if(i<0||i>=hist.length)return;
    hi=i;const s=hist[i],reload=s.bytes!==bytes;
    bytes=s.bytes;annots=clone(s.annots);resetSel();dirty=true;btns();
    if(reload)await load(bytes,{page:cur});else drawAll();
  }
  async function apply(nb,page,keep){bytes=nb;annots=keep||[];resetSel();push();await load(nb,{page});}
  const run=async fn=>{if(locked)return;locked=true;try{await fn();}catch(e){console.error(e);toast('Gagal: '+(e&&e.message||e));}finally{locked=false;}};

  /* ---------- memuat & merender ---------- */
  async function load(buf,o={}){
    await initPdfjs();
    const task=pdfjsLib.getDocument({data:buf.slice()});
    task.onPassword=(upd,reason)=>{const p=prompt(reason===2?'Kata sandi salah. Coba lagi:':'PDF ini dilindungi kata sandi:');if(p===null)task.destroy();else upd(p);};
    const doc=await task.promise;
    if(io){io.disconnect();io=null;}
    pages.forEach(p=>{p.tok++;try{p.task&&p.task.cancel();}catch{}});
    if(pdfDoc)try{pdfDoc.destroy();}catch{}
    pdfDoc=doc;textCache.clear();
    const pg=await Promise.all(Array.from({length:doc.numPages},(_,i)=>doc.getPage(i+1)));
    pages=pg.map(p=>{const vp=p.getViewport({scale:1});return{w:vp.width,h:vp.height,rot:p.rotate||0,vp,tok:0,rendered:0,pr:null,divs:null,strs:null,blocks:null};});
    pgs.innerHTML=pages.map((_,i)=>`<div class="pg" data-i="${i}"><canvas></canvas><div class="hll"></div><div class="tl textLayer"></div><div class="annl"></div></div>`).join('');
    if(autoFit)zoom=Math.max(10,Math.min(100,fitW()));
    S=zoom/100*4/3;layout();updZoomUI();
    io=new IntersectionObserver(es=>es.forEach(en=>{const i=+en.target.dataset.i;if(en.isIntersecting)renderPage(i);else freePage(i);}),{root:desk,rootMargin:'800px 0px'});
    [...pgs.children].forEach(el=>io.observe(el));
    drawAll();
    const t=Math.max(0,Math.min(pages.length-1,o.page??cur));
    cur=t;goto(t,false);updPg();
    if(fq)doFind(fq,true);
  }
  function fitW(){const mw=Math.max(...pages.map(p=>p.w),1);return Math.floor((desk.clientWidth-40)/mw*75);}
  function layout(){
    [...pgs.children].forEach((el,i)=>{const p=pages[i];el.style.width=p.w*S+'px';el.style.height=p.h*S+'px';el.style.setProperty('--s',S);
      const al=el.querySelector('.annl');al.style.transform=`scale(${S})`;al.style.width=p.w+'px';al.style.height=p.h+'px';});
  }
  function renderPage(i){
    const p=pages[i];if(!p||!pdfDoc)return Promise.resolve();
    if(p.rendered===zoom&&p.pr)return p.pr;
    p.rendered=zoom;const tok=++p.tok,s=S,el=pgs.children[i];
    p.pr=(async()=>{
      const page=await pdfDoc.getPage(i+1);if(tok!==p.tok)return;
      const vp=page.getViewport({scale:s});
      let k=Math.min(window.devicePixelRatio||1,2);if(vp.width*vp.height*k*k>16e6)k=Math.sqrt(16e6/(vp.width*vp.height));
      const c=document.createElement('canvas');c.width=Math.max(1,Math.floor(vp.width*k));c.height=Math.max(1,Math.floor(vp.height*k));
      const rt=page.render({canvasContext:c.getContext('2d'),viewport:vp,transform:k!==1?[k,0,0,k,0,0]:null,background:'#fff'});
      p.task=rt;
      try{await rt.promise;}catch{return;}
      if(tok!==p.tok)return;
      const old=el.querySelector('canvas');old.replaceWith(c);
      const tl=el.querySelector('.tl');tl.textContent='';tl.style.setProperty('--scale-factor',s);
      const tc=await getText(i);if(tok!==p.tok)return;
      const divs=[],strs=[];
      try{const tt=pdfjsLib.renderTextLayer({textContentSource:tc,container:tl,viewport:vp,textDivs:divs,textContentItemsStr:strs});await tt.promise;}catch{}
      if(tok!==p.tok)return;
      p.divs=divs;p.strs=strs;applyHits(i);
      if(!p.blocks){try{p.blocks=(p.rot%360===0)?U.pageBlocks(page,tc,p.vp):[];}catch(err){console.warn('blok teks:',err);p.blocks=[];}}
    })();
    return p.pr;
  }
  function freePage(i){
    const p=pages[i];if(!p||!p.rendered)return;
    p.tok++;try{p.task&&p.task.cancel();}catch{}
    p.rendered=0;p.pr=null;p.divs=p.strs=null;
    const el=pgs.children[i];if(!el)return;
    const c=el.querySelector('canvas');c.width=c.height=1;el.querySelector('.tl').textContent='';
  }
  const getText=i=>{if(!textCache.has(i))textCache.set(i,pdfDoc.getPage(i+1).then(p=>p.getTextContent()));return textCache.get(i);};
  function renderVisible(){
    const dr=desk.getBoundingClientRect();
    [...pgs.children].forEach((el,i)=>{const r=el.getBoundingClientRect();if(r.bottom>dr.top-800&&r.top<dr.bottom+800)renderPage(i);});
  }

  /* ---------- zoom, navigasi ---------- */
  function setZoom(v,keep=true){
    if(!pages.length)return;
    v=Math.max(10,Math.min(500,Math.round(v)));
    const fy=keep&&desk.scrollHeight?(desk.scrollTop+desk.clientHeight/2)/desk.scrollHeight:0,fx=keep&&desk.scrollWidth?(desk.scrollLeft+desk.clientWidth/2)/desk.scrollWidth:0;
    zoom=v;S=v/100*4/3;layout();
    pages.forEach((p,i)=>{p.tok++;try{p.task&&p.task.cancel();}catch{}p.rendered=0;p.pr=null;pgs.children[i].querySelector('.tl').textContent='';});
    if(keep){desk.scrollTop=fy*desk.scrollHeight-desk.clientHeight/2;desk.scrollLeft=fx*desk.scrollWidth-desk.clientWidth/2;}
    updZoomUI();renderVisible();
  }
  function updZoomUI(){$('.zs').value=Math.max(10,Math.min(400,zoom));$('.zv').textContent=zoom+'%';}
  const zstep=d=>{autoFit=false;const k=d>0?ZL.find(z=>z>zoom):[...ZL].reverse().find(z=>z<zoom);if(k)setZoom(k);};
  function goto(i,smooth=true){const el=pgs.children[i];if(el)desk.scrollTo({top:el.offsetTop-10,behavior:smooth?'smooth':'auto'});}
  function updPg(){$('.pi').value=cur+1;$('.pi').max=pages.length;$('.of').textContent='OF '+pages.length;}
  let raf=0;
  desk.addEventListener('scroll',()=>{if(raf)return;raf=requestAnimationFrame(()=>{raf=0;const mid=desk.scrollTop+desk.clientHeight*0.35;let c=0;
    for(let i=0;i<pgs.children.length;i++){if(pgs.children[i].offsetTop-10<=mid)c=i;else break;}
    if(c!==cur){cur=c;updPg();}renderVisible();});});

  /* ---------- anotasi ---------- */
  const annEl=id=>pgs.querySelector(`.an[data-id="${id}"]`);
  const annEls=id=>[...pgs.querySelectorAll(`[data-id="${id}"]`)];
  let hvEl=null;
  const HD=['nw','n','ne','e','se','s','sw','w'].map(h=>`<span class="hd" data-h="${h}"></span>`).join(''),MV=['t','b','l','r'].map(m=>`<span class="mv" data-m="${m}"></span>`).join('');
  const runsHtml=rs=>rs.map(r=>{let h=esc(r.t).replace(/\n/g,'<br>');if(r.b)h=`<b>${h}</b>`;if(r.i)h=`<i>${h}</i>`;return h;}).join('');
  function readRuns(root){
    const out=[];
    const add=(t,i,b)=>{if(!t)return;const l=out[out.length-1];if(l&&l.i===i&&l.b===b)l.t+=t;else out.push({t,i,b});};
    const walk=(n,i,b)=>{
      for(const c of n.childNodes){
        if(c.nodeType===3)add(c.nodeValue.replace(/\u00A0/g,' ').replace(/[\u200B\uFEFF]/g,''),i,b);
        else if(c.nodeType===1){
          const tg=c.tagName;
          if(tg==='BR'){if(!(n!==root&&n.childNodes.length===1&&(n.tagName==='DIV'||n.tagName==='P')))add('\n',i,b);continue;}
          const st=c.style||{};
          const ci=i||tg==='I'||tg==='EM'||st.fontStyle==='italic',cb=b||tg==='B'||tg==='STRONG'||st.fontWeight==='bold'||+st.fontWeight>=600;
          if((tg==='DIV'||tg==='P')&&out.length&&!/\n$/.test(out[out.length-1].t))add('\n',ci,cb);
          walk(c,ci,cb);
        }
      }
    };
    walk(root,false,false);
    if(out.length){const l=out[out.length-1];l.t=l.t.replace(/\n+$/,'');if(!l.t)out.pop();}
    return out;
  }
  function drawAll(){pages.forEach((_,i)=>drawPageAnn(i));}
  function drawPageAnn(i){
    const el=pgs.children[i];if(!el)return;
    const al=el.querySelector('.annl'),hl=el.querySelector('.hll');al.textContent='';hl.textContent='';
    const list=annots.filter(a=>a.page===i);
    list.forEach(a=>{if(a.type==='tb')al.appendChild(mkCover(a));});
    list.forEach(a=>placeAnn(a,al,hl));
  }
  function placeAnn(a,al,hl){
    if(a.type==='hl')a.rects.forEach(r=>hl.appendChild(mkHl(a,r)));else al.appendChild(mkAnn(a));
  }
  function mkCover(a){const d=document.createElement('div'),c=a.cover;d.className='cov';d.dataset.for=a.id;d.style.cssText=`left:${c.x}px;top:${c.y}px;width:${c.w}px;height:${c.h}px`;return d;}
  function mkHl(a,r){
    const d=document.createElement('i');d.dataset.id=a.id;d.className=(a.kind&&a.kind!=='hl'?a.kind:'')+(a.cm?' cm':'')+(sel.includes(a.id)?' sel':'');
    d.style.cssText=`left:calc(var(--s)*${r.x}px);top:calc(var(--s)*${r.y}px);width:calc(var(--s)*${r.w}px);height:calc(var(--s)*${r.h}px)`;
    return d;
  }
  function mkAnn(a){
    const d=document.createElement('div');d.dataset.id=a.id;
    if(a.type==='tx'){d.className=txCls(a);d.style.cssText=txCss(a);d.textContent=a.text;}
    else if(a.type==='sh'){d.className='an sh';d.style.cssText=`left:${a.x}px;top:${a.y}px;width:${a.w}px;height:${a.h}px`;d.innerHTML=U.shapeSvg(a)+'<span class="rs"></span>';}
    else if(a.type==='dr'){d.className='an dr';d.style.cssText=`left:${a.x}px;top:${a.y}px;width:${a.w}px;height:${a.h}px${a.mul?';mix-blend-mode:multiply':''}`;d.innerHTML=`<svg viewBox="0 0 ${a.w} ${a.h}"><polyline points="${a.pts.map(q=>q.join(',')).join(' ')}" fill="none" stroke="${a.color}" stroke-width="${a.lw}" stroke-linecap="round" stroke-linejoin="round"/></svg>`;}
    else if(a.type==='lk'){d.className='an lk';d.style.cssText=`left:${a.x}px;top:${a.y}px;width:${a.w}px;height:${a.h}px`;d.title=a.url?a.url:'Ke halaman '+(a.dest+1);}
    else if(a.type==='im'){d.className='an im';d.style.cssText=`left:${a.x}px;top:${a.y}px;width:${a.w}px;height:${a.h}px`;d.innerHTML=`<img src="${a.src}" alt=""><span class="rs"></span>`;}
    else if(a.type==='tb'){
      d.className='an tb '+(a.fam||'serif');
      d.style.cssText=`left:${a.x}px;top:${a.y}px;width:${a.w}px;min-height:${a.h}px;font-size:${a.size}px;line-height:${a.lh};text-align:${a.align};color:${a.color||'#000'}`;
      d.innerHTML=`<div class="tbt" style="text-indent:${a.indent||0}px">${runsHtml(a.runs)}</div>${MV}${HD}`;
    }
    if(a.op!=null&&a.op<100&&a.type!=='tx')d.style.opacity=a.op/100;
    if(a.cm)d.classList.add('cm');
    if(sel.includes(a.id))d.classList.add('sel');
    return d;
  }
  function dropUntouched(id){
    const a=annots.find(x=>x.id===id);if(!a||a.type!=='tb'||a.touched)return;
    annots=annots.filter(x=>x!==a);annEls(id).forEach(e=>e.remove());
    pgs.querySelectorAll(`[data-for="${id}"]`).forEach(e=>e.remove());
  }
  function select(id,add){
    const prev=sel.slice();
    if(id==null) sel=[];
    else if(add){
      const a=annots.find(x=>x.id===id);
      if(a&&a.type!=='tb'&&a.type!=='hl'){
        if(sel.includes(id)) sel=sel.filter(x=>x!==id);
        else{const f=selA()[0];sel=(f&&f.type!=='tb'&&f.type!=='hl'&&f.page===a.page)?[...sel,id]:[id];}
      }
    }else sel=[id];
    selId=sel.length?sel[sel.length-1]:null;
    prev.filter(x=>!sel.includes(x)).forEach(dropUntouched);
    pgs.querySelectorAll('.sel').forEach(e=>e.classList.remove('sel'));
    sel.forEach(x=>annEls(x).forEach(e=>e.classList.add('sel')));
    pgs.classList.toggle('multi',sel.length>1);
    if(sel.length&&!pd.contains(document.activeElement))pd.focus({preventScroll:true}); /* agar Delete/Esc bekerja tanpa klik lagi */
    updCtx();
  }
  function setTool(t){
    tool=t;desk.className='desk t-'+t+(hidecm?' hidecm':'');
    $$('[data-t]').forEach(b=>b.classList.toggle('on',b.dataset.t===t));
    if(t!=='edit'&&t!=='select'&&!TXT[t])select(null);
    if(t!=='edit')hideHv();
    if(!SELTOOL[t]&&t!=='select')getSelection().removeAllRanges();
  }
  function startEdit(el){
    if(curEditDone)curEditDone();
    el.contentEditable='true';el.focus();
    const r=document.createRange();r.selectNodeContents(el);const sg=getSelection();sg.removeAllRanges();sg.addRange(r);
    const id=el.dataset.id;let fin=false;
    const done=()=>{
      if(fin)return;fin=true;
      el.removeEventListener('blur',onblur);pd.removeEventListener('pointerdown',outside,true);
      el.contentEditable='false';curEditDone=null;editId=null;
      const a=annots.find(x=>x.id===id);
      if(a){
        const t=el.innerText.replace(/\n+$/,'');
        if(!t.trim()){annots=annots.filter(x=>x.id!==id);el.remove();sel=sel.filter(x=>x!==id);selId=sel.length?sel[sel.length-1]:null;if(a.text)push();}
        else{if(t!==a.text){a.text=t;push();}if(TXT[tool]){setTool('select');select(id);}}
      }
      if(!pd.contains(document.activeElement))pd.focus({preventScroll:true});
      updCtx();
    };
    /* fokus pindah ke ribbon (tab Format) tidak mengakhiri pengetikan; klik di luar / Finish yang mengakhiri */
    const onblur=e=>{const rt=e&&e.relatedTarget;if(rt&&rt.closest&&rt.closest('.rib,.tabs'))return;done();};
    const outside=e=>{const t=e.target;if(el.contains(t)||(t.closest&&t.closest('.rib,.tabs,.pmenu,.ov')))return;done();};
    el.addEventListener('blur',onblur);pd.addEventListener('pointerdown',outside,true);
    curEditDone=done;editId=id;updCtx();
    el.onkeydown=ev=>{if(ev.key==='Escape'){ev.stopPropagation();done();}};
  }
  function addAnn(a){
    a.id='a'+(++idc);if(a.type!=='tb'){if(a.au==null&&author)a.au=author;if(!a.sj)a.sj=U.kindLabel(a);}annots.push(a);
    const pe=pgs.children[a.page],al=pe.querySelector('.annl'),hl=pe.querySelector('.hll');
    if(a.type==='tb')al.prepend(mkCover(a));
    placeAnn(a,al,hl);
    return a.type==='hl'?null:annEl(a.id);
  }
  /* ---------- blok teks asli (alat Edit ala Nitro) ---------- */
  function hitBlock(bl,x,y){
    let best=null,ba=1e18;
    for(const b of bl){const c=b.cover,m=3;
      if(x>=c.x-m&&x<=c.x+c.w+m&&y>=c.y-m&&y<=c.y+c.h+m){const ar=c.w*c.h;if(ar<ba){ba=ar;best=b;}}}
    return best;
  }
  function hideHv(){if(hvEl)hvEl.hidden=true;desk.style.cursor='';}
  function showHv(i,b){
    const al=pgs.children[i].querySelector('.annl');
    if(!hvEl){hvEl=document.createElement('div');hvEl.className='hv';}
    if(hvEl.parentNode!==al)al.appendChild(hvEl);
    const c=b.cover;hvEl.style.cssText=`left:${c.x}px;top:${c.y}px;width:${c.w}px;height:${c.h}px`;
    hvEl.classList.toggle('bad',!b.ok);hvEl.hidden=false;desk.style.cursor=b.ok?'text':'not-allowed';
  }
  function openBlockAt(e,pgEl){
    const i=+pgEl.dataset.i,p=pages[i];if(!p||!p.blocks)return false;
    const[x,y]=pt(e,pgEl),b=hitBlock(p.blocks,x,y);if(!b)return false;
    let a=annots.find(t=>t.type==='tb'&&t.page===i&&t.bid===b.id);
    if(!a){
      if(!b.ok){toast('Paragraf ini berisi rumus/simbol khusus yang tidak bisa diedit tanpa merusak tampilannya.');return true;}
      const el=addAnn({type:'tb',page:i,bid:b.id,x:b.x,y:b.y,w:b.w,h:b.h,size:b.size,fam:b.fam,lh:b.lh,indent:b.indent,align:b.align,runs:clone(b.runs),cover:{...b.cover},color:'#000000',touched:false});
      a=annots.find(t=>t.id===el.dataset.id);
    }
    select(a.id);startTb(annEl(a.id),e.clientX,e.clientY);hideHv();
    return true;
  }
  function startTb(el,cx,cy){
    const a=annots.find(x=>x.id===el.dataset.id),t=el.querySelector('.tbt');
    if(!a||!t||t.isContentEditable)return;
    t.contentEditable='true';t.focus({preventScroll:true});
    if(cx!=null){
      let r=null;
      if(document.caretRangeFromPoint)r=document.caretRangeFromPoint(cx,cy);
      else if(document.caretPositionFromPoint){const p=document.caretPositionFromPoint(cx,cy);if(p){r=document.createRange();r.setStart(p.offsetNode,p.offset);r.collapse(true);}}
      if(r&&t.contains(r.startContainer)){const s=getSelection();s.removeAllRanges();s.addRange(r);}
    }
    let ch=false;
    const inp=()=>{ch=true;a.touched=true;};
    const ps=e=>{e.preventDefault();document.execCommand('insertText',false,(e.clipboardData||window.clipboardData).getData('text/plain'));};
    const kd=e=>{if(e.key==='Escape'){e.stopPropagation();t.blur();pd.focus();}};
    const done=()=>{
      t.removeEventListener('input',inp);t.removeEventListener('blur',done);t.removeEventListener('keydown',kd);t.removeEventListener('paste',ps);
      t.contentEditable='false';
      if(!annots.includes(a))return;
      if(ch){a.runs=readRuns(t);t.innerHTML=runsHtml(a.runs);push();}
    };
    t.addEventListener('input',inp);t.addEventListener('blur',done);t.addEventListener('keydown',kd);t.addEventListener('paste',ps);
  }
  function tbPointer(e,an,a){
    const hd=e.target.closest('.hd'),mv=e.target.closest('.mv'),t=an.querySelector('.tbt');
    if(!hd&&!mv){select(a.id);if(!t.isContentEditable)startTb(an,e.clientX,e.clientY);return;}
    e.preventDefault();select(a.id);if(document.activeElement===t)t.blur();
    const dir=hd?hd.dataset.h:'',sx=e.clientX,sy=e.clientY,o={x:a.x,y:a.y,w:a.w,h:Math.max(a.h,an.offsetHeight)};let moved=false;
    an.setPointerCapture&&an.setPointerCapture(e.pointerId);
    const mf=m=>{
      const dx=(m.clientX-sx)/S,dy=(m.clientY-sy)/S;moved=true;
      if(mv){a.x=o.x+dx;a.y=o.y+dy;}
      else{
        let w=o.w,h=o.h,x=o.x,y=o.y;
        if(dir.includes('e'))w=o.w+dx;
        if(dir.includes('w')){w=o.w-dx;x=o.x+dx;}
        if(w<24){if(dir.includes('w'))x=o.x+o.w-24;w=24;}
        an.style.width=w+'px';
        const mh=t.offsetHeight;
        if(dir.includes('s'))h=o.h+dy;
        if(dir.includes('n')){h=o.h-dy;y=o.y+dy;}
        if(h<mh){if(dir.includes('n'))y=o.y+o.h-mh;h=mh;}
        a.x=x;a.y=y;a.w=w;a.h=h;
      }
      an.style.left=a.x+'px';an.style.top=a.y+'px';an.style.width=a.w+'px';an.style.minHeight=a.h+'px';
    };
    const up=()=>{an.removeEventListener('pointermove',mf);an.removeEventListener('pointerup',up);an.removeEventListener('pointercancel',up);if(moved){a.touched=true;push();}};
    an.addEventListener('pointermove',mf);an.addEventListener('pointerup',up);an.addEventListener('pointercancel',up);
  }
  function pt(e,el){const r=el.getBoundingClientRect();return[(e.clientX-r.left)/S,(e.clientY-r.top)/S];}

  desk.addEventListener('click',async e=>{
    const pgEl=e.target.closest('.pg');
    if(tool==='zoom'){autoFit=false;if(e.shiftKey||e.altKey)zstep(-1);else zstep(1);return;}
    if(!pgEl)return;
    const i=+pgEl.dataset.i;
    if(TXT[tool]){
      const tx=e.target.closest('.an.tx');if(tx){if(!tx.isContentEditable)startEdit(tx);return;}
      if(e.target.closest('.an'))return;
      const o=TXT[tool],[x,y]=pt(e,pgEl),size=o.size,pad=(o.box||o.bg)?4:0;
      const el=addAnn({type:'tx',page:i,x:Math.max(0,x),y:Math.max(0,y-size*0.6-pad),size,color:'#000000',text:'',...o.p,...(defs[tool]||{})});
      select(null);startEdit(el);
    }else if(tool==='stamp'){
      placeImg(i,pt(e,pgEl),stampImg(stampSel),140,true);
    }else if(tool==='sign'){
      if(!sign){const s=await signDlg();if(!s)return;sign=s;store.set('sign',s);}
      placeSign(i,pt(e,pgEl));
    }else if(tool==='edit'){
      if(e.target.closest('.an'))return;
      openBlockAt(e,pgEl);
    }else if(tool==='select'){
      if(e.target.closest('.an'))return;
      const sg=getSelection();if(sg&&!sg.isCollapsed)return;
      const[x,y]=pt(e,pgEl),h=[...annots].reverse().find(z=>z.type==='hl'&&z.page===i&&z.rects.some(r=>x>=r.x&&x<=r.x+r.w&&y>=r.y&&y<=r.y+r.h));
      if(h)select(h.id);
    }else if(SELTOOL[tool]){
      const s=getSelection();if(s&&!s.isCollapsed)return;
      const kk=SELTOOL[tool]==='rp'?'st':SELTOOL[tool];
      const[x,y]=pt(e,pgEl),h=annots.find(z=>z.type==='hl'&&(z.kind||'hl')===kk&&z.page===i&&z.rects.some(r=>x>=r.x&&x<=r.x+r.w&&y>=r.y&&y<=r.y+r.h));
      if(h){annots=annots.filter(z=>z!==h);annEls(h.id).forEach(el=>el.remove());push();toast(kk==='rd'?'Tanda redaksi dihapus':'Markup dihapus');}
    }
  });
  function placeImg(i,[x,y],src,w0,cm){
    const im=new Image();
    im.onload=()=>{const w=w0,h=w*im.height/im.width;
      const el=addAnn({type:'im',page:i,x:Math.max(0,x-w/2),y:Math.max(0,y-h/2),w,h,src,...(cm?{cm:true}:{})});
      push();setTool('select');selectQuiet(el.dataset.id);};
    im.src=src;
  }
  const placeSign=(i,p)=>placeImg(i,p,sign,150);
  /* geser / ubah ukuran / hapus anotasi (alat Edit) */
  desk.addEventListener('pointerdown',e=>{
    if(tool==='hand'){
      const sx=e.clientX,sy=e.clientY,l=desk.scrollLeft,t=desk.scrollTop;desk.classList.add('drag');desk.setPointerCapture(e.pointerId);
      const mv=m=>{desk.scrollLeft=l-(m.clientX-sx);desk.scrollTop=t-(m.clientY-sy);};
      const up=()=>{desk.classList.remove('drag');desk.removeEventListener('pointermove',mv);desk.removeEventListener('pointerup',up);desk.removeEventListener('pointercancel',up);};
      desk.addEventListener('pointermove',mv);desk.addEventListener('pointerup',up);desk.addEventListener('pointercancel',up);return;
    }
    if(DRAWT[tool]){drawStart(e);return;}
    if(tool!=='edit'&&tool!=='select')return;
    const an=e.target.closest('.an'),multi=e.shiftKey||e.ctrlKey||e.metaKey;
    if(!an){if(!multi)select(null);return;}
    const a=annots.find(x=>x.id===an.dataset.id);if(!a)return;
    if(a.type==='tb'){if(tool==='edit')tbPointer(e,an,a);else select(a.id);return;}
    pd.focus&&pd.focus({preventScroll:true}); /* supaya tombol Delete sampai ke aplikasi */
    if(multi){e.preventDefault();select(a.id,true);return;}
    if(!(sel.includes(a.id)&&sel.length>1))select(a.id);
    {const ac=$('.tab.on'),pt_=$('.tab[data-tab=props]');if(ac&&!ac.classList.contains('ctx')&&pt_&&!pt_.hidden&&!matchMedia('(max-width:700px)').matches)showTab('props');}
    if(an.isContentEditable)return;
    e.preventDefault();
    const grp=selA().filter(x=>x.type!=='tb'&&x.type!=='hl'),rs=e.target.classList.contains('rs')&&grp.length===1,sx=e.clientX,sy=e.clientY,orig=grp.map(g=>({a:g,x:g.x,y:g.y})),ow=a.w,oh=a.h;let moved=false;
    an.setPointerCapture&&an.setPointerCapture(e.pointerId);
    const mv=m=>{const dx=(m.clientX-sx)/S,dy=(m.clientY-sy)/S;moved=true;
      if(rs&&a.type==='im'){a.w=Math.max(20,ow+dx);a.h=a.w*oh/ow;an.style.width=a.w+'px';an.style.height=a.h+'px';}
      else if(rs&&a.type==='sh'){a.w=Math.max(4,ow+dx);a.h=Math.max(4,oh+dy);an.style.width=a.w+'px';an.style.height=a.h+'px';const sv=an.querySelector('svg');sv.insertAdjacentHTML('beforebegin',U.shapeSvg(a));sv.remove();}
      else{for(const o of orig){o.a.x=o.x+dx;o.a.y=o.y+dy;const el=annEl(o.a.id);if(el){el.style.left=o.a.x+'px';el.style.top=o.a.y+'px';}}}};
    const up=()=>{an.removeEventListener('pointermove',mv);an.removeEventListener('pointerup',up);an.removeEventListener('pointercancel',up);if(moved)push();};
    an.addEventListener('pointermove',mv);an.addEventListener('pointerup',up);an.addEventListener('pointercancel',up);
  });
  desk.addEventListener('pointermove',e=>{
    if(tool!=='edit'||e.buttons){if(tool!=='edit')return;hideHv();return;}
    const pgEl=e.target.closest&&e.target.closest('.pg');
    if(!pgEl||e.target.closest('.an')){hideHv();return;}
    const i=+pgEl.dataset.i,p=pages[i];if(!p||!p.blocks){hideHv();return;}
    const[x,y]=pt(e,pgEl),b=hitBlock(p.blocks,x,y);
    if(!b){hideHv();return;}
    const a=annots.find(t=>t.type==='tb'&&t.page===i&&t.bid===b.id&&t.id===selId);
    if(a){hideHv();return;}
    showHv(i,b);
  });
  desk.addEventListener('pointerleave',hideHv);
  desk.addEventListener('dblclick',e=>{const tx=e.target.closest('.an.tx');if((tool==='edit'||tool==='select')&&tx)startEdit(tx);});
  /* sorot teks */
  desk.addEventListener('pointerup',()=>{if(SELTOOL[tool])setTimeout(()=>highlightSel(SELTOOL[tool]),0);});
  /* gabungkan kotak per-span menjadi satu pita per baris (tanpa celah & tanpa tumpang tindih) */
  function mergeLine(rs){
    rs.sort((a,b)=>a.y-b.y||a.x-b.x);const out=[];
    for(const r of rs){
      let m=null;
      for(const o of out){
        const ov=Math.min(o.y+o.h,r.y+r.h)-Math.max(o.y,r.y),gap=Math.max(o.h,r.h)*2.5;
        if(ov>0.5*Math.min(o.h,r.h)&&r.x<=o.x+o.w+gap&&r.x+r.w>=o.x-gap){m=o;break;}
      }
      if(m){const x2=Math.max(m.x+m.w,r.x+r.w),y2=Math.max(m.y+m.h,r.y+r.h);m.x=Math.min(m.x,r.x);m.y=Math.min(m.y,r.y);m.w=x2-m.x;m.h=y2-m.y;}
      else out.push({...r});
    }
    return out;
  }
  async function highlightSel(k){
    const s=getSelection();if(!s.rangeCount||s.isCollapsed)return;
    const rg=s.getRangeAt(0),by=new Map();
    [...pgs.children].forEach((pe,i)=>{
      const b=pe.getBoundingClientRect(),rs=[];
      pe.querySelectorAll('.tl span').forEach(sp=>{
        if(sp.classList.contains('markedContent')||!sp.firstChild||!rg.intersectsNode(sp))return;
        const sub=document.createRange();sub.selectNodeContents(sp);
        if(rg.compareBoundaryPoints(Range.START_TO_START,sub)>0)sub.setStart(rg.startContainer,rg.startOffset);
        if(rg.compareBoundaryPoints(Range.END_TO_END,sub)<0)sub.setEnd(rg.endContainer,rg.endOffset);
        if(sub.collapsed)return;
        for(const r of sub.getClientRects())if(r.width>0.5&&r.height>0.5)rs.push({x:(r.left-b.left)/S,y:(r.top-b.top)/S,w:r.width/S,h:r.height/S});
      });
      if(rs.length)by.set(i,mergeLine(rs));
    });
    if(!by.size)return;
    s.removeAllRanges();
    const kind=k==='rp'?'st':k;let last=null;
    by.forEach((rs,i)=>{addAnn({type:'hl',kind,page:i,rects:rs,...(kind==='rd'?{}:{cm:true})});last={i,r:rs[rs.length-1]};});
    if(k==='rp'&&last){
      const b=await dialog('Ganti teks','<label>Teks pengganti:<input type="text" class="rt"></label>','Tambah');
      const t=b&&b.querySelector('.rt').value.trim();
      if(t)addAnn({type:'tx',page:last.i,x:last.r.x+last.r.w+3,y:last.r.y,size:Math.max(8,Math.round(last.r.h*0.8)),color:'#d90000',text:t,cm:true});
    }
    push();
  }
  /* ====== seleksi objek & tab kontekstual (Properties / Alignment / Format) ====== */
  const selA=()=>sel.map(id=>annots.find(a=>a.id===id)).filter(Boolean);
  function resetSel(){sel=[];selId=null;updCtx();}
  /* objek yang baru dibuat: tab kontekstual muncul, tetapi tab aktif tidak berpindah (agar bisa menggambar lagi) */
  function selectQuiet(id){quiet=true;try{select(id);}finally{quiet=false;}}
  const txCls=a=>'an tx'+(a.box||a.bg?' bx':'')+(a.bg?' nt':'')+(a.co?' co':'');
  function txCss(a){
    const box=!!(a.box||a.bg);
    return `left:${a.x}px;top:${a.y}px;font-size:${a.size}px;color:${a.color||'#000'};font-family:${U.FAM[a.font||'helv']};font-weight:${a.bold?700:400};font-style:${a.italic?'italic':'normal'};text-align:${({l:'left',c:'center',r:'right'})[a.align||'l']};letter-spacing:${a.cs||0}px;word-spacing:${a.ws||0}px;line-height:${a.ls||1.2}`
      +(box?`;--tbc:${a.box||'#444444'};--tbw:${a.bw||1}px;--tbs:${a.dash==='dashed'?'dashed':a.dash==='dotted'?'dotted':'solid'}`:'')
      +(a.bg?`;--tbg:${a.bg}`:'')+(a.co?`;--tll:${a.ll||40}px;--tlo:${a.lo||0}px`:'')+(a.op!=null&&a.op<100?`;opacity:${a.op/100}`:'');
  }
  function updTx(el,a){const c=txCls(a).split(' ');['bx','nt','co'].forEach(k=>el.classList.toggle(k,c.includes(k)));el.style.cssText=txCss(a);}
  function refresh(a){
    const el=annEl(a.id);if(!el)return;
    if(a.type==='tx'){updTx(el,a);return;}
    el.replaceWith(mkAnn(a));
  }
  function showTab(n){$$('.tab[data-tab]').forEach(x=>x.classList.toggle('on',x.dataset.tab===n));$$('[data-pane]').forEach(p=>p.hidden=p.dataset.pane!==n);}
  const HEX=v=>/^#[0-9a-f]{6}$/i.test(v||'')?v.toLowerCase():null;
  const clampN=(v,a,b)=>Math.max(a,Math.min(b,+v));
  function ppTargets(){const ed=editId&&annots.find(a=>a.id===editId);return ed?[ed]:selA().filter(a=>a.type!=='tb');}
  const lineSh=a=>a.type==='sh'&&(a.shape==='line'||a.shape==='arrow');
  const boxTx=a=>a.type==='tx'&&(a.box||a.bg);
  const PPSET={
    author:(a,v)=>{a.au=String(v);return true;},
    subject:(a,v)=>{a.sj=String(v);return true;},
    opacity:(a,v)=>{if(!['tx','sh','dr','im'].includes(a.type))return false;a.op=clampN(v,5,100);return true;},
    border:(a,v)=>{if(a.type==='tx')a.box=v;else if(a.type==='sh'||a.type==='dr')a.color=v;else return false;return true;},
    fill:(a,v)=>{
      if(a.type==='tx'){if(v)a.bg=v;else delete a.bg;return true;}
      if(a.type==='sh'&&(a.shape==='rect'||a.shape==='ell')){if(v)a.fill=v;else delete a.fill;return true;}
      return false;
    },
    thick:(a,v)=>{if(boxTx(a)){a.bw=clampN(v,.5,20);return true;}if(a.type==='sh'||a.type==='dr'){a.lw=clampN(v,.5,20);return true;}return false;},
    style:(a,v)=>{if(boxTx(a)||a.type==='sh'){if(v==='solid')delete a.dash;else a.dash=v;return true;}return false;},
    start:(a,v)=>{if(lineSh(a)){a.s=v;return true;}return false;},
    end:(a,v)=>{if(lineSh(a)){a.e=v;return true;}return false;},
    ll:(a,v)=>{if(a.type==='tx'&&a.co){a.ll=clampN(v,5,300);return true;}return false;},
    lo:(a,v)=>{if(a.type==='tx'&&a.co){a.lo=clampN(v,0,1000);return true;}return false;},
    font:(a,v)=>{if(a.type!=='tx'||!U.FAM[v])return false;a.font=v;return true;},
    size:(a,v)=>{if(a.type!=='tx')return false;a.size=clampN(v,4,200);return true;},
    tcolor:(a,v)=>{if(a.type!=='tx')return false;a.color=v;return true;},
    fstyle:(a,v)=>{if(a.type!=='tx')return false;a.bold=v==='b'||v==='bi';a.italic=v==='i'||v==='bi';return true;},
    align:(a,v)=>{if(a.type!=='tx')return false;a.align=v;return true;},
    cs:(a,v)=>{if(a.type!=='tx')return false;a.cs=clampN(v,-5,50);return true;},
    ws:(a,v)=>{if(a.type!=='tx')return false;a.ws=clampN(v,-10,100);return true;},
    ls:(a,v)=>{if(a.type!=='tx')return false;a.ls=clampN(v,.8,5);return true;}
  };
  function ppApply(k,v){let n=0;const f=PPSET[k];if(!f)return 0;for(const a of ppTargets())if(f(a,v)){n++;refresh(a);}syncProps();return n;}
  function ppEvent(e,commit){
    const c=e.target.closest&&e.target.closest('[data-pp]');if(!c||c.disabled)return;
    const k=c.dataset.pp;let v=c.value;
    if(c.type==='number'||k==='ls'){v=parseFloat(v);if(!isFinite(v))return;}
    const n=ppApply(k,v);
    if(k==='author'&&commit){author=String(v).trim();store.set('author',author);}
    if(commit&&n)push();
  }
  pd.addEventListener('input',e=>ppEvent(e,false));
  pd.addEventListener('change',e=>ppEvent(e,true));
  /* klik tombol ribbon saat mengetik tidak boleh merebut fokus dari kotak teks */
  pd.addEventListener('mousedown',e=>{if(editId&&e.target.closest&&e.target.closest('.rib .b,.pmenu button')&&!e.target.closest('select,input'))e.preventDefault();});
  function updCtx(){
    const ed=editId?annots.find(a=>a.id===editId):null;
    const objs=ed?[]:selA().filter(a=>a.type!=='tb');
    const showP=objs.length>0,showA=objs.length>0&&objs.every(a=>a.type!=='hl'),showF=!!ed||objs.some(a=>a.type==='tx');
    const vis={props:showP,align:showA,format:showF};
    for(const k in vis){const t=$('.tab[data-tab="'+k+'"]');if(t)t.hidden=!vis[k];}
    const any=showP||showA||showF,act=$('.tab.on'),mob=matchMedia('(max-width:700px)').matches;
    if(!mob){
      if(any){
        if(ed&&!edShown)showTab('format');
        else if((!ctxOn||(edShown&&!ed))&&!quiet)showTab(showP?'props':'format');
        else if(act&&act.classList.contains('ctx')&&act.hidden)showTab(showP?'props':'format');
      }else if(act&&act.classList.contains('ctx'))showTab(lastMain);
    }
    ctxOn=any;edShown=!!ed;
    const n=objs.filter(a=>a.type!=='hl').length;
    $$('[data-n]').forEach(b=>{b.disabled=n<+b.dataset.n;});
    syncProps();
  }
  function syncProps(){
    const T=ppTargets(),p=T[0];if(!p)return;
    const any=f=>T.some(f);
    const put=(k,v)=>{const c=$('[data-pp="'+k+'"]');if(!c||document.activeElement===c)return;if(c.tagName==='SELECT'&&![...c.options].some(o=>o.value===String(v)))c.add(new Option(String(v),String(v)));c.value=v;};
    const en=(k,ok)=>{const c=$('[data-pp="'+k+'"]');if(!c)return;c.disabled=!ok;const f=c.closest('.fld');if(f)f.classList.toggle('off',!ok);};
    const sw=(k,v)=>$$('[data-sw="'+k+'"]').forEach(e=>{e.style.background=v||'';e.classList.toggle('none',!v);});
    const isTx=any(a=>a.type==='tx'),fillOK=any(a=>a.type==='tx'||(a.type==='sh'&&(a.shape==='rect'||a.shape==='ell'))),co=any(a=>a.type==='tx'&&a.co);
    put('author',p.au||author);put('subject',p.sj||U.kindLabel(p));
    en('opacity',any(a=>['tx','sh','dr','im'].includes(a.type)));put('opacity',p.op==null?100:p.op);
    en('border',any(a=>['tx','sh','dr'].includes(a.type)));
    const bc=p.type==='tx'?(p.box||'#444444'):p.color;put('border',HEX(bc)||'#000000');sw('border',p.type==='tx'||p.type==='sh'||p.type==='dr'?bc:'');
    const fb=$('[data-dd="fillc"]');if(fb)fb.disabled=!fillOK;
    const fc=p.type==='tx'?p.bg:p.fill;sw('fill',fc);put('fill',HEX(fc)||'#ffffff');
    en('thick',any(a=>boxTx(a)||a.type==='sh'||a.type==='dr'));put('thick',p.type==='tx'?(p.bw||1):(p.lw||1));
    en('style',any(a=>boxTx(a)||a.type==='sh'));put('style',p.dash||'solid');
    const ln=any(lineSh);en('start',ln);en('end',ln);if(p.type==='sh'){const E=U.ends(p);put('start',E.s);put('end',E.e);}
    en('ll',co);en('lo',co);put('ll',p.ll||40);put('lo',p.lo||0);
    ['font','size','tcolor','cs','ws','ls'].forEach(k=>{const c=$('[data-pp="'+k+'"]');if(c)c.disabled=!isTx;});
    $$('[data-pane=format] [data-a],[data-pane=format] [data-dd]').forEach(b=>{if(b.dataset.a!=='tfinish')b.disabled=!isTx;});
    const fin=$('[data-a=tfinish]');if(fin)fin.disabled=!editId;
    put('font',p.font||'helv');put('size',p.size||12);put('tcolor',HEX(p.color)||'#000000');sw('tcolor',p.color||'#000000');
    put('cs',p.cs||0);put('ws',p.ws||0);put('ls',String(p.ls||1.2));
    $$('[data-a=talign]').forEach(b=>b.classList.toggle('on',(p.align||'l')===b.dataset.v));
  }
  function delSel(){
    const ids=[...sel];
    for(const id of ids){
      const sa=annots.find(a=>a.id===id);if(!sa)continue;
      if(sa.type==='tb'){sa.runs=[];sa.touched=true;const old=annEl(sa.id);if(old)old.replaceWith(mkAnn(sa));continue;}
      annots=annots.filter(a=>a.id!==id);annEls(id).forEach(el=>el.remove());
    }
    sel=sel.filter(id=>annots.some(a=>a.id===id));selId=sel.length?sel[sel.length-1]:null;
    updCtx();push();
  }
  /* ---- perataan & jarak antar objek ---- */
  function boxes(){
    return ppTargets().filter(a=>a.type!=='hl'&&a.type!=='tb').map(a=>{
      const el=annEl(a.id),bb=U.bbox(a);let w=a.w,h=a.h;
      if(a.type==='tx'){w=(el&&el.offsetWidth)||bb.w;h=(el&&el.offsetHeight)||bb.h;}
      return{a,x:a.x,y:a.y,w:w||0,h:h||0};
    });
  }
  function commitBoxes(bs){bs.forEach(b=>{b.a.x=b.x;b.a.y=b.y;const el=annEl(b.a.id);if(el){el.style.left=b.x+'px';el.style.top=b.y+'px';}});push();}
  const ext=bs=>({x0:Math.min(...bs.map(b=>b.x)),y0:Math.min(...bs.map(b=>b.y)),x1:Math.max(...bs.map(b=>b.x+b.w)),y1:Math.max(...bs.map(b=>b.y+b.h))});
  const ALN={
    alnL:bs=>{const m=ext(bs).x0;bs.forEach(b=>{b.x=m;});},
    alnR:bs=>{const m=ext(bs).x1;bs.forEach(b=>{b.x=m-b.w;});},
    alnT:bs=>{const m=ext(bs).y0;bs.forEach(b=>{b.y=m;});},
    alnB:bs=>{const m=ext(bs).y1;bs.forEach(b=>{b.y=m-b.h;});},
    alnV:bs=>{const e=ext(bs),c=(e.y0+e.y1)/2;bs.forEach(b=>{b.y=c-b.h/2;});},
    alnH:bs=>{const e=ext(bs),c=(e.x0+e.x1)/2;bs.forEach(b=>{b.x=c-b.w/2;});},
    ctrV:bs=>{const e=ext(bs),d=pages[bs[0].a.page].h/2-(e.y0+e.y1)/2;bs.forEach(b=>{b.y+=d;});},
    ctrH:bs=>{const e=ext(bs),d=pages[bs[0].a.page].w/2-(e.x0+e.x1)/2;bs.forEach(b=>{b.x+=d;});},
    ctrB:bs=>{ALN.ctrV(bs);ALN.ctrH(bs);},
    distV:bs=>{const o=[...bs].sort((p,q)=>p.y-q.y),e=ext(bs),gap=((e.y1-e.y0)-o.reduce((t,b)=>t+b.h,0))/(o.length-1);let y=e.y0;o.forEach(b=>{b.y=y;y+=b.h+gap;});},
    distH:bs=>{const o=[...bs].sort((p,q)=>p.x-q.x),e=ext(bs),gap=((e.x1-e.x0)-o.reduce((t,b)=>t+b.w,0))/(o.length-1);let x=e.x0;o.forEach(b=>{b.x=x;x+=b.w+gap;});},
    spHp:bs=>{[...bs].sort((p,q)=>p.x-q.x).forEach((b,k)=>{b.x+=k*5;});},
    spHm:bs=>{[...bs].sort((p,q)=>p.x-q.x).forEach((b,k)=>{b.x-=k*5;});},
    spVp:bs=>{[...bs].sort((p,q)=>p.y-q.y).forEach((b,k)=>{b.y+=k*5;});},
    spVm:bs=>{[...bs].sort((p,q)=>p.y-q.y).forEach((b,k)=>{b.y-=k*5;});}
  };
  const ALNMIN={alnL:2,alnR:2,alnT:2,alnB:2,alnV:2,alnH:2,ctrV:1,ctrH:1,ctrB:1,distV:3,distH:3,spHp:2,spHm:2,spVp:2,spVm:2};
  function sizeTo(w,h){
    const bs=boxes();if(bs.length<2)return;
    const ref=bs[0];let n=0,skip=0;
    for(const b of bs.slice(1)){
      if(['sh','im','lk'].includes(b.a.type)){if(w)b.a.w=ref.w;if(h)b.a.h=ref.h;refresh(b.a);n++;}else skip++;
    }
    if(n)push();
    if(skip)toast(skip+' objek (teks/coretan) tidak bisa diubah ukurannya');
  }

  /* ====== Review / Page Layout / Protect: alat gambar & helper ====== */
  function stampImg(key){
    const col=({DISETUJUI:'#2e7d32',DITOLAK:'#c62828',DRAF:'#1565c0',RAHASIA:'#c62828',FINAL:'#2e7d32',SELESAI:'#2e7d32',SALINAN:'#6a1b9a'})[key]||'#1565c0';
    const text=key==='TGL'?new Date().toLocaleDateString('id-ID'):key,fs=44,font='700 '+fs+'px Arial,Helvetica,sans-serif';
    const c=document.createElement('canvas'),g=c.getContext('2d');
    g.font=font;const tw=Math.ceil(g.measureText(text).width);
    c.width=tw+44;c.height=fs+30;
    g.font=font;g.strokeStyle=g.fillStyle=col;g.lineWidth=5;
    const r=10,w=c.width-6,h=c.height-6;
    g.beginPath();g.moveTo(3+r,3);g.arcTo(3+w,3,3+w,3+h,r);g.arcTo(3+w,3+h,3,3+h,r);g.arcTo(3,3+h,3,3,r);g.arcTo(3,3,3+w,3,r);g.closePath();g.stroke();
    g.textAlign='center';g.textBaseline='middle';g.fillText(text,c.width/2,c.height/2+2);
    return c.toDataURL('image/png');
  }
  async function blobToPng(blob){
    const bm=await createImageBitmap(blob),k=Math.min(1,1600/Math.max(bm.width,bm.height)),c=document.createElement('canvas');
    c.width=Math.max(1,Math.round(bm.width*k));c.height=Math.max(1,Math.round(bm.height*k));
    c.getContext('2d').drawImage(bm,0,0,c.width,c.height);return c.toDataURL('image/png');
  }
  function placeCenter(src){const pg=pages[cur]||pages[0];placeImg(cur,[pg.w/2,pg.h/2],src,Math.min(300,pg.w*0.6));}
  /* seret di halaman untuk membuat bentuk / coretan / ukuran / tautan / area redaksi */
  function drawStart(e){
    const pgEl=e.target.closest&&e.target.closest('.pg');if(!pgEl||e.button>0)return;
    e.preventDefault();
    const i=+pgEl.dataset.i,al=pgEl.querySelector('.annl'),pw=pages[i].w,ph=pages[i].h,t=tool,cl=(v,m)=>Math.max(0,Math.min(m,v));
    const[x0,y0]=[cl(pt(e,pgEl)[0],pw),cl(pt(e,pgEl)[1],ph)];
    let x1=x0,y1=y0,tmp=null;const pts=[[x0,y0]];
    const build=()=>{
      if(t==='pen')return U.normDr(pts,U.PENS[penSel],i);
      const x=Math.min(x0,x1),y=Math.min(y0,y1),w=Math.max(1,Math.abs(x1-x0)),h=Math.max(1,Math.abs(y1-y0)),base={page:i,x,y,w,h,fx:x1<x0,fy:y1<y0};
      if(t==='shape')return{type:'sh',shape:shapeSel,color:'#d90000',lw:2,cm:true,...(defs[shapeSel]||{}),...base};
      if(t==='measure')return{type:'sh',shape:'measure',unit:measureUnit,color:'#0b6bcb',lw:1.5,cm:true,...(defs.measure||{}),...base};
      return{type:'lk',...base};
    };
    const show=()=>{
      if(tmp)tmp.remove();
      const o=build();
      if(t==='redarea'||t==='link'){tmp=document.createElement('div');tmp.className='dprev '+t;tmp.style.cssText=`left:${o.x}px;top:${o.y}px;width:${o.w}px;height:${o.h}px`;}
      else{tmp=mkAnn({...o,id:'tmp'});tmp.style.pointerEvents='none';}
      al.appendChild(tmp);
    };
    const mv=m=>{const[x,y]=pt(m,pgEl);x1=cl(x,pw);y1=cl(y,ph);if(t==='pen'){const L=pts[pts.length-1];if(Math.hypot(x1-L[0],y1-L[1])>=1.2)pts.push([x1,y1]);}show();};
    const fin=async()=>{
      desk.removeEventListener('pointermove',mv);desk.removeEventListener('pointerup',fin);desk.removeEventListener('pointercancel',fin);
      if(tmp)tmp.remove();
      const tiny=Math.abs(x1-x0)<3&&Math.abs(y1-y0)<3;
      if(t==='pen'){if(pts.length<2)pts.push([x0+.01,y0]);addAnn(U.normDr(pts,U.PENS[penSel],i));push();return;}
      if(tiny&&(t==='measure'||t==='link'||t==='redarea'||(t==='shape'&&(shapeSel==='line'||shapeSel==='arrow'))))return;
      const o=build();
      if(t==='redarea'){addAnn({type:'hl',kind:'rd',page:i,rects:[{x:o.x,y:o.y,w:o.w,h:o.h}]});push();return;}
      if(t==='link'){const tg=await askLink();if(!tg)return;addAnn({...o,...tg});push();return;}
      if(tiny){o.x=Math.min(x0,pw-100);o.y=Math.min(y0,ph-60);o.w=100;o.h=60;}
      const el=addAnn(o);push();setTool('select');selectQuiet(el.dataset.id);
    };
    desk.setPointerCapture&&desk.setPointerCapture(e.pointerId);
    desk.addEventListener('pointermove',mv);desk.addEventListener('pointerup',fin);desk.addEventListener('pointercancel',fin);
  }
  async function askLink(){
    if(linkMode==='page'){
      const b=await dialog('Tautan ke halaman',`<label>Nomor halaman tujuan (1-${pages.length}):<input type="text" class="lp" value="1"></label>`,'Buat');
      if(!b)return null;const n=parseInt(b.querySelector('.lp').value);
      if(!(n>=1&&n<=pages.length)){toast('Nomor halaman tidak valid');return null;}
      return{dest:n-1};
    }
    const b=await dialog('Tautan ke URL','<label>Alamat URL:<input type="text" class="lu" value="https://"></label>','Buat');
    if(!b)return null;let u=b.querySelector('.lu').value.trim();
    if(!u||u==='https://'){toast('URL kosong');return null;}
    if(!/^[a-z][a-z0-9+.-]*:/i.test(u))u='https://'+u;
    if(/^(javascript|data|vbscript|file):/i.test(u)){toast('Skema URL tidak diizinkan');return null;}
    return{url:u};
  }
  const rangeIn=(b,sel='.rg')=>{const idx=U.parseRange(b.querySelector(sel).value,pages.length);if(!idx||!idx.length){toast('Rentang halaman tidak valid');return null;}return idx;};
  /* gambar teks pada titik menurut tampilan (memperhitungkan putaran halaman & cropbox) */
  function drawVis(PL,p,font,txt,pos,size){
    const cb=p.getCropBox(),R=U.rot(p.getRotation().angle),W=cb.width,H=cb.height,vw=R%180?H:W,vh=R%180?W:H,tw=font.widthOfTextAtSize(txt,size);
    const vx=pos.h==='l'?pos.m:pos.h==='r'?vw-pos.m-tw:(vw-tw)/2,vy=pos.v==='b'?pos.m:vh-pos.m-size*0.72,[u,v]=U.v2p(R,W,H,vx,vy);
    p.drawText(txt,{x:cb.x+u,y:cb.y+v,size,font,color:PL.rgb(0.15,0.15,0.15),rotate:PL.degrees(R)});
  }
  /* kotak persegi tiap kecocokan teks, dihitung dari lapisan teks yang sudah dirender */
  function rectsOfMatches(i,p,q){
    const re=reOf(q),strs=p.strs,offs=[];let pos=0;for(const t of strs){offs.push(pos);pos+=t.length;}
    const text=strs.join(''),b=pgs.children[i].getBoundingClientRect(),rs=[];let m;
    while((m=re.exec(text))){
      if(!m[0].length){re.lastIndex++;continue;}
      const a=m.index,z=a+m[0].length;
      for(let k=0;k<strs.length;k++){
        const d=p.divs[k];if(!d||!d.firstChild||offs[k]>=z||offs[k]+strs[k].length<=a)continue;
        const r=document.createRange();r.setStart(d.firstChild,Math.max(0,a-offs[k]));r.setEnd(d.firstChild,Math.min(strs[k].length,z-offs[k]));
        for(const c of r.getClientRects())if(c.width>.5&&c.height>.5)rs.push({x:(c.left-b.left)/S,y:(c.top-b.top)/S,w:c.width/S,h:c.height/S});
      }
    }
    return rs;
  }
  async function findRects(q){
    const out=[];let missed=0;
    for(let i=0;i<pages.length;i++){
      const tc=await getText(i),exp=(tc.items.map(t=>t.str||'').join('').match(reOf(q))||[]).length;
      if(!exp)continue;
      let rs=[];
      for(let k=0;k<3&&!rs.length;k++){await renderPage(i);const p=pages[i];if(p.divs)rs=rectsOfMatches(i,p,q);}
      if(rs.length)out.push({page:i,rects:mergeLine(rs)});else missed+=exp;
    }
    return{out,missed};
  }
  function commentList(){
    const KIND={hl:'Sorotan',ul:'Garis bawah',st:'Coret'},out=[];
    for(const a of annots){
      if(!a.cm)continue;let t='',k='';
      if(a.type==='hl')k=KIND[a.kind||'hl'];
      else if(a.type==='tx'){k=a.co?'Callout':a.bg?'Catatan':a.box?'Kotak teks':'Komentar';t=a.text;}
      else if(a.type==='sh'){k=a.shape==='measure'?'Ukuran':'Bentuk';t=a.shape==='measure'?U.measureLabel(a):a.shape;}
      else if(a.type==='dr')k='Coretan';
      else if(a.type==='im')k='Stempel';
      else continue;
      out.push({page:a.page+1,kind:k,text:t});
    }
    return out.sort((x,y)=>x.page-y.page);
  }

  /* tanda tangan */
  function signDlg(){
    return dialog('Gambar tanda tangan','<canvas width="420" height="160"></canvas><small>Gambar dengan mouse atau jari.</small><button class="clr" style="align-self:flex-start">Bersihkan</button>','Gunakan',b=>{
      const c=b.querySelector('canvas'),x=c.getContext('2d');x.lineWidth=2.6;x.lineCap=x.lineJoin='round';x.strokeStyle='#14233c';let dr=false;
      const p=e=>{const r=c.getBoundingClientRect();return[(e.clientX-r.left)*c.width/r.width,(e.clientY-r.top)*c.height/r.height];};
      c.onpointerdown=e=>{dr=true;c.setPointerCapture(e.pointerId);const[a,d]=p(e);x.beginPath();x.moveTo(a,d);x.lineTo(a+.1,d+.1);x.stroke();};
      c.onpointermove=e=>{if(!dr)return;const[a,d]=p(e);x.lineTo(a,d);x.stroke();};
      c.onpointerup=c.onpointercancel=()=>{dr=false;};
      b.querySelector('.clr').onclick=()=>x.clearRect(0,0,c.width,c.height);
    }).then(b=>{
      if(!b)return null;
      const c=b.querySelector('canvas'),d=c.getContext('2d').getImageData(0,0,c.width,c.height).data;
      let x0=c.width,y0=c.height,x1=-1,y1=-1;
      for(let y=0;y<c.height;y++)for(let x=0;x<c.width;x++)if(d[(y*c.width+x)*4+3]>8){if(x<x0)x0=x;if(x>x1)x1=x;if(y<y0)y0=y;if(y>y1)y1=y;}
      if(x1<0){toast('Tanda tangan masih kosong');return null;}
      const m=4,o=document.createElement('canvas');o.width=x1-x0+1+m*2;o.height=y1-y0+1+m*2;
      o.getContext('2d').drawImage(c,x0-m,y0-m,o.width,o.height,0,0,o.width,o.height);
      return o.toDataURL('image/png');
    });
  }

  /* ---------- menu tarik-turun ---------- */
  function closeDD(){const m=pd.querySelector('.pmenu');if(m)m.remove();ddBtn=null;}
  function openDD(btn){
    const key=btn.dataset.dd,items=DD[key](sign);
    const m=document.createElement('div');m.className='pmenu';
    m.innerHTML=items.map(it=>it==='-'?'<hr>':`<button data-a="${it[1].split(':')[0]}" data-v="${it[1].split(':')[1]??''}">${esc(it[0])}</button>`).join('');
    pd.appendChild(m);
    const b=btn.getBoundingClientRect(),p=pd.getBoundingClientRect();
    m.style.left=Math.max(0,Math.min(b.left-p.left,p.width-m.offsetWidth-4))+'px';m.style.top=(b.bottom-p.top)+'px';
    ddBtn=btn;
  }

  /* ---------- operasi dokumen (pdf-lib) ---------- */
  const needDoc=()=>{if(!pdfDoc){toast('Belum ada dokumen');return false;}return true;};
  const redOk=()=>!annots.some(isRd)||confirm('Ada tanda redaksi yang belum diterapkan: teks di bawahnya MASIH ada di berkas. Lanjutkan tanpa menerapkan redaksi?');
  async function commit(){
    const keep=annots.filter(isRd),flat=annots.filter(a=>!isRd(a));
    if(!flat.length)return;
    await lib('pdflib',()=>window.PDFLib);
    const doc=await PDFLib.PDFDocument.load(bytes,{ignoreEncryption:true});
    await U.flatten(PDFLib,doc,flat,pages);
    const nb=await doc.save();pristine=false;await apply(nb,cur,keep);
  }
  async function edit(fn,page){
    await commit();await lib('pdflib',()=>window.PDFLib);
    const doc=await PDFLib.PDFDocument.load(bytes,{ignoreEncryption:true});
    await fn(doc,PDFLib);
    const nb=await doc.save();pristine=false;await apply(nb,page);
  }
  /* ---------- banyak berkas: tiap berkas menyimpan state & riwayat undo/redo sendiri ---------- */
  const D=DOCS;
  function snap(){
    const d=D.list.find(x=>x.id===D.active);if(!d)return;
    Object.assign(d,{name,bytes,annots:clone(annots),hist,hi,dirty,pristine,cur,zoom,autoFit,top:desk.scrollTop});
  }
  async function enter(d){
    fq='';fm=[];fi=-1;fd.hidden=true;
    name=d.name;bytes=d.bytes;annots=clone(d.annots);hist=d.hist;hi=d.hi;dirty=d.dirty;pristine=d.pristine;
    cur=d.cur;zoom=d.zoom;autoFit=d.autoFit;resetSel();D.active=d.id;
    setName(name);btns();
    await load(bytes,{page:cur});
    if(d.top)desk.scrollTop=d.top;
    renderList();
  }
  async function addDoc(o){
    snap();const prev=D.list.find(x=>x.id===D.active);
    const d={id:'d'+(++D.seq),name:'Untitled1.pdf',bytes:null,annots:[],hist:[],hi:-1,dirty:false,pristine:false,cur:0,zoom:100,autoFit:true,top:0,...o};
    D.list.push(d);
    try{await enter(d);return true;}
    catch(e){
      console.error(e);D.list=D.list.filter(x=>x!==d);D.active=prev?prev.id:null;
      if(prev)await enter(prev).catch(()=>{});
      toast('Tidak bisa membuka PDF: '+(e&&e.message||e));return false;
    }
  }
  function renderList(){
    const el=$('.ofl');if(!el)return;
    el.innerHTML=D.list.length?'<hr><div class="lbl">Berkas terbuka</div>'+D.list.map(d=>{
      const on=d.id===D.active,nm=on?name:d.name,dt=on?dirty:d.dirty;
      return`<div class="of${on?' on':''}"><button class="ofn" data-a="sw" data-v="${d.id}" title="${esc(nm)}">${esc(nm)}${dt?' •':''}</button><button class="ofx" data-a="cl" data-v="${d.id}" title="Tutup berkas ini" aria-label="Tutup ${esc(nm)}">×</button></div>`;
    }).join(''):'';
  }
  async function closeDoc(id){
    const d=D.list.find(x=>x.id===id);if(!d)return;
    const isA=id===D.active,nm=isA?name:d.name;
    if((isA?dirty:d.dirty)&&!confirm('"'+nm+'" punya perubahan yang belum disimpan. Tutup tanpa menyimpan?'))return;
    const i=D.list.indexOf(d);D.list.splice(i,1);
    if(!isA){renderList();return;}
    if(!D.list.length){D.active=null;menu.hidden=true;opts.onHome&&opts.onHome();return;}
    await enter(D.list[Math.min(i,D.list.length-1)]);
  }
  function nextName(){
    const used=new Set(D.list.map(x=>x.id===D.active?name:x.name));
    let n=1;while(used.has('Untitled'+n+'.pdf'))n++;return 'Untitled'+n+'.pdf';
  }
  async function newPdf(buildFn,label){
    await lib('pdflib',()=>window.PDFLib);
    const doc=await PDFLib.PDFDocument.create();await buildFn(doc,PDFLib);
    if(!doc.getPageCount())return;
    const nb=await doc.save();
    await addDoc({name:label||nextName(),bytes:nb,hist:[{bytes:nb,annots:[]}],hi:0});
  }
  async function blankDoc(){
    await lib('pdflib',()=>window.PDFLib);
    const doc=await PDFLib.PDFDocument.create();doc.addPage([595.28,841.89]);
    const nb=await doc.save();
    await addDoc({name:nextName(),bytes:nb,hist:[{bytes:nb,annots:[]}],hi:0,pristine:true});
  }
  async function openBytes(buf,fname){
    if(String.fromCharCode(...buf.slice(0,5))!=='%PDF-'){toast('Berkas bukan PDF yang valid: '+fname);return;}
    await addDoc({name:fname,bytes:buf,hist:[{bytes:buf,annots:[]}],hi:0});
  }
  async function openFile(f){await openBytes(new Uint8Array(await f.arrayBuffer()),f.name);}
  async function embedImg(doc,f){
    let b=new Uint8Array(await f.arrayBuffer());
    if(f.type==='image/png')return doc.embedPng(b);
    if(f.type==='image/jpeg')return doc.embedJpg(b);
    const bm=await createImageBitmap(f),c=document.createElement('canvas');c.width=bm.width;c.height=bm.height;c.getContext('2d').drawImage(bm,0,0);
    const blob=await new Promise(r=>c.toBlob(r,'image/png'));return doc.embedPng(new Uint8Array(await blob.arrayBuffer()));
  }
  async function imgPage(doc,f,at){
    const im=await embedImg(doc,f);
    let w=im.width*0.75,h=im.height*0.75;const [mw,mh]=w>h?[841.89,595.28]:[595.28,841.89],k=Math.min(1,mw/w,mh/h);w*=k;h*=k;
    const pg=at==null?doc.addPage([w,h]):doc.insertPage(at,[w,h]);pg.drawImage(im,{x:0,y:0,width:w,height:h});
  }

  /* ---------- ekstraksi teks / gambar ---------- */
  async function pageLines(i){
    const tc=await getText(i);
    return U.groupLines(tc.items.filter(t=>typeof t.str==='string').map(t=>({s:t.str,x:t.transform[4],y:t.transform[5],w:t.width,h:Math.abs(t.transform[3])||t.height||10})));
  }
  async function renderCanvas(i,scale){
    const pg=await pdfDoc.getPage(i+1),vp=pg.getViewport({scale}),c=document.createElement('canvas');
    c.width=Math.ceil(vp.width);c.height=Math.ceil(vp.height);
    await pg.render({canvasContext:c.getContext('2d'),viewport:vp,background:'#fff'}).promise;return c;
  }
  const toBlob=(c,m,q)=>new Promise(r=>c.toBlob(r,m,q));
  function imgToCanvas(img){
    const c=document.createElement('canvas');
    if(img.bitmap){c.width=img.bitmap.width;c.height=img.bitmap.height;c.getContext('2d').drawImage(img.bitmap,0,0);return c;}
    const w=img.width,h=img.height,d=img.data;if(!w||!h||!d)return null;
    c.width=w;c.height=h;const ctx=c.getContext('2d'),id=ctx.createImageData(w,h),o=id.data;
    if(d.length===w*h*4)o.set(d);
    else if(d.length===w*h*3){for(let s=0,t=0;t<o.length;s+=3,t+=4){o[t]=d[s];o[t+1]=d[s+1];o[t+2]=d[s+2];o[t+3]=255;}}
    else if(d.length===w*h){for(let s=0,t=0;t<o.length;s++,t+=4){o[t]=o[t+1]=o[t+2]=d[s];o[t+3]=255;}}
    else if(d.length>=((w+7)>>3)*h){const rb=(w+7)>>3;for(let y=0;y<h;y++)for(let x=0;x<w;x++){const v=((d[y*rb+(x>>3)]>>(7-(x&7)))&1)?255:0,t=(y*w+x)*4;o[t]=o[t+1]=o[t+2]=v;o[t+3]=255;}}
    else return null;
    ctx.putImageData(id,0,0);return c;
  }
  async function saveMany(blobs,ext,suffix){
    if(blobs.length===1){dl(blobs[0],`${base()}.${ext}`);return;}
    await lib('jszip',()=>window.JSZip);const z=new JSZip(),pad=String(blobs.length).length;
    blobs.forEach((b,i)=>z.file(`${base()}_${String(i+1).padStart(pad,'0')}.${ext}`,b));
    dl(await z.generateAsync({type:'blob'}),`${base()}_${suffix}.zip`);
  }

  /* ---------- konversi ---------- */
  async function convert(kind,arg){
    if(!needDoc()||!redOk())return;
    if(pristine){toast('Dokumen masih kosong, tidak ada yang dikonversi');return;}
    await commit();
    const n=pages.length,pr=note('Mengonversi…');
    try{
      if(kind==='word'||kind==='rich'||kind==='plain'||kind==='excel'){
        const all=[];for(let i=0;i<n;i++){pr.set(`Membaca teks… ${i+1}/${n}`);all.push(await pageLines(i));}
        if(!all.some(l=>l.length)){toast('Tidak ada teks di PDF ini (PDF hasil pindai membutuhkan OCR)');return;}
        if(kind==='plain'){dl(new Blob([all.map(ls=>ls.map(l=>l.text).join('\n')).join('\n\n')],{type:'text/plain;charset=utf-8'}),base()+'.txt');}
        else if(kind==='rich'){dl(new Blob([U.rtf(all.map(U.paras))],{type:'application/rtf'}),base()+'.rtf');}
        else if(kind==='excel'){
          await lib('xlsx',()=>window.XLSX);const wb=XLSX.utils.book_new();
          all.forEach((ls,i)=>{const rows=U.sheetRows(ls);XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows.length?rows:[['']]),'Hal '+(i+1));});
          dl(new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array'})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),base()+'.xlsx');
        }else{
          await lib('docx',()=>window.docx);const{Document,Packer,Paragraph,TextRun}=docx,kids=[];
          all.forEach((ls,i)=>{const ps=U.paras(ls);
            if(!ps.length)kids.push(new Paragraph({pageBreakBefore:i>0,children:[]}));
            ps.forEach((p,k)=>kids.push(new Paragraph({pageBreakBefore:i>0&&k===0,spacing:{after:120},children:[new TextRun({text:p.text,bold:p.bold,size:Math.max(16,Math.min(144,Math.round(p.size*2)))})]})));});
          const d=new Document({styles:{default:{document:{run:{font:'Calibri'}}}},sections:[{properties:{page:{size:{width:Math.round(pages[0].w*20),height:Math.round(pages[0].h*20)},margin:{top:1134,bottom:1134,left:1134,right:1134}}},children:kids}]});
          dl(await Packer.toBlob(d),base()+'.docx');
        }
      }else if(kind==='ppt'){
        await lib('pptx',()=>window.PptxGenJS);const px=new PptxGenJS(),W=pages[0].w/72,H=pages[0].h/72;
        px.defineLayout({name:'PDFPG',width:W,height:H});px.layout='PDFPG';
        for(let i=0;i<n;i++){
          pr.set(`Membuat slide… ${i+1}/${n}`);
          const c=await renderCanvas(i,2),sl=px.addSlide(),k=Math.min(W/(pages[i].w/72),H/(pages[i].h/72)),w=pages[i].w/72*k,h=pages[i].h/72*k;
          sl.addImage({data:c.toDataURL('image/jpeg',.9).slice(5),x:(W-w)/2,y:(H-h)/2,w,h});
          const t=(await pageLines(i)).map(l=>l.text).join('\n');if(t)sl.addNotes(t);
        }
        dl(await px.write({outputType:'blob'}),base()+'.pptx');
      }else if(kind==='image'){
        const{fmt,scale,idx}=arg,blobs=[];
        for(let k=0;k<idx.length;k++){pr.set(`Merender halaman ${idx[k]+1}… (${k+1}/${idx.length})`);
          blobs.push(await toBlob(await renderCanvas(idx[k],scale),fmt==='jpg'?'image/jpeg':'image/png',.92));}
        await saveMany(blobs,fmt,'gambar');
      }else if(kind==='extimg'){
        const OPS=pdfjsLib.OPS,blobs=[],seen=new Set();
        for(let i=0;i<n;i++){
          pr.set(`Mencari gambar… ${i+1}/${n}`);
          const pg=await pdfDoc.getPage(i+1),ol=await pg.getOperatorList();
          for(let k=0;k<ol.fnArray.length;k++){
            const f=ol.fnArray[k];let img=null;
            if(f===OPS.paintImageXObject){
              const id=ol.argsArray[k][0];if(seen.has(id))continue;seen.add(id);
              const st=String(id).startsWith('g_')?pg.commonObjs:pg.objs;
              img=await new Promise(r=>{let ok=false;try{st.get(id,v=>{ok=true;r(v);});}catch{r(null);}setTimeout(()=>{if(!ok)r(null);},3000);});
            }else if(f===OPS.paintInlineImageXObject)img=ol.argsArray[k][0];
            if(!img)continue;
            const c=imgToCanvas(img);if(!c||c.width<24||c.height<24)continue;
            blobs.push(await toBlob(c,'image/png'));
          }
        }
        if(!blobs.length){toast('Tidak ada gambar yang bisa diekstrak');return;}
        await saveMany(blobs,'png','gambar-ekstrak');toast(blobs.length+' gambar diekstrak');
      }
    }catch(e){console.error(e);toast('Gagal mengonversi: '+(e&&e.message||e));}
    finally{pr.end();}
  }

  /* ---------- aksi ---------- */
  const acts={
    async save(){if(!needDoc()||!redOk())return;await run(async()=>{await commit();dl(new Blob([bytes],{type:'application/pdf'}),/\.pdf$/i.test(name)?name:name+'.pdf');dirty=false;toast('Tersimpan: '+name);});},
    undo(){run(()=>restore(hi-1));},redo(){run(()=>restore(hi+1));},
    menu(){menu.hidden=!menu.hidden;if(!menu.hidden)renderList();},
    home(){snap();opts.onHome&&opts.onHome();},
    sw(id){if(id===D.active){menu.hidden=true;return;}run(async()=>{snap();menu.hidden=true;await enter(D.list.find(x=>x.id===id));});},
    cl(id){run(()=>closeDoc(id));},
    blank(){menu.hidden=true;run(blankDoc);},
    async open(){const fs=await pickFiles('.pdf,application/pdf',true);if(!fs.length)return;menu.hidden=true;run(async()=>{for(const f of fs)await openFile(f);});},
    async print(){if(!needDoc()||!redOk())return;await run(async()=>{
      await commit();const pr=note('Menyiapkan cetak…'),box=document.createElement('div');box.className='pd-print';
      try{for(let i=0;i<pages.length;i++){const im=new Image();im.src=(await renderCanvas(i,2)).toDataURL('image/jpeg',.92);await im.decode().catch(()=>{});box.appendChild(im);}}finally{pr.end();}
      document.body.appendChild(box);const clean=()=>{box.remove();removeEventListener('afterprint',clean);};addEventListener('afterprint',clean);setTimeout(()=>print(),50);});},
    /* Create */
    async fromimg(){const fs=await pickFiles('image/*',true);if(!fs.length)return;run(()=>newPdf(async d=>{for(const f of fs)await imgPage(d,f);},'Dari-gambar.pdf'));},
    async fromtxt(){const f=(await pickFiles('.txt,text/plain'))[0];if(!f)return;run(()=>newPdf(async(d,PL)=>{
      const font=await d.embedFont(PL.StandardFonts.Helvetica),lines=U.wrap(await f.text(),font,11,595.28-112);
      let pg=null,y=0;for(const ln of lines){if(!pg||y<56){pg=d.addPage([595.28,841.89]);y=841.89-56;}if(ln)pg.drawText(ln,{x:56,y:y-11,size:11,font});y-=14;}
      if(!lines.length)d.addPage([595.28,841.89]);},f.name.replace(/\.txt$/i,'')+'.pdf'));},
    async combine(){
      const fs=await pickFiles('.pdf,application/pdf',true);if(!fs.length)return;
      run(async()=>{await commit();await lib('pdflib',()=>window.PDFLib);
        const out=await PDFLib.PDFDocument.create(),srcs=[];if(pdfDoc&&!pristine)srcs.push(bytes);
        for(const f of fs)srcs.push(new Uint8Array(await f.arrayBuffer()));
        for(const b of srcs){const s=await PDFLib.PDFDocument.load(b,{ignoreEncryption:true});(await out.copyPages(s,s.getPageIndices())).forEach(p=>out.addPage(p));}
        const nb=await out.save();pristine=false;await apply(nb,0);if(/^Untitled\d+\.pdf$/.test(name))setName('Gabungan.pdf');toast(`Digabung: ${out.getPageCount()} halaman`);});
    },
    /* Page layout */
    insblank(){if(!needDoc())return;run(()=>edit((d)=>{const p=d.getPage(cur),{width,height}=p.getSize();d.insertPage(cur+1,[width,height]);},cur+1));},
    async inspdf(){if(!needDoc())return;const f=(await pickFiles('.pdf,application/pdf'))[0];if(!f)return;
      run(()=>edit(async d=>{const s=await PDFLib.PDFDocument.load(new Uint8Array(await f.arrayBuffer()),{ignoreEncryption:true});
        (await d.copyPages(s,s.getPageIndices())).forEach((p,k)=>d.insertPage(cur+1+k,p));},cur+1));},
    async insimg(){if(!needDoc())return;const fs=await pickFiles('image/*',true);if(!fs.length)return;
      run(()=>edit(async d=>{for(let k=0;k<fs.length;k++)await imgPage(d,fs[k],cur+1+k);},cur+1));},
    rot(v){if(!needDoc())return;const a=+v;run(()=>edit((d,PL)=>{const p=d.getPage(cur);p.setRotation(PL.degrees(((p.getRotation().angle+a)%360+360)%360));},cur));},
    rotall(v){if(!needDoc())return;const a=+v;run(()=>edit((d,PL)=>{d.getPages().forEach(p=>p.setRotation(PL.degrees(((p.getRotation().angle+a)%360+360)%360)));},cur));},
    del(){if(!needDoc())return;if(pages.length<2){toast('Tidak bisa menghapus satu-satunya halaman');return;}const t=cur;run(()=>edit(d=>d.removePage(t),Math.min(t,pages.length-2)));},
    async extract(){
      if(!needDoc())return;
      const b=await dialog('Ekstrak halaman',`<label>Halaman (mis. 1-3, 5):<input type="text" class="rg" value="${cur+1}"></label><small>Dokumen ini memiliki ${pages.length} halaman.</small>`,'Ekstrak');if(!b)return;
      const idx=U.parseRange(b.querySelector('.rg').value,pages.length);if(!idx||!idx.length){toast('Rentang halaman tidak valid');return;}
      run(async()=>{await commit();await lib('pdflib',()=>window.PDFLib);
        const s=await PDFLib.PDFDocument.load(bytes,{ignoreEncryption:true}),o=await PDFLib.PDFDocument.create();
        (await o.copyPages(s,idx)).forEach(p=>o.addPage(p));
        dl(new Blob([await o.save()],{type:'application/pdf'}),base()+'_ekstrak.pdf');toast(idx.length+' halaman diekstrak');});
    },
    find(){fd.hidden=false;fqIn.focus();fqIn.select();},
    fclose(){fd.hidden=true;fq='';fm=[];fi=-1;pages.forEach((_,i)=>applyHits(i));},
    fnext(){stepFind(1);},fprev(){stepFind(-1);},
    addtools(){
      dialog('Favorite Tools','<small>Pilih alat yang ingin ditampilkan di Home.</small>'+FAVS.map(f=>`<label><input type="checkbox" value="${f[0]}"${favs.includes(f[0])?' checked':''}>${esc(f[1])}</label>`).join(''),'Simpan').then(b=>{
        if(!b)return;favs=[...b.querySelectorAll('input:checked')].map(i=>i.value);store.set('favs',favs);drawFavs();});
    },
    /* Convert */
    toword(){convert('word');},toexcel(){convert('excel');},topptx(){convert('ppt');},torich(){convert('rich');},toplain(){convert('plain');},extimg(){convert('extimg');},
    async toimage(){
      if(!needDoc())return;
      const b=await dialog('Ke gambar',`<label>Format:<select class="fm"><option value="png">PNG</option><option value="jpg">JPEG</option></select></label><label>Resolusi:<select class="dp"><option value="72">72 dpi</option><option value="150" selected>150 dpi</option><option value="300">300 dpi</option></select></label><label>Halaman:<input type="text" class="rg" value="semua"></label>`,'Konversi');if(!b)return;
      const idx=U.parseRange(b.querySelector('.rg').value,pages.length);if(!idx||!idx.length){toast('Rentang halaman tidak valid');return;}
      run(()=>convert('image',{fmt:b.querySelector('.fm').value,scale:+b.querySelector('.dp').value/72,idx}));
    },
    pdfa(){toast('Konversi PDF/A belum didukung di versi web ini (butuh font tertanam & profil warna ICC).');},
    async portfolio(){
      const fs=await pickFiles('*/*',true);if(!fs.length)return;
      run(()=>newPdf(async(d,PL)=>{
        const font=await d.embedFont(PL.StandardFonts.Helvetica),bold=await d.embedFont(PL.StandardFonts.HelveticaBold),pg=d.addPage([595.28,841.89]);
        pg.drawText('PDF Portfolio',{x:56,y:770,size:24,font:bold});
        pg.drawText(U.safe(`Berisi ${fs.length} berkas terlampir. Buka di pembaca PDF yang mendukung portfolio.`),{x:56,y:745,size:11,font});
        let y=710;
        for(const f of fs){
          await d.attach(new Uint8Array(await f.arrayBuffer()),f.name,{mimeType:f.type||'application/octet-stream',creationDate:new Date(f.lastModified),modificationDate:new Date(f.lastModified)});
          if(y>60){pg.drawText(U.safe(`• ${f.name}  (${Math.max(1,Math.round(f.size/1024))} KB)`),{x:56,y,size:11,font});y-=18;}
        }
        d.catalog.set(PL.PDFName.of('Collection'),d.context.obj({Type:'Collection',View:'D'}));
      },'Portfolio.pdf'));
    },
    /* tanda tangan */
    async signnew(){const s=await signDlg();if(!s)return;sign=s;store.set('sign',s);setTool('sign');toast('Klik di halaman untuk menempatkan tanda tangan');},
    signuse(){setTool('sign');toast('Klik di halaman untuk menempatkan tanda tangan');},
    signdel(){sign=null;store.set('sign',null);if(tool==='sign')setTool('select');toast('Tanda tangan tersimpan dihapus');},
    /* zoom / navigasi */
    zin(){zstep(1);},zout(){zstep(-1);},z(v){autoFit=false;setZoom(+v);},
    zfitw(){autoFit=false;setZoom(fitW());},
    zfitp(){if(!pages.length)return;autoFit=false;const p=pages[cur]||pages[0];setZoom(Math.min(fitW(),Math.floor((desk.clientHeight-30)/p.h*75)));},
    first(){goto(0);},last(){goto(pages.length-1);},prev(){goto(Math.max(0,cur-1));},next(){goto(Math.min(pages.length-1,cur+1));}
  };
  /* ====== aksi tab Review / Page Layout / Protect ====== */
  const NOSUP={
    ocr:['OCR','butuh mesin pengenalan teks (OCR) yang belum tersedia di versi web ini'],
    migrate:['Migrate','belum tersedia di versi web ini'],
    deskew:['Auto De-skew','butuh pemrosesan gambar halaman pindaian yang belum tersedia'],
    certsec:['Certificate Security','butuh infrastruktur sertifikat digital (PKI)'],
    profiles:['Manage Profiles','butuh infrastruktur sertifikat digital (PKI)'],
    certify:['Certify','butuh sertifikat digital (PKI)'],
    digid:['My Digital IDs','butuh penyimpanan sertifikat digital'],
    trusted:['Trusted Contacts','butuh penyimpanan sertifikat digital']
  };
  Object.keys(NOSUP).forEach(k=>{acts[k]=()=>toast(NOSUP[k][0]+' belum didukung: '+NOSUP[k][1]+'.');});
  Object.assign(acts,{
    /* --- Review --- */
    stamp(v){stampSel=v;setTool('stamp');toast('Klik di halaman untuk menempatkan stempel');},
    shape(v){shapeSel=v;setTool('shape');toast('Seret di halaman untuk menggambar');},
    pen(v){penSel=v;setTool('pen');toast('Gambar dengan mouse atau jari');},
    measure(v){measureUnit=v;setTool('measure');toast('Seret dari titik awal ke titik akhir');},
    link(v){linkMode=v;setTool('link');toast('Seret di halaman untuk membuat area tautan');},
    async addimg(){if(!needDoc())return;const f=(await pickFiles('image/*'))[0];if(!f)return;try{placeCenter(await blobToPng(f));}catch(e){toast('Gagal memuat gambar');}},
    async pasteimg(){
      if(!needDoc())return;
      try{for(const it of await navigator.clipboard.read()){const ty=it.types.find(x=>x.startsWith('image/'));if(ty){placeCenter(await blobToPng(await it.getType(ty)));return;}}toast('Tidak ada gambar di clipboard');}
      catch(e){toast('Tidak bisa membaca clipboard (izin ditolak?)');}
    },
    async attach(){
      if(!needDoc())return;const fs=await pickFiles('*/*',true);if(!fs.length)return;
      const data=[];for(const f of fs)data.push([f.name,new Uint8Array(await f.arrayBuffer()),f.type||'application/octet-stream']);
      run(()=>edit(async d=>{for(const[n,b,m]of data)await d.attach(b,n,{mimeType:m});},cur));
      toast(fs.length+' berkas dilampirkan');
    },
    async compare(){
      if(!needDoc())return;const f=(await pickFiles('.pdf,application/pdf'))[0];if(!f)return;
      run(async()=>{
        await commit();const pr=note('Membandingkan teks…');let od=null;
        try{
          await initPdfjs();od=await pdfjsLib.getDocument({data:new Uint8Array(await f.arrayBuffer())}).promise;
          const n=Math.max(pages.length,od.numPages),norm=t=>t.replace(/\s+/g,' ').trim(),diffs=[];
          for(let i=0;i<n;i++){
            const a=i<pages.length?norm((await getText(i)).items.map(t=>t.str).join(' ')):null;
            const b=i<od.numPages?norm((await(await od.getPage(i+1)).getTextContent()).items.map(t=>t.str).join(' ')):null;
            if(a===null)diffs.push(`Hal. ${i+1}: hanya ada di "${f.name}"`);
            else if(b===null)diffs.push(`Hal. ${i+1}: hanya ada di dokumen ini`);
            else if(a!==b){const d=U.wordDiff(a,b);diffs.push(`Hal. ${i+1}: teks berbeda (+${d.add} / −${d.del} kata di berkas pembanding)`);}
          }
          pr.end();
          await dialog('Hasil perbandingan',`<div style="max-height:40vh;overflow:auto">${diffs.length?diffs.map(x=>`<div>${esc(x)}</div>`).join(''):'<div>Teks kedua dokumen identik.</div>'}</div><small>Hanya membandingkan teks per halaman, bukan tampilan visual.</small>`,'Tutup');
        }finally{pr.end();if(od)od.destroy();}
      });
    },
    sumshow(){
      const L=commentList();if(!L.length){toast('Belum ada komentar di dokumen ini');return;}
      dialog('Ringkasan komentar',`<div style="max-height:40vh;overflow:auto">${L.map(c=>`<div>Hal. ${c.page} — <b>${esc(c.kind)}</b>${c.text?': '+esc(c.text):''}</div>`).join('')}</div><small>${L.length} komentar yang belum dibakukan ke dokumen.</small>`,'Ekspor .txt').then(b=>{if(b)acts.sumexp();});
    },
    sumexp(){
      const L=commentList();if(!L.length){toast('Belum ada komentar');return;}
      dl(new Blob([`Ringkasan komentar — ${name}\n\n`+L.map(c=>`Hal. ${c.page} — ${c.kind}${c.text?': '+c.text.replace(/\n/g,' '):''}`).join('\n')+'\n'],{type:'text/plain;charset=utf-8'}),base()+'_komentar.txt');
    },
    hidecm(){
      hidecm=!hidecm;desk.classList.toggle('hidecm',hidecm);
      const b=$('[data-a=hidecm]');if(b){b.classList.toggle('on',hidecm);b.querySelector('span').textContent=hidecm?'Show Comments':'Hide Comments';}
    },
    /* --- Page Layout --- */
    async crop(){
      if(!needDoc())return;
      const b=await dialog('Pangkas halaman (Crop)',`<small>Margin yang dipangkas dari tiap sisi (mm), menurut tampilan halaman.</small><label>Atas:<input type="text" class="m-t" value="10"></label><label>Bawah:<input type="text" class="m-b" value="10"></label><label>Kiri:<input type="text" class="m-l" value="10"></label><label>Kanan:<input type="text" class="m-r" value="10"></label><label>Halaman:<input type="text" class="rg" value="${cur+1}"></label><small>Isi "semua" untuk seluruh dokumen (${pages.length} halaman).</small>`,'Pangkas');
      if(!b)return;
      const mm=c=>Math.max(0,parseFloat(b.querySelector(c).value.replace(',','.'))||0)*72/25.4,m={t:mm('.m-t'),b:mm('.m-b'),l:mm('.m-l'),r:mm('.m-r')},idx=rangeIn(b);
      if(!idx)return;if(!(m.t+m.b+m.l+m.r)){toast('Isi minimal satu margin');return;}
      run(()=>edit(d=>{
        let bad=0;
        for(const i of idx){const pg=d.getPage(i),cb=pg.getCropBox(),M=U.cropMargins(U.rot(pg.getRotation().angle),m),w=cb.width-M.L-M.R,h=cb.height-M.T-M.B;
          if(w<20||h<20){bad++;continue;}pg.setCropBox(cb.x+M.L,cb.y+M.B,w,h);}
        if(bad)setTimeout(()=>toast(bad+' halaman dilewati: margin terlalu besar'),0);
      },cur));
    },
    async replpg(){
      if(!needDoc())return;const f=(await pickFiles('.pdf,application/pdf'))[0];if(!f)return;
      await lib('pdflib',()=>window.PDFLib);
      const src=new Uint8Array(await f.arrayBuffer());let sn;
      try{sn=(await PDFLib.PDFDocument.load(src,{ignoreEncryption:true})).getPageCount();}catch(e){toast('Berkas bukan PDF yang valid');return;}
      const b=await dialog('Ganti halaman',`<label>Ganti halaman dokumen ini (mis. 2-3):<input type="text" class="rg" value="${cur+1}"></label><label>Dengan halaman dari "${esc(f.name)}" (${sn} hal.):<input type="text" class="rg2" value="semua"></label>`,'Ganti');
      if(!b)return;
      const a=rangeIn(b),c=U.parseRange(b.querySelector('.rg2').value,sn);
      if(!a)return;if(!c||!c.length){toast('Rentang halaman berkas pembanding tidak valid');return;}
      run(()=>edit(async d=>{
        const sd=await PDFLib.PDFDocument.load(src,{ignoreEncryption:true}),cp=await d.copyPages(sd,c),at=a[0];
        [...a].reverse().forEach(i=>d.removePage(i));cp.forEach((pg,k)=>d.insertPage(at+k,pg));
      },a[0]));
    },
    async splitdoc(){
      if(!needDoc()||!redOk())return;
      const b=await dialog('Pisah dokumen',`<label>Pisah per berapa halaman / rentang:<input type="text" class="sp" value="1"></label><small>Contoh: "2" = tiap 2 halaman satu berkas; "1-3; 4-6; 7" = bagian menurut rentang. Dokumen ini ${pages.length} halaman. Hasil diunduh sebagai .zip.</small>`,'Pisah');
      if(!b)return;
      const plan=U.splitPlan(b.querySelector('.sp').value,pages.length);if(!plan){toast('Format tidak valid');return;}
      run(async()=>{
        await commit();await lib('pdflib',()=>window.PDFLib);await lib('jszip',()=>window.JSZip);const pr=note('Memisahkan…');
        try{
          const sd=await PDFLib.PDFDocument.load(bytes,{ignoreEncryption:true}),z=new JSZip(),pad=String(plan.length).length;
          for(let k=0;k<plan.length;k++){const o=await PDFLib.PDFDocument.create();(await o.copyPages(sd,plan[k])).forEach(pg=>o.addPage(pg));z.file(`${base()}_bagian${String(k+1).padStart(pad,'0')}.pdf`,await o.save());}
          dl(await z.generateAsync({type:'blob'}),base()+'_pisah.zip');toast(plan.length+' berkas dibuat');
        }finally{pr.end();}
      });
    },
    async hf(v){
      if(!needDoc())return;
      const f=(c,l,val='')=>`<label>${l}<input type="text" class="s-${c}" value="${esc(val)}"></label>`;
      const b=await dialog('Header & Footer',f('hl','Header kiri:')+f('hc','Header tengah:')+f('hr','Header kanan:')+f('fl','Footer kiri:')+f('fc','Footer tengah:',v==='num'?'Halaman {page} dari {total}':'')+f('fr','Footer kanan:')+
        '<label>Ukuran font:<select class="fs"><option>8</option><option>9</option><option selected>10</option><option>11</option><option>12</option><option>14</option></select></label><label>Halaman:<input type="text" class="rg" value="semua"></label><small>{page} = nomor halaman, {total} = jumlah halaman, {date} = tanggal hari ini.</small>','Terapkan');
      if(!b)return;const idx=rangeIn(b);if(!idx)return;
      const slots=[['hl','l','t'],['hc','c','t'],['hr','r','t'],['fl','l','b'],['fc','c','b'],['fr','r','b']].map(([c,h,vv])=>[b.querySelector('.s-'+c).value,h,vv]).filter(x=>x[0].trim());
      if(!slots.length){toast('Isi minimal satu kolom');return;}
      const size=+b.querySelector('.fs').value,tot=pages.length;
      run(()=>edit(async(d,PL)=>{const font=await d.embedFont(PL.StandardFonts.Helvetica);for(const i of idx)for(const[t,h,vv]of slots)drawVis(PL,d.getPage(i),font,U.safe(U.hfText(t,i+1,tot)),{h,v:vv,m:28},size);},cur));
    },
    async wm(v){
      if(!needDoc())return;
      const PRE={draf:'DRAF',rahasia:'RAHASIA',salinan:'SALINAN',jangan:'JANGAN DISALIN'};
      let o={text:PRE[v],size:96,op:25,color:'#cc0000',ang:45},idx=[...Array(pages.length).keys()];
      if(!PRE[v]){
        const b=await dialog('Watermark teks','<label>Teks:<input type="text" class="wt" value="CONFIDENTIAL"></label><label>Ukuran font:<input type="text" class="ws" value="90"></label><label>Opasitas (%):<input type="text" class="wo" value="25"></label><label>Warna:<select class="wc"><option value="#cc0000">Merah</option><option value="#808080">Abu-abu</option><option value="#1a56db">Biru</option><option value="#000000">Hitam</option></select></label><label>Sudut (°):<input type="text" class="wa" value="45"></label><label>Halaman:<input type="text" class="rg" value="semua"></label>','Terapkan');
        if(!b)return;const t=b.querySelector('.wt').value.trim();if(!t){toast('Teks kosong');return;}
        idx=rangeIn(b);if(!idx)return;
        o={text:t,size:Math.max(8,Math.min(300,+b.querySelector('.ws').value||90)),op:Math.max(5,Math.min(100,+b.querySelector('.wo').value||25)),color:b.querySelector('.wc').value,ang:+b.querySelector('.wa').value||0};
      }
      run(()=>edit(async(d,PL)=>{
        const font=await d.embedFont(PL.StandardFonts.HelveticaBold),txt=U.safe(o.text),[cr,cg,cb2]=U.col(o.color),rad=o.ang*Math.PI/180,ca=Math.abs(Math.cos(rad)),sa=Math.abs(Math.sin(rad));
        for(const i of idx){
          const pg=d.getPage(i),cb=pg.getCropBox(),R=U.rot(pg.getRotation().angle),W=cb.width,H=cb.height,vw=R%180?H:W,vh=R%180?W:H;
          const tw0=font.widthOfTextAtSize(txt,o.size),th0=o.size*0.72,k=Math.min(1,0.9*vw/((tw0*ca+th0*sa)||1),0.9*vh/((tw0*sa+th0*ca)||1)),size=o.size*k,tw=tw0*k,th=th0*k;
          const vx=vw/2-(tw/2)*Math.cos(rad)+(th/2)*Math.sin(rad),vy=vh/2-(tw/2)*Math.sin(rad)-(th/2)*Math.cos(rad),[u,w2]=U.v2p(R,W,H,vx,vy);
          pg.drawText(txt,{x:cb.x+u,y:cb.y+w2,size,font,color:PL.rgb(cr,cg,cb2),opacity:o.op/100,rotate:PL.degrees(o.ang+R)});
        }
      },cur));
    },
    async bates(){
      if(!needDoc())return;
      const b=await dialog('Nomor Bates','<label>Awalan:<input type="text" class="bp" value="DOK-"></label><label>Nomor awal:<input type="text" class="bs" value="1"></label><label>Jumlah digit:<input type="text" class="bd" value="6"></label><label>Akhiran:<input type="text" class="bx" value=""></label><label>Posisi:<select class="bpos"><option value="br">Bawah kanan</option><option value="bc">Bawah tengah</option><option value="bl">Bawah kiri</option><option value="tr">Atas kanan</option><option value="tc">Atas tengah</option><option value="tl">Atas kiri</option></select></label><label>Ukuran font:<input type="text" class="bf" value="10"></label><label>Halaman:<input type="text" class="rg" value="semua"></label>','Terapkan');
      if(!b)return;const idx=rangeIn(b);if(!idx)return;
      const q=c=>b.querySelector(c).value,pre=U.safe(q('.bp')),suf=U.safe(q('.bx')),st=Math.max(0,parseInt(q('.bs'))||1),dg=Math.max(1,Math.min(12,parseInt(q('.bd'))||6)),sz=Math.max(6,Math.min(24,+q('.bf')||10)),pos=q('.bpos');
      run(()=>edit(async(d,PL)=>{const font=await d.embedFont(PL.StandardFonts.HelveticaBold);idx.forEach((i,k)=>drawVis(PL,d.getPage(i),font,U.bates(pre,st+k,dg,suf),{h:pos[1],v:pos[0],m:28},sz));},cur));
    },
    async bmadd(){
      if(!needDoc())return;
      const b=await dialog('Tambah bookmark',`<label>Judul:<input type="text" class="bt" value="Halaman ${cur+1}"></label>`,'Tambah');if(!b)return;
      const title=b.querySelector('.bt').value.trim();if(!title){toast('Judul kosong');return;}
      const at=cur;run(()=>edit((d,PL)=>{U.addBookmark(PL,d,title,at);},at));
    },
    async bmlist(){
      if(!needDoc())return;
      const ol=await pdfDoc.getOutline();if(!ol||!ol.length){toast('Dokumen ini belum punya bookmark');return;}
      const flat=[];const walk=(it,dep)=>it.forEach(o=>{flat.push({t:o.title,dep,dest:o.dest});walk(o.items||[],dep+1);});walk(ol,0);
      const idxOf=async dest=>{try{let d=dest;if(typeof d==='string')d=await pdfDoc.getDestination(d);if(!Array.isArray(d))return -1;const r=d[0];return typeof r==='object'?await pdfDoc.getPageIndex(r):+r;}catch(e){return -1;}};
      const rows=[];for(const f of flat)rows.push({...f,p:await idxOf(f.dest)});
      await dialog('Daftar bookmark',rows.map((r,k)=>`<button type="button" class="bmi" data-k="${k}" style="padding-left:${8+r.dep*16}px">${esc(r.t)}${r.p>=0?' — hal. '+(r.p+1):''}</button>`).join(''),'Tutup',body=>{
        body.addEventListener('click',e=>{const k=e.target.closest('.bmi');if(!k)return;const r=rows[+k.dataset.k];if(r.p>=0){goto(r.p);body.closest('.ov').querySelector('.dbt button').click();}});
      });
    },
    /* --- Protect --- */
    redarea(){setTool('redarea');toast('Seret untuk menandai area yang akan dihapus');},
    redtext(){setTool('redtext');toast('Sorot teks yang akan dihapus');},
    redclear(){
      const n=annots.filter(isRd).length;if(!n){toast('Tidak ada tanda redaksi');return;}
      annots=annots.filter(a=>!isRd(a));drawAll();push();toast(n+' tanda redaksi dihapus');
    },
    async redsearch(){
      if(!needDoc())return;
      const b=await dialog('Cari dan redaksi','<label>Teks yang dicari:<input type="text" class="rq"></label><small>Setiap kecocokan hanya DITANDAI dulu. Tinjau, lalu pilih "Apply All Redactions" untuk menghapusnya.</small>','Tandai');
      if(!b)return;const q=b.querySelector('.rq').value;if(!q.trim())return;
      run(async()=>{
        const pr=note('Mencari…');
        try{
          const{out,missed}=await findRects(q);
          out.forEach(r=>addAnn({type:'hl',kind:'rd',page:r.page,rects:r.rects}));
          if(out.length)push();
          const n=out.reduce((t,r)=>t+r.rects.length,0);
          toast(out.length?`${n} area ditandai di ${out.length} halaman`+(missed?` — ${missed} kecocokan tidak bisa ditandai, tandai manual`:''):(missed?`${missed} kecocokan ditemukan tetapi tidak bisa ditandai, tandai manual`:'Teks tidak ditemukan'));
        }finally{pr.end();}
      });
    },
    async redapply(){
      if(!needDoc())return;
      const marks=annots.filter(isRd);if(!marks.length){toast('Belum ada tanda redaksi');return;}
      const pg2=[...new Set(marks.map(a=>a.page))].sort((a,b)=>a-b),nr=marks.reduce((n,a)=>n+a.rects.length,0);
      if(!confirm(`Terapkan ${nr} area redaksi pada ${pg2.length} halaman?\n\nHalaman terkait diubah menjadi gambar dengan area hitam sehingga teks aslinya benar-benar hilang (halaman itu tak bisa lagi dicari/disalin). Anda masih bisa membatalkan dengan Urungkan sebelum menyimpan.`))return;
      run(async()=>{
        const pr=note('Menerapkan redaksi…');
        try{
          await commit();
          const K=2.5,jpgs=new Map();
          for(const i of pg2){
            pr.set(`Merender halaman ${i+1}…`);
            const c=await renderCanvas(i,K),g=c.getContext('2d');g.fillStyle='#000';
            for(const a of marks)if(a.page===i)for(const r of a.rects){const pad=1.5;g.fillRect((r.x-pad)*K,(r.y-pad)*K,(r.w+2*pad)*K,(r.h+2*pad)*K);}
            jpgs.set(i,new Uint8Array(await(await toBlob(c,'image/jpeg',.92)).arrayBuffer()));
          }
          const dims=new Map(pg2.map(i=>[i,[pages[i].w,pages[i].h]]));
          annots=annots.filter(a=>!isRd(a));
          await edit(async(d,PL)=>{for(const i of pg2){const[w,h]=dims.get(i);await U.redactPage(PL,d,i,jpgs.get(i),w,h);}U.gc(PL,d);},cur);
          toast('Redaksi diterapkan');
        }finally{pr.end();}
      });
    },
    rmmeta(){
      if(!needDoc())return;
      run(()=>edit((d,PL)=>{
        const ctx=d.context,ref=ctx.trailerInfo.Info,info=ref&&ctx.lookup(ref,PL.PDFDict);
        if(info)for(const k of [...info.keys()])info.delete(k);
        d.catalog.delete(PL.PDFName.of('Metadata'));
        U.gc(PL,d);
      },cur).then(()=>toast('Metadata dokumen dihapus')));
    },
    async pwd(){
      if(!needDoc()||!redOk())return;
      const ck=(c,l)=>`<label><input type="checkbox" class="${c}" checked>${l}</label>`;
      const b=await dialog('Keamanan kata sandi',`<label>Kata sandi untuk membuka:<input type="password" class="pu" autocomplete="new-password"></label><label>Kata sandi izin (pemilik):<input type="password" class="po" autocomplete="new-password"></label>${ck('ap','Izinkan mencetak')}${ck('ac','Izinkan menyalin teks')}${ck('am','Izinkan mengubah dokumen')}${ck('aa','Izinkan menambah anotasi')}<small>Hasilnya salinan baru (${esc(base())}_terlindungi.pdf); berkas yang terbuka tidak berubah. Kata sandi tidak bisa dipulihkan bila lupa. Bila kata sandi izin dikosongkan dan ada izin yang dimatikan, kata sandi pemilik dibuat acak agar pembatasan tetap berlaku.</small>`,'Lindungi');
      if(!b)return;
      const q=c=>b.querySelector(c),up=q('.pu').value,perms={printing:q('.ap').checked?'highResolution':false,copying:q('.ac').checked,modifying:q('.am').checked,annotating:q('.aa').checked};
      const restr=!(q('.ap').checked&&q('.ac').checked&&q('.am').checked&&q('.aa').checked);
      const op=q('.po').value||(restr?[...crypto.getRandomValues(new Uint8Array(12))].map(x=>x.toString(16).padStart(2,'0')).join(''):up);
      if(!up&&!q('.po').value){toast('Isi minimal satu kata sandi');return;}
      run(async()=>{
        const pr=note('Mengenkripsi…');
        try{
          await commit();const E=await loadEnc(),d=await E.PDFDocument.load(bytes,{ignoreEncryption:true});
          d.encrypt({userPassword:up,ownerPassword:op,permissions:perms});
          dl(new Blob([await d.save()],{type:'application/pdf'}),base()+'_terlindungi.pdf');toast('Salinan terlindungi diunduh');
        }finally{pr.end();}
      });
    }
  });
  Object.assign(acts,{
    talign(v){if(ppApply('align',v))push();},
    ftt(v){if(ppApply('fstyle',v))push();},
    tfinish(){if(curEditDone)curEditDone();},
    delsel(){if(sel.length)delSel();else toast('Tidak ada objek yang dipilih');},
    fillpick(){const c=$('[data-pp=fill]');if(c)c.click();},
    fillnone(){if(ppApply('fill',''))push();},
    szH(){sizeTo(false,true);},szW(){sizeTo(true,false);},szB(){sizeTo(true,true);},
    pdefault(){
      const p=ppTargets()[0],k=p&&(p.type==='tx'?(p.co?'callout':p.bg?'note':p.box?'textbox':'text'):p.type==='sh'?p.shape:null);
      if(!k){toast('Set As Default hanya untuk teks dan bentuk');return;}
      const F={tx:['size','color','font','bold','italic','align','cs','ws','ls','box','bg','bw','dash','op','ll','lo'],sh:['color','fill','lw','dash','op','s','e']}[p.type],o={};
      for(const f of F)if(p[f]!==undefined)o[f]=p[f];
      defs[k]=o;store.set('defs',defs);toast('Gaya disimpan sebagai default untuk '+U.kindLabel(p));
    },
    pinfo(){
      const p=ppTargets()[0];if(!p){toast('Pilih objek terlebih dahulu');return;}
      const bb=boxes().find(b=>b.a===p)||(p.rects?{x:p.rects[0].x,y:p.rects[0].y,w:p.rects[0].w,h:p.rects[0].h}:{x:0,y:0,w:0,h:0}),mm=v=>(v*25.4/72).toFixed(1);
      const row=(k,v)=>`<div style="display:flex;gap:10px;padding:2px 0"><b style="width:90px">${k}</b><span>${esc(String(v))}</span></div>`;
      dialog('Properties',row('Jenis',U.kindLabel(p))+row('Halaman',p.page+1)+row('Posisi',`${bb.x.toFixed(1)}, ${bb.y.toFixed(1)} pt (${mm(bb.x)}, ${mm(bb.y)} mm)`)+row('Ukuran',`${bb.w.toFixed(1)} × ${bb.h.toFixed(1)} pt (${mm(bb.w)} × ${mm(bb.h)} mm)`)+row('Author',p.au||author||'-')+row('Subject',p.sj||U.kindLabel(p)),'Tutup');
    }
  });
  Object.keys(ALN).forEach(k=>{acts[k]=()=>{const bs=boxes();if(bs.length<ALNMIN[k])return;ALN[k](bs);commitBoxes(bs);};});
  function drawFavs(){
    $('.favs').innerHTML=favs.map(id=>{const f=FAVS.find(x=>x[0]===id);if(!f)return'';const[a,v]=id.split(':');return`<button class="b rw" data-a="${a}" data-v="${v??''}" title="${esc(f[1])}">${IC[f[2]](16)}<span>${esc(f[1])}</span></button>`;}).join('');
    $('.favs').style.display=favs.length?'flex':'none';
  }

  /* ---------- cari ---------- */
  const reOf=q=>new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'gi');
  async function doFind(q,keep){
    const my=++ft;fq=q;fm=[];if(!keep)fi=-1;
    if(!q||!pdfDoc){pages.forEach((_,i)=>applyHits(i));fc();return;}
    for(let i=0;i<pages.length;i++){
      const tc=await getText(i);if(my!==ft)return;
      const s=tc.items.map(t=>t.str||'').join(''),re=reOf(q);let m,n=0;
      while((m=re.exec(s))){fm.push({page:i,n:n++});if(m[0].length===0)re.lastIndex++;}
    }
    if(!fm.length){fi=-1;fc();pages.forEach((_,i)=>applyHits(i));return;}
    if(!keep||fi<0||fi>=fm.length){fi=Math.max(0,fm.findIndex(m=>m.page>=cur));}
    fc();pages.forEach((_,i)=>applyHits(i));if(!keep)showMatch();
  }
  const fc=()=>{$('.fc').textContent=fq?(fm.length?`${fi+1}/${fm.length}`:'0/0'):'';};
  function applyHits(i){
    const p=pages[i];if(!p||!p.divs)return;
    p.divs.forEach(d=>d.classList.remove('hit','cur'));
    if(!fq)return;
    const strs=p.strs,offs=[];let pos=0;for(const s of strs){offs.push(pos);pos+=s.length;}
    const text=strs.join(''),re=reOf(fq);let m,n=0;
    while((m=re.exec(text))){
      if(m[0].length===0){re.lastIndex++;continue;}
      const a=m.index,b=a+m[0].length,isCur=fm[fi]&&fm[fi].page===i&&fm[fi].n===n;
      for(let k=0;k<strs.length;k++)if(offs[k]<b&&offs[k]+strs[k].length>a&&p.divs[k]){p.divs[k].classList.add('hit');if(isCur)p.divs[k].classList.add('cur');}
      n++;
    }
  }
  async function showMatch(){
    const m=fm[fi];if(!m)return;
    goto(m.page,false);await renderPage(m.page);applyHits(m.page);
    const e=pgs.children[m.page].querySelector('.hit.cur');if(e)e.scrollIntoView({block:'center',inline:'center'});
  }
  function stepFind(d){if(!fm.length)return;const old=fm[fi]&&fm[fi].page;fi=(fi+d+fm.length)%fm.length;fc();if(old!=null)applyHits(old);showMatch();}
  fqIn.addEventListener('input',()=>{clearTimeout(fqIn._t);fqIn._t=setTimeout(()=>doFind(fqIn.value),250);});
  fqIn.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();if(fqIn.value!==fq)doFind(fqIn.value);else stepFind(e.shiftKey?-1:1);}else if(e.key==='Escape')acts.fclose();});

  /* ---------- event utama ---------- */
  pd.addEventListener('click',e=>{
    const was=ddBtn,dd=e.target.closest('[data-dd]');closeDD();
    if(!e.target.closest('.menu')&&!e.target.closest('[data-a=menu]'))menu.hidden=true;
    if(dd){if(was!==dd)openDD(dd);return;}
    const t=e.target.closest('[data-tab]');
    if(t){if(!t.classList.contains('ctx'))lastMain=t.dataset.tab;showTab(t.dataset.tab);return;}
    const tb=e.target.closest('[data-t]');
    if(tb){setTool(tool===tb.dataset.t&&tb.dataset.t!=='select'?'select':tb.dataset.t);return;}
    const b=e.target.closest('[data-a]');
    if(b&&acts[b.dataset.a])acts[b.dataset.a](b.dataset.v||undefined);
  });
  $('.zs').oninput=e=>{autoFit=false;setZoom(+e.target.value);};
  $('.pi').addEventListener('keydown',e=>{if(e.key==='Enter'){const v=Math.max(1,Math.min(pages.length,parseInt(e.target.value)||1));goto(v-1);}});
  pd.addEventListener('keydown',e=>{
    if(e.key==='Escape'){closeDD();menu.hidden=true;if(!fd.hidden)acts.fclose();else if(!e.target.isContentEditable){if(sel.length)select(null);setTool('select');}return;}
    if(e.target.isContentEditable||e.target.tagName==='INPUT'||e.target.tagName==='SELECT')return;
    if(e.ctrlKey||e.metaKey){
      const k=e.key.toLowerCase();
      if(k==='s'){e.preventDefault();acts.save();}
      else if(k==='f'){e.preventDefault();acts.find();}
      else if(k==='z'){e.preventDefault();e.shiftKey?acts.redo():acts.undo();}
      else if(k==='y'){e.preventDefault();acts.redo();}
      else if(k==='p'){e.preventDefault();acts.print();}
      else if(e.key==='+'||e.key==='='){e.preventDefault();acts.zin();}
      else if(e.key==='-'){e.preventDefault();acts.zout();}
      return;
    }
    if((e.key==='Delete'||e.key==='Backspace')&&sel.length){e.preventDefault();delSel();}
    else if(e.key==='PageDown')acts.next();else if(e.key==='PageUp')acts.prev();
  });
  if(window.ResizeObserver){let rt=0;new ResizeObserver(()=>{clearTimeout(rt);rt=setTimeout(()=>{if(!root.isConnected||!pages.length||!autoFit)return;const v=Math.max(10,Math.min(100,fitW()));if(v!==zoom)setZoom(v,false);},80);}).observe(desk);}
  if(matchMedia('(max-width:700px)').matches){$$('.tab[data-tab]').forEach(x=>x.classList.remove('on'));$$('[data-pane]').forEach(p=>p.hidden=true);}

  drawFavs();setTool('select');btns();
  (async()=>{
    try{
      const last=D.list.find(x=>x.id===D.active)||D.list[D.list.length-1];
      if(last)await enter(last); /* pulihkan berkas yang terbuka sebelum kembali ke menu utama */
      else D.active=null;
      if(opts.file)await openFile(opts.file);else await blankDoc();
    }catch(e){console.error(e);toast('Gagal memuat pustaka PDF: '+(e&&e.message||e)+'. Periksa koneksi internet.');}
    pd.focus();
  })();
}

window.Apps=window.Apps||{};
Apps.pdf={mount,util:U};
})();
