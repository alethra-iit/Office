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
  pptx:'https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundle.js'
};
const libs={};
function lib(n,test){
  if(test()) return Promise.resolve();
  return libs[n]||(libs[n]=new Promise((ok,err)=>{const s=document.createElement('script');s.src=CDN[n];s.onload=ok;s.onerror=()=>{delete libs[n];err(new Error('Gagal memuat '+n));};document.head.appendChild(s);}));
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
/* tempelkan anotasi ke dokumen pdf-lib. pages[i]={vp,rot} */
U.flatten=async(PL,doc,annots,pages)=>{
  const {StandardFonts,rgb,degrees,BlendMode}=PL;
  const font=await doc.embedFont(StandardFonts.Helvetica),pgs=doc.getPages();
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
    if(a.type==='tx'){
      const m=/^#?([0-9a-f]{6})$/i.exec(a.color||'#000000'),c=m?[0,2,4].map(i=>parseInt(m[1].slice(i,i+2),16)/255):[0,0,0];
      U.safe(a.text).split('\n').forEach((ln,i)=>{
        if(!ln) return;
        const [px,py]=vp.convertToPdfPoint(a.x,a.y+i*a.size*1.2+a.size*TXT_ASC);
        pg.drawText(ln,{x:px,y:py,size:a.size,font,color:rgb(c[0],c[1],c[2]),rotate:degrees(rot)});
      });
    }else if(a.type==='hl'){
      for(const r of a.rects){
        const [x1,y1]=vp.convertToPdfPoint(r.x,r.y),[x2,y2]=vp.convertToPdfPoint(r.x+r.w,r.y+r.h);
        pg.drawRectangle({x:Math.min(x1,x2),y:Math.min(y1,y2),width:Math.abs(x2-x1),height:Math.abs(y2-y1),color:rgb(1,0.9,0),borderWidth:0,blendMode:BlendMode.Multiply});
      }
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
      pg.drawImage(img,{x:px,y:py,width:a.w,height:a.h,rotate:degrees(rot)});
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
  tools:s=>svg('<path d="M4 12h10M9 7v10" stroke="#e5622c" stroke-width="3.5"/><path d="M19 29l8-14a5 5 0 1 0-4-4z" fill="#555" stroke="#555"/>',s)
};

/* ====================== ribbon ====================== */
const G=(n,b)=>`<div class="grp"><div class="gb">${b}</div><div class="gn">${n}</div></div>`;
const lab=(l,dd)=>{const t=l.replace(' ','<br>');return dd?(l.includes(' ')?t+' ▾':t+'<br>▾'):t;};
const big=(a,l,i,o={})=>`<button class="b big${o.t?'':''}" ${o.dd?`data-dd="${a}"`:o.t?`data-t="${a}"`:`data-a="${a}"`} title="${l}">${IC[i](32)}<span>${lab(l,o.dd)}</span></button>`;
const sm=(a,l,i,o={})=>`<button class="b rw" ${o.dd?`data-dd="${a}"`:o.t?`data-t="${a}"`:`data-a="${a}"`} title="${l}">${IC[i](16)}<span>${l}</span>${o.dd?'<i>▾</i>':''}</button>`;
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
 G('Create',big('portfolio','Portfolio','portfolio'))
};
const DD={
  create:()=>[['PDF kosong baru','blank'],['PDF dari gambar…','fromimg'],['PDF dari teks (.txt)…','fromtxt']],
  sign:sg=>[['Gambar tanda tangan baru…','signnew'],...(sg?[['Pakai tanda tangan tersimpan','signuse'],['Hapus tanda tangan tersimpan','signdel']]:[])],
  insert:()=>[['Halaman kosong (setelah halaman ini)','insblank'],['Dari berkas PDF…','inspdf'],['Dari gambar…','insimg']],
  rotate:()=>[['Putar kanan 90° (halaman ini)','rot:90'],['Putar kiri 90° (halaman ini)','rot:-90'],['Putar 180° (halaman ini)','rot:180'],'-',['Putar semua halaman ke kanan','rotall:90'],['Putar semua halaman ke kiri','rotall:-90']],
  zoom:()=>[['Perbesar','zin'],['Perkecil','zout'],'-',['50%','z:50'],['75%','z:75'],['100%','z:100'],['150%','z:150'],['200%','z:200'],'-',['Sesuaikan lebar','zfitw'],['Sesuaikan halaman','zfitp']],
  pdfa:()=>[['PDF/A-1b','pdfa'],['PDF/A-2b','pdfa'],['PDF/A-3b','pdfa']]
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
.pd .im.sel .rs{display:block}
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
.pd .db input[type=text],.pd .db select{height:26px;border:1px solid #aaa;padding:0 4px;max-width:100%}
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
  const textCache=new Map(),ZL=[10,25,50,75,100,125,150,200,300,400,500];

  root.innerHTML=`<div class="pd" tabindex="-1">
    <div class="qat"><div class="l"><span class="w">A</span><button data-a="save" title="Simpan (Ctrl+S)">💾</button><button data-a="undo" title="Urungkan (Ctrl+Z)">↶</button><button data-a="redo" title="Ulangi (Ctrl+Y)">↷</button></div><span class="fn"></span></div>
    <div class="tabs"><button class="tab file" data-a="menu">File</button><button class="tab on" data-tab="home">Home</button><button class="tab" data-tab="convert">Convert</button></div>
    <div class="rib">${MODES}<div class="ribbon" data-pane="home">${RIB.home}</div><div class="ribbon" data-pane="convert" hidden>${RIB.convert}</div></div>
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
    bytes=s.bytes;annots=clone(s.annots);selId=null;dirty=true;btns();
    if(reload)await load(bytes,{page:cur});else drawAll();
  }
  async function apply(nb,page){bytes=nb;annots=[];selId=null;push();await load(nb,{page});}
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
    const d=document.createElement('i');d.dataset.id=a.id;
    d.style.cssText=`left:calc(var(--s)*${r.x}px);top:calc(var(--s)*${r.y}px);width:calc(var(--s)*${r.w}px);height:calc(var(--s)*${r.h}px)`;
    return d;
  }
  function mkAnn(a){
    const d=document.createElement('div');d.dataset.id=a.id;
    if(a.type==='tx'){d.className='an tx';d.style.cssText=`left:${a.x}px;top:${a.y}px;font-size:${a.size}px;color:${a.color||'#000'}`;d.textContent=a.text;}
    else if(a.type==='im'){d.className='an im';d.style.cssText=`left:${a.x}px;top:${a.y}px;width:${a.w}px;height:${a.h}px`;d.innerHTML=`<img src="${a.src}" alt=""><span class="rs"></span>`;}
    else if(a.type==='tb'){
      d.className='an tb '+(a.fam||'serif');
      d.style.cssText=`left:${a.x}px;top:${a.y}px;width:${a.w}px;min-height:${a.h}px;font-size:${a.size}px;line-height:${a.lh};text-align:${a.align};color:${a.color||'#000'}`;
      d.innerHTML=`<div class="tbt" style="text-indent:${a.indent||0}px">${runsHtml(a.runs)}</div>${MV}${HD}`;
    }
    if(a.id===selId)d.classList.add('sel');
    return d;
  }
  function dropUntouched(id){
    const a=annots.find(x=>x.id===id);if(!a||a.type!=='tb'||a.touched)return;
    annots=annots.filter(x=>x!==a);annEls(id).forEach(e=>e.remove());
    pgs.querySelectorAll(`[data-for="${id}"]`).forEach(e=>e.remove());
  }
  function select(id){
    if(selId&&selId!==id)dropUntouched(selId);
    selId=id;pgs.querySelectorAll('.sel').forEach(e=>e.classList.remove('sel'));const e=id&&annEl(id);if(e)e.classList.add('sel');
  }
  function setTool(t){
    tool=t;desk.className='desk t-'+t;
    $$('[data-t]').forEach(b=>b.classList.toggle('on',b.dataset.t===t));
    if(t!=='edit'&&t!=='text')select(null);
    if(t!=='edit')hideHv();
    if(t!=='highlight'&&t!=='select')getSelection().removeAllRanges();
  }
  function startEdit(el){
    el.contentEditable='true';el.focus();
    const r=document.createRange();r.selectNodeContents(el);const s=getSelection();s.removeAllRanges();s.addRange(r);
    const id=el.dataset.id;
    const done=()=>{
      el.removeEventListener('blur',done);el.contentEditable='false';
      const a=annots.find(x=>x.id===id);if(!a)return;
      const t=el.innerText.replace(/\n+$/,'');
      if(!t.trim()){annots=annots.filter(x=>x.id!==id);el.remove();if(a.text)push();return;}
      if(t!==a.text){a.text=t;push();}
    };
    el.addEventListener('blur',done);
    el.onkeydown=e=>{if(e.key==='Escape'){e.stopPropagation();el.blur();}};
  }
  function addAnn(a){
    a.id='a'+(++idc);annots.push(a);
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
    if(tool==='text'){
      const tx=e.target.closest('.an.tx');if(tx){if(!tx.isContentEditable)startEdit(tx);return;}
      if(e.target.closest('.an'))return;
      const[x,y]=pt(e,pgEl),size=12;
      const el=addAnn({type:'tx',page:i,x:Math.max(0,x),y:Math.max(0,y-size*0.6),size,color:'#000000',text:''});
      select(null);startEdit(el);
    }else if(tool==='sign'){
      if(!sign){const s=await signDlg();if(!s)return;sign=s;store.set('sign',s);}
      placeSign(i,pt(e,pgEl));
    }else if(tool==='edit'){
      if(e.target.closest('.an'))return;
      openBlockAt(e,pgEl);
    }else if(tool==='highlight'){
      const s=getSelection();if(s&&!s.isCollapsed)return;
      const[x,y]=pt(e,pgEl),h=annots.find(z=>z.type==='hl'&&z.page===i&&z.rects.some(r=>x>=r.x&&x<=r.x+r.w&&y>=r.y&&y<=r.y+r.h));
      if(h){annots=annots.filter(z=>z!==h);annEls(h.id).forEach(el=>el.remove());push();toast('Highlight dihapus');}
    }
  });
  function placeSign(i,[x,y]){
    const im=new Image();
    im.onload=()=>{const w=150,h=w*im.height/im.width;
      const el=addAnn({type:'im',page:i,x:Math.max(0,x-w/2),y:Math.max(0,y-h/2),w,h,src:sign});
      push();setTool('edit');select(el.dataset.id);};
    im.src=sign;
  }
  /* geser / ubah ukuran / hapus anotasi (alat Edit) */
  desk.addEventListener('pointerdown',e=>{
    if(tool==='hand'){
      const sx=e.clientX,sy=e.clientY,l=desk.scrollLeft,t=desk.scrollTop;desk.classList.add('drag');desk.setPointerCapture(e.pointerId);
      const mv=m=>{desk.scrollLeft=l-(m.clientX-sx);desk.scrollTop=t-(m.clientY-sy);};
      const up=()=>{desk.classList.remove('drag');desk.removeEventListener('pointermove',mv);desk.removeEventListener('pointerup',up);desk.removeEventListener('pointercancel',up);};
      desk.addEventListener('pointermove',mv);desk.addEventListener('pointerup',up);desk.addEventListener('pointercancel',up);return;
    }
    if(tool!=='edit')return;
    const an=e.target.closest('.an');
    if(!an){select(null);return;}
    const a=annots.find(x=>x.id===an.dataset.id);if(!a)return;
    if(a.type==='tb'){tbPointer(e,an,a);return;}
    select(a.id);
    if(an.isContentEditable)return;
    e.preventDefault();
    const rs=e.target.classList.contains('rs'),sx=e.clientX,sy=e.clientY,ox=a.x,oy=a.y,ow=a.w,oh=a.h;let moved=false;
    an.setPointerCapture&&an.setPointerCapture(e.pointerId);
    const mv=m=>{const dx=(m.clientX-sx)/S,dy=(m.clientY-sy)/S;moved=true;
      if(rs&&a.type==='im'){a.w=Math.max(20,ow+dx);a.h=a.w*oh/ow;an.style.width=a.w+'px';an.style.height=a.h+'px';}
      else{a.x=ox+dx;a.y=oy+dy;an.style.left=a.x+'px';an.style.top=a.y+'px';}};
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
  desk.addEventListener('dblclick',e=>{const tx=e.target.closest('.an.tx');if(tool==='edit'&&tx)startEdit(tx);});
  /* sorot teks */
  desk.addEventListener('pointerup',()=>{if(tool==='highlight')setTimeout(highlightSel,0);});
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
  function highlightSel(){
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
    by.forEach((rs,i)=>addAnn({type:'hl',page:i,rects:rs}));
    push();s.removeAllRanges();
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
  async function commit(){
    if(!annots.length)return;
    await lib('pdflib',()=>window.PDFLib);
    const doc=await PDFLib.PDFDocument.load(bytes,{ignoreEncryption:true});
    await U.flatten(PDFLib,doc,annots,pages);
    const nb=await doc.save();pristine=false;await apply(nb,cur);
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
    cur=d.cur;zoom=d.zoom;autoFit=d.autoFit;selId=null;D.active=d.id;
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
    if(!needDoc())return;
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
    async save(){if(!needDoc())return;await run(async()=>{await commit();dl(new Blob([bytes],{type:'application/pdf'}),/\.pdf$/i.test(name)?name:name+'.pdf');dirty=false;toast('Tersimpan: '+name);});},
    undo(){run(()=>restore(hi-1));},redo(){run(()=>restore(hi+1));},
    menu(){menu.hidden=!menu.hidden;if(!menu.hidden)renderList();},
    home(){snap();opts.onHome&&opts.onHome();},
    sw(id){if(id===D.active){menu.hidden=true;return;}run(async()=>{snap();menu.hidden=true;await enter(D.list.find(x=>x.id===id));});},
    cl(id){run(()=>closeDoc(id));},
    blank(){menu.hidden=true;run(blankDoc);},
    async open(){const fs=await pickFiles('.pdf,application/pdf',true);if(!fs.length)return;menu.hidden=true;run(async()=>{for(const f of fs)await openFile(f);});},
    async print(){if(!needDoc())return;await run(async()=>{
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
  root.addEventListener('click',e=>{
    const was=ddBtn,dd=e.target.closest('[data-dd]');closeDD();
    if(!e.target.closest('.menu')&&!e.target.closest('[data-a=menu]'))menu.hidden=true;
    if(dd){if(was!==dd)openDD(dd);return;}
    const t=e.target.closest('[data-tab]');
    if(t){$$('.tab[data-tab]').forEach(x=>x.classList.toggle('on',x===t));$$('[data-pane]').forEach(p=>p.hidden=p.dataset.pane!==t.dataset.tab);return;}
    const tb=e.target.closest('[data-t]');
    if(tb){setTool(tool===tb.dataset.t&&tb.dataset.t!=='select'?'select':tb.dataset.t);return;}
    const b=e.target.closest('[data-a]');
    if(b&&acts[b.dataset.a])acts[b.dataset.a](b.dataset.v||undefined);
  });
  $('.zs').oninput=e=>{autoFit=false;setZoom(+e.target.value);};
  $('.pi').addEventListener('keydown',e=>{if(e.key==='Enter'){const v=Math.max(1,Math.min(pages.length,parseInt(e.target.value)||1));goto(v-1);}});
  pd.addEventListener('keydown',e=>{
    if(e.key==='Escape'){closeDD();menu.hidden=true;if(!fd.hidden)acts.fclose();else if(!e.target.isContentEditable)setTool('select');return;}
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
    if((e.key==='Delete'||e.key==='Backspace')&&selId){
      e.preventDefault();const sa=annots.find(a=>a.id===selId);
      if(sa&&sa.type==='tb'){sa.runs=[];sa.touched=true;const old=annEl(sa.id);if(old)old.replaceWith(mkAnn(sa));push();return;}
      annots=annots.filter(a=>a.id!==selId);annEls(selId).forEach(el=>el.remove());selId=null;push();}
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
