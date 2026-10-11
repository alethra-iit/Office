/* Baca/tulis .pptx untuk editor PowerPoint.
   - write(data,PptxGenJS,JSZip): model slide -> berkas .pptx sungguhan (teks, bentuk, gambar, catatan, bagian, latar).
   - read(buf,JSZip): berkas .pptx -> model slide (placeholder, teks berformat, bentuk, garis, gambar, grup, catatan, bagian, latar).
   Mengekspos window.PptxIO. Model: lihat mk()/mkSlide() di ppt.js. Kanvas 1280x720 px = 13,333x7,5 inci (96 dpi). */
(function(root){
'use strict';
const W=1280,H=720,EMU_PX=9525;
const NS_R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const SHAPE_OUT={rect:'rect',round:'roundRect',oval:'ellipse',tri:'triangle',rtri:'rtTriangle',diamond:'diamond',pent:'pentagon',hex:'hexagon',star:'star5',
  rarrow:'rightArrow',larrow:'leftArrow',uarrow:'upArrow',darrow:'downArrow',chevron:'chevron',plus:'mathPlus'};
const SHAPE_IN={rect:'rect',roundRect:'round',round1Rect:'round',round2SameRect:'round',round2DiagRect:'round',snip1Rect:'rect',snip2SameRect:'rect',snip2DiagRect:'rect',
  ellipse:'oval',triangle:'tri',rtTriangle:'rtri',diamond:'diamond',pentagon:'pent',hexagon:'hex',star5:'star',rightArrow:'rarrow',leftArrow:'larrow',upArrow:'uarrow',
  downArrow:'darrow',chevron:'chevron',mathPlus:'plus',plus:'plus',line:'line',straightConnector1:'line',flowChartProcess:'rect',flowChartAlternateProcess:'round',
  flowChartDecision:'diamond',flowChartTerminator:'round',homePlate:'chevron',textBox:'rect'};
const MIME={png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',gif:'image/gif',bmp:'image/bmp',svg:'image/svg+xml',webp:'image/webp'};
const esc=t=>String(t).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const num=(v,d)=>{const n=parseFloat(v);return isNaN(n)?d:n;};

/* ---------- warna ---------- */
const hex6=c=>{
  if(!c)return null;c=String(c).trim().toLowerCase();
  let m=/^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?\s*\)$/.exec(c);
  if(m){if(m[4]!==undefined&&+m[4]===0)return null;return[m[1],m[2],m[3]].map(x=>Math.min(255,+x).toString(16).padStart(2,'0')).join('').toUpperCase();}
  m=/^#([0-9a-f]{3})$/.exec(c);if(m)return m[1].split('').map(x=>x+x).join('').toUpperCase();
  m=/^#([0-9a-f]{6})$/.exec(c);if(m)return m[1].toUpperCase();
  return({black:'000000',white:'FFFFFF',red:'FF0000',blue:'0000FF',green:'008000',yellow:'FFFF00',gray:'808080',grey:'808080',orange:'FFA500',purple:'800080'})[c]||null;
};
const rgb=h=>[0,2,4].map(i=>parseInt(h.slice(i,i+2),16));
const toHex=a=>a.map(v=>Math.max(0,Math.min(255,Math.round(v))).toString(16).padStart(2,'0')).join('').toUpperCase();
function lumShift(h,mod,off){   // lumMod/lumOff pada HSL
  let [r,g,b]=rgb(h).map(v=>v/255);const mx=Math.max(r,g,b),mn=Math.min(r,g,b);let l=(mx+mn)/2,s=0,hh=0;
  if(mx!==mn){const d=mx-mn;s=l>.5?d/(2-mx-mn):d/(mx+mn);hh=mx===r?(g-b)/d+(g<b?6:0):mx===g?(b-r)/d+2:(r-g)/d+4;hh/=6;}
  l=Math.max(0,Math.min(1,l*mod+off));
  const f=(p,q,t)=>{if(t<0)t+=1;if(t>1)t-=1;return t<1/6?p+(q-p)*6*t:t<1/2?q:t<2/3?p+(q-p)*(2/3-t)*6:p;};
  if(!s)return toHex([l*255,l*255,l*255]);
  const q=l<.5?l*(1+s):l+s-l*s,p=2*l-q;return toHex([f(p,q,hh+1/3)*255,f(p,q,hh)*255,f(p,q,hh-1/3)*255]);
}

/* ======================================================================
   TULIS: model -> .pptx
   ====================================================================== */
function parseParas(html,o){
  const D={font:o.ts.font||'Calibri',size:+o.ts.size||18,color:hex6(o.ts.color)||'000000',align:o.ts.align||'left'};
  const box=document.createElement('div');box.innerHTML=html||'';
  const paras=[];let cur=null;
  const flush=()=>{if(cur&&cur.runs.length)paras.push(cur);cur=null;};
  const fmtOf=el=>{
    const t=el.tagName.toLowerCase(),f={},s=el.style||{};
    if(t==='b'||t==='strong')f.bold=1;if(t==='i'||t==='em')f.italic=1;if(t==='u')f.underline=1;if(t==='s'||t==='strike'||t==='del')f.strike=1;
    if(t==='a'&&/^https?:\/\//i.test(el.getAttribute('href')||''))f.link=el.getAttribute('href');
    if(t==='font'){if(el.getAttribute('face'))f.font=el.getAttribute('face');const c=hex6(el.getAttribute('color'));if(c)f.color=c;}
    if(s.fontWeight)f.bold=s.fontWeight==='bold'||+s.fontWeight>=600?1:0;
    if(s.fontStyle)f.italic=s.fontStyle==='italic'?1:0;
    const td=(s.textDecorationLine||s.textDecoration||'');if(/underline/.test(td))f.underline=1;if(/line-through/.test(td))f.strike=1;
    if(s.fontFamily)f.font=s.fontFamily.split(',')[0].replace(/["']/g,'').trim();
    if(s.fontSize){const m=/^([\d.]+)(pt|px)$/.exec(s.fontSize);if(m)f.size=m[2]==='pt'?+m[1]:+m[1]*.75;}
    if(s.color){const c=hex6(s.color);if(c)f.color=c;}
    if(s.letterSpacing){const m=/^(-?[\d.]+)px$/.exec(s.letterSpacing);if(m)f.spc=+m[1]*.75;}
    return f;
  };
  const BLOCK=/^(div|p|li|h[1-6]|blockquote|pre)$/,LIST=/^(ul|ol)$/;
  function walk(node,fmt,pc){
    node.childNodes.forEach(n=>{
      if(n.nodeType===3){
        const t=n.data.replace(/\u00a0/g,' ').replace(/[\r\n]+/g,' ');if(!t)return;
        if(!cur)cur={runs:[],bul:pc.bul,lvl:pc.lvl,align:pc.align};cur.runs.push({t,f:fmt});return;}
      if(n.nodeType!==1)return;
      const tag=n.tagName.toLowerCase();
      if(tag==='br'){flush();return;}
      if(tag==='img'||tag==='script'||tag==='style')return;
      if(LIST.test(tag)){flush();walk(n,fmt,{bul:tag==='ul'?'b':'n',lvl:pc.ld||0,ld:(pc.ld||0)+1,align:pc.align});flush();return;}
      if(BLOCK.test(tag)){
        flush();const k=paras.length,al=(n.style&&n.style.textAlign)||n.getAttribute('align')||pc.align;
        const pc2=Object.assign({},pc,{align:al&&al!=='start'?al:pc.align});
        walk(n,Object.assign({},fmt,fmtOf(n)),pc2);flush();
        if(paras.length===k)paras.push({runs:[],bul:pc2.bul,lvl:pc2.lvl,align:pc2.align});
        return;}
      walk(n,Object.assign({},fmt,fmtOf(n)),pc);
    });
  }
  walk(box,{},{bul:null,lvl:0,ld:0,align:null});flush();
  return{D,paras};
}

function textItems(o){
  const {D,paras}=parseParas(o.html,o),lh=+o.ts.lh||1.1,items=[];
  const AL={left:'left',center:'center',right:'right',justify:'justify',start:'left',end:'right'};
  const list=paras.length?paras:[{runs:[],bul:null,lvl:0,align:null}];
  list.forEach(p=>{
    const po={align:AL[p.align||D.align]||'left',lineSpacingMultiple:lh};
    if(p.bul)po.indentLevel=Math.min(8,p.lvl||0);
    const runs=p.runs.length?p.runs:[{t:'',f:{}}];
    runs.forEach((r,i)=>{
      const f=r.f,op=Object.assign({},po,{fontFace:f.font||D.font,fontSize:Math.max(1,+(f.size||D.size).toFixed(1)),color:f.color||D.color,breakLine:i===runs.length-1});
      if(p.bul)op.bullet=p.bul==='n'?{type:'number'}:true;
      if(f.link)op.hyperlink={url:f.link};if(f.bold)op.bold=true;if(f.italic)op.italic=true;if(f.underline)op.underline={style:'sng'};if(f.strike)op.strike='sngStrike';if(f.spc)op.charSpacing=f.spc;
      items.push({text:r.t,options:op});
    });
  });
  return{items,D};
}

const hasText=h=>{const d=document.createElement('div');d.innerHTML=h||'';return!!d.textContent.trim();};
const inch=px=>px/96;

async function write(data,PptxGenJS,JSZip,type){
  const px=new PptxGenJS();px.layout='LAYOUT_WIDE';px.title=(data.name||'Presentasi').replace(/\.pptx$/i,'');
  const slides=data.slides,secs=slides.some(s=>s.sec!=null);
  const secOf=[];
  if(secs){const first=slides[0].sec!=null?slides[0].sec:'Bagian Default';let cur=first;
    slides.forEach((s,i)=>{if(i===0||s.sec!=null){cur=s.sec!=null?s.sec:first;px.addSection({title:cur});}secOf[i]=cur;});}
  const cols=[];                                   // per slide: daftar [indeks p:sp -> jumlah kolom]
  let nph=0;
  slides.forEach((s,si)=>{
    const sl=secs?px.addSlide({sectionTitle:secOf[si]}):px.addSlide(),spCols=[];let spIdx=0;
    if(s.bg){const c=hex6(s.bg);if(c)sl.background={color:c};}
    s.objs.forEach(o=>{
      const line=o.shape==='line'||o.shape==='arrow',rot=o.rot||0,common={rotate:rot||undefined,flipH:o.fh?true:undefined,flipV:o.fv?true:undefined};
      if(o.media)return;
      if(/<table/i.test(o.html||'')){                       // tabel (sisipan dari menu SISIPKAN)
        const d=document.createElement('div');d.innerHTML=o.html;const tb=d.querySelector('table');
        if(tb&&tb.rows.length){
          const rows=[...tb.rows].map(r=>[...r.cells].map(c=>{const st=c.style,op={fontFace:o.ts.font||'Calibri',fontSize:+o.ts.size||18,color:hex6(st.color)||hex6(o.ts.color)||'000000',valign:'top',border:{type:'solid',pt:1,color:'FFFFFF'}};
            const bg=hex6(st.backgroundColor||st.background);if(bg)op.fill={color:bg};
            if(st.fontWeight==='bold'||+st.fontWeight>=600)op.bold=true;
            return{text:c.textContent.replace(/\u00a0/g,' ').trim(),options:op};}));
          const nc=Math.max(...rows.map(r=>r.length));rows.forEach(r=>{while(r.length<nc)r.push({text:'',options:{}});});
          sl.addTable(rows,{x:inch(o.x),y:inch(o.y),w:inch(o.w),colW:Array(nc).fill(inch(o.w)/nc),rowH:Array(rows.length).fill(inch(o.h)/rows.length)});}
        return;}
      if(o.img){
        if(!/^data:/.test(o.img))return;                    // gambar dari URL (tanpa data) dilewati
        sl.addImage(Object.assign({data:o.img.replace(/^data:/,''),x:inch(o.x),y:inch(o.y),w:inch(o.w),h:inch(o.h)},common));return;}
      if(line){
        const st=hex6(o.stroke)||'5B9BD5',lo={color:st,width:+o.sw||2.25};
        if(o.shape==='arrow')lo.endArrowType='triangle';if(o.dash)lo.dashType=o.dash==='2 4'?'sysDot':'dash';
        sl.addShape('line',Object.assign({x:inch(o.x),y:inch(o.y+o.h/2),w:inch(o.w),h:0,line:lo},common));spIdx++;return;}
      const ph=o.k==='ph',boxy=!!(o.fill||o.stroke||o.k==='sh');
      let name;if(ph){nph++;name=(o.role==='title'?'Title ':o.role==='sub'?'Subtitle ':'Content Placeholder ')+nph;}
      const {items}=textItems(o);
      const opts=Object.assign({x:inch(o.x),y:inch(o.y),w:inch(o.w),h:inch(o.h),margin:[7.2,7.2,3.6,3.6],
        valign:{t:'top',m:'middle',b:'bottom'}[o.anchor]||'top',wrap:true},common);
      if(name)opts.objectName=name;
      if(boxy){opts.shape=SHAPE_OUT[o.shape]||'rect';
        if(o.fill)opts.fill={color:hex6(o.fill)||'FFFFFF'};
        opts.line=o.stroke?{color:hex6(o.stroke)||'000000',width:+o.sw||1,dashType:o.dash?(o.dash==='2 4'?'sysDot':'dash'):'solid'}:{type:'none'};}
      else if(ph)opts.line={type:'none'};
      else opts.isTextBox=true;
      if(o.vert&&o.vert!=='h')opts.vert={v90:'vert',v270:'vert270',st:'wordArtVert'}[o.vert];
      if(o.fx==='shadow')opts.shadow={type:'outer',blur:4,offset:3,angle:45,color:'000000',opacity:.45};
      if(!hasText(o.html)&&!ph&&!boxy)return;
      sl.addText(items,opts);
      if(o.cols>1)spCols[spIdx]=o.cols;
      spIdx++;
    });
    if(s.notes){const d=document.createElement('div');d.innerHTML=String(s.notes).replace(/<\/(div|p|li)>|<br\s*\/?>/gi,'\n');sl.addNotes(d.textContent.replace(/\n+$/,''));}
    cols.push(spCols);
  });
  const buf=await px.write({outputType:'uint8array'});
  return fixup(buf,cols,JSZip,type);
}

/* perbaikan paska-proses: pptxgenjs menulis <a:pPr> sebelum tiap run (tidak valid) + kolom teks */
async function fixup(buf,cols,JSZip,type){
  const z=await JSZip.loadAsync(buf);
  for(let i=0;i<cols.length;i++){
    const p='ppt/slides/slide'+(i+1)+'.xml',f=z.file(p);if(!f)continue;
    let x=await f.async('string'),k=0;
    x=x.replace(/<p:sp>[\s\S]*?<\/p:sp>/g,sp=>{const c=cols[i][k++];return c?sp.replace(/<a:bodyPr /,'<a:bodyPr numCol="'+c+'" spcCol="228600" '):sp;});
    x=x.replace(/<a:p>[\s\S]*?<\/a:p>/g,para=>{let seen=false;return para.replace(/<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g,m=>seen?'':(seen=true,m));});
    z.file(p,x);
  }
  /* catatan pembicara: pptxgenjs menaruh semua baris dalam satu <a:t>; pecah jadi paragraf sungguhan */
  for(const n of Object.keys(z.files).filter(n=>/^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(n))){
    const x=(await z.file(n).async('string')).replace(/<a:p><a:r>(<a:rPr[^>]*\/>)<a:t>([^<]*[\r\n][^<]*)<\/a:t><\/a:r>(?:<a:endParaRPr[^>]*\/>)?<\/a:p>/,
      (m,rpr,t)=>t.split(/\r?\n/).map(l=>'<a:p><a:r>'+rpr+'<a:t>'+l+'</a:t></a:r></a:p>').join(''));
    z.file(n,x);}
  return z.generateAsync({type:type||'blob',mimeType:'application/vnd.openxmlformats-officedocument.presentationml.presentation',compression:'DEFLATE'});
}

/* ======================================================================
   BACA: .pptx -> model
   ====================================================================== */
async function read(buf,JSZip){
  const zip=await JSZip.loadAsync(buf),DP=new DOMParser(),warn={table:0,chart:0,media:0,other:0};
  const cache={};
  const parse=async p=>{if(p in cache)return cache[p];const f=zip.file(p);let d=null;
    if(f){d=DP.parseFromString(await f.async('string'),'application/xml');if(d.getElementsByTagName('parsererror').length)d=null;}return cache[p]=d;};
  const kids=(e,n)=>e?[...e.childNodes].filter(c=>c.nodeType===1&&c.localName===n):[];
  const kid=(e,n)=>kids(e,n)[0]||null;
  const path=(e,...ns)=>ns.reduce((x,n)=>x&&kid(x,n),e);
  const at=(e,n)=>e?e.getAttribute(n):null;
  const resolve=(base,t)=>{if(t[0]==='/')return t.slice(1);const parts=base.split('/').slice(0,-1);t.split('/').forEach(s=>{if(s==='..')parts.pop();else if(s&&s!=='.')parts.push(s);});return parts.join('/');};
  const relsOf=async part=>{const i=part.lastIndexOf('/'),rp=part.slice(0,i+1)+'_rels/'+part.slice(i+1)+'.rels',d=await parse(rp),m={};
    if(d)[...d.getElementsByTagName('Relationship')].forEach(r=>{m[r.getAttribute('Id')]={type:r.getAttribute('Type')||'',target:resolve(part,r.getAttribute('Target')||'')};});return m;};
  const rid=(e,n)=>e&&(e.getAttributeNS(NS_R,n)||e.getAttribute('r:'+n));

  const pres=await parse('ppt/presentation.xml');if(!pres)throw new Error('presentation.xml tidak ditemukan — bukan berkas .pptx yang valid');
  const prels=await relsOf('ppt/presentation.xml');
  const sz=path(pres.documentElement,'sldSz'),cx0=num(at(sz,'cx'),12192000),cy0=num(at(sz,'cy'),6858000);
  const K=Math.min(W/cx0,H/cy0),offX=(W-cx0*K)/2,offY=(H-cy0*K)/2,FS=K*EMU_PX;   // FS: skala ukuran font
  const defTextStyle=path(pres.documentElement,'defaultTextStyle');

  /* tema + peta warna dari master pertama yang dipakai */
  const themes={};
  async function themeOf(master){
    const mr=await relsOf(master),t=Object.values(mr).find(r=>/\/theme$/.test(r.type));if(!t)return{clr:{},major:'Calibri Light',minor:'Calibri'};
    if(themes[t.target])return themes[t.target];
    const d=await parse(t.target),th={clr:{},major:'Calibri Light',minor:'Calibri'};
    if(d){const cs=d.getElementsByTagNameNS('*','clrScheme')[0];
      if(cs)[...cs.childNodes].filter(n=>n.nodeType===1).forEach(n=>{const c=n.firstElementChild;if(c){let v=(c.localName==='sysClr'?at(c,'lastClr')||({window:'FFFFFF',windowText:'000000'})[at(c,'val')]:at(c,'val'))||'000000';th.clr[n.localName]=/^[0-9a-fA-F]{6}$/.test(v)?v.toUpperCase():'000000';}});
      const mj=d.getElementsByTagNameNS('*','majorFont')[0],mn=d.getElementsByTagNameNS('*','minorFont')[0];
      if(mj&&kid(mj,'latin'))th.major=at(kid(mj,'latin'),'typeface');if(mn&&kid(mn,'latin'))th.minor=at(kid(mn,'latin'),'typeface');}
    return themes[t.target]=th;
  }

  /* konteks per slide: layout, master, tema */
  const ctxCache={};
  async function ctxFor(slidePath){
    const sr=await relsOf(slidePath),lp=Object.values(sr).find(r=>/\/slideLayout$/.test(r.type));
    const layoutP=lp&&lp.target,layout=layoutP?await parse(layoutP):null;
    let masterP=null,master=null;if(layoutP){const lr=await relsOf(layoutP),mp=Object.values(lr).find(r=>/\/slideMaster$/.test(r.type));if(mp){masterP=mp.target;master=await parse(masterP);}}
    const key=layoutP||slidePath;if(ctxCache[key])return Object.assign({sr},ctxCache[key]);
    const th=masterP?await themeOf(masterP):{clr:{},major:'Calibri Light',minor:'Calibri'};
    const cm=master&&path(master.documentElement,'clrMap'),map={bg1:'lt1',tx1:'dk1',bg2:'lt2',tx2:'dk2'};
    if(cm)['bg1','tx1','bg2','tx2'].forEach(k=>{if(at(cm,k))map[k]=at(cm,k);});
    const c={layout,master,th,map,layoutP,masterP};ctxCache[key]=c;return Object.assign({sr},c);
  }
  const mkColor=C=>{
    const sch=n=>{n=C.map[n]||n;return C.th.clr[n]||null;};
    return function color(holder){
      if(!holder)return null;
      const el=[...holder.childNodes].find(n=>n.nodeType===1&&/^(srgbClr|schemeClr|sysClr|prstClr)$/.test(n.localName));if(!el)return null;
      let h=el.localName==='srgbClr'?at(el,'val'):el.localName==='schemeClr'?sch(at(el,'val')):el.localName==='sysClr'?at(el,'lastClr')||(at(el,'val')==='window'?'FFFFFF':'000000'):hex6(at(el,'val'));
      if(!h||!/^[0-9a-fA-F]{6}$/.test(h))return null;h=h.toUpperCase();
      const lm=kid(el,'lumMod'),lo=kid(el,'lumOff'),sh=kid(el,'shade'),ti=kid(el,'tint');
      if(lm||lo)h=lumShift(h,lm?num(at(lm,'val'),100000)/100000:1,lo?num(at(lo,'val'),0)/100000:0);
      if(sh)h=toHex(rgb(h).map(v=>v*num(at(sh,'val'),100000)/100000));
      if(ti){const t=num(at(ti,'val'),100000)/100000;h=toHex(rgb(h).map(v=>v*t+255*(1-t)));}
      return '#'+h;
    };
  };
  const fillOf=(spPr,color)=>{   // -> '#rrggbb' | '' (tanpa isian) | null (tidak ditentukan)
    if(!spPr)return null;if(kid(spPr,'noFill'))return '';
    const s=kid(spPr,'solidFill');if(s)return color(s)||null;
    const g=kid(spPr,'gradFill');if(g){const gs=path(g,'gsLst','gs');return gs?color(gs)||null:null;}
    return null;
  };

  /* urutan slide, id, bagian */
  const slideRefs=kids(path(pres.documentElement,'sldIdLst'),'sldId').map(s=>({id:at(s,'id'),part:prels[rid(s,'id')]&&prels[rid(s,'id')].target})).filter(s=>s.part);
  if(!slideRefs.length)throw new Error('Berkas tidak berisi slide');
  const sectionOf={};
  [...pres.getElementsByTagNameNS('*','section')].forEach(sec=>{const ids=[...sec.getElementsByTagNameNS('*','sldId')].map(e=>at(e,'id'));
    const first=ids.map(id=>slideRefs.findIndex(s=>s.id===id)).filter(i=>i>=0).sort((a,b)=>a-b)[0];if(first!==undefined)sectionOf[first]=at(sec,'name')||'Bagian';});

  const out=[];
  for(let si=0;si<slideRefs.length;si++){
    const part=slideRefs[si].part,doc=await parse(part);if(!doc){out.push({layout:'blank',objs:[],notes:''});continue;}
    const C=await ctxFor(part),color=mkColor(C),sld=doc.documentElement;
    const tree=path(sld,'cSld','spTree'),objs=[],roles=[];

    /* latar slide: slide -> layout -> master */
    const bgOf=d=>{const bg=d&&path(d.documentElement,'cSld','bg');if(!bg)return null;
      const pr=kid(bg,'bgPr');if(pr)return fillOf(pr,color);const rf=kid(bg,'bgRef');return rf?color(rf):null;};
    let bg=bgOf(doc);if(bg==null)bg=bgOf(C.layout);if(bg==null)bg=bgOf(C.master);

    /* pencarian placeholder di layout/master */
    const phKey=sp=>{const ph=path(sp,'nvSpPr','nvPr','ph');return ph?{type:at(ph,'type')||'obj',idx:at(ph,'idx')}:null;};
    const norm=t=>t==='ctrTitle'?'title':(t==='subTitle'||t==='obj')?'body':t;
    const findPh=(d,k)=>{const t=d&&path(d.documentElement,'cSld','spTree');if(!t||!k)return null;
      const all=kids(t,'sp').map(sp=>({sp,k:phKey(sp)})).filter(x=>x.k);
      return(k.idx&&all.find(x=>x.k.idx===k.idx))||all.find(x=>norm(x.k.type)===norm(k.type))||null;};

    /* gaya teks bertingkat */
    const rInfo=d=>{const r={};if(!d)return r;
      if(at(d,'sz'))r.sz=+at(d,'sz')/100;if(at(d,'b')!=null)r.b=at(d,'b')==='1'||at(d,'b')==='true';if(at(d,'i')!=null)r.i=at(d,'i')==='1'||at(d,'i')==='true';
      if(at(d,'u'))r.u=at(d,'u')!=='none';if(at(d,'strike'))r.s=at(d,'strike')!=='noStrike';if(at(d,'spc'))r.spc=+at(d,'spc')/100;
      const sf=kid(d,'solidFill');if(sf){const c=color(sf);if(c)r.color=c;}
      const lat=kid(d,'latin');if(lat&&at(lat,'typeface')){const f=at(lat,'typeface');r.font=f==='+mj-lt'?C.th.major:f==='+mn-lt'?C.th.minor:f;}
      return r;};
    const pInfo=p=>{const r={};if(!p)return r;if(at(p,'algn'))r.algn={l:'left',ctr:'center',r:'right',just:'justify'}[at(p,'algn')];
      const ls=path(p,'lnSpc','spcPct');if(ls)r.lh=num(at(ls,'val'),100000)/100000;
      if(kid(p,'buNone'))r.bu='none';else if(kid(p,'buAutoNum'))r.bu='num';else if(kid(p,'buChar'))r.bu='char';
      Object.assign(r,rInfo(kid(p,'defRPr')));return r;};
    const lvlOf=(srcs,l)=>srcs.reduce((a,ls)=>Object.assign(a,pInfo(ls&&kid(ls,'lvl'+(l+1)+'pPr'))),{});

    async function shapes(parent,T){
      for(const el of [...parent.childNodes].filter(n=>n.nodeType===1)){
        const ln=el.localName;
        if(ln==='grpSp'){
          const gx=path(el,'grpSpPr','xfrm');
          if(gx){const off=kid(gx,'off'),ext=kid(gx,'ext'),co=kid(gx,'chOff'),ce=kid(gx,'chExt');
            const sx=num(at(ext,'cx'),1)/(num(at(ce,'cx'),1)||1),sy=num(at(ext,'cy'),1)/(num(at(ce,'cy'),1)||1);
            const T2={sx:T.sx*sx,sy:T.sy*sy,tx:T.tx+(num(at(off,'x'),0)-num(at(co,'x'),0)*sx)*T.sx,ty:T.ty+(num(at(off,'y'),0)-num(at(co,'y'),0)*sy)*T.sy};
            await shapes(el,T2);}else await shapes(el,T);
          continue;}
        if(ln==='graphicFrame'){const u=el.getElementsByTagNameNS('*','graphicData')[0],uri=u?at(u,'uri')||'':'';
          if(/table/.test(uri))warn.table++;else if(/chart/.test(uri))warn.chart++;else warn.other++;continue;}
        if(ln==='pic'){await pic(el,T);continue;}
        if(ln==='sp'||ln==='cxnSp')await shape(el,T);
      }
    }
    const box=(xf,T,inh)=>{
      let e=xf;if(!e&&inh)e=inh;if(!e)return null;
      const off=kid(e,'off'),ext=kid(e,'ext');if(!off||!ext)return null;
      const x=(T.tx+num(at(off,'x'),0)*T.sx)*K+offX,y=(T.ty+num(at(off,'y'),0)*T.sy)*K+offY,w=num(at(ext,'cx'),0)*T.sx*K,h=num(at(ext,'cy'),0)*T.sy*K;
      return{x,y,w,h,rot:num(at(e,'rot'),0)/60000,fh:at(e,'flipH')==='1'?1:0,fv:at(e,'flipV')==='1'?1:0};
    };
    async function pic(el,T){
      const bl=path(el,'blipFill','blip'),id=rid(bl,'embed'),r=id&&C.sr[id];if(!r)return;
      const ext=r.target.split('.').pop().toLowerCase(),mime=MIME[ext],f=zip.file(r.target);
      if(!mime||!f){warn.media++;return;}
      const b=box(path(el,'spPr','xfrm'),T);if(!b)return;
      objs.push({k:'sh',shape:'rect',x:Math.round(b.x),y:Math.round(b.y),w:Math.round(b.w),h:Math.round(b.h),rot:b.rot,fh:b.fh,fv:b.fv,img:'data:'+mime+';base64,'+await f.async('base64'),html:''});
    }
    async function shape(el,T){
      const nv=path(el,'nvSpPr')||path(el,'nvCxnSpPr'),cNv=kid(nv,'cNvPr'),nme=at(cNv,'name')||'';
      const spPr=kid(el,'spPr'),style=kid(el,'style'),ph=path(nv,'nvPr','ph'),k=ph?{type:at(ph,'type')||'obj',idx:at(ph,'idx')}:null;
      const lay=k&&findPh(C.layout,k),mas=k&&findPh(C.master,k);
      let role=null;
      if(k){const t=k.type;role=(t==='title'||t==='ctrTitle')?'title':t==='subTitle'?'sub':(t==='body'||t==='obj')?'body':null;}
      else if(/^Title\b/.test(nme))role='title';else if(/^Subtitle\b/.test(nme))role='sub';else if(/^Content Placeholder\b/.test(nme))role='body';
      const b=box(path(spPr,'xfrm'),T,null)||(k&&(box(lay&&path(lay.sp,'spPr','xfrm'),{sx:1,sy:1,tx:0,ty:0})||box(mas&&path(mas.sp,'spPr','xfrm'),{sx:1,sy:1,tx:0,ty:0})));
      if(!b)return;
      const prst=at(path(spPr,'prstGeom'),'prst')||'rect';
      /* garis & konektor */
      if(el.localName==='cxnSp'||prst==='line'||prst==='straightConnector1'){
        const ln0=kid(spPr,'ln'),lc=ln0&&kid(ln0,'solidFill');let stroke=lc?color(lc):null;
        if(!stroke&&style){const lr=kid(style,'lnRef');stroke=color(lr);}
        const cx=b.w,cy=b.h,vx=b.fh?-cx:cx,vy=b.fv?-cy:cy,L=Math.hypot(cx,cy);
        let ang=Math.atan2(vy,vx)*180/Math.PI+b.rot;
        const tail=ln0&&kid(ln0,'tailEnd'),head=ln0&&kid(ln0,'headEnd'),isA=e=>e&&at(e,'type')&&at(e,'type')!=='none';
        if(isA(head)&&!isA(tail))ang+=180;
        const ccx=b.x+cx/2,ccy=b.y+cy/2,len=Math.max(L,1);
        objs.push({k:'sh',shape:(isA(tail)||isA(head))?'arrow':'line',x:Math.round(ccx-len/2),y:Math.round(ccy-12),w:Math.round(len),h:24,rot:((Math.round(ang)%360)+360)%360,
          stroke:stroke||'#000000',sw:ln0&&at(ln0,'w')?Math.max(.5,num(at(ln0,'w'),12700)/12700):1,html:''});
        return;}
      /* isian & garis tepi */
      let fill=fillOf(spPr,color),stroke='',sw=1,dash='';
      const ln0=kid(spPr,'ln');
      if(ln0){if(kid(ln0,'noFill'))stroke='';else{const lc=kid(ln0,'solidFill');stroke=lc?color(lc)||'':'';
        if(!stroke&&style)stroke=color(kid(style,'lnRef'))||'';
        if(at(ln0,'w'))sw=Math.max(.5,num(at(ln0,'w'),12700)/12700);
        const pd=at(kid(ln0,'prstDash'),'val');dash=/dot|sysDot/.test(pd||'')&&!/dash/i.test(pd||'')?'2 4':(pd&&pd!=='solid')?'8 5':'';}}
      else if(style&&!role){const lr=kid(style,'lnRef');if(lr&&at(lr,'idx')!=='0')stroke=color(lr)||'';}
      if(fill==null){fill='';if(style&&!role){const fr=kid(style,'fillRef');if(fr&&at(fr,'idx')!=='0'&&at(fr,'idx')!=='1000')fill=color(fr)||'';}}
      const eff=kid(spPr,'effectLst'),fx=eff&&kid(eff,'outerShdw')?'shadow':'';
      /* teks */
      const tb=kid(el,'txBody'),bp=kid(tb,'bodyPr'),bpL=lay&&path(lay.sp,'txBody','bodyPr'),bpM=mas&&path(mas.sp,'txBody','bodyPr');
      const bAttr=n=>at(bp,n)||at(bpL,n)||at(bpM,n);
      const anchor={t:'t',ctr:'m',b:'b'}[bAttr('anchor')]||'t';
      const vert={vert:'v90',vert270:'v270',wordArtVert:'st',eaVert:'v90'}[at(bp,'vert')]||'h';
      const cols=Math.max(1,Math.min(3,num(at(bp,'numCol'),1)));
      const fsc=num(at(path(bp,'normAutofit'),'fontScale'),100000)/100000;
      const ptype=k?(norm(k.type)==='title'?'title':norm(k.type)==='body'?'body':'other'):'other';
      const styleEl=C.master&&path(C.master.documentElement,'txStyles',ptype==='title'?'titleStyle':ptype==='body'?'bodyStyle':'otherStyle');
      const srcs=[k?styleEl:defTextStyle,mas&&path(mas.sp,'txBody','lstStyle'),lay&&path(lay.sp,'txBody','lstStyle'),tb&&kid(tb,'lstStyle')];
      const dflt={sz:ptype==='title'?44:18,color:'#'+(C.th.clr[C.map.tx1]||'000000'),font:C.th.minor};
      const fontRefClr=!k&&style?color(kid(style,'fontRef')):null;   // teks bentuk berbasis style (mis. putih di atas biru)
      const paras=[];
      for(const p of kids(tb,'p')){
        const pPr=kid(p,'pPr'),lvl=Math.min(8,num(at(pPr,'lvl'),0)),base=Object.assign({},dflt,lvlOf(srcs,lvl)),runs=[];
        if(fontRefClr&&!lvlOf([tb&&kid(tb,'lstStyle')],lvl).color)base.color=fontRefClr;
        Object.assign(base,pInfo(pPr));
        for(const c of [...p.childNodes].filter(n=>n.nodeType===1)){
          if(c.localName==='br'){runs.push({br:1});continue;}
          if(c.localName!=='r'&&c.localName!=='fld')continue;
          const t=kid(c,'t');if(!t)continue;
          const ri=Object.assign({},base,rInfo(kid(c,'rPr')));delete ri.algn;
          runs.push({text:t.textContent,f:ri});
        }
        const bu=base.bu==='num'?'n':base.bu==='char'?'b':null;
        paras.push({runs,base,lvl,bu});
      }
      const anyText=paras.some(p=>p.runs.some(r=>r.text&&r.text.trim()));
      if(!anyText&&!role&&!fill&&!stroke)return;     // bentuk kosong tak kasatmata
      const first=paras.map(p=>p.runs.find(r=>r.text))[paras.findIndex(p=>p.runs.some(r=>r.text))]||{f:paras[0]?paras[0].base:dflt};
      const f0=first.f||dflt,ts={font:f0.font||dflt.font,size:Math.round((f0.sz||dflt.sz)*FS*fsc*10)/10,color:f0.color||dflt.color,align:(paras[0]&&paras[0].base.algn)||'left',lh:Math.round(((paras[0]&&paras[0].base.lh)||1)*100)/100};
      const runHTML=r=>{
        if(r.br)return '<br>';
        let s=esc(r.text),st='';const f=r.f;
        const size=Math.round((f.sz||dflt.sz)*FS*fsc*10)/10;
        if(f.font&&f.font!==ts.font)st+=`font-family:'${f.font}';`;
        if(size!==ts.size)st+=`font-size:${size}pt;`;
        if(f.color&&f.color.toLowerCase()!==ts.color.toLowerCase())st+=`color:${f.color};`;
        if(f.spc)st+=`letter-spacing:${(f.spc/.75).toFixed(1)}px;`;
        if(st)s=`<span style="${st}">${s}</span>`;
        if(f.b)s=`<b>${s}</b>`;if(f.i)s=`<i>${s}</i>`;if(f.u)s=`<u>${s}</u>`;if(f.s)s=`<s>${s}</s>`;
        return s;};
      let html='',stack=[];
      const closeTo=n=>{while(stack.length>n)html+='</'+stack.pop()+'>';};
      if(anyText)for(const p of paras){
        const inner=p.runs.map(runHTML).join('')||'<br>',al=p.base.algn&&p.base.algn!==ts.align?` style="text-align:${p.base.algn}"`:'';
        if(p.bu){const tag=p.bu==='n'?'ol':'ul',want=p.lvl+1;
          closeTo(Math.min(stack.length,want));
          if(stack.length===want&&stack[want-1]!==tag)closeTo(want-1);
          while(stack.length<want){html+='<'+tag+'>';stack.push(tag);}
          html+=`<li${al}>${inner}</li>`;}
        else{closeTo(0);html+=`<div${al}>${inner}</div>`;}
      }
      closeTo(0);
      const isTx=!role&&!fill&&!stroke&&(prst==='rect'||prst==='textBox'),shp=SHAPE_IN[prst]||'rect';
      const o={k:role?'ph':isTx?'tb':'sh',shape:shp==='line'?'rect':shp,x:Math.round(b.x),y:Math.round(b.y),w:Math.round(b.w),h:Math.round(b.h),rot:b.rot,fh:b.fh,fv:b.fv,
        fill:fill||'',stroke:stroke||'',sw,dash,fx,anchor,vert,cols,html:anyText?html:'',ts};
      if(role){o.role=role;if(role==='body')o.bul=1;roles.push(role);}
      objs.push(o);
    }
    await shapes(tree,{sx:1,sy:1,tx:0,ty:0});

    /* notes */
    let notes='';const nr=Object.values(C.sr).find(r=>/\/notesSlide$/.test(r.type));
    if(nr){const nd=await parse(nr.target);if(nd){const t=path(nd.documentElement,'cSld','spTree');
      const body=kids(t,'sp').find(sp=>{const ph=path(sp,'nvSpPr','nvPr','ph');return ph&&at(ph,'type')==='body';});
      if(body){const ls=kids(path(body,'txBody'),'p').map(p=>[...p.getElementsByTagNameNS('*','t')].map(t=>t.textContent).join(''));
        if(ls.some(x=>x.trim()))notes=ls.map(x=>'<div>'+(esc(x)||'<br>')+'</div>').join('');}}}

    /* layout */
    const lt=C.layout&&at(C.layout.documentElement,'type'),LMAP={title:'title',obj:'content',tx:'content',secHead:'section',twoObj:'two',twoTxTwoObj:'two',titleOnly:'tonly',blank:'blank',objOnly:'content'};
    let layout=LMAP[lt];
    if(!layout){const nb=roles.filter(r=>r==='body').length;layout=roles.includes('title')?(roles.includes('sub')?'title':nb>=2?'two':nb===1?'content':'tonly'):(nb?'content':'blank');}
    const s={layout,objs,notes};if(bg&&bg.toLowerCase()!=='#ffffff')s.bg=bg;if(sectionOf[si]!=null)s.sec=sectionOf[si];
    out.push(s);
  }
  return{slides:out,warn};
}

root.PptxIO={write,read};
})(typeof window!=='undefined'?window:globalThis);
