/* Baca/tulis .docx untuk editor Word.
   - toBlob(ctx): isi halaman -> dokumen Word sungguhan (library `docx`), lengkap dengan header/footer, nomor halaman, ukuran/margin/kolom.
   - readDocx(buf,JSZip): ambil header, footer, dan tata letak halaman asli dari berkas .docx.
   Mengekspos window.WordDocx. */
(function(root){
'use strict';
const TW=15;                                   // 1 px = 15 twip
const PAPER={a4:[11906,16838],letter:[12240,15840],legal:[12240,20160]};
const W_NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/* ---------- util ---------- */
const len=v=>{                                 // CSS length -> px
  if(!v)return null;const m=/^(-?[\d.]+)(px|pt|cm|mm|in|em)?$/.exec(String(v).trim());if(!m)return null;
  const n=+m[1],u=m[2]||'px';return u==='px'?n:u==='pt'?n/.75:u==='cm'?n*37.795:u==='mm'?n*3.7795:u==='in'?n*96:u==='em'?n*16:null;
};
const hex=c=>{
  if(!c)return null;c=String(c).trim().toLowerCase();
  let m=/^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?\s*\)$/.exec(c);
  if(m){if(m[4]!==undefined&&+m[4]===0)return null;return[m[1],m[2],m[3]].map(x=>(+x).toString(16).padStart(2,'0')).join('').toUpperCase();}
  m=/^#([0-9a-f]{3})$/.exec(c);if(m)return m[1].split('').map(x=>x+x).join('').toUpperCase();
  m=/^#([0-9a-f]{6})$/.exec(c);if(m)return m[1].toUpperCase();
  const N={black:'000000',white:'FFFFFF',red:'FF0000',blue:'0000FF',green:'008000',yellow:'FFFF00',gray:'808080',grey:'808080'};
  return N[c]||null;
};
const FONT_SIZES=[8,10,12,14,18,24,36];        // <font size=1..7>
const b64=s=>{const bin=atob(s),u=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)u[i]=bin.charCodeAt(i);return u;};
const imgSize=u=>{                             // ukuran asli gambar dari header berkas
  try{
    if(u[0]===0x89&&u[1]===0x50)return[(u[16]<<24|u[17]<<16|u[18]<<8|u[19])>>>0,(u[20]<<24|u[21]<<16|u[22]<<8|u[23])>>>0];
    if(u[0]===0x47&&u[1]===0x49)return[u[6]|u[7]<<8,u[8]|u[9]<<8];
    if(u[0]===0xFF&&u[1]===0xD8){let i=2;while(i<u.length){if(u[i]!==0xFF){i++;continue;}const t=u[i+1];
      if(t>=0xC0&&t<=0xCF&&t!==0xC4&&t!==0xC8&&t!==0xCC)return[u[i+7]<<8|u[i+8],u[i+5]<<8|u[i+6]];i+=2+(u[i+2]<<8|u[i+3]);}}
    if(u[0]===0x42&&u[1]===0x4D)return[u[18]|u[19]<<8|u[20]<<16,Math.abs(u[22]|u[23]<<8|u[24]<<16)];
  }catch(e){}
  return null;
};

/* ======================================================================
   TULIS: isi halaman -> .docx
   ====================================================================== */
async function toBlob(c){
  const X=root.docx;if(!X)throw new Error('library docx belum dimuat');
  const {page,cfg,D}=c;
  const doc0=page.ownerDocument;
  const paper=PAPER[cfg.size]||PAPER.a4;
  const mTw=cfg.m.map(v=>Math.round(v*TW));                    // atas, kanan, bawah, kiri
  const pageW=cfg.land?paper[1]:paper[0];
  const contentW=Math.max(1440,pageW-mTw[1]-mTw[3]);           // lebar isi (twip), dipakai untuk tab & tabel
  const colCount=cfg.cols>1?cfg.cols:1;
  const colSpace=36*TW;
  const bodyW=colCount>1?Math.floor((contentW-colSpace*(colCount-1))/colCount):contentW;
  const st={ol:0,pendingBreak:false,rtl:c.rtl};
  let tabW=bodyW;

  /* --- konteks run --- */
  const CLS={
    's-title':{size:56,font:D.h},'s-sub':{color:'5A5A5A',spacing:15},'s-em1':{italics:true,color:'404040'},'s-em2':{italics:true},
    's-em3':{italics:true,bold:true,color:(hex(D.hc)||'2F5496')},'s-strong':{bold:true}
  };
  function ctxFor(el,c0){
    const n={...c0},s=el.style||{},t=el.tagName;
    if(t==='B'||t==='STRONG')n.bold=true;
    if(/^(I|EM|CITE|DFN)$/.test(t))n.italics=true;
    if(t==='U'||t==='INS')n.underline=true;
    if(/^(S|STRIKE|DEL)$/.test(t))n.strike=true;
    if(t==='SUB'){n.sub=true;n.sup=false;}
    if(t==='SUP'){n.sup=true;n.sub=false;}
    if(/^(CODE|KBD|TT|SAMP)$/.test(t))n.font='Courier New';
    if(t==='MARK')n.shade='FFFF00';
    if(t==='A'&&el.getAttribute('href')){n.color='0563C1';n.underline=true;}
    if(t==='FONT'){
      if(el.getAttribute('face'))n.font=el.getAttribute('face').split(',')[0].replace(/["']/g,'').trim();
      const col=hex(el.getAttribute('color'));if(col)n.color=col;
      const sz=FONT_SIZES[(+el.getAttribute('size')||0)-1];if(sz)n.size=sz*2;
    }
    el.classList&&[...el.classList].forEach(k=>{if(CLS[k])Object.assign(n,CLS[k]);});
    const fw=s.fontWeight;if(fw)n.bold=fw==='bold'||fw==='bolder'||+fw>=600;
    const fs=s.fontStyle;if(fs)n.italics=fs==='italic'||fs==='oblique';
    const td=(s.textDecorationLine||s.textDecoration||'');
    if(td){if(/underline/.test(td))n.underline=true;if(/line-through/.test(td))n.strike=true;if(td==='none')n.underline=n.strike=false;}
    if(s.fontFamily)n.font=s.fontFamily.split(',')[0].replace(/["']/g,'').trim();
    const fz=len(s.fontSize);if(fz)n.size=Math.max(2,Math.round(fz*.75*2));
    const col=hex(s.color);if(col)n.color=col;
    const bg=hex(s.backgroundColor||(/^(#|rgb)/i.test(s.background||'')?s.background:''));if(bg)n.shade=bg;
    if(s.verticalAlign==='super'){n.sup=true;n.sub=false;}else if(s.verticalAlign==='sub'){n.sub=true;n.sup=false;}
    if(s.letterSpacing){const ls=len(s.letterSpacing);if(ls)n.spacing=Math.round(ls*.75*20);}
    return n;
  }
  const runOpts=c0=>({
    bold:c0.bold||undefined,italics:c0.italics||undefined,
    underline:c0.underline?{type:X.UnderlineType.SINGLE}:undefined,strike:c0.strike||undefined,
    subScript:c0.sub||undefined,superScript:c0.sup||undefined,font:c0.font,size:c0.size,color:c0.color,
    characterSpacing:c0.spacing,shading:c0.shade?{type:X.ShadingType.CLEAR,fill:c0.shade,color:'auto'}:undefined
  });
  const mkRun=(text,c0,extra)=>new X.TextRun({text,...runOpts(c0),...extra});

  /* --- gambar --- */
  function imageRun(img,c0){
    const m=/^data:image\/(png|jpe?g|gif|bmp);base64,(.*)$/is.exec(img.getAttribute('src')||'');
    if(!m)return null;
    const data=b64(m[2].replace(/\s/g,'')),type=m[1].toLowerCase().replace('jpeg','jpg'),nat=imgSize(data)||[img.naturalWidth||200,img.naturalHeight||100];
    const s=img.style;let w=len(s.width)||len(img.getAttribute('width')),h=len(s.height)||len(img.getAttribute('height'));
    if(w&&!h)h=w*nat[1]/nat[0];else if(h&&!w)w=h*nat[0]/nat[1];else if(!w&&!h){w=nat[0];h=nat[1];}
    const maxW=Math.floor(bodyW/TW);if(w>maxW){h=h*maxW/w;w=maxW;}
    const tr={width:Math.max(1,Math.round(w)),height:Math.max(1,Math.round(h))};
    const rot=((+img.dataset.r||0)%360+360)%360;if(rot)tr.rotation=rot;
    const fx=+img.dataset.fx<0,fy=+img.dataset.fy<0;if(fx||fy)tr.flip={horizontal:fx||undefined,vertical:fy||undefined};
    const o={type,data,transformation:tr,altText:{name:'Gambar',title:img.getAttribute('alt')||'Gambar',description:img.getAttribute('alt')||'Gambar'}};
    if(s.float==='left'||s.float==='right'){
      o.floating={horizontalPosition:{relative:X.HorizontalPositionRelativeFrom.MARGIN,align:s.float==='left'?X.HorizontalPositionAlign.LEFT:X.HorizontalPositionAlign.RIGHT},
        verticalPosition:{relative:X.VerticalPositionRelativeFrom.PARAGRAPH,offset:0},
        wrap:{type:X.TextWrappingType.SQUARE,side:X.TextWrappingSide.BOTH_SIDES},
        margins:{left:s.float==='right'?114300:0,right:s.float==='left'?114300:0,top:0,bottom:76200}};
    }
    return new X.ImageRun(o);
  }
  const imgAlign=img=>{const s=img.style;
    if(s.marginLeft==='auto'&&s.marginRight==='auto')return X.AlignmentType.CENTER;
    if(s.marginLeft==='auto')return X.AlignmentType.RIGHT;return null;};

  /* --- inline --- */
  const hidden=n=>{const s=n.style;return n.classList.contains('pgap')||/^(SCRIPT|STYLE)$/.test(n.tagName)||(s&&(s.display==='none'||s.visibility==='hidden'));};
  function inline(nodes,c0,ps){                   // ps = state paragraf {tabRight}
    const out=[];
    const walk=(list,cx)=>{
      for(const n of list){
        if(n.nodeType===3){
          const t=n.nodeValue.replace(/[ \t\r\n]+/g,' ').replace(/​/g,'');
          if(t)out.push(mkRun(t,cx));continue;
        }
        if(n.nodeType!==1||hidden(n))continue;
        const t=n.tagName;
        if(t==='BR'){out.push(new X.TextRun({break:1}));continue;}
        if(t==='IMG'){const r=imageRun(n,cx);if(r)out.push(r);else if(n.getAttribute('alt'))out.push(mkRun(n.getAttribute('alt'),cx));continue;}
        if(n.classList.contains('pn')){out.push(new X.TextRun({children:[X.PageNumber.CURRENT],...runOpts(ctxFor(n,cx))}));continue;}
        if(n.classList.contains('pt')){out.push(new X.TextRun({children:[X.PageNumber.TOTAL_PAGES],...runOpts(ctxFor(n,cx))}));continue;}
        const nc=ctxFor(n,cx);
        if(n.style&&n.style.float==='right'){                 // nomor halaman di daftar isi -> tab rata kanan
          ps.tabRight=true;out.push(new X.TextRun({children:[new X.Tab()]}));walk(n.childNodes,nc);continue;
        }
        if(t==='A'&&/^https?:|^mailto:/i.test(n.getAttribute('href')||'')){
          const sub=[];const keep=out.splice(0,out.length);walk(n.childNodes,nc);sub.push(...out.splice(0,out.length));out.push(...keep);
          if(sub.length)out.push(new X.ExternalHyperlink({link:n.getAttribute('href'),children:sub}));continue;
        }
        walk(n.childNodes,nc);
      }
    };
    // buang <br> penutup (placeholder browser)
    let list=[...nodes];while(list.length&&list[list.length-1].nodeType===1&&list[list.length-1].tagName==='BR')list.pop();
    while(list.length&&list[list.length-1].nodeType===3&&!list[list.length-1].nodeValue.trim()&&list.length>1)list.pop();
    walk(list,c0);return out;
  }

  /* --- paragraf --- */
  const AL={left:X.AlignmentType.LEFT,center:X.AlignmentType.CENTER,right:X.AlignmentType.RIGHT,justify:X.AlignmentType.BOTH,start:X.AlignmentType.LEFT,end:X.AlignmentType.RIGHT};
  function pProps(el,o){
    if(!el||!el.style)return o;
    const s=el.style;
    const al=AL[s.textAlign||el.getAttribute&&el.getAttribute('align')];if(al&&!o.alignment)o.alignment=al;
    const sp={...(o.spacing||{})};
    const mt=len(s.marginTop),mb=len(s.marginBottom);
    if(mt!==null)sp.before=Math.round(mt*TW);if(mb!==null)sp.after=Math.round(mb*TW);
    const lh=s.lineHeight;if(lh){if(/^[\d.]+$/.test(lh)){sp.line=Math.round(+lh*240);sp.lineRule=X.LineRuleType.AUTO;}
      else{const px=len(lh);if(px){sp.line=Math.round(px*TW);sp.lineRule=X.LineRuleType.AT_LEAST;}}}
    if(el.classList.contains('s-nosp')){if(sp.after===undefined)sp.after=0;sp.line=240;sp.lineRule=X.LineRuleType.AUTO;}
    if(el.classList.contains('s-title')){if(sp.after===undefined)sp.after=160;sp.line=240;sp.lineRule=X.LineRuleType.AUTO;}
    if(Object.keys(sp).length)o.spacing=sp;
    const ind={...(o.indent||{})};
    const L=(len(s.marginLeft)||0)+(len(s.paddingLeft)||0),R=(len(s.marginRight)||0)+(len(s.paddingRight)||0);
    if(L)ind.left=(ind.left||0)+Math.round(L*TW);if(R)ind.right=(ind.right||0)+Math.round(R*TW);
    const ti=len(s.textIndent);if(ti){if(ti<0){ind.hanging=Math.round(-ti*TW);}else ind.firstLine=Math.round(ti*TW);}
    if(Object.keys(ind).length)o.indent=ind;
    const bg=hex(s.backgroundColor||(/^(#|rgb)/i.test(s.background||'')?s.background:''));
    if(bg)o.shading={type:X.ShadingType.CLEAR,fill:bg,color:'auto'};
    if(el.tagName==='BLOCKQUOTE'){o.indent={left:720,right:720,...(o.indent||{})};o.spacing={before:160,after:160,...(o.spacing||{})};}
    return o;
  }
  function makePara(nodes,c0,el,extra){
    nodes=[...nodes];
    const ps={tabRight:false};
    const ctx=el?ctxFor(el,c0):c0;
    if(el&&el.tagName==='BLOCKQUOTE'){ctx.italics=true;if(!ctx.color)ctx.color='404040';}
    let runs=inline(nodes,ctx,ps);
    const o={children:runs,...extra};
    if(st.rtl)o.bidirectional=true;
    if(st.pendingBreak){o.pageBreakBefore=true;st.pendingBreak=false;}
    if(ps.tabRight)o.tabStops=[{type:X.TabStopType.RIGHT,position:tabW,leader:X.LeaderType.DOT}];
    pProps(el,o);
    // gambar tunggal dengan margin auto -> rata paragraf
    const real=nodes.filter(n=>!(n.nodeType===3&&!n.nodeValue.trim()));
    if(real.length===1&&real[0].tagName==='IMG'&&!o.alignment){const a=imgAlign(real[0]);if(a)o.alignment=a;}
    return new X.Paragraph(o);
  }
  function flexPara(el,c0){                       // baris beberapa kolom (header/footer "Tiga kolom")
    const kids=[...el.children].filter(k=>!hidden(k)),ps={tabRight:false},runs=[];
    kids.forEach((k,i)=>{if(i)runs.push(new X.TextRun({children:[new X.Tab()]}));runs.push(...inline([k],c0,ps));});
    const o={children:runs};
    if(kids.length===3)o.tabStops=[{type:X.TabStopType.CENTER,position:Math.round(tabW/2)},{type:X.TabStopType.RIGHT,position:tabW}];
    else if(kids.length===2)o.tabStops=[{type:X.TabStopType.RIGHT,position:tabW}];
    if(st.rtl)o.bidirectional=true;
    return new X.Paragraph(o);
  }

  /* --- blok --- */
  const BLOCK=/^(P|H[1-6]|UL|OL|TABLE|DIV|BLOCKQUOTE|HR|PRE|SECTION|ARTICLE|FIGURE|CENTER|HEADER|FOOTER|MAIN|ASIDE|NAV)$/;
  const hasBlock=n=>[...n.children].some(k=>BLOCK.test(k.tagName));
  const HEAD={H1:X.HeadingLevel.HEADING_1,H2:X.HeadingLevel.HEADING_2,H3:X.HeadingLevel.HEADING_3,H4:X.HeadingLevel.HEADING_4,H5:X.HeadingLevel.HEADING_5,H6:X.HeadingLevel.HEADING_6};
  function blocks(node,c0){
    const out=[];let inl=[];
    const flush=()=>{if(inl.some(n=>n.nodeType===1||n.nodeValue.trim()))out.push(makePara(inl,c0,null));inl=[];};
    for(const n of node.childNodes){
      if(n.nodeType===3){inl.push(n);continue;}
      if(n.nodeType!==1||hidden(n))continue;
      if(BLOCK.test(n.tagName)){flush();out.push(...blockEl(n,c0));}else inl.push(n);
    }
    flush();return out;
  }
  function blockEl(n,c0){
    const t=n.tagName;
    if(t==='UL'||t==='OL')return list(n,c0,0,t==='OL'?++st.ol:0);
    if(t==='TABLE')return [table(n,c0)];
    if(t==='HR')return [new X.Paragraph({border:{bottom:{style:X.BorderStyle.SINGLE,size:6,color:'808080',space:1}},spacing:{after:120}})];
    if(n.classList.contains('pb')){st.pendingBreak=true;return [];}
    if(n.classList.contains('cb'))return [new X.Paragraph({children:[new X.ColumnBreak()]})];
    if(t==='PRE'){
      return n.textContent.replace(/\n$/,'').split('\n').map(l=>new X.Paragraph({spacing:{after:0},children:[new X.TextRun({text:l,font:'Courier New',size:20})]}));
    }
    if(t==='DIV'&&n.style.display==='flex')return [flexPara(n,c0)];
    if(/^H[1-6]$/.test(t)){
      return [makePara(n.childNodes,c0,n,{heading:HEAD[t]})];
    }
    if((t==='P'||t==='BLOCKQUOTE'||t==='DIV'||t==='CENTER')&&!hasBlock(n)){
      const extra={};if(t==='CENTER')extra.alignment=X.AlignmentType.CENTER;
      return [makePara(n.childNodes,c0,n,extra)];
    }
    return blocks(n,ctxFor(n,c0));
  }

  /* --- daftar --- */
  const bullets=['•','o','▪'];
  function list(el,c0,lvl,inst){
    const ordered=el.tagName==='OL',out=[],level=Math.min(lvl,8);
    const numbering=()=>({reference:ordered?'ol':'ul',level,...(ordered?{instance:inst}:{})});
    for(const li of el.children){
      if(li.tagName!=='LI'||hidden(li))continue;
      const lc=ctxFor(li,c0);let group=[],first=true,any=false;
      const flush=()=>{
        if(group.some(n=>n.nodeType===1||n.nodeValue.trim())){
          out.push(makePara(group,lc,li,first?{numbering:numbering()}:{indent:{left:720+360*level}}));first=false;any=true;
        }
        group=[];
      };
      for(const n of li.childNodes){
        if(n.nodeType===1&&/^(UL|OL)$/.test(n.tagName)){flush();out.push(...list(n,lc,lvl+1,n.tagName==='OL'?++st.ol:0));any=true;}
        else if(n.nodeType===1&&/^(P|DIV)$/.test(n.tagName)&&!hidden(n)){flush();group=[...n.childNodes];flush();}
        else group.push(n);
      }
      flush();
      if(!any)out.push(makePara([],lc,li,{numbering:numbering()}));
    }
    return out;
  }

  /* --- tabel --- */
  const BD={style:X.BorderStyle.SINGLE,size:4,color:'000000'};
  function table(el,c0){
    const rows=[...el.rows];if(!rows.length)return new X.Paragraph({});
    const grid=rows.map(r=>[...r.cells]);
    const ncols=Math.max(1,...grid.map(r=>r.reduce((a,cl)=>a+(cl.colSpan||1),0)));
    let widths=null;
    for(const r of grid){if(r.length===ncols&&r.every(cl=>(cl.rowSpan||1)===1)){const w=r.map(cl=>cl.getBoundingClientRect?cl.getBoundingClientRect().width:0);
      if(w.every(x=>x>0)){widths=w.map(x=>Math.round(x*TW));break;}}}
    if(!widths)widths=Array(ncols).fill(Math.floor(bodyW/ncols));
    const sum=widths.reduce((a,b)=>a+b,0);if(sum>bodyW)widths=widths.map(w=>Math.floor(w*bodyW/sum));
    let occ=Array(ncols).fill(0);
    const trs=rows.map(r=>{
      let ci=0;
      const row=new X.TableRow({
      tableHeader:r.parentNode&&r.parentNode.tagName==='THEAD'||undefined,cantSplit:true,
      children:[...r.cells].map(cl=>{
        while(ci<ncols&&occ[ci]>0)ci++;
        const span=Math.max(1,cl.colSpan||1),cw=widths.slice(ci,ci+span).reduce((x,y)=>x+y,0);
        for(let k=0;k<span&&ci+k<ncols;k++)occ[ci+k]=Math.max(1,cl.rowSpan||1);
        ci+=span;
        const cc=ctxFor(cl,cl.tagName==='TH'?{...c0,bold:true}:c0);
        let ch=blocks(cl,cc);if(!ch.length||ch[ch.length-1]instanceof X.Table)ch.push(new X.Paragraph({}));
        // tinggi paragraf di sel: tanpa jarak bawah bawaan agar sel tidak terlalu tinggi
        const bg=hex(cl.style.backgroundColor||(/^(#|rgb)/i.test(cl.style.background||'')?cl.style.background:''));
        const o={children:ch,width:{size:cw,type:X.WidthType.DXA},columnSpan:cl.colSpan>1?cl.colSpan:undefined,rowSpan:cl.rowSpan>1?cl.rowSpan:undefined,
          shading:bg?{type:X.ShadingType.CLEAR,fill:bg,color:'auto'}:undefined,
          verticalAlign:cl.style.verticalAlign==='middle'?X.VerticalAlign.CENTER:cl.style.verticalAlign==='bottom'?X.VerticalAlign.BOTTOM:X.VerticalAlign.TOP,
          borders:{top:BD,bottom:BD,left:BD,right:BD}};
        return new X.TableCell(o);
      })
    });
      occ=occ.map(v=>Math.max(0,v-1));
      return row;
    });
    const o={rows:trs,columnWidths:widths,width:{size:widths.reduce((a,b)=>a+b,0),type:X.WidthType.DXA},layout:X.TableLayoutType.FIXED};
    if(el.style.marginLeft==='auto'&&el.style.marginRight==='auto')o.alignment=X.AlignmentType.CENTER;
    return new X.Table(o);
  }

  /* --- header / footer --- */
  function hf(html,Cls){
    if(!html)return undefined;
    const box=doc0.createElement('div');box.innerHTML=html;
    if(!box.textContent.replace(/[#\s ]/g,'')&&!box.querySelector('img,.pn,.pt'))return undefined;
    tabW=contentW;const ch=blocks(box,{size:20,color:'555555'});tabW=bodyW;
    st.pendingBreak=false;
    return ch.length?new Cls({children:ch}):undefined;
  }

  /* --- isi dokumen --- */
  const body=blocks(page,{});
  st.pendingBreak=false;
  if(!body.length)body.push(new X.Paragraph({}));
  const header=hf(c.hdr,X.Header),footer=hf(c.ftr,X.Footer);

  /* --- gaya, penomoran, bagian --- */
  const hc=hex(D.hc)||'2F5496',ps=Math.round((+D.ps||0)*20),line=Math.round((+D.lh||1)*240);
  const lvls=(ordered)=>Array.from({length:9},(_,i)=>({
    level:i,format:ordered?[X.LevelFormat.DECIMAL,X.LevelFormat.LOWER_LETTER,X.LevelFormat.LOWER_ROMAN][i%3]:X.LevelFormat.BULLET,
    text:ordered?`%${i+1}.`:bullets[i%3],alignment:X.AlignmentType.LEFT,
    style:{paragraph:{indent:{left:720+360*i,hanging:360}},...(ordered?{}:{run:{font:i%3===1?'Courier New':undefined}})}
  }));
  const hs=(size,color,before)=>({run:{font:D.h,size,color,bold:false},paragraph:{spacing:{before,after:0},keepNext:true,keepLines:true}});
  const pageProps={
    size:{width:paper[0],height:paper[1],orientation:cfg.land?X.PageOrientation.LANDSCAPE:X.PageOrientation.PORTRAIT},
    margin:{top:mTw[0],right:mTw[1],bottom:mTw[2],left:mTw[3],header:540,footer:540}
  };
  if(c.border&&c.border.style){
    const bs=({solid:'SINGLE',double:'DOUBLE',dashed:'DASHED',groove:'THREE_D_ENGRAVE'})[c.border.style]||'SINGLE';
    const b={style:X.BorderStyle[bs],size:Math.max(2,Math.round((c.border.w||1)*6)),color:hex(c.border.color)||'2B579A',space:24};
    pageProps.borders={pageBorders:{display:X.PageBorderDisplay.ALL_PAGES,offsetFrom:X.PageBorderOffsetFrom.PAGE},
      pageBorderTop:b,pageBorderRight:b,pageBorderBottom:b,pageBorderLeft:b};
  }
  const props={page:pageProps};
  if(colCount>1)props.column={count:colCount,space:colSpace,equalWidth:true};
  if(c.lineNumbers)props.lineNumbers={countBy:1,restart:X.LineNumberRestartFormat.CONTINUOUS};
  const section={properties:props,children:body};
  if(header)section.headers={default:header};
  if(footer)section.footers={default:footer};

  const docOpts={
    creator:'Office Web',title:(c.name||'Dokumen').replace(/\.docx$/i,''),
    styles:{default:{
      document:{run:{font:D.b,size:22},paragraph:{spacing:{after:ps,line,lineRule:X.LineRuleType.AUTO}}},
      heading1:hs(32,hc,240),heading2:hs(26,hc,40),heading3:hs(24,'1F3763',40)
    }},
    numbering:{config:[{reference:'ul',levels:lvls(false)},{reference:'ol',levels:lvls(true)}]},
    sections:[section]
  };
  const pc=hex(c.pageColor);if(pc)docOpts.background={color:pc};
  return X.Packer.toBlob(new X.Document(docOpts));
}

/* ======================================================================
   BACA: .docx -> header, footer, tata letak halaman
   ====================================================================== */
async function readDocx(buf,JSZip,DP){
  DP=DP||root.DOMParser;
  const zip=await JSZip.loadAsync(buf);
  const parse=async p=>{const f=zip.file(p);return f?new DP().parseFromString(await f.async('string'),'application/xml'):null;};
  const kids=(el,name)=>[...el.childNodes].filter(n=>n.nodeType===1&&n.localName===name);
  const kid=(el,name)=>kids(el,name)[0]||null;
  const wa=(el,a)=>el?el.getAttributeNS(W_NS,a)||el.getAttribute('w:'+a):null;
  const rels=async part=>{
    const d=await parse(part);const m={};
    if(d)[...d.getElementsByTagName('Relationship')].forEach(r=>{m[r.getAttribute('Id')]=r.getAttribute('Target');});
    return m;
  };
  const resolve=(base,target)=>{   // base: 'word/header1.xml'
    if(target.startsWith('/'))return target.slice(1);
    const parts=base.split('/').slice(0,-1);
    for(const seg of target.split('/')){if(seg==='..')parts.pop();else if(seg!=='.')parts.push(seg);}
    return parts.join('/');
  };
  const res={hdr:'',ftr:'',size:null,land:false,m:null,cols:1};
  const doc=await parse('word/document.xml');if(!doc)return res;
  const bodyEl=doc.getElementsByTagNameNS(W_NS,'body')[0];if(!bodyEl)return res;
  const sect=kid(bodyEl,'sectPr')||[...doc.getElementsByTagNameNS(W_NS,'sectPr')].pop();
  if(!sect)return res;

  /* tata letak halaman */
  const pgSz=kid(sect,'pgSz'),pgMar=kid(sect,'pgMar'),cols=kid(sect,'cols');
  if(pgSz){
    const w=+wa(pgSz,'w'),h=+wa(pgSz,'h');
    if(w&&h){
      const pw=Math.min(w,h),ph=Math.max(w,h);let best=null,bd=1e9;
      for(const k in PAPER){const d=Math.abs(PAPER[k][0]-pw)+Math.abs(PAPER[k][1]-ph);if(d<bd){bd=d;best=k;}}
      res.size=best;res.land=wa(pgSz,'orient')==='landscape'||w>h;
    }
  }
  if(pgMar){const g=a=>Math.round(Math.abs(+wa(pgMar,a)||0)/TW);res.m=[g('top'),g('right'),g('bottom'),g('left')];}
  if(cols&&+wa(cols,'num')>1)res.cols=Math.min(4,+wa(cols,'num'));

  /* header & footer (tipe "default") */
  const docRels=await rels('word/_rels/document.xml.rels');
  const esc=t=>t.replace(/[&<>]/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[ch]));
  const MIME={png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',gif:'image/gif',bmp:'image/bmp'};
  async function partHTML(refEl){
    if(!refEl)return '';
    const rid=refEl.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id')||refEl.getAttribute('r:id');
    const target=docRels[rid];if(!target)return '';
    const path=resolve('word/document.xml',target),xml=await parse(path);if(!xml)return '';
    const prels=await rels(path.replace(/([^/]+)$/,'_rels/$1.rels'));
    const imgs={};
    for(const id in prels){
      const ext=(prels[id].split('.').pop()||'').toLowerCase();
      if(MIME[ext]){const f=zip.file(resolve(path,prels[id]));if(f)imgs[id]='data:'+MIME[ext]+';base64,'+await f.async('base64');}
    }
    const on=v=>v!==null&&v!=='0'&&v!=='false'&&v!=='none';
    const rpr=run=>{
      const p=kid(run,'rPr');if(!p)return {tags:[],css:[]};const tags=[],css=[];
      const b=kid(p,'b'),i=kid(p,'i'),u=kid(p,'u'),s=kid(p,'strike'),col=kid(p,'color'),sz=kid(p,'sz'),f=kid(p,'rFonts'),va=kid(p,'vertAlign');
      if(b&&on(wa(b,'val')??'1'))tags.push('b');if(i&&on(wa(i,'val')??'1'))tags.push('i');
      if(u&&on(wa(u,'val')))tags.push('u');if(s&&on(wa(s,'val')??'1'))tags.push('s');
      if(col&&/^[0-9a-f]{6}$/i.test(wa(col,'val')||''))css.push('color:#'+wa(col,'val'));
      if(sz&&+wa(sz,'val'))css.push('font-size:'+(+wa(sz,'val')/2)+'pt');
      if(f&&(wa(f,'ascii')))css.push(`font-family:'${wa(f,'ascii').replace(/['"]/g,'')}'`);
      if(va&&wa(va,'val')==='superscript')tags.push('sup');if(va&&wa(va,'val')==='subscript')tags.push('sub');
      return {tags,css};
    };
    const wrap=(html,run)=>{
      if(!html)return html;const {tags,css}=rpr(run);
      if(css.length)html=`<span style="${css.join(';')}">${html}</span>`;
      tags.forEach(t=>{html=`<${t}>${html}</${t}>`;});return html;
    };
    const FLD=(instr,run)=>{
      if(/\bPAGE\b/i.test(instr)&&!/NUMPAGES|SECTIONPAGES/i.test(instr))return wrap('<span class="pn" contenteditable="false">#</span>',run);
      if(/NUMPAGES|SECTIONPAGES/i.test(instr))return wrap('<span class="pt" contenteditable="false">#</span>',run);
      return null;
    };
    function para(p){
      const segs=[''];let fld=null;
      const add=h=>{segs[segs.length-1]+=h;};
      const walk=el=>{
        for(const n of el.childNodes){
          if(n.nodeType!==1)continue;
          const ln=n.localName;
          if(ln==='r'){
            for(const c of n.childNodes){
              if(c.nodeType!==1)continue;const cn=c.localName;
              if(cn==='fldChar'){
                const ty=wa(c,'fldCharType');
                if(ty==='begin')fld={instr:'',skip:false,phase:'instr',run:n};
                else if(ty==='separate'&&fld){fld.phase='result';const h=FLD(fld.instr,fld.run);if(h){add(h);fld.skip=true;}}
                else if(ty==='end')fld=null;
              }else if(cn==='instrText'){if(fld)fld.instr+=c.textContent;}
              else if(fld&&(fld.phase==='instr'||fld.skip)){continue;}
              else if(cn==='t')add(wrap(esc(c.textContent),n));
              else if(cn==='tab')segs.push('');
              else if(cn==='br'&&wa(c,'type')!=='page')add('<br>');
              else if(cn==='drawing'||cn==='pict'){
                const blip=c.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main','blip')[0];
                const ext=c.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing','extent')[0];
                const rid=blip&&(blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','embed')||blip.getAttribute('r:embed'));
                if(rid&&imgs[rid]){const wpx=ext?Math.round(+ext.getAttribute('cx')/9525):0,hpx=ext?Math.round(+ext.getAttribute('cy')/9525):0;
                  add(`<img src="${imgs[rid]}"${wpx?` style="width:${wpx}px;height:${hpx}px"`:''}>`);}
              }
            }
          }else if(ln==='fldSimple'){
            const h=FLD(wa(n,'instr')||'',n.getElementsByTagNameNS(W_NS,'r')[0]||n);
            if(h)add(h);else walk(n);
          }else if(/^(hyperlink|sdt|sdtContent|smartTag|ins|customXml)$/.test(ln))walk(n);
        }
      };
      walk(p);
      const pPr=kid(p,'pPr'),jc=pPr&&kid(pPr,'jc'),al=jc?({center:'center',right:'right',end:'right',both:'justify'})[wa(jc,'val')]:null;
      const isEmpty=h=>!h.replace(/<[^>]*>/g,'').trim()&&!/<img|class="p[nt]"/.test(h);
      if(segs.length===1){
        if(isEmpty(segs[0]))return '<div><br></div>';
        return `<div${al?` style="text-align:${al}"`:''}>${segs[0]}</div>`;
      }
      if(segs.every(isEmpty))return '<div><br></div>';
      if(segs.length===2&&isEmpty(segs[0]))return `<div style="text-align:center">${segs[1]}</div>`;
      if(segs.length===2||segs.length===3){
        const a=segs.length===2?['left','right']:['left','center','right'];
        return `<div style="display:flex">${segs.map((s,i)=>`<span style="flex:1;text-align:${a[i]}">${s}</span>`).join('')}</div>`;
      }
      return `<div>${segs.join('&emsp;')}</div>`;
    }
    const lines=[...xml.getElementsByTagNameNS(W_NS,'p')].filter(p=>{for(let a=p.parentNode;a;a=a.parentNode)if(a.localName==='txbxContent')return false;return true;}).map(para);
    while(lines.length&&lines[lines.length-1]==='<div><br></div>')lines.pop();
    while(lines.length&&lines[0]==='<div><br></div>')lines.shift();
    return lines.join('');
  }
  const pick=name=>kids(sect,name).find(r=>(wa(r,'type')||'default')==='default');
  res.hdr=await partHTML(pick('headerReference'));
  res.ftr=await partHTML(pick('footerReference'));
  return res;
}

root.WordDocx={toBlob,readDocx};
})(typeof window!=='undefined'?window:globalThis);
