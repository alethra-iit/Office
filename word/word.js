/* Word editor (gaya Word 2013). Mengekspos Apps.word.mount(root,{file,onHome}) */
(function(){
const CDN={
  mammoth:'https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js',
  htmldocx:'https://cdn.jsdelivr.net/npm/html-docx-js@0.3.1/dist/html-docx.js'
};
const libs={};
function lib(n,test){
  if(test()) return Promise.resolve();
  return libs[n]||(libs[n]=new Promise((ok,err)=>{const s=document.createElement('script');s.src=CDN[n];s.onload=ok;s.onerror=()=>err(new Error('Gagal memuat '+n));document.head.appendChild(s);}));
}

const CSS=`
.wd{--blue:#2b579a;--ln:#d4d4d4;height:100%;display:flex;flex-direction:column;font:12px "Segoe UI",system-ui,sans-serif;color:#444;background:#fff;position:relative}
.wd *{box-sizing:border-box}
.wd button,.wd select,.wd input{font:inherit;color:inherit}
.wd [hidden]{display:none!important}
.wd em{font-style:normal}
.wd .qat{height:26px;display:flex;align-items:center;justify-content:center;position:relative;font-size:13px}
.wd .qat .l{position:absolute;left:6px;display:flex;gap:2px;align-items:center}
.wd .qat .w{background:var(--blue);color:#fff;font-weight:700;width:18px;height:16px;display:grid;place-items:center;font-size:11px;margin-right:6px}
.wd .qat button{background:none;border:0;cursor:pointer;width:24px;height:22px;color:#555}
.wd .qat button:hover{background:#e1e1e1}
.wd .tabs{display:flex;height:26px;border-bottom:1px solid var(--ln);background:#fff}
.wd .tab{background:none;border:1px solid transparent;border-bottom:0;padding:0 14px;cursor:pointer;height:26px;font-size:12px}
.wd .tab:hover{color:var(--blue)}
.wd .tab.on{border-color:var(--ln);background:#fff;color:var(--blue);height:27px}
.wd .tab.file{background:var(--blue);color:#fff;padding:0 22px;margin-right:6px;border:0}
.wd .ribbon{height:92px;background:#fff;border-bottom:1px solid var(--ln);display:flex;padding:3px 4px;overflow-x:auto;white-space:nowrap;flex:none}
.wd .grp{display:flex;flex-direction:column;border-right:1px solid #e3e3e3;padding:0 6px}
.wd .gb{flex:1;display:flex;gap:4px;align-items:flex-start}
.wd .gn{text-align:center;font-size:11px;color:#666;height:16px}
.wd .rows{display:flex;flex-direction:column;gap:4px}
.wd .row{display:flex;gap:2px;align-items:center;height:26px}
.wd .b{min-width:24px;height:24px;border:1px solid transparent;background:none;display:inline-flex;align-items:center;justify-content:center;gap:5px;padding:0 4px;cursor:pointer}
.wd .b:hover{background:#e6e6e6;border-color:#cfcfcf}
.wd .b.on{background:#cde6f7;border-color:#92c0e0}
.wd .b.big{flex-direction:column;height:62px;min-width:48px;gap:3px;font-size:11px}
.wd .b.big em{font-size:24px;line-height:1.1}
.wd .b.rw{height:22px;width:100%;justify-content:flex-start}
.wd select{height:23px;border:1px solid #ababab;background:#fff}
.wd .clr{position:relative;width:26px;height:24px;display:grid;place-items:center;cursor:pointer}
.wd .clr:hover{background:#e6e6e6}
.wd .clr span{padding:0 3px;line-height:1.1}
.wd .clr input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.wd .gal{display:flex;gap:2px;width:600px;height:64px;border:1px solid var(--ln);padding:2px;overflow:hidden}
.wd .sty{flex:none;width:86px;height:58px;background:none;border:1px solid transparent;cursor:pointer;display:flex;flex-direction:column;justify-content:flex-end;align-items:center;gap:4px;overflow:hidden}
.wd .sty:hover{border-color:#f5c35d;background:#fdf1c9}
.wd .sty span{font:11pt Calibri,Carlito,sans-serif;white-space:nowrap}
.wd .sty small{font-size:11px;color:#555;white-space:nowrap}
.wd .rbar{height:20px;background:#fff;border-bottom:1px solid var(--ln);display:flex;justify-content:center;overflow:hidden;flex:none}
.wd .ruler{position:relative;height:20px;flex:none;font-size:9px;color:#555;user-select:none}
.wd .ruler .rs{position:absolute;top:4px;height:12px;background:#dfe3e8}
.wd .ruler .mid{position:absolute;top:4px;height:12px;background:#fff;border:1px solid #cfd4da}
.wd .ruler .tk{position:absolute;top:8px;width:1px;height:4px;background:#888}
.wd .ruler b{position:absolute;top:3px;transform:translateX(-50%);font-weight:400}
.wd .ruler .mk{position:absolute;top:12px;width:0;height:0;border:5px solid transparent;border-bottom:7px solid #777;cursor:ew-resize}
.wd .desk{flex:1;overflow:auto;padding:12px 0 24px;background:linear-gradient(#f8f8f8,#ececec)}
.wd .sheet{width:max-content;min-width:100%;display:flex;justify-content:center}
.wd .paper{position:relative;background:#fff;border:1px solid #ababab;flex:none}
.wd .paper.bd{outline:2px solid #2b579a;outline-offset:-24px}
.wd .paper.wm::after{content:"DRAF";position:absolute;inset:0;display:grid;place-items:center;font:700 140px Calibri,sans-serif;color:rgba(0,0,0,.08);transform:rotate(-45deg);pointer-events:none}
.wd .cm{position:absolute;width:30px;height:30px;border:0 solid #bbb;pointer-events:none}
.wd .cm.a{left:calc(var(--ml) - 30px);top:calc(var(--mt) - 30px);border-width:0 1px 1px 0}
.wd .cm.b{left:calc(100% - var(--mr));top:calc(var(--mt) - 30px);border-width:0 0 1px 1px}
.wd .cm.c{left:calc(var(--ml) - 30px);top:calc(100% - var(--mb));border-width:1px 1px 0 0}
.wd .cm.d{left:calc(100% - var(--mr));top:calc(100% - var(--mb));border-width:1px 0 0 1px}
.wd .page{--hc:#2f5496;outline:none;font:11pt/1.08 Calibri,Carlito,"Segoe UI",sans-serif;color:#000;word-wrap:break-word;background-image:repeating-linear-gradient(to bottom,transparent 0,transparent calc(var(--ph) - 2px),#c9c9c9 calc(var(--ph) - 2px),#c9c9c9 var(--ph))}
.wd .page p{margin:0 0 8pt}
.wd .page h1,.wd .g1{font:400 16pt "Calibri Light",Carlito,Calibri,sans-serif;color:var(--hc);margin:12pt 0 0}
.wd .page h2,.wd .g2{font:400 13pt "Calibri Light",Carlito,Calibri,sans-serif;color:var(--hc);margin:2pt 0 0}
.wd .page h3{font:400 12pt "Calibri Light",Carlito,Calibri,sans-serif;color:#1f3763;margin:2pt 0 0}
.wd .s-title{font:28pt/1 "Calibri Light",Carlito,Calibri,sans-serif;margin:0 0 8pt}
.wd .s-sub{color:#5a5a5a;letter-spacing:.75pt}
.wd .s-em1{font-style:italic;color:#404040}.wd .s-em2{font-style:italic}
.wd .s-em3{font-style:italic;font-weight:700;color:var(--hc,#2f5496)}.wd .s-strong{font-weight:700}
.wd .page p.s-nosp{margin:0}
.wd .sty .s-title{font-size:20pt}.wd .sty .g1,.wd .sty .g2{margin:0}
.wd .page table{border-collapse:collapse;margin:0 0 8pt}
.wd .page td,.wd .page th{border:1px solid #000;padding:2pt 5.4pt;min-width:60px;vertical-align:top}
.wd .page img{max-width:100%}
.wd .page blockquote{margin:8pt 36pt;font-style:italic;color:#404040}
.wd .page .fnote{font-size:9pt;margin:0}
.wd .page.marks p::after,.wd .page.marks h1::after,.wd .page.marks h2::after,.wd .page.marks h3::after,.wd .page.marks li::after{content:"¶";color:#2b579a;font-size:10pt}
.wd .find{position:absolute;right:24px;top:150px;z-index:5;background:#fff;border:1px solid #aaa;box-shadow:0 2px 8px rgba(0,0,0,.25);padding:10px;display:grid;grid-template-columns:auto 170px;gap:6px;align-items:center}
.wd .find input[type=text]{height:24px;border:1px solid #aaa;padding:0 4px}
.wd .find .fb{grid-column:1/3;display:flex;gap:4px;justify-content:flex-end}
.wd .find button{height:24px;border:1px solid #aaa;background:#f3f3f3;cursor:pointer;padding:0 8px}
.wd .status{background:var(--blue);color:#fff;height:22px;display:flex;align-items:center;justify-content:space-between;padding:0 8px;font-size:11px;text-transform:uppercase;flex:none}
.wd .status .s{display:flex;gap:18px;align-items:center}
.wd .status button{background:none;border:0;color:#fff;cursor:pointer;width:20px;height:20px;font-size:13px;text-transform:none}
.wd .status button.on{background:rgba(255,255,255,.25)}
.wd .status input{width:130px}
.wd .menu{position:absolute;z-index:10;top:52px;left:0;bottom:22px;width:260px;background:#2b579a;color:#fff;padding:8px 0}
.wd .menu button{display:block;width:100%;background:none;border:0;color:#fff;text-align:left;padding:12px 24px;cursor:pointer;font-size:14px}
.wd .menu button:hover{background:rgba(255,255,255,.18)}
.wd .menu hr{border:0;border-top:1px solid rgba(255,255,255,.3);margin:6px 0}
.wd .toast{position:absolute;right:16px;bottom:36px;background:#323130;color:#fff;padding:8px 14px;z-index:20;max-width:360px}
@media print{body *{visibility:hidden}.wd .paper,.wd .paper *{visibility:visible}.wd .cm{display:none}.wd .paper{position:absolute;left:0;top:0;border:0;zoom:1!important}}
`;

const FONTS=['Calibri','Arial','Times New Roman','Cambria','Georgia','Verdana','Courier New','Segoe UI','Tahoma','Trebuchet MS'];
const SIZES=[8,9,10,11,12,14,16,18,20,24,28,36,48,72];
const SZ={a4:[794,1123],letter:[816,1056],legal:[816,1344]};
const MG={normal:[96,96,96,96],narrow:[48,48,48,48],moderate:[96,72,96,72],wide:[96,192,96,192]};
const STATE=['bold','italic','underline','strikeThrough','subscript','superscript','insertUnorderedList','insertOrderedList','justifyLeft','justifyCenter','justifyRight','justifyFull'];
const CMDS=new Set([...STATE,'outdent','indent','undo','redo','removeFormat','insertHorizontalRule','selectAll','cut','copy']);
const STY=[['Normal','p',''],['Tanpa Spasi','p','s-nosp'],['Judul 1','h1',''],['Judul 2','h2',''],['Judul','p','s-title'],['Subjudul','p','s-sub'],
  ['Penekanan Halus','i','s-em1'],['Penekanan','i','s-em2'],['Penekanan Kuat','i','s-em3'],['Tebal','i','s-strong']];
const TABS=[['home','BERANDA'],['insert','SISIPKAN'],['design','DESAIN'],['layout','TATA LETAK'],['ref','REFERENSI'],['review','TINJAU'],['view','TAMPILAN']];

const esc=t=>t.replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const B=(ic,tip,a,m)=>`<button class="b ${m||''}" title="${tip}" data-a="${a}"><em>${ic}</em>${m?`<span>${tip}</span>`:''}</button>`;
const S=(n,o,w,tip,r)=>`<select data-s="${n}" title="${tip||''}" style="width:${w}px"${r?' data-r':''}>${o.map(x=>Array.isArray(x)?`<option value="${x[0]}">${x[1]}</option>`:`<option>${x}</option>`).join('')}</select>`;
const C=(n,v,t,ic)=>`<label class="clr" title="${t}"><span style="border-bottom:3px solid ${v}">${ic}</span><input type="color" data-s="${n}" value="${v}"></label>`;
const G=(n,h)=>`<div class="grp"><div class="gb">${h}</div><div class="gn">${n}</div></div>`;
const R2=(a,b)=>`<div class="rows"><div class="row">${a}</div><div class="row">${b}</div></div>`;
const AL=t=>{const w=[14,9,14,9],x=i=>t==='L'||t==='J'?0:t==='C'?(14-w[i])/2:14-w[i];if(t==='J')w.fill(14);return `<svg width="16" height="14">${w.map((v,i)=>`<rect x="${1+x(i)}" y="${1+i*3.5}" width="${v}" height="1.6" fill="#444"/>`).join('')}</svg>`;};

const RIB={
home:
 G('Papan klip',B('📋','Tempel','paste','big')+`<div class="rows">${B('✂','Potong','cut','rw')}${B('⎘','Salin','copy','rw')}${B('🖌','Kuas Format','painter','rw')}</div>`)+
 G('Font',R2(
  S('font',FONTS,128,'Font')+S('size',SIZES,48,'Ukuran font')+B('A⁺','Perbesar font','grow')+B('A⁻','Perkecil font','shrink')+
  S('case',[['','Aa'],['upper','HURUF BESAR'],['lower','huruf kecil'],['title','Kapital Tiap Kata']],50,'Ubah huruf besar/kecil',1)+B('🧽','Hapus semua format','removeFormat'),
  B('<b>B</b>','Tebal (Ctrl+B)','bold')+B('<i>I</i>','Miring (Ctrl+I)','italic')+B('<u>U</u>','Garis bawah (Ctrl+U)','underline')+B('<s>abc</s>','Coret','strikeThrough')+
  B('x₂','Subskrip','subscript')+B('x²','Superskrip','superscript')+C('hilite','#ffff00','Warna sorot teks','🖍')+C('color','#ff0000','Warna font','<b>A</b>')))+
 G('Paragraf',R2(
  B('•≡','Poin','insertUnorderedList')+B('1≡','Penomoran','insertOrderedList')+B('⇤','Kurangi indentasi','outdent')+B('⇥','Tambah indentasi','indent')+B('¶','Tampilkan/sembunyikan ¶','marks'),
  B(AL('L'),'Rata kiri (Ctrl+L)','justifyLeft')+B(AL('C'),'Tengah (Ctrl+E)','justifyCenter')+B(AL('R'),'Rata kanan (Ctrl+R)','justifyRight')+B(AL('J'),'Rata kiri-kanan (Ctrl+J)','justifyFull')+
  S('lh',[['','↕ Spasi'],['1','1,0'],['1.15','1,15'],['1.5','1,5'],['2','2,0'],['3','3,0']],64,'Spasi baris',1)+C('shade','#d9d9d9','Pengarsiran paragraf','▧')))+
 G('Gaya',`<div class="gal">${STY.map((s,i)=>`<button class="sty" data-a="style:${i}"><span class="${s[2]||(s[1]==='h1'?'g1':s[1]==='h2'?'g2':'')}">AaBbCcDd</span><small>${s[0]}</small></button>`).join('')}</div>`)+
 G('Penyuntingan',`<div class="rows">${B('🔍','Cari','find','rw')}${B('ab','Ganti','replace','rw')}${B('⬚','Pilih semua','selectAll','rw')}</div>`),
insert:
 G('Halaman',B('⤓','Pemisah halaman','pagebreak','big'))+G('Tabel',B('▦','Tabel','table','big'))+
 G('Ilustrasi',B('🖼','Gambar','image','big'))+G('Tautan',B('🔗','Tautan','link','big'))+
 G('Simbol',B('―','Garis horizontal','insertHorizontalRule','big')+S('sym',[['','Ω Simbol'],'©','®','™','°','±','×','÷','€','£','→','←','✓'],80,'Sisipkan simbol',1))+
 G('Teks',B('📅','Tanggal & waktu','date','big')),
design:
 G('Pemformatan dokumen',S('theme',[['#2f5496','Tema: Biru'],['#385723','Tema: Hijau'],['#c00000','Tema: Merah'],['#7f7f7f','Tema: Abu-abu'],['#000000','Tema: Hitam']],110,'Warna judul'))+
 G('Latar halaman',B('◫','Watermark DRAF','wm','big')+C('pagecolor','#ffffff','Warna halaman','🎨')+B('▢','Batas halaman','border','big')),
layout:
 G('Penataan halaman',R2(S('margin',[['normal','Margin: Normal'],['narrow','Margin: Sempit'],['moderate','Margin: Sedang'],['wide','Margin: Lebar']],120,'Margin')+S('orient',[['p','Tegak'],['l','Mendatar']],80,'Orientasi'),
  S('pgsize',[['a4','A4'],['letter','Letter'],['legal','Legal']],120,'Ukuran kertas')+S('cols',[['1','1 Kolom'],['2','2 Kolom'],['3','3 Kolom']],80,'Kolom')))+
 G('Halaman',B('⤓','Pemisah halaman','pagebreak','big')),
ref:
 G('Daftar isi',B('📑','Daftar isi','toc','big'))+G('Catatan kaki',B('ab¹','Sisipkan catatan kaki','footnote','big')),
review:
 G('Pemeriksaan',B('✔','Ejaan','spell','big')+B('123','Hitung kata','wordcount','big'))+G('Mode',B('📖','Mode baca','read','big')),
view:
 G('Tampilan',B('📖','Mode baca','read','big'))+G('Tampilkan',B('📏','Penggaris','ruler','big')+B('¶','Tanda paragraf','marks','big'))+
 G('Zoom',B('🔍','100%','zoom:100','big')+B('🔎','75%','zoom:75','big')+B('🔬','150%','zoom:150','big'))
};

function mount(root,opts={}){
  if(!document.getElementById('wd-css')){const s=document.createElement('style');s.id='wd-css';s.textContent=CSS;document.head.appendChild(s);}
  let name='Dokumen1.docx',zoom=100,lastRange=null,painting=null;
  const cfg={size:'a4',land:false,m:[...MG.normal],cols:1};
  root.innerHTML=`<div class="wd">
    <div class="qat"><div class="l"><span class="w">W</span><button data-a="save" title="Simpan (Ctrl+S)">💾</button><button data-a="undo" title="Urungkan (Ctrl+Z)">↶</button><button data-a="redo" title="Ulangi (Ctrl+Y)">↷</button></div><span class="fn"></span></div>
    <div class="tabs"><button class="tab file" data-a="menu">BERKAS</button>${TABS.map((t,i)=>`<button class="tab${i?'':' on'}" data-tab="${t[0]}">${t[1]}</button>`).join('')}</div>
    ${TABS.map((t,i)=>`<div class="ribbon" data-pane="${t[0]}"${i?' hidden':''}>${RIB[t[0]]}</div>`).join('')}
    <div class="rbar"><div class="ruler"></div></div>
    <div class="desk"><div class="sheet"><div class="paper"><i class="cm a"></i><i class="cm b"></i><i class="cm c"></i><i class="cm d"></i>
      <div class="page" contenteditable="true" spellcheck="true"></div></div></div></div>
    <div class="find" hidden><span>Cari:</span><input type="text" class="f1"><span>Ganti:</span><input type="text" class="f2">
      <div class="fb"><button data-a="fnext">Berikutnya</button><button data-a="frep">Ganti</button><button data-a="frall">Ganti semua</button><button data-a="fclose">Tutup</button></div></div>
    <div class="status"><div class="s"><span class="pg">Halaman 1 dari 1</span><span class="wc">0 kata</span><span>Bahasa Indonesia</span></div>
      <div class="s"><button class="on" title="Tata letak cetak">▭</button><button data-a="read" title="Mode baca">📖</button>
      <span><button data-a="zoomd:-10">−</button><input type="range" min="40" max="200" value="100" class="zr"><button data-a="zoomd:10">+</button> <span class="zv">100%</span></span></div></div>
    <div class="menu" hidden>
      <button data-a="new">📄  Dokumen kosong baru</button><button data-a="open">📂  Buka…</button>
      <button data-a="save">💾  Simpan sebagai .docx</button><button data-a="print">🖨  Cetak / simpan PDF</button>
      <hr><button data-a="home">⬅  Kembali ke menu utama</button></div>
    <input type="file" class="fin" accept=".docx,.html,.htm,.txt" hidden><input type="file" class="img" accept="image/*" hidden>
  </div>`;

  const $=s=>root.querySelector(s),$$=s=>[...root.querySelectorAll(s)];
  const page=$('.page'),paper=$('.paper'),ruler=$('.ruler'),sheet=$('.sheet'),menu=$('.menu'),fin=$('.fin'),fd=$('.find');
  const toast=m=>{const t=document.createElement('div');t.className='toast';t.textContent=m;$('.wd').appendChild(t);setTimeout(()=>t.remove(),3000);};
  const setName=n=>{name=n;$('.fn').textContent=n.replace(/\.docx$/i,'')+' - Word';document.title=$('.fn').textContent;};
  const dims=()=>{const[w,h]=SZ[cfg.size];return cfg.land?[h,w]:[w,h];};

  /* seleksi & perintah */
  document.addEventListener('selectionchange',function sc(){
    if(!root.isConnected){document.removeEventListener('selectionchange',sc);return;}
    const s=getSelection(); if(s.rangeCount&&page.contains(s.anchorNode)) lastRange=s.getRangeAt(0).cloneRange();
  });
  function focus(){
    const s=getSelection(),inside=s.rangeCount&&page.contains(s.anchorNode);
    page.focus({preventScroll:true});
    if(!inside&&lastRange){s.removeAllRanges();s.addRange(lastRange);}
  }
  const exec=(c,v=null)=>{focus();document.execCommand(c,false,v);refresh();stats();};
  const curSize=()=>{const s=getSelection();let n=s.anchorNode;if(!n||!page.contains(n))return 11;n=n.nodeType===3?n.parentElement:n;return Math.round(parseFloat(getComputedStyle(n).fontSize)*.75);};
  function setSize(pt){
    focus();document.execCommand('styleWithCSS',false,false);document.execCommand('fontSize',false,7);document.execCommand('styleWithCSS',false,true);
    page.querySelectorAll('font[size="7"]').forEach(el=>{const s=document.createElement('span');s.style.fontSize=pt+'pt';s.innerHTML=el.innerHTML;el.replaceWith(s);});
    refresh();
  }
  function blocks(){
    focus();const s=getSelection();if(!s.rangeCount)return[];
    let bl=[...page.querySelectorAll('p,h1,h2,h3,li,blockquote')].filter(el=>s.containsNode(el,true));
    if(!bl.length){let n=s.anchorNode;n=n.nodeType===3?n.parentElement:n;const e=n.closest('p,h1,h2,h3,li,blockquote');if(e&&page.contains(e))bl=[e];}
    return bl;
  }
  function refresh(){
    STATE.forEach(c=>{const b=root.querySelector(`[data-a="${c}"]`);if(b)b.classList.toggle('on',document.queryCommandState(c));});
    const f=(document.queryCommandValue('fontName')||'').replace(/["']/g,'');
    if(FONTS.includes(f))$('[data-s=font]').value=f;
    const sz=curSize();if(SIZES.includes(sz))$('[data-s=size]').value=sz;
  }
  function stats(){
    const t=page.innerText.trim(),h=dims()[1];
    $('.wc').textContent=(t?t.split(/\s+/).length:0)+' kata';
    $('.pg').textContent='Halaman 1 dari '+Math.max(1,Math.ceil(page.offsetHeight/h));
  }

  /* tata letak halaman & penggaris */
  function drawRuler(){
    const[w]=dims(),[,mr,,ml]=cfg.m,u=37.795/2;
    let h=`<div class="rs" style="left:0;width:${ml}px"></div><div class="rs" style="right:0;width:${mr}px"></div><div class="mid" style="left:${ml}px;width:${w-ml-mr}px"></div>`;
    for(let i=-Math.ceil(ml/u);i<=Math.ceil((w-ml)/u);i++){const x=ml+i*u;if(x<0||x>w)continue;
      h+=i%2?`<i class="tk" style="left:${x}px"></i>`:(i?`<b style="left:${x}px">${Math.abs(i/2)}</b>`:`<b style="left:${x}px">0</b>`);}
    ruler.innerHTML=h+`<div class="mk" data-side="l" style="left:${ml-5}px"></div><div class="mk" data-side="r" style="left:${w-mr-5}px"></div>`;
    ruler.style.width=w+'px';
  }
  function apply(){
    const[w,h]=dims(),[mt,mr,mb,ml]=cfg.m;
    Object.assign(page.style,{width:w+'px',minHeight:h+'px',padding:`${mt}px ${mr}px ${mb}px ${ml}px`,columnCount:cfg.cols>1?cfg.cols:'auto',columnGap:'36px'});
    page.style.setProperty('--ph',h+'px');
    [['--mt',mt],['--mr',mr],['--mb',mb],['--ml',ml]].forEach(([k,v])=>paper.style.setProperty(k,v+'px'));
    drawRuler();stats();
  }
  ruler.addEventListener('mousedown',e=>{
    const mk=e.target.closest('.mk');if(!mk)return;e.preventDefault();
    const side=mk.dataset.side,left=ruler.getBoundingClientRect().left;
    const mv=ev=>{const x=(ev.clientX-left)/(zoom/100),[w]=dims();
      if(side==='l')cfg.m[3]=Math.round(Math.max(0,Math.min(x,w-cfg.m[1]-120)));else cfg.m[1]=Math.round(Math.max(0,Math.min(w-x,w-cfg.m[3]-120)));apply();};
    addEventListener('mousemove',mv);addEventListener('mouseup',()=>removeEventListener('mousemove',mv),{once:true});
  });
  function setZoom(v){zoom=v;sheet.style.zoom=ruler.style.zoom=v/100;$('.zr').value=v;$('.zv').textContent=v+'%';}

  /* berkas */
  const setContent=h=>{page.innerHTML=h||'<p><br></p>';stats();};
  async function openFile(file){
    const ext=file.name.split('.').pop().toLowerCase();
    try{
      if(ext==='docx'){await lib('mammoth',()=>window.mammoth);setContent((await mammoth.convertToHtml({arrayBuffer:await file.arrayBuffer()})).value);}
      else if(ext==='txt'){setContent((await file.text()).split(/\n/).map(l=>'<p>'+(esc(l)||'<br>')+'</p>').join(''));}
      else setContent(new DOMParser().parseFromString(await file.text(),'text/html').body.innerHTML);
      setName(file.name.replace(/\.(html?|txt)$/i,'.docx'));
    }catch(e){toast('Tidak bisa membuka berkas: '+e.message);}
  }
  async function saveDocx(){
    try{
      await lib('htmldocx',()=>window.htmlDocx);
      const css='body{font-family:Calibri;font-size:11pt}h1{font:16pt Calibri;color:#2f5496}h2{font:13pt Calibri;color:#2f5496}td,th{border:1px solid #000}.s-title{font-size:28pt}.s-em1,.s-em2{font-style:italic}.s-em3{font-style:italic;font-weight:bold;color:#2f5496}.s-strong{font-weight:bold}';
      const a=document.createElement('a');a.href=URL.createObjectURL(htmlDocx.asBlob(`<!DOCTYPE html><html><head><meta charset="utf-8"><style>${css}</style></head><body>${page.innerHTML}</body></html>`));
      a.download=/\.docx$/i.test(name)?name:name+'.docx';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);toast('Tersimpan: '+a.download);
    }catch(e){toast('Gagal menyimpan: '+e.message);}
  }

  /* cari & ganti */
  const f1=$('.f1'),f2=$('.f2');
  function findNext(){const t=f1.value;if(!t)return;
    for(let i=0;i<40;i++){if(!window.find(t,false,false,true)){toast('Tidak ditemukan');return;}const n=getSelection().anchorNode;if(n&&page.contains(n))return;}}
  function replaceAll(){const t=f1.value;if(!t)return;let c=0;const re=new RegExp(t.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'gi'),w=document.createTreeWalker(page,NodeFilter.SHOW_TEXT);
    for(let n;n=w.nextNode();)n.nodeValue=n.nodeValue.replace(re,()=>{c++;return f2.value;});toast(c+' penggantian dibuat');stats();}

  /* aksi */
  const showFind=()=>{fd.hidden=false;f1.focus();f1.select();};
  const acts={
    menu(){menu.hidden=!menu.hidden;},
    new(){setName('Dokumen1.docx');setContent('');focus();},
    open(){fin.click();},save:saveDocx,print(){window.print();},home(){opts.onHome&&opts.onHome();},
    async paste(){focus();try{document.execCommand('insertText',false,await navigator.clipboard.readText());}catch{toast('Gunakan Ctrl+V untuk menempel');}},
    painter(){painting={bold:document.queryCommandState('bold'),italic:document.queryCommandState('italic'),underline:document.queryCommandState('underline'),
      font:(document.queryCommandValue('fontName')||'Calibri').replace(/["']/g,''),size:curSize(),color:document.queryCommandValue('foreColor')};toast('Pilih teks untuk menerapkan format');},
    grow(){const n=SIZES.find(s=>s>curSize());if(n)setSize(n);},
    shrink(){const n=[...SIZES].reverse().find(s=>s<curSize());if(n)setSize(n);},
    marks(){page.classList.toggle('marks');},
    style(i){const[,k,c]=STY[i];
      if(k==='i'){const t=getSelection().toString();if(t)exec('insertHTML',`<span class="${c}">${esc(t)}</span>`);return;}
      exec('formatBlock',k==='p'?'P':k.toUpperCase());blocks().forEach(b=>b.className=c);},
    find:showFind,replace:showFind,fnext:findNext,fclose(){fd.hidden=true;},
    frep(){const s=getSelection();if(s.toString().toLowerCase()===f1.value.toLowerCase()&&page.contains(s.anchorNode)){focus();document.execCommand('insertText',false,f2.value);}findNext();},
    frall:replaceAll,
    pagebreak(){exec('insertHTML','<div class="pb" style="page-break-after:always;border-top:1px dashed #999;text-align:center;color:#999;font-size:9pt" contenteditable="false">Pemisah halaman</div><p><br></p>');},
    table(){const r=parseInt(prompt('Jumlah baris','3')),c=parseInt(prompt('Jumlah kolom','3'));if(!r||!c)return;
      exec('insertHTML','<table>'+('<tr>'+'<td><br></td>'.repeat(c)+'</tr>').repeat(r)+'</table><p><br></p>');},
    image(){$('.img').click();},
    link(){const u=prompt('Alamat tautan (URL)','https://');if(u)exec('createLink',u);},
    date(){exec('insertText',new Date().toLocaleString('id-ID',{day:'numeric',month:'long',year:'numeric',hour:'2-digit',minute:'2-digit'}));},
    wm(){paper.classList.toggle('wm');},border(){paper.classList.toggle('bd');},
    toc(){const hs=[...page.querySelectorAll('h1,h2,h3')];if(!hs.length)return toast('Belum ada Judul 1/2/3 di dokumen');
      exec('insertHTML','<p><b>Daftar Isi</b></p>'+hs.map(h=>`<p style="margin:0 0 4pt ${(h.tagName[1]-1)*20}px">${esc(h.textContent)}</p>`).join(''));},
    footnote(){const t=prompt('Teks catatan kaki');if(!t)return;const n=page.querySelectorAll('sup.fn').length+1;
      exec('insertHTML',`<sup class="fn">${n}</sup>`);page.insertAdjacentHTML('beforeend',(n===1?'<hr>':'')+`<p class="fnote"><sup>${n}</sup> ${esc(t)}</p>`);},
    spell(){page.spellcheck=!page.spellcheck;toast('Pemeriksa ejaan '+(page.spellcheck?'aktif':'nonaktif'));},
    wordcount(){const t=page.innerText;toast(`${t.trim()?t.trim().split(/\s+/).length:0} kata · ${t.replace(/\n/g,'').length} karakter · ${page.querySelectorAll('p,h1,h2,h3,li').length} paragraf`);},
    read(){const e=page.isContentEditable;page.contentEditable=e?'false':'true';toast(e?'Mode baca':'Mode edit');},
    ruler(){root.querySelector('.rbar').hidden=!root.querySelector('.rbar').hidden;},
    zoom(v){setZoom(+v);},zoomd(d){setZoom(Math.min(200,Math.max(40,zoom+ +d)));}
  };
  const caseOf=(t,m)=>m==='upper'?t.toUpperCase():m==='lower'?t.toLowerCase():t.toLowerCase().replace(/(^|\s)\S/g,c=>c.toUpperCase());
  const sels={
    font:v=>exec('fontName',v),size:v=>setSize(v),color:v=>exec('foreColor',v),hilite:v=>exec('hiliteColor',v),
    case(v){focus();const t=getSelection().toString();if(t&&v)exec('insertText',caseOf(t,v));},
    lh(v){if(v)blocks().forEach(b=>b.style.lineHeight=v);},
    shade(v){blocks().forEach(b=>b.style.background=v);},
    sym:v=>v&&exec('insertText',v),
    theme:v=>page.style.setProperty('--hc',v),pagecolor:v=>paper.style.background=v,
    margin(v){cfg.m=[...MG[v]];apply();},orient(v){cfg.land=v==='l';apply();},pgsize(v){cfg.size=v;apply();},cols(v){cfg.cols=+v;apply();}
  };

  /* event */
  root.addEventListener('mousedown',e=>{if(e.target.closest('.b,.sty,.tab,.status button'))e.preventDefault();});
  root.addEventListener('click',e=>{
    const t=e.target.closest('[data-a],[data-tab]');if(!t)return;
    if(t.dataset.tab){$$('.tab[data-tab]').forEach(x=>x.classList.toggle('on',x===t));$$('[data-pane]').forEach(p=>p.hidden=p.dataset.pane!==t.dataset.tab);return;}
    const[k,v]=t.dataset.a.split(':');
    if(CMDS.has(k))exec(k);else if(acts[k])acts[k](v);
    if(k!=='menu')menu.hidden=true;
  });
  const onSel=e=>{const s=e.target.dataset&&e.target.dataset.s;if(!s||(e.type==='input'&&e.target.tagName==='SELECT'))return;
    sels[s](e.target.value);if(e.target.tagName==='SELECT'&&e.target.hasAttribute('data-r'))e.target.selectedIndex=0;
    if(e.target.type==='color')e.target.previousElementSibling.style.borderBottomColor=e.target.value;};
  root.addEventListener('change',onSel);root.addEventListener('input',onSel);
  $('.desk').addEventListener('mousedown',()=>menu.hidden=true);
  $('.zr').oninput=e=>setZoom(+e.target.value);
  fin.onchange=()=>{if(fin.files[0])openFile(fin.files[0]);fin.value='';};
  $('.img').onchange=e=>{const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=()=>exec('insertImage',r.result);r.readAsDataURL(f);e.target.value='';};
  page.addEventListener('input',stats);
  page.addEventListener('keyup',refresh);
  page.addEventListener('mouseup',()=>{refresh();
    if(painting&&!getSelection().isCollapsed){const p=painting;painting=null;
      ['bold','italic','underline'].forEach(c=>{if(document.queryCommandState(c)!==p[c])exec(c);});
      exec('fontName',p.font);setSize(p.size);if(p.color)exec('foreColor',p.color);}});
  page.addEventListener('keydown',e=>{if(e.key==='Tab'){e.preventDefault();document.execCommand('insertHTML',false,'&emsp;');}});
  root.addEventListener('keydown',e=>{
    if(!(e.ctrlKey||e.metaKey))return;const k=e.key.toLowerCase();
    if(k==='s'){e.preventDefault();saveDocx();}else if(k==='f'||k==='h'){e.preventDefault();showFind();}});
  f1.addEventListener('keydown',e=>{if(e.key==='Enter')findNext();});

  document.execCommand('styleWithCSS',false,true);
  $('[data-s=font]').value='Calibri';$('[data-s=size]').value=11;
  setName(opts.file?opts.file.name:'Dokumen1.docx');
  setContent('');apply();
  if(opts.file)openFile(opts.file);else focus();
}

window.Apps=window.Apps||{};
Apps.word={mount};
})();
