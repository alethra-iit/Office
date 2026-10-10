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
.wd{--blue:#2b579a;--ln:#d4d4d4;--hf:"Calibri Light",Carlito,Calibri,sans-serif;--bf:Calibri,Carlito,"Segoe UI",sans-serif;height:100%;display:flex;flex-direction:column;font:12px "Segoe UI",system-ui,sans-serif;color:#444;background:#fff;position:relative}
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
.wd .paper.bd{outline-offset:-24px}
.wd .paper.wm::after{content:var(--wmt,"DRAF");position:absolute;inset:0;display:grid;place-items:center;font:700 var(--wms,110px) Calibri,sans-serif;color:rgba(0,0,0,.08);transform:rotate(-45deg);pointer-events:none}
.wd .cm{position:absolute;width:30px;height:30px;border:0 solid #bbb;pointer-events:none}
.wd .cm.a{left:calc(var(--ml) - 30px);top:calc(var(--mt) - 30px);border-width:0 1px 1px 0}
.wd .cm.b{left:calc(100% - var(--mr));top:calc(var(--mt) - 30px);border-width:0 0 1px 1px}
.wd .cm.c{left:calc(var(--ml) - 30px);top:calc(100% - var(--mb));border-width:1px 1px 0 0}
.wd .cm.d{left:calc(100% - var(--mr));top:calc(100% - var(--mb));border-width:1px 0 0 1px}
.wd .page{--hc:#2f5496;outline:none;font:11pt/var(--lh,1.08) var(--bf);color:#000;word-wrap:break-word}
.wd .page.col{background-image:repeating-linear-gradient(to bottom,transparent 0,transparent calc(var(--ph) - 2px),#c9c9c9 calc(var(--ph) - 2px),#c9c9c9 var(--ph))}
.wd .pgap{user-select:none;pointer-events:none;display:flow-root}
.wd .hfl{position:absolute;inset:0;pointer-events:none}
.wd .hf{position:absolute;pointer-events:auto;font:10pt/1.2 var(--bf);color:#555;outline:none;overflow:hidden;min-height:18px;cursor:text;word-wrap:break-word}
.wd .hf:hover{outline:1px dashed #9ab}
.wd .hf:focus{outline:1px dashed #2b579a;background:rgba(43,87,154,.06);color:#000}
.wd .hf:empty:not(:focus):hover::before{content:attr(data-ph);color:#999}
@media print{.wd .hfl{display:none}}
@media print{.wd .pgap{height:0!important;margin:0!important;background:none!important;break-after:page}}
.wd .page p{margin:0 0 var(--ps,8pt)}
.wd .page h1,.wd .g1{font:400 16pt var(--hf);color:var(--hc);margin:12pt 0 0}
.wd .page h2,.wd .g2{font:400 13pt var(--hf);color:var(--hc);margin:2pt 0 0}
.wd .page h3{font:400 12pt var(--hf);color:#1f3763;margin:2pt 0 0}
.wd .s-title{font:28pt/1 var(--hf);margin:0 0 8pt}
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
.wd .page a{color:#0563c1}
.wd .page .toc a,.wd .page .tof a,.wd .page .bib a{color:inherit;text-decoration:none}
.wd .page.fx1 :is(h1,h2,h3){text-shadow:1px 1px 2px rgba(0,0,0,.3)}
.wd .page.fx2 :is(h1,h2){border-bottom:1px solid var(--hc)}
.wd .page.ln{counter-reset:ln}
.wd .page.ln :is(p,h1,h2,h3,li){position:relative;counter-increment:ln}
.wd .page.ln :is(p,h1,h2,h3,li)::before{content:counter(ln);position:absolute;left:-34px;width:26px;text-align:right;font:9pt Calibri,sans-serif;color:#777}
.wd .page img.sel,.wd .page table.sel,.wd .page hr.sel{outline:2px solid #2b579a}
.wd .page.marks .xe::after{content:" XE";color:#2b579a;font-size:8pt}
.wd .page.marks .ta::after{content:" TA";color:#7030a0;font-size:8pt}
.wd .page .xe,.wd .page .ta{border-bottom:1px dotted #999}
.wd input[type=number]{height:23px;border:1px solid #ababab;width:58px}
.wd .nl{display:flex;align-items:center;gap:3px}
.wd .ov{position:absolute;inset:0;background:rgba(0,0,0,.25);z-index:30;display:grid;place-items:center}
.wd .dg{background:#fff;border:1px solid #888;box-shadow:0 4px 18px rgba(0,0,0,.3);padding:14px 16px;min-width:340px;max-width:92%}
.wd .dg h3{margin:0 0 10px;font-size:14px;font-weight:600;color:var(--blue)}
.wd .db{display:flex;flex-direction:column;gap:8px;max-height:60vh;overflow:auto}
.wd .db input,.wd .db select{height:26px;border:1px solid #aaa;padding:0 4px}
.wd .db select[size]{height:auto}
.wd .dbt{display:flex;gap:6px;justify-content:flex-end;margin-top:12px}
.wd .dbt button,.wd .db button{height:26px;border:1px solid #aaa;background:#f3f3f3;padding:0 12px;cursor:pointer}
.wd .dbt .ok{background:var(--blue);color:#fff;border-color:var(--blue)}
.wd .sli{display:flex;justify-content:space-between;gap:8px;align-items:center;border-bottom:1px solid #eee;padding:3px 0}
.wd .sp{position:absolute;right:24px;top:150px;z-index:5;width:220px;background:#fff;border:1px solid #aaa;box-shadow:0 2px 8px rgba(0,0,0,.25);padding:8px}
.wd .spi{display:flex;justify-content:space-between;align-items:center;padding:3px 6px;cursor:pointer}
.wd .spi:hover{background:#e6f1fb}.wd .spi.on{background:#cde6f7}
.wd .spi button{background:none;border:0;cursor:pointer}
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

const N=(n,v,t,st,l)=>`<label class="nl" title="${t}">${l}<input type="number" data-s="${n}" value="${v}" step="${st}"></label>`;
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
 G('Teks',B('📅','Tanggal & waktu','date','big'))+
 G('Header & Footer',R2(S('hdrs',[['','Header ▾'],['blank','Header kosong'],['title','Judul dokumen'],['three','Tiga kolom'],['rm','Hapus header']],118,'Header',1)+S('ftrs',[['','Footer ▾'],['blank','Footer kosong'],['title','Judul dokumen'],['three','Tiga kolom'],['rm','Hapus footer']],118,'Footer',1),
  S('pnum',[['','Nomor Halaman ▾'],['bc','Bawah – tengah'],['br','Bawah – kanan'],['bl','Bawah – kiri'],['tr','Atas – kanan'],['pc','Halaman X dari Y'],['rm','Hapus nomor halaman']],140,'Nomor halaman',1))),
design:
 G('Pemformatan dokumen',R2(
  S('thm',[['','Tema: Office'],['ion','Tema: Ion'],['facet','Tema: Facet'],['retro','Tema: Retrospect'],['serif','Tema: Serif'],['mono','Tema: Mono']],110,'Tema')+
  S('hcol',[['#2f5496','Warna: Biru'],['#385723','Warna: Hijau'],['#c00000','Warna: Merah'],['#c55a11','Warna: Oranye'],['#7030a0','Warna: Ungu'],['#7f7f7f','Warna: Abu-abu'],['#000000','Warna: Hitam']],110,'Warna tema'),
  S('fpair',[['Calibri Light|Calibri','Font: Office'],['Arial|Arial','Font: Arial'],['Cambria|Calibri','Font: Cambria-Calibri'],['Georgia|Verdana','Font: Georgia-Verdana'],['Times New Roman|Times New Roman','Font: Times New Roman'],['Trebuchet MS|Segoe UI','Font: Trebuchet-Segoe']],110,'Font tema')+
  S('pspace',[['','Spasi: Default'],['none','Tanpa spasi'],['compact','Kompak'],['tight','Ketat'],['open','Terbuka'],['relaxed','Santai'],['double','Dobel']],110,'Spasi paragraf')))+
 G('Efek',`<div class="rows">${S('fx',[['','Efek: Tanpa'],['fx1','Efek: Bayangan'],['fx2','Efek: Garis bawah']],110,'Efek judul')}${B('✔','Atur sebagai default','setdef','rw')}</div>`)+
 G('Latar halaman',R2(
  S('wmk',[['','Watermark: Tanpa'],['DRAF','DRAF'],['RAHASIA','RAHASIA'],['CONTOH','CONTOH'],['JANGAN DISALIN','JANGAN DISALIN'],['*','Kustom…']],118,'Watermark')+C('pagecolor','#ffffff','Warna halaman','🎨'),
  S('bsty',[['','Batas: Tanpa'],['solid','Garis'],['double','Ganda'],['dashed','Putus-putus'],['groove','Alur']],100,'Batas halaman')+S('bw',[['1','1 px'],['2','2 px'],['4','4 px'],['6','6 px']],54,'Tebal batas')+C('bcol','#2b579a','Warna batas','▢'))),
layout:
 G('Penataan halaman',R2(
  S('margin',[['normal','Margin: Normal'],['narrow','Margin: Sempit'],['moderate','Margin: Sedang'],['wide','Margin: Lebar']],120,'Margin')+S('orient',[['p','Tegak'],['l','Mendatar']],80,'Orientasi')+S('pgsize',[['a4','A4'],['letter','Letter'],['legal','Legal']],70,'Ukuran kertas')+S('cols',[['1','1 Kolom'],['2','2 Kolom'],['3','3 Kolom']],90,'Kolom'),
  S('dir',[['ltr','Arah: Kiri-Kanan'],['rtl','Arah: Kanan-Kiri']],120,'Arah teks')+S('brk',[['','Pemisah…'],['p','Halaman'],['c','Kolom']],90,'Pemisah',1)+S('ln',[['','Nomor baris: Tidak'],['1','Nomor baris: Kontinu']],130,'Nomor baris')+S('hyph',[['','Tanpa tanda hubung'],['1','Hubung otomatis']],120,'Pemenggalan kata')))+
 G('Paragraf',R2(N('indl',0,'Indentasi kiri (cm)','0.5','Kiri (cm)')+N('indr',0,'Indentasi kanan (cm)','0.5','Kanan (cm)'),N('spb',0,'Spasi sebelum (pt)','6','Sebelum (pt)')+N('spa',8,'Spasi sesudah (pt)','6','Sesudah (pt)')))+
 G('Susun (pilih gambar/tabel)',R2(
  S('pos',[['','Posisi'],['i','Sebaris'],['l','Kiri (membungkus)'],['c','Tengah'],['r','Kanan (membungkus)']],118,'Posisi objek',1)+S('wrap',[['','Bungkus Teks'],['in','Sebaris dengan teks'],['sq','Persegi'],['tb','Atas dan bawah']],118,'Bungkus teks',1)+B('▲','Majukan','fwd')+B('▼','Mundurkan','bwd'),
  B('☰ Panel Seleksi','Panel Seleksi','selpane')+S('algn',[['','Rata'],['l','Kiri'],['c','Tengah'],['r','Kanan']],70,'Ratakan objek / paragraf',1)+S('rot',[['','Putar'],['cw','Kanan 90°'],['ccw','Kiri 90°'],['fh','Balik horizontal'],['fv','Balik vertikal']],90,'Putar objek',1))),
ref:
 G('Daftar isi',B('📑','Daftar isi','toc','big')+`<div class="rows">${S('addtxt',[['','Tambah Teks'],['h1','Level 1'],['h2','Level 2'],['h3','Level 3'],['p','Tidak ditampilkan']],96,'Tambah Teks',1)}${B('⟳','Perbarui Tabel','tocup','rw')}</div>`)+
 G('Catatan kaki',B('ab¹','Sisipkan catatan kaki','footnote','big')+`<div class="rows">${B('ⅰ Sisipkan Catatan Akhir','Sisipkan catatan akhir','endnote','rw')}${B('↓ Catatan Kaki Berikutnya','Catatan kaki berikutnya','nextfn','rw')}${B('👁 Tampilkan Catatan','Tampilkan catatan','shownotes','rw')}</div>`)+
 G('Kutipan & Daftar Pustaka',B('❝','Sisipkan kutipan','cite','big')+`<div class="rows">${B('📚 Kelola Sumber','Kelola sumber','sources','rw')}<div class="row"><span>Gaya:</span>${S('cstyle',['APA','MLA','Chicago','IEEE'],76,'Gaya sitasi')}</div>${B('📖 Daftar Pustaka','Daftar pustaka','bib','rw')}</div>`)+
 G('Keterangan',B('🏷','Sisipkan keterangan','caption','big')+`<div class="rows">${B('▤ Sisipkan Tabel Gambar','Sisipkan tabel gambar','tof','rw')}${B('⟳ Perbarui Tabel','Perbarui tabel','capup','rw')}${B('⌗ Referensi Silang','Referensi silang','xref','rw')}</div>`)+
 G('Indeks',B('✚','Tandai entri','markxe','big')+`<div class="rows">${B('☷ Sisipkan Indeks','Sisipkan indeks','idx','rw')}${B('⟳ Perbarui Indeks','Perbarui indeks','idxup','rw')}</div>`)+
 G('Daftar kutipan hukum',B('⚖','Tandai kutipan','markta','big')+`<div class="rows">${B('☷ Sisipkan Daftar Kutipan','Sisipkan daftar kutipan','toa','rw')}${B('⟳ Perbarui Tabel','Perbarui tabel','toaup','rw')}</div>`),

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
    <div class="desk"><div class="sheet"><div class="paper"><div class="hfl"></div>
      <div class="page" contenteditable="true" spellcheck="true"></div></div></div></div>
    <div class="find" hidden><span>Cari:</span><input type="text" class="f1"><span>Ganti:</span><input type="text" class="f2">
      <div class="fb"><button data-a="fnext">Berikutnya</button><button data-a="frep">Ganti</button><button data-a="frall">Ganti semua</button><button data-a="fclose">Tutup</button></div></div>
    <div class="sp" hidden></div>
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
    const s=getSelection();if(s.rangeCount&&root.querySelector('.hfl').contains(s.anchorNode))return;
    const inside=s.rangeCount&&page.contains(s.anchorNode);
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
    let bl=[...page.querySelectorAll('p,h1,h2,h3,li,blockquote')].filter(el=>s.containsNode(el,true)&&!el.classList.contains('pgap'));
    if(!bl.length){let n=s.anchorNode;n=n.nodeType===3?n.parentElement:n;const e=n.closest('p,h1,h2,h3,li,blockquote');if(e&&page.contains(e))bl=[e];}
    return bl;
  }
  function refresh(){
    STATE.forEach(c=>{const b=root.querySelector(`[data-a="${c}"]`);if(b)b.classList.toggle('on',document.queryCommandState(c));});
    const f=(document.queryCommandValue('fontName')||'').replace(/["']/g,'');
    if(FONTS.includes(f))$('[data-s=font]').value=f;
    const sz=curSize();if(SIZES.includes(sz))$('[data-s=size]').value=sz;
    syncNums();
  }
  function stats(){
    const t=page.innerText.trim(),h=dims()[1];
    $('.wc').textContent=(t?t.split(/\s+/).length:0)+' kata';
    $('.pg').textContent='Halaman 1 dari '+(cfg.cols>1?Math.max(1,Math.ceil(page.offsetHeight/h)):pgCount);
    if(!paging)sched();
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
    page.style.setProperty('--ph',h+'px');page.classList.toggle('col',cfg.cols>1);
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
      pick(null);await lib('htmldocx',()=>window.htmlDocx);
      const css=`body{font-family:"${D.b}";font-size:11pt}h1{font:16pt "${D.h}";color:${D.hc}}h2{font:13pt "${D.h}";color:${D.hc}}td,th{border:1px solid #000}.s-title{font-size:28pt}.s-em1,.s-em2{font-style:italic}.s-em3{font-style:italic;font-weight:bold;color:${D.hc}}.s-strong{font-weight:bold}`;
      const a=document.createElement('a');a.href=URL.createObjectURL(htmlDocx.asBlob(`<!DOCTYPE html><html><head><meta charset="utf-8"><style>${css}</style></head><body>${cleanHTML()}</body></html>`));
      a.download=/\.docx$/i.test(name)?name:name+'.docx';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);toast('Tersimpan: '+a.download);
    }catch(e){toast('Gagal menyimpan: '+e.message);}
  }

  /* cari & ganti */
  const f1=$('.f1'),f2=$('.f2');
  function findNext(){const t=f1.value;if(!t)return;
    for(let i=0;i<40;i++){if(!window.find(t,false,false,true)){toast('Tidak ditemukan');return;}const n=getSelection().anchorNode;if(n&&page.contains(n))return;}}
  function replaceAll(){const t=f1.value;if(!t)return;let c=0;const re=new RegExp(t.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'gi'),w=document.createTreeWalker(page,NodeFilter.SHOW_TEXT);
    for(let n;n=w.nextNode();)n.nodeValue=n.nodeValue.replace(re,()=>{c++;return f2.value;});toast(c+' penggantian dibuat');stats();}


/* ===== Paginasi: tiap halaman terpisah, ukuran sama ===== */
const PG_GAP=16;let pgCount=1,paging=false,pgT=0;
const cleanHTML=()=>{const c=page.cloneNode(true);c.querySelectorAll('.pgap').forEach(x=>x.remove());return c.innerHTML;};
function sched(){clearTimeout(pgT);pgT=setTimeout(paginate,120);}
function paginate(){
  if(!root.isConnected)return;
  paging=true;
  page.querySelectorAll('.pgap').forEach(x=>x.remove());
  const h=dims()[1],mt=cfg.m[0],mr=cfg.m[1],mb=cfg.m[2],ml=cfg.m[3],G=PG_GAP,T=h+G,ch=h-mt-mb;
  if(cfg.cols>1){page.style.minHeight=h+'px';pgCount=Math.max(1,Math.ceil(page.offsetHeight/h));paging=false;renderHF(pgCount,0);return;}
  const sc=()=>page.offsetHeight?page.getBoundingClientRect().height/page.offsetHeight:1;
  const Y=el=>(el.getBoundingClientRect().top-page.getBoundingClientRect().top)/sc(),H=el=>el.getBoundingClientRect().height/sc();
  const units=[];[...page.children].forEach(c=>{if(/^(UL|OL)$/.test(c.tagName))units.push(...c.children);else units.push(c);});
  let lastBot=mt;
  for(const el of units){
    const ht=H(el);if(!ht)continue;
    const top=Y(el),k=Math.floor(top/T);
    if(top+ht>k*T+h-mb&&ht<=ch&&top>k*T+mt+2){
      const sp=document.createElement(el.parentNode===page?'div':'li');
      sp.className='pgap';sp.contentEditable='false';sp.style.display='flow-root';el.before(sp);
      const y0=Y(sp),hh=(k+1)*T+mt-Y(el),a=Math.max(0,k*T+h-y0);
      if(hh>0)sp.style.cssText=`display:flow-root;height:${hh}px;`+(sp.tagName==='DIV'?`margin:0 -${mr}px 0 -${ml}px;`:'list-style:none;')+
        `background:linear-gradient(to bottom,transparent ${a}px,#aaa ${a}px,#aaa ${a+1}px,#ececec ${a+1}px,#ececec ${a+G-1}px,#aaa ${a+G-1}px,#aaa ${a+G}px,transparent ${a+G}px)`;
      else sp.remove();
    }
    lastBot=Math.max(lastBot,Y(el)+ht);
  }
  pgCount=Math.max(1,Math.ceil((lastBot+mb+G)/T));
  page.style.minHeight=(pgCount*T-G)+'px';
  paging=false;renderHF(pgCount,G);
  $('.pg').textContent='Halaman 1 dari '+pgCount;
}
page.addEventListener('load',sched,true);
if(document.fonts&&document.fonts.ready)document.fonts.ready.then(sched);

/* ===== Header, Footer, Nomor Halaman, tanda sudut tiap halaman ===== */
let hdr='',ftr='';
const pnHTML=(html,k,n)=>{const t=document.createElement('div');t.innerHTML=html;t.querySelectorAll('.pn').forEach(x=>x.textContent=k+1);t.querySelectorAll('.pt').forEach(x=>x.textContent=n);return t.innerHTML;};
function renderHF(n,G){
  const L=$('.hfl');if(!L||L.contains(document.activeElement))return;
  const[w,h]=dims(),[mt,mr,mb,ml]=cfg.m;let o='';
  const cm=(x,y,bw)=>`<i class="cm" style="left:${x}px;top:${y}px;border-width:${bw}"></i>`;
  for(let k=0;k<n;k++){
    const b=k*(h+G),bx=`left:${ml}px;width:${Math.max(40,w-ml-mr)}px;`;
    o+=cm(ml-30,b+mt-30,'0 1px 1px 0')+cm(w-mr,b+mt-30,'0 0 1px 1px')+cm(ml-30,b+h-mb,'1px 1px 0 0')+cm(w-mr,b+h-mb,'1px 0 0 1px');
    o+=`<div class="hf" data-t="h" contenteditable="true" data-ph="Header" style="${bx}top:${b+36}px;max-height:${Math.max(20,mt-44)}px">${pnHTML(hdr,k,n)}</div>`;
    o+=`<div class="hf" data-t="f" contenteditable="true" data-ph="Footer" style="${bx}top:${b+h-mb+14}px;max-height:${Math.max(20,mb-20)}px">${pnHTML(ftr,k,n)}</div>`;
  }
  L.innerHTML=o;
}
$('.hfl').addEventListener('input',e=>{
  const el=e.target.closest('.hf');if(!el)return;
  const c=el.cloneNode(true);c.querySelectorAll('.pn,.pt').forEach(x=>x.textContent='#');
  const html=c.innerHTML==='<br>'?'':c.innerHTML;if(el.dataset.t==='h')hdr=html;else ftr=html;
  [...$('.hfl').querySelectorAll(`.hf[data-t="${el.dataset.t}"]`)].forEach((o,k)=>{if(o!==el)o.innerHTML=pnHTML(html,k,pgCount);});
});
const HFT={blank:'<div><br></div>',title:()=>`<div style="text-align:center">${esc(name.replace(/\.docx$/i,''))}</div>`,three:'<div style="display:flex;justify-content:space-between"><span>[Kiri]</span><span>[Tengah]</span><span>[Kanan]</span></div>'};
const redoHF=()=>renderHF(pgCount,cfg.cols>1?0:PG_GAP);
function setHF(t,v){
  if(!v)return;const html=v==='rm'?'':typeof HFT[v]==='function'?HFT[v]():HFT[v];
  if(t==='h')hdr=html;else ftr=html;redoHF();
  const el=$(`.hf[data-t="${t}"]`);if(el&&html){el.scrollIntoView({block:'center'});el.focus();}
}
const PNM={bc:['f','center'],br:['f','right'],bl:['f','left'],tr:['h','right'],pc:['f','center',1]};
function pageNum(v){
  if(!v)return;
  const strip=html=>{const d=document.createElement('div');d.innerHTML=html;[...d.children].forEach(c=>{if(c.querySelector('.pn')||c.classList.contains('pn'))c.remove();});return d.innerHTML;};
  if(v==='rm'){hdr=strip(hdr);ftr=strip(ftr);}
  else{const[t,al,tot]=PNM[v],pn='<span class="pn" contenteditable="false">#</span>',
    row=`<div style="text-align:${al}">${tot?`Halaman ${pn} dari <span class="pt" contenteditable="false">#</span>`:pn}</div>`;
    if(t==='h')hdr=strip(hdr)+row;else ftr=strip(ftr)+row;}
  redoHF();
}

/* ===== Desain, Tata Letak, Referensi ===== */
const store={get(k,d){try{const v=localStorage.getItem('wd.'+k);return v===null?d:JSON.parse(v);}catch{return d;}},set(k,v){try{localStorage.setItem('wd.'+k,JSON.stringify(v));}catch{}}};
const D={h:'Calibri Light',b:'Calibri',hc:'#2f5496',lh:1.08,ps:8,fx:''};
const THM={'':['Calibri Light','Calibri','#2f5496',1.08],ion:['Trebuchet MS','Verdana','#c55a11',1.15],facet:['Trebuchet MS','Trebuchet MS','#385723',1.15],retro:['Georgia','Georgia','#7030a0',1.2],serif:['Times New Roman','Times New Roman','#000000',1.15],mono:['Courier New','Courier New','#7f7f7f',1.1]};
const PSP={'':[8,1.08],none:[0,1],compact:[4,1],tight:[6,1.15],open:[10,1.3],relaxed:[8,1.5],double:[8,2]};
const TAC=['Kasus','Undang-Undang','Peraturan','Lainnya'],CAPL=['Gambar','Tabel','Persamaan'];
let selObj=null,cstyle=store.get('cs','APA');const src=store.get('src',[]);
const setSel=(n,v)=>{const s=$(`[data-s=${n}]`);if(s){s.value=v;if(s.selectedIndex<0)s.selectedIndex=0;}};
const roman=n=>{let r='';for(const[v,s]of[[10,'x'],[9,'ix'],[5,'v'],[4,'iv'],[1,'i']])while(n>=v){r+=s;n-=v;}return r;};
const tid=e=>{if(!e.id)e.id=(e.classList.contains('cap')?'cap':'h')+Math.random().toString(36).slice(2,8);return e.id;};
const pageOf=el=>{const r=el.getBoundingClientRect(),p=page.getBoundingClientRect(),k=page.offsetHeight?p.height/page.offsetHeight:1;return Math.max(1,Math.floor((r.top-p.top)/k/(dims()[1]+(cfg.cols>1?0:PG_GAP)))+1);};
const attr=t=>esc(t).replace(/"/g,'&quot;');

function dlg(title,body,ok,okText){
  const d=document.createElement('div');d.className='ov';
  d.innerHTML=`<div class="dg"><h3>${title}</h3><div class="db">${body}</div><div class="dbt"><button class="x">Batal</button><button class="ok">${okText||'OK'}</button></div></div>`;
  $('.wd').appendChild(d);d.querySelector('.x').onclick=()=>d.remove();
  d.querySelector('.ok').onclick=()=>{if(ok(d)!==false)d.remove();};return d;
}

/* Desain */
function applyDesign(d){
  Object.assign(D,d);const w=$('.wd');
  w.style.setProperty('--hf',`"${D.h}",Carlito,sans-serif`);w.style.setProperty('--bf',`"${D.b}",Carlito,"Segoe UI",sans-serif`);
  page.style.setProperty('--hc',D.hc);page.style.setProperty('--lh',D.lh);page.style.setProperty('--ps',D.ps+'pt');
  page.classList.toggle('fx1',D.fx==='fx1');page.classList.toggle('fx2',D.fx==='fx2');
  setSel('hcol',D.hc);setSel('fpair',D.h+'|'+D.b);setSel('fx',D.fx);
  const pk=Object.keys(PSP).find(k=>PSP[k][0]===+D.ps&&PSP[k][1]===+D.lh);setSel('pspace',pk===undefined?'':pk);
}
function border(){
  const s=$('[data-s=bsty]').value,w=+$('[data-s=bw]').value,c=$('[data-s=bcol]').value;
  paper.classList.toggle('bd',!!s);paper.style.outline=s?`${s==='double'&&w<3?3:w}px ${s} ${c}`:'';
}

/* Tata letak */
function curBlock(){
  const s=getSelection();let n=s.anchorNode;if(!n||!page.contains(n))return null;
  n=n.nodeType===3?n.parentElement:n;const e=n.closest('p,h1,h2,h3,li,blockquote');return e&&page.contains(e)?e:null;
}
function syncNums(){
  const n=curBlock();if(!n)return;const c=getComputedStyle(n);
  const m={indl:c.marginLeft,indr:c.marginRight,spb:c.marginTop,spa:c.marginBottom};
  for(const k in m){const i=$(`[data-s=${k}]`);if(i&&document.activeElement!==i)i.value=k[0]==='i'?+(parseFloat(m[k])/37.795).toFixed(1):Math.round(parseFloat(m[k])*.75);}
}
function pick(t){if(selObj)selObj.classList.remove('sel');selObj=t;if(t)t.classList.add('sel');if(!$('.sp').hidden)drawSP();}
const need=()=>selObj||(toast('Pilih gambar atau tabel terlebih dahulu'),null);
const objs=()=>[...page.querySelectorAll('img,table,hr')];
function drawSP(){
  const c={IMG:0,TABLE:0,HR:0},nm={IMG:'Gambar',TABLE:'Tabel',HR:'Garis'};
  $('.sp').innerHTML='<b>Panel Seleksi</b>'+(objs().map((o,i)=>`<div class="spi${o===selObj?' on':''}" data-i="${i}"><span>${nm[o.tagName]} ${++c[o.tagName]}</span><button data-v="${i}" title="Tampilkan/sembunyikan">${o.style.visibility==='hidden'?'🚫':'👁'}</button></div>`).join('')||'<div><i>Tidak ada objek</i></div>');
}
function selpane(){const p=$('.sp');p.hidden=!p.hidden;if(!p.hidden)drawSP();}
function z(d){const o=need();if(!o)return;o.style.position='relative';const v=Math.max(0,(+o.style.zIndex||0)+d);o.style.zIndex=v;toast('Urutan tumpukan: '+v);}
$('.sp').addEventListener('click',e=>{
  const os=objs(),b=e.target.closest('button[data-v]');
  if(b){const o=os[+b.dataset.v];o.style.visibility=o.style.visibility==='hidden'?'':'hidden';drawSP();return;}
  const r=e.target.closest('.spi');if(r){const o=os[+r.dataset.i];pick(o);o.scrollIntoView({block:'center'});}
});
page.addEventListener('click',e=>{
  const a=e.target.closest('a[href^="#"]');
  if(a){e.preventDefault();const t=page.querySelector('#'+CSS.escape(a.getAttribute('href').slice(1)));if(t)t.scrollIntoView({block:'center'});return;}
  const o=e.target.closest('img,table,hr');pick(o&&page.contains(o)?o:null);
});

/* Daftar isi, catatan, keterangan, indeks */
const tocHTML=()=>{const hs=[...page.querySelectorAll('h1,h2,h3')];
  return '<p><b>Daftar Isi</b></p>'+(hs.length?hs.map(h=>`<p style="margin:0 0 4pt ${(h.tagName[1]-1)*20}px"><a href="#${tid(h)}">${esc(h.textContent)}</a><span style="float:right">${pageOf(h)}</span></p>`).join(''):'<p><i>Belum ada judul</i></p>');};
function renum(){
  const c={};
  page.querySelectorAll('p.cap').forEach(p=>{const l=p.dataset.l;c[l]=(c[l]||0)+1;tid(p);p.querySelector('b').textContent=l+' '+c[l];});
  page.querySelectorAll('a.xr').forEach(a=>{const t=page.querySelector('#'+CSS.escape(a.dataset.t));if(t)a.textContent=t.classList.contains('cap')?t.querySelector('b').textContent:t.textContent;});
}
const tofHTML=l=>`<p><b>Daftar ${l}</b></p>`+([...page.querySelectorAll(`p.cap[data-l="${l}"]`)].map(p=>`<p><a href="#${tid(p)}">${esc(p.textContent)}</a><span style="float:right">${pageOf(p)}</span></p>`).join('')||`<p><i>Belum ada keterangan ${l}</i></p>`);
const idxHTML=()=>{
  const m={};page.querySelectorAll('.xe').forEach(x=>{(m[x.dataset.xe]=m[x.dataset.xe]||new Set()).add(pageOf(x));});
  const k=Object.keys(m).sort((a,b)=>a.localeCompare(b,'id'));let L='';
  return '<p><b>Indeks</b></p>'+(k.length?k.map(e=>{const l=e[0].toUpperCase(),hd=l!==L?(L=l,`<p style="margin:6pt 0 2pt"><b>${esc(l)}</b></p>`):'';return hd+`<p style="margin:0 0 2pt 12px">${esc(e)}, ${[...m[e]].sort((a,b)=>a-b).join(', ')}</p>`;}).join(''):'<p><i>Belum ada entri yang ditandai</i></p>');
};
const toaHTML=()=>{
  const m={};page.querySelectorAll('.ta').forEach(x=>{const g=m[x.dataset.cat]=m[x.dataset.cat]||{},e=x.dataset.ta;(g[e]=g[e]||new Set()).add(pageOf(x));});
  const cs=TAC.filter(c=>m[c]);
  return '<p><b>Daftar Kutipan Hukum</b></p>'+(cs.length?cs.map(c=>`<p style="margin:6pt 0 2pt"><b>${c}</b></p>`+Object.keys(m[c]).sort((a,b)=>a.localeCompare(b,'id')).map(e=>`<p style="margin:0 0 2pt 12px">${esc(e)}, ${[...m[c][e]].sort((a,b)=>a-b).join(', ')}</p>`).join('')).join(''):'<p><i>Belum ada kutipan yang ditandai</i></p>');
};
function updAll(){
  renum();
  page.querySelectorAll('.toc').forEach(t=>t.innerHTML=tocHTML());
  page.querySelectorAll('.tof').forEach(t=>t.innerHTML=tofHTML(t.dataset.l));
  page.querySelectorAll('.idx').forEach(t=>t.innerHTML=idxHTML());
  page.querySelectorAll('.toa').forEach(t=>t.innerHTML=toaHTML());
  refreshCites();toast('Tabel diperbarui');
}
function note(k){
  const t=prompt(k==='fn'?'Teks catatan kaki':'Teks catatan akhir');if(!t)return;
  const n=page.querySelectorAll('sup.'+k).length+1,lab=k==='fn'?n:roman(n);
  exec('insertHTML',`<sup class="${k}">${lab}</sup>`);
  let box=page.querySelector('.'+k+'s');
  if(!box){box=document.createElement('div');box.className=k+'s';box.innerHTML='<hr>'+(k==='en'?'<p class="fnote"><b>Catatan Akhir</b></p>':'');
    const en=page.querySelector('.ens');if(k==='fn'&&en)en.before(box);else page.appendChild(box);}
  box.insertAdjacentHTML('beforeend',`<p class="fnote"><sup>${lab}</sup> ${esc(t)}</p>`);
}
function nextfn(){
  const all=[...page.querySelectorAll('sup.fn,sup.en')];if(!all.length)return toast('Belum ada catatan');
  const s=getSelection(),nx=(lastRange&&all.find(x=>lastRange.comparePoint(x,0)>0))||all[0];
  nx.scrollIntoView({block:'center'});const r=document.createRange();r.selectNode(nx);s.removeAllRanges();s.addRange(r);
}
function shownotes(){const b=page.querySelector('.fns')||page.querySelector('.ens');b?b.scrollIntoView({block:'center'}):toast('Belum ada catatan');}
function caption(){
  dlg('Sisipkan Keterangan',`<label>Label <select class="cl">${CAPL.map(l=>`<option>${l}</option>`).join('')}</select></label><label>Teks <input class="ct" style="width:240px"></label>`,d=>{
    const l=d.querySelector('.cl').value,t=d.querySelector('.ct').value;
    const h=`<p class="cap" data-l="${l}" style="font-size:9pt;color:#44546a;font-style:italic"><b>${l}</b>${t?': '+esc(t):''}</p>`;
    if(selObj)(selObj.closest('p')||selObj).insertAdjacentHTML('afterend',h);else exec('insertHTML',h);
    renum();
  });
}
function tof(){
  dlg('Sisipkan Tabel Gambar',`<label>Label <select class="cl">${CAPL.map(l=>`<option>${l}</option>`).join('')}</select></label>`,d=>{
    const l=d.querySelector('.cl').value;renum();exec('insertHTML',`<div class="tof" data-l="${l}">${tofHTML(l)}</div><p><br></p>`);
  });
}
function xref(){
  const items=[...page.querySelectorAll('h1,h2,h3,p.cap')];if(!items.length)return toast('Belum ada judul atau keterangan');renum();
  dlg('Referensi Silang',`<select class="xs" size="8" style="width:320px">${items.map((e,i)=>`<option value="${i}">${e.classList.contains('cap')?esc(e.textContent):'Judul: '+esc(e.textContent)}</option>`).join('')}</select>`,d=>{
    const e=items[d.querySelector('.xs').value];if(!e)return false;
    exec('insertHTML',`<a class="xr" href="#${tid(e)}" data-t="${tid(e)}">${esc(e.classList.contains('cap')?e.querySelector('b').textContent:e.textContent)}</a>&nbsp;`);
  });
}
function markxe(){
  focus();const t=getSelection().toString().trim();if(!t)return toast('Pilih teks yang akan ditandai');
  const e=prompt('Entri indeks',t);if(!e)return;
  exec('insertHTML',`<span class="xe" data-xe="${attr(e)}">${esc(t)}</span>`);
}
function idx(){exec('insertHTML',`<div class="idx">${idxHTML()}</div><p><br></p>`);}
function markta(){
  focus();const t=getSelection().toString().trim();if(!t)return toast('Pilih teks kutipan yang akan ditandai');
  const d=dlg('Tandai Kutipan',`<label>Kutipan <input class="c" style="width:260px"></label><label>Kategori <select class="k">${TAC.map(k=>`<option>${k}</option>`).join('')}</select></label>`,d=>{
    exec('insertHTML',`<span class="ta" data-ta="${attr(d.querySelector('.c').value||t)}" data-cat="${d.querySelector('.k').value}">${esc(t)}</span>`);
  });
  d.querySelector('.c').value=t;
}
function toa(){exec('insertHTML',`<div class="toa">${toaHTML()}</div><p><br></p>`);}

/* Kutipan & daftar pustaka */
const A=s=>s.author||'Anonim',last=a=>a.includes(',')?a.split(',')[0].trim():a.trim().split(/\s+/).pop();
const fmtCite=(s,i)=>cstyle==='IEEE'?`[${i+1}]`:cstyle==='MLA'?`(${last(A(s))})`:cstyle==='Chicago'?`(${last(A(s))} ${s.year||'t.t.'})`:`(${last(A(s))}, ${s.year||'t.t.'})`;
const bibEntry=(s,i)=>{const a=esc(A(s)),y=esc(s.year||'t.t.'),ti=esc(s.title||''),p=esc(s.pub||'');
  return cstyle==='MLA'?`${a}. <i>${ti}</i>. ${p?p+', ':''}${y}.`:cstyle==='Chicago'?`${a}. ${y}. <i>${ti}</i>. ${p?p+'.':''}`:cstyle==='IEEE'?`[${i+1}] ${a}, “${ti},” ${p?p+', ':''}${y}.`:`${a}. (${y}). <i>${ti}</i>. ${p?p+'.':''}`;};
const bibHTML=()=>'<p><b>Daftar Pustaka</b></p>'+(src.length?src.map((s,i)=>[s,i]).sort((x,y)=>cstyle==='IEEE'?x[1]-y[1]:A(x[0]).localeCompare(A(y[0]),'id')).map(([s,i])=>`<p style="padding-left:36px;text-indent:-36px">${bibEntry(s,i)}</p>`).join(''):'<p><i>Belum ada sumber</i></p>');
function refreshCites(){
  page.querySelectorAll('.cite').forEach(c=>{const i=src.findIndex(s=>s.id===c.dataset.src);if(i>=0)c.textContent=fmtCite(src[i],i);});
  const b=page.querySelector('.bib');if(b)b.innerHTML=bibHTML();
}
function sources(){
  const d=dlg('Kelola Sumber',`<div class="sl"></div><input class="a" placeholder="Pengarang (Belakang, Depan)"><input class="y" placeholder="Tahun"><input class="t" placeholder="Judul"><input class="p" placeholder="Penerbit / situs"><button class="add">Tambah sumber</button>`,()=>{refreshCites();},'Selesai');
  const q=s=>d.querySelector(s),draw=()=>{q('.sl').innerHTML=src.length?src.map((s,i)=>`<div class="sli"><span>${esc(A(s))} (${esc(s.year||'t.t.')}) — ${esc(s.title||'')}</span><button data-d="${i}">✕</button></div>`).join(''):'<i>Belum ada sumber</i>';};
  q('.add').onclick=()=>{const s={id:'s'+Date.now().toString(36),author:q('.a').value.trim(),year:q('.y').value.trim(),title:q('.t').value.trim(),pub:q('.p').value.trim()};
    if(!s.title&&!s.author)return;src.push(s);store.set('src',src);['.a','.y','.t','.p'].forEach(k=>q(k).value='');draw();};
  q('.sl').onclick=e=>{const i=e.target.dataset.d;if(i!==undefined){src.splice(+i,1);store.set('src',src);draw();}};
  draw();
}
function cite(){
  if(!src.length){toast('Tambahkan sumber terlebih dahulu');return sources();}
  dlg('Sisipkan Kutipan',`<select class="ss" size="6">${src.map((s,i)=>`<option value="${i}"${i?'':' selected'}>${esc(A(s))} (${esc(s.year||'t.t.')}) — ${esc(s.title||'')}</option>`).join('')}</select>`,d=>{
    const i=+d.querySelector('.ss').value,s=src[i];if(!s)return false;
    exec('insertHTML',`<span class="cite" data-src="${s.id}">${fmtCite(s,i)}</span>&nbsp;`);
  });
}
function bib(){
  if(page.querySelector('.bib')){refreshCites();return toast('Daftar pustaka diperbarui');}
  exec('insertHTML',`<div class="bib">${bibHTML()}</div><p><br></p>`);
}


  /* aksi */
  const showFind=()=>{fd.hidden=false;f1.focus();f1.select();};
  const acts={
    menu(){menu.hidden=!menu.hidden;},
    new(){setName('Dokumen1.docx');hdr=ftr='';setContent('');focus();},
    open(){fin.click();},save:saveDocx,print(){pick(null);window.print();},home(){opts.onHome&&opts.onHome();},
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
    toc(){if(!page.querySelector('h1,h2,h3'))return toast('Belum ada Judul 1/2/3 di dokumen');exec('insertHTML',`<div class="toc">${tocHTML()}</div><p><br></p>`);},
    tocup:updAll,capup:updAll,idxup:updAll,toaup:updAll,
    footnote(){note('fn');},endnote(){note('en');},nextfn,shownotes,
    cite,sources,bib,caption,tof,xref,markxe,idx,markta,toa,selpane,
    fwd(){z(1);},bwd(){z(-1);},
    setdef(){store.set('design',D);toast('Format disimpan sebagai default');},
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
    thm(v){const t=THM[v];applyDesign({h:t[0],b:t[1],hc:t[2],lh:t[3],ps:8});},
    hcol:v=>applyDesign({hc:v}),fpair(v){const[h,b]=v.split('|');applyDesign({h,b});},
    pspace(v){const p=PSP[v];applyDesign({ps:p[0],lh:p[1]});},fx:v=>applyDesign({fx:v}),
    pagecolor:v=>paper.style.background=v,
    wmk(v){if(v==='*')v=prompt('Teks watermark','RAHASIA')||'';paper.classList.toggle('wm',!!v);
      if(v){paper.style.setProperty('--wmt',JSON.stringify(v));paper.style.setProperty('--wms',Math.round(Math.min(140,760/(v.length*.6)))+'px');}},
    bsty:border,bw:border,bcol:border,
    dir:v=>{page.dir=v;},ln:v=>page.classList.toggle('ln',v==='1'),hyph(v){page.lang='id';page.style.hyphens=v?'auto':'manual';},
    brk(v){if(v==='p')acts.pagebreak();else if(v==='c')exec('insertHTML','<div class="cb" style="break-after:column;border-top:1px dashed #999;font-size:9pt;color:#999;text-align:center" contenteditable="false">Pemisah kolom</div><p><br></p>');},
    indl:v=>blocks().forEach(b=>b.style.marginLeft=(+v||0)+'cm'),indr:v=>blocks().forEach(b=>b.style.marginRight=(+v||0)+'cm'),
    spb:v=>blocks().forEach(b=>b.style.marginTop=(+v||0)+'pt'),spa:v=>blocks().forEach(b=>b.style.marginBottom=(+v||0)+'pt'),
    pos(v){const o=need();if(!o||!v)return;const st=o.style;st.float='';st.margin='';if(o.tagName==='IMG')st.display='';
      if(v==='l'){st.float='left';st.margin='0 12px 8px 0';}else if(v==='r'){st.float='right';st.margin='0 0 8px 12px';}else if(v==='c'){if(o.tagName==='IMG')st.display='block';st.margin='0 auto';}},
    wrap(v){const o=need();if(!o||!v)return;const st=o.style;
      if(v==='in'){st.float='';st.margin='';st.clear='';if(o.tagName==='IMG')st.display='';}
      else if(v==='sq'){if(!st.float||st.float==='none'){st.float='left';st.margin='0 12px 8px 0';}}
      else{st.float='none';st.clear='both';st.margin='8px 0';if(o.tagName==='IMG')st.display='block';}},
    algn(v){if(!v)return;const o=selObj;if(!o)return exec(v==='c'?'justifyCenter':v==='r'?'justifyRight':'justifyLeft');
      const st=o.style;st.float='';if(o.tagName==='IMG')st.display='block';st.marginLeft=v==='l'?'0':'auto';st.marginRight=v==='r'?'0':'auto';},
    rot(v){if(!v)return;const o=need();if(!o)return;const d=o.dataset;let r=+d.r||0,fx=+d.fx||1,fy=+d.fy||1;
      if(v==='cw')r+=90;else if(v==='ccw')r-=90;else if(v==='fh')fx*=-1;else fy*=-1;
      Object.assign(d,{r,fx,fy});o.style.transform=`rotate(${r}deg) scale(${fx},${fy})`;},
    hdrs:v=>setHF('h',v),ftrs:v=>setHF('f',v),pnum:pageNum,
    addtxt(v){if(v){exec('formatBlock',v==='p'?'P':v.toUpperCase());blocks().forEach(b=>b.className='');}},
    cstyle(v){cstyle=v;store.set('cs',v);refreshCites();},
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
  const onSel=e=>{const s=e.target.dataset&&e.target.dataset.s;if(!s||(e.type==='input'&&(e.target.tagName==='SELECT'||e.target.type==='number')))return;
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
  applyDesign(store.get('design',{}));setSel('cstyle',cstyle);
  if(opts.file)openFile(opts.file);else focus();
}

window.Apps=window.Apps||{};
Apps.word={mount};
})();
