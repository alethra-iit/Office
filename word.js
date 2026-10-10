/* Word editor — tampilan mirip Microsoft Word. Mengekspos Apps.word.mount(root,{file,onHome}) */
(function(){
const CDN={
  mammoth:'https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js',
  htmldocx:'https://cdn.jsdelivr.net/npm/html-docx-js@0.3.1/dist/html-docx.js'
};
const libs={};
function lib(name,test){
  if(test()) return Promise.resolve();
  return libs[name]||(libs[name]=new Promise((ok,err)=>{
    const s=document.createElement('script'); s.src=CDN[name]; s.onload=ok; s.onerror=()=>err(new Error('Gagal memuat '+name)); document.head.appendChild(s);
  }));
}

const CSS=`
.wd{--blue:#2b579a;--ribbon:#f3f2f1;--line:#d2d0ce;--hover:#e1dfdd;height:100%;display:flex;flex-direction:column;font:13px "Segoe UI",system-ui,sans-serif;color:#201f1e;background:#e6e6e6;position:relative}
.wd *{box-sizing:border-box}
.wd button,.wd select{font:inherit;color:inherit}
.wd .title{background:var(--blue);color:#fff;height:32px;display:flex;align-items:center;justify-content:center;position:relative;font-size:12px}
.wd .title .l{position:absolute;left:8px;display:flex;gap:2px}
.wd .title .l button{background:none;border:0;color:#fff;width:28px;height:26px;cursor:pointer;border-radius:2px;font-size:14px}
.wd .title .l button:hover{background:rgba(255,255,255,.2)}
.wd .tabs{background:var(--blue);display:flex;height:30px}
.wd .tab{background:none;border:0;color:#fff;padding:0 16px;cursor:pointer;height:30px}
.wd .tab:hover{background:rgba(255,255,255,.15)}
.wd .tab.on{background:var(--ribbon);color:var(--blue);border-radius:3px 3px 0 0}
.wd .tab.file{background:#1e3f73;margin-right:6px;font-weight:600}
.wd .ribbon{background:var(--ribbon);border-bottom:1px solid var(--line);height:86px;display:flex;align-items:stretch;padding:4px 6px;overflow-x:auto;white-space:nowrap}
.wd .grp{display:flex;flex-direction:column;justify-content:space-between;padding:0 8px;border-right:1px solid var(--line)}
.wd .grp .row{display:flex;gap:2px;align-items:center;height:28px}
.wd .grp .name{text-align:center;font-size:11px;color:#605e5c;height:16px}
.wd .grp .rows{display:flex;flex-direction:column;gap:4px;padding-top:4px}
.wd .b{min-width:28px;height:26px;border:1px solid transparent;background:none;border-radius:2px;cursor:pointer;padding:0 5px}
.wd .b:hover{background:var(--hover);border-color:var(--line)}
.wd .b.on{background:#c7e0f4;border-color:#8ab4dc}
.wd .big{flex-direction:column;height:60px;min-width:52px;display:flex;align-items:center;justify-content:center;gap:2px;font-size:11px}
.wd .big span{font-size:22px;line-height:1}
.wd select{height:26px;border:1px solid var(--line);background:#fff;border-radius:2px}
.wd input[type=color]{width:26px;height:22px;padding:0;border:1px solid var(--line);background:#fff;cursor:pointer}
.wd .styles{display:flex;gap:4px;align-items:center;height:60px}
.wd .sty{width:78px;height:56px;border:1px solid var(--line);background:#fff;cursor:pointer;text-align:left;padding:4px 6px;overflow:hidden}
.wd .sty:hover{border-color:var(--blue)}
.wd .desk{flex:1;overflow:auto;padding:24px 0}
.wd .zoomer{width:max-content;min-width:100%;display:flex;justify-content:center}
.wd .page{width:794px;min-height:1123px;background:#fff;padding:96px;box-shadow:0 1px 4px rgba(0,0,0,.35);outline:none;font:11pt/1.08 Calibri,Carlito,"Segoe UI",sans-serif;color:#000;word-wrap:break-word;flex:none}
.wd .page p{margin:0 0 8pt;line-height:1.08}
.wd .page h1{font:400 16pt "Calibri Light",Carlito,Calibri,sans-serif;color:#2f5496;margin:12pt 0 0}
.wd .page h2{font:400 13pt "Calibri Light",Carlito,Calibri,sans-serif;color:#2f5496;margin:2pt 0 0}
.wd .page h3{font:400 12pt "Calibri Light",Carlito,Calibri,sans-serif;color:#1f3763;margin:2pt 0 0}
.wd .page table{border-collapse:collapse;margin:0 0 8pt}
.wd .page td,.wd .page th{border:1px solid #000;padding:2pt 5.4pt;min-width:60px;vertical-align:top}
.wd .page img{max-width:100%}
.wd .page blockquote{margin:8pt 36pt;font-style:italic;color:#404040}
.wd .status{background:var(--blue);color:#fff;height:24px;display:flex;align-items:center;justify-content:space-between;padding:0 12px;font-size:12px}
.wd .status input{width:120px;vertical-align:middle}
.wd .menu{position:absolute;z-index:10;top:62px;left:0;bottom:24px;width:260px;background:#1e3f73;color:#fff;padding:8px 0}
.wd .menu button{display:block;width:100%;background:none;border:0;color:#fff;text-align:left;padding:12px 24px;cursor:pointer;font-size:14px}
.wd .menu button:hover{background:rgba(255,255,255,.15)}
.wd .menu hr{border:0;border-top:1px solid rgba(255,255,255,.25);margin:6px 0}
.wd .toast{position:absolute;right:16px;bottom:40px;background:#323130;color:#fff;padding:8px 14px;border-radius:2px}
.wd [hidden]{display:none!important}
@media print{body *{visibility:hidden}.wd .page,.wd .page *{visibility:visible}.wd .page{position:absolute;left:0;top:0;box-shadow:none;margin:0;zoom:1!important}}
`;

const FONTS=['Calibri','Arial','Times New Roman','Cambria','Georgia','Verdana','Courier New','Segoe UI','Tahoma','Trebuchet MS'];
const SIZES=[8,9,10,11,12,14,16,18,20,24,28,36,48,72];
const STATE=['bold','italic','underline','strikeThrough','subscript','superscript','insertUnorderedList','insertOrderedList','justifyLeft','justifyCenter','justifyRight','justifyFull'];

function mount(root,opts={}){
  if(!document.getElementById('wd-css')){const s=document.createElement('style');s.id='wd-css';s.textContent=CSS;document.head.appendChild(s);}
  let name='Dokumen1.docx';
  root.innerHTML=`
  <div class="wd">
    <div class="title">
      <div class="l"><button data-act="save" title="Simpan (Ctrl+S)">💾</button><button data-cmd="undo" title="Urungkan (Ctrl+Z)">↶</button><button data-cmd="redo" title="Ulangi (Ctrl+Y)">↷</button></div>
      <span class="fn"></span> <span style="opacity:.7;margin-left:6px">- Word</span>
    </div>
    <div class="tabs">
      <button class="tab file" data-act="menu">Berkas</button>
      <button class="tab on" data-tab="home">Beranda</button>
      <button class="tab" data-tab="insert">Sisipkan</button>
      <button class="tab" data-tab="view">Tampilan</button>
    </div>
    <div class="ribbon" data-pane="home">
      <div class="grp"><div class="row" style="height:60px"><button class="b big" data-act="paste"><span>📋</span>Tempel</button></div><div class="name">Papan klip</div></div>
      <div class="grp"><div class="rows">
        <div class="row"><select data-f="font" style="width:130px">${FONTS.map(f=>`<option>${f}</option>`).join('')}</select>
          <select data-f="size" style="width:52px">${SIZES.map(s=>`<option${s==11?' selected':''}>${s}</option>`).join('')}</select></div>
        <div class="row">
          <button class="b" data-cmd="bold" title="Tebal (Ctrl+B)"><b>B</b></button>
          <button class="b" data-cmd="italic" title="Miring (Ctrl+I)"><i>I</i></button>
          <button class="b" data-cmd="underline" title="Garis bawah (Ctrl+U)"><u>U</u></button>
          <button class="b" data-cmd="strikeThrough" title="Coret"><s>ab</s></button>
          <button class="b" data-cmd="subscript">x₂</button><button class="b" data-cmd="superscript">x²</button>
          <input type="color" data-f="color" value="#000000" title="Warna font">
          <input type="color" data-f="hilite" value="#ffff00" title="Sorot teks">
        </div></div><div class="name">Font</div></div>
      <div class="grp"><div class="rows">
        <div class="row"><button class="b" data-cmd="insertUnorderedList" title="Poin">• ≡</button><button class="b" data-cmd="insertOrderedList" title="Nomor">1. ≡</button>
          <button class="b" data-cmd="outdent" title="Kurangi indentasi">⇤</button><button class="b" data-cmd="indent" title="Tambah indentasi">⇥</button></div>
        <div class="row"><button class="b" data-cmd="justifyLeft" title="Rata kiri">⯇≡</button><button class="b" data-cmd="justifyCenter" title="Tengah">≡</button>
          <button class="b" data-cmd="justifyRight" title="Rata kanan">≡⯈</button><button class="b" data-cmd="justifyFull" title="Rata kanan-kiri">☰</button></div>
        </div><div class="name">Paragraf</div></div>
      <div class="grp"><div class="styles">
        <button class="sty" data-block="P" style="font:11pt Calibri">Normal</button>
        <button class="sty" data-block="H1" style="font:16pt 'Calibri Light',Calibri;color:#2f5496">Judul 1</button>
        <button class="sty" data-block="H2" style="font:13pt 'Calibri Light',Calibri;color:#2f5496">Judul 2</button>
        <button class="sty" data-block="H3" style="font:12pt 'Calibri Light',Calibri;color:#1f3763">Judul 3</button>
        <button class="sty" data-block="BLOCKQUOTE" style="font:italic 11pt Calibri">Kutipan</button>
      </div><div class="name">Gaya</div></div>
      <div class="grp"><div class="row" style="height:60px"><button class="b big" data-act="clear"><span>🧹</span>Hapus format</button></div><div class="name">Penyuntingan</div></div>
    </div>
    <div class="ribbon" data-pane="insert" hidden>
      <div class="grp"><div class="row" style="height:60px"><button class="b big" data-act="table"><span>▦</span>Tabel</button></div><div class="name">Tabel</div></div>
      <div class="grp"><div class="row" style="height:60px"><button class="b big" data-act="image"><span>🖼</span>Gambar</button><button class="b big" data-act="link"><span>🔗</span>Tautan</button></div><div class="name">Ilustrasi &amp; Tautan</div></div>
      <div class="grp"><div class="row" style="height:60px"><button class="b big" data-cmd="insertHorizontalRule"><span>―</span>Garis</button><button class="b big" data-act="date"><span>📅</span>Tanggal</button></div><div class="name">Teks</div></div>
    </div>
    <div class="ribbon" data-pane="view" hidden>
      <div class="grp"><div class="row" style="height:60px"><button class="b big" data-zoom="100"><span>🔍</span>100%</button><button class="b big" data-zoom="75"><span>🔎</span>75%</button></div><div class="name">Zoom</div></div>
      <div class="grp"><div class="row" style="height:60px"><button class="b big" data-act="print"><span>🖨</span>Cetak / PDF</button></div><div class="name">Cetak</div></div>
    </div>
    <div class="desk"><div class="zoomer"><div class="page" contenteditable="true" spellcheck="true"></div></div></div>
    <div class="status"><span class="wc">0 kata</span><span>Bahasa Indonesia</span>
      <span>Zoom <input type="range" min="40" max="200" value="100" class="zr"> <span class="zv">100%</span></span></div>
    <div class="menu" hidden>
      <button data-act="new">📄  Dokumen kosong baru</button>
      <button data-act="open">📂  Buka…</button>
      <button data-act="save">💾  Simpan sebagai .docx</button>
      <button data-act="print">🖨  Cetak / simpan PDF</button>
      <hr><button data-act="home">⬅  Kembali ke menu utama</button>
    </div>
    <input type="file" class="fin" accept=".docx,.html,.htm,.txt" hidden>
    <input type="file" class="img" accept="image/*" hidden>
  </div>`;

  const $=s=>root.querySelector(s), $$=s=>[...root.querySelectorAll(s)];
  const page=$('.page'), menu=$('.menu'), fin=$('.fin');
  let saved=null;

  const setName=n=>{name=n;$('.fn').textContent=n;document.title=n+' - Word';};
  const toast=m=>{const t=document.createElement('div');t.className='toast';t.textContent=m;$('.wd').appendChild(t);setTimeout(()=>t.remove(),2500);};
  const focus=()=>page.focus();
  const count=()=>{const t=page.innerText.trim();$('.wc').textContent=(t?t.split(/\s+/).length:0)+' kata';};
  function refresh(){
    $$('[data-cmd]').forEach(b=>{ if(STATE.includes(b.dataset.cmd)) b.classList.toggle('on',document.queryCommandState(b.dataset.cmd)); });
    const f=(document.queryCommandValue('fontName')||'').replace(/["']/g,'');
    if(FONTS.includes(f)) $('[data-f=font]').value=f;
  }
  const exec=(c,v=null)=>{focus();document.execCommand(c,false,v);refresh();count();};
  function setSize(pt){
    focus(); document.execCommand('fontSize',false,7);
    page.querySelectorAll('font[size="7"]').forEach(el=>{const s=document.createElement('span');s.style.fontSize=pt+'pt';s.innerHTML=el.innerHTML;el.replaceWith(s);});
  }
  const setContent=html=>{ page.innerHTML=html||'<p><br></p>'; count(); };

  async function openFile(file){
    const ext=file.name.split('.').pop().toLowerCase();
    try{
      if(ext==='docx'){
        await lib('mammoth',()=>window.mammoth);
        setContent((await mammoth.convertToHtml({arrayBuffer:await file.arrayBuffer()})).value);
      }else if(ext==='txt'){
        const t=await file.text();
        setContent(t.split(/\n/).map(l=>'<p>'+(l.replace(/&/g,'&amp;').replace(/</g,'&lt;')||'<br>')+'</p>').join(''));
      }else{
        setContent(new DOMParser().parseFromString(await file.text(),'text/html').body.innerHTML);
      }
      setName(file.name.replace(/\.(html?|txt)$/i,'.docx'));
    }catch(e){ toast('Tidak bisa membuka berkas: '+e.message); }
  }
  async function saveDocx(){
    try{
      await lib('htmldocx',()=>window.htmlDocx);
      const html='<!DOCTYPE html><html><head><meta charset="utf-8"><style>body{font-family:Calibri;font-size:11pt}h1{font-size:16pt;color:#2f5496}h2{font-size:13pt;color:#2f5496}td,th{border:1px solid #000}</style></head><body>'+page.innerHTML+'</body></html>';
      const a=document.createElement('a'); a.href=URL.createObjectURL(htmlDocx.asBlob(html));
      a.download=/\.docx$/i.test(name)?name:name+'.docx'; a.click();
      setTimeout(()=>URL.revokeObjectURL(a.href),1000); toast('Tersimpan: '+a.download);
    }catch(e){ toast('Gagal menyimpan: '+e.message); }
  }

  const actions={
    menu(){menu.hidden=!menu.hidden;},
    new(){setName('Dokumen1.docx');setContent('');focus();},
    open(){fin.click();},
    save:saveDocx,
    print(){window.print();},
    home(){opts.onHome&&opts.onHome();},
    async paste(){ focus(); try{ document.execCommand('insertText',false,await navigator.clipboard.readText()); }catch{ toast('Gunakan Ctrl+V untuk menempel'); } },
    clear(){exec('removeFormat');},
    table(){
      const r=parseInt(prompt('Jumlah baris','3')),c=parseInt(prompt('Jumlah kolom','3')); if(!r||!c) return;
      exec('insertHTML','<table>'+('<tr>'+'<td><br></td>'.repeat(c)+'</tr>').repeat(r)+'</table><p><br></p>');
    },
    image(){ const s=getSelection(); saved=s.rangeCount?s.getRangeAt(0):null; $('.img').click(); },
    link(){ const u=prompt('Alamat tautan (URL)','https://'); if(u) exec('createLink',u); },
    date(){ exec('insertText',new Date().toLocaleDateString('id-ID',{day:'numeric',month:'long',year:'numeric'})); }
  };

  root.addEventListener('mousedown',e=>{ if(e.target.closest('.b,.sty,.tab')) e.preventDefault(); });
  root.addEventListener('click',e=>{
    const t=e.target.closest('[data-cmd],[data-act],[data-block],[data-tab],[data-zoom]'); if(!t) return;
    if(t.dataset.cmd) exec(t.dataset.cmd);
    else if(t.dataset.block) exec('formatBlock',t.dataset.block);
    else if(t.dataset.zoom) setZoom(+t.dataset.zoom);
    else if(t.dataset.tab){ $$('.tab[data-tab]').forEach(x=>x.classList.toggle('on',x===t)); $$('[data-pane]').forEach(p=>p.hidden=p.dataset.pane!==t.dataset.tab); }
    else { actions[t.dataset.act](); if(t.dataset.act!=='menu') menu.hidden=true; }
  });
  $('.desk').addEventListener('mousedown',()=>menu.hidden=true);

  $('[data-f=font]').onchange=e=>exec('fontName',e.target.value);
  $('[data-f=size]').onchange=e=>setSize(e.target.value);
  $('[data-f=color]').oninput=e=>exec('foreColor',e.target.value);
  $('[data-f=hilite]').oninput=e=>exec('hiliteColor',e.target.value);

  function setZoom(z){ page.style.zoom=z/100; $('.zr').value=z; $('.zv').textContent=z+'%'; }
  $('.zr').oninput=e=>setZoom(+e.target.value);

  fin.onchange=()=>{ if(fin.files[0]) openFile(fin.files[0]); fin.value=''; };
  $('.img').onchange=e=>{
    const f=e.target.files[0]; if(!f) return;
    const r=new FileReader();
    r.onload=()=>{ focus(); if(saved){const s=getSelection();s.removeAllRanges();s.addRange(saved);} document.execCommand('insertImage',false,r.result); };
    r.readAsDataURL(f); e.target.value='';
  };

  page.addEventListener('input',count);
  page.addEventListener('keyup',refresh); page.addEventListener('mouseup',refresh);
  page.addEventListener('keydown',e=>{ if(e.key==='Tab'){e.preventDefault();document.execCommand('insertHTML',false,'&emsp;');} });
  root.addEventListener('keydown',e=>{ if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='s'){e.preventDefault();saveDocx();} });

  document.execCommand('styleWithCSS',false,true);
  setName(opts.file?opts.file.name:'Dokumen1.docx');
  setContent('');
  if(opts.file) openFile(opts.file); else focus();
}

window.Apps=window.Apps||{};
Apps.word={mount};
})();
