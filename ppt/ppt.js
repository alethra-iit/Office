/* PowerPoint editor (gaya PowerPoint 2013) — Apps.ppt.mount(root,{file,onHome})
   Tahap 1: tab BERANDA (Papan klip, Slide, Font, Paragraf, Gambar, Penyuntingan) + panel slide + kanvas slide.
   Buka/Simpan .pptx: lihat pptx-io.js (window.PptxIO). */
(function(){
const W=1280,H=720;
const CDN={
  jszip:'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
  pptx:'https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundle.js',
  pptxio:'ppt/pptx-io.js'
};
const libs={};
function lib(n,test){
  if(test()) return Promise.resolve();
  return libs[n]||(libs[n]=new Promise((ok,err)=>{const s=document.createElement('script');s.src=CDN[n];s.onload=ok;s.onerror=()=>{delete libs[n];err(new Error('Gagal memuat '+n));};document.head.appendChild(s);}));
}
const CSS=`
.pp{--acc:#d24726;--ln:#d4d4d4;height:100%;display:flex;flex-direction:column;font:12px "Segoe UI",system-ui,sans-serif;color:#444;background:#fff;position:relative;user-select:none;-webkit-user-select:none}
.pp *{box-sizing:border-box}
.pp button,.pp select,.pp input{font:inherit;color:inherit}
.pp [hidden]{display:none!important}
.pp em{font-style:normal}
.pp .tx,.pp .notes,.pp input[type=text],.pp input[type=number]{user-select:text;-webkit-user-select:text}
.pp .qat{height:26px;display:flex;align-items:center;justify-content:center;position:relative;font-size:13px}
.pp .qat .l{position:absolute;left:6px;display:flex;gap:2px;align-items:center}
.pp .qat .w{background:var(--acc);color:#fff;font-weight:700;width:18px;height:16px;display:grid;place-items:center;font-size:11px;margin-right:6px}
.pp .qat button{background:none;border:0;cursor:pointer;width:24px;height:22px;color:#555}
.pp .qat button:hover{background:#e1e1e1}
.pp .qat .fn{max-width:calc(100% - 240px);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.pp .tabs{display:flex;height:26px;border-bottom:1px solid var(--ln);background:#fff;overflow-x:auto;scrollbar-width:none}
.pp .tabs::-webkit-scrollbar{display:none}
.pp .tab{flex:none;white-space:nowrap;background:none;border:1px solid transparent;border-bottom:0;padding:0 14px;cursor:pointer;height:26px;font-size:12px}
.pp .tab:hover{color:var(--acc)}
.pp .tab.on{border-color:var(--ln);background:#fff;color:var(--acc);height:27px}
.pp .tab.file{background:var(--acc);color:#fff;padding:0 22px;margin-right:6px;border:0}
.pp .ribbon{height:98px;background:#fff;border-bottom:1px solid var(--ln);display:flex;padding:3px 4px;overflow-x:auto;white-space:nowrap;flex:none;-webkit-overflow-scrolling:touch}
.pp .soon{color:#888;padding:30px 16px;font-size:13px}
.pp .grp{display:flex;flex-direction:column;border-right:1px solid #e3e3e3;padding:0 6px}
.pp .gb{flex:1;display:flex;gap:4px;align-items:flex-start}
.pp .gn{text-align:center;font-size:11px;color:#666;height:16px}
.pp .rows{display:flex;flex-direction:column;gap:1px}
.pp .row{display:flex;gap:2px;align-items:center;height:24px}
.pp .b{min-width:22px;height:22px;border:1px solid transparent;background:none;display:inline-flex;align-items:center;justify-content:center;gap:5px;padding:0 4px;cursor:pointer}
.pp .b:hover{background:#fdeee8;border-color:#e8b5a3}
.pp .b.on{background:#f9d3c5;border-color:#e39b83}
.pp .b.rw{width:100%;justify-content:flex-start}
.pp .b.big{flex-direction:column;height:68px;min-width:48px;gap:3px;font-size:11px}
.pp .b.big em{font-size:24px;line-height:1.1}
.pp .sp1{display:flex;flex-direction:column;align-items:center}
.pp .sp1 .b.big{height:46px;min-width:52px;justify-content:center}
.pp .sp1 .b.sm{height:22px;font-size:11px;padding:0 4px}
.pp select{height:22px;border:1px solid #ababab;background:#fff}
.pp input[type=number]{height:22px;border:1px solid #ababab;width:58px}
.pp .clr{position:relative;width:26px;height:22px;display:grid;place-items:center;cursor:pointer}
.pp .clr:hover{background:#fdeee8}
.pp .clr span{padding:0 3px;line-height:1.1}
.pp .clr input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.pp .clr.lb{width:112px;display:flex;align-items:center;justify-content:flex-start;gap:5px;padding:0 4px}
.pp .shg{display:grid;grid-template-columns:repeat(6,24px);grid-auto-rows:22px;border:1px solid var(--ln);width:150px;height:70px;overflow:hidden;padding:1px}
.pp .shg .b{min-width:0;width:24px;height:22px;padding:0}
.pp .main{flex:1;display:flex;min-height:0}
.pp .thumbs{width:192px;flex:none;overflow:auto;background:linear-gradient(#fbfbfb,#efefef);border-right:1px solid #ddd;padding:10px 8px 10px 4px;outline:none}
.pp .th{display:flex;gap:4px;margin-bottom:10px;cursor:pointer}
.pp .th .n{width:20px;text-align:right;color:#777;font-size:12px;padding-top:2px;flex:none}
.pp .th.on .n{color:var(--acc)}
.pp .tv{position:relative;overflow:hidden;background:#fff;border:1px solid #bbb;flex:none}
.pp .th.on .tv,.pp .th2.on .tv{outline:2px solid var(--acc);outline-offset:-1px}
.pp .th.dr .tv{box-shadow:0 -3px 0 var(--acc)}
.pp .sec{display:flex;gap:4px;font-weight:600;margin:4px 0 8px 4px;cursor:pointer;color:#555}
.pp .ws{flex:1;display:flex;flex-direction:column;min-width:0}
.pp .desk{flex:1;overflow:auto;padding:24px;display:flex;background:linear-gradient(#fafafa,#efefef)}
.pp .holder{margin:auto;position:relative;flex:none}
.pp .slide{position:absolute;left:0;top:0;width:${W}px;height:${H}px;background:#fff;transform-origin:0 0;box-shadow:0 0 0 1px #cfcfcf,0 2px 10px rgba(0,0,0,.12);outline:none}
.pp .slide.draw{cursor:crosshair}.pp .slide.paint{cursor:copy}
.pp .notes{height:72px;flex:none;border-top:1px solid #ddd;background:#fff;padding:6px 12px;overflow:auto;outline:none;font-size:13px}
.pp .notes:empty::before{content:attr(data-ph);color:#999}
.pp .sorter{flex:1;overflow:auto;padding:16px;display:flex;flex-wrap:wrap;gap:18px;align-content:flex-start;background:#f2f2f2}
.pp .th2{display:flex;flex-direction:column;align-items:center;gap:4px;cursor:pointer}
.pp .sc{position:absolute;left:0;top:0;width:${W}px;height:${H}px;transform-origin:0 0}
.pp .ob>img.pic,.pp-print .ob>img.pic{position:absolute;inset:0;width:100%;height:100%;pointer-events:none;-webkit-user-drag:none}
.pp .ob{position:absolute;display:flex;flex-direction:column}
.pp .ob>svg.bg{position:absolute;inset:0;width:100%;height:100%;overflow:visible;pointer-events:none}
.pp .tx{position:relative;width:100%;outline:none;padding:4.8px 9.6px;word-wrap:break-word;overflow-wrap:anywhere;flex:none;max-height:none}
.pp .tx ul,.pp .tx ol{margin:0;padding-left:.9em}
.pp .tx ul ul,.pp .tx ol ol,.pp .tx ul ol,.pp .tx ol ul{padding-left:1.1em}
.pp .tx li{margin:.1em 0}
.pp .tx div,.pp .tx p{margin:0}
.pp .stg .tx:empty::before{content:attr(data-ph)}
.pp .stg .tx:focus:empty::before{content:""}
.pp .stg .ob.ph:not(.sel){outline:1px dotted #8b8b8b}
.pp .stg .ob.sh .tx{pointer-events:none}
.pp .stg .ob.sh.edt .tx{pointer-events:auto}
.pp .stg .ob.sel{outline:calc(1px*var(--inv)) solid #8a8a8a}
.pp .hd{position:absolute;width:calc(10px*var(--inv));height:calc(10px*var(--inv));background:#fff;border:calc(1px*var(--inv)) solid #666;border-radius:50%;transform:translate(-50%,-50%);z-index:3}
.pp .ed{position:absolute;cursor:move;z-index:2}
.pp .ed.t{left:0;right:0;top:calc(-6px*var(--inv));height:calc(12px*var(--inv))}
.pp .ed.b{left:0;right:0;bottom:calc(-6px*var(--inv));height:calc(12px*var(--inv))}
.pp .ed.l{top:0;bottom:0;left:calc(-6px*var(--inv));width:calc(12px*var(--inv))}
.pp .ed.r{top:0;bottom:0;right:calc(-6px*var(--inv));width:calc(12px*var(--inv))}
.pp .mq{position:absolute;border:calc(1px*var(--inv)) dashed #4a7ec8;background:rgba(74,126,200,.12);z-index:5;pointer-events:none}
.pp .find{position:absolute;right:24px;top:150px;z-index:5;background:#fff;border:1px solid #aaa;box-shadow:0 2px 8px rgba(0,0,0,.25);padding:10px;display:grid;grid-template-columns:auto 170px;gap:6px;align-items:center}
.pp .find input{height:24px;border:1px solid #aaa;padding:0 4px}
.pp .find .fb{grid-column:1/3;display:flex;gap:4px;justify-content:flex-end}
.pp .find button{height:24px;border:1px solid #aaa;background:#f3f3f3;cursor:pointer;padding:0 8px}
.pp .status{background:var(--acc);color:#fff;height:22px;display:flex;align-items:center;justify-content:space-between;padding:0 8px;font-size:11px;text-transform:uppercase;flex:none}
.pp .status .s{display:flex;gap:16px;align-items:center}
.pp .status button{background:none;border:0;color:#fff;cursor:pointer;min-width:20px;height:20px;font-size:12px;text-transform:uppercase;padding:0 4px}
.pp .status button.on{background:rgba(255,255,255,.25)}
.pp .status input{width:130px}
.pp .menu{position:absolute;z-index:10;top:52px;left:0;bottom:22px;width:260px;background:var(--acc);color:#fff;padding:8px 0}
.pp .menu button{display:block;width:100%;background:none;border:0;color:#fff;text-align:left;padding:12px 24px;cursor:pointer;font-size:14px}
.pp .menu button:hover{background:rgba(255,255,255,.18)}
.pp .menu hr{border:0;border-top:1px solid rgba(255,255,255,.3);margin:6px 0}
.pp .toast{position:absolute;right:16px;bottom:36px;background:#323130;color:#fff;padding:8px 14px;z-index:60;max-width:min(360px,calc(100% - 32px))}
.pp .pm{position:absolute;z-index:40;background:#fff;border:1px solid #aaa;box-shadow:0 2px 8px rgba(0,0,0,.25);min-width:170px;padding:2px 0;max-height:70vh;overflow:auto}
.pp .pm button{display:flex;align-items:center;gap:8px;width:100%;background:none;border:0;text-align:left;padding:5px 12px;cursor:pointer;white-space:nowrap}
.pp .pm button:hover{background:#fbe3da}
.pp .pm hr{border:0;border-top:1px solid #ddd;margin:3px 0}
.pp .pm.grid{display:grid;grid-template-columns:repeat(3,96px);gap:2px;padding:6px;min-width:0}
.pp .pm.grid button{flex-direction:column;padding:6px 2px;font-size:11px;text-align:center;gap:3px;line-height:1.1;white-space:normal}
.pp .pm.qs{grid-template-columns:repeat(4,56px)}
.pp .pm.qs button{padding:4px}
.pp .sw{display:block;width:44px;height:26px;border-radius:2px}
.pp .show{position:fixed;inset:0;z-index:100;background:#000;display:grid;place-items:center;outline:none}
.pp .sv{position:relative;overflow:hidden;background:#fff}
.pp.rc .ribbon{display:none}
@media (pointer:coarse){
 .pp .qat{height:36px}.pp .qat button{width:34px;height:32px;font-size:16px}
 .pp .tabs{height:36px}.pp .tab{height:36px;padding:0 16px;font-size:13px}.pp .tab.on{height:37px}
 .pp .ribbon{height:148px}.pp .row{height:36px}.pp .b{min-width:34px;height:34px}.pp .b.rw{height:34px}
 .pp select{height:34px;font-size:16px}.pp .clr{height:34px;width:36px}.pp .clr.lb{width:120px}
 .pp .shg{grid-template-columns:repeat(6,34px);grid-auto-rows:34px;width:210px;height:106px}.pp .shg .b{width:34px;height:34px}
 .pp .status{height:30px}.pp .menu button{padding:16px 24px}
}
@media (max-width:700px){
 .pp .main{flex-direction:column}
 .pp .thumbs{width:auto;height:86px;display:flex;gap:8px;overflow-x:auto;border-right:0;border-bottom:1px solid #ddd;padding:6px}
 .pp .th{margin:0;flex:none}.pp .th .n{display:none}.pp .sec{display:none}
 .pp .desk{padding:10px}.pp .status .s>span:nth-child(2),.pp .status .zr{display:none}
 .pp .find{left:8px;right:8px;top:auto;bottom:40px;grid-template-columns:auto 1fr}
}
@media print{
 body>*:not(.pp-print){display:none!important}
 .pp-print{display:block!important}
 .pp-print .pg{position:relative;width:${W}px;height:${H}px;overflow:hidden;page-break-after:always;break-after:page;-webkit-print-color-adjust:exact;print-color-adjust:exact}
 @page{size:${W}px ${H}px;margin:0}
}
.pp-print{display:none}
.pp-print .ob{position:absolute;display:flex;flex-direction:column}
.pp-print .ob>svg.bg{position:absolute;inset:0;width:100%;height:100%;overflow:visible}
.pp-print .tx{width:100%;padding:4.8px 9.6px;overflow-wrap:anywhere}
.pp-print .tx ul,.pp-print .tx ol{margin:0;padding-left:.9em}.pp-print .tx div,.pp-print .tx p{margin:0}
`;

const FONTS=['Calibri','Calibri Light','Arial','Times New Roman','Cambria','Georgia','Verdana','Courier New','Segoe UI','Tahoma','Trebuchet MS'];
const SIZES=[8,9,10,11,12,14,16,18,20,24,28,32,36,40,44,48,54,60,72,96];
const TABS=[['home','BERANDA'],['insert','SISIPKAN'],['design','DESAIN'],['trans','TRANSISI'],['anim','ANIMASI'],['show','SLIDE SHOW'],['review','TINJAU'],['view','TAMPILAN']];
const LAY=[['title','Slide Judul'],['content','Judul dan Konten'],['section','Header Bagian'],['two','Dua Konten'],['tonly','Hanya Judul'],['blank','Kosong']];
const LI={title:[[6,4,24,7],[10,13,16,4]],content:[[3,2,30,4],[3,8,30,11]],section:[[4,6,28,6],[4,14,28,3]],two:[[3,2,30,4],[3,8,14,11],[19,8,14,11]],tonly:[[3,2,30,5]],blank:[]};
const SHAPES=[['tb','Kotak Teks'],['line','Garis'],['arrow','Garis Panah'],['rect','Persegi Panjang'],['round','Persegi Panjang Sudut Bulat'],['oval','Elips'],
 ['tri','Segitiga'],['rtri','Segitiga Siku-siku'],['diamond','Belah Ketupat'],['pent','Segi Lima'],['hex','Segi Enam'],['star','Bintang'],
 ['rarrow','Panah Kanan'],['larrow','Panah Kiri'],['uarrow','Panah Atas'],['darrow','Panah Bawah'],['chevron','Chevron'],['plus','Tanda Tambah']];
const PTS={tri:[[.5,0],[1,1],[0,1]],rtri:[[0,0],[1,1],[0,1]],diamond:[[.5,0],[1,.5],[.5,1],[0,.5]],pent:[[.5,0],[1,.38],[.81,1],[.19,1],[0,.38]],
 hex:[[.25,0],[.75,0],[1,.5],[.75,1],[.25,1],[0,.5]],
 rarrow:[[0,.25],[.6,.25],[.6,0],[1,.5],[.6,1],[.6,.75],[0,.75]],larrow:[[1,.25],[.4,.25],[.4,0],[0,.5],[.4,1],[.4,.75],[1,.75]],
 uarrow:[[.25,1],[.25,.4],[0,.4],[.5,0],[1,.4],[.75,.4],[.75,1]],darrow:[[.25,0],[.75,0],[.75,.6],[1,.6],[.5,1],[0,.6],[.25,.6]],
 chevron:[[0,0],[.7,0],[1,.5],[.7,1],[0,1],[.3,.5]],plus:[[.35,0],[.65,0],[.65,.35],[1,.35],[1,.65],[.65,.65],[.65,1],[.35,1],[.35,.65],[0,.65],[0,.35],[.35,.35]]};
PTS.star=Array.from({length:10},(_,i)=>{const a=-Math.PI/2+i*Math.PI/5,r=i%2?.19:.5;return[.5+r*Math.cos(a),.5+r*Math.sin(a)];});
const QS=[['Isi Biru',{fill:'#5b9bd5',stroke:'#41719c',sw:1,color:'#fff'}],['Isi Oranye',{fill:'#ed7d31',stroke:'#ae5a21',sw:1,color:'#fff'}],['Isi Abu-abu',{fill:'#a5a5a5',stroke:'#787878',sw:1,color:'#fff'}],['Isi Emas',{fill:'#ffc000',stroke:'#bc8c00',sw:1,color:'#000'}],
 ['Isi Hijau',{fill:'#70ad47',stroke:'#507e32',sw:1,color:'#fff'}],['Garis Biru',{fill:'',stroke:'#5b9bd5',sw:2.25,color:'#41719c'}],['Garis Gelap',{fill:'',stroke:'#262626',sw:2.25,color:'#262626'}],['Isi Gelap',{fill:'#262626',stroke:'#000',sw:1,color:'#fff'}]];
const STATE=['bold','italic','underline','strikeThrough','insertUnorderedList','insertOrderedList','justifyLeft','justifyCenter','justifyRight','justifyFull'];
const CMDS=new Set(['bold','italic','underline','strikeThrough','insertUnorderedList','insertOrderedList','indent','outdent','justifyLeft','justifyCenter','justifyRight','justifyFull','removeFormat']);
const AN={t:'flex-start',m:'center',b:'flex-end'};
const VERT={v90:'writing-mode:vertical-rl;height:100%;width:auto;',v270:'writing-mode:vertical-rl;transform:rotate(180deg);height:100%;width:auto;',st:'writing-mode:vertical-rl;text-orientation:upright;height:100%;width:auto;'};

const esc=t=>String(t).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const B=(ic,tip,a,m)=>`<button class="b ${m||''}" title="${tip}" data-a="${a}"><em>${ic}</em>${m?`<span>${tip}</span>`:''}</button>`;
const BM=(ic,tip,m)=>`<button class="b rw" title="${tip}" data-m="${m}"><em>${ic}</em><span>${tip} ▾</span></button>`;
const S=(n,o,w,tip,r)=>`<select data-s="${n}" title="${tip||''}" style="width:${w}px"${r?' data-r':''}>${o.map(x=>Array.isArray(x)?`<option value="${x[0]}">${x[1]}</option>`:`<option>${x}</option>`).join('')}</select>`;
const C=(n,v,t,ic)=>`<label class="clr" title="${t}"><span style="border-bottom:3px solid ${v}">${ic}</span><input type="color" data-s="${n}" value="${v}"></label>`;
const CL=(n,v,t,ic)=>`<label class="clr lb" title="${t}"><span style="border-bottom:3px solid ${v}">${ic}</span><span>${t}</span><input type="color" data-s="${n}" value="${v}"></label>`;
const G=(n,h)=>`<div class="grp"><div class="gb">${h}</div><div class="gn">${n}</div></div>`;
const R=(...r)=>`<div class="rows">${r.map(x=>`<div class="row">${x}</div>`).join('')}</div>`;
const AL=t=>{const w=[14,9,14,9],x=i=>t==='L'||t==='J'?0:t==='C'?(14-w[i])/2:14-w[i];if(t==='J')w.fill(14);return `<svg width="16" height="14">${w.map((v,i)=>`<rect x="${1+x(i)}" y="${1+i*3.5}" width="${v}" height="1.6" fill="#444"/>`).join('')}</svg>`;};
const layIcon=id=>`<svg width="38" height="24" viewBox="0 0 36 22"><rect x=".5" y=".5" width="35" height="21" fill="#fff" stroke="#999"/>${LI[id].map(r=>`<rect x="${r[0]}" y="${r[1]}" width="${r[2]}" height="${r[3]}" fill="none" stroke="#777" stroke-dasharray="2 1"/>`).join('')}</svg>`;
const shapeIcon=id=>id==='tb'?`<svg width="16" height="16" viewBox="0 0 16 16"><rect x="1.5" y="2.5" width="13" height="11" fill="none" stroke="#4a4a4a"/><text x="8" y="11.5" font-size="9" text-anchor="middle" fill="#4a4a4a">A</text></svg>`
  :(id==='line'||id==='arrow')?`<svg width="16" height="16" viewBox="0 0 16 16"><path d="M2 14L${id==='arrow'?11:14} ${id==='arrow'?5:2}" stroke="#4a4a4a" fill="none"/>${id==='arrow'?'<path d="M14 2L8 4.5L11.5 8Z" fill="#4a4a4a"/>':''}</svg>`
  :`<svg width="16" height="16" viewBox="0 0 16 16">${shapeSVG({shape:id,w:16,h:16,fill:'none',stroke:'#4a4a4a',sw:1,dash:''})}</svg>`;

function shapeSVG(o){
  const sw=o.stroke?(+o.sw||1):0,i=sw/2,w=Math.max(1,o.w),h=Math.max(1,o.h),x0=i,y0=i,x1=w-i,y1=h-i,t=o.shape;
  const A=`fill="${o.fill||'none'}" stroke="${o.stroke||'none'}" stroke-width="${sw}" stroke-linejoin="round"${o.dash?` stroke-dasharray="${o.dash}"`:''}`;
  if(t==='line'||t==='arrow'){const y=h/2;
    return `<line x1="${x0}" y1="${y}" x2="${t==='arrow'?w-12:w-i}" y2="${y}" ${A} fill="none"/>`+(t==='arrow'?`<polygon points="${w},${y} ${w-14},${y-7} ${w-14},${y+7}" fill="${o.stroke||'#000'}" stroke="none"/>`:'');}
  if(t==='oval')return `<ellipse cx="${w/2}" cy="${h/2}" rx="${Math.max(0,(x1-x0)/2)}" ry="${Math.max(0,(y1-y0)/2)}" ${A}/>`;
  if(PTS[t])return `<polygon points="${PTS[t].map(([a,b])=>(x0+a*(x1-x0)).toFixed(1)+','+(y0+b*(y1-y0)).toFixed(1)).join(' ')}" ${A}/>`;
  return `<rect x="${x0}" y="${y0}" width="${Math.max(0,x1-x0)}" height="${Math.max(0,y1-y0)}"${t==='round'?` rx="${Math.min(w,h)/6}"`:''} ${A}/>`;
}
function fxCSS(o){
  const c=o.fill||o.stroke||'#5b9bd5';let f='';
  if(o.fx==='shadow')f='filter:drop-shadow(4px 4px 4px rgba(0,0,0,.45));';
  else if(o.fx==='glow')f=`filter:drop-shadow(0 0 10px ${c});`;
  else if(o.fx==='soft')f='filter:blur(3px);';
  else if(o.fx==='bevel')f='filter:drop-shadow(2px 2px 0 rgba(0,0,0,.35)) drop-shadow(-1px -1px 0 rgba(255,255,255,.7));';
  else if(o.fx==='refl')f='-webkit-box-reflect:below 2px linear-gradient(transparent 65%,rgba(255,255,255,.45));';
  if(o.fh||o.fv)f+=`transform:scale(${o.fh?-1:1},${o.fv?-1:1});`;
  return f;
}
function txCSS(o){
  const t=o.ts||{};
  return `font-family:'${t.font||'Calibri'}',Carlito,Calibri,sans-serif;font-size:${t.size||18}pt;color:${t.color||'#000'};text-align:${t.align||'left'};line-height:${t.lh||1.1};${t.bold?'font-weight:700;':''}${VERT[o.vert]||''}${o.cols>1?`column-count:${o.cols};column-gap:24px;`:''}`;
}
function obHTML(o,edit){
  const noTx=o.shape==='line'||o.shape==='arrow';
  return `<div class="ob ${o.k}" data-id="${o.id}" style="left:${o.x}px;top:${o.y}px;width:${o.w}px;height:${o.h}px;transform:rotate(${o.rot||0}deg);justify-content:${AN[o.anchor]||'flex-start'}">`+
    `<svg class="bg" viewBox="0 0 ${Math.max(1,o.w)} ${Math.max(1,o.h)}" style="${fxCSS(o)}">${shapeSVG(o)}</svg>`+
    (o.img?`<img class="pic" src="${o.img}" alt="" draggable="false" style="${fxCSS(o)}">`:'')+
    (noTx?'':`<div class="tx"${edit?' contenteditable="true" spellcheck="true"':''} data-ph="${esc(o.ph||'')}" style="${txCSS(o)}">${o.html}</div>`)+`</div>`;
}

let uid=1;
const mk=p=>Object.assign({id:'o'+(uid++),k:'tb',shape:'rect',fill:'',stroke:'',sw:1,dash:'',fx:'',rot:0,fh:0,fv:0,anchor:'t',vert:'h',cols:1,html:'',ph:'',x:100,y:100,w:300,h:80},p,
  {ts:Object.assign({font:'Calibri',size:18,color:'#000',align:'left'},p.ts||{})});
function layoutObjs(l){
  const ti=(p)=>mk(Object.assign({k:'ph',role:'title',ph:'Klik untuk menambahkan judul',x:88,y:29,w:1104,h:140,anchor:'m',ts:{font:'Calibri Light',size:44,align:'left'}},p));
  const bd=(x,w)=>mk({k:'ph',role:'body',bul:1,ph:'Klik untuk menambahkan teks',x,y:168,w:w,h:500,anchor:'t',ts:{size:28}});
  switch(l){
    case'title':return[ti({x:160,y:118,w:960,h:251,anchor:'b',ts:{size:60,align:'center'}}),mk({k:'ph',role:'sub',ph:'Klik untuk menambahkan subjudul',x:160,y:379,w:960,h:174,anchor:'t',ts:{size:24,align:'center'}})];
    case'content':return[ti(),bd(88,1104)];
    case'section':return[ti({x:101,y:180,w:1078,h:300,anchor:'b',ts:{size:60}}),mk({k:'ph',role:'sub',ph:'Klik untuk menambahkan teks',x:101,y:489,w:1078,h:160,anchor:'t',ts:{size:24,color:'#898989'}})];
    case'two':return[ti(),bd(88,540),bd(652,540)];
    case'tonly':return[ti()];
    default:return[];
  }
}
const mkSlide=l=>({layout:l,objs:layoutObjs(l),notes:''});
const textOf=h=>{const d=document.createElement('div');d.innerHTML=h;return d.textContent;};
const lines=h=>{const d=document.createElement('div');d.innerHTML=String(h).replace(/<\/(li|div|p|h\d)>|<br\s*\/?>/gi,'\n');return d.textContent.split('\n').filter((x,i,a)=>x.trim()||i<a.length-1&&0);};
const plainHTML=(o,h)=>{const L=lines(h).map(x=>x.trim()).filter(Boolean);if(!L.length)return'';
  return o.k==='ph'&&o.bul?'<ul>'+L.map(x=>'<li>'+esc(x)+'</li>').join('')+'</ul>':L.map(x=>'<div>'+esc(x)+'</div>').join('');};
const TXR=['body','sub','text'];

const RIB_HOME=
 G('Papan klip',`<div class="sp1"><button class="b big" data-a="paste" title="Tempel (Ctrl+V)"><em>📋</em></button><button class="b sm" data-m="paste">Tempel ▾</button></div>`+
   `<div class="rows">${B('✂','Potong','cut','rw')}${B('⎘','Salin','copy','rw')}${B('🖌','Kuas Format','painter','rw')}</div>`)+
 G('Slide',`<div class="sp1"><button class="b big" data-a="newslide" title="Slide Baru (Ctrl+M)"><em>🗎</em></button><button class="b sm" data-m="newslide">Slide Baru ▾</button></div>`+
   `<div class="rows">${BM('▦','Tata Letak','layout')}${B('↺','Atur Ulang','reset','rw')}${BM('☰','Bagian','section')}</div>`)+
 G('Font',R(
   S('font',FONTS,128,'Font')+S('size',SIZES,48,'Ukuran font')+B('A⁺','Perbesar font','grow')+B('A⁻','Perkecil font','shrink')+B('🧽','Hapus semua format','removeFormat'),
   B('<b>B</b>','Tebal (Ctrl+B)','bold')+B('<i>I</i>','Miring (Ctrl+I)','italic')+B('<u>U</u>','Garis bawah (Ctrl+U)','underline')+B('<span style="text-shadow:1px 1px 1px #888">S</span>','Bayangan teks','tshadow')+B('<s>abc</s>','Coret','strikeThrough')+
   S('spc',[['','AV'],['-1','Rapat'],['0','Normal'],['1','Longgar'],['3','Sangat longgar']],52,'Jarak karakter',1)+
   S('case',[['','Aa'],['upper','HURUF BESAR'],['lower','huruf kecil'],['title','Kapital Tiap Kata'],['sent','Kapital kalimat']],50,'Ubah huruf besar/kecil',1)+
   C('color','#ff0000','Warna font','<b>A</b>')))+
 G('Paragraf',R(
   B('•≡','Poin','insertUnorderedList')+B('1≡','Penomoran','insertOrderedList')+B('⇤','Kurangi indentasi','outdent')+B('⇥','Tambah indentasi','indent')+
   S('lh',[['','↕ Spasi'],['1','1,0'],['1.15','1,15'],['1.5','1,5'],['2','2,0'],['3','3,0']],74,'Spasi baris',1),
   S('dir',[['','Arah Teks'],['h','Horizontal'],['v90','Putar 90°'],['v270','Putar 270°'],['st','Bertumpuk']],92,'Arah teks',1)+
   B(AL('L'),'Rata kiri (Ctrl+L)','justifyLeft')+B(AL('C'),'Tengah (Ctrl+E)','justifyCenter')+B(AL('R'),'Rata kanan (Ctrl+R)','justifyRight')+B(AL('J'),'Rata kiri-kanan (Ctrl+J)','justifyFull')+
   S('cols',[['1','1 Kol'],['2','2 Kol'],['3','3 Kol']],58,'Kolom'),
   S('anchor',[['','Ratakan Teks'],['t','Atas'],['m','Tengah'],['b','Bawah']],112,'Ratakan teks')+
   `<button class="b" data-m="smart" title="Konversi ke SmartArt"><em>⬚</em><span>Konversi ke SmartArt ▾</span></button>`))+
 G('Gambar',`<div class="shg">${SHAPES.map(([id,n])=>`<button class="b" data-a="shape:${id}" title="${n}">${shapeIcon(id)}</button>`).join('')}</div>`+
   `<button class="b big" data-m="arrange" title="Susun"><em>⧉</em><span>Susun ▾</span></button><button class="b big" data-m="qstyle" title="Gaya Cepat"><em>◧</em><span>Gaya Cepat ▾</span></button>`+
   R(`${CL('sfill','#5b9bd5','Isian Bentuk','🪣')}<button class="b" data-a="nofill" title="Tanpa isian">∅</button>`,
     `${CL('sline','#41719c','Kerangka','▢')}${S('sw',[['0.5','½'],['1','1'],['2.25','2¼'],['3','3'],['4.5','4½'],['6','6']],46,'Tebal garis')}${S('dash',[['','—'],['8 5','- -'],['2 4','···']],42,'Jenis garis')}`,
     S('sfx',[['','Efek Bentuk ▾'],['shadow','Bayangan'],['refl','Pantulan'],['glow','Cahaya'],['soft','Tepi Halus'],['bevel','Timbul'],['none','Hapus efek']],112,'Efek bentuk',1)))+
 G('Penyuntingan',`<div class="rows">${B('🔍','Cari','find','rw')}${B('ab','Ganti','replace','rw')}${BM('⬚','Pilih','pilih')}</div>`);

function mount(root,opts={}){
  if(!document.getElementById('pp-css')){const s=document.createElement('style');s.id='pp-css';s.textContent=CSS;document.head.appendChild(s);}
  let name='Presentasi1.pptx',scale=.8,autoFit=true,slides=[mkSlide('title')],cur=0,sel=[],hist=[],hi=-1,clip=null,painting=null,lastR=null,lastTx=null,drawing=null,view='normal',fk=-1,timer=0,sortFrom=null;
  const panes=TABS.map((t,i)=>`<div class="ribbon" data-pane="${t[0]}"${i?' hidden':''}>${t[0]==='home'?RIB_HOME:`<div class="soon">Menu ${t[1]} akan ditambahkan pada tahap berikutnya.</div>`}</div>`).join('');
  root.innerHTML=`<div class="pp">
    <div class="qat"><div class="l"><span class="w">P</span><button data-a="save" title="Simpan (Ctrl+S)">💾</button><button data-a="undo" title="Urungkan (Ctrl+Z)">↶</button><button data-a="redo" title="Ulangi (Ctrl+Y)">↷</button><button data-a="showfirst" title="Mulai dari awal">▶</button></div><span class="fn"></span></div>
    <div class="tabs"><button class="tab file" data-a="menu">BERKAS</button>${TABS.map((t,i)=>`<button class="tab${i?'':' on'}" data-tab="${t[0]}">${t[1]}</button>`).join('')}</div>
    ${panes}
    <div class="main"><div class="thumbs" tabindex="0"></div>
      <div class="ws"><div class="desk"><div class="holder"><div class="slide stg" tabindex="-1"></div></div></div><div class="notes" contenteditable="true" data-ph="Klik untuk menambahkan catatan"></div></div>
      <div class="sorter" hidden></div></div>
    <div class="find" hidden><span>Cari:</span><input type="text" class="f1"><span>Ganti:</span><input type="text" class="f2">
      <div class="fb"><button data-a="fnext">Berikutnya</button><button data-a="frep">Ganti</button><button data-a="frall">Ganti semua</button><button data-a="fclose">Tutup</button></div></div>
    <div class="status"><div class="s"><span class="sn">SLIDE 1 DARI 1</span><span>BAHASA INDONESIA</span></div>
      <div class="s"><button data-a="notes" class="on">☰ CATATAN</button><button data-a="comments">💬 KOMENTAR</button>
      <button data-a="vnormal" class="on vb" title="Normal">▭</button><button data-a="vsorter" class="vb" title="Penyortir Slide">▦</button><button data-a="show" title="Slide Show">▶</button>
      <span><button data-a="zoomd:-10">−</button><input type="range" min="10" max="200" value="80" class="zr"><button data-a="zoomd:10">+</button> <span class="zv">80%</span> <button data-a="fit" title="Sesuaikan ke jendela">⛶</button></span></div></div>
    <input type="file" class="fin" accept=".pptx" hidden>
    <div class="menu" hidden><button data-a="new">📄  Presentasi kosong baru</button><button data-a="open">📂  Buka…</button><button data-a="save">💾  Simpan sebagai .pptx</button><button data-a="print">🖨  Cetak / simpan PDF</button><hr><button data-a="home">⬅  Kembali ke menu utama</button></div>
  </div>`;

  const $=s=>root.querySelector(s),$$=s=>[...root.querySelectorAll(s)];
  const pp=$('.pp'),stg=$('.stg'),hold=$('.holder'),desk=$('.desk'),notes=$('.notes'),thumbsEl=$('.thumbs'),sorter=$('.sorter'),fm=$('.menu'),fd=$('.find'),f1=$('.f1'),f2=$('.f2'),zr=$('.zr'),zv=$('.zv'),fin=$('.fin');
  const toast=(m,ms)=>{$$('.toast').forEach(x=>x.remove());const t=document.createElement('div');t.className='toast';t.textContent=m;pp.appendChild(t);setTimeout(()=>t.remove(),ms||3000);};
  const setName=()=>{$('.fn').textContent=name.replace(/\.pptx$/i,'')+' - PowerPoint';document.title=$('.fn').textContent;};
  const S0=()=>slides[cur],objs=()=>S0().objs,byId=id=>objs().find(o=>o.id===id),selObjs=()=>sel.map(byId).filter(Boolean);
  const elOf=id=>stg.querySelector(`.ob[data-id="${id}"]`);
  const txOf=n=>{const e=n&&(n.nodeType===3?n.parentElement:n);const t=e&&e.closest&&e.closest('.tx');return t&&stg.contains(t)?t:null;};
  const status=()=>{$('.sn').textContent=`SLIDE ${cur+1} DARI ${slides.length}`;};

  /* ===== model ⇄ DOM, riwayat ===== */
  function sync(){
    $$('.stg .ob').forEach(el=>{const o=byId(el.dataset.id),t=el.querySelector('.tx');if(o&&t)o.html=(t.textContent.trim()||t.querySelector('img,li'))?(t.textContent.trim()?t.innerHTML:''):'';});
    S0().notes=notes.textContent.trim()?notes.innerHTML:'';
  }
  const imgReg=new Map(),imgList=[];     // data gambar disimpan sekali; snapshot hanya memuat rujukan
  const snap=()=>JSON.stringify({slides,cur},(k,v)=>{if(k==='img'&&typeof v==='string'&&v[0]!=='@'){let i=imgReg.get(v);if(i==null){i=imgList.push(v)-1;imgReg.set(v,i);}return '@img'+i;}return v;});
  const unsnap=t=>JSON.parse(t,(k,v)=>k==='img'&&typeof v==='string'&&v.startsWith('@img')?imgList[+v.slice(4)]:v);
  function commit(){
    clearTimeout(timer);sync();
    const s=snap();if(hist[hi]===s)return;
    hist=hist.slice(0,hi+1);hist.push(s);if(hist.length>100)hist.shift();hi=hist.length-1;thumbs();
  }
  const commitSoon=()=>{clearTimeout(timer);timer=setTimeout(commit,500);};
  function restore(i){const d=unsnap(hist[i]);slides=d.slides;cur=Math.min(d.cur,slides.length-1);sel=[];draw();thumbs();notes.innerHTML=S0().notes||'';}
  const undo=()=>{commit();if(hi>0){hi--;restore(hi);}};
  const redo=()=>{if(hi<hist.length-1){hi++;restore(hi);}};

  /* ===== render ===== */
  function draw(){stg.innerHTML=objs().map(o=>obHTML(o,true)).join('');stg.style.background=S0().bg||'';decorate();status();}
  function decorate(){
    $$('.stg .ob').forEach(el=>{const on=sel.includes(el.dataset.id);el.classList.toggle('sel',on);if(!on)el.classList.remove('edt');el.querySelectorAll('.hd,.ed').forEach(x=>x.remove());});
    if(sel.length===1){const o=byId(sel[0]),el=o&&elOf(o.id);if(el){
      el.insertAdjacentHTML('beforeend',['t','b','l','r'].map(s=>`<div class="ed ${s}" data-ed="1"></div>`).join(''));
      const line=o.shape==='line'||o.shape==='arrow',hs=line?['w','e']:['nw','n','ne','e','se','s','sw','w'];
      el.insertAdjacentHTML('beforeend',hs.map(h=>{const x=h.includes('w')?0:h.includes('e')?100:50,y=h.includes('n')?0:h.includes('s')?100:50;
        return `<div class="hd" data-h="${h}" style="left:${x}%;top:${y}%;cursor:${h}-resize"></div>`;}).join(''));}}
    syncUI();
  }
  function redrawOb(o){sync();const el=elOf(o.id);if(el){el.outerHTML=obHTML(o,true);}decorate();}
  function thumbHTML(s,i,w,c){const k=w/W;return `<div class="${c}${i===cur?' on':''}" data-i="${i}" draggable="true"><span class="n">${i+1}</span><div class="tv" style="width:${w}px;height:${Math.round(H*k)}px${s.bg?';background:'+s.bg:''}"><div class="sc" style="transform:scale(${k})">${s.objs.map(o=>obHTML(o,false)).join('')}</div></div></div>`;}
  function thumbs(){
    const tp=thumbsEl.scrollTop,secs=slides.map((s,i)=>s.sec!=null?i:-1).filter(i=>i>=0);let h='';
    slides.forEach((s,i)=>{if(s.sec!=null){const nx=secs[secs.indexOf(i)+1];h+=`<div class="sec" data-sec="${i}"><i>▾</i><b>${esc(s.sec)}</b> (${(nx==null?slides.length:nx)-i})</div>`;}h+=thumbHTML(s,i,160,'th');});
    thumbsEl.innerHTML=h;thumbsEl.scrollTop=tp;
    if(view==='sorter')sorter.innerHTML=slides.map((s,i)=>thumbHTML(s,i,220,'th2')).join('');
    status();
  }
  function go(i){
    if(i<0||i>=slides.length)return;sync();cur=i;sel=[];lastR=null;draw();notes.innerHTML=S0().notes||'';
    $$('.th,.th2').forEach(t=>t.classList.toggle('on',+t.dataset.i===cur));
    const t=thumbsEl.querySelector('.th.on');if(t)t.scrollIntoView({block:'nearest'});
  }
  function setScale(s){scale=Math.max(.1,Math.min(2,s));hold.style.width=W*scale+'px';hold.style.height=H*scale+'px';stg.style.transform=`scale(${scale})`;stg.style.setProperty('--inv',1/scale);zr.value=Math.round(scale*100);zv.textContent=Math.round(scale*100)+'%';}
  function fit(){const dw=desk.clientWidth-48,dh=desk.clientHeight-48;if(dw>0&&dh>0)setScale(Math.min(dw/W,dh/H));}

  /* ===== seleksi teks & pemformatan ===== */
  document.addEventListener('selectionchange',function f(){
    if(!root.isConnected){document.removeEventListener('selectionchange',f);return;}
    const s=getSelection();if(!s.rangeCount)return;
    const tx=txOf(s.anchorNode);
    if(tx){lastR=s.getRangeAt(0).cloneRange();lastTx=tx;syncUI();}
    else if(stg.contains(s.anchorNode)||s.anchorNode===stg){lastR=null;}
  });
  const hasRange=()=>lastR&&lastTx&&lastTx.isConnected&&lastTx.contains(lastR.commonAncestorContainer);
  function target(fn,whole){
    const list=[];let range=null;
    if(hasRange()&&!(whole&&lastR.collapsed)){range=lastR;list.push(lastTx);}
    else if(hasRange()&&whole){list.push(lastTx);}
    else selObjs().forEach(o=>{const t=elOf(o.id)&&elOf(o.id).querySelector('.tx');if(t)list.push(t);});
    if(!list.length)return toast('Pilih objek atau teks terlebih dahulu');
    const s=getSelection();
    list.forEach(t=>{t.focus({preventScroll:true});s.removeAllRanges();
      if(range)s.addRange(range);else{const r=document.createRange();r.selectNodeContents(t);s.addRange(r);}
      fn(t);});
    if(!range){s.removeAllRanges();if(document.activeElement&&document.activeElement.blur)document.activeElement.blur();stg.focus({preventScroll:true});lastR=null;}
    commitSoon();syncUI();
  }
  const cmd=(c,v)=>target(()=>document.execCommand(c,false,v===undefined?null:v));
  function spanStyle(prop,val){
    document.execCommand('styleWithCSS',false,false);document.execCommand('fontSize',false,7);document.execCommand('styleWithCSS',false,true);
    let f;while((f=stg.querySelector('font[size="7"]'))){const s=document.createElement('span');s.style[prop]=val;s.innerHTML=f.innerHTML;s.querySelectorAll('[style]').forEach(e=>e.style[prop]='');f.replaceWith(s);}
  }
  const curSize=()=>{const t=hasRange()?lastTx:(selObjs()[0]&&elOf(sel[0])&&elOf(sel[0]).querySelector('.tx'));if(!t)return 18;const s=getSelection();let n=hasRange()&&s.rangeCount&&txOf(s.anchorNode)?(s.anchorNode.nodeType===3?s.anchorNode.parentElement:s.anchorNode):t;return Math.round(parseFloat(getComputedStyle(n).fontSize)*.75);};
  function syncUI(){
    const t=hasRange()?lastTx:(sel.length&&elOf(sel[0])?elOf(sel[0]).querySelector('.tx'):null);
    if(!t)return;
    const s=getSelection(),inTx=s.rangeCount&&txOf(s.anchorNode),n=inTx?(s.anchorNode.nodeType===3?s.anchorNode.parentElement:s.anchorNode):t,cs=getComputedStyle(n);
    const fam=cs.fontFamily.split(',')[0].replace(/["']/g,'').trim(),sz=Math.round(parseFloat(cs.fontSize)*.75);
    const fs=$('[data-s=font]'),ss=$('[data-s=size]');if(FONTS.includes(fam))fs.value=fam;if(SIZES.includes(sz))ss.value=sz;
    STATE.forEach(c=>{const b=root.querySelector(`[data-a="${c}"]`);if(!b)return;let on=false;
      if(inTx){try{on=document.queryCommandState(c);}catch{}}
      else if(c==='bold')on=+cs.fontWeight>=600;else if(c==='italic')on=cs.fontStyle==='italic';
      else if(c==='justifyLeft'||c==='justifyCenter'||c==='justifyRight'||c==='justifyFull'){const a=cs.textAlign;on=c==={left:'justifyLeft',start:'justifyLeft',center:'justifyCenter',right:'justifyRight',end:'justifyRight',justify:'justifyFull'}[a]?true:false;on=on||(c==='justifyLeft'&&(a==='left'||a==='start'));}
      b.classList.toggle('on',!!on);});
    const ts=$('[data-a=tshadow]');if(ts)ts.classList.toggle('on',cs.textShadow&&cs.textShadow!=='none');
  }
  const fmtNow=()=>{const s=getSelection(),inTx=hasRange()&&s.rangeCount&&txOf(s.anchorNode);const t=hasRange()?lastTx:(sel.length&&elOf(sel[0])&&elOf(sel[0]).querySelector('.tx'));if(!t)return null;
    const n=inTx?(s.anchorNode.nodeType===3?s.anchorNode.parentElement:s.anchorNode):t,cs=getComputedStyle(n);const q=c=>{try{return inTx&&document.queryCommandState(c);}catch{return false;}};
    return{bold:inTx?q('bold'):+cs.fontWeight>=600,italic:inTx?q('italic'):cs.fontStyle==='italic',underline:q('underline'),font:cs.fontFamily.split(',')[0].replace(/["']/g,'').trim(),size:Math.round(parseFloat(cs.fontSize)*.75),color:cs.color,align:cs.textAlign};};
  function applyPaint(){
    const p=painting;painting=null;stg.classList.remove('paint');if(!p)return;
    target(()=>{['bold','italic','underline'].forEach(c=>{let cur;try{cur=document.queryCommandState(c);}catch{}if(!!cur!==p[c])document.execCommand(c,false,null);});
      document.execCommand('fontName',false,p.font);spanStyle('fontSize',p.size+'pt');document.execCommand('foreColor',false,p.color);
      document.execCommand({center:'justifyCenter',right:'justifyRight',justify:'justifyFull'}[p.align]||'justifyLeft',false,null);},false);
  }
  const caseOf=(t,m)=>m==='upper'?t.toUpperCase():m==='lower'?t.toLowerCase():m==='sent'?t.toLowerCase().replace(/(^\s*|[.!?]\s+)(\S)/g,(a,b,c)=>b+c.toUpperCase()):t.toLowerCase().replace(/(^|\s)\S/g,c=>c.toUpperCase());
  function changeCase(m){
    if(!m)return;
    if(hasRange()&&!lastR.collapsed)return target(()=>{const t=getSelection().toString();if(t)document.execCommand('insertText',false,caseOf(t,m));});
    const list=hasRange()?[lastTx]:selObjs().map(o=>elOf(o.id)&&elOf(o.id).querySelector('.tx')).filter(Boolean);
    list.forEach(t=>{const w=document.createTreeWalker(t,NodeFilter.SHOW_TEXT);let n;while(n=w.nextNode())n.data=caseOf(n.data,m);});commitSoon();
  }
  const forObjs=fn=>{if(!sel.length)return toast('Pilih objek terlebih dahulu');sync();selObjs().forEach(fn);draw();commit();};

  /* ===== menu popup ===== */
  const closePop=()=>$$('.pm').forEach(m=>m.remove());
  function pop(items,x,y,cls){
    closePop();const m=document.createElement('div');m.className='pm '+(cls||'');
    m.innerHTML=items.map(it=>it==='-'?'<hr>':`<button data-a="${it[1]}" title="${esc(it[0])}">${it[2]||''}<span>${it[0]}</span></button>`).join('');
    pp.appendChild(m);const rr=pp.getBoundingClientRect();m.style.left=Math.max(0,Math.min(x-rr.left,rr.width-m.offsetWidth-4))+'px';m.style.top=Math.max(0,Math.min(y-rr.top,rr.height-m.offsetHeight-4))+'px';
  }
  const MENUS={
    paste:()=>[['Tempel','paste','📋'],['Tempel Teks Saja','pastetext','🗒']],
    newslide:()=>[LAY.map(([id,n])=>[n,'newslide:'+id,layIcon(id)]),'grid'],
    layout:()=>[LAY.map(([id,n])=>[n,'layout:'+id,layIcon(id)]),'grid'],
    section:()=>[['Tambah Bagian','section:add'],['Ganti Nama Bagian','section:ren'],['Hapus Bagian','section:rm'],['Hapus Semua Bagian','section:rmall']],
    arrange:()=>[['Bawa ke Depan','arrange:front'],['Kirim ke Belakang','arrange:back'],['Majukan','arrange:fwd'],['Mundurkan','arrange:bwd'],'-',
      ['Rata Kiri','arrange:al'],['Rata Tengah','arrange:ac'],['Rata Kanan','arrange:ar'],['Rata Atas','arrange:at'],['Rata Tengah Vertikal','arrange:am'],['Rata Bawah','arrange:ab'],'-',
      ['Putar Kanan 90°','arrange:rcw'],['Putar Kiri 90°','arrange:rccw'],['Balik Horizontal','arrange:fh'],['Balik Vertikal','arrange:fv']],
    qstyle:()=>[QS.map(([n,s],i)=>[n,'qs:'+i,`<i class="sw" style="background:${s.fill||'#fff'};border:${s.sw}px solid ${s.stroke}"></i>`]),'grid qs'],
    smart:()=>[['Daftar Blok Dasar','smart:blocks'],['Proses Dasar','smart:process'],['Daftar Vertikal','smart:list']],
    pilih:()=>[['Pilih Semua','selall'],['Batalkan Pilihan','selnone']]
  };
  function openMenu(btn,k){let m=MENUS[k]();let cls='';if(m.length===2&&typeof m[1]==='string'&&Array.isArray(m[0])&&Array.isArray(m[0][0])){cls=m[1];m=m[0];}
    const r=btn.getBoundingClientRect();pop(m,r.left,r.bottom,cls);}

  /* ===== objek: tambah, hapus, salin ===== */
  const cloneO=o=>Object.assign(JSON.parse(JSON.stringify(o)),{id:'o'+(uid++)});
  function addObj(o,edit){sync();objs().push(o);sel=[o.id];draw();commit();if(edit){const t=elOf(o.id)&&elOf(o.id).querySelector('.tx');if(t)t.focus();}}
  function delSel(){if(!sel.length)return;sync();S0().objs=objs().filter(o=>!sel.includes(o.id));sel=[];draw();commit();}
  function copyObjs(){const l=selObjs();if(!l.length)return false;sync();clip=JSON.parse(JSON.stringify(l));return true;}
  function pasteObjs(){if(!clip)return false;sync();const n=clip.map(o=>Object.assign(cloneO(o),{x:o.x+24,y:o.y+24}));objs().push(...n);clip=JSON.parse(JSON.stringify(n));sel=n.map(o=>o.id);draw();commit();return true;}
  const selectedText=()=>hasRange()&&!lastR.collapsed;
  function newSlide(l){
    sync();const c=S0().layout;l=l||(c==='title'||c==='section'?'content':c);
    slides.splice(cur+1,0,mkSlide(l));cur++;sel=[];draw();notes.innerHTML='';commit();
  }
  function applyLayout(s,l,reset){
    const old=s.objs.filter(o=>o.k==='ph'),nw=layoutObjs(l),used=new Set();
    const take=r=>{const o=old.find(x=>!used.has(x.id)&&x.role===r)||old.find(x=>!used.has(x.id)&&TXR.includes(x.role)&&TXR.includes(r));if(o)used.add(o.id);return o;};
    nw.forEach(n=>{const o=take(n.role);if(o)n.html=reset?plainHTML(n,o.html):o.html;});
    const orphans=old.filter(o=>!used.has(o.id)&&textOf(o.html).trim());
    s.objs=[...nw,...s.objs.filter(o=>o.k!=='ph'),...orphans];s.layout=l;
  }
  function smart(kind){
    const o=selObjs()[0];if(!o||o.shape==='line'||o.shape==='arrow')return toast('Pilih kotak teks berisi daftar poin');
    const L=lines(o.html).map(x=>x.trim()).filter(Boolean);if(!L.length)return toast('Objek tidak berisi teks');
    const n=L.length,pad=16,out=[];
    const box=(x,y,w,h,t)=>mk({k:'sh',shape:'round',x,y,w,h,fill:'#5b9bd5',stroke:'#41719c',sw:1,anchor:'m',html:'<div>'+esc(t)+'</div>',
      ts:{size:Math.max(10,Math.min(28,Math.floor(Math.min(h*.35,(w-20)/Math.max(6,t.length)*1.7)*.75))),color:'#fff',align:'center'}});
    if(kind==='blocks'){const c=Math.ceil(Math.sqrt(n)),r=Math.ceil(n/c),w=(o.w-pad*(c-1))/c,h=(o.h-pad*(r-1))/r;L.forEach((t,i)=>out.push(box(o.x+(i%c)*(w+pad),o.y+Math.floor(i/c)*(h+pad),w,h,t)));}
    else if(kind==='process'){const g=44,w=(o.w-g*(n-1))/n,h=Math.min(o.h,180),y=o.y+(o.h-h)/2;
      L.forEach((t,i)=>{out.push(box(o.x+i*(w+g),y,w,h,t));if(i<n-1)out.push(mk({k:'sh',shape:'rarrow',x:o.x+i*(w+g)+w+8,y:y+h/2-14,w:g-16,h:28,fill:'#a5a5a5',stroke:'#787878'}));});}
    else{const h=(o.h-pad*(n-1))/n;L.forEach((t,i)=>out.push(box(o.x,o.y+i*(h+pad),o.w,h,t)));}
    sync();const i=objs().indexOf(o);objs().splice(i,1,...out);sel=out.map(x=>x.id);draw();commit();
  }
  function arrange(v){
    if(!sel.length)return toast('Pilih objek terlebih dahulu');sync();const a=objs(),l=selObjs();
    if(v==='front'){S0().objs=[...a.filter(o=>!sel.includes(o.id)),...l];}
    else if(v==='back'){S0().objs=[...l,...a.filter(o=>!sel.includes(o.id))];}
    else if(v==='fwd'||v==='bwd'){const d=v==='fwd'?1:-1;(d>0?[...l].reverse():l).forEach(o=>{const i=a.indexOf(o),j=i+d;if(j>=0&&j<a.length&&!sel.includes(a[j].id)){a.splice(i,1);a.splice(j,0,o);}});}
    else if('al ac ar at am ab'.split(' ').includes(v)){
      const bx=l.length>1?{x:Math.min(...l.map(o=>o.x)),y:Math.min(...l.map(o=>o.y)),r:Math.max(...l.map(o=>o.x+o.w)),b:Math.max(...l.map(o=>o.y+o.h))}:{x:0,y:0,r:W,b:H};
      l.forEach(o=>{if(v==='al')o.x=bx.x;else if(v==='ac')o.x=(bx.x+bx.r-o.w)/2;else if(v==='ar')o.x=bx.r-o.w;else if(v==='at')o.y=bx.y;else if(v==='am')o.y=(bx.y+bx.b-o.h)/2;else o.y=bx.b-o.h;
        ['x','y'].forEach(k=>o[k]=Math.round(o[k]));});}
    else l.forEach(o=>{if(v==='rcw')o.rot=((o.rot||0)+90)%360;else if(v==='rccw')o.rot=((o.rot||0)-90+360)%360;else if(v==='fh')o.fh=o.fh?0:1;else if(v==='fv')o.fv=o.fv?0:1;});
    draw();commit();
  }
  function secOp(v){
    const s=S0(),has=slides.some(x=>x.sec!=null);
    if(v==='add'){const n=prompt('Nama bagian','Bagian Tanpa Judul');if(n==null)return;if(cur>0&&!has&&slides[0].sec==null)slides[0].sec='Bagian Default';s.sec=n||'Bagian Tanpa Judul';}
    else if(v==='ren'){let i=cur;while(i>=0&&slides[i].sec==null)i--;if(i<0)return toast('Belum ada bagian');const n=prompt('Ganti nama bagian',slides[i].sec);if(n)slides[i].sec=n;}
    else if(v==='rm'){let i=cur;while(i>=0&&slides[i].sec==null)i--;if(i<0)return toast('Belum ada bagian');delete slides[i].sec;}
    else slides.forEach(x=>delete x.sec);
    commit();
  }

  /* ===== cari & ganti ===== */
  function matches(q){const r=[],lq=q.toLowerCase();slides.forEach((s,si)=>s.objs.forEach(o=>{const t=textOf(o.html).toLowerCase();let i=t.indexOf(lq);while(i>=0){r.push({si,id:o.id,i});i=t.indexOf(lq,i+Math.max(1,lq.length));}}));return r;}
  function selText(el,start,len){
    const w=document.createTreeWalker(el,NodeFilter.SHOW_TEXT),r=document.createRange();let n,pos=0,a=false;
    while(n=w.nextNode()){const L=n.data.length;if(!a&&start<pos+L){r.setStart(n,start-pos);a=true;}if(a&&start+len<=pos+L){r.setEnd(n,start+len-pos);break;}pos+=L;}
    el.focus({preventScroll:true});const s=getSelection();s.removeAllRanges();s.addRange(r);
  }
  function findNext(){
    sync();const q=f1.value;if(!q)return;const m=matches(q);if(!m.length)return toast('Tidak ditemukan');
    fk=(fk+1)%m.length;const x=m[fk];if(x.si!==cur)go(x.si);
    sel=[x.id];decorate();const t=elOf(x.id).querySelector('.tx');selText(t,x.i,q.length);
  }
  function replaceAll(){
    sync();const q=f1.value;if(!q)return;const re=new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'gi');let n=0;
    slides.forEach(s=>s.objs.forEach(o=>{const d=document.createElement('div');d.innerHTML=o.html;const w=document.createTreeWalker(d,NodeFilter.SHOW_TEXT);let t;
      while(t=w.nextNode()){t.data=t.data.replace(re,()=>{n++;return f2.value;});}o.html=d.innerHTML;}));
    draw();commit();toast(n+' penggantian');
  }
  const showFind=()=>{fd.hidden=false;f1.focus();f1.select();};

  /* ===== slide show & cetak ===== */
  function show(from){
    sync();let i=from;const ov=document.createElement('div');ov.className='show';ov.tabIndex=0;pp.appendChild(ov);
    const paint=()=>{const k=Math.min(ov.clientWidth/W,ov.clientHeight/H);ov.innerHTML=`<div class="sv" style="width:${W*k}px;height:${H*k}px${slides[i].bg?';background:'+slides[i].bg:''}"><div class="sc" style="transform:scale(${k})">${slides[i].objs.map(o=>obHTML(o,false)).join('')}</div></div>`;};
    const close=()=>{document.removeEventListener('fullscreenchange',fc);ov.remove();if(document.fullscreenElement)document.exitFullscreen().catch(()=>{});go(i);};
    const fc=()=>{if(!document.fullscreenElement)close();};
    const next=()=>{if(i<slides.length-1){i++;paint();}else close();},prev=()=>{if(i>0){i--;paint();}};
    ov.onclick=next;
    ov.onkeydown=e=>{const k=e.key;if(k==='Escape')close();else if(['ArrowRight','ArrowDown','PageDown',' ','Enter'].includes(k)){e.preventDefault();next();}else if(['ArrowLeft','ArrowUp','PageUp','Backspace'].includes(k)){e.preventDefault();prev();}};
    paint();ov.focus();addEventListener('resize',paint);
    if(ov.requestFullscreen){ov.requestFullscreen().then(()=>document.addEventListener('fullscreenchange',fc)).catch(()=>{});}
  }
  function printAll(){
    sync();const d=document.createElement('div');d.className='pp-print';
    d.innerHTML=slides.map(s=>`<div class="pg"${s.bg?` style="background:${s.bg}"`:''}>${s.objs.map(o=>obHTML(o,false)).join('')}</div>`).join('');document.body.appendChild(d);
    const done=()=>{d.remove();removeEventListener('afterprint',done);};addEventListener('afterprint',done);setTimeout(()=>{print();},50);
  }

  /* ===== berkas .pptx ===== */
  const PH={title:'Klik untuk menambahkan judul',sub:'Klik untuk menambahkan subjudul',body:'Klik untuk menambahkan teks'};
  async function openFile(file){
    if(!/\.pptx$/i.test(file.name))return toast('Hanya berkas .pptx yang bisa dibuka');
    try{
      toast('Membuka '+file.name+'…',60000);
      await lib('jszip',()=>window.JSZip);await lib('pptxio',()=>window.PptxIO);
      const r=await PptxIO.read(await file.arrayBuffer(),window.JSZip);
      slides=r.slides.map(s=>{const n={layout:s.layout,notes:s.notes||'',objs:s.objs.map(o=>mk(o.k==='ph'?Object.assign({},o,{ph:PH[o.role]||PH.body}):o))};
        if(s.bg)n.bg=s.bg;if(s.sec!=null)n.sec=s.sec;return n;});
      cur=0;sel=[];lastR=null;name=file.name;setName();hist=[];hi=-1;
      draw();thumbs();notes.innerHTML=S0().notes||'';commit();
      const w=r.warn,skip=[w.table&&w.table+' tabel',w.chart&&w.chart+' bagan',w.media&&w.media+' gambar (format tidak didukung)',w.other&&w.other+' objek lain'].filter(Boolean);
      if(skip.length)toast('Dibuka: '+file.name+'. Belum didukung dan dilewati: '+skip.join(', '),7000);else toast('Dibuka: '+file.name);
    }catch(e){toast('Tidak bisa membuka berkas: '+e.message,6000);}
  }
  async function savePptx(){
    try{
      sync();toast('Menyiapkan .pptx…',60000);
      await lib('jszip',()=>window.JSZip);await lib('pptx',()=>window.PptxGenJS);await lib('pptxio',()=>window.PptxIO);
      const blob=await PptxIO.write({name,slides},window.PptxGenJS,window.JSZip);
      const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=/\.pptx$/i.test(name)?name:name+'.pptx';
      document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),1000);toast('Tersimpan: '+a.download);
    }catch(e){toast('Gagal menyimpan: '+e.message,6000);}
  }
  fin.addEventListener('change',()=>{const f=fin.files[0];fin.value='';if(f)openFile(f);});

  /* ===== aksi ===== */
  const acts={
    menu(){fm.hidden=!fm.hidden;},
    new(){slides=[mkSlide('title')];cur=0;sel=[];name='Presentasi1.pptx';setName();hist=[];hi=-1;draw();thumbs();notes.innerHTML='';commit();},
    open(){fin.click();},
    save:savePptx,
    print:printAll,home(){opts.onHome&&opts.onHome();},undo,redo,
    showfirst(){show(0);},show(){show(cur);},
    cut(){if(selectedText()){target(()=>document.execCommand('cut'));}else if(copyObjs()){delSel();}},
    copy(){if(selectedText()){target(()=>document.execCommand('copy'));}else if(!copyObjs())toast('Pilih teks atau objek untuk disalin');},
    async paste(){
      if(hasRange()){try{const t=await navigator.clipboard.readText();target(()=>document.execCommand('insertText',false,t));}catch{toast('Gunakan Ctrl+V untuk menempel');}return;}
      if(!pasteObjs())toast('Tidak ada yang ditempel');
    },
    async pastetext(){try{const t=await navigator.clipboard.readText();if(hasRange())target(()=>document.execCommand('insertText',false,t));else addObj(mk({k:'tb',x:400,y:300,w:480,h:60,html:'<div>'+esc(t)+'</div>'}),false);}catch{toast('Gunakan Ctrl+Shift+V untuk menempel');}},
    painter(){const f=fmtNow();if(!f)return toast('Pilih teks atau objek sumber format');painting=f;stg.classList.add('paint');toast('Pilih teks atau objek untuk menerapkan format');},
    newslide:newSlide,
    layout(l){sync();applyLayout(S0(),l,false);sel=[];draw();commit();},
    reset(){sync();applyLayout(S0(),S0().layout,true);sel=[];draw();commit();},
    section:secOp,
    grow(){const n=SIZES.find(s=>s>curSize());if(n)target(()=>spanStyle('fontSize',n+'pt'),true);},
    shrink(){const n=[...SIZES].reverse().find(s=>s<curSize());if(n)target(()=>spanStyle('fontSize',n+'pt'),true);},
    tshadow(){const sh=getComputedStyle(hasRange()?lastTx:(elOf(sel[0])||stg).querySelector('.tx')||stg).textShadow;target(()=>spanStyle('textShadow',sh&&sh!=='none'&&$('[data-a=tshadow]').classList.contains('on')?'none':'2px 2px 3px rgba(0,0,0,.45)'),true);},
    shape(id){drawing=id;stg.classList.add('draw');toast('Seret pada slide untuk menggambar (Esc untuk batal)');},
    arrange,smart,
    qs(i){const q=QS[i][1];forObjs(o=>{o.fill=q.fill;o.stroke=q.stroke;o.sw=q.sw;o.ts.color=q.color;});},
    nofill(){forObjs(o=>o.fill='');},
    find:showFind,replace:showFind,fnext:findNext,fclose(){fd.hidden=true;},
    frep(){const s=getSelection();if(s.toString().toLowerCase()===f1.value.toLowerCase()&&txOf(s.anchorNode)){document.execCommand('insertText',false,f2.value);commit();fk--;}findNext();},
    frall:replaceAll,
    selall(){if(hasRange()){lastTx.focus();const r=document.createRange();r.selectNodeContents(lastTx);const s=getSelection();s.removeAllRanges();s.addRange(r);}else{sel=objs().map(o=>o.id);decorate();stg.focus();}},
    selnone(){sel=[];lastR=null;getSelection().removeAllRanges();decorate();},
    notes(e){const b=$('[data-a=notes]');notes.hidden=!notes.hidden;b.classList.toggle('on',!notes.hidden);if(view==='normal')fit();},
    comments(){toast('Komentar akan tersedia di menu TINJAU');},
    vnormal(){view='normal';$('.thumbs').hidden=false;$('.ws').hidden=false;sorter.hidden=true;$$('.vb').forEach(b=>b.classList.toggle('on',b.dataset.a==='vnormal'));draw();fit();},
    vsorter(){commit();view='sorter';$('.thumbs').hidden=true;$('.ws').hidden=true;sorter.hidden=false;$$('.vb').forEach(b=>b.classList.toggle('on',b.dataset.a==='vsorter'));thumbs();},
    zoom(v){autoFit=false;setScale(+v/100);},zoomd(d){autoFit=false;setScale(Math.min(2,Math.max(.1,scale+(+d)/100)));},fit(){autoFit=true;fit();}
  };
  const sels={
    font:v=>target(()=>document.execCommand('fontName',false,v)),
    size:v=>target(()=>spanStyle('fontSize',v+'pt'),true),
    color:v=>target(()=>document.execCommand('foreColor',false,v)),
    spc:v=>v!==''&&target(()=>spanStyle('letterSpacing',v+'px'),true),
    case:changeCase,
    lh:v=>v&&forObjs(o=>o.ts.lh=+v),
    dir:v=>v&&forObjs(o=>o.vert=v),
    cols:v=>forObjs(o=>o.cols=+v),
    anchor:v=>v&&forObjs(o=>o.anchor=v),
    sfill:v=>forObjs(o=>o.fill=v),
    sline:v=>forObjs(o=>{o.stroke=v;if(!o.sw)o.sw=1;}),
    sw:v=>forObjs(o=>{o.sw=+v;if(!o.stroke)o.stroke='#41719c';}),
    dash:v=>forObjs(o=>{o.dash=v;if(!o.stroke)o.stroke='#41719c';}),
    sfx:v=>v&&forObjs(o=>o.fx=v==='none'?'':v)
  };

  /* ===== event ===== */
  root.addEventListener('mousedown',e=>{
    if(!e.target.closest('.pm'))closePop();
    if(e.target.closest('.b,.tab,.status button,.pm button,.menu button,.clr span'))e.preventDefault();
  });
  root.addEventListener('click',e=>{
    const m=e.target.closest('[data-m]');if(m){openMenu(m,m.dataset.m);return;}
    const t=e.target.closest('[data-a],[data-tab]');if(!t)return;
    if(t.dataset.tab){
      const small=matchMedia('(max-width:700px),(max-height:500px)').matches;
      if(small&&t.classList.contains('on')&&!pp.classList.contains('rc')){pp.classList.add('rc');t.classList.remove('on');return;}
      pp.classList.remove('rc');$$('.tab[data-tab]').forEach(x=>x.classList.toggle('on',x===t));$$('[data-pane]').forEach(p=>p.hidden=p.dataset.pane!==t.dataset.tab);return;}
    const i=t.dataset.a.indexOf(':'),k=i<0?t.dataset.a:t.dataset.a.slice(0,i),v=i<0?undefined:t.dataset.a.slice(i+1);
    closePop();
    if(CMDS.has(k))cmd(k);else if(acts[k])acts[k](v);
    if(k!=='menu')fm.hidden=true;
  });
  const onSel=e=>{const s=e.target.dataset&&e.target.dataset.s;if(!s||!sels[s])return;
    if(e.type==='input'&&(e.target.tagName==='SELECT'||e.target.type==='number'))return;
    if(e.type==='input'&&e.target.type==='color'&&['sfill','sline'].indexOf(s)<0&&s!=='color')return;
    sels[s](e.target.value);if(e.target.hasAttribute('data-r'))e.target.selectedIndex=0;
    if(e.target.type==='color'){const sp=e.target.previousElementSibling;if(sp)sp.style.borderBottomColor=e.target.value;}};
  root.addEventListener('change',onSel);root.addEventListener('input',onSel);
  zr.oninput=e=>{autoFit=false;setScale(+e.target.value/100);};
  $('.desk').addEventListener('mousedown',()=>{fm.hidden=true;});

  /* kanvas: seleksi, geser, ubah ukuran, gambar */
  const pt=e=>{const r=stg.getBoundingClientRect();return[(e.clientX-r.left)/scale,(e.clientY-r.top)/scale];};
  function drag(move,end){const mv=ev=>move(ev),up=ev=>{removeEventListener('mousemove',mv);removeEventListener('mouseup',up);end&&end(ev);};addEventListener('mousemove',mv);addEventListener('mouseup',up);}
  function startMove(e){
    const [sx,sy]=pt(e),l=selObjs(),o0=l.map(o=>[o.x,o.y]);let moved=false;
    drag(ev=>{const [x,y]=pt(ev),dx=x-sx,dy=y-sy;if(!moved&&Math.hypot(dx,dy)*scale<3)return;moved=true;
      l.forEach((o,i)=>{o.x=Math.round(o0[i][0]+dx);o.y=Math.round(o0[i][1]+dy);const el=elOf(o.id);el.style.left=o.x+'px';el.style.top=o.y+'px';});},
      ()=>{if(moved)commit();if(painting)applyPaint();});
  }
  function startResize(e,h){
    const o=byId(sel[0]),el=elOf(o.id),[sx,sy]=pt(e),b={x:o.x,y:o.y,w:o.w,h:o.h};
    drag(ev=>{const [x,y]=pt(ev),dx=x-sx,dy=y-sy;let {x:nx,y:ny,w:nw,h:nh}=b;
      if(h.includes('e'))nw=Math.max(12,b.w+dx);if(h.includes('s'))nh=Math.max(12,b.h+dy);
      if(h.includes('w')){nw=Math.max(12,b.w-dx);nx=b.x+b.w-nw;}if(h.includes('n')){nh=Math.max(12,b.h-dy);ny=b.y+b.h-nh;}
      Object.assign(o,{x:Math.round(nx),y:Math.round(ny),w:Math.round(nw),h:Math.round(nh)});
      Object.assign(el.style,{left:o.x+'px',top:o.y+'px',width:o.w+'px',height:o.h+'px'});
      const sv=el.querySelector('svg.bg');sv.setAttribute('viewBox',`0 0 ${o.w} ${o.h}`);sv.innerHTML=shapeSVG(o);},commit);
  }
  function startDraw(e){
    const [sx,sy]=pt(e),id=drawing,mq=document.createElement('div');mq.className='mq';stg.appendChild(mq);
    drag(ev=>{const [x,y]=pt(ev);Object.assign(mq.style,{left:Math.min(sx,x)+'px',top:Math.min(sy,y)+'px',width:Math.abs(x-sx)+'px',height:Math.abs(y-sy)+'px'});},
      ev=>{const [x,y]=pt(ev);mq.remove();drawing=null;stg.classList.remove('draw');
        let w=Math.abs(x-sx),h=Math.abs(y-sy),px=Math.min(sx,x),py=Math.min(sy,y);const click=w<8&&h<8;
        if(id==='tb'){if(click){w=360;h=60;}addObj(mk({k:'tb',x:Math.round(px),y:Math.round(py),w:Math.round(Math.max(w,60)),h:Math.round(Math.max(h,40))}),true);return;}
        const ln=id==='line'||id==='arrow';
        if(click){w=ln?300:220;h=ln?24:ln?24:140;px=sx-w/2;py=sy-h/2;}else if(ln)h=Math.max(24,h);
        addObj(mk(ln?{k:'sh',shape:id,x:Math.round(px),y:Math.round(py),w:Math.round(w),h:24,stroke:'#5b9bd5',sw:2.25}:
          {k:'sh',shape:id,x:Math.round(px),y:Math.round(py),w:Math.round(Math.max(w,24)),h:Math.round(Math.max(h,24)),fill:'#5b9bd5',stroke:'#41719c',sw:1,anchor:'m',ts:{color:'#fff',align:'center'}}),false);});
  }
  function startMarquee(e,add){
    const [sx,sy]=pt(e),mq=document.createElement('div');mq.className='mq';stg.appendChild(mq);const base=add?[...sel]:[];
    drag(ev=>{const [x,y]=pt(ev),l=Math.min(sx,x),t=Math.min(sy,y),r=Math.max(sx,x),b=Math.max(sy,y);
      Object.assign(mq.style,{left:l+'px',top:t+'px',width:r-l+'px',height:b-t+'px'});
      sel=[...new Set([...base,...objs().filter(o=>o.x<r&&o.x+o.w>l&&o.y<b&&o.y+o.h>t).map(o=>o.id)])];decorate();},()=>mq.remove());
  }
  stg.addEventListener('mousedown',e=>{
    if(e.button)return;
    if(drawing){e.preventDefault();startDraw(e);return;}
    const hd=e.target.closest('.hd');if(hd){e.preventDefault();startResize(e,hd.dataset.h);return;}
    const el=e.target.closest('.ob');
    if(!el){getSelection().removeAllRanges();lastR=null;if(document.activeElement&&document.activeElement!==document.body)document.activeElement.blur();
      if(!e.shiftKey){sel=[];decorate();}e.preventDefault();stg.focus({preventScroll:true});startMarquee(e,e.shiftKey);if(painting)applyPaint();return;}
    const o=byId(el.dataset.id),ed=e.target.closest('.ed'),inTx=!!e.target.closest('.tx')&&(o.k!=='sh'||el.classList.contains('edt'));
    if(ed||(o.k==='sh'&&!el.classList.contains('edt'))){
      e.preventDefault();
      if(e.shiftKey)sel=sel.includes(o.id)?sel.filter(x=>x!==o.id):[...sel,o.id];else if(!sel.includes(o.id))sel=[o.id];
      getSelection().removeAllRanges();lastR=null;decorate();stg.focus({preventScroll:true});startMove(e);return;}
    if(!inTx){e.preventDefault();const t=el.querySelector('.tx');if(t){sel=[o.id];decorate();t.focus();const r=document.createRange();r.selectNodeContents(t);r.collapse(false);const s=getSelection();s.removeAllRanges();s.addRange(r);}return;}
    if(sel.length!==1||sel[0]!==o.id){sel=[o.id];decorate();}
  });
  stg.addEventListener('dblclick',e=>{const el=e.target.closest('.ob'),o=el&&byId(el.dataset.id);
    if(o&&o.k==='sh'&&o.shape!=='line'&&o.shape!=='arrow'){el.classList.add('edt');const t=el.querySelector('.tx');t.focus();const r=document.createRange();r.selectNodeContents(t);r.collapse(false);const s=getSelection();s.removeAllRanges();s.addRange(r);}});
  stg.addEventListener('mouseup',e=>{if(painting&&(selectedText()||sel.length))setTimeout(applyPaint,0);syncUI();});
  stg.addEventListener('focusin',e=>{const t=e.target.closest&&e.target.closest('.tx');if(!t)return;const el=t.closest('.ob'),o=byId(el.dataset.id);if(!o)return;
    if(sel.length!==1||sel[0]!==o.id){sel=[o.id];decorate();}
    if(o.bul&&!t.textContent.trim()&&!t.querySelector('li')){t.innerHTML='<ul><li><br></li></ul>';const li=t.querySelector('li'),r=document.createRange();r.selectNodeContents(li);r.collapse(true);const s=getSelection();s.removeAllRanges();s.addRange(r);}});
  stg.addEventListener('focusout',e=>{const t=e.target.closest&&e.target.closest('.tx');if(t&&!t.textContent.trim()&&!t.querySelector('img'))t.innerHTML='';});
  stg.addEventListener('input',e=>{if(e.target.closest('.tx'))commitSoon();});
  stg.addEventListener('paste',e=>{const t=e.target.closest&&e.target.closest('.tx');if(!t)return;e.preventDefault();const x=(e.clipboardData||window.clipboardData).getData('text/plain');if(x)document.execCommand('insertText',false,x);});
  stg.addEventListener('keydown',e=>{
    const intx=e.target.closest&&e.target.closest('.tx');
    if(e.key==='Escape'){if(drawing){drawing=null;stg.classList.remove('draw');return;}if(painting){painting=null;stg.classList.remove('paint');return;}
      if(intx){e.preventDefault();intx.blur();getSelection().removeAllRanges();lastR=null;stg.focus({preventScroll:true});}else{sel=[];decorate();}return;}
    if(intx||!sel.length)return;
    const k=e.key;
    if(k==='Delete'||k==='Backspace'){e.preventDefault();delSel();}
    else if(k==='Enter'){const t=elOf(sel[0]).querySelector('.tx');if(t){e.preventDefault();t.focus();const r=document.createRange();r.selectNodeContents(t);r.collapse(false);const s=getSelection();s.removeAllRanges();s.addRange(r);}}
    else if(k.startsWith('Arrow')){e.preventDefault();const st=e.ctrlKey?1:8,d={ArrowLeft:[-st,0],ArrowRight:[st,0],ArrowUp:[0,-st],ArrowDown:[0,st]}[k];
      selObjs().forEach(o=>{o.x+=d[0];o.y+=d[1];const el=elOf(o.id);el.style.left=o.x+'px';el.style.top=o.y+'px';});commitSoon();}
  });
  notes.addEventListener('input',()=>{S0().notes=notes.textContent.trim()?notes.innerHTML:'';commitSoon();});

  /* panel slide & penyortir */
  const listClick=e=>{const t=e.target.closest('.th,.th2,.sec');if(!t)return;go(+(t.dataset.i??t.dataset.sec));};
  thumbsEl.addEventListener('click',listClick);sorter.addEventListener('click',listClick);
  sorter.addEventListener('dblclick',e=>{if(e.target.closest('.th2'))acts.vnormal();});
  const ctx=e=>{const t=e.target.closest('.th,.th2');if(!t)return;e.preventDefault();go(+t.dataset.i);
    pop([['Slide Baru','newslide'],['Duplikat Slide','dup'],['Hapus Slide','delslide'],'-',['Mulai dari Slide Ini','show']],e.clientX,e.clientY);};
  thumbsEl.addEventListener('contextmenu',ctx);sorter.addEventListener('contextmenu',ctx);
  acts.dup=()=>{sync();const c=JSON.parse(JSON.stringify(S0()));delete c.sec;c.objs.forEach(o=>o.id='o'+(uid++));slides.splice(cur+1,0,c);cur++;sel=[];draw();notes.innerHTML=c.notes||'';commit();};
  acts.delslide=()=>{sync();if(slides.length===1)slides=[mkSlide('title')];else{slides.splice(cur,1);if(cur>=slides.length)cur=slides.length-1;}sel=[];draw();notes.innerHTML=S0().notes||'';commit();};
  thumbsEl.addEventListener('keydown',e=>{
    if(e.key==='Delete'||e.key==='Backspace'){e.preventDefault();acts.delslide();}
    else if(e.key==='Enter'){e.preventDefault();newSlide();}
    else if(e.key==='ArrowDown'||e.key==='ArrowRight'){e.preventDefault();go(cur+1);}
    else if(e.key==='ArrowUp'||e.key==='ArrowLeft'){e.preventDefault();go(cur-1);}});
  const dnd=(host)=>{
    host.addEventListener('dragstart',e=>{const t=e.target.closest('.th,.th2');if(!t)return;sync();sortFrom=+t.dataset.i;e.dataTransfer.effectAllowed='move';e.dataTransfer.setData('text/plain',String(sortFrom));});
    host.addEventListener('dragover',e=>{if(sortFrom==null)return;const t=e.target.closest('.th,.th2');if(!t)return;e.preventDefault();$$('.dr').forEach(x=>x.classList.remove('dr'));t.classList.add('dr');});
    host.addEventListener('drop',e=>{const t=e.target.closest('.th,.th2');if(!t||sortFrom==null)return;e.preventDefault();const to=+t.dataset.i,from=sortFrom;sortFrom=null;
      if(to!==from){const [m]=slides.splice(from,1);slides.splice(to,0,m);cur=to;sel=[];draw();commit();}$$('.dr').forEach(x=>x.classList.remove('dr'));});
    host.addEventListener('dragend',()=>{sortFrom=null;$$('.dr').forEach(x=>x.classList.remove('dr'));});
  };
  dnd(thumbsEl);dnd(sorter);

  root.addEventListener('keydown',e=>{
    if(!(e.ctrlKey||e.metaKey))return;const k=e.key.toLowerCase(),intx=e.target.closest&&e.target.closest('.tx'),inInput=e.target.matches&&e.target.matches('input[type=text],input[type=number],.notes');
    if(k==='z'&&!inInput){e.preventDefault();undo();}else if(k==='y'&&!inInput){e.preventDefault();redo();}
    else if(k==='s'){e.preventDefault();acts.save();}
    else if(k==='o'){e.preventDefault();acts.open();}
    else if(k==='f'||k==='h'){e.preventDefault();showFind();}
    else if(k==='m'){e.preventDefault();newSlide();}
    else if(!intx&&!inInput&&sel.length&&(k==='c'||k==='x'||k==='v'||k==='d')){
      if(k==='c'){e.preventDefault();copyObjs();}else if(k==='x'){e.preventDefault();if(copyObjs())delSel();}
      else if(k==='v'&&clip){e.preventDefault();pasteObjs();}else if(k==='d'){e.preventDefault();if(copyObjs())pasteObjs();}}
  });
  f1.addEventListener('keydown',e=>{if(e.key==='Enter')findNext();});f1.addEventListener('input',()=>fk=-1);
  if(window.ResizeObserver){let rt=0;new ResizeObserver(()=>{clearTimeout(rt);rt=setTimeout(()=>{if(root.isConnected&&autoFit&&view==='normal')fit();},80);}).observe(desk);}

  /* mulai */
  document.execCommand('styleWithCSS',false,true);
  $('[data-s=font]').value='Calibri';$('[data-s=size]').value=18;
  if(matchMedia('(max-width:700px),(max-height:500px)').matches){pp.classList.add('rc');$$('.tab[data-tab]').forEach(x=>x.classList.remove('on'));}
  setName();draw();thumbs();fit();commit();
  if(opts.file)openFile(opts.file);
}

window.Apps=window.Apps||{};
Apps.ppt={mount};
})();
