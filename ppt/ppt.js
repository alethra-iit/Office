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
.pp .b:disabled{opacity:.38;pointer-events:none}
.pp .b.big em svg{display:block}
.pp .ibig{flex-direction:column;height:68px;min-width:50px;gap:2px;font-size:11px;line-height:1.1;white-space:normal;text-align:center;padding:2px 3px}
.pp .ibig em{font-size:26px;line-height:1.1}
.pp .ibig span{display:block}
.pp .ob.tbl .tx{flex:1 1 auto;height:100%}
.pp .ob.tbl .tx table{height:100%}
.pp .tx table,.pp-print .tx table{border-collapse:collapse;width:100%;table-layout:fixed;font-size:inherit}
.pp .tx td,.pp-print .tx td{border:1px solid #fff;padding:4px 8px;vertical-align:top;height:40px;overflow-wrap:anywhere}
.pp .tx a,.pp-print .tx a{color:#0563c1;text-decoration:underline}
.pp .ob[data-lk]{cursor:pointer}
.pp .stg .ob[data-lk]{cursor:default}
.pp .med{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;background:#222;color:#fff;font-size:15px;gap:10px;overflow:hidden}
.pp .med video,.pp .med iframe{width:100%;height:100%;border:0;background:#000}
.pp .med audio{width:100%}
.pp .med.au{background:#eef1f5;color:#333}
.pp .med em{font-size:34px}
.pp .stg .ob.mdo .med{pointer-events:none}
.pp .cm{position:absolute;width:26px;height:26px;background:#ffc000;border:1px solid #b88a00;border-radius:3px 3px 3px 0;display:grid;place-items:center;font-size:13px;cursor:pointer;z-index:6;transform:scale(var(--inv));transform-origin:0 0}
.pp .pm.sg{grid-template-columns:repeat(6,34px);gap:1px;padding:4px;min-width:0}
.pp .pm.sg button{flex-direction:column;justify-content:center;padding:0;height:30px;width:34px}
.pp .pm.sg button span{display:none}
.pp .pm.wa{grid-template-columns:repeat(4,64px);gap:3px;padding:6px;min-width:0}
.pp .pm.wa button{padding:6px 0;justify-content:center;font-size:30px;line-height:1.1;font-weight:700;height:50px}
.pp .pm.wa button span{display:none}
.pp .tbg{padding:6px}
.pp .tbg .gr{display:grid;grid-template-columns:repeat(10,18px);gap:2px}
.pp .tbg .gr i{width:18px;height:18px;border:1px solid #aaa;background:#fff;display:block}
.pp .tbg .gr i.on{background:#fbe3da;border-color:var(--acc)}
.pp .tbg .lb{margin:4px 0;text-align:center;color:#555;min-height:16px}
.pp .tbg hr{border:0;border-top:1px solid #ddd;margin:4px 0}
.pp .tbg button{display:block;width:100%;text-align:left;background:none;border:0;padding:4px 6px;cursor:pointer}
.pp .tbg button:hover{background:#fbe3da}
.pp .dlgbg{position:absolute;inset:0;z-index:70;background:rgba(0,0,0,.25);display:grid;place-items:center}
.pp .dlg{background:#fff;border:1px solid #888;box-shadow:0 6px 24px rgba(0,0,0,.35);padding:14px 16px;min-width:300px;max-width:min(520px,calc(100% - 24px));max-height:calc(100% - 24px);overflow:auto;font-size:13px;user-select:text;-webkit-user-select:text}
.pp .dlg h3{margin:0 0 10px;font-size:15px;font-weight:600;color:var(--acc)}
.pp .dlg label{display:flex;gap:6px;align-items:center;margin:5px 0}
.pp .dlg input[type=text]{height:26px;border:1px solid #aaa;padding:0 6px;flex:1;min-width:0}
.pp .dlg select{height:26px;width:100%}
.pp .dlg .bt{display:flex;gap:6px;justify-content:flex-end;margin-top:12px;flex-wrap:wrap}
.pp .dlg .bt button{height:28px;border:1px solid #aaa;background:#f3f3f3;cursor:pointer;padding:0 12px}
.pp .dlg .bt button.pri{background:var(--acc);border-color:var(--acc);color:#fff}
.pp .dlg .sy{display:grid;grid-template-columns:repeat(10,1fr);gap:2px;margin-top:8px}
.pp .dlg .sy button{height:32px;font-size:18px;border:1px solid #ddd;background:#fff;cursor:pointer;padding:0}
.pp .dlg .sy button:hover{background:#fbe3da;border-color:var(--acc)}
.pp .dlg .dl{border:1px solid #ccc;max-height:200px;overflow:auto}
.pp .dlg .dl label{margin:0;padding:5px 8px}
.pp .dlg .dl label:hover{background:#fbe3da}
.pp .tgal{display:flex;gap:4px;width:600px;overflow-x:auto;scrollbar-width:none;padding:2px}
.pp .tgal::-webkit-scrollbar{display:none}
.pp .vgal{display:flex;gap:4px;padding:2px}
.pp .tt{flex:none;position:relative;width:106px;height:70px;border:1px solid #bbb;background:#fff;padding:0;cursor:pointer;overflow:hidden}
.pp .tt:hover{border-color:var(--acc)}.pp .tt.on{outline:2px solid #e8b5a3;outline-offset:1px;border-color:var(--acc)}
.pp .tb{position:absolute;inset:0;display:block}
.pp .tb>i.d{position:absolute;display:block}
.pp .tb>b{position:absolute;left:10px;top:12px;font-size:24px;font-weight:400;line-height:1}
.pp .tb>.sws{position:absolute;left:8px;bottom:6px;display:flex;gap:2px}
.pp .tb>.sws i{display:block;width:11px;height:5px}
.pp .tbx{display:block;position:relative;width:88px;height:56px;border:1px solid #bbb}
.pp .pm.thm{grid-template-columns:repeat(4,100px)}
.pp .trg{width:640px}
.pp .trb{flex:none;width:76px;height:66px;border:1px solid transparent;background:none;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;font-size:11px;cursor:pointer;padding:2px;white-space:normal;text-align:center}
.pp .trb:hover{background:#fdeee8;border-color:#e8b5a3}.pp .trb.on{background:#f9d3c5;border-color:#e39b83}
.pp .tmw .row{gap:4px}.pp .tmw label{display:flex;align-items:center;gap:4px}.pp .tmw input[type=number]{width:62px}
.pp .trm{color:var(--acc);font-size:10px;font-style:normal}
.pp .trb:disabled{opacity:.45;filter:grayscale(1);pointer-events:none}
.pp .apane{width:230px;flex:none;border-left:1px solid #ddd;background:#fafafa;display:flex;flex-direction:column;font-size:12px}
.pp .apane .ah{display:flex;justify-content:space-between;align-items:center;padding:6px 8px;border-bottom:1px solid #ddd;font-weight:600}
.pp .apane .ah button{height:24px;border:1px solid #aaa;background:#f3f3f3;cursor:pointer;padding:0 8px}
.pp .apane .apl{flex:1;overflow:auto}
.pp .apane .ai{display:flex;gap:6px;align-items:center;padding:5px 8px;cursor:pointer;border-bottom:1px solid #eee}
.pp .apane .ai.on{background:#f9d3c5}.pp .apane .ai i{font-style:normal;color:#555;width:14px}
.pp .apane .ai b{font-weight:400;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1}
.pp .apane .ai em{color:#3d7a3d}.pp .apane .am{padding:10px;color:#888}
.pp .anb{display:none;position:absolute;min-width:22px;height:20px;padding:0 4px;border:1px solid #888;background:#fff;color:#444;font-size:12px;line-height:18px;text-align:center;z-index:6;pointer-events:none;transform:scale(var(--inv));transform-origin:0 0}
.pp.ant .anb,.pp.apo .anb{display:block}
.pp .th.hid .tv,.pp .th2.hid .tv{opacity:.4}
.pp .th.hid .n,.pp .th2.hid .n{text-decoration:line-through}
.pp .show{display:grid;place-items:center}
.pp .show.win{inset:6vh 8vw;border:1px solid #777;box-shadow:0 8px 40px rgba(0,0,0,.6)}
.pp .showbk{position:fixed;inset:0;z-index:99;background:rgba(0,0,0,.5)}
.pp .shh{display:grid;place-items:center;width:100%;height:100%;min-width:0;min-height:0;overflow:hidden}
.pp .show.pv{grid-template-columns:minmax(0,1fr) 300px;place-items:stretch}
.pp .pvs{background:#1f1f1f;color:#eee;padding:10px;display:flex;flex-direction:column;gap:8px;min-height:0;overflow:hidden;font-size:13px}
.pp .pvn{position:relative;width:100%;aspect-ratio:16/9;background:#fff;overflow:hidden;flex:none;display:grid;place-items:center;color:#555}
.pp .pvx{flex:1;min-height:0;overflow:auto;background:#2b2b2b;padding:8px;font-size:15px;line-height:1.35;user-select:text;-webkit-user-select:text}
.pp .pvi{display:flex;justify-content:space-between;gap:8px;flex:none}
.pp .shtb,.pp .shrec{position:absolute;display:flex;gap:2px;align-items:center;background:rgba(30,30,30,.85);color:#fff;padding:3px 6px;border-radius:4px;z-index:3;font-size:14px}
.pp .shtb{left:10px;bottom:10px;opacity:0;pointer-events:none;transition:opacity .2s}
.pp .show.mv .shtb{opacity:1;pointer-events:auto}
.pp .shrec{top:10px;left:50%;transform:translateX(-50%);gap:8px}
.pp .shtb button,.pp .shrec button{background:none;border:0;color:#fff;font-size:15px;min-width:28px;height:26px;cursor:pointer;border-radius:3px}
.pp .shtb button:hover,.pp .shrec button:hover,.pp .shtb button.on{background:rgba(255,255,255,.25)}
.pp .shrec .rt,.pp .shrec .rtt{font-variant-numeric:tabular-nums;min-width:38px;text-align:center}
.pp .shrec .rtt{color:#bbb}
.pp .rdot{color:#e33;animation:pprd 1s infinite}
@keyframes pprd{50%{opacity:.2}}
.pp .shink{position:absolute;pointer-events:none;z-index:2}
.pp .show.tpen .shink{pointer-events:auto;cursor:crosshair}
.pp .shlz{position:absolute;width:20px;height:20px;margin:-10px 0 0 -10px;border-radius:50%;background:radial-gradient(#f00 25%,rgba(255,0,0,.4) 55%,transparent 70%);pointer-events:none;display:none;z-index:4}
.pp .show.tlaser .shlz{display:block}
.pp .show.tlaser,.pp .show:not(.mv):not(.tpen){cursor:none}
.pp .shcv{position:absolute;inset:0;z-index:5}
.pp .dlg fieldset{border:1px solid #ccc;margin:6px 0;padding:2px 10px 4px}
.pp .dlg legend{padding:0 4px;color:#666}
.pp .dlg input[type=number]{width:60px;height:24px;border:1px solid #aaa}
.pp .dlg select[multiple]{height:150px}
.pp .dlg .cs2{display:grid;grid-template-columns:1fr auto 1fr;gap:8px;align-items:center;margin:8px 0}
.pp .dlg .cs2 .cb{display:flex;flex-direction:column;gap:4px}
.pp .dlg .cs2 button{height:26px;border:1px solid #aaa;background:#f3f3f3;cursor:pointer;padding:0 8px}
.pp .shsel{height:22px;width:150px}
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
  return `<div class="ob ${o.k}${o.media?' mdo':''}${/<table/i.test(o.html||'')?' tbl':''}" data-id="${o.id}"${o.link?' data-lk="1"':''} style="left:${o.x}px;top:${o.y}px;width:${o.w}px;height:${o.h}px;transform:rotate(${o.rot||0}deg);justify-content:${AN[o.anchor]||'flex-start'}">`+
    `<svg class="bg" viewBox="0 0 ${Math.max(1,o.w)} ${Math.max(1,o.h)}" style="${fxCSS(o)}">${shapeSVG(o)}</svg>`+
    (o.img?`<img class="pic" src="${o.img}" alt="" draggable="false" style="${fxCSS(o)}">`:'')+
    (o.media?mediaHTML(o,edit):'')+
    (noTx||o.media?'':`<div class="tx"${edit===true?' contenteditable="true" spellcheck="true"':''} data-ph="${esc(o.ph||'')}" style="${txCSS(o)}">${o.html}</div>`)+`</div>`;
}
/* media: pemutar sungguhan hanya saat slide show (edit==='show'); di kanvas/thumbnail berupa penanda */
function mediaHTML(o,mode){
  const m=o.media,au=m.type==='audio';
  if(mode==='show'){
    if(m.embed)return `<div class="med"><iframe src="${esc(m.src)}" allow="autoplay; encrypted-media; fullscreen" allowfullscreen></iframe></div>`;
    return au?`<div class="med au"><audio controls src="${esc(m.src)}"></audio></div>`:`<div class="med"><video controls playsinline src="${esc(m.src)}"></video></div>`;}
  return `<div class="med${au?' au':''}"><em>${au?'🔊':'▶'}</em><span>${esc(m.name||(au?'Audio':'Video'))}</span></div>`;
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

/* ---- tab SISIPKAN ---- */
const IB=(ic,tip,a,need,lab)=>`<button class="b big" title="${tip}" data-a="${a}"${need?` data-need="${need}"`:''}><em>${ic}</em><span>${lab||tip}</span></button>`;
const IM=(ic,tip,m,lab)=>`<button class="b big" title="${tip}" data-m="${m}"><em>${ic}</em><span>${lab||tip} ▾</span></button>`;
const SPLIT=(ic,tip,a,m,lab)=>`<div class="sp1"><button class="b big" title="${tip}" data-a="${a}"><em>${ic}</em></button><button class="b sm" data-m="${m}">${lab} ▾</button></div>`;
const ICO={
  table:'<svg width="28" height="26" viewBox="0 0 28 26"><rect x="2" y="3" width="24" height="20" fill="#fff" stroke="#4a5f8a"/><rect x="2" y="3" width="24" height="6" fill="#4a73b5"/><path d="M2 9.5h24M2 16h24M10 3v20M18 3v20" stroke="#4a5f8a"/></svg>',
  chart:'<svg width="28" height="26" viewBox="0 0 28 26"><rect x="3" y="12" width="6" height="12" fill="#5b9bd5"/><rect x="11" y="5" width="6" height="19" fill="#ed7d31"/><rect x="19" y="9" width="6" height="15" fill="#a5a5a5"/><path d="M2 24.5h25" stroke="#555"/></svg>',
  smart:'<svg width="28" height="26" viewBox="0 0 28 26"><rect x="2" y="3" width="11" height="7" fill="#4a9a6a" stroke="#2c6e47"/><rect x="15" y="3" width="11" height="7" fill="#5b9bd5" stroke="#41719c"/><rect x="2" y="14" width="24" height="8" fill="#9fc3e6" stroke="#41719c"/></svg>',
  apps:'<svg width="28" height="26" viewBox="0 0 28 26"><polygon points="14,2 24,8 24,19 14,25 4,19 4,8" fill="#0a64c8"/><polygon points="14,2 24,8 14,13 4,8" fill="#3b8be0"/></svg>',
  hdr:'<svg width="28" height="26" viewBox="0 0 28 26"><rect x="4" y="2" width="20" height="22" fill="#fff" stroke="#777"/><rect x="4" y="2" width="20" height="5" fill="#f1c9a8" stroke="#777"/><path d="M7 11h14M7 14h14M7 17h14" stroke="#aaa"/></svg>',
  wa:'<svg width="28" height="26" viewBox="0 0 28 26"><text x="14" y="22" font-size="26" font-weight="700" font-style="italic" text-anchor="middle" fill="#2b579a" font-family="Georgia,serif">A</text></svg>'
};
const RIB_INSERT=
 G('Slide',SPLIT('🗎','Slide Baru (Ctrl+M)','newslide','newslide','Slide Baru'))+
 G('Tabel',IM(ICO.table,'Tabel','table'))+
 G('Gambar',IB('🖼','Gambar dari berkas','pic',0,'Gambar')+IB('🌐','Gambar online (dari URL)','onlinepic',0,'Gambar<br>Online')+IM('📷','Cuplikan layar','screenshot','Cuplikan<br>Layar')+IM('🗂','Album foto','album','Album<br>Foto'))+
 G('Ilustrasi',IM('◇','Bentuk','shapes')+IM(ICO.smart,'SmartArt','smartins')+IM(ICO.chart,'Bagan','chartm'))+
 G('Aplikasi',IM(ICO.apps,'Aplikasi untuk Office','apps','Aplikasi<br>untuk Office'))+
 G('Tautan',IB('🔗','Hyperlink (Ctrl+K)','hlink','link','Hyperlink')+IB('★','Tindakan','action','obj','Tindakan'))+
 G('Komentar',IB('💬','Komentar','comment',0,'Komentar'))+
 G('Teks',SPLIT('<span style="font-family:Georgia;border:1px solid #4a5f8a;padding:0 5px;font-size:20px">A</span>','Kotak Teks','shape:tb','textbox','Kotak<br>Teks')+IB(ICO.hdr,'Header &amp; Footer','hf',0,'Header<br>&amp; Footer')+IM(ICO.wa,'WordArt','wordart')+IB('📅','Tanggal &amp; Waktu','datetime',0,'Tanggal<br>&amp; Waktu')+IB('#','Nomor Slide','slidenum',0,'Nomor<br>Slide')+IB('▣','Objek (gambar, teks, CSV)','object',0,'Objek'))+
 G('Simbol',IM('<span style="color:#2b579a;font-family:Cambria,serif">π</span>','Persamaan','eqm')+IB('Ω','Simbol','symbol','text','Simbol'))+
 G('Media',IM('🎞','Video','video')+IM('🔊','Audio','audio'));

/* ---- tab ANIMASI ---- */
const AEFF=[['none','Tidak Ada'],['appear','Muncul'],['fade','Pudar'],['fly','Terbang Masuk'],['float','Melayang Masuk'],['split','Belah'],['wipe','Sapu'],['shape','Bentuk'],['wheel','Roda'],['bars','Batang Acak'],['grow','Membesar & Putar']];
const FROMS=[['b','Dari Bawah'],['l','Dari Kiri'],['t','Dari Atas'],['r','Dari Kanan']];
const AOPT={fly:FROMS,wipe:FROMS,float:[['u','Naik'],['d','Turun']],split:[['v','Vertikal Keluar'],['h','Horizontal Keluar']],shape:[['c','Lingkaran'],['d','Belah Ketupat']],bars:[['h','Horizontal'],['v','Vertikal']]};
const getAn=o=>Object.assign({t:'none',o:'',x:0,start:'click',dur:.5,delay:0},o&&o.an||{});
const hasAn=o=>!!(o.an&&o.an.t&&o.an.t!=='none');
const objName=o=>{const t=textOf(o.html||'').trim();return t?t.slice(0,24):({title:'Judul',sub:'Subjudul',body:'Teks Isi'}[o.role]||(o.img?'Gambar':o.media?'Media':o.shape&&o.shape!=='rect'?'Bentuk':'Objek'));};
function anList(s){const have=s.objs.filter(hasAn),ord=(s.ao||[]).map(id=>have.find(o=>o.id===id)).filter(Boolean);return[...ord,...have.filter(o=>!ord.includes(o))];}
function anSteps(s){
  const steps=[];let st=null,pS=0,pE=0;
  anList(s).forEach(o=>{const a=getAn(o),D=(a.dur>0?+a.dur:.5)*1000,dl=(+a.delay||0)*1000;let at;
    if(a.start==='click'||!st){st={auto:a.start!=='click',items:[],ms:0};steps.push(st);at=dl;}
    else at=(a.start==='after'?pE:pS)+dl;
    st.items.push({o,a,at,D});pS=at;pE=at+D;st.ms=Math.max(st.ms,pE);});
  return steps;
}
function barsKF(op){const n=8,idx=[...Array(n).keys()].sort(()=>Math.random()-.5),half=new Set(idx.slice(0,n/2));
  const poly=f=>'polygon('+[...Array(n)].map((_,i)=>{const a=i*100/n,A=(i+1)*100/n,g=f(i);return op==='v'?`${a}% 0%,${a}% ${g}%,${A}% ${g}%,${A}% 0%`:`0% ${a}%,${g}% ${a}%,${g}% ${A}%,0% ${A}%`;}).join(',')+')';
  return[{clipPath:poly(()=>0),offset:0},{clipPath:poly(i=>half.has(i)?100:0),offset:.5},{clipPath:poly(()=>100),offset:1}];}
function anKF(o,a){
  const b=`rotate(${o.rot||0}deg)`,op=a.o||(AOPT[a.t]||[['']])[0][0],T=f=>f+' '+b;
  switch(a.t){
    case'appear':return[{opacity:0,offset:0},{opacity:1,offset:.001},{opacity:1,offset:1}];
    case'fly':{const f={l:`translateX(${-(o.x+o.w)}px)`,r:`translateX(${W-o.x}px)`,t:`translateY(${-(o.y+o.h)}px)`,b:`translateY(${H-o.y}px)`}[op];
      return[{transform:T(f),opacity:0,offset:0},{opacity:1,offset:.12},{transform:b,opacity:1,offset:1}];}
    case'float':return[{transform:T(`translateY(${op==='d'?-50:50}px)`),opacity:0},{transform:b,opacity:1}];
    case'split':return[{clipPath:op==='h'?'inset(50% 0 50% 0)':'inset(0 50% 0 50%)'},{clipPath:'inset(0 0 0 0)'}];
    case'wipe':return[{clipPath:{b:'inset(100% 0 0 0)',l:'inset(0 100% 0 0)',t:'inset(0 0 100% 0)',r:'inset(0 0 0 100%)'}[op]},{clipPath:'inset(0 0 0 0)'}];
    case'shape':return op==='d'?[{clipPath:'polygon(50% 50%,50% 50%,50% 50%,50% 50%)'},{clipPath:'polygon(50% -50%,150% 50%,50% 150%,-50% 50%)'}]:[{clipPath:'circle(0% at 50% 50%)'},{clipPath:'circle(75% at 50% 50%)'}];
    case'wheel':return[{transform:T('rotate(-360deg) scale(.1)'),opacity:0},{transform:b,opacity:1}];
    case'bars':return barsKF(op);
    case'grow':return[{transform:T('scale(.1) rotate(-90deg)'),opacity:0},{transform:b,opacity:1}];
    default:return[{opacity:0},{opacity:1}];
  }
}
const anEl=(root,o)=>root.querySelector(`.ob[data-id="${o.id}"]`);
function anInit(root,s,only){s.objs.forEach(o=>{if(!hasAn(o)||o.an.x||(only&&!only.includes(o)))return;const e=anEl(root,o);if(e)e.style.visibility='hidden';});}
function anPlayItems(root,items){items.forEach(({o,a,at,D})=>{const e=anEl(root,o);if(!e)return;
  e.animate(anKF(o,a).map(f=>Object.assign({visibility:'visible'},f)),{duration:D,delay:at,fill:'both',easing:'ease-out',direction:a.x?'reverse':'normal'});});}
const aIcon=t=>{const M={appear:'<path d="M3 3l4 4M31 3l-4 4M3 29l4-4M31 29l-4-4" stroke="#3d7a3d" stroke-width="1.6"/>',fade:'<circle cx="17" cy="16" r="15" fill="none" stroke="#7bbd7b" stroke-dasharray="2 2"/>',
  fly:'<path d="M1 30l9-9M10 21v5M10 21h5" stroke="#3d7a3d" stroke-width="1.6" fill="none"/>',float:'<path d="M4 31v-9M1 25l3-3 3 3" stroke="#3d7a3d" stroke-width="1.6" fill="none"/>',
  split:'<path d="M17 4v24" stroke="#fff" stroke-width="2"/>',wipe:'<path d="M1 16h9M7 12l4 4-4 4" stroke="#3d7a3d" stroke-width="1.6" fill="none"/>',shape:'<circle cx="17" cy="16" r="5" fill="#fff"/>',
  wheel:'<path d="M17 16V6M17 16l9 5" stroke="#fff" stroke-width="2"/>',bars:'<path d="M5 12h24M5 17h24M5 22h24" stroke="#fff" stroke-width="1.2"/>',grow:'<path d="M24 24l7 7M31 24v7h-7" stroke="#3d7a3d" stroke-width="1.6" fill="none"/>'};
  return `<svg width="34" height="32" viewBox="0 0 34 32"><path d="M17 2l4.2 8.6 9.3 1.3-6.8 6.6 1.7 9.3L17 23.3l-8.4 4.5 1.7-9.3-6.8-6.6 9.3-1.3z" fill="${t==='none'?'#d0d0d0':'#5ba85b'}" stroke="${t==='none'?'#aaa':'#3d7a3d'}" stroke-width=".8"/>${M[t]||''}</svg>`;};
const RIB_ANIM=
 G('Pratinjau',IB('★','Pratinjau animasi','anprev',0,'Pratinjau'))+
 G('Animasi',`<div class="tgal trg">${AEFF.map(([id,n])=>`<button class="trb anb2" data-an="${id}" data-a="anset:${id}" title="${n}" disabled>${aIcon(id)}<span>${n}</span></button>`).join('')}</div>`+
   `<div class="rows"><button class="b" data-a="anscroll:-1" title="Sebelumnya">◀</button><button class="b" data-a="anscroll:1" title="Berikutnya">▶</button></div>`+
   `<button class="b big aneff" data-m="aneff" disabled title="Opsi Efek"><em>⚙</em><span>Opsi Efek ▾</span></button>`)+
 G('Animasi Lanjutan',`<button class="b big" data-m="anadd" title="Tambah Animasi"><em>✦</em><span>Tambah<br>Animasi ▾</span></button><div class="rows">${B('▤','Panel Animasi','anpane','rw')}${B('🖌','Pelukis Animasi','anpaint','rw')}</div>`)+
 G('Pengaturan Waktu',`<div class="rows tmw"><div class="row">Mulai: ${S('anstart',[['click','Saat Klik'],['with','Bersama Sebelumnya'],['after','Setelah Sebelumnya']],150,'Mulai')}</div>`+
   `<div class="row">Durasi: <input type="number" data-s="andur" step="0.25" min="0.1" max="60" value="0.5"> dtk</div><div class="row">Tunda: <input type="number" data-s="andelay" step="0.25" min="0" max="60" value="0"> dtk</div></div>`+
   `<div class="rows tmw"><div class="row"><b>Urutkan Ulang Animasi</b></div>${B('▲','Pindah Lebih Awal','anup','rw')}${B('▼','Pindah Lebih Lambat','andown','rw')}</div>`);

/* ---- tab TRANSISI ---- */
const TRS=[['none','Tidak Ada'],['cut','Potong'],['fade','Pudar'],['push','Dorong'],['wipe','Sapu'],['split','Belah'],['reveal','Ungkap'],['bars','Batang Acak'],['shape','Bentuk'],['uncover','Singkap'],['cover','Tutupi'],['flash','Kilat'],['fall','Jatuh'],['zoom','Zoom']];
const DIRS=[['l','Dari Kanan'],['r','Dari Kiri'],['u','Dari Bawah'],['d','Dari Atas']],OUTS=[['l','Ke Kiri'],['r','Ke Kanan'],['u','Ke Atas'],['d','Ke Bawah']];
const TROPT={fade:[['s','Mulus'],['k','Melalui Hitam']],push:DIRS,wipe:DIRS,cover:DIRS,uncover:OUTS,reveal:OUTS,
  split:[['vo','Vertikal Keluar'],['ho','Horizontal Keluar'],['vi','Vertikal Masuk'],['hi','Horizontal Masuk']],bars:[['v','Vertikal'],['h','Horizontal']],
  shape:[['c','Lingkaran'],['d','Belah Ketupat']],fall:[['l','Kiri'],['r','Kanan']],zoom:[['i','Masuk'],['o','Keluar']]};
const getTr=s=>Object.assign({t:'none',o:'',dur:1,snd:'',click:1,afterOn:0,after:0},s&&s.tr||{});
const trIcon=t=>{const f=(x,y,w,h,o)=>`<rect x="${x}" y="${y}" width="${w}" height="${h}" fill="#3f6fb0"${o?` opacity="${o}"`:''}/>`;
  const m={none:'',cut:f(8,6,22,16),fade:f(1.5,1.5,33,25,.5),push:f(1.5,1.5,22,25)+'<path d="M26 9l7 5-7 5z" fill="#3f6fb0"/>',wipe:f(1.5,1.5,18,25)+'<path d="M24 4v20" stroke="#3f6fb0" stroke-width="3"/>',
    split:f(1.5,1.5,14,25)+f(20,1.5,14.5,25),reveal:f(1.5,1.5,33,25,.3)+f(1.5,1.5,14,25),bars:[0,1,2,3,4].map(i=>f(3+i*6.5,1.5,3.5,25)).join(''),
    shape:'<circle cx="18" cy="14" r="9" fill="#3f6fb0"/>',uncover:f(14,1.5,20.5,25)+f(1.5,1.5,12,25,.4),cover:f(10,6,25,20),
    flash:'<circle cx="18" cy="14" r="8" fill="#fff"/><circle cx="18" cy="14" r="3.5" fill="#3f6fb0"/>',fall:'<path d="M3 26V8l25 10z" fill="#3f6fb0"/>',zoom:f(11,8,14,11)};
  return `<svg width="38" height="30" viewBox="0 0 36 28"><rect x="1.5" y="1.5" width="33" height="25" fill="#cfe0f5" stroke="#4a73b5"/>${m[t]||''}</svg>`;};
function svEl(s,k,mode){const d=document.createElement('div');d.className='sv';d.style.cssText='position:absolute;inset:0;'+(s&&s.bg?'background:'+s.bg:'');
  d.innerHTML=`<div class="sc" style="transform:scale(${k})">${s?s.objs.map(o=>obHTML(o,mode)).join(''):''}</div>`;return d;}
function playSnd(n){if(!n)return;try{const A=window.AudioContext||window.webkitAudioContext,c=new A(),g=c.createGain(),o=c.createOscillator();
  const P={click:[900,.06,'square'],chime:[1320,.6,'sine'],pop:[520,.15,'triangle'],whoosh:[250,.5,'sawtooth']}[n];if(!P)return;
  o.type=P[2];o.frequency.setValueAtTime(P[0],c.currentTime);if(n==='whoosh')o.frequency.exponentialRampToValueAtTime(1400,c.currentTime+P[1]);
  g.gain.setValueAtTime(.08,c.currentTime);g.gain.exponentialRampToValueAtTime(.0001,c.currentTime+P[1]);o.connect(g);g.connect(c.destination);o.start();o.stop(c.currentTime+P[1]);setTimeout(()=>c.close(),P[1]*1000+200);}catch{}}
/* jalankan transisi: oldEl (slide lama) -> newEl (slide baru) di dalam box; mengembalikan fungsi untuk menyelesaikan seketika */
function runTr(box,oldEl,newEl,tr,done){
  const dur=Math.max(150,(tr.dur||1)*1000),o=tr.o||(TROPT[tr.t]||[['']])[0][0],anims=[],extra=[];
  const an=(el,kf,opt)=>{anims.push(el.animate(kf,Object.assign({duration:dur,easing:'cubic-bezier(.4,0,.2,1)',fill:'both'},opt||{})));};
  const mv={l:['translateX(100%)','translateX(-100%)'],r:['translateX(-100%)','translateX(100%)'],u:['translateY(100%)','translateY(-100%)'],d:['translateY(-100%)','translateY(100%)']}[o]||['none','none'];
  const wc={l:'inset(0 0 0 100%)',r:'inset(0 100% 0 0)',u:'inset(100% 0 0 0)',d:'inset(0 0 100% 0)'}[o]||'inset(0 0 0 100%)';
  const clones=(specs)=>{newEl.style.visibility='hidden';specs.forEach(([clip,from,delay,d2])=>{const c=newEl.cloneNode(true);c.style.visibility='visible';c.style.clipPath=clip;box.appendChild(c);extra.push(c);
    an(c,from?[{transform:from},{transform:'none'}]:[{opacity:0},{opacity:1}],{delay:delay||0,duration:d2||dur});});};
  const flat=e=>e;
  switch(tr.t){
    case'fade':if(o==='k'){box.style.background='#000';an(oldEl,[{opacity:1,offset:0},{opacity:0,offset:.5},{opacity:0,offset:1}],{easing:'linear'});an(newEl,[{opacity:0,offset:0},{opacity:0,offset:.5},{opacity:1,offset:1}],{easing:'linear'});}
      else an(newEl,[{opacity:0},{opacity:1}]);break;
    case'push':an(newEl,[{transform:mv[0]},{transform:'none'}]);an(oldEl,[{transform:'none'},{transform:mv[1]}]);break;
    case'wipe':an(newEl,[{clipPath:wc},{clipPath:'inset(0 0 0 0)'}]);break;
    case'cover':an(newEl,[{transform:mv[0]},{transform:'none'}]);break;
    case'uncover':oldEl.style.zIndex=2;an(oldEl,[{transform:'none'},{transform:mv[1]}]);break;
    case'reveal':oldEl.style.zIndex=2;an(oldEl,[{transform:'none',opacity:1},{transform:mv[1],opacity:.2}]);an(newEl,[{transform:'scale(.92)'},{transform:'none'}]);break;
    case'split':
      if(o==='vo')an(newEl,[{clipPath:'inset(0 50% 0 50%)'},{clipPath:'inset(0 0 0 0)'}]);
      else if(o==='ho')an(newEl,[{clipPath:'inset(50% 0 50% 0)'},{clipPath:'inset(0 0 0 0)'}]);
      else if(o==='vi')clones([['inset(0 50% 0 0)','translateX(-100%)'],['inset(0 0 0 50%)','translateX(100%)']]);
      else clones([['inset(0 0 50% 0)','translateY(-100%)'],['inset(50% 0 0 0)','translateY(100%)']]);break;
    case'bars':{const n=16,sp=[];for(let i=0;i<n;i++){const a=i*100/n,b=100-(i+1)*100/n;sp.push([o==='h'?`inset(${a}% 0 ${b}% 0)`:`inset(0 ${b}% 0 ${a}%)`,null,Math.random()*dur*.6,dur*.4]);}clones(sp);break;}
    case'shape':if(o==='d')an(newEl,[{clipPath:'polygon(50% 50%,50% 50%,50% 50%,50% 50%)'},{clipPath:'polygon(50% -50%,150% 50%,50% 150%,-50% 50%)'}]);
      else an(newEl,[{clipPath:'circle(0% at 50% 50%)'},{clipPath:'circle(75% at 50% 50%)'}]);break;
    case'flash':{const w=document.createElement('div');w.style.cssText='position:absolute;inset:0;background:#fff;z-index:5;opacity:0';box.appendChild(w);extra.push(w);
      an(w,[{opacity:0,offset:0},{opacity:1,offset:.45},{opacity:0,offset:1}],{easing:'linear'});
      an(newEl,[{opacity:0,offset:0},{opacity:0,offset:.45},{opacity:1,offset:.46},{opacity:1,offset:1}],{easing:'linear'});break;}
    case'fall':oldEl.style.zIndex=2;oldEl.style.transformOrigin=o==='r'?'100% 100%':'0% 100%';an(oldEl,[{transform:'rotate(0deg)',opacity:1},{transform:`rotate(${o==='r'?90:-90}deg)`,opacity:0}],{easing:'cubic-bezier(.5,0,1,.6)'});break;
    case'zoom':if(o==='o'){oldEl.style.zIndex=2;an(oldEl,[{transform:'none',opacity:1},{transform:'scale(1.6)',opacity:0}]);}else an(newEl,[{transform:'scale(.2)',opacity:0},{transform:'none',opacity:1}]);break;
  }
  const fin=()=>{clearTimeout(tm);anims.forEach(a=>{try{a.cancel();}catch{}});extra.forEach(e=>e.remove());if(oldEl)oldEl.remove();newEl.style.visibility='';box.style.background='';const d=done;done=null;if(d)d();};
  const tm=setTimeout(fin,dur+60);
  return fin;
}
const RIB_TRANS=
 G('Pratinjau',IB('▶','Pratinjau','trprev',0,'Pratinjau'))+
 G('Transisi ke Slide Ini',`<div class="tgal trg">${TRS.map(([id,n])=>`<button class="trb" data-tr="${id}" data-a="trset:${id}" title="${n}">${trIcon(id)}<span>${n}</span></button>`).join('')}</div>`+
   `<div class="rows"><button class="b" data-a="trscroll:-1" title="Sebelumnya">◀</button><button class="b" data-a="trscroll:1" title="Berikutnya">▶</button><button class="b" data-m="trmenu" title="Semua transisi">▾</button></div>`+
   `<button class="b big treff" title="Opsi Efek" data-m="treff"><em>⚙</em><span>Opsi Efek ▾</span></button>`)+
 G('Pengaturan Waktu',`<div class="rows tmw"><div class="row">Suara: ${S('trsnd',[['','Tanpa Suara'],['click','Klik'],['chime','Denting'],['whoosh','Whoosh'],['pop','Pop']],112,'Suara')}</div>`+
   `<div class="row">Durasi: <input type="number" data-s="trdur" step="0.25" min="0.25" max="60" value="1" title="Durasi (detik)"> dtk</div>`+
   `<div class="row"><button class="b" data-a="trall" title="Terapkan transisi ke semua slide">⧉ Terapkan ke Semua</button></div></div>`+
   `<div class="rows tmw"><div class="row"><b>Maju Slide</b></div><div class="row"><label><input type="checkbox" data-s="trclk" checked> Saat Klik Mouse</label></div>`+
   `<div class="row"><label><input type="checkbox" data-s="traft"> Setelah:</label><input type="number" data-s="traftv" step="1" min="0" value="0" title="Detik"> dtk</div></div>`);

/* ---- tab DESAIN: tema, varian, latar belakang ---- */
const THEMES=[
 {n:'Office',hf:'Calibri Light',bf:'Calibri',acc:['#5b9bd5','#ed7d31','#a5a5a5','#ffc000','#4472c4','#70ad47'],d:'plain'},
 {n:'Pita Biru',hf:'Trebuchet MS',bf:'Calibri',acc:['#2e75b6','#00b0f0','#7f7f7f','#ffc000','#4472c4','#70ad47'],d:'bar'},
 {n:'Hijau Segar',hf:'Georgia',bf:'Calibri',acc:['#70ad47','#a9d18e','#ed7d31','#ffc000','#4472c4','#a5a5a5'],d:'side'},
 {n:'Senja',hf:'Trebuchet MS',bf:'Calibri',acc:['#ed7d31','#c55a11','#ffc000','#a5a5a5','#5b9bd5','#70ad47'],d:'band'},
 {n:'Samudra',hf:'Segoe UI',bf:'Segoe UI',acc:['#1a8a9a','#2eb8c9','#0b5563','#ffc000','#ed7d31','#a5a5a5'],d:'frame'},
 {n:'Anggur',hf:'Cambria',bf:'Calibri',acc:['#7030a0','#b565d8','#d9a3f0','#ffc000','#ed7d31','#a5a5a5'],d:'corner'},
 {n:'Hutan',hf:'Georgia',bf:'Cambria',acc:['#4e7a3e','#8fb573','#c9a227','#a5a5a5','#5b9bd5','#ed7d31'],d:'side'},
 {n:'Grafit',hf:'Arial',bf:'Arial',acc:['#595959','#7f7f7f','#d24726','#ffc000','#5b9bd5','#70ad47'],d:'bar'},
 {n:'Marun',hf:'Times New Roman',bf:'Calibri',acc:['#a6192e','#d9534f','#f0ad4e','#a5a5a5','#5b9bd5','#70ad47'],d:'band'},
 {n:'Langit',hf:'Verdana',bf:'Verdana',acc:['#00a0d6','#5ad0f5','#1b3a6b','#ffc000','#ed7d31','#a5a5a5'],d:'corner'}
];
/* dekorasi: [jenis(r=isi,o=elips,f=bingkai),x,y,w,h] sebagai pecahan slide */
const DECO={plain:[],bar:[['r',0,0,1,.045]],side:[['r',0,0,.03,1]],band:[['r',0,.91,1,.09]],frame:[['f',.02,.035,.96,.93]],corner:[['o',.94,.05,.045,.08],['r',0,.95,1,.05]]};
const mixC=(a,b,k)=>{const q=h=>[1,3,5].map(i=>parseInt(h.slice(i,i+2),16)),x=q(a),y=q(b);return '#'+x.map((v,i)=>Math.round(v+(y[i]-v)*k).toString(16).padStart(2,'0')).join('');};
function pal(t,v){const a=t.acc[0];
  return [{bg:'#ffffff',title:mixC(a,'#000000',.35),body:'#404040',mute:'#898989',deco:a},
          {bg:mixC(a,'#ffffff',.9),title:mixC(a,'#000000',.45),body:'#333333',mute:'#7f7f7f',deco:mixC(a,'#000000',.1)},
          {bg:'#1e1e1e',title:'#ffffff',body:'#e0e0e0',mute:'#a6a6a6',deco:a},
          {bg:mixC(a,'#000000',.7),title:'#ffffff',body:'#f0f0f0',mute:'#c8c8c8',deco:mixC(a,'#ffffff',.25)}][v];}
function tThumb(t,v){const c=pal(t,v);
  const d=DECO[t.d].map(([k,x,y,w,h])=>`<i class="d" style="left:${x*100}%;top:${y*100}%;width:${w*100}%;height:${h*100}%;${k==='f'?`border:2px solid ${c.deco}`:`background:${c.deco}`}${k==='o'?';border-radius:50%':''}"></i>`).join('');
  return `<span class="tb" style="background:${c.bg}">${d}<b style="color:${c.title};font-family:'${t.hf}',sans-serif">Aa</b><span class="sws">${t.acc.map(x=>`<i style="background:${x}"></i>`).join('')}</span></span>`;}
const RIB_DESIGN=
 G('Tema',`<div class="tgal">${THEMES.map((t,i)=>`<button class="tt" data-t="${i}" data-a="theme:${i}" title="${t.n}">${tThumb(t,0)}</button>`).join('')}</div>`+
   `<div class="rows"><button class="b" data-a="tscroll:-1" title="Sebelumnya">◀</button><button class="b" data-a="tscroll:1" title="Berikutnya">▶</button><button class="b" data-m="themes" title="Semua tema">▾</button></div>`)+
 G('Varian','<div class="vgal"></div>')+
 G('Sesuaikan',IB('🎨','Format Latar Belakang','bgfmt',0,'Format<br>Latar Belakang'));

/* ---- tab SLIDE SHOW ---- */
const SHI=(()=>{
  const bd='<rect x="5" y="2.5" width="24" height="16" fill="#fff" stroke="#6b6b6b"/><path d="M5 5.5h24" stroke="#6b6b6b"/><path d="M17 18.5v5M17 23.5l-6 7M17 23.5l6 7" stroke="#6b6b6b" fill="none" stroke-width="1.4"/>';
  const S=x=>`<svg width="34" height="32" viewBox="0 0 34 32">${x}</svg>`;
  return{
    start:S(bd+'<circle cx="24" cy="21" r="7.5" fill="#4caf50" stroke="#2e7d32"/><path d="M21.5 17.5l6 3.5-6 3.5z" fill="#fff"/>'),
    cur:S(bd+'<circle cx="17" cy="11" r="4.6" fill="#ed7d31"/><path d="M17 11V6.4a4.6 4.6 0 0 1 4.6 4.6z" fill="#5b9bd5"/>'),
    online:S(bd+'<circle cx="24" cy="21" r="7.5" fill="#3b8be0" stroke="#1f5fa8"/><path d="M16.5 21h15M24 13.5v15M20 15.5q4 5.5 0 11M28 15.5q-4 5.5 0 11" stroke="#fff" fill="none" stroke-width=".9"/>'),
    cust:S(bd+'<rect x="19" y="14" width="12" height="15" fill="#fff" stroke="#4a73b5"/><path d="M21.5 18h7M21.5 21.5h7M21.5 25h7" stroke="#4a73b5" stroke-width="1.5"/>'),
    setup:S('<rect x="6" y="3" width="17" height="22" fill="#fff" stroke="#6b6b6b"/><path d="M9 8h11M9 12h11M9 16h8" stroke="#9a9a9a"/><circle cx="12" cy="20" r="2.5" fill="#ed7d31"/><circle cx="23" cy="21" r="6.5" fill="#e8e8e8" stroke="#6b6b6b"/><circle cx="23" cy="21" r="2.4" fill="#fff" stroke="#6b6b6b"/><path d="M23 13.5v3M23 25.5v3M15.5 21h3M27.5 21h3" stroke="#6b6b6b" stroke-width="2"/>'),
    hide:S('<rect x="5" y="5" width="24" height="18" fill="#fff" stroke="#6b6b6b"/><circle cx="17" cy="14" r="4.6" fill="#ed7d31"/><path d="M17 14V9.4a4.6 4.6 0 0 1 4.6 4.6z" fill="#5b9bd5"/><path d="M3 27L31 1" stroke="#444" stroke-width="2.2"/>'),
    reh:S('<rect x="3" y="3" width="24" height="18" fill="#fff" stroke="#6b6b6b"/><path d="M3 6.5h24" stroke="#6b6b6b"/><circle cx="23" cy="21" r="8" fill="#fff" stroke="#44546a" stroke-width="1.6"/><path d="M23 15.5V21l4 2.5" stroke="#44546a" stroke-width="1.6" fill="none"/>'),
    rec:S('<rect x="3" y="3" width="24" height="18" fill="#fff" stroke="#6b6b6b"/><path d="M3 6.5h24" stroke="#6b6b6b"/><circle cx="12" cy="17" r="7" fill="#e53935" stroke="#a31515"/><circle cx="24" cy="21" r="7" fill="#fff" stroke="#44546a" stroke-width="1.5"/><path d="M24 16.5V21l3.4 2" stroke="#44546a" stroke-width="1.5" fill="none"/>')
  };
})();
const SB=(ic,tip,a,lab)=>`<button class="b big" title="${tip}" data-a="${a}"><em>${ic}</em><span>${lab||tip}</span></button>`;
const SM=(ic,tip,m,lab)=>`<button class="b big" title="${tip}" data-m="${m}"><em>${ic}</em><span>${lab||tip} ▾</span></button>`;
const RIB_SHOW=
 G('Mulai Slide Show',SB(SHI.start,'Mulai dari slide pertama (F5)','showfirst','Dari<br>Awal')+SB(SHI.cur,'Mulai dari slide saat ini (Shift+F5)','show','Dari Slide<br>Saat Ini')+
   SM(SHI.online,'Presentasikan Online','online','Presentasikan<br>Online')+SM(SHI.cust,'Slide Show Kustom','cshow','Slide Show<br>Kustom'))+
 G('Siapkan',SB(SHI.setup,'Siapkan Slide Show','shsetup','Siapkan<br>Slide Show')+SB(SHI.hide,'Sembunyikan slide yang dipilih','hideslide','Sembunyikan<br>Slide')+
   SB(SHI.reh,'Latih Pengaturan Waktu','rehearse','Latih Pengaturan<br>Waktu')+SM(SHI.rec,'Rekam Slide Show','rec','Rekam Slide<br>Show')+
   `<div class="rows tmw"><div class="row"><label><input type="checkbox" data-s="shnarr" checked> Putar Narasi</label></div><div class="row"><label><input type="checkbox" data-s="shtim" checked> Gunakan Pengaturan Waktu</label></div><div class="row"><label><input type="checkbox" data-s="shmed" checked> Tampilkan Kontrol Media</label></div></div>`)+
 G('Monitor',`<div class="rows tmw"><div class="row">Monitor: <select data-s="shmon" class="shsel" title="Pilih monitor untuk slide show"><option value="auto">Otomatis</option><option value="main">Monitor Utama</option></select></div><div class="row"><label><input type="checkbox" data-s="shpv"> Gunakan Tampilan Presenter</label></div></div>`);

function mount(root,opts={}){
  if(!document.getElementById('pp-css')){const s=document.createElement('style');s.id='pp-css';s.textContent=CSS;document.head.appendChild(s);}
  let name='Presentasi1.pptx',scale=.8,autoFit=true,slides=[mkSlide('title')],cur=0,sel=[],hist=[],hi=-1,clip=null,painting=null,lastR=null,lastTx=null,drawing=null,view='normal',fk=-1,timer=0,sortFrom=null;
  let ss={type:'full',loop:0,playNarr:1,noAnim:0,pen:'#ff0000',range:'all',from:1,to:0,custom:'',useTimings:1,mediaCtrl:1,monitor:'auto',presenter:0},customShows=[],cidn=1,scrList=[],scrTried=false;
  const panes=TABS.map((t,i)=>`<div class="ribbon" data-pane="${t[0]}"${i?' hidden':''}>${t[0]==='home'?RIB_HOME:t[0]==='insert'?RIB_INSERT:t[0]==='design'?RIB_DESIGN:t[0]==='trans'?RIB_TRANS:t[0]==='anim'?RIB_ANIM:t[0]==='show'?RIB_SHOW:`<div class="soon">Menu ${t[1]} akan ditambahkan pada tahap berikutnya.</div>`}</div>`).join('');
  root.innerHTML=`<div class="pp">
    <div class="qat"><div class="l"><span class="w">P</span><button data-a="save" title="Simpan (Ctrl+S)">💾</button><button data-a="undo" title="Urungkan (Ctrl+Z)">↶</button><button data-a="redo" title="Ulangi (Ctrl+Y)">↷</button><button data-a="showfirst" title="Mulai dari awal">▶</button></div><span class="fn"></span></div>
    <div class="tabs"><button class="tab file" data-a="menu">BERKAS</button>${TABS.map((t,i)=>`<button class="tab${i?'':' on'}" data-tab="${t[0]}">${t[1]}</button>`).join('')}</div>
    ${panes}
    <div class="main"><div class="thumbs" tabindex="0"></div>
      <div class="ws"><div class="desk"><div class="holder"><div class="slide stg" tabindex="-1"></div></div></div><div class="notes" contenteditable="true" data-ph="Klik untuk menambahkan catatan"></div></div>
      <div class="sorter" hidden></div>
      <div class="apane" hidden><div class="ah"><span>Panel Animasi</span><span><button data-a="anprev" title="Putar semua">▶ Putar</button> <button data-a="anpane" title="Tutup">✕</button></span></div><div class="apl"></div></div></div>
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
    $$('.stg .ob').forEach(el=>{const o=byId(el.dataset.id),t=el.querySelector('.tx');if(o&&t)o.html=(t.textContent.trim()||t.querySelector('table'))?t.innerHTML:'';});
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
  function draw(){renumber();stg.innerHTML=objs().map(o=>obHTML(o,true)).join('');stg.style.background=S0().bg||'';
    (S0().comments||[]).forEach((c,i)=>stg.insertAdjacentHTML('beforeend',`<div class="cm" data-ci="${i}" style="left:${c.x}px;top:${c.y}px" title="${esc(c.t)}">💬</div>`));animBadges();decorate();status();syncTr();syncSs();}
  function decorate(){
    $$('.stg .ob').forEach(el=>{const on=sel.includes(el.dataset.id);el.classList.toggle('sel',on);if(!on)el.classList.remove('edt');el.querySelectorAll('.hd,.ed').forEach(x=>x.remove());});
    if(sel.length===1){const o=byId(sel[0]),el=o&&elOf(o.id);if(el){
      el.insertAdjacentHTML('beforeend',['t','b','l','r'].map(s=>`<div class="ed ${s}" data-ed="1"></div>`).join(''));
      const line=o.shape==='line'||o.shape==='arrow',hs=line?['w','e']:['nw','n','ne','e','se','s','sw','w'];
      el.insertAdjacentHTML('beforeend',hs.map(h=>{const x=h.includes('w')?0:h.includes('e')?100:50,y=h.includes('n')?0:h.includes('s')?100:50;
        return `<div class="hd" data-h="${h}" style="left:${x}%;top:${y}%;cursor:${h}-resize"></div>`;}).join(''));}}
    syncUI();syncAn();
  }
  function redrawOb(o){sync();const el=elOf(o.id);if(el){el.outerHTML=obHTML(o,true);}decorate();}
  function thumbHTML(s,i,w,c){const k=w/W;return `<div class="${c}${i===cur?' on':''}${s.hid?' hid':''}" data-i="${i}" draggable="true"><span class="n">${i+1}${s.tr&&s.tr.t&&s.tr.t!=='none'?'<br><i class="trm">★</i>':''}</span><div class="tv" style="width:${w}px;height:${Math.round(H*k)}px${s.bg?';background:'+s.bg:''}"><div class="sc" style="transform:scale(${k})">${s.objs.map(o=>obHTML(o,false)).join('')}</div></div></div>`;}
  function thumbs(){
    renumber();const tp=thumbsEl.scrollTop,secs=slides.map((s,i)=>s.sec!=null?i:-1).filter(i=>i>=0);let h='';
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
    else if(stg.contains(s.anchorNode)||s.anchorNode===stg){lastR=null;updDis();}
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
    updDis();
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
    slides.splice(cur+1,0,mkSlide(l));if(themeOn)styleSlide(slides[cur+1]);if(hfs.all)hfBuild(slides[cur+1],cur+1,hfs,new Date());cur++;sel=[];draw();notes.innerHTML='';commit();
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
    const out=smartBoxes(kind,L,o);
    sync();const i=objs().indexOf(o);objs().splice(i,1,...out);sel=out.map(x=>x.id);draw();commit();
  }
  function smartBoxes(kind,L,o){
    const n=L.length,pad=16,out=[];
    const box=(x,y,w,h,t)=>mk({k:'sh',shape:'round',x,y,w,h,fill:'#5b9bd5',stroke:'#41719c',sw:1,anchor:'m',html:'<div>'+esc(t)+'</div>',
      ts:{size:Math.max(10,Math.min(28,Math.floor(Math.min(h*.35,(w-20)/Math.max(6,t.length)*1.7)*.75))),color:'#fff',align:'center'}});
    if(kind==='blocks'){const c=Math.ceil(Math.sqrt(n)),r=Math.ceil(n/c),w=(o.w-pad*(c-1))/c,h=(o.h-pad*(r-1))/r;L.forEach((t,i)=>out.push(box(o.x+(i%c)*(w+pad),o.y+Math.floor(i/c)*(h+pad),w,h,t)));}
    else if(kind==='process'){const g=44,w=(o.w-g*(n-1))/n,h=Math.min(o.h,180),y=o.y+(o.h-h)/2;
      L.forEach((t,i)=>{out.push(box(o.x+i*(w+g),y,w,h,t));if(i<n-1)out.push(mk({k:'sh',shape:'rarrow',x:o.x+i*(w+g)+w+8,y:y+h/2-14,w:g-16,h:28,fill:'#a5a5a5',stroke:'#787878'}));});}
    else{const h=(o.h-pad*(n-1))/n;L.forEach((t,i)=>out.push(box(o.x,o.y+i*(h+pad),o.w,h,t)));}
    return out;
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
  const fmtT=s=>{s=Math.max(0,Math.round(s));const h=Math.floor(s/3600),m=Math.floor(s%3600/60),c=s%60;return(h?h+':'+String(m).padStart(2,'0'):m)+':'+String(c).padStart(2,'0');};
  const cidIdx=id=>slides.findIndex(s=>s.cid===id);
  function showOrder(from,all,cs){
    const vis=i=>!slides[i].hid,every=slides.map((_,i)=>i);let list;
    if(from!=null||all){list=every.filter(i=>vis(i)||i===from);return{list,p:Math.max(0,list.indexOf(from!=null?from:0))};}
    if(cs)list=cs.ids.map(cidIdx).filter(i=>i>=0);
    else if(ss.range==='range'){const a=Math.max(1,Math.min(slides.length,+ss.from||1)),b=Math.max(a,Math.min(slides.length,+ss.to||slides.length));list=every.slice(a-1,b);}
    else list=every;
    return{list:list.filter(vis),p:0};
  }
  /* o: {mode:'play'|'rehearse'|'record', custom, timing, narr, mic} */
  function show(from,o){
    o=o||{};sync();
    const mode=o.mode||'play',rec=mode!=='play',kiosk=!rec&&ss.type==='kiosk',win=!rec&&ss.type==='win';
    const cs=o.custom||(!rec&&from==null&&ss.range==='custom'?customShows.find(x=>x.name===ss.custom):null);
    const ord=showOrder(from,rec,cs),order=ord.list;let p=ord.p;
    if(!order.length)return toast('Tidak ada slide untuk ditampilkan');
    let i=order[p];
    const pv=!rec&&ss.presenter&&!kiosk;
    const ov=document.createElement('div');ov.className='show'+(win?' win':'')+(pv?' pv':'');ov.tabIndex=0;
    let bk=null;if(win){bk=document.createElement('div');bk.className='showbk';pp.appendChild(bk);}
    const host=document.createElement('div');host.className='shh';ov.appendChild(host);
    let side=null;
    if(pv){side=document.createElement('div');side.className='pvs';side.innerHTML='<div class="pvn"></div><div class="pvx"></div><div class="pvi"><b class="pvc"></b><span class="pvt">0:00</span></div>';ov.appendChild(side);}
    const lz=document.createElement('div');lz.className='shlz';ov.appendChild(lz);
    if(!rec&&!kiosk){const tb=document.createElement('div');tb.className='shtb';
      tb.innerHTML='<button data-k="prev" title="Sebelumnya">◀</button><button data-k="next" title="Berikutnya">▶</button><button data-k="pen" title="Pena (Ctrl+P)">✎</button><button data-k="laser" title="Penunjuk laser (Ctrl+L)">◉</button><button data-k="erase" title="Hapus coretan (E)">⌫</button><button data-k="blk" title="Layar hitam (B)">■</button><button data-k="end" title="Akhiri slide show (Esc)">✕</button>';ov.appendChild(tb);}
    let rb=null;
    if(rec){rb=document.createElement('div');rb.className='shrec';
      rb.innerHTML=`<button data-k="next" title="Slide berikutnya">➜</button><button data-k="pause" title="Jeda">⏸</button><span class="rt" title="Waktu slide ini">0:00</span><button data-k="repeat" title="Ulangi slide ini">↺</button><span class="rtt" title="Total waktu">0:00</span>${o.mic?'<b class="rdot" title="Merekam narasi">●</b>':''}<button data-k="end" title="Selesai (Esc)">✕</button>`;ov.appendChild(rb);}
    pp.appendChild(ov);

    let box=null,cancelTr=null,autoT=0,steps=[],si=0,aud=null,tool=null,ink=null,cv=null,mvT=0,closed=false;
    const noAnim=!rec&&!!ss.noAnim;

    /* waktu & narasi (latih / rekam) */
    let t0=performance.now(),acc=0,paused=false;const T={},Bl={},pend=[];let mr=null;
    const elapsed=()=>acc+(paused?0:(performance.now()-t0)/1000),resetT=()=>{acc=0;t0=performance.now();};
    const nrStart=()=>{if(!o.mic)return;try{const ch=[],idx=i,r=new MediaRecorder(o.mic);let res;const pr=new Promise(x=>res=x);
      r.ondataavailable=e=>{if(e.data&&e.data.size)ch.push(e.data);};
      r.onstop=()=>{if(!r.drop&&ch.length)Bl[idx]=new Blob(ch,{type:r.mimeType||'audio/webm'});res();};
      r.start();mr=r;r.done=pr;}catch{mr=null;}};
    const nrStop=drop=>{if(!mr)return;const r=mr;mr=null;r.drop=!!drop;pend.push(r.done);try{if(r.state!=='inactive')r.stop();}catch{}};
    const leave=()=>{if(!rec)return;T[i]=Math.max(1,Math.round(elapsed()));nrStop();resetT();};

    const playStep=()=>{const st=steps[si++];if(st&&box)anPlayItems(box.lastElementChild,st.items);};
    const arm=()=>{
      if(steps[0]&&steps[0].auto&&!si)playStep();
      const a=getTr(slides[i]);
      if(!rec&&a.afterOn&&(ss.useTimings||kiosk))autoT=setTimeout(()=>next(),(+a.after||0)*1000);
    };
    const stopAud=()=>{if(aud){try{aud.pause();}catch{}aud=null;}};
    const dropInk=()=>{if(ink){ink.remove();ink=null;}};
    const place=()=>{if(!ink||!box)return;const a=box.getBoundingClientRect(),b=ov.getBoundingClientRect(),w=Math.round(a.width),h=Math.round(a.height);
      ink.style.left=a.left-b.left+'px';ink.style.top=a.top-b.top+'px';ink.style.width=w+'px';ink.style.height=h+'px';if(ink.width!==w||ink.height!==h){ink.width=w;ink.height=h;}};
    const pvUpd=()=>{if(!side)return;const nx=order[p+1],pn=side.querySelector('.pvn');pn.innerHTML='';
      if(nx!=null){pn.appendChild(svEl(slides[nx],(pn.clientWidth||280)/W,false));}else pn.textContent='Akhir slide show';
      side.querySelector('.pvx').innerHTML=slides[i].notes||'<i style="color:#888">(Tidak ada catatan untuk slide ini)</i>';
      side.querySelector('.pvc').textContent='Slide '+(p+1)+' dari '+order.length;};
    const mediaFix=()=>{if(!ss.mediaCtrl)box.querySelectorAll('video,audio').forEach(m=>{m.controls=false;});};
    const paint=(play,quiet)=>{
      const k=Math.min(host.clientWidth/W,host.clientHeight/H);if(cancelTr){cancelTr();cancelTr=null;}clearTimeout(autoT);steps=play&&!noAnim?anSteps(slides[i]):[];si=0;
      const prev=box&&box.lastElementChild,t=getTr(slides[i]);
      if(play&&prev&&box.dataset.k==k&&t.t!=='none'&&t.t!=='cut'){const nw=svEl(slides[i],k,'show');box.appendChild(nw);if(!noAnim)anInit(nw,slides[i]);playSnd(t.snd);cancelTr=runTr(box,prev,nw,t,()=>{cancelTr=null;arm();});}
      else{host.innerHTML='';box=document.createElement('div');box.dataset.k=k;box.style.cssText=`position:relative;overflow:hidden;width:${W*k}px;height:${H*k}px`;{const nw=svEl(slides[i],k,'show');box.appendChild(nw);if(play&&!noAnim)anInit(nw,slides[i]);}host.appendChild(box);if(play)playSnd(t.snd);arm();}
      mediaFix();
      if(!quiet){stopAud();dropInk();if(tool==='pen')mkInk();if(rec)nrStart();else if(ss.playNarr&&slides[i].narr){aud=new Audio(slides[i].narr.src);aud.play().catch(()=>{});}}
      place();pvUpd();
    };
    const finish=()=>{
      if(closed)return;closed=true;clearTimeout(autoT);clearTimeout(mvT);clearInterval(tick);if(cancelTr)cancelTr();stopAud();if(rec)leave();
      document.removeEventListener('fullscreenchange',fc);removeEventListener('resize',rs);
      ov.remove();if(bk)bk.remove();if(document.fullscreenElement)document.exitFullscreen().catch(()=>{});go(i);
      if(rec)Promise.all(pend).then(()=>{if(o.mic)o.mic.getTracks().forEach(t=>t.stop());finishRec(T,Bl,o);});
    };
    const close=finish;
    const fc=()=>{if(!document.fullscreenElement)close();};
    const rs=()=>{if(!closed)paint(0,1);};
    const cvOff=()=>{if(cv){cv.remove();cv=null;}};
    const blank=c=>{if(cv&&cv.dataset.c===c){cvOff();return;}cvOff();cv=document.createElement('div');cv.className='shcv';cv.dataset.c=c;cv.style.background=c;ov.appendChild(cv);};
    const goTo=n=>{leave();p=n;i=order[p];paint(1);};
    const next=()=>{
      if(cv){cvOff();return;}
      if(si<steps.length){playStep();return;}
      if(p<order.length-1)goTo(p+1);
      else if(!rec&&ss.loop)goTo(0);
      else close();
    };
    const prev=()=>{if(cv){cvOff();return;}if(p>0){leave();p--;i=order[p];paint();}};
    const mkInk=()=>{if(ink)return;ink=document.createElement('canvas');ink.className='shink';ov.appendChild(ink);place();
      ink.onpointerdown=e=>{if(tool!=='pen')return;e.preventDefault();ink.setPointerCapture(e.pointerId);const c=ink.getContext('2d'),r=ink.getBoundingClientRect();
        c.strokeStyle=ss.pen;c.lineWidth=Math.max(3,ink.width/280);c.lineCap=c.lineJoin='round';c.beginPath();c.moveTo(e.clientX-r.left,e.clientY-r.top);c.lineTo(e.clientX-r.left+.1,e.clientY-r.top);c.stroke();
        ink.onpointermove=ev=>{c.lineTo(ev.clientX-r.left,ev.clientY-r.top);c.stroke();c.beginPath();c.moveTo(ev.clientX-r.left,ev.clientY-r.top);};
        ink.onpointerup=ink.onpointercancel=()=>{ink.onpointermove=null;};};};
    const setTool=t=>{tool=tool===t?null:t;ov.classList.toggle('tpen',tool==='pen');ov.classList.toggle('tlaser',tool==='laser');
      ov.querySelectorAll('.shtb [data-k=pen]').forEach(b=>b.classList.toggle('on',tool==='pen'));ov.querySelectorAll('.shtb [data-k=laser]').forEach(b=>b.classList.toggle('on',tool==='laser'));
      if(tool==='pen')mkInk();};
    const api={set:n=>{n=Math.max(0,Math.min(slides.length-1,n));let q=order.indexOf(n);if(q<0){order.splice(p+1,0,n);q=p+1;}leave();p=q;i=n;paint();},end:close};
    const act=k=>{
      if(k==='next')next();else if(k==='prev')prev();else if(k==='end')close();
      else if(k==='pen')setTool('pen');else if(k==='laser')setTool('laser');else if(k==='erase'){dropInk();if(tool==='pen')mkInk();}else if(k==='blk')blank('#000');
      else if(k==='pause'){paused=!paused;if(paused)acc+=(performance.now()-t0)/1000;else t0=performance.now();
        const b=rb.querySelector('[data-k=pause]');b.textContent=paused?'▶':'⏸';b.title=paused?'Lanjutkan':'Jeda';if(mr){try{paused?mr.pause():mr.resume();}catch{}}}
      else if(k==='repeat'){nrStop(true);resetT();paused=false;const b=rb.querySelector('[data-k=pause]');b.textContent='⏸';box=null;paint(1);}
    };
    ov.onclick=e=>{
      const kb=e.target.closest&&e.target.closest('[data-k]');if(kb){act(kb.dataset.k);return;}
      const a=e.target.closest&&e.target.closest('a[href]');if(a){e.preventDefault();if(/^https?:/i.test(a.href))open(a.href,'_blank','noopener');return;}
      if(e.target.closest&&e.target.closest('.med video,.med audio,.med iframe')){const m=e.target.closest('video,audio');if(m&&!ss.mediaCtrl){if(m.paused)m.play();else m.pause();}return;}
      if(tool==='pen')return;
      const el=e.target.closest&&e.target.closest('.ob[data-lk]'),ob=el&&slides[i].objs.find(x=>x.id===el.dataset.id);if(ob&&runLink(ob.link,i,api))return;
      if(rec){next();return;}if(kiosk)return;if(getTr(slides[i]).click)next();};
    ov.onmousemove=e=>{ov.classList.add('mv');clearTimeout(mvT);mvT=setTimeout(()=>ov.classList.remove('mv'),2500);
      if(tool==='laser'){const r=ov.getBoundingClientRect();lz.style.left=e.clientX-r.left+'px';lz.style.top=e.clientY-r.top+'px';}};
    ov.onkeydown=e=>{
      const k=e.key,kl=k.toLowerCase();
      if(k==='Escape'){if(tool)setTool(tool);else if(cv)cvOff();else close();return;}
      if(kiosk)return;
      if(!rec){
        if(e.ctrlKey&&kl==='p'){e.preventDefault();setTool('pen');return;}
        if(e.ctrlKey&&kl==='l'){e.preventDefault();setTool('laser');return;}
        if(e.ctrlKey&&kl==='a'){e.preventDefault();if(tool)setTool(tool);return;}
        if(!e.ctrlKey&&!e.metaKey&&!e.altKey){
          if(kl==='e'){dropInk();if(tool==='pen')mkInk();return;}
          if(kl==='b'||k==='.'){blank('#000');return;}
          if(kl==='w'||k===','){blank('#fff');return;}
          if(k==='Home'){e.preventDefault();goTo(0);return;}
          if(k==='End'){e.preventDefault();goTo(order.length-1);return;}}}
      if(['ArrowRight','ArrowDown','PageDown',' ','Enter'].includes(k)){e.preventDefault();next();}
      else if(['ArrowLeft','ArrowUp','PageUp','Backspace'].includes(k)){e.preventDefault();prev();}};
    const t00=Date.now();
    const tick=setInterval(()=>{
      if(rb){rb.querySelector('.rt').textContent=fmtT(elapsed());rb.querySelector('.rtt').textContent=fmtT(Object.keys(T).reduce((s,x)=>s+(+x===i?0:T[x]),0)+elapsed());}
      if(side)side.querySelector('.pvt').textContent=fmtT((Date.now()-t00)/1000);},250);
    paint(1);ov.focus();addEventListener('resize',rs);
    if(!win&&ov.requestFullscreen){
      const sc=ss.monitor!=='auto'&&ss.monitor!=='main'?scrList[+ss.monitor]:null;
      Promise.resolve().then(()=>ov.requestFullscreen(sc?{screen:sc}:undefined)).then(()=>document.addEventListener('fullscreenchange',fc)).catch(()=>{});}
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
        if(s.bg)n.bg=s.bg;if(s.tr)n.tr=s.tr;if(s.hid)n.hid=1;if(s.sec!=null)n.sec=s.sec;return n;});
      themeOn=false;customShows=[];cur=0;sel=[];lastR=null;name=file.name;setName();hist=[];hi=-1;
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
      document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),1000);const nm=slides.reduce((c,s)=>c+s.objs.filter(o=>o.media).length,0),nx=slides.reduce((c,s)=>c+s.objs.filter(o=>o.img&&!/^data:/.test(o.img)).length,0),na=slides.reduce((c,s)=>c+s.objs.filter(o=>o.an&&o.an.t&&o.an.t!=='none').length,0);
      toast('Tersimpan: '+a.download+(nm?' — '+nm+' video/audio tidak ikut tersimpan':'')+(nx?' — '+nx+' gambar online tidak ikut tersimpan':'')+(na?' — '+na+' animasi tidak ikut tersimpan':'')+(slides.some(x=>x.narr)?' — narasi tidak ikut tersimpan':''),nm||nx||na||slides.some(x=>x.narr)?7000:3000);
    }catch(e){toast('Gagal menyimpan: '+e.message,6000);}
  }
  fin.addEventListener('change',()=>{const f=fin.files[0];fin.value='';if(f)openFile(f);});

  /* ===== aksi ===== */
  const acts={
    menu(){fm.hidden=!fm.hidden;},
    new(){hfs.all=0;customShows=[];slides=[mkSlide('title')];cur=0;sel=[];name='Presentasi1.pptx';setName();hist=[];hi=-1;draw();thumbs();notes.innerHTML='';commit();},
    open(){fin.click();},
    save:savePptx,
    print:printAll,home(){opts.onHome&&opts.onHome();},undo,redo,
    showfirst(){show();},show(){show(cur);},
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

  /* ===== tab SISIPKAN ===== */
  let drawVert=false,hfs={date:0,auto:1,fmt:0,dtxt:'',num:0,foot:0,ftxt:'',notitle:0,all:0},recorder=null;
  const DF=[{weekday:'long',day:'numeric',month:'long',year:'numeric'},{day:'2-digit',month:'2-digit',year:'numeric'},{day:'numeric',month:'long',year:'numeric'},
    {day:'numeric',month:'short',year:'numeric'},'iso',{day:'numeric',month:'long',year:'numeric',hour:'2-digit',minute:'2-digit'},{hour:'2-digit',minute:'2-digit'}];
  const p2=n=>String(n).padStart(2,'0');
  const fmtDate=(f,d)=>DF[f]==='iso'?`${d.getFullYear()}-${p2(d.getMonth()+1)}-${p2(d.getDate())}`:d.toLocaleString('id-ID',DF[f]);
  /* nomor slide & tanggal otomatis ikut diperbarui */
  function renumber(){const now=new Date();slides.forEach((s,i)=>s.objs.forEach(o=>{
    if(o.hf==='num')o.html='<div>'+(i+1)+'</div>';else if(o.auto)o.html='<div>'+esc(fmtDate(o.fmt||0,now))+'</div>';}));}
  function updDis(){$$('[data-need]').forEach(b=>{const n=b.dataset.need;b.disabled=!(n==='text'?hasRange():n==='link'?(selectedText()||sel.length>0):sel.length>0);});}

  /* dialog sederhana */
  function dlg(title,body,buttons){
    const bg=document.createElement('div');bg.className='dlgbg';
    bg.innerHTML=`<div class="dlg" role="dialog" aria-label="${esc(title)}"><h3>${title}</h3>${body}<div class="bt"></div></div>`;
    const close=()=>bg.remove();
    buttons.forEach(([l,fn,pri])=>{const b=document.createElement('button');b.textContent=l;if(pri)b.className='pri';b.onclick=()=>{if(!fn||fn(bg)!==false)close();};bg.querySelector('.bt').appendChild(b);});
    bg.addEventListener('mousedown',e=>{if(e.target===bg)close();if(e.target.closest('button'))e.preventDefault();});
    bg.addEventListener('keydown',e=>{if(e.key==='Escape'){e.stopPropagation();close();}});
    pp.appendChild(bg);const f=bg.querySelector('input[type=text],select,button');if(f)f.focus({preventScroll:true});return bg;
  }
  const pick=(accept,multi)=>new Promise(ok=>{const i=document.createElement('input');i.type='file';i.accept=accept;i.multiple=!!multi;i.hidden=true;pp.appendChild(i);
    i.onchange=()=>{const f=[...i.files];i.remove();ok(f);};i.addEventListener('cancel',()=>{i.remove();ok([]);});i.click();});
  const readFile=(f,as)=>new Promise((ok,err)=>{const r=new FileReader();r.onload=()=>ok(r.result);r.onerror=()=>err(r.error);as==='text'?r.readAsText(f):r.readAsDataURL(f);});
  const loadImg=src=>new Promise((ok,err)=>{const i=new Image();i.onload=()=>ok(i);i.onerror=()=>err(new Error('Gambar tidak dapat dimuat'));i.src=src;});
  async function shrink(src){                       // perkecil gambar >1600px agar ringan
    const im=await loadImg(src),w=im.naturalWidth||400,h=im.naturalHeight||300,k=Math.min(1,1600/Math.max(w,h));
    if(k===1||/^data:image\/(svg|gif)/.test(src))return{src,w,h};
    const c=document.createElement('canvas');c.width=Math.round(w*k);c.height=Math.round(h*k);c.getContext('2d').drawImage(im,0,0,c.width,c.height);
    return{src:c.toDataURL(/^data:image\/png/.test(src)?'image/png':'image/jpeg',.9),w,h};
  }
  const picObj=(r,mw,mh)=>{const k=Math.min(1,mw/r.w,mh/r.h),w=Math.max(8,Math.round(r.w*k)),h=Math.max(8,Math.round(r.h*k));
    return mk({k:'sh',shape:'rect',img:r.src,x:Math.round((W-w)/2),y:Math.round((H-h)/2),w,h,html:''});};
  const addPic=r=>addObj(picObj(r,900,520),false);
  const addMedia=m=>addObj(mk({k:'tb',x:m.type==='audio'?490:240,y:m.type==='audio'?320:130,w:m.type==='audio'?300:800,h:m.type==='audio'?70:450,media:m,html:''}),false);
  function insText(t,props){
    if(hasRange())target(()=>document.execCommand('insertText',false,t));
    else addObj(mk(Object.assign({k:'tb',x:340,y:300,w:600,h:70,html:'<div>'+esc(t)+'</div>'},props)),false);
  }

  /* tabel */
  function tableHTML(r,c,data){
    let h='<table>';
    for(let i=0;i<r;i++){h+='<tr>';for(let j=0;j<c;j++){const hd=i===0,bg=hd?'#5b9bd5':(i%2?'#cfd5ea':'#e9ebf5');
      h+=`<td style="background:${bg};color:${hd?'#ffffff':'#000000'};${hd?'font-weight:700;':''}">${data?esc((data[i]&&data[i][j])||''):'<br>'}</td>`;}h+='</tr>';}
    return h+'</table>';
  }
  function insertTable(r,c,data){
    r=Math.max(1,Math.min(12,r|0));c=Math.max(1,Math.min(10,c|0));const w=Math.min(1104,c*200),h=r*46;
    const o=mk({k:'tb',x:Math.round((W-w)/2),y:Math.max(30,Math.round((H-h)/2)),w,h,html:tableHTML(r,c,data),ts:{size:18}});
    addObj(o,false);
    const td=elOf(o.id)&&elOf(o.id).querySelector('td');
    if(td&&!data){const t=td.closest('.tx');t.focus({preventScroll:true});const rg=document.createRange();rg.selectNodeContents(td);rg.collapse(true);const g=getSelection();g.removeAllRanges();g.addRange(rg);}
  }
  function tablePicker(btn){
    closePop();const m=document.createElement('div');m.className='pm tbg';let g='';
    for(let i=0;i<80;i++)g+=`<i data-r="${Math.floor(i/10)+1}" data-c="${i%10+1}"></i>`;
    m.innerHTML=`<div class="lb">Sisipkan Tabel</div><div class="gr">${g}</div><hr><button data-tb="dlg">▦  Sisipkan Tabel…</button>`;
    pp.appendChild(m);const rr=pp.getBoundingClientRect(),br=btn.getBoundingClientRect();
    m.style.left=Math.max(0,Math.min(br.left-rr.left,rr.width-m.offsetWidth-4))+'px';m.style.top=Math.max(0,Math.min(br.bottom-rr.top,rr.height-m.offsetHeight-4))+'px';
    const cells=[...m.querySelectorAll('i')],lb=m.querySelector('.lb');
    m.addEventListener('mouseover',e=>{const i=e.target.closest('i');if(!i)return;const r=+i.dataset.r,c=+i.dataset.c;
      cells.forEach(x=>x.classList.toggle('on',+x.dataset.r<=r&&+x.dataset.c<=c));lb.textContent=c+' x '+r+' Tabel';});
    m.addEventListener('click',e=>{const i=e.target.closest('i');if(i){closePop();insertTable(+i.dataset.r,+i.dataset.c);return;}
      if(e.target.closest('[data-tb]')){closePop();tableDlg();}});
  }
  function tableDlg(){
    dlg('Sisipkan Tabel',`<label>Jumlah kolom (1–10): <input type="text" class="tc" value="5"></label><label>Jumlah baris (1–12): <input type="text" class="tr" value="4"></label>`,
      [['Batal'],['OK',bg=>{const c=parseInt(bg.querySelector('.tc').value,10),r=parseInt(bg.querySelector('.tr').value,10);
        if(!(c>=1&&r>=1)){toast('Masukkan angka yang valid');return false;}insertTable(r,c);},1]]);
  }
  const parseCSV=t=>{const rows=[];let r=[],c='',q=false;
    for(let i=0;i<t.length;i++){const ch=t[i];
      if(q){if(ch==='"'){if(t[i+1]==='"'){c+='"';i++;}else q=false;}else c+=ch;}
      else if(ch==='"')q=true;else if(ch===','||ch===';'||ch==='\t'){r.push(c);c='';}
      else if(ch==='\n'||ch==='\r'){if(ch==='\r'&&t[i+1]==='\n')i++;r.push(c);c='';rows.push(r);r=[];}else c+=ch;}
    if(c!==''||r.length){r.push(c);rows.push(r);}return rows.filter(r=>r.some(x=>x.trim()));};

  /* bagan (digambar sebagai SVG lalu dijadikan gambar PNG agar ikut tersimpan di .pptx) */
  const CHARTS={col:'Kolom',bar:'Batang',line:'Garis',pie:'Pai'};
  function chartSVG(t,d,w,h){
    const P=['#5b9bd5','#ed7d31','#a5a5a5','#ffc000','#4472c4','#70ad47','#255e91','#9e480e','#636363','#997300','#264478','#43682b'];
    const F='font-family="Calibri,Carlito,Arial,sans-serif" font-size="22" fill="#595959"';
    const nice=v=>{if(v<=0)return 1;const p=Math.pow(10,Math.floor(Math.log10(v))),f=v/p;return(f<=1?1:f<=2?2:f<=5?5:10)*p;},fmt=v=>String(Math.round(v*100)/100);
    const n=d.length,max=Math.max(...d.map(x=>x[1]));
    let s=`<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><rect width="${w}" height="${h}" fill="#fff"/>`;
    if(t==='pie'){
      const tot=d.reduce((a,x)=>a+x[1],0)||1,cx=w*.34,cy=h/2,r=Math.min(h/2-30,w*.3);let a0=-Math.PI/2;
      d.forEach((x,i)=>{const a=x[1]/tot*Math.PI*2,a1=a0+a;
        if(d.length===1)s+=`<circle cx="${cx}" cy="${cy}" r="${r}" fill="${P[0]}" stroke="#fff" stroke-width="2"/>`;
        else s+=`<path d="M${cx} ${cy}L${cx+r*Math.cos(a0)} ${cy+r*Math.sin(a0)}A${r} ${r} 0 ${a>Math.PI?1:0} 1 ${cx+r*Math.cos(a1)} ${cy+r*Math.sin(a1)}Z" fill="${P[i%P.length]}" stroke="#fff" stroke-width="2"/>`;
        if(x[1]/tot>.04){const am=a0+a/2;s+=`<text x="${cx+r*.66*Math.cos(am)}" y="${cy+r*.66*Math.sin(am)+7}" text-anchor="middle" ${F.replace('fill="#595959"','fill="#ffffff" font-weight="700"')}>${Math.round(x[1]/tot*100)}%</text>`;}
        a0=a1;});
      d.forEach((x,i)=>{const y=h/2-(n*34)/2+i*34+17;s+=`<rect x="${w*.7}" y="${y-13}" width="20" height="20" fill="${P[i%P.length]}"/><text x="${w*.7+30}" y="${y+4}" ${F}>${esc(x[0])}</text>`;});
      return s+'</svg>';
    }
    const top=nice(max),bar=t==='bar',m={l:bar?Math.min(240,30+Math.max(...d.map(x=>x[0].length))*11):70,r:40,t:30,b:bar?50:70},pw=w-m.l-m.r,ph=h-m.t-m.b;
    for(let i=0;i<=5;i++){const v=top*i/5;
      if(bar){const x=m.l+pw*i/5;s+=`<path d="M${x} ${m.t}V${m.t+ph}" stroke="#d9d9d9"/><text x="${x}" y="${m.t+ph+30}" text-anchor="middle" ${F}>${fmt(v)}</text>`;}
      else{const y=m.t+ph-ph*i/5;s+=`<path d="M${m.l} ${y}H${m.l+pw}" stroke="#d9d9d9"/><text x="${m.l-10}" y="${y+7}" text-anchor="end" ${F}>${fmt(v)}</text>`;}}
    if(t==='col'){const bw=pw/n;d.forEach((x,i)=>{const bh=ph*x[1]/top,bx=m.l+i*bw+bw*.18;
      s+=`<rect x="${bx}" y="${m.t+ph-bh}" width="${bw*.64}" height="${bh}" fill="${P[0]}"/><text x="${bx+bw*.32}" y="${m.t+ph-bh-8}" text-anchor="middle" ${F}>${fmt(x[1])}</text><text x="${m.l+i*bw+bw/2}" y="${m.t+ph+32}" text-anchor="middle" ${F}>${esc(x[0])}</text>`;});}
    else if(bar){const bh=ph/n;d.forEach((x,i)=>{const bw=pw*x[1]/top,by=m.t+i*bh+bh*.18;
      s+=`<rect x="${m.l}" y="${by}" width="${bw}" height="${bh*.64}" fill="${P[0]}"/><text x="${m.l+bw+8}" y="${by+bh*.32+7}" ${F}>${fmt(x[1])}</text><text x="${m.l-10}" y="${by+bh*.32+7}" text-anchor="end" ${F}>${esc(x[0])}</text>`;});}
    else{const st=pw/n,pts=d.map((x,i)=>[m.l+st*i+st/2,m.t+ph-ph*x[1]/top]);
      s+=`<polyline points="${pts.map(p=>p.join(',')).join(' ')}" fill="none" stroke="${P[0]}" stroke-width="5" stroke-linejoin="round"/>`;
      pts.forEach((p,i)=>{s+=`<circle cx="${p[0]}" cy="${p[1]}" r="8" fill="${P[0]}"/><text x="${p[0]}" y="${p[1]-16}" text-anchor="middle" ${F}>${fmt(d[i][1])}</text><text x="${p[0]}" y="${m.t+ph+32}" text-anchor="middle" ${F}>${esc(d[i][0])}</text>`;});}
    s+=bar?`<path d="M${m.l} ${m.t}V${m.t+ph}" stroke="#888"/>`:`<path d="M${m.l} ${m.t+ph}H${m.l+pw}" stroke="#888"/>`;
    return s+'</svg>';
  }
  async function addChart(t,d){
    const W2=900,H2=520,svg='data:image/svg+xml;charset=utf-8,'+encodeURIComponent(chartSVG(t,d,W2,H2));
    try{const im=await loadImg(svg),c=document.createElement('canvas');c.width=W2*2;c.height=H2*2;const g=c.getContext('2d');g.fillStyle='#fff';g.fillRect(0,0,c.width,c.height);g.drawImage(im,0,0,c.width,c.height);
      addObj(picObj({src:c.toDataURL('image/png'),w:W2,h:H2},W2,H2),false);}
    catch(e){toast('Gagal membuat bagan: '+e.message,5000);}
  }
  function chartDlg(t){
    dlg('Sisipkan Bagan '+CHARTS[t],`<div>Data (format Label:Nilai, pisahkan dengan koma)</div><label><input type="text" class="cd" value="Kategori 1:4.3, Kategori 2:2.5, Kategori 3:3.5, Kategori 4:4.5"></label>`,
      [['Batal'],['OK',bg=>{
        const d=bg.querySelector('.cd').value.split(',').map(x=>{const i=x.lastIndexOf(':');return i<0?null:[x.slice(0,i).trim(),parseFloat(x.slice(i+1))];}).filter(x=>x&&x[0]&&isFinite(x[1])&&x[1]>=0).slice(0,12);
        if(!d.length){toast('Format data salah. Contoh: Jan:10, Feb:20, Mar:15 (nilai tidak boleh negatif)',5000);return false;}
        addChart(t,d);},1]]);
  }

  /* WordArt, persamaan, simbol */
  const WA=['color:#000000','color:#2f5597;text-shadow:3px 3px 3px rgba(0,0,0,.35)','color:#ed7d31;-webkit-text-stroke:1px #843c0c','color:#ffffff;-webkit-text-stroke:2px #2f5597',
    'color:#70ad47;text-shadow:2px 2px 0 #385723','color:#c00000;font-style:italic;text-shadow:0 0 8px rgba(192,0,0,.5)','color:#7030a0;-webkit-text-stroke:1px #3b1857','color:#ffc000;text-shadow:2px 2px 0 #7f6000'];
  const EQ=[['Luas Lingkaran','A = πr²'],['Teorema Pythagoras','a² + b² = c²'],['Rumus Kuadrat','x = (−b ± √(b² − 4ac)) / 2a'],['Teorema Binomial','(x + a)ⁿ = Σ [n! / (k!(n−k)!)] xᵏ aⁿ⁻ᵏ'],
    ['Identitas Euler','e^(iπ) + 1 = 0'],['Identitas Trigonometri','sin²θ + cos²θ = 1'],['Luas Segitiga','L = ½ × a × t'],['Volume Bola','V = (4/3)πr³']];
  const SYM={'Umum':'©®™§¶†‡•…‰′″°±×÷¼½¾µ¿¡«»‹›–—‘’“”','Panah':'←↑→↓↔↕↖↗↘↙⇐⇑⇒⇓⇔⇕↩↪↺↻➔➜➤','Matematika':'±×÷≠≈≤≥∞∑∏√∫∂∆∇∈∉∪∩⊂⊃⊆⊇∀∃∅∧∨¬≡∝∠⊥∥',
    'Huruf Yunani':'αβγδεζηθικλμνξοπρστυφχψωΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ','Mata Uang':'€£¥¢₹₩₽₺฿₫$¤₿','Bentuk':'★☆♠♣♥♦♪♫☀☁☂☎☑☒✓✔✗✘✚❤◆◇○●□■△▲▽▼'};
  function symbolDlg(){
    const cats=Object.keys(SYM),grid=c=>[...SYM[c]].map(ch=>`<button type="button" title="${ch}">${ch}</button>`).join('');
    const bg=dlg('Simbol',`<label>Kategori: <select class="sc">${cats.map(c=>`<option>${c}</option>`).join('')}</select></label><div class="sy">${grid(cats[0])}</div>`,[['Tutup',null,1]]);
    const sy=bg.querySelector('.sy');
    bg.querySelector('.sc').onchange=e=>{sy.innerHTML=grid(e.target.value);};
    sy.addEventListener('click',e=>{const b=e.target.closest('button');if(b)insText(b.textContent);});
  }

  /* header & footer, tanggal */
  function hfBuild(s,idx,h,now){
    s.objs=s.objs.filter(o=>!o.hf);
    if(h.notitle&&s.layout==='title')return;
    const f=(kind,x,w,al,html,extra)=>mk(Object.assign({k:'tb',hf:kind,x,y:664,w,h:40,anchor:'m',html:'<div>'+html+'</div>',ts:{size:12,color:'#8b8b8b',align:al}},extra));
    if(h.date)s.objs.push(f('date',88,288,'left',esc(h.auto?fmtDate(h.fmt,now):h.dtxt),h.auto?{auto:1,fmt:h.fmt}:{}));
    if(h.foot)s.objs.push(f('foot',436,408,'center',esc(h.ftxt)));
    if(h.num)s.objs.push(f('num',904,288,'right',String(idx+1)));
  }
  function hfDialog(pre){
    const now=new Date(),h=Object.assign({},hfs,pre||{}),ck=v=>v?' checked':'';
    dlg('Header dan Footer',
      `<label><input type="checkbox" class="hd"${ck(h.date)}> Tanggal dan waktu</label>
       <div style="margin-left:24px"><label><input type="radio" name="hdm" value="auto"${ck(h.auto)}> Perbarui otomatis</label>
       <label style="margin-left:24px"><select class="hdf">${DF.map((f,i)=>`<option value="${i}"${h.fmt===i?' selected':''}>${esc(fmtDate(i,now))}</option>`).join('')}</select></label>
       <label><input type="radio" name="hdm" value="fix"${ck(!h.auto)}> Tetap</label><label style="margin-left:24px"><input type="text" class="hdt" value="${esc(h.dtxt)}"></label></div>
       <label><input type="checkbox" class="hn"${ck(h.num)}> Nomor slide</label>
       <label><input type="checkbox" class="hf"${ck(h.foot)}> Footer</label><label style="margin-left:24px"><input type="text" class="hft" value="${esc(h.ftxt)}"></label>
       <label><input type="checkbox" class="hs"${ck(h.notitle)}> Jangan tampilkan pada slide judul</label>`,
      [['Batal'],['Terapkan',bg=>hfApply(bg,false)],['Terapkan ke Semua',bg=>hfApply(bg,true),1]]);
  }
  function hfApply(bg,all){
    const q=s=>bg.querySelector(s);
    hfs={date:q('.hd').checked?1:0,auto:q('[name=hdm]:checked').value==='auto'?1:0,fmt:+q('.hdf').value,dtxt:q('.hdt').value,num:q('.hn').checked?1:0,foot:q('.hf').checked?1:0,ftxt:q('.hft').value,notitle:q('.hs').checked?1:0,all:all?1:0};
    sync();const now=new Date();(all?slides:[S0()]).forEach(s=>hfBuild(s,slides.indexOf(s),hfs,now));draw();commit();
    toast(all?'Header dan footer diterapkan ke semua slide':'Header dan footer diterapkan ke slide ini');
  }

  /* tindakan objek saat slide show */
  function runLink(l,i,api){
    if(!l)return false;
    const a=l.a;
    if(a==='url'&&/^https?:\/\//i.test(l.url||'')){open(l.url,'_blank','noopener');return true;}
    if(a==='next'){api.set(i+1);return true;}if(a==='prev'){api.set(i-1);return true;}
    if(a==='first'){api.set(0);return true;}if(a==='last'){api.set(slides.length-1);return true;}
    if(a==='n'){api.set((+l.n||1)-1);return true;}if(a==='end'){api.end();return true;}
    return false;
  }
  const normUrl=u=>/^[a-z][a-z0-9+.-]*:/i.test(u)?u:'https://'+u;

  Object.assign(acts,{
    table:tablePicker,
    async pic(){const fs=await pick('image/*',true);if(!fs.length)return;
      try{for(const f of fs)addPic(await shrink(await readFile(f)));}catch(e){toast('Gagal menyisipkan gambar: '+e.message,5000);}},
    async onlinepic(){
      const u=(prompt('Alamat gambar (URL):','https://')||'').trim();if(!/^https?:\/\/./i.test(u))return;
      toast('Mengunduh gambar…',20000);
      try{const r=await fetch(u);if(!r.ok)throw new Error('x');const b=await r.blob();if(!/^image\//.test(b.type))throw new Error('x');addPic(await shrink(await readFile(b)));toast('Gambar disisipkan');}
      catch{try{const im=await loadImg(u);addPic({src:u,w:im.naturalWidth||400,h:im.naturalHeight||300});toast('Gambar disisipkan, tetapi situs tidak mengizinkan penyalinan sehingga tidak ikut tersimpan ke .pptx',7000);}
        catch{toast('Gambar tidak dapat dimuat dari alamat itu',5000);}}
    },
    async screen(){
      if(!navigator.mediaDevices||!navigator.mediaDevices.getDisplayMedia)return toast('Peramban ini tidak mendukung cuplikan layar');
      try{const st=await navigator.mediaDevices.getDisplayMedia({video:true,audio:false});const v=document.createElement('video');v.srcObject=st;v.muted=true;await v.play();
        await new Promise(r=>setTimeout(r,400));const c=document.createElement('canvas');c.width=v.videoWidth;c.height=v.videoHeight;c.getContext('2d').drawImage(v,0,0);
        st.getTracks().forEach(t=>t.stop());addPic(await shrink(c.toDataURL('image/png')));}
      catch(e){if(!e||e.name!=='NotAllowedError')toast('Cuplikan layar gagal: '+(e&&e.message||e));}
    },
    async album(){
      const fs=await pick('image/*',true);if(!fs.length)return;
      try{const ns=[];for(const f of fs){const s=mkSlide('blank');s.objs.push(picObj(await shrink(await readFile(f)),1100,640));ns.push(s);}
        sync();slides.splice(cur+1,0,...ns);cur++;sel=[];draw();notes.innerHTML='';commit();toast('Album foto: '+ns.length+' slide ditambahkan');}
      catch(e){toast('Gagal membuat album: '+e.message,5000);}
    },
    chart(t){chartDlg(t);},
    sm(kind){
      const L={blocks:['Teks','Teks','Teks','Teks'],process:['Langkah 1','Langkah 2','Langkah 3'],list:['Teks','Teks','Teks']}[kind];
      const out=smartBoxes(kind,L,{x:140,y:190,w:1000,h:340});sync();objs().push(...out);sel=out.map(x=>x.id);draw();commit();
    },
    apps(){toast('Aplikasi untuk Office belum tersedia pada versi web ini');},
    hlink(){
      if(selectedText()){const u=(prompt('Alamat hyperlink:','https://')||'').trim();if(!u||u==='https://')return;target(()=>document.execCommand('createLink',false,normUrl(u)));return;}
      if(sel.length){const c=selObjs()[0].link,u=prompt('Alamat hyperlink untuk objek (kosongkan untuk menghapus):',(c&&c.url)||'https://');if(u===null)return;
        const t=u.trim();forObjs(o=>{if(t&&t!=='https://')o.link={a:'url',url:normUrl(t)};else delete o.link;});return;}
      toast('Pilih teks atau objek terlebih dahulu');
    },
    action(){
      const o=selObjs()[0];if(!o)return toast('Pilih objek terlebih dahulu');const c=o.link||{};
      const opt=[['','Tidak ada'],['next','Slide Berikutnya'],['prev','Slide Sebelumnya'],['first','Slide Pertama'],['last','Slide Terakhir'],['n','Slide nomor…'],['url','Hyperlink ke URL…'],['end','Akhiri Show']];
      dlg('Pengaturan Tindakan',`<div>Saat objek diklik pada Slide Show:</div><label><select class="at">${opt.map(([v,n])=>`<option value="${v}"${(c.a||'')===v?' selected':''}>${n}</option>`).join('')}</select></label><label>Nomor slide / URL: <input type="text" class="av" value="${esc(c.a==='n'?c.n:c.url||'')}"></label>`,
        [['Batal'],['OK',bg=>{const a=bg.querySelector('.at').value,v=bg.querySelector('.av').value.trim();
          if(a==='n'&&!(parseInt(v,10)>=1)){toast('Masukkan nomor slide yang valid');return false;}
          if(a==='url'&&!v){toast('Masukkan alamat URL');return false;}
          forObjs(x=>{if(!a)delete x.link;else x.link=a==='n'?{a,n:parseInt(v,10)}:a==='url'?{a,url:normUrl(v)}:{a};});},1]]);
    },
    comment(){
      const t=(prompt('Tulis komentar:')||'').trim();if(!t)return;sync();const s=S0();s.comments=s.comments||[];
      s.comments.push({t,x:W-52,y:16+s.comments.length*34});draw();commit();toast('Komentar ditambahkan');
    },
    comments(){
      const c=S0().comments||[];if(!c.length)return toast('Belum ada komentar pada slide ini');
      dlg('Komentar slide '+(cur+1),c.map(x=>`<div style="margin:6px 0;padding:6px 8px;background:#fff8dc;border:1px solid #e6d28a">${esc(x.t)}</div>`).join(''),[['Tutup',null,1]]);
    },
    hf(){hfDialog();},
    slidenum(){hfDialog({num:1});},
    datetime(){
      const now=new Date();
      dlg('Tanggal dan Waktu',`<div>Format yang tersedia:</div><div class="dl">${DF.map((f,i)=>`<label><input type="radio" name="dtf" value="${i}"${i===0?' checked':''}> ${esc(fmtDate(i,now))}</label>`).join('')}</div><label><input type="checkbox" class="au"> Perbarui otomatis</label>`,
        [['Batal'],['OK',bg=>{const i=+bg.querySelector('[name=dtf]:checked').value,au=bg.querySelector('.au').checked,t=fmtDate(i,new Date());
          if(hasRange()&&!au)insText(t);else addObj(mk(Object.assign({k:'tb',x:340,y:300,w:600,h:60,html:'<div>'+esc(t)+'</div>',ts:{size:24}},au?{auto:1,fmt:i}:{})),false);},1]]);
    },
    async object(){
      const f=(await pick('image/*,.txt,.csv,.tsv,.html,.htm,.md',false))[0];if(!f)return;
      try{
        if(/^image\//.test(f.type)){addPic(await shrink(await readFile(f)));return;}
        let t=await readFile(f,'text');
        if(/\.(csv|tsv)$/i.test(f.name)){const rows=parseCSV(t).slice(0,12).map(r=>r.slice(0,10));if(!rows.length)return toast('Berkas kosong');insertTable(rows.length,Math.max(...rows.map(r=>r.length)),rows);return;}
        if(/\.html?$/i.test(f.name)){const d=new DOMParser().parseFromString(t,'text/html');d.querySelectorAll('script,style').forEach(e=>e.remove());t=d.body?d.body.textContent:'';}
        const L=t.split(/\r?\n/).map(x=>x.trim()).filter(Boolean).slice(0,200);if(!L.length)return toast('Berkas tidak berisi teks');
        addObj(mk({k:'tb',x:100,y:120,w:1080,h:Math.min(560,L.length*34+20),html:L.map(x=>'<div>'+esc(x)+'</div>').join(''),ts:{size:20}}),false);
      }catch(e){toast('Gagal menyisipkan objek: '+e.message,5000);}
    },
    vtb(){drawVert=true;drawing='tb';stg.classList.add('draw');toast('Seret pada slide untuk menggambar kotak teks vertikal (Esc untuk batal)');},
    wa(i){const c=WA[+i];addObj(mk({k:'tb',x:240,y:250,w:800,h:160,anchor:'m',ts:{size:54,bold:1,align:'center'},html:`<div><span style="${c}">Teks Anda di sini</span></div>`}),false);},
    eq(i){const e=+i<0?['','Ketik persamaan di sini']:EQ[+i];insText(e[1],{x:240,w:800,h:90,anchor:'m',ts:{font:'Cambria Math',size:32,align:'center'}});},
    symbol:symbolDlg,
    video(v){
      if(v==='file'){pick('video/*').then(fs=>{const f=fs[0];if(f)addMedia({type:'video',src:URL.createObjectURL(f),name:f.name});});return;}
      const u=(prompt('Alamat video (YouTube atau berkas video langsung):','https://')||'').trim();if(!/^https?:\/\/./i.test(u))return;
      const m=/(?:youtu\.be\/|youtube\.com\/(?:watch\?(?:.*&)?v=|embed\/|shorts\/))([\w-]{11})/.exec(u);
      addMedia(m?{type:'video',embed:1,src:'https://www.youtube.com/embed/'+m[1],name:'Video YouTube'}:{type:'video',src:u,name:u.split('/').pop().split('?')[0]||'Video'});
    },
    audio(v){
      if(v==='file'){pick('audio/*').then(fs=>{const f=fs[0];if(f)addMedia({type:'audio',src:URL.createObjectURL(f),name:f.name});});return;}
      if(recorder){recorder.stop();return;}
      if(!navigator.mediaDevices||!window.MediaRecorder)return toast('Perekaman audio tidak didukung di peramban ini');
      navigator.mediaDevices.getUserMedia({audio:true}).then(st=>{const ch=[],r=new MediaRecorder(st);r.ondataavailable=e=>e.data&&e.data.size&&ch.push(e.data);
        r.onstop=()=>{st.getTracks().forEach(t=>t.stop());recorder=null;if(ch.length){addMedia({type:'audio',src:URL.createObjectURL(new Blob(ch,{type:r.mimeType||'audio/webm'})),name:'Rekaman audio'});toast('Rekaman disisipkan');}};
        r.start();recorder=r;toast('Merekam… pilih Audio ▾ Hentikan Rekaman untuk selesai',8000);}).catch(()=>toast('Mikrofon tidak dapat diakses'));
    },
    shape(id){drawVert=false;drawing=id;stg.classList.add('draw');toast('Seret pada slide untuk menggambar (Esc untuk batal)');}
  });
  Object.assign(MENUS,{
    shapes:()=>[SHAPES.map(([id,n])=>[n,'shape:'+id,shapeIcon(id)]),'sg'],
    smartins:()=>[['Daftar Blok Dasar','sm:blocks'],['Proses Dasar','sm:process'],['Daftar Vertikal','sm:list']],
    chartm:()=>Object.keys(CHARTS).map(k=>['Bagan '+CHARTS[k],'chart:'+k,{col:'▥',bar:'▤',line:'📈',pie:'◔'}[k]]),
    screenshot:()=>[['Tangkap layar / jendela / tab…','screen','🖥']],
    album:()=>[['Album Foto Baru…','album','🖼']],
    apps:()=>[['Toko Office','apps'],['Aplikasi Saya','apps']],
    textbox:()=>[['Kotak Teks Horizontal','shape:tb'],['Kotak Teks Vertikal','vtb']],
    wordart:()=>[WA.map((c,i)=>['Gaya WordArt '+(i+1),'wa:'+i,`<b style="${c}">A</b>`]),'wa'],
    eqm:()=>[...EQ.map((e,i)=>[e[0],'eq:'+i,'π']),'-',['Persamaan Baru','eq:-1','✎']],
    video:()=>[['Video dari Berkas…','video:file','🎞'],['Video Online (URL)…','video:url','🌐']],
    audio:()=>[['Audio dari Berkas…','audio:file','🔊'],[recorder?'Hentikan Rekaman':'Rekam Audio…','audio:rec','🎙']]
  });

  /* ===== tab DESAIN ===== */
  let themeOn=false,curT=0,curV=0;
  function styleSlide(sl){
    const t=THEMES[curT],c=pal(t,curV);
    sl.bg=c.bg;sl.objs=sl.objs.filter(o=>!o.deco);
    sl.objs.forEach(o=>{if(o.k!=='ph')return;
      o.ts.font=o.role==='title'?t.hf:t.bf;
      o.ts.color=o.role==='title'?c.title:(o.role==='sub'&&sl.layout==='section'?c.mute:c.body);});
    sl.objs.unshift(...DECO[t.d].map(([k,x,y,w,h])=>mk({k:'sh',shape:k==='o'?'oval':'rect',deco:1,x:Math.round(x*W),y:Math.round(y*H),w:Math.round(w*W),h:Math.round(h*H),
      fill:k==='f'?'':c.deco,stroke:k==='f'?c.deco:'',sw:4,html:''})));
  }
  function refreshDesign(){
    $$('.tt[data-t]').forEach(b=>b.classList.toggle('on',+b.dataset.t===curT));
    const vg=$('.vgal');if(vg)vg.innerHTML=[0,1,2,3].map(v=>`<button class="tt${v===curV?' on':''}" data-a="variant:${v}" title="Varian ${v+1}">${tThumb(THEMES[curT],v)}</button>`).join('');
  }
  function applyTheme(i,v){
    sync();curT=i;curV=v;themeOn=true;sel=[];slides.forEach(styleSlide);draw();commit();refreshDesign();
    toast('Tema "'+THEMES[i].n+'" (varian '+(v+1)+') diterapkan ke semua slide');
  }
  function setBg(v,all){
    sync();(all?slides:[S0()]).forEach(x=>{if(v)x.bg=v;else delete x.bg;});draw();commit();
  }
  Object.assign(acts,{
    theme(i){applyTheme(+i,curV);},
    variant(v){applyTheme(curT,+v);},
    tscroll(d){$('.tgal').scrollBy({left:+d*330,behavior:'smooth'});},
    bgfmt(){
      const cur=/^#[0-9a-f]{6}$/i.test(S0().bg||'')?S0().bg:'#ffffff';
      dlg('Format Latar Belakang',`<label>Warna isian: <input type="color" class="bgc" value="${cur}" style="width:64px;height:28px"></label>`,
        [['Setel Ulang',()=>setBg(themeOn?pal(THEMES[curT],curV).bg:null,false)],
         ['Terapkan ke Semua',bg=>setBg(bg.querySelector('.bgc').value,true)],
         ['Batal'],['Terapkan',bg=>setBg(bg.querySelector('.bgc').value,false),1]]);
    }
  });
  Object.assign(MENUS,{themes:()=>[THEMES.map((t,i)=>[t.n,'theme:'+i,`<span class="tbx">${tThumb(t,0)}</span>`]),'grid thm']});
  const _lay=acts.layout,_rst=acts.reset,_new=acts.new;
  acts.layout=l=>{_lay(l);if(themeOn){styleSlide(S0());draw();commit();}};
  acts.reset=()=>{_rst();if(themeOn){styleSlide(S0());draw();commit();}};
  acts.new=()=>{_new();themeOn=false;curT=0;curV=0;refreshDesign();};
  refreshDesign();

  /* ===== tab TRANSISI ===== */
  function syncTr(){
    const t=getTr(S0()),q=n=>$('[data-s='+n+']');if(!q('trdur'))return;
    $$('.trb').forEach(b=>b.classList.toggle('on',b.dataset.tr===t.t));
    q('trsnd').value=t.snd;q('trdur').value=t.dur;q('trclk').checked=!!t.click;q('traft').checked=!!t.afterOn;q('traftv').value=t.after;
    const e=$('.treff');if(e)e.disabled=!TROPT[t.t];
  }
  function setTr(p){sync();S0().tr=Object.assign(getTr(S0()),p);commit();syncTr();}
  Object.assign(acts,{
    trset(id){setTr({t:id,o:(TROPT[id]||[['']])[0][0]});acts.trprev(true);},
    treff(v){setTr({o:v});acts.trprev(true);},
    trscroll(d){$('.trg').scrollBy({left:+d*320,behavior:'smooth'});},
    trall(){sync();const t=getTr(S0());slides.forEach(x=>x.tr=Object.assign({},t));commit();toast('Transisi diterapkan ke semua slide');},
    trprev(quiet){
      sync();const t=getTr(S0());
      if(t.t==='none'||t.t==='cut'){if(!quiet)toast('Pilih transisi terlebih dahulu');return;}
      hold.querySelectorAll('.trpv').forEach(x=>x.remove());
      const b=document.createElement('div');b.className='trpv';b.style.cssText=`position:absolute;left:0;top:0;width:${W*scale}px;height:${H*scale}px;overflow:hidden;z-index:8;background:#fff`;
      const od=svEl(cur>0?slides[cur-1]:null,scale,false),nw=svEl(S0(),scale,false);b.append(od,nw);hold.appendChild(b);playSnd(t.snd);runTr(b,od,nw,t,()=>b.remove());
    }
  });
  Object.assign(sels,{
    trsnd:v=>{setTr({snd:v});playSnd(v);},
    trdur:v=>{const d=parseFloat(String(v).replace(',','.'));if(d>0)setTr({dur:Math.min(60,d)});},
    trclk:()=>setTr({click:$('[data-s=trclk]').checked?1:0}),
    traft:()=>setTr({afterOn:$('[data-s=traft]').checked?1:0}),
    traftv:v=>setTr({after:Math.max(0,parseFloat(v)||0)})
  });
  Object.assign(MENUS,{
    trmenu:()=>[TRS.map(([id,n])=>[n,'trset:'+id,trIcon(id)]),'grid thm'],
    treff:()=>{const t=getTr(S0()),L=TROPT[t.t]||[],cu=t.o||(L[0]&&L[0][0]);return L.map(([v,n])=>[(v===cu?'✓ ':'\u2003')+n,'treff:'+v]);}
  });
  syncTr();

  /* ===== tab ANIMASI ===== */
  let anPaint=null,anT=0;
  const addAo=(sl,id)=>{sl.ao=sl.ao||[];if(!sl.ao.includes(id))sl.ao.push(id);};
  function animBadges(){anList(S0()).forEach((o,i)=>stg.insertAdjacentHTML('beforeend',`<div class="anb" style="left:${Math.max(0,o.x-30)}px;top:${o.y}px">${i+1}</div>`));}
  function renderPane(){
    const pn=$('.apane');if(!pn||pn.hidden)return;
    const L=anList(S0()),ic={click:'🖱',with:'◐',after:'◔'};
    $('.apl').innerHTML=L.length?L.map((o,i)=>{const a=getAn(o),n=(AEFF.find(e=>e[0]===a.t)||[,''])[1];
      return `<div class="ai${sel.includes(o.id)?' on':''}" data-a="anpick:${o.id}" title="${esc(n+(a.x?' (keluar)':''))}"><i>${i+1}</i><span>${ic[a.start]}</span><b>${esc(objName(o))}</b><em>${a.x?'⇥':'★'}</em></div>`;}).join('')
      :'<div class="am">Belum ada animasi. Pilih objek pada slide, lalu pilih sebuah efek.</div>';
  }
  function syncAn(){
    const q=n=>$('[data-s='+n+']');if(!q('andur'))return;
    const os=selObjs(),o=os[0],a=getAn(o),has=os.length>0,act=has&&hasAn(o);
    $$('.anb2').forEach(b=>{b.disabled=!has;b.classList.toggle('on',act&&a.t===b.dataset.an&&!a.x);});
    ['anstart','andur','andelay'].forEach(n=>q(n).disabled=!act);
    q('anstart').value=a.start;q('andur').value=a.dur;q('andelay').value=a.delay;
    $('.aneff').disabled=!(act&&AOPT[a.t]);
    ['anup','andown'].forEach(n=>$('[data-a='+n+']').disabled=!act);
    if(anPaint&&os.length===1&&o.id!==anPaint.src){const pt=anPaint;anPaint=null;
      setTimeout(()=>{sync();o.an=Object.assign({},pt.an);addAo(S0(),o.id);draw();commit();toast('Animasi disalin');},0);}
    renderPane();
  }
  function setAnP(p){forObjs(o=>{if(!hasAn(o))return;o.an=Object.assign(getAn(o),p);});}
  function anMove(d){
    const o=selObjs().find(hasAn);if(!o)return;sync();const sl=S0(),ids=anList(sl).map(x=>x.id),i=ids.indexOf(o.id),j=i+d;
    if(j<0||j>=ids.length)return;[ids[i],ids[j]]=[ids[j],ids[i]];sl.ao=ids;draw();commit();
  }
  function playIn(root,sl,steps,only){
    clearTimeout(anT);anInit(root,sl,only);let t=0;
    steps.forEach(st=>{setTimeout(()=>anPlayItems(root,st.items),t);t+=st.ms+250;});
    anT=setTimeout(()=>draw(),t+150);
  }
  Object.assign(acts,{
    anset(id){
      forObjs(o=>{if(id==='none'){delete o.an;if(S0().ao)S0().ao=S0().ao.filter(x=>x!==o.id);}
        else{o.an=Object.assign(getAn(o),{t:id,o:(AOPT[id]||[['']])[0][0],x:0});addAo(S0(),o.id);}});
      if(id!=='none')acts.anprev(true);
    },
    anadd(v){const[k,id]=v.split(':');
      forObjs(o=>{o.an=Object.assign(getAn(o),{t:id,o:(AOPT[id]||[['']])[0][0],x:k==='out'?1:0});addAo(S0(),o.id);});acts.anprev(true);},
    aneff(v){setAnP({o:v});acts.anprev(true);},
    anscroll(d){const g=$$('.trg')[1];if(g)g.scrollBy({left:+d*320,behavior:'smooth'});},
    anprev(only){
      sync();const sl=S0();
      if(only===true){const os=selObjs().filter(hasAn);if(!os.length)return;
        const items=os.map(o=>{const a=getAn(o);return{o,a,at:(+a.delay||0)*1000,D:(a.dur>0?+a.dur:.5)*1000};});
        playIn(stg,sl,[{items,ms:Math.max(...items.map(x=>x.at+x.D))}],os);return;}
      const steps=anSteps(sl);if(!steps.length)return toast('Slide ini belum memiliki animasi');
      sel=[];decorate();playIn(stg,sl,steps);
    },
    anpane(){const pn=$('.apane');pn.hidden=!pn.hidden;pp.classList.toggle('apo',!pn.hidden);renderPane();if(autoFit)fit();},
    anpick(id){sel=[id];decorate();},
    anpaint(){const o=selObjs()[0];if(!o||!hasAn(o))return toast('Pilih objek yang sudah memiliki animasi');
      anPaint={an:JSON.parse(JSON.stringify(o.an)),src:o.id};toast('Pilih objek tujuan untuk menyalin animasi');},
    anup(){anMove(-1);},andown(){anMove(1);}
  });
  Object.assign(sels,{
    anstart:v=>setAnP({start:v}),
    andur:v=>{const d=parseFloat(v);if(d>0)setAnP({dur:Math.min(60,d)});},
    andelay:v=>setAnP({delay:Math.max(0,parseFloat(v)||0)})
  });
  Object.assign(MENUS,{
    aneff:()=>{const o=selObjs()[0],a=getAn(o),L=AOPT[a.t]||[],cu=a.o||(L[0]&&L[0][0]);return L.map(([v,n])=>[(v===cu?'✓ ':'\u2003')+n,'aneff:'+v]);},
    anadd:()=>[...AEFF.filter(e=>e[0]!=='none').map(([id,n])=>['Masuk · '+n,'anadd:in:'+id,aIcon(id)]),'-',...AEFF.filter(e=>e[0]!=='none').map(([id,n])=>['Keluar · '+n,'anadd:out:'+id,aIcon(id)])]
  });

  /* ===== tab SLIDE SHOW ===== */
  const slideTitle=s=>{const t=s.objs.map(o=>textOf(o.html||'').trim()).find(Boolean)||'(tanpa teks)';return t.length>38?t.slice(0,38)+'…':t;};
  const ensureCid=s=>{if(!s.cid)s.cid='c'+(cidn++);return s.cid;};
  function syncSs(){
    const q=n=>$('[data-s='+n+']');if(!q('shnarr'))return;
    q('shnarr').checked=!!ss.playNarr;q('shtim').checked=!!ss.useTimings;q('shmed').checked=!!ss.mediaCtrl;q('shpv').checked=!!ss.presenter;
    if([...q('shmon').options].some(x=>x.value===ss.monitor))q('shmon').value=ss.monitor;
    const h=$('[data-a=hideslide]');if(h&&S0())h.classList.toggle('on',!!S0().hid);
  }
  async function loadScreens(){
    if(scrTried||!window.getScreenDetails)return;scrTried=true;
    try{const d=await window.getScreenDetails();scrList=[...d.screens];const sel=$('[data-s=shmon]');if(!sel||scrList.length<2)return;
      sel.innerHTML='<option value="auto">Otomatis</option>'+scrList.map((s,i)=>`<option value="${i}">Monitor ${i+1}: ${esc(s.label||(s.width+'×'+s.height))}${s.isPrimary?' (utama)':''}</option>`).join('');
      if(ss.monitor==='main')ss.monitor='auto';sel.value=ss.monitor;}catch{}
  }
  root.addEventListener('focusin',e=>{if(e.target.matches&&e.target.matches('[data-s=shmon]'))loadScreens();});
  root.addEventListener('mousedown',e=>{if(e.target.matches&&e.target.matches('[data-s=shmon]'))loadScreens();});

  function setupDlg(){
    const c=ss,n=slides.length,ck=v=>v?' checked':'',opt=customShows.map(x=>`<option${c.custom===x.name?' selected':''}>${esc(x.name)}</option>`).join('');
    dlg('Siapkan Slide Show',`
     <fieldset><legend>Jenis presentasi</legend>
      <label><input type="radio" name="sst" value="full"${ck(c.type==='full')}> Dipresentasikan oleh pembicara (layar penuh)</label>
      <label><input type="radio" name="sst" value="win"${ck(c.type==='win')}> Dilihat oleh individu (jendela)</label>
      <label><input type="radio" name="sst" value="kiosk"${ck(c.type==='kiosk')}> Dilihat di kios (layar penuh)</label></fieldset>
     <fieldset><legend>Opsi tampilan</legend>
      <label><input type="checkbox" class="sl"${ck(c.loop)}> Ulangi terus sampai 'Esc' ditekan</label>
      <label><input type="checkbox" class="sn"${ck(!c.playNarr)}> Tampilkan tanpa narasi</label>
      <label><input type="checkbox" class="sa"${ck(c.noAnim)}> Tampilkan tanpa animasi</label>
      <label>Warna pena: <input type="color" class="sp" value="${c.pen}"></label></fieldset>
     <fieldset><legend>Tampilkan slide</legend>
      <label><input type="radio" name="ssr" value="all"${ck(c.range==='all')}> Semua</label>
      <label><input type="radio" name="ssr" value="range"${ck(c.range==='range')}> Dari: <input type="number" class="sf" min="1" max="${n}" value="${Math.min(c.from,n)}"> Sampai: <input type="number" class="se" min="1" max="${n}" value="${Math.min(c.to||n,n)}"></label>
      <label><input type="radio" name="ssr" value="custom"${ck(c.range==='custom')}> Slide show kustom: <select class="sc">${opt||'<option value="">(belum ada)</option>'}</select></label></fieldset>
     <fieldset><legend>Slide berikutnya</legend>
      <label><input type="radio" name="ssa" value="m"${ck(!c.useTimings)}> Manual</label>
      <label><input type="radio" name="ssa" value="t"${ck(c.useTimings)}> Menggunakan pengaturan waktu, jika ada</label></fieldset>`,
     [['Batal'],['OK',bg=>{
       const g=q=>bg.querySelector(q),v=q=>bg.querySelector('input[name='+q+']:checked').value,r=v('ssr');
       if(r==='custom'&&!g('.sc').value){toast('Belum ada slide show kustom. Buat dulu lewat Slide Show Kustom.');return false;}
       const a=Math.max(1,Math.min(n,parseInt(g('.sf').value)||1)),b=Math.max(a,Math.min(n,parseInt(g('.se').value)||n));
       Object.assign(ss,{type:v('sst'),loop:g('.sl').checked?1:0,playNarr:g('.sn').checked?0:1,noAnim:g('.sa').checked?1:0,pen:g('.sp').value,range:r,from:a,to:b,custom:g('.sc').value,useTimings:v('ssa')==='t'?1:0});
       if(ss.type==='kiosk')ss.loop=1;syncSs();}, 1]]);
  }

  /* slide show kustom */
  function customMgr(){
    const bg=dlg('Slide Show Kustom',customShows.length?`<div>Slide show kustom:</div><select multiple size="6" class="cl" style="height:120px">${customShows.map((x,k)=>`<option value="${k}">${esc(x.name)} (${x.ids.length} slide)</option>`).join('')}</select>`:'<div>Belum ada slide show kustom. Klik <b>Baru…</b> untuk membuat.</div>',
      [['Baru…',()=>{customEdit(-1);}],['Edit…',b=>{const s=b.querySelector('.cl'),k=s&&s.selectedIndex;if(!s||k<0){toast('Pilih slide show terlebih dahulu');return false;}customEdit(k);}],
       ['Hapus',b=>{const s=b.querySelector('.cl'),k=s&&s.selectedIndex;if(!s||k<0){toast('Pilih slide show terlebih dahulu');return false;}const nm=customShows[k].name;customShows.splice(k,1);if(ss.custom===nm){ss.custom='';if(ss.range==='custom')ss.range='all';}customMgr();}],
       ['Putar',b=>{const s=b.querySelector('.cl'),k=s&&s.selectedIndex;if(!s||k<0){toast('Pilih slide show terlebih dahulu');return false;}setTimeout(()=>acts.cshowplay(k),0);}],
       ['Tutup',null,1]]);
    return bg;
  }
  function customEdit(k){
    sync();const ex=k>=0?customShows[k]:null,opts=slides.map((s,i)=>`<option value="${ensureCid(s)}">${i+1}. ${esc(slideTitle(s))}</option>`).join('');
    const label=id=>{const j=cidIdx(id);return j<0?null:`${j+1}. ${slideTitle(slides[j])}`;};
    const bg=dlg(ex?'Edit Slide Show Kustom':'Slide Show Kustom Baru',`<label>Nama slide show: <input type="text" class="cn" value="${esc(ex?ex.name:'Slide Show Kustom '+(customShows.length+1))}"></label>
      <div class="cs2"><div><div>Slide dalam presentasi:</div><select multiple class="ca">${opts}</select></div>
      <div class="cb"><button type="button" class="ad">Tambah ▸</button><button type="button" class="rm">◂ Hapus</button><button type="button" class="up">▲</button><button type="button" class="dn">▼</button></div>
      <div><div>Slide dalam slide show kustom:</div><select multiple class="cu"></select></div></div>`,
      [['Batal',()=>{customMgr();}],['OK',b=>{
        const nm=b.querySelector('.cn').value.trim(),ids=[...b.querySelector('.cu').options].map(x=>x.value);
        if(!nm){toast('Isi nama slide show');return false;}
        if(customShows.some((x,j)=>j!==k&&x.name===nm)){toast('Nama sudah dipakai');return false;}
        if(!ids.length){toast('Tambahkan minimal satu slide');return false;}
        if(ex){if(ss.custom===ex.name)ss.custom=nm;ex.name=nm;ex.ids=ids;}else customShows.push({name:nm,ids});
        customMgr();},1]]);
    bg.querySelector('.dlg').style.minWidth='min(520px,calc(100% - 24px))';
    const A=bg.querySelector('.ca'),U=bg.querySelector('.cu'),mkOpt=id=>{const x=document.createElement('option');x.value=id;x.textContent=label(id)||'(slide dihapus)';return x;};
    if(ex)ex.ids.filter(id=>cidIdx(id)>=0).forEach(id=>U.appendChild(mkOpt(id)));
    bg.querySelector('.ad').onclick=()=>[...A.selectedOptions].forEach(o=>U.appendChild(mkOpt(o.value)));
    bg.querySelector('.rm').onclick=()=>[...U.selectedOptions].forEach(o=>o.remove());
    bg.querySelector('.up').onclick=()=>[...U.selectedOptions].forEach(o=>{const pr=o.previousElementSibling;if(pr&&!pr.selected)U.insertBefore(o,pr);});
    bg.querySelector('.dn').onclick=()=>[...U.selectedOptions].reverse().forEach(o=>{const nx=o.nextElementSibling;if(nx&&!nx.selected)U.insertBefore(nx,o);});
  }

  /* hasil latih / rekam */
  function finishRec(T,Bl,o){
    const ks=Object.keys(T).map(Number);if(!ks.length)return;
    const tot=ks.reduce((s,x)=>s+T[x],0),what=[o.timing?'pengaturan waktu':'',o.narr&&Object.keys(Bl).length?'narasi':''].filter(Boolean).join(' dan ');
    if(!what)return;
    dlg('Slide show berakhir',`<div>Total waktu: <b>${fmtT(tot)}</b> untuk ${ks.length} slide.</div><div style="margin-top:6px">Simpan ${what}?</div>`,
      [['Jangan simpan'],['Simpan',()=>{
        sync();
        ks.forEach(k=>{const s=slides[k];if(!s)return;
          if(o.timing)s.tr=Object.assign(getTr(s),{afterOn:1,after:T[k]});
          if(o.narr&&Bl[k])s.narr={src:URL.createObjectURL(Bl[k]),t:T[k]};});
        commit();syncSs();toast('Tersimpan: '+what+'. Narasi tidak ikut tersimpan ke berkas .pptx.',5000);},1]]);
  }
  function recDlg(from){
    dlg('Rekam Slide Show','<div>Pilih yang ingin direkam sebelum memulai:</div><label><input type="checkbox" class="r1" checked> Waktu slide dan animasi</label><label><input type="checkbox" class="r2" checked> Narasi (membutuhkan mikrofon)</label>',
      [['Batal'],['Mulai Merekam',bg=>{
        const timing=bg.querySelector('.r1').checked,narr=bg.querySelector('.r2').checked;
        if(!timing&&!narr){toast('Pilih minimal satu opsi');return false;}
        (async()=>{let mic=null;
          if(narr){try{mic=await navigator.mediaDevices.getUserMedia({audio:true});}catch{
            if(!timing){toast('Mikrofon tidak dapat diakses');return;}toast('Mikrofon tidak dapat diakses — hanya merekam waktu slide',5000);}}
          show(from,{mode:'record',timing,narr:!!mic,mic});})();
      },1]]);
  }

  Object.assign(acts,{
    shsetup:setupDlg,
    hideslide(){sync();const s=S0();if(s.hid)delete s.hid;else s.hid=1;commit();syncSs();toast(s.hid?'Slide '+(cur+1)+' disembunyikan dan dilewati saat slide show':'Slide '+(cur+1)+' ditampilkan kembali');},
    rehearse(){show(null,{mode:'rehearse',timing:1});},
    recfrom(v){recDlg(v==='cur'?cur:null);},
    recclr(v){const[w,sc]=v.split(':');sync();(sc==='cur'?[S0()]:slides).forEach(s=>{if(w==='time'){if(s.tr){s.tr.afterOn=0;s.tr.after=0;}}else delete s.narr;});commit();syncSs();toast(w==='time'?'Pengaturan waktu dihapus':'Narasi dihapus');},
    cshowplay(v){const c=customShows[+v];if(!c)return;show(null,{custom:c});},
    cshowmgr:customMgr
  });
  Object.assign(sels,{
    shnarr:()=>{ss.playNarr=$('[data-s=shnarr]').checked?1:0;},
    shtim:()=>{ss.useTimings=$('[data-s=shtim]').checked?1:0;},
    shmed:()=>{ss.mediaCtrl=$('[data-s=shmed]').checked?1:0;},
    shpv:()=>{ss.presenter=$('[data-s=shpv]').checked?1:0;},
    shmon:v=>{ss.monitor=v;}
  });
  Object.assign(MENUS,{
    online:()=>[['Unduh .pptx untuk dibagikan','save','💾'],['Cetak / simpan sebagai PDF','print','🖨']],
    cshow:()=>[...customShows.map((c,k)=>[c.name,'cshowplay:'+k,'▶']),...(customShows.length?['-']:[]),['Slide Show Kustom…','cshowmgr']],
    rec:()=>[['Mulai Merekam dari Awal…','recfrom:0','⏺'],['Mulai Merekam dari Slide Saat Ini…','recfrom:cur','⏺'],'-',
      ['Hapus Waktu pada Slide Saat Ini','recclr:time:cur'],['Hapus Waktu pada Semua Slide','recclr:time:all'],['Hapus Narasi pada Slide Saat Ini','recclr:narr:cur'],['Hapus Narasi pada Semua Slide','recclr:narr:all']]
  });
  root.addEventListener('keydown',e=>{
    if(e.key!=='F5'||pp.querySelector('.show'))return;e.preventDefault();if(e.shiftKey)acts.show();else acts.showfirst();});
  syncSs();

  /* ===== event ===== */
  root.addEventListener('mousedown',e=>{
    if(!e.target.closest('.pm'))closePop();
    if(e.target.closest('.b,.tab,.status button,.pm button,.menu button,.clr span'))e.preventDefault();
  });
  root.addEventListener('click',e=>{
    const m=e.target.closest('[data-m]');if(m){if(m.dataset.m==='table')tablePicker(m);else openMenu(m,m.dataset.m);return;}
    const t=e.target.closest('[data-a],[data-tab]');if(!t)return;
    if(t.dataset.tab){
      const small=matchMedia('(max-width:700px),(max-height:500px)').matches;
      if(small&&t.classList.contains('on')&&!pp.classList.contains('rc')){pp.classList.add('rc');t.classList.remove('on');return;}
      pp.classList.remove('rc');pp.classList.toggle('ant',t.dataset.tab==='anim');$$('.tab[data-tab]').forEach(x=>x.classList.toggle('on',x===t));$$('[data-pane]').forEach(p=>p.hidden=p.dataset.pane!==t.dataset.tab);return;}
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
        if(id==='tb'){const dv=drawVert;drawVert=false;if(click){w=dv?60:360;h=dv?360:60;}addObj(mk({k:'tb',vert:dv?'v90':'h',x:Math.round(px),y:Math.round(py),w:Math.round(Math.max(w,60)),h:Math.round(Math.max(h,40))}),true);return;}
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
    if(e.target.closest('.cm'))return;
    if(drawing){e.preventDefault();startDraw(e);return;}
    const hd=e.target.closest('.hd');if(hd){e.preventDefault();startResize(e,hd.dataset.h);return;}
    const el=e.target.closest('.ob');
    if(!el){getSelection().removeAllRanges();lastR=null;if(document.activeElement&&document.activeElement!==document.body)document.activeElement.blur();
      if(!e.shiftKey){sel=[];decorate();}e.preventDefault();stg.focus({preventScroll:true});startMarquee(e,e.shiftKey);if(painting)applyPaint();return;}
    const o=byId(el.dataset.id),ed=e.target.closest('.ed'),inTx=!!e.target.closest('.tx')&&(o.k!=='sh'||el.classList.contains('edt'));
    if(ed||((o.k==='sh'||o.media)&&!el.classList.contains('edt'))){
      e.preventDefault();
      if(e.shiftKey)sel=sel.includes(o.id)?sel.filter(x=>x!==o.id):[...sel,o.id];else if(!sel.includes(o.id))sel=[o.id];
      getSelection().removeAllRanges();lastR=null;decorate();stg.focus({preventScroll:true});startMove(e);return;}
    if(!inTx){e.preventDefault();const t=el.querySelector('.tx');if(t){sel=[o.id];decorate();t.focus();const r=document.createRange();r.selectNodeContents(t);r.collapse(false);const s=getSelection();s.removeAllRanges();s.addRange(r);}return;}
    if(sel.length!==1||sel[0]!==o.id){sel=[o.id];decorate();}
  });
  stg.addEventListener('click',e=>{const c=e.target.closest('.cm');if(!c)return;const k=+c.dataset.ci,cm=(S0().comments||[])[k];if(!cm)return;
    dlg('Komentar',`<div style="padding:6px 8px;background:#fff8dc;border:1px solid #e6d28a">${esc(cm.t)}</div>`,[['Hapus komentar',()=>{sync();S0().comments.splice(k,1);draw();commit();}],['Tutup',null,1]]);});
  stg.addEventListener('dblclick',e=>{const el=e.target.closest('.ob'),o=el&&byId(el.dataset.id);
    if(o&&o.k==='sh'&&o.shape!=='line'&&o.shape!=='arrow'){el.classList.add('edt');const t=el.querySelector('.tx');t.focus();const r=document.createRange();r.selectNodeContents(t);r.collapse(false);const s=getSelection();s.removeAllRanges();s.addRange(r);}});
  stg.addEventListener('mouseup',e=>{if(painting&&(selectedText()||sel.length))setTimeout(applyPaint,0);syncUI();});
  stg.addEventListener('focusin',e=>{const t=e.target.closest&&e.target.closest('.tx');if(!t)return;const el=t.closest('.ob'),o=byId(el.dataset.id);if(!o)return;
    if(sel.length!==1||sel[0]!==o.id){sel=[o.id];decorate();}
    if(o.bul&&!t.textContent.trim()&&!t.querySelector('li')){t.innerHTML='<ul><li><br></li></ul>';const li=t.querySelector('li'),r=document.createRange();r.selectNodeContents(li);r.collapse(true);const s=getSelection();s.removeAllRanges();s.addRange(r);}});
  stg.addEventListener('focusout',e=>{const t=e.target.closest&&e.target.closest('.tx');if(t&&!t.textContent.trim()&&!t.querySelector('img,table'))t.innerHTML='';});
  stg.addEventListener('input',e=>{if(e.target.closest('.tx'))commitSoon();});
  stg.addEventListener('paste',e=>{const t=e.target.closest&&e.target.closest('.tx');if(!t)return;e.preventDefault();const x=(e.clipboardData||window.clipboardData).getData('text/plain');if(x)document.execCommand('insertText',false,x);});
  stg.addEventListener('keydown',e=>{
    const intx=e.target.closest&&e.target.closest('.tx');
    if(e.key==='Escape'){if(drawing){drawing=null;stg.classList.remove('draw');return;}if(painting){painting=null;stg.classList.remove('paint');return;}
      if(intx){e.preventDefault();intx.blur();getSelection().removeAllRanges();lastR=null;stg.focus({preventScroll:true});}else{sel=[];decorate();}return;}
    if(e.key==='Tab'&&intx){const s=getSelection(),n=s.anchorNode,el=n&&(n.nodeType===3?n.parentElement:n),td=el&&el.closest&&el.closest('td');
      if(td&&intx.contains(td)){e.preventDefault();const tb=td.closest('table'),cells=[...tb.querySelectorAll('td')],i=cells.indexOf(td);let nx=cells[i+(e.shiftKey?-1:1)];
        if(!nx&&!e.shiftKey){const tr=td.parentElement,nr=tr.cloneNode(true);tr.parentNode.appendChild(nr);
          nr.querySelectorAll('td').forEach(c=>{c.innerHTML='<br>';c.style.background=(tb.rows.length-1)%2?'#cfd5ea':'#e9ebf5';c.style.color='#000000';c.style.fontWeight='';});
          nx=nr.cells[0];const ob=intx.closest('.ob'),o=byId(ob.dataset.id);if(o){o.h+=46;ob.style.height=o.h+'px';}}
        if(nx){const r=document.createRange();r.selectNodeContents(nx);s.removeAllRanges();s.addRange(r);}commitSoon();return;}}
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
    pop([['Slide Baru','newslide'],['Duplikat Slide','dup'],['Hapus Slide','delslide'],'-',['Mulai dari Slide Ini','show'],['Sembunyikan / Tampilkan Slide','hideslide']],e.clientX,e.clientY);};
  thumbsEl.addEventListener('contextmenu',ctx);sorter.addEventListener('contextmenu',ctx);
  acts.dup=()=>{sync();const c=JSON.parse(JSON.stringify(S0()));delete c.sec;delete c.cid;c.objs.forEach(o=>o.id='o'+(uid++));slides.splice(cur+1,0,c);cur++;sel=[];draw();notes.innerHTML=c.notes||'';commit();};
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
    else if(k==='k'&&!inInput){e.preventDefault();acts.hlink();}
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
