(function(){
  window.Apps=window.Apps||{};
  Apps.excel={
    mount(root,opts){
      root.innerHTML='<div style="height:100%;display:grid;place-items:center;font-family:Segoe UI,system-ui,sans-serif;text-align:center">'+
        '<div><h1 style="font-weight:600">Hello World</h1><p>Excel (belum dibuat)</p>'+
        '<button id="back" style="font:inherit;padding:6px 14px;cursor:pointer">Kembali ke menu</button></div></div>';
      root.querySelector('#back').onclick=opts.onHome;
    }
  };
})();
