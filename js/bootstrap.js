/* Loads the JSON datasets into the globals app.js expects, then starts the app. */
(function(){
  const FILES={DATA:"data/locations_weather.json",ALERTS:"data/alerts.json",ACC:"data/accuracy.json",CROP:"data/crops.json"};
  Promise.all(Object.keys(FILES).map(k=>fetch(FILES[k]).then(r=>{
    if(!r.ok)throw new Error(FILES[k]+" "+r.status);
    return r.json();
  }).then(d=>{window[k]=d;}))).then(()=>{
    const s=document.createElement("script");s.src="js/app.js";document.body.appendChild(s);
  }).catch(e=>{
    document.body.insertAdjacentHTML("afterbegin",'<p style="padding:16px;color:#b00">Could not load data: '+e.message+'. Serve this folder over HTTP (for example <code>python3 -m http.server</code>); opening index.html from disk blocks JSON loading.</p>');
  });
})();
