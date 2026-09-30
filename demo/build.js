const fs=require('fs'), path=require('path');
const APP=path.join(__dirname,'../apps-script');
const rd=f=>fs.readFileSync(path.join(APP,f),'utf8');
const expand=h=>h.replace(/<\?!= include\('(\w+)'\) \?>/g,(_,n)=>rd(n+'.html'));
const js=s=>JSON.stringify(s).replace(/<\//g,'<\\/').replace(/<!--/g,'<\\!--');
const code=rd('Code.gs');
const names=[...code.matchAll(/^function (\w+)/gm)].map(m=>m[1]);
const SHIM=`<script>(function(){var n=0,cb={};window.addEventListener('message',function(e){var d=e.data;if(d&&d.__gcRes){var c=cb[d.id];if(c){delete cb[d.id];if(d.ok){c.s&&c.s(d.v)}else{c.f&&c.f(new Error(d.err))}}}});function mk(s,f){return new Proxy({},{get:function(_,k){if(k==='withSuccessHandler')return function(h){return mk(h,f)};if(k==='withFailureHandler')return function(h){return mk(s,h)};return function(){var id=++n;cb[id]={s:s,f:f};parent.postMessage({__gcReq:1,id:id,fn:k,args:[].slice.call(arguments)},'*')}}})}window.google={script:{run:mk(null,null)}};window.confirm=function(){return true};window.prompt=function(){return null};window.alert=function(){};document.addEventListener('click',function(e){var a=e.target.closest&&e.target.closest('a');if(!a)return;e.preventDefault();parent.postMessage({__gcNav:1,href:a.getAttribute('href')||''},'*')},true);})();</script>`;
const wrapper=fs.readFileSync(path.join(__dirname,'wrapper.html'),'utf8');
const script=`
${fs.readFileSync(path.join(__dirname,'mocks.js'),'utf8')}
function bootBackend() {
  var G = makeGoogle('me@example.com');
  var PropertiesService = G.PropertiesService, CacheService = G.CacheService, SpreadsheetApp = G.SpreadsheetApp, Session = G.Session,
      LockService = G.LockService, MailApp = G.MailApp, Logger = G.Logger, Utilities = G.Utilities, Calendar = G.Calendar,
      CalendarApp = G.CalendarApp, ScriptApp = G.ScriptApp, HtmlService = G.HtmlService, UrlFetchApp = G.UrlFetchApp;
${code}
  return { G: G, api: { ${names.map(n=>n+': '+n).join(', ')} } };
}
var TPL_ADMIN = ${js(expand(rd('Admin.html')).replace('<head>','<head>'+SHIM))};
var TPL_BOOK = ${js(expand(rd('Book.html')).replace('<head>','<head>'+SHIM))};
var TPL_SHARE = ${js(expand(rd('Share.html')).replace('<head>','<head>'+SHIM))};
`;
const out=wrapper.replace('/*__BUILD__*/',()=>script);
fs.writeFileSync(path.join(__dirname,'index.html'),out);
console.log('built',(out.length/1024).toFixed(0)+'KB', names.length,'functions');
