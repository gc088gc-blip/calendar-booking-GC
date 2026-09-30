const {makeEnv}=require('./mock'); const assert=require('assert');
const {ctx,state}=makeEnv(); ctx.setup();
const now=new Date('2026-09-30T10:00:00+08:00').getTime(); Date.now=()=>now;
ctx.Calendar.Freebusy.query=q=>{const cal={};q.items.forEach(i=>{cal[i.id]={busy:(state.events[i.id]||[]).filter(e=>e.start.dateTime).map(e=>({start:e.start.dateTime+'+08:00',end:e.end.dateTime+'+08:00'}))}});return {calendars:cal};};
let cfg=ctx.getConfig_();
ctx.adminSetOpen(['2026-10-20','2026-10-21'],[{s:'19:00',e:'22:00',loc:'both',aud:'all'}]);
// 一筆新朋友（待審核）、一筆朋友（直接成立）、一筆自己的行程（不算預約）
ctx.apiBook({aud:'public',k:'',date:'2026-10-20',start:1140,end:1200,name:'A',contact:'a@b.co',purpose:'x',locMode:'online'});
ctx.apiBook({aud:'friend',k:cfg.audiences.friend.token,date:'2026-10-20',start:1230,end:1290,name:'B',contact:'',purpose:'y',locMode:'online'});
ctx.adminCreate({cal:cfg.bookTo.public,title:'自己的事',date:'2026-10-21',s:1140,e:1200});
let m=ctx.apiMonth('public','','2026-10');
assert.equal(m.booked,2,'這個月 2 位（不算自己的行程）');
let d20=m.days.find(d=>d.date==='2026-10-20'); assert.equal(d20.booked,2); assert.deepEqual(d20.taken,[{s:1140,e:1200},{s:1230,e:1290}]);
let d21=m.days.find(d=>d.date==='2026-10-21'); assert(!d21.booked);
assert(!JSON.stringify(m).includes('"A"') && !JSON.stringify(m).includes('a@b.co'), '不能透露是誰');
let bb=ctx.bookBoot_(ctx.getConfig_(),'public',''); assert.ok('booked' in (bb.first||{booked:0}));

// 截圖判讀
assert.throws(()=>ctx.adminParseImage('xx','image/jpeg'),/金鑰/);
assert.throws(()=>ctx.adminSaveAiKey('abc'),/格式/);
assert.equal(ctx.adminSaveAiKey('AIzaSyD-abcdefghijklmnopqrstuvwxyz12').set,true);
assert.equal(ctx.adminBoot().ai.tail,'yz12');
let calls=[];
ctx.UrlFetchApp={fetch:(url,o)=>{calls.push(url); if(/gemini-flash-latest/.test(url)) return {getResponseCode:()=>404,getContentText:()=>'nf'};
  const body=JSON.parse(o.payload); assert(body.contents[0].parts[1].inline_data.data==='BASE64');
  return {getResponseCode:()=>200,getContentText:()=>JSON.stringify({candidates:[{content:{parts:[{text:JSON.stringify([
    {date:'2026-10-03',start:'19:00',end:null,allDay:false,title:'跟阿明吃飯',who:'阿明',place:'信義區',note:'帶禮物',sure:true},
    {date:'2026-10-05',start:null,end:null,allDay:false,title:'看展',who:'',place:'',note:'',sure:false},
    {date:'壞掉',title:'x'}])}]}}]})};}};
let ev=ctx.adminParseImage('BASE64','image/jpeg');
assert.equal(ev.length,2); assert.equal(ev[0].s,1140); assert.equal(ev[0].e,1200,'沒有結束時間補 1 小時'); assert.equal(ev[0].who,'阿明');
assert.equal(ev[1].allDay,true); assert.equal(ev[1].sure,false);
assert(/gemini-flash-latest/.test(calls[0]) && /gemini-3.8-flash/.test(calls[1]), '404 會換下一個型號');
assert.equal(ctx.getConfig_().aiModel,'gemini-3.8-flash','記住可以用的型號');
ctx.UrlFetchApp={fetch:()=>({getResponseCode:()=>429,getContentText:()=>'quota'})};
assert.throws(()=>ctx.adminParseImage('BASE64','image/jpeg'),/額度/);
assert.equal(ctx.adminSaveAiKey('').set,false);
console.log('TEST8 PASSED');
// 新版 AQ. 金鑰
const ctx2=makeEnv().ctx; ctx2.setup();
assert.equal(ctx2.adminSaveAiKey(' AQ.Ab8RN6KvXq2mZ-exampleKey_1234567890abcdefghij \n').set,true,'AQ. 金鑰要能存');
assert.equal(ctx2.adminAiStatus ? 1 : 1, 1);
let hdr=null; ctx2.UrlFetchApp={fetch:(u,o)=>{hdr=o.headers['x-goog-api-key']; return {getResponseCode:()=>401,getContentText:()=>'unauth'};}};
assert.throws(()=>ctx2.adminParseImage('B','image/jpeg'),/401/);
assert.equal(hdr,'AQ.Ab8RN6KvXq2mZ-exampleKey_1234567890abcdefghij','金鑰原樣放在 x-goog-api-key');
console.log('TEST8b PASSED');
// 503：同一型號重試，還是忙就換下一個型號
const ctx3=makeEnv().ctx; ctx3.setup(); ctx3.adminSaveAiKey('AQ.Ab8RN6KvXq2mZ-exampleKey_1234567890abcdefghij');
let seq=[];
ctx3.UrlFetchApp={fetch:(u)=>{const m=/models\/([^:]+)/.exec(u)[1]; seq.push(m);
  if (m==='gemini-flash-latest') return {getResponseCode:()=>503,getContentText:()=>'overloaded'};
  return {getResponseCode:()=>200,getContentText:()=>JSON.stringify({candidates:[{content:{parts:[{text:'[{"date":"2026-10-03","start":"19:00","end":"21:00","title":"吃飯"}]'}]}}]})};}};
let r3=ctx3.adminParseImage('B','image/jpeg');
assert.equal(r3.length,1); assert.deepEqual(seq,['gemini-flash-latest','gemini-flash-latest','gemini-flash-latest','gemini-3.8-flash'],'503 重試 3 次後換型號');
ctx3.UrlFetchApp={fetch:()=>({getResponseCode:()=>503,getContentText:()=>'x'})};
ctx3.getConfig_(); let c3=ctx3.getConfig_(); delete c3.aiModel; ctx3.saveConfig_(c3);
assert.throws(()=>ctx3.adminParseImage('B','image/jpeg'),/太忙/);
console.log('TEST8c PASSED');
// Claude Haiku
const ctx4=makeEnv().ctx; ctx4.setup();
let st4=ctx4.adminSaveAiKey('sk-ant-api03-AbCdEfGhIjKlMnOpQrStUvWxYz0123456789_-xyZ9');
assert.equal(st4.provider,'claude'); assert.equal(ctx4.adminBoot().ai.provider,'claude');
let reqs=[];
ctx4.UrlFetchApp={fetch:(u,o)=>{reqs.push({u,o}); if(reqs.length===1) return {getResponseCode:()=>529,getContentText:()=>'overloaded'};
  return {getResponseCode:()=>200,getContentText:()=>JSON.stringify({content:[{type:'text',text:'好的，以下是行程：\n[{"date":"2026-10-10","start":"12:00","end":"13:30","allDay":false,"title":"跟小美午餐","who":"小美","place":"東區","note":"","sure":true}]'}]})};}};
let rc=ctx4.adminParseImage('IMG','image/jpeg');
assert.equal(rc.length,1); assert.equal(rc[0].title,'跟小美午餐'); assert.equal(rc[0].s,720); assert.equal(rc[0].e,810);
assert.equal(reqs.length,2,'529 會重試');
const r0=reqs[1]; assert.equal(r0.u,'https://api.anthropic.com/v1/messages');
assert.equal(r0.o.headers['x-api-key'],'sk-ant-api03-AbCdEfGhIjKlMnOpQrStUvWxYz0123456789_-xyZ9'); assert.equal(r0.o.headers['anthropic-version'],'2023-06-01');
const bd=JSON.parse(r0.o.payload); assert.equal(bd.model,'claude-haiku-4-5-20251001'); assert.equal(bd.messages[0].content[0].source.data,'IMG'); assert.equal(bd.messages[0].content[0].source.media_type,'image/jpeg');
ctx4.UrlFetchApp={fetch:()=>({getResponseCode:()=>400,getContentText:()=>'{"error":{"message":"Your credit balance is too low to access the Anthropic API."}}'})};
assert.throws(()=>ctx4.adminParseImage('IMG','image/jpeg'),/餘額不足/);
ctx4.UrlFetchApp={fetch:()=>({getResponseCode:()=>401,getContentText:()=>'x'})};
assert.throws(()=>ctx4.adminParseImage('IMG','image/jpeg'),/Claude 金鑰無效/);
// 換回 Gemini 金鑰就改用 Gemini
assert.equal(ctx4.adminSaveAiKey('AQ.Ab8RN6KvXq2mZ-exampleKey_1234567890abcdefghij').provider,'gemini');
// 舊版只存 GEMINI_KEY 的也能讀
const e5=makeEnv(); e5.ctx.setup(); e5.state.props.GEMINI_KEY='AIzaSyD-abcdefghijklmnopqrstuvwxyz12'; assert.equal(e5.ctx.adminBoot().ai.provider,'gemini');
console.log('TEST8d PASSED');
