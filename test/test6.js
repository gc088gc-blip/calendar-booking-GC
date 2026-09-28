const {makeEnv}=require('./mock'); const assert=require('assert');
const {ctx,state}=makeEnv(); ctx.setup();
const now=new Date('2026-09-22T10:00:00+08:00').getTime(); Date.now=()=>now;
ctx.Calendar.Freebusy.query=q=>{const cal={};q.items.forEach(i=>{cal[i.id]={busy:(state.events[i.id]||[]).filter(e=>e.start.dateTime).map(e=>({start:e.start.dateTime+'+08:00',end:e.end.dateTime+'+08:00'}))}});return {calendars:cal};};
let cfg=ctx.getConfig_(); cfg.publicUrl='https://script.google.com/macros/s/X/exec'; ctx.saveConfig_(cfg);
ctx.adminSetOpen(['2026-10-18'],[{s:'19:00',e:'22:00',loc:'both',aud:'all'}]);
// 新朋友：待審核
let r=ctx.apiBook({aud:'public',k:'',date:'2026-10-18',start:1140,end:1200,name:'Karen',contact:'karen@example.com',purpose:'聊轉職',locMode:'online'});
assert.equal(r.pending,true); assert.equal(r.meet,'');
let ib=ctx.adminInbox(); assert.equal(ib[0].status,'待審核'); assert.equal(ib[0].aud,'新朋友');
let evs=state.events[ctx.getConfig_().bookTo.public]; assert(/^【待審核】Karen/.test(evs[0].summary), evs[0].summary);
// 該時段已被卡住
assert.throws(()=>ctx.apiBook({aud:'public',k:'',date:'2026-10-18',start:1140,end:1200,name:'B',contact:'b@b.co',purpose:'x',locMode:'online'}),/約走/);
// 朋友：直接成立
let fr=ctx.apiBook({aud:'friend',k:cfg.audiences.friend.token,date:'2026-10-18',start:1230,end:1290,name:'小明',contact:'',purpose:'吃飯',locMode:'online'});
assert.equal(!!fr.pending,false); assert.equal(ctx.adminInbox()[0].status,'已排');
// 通過
const before=state.mails.length;
let l=ctx.adminApprove(ib[0].id);
assert.equal(l.filter(x=>x.name==='Karen')[0].status,'已排');
assert(/^【新朋友】Karen/.test(evs[0].summary), evs[0].summary);
assert(state.mails.length>before, '應該寄確認信給對方');
assert.equal(state.mails[state.mails.length-1].to,'karen@example.com');
// 婉拒
let r2=ctx.apiBook({aud:'public',k:'',date:'2026-10-18',start:1290,end:1320,name:'林',contact:'lin@example.com',purpose:'問實習',locMode:'online'});
let id2=ctx.adminInbox()[0].id; const n2=state.events[ctx.getConfig_().bookTo.public].length;
ctx.adminReject(id2,'那天剛好有事');
assert.equal(ctx.adminInbox()[0].status,'婉拒');
assert.equal(state.events[ctx.getConfig_().bookTo.public].filter(e=>e.status!=='cancelled').length, n2-1, '婉拒要刪掉行程');
assert(/那天剛好有事/.test(state.mails[state.mails.length-1].body));
// 關掉審核
ctx.adminSaveReview('public',false);
let r3=ctx.apiBook({aud:'public',k:'',date:'2026-10-18',start:1290,end:1320,name:'阿明',contact:'a@b.co',purpose:'x',locMode:'online'});
assert.equal(!!r3.pending,false);
console.log('TEST6 PASSED');
// 每月的話：朋友和新朋友分開
ctx.adminSaveNote('2026-09',{f:'這個月大多在台北',p:'這個月只接線上'});
ctx.adminSaveNote('2026-10',{f:'',p:'10 月只接線上'});
assert.equal(ctx.bookBoot_(ctx.getConfig_(),'friend','t').notes['2026-09'],'這個月大多在台北');
assert.equal(ctx.bookBoot_(ctx.getConfig_(),'public','').notes['2026-09'],'這個月只接線上');
assert.equal(ctx.bookBoot_(ctx.getConfig_(),'friend','t').notes['2026-10'],undefined);
assert.equal(ctx.bookBoot_(ctx.getConfig_(),'public','').notes['2026-10'],'10 月只接線上');
// 舊格式自動轉換
let c=ctx.getConfig_(); c.notes['2026-11']={t:'舊的話',aud:'friend'}; ctx.saveConfig_(c);
ctx.CacheService.getScriptCache().remove('cfg');
assert.equal(ctx.bookBoot_(ctx.getConfig_(),'friend','t').notes['2026-11'],'舊的話');
assert.equal(ctx.bookBoot_(ctx.getConfig_(),'public','').notes['2026-11'],undefined);
ctx.adminSaveNote('2026-09',{f:'',p:''});
assert.equal(ctx.bookBoot_(ctx.getConfig_(),'friend','t').notes['2026-09'],undefined);
console.log('TEST6b PASSED');
