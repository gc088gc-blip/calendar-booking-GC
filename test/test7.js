const {makeEnv}=require('./mock'); const assert=require('assert');
const {ctx,state}=makeEnv(); ctx.setup();
const C=ctx.getConfig_().calendars.map(c=>c.id), FAM='family@group.calendar.google.com';
// 沒設定前勾家人 → 錯誤
assert.throws(()=>ctx.adminCreate({cal:C[0],title:'x',date:'2026-10-03',s:600,e:660,family:true}),/家庭日曆/);
let opts=ctx.adminFamilyOptions(); assert.equal(opts.guess,FAM); assert(opts.items.every(i=>C.indexOf(i.id)<0 && i.id!=='me@example.com'));
let cfg=ctx.adminSaveFamilyCal(FAM); assert.equal(cfg.familyName,'家庭');
// 新增：家庭日曆多一份，只有標題時間地點
ctx.adminCreate({cal:C[0],title:'回家吃飯',date:'2026-10-03',s:1080,e:1200,loc:'老家',who:'家人',note:'帶水果',family:true});
assert.equal(state.events[FAM].length,1); let fe=state.events[FAM][0];
assert.equal(fe.summary,'回家吃飯'); assert.equal(fe.location,'老家'); assert(!fe.extendedProperties, '家庭那份不帶對象備註');
let main=state.events[C[0]].find(e=>e.summary==='回家吃飯'); assert.equal(main.extendedProperties.private.famId,fe.id);
// 列表帶 famId
let ev=ctx.adminEvents('2026-10-03','2026-10-03').find(e=>e.title==='回家吃飯'); assert.equal(ev.famId,fe.id);
// 改時間 → 家庭那份跟著改
ctx.adminUpdate({cal:C[0],id:main.id,title:'回家吃晚餐',date:'2026-10-03',s:1110,e:1230,loc:'老家',family:true,famId:fe.id});
assert.equal(state.events[FAM][0].summary,'回家吃晚餐'); assert(/18:30/.test(state.events[FAM][0].start.dateTime));
// 關掉 → 家庭那份刪掉
ctx.adminUpdate({cal:C[0],id:main.id,title:'回家吃晚餐',date:'2026-10-03',s:1110,e:1230,family:false,famId:fe.id});
assert.equal(state.events[FAM].length,0); assert.equal(main.extendedProperties.private.famId,'');
// 再打開 → 重新建一份
ctx.adminUpdate({cal:C[0],id:main.id,title:'回家吃晚餐',date:'2026-10-03',s:1110,e:1230,family:true,famId:''});
assert.equal(state.events[FAM].length,1);
// 刪除主行程 → 家庭那份一起刪
ctx.adminDelete(C[0],main.id,'',false);
assert.equal(state.events[FAM].length,0);
// 家庭日曆不算忙碌
assert(ctx.getConfig_().calendars.every(c=>c.id!==FAM));
console.log('TEST7 PASSED');
