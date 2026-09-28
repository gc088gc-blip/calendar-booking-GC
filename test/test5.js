const {makeEnv}=require('./mock'); const assert=require('assert');
const {ctx,state}=makeEnv(); ctx.setup();
const cfg=ctx.getConfig_(), C=cfg.calendars.map(c=>c.id);
ctx.adminCreate({cal:C[0],title:'吃飯',date:'2026-09-20',s:1080,e:1200,note:'訂位\n帶禮物'});
let ev=state.lastInsert.ev; assert.equal(ev.description,'訂位\n帶禮物'); assert.equal(ev.extendedProperties.private.note,'訂位\n帶禮物');
let bd=ctx.adminBoard(); assert.equal(bd.length,8); assert.equal(bd[0].t,'');
ctx.adminBoardSave([{t:'=SUM(1)'},{t:'第二件',d:true}]);
bd=ctx.adminBoard(); assert.equal(bd.length,8); assert.equal(bd[0].t,'=SUM(1)'); assert.equal(bd[1].d,true); assert.equal(ctx.adminBoot().board.length,8);
// 舊格式：一整塊文字 → 拆成格子；超過 8 行併進最後一格
state.sheets['_白板'].getRange('A1').setValue(JSON.stringify({t:'a\nb\n\nc\nd\ne\nf\ng\nh\ni\nj'}));
bd=ctx.adminBoard(); assert.equal(bd.length,8); assert.equal(bd[0].t,'a'); assert.equal(bd[7].t,'h / i / j');
// 更舊的便條格式
state.sheets['_白板'].getRange('A1').setValue(JSON.stringify([{t:'舊1',done:true},{t:'舊2'}]));
bd=ctx.adminBoard(); assert.equal(bd[0].t,'舊1'); assert.equal(bd[0].d,true); assert.equal(bd[2].t,'');
// 多行事曆
const CC=ctx.getConfig_().calendars.map(c=>c.id);
ctx.adminCreate({cal:CC[0],tags:[CC[2],CC[0],'bogus',CC[2]],title:'共同',date:'2026-09-21',s:600,e:660});
let ev2=state.lastInsert.ev; assert.equal(ev2.extendedProperties.private.tags, CC[2]);
let list=ctx.adminEvents('2026-09-21','2026-09-21'); let e=list.filter(x=>x.title==='共同')[0]; assert.deepEqual(e.tags,[CC[2]]);
console.log('TEST5 PASSED');
