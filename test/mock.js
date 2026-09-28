// 模擬 Apps Script 環境，測後端邏輯
const fs=require('fs'), vm=require('vm');
function makeEnv(opts={}){
  const owner='me@example.com';
  const state={props:{}, cache:{}, sheets:{}, events:{}, mails:[], busy:[], activeUser: opts.activeUser ?? owner, cals:{}};
  function sheet(name){ const rows=[]; const s={name,rows,hidden:false,
    getLastRow:()=>rows.length, appendRow:r=>rows.push(r.slice()), setFrozenRows(){}, setColumnWidth(){}, hideSheet(){s.hidden=true}, deleteRow:i=>rows.splice(i-1,1),
    setName(n){s.name=n; state.sheets[n]=s; return s;},
    getRange(a,b,c,d){ if(typeof a==='string'){ return {getValue:()=> (rows[0]||[])[0]||'', setValue:v=>{rows[0]=[v]}} }
      const r=a,col=b,nr=c||1,nc=d||1;
      const rng={ setFontWeight:()=>rng,setBackground:()=>rng,setFontColor:()=>rng,
        getValues:()=>{const out=[];for(let i=0;i<nr;i++){out.push([]);for(let j=0;j<nc;j++)out[i].push((rows[r-1+i]||[])[col-1+j]??'')}return out;},
        getDisplayValues(){return rng.getValues().map(x=>x.map(String))},
        setValue:v=>{rows[r-1][col-1]=v}, setValues:vs=>{vs.forEach((row,i)=>row.forEach((v,j)=>{rows[r-1+i][col-1+j]=v}))} };
      return rng; } }; return s; }
  const ss={ getSheetByName:n=>state.sheets[n]||null, getSheets:()=>Object.values(state.sheets), insertSheet:n=>(state.sheets[n]=sheet(n)), getUrl:()=>'https://sheet', getId:()=>'ssid1'};
  state.sheets['工作表1']=sheet('工作表1');
  let evId=0;
  const ctx={
    console, Math, Date, JSON, String, Number, Object, Array, RegExp, Error, parseInt, parseFloat, isNaN,
    PropertiesService:{getScriptProperties:()=>({getProperty:k=>state.props[k]??null,setProperty:(k,v)=>{state.props[k]=v}})},
    CacheService:{getScriptCache:()=>({get:k=>state.cache[k]??null,put:(k,v)=>{state.cache[k]=v},remove:k=>delete state.cache[k]})},
    SpreadsheetApp:{create:()=>ss, openById:()=>ss},
    Session:{getActiveUser:()=>({getEmail:()=>state.activeUser}), getEffectiveUser:()=>({getEmail:()=>owner})},
    LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock(){}})},
    MailApp:{sendEmail:o=>state.mails.push(o)},
    Logger:{log:()=>{}},
    Utilities:{
      formatDate:(d,tz,f)=>{const p=Object.fromEntries(new Intl.DateTimeFormat('en-CA',{timeZone:tz,year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',hourCycle:'h23'}).formatToParts(d).map(x=>[x.type,x.value]));
        if(f==='yyyy-MM-dd')return `${p.year}-${p.month}-${p.day}`; if(f==='H')return String(+p.hour); if(f==='m')return String(+p.minute); if(f==='yyyy-MM-dd HH:mm')return `${p.year}-${p.month}-${p.day} ${p.hour}:${p.minute}`; throw f;},
      getUuid:()=>'uuid-'+Math.random().toString(36).slice(2,10)+'-xxxx'},
    CalendarApp:{getCalendarsByName:()=>[], createCalendar:(n,o)=>{const id='cal_'+n; state.cals[id]={n,o}; state.events[id]=[]; return {getId:()=>id}},
      getCalendarById:id=>({setName(n){state.cals[id].n=n},setColor(c){state.cals[id].o.color=c},deleteCalendar(){delete state.cals[id]}})},
    Calendar:{
      Freebusy:{query:q=>({calendars:Object.fromEntries(q.items.map(i=>[i.id,{busy:state.busy.filter(b=>b.cal===i.id||(!b.cal&&i.id===q.items[0].id)).map(b=>({start:b.start,end:b.end}))}]))})},
      Events:{
        insert:(ev,cal,opt)=>{const e=Object.assign({id:'ev'+(++evId),status:'confirmed'},ev); if(opt.conferenceDataVersion) e.hangoutLink='https://meet.google.com/abc-defg-hij'; (state.events[cal]=state.events[cal]||[]).push(e); state.lastInsert={ev:e,cal,opt}; return e;},
        list:(cal,o)=>({items:(state.events[cal]||[]).map(e=>Object.assign({},e,{start:e.start.dateTime?{dateTime:e.start.dateTime+'+08:00'}:e.start,end:e.end.dateTime?{dateTime:e.end.dateTime+'+08:00'}:e.end}))}),
        patch:(p,cal,id)=>{const e=state.events[cal].find(x=>x.id===id); Object.assign(e,p); return e;},
        move:(cal,id,dest)=>{const i=state.events[cal].findIndex(x=>x.id===id); const [e]=state.events[cal].splice(i,1); state.events[dest].push(e);},
        remove:(cal,id)=>{state.events[cal]=(state.events[cal]||[]).filter(x=>x.id!==id)},
        get:(cal,id)=>{const e=(state.events[cal]||[]).find(x=>x.id===id); if(!e) throw new Error('Not Found'); return e;}
      },
      CalendarList:{list:()=>({items:Object.keys(state.cals).map(id=>({id,summary:state.cals[id].n,accessRole:'owner'})).concat([{id:'family@group.calendar.google.com',summary:'家庭',accessRole:'writer'},{id:'me@example.com',summary:'me@example.com',accessRole:'owner',primary:true}])})}
      },
    ScriptApp:{getService:()=>({getUrl:()=>'https://script.google.com/macros/s/XYZ/exec'}),getProjectTriggers:()=>(state.triggers||[]).map(h=>({getHandlerFunction:()=>h})),deleteTrigger:t=>{state.triggers=(state.triggers||[]).filter(h=>h!==t.getHandlerFunction())},newTrigger:h=>{const b={timeBased:()=>b,everyDays:()=>b,atHour:()=>b,inTimezone:()=>b,create:()=>{(state.triggers=state.triggers||[]).push(h)}};return b;}},
    HtmlService:{}
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(__dirname+'/../apps-script/Code.gs','utf8'),ctx);
  return {ctx,state};
}
module.exports={makeEnv};
