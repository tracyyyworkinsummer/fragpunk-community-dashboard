const SHEET_ID = '1YLUD6y0Q7MfW7pKZjbkIlE6jmeN3uVteJ1D-oENnyF0';
const CACHE_KEY = 'fragpunk-dashboard-v3';
const MONTH_NAMES = {jan:1,january:1,feb:2,february:2,mar:3,march:3,apr:4,april:4,may:5,jun:6,june:6,jul:7,july:7,aug:8,august:8,sep:9,sept:9,september:9,oct:10,october:10,nov:11,november:11,dec:12,december:12};

function doGet(e) {
  const force = e && e.parameter && e.parameter.refresh === '1';
  const json = JSON.stringify(buildDashboard_(force));
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function buildDashboard_(force) {
  const cache = CacheService.getScriptCache();
  const cached = !force && cache.get(CACHE_KEY);
  if (cached) return JSON.parse(cached);
  const book = SpreadsheetApp.openById(SHEET_ID);
  const months = book.getSheets().map(sheet => {
    const period = monthFromTitle_(sheet.getName());
    return period ? parseMonth_(sheet, period) : null;
  }).filter(Boolean).sort((a,b) => a.key.localeCompare(b.key));
  const result = {source:book.getUrl(),syncedAt:new Date().toISOString(),months:months};
  const json = JSON.stringify(result);
  if (json.length < 95000) cache.put(CACHE_KEY, json, 300);
  return result;
}

function monthFromTitle_(title) {
  const low = String(title).toLowerCase();
  if (!low.includes('data') || /data update|only for|payment/.test(low)) return null;
  let month = 0;
  Object.keys(MONTH_NAMES).some(name => {
    if (new RegExp('\\b'+name+'\\b','i').test(low)) { month=MONTH_NAMES[name]; return true; }
    return false;
  });
  if (!month) return null;
  const match = low.match(/20\d{2}/), year = match ? Number(match[0]) : 2025;
  return {key:year+'-'+String(month).padStart(2,'0'),label:year+'年'+month+'月',year:year,month:month};
}

function parseMonth_(sheet, period) {
  const range=sheet.getDataRange(), rows=range.getValues(), shown=range.getDisplayValues(), rich=range.getRichTextValues();
  if (!rows.length) return null;
  const col=columns_(shown[0]), current={id:'',platform:'',region:'',channel:'',follower:0,kind:''};
  const creators={}, platforms={}, topics={}, weeks={}, contentRows=[], seen={};
  for (let r=1;r<rows.length;r++) {
    ['id','platform','region','channel','follower','kind'].forEach(name => {
      const c=col[name]; if (c==null || shown[r][c]==='') return;
      current[name]=name==='platform'?platform_(shown[r][c]):name==='follower'?num_(rows[r][c]):String(shown[r][c]).trim();
    });
    const id=current.id, platform=current.platform, link=link_(shown[r],rich[r],col.video);
    if (!id || !platform || !link || seen[link.toLowerCase()]) continue;
    const week=week_(col.week==null?'':shown[r][col.week]);
    if (!week.key) continue;
    seen[link.toLowerCase()]=true;
    const views=col.views==null?0:num_(rows[r][col.views]);
    const ccv=col.ccv==null?0:num_(rows[r][col.ccv]);
    let engagement=col.engagement==null?0:num_(rows[r][col.engagement]); if (engagement>1) engagement/=100;
    const salary=col.salary==null?0:num_(rows[r][col.salary]);
    const fragcoin=col.fragcoin==null?0:num_(rows[r][col.fragcoin]);
    const topic=String(col.topic==null?(current.kind||'Unknown'):(shown[r][col.topic]||current.kind||'Unknown')).trim();
    const keyword=col.keyword==null?null:bool_(shown[r][col.keyword]);
    const markedValid=col.valid==null?null:bool_(shown[r][col.valid]);
    const standard=standard_(platform,views,engagement), isTopic=keyword===true, isValid=markedValid==null?standard.valid:markedValid;

    if (!creators[id]) creators[id]=Object.assign({},current,{content:0,views:0,salary:0,fragcoin:0,ccvs:[]});
    const creator=creators[id]; creator.content++; creator.views+=views; creator.salary+=salary; creator.fragcoin+=fragcoin; if(ccv)creator.ccvs.push(ccv);
    add_(bucket_(platforms,platform),id,views,salary,fragcoin,ccv);
    if(!topics[topic])topics[topic]={content:0,views:0}; topics[topic].content++; topics[topic].views+=views;
    if(!weeks[week.key])weeks[week.key]={};
    const weekly=bucket_(weeks[week.key],platform,true); add_(weekly,id,views,salary,fragcoin,ccv);
    if(isTopic)weekly.topicContent++;
    if(isTopic&&isValid)weekly.validTopic++;
    if(isTopic&&!isValid&&(platform==='TikTok'||platform==='YouTube')){
      if(standard.reason.includes('曝光'))weekly.invalidExposure++;
      if(standard.reason.includes('互动'))weekly.invalidEngagement++;
    }
    if(views>10000||ccv>10000)weekly.hotLinks.push({link:link,views:Math.round(Math.max(views,ccv)),creator:id,topic:topic});
    contentRows.push({creator:id,platform:platform,region:current.region||'-',week:week.key,weekNo:week.number,weekRaw:week.raw,topic:topic,views:Math.round(views),ccv:round_(ccv,1),engagement:round_(engagement,4),validTopic:Boolean(isValid),invalidReason:standard.reason,salary:round_(salary,2),fragcoin:round_(fragcoin,2),link:link});
  }

  const service=service_(shown,rows), platformValues=Object.values(platforms);
  const month={key:period.key,label:period.label,year:period.year,quarter:Math.floor((period.month-1)/3)+1,months:1,
    creators:Object.keys(creators).length,content:contentRows.length,views:Math.round(sum_(platformValues,'views')),
    salary:round_(sum_(platformValues,'salary'),2),fragcoin:round_(sum_(platformValues,'fragcoin'),2),
    serviceSpend:service.spend,serviceCpm:service.cpm,avgCcv:avgCcv_(platformValues),platforms:{},topics:{},weeks:{},
    topCreators:Object.keys(creators).map(id=>creator_(id,creators[id])).sort((a,b)=>b.views-a.views).slice(0,80),contentRows:contentRows};
  month.cpm=month.views?round_(month.salary/month.views*1000,2):0;
  Object.keys(platforms).forEach(name=>month.platforms[name]=output_(platforms[name]));
  Object.keys(topics).sort((a,b)=>topics[b].views-topics[a].views).forEach(name=>month.topics[name]={content:topics[name].content,views:Math.round(topics[name].views)});
  Object.keys(weeks).forEach(wk=>{month.weeks[wk]={};Object.keys(weeks[wk]).forEach(name=>{
    const src=weeks[wk][name], out=output_(src); out.activeIds=Object.keys(src.activeIds).sort();
    ['topicContent','validTopic','invalidExposure','invalidEngagement'].forEach(k=>out[k]=src[k]);
    out.hotLinks=src.hotLinks.sort((a,b)=>b.views-a.views); month.weeks[wk][name]=out;
  });});
  return month;
}

function columns_(headers){const week=pick_(headers,['Week']),topic=pick_(headers,['Topic']);let video=pick_(headers,['Video link','Posting Link']);if(video==null&&week!=null&&topic!=null&&topic>week){for(let i=week+1;i<topic;i++){if(!norm_(headers[i])){video=i;break;}}}return{id:pick_(headers,['KOC DC ID']),channel:pick_(headers,['Channel Link']),platform:pick_(headers,['Main Platform']),region:pick_(headers,['受众地区','主要受众','受众市场','Region']),kind:pick_(headers,['Video/Stream']),follower:pick_(headers,['Follower']),week:week,video:video,topic:topic,views:pick_(headers,['Views','Views/UV']),ccv:pick_(headers,['CCV']),engagement:pick_(headers,['Engagement Rate','互动率']),keyword:pick_(headers,['是否包含专项关键词']),valid:pick_(headers,['是否为有效专项']),salary:pick_(headers,['Salary','Total Salary']),fragcoin:pick_(headers,['FragCoin'])};}
function norm_(v){return String(v==null?'':v).replace(/\s+/g,' ').trim().toLowerCase();}
function pick_(headers,names){const h=headers.map(norm_);for(const n of names){const i=h.indexOf(norm_(n));if(i>=0)return i;}for(const n of names){const t=norm_(n),i=h.findIndex(v=>t&&v.includes(t));if(i>=0)return i;}return null;}
function num_(v){if(typeof v==='number')return isFinite(v)?v:0;const n=Number(String(v==null?'':v).replace(/[$,%\s]/g,'').replace(/,/g,''));return isFinite(n)?n:0;}
function platform_(v){const s=norm_(v);if(s.includes('tiktok')||s==='tt')return'TikTok';if(s.includes('twitch'))return'Twitch';if(s.includes('youtube')||s==='yt'||s==='ytb')return'YouTube';return String(v||'').trim()||'Unknown';}
function bool_(v){const s=norm_(v);if(!s)return null;if(['是','yes','valid','有效','合格','达标','true','1'].includes(s))return true;if(['否','no','invalid','无效','不合格','未达标','false','0'].includes(s))return false;if(/不计专项|非专项|不包含/.test(s))return false;if(/计专项|包含|有效/.test(s))return true;return null;}
function week_(v){const raw=String(v||'').trim(),m=raw.match(/weeks?\s*(\d+)/i)||raw.match(/第?\s*(\d+)\s*周/),n=m?Number(m[1]):0;return{key:n?'W'+n:'',number:n,raw:raw};}
function link_(shown,rich,i){if(i==null)return'';const cell=rich[i],url=cell&&cell.getLinkUrl?cell.getLinkUrl():'';return String(url||shown[i]||'').trim();}
function standard_(platform,views,engagement){const reasons=[];if(platform==='TikTok'){if(views<=500)reasons.push('曝光未达标');if(engagement<=.05)reasons.push('互动未达标');}else if(platform==='YouTube'&&views<=100)reasons.push('曝光未达标');return{valid:!reasons.length,reason:reasons.join('、')};}
function bucket_(where,name,weekly){if(!where[name]){where[name]={content:0,views:0,salary:0,fragcoin:0,activeIds:{},ccvSum:0,ccvN:0};if(weekly)Object.assign(where[name],{topicContent:0,validTopic:0,invalidExposure:0,invalidEngagement:0,hotLinks:[]});}return where[name];}
function add_(b,id,views,salary,fragcoin,ccv){b.content++;b.views+=views;b.salary+=salary;b.fragcoin+=fragcoin;b.activeIds[id]=true;if(ccv){b.ccvSum+=ccv;b.ccvN++;}}
function output_(b){return{content:b.content,views:Math.round(b.views),salary:round_(b.salary,2),fragcoin:round_(b.fragcoin,2),active:Object.keys(b.activeIds).length,avgCcv:b.ccvN?round_(b.ccvSum/b.ccvN,1):0};}
function creator_(id,c){return{name:id,platform:c.platform,region:c.region||'-',follower:c.follower,views:Math.round(c.views),salary:round_(c.salary,2),fragcoin:round_(c.fragcoin,2),content:c.content,avgCcv:c.ccvs.length?round_(c.ccvs.reduce((a,b)=>a+b,0)/c.ccvs.length,1):0,channel:c.channel};}
function avgCcv_(items){const total=items.reduce((s,x)=>s+x.ccvSum,0),count=items.reduce((s,x)=>s+x.ccvN,0);return count?round_(total/count,1):0;}
function service_(shown,values){let spend=0,cpm=0;for(let r=0;r<shown.length;r++)for(let c=0;c<shown[r].length;c++){const text=String(shown[r][c]||'').trim();if(!text)continue;if((text.includes('总计消耗')&&text.includes('含服务费')&&!text.includes('不含'))||text==='总消耗'||text==='总消耗（含服务费）'){if(!spend)spend=near_(values,r,c);}else if(text.toUpperCase()==='CPM'&&!cpm)cpm=near_(values,r,c);}return{spend:round_(spend,2),cpm:round_(cpm,2)};}
function near_(rows,r,c){for(const p of [[r+1,c],[r,c+1],[r+1,c+1],[r+1,c-1]])if(p[0]>=0&&p[0]<rows.length&&p[1]>=0&&p[1]<rows[p[0]].length){const n=num_(rows[p[0]][p[1]]);if(n)return n;}return 0;}
function sum_(items,key){return items.reduce((s,x)=>s+Number(x[key]||0),0);}
function round_(n,d){const p=Math.pow(10,d||0);return Math.round((Number(n)||0)*p)/p;}
