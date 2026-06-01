import React, { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

const STORAGE_KEY = "csm_dash_v4";
const DASHBOARD_PIN = "thryv2025";
const PIN_KEY = "csm_pin_unlocked";
const SHEET_ID = "1ZjakNUSbcSaiO3KVXcg6q3FetFbLISaFsGKLNtVOhyI";

// Tab GIDs - update these after creating tabs in your sheet
// Sheet1/revenue = gid 0, create "email" and "cadence" tabs and put their gids here
const GID_REV     = "322916467";
const GID_EMAIL   = "0";
const GID_CADENCE = "1973544046";

function sheetCSVUrl(gid) {
  return `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&gid=${gid}`;
}

async function fetchSheetTab(gid) {
  // Try multiple URL formats - Google sometimes blocks one but not another
  const urls = [
    `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&gid=${gid}`,
    `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${gid}`,
    `https://opensheet.elk.sh/${SHEET_ID}/${gid}`,
  ];
  for(const url of urls) {
    try {
      const res = await fetch(url);
      if(!res.ok) continue;
      const text = await res.text();
      if(!text || text.trim()==="" || text.includes("<!DOCTYPE")) continue;
      // Try JSON format first (opensheet returns JSON)
      if(text.trim().startsWith("[")) {
        try { return JSON.parse(text); } catch(e) {}
      }
      // Parse CSV
      const lines = text.split("\n").filter(l=>l.trim());
      if(lines.length<2) continue;
      const headers = parseCSVLine(lines[0]);
      const rows = lines.slice(1).map(line=>{
        const vals = parseCSVLine(line);
        const obj={};
        headers.forEach((h,i)=>obj[h.trim()]=vals[i]||"");
        return obj;
      }).filter(r=>Object.values(r).some(v=>v));
      if(rows.length>0) return rows;
    } catch(e) { continue; }
  }
  return [];
}

function parseCSVLine(line) {
  const result=[], re=/("(?:[^"]|"")*"|[^,]*)(,|$)/g;
  let m;
  while((m=re.exec(line))!==null) {
    let val=m[1];
    if(val.startsWith('"')&&val.endsWith('"')) val=val.slice(1,-1).replace(/""/g,'"');
    result.push(val);
    if(m[2]==="") break;
  }
  return result;
}

async function pullFromSheets() {
  const [revRows, emailRows, cadRows] = await Promise.all([
    fetchSheetTab(GID_REV),
    fetchSheetTab(GID_EMAIL),
    fetchSheetTab(GID_CADENCE),
  ]);
  return { revRows, emailRows, cadRows };
}

// Map Google Sheet rows → dashboard format
function sheetToRev(rows) {
  return rows.map(r=>{
    // Handle both raw JotForm format and pre-processed format
    const name = r["CSM Name"]||r["name"]||"";
    if(!name) return null;
    return {
      name: name.trim(),
      team: (r["CSM Team! "]||r["CSM Team"]||r["team"]||"").trim(),
      tier: (r["CSM Tier"]||r["tier"]||"").trim(),
      type: (r["Type of Integration"]||r["type"]||"").trim(),
      quarter: (r["Quarter for Consideration"]||r["quarter"]||"").trim(),
      mrr:  +((r["MRR $ Added"]||r["mrr"]||"0").replace(/[$,]/g,"")||0),
      otr:  +((r["OTR $ Added"]||r["otr"]||"0").replace(/[$,]/g,"")||0),
      total:+((r["Total Revenue Added"]||r["total"]||"0").replace(/[$,]/g,"")||0),
      nonrev: (r["Non-Revenue Integrations"]||r["nonrev"]||"").trim()?1:0,
    };
  }).filter(Boolean);
}

function sheetToEmail(rows) {
  return rows.map(r=>{
    const name=r["name"]||r["Name"]||r["CSM"]||"";
    if(!name) return null;
    return {
      name:name.trim(),
      sent:   +(r["sent"]||r["Sent"]||0),
      opens:  +(r["opens"]||r["Opens"]||0),
      replies:+(r["replies"]||r["Replies"]||0),
      openRate: +(r["openRate"]||r["Open Rate"]||r["open_rate"]||0),
      replyRate:+(r["replyRate"]||r["Reply Rate"]||r["reply_rate"]||0),
      replyWhenOpened:+(r["replyWhenOpened"]||0),
    };
  }).filter(Boolean);
}

function sheetToCadence(rows) {
  return rows.map(r=>{
    const name=r["name"]||r["Name"]||r["CSM"]||"";
    if(!name) return null;
    return {
      name:name.trim(),
      count:  +(r["count"]||r["Count"]||0),
      pct:    +(r["pct"]||r["Pct"]||r["Completion"]||0),
      removed:+(r["removed"]||r["Removed"]||0),
    };
  }).filter(Boolean);
}

function storageSave(data) {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(data)); } catch(e) {}
}
function storageLoad() {
  try { const r=localStorage.getItem(STORAGE_KEY); return r?JSON.parse(r):null; } catch(e) { return null; }
}
function exportSnap(loaded) {
  const blob=new Blob([JSON.stringify({v:"csm1",loaded,at:new Date().toISOString()})],{type:"application/json"});
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob);
  a.download="csm_snapshot_"+new Date().toLocaleDateString("en-US").replace(/\//g,"-")+".json";
  a.click(); URL.revokeObjectURL(a.href);
}

function parseRev(wb) {
  const ws = wb.Sheets[wb.SheetNames.includes("Sheet1")?"Sheet1":wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws,{defval:""});
  if (rows[0] && rows[0]["CSM Name"]!==undefined) {
    return rows.map(r=>({
      name:String(r["CSM Name"]||"").trim(),
      team:String(r["CSM Team! "]||r["CSM Team"]||r["Team"]||"").trim(),
      tier:String(r["CSM Tier"]||"").trim(),
      mrr:pm(r["MRR $ Added"]),otr:pm(r["OTR $ Added"]),
      total:pm(r["Total Revenue Added"]),nonrev:String(r["Non-Revenue Integrations"]||"").trim()?1:0,
    })).filter(r=>r.name);
  }
  const raw = XLSX.utils.sheet_to_json(ws,{defval:"",header:1});
  const out=[]; let team="";
  raw.forEach(r=>{
    const lab=String(r[0]||"").trim();
    if(!lab||lab==="Row Labels"||lab==="Grand Total") return;
    if(lab.includes("@")) {
      const nm=lab.split("@")[0].split(".").map(p=>p[0].toUpperCase()+p.slice(1)).join(" ");
      out.push({name:nm,team,tier:"",mrr:pm(r[2]),otr:pm(r[1]),total:pm(r[3]),nonrev:pn(r[4])});
    } else { team=lab; }
  });
  return out;
}
function parseEmail(wb) {
  const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:"",header:1});
  let hdr=-1;
  raw.forEach((r,i)=>{ if(String(r[1]||"").includes("Owner Name")) hdr=i; });
  if(hdr<0) return [];
  const out=[];
  for(let i=hdr+1;i<raw.length;i++){
    const r=raw[i], name=String(r[1]||"").trim();
    if(!name||name==="Total"||name.includes("©")) break;
    const sent=pn(r[2]);
    if(sent>0||pn(r[11])>0) out.push({name,sent,opens:pn(r[3]),replies:pn(r[4]),openRate:pn(r[8]),replyRate:pn(r[9])});
  }
  return out;
}
function parseCadence(wb) {
  const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:"",header:1});
  const out=[]; let cur="",rem=0;
  raw.forEach(r=>{
    const c1=String(r[1]||"").trim(),c2=String(r[2]||"").trim(),c3=String(r[3]||"").trim();
    if(c1&&c1!=="Subtotal"&&c2&&!c2.includes("Subtotal")&&(c2.startsWith("CSM")||c2.includes("Cadence"))){cur=c1;rem=0;}
    if(cur&&String(r[5]||"").trim()==="Removed") rem++;
    if(c1==="Subtotal"&&!c2&&!c3){
      const count=pn(r[5]),pct=pn(r[13]);
      if(cur&&count>0){out.push({name:cur,count,pct,removed:rem});cur="";rem=0;}
    }
  });
  return out.filter(r=>r.name);
}

function buildCSMs(data) {
  const m={};
  const get=(name,team,tier)=>{
    if(!m[name]){
      const info=lk(name);
      m[name]={name,team:team||(info&&info.t)||"",coach:(info&&info.c)||"",tier:tier||(info&&info.r)||"",
        rev:0,mrr:0,otr:0,ints:0,nonrev:0,sent:0,opens:0,replies:0,openRate:0,replyRate:0,cadCount:0,cadPct:0};
    }
    if(team&&!m[name].team) m[name].team=team;
    if(!m[name].coach){const info=lk(name);if(info)m[name].coach=info.c;}
    return m[name];
  };
  if(data.rev){
    const byN={};
    data.rev.forEach(r=>{
      if(!r.name) return;
      const info=lk(r.name);
      if(!byN[r.name]) byN[r.name]={team:r.team||(info&&info.t)||"",tier:r.tier||(info&&info.r)||"",coach:(info&&info.c)||"",rev:0,mrr:0,otr:0,ints:0,nonrev:0};
      byN[r.name].rev+=r.total; byN[r.name].mrr+=r.mrr; byN[r.name].otr+=r.otr;
      if(r.total>0) byN[r.name].ints++;
      byN[r.name].nonrev+=r.nonrev;
    });
    Object.entries(byN).forEach(([name,v])=>{
      const c=get(name,v.team,v.tier);
      c.coach=v.coach||c.coach; c.rev=v.rev; c.mrr=v.mrr; c.otr=v.otr; c.ints=v.ints+v.nonrev; c.nonrev=v.nonrev;
    });
  }
  if(data.email) data.email.forEach(r=>{
    const c=get(r.name,"",""); c.sent=r.sent; c.opens=r.opens; c.replies=r.replies; c.openRate=r.openRate; c.replyRate=r.replyRate;
  });
  if(data.cadence) data.cadence.forEach(r=>{
    const c=get(r.name,"",""); c.cadCount=r.count; c.cadPct=r.pct;
  });
  return Object.values(m);
}

function Bar({label,val,hi,lo}) {
  const pct = val!=null?Math.min(val*100,100):0;
  const col = val!=null?bCol(val,hi,lo):"#e5e7eb";
  return (
    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
      <span style={{width:70,flexShrink:0,fontSize:10,fontWeight:700,color:MUT,textTransform:"uppercase",letterSpacing:".04em"}}>{label}</span>
      <div style={{flex:1,height:6,background:"#ECEEF1",borderRadius:3,overflow:"hidden"}}>
        <div style={{width:pct.toFixed(1)+"%",height:"100%",background:col,borderRadius:3}}/>
      </div>
      {val!=null
        ? <span style={{width:30,textAlign:"right",fontSize:11,fontWeight:700,color:col,flexShrink:0}}>{Math.round(val*100)+"%"}</span>
        : <span style={{width:30,textAlign:"right",fontSize:11,color:MUT,flexShrink:0}}>—</span>}
    </div>
  );
}

function Bdg({p,label}) {
  const bg=p>=0.9?"rgba(22,163,74,.12)":p>=0.5?"rgba(217,119,6,.12)":p>0?"rgba(220,38,38,.12)":"rgba(128,128,128,.1)";
  const fg=p>=0.9?"#166534":p>=0.5?"#854d0e":p>0?"#991b1b":"#808080";
  const txt=label||(p>=0.9?"Win":p>=0.5?"Watch":p>0?"Coach":"—");
  return <span style={{fontSize:10,fontWeight:700,padding:"1px 7px",borderRadius:20,background:bg,color:fg,flexShrink:0}}>{txt}</span>;
}

function CoachCard({coach,csms}) {
  const team = csms.filter(c=>{ const i=lk(c.name); return (i&&i.c===coach.email)||c.coach===coach.email; });
  const cadCSMs = team.filter(c=>c.cadCount>0);
  const avgCad = cadCSMs.length ? cadCSMs.reduce((s,c)=>s+c.cadPct,0)/cadCSMs.length : null;
  const emailCSMs = team.filter(c=>c.sent>0);
  const avgOpen = emailCSMs.length ? emailCSMs.reduce((s,c)=>s+c.openRate,0)/emailCSMs.length : null;
  const revPct = team.length ? team.filter(c=>c.rev>0||c.ints>0).length/team.length : null;
  let sc=0,sf=0;
  if(avgCad!=null){sc+=Math.min(avgCad/0.9,1)*40;sf++;}
  if(avgOpen!=null){sc+=Math.min(avgOpen/0.7,1)*30;sf++;}
  if(revPct!=null){sc+=Math.min(revPct/0.7,1)*30;sf++;}
  const score = sf?Math.round(Math.min(sc,100)):null;
  const status = score==null?"none":score>=75?"win":score>=50?"warn":"att";
  const sCol = status==="win"?GRN:status==="warn"?YLW:status==="att"?RED:MUT;
  const topC = status==="win"?GRN:status==="warn"?YLW:status==="att"?RED:"#e5e7eb";
  const wins=cadCSMs.filter(c=>c.cadPct>=0.9).length;
  const warns=cadCSMs.filter(c=>c.cadPct>=0.5&&c.cadPct<0.9).length;
  const atts=cadCSMs.filter(c=>c.cadPct>0&&c.cadPct<0.5).length;
  const sorted=[...team].sort((a,b)=>(a.cadPct||0)-(b.cadPct||0)).slice(0,7);
  return (
    <div style={{background:"#fff",border:"1px solid rgba(41,53,93,.09)",borderRadius:12,padding:18,boxShadow:"0 2px 10px rgba(41,53,93,.06)",position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:0,left:0,right:0,height:3,background:topC,borderRadius:"12px 12px 0 0"}}/>
      <div style={{fontFamily:"Nunito,sans-serif",fontSize:14,fontWeight:800,color:MID,marginTop:4}}>{coach.name}</div>
      <div style={{fontSize:11,fontWeight:600,color:tc(coach.team),marginBottom:10}}>● {coach.team}</div>
      <div style={{display:"flex",alignItems:"baseline",gap:6,marginBottom:8}}>
        <span style={{fontFamily:"Nunito,sans-serif",fontSize:44,fontWeight:800,lineHeight:1,letterSpacing:"-.03em",color:sCol}}>{score!=null?score:"—"}</span>
        <div style={{fontSize:11,color:MUT,paddingBottom:6,lineHeight:1.4}}>/ 100<br/><span style={{fontSize:10}}>{wins}✓ {warns}⚠ {atts}✗</span></div>
      </div>
      <Bar label="Cadence" val={avgCad} hi={0.9} lo={0.5}/>
      <Bar label="Email" val={avgOpen} hi={0.7} lo={0.35}/>
      <Bar label="Revenue" val={revPct} hi={0.7} lo={0.4}/>
      <div style={{height:1,background:"rgba(41,53,93,.07)",margin:"10px 0"}}/>
      <div style={{fontSize:10,textTransform:"uppercase",letterSpacing:".08em",color:MUT,fontWeight:700,marginBottom:8}}>CSM Status</div>
      {sorted.map(c=>(
        <div key={c.name} style={{display:"flex",alignItems:"center",gap:6,padding:"4px 0",borderBottom:"1px solid rgba(41,53,93,.05)"}}>
          <span style={{flex:1,fontSize:11,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.name}</span>
          <span style={{fontSize:11,fontWeight:700,width:30,textAlign:"right",flexShrink:0,color:pCol(c.cadPct)}}>{c.cadPct>0?Math.round(c.cadPct*100)+"%":"—"}</span>
          <Bdg p={c.cadPct}/>
        </div>
      ))}
      {team.length>7&&<div style={{fontSize:11,color:MUT,textAlign:"center",marginTop:8}}>+{team.length-7} more</div>}
    </div>
  );
}

function CoachingView({csms,fc,ft}) {
  const coaches = COACHES.filter(c=>(!fc||c.email===fc)&&(!ft||c.team===ft));
  const cols = coaches.length===1?1:coaches.length===2?2:3;
  const tips = p => p===0?"No touchpoints — check cadence is active":p<0.5?"Significantly behind — prioritize in 1:1":p<0.7?"Below average — discuss blockers":"Close to target — a nudge may close the gap";
  const attn = csms.filter(c=>{ const i=lk(c.name); if(!i) return false; if(fc&&i.c!==fc) return false; if(ft&&i.t!==ft) return false; return c.cadPct<0.9&&(c.cadCount>0||c.sent>0); }).sort((a,b)=>(a.cadPct||0)-(b.cadPct||0));
  const wins = csms.filter(c=>{ const i=lk(c.name); if(!i) return false; if(fc&&i.c!==fc) return false; if(ft&&i.t!==ft) return false; return c.cadPct>=0.9&&c.cadCount>0; }).sort((a,b)=>b.cadPct-a.cadPct);
  return (
    <div style={{padding:"24px 28px"}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat("+cols+",1fr)",gap:18,marginBottom:32}}>
        {coaches.map(c=><CoachCard key={c.email} coach={c} csms={csms}/>)}
      </div>
      <div style={{marginBottom:28}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
          <span style={{fontFamily:"Nunito,sans-serif",fontSize:18,fontWeight:800,color:MID}}>🚨 Needs Attention</span>
          <span style={{fontSize:10,fontWeight:700,padding:"2px 8px",borderRadius:20,background:"rgba(220,38,38,.1)",color:RED}}>{attn.length} CSMs</span>
        </div>
        <div style={{background:"#fff",border:"1px solid rgba(41,53,93,.09)",borderRadius:12,padding:18,boxShadow:"0 2px 10px rgba(41,53,93,.06)"}}>
          {attn.length===0
            ? <div style={{textAlign:"center",padding:24,color:GRN,fontWeight:700}}>🎉 Everyone at 90%+ today!</div>
            : <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr>
                  {["CSM","Team","Cadence","Gap to 90%"].map(h=><th key={h} style={{fontSize:10,textTransform:"uppercase",letterSpacing:".08em",color:MUT,fontWeight:700,padding:"0 8px 10px 0",textAlign:h==="CSM"||h==="Team"?"left":"right",borderBottom:"1px solid rgba(41,53,93,.08)"}}>{h}</th>)}
                </tr></thead>
                <tbody>{attn.map(c=>{
                  const i=lk(c.name)||{};
                  return (
                    <tr key={c.name}>
                      <td style={{padding:"9px 8px 9px 0",borderBottom:"1px solid rgba(41,53,93,.06)",verticalAlign:"top"}}>
                        <div style={{fontWeight:700}}>{c.name}</div>
                        <div style={{fontSize:11,color:MUT,fontStyle:"italic",marginTop:2}}>{tips(c.cadPct)}</div>
                      </td>
                      <td style={{padding:"9px 8px 9px 0",borderBottom:"1px solid rgba(41,53,93,.06)",fontSize:12,color:MUT}}>{st(i.t)}</td>
                      <td style={{padding:"9px 8px 9px 0",borderBottom:"1px solid rgba(41,53,93,.06)",textAlign:"right"}}><Bdg p={c.cadPct} label={c.cadPct>0?Math.round(c.cadPct*100)+"%":"—"}/></td>
                      <td style={{padding:"9px 0",borderBottom:"1px solid rgba(41,53,93,.06)",textAlign:"right",fontSize:12,fontWeight:700,color:RED}}>{c.cadPct>0?Math.round((0.9-c.cadPct)*100)+"%":"—"}</td>
                    </tr>
                  );
                })}</tbody>
              </table>}
        </div>
      </div>
      <div>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
          <span style={{fontFamily:"Nunito,sans-serif",fontSize:18,fontWeight:800,color:MID}}>🏅 Winning Today</span>
          <span style={{fontSize:10,fontWeight:700,padding:"2px 8px",borderRadius:20,background:"rgba(22,163,74,.1)",color:GRN}}>{wins.length} CSMs</span>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12}}>
          {wins.map(c=>{
            const i=lk(c.name)||{};
            return (
              <div key={c.name} style={{display:"flex",alignItems:"center",gap:12,padding:"13px 15px",background:"#fff",border:"1px solid rgba(41,53,93,.08)",borderLeft:"3px solid "+GRN,borderRadius:10}}>
                <div style={{fontSize:18}}>✅</div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:700,fontSize:13,color:MID,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.name}</div>
                  <div style={{fontSize:11,color:MUT,marginTop:2}}>{st(i.t||"")} · {Math.round(c.cadPct*100)}%</div>
                </div>
                <span style={{fontSize:10,fontWeight:700,padding:"3px 9px",borderRadius:20,background:"rgba(22,163,74,.1)",color:GRN,flexShrink:0}}>✓ Today</span>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

function OverviewView({data,csms}) {
  const rev=data.rev||[], email=data.email||[], cad=data.cadence||[];
  const totalRev=rev.reduce((s,r)=>s+r.total,0);
  const totalMRR=rev.reduce((s,r)=>s+r.mrr,0);
  const totalSent=email.reduce((s,r)=>s+r.sent,0);
  const avgOpen=email.length?email.reduce((s,r)=>s+r.openRate,0)/email.length:0;
  const avgCad=cad.length?cad.reduce((s,r)=>s+r.pct,0)/cad.length:0;
  const metrics=[
    rev.length>0?{l:"Total Revenue",v:f$(totalRev),s:"MRR "+fk(totalMRR),col:ACC}:null,
    email.length>0?{l:"Emails Sent",v:totalSent,s:email.length+" senders",col:"#5378FC"}:null,
    email.length>0?{l:"Avg Open Rate",v:fp(avgOpen),s:"Target 70%+",col:avgOpen>=0.7?GRN:YLW}:null,
    cad.length>0?{l:"Cadence Completion",v:fp(avgCad),s:"Target 90%+",col:avgCad>=0.9?GRN:avgCad>=0.5?YLW:RED}:null,
  ].filter(Boolean);
  const teamRev={};
  rev.forEach(r=>{if(r.team) teamRev[r.team]=(teamRev[r.team]||0)+r.total;});
  const tRevSort=Object.entries(teamRev).sort((a,b)=>b[1]-a[1]);
  const maxRev=(tRevSort[0]&&tRevSort[0][1])||1;
  const emailSort=[...email].sort((a,b)=>b.sent-a.sent).slice(0,10);
  const maxSent=(emailSort[0]&&emailSort[0].sent)||1;
  const cadSort=[...cad].filter(r=>r.count>0).sort((a,b)=>b.pct-a.pct);
  const card={background:"#fff",border:"1px solid rgba(41,53,93,.09)",borderRadius:12,padding:18,boxShadow:"0 2px 10px rgba(41,53,93,.06)"};
  return (
    <div style={{padding:"24px 28px"}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat("+metrics.length+",1fr)",gap:14,marginBottom:24}}>
        {metrics.map(m=>(
          <div key={m.l} style={{...card,position:"relative",overflow:"hidden"}}>
            <div style={{position:"absolute",top:0,left:0,right:0,height:3,background:m.col,borderRadius:"12px 12px 0 0"}}/>
            <div style={{fontSize:10,color:MUT,textTransform:"uppercase",letterSpacing:".08em",fontWeight:700,marginBottom:8}}>{m.l}</div>
            <div style={{fontFamily:"Nunito,sans-serif",fontSize:24,fontWeight:800,color:MID,letterSpacing:"-.02em",lineHeight:1,marginBottom:4}}>{m.v}</div>
            <div style={{fontSize:11,color:MUT}}>{m.s}</div>
          </div>
        ))}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20,marginBottom:20}}>
        <div style={card}>
          <div style={{fontSize:11,textTransform:"uppercase",letterSpacing:".1em",color:MUT,fontWeight:700,marginBottom:14}}>Revenue by Team</div>
          {tRevSort.length===0?<div style={{color:MUT,fontSize:13}}>No revenue data</div>:tRevSort.map(([t,v])=>(
            <div key={t} style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,fontSize:12}}>
              <span style={{width:140,flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                <span style={{display:"inline-block",width:6,height:6,borderRadius:"50%",background:tc(t),marginRight:6,verticalAlign:"middle"}}/>
                {st(t)}
              </span>
              <div style={{flex:1,height:5,background:"#ECEEF1",borderRadius:3,overflow:"hidden"}}>
                <div style={{width:(v/maxRev*100).toFixed(1)+"%",height:"100%",background:tc(t),opacity:.85,borderRadius:3}}/>
              </div>
              <span style={{width:56,textAlign:"right",fontSize:11,color:MUT,fontWeight:600,flexShrink:0}}>{fk(v)}</span>
            </div>
          ))}
        </div>
        <div style={card}>
          <div style={{fontSize:11,textTransform:"uppercase",letterSpacing:".1em",color:MUT,fontWeight:700,marginBottom:14}}>Email Activity — Top 10</div>
          {emailSort.length===0?<div style={{color:MUT,fontSize:13}}>No email data</div>:emailSort.map(r=>(
            <div key={r.name} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,fontSize:12}}>
              <span style={{width:120,flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.name}</span>
              <div style={{flex:1,height:5,background:"#ECEEF1",borderRadius:3,overflow:"hidden"}}>
                <div style={{width:(r.sent/maxSent*100).toFixed(1)+"%",height:"100%",background:"#5378FC",opacity:.8,borderRadius:3}}/>
              </div>
              <span style={{width:34,textAlign:"right",fontSize:11,color:MUT,fontWeight:600,flexShrink:0}}>{r.sent}</span>
              <span style={{width:42,textAlign:"right",fontSize:11,fontWeight:700,flexShrink:0,color:pCol(r.openRate)}}>{fp(r.openRate)}</span>
            </div>
          ))}
        </div>
      </div>
      <div style={card}>
        <div style={{fontSize:11,textTransform:"uppercase",letterSpacing:".1em",color:MUT,fontWeight:700,marginBottom:14}}>Cadence Completion — Yesterday</div>
        {cadSort.length===0?<div style={{color:MUT,fontSize:13}}>No cadence data</div>:
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"4px 32px"}}>
            {cadSort.map(r=>{
              const p=Math.min(r.pct,1), col=bCol(p,0.9,0.5);
              return (
                <div key={r.name} style={{display:"flex",alignItems:"center",gap:8,padding:"4px 0",borderBottom:"1px solid rgba(41,53,93,.05)"}}>
                  <span style={{flex:1,fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.name}</span>
                  <div style={{width:80,height:5,background:"#ECEEF1",borderRadius:3,overflow:"hidden",flexShrink:0}}>
                    <div style={{width:(p*100).toFixed(1)+"%",height:"100%",background:col,borderRadius:3}}/>
                  </div>
                  <span style={{width:32,textAlign:"right",fontSize:11,fontWeight:700,flexShrink:0,color:col}}>{Math.round(p*100)+"%"}</span>
                </div>
              );
            })}
          </div>}
      </div>
    </div>
  );
}

function LeaderboardView({csms,fc,ft,tier}) {
  const [sort,setSort]=useState({col:"rev",dir:"desc"});
  const [all,setAll]=useState(false);
  const medals=["🥇","🥈","🥉"];
  const filtered=csms.filter(c=>{
    const i=lk(c.name);
    if(fc&&(i&&i.c)!==fc) return false;
    if(ft&&(i&&i.t||c.team)!==ft) return false;
    if(tier&&(i&&i.r||c.tier)!==tier) return false;
    return true;
  }).sort((a,b)=>{
    const mul=sort.dir==="desc"?-1:1;
    return ((b[sort.col]||0)-(a[sort.col]||0))*mul;
  });
  const vis=all?filtered:filtered.slice(0,15);
  const th=(col,lbl)=>(
    <th onClick={()=>setSort(s=>({col,dir:s.col===col&&s.dir==="desc"?"asc":"desc"}))}
      style={{fontSize:10,textTransform:"uppercase",letterSpacing:".08em",color:MUT,fontWeight:700,padding:"0 0 10px",textAlign:"right",cursor:"pointer",userSelect:"none",borderBottom:"1px solid rgba(41,53,93,.08)"}}>
      {lbl}{sort.col===col?(sort.dir==="desc"?" ▼":" ▲"):""}
    </th>
  );
  return (
    <div style={{padding:"24px 28px"}}>
      <div style={{background:"#fff",border:"1px solid rgba(41,53,93,.09)",borderRadius:12,padding:18,boxShadow:"0 2px 10px rgba(41,53,93,.06)"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
          <thead><tr>
            <th style={{width:32,fontSize:10,color:MUT,fontWeight:700,padding:"0 0 10px",textAlign:"left",borderBottom:"1px solid rgba(41,53,93,.08)"}}>#</th>
            <th style={{fontSize:10,textTransform:"uppercase",letterSpacing:".08em",color:MUT,fontWeight:700,padding:"0 0 10px",textAlign:"left",borderBottom:"1px solid rgba(41,53,93,.08)"}}>CSM</th>
            <th style={{fontSize:10,textTransform:"uppercase",letterSpacing:".08em",color:MUT,fontWeight:700,padding:"0 0 10px",textAlign:"left",borderBottom:"1px solid rgba(41,53,93,.08)"}}>Team</th>
            {th("rev","Revenue")}
            {th("ints","Integrations")}
            {th("sent","Emails")}
            {th("openRate","Open %")}
            {th("cadPct","Cadence %")}
          </tr></thead>
          <tbody>{vis.map((c,i)=>{
            const info=lk(c.name)||{};
            const color=tc(info.t||c.team);
            return (
              <tr key={c.name}>
                <td style={{padding:"10px 0",borderBottom:"1px solid rgba(41,53,93,.06)"}}>{i<3?medals[i]:(i+1)+"."}</td>
                <td style={{padding:"10px 8px 10px 0",borderBottom:"1px solid rgba(41,53,93,.06)"}}>
                  <span style={{fontWeight:700}}>{c.name}</span>
                  {(info.r||c.tier)&&<span style={{fontSize:10,background:"rgba(41,53,93,.07)",color:MID,padding:"2px 6px",borderRadius:4,fontWeight:600,marginLeft:6}}>{info.r||c.tier}</span>}
                </td>
                <td style={{padding:"10px 8px 10px 0",borderBottom:"1px solid rgba(41,53,93,.06)"}}>
                  <span style={{display:"inline-block",width:6,height:6,borderRadius:"50%",background:color,marginRight:6,verticalAlign:"middle"}}/>
                  <span style={{color:MUT,fontSize:12}}>{st(info.t||c.team)}</span>
                </td>
                <td style={{padding:"10px 0",borderBottom:"1px solid rgba(41,53,93,.06)",textAlign:"right",fontWeight:700,color:ACC}}>{c.rev>0?f$(c.rev):"—"}</td>
                <td style={{padding:"10px 0",borderBottom:"1px solid rgba(41,53,93,.06)",textAlign:"right"}}>{c.ints>0?c.ints:"—"}</td>
                <td style={{padding:"10px 0",borderBottom:"1px solid rgba(41,53,93,.06)",textAlign:"right"}}>{c.sent>0?c.sent:"—"}</td>
                <td style={{padding:"10px 0",borderBottom:"1px solid rgba(41,53,93,.06)",textAlign:"right",fontWeight:700,color:pCol(c.openRate)}}>{c.sent>0?fp(c.openRate):"—"}</td>
                <td style={{padding:"10px 0",borderBottom:"1px solid rgba(41,53,93,.06)",textAlign:"right",fontWeight:700,color:pCol(c.cadPct)}}>{c.cadPct>0?fp(c.cadPct):"—"}</td>
              </tr>
            );
          })}</tbody>
        </table>
        {filtered.length>15&&(
          <button onClick={()=>setAll(s=>!s)} style={{display:"block",width:"100%",marginTop:14,background:"transparent",border:"1px solid rgba(41,53,93,.15)",color:MUT,fontSize:12,fontWeight:600,padding:8,borderRadius:8,cursor:"pointer"}}>
            {all?"Show top 15 ↑":"Show all "+filtered.length+" CSMs ↓"}
          </button>
        )}
      </div>
    </div>
  );
}

function ActivityView({data}) {
  const email=[...data.email||[]].sort((a,b)=>b.sent-a.sent);
  const cad=[...data.cadence||[]].sort((a,b)=>b.pct-a.pct);
  const card={background:"#fff",border:"1px solid rgba(41,53,93,.09)",borderRadius:12,padding:18,boxShadow:"0 2px 10px rgba(41,53,93,.06)"};
  return (
    <div style={{padding:"24px 28px"}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
        <div style={card}>
          <div style={{fontSize:11,textTransform:"uppercase",letterSpacing:".1em",color:MUT,fontWeight:700,marginBottom:14}}>Email Performance</div>
          {email.length===0?<div style={{color:MUT,fontSize:13}}>No email data</div>:
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead><tr>{["CSM","Sent","Opens","Open %","Reply %"].map(h=>(
                <th key={h} style={{fontSize:10,textTransform:"uppercase",color:MUT,fontWeight:700,padding:"0 0 8px",textAlign:h==="CSM"?"left":"right",borderBottom:"1px solid rgba(41,53,93,.08)"}}>{h}</th>
              ))}</tr></thead>
              <tbody>{email.map(r=>(
                <tr key={r.name}>
                  <td style={{padding:"7px 0",borderBottom:"1px solid rgba(41,53,93,.05)"}}>{r.name}</td>
                  <td style={{padding:"7px 0",borderBottom:"1px solid rgba(41,53,93,.05)",textAlign:"right"}}>{r.sent}</td>
                  <td style={{padding:"7px 0",borderBottom:"1px solid rgba(41,53,93,.05)",textAlign:"right"}}>{r.opens}</td>
                  <td style={{padding:"7px 0",borderBottom:"1px solid rgba(41,53,93,.05)",textAlign:"right",fontWeight:700,color:pCol(r.openRate)}}>{fp(r.openRate)}</td>
                  <td style={{padding:"7px 0",borderBottom:"1px solid rgba(41,53,93,.05)",textAlign:"right",fontWeight:700,color:pCol(r.replyRate)}}>{fp(r.replyRate)}</td>
                </tr>
              ))}</tbody>
            </table>}
        </div>
        <div style={card}>
          <div style={{fontSize:11,textTransform:"uppercase",letterSpacing:".1em",color:MUT,fontWeight:700,marginBottom:14}}>Cadence Touchpoints</div>
          {cad.length===0?<div style={{color:MUT,fontSize:13}}>No cadence data</div>:
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead><tr>{["CSM","Tasks","Removed","Done"].map(h=>(
                <th key={h} style={{fontSize:10,textTransform:"uppercase",color:MUT,fontWeight:700,padding:"0 0 8px",textAlign:h==="CSM"?"left":"right",borderBottom:"1px solid rgba(41,53,93,.08)"}}>{h}</th>
              ))}</tr></thead>
              <tbody>{cad.map(r=>{
                const p=Math.min(r.pct,1);
                const bg=p>=0.9?"rgba(22,163,74,.1)":p>=0.5?"rgba(217,119,6,.1)":"rgba(220,38,38,.1)";
                const fg=p>=0.9?"#166534":p>=0.5?"#854d0e":"#991b1b";
                return (
                  <tr key={r.name}>
                    <td style={{padding:"7px 0",borderBottom:"1px solid rgba(41,53,93,.05)"}}>{r.name}</td>
                    <td style={{padding:"7px 0",borderBottom:"1px solid rgba(41,53,93,.05)",textAlign:"right"}}>{r.count}</td>
                    <td style={{padding:"7px 0",borderBottom:"1px solid rgba(41,53,93,.05)",textAlign:"right",color:r.removed>0?RED:MUT}}>{r.removed>0?r.removed:"—"}</td>
                    <td style={{padding:"7px 0",borderBottom:"1px solid rgba(41,53,93,.05)",textAlign:"right"}}>
                      <span style={{fontSize:10,fontWeight:700,padding:"2px 8px",borderRadius:20,background:bg,color:fg}}>{Math.round(p*100)+"%"}</span>
                    </td>
                  </tr>
                );
              })}</tbody>
            </table>}
        </div>
      </div>
    </div>
  );
}

function AIPanel({csms,data,open,onClose}) {
  const [msgs,setMsgs]=useState([{role:"assistant",text:"Hi! I have your full team data. Ask me who needs coaching, which team is winning, or anything about performance."}]);
  const [inp,setInp]=useState("");
  const [busy,setBusy]=useState(false);
  const [hist,setHist]=useState([]);
  const ref=useRef();
  useEffect(()=>{ if(open&&ref.current) ref.current.scrollTop=ref.current.scrollHeight; },[msgs,open]);
  const ctx=()=>{
    const teamSums=COACHES.map(coach=>{
      const tcsms=csms.filter(c=>{ const i=lk(c.name); return(i&&i.c===coach.email)||c.coach===coach.email; });
      if(!tcsms.length) return coach.name+" ("+coach.team+"): no data";
      const cadC=tcsms.filter(c=>c.cadCount>0);
      const avgC=cadC.length?cadC.reduce((s,c)=>s+c.cadPct,0)/cadC.length:null;
      const emaC=tcsms.filter(c=>c.sent>0);
      const avgO=emaC.length?emaC.reduce((s,c)=>s+c.openRate,0)/emaC.length:null;
      const totR=tcsms.reduce((s,c)=>s+c.rev,0);
      const win=cadC.filter(c=>c.cadPct>=0.9).map(c=>c.name);
      const str=cadC.filter(c=>c.cadPct<0.9).sort((a,b)=>a.cadPct-b.cadPct).map(c=>c.name+" "+Math.round(c.cadPct*100)+"%");
      const nd=tcsms.filter(c=>c.cadCount===0).map(c=>c.name);
      const lines=[
        "COACH: "+coach.name+" | TEAM: "+coach.team,
        "  Cadence: "+(avgC!=null?Math.round(avgC*100)+"%":"no data")+" (target 90%+)",
        "  Email open: "+(avgO!=null?Math.round(avgO*100)+"%":"no data")+" (target 70%+)",
        "  Revenue: $"+totR.toLocaleString(),
        "  Winning: "+(win.join(", ")||"none"),
        "  Needs coaching: "+(str.join(", ")||"none"),
      ];
      if(nd.length) lines.push("  No data: "+nd.join(", "));
      return lines.join("\n");
    }).join("\n\n");
    return "You are a CSM coaching assistant for Thryv B2B SaaS. Give specific, actionable advice using real names and numbers. Be direct and concise. Winning threshold: 90%+ cadence completion.\n\n"+teamSums;
  };
  const send=async()=>{
    const msg=inp.trim(); if(!msg||busy) return;
    setInp(""); setBusy(true);
    const newMsgs=[...msgs,{role:"assistant",text:"Thinking..."}];
    const msgIdx=newMsgs.length-1;
    setMsgs([...msgs,{role:"user",text:msg}]);
    const newHist=[...hist,{role:"user",content:msg}];
    setHist(newHist);
    try {
      const res=await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:800,system:ctx(),messages:newHist.slice(-8)})
      });
      const d=await res.json();
      const reply=d.content&&d.content[0]?d.content[0].text:"No response.";
      setMsgs(m=>[...m,{role:"assistant",text:reply}]);
      setHist(h=>[...h,{role:"assistant",content:reply}]);
    } catch(e) {
      setMsgs(m=>[...m,{role:"assistant",text:"Error connecting. Please try again."}]);
    }
    setBusy(false);
  };
  const chips=["Who needs coaching today?","Which team is winning?","Who has the best cadence?","Compare all 6 teams"];
  if(!open) return null;
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(41,53,93,.4)",zIndex:500,display:"flex",alignItems:"flex-end",justifyContent:"flex-end",padding:20}}>
      <div style={{width:420,height:"72vh",maxHeight:680,background:"#fff",borderRadius:16,display:"flex",flexDirection:"column",overflow:"hidden",boxShadow:"0 8px 40px rgba(41,53,93,.2)"}}>
        <div style={{padding:"14px 18px",background:MID,display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
          <span style={{fontFamily:"Nunito,sans-serif",fontSize:14,fontWeight:800,color:"#fff"}}>🤖 AI Coaching Assistant</span>
          <button onClick={onClose} style={{background:"transparent",border:"none",color:"rgba(255,255,255,.6)",fontSize:20,cursor:"pointer",padding:"2px 6px",borderRadius:6}}>✕</button>
        </div>
        <div ref={ref} style={{flex:1,overflowY:"auto",padding:16,display:"flex",flexDirection:"column",gap:10}}>
          {msgs.map((m,i)=>(
            <div key={i} style={{display:"flex",flexDirection:"column",alignItems:m.role==="user"?"flex-end":"flex-start"}}>
              <div style={{maxWidth:"88%",padding:"10px 14px",fontSize:13,lineHeight:1.55,
                background:m.role==="user"?ACC:"#F4F6FB",color:m.role==="user"?"#fff":"#121212",
                border:m.role==="user"?"none":"1px solid rgba(41,53,93,.08)",
                borderRadius:m.role==="user"?"12px 12px 3px 12px":"12px 12px 12px 3px"}}
                dangerouslySetInnerHTML={{__html:m.text.replace(/\*\*(.*?)\*\*/g,"<strong>$1</strong>").replace(/\n/g,"<br/>")}}
              />
            </div>
          ))}
        </div>
        {msgs.length<=1&&(
          <div style={{display:"flex",flexWrap:"wrap",gap:6,padding:"0 14px 10px"}}>
            {chips.map(c=>(
              <span key={c} onClick={()=>{setInp(c);}} style={{fontSize:11,fontWeight:600,padding:"4px 10px",borderRadius:20,border:"1px solid rgba(41,53,93,.15)",background:"#F4F6FB",color:MID,cursor:"pointer",whiteSpace:"nowrap"}}>{c}</span>
            ))}
          </div>
        )}
        <div style={{padding:"12px 14px",borderTop:"1px solid rgba(41,53,93,.08)",flexShrink:0}}>
          <div style={{display:"flex",gap:8,alignItems:"flex-end"}}>
            <textarea value={inp} onChange={e=>setInp(e.target.value)}
              onKeyDown={e=>{ if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();send();} }}
              placeholder="Ask about your team..."
              style={{flex:1,border:"1px solid rgba(41,53,93,.15)",borderRadius:10,padding:"9px 12px",fontSize:13,resize:"none",outline:"none",minHeight:40,maxHeight:100,lineHeight:1.5,fontFamily:"inherit"}}/>
            <button onClick={send} disabled={busy||!inp.trim()}
              style={{background:ACC,border:"none",borderRadius:8,padding:"0 14px",cursor:"pointer",color:"#fff",fontSize:16,flexShrink:0,height:40,minWidth:40,opacity:(busy||!inp.trim())?0.4:1}}>↑</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function UploadZone({loaded,onFile,onSnap}) {
  const [drag,setDrag]=useState(null);
  const cards=[
    {type:"rev",icon:"💰",title:"Revenue & Integrations",sub:"Weekly OTR, MRR, integrations"},
    {type:"email",icon:"📧",title:"Email Performance",sub:"CT-Conquer cadence email stats"},
    {type:"cadence",icon:"✅",title:"Cadence Completed",sub:"Yesterday touchpoint completion"},
  ];
  return (
    <div style={{padding:"32px 32px 24px"}}>
      <div style={{textAlign:"center",marginBottom:24}}>
        <div style={{fontFamily:"Nunito,sans-serif",fontSize:22,fontWeight:800,color:MID,marginBottom:6}}>Upload your reports</div>
        <div style={{fontSize:13,color:MUT,marginBottom:12}}>Upload CSVs below — or import a snapshot shared by a teammate</div>
        <label style={{display:"inline-flex",alignItems:"center",gap:6,fontSize:12,fontWeight:700,padding:"7px 16px",borderRadius:8,cursor:"pointer",background:MID,color:"#fff"}}>
          📥 Import teammate snapshot
          <input type="file" accept=".json" style={{display:"none"}} onChange={e=>{ if(e.target.files[0]){onSnap(e.target.files[0]);e.target.value="";} }}/>
        </label>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16,maxWidth:780,margin:"0 auto"}}>
        {cards.map(({type,icon,title,sub})=>{
          const isLoaded=!!loaded[type];
          return (
            <div key={type}
              onDragOver={e=>{e.preventDefault();setDrag(type);}}
              onDragLeave={()=>setDrag(null)}
              onDrop={e=>{e.preventDefault();setDrag(null);onFile(type,e.dataTransfer.files[0]);}}
              style={{border:"1.5px dashed "+(isLoaded?"#16a34a":drag===type?ACC:"rgba(41,53,93,.2)"),borderRadius:14,padding:"28px 20px",background:isLoaded?"rgba(22,163,74,.02)":"#fff",textAlign:"center",transition:"all .2s"}}>
              <div style={{fontSize:24,marginBottom:12}}>{icon}</div>
              <div style={{fontFamily:"Nunito,sans-serif",fontSize:14,fontWeight:800,color:MID,marginBottom:4}}>{title}</div>
              <div style={{fontSize:11,color:MUT,marginBottom:16,lineHeight:1.5}}>{sub}</div>
              <label style={{display:"inline-flex",alignItems:"center",gap:6,background:isLoaded?"#16a34a":ACC,color:"#fff",fontSize:12,fontWeight:700,padding:"7px 16px",borderRadius:8,cursor:"pointer"}}>
                {isLoaded?"✓ "+loaded[type].name.replace(/\.(xlsx?|csv)$/i,"").substring(0,22):"↑ Choose file"}
                <input type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>{onFile(type,e.target.files[0]);e.target.value="";}}/>
              </label>
              <div style={{fontSize:10,color:MUT,marginTop:8}}>{isLoaded?loaded[type].rows.length+" records loaded":".xlsx · .csv"}</div>
            </div>
          );
        })}
      </div>
    </div>
  );
}


function PinLock({onUnlock}) {
  const [val, setVal] = React.useState("");
  const [err, setErr] = React.useState(false);
  const check = () => {
    if(val === DASHBOARD_PIN) {
      try { sessionStorage.setItem(PIN_KEY, "1"); } catch(e) {}
      onUnlock();
    } else {
      setErr(true);
      setVal("");
      setTimeout(()=>setErr(false), 1500);
    }
  };
  return (
    <div style={{minHeight:"100vh",background:"#F4F6FB",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{background:"#fff",borderRadius:16,padding:"40px 36px",boxShadow:"0 4px 32px rgba(41,53,93,.12)",textAlign:"center",width:340}}>
        <div style={{width:56,height:56,borderRadius:"50%",background:"#29355D",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 20px",fontSize:24}}>🔒</div>
        <div style={{fontFamily:"Nunito,sans-serif",fontSize:20,fontWeight:800,color:"#29355D",marginBottom:6}}>CSM Coaching Dashboard</div>
        <div style={{fontSize:13,color:"#808080",marginBottom:24}}>Enter your team PIN to continue</div>
        <input
          type="password"
          value={val}
          onChange={e=>setVal(e.target.value)}
          onKeyDown={e=>e.key==="Enter"&&check()}
          placeholder="Enter PIN"
          autoFocus
          style={{width:"100%",padding:"10px 14px",fontSize:15,borderRadius:10,border:"1px solid "+(err?"#dc2626":"rgba(41,53,93,.2)"),outline:"none",textAlign:"center",letterSpacing:4,marginBottom:12,background:err?"rgba(220,38,38,.04)":"#fff",transition:"border-color .2s"}}
        />
        {err&&<div style={{color:"#dc2626",fontSize:12,marginBottom:10}}>Incorrect PIN — try again</div>}
        <button onClick={check}
          style={{width:"100%",padding:"11px",background:"#FF5000",border:"none",borderRadius:10,color:"#fff",fontSize:14,fontWeight:700,cursor:"pointer"}}>
          Unlock Dashboard
        </button>
        <div style={{fontSize:11,color:"#808080",marginTop:16}}>Contact your coach for the PIN</div>
      </div>
    </div>
  );
}

export default function App() {
  const [unlocked,setUnlocked]=useState(()=>{ try { return sessionStorage.getItem(PIN_KEY)==="1"; } catch(e){ return false; } });
  const [loaded,setLoaded]=useState({rev:null,email:null,cadence:null});
  const [tab,setTab]=useState("upload");
  const [filters,setFilters]=useState({coach:"",team:"",tier:""});
  const [aiOpen,setAiOpen]=useState(false);
  const [snapDate,setSnapDate]=useState(null);
  const [syncStatus,setSyncStatus]=useState("idle"); // idle | syncing | ok | error | loading

  const raw={rev:(loaded.rev&&loaded.rev.rows)||null,email:(loaded.email&&loaded.email.rows)||null,cadence:(loaded.cadence&&loaded.cadence.rows)||null};
  const csms=raw.rev||raw.email||raw.cadence?buildCSMs(raw):[];
  const hasData=!!(raw.rev||raw.email||raw.cadence);

  const CTMAP={"odirlm01@thryv.com":"The Dominican Dream Team","chase.boyd@thryv.com":"Boyd Meets World","elizabeth.white@thryv.com":"White Wave Warriors","kendra.morelli@thryv.com":"Team Thryv-More(lli)","trisha.stalnaker@thryv.com":"Team Status Engaged","aaron.taylor@thryv.com":"Team Aurorians"};
  const ALLTEAMS=["The Dominican Dream Team","Boyd Meets World","White Wave Warriors","Team Thryv-More(lli)","Team Status Engaged","Team Aurorians"];
  const TIERS=["CSMI","CSMII","CSMIII","SSMI","SSMII"];

  const handleFile=(type,file)=>{
    if(!file) return;
    const reader=new FileReader();
    reader.onload=e=>{
      try {
        const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});
        const parsed=type==="rev"?parseRev(wb):type==="email"?parseEmail(wb):parseCadence(wb);
        setLoaded(prev=>{
          const next={...prev,[type]:{name:file.name,rows:parsed}};
          storageSave(next);
          return next;
        });
        const now=new Date().toISOString();
        setSnapDate(now);
        if(tab==="upload") setTab("coaching");
        setSyncStatus("ok");
        // Data goes to Google Sheet manually - see instructions in upload zone
      } catch(err){ alert("Error: "+err.message); }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleSnap=(file)=>{
    const reader=new FileReader();
    reader.onload=e=>{
      try {
        const snap=JSON.parse(e.target.result);
        if(snap.v!=="csm1"){ alert("Invalid snapshot file."); return; }
        setLoaded(snap.loaded||{rev:null,email:null,cadence:null});
        setSnapDate(snap.at||null);
        storageSave(snap.loaded||{});
        setTab("coaching");
      } catch(err){ alert("Error reading snapshot: "+err.message); }
    };
    reader.readAsText(file);
  };

  useEffect(()=>{
    // Load from localStorage first (instant)
    const saved = storageLoad();
    if(saved&&(saved.rev||saved.email||saved.cadence)){
      setLoaded(saved); setTab("coaching");
    }
    // Then pull fresh from Google Sheets
    setSyncStatus("loading");
    pullFromSheets().then(({revRows,emailRows,cadRows})=>{
      const revParsed   = sheetToRev(revRows);
      const emailParsed = sheetToEmail(emailRows);
      const cadParsed   = sheetToCadence(cadRows);
      const hasData = revParsed.length>0||emailParsed.length>0||cadParsed.length>0;
      if(hasData){
        const next={
          rev:     revParsed.length>0   ? {name:"Google Sheet", rows:revParsed}   : null,
          email:   emailParsed.length>0 ? {name:"Google Sheet", rows:emailParsed} : null,
          cadence: cadParsed.length>0   ? {name:"Google Sheet", rows:cadParsed}   : null,
        };
        setLoaded(next);
        storageSave(next);
        setSnapDate(new Date().toISOString());
        setTab("coaching");
      }
      setSyncStatus("ok");
    }).catch(()=>setSyncStatus("error"));
  },[]);

  const selStyle=(active)=>({
    fontSize:12,fontWeight:600,
    color:active?ACC:MID,
    background:active?"rgba(255,80,0,.04)":"#F4F6FB",
    border:"1px solid "+(active?ACC:"rgba(41,53,93,.15)"),
    borderRadius:8,padding:"6px 12px",cursor:"pointer",
  });

  const tabs=[{id:"coaching",label:"🏆 Coaching"},{id:"overview",label:"Overview"},{id:"leaderboard",label:"Leaderboard"},{id:"activity",label:"Activity"}];

  if(!unlocked) return <PinLock onUnlock={()=>setUnlocked(true)}/>;

  return (
    <div style={{fontFamily:"Nunito Sans,sans-serif",background:"#F4F6FB",minHeight:"100vh",fontSize:14}}>
      <div style={{padding:"14px 28px",background:MID,display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:100,boxShadow:"0 1px 12px rgba(41,53,93,.18)"}}>
        <div style={{display:"flex",alignItems:"center",gap:14}}>
          <img src="https://assets.thryv.com/prod/media/thryv-main-logo-white.png" alt="Thryv" style={{height:26}} onError={e=>{e.target.style.display="none";}}/>
          <div style={{width:1,height:20,background:"rgba(255,255,255,.2)"}}/>
          <span style={{fontSize:13,fontWeight:500,color:"rgba(255,255,255,.65)"}}>CSM Coaching Dashboard</span>
        </div>
        <div style={{display:"flex",gap:3,background:"rgba(255,255,255,.08)",borderRadius:8,padding:3}}>
          {hasData&&tabs.map(t=>(
            <button key={t.id} onClick={()=>setTab(t.id)}
              style={{fontSize:12,fontWeight:600,padding:"6px 14px",borderRadius:6,border:"none",cursor:"pointer",
                background:tab===t.id?ACC:"transparent",color:tab===t.id?"#fff":"rgba(255,255,255,.55)",transition:"all .15s"}}>
              {t.label}
            </button>
          ))}
          {!hasData&&<span style={{fontSize:12,color:"rgba(255,255,255,.4)",padding:"6px 14px"}}>Upload files to get started</span>}
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          {syncStatus==="syncing"&&<span style={{fontSize:11,padding:"4px 10px",borderRadius:20,background:"rgba(255,255,255,.1)",color:"rgba(255,255,255,.6)"}}>⟳ Syncing...</span>}
          {syncStatus==="ok"&&<span style={{fontSize:11,padding:"4px 10px",borderRadius:20,background:"rgba(22,163,74,.25)",color:"#86efac"}}>✓ Synced</span>}
          {syncStatus==="loading"&&<span style={{fontSize:11,padding:"4px 10px",borderRadius:20,background:"rgba(255,255,255,.1)",color:"rgba(255,255,255,.6)"}}>⟳ Loading...</span>}
          {syncStatus==="error"&&<span style={{fontSize:11,padding:"4px 10px",borderRadius:20,background:"rgba(220,38,38,.25)",color:"#fca5a5"}}>✗ Sync error</span>}
          {hasData&&(
            <button onClick={()=>exportSnap(loaded)}
              style={{fontSize:11,fontWeight:700,padding:"5px 12px",borderRadius:20,background:"rgba(255,255,255,.1)",color:"rgba(255,255,255,.8)",cursor:"pointer",border:"1px solid rgba(255,255,255,.2)"}}>
              ↓ Export
            </button>
          )}
        </div>
      </div>

      {hasData&&(
        <div style={{background:"#fff",borderBottom:"1px solid rgba(41,53,93,.08)",padding:"10px 28px",display:"flex",gap:10,flexWrap:"wrap",alignItems:"center",boxShadow:"0 2px 8px rgba(41,53,93,.06)"}}>
          <select value={filters.coach} onChange={e=>{ const v=e.target.value; setFilters({coach:v,team:v?CTMAP[v]||"":"",tier:filters.tier}); }} style={selStyle(!!filters.coach)}>
            <option value="">All Coaches</option>
            {COACHES.map(c=><option key={c.email} value={c.email}>{c.name}</option>)}
          </select>
          <select value={filters.team} onChange={e=>setFilters(f=>({...f,team:e.target.value}))} style={selStyle(!!filters.team)}>
            <option value="">All Teams</option>
            {ALLTEAMS.map(t=><option key={t} value={t}>{st(t)}</option>)}
          </select>
          <select value={filters.tier} onChange={e=>setFilters(f=>({...f,tier:e.target.value}))} style={selStyle(!!filters.tier)}>
            <option value="">All Tiers</option>
            {TIERS.map(t=><option key={t} value={t}>{t}</option>)}
          </select>
          {Object.values(filters).some(v=>v)&&(
            <button onClick={()=>setFilters({coach:"",team:"",tier:""})} style={{fontSize:11,fontWeight:700,color:ACC,background:"none",border:"none",cursor:"pointer",padding:"4px 8px"}}>✕ Clear</button>
          )}
          <span style={{marginLeft:"auto",fontSize:11,color:MUT,fontWeight:600}}>
            {snapDate?"Updated: "+new Date(snapDate).toLocaleDateString("en-US",{month:"short",day:"numeric"}):""}
          </span>
          <button onClick={()=>setTab("upload")} style={{fontSize:12,fontWeight:600,color:MUT,background:"transparent",border:"1px solid rgba(41,53,93,.15)",borderRadius:8,padding:"6px 14px",cursor:"pointer"}}>↑ Upload new files</button>
        </div>
      )}

      {tab==="upload"&&<UploadZone loaded={loaded} onFile={handleFile} onSnap={handleSnap}/>}
      {tab==="coaching"&&hasData&&<CoachingView csms={csms} fc={filters.coach} ft={filters.team}/>}
      {tab==="overview"&&hasData&&<OverviewView data={raw} csms={csms}/>}
      {tab==="leaderboard"&&hasData&&<LeaderboardView csms={csms} fc={filters.coach} ft={filters.team} tier={filters.tier}/>}
      {tab==="activity"&&hasData&&<ActivityView data={raw}/>}

      {hasData&&(
        <button onClick={()=>setAiOpen(true)}
          style={{position:"fixed",bottom:24,right:24,width:54,height:54,borderRadius:"50%",background:MID,border:"none",cursor:"pointer",boxShadow:"0 4px 20px rgba(41,53,93,.35)",fontSize:22,zIndex:499,display:"flex",alignItems:"center",justifyContent:"center"}}
          onMouseOver={e=>{ e.currentTarget.style.background=ACC; }}
          onMouseOut={e=>{ e.currentTarget.style.background=MID; }}>
          🤖
        </button>
      )}
      <AIPanel csms={csms} data={raw} open={aiOpen} onClose={()=>setAiOpen(false)}/>
    </div>
  );
}
