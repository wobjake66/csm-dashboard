import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

const PIN = "thryv2025";
const PIN_KEY = "csm_pin_v1";

const CSV_REV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=322916467&single=true&output=csv";
const CSV_EMAIL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=0&single=true&output=csv";
const CSV_CAD = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=1973544046&single=true&output=csv";

const COACHES = [
  {email:"odirlm01@thryv.com",      name:"Mia O'Dirling",    team:"The Dominican Dream Team"},
  {email:"chase.boyd@thryv.com",    name:"Chase Boyd",        team:"Boyd Meets World"},
  {email:"elizabeth.white@thryv.com",name:"Elizabeth White",  team:"White Wave Warriors"},
  {email:"kendra.morelli@thryv.com", name:"Kendra Morelli",   team:"Team Thryv-More(lli)"},
  {email:"trisha.stalnaker@thryv.com",name:"Trisha Stalnaker",team:"Team Status Engaged"},
  {email:"aaron.taylor@thryv.com",  name:"Aaron Taylor",      team:"Team Aurorians"},
];

const TCOLORS = {
  "The Dominican Dream Team":"#FF5000","Boyd Meets World":"#4A5D8C",
  "White Wave Warriors":"#808080","Team Thryv-More(lli)":"#29355D",
  "Team Status Engaged":"#E03000","Team Aurorians":"#5378FC",
};

const ROSTER = {
  "darling danais santos taveras":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "darling danais":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "heidi torres uribe":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "irina larianni molina molina":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "irina molina molina":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "jathzelyn elizabeth fortuna paulino":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "jathzelyn fortuna":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "jazz fortuna":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "johnny cornielle":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII"},
  "johnny cornielle montas":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII"},
  "joseph guillermo carmona garcia":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "joseph carmona garcia":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "samuel frias de paula":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII"},
  "sam frias":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII"},
  "sati ananda pimentel malespin":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "sati pimentel":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "victor abner moscoso fernandez":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "victor moscoso":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "wilson mercedes":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII"},
  "yessica montero urena":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "yessica montero":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "barbara larrosa presinal":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMII"},
  "barbara larrosa":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMII"},
  "deivis pena":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "elianny tena antigua":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "eric johnson":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMII"},
  "katelyn ankrom":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMII"},
  "kyle dye":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMIII"},
  "luis aguasvivas peralta":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "juan liberato":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "sarah swanson":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "tyler moeggenberg":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "tyler popplewell":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "april hall":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII"},
  "anthony yen":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII"},
  "damita hill":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII"},
  "dorka frias lantigua":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI"},
  "florence francois nova":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI"},
  "kennedy sanchez":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII"},
  "matt sword":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI"},
  "michael furlong":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII"},
  "rossi valerio tejeda":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI"},
  "rossi valerio":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI"},
  "steven saunders":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII"},
  "yolanda ramirez":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII"},
  "yolanda ramirez-drake":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII"},
  "alejandro rodriguez-medina":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"SSMII"},
  "alejandro rodriguez medina":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"SSMII"},
  "chelsea dingus":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"SSMII"},
  "karmita turner":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMII"},
  "lauren carter":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMII"},
  "libby booher":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMIII"},
  "misti dixon":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMIII"},
  "misty decatur":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"SSMI"},
  "saira julian guzman":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMI"},
  "scott mather":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"SSMII"},
  "ashley shaffer":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMII"},
  "ashley vasquez mena":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "ashley vasquez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "felix caba jimenez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "karen capellan tavarez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "karen capellan":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "karissa hernandez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMII"},
  "kellie lester":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMIII"},
  "mark velazquez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"SSMII"},
  "merve (mj) brielmann":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMII"},
  "mj brielmann":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMII"},
  "rafael sencion sencion":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "rafael sencion":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "stacy roers":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMIII"},
  "taylor kidd":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "david crisler":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMII"},
  "dave crisler":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMII"},
  "ellise payne":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "indu vijay":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "matt daly":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "nikita siepen-bowers":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "nikita siepen bowers":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "peter manalac":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "sakshi mahalwal":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "sylvia appla":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "sylvia":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "tracy-ann gaudencio":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "warda gul":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "zoltan rudolf":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMII"},
};

const lk = n => n ? ROSTER[n.toLowerCase().trim()] || null : null;
const pm = v => { const n = parseFloat(String(v||0).replace(/[$,]/g,"")); return isNaN(n)?0:n; };
const pn = v => { const n = parseFloat(String(v||0).replace(/[,%]/g,"")); return isNaN(n)?0:n; };
const f$ = n => "$"+(+n).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});
const fk = n => n>=1000?"$"+(n/1000).toFixed(1).replace(/\.0$/,"")+"k":"$"+Math.round(n);
const fp = n => (isNaN(n)||n==null)?"—":(n*100).toFixed(1)+"%";
const tc = t => TCOLORS[t]||"#808080";
const st = t => (t||"").replace("The ","").replace("Team ","");
const MID="#29355D",ACC="#FF5000",GRN="#16a34a",YLW="#d97706",RED="#dc2626",MUT="#808080";
const bCol=(v,hi,lo)=>v>=hi?GRN:v>=lo?YLW:RED;
const pCol=p=>p>=0.9?GRN:p>=0.5?YLW:p>0?RED:MUT;

// ── CSV PARSER ────────────────────────────────────────────
function parseCSVLine(line) {
  const result = [];
  let cur = "", inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQ && line[i+1] === '"') { cur += '"'; i++; }
      else inQ = !inQ;
    } else if (ch === ',' && !inQ) {
      result.push(cur.trim());
      cur = "";
    } else {
      cur += ch;
    }
  }
  result.push(cur.trim());
  return result;
}

async function fetchCSV(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error("HTTP " + res.status);
  const text = await res.text();
  const lines = text.split("\n").filter(l => l.trim());
  if (lines.length < 2) return [];
  const headers = parseCSVLine(lines[0]);
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const vals = parseCSVLine(lines[i]);
    if (!vals.some(v => v)) continue;
    const obj = {};
    headers.forEach((h, j) => { obj[h.trim()] = (vals[j] || "").trim(); });
    rows.push(obj);
  }
  return rows;
}

// ── DATA MAPPERS ──────────────────────────────────────────
function mapRev(rows) {
  const byName = {};
  rows.forEach(r => {
    const name = (r["CSM Name"] || "").trim();
    if (!name) return;
    if (!byName[name]) byName[name] = { name, team:(r["CSM Team!"]||r["CSM Team! "]||"").trim(), tier:(r["CSM Tier"]||"").trim(), mrr:0, otr:0, total:0, nonrev:0 };
    byName[name].mrr   += pm(r["MRR $ Added"]);
    byName[name].otr   += pm(r["OTR $ Added"]);
    byName[name].total += pm(r["Total Revenue Added"]);
    byName[name].nonrev += (r["Non-Revenue Integrations"]||"").trim() ? 1 : 0;
  });
  return Object.values(byName);
}

function mapEmail(rows) {
  return rows.map(r => {
    const name = (r["Touchpoint: Owner Name"]||r["Name"]||r["name"]||"").trim();
    if (!name || name === "Total") return null;
    return {
      name,
      sent: pn(r["Sum of Emails Sent"]||r["sent"]||0),
      opens: pn(r["Sum of Email Opens"]||r["opens"]||0),
      replies: pn(r["Sum of Email Replies"]||r["replies"]||0),
      openRate: pn(r["Open Rate"]||r["openRate"]||0),
      replyRate: pn(r["Reply Rate"]||r["replyRate"]||0),
    };
  }).filter(Boolean);
}

function mapCadence(rows) {
  return rows.map(r => {
    const name = (r["name"]||r["Name"]||r["CSM"]||r["Touchpoint: Owner Name"]||"").trim();
    if (!name || name === "Total") return null;
    let pct = pn(r["pct"]||r["Pct"]||r["% Completed"]||r["Completion"]||0);
    if (pct > 1) pct = pct / 100;
    return {
      name,
      count: pn(r["count"]||r["Count"]||r["Tasks"]||0),
      pct,
      removed: pn(r["removed"]||r["Removed"]||0),
    };
  }).filter(Boolean);
}

// ── BUILD CSM LIST ────────────────────────────────────────
function buildCSMs(rev, email, cad) {
  const m = {};
  const get = name => {
    if (!m[name]) {
      const info = lk(name);
      m[name] = { name, team:(info&&info.t)||"", coach:(info&&info.c)||"", tier:(info&&info.r)||"",
        rev:0, mrr:0, otr:0, ints:0, nonrev:0, sent:0, opens:0, replies:0,
        openRate:0, replyRate:0, cadCount:0, cadPct:0 };
    }
    return m[name];
  };
  (rev||[]).forEach(r => {
    const c = get(r.name);
    if (!c.team && r.team) c.team = r.team;
    if (!c.tier && r.tier) c.tier = r.tier;
    c.rev += r.total; c.mrr += r.mrr; c.otr += r.otr;
    if (r.total > 0) c.ints++;
    c.nonrev += r.nonrev;
  });
  (email||[]).forEach(r => {
    const c = get(r.name);
    c.sent = r.sent; c.opens = r.opens; c.replies = r.replies;
    c.openRate = r.openRate; c.replyRate = r.replyRate;
  });
  (cad||[]).forEach(r => {
    const c = get(r.name);
    c.cadCount = r.count; c.cadPct = r.pct;
  });
  return Object.values(m);
}

// ── SUB COMPONENTS ────────────────────────────────────────
const card = { background:"#fff", border:"1px solid rgba(41,53,93,.09)", borderRadius:12, padding:18, boxShadow:"0 2px 10px rgba(41,53,93,.06)" };

function Bar({label,val,hi,lo}) {
  const pct = val!=null ? Math.min(val*100,100) : 0;
  const col = val!=null ? bCol(val,hi,lo) : "#e5e7eb";
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
  return <span style={{fontSize:10,fontWeight:700,padding:"1px 7px",borderRadius:20,background:bg,color:fg,flexShrink:0}}>{label||(p>=0.9?"Win":p>=0.5?"Watch":p>0?"Coach":"—")}</span>;
}

function CoachCard({coach,csms}) {
  const team = csms.filter(c => { const i=lk(c.name); return (i&&i.c===coach.email)||c.coach===coach.email; });
  const cadC = team.filter(c=>c.cadCount>0);
  const avgCad = cadC.length ? cadC.reduce((s,c)=>s+c.cadPct,0)/cadC.length : null;
  const emC = team.filter(c=>c.sent>0);
  const avgOpen = emC.length ? emC.reduce((s,c)=>s+c.openRate,0)/emC.length : null;
  const revPct = team.length ? team.filter(c=>c.rev>0||c.ints>0).length/team.length : null;
  let sc=0,sf=0;
  if(avgCad!=null){sc+=Math.min(avgCad/0.9,1)*40;sf++;}
  if(avgOpen!=null){sc+=Math.min(avgOpen/0.7,1)*30;sf++;}
  if(revPct!=null){sc+=Math.min(revPct/0.7,1)*30;sf++;}
  const score = sf?Math.round(Math.min(sc,100)):null;
  const st2 = score==null?"none":score>=75?"win":score>=50?"warn":"att";
  const sCol = st2==="win"?GRN:st2==="warn"?YLW:st2==="att"?RED:MUT;
  const topC = st2==="win"?GRN:st2==="warn"?YLW:st2==="att"?RED:"#e5e7eb";
  const wins=cadC.filter(c=>c.cadPct>=0.9).length;
  const warns=cadC.filter(c=>c.cadPct>=0.5&&c.cadPct<0.9).length;
  const atts=cadC.filter(c=>c.cadPct>0&&c.cadPct<0.5).length;
  const sorted=[...team].sort((a,b)=>(a.cadPct||0)-(b.cadPct||0)).slice(0,7);
  return (
    <div style={{...card,position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:0,left:0,right:0,height:3,background:topC,borderRadius:"12px 12px 0 0"}}/>
      <div style={{fontFamily:"Nunito,sans-serif",fontSize:14,fontWeight:800,color:MID,marginTop:4}}>{coach.name}</div>
      <div style={{fontSize:11,fontWeight:600,color:tc(coach.team),marginBottom:10}}>● {coach.team}</div>
      <div style={{display:"flex",alignItems:"baseline",gap:6,marginBottom:8}}>
        <span style={{fontFamily:"Nunito,sans-serif",fontSize:44,fontWeight:800,lineHeight:1,color:sCol}}>{score!=null?score:"—"}</span>
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
        <div style={card}>
          {attn.length===0
            ? <div style={{textAlign:"center",padding:24,color:GRN,fontWeight:700}}>🎉 Everyone at 90%+ today!</div>
            : <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr>{["CSM","Team","Cadence","Gap to 90%"].map(h=><th key={h} style={{fontSize:10,textTransform:"uppercase",color:MUT,fontWeight:700,padding:"0 8px 10px 0",textAlign:h==="CSM"||h==="Team"?"left":"right",borderBottom:"1px solid rgba(41,53,93,.08)"}}>{h}</th>)}</tr></thead>
                <tbody>{attn.map(c=>{
                  const i=lk(c.name)||{};
                  return (
                    <tr key={c.name}>
                      <td style={{padding:"9px 8px 9px 0",borderBottom:"1px solid rgba(41,53,93,.06)",verticalAlign:"top"}}>
                        <div style={{fontWeight:700}}>{c.name}</div>
                        <div style={{fontSize:11,color:MUT,fontStyle:"italic",marginTop:2}}>{tips(c.cadPct)}</div>
                      </td>
                      <td style={{padding:"9px 8px 9px 0",borderBottom:"1px solid rgba(41,53,93,.06)",fontSize:12,color:MUT}}>{st(i.t||"")}</td>
                      <td style={{padding:"9px 8px 9px 0",borderBottom:"1px solid rgba(41,53,93,.06)",textAlign:"right"}}><Bdg p={c.cadPct} label={c.cadPct>0?Math.round(c.cadPct*100)+"%":"—"}/></td>
                      <td style={{padding:"9px 0",borderBottom:"1px solid rgba(41,53,93,.06)",textAlign:"right",fontSize:12,fontWeight:700,color:RED}}>{c.cadPct>0&&c.cadPct<0.9?Math.round((0.9-c.cadPct)*100)+"%":"—"}</td>
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

function OverviewView({rev,email,cad,csms}) {
  const totalRev=rev.reduce((s,r)=>s+r.total,0);
  const totalMRR=rev.reduce((s,r)=>s+r.mrr,0);
  const totalSent=email.reduce((s,r)=>s+r.sent,0);
  const avgOpen=email.length?email.reduce((s,r)=>s+r.openRate,0)/email.length:0;
  const avgCad=cad.length?cad.reduce((s,r)=>s+r.pct,0)/cad.length:0;
  const metrics=[
    rev.length>0?{l:"Total Revenue",v:f$(totalRev),s:"MRR "+fk(totalMRR),col:ACC}:null,
    email.length>0?{l:"Emails Sent",v:totalSent,s:email.length+" senders",col:"#5378FC"}:null,
    email.length>0?{l:"Avg Open Rate",v:fp(avgOpen),s:"Target 70%+",col:avgOpen>=0.7?GRN:YLW}:null,
    cad.length>0?{l:"Cadence",v:fp(avgCad),s:"Target 90%+",col:avgCad>=0.9?GRN:avgCad>=0.5?YLW:RED}:null,
  ].filter(Boolean);
  const teamRev={};
  rev.forEach(r=>{if(r.team) teamRev[r.team]=(teamRev[r.team]||0)+r.total;});
  const trs=Object.entries(teamRev).sort((a,b)=>b[1]-a[1]);
  const maxR=trs[0]&&trs[0][1]||1;
  const emSort=[...email].sort((a,b)=>b.sent-a.sent).slice(0,10);
  const maxS=emSort[0]&&emSort[0].sent||1;
  const cadSort=[...cad].filter(r=>r.count>0).sort((a,b)=>b.pct-a.pct);
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
          {trs.length===0?<div style={{color:MUT,fontSize:13}}>No data</div>:trs.map(([t,v])=>(
            <div key={t} style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,fontSize:12}}>
              <span style={{width:140,flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                <span style={{display:"inline-block",width:6,height:6,borderRadius:"50%",background:tc(t),marginRight:6,verticalAlign:"middle"}}/>
                {st(t)}
              </span>
              <div style={{flex:1,height:5,background:"#ECEEF1",borderRadius:3,overflow:"hidden"}}>
                <div style={{width:(v/maxR*100).toFixed(1)+"%",height:"100%",background:tc(t),opacity:.85,borderRadius:3}}/>
              </div>
              <span style={{width:56,textAlign:"right",fontSize:11,color:MUT,fontWeight:600,flexShrink:0}}>{fk(v)}</span>
            </div>
          ))}
        </div>
        <div style={card}>
          <div style={{fontSize:11,textTransform:"uppercase",letterSpacing:".1em",color:MUT,fontWeight:700,marginBottom:14}}>Email Activity — Top 10</div>
          {emSort.length===0?<div style={{color:MUT,fontSize:13}}>No data</div>:emSort.map(r=>(
            <div key={r.name} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,fontSize:12}}>
              <span style={{width:120,flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.name}</span>
              <div style={{flex:1,height:5,background:"#ECEEF1",borderRadius:3,overflow:"hidden"}}>
                <div style={{width:(r.sent/maxS*100).toFixed(1)+"%",height:"100%",background:"#5378FC",opacity:.8,borderRadius:3}}/>
              </div>
              <span style={{width:34,textAlign:"right",fontSize:11,color:MUT,fontWeight:600,flexShrink:0}}>{r.sent}</span>
              <span style={{width:42,textAlign:"right",fontSize:11,fontWeight:700,flexShrink:0,color:pCol(r.openRate)}}>{fp(r.openRate)}</span>
            </div>
          ))}
        </div>
      </div>
      <div style={card}>
        <div style={{fontSize:11,textTransform:"uppercase",letterSpacing:".1em",color:MUT,fontWeight:700,marginBottom:14}}>Cadence Completion</div>
        {cadSort.length===0?<div style={{color:MUT,fontSize:13}}>No data</div>:
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"4px 32px"}}>
            {cadSort.map(r=>{
              const p=Math.min(r.pct,1),col=bCol(p,0.9,0.5);
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
      style={{fontSize:10,textTransform:"uppercase",color:MUT,fontWeight:700,padding:"0 0 10px",textAlign:"right",cursor:"pointer",borderBottom:"1px solid rgba(41,53,93,.08)"}}>
      {lbl}{sort.col===col?(sort.dir==="desc"?" ▼":" ▲"):""}
    </th>
  );
  return (
    <div style={{padding:"24px 28px"}}>
      <div style={card}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
          <thead><tr>
            <th style={{width:32,fontSize:10,color:MUT,fontWeight:700,padding:"0 0 10px",textAlign:"left",borderBottom:"1px solid rgba(41,53,93,.08)"}}>#</th>
            <th style={{fontSize:10,textTransform:"uppercase",color:MUT,fontWeight:700,padding:"0 0 10px",textAlign:"left",borderBottom:"1px solid rgba(41,53,93,.08)"}}>CSM</th>
            <th style={{fontSize:10,textTransform:"uppercase",color:MUT,fontWeight:700,padding:"0 0 10px",textAlign:"left",borderBottom:"1px solid rgba(41,53,93,.08)"}}>Team</th>
            {th("rev","Revenue")}
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
                <td style={{padding:"10px 8px 10px 0",borderBottom:"1px solid rgba(41,53,93,.06)"}}><span style={{fontWeight:700}}>{c.name}</span></td>
                <td style={{padding:"10px 8px 10px 0",borderBottom:"1px solid rgba(41,53,93,.06)"}}>
                  <span style={{display:"inline-block",width:6,height:6,borderRadius:"50%",background:color,marginRight:6,verticalAlign:"middle"}}/>
                  <span style={{color:MUT,fontSize:12}}>{st(info.t||c.team)}</span>
                </td>
                <td style={{padding:"10px 0",borderBottom:"1px solid rgba(41,53,93,.06)",textAlign:"right",fontWeight:700,color:ACC}}>{c.rev>0?f$(c.rev):"—"}</td>
                <td style={{padding:"10px 0",borderBottom:"1px solid rgba(41,53,93,.06)",textAlign:"right"}}>{c.sent>0?c.sent:"—"}</td>
                <td style={{padding:"10px 0",borderBottom:"1px solid rgba(41,53,93,.06)",textAlign:"right",fontWeight:700,color:pCol(c.openRate)}}>{c.sent>0?fp(c.openRate):"—"}</td>
                <td style={{padding:"10px 0",borderBottom:"1px solid rgba(41,53,93,.06)",textAlign:"right",fontWeight:700,color:pCol(c.cadPct)}}>{c.cadPct>0?fp(c.cadPct):"—"}</td>
              </tr>
            );
          })}</tbody>
        </table>
        {filtered.length>15&&(
          <button onClick={()=>setAll(s=>!s)} style={{display:"block",width:"100%",marginTop:14,background:"transparent",border:"1px solid rgba(41,53,93,.15)",color:MUT,fontSize:12,fontWeight:600,padding:8,borderRadius:8,cursor:"pointer"}}>
            {all?"Show top 15 ↑":"Show all "+filtered.length+" ↓"}
          </button>
        )}
      </div>
    </div>
  );
}

function ActivityView({email,cad}) {
  const emSort=[...email].sort((a,b)=>b.sent-a.sent);
  const cadSort=[...cad].sort((a,b)=>b.pct-a.pct);
  return (
    <div style={{padding:"24px 28px"}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20}}>
        <div style={card}>
          <div style={{fontSize:11,textTransform:"uppercase",color:MUT,fontWeight:700,marginBottom:14}}>Email Performance</div>
          {emSort.length===0?<div style={{color:MUT,fontSize:13}}>No data</div>:
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead><tr>{["CSM","Sent","Open %","Reply %"].map(h=><th key={h} style={{fontSize:10,textTransform:"uppercase",color:MUT,fontWeight:700,padding:"0 0 8px",textAlign:h==="CSM"?"left":"right",borderBottom:"1px solid rgba(41,53,93,.08)"}}>{h}</th>)}</tr></thead>
              <tbody>{emSort.map(r=>(
                <tr key={r.name}>
                  <td style={{padding:"7px 0",borderBottom:"1px solid rgba(41,53,93,.05)"}}>{r.name}</td>
                  <td style={{padding:"7px 0",borderBottom:"1px solid rgba(41,53,93,.05)",textAlign:"right"}}>{r.sent}</td>
                  <td style={{padding:"7px 0",borderBottom:"1px solid rgba(41,53,93,.05)",textAlign:"right",fontWeight:700,color:pCol(r.openRate)}}>{fp(r.openRate)}</td>
                  <td style={{padding:"7px 0",borderBottom:"1px solid rgba(41,53,93,.05)",textAlign:"right",fontWeight:700,color:pCol(r.replyRate)}}>{fp(r.replyRate)}</td>
                </tr>
              ))}</tbody>
            </table>}
        </div>
        <div style={card}>
          <div style={{fontSize:11,textTransform:"uppercase",color:MUT,fontWeight:700,marginBottom:14}}>Cadence Touchpoints</div>
          {cadSort.length===0?<div style={{color:MUT,fontSize:13}}>No data</div>:
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead><tr>{["CSM","Tasks","Done"].map(h=><th key={h} style={{fontSize:10,textTransform:"uppercase",color:MUT,fontWeight:700,padding:"0 0 8px",textAlign:h==="CSM"?"left":"right",borderBottom:"1px solid rgba(41,53,93,.08)"}}>{h}</th>)}</tr></thead>
              <tbody>{cadSort.map(r=>{
                const p=Math.min(r.pct,1);
                const bg=p>=0.9?"rgba(22,163,74,.1)":p>=0.5?"rgba(217,119,6,.1)":"rgba(220,38,38,.1)";
                const fg=p>=0.9?"#166534":p>=0.5?"#854d0e":"#991b1b";
                return (
                  <tr key={r.name}>
                    <td style={{padding:"7px 0",borderBottom:"1px solid rgba(41,53,93,.05)"}}>{r.name}</td>
                    <td style={{padding:"7px 0",borderBottom:"1px solid rgba(41,53,93,.05)",textAlign:"right"}}>{r.count}</td>
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

function AIPanel({csms,rev,email,cad,open,onClose}) {
  const [msgs,setMsgs]=useState([{role:"assistant",text:"Hi! I have your team data loaded. Ask me who needs coaching, which team is winning, or anything about performance."}]);
  const [inp,setInp]=useState("");
  const [busy,setBusy]=useState(false);
  const [hist,setHist]=useState([]);
  const ref=useRef();
  useEffect(()=>{ if(ref.current) ref.current.scrollTop=ref.current.scrollHeight; },[msgs]);
  const ctx=()=>{
    const sums=COACHES.map(coach=>{
      const tc2=csms.filter(c=>{ const i=lk(c.name); return(i&&i.c===coach.email)||c.coach===coach.email; });
      if(!tc2.length) return coach.name+": no data";
      const cadC=tc2.filter(c=>c.cadCount>0);
      const avgC=cadC.length?cadC.reduce((s,c)=>s+c.cadPct,0)/cadC.length:null;
      const emC=tc2.filter(c=>c.sent>0);
      const avgO=emC.length?emC.reduce((s,c)=>s+c.openRate,0)/emC.length:null;
      const totR=tc2.reduce((s,c)=>s+c.rev,0);
      const win=cadC.filter(c=>c.cadPct>=0.9).map(c=>c.name);
      const str=cadC.filter(c=>c.cadPct<0.9).sort((a,b)=>a.cadPct-b.cadPct).map(c=>c.name+" "+Math.round(c.cadPct*100)+"%");
      const nd=tc2.filter(c=>c.cadCount===0).map(c=>c.name);
      return "COACH: "+coach.name+" | "+coach.team+
        "\n  Cadence: "+(avgC!=null?Math.round(avgC*100)+"%":"no data")+
        "\n  Email open: "+(avgO!=null?Math.round(avgO*100)+"%":"no data")+
        "\n  Revenue: $"+totR.toLocaleString()+
        "\n  Winning: "+(win.join(", ")||"none")+
        "\n  Needs coaching: "+(str.join(", ")||"none")+
        (nd.length?"\n  No data: "+nd.join(", "):"");
    }).join("\n\n");
    return "You are a CSM coaching assistant for Thryv. Give specific, actionable advice. Winning = 90%+ cadence. Be direct.\n\n"+sums;
  };
  const send=async()=>{
    const msg=inp.trim(); if(!msg||busy) return;
    setInp(""); setBusy(true);
    setMsgs(m=>[...m,{role:"user",text:msg}]);
    const newH=[...hist,{role:"user",content:msg}];
    setHist(newH);
    try {
      const res=await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:800,system:ctx(),messages:newH.slice(-8)})
      });
      const d=await res.json();
      const reply=d.content&&d.content[0]?d.content[0].text:"No response.";
      setMsgs(m=>[...m,{role:"assistant",text:reply}]);
      setHist(h=>[...h,{role:"assistant",content:reply}]);
    } catch(e) {
      setMsgs(m=>[...m,{role:"assistant",text:"Error. Please try again."}]);
    }
    setBusy(false);
  };
  if(!open) return null;
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(41,53,93,.4)",zIndex:500,display:"flex",alignItems:"flex-end",justifyContent:"flex-end",padding:20}}>
      <div style={{width:420,height:"72vh",maxHeight:680,background:"#fff",borderRadius:16,display:"flex",flexDirection:"column",overflow:"hidden",boxShadow:"0 8px 40px rgba(41,53,93,.2)"}}>
        <div style={{padding:"14px 18px",background:MID,display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
          <span style={{fontFamily:"Nunito,sans-serif",fontSize:14,fontWeight:800,color:"#fff"}}>🤖 AI Coaching Assistant</span>
          <button onClick={onClose} style={{background:"transparent",border:"none",color:"rgba(255,255,255,.6)",fontSize:20,cursor:"pointer"}}>✕</button>
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
          {busy&&<div style={{alignSelf:"flex-start",padding:"10px 14px",borderRadius:"12px 12px 12px 3px",background:"#F4F6FB",fontSize:13,color:MUT,fontStyle:"italic"}}>Thinking...</div>}
        </div>
        <div style={{padding:"12px 14px",borderTop:"1px solid rgba(41,53,93,.08)",flexShrink:0}}>
          <div style={{display:"flex",gap:8,alignItems:"flex-end"}}>
            <textarea value={inp} onChange={e=>setInp(e.target.value)}
              onKeyDown={e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();send();}}}
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

function PinLock({onUnlock}) {
  const [val,setVal]=useState("");
  const [err,setErr]=useState(false);
  const check=()=>{
    if(val===PIN){try{sessionStorage.setItem(PIN_KEY,"1");}catch(e){}onUnlock();}
    else{setErr(true);setVal("");setTimeout(()=>setErr(false),1500);}
  };
  return (
    <div style={{minHeight:"100vh",background:"#F4F6FB",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{background:"#fff",borderRadius:16,padding:"40px 36px",boxShadow:"0 4px 32px rgba(41,53,93,.12)",textAlign:"center",width:340}}>
        <div style={{width:56,height:56,borderRadius:"50%",background:MID,display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 20px",fontSize:24}}>🔒</div>
        <div style={{fontFamily:"Nunito,sans-serif",fontSize:20,fontWeight:800,color:MID,marginBottom:6}}>CSM Coaching Dashboard</div>
        <div style={{fontSize:13,color:MUT,marginBottom:24}}>Enter your team PIN to continue</div>
        <input type="password" value={val} onChange={e=>setVal(e.target.value)}
          onKeyDown={e=>e.key==="Enter"&&check()} placeholder="Enter PIN" autoFocus
          style={{width:"100%",padding:"10px 14px",fontSize:15,borderRadius:10,border:"1px solid "+(err?RED:"rgba(41,53,93,.2)"),outline:"none",textAlign:"center",letterSpacing:4,marginBottom:12}}/>
        {err&&<div style={{color:RED,fontSize:12,marginBottom:10}}>Incorrect PIN</div>}
        <button onClick={check} style={{width:"100%",padding:11,background:ACC,border:"none",borderRadius:10,color:"#fff",fontSize:14,fontWeight:700,cursor:"pointer"}}>Unlock Dashboard</button>
        <div style={{fontSize:11,color:MUT,marginTop:16}}>Contact your coach for the PIN</div>
      </div>
    </div>
  );
}

// ── MAIN APP ──────────────────────────────────────────────
export default function App() {
  const [unlocked,setUnlocked]=useState(()=>{ try{return sessionStorage.getItem(PIN_KEY)==="1";}catch(e){return false;} });
  const [rev,setRev]=useState([]);
  const [email,setEmail]=useState([]);
  const [cad,setCad]=useState([]);
  const [tab,setTab]=useState("coaching");
  const [filters,setFilters]=useState({coach:"",team:"",tier:""});
  const [aiOpen,setAiOpen]=useState(false);
  const [status,setStatus]=useState("loading"); // loading | ok | error
  const [updatedAt,setUpdatedAt]=useState(null);

  const csms = buildCSMs(rev,email,cad);
  const hasData = rev.length>0||email.length>0||cad.length>0;

  useEffect(()=>{
    if(!unlocked) return;
    setStatus("loading");
    Promise.all([
      fetchCSV(CSV_REV).then(rows=>{ const d=mapRev(rows); console.log("Rev rows:",d.length); return d; }),
      fetchCSV(CSV_EMAIL).then(rows=>{ const d=mapEmail(rows); console.log("Email rows:",d.length); return d; }),
      fetchCSV(CSV_CAD).then(rows=>{ const d=mapCadence(rows); console.log("Cad rows:",d.length); return d; }),
    ]).then(([r,e,c])=>{
      setRev(r); setEmail(e); setCad(c);
      setUpdatedAt(new Date().toLocaleTimeString());
      setStatus("ok");
      if(r.length>0||e.length>0||c.length>0) setTab("coaching");
    }).catch(err=>{
      console.error("Sheet load error:", err);
      setStatus("error");
    });
  },[unlocked]);

  const CTMAP={"odirlm01@thryv.com":"The Dominican Dream Team","chase.boyd@thryv.com":"Boyd Meets World","elizabeth.white@thryv.com":"White Wave Warriors","kendra.morelli@thryv.com":"Team Thryv-More(lli)","trisha.stalnaker@thryv.com":"Team Status Engaged","aaron.taylor@thryv.com":"Team Aurorians"};
  const ALLTEAMS=["The Dominican Dream Team","Boyd Meets World","White Wave Warriors","Team Thryv-More(lli)","Team Status Engaged","Team Aurorians"];
  const TIERS=["CSMI","CSMII","CSMIII","SSMI","SSMII"];
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
          {tabs.map(t=>(
            <button key={t.id} onClick={()=>setTab(t.id)}
              style={{fontSize:12,fontWeight:600,padding:"6px 14px",borderRadius:6,border:"none",cursor:"pointer",
                background:tab===t.id?ACC:"transparent",color:tab===t.id?"#fff":"rgba(255,255,255,.55)"}}>
              {t.label}
            </button>
          ))}
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          {status==="loading"&&<span style={{fontSize:11,padding:"4px 10px",borderRadius:20,background:"rgba(255,255,255,.1)",color:"rgba(255,255,255,.6)"}}>⟳ Loading...</span>}
          {status==="ok"&&<span style={{fontSize:11,padding:"4px 10px",borderRadius:20,background:"rgba(22,163,74,.25)",color:"#86efac"}}>✓ Live{updatedAt?" · "+updatedAt:""}</span>}
          {status==="error"&&<span style={{fontSize:11,padding:"4px 10px",borderRadius:20,background:"rgba(220,38,38,.25)",color:"#fca5a5"}}>✗ Error</span>}
        </div>
      </div>

      {hasData&&(
        <div style={{background:"#fff",borderBottom:"1px solid rgba(41,53,93,.08)",padding:"10px 28px",display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
          <select value={filters.coach} onChange={e=>{ const v=e.target.value; setFilters({coach:v,team:v?CTMAP[v]||"":"",tier:filters.tier}); }}
            style={{fontSize:12,fontWeight:600,color:filters.coach?ACC:MID,background:"#F4F6FB",border:"1px solid "+(filters.coach?ACC:"rgba(41,53,93,.15)"),borderRadius:8,padding:"6px 12px",cursor:"pointer"}}>
            <option value="">All Coaches</option>
            {COACHES.map(c=><option key={c.email} value={c.email}>{c.name}</option>)}
          </select>
          <select value={filters.team} onChange={e=>setFilters(f=>({...f,team:e.target.value}))}
            style={{fontSize:12,fontWeight:600,color:filters.team?ACC:MID,background:"#F4F6FB",border:"1px solid "+(filters.team?ACC:"rgba(41,53,93,.15)"),borderRadius:8,padding:"6px 12px",cursor:"pointer"}}>
            <option value="">All Teams</option>
            {ALLTEAMS.map(t=><option key={t} value={t}>{st(t)}</option>)}
          </select>
          <select value={filters.tier} onChange={e=>setFilters(f=>({...f,tier:e.target.value}))}
            style={{fontSize:12,fontWeight:600,color:filters.tier?ACC:MID,background:"#F4F6FB",border:"1px solid "+(filters.tier?ACC:"rgba(41,53,93,.15)"),borderRadius:8,padding:"6px 12px",cursor:"pointer"}}>
            <option value="">All Tiers</option>
            {TIERS.map(t=><option key={t} value={t}>{t}</option>)}
          </select>
          {Object.values(filters).some(v=>v)&&(
            <button onClick={()=>setFilters({coach:"",team:"",tier:""})} style={{fontSize:11,fontWeight:700,color:ACC,background:"none",border:"none",cursor:"pointer"}}>✕ Clear</button>
          )}
        </div>
      )}

      {status==="loading"&&!hasData&&(
        <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"60vh",flexDirection:"column",gap:16}}>
          <div style={{fontSize:32}}>⟳</div>
          <div style={{fontSize:15,color:MUT}}>Loading data from Google Sheet...</div>
        </div>
      )}

      {status==="error"&&!hasData&&(
        <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"60vh",flexDirection:"column",gap:16}}>
          <div style={{fontSize:32}}>⚠️</div>
          <div style={{fontSize:15,color:RED,fontWeight:700}}>Could not load data from Google Sheet</div>
          <div style={{fontSize:13,color:MUT}}>Check the browser console for details</div>
          <button onClick={()=>window.location.reload()} style={{background:ACC,color:"#fff",border:"none",borderRadius:8,padding:"10px 20px",cursor:"pointer",fontSize:13,fontWeight:700}}>Try Again</button>
        </div>
      )}

      {hasData&&tab==="coaching"&&<CoachingView csms={csms} fc={filters.coach} ft={filters.team}/>}
      {hasData&&tab==="overview"&&<OverviewView rev={rev} email={email} cad={cad} csms={csms}/>}
      {hasData&&tab==="leaderboard"&&<LeaderboardView csms={csms} fc={filters.coach} ft={filters.team} tier={filters.tier}/>}
      {hasData&&tab==="activity"&&<ActivityView email={email} cad={cad}/>}

      {hasData&&(
        <button onClick={()=>setAiOpen(true)}
          style={{position:"fixed",bottom:24,right:24,width:54,height:54,borderRadius:"50%",background:MID,border:"none",cursor:"pointer",boxShadow:"0 4px 20px rgba(41,53,93,.35)",fontSize:22,zIndex:499,display:"flex",alignItems:"center",justifyContent:"center"}}
          onMouseOver={e=>{e.currentTarget.style.background=ACC;}}
          onMouseOut={e=>{e.currentTarget.style.background=MID;}}>
          🤖
        </button>
      )}
      <AIPanel csms={csms} rev={rev} email={email} cad={cad} open={aiOpen} onClose={()=>setAiOpen(false)}/>
    </div>
  );
}
