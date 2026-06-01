import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";

const PIN = "thryv2025";
const PIN_KEY = "csm_pin_v1";

const CSV_REV     = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=322916467&single=true&output=csv";
const CSV_EMAIL   = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=0&single=true&output=csv";
const CSV_CAD     = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=1973544046&single=true&output=csv";
const CSV_DUE     = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=341836664&single=true&output=csv";
const CSV_ONTIME  = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=459845057&single=true&output=csv";

const COACHES = [
  {e:"odirlm01@thryv.com",      n:"Mia O\u2019Dirling",   t:"The Dominican Dream Team", col:"#FF5000"},
  {e:"chase.boyd@thryv.com",    n:"Chase Boyd",            t:"Boyd Meets World",          col:"#4A5D8C"},
  {e:"elizabeth.white@thryv.com",n:"Elizabeth White",      t:"White Wave Warriors",       col:"#808080"},
  {e:"kendra.morelli@thryv.com", n:"Kendra Morelli",       t:"Team Thryv-More(lli)",      col:"#29355D"},
  {e:"trisha.stalnaker@thryv.com",n:"Trisha Stalnaker",    t:"Team Status Engaged",       col:"#E03000"},
  {e:"aaron.taylor@thryv.com",  n:"Aaron Taylor",          t:"Team Aurorians",            col:"#5378FC"},
];

const TEAM_COLS = {
  "The Dominican Dream Team":"#FF5000","Boyd Meets World":"#4A5D8C",
  "White Wave Warriors":"#808080","Team Thryv-More(lli)":"#29355D",
  "Team Status Engaged":"#E03000","Team Aurorians":"#5378FC",
};

// ── ORG CHART (source of truth for coach assignments) ──────────────────────
// Updated from: https://docs.google.com/spreadsheets/d/e/...?gid=352807280
// Key changes vs previous: Dorka Frias Lantigua → Kendra Morelli,
//   Steven Saunders → Kendra Morelli, Sidharta Goris removed from roster

const NAME_NORM = {
  "darling danais":"Darling Danais Santos Taveras","darling taveras":"Darling Danais Santos Taveras","darling danais santos taveras":"Darling Danais Santos Taveras",
  "heidi torres":"Heidi Torres Uribe","heidi uribe":"Heidi Torres Uribe","heidi torres uribe":"Heidi Torres Uribe",
  "irina larianni":"Irina Larianni Molina Molina","irina molina":"Irina Larianni Molina Molina","irina molina molina":"Irina Larianni Molina Molina","irina larianni molina molina":"Irina Larianni Molina Molina",
  "jathzelyn elizabeth":"Jathzelyn Elizabeth Fortuna Paulino","jathzelyn fortuna":"Jathzelyn Elizabeth Fortuna Paulino","jathzelyn fortuna paulino":"Jathzelyn Elizabeth Fortuna Paulino","jathzelyn paulino":"Jathzelyn Elizabeth Fortuna Paulino","jazz fortuna":"Jathzelyn Elizabeth Fortuna Paulino","jathzelyn elizabeth fortuna paulino":"Jathzelyn Elizabeth Fortuna Paulino",
  "johnny cornielle montas":"Johnny Cornielle","johnny cornielle":"Johnny Cornielle",
  "joseph carmona":"Joseph Guillermo Carmona Garcia","joseph carmona garcia":"Joseph Guillermo Carmona Garcia","joseph garcia":"Joseph Guillermo Carmona Garcia","joseph guillermo":"Joseph Guillermo Carmona Garcia","joseph guillermo carmona garcia":"Joseph Guillermo Carmona Garcia",
  "sam frias":"Samuel Frias De Paula","sam frias de paula":"Samuel Frias De Paula","samuel frias":"Samuel Frias De Paula","samuel paula":"Samuel Frias De Paula","samuel frias de paula":"Samuel Frias De Paula",
  "sati ananda":"Sati Ananda Pimentel Malespin","sati malespin":"Sati Ananda Pimentel Malespin","sati pimentel":"Sati Ananda Pimentel Malespin","sati pimentel malespin":"Sati Ananda Pimentel Malespin","sati ananda pimentel malespin":"Sati Ananda Pimentel Malespin",
  "victor abner":"Victor Abner Moscoso Fernandez","victor fernandez":"Victor Abner Moscoso Fernandez","victor moscoso":"Victor Abner Moscoso Fernandez","victorabner moscoso fernandez":"Victor Abner Moscoso Fernandez","victor abner moscoso fernandez":"Victor Abner Moscoso Fernandez",
  "wilson mercedes":"Wilson Mercedes",
  "yessica montero":"Yessica Montero Urena","yessica urena":"Yessica Montero Urena","yessica montero urena":"Yessica Montero Urena",
  "barbara larrosa":"Barbara Larrosa Presinal","barbara presinal":"Barbara Larrosa Presinal","barbara larrosa presinal":"Barbara Larrosa Presinal",
  "deivis pe\u00f1a":"Deivis Pena","deivis pena":"Deivis Pena",
  "eric johnson":"Eric Johnson","kyle dye":"Kyle Dye","sarah swanson":"Sarah Swanson",
  "tyler moeggenberg":"Tyler Moeggenberg","tyler popplewell":"Tyler Popplewell",
  "luis aguasvivas":"Luis Aguasvivas Peralta","luis peralta":"Luis Aguasvivas Peralta","luis aguasvivas peralta":"Luis Aguasvivas Peralta",
  "juan liberato paula":"Juan Liberato","juan liberato":"Juan Liberato",
  "elianny antigua":"Elianny Tena Antigua","elianny tena":"Elianny Tena Antigua","elianny tena antigua":"Elianny Tena Antigua",
  "damita hill":"Damita Hill","anthony yen":"Anthony Yen","april hall":"April Hall",
  "katelyn ankrom":"Katelyn Ankrom","kennedy sanchez":"Kennedy Sanchez","matt sword":"Matt Sword","michael furlong":"Michael Furlong",
  "yolanda ramirez-drake":"Yolanda Ramirez","yolanda ramirez":"Yolanda Ramirez",
  "florence francois":"Florence Francois Nova","florence nova":"Florence Francois Nova","francois nova":"Florence Francois Nova","florence francois nova":"Florence Francois Nova",
  "rossi tejeda":"Rossi Valerio Tejeda","rossi valerio":"Rossi Valerio Tejeda","rossi valerio tejeda":"Rossi Valerio Tejeda",
  "alejandro rodriguez medina":"Alejandro Rodriguez-Medina","alejandro rodriguez-medina":"Alejandro Rodriguez-Medina",
  "chelsea dingus":"Chelsea Dingus",
  "dorka frias":"Dorka Frias Lantigua","dorka lantigua":"Dorka Frias Lantigua","dorka frias lantigua":"Dorka Frias Lantigua",
  "karmita k turner":"Karmita Turner","karmita turner":"Karmita Turner",
  "lauren carter":"Lauren Carter","libby booher":"Libby Booher","misti dixon":"Misti Dixon","misty decatur":"Misty Decatur",
  "saira guzman":"Saira Julian Guzman","saira julian":"Saira Julian Guzman","saira julian guzman":"Saira Julian Guzman",
  "scott mather":"Scott Mather","steven saunders":"Steven Saunders",
  "ashley shaffer":"Ashley Shaffer",
  "ashley mena":"Ashley Vasquez Mena","ashley vasquez":"Ashley Vasquez Mena","ashley vasquez mena":"Ashley Vasquez Mena",
  "karen capellan":"Karen Capellan Tavarez","karen tavarez":"Karen Capellan Tavarez","karen capellan tavarez":"Karen Capellan Tavarez",
  "karissa hernandez":"Karissa Hernandez","kellie lester":"Kellie Lester","mark velazquez":"Mark Velazquez",
  "merve (mj)":"Merve (MJ) Brielmann","merve brielmann":"Merve (MJ) Brielmann","mj brielmann":"Merve (MJ) Brielmann","merve (mj) brielmann":"Merve (MJ) Brielmann",
  "rafael sencion":"Rafael Sencion Sencion","rafael sencion sencion":"Rafael Sencion Sencion",
  "stacy miron":"Stacy Roers","stacy roers":"Stacy Roers",
  "taylor kidd":"Taylor Kidd",
  "felix caba":"Felix Caba Jimenez","felix jimenez":"Felix Caba Jimenez","felix caba jimenez":"Felix Caba Jimenez",
  "dave crisler":"David Crisler","david crisler":"David Crisler",
  "ellise payne":"Ellise Payne","indu vijay":"Indu Vijay","matt daly":"Matt Daly",
  "nikita siepen bowers":"Nikita Siepen-Bowers","nikita siepen-bowers":"Nikita Siepen-Bowers",
  "peter manalac":"Peter Manalac","sakshi mahalwal":"Sakshi Mahalwal",
  "sylvia":"Sylvia Appla","sylvia appla":"Sylvia Appla",
  "tracy ann gaudencio":"Tracy-Ann Gaudencio","tracy-ann gaudencio":"Tracy-Ann Gaudencio",
  "warda gul":"Warda Gul","zoltan rudolf":"Zoltan Rudolf",
};

const ROSTER = {
  // ── Mia O'Dirling — The Dominican Dream Team ──────────────────────────
  "darling danais":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "darling taveras":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "darling danais santos taveras":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "heidi torres":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "heidi uribe":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "heidi torres uribe":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "irina larianni":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "irina molina":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "irina molina molina":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "irina larianni molina molina":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "jathzelyn elizabeth":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "jathzelyn fortuna":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "jathzelyn fortuna paulino":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "jathzelyn paulino":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "jazz fortuna":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "jathzelyn elizabeth fortuna paulino":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "johnny cornielle":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII"},
  "johnny cornielle montas":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII"},
  "joseph carmona":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "joseph carmona garcia":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "joseph garcia":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "joseph guillermo":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "joseph guillermo carmona garcia":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "sam frias":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII"},
  "sam frias de paula":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII"},
  "samuel frias":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII"},
  "samuel paula":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII"},
  "samuel frias de paula":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII"},
  "sati ananda":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "sati malespin":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "sati pimentel":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "sati pimentel malespin":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "sati ananda pimentel malespin":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "victor abner":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "victor fernandez":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "victor moscoso":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "victorabner moscoso fernandez":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "victor abner moscoso fernandez":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "wilson mercedes":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII"},
  "yessica montero":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "yessica urena":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "yessica montero urena":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  // ── Chase Boyd — Boyd Meets World ─────────────────────────────────────
  "barbara larrosa":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMII"},
  "barbara presinal":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMII"},
  "barbara larrosa presinal":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMII"},
  "deivis pena":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "deivis pe\u00f1a":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "eric johnson":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMII"},
  "kyle dye":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMIII"},
  "sarah swanson":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "tyler moeggenberg":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "tyler popplewell":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "luis aguasvivas":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "luis peralta":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "luis aguasvivas peralta":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "juan liberato paula":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "juan liberato":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "elianny antigua":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "elianny tena":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "elianny tena antigua":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  // ── Elizabeth White — White Wave Warriors ─────────────────────────────
  "damita hill":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII"},
  "anthony yen":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII"},
  "april hall":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII"},
  "katelyn ankrom":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII"},
  "kennedy sanchez":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII"},
  "matt sword":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI"},
  "michael furlong":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII"},
  "yolanda ramirez-drake":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII"},
  "yolanda ramirez":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII"},
  "florence francois":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI"},
  "florence nova":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI"},
  "francois nova":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI"},
  "florence francois nova":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI"},
  "rossi tejeda":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI"},
  "rossi valerio":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI"},
  "rossi valerio tejeda":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI"},
  // ── Kendra Morelli — Team Thryv-More(lli) ────────────────────────────
  "alejandro rodriguez medina":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"SSMII"},
  "alejandro rodriguez-medina":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"SSMII"},
  "chelsea dingus":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"SSMII"},
  "dorka frias":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMI"},
  "dorka lantigua":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMI"},
  "dorka frias lantigua":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMI"},
  "karmita k turner":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMII"},
  "karmita turner":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMII"},
  "lauren carter":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMII"},
  "libby booher":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMIII"},
  "misti dixon":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMIII"},
  "misty decatur":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"SSMI"},
  "saira guzman":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMI"},
  "saira julian":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMI"},
  "saira julian guzman":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMI"},
  "scott mather":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"SSMII"},
  "steven saunders":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMII"},
  // ── Trisha Stalnaker — Team Status Engaged ────────────────────────────
  "ashley shaffer":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMII"},
  "ashley mena":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "ashley vasquez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "ashley vasquez mena":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "karen capellan":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "karen tavarez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "karen capellan tavarez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "karissa hernandez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMII"},
  "kellie lester":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMIII"},
  "mark velazquez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"SSMII"},
  "merve (mj)":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMII"},
  "merve brielmann":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMII"},
  "mj brielmann":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMII"},
  "merve (mj) brielmann":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMII"},
  "rafael sencion":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "rafael sencion sencion":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "stacy miron":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMIII"},
  "stacy roers":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMIII"},
  "taylor kidd":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "felix caba":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "felix jimenez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "felix caba jimenez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  // ── Aaron Taylor — Team Aurorians ─────────────────────────────────────
  "dave crisler":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMII"},
  "david crisler":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMII"},
  "ellise payne":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "indu vijay":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "matt daly":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "nikita siepen bowers":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "nikita siepen-bowers":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "peter manalac":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "sakshi mahalwal":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "sylvia":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "sylvia appla":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "tracy ann gaudencio":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "tracy-ann gaudencio":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "warda gul":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "zoltan rudolf":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMII"},
};

function lk(n) { return n ? ROSTER[n.toLowerCase().trim()] || null : null; }
function norm(n) { return NAME_NORM[n.toLowerCase().trim()] || n.trim(); }
function pm(v) { const x = parseFloat(String(v||0).replace(/[$,]/g,"")); return isNaN(x) ? 0 : x; }
function pn(v) { const x = parseFloat(String(v||0).replace(/,/g,"")); return isNaN(x) ? 0 : x; }
function pp(v) { return isNaN(v)||v==null ? "--" : (v*100).toFixed(1)+"%"; }
function fd(n) { return "$"+Number(n).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2}); }
function fk(n) { return n>=1000 ? "$"+(n/1000).toFixed(1).replace(/\.0$/,"")+"k" : "$"+Math.round(n); }
function bc(v,hi,lo) { return v>=hi ? "#16a34a" : v>=lo ? "#d97706" : "#dc2626"; }
function pc(p) { return p>=0.9 ? "#16a34a" : p>=0.5 ? "#d97706" : p>0 ? "#dc2626" : "#888"; }
function st(t) { return (t||"").replace("The ","").replace("Team ",""); }

// ── CSV PARSER ─────────────────────────────────────────────────────────────
function parseCSVLine(line) {
  const result = []; let cur = "", inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') { if (inQ && line[i+1]==='"'){cur+='"';i++;} else inQ=!inQ; }
    else if (ch===',' && !inQ) { result.push(cur.trim()); cur=""; }
    else cur += ch;
  }
  result.push(cur.trim()); return result;
}

async function fetchCSV(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error("HTTP "+res.status);
  const text = await res.text();
  if (!text||text.includes("<!DOCTYPE")) throw new Error("bad response");
  const lines = text.split("\n").filter(l=>l.trim());
  if (lines.length < 2) return [];
  const headers = parseCSVLine(lines[0]);
  const rows = [];
  for (let i=1; i<lines.length; i++) {
    const vals = parseCSVLine(lines[i]);
    if (!vals.some(v=>v)) continue;
    const obj = {};
    headers.forEach((h,j) => { obj[h.trim()] = (vals[j]||"").trim(); });
    rows.push(obj);
  }
  console.log("CSV from "+url.slice(-30)+":", rows.length, "rows");
  return rows;
}

// ── DATA MAPPERS ───────────────────────────────────────────────────────────
function mapRev(rows) {
  const by = {};
  rows.forEach(r => {
    const raw = r["CSM Name"]||r["csm name"]||"";
    const name = norm(raw.trim());
    if (!name) return;
    if (!by[name]) by[name] = {name, team:r["CSM Team! "]||r["CSM Team!"]||"", mrr:0, otr:0, total:0, nonrev:0, accts:[]};
    by[name].mrr   += pm(r["MRR $ Added"]||0);
    by[name].otr   += pm(r["OTR $ Added"]||0);
    by[name].total += pm(r["Total Revenue Added"]||0);
    if ((r["Non-Revenue Integrations"]||"").trim()) by[name].nonrev++;
    const biz = (r["Business Name"]||"").trim();
    if (biz) by[name].accts.push({
      b: biz,
      t: r["Type of Integration"]||"",
      m: pm(r["MRR $ Added"]||0),
      o: pm(r["OTR $ Added"]||0),
      n: r["Non-Revenue Integrations"]||"",
    });
  });
  return Object.values(by);
}

function mapEmail(rows) {
  const by = {};
  rows.forEach(r => {
    const raw = r["Touchpoint: Owner Name \u2191"]||r["Touchpoint: Owner Name"]||r["Name"]||"";
    const name = norm(raw.trim());
    if (!name||name==="Total") return;
    if (!by[name]) by[name] = {name, sent:0, uniqueOpens:0, replies:0};
    by[name].sent        += pn(r["Sum of Emails Sent"]||r["Emails Sent"]||0);
    by[name].uniqueOpens += pn(r["Sum of Unique Opens"]||r["Unique Opens"]||0);
    by[name].replies     += pn(r["Sum of Email Replies"]||r["Email Replies"]||0);
  });
  return Object.values(by).map(d => ({
    name: d.name,
    sent: d.sent,
    openRate: d.sent > 0 ? Math.min(d.uniqueOpens/d.sent, 1) : 0,
    replyRate: d.sent > 0 ? Math.min(d.replies/d.sent, 1) : 0,
  }));
}

function mapCadence(rows) {
  const by = {};
  rows.forEach(r => {
    const raw = r["Touchpoint: Owner Name \u2191"]||r["Touchpoint: Owner Name"]||r["name"]||r["Name"]||"";
    const name = norm(raw.trim());
    const status = (r["Cadence Member: Status"]||r["Status"]||"").trim();
    if (!name||name==="Total") return;
    if (!by[name]) by[name] = {name, total:0, completed:0, removed:0, pctField:null};
    by[name].total++;
    if (status==="Removed") by[name].removed++;
    else by[name].completed++;
    const pv = r["% Completed"]||r["pct"]||"";
    if (pv && by[name].pctField===null) by[name].pctField = pn(pv);
  });
  return Object.values(by).map(d => {
    let pct = d.pctField !== null ? d.pctField : (d.total>0 ? (d.total-d.removed)/d.total : 0);
    if (pct > 1) pct = pct/100;
    return {name:d.name, total:d.total, pct};
  }).filter(d=>d.total>0);
}

function mapDue(rows) {
  const by = {};
  rows.forEach(r => {
    const raw = r["Cadence Member: Assigned"]||r["Assigned"]||r["CSM"]||"";
    const name = norm(raw.trim());
    if (!name) return;
    if (!by[name]) by[name] = {name, due:0, overdue:0, newToday:0};
    by[name].due++;
    if ((r["Overdue"]||"").trim()==="1") by[name].overdue++;
    if ((r["New Today"]||"").trim()==="1") by[name].newToday++;
  });
  return Object.values(by);
}

function mapOnTime(rows) {
  const by = {};
  rows.forEach(r => {
    const raw = r["Cadence Member: Assigned"]||r["Assigned"]||r["CSM"]||"";
    const name = norm(raw.trim());
    if (!name) return;
    if (!by[name]) by[name] = {name, total:0, onTime:0};
    by[name].total++;
    if ((r["Completed On Time?"]||r["On Time"]||"").trim()==="1") by[name].onTime++;
  });
  return Object.values(by).map(d => ({
    name: d.name,
    total: d.total,
    onTime: d.onTime,
    pct: d.total > 0 ? d.onTime/d.total : 0,
  })).filter(d=>d.total>=1);
}

// ── BUILD UNIFIED CSM LIST ─────────────────────────────────────────────────
function buildCSMs(rev, email, cad, due, ontime) {
  const m = {};
  const get = name => {
    if (!m[name]) {
      const i = lk(name);
      m[name] = {name, team:(i&&i.t)||"", coach:(i&&i.c)||"", tier:(i&&i.r)||"",
        rev:0, mrr:0, ints:0, accts:[],
        sent:0, openRate:0, replyRate:0,
        cadCount:0, cadPct:0,
        dueCount:0, overdueCount:0, newToday:0,
        otTotal:0, otOnTime:0, otPct:null};
    }
    return m[name];
  };
  (rev||[]).forEach(d => {
    const c = get(d.name);
    if (!c.team && d.team) c.team = d.team.trim();
    if (!c.coach) { const i=lk(d.name); if(i) c.coach=i.c; }
    c.rev += d.total; c.mrr += d.mrr;
    if (d.total>0||d.nonrev>0) c.ints++;
    c.accts = c.accts.concat(d.accts||[]);
  });
  (email||[]).forEach(d => {
    const c = get(d.name);
    c.sent=d.sent; c.openRate=d.openRate; c.replyRate=d.replyRate;
  });
  (cad||[]).forEach(d => {
    const c = get(d.name);
    c.cadCount=d.total; c.cadPct=d.pct;
  });
  (due||[]).forEach(d => {
    const c = get(d.name);
    c.dueCount=d.due; c.overdueCount=d.overdue; c.newToday=d.newToday;
  });
  (ontime||[]).forEach(d => {
    const c = get(d.name);
    c.otTotal=d.total; c.otOnTime=d.onTime; c.otPct=d.pct;
  });
  return Object.values(m);
}

// ── STYLES ─────────────────────────────────────────────────────────────────
const S = {
  card: {background:"#fff",border:"0.5px solid rgba(41,53,93,.09)",borderRadius:12,padding:16},
  nav:  {background:"#29355D"},
};

// ── BAR COMPONENT ──────────────────────────────────────────────────────────
function Bar({label,val,hi,lo}) {
  const p = val!=null ? Math.min(val*100,100) : 0;
  const col = val!=null ? bc(val,hi,lo) : "#e5e7eb";
  return (
    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:5}}>
      <span style={{width:65,flexShrink:0,fontSize:10,fontWeight:500,color:"#808080",textTransform:"uppercase"}}>{label}</span>
      <div style={{flex:1,height:5,background:"#ECEEF1",borderRadius:3,overflow:"hidden"}}>
        <div style={{width:p.toFixed(1)+"%",height:"100%",background:col,borderRadius:3}}/>
      </div>
      <span style={{width:30,textAlign:"right",fontSize:11,fontWeight:500,color:col,flexShrink:0}}>
        {val!=null ? Math.round(val*100)+"%" : "--"}
      </span>
    </div>
  );
}

// ── COACH CARD ─────────────────────────────────────────────────────────────
function CoachCard({coach, csms, onSelectCSM, onSelectCoach}) {
  const team = csms.filter(c => (lk(c.name)&&lk(c.name).c===coach.e)||c.coach===coach.e);
  const cadC = team.filter(c=>c.cadCount>0);
  const emC  = team.filter(c=>c.sent>0);
  const otC  = team.filter(c=>c.otTotal>=3);
  const avgCad  = cadC.length ? cadC.reduce((s,c)=>s+c.cadPct,0)/cadC.length : null;
  const avgOpen = emC.length  ? emC.reduce((s,c)=>s+c.openRate,0)/emC.length : null;
  const avgOT   = otC.length  ? otC.reduce((s,c)=>s+c.otPct,0)/otC.length : null;
  const revPct  = team.length ? team.filter(c=>c.rev>0||c.ints>0).length/team.length : null;
  const teamRev = team.reduce((s,c)=>s+c.rev,0);
  let sc=0,sf=0;
  if(cadC.length>0&&avgCad!=null){sc+=Math.min(avgCad/0.9,1)*30;sf++;}
  if(avgOpen!=null){sc+=Math.min(avgOpen/0.7,1)*25;sf++;}
  if(avgOT!=null){sc+=Math.min(avgOT/0.8,1)*20;sf++;}
  if(revPct!=null){sc+=Math.min(revPct/0.7,1)*25;sf++;}
  const score = sf ? Math.round(Math.min(sc,100)) : null;
  const cls = score==null?"none":score>=75?"win":score>=50?"warn":"att";
  const sCol = {win:"#16a34a",warn:"#d97706",att:"#dc2626",none:"#888"}[cls];
  const topCol = {win:"#16a34a",warn:"#d97706",att:"#dc2626",none:"#e5e7eb"}[cls];
  const wins=cadC.filter(c=>c.cadPct>=0.9).length;
  const warns=cadC.filter(c=>c.cadPct>=0.5&&c.cadPct<0.9).length;
  const atts=cadC.filter(c=>c.cadPct>0&&c.cadPct<0.9).length;
  const sorted=[...team].sort((a,b)=>(a.cadPct||0)-(b.cadPct||0)).slice(0,7);
  return (
    <div style={{...S.card,position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:0,left:0,right:0,height:3,background:topCol}}/>
      <div style={{fontSize:13,fontWeight:500,color:"#29355D",marginTop:4,cursor:"pointer"}} onClick={()=>onSelectCoach(coach.e)}>
        {coach.n}
      </div>
      <div style={{fontSize:11,fontWeight:500,color:coach.col,marginBottom:6}}>{coach.t}</div>
      <div style={{fontSize:11,color:"#808080",marginBottom:8}}>Revenue: <strong style={{color:"#FF5000"}}>{fk(teamRev)}</strong></div>
      <div style={{display:"flex",alignItems:"baseline",gap:6,marginBottom:8}}>
        <span style={{fontSize:40,fontWeight:500,lineHeight:1,color:sCol}}>{score!=null?score:"--"}</span>
        <div style={{fontSize:11,color:"#808080",lineHeight:1.4}}>/ 100<br/><span style={{fontSize:10}}>{wins}✓ {warns}⚠ {atts}✗</span></div>
      </div>
      <Bar label="Cadence"  val={avgCad}  hi={0.9} lo={0.5}/>
      <Bar label="On-time"  val={avgOT}   hi={0.8} lo={0.6}/>
      <Bar label="Email"    val={avgOpen}  hi={0.7} lo={0.35}/>
      <Bar label="Revenue"  val={revPct}  hi={0.7} lo={0.4}/>
      <div style={{height:.5,background:"rgba(41,53,93,.07)",margin:"10px 0"}}/>
      <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:6}}>CSM Status</div>
      {sorted.map(c => {
        const hc = c.cadCount>0;
        const bdgTxt = hc?(c.cadPct>=0.9?"Win":c.cadPct>=0.5?"Watch":"Coach"):"No tasks";
        const bdgCol = hc?(c.cadPct>=0.9?"rgba(22,163,74,.12)":c.cadPct>=0.5?"rgba(217,119,6,.12)":"rgba(220,38,38,.12)"):"rgba(128,128,128,.1)";
        const bdgFg  = hc?(c.cadPct>=0.9?"#166534":c.cadPct>=0.5?"#854d0e":"#991b1b"):"#808080";
        return (
          <div key={c.name} style={{display:"flex",alignItems:"center",gap:6,padding:"3px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",cursor:"pointer"}} onClick={()=>onSelectCSM(c.name)}>
            <span style={{flex:1,fontSize:11,fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"#121212"}}>{c.name}</span>
            <span style={{fontSize:11,fontWeight:500,width:28,textAlign:"right",flexShrink:0,color:hc?pc(c.cadPct):"#888"}}>{hc?Math.round(c.cadPct*100)+"%":"--"}</span>
            <span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:bdgCol,color:bdgFg,flexShrink:0}}>{bdgTxt}</span>
            {c.overdueCount>0&&<span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:"rgba(220,38,38,.1)",color:"#991b1b",flexShrink:0}}>{c.overdueCount}!</span>}
            {c.otTotal>=3&&<span style={{fontSize:10,color:pc(c.otPct),marginLeft:2,flexShrink:0}}>{Math.round(c.otPct*100)}%⏱</span>}
          </div>
        );
      })}
      {team.length>7&&<div style={{fontSize:11,color:"#808080",textAlign:"center",marginTop:6}}>+{team.length-7} more</div>}
    </div>
  );
}

// ── CSM DETAIL VIEW ────────────────────────────────────────────────────────
function CSMDetail({csm, onClear}) {
  const i = lk(csm.name)||{};
  const coach = COACHES.find(c=>c.e===(i.c||csm.coach));
  const ot = csm.otTotal>=1 ? csm : null;
  const totalAcctRev = csm.accts.reduce((s,a)=>s+a.m+a.o,0);

  const wins=[], atts=[];
  if(csm.rev>0) { wins.push("Revenue this period: "+fd(csm.rev)+(csm.mrr>0?" (MRR "+fd(csm.mrr)+")":" ")); }
  else { atts.push("No revenue this period — identify upsell or upgrade opportunities"); }
  if(csm.sent>0) {
    if(csm.openRate>=0.7) wins.push("Excellent email open rate of "+pp(csm.openRate)+" — above 70% target");
    else if(csm.openRate>=0.5) atts.push("Email open rate "+pp(csm.openRate)+" is below target — review subject lines");
    else atts.push("Low email open rate of "+pp(csm.openRate)+" — cadence emails need refreshing");
    if(csm.replyRate>0.15) wins.push("Strong reply rate of "+pp(csm.replyRate)+" — clients engaging");
  } else { atts.push("No email activity — confirm cadence is active"); }
  if(ot&&ot.otTotal>=3) {
    if(ot.otPct>=0.8) wins.push("On-time cadence rate of "+pp(ot.otPct)+" across "+ot.otTotal+" tasks");
    else if(ot.otPct>=0.6) atts.push("On-time rate of "+pp(ot.otPct)+" — some tasks completed late");
    else atts.push("On-time rate of only "+pp(ot.otPct)+" ("+ot.otTotal+" tasks) — timeliness needs focus");
  }
  if(csm.overdueCount>50) atts.push("Critical: "+csm.overdueCount+" overdue tasks — needs immediate action");
  else if(csm.overdueCount>0) atts.push(csm.overdueCount+" overdue cadence tasks to clear");
  if(csm.cadCount>0) {
    if(csm.cadPct>=0.9) wins.push("Cadence completion on track at "+pp(csm.cadPct));
    else atts.push("Cadence completion "+pp(csm.cadPct)+" is below the 90% target");
  }

  const statBox = (label, val, sub, col) => (
    <div style={{background:"#F4F6FB",borderRadius:8,padding:14}}>
      <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:4}}>{label}</div>
      <div style={{fontSize:20,fontWeight:500,color:col||"#29355D",lineHeight:1.2}}>{val}</div>
      {sub&&<div style={{fontSize:11,color:"#808080",marginTop:3}}>{sub}</div>}
    </div>
  );

  return (
    <div>
      <div style={{...S.card,marginBottom:16}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:4}}>
          <div style={{fontSize:20,fontWeight:500,color:"#29355D"}}>{csm.name}</div>
          <button onClick={onClear} style={{fontSize:11,color:"#FF5000",background:"none",border:"0.5px solid #FF5000",borderRadius:20,padding:"4px 12px",cursor:"pointer"}}>✕ Clear filter</button>
        </div>
        <div style={{fontSize:12,color:"#808080",marginBottom:16}}>
          {i.t||csm.team||""}
          {i.r?" · "+i.r:""}
          {coach?" · Coach: "+coach.n:""}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,minmax(0,1fr))",gap:12}}>
          {statBox("Revenue", csm.rev>0?fd(csm.rev):"--", csm.mrr>0?"MRR "+fd(csm.mrr):null, csm.rev>0?"#FF5000":null)}
          {statBox("Email open", csm.sent>0?pp(csm.openRate):"--", csm.sent>0?csm.sent+" sent · "+pp(csm.replyRate)+" reply":"No emails", csm.sent>0?pc(csm.openRate):null)}
          {statBox("Cadence", csm.cadCount>0?pp(csm.cadPct):csm.overdueCount>0?csm.overdueCount+" overdue":"No tasks", csm.dueCount>0?csm.dueCount+" due · "+csm.newToday+" new":"Nothing due yesterday", csm.cadCount>0?pc(csm.cadPct):csm.overdueCount>0?"#dc2626":null)}
          {statBox("On-time %", ot&&ot.otTotal>=1?pp(ot.otPct):"--", ot&&ot.otTotal>=1?ot.otOnTime+"/"+ot.otTotal+" on time":"No data", ot&&ot.otTotal>=1?pc(ot.otPct):null)}
          {statBox("Integrations", csm.accts.length||"--", csm.accts.length>0?fd(totalAcctRev)+" total":null, null)}
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
        <div style={{background:"rgba(22,163,74,.06)",border:"0.5px solid rgba(22,163,74,.2)",borderRadius:12,padding:16}}>
          <div style={{fontSize:12,fontWeight:500,color:"#166534",marginBottom:10}}>✓ What's working</div>
          {wins.length===0
            ? <div style={{fontSize:12,color:"#808080",fontStyle:"italic"}}>No clear wins yet this period</div>
            : wins.map((w,i)=><div key={i} style={{fontSize:12,padding:"5px 0",borderBottom:"0.5px solid rgba(0,0,0,.06)",display:"flex",gap:6}}><span style={{color:"#16a34a",flexShrink:0}}>✓</span>{w}</div>)}
        </div>
        <div style={{background:"rgba(220,38,38,.05)",border:"0.5px solid rgba(220,38,38,.15)",borderRadius:12,padding:16}}>
          <div style={{fontSize:12,fontWeight:500,color:"#991b1b",marginBottom:10}}>⚠ Coaching focus</div>
          {atts.length===0
            ? <div style={{fontSize:12,color:"#808080",fontStyle:"italic"}}>No urgent issues identified</div>
            : atts.map((a,i)=><div key={i} style={{fontSize:12,padding:"5px 0",borderBottom:"0.5px solid rgba(0,0,0,.06)",display:"flex",gap:6}}><span style={{color:"#dc2626",flexShrink:0}}>⚠</span>{a}</div>)}
        </div>
      </div>

      <div style={S.card}>
        <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>
          Account integrations — {csm.accts.length} submissions · {fd(totalAcctRev)} total
        </div>
        {csm.accts.length===0
          ? <div style={{color:"#808080"}}>No submissions this period</div>
          : <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,tableLayout:"fixed"}}>
              <colgroup><col style={{width:"38%"}}/><col style={{width:"10%"}}/><col style={{width:"11%"}}/><col style={{width:"11%"}}/><col style={{width:"30%"}}/></colgroup>
              <thead><tr>{["Business","Type","MRR","OTR","Integration"].map(h=><th key={h} style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,padding:"0 8px 8px 0",textAlign:h==="MRR"||h==="OTR"?"right":"left",borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>{h}</th>)}</tr></thead>
              <tbody>{csm.accts.map((a,i)=>{
                const cls=a.t==="Monthly Recurring Revenue"?"mrr":a.t==="One-Time Revenue"?"otr":"non";
                const sh=a.t==="Monthly Recurring Revenue"?"MRR":a.t==="One-Time Revenue"?"OTR":"Non-Rev";
                const bg=cls==="mrr"?"rgba(22,163,74,.1)":cls==="otr"?"rgba(83,120,252,.1)":"rgba(128,128,128,.1)";
                const fg=cls==="mrr"?"#166534":cls==="otr"?"#1e3a8a":"#808080";
                return <tr key={i}>
                  <td style={{padding:"7px 8px 7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.b}</td>
                  <td style={{padding:"7px 8px 7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)"}}><span style={{fontSize:10,fontWeight:500,padding:"2px 7px",borderRadius:20,background:bg,color:fg}}>{sh}</span></td>
                  <td style={{padding:"7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}>{a.m>0?fd(a.m):"--"}</td>
                  <td style={{padding:"7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}>{a.o>0?fd(a.o):"--"}</td>
                  <td style={{padding:"7px 8px 7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",color:"#808080",fontSize:11,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.n||"--"}</td>
                </tr>;
              })}</tbody>
            </table>}
      </div>
    </div>
  );
}

// ── COACHING TAB ───────────────────────────────────────────────────────────
function CoachingView({csms, coach, onSelectCSM, onSelectCoach, onClear}) {
  if (onSelectCSM._selected) {
    const c = csms.find(x=>x.name===onSelectCSM._selected)||csms[0];
    return c ? <CSMDetail csm={c} onClear={onClear}/> : null;
  }
  const coaches = coach ? COACHES.filter(c=>c.e===coach) : COACHES;
  const cols = coaches.length===1?1:coaches.length===2?2:3;
  const overdue = csms.filter(c=>c.overdueCount>0).sort((a,b)=>b.overdueCount-a.overdueCount).slice(0,6);
  const attn = csms.filter(c=>c.cadCount>0&&c.cadPct<0.9).sort((a,b)=>a.cadPct-b.cadPct);
  const wins = csms.filter(c=>c.cadCount>0&&c.cadPct>=0.9).sort((a,b)=>b.cadPct-a.cadPct);
  return (
    <div>
      {overdue.length>0&&<div style={{background:"rgba(220,38,38,.04)",border:"0.5px solid rgba(220,38,38,.18)",borderRadius:12,padding:16,marginBottom:24}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
          <span style={{fontSize:16,fontWeight:500,color:"#dc2626"}}>⚠ Overdue cadence tasks</span>
          <span style={{fontSize:10,fontWeight:500,padding:"2px 8px",borderRadius:20,background:"rgba(220,38,38,.1)",color:"#991b1b"}}>{overdue.length} CSMs</span>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>
          {overdue.map(c=>{
            const i=lk(c.name)||{};
            return <div key={c.name} style={{...S.card,display:"flex",alignItems:"center",gap:10,cursor:"pointer",padding:12}} onClick={()=>onSelectCSM(c.name)}>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:12,fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.name}</div>
                <div style={{fontSize:11,color:"#808080",marginTop:2}}>{st(i.t||"")}{c.newToday>0?" · "+c.newToday+" new":""}</div>
              </div>
              <div style={{textAlign:"center",flexShrink:0}}>
                <div style={{fontSize:20,fontWeight:500,color:"#dc2626"}}>{c.overdueCount}</div>
                <div style={{fontSize:10,color:"#dc2626"}}>overdue</div>
              </div>
            </div>;
          })}
        </div>
      </div>}

      <div style={{display:"grid",gridTemplateColumns:`repeat(${cols},1fr)`,gap:16,marginBottom:24}}>
        {coaches.map(c=><CoachCard key={c.e} coach={c} csms={csms} onSelectCSM={name=>onSelectCSM(name)} onSelectCoach={onSelectCoach}/>)}
      </div>

      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12}}>
        <span style={{fontSize:16,fontWeight:500,color:"#29355D"}}>🚨 Needs attention</span>
        <span style={{fontSize:10,fontWeight:500,padding:"2px 8px",borderRadius:20,background:"rgba(220,38,38,.1)",color:"#991b1b"}}>{attn.length} CSMs</span>
      </div>
      <div style={{...S.card,marginBottom:24}}>
        {attn.length===0
          ? <div style={{textAlign:"center",padding:20,color:"#16a34a",fontWeight:500}}>🎉 Everyone at 90%+ today!</div>
          : <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead><tr>{["CSM","Team","Cadence","On-time","Overdue","Gap"].map((h,j)=><th key={h} style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,padding:"0 8px 8px 0",textAlign:j>=2?"right":"left",borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>{h}</th>)}</tr></thead>
              <tbody>{attn.map(c=>{
                const i=lk(c.name)||{};
                return <tr key={c.name} style={{cursor:"pointer"}} onClick={()=>onSelectCSM(c.name)}>
                  <td style={{padding:"7px 8px 7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)"}}>
                    <div style={{fontWeight:500,color:"#29355D"}}>{c.name}</div>
                    <div style={{fontSize:11,color:"#808080",fontStyle:"italic",marginTop:2}}>{c.cadPct===0?"No completions":c.cadPct<0.5?"Significantly behind":"Below average"}</div>
                  </td>
                  <td style={{padding:"7px 8px 7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",color:"#808080"}}>{st(i.t||"")}</td>
                  <td style={{padding:"7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}><span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:"rgba(220,38,38,.12)",color:"#991b1b"}}>{Math.round(c.cadPct*100)}%</span></td>
                  <td style={{padding:"7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",fontWeight:500,color:c.otTotal>=3?pc(c.otPct):"#888"}}>{c.otTotal>=3?pp(c.otPct):"--"}</td>
                  <td style={{padding:"7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}>{c.overdueCount>0?<span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:"rgba(220,38,38,.1)",color:"#991b1b"}}>{c.overdueCount}</span>:"--"}</td>
                  <td style={{padding:"7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",fontWeight:500,color:"#dc2626"}}>{Math.round((0.9-c.cadPct)*100)}%</td>
                </tr>;
              })}</tbody>
            </table>}
      </div>

      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12}}>
        <span style={{fontSize:16,fontWeight:500,color:"#29355D"}}>🏅 Winning today</span>
        <span style={{fontSize:10,fontWeight:500,padding:"2px 8px",borderRadius:20,background:"rgba(22,163,74,.1)",color:"#166534"}}>{wins.length} CSMs</span>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>
        {wins.map(c=>{
          const i=lk(c.name)||{};
          return <div key={c.name} style={{...S.card,display:"flex",alignItems:"center",gap:10,borderLeft:"3px solid #16a34a",cursor:"pointer",padding:12}} onClick={()=>onSelectCSM(c.name)}>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontSize:12,fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.name}</div>
              <div style={{fontSize:11,color:"#808080",marginTop:2}}>{st(i.t||"")} · {Math.round(c.cadPct*100)}%</div>
            </div>
            <span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:"rgba(22,163,74,.12)",color:"#166534",flexShrink:0}}>Win</span>
          </div>;
        })}
      </div>
    </div>
  );
}

// ── OVERVIEW TAB ───────────────────────────────────────────────────────────
function OverviewView({csms, allCSMs}) {
  const totalRev = csms.reduce((s,c)=>s+c.rev,0);
  const totalMRR = csms.reduce((s,c)=>s+c.mrr,0);
  const totalSent = csms.reduce((s,c)=>s+c.sent,0);
  const emC = csms.filter(c=>c.sent>0);
  const avgOpen = emC.length ? emC.reduce((s,c)=>s+c.openRate,0)/emC.length : 0;
  const totalOD = csms.reduce((s,c)=>s+c.overdueCount,0);
  const otC = csms.filter(c=>c.otTotal>=3);
  const avgOT = otC.length ? otC.reduce((s,c)=>s+c.otPct,0)/otC.length : 0;
  const teamRev={};
  csms.forEach(c=>{const t=c.team||(lk(c.name)&&lk(c.name).t)||"";if(t)(teamRev[t]=(teamRev[t]||0)+c.rev);});
  const trs=Object.entries(teamRev).sort((a,b)=>b[1]-a[1]);
  const maxR=trs[0]&&trs[0][1]||1;
  const topRev=[...csms].filter(c=>c.rev>0).sort((a,b)=>b.rev-a.rev).slice(0,8);
  const maxRI=topRev[0]&&topRev[0].rev||1;
  const topOT=[...csms].filter(c=>c.otTotal>=3).sort((a,b)=>b.otPct-a.otPct).slice(0,8);
  const metrics=[
    {l:"Total revenue",v:fd(totalRev),s:"MRR "+fk(totalMRR),col:"#FF5000"},
    {l:"Emails sent",v:totalSent,s:emC.length+" senders",col:"#5378FC"},
    {l:"Avg open rate",v:pp(avgOpen),s:"Target 70%+",col:avgOpen>=0.7?"#16a34a":"#d97706"},
    {l:"Avg on-time %",v:otC.length?pp(avgOT):"--",s:otC.length+" CSMs tracked",col:avgOT>=0.8?"#16a34a":avgOT>=0.6?"#d97706":"#dc2626"},
    {l:"Overdue tasks",v:totalOD,s:csms.filter(c=>c.overdueCount>0).length+" CSMs",col:"#dc2626"},
  ];
  const hbar=(name,pct,val,max,col,onClick)=>(
    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,fontSize:12,cursor:"pointer"}} onClick={onClick}>
      <span style={{width:140,flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{name}</span>
      <div style={{flex:1,height:4,background:"#ECEEF1",borderRadius:2,overflow:"hidden"}}>
        <div style={{width:((pct||0)*100).toFixed(1)+"%",height:"100%",background:col,opacity:.85,borderRadius:2}}/>
      </div>
      <span style={{width:50,textAlign:"right",fontSize:11,color:"#808080",flexShrink:0}}>{val}</span>
    </div>
  );
  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(5,minmax(0,1fr))",gap:12,marginBottom:20}}>
        {metrics.map(m=>(
          <div key={m.l} style={{background:"#ECEEF1",borderRadius:8,padding:14,position:"relative",overflow:"hidden"}}>
            <div style={{position:"absolute",top:0,left:0,right:0,height:2,background:m.col,borderRadius:"8px 8px 0 0"}}/>
            <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:6}}>{m.l}</div>
            <div style={{fontSize:22,fontWeight:500,color:"#29355D",lineHeight:1,marginBottom:3}}>{m.v}</div>
            <div style={{fontSize:11,color:"#808080"}}>{m.s}</div>
          </div>
        ))}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
        <div style={S.card}>
          <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>Revenue by team</div>
          {trs.map(([t,v])=>{
            const coach=COACHES.find(c=>c.t===t);
            return hbar(st(t),v/maxR,fk(v),maxR,TEAM_COLS[t]||"#888",()=>{});
          })}
        </div>
        <div style={S.card}>
          <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>Top CSM revenue</div>
          {topRev.map(c=>hbar(c.name.split(" ").slice(0,2).join(" "),c.rev/maxRI,fk(c.rev),maxRI,"#FF5000",()=>{}))}
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        <div style={S.card}>
          <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>On-time cadence % — top 8</div>
          {topOT.length===0
            ? <div style={{color:"#808080"}}>No on-time data available</div>
            : topOT.map(c=>hbar(c.name.split(" ").slice(0,2).join(" "),c.otPct,pp(c.otPct),1,bc(c.otPct,0.8,0.6),()=>{}))}
        </div>
        <div style={S.card}>
          <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>Overdue cadence</div>
          {csms.filter(c=>c.overdueCount>0).sort((a,b)=>b.overdueCount-a.overdueCount).map(c=>(
            <div key={c.name} style={{display:"flex",alignItems:"center",gap:8,padding:"4px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)"}}>
              <span style={{flex:1,fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.name}</span>
              <span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:"rgba(220,38,38,.1)",color:"#991b1b"}}>{c.overdueCount}</span>
              {c.newToday>0&&<span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:"rgba(83,120,252,.1)",color:"#1e3a8a"}}>{c.newToday} new</span>}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ── LEADERBOARD TAB ────────────────────────────────────────────────────────
function LeaderboardView({csms}) {
  const [sort,setSort]=useState({col:"rev",dir:"desc"});
  const sorted=[...csms].sort((a,b)=>{
    const av=a[sort.col]||0,bv=b[sort.col]||0;
    return sort.dir==="desc"?bv-av:av-bv;
  });
  const medals=["🥇","🥈","🥉"];
  const th=(col,lbl)=>(
    <th onClick={()=>setSort(s=>({col,dir:s.col===col&&s.dir==="desc"?"asc":"desc"}))}
      style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,padding:"0 0 8px",textAlign:"right",cursor:"pointer",borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>
      {lbl}{sort.col===col?(sort.dir==="desc"?" ▼":" ▲"):""}
    </th>
  );
  return (
    <div style={S.card}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
        <thead><tr>
          <th style={{width:28,fontSize:10,color:"#808080",fontWeight:500,padding:"0 0 8px",textAlign:"left",borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>#</th>
          <th style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,padding:"0 0 8px",textAlign:"left",borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>CSM</th>
          <th style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,padding:"0 0 8px",textAlign:"left",borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>Team</th>
          {th("rev","Revenue")}{th("sent","Emails")}{th("openRate","Open %")}{th("cadPct","Cadence")}{th("otPct","On-time %")}{th("overdueCount","Overdue")}
        </tr></thead>
        <tbody>{sorted.map((c,i)=>{
          const info=lk(c.name)||{};
          const col=TEAM_COLS[info.t||c.team]||"#888";
          return <tr key={c.name}>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)"}}>{i<3?medals[i]:(i+1)+"."}</td>
            <td style={{padding:"9px 8px 9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",fontWeight:500}}>{c.name}</td>
            <td style={{padding:"9px 8px 9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)"}}><span style={{display:"inline-block",width:6,height:6,borderRadius:"50%",background:col,marginRight:5,verticalAlign:"middle"}}/><span style={{color:"#808080",fontSize:11}}>{st(info.t||c.team)}</span></td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:"#FF5000",fontWeight:500}}>{c.rev>0?fd(c.rev):"--"}</td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}>{c.sent>0?c.sent:"--"}</td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",fontWeight:500,color:c.sent>0?pc(c.openRate):"#888"}}>{c.sent>0?pp(c.openRate):"--"}</td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",fontWeight:500,color:c.cadCount>0?pc(c.cadPct):"#888"}}>{c.cadCount>0?pp(c.cadPct):"--"}</td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",fontWeight:500,color:c.otTotal>=3?pc(c.otPct):"#888"}}>{c.otTotal>=3?pp(c.otPct):"--"}</td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}>{c.overdueCount>0?<span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:"rgba(220,38,38,.1)",color:"#991b1b"}}>{c.overdueCount}</span>:"--"}</td>
          </tr>;
        })}</tbody>
      </table>
    </div>
  );
}

// ── ACTIVITY TAB ───────────────────────────────────────────────────────────
function SortableTable({title, cols, rows, defaultCol, defaultDir="desc"}) {
  const [sort, setSort] = useState({col: defaultCol, dir: defaultDir});
  const sorted = [...rows].sort((a, b) => {
    const av = a[sort.col] ?? (typeof a[sort.col]==="string" ? "" : -Infinity);
    const bv = b[sort.col] ?? (typeof b[sort.col]==="string" ? "" : -Infinity);
    if (typeof av === "string") return sort.dir==="asc" ? av.localeCompare(bv) : bv.localeCompare(av);
    return sort.dir === "asc" ? av - bv : bv - av;
  });
  const toggle = col => setSort(s => ({col, dir: s.col===col && s.dir==="desc" ? "asc" : "desc"}));
  const thStyle = (right) => ({fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,padding:"0 0 8px",textAlign:right?"right":"left",borderBottom:"0.5px solid rgba(41,53,93,.08)",cursor:"pointer",userSelect:"none",whiteSpace:"nowrap"});
  const arrow = col => sort.col===col ? (sort.dir==="desc"?" ▼":" ▲") : " ↕";
  return (
    <div style={S.card}>
      <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>{title}</div>
      {rows.length===0
        ? <div style={{color:"#808080",fontSize:12}}>No data available</div>
        : <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr>
              {cols.map(({key,label,right})=>(
                <th key={key} style={thStyle(right)} onClick={()=>toggle(key)}>
                  {label}<span style={{color: sort.col===key ? "#FF5000" : "#ccc", fontSize:9}}>{arrow(key)}</span>
                </th>
              ))}
            </tr></thead>
            <tbody>{sorted.map((row,i)=>row._render(i))}</tbody>
          </table>}
    </div>
  );
}

function ActivityView({csms}) {
  const tdBase = {padding:"6px 0", borderBottom:"0.5px solid rgba(41,53,93,.05)"};

  // Email table rows
  const emRows = csms.filter(c=>c.sent>0).map(c=>({
    name: c.name, sent: c.sent, openRate: c.openRate, replyRate: c.replyRate,
    _render: (i) => <tr key={c.name}>
      <td style={tdBase}>{c.name}</td>
      <td style={{...tdBase,textAlign:"right"}}>{c.sent}</td>
      <td style={{...tdBase,textAlign:"right",fontWeight:500,color:pc(c.openRate)}}>{pp(c.openRate)}</td>
      <td style={{...tdBase,textAlign:"right",fontWeight:500,color:pc(c.replyRate)}}>{pp(c.replyRate)}</td>
    </tr>
  }));

  // On-time table rows
  const otRows = csms.filter(c=>c.otTotal>=3).map(c=>({
    name: c.name, otTotal: c.otTotal, otOnTime: c.otOnTime, otPct: c.otPct,
    _render: (i) => <tr key={c.name}>
      <td style={tdBase}>{c.name}</td>
      <td style={{...tdBase,textAlign:"right",color:"#808080"}}>{c.otTotal}</td>
      <td style={{...tdBase,textAlign:"right"}}>{c.otOnTime}</td>
      <td style={{...tdBase,textAlign:"right",fontWeight:500,color:bc(c.otPct,0.8,0.6)}}>{pp(c.otPct)}</td>
    </tr>
  }));

  // Due/overdue table rows
  const dueRows = csms.filter(c=>c.dueCount>0).map(c=>({
    name: c.name, dueCount: c.dueCount, overdueCount: c.overdueCount, newToday: c.newToday,
    _render: (i) => <tr key={c.name}>
      <td style={tdBase}>{c.name}</td>
      <td style={{...tdBase,textAlign:"right"}}>{c.dueCount}</td>
      <td style={{...tdBase,textAlign:"right"}}>
        <span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:"rgba(220,38,38,.1)",color:"#991b1b"}}>{c.overdueCount}</span>
      </td>
      <td style={{...tdBase,textAlign:"right"}}>
        {c.newToday>0
          ? <span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:"rgba(83,120,252,.1)",color:"#1e3a8a"}}>{c.newToday}</span>
          : "--"}
      </td>
    </tr>
  }));

  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
        <SortableTable
          title="Email performance"
          defaultCol="sent"
          cols={[
            {key:"name",      label:"CSM",     right:false},
            {key:"sent",      label:"Sent",    right:true},
            {key:"openRate",  label:"Open %",  right:true},
            {key:"replyRate", label:"Reply %", right:true},
          ]}
          rows={emRows}
        />
        <SortableTable
          title="On-time cadence %"
          defaultCol="otPct"
          cols={[
            {key:"name",     label:"CSM",     right:false},
            {key:"otTotal",  label:"Tasks",   right:true},
            {key:"otOnTime", label:"On-time", right:true},
            {key:"otPct",    label:"Rate",    right:true},
          ]}
          rows={otRows}
        />
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        <SortableTable
          title="Due / past due cadence"
          defaultCol="overdueCount"
          cols={[
            {key:"name",         label:"CSM",     right:false},
            {key:"dueCount",     label:"Due",     right:true},
            {key:"overdueCount", label:"Overdue", right:true},
            {key:"newToday",     label:"New",     right:true},
          ]}
          rows={dueRows}
        />
        <div style={S.card}>
          <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>Cadence completions</div>
          <div style={{color:"#808080",fontSize:12}}>Data from cadence tab — upload to sheet weekly to see here.</div>
        </div>
      </div>
    </div>
  );
}

// ── PIN LOCK ───────────────────────────────────────────────────────────────
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
        <div style={{width:56,height:56,borderRadius:"50%",background:"#29355D",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 20px",fontSize:24}}>🔒</div>
        <div style={{fontFamily:"Nunito,sans-serif",fontSize:20,fontWeight:800,color:"#29355D",marginBottom:6}}>CSM Coaching Dashboard</div>
        <div style={{fontSize:13,color:"#808080",marginBottom:24}}>Enter your team PIN to continue</div>
        <input type="password" value={val} onChange={e=>setVal(e.target.value)} onKeyDown={e=>e.key==="Enter"&&check()}
          placeholder="Enter PIN" autoFocus
          style={{width:"100%",padding:"10px 14px",fontSize:15,borderRadius:10,border:"1px solid "+(err?"#dc2626":"rgba(41,53,93,.2)"),outline:"none",textAlign:"center",letterSpacing:4,marginBottom:12,fontFamily:"inherit"}}/>
        {err&&<div style={{color:"#dc2626",fontSize:12,marginBottom:10}}>Incorrect PIN</div>}
        <button onClick={check} style={{width:"100%",padding:11,background:"#FF5000",border:"none",borderRadius:10,color:"#fff",fontSize:14,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>Unlock Dashboard</button>
        <div style={{fontSize:11,color:"#808080",marginTop:16}}>Contact your coach for the PIN</div>
      </div>
    </div>
  );
}

// ── MAIN APP ───────────────────────────────────────────────────────────────
export default function App() {
  const [unlocked, setUnlocked] = useState(()=>{try{return sessionStorage.getItem(PIN_KEY)==="1";}catch(e){return false;}});
  const [csms, setCSMs] = useState([]);
  const [tab, setTab] = useState("coaching");
  const [filterCoach, setFilterCoach] = useState("");
  const [filterCSM, setFilterCSM] = useState("");
  const [status, setStatus] = useState("loading");
  const [updatedAt, setUpdatedAt] = useState(null);

  const allCSMNames = [...new Set(csms.map(c=>c.name))].sort();

  const filteredCSMs = csms.filter(c => {
    const i = lk(c.name);
    if (filterCoach && (i&&i.c||c.coach) !== filterCoach) return false;
    if (filterCSM && c.name !== filterCSM) return false;
    return true;
  });

  const activeCoaches = filterCoach ? COACHES.filter(c=>c.e===filterCoach)
    : filterCSM ? COACHES.filter(c=>{const i=lk(filterCSM);return i&&i.c===c.e;})
    : COACHES;

  useEffect(()=>{
    if (!unlocked) return;
    setStatus("loading");
    Promise.all([
      fetchCSV(CSV_REV),
      fetchCSV(CSV_EMAIL),
      fetchCSV(CSV_CAD),
      fetchCSV(CSV_DUE),
      fetchCSV(CSV_ONTIME),
    ]).then(([revRows, emailRows, cadRows, dueRows, ontimeRows]) => {
      const rev    = mapRev(revRows);
      const email  = mapEmail(emailRows);
      const cad    = mapCadence(cadRows);
      const due    = mapDue(dueRows);
      const ontime = mapOnTime(ontimeRows);
      const built  = buildCSMs(rev, email, cad, due, ontime);
      console.log("Built CSMs:", built.length);
      setCSMs(built);
      setUpdatedAt(new Date().toLocaleTimeString());
      setStatus("ok");
    }).catch(err => {
      console.error("Sheet load error:", err);
      setStatus("error");
    });
  }, [unlocked]);

  function buildAIPrompt() {
    const lines = [];
    if (filterCSM) {
      const c = csms.find(x=>x.name===filterCSM);
      if (!c) return "No data for "+filterCSM;
      const i = lk(c.name)||{};
      const coach = COACHES.find(x=>x.e===(i.c||c.coach));
      lines.push("I need coaching tips for a specific CSM at Thryv.");
      lines.push("CSM: "+c.name+" | Team: "+(i.t||"")+" | Tier: "+(i.r||"")+(coach?" | Coach: "+coach.n:""));
      lines.push("");
      lines.push("PERFORMANCE:");
      lines.push("  Revenue: "+(c.rev>0?fd(c.rev)+" (MRR "+fd(c.mrr)+")":"None this period"));
      lines.push("  Email: "+(c.sent>0?c.sent+" sent, "+pp(c.openRate)+" open rate, "+pp(c.replyRate)+" reply rate":"No email activity"));
      lines.push("  On-time cadence: "+(c.otTotal>=3?pp(c.otPct)+" ("+c.otOnTime+"/"+c.otTotal+" tasks)":"Not enough data"));
      lines.push("  Overdue tasks: "+(c.overdueCount>0?c.overdueCount:"None"));
      if (c.accts.length>0) lines.push("  Accounts: "+c.accts.slice(0,5).map(a=>a.b+(a.m>0?" MRR "+fd(a.m):a.o>0?" OTR "+fd(a.o):"")).join(", "));
      lines.push("");
      lines.push("Give me: (1) What is this CSM doing well? (2) Top 2-3 coaching priorities for our next 1:1? (3) Specific action items for this week. (4) Any red flags?");
    } else if (filterCoach) {
      const coach = COACHES.find(c=>c.e===filterCoach);
      lines.push("Coaching insights for "+coach.n+"'s team ("+coach.t+") at Thryv.");
      lines.push("");
      filteredCSMs.forEach(c=>{
        lines.push(c.name+": rev="+(c.rev>0?fd(c.rev):"none")+", email open="+(c.sent>0?pp(c.openRate):"n/a")+", on-time="+(c.otTotal>=3?pp(c.otPct):"n/a")+", overdue="+(c.overdueCount>0?c.overdueCount:"none"));
      });
      lines.push("");
      lines.push("(1) Who needs most urgent attention? (2) Team patterns? (3) What should "+coach.n+" prioritize this week? (4) Who is excelling?");
    } else {
      lines.push("Full CSM team coaching overview — Thryv.");
      lines.push("");
      COACHES.forEach(coach=>{
        const team=csms.filter(c=>(lk(c.name)&&lk(c.name).c===coach.e)||c.coach===coach.e);
        const cadC=team.filter(c=>c.cadCount>0);
        const avgCad=cadC.length?cadC.reduce((s,c)=>s+c.cadPct,0)/cadC.length:null;
        const otC=team.filter(c=>c.otTotal>=3);
        const avgOT=otC.length?otC.reduce((s,c)=>s+c.otPct,0)/otC.length:null;
        const totR=team.reduce((s,c)=>s+c.rev,0);
        lines.push("COACH: "+coach.n+" | "+coach.t);
        lines.push("  Revenue: $"+Math.round(totR).toLocaleString()+" | Cadence: "+(avgCad!=null?Math.round(avgCad*100)+"%":"n/a")+" | On-time: "+(avgOT!=null?Math.round(avgOT*100)+"%":"n/a"));
        lines.push("  Needs coaching: "+cadC.filter(c=>c.cadPct<0.9).map(c=>c.name).join(", ")||"none");
        lines.push("");
      });
      lines.push("(1) Most urgent attention? (2) Biggest risks this week? (3) Top 3 action items for the leader?");
    }
    return lines.join("\n");
  }

  function openAI() {
    window.open("https://claude.ai/new?q="+encodeURIComponent(buildAIPrompt()), "_blank");
  }

  if (!unlocked) return <PinLock onUnlock={()=>setUnlocked(true)}/>;

  const aiLabel = filterCSM ? "Ask AI: "+filterCSM.split(" ")[0]
    : filterCoach ? "Ask AI: "+COACHES.find(c=>c.e===filterCoach).n.split(" ")[0]+"'s team"
    : "Ask AI Coach";

  const hasData = csms.length > 0;

  // Pass selected CSM through a trick so CoachingView can access it
  const selectCSMFn = name => { setFilterCSM(name); setFilterCoach(""); };
  selectCSMFn._selected = filterCSM;

  return (
    <div style={{fontFamily:"Nunito Sans,sans-serif",background:"#F4F6FB",minHeight:"100vh",fontSize:13}}>
      {/* NAV */}
      <div style={S.nav}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"12px 24px",borderBottom:"1px solid rgba(255,255,255,.1)"}}>
          <div style={{color:"#fff",fontSize:14,fontWeight:500,display:"flex",alignItems:"center",gap:10}}>
            <img src="https://assets.thryv.com/prod/media/thryv-main-logo-white.png" alt="Thryv" style={{height:24}} onError={e=>{e.target.style.display="none";}}/>
            <div style={{width:1,height:18,background:"rgba(255,255,255,.25)"}}/>
            <span style={{color:"rgba(255,255,255,.6)",fontSize:13}}>CSM Coaching Dashboard</span>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            {status==="loading"&&<span style={{fontSize:11,padding:"4px 10px",borderRadius:20,background:"rgba(255,255,255,.1)",color:"rgba(255,255,255,.6)"}}>⟳ Loading...</span>}
            {status==="ok"&&<span style={{fontSize:11,padding:"4px 10px",borderRadius:20,background:"rgba(22,163,74,.25)",color:"#86efac"}}>✓ Live{updatedAt?" · "+updatedAt:""}</span>}
            {status==="error"&&<span style={{fontSize:11,padding:"4px 10px",borderRadius:20,background:"rgba(220,38,38,.25)",color:"#fca5a5"}}>✗ Sync error</span>}
            {hasData&&<button onClick={openAI} style={{background:"#FF5000",border:"none",color:"#fff",fontSize:12,fontWeight:500,padding:"7px 14px",borderRadius:6,cursor:"pointer"}}>{aiLabel} ↗</button>}
          </div>
        </div>
        <div style={{display:"flex",alignItems:"stretch",padding:"0 24px"}}>
          {["coaching","overview","leaderboard","activity"].map(t=>(
            <button key={t} onClick={()=>setTab(t)}
              style={{padding:"10px 18px",fontSize:13,fontWeight:500,color:tab===t?"#fff":"rgba(255,255,255,.55)",background:"transparent",border:"none",cursor:"pointer",borderBottom:tab===t?"3px solid #FF5000":"3px solid transparent",whiteSpace:"nowrap"}}>
              {t==="coaching"?"Coaching":t.charAt(0).toUpperCase()+t.slice(1)}
            </button>
          ))}
        </div>
      </div>

      {/* FILTERS */}
      {hasData&&(
        <div style={{background:"#fff",borderBottom:"0.5px solid rgba(41,53,93,.08)",padding:"8px 24px",display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          <select value={filterCoach} onChange={e=>{setFilterCoach(e.target.value);setFilterCSM("");}}
            style={{fontSize:12,fontWeight:500,padding:"5px 10px",borderRadius:8,border:"0.5px solid "+(filterCoach?"#FF5000":"rgba(41,53,93,.15)"),background:"#F4F6FB",color:filterCoach?"#FF5000":"#29355D",cursor:"pointer"}}>
            <option value="">All coaches</option>
            {COACHES.map(c=><option key={c.e} value={c.e}>{c.n}</option>)}
          </select>
          <select value={filterCSM} onChange={e=>{setFilterCSM(e.target.value);setFilterCoach("");}}
            style={{fontSize:12,fontWeight:500,padding:"5px 10px",borderRadius:8,border:"0.5px solid "+(filterCSM?"#FF5000":"rgba(41,53,93,.15)"),background:"#F4F6FB",color:filterCSM?"#FF5000":"#29355D",cursor:"pointer"}}>
            <option value="">All CSMs</option>
            {allCSMNames.map(n=><option key={n} value={n}>{n}</option>)}
          </select>
          {(filterCoach||filterCSM)&&(
            <span style={{background:"#FF5000",color:"#fff",fontSize:11,fontWeight:500,padding:"4px 10px",borderRadius:20,display:"inline-flex",alignItems:"center",gap:6}}>
              {filterCSM||COACHES.find(c=>c.e===filterCoach).n}
              <button onClick={()=>{setFilterCoach("");setFilterCSM("");}} style={{background:"none",border:"none",color:"rgba(255,255,255,.8)",cursor:"pointer",fontSize:14,lineHeight:1,padding:0}}>✕</button>
            </span>
          )}
        </div>
      )}

      {/* CONTENT */}
      {status==="loading"&&!hasData&&(
        <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"60vh",flexDirection:"column",gap:16}}>
          <div style={{fontSize:32}}>⟳</div>
          <div style={{fontSize:15,color:"#808080"}}>Loading data from Google Sheet...</div>
        </div>
      )}
      {status==="error"&&!hasData&&(
        <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"60vh",flexDirection:"column",gap:16}}>
          <div style={{fontSize:32}}>⚠️</div>
          <div style={{fontSize:15,color:"#dc2626",fontWeight:700}}>Could not load data from Google Sheet</div>
          <div style={{fontSize:13,color:"#808080"}}>Check the browser console for details</div>
          <button onClick={()=>window.location.reload()} style={{background:"#FF5000",color:"#fff",border:"none",borderRadius:8,padding:"10px 20px",cursor:"pointer",fontSize:13,fontWeight:700}}>Try Again</button>
        </div>
      )}
      {hasData&&(
        <div style={{padding:"20px 24px"}}>
          {tab==="coaching"&&<CoachingView csms={filteredCSMs} coach={filterCoach} onSelectCSM={selectCSMFn} onSelectCoach={e=>{setFilterCoach(e);setFilterCSM("");}} onClear={()=>{setFilterCoach("");setFilterCSM("");}}/>}
          {tab==="overview"&&<OverviewView csms={filteredCSMs} allCSMs={csms}/>}
          {tab==="leaderboard"&&<LeaderboardView csms={filteredCSMs}/>}
          {tab==="activity"&&<ActivityView csms={filteredCSMs}/>}
        </div>
      )}
    </div>
  );
}
