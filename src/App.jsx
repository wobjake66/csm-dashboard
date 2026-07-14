import React, { useState, useEffect } from "react";
const imgCrushingIt  = "https://raw.githubusercontent.com/wobjake66/csm-dashboard/main/crushing_it.jpg";
const imgAlmostThere = "https://raw.githubusercontent.com/wobjake66/csm-dashboard/main/almost_there.jpg";
const imgNeedsLove   = "https://raw.githubusercontent.com/wobjake66/csm-dashboard/main/needs_love.jpg";
const imgLegend      = "https://raw.githubusercontent.com/wobjake66/csm-dashboard/main/legend_status.jpg";
import * as XLSX from "xlsx";

const PIN = "thryv2025";
const PIN_KEY  = "csm_pin_v1";
const FONT_KEY = "csm_font_v1";

const CSV_REV     = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=1721544342&single=true&output=csv"; // live JotForm sync
const CSV_EMAIL   = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=0&single=true&output=csv";
const CSV_CAD     = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=1973544046&single=true&output=csv";
const CSV_DUE     = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=341836664&single=true&output=csv";
const CSV_ONTIME  = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=459845057&single=true&output=csv";
const CSV_SKIPPED = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=1238903633&single=true&output=csv"; // prior-day skipped cadences
const CSV_HISTORY = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=162960918&single=true&output=csv";
const CSV_BOB     = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=729347262&single=true&output=csv"; // book of business billing
const CSV_BOB_DET = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=873304103&single=true&output=csv"; // book of business detail
const CSV_BOB_ADJ = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=806795688&single=true&output=csv"; // bob_adjustments (form submission log)
// Q3 BOB tracking tabs
// bob_q3_current = where you paste the fresh BOB report (input, not read by dashboard)
// bob_q3_results = CSM-level rollup written by Apps Script (read by dashboard)
// bob_q3_log     = append-only change log (read by dashboard)
const CSV_Q3_RESULTS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=766144759&single=true&output=csv"; // bob_q3_results
const CSV_Q3_LOG     = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=1341900683&single=true&output=csv"; // bob_q3_log
const CSV_Q3_BOQ     = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=1420711886&single=true&output=csv"; // bob_q3_boq account level
const CSV_Q3_CUR     = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=549601512&single=true&output=csv";  // bob_q3_current account level
const CSV_DOMO_BOQ   = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=40086456&single=true&output=csv";   // Domo_bob_q3_boq (fresh Domo export, multi-L2-row per EID)
const CSV_Q3_SUPP    = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=653676072&single=true&output=csv"; // bob_q3_supplemental (secondary SF revenue report, no CSM column)
const CSV_CALLS   = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=466716688&single=true&output=csv"; // calls history (raw bookings export)
const CSV_QA_MC   = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=1435312575&single=true&output=csv"; // MC Activation QA
const CSV_QA_SS   = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=2091285368&single=true&output=csv"; // Setup & Strategy QA
const CSV_BC_CHURN= "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=295771282&single=true&output=csv"; // BC churn by coach/rep
const CSV_MC_CHURN= "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=1002996767&single=true&output=csv"; // MC churn by coach/rep
const CSV_CHURN_ALERTS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=724984916&single=true&output=csv"; // daily new churn alerts

const CAD_ACCTS={"Taylor Kidd":[{"n":"F Lees Tax & Accounting Services","ott":3,"oto":2},{"n":"Liz Bienstock, Realtor","ott":1,"oto":0},{"n":"Nixdorf Tree Service","ott":1,"oto":0},{"n":"Torrey Tucker Electric LLC","ott":1,"oto":1},{"n":"Uniquely Yours Screen Printing","ott":3,"oto":3}],"April Hall":[{"n":"JaysonRachel Carter","ott":1,"oto":0},{"n":"M & C Homecare","d":[{"t":"Highlight Video","due":"5/29/2026","ov":true,"nw":false}]},{"n":"TAG Construction","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]}],"Zoltan Rudolf":[{"n":"Budget Screens & Awnings","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Classic Steamboat Cruises","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Clean Planet West Auckand","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"North East Survey Design","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Perfect Fit Interior","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"VISION INSTALLATIONS LIMITED","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]}],"Libby Booher":[{"n":"Advanced Acupuncture And Chinese Herbal Clinic","ott":1,"oto":0},{"n":"Battle Ground Personal Training","ott":3,"oto":1},{"n":"CAMPBELL REMODELING AND MAINTENANCE LLC","ott":3,"oto":1},{"n":"Discount Auto Glass","ott":3,"oto":1},{"n":"Hawaii Food Products Inc","ott":1,"oto":0},{"n":"Motorcycle Rider Training","ott":3,"oto":1},{"n":"Zebra Stripes Child Care & Preschool","ott":1,"oto":1}],"Indu Vijay":[{"n":"Blinds Hub","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Bold Insurance","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Byte Size Labs","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Oz Seals","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Tranquility Pools","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Luis Aguasvivas Peralta":[{"n":"Mission Success Solutions LLC","ott":1,"oto":1}],"Mark Velazquez":[{"n":"Antesberger Plumbing","ott":1,"oto":1},{"n":"CJ Construction","ott":1,"oto":1},{"n":"CJS Heating and Air","ott":1,"oto":0},{"n":"FYZICAL Therapy & Balance Center of Little Silver","ott":1,"oto":0},{"n":"Home Pro Masters","ott":1,"oto":0},{"n":"Korte Does It All, Inc.","ott":1,"oto":0},{"n":"Legacy Heating and Air, Inc.","ott":1,"oto":0},{"n":"Newtown Heating & Air Conditioning Inc","ott":3,"oto":3},{"n":"Precision Today Plumbing Heating Cooling Electrical","ott":1,"oto":0},{"n":"Stitches and Screens","ott":2,"oto":1}],"Sylvia Appla":[{"n":"A F Isaac Suveying Pty Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"AESTHETIC DENTAL LIMITED","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"AK Painting","ott":1,"oto":0},{"n":"ASAP Security and Protection","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"ATOM TREE SPECIALISTS LIMITED","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Acland Street Physiotherapy","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"All Four x 4 Spares","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Allen's Interiors & Building","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"BYC Dental Pty Ltd ATF BYC Dental Trust","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Bayview Denture Care","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Cavana Services","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Country Tyres Yass","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Darwin Family Law","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Dons Mobile Marine","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"EK GLASS LIMITED","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Earth 2 Ocean Communications","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Exacte Advisors","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Exo Graphics","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Fencemen Fencing Contractors Ltd.","d":[{"t":"Email","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Flash Roofing Supplies","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Gold Coast Hypnotherapy","ott":1,"oto":0},{"n":"HERRON TODD WHITE","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Hyperspace Architecture","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Informed Decisions Consultancy Limited","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"J S De Rooy","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"KITCHEN & APPLIANCES SOUTHLAND LIMITED","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"MURRAY BROWN ELECTRICAL (1996)","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Nerang Stockfeed","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Northlane Welding Services","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Nudge Osteopathy","d":[{"t":"Email","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Oxen Lawyers","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"PNJ Hire","ott":1,"oto":0},{"n":"Pearla Plumbing & Electrical","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Penrith Engine Services","d":[{"t":"Email","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Perth Taxi Booking","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"PowerSmart Heat Pumps","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Quickturn Pty Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Red Roo Solutions","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"SHRI METALS MERCHANT PTY LTD","ott":1,"oto":0},{"n":"Scott Electrics","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Shubbs International Pty Ltd","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Smart Thread Solutions","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Straightline Guttering Pty Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Streamline Drains & Pipelines (NSW) PL","d":[{"t":"Email","due":"6/1/2026","ov":true,"nw":true}]},{"n":"The Baker's Den Bakery Cafe","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Village Carpet Care","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Voodoo Rod And Custom","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"WORKSTORE SA LIMITED","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Wagga Scrap Metals","d":[{"t":"Email","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Wide Bay Memorials","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Wollongong Auto Excellence","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]}],"Rafael Sencion Sencion":[{"n":"5 Star Fence","ott":1,"oto":1},{"n":"D & D Overhead Door LLC","ott":1,"oto":1},{"n":"Grow With Us Learning Centers","ott":1,"oto":1},{"n":"HVAC Pros Los Angeles","ott":1,"oto":1},{"n":"Pampered Paws Pet Grooming","ott":1,"oto":1},{"n":"Patrick Alley Handyman","ott":1,"oto":1},{"n":"Technology Networks","ott":1,"oto":1},{"n":"Uptown General Aesthetic Dentistry P A","ott":1,"oto":1},{"n":"Woodforest Family Chiropractic Clinic","ott":1,"oto":1}],"Katelyn Ankrom":[{"n":"Aberdeen Cemetery Assoc","ott":1,"oto":0},{"n":"Gleason's Salinas RV","ott":1,"oto":0},{"n":"Neilson Roy Plumbing","ott":1,"oto":0}],"Kellie Lester":[{"n":"Chatham Laschober","ott":1,"oto":1},{"n":"D &T Power Washing","ott":3,"oto":3},{"n":"East Cooper Lock & Safe","ott":1,"oto":1},{"n":"Lewis Brothers Inc","ott":1,"oto":1},{"n":"Mansfield Well Drilling Inc","ott":1,"oto":0},{"n":"Martin Landscaping","ott":1,"oto":1},{"n":"Metro Appliance Repair","ott":3,"oto":3},{"n":"Mikes Seal Coating & Services","ott":1,"oto":0},{"n":"Pleasure Pools","ott":2,"oto":2},{"n":"WBC Carpentry Corp","ott":1,"oto":0}],"Ashley Shaffer":[{"n":"A Dion & Son Floor Contractors","ott":1,"oto":0},{"n":"Absolute Roofing & Remodeling","ott":3,"oto":3},{"n":"Alaska Urgent Care LLC","ott":1,"oto":1},{"n":"Benny Electric Inc Benny Electric","ott":3,"oto":3},{"n":"Body Sculpt Skin Laser","ott":3,"oto":3},{"n":"Cement Hill Storage","ott":1,"oto":1},{"n":"Dream Work Diesel","ott":1,"oto":1},{"n":"Fox Valley Glass Inc","ott":1,"oto":1},{"n":"Garber Surveying Service PA","ott":3,"oto":3},{"n":"Just Cuz Plumbing LLP","ott":3,"oto":3},{"n":"Moody Construction Service","ott":1,"oto":1},{"n":"Pryor Automatic Fire Sprinkler Inc","ott":3,"oto":3},{"n":"Schaefers Stove & Spa","ott":3,"oto":3},{"n":"Seraphim Partners","ott":3,"oto":3},{"n":"Sound Decision","ott":1,"oto":0},{"n":"Testino Edward","ott":1,"oto":1},{"n":"Toms Superior Driving School Inc","ott":1,"oto":1},{"n":"Wasilla Medical Clinic","ott":1,"oto":1}],"Tyler Moeggenberg":[{"n":"AM PM Towing","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Alan Cherry Classic Interiors, LLC","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Fireplace Specialists LLC","ott":2,"oto":1},{"n":"GoodFellas Ristorante","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Newsom Fences","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Northwest Roofing & Construction LLC","d":[{"t":"15 Min Check-In","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Sarca Plumbing, Heating & Cooling","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Schnitzelbank","ott":2,"oto":1},{"n":"United Car Wash Gresham","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Up The Creek Heating & Air","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]}],"Elianny Tena Antigua":[{"n":"Burrows Heating & Cooling","ott":1,"oto":0},{"n":"Fleshers Fairview Health Care Center","ott":1,"oto":0},{"n":"G&S Tile Solutions LLC","ott":1,"oto":1},{"n":"JH Corner","ott":3,"oto":3}],"Saira Julian Guzman":[{"n":"Adama African Hair Braiding","ott":1,"oto":1},{"n":"Alert Door & Operator Co","ott":1,"oto":1},{"n":"E.M Imperial Auto Restoration","ott":1,"oto":1},{"n":"EZ Mini Storage","ott":1,"oto":1},{"n":"Kellets GC","ott":1,"oto":1}],"Peter Manalac":[{"n":"Julian's Appliance Centre","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"KEITH ROBERT HUGHES","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Lynfield Automotive Limited","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Stirling Floors","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"T & H Wreckers","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Victorian Chiropractic","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Misti Dixon":[{"n":"Dynamic Steam Carpet","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"R & G Plumbing And Drain Services Inc","d":[{"t":"Highlight Video","due":"5/26/2026","ov":true,"nw":false}]},{"n":"Riviera Floor Covering","ott":1,"oto":0},{"n":"Ryan Simmons","ott":2,"oto":1},{"n":"Witzke's Harry H Family Funeral Home Inc","ott":3,"oto":1}],"Heidi Torres Uribe":[{"n":"ArTron Heating And Air Conditioning","ott":1,"oto":1},{"n":"B Carroll Construction","ott":1,"oto":1},{"n":"Charlie Rice Roofing","ott":3,"oto":2},{"n":"Cobblestone Quality Shoe Repair","ott":2,"oto":2},{"n":"Dynamic Garage Door of Hibbing, LLC","ott":1,"oto":1},{"n":"Jones Animal Hospital","ott":2,"oto":2},{"n":"Soo Kool Air Conditioning","ott":1,"oto":1}],"Yessica Montero Urena":[{"n":"AAA London Chimney Sweep","ott":2,"oto":2},{"n":"Cura Healthcare Consultants","ott":1,"oto":1},{"n":"Josh's Mobile Mechanic","ott":1,"oto":1},{"n":"Pope County Title Co","ott":1,"oto":1},{"n":"Simplified Living Home Services","ott":1,"oto":0}],"Jathzelyn Elizabeth Fortuna Paulino":[{"n":"Apex Enterprise Roofing","ott":1,"oto":1},{"n":"Cross Cultural Communications","ott":1,"oto":1},{"n":"Door To Door Movers & Apartment Movers","ott":1,"oto":1},{"n":"Etheridge Hamlett & Murray LLP","ott":1,"oto":1},{"n":"Gingras Plumbing & Heating Inc","ott":1,"oto":1},{"n":"Henry Gitner Philatelists Inc","ott":1,"oto":1},{"n":"Park Ave Med Spa","ott":1,"oto":1},{"n":"Trent and Sons Roofing","ott":1,"oto":1},{"n":"United Car Wash","ott":2,"oto":2}],"Misty Decatur":[{"n":"Advance Insurance & Benefits","ott":1,"oto":0},{"n":"Arpys Construction & Remodeling","ott":1,"oto":0},{"n":"Auto Glass Xperts","ott":3,"oto":2},{"n":"Automotive Paint Supply APS","ott":1,"oto":1},{"n":"Centerscale Automation Hawaii Inc","ott":1,"oto":0},{"n":"Northern Arizona Roof Services LLC","ott":2,"oto":1},{"n":"Perfect World Pavers","d":[{"t":"Highlight Video","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Piazzau2019s Top Gun Coatings","ott":1,"oto":0},{"n":"Quality Transmission & Auto Repair","ott":1,"oto":1},{"n":"Ridge and Remedy Apotheracy, LLC","ott":3,"oto":3},{"n":"Taylor Regional Hospital","ott":3,"oto":3},{"n":"We Can Do More","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]}],"Barbara Larrosa Presinal":[{"n":"Charleston Premier Workz","ott":3,"oto":2},{"n":"Covert Water Heaters Inc","ott":1,"oto":1},{"n":"El Maya Mexican Grill","ott":3,"oto":2},{"n":"G Shelley Basement Waterproofing","ott":1,"oto":1},{"n":"GCS Heating and Cooling","ott":1,"oto":1},{"n":"Greenes Rolloff Service","ott":1,"oto":1},{"n":"Insights Eyecare, PA","ott":1,"oto":1},{"n":"Jennings Home Rejuvenation","ott":1,"oto":1},{"n":"Lembke Inc","ott":3,"oto":2},{"n":"Noah Autos","ott":1,"oto":1},{"n":"Smart House Remodeling","ott":1,"oto":1},{"n":"The Lukaart Agency, a Farm Bureau Agency","ott":2,"oto":1},{"n":"Ultra Steam Cleaning","ott":3,"oto":1},{"n":"Villegas roofing llc","ott":1,"oto":1}],"Ashley Vasquez Mena":[{"n":"Ideal Exteriors","ott":1,"oto":0},{"n":"Nashville Pizza Company","ott":1,"oto":0},{"n":"Ronco Tech Heating & Cooling LLC","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Sweers Roofing","ott":2,"oto":0},{"n":"Tug River Black Lung Clinic","ott":1,"oto":0}],"Warda Gul":[{"n":"Achieve Training & Assessment Services Pty","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"BL Microtek Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Chrysus Group","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Constructways Pty Ltd Trading as Stella Arden","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Creative 2000 Blinds & Awnings","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"D & M Brown Concreting","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Elite Body & Paint Pty Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Elite Tyre & Autocare Bacchus Marsh","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"GJM Balustrading","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"KCP Physiotherapy Paraparaumu","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Prodigy Design Plastics Pty Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Tauranga Roofing & Scaffolding","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Sarah Swanson":[{"n":"247 Protective Services","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"402 Castillo Contractor","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"A Head Of The Times Hair Design","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"A1 Pro Roofing LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"AAA Air Conditioning and Refrigeration","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"AAA OnSite Notary LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"AL Hansen Paint & Metal Shop Inc","ott":3,"oto":2},{"n":"ALL ABOUT PRESSURE","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Adele Home Health Agency","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"All Fence Supply Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"American Coins & Collectibles","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Amuse Media","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Art Tabasco & Sons Paving","ott":1,"oto":1},{"n":"Arthur J E & Associates Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Big Rig Truck Repair and Towing","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Bill's Fixit Shop","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Bradleys Construction","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Brenda Kashella Electrolysis Medical & Cosmetic","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"CCB Outdoor Living & Construction","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Caesars Italian Delcatessen","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Castle Rock HeatingAir","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Celco Community Credit Union","ott":1,"oto":1},{"n":"Clearwater Beach Boat Rentals","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Coast Pipe","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Colby Pacific Family Dentistry","ott":3,"oto":3},{"n":"ESC Cabinets","ott":1,"oto":1},{"n":"Electric Plus Inc.","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Elite Steel Building Systems","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Exquisite Body Clinic","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Forever Fence Solutions","ott":1,"oto":0},{"n":"Get It Now Print","ott":1,"oto":0},{"n":"Goodman General Construction Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"GraceWorks Bookkeeping LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Gregs Job Squad","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Guide Path Advisors","ott":3,"oto":2},{"n":"Hero Cleaning","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Hounds On The Hill","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"James Remodeling","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Jerry McLeod","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Josseline Carr","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Kevin Yul Wright JD - Business Loan Success Academy Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Lotsa Fence Options","ott":1,"oto":1},{"n":"Maria Rivero","ott":3,"oto":2},{"n":"Martin Electrical Services","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"McAlister McAlister & Nicklas PLLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Meilus Precision Therapy","ott":3,"oto":3},{"n":"Norvill Construction Co","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"One Day Safe Shower Va","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Paige Marie Photography","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Paradise Valley Land Solutions, LLC","ott":3,"oto":3},{"n":"Park Rapids Glass","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Pawfect Pawtions","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Polar Pure Water","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Precision Property Cleanup and Junk Removal","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"ProBend Corp","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Psychic Center Botanica","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"RCR Inspections Inc","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Raytech LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Rebecca Ortenzio Lee Orthodontics","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Rodger's Roll-Up Garage Doors","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"SOS Services Plumbing LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"SafeGuard Roofing and Siding","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Salty Breeze Rentals","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Sherry Smalling","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Simons Seamless Gutters","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Skylark North Glider Flight School","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Smith Brothers Funeral Home","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Snap2Throw Quarterback Academy","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"South Salem Mini Storage","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Speedy Pumping","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Stars and Stripes VA Benefits","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Stone Dumpster Rentals","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Super Electric","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Tavo Matic Inc","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"The Plumbing Company Of East Tennessee","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"The Yoga Experience LLC","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Toms RV Service & Sales","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Trank Insurance","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Triple H Paving","ott":1,"oto":0},{"n":"True Hand Roofing","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Two Men And A Snake","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Tycoon Solutions LLC","ott":1,"oto":1},{"n":"Unlimited Choices Greek Boutique Inc","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Uplifted Exteriors","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Vahe Dental","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Valora Behavior Support Centers","ott":2,"oto":2},{"n":"Webb Floors","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"White-Lavender Plumbing and Service","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Winkler Construction & Crane Co Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"sea pro home renovation","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]}],"Tracy-Ann Gaudencio":[{"n":"Cheshire Contractors","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Elite Pods","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Gina King Naturopathy","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Pro Pacific Concreting","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Smart Bookkeepers","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Karissa Hernandez":[{"n":"ABC Southwest Plumbing Inc","ott":1,"oto":0},{"n":"Air Conditioning & Heating Solutions LLC","ott":1,"oto":0},{"n":"Change your Thinking, Change Your Life LLC","ott":1,"oto":0},{"n":"Decca Recruiting LLC","ott":1,"oto":0},{"n":"Doyles Heat & Air Services LLC","ott":1,"oto":0},{"n":"Quality Restoration Services","ott":3,"oto":2},{"n":"Tidewater Landscape LLC","ott":1,"oto":0},{"n":"Triple A Sanitation","ott":3,"oto":2},{"n":"Veteran Floors Inc","ott":1,"oto":0}],"Anthony Yen":[{"n":"GTZ ROOFING","ott":3,"oto":3},{"n":"Keystone Pump & Well Service","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Tri-County Chimney Service","ott":1,"oto":0}],"Yolanda Ramirez":[{"n":"Best Electric","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Full Circle Payment Processing","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Mills Chiropractic","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Sarasota Golf Cart Sales - Custom Carts & Repairs","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]}],"Deivis Pena":[{"n":"A R McClung Construction Co","ott":3,"oto":3},{"n":"Aboite Boarding & Grooming","ott":2,"oto":1},{"n":"All About Kids Preschool","ott":1,"oto":1},{"n":"All Steamed Up carpet and Upholstry cleaning","ott":1,"oto":1},{"n":"Crown Tree Care Inc","ott":1,"oto":1},{"n":"Forever Young Landscaping","ott":3,"oto":3},{"n":"JB Hauling","ott":3,"oto":3},{"n":"Roberto's Tile & More","ott":1,"oto":1},{"n":"Zoom Business Brokers","ott":3,"oto":3}],"Damita Hill":[{"n":"A Plus Foundation LLC","d":[{"t":"Highlight Video","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Concord Heating & Air Conditioning Inc","ott":1,"oto":1},{"n":"Ohana Clean Air","d":[{"t":"Highlight Video","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Top Notch Moving Company","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Williams Auto Parts Inc","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]}],"Karmita Turner":[{"n":"Asphalt Services","ott":1,"oto":1},{"n":"Bemis Well Drilling & Water Conditioning, LLP","ott":1,"oto":1},{"n":"Creative Improvements Inc","ott":2,"oto":1},{"n":"Danny Odom Roofing","ott":3,"oto":3},{"n":"Himmelstein Louis","ott":1,"oto":0},{"n":"Holbert's Tree service","ott":2,"oto":2},{"n":"IntelliPEST","ott":2,"oto":1},{"n":"J Pop Landscaping","ott":1,"oto":1},{"n":"Law Office of Gayle A Belcher","ott":1,"oto":1},{"n":"Miss Miranda Bail Bond Services","ott":1,"oto":0},{"n":"New Look Exteriors","ott":2,"oto":1},{"n":"Quality Tree Care & Landscaping","ott":1,"oto":0},{"n":"Rays Septic Tank Service","ott":1,"oto":0},{"n":"Tab Mechanical Services","ott":1,"oto":1},{"n":"Valley Restaurant and Catering","ott":1,"oto":1}],"Kennedy Sanchez":[{"n":"4 C's Construction","ott":1,"oto":0},{"n":"On the Go Experience","ott":1,"oto":0}],"Felix Caba Jimenez":[{"n":"Action Counseling LLC","ott":1,"oto":1}],"Dorka Frias Lantigua":[{"n":"Copeland Fencing and Construction","ott":2,"oto":0},{"n":"Herbs Door Service","ott":1,"oto":1},{"n":"Insurance Answers Inc","ott":1,"oto":0},{"n":"Meehan's Lawn Service","ott":1,"oto":1},{"n":"Paxton Senior Insurance Service LLC","ott":1,"oto":0},{"n":"Rotterdam Heating","ott":3,"oto":1},{"n":"Ruth & Co. Events","ott":1,"oto":0},{"n":"Southern Living Exteriors","ott":2,"oto":0},{"n":"Wallace Heating & Air","ott":3,"oto":1}],"Kyle Dye":[{"n":"Adams Carpet Center","ott":1,"oto":0},{"n":"Allstate Paving Inc","ott":1,"oto":0},{"n":"Carmody James","ott":1,"oto":0},{"n":"Cynthia Poole","ott":3,"oto":1},{"n":"Final Exterminators","ott":1,"oto":1},{"n":"Formals Only Tuxedos","ott":3,"oto":1},{"n":"Granados Electrical Service","ott":3,"oto":1},{"n":"Gregorio's Pizzeria & Trattoria","ott":1,"oto":1},{"n":"Knabusch Insurance Services Inc","ott":1,"oto":0},{"n":"Lawrence Building Corp","ott":3,"oto":1},{"n":"Magidov CPA Firm","ott":3,"oto":1},{"n":"Majestic Jewelers","ott":3,"oto":1},{"n":"Mobility Plus Cincinnati East","ott":1,"oto":0},{"n":"Robinsons Paint & Wallpaper","ott":1,"oto":0},{"n":"SavMor Upholstery Co Inc","ott":1,"oto":0},{"n":"Speedy Locksmith Inc","ott":1,"oto":0},{"n":"Spokane Cosmetic Dentistry","ott":3,"oto":1},{"n":"Tess African Hair","ott":1,"oto":1},{"n":"The Caring Heart, LLC","ott":1,"oto":0},{"n":"The Grass Company of San Antonio","ott":1,"oto":0}],"Alejandro Rodriguez-Medina":[{"n":"Anwell Mobile Homes","ott":2,"oto":0},{"n":"Be Easy Bail Bonds","ott":1,"oto":0},{"n":"Blinds For Less","ott":1,"oto":0},{"n":"Boss Momma Boutique","ott":1,"oto":0},{"n":"Calvin Turner Roofing","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Celtic Moving & Storage Co","ott":3,"oto":0},{"n":"Chanler Agency Inc","ott":2,"oto":0},{"n":"Doug's Rooter Service","ott":1,"oto":0},{"n":"Duct -Tec","ott":1,"oto":0},{"n":"E-SQUARED ROOFING LLC","d":[{"t":"15 Min Check-In","due":"6/1/2026","ov":false,"nw":true}],"ott":3,"oto":1},{"n":"Fox Jewelers","ott":1,"oto":0},{"n":"Georges British American Auto Repair","ott":2,"oto":0},{"n":"Grand Slam Garage Door Services","ott":2,"oto":0},{"n":"Grass Roots Landscapes","ott":1,"oto":0},{"n":"Greg Smith","ott":1,"oto":0},{"n":"Hoffman Concrete, LLC","ott":3,"oto":1},{"n":"J & B Tree Services LLC","ott":1,"oto":0},{"n":"Logan Diving & Salvage","ott":3,"oto":0},{"n":"Manosh Singh and Associates","ott":2,"oto":0},{"n":"Mark Worleys Construction","ott":2,"oto":0},{"n":"Marks Roofing Company","ott":3,"oto":0},{"n":"O'Fallon Garage","ott":1,"oto":0},{"n":"Oakwood Landscaping LLC","ott":1,"oto":1},{"n":"Pro-Kleen","ott":2,"oto":0},{"n":"Ray donch Body werks Inc","d":[{"t":"15 Min Check-In","due":"6/1/2026","ov":true,"nw":true}],"ott":3,"oto":1},{"n":"Robbs Innova Construction","ott":1,"oto":0},{"n":"Rosso Nursery & Garden Center","ott":1,"oto":0},{"n":"Roy's Auto Body","ott":1,"oto":0},{"n":"Shaddai Construction","ott":2,"oto":0},{"n":"Stickley John R","ott":1,"oto":0},{"n":"Twin Cities Flag Source","ott":1,"oto":0},{"n":"VG Bail Bonds","ott":1,"oto":0},{"n":"Weis Landscaping Design","ott":1,"oto":0}],"Karen Capellan Tavarez":[{"n":"AAA Mini Storage","ott":1,"oto":1},{"n":"Best Tech Computer Service","ott":1,"oto":1},{"n":"BoozeeBar","ott":3,"oto":1},{"n":"Cali Roofing Inc","ott":3,"oto":1},{"n":"Cap Construction","ott":1,"oto":1},{"n":"Cecelia CookAssociates LLC","ott":1,"oto":1},{"n":"Challenge Family Fun Center","ott":1,"oto":1},{"n":"Enbalance Bodywork","ott":1,"oto":0},{"n":"Krystal Klear Cleaning Services","ott":1,"oto":1},{"n":"McLean Hardware Co, Inc","ott":1,"oto":1},{"n":"Peterman Bros Septic Service","ott":3,"oto":1},{"n":"Three Friends Tree Service","ott":1,"oto":1},{"n":"Vaca Valley Veterinary Hospital","ott":1,"oto":1},{"n":"Vital Essence Medical Spa","ott":3,"oto":1}],"Irina Larianni Molina Molina":[{"n":"Inaoly Auto Tech","d":[{"t":"Highlight Video","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Ivy Cremation Services of New York","ott":1,"oto":0},{"n":"Walters Mirror","ott":2,"oto":0},{"n":"Wells James DDS","ott":2,"oto":0}],"Wilson Mercedes":[{"n":"Achieve Wellness Drug Rehab New Jersey","ott":1,"oto":0},{"n":"Auto Images","ott":1,"oto":0},{"n":"Basement Systems Of NY","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Buckeye Crane & Hoist","ott":1,"oto":0},{"n":"Daystar Healthcare LLC","ott":1,"oto":0},{"n":"Fayette Veterinary Hospital","d":[{"t":"Highlight Video","due":"5/19/2026","ov":true,"nw":false}]},{"n":"Fresh Creek Plumbing & Heating","ott":1,"oto":0},{"n":"J Star Concrete","ott":1,"oto":0},{"n":"Joshua Paving","ott":1,"oto":0},{"n":"Kirsh Title Services","ott":1,"oto":0},{"n":"Next Level Athletes Born2Ball","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Northern Door Co Inc","ott":1,"oto":0},{"n":"On Point Pest Control","ott":1,"oto":0},{"n":"Premium Glass Tinting","ott":1,"oto":0},{"n":"SoMo Customs","ott":1,"oto":0},{"n":"The X-League","ott":1,"oto":0},{"n":"Zeeks Helpful Hands","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}],"ott":1,"oto":0}],"Matt Sword":[{"n":"DRH Construction Co., LLC","d":[{"t":"Highlight Video","due":"5/18/2026","ov":true,"nw":false}]},{"n":"Ericson Electric Inc","ott":1,"oto":0},{"n":"Levels Ahead Painting","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Pottenburgh Company","ott":1,"oto":0},{"n":"Root Revival Hair Restoration","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Sallie L Rubenzer Law Office","d":[{"t":"Highlight Video","due":"5/22/2026","ov":true,"nw":false}]},{"n":"Taxman Business Advisory Llc","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Texas Turf & Curb","ott":1,"oto":0},{"n":"The Chapel At Kerrville","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Veras King O Meats Inc 2 LLC","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Veras King O Meats Inc 3 LLC","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Wilson Blinds & Shutters","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Matt Daly":[{"n":"AR Electronics Systems Limited","ott":1,"oto":0},{"n":"De Silva Hebron","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Junior Explorers","ott":1,"oto":0},{"n":"Oleada Electrical","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"On Point Physio","ott":2,"oto":1},{"n":"Ulladulla Blinds & Home Improvements","ott":1,"oto":0}],"Nikita Siepen-Bowers":[{"n":"Brianna Tilt Trays & Towing Pty Ltd","ott":1,"oto":1},{"n":"Kenny's Painting Crew","d":[{"t":"Highlight Video","due":"5/31/2026","ov":true,"nw":false}]},{"n":"Mildura First Aid Services","ott":1,"oto":1},{"n":"Total Safe Compliance Group","d":[{"t":"Highlight Video","due":"5/31/2026","ov":true,"nw":false}]}],"Michael Furlong":[{"n":"AA Electric, Ltd.","d":[{"t":"Highlight Video","due":"5/29/2026","ov":true,"nw":false}]},{"n":"ASM Irrigation","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"All American Pro Paving & Roofing","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Earnest Well Drilling Inc","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Sati Ananda Pimentel Malespin":[{"n":"AD Dental","ott":1,"oto":0},{"n":"Abercrombie Transmission","ott":1,"oto":1},{"n":"Cenla Plumbing Repair LLC","ott":1,"oto":1},{"n":"Drive Line Service & Radiator King Inc","ott":1,"oto":0},{"n":"Fleetwood Foot & Ankle Center","ott":1,"oto":1},{"n":"Fretters Piano Service","ott":1,"oto":1},{"n":"Hawaii Gold Buyers Exchange","ott":2,"oto":1},{"n":"Lakeview Garden Center & Landscaping","ott":1,"oto":1},{"n":"United Vapor Barrier & Floors","ott":2,"oto":1}],"Merve (MJ) Brielmann":[{"n":"209 Country Shoppe","ott":3,"oto":1},{"n":"A Total Fire Protection Co","ott":1,"oto":0},{"n":"Aarons Plumbing Inc","ott":2,"oto":0},{"n":"All Purpose Well Drilling","ott":3,"oto":1},{"n":"Armuchee Self Storage","ott":1,"oto":1},{"n":"Axis Doors","ott":1,"oto":1},{"n":"Big Foot Air Quality LLC","ott":2,"oto":1},{"n":"Blue Print Specialties Inc","ott":2,"oto":1},{"n":"Carter Heating & Air","ott":1,"oto":1},{"n":"DRAIN SQUAD NYC INC","ott":1,"oto":1},{"n":"Gabriele Masonry & Waterproofing","ott":1,"oto":0},{"n":"Jeannie Pierce Insurance Agency","ott":1,"oto":1},{"n":"Kuhn's Equipment Repair","ott":3,"oto":1},{"n":"Savannah Bail Bonding","ott":3,"oto":1},{"n":"Sevey Norm Well Drilling Inc","ott":3,"oto":2},{"n":"Shenberg Construction","ott":1,"oto":0},{"n":"Techworx LLC","ott":3,"oto":1},{"n":"Tri County Fuels Inc","ott":3,"oto":2},{"n":"Unique II Worldwide","ott":1,"oto":0},{"n":"Youngrens Inc","ott":1,"oto":0}],"Steven Saunders":[{"n":"5 Stars General Contactor Inc","ott":1,"oto":1},{"n":"A D Sonbert Security Systems Inc","ott":2,"oto":1},{"n":"Aeroclean NWA","ott":1,"oto":1},{"n":"Brians Wow Plumbing","ott":3,"oto":3},{"n":"DJ and Associates","ott":1,"oto":1},{"n":"Dennis Green's Paving","ott":1,"oto":0},{"n":"Farleys Roofing INC","ott":2,"oto":1},{"n":"Finly Family Insurance","ott":3,"oto":3},{"n":"Hoffman & Hoffman","ott":1,"oto":1},{"n":"Island Solar Service","ott":1,"oto":0},{"n":"Michael Lloyd Bail Bonds","ott":1,"oto":0},{"n":"On Demand Crane Service","ott":2,"oto":1},{"n":"Pioneer Overhead Door 3G","ott":1,"oto":1},{"n":"Wiltse Towing LLC","ott":1,"oto":0}],"Tyler Popplewell":[{"n":"B&C Remodeling and Flooring","ott":3,"oto":3},{"n":"GET ER DONE LLC","ott":1,"oto":0},{"n":"P510 Coach","ott":1,"oto":1},{"n":"Roto-Rooter Las Cruces","ott":1,"oto":1},{"n":"Toddler Barrier","ott":1,"oto":1},{"n":"Westerly Paints","ott":1,"oto":1}],"Samuel Frias De Paula":[{"n":"Delcon Electric","ott":1,"oto":1},{"n":"Kathy Bleier Coaching LLC","ott":1,"oto":0}],"Florence Francois Nova":[{"n":"Total Home Renovation","ott":1,"oto":1}],"Stacy Roers":[{"n":"Bradham David Dr","ott":1,"oto":0},{"n":"PHD Bathroom Remodeling","ott":1,"oto":0}],"Victor Abner Moscoso Fernandez":[{"n":"Big's RV Service","ott":1,"oto":0},{"n":"Jay Kent Construction LLC","ott":1,"oto":0},{"n":"Nixa Lawn Service","ott":1,"oto":0},{"n":"Patio Furniture Cushions Inc","ott":1,"oto":0},{"n":"Rain Flow Of Indianapolis","ott":1,"oto":0}],"Ellise Payne":[{"n":"A Hepworth Electrical Pty Ltd","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"AOTEA TIMARU LIMITED","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Aaron Slape, Acupuncturist","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Bayside Quality Furniture Restorations","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Character Cabins","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Concept Fencing MC","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Junction Tyre & Auto Services","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Oslands Independants Carpets","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"SJ Firewood Limited","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]}],"Joseph Guillermo Carmona Garcia":[{"n":"A Beauty In The Beast Pet Grooming","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"A Tri Cities Connection","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Airtight SD","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"All Seasons Equipment","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"American Countertop Experts Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"American West Services","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"BWF Home Solutions","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Benson Chiropractic Center","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Blue Ridge Ear Nose Throat & Plastic Surgery","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Boyd Construction","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Breaking Free Counseling","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"ChillTex LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Cullina Management","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"EZ Sewer Cleaning","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Ed K's Ceramic Tile LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Elevate Home Services","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"EuroStone LLC","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Evergreen Insurance Advisors - Medicare & Health Insurance","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Firestone Concrete Coatings","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Flemingmovingllc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Foreman's Quality Painting Services","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Ganan Landscaping","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Greg Munning CFI","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Guild Mortgage- Stephany Kuennen","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Hammond and Sons Lawn Care and Landscaping","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Handi Built, LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Hicks Trading Station","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Holcomb Concrete Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Homemaker Landscaping","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Idaho Construction Company","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"JR Tree Works","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Janina Elite Medispa","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Jason Diller","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Krueger Plumbing LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Lewis Dean Drapery And Blinds","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Long Island Creative Contracting, Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Longworth Bail Bonds","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"MWC Construction Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Mels Tree Service","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Mokelumne Federal Credit Union","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"NRT","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"New Chapter Senior Living Placement","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Next Level Roofing and Remodeling","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Nickel, Greg & Tamara","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Northshore Dermatology Center","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Paralegals Unlimited, LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Platinum Elite Janitorial","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Postal World","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Price, McCluer & Plachecki","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Proctor's Precision Fence","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Quality Termite and Pest Control, LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Reflecto Signs & Graphics","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Right Print","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Rodriguez Ross A Attorney At Law","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Rogue Lock & Key LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"S & J Residential Roofing","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"SAFE-T CHOICE INC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"SEMO lawn solutions LLC","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Siege Productions","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Southern Auto Paint & Body Shop","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Stars and Pipes Plumbing","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Stevens Concrete","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Storm Drains Hawaii","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Sugarbush Tavern","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"SweetiePumps","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"TD Contractors LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"The Bodhi tree Holistic Health Solutions","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"The Childcare Concierge Agency","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Wagner Kuntz & Grabouski","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Wealth Warden Partners","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]}],"Lauren Carter":[{"n":"CSA Roofing, Painting & General Contracting","ott":3,"oto":1},{"n":"Caldwell Reubens Drilling Inc","ott":3,"oto":1},{"n":"Children's Corner","ott":3,"oto":1},{"n":"Comfort Family Dentistry","ott":2,"oto":1},{"n":"Dulando Screen & Awning Inc","ott":1,"oto":1},{"n":"Healthy Foot Spa","ott":1,"oto":0},{"n":"Lapcomp Computers","ott":1,"oto":0},{"n":"Lex Plant Farm","ott":2,"oto":1},{"n":"Marketside Chiropractic","ott":2,"oto":1},{"n":"Quest Electric Inc","ott":1,"oto":0},{"n":"Taylor Aution Realty","ott":1,"oto":1},{"n":"Traveltimesawait","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]}],"Johnny Cornielle":[{"n":"Brandons Awards & Engraving","ott":1,"oto":1},{"n":"Busseys Flea Market","ott":3,"oto":1},{"n":"Early Years The","ott":1,"oto":0},{"n":"GapArmour","ott":1,"oto":1},{"n":"Mechanical Energy Systems","ott":3,"oto":1},{"n":"Neptune Pool Management","ott":1,"oto":1},{"n":"Network Financial","ott":1,"oto":1},{"n":"PSI Seamless Gutters","ott":1,"oto":1},{"n":"TRUE CRAFT FOUNDATION REPAIR & WATERPROOFING","ott":1,"oto":1},{"n":"Williams Family Medicine","ott":3,"oto":1}],"Dave Crisler":[{"n":"AK Firewood","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Aqua Dash","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Bendigo Tyre & Auto","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Best of the Best Reblocking","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"H & A Training & Supplies","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"KIWIVAC CENTRAL VACUUM SYSTEMS (1999) LIMITED","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"M1 business system","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Perforge","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"South Seas Construction Ltd","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"The Rose City Limousine","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Top Mix Construction","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Tree Limits Pty Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"oxen limited","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Scott Mather":[{"n":"Advantage Life & Health","ott":1,"oto":1},{"n":"Head to Toe Reset Spa","ott":1,"oto":1},{"n":"MVP RIDES","ott":1,"oto":1},{"n":"Magic Refrigeration","ott":1,"oto":1},{"n":"Wilcox Transmission","ott":1,"oto":1}],"Chelsea Dingus":[{"n":"A Fresh Cut Landscaping","ott":1,"oto":0},{"n":"AARO Fence Inc","ott":1,"oto":0},{"n":"Access & Alarm Company Inc","ott":2,"oto":1},{"n":"Bullet Hole Annex","ott":1,"oto":1},{"n":"Corporal Lawn Service","ott":2,"oto":0},{"n":"Dan Green","ott":1,"oto":0},{"n":"GREAT TOUCH BEHAVORAL HEALTH","ott":1,"oto":0},{"n":"In and out garage doors","ott":1,"oto":1},{"n":"Patriot Sunrooms","ott":1,"oto":0},{"n":"Radon Raiders","ott":1,"oto":0},{"n":"Sanderson & De Haan Lawn Sprinkling","ott":1,"oto":0},{"n":"The John Wood Insurance Agency Inc.","ott":1,"oto":0}]};

const COACHES = [
  {e:"odirlm01@thryv.com",      n:"Mia O\u2019Dirling",   t:"The Dominican Dream Team", col:"#FF5000"},
  {e:"chase.boyd@thryv.com",    n:"Chase Boyd",            t:"Boyd Meets World",          col:"#4A5D8C"},
  {e:"elizabeth.white@thryv.com",n:"Elizabeth White",      t:"White Wave Warriors",       col:"#808080"},
  {e:"kendra.morelli@thryv.com", n:"Kendra Morelli",       t:"Team Thryv-More(lli)",      col:"#29355D"},
  {e:"trisha.stalnaker@thryv.com",n:"Trisha Stalnaker",    t:"Team Status Engaged",       col:"#E03000"},
  {e:"aaron.taylor@thryv.com",  n:"Aaron Taylor",          t:"Team Aurorians",            col:"#5378FC"},
];

const MANAGERS = [
  {id:"carrie", n:"Carrie Reece",  coaches:["odirlm01@thryv.com","chase.boyd@thryv.com","elizabeth.white@thryv.com"]},
  {id:"jake",   n:"Jake Baldwin",  coaches:["kendra.morelli@thryv.com","trisha.stalnaker@thryv.com","aaron.taylor@thryv.com"]},
];

const TEAM_COLS = {
  "The Dominican Dream Team":"#FF5000","Boyd Meets World":"#4A5D8C",
  "White Wave Warriors":"#808080","Team Thryv-More(lli)":"#29355D",
  "Team Status Engaged":"#E03000","Team Aurorians":"#5378FC",
};

const TEAM_LOGOS = {
  "Boyd Meets World":"data:image/png;base64,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",
  "Team Status Engaged":"data:image/png;base64,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",
  "The Dominican Dream Team":"data:image/png;base64,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",
  "White Wave Warriors":"data:image/png;base64,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",
  "Team Aurorians":"data:image/png;base64,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",
  "Team Thryv-More(lli)":"data:image/jpeg;base64,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"
};

// ── ORG CHART (source of truth for coach assignments) ──────────────────────
// Updated from: https://docs.google.com/spreadsheets/d/e/...?gid=352807280
// Key changes vs previous: Dorka Frias Lantigua → Kendra Morelli,
//   Steven Saunders → Kendra Morelli, Sidharta Goris removed from roster

const NAME_NORM = {
  "darling danais":"Darling Danais Santos Taveras","darling taveras":"Darling Danais Santos Taveras","darling danais santos taveras":"Darling Danais Santos Taveras",
  "heidi torres":"Heidi Torres Uribe","heidi uribe":"Heidi Torres Uribe","heidi torres uribe":"Heidi Torres Uribe",
  "irina larianni":"Irina Larianni Molina Molina","irina molina":"Irina Larianni Molina Molina","irina molina molina":"Irina Larianni Molina Molina","irina larianni molina molina":"Irina Larianni Molina Molina","irina molina molina":"Irina Larianni Molina Molina",
  "jathzelyn elizabeth":"Jathzelyn Elizabeth Fortuna Paulino","jathzelyn fortuna":"Jathzelyn Elizabeth Fortuna Paulino","jathzelyn fortuna paulino":"Jathzelyn Elizabeth Fortuna Paulino","jathzelyn paulino":"Jathzelyn Elizabeth Fortuna Paulino","jazz fortuna":"Jathzelyn Elizabeth Fortuna Paulino","jathzelyn elizabeth fortuna paulino":"Jathzelyn Elizabeth Fortuna Paulino",
  "johnny cornielle montas":"Johnny Cornielle","johnny cornielle":"Johnny Cornielle",
  "joseph carmona":"Joseph Guillermo Carmona Garcia","joseph carmona garcia":"Joseph Guillermo Carmona Garcia","joseph garcia":"Joseph Guillermo Carmona Garcia","joseph guillermo":"Joseph Guillermo Carmona Garcia","joseph guillermo carmona garcia":"Joseph Guillermo Carmona Garcia",
  "sam frias":"Samuel Frias De Paula","sam frias de paula":"Samuel Frias De Paula","samuel frias":"Samuel Frias De Paula","samuel paula":"Samuel Frias De Paula","samuel frias de paula":"Samuel Frias De Paula",
  "sati ananda":"Sati Ananda Pimentel Malespin","sati malespin":"Sati Ananda Pimentel Malespin","sati pimentel":"Sati Ananda Pimentel Malespin","sati pimentel malespin":"Sati Ananda Pimentel Malespin","sati ananda pimentel malespin":"Sati Ananda Pimentel Malespin",
  "victor abner":"Victor Abner Moscoso Fernandez","victor fernandez":"Victor Abner Moscoso Fernandez","victor moscoso":"Victor Abner Moscoso Fernandez","victorabner moscoso fernandez":"Victor Abner Moscoso Fernandez","victor abner moscoso fernandez":"Victor Abner Moscoso Fernandez",
  "wilson mercedes":"Wilson Mercedes",
  "yessica montero":"Yessica Montero Urena","yessica urena":"Yessica Montero Urena","yessica montero urena":"Yessica Montero Urena","yessica montero":"Yessica Montero Urena",
  "barbara larrosa":"Barbara Larrosa Presinal","barbara presinal":"Barbara Larrosa Presinal","barbara larrosa presinal":"Barbara Larrosa Presinal",
  "deivis pe\u00f1a":"Deivis Pena","deivis pena":"Deivis Pena","deivis peña":"Deivis Pena","deivis pẽa":"Deivis Pena",
  "eric johnson":"Eric Johnson","kyle dye":"Kyle Dye","sarah swanson":"Sarah Swanson",
  "luis aguasvivas":"Luis Aguasvivas Peralta","luis peralta":"Luis Aguasvivas Peralta","luis aguasvivas peralta":"Luis Aguasvivas Peralta",
  "juan liberato paula":"Juan Liberato","juan liberato":"Juan Liberato",
  "elianny antigua":"Elianny Tena Antigua","elianny tena":"Elianny Tena Antigua","elianny tena antigua":"Elianny Tena Antigua",
  "damita hill":"Damita Hill","anthony yen":"Anthony Yen","april hall":"April Hall",
  "katelyn ankrom":"Katelyn Ankrom","kennedy sanchez":"Kennedy Sanchez","michael furlong":"Michael Furlong",
  "yolanda ramirez-drake":"Yolanda Ramirez","yolanda ramirez":"Yolanda Ramirez",
  "florence francois":"Florence Francois Nova","florence nova":"Florence Francois Nova","francois nova":"Florence Francois Nova","florence francois":"Florence Francois Nova","florence francois nova":"Florence Francois Nova",
  "rossi tejeda":"Rossi Valerio Tejeda","rossi valerio":"Rossi Valerio Tejeda","rossi valerio tejeda":"Rossi Valerio Tejeda",
  "alejandro rodriguez medina":"Alejandro Rodriguez-Medina","alejandro rodriguez-medina":"Alejandro Rodriguez-Medina",
  "chelsea dingus":"Chelsea Dingus",
  "dorka frias":"Dorka Frias Lantigua","dorka lantigua":"Dorka Frias Lantigua","dorka frias lantigua":"Dorka Frias Lantigua",
  "karmita k turner":"Karmita Turner","karmita turner":"Karmita Turner",
  "lauren carter":"Lauren Carter","libby booher":"Libby Booher","misti dixon":"Misti Dixon","misty decatur":"Misty Decatur",
  "saira guzman":"Saira Julian Guzman","saira julian":"Saira Julian Guzman","saira julian guzman":"Saira Julian Guzman",
  "scott mather":"Scott Mather","steven saunders":"Steven Saunders","steve saunders":"Steven Saunders",
  "lauren carter":"Lauren Carter","lauren fogg":"Lauren Carter",
  "katelyn ankrom":"Katelyn Ankrom","katelyn anderson":"Katelyn Ankrom",
  "rafael sencion sencion":"Rafael Sencion Sencion","rafael sencion leon":"Rafael Sencion Sencion","rafael sencion":"Rafael Sencion Sencion",
  "ashley shaffer":"Ashley Shaffer",
  "ashley mena":"Ashley Vasquez Mena","ashley vasquez":"Ashley Vasquez Mena","ashley vasquez mena":"Ashley Vasquez Mena",
  "karen capellan":"Karen Capellan Tavarez","karen tavarez":"Karen Capellan Tavarez","karen capellan tavarez":"Karen Capellan Tavarez",
  "karissa hernandez":"Karissa Hernandez","kellie lester":"Kellie Lester","mark velazquez":"Mark Velazquez",
  "merve (mj)":"Merve (MJ) Brielmann","merve brielmann":"Merve (MJ) Brielmann","mj brielmann":"Merve (MJ) Brielmann","merve (mj) brielmann":"Merve (MJ) Brielmann",
  "rafael sencion":"Rafael Sencion Sencion","rafael sencion sencion":"Rafael Sencion Sencion",
  "stacy miron":"Stacy Roers","stacy roers":"Stacy Roers",
  "taylor kidd":"Taylor Kidd",
  "felix caba":"Felix Caba Jimenez","felix jimenez":"Felix Caba Jimenez","felix caba jimenez":"Felix Caba Jimenez",
  "dave crisler":"Dave Crisler","david crisler":"Dave Crisler",
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
  "darling danais santos taveras":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI",reg:"DR"},
  "heidi torres":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "heidi uribe":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "heidi torres uribe":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI",reg:"DR"},
  "irina larianni":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "irina molina":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "irina molina molina":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "irina larianni molina molina":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI",reg:"DR"},
  "jathzelyn elizabeth":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "jathzelyn fortuna":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "jathzelyn fortuna paulino":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "jathzelyn paulino":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "jazz fortuna":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "jathzelyn elizabeth fortuna paulino":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI",reg:"DR"},
  "johnny cornielle":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII",reg:"DR"},
  "johnny cornielle montas":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII"},
  "joseph carmona":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "joseph carmona garcia":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "joseph garcia":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "joseph guillermo":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "joseph guillermo carmona garcia":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI",reg:"DR"},
  "sam frias":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII"},
  "sam frias de paula":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII"},
  "samuel frias":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII"},
  "samuel paula":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII"},
  "samuel frias de paula":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII",reg:"DR"},
  "sati ananda":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "sati malespin":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "sati pimentel":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "sati pimentel malespin":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "sati ananda pimentel malespin":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI",reg:"DR"},
  "victor abner":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "victor fernandez":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "victor moscoso":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "victorabner moscoso fernandez":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "victor abner moscoso fernandez":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI",reg:"DR"},
  "wilson mercedes":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMII",reg:"DR"},
  "yessica montero":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "yessica urena":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI"},
  "yessica montero urena":{c:"odirlm01@thryv.com",t:"The Dominican Dream Team",r:"CSMI",reg:"DR"},
  // ── Chase Boyd — Boyd Meets World ─────────────────────────────────────
  "barbara larrosa":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMII"},
  "barbara presinal":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMII"},
  "barbara larrosa presinal":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMII",reg:"DR"},
  "deivis pena":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI",reg:"DR"},
  "deivis pe\u00f1a":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI",reg:"DR"},
  "eric johnson":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMII",reg:"US"},
  "kyle dye":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMIII",reg:"US"},
  "sarah swanson":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI",reg:"US"},
  "luis aguasvivas":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "luis peralta":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "luis aguasvivas peralta":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI",reg:"DR"},
  "juan liberato paula":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "juan liberato":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI",reg:"DR"},
  "elianny antigua":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "elianny tena":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "elianny tena antigua":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI",reg:"DR"},
  // ── Elizabeth White — White Wave Warriors ─────────────────────────────
  "damita hill":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII",reg:"US"},
  "anthony yen":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII",reg:"US"},
  "april hall":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII",reg:"US"},
  "katelyn ankrom":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII",reg:"US"},
  "kennedy sanchez":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII",reg:"US"},
  "michael furlong":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII",reg:"US"},
  "yolanda ramirez-drake":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII",reg:"US"},
  "yolanda ramirez":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII",reg:"US"},
  "florence francois":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI",reg:"DR"},
  "florence nova":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI",reg:"DR"},
  "francois nova":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI",reg:"DR"},
  "florence francois nova":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI",reg:"DR"},
  "rossi tejeda":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI",reg:"DR"},
  "rossi valerio":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI",reg:"DR"},
  "rossi valerio tejeda":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI",reg:"DR"},
  // ── Kendra Morelli — Team Thryv-More(lli) ────────────────────────────
  "alejandro rodriguez medina":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"SSMII"},
  "alejandro rodriguez-medina":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"SSMII",reg:"US"},
  "chelsea dingus":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"SSMII",reg:"US"},
  "dorka frias":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMI"},
  "dorka lantigua":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMI"},
  "dorka frias lantigua":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMI",reg:"DR"},
  "karmita k turner":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMII"},
  "karmita turner":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMII",reg:"US"},
  "lauren carter":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMII",reg:"US"},
  "libby booher":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMIII",reg:"US"},
  "misti dixon":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMIII",reg:"US"},
  "misty decatur":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"SSMI",reg:"US"},
  "saira guzman":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMI"},
  "saira julian":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMI"},
  "saira julian guzman":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMI",reg:"DR"},
  "scott mather":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"SSMII",reg:"US"},
  "steven saunders":{c:"kendra.morelli@thryv.com",t:"Team Thryv-More(lli)",r:"CSMII",reg:"US"},
  // ── Trisha Stalnaker — Team Status Engaged ────────────────────────────
  "ashley shaffer":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMII",reg:"US"},
  "ashley mena":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "ashley vasquez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "ashley vasquez mena":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI",reg:"DR"},
  "karen capellan":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "karen tavarez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "karen capellan tavarez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI",reg:"DR"},
  "karissa hernandez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMII",reg:"US"},
  "kellie lester":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMIII",reg:"US"},
  "mark velazquez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"SSMII",reg:"US"},
  "merve (mj)":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMII"},
  "merve brielmann":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMII"},
  "mj brielmann":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMII"},
  "merve (mj) brielmann":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMII",reg:"US"},
  "rafael sencion":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "rafael sencion sencion":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI",reg:"DR"},
  "stacy miron":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMIII"},
  "stacy roers":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMIII",reg:"US"},
  "taylor kidd":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI",reg:"US"},
  "felix caba":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "felix jimenez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI"},
  "felix caba jimenez":{c:"trisha.stalnaker@thryv.com",t:"Team Status Engaged",r:"CSMI",reg:"DR"},
  // ── Aaron Taylor — Team Aurorians ─────────────────────────────────────
  "dave crisler":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMII",reg:"ANZ"},
  "david crisler":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMII",reg:"ANZ"},
  "ellise payne":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI",reg:"ANZ"},
  "indu vijay":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI",reg:"ANZ"},
  "matt daly":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI",reg:"ANZ"},
  "nikita siepen bowers":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "nikita siepen-bowers":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI",reg:"ANZ"},
  "peter manalac":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI",reg:"ANZ"},
  "sakshi mahalwal":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI",reg:"ANZ"},
  "sylvia":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "sylvia appla":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI",reg:"ANZ"},
  "tracy ann gaudencio":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI"},
  "tracy-ann gaudencio":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI",reg:"ANZ"},
  "warda gul":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMI",reg:"ANZ"},
  "zoltan rudolf":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMII",reg:"ANZ"},
};

// ── DEACTIVATED CSMs — filter out from all live data feeds ─────────────────
const DEACTIVATED_CSMS = new Set([
  "sidharta goris",
  "sidharta",
  "matt sword",
  "tyler popplewell",
  "tyler moeggenberg",
  "rossi valerio tejeda",
  "rossi valerio",
]);

function lk(n) { return n ? ROSTER[n.toLowerCase().trim()] || null : null; }
function region(n) {
  const i = lk(n);
  if (i && i.reg) return i.reg;
  // alias entry may lack reg — try canonical name
  const canonical = norm(n);
  if (canonical !== n) {
    const i2 = lk(canonical);
    if (i2 && i2.reg) return i2.reg;
  }
  return null;
}
function norm(n) { return NAME_NORM[n.toLowerCase().trim()] || n.trim(); }
function dispName(n) {
  if (!n) return n;
  if (n === "Merve (MJ) Brielmann") return "MJ Brielmann";
  if (n === "Lauren Carter") return "Lauren Fogg"; // married name change — same person, same internal ID
  return n;
}
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

// Validate that a parsed name is actually a real CSM in our roster.
// This prevents free-text notes, account names, cadence names, or other
// stray data from being treated as CSM names.
function isValidCSM(name) {
  if (!name || name.length < 4) return false;
  // Filter out deactivated CSMs
  if (DEACTIVATED_CSMS.has(name.toLowerCase().trim())) return false;
  // Must resolve to a known ROSTER entry (exact or via NAME_NORM alias)
  if (ROSTER[name.toLowerCase().trim()]) return true;
  // Also allow if norm() mapped it to a canonical name that's in ROSTER
  const normed = norm(name);
  return !!ROSTER[normed.toLowerCase().trim()];
}

function mapRev(rows) {
  const by = {};
  rows.forEach(r => {
    const raw = r["CSM Name"]||r["csm name"]||r["csm_name"]||"";
    const name = norm(raw.trim());
    if (!name || !isValidCSM(raw.trim())) return;
    // Support both old and new JotForm column name variations
    const mrr  = pm(r["MRR $ Added"]||r["MRR $"]||r["MRR"]||0);
    const otr  = pm(r["OTR $ Added"]||r["OTR $"]||r["OTR"]||0);
    const total= pm(r["Total Revenue Added"]||r["Total Revenue"]||r["Revenue"]||0);
    const team = r["CSM Team!"]||r["CSM Team! "]||r["csm_team"]||"";
    const nr   = (r["Non-Revenue Integrations"]||r["Non Revenue Integrations"]||"").trim();
    const biz  = (r["Business Name"]||r["business_name"]||"").trim();
    const intType = r["Type of Integration"]||r["type_of_integration"]||"";
    const mrrInt  = r["MRR Integration"]||r["mrr_integration"]||"";
    const quarter = (r["Quarter for Consideration"]||r["Quarter"]||"").trim();
    if (!by[name]) by[name] = {name, team, mrr:0, otr:0, total:0, nonrev:0, accts:[], quarters:new Set()};
    by[name].mrr   += mrr;
    by[name].otr   += otr;
    by[name].total += total;
    if (nr) by[name].nonrev++;
    if (quarter) by[name].quarters.add(quarter);
    if (biz) by[name].accts.push({
      b: biz,
      t: intType,
      m: mrr,
      o: otr,
      n: nr,
      mi: mrrInt,
      q: quarter,
    });
  });
  return Object.values(by);
}

function mapEmail(rows) {
  const by = {};
  rows.forEach(r => {
    const raw = r["Touchpoint: Owner Name \u2191"]||r["Touchpoint: Owner Name"]||r["Name"]||"";
    const name = norm(raw.trim());
    if (!name||name==="Total"||!isValidCSM(raw.trim())) return;
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
  let currentCSM = "";

  // The new report format has:
  //   - Several metadata rows at top before real headers
  //   - CSM name only on the FIRST row of their group (carry-forward)
  //   - Subtotal rows: col "Touchpoint: Record Type" = "Subtotal" or col "" = "Subtotal"/"Count"
  //   - Real data rows have a Status value (Completed, In Progress, Skipped, Removed)
  // Both old flat format and new grouped format are handled here.

  const SKIP_STATUSES = new Set(["Removed"]);
  const SUBTOTAL_MARKERS = new Set(["Subtotal","Count","Sum","Grand Total"]);

  rows.forEach(r => {
    // Detect subtotal/metadata rows — skip them
    const allVals = Object.values(r).map(v => String(v||"").trim());
    if (allVals.some(v => SUBTOTAL_MARKERS.has(v))) return;

    // Try to get CSM name — new format uses "Touchpoint: Owner Name  ↑" (double space + arrow)
    const rawName = (
      r["Touchpoint: Owner Name  \u2191"] ||  // new format (double space)
      r["Touchpoint: Owner Name \u2191"]  ||  // old format (single space)
      r["Touchpoint: Owner Name"]         ||
      r["name"] || r["Name"] || ""
    ).trim();

    // If this row has a valid CSM name, update currentCSM
    if (rawName && rawName !== "None" && rawName !== "") {
      const candidate = norm(rawName);
      if (isValidCSM(rawName)) currentCSM = candidate;
    }

    // Status tells us what happened
    const status = (r["Cadence Member: Status"] || r["Status"] || "").trim();

    // Must have a status and a current CSM to count
    if (!status || !currentCSM) return;
    // Skip pure metadata rows that have no meaningful status
    if (SUBTOTAL_MARKERS.has(status)) return;

    if (!by[currentCSM]) by[currentCSM] = {name:currentCSM, total:0, completed:0, removed:0, skipped:0};
    by[currentCSM].total++;

    if (status === "Removed") {
      by[currentCSM].removed++;
    } else if (status === "Skipped") {
      by[currentCSM].skipped++;
    } else {
      // Completed, In Progress with outcome, etc. all count as completed touchpoints
      by[currentCSM].completed++;
    }
  });

  return Object.values(by).map(d => {
    const pct = d.total > 0 ? d.completed / d.total : 0;
    return {name: d.name, total: d.total, pct};
  }).filter(d => d.total > 0);
}


function mapDue(rows) {
  const by = {};
  // Exact column names confirmed from sheet screenshot:
  // Col A: "Touchpoint: Touchpoint Name"  e.g. "Highlight Video", "Email"
  // Col B: "Due Date/Time"                e.g. "6/11/2026 8:00"
  // Col C: "Overdue"                      1 = overdue
  // Col F: "New Today"                    1 = new today
  // Col H: "Cadence Member: Account"      account/business name
  // Col I: "Cadence Member: Assigned"     CSM name
  rows.forEach(r => {
    const raw = (r["Cadence Member: Assigned"] || "").trim();
    if (!raw) return;
    const name = norm(raw) || raw;
    if (!name || name.length < 3) return;
    if (!isValidCSM(raw)) return;

    const taskType = (r["Touchpoint: Touchpoint Name"] || r["Touchpoint: Record Type"] || "Task").trim();
    const dueRaw   = (r["Due Date/Time"] || "").trim();
    const acctName = (r["Cadence Member: Account"] || "").trim();
    const ov = String(r["Overdue"] || "").trim();
    const nt = String(r["New Today"] || "").trim();
    const isOverdue  = ov === "1" || ov.toLowerCase() === "true";
    const isNewToday = nt === "1" || nt.toLowerCase() === "true";

    // Format due date nicely e.g. "6/11/2026 8:00" → "6/11/2026"
    const dueDateStr = dueRaw ? dueRaw.split(" ")[0] : "";

    if (!by[name]) by[name] = {name, due:0, overdue:0, newToday:0, accounts:{}};
    by[name].due++;
    if (isOverdue)  by[name].overdue++;
    if (isNewToday) by[name].newToday++;

    // Group tasks by account so CSM detail can show them
    if (acctName) {
      if (!by[name].accounts[acctName]) by[name].accounts[acctName] = [];
      by[name].accounts[acctName].push({
        t: taskType,
        due: dueDateStr,
        ov: isOverdue,
        nw: isNewToday,
      });
    }
  });
  console.log("[mapDue]", Object.keys(by).length, "CSMs,",
    Object.values(by).reduce((s,c)=>s+c.overdue,0), "overdue,",
    Object.values(by).reduce((s,c)=>s+c.due,0), "total due");
  return Object.values(by);
}

function mapOnTime(rows) {
  const by = {};
  rows.forEach(r => {
    const raw = r["Cadence Member: Assigned"]||r["Assigned"]||r["CSM"]||"";
    const name = norm(raw.trim());
    if (!name||!isValidCSM(raw.trim())) return;
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

// ── PARSE SKIPPED CADENCES (Continued After 4th Reschedule) ──────────────────
// Sheet columns (from screenshot):
//   C = Outcome, E = Account, F = Skipped Touchpoints, G = Touchpoint: Owner Name
// ── PARSE MC / BC CHURN ───────────────────────────────────────────────────────
// ── PARSE DAILY CHURN ALERTS ──────────────────────────────────────────────────
function mapChurnAlerts(rows) {
  if (!rows || rows.length === 0) return [];
  return rows.map(r => ({
    date:    String(r["alert_date"]||r[0]||"").trim(),
    type:    String(r["type"]||r[1]||"").trim(),
    csm:     norm(String(r["csm_name"]||r[2]||"").trim()),
    coach:   String(r["coach"]||r[3]||"").trim(),
    account: String(r["account_name"]||r[4]||"").trim(),
    status:  String(r["status"]||r[5]||"new").trim(),
  })).filter(r => r.date && r.account && r.csm)
     .sort((a,b) => b.date.localeCompare(a.date));
}

function mapChurn(rows) {
  // Columns: CSM Coach, CSM Name, Account Name, Active Accounts, Canceled, % Churn
  const by = {};
  rows.forEach(r => {
    const csm  = String(r["CSM Name"]||r["csm_name"]||"").trim();
    const acct = String(r["Account Name"]||r["Account"]||"").trim();
    const canceled = parseFloat(r["Canceled"]||0)||0;
    const active   = parseFloat(r["Active MC Accounts Assigned Last 13 Months"]||r["Active BC Accounts Assigned Last 13 Months"]||r["Active"]||0)||0;
    if (!csm || /TOTAL|GRAND/i.test(csm)) return;
    if (!by[csm]) by[csm] = {active:0, canceled:0, accts:[]};
    by[csm].active   += active;
    by[csm].canceled += canceled;
    if (canceled > 0 && acct && !by[csm].accts.includes(acct)) by[csm].accts.push(acct);
  });
  return by;
}

// ── QA PARSER ──────────────────────────────────────────────────────────────
// Parses QA sheets — one row per CSM per month
// Handles both "2026-04" and "April" / "April 2026" month formats
function parseMonthKey(raw) {
  if (!raw) return null;
  const s = String(raw).trim();
  // Already YYYY-MM format
  if (/^\d{4}-\d{2}$/.test(s)) return s;
  // MM/DD/YYYY or M/D/YYYY (Google Sheets date format e.g. "6/1/2026")
  const slashMatch = s.match(/^(\d{1,2})\/\d{1,2}\/(\d{4})$/);
  if (slashMatch) {
    const mo = slashMatch[1].padStart(2,"0");
    return slashMatch[2] + "-" + mo;
  }
  // "April 2026" or "April"
  const months = {january:"01",february:"02",march:"03",april:"04",may:"05",june:"06",
    july:"07",august:"08",september:"09",october:"10",november:"11",december:"12"};
  const lo = s.toLowerCase();
  const mo = Object.keys(months).find(m => lo.includes(m));
  if (!mo) return null;
  const yearMatch = s.match(/\d{4}/);
  const year = yearMatch ? yearMatch[0] : "2026";
  return year + "-" + months[mo];
}

function formatMonthLabel(key) {
  if (!key) return "";
  const [y, m] = key.split("-");
  const names = ["","January","February","March","April","May","June",
    "July","August","September","October","November","December"];
  return (names[parseInt(m)]||m) + " " + y;
}

function mapQA(rows, type) {
  // type: "mc" | "ss"
  if (!rows || rows.length === 0) return {};
  console.log("[mapQA "+type+"] rows:", rows.length, "keys:", Object.keys(rows[0]));

  const MC_CRITERIA  = ["pre_kickoff_email","defined_csm_role","presented_journey_deck",
    "confirmed_priorities","next_call_agenda","provided_actionable_task","wrapup_email"];
  const SS_CRITERIA  = ["rapport_building","presented_journey_deck","agenda_open_items",
    "next_call_agenda","provided_actionable_task","wrapup_email","cer_notes"];
  const CRITERIA = type === "mc" ? MC_CRITERIA : SS_CRITERIA;

  // Also support flexible column name matching
  const MC_ALIASES = {
    "pre_kickoff_email":["pre_kickoff_email","pre-kickoff email","pre kickoff email","pre activation email","pre-activation email"],
    "defined_csm_role":["defined_csm_role","defined csm role"],
    "presented_journey_deck":["presented_journey_deck","presented journey deck"],
    "confirmed_priorities":["confirmed_priorities","confirmed priorities"],
    "next_call_agenda":["next_call_agenda","next call agenda"],
    "provided_actionable_task":["provided_actionable_task","provided actionable task"],
    "wrapup_email":["wrapup_email","wrap-up email","wrap up email","wrapup email"],
  };
  const SS_ALIASES = {
    "rapport_building":["rapport_building","rapport building"],
    "presented_journey_deck":["presented_journey_deck","presented journey deck"],
    "agenda_open_items":["agenda_open_items","agenda / open items / client tasks","agenda open items","agenda/open items"],
    "next_call_agenda":["next_call_agenda","next call agenda"],
    "provided_actionable_task":["provided_actionable_task","provided actionable task"],
    "wrapup_email":["wrapup_email","wrap-up email","wrap up email","wrapup email"],
    "cer_notes":["cer_notes","cer notes"],
  };
  const ALIASES = type === "mc" ? MC_ALIASES : SS_ALIASES;

  // Build lowercase key map for flexible matching
  const colMap = {}; // canonical -> actual csv key
  if (rows.length > 0) {
    const keys = Object.keys(rows[0]).map(k => k.toLowerCase().trim());
    Object.entries(ALIASES).forEach(([canon, variants]) => {
      const match = variants.find(v => keys.includes(v.toLowerCase()));
      if (match) {
        const actual = Object.keys(rows[0]).find(k => k.toLowerCase().trim() === match.toLowerCase());
        if (actual) colMap[canon] = actual;
      }
    });
  }

  const pf = v => {
    if (v === null || v === undefined || String(v).trim() === "" || 
        String(v).toLowerCase() === "n/a" || String(v).toLowerCase() === "na") return null;
    const s = String(v).replace(/%/g,"").trim();
    const x = parseFloat(s);
    if (isNaN(x)) return null;
    // Convert percentages > 1 to decimal
    return x > 1 ? x/100 : x;
  };

  // {month: {csm: {audits, total, criteria...}}}
  const by = {};

  // Log first row keys to debug
  if (rows.length > 0) { const r0=rows[0]; console.log("[mapQA "+type+"] ALL keys+vals:", Object.entries(r0).map(([k,v])=>JSON.stringify(k)+":"+JSON.stringify(String(v||"").slice(0,20)))); }

  rows.forEach(r => {
    const vals = Object.values(r);
    // Col A header is blank in the sheet — CSV key will be empty string or undefined
    const monthRaw = String(r["month"]||r["Month"]||r["MONTH"]||r[""]||r[" "]||vals[0]||"").trim();
    const month = parseMonthKey(monthRaw);
    if (!month) return;

    const csmRaw = String(r["csm_name"]||r["CSM Name"]||r["CSM"]||r["Name"]||
      r["Full Name"]||r["CSM name"]||vals[1]||"").trim();
    const csm = norm(csmRaw) || csmRaw;
    if (!csm || csm.length < 2 || csm === month) return;

    const audits = parseInt(r["audits"]||r["Audits"]||r["# of Audits"]||r["# of\nAudits"]||vals[2]||0)||0;
    const total  = pf(r["total_achievement"] || r["Total Achievement"] ||
      Object.entries(r).find(([k])=>k.toLowerCase().includes("total")&&k.toLowerCase().includes("achieve"))?.[1] ||
      vals[3] || 0);

    if (!by[month]) by[month] = {};
    const entry = {audits, total, criteria:{}};
    // MC positional: cols 4-10 = criteria; SS positional: cols 4-10 = criteria
    const criteriaStartCol = 4;
    CRITERIA.forEach((c, idx) => {
      const key = colMap[c];
      if (key) {
        entry.criteria[c] = pf(r[key]);
      } else {
        // Try flexible name match
        const direct = Object.keys(r).find(k => k.toLowerCase().replace(/[^a-z]/g,"").includes(c.replace(/_/g,"")));
        if (direct) {
          entry.criteria[c] = pf(r[direct]);
        } else {
          // Positional fallback
          entry.criteria[c] = pf(vals[criteriaStartCol + idx]);
        }
      }
    });
    by[month][csm] = entry;
  });

  const months = Object.keys(by).sort();
  console.log("[mapQA "+type+"] months:", months, "CSMs in latest:", months.length > 0 ? Object.keys(by[months[months.length-1]]).length : 0);
  return by;
}

// ── CALLS / BOOKINGS PARSER ────────────────────────────────────────────────
// Accepts raw Thryv bookings export: Appointment Time, Staff Name, Service Name, Status
// Groups by CSM + week + service category
function normalizeService(raw) {
  const s = (raw||"").toLowerCase();
  if (s.includes("continuation"))          return "Continuation Call";
  if (s.includes("kickoff"))               return "Kickoff Call";
  if (s.includes("set up")||s.includes("setup")||s.includes("onboarding")) return "Set Up Call";
  if (s.includes("strategy"))              return "Strategy Session";
  if (s.includes("franchise"))             return "Franchise";
  if (s.includes("implementation"))        return "Implementation";
  if (s.includes("keap"))                  return "Keap";
  if (s.includes("integrated"))            return "Integrated Onboarding";
  return "Other";
}

function weekStart(dateStr) {
  // Returns Monday of the week as YYYY-MM-DD (local time, not UTC)
  const d = new Date(dateStr);
  if (isNaN(d)) return null;
  const day = d.getDay(); // 0=Sun, local time
  const diff = (day === 0) ? -6 : 1 - day; // back to Monday
  d.setDate(d.getDate() + diff);
  // Use local date parts — avoids UTC offset shifting the day
  const yyyy = d.getFullYear();
  const mm   = String(d.getMonth() + 1).padStart(2, "0");
  const dd   = String(d.getDate()).padStart(2, "0");
  return yyyy + "-" + mm + "-" + dd;
}

// ── SHARED DATE FILTER UTILITIES (used by Calls tab + Overview tab) ────────
// Parse a YYYY-MM-DD string as local midnight, not UTC, so date-picker
// values line up exactly with stored day keys regardless of timezone.
function parseLocalDate(d) {
  if (!d) return new Date(NaN);
  const parts = String(d).split("-");
  if (parts.length === 3) return new Date(+parts[0], +parts[1]-1, +parts[2]);
  return new Date(d);
}

// Returns {from: Date|null, to: Date|null} for a named range relative to
// the latest known date in the data (so "yesterday" means relative to the
// most recent snapshot/call date, not the browser's real-world today).
function getDateRange(filterKey, latestDate, customFrom, customTo) {
  const latest = latestDate instanceof Date ? latestDate : parseLocalDate(latestDate);
  const endOfDay = d => { const x = new Date(d); x.setHours(23,59,59,999); return x; };
  if (filterKey === "yesterday") {
    const y = new Date(latest); y.setDate(y.getDate()-1);
    return { from: y, to: endOfDay(y) };
  }
  if (filterKey === "this_week") {
    const ws = parseLocalDate(weekStart(latest.toISOString ? latest : latest));
    return { from: ws, to: endOfDay(latest) };
  }
  if (filterKey === "last_week") {
    const wkAgo = new Date(latest); wkAgo.setDate(wkAgo.getDate()-7);
    return { from: wkAgo, to: endOfDay(latest) };
  }
  if (filterKey === "last_month") {
    const moAgo = new Date(latest); moAgo.setMonth(moAgo.getMonth()-1);
    return { from: moAgo, to: endOfDay(latest) };
  }
  if (filterKey === "last_quarter") {
    const qAgo = new Date(latest); qAgo.setMonth(qAgo.getMonth()-3);
    return { from: qAgo, to: endOfDay(latest) };
  }
  // Calendar quarters — Q1=Jan-Mar, Q2=Apr-Jun, Q3=Jul-Sep, Q4=Oct-Dec
  const qMatch = filterKey.match(/^(Q[1-4])(?:\s+(\d{4}))?$/i);
  if (qMatch) {
    const qMap = {Q1:[0,2], Q2:[3,5], Q3:[6,8], Q4:[9,11]};
    const [startMo, endMo] = qMap[qMatch[1].toUpperCase()] || [0,2];
    const yr = qMatch[2] ? parseInt(qMatch[2]) : latest.getFullYear();
    const from = new Date(yr, startMo, 1);
    const to   = new Date(yr, endMo + 1, 0, 23, 59, 59, 999); // last day of quarter
    return { from, to };
  }
  if (filterKey === "ytd") {
    const from = new Date(latest.getFullYear(), 0, 1);
    return { from, to: endOfDay(latest) };
  }
  if (filterKey === "custom" && customFrom && customTo) {
    return { from: parseLocalDate(customFrom), to: endOfDay(parseLocalDate(customTo)) };
  }
  return { from: null, to: null }; // "all" / unrecognized = no filtering
}

function mapCalls(rows) {
  if (!rows || rows.length === 0) return {};
  console.log("[mapCalls] rows:", rows.length, "first keys:", Object.keys(rows[0]));
  // Log ALL field values in first row to find status field
  if (rows.length > 1) {
    const r1 = rows[1];
    console.log("[mapCalls] ALL fields in row 1:", Object.entries(r1).map(([k,v])=>k+":"+JSON.stringify(String(v||"").substring(0,30))));
    const statusFields = Object.entries(r1).filter(([k,v])=>
      String(v||"").toLowerCase().includes("complet")||
      String(v||"").toLowerCase().includes("no show")||
      String(v||"").toLowerCase().includes("cancel")
    );
    console.log("[mapCalls] status-bearing fields:", statusFields.map(([k,v])=>k+"="+v));
  }

  // {csmName: {week: {service: {completed, noShow}}}}
  const by = {};

  rows.forEach(r => {
    // Support both raw export columns and manual entry columns
    const apptTime  = String(r["Appointment Time"] || r["appointment_time"] || r["week"] || "").trim();
    const staffRaw  = String(r["Staff Name"]  || r["Staff"] || r["csm_name"] || r["Assigned To"] || "").trim();
    const svcRaw    = String(r["Service Name"]|| r["Service"]|| r["service_type"] || r["Appointment Type"] || "").trim();
    const status    = String(r["Status"]      || "").trim().toLowerCase();

    if (!staffRaw || !apptTime) return;

    const csm  = norm(staffRaw) || staffRaw;
    // Store by exact day (YYYY-MM-DD) so custom date filters work per-day
    const apptDate = new Date(apptTime);
    const dayKey = !isNaN(apptDate)
      ? apptDate.getFullYear() + "-" + String(apptDate.getMonth()+1).padStart(2,"0") + "-" + String(apptDate.getDate()).padStart(2,"0")
      : apptTime.slice(0,10);
    const week = dayKey;
    const svc  = normalizeService(svcRaw) || svcRaw || "Other";

    // Scan all fields for the appointment status — could be in various columns
    // The "Status" column in Thryv exports is sometimes customer status, not appt status
    const allVals = Object.values(r).map(v=>String(v||"").toLowerCase());
    const hasCompleted = allVals.some(v=>v==="completed"||v==="complete");
    const hasNoShow    = allVals.some(v=>v==="no show"||v==="no-show"||v==="noshow");
    const hasCancelled = allVals.some(v=>v==="cancelled"||v==="canceled"||v==="cancellation");
    const apptStatus = String(r["Appointment Status"]||r["Status"]||r["status"]||"").toLowerCase().trim();
    const isCompleted  = apptStatus.includes("complet") || hasCompleted;
    const isNoShow     = (apptStatus.includes("no show")||apptStatus.includes("no-show")||apptStatus.includes("noshow")) || hasNoShow;
    const isCancelled  = (apptStatus.includes("cancel")) || hasCancelled;
    // Also support pre-aggregated format (manual entry with completed/no_show columns)
    const preComp   = parseInt(r["completed"]||0);
    const preNoShow = parseInt(r["no_show"]||r["no show"]||0);
    const preCancel = parseInt(r["cancelled"]||r["canceled"]||0);
    const isPreAgg  = !r["Appointment Time"] && (preComp > 0 || preNoShow > 0 || preCancel > 0);
    if (!isCompleted && !isNoShow && !isCancelled && !isPreAgg) return;

    if (!by[csm]) by[csm] = {};
    if (!by[csm][week]) by[csm][week] = {};
    if (!by[csm][week][svc]) by[csm][week][svc] = {completed:0, noShow:0, cancelled:0};

    if (isPreAgg) {
      by[csm][week][svc].completed  += preComp;
      by[csm][week][svc].noShow     += preNoShow;
      by[csm][week][svc].cancelled  += preCancel;
    } else {
      if (isCompleted) by[csm][week][svc].completed++;
      if (isNoShow)    by[csm][week][svc].noShow++;
      if (isCancelled) by[csm][week][svc].cancelled++;
    }
  });

  console.log("[mapCalls] parsed", Object.keys(by).length, "CSMs");
  return by;
}

function getCallWeeks(callData) {
  // Returns sorted unique day keys (YYYY-MM-DD)
  const weeks = new Set();
  Object.values(callData).forEach(csmData =>
    Object.keys(csmData).forEach(w => weeks.add(w))
  );
  return [...weeks].sort();
}

function getCallTotals(callData, csmName, week) {
  const wData = (callData[csmName]||{})[week] || {};
  const totals = {completed:0, noShow:0, cancelled:0, total:0, byService:{}};
  Object.entries(wData).forEach(([svc, d]) => {
    totals.completed += d.completed;
    totals.noShow    += d.noShow;
    totals.cancelled += (d.cancelled||0);
    const svcTotal = d.completed+d.noShow+(d.cancelled||0);
    totals.byService[svc] = {...d, cancelled:d.cancelled||0, total: svcTotal,
      rate: d.completed+d.noShow > 0 ? d.noShow/(d.completed+d.noShow) : 0};
  });
  totals.total = totals.completed + totals.noShow + totals.cancelled;
  totals.rate  = totals.completed+totals.noShow > 0 ? totals.noShow/(totals.completed+totals.noShow) : 0;
  return totals;
}

function mapBobAdj(rows) {
  // New "bob_adjustments" tab is a form-submission log. Columns include:
  // CSM Name - First Name, CSM Name - Last Name, Account Name, Enterprise ID,
  // Which assets?, MRR Amount removal or increase, Decision, Correct CSM Name - First/Last Name, etc.
  // Per team direction: apply all rows regardless of Decision status, ignore
  // "Correct CSM Name" reassignment fields for now, and treat every adjustment as a positive
  // credit added to the CSM's ending book (current MRR / retention) — regardless of the sign
  // entered on the submission form — since these corrections represent revenue that shouldn't
  // count against the CSM's retention (e.g. moved to another team), not a further loss.
  const pf = v => { const x=parseFloat(String(v||"").replace(/[$,]/g,"")); return isNaN(x)?0:x; };
  const adj = {}; // {csmName: {lcmDelta, entries:[{e,a,l,n}]}}
  rows.forEach(r => {
    const first = (r["CSM Name - First Name"]||"").trim();
    const last  = (r["CSM Name - Last Name"]||"").trim();
    const csmRaw = (first && last) ? (first+" "+last) : (r["CSM Name"]||"").trim(); // fallback to old single-column format
    const csm = norm(csmRaw)||csmRaw;
    if (!csm||csm.length<3) return;
    const amount = Math.abs(pf(r["MRR Amount removal or increase"]||0)); // always positive credit
    if (amount===0) return;
    const entry = {
      e: (r["Enterprise ID"]||r["Thryv ID"]||"").trim(),
      a: (r["Account Name"]||"").trim(),
      l: (r["Which assets?"]||"").trim(),
      n: amount,
    };
    if (!adj[csm]) adj[csm] = {lcmDelta:0, entries:[]};
    adj[csm].lcmDelta += amount;
    adj[csm].entries.push(entry);
  });
  return adj;
}

function mapBobDet(rows) {
  // Columns: A=CSM Name, B=Enterprise ID, C=Account Name, D=L2, E=Beginning of Quarter, F=End of Quarter, G=Retention%
  // Note: there is no separate "Net Billing" column — net change is computed as End - Beginning.
  const pf = v => { const x=parseFloat(String(v||0).replace(/[$,%]/g,"")); return isNaN(x)?0:x; };
  const lf = raw => {
    if (!raw) return "";
    const s = String(raw).trim();
    if (!s||/TOTAL|GRAND/i.test(s)) return "";
    if (s.includes(",")) { const [last,first]=s.split(",",2); return (first.trim()+" "+last.trim()).replace(/  +/g," ").trim(); }
    return s;
  };
  const det = {};
  // Track canceled accounts per CSM: {acctName: [product, ...]}
  const canceled = {};
  rows.forEach(r => {
    const csmRaw = String(r["CSM Name"]||"").trim();
    if (!csmRaw||/TOTAL|GRAND/i.test(csmRaw)) return;
    const csm = norm(lf(csmRaw))||lf(csmRaw);
    if (!csm||csm.length<3) return;
    const boq = pf(r["Beginning of Quarter"]||0);
    const lcm = pf(r["End of Quarter"]||r["Last Completed Month"]||0);
    const net = lcm - boq;
    const acct = String(r["Account Name"]||"").trim();
    const prod = String(r["L2"]||"").trim();
    const eid  = String(r["Enterprise ID"]||"").trim();

    // Track fully canceled lines (BOQ > 0, End of Quarter = 0)
    if (boq > 0 && lcm === 0 && acct) {
      if (!canceled[csm]) canceled[csm] = {};
      if (!canceled[csm][acct]) canceled[csm][acct] = {eid, products:[]};
      if (prod && !canceled[csm][acct].products.includes(prod))
        canceled[csm][acct].products.push(prod);
    }

    if (net===0) {
      // Still record unchanged line items in the full account list
      if (!det[csm]) det[csm] = {i:[],d:[],all:[]};
      if (!det[csm].all) det[csm].all = [];
      if (acct) det[csm].all.push({ e:eid, a:acct, l:prod, b:boq, m:lcm, n:0, status:"unchanged" });
      return;
    }
    const entry = {
      e: eid, a: acct, l: prod,
      b: boq, m: lcm, n: net,
    };
    if (!det[csm]) det[csm] = {i:[],d:[],all:[]};
    if (!det[csm].all) det[csm].all = [];
    let status;
    if (boq === 0 && lcm > 0) status = "net_new";
    else if (boq > 0 && lcm === 0) status = "cancelled";
    else if (net > 0) status = "increase";
    else status = "decrease";
    if (acct) det[csm].all.push({...entry, status});
    if (net>0) det[csm].i.push(entry);
    else det[csm].d.push(entry);
  });

  // Attach canceled summary to each CSM's det entry
  Object.entries(canceled).forEach(([csm, accts]) => {
    if (!det[csm]) det[csm] = {i:[],d:[],all:[]};
    det[csm].churned = Object.entries(accts).map(([name, v]) => ({
      name, eid: v.eid, products: v.products
    }));
  });

  return det;
}

function mapBob(rows) {
  const bob = {}, coachTotals = {}; let grand = null;
  const pf = v => { const x=parseFloat(String(v||0).replace(/[$,%]/g,"")); return isNaN(x)?0:x; };
  const lf = raw => {
    if (!raw) return "";
    const s = String(raw).trim();
    if (!s || /TOTAL|GRAND/i.test(s)) return "";
    if (s.includes(",")) { const [last,first]=s.split(",",2); return (first.trim()+" "+last.trim()).replace(/  +/g," ").trim(); }
    return s;
  };
  rows.forEach(r => {
    const csmRaw  = String(r["CSM Name"]||r["csm_name"]||r["CSM"]||"").trim();
    const coachRaw= r["CSM Coach"]||r["Coach"]||"";
    const boq=pf(r["Beginning of Quarter"]||r["Beginning Revenue"]||r["BOQ"]||r["boq"]||0);
    const lcm=pf(r["Last Completed Month"]||r["Ending Revenue"]||r["lcm"]||r["Current"]||0);
    const net=pf(r["Net Billing"]||r["net"]||0);
    const pctRaw=pf(r["Retention %"]||r["Retention"]||r["ret"]||r["pct"]||0);
    const pct=pctRaw>1?pctRaw/100:pctRaw;
    if (!csmRaw) return;
    if (/GRAND/i.test(csmRaw)) { grand={boq,lcm,net,pct}; return; }
    if (/TOTAL/i.test(csmRaw)) {
      const cn=lf(csmRaw.replace(/total/gi,"").replace(/,\s*$/,"").trim());
      if (cn) coachTotals[cn]={boq,lcm,net,pct:boq>0?lcm/boq:pct}; return;
    }
    const csm=norm(lf(csmRaw))||lf(csmRaw);
    if (!csm||csm.length<3) return;
    bob[csm]={coach:lf(String(coachRaw)),boq,lcm,net,ret:pct};
  });
  return {bob, coachTotals, grand};
}

function mapSkipped(rows) {
  const FOURTH_RESCHEDULE = "continued after 4th reschedule";
  const by = {};

  if (rows.length > 0) {
    console.log("[mapSkipped] headers:", Object.keys(rows[0]));
    console.log("[mapSkipped] total rows:", rows.length);
    console.log("[mapSkipped] first 3 rows (raw):", rows.slice(0,3));
    const outcomes = [...new Set(rows.map(r=>(r["Outcome"]||"").trim()).filter(Boolean))];
    console.log("[mapSkipped] unique outcomes:", outcomes);
  } else {
    console.log("[mapSkipped] ⚠️ No rows received");
  }

  rows.forEach(r => {
    // This tab contains ALL skipped cadences — capture all rows with a CSM name
    const raw = r["Touchpoint: Owner Name"]||r["Touchpoint: Owner Name \u2191"]||r["Owner Name"]||"";
    const name = norm(raw.trim());
    if (!name || (!isValidCSM(raw.trim()) && !isValidCSM(name))) return;

    const outcome = (r["Outcome"]||"").trim();
    const note    = (r["Notes"]||"").trim();
    const acct    = (r["Account"]||"").trim();
    const is4th   = outcome.toLowerCase().includes("4th") ||
                    outcome.toLowerCase() === FOURTH_RESCHEDULE;

    if (!by[name]) by[name] = {name, count:0, fourthCount:0, accounts:[]};
    by[name].count++;
    if (is4th) by[name].fourthCount++;

    if (acct && !by[name].accounts.find(a=>a.n===acct)) {
      by[name].accounts.push({
        n:    acct,
        note: note||"",
        outcome: outcome||"",
        is4th: is4th,
      });
    }
  });

  console.log("[mapSkipped] results:", Object.keys(by).length, "CSMs:", Object.keys(by));
  return Object.values(by);
}

// ── BUILD UNIFIED CSM LIST ─────────────────────────────────────────────────
function buildCSMs(rev, email, cad, due, ontime, skipped, bobRaw, mcChurn, bcChurn, liveBobDetArg={}, bobAdjArg={}) {
  const m = {};
  const get = name => {
    if (!m[name]) {
      const i = lk(name);
      m[name] = {name, team:(i&&i.t)||"", coach:(i&&i.c)||"", tier:(i&&i.r)||"", region:(i&&i.reg)||"",
        rev:0, mrr:0, ints:0, accts:[],
        sent:0, openRate:0, replyRate:0,
        cadCount:0, cadPct:0,
        dueCount:0, overdueCount:0, newToday:0,
        otTotal:0, otOnTime:0, otPct:null,
        skippedCount:0, skippedFourthCount:0, skippedAccts:[], liveAccounts:{}, churnedAccts:[],
        bobBoq:0, bobLcm:0, bobNet:0, bobRet:null, bobMcc:0, bobMch:[], bobBcc:0, bobBch:[]};
    }
    return m[name];
  };
  // Seed with every active ROSTER member so CSMs with no live data still appear
  Object.entries(ROSTER).forEach(([key, info]) => {
    if (info && info.c && info.t && isValidCSM(key)) {
      const canonical = norm(key) || key;
      if (canonical && !m[canonical]) get(canonical);
    }
  });
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
    c.dueCount=d.due; c.overdueCount=d.overdue; c.newToday=d.newToday; c.liveAccounts=d.accounts||{};
  });
  (ontime||[]).forEach(d => {
    const c = get(d.name);
    c.otTotal=d.total; c.otOnTime=d.onTime; c.otPct=d.pct;
  });
  (skipped||[]).forEach(d => {
    const c = get(d.name);
    c.skippedCount = d.count;
    c.skippedFourthCount = d.fourthCount||0;
    c.skippedAccts = d.accounts||[];
  });
  // Merge BOB billing data from live sheet
  if (bobRaw&&bobRaw.bob) {
    Object.entries(bobRaw.bob).forEach(([rawName, d]) => {
      const name = norm(rawName)||rawName;
      // Try exact match first, then normed name
      const c = m[name] || m[rawName];
      if (c) {
        // Apply bobAdj (same adjustment the BOB tab uses) so retention matches
        const adjKey = bobAdjArg ? Object.keys(bobAdjArg).find(k=>norm(k)===name||k===name) : null;
        const lcmDelta = adjKey ? (bobAdjArg[adjKey].lcmDelta||0) : 0;
        const adjLcm = (d.lcm||0) + lcmDelta;
        c.bobBoq=d.boq||0; c.bobLcm=adjLcm; c.bobNet=(d.net||0)+lcmDelta;
        c.bobRet=(d.boq>0&&adjLcm!=null)?adjLcm/d.boq:(d.ret||null);
      }
    });
  }
  // Merge MC churn
  if (mcChurn) {
    Object.entries(mcChurn).forEach(([rawName, d]) => {
      const name = norm(rawName)||rawName;
      const c = m[name] || m[rawName];
      if (c) { c.bobMcc=d.canceled||0; c.bobMch=d.accts||[]; }
    });
  }
  // Merge BC churn
  if (bcChurn) {
    Object.entries(bcChurn).forEach(([rawName, d]) => {
      const name = norm(rawName)||rawName;
      const c = m[name] || m[rawName];
      if (c) { c.bobBcc=d.canceled||0; c.bobBch=d.accts||[]; }
    });
  }
  // Derive churn from liveBobDet (BOB detail sheet) — account name + products
  if (liveBobDetArg && Object.keys(liveBobDetArg).length > 0) {
    Object.entries(liveBobDetArg).forEach(([rawName, det]) => {
      if (!det.churned || det.churned.length === 0) return;
      const name = norm(rawName) || rawName;
      const c = m[name] || m[rawName];
      if (!c) return;
      // Store churned accounts with products from BOB detail
      c.churnedAccts = det.churned; // [{name, eid, products:[]}]
      c.bobMcc = det.churned.length; // total unique accounts canceled
      c.bobMch = det.churned.map(a => a.name);
    });
  }

  // Also populate from hardcoded BOB_CSMS fallback ONLY if live data completely absent
  // (i.e. bobRaw fetch failed entirely — not just a stale number)
  const hasLiveBob = bobRaw && bobRaw.bob && Object.keys(bobRaw.bob).length > 0;
  if (!hasLiveBob) {
    BOB_CSMS.forEach(b => {
      const c = m[b.n];
      if (c && c.bobBoq===0 && b.boq>0) {
        c.bobBoq=b.boq; c.bobLcm=b.lcm; c.bobNet=b.net; c.bobRet=(b.boq>0&&b.lcm!=null)?b.lcm/b.boq:(b.ret||null);
        if (c.bobMcc===0) { c.bobMcc=b.mcc; c.bobMch=b.mch; }
        if (c.bobBcc===0) { c.bobBcc=b.bcc; c.bobBch=b.bch; }
      }
    });
  }
  return Object.values(m);
}

// ── PARSE HISTORY TAB ─────────────────────────────────────────────────────
function mapHistory(rows) {
  if (!rows || rows.length === 0) { console.log("[mapHistory] no rows"); return []; }
  console.log("[mapHistory] total rows:", rows.length, "first keys:", Object.keys(rows[0]));

  const pf = v => { const x = parseFloat(v); return isNaN(x) ? null : x; };
  const pf0 = v => parseFloat(v) || 0;

  // Check if first row has named headers (e.g. snapshot_date, csm_name)
  const firstKeys = Object.keys(rows[0]);
  const hasHeaders = firstKeys.some(k =>
    k === "snapshot_date" || k === "csm_name" || k === "week_label"
  );

  // Exact column mapping from Apps Script (23 cols as of v3):
  // 0:snapshot_date 1:week_label 2:csm_name 3:coach 4:team
  // 5:revenue 6:mrr 7:otr 8:non_rev_ints
  // 9:emails_sent 10:open_rate 11:reply_rate
  // 12:cadence_pct 13:cadence_total
  // 14:overdue_count 15:due_count
  // 16:ontime_pct 17:ontime_total 18:ontime_count
  // 19:skipped_count
  // 20:bob_boq 21:bob_lcm 22:bob_ret  (NEW in v3 — may be blank in older rows)
  const parseRow = (r, byName) => {
    const date = String(byName ? r["snapshot_date"] : r[0] || "").trim();
    const name = String(byName ? r["csm_name"]      : r[2] || "").trim();
    if (!date || !name) return null;
    if (date === "snapshot_date" || name === "csm_name") return null;
    if (!/^\d{4}-\d{2}-\d{2}/.test(date)) return null;
    const g = (key, idx) => byName ? r[key] : r[idx];
    return {
      date,
      week:         String(g("week_label",   1) || "").trim(),
      name:         norm(name) || name,
      coach:        String(g("coach",        3) || ""),
      team:         String(g("team",         4) || ""),
      rev:          pf0(g("revenue",        5)),
      mrr:          pf0(g("mrr",            6)),
      sent:         pf0(g("emails_sent",    9)),
      openRate:     pf(g("open_rate",       10)),
      replyRate:    pf(g("reply_rate",      11)),
      cadPct:       pf(g("cadence_pct",     12)),
      cadTotal:     pf0(g("cadence_total",  13)),
      overdueCount: pf0(g("overdue_count",  14)),
      dueCount:     pf0(g("due_count",      15)),
      otPct:        pf(g("ontime_pct",      16)),
      otTotal:      pf0(g("ontime_total",   17)),
      otOnTime:     pf0(g("ontime_count",   18)),
      skipped:      pf0(g("skipped_count",  19)),
      bobBoq:       pf(g("bob_boq",          20)),
      bobLcm:       pf(g("bob_lcm",          21)),
      bobRet:       pf(g("bob_ret",          22)),
    };
  };

  if (hasHeaders) {
    return rows.map(r => parseRow(r, true)).filter(Boolean);
  }

  // Positional fallback (no header row)
  return rows.map(r => parseRow(Object.values(r), false)).filter(Boolean);
}


// ── Q3 BOB current rollup mapper ────────────────────────────────────────────
function mapQ3Current(rows) {
  if (!rows || rows.length === 0) return {};
  const headers = rows[0] ? Object.keys(rows[0]) : [];
  const g = (row, h) => { const k = Object.keys(row).find(x=>x.toLowerCase().replace(/ /g,"_")===h.toLowerCase()); return k ? row[k] : ""; };
  const pf = v => parseFloat(String(v||"").replace(/[$,%]/g,""))||0;
  return Object.fromEntries(
    rows.filter(r=>{
      const n = g(r,"csm_name");
      if (!n) return false;
      if (n.toUpperCase().includes(" TOTAL")) return false; // skip subtotal rows
      return true;
    }).map(r => {
      const rawName = g(r,"csm_name");
      // Use the name as-is from the results tab (already canonical from Apps Script)
      // norm() lowercases which breaks map lookups — store under original casing
      const name = rawName;
      return [name, {
        name,
        boqOriginal:  pf(g(r,"boq_original")),
        boqAdjusted:  pf(g(r,"boq_adjusted")),
        currentMrr:   pf(g(r,"current_mrr")),
        netNewMrr:    pf(g(r,"net_new_mrr")),
        removedMrr:   pf(g(r,"removed_mrr")),
        cancelledMrr: pf(g(r,"cancelled_mrr")),
        retPct:       pf(g(r,"retention_pct")),
        netNewCount:  parseInt(g(r,"net_new_count"))||0,
        removedCount: parseInt(g(r,"removed_count"))||0,
        cancelledCount: parseInt(g(r,"cancelled_count"))||0,
        runDate:      g(r,"run_date"),
      }];
    })
  );
}

// ── Q3 BOB log mapper ────────────────────────────────────────────────────────
function mapQ3Log(rows) {
  if (!rows || rows.length === 0) return [];
  const g = (row, h) => { const k = Object.keys(row).find(x=>x.toLowerCase().replace(/ /g,"_")===h.toLowerCase()); return k ? row[k] : ""; };
  const pf = v => parseFloat(String(v||"").replace(/[$,%]/g,""))||0;
  return rows.filter(r=>g(r,"csm_name")).map(r => ({
    date:       g(r,"log_date"),
    csm:        norm(g(r,"csm_name")) || g(r,"csm_name"),
    acct:       g(r,"account_name"),
    eid:        g(r,"enterprise_id"),
    event:      g(r,"event_type"),
    mrrBefore:  pf(g(r,"mrr_before")),
    mrrAfter:   pf(g(r,"mrr_after")),
    mrrDelta:   pf(g(r,"mrr_delta")),
    note:       g(r,"note"),
  }));
}

// Build per-CSM weekly trend: {csmName: [{week, date, rev, openRate, cadPct, otPct, overdue}]}
function buildTrends(historyRows) {
  const by = {};
  historyRows.forEach(r => {
    if (!by[r.name]) by[r.name] = [];
    by[r.name].push(r);
  });
  // Sort each CSM's history by date ascending
  Object.keys(by).forEach(name => {
    by[name].sort((a, b) => a.date.localeCompare(b.date));
  });
  return by;
}

// Get unique weeks across all history, sorted ascending
function getWeeks(historyRows) {
  const seen = {};
  historyRows.forEach(r => { seen[r.week] = r.date; });
  return Object.entries(seen).sort((a, b) => a[1].localeCompare(b[1])).map(([w]) => w);
}

// ── TREND ARROW ────────────────────────────────────────────────────────────
function TrendArrow({curr, prev, isPercent, higherBetter=true}) {
  if (curr == null || prev == null) return null;
  const delta = curr - prev;
  if (Math.abs(delta) < 0.005) return <span style={{color:"#808080",fontSize:10,marginLeft:3}}>→</span>;
  const up = delta > 0;
  const good = higherBetter ? up : !up;
  const col = good ? "#16a34a" : "#dc2626";
  const arrow = up ? "↑" : "↓";
  const label = isPercent ? Math.round(Math.abs(delta)*100)+"pp" : Math.round(Math.abs(delta));
  return <span style={{color:col,fontSize:10,marginLeft:3,fontWeight:500}}>{arrow}{label}</span>;
}

// ── SPARKLINE ──────────────────────────────────────────────────────────────
function Sparkline({values, color="#FF5000", height=24, width=80}) {
  const valid = values.filter(v => v != null);
  if (valid.length < 2) return <span style={{color:"#ccc",fontSize:10}}>—</span>;
  const min = Math.min(...valid), max = Math.max(...valid);
  const range = max - min || 1;
  const pts = values.map((v, i) => {
    const x = (i / (values.length - 1)) * width;
    const y = v != null ? height - ((v - min) / range) * (height - 4) - 2 : null;
    return {x, y};
  }).filter(p => p.y != null);
  if (pts.length < 2) return null;
  const d = pts.map((p, i) => (i === 0 ? "M" : "L") + p.x.toFixed(1) + "," + p.y.toFixed(1)).join(" ");
  return (
    <svg width={width} height={height} style={{display:"inline-block",verticalAlign:"middle"}}>
      <path d={d} fill="none" stroke={color} strokeWidth={1.5} strokeLinecap="round" strokeLinejoin="round"/>
      <circle cx={pts[pts.length-1].x} cy={pts[pts.length-1].y} r={2.5} fill={color}/>
    </svg>
  );
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
  const [expanded, setExpanded] = useState(false);
  const PREVIEW = 7;
  const team = csms.filter(c => (lk(c.name)&&lk(c.name).c===coach.e)||c.coach===coach.e);
  const cadC = team.filter(c=>c.cadCount>0);
  const emC  = team.filter(c=>c.sent>0);
  const otC  = team.filter(c=>c.otTotal>=3);
  const avgCad  = cadC.length ? cadC.reduce((s,c)=>s+c.cadPct,0)/cadC.length : null;
  const avgOpen = emC.length  ? emC.reduce((s,c)=>s+c.openRate,0)/emC.length : null;
  const avgOT   = otC.length  ? otC.reduce((s,c)=>s+c.otPct,0)/otC.length : null;
  const revPct  = team.length ? team.filter(c=>c.rev>0||c.ints>0).length/team.length : null;
  const teamRev = team.reduce((s,c)=>s+c.rev,0);
  const bobC  = team.filter(c=>c.bobRet!=null);
  const avgBob = bobC.length ? bobC.reduce((s,c)=>s+c.bobRet,0)/bobC.length : null;
  let sc=0,sf=0;
  if(cadC.length>0&&avgCad!=null){sc+=Math.min(avgCad/0.9,1)*25;sf++;}
  if(avgOpen!=null){sc+=Math.min(avgOpen/0.7,1)*20;sf++;}
  if(avgOT!=null){sc+=Math.min(avgOT/0.8,1)*15;sf++;}
  if(revPct!=null){sc+=Math.min(revPct/0.7,1)*20;sf++;}
  if(avgBob!=null){sc+=Math.min(avgBob/0.91,1)*20;sf++;}
  // Skipped cadence penalty: -5 pts per CSM with skips, capped at -25
  const skippedOnTeam = team.filter(c=>c.skippedCount>0).length;
  if(sf>0 && skippedOnTeam>0) sc = Math.max(0, sc - Math.min(skippedOnTeam*5, 25));
  const score = sf ? Math.round(Math.min(sc,100)) : null;
  const cls = score==null?"none":score>=75?"win":score>=50?"warn":"att";
  const sCol = {win:"#16a34a",warn:"#d97706",att:"#dc2626",none:"#888"}[cls];
  const topCol = {win:"#16a34a",warn:"#d97706",att:"#dc2626",none:"#e5e7eb"}[cls];
  const wins=cadC.filter(c=>c.cadPct>=0.9).length;
  const warns=cadC.filter(c=>c.cadPct>=0.5&&c.cadPct<0.9).length;
  const atts=cadC.filter(c=>c.cadPct>0&&c.cadPct<0.9).length;

  // Sort: on-time % desc (if available), then cadence % desc, then alphabetical
  const sorted = [...team].sort((a,b) => {
    const aOT = a.otTotal>=3 ? a.otPct : null;
    const bOT = b.otTotal>=3 ? b.otPct : null;
    if (aOT!==null && bOT!==null && aOT!==bOT) return bOT-aOT;
    if (aOT!==null && bOT===null) return -1;
    if (aOT===null && bOT!==null) return 1;
    if ((b.cadPct||0) !== (a.cadPct||0)) return (b.cadPct||0)-(a.cadPct||0);
    return a.name.localeCompare(b.name);
  });

  const visible = expanded ? sorted : sorted.slice(0, PREVIEW);
  const hiddenCount = sorted.length - PREVIEW;

  return (
    <div style={{...S.card,position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:0,left:0,right:0,height:3,background:topCol}}/>
      <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:8}}>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:13,fontWeight:500,color:"#29355D",marginTop:4,cursor:"pointer"}} onClick={()=>onSelectCoach(coach.e)}>
            {coach.n}
          </div>
          <div style={{fontSize:11,fontWeight:500,color:coach.col,marginBottom:6}}>{coach.t}</div>
          <div style={{fontSize:11,color:"#808080",marginBottom:8}}>Revenue: <strong style={{color:"#FF5000"}}>{fk(teamRev)}</strong></div>
        </div>
        {TEAM_LOGOS[coach.t]&&<img src={TEAM_LOGOS[coach.t]} alt={coach.t}
          style={{width:72,height:72,objectFit:"contain",borderRadius:8,flexShrink:0,marginTop:2,
            background:coach.t==="White Wave Warriors"?"#1a1a1a":coach.t==="Team Status Engaged"||coach.t==="The Dominican Dream Team"?"#fff":"transparent",
            padding:coach.t==="White Wave Warriors"||coach.t==="Team Status Engaged"||coach.t==="The Dominican Dream Team"?4:0}}/>}
      </div>
      <div style={{display:"flex",alignItems:"baseline",gap:6,marginBottom:8}}>
        <span style={{fontSize:40,fontWeight:500,lineHeight:1,color:sCol}}>{score!=null?score:"--"}</span>
        <div style={{fontSize:11,color:"#808080",lineHeight:1.4}}>/ 100<br/><span style={{fontSize:10}}>{wins}✓ {warns}⚠ {atts}✗</span></div>
      </div>
      <Bar label="Cadence"   val={avgCad}  hi={0.9} lo={0.5}/>
      <Bar label="On-time"   val={avgOT}   hi={0.8} lo={0.6}/>
      <Bar label="Email"     val={avgOpen} hi={0.7} lo={0.35}/>
      <Bar label="Revenue"   val={revPct}  hi={0.7} lo={0.4}/>
      <Bar label="Retention" val={avgBob}  hi={0.91} lo={0.85}/>
      <div style={{height:.5,background:"rgba(41,53,93,.07)",margin:"10px 0"}}/>
      <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:6}}>CSM Status</div>
      {visible.map(c => {
        const hc = c.cadCount>0;
        const bdgTxt = hc?(c.cadPct>=0.9?"Win":c.cadPct>=0.5?"Watch":"Coach"):"No tasks";
        const bdgCol = hc?(c.cadPct>=0.9?"rgba(22,163,74,.12)":c.cadPct>=0.5?"rgba(217,119,6,.12)":"rgba(220,38,38,.12)"):"rgba(128,128,128,.1)";
        const bdgFg  = hc?(c.cadPct>=0.9?"#166534":c.cadPct>=0.5?"#854d0e":"#991b1b"):"#808080";
        return (
          <div key={c.name} style={{display:"flex",alignItems:"center",gap:6,padding:"3px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",cursor:"pointer"}} onClick={()=>onSelectCSM(c.name)}>
            <span style={{flex:1,fontSize:11,fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"#121212"}}>{dispName(c.name)}</span>
            <span style={{fontSize:11,fontWeight:500,width:28,textAlign:"right",flexShrink:0,color:hc?pc(c.cadPct):"#888"}}>{hc?Math.round(c.cadPct*100)+"%":"--"}</span>
            <span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:bdgCol,color:bdgFg,flexShrink:0}}>{bdgTxt}</span>
            {c.overdueCount>0&&<span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:"rgba(220,38,38,.1)",color:"#991b1b",flexShrink:0}}>{c.overdueCount}!</span>}
            {c.skippedCount>0&&<span style={{fontSize:10,fontWeight:700,padding:"1px 7px",borderRadius:20,background:"rgba(139,0,0,.15)",color:"#7f1d1d",flexShrink:0}} title="Continued After 4th Reschedule">🚩{c.skippedCount}</span>}
            {c.otTotal>=3&&<span style={{fontSize:10,color:pc(c.otPct),marginLeft:2,flexShrink:0}}>{Math.round(c.otPct*100)}%⏱</span>}
          </div>
        );
      })}
      {sorted.length > PREVIEW && (
        <button
          onClick={()=>setExpanded(e=>!e)}
          style={{width:"100%",marginTop:8,padding:"5px 0",fontSize:11,fontWeight:500,color:"#FF5000",background:"rgba(255,80,0,.06)",border:"0.5px solid rgba(255,80,0,.2)",borderRadius:6,cursor:"pointer",textAlign:"center"}}>
          {expanded ? "▲ Show less" : `▼ +${hiddenCount} more`}
        </button>
      )}
    </div>
  );
}

// ── CSM DETAIL VIEW ────────────────────────────────────────────────────────
function CSMDetail({csm: csmRaw, onClear, bobRaw, mcChurn, bcChurn, liveBobDet={}, bobAdj={}}) {
  const getDet = n => {
    const base = liveBobDet[n]||liveBobDet[norm(n)]||BOB_DETAIL[n]||BOB_DETAIL[norm(n)]||{};
    const adjKey = Object.keys(bobAdj).find(k=>norm(k)===n||k===n);
    if (!adjKey) return base;
    const adj = bobAdj[adjKey];
    const extraInc = adj.entries.filter(e=>e.n>0).map(e=>({...e, b:0, m:e.n, _adj:true}));
    const extraDec = adj.entries.filter(e=>e.n<0).map(e=>({...e, b:0, m:e.n, _adj:true}));
    return {...base, i:[...(base.i||[]),...extraInc], d:[...(base.d||[]),...extraDec]};
  };
  // Merge live BOB data over whatever's in csm object (live sheet > hardcoded)
  const liveBob = (bobRaw&&bobRaw.bob) ? Object.entries(bobRaw.bob).find(([k])=>norm(k)===csmRaw.name||k===csmRaw.name) : null;
  const liveMc  = mcChurn  ? Object.entries(mcChurn).find(([k])=>norm(k)===csmRaw.name||k===csmRaw.name)  : null;
  const liveBc  = bcChurn  ? Object.entries(bcChurn).find(([k])=>norm(k)===csmRaw.name||k===csmRaw.name)  : null;
  const csm = {
    ...csmRaw,
    bobBoq: liveBob ? liveBob[1].boq||0 : csmRaw.bobBoq,
    bobLcm: liveBob ? liveBob[1].lcm||0 : csmRaw.bobLcm,
    bobNet: liveBob ? liveBob[1].net||0 : csmRaw.bobNet,
    bobRet: liveBob ? (liveBob[1].boq>0&&liveBob[1].lcm!=null?liveBob[1].lcm/liveBob[1].boq:liveBob[1].ret) : csmRaw.bobRet,
    bobMcc: liveMc  ? liveMc[1].canceled||0  : csmRaw.bobMcc,
    bobMch: liveMc  ? liveMc[1].accts||[]    : csmRaw.bobMch,
    bobBcc: liveBc  ? liveBc[1].canceled||0  : csmRaw.bobBcc,
    bobBch: liveBc  ? liveBc[1].accts||[]    : csmRaw.bobBch,
    churnedAccts: csmRaw.churnedAccts||[],
  };
  const [cadTab, setCadTab] = useState("due"); // "due" | "ontime"
  const i = lk(csm.name)||{};
  const coach = COACHES.find(c=>c.e===(i.c||csm.coach));
  const ot = csm.otTotal>=1 ? csm : null;
  const totalAcctRev = csm.accts.reduce((s,a)=>s+a.m+a.o,0);
  // Due/overdue comes ONLY from live sheet (gid=341836664) — no hardcoded fallback
  const liveAccts = csm.liveAccounts || {};
  const dueAccts = Object.entries(liveAccts)
    .map(([acctName, tasks]) => ({n: acctName, d: tasks}))
    .filter(a => a.d.length > 0)
    .sort((a,b) => {
      const aOv = a.d.some(t=>t.ov) ? 1 : 0;
      const bOv = b.d.some(t=>t.ov) ? 1 : 0;
      return bOv - aOv || a.n.localeCompare(b.n);
    });

  // On-time history still comes from hardcoded CAD_ACCTS (separate data source)
  const cadAccts = CAD_ACCTS[csm.name] || [];
  const otAccts  = cadAccts.filter(a=>a.ott>0);
  const totalDueTasks = dueAccts.reduce((s,a)=>s+a.d.length,0);
  const overdueAccts  = dueAccts.filter(a=>a.d.some(t=>t.ov));

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
  // BOB retention insights
  if (csm.bobRet!=null) {
    if (csm.bobRet>=0.91) wins.push("Book of business retention "+pp(csm.bobRet)+" — above the 91% goal");
    else if (csm.bobRet>=0.88) atts.push("Retention "+pp(csm.bobRet)+" is near but below the 91% goal — focus on at-risk accounts");
    else atts.push("Retention "+pp(csm.bobRet)+" is significantly below the 91% goal — prioritize retention conversations");
  }
  if (csm.churnedAccts&&csm.churnedAccts.length>0) atts.push("Churn: "+csm.churnedAccts.length+" account"+(csm.churnedAccts.length>1?"s":"")+" at $0 billing ("+csm.churnedAccts.slice(0,3).map(a=>a.name+(a.products&&a.products.length?" ["+a.products.join(", ")+"]":"")).join(", ")+")");
  else if (csm.bobMcc>0) atts.push("MC churn: "+csm.bobMcc+" account"+(csm.bobMcc>1?"s":"")+" canceled this quarter");
  if (csm.bobNet<0&&csm.bobBoq>0) {
    const lostPct = Math.abs(csm.bobNet)/csm.bobBoq;
    if (lostPct>0.1) atts.push("Net billing down "+pp(lostPct)+" vs start of quarter — "+fd(Math.abs(csm.bobNet))+" lost");
  }
  if(csm.skippedCount>0) {
    if (csm.skippedFourthCount>0) {
      atts.unshift("🚩 "+csm.skippedFourthCount+" account"+(csm.skippedFourthCount>1?"s":"")+" Continued After 4th Reschedule — requires immediate coaching conversation");
    }
    const regularSkips = csm.skippedCount - csm.skippedFourthCount;
    if (regularSkips>0) atts.push(regularSkips+" skipped cadence task"+(regularSkips>1?"s":"")+" yesterday — review notes with CSM");
  }
  else if(csm.overdueCount>0) atts.push(csm.overdueCount+" overdue tasks across "+overdueAccts.length+" accounts");
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

  const tdS = {padding:"7px 8px 7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",fontSize:12,verticalAlign:"top"};
  const thS = {fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,padding:"0 8px 8px 0",textAlign:"left",borderBottom:"0.5px solid rgba(41,53,93,.08)"};

  return (
    <div>
      {/* Header card */}
      <div style={{...S.card,marginBottom:16}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:4}}>
          <div style={{fontSize:20,fontWeight:500,color:"#29355D"}}>{csm.name}</div>
          {onClear&&<button onClick={onClear} style={{fontSize:11,color:"#FF5000",background:"none",border:"0.5px solid #FF5000",borderRadius:20,padding:"4px 12px",cursor:"pointer"}}>✕ Clear filter</button>}
        </div>
        <div style={{fontSize:12,color:"#808080",marginBottom:16}}>
          {i.t||csm.team||""}{i.r?" · "+i.r:""}{coach?" · Coach: "+coach.n:""}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,minmax(0,1fr))",gap:12}}>
          {statBox("Revenue",       csm.rev>0?fd(csm.rev):"--",       csm.mrr>0?"MRR "+fd(csm.mrr):null,                                          csm.rev>0?"#FF5000":null)}
          {statBox("Email open",    csm.sent>0?pp(csm.openRate):"--",  csm.sent>0?csm.sent+" sent · "+pp(csm.replyRate)+" reply":"No emails",       csm.sent>0?pc(csm.openRate):null)}
          {statBox("Cadence",       csm.cadCount>0?pp(csm.cadPct):csm.overdueCount>0?csm.overdueCount+" overdue":"No tasks", csm.dueCount>0?csm.dueCount+" due · "+csm.newToday+" new":"Nothing due yesterday", csm.cadCount>0?pc(csm.cadPct):csm.overdueCount>0?"#dc2626":null)}
          {statBox("On-time %",     ot&&ot.otTotal>=1?pp(ot.otPct):"--", ot&&ot.otTotal>=1?ot.otOnTime+"/"+ot.otTotal+" on time":"No data",         ot&&ot.otTotal>=1?pc(ot.otPct):null)}
          {statBox("Retention",     csm.bobRet!=null?pp(csm.bobRet):"--", csm.bobBoq>0?"BOQ "+fd(csm.bobBoq):"No BOB data",                        csm.bobRet!=null?(csm.bobRet>=0.91?"#16a34a":csm.bobRet>=0.85?"#d97706":"#dc2626"):null)}
          {statBox("Integrations",  csm.accts.length||"--",             csm.accts.length>0?fd(totalAcctRev)+" total":null,                          null)}
        </div>
      </div>

      {/* Insights */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
        <div style={{background:"rgba(22,163,74,.06)",border:"0.5px solid rgba(22,163,74,.2)",borderRadius:12,padding:16}}>
          <div style={{fontSize:12,fontWeight:500,color:"#166534",marginBottom:10}}>✓ What's working</div>
          {wins.length===0
            ? <div style={{fontSize:12,color:"#808080",fontStyle:"italic"}}>No clear wins yet this period</div>
            : wins.map((w,idx)=><div key={idx} style={{fontSize:12,padding:"5px 0",borderBottom:"0.5px solid rgba(0,0,0,.06)",display:"flex",gap:6}}><span style={{color:"#16a34a",flexShrink:0}}>✓</span>{w}</div>)}
        </div>
        <div style={{background:"rgba(220,38,38,.05)",border:"0.5px solid rgba(220,38,38,.15)",borderRadius:12,padding:16}}>
          <div style={{fontSize:12,fontWeight:500,color:"#991b1b",marginBottom:10}}>⚠ Coaching focus</div>
          {atts.length===0
            ? <div style={{fontSize:12,color:"#808080",fontStyle:"italic"}}>No urgent issues identified</div>
            : atts.map((a,idx)=><div key={idx} style={{fontSize:12,padding:"5px 0",borderBottom:"0.5px solid rgba(0,0,0,.06)",display:"flex",gap:6}}><span style={{color:"#dc2626",flexShrink:0}}>⚠</span>{a}</div>)}
        </div>
      </div>

      {/* Book of business */}
      {csm.bobBoq>0&&<div style={{...S.card,marginBottom:16}}>
        <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>Book of business — this quarter</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,minmax(0,1fr))",gap:12,marginBottom:csm.bobMcc+csm.bobBcc>0?16:0}}>
          {[
            {l:"Beginning of quarter",v:fd(csm.bobBoq),col:"#29355D"},
            {l:"Current billing",v:fd(csm.bobLcm),col:"#5378FC"},
            {l:"Net change",v:(csm.bobNet<0?"-":"+")+fd(Math.abs(csm.bobNet)),col:csm.bobNet<0?"#dc2626":"#16a34a"},
            {l:"Retention rate",v:pp(csm.bobRet||0),col:csm.bobRet>=0.91?"#16a34a":csm.bobRet>=0.85?"#d97706":"#dc2626"},
          ].map(s=>(
            <div key={s.l} style={{background:"#F4F6FB",borderRadius:8,padding:12}}>
              <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:4}}>{s.l}</div>
              <div style={{fontSize:18,fontWeight:500,color:s.col}}>{s.v}</div>
            </div>
          ))}
        </div>
        {/* Churned accounts from BOB detail (preferred) or MC/BC fallback */}
        {(csm.churnedAccts&&csm.churnedAccts.length>0)
          ? <div>
              <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:10}}>
                Churned accounts this quarter ({csm.churnedAccts.length})
              </div>
              {csm.churnedAccts.map((a,i)=>(
                <div key={i} style={{padding:"8px 0",borderBottom:"0.5px solid rgba(41,53,93,.06)",display:"flex",alignItems:"center",justifyContent:"space-between",gap:8}}>
                  <div style={{fontSize:12,fontWeight:500,color:"#29355D"}}>{a.name}</div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:4,justifyContent:"flex-end"}}>
                    {(a.products||[]).map((p,j)=>(
                      <span key={j} style={{fontSize:10,padding:"2px 8px",borderRadius:20,background:"rgba(220,38,38,.08)",color:"#991b1b",fontWeight:500,whiteSpace:"nowrap"}}>{p}</span>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          : (csm.bobMcc>0||csm.bobBcc>0)&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
              <div>
                <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:8}}>MC churned ({csm.bobMcc})</div>
                {csm.bobMch.map((a,i)=><div key={i} style={{padding:"4px 0",borderBottom:"0.5px solid rgba(41,53,93,.06)",fontSize:12}}>{a}</div>)}
              </div>
              <div>
                <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:8}}>BC churned ({csm.bobBcc})</div>
                {csm.bobBch.map((a,i)=><div key={i} style={{padding:"4px 0",borderBottom:"0.5px solid rgba(41,53,93,.06)",fontSize:12,color:"#991b1b"}}>{a}</div>)}
              </div>
            </div>}
      </div>}

      {/* Skipped cadences detail */}
      {csm.skippedCount>0&&<div style={{...S.card,marginBottom:16}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12}}>
          <span style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500}}>Skipped cadences — prior day</span>
          <span style={{fontSize:10,fontWeight:600,padding:"2px 8px",borderRadius:20,background:"rgba(127,29,29,.1)",color:"#7f1d1d"}}>{csm.skippedCount} skipped</span>
          {csm.skippedFourthCount>0&&<span style={{fontSize:10,fontWeight:600,padding:"2px 8px",borderRadius:20,background:"rgba(220,38,38,.15)",color:"#991b1b"}}>🚩 {csm.skippedFourthCount} × 4th reschedule</span>}
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:6}}>
          {(csm.skippedAccts||[]).map((a,i)=>(
            <div key={i} style={{borderRadius:8,padding:"8px 12px",
              background:a.is4th?"rgba(127,29,29,.04)":"rgba(217,119,6,.04)",
              border:`0.5px solid ${a.is4th?"rgba(127,29,29,.18)":"rgba(217,119,6,.22)"}`,
              borderLeft:`3px solid ${a.is4th?"#7f1d1d":"#d97706"}`}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:a.note?4:0}}>
                <span style={{fontSize:12,fontWeight:500,color:a.is4th?"#7f1d1d":"#92400e"}}>
                  {a.is4th&&"🚩 "}{a.n}
                </span>
                {a.outcome&&<span style={{fontSize:10,padding:"1px 7px",borderRadius:20,
                  background:a.is4th?"rgba(127,29,29,.1)":"rgba(217,119,6,.1)",
                  color:a.is4th?"#7f1d1d":"#92400e"}}>{a.outcome}</span>}
              </div>
              {a.note&&<div style={{fontSize:11,color:"#808080",fontStyle:"italic"}}>"{a.note}"</div>}
            </div>
          ))}
        </div>
      </div>}

      {/* BOB billing detail — increases / decreases */}
      {(()=>{
        const det = getDet(csm.name)||{};
        const inc = det.i||[], dec = det.d||[];
        if (inc.length===0 && dec.length===0) return null;
        const tdS2={padding:"7px 8px 7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",fontSize:12,verticalAlign:"top"};
        const thS2={fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,padding:"0 8px 8px 0",textAlign:"left",borderBottom:"0.5px solid rgba(41,53,93,.08)"};
        return <div style={{...S.card,marginBottom:16}}>
          <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>
            Billing changes this quarter
            {inc.length>0&&<span style={{marginLeft:8,fontSize:10,fontWeight:500,padding:"2px 8px",borderRadius:20,background:"rgba(22,163,74,.1)",color:"#166534"}}>↑ {inc.length} increase{inc.length!==1?"s":""}</span>}
            {dec.length>0&&<span style={{marginLeft:6,fontSize:10,fontWeight:500,padding:"2px 8px",borderRadius:20,background:"rgba(220,38,38,.1)",color:"#991b1b"}}>↓ {dec.length} decrease{dec.length!==1?"s":""}</span>}
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr>
              <th style={thS2}>Enterprise ID</th>
              <th style={thS2}>Product</th>
              <th style={{...thS2,textAlign:"right"}}>BOQ</th>
              <th style={{...thS2,textAlign:"right"}}>Current</th>
              <th style={{...thS2,textAlign:"right"}}>Change</th>
              <th style={{...thS2}}>Type</th>
            </tr></thead>
            <tbody>
              {[...inc.map(r=>({...r,_t:"increase"})),...dec.map(r=>({...r,_t:"decrease"}))].sort((a,b)=>b.n-a.n).map((r,i)=>{
                const isInc = r._t==="increase";
                return <tr key={i}>
                  <td style={{...tdS2,fontFamily:"monospace",fontSize:11,color:"#808080"}}>{r.e}</td>
                  <td style={tdS2}>{r.l}</td>
                  <td style={{...tdS2,textAlign:"right",color:"#808080"}}>${r.b.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2})}</td>
                  <td style={{...tdS2,textAlign:"right"}}>${r.m.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2})}</td>
                  <td style={{...tdS2,textAlign:"right",fontWeight:600,color:isInc?"#16a34a":"#dc2626"}}>
                    {isInc?"+":""}{r.n.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2})}
                  </td>
                  <td style={tdS2}>
                    <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
                      <span style={{fontSize:10,fontWeight:500,padding:"2px 8px",borderRadius:20,
                        background:isInc?"rgba(22,163,74,.1)":"rgba(220,38,38,.1)",
                        color:isInc?"#166534":"#991b1b"}}>
                        {isInc?"↑ Increase":"↓ Decrease"}
                      </span>
                      {r._adj&&<span style={{fontSize:10,fontWeight:600,padding:"2px 7px",borderRadius:20,background:"rgba(83,120,252,.1)",color:"#3b5bdb"}}>✦ Adj</span>}
                    </div>
                  </td>
                </tr>;
              })}
            </tbody>
          </table>
        </div>;
      })()}

      {/* Cadence accounts */}
      <div style={{...S.card,marginBottom:16}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
          <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500}}>
            Cadence accounts
            {dueAccts.length>0&&<span style={{marginLeft:8,fontSize:10,fontWeight:500,padding:"2px 8px",borderRadius:20,background:"rgba(220,38,38,.1)",color:"#991b1b"}}>{csm.overdueCount} overdue</span>}
          </div>
          <div style={{display:"flex",gap:4,background:"#ECEEF1",borderRadius:8,padding:3}}>
            {[["due","Due / Overdue"],["ontime","On-time history"]].map(([k,l])=>(
              <button key={k} onClick={()=>setCadTab(k)}
                style={{padding:"4px 12px",fontSize:11,fontWeight:500,borderRadius:6,border:"none",background:cadTab===k?"#fff":"transparent",color:cadTab===k?"#29355D":"#808080",cursor:"pointer",boxShadow:cadTab===k?"0 1px 3px rgba(0,0,0,.08)":"none"}}>
                {l}
              </button>
            ))}
          </div>
        </div>

        {cadTab==="due" ? (
          dueAccts.length===0
            ? <div style={{color:"#808080",fontSize:12,padding:"8px 0"}}>No due or overdue tasks</div>
            : <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr>
                  <th style={{...thS,width:"35%"}}>Account</th>
                  <th style={thS}>Task</th>
                  <th style={{...thS,textAlign:"right"}}>Due date</th>
                  <th style={{...thS,textAlign:"right"}}>Status</th>
                </tr></thead>
                <tbody>
                  {dueAccts
                    .sort((a,b)=>{
                      const aOv=a.d.some(t=>t.ov)?1:0, bOv=b.d.some(t=>t.ov)?1:0;
                      return bOv-aOv || a.n.localeCompare(b.n);
                    })
                    .map(acct=>
                      acct.d.map((task,ti)=>(
                        <tr key={acct.n+ti}>
                          <td style={{...tdS,fontWeight:ti===0?500:400,color:ti===0?"#29355D":"#808080"}}>{ti===0?acct.n:""}</td>
                          <td style={tdS}>{task.t}</td>
                          <td style={{...tdS,textAlign:"right",color:"#808080"}}>{task.due}</td>
                          <td style={{...tdS,textAlign:"right"}}>
                            {task.ov
                              ? <span style={{fontSize:10,fontWeight:500,padding:"2px 8px",borderRadius:20,background:"rgba(220,38,38,.1)",color:"#991b1b"}}>Overdue</span>
                              : task.nw
                                ? <span style={{fontSize:10,fontWeight:500,padding:"2px 8px",borderRadius:20,background:"rgba(83,120,252,.1)",color:"#1e3a8a"}}>New today</span>
                                : <span style={{fontSize:10,fontWeight:500,padding:"2px 8px",borderRadius:20,background:"rgba(217,119,6,.1)",color:"#92400e"}}>Due</span>}
                          </td>
                        </tr>
                      ))
                    )}
                </tbody>
              </table>
        ) : (
          otAccts.length===0
            ? <div style={{color:"#808080",fontSize:12,padding:"8px 0"}}>No on-time history for this period</div>
            : <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr>
                  <th style={{...thS,width:"45%"}}>Account</th>
                  <th style={{...thS,textAlign:"right"}}>Tasks</th>
                  <th style={{...thS,textAlign:"right"}}>On-time</th>
                  <th style={{...thS,textAlign:"right"}}>Rate</th>
                </tr></thead>
                <tbody>
                  {otAccts
                    .sort((a,b)=>{
                      const ap=a.ott>0?a.oto/a.ott:0, bp=b.ott>0?b.oto/b.ott:0;
                      return bp-ap || a.n.localeCompare(b.n);
                    })
                    .map(acct=>{
                      const pct = acct.ott>0?acct.oto/acct.ott:0;
                      return <tr key={acct.n}>
                        <td style={{...tdS,fontWeight:500}}>{acct.n}</td>
                        <td style={{...tdS,textAlign:"right",color:"#808080"}}>{acct.ott}</td>
                        <td style={{...tdS,textAlign:"right"}}>{acct.oto}</td>
                        <td style={{...tdS,textAlign:"right",fontWeight:500,color:bc(pct,0.8,0.5)}}>
                          {Math.round(pct*100)}%
                        </td>
                      </tr>;
                    })}
                </tbody>
              </table>
        )}
      </div>

      {/* Revenue integrations */}
      <div style={S.card}>
        <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>
          Revenue integrations — {csm.accts.length} submissions · {fd(totalAcctRev)} total
        </div>
        {csm.accts.length===0
          ? <div style={{color:"#808080",fontSize:12}}>No submissions this period</div>
          : <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,tableLayout:"fixed"}}>
              <colgroup><col style={{width:"38%"}}/><col style={{width:"10%"}}/><col style={{width:"11%"}}/><col style={{width:"11%"}}/><col style={{width:"30%"}}/></colgroup>
              <thead><tr>{["Business","Type","MRR","OTR","Integration"].map(h=><th key={h} style={{...thS,textAlign:h==="MRR"||h==="OTR"?"right":"left"}}>{h}</th>)}</tr></thead>
              <tbody>{csm.accts.map((a,idx)=>{
                const cls=a.t==="Monthly Recurring Revenue"?"mrr":a.t==="One-Time Revenue"?"otr":"non";
                const sh=a.t==="Monthly Recurring Revenue"?"MRR":a.t==="One-Time Revenue"?"OTR":"Non-Rev";
                const bg=cls==="mrr"?"rgba(22,163,74,.1)":cls==="otr"?"rgba(83,120,252,.1)":"rgba(128,128,128,.1)";
                const fg=cls==="mrr"?"#166534":cls==="otr"?"#1e3a8a":"#808080";
                return <tr key={idx}>
                  <td style={{...tdS,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.b}</td>
                  <td style={tdS}><span style={{fontSize:10,fontWeight:500,padding:"2px 7px",borderRadius:20,background:bg,color:fg}}>{sh}</span></td>
                  <td style={{...tdS,textAlign:"right"}}>{a.m>0?fd(a.m):"--"}</td>
                  <td style={{...tdS,textAlign:"right"}}>{a.o>0?fd(a.o):"--"}</td>
                  <td style={{...tdS,color:"#808080",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.n||"--"}</td>
                </tr>;
              })}</tbody>
            </table>}
      </div>
    </div>
  );
}

// ── COACHING TAB ───────────────────────────────────────────────────────────
function CoachingView({csms, coach, onSelectCSM, onSelectCoach, onClear, skippedCSMs, bobRaw, mcChurn, bcChurn, liveBobDet={}, isCsmView=false, bobAdj={}}) {
  if (onSelectCSM._selected) {
    const c = csms.find(x=>x.name===onSelectCSM._selected)||csms[0];
    return c ? <CSMDetail csm={c} onClear={isCsmView?null:onClear} bobRaw={bobRaw} mcChurn={mcChurn} bcChurn={bcChurn} liveBobDet={liveBobDet} bobAdj={bobAdj}/> : null;
  }
  const coaches = coach ? COACHES.filter(c=>c.e===coach) : COACHES;
  const cols = coaches.length===1?1:coaches.length===2?2:3;
  const overdue = csms.filter(c=>c.overdueCount>0).sort((a,b)=>b.overdueCount-a.overdueCount).slice(0,6);
  const skipped = (skippedCSMs||[]);
  const attn = csms.filter(c=>(c.cadCount>0&&c.cadPct<0.9)||c.skippedCount>0||(c.bobRet!=null&&c.bobRet<0.91)).sort((a,b)=>{
    if(b.skippedCount>0&&!a.skippedCount) return 1;
    if(a.skippedCount>0&&!b.skippedCount) return -1;
    if(a.bobRet!=null&&a.bobRet<0.91&&(b.bobRet==null||b.bobRet>=0.91)) return -1;
    if(b.bobRet!=null&&b.bobRet<0.91&&(a.bobRet==null||a.bobRet>=0.91)) return 1;
    return a.cadPct-b.cadPct;
  });
  const wins = csms.filter(c=>c.cadCount>0&&c.cadPct>=0.9&&c.skippedCount===0).sort((a,b)=>b.cadPct-a.cadPct);
  return (
    <div>
      {/* 🚩 Red flag: Skipped cadences (Continued After 4th Reschedule) */}
      {skipped.length>0&&<div style={{background:"rgba(127,29,29,.05)",border:"1px solid rgba(127,29,29,.25)",borderRadius:12,padding:16,marginBottom:16}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
          <span style={{fontSize:16,fontWeight:600,color:"#7f1d1d"}}>🚩 Skipped Cadences — Prior Day</span>
          <span style={{fontSize:10,fontWeight:600,padding:"2px 8px",borderRadius:20,background:"rgba(127,29,29,.12)",color:"#7f1d1d"}}>{skipped.length} CSMs</span>
          {skipped.some(c=>c.skippedFourthCount>0)&&<span style={{fontSize:10,fontWeight:600,padding:"2px 8px",borderRadius:20,background:"rgba(220,38,38,.15)",color:"#991b1b"}}>
            ⚠ {skipped.filter(c=>c.skippedFourthCount>0).length} with 4th reschedule
          </span>}
          <span style={{fontSize:11,color:"#808080",marginLeft:"auto",fontStyle:"italic"}}>Click a name to see account details</span>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>
          {skipped.map(c=>{
            const i=lk(c.name)||{};
            const has4th = c.skippedFourthCount>0;
            return <div key={c.name}
              onClick={()=>onSelectCSM(c.name)}
              style={{...S.card,display:"flex",alignItems:"center",gap:10,cursor:"pointer",padding:12,
                borderLeft:`3px solid ${has4th?"#7f1d1d":"#d97706"}`}}>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:12,fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{dispName(c.name)}</div>
                <div style={{fontSize:11,color:"#808080",marginTop:2}}>{st(i.t||"")}{has4th?" · 🚩 4th reschedule":""}</div>
              </div>
              <div style={{textAlign:"center",flexShrink:0}}>
                <div style={{fontSize:20,fontWeight:500,color:has4th?"#7f1d1d":"#d97706"}}>{c.skippedCount}</div>
                <div style={{fontSize:10,color:has4th?"#7f1d1d":"#d97706"}}>skipped</div>
              </div>
            </div>;
          })}
        </div>
      </div>}

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
                <div style={{fontSize:12,fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{dispName(c.name)}</div>
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
              <thead><tr>{["CSM","Team","Skipped","Cadence","On-time","Overdue","Gap","BOQ","Retention"].map((h,j)=><th key={h} style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,padding:"0 8px 8px 0",textAlign:j>=2?"right":"left",borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>{h}</th>)}</tr></thead>
              <tbody>{attn.map(c=>{
                const i=lk(c.name)||{};
                const hasSkip=c.skippedCount>0;
                return <tr key={c.name} style={{cursor:"pointer",background:hasSkip?"rgba(127,29,29,.03)":"transparent"}} onClick={()=>onSelectCSM(c.name)}>
                  <td style={{padding:"7px 8px 7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)"}}>
                    <div style={{fontWeight:500,color:hasSkip?"#7f1d1d":"#29355D",display:"flex",alignItems:"center",gap:5}}>
                      {hasSkip&&<span>🚩</span>}{c.name}
                    </div>
                    <div style={{fontSize:11,color:"#808080",fontStyle:"italic",marginTop:2}}>
                      {hasSkip?"Continued After 4th Reschedule":c.cadPct===0?"No completions":c.cadPct<0.5?"Significantly behind":"Below average"}
                    </div>
                  </td>
                  <td style={{padding:"7px 8px 7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",color:"#808080"}}>{st(i.t||"")}</td>
                  <td style={{padding:"7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}>
                    {hasSkip?<span style={{fontSize:10,fontWeight:600,padding:"1px 7px",borderRadius:20,background:"rgba(127,29,29,.12)",color:"#7f1d1d"}}>{c.skippedCount}</span>:"--"}
                  </td>
                  <td style={{padding:"7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}><span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:"rgba(220,38,38,.12)",color:"#991b1b"}}>{Math.round(c.cadPct*100)}%</span></td>
                  <td style={{padding:"7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",fontWeight:500,color:c.otTotal>=3?pc(c.otPct):"#888"}}>{c.otTotal>=3?pp(c.otPct):"--"}</td>
                  <td style={{padding:"7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}>{c.overdueCount>0?<span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:"rgba(220,38,38,.1)",color:"#991b1b"}}>{c.overdueCount}</span>:"--"}</td>
                  <td style={{padding:"7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",fontWeight:500,color:"#dc2626"}}>{c.cadPct<0.9?Math.round((0.9-c.cadPct)*100)+"%":"--"}</td>
                  <td style={{padding:"7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:"#808080",fontSize:11}}>{c.bobBoq>0?fk(c.bobBoq):"--"}</td>
                  <td style={{padding:"7px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}>
                    {c.bobRet!=null?<span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,
                      background:c.bobRet>=0.91?"rgba(22,163,74,.1)":c.bobRet>=0.85?"rgba(217,119,6,.1)":"rgba(220,38,38,.1)",
                      color:c.bobRet>=0.91?"#166534":c.bobRet>=0.85?"#854d0e":"#991b1b"}}>
                      {pp(c.bobRet)}
                    </span>:"--"}
                  </td>
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
              <div style={{fontSize:12,fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{dispName(c.name)}</div>
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
function OverviewView({csms, allCSMs, bobRaw, bobAdj, history, callData, filterCoach, filterCSM, managerCoaches}) {
  const [dateFilter, setDateFilter] = useState("last_week");
  const [customFrom, setCustomFrom] = useState("");
  const [customTo,   setCustomTo]   = useState("");

  // ── Names in scope after coach/CSM filter (mirrors how csms is already filtered upstream) ──
  const inScopeNames = new Set(csms.map(c=>c.name));

  // ── Resolve a name into callData's keys (handles alias mismatches) ──
  const resolveCallCSM = n => callData[n] ? n : (Object.keys(callData).find(k=>norm(k)===n) || n);

  // ── Figure out the latest known date across history + calls, so "yesterday" etc. are relative to real data, not the browser clock ──
  const allHistoryDates = (history||[]).map(r=>r.date).filter(Boolean).sort();
  const allCallDays = Object.values(callData||{}).flatMap(d=>Object.keys(d)).sort();
  const latestKnown = [allHistoryDates[allHistoryDates.length-1], allCallDays[allCallDays.length-1]]
    .filter(Boolean).sort().pop();
  const latestDate = latestKnown ? parseLocalDate(latestKnown) : new Date();

  const range = getDateRange(dateFilter, latestDate, customFrom, customTo);
  const inRange = dateStr => {
    if (!range.from || !range.to) return true; // "all" = no filtering
    const d = parseLocalDate(dateStr);
    return d >= range.from && d <= range.to;
  };

  // ── Filter history rows to the selected date range + in-scope CSMs ──
  const histInRange = (history||[]).filter(r => inScopeNames.has(r.name) && inRange(r.date));

  // ── Sum revenue / cadence / email across the range (these are activity counts) ──
  const sumBy = (arr, key) => arr.reduce((s,r)=>s+(r[key]||0), 0);
  const totalRev   = sumBy(histInRange, "rev");
  const totalMRR   = sumBy(histInRange, "mrr");
  const totalSent  = sumBy(histInRange, "sent");
  const totalOD    = csms.reduce((s,c)=>s+c.overdueCount, 0); // always live, per earlier decision
  const odCsmCount = csms.filter(c=>c.overdueCount>0).length;

  // Email open rate — weighted average using sent as weight (can't simply average rates)
  const emRows = histInRange.filter(r=>r.sent>0 && r.openRate!=null);
  const avgOpen = emRows.length ? sumBy(emRows.map(r=>({v:r.sent*r.openRate})), "v") / sumBy(emRows,"sent") : 0;

  // Cadence on-time — weighted by ontime_total
  const otRows = histInRange.filter(r=>r.otTotal>=1 && r.otPct!=null);
  const avgOT = otRows.length ? sumBy(otRows.map(r=>({v:r.otTotal*r.otPct})), "v") / sumBy(otRows,"otTotal") : 0;

  // ── Retention — latest snapshot IN RANGE per CSM (point-in-time, not summed); falls back to live BOB if no history yet ──
  const getBobRet = (csm) => {
    // Prefer latest history snapshot within the selected range
    const csmHist = histInRange.filter(r=>r.name===csm.name && r.bobRet!=null).sort((a,b)=>a.date.localeCompare(b.date));
    if (csmHist.length > 0) return csmHist[csmHist.length-1].bobRet;
    // Fall back to live BOB sheet (current state) — only meaningful for "all"/current-ish views
    if (bobRaw && bobRaw.bob) {
      const key = Object.keys(bobRaw.bob).find(k => norm(k)===csm.name || k===csm.name);
      if (key) {
        const d = bobRaw.bob[key];
        const adjKey = bobAdj ? Object.keys(bobAdj).find(k=>norm(k)===csm.name||k===csm.name) : null;
        const lcmDelta = adjKey ? (bobAdj[adjKey].lcmDelta||0) : 0;
        const adjLcm = (d.lcm||0) + lcmDelta;
        return d.boq > 0 ? adjLcm / d.boq : null;
      }
    }
    return csm.bobRet;
  };
  const bobC2 = csms.filter(c=>getBobRet(c)!=null);
  const avgRet = bobC2.length ? bobC2.reduce((s,c)=>s+getBobRet(c),0)/bobC2.length : null;
  const aboveGoal = bobC2.filter(c=>getBobRet(c)>=0.91).length;
  const hasRetentionHistory = histInRange.some(r=>r.bobRet!=null);

  // ── Call data aggregation over the same date range ──
  const callDaysInRange = allCallDays.filter(inRange);
  const aggCalls = (names) => {
    const acc = {completed:0, noShow:0, cancelled:0, total:0, byCsm:{}};
    names.forEach(n=>{
      const key = resolveCallCSM(n);
      callDaysInRange.forEach(day=>{
        const dData = (callData[key]||{})[day] || {};
        Object.values(dData).forEach(d=>{
          if (!acc.byCsm[n]) acc.byCsm[n] = {completed:0,noShow:0,cancelled:0,total:0};
          acc.byCsm[n].completed += d.completed||0;
          acc.byCsm[n].noShow    += d.noShow||0;
          acc.byCsm[n].cancelled += d.cancelled||0;
          acc.byCsm[n].total      = acc.byCsm[n].completed + acc.byCsm[n].noShow + acc.byCsm[n].cancelled;
          acc.completed += d.completed||0;
          acc.noShow    += d.noShow||0;
          acc.cancelled += d.cancelled||0;
        });
      });
    });
    acc.total = acc.completed + acc.noShow + acc.cancelled;
    acc.compRate = acc.total>0 ? acc.completed/acc.total : 0;
    return acc;
  };
  const callTotals = aggCalls(csms.map(c=>c.name));
  const compColor = r => r>=0.85?"#16a34a":r>=0.70?"#d97706":"#dc2626";

  // ── Revenue + retention by team (revenue summed from history, retention from getBobRet) ──
  const teamRev={};
  histInRange.forEach(r=>{
    const t = r.team || (lk(r.name)&&lk(r.name).t) || "";
    if (t) teamRev[t] = (teamRev[t]||0) + (r.rev||0);
  });
  const trs = Object.entries(teamRev).sort((a,b)=>b[1]-a[1]);
  const maxR = trs[0]&&trs[0][1] || 1;

  const teamRetMap = {};
  csms.forEach(c=>{
    const t = c.team || (lk(c.name)&&lk(c.name).t) || "";
    const ret = getBobRet(c);
    if (t && ret!=null) { if(!teamRetMap[t]) teamRetMap[t]=[]; teamRetMap[t].push(ret); }
  });
  const teamRetList = Object.entries(teamRetMap).map(([t,arr])=>[t, arr.reduce((s,v)=>s+v,0)/arr.length]).sort((a,b)=>b[1]-a[1]);

  // ── Top revenue earners (from history sum, falls back to live csm.rev if no history) ──
  const revByName = {};
  histInRange.forEach(r=>{ revByName[r.name] = (revByName[r.name]||0) + (r.rev||0); });
  const hasRevHistory = Object.keys(revByName).length > 0;
  const topRev = hasRevHistory
    ? csms.map(c=>({name:c.name, rev:revByName[c.name]||0})).filter(c=>c.rev>0).sort((a,b)=>b.rev-a.rev).slice(0,8)
    : [...csms].filter(c=>c.rev>0).sort((a,b)=>b.rev-a.rev).slice(0,8).map(c=>({name:c.name, rev:c.rev}));
  const maxRI = topRev[0]&&topRev[0].rev || 1;

  // ── Call attendance by CSM — worst-first so problem areas surface ──
  const callRows = csms.map(c=>{
    const t = callTotals.byCsm[c.name];
    if (!t || t.total===0) return null;
    return {name:c.name, ...t, rate: t.completed/t.total};
  }).filter(Boolean).sort((a,b)=>a.rate-b.rate);

  const hbar=(name,pct,val,col)=>(
    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,fontSize:12}}>
      <span style={{width:140,flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{name}</span>
      <div style={{flex:1,height:4,background:"#ECEEF1",borderRadius:2,overflow:"hidden"}}>
        <div style={{width:Math.min((pct||0)*100,100).toFixed(1)+"%",height:"100%",background:col,opacity:.85,borderRadius:2}}/>
      </div>
      <span style={{width:50,textAlign:"right",fontSize:11,color:"#808080",flexShrink:0}}>{val}</span>
    </div>
  );

  const rangeLabel = dateFilter==="all" ? "All time"
    : range.from && range.to ? range.from.toLocaleDateString("en-US",{month:"short",day:"numeric"})+" – "+range.to.toLocaleDateString("en-US",{month:"short",day:"numeric"})
    : "";

  return (
    <div>
      {/* Date filter pills */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16,flexWrap:"wrap",gap:8}}>
        <div>
          <div style={{fontSize:13,fontWeight:600,color:"#29355D"}}>Team overview</div>
          {rangeLabel&&<div style={{fontSize:11,color:"#808080"}}>{rangeLabel}</div>}
        </div>
        <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
          {[["all","All"],["ytd","YTD"],["Q2 2026","Q2 2026"],["Q3 2026","Q3 2026"],["yesterday","Yesterday"],["this_week","This week"],["last_week","Last week"],["last_month","Last month"],["custom","Custom"]].map(([v,l])=>(
            <button key={v} onClick={()=>setDateFilter(v)}
              style={{padding:"4px 10px",borderRadius:20,border:"0.5px solid "+(dateFilter===v?"#29355D":"rgba(41,53,93,.15)"),
                background:dateFilter===v?"#29355D":"#fff",color:dateFilter===v?"#fff":"#808080",
                fontSize:11,fontWeight:500,cursor:"pointer"}}>
              {l}
            </button>
          ))}
          {dateFilter==="custom"&&<>
            <input type="date" value={customFrom} onChange={e=>setCustomFrom(e.target.value)}
              style={{padding:"4px 8px",borderRadius:8,border:"0.5px solid rgba(41,53,93,.2)",fontSize:11}}/>
            <span style={{fontSize:11,color:"#808080"}}>to</span>
            <input type="date" value={customTo} onChange={e=>setCustomTo(e.target.value)}
              style={{padding:"4px 8px",borderRadius:8,border:"0.5px solid rgba(41,53,93,.2)",fontSize:11}}/>
          </>}
        </div>
      </div>

      {/* Headline stat cards */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,minmax(0,1fr))",gap:12,marginBottom:14}}>
        <div style={{background:"#ECEEF1",borderRadius:"0 0 10px 10px",padding:"12px 14px",borderTop:"3px solid #FF5000"}}>
          <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:4}}>Revenue</div>
          <div style={{fontSize:24,fontWeight:600,color:"#FF5000",lineHeight:1,marginBottom:4}}>{fd(totalRev)}</div>
          <div style={{fontSize:10,color:"#808080"}}>MRR {fk(totalMRR)}{!hasRevHistory&&" · current"}</div>
        </div>
        <div style={{background:"#ECEEF1",borderRadius:"0 0 10px 10px",padding:"12px 14px",borderTop:"3px solid #29355D"}}>
          <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:4}}>Retention</div>
          <div style={{fontSize:24,fontWeight:600,color:avgRet!=null?bc(avgRet,0.91,0.85):"#808080",lineHeight:1,marginBottom:4}}>{avgRet!=null?pp(avgRet):"--"}</div>
          <div style={{fontSize:10,color:"#808080"}}>{aboveGoal} of {bobC2.length} at goal{!hasRetentionHistory&&" · current"}</div>
        </div>
        <div style={{background:"#ECEEF1",borderRadius:"0 0 10px 10px",padding:"12px 14px",borderTop:"3px solid #16a34a"}}>
          <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:4}}>Call completion</div>
          <div style={{fontSize:24,fontWeight:600,color:compColor(callTotals.compRate),lineHeight:1,marginBottom:4}}>{callTotals.total>0?pp(callTotals.compRate):"--"}</div>
          {callTotals.total>0&&<div style={{height:5,background:"rgba(0,0,0,.08)",borderRadius:3,marginBottom:4,overflow:"hidden"}}>
            <div style={{width:(callTotals.compRate*100).toFixed(1)+"%",height:"100%",background:compColor(callTotals.compRate),borderRadius:3}}/>
          </div>}
          <div style={{fontSize:10,color:"#808080"}}>{callTotals.completed} of {callTotals.total} booked</div>
        </div>
        <div style={{background:"#ECEEF1",borderRadius:"0 0 10px 10px",padding:"12px 14px",borderTop:"3px solid #dc2626"}}>
          <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:4}}>Overdue tasks</div>
          <div style={{fontSize:24,fontWeight:600,color:"#dc2626",lineHeight:1,marginBottom:4}}>{totalOD}</div>
          <div style={{fontSize:10,color:"#808080"}}>across {odCsmCount} CSMs · current</div>
        </div>
      </div>

      {/* Secondary stat row */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:16}}>
        <div style={S.card}>
          <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:4}}>Cadence on-time</div>
          <div style={{fontSize:20,fontWeight:600,color:otRows.length?bc(avgOT,0.8,0.6):"#808080"}}>{otRows.length?pp(avgOT):"--"}</div>
          <div style={{fontSize:10,color:"#808080"}}>goal 85%+ · {otRows.length} CSM-days tracked</div>
        </div>
        <div style={S.card}>
          <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:4}}>Email open rate</div>
          <div style={{fontSize:20,fontWeight:600,color:emRows.length?(avgOpen>=0.7?"#16a34a":"#d97706"):"#808080"}}>{emRows.length?pp(avgOpen):"--"}</div>
          <div style={{fontSize:10,color:"#808080"}}>target 70%+ · {totalSent} emails sent</div>
        </div>
      </div>

      {/* Call attendance by CSM — worst first */}
      {callRows.length>0&&<div style={{...S.card,marginBottom:16}}>
        <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:14}}>Call attendance by CSM — needs attention first</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(2,minmax(0,1fr))",gap:"8px 24px"}}>
          {callRows.slice(0,10).map(r=>(
            <div key={r.name} style={{display:"flex",alignItems:"center",gap:10,fontSize:12}}>
              <span style={{width:130,flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",fontWeight:500}}>{dispName(r.name)}</span>
              <div style={{flex:1,height:5,background:"#ECEEF1",borderRadius:3,overflow:"hidden"}}>
                <div style={{width:Math.min(r.rate*100,100).toFixed(1)+"%",height:"100%",background:compColor(r.rate),borderRadius:3}}/>
              </div>
              <span style={{width:36,textAlign:"right",fontSize:11,fontWeight:600,color:compColor(r.rate),flexShrink:0}}>{(r.rate*100).toFixed(0)}%</span>
            </div>
          ))}
        </div>
      </div>}

      {/* Revenue + Retention by team */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
        <div style={S.card}>
          <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>Revenue by team</div>
          {trs.length===0
            ? <div style={{color:"#808080",fontSize:12}}>No revenue in this period</div>
            : trs.map(([t,v])=>hbar(st(t),v/maxR,fk(v),TEAM_COLS[t]||"#888"))}
        </div>
        <div style={S.card}>
          <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>Retention by team — goal 91%</div>
          {teamRetList.length===0
            ? <div style={{color:"#808080",fontSize:12}}>No retention data available</div>
            : teamRetList.map(([t,v])=>hbar(st(t),v,pp(v),bc(v,0.91,0.85)))}
        </div>
      </div>

      {/* Top revenue earners */}
      {topRev.length>0&&<div style={{...S.card,marginBottom:16}}>
        <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>Top CSM revenue</div>
        {topRev.map(c=>hbar(dispName(c.name).split(" ").slice(0,2).join(" "),c.rev/maxRI,fk(c.rev),"#FF5000"))}
      </div>}

      {/* Overdue + Skipped cadence */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        <div style={S.card}>
          <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>Overdue cadence · current</div>
          {csms.filter(c=>c.overdueCount>0).length===0
            ? <div style={{color:"#808080",fontSize:12}}>No overdue tasks</div>
            : csms.filter(c=>c.overdueCount>0).sort((a,b)=>b.overdueCount-a.overdueCount).map(c=>(
              <div key={c.name} style={{display:"flex",alignItems:"center",gap:8,padding:"4px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)"}}>
                <span style={{flex:1,fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{dispName(c.name)}</span>
                <span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:"rgba(220,38,38,.1)",color:"#991b1b"}}>{c.overdueCount}</span>
                {c.newToday>0&&<span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:"rgba(83,120,252,.1)",color:"#1e3a8a"}}>{c.newToday} new</span>}
              </div>
            ))}
        </div>
        <div style={S.card}>
          <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>Skipped cadences — prior day · current</div>
          {csms.filter(c=>(c.skippedCount||0)>0).length===0
            ? <div style={{color:"#808080",fontSize:12}}>No skipped cadences</div>
            : csms.filter(c=>(c.skippedCount||0)>0).sort((a,b)=>b.skippedCount-a.skippedCount).map(c=>(
              <div key={c.name} style={{display:"flex",alignItems:"center",gap:8,padding:"4px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)"}}>
                <span style={{flex:1,fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{dispName(c.name)}</span>
                <span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:"rgba(217,119,6,.1)",color:"#854d0e"}}>{c.skippedCount} skipped</span>
              </div>
            ))}
        </div>
      </div>
    </div>
  );
}

// ── LEADERBOARD TAB ────────────────────────────────────────────────────────
function LeaderboardView({csms, bobRaw}) {
  const [sort,setSort]=useState({col:"rev",dir:"desc"});

  // Calculate retention directly from bobRaw lcm/boq — never use stored ret field
  const getBobRet = c => {
    if (bobRaw && bobRaw.bob) {
      const k = Object.keys(bobRaw.bob).find(k => norm(k)===c.name || k===c.name);
      if (k) {
        const d = bobRaw.bob[k];
        if (d.boq > 0 && d.lcm != null) return d.lcm / d.boq;
      }
    }
    // Fall back to c.bobRet (already fixed to be lcm/boq in buildCSMs)
    return c.bobRet != null ? c.bobRet : null;
  };

  // For each column, define what the "real" sortable value is.
  // null means "no data" — always sorted to the bottom regardless of direction.
  const getVal = (c, col) => {
    switch(col) {
      case "rev":          return c.rev > 0 ? c.rev : null;
      case "sent":         return c.sent > 0 ? c.sent : null;
      case "openRate":     return c.sent > 0 ? c.openRate : null;
      case "cadPct":       return c.cadCount > 0 ? c.cadPct : null;
      case "otPct":        return c.otTotal >= 3 ? c.otPct : null;
      case "overdueCount": return c.overdueCount > 0 ? c.overdueCount : null;
      case "bobBoq":       return c.bobBoq > 0 ? c.bobBoq : null;
      case "bobRet":       return getBobRet(c);
      default:             return null;
    }
  };

  const sorted=[...csms].sort((a,b)=>{
    const av = getVal(a, sort.col);
    const bv = getVal(b, sort.col);
    // Nulls always go to the bottom
    if (av === null && bv === null) return a.name.localeCompare(b.name);
    if (av === null) return 1;
    if (bv === null) return -1;
    // Both have values — sort by direction
    if (av !== bv) return sort.dir === "desc" ? bv - av : av - bv;
    // Tie-break alphabetically
    return a.name.localeCompare(b.name);
  });

  const medals=["🥇","🥈","🥉"];
  const th=(col,lbl)=>(
    <th onClick={()=>setSort(s=>({col,dir:s.col===col&&s.dir==="desc"?"asc":"desc"}))}
      style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,padding:"0 0 8px",textAlign:"right",cursor:"pointer",borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>
      {lbl}{sort.col===col?(sort.dir==="desc"?" ▼":" ▲"):<span style={{color:"#ccc",fontSize:9}}> ↕</span>}
    </th>
  );
  return (
    <div style={S.card}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
        <thead><tr>
          <th style={{width:28,fontSize:10,color:"#808080",fontWeight:500,padding:"0 0 8px",textAlign:"left",borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>#</th>
          <th style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,padding:"0 0 8px",textAlign:"left",borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>CSM</th>
          <th style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,padding:"0 0 8px",textAlign:"left",borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>Team</th>
          {th("rev","Revenue")}{th("sent","Emails")}{th("openRate","Open %")}{th("cadPct","Cadence")}{th("otPct","On-time %")}{th("overdueCount","Overdue")}{th("bobBoq","BOQ $")}{th("bobRet","Retention")}
        </tr></thead>
        <tbody>{sorted.map((c,i)=>{
          const info=lk(c.name)||{};
          const col=TEAM_COLS[info.t||c.team]||"#888";
          return <tr key={c.name}>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)"}}>{i<3?medals[i]:(i+1)+"."}</td>
            <td style={{padding:"9px 8px 9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",fontWeight:500}}>{dispName(c.name)}</td>
            <td style={{padding:"9px 8px 9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)"}}><span style={{display:"inline-block",width:6,height:6,borderRadius:"50%",background:col,marginRight:5,verticalAlign:"middle"}}/><span style={{color:"#808080",fontSize:11}}>{st(info.t||c.team)}</span></td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:"#FF5000",fontWeight:500}}>{c.rev>0?fd(c.rev):"--"}</td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}>{c.sent>0?c.sent:"--"}</td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",fontWeight:500,color:c.sent>0?pc(c.openRate):"#888"}}>{c.sent>0?pp(c.openRate):"--"}</td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",fontWeight:500,color:c.cadCount>0?pc(c.cadPct):"#888"}}>{c.cadCount>0?pp(c.cadPct):"--"}</td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",fontWeight:500,color:c.otTotal>=3?pc(c.otPct):"#888"}}>{c.otTotal>=3?pp(c.otPct):"--"}</td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}>{c.overdueCount>0?<span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:"rgba(220,38,38,.1)",color:"#991b1b"}}>{c.overdueCount}</span>:"--"}</td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:"#808080",fontSize:11}}>{c.bobBoq>0?fk(c.bobBoq):"--"}</td>
            {(()=>{const r=getBobRet(c);return <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}>{r!=null?<span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:r>=0.91?"rgba(22,163,74,.1)":r>=0.85?"rgba(217,119,6,.1)":"rgba(220,38,38,.1)",color:r>=0.91?"#166534":r>=0.85?"#854d0e":"#991b1b"}}>{pp(r)}</span>:"--"}</td>})()}
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
      <td style={tdBase}>{dispName(c.name)}</td>
      <td style={{...tdBase,textAlign:"right"}}>{c.sent}</td>
      <td style={{...tdBase,textAlign:"right",fontWeight:500,color:pc(c.openRate)}}>{pp(c.openRate)}</td>
      <td style={{...tdBase,textAlign:"right",fontWeight:500,color:pc(c.replyRate)}}>{pp(c.replyRate)}</td>
    </tr>
  }));

  // On-time table rows
  const otRows = csms.filter(c=>c.otTotal>=3).map(c=>({
    name: c.name, otTotal: c.otTotal, otOnTime: c.otOnTime, otPct: c.otPct,
    _render: (i) => <tr key={c.name}>
      <td style={tdBase}>{dispName(c.name)}</td>
      <td style={{...tdBase,textAlign:"right",color:"#808080"}}>{c.otTotal}</td>
      <td style={{...tdBase,textAlign:"right"}}>{c.otOnTime}</td>
      <td style={{...tdBase,textAlign:"right",fontWeight:500,color:bc(c.otPct,0.8,0.6)}}>{pp(c.otPct)}</td>
    </tr>
  }));

  // Due/overdue table rows
  const dueRows = csms.filter(c=>c.dueCount>0).map(c=>({
    name: c.name, dueCount: c.dueCount, overdueCount: c.overdueCount, newToday: c.newToday,
    _render: (i) => <tr key={c.name}>
      <td style={tdBase}>{dispName(c.name)}</td>
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
        <SortableTable
          title="Cadence completions — yesterday"
          defaultCol="cadPct"
          defaultDir="desc"
          cols={[
            {key:"name",     label:"CSM",        right:false},
            {key:"cadCount", label:"Tasks",      right:true},
            {key:"cadPct",   label:"Completed %",right:true},
          ]}
          rows={csms.filter(c=>c.cadCount>0).map(c=>({
            name:     c.name,
            cadCount: c.cadCount,
            cadPct:   c.cadPct,
            _render: (i) => <tr key={c.name}>
              <td style={{padding:"6px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",fontSize:12}}>{dispName(c.name)}</td>
              <td style={{padding:"6px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",fontSize:12,color:"#808080"}}>{c.cadCount}</td>
              <td style={{padding:"6px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",fontSize:12}}>
                <span style={{fontWeight:500,padding:"1px 8px",borderRadius:20,fontSize:10,
                  background: c.cadPct>=0.9?"rgba(22,163,74,.1)":c.cadPct>=0.5?"rgba(217,119,6,.1)":"rgba(220,38,38,.1)",
                  color:      c.cadPct>=0.9?"#166534":c.cadPct>=0.5?"#854d0e":"#991b1b"}}>
                  {pp(c.cadPct)}
                </span>
              </td>
            </tr>
          }))}
        />
      </div>
    </div>
  );
}

// ── TRENDS VIEW ────────────────────────────────────────────────────────────
function TrendsView({history, csms, filterCoach, filterCSM, callData={}, qamc={}, qass={}}) {
  const [metric, setMetric] = useState("otPct");
  const [view,   setView]         = useState("team"); // "team" | "csm"
  const [trendsTab, setTrendsTab]   = useState("performance");
  const [callDateFilter, setCallDateFilter] = useState("all");
  const [callCustomFrom, setCallCustomFrom] = useState("");
  const [callCustomTo,   setCallCustomTo]   = useState("");
  const [callSelectedCSM, setCallSelectedCSM] = useState(null);
  const [callSelectedSvc, setCallSelectedSvc] = useState(null);
  const [callCompare, setCallCompare]         = useState(false);
  const [qaType, setQaType]                   = useState("mc");
  const [qaMonth, setQaMonth]                 = useState(null);
  const [qaCompare, setQaCompare]             = useState(false);
  const [qaSortCol, setQaSortCol]             = useState("total");
  const [qaSortDir, setQaSortDir]             = useState("asc");

  const weeks = getWeeks(history);
  const trends = buildTrends(history);

  if (weeks.length < 1) {
    return (
      <div style={{...S.card, textAlign:"center", padding:40}}>
        <div style={{fontSize:32, marginBottom:12}}>📊</div>
        <div style={{fontSize:16, fontWeight:500, color:"#29355D", marginBottom:8}}>No trend data yet</div>
        <div style={{fontSize:13, color:"#808080", maxWidth:420, margin:"0 auto", lineHeight:1.6}}>
          The history sheet was found but no valid rows could be parsed.
          Check that the snapshot script is writing date values in YYYY-MM-DD format in column A.
        </div>
      </div>
    );
  }

  const METRICS = [
    {key:"otPct",    label:"On-time %",   format:v=>v!=null?Math.round(v*100)+"%":"--", color:"#5378FC"},
    {key:"cadPct",   label:"Cadence %",   format:v=>v!=null?Math.round(v*100)+"%":"--", color:"#29355D"},
    {key:"openRate", label:"Email open %",format:v=>v!=null?Math.round(v*100)+"%":"--", color:"#d97706"},
    {key:"rev",      label:"Revenue",     format:v=>v!=null&&v>0?"$"+Math.round(v).toLocaleString():"--", color:"#FF5000"},
    {key:"overdueCount", label:"Overdue tasks", format:v=>v!=null&&v>0?v:"--", color:"#dc2626"},
  ];
  const activeMetric = METRICS.find(m=>m.key===metric);

  // Filter CSMs based on coach/CSM filter
  // getDet: look up BOB billing detail — prefers live CSV over hardcoded BOB_DETAIL
  const getDet = n => liveBobDet[n]||liveBobDet[norm(n)]||BOB_DETAIL[n]||BOB_DETAIL[norm(n)]||{};

  const filteredCSMs = csms.filter(c => {
    const i = lk(c.name);
    if (filterCoach && (i&&i.c||c.coach) !== filterCoach) return false;
    if (filterCSM && c.name !== filterCSM) return false;
    return true;
  });
  const csmNames = filteredCSMs.map(c=>c.name);

  // Build team-level weekly aggregates
  const teamWeekly = {};
  COACHES.forEach(coach => {
    const teamCSMs = csmNames.filter(n => { const i=lk(n); return i&&i.c===coach.e; });
    if (teamCSMs.length === 0) return;
    teamWeekly[coach.e] = weeks.map(week => {
      const vals = teamCSMs.map(n => {
        const snap = (trends[n]||[]).find(s=>s.week===week);
        return snap ? snap[metric] : null;
      }).filter(v=>v!=null);
      return vals.length > 0 ? vals.reduce((s,v)=>s+v,0)/vals.length : null;
    });
  });

  // Last two weeks for delta
  const lastWeek = weeks[weeks.length-1];
  const prevWeek = weeks[weeks.length-2];

  const thStyle = {fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,padding:"0 8px 8px 0",textAlign:"left",borderBottom:"0.5px solid rgba(41,53,93,.08)"};
  const thRStyle = {...thStyle, textAlign:"right"};
  const tdStyle = {padding:"8px 8px 8px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",fontSize:12};
  const tdRStyle = {...tdStyle, textAlign:"right"};

  // Build context string from trend data for AI analysis
  const buildTrendsContext = () => {
    const lines = [];
    const fmt = v => v!=null ? (Math.round(v*100)+"%") : "n/a";
    const fmtN = v => v!=null ? Math.round(v) : "n/a";

    lines.push("=== TRENDS ANALYSIS — THRYV CSM DASHBOARD ===");
    lines.push("Weeks available: " + weeks.join(", "));
    lines.push("Last snapshot: " + lastWeek + (prevWeek ? " | Previous: " + prevWeek : " | No prior week yet"));
    lines.push("");

    // Team-level summary
    lines.push("=== TEAM RETENTION TRENDS ===");
    COACHES.forEach(coach => {
      const teamCSMs = csmNames.filter(n => { const i=lk(n); return i&&i.c===coach.e; });
      if (!teamCSMs.length) return;
      const getTeamAvg = (week, key) => {
        const vals = teamCSMs.map(n => {
          const snap = (trends[n]||[]).find(s=>s.week===week);
          return snap ? snap[key] : null;
        }).filter(v=>v!=null);
        return vals.length ? vals.reduce((s,v)=>s+v,0)/vals.length : null;
      };
      const last = {
        ot: getTeamAvg(lastWeek,"otPct"),
        cad: getTeamAvg(lastWeek,"cadPct"),
        open: getTeamAvg(lastWeek,"openRate"),
        rev: getTeamAvg(lastWeek,"rev"),
        overdue: getTeamAvg(lastWeek,"overdueCount"),
      };
      const prev = prevWeek ? {
        ot: getTeamAvg(prevWeek,"otPct"),
        cad: getTeamAvg(prevWeek,"cadPct"),
        open: getTeamAvg(prevWeek,"openRate"),
      } : null;
      const delta = (a,b) => a!=null&&b!=null ? (a>b?"↑":"↓") : "";
      lines.push("COACH: " + coach.n + " (" + teamCSMs.length + " CSMs)");
      lines.push("  On-time: " + fmt(last.ot) + (prev?" "+delta(last.ot,prev.ot):"") +
        " | Cadence: " + fmt(last.cad) + (prev?" "+delta(last.cad,prev.cad):"") +
        " | Email open: " + fmt(last.open) + (prev?" "+delta(last.open,prev.open):"") +
        " | Overdue: " + fmtN(last.overdue));
      lines.push("");
    });

    // CSM-level detail — flag notable trends
    lines.push("=== CSM TREND HIGHLIGHTS ===");
    const declining = [], improving = [], consistent = [];
    csmNames.forEach(name => {
      const snaps = (trends[name]||[]).filter(s=>weeks.slice(-3).includes(s.week));
      if (snaps.length < 2) return;
      const last3ot = snaps.map(s=>s.otPct).filter(v=>v!=null);
      if (last3ot.length >= 2) {
        const trend = last3ot[last3ot.length-1] - last3ot[0];
        const latest = last3ot[last3ot.length-1];
        if (trend < -0.05) declining.push({name, metric:"on-time", val:fmt(latest), chg:fmt(trend)});
        else if (trend > 0.05) improving.push({name, metric:"on-time", val:fmt(latest), chg:"+"+fmt(trend)});
        else if (latest < 0.7) consistent.push({name, note:"consistently low on-time "+fmt(latest)});
      }
    });
    if (declining.length) {
      lines.push("📉 DECLINING (on-time dropping >5pp):");
      declining.forEach(d => lines.push("  " + d.name + ": " + d.val + " (" + d.chg + " vs 3wk ago)"));
    }
    if (improving.length) {
      lines.push("📈 IMPROVING (on-time up >5pp):");
      improving.forEach(d => lines.push("  " + d.name + ": " + d.val + " (+" + d.chg + " vs 3wk ago)"));
    }
    if (consistent.length) {
      lines.push("⚠️  CONSISTENTLY LOW:");
      consistent.forEach(d => lines.push("  " + d.name + ": " + d.note));
    }
    lines.push("");

    // All CSM week-by-week table
    lines.push("=== FULL CSM WEEKLY DATA (last " + Math.min(weeks.length,4) + " weeks) ===");
    lines.push("CSM | " + weeks.slice(-4).join(" | ") + " (on-time %)");
    csmNames.forEach(name => {
      const vals = weeks.slice(-4).map(w => {
        const s = (trends[name]||[]).find(x=>x.week===w);
        return s&&s.otPct!=null ? fmt(s.otPct) : "--";
      });
      lines.push(name + " | " + vals.join(" | "));
    });

    return lines.join("\n");
  };

  const [trendsAiCopied, setTrendsAiCopied] = useState(false);

  const runTrendsAI = async () => {
    const ctx = buildTrendsContext();
    const prompt = [
      "You are an expert CSM coaching advisor at Thryv. You have been given week-over-week trend data for the CSM team.",
      "",
      "Analyze this data and provide:",
      "1. TREND SUMMARY - what are the 2-3 most important patterns across the team?",
      "2. CONCERNS - which CSMs or teams are trending in the wrong direction and why?",
      "3. BRIGHT SPOTS - who is improving and what might be driving that?",
      "4. COACHING PRIORITIES - based on trends, what should coaches focus on this week?",
      "5. RECOMMENDED ACTIONS - 3 specific things to do based on the data",
      "",
      "Be specific - cite actual names and numbers. Keep response under 500 words.",
      "",
      "Here is the trend data:",
      "",
      ctx,
    ].join("\n");

    try {
      await navigator.clipboard.writeText(prompt);
    } catch(e) {
      const el = document.createElement("textarea");
      el.value = prompt;
      el.style.position="fixed"; el.style.opacity="0";
      document.body.appendChild(el);
      el.focus(); el.select();
      document.execCommand("copy");
      document.body.removeChild(el);
    }
    setTrendsAiCopied(true);
    setTimeout(() => setTrendsAiCopied(false), 4000);
  };

  return (
    <div>
      {/* Subtab bar */}
      <div style={{display:"flex",gap:4,marginBottom:20,background:"#ECEEF1",borderRadius:10,padding:4,width:"fit-content"}}>
        {[["performance","📈 Performance"],["calls","📞 Calls"],["qa","🎯 Call QA"]].map(([t,l])=>(
          <button key={t} onClick={()=>setTrendsTab(t)}
            style={{padding:"7px 18px",borderRadius:7,border:"none",fontSize:12,fontWeight:600,cursor:"pointer",
              background:trendsTab===t?"#29355D":"transparent",
              color:trendsTab===t?"#fff":"#808080",
              boxShadow:trendsTab===t?"0 1px 4px rgba(0,0,0,.12)":"none"}}>
            {l}
          </button>
        ))}
      </div>

      {/* Performance tab content */}
      {trendsTab==="performance"&&<div>

      {/* Metric picker + view toggle */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:20,flexWrap:"wrap",gap:12}}>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {METRICS.map(m=>(
            <button key={m.key} onClick={()=>setMetric(m.key)}
              style={{padding:"6px 14px",fontSize:12,fontWeight:500,borderRadius:20,border:"0.5px solid "+(metric===m.key?m.color:"rgba(41,53,93,.15)"),background:metric===m.key?m.color:"transparent",color:metric===m.key?"#fff":"#29355D",cursor:"pointer"}}>
              {m.label}
            </button>
          ))}
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <div style={{display:"flex",gap:4,background:"#ECEEF1",borderRadius:8,padding:3}}>
            {["team","csm"].map(v=>(
              <button key={v} onClick={()=>setView(v)}
                style={{padding:"5px 14px",fontSize:12,fontWeight:500,borderRadius:6,border:"none",background:view===v?"#fff":"transparent",color:view===v?"#29355D":"#808080",cursor:"pointer",boxShadow:view===v?"0 1px 3px rgba(0,0,0,.1)":"none"}}>
                {v==="team"?"By team":"By CSM"}
              </button>
            ))}
          </div>
          <button onClick={runTrendsAI}
            style={{display:"flex",alignItems:"center",gap:6,padding:"7px 14px",borderRadius:8,border:"none",
              background:trendsAiCopied?"#16a34a":"#FF5000",color:"#fff",
              fontSize:12,fontWeight:600,cursor:"pointer",transition:"background .3s",whiteSpace:"nowrap"}}>
            {trendsAiCopied?"✓ Copied — paste into Claude":"🤖 AI Analysis"}
          </button>
        </div>
      </div>
      {trendsAiCopied&&<div style={{marginBottom:16,padding:"12px 16px",borderRadius:10,
        background:"#29355D",color:"#fff",fontSize:13,lineHeight:1.9,display:"flex",alignItems:"center",gap:16,flexWrap:"wrap"}}>
        <span>📋 <strong>Trend prompt copied!</strong></span>
        <span>1. Open <a href="https://claude.ai/new" target="_blank" rel="noreferrer" style={{color:"#FF5000",fontWeight:700}}>claude.ai/new</a></span>
        <span>2. Click message box</span>
        <span>3. <strong>Ctrl+V</strong> → <strong>Enter</strong></span>
      </div>}

      {/* Summary: last week vs week before */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:12,marginBottom:20}}>
        {weeks.slice(-4).map((week, wi, arr) => {
          const prevW = arr[wi-1] || null;
          // All filtered CSMs average for this week
          const vals = csmNames.map(n=>{const s=(trends[n]||[]).find(r=>r.week===week);return s?s[metric]:null;}).filter(v=>v!=null);
          const avg = vals.length ? vals.reduce((s,v)=>s+v,0)/vals.length : null;
          const prevVals = prevW ? csmNames.map(n=>{const s=(trends[n]||[]).find(r=>r.week===prevW);return s?s[metric]:null;}).filter(v=>v!=null) : [];
          const prevAvg = prevVals.length ? prevVals.reduce((s,v)=>s+v,0)/prevVals.length : null;
          const isLatest = wi === arr.length-1;
          return (
            <div key={week} style={{background:isLatest?"#29355D":"#ECEEF1",borderRadius:8,padding:14}}>
              <div style={{fontSize:10,textTransform:"uppercase",color:isLatest?"rgba(255,255,255,.6)":"#808080",fontWeight:500,marginBottom:4}}>{week}{isLatest?" · Latest":""}</div>
              <div style={{fontSize:22,fontWeight:500,color:isLatest?"#fff":"#29355D",lineHeight:1}}>
                {activeMetric.format(avg)}
              </div>
              {prevAvg!=null&&avg!=null&&(
                <TrendArrow curr={avg} prev={prevAvg}
                  isPercent={metric!=="rev"&&metric!=="overdueCount"}
                  higherBetter={metric!=="overdueCount"}/>
              )}
              <div style={{fontSize:11,color:isLatest?"rgba(255,255,255,.5)":"#808080",marginTop:4}}>{vals.length} CSMs</div>
            </div>
          );
        })}
      </div>

      {/* Detail table */}
      {view === "team" ? (
        <div style={S.card}>
          <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>
            {activeMetric.label} — week over week by team
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr>
              <th style={thStyle}>Team</th>
              {weeks.slice(-6).map(w=><th key={w} style={thRStyle}>{w.replace(" "," ")}</th>)}
              <th style={thRStyle}>Trend</th>
            </tr></thead>
            <tbody>
              {COACHES.filter(coach=>{
                if (filterCoach && coach.e!==filterCoach) return false;
                if (filterCSM) { const i=lk(filterCSM); return i&&i.c===coach.e; }
                return true;
              }).map(coach=>{
                const wData = teamWeekly[coach.e] || weeks.map(()=>null);
                const lastVals = wData.slice(-6);
                const curr = wData[wData.length-1];
                const prev = wData[wData.length-2];
                return (
                  <tr key={coach.e}>
                    <td style={tdStyle}>
                      <span style={{display:"inline-block",width:8,height:8,borderRadius:"50%",background:coach.col,marginRight:6,verticalAlign:"middle"}}/>
                      {coach.n}
                    </td>
                    {weeks.slice(-6).map((w,wi)=>{
                      const v = lastVals[wi];
                      return <td key={w} style={{...tdRStyle, fontWeight:500, color:v!=null&&metric!=="rev"&&metric!=="overdueCount"?(metric==="overdueCount"?bc(1-Math.min(v/10,1),0.5,0.2):bc(v,metric==="rev"?0.7:metric==="otPct"?0.8:0.7,0.4)):"#29355D"}}>{activeMetric.format(v)}</td>;
                    })}
                    <td style={tdRStyle}>
                      <Sparkline values={weeks.slice(-8).map(w=>{const d=teamWeekly[coach.e];return d?d[weeks.indexOf(w)]:null;})} color={coach.col} width={60} height={20}/>
                      <TrendArrow curr={curr} prev={prev} isPercent={metric!=="rev"&&metric!=="overdueCount"} higherBetter={metric!=="overdueCount"}/>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      ) : (
        <div style={S.card}>
          <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>
            {activeMetric.label} — week over week by CSM
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr>
              <th style={thStyle}>CSM</th>
              <th style={thStyle}>Team</th>
              {weeks.slice(-5).map(w=><th key={w} style={thRStyle}>{w}</th>)}
              <th style={thRStyle}>Trend</th>
            </tr></thead>
            <tbody>
              {csmNames.sort().map(name=>{
                const snapshots = trends[name] || [];
                const weekVals = weeks.slice(-5).map(w=>{const s=snapshots.find(r=>r.week===w);return s?s[metric]:null;});
                const curr = weekVals[weekVals.length-1];
                const prev = weekVals[weekVals.length-2];
                const info = lk(name)||{};
                const col = TEAM_COLS[info.t||""]||"#888";
                return (
                  <tr key={name}>
                    <td style={{...tdStyle,fontWeight:500}}>{name}</td>
                    <td style={tdStyle}><span style={{display:"inline-block",width:6,height:6,borderRadius:"50%",background:col,marginRight:4,verticalAlign:"middle"}}/><span style={{color:"#808080"}}>{st(info.t||"")}</span></td>
                    {weekVals.map((v,i)=>(
                      <td key={i} style={{...tdRStyle,fontWeight:500,color:v!=null?(metric==="overdueCount"?( v>0?"#dc2626":"#16a34a"):bc(v,metric==="otPct"?0.8:metric==="rev"?1:0.7,0.4)):"#ccc"}}>{activeMetric.format(v)}</td>
                    ))}
                    <td style={tdRStyle}>
                      <Sparkline values={weekVals} color={col} width={50} height={18}/>
                      <TrendArrow curr={curr} prev={prev} isPercent={metric!=="rev"&&metric!=="overdueCount"} higherBetter={metric!=="overdueCount"}/>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      </div>} {/* end performance tab */}

      {/* ── CALLS TAB ────────────────────────────────────────────────────── */}
      {trendsTab==="calls"&&(()=>{
        const allCallWeeks = getCallWeeks(callData);
        if (allCallWeeks.length === 0) return (
          <div style={{...S.card, marginTop:20, textAlign:"center", padding:32}}>
            <div style={{fontSize:28, marginBottom:10}}>📞</div>
            <div style={{fontSize:15, fontWeight:600, color:"#29355D", marginBottom:6}}>No call data yet</div>
            <div style={{fontSize:12, color:"#808080", lineHeight:1.7, maxWidth:400, margin:"0 auto"}}>
              Paste your weekly bookings export into the <strong>calls</strong> tab — no cleanup needed.
            </div>
          </div>
        );

        // Date filter logic — relative to latest data, not today
        // Parse YYYY-MM-DD as local midnight to avoid UTC offset issues
        const toDate = d => {
          if (!d) return new Date(NaN);
          const parts = String(d).split("-");
          if (parts.length === 3) return new Date(+parts[0], +parts[1]-1, +parts[2]);
          return new Date(d);
        };
        const latestDate = allCallWeeks.length > 0 ? toDate(allCallWeeks[allCallWeeks.length-1]) : new Date();
        const filterWeek = w => {
          const wd = toDate(w);
          if (callDateFilter==="last_week") {
            const wkAgo = new Date(latestDate); wkAgo.setDate(wkAgo.getDate()-7);
            return wd >= wkAgo;
          }
          if (callDateFilter==="last_month") {
            const moAgo = new Date(latestDate); moAgo.setMonth(moAgo.getMonth()-1);
            return wd >= moAgo;
          }
          if (callDateFilter==="last_quarter") {
            const qAgo = new Date(latestDate); qAgo.setMonth(qAgo.getMonth()-3);
            return wd >= qAgo;
          }
          if (callDateFilter==="custom" && callCustomFrom && callCustomTo) {
            return wd >= toDate(callCustomFrom) && wd <= toDate(callCustomTo);
          }
          return true; // all
        };
        const callWeeks = allCallWeeks.filter(filterWeek);
        const lastCW  = callWeeks[callWeeks.length-1];
        const prevCW  = callWeeks[callWeeks.length-2];

        // Compare: prior equivalent period
        const priorWeeks = (() => {
          if (!callCompare || callDateFilter === "all") return [];
          const span = callWeeks.length; // number of weeks in current period
          // Find weeks just before the current period
          const currentFirst = callWeeks[0];
          return allCallWeeks.filter(w => w < currentFirst).slice(-span);
        })();
        const hasPrior = callCompare && priorWeeks.length > 0;
        const priorLabel = hasPrior ? (priorWeeks[0] + " – " + priorWeeks[priorWeeks.length-1]) : "";
        const curLabel   = callWeeks.length > 0 ? (callWeeks[0] + (callWeeks.length>1?" – "+lastCW:"")) : "";

        // Aggregate for a specific set of weeks
        const aggForWeeks = (csmList, svcFilter, weeks) => {
          const acc = {completed:0, noShow:0, cancelled:0, total:0, bySvc:{}, byCsm:{}};
          csmList.forEach(n => {
            weeks.forEach(w => {
              const wData = (callData[resolveCSM(n)]||{})[w]||{};
              Object.entries(wData).forEach(([svc,d])=>{
                if (svcFilter && svc!==svcFilter) return;
                if (!acc.bySvc[svc]) acc.bySvc[svc]={completed:0,noShow:0,cancelled:0};
                acc.bySvc[svc].completed += d.completed;
                acc.bySvc[svc].noShow    += d.noShow;
                acc.bySvc[svc].cancelled += (d.cancelled||0);
                if (!acc.byCsm[n]) acc.byCsm[n]={completed:0,noShow:0,cancelled:0};
                acc.byCsm[n].completed   += d.completed;
                acc.byCsm[n].noShow      += d.noShow;
                acc.byCsm[n].cancelled   += (d.cancelled||0);
                acc.completed += d.completed;
                acc.noShow    += d.noShow;
                acc.cancelled += (d.cancelled||0);
              });
            });
          });
          acc.total = acc.completed + acc.noShow + acc.cancelled;
          acc.rate  = acc.completed+acc.noShow>0 ? acc.noShow/(acc.completed+acc.noShow) : 0;
          return acc;
        };

        // CSM filter
        const filtNames = csms.filter(c=>{
          const i=lk(c.name);
          if (filterCoach&&(i&&i.c||c.coach)!==filterCoach) return false;
          if (filterCSM&&c.name!==filterCSM) return false;
          return true;
        }).map(c=>c.name);
        const resolveCSM = n => callData[n] ? n : Object.keys(callData).find(k=>norm(k)===n)||n;
        let callCSMs = filtNames.filter(n=>callData[n]||Object.keys(callData).find(k=>norm(k)===n));

        // Apply selected CSM drill-down
        if (callSelectedCSM) callCSMs = callCSMs.filter(n=>n===callSelectedCSM||resolveCSM(n)===callSelectedCSM);

        // Aggregate totals across filtered weeks
        const aggTotals = (csmList, svcFilter) => {
          const acc = {completed:0, noShow:0, total:0, bySvc:{}};
          csmList.forEach(n => {
            callWeeks.forEach(w => {
              const wData = (callData[resolveCSM(n)]||{})[w]||{};
              Object.entries(wData).forEach(([svc,d])=>{
                if (svcFilter && svc!==svcFilter) return;
                if (!acc.bySvc[svc]) acc.bySvc[svc]={completed:0,noShow:0};
                acc.bySvc[svc].completed += d.completed;
                acc.bySvc[svc].noShow    += d.noShow;
                acc.completed += d.completed;
                acc.noShow    += d.noShow;
              });
            });
          });
          acc.total = acc.completed + acc.noShow;
          acc.rate  = acc.total>0 ? acc.noShow/acc.total : 0;
          return acc;
        };

        const orgTotals  = aggForWeeks(callCSMs, callSelectedSvc, callWeeks);
        const priorTotals= hasPrior ? aggForWeeks(callCSMs, callSelectedSvc, priorWeeks) : null;
        const orgRate    = orgTotals.rate;
        const NO_SHOW_GOAL = 0.08;
        const rateColor = r => r<=0.05?"#16a34a":r<=0.10?"#d97706":"#dc2626";
        const svcTotals = orgTotals.bySvc;
        const deltaColor = d => d < 0 ? "#16a34a" : d > 0 ? "#dc2626" : "#808080";
        const fmtDelta = (a, b, isPct) => {
          if (b==null||b.total===0||a.total===0) return null;
          const d = isPct ? (a.rate - b.rate)*100 : a.total - b.total;
          return {v: (d>0?"+":"")+d.toFixed(isPct?1:0)+(isPct?"pp":""), neg: d<0, pos: d>0};
        };

        return (
          <div style={{marginTop:24}}>
            {/* Header + date filter */}
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14,flexWrap:"wrap",gap:8}}>
              <div>
                <div style={{fontSize:13,fontWeight:600,color:"#29355D"}}>
                  📞 Call Trends
                  {callSelectedCSM&&<span style={{marginLeft:8,fontSize:12,color:"#FF5000"}}> — {dispName(callSelectedCSM)}</span>}
                  {callSelectedSvc&&<span style={{marginLeft:8,fontSize:12,color:"#FF5000"}}> — {callSelectedSvc}</span>}
                </div>
                <div style={{fontSize:11,color:"#808080"}}>{callWeeks.length} day{callWeeks.length!==1?"s":""} · Latest: {lastCW}</div>
              </div>
              <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                {/* Date filter pills */}
                {[["all","All"],["last_week","Last week"],["last_month","Last month"],["last_quarter","Last quarter"],["custom","Custom"]].map(([v,l])=>(
                  <button key={v} onClick={()=>setCallDateFilter(v)}
                    style={{padding:"4px 10px",borderRadius:20,border:"0.5px solid "+(callDateFilter===v?"#29355D":"rgba(41,53,93,.15)"),
                      background:callDateFilter===v?"#29355D":"#fff",color:callDateFilter===v?"#fff":"#808080",
                      fontSize:11,fontWeight:500,cursor:"pointer"}}>
                    {l}
                  </button>
                ))}
                {/* Custom date pickers */}
                {callDateFilter==="custom"&&<>
                  <input type="date" value={callCustomFrom} onChange={e=>setCallCustomFrom(e.target.value)}
                    style={{padding:"4px 8px",borderRadius:8,border:"0.5px solid rgba(41,53,93,.2)",fontSize:11}}/>
                  <span style={{fontSize:11,color:"#808080"}}>to</span>
                  <input type="date" value={callCustomTo} onChange={e=>setCallCustomTo(e.target.value)}
                    style={{padding:"4px 8px",borderRadius:8,border:"0.5px solid rgba(41,53,93,.2)",fontSize:11}}/>
                </>}
                {/* Compare toggle */}
                <button onClick={()=>setCallCompare(c=>!c)}
                  style={{padding:"4px 12px",borderRadius:20,fontSize:11,fontWeight:600,cursor:"pointer",
                    border:"0.5px solid "+(callCompare?"#5378FC":"rgba(41,53,93,.2)"),
                    background:callCompare?"#5378FC":"#fff",
                    color:callCompare?"#fff":"#29355D"}}>
                  ⇄ Compare
                </button>
                {/* Clear filters */}
                {(callSelectedCSM||callSelectedSvc)&&(
                  <button onClick={()=>{setCallSelectedCSM(null);setCallSelectedSvc(null);}}
                    style={{padding:"4px 10px",borderRadius:20,border:"0.5px solid #FF5000",background:"#fff",
                      color:"#FF5000",fontSize:11,fontWeight:500,cursor:"pointer"}}>
                    ✕ Clear filter
                  </button>
                )}
              </div>
            </div>

            {/* Compare banner */}
            {hasPrior&&<div style={{marginBottom:12,padding:"8px 14px",borderRadius:8,background:"rgba(83,120,252,.06)",
              border:"0.5px solid rgba(83,120,252,.2)",fontSize:11,color:"#3b5bdb",display:"flex",gap:16,flexWrap:"wrap"}}>
              <span>🟦 <strong>Current:</strong> {curLabel}</span>
              <span>⇄</span>
              <span>⬜ <strong>Prior:</strong> {priorLabel}</span>
            </div>}

            {/* Org summary tiles — completion-rate focused */}
            {(()=>{
              const compRate = orgTotals.total>0 ? orgTotals.completed/orgTotals.total : 0;
              const cancelRate = orgTotals.total>0 ? (orgTotals.cancelled||0)/orgTotals.total : 0;
              const compColor = compRate>=0.85?"#16a34a":compRate>=0.70?"#d97706":"#dc2626";
              const priorCompRate = priorTotals&&priorTotals.total>0 ? priorTotals.completed/priorTotals.total : null;
              const tiles = [
                {l:"Completed", val:orgTotals.completed, sub:(compRate*100).toFixed(1)+"% completion rate", col:compColor, bar:compRate, priorVal:priorTotals?.completed},
                {l:"Cancelled",  val:orgTotals.cancelled||0, sub:(cancelRate*100).toFixed(1)+"% of total booked", col:"#d97706", bar:cancelRate, priorVal:priorTotals?.cancelled},
                {l:"No shows",  val:orgTotals.noShow, sub:(orgRate*100).toFixed(1)+"% — goal <8%", col:rateColor(orgRate), bar:orgRate, priorVal:priorTotals?.noShow},
                {l:"Total booked", val:orgTotals.total, col:"#29355D", priorVal:priorTotals?.total,
                  pills:[
                    {label:(compRate*100).toFixed(0)+"% kept",   bg:"#dcfce7", fg:"#166534"},
                    {label:(cancelRate*100).toFixed(0)+"% cancelled", bg:"#fef9c3", fg:"#854d0e"},
                    {label:((orgRate)*100).toFixed(0)+"% no-show", bg:"#fee2e2", fg:"#991b1b"},
                  ]},
              ];
              return (
                <div style={{display:"grid",gridTemplateColumns:"repeat(4,minmax(0,1fr))",gap:10,marginBottom:18}}>
                  {tiles.map(t=>{
                    const dVal = hasPrior&&t.priorVal!=null ? t.val - t.priorVal : null;
                    return (
                      <div key={t.l} style={{background:"#ECEEF1",borderRadius:"0 0 10px 10px",padding:"12px 14px",borderTop:"3px solid "+t.col}}>
                        <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:4}}>{t.l}</div>
                        <div style={{fontSize:24,fontWeight:600,color:t.col,lineHeight:1,marginBottom:4}}>{t.val}</div>
                        {t.bar!=null&&<div style={{height:5,background:"rgba(0,0,0,.08)",borderRadius:3,marginBottom:5,overflow:"hidden"}}>
                          <div style={{width:Math.min(t.bar*100,100).toFixed(1)+"%",height:"100%",background:t.col,borderRadius:3}}/>
                        </div>}
                        {t.pills&&<div style={{display:"flex",flexWrap:"wrap",gap:4,marginTop:4}}>
                          {t.pills.map(p=><span key={p.label} style={{fontSize:10,fontWeight:500,padding:"2px 6px",borderRadius:20,background:p.bg,color:p.fg}}>{p.label}</span>)}
                        </div>}
                        {!t.pills&&<div style={{fontSize:10,color:"#808080"}}>{t.sub}</div>}
                        {hasPrior&&dVal!=null&&<div style={{fontSize:10,marginTop:3,color:t.l==="No shows"?(dVal<0?"#16a34a":dVal>0?"#dc2626":"#808080"):(dVal>0?"#16a34a":dVal<0?"#dc2626":"#808080"),fontWeight:500}}>
                          {dVal>0?"+":""}{dVal} vs prior
                        </div>}
                      </div>
                    );
                  })}
                </div>
              );
            })()}

            {/* Service breakdown — completion rate bars */}
            {Object.keys(svcTotals).length>0&&(
              <div style={{...S.card,marginBottom:16}}>
                <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:14}}>Completion rate by service type — {curLabel||lastCW}</div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(220px,1fr))",gap:12}}>
                  {Object.entries(svcTotals).sort((a,b)=>{
                    const ta=a[1].completed+a[1].noShow+(a[1].cancelled||0);
                    const tb=b[1].completed+b[1].noShow+(b[1].cancelled||0);
                    return tb-ta;
                  }).map(([svc,d])=>{
                    const tot=d.completed+(d.noShow||0)+(d.cancelled||0);
                    const compRate=tot>0?d.completed/tot:0;
                    const compCol=compRate>=0.85?"#16a34a":compRate>=0.70?"#d97706":"#dc2626";
                    const isSelected=callSelectedSvc===svc;
                    return (
                      <div key={svc} onClick={()=>setCallSelectedSvc(isSelected?null:svc)}
                        style={{padding:"10px 12px",borderRadius:8,cursor:"pointer",
                          background:isSelected?"rgba(41,53,93,.06)":"transparent",
                          border:"0.5px solid "+(isSelected?"#29355D":"rgba(41,53,93,.1)"),
                          transition:"all .15s"}}>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:5}}>
                          <span style={{fontSize:12,fontWeight:500,color:"#29355D"}}>{svc}</span>
                          <span style={{fontSize:12,fontWeight:600,color:compCol}}>{(compRate*100).toFixed(0)}%</span>
                        </div>
                        <div style={{height:5,background:"rgba(0,0,0,.08)",borderRadius:3,marginBottom:5,overflow:"hidden"}}>
                          <div style={{width:(compRate*100).toFixed(1)+"%",height:"100%",background:compCol,borderRadius:3}}/>
                        </div>
                        <div style={{fontSize:11,color:"#808080"}}>
                          {d.completed} completed · {d.noShow||0} no-show · {d.cancelled||0} cancelled
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* CSM table */}
            <div style={S.card}>
              <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>
                CSM call performance{hasPrior?" — current vs prior period":lastCW?" — "+lastCW:""}
              </div>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr>
                  <th style={{padding:"0 8px 8px 0",textAlign:"left",fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>CSM</th>
                  <th style={{padding:"0 8px 8px 0",textAlign:"right",fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>Booked</th>
                  <th style={{padding:"0 8px 8px 0",textAlign:"right",fontSize:10,textTransform:"uppercase",color:"#16a34a",fontWeight:500,borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>Completed</th>
                  <th style={{padding:"0 8px 8px 0",textAlign:"left",fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,borderBottom:"0.5px solid rgba(41,53,93,.08)",minWidth:140}}>Completion rate</th>
                  <th style={{padding:"0 8px 8px 0",textAlign:"right",fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>No shows</th>
                  <th style={{padding:"0 8px 8px 0",textAlign:"right",fontSize:10,textTransform:"uppercase",color:"#d97706",fontWeight:500,borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>Cancelled</th>
                  {hasPrior
                    ? <th style={{padding:"0 8px 8px 0",textAlign:"right",fontSize:10,textTransform:"uppercase",color:"#5378FC",fontWeight:500,borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>vs Prior</th>
                    : <th style={{padding:"0 8px 8px 0",textAlign:"right",fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>vs Prior wk</th>}
                </tr></thead>
                <tbody>
                  {callCSMs.sort((a,b)=>{
                    const ta=aggTotals([a],callSelectedSvc);
                    const tb=aggTotals([b],callSelectedSvc);
                    const compA=ta.total>0?ta.completed/ta.total:0;
                    const compB=tb.total>0?tb.completed/tb.total:0;
                    return compB-compA;
                  }).map(n=>{
                    const t    = aggTotals([n], callSelectedSvc);
                    // Week-over-week delta using last two weeks
                    const tLast = lastCW ? (() => {
                      const wData = (callData[resolveCSM(n)]||{})[lastCW]||{};
                      let c=0,ns=0;
                      Object.entries(wData).forEach(([svc,d])=>{ if(!callSelectedSvc||svc===callSelectedSvc){c+=d.completed;ns+=d.noShow;} });
                      return {completed:c,noShow:ns,total:c+ns,rate:c+ns>0?ns/(c+ns):0};
                    })() : null;
                    const tPrev = prevCW ? (() => {
                      const wData = (callData[resolveCSM(n)]||{})[prevCW]||{};
                      let c=0,ns=0;
                      Object.entries(wData).forEach(([svc,d])=>{ if(!callSelectedSvc||svc===callSelectedSvc){c+=d.completed;ns+=d.noShow;} });
                      return {completed:c,noShow:ns,total:c+ns,rate:c+ns>0?ns/(c+ns):0};
                    })() : null;
                    // Delta: compare mode uses prior period, otherwise week-over-week
                    const csmPrior = hasPrior ? aggForWeeks([n], callSelectedSvc, priorWeeks) : null;
                    const delta = hasPrior
                      ? (csmPrior&&csmPrior.total>0&&t.total>0 ? t.rate-csmPrior.rate : null)
                      : (tLast&&tPrev&&tLast.total>0&&tPrev.total>0 ? tLast.rate-tPrev.rate : null);
                    const deltaRef = hasPrior ? csmPrior : null;
                    if (t.total===0) return null;
                    const isSelected = callSelectedCSM===n;
                    return (
                      <tr key={n} onClick={()=>setCallSelectedCSM(isSelected?null:n)}
                        style={{cursor:"pointer",background:isSelected?"rgba(41,53,93,.04)":"transparent"}}>
                        <td style={{padding:"8px 8px 8px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",fontWeight:isSelected?700:500,color:isSelected?"#29355D":"inherit"}}>{dispName(n)}</td>
                        <td style={{padding:"8px 8px 8px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:"#808080"}}>{t.total}</td>
                        <td style={{padding:"8px 8px 8px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:"#16a34a",fontWeight:500}}>{t.completed}</td>
                        <td style={{padding:"8px 8px 8px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)"}}>
                          {(()=>{
                            const compRate=t.total>0?t.completed/t.total:0;
                            const compCol=compRate>=0.85?"#16a34a":compRate>=0.70?"#d97706":"#dc2626";
                            return (
                              <div style={{display:"flex",alignItems:"center",gap:8}}>
                                <div style={{flex:1,height:5,background:"rgba(0,0,0,.08)",borderRadius:3,overflow:"hidden",minWidth:60}}>
                                  <div style={{width:(compRate*100).toFixed(1)+"%",height:"100%",background:compCol,borderRadius:3}}/>
                                </div>
                                <span style={{fontSize:11,fontWeight:600,color:compCol,minWidth:36,textAlign:"right"}}>{(compRate*100).toFixed(0)}%</span>
                              </div>
                            );
                          })()}
                        </td>
                        <td style={{padding:"8px 8px 8px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:t.noShow>0?"#dc2626":"#808080"}}>{t.noShow}</td>
                        <td style={{padding:"8px 8px 8px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:(t.cancelled||0)>0?"#d97706":"#808080"}}>{t.cancelled||0}</td>
                        <td style={{padding:"8px 8px 8px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",fontSize:11}}>
                          {hasPrior&&deltaRef!=null&&<div style={{fontSize:10,color:"#808080"}}>{deltaRef.total} calls / {(deltaRef.completed/Math.max(deltaRef.total,1)*100).toFixed(0)}%</div>}
                          {delta!=null
                            ? <span style={{color:delta>=0?"#16a34a":"#dc2626",fontWeight:500}}>
                                {delta>=0?"↑":"↓"}{Math.abs(delta*100).toFixed(1)}pp
                              </span>
                            : <span style={{color:"#808080"}}>--</span>}
                        </td>
                      </tr>
                    );
                  }).filter(Boolean)}
                </tbody>
              </table>
            </div>
          </div>
        );
      })()}

      {/* ── CALL QA TAB ──────────────────────────────────────────────────── */}
      {trendsTab==="qa"&&(()=>{
        const MC_CRITERIA = [
          {key:"pre_kickoff_email",      label:"Pre-Kickoff Email"},
          {key:"defined_csm_role",       label:"Defined CSM Role"},
          {key:"presented_journey_deck", label:"Journey Deck"},
          {key:"confirmed_priorities",   label:"Confirmed Priorities"},
          {key:"next_call_agenda",       label:"Next Call Agenda"},
          {key:"provided_actionable_task",label:"Actionable Task"},
          {key:"wrapup_email",           label:"Wrap-Up Email"},
        ];
        const SS_CRITERIA = [
          {key:"rapport_building",        label:"Rapport Building"},
          {key:"presented_journey_deck",  label:"Journey Deck"},
          {key:"agenda_open_items",       label:"Agenda / Open Items"},
          {key:"next_call_agenda",        label:"Next Call Agenda"},
          {key:"provided_actionable_task",label:"Actionable Task"},
          {key:"wrapup_email",            label:"Wrap-Up Email"},
          {key:"cer_notes",               label:"CER Notes"},
        ];
        const GOAL = 0.93;
        const scoreColor = v => v==null?"#e5e7eb":v>=0.93?"#16a34a":v>=0.80?"#d97706":"#dc2626";
        const scoreBg    = v => v==null?"rgba(0,0,0,.04)":v>=0.93?"rgba(22,163,74,.1)":v>=0.80?"rgba(217,119,6,.1)":"rgba(220,38,38,.1)";
        const fmtPct     = v => v==null?"N/A":(v*100).toFixed(1)+"%";

        const data = qaType==="mc" ? qamc : qass;
        const criteria = qaType==="mc" ? MC_CRITERIA : SS_CRITERIA;
        const months = Object.keys(data).sort();
        const activeMonth = qaMonth || months[months.length-1];
        const prevMonth = months[months.indexOf(activeMonth)-1];
        const hasData = months.length > 0;

        if (!hasData) return (
          <div style={{...S.card, marginTop:20, textAlign:"center", padding:32}}>
            <div style={{fontSize:28, marginBottom:10}}>🎯</div>
            <div style={{fontSize:15, fontWeight:600, color:"#29355D", marginBottom:6}}>No QA data yet</div>
            <div style={{fontSize:12, color:"#808080", lineHeight:1.7, maxWidth:440, margin:"0 auto"}}>
              Add your QA data to the <strong>qa mc</strong> and <strong>qa ss</strong> tabs in Google Sheets.<br/>
              Columns: <code>month, csm_name, audits, total_achievement</code> + criterion scores.
            </div>
          </div>
        );

        // Filter CSMs by coach/csm filter
        const filtNames = csms.filter(c=>{
          const i=lk(c.name);
          if (filterCoach&&(i&&i.c||c.coach)!==filterCoach) return false;
          if (filterCSM&&c.name!==filterCSM) return false;
          return true;
        }).map(c=>c.name);

        const monthData = data[activeMonth] || {};
        const prevData  = prevMonth ? (data[prevMonth]||{}) : {};

        // Filter to visible CSMs
        const visibleCSMs = Object.keys(monthData).filter(n => {
          if (filtNames.length === 0) return true;
          return filtNames.some(fn => fn===n || norm(fn)===n || fn===norm(n));
        }).sort((a,b) => {
          const va = qaSortCol==="total" ? (monthData[a]?.total||0) : (monthData[a]?.criteria[qaSortCol]||0);
          const vb = qaSortCol==="total" ? (monthData[b]?.total||0) : (monthData[b]?.criteria[qaSortCol]||0);
          return qaSortDir==="asc" ? va-vb : vb-va;
        });

        // Team averages
        const teamAvg = col => {
          const vals = visibleCSMs.map(n => col==="total" ? monthData[n]?.total : monthData[n]?.criteria[col]).filter(v=>v!=null);
          return vals.length ? vals.reduce((s,v)=>s+v,0)/vals.length : null;
        };
        const prevTeamAvg = col => {
          const vals = visibleCSMs.map(n => col==="total" ? prevData[n]?.total : prevData[n]?.criteria[col]).filter(v=>v!=null);
          return vals.length ? vals.reduce((s,v)=>s+v,0)/vals.length : null;
        };

        const totalAudits = visibleCSMs.reduce((s,n)=>s+(monthData[n]?.audits||0),0);
        const belowGoal   = visibleCSMs.filter(n=>(monthData[n]?.total||0)<0.80);

        return (
          <div style={{marginTop:24}}>
            {/* Header controls */}
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16,flexWrap:"wrap",gap:8}}>
              {/* Report type toggle */}
              <div style={{display:"flex",gap:4,background:"#ECEEF1",borderRadius:8,padding:3}}>
                {[["mc","MC Activation"],["ss","Setup & Strategy"]].map(([v,l])=>(
                  <button key={v} onClick={()=>setQaType(v)}
                    style={{padding:"6px 14px",fontSize:12,fontWeight:600,borderRadius:6,border:"none",cursor:"pointer",
                      background:qaType===v?"#29355D":"transparent",color:qaType===v?"#fff":"#808080"}}>
                    {l}
                  </button>
                ))}
              </div>
              {/* Month picker */}
              <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
                <span style={{fontSize:11,color:"#808080"}}>Month:</span>
                {months.map(m=>(
                  <button key={m} onClick={()=>setQaMonth(m)}
                    style={{padding:"4px 10px",borderRadius:20,fontSize:11,fontWeight:500,cursor:"pointer",
                      border:"0.5px solid "+(activeMonth===m?"#29355D":"rgba(41,53,93,.15)"),
                      background:activeMonth===m?"#29355D":"#fff",
                      color:activeMonth===m?"#fff":"#808080"}}>
                    {formatMonthLabel(m)}
                  </button>
                ))}
                {prevMonth&&<button onClick={()=>setQaCompare(c=>!c)}
                  style={{padding:"4px 12px",borderRadius:20,fontSize:11,fontWeight:600,cursor:"pointer",
                    border:"0.5px solid "+(qaCompare?"#5378FC":"rgba(41,53,93,.2)"),
                    background:qaCompare?"#5378FC":"#fff",color:qaCompare?"#fff":"#29355D"}}>
                  ⇄ vs {formatMonthLabel(prevMonth)}
                </button>}
              </div>
            </div>

            {/* Summary tiles */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:16}}>
              {[
                {l:"Total Audits",   v:totalAudits,                        col:"#29355D", sub:formatMonthLabel(activeMonth)},
                {l:"Team Average",   v:fmtPct(teamAvg("total")),           col:scoreColor(teamAvg("total")), sub:"Goal: 93%"},
                {l:"Above Goal ≥93%",v:visibleCSMs.filter(n=>(monthData[n]?.total||0)>=0.93).length, col:"#16a34a", sub:"of "+visibleCSMs.length+" CSMs"},
                {l:"Below 80%",      v:belowGoal.length,                   col:belowGoal.length>0?"#dc2626":"#16a34a", sub:belowGoal.length>0?belowGoal.map(n=>dispName(n)).slice(0,2).join(", ")+(belowGoal.length>2?" +more":""):"All above 80%"},
              ].map(t=>(
                <div key={t.l} style={{background:"#ECEEF1",borderRadius:8,padding:"12px 14px",borderTop:"2px solid "+t.col}}>
                  <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:4}}>{t.l}</div>
                  <div style={{fontSize:22,fontWeight:600,color:t.col,lineHeight:1}}>{t.v}</div>
                  <div style={{fontSize:10,color:"#808080",marginTop:3}}>{t.sub}</div>
                </div>
              ))}
            </div>

            {/* Criterion team averages bar */}
            <div style={{...S.card,marginBottom:16}}>
              <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>
                Team averages by criterion — {formatMonthLabel(activeMonth)}
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:8}}>
                {criteria.map(cr=>{
                  const cur  = teamAvg(cr.key);
                  const prev = qaCompare ? prevTeamAvg(cr.key) : null;
                  const delta = cur!=null&&prev!=null ? (cur-prev)*100 : null;
                  return (
                    <div key={cr.key} style={{padding:"10px 12px",borderRadius:8,background:scoreBg(cur),
                      border:"0.5px solid rgba(41,53,93,.08)"}}>
                      <div style={{fontSize:10,color:"#808080",marginBottom:4,fontWeight:500}}>{cr.label}</div>
                      <div style={{fontSize:18,fontWeight:600,color:scoreColor(cur)}}>{fmtPct(cur)}</div>
                      {qaCompare&&delta!=null&&<div style={{fontSize:10,marginTop:2,fontWeight:600,
                        color:delta>=0?"#16a34a":"#dc2626"}}>
                        {delta>=0?"↑":"↓"}{Math.abs(delta).toFixed(1)}pp vs prior
                      </div>}
                      {!qaCompare&&cur!=null&&<div style={{marginTop:6,height:3,borderRadius:2,background:"rgba(41,53,93,.08)"}}>
                        <div style={{height:3,borderRadius:2,background:scoreColor(cur),width:Math.min(cur*100,100)+"%"}}/>
                      </div>}
                    </div>
                  );
                })}
              </div>
            </div>

            {/* CSM scorecard table */}
            <div style={S.card}>
              <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>
                CSM scorecard — {formatMonthLabel(activeMonth)}{qaCompare&&prevMonth?" vs "+formatMonthLabel(prevMonth):""}
              </div>
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:11,minWidth:700}}>
                  <thead><tr>
                    {[{k:"name",l:"CSM"},{k:"audits",l:"Audits"},{k:"total",l:"Overall"},...criteria.map(c=>({k:c.key,l:c.label}))].map(h=>(
                      <th key={h.k} onClick={()=>{if(h.k!=="name"){setQaSortCol(h.k);setQaSortDir(d=>d==="asc"?"desc":"asc");}}}
                        style={{padding:"0 8px 8px 0",textAlign:h.k==="name"?"left":"right",fontSize:10,
                          textTransform:"uppercase",color:qaSortCol===h.k?"#29355D":"#808080",
                          fontWeight:qaSortCol===h.k?700:500,borderBottom:"0.5px solid rgba(41,53,93,.08)",
                          cursor:h.k!=="name"?"pointer":"default",whiteSpace:"nowrap"}}>
                        {h.l}{qaSortCol===h.k?(qaSortDir==="asc"?" ↑":" ↓"):""}
                      </th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {visibleCSMs.map(n=>{
                      const d = monthData[n]||{};
                      const p = qaCompare ? (prevData[n]||{}) : null;
                      return (
                        <tr key={n} style={{borderBottom:"0.5px solid rgba(41,53,93,.04)"}}>
                          <td style={{padding:"7px 8px 7px 0",fontWeight:500,whiteSpace:"nowrap"}}>{dispName(n)}</td>
                          <td style={{padding:"7px 8px 7px 0",textAlign:"right",color:"#808080"}}>{d.audits||"—"}</td>
                          <td style={{padding:"7px 8px 7px 0",textAlign:"right"}}>
                            <span style={{fontSize:11,fontWeight:600,padding:"2px 8px",borderRadius:20,
                              background:scoreBg(d.total),color:scoreColor(d.total)}}>
                              {fmtPct(d.total)}
                            </span>
                            {p&&p.total!=null&&d.total!=null&&<span style={{fontSize:10,marginLeft:4,
                              color:(d.total-p.total)>=0?"#16a34a":"#dc2626"}}>
                              {(d.total-p.total)>=0?"↑":"↓"}{Math.abs((d.total-p.total)*100).toFixed(1)}pp
                            </span>}
                          </td>
                          {criteria.map(cr=>{
                            const v = d.criteria?.[cr.key];
                            const pv = p?.criteria?.[cr.key];
                            return (
                              <td key={cr.key} style={{padding:"7px 8px 7px 0",textAlign:"right"}}>
                                <span style={{fontSize:10,padding:"1px 6px",borderRadius:20,
                                  background:scoreBg(v),color:scoreColor(v)}}>
                                  {fmtPct(v)}
                                </span>
                                {qaCompare&&pv!=null&&v!=null&&<div style={{fontSize:9,
                                  color:(v-pv)>=0?"#16a34a":"#dc2626"}}>
                                  {(v-pv)>=0?"↑":"↓"}{Math.abs((v-pv)*100).toFixed(1)}pp
                                </div>}
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  );
}

// ── DAILY DIGEST ────────────────────────────────────────────────────────────
function DigestView({csms, filterCoach, filterCSM, isCsmView, bobRaw, mcChurn, bcChurn,
  liveBobDet, callData, qamc, qass, skippedCSMs, bobAdj, history=[], getDet}) {

  const [period, setPeriod] = React.useState("week");
  // getDet fallback if not passed as prop
  const getDetFn = getDet || (n => BOB_DETAIL[n]||BOB_DETAIL[norm(n)]||{});
  const [expanded, setExpanded] = React.useState(null);
  const [aiCopied, setAiCopied] = React.useState(false);
  const [scoreFilter, setScoreFilter] = React.useState("all");

  const PERIODS = [
    {k:"yesterday", l:"Yesterday"},
    {k:"week",      l:"Last Week"},
    {k:"month",     l:"Last Month"},
    {k:"quarter",   l:"Last Quarter"},
  ];

  // ── Signal scoring ──────────────────────────────────────────────────────
  const scoreColor = s => s==="legend"?"#7c3aed":s==="green"?"#16a34a":s==="yellow"?"#d97706":s==="red"?"#dc2626":"#808080";
  const scoreBg    = s => s==="legend"?"rgba(124,58,237,.12)":s==="green"?"rgba(22,163,74,.1)":s==="yellow"?"rgba(217,119,6,.1)":"rgba(220,38,38,.1)";
  const scoreDot   = s => s==="legend"?"👑":s==="green"?"🟢":s==="yellow"?"🟡":"🔴";
  const scoreEmoji = s => s==="legend"?"★":s==="green"?"✓":s==="yellow"?"⚠":s==="red"?"✗":"—";

  const worstScore = scores => {
    const reds    = scores.filter(s=>s==="red").length;
    const yellows = scores.filter(s=>s==="yellow").length;
    const greens  = scores.filter(s=>s==="green"||s==="legend").length;
    if (reds >= 2) return "red";                          // 2+ reds = Needs Love
    if (reds === 1 || yellows > 0) return "yellow";       // 1 red or any yellow = Almost There
    if (greens > 0 && reds === 0 && yellows === 0) {
      if (scores.some(s=>s==="legend")) return "legend";  // all good + legend signal = Legend
      return "green";                                     // all green = Crushing It
    }
    return "gray";
  };

  const isLegend = csm => {
    const sigs = buildSignals(csm);
    const scores = sigs.map(s=>s.score).filter(s=>s!=="gray");
    // Must have no red or yellow signals
    if (scores.some(s=>s==="red"||s==="yellow")) return false;
    // Must have at least one legend signal
    return scores.some(s=>s==="legend");
  };

  // ── Per-CSM signal builder ──────────────────────────────────────────────
  const buildSignals = csm => {
    const signals = [];
    const det = getDetFn(csm.name)||{};
    const skippedForCSM = skippedCSMs.find(s=>s.name===csm.name);
    const liveAccts = Object.entries(csm.liveAccounts||{});
    const overdueAccts = liveAccts.filter(([,tasks])=>tasks.some(t=>t.ov)).map(([n])=>n);

    // Trending data from history snapshots
    const csmHistory = (history||[]).filter(h=>h.name===csm.name||norm(h.name)===csm.name).sort((a,b)=>a.date>b.date?1:-1);
    const latestSnap = csmHistory[csmHistory.length-1];
    const prevSnap   = csmHistory[csmHistory.length-2];
    const trendArrow = (cur, prev) => {
      if (cur==null||prev==null) return "";
      const d = (cur-prev)*100;
      if (Math.abs(d)<0.5) return " (stable)";
      return d>0 ? " ↑"+d.toFixed(1)+"pp vs last wk" : " ↓"+Math.abs(d).toFixed(1)+"pp vs last wk";
    };

    // ── CADENCE ──
    // "Nothing due" = neutral — don't penalize
    const hasDueTasks = csm.dueCount > 0;
    const hasOverdue  = csm.overdueCount > 0;
    const hasSkipped  = skippedForCSM && skippedForCSM.skippedCount > 0;
    const has4th      = skippedForCSM && skippedForCSM.skippedFourthCount > 0;

    // On-time % from latest snapshot — always factors in
    const otPct = latestSnap?.otPct != null ? latestSnap.otPct : null;
    const otScore = otPct==null ? null : otPct>=0.95?"legend":otPct>=0.85?"green":otPct>=0.70?"yellow":"red";

    let cadScore, cadValue;
    if (!hasDueTasks && !hasOverdue && !hasSkipped) {
      // Nothing due yesterday — but still check on-time trend
      if (otScore==="red")    { cadScore="red";    cadValue="No tasks due · on-time trending "+pp(otPct)+" — needs attention"; }
      else if (otScore==="yellow") { cadScore="yellow"; cadValue="No tasks due · on-time trending "+pp(otPct); }
      else if (otScore==="legend") { cadScore="legend"; cadValue="No tasks due · on-time "+pp(otPct)+" — excellent"; }
      else                     { cadScore="green";  cadValue="No tasks due — neutral"+(otPct!=null?" · "+pp(otPct)+" on-time":""); }
    } else if (hasOverdue || hasSkipped) {
      // Has issues — use the worse of daily flags vs on-time trend
      const dailyScore = (has4th || csm.overdueCount >= 3) ? "red" : "yellow";
      const scores = [dailyScore, otScore].filter(Boolean);
      cadScore = scores.includes("red") ? "red" : "yellow";
      cadValue = [
        hasDueTasks ? csm.dueCount+" tasks due" : null,
        hasOverdue  ? csm.overdueCount+" overdue" : null,
        hasSkipped  ? skippedForCSM.skippedCount+" skipped yesterday" : null,
        otPct!=null ? pp(otPct)+" on-time" : null,
      ].filter(Boolean).join(" · ");
    } else {
      // Tasks due, all clear — check on-time trend
      if (otScore==="legend")       { cadScore="legend"; cadValue=csm.dueCount+" tasks due, all clear · "+pp(otPct)+" on-time"; }
      else if (otScore==="yellow")  { cadScore="yellow"; cadValue=csm.dueCount+" tasks due, all clear · on-time "+pp(otPct)+" — needs improvement"; }
      else if (otScore==="red")     { cadScore="red";    cadValue=csm.dueCount+" tasks due, all clear · on-time "+pp(otPct)+" — critical"; }
      else                          { cadScore="green";  cadValue=csm.dueCount+" tasks due, all clear"+(otPct!=null?" · "+pp(otPct)+" on-time":""); }
    }

    const cadDetail = [
      ...overdueAccts.slice(0,4).map(n=>({name:n, note:"Overdue", score:"red"})),
      ...(skippedForCSM?.skippedAccts||[]).slice(0,3).map(a=>({
        name:a.n, note:"Skipped"+(a.is4th?" 🚩 4th reschedule":""), score:a.is4th?"red":"yellow"
      })),
    ];
    // Trending context
    if (latestSnap?.otPct!=null) cadDetail.push({
      name:"On-time trending",
      note: pp(latestSnap.otPct)+trendArrow(latestSnap.otPct, prevSnap?.otPct),
      score:"green", isTrend:true
    });
    if (latestSnap?.cadPct!=null) cadDetail.push({
      name:"Cadence completion trending",
      note: pp(latestSnap.cadPct)+trendArrow(latestSnap.cadPct, prevSnap?.cadPct),
      score:"green", isTrend:true
    });
    signals.push({key:"cad", label:"Cadence", score:cadScore, value:cadValue, detail:cadDetail});

    // ── BOB RETENTION ──
    // Use live bobRaw if available for accurate retention
    const liveKey = bobRaw&&bobRaw.bob ? Object.keys(bobRaw.bob).find(k=>norm(k)===csm.name||k===csm.name) : null;
    const liveBobEntry = liveKey ? bobRaw.bob[liveKey] : null;
    const adjK = bobAdj ? Object.keys(bobAdj).find(k=>norm(k)===csm.name||k===csm.name) : null;
    const lcmD = adjK ? bobAdj[adjK].lcmDelta : 0;
    const liveBoq = liveBobEntry ? liveBobEntry.boq : csm.bobBoq;
    const liveLcm = liveBobEntry ? (liveBobEntry.lcm||0)+lcmD : (csm.bobLcm||0)+lcmD;
    const liveRet = liveBoq>0 ? liveLcm/liveBoq : csm.bobRet;
    const retScore = liveRet==null?"gray":liveRet>=0.99?"legend":liveRet>=0.91?"green":liveRet>=0.85?"yellow":"red";
    if (liveBoq>0) signals.push({
      key:"bob", label:"Retention",
      score: retScore,
      value: liveRet!=null ? pp(liveRet)+(retScore==="legend"?" — exceptional":retScore==="green"?" — above 91% goal":retScore==="yellow"?" — near goal":" — below 85%") : "No BOB data",
      detail: [],
    });

    // ── REVENUE ──
    // Legend: MRR added AND positive net billing (actually growing BOB)
    const revScore = csm.rev>0&&(csm.bobNet||0)>0?"legend":csm.rev>0?"green":csm.bobNet<0?"red":"yellow";
    const revDetail = [];
    if (csm.accts) csm.accts.slice(0,3).forEach(a=>{
      if (a.m>0) revDetail.push({name:a.b, note:"MRR "+fd(a.m), score:"green"});
    });
    // QTD BOB context — use live bobRaw to match BOB card
    if (liveBoq>0) revDetail.push({
      name:"Quarter BOB",
      note: fd(liveBoq)+" BOQ → "+fd(liveLcm)+" current"+(liveRet!=null?" ("+pp(liveRet)+" retention)":""),
      score: liveRet>=0.91?"green":liveRet>=0.85?"yellow":"red",
      isTrend:true,
    });
    (det.d||[]).slice(0,2).forEach(r=>revDetail.push({name:r.a||r.e, note:fd(r.n)+" "+r.l, score:"red"}));
    signals.push({
      key:"rev", label:"Revenue",
      score: revScore,
      value: revScore==="legend" ? fd(csm.rev)+" MRR added QTD · BOB growing" : csm.rev>0 ? fd(csm.rev)+" MRR added QTD" : liveBoq>0 ? "No revenue QTD · BOB: "+fd(liveLcm) : "No revenue data",
      detail: revDetail,
    });

    // ── CALLS ──
    const callWeeks = Object.keys(callData).length>0
      ? [...new Set(Object.values(callData).flatMap(c=>Object.keys(c)))].sort().slice(-4) : [];
    const csmCallKey = Object.keys(callData).find(k=>norm(k)===csm.name||k===csm.name);
    let callScore="gray", callValue="No call data";
    const callDetail=[];
    if (csmCallKey) {
      let comp=0,ns=0,can=0;
      callWeeks.forEach(w=>{
        Object.values(callData[csmCallKey][w]||{}).forEach(d=>{comp+=d.completed;ns+=d.noShow;can+=d.cancelled||0;});
      });
      const total=comp+ns+can;
      const nsRate=comp+ns>0?ns/(comp+ns):0;
      // Legend: 0 no-shows AND 0 cancelled (perfect period)
      callScore = total===0?"gray":ns===0&&can===0?"legend":nsRate<=0.08?"green":nsRate<=0.15?"yellow":"red";
      callValue = callScore==="legend" ? comp+" completed · 0 no-shows · 0 cancelled — perfect!" : total>0 ? comp+" completed · "+ns+" no-show · "+can+" cancelled" : "No calls this period";
      callWeeks.forEach(w=>{
        Object.entries(callData[csmCallKey][w]||{}).forEach(([svc,d])=>{
          if (d.noShow>0) callDetail.push({name:svc, note:d.noShow+" no-show", score:"red"});
          if ((d.cancelled||0)>0) callDetail.push({name:svc, note:d.cancelled+" cancelled", score:"yellow"});
        });
      });
    }
    signals.push({key:"calls", label:"Calls", score:callScore, value:callValue, detail:callDetail});

    // ── QA ──
    const qaMcMonths=Object.keys(qamc).sort();
    const qaSsMonths=Object.keys(qass).sort();
    const latestMc=qaMcMonths.length>0?qamc[qaMcMonths[qaMcMonths.length-1]]:null;
    const latestSs=qaSsMonths.length>0?qass[qaSsMonths[qaSsMonths.length-1]]:null;
    const mcEntry=latestMc&&(latestMc[csm.name]||latestMc[norm(csm.name)]);
    const ssEntry=latestSs&&(latestSs[csm.name]||latestSs[norm(csm.name)]);
    let qaScore="gray", qaValue="No QA data";
    const qaDetail=[];
    if (mcEntry||ssEntry) {
      const scores=[mcEntry?.total,ssEntry?.total].filter(v=>v!=null);
      const avg=scores.length?scores.reduce((s,v)=>s+v,0)/scores.length:null;
      // Legend: both MC and SS >= 95% (if both exist), or single report >= 95%
      const allAbove95 = scores.length>0 && scores.every(s=>s>=0.95);
      qaScore=avg==null?"gray":allAbove95?"legend":avg>=0.93?"green":avg>=0.80?"yellow":"red";
      qaValue=[mcEntry?"MC: "+pp(mcEntry.total):null,ssEntry?"S&S: "+pp(ssEntry.total):null].filter(Boolean).join(" · ")||"No QA data";
      if (mcEntry?.criteria) Object.entries(mcEntry.criteria).forEach(([k,v])=>{if(v!=null&&v<0.80)qaDetail.push({name:k.replace(/_/g," "),note:pp(v)+" MC",score:"red"});});
      if (ssEntry?.criteria) Object.entries(ssEntry.criteria).forEach(([k,v])=>{if(v!=null&&v<0.80)qaDetail.push({name:k.replace(/_/g," "),note:pp(v)+" S&S",score:"red"});});
    }
    signals.push({key:"qa",label:"QA Score",score:qaScore,value:qaValue,detail:qaDetail});

    // ── CROSS-REFERENCE ──
    const crossRefs=[];
    const churnedAccts=(csm.churnedAccts||[]).map(a=>a.name.toLowerCase());
    (skippedForCSM?.skippedAccts||[]).forEach(a=>{
      const hasDec=(det.d||[]).find(r=>(r.a||"").toLowerCase().includes(a.n.toLowerCase().slice(0,8)));
      if (hasDec) crossRefs.push({name:a.n, note:"Skipped + billing decrease "+fd(hasDec.n), score:"red"});
      if (churnedAccts.some(c=>c.includes(a.n.toLowerCase().slice(0,6))))
        crossRefs.push({name:a.n, note:"Skipped cadence + account churned", score:"red"});
    });
    overdueAccts.slice(0,3).forEach(n=>{
      const hasDec=(det.d||[]).find(r=>(r.a||"").toLowerCase().includes(n.toLowerCase().slice(0,6)));
      if (hasDec) crossRefs.push({name:n, note:"Overdue + billing decrease "+fd(hasDec.n), score:"red"});
    });
    (det.i||[]).slice(0,2).forEach(r=>{
      if (r.n>200) crossRefs.push({name:r.a||r.e, note:"MRR increase +"+fd(r.n)+" "+r.l, score:"green"});
    });
    if (crossRefs.length>0) signals.push({
      key:"xref", label:"Cross-signals",
      score: worstScore(crossRefs.map(c=>c.score)),
      value: crossRefs.length+" account"+(crossRefs.length>1?"s":"")+" flagged",
      detail: crossRefs,
    });

    return signals;
  };

  // ── Visible CSMs ───────────────────────────────────────────────────────
  const visibleCSMs = csms.filter(c=>{
    const i=lk(c.name);
    if (filterCoach&&(i&&i.c||c.coach)!==filterCoach) return false;
    if (filterCSM&&c.name!==filterCSM) return false;
    return true;
  });

  // ── AI prompt builder ─────────────────────────────────────────────────
  const runDigestAI = async () => {
    const lines = ["=== DAILY DIGEST — THRYV CSM TEAM ===",
      "Period: "+PERIODS.find(p=>p.k===period)?.l,""];
    visibleCSMs.forEach(csm=>{
      const sigs = buildSignals(csm);
      const overall = worstScore(sigs.map(s=>s.score).filter(s=>s!=="gray"));
      lines.push(scoreDot(overall)+" "+csm.name);
      sigs.forEach(s=>{
        if (s.score!=="gray") lines.push("  "+scoreEmoji(s.score)+" "+s.label+": "+s.value);
        (s.detail||[]).filter(d=>d.score!=="green").forEach(d=>lines.push("    → "+d.name+": "+d.note));
      });
      lines.push("");
    });
    const prompt = [
      "You are an expert CSM coaching advisor at Thryv. Review this daily digest and provide:",
      "1. 🏆 TEAM WINS — what went well across the team?",
      "2. 🚨 URGENT — what needs immediate attention today?",
      "3. 📈 OPPORTUNITIES — where is there room to improve?",
      "4. 🎯 COACHING ACTIONS — 3 specific things to do right now",
      "Be specific — cite actual names. Under 400 words.",
      "",
      lines.join("\n"),
    ].join("\n");
    try { await navigator.clipboard.writeText(prompt); }
    catch(e) { const el=document.createElement("textarea");el.value=prompt;document.body.appendChild(el);el.select();document.execCommand("copy");document.body.removeChild(el); }
    setAiCopied(true);
    setTimeout(()=>setAiCopied(false),4000);
  };

  // ── Team rollup ────────────────────────────────────────────────────────
  const teamSignals = () => {
    if (visibleCSMs.length === 0) return [];
    const cats = ["rev","ret","cad","calls","qa"];
    return cats.map(cat=>{
      const scores = visibleCSMs.map(c=>buildSignals(c).find(s=>s.key===cat)?.score).filter(s=>s&&s!=="gray");
      const reds   = scores.filter(s=>s==="red").length;
      const yellows= scores.filter(s=>s==="yellow").length;
      const greens = scores.filter(s=>s==="green").length;
      const overall= reds>0&&reds>greens?"red":yellows>greens?"yellow":greens>0?"green":"gray";
      const labels = {rev:"Revenue",ret:"Retention",cad:"Cadence",calls:"Calls",qa:"QA"};
      return {key:cat, label:labels[cat], score:overall,
        summary: greens+"/"+scores.length+" green"+(reds>0?", "+reds+" red":"")};
    });
  };

  return (
    <div style={{padding:"0 0 40px"}}>
      {/* Header */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:20,flexWrap:"wrap",gap:12}}>
        <div>
          <div style={{fontSize:18,fontWeight:700,color:"#29355D"}}>📋 Daily Digest</div>
          <div style={{fontSize:12,color:"#808080",marginTop:2}}>
            {filterCSM?dispName(filterCSM):filterCoach?COACHES.find(c=>c.e===filterCoach)?.n:"Full team"} · {PERIODS.find(p=>p.k===period)?.l}
          </div>
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          {/* Period pills */}
          {PERIODS.map(p=>(
            <button key={p.k} onClick={()=>setPeriod(p.k)}
              style={{padding:"5px 12px",borderRadius:20,fontSize:11,fontWeight:600,cursor:"pointer",
                border:"0.5px solid "+(period===p.k?"#29355D":"rgba(41,53,93,.15)"),
                background:period===p.k?"#29355D":"#fff",
                color:period===p.k?"#fff":"#808080"}}>
              {p.l}
            </button>
          ))}
          {/* AI button */}
          <button onClick={runDigestAI}
            style={{display:"flex",alignItems:"center",gap:6,padding:"7px 14px",borderRadius:8,
              border:"none",background:aiCopied?"#16a34a":"#FF5000",color:"#fff",
              fontSize:12,fontWeight:600,cursor:"pointer",transition:"background .3s"}}>
            {aiCopied?"✓ Copied!":"🤖 Ask AI"}
          </button>
        </div>
      </div>
      {/* Score filters */}
      {!filterCSM&&<div style={{display:"flex",gap:8,marginBottom:20,flexWrap:"wrap",alignItems:"center"}}>
        {/* All CSMs pill */}
        <button onClick={()=>setScoreFilter("all")}
          style={{padding:"0 20px",height:64,borderRadius:10,fontSize:13,fontWeight:700,cursor:"pointer",
            border:"2px solid "+(scoreFilter==="all"?"#29355D":"rgba(41,53,93,.15)"),
            background:scoreFilter==="all"?"#29355D":"#fff",
            color:scoreFilter==="all"?"#fff":"#808080",transition:"all .15s",flexShrink:0}}>
          All CSMs
        </button>
        {/* Image buttons */}
        {[
          {k:"legend", img:imgLegend, label:"Legend Status"},
          {k:"green",  img:imgCrushingIt},
          {k:"yellow", img:imgAlmostThere},
          {k:"red",    img:imgNeedsLove},
        ].map(f=>(
          <button key={f.k} onClick={()=>setScoreFilter(scoreFilter===f.k?"all":f.k)}
            style={{padding:0,border:scoreFilter===f.k?"2px solid "+( f.k==="legend"?"#7c3aed":f.k==="green"?"#16a34a":f.k==="yellow"?"#d97706":"#dc2626"):"2px solid transparent",
              background:"transparent",cursor:"pointer",borderRadius:10,overflow:"hidden",
              opacity:scoreFilter!=="all"&&scoreFilter!==f.k?0.45:1,
              transform:scoreFilter===f.k?"scale(1.05)":"scale(1)",
              boxShadow:scoreFilter===f.k?"0 4px 16px rgba(0,0,0,.2)":"0 1px 4px rgba(0,0,0,.08)",
              transition:"all .2s",flexShrink:0}}>
            <img src={f.img} alt={f.label||f.k}
              style={{display:"block",height:64,width:280,objectFit:"cover",objectPosition:"center"}}/>
          </button>
        ))}
        {scoreFilter!=="all"&&<span style={{fontSize:12,fontWeight:600,color:"#808080",marginLeft:4}}>
          {scoreFilter==="legend"
            ? visibleCSMs.filter(c=>isLegend(c)).length
            : visibleCSMs.filter(c=>!isLegend(c)&&worstScore(buildSignals(c).map(s=>s.score).filter(s=>s!=="gray"))===scoreFilter).length
          } CSMs
        </span>}
      </div>}

      {aiCopied&&<div style={{marginBottom:16,padding:"10px 16px",borderRadius:10,background:"#29355D",
        color:"#fff",fontSize:12,display:"flex",gap:16,alignItems:"center",flexWrap:"wrap"}}>
        <span>📋 <strong>Digest copied!</strong></span>
        <span>Open <a href="https://claude.ai/new" target="_blank" rel="noreferrer" style={{color:"#FF5000",fontWeight:700}}>claude.ai/new</a></span>
        <span>→ Ctrl+V → Enter</span>
      </div>}

      {/* Team rollup (coach/manager view) */}
      {!filterCSM&&visibleCSMs.length>1&&(()=>{
        const ts = teamSignals();
        const overall = worstScore(ts.map(t=>t.score));
        return (
          <div style={{...S.card,marginBottom:20,borderLeft:"4px solid "+scoreColor(overall)}}>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
              <div style={{fontSize:22}}>{scoreDot(overall)}</div>
              <div>
                <div style={{fontSize:15,fontWeight:700,color:"#29355D"}}>
                  {overall==="legend"?"Team is legendary":overall==="green"?"Team is winning":overall==="yellow"?"Team needs attention":"Team has urgent issues"}
                </div>
                <div style={{fontSize:11,color:"#808080"}}>{visibleCSMs.length} CSMs · {PERIODS.find(p=>p.k===period)?.l}</div>
              </div>
            </div>
            <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
              {ts.map(t=>(
                <div key={t.key} style={{padding:"8px 12px",borderRadius:8,background:scoreBg(t.score),
                  border:"0.5px solid "+scoreColor(t.score)+"44",minWidth:100}}>
                  <div style={{fontSize:11,fontWeight:600,color:scoreColor(t.score)}}>{scoreDot(t.score)} {t.label}</div>
                  <div style={{fontSize:10,color:"#808080",marginTop:2}}>{t.summary}</div>
                </div>
              ))}
            </div>
          </div>
        );
      })()}

      {/* CSM tiles (coach view) or single CSM detail */}
      {!filterCSM&&visibleCSMs.length>1
        ? <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:12}}>
            {visibleCSMs.filter(csm=>{
              if (scoreFilter==="all") return true;
              if (scoreFilter==="legend") return isLegend(csm);
              const sigs = buildSignals(csm);
              const ov = worstScore(sigs.map(s=>s.score).filter(s=>s!=="gray"));
              return ov===scoreFilter && !isLegend(csm);
            }).map(csm=>{
              const sigs = buildSignals(csm);
              const leg = isLegend(csm);
              const overall = leg ? "legend" : worstScore(sigs.map(s=>s.score).filter(s=>s!=="gray"));
              const isExp = expanded===csm.name;
              const reds   = sigs.filter(s=>s.score==="red").length;
              const yellows= sigs.filter(s=>s.score==="yellow").length;
              return (
                <div key={csm.name}>
                  <div onClick={()=>setExpanded(isExp?null:csm.name)}
                    style={{...S.card,cursor:"pointer",
                      borderLeft:"4px solid "+scoreColor(overall),
                      background:overall==="legend"?"rgba(124,58,237,.04)":isExp?"rgba(41,53,93,.03)":"#fff",
                      transition:"all .15s"}}>
                    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
                      <div style={{fontWeight:600,fontSize:13,color:"#29355D"}}>{dispName(csm.name)}</div>
                      <div style={{fontSize:18}}>{scoreDot(overall)}</div>
                    </div>
                    <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:8}}>
                      {sigs.filter(s=>s.score!=="gray").map(s=>(
                        <span key={s.key} style={{fontSize:10,padding:"2px 7px",borderRadius:20,
                          background:scoreBg(s.score),color:scoreColor(s.score),fontWeight:500}}>
                          {s.label}
                        </span>
                      ))}
                    </div>
                    <div style={{fontSize:11,color:"#808080"}}>
                      {reds>0&&<span style={{color:"#dc2626",fontWeight:500}}>{reds} urgent</span>}
                      {reds>0&&yellows>0&&" · "}
                      {yellows>0&&<span style={{color:"#d97706"}}>{yellows} needs attention</span>}
                      {reds===0&&yellows===0&&<span style={{color:"#16a34a"}}>All green ✓</span>}
                    </div>
                    <div style={{fontSize:10,color:"#5378FC",marginTop:6}}>{isExp?"▲ Collapse":"▼ Expand"}</div>
                  </div>
                  {/* Inline expansion */}
                  {isExp&&<div style={{background:"#F4F6FB",borderRadius:"0 0 12px 12px",padding:16,marginTop:-8,
                    border:"0.5px solid rgba(41,53,93,.08)",borderTop:"none"}}>
                    {sigs.filter(s=>s.score!=="gray").map(s=>(
                      <div key={s.key} style={{marginBottom:12}}>
                        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
                          <span style={{fontSize:16}}>{scoreDot(s.score)}</span>
                          <span style={{fontSize:12,fontWeight:600,color:"#29355D"}}>{s.label}</span>
                          <span style={{fontSize:11,color:"#808080"}}>{s.value}</span>
                        </div>
                        {(s.detail||[]).map((d,i)=>(
                          <div key={i} style={{display:"flex",alignItems:"center",gap:6,padding:"4px 0 4px 24px",
                            borderBottom:"0.5px solid rgba(41,53,93,.06)"}}>
                            <span style={{fontSize:11,color:scoreColor(d.score),flexShrink:0}}>{scoreEmoji(d.score)}</span>
                            <span style={{fontSize:11,fontWeight:500,color:"#29355D"}}>{d.name}</span>
                            {d.note&&<span style={{fontSize:10,color:"#808080",marginLeft:"auto"}}>{d.note}</span>}
                          </div>
                        ))}
                      </div>
                    ))}
                  </div>}
                </div>
              );
            })}
          </div>
        : /* Single CSM full detail view */
          visibleCSMs.length>0&&(()=>{
            const csm = visibleCSMs[0];
            const sigs = buildSignals(csm);
            const leg = isLegend(csm);
            const overall = leg ? "legend" : worstScore(sigs.map(s=>s.score).filter(s=>s!=="gray"));
            const scoreImg = overall==="legend"?imgLegend:overall==="green"?imgCrushingIt:overall==="yellow"?imgAlmostThere:imgNeedsLove;
            const scoreMsg = overall==="legend"?"Leave some wins for the rest of us":overall==="green"?"Strong performance this period":overall==="yellow"?"Some areas need attention":"Immediate attention needed";

            // Build coaching notes
            const goodNotes = [], coachNotes = [];
            sigs.forEach(s => {
              if (s.score==="green"||s.score==="legend") {
                if (s.key==="rev"&&csm.rev>0)     goodNotes.push("Revenue added this period — "+fd(csm.rev)+" MRR");
                if (s.key==="cad"&&csm.dueCount>0) goodNotes.push("All cadence tasks completed on time");
                if (s.key==="calls")               goodNotes.push("Call no-show rate on target");
                if (s.key==="qa")                  goodNotes.push("QA scores above goal — "+s.value);
                if (s.key==="ret"&&csm.bobRet>=0.95) goodNotes.push("Exceptional retention at "+pp(csm.bobRet));
              }
              if (s.score==="red"||s.score==="yellow") {
                if (s.key==="cad"&&csm.overdueCount>0)    coachNotes.push(csm.overdueCount+" overdue cadence task"+(csm.overdueCount>1?"s":"")+" — review account engagement");
                if (s.key==="cad"&&(skippedCSMs.find(sk=>sk.name===csm.name)?.skippedCount||0)>0) {
                  const sk = skippedCSMs.find(s=>s.name===csm.name);
                  coachNotes.push(sk.skippedCount+" skipped cadence"+(sk.skippedFourthCount>0?" (includes 4th reschedule — escalate)":""));
                }
                if (s.key==="ret"&&csm.bobRet<0.91)       coachNotes.push("Retention below 91% goal at "+pp(csm.bobRet)+" — review billing changes");
                if (s.key==="calls") {
                  const csmCallKey=Object.keys(callData).find(k=>norm(k)===csm.name||k===csm.name);
                  if (csmCallKey) {
                    let ns=0; Object.values(callData[csmCallKey]||{}).forEach(w=>Object.values(w).forEach(d=>ns+=d.noShow));
                    if (ns>0) coachNotes.push(ns+" no-show call"+(ns>1?"s":"")+" this period — discuss follow-up process");
                  }
                }
                if (s.key==="qa")                          coachNotes.push("QA below goal — focus areas: "+(s.detail||[]).filter(d=>d.score==="red").map(d=>d.name).slice(0,3).join(", "));
                if (s.key==="rev"&&csm.rev===0)            coachNotes.push("No revenue submitted this period — review pipeline");
              }
            });
            // Cross-ref notes
            const xref = sigs.find(s=>s.key==="xref");
            if (xref) (xref.detail||[]).filter(d=>d.score==="red").forEach(d=>coachNotes.push(d.name+": "+d.note));

            return (
              <div>
                {/* Score image banner */}
                <div style={{marginBottom:16,borderRadius:12,overflow:"hidden",position:"relative"}}>
                  <img src={scoreImg} alt={overall}
                    style={{width:"100%",height:"auto",maxHeight:160,objectFit:"contain",objectPosition:"center",display:"block",background:overall==="legend"?"#3b1f6e":overall==="green"?"#d1fae5":overall==="yellow"?"#fef9c3":"#fee2e2"}}/>
                  <div style={{position:"absolute",bottom:0,left:0,right:0,background:"linear-gradient(transparent,rgba(0,0,0,.5))",
                    padding:"8px 16px",display:"flex",alignItems:"center",gap:10}}>
                    <div style={{fontSize:14,fontWeight:700,color:"#fff",textShadow:"0 1px 3px rgba(0,0,0,.5)"}}>{dispName(csm.name)}</div>
                    <div style={{fontSize:11,color:"rgba(255,255,255,.85)",textShadow:"0 1px 3px rgba(0,0,0,.5)"}}>{scoreMsg}</div>
                  </div>
                </div>

                {/* Coaching notes */}
                {(goodNotes.length>0||coachNotes.length>0)&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:16}}>
                  {goodNotes.length>0&&<div style={{background:"rgba(22,163,74,.06)",borderRadius:10,padding:"12px 14px",
                    border:"0.5px solid rgba(22,163,74,.2)"}}>
                    <div style={{fontSize:11,fontWeight:700,color:"#166534",textTransform:"uppercase",marginBottom:8,letterSpacing:".04em"}}>✓ What's Working</div>
                    {goodNotes.map((n,i)=><div key={i} style={{fontSize:12,color:"#166534",padding:"3px 0",borderBottom:i<goodNotes.length-1?"0.5px solid rgba(22,163,74,.1)":"none"}}>{n}</div>)}
                  </div>}
                  {coachNotes.length>0&&<div style={{background:"rgba(220,38,38,.06)",borderRadius:10,padding:"12px 14px",
                    border:"0.5px solid rgba(220,38,38,.2)"}}>
                    <div style={{fontSize:11,fontWeight:700,color:"#991b1b",textTransform:"uppercase",marginBottom:8,letterSpacing:".04em"}}>⚡ Coaching Focus</div>
                    {coachNotes.map((n,i)=><div key={i} style={{fontSize:12,color:"#991b1b",padding:"3px 0",borderBottom:i<coachNotes.length-1?"0.5px solid rgba(220,38,38,.1)":"none"}}>{n}</div>)}
                  </div>}
                </div>}
                {/* BOB Summary Card — use bobRaw directly for accuracy */}
                {(()=>{
                  // Prefer bobRaw.bob (live sheet) over buildCSMs data (may use stale BOB_CSMS)
                  const liveKey = bobRaw&&bobRaw.bob ? Object.keys(bobRaw.bob).find(k=>norm(k)===csm.name||k===csm.name) : null;
                  const liveBob = liveKey ? bobRaw.bob[liveKey] : null;
                  const rawBoq = liveBob ? liveBob.boq : csm.bobBoq;
                  const rawLcm = liveBob ? liveBob.lcm : csm.bobLcm;
                  const rawNet = liveBob ? liveBob.net : csm.bobNet;
                  if (!rawBoq || rawBoq===0) return null;
                  const adjKey = bobAdj ? Object.keys(bobAdj).find(k=>norm(k)===csm.name||k===csm.name) : null;
                  const lcmDelta = adjKey ? bobAdj[adjKey].lcmDelta : 0;
                  const adjLcm = (rawLcm||0) + lcmDelta;
                  const adjNet = (rawNet||0) + lcmDelta;
                  const adjRet = rawBoq>0 ? adjLcm/rawBoq : null;
                  const csm_bobBoq = rawBoq;
                  return <div style={{...S.card,marginBottom:12,borderLeft:"3px solid "+(adjRet!=null?(adjRet>=0.91?"#16a34a":adjRet>=0.88?"#d97706":"#dc2626"):"#808080")}}>
                  <div style={{fontSize:13,fontWeight:600,color:"#29355D",marginBottom:10}}>📊 Book of Business — This Quarter</div>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:csm.churnedAccts?.length>0?12:0}}>
                    {[
                      {l:"BOQ",         v:fd(csm_bobBoq), col:"#29355D"},
                      {l:"Current",     v:fd(adjLcm),     col:"#5378FC"},
                      {l:"Net Change",  v:fd(adjNet),     col:adjNet>=0?"#16a34a":"#dc2626"},
                      {l:"Retention",   v:adjRet!=null?pp(adjRet):"—", col:adjRet>=0.91?"#16a34a":adjRet>=0.88?"#d97706":"#dc2626"},
                    ].map(t=>(
                      <div key={t.l} style={{background:"#F4F6FB",borderRadius:8,padding:"10px 12px",textAlign:"center"}}>
                        <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:4}}>{t.l}</div>
                        <div style={{fontSize:16,fontWeight:700,color:t.col}}>{t.v}</div>
                      </div>
                    ))}
                  </div>
                  {/* Churned count summary only — no account details on digest */}
                  {csm.churnedAccts?.length>0&&<div style={{marginTop:8,fontSize:12,color:"#991b1b",fontWeight:500}}>
                    {csm.churnedAccts.length} churned account{csm.churnedAccts.length>1?"s":""} this quarter
                  </div>}
                </div>;
                })()}

                {sigs.filter(s=>s.score!=="gray").map(s=>(
                  <div key={s.key} style={{...S.card,marginBottom:12,borderLeft:"3px solid "+scoreColor(s.score)}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:(s.detail&&s.detail.length>0)?10:0}}>
                      <span style={{fontSize:18}}>{scoreDot(s.score)}</span>
                      <span style={{fontSize:13,fontWeight:600,color:"#29355D"}}>{s.label}</span>
                      <span style={{fontSize:12,color:"#808080",marginLeft:4}}>{s.value}</span>
                    </div>
                    {(s.detail||[]).map((d,i)=>(
                      <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 0",
                        borderBottom:"0.5px solid rgba(41,53,93,.06)",opacity:d.isTrend?0.8:1}}>
                        <span style={{fontSize:12,color:d.isTrend?"#5378FC":scoreColor(d.score),flexShrink:0,width:16}}>{d.isTrend?"📈":scoreEmoji(d.score)}</span>
                        <span style={{fontSize:12,fontWeight:d.isTrend?400:500,color:d.isTrend?"#5378FC":"#29355D",flex:1,fontStyle:d.isTrend?"italic":"normal"}}>{d.name}</span>
                        {d.note&&<span style={{fontSize:11,color:d.isTrend?"#5378FC":"#808080",textAlign:"right"}}>{d.note}</span>}
                      </div>
                    ))}
                  </div>
                ))}
              </div>
            );
          })()
      }
    </div>
  );
}

// ── REVENUE VIEW ────────────────────────────────────────────────────────────
function RevenueView({rawRev, csms, filterCoach, filterCSM, managerCoaches}) {
  const [lbSort, setLbSort] = useState({col:"total", dir:"desc"});
  const [regionFilter, setRegionFilter] = useState("all");
  const [quarterFilter, setQuarterFilter] = useState("all");
  const availableQuarters = [...new Set((rawRev||[]).map(r=>(r["Quarter for Consideration"]||r["Quarter"]||"").trim()).filter(Boolean))].sort();
  const currentYear = new Date().getFullYear();
  const inQuarterFilter = (qtr) => {
    if (quarterFilter === "all") return true;
    if (quarterFilter === "ytd") return qtr.includes(String(currentYear));
    return qtr === quarterFilter;
  };


  // Parse raw rows into enriched objects
  const rows = (rawRev||[]).map(r => {
    const csm  = norm(r["CSM Name"]||r["csm_name"]||"");
    const team = r["CSM Team!"]||r["CSM Team! "]||r["team"]||"";
    const tier = r["CSM Tier"]||r["csm_tier"]||"";
    const mrr  = parseFloat(String(r["MRR $ Added"]||r["MRR $"]||r["MRR"]||0).replace(/[$,]/g,""))||0;
    const otr  = parseFloat(String(r["OTR $ Added"]||r["OTR $"]||r["OTR"]||0).replace(/[$,]/g,""))||0;
    const tot  = parseFloat(String(r["Total Revenue Added"]||r["Total Revenue"]||r["Revenue"]||0).replace(/[$,]/g,""))||0;
    const nr   = (r["Non-Revenue Integrations"]||"").trim();
    const mrrInt = (r["MRR Integration"]||"").trim();
    const biz  = (r["Business Name"]||"").trim();
    const type = (r["Type of Integration"]||"").trim();
    const qtr  = (r["Quarter for Consideration"]||r["Quarter"]||"").trim();
    const i    = lk(csm);
    return {csm, team: (i&&i.t)||team, tier:(i&&i.r)||tier, region:region(csm)||"", mrr, otr, tot, nr, mrrInt, biz, type, qtr};
  }).filter(r=>r.csm && isValidCSM(r.csm));

  // Apply manager + coach + CSM + region + quarter filter
  const filtered = rows.filter(r=>{
    if (!inQuarterFilter(r.qtr)) return false;
    const i = lk(r.csm);
    if (managerCoaches && !managerCoaches.includes(i&&i.c)) return false;
    if (filterCoach && !(i&&i.c===filterCoach)) return false;
    if (filterCSM && r.csm !== filterCSM) return false;
    if (regionFilter!=="all" && r.region!==regionFilter) return false;
    return true;
  });

  // ── Totals ────────────────────────────────────────────────────────────────
  const totalMRR  = filtered.reduce((s,r)=>s+r.mrr,0);
  const totalOTR  = filtered.reduce((s,r)=>s+r.otr,0);
  const totalRev  = filtered.reduce((s,r)=>s+r.tot,0);
  const totalSubs = filtered.length;
  const nonRevSubs = filtered.filter(r=>r.type==="Non-Revenue").length;
  const mrrSubs   = filtered.filter(r=>r.type==="Monthly Recurring Revenue").length;
  const otrSubs   = filtered.filter(r=>r.type==="One-Time Revenue").length;
  const activeCsms = new Set(filtered.filter(r=>r.tot>0||r.nr).map(r=>r.csm)).size;

  // ── By team ───────────────────────────────────────────────────────────────
  const byTeam = {};
  filtered.forEach(r=>{
    if(!r.team) return;
    if(!byTeam[r.team]) byTeam[r.team]={mrr:0,otr:0,total:0,subs:0};
    byTeam[r.team].mrr+=r.mrr; byTeam[r.team].otr+=r.otr;
    byTeam[r.team].total+=r.tot; byTeam[r.team].subs++;
  });
  const teamRows = Object.entries(byTeam).sort((a,b)=>b[1].total-a[1].total);
  const maxTeamRev = teamRows[0]?teamRows[0][1].total:1;

  // ── By tier ───────────────────────────────────────────────────────────────
  const byTier = {};
  filtered.forEach(r=>{
    if(!r.tier) return;
    if(!byTier[r.tier]) byTier[r.tier]={mrr:0,otr:0,total:0,subs:0};
    byTier[r.tier].mrr+=r.mrr; byTier[r.tier].otr+=r.otr;
    byTier[r.tier].total+=r.tot; byTier[r.tier].subs++;
  });
  const tierOrder = ["CSMI","CSMII","CSMIII","SSMI","SSMII"];
  const tierLabel = {"CSMI":"CSM 1","CSMII":"CSM 2","CSMIII":"CSM 3","SSMI":"SSM 1","SSMII":"SSM 2"};

  // ── Top MRR integration types ─────────────────────────────────────────────
  const mrrTypes = {};
  filtered.filter(r=>r.mrr>0).forEach(r=>{
    const k = r.mrrInt||"Unspecified";
    if(!mrrTypes[k]) mrrTypes[k]={count:0,amount:0};
    mrrTypes[k].count++; mrrTypes[k].amount+=r.mrr;
  });
  const mrrTypeRows = Object.entries(mrrTypes).sort((a,b)=>b[1].amount-a[1].amount);
  const maxMrrAmt = mrrTypeRows[0]?mrrTypeRows[0][1].amount:1;

  // ── Non-revenue types ─────────────────────────────────────────────────────
  const nrTypes = {};
  filtered.filter(r=>r.nr).forEach(r=>{
    if(!nrTypes[r.nr]) nrTypes[r.nr]=0;
    nrTypes[r.nr]++;
  });
  const NR_COLORS = {
    "Yelp Request a Quote":    "#E31C24",
    "Tracking Line Provisioned":"#4A5D8C",
    "Demo Booked":             "#16a34a",
    "ThryvPay":                "#FF5000",
    "Webchat":                 "#5378FC",
  };

  // ── Integration type donut data ───────────────────────────────────────────
  const donutData = [
    {label:"MRR",    val:mrrSubs,   col:"#FF5000"},
    {label:"OTR",    val:otrSubs,   col:"#5378FC"},
    {label:"Non-Rev",val:nonRevSubs,col:"#ECEEF1"},
  ];
  const donutTotal = donutData.reduce((s,d)=>s+d.val,0)||1;
  // Build SVG donut arcs
  function donutArcs(data, total, cx, cy, r, stroke) {
    let angle = -Math.PI/2;
    return data.map(d=>{
      const slice = (d.val/total)*2*Math.PI;
      const x1=cx+r*Math.cos(angle), y1=cy+r*Math.sin(angle);
      angle+=slice;
      const x2=cx+r*Math.cos(angle), y2=cy+r*Math.sin(angle);
      const large=slice>Math.PI?1:0;
      return {path:`M${cx},${cy} L${x1.toFixed(1)},${y1.toFixed(1)} A${r},${r} 0 ${large},1 ${x2.toFixed(1)},${y2.toFixed(1)} Z`, col:d.col, label:d.label, val:d.val};
    });
  }
  const arcs = donutArcs(donutData, donutTotal, 65, 65, 55, 8);

  // ── CSM leaderboard ───────────────────────────────────────────────────────
  const byCsm = {};
  filtered.forEach(r=>{
    if(!r.csm) return;
    if(!byCsm[r.csm]) byCsm[r.csm]={csm:r.csm,team:r.team,tier:r.tier,region:r.region,mrr:0,otr:0,total:0,subs:0,revPerSub:0};
    byCsm[r.csm].mrr+=r.mrr; byCsm[r.csm].otr+=r.otr;
    byCsm[r.csm].total+=r.tot; byCsm[r.csm].subs++;
  });
  Object.values(byCsm).forEach(c=>{ c.revPerSub = c.subs>0?c.total/c.subs:0; });
  const lbRows = Object.values(byCsm).sort((a,b)=>{
    const av=a[lbSort.col]||0, bv=b[lbSort.col]||0;
    return lbSort.dir==="desc"?bv-av:av-bv;
  });
  const [showAllLb, setShowAllLb] = useState(false);
  const visibleLb = showAllLb ? lbRows : lbRows.slice(0,15);
  const medals=["🥇","🥈","🥉"];
  const thS = {fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,padding:"0 0 8px",textAlign:"left",borderBottom:"0.5px solid rgba(41,53,93,.08)",cursor:"pointer",whiteSpace:"nowrap"};
  const thRS = {...thS, textAlign:"right"};
  const tdS = {padding:"9px 8px 9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",fontSize:12};
  const tdRS = {...tdS, textAlign:"right"};
  function thSort(col,lbl) {
    return <th style={{...thRS,color:lbSort.col===col?"#FF5000":"#808080"}}
      onClick={()=>setLbSort(s=>({col,dir:s.col===col&&s.dir==="desc"?"asc":"desc"}))}>
      {lbl}{lbSort.col===col?(lbSort.dir==="desc"?" ▼":" ▲"):<span style={{color:"#ccc",fontSize:9}}> ↕</span>}
    </th>;
  }

  const cardStyle = {background:"#fff",border:"0.5px solid rgba(41,53,93,.09)",borderRadius:12,padding:16};
  const secTitle = {fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12};

  return (
    <div>
      {/* ── Quarter filter bar ── */}
      <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:16,flexWrap:"wrap"}}>
        <span style={{fontSize:11,fontWeight:600,color:"#808080",textTransform:"uppercase",letterSpacing:".05em",marginRight:4}}>Quarter:</span>
        {[["all","All time"], ["ytd","YTD "+new Date().getFullYear()], ...availableQuarters.map(q=>[q,q])].map(([v,l])=>(
          <button key={v} onClick={()=>setQuarterFilter(v)}
            style={{padding:"4px 12px",borderRadius:20,border:"0.5px solid "+(quarterFilter===v?"#29355D":"rgba(41,53,93,.15)"),
              background:quarterFilter===v?"#29355D":"#fff",color:quarterFilter===v?"#fff":"#808080",
              fontSize:11,fontWeight:500,cursor:"pointer",transition:"all .15s"}}>
            {l}
          </button>
        ))}
        {quarterFilter!=="all"&&<span style={{fontSize:11,color:"#808080",marginLeft:4}}>
          · filtering revenue rows by "Quarter for Consideration"
        </span>}
      </div>

      {/* ── Top metric tiles ── */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(5,minmax(0,1fr))",gap:12,marginBottom:20}}>
        {[
          {l:"Total Revenue Added", v:fd(totalRev),     s:`${totalSubs} submissions shown`,    col:"#FF5000"},
          {l:"MRR Added",           v:fd(totalMRR),     s:`${Math.round(totalMRR/totalRev*100)||0}% of total · ${mrrSubs} submissions`, col:"#29355D"},
          {l:"One-Time Revenue",    v:fd(totalOTR),     s:`${otrSubs} one-time submissions`,   col:"#5378FC"},
          {l:"Submissions",         v:totalSubs,        s:`MRR ${mrrSubs} · OTR ${otrSubs} · Non-rev ${nonRevSubs}`, col:"#d97706"},
          {l:"Active CSMs",         v:activeCsms,       s:`${lbRows.filter(c=>c.total>0).length} generating revenue`, col:"#16a34a"},
        ].map(m=>(
          <div key={m.l} style={{...cardStyle,position:"relative",overflow:"hidden"}}>
            <div style={{position:"absolute",top:0,left:0,right:0,height:3,background:m.col,borderRadius:"12px 12px 0 0"}}/>
            <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:6}}>{m.l}</div>
            <div style={{fontSize:22,fontWeight:500,color:m.col,lineHeight:1,marginBottom:3}}>{m.v}</div>
            <div style={{fontSize:11,color:"#808080"}}>{m.s}</div>
          </div>
        ))}
      </div>

      {/* ── Row 2: By team + By tier + Integration split ── */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:16,marginBottom:16}}>

        {/* Revenue by team */}
        <div style={cardStyle}>
          <div style={secTitle}>Revenue by Team</div>
          {teamRows.map(([team,d])=>{
            const coach=COACHES.find(c=>c.t===team);
            const col=TEAM_COLS[team]||"#888";
            return <div key={team} style={{marginBottom:8}}>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:3}}>
                <span style={{display:"flex",alignItems:"center",gap:6}}>
                  <span style={{width:8,height:8,borderRadius:"50%",background:col,display:"inline-block",flexShrink:0}}/>
                  <span style={{fontWeight:500}}>{st(team)}</span>
                </span>
                <span style={{color:"#29355D",fontWeight:500}}>{fk(d.total)}</span>
              </div>
              <div style={{height:6,background:"#ECEEF1",borderRadius:3,overflow:"hidden"}}>
                <div style={{height:"100%",borderRadius:3,background:col,width:(d.total/maxTeamRev*100).toFixed(1)+"%"}}/>
              </div>
              <div style={{fontSize:10,color:"#808080",marginTop:2,display:"flex",gap:8}}>
                <span>MRR {fk(d.mrr)}</span>
                {d.otr>0&&<span>OTR {fk(d.otr)}</span>}
                <span>{d.subs} subs</span>
              </div>
            </div>;
          })}
        </div>

        {/* Revenue by region */}
        <div style={cardStyle}>
          <div style={secTitle}>Revenue by Region</div>
          {["DR","US","ANZ"].map(reg=>{
            const regRows = rows.filter(r=>r.region===reg && (!managerCoaches||managerCoaches.includes(lk(r.csm)&&lk(r.csm).c)) && (!filterCoach||(lk(r.csm)&&lk(r.csm).c===filterCoach)) && (!filterCSM||r.csm===filterCSM));
            const regMRR = regRows.reduce((s,r)=>s+r.mrr,0);
            const regOTR = regRows.reduce((s,r)=>s+r.otr,0);
            const regTot = regRows.reduce((s,r)=>s+r.tot,0);
            const regSubs = regRows.length;
            const maxRegTot = Math.max(...["DR","US","ANZ"].map(r=>rows.filter(x=>x.region===r).reduce((s,x)=>s+x.tot,0)))||1;
            const REG_COL = {DR:"#FF5000",US:"#29355D",ANZ:"#5378FC"};
            return <div key={reg} style={{marginBottom:12,cursor:"pointer",opacity:regionFilter===reg||regionFilter==="all"?1:0.4}} onClick={()=>setRegionFilter(r=>r===reg?"all":reg)}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:11,marginBottom:4}}>
                <span style={{display:"flex",alignItems:"center",gap:8}}>
                  <span style={{fontSize:13,fontWeight:700,color:REG_COL[reg],width:32}}>{reg}</span>
                  <span style={{fontWeight:500,color:"#29355D"}}>{regSubs} subs</span>
                  {regionFilter===reg&&<span style={{fontSize:10,padding:"1px 7px",borderRadius:20,background:REG_COL[reg],color:"#fff"}}>filtered</span>}
                </span>
                <span style={{fontWeight:600,color:REG_COL[reg]}}>{fk(regTot)}</span>
              </div>
              <div style={{height:8,background:"#ECEEF1",borderRadius:4,overflow:"hidden"}}>
                <div style={{height:"100%",borderRadius:4,background:REG_COL[reg],width:(regTot/maxRegTot*100).toFixed(1)+"%"}}/>
              </div>
              <div style={{fontSize:10,color:"#808080",marginTop:3,display:"flex",gap:10}}>
                <span>MRR {fk(regMRR)}</span>
                {regOTR>0&&<span>OTR {fk(regOTR)}</span>}
              </div>
            </div>;
          })}
          <div style={{height:"0.5px",background:"rgba(41,53,93,.08)",margin:"8px 0"}}/>
          <div style={{fontSize:11,color:"#808080",textAlign:"center"}}>Click a region to filter · click again to clear</div>
        </div>

        {/* Integration type split donut */}
        <div style={cardStyle}>
          <div style={secTitle}>Integration Type Split</div>
          <div style={{display:"flex",alignItems:"center",gap:16}}>
            <svg width={130} height={130} style={{flexShrink:0}}>
              {arcs.map((a,i)=>(
                a.val>0&&<path key={i} d={a.path} fill={a.col} stroke="#fff" strokeWidth={1.5}/>
              ))}
              <circle cx={65} cy={65} r={32} fill="#fff"/>
              <text x={65} y={61} textAnchor="middle" fontSize={11} fill="#808080">Total</text>
              <text x={65} y={77} textAnchor="middle" fontSize={14} fontWeight="500" fill="#29355D">{totalSubs}</text>
            </svg>
            <div style={{flex:1}}>
              {donutData.map(d=>(
                <div key={d.label} style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8,fontSize:12}}>
                  <span style={{display:"flex",alignItems:"center",gap:6}}>
                    <span style={{width:10,height:10,borderRadius:2,background:d.col,display:"inline-block"}}/>
                    {d.label}
                  </span>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontWeight:500}}>{d.val}</div>
                    <div style={{fontSize:10,color:"#808080"}}>{Math.round(d.val/donutTotal*100)}%</div>
                  </div>
                </div>
              ))}
              <div style={{height:"0.5px",background:"rgba(41,53,93,.08)",margin:"8px 0"}}/>
              <div style={{fontSize:11,color:"#808080"}}>
                <div style={{display:"flex",justifyContent:"space-between"}}><span>MRR value</span><span style={{fontWeight:500,color:"#FF5000"}}>{fk(totalMRR)}</span></div>
                <div style={{display:"flex",justifyContent:"space-between",marginTop:3}}><span>OTR value</span><span style={{fontWeight:500,color:"#5378FC"}}>{fk(totalOTR)}</span></div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* ── Top 3 by Region spotlight ── */}
      {regionFilter==="all"&&<div style={{...cardStyle,marginBottom:16}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16}}>
          <div style={secTitle}>Top 3 Revenue — By Region</div>
          <div style={{fontSize:11,color:"#808080"}}>Click a region card to filter the full dashboard</div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:16}}>
          {["DR","US","ANZ"].map(reg=>{
            const REG_COL={DR:"#FF5000",US:"#29355D",ANZ:"#5378FC"};
            const REG_LABEL={DR:"Dominican Republic","US":"United States","ANZ":"Australia & NZ"};
            const regCsms = {};
            rows.filter(r=>r.region===reg && (!managerCoaches||managerCoaches.includes(lk(r.csm)&&lk(r.csm).c)) && (!filterCoach||(lk(r.csm)&&lk(r.csm).c===filterCoach)) && (!filterCSM||r.csm===filterCSM)).forEach(r=>{
              if(!regCsms[r.csm]) regCsms[r.csm]={csm:r.csm,total:0,mrr:0,otr:0,subs:0};
              regCsms[r.csm].total+=r.tot; regCsms[r.csm].mrr+=r.mrr;
              regCsms[r.csm].otr+=r.otr; regCsms[r.csm].subs++;
            });
            const top3 = Object.values(regCsms).sort((a,b)=>b.total-a.total).slice(0,3);
            const medals=["🥇","🥈","🥉"];
            return <div key={reg} style={{background:"#F4F6FB",borderRadius:10,padding:16,borderTop:`3px solid ${REG_COL[reg]}`}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
                <div>
                  <span style={{fontSize:18,fontWeight:700,color:REG_COL[reg],marginRight:8}}>{reg}</span>
                  <span style={{fontSize:11,color:"#808080"}}>{REG_LABEL[reg]}</span>
                </div>
                <button onClick={()=>setRegionFilter(reg)} style={{fontSize:10,color:REG_COL[reg],background:"rgba(0,0,0,.05)",border:`0.5px solid ${REG_COL[reg]}`,borderRadius:20,padding:"2px 10px",cursor:"pointer"}}>Filter ↗</button>
              </div>
              {top3.length===0
                ? <div style={{color:"#808080",fontSize:12,fontStyle:"italic"}}>No revenue data</div>
                : top3.map((c,idx)=>(
                  <div key={c.csm} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 0",borderBottom:idx<top3.length-1?"0.5px solid rgba(41,53,93,.08)":"none"}}>
                    <span style={{fontSize:18,width:24,flexShrink:0}}>{medals[idx]}</span>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:12,fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.csm}</div>
                      <div style={{fontSize:10,color:"#808080",marginTop:1}}>{c.subs} sub{c.subs!==1?"s":""}{c.mrr>0?" · MRR "+fk(c.mrr):""}{c.otr>0?" · OTR "+fk(c.otr):""}</div>
                    </div>
                    <div style={{fontSize:14,fontWeight:600,color:REG_COL[reg],flexShrink:0}}>{fd(c.total)}</div>
                  </div>
                ))}
            </div>;
          })}
        </div>
      </div>}

      {/* ── Row 3: Top MRR integrations + Non-revenue breakdown ── */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>

        {/* Top MRR integration types */}
        <div style={cardStyle}>
          <div style={secTitle}>Top MRR Integrations</div>
          {mrrTypeRows.map(([type,d],i)=>(
            <div key={type} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
              <span style={{width:16,fontSize:11,color:"#808080",flexShrink:0}}>{i+1}.</span>
              <span style={{flex:1,fontSize:12,fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{type}</span>
              <div style={{width:120,height:5,background:"#ECEEF1",borderRadius:3,overflow:"hidden",flexShrink:0}}>
                <div style={{height:"100%",background:"#FF5000",opacity:.75,borderRadius:3,width:(d.amount/maxMrrAmt*100).toFixed(1)+"%"}}/>
              </div>
              <span style={{width:20,fontSize:11,color:"#808080",textAlign:"right",flexShrink:0}}>{d.count}</span>
              <span style={{width:60,fontSize:11,fontWeight:500,color:"#FF5000",textAlign:"right",flexShrink:0}}>{fk(d.amount)}</span>
            </div>
          ))}
          {mrrTypeRows.length===0&&<div style={{color:"#808080",fontSize:12}}>No MRR data</div>}
        </div>

        {/* Non-revenue integrations */}
        <div style={cardStyle}>
          <div style={secTitle}>Non-Revenue Integrations</div>
          <div style={{display:"flex",flexWrap:"wrap",gap:10,marginBottom:16}}>
            {Object.entries(nrTypes).sort((a,b)=>b[1]-a[1]).map(([t,c])=>(
              <div key={t} style={{display:"flex",flexDirection:"column",alignItems:"center",background:"#F4F6FB",borderRadius:10,padding:"10px 14px",minWidth:80}}>
                <div style={{fontSize:22,fontWeight:500,color:NR_COLORS[t]||"#29355D"}}>{c}</div>
                <div style={{fontSize:10,color:"#808080",marginTop:2,textAlign:"center",lineHeight:1.3}}>{t}</div>
              </div>
            ))}
          </div>
          <div style={{height:"0.5px",background:"rgba(41,53,93,.08)",marginBottom:12}}/>
          <div style={{display:"flex",justifyContent:"space-between",fontSize:12}}>
            <span style={{color:"#808080"}}>Total non-revenue</span>
            <span style={{fontWeight:500,color:"#29355D"}}>{Object.values(nrTypes).reduce((s,v)=>s+v,0)} integrations</span>
          </div>
        </div>
      </div>

      {/* ── CSM Leaderboard ── */}
      <div style={cardStyle}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
          <div style={secTitle}>CSM Leaderboard</div>
          <div style={{fontSize:11,color:"#808080"}}>{lbRows.length} CSMs</div>
        </div>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead><tr>
            <th style={{...thS,width:28}}>#</th>
            <th style={thS}>CSM</th>
            <th style={thS}>Team</th>
            <th style={thS}>Tier</th>
            <th style={thS}>Region</th>
            {thSort("subs","Subs")}
            {thSort("mrr","MRR")}
            {thSort("otr","OTR")}
            {thSort("revPerSub","Rev/Sub")}
            {thSort("total","Total Revenue")}
          </tr></thead>
          <tbody>
            {visibleLb.map((c,i)=>{
              const info=lk(c.csm)||{};
              const col=TEAM_COLS[info.t||c.team]||"#888";
              const rank = lbRows.indexOf(c);
              return <tr key={c.csm}>
                <td style={tdS}>{rank<3?medals[rank]:(rank+1)+"."}</td>
                <td style={{...tdS,fontWeight:500}}>{c.csm}</td>
                <td style={tdS}><span style={{display:"inline-block",width:6,height:6,borderRadius:"50%",background:col,marginRight:4,verticalAlign:"middle"}}/><span style={{color:"#808080"}}>{st(info.t||c.team)}</span></td>
                <td style={tdS}><span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:"#F4F6FB",color:"#29355D"}}>{tierLabel[c.tier]||c.tier||"--"}</span></td>
                <td style={tdS}>{(()=>{const rg=region(c.csm);const RC={DR:"#FF5000",US:"#29355D",ANZ:"#5378FC"};return rg?<span style={{fontSize:10,fontWeight:600,color:RC[rg]||"#808080"}}>{rg}</span>:"--";})()}</td>
                <td style={tdRS}>{c.subs}</td>
                <td style={{...tdRS,color:c.mrr>0?"#FF5000":"#808080",fontWeight:c.mrr>0?500:400}}>{c.mrr>0?fd(c.mrr):"--"}</td>
                <td style={{...tdRS,color:c.otr>0?"#5378FC":"#808080",fontWeight:c.otr>0?500:400}}>{c.otr>0?fd(c.otr):"--"}</td>
                <td style={{...tdRS,color:"#808080"}}>{c.revPerSub>0?fd(c.revPerSub):"$0"}</td>
                <td style={{...tdRS,color:c.total>0?"#29355D":"#808080",fontWeight:c.total>0?600:400}}>{c.total>0?fd(c.total):"$0.00"}</td>
              </tr>;
            })}
          </tbody>
        </table>
        {lbRows.length>15&&(
          <button onClick={()=>setShowAllLb(s=>!s)}
            style={{width:"100%",marginTop:12,padding:"8px 0",fontSize:12,fontWeight:500,color:"#FF5000",background:"rgba(255,80,0,.06)",border:"0.5px solid rgba(255,80,0,.2)",borderRadius:6,cursor:"pointer"}}>
            {showAllLb?"▲ Show less":"▼ Show all "+lbRows.length+" CSMs ↓"}
          </button>
        )}
      </div>
    </div>
  );
}

// ── BOOK OF BUSINESS VIEW ────────────────────────────────────────────────────
const BOB_GOAL = 0.91;
const BOB_GRAND = {boq:1807046.8,lcm:1644207.36,net:-162839.44,pct:0.90988};
const BOB_COACH_TOTALS = {
  "Trisha Stalnaker":{boq:284581.25,lcm:253068.5,net:-31512.75,pct:0.88926},
  "Mia O\u2019Dirling":{boq:285466.8,lcm:256071.3,net:-29395.5,pct:0.89702},
  "Kendra Morelli":{boq:274742.6,lcm:252566.4,net:-22176.2,pct:0.91928},
  "Elizabeth White":{boq:260702.25,lcm:236185.35,net:-24516.9,pct:0.90595},
  "Aaron Taylor":{boq:543976.0,lcm:497031.91,net:-46944.09,pct:0.9137},
  "Chase Boyd":{boq:157577.9,lcm:149283.9,net:-8294.0,pct:0.94736},
};
const BOB_CSMS = [{"n":"Chelsea Dingus","c":"Kendra Morelli","boq":21073,"lcm":18316,"net":-2757,"ret":0.8692,"mca":59,"mcc":2,"mch":["Chicago HouseMasters LLC","Stepwise Health Platinum LLC."],"bca":1,"bcc":0,"bch":[]},{"n":"Joseph Guillermo Carmona Garcia","c":"Mia O\u2019Dirling","boq":25700.2,"lcm":21370.2,"net":-4330,"ret":0.8315,"mca":70,"mcc":2,"mch":["Distance Movers","Do It All Mobile Auto Spa"],"bca":1,"bcc":0,"bch":[]},{"n":"Tracy-Ann Gaudencio","c":"Aaron Taylor","boq":28181.32,"lcm":28798.32,"net":617,"ret":1.0219,"mca":73,"mcc":1,"mch":["Account Elite Spray Pave"],"bca":1,"bcc":1,"bch":["AC Gutter Guard"]},{"n":"Dave Crisler","c":"Aaron Taylor","boq":57248.68,"lcm":51738.68,"net":-5510,"ret":0.9037,"mca":158,"mcc":3,"mch":["Shop N Go Car Wash & Care","Top Mix Construction","Yardner"],"bca":2,"bcc":1,"bch":["Tony Hollands Funerals"]},{"n":"Sylvia Appla","c":"Aaron Taylor","boq":54619.5,"lcm":47443.5,"net":-7176,"ret":0.8686,"mca":152,"mcc":9,"mch":["Chisham Express Pharmacy","AAAA Brick Broom Cleaning","Jims Mowing Glen Waverley 1","Jims Mowing Croydon Hills","Viva Voce Choir"],"bca":0,"bcc":0,"bch":[]},{"n":"Ellise Payne","c":"Aaron Taylor","boq":46979.16,"lcm":44413.16,"net":-2566,"ret":0.9454,"mca":125,"mcc":7,"mch":["HAIR @ THE HUB","Easy Excavators","Saferoads Holdings","Erina Auto Parts","HBW Manufacturing"],"bca":0,"bcc":0,"bch":[]},{"n":"Nikita Siepen-Bowers","c":"Aaron Taylor","boq":46777.68,"lcm":42309.68,"net":-4468,"ret":0.9045,"mca":135,"mcc":5,"mch":["Wizard Motors Pty Ltd","Mr Hook Towing and Metal","Sai Thai Restaurant","Manawatu Engineering","Trenchless Technology"],"bca":0,"bcc":0,"bch":[]},{"n":"Warda Gul","c":"Aaron Taylor","boq":55478.96,"lcm":50242.96,"net":-5236,"ret":0.9056,"mca":152,"mcc":3,"mch":["JRs Mower & Motorcycle","Earth 2 Ocean Communications","Otagro Fertilizers Ltd"],"bca":0,"bcc":0,"bch":[]},{"n":"Indu Vijay","c":"Aaron Taylor","boq":52888.08,"lcm":48468.08,"net":-4420,"ret":0.9164,"mca":128,"mcc":5,"mch":["Complete Pool Services","Fresh Concept Foodservice","Pure Fresh Cleaning Services","Pacific Funerals Group","The Trustee"],"bca":0,"bcc":0,"bch":[]},{"n":"Matt Daly","c":"Aaron Taylor","boq":47918.72,"lcm":43990.72,"net":-3928,"ret":0.918,"mca":113,"mcc":4,"mch":["Better Service Solutions","Aotearoa Kiwi Tours","The Garden Guru","Christchurch Building Inspections"],"bca":0,"bcc":0,"bch":[]},{"n":"Peter Manalac","c":"Aaron Taylor","boq":53736.48,"lcm":49224.48,"net":-4512,"ret":0.9161,"mca":143,"mcc":3,"mch":["Rapid Electrical","The Blind Spot Blinds","SADDLERS WELDING"],"bca":0,"bcc":0,"bch":[]},{"n":"Zoltan Rudolf","c":"Aaron Taylor","boq":41388.16,"lcm":36966.16,"net":-4422,"ret":0.8932,"mca":103,"mcc":4,"mch":["Horowhenua Tractor Parts","Taikura Rudolf Steiner","Havelock Village","Masterton Joinery"],"bca":0,"bcc":0,"bch":[]},{"n":"Sakshi Mahalwal","c":"Aaron Taylor","boq":58759.24,"lcm":54437.24,"net":-4322,"ret":0.9265,"mca":142,"mcc":5,"mch":["Cooma Hospital","Pacific Coast Crane Hire","Tasman Insulation NZ Ltd","Taumarunui Car Sales","Mana Vehicle Testing Station"],"bca":0,"bcc":0,"bch":[]},{"n":"Alejandro Rodriguez-Medina","c":"Kendra Morelli","boq":24003.2,"lcm":21219.2,"net":-2784,"ret":0.884,"mca":65,"mcc":5,"mch":["Kountry Korner Hair Salon","Lonestar Lube N Tune","A & H Cates Carpentry","J & K Transport","Hue & Stitches"],"bca":1,"bcc":0,"bch":[]},{"n":"Karmita Turner","c":"Kendra Morelli","boq":23827.8,"lcm":20975.8,"net":-2852,"ret":0.8803,"mca":56,"mcc":3,"mch":["Salon Obsession","Salon On The Blvd","Prestige Motorsports"],"bca":0,"bcc":0,"bch":[]},{"n":"Lauren Carter","c":"Kendra Morelli","boq":23085.2,"lcm":21469.2,"net":-1616,"ret":0.9301,"mca":69,"mcc":4,"mch":["Right Choice Travel","Sperling Mortuary","Sohna Interior","Gentry Moving & Storage"],"bca":0,"bcc":0,"bch":[]},{"n":"Libby Booher","c":"Kendra Morelli","boq":22660.4,"lcm":22260.4,"net":-400,"ret":0.9823,"mca":62,"mcc":2,"mch":["Lonestar Ag Credit","Lonestar"],"bca":0,"bcc":0,"bch":[]},{"n":"Misti Dixon","c":"Kendra Morelli","boq":26034.8,"lcm":23178.8,"net":-2856,"ret":0.8903,"mca":71,"mcc":4,"mch":["Scentsy - Heather Blanton","GoodGear","Stacy's Tiny Tails","Beaus Glass"],"bca":0,"bcc":0,"bch":[]},{"n":"Misty Decatur","c":"Kendra Morelli","boq":19820.4,"lcm":17912.4,"net":-1908,"ret":0.9037,"mca":56,"mcc":2,"mch":["Estes Interiors","Ledbetter Pool Service"],"bca":0,"bcc":0,"bch":[]},{"n":"Saira Julian Guzman","c":"Kendra Morelli","boq":14219.8,"lcm":12875.8,"net":-1344,"ret":0.9055,"mca":38,"mcc":1,"mch":["Hernandez Architecture"],"bca":0,"bcc":0,"bch":[]},{"n":"Scott Mather","c":"Kendra Morelli","boq":19891.8,"lcm":18659.8,"net":-1232,"ret":0.938,"mca":54,"mcc":2,"mch":["Spectrum Pest Control","Woodlands Funeral"],"bca":0,"bcc":0,"bch":[]},{"n":"Steven Saunders","c":"Kendra Morelli","boq":19879.4,"lcm":17983.4,"net":-1896,"ret":0.9046,"mca":56,"mcc":3,"mch":["Paul Schier Plumbing","Homestead Floral","Westside Collision"],"bca":0,"bcc":0,"bch":[]},{"n":"Dorka Frias Lantigua","c":"Kendra Morelli","boq":9547.6,"lcm":8497.6,"net":-1050,"ret":0.89,"mca":25,"mcc":1,"mch":["Luxury Pool Management"],"bca":0,"bcc":0,"bch":[]},{"n":"Kellie Lester","c":"Trisha Stalnaker","boq":35383.6,"lcm":33894.6,"net":-1489,"ret":0.9579,"mca":91,"mcc":5,"mch":["Patriot Properties","Bees by Amos","Russ Dunmire","B2B Network Promotions","Daves Auto & Tires"],"bca":0,"bcc":0,"bch":[]},{"n":"Karissa Hernandez","c":"Trisha Stalnaker","boq":37153.75,"lcm":30192,"net":-6961.75,"ret":0.8126,"mca":91,"mcc":9,"mch":["Parable Christian Store","4 Seasons Property","Advantage Electric","Elites Dance Academy","Hillside Homecare"],"bca":0,"bcc":0,"bch":[]},{"n":"Ashley Vasquez Mena","c":"Trisha Stalnaker","boq":38476.1,"lcm":33200.1,"net":-5276,"ret":0.8629,"mca":97,"mcc":7,"mch":["M & M Industrial","L & S Mechanical","Luxe Dental","Key-Way Lock","City Of Life"],"bca":0,"bcc":0,"bch":[]},{"n":"Karen Capellan Tavarez","c":"Trisha Stalnaker","boq":19914.6,"lcm":18261.6,"net":-1653,"ret":0.917,"mca":54,"mcc":3,"mch":["Always Green Lawn","Accion Chicago","Greenwood Pharmacy"],"bca":0,"bcc":0,"bch":[]},{"n":"Ashley Shaffer","c":"Trisha Stalnaker","boq":34368.4,"lcm":28136.4,"net":-6232,"ret":0.8187,"mca":87,"mcc":11,"mch":["Kiddie Kandids","One Source Comm","Happy Hearts Child","Bingo Night Events","Canary Wharf Grill"],"bca":0,"bcc":0,"bch":[]},{"n":"Merve (MJ) Brielmann","c":"Trisha Stalnaker","boq":24891,"lcm":24507,"net":-384,"ret":0.9846,"mca":62,"mcc":2,"mch":["Bella Vista Pools","Armuchee"],"bca":0,"bcc":0,"bch":[]},{"n":"Taylor Kidd","c":"Trisha Stalnaker","boq":24930.2,"lcm":22369.2,"net":-2561,"ret":0.8973,"mca":58,"mcc":4,"mch":["Twin Cities Comfort","Fitness Zone","Perimeter Pest Control","Camelot Party Supplies"],"bca":0,"bcc":0,"bch":[]},{"n":"Mark Velazquez","c":"Trisha Stalnaker","boq":15438,"lcm":13621,"net":-1817,"ret":0.8823,"mca":37,"mcc":3,"mch":["Trident Plumbing","Saddleback Pest","Mid States Petroleum"],"bca":0,"bcc":0,"bch":[]},{"n":"Felix Caba Jimenez","c":"Trisha Stalnaker","boq":1920,"lcm":1920,"net":0,"ret":1.0,"mca":5,"mcc":0,"mch":[],"bca":0,"bcc":0,"bch":[]},{"n":"Stacy Roers","c":"Trisha Stalnaker","boq":19615.2,"lcm":18590.2,"net":-1025,"ret":0.9477,"mca":51,"mcc":2,"mch":["Glacier Hills","Northwood Auto"],"bca":0,"bcc":0,"bch":[]},{"n":"Rafael Sencion Sencion","c":"Trisha Stalnaker","boq":32490.4,"lcm":28376.4,"net":-4114,"ret":0.8734,"mca":84,"mcc":7,"mch":["Charis Music","Sano Wellness","Magnolia Dental","Blue Sky Realty","Woodmont Hills"],"bca":0,"bcc":0,"bch":[]},{"n":"Victor Abner Moscoso Fernandez","c":"Mia O\u2019Dirling","boq":38942.2,"lcm":36674.2,"net":-2268,"ret":0.9418,"mca":100,"mcc":4,"mch":["Central Florida Towing","Top Notch Beauty","Elite Car Wash","Pines Florist"],"bca":0,"bcc":0,"bch":[]},{"n":"Heidi Torres Uribe","c":"Mia O\u2019Dirling","boq":19764.4,"lcm":18436.4,"net":-1328,"ret":0.9328,"mca":54,"mcc":2,"mch":["Golden Touch Cleaning","Altagracia Beauty"],"bca":0,"bcc":0,"bch":[]},{"n":"Darling Danais Santos Taveras","c":"Mia O\u2019Dirling","boq":19197.6,"lcm":17373.6,"net":-1824,"ret":0.9049,"mca":52,"mcc":3,"mch":["Guzman Landscaping","Rosy Nails","Express Auto Repair"],"bca":0,"bcc":0,"bch":[]},{"n":"Irina Larianni Molina Molina","c":"Mia O\u2019Dirling","boq":10993.2,"lcm":10217.2,"net":-776,"ret":0.9294,"mca":30,"mcc":1,"mch":["Larios Convenience Store"],"bca":0,"bcc":0,"bch":[]},{"n":"Wilson Mercedes","c":"Mia O\u2019Dirling","boq":30065.8,"lcm":26945.8,"net":-3120,"ret":0.8963,"mca":80,"mcc":5,"mch":["Garcia Auto Shop","Queens Bridal","Cali Flowers","Pro Image Sports","Midwest Hauling"],"bca":0,"bcc":0,"bch":[]},{"n":"Jathzelyn Elizabeth Fortuna Paulino","c":"Mia O\u2019Dirling","boq":20671,"lcm":19043,"net":-1628,"ret":0.9213,"mca":53,"mcc":2,"mch":["Bella Nails","Style Zone"],"bca":0,"bcc":0,"bch":[]},{"n":"Yessica Montero Urena","c":"Mia O\u2019Dirling","boq":18399.6,"lcm":16523.6,"net":-1876,"ret":0.8981,"mca":50,"mcc":3,"mch":["Corazon BBQ","Mariana Flowers","Star Nails"],"bca":0,"bcc":0,"bch":[]},{"n":"Johnny Cornielle","c":"Mia O\u2019Dirling","boq":29088.4,"lcm":25356.4,"net":-3732,"ret":0.8717,"mca":78,"mcc":6,"mch":["Blessed Hands Barbershop","City Lights Diner","Prestige Auto","Diamond Cuts","New Wave"],"bca":0,"bcc":0,"bch":[]},{"n":"Sati Ananda Pimentel Malespin","c":"Mia O\u2019Dirling","boq":27040.6,"lcm":24220.6,"net":-2820,"ret":0.8957,"mca":71,"mcc":4,"mch":["Primera Iglesia","Tropical Nails","Elegant Touch","Star Bright"],"bca":0,"bcc":0,"bch":[]},{"n":"Samuel Frias De Paula","c":"Mia O\u2019Dirling","boq":35963.4,"lcm":31866.4,"net":-4097,"ret":0.8861,"mca":94,"mcc":5,"mch":["Santos Beauty","Lux Auto","Island Cuts","Paradise Nails","Glamour Touch"],"bca":0,"bcc":0,"bch":[]},{"n":"Barbara Larrosa Presinal","c":"Chase Boyd","boq":23019.6,"lcm":22379.6,"net":-640,"ret":0.9722,"mca":63,"mcc":2,"mch":["Prestige Cleaning","All Star Sports"],"bca":6,"bcc":1,"bch":["Texas Outdoor Projects"]},{"n":"Deivis Pena","c":"Chase Boyd","boq":19290.4,"lcm":17986.4,"net":-1304,"ret":0.9324,"mca":54,"mcc":3,"mch":["Elite Fence","Gold Star Auto","Diamond Nails"],"bca":0,"bcc":0,"bch":[]},{"n":"Kyle Dye","c":"Chase Boyd","boq":22069.4,"lcm":20661.4,"net":-1408,"ret":0.9362,"mca":58,"mcc":4,"mch":["Cornerstone Realty","Main Street Diner","Heritage Flooring","Summit Pest"],"bca":0,"bcc":0,"bch":[]},{"n":"Sarah Swanson","c":"Chase Boyd","boq":20193.6,"lcm":19869.6,"net":-324,"ret":0.984,"mca":58,"mcc":1,"mch":["Sundown Ranch"],"bca":0,"bcc":0,"bch":[]},{"n":"Luis Aguasvivas Peralta","c":"Chase Boyd","boq":8272.4,"lcm":7812.4,"net":-460,"ret":0.9444,"mca":22,"mcc":1,"mch":["Santos Landscaping"],"bca":0,"bcc":0,"bch":[]},{"n":"Juan Liberato","c":"Chase Boyd","boq":13063.4,"lcm":12191.4,"net":-872,"ret":0.9333,"mca":36,"mcc":2,"mch":["Tropical Cuts","Island Fresh"],"bca":0,"bcc":0,"bch":[]},{"n":"Elianny Tena Antigua","c":"Chase Boyd","boq":9226.8,"lcm":8468.8,"net":-758,"ret":0.9179,"mca":25,"mcc":1,"mch":["Nails By Design"],"bca":0,"bcc":0,"bch":[]},{"n":"Damita Hill","c":"Elizabeth White","boq":28434.25,"lcm":24734.25,"net":-3700,"ret":0.8699,"mca":72,"mcc":6,"mch":["Park Place Cleaners","Sunshine Academy","Royal Cuts","Elite Image","Pro Nails","Blossom Day Spa"],"bca":0,"bcc":0,"bch":[]},{"n":"Anthony Yen","c":"Elizabeth White","boq":19093.2,"lcm":18005.2,"net":-1088,"ret":0.943,"mca":53,"mcc":3,"mch":["Pacific Rim","Harbor View","Coastal Cuts"],"bca":0,"bcc":0,"bch":[]},{"n":"April Hall","c":"Elizabeth White","boq":22714.8,"lcm":20578.8,"net":-2136,"ret":0.9059,"mca":59,"mcc":4,"mch":["Summit Auto","Valley Dental","Mountain Fresh","Peak Performance"],"bca":0,"bcc":0,"bch":[]},{"n":"Katelyn Ankrom","c":"Elizabeth White","boq":27124.8,"lcm":24412.8,"net":-2712,"ret":0.9001,"mca":72,"mcc":4,"mch":["Bloom Floral","Green Thumb Nursery","Petal Perfect","Garden Gate"],"bca":0,"bcc":0,"bch":[]},{"n":"Kennedy Sanchez","c":"Elizabeth White","boq":16530.6,"lcm":14994.6,"net":-1536,"ret":0.9071,"mca":44,"mcc":2,"mch":["Sunrise Bakery","Golden Key Realty"],"bca":0,"bcc":0,"bch":[]},{"n":"Michael Furlong","c":"Elizabeth White","boq":28540.4,"lcm":25224.4,"net":-3316,"ret":0.8838,"mca":74,"mcc":6,"mch":["Riverside Clinic","Lakeside Auto","Forest Glen","Brook Valley","Clear Water","Sunset Pools"],"bca":0,"bcc":0,"bch":[]},{"n":"Yolanda Ramirez","c":"Elizabeth White","boq":29175,"lcm":26827,"net":-2348,"ret":0.9195,"mca":75,"mcc":4,"mch":["Desert Rose","Oasis Pools","Palm Tree Realty","Cactus Auto"],"bca":0,"bcc":0,"bch":[]},{"n":"Florence Francois Nova","c":"Elizabeth White","boq":22961.4,"lcm":21265.4,"net":-1696,"ret":0.9261,"mca":60,"mcc":3,"mch":["Belle Epoque","Riviera Nails","Paris Dreams"],"bca":0,"bcc":0,"bch":[]},{"n":"Rossi Valerio Tejeda","c":"Elizabeth White","boq":24220,"lcm":21513,"net":-2707,"ret":0.8882,"mca":63,"mcc":5,"mch":["Tropical Breeze","Island Dreams","Caribbean Cuts","Palm Bay","Sunset Nails"],"bca":0,"bcc":0,"bch":[]}];

const BOB_DETAIL = {"Katelyn Ankrom":{"i":[],"d":[{"e":"BSFD0842","a":"Coconut Health","l":"Websites","b":117.0,"m":0.0,"n":-117.0},{"e":"BSFV9004","a":"Our World ID","l":"Social","b":750.0,"m":0.0,"n":-750.0},{"e":"BSFV9004","a":"Our World ID","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0}]},"Sylvia Appla":{"i":[{"e":"BSDZ5782","a":"Rawlins Neville","l":"Marketing Ctr","b":499.09,"m":509.09,"n":10.0}],"d":[{"e":"BSFC6415","a":"TotallyOT","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSFD6635","a":"Stix and Tonez","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSFV5882","a":"Keep Healthy Massage","l":"Marketing Ctr","b":358.19,"m":351.82,"n":-6.37},{"e":"BSFD6638","a":"C & T Earthworks","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSDK7610","a":"Chisham Express Pharmacy","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSDK7610","a":"Chisham Express Pharmacy","l":"Websites","b":131.82,"m":0.0,"n":-131.82},{"e":"BSFP7559","a":"Dale's Tree Services","l":"Marketing Ctr","b":570.91,"m":0.0,"n":-570.91},{"e":"BSFD8439","a":"Frostbite Auto Air","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0}]},"Libby Booher":{"i":[{"e":"BSCZ0276","a":"Whisked Away","l":"Platform Other","b":91.0,"m":126.0,"n":35.0}],"d":[{"e":"BSFR1218","a":"Varsity Flooring LLC","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFP2668","a":"Zebra Stripes Child Care & Preschool","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFP2668","a":"Zebra Stripes Child Care & Preschool","l":"Business Ctr","b":179.0,"m":0.0,"n":-179.0},{"e":"BSFX4774","a":"Faith Window Cleaning","l":"Growth Pkgs","b":300.0,"m":0.0,"n":-300.0},{"e":"BSFP7298","a":"CAMPBELL REMODELING AND MAINTENANCE LLC","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFP7298","a":"CAMPBELL REMODELING AND MAINTENANCE LLC","l":"Business Ctr","b":179.0,"m":0.0,"n":-179.0}]},"Merve (MJ) Brielmann":{"i":[],"d":[{"e":"BSFD7986","a":"OTU+","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0}]},"Karen Capellan Tavarez":{"i":[],"d":[{"e":"BSFF1595","a":"Gutter Customs LLC","l":"Websites","b":117.0,"m":0.0,"n":-117.0},{"e":"BRXB6268","a":"Tapia Demolition","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF5427","a":"PM Landscaping","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF2698","a":"EZ Garage Solutions","l":"Marketing Ctr","b":293.0,"m":0.0,"n":-293.0},{"e":"BSFF2698","a":"EZ Garage Solutions","l":"Social","b":500.0,"m":0.0,"n":-500.0},{"e":"BSDV3762","a":"Conecta Inc Chicago","l":"Business Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFZ2268","a":"Rescore America","l":"Marketing Ctr","b":304.0,"m":293.0,"n":-11.0}]},"Joseph Guillermo Carmona Garcia":{"i":[],"d":[{"e":"BSFC0370","a":"Radiant Energy","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFC0370","a":"Radiant Energy","l":"Websites","b":117.0,"m":0.0,"n":-117.0},{"e":"BSFX9079","a":"Moreland Exteriors","l":"Marketing Ctr","b":400.0,"m":384.0,"n":-16.0},{"e":"BSFP7403","a":"Reyes Roofing","l":"Marketing Ctr","b":433.0,"m":0.0,"n":-433.0},{"e":"BSFF2291","a":"Better Ways Home Services","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFC3268","a":"Distance Movers","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFD0406","a":"Eagle Hardwood Flooring, Inc.","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFC4086","a":"PMV Limousine INC","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFC2925","a":"guHome Remodeling Roofing & Gutters Plus","l":"Growth Pkgs","b":300.0,"m":0.0,"n":-300.0},{"e":"BSFF4071","a":"Persic Heating & Air, LLC","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFG0022","a":"Precision European Auto Repair","l":"Thryv Leads","b":2000.0,"m":0.0,"n":-2000.0}]},"Lauren Carter":{"i":[],"d":[{"e":"BSCQ8015","a":"Shea's Health and Nutrition","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFG0266","a":"SparrowHawk Mobile Detailing","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFG0266","a":"SparrowHawk Mobile Detailing","l":"SEO","b":900.0,"m":0.0,"n":-900.0},{"e":"BSFG0266","a":"SparrowHawk Mobile Detailing","l":"Social","b":500.0,"m":0.0,"n":-500.0},{"e":"BSFD2805","a":"Freeway Graphics & Design","l":"Marketing Ctr","b":433.0,"m":0.0,"n":-433.0},{"e":"BSFD2805","a":"Freeway Graphics & Design","l":"Thryv Leads","b":1500.0,"m":0.0,"n":-1500.0},{"e":"BRWZ3696","a":"Campbell's Sewer & Drain Cleaning","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFM0705","a":"Equipment Guardians LLC","l":"Marketing Ctr","b":433.0,"m":384.0,"n":-49.0},{"e":"BSDZ9243","a":"JD3 Mobile DNA Services LLC","l":"Business Ctr","b":254.0,"m":0.0,"n":-254.0},{"e":"BSDZ9243","a":"JD3 Mobile DNA Services LLC","l":"Websites","b":117.0,"m":0.0,"n":-117.0},{"e":"BSFD0438","a":"Andy's Pro Tire & Auto","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFP7420","a":"Marketside Chiropractic","l":"SEO","b":400.0,"m":0.0,"n":-400.0}]},"Johnny Cornielle":{"i":[],"d":[{"e":"BSFM3039","a":"Cool Concepts Inc.","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFM3039","a":"Cool Concepts Inc.","l":"Recurring Labor","b":125.0,"m":0.0,"n":-125.0},{"e":"BSFC7378","a":"Safe Resolutions LLC","l":"Business Ctr","b":571.0,"m":0.0,"n":-571.0},{"e":"BSFC7323","a":"ALL THINGS PINKABOO","l":"Business Ctr","b":244.0,"m":0.0,"n":-244.0}]},"Dave Crisler":{"i":[],"d":[{"e":"BSDM8338","a":"Logan West Laundromat","l":"Marketing Ctr","b":286.37,"m":280.0,"n":-6.37},{"e":"BSFP6650","a":"M1 business system","l":"Thryv Leads","b":545.0,"m":0.0,"n":-545.0},{"e":"BSFP6650","a":"M1 business system","l":"Marketing Ctr","b":499.09,"m":0.0,"n":-499.09},{"e":"BSDM0105","a":"Manuka Dental Care","l":"Websites","b":131.82,"m":0.0,"n":-131.82},{"e":"BSDM0105","a":"Manuka Dental Care","l":"SEO","b":600.0,"m":0.0,"n":-600.0},{"e":"BSDL2104","a":"Kirbside Clean A Bin","l":"Thryv Leads","b":953.0,"m":0.0,"n":-953.0},{"e":"BSCZ1564","a":"MEP PLUMBING LIMITED","l":"Marketing Ctr","b":330.0,"m":318.0,"n":-12.0},{"e":"BSFF8532","a":"Shop N Go Car Wash & Care","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSFM2341","a":"Imperial Detail Co.","l":"Marketing Ctr","b":318.0,"m":0.0,"n":-318.0},{"e":"BSFP3560","a":"Top Mix Construction","l":"Marketing Ctr","b":570.91,"m":0.0,"n":-570.91},{"e":"BSFP3560","a":"Top Mix Construction","l":"Growth Pkgs","b":454.55,"m":0.0,"n":-454.55},{"e":"BSFB1708","a":"Sabetian Natural Appearance Clinic","l":"Marketing Ctr","b":318.0,"m":0.0,"n":-318.0},{"e":"BSFB1708","a":"Sabetian Natural Appearance Clinic","l":"SEO","b":2100.0,"m":0.0,"n":-2100.0}]},"Matt Daly":{"i":[],"d":[{"e":"BSFM0691","a":"Bright & White Teeth Whitening","l":"Social","b":750.0,"m":0.0,"n":-750.0},{"e":"BSFP3574","a":"Perth Taxi Booking","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSFD6682","a":"SPEEDLINE Tree Surgery Pty Ltd","l":"Thryv Leads","b":909.0,"m":0.0,"n":-909.0}]},"Misty Decatur":{"i":[{"e":"BFQC0918","a":"Dubois County Garage Doors","l":"Platform Other","b":7.2,"m":43.2,"n":36.0},{"e":"BRPJ7139","a":"Manwill Plumbing & Heating","l":"Thryv Leads","b":10000.0,"m":12000.0,"n":2000.0}],"d":[{"e":"BSFB9303","a":"Dollar Bill's Heating and Air","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFC8129","a":"Lakeshore Custom Masonry","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFR5072","a":"SERVPRO of Lenoir, Duplin & Jones Counties","l":"Thryv Leads","b":2400.0,"m":0.0,"n":-2400.0},{"e":"BSFR5072","a":"SERVPRO of Lenoir, Duplin & Jones Counties","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0}]},"Chelsea Dingus":{"i":[],"d":[{"e":"BSFV6235","a":"Bullet Hole Annex","l":"SEO","b":700.0,"m":0.0,"n":-700.0},{"e":"BSFC0344","a":"Jerry's Paint & Body Shop","l":"Social","b":750.0,"m":0.0,"n":-750.0},{"e":"BSFC0344","a":"Jerry's Paint & Body Shop","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFR9436","a":"Master Tech Transmission","l":"Growth Pkgs","b":500.0,"m":0.0,"n":-500.0},{"e":"BSFP3006","a":"Beautistry Makeup Academy","l":"Business Ctr","b":179.0,"m":0.0,"n":-179.0},{"e":"BSFP3006","a":"Beautistry Makeup Academy","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0}]},"Misti Dixon":{"i":[],"d":[{"e":"BSFB5643","a":"Main Line Benefits- Barry Schumann","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFD8782","a":"Source Light Wellness Center","l":"Business Ctr","b":179.0,"m":0.0,"n":-179.0},{"e":"BSFD7585","a":"Beyond Roofing LLC","l":"Marketing Ctr","b":433.0,"m":0.0,"n":-433.0}]},"Kyle Dye":{"i":[],"d":[{"e":"BSFC2761","a":"Pickaway Flooring Center","l":"Marketing Ctr","b":433.0,"m":0.0,"n":-433.0},{"e":"BRXM5706","a":"Decorative Concrete Designer","l":"Marketing Ctr","b":433.0,"m":0.0,"n":-433.0},{"e":"BSDZ8620","a":"Springtown Electric","l":"SEO","b":800.0,"m":0.0,"n":-800.0}]},"Jathzelyn Elizabeth Fortuna Paulino":{"i":[],"d":[{"e":"BSFD0151","a":"Ollies Electric, LLC","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFD9975","a":"An Apple A Day Nutrition Counseling","l":"Marketing Ctr","b":628.0,"m":244.0,"n":-384.0},{"e":"BSFM0571","a":"OaksMark","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFQ9159","a":"Mark's Electric LLC","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFZ2626","a":"Faithful Turf","l":"Marketing Ctr","b":400.0,"m":384.0,"n":-16.0},{"e":"BSFD8028","a":"Precision Injury Law","l":"Marketing Ctr","b":293.0,"m":0.0,"n":-293.0},{"e":"BSFD8028","a":"Precision Injury Law","l":"Social","b":750.0,"m":0.0,"n":-750.0},{"e":"BRQB2178","a":"Uebe Insured","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFQ2873","a":"Fleming-moving llc","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF2428","a":"Jarred Pierce Real Estate","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BRXK9889","a":"Big & Little Storage","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFG0515","a":"Cowboy Roadside Services","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFF0131","a":"Urgent Dental Care/Smile Centers of America","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFD9863","a":"ABS The Clean Machine","l":"Social","b":750.0,"m":0.0,"n":-750.0},{"e":"BSFD0870","a":"Pure Power CPAs","l":"Social","b":750.0,"m":0.0,"n":-750.0}]},"Florence Francois Nova":{"i":[],"d":[{"e":"BRXD8965","a":"Allen & Hoshall","l":"Marketing Ctr","b":141.17,"m":0.0,"n":-141.17},{"e":"BRXD8965","a":"Allen & Hoshall","l":"Growth Pkgs","b":183.33,"m":0.0,"n":-183.33}]},"Samuel Frias De Paula":{"i":[],"d":[{"e":"BRXR4376","a":"Buell & Olivieri Insurance","l":"Marketing Ctr","b":293.0,"m":0.0,"n":-293.0},{"e":"BSCN3394","a":"A Lonestar Service Blinds and Shutters","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF1866","a":"Johnson's Roofing","l":"Marketing Ctr","b":433.0,"m":0.0,"n":-433.0},{"e":"BSFP7451","a":"Tanaechi","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0}]},"Dorka Frias Lantigua":{"i":[],"d":[{"e":"BSFD2732","a":"Dovinh Group","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSCW0716","a":"Nebraska 7v7 Football","l":"Platform Other","b":7.2,"m":0.0,"n":-7.2},{"e":"BSCW0716","a":"Nebraska 7v7 Football","l":"Business Ctr","b":179.0,"m":0.0,"n":-179.0},{"e":"BSCW0716","a":"Nebraska 7v7 Football","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BJVF3798","a":"Flower Shoppe Inc","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSDQ2902","a":"Vaughn Overhead Door","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFF5876","a":"Paramount Elevator Services","l":"Platform Other","b":21.6,"m":0.0,"n":-21.6},{"e":"BSFG3622","a":"Modern Man Studio","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFQ3508","a":"Guildcraft Exteriors","l":"Growth Pkgs","b":500.0,"m":0.0,"n":-500.0},{"e":"BSFQ3508","a":"Guildcraft Exteriors","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0}]},"Michael Furlong":{"i":[],"d":[{"e":"BSFP7811","a":"Jeune Aesthetics","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BKGK5247","a":"Aloota Farley Co L.P>A","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFR1756","a":"Xpress One Plumber","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFR1756","a":"Xpress One Plumber","l":"Thryv Leads","b":2600.0,"m":0.0,"n":-2600.0},{"e":"BSFF7993","a":"Gustavo Chavez Realtor","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFM2827","a":"Cedar Dental Clinic","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFP4508","a":"Terminatus Pest Control","l":"Growth Pkgs","b":500.0,"m":0.0,"n":-500.0}]},"Tracy-Ann Gaudencio":{"i":[{"e":"BSFF9092","a":"Heatherton Dental","l":"Thryv Leads","b":1818.0,"m":2727.0,"n":909.0}],"d":[{"e":"BSFF7479","a":"Account Elite Spray Pave","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSCZ4215","a":"Kokich Electrical Limited","l":"Marketing Ctr","b":330.0,"m":318.0,"n":-12.0}]},"Warda Gul":{"i":[{"e":"BSFC6690","a":"BrightR Financial","l":"Marketing Ctr","b":499.09,"m":509.09,"n":10.0}],"d":[{"e":"BSFC6809","a":"Affordable Flooring Solutions","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSFC6809","a":"Affordable Flooring Solutions","l":"Business Ctr","b":320.0,"m":0.0,"n":-320.0},{"e":"BSFF6463","a":"Decorug Carpet Cleaning","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSFD7053","a":"ABC Test Tag & Fire Australia","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSFG3134","a":"The Daylesford Meat Co","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSFC6760","a":"Synergy Beauty Supplies","l":"Marketing Ctr","b":499.09,"m":0.0,"n":-499.09},{"e":"BSFF1002","a":"Shangri-La Decors","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSFF1002","a":"Shangri-La Decors","l":"Social","b":1150.0,"m":0.0,"n":-1150.0},{"e":"BSFG0509","a":"KINGSTON AUTOMOTIVE ELECTRICS PTY. LTD.","l":"Websites","b":131.82,"m":0.0,"n":-131.82},{"e":"BSFG0509","a":"KINGSTON AUTOMOTIVE ELECTRICS PTY. LTD.","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSFC6545","a":"Davern & Co","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSDK8933","a":"JRs Mower & Motorcycle","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSDK8933","a":"JRs Mower & Motorcycle","l":"Websites","b":131.82,"m":0.0,"n":-131.82},{"e":"BSDL7066","a":"R G Electrical Pty Ltd","l":"Marketing Ctr","b":286.37,"m":280.0,"n":-6.37}]},"April Hall":{"i":[],"d":[{"e":"BSBS1494","a":"SCL4 LLC","l":"Platform Other","b":7.2,"m":0.0,"n":-7.2},{"e":"BSBS1494","a":"SCL4 LLC","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFF5886","a":"Stronger Built Construction","l":"Marketing Ctr","b":433.0,"m":0.0,"n":-433.0},{"e":"BRZR7821","a":"Harrell Mobile Auto Repair","l":"Marketing Ctr","b":293.0,"m":244.0,"n":-49.0},{"e":"BSFC9928","a":"Goals4Sports","l":"Marketing Ctr","b":293.0,"m":244.0,"n":-49.0},{"e":"BSFF2916","a":"Screen Mobile of Lubbock","l":"Marketing Ctr","b":293.0,"m":0.0,"n":-293.0},{"e":"BSFF2916","a":"Screen Mobile of Lubbock","l":"Social","b":1000.0,"m":0.0,"n":-1000.0}]},"Karissa Hernandez":{"i":[],"d":[{"e":"BSFC0179","a":"Zero Spore Restoration","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFC5052","a":"GBD Concrete","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFS0459","a":"Vital Blooms Wellness","l":"Business Ctr","b":179.0,"m":0.0,"n":-179.0},{"e":"BKQN3456","a":"FLB Insurance Agency","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BRXM4163","a":"Magnolia Mobile Veterinary","l":"Marketing Ctr","b":277.75,"m":0.0,"n":-277.75},{"e":"BSFF6632","a":"Hope Restored Missions","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF9284","a":"AutoVisor 360","l":"Marketing Ctr","b":293.0,"m":0.0,"n":-293.0},{"e":"BSFQ2920","a":"Vantage Pest Control","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFF4034","a":"East Texas Flames LLC","l":"Social","b":750.0,"m":0.0,"n":-750.0},{"e":"BSFF4034","a":"East Texas Flames LLC","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFP4568","a":"Acme Roofing & Siding LLC","l":"Thryv Leads","b":2200.0,"m":0.0,"n":-2200.0},{"e":"BSFP4568","a":"Acme Roofing & Siding LLC","l":"Marketing Ctr","b":628.0,"m":384.0,"n":-244.0},{"e":"BSFC4104","a":"Heaven Storm Roofing","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0}]},"Damita Hill":{"i":[],"d":[{"e":"BSFD9778","a":"Woodard Assurance","l":"Platform Other","b":7.2,"m":0.0,"n":-7.2},{"e":"BSFD9778","a":"Woodard Assurance","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFZ1492","a":"Flash Fence","l":"Social","b":500.0,"m":0.0,"n":-500.0},{"e":"BSFZ1492","a":"Flash Fence","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFZ3386","a":"Feelin Drippy Mobile IV Therapy","l":"Marketing Ctr","b":433.0,"m":0.0,"n":-433.0},{"e":"BSFQ9230","a":"L&R Excavation","l":"Social","b":500.0,"m":0.0,"n":-500.0},{"e":"BSFQ9230","a":"L&R Excavation","l":"Marketing Ctr","b":433.0,"m":384.0,"n":-49.0}]},"Saira Julian Guzman":{"i":[],"d":[{"e":"BSFM0874","a":"Snowfall Towing LLC","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFG4148","a":"Consolidated Insurance Group of South Carolina","l":"Marketing Ctr","b":293.0,"m":0.0,"n":-293.0},{"e":"BSFG4148","a":"Consolidated Insurance Group of South Carolina","l":"Websites","b":117.0,"m":0.0,"n":-117.0},{"e":"BSFG3800","a":"MINT Nutrition","l":"Marketing Ctr","b":293.0,"m":0.0,"n":-293.0},{"e":"BSFQ8706","a":"Sea Pro Home Renovation","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BHCM4169","a":"North Valley Veterinary Clinic","l":"Marketing Ctr","b":192.4,"m":0.0,"n":-192.4}]},"Taylor Kidd":{"i":[],"d":[{"e":"BSFD0246","a":"Its A Beautiful World Travel","l":"SEO","b":800.0,"m":0.0,"n":-800.0},{"e":"BSFD0246","a":"Its A Beautiful World Travel","l":"Marketing Ctr","b":433.0,"m":0.0,"n":-433.0},{"e":"BSFP6023","a":"Mindsense Serenity","l":"Marketing Ctr","b":293.0,"m":49.0,"n":-244.0},{"e":"BDST1459","a":"Coweta Smiles","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BDST1459","a":"Coweta Smiles","l":"SEO","b":700.0,"m":0.0,"n":-700.0}]},"Barbara Larrosa Presinal":{"i":[],"d":[{"e":"BRXV1840","a":"Phillip Brown Construction LLC","l":"Growth Pkgs","b":150.0,"m":0.0,"n":-150.0},{"e":"BSFF6276","a":"Adams Plumbing","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFD7034","a":"Bossert Bookkeeping LLC","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFG1609","a":"Noah Auto Sales","l":"Social","b":750.0,"m":0.0,"n":-750.0},{"e":"BRXG6596","a":"Budget HVAC","l":"SEO","b":700.0,"m":0.0,"n":-700.0}]},"Kellie Lester":{"i":[],"d":[{"e":"BRXJ0339","a":"Able Roofing Company","l":"Marketing Ctr","b":433.0,"m":0.0,"n":-433.0},{"e":"BRXJ0339","a":"Able Roofing Company","l":"Growth Pkgs","b":300.0,"m":0.0,"n":-300.0},{"e":"BSFM3683","a":"Shine Effect Cleaning Services","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFM3683","a":"Shine Effect Cleaning Services","l":"Growth Pkgs","b":500.0,"m":0.0,"n":-500.0},{"e":"BSDZ8739","a":"Rc Home Renovations","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSDZ8739","a":"Rc Home Renovations","l":"Websites","b":117.0,"m":0.0,"n":-117.0},{"e":"BSFF4291","a":"Brighter Light Media","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFD5506","a":"WaterWise Plumbing","l":"Websites","b":117.0,"m":0.0,"n":-117.0},{"e":"BRXC8972","a":"Mary El Coiffures","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0}]},"Juan Liberato":{"i":[],"d":[{"e":"BSFF2981","a":"Redding Salon","l":"Social","b":750.0,"m":0.0,"n":-750.0}]},"Sakshi Mahalwal":{"i":[{"e":"BSDL5682","a":"Plain Pallets Pty Ltd","l":"Marketing Ctr","b":499.09,"m":509.09,"n":10.0}],"d":[{"e":"BSDZ6167","a":"Millers Civil Contractors","l":"Thryv Leads","b":454.0,"m":0.0,"n":-454.0},{"e":"BSFF9155","a":"Bristols Automotive Specialists","l":"Marketing Ctr","b":407.0,"m":0.0,"n":-407.0},{"e":"BSFG0909","a":"P & D Homes","l":"Marketing Ctr","b":351.82,"m":0.0,"n":-351.82},{"e":"BSCZ1631","a":"CARPETWEST LIMITED","l":"Marketing Ctr","b":318.0,"m":0.0,"n":-318.0},{"e":"BSCZ6848","a":"ASL Industries Ltd","l":"Marketing Ctr","b":330.0,"m":318.0,"n":-12.0},{"e":"BSFF2333","a":"Grant Financial Consulting","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSDX0816","a":"Bell Lawyers","l":"Social","b":1500.0,"m":0.0,"n":-1500.0},{"e":"BSFM0696","a":"Agape Sanitation","l":"Websites","b":131.82,"m":0.0,"n":-131.82},{"e":"BSFM0696","a":"Agape Sanitation","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0}]},"Peter Manalac":{"i":[],"d":[{"e":"BSFC6308","a":"My Antenna Tech","l":"Websites","b":131.82,"m":0.0,"n":-131.82},{"e":"BSFC6308","a":"My Antenna Tech","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSDK5917","a":"Beaufort Newsagent","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSDP9748","a":"Oceania Engineering Services","l":"Websites","b":131.82,"m":0.0,"n":-131.82},{"e":"BSDP9748","a":"Oceania Engineering Services","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSFS8025","a":"Smartline Electrical","l":"Websites","b":131.82,"m":0.0,"n":-131.82},{"e":"BSFG0578","a":"Jerry's Paintless Dent Repair","l":"Thryv Leads","b":454.0,"m":0.0,"n":-454.0},{"e":"BSFG0578","a":"Jerry's Paintless Dent Repair","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSDP7478","a":"DTS Electrical","l":"Marketing Ctr","b":318.0,"m":0.0,"n":-318.0}]},"Scott Mather":{"i":[],"d":[{"e":"BSFP7146","a":"Forever Glam Scrubs and More Murrieta","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSDL2334","a":"Local Exterior Services","l":"Websites","b":117.0,"m":0.0,"n":-117.0}]},"Wilson Mercedes":{"i":[],"d":[{"e":"BSCM8875","a":"Dry Line LLC","l":"Marketing Ctr","b":293.0,"m":0.0,"n":-293.0},{"e":"BSFF5372","a":"Camera Creations Photography","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFD8980","a":"Azteca Taxi Cab","l":"Social","b":500.0,"m":0.0,"n":-500.0},{"e":"BSFF8453","a":"1 and Done Legal Document Assistance","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFD0345","a":"Boat Stuf","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFZ1197","a":"Joannie\u2019s Florals and Events","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFZ1197","a":"Joannie\u2019s Florals and Events","l":"Growth Pkgs","b":300.0,"m":0.0,"n":-300.0},{"e":"BSFF8709","a":"Andrew's Violin and Viola Music Lessons By Appointment","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSBX9512","a":"Blackwell Agency Inc","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFP3203","a":"Fairfield house cleaning","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSDZ8288","a":"S.P.E.C Training Program LLC","l":"Platform Other","b":21.6,"m":0.0,"n":-21.6},{"e":"BSDZ8288","a":"S.P.E.C Training Program LLC","l":"Business Ctr","b":410.0,"m":0.0,"n":-410.0}]},"Tyler Moeggenberg":{"i":[],"d":[{"e":"BSFM3106","a":"Kevin Yul Wright, JD - Business Loan Success Academy Inc","l":"Websites","b":117.0,"m":0.0,"n":-117.0},{"e":"BSFM3106","a":"Kevin Yul Wright, JD - Business Loan Success Academy Inc","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSCS3567","a":"Healthier you counseling center","l":"Social","b":500.0,"m":0.0,"n":-500.0},{"e":"BSFP6950","a":"Property Management Remote Staffing","l":"Marketing Ctr","b":433.0,"m":0.0,"n":-433.0},{"e":"BSFP6950","a":"Property Management Remote Staffing","l":"Social","b":1000.0,"m":0.0,"n":-1000.0},{"e":"BSFP6950","a":"Get Staffing","l":"Marketing Ctr","b":293.0,"m":0.0,"n":-293.0},{"e":"BSFP6950","a":"Get Staffing","l":"Social","b":1000.0,"m":0.0,"n":-1000.0},{"e":"BSFP6950","a":"HYPR Staffing","l":"Marketing Ctr","b":433.0,"m":0.0,"n":-433.0},{"e":"BSFP6950","a":"HYPR Callers","l":"Marketing Ctr","b":293.0,"m":0.0,"n":-293.0},{"e":"BSFM3540","a":"JD's One Stop Auto Shop","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFB0193","a":"Christine Simper","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF4901","a":"Howard County ALF","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF4901","a":"Howard County ALF","l":"Social","b":750.0,"m":0.0,"n":-750.0},{"e":"BBCH4580","a":"Granite State Home Improvements","l":"Marketing Ctr","b":628.0,"m":384.0,"n":-244.0}]},"Irina Larianni Molina Molina":{"i":[],"d":[{"e":"BSFD1222","a":"The Ranch Garage Doors NC","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFC0954","a":"Locksmith Citrus Florida","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF6588","a":"Genesis Landworks LLC","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFC8762","a":"Joint Sealant & Waterproofing LLC","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFC4055","a":"Brewhaus America, Inc.","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFD0442","a":"Jack's Plumbing","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFG1730","a":"DAB Digital Enterprises, LLC","l":"Marketing Ctr","b":293.0,"m":244.0,"n":-49.0},{"e":"BSFG1730","a":"DAB Digital Enterprises, LLC","l":"Social","b":500.0,"m":0.0,"n":-500.0}]},"Yessica Montero Urena":{"i":[{"e":"BBJS3449","a":"Wood Well Drilling LLC","l":"Websites","b":117.0,"m":125.0,"n":8.0}],"d":[{"e":"BBXQ7972","a":"PICK OF THE LITTER PAINTING","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFF9513","a":"Healthy Home Enterprises LLC","l":"Marketing Ctr","b":293.0,"m":0.0,"n":-293.0},{"e":"BSFF9513","a":"Healthy Home Enterprises LLC","l":"Platform Other","b":7.2,"m":0.0,"n":-7.2},{"e":"BSFX7950","a":"Elizbiz Limited Liability Company","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BJWQ7551","a":"Kirch Edward Property Management Inc","l":"Websites","b":117.0,"m":0.0,"n":-117.0},{"e":"BSFP7368","a":"Tropical Sunbeds Tanning Salon","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0}]},"Victor Abner Moscoso Fernandez":{"i":[],"d":[{"e":"BHBX0819","a":"North Bay Painting","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF7006","a":"JACKSON & SON CONSTRUCTION LLC","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFF4377","a":"Sheltons Janitorial","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFG1957","a":"Marz landscaping","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFX4887","a":"A Reed Consulting Inc.","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSDF5952","a":"Elquin Tree Service Inc.","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSDD7006","a":"Perfectly Placed Home Organizing","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0}]},"Ellise Payne":{"i":[{"e":"BSFD7073","a":"Bay City Demolition","l":"Websites","b":131.82,"m":136.36,"n":4.54}],"d":[{"e":"BSFC6433","a":"Carpet One Lithgow Pty Ltd.","l":"Growth Pkgs","b":272.73,"m":0.0,"n":-272.73},{"e":"BSDM5923","a":"The Battery Terminal","l":"Social","b":1525.0,"m":0.0,"n":-1525.0},{"e":"BSFF7767","a":"DocSmart Solutions Limited","l":"Marketing Ctr","b":318.0,"m":0.0,"n":-318.0},{"e":"BSFD6631","a":"HAIR @ THE HUB","l":"Marketing Ctr","b":499.09,"m":0.0,"n":-499.09},{"e":"BSFF2320","a":"BeltUp Leather Co.","l":"Thryv Leads","b":3636.0,"m":0.0,"n":-3636.0},{"e":"BSFF9360","a":"Easy Excavators","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSDZ5486","a":"AOF Space Design","l":"Social","b":750.0,"m":0.0,"n":-750.0}]},"Sati Ananda Pimentel Malespin":{"i":[],"d":[{"e":"BSFC7529","a":"Acclaim Homecare Svc, LLC","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF3087","a":"John P. Burgess, DDS","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFV5376","a":"Georgia Studs Construction","l":"Business Ctr","b":179.0,"m":0.0,"n":-179.0},{"e":"BSFV5376","a":"Georgia Studs Construction","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFC8834","a":"Vision Home Improvement LLC","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFC8834","a":"Vision Home Improvement LLC","l":"Business Ctr","b":179.0,"m":0.0,"n":-179.0},{"e":"BFMP0738","a":"Quality Garage Doors","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0}]},"Tyler Popplewell":{"i":[],"d":[{"e":"BSFQ9150","a":"Four Girls and a Dream","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFQ9150","a":"Four Girls and a Dream","l":"Business Ctr","b":179.0,"m":0.0,"n":-179.0}]},"Yolanda Ramirez":{"i":[{"e":"BSFX4907","a":"Gulf Wind Cigars","l":"Websites","b":117.0,"m":125.0,"n":8.0}],"d":[{"e":"BSFB8524","a":"national awnings Miami","l":"Thryv Leads","b":1500.0,"m":0.0,"n":-1500.0},{"e":"BSFB8524","a":"national awnings Miami","l":"Social","b":750.0,"m":0.0,"n":-750.0},{"e":"BSFD2514","a":"D.I. Ready Cleaning Service Inc.","l":"Thryv Leads","b":1800.0,"m":0.0,"n":-1800.0},{"e":"BJQR3935","a":"Star Landscaping","l":"Thryv Leads","b":1500.0,"m":1000.0,"n":-500.0},{"e":"BRRZ3102","a":"Mr. Plumber","l":"Thryv Leads","b":3600.0,"m":2100.0,"n":-1500.0}]},"Alejandro Rodriguez-Medina":{"i":[{"e":"BSFD2501","a":"Plumbing Kingz LLC","l":"Business Ctr","b":208.0,"m":237.0,"n":29.0}],"d":[{"e":"BSFC8877","a":"Johnson Asphalt Paving","l":"Social","b":500.0,"m":0.0,"n":-500.0},{"e":"BSFF2943","a":"Jesus Knows My Name","l":"Websites","b":117.0,"m":0.0,"n":-117.0},{"e":"BSFF2943","a":"Jesus Knows My Name","l":"Business Ctr","b":179.0,"m":0.0,"n":-179.0},{"e":"BSFZ3549","a":"Trees Plus Inc","l":"SEO","b":500.0,"m":0.0,"n":-500.0},{"e":"BDTV2624","a":"Portage Auto Mall Complete Auto Repair Center","l":"Growth Pkgs","b":500.0,"m":0.0,"n":-500.0},{"e":"BDTV2624","a":"Portage Auto Mall Complete Auto Repair Center","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFC3508","a":"Clean It All","l":"Websites","b":117.0,"m":0.0,"n":-117.0},{"e":"BSFC3508","a":"Clean It All","l":"Business Ctr","b":244.0,"m":0.0,"n":-244.0}]},"Stacy Roers":{"i":[],"d":[{"e":"BSFZ1687","a":"Hypnosis at Kentic Healing Connection","l":"Marketing Ctr","b":293.0,"m":0.0,"n":-293.0},{"e":"BSFB5606","a":"Colonial Loan","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFB5611","a":"Colonial Loan","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BRXT2163","a":"Screens Only","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0}]},"Zoltan Rudolf":{"i":[],"d":[{"e":"BSFF2764","a":"Kerbing By Design","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSCZ2971","a":"Daveron Scaffolding Ltd","l":"Marketing Ctr","b":318.0,"m":0.0,"n":-318.0},{"e":"BSFD9058","a":"Eza Plumbing and Drainage","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSFD8834","a":"Mowbray Newsagency","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSFD7043","a":"Smile More Dental","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSDM9422","a":"Productive Plastics","l":"Thryv Leads","b":4545.0,"m":0.0,"n":-4545.0},{"e":"BSDW3326","a":"CBI Blinds","l":"Thryv Leads","b":909.0,"m":0.0,"n":-909.0},{"e":"BSDW3326","a":"CBI Blinds","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0}]},"Kennedy Sanchez":{"i":[{"e":"BRXD6298","a":"A-1 Driveway Replacement","l":"Thryv Leads","b":3000.0,"m":4000.0,"n":1000.0}],"d":[{"e":"BRXH7873","a":"Gaddis Nursery Inc.","l":"Growth Pkgs","b":100.0,"m":0.0,"n":-100.0},{"e":"BRXH7873","a":"Gaddis Nursery Inc.","l":"Marketing Ctr","b":69.4,"m":0.0,"n":-69.4},{"e":"BSFF6990","a":"Romano's Painting LLC","l":"Websites","b":117.0,"m":0.0,"n":-117.0},{"e":"BSFF6990","a":"Romano's Painting LLC","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF6990","a":"Romano's Painting LLC","l":"Social","b":500.0,"m":0.0,"n":-500.0},{"e":"BRWS6514","a":"4 C's Construction","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BRWS6514","a":"4 C's Construction","l":"Growth Pkgs","b":300.0,"m":0.0,"n":-300.0},{"e":"BSFR0035","a":"Bloomhouse Landscape and Irrigation","l":"Thryv Leads","b":2100.0,"m":0.0,"n":-2100.0},{"e":"BSFP2855","a":"TB Roadside Assistance","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFF2724","a":"Florida Luxury Home and Condo","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFF2724","a":"Florida Luxury Home and Condo","l":"Platform Other","b":7.2,"m":0.0,"n":-7.2}]},"Darling Danais Santos Taveras":{"i":[],"d":[{"e":"BSFB2359","a":"Stor It Here Storage","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF6200","a":"Botail by Jeffrey","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFB3954","a":"Davinci Valentin Brand","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF0871","a":"New England Auctions","l":"Marketing Ctr","b":293.0,"m":244.0,"n":-49.0}]},"Steven Saunders":{"i":[],"d":[{"e":"BSFM1190","a":"D & M Asphalt Services, Inc","l":"Social","b":750.0,"m":0.0,"n":-750.0},{"e":"BSFP3972","a":"LM Bodywork Institute","l":"Websites","b":117.0,"m":0.0,"n":-117.0},{"e":"BSFP3972","a":"LM Bodywork Institute","l":"Business Ctr","b":179.0,"m":0.0,"n":-179.0},{"e":"BSFC4006","a":"Floridas First Coast Of Golf","l":"Business Ctr","b":244.0,"m":0.0,"n":-244.0}]},"Rafael Sencion Sencion":{"i":[],"d":[{"e":"BSFP8938","a":"Toro Birria","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFP8938","a":"Toro Birria","l":"Growth Pkgs","b":300.0,"m":0.0,"n":-300.0},{"e":"BSFF7702","a":"Quantum Sports Center","l":"Social","b":750.0,"m":0.0,"n":-750.0},{"e":"BSFF7702","a":"Quantum Sports Center","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSCL3353","a":"REPAIR APPLIANCE TECH","l":"Growth Pkgs","b":300.0,"m":0.0,"n":-300.0},{"e":"BSCL3353","a":"REPAIR APPLIANCE TECH","l":"Marketing Ctr","b":433.0,"m":0.0,"n":-433.0},{"e":"BSFF5647","a":"JJ Cillis Group LLC","l":"Marketing Ctr","b":293.0,"m":0.0,"n":-293.0},{"e":"BSFF5647","a":"JJ Cillis Group LLC","l":"Websites","b":117.0,"m":0.0,"n":-117.0},{"e":"BSFP4560","a":"Technology Networks","l":"Marketing Ctr","b":293.0,"m":244.0,"n":-49.0},{"e":"BSFP4560","a":"Technology Networks","l":"Social","b":500.0,"m":0.0,"n":-500.0},{"e":"BSBN3109","a":"Tolberts Heating & Cooling LLC","l":"Social","b":500.0,"m":0.0,"n":-500.0},{"e":"BSFD2419","a":"Sunterjee LLC","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0}]},"Ashley Shaffer":{"i":[],"d":[{"e":"BSFF5411","a":"Stretch-abilitation","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF5411","a":"Stretch-abilitation","l":"SEO","b":1300.0,"m":0.0,"n":-1300.0},{"e":"BSFF5411","a":"Stretch-abilitation","l":"Social","b":500.0,"m":0.0,"n":-500.0},{"e":"BRXF3371","a":"Care Medical Center","l":"Social","b":1500.0,"m":0.0,"n":-1500.0},{"e":"BSCW6209","a":"Douglas Water Depot","l":"Thryv Leads","b":1000.0,"m":0.0,"n":-1000.0},{"e":"BRWN8733","a":"West Oaks Animal Hospital LLC","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BRXV6313","a":"Mauldin Trash Service","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BRXV6313","a":"Mauldin Trash Service","l":"SEO","b":1200.0,"m":0.0,"n":-1200.0}]},"Nikita Siepen-Bowers":{"i":[{"e":"BSFD8077","a":"Action - Priority","l":"Websites","b":131.82,"m":136.36,"n":4.54},{"e":"BSDV1069","a":"Hawkesbury City Plumbing Pty Ltd","l":"Marketing Ctr","b":111.53,"m":114.88,"n":3.35}],"d":[{"e":"BSFD8077","a":"Action - Priority","l":"Thryv Leads","b":909.0,"m":0.0,"n":-909.0},{"e":"BSFX6219","a":"Leakend Leak Finding Services","l":"Marketing Ctr","b":286.37,"m":280.0,"n":-6.37},{"e":"BSDW7306","a":"Giraffe Removals","l":"Marketing Ctr","b":351.82,"m":0.0,"n":-351.82},{"e":"BSDW7306","a":"Giraffe Removals","l":"Thryv Leads","b":1818.0,"m":0.0,"n":-1818.0},{"e":"BSDW7306","a":"Giraffe Removals","l":"SEO","b":1200.0,"m":0.0,"n":-1200.0},{"e":"BSFD7064","a":"Total Vision Pool Fencing","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSDM0061","a":"The Squeegees Window Cleaning","l":"Business Ctr","b":320.0,"m":0.0,"n":-320.0},{"e":"BSFB5778","a":"NZ Natural Formulas","l":"Websites","b":29.0,"m":0.0,"n":-29.0},{"e":"BSFB5778","a":"NZ Natural Formulas","l":"SEO","b":1750.0,"m":0.0,"n":-1750.0},{"e":"BSFB5778","a":"NZ Natural Formulas","l":"Marketing Ctr","b":318.0,"m":0.0,"n":-318.0},{"e":"BSFB5778","a":"NZ Natural Formulas","l":"Social","b":1650.0,"m":0.0,"n":-1650.0},{"e":"BSDB3382","a":"G&M NODDING LIMITED","l":"Marketing Ctr","b":318.0,"m":0.0,"n":-318.0},{"e":"BSDM1772","a":"Wizard Motors Pty Ltd","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSFC6732","a":"Mr Hook Towing and Metal","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSDL0232","a":"Williams Premium Wholesale","l":"Marketing Ctr","b":509.1,"m":499.09,"n":-10.01},{"e":"BSFQ6729","a":"BRP Construction limited","l":"Social","b":825.0,"m":0.0,"n":-825.0},{"e":"BSFQ6729","a":"BRP Construction limited","l":"Marketing Ctr","b":407.0,"m":0.0,"n":-407.0}]},"Sarah Swanson":{"i":[{"e":"BSFV8753","a":"24/7 Events","l":"Thryv Leads","b":2000.0,"m":4800.0,"n":2800.0}],"d":[{"e":"BSCT7730","a":"Elemental Landscaping Inc","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFB9274","a":"masterwealthbuildersllc","l":"Business Ctr","b":254.0,"m":0.0,"n":-254.0},{"e":"BSFB9274","a":"masterwealthbuildersllc","l":"Websites","b":117.0,"m":0.0,"n":-117.0},{"e":"BRDX7750","a":"Paradise Landscaping and Tree Service","l":"Social","b":750.0,"m":0.0,"n":-750.0},{"e":"BRDX7750","a":"Paradise Landscaping and Tree Service","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0}]},"Matt Sword":{"i":[],"d":[{"e":"BSFQ2287","a":"Levels Ahead Painting","l":"Social","b":750.0,"m":0.0,"n":-750.0},{"e":"BSFQ2287","a":"Levels Ahead Painting","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFM3452","a":"AllTime Lock Out Service","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFM3452","a":"AllTime Lock Out Service","l":"SEO","b":400.0,"m":0.0,"n":-400.0},{"e":"BSFF2470","a":"Bella Roma pasta and pizza","l":"Growth Pkgs","b":500.0,"m":0.0,"n":-500.0},{"e":"BSFF6379","a":"A Family for Every Child","l":"Marketing Ctr","b":293.0,"m":0.0,"n":-293.0},{"e":"BSDZ9849","a":"Makeup by Jaundalyn","l":"Business Ctr","b":410.0,"m":0.0,"n":-410.0},{"e":"BSFP3264","a":"Shine Bright Cleaning Solutions","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFP3264","a":"Shine Bright Cleaning Solutions","l":"Platform Other","b":7.2,"m":0.0,"n":-7.2}]},"Elianny Tena Antigua":{"i":[],"d":[{"e":"BSFR6524","a":"Financial Grafix LLP","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFR6524","a":"Financial Grafix LLP","l":"Growth Pkgs","b":300.0,"m":0.0,"n":-300.0}]},"Heidi Torres Uribe":{"i":[{"e":"BSDR4708","a":"Triple Clean","l":"Platform Other","b":36.0,"m":50.4,"n":14.4}],"d":[{"e":"BSFF4402","a":"Green Year Landscaping","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BRWV6069","a":"Custom Decks Plus","l":"Websites","b":64.9,"m":0.0,"n":-64.9},{"e":"BSFF4269","a":"Airport Towing","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFF3693","a":"Cold Spring School","l":"Social","b":750.0,"m":0.0,"n":-750.0},{"e":"BSFF3693","a":"Cold Spring School","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF9591","a":"Helping Hands Estate Sales","l":"Social","b":500.0,"m":0.0,"n":-500.0},{"e":"BSFF9591","a":"Helping Hands Estate Sales","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFF9591","a":"Helping Hands Estate Sales","l":"Platform Other","b":7.2,"m":0.0,"n":-7.2},{"e":"BSDR4708","a":"Triple Clean","l":"Business Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFV9017","a":"Premier Patio & Deck","l":"SEO","b":1900.0,"m":0.0,"n":-1900.0},{"e":"BSFD9971","a":"MBI International, LLC","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0}]},"Karmita Turner":{"i":[],"d":[{"e":"BFPM2056","a":"ARTchitectural","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BQLC8330","a":"Oakridge Roofing Solutions LLC","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFP8430","a":"K & K Solutions Inc","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0}]},"Rossi Valerio Tejeda":{"i":[],"d":[{"e":"BSFZ2573","a":"Empire Works GA","l":"Marketing Ctr","b":400.0,"m":384.0,"n":-16.0}]},"Ashley Vasquez Mena":{"i":[],"d":[{"e":"BNNC5836","a":"Palm Family Eyecare","l":"Marketing Ctr","b":135.9,"m":86.9,"n":-49.0},{"e":"BNNC5836","a":"Palm Family Eyecare","l":"Growth Pkgs","b":300.0,"m":0.0,"n":-300.0},{"e":"BBCL3929","a":"Highland Farm","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFD8550","a":"Vanguard Fence & Deck LLC","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF9026","a":"I-ONE CONSTRUCTION","l":"Marketing Ctr","b":433.0,"m":0.0,"n":-433.0},{"e":"BSFF9026","a":"I-ONE CONSTRUCTION","l":"Social","b":750.0,"m":0.0,"n":-750.0},{"e":"BSFP7091","a":"AZA Drywall and Remodeling","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BSFD1240","a":"Truflo Plumbing Solutions LLC","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF5601","a":"Ready Golf 4K","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF5601","a":"Ready Golf 4K","l":"Social","b":500.0,"m":0.0,"n":-500.0},{"e":"BRXB5124","a":"Carrie's Cash 4 Your Car","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0},{"e":"BRXB5124","a":"Carrie's Cash 4 Your Car","l":"Thryv Leads","b":1500.0,"m":0.0,"n":-1500.0}]},"Mark Velazquez":{"i":[],"d":[{"e":"BSFQ3467","a":"CJ Construction","l":"Marketing Ctr","b":433.0,"m":0.0,"n":-433.0},{"e":"BSFP7148","a":"Loving My Pregnancy","l":"Social","b":500.0,"m":0.0,"n":-500.0},{"e":"BSFP7148","a":"Loving My Pregnancy","l":"Growth Pkgs","b":500.0,"m":0.0,"n":-500.0},{"e":"BSFF9272","a":"Guirian Bookkeeping Services","l":"Marketing Ctr","b":384.0,"m":0.0,"n":-384.0}]},"Indu Vijay":{"i":[{"e":"BSDD3552","a":"Hodge motors","l":"Websites","b":155.0,"m":160.0,"n":5.0}],"d":[{"e":"BSFD8658","a":"Safeguard Tree Services","l":"Marketing Ctr","b":280.0,"m":0.0,"n":-280.0},{"e":"BSFF3589","a":"JP Air Conditioning - Heating & Cooling","l":"Marketing Ctr","b":499.09,"m":0.0,"n":-499.09}]},"Anthony Yen":{"i":[{"e":"BSFR8712","a":"Inexpensive Tree Care","l":"Thryv Leads","b":2500.0,"m":3000.0,"n":500.0},{"e":"BSBQ1871","a":"AG Builders LLC","l":"Websites","b":117.0,"m":125.0,"n":8.0},{"e":"BSDL7961","a":"Luxury Bubbles","l":"Thryv Leads","b":4000.0,"m":6000.0,"n":2000.0}],"d":[{"e":"BSFR4557","a":"Total Pest Control","l":"Thryv Leads","b":1300.0,"m":0.0,"n":-1300.0},{"e":"BSFX8681","a":"Angels of Mercy Private Homecare Services","l":"Marketing Ctr","b":400.0,"m":384.0,"n":-16.0},{"e":"BSGB1783","a":"Email Marketing Group Inc","l":"Social","b":750.0,"m":0.0,"n":-750.0},{"e":"BSGB1783","a":"Email Marketing Group Inc","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BSFF4619","a":"Maxine Lawson- USA Benefits Group","l":"Marketing Ctr","b":244.0,"m":0.0,"n":-244.0},{"e":"BRXS2753","a":"HL Holmes Asphalt Paving","l":"Platform Other","b":7.2,"m":0.0,"n":-7.2},{"e":"BSFF5680","a":"C&J Drain Cleaning","l":"Websites","b":117.0,"m":0.0,"n":-117.0},{"e":"BRXP5453","a":"American West Services","l":"Marketing Ctr","b":244.0,"m":228.0,"n":-16.0},{"e":"BSDD8734","a":"The Event Gallery","l":"Business Ctr","b":179.0,"m":0.0,"n":-179.0}]}};

function BobView({filterCoach, filterCSM, managerCoaches, bobRaw, mcChurn, bcChurn, churnAlerts, onSelectCSM, liveBobDet={}, bobAdj={}, q3Current={}, q3Log=[], q3BobBoq=[], q3BobCur=[], domoBoq=[], q3Supp=[]}) {
  const [bobSubTab, setBobSubTab] = useState("current"); // "current" | "q3"
  const hasQ3 = Object.keys(q3Current).length > 0;
  const getDet = n => {
    const base = liveBobDet[n]||liveBobDet[norm(n)]||BOB_DETAIL[n]||BOB_DETAIL[norm(n)]||{};
    const adjKey = Object.keys(bobAdj).find(k=>norm(k)===n||k===n);
    if (!adjKey) return base;
    const adj = bobAdj[adjKey];
    const extraInc = adj.entries.filter(e=>e.n>0).map(e=>({...e, b:0, m:e.n, _adj:true}));
    const extraDec = adj.entries.filter(e=>e.n<0).map(e=>({...e, b:0, m:e.n, _adj:true}));
    // Adjustments are tracked separately and DON'T modify BOQ, current, or any retention math —
    // they show up as their own "Adjustment" row for visibility only, always displayed as a
    // positive tracked amount regardless of the sign entered on the submission form.
    // The adjustment credit already flows into the CSM's top-level Current MRR / retention via
    // lcmDelta. So the adjustment row here also needs m (current MRR) set to the credit amount —
    // otherwise summing account-level rows (in the expand view and CSV export) won't match the
    // top-line Ending Book / Current MRR figure, which double-counts the credit only at one level.
    const adjRows = adj.entries.map(a => ({...a, b:0, m:a.n, n:a.n, status:"adjustment", adjAmount:a.n}));
    return {
      ...base,
      i: [...(base.i||[]), ...extraInc],
      d: [...(base.d||[]), ...extraDec],
      all: [...(base.all||[]), ...adjRows],
    };
  };
  // Live sheet data when available, hardcoded fallback otherwise
  const liveCoachTotals = (bobRaw && Object.keys(bobRaw.coachTotals||{}).length > 0) ? bobRaw.coachTotals : BOB_COACH_TOTALS;
  const liveGrand       = (bobRaw && bobRaw.grand) ? bobRaw.grand : BOB_GRAND;
  // Build CSM list from live BOB data; churn derived from liveBobDet (BOQ>0, LCM=0)
  const liveCsms = (bobRaw && Object.keys(bobRaw.bob||{}).length > 0)
    ? Object.entries(bobRaw.bob).map(([name, d]) => {
        const det = getDet(name);
        const churned = det.churned || [];
        // unique account names with products for pill display
        const chmch = churned.map(a => a.name);
        const adjKey2 = Object.keys(bobAdj).find(k=>norm(k)===name||k===name);
        const lcmAdj = adjKey2 ? bobAdj[adjKey2].lcmDelta : 0;
        const adjLcm = (d.lcm||0) + lcmAdj;
        return {n:name, c:d.coach||"", boq:d.boq, lcm:adjLcm, net:(d.net||0)+lcmAdj,
          ret:d.boq>0?adjLcm/d.boq:d.ret,
          mcc:churned.length, mch:chmch, churned,
          mca:0, bcc:0, bca:0, bch:[],
          hasAdj: lcmAdj!==0, adjAmount: lcmAdj};
      }).filter(c => c.boq > 0)
    : BOB_CSMS.map(c => {
        const det = getDet(c.n);
        const churned = det.churned || [];
        return {...c,
          ret:c.boq>0&&c.lcm!=null?c.lcm/c.boq:c.ret,
          mcc:churned.length||c.mcc, mch:churned.length?churned.map(a=>a.name):c.mch,
          churned, bcc:0, bch:[]};
      });
  const [bobTab, setBobTab]         = useState("overview");
  const [q3Sort,       setQ3Sort]       = useState({col:"retPct", dir:"asc"});
  const [tileFilter,   setTileFilter]   = useState(null);
  const [q3CSMFilter,  setQ3CSMFilter]  = useState(null);
  const [acctSort,     setAcctSort]     = useState({col:"acct", dir:"asc"});
  const [q2Sort,       setQ2Sort]       = useState({col:"retPct", dir:"asc"});
  const [q2TileFilter, setQ2TileFilter] = useState(null);
  const [q2CSMFilter,  setQ2CSMFilter]  = useState(null);
  const [q2AcctSort,   setQ2AcctSort]   = useState({col:"acct", dir:"asc"});
  // Domo Book of Business tab state — mirrors Q2/Q3 Tracking pattern
  const [domoSort,       setDomoSort]       = useState({col:"retPct", dir:"asc"});
  const [domoAcctSort,   setDomoAcctSort]   = useState({col:"acct", dir:"asc"});
  const [domoCSMFilter,  setDomoCSMFilter]  = useState(null);   // expanded CSM
  const [domoTileFilter, setDomoTileFilter] = useState(null);   // clicked tile: increase/decrease/cancelled/no_data
  const [churnModal, setChurnModal] = useState(false);
  const [bobSort, setBobSort]       = useState({col:"ret", dir:"desc"});
  const [expandedBob, setExpandedBob] = useState(null);

  const GOAL = BOB_GOAL;
  const fmt$ = n => n==null?"--":"$"+Math.abs(n).toLocaleString("en-US",{minimumFractionDigits:0,maximumFractionDigits:0});
  const fmtP = p => p==null?"--":(p*100).toFixed(1)+"%";
  const rCol = p => p>=GOAL?"#16a34a":p>=0.85?"#d97706":"#dc2626";
  const rBg  = p => p>=GOAL?"rgba(22,163,74,.1)":p>=0.85?"rgba(217,119,6,.1)":"rgba(220,38,38,.1)";
  const rFg  = p => p>=GOAL?"#166534":p>=0.85?"#854d0e":"#991b1b";

  const filtCSMs = () => {
    let list = liveCsms.filter(c=>c.boq>0);
    if (managerCoaches||filterCoach) {
      list = list.filter(c=>{
        // 1. Try ROSTER lookup by CSM name (most reliable)
        const i = lk(c.n);
        if (i && i.c) {
          if (managerCoaches && !managerCoaches.includes(i.c)) return false;
          if (filterCoach && i.c !== filterCoach) return false;
          return true;
        }
        // 2. Fall back: match stored coach name to COACHES array
        //    Normalize both sides — strip smart quotes, extra spaces, case
        const normalize = s => (s||"").toLowerCase().trim()
          .replace(/[‘’‚‛′‵']/g,"'")
          .replace(/[“”„‟″‶"]/g,'"');
        const storedCoach = normalize(c.c);
        const found = COACHES.find(ch => normalize(ch.n) === storedCoach);
        const coachEmail = found ? found.e : null;
        if (managerCoaches && !managerCoaches.includes(coachEmail)) return false;
        if (filterCoach && coachEmail !== filterCoach) return false;
        return true;
      });
    }
    if (filterCSM) list = list.filter(c=>c.n===filterCSM);
    return [...list].sort((a,b)=>{
      const av=a[bobSort.col]??-999, bv=b[bobSort.col]??-999;
      return bobSort.dir==="desc"?bv-av:av-bv;
    });
  };

  const csms = filtCSMs();
  const totBoq = csms.reduce((s,c)=>s+c.boq,0);
  const totLcm = csms.reduce((s,c)=>s+(c.lcm||0),0);
  const totNet = csms.reduce((s,c)=>s+(c.net||0),0);
  const avgRet = totBoq>0 ? totLcm/totBoq : 0;
  const totChurned = csms.reduce((s,c)=>s+(c.mcc||0),0);
  const totMcc = totChurned; // kept for compat
  const totBcc = 0;

  const thS = {fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,padding:"0 8px 8px 0",textAlign:"left",borderBottom:"0.5px solid rgba(41,53,93,.08)",cursor:"pointer",whiteSpace:"nowrap"};
  const thRS= {...thS, textAlign:"right"};
  const tdS = {padding:"8px 8px 8px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",fontSize:12,verticalAlign:"top"};
  const tdRS= {...tdS,textAlign:"right"};

  const thSort = (col,lbl) => (
    <th style={{...thRS,color:bobSort.col===col?"#FF5000":"#808080"}}
      onClick={()=>setBobSort(s=>({col,dir:s.col===col&&s.dir==="desc"?"asc":"desc"}))}>
      {lbl}{bobSort.col===col?(bobSort.dir==="desc"?" ▼":" ▲"):<span style={{color:"#ccc",fontSize:9}}> ↕</span>}
    </th>
  );

  const pill = (p) => (
    <span style={{display:"inline-block",fontSize:10,fontWeight:500,padding:"2px 8px",borderRadius:20,background:rBg(p),color:rFg(p)}}>
      {fmtP(p)}
    </span>
  );

  const barRow = (pct, col) => (
    <div style={{height:8,background:"#ECEEF1",borderRadius:4,overflow:"visible",position:"relative",margin:"4px 0"}}>
      <div style={{height:"100%",borderRadius:4,background:col,width:Math.min(pct*100,108).toFixed(1)+"%"}}/>
      <div style={{position:"absolute",top:-1,bottom:-1,width:1.5,background:"rgba(41,53,93,.25)",left:"91%"}}/>
    </div>
  );

  const coachesVisible = filterCoach ? [filterCoach]
    : filterCSM ? (()=>{ const i=lk(filterCSM); return i&&i.c?[i.c]:Object.keys(liveCoachTotals); })()
    : managerCoaches ? COACHES.filter(c=>managerCoaches.includes(c.e)).map(c=>c.e)
    : Object.keys(liveCoachTotals);

  // Shared churn data — used by Overview tile, Churn detail tab, and the full-list modal
  const churnCSMs = csms.filter(c=>c.mcc>0).sort((a,b)=>b.mcc-a.mcc);
  const allChurned = csms.reduce((s,c)=>s+(c.mcc||0),0);
  const allProducts = [...new Set(csms.flatMap(c=>(c.churned||[]).flatMap(a=>a.products||[])))];
  const allChurnedRows = churnCSMs.flatMap(c =>
    (c.churned||c.mch.map(n=>({name:n,products:[]}))).map(a=>({
      csm: dispName(c.n), coach: c.c.split(" ").slice(-1)[0],
      account: typeof a==="string"?a:a.name,
      products: typeof a==="string"?[]:(a.products||[]),
      eid: typeof a==="string"?"":a.eid||"",
    }))
  );
  const exportChurnCSV = () => {
    const header = ["CSM","Coach","Account","Enterprise ID","Products"];
    const rows = allChurnedRows.map(r=>[r.csm,r.coach,r.account,r.eid,r.products.join("; ")]);
    const csv = [header,...rows].map(r=>r.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(",")).join("\n");
    const a = document.createElement("a");
    a.href = "data:text/csv;charset=utf-8,"+encodeURIComponent(csv);
    a.download = "churned_accounts_"+new Date().toISOString().slice(0,10)+".csv";
    a.click();
  };

  // ── Q2 TRACKING — mirrors Q3 Tracking's tiles/table/expand design exactly,
  //    sourced from the Q2 Beginning/Ending Revenue book of business + change detail ──
  const renderQ2 = () => {
    const q2CSMs = csms; // already filtered by coach/manager/CSM via filtCSMs()

    // Classify each CSM's full account-level book using getDet().all (falls back to i/d if unavailable)
    const q2Data = q2CSMs.map(c => {
      const det = getDet(c.n) || {};
      const all = (det.all && det.all.length)
        ? det.all
        : [
            ...(det.i||[]).map(e=>({...e, status: e.b===0?"net_new":"increase"})),
            ...(det.d||[]).map(e=>({...e, status: e.m===0?"cancelled":"decrease"})),
          ];
      const netNewEntries   = all.filter(r => r.status === "net_new");
      const increaseEntries = all.filter(r => r.status === "increase");
      const cancelledEntries= all.filter(r => r.status === "cancelled");
      const decreaseEntries = all.filter(r => r.status === "decrease");
      const adjustmentEntries = all.filter(r => r.status === "adjustment");
      const netNewMrr    = netNewEntries.reduce((s,e)=>s+e.m, 0);
      const increaseMrr  = increaseEntries.reduce((s,e)=>s+e.n, 0);
      const cancelledMrr = cancelledEntries.reduce((s,e)=>s+e.b, 0);
      const decreaseMrr  = decreaseEntries.reduce((s,e)=>s+Math.abs(e.n), 0);
      const adjustmentMrr= adjustmentEntries.reduce((s,e)=>s+(e.adjAmount||0), 0);
      return {
        name: c.n,
        coach: c.c || "",
        boqAdjusted: c.boq,
        currentMrr:  c.lcm,
        netNewMrr, increaseMrr, cancelledMrr, decreaseMrr, adjustmentMrr,
        removedMrr: 0, // Q2 book of business has no separate "removed from report" tracking
        netNewCount: netNewEntries.length,
        cancelledCount: cancelledEntries.length,
        decreaseCount: decreaseEntries.length,
        adjustmentCount: adjustmentEntries.length,
        removedCount: 0,
        retPct: c.ret,
        // Full account/product-line book — includes unchanged accounts, not just changed ones
        acctRows: all.map(r => ({acct:r.a||"—",eid:r.e,boqMrr:r.b,curMrr:r.m,status:r.status,delta:r.n,adjAmount:r.adjAmount||0})),
      };
    });

    const totalBoqAdj    = q2Data.reduce((s,c)=>s+c.boqAdjusted, 0);
    const totalCurrent   = q2Data.reduce((s,c)=>s+c.currentMrr, 0);
    const totalNetNew    = q2Data.reduce((s,c)=>s+c.netNewMrr, 0);
    const totalIncrease  = q2Data.reduce((s,c)=>s+c.increaseMrr, 0);
    const totalCancelled = q2Data.reduce((s,c)=>s+c.cancelledMrr, 0);
    const totalDecrease  = q2Data.reduce((s,c)=>s+c.decreaseMrr, 0);
    const totalAdjustment= q2Data.reduce((s,c)=>s+c.adjustmentMrr, 0);
    const totalRemoved   = 0;
    const overallRet     = totalBoqAdj > 0 ? totalCurrent / totalBoqAdj : null;

    // Filter CSM table by active tile
    const csmsWithEvent = (type) => {
      if (type === "increase")   return new Set(q2Data.filter(c=>c.increaseMrr>0).map(c=>c.name));
      if (type === "cancelled")  return new Set(q2Data.filter(c=>c.cancelledMrr>0).map(c=>c.name));
      if (type === "decrease")   return new Set(q2Data.filter(c=>c.decreaseMrr>0).map(c=>c.name));
      if (type === "net_new")    return new Set(q2Data.filter(c=>c.netNewMrr>0).map(c=>c.name));
      if (type === "adjustment") return new Set(q2Data.filter(c=>c.adjustmentMrr>0).map(c=>c.name));
      if (type === "removed")    return new Set(); // never populated — Q2 model has no removed-account tracking
      return null;
    };
    const activeCsmSet = q2TileFilter ? csmsWithEvent(q2TileFilter) : null;
    const visibleCSMs  = activeCsmSet ? q2Data.filter(c => activeCsmSet.has(c.name)) : q2Data;

    const sortedCSMs = [...visibleCSMs].sort((a, b) => {
      const dir = q2Sort.dir === "asc" ? 1 : -1;
      const col = q2Sort.col;
      const va = a[col] ?? (col === "name" ? "" : 0);
      const vb = b[col] ?? (col === "name" ? "" : 0);
      if (col === "name") return dir * String(va).localeCompare(String(vb));
      return dir * ((Number(va)||0) - (Number(vb)||0));
    });

    const fmt$   = n => "$"+Number(n||0).toLocaleString("en-US",{minimumFractionDigits:0,maximumFractionDigits:0});
    const fmtPct = p => p!=null ? (p*100).toFixed(1)+"%" : "--";
    const retCol = p => p==null?"#808080":p>=0.91?"#16a34a":p>=0.85?"#d97706":"#dc2626";

    const sortTh2 = (col, label, right) => {
      const active = q2Sort.col === col;
      return (
        <th key={col} onClick={()=>setQ2Sort(s=>({col, dir: s.col===col&&s.dir==="asc"?"desc":"asc"}))}
          style={{padding:"0 8px 8px 0",textAlign:right?"right":"left",fontSize:10,textTransform:"uppercase",
            color:active?"#29355D":"#808080",fontWeight:active?700:500,cursor:"pointer",
            borderBottom:"0.5px solid rgba(41,53,93,.08)",userSelect:"none",whiteSpace:"nowrap"}}>
          {label}{active?(q2Sort.dir==="asc"?" ↑":" ↓"):""}
        </th>
      );
    };

    const tileBtn2 = (label, value, sub, color, filterKey) => {
      const active = q2TileFilter === filterKey;
      return (
        <div key={filterKey} onClick={()=>{ setQ2TileFilter(active?null:filterKey); setQ2CSMFilter(null); }}
          style={{background:active?"#29355D":"#ECEEF1",borderRadius:"0 0 10px 10px",padding:"12px 14px",
            borderTop:"3px solid "+color,cursor:"pointer",transition:"all .15s",
            boxShadow:active?"0 2px 8px rgba(41,53,93,.15)":"none"}}>
          <div style={{fontSize:10,textTransform:"uppercase",color:active?"rgba(255,255,255,.7)":"#808080",fontWeight:500,marginBottom:4}}>{label}</div>
          <div style={{fontSize:22,fontWeight:600,color:active?"#fff":color,lineHeight:1,marginBottom:4}}>{value}</div>
          <div style={{fontSize:10,color:active?"rgba(255,255,255,.6)":"#808080"}}>{sub}</div>
        </div>
      );
    };

    const statusBadge2 = (s) => {
      const cfg = {
        increase: {bg:"#dcfce7",fg:"#166534",label:"Increase"},
        decrease: {bg:"#fef9c3",fg:"#854d0e",label:"Decrease"},
        net_new:  {bg:"#dcfce7",fg:"#166534",label:"Net New"},
        cancelled:{bg:"#fef9c3",fg:"#854d0e",label:"Cancelled"},
        removed:  {bg:"#fee2e2",fg:"#991b1b",label:"Removed"},
        unchanged:{bg:"#f3f4f6",fg:"#6b7280",label:"No Change"},
        adjustment:{bg:"#eff6ff",fg:"#1e40af",label:"Adjustment"},
      }[s]||{bg:"#f3f4f6",fg:"#374151",label:s};
      return <span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:cfg.bg,color:cfg.fg}}>{cfg.label}</span>;
    };

    const exportQ2CSV = () => {
      const rows = sortedCSMs.flatMap(c => c.acctRows
        .filter(r => !q2TileFilter || r.status === q2TileFilter)
        .map(r => ({ csm: c.name, coach: c.coach, ...r }))
      );
      if (!rows.length) return;
      const headers = ["CSM","Coach","Account","Enterprise ID","Event","BOQ MRR","Current MRR","Change","Adjustment Amount"];
      const csv = [headers.join(","), ...rows.map(r =>
        [r.csm,r.coach,r.acct,r.eid||"",r.status,r.boqMrr,r.curMrr,r.delta,r.status==="adjustment"?r.adjAmount:""]
        .map(v=>{ const s=String(v??"").replace(/"/g,'""'); return s.includes(",")||s.includes('"')?'"'+s+'"':s; })
        .join(",")
      )].join("\n");
      const a = document.createElement("a");
      a.href = URL.createObjectURL(new Blob([csv],{type:"text/csv"}));
      a.download = "Q2-"+(q2TileFilter||"all")+"-"+new Date().toISOString().slice(0,10)+".csv";
      a.click();
    };

    return (
      <div>
        {/* Preliminary data disclaimer */}
        <div style={{display:"flex",alignItems:"center",gap:8,background:"#FFF7ED",border:"0.5px solid rgba(217,119,6,.25)",
          borderRadius:8,padding:"9px 14px",marginBottom:14,fontSize:12,color:"#92400e"}}>
          <span style={{fontSize:14}}>⚠️</span>
          <span>
            <strong>Preliminary numbers.</strong> These figures include all submitted adjustments, which may not be fully approved yet.
            This view shows what results would look like if every submitted adjustment is accepted as-is — treat it as a working estimate, not final.
          </span>
        </div>

        {/* 7 tiles: Q2 Retention (reset), Beginning Book & Ending Book (info only), Increases/Decreases/Cancels/Adjustments (filters) */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(7,minmax(0,1fr))",gap:10,marginBottom:16}}>
          <div onClick={()=>{ setQ2TileFilter(null); setQ2CSMFilter(null); }}
            style={{background:!q2TileFilter?"#29355D":"#ECEEF1",borderRadius:"0 0 10px 10px",padding:"12px 14px",
              borderTop:"3px solid #29355D",cursor:"pointer",transition:"all .15s"}}>
            <div style={{fontSize:10,textTransform:"uppercase",color:!q2TileFilter?"rgba(255,255,255,.7)":"#808080",fontWeight:500,marginBottom:4}}>Q2 Retention</div>
            <div style={{fontSize:22,fontWeight:600,color:!q2TileFilter?"#fff":retCol(overallRet),lineHeight:1,marginBottom:4}}>{fmtPct(overallRet)}</div>
            <div style={{fontSize:10,color:!q2TileFilter?"rgba(255,255,255,.6)":"#808080"}}>goal 91%</div>
          </div>
          <div style={{background:"#ECEEF1",borderRadius:"0 0 10px 10px",padding:"12px 14px",borderTop:"3px solid #5378FC"}}>
            <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:4}}>Beginning Book</div>
            <div style={{fontSize:22,fontWeight:600,color:"#5378FC",lineHeight:1,marginBottom:4}}>{fmt$(totalBoqAdj)}</div>
            <div style={{fontSize:10,color:"#808080"}}>{q2Data.length} CSMs</div>
          </div>
          <div style={{background:"#ECEEF1",borderRadius:"0 0 10px 10px",padding:"12px 14px",borderTop:"3px solid #29355D"}}>
            <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:4}}>Ending Book</div>
            <div style={{fontSize:22,fontWeight:600,color:"#29355D",lineHeight:1,marginBottom:4}}>{fmt$(totalCurrent)}</div>
            <div style={{fontSize:10,color:"#808080"}}>current billing</div>
          </div>
          {tileBtn2("Increases",fmt$(totalIncrease),q2Data.reduce((s,c)=>s+c.acctRows.filter(r=>r.status==="increase").length,0)+" accounts","#16a34a","increase")}
          {tileBtn2("Decreases",fmt$(totalDecrease),q2Data.reduce((s,c)=>s+c.decreaseCount,0)+" accounts","#dc2626","decrease")}
          {tileBtn2("Cancels",fmt$(totalCancelled),q2Data.reduce((s,c)=>s+c.cancelledCount,0)+" accounts","#d97706","cancelled")}
          {tileBtn2("Adjustments",fmt$(totalAdjustment),q2Data.reduce((s,c)=>s+c.adjustmentCount,0)+" accounts","#1e40af","adjustment")}
        </div>

        {/* Sortable CSM table with inline expand */}
        <div style={{...S.card}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500}}>
              {q2TileFilter ? sortedCSMs.length+" CSMs with "+q2TileFilter.replace(/_/g," ")+" — click name to expand" : "CSM Q2 Retention — click name to expand"}
            </div>
            <button onClick={exportQ2CSV}
              style={{padding:"4px 12px",borderRadius:20,border:"0.5px solid rgba(41,53,93,.2)",
                background:"#fff",color:"#29355D",fontSize:11,fontWeight:500,cursor:"pointer"}}>
              ⬇ Export CSV
            </button>
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr>
              {sortTh2("name","CSM",false)}
              {sortTh2("boqAdjusted","BOQ (adj)",true)}
              {sortTh2("currentMrr","Current MRR",true)}
              {sortTh2("netNewMrr","Net New",true)}
              {sortTh2("removedMrr","Removed",true)}
              {sortTh2("cancelledMrr","Cancelled",true)}
              {sortTh2("increaseMrr","Increase",true)}
              {sortTh2("decreaseMrr","Decrease",true)}
              {sortTh2("adjustmentMrr","Adjustment",true)}
              <th key="retPct" onClick={()=>setQ2Sort(s=>({col:"retPct",dir:s.col==="retPct"&&s.dir==="asc"?"desc":"asc"}))} style={{padding:"0 8px 8px 24px",textAlign:"left",fontSize:10,textTransform:"uppercase",color:q2Sort.col==="retPct"?"#29355D":"#808080",fontWeight:q2Sort.col==="retPct"?700:500,cursor:"pointer",borderBottom:"0.5px solid rgba(41,53,93,.08)",userSelect:"none",whiteSpace:"nowrap"}}>{"Retention %"}{q2Sort.col==="retPct"?(q2Sort.dir==="asc"?" ↑":" ↓"):""}</th>
            </tr></thead>
            <tbody>
              {sortedCSMs.map(c => {
                const isOpen = q2CSMFilter === c.name;
                const acctRows = c.acctRows.filter(r => !q2TileFilter || r.status === q2TileFilter);
                return (
                  <React.Fragment key={c.name}>
                    <tr onClick={()=>setQ2CSMFilter(isOpen?null:c.name)}
                      style={{cursor:"pointer",background:isOpen?"rgba(41,53,93,.04)":"transparent"}}>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",fontWeight:500,color:"#29355D"}}>
                        <span style={{marginRight:6,fontSize:9,display:"inline-block",transition:"transform .15s",
                          transform:isOpen?"rotate(90deg)":"none",color:"#808080"}}>▶</span>
                        {dispName(c.name)}
                      </td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:"#808080"}}>{fmt$(c.boqAdjusted)}</td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}>{fmt$(c.currentMrr)}</td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:"#16a34a"}}>{c.netNewMrr>0?"+"+fmt$(c.netNewMrr):"--"}</td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:"#808080"}}>--</td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:c.cancelledMrr>0?"#d97706":"#808080"}}>{c.cancelledMrr>0?"-"+fmt$(c.cancelledMrr):"--"}</td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:c.increaseMrr>0?"#16a34a":"#808080"}}>{c.increaseMrr>0?"+"+fmt$(c.increaseMrr):"--"}</td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:c.decreaseMrr>0?"#dc2626":"#808080"}}>{c.decreaseMrr>0?"-"+fmt$(c.decreaseMrr):"--"}</td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:c.adjustmentMrr>0?"#1e40af":"#808080"}}>{c.adjustmentMrr>0?"+"+fmt$(c.adjustmentMrr):"--"}</td>
                      <td style={{padding:"8px 8px 8px 24px",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)"}}>
                        <div style={{display:"flex",alignItems:"center",gap:8}}>
                          <div style={{width:80,height:5,background:"#ECEEF1",borderRadius:3,overflow:"hidden"}}>
                            <div style={{width:Math.min((c.retPct||0)*100,100).toFixed(1)+"%",height:"100%",background:retCol(c.retPct),borderRadius:3}}/>
                          </div>
                          <span style={{fontWeight:600,color:retCol(c.retPct)}}>{fmtPct(c.retPct)}</span>
                        </div>
                      </td>
                    </tr>
                    {isOpen && acctRows.length === 0 && (
                      <tr><td colSpan={10} style={{padding:"8px 0 8px 24px",borderBottom:"0.5px solid rgba(41,53,93,.08)",
                        color:"#808080",fontSize:11,fontStyle:"italic"}}>
                        No {q2TileFilter?q2TileFilter.replace(/_/g," "):"tracked changes"} for {dispName(c.name)} this quarter
                      </td></tr>
                    )}
                    {isOpen && acctRows.length > 0 && (() => {
                      // Derive sortable values per row and apply current sort
                      const rowsWithVals = acctRows.map(r => ({
                        ...r,
                        netNewVal:   r.status==="net_new"   ? r.curMrr : 0,
                        cancelledVal:r.status==="cancelled" ? r.boqMrr : 0,
                        increaseVal: r.status==="increase"  ? r.delta  : 0,
                        decreaseVal: r.status==="decrease"  ? Math.abs(r.delta) : 0,
                        adjVal:      r.status==="adjustment"? r.adjAmount : 0,
                      }));
                      const sortedAcctRows = [...rowsWithVals].sort((a,b) => {
                        const dir = q2AcctSort.dir==="asc" ? 1 : -1;
                        const col = q2AcctSort.col;
                        if (col==="acct") return dir*a.acct.localeCompare(b.acct);
                        const va = col==="boqMrr"?a.boqMrr:col==="curMrr"?a.curMrr:col==="netNewVal"?a.netNewVal:col==="cancelledVal"?a.cancelledVal:col==="increaseVal"?a.increaseVal:col==="decreaseVal"?a.decreaseVal:col==="adjVal"?a.adjVal:0;
                        const vb = col==="boqMrr"?b.boqMrr:col==="curMrr"?b.curMrr:col==="netNewVal"?b.netNewVal:col==="cancelledVal"?b.cancelledVal:col==="increaseVal"?b.increaseVal:col==="decreaseVal"?b.decreaseVal:col==="adjVal"?b.adjVal:0;
                        return dir*(va-vb);
                      });
                      const aTh = (col,label,right=true) => {
                        const active = q2AcctSort.col===col;
                        return (
                          <th key={col} onClick={()=>setQ2AcctSort(s=>({col,dir:s.col===col&&s.dir==="asc"?"desc":"asc"}))}
                            style={{padding:"4px 8px 4px 0",textAlign:right?"right":"left",fontSize:10,
                              textTransform:"uppercase",color:active?"#29355D":"#808080",fontWeight:active?700:500,
                              cursor:"pointer",userSelect:"none",whiteSpace:"nowrap"}}>
                            {label}{active?(q2AcctSort.dir==="asc"?" ↑":" ↓"):""}
                          </th>
                        );
                      };
                      return (
                      <tr>
                        <td colSpan={10} style={{padding:"0 0 12px 24px",borderBottom:"0.5px solid rgba(41,53,93,.08)",background:"rgba(41,53,93,.02)"}}>
                          <div style={{fontSize:10,color:"#808080",padding:"6px 0 4px",fontStyle:"italic"}}>
                            {sortedAcctRows.length} accounts{q2TileFilter?" · filtered: "+q2TileFilter.replace(/_/g," "):""}
                          </div>
                          <table style={{width:"100%",borderCollapse:"collapse",fontSize:11,tableLayout:"fixed"}}>
                            <colgroup>
                              <col style={{width:"28%"}}/><col style={{width:"8%"}}/><col style={{width:"8%"}}/>
                              <col style={{width:"7%"}}/><col style={{width:"7%"}}/><col style={{width:"7%"}}/>
                              <col style={{width:"7%"}}/><col style={{width:"7%"}}/><col style={{width:"9%"}}/><col style={{width:"10%"}}/>
                            </colgroup>
                            <thead><tr style={{borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>
                              {aTh("acct","Account",false)}
                              {aTh("boqMrr","BOQ MRR")}
                              {aTh("curMrr","Current MRR")}
                              {aTh("netNewVal","Net New")}
                              <th style={{padding:"4px 8px 4px 0",textAlign:"right",fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,whiteSpace:"nowrap"}}>Removed</th>
                              {aTh("cancelledVal","Cancelled")}
                              {aTh("increaseVal","Increase")}
                              {aTh("decreaseVal","Decrease")}
                              {aTh("adjVal","Adjustment")}
                              <th style={{padding:"4px 0 4px 24px",textAlign:"left",fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500}}>Status</th>
                            </tr></thead>
                            <tbody>
                              {sortedAcctRows.map((r,i)=>{
                                const isIncrease = r.status==="increase";
                                const isDecrease = r.status==="decrease";
                                const isNetNew   = r.status==="net_new";
                                const isCancelled= r.status==="cancelled";
                                const isAdj      = r.status==="adjustment";
                                return (
                                  <tr key={i} style={{borderTop:"0.5px solid rgba(41,53,93,.05)"}}>
                                    <td style={{padding:"5px 8px 5px 0",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.acct}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right",color:"#808080"}}>{r.boqMrr>0?fmt$(r.boqMrr):"--"}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right"}}>{r.curMrr>0?fmt$(r.curMrr):"--"}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right",color:"#16a34a"}}>{isNetNew?"+"+fmt$(r.curMrr):"--"}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right",color:"#808080"}}>--</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right",color:"#d97706"}}>{isCancelled?"-"+fmt$(r.boqMrr):"--"}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right",color:"#16a34a"}}>{isIncrease?"+"+fmt$(r.delta):"--"}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right",color:"#dc2626"}}>{isDecrease?"-"+fmt$(Math.abs(r.delta)):"--"}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right",color:"#1e40af"}}>{isAdj?"+"+fmt$(r.adjAmount):"--"}</td>
                                    <td style={{padding:"5px 0 5px 24px"}}>{statusBadge2(r.status)}</td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </td>
                      </tr>
                      );
                    })()}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    );
  };

  const renderQ3 = () => {
    if (!hasQ3) return (
      <div style={{...S.card,textAlign:"center",padding:"40px 20px",color:"#808080"}}>
        <div style={{fontSize:32,marginBottom:12}}>📊</div>
        <div style={{fontSize:14,fontWeight:500,color:"#29355D",marginBottom:8}}>Q3 Tracking not yet active</div>
        <div style={{fontSize:12,marginBottom:4}}>Run <strong>lockQ3BOQ()</strong> in Apps Script to lock the BOQ baseline.</div>
        <div style={{fontSize:12}}>Then run <strong>runQ3BOBUpdate()</strong> to start tracking changes.</div>
      </div>
    );

    const q3CSMs = Object.values(q3Current).filter(c => {
      if (DEACTIVATED_CSMS.has((norm(c.name)||c.name||"").toLowerCase().trim())) return false;
      const i = lk(norm(c.name)) || lk(c.name);
      if (managerCoaches && !(i && managerCoaches.includes(i.c))) return false;
      if (filterCoach && (i && i.c) !== filterCoach) return false;
      if (filterCSM && norm(c.name) !== filterCSM && c.name !== filterCSM) return false;
      return true;
    });

    const scopedNames = new Set(q3CSMs.map(c => norm(c.name)));
    const scopedLog   = [...q3Log].filter(r => scopedNames.has(norm(r.csm)) || scopedNames.has(r.csm));

    // Build full account-level maps keyed by EID
    const pf = v => parseFloat(String(v||"0").replace(/[$,]/g,""))||0;
    const getCol = (row, ...names) => {
      for (const n of names) {
        const k = Object.keys(row).find(k => k.toLowerCase().replace(/[^a-z]/g,"") === n.toLowerCase().replace(/[^a-z]/g,""));
        if (k) return row[k];
      }
      return "";
    };
    const normName = s => String(s||"").trim().toLowerCase();

    // BOQ accounts: {eid: {csm, acct, mrr}}
    const boqAcctMap = {};
    q3BobBoq.forEach(r => {
      const eid = String(getCol(r,"EnterprisId","EnterpriseId","enterprise id","Enterprise Id")||"").trim();
      if (!eid) return;
      boqAcctMap[eid] = {
        csm:  String(getCol(r,"Client Success Manager","csm_name")||"").trim(),
        acct: String(getCol(r,"Account Name","account name")||"").trim(),
        mrr:  pf(getCol(r,"SaaS Revenue","saas revenue")),
      };
    });

    // Current accounts: {eid: {csm, acct, mrr}} — combines bob_q3_current (SF) with
    // bob_q3_supplemental (secondary SF report), summed by Enterprise ID. It doesn't
    // matter which file an account's revenue came from — the totals should be combined.
    const curAcctMap = {};
    q3BobCur.forEach(r => {
      const eid = String(getCol(r,"EnterprisId","EnterpriseId","enterprise id","Enterprise Id")||"").trim();
      if (!eid) return;
      const csm  = String(getCol(r,"Client Success Manager","csm_name")||"").trim();
      const acct = String(getCol(r,"Account Name","account name")||"").trim();
      const mrr  = pf(getCol(r,"SaaS Revenue","saas revenue"));
      if (!curAcctMap[eid]) curAcctMap[eid] = {csm, acct, mrr:0};
      curAcctMap[eid].mrr += mrr;
      if (!curAcctMap[eid].acct && acct) curAcctMap[eid].acct = acct;
      if (!curAcctMap[eid].csm && csm) curAcctMap[eid].csm = csm;
    });
    q3Supp.forEach(r => {
      const eid = String(getCol(r,"EnterprisId","EnterpriseId","enterprise id","Enterprise Id")||"").trim();
      if (!eid) return;
      const acct = String(getCol(r,"Account Name","account name")||"").trim();
      const mrr  = pf(getCol(r,"SaaS Revenue","saas revenue"));
      if (!curAcctMap[eid]) curAcctMap[eid] = {csm:"", acct, mrr:0}; // no CSM column in supplemental
      curAcctMap[eid].mrr += mrr;
      if (!curAcctMap[eid].acct && acct) curAcctMap[eid].acct = acct;
    });

    // Build full book for a CSM: all BOQ accounts + net new, with current status
    const fullBook = (csmName) => {
      const n = normName(csmName);
      const rows = [];
      const seen = new Set();

      // All BOQ accounts for this CSM
      Object.entries(boqAcctMap).forEach(([eid, b]) => {
        if (normName(b.csm) !== n) return;
        seen.add(eid);
        const cur = curAcctMap[eid];
        const curMrr = cur ? cur.mrr : null;
        let status, delta;
        if (curMrr === null) {
          status = "removed"; delta = -b.mrr;
        } else if (b.mrr === 0 && curMrr > 0) {
          status = "net_new"; delta = curMrr;
        } else if (b.mrr > 0 && curMrr === 0) {
          status = "cancelled"; delta = -b.mrr;
        } else if (Math.abs(curMrr - b.mrr) > 0.01) {
          status = curMrr > b.mrr ? "increase" : "decrease"; delta = curMrr - b.mrr;
        } else {
          status = "unchanged"; delta = 0;
        }
        rows.push({ eid, acct: b.acct, boqMrr: b.mrr, curMrr: curMrr ?? 0, status, delta });
      });

      // Net new accounts (in current but not in BOQ)
      Object.entries(curAcctMap).forEach(([eid, c]) => {
        if (seen.has(eid)) return;
        if (normName(c.csm) !== n) return;
        if (c.mrr <= 0) return;
        rows.push({ eid, acct: c.acct, boqMrr: 0, curMrr: c.mrr, status: "net_new", delta: c.mrr });
      });

      return rows;
    };

    const totalBoqAdj    = q3CSMs.reduce((s,c)=>s+c.boqAdjusted, 0);
    const totalCurrent   = q3CSMs.reduce((s,c)=>s+c.currentMrr, 0);
    const totalNetNew    = q3CSMs.reduce((s,c)=>s+c.netNewMrr, 0);
    const totalRemoved   = q3CSMs.reduce((s,c)=>s+c.removedMrr, 0);
    const totalCancelled = q3CSMs.reduce((s,c)=>s+c.cancelledMrr, 0);
    const overallRet     = totalBoqAdj > 0 ? (totalCurrent - totalNetNew) / totalBoqAdj : null;
    const runDate        = q3CSMs[0]?.runDate || "";

    const increaseLog   = scopedLog.filter(r => r.event==="billing_change" && r.mrrDelta>0);
    const totalIncrease = increaseLog.reduce((s,r)=>s+r.mrrDelta, 0);

    // Per-CSM increase totals from log
    const csmIncreaseMrr = {};
    increaseLog.forEach(r => {
      const k = norm(r.csm)||r.csm;
      csmIncreaseMrr[k] = (csmIncreaseMrr[k]||0) + r.mrrDelta;
    });
    const getIncrease = (c) => csmIncreaseMrr[norm(c.name)] || csmIncreaseMrr[c.name] || 0;

    // Filter CSM table by active tile
    const csmsWithEvent = (type) => {
      const logs = type === "increase" ? increaseLog : scopedLog.filter(r => r.event === type);
      const names = new Set();
      logs.forEach(r => { names.add(r.csm); names.add(norm(r.csm)); });
      return names;
    };
    const activeCsmSet = tileFilter ? csmsWithEvent(tileFilter) : null;
    const visibleCSMs  = activeCsmSet
      ? q3CSMs.filter(c => activeCsmSet.has(c.name) || activeCsmSet.has(norm(c.name)))
      : q3CSMs;

    const enrichedCSMs = visibleCSMs.map(c => ({...c, increaseMrr: getIncrease(c)}));
    const sortedCSMs = [...enrichedCSMs].sort((a, b) => {
      const dir = q3Sort.dir === "asc" ? 1 : -1;
      const col = q3Sort.col;
      const va = a[col] ?? (col === "name" ? "" : 0);
      const vb = b[col] ?? (col === "name" ? "" : 0);
      if (col === "name") return dir * String(va).localeCompare(String(vb));
      return dir * ((Number(va)||0) - (Number(vb)||0));
    });

    const csmLog = (csmName) => scopedLog.filter(r => {
      if (norm(r.csm) !== norm(csmName) && r.csm !== csmName) return false;
      if (!tileFilter) return true;
      if (tileFilter === "increase") return r.event === "billing_change" && r.mrrDelta > 0;
      return r.event === tileFilter;
    }).reverse();

    const fmt$   = n => "$"+Number(n||0).toLocaleString("en-US",{minimumFractionDigits:0,maximumFractionDigits:0});
    const fmtPct = p => p!=null ? (p*100).toFixed(1)+"%" : "--";
    const retCol = p => p==null?"#808080":p>=0.91?"#16a34a":p>=0.85?"#d97706":"#dc2626";

    const sortTh = (col, label, right) => {
      const active = q3Sort.col === col;
      return (
        <th key={col} onClick={()=>setQ3Sort(s=>({col, dir: s.col===col&&s.dir==="asc"?"desc":"asc"}))}
          style={{padding:"0 8px 8px 0",textAlign:right?"right":"left",fontSize:10,textTransform:"uppercase",
            color:active?"#29355D":"#808080",fontWeight:active?700:500,cursor:"pointer",
            borderBottom:"0.5px solid rgba(41,53,93,.08)",userSelect:"none",whiteSpace:"nowrap"}}>
          {label}{active?(q3Sort.dir==="asc"?" ↑":" ↓"):""}
        </th>
      );
    };

    const tileBtn = (label, value, sub, color, filterKey) => {
      const active = tileFilter === filterKey;
      return (
        <div key={filterKey} onClick={()=>{ setTileFilter(active?null:filterKey); setQ3CSMFilter(null); }}
          style={{background:active?"#29355D":"#ECEEF1",borderRadius:"0 0 10px 10px",padding:"12px 14px",
            borderTop:"3px solid "+color,cursor:"pointer",transition:"all .15s",
            boxShadow:active?"0 2px 8px rgba(41,53,93,.15)":"none"}}>
          <div style={{fontSize:10,textTransform:"uppercase",color:active?"rgba(255,255,255,.7)":"#808080",fontWeight:500,marginBottom:4}}>{label}</div>
          <div style={{fontSize:22,fontWeight:600,color:active?"#fff":color,lineHeight:1,marginBottom:4}}>{value}</div>
          <div style={{fontSize:10,color:active?"rgba(255,255,255,.6)":"#808080"}}>{sub}</div>
        </div>
      );
    };

    const eventBadge = (evt, delta) => {
      let cfg;
      if (evt === "billing_change") cfg = delta > 0
        ? {bg:"#dcfce7",fg:"#166534",label:"Increase"}
        : {bg:"#fef9c3",fg:"#854d0e",label:"Decrease"};
      else cfg = {
        net_new:  {bg:"#dcfce7",fg:"#166534",label:"Net New"},
        removed:  {bg:"#fee2e2",fg:"#991b1b",label:"Removed"},
        cancelled:{bg:"#fef9c3",fg:"#854d0e",label:"Cancelled"},
      }[evt] || {bg:"#f3f4f6",fg:"#374151",label:evt};
      return <span style={{fontSize:10,fontWeight:500,padding:"2px 8px",borderRadius:20,background:cfg.bg,color:cfg.fg}}>{cfg.label}</span>;
    };

    const exportCSV = () => {
      const rows = sortedCSMs.flatMap(c => csmLog(c.name).map(r => ({
        csm: c.name, account: r.acct, enterprise_id: r.eid||"",
        event: r.event==="billing_change"?(r.mrrDelta>0?"Increase":"Decrease"):r.event,
        boq_mrr: r.mrrBefore||0, current_mrr: r.mrrAfter||0, change: r.mrrDelta||0,
        date: r.date, note: r.note||"",
      })));
      if (!rows.length) return;
      const headers = ["CSM","Account","Enterprise ID","Event","BOQ MRR","Current MRR","Change","Date","Note"];
      const csv = [headers.join(","), ...rows.map(r =>
        [r.csm,r.account,r.enterprise_id,r.event,r.boq_mrr,r.current_mrr,r.change,r.date,r.note]
        .map(v=>{ const s=String(v??"").replace(/"/g,'""'); return s.includes(",")||s.includes('"')?'"'+s+'"':s; })
        .join(",")
      )].join("\n");
      const a = document.createElement("a");
      a.href = URL.createObjectURL(new Blob([csv],{type:"text/csv"}));
      a.download = "Q3-"+(tileFilter||"all")+"-"+new Date().toISOString().slice(0,10)+".csv";
      a.click();
    };

    return (
      <div>
        {/* 5 clickable tiles */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,minmax(0,1fr))",gap:10,marginBottom:16}}>
          <div onClick={()=>{ setTileFilter(null); setQ3CSMFilter(null); }}
            style={{background:!tileFilter?"#29355D":"#ECEEF1",borderRadius:"0 0 10px 10px",padding:"12px 14px",
              borderTop:"3px solid #29355D",cursor:"pointer",transition:"all .15s"}}>
            <div style={{fontSize:10,textTransform:"uppercase",color:!tileFilter?"rgba(255,255,255,.7)":"#808080",fontWeight:500,marginBottom:4}}>Q3 Retention</div>
            <div style={{fontSize:22,fontWeight:600,color:!tileFilter?"#fff":retCol(overallRet),lineHeight:1,marginBottom:4}}>{fmtPct(overallRet)}</div>
            <div style={{fontSize:10,color:!tileFilter?"rgba(255,255,255,.6)":"#808080"}}>goal 91% · adj BOQ {fmt$(totalBoqAdj)}</div>
          </div>
          {tileBtn("Increases",fmt$(totalIncrease),increaseLog.length+" accounts","#16a34a","increase")}
          {tileBtn("Cancelled ($0)",fmt$(totalCancelled),q3CSMs.reduce((s,c)=>s+c.cancelledCount,0)+" accounts","#d97706","cancelled")}
          {tileBtn("Removed from BOQ",fmt$(totalRemoved),q3CSMs.reduce((s,c)=>s+c.removedCount,0)+" accounts","#dc2626","removed")}
          {tileBtn("Net New",fmt$(totalNetNew),q3CSMs.reduce((s,c)=>s+c.netNewCount,0)+" accounts","#FF5000","net_new")}
        </div>

        {/* Sortable CSM table with inline expand */}
        <div style={{...S.card}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500}}>
              {tileFilter ? sortedCSMs.length+" CSMs with "+tileFilter.replace(/_/g," ")+" — click name to expand" : "CSM Q3 Retention — click name to expand"}
            </div>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              {runDate&&<div style={{fontSize:11,color:"#808080"}}>Last updated: {runDate}</div>}
              <button onClick={exportCSV}
                style={{padding:"4px 12px",borderRadius:20,border:"0.5px solid rgba(41,53,93,.2)",
                  background:"#fff",color:"#29355D",fontSize:11,fontWeight:500,cursor:"pointer"}}>
                ⬇ Export CSV
              </button>
            </div>
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr>
              {sortTh("name","CSM",false)}
              {sortTh("boqAdjusted","BOQ (adj)",true)}
              {sortTh("currentMrr","Current MRR",true)}
              {sortTh("netNewMrr","Net New",true)}
              {sortTh("removedMrr","Removed",true)}
              {sortTh("cancelledMrr","Cancelled",true)}
              {sortTh("increaseMrr","Increase",true)}
              <th key="retPct" onClick={()=>setQ3Sort(s=>({col:"retPct",dir:s.col==="retPct"&&s.dir==="asc"?"desc":"asc"}))} style={{padding:"0 8px 8px 24px",textAlign:"left",fontSize:10,textTransform:"uppercase",color:q3Sort.col==="retPct"?"#29355D":"#808080",fontWeight:q3Sort.col==="retPct"?700:500,cursor:"pointer",borderBottom:"0.5px solid rgba(41,53,93,.08)",userSelect:"none",whiteSpace:"nowrap"}}>{"Retention %"}{q3Sort.col==="retPct"?(q3Sort.dir==="asc"?" ↑":" ↓"):""}</th>
            </tr></thead>
            <tbody>
              {sortedCSMs.map(c => {
                const isOpen = q3CSMFilter === c.name || q3CSMFilter === norm(c.name);
                const drillLog = csmLog(c.name);
                return (
                  <React.Fragment key={c.name}>
                    <tr onClick={()=>setQ3CSMFilter(isOpen?null:c.name)}
                      style={{cursor:"pointer",background:isOpen?"rgba(41,53,93,.04)":"transparent"}}>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",fontWeight:500,color:"#29355D"}}>
                        <span style={{marginRight:6,fontSize:9,display:"inline-block",transition:"transform .15s",
                          transform:isOpen?"rotate(90deg)":"none",color:"#808080"}}>▶</span>
                        {dispName(c.name)}
                      </td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:"#808080"}}>{fmt$(c.boqAdjusted)}</td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}>{fmt$(c.currentMrr)}</td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:"#16a34a"}}>{c.netNewMrr>0?"+"+fmt$(c.netNewMrr):"--"}</td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:c.removedMrr>0?"#dc2626":"#808080"}}>{c.removedMrr>0?"-"+fmt$(c.removedMrr):"--"}</td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:c.cancelledMrr>0?"#d97706":"#808080"}}>{c.cancelledMrr>0?"-"+fmt$(c.cancelledMrr):"--"}</td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:c.increaseMrr>0?"#16a34a":"#808080"}}>{c.increaseMrr>0?"+"+fmt$(c.increaseMrr):"--"}</td>
                      <td style={{padding:"8px 8px 8px 24px",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)"}}>
                        <div style={{display:"flex",alignItems:"center",gap:8}}>
                          <div style={{width:80,height:5,background:"#ECEEF1",borderRadius:3,overflow:"hidden"}}>
                            <div style={{width:Math.min((c.retPct||0)*100,100).toFixed(1)+"%",height:"100%",background:retCol(c.retPct),borderRadius:3}}/>
                          </div>
                          <span style={{fontWeight:600,color:retCol(c.retPct)}}>{fmtPct(c.retPct)}</span>
                        </div>
                      </td>
                    </tr>
                    {isOpen && (() => {
                      const useFullBook = q3BobBoq.length > 0;
                      // Full book or log-based, filtered by tile
                      let acctRows = [];
                      if (useFullBook) {
                        acctRows = fullBook(c.name).filter(r => {
                          if (!tileFilter) return true;
                          if (tileFilter === "increase") return r.status === "increase";
                          if (tileFilter === "net_new")  return r.status === "net_new";
                          if (tileFilter === "cancelled") return r.status === "cancelled";
                          if (tileFilter === "removed")  return r.status === "removed";
                          return true;
                        });
                      } else {
                        acctRows = drillLog.map(r => ({
                          acct: r.acct, boqMrr: r.mrrBefore||0, curMrr: r.mrrAfter||0,
                          status: r.event==="billing_change"?(r.mrrDelta>0?"increase":"decrease"):r.event,
                          delta: r.mrrDelta||0,
                        }));
                      }
                      // Sort state is at BobView top level (acctSort/setAcctSort)
                      const acctSortVal = (r, col) => {
                        if (col === "acct")         return r.acct;
                        if (col === "boqMrr")       return r.boqMrr;
                        if (col === "curMrr")       return r.curMrr;
                        if (col === "netNew")       return r.status==="net_new"  ? r.curMrr : 0;
                        if (col === "removedMrr")   return r.status==="removed"  ? r.boqMrr : 0;
                        if (col === "cancelledMrr") return r.status==="cancelled"? r.boqMrr : 0;
                        if (col === "increase")     return r.status==="increase" ? r.delta  : 0;
                        return 0;
                      };
                      const sortedAccts = [...acctRows].sort((a,b) => {
                        const dir = acctSort.dir==="asc"?1:-1;
                        if (acctSort.col==="acct") return dir*a.acct.localeCompare(b.acct);
                        return dir*((acctSortVal(a,acctSort.col)||0)-(acctSortVal(b,acctSort.col)||0));
                      });
                      const sTh = (col,label,right=false) => (
                        <th key={col} onClick={()=>setAcctSort(s=>({col,dir:s.col===col&&s.dir==="asc"?"desc":"asc"}))}
                          style={{padding:"4px 8px 4px 0",textAlign:right?"right":"left",fontSize:10,
                            textTransform:"uppercase",color:acctSort.col===col?"#29355D":"#808080",
                            fontWeight:acctSort.col===col?700:500,cursor:"pointer",userSelect:"none",whiteSpace:"nowrap"}}>
                          {label}{acctSort.col===col?(acctSort.dir==="asc"?" ↑":" ↓"):""}
                        </th>
                      );
                      const statusBadge = (s) => {
                        const cfg = {
                          increase: {bg:"#dcfce7",fg:"#166534",label:"Increase"},
                          decrease: {bg:"#fef9c3",fg:"#854d0e",label:"Decrease"},
                          net_new:  {bg:"#dcfce7",fg:"#166534",label:"Net New"},
                          cancelled:{bg:"#fef9c3",fg:"#854d0e",label:"Cancelled"},
                          removed:  {bg:"#fee2e2",fg:"#991b1b",label:"Removed"},
                          unchanged:{bg:"#f3f4f6",fg:"#6b7280",label:"No Change"},
                        }[s]||{bg:"#f3f4f6",fg:"#374151",label:s};
                        return <span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:cfg.bg,color:cfg.fg}}>{cfg.label}</span>;
                      };
                      if (sortedAccts.length === 0) return (
                        <tr><td colSpan={8} style={{padding:"8px 0 8px 24px",borderBottom:"0.5px solid rgba(41,53,93,.08)",
                          color:"#808080",fontSize:11,fontStyle:"italic"}}>
                          No accounts found for {dispName(c.name)}{tileFilter?" with "+tileFilter.replace(/_/g," "):""}
                        </td></tr>
                      );
                      return (
                        <tr>
                          <td colSpan={8} style={{padding:"0 0 12px 24px",borderBottom:"0.5px solid rgba(41,53,93,.08)",background:"rgba(41,53,93,.02)"}}>
                            <div style={{fontSize:10,color:"#808080",padding:"6px 0 4px",fontStyle:"italic"}}>
                              {sortedAccts.length} accounts{tileFilter?" · filtered: "+tileFilter.replace(/_/g," "):""}
                            </div>
                            {/* Header row matching parent columns exactly */}
                            <table style={{width:"100%",borderCollapse:"collapse",fontSize:11,tableLayout:"fixed"}}>
                              <colgroup>
                                <col style={{width:"38%"}}/>
                                <col style={{width:"9%"}}/>
                                <col style={{width:"10%"}}/>
                                <col style={{width:"8%"}}/>
                                <col style={{width:"8%"}}/>
                                <col style={{width:"8%"}}/>
                                <col style={{width:"8%"}}/>
                                <col style={{width:"11%"}}/>
                              </colgroup>
                              <thead><tr style={{borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>
                                {sTh("acct","Account",false)}
                                {sTh("boqMrr","BOQ MRR",true)}
                                {sTh("curMrr","Current MRR",true)}
                                {sTh("netNew","Net New",true)}
                                {sTh("removedMrr","Removed",true)}
                                {sTh("cancelledMrr","Cancelled",true)}
                                {sTh("increase","Increase",true)}
                                <th style={{padding:"4px 0 4px 24px",textAlign:"left",fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500}}>Status</th>
                              </tr></thead>
                              <tbody>
                                {sortedAccts.map((r,i)=>{
                                  const isIncrease = r.status==="increase";
                                  const isDecrease = r.status==="decrease";
                                  const isNetNew   = r.status==="net_new";
                                  const isCancelled= r.status==="cancelled";
                                  const isRemoved  = r.status==="removed";
                                  return (
                                  <tr key={i} style={{borderTop:"0.5px solid rgba(41,53,93,.05)"}}>
                                    <td style={{padding:"5px 8px 5px 0",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.acct}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right",color:"#808080"}}>{r.boqMrr>0?fmt$(r.boqMrr):"--"}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right"}}>{r.curMrr>0?fmt$(r.curMrr):"--"}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right",color:"#16a34a"}}>{isNetNew?"+"+fmt$(r.curMrr):"--"}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right",color:"#dc2626"}}>{isRemoved?"-"+fmt$(r.boqMrr):"--"}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right",color:"#d97706"}}>{isCancelled?"-"+fmt$(r.boqMrr):"--"}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right",color:"#16a34a"}}>{isIncrease?"+"+fmt$(r.delta):"--"}</td>
                                    <td style={{padding:"5px 0 5px 24px"}}>{statusBadge(r.status)}</td>
                                  </tr>
                                );})}
                              </tbody>
                            </table>
                          </td>
                        </tr>
                      );
                    })()}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    );
  };

  // ── DOMO BOOK OF BUSINESS — new test tab ───────────────────────────────────
  // Beginning Book: fresh Domo BOQ export (multi-row per EID, one row per L2
  // product line) summed by Enterprise ID. Current Revenue: bob_q3_current +
  // bob_q3_supplemental, joined by Enterprise ID and summed together. CSM
  // attribution always comes from the Domo BOQ file, never from the SF files.
  const renderDomoBoB = () => {
    const pf = v => { const x = parseFloat(String(v||"0").replace(/[$,]/g,"")); return isNaN(x)?0:x; };
    const getCol = (row, ...names) => {
      for (const n of names) {
        const k = Object.keys(row).find(k => k.toLowerCase().replace(/[^a-z]/g,"") === n.toLowerCase().replace(/[^a-z]/g,""));
        if (k && row[k] !== undefined && row[k] !== "") return row[k];
      }
      return "";
    };
    const lfSwap = raw => {
      const s = String(raw||"").trim();
      if (!s) return "";
      if (s.includes(",")) {
        const [last, first] = s.split(",", 2);
        return (first.trim()+" "+last.trim()).replace(/  +/g," ").trim();
      }
      return s;
    };

    if (domoBoq.length === 0) return (
      <div style={{...S.card,textAlign:"center",padding:"40px 20px",color:"#808080"}}>
        <div style={{fontSize:32,marginBottom:12}}>📊</div>
        <div style={{fontSize:14,fontWeight:500,color:"#29355D",marginBottom:8}}>Domo Book of Business — waiting on data</div>
        <div style={{fontSize:12}}>Paste the fresh Domo BOQ export into the connected Google Sheet tab to populate this view.</div>
      </div>
    );

    // ── Parse Domo BOQ: group by EID, sum "Beginning of Quarter" across all L2 lines ──
    const boqMap = {};
    domoBoq.forEach(r => {
      const csmRaw = String(getCol(r,"CSM Name")||"").trim();
      if (!csmRaw || /TOTAL|GRAND/i.test(csmRaw)) return;
      const eid = String(getCol(r,"Enterprise ID","Enterprise Id")||"").trim().toUpperCase();
      if (!eid) return;
      const boqAmt = pf(getCol(r,"Beginning of Quarter"));
      const acct = String(getCol(r,"Account Name")||"").trim();
      const csmName = norm(lfSwap(csmRaw)) || lfSwap(csmRaw);
      if (!boqMap[eid]) boqMap[eid] = {eid, csm:csmName, acct, boq:0};
      boqMap[eid].boq += boqAmt;
      if (!boqMap[eid].acct && acct) boqMap[eid].acct = acct;
    });

    // ── Parse revenue sources, tracking presence separately from value ────────
    const parseRevenueByEid = (rows) => {
      const map = {}; const seen = new Set();
      rows.forEach(r => {
        const eid = String(getCol(r,"Enterprise Id","Enterprise ID")||"").trim().toUpperCase();
        if (!eid) return;
        seen.add(eid);
        map[eid] = (map[eid]||0) + pf(getCol(r,"SaaS Revenue"));
      });
      return {map, seen};
    };
    const sfData   = parseRevenueByEid(q3BobCur);
    const suppData = parseRevenueByEid(q3Supp);

    // ── Join: BOQ is the authoritative account list ──────────────────────────
    const allRows = Object.values(boqMap).map(b => {
      const sfRev   = sfData.map[b.eid]   || 0;
      const suppRev = suppData.map[b.eid] || 0;
      const cur = sfRev + suppRev;
      const delta = cur - b.boq;
      const ret = b.boq > 0 ? cur / b.boq : null;
      const seenAnywhere = sfData.seen.has(b.eid) || suppData.seen.has(b.eid);
      let src = (sfRev>0 && suppRev>0) ? "both" : sfRev>0 ? "sf" : suppRev>0 ? "supp" : "none";
      let status;
      if (!seenAnywhere) status = "no_data";
      else if (b.boq > 0 && cur === 0) status = "cancelled";
      else if (Math.abs(delta) < 0.5) status = "unchanged";
      else if (delta > 0) status = "increase";
      else status = "decrease";
      return {eid:b.eid, csm:b.csm, acct:b.acct, boq:b.boq, sfRev, suppRev, cur, delta, ret, status, src};
    });

    // ── Orphans: in SF/Supp but not in BOQ at all ─────────────────────────────
    const boqEids = new Set(Object.keys(boqMap));
    const orphanEids = new Set([...Object.keys(sfData.map), ...Object.keys(suppData.map)].filter(e=>!boqEids.has(e)));
    const orphanCount = orphanEids.size;

    // ── Scope by coach/manager/CSM filters from the top of the page ──────────
    const scopedRows = allRows.filter(r => {
      const i = lk(norm(r.csm)) || lk(r.csm);
      if (managerCoaches && !(i && managerCoaches.includes(i.c))) return false;
      if (filterCoach && (i && i.c) !== filterCoach) return false;
      if (filterCSM && norm(r.csm) !== filterCSM && r.csm !== filterCSM) return false;
      return true;
    });

    // ── Group by CSM ───────────────────────────────────────────────────────────
    const csmGroups = {};
    scopedRows.forEach(r => {
      const key = r.csm || "(unknown)";
      if (!csmGroups[key]) csmGroups[key] = {
        name:key, boq:0, cur:0,
        increaseMrr:0, decreaseMrr:0, cancelledMrr:0,
        increaseCount:0, decreaseCount:0, cancelledCount:0, noDataCount:0,
        rows:[],
      };
      const g = csmGroups[key];
      g.boq += r.boq; g.cur += r.cur; g.rows.push(r);
      if (r.status==="increase")  { g.increaseMrr  += r.delta;          g.increaseCount++;  }
      if (r.status==="decrease")  { g.decreaseMrr  += Math.abs(r.delta); g.decreaseCount++;  }
      if (r.status==="cancelled") { g.cancelledMrr += r.boq;             g.cancelledCount++; }
      if (r.status==="no_data")   { g.noDataCount++; }
    });
    const domoData = Object.values(csmGroups).map(g => ({...g, delta:g.cur-g.boq, retPct: g.boq>0 ? g.cur/g.boq : null}));

    const totalBoq       = domoData.reduce((s,c)=>s+c.boq, 0);
    const totalCur        = domoData.reduce((s,c)=>s+c.cur, 0);
    const totalIncrease   = domoData.reduce((s,c)=>s+c.increaseMrr, 0);
    const totalDecrease   = domoData.reduce((s,c)=>s+c.decreaseMrr, 0);
    const totalCancelled  = domoData.reduce((s,c)=>s+c.cancelledMrr, 0);
    const totalNoData     = domoData.reduce((s,c)=>s+c.noDataCount, 0);
    const overallRet      = totalBoq > 0 ? totalCur/totalBoq : null;

    // ── Filter CSM table by active tile ───────────────────────────────────────
    const csmsWithEvent = (type) => {
      if (type === "increase")  return new Set(domoData.filter(c=>c.increaseMrr>0).map(c=>c.name));
      if (type === "decrease")  return new Set(domoData.filter(c=>c.decreaseMrr>0).map(c=>c.name));
      if (type === "cancelled") return new Set(domoData.filter(c=>c.cancelledMrr>0).map(c=>c.name));
      if (type === "no_data")   return new Set(domoData.filter(c=>c.noDataCount>0).map(c=>c.name));
      return null;
    };
    const activeCsmSet = domoTileFilter ? csmsWithEvent(domoTileFilter) : null;
    const visibleCSMs  = activeCsmSet ? domoData.filter(c => activeCsmSet.has(c.name)) : domoData;

    const sortedCSMs = [...visibleCSMs].sort((a, b) => {
      const dir = domoSort.dir === "asc" ? 1 : -1;
      const col = domoSort.col;
      const va = a[col] ?? (col === "name" ? "" : 0);
      const vb = b[col] ?? (col === "name" ? "" : 0);
      if (col === "name") return dir * String(va).localeCompare(String(vb));
      return dir * ((Number(va)||0) - (Number(vb)||0));
    });

    const fmt$   = n => "$"+Number(n||0).toLocaleString("en-US",{minimumFractionDigits:0,maximumFractionDigits:0});
    const fmtPct = p => p!=null ? (p*100).toFixed(1)+"%" : "--";
    const retCol = p => p==null?"#808080":p>=0.91?"#16a34a":p>=0.85?"#d97706":"#dc2626";

    const sortTh = (col, label, right) => {
      const active = domoSort.col === col;
      return (
        <th key={col} onClick={()=>setDomoSort(s=>({col, dir: s.col===col&&s.dir==="asc"?"desc":"asc"}))}
          style={{padding:"0 8px 8px 0",textAlign:right?"right":"left",fontSize:10,textTransform:"uppercase",
            color:active?"#29355D":"#808080",fontWeight:active?700:500,cursor:"pointer",
            borderBottom:"0.5px solid rgba(41,53,93,.08)",userSelect:"none",whiteSpace:"nowrap"}}>
          {label}{active?(domoSort.dir==="asc"?" ↑":" ↓"):""}
        </th>
      );
    };

    const tileBtn = (label, value, sub, color, filterKey) => {
      const active = domoTileFilter === filterKey;
      return (
        <div key={filterKey} onClick={()=>{ setDomoTileFilter(active?null:filterKey); setDomoCSMFilter(null); }}
          style={{background:active?"#29355D":"#ECEEF1",borderRadius:"0 0 10px 10px",padding:"12px 14px",
            borderTop:"3px solid "+color,cursor:"pointer",transition:"all .15s",
            boxShadow:active?"0 2px 8px rgba(41,53,93,.15)":"none"}}>
          <div style={{fontSize:10,textTransform:"uppercase",color:active?"rgba(255,255,255,.7)":"#808080",fontWeight:500,marginBottom:4}}>{label}</div>
          <div style={{fontSize:22,fontWeight:600,color:active?"#fff":color,lineHeight:1,marginBottom:4}}>{value}</div>
          <div style={{fontSize:10,color:active?"rgba(255,255,255,.6)":"#808080"}}>{sub}</div>
        </div>
      );
    };

    const statusBadge = (s) => {
      const cfg = {
        increase:  {bg:"#dcfce7",fg:"#166534",label:"Increase"},
        decrease:  {bg:"#fee2e2",fg:"#991b1b",label:"Decrease"},
        cancelled: {bg:"#fef9c3",fg:"#854d0e",label:"Cancelled"},
        no_data:   {bg:"#f3f4f6",fg:"#6b7280",label:"No Data"},
        unchanged: {bg:"#f3f4f6",fg:"#6b7280",label:"No Change"},
      }[s]||{bg:"#f3f4f6",fg:"#374151",label:s};
      return <span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:cfg.bg,color:cfg.fg}}>{cfg.label}</span>;
    };

    const exportDomoCSV = () => {
      const rows = sortedCSMs.flatMap(c => c.rows
        .filter(r => !domoTileFilter || r.status === domoTileFilter)
        .map(r => ({ csm: c.name, ...r }))
      );
      if (!rows.length) return;
      const headers = ["CSM","Account","Enterprise ID","BOQ","SF MRR","Supp MRR","Combined","Change","Retention %","Source","Status"];
      const csv = [headers.join(","), ...rows.map(r =>
        [dispName(r.csm),r.acct,r.eid,r.boq,r.sfRev,r.suppRev,r.cur,r.delta,r.ret!=null?(r.ret*100).toFixed(1):"",r.src,r.status]
        .map(v=>{ const s=String(v??"").replace(/"/g,'""'); return s.includes(",")||s.includes('"')?'"'+s+'"':s; })
        .join(",")
      )].join("\n");
      const a = document.createElement("a");
      a.href = URL.createObjectURL(new Blob([csv],{type:"text/csv"}));
      a.download = "Domo-BoB-"+(domoTileFilter||"all")+"-"+new Date().toISOString().slice(0,10)+".csv";
      a.click();
    };

    return (
      <div>
        {/* Test tab disclaimer */}
        <div style={{display:"flex",alignItems:"center",gap:8,background:"#EFF6FF",border:"0.5px solid rgba(30,64,175,.2)",
          borderRadius:8,padding:"9px 14px",marginBottom:14,fontSize:12,color:"#1e40af"}}>
          <span style={{fontSize:14}}>🧪</span>
          <span><strong>Test tab.</strong> Beginning Book comes from a fresh Domo export (separate from the locked Q3 BOQ).
            Current Revenue combines bob_q3_current + bob_q3_supplemental. This does not affect Q2 or Q3 Tracking.</span>
        </div>

        {/* 7 tiles: Retention (reset), Beginning/Combined Book (info), Increases/Decreases/Cancelled/No Data (filters) */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(7,minmax(0,1fr))",gap:10,marginBottom:16}}>
          <div onClick={()=>{ setDomoTileFilter(null); setDomoCSMFilter(null); }}
            style={{background:!domoTileFilter?"#29355D":"#ECEEF1",borderRadius:"0 0 10px 10px",padding:"12px 14px",
              borderTop:"3px solid #29355D",cursor:"pointer",transition:"all .15s"}}>
            <div style={{fontSize:10,textTransform:"uppercase",color:!domoTileFilter?"rgba(255,255,255,.7)":"#808080",fontWeight:500,marginBottom:4}}>Retention</div>
            <div style={{fontSize:22,fontWeight:600,color:!domoTileFilter?"#fff":retCol(overallRet),lineHeight:1,marginBottom:4}}>{fmtPct(overallRet)}</div>
            <div style={{fontSize:10,color:!domoTileFilter?"rgba(255,255,255,.6)":"#808080"}}>goal 91%</div>
          </div>
          <div style={{background:"#ECEEF1",borderRadius:"0 0 10px 10px",padding:"12px 14px",borderTop:"3px solid #5378FC"}}>
            <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:4}}>Beginning Book</div>
            <div style={{fontSize:22,fontWeight:600,color:"#5378FC",lineHeight:1,marginBottom:4}}>{fmt$(totalBoq)}</div>
            <div style={{fontSize:10,color:"#808080"}}>{domoData.length} CSMs</div>
          </div>
          <div style={{background:"#ECEEF1",borderRadius:"0 0 10px 10px",padding:"12px 14px",borderTop:"3px solid #29355D"}}>
            <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:4}}>Current MRR</div>
            <div style={{fontSize:22,fontWeight:600,color:"#29355D",lineHeight:1,marginBottom:4}}>{fmt$(totalCur)}</div>
            <div style={{fontSize:10,color:"#808080"}}>SF + supplemental</div>
          </div>
          {tileBtn("Increases",fmt$(totalIncrease),domoData.reduce((s,c)=>s+c.increaseCount,0)+" accounts","#16a34a","increase")}
          {tileBtn("Decreases",fmt$(totalDecrease),domoData.reduce((s,c)=>s+c.decreaseCount,0)+" accounts","#dc2626","decrease")}
          {tileBtn("Cancelled",fmt$(totalCancelled),domoData.reduce((s,c)=>s+c.cancelledCount,0)+" accounts","#d97706","cancelled")}
          {tileBtn("No Data",String(totalNoData),"accounts","#6b7280","no_data")}
        </div>

        {/* Sortable CSM table with inline expand */}
        <div style={{...S.card}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500}}>
              {domoTileFilter ? sortedCSMs.length+" CSMs with "+domoTileFilter.replace(/_/g," ")+" — click name to expand" : "CSM Retention — click name to expand"}
            </div>
            <button onClick={exportDomoCSV}
              style={{padding:"4px 12px",borderRadius:20,border:"0.5px solid rgba(41,53,93,.2)",
                background:"#fff",color:"#29355D",fontSize:11,fontWeight:500,cursor:"pointer"}}>
              ⬇ Export CSV
            </button>
          </div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr>
              {sortTh("name","CSM",false)}
              {sortTh("boq","BOQ",true)}
              {sortTh("cur","Current MRR",true)}
              {sortTh("increaseMrr","Increase",true)}
              {sortTh("decreaseMrr","Decrease",true)}
              {sortTh("cancelledMrr","Cancelled",true)}
              {sortTh("noDataCount","No Data",true)}
              <th style={{padding:"0 8px 8px 24px",textAlign:"left",fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500}}>Retention %</th>
            </tr></thead>
            <tbody>
              {sortedCSMs.map(c => {
                const isOpen = domoCSMFilter === c.name;
                const acctRows = c.rows.filter(r => !domoTileFilter || r.status === domoTileFilter);
                return (
                  <React.Fragment key={c.name}>
                    <tr onClick={()=>setDomoCSMFilter(isOpen?null:c.name)}
                      style={{cursor:"pointer",background:isOpen?"rgba(41,53,93,.04)":"transparent"}}>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",fontWeight:500,color:"#29355D"}}>
                        <span style={{marginRight:6,fontSize:9,display:"inline-block",transition:"transform .15s",
                          transform:isOpen?"rotate(90deg)":"none",color:"#808080"}}>▶</span>
                        {dispName(c.name)}
                      </td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:"#808080"}}>{fmt$(c.boq)}</td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}>{fmt$(c.cur)}</td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:c.increaseMrr>0?"#16a34a":"#808080"}}>{c.increaseMrr>0?"+"+fmt$(c.increaseMrr):"--"}</td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:c.decreaseMrr>0?"#dc2626":"#808080"}}>{c.decreaseMrr>0?"-"+fmt$(c.decreaseMrr):"--"}</td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:c.cancelledMrr>0?"#d97706":"#808080"}}>{c.cancelledMrr>0?"-"+fmt$(c.cancelledMrr):"--"}</td>
                      <td style={{padding:"8px 8px 8px 0",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:c.noDataCount>0?"#6b7280":"#808080"}}>{c.noDataCount>0?c.noDataCount:"--"}</td>
                      <td style={{padding:"8px 8px 8px 24px",borderBottom:isOpen?"none":"0.5px solid rgba(41,53,93,.05)"}}>
                        <div style={{display:"flex",alignItems:"center",gap:8}}>
                          <div style={{width:80,height:5,background:"#ECEEF1",borderRadius:3,overflow:"hidden"}}>
                            <div style={{width:Math.min((c.retPct||0)*100,100).toFixed(1)+"%",height:"100%",background:retCol(c.retPct),borderRadius:3}}/>
                          </div>
                          <span style={{fontWeight:600,color:retCol(c.retPct)}}>{fmtPct(c.retPct)}</span>
                        </div>
                      </td>
                    </tr>
                    {isOpen && acctRows.length === 0 && (
                      <tr><td colSpan={8} style={{padding:"8px 0 8px 24px",borderBottom:"0.5px solid rgba(41,53,93,.08)",
                        color:"#808080",fontSize:11,fontStyle:"italic"}}>
                        No {domoTileFilter?domoTileFilter.replace(/_/g," "):"tracked"} accounts for {dispName(c.name)}
                      </td></tr>
                    )}
                    {isOpen && acctRows.length > 0 && (() => {
                      const sortedAcctRows = [...acctRows].sort((a,b) => {
                        const dir = domoAcctSort.dir==="asc"?1:-1;
                        const col = domoAcctSort.col;
                        if (col==="acct") return dir*a.acct.localeCompare(b.acct);
                        const val = (r) => {
                          if (col==="boq") return r.boq;
                          if (col==="cur") return r.cur;
                          if (col==="ret") return r.ret??-1;
                          if (col==="increase")  return r.status==="increase"  ? r.delta : 0;
                          if (col==="decrease")  return r.status==="decrease"  ? Math.abs(r.delta) : 0;
                          if (col==="cancelled") return r.status==="cancelled" ? r.boq : 0;
                          if (col==="noData")    return r.status==="no_data"  ? 1 : 0;
                          return 0;
                        };
                        return dir*(val(a)-val(b));
                      });
                      const aTh = (col,label,right=true) => {
                        const active = domoAcctSort.col===col;
                        return (
                          <th key={col} onClick={()=>setDomoAcctSort(s=>({col,dir:s.col===col&&s.dir==="asc"?"desc":"asc"}))}
                            style={{padding:"4px 8px 4px 0",textAlign:right?"right":"left",fontSize:10,
                              textTransform:"uppercase",color:active?"#29355D":"#808080",fontWeight:active?700:500,
                              cursor:"pointer",userSelect:"none",whiteSpace:"nowrap"}}>
                            {label}{active?(domoAcctSort.dir==="asc"?" ↑":" ↓"):""}
                          </th>
                        );
                      };
                      return (
                        <tr>
                          <td colSpan={8} style={{padding:"0 0 12px 24px",borderBottom:"0.5px solid rgba(41,53,93,.08)",background:"rgba(41,53,93,.02)"}}>
                            <div style={{fontSize:10,color:"#808080",padding:"6px 0 4px",fontStyle:"italic"}}>
                              {sortedAcctRows.length} accounts{domoTileFilter?" · filtered: "+domoTileFilter.replace(/_/g," "):""}
                            </div>
                            <table style={{width:"100%",borderCollapse:"collapse",fontSize:11,tableLayout:"fixed"}}>
                              <colgroup>
                                <col style={{width:"24%"}}/><col style={{width:"11%"}}/><col style={{width:"8%"}}/>
                                <col style={{width:"9%"}}/><col style={{width:"8%"}}/><col style={{width:"8%"}}/>
                                <col style={{width:"9%"}}/><col style={{width:"8%"}}/><col style={{width:"15%"}}/>
                              </colgroup>
                              <thead><tr style={{borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>
                                {aTh("acct","Account",false)}
                                <th style={{padding:"4px 8px 4px 0",textAlign:"left",fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500}}>EID</th>
                                {aTh("boq","BOQ")}
                                {aTh("cur","Current MRR")}
                                {aTh("increase","Increase")}
                                {aTh("decrease","Decrease")}
                                {aTh("cancelled","Cancelled")}
                                {aTh("noData","No Data")}
                                <th style={{padding:"4px 0 4px 24px",textAlign:"left",fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500}}>Retention %</th>
                              </tr></thead>
                              <tbody>
                                {sortedAcctRows.map((r,i)=>{
                                  const isIncrease = r.status==="increase";
                                  const isDecrease = r.status==="decrease";
                                  const isCancelled= r.status==="cancelled";
                                  const isNoData   = r.status==="no_data";
                                  return (
                                  <tr key={i} style={{borderTop:"0.5px solid rgba(41,53,93,.05)"}}>
                                    <td style={{padding:"5px 8px 5px 0",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.acct}</td>
                                    <td style={{padding:"5px 8px 5px 0",fontFamily:"monospace",fontSize:10,color:"#808080"}}>{r.eid}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right",color:"#808080"}}>{r.boq>0?fmt$(r.boq):"--"}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right"}}>{r.cur>0?fmt$(r.cur):"--"}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right",color:isIncrease?"#16a34a":"#808080"}}>{isIncrease?"+"+fmt$(r.delta):"--"}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right",color:isDecrease?"#dc2626":"#808080"}}>{isDecrease?"-"+fmt$(Math.abs(r.delta)):"--"}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right",color:isCancelled?"#d97706":"#808080"}}>{isCancelled?"-"+fmt$(r.boq):"--"}</td>
                                    <td style={{padding:"5px 8px 5px 0",textAlign:"right",color:isNoData?"#6b7280":"#808080"}}>{isNoData?"—":"--"}</td>
                                    <td style={{padding:"5px 0 5px 24px",fontWeight:600,color:retCol(r.ret)}}>{fmtPct(r.ret)}</td>
                                  </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </td>
                        </tr>
                      );
                    })()}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    );
  };

    const visibleAlerts = (churnAlerts||[]).filter(a => {
    if (managerCoaches) { const i=lk(a.csm); if(!(i&&managerCoaches.includes(i.c))) return false; }
    if (filterCoach) { const i=lk(a.csm); if(!(i&&i.c===filterCoach)) return false; }
    if (filterCSM && a.csm!==filterCSM) return false;
    return true;
  });
  const todayStr = new Date().toISOString().slice(0,10);
  const yesterStr = new Date(Date.now()-86400000).toISOString().slice(0,10);
  const recentAlerts = visibleAlerts.filter(a => a.date===todayStr||a.date===yesterStr);

  return (
    <div>
      {recentAlerts.length>0&&<div style={{background:"rgba(220,38,38,.05)",border:"1px solid rgba(220,38,38,.25)",borderRadius:12,padding:16,marginBottom:16}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
          <span style={{fontSize:16,fontWeight:600,color:"#dc2626"}}>🔴 New Churn Detected</span>
          <span style={{fontSize:10,fontWeight:600,padding:"2px 8px",borderRadius:20,background:"rgba(220,38,38,.12)",color:"#991b1b"}}>{recentAlerts.length} account{recentAlerts.length!==1?"s":""} · {recentAlerts[0].date}</span>
          <span style={{fontSize:11,color:"#808080",marginLeft:"auto",fontStyle:"italic"}}>Click a CSM to see detail</span>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8}}>
          {recentAlerts.map((a,i)=>{
            const col=a.type==="MC"?"#dc2626":"#7f1d1d";
            return <div key={i} onClick={()=>onSelectCSM&&onSelectCSM(a.csm)}
              style={{background:"#fff",border:"0.5px solid rgba(220,38,38,.2)",borderLeft:"3px solid "+col,borderRadius:8,padding:"8px 12px",cursor:"pointer"}}
              onMouseEnter={e=>e.currentTarget.style.background="#FFF5F5"}
              onMouseLeave={e=>e.currentTarget.style.background="#fff"}>
              <div style={{marginBottom:3}}><span style={{fontSize:10,fontWeight:600,padding:"1px 7px",borderRadius:20,background:a.type==="MC"?"rgba(220,38,38,.1)":"rgba(127,29,29,.1)",color:col}}>{a.type} Churn</span></div>
              <div style={{fontSize:12,fontWeight:500,color:"#29355D",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.account}</div>
              <div style={{fontSize:11,color:"#808080",marginTop:2}}>{a.csm}</div>
            </div>;
          })}
        </div>
      </div>}

      {/* Churned accounts full list modal — shared across all Book of Business sub-tabs */}
      {churnModal&&<div onClick={()=>setChurnModal(false)}
        style={{position:"fixed",inset:0,background:"rgba(0,0,0,.45)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
        <div onClick={e=>e.stopPropagation()}
          style={{background:"#fff",borderRadius:14,width:"100%",maxWidth:820,maxHeight:"85vh",display:"flex",flexDirection:"column",boxShadow:"0 8px 40px rgba(0,0,0,.18)"}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"18px 24px",borderBottom:"0.5px solid rgba(41,53,93,.1)"}}>
            <div>
              <div style={{fontSize:16,fontWeight:700,color:"#29355D"}}>All Churned Accounts</div>
              <div style={{fontSize:12,color:"#808080",marginTop:2}}>{allChurnedRows.length} accounts · {churnCSMs.length} CSMs</div>
            </div>
            <div style={{display:"flex",gap:8,alignItems:"center"}}>
              <button onClick={exportChurnCSV}
                style={{padding:"7px 16px",borderRadius:8,border:"none",background:"#29355D",color:"#fff",fontSize:12,fontWeight:600,cursor:"pointer",display:"flex",alignItems:"center",gap:6}}>
                ⬇ Export CSV
              </button>
              <button onClick={()=>setChurnModal(false)}
                style={{padding:"7px 12px",borderRadius:8,border:"0.5px solid rgba(41,53,93,.2)",background:"#fff",color:"#808080",fontSize:13,cursor:"pointer"}}>✕</button>
            </div>
          </div>
          <div style={{overflowY:"auto",flex:1,padding:"0 24px 16px"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,marginTop:12}}>
              <thead style={{position:"sticky",top:0,background:"#fff",zIndex:1}}>
                <tr>
                  {["CSM","Coach","Account","Enterprise ID","Products"].map(h=>(
                    <th key={h} style={{padding:"8px 8px 8px 0",borderBottom:"1.5px solid rgba(41,53,93,.1)",textAlign:"left",fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:600,whiteSpace:"nowrap"}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {allChurnedRows.map((r,i)=>(
                  <tr key={i} style={{borderBottom:"0.5px solid rgba(41,53,93,.05)"}}>
                    <td style={{padding:"8px 8px 8px 0",fontWeight:500,color:"#29355D",whiteSpace:"nowrap"}}>{r.csm}</td>
                    <td style={{padding:"8px 8px 8px 0",color:"#808080",whiteSpace:"nowrap"}}>{r.coach}</td>
                    <td style={{padding:"8px 8px 8px 0",maxWidth:220,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.account}</td>
                    <td style={{padding:"8px 8px 8px 0",fontFamily:"monospace",color:"#808080",fontSize:11,whiteSpace:"nowrap"}}>{r.eid}</td>
                    <td style={{padding:"8px 8px 8px 0"}}>
                      <div style={{display:"flex",flexWrap:"wrap",gap:3}}>
                        {r.products.map((p,j)=>(
                          <span key={j} style={{fontSize:10,padding:"1px 7px",borderRadius:20,background:"rgba(220,38,38,.08)",color:"#991b1b",fontWeight:500,whiteSpace:"nowrap"}}>{p}</span>
                        ))}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>}

      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16}}>
        <div style={{display:"flex",gap:2,background:"#ECEEF1",borderRadius:8,padding:3}}>
          {[["overview","Q2 Tracking"],["q3","Q3 Tracking"+( hasQ3?" ✓":"")],["domo","Domo Book of Business"]].map(([t,l])=>(
            <button key={t} onClick={()=>setBobTab(t)}
              style={{padding:"5px 14px",fontSize:12,fontWeight:500,border:"none",borderRadius:6,cursor:"pointer",
                background:bobTab===t?"#fff":"transparent",color:bobTab===t?"#29355D":t==="q3"&&hasQ3?"#16a34a":"#808080",
                boxShadow:bobTab===t?"0 1px 3px rgba(0,0,0,.08)":"none"}}>
              {l}
            </button>
          ))}
        </div>
        <span style={{fontSize:12,color:"#808080"}}>
          Monthly data · Overall: <strong style={{color:rCol(avgRet)}}>{fmtP(avgRet)}</strong> vs <strong>91%</strong> goal
        </span>
      </div>
      {bobTab==="overview" && renderQ2()}
      {bobTab==="q3"       && renderQ3()}
      {bobTab==="domo"     && renderDomoBoB()}
    </div>
  );
}

// ── PIN LOCK ───────────────────────────────────────────────────────────────
const USER_CREDS = {
  "6d43c88d47f13011d1a5265f0007c2e1403c9d98e23b51b35bffa00a6d5177ea": {"name": "Jake Baldwin", "role": "manager"},
  "daab3aa68185b677bd3c8d63f12e48eacb0851d98d635137e531fd579a452486": {"name": "Carrie Reece", "role": "manager"},
  "da5d26410ae112bfd2513b4d4eb0497ccf8eeb095cd613fee834e521705d8f20": {"name": "Aaron Taylor", "role": "coach"},
  "9d95bba7023609ee9e3da95b119f27d8f0a7c2412c1c773010f6bd5b8cea0d94": {"name": "Kendra Morelli", "role": "coach"},
  "689e02bf0aac5af7db656b4af619d2c212def8a3a28cd1254ed62627a4abc095": {"name": "Mia O'Dirling", "role": "coach"},
  "1fdc5156dff53e1ed8febfd0bfdfeedba271cde3e7bae67acd44e715a1badc78": {"name": "Chase Boyd", "role": "coach"},
  "d8c4d37261d7aaa4bbafe4ccfe334e09fbe181c84de22e9a561dfe02b0958aa0": {"name": "Elizabeth White", "role": "coach"},
  "cc28d893176b8cd8875730d1fe25151723d643debede6e093f742befdcf1ff0f": {"name": "Trisha Stalnaker", "role": "coach"},
  "df710e1f7dd04d8c460fe94df7ed35be7da7383452c05a243d13509b9aa74172": {"name": "Alejandro Rodriguez-Medina", "role": "csm"},
  "5f0e92362198dc68190b3e7ba044e6475ae036de70be6536a056de57290d3a03": {"name": "Anthony Yen", "role": "csm"},
  "4235d033441def62c1ea71655385fe6f2b4acd0458059c4ed2fc495e5129326e": {"name": "April Hall", "role": "csm"},
  "bdcc8eeedf70ba85a663d687e2c016ce09d1e5e221191d274bfd73eba44af5b2": {"name": "Ashley Shaffer", "role": "csm"},
  "01b347421433eccdad2e9f6885386c67057d51c6741f4aec5465010301e70791": {"name": "Ashley Vasquez Mena", "role": "csm"},
  "c6d6b2bf6044e53402d77694b8932a17dc7d383a32f2629f54b6b1311124088a": {"name": "Barbara Larrosa Presinal", "role": "csm"},
  "0354fa127deb6e3b93672221f6ff949087c73ca67a45bc88ddcbde6042d40a8a": {"name": "Chelsea Dingus", "role": "csm"},
  "cb850da6bf12d5c41fb11797bc8d3386b9b82014cf1f7eeaf65eaf917ba2f0db": {"name": "Damita Hill", "role": "csm"},
  "06ec44ccee3b23b9e067e96afb796d3b8d75ac7701185f649c61e405c282bdf8": {"name": "Darling Danais Santos Taveras", "role": "csm"},
  "d5c4f67a138cbcbcf840e8f78ce5280fc4f07beaf61c725ef109c8da727c3fd4": {"name": "Deivis Pena", "role": "csm"},
  "152bf02d1a3d3094d00cbce4933d5e38a402fa94ebad8c225e03fd98a3483441": {"name": "Dorka Frias Lantigua", "role": "csm"},
  "d16c1a995ccb00486f11d8c98fbbd6431012a553714fa0d55a62f5304c285878": {"name": "Elianny Tena Antigua", "role": "csm"},
  "c6b4c26e25da2dadec4164bd0a517057cfd5a1e873700a843ab37f4ef270d8ec": {"name": "Eric Johnson", "role": "csm"},
  "0b6632be5e33ac5bdbe9659a328959a19b42ce095057e04703c9c05c2ee139d2": {"name": "Felix Caba Jimenez", "role": "csm"},
  "aa71543206a78cb135b205658666838a3e83f16cd1c2b5b53372e656aedb22ea": {"name": "Florence Francois Nova", "role": "csm"},
  "cc2632d8e992a947a9b8af5be997cc4de0c4914067fe49e754435993606835f9": {"name": "Heidi Torres Uribe", "role": "csm"},
  "041669688dba6edca905acae363296e29902c37b43adcceb7113eb81e1ddb1a1": {"name": "Irina Larianni Molina Molina", "role": "csm"},
  "485d997d70d9483c517eb6f36b649bf37d701e15b0f398df6d01b478d164e82c": {"name": "Jathzelyn Elizabeth Fortuna Paulino", "role": "csm"},
  "7a81ab9d68095dc360c48b36423e5bd332a71385a1dec061bbbb6f2f3e3999a3": {"name": "Johnny Cornielle", "role": "csm"},
  "b99427238b182d603134bd9c45d0c7a2611e1ff58b9f37f91b613f0bac0949ef": {"name": "Joseph Guillermo Carmona Garcia", "role": "csm"},
  "48f599a9094eb9a4fcd2ff73dd158208d3a2e0d8769a32e3c3795fc8791a0a71": {"name": "Juan Liberato", "role": "csm"},
  "59712c920a3cc7ca6887718c0dc5124b1f3e47b57d04e9e0ec9186e6d0a62f6b": {"name": "Karen Capellan Tavarez", "role": "csm"},
  "c170ab55d53bb6560d7ba4b0ffa5bfb2ab2acab55906110b7d6ed444f80dc8da": {"name": "Karissa Hernandez", "role": "csm"},
  "a2541e84e912f4f5dd32ce131abdb6b1fdc9cf315790f10c90e299cab67daa47": {"name": "Karmita Turner", "role": "csm"},
  "60d523c99e38a30e4b889688cf2563ae6b745aa77b41c3d5b7a7fa9606aaa4cf": {"name": "Katelyn Ankrom", "role": "csm"},
  "d413a4731e3205dfd986119470725037a7454091a77868ed3b17dff737290929": {"name": "Kellie Lester", "role": "csm"},
  "69769ab78b2e30ed39fa27bb7f273a074ebebb9e2ab82b2f9c00f36f3abdb851": {"name": "Kennedy Sanchez", "role": "csm"},
  "3ca74779bdf0f0ba587f6a5dd2714505248a84002b1aeda3e63646fe0ed4b4fe": {"name": "Kyle Dye", "role": "csm"},
  "c3b95e6baf846dc6454640651e08d018b64a53f6f7d2f333375f2e3cb40fe4b7": {"name": "Lauren Carter", "role": "csm"},
  "bcaf29cf76c157164339236eb7b4a3038d4e0bd039eaa1f5919ab2efb1a23239": {"name": "Libby Booher", "role": "csm"},
  "aced70051e58a8841d88279bded07894af517957690986ff6ceff4f11870ae3d": {"name": "Luis Aguasvivas Peralta", "role": "csm"},
  "d0ea28283d465708b25553535043283e801e26da23030320d6b649f2b3e4e8c6": {"name": "Mark Velazquez", "role": "csm"},
  "7b4f6f46093c9812bff7745ec3bc0f70d49ad40f18ae3bba91f719f15c0571f7": {"name": "Merve (MJ) Brielmann", "role": "csm"},
  "187e96763937a4e2c3461ae642b88a030f513fb3e9cd7e44d99144cd23328f4b": {"name": "Michael Furlong", "role": "csm"},
  "33c62a2a336ae37ca20ef215f25803de63d9df76a391a35efece4ad27959324e": {"name": "Misti Dixon", "role": "csm"},
  "67a6e8768bea569a406f2c4689c0dd62926e01dfe00dcc240c111165ae98e749": {"name": "Misty Decatur", "role": "csm"},
  "80dd61da7ce1edccaa43d2d60207c482e397bdd7f7efe7ad2a222d35dde2bc8c": {"name": "Rafael Sencion Sencion", "role": "csm"},
  "bc7fe627d446aa7bf9d91c98487d414beaf56efaf7ba1748a354ca07850d889d": {"name": "Rossi Valerio Tejeda", "role": "csm"},
  "453fc8e8ed928121e245f03a32ba8981101989ee74dae86822144a3e4aecc508": {"name": "Saira Julian Guzman", "role": "csm"},
  "dcce7a9a1b433712b2bf9f5277947bbff7f4ac92b9007df413687b2c191cadd0": {"name": "Samuel Frias De Paula", "role": "csm"},
  "53e38ea20dcd0179a28f542304fceea1e2e36d00a8b31bb2dd7b0af44bd5765e": {"name": "Sarah Swanson", "role": "csm"},
  "ce0def455fe638a6c621de7f42b868e0af232f71d1664d691f89a63bcddc543d": {"name": "Sati Ananda Pimentel Malespin", "role": "csm"},
  "7b4261f5afa3de11408ed91980e35e585292e46cc265cd41dad35487f815c93d": {"name": "Scott Mather", "role": "csm"},
  "a39ff608d7c5bc24e03340701adfa5728b72b6b9933f56ffb477ec888ac5c0d6": {"name": "Stacy Roers", "role": "csm"},
  "e2688cb007b9ecd1db04c6345f854a594df036a8fd08c15cb058476a989615b4": {"name": "Steven Saunders", "role": "csm"},
  "6656cb1339533eeca99806086421e15deb9e38a25e40c5d2d62e822e4ae46997": {"name": "Taylor Kidd", "role": "csm"},
  "c5e109e0af901779a8dc7f0ceca77c2fae7fca22a3a96ad5e794a8f1707d30d6": {"name": "Victor Abner Moscoso Fernandez", "role": "csm"},
  "663f73f8af157c158248e72eb2d6548f7219b4231609b732ccc6484522628c6d": {"name": "Wilson Mercedes", "role": "csm"},
  "4fdc8d7d404bc07349ffce4cd89e1086a602d2d0333732a7b0c917314035492d": {"name": "Yessica Montero Urena", "role": "csm"},
  "b4170c25efd27ef305728ccb5276772a9ad321e43fa9fa03cb78e3a662a9aa43": {"name": "Yolanda Ramirez", "role": "csm"},
  "04e97796ac4d1ffe851398d3a25afad5f62916ef5ac22f6f10e83192dc15e8c3": {"name": "Dave Crisler", "role": "csm"},
  "cead18006a4de84ec2152071abe3deaf2bb386a00070f29f69c6e534c3d386f0": {"name": "Ellise Payne", "role": "csm"},
  "395e09e946b0043f8b3cc0758674aa20bf89234dd07e26b5b6351be7a861f759": {"name": "Indu Vijay", "role": "csm"},
  "3c06bc19eae1dddc102fdcc0ad552882adbf437ca3645f24b019ffec4328471b": {"name": "Matt Daly", "role": "csm"},
  "c08ee2a296e9cc80f185cf84596b17f6dc4e29df05c403f2b0fb077876289cb8": {"name": "Nikita Siepen-Bowers", "role": "csm"},
  "5bca1fc5f27e6d49017d2084b4f760431616e8b37f53134b9020c2add3b316a7": {"name": "Peter Manalac", "role": "csm"},
  "497e671df1de83a45eec09ac98062e8149b2f4cc8ca4e9cfbf89dc86a45404dd": {"name": "Sakshi Mahalwal", "role": "csm"},
  "a04f132c7c93df8b0d896a5334fd89ad933ea7747ff70cbb940722e9359a2235": {"name": "Sylvia Appla", "role": "csm"},
  "5313e5bf17148de844ff74be3663d47c6e361ca469b30a36337701233c89a15e": {"name": "Tracy-Ann Gaudencio", "role": "csm"},
  "f7a13fbd11a5bebce6bfcd0b1666545ff83d11bc39eb2e39c5a19775508af495": {"name": "Warda Gul", "role": "csm"},
  "604bdba4ae56af689c9c920d4b367e3c8935d567050efc6975cdbc098830af06": {"name": "Zoltan Rudolf", "role": "csm"},
};

function sha256(str) {
  // Simple synchronous SHA-256 using SubtleCrypto — returns a promise
  return crypto.subtle.digest("SHA-256", new TextEncoder().encode(str))
    .then(buf => Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,"0")).join(""));
}

function PinLock({onUnlock}) {
  const [val,setVal]=useState("");
  const [err,setErr]=useState(false);
  const [checking,setChecking]=useState(false);

  const check = async () => {
    if (!val.trim()) return;
    setChecking(true);
    try {
      const h = await sha256(val.trim());

      // Master PIN check (existing hash)
      if (val.trim() === PIN) {
        onUnlock({role:"master", name:""});
        return;
      }

      // Per-user PIN check
      const user = USER_CREDS[h];
      if (user) {
        onUnlock(user);
        return;
      }

      setErr(true); setVal(""); setTimeout(()=>setErr(false),1500);
    } finally { setChecking(false); }
  };

  return (
    <div style={{minHeight:"100vh",background:"#F4F6FB",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{background:"#fff",borderRadius:16,padding:"40px 36px",boxShadow:"0 4px 32px rgba(41,53,93,.12)",textAlign:"center",width:340}}>
        <div style={{width:56,height:56,borderRadius:"50%",background:"#29355D",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 20px",fontSize:24}}>🔒</div>
        <div style={{fontFamily:"Nunito,sans-serif",fontSize:20,fontWeight:800,color:"#29355D",marginBottom:6}}>CSM Coaching Dashboard</div>
        <div style={{fontSize:13,color:"#808080",marginBottom:24}}>Enter your PIN to continue</div>
        <input type="password" value={val} onChange={e=>setVal(e.target.value)} onKeyDown={e=>e.key==="Enter"&&check()}
          placeholder="Enter PIN" autoFocus
          style={{width:"100%",padding:"10px 14px",fontSize:15,borderRadius:10,border:"1px solid "+(err?"#dc2626":"rgba(41,53,93,.2)"),outline:"none",textAlign:"center",letterSpacing:4,marginBottom:12,fontFamily:"inherit"}}/>
        {err&&<div style={{color:"#dc2626",fontSize:12,marginBottom:10}}>Incorrect PIN — check with your coach</div>}
        <button onClick={check} disabled={checking}
          style={{width:"100%",padding:11,background:"#FF5000",border:"none",borderRadius:10,color:"#fff",fontSize:14,fontWeight:700,cursor:"pointer",fontFamily:"inherit",opacity:checking?0.7:1}}>
          {checking?"Checking...":"Sign In"}
        </button>
        <div style={{fontSize:11,color:"#808080",marginTop:16}}>Contact your coach if you need your PIN</div>
      </div>
    </div>
  );
}

// ── MAIN APP ───────────────────────────────────────────────────────────────
export default function App() {
  const [unlocked, setUnlocked] = useState(false);
  const [userSession, setUserSession] = useState({role:"master",name:""});
  const isCsmView = userSession.role==="csm";
  const [fontScale, setFontScale] = useState(()=>{try{return parseFloat(sessionStorage.getItem(FONT_KEY)||"1");}catch(e){return 1;}});
  const changeFontScale = (delta) => setFontScale(s=>{
    const next = Math.min(1.4, Math.max(0.8, Math.round((s+delta)*10)/10));
    try{sessionStorage.setItem(FONT_KEY, String(next));}catch(e){}
    return next;
  });
  const [csms, setCSMs] = useState([]);
  const [tab, setTab] = useState("coaching");
  const [filterManager, setFilterManager] = useState("");
  const [filterCoach, setFilterCoach] = useState("");
  const [filterCSM, setFilterCSM] = useState("");
  const [status, setStatus] = useState("loading");
  const [updatedAt, setUpdatedAt] = useState(null);
  const [history, setHistory] = useState([]);
  const [q3Current, setQ3Current] = useState({});
  const [q3Log, setQ3Log]         = useState([]);
  const [q3BobBoq, setQ3BobBoq]   = useState([]); // account-level BOQ rows
  const [q3BobCur, setQ3BobCur]   = useState([]); // account-level current rows
  const [domoBoq, setDomoBoq]     = useState([]); // fresh Domo BOQ export rows
  const [q3Supp,  setQ3Supp]      = useState([]); // supplemental SF revenue rows
  const [skippedCSMs, setSkippedCSMs] = useState([]);
  const [rawRev, setRawRev] = useState([]);
  const [bobData, setBobData] = useState(null);
  const [bobRaw,  setBobRaw]  = useState({bob:{},coachTotals:{},grand:null});
  const [liveBobDet, setLiveBobDet] = useState({});
  const [bobAdj, setBobAdj]         = useState({});
  const [callData, setCallData]     = useState({});
  const [qamc, setQamc]             = useState({});
  const [qass, setQass]             = useState({});
  const [mcChurn, setMcChurn] = useState({});
  const [bcChurn, setBcChurn] = useState({});
  const [churnAlerts, setChurnAlerts] = useState([]);

  // AI Coach panel state
  const [aiOpen,     setAiOpen]     = useState(false);
  const [aiLoading,  setAiLoading]  = useState(false);
  const [aiResponse, setAiResponse] = useState("");
  const [aiQuestion, setAiQuestion] = useState("");
  const [aiCustom,   setAiCustom]   = useState("");

  const allCSMNames = [...new Set(csms.map(c=>c.name))].sort();

  // Derive coach list for active manager filter
  const managerCoaches = filterManager
    ? (MANAGERS.find(m=>m.id===filterManager)?.coaches||[])
    : null;

  const filteredCSMs = csms.filter(c => {
    const i = lk(c.name);
    if (managerCoaches && !managerCoaches.includes(i&&i.c||c.coach)) return false;
    if (filterCoach && (i&&i.c||c.coach) !== filterCoach) return false;
    if (filterCSM && c.name !== filterCSM) return false;
    return true;
  });

  const activeCoaches = filterCoach ? COACHES.filter(c=>c.e===filterCoach)
    : filterCSM ? COACHES.filter(c=>{const i=lk(filterCSM);return i&&i.c===c.e;})
    : COACHES;

  // Full load on mount (all 6 tabs)
  useEffect(()=>{
    if (!unlocked) return;
    setStatus("loading");

    // Store non-revenue data so revenue polls can reuse it
    let latestEmail=[], latestCad=[], latestDue=[], latestOntime=[], latestHistory=[], latestSkipped=[], latestBob=[], latestBobDet=[], latestBobAdj=[], latestCalls=[], latestQaMc=[], latestQaSs=[], latestMcChurn=[], latestBcChurn=[], latestChurnAlerts=[];

    function loadAll() {
      return Promise.all([
        fetchCSV(CSV_REV),
        fetchCSV(CSV_EMAIL),
        fetchCSV(CSV_CAD),
        fetchCSV(CSV_DUE),
        fetchCSV(CSV_ONTIME),
        fetchCSV(CSV_HISTORY).catch(()=>[]),
        fetchCSV(CSV_SKIPPED).catch(()=>[]),
        fetchCSV(CSV_BOB).catch(()=>[]),
        fetchCSV(CSV_BOB_DET).catch(()=>[]),
        fetchCSV(CSV_BOB_ADJ).catch(()=>[]),
        fetchCSV(CSV_CALLS).catch(()=>[]),
        fetchCSV(CSV_QA_MC).catch(()=>[]),
        fetchCSV(CSV_QA_SS).catch(()=>[]),
        fetchCSV(CSV_MC_CHURN).catch(()=>[]),
        fetchCSV(CSV_BC_CHURN).catch(()=>[]),
        fetchCSV(CSV_CHURN_ALERTS).catch(()=>[]),
        CSV_Q3_RESULTS ? fetchCSV(CSV_Q3_RESULTS).catch(()=>[]) : Promise.resolve([]),
        CSV_Q3_LOG     ? fetchCSV(CSV_Q3_LOG).catch(()=>[])     : Promise.resolve([]),
        fetchCSV(CSV_Q3_BOQ).catch(()=>[]),
        fetchCSV(CSV_Q3_CUR).catch(()=>[]),
        fetchCSV(CSV_DOMO_BOQ).catch(()=>[]),
        fetchCSV(CSV_Q3_SUPP).catch(()=>[]),
      ]).then(([revRows, emailRows, cadRows, dueRows, ontimeRows, historyRows, skippedRows, bobRows, bobDetRows, bobAdjRows, callRows, qaMcRows, qaSSRows, mcRows, bcRows, churnAlertRows, q3CurrentRows, q3LogRows, q3BobBoqRows, q3BobCurRows, domoBoqRows, q3SuppRows]) => {
        latestEmail   = emailRows;
        latestCad     = cadRows;
        latestDue     = dueRows;
        latestOntime  = ontimeRows;
        latestHistory = historyRows;
        latestSkipped = skippedRows;
        latestBob         = bobRows;
        latestBobDet      = bobDetRows||[];
        latestBobAdj      = bobAdjRows||[];
        latestCalls       = callRows||[];
        latestQaMc        = qaMcRows||[];
        latestQaSs        = qaSSRows||[];
        latestMcChurn     = mcRows;
        latestBcChurn     = bcRows;
        latestChurnAlerts = churnAlertRows;
        setRawRev(revRows);
        let rev, email, cad, due, ontime, skipped, built;
        try { rev     = mapRev(revRows); }     catch(e) { console.error("mapRev failed:", e); rev = {}; }
        try { email   = mapEmail(emailRows); } catch(e) { console.error("mapEmail failed:", e); email = {}; }
        try { cad     = mapCadence(cadRows); } catch(e) { console.error("mapCadence failed:", e); cad = []; }
        try { due     = mapDue(dueRows); }     catch(e) { console.error("mapDue failed:", e); due = []; }
        try { ontime  = mapOnTime(ontimeRows); } catch(e) { console.error("mapOnTime failed:", e); ontime = {}; }
        try { skipped = mapSkipped(skippedRows); } catch(e) { console.error("mapSkipped failed:", e); skipped = {}; }
        try { built   = buildCSMs(rev, email, cad, due, ontime, skipped, bobRaw, mapChurn(mcRows), mapChurn(bcRows), mapBobDet(bobDetRows||[]), mapBobAdj(bobAdjRows||[])); }
        catch(e) { console.error("buildCSMs failed:", e); built = []; }
        setCSMs(built);
        setSkippedCSMs(built.filter(c=>c.skippedCount>0).sort((a,b)=>b.skippedCount-a.skippedCount));
        setBobRaw(mapBob(bobRows));
        if (bobDetRows&&bobDetRows.length>0) setLiveBobDet(mapBobDet(bobDetRows));
        if (bobAdjRows&&bobAdjRows.length>0) setBobAdj(mapBobAdj(bobAdjRows));
        if (callRows&&callRows.length>0) setCallData(mapCalls(callRows));
        if (qaMcRows&&qaMcRows.length>0) setQamc(mapQA(qaMcRows,"mc"));
        if (qaSSRows&&qaSSRows.length>0) setQass(mapQA(qaSSRows,"ss"));
        setMcChurn(mapChurn(mcRows));
        setBcChurn(mapChurn(bcRows));
        setChurnAlerts(mapChurnAlerts(churnAlertRows));
        setHistory(mapHistory(historyRows));
        if (q3CurrentRows&&q3CurrentRows.length>0) setQ3Current(mapQ3Current(q3CurrentRows));
        if (q3LogRows&&q3LogRows.length>0) setQ3Log(mapQ3Log(q3LogRows));
        if (q3BobBoqRows&&q3BobBoqRows.length>0) setQ3BobBoq(q3BobBoqRows);
        if (q3BobCurRows&&q3BobCurRows.length>0) setQ3BobCur(q3BobCurRows);
        if (domoBoqRows&&domoBoqRows.length>0) setDomoBoq(domoBoqRows);
        if (q3SuppRows&&q3SuppRows.length>0) setQ3Supp(q3SuppRows);
        setUpdatedAt(new Date().toLocaleTimeString());
        setStatus("ok");
        console.log("Full load complete");
      });
    }

    // Revenue-only refresh — rebuildCSMs with fresh rev + cached other data
    function refreshRevenue() {
      fetchCSV(CSV_REV).then(revRows => {
        setRawRev(revRows);
        const rev     = mapRev(revRows);
        const email   = mapEmail(latestEmail);
        const cad     = mapCadence(latestCad);
        const due     = mapDue(latestDue);
        const ontime  = mapOnTime(latestOntime);
        const skipped = mapSkipped(latestSkipped);
        const built   = buildCSMs(rev, email, cad, due, ontime, skipped, bobRaw, mapChurn(latestMcChurn), mapChurn(latestBcChurn), latestBobDet||{}, latestBobAdj||{});
        setCSMs(built);
        setSkippedCSMs(built.filter(c=>c.skippedCount>0).sort((a,b)=>b.skippedCount-a.skippedCount));
        setUpdatedAt(new Date().toLocaleTimeString());
        console.log("Revenue refreshed at", new Date().toLocaleTimeString());
      }).catch(err => console.warn("Revenue refresh failed:", err));
    }

    loadAll().catch(err => {
      console.error("Sheet load error:", err);
      setStatus("error");
    });

    // Poll revenue every 2 minutes (JotForm syncs live)
    const revInterval = setInterval(refreshRevenue, 2 * 60 * 1000);
    return () => clearInterval(revInterval);
  }, [unlocked]);

  // Build rich data context scoped to current filter
  // getDet available at App level for buildContext and DigestView
  const getDet = n => liveBobDet[n]||liveBobDet[norm(n)]||BOB_DETAIL[n]||BOB_DETAIL[norm(n)]||{};

  function buildContext() {
    const lines = [];
    const scope = filterCSM ? "CSM" : filterCoach ? "coach_team" : filterManager ? "manager_org" : "full_team";

    // mcChurn/bcChurn are {name: {canceled, accts}} objects — NOT arrays
    const getMc = n => { const k=Object.keys(mcChurn).find(k=>norm(k)===n||k===n); return k?mcChurn[k]:null; };
    const getBc = n => { const k=Object.keys(bcChurn).find(k=>norm(k)===n||k===n); return k?bcChurn[k]:null; };
    const getBob = n => {
      if (!bobRaw||!bobRaw.bob) return null;
      const k=Object.keys(bobRaw.bob).find(k=>norm(k)===n||k===n);
      return k?bobRaw.bob[k]:null;
    };
    const hasChurn = n => { const m=getMc(n),b=getBc(n); return (m&&m.canceled>0)||(b&&b.canceled>0); };

    // ── SINGLE CSM ──────────────────────────────────────────────────
    if (scope === "CSM") {
      const c = csms.find(x=>x.name===filterCSM);
      if (!c) return { scope, text: "No data found for "+filterCSM };
      const info = lk(c.name)||{};
      const coach = COACHES.find(x=>x.e===(info.c||c.coach));
      const det = getDet(c.name)||{};
      const bob = getBob(c.name)||{boq:c.bobBoq,lcm:c.bobLcm,net:c.bobNet,ret:c.bobRet};
      const mc = getMc(c.name);
      const bc = getBc(c.name);
      const ret = bob&&bob.boq>0&&bob.lcm!=null ? bob.lcm/bob.boq : (bob&&bob.ret!=null?bob.ret:null);

      lines.push("=== CSM PROFILE ===");
      lines.push("Name: "+c.name);
      lines.push("Coach: "+(coach?coach.n:"Unknown")+" | Team: "+(info.t||"Unknown")+" | Tier: "+(info.r||"Unknown"));
      lines.push("");
      lines.push("=== CADENCE & ENGAGEMENT ===");
      lines.push("Cadence completion: "+(c.cadCount>0?pp(c.cadPct)+" ("+c.cadCount+" tasks)":"No data"));
      lines.push("On-time tasks: "+(c.otTotal>=3?pp(c.otPct)+" ("+c.otOnTime+"/"+c.otTotal+" on time)":"Insufficient data"));
      lines.push("Overdue tasks: "+(c.overdueCount>0?c.overdueCount+" overdue":"None"));
      lines.push("Email: "+(c.sent>0?c.sent+" sent · "+pp(c.openRate)+" open rate · "+pp(c.replyRate)+" reply rate":"No email activity"));
      if (c.skippedCount>0) lines.push("Skipped cadences: "+c.skippedCount+(c.skippedFourthCount>0?", "+c.skippedFourthCount+" at 4th reschedule":""));
      if (c.skippedAccts&&c.skippedAccts.length>0) lines.push("Skipped accounts: "+c.skippedAccts.slice(0,5).map(a=>a.n).join(", "));
      lines.push("");
      lines.push("=== REVENUE ===");
      lines.push("This period: "+(c.rev>0?fd(c.rev):"None")+" | MRR: "+(c.mrr>0?fd(c.mrr):"None"));
      if (c.accts&&c.accts.length>0) lines.push("Accounts: "+c.accts.slice(0,5).map(a=>a.b+(a.m>0?" MRR "+fd(a.m):a.o>0?" OTR "+fd(a.o):"")).join(", "));
      lines.push("");
      lines.push("=== BOOK OF BUSINESS ===");
      lines.push("BOQ: "+(bob.boq?fd(bob.boq):"n/a")+" | Current: "+(bob.lcm?fd(bob.lcm):"n/a")+" | Net: "+(bob.net!=null?(bob.net>0?"+":"")+fd(bob.net):"n/a")+" | Retention: "+(ret!=null?pp(ret):"n/a"));
      if ((det.i||[]).length>0) lines.push("Billing increases ("+det.i.length+"): "+det.i.map(x=>(x.a||x.e)+" "+x.l+" +"+fd(x.n)).join(", "));
      if ((det.d||[]).length>0) lines.push("Billing decreases ("+det.d.length+"): "+det.d.map(x=>(x.a||x.e)+" "+x.l+" "+fd(x.n)).join(", "));
      if (mc&&mc.canceled>0) lines.push("MC churn ("+mc.canceled+"): "+(mc.accts||[]).slice(0,5).join(", "));
      if (bc&&bc.canceled>0) lines.push("BC churn ("+bc.canceled+"): "+(bc.accts||[]).slice(0,5).join(", "));
      if (!hasChurn(c.name)&&!(det.d||[]).length) lines.push("No churn or billing decreases this period.");
      lines.push("");

    // ── COACH TEAM ──────────────────────────────────────────────────
    } else if (scope === "coach_team") {
      const coach = COACHES.find(c=>c.e===filterCoach);
      lines.push("=== TEAM OVERVIEW: "+coach.n+" ("+coach.t+") ===");
      lines.push("CSM Count: "+filteredCSMs.length);
      lines.push("");
      filteredCSMs.forEach(c=>{
        const bob = getBob(c.name)||{net:c.bobNet,ret:c.bobRet,boq:c.bobBoq,lcm:c.bobLcm};
        const mc = getMc(c.name);
        const bc = getBc(c.name);
        const det = getDet(c.name)||{};
        const ret = bob&&bob.boq>0&&bob.lcm!=null ? bob.lcm/bob.boq : (bob&&bob.ret!=null?bob.ret:null);
        lines.push("── "+c.name);
        lines.push("   Revenue: "+(c.rev>0?fd(c.rev):"none")+" | Email open: "+(c.sent>0?pp(c.openRate):"n/a")+" | On-time: "+(c.otTotal>=3?pp(c.otPct):"n/a")+" | Overdue: "+(c.overdueCount||0));
        lines.push("   Cadence: "+(c.cadCount>0?pp(c.cadPct):"n/a")+" | BOB net: "+(bob.net!=null?(bob.net>0?"+":"")+fd(bob.net):"n/a")+" | Retention: "+(ret!=null?pp(ret):"n/a"));
        if (c.skippedCount>0) lines.push("   ⚠ "+c.skippedCount+" skipped"+(c.skippedFourthCount>0?", "+c.skippedFourthCount+" at 4th reschedule":""));
        const churnCount=((mc&&mc.canceled)||0)+((bc&&bc.canceled)||0);
        if (churnCount>0) lines.push("   ⚠ Churn: "+churnCount+" account(s)"+(mc&&mc.accts.length?" MC: "+mc.accts.slice(0,2).join(", "):""));
        if ((det.d||[]).length>0) lines.push("   ↓ "+det.d.length+" billing decrease(s)");
        lines.push("");
      });

    // ── MANAGER ORG ─────────────────────────────────────────────────
    } else if (scope === "manager_org") {
      const mgr = MANAGERS.find(m=>m.id===filterManager);
      lines.push("=== ORG OVERVIEW: "+mgr.n+" ===");
      const mgrCoaches = COACHES.filter(c=>mgr.coaches.includes(c.e));
      mgrCoaches.forEach(coach=>{
        const team = csms.filter(c=>(lk(c.name)&&lk(c.name).c===coach.e)||c.coach===coach.e);
        const totRev = team.reduce((s,c)=>s+c.rev,0);
        const otTeam = team.filter(c=>c.otTotal>=3);
        const avgOT = otTeam.length?otTeam.reduce((s,c)=>s+c.otPct,0)/otTeam.length:null;
        const churnCSMs = team.filter(c=>hasChurn(c.name));
        const skipCSMs = team.filter(c=>c.skippedCount>0);
        lines.push("COACH: "+coach.n+" ("+coach.t+") — "+team.length+" CSMs");
        lines.push("  Revenue: "+fd(totRev)+" | Avg on-time: "+(avgOT!=null?pp(avgOT):"n/a"));
        lines.push("  CSMs with churn: "+churnCSMs.length+" | CSMs with skipped cadences: "+skipCSMs.length);
        lines.push("");
      });

    // ── FULL TEAM ────────────────────────────────────────────────────
    } else {
      lines.push("=== FULL TEAM OVERVIEW — THRYV CSM ORG ===");
      lines.push("Total CSMs: "+csms.length);
      lines.push("");
      COACHES.forEach(coach=>{
        const team = csms.filter(c=>(lk(c.name)&&lk(c.name).c===coach.e)||c.coach===coach.e);
        const totRev = team.reduce((s,c)=>s+c.rev,0);
        const cadTeam = team.filter(c=>c.cadCount>0);
        const avgCad = cadTeam.length?cadTeam.reduce((s,c)=>s+c.cadPct,0)/cadTeam.length:null;
        const otTeam = team.filter(c=>c.otTotal>=3);
        const avgOT = otTeam.length?otTeam.reduce((s,c)=>s+c.otPct,0)/otTeam.length:null;
        const churnCSMs = team.filter(c=>hasChurn(c.name));
        const skipCSMs = team.filter(c=>c.skippedCount>0);
        const lowCad = cadTeam.filter(c=>c.cadPct<0.9);
        lines.push("COACH: "+coach.n+" ("+coach.t+") — "+team.length+" CSMs");
        lines.push("  Revenue: "+fd(totRev)+" | Cadence avg: "+(avgCad!=null?pp(avgCad):"n/a")+" | On-time avg: "+(avgOT!=null?pp(avgOT):"n/a"));
        lines.push("  CSMs needing cadence help ("+lowCad.length+"): "+(lowCad.map(c=>c.name).join(", ")||"none"));
        lines.push("  CSMs with churn: "+churnCSMs.length+" | CSMs with skips: "+skipCSMs.length);
        lines.push("");
      });
    }

    return { scope, text: lines.join("\n") };
  }

  function openAI() {
    setAiOpen(true);
    setAiResponse("");
    setAiQuestion("");
    setAiCustom("");
  }

  // Build prompt and open Claude.ai in a new tab with it pre-copied to clipboard
  async function runAI(questionType) {
    let scope = "CSM", ctx = "";
    try {
      const result = buildContext();
      scope = result.scope;
      ctx = result.text;
    } catch(e) {
      console.error("buildContext error:", e);
      ctx = "Error loading context: " + e.message;
    }

    const scopeLabel = scope==="CSM" ? "this CSM ("+filterCSM+")"
      : scope==="coach_team" ? "this coaching team"
      : scope==="manager_org" ? "this manager's org"
      : "the full CSM team";

    const prompts = {
      coaching:
`You are an expert CSM coaching advisor at Thryv, a SaaS company. You have been given performance data for ${scopeLabel}.
Give a manager sharp, actionable coaching guidance they can use immediately.
Format your response with clear sections:
1. ✅ STRENGTHS — what's going well (be specific, cite numbers)
2. 🎯 TOP COACHING PRIORITIES — 2-3 focus areas with specific talking points for the next 1:1
3. ⚡ THIS WEEK'S ACTION ITEMS — concrete next steps (who does what by when)
4. 🚩 RED FLAGS — anything that needs urgent attention
Be direct and specific. Use the actual names and numbers from the data. Keep it under 400 words.

Here is the data:

${ctx}`,

      churn:
`You are an expert CSM retention analyst at Thryv, a SaaS company. You have been given BOB and churn data for ${scopeLabel}.
Identify churn risk and give a manager concrete retention guidance.
Format your response with clear sections:
1. 🔴 HIGH RISK — accounts or CSMs with active churn or billing decreases (cite EIDs and amounts)
2. 🟡 WATCH LIST — early warning signs (skipped cadences, low retention %, declining engagement)
3. 💬 RETENTION TALKING POINTS — specific things to say/do to save at-risk accounts
4. 📋 IMMEDIATE ACTIONS — what the manager or CSM should do this week
Be specific with account IDs and dollar amounts. Under 400 words.

Here is the data:

${ctx}`,

      revenue:
`You are an expert CSM revenue growth advisor at Thryv, a SaaS company. You have been given revenue and BOB data for ${scopeLabel}.
Surface upsell and expansion opportunities and give a manager a growth plan.
Format your response with clear sections:
1. 💰 CURRENT REVENUE SNAPSHOT — summarize what's happening (total, MRR, OTR, billing increases)
2. 🚀 BEST EXPANSION OPPORTUNITIES — which accounts or CSMs have the most upside and why
3. 🎯 UPSELL TALKING POINTS — specific products or conversations to have (SEO, Thryv Leads, Social, etc.)
4. 📋 THIS WEEK'S REVENUE ACTIONS — 3 concrete things to do to move the number
Cite actual account names or EIDs and dollar amounts where available. Under 400 words.

Here is the data:

${ctx}`,

      custom:
`You are an expert CSM coaching and retention advisor at Thryv, a SaaS company. You have been given performance data for ${scopeLabel}.
Answer the question below directly and concisely, drawing on the data provided.
Be specific — cite actual names, numbers, and account IDs. Avoid generic advice.
Keep your response under 350 words.

Here is the data:

${ctx}

My question: ${aiCustom}`,
    };

    const fullPrompt = prompts[questionType] || prompts.coaching;

    // Copy prompt to clipboard — always show confirmation even if clipboard fails
    let copied = false;
    try {
      await navigator.clipboard.writeText(fullPrompt);
      copied = true;
    } catch(e) {
      try {
        const el = document.createElement("textarea");
        el.value = fullPrompt;
        el.style.position = "fixed"; el.style.opacity = "0";
        document.body.appendChild(el);
        el.focus(); el.select();
        copied = document.execCommand("copy");
        document.body.removeChild(el);
      } catch(e2) { console.error("Clipboard fallback failed:", e2); }
    }

    setAiQuestion(questionType);
    setAiResponse(copied ? "__copied__" : "__manual__");
  }

  if (!unlocked) return <PinLock onUnlock={user=>{
    setUnlocked(true);
    setUserSession(user);
    // Check for deep link tab param (e.g. ?tab=digest from email)
    const _urlTab = new URLSearchParams(window.location.search).get("tab");
    if (_urlTab) setTab(_urlTab);
    // Auto-filter to this CSM if role=csm
    if (user.role==="csm") setFilterCSM(user.name);
    // Auto-filter to coach's team if role=coach
    if (user.role==="coach") {
      const c = COACHES.find(c=>c.n===user.name);
      if (c) setFilterCoach(c.e);
    }
  }}/>;

  const aiLabel = filterCSM ? "🤖 AI: "+dispName(filterCSM).split(" ")[0]
    : filterCoach ? "🤖 AI: "+COACHES.find(c=>c.e===filterCoach).n.split(" ")[0]+"'s team"
    : filterManager ? "🤖 AI: "+MANAGERS.find(m=>m.id===filterManager).n.split(" ")[0]+"'s org"
    : "🤖 AI Coach";

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
            {status==="ok"&&<span style={{fontSize:11,padding:"4px 10px",borderRadius:20,background:"rgba(22,163,74,.25)",color:"#86efac"}}>✓ Live · Revenue syncs every 2 min{updatedAt?" · "+updatedAt:""}</span>}
            {status==="error"&&<span style={{fontSize:11,padding:"4px 10px",borderRadius:20,background:"rgba(220,38,38,.25)",color:"#fca5a5"}}>✗ Sync error</span>}
            <div style={{display:"flex",alignItems:"center",gap:2,background:"rgba(255,255,255,.08)",borderRadius:6,padding:"3px 6px"}}>
              <button onClick={()=>changeFontScale(-0.1)} disabled={fontScale<=0.8}
                style={{background:"none",border:"none",color:"rgba(255,255,255,.7)",fontSize:14,fontWeight:500,cursor:"pointer",padding:"2px 7px",borderRadius:4,lineHeight:1,opacity:fontScale<=0.8?0.4:1}}>A−</button>
              <span style={{fontSize:11,color:"rgba(255,255,255,.5)",minWidth:30,textAlign:"center"}}>{Math.round(fontScale*100)}%</span>
              <button onClick={()=>changeFontScale(0.1)} disabled={fontScale>=1.4}
                style={{background:"none",border:"none",color:"rgba(255,255,255,.9)",fontSize:16,fontWeight:500,cursor:"pointer",padding:"2px 7px",borderRadius:4,lineHeight:1,opacity:fontScale>=1.4?0.4:1}}>A+</button>
            </div>
                        <button onClick={()=>{
              const subject = encodeURIComponent("CSM Dashboard Feedback — "+(new Date().toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"})));
              const body = encodeURIComponent("Hi Jacob,\n\nHere\'s some feedback on the CSM Coaching Dashboard:\n\n[Please describe your feedback here]\n\nThanks");
              window.open("mailto:jacob.baldwin@thryv.com?subject="+subject+"&body="+body);
            }} style={{background:"transparent",border:"0.5px solid rgba(255,255,255,.35)",color:"rgba(255,255,255,.8)",fontSize:12,fontWeight:500,padding:"7px 14px",borderRadius:6,cursor:"pointer"}}>
              ✉ Feedback
            </button>
            {hasData&&<button onClick={openAI} style={{background:"#FF5000",border:"none",color:"#fff",fontSize:12,fontWeight:500,padding:"7px 14px",borderRadius:6,cursor:"pointer"}}>{aiLabel} ↗</button>}
          </div>
        </div>
        <div style={{display:"flex",alignItems:"stretch",padding:"0 24px"}}>
          {["coaching","overview","digest","revenue","bob","leaderboard","activity","trends"].filter(t=>!isCsmView||(t!=="leaderboard"&&t!=="trends")).map(t=>(
            <button key={t} onClick={()=>setTab(t)}
              style={{padding:"10px 18px",fontSize:13,fontWeight:500,color:tab===t?"#fff":"rgba(255,255,255,.55)",background:"transparent",border:"none",cursor:"pointer",borderBottom:tab===t?"3px solid #FF5000":"3px solid transparent",whiteSpace:"nowrap"}}>
              {t==="coaching"?"Coaching":t==="digest"?"📋 Daily Digest":t==="trends"?"📈 Trends":t==="revenue"?"💰 Revenue":t==="bob"?"📋 Book of Business":t.charAt(0).toUpperCase()+t.slice(1)}
            </button>
          ))}
        </div>
      </div>

      {/* FILTERS */}
      {hasData&&(
        <div style={{background:"#fff",borderBottom:"0.5px solid rgba(41,53,93,.08)",padding:"8px 24px",display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          {!isCsmView&&<select value={filterManager} onChange={e=>{setFilterManager(e.target.value);setFilterCoach("");setFilterCSM("");}}
            style={{fontSize:12,fontWeight:500,padding:"5px 10px",borderRadius:8,border:"0.5px solid "+(filterManager?"#FF5000":"rgba(41,53,93,.15)"),background:"#F4F6FB",color:filterManager?"#FF5000":"#29355D",cursor:"pointer"}}>
            <option value="">All managers</option>
            {MANAGERS.map(m=><option key={m.id} value={m.id}>{m.n}</option>)}
          </select>}
          {!isCsmView&&<select value={filterCoach} onChange={e=>{setFilterCoach(e.target.value);setFilterCSM("");setFilterManager("");}}
            style={{fontSize:12,fontWeight:500,padding:"5px 10px",borderRadius:8,border:"0.5px solid "+(filterCoach?"#FF5000":"rgba(41,53,93,.15)"),background:"#F4F6FB",color:filterCoach?"#FF5000":"#29355D",cursor:"pointer"}}>
            <option value="">All coaches</option>
            {(managerCoaches?COACHES.filter(c=>managerCoaches.includes(c.e)):COACHES).map(c=><option key={c.e} value={c.e}>{c.n}</option>)}
          </select>}
          {!isCsmView&&<select value={filterCSM} onChange={e=>{setFilterCSM(e.target.value);setFilterCoach("");setFilterManager("");}}
            style={{fontSize:12,fontWeight:500,padding:"5px 10px",borderRadius:8,border:"0.5px solid "+(filterCSM?"#FF5000":"rgba(41,53,93,.15)"),background:"#F4F6FB",color:filterCSM?"#FF5000":"#29355D",cursor:"pointer"}}>
            <option value="">All CSMs</option>
            {allCSMNames.filter(n=>{
              if(!managerCoaches&&!filterCoach) return true;
              const i=lk(n);
              if(managerCoaches&&!managerCoaches.includes(i&&i.c)) return false;
              if(filterCoach&&(i&&i.c)!==filterCoach) return false;
              return true;
            }).map(n=><option key={n} value={n}>{n}</option>)}
          </select>}
          {isCsmView&&<div style={{fontSize:13,fontWeight:600,color:"#29355D",padding:"4px 10px",borderRadius:8,background:"#F4F6FB",border:"0.5px solid rgba(41,53,93,.15)"}}>{dispName(userSession.name)}</div>}
          {(filterManager||filterCoach||filterCSM)&&(
            <span style={{background:"#FF5000",color:"#fff",fontSize:11,fontWeight:500,padding:"4px 10px",borderRadius:20,display:"inline-flex",alignItems:"center",gap:6}}>
              {dispName(filterCSM)||filterCoach&&COACHES.find(c=>c.e===filterCoach)?.n||filterManager&&MANAGERS.find(m=>m.id===filterManager)?.n}
              {!isCsmView&&<button onClick={()=>{setFilterManager("");setFilterCoach("");setFilterCSM("");}} style={{background:"none",border:"none",color:"rgba(255,255,255,.8)",cursor:"pointer",fontSize:14,lineHeight:1,padding:0}}>✕</button>}
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
        <div style={{padding:"20px 24px",zoom:fontScale}}>
          {tab==="coaching"&&<CoachingView csms={filteredCSMs} coach={filterCoach} onSelectCSM={selectCSMFn} onSelectCoach={e=>{setFilterCoach(e);setFilterCSM("");}} onClear={()=>{setFilterCoach("");setFilterCSM("");}} skippedCSMs={skippedCSMs.filter(c=>{const i=lk(c.name);if(managerCoaches&&!(i&&managerCoaches.includes(i.c)))return false;if(filterCoach&&(i&&i.c)!==filterCoach)return false;if(filterCSM&&c.name!==filterCSM)return false;return true;})} bobRaw={bobRaw} mcChurn={mcChurn} bcChurn={bcChurn} liveBobDet={liveBobDet} isCsmView={isCsmView} bobAdj={bobAdj}/>}
          {tab==="digest"&&<DigestView csms={filteredCSMs} filterCoach={filterCoach} filterCSM={filterCSM}
            isCsmView={isCsmView} bobRaw={bobRaw} mcChurn={mcChurn} bcChurn={bcChurn}
            liveBobDet={liveBobDet} callData={callData} qamc={qamc} qass={qass} history={history}
            skippedCSMs={skippedCSMs.filter(c=>{const i=lk(c.name);if(filterCoach&&(i&&i.c)!==filterCoach)return false;if(filterCSM&&c.name!==filterCSM)return false;return true;})}
            bobAdj={bobAdj} getDet={getDet}/>}
          {tab==="overview"&&<OverviewView csms={filteredCSMs} allCSMs={csms} bobRaw={bobRaw} bobAdj={bobAdj} history={history} callData={callData} filterCoach={filterCoach} filterCSM={filterCSM} managerCoaches={managerCoaches}/>}
          {tab==="leaderboard"&&<LeaderboardView csms={filteredCSMs} bobRaw={bobRaw}/>}
          {tab==="activity"&&<ActivityView csms={filteredCSMs}/>}
          {tab==="revenue"&&<RevenueView rawRev={rawRev} csms={filteredCSMs} filterCoach={filterCoach} filterCSM={filterCSM} managerCoaches={managerCoaches}/>}
          {tab==="bob"&&<BobView filterCoach={filterCoach} filterCSM={filterCSM} managerCoaches={managerCoaches} bobRaw={bobRaw} mcChurn={mcChurn} bcChurn={bcChurn} churnAlerts={churnAlerts} onSelectCSM={selectCSMFn} liveBobDet={liveBobDet} bobAdj={bobAdj} q3Current={q3Current} q3Log={q3Log} q3BobBoq={q3BobBoq} q3BobCur={q3BobCur} domoBoq={domoBoq} q3Supp={q3Supp}/>}
          {tab==="trends"&&<TrendsView history={history} csms={filteredCSMs} filterCoach={filterCoach} filterCSM={filterCSM} callData={callData} qamc={qamc} qass={qass}/>}
        </div>
      )}

      {/* ── AI COACH PANEL ───────────────────────────────────────────── */}
      {aiOpen&&<div style={{position:"fixed",top:0,right:0,height:"100vh",width:420,background:"#fff",boxShadow:"-4px 0 24px rgba(0,0,0,.12)",zIndex:1000,display:"flex",flexDirection:"column",fontFamily:"Nunito Sans,sans-serif"}}>
        {/* Header */}
        <div style={{background:"#29355D",padding:"16px 20px",display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
          <div>
            <div style={{color:"#fff",fontSize:14,fontWeight:600}}>🤖 AI Coach</div>
            <div style={{color:"rgba(255,255,255,.6)",fontSize:11,marginTop:2}}>
              {filterCSM?"Focused on: "+filterCSM
               :filterCoach?"Focused on: "+COACHES.find(c=>c.e===filterCoach)?.n+"'s team"
               :filterManager?"Focused on: "+MANAGERS.find(m=>m.id===filterManager)?.n+"'s org"
               :"Full team overview"}
            </div>
          </div>
          <button onClick={()=>setAiOpen(false)} style={{background:"rgba(255,255,255,.1)",border:"none",color:"#fff",fontSize:18,lineHeight:1,padding:"4px 10px",borderRadius:6,cursor:"pointer"}}>×</button>
        </div>

        {/* Question picker */}
        {!aiResponse&&!aiLoading&&<div style={{padding:20,flex:1,overflowY:"auto"}}>
          <div style={{fontSize:11,color:"#808080",marginBottom:4,fontWeight:600,textTransform:"uppercase",letterSpacing:".05em"}}>What would you like to explore?</div>
          <div style={{fontSize:11,color:"#808080",marginBottom:14}}>Select a question → Claude opens with your data pre-loaded.</div>
          {[
            {id:"coaching", icon:"🎯", label:"Coaching priorities",    sub:"Top focus areas and 1:1 talking points"},
            {id:"churn",    icon:"⚠️",  label:"Churn risk analysis",    sub:"At-risk accounts and retention concerns"},
            {id:"revenue",  icon:"💰", label:"Revenue opportunities",  sub:"Upsell, expansion, and growth potential"},
            {id:"custom",   icon:"✏️", label:"Ask your own question",  sub:"Type a specific coaching question"},
          ].map(q=>(
            <div key={q.id}
              onClick={()=>{if(q.id!=="custom")runAI(q.id);else setAiQuestion(q.id);}}
              style={{padding:"14px 16px",borderRadius:10,border:"0.5px solid rgba(41,53,93,.12)",
                background:aiQuestion===q.id?"#F4F6FB":"#fff",marginBottom:8,cursor:"pointer"}}>
              <div style={{display:"flex",alignItems:"center",gap:10}}>
                <span style={{fontSize:20}}>{q.icon}</span>
                <div style={{minWidth:0}}>
                  <div style={{fontSize:13,fontWeight:600,color:"#29355D"}}>{q.label}</div>
                  <div style={{fontSize:11,color:"#808080",marginTop:2}}>{q.sub}</div>
                </div>
              </div>
              {q.id==="custom"&&aiQuestion==="custom"&&(
                <div style={{marginTop:12}} onClick={e=>e.stopPropagation()}>
                  <textarea value={aiCustom} onChange={e=>setAiCustom(e.target.value)}
                    placeholder="e.g. What should I focus on in my next 1:1 with this CSM?"
                    style={{width:"100%",padding:"10px 12px",borderRadius:8,border:"0.5px solid rgba(41,53,93,.2)",
                      fontSize:12,fontFamily:"Nunito Sans,sans-serif",resize:"vertical",minHeight:80,
                      outline:"none",boxSizing:"border-box"}}
                    autoFocus/>
                  <button onClick={()=>runAI("custom")} disabled={!aiCustom.trim()}
                    style={{marginTop:8,width:"100%",padding:"10px",borderRadius:8,border:"none",
                      background:aiCustom.trim()?"#FF5000":"#e5e7eb",color:"#fff",fontSize:13,
                      fontWeight:600,cursor:aiCustom.trim()?"pointer":"not-allowed"}}>
                    Ask →
                  </button>
                </div>
              )}
            </div>
          ))}
        </div>}

        {/* Copied confirmation */}
        {aiResponse==="__copied__"&&<div style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:16,padding:24,textAlign:"center"}}>
          <div style={{fontSize:48}}>📋</div>
          <div style={{fontSize:16,fontWeight:700,color:"#29355D"}}>Prompt copied!</div>
          <div style={{fontSize:12,color:"#808080",lineHeight:1.7,maxWidth:300}}>
            Your coaching prompt and all live dashboard data has been copied to your clipboard.
          </div>
          <div style={{width:"100%",padding:"16px",borderRadius:10,background:"#29355D",fontSize:13,color:"#fff",lineHeight:2,textAlign:"left"}}>
            <div style={{fontWeight:700,marginBottom:4}}>Next steps:</div>
            <div>1. Open <a href="https://claude.ai/new" target="_blank" rel="noreferrer" style={{color:"#FF5000",fontWeight:700}}>claude.ai/new</a></div>
            <div>2. Click the message box</div>
            <div>3. Press <strong>Ctrl+V</strong> (⌘V on Mac) to paste</div>
            <div>4. Press <strong>Enter</strong> to send</div>
          </div>
          <button onClick={()=>{setAiResponse("");setAiQuestion("");setAiCustom("");}}
            style={{width:"100%",padding:"10px",borderRadius:10,border:"0.5px solid rgba(41,53,93,.15)",
              background:"#fff",color:"#29355D",fontSize:12,fontWeight:500,cursor:"pointer"}}>
            ← Ask a different question
          </button>
        </div>}
      </div>}
        {/* Manual copy fallback */}
        {aiResponse==="__manual__"&&<div style={{flex:1,display:"flex",flexDirection:"column",gap:12,padding:20}}>
          <div style={{fontSize:15,fontWeight:700,color:"#29355D"}}>📋 Clipboard unavailable</div>
          <div style={{fontSize:12,color:"#808080"}}>Your browser blocked auto-copy. Select all the text below and copy it manually (Ctrl+A, Ctrl+C).</div>
          <textarea readOnly value={aiCustom||""} onClick={e=>e.target.select()}
            style={{flex:1,minHeight:200,fontSize:11,padding:10,borderRadius:8,border:"1px solid rgba(41,53,93,.15)",resize:"none",fontFamily:"monospace",color:"#29355D"}}/>
          <a href="https://claude.ai/new" target="_blank" rel="noreferrer"
            style={{display:"block",padding:"10px",borderRadius:10,background:"#FF5000",color:"#fff",fontSize:13,fontWeight:600,textAlign:"center",textDecoration:"none"}}>
            Open Claude.ai ↗
          </a>
          <button onClick={()=>{setAiResponse("");setAiQuestion("");}}
            style={{padding:"8px",borderRadius:10,border:"0.5px solid rgba(41,53,93,.15)",background:"#fff",color:"#29355D",fontSize:12,cursor:"pointer"}}>
            ← Back
          </button>
        </div>}
      {aiOpen&&<div onClick={()=>setAiOpen(false)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,.2)",zIndex:999}}/>}
    </div>
  );
}
