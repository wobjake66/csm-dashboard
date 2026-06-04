import React, { useState, useEffect } from "react";
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
const CSV_BC_CHURN= "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=295771282&single=true&output=csv"; // BC churn by coach/rep
const CSV_MC_CHURN= "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=1002996767&single=true&output=csv"; // MC churn by coach/rep

const CAD_ACCTS={"Taylor Kidd":[{"n":"F Lees Tax & Accounting Services","ott":3,"oto":2},{"n":"Liz Bienstock, Realtor","ott":1,"oto":0},{"n":"Nixdorf Tree Service","ott":1,"oto":0},{"n":"Torrey Tucker Electric LLC","ott":1,"oto":1},{"n":"Uniquely Yours Screen Printing","ott":3,"oto":3}],"April Hall":[{"n":"JaysonRachel Carter","ott":1,"oto":0},{"n":"M & C Homecare","d":[{"t":"Highlight Video","due":"5/29/2026","ov":true,"nw":false}]},{"n":"TAG Construction","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]}],"Zoltan Rudolf":[{"n":"Budget Screens & Awnings","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Classic Steamboat Cruises","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Clean Planet West Auckand","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"North East Survey Design","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Perfect Fit Interior","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"VISION INSTALLATIONS LIMITED","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]}],"Libby Booher":[{"n":"Advanced Acupuncture And Chinese Herbal Clinic","ott":1,"oto":0},{"n":"Battle Ground Personal Training","ott":3,"oto":1},{"n":"CAMPBELL REMODELING AND MAINTENANCE LLC","ott":3,"oto":1},{"n":"Discount Auto Glass","ott":3,"oto":1},{"n":"Hawaii Food Products Inc","ott":1,"oto":0},{"n":"Motorcycle Rider Training","ott":3,"oto":1},{"n":"Zebra Stripes Child Care & Preschool","ott":1,"oto":1}],"Indu Vijay":[{"n":"Blinds Hub","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Bold Insurance","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Byte Size Labs","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Oz Seals","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Tranquility Pools","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Luis Aguasvivas Peralta":[{"n":"Mission Success Solutions LLC","ott":1,"oto":1}],"Mark Velazquez":[{"n":"Antesberger Plumbing","ott":1,"oto":1},{"n":"CJ Construction","ott":1,"oto":1},{"n":"CJS Heating and Air","ott":1,"oto":0},{"n":"FYZICAL Therapy & Balance Center of Little Silver","ott":1,"oto":0},{"n":"Home Pro Masters","ott":1,"oto":0},{"n":"Korte Does It All, Inc.","ott":1,"oto":0},{"n":"Legacy Heating and Air, Inc.","ott":1,"oto":0},{"n":"Newtown Heating & Air Conditioning Inc","ott":3,"oto":3},{"n":"Precision Today Plumbing Heating Cooling Electrical","ott":1,"oto":0},{"n":"Stitches and Screens","ott":2,"oto":1}],"Sylvia Appla":[{"n":"A F Isaac Suveying Pty Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"AESTHETIC DENTAL LIMITED","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"AK Painting","ott":1,"oto":0},{"n":"ASAP Security and Protection","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"ATOM TREE SPECIALISTS LIMITED","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Acland Street Physiotherapy","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"All Four x 4 Spares","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Allen's Interiors & Building","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"BYC Dental Pty Ltd ATF BYC Dental Trust","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Bayview Denture Care","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Cavana Services","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Country Tyres Yass","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Darwin Family Law","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Dons Mobile Marine","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"EK GLASS LIMITED","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Earth 2 Ocean Communications","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Exacte Advisors","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Exo Graphics","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Fencemen Fencing Contractors Ltd.","d":[{"t":"Email","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Flash Roofing Supplies","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Gold Coast Hypnotherapy","ott":1,"oto":0},{"n":"HERRON TODD WHITE","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Hyperspace Architecture","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Informed Decisions Consultancy Limited","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"J S De Rooy","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"KITCHEN & APPLIANCES SOUTHLAND LIMITED","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"MURRAY BROWN ELECTRICAL (1996)","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Nerang Stockfeed","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Northlane Welding Services","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Nudge Osteopathy","d":[{"t":"Email","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Oxen Lawyers","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"PNJ Hire","ott":1,"oto":0},{"n":"Pearla Plumbing & Electrical","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Penrith Engine Services","d":[{"t":"Email","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Perth Taxi Booking","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"PowerSmart Heat Pumps","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Quickturn Pty Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Red Roo Solutions","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"SHRI METALS MERCHANT PTY LTD","ott":1,"oto":0},{"n":"Scott Electrics","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Shubbs International Pty Ltd","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Smart Thread Solutions","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Straightline Guttering Pty Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Streamline Drains & Pipelines (NSW) PL","d":[{"t":"Email","due":"6/1/2026","ov":true,"nw":true}]},{"n":"The Baker's Den Bakery Cafe","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Village Carpet Care","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Voodoo Rod And Custom","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"WORKSTORE SA LIMITED","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Wagga Scrap Metals","d":[{"t":"Email","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Wide Bay Memorials","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Wollongong Auto Excellence","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]}],"Rafael Sencion Sencion":[{"n":"5 Star Fence","ott":1,"oto":1},{"n":"D & D Overhead Door LLC","ott":1,"oto":1},{"n":"Grow With Us Learning Centers","ott":1,"oto":1},{"n":"HVAC Pros Los Angeles","ott":1,"oto":1},{"n":"Pampered Paws Pet Grooming","ott":1,"oto":1},{"n":"Patrick Alley Handyman","ott":1,"oto":1},{"n":"Technology Networks","ott":1,"oto":1},{"n":"Uptown General Aesthetic Dentistry P A","ott":1,"oto":1},{"n":"Woodforest Family Chiropractic Clinic","ott":1,"oto":1}],"Katelyn Ankrom":[{"n":"Aberdeen Cemetery Assoc","ott":1,"oto":0},{"n":"Gleason's Salinas RV","ott":1,"oto":0},{"n":"Neilson Roy Plumbing","ott":1,"oto":0}],"Kellie Lester":[{"n":"Chatham Laschober","ott":1,"oto":1},{"n":"D &T Power Washing","ott":3,"oto":3},{"n":"East Cooper Lock & Safe","ott":1,"oto":1},{"n":"Lewis Brothers Inc","ott":1,"oto":1},{"n":"Mansfield Well Drilling Inc","ott":1,"oto":0},{"n":"Martin Landscaping","ott":1,"oto":1},{"n":"Metro Appliance Repair","ott":3,"oto":3},{"n":"Mikes Seal Coating & Services","ott":1,"oto":0},{"n":"Pleasure Pools","ott":2,"oto":2},{"n":"WBC Carpentry Corp","ott":1,"oto":0}],"Ashley Shaffer":[{"n":"A Dion & Son Floor Contractors","ott":1,"oto":0},{"n":"Absolute Roofing & Remodeling","ott":3,"oto":3},{"n":"Alaska Urgent Care LLC","ott":1,"oto":1},{"n":"Benny Electric Inc Benny Electric","ott":3,"oto":3},{"n":"Body Sculpt Skin Laser","ott":3,"oto":3},{"n":"Cement Hill Storage","ott":1,"oto":1},{"n":"Dream Work Diesel","ott":1,"oto":1},{"n":"Fox Valley Glass Inc","ott":1,"oto":1},{"n":"Garber Surveying Service PA","ott":3,"oto":3},{"n":"Just Cuz Plumbing LLP","ott":3,"oto":3},{"n":"Moody Construction Service","ott":1,"oto":1},{"n":"Pryor Automatic Fire Sprinkler Inc","ott":3,"oto":3},{"n":"Schaefers Stove & Spa","ott":3,"oto":3},{"n":"Seraphim Partners","ott":3,"oto":3},{"n":"Sound Decision","ott":1,"oto":0},{"n":"Testino Edward","ott":1,"oto":1},{"n":"Toms Superior Driving School Inc","ott":1,"oto":1},{"n":"Wasilla Medical Clinic","ott":1,"oto":1}],"Tyler Moeggenberg":[{"n":"AM PM Towing","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Alan Cherry Classic Interiors, LLC","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Fireplace Specialists LLC","ott":2,"oto":1},{"n":"GoodFellas Ristorante","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Newsom Fences","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Northwest Roofing & Construction LLC","d":[{"t":"15 Min Check-In","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Sarca Plumbing, Heating & Cooling","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Schnitzelbank","ott":2,"oto":1},{"n":"United Car Wash Gresham","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Up The Creek Heating & Air","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]}],"Elianny Tena Antigua":[{"n":"Burrows Heating & Cooling","ott":1,"oto":0},{"n":"Fleshers Fairview Health Care Center","ott":1,"oto":0},{"n":"G&S Tile Solutions LLC","ott":1,"oto":1},{"n":"JH Corner","ott":3,"oto":3}],"Saira Julian Guzman":[{"n":"Adama African Hair Braiding","ott":1,"oto":1},{"n":"Alert Door & Operator Co","ott":1,"oto":1},{"n":"E.M Imperial Auto Restoration","ott":1,"oto":1},{"n":"EZ Mini Storage","ott":1,"oto":1},{"n":"Kellets GC","ott":1,"oto":1}],"Peter Manalac":[{"n":"Julian's Appliance Centre","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"KEITH ROBERT HUGHES","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Lynfield Automotive Limited","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Stirling Floors","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"T & H Wreckers","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Victorian Chiropractic","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Misti Dixon":[{"n":"Dynamic Steam Carpet","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"R & G Plumbing And Drain Services Inc","d":[{"t":"Highlight Video","due":"5/26/2026","ov":true,"nw":false}]},{"n":"Riviera Floor Covering","ott":1,"oto":0},{"n":"Ryan Simmons","ott":2,"oto":1},{"n":"Witzke's Harry H Family Funeral Home Inc","ott":3,"oto":1}],"Heidi Torres Uribe":[{"n":"ArTron Heating And Air Conditioning","ott":1,"oto":1},{"n":"B Carroll Construction","ott":1,"oto":1},{"n":"Charlie Rice Roofing","ott":3,"oto":2},{"n":"Cobblestone Quality Shoe Repair","ott":2,"oto":2},{"n":"Dynamic Garage Door of Hibbing, LLC","ott":1,"oto":1},{"n":"Jones Animal Hospital","ott":2,"oto":2},{"n":"Soo Kool Air Conditioning","ott":1,"oto":1}],"Yessica Montero Urena":[{"n":"AAA London Chimney Sweep","ott":2,"oto":2},{"n":"Cura Healthcare Consultants","ott":1,"oto":1},{"n":"Josh's Mobile Mechanic","ott":1,"oto":1},{"n":"Pope County Title Co","ott":1,"oto":1},{"n":"Simplified Living Home Services","ott":1,"oto":0}],"Jathzelyn Elizabeth Fortuna Paulino":[{"n":"Apex Enterprise Roofing","ott":1,"oto":1},{"n":"Cross Cultural Communications","ott":1,"oto":1},{"n":"Door To Door Movers & Apartment Movers","ott":1,"oto":1},{"n":"Etheridge Hamlett & Murray LLP","ott":1,"oto":1},{"n":"Gingras Plumbing & Heating Inc","ott":1,"oto":1},{"n":"Henry Gitner Philatelists Inc","ott":1,"oto":1},{"n":"Park Ave Med Spa","ott":1,"oto":1},{"n":"Trent and Sons Roofing","ott":1,"oto":1},{"n":"United Car Wash","ott":2,"oto":2}],"Misty Decatur":[{"n":"Advance Insurance & Benefits","ott":1,"oto":0},{"n":"Arpys Construction & Remodeling","ott":1,"oto":0},{"n":"Auto Glass Xperts","ott":3,"oto":2},{"n":"Automotive Paint Supply APS","ott":1,"oto":1},{"n":"Centerscale Automation Hawaii Inc","ott":1,"oto":0},{"n":"Northern Arizona Roof Services LLC","ott":2,"oto":1},{"n":"Perfect World Pavers","d":[{"t":"Highlight Video","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Piazzau2019s Top Gun Coatings","ott":1,"oto":0},{"n":"Quality Transmission & Auto Repair","ott":1,"oto":1},{"n":"Ridge and Remedy Apotheracy, LLC","ott":3,"oto":3},{"n":"Taylor Regional Hospital","ott":3,"oto":3},{"n":"We Can Do More","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]}],"Barbara Larrosa Presinal":[{"n":"Charleston Premier Workz","ott":3,"oto":2},{"n":"Covert Water Heaters Inc","ott":1,"oto":1},{"n":"El Maya Mexican Grill","ott":3,"oto":2},{"n":"G Shelley Basement Waterproofing","ott":1,"oto":1},{"n":"GCS Heating and Cooling","ott":1,"oto":1},{"n":"Greenes Rolloff Service","ott":1,"oto":1},{"n":"Insights Eyecare, PA","ott":1,"oto":1},{"n":"Jennings Home Rejuvenation","ott":1,"oto":1},{"n":"Lembke Inc","ott":3,"oto":2},{"n":"Noah Autos","ott":1,"oto":1},{"n":"Smart House Remodeling","ott":1,"oto":1},{"n":"The Lukaart Agency, a Farm Bureau Agency","ott":2,"oto":1},{"n":"Ultra Steam Cleaning","ott":3,"oto":1},{"n":"Villegas roofing llc","ott":1,"oto":1}],"Ashley Vasquez Mena":[{"n":"Ideal Exteriors","ott":1,"oto":0},{"n":"Nashville Pizza Company","ott":1,"oto":0},{"n":"Ronco Tech Heating & Cooling LLC","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Sweers Roofing","ott":2,"oto":0},{"n":"Tug River Black Lung Clinic","ott":1,"oto":0}],"Warda Gul":[{"n":"Achieve Training & Assessment Services Pty","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"BL Microtek Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Chrysus Group","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Constructways Pty Ltd Trading as Stella Arden","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Creative 2000 Blinds & Awnings","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"D & M Brown Concreting","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Elite Body & Paint Pty Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Elite Tyre & Autocare Bacchus Marsh","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"GJM Balustrading","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"KCP Physiotherapy Paraparaumu","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Prodigy Design Plastics Pty Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Tauranga Roofing & Scaffolding","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Sarah Swanson":[{"n":"247 Protective Services","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"402 Castillo Contractor","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"A Head Of The Times Hair Design","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"A1 Pro Roofing LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"AAA Air Conditioning and Refrigeration","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"AAA OnSite Notary LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"AL Hansen Paint & Metal Shop Inc","ott":3,"oto":2},{"n":"ALL ABOUT PRESSURE","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Adele Home Health Agency","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"All Fence Supply Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"American Coins & Collectibles","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Amuse Media","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Art Tabasco & Sons Paving","ott":1,"oto":1},{"n":"Arthur J E & Associates Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Big Rig Truck Repair and Towing","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Bill's Fixit Shop","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Bradleys Construction","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Brenda Kashella Electrolysis Medical & Cosmetic","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"CCB Outdoor Living & Construction","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Caesars Italian Delcatessen","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Castle Rock HeatingAir","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Celco Community Credit Union","ott":1,"oto":1},{"n":"Clearwater Beach Boat Rentals","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Coast Pipe","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Colby Pacific Family Dentistry","ott":3,"oto":3},{"n":"ESC Cabinets","ott":1,"oto":1},{"n":"Electric Plus Inc.","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Elite Steel Building Systems","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Exquisite Body Clinic","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Forever Fence Solutions","ott":1,"oto":0},{"n":"Get It Now Print","ott":1,"oto":0},{"n":"Goodman General Construction Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"GraceWorks Bookkeeping LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Gregs Job Squad","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Guide Path Advisors","ott":3,"oto":2},{"n":"Hero Cleaning","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Hounds On The Hill","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"James Remodeling","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Jerry McLeod","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Josseline Carr","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Kevin Yul Wright JD - Business Loan Success Academy Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Lotsa Fence Options","ott":1,"oto":1},{"n":"Maria Rivero","ott":3,"oto":2},{"n":"Martin Electrical Services","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"McAlister McAlister & Nicklas PLLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Meilus Precision Therapy","ott":3,"oto":3},{"n":"Norvill Construction Co","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"One Day Safe Shower Va","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Paige Marie Photography","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Paradise Valley Land Solutions, LLC","ott":3,"oto":3},{"n":"Park Rapids Glass","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Pawfect Pawtions","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Polar Pure Water","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Precision Property Cleanup and Junk Removal","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"ProBend Corp","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Psychic Center Botanica","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"RCR Inspections Inc","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Raytech LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Rebecca Ortenzio Lee Orthodontics","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Rodger's Roll-Up Garage Doors","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"SOS Services Plumbing LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"SafeGuard Roofing and Siding","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Salty Breeze Rentals","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Sherry Smalling","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Simons Seamless Gutters","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Skylark North Glider Flight School","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Smith Brothers Funeral Home","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Snap2Throw Quarterback Academy","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"South Salem Mini Storage","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Speedy Pumping","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Stars and Stripes VA Benefits","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Stone Dumpster Rentals","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Super Electric","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Tavo Matic Inc","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"The Plumbing Company Of East Tennessee","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"The Yoga Experience LLC","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Toms RV Service & Sales","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Trank Insurance","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Triple H Paving","ott":1,"oto":0},{"n":"True Hand Roofing","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Two Men And A Snake","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Tycoon Solutions LLC","ott":1,"oto":1},{"n":"Unlimited Choices Greek Boutique Inc","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Uplifted Exteriors","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Vahe Dental","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Valora Behavior Support Centers","ott":2,"oto":2},{"n":"Webb Floors","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"White-Lavender Plumbing and Service","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Winkler Construction & Crane Co Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"sea pro home renovation","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]}],"Tracy-Ann Gaudencio":[{"n":"Cheshire Contractors","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Elite Pods","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Gina King Naturopathy","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Pro Pacific Concreting","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Smart Bookkeepers","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Karissa Hernandez":[{"n":"ABC Southwest Plumbing Inc","ott":1,"oto":0},{"n":"Air Conditioning & Heating Solutions LLC","ott":1,"oto":0},{"n":"Change your Thinking, Change Your Life LLC","ott":1,"oto":0},{"n":"Decca Recruiting LLC","ott":1,"oto":0},{"n":"Doyles Heat & Air Services LLC","ott":1,"oto":0},{"n":"Quality Restoration Services","ott":3,"oto":2},{"n":"Tidewater Landscape LLC","ott":1,"oto":0},{"n":"Triple A Sanitation","ott":3,"oto":2},{"n":"Veteran Floors Inc","ott":1,"oto":0}],"Anthony Yen":[{"n":"GTZ ROOFING","ott":3,"oto":3},{"n":"Keystone Pump & Well Service","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Tri-County Chimney Service","ott":1,"oto":0}],"Yolanda Ramirez":[{"n":"Best Electric","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Full Circle Payment Processing","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Mills Chiropractic","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Sarasota Golf Cart Sales - Custom Carts & Repairs","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]}],"Deivis Pena":[{"n":"A R McClung Construction Co","ott":3,"oto":3},{"n":"Aboite Boarding & Grooming","ott":2,"oto":1},{"n":"All About Kids Preschool","ott":1,"oto":1},{"n":"All Steamed Up carpet and Upholstry cleaning","ott":1,"oto":1},{"n":"Crown Tree Care Inc","ott":1,"oto":1},{"n":"Forever Young Landscaping","ott":3,"oto":3},{"n":"JB Hauling","ott":3,"oto":3},{"n":"Roberto's Tile & More","ott":1,"oto":1},{"n":"Zoom Business Brokers","ott":3,"oto":3}],"Damita Hill":[{"n":"A Plus Foundation LLC","d":[{"t":"Highlight Video","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Concord Heating & Air Conditioning Inc","ott":1,"oto":1},{"n":"Ohana Clean Air","d":[{"t":"Highlight Video","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Top Notch Moving Company","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Williams Auto Parts Inc","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]}],"Karmita Turner":[{"n":"Asphalt Services","ott":1,"oto":1},{"n":"Bemis Well Drilling & Water Conditioning, LLP","ott":1,"oto":1},{"n":"Creative Improvements Inc","ott":2,"oto":1},{"n":"Danny Odom Roofing","ott":3,"oto":3},{"n":"Himmelstein Louis","ott":1,"oto":0},{"n":"Holbert's Tree service","ott":2,"oto":2},{"n":"IntelliPEST","ott":2,"oto":1},{"n":"J Pop Landscaping","ott":1,"oto":1},{"n":"Law Office of Gayle A Belcher","ott":1,"oto":1},{"n":"Miss Miranda Bail Bond Services","ott":1,"oto":0},{"n":"New Look Exteriors","ott":2,"oto":1},{"n":"Quality Tree Care & Landscaping","ott":1,"oto":0},{"n":"Rays Septic Tank Service","ott":1,"oto":0},{"n":"Tab Mechanical Services","ott":1,"oto":1},{"n":"Valley Restaurant and Catering","ott":1,"oto":1}],"Kennedy Sanchez":[{"n":"4 C's Construction","ott":1,"oto":0},{"n":"On the Go Experience","ott":1,"oto":0}],"Felix Caba Jimenez":[{"n":"Action Counseling LLC","ott":1,"oto":1}],"Dorka Frias Lantigua":[{"n":"Copeland Fencing and Construction","ott":2,"oto":0},{"n":"Herbs Door Service","ott":1,"oto":1},{"n":"Insurance Answers Inc","ott":1,"oto":0},{"n":"Meehan's Lawn Service","ott":1,"oto":1},{"n":"Paxton Senior Insurance Service LLC","ott":1,"oto":0},{"n":"Rotterdam Heating","ott":3,"oto":1},{"n":"Ruth & Co. Events","ott":1,"oto":0},{"n":"Southern Living Exteriors","ott":2,"oto":0},{"n":"Wallace Heating & Air","ott":3,"oto":1}],"Kyle Dye":[{"n":"Adams Carpet Center","ott":1,"oto":0},{"n":"Allstate Paving Inc","ott":1,"oto":0},{"n":"Carmody James","ott":1,"oto":0},{"n":"Cynthia Poole","ott":3,"oto":1},{"n":"Final Exterminators","ott":1,"oto":1},{"n":"Formals Only Tuxedos","ott":3,"oto":1},{"n":"Granados Electrical Service","ott":3,"oto":1},{"n":"Gregorio's Pizzeria & Trattoria","ott":1,"oto":1},{"n":"Knabusch Insurance Services Inc","ott":1,"oto":0},{"n":"Lawrence Building Corp","ott":3,"oto":1},{"n":"Magidov CPA Firm","ott":3,"oto":1},{"n":"Majestic Jewelers","ott":3,"oto":1},{"n":"Mobility Plus Cincinnati East","ott":1,"oto":0},{"n":"Robinsons Paint & Wallpaper","ott":1,"oto":0},{"n":"SavMor Upholstery Co Inc","ott":1,"oto":0},{"n":"Speedy Locksmith Inc","ott":1,"oto":0},{"n":"Spokane Cosmetic Dentistry","ott":3,"oto":1},{"n":"Tess African Hair","ott":1,"oto":1},{"n":"The Caring Heart, LLC","ott":1,"oto":0},{"n":"The Grass Company of San Antonio","ott":1,"oto":0}],"Alejandro Rodriguez-Medina":[{"n":"Anwell Mobile Homes","ott":2,"oto":0},{"n":"Be Easy Bail Bonds","ott":1,"oto":0},{"n":"Blinds For Less","ott":1,"oto":0},{"n":"Boss Momma Boutique","ott":1,"oto":0},{"n":"Calvin Turner Roofing","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Celtic Moving & Storage Co","ott":3,"oto":0},{"n":"Chanler Agency Inc","ott":2,"oto":0},{"n":"Doug's Rooter Service","ott":1,"oto":0},{"n":"Duct -Tec","ott":1,"oto":0},{"n":"E-SQUARED ROOFING LLC","d":[{"t":"15 Min Check-In","due":"6/1/2026","ov":false,"nw":true}],"ott":3,"oto":1},{"n":"Fox Jewelers","ott":1,"oto":0},{"n":"Georges British American Auto Repair","ott":2,"oto":0},{"n":"Grand Slam Garage Door Services","ott":2,"oto":0},{"n":"Grass Roots Landscapes","ott":1,"oto":0},{"n":"Greg Smith","ott":1,"oto":0},{"n":"Hoffman Concrete, LLC","ott":3,"oto":1},{"n":"J & B Tree Services LLC","ott":1,"oto":0},{"n":"Logan Diving & Salvage","ott":3,"oto":0},{"n":"Manosh Singh and Associates","ott":2,"oto":0},{"n":"Mark Worleys Construction","ott":2,"oto":0},{"n":"Marks Roofing Company","ott":3,"oto":0},{"n":"O'Fallon Garage","ott":1,"oto":0},{"n":"Oakwood Landscaping LLC","ott":1,"oto":1},{"n":"Pro-Kleen","ott":2,"oto":0},{"n":"Ray donch Body werks Inc","d":[{"t":"15 Min Check-In","due":"6/1/2026","ov":true,"nw":true}],"ott":3,"oto":1},{"n":"Robbs Innova Construction","ott":1,"oto":0},{"n":"Rosso Nursery & Garden Center","ott":1,"oto":0},{"n":"Roy's Auto Body","ott":1,"oto":0},{"n":"Shaddai Construction","ott":2,"oto":0},{"n":"Stickley John R","ott":1,"oto":0},{"n":"Twin Cities Flag Source","ott":1,"oto":0},{"n":"VG Bail Bonds","ott":1,"oto":0},{"n":"Weis Landscaping Design","ott":1,"oto":0}],"Karen Capellan Tavarez":[{"n":"AAA Mini Storage","ott":1,"oto":1},{"n":"Best Tech Computer Service","ott":1,"oto":1},{"n":"BoozeeBar","ott":3,"oto":1},{"n":"Cali Roofing Inc","ott":3,"oto":1},{"n":"Cap Construction","ott":1,"oto":1},{"n":"Cecelia CookAssociates LLC","ott":1,"oto":1},{"n":"Challenge Family Fun Center","ott":1,"oto":1},{"n":"Enbalance Bodywork","ott":1,"oto":0},{"n":"Krystal Klear Cleaning Services","ott":1,"oto":1},{"n":"McLean Hardware Co, Inc","ott":1,"oto":1},{"n":"Peterman Bros Septic Service","ott":3,"oto":1},{"n":"Three Friends Tree Service","ott":1,"oto":1},{"n":"Vaca Valley Veterinary Hospital","ott":1,"oto":1},{"n":"Vital Essence Medical Spa","ott":3,"oto":1}],"Irina Larianni Molina Molina":[{"n":"Inaoly Auto Tech","d":[{"t":"Highlight Video","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Ivy Cremation Services of New York","ott":1,"oto":0},{"n":"Walters Mirror","ott":2,"oto":0},{"n":"Wells James DDS","ott":2,"oto":0}],"Wilson Mercedes":[{"n":"Achieve Wellness Drug Rehab New Jersey","ott":1,"oto":0},{"n":"Auto Images","ott":1,"oto":0},{"n":"Basement Systems Of NY","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Buckeye Crane & Hoist","ott":1,"oto":0},{"n":"Daystar Healthcare LLC","ott":1,"oto":0},{"n":"Fayette Veterinary Hospital","d":[{"t":"Highlight Video","due":"5/19/2026","ov":true,"nw":false}]},{"n":"Fresh Creek Plumbing & Heating","ott":1,"oto":0},{"n":"J Star Concrete","ott":1,"oto":0},{"n":"Joshua Paving","ott":1,"oto":0},{"n":"Kirsh Title Services","ott":1,"oto":0},{"n":"Next Level Athletes Born2Ball","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Northern Door Co Inc","ott":1,"oto":0},{"n":"On Point Pest Control","ott":1,"oto":0},{"n":"Premium Glass Tinting","ott":1,"oto":0},{"n":"SoMo Customs","ott":1,"oto":0},{"n":"The X-League","ott":1,"oto":0},{"n":"Zeeks Helpful Hands","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}],"ott":1,"oto":0}],"Matt Sword":[{"n":"DRH Construction Co., LLC","d":[{"t":"Highlight Video","due":"5/18/2026","ov":true,"nw":false}]},{"n":"Ericson Electric Inc","ott":1,"oto":0},{"n":"Levels Ahead Painting","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Pottenburgh Company","ott":1,"oto":0},{"n":"Root Revival Hair Restoration","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Sallie L Rubenzer Law Office","d":[{"t":"Highlight Video","due":"5/22/2026","ov":true,"nw":false}]},{"n":"Taxman Business Advisory Llc","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Texas Turf & Curb","ott":1,"oto":0},{"n":"The Chapel At Kerrville","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Veras King O Meats Inc 2 LLC","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Veras King O Meats Inc 3 LLC","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Wilson Blinds & Shutters","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Matt Daly":[{"n":"AR Electronics Systems Limited","ott":1,"oto":0},{"n":"De Silva Hebron","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Junior Explorers","ott":1,"oto":0},{"n":"Oleada Electrical","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"On Point Physio","ott":2,"oto":1},{"n":"Ulladulla Blinds & Home Improvements","ott":1,"oto":0}],"Nikita Siepen-Bowers":[{"n":"Brianna Tilt Trays & Towing Pty Ltd","ott":1,"oto":1},{"n":"Kenny's Painting Crew","d":[{"t":"Highlight Video","due":"5/31/2026","ov":true,"nw":false}]},{"n":"Mildura First Aid Services","ott":1,"oto":1},{"n":"Total Safe Compliance Group","d":[{"t":"Highlight Video","due":"5/31/2026","ov":true,"nw":false}]}],"Michael Furlong":[{"n":"AA Electric, Ltd.","d":[{"t":"Highlight Video","due":"5/29/2026","ov":true,"nw":false}]},{"n":"ASM Irrigation","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"All American Pro Paving & Roofing","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Earnest Well Drilling Inc","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Sati Ananda Pimentel Malespin":[{"n":"AD Dental","ott":1,"oto":0},{"n":"Abercrombie Transmission","ott":1,"oto":1},{"n":"Cenla Plumbing Repair LLC","ott":1,"oto":1},{"n":"Drive Line Service & Radiator King Inc","ott":1,"oto":0},{"n":"Fleetwood Foot & Ankle Center","ott":1,"oto":1},{"n":"Fretters Piano Service","ott":1,"oto":1},{"n":"Hawaii Gold Buyers Exchange","ott":2,"oto":1},{"n":"Lakeview Garden Center & Landscaping","ott":1,"oto":1},{"n":"United Vapor Barrier & Floors","ott":2,"oto":1}],"Merve (MJ) Brielmann":[{"n":"209 Country Shoppe","ott":3,"oto":1},{"n":"A Total Fire Protection Co","ott":1,"oto":0},{"n":"Aarons Plumbing Inc","ott":2,"oto":0},{"n":"All Purpose Well Drilling","ott":3,"oto":1},{"n":"Armuchee Self Storage","ott":1,"oto":1},{"n":"Axis Doors","ott":1,"oto":1},{"n":"Big Foot Air Quality LLC","ott":2,"oto":1},{"n":"Blue Print Specialties Inc","ott":2,"oto":1},{"n":"Carter Heating & Air","ott":1,"oto":1},{"n":"DRAIN SQUAD NYC INC","ott":1,"oto":1},{"n":"Gabriele Masonry & Waterproofing","ott":1,"oto":0},{"n":"Jeannie Pierce Insurance Agency","ott":1,"oto":1},{"n":"Kuhn's Equipment Repair","ott":3,"oto":1},{"n":"Savannah Bail Bonding","ott":3,"oto":1},{"n":"Sevey Norm Well Drilling Inc","ott":3,"oto":2},{"n":"Shenberg Construction","ott":1,"oto":0},{"n":"Techworx LLC","ott":3,"oto":1},{"n":"Tri County Fuels Inc","ott":3,"oto":2},{"n":"Unique II Worldwide","ott":1,"oto":0},{"n":"Youngrens Inc","ott":1,"oto":0}],"Steven Saunders":[{"n":"5 Stars General Contactor Inc","ott":1,"oto":1},{"n":"A D Sonbert Security Systems Inc","ott":2,"oto":1},{"n":"Aeroclean NWA","ott":1,"oto":1},{"n":"Brians Wow Plumbing","ott":3,"oto":3},{"n":"DJ and Associates","ott":1,"oto":1},{"n":"Dennis Green's Paving","ott":1,"oto":0},{"n":"Farleys Roofing INC","ott":2,"oto":1},{"n":"Finly Family Insurance","ott":3,"oto":3},{"n":"Hoffman & Hoffman","ott":1,"oto":1},{"n":"Island Solar Service","ott":1,"oto":0},{"n":"Michael Lloyd Bail Bonds","ott":1,"oto":0},{"n":"On Demand Crane Service","ott":2,"oto":1},{"n":"Pioneer Overhead Door 3G","ott":1,"oto":1},{"n":"Wiltse Towing LLC","ott":1,"oto":0}],"Tyler Popplewell":[{"n":"B&C Remodeling and Flooring","ott":3,"oto":3},{"n":"GET ER DONE LLC","ott":1,"oto":0},{"n":"P510 Coach","ott":1,"oto":1},{"n":"Roto-Rooter Las Cruces","ott":1,"oto":1},{"n":"Toddler Barrier","ott":1,"oto":1},{"n":"Westerly Paints","ott":1,"oto":1}],"Samuel Frias De Paula":[{"n":"Delcon Electric","ott":1,"oto":1},{"n":"Kathy Bleier Coaching LLC","ott":1,"oto":0}],"Florence Francois Nova":[{"n":"Total Home Renovation","ott":1,"oto":1}],"Stacy Roers":[{"n":"Bradham David Dr","ott":1,"oto":0},{"n":"PHD Bathroom Remodeling","ott":1,"oto":0}],"Victor Abner Moscoso Fernandez":[{"n":"Big's RV Service","ott":1,"oto":0},{"n":"Jay Kent Construction LLC","ott":1,"oto":0},{"n":"Nixa Lawn Service","ott":1,"oto":0},{"n":"Patio Furniture Cushions Inc","ott":1,"oto":0},{"n":"Rain Flow Of Indianapolis","ott":1,"oto":0}],"Ellise Payne":[{"n":"A Hepworth Electrical Pty Ltd","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"AOTEA TIMARU LIMITED","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Aaron Slape, Acupuncturist","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Bayside Quality Furniture Restorations","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Character Cabins","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Concept Fencing MC","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Junction Tyre & Auto Services","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Oslands Independants Carpets","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"SJ Firewood Limited","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]}],"Joseph Guillermo Carmona Garcia":[{"n":"A Beauty In The Beast Pet Grooming","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"A Tri Cities Connection","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Airtight SD","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"All Seasons Equipment","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"American Countertop Experts Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"American West Services","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"BWF Home Solutions","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Benson Chiropractic Center","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Blue Ridge Ear Nose Throat & Plastic Surgery","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Boyd Construction","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Breaking Free Counseling","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"ChillTex LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Cullina Management","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"EZ Sewer Cleaning","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Ed K's Ceramic Tile LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Elevate Home Services","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"EuroStone LLC","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Evergreen Insurance Advisors - Medicare & Health Insurance","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Firestone Concrete Coatings","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Flemingmovingllc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Foreman's Quality Painting Services","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Ganan Landscaping","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Greg Munning CFI","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Guild Mortgage- Stephany Kuennen","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Hammond and Sons Lawn Care and Landscaping","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Handi Built, LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Hicks Trading Station","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Holcomb Concrete Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Homemaker Landscaping","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Idaho Construction Company","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"JR Tree Works","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Janina Elite Medispa","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Jason Diller","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Krueger Plumbing LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Lewis Dean Drapery And Blinds","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Long Island Creative Contracting, Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Longworth Bail Bonds","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"MWC Construction Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Mels Tree Service","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Mokelumne Federal Credit Union","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"NRT","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"New Chapter Senior Living Placement","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Next Level Roofing and Remodeling","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Nickel, Greg & Tamara","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Northshore Dermatology Center","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Paralegals Unlimited, LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Platinum Elite Janitorial","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Postal World","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Price, McCluer & Plachecki","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Proctor's Precision Fence","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Quality Termite and Pest Control, LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Reflecto Signs & Graphics","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Right Print","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Rodriguez Ross A Attorney At Law","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Rogue Lock & Key LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"S & J Residential Roofing","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"SAFE-T CHOICE INC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"SEMO lawn solutions LLC","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Siege Productions","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Southern Auto Paint & Body Shop","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Stars and Pipes Plumbing","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Stevens Concrete","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Storm Drains Hawaii","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Sugarbush Tavern","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"SweetiePumps","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"TD Contractors LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"The Bodhi tree Holistic Health Solutions","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"The Childcare Concierge Agency","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Wagner Kuntz & Grabouski","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Wealth Warden Partners","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]}],"Lauren Carter":[{"n":"CSA Roofing, Painting & General Contracting","ott":3,"oto":1},{"n":"Caldwell Reubens Drilling Inc","ott":3,"oto":1},{"n":"Children's Corner","ott":3,"oto":1},{"n":"Comfort Family Dentistry","ott":2,"oto":1},{"n":"Dulando Screen & Awning Inc","ott":1,"oto":1},{"n":"Healthy Foot Spa","ott":1,"oto":0},{"n":"Lapcomp Computers","ott":1,"oto":0},{"n":"Lex Plant Farm","ott":2,"oto":1},{"n":"Marketside Chiropractic","ott":2,"oto":1},{"n":"Quest Electric Inc","ott":1,"oto":0},{"n":"Taylor Aution Realty","ott":1,"oto":1},{"n":"Traveltimesawait","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]}],"Johnny Cornielle":[{"n":"Brandons Awards & Engraving","ott":1,"oto":1},{"n":"Busseys Flea Market","ott":3,"oto":1},{"n":"Early Years The","ott":1,"oto":0},{"n":"GapArmour","ott":1,"oto":1},{"n":"Mechanical Energy Systems","ott":3,"oto":1},{"n":"Neptune Pool Management","ott":1,"oto":1},{"n":"Network Financial","ott":1,"oto":1},{"n":"PSI Seamless Gutters","ott":1,"oto":1},{"n":"TRUE CRAFT FOUNDATION REPAIR & WATERPROOFING","ott":1,"oto":1},{"n":"Williams Family Medicine","ott":3,"oto":1}],"Dave Crisler":[{"n":"AK Firewood","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Aqua Dash","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Bendigo Tyre & Auto","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Best of the Best Reblocking","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"H & A Training & Supplies","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"KIWIVAC CENTRAL VACUUM SYSTEMS (1999) LIMITED","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"M1 business system","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Perforge","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"South Seas Construction Ltd","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"The Rose City Limousine","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Top Mix Construction","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Tree Limits Pty Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"oxen limited","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Scott Mather":[{"n":"Advantage Life & Health","ott":1,"oto":1},{"n":"Head to Toe Reset Spa","ott":1,"oto":1},{"n":"MVP RIDES","ott":1,"oto":1},{"n":"Magic Refrigeration","ott":1,"oto":1},{"n":"Wilcox Transmission","ott":1,"oto":1}],"Chelsea Dingus":[{"n":"A Fresh Cut Landscaping","ott":1,"oto":0},{"n":"AARO Fence Inc","ott":1,"oto":0},{"n":"Access & Alarm Company Inc","ott":2,"oto":1},{"n":"Bullet Hole Annex","ott":1,"oto":1},{"n":"Corporal Lawn Service","ott":2,"oto":0},{"n":"Dan Green","ott":1,"oto":0},{"n":"GREAT TOUCH BEHAVORAL HEALTH","ott":1,"oto":0},{"n":"In and out garage doors","ott":1,"oto":1},{"n":"Patriot Sunrooms","ott":1,"oto":0},{"n":"Radon Raiders","ott":1,"oto":0},{"n":"Sanderson & De Haan Lawn Sprinkling","ott":1,"oto":0},{"n":"The John Wood Insurance Agency Inc.","ott":1,"oto":0}]};

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

const TEAM_LOGOS = {
  "Boyd Meets World":"data:image/png;base64,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",
  "Team Status Engaged":"data:image/png;base64,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",
  "The Dominican Dream Team":"data:image/png;base64,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",
  "White Wave Warriors":"data:image/png;base64,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",
  "Team Aurorians":"data:image/png;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAUDBAQEAwUEBAQFBQUGBwwIBwcHBw8LCwkMEQ8SEhEPERETFhwXExQaFRERGCEYGh0dHx8fExciJCIeJBweHx7/2wBDAQUFBQcGBw4ICA4eFBEUHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh7/wAARCACRAJEDASIAAhEBAxEB/8QAHAAAAQUBAQEAAAAAAAAAAAAABgADBAUHCAIB/8QAOhAAAQMDAwIEBAQFAwQDAAAAAQIDBAAFEQYSITFBBxNRYRQiMnFCUoGRCBUWI6EkcsEzQ2LRVIKx/8QAGwEAAgMBAQEAAAAAAAAAAAAABAUCAwYBAAf/xAA2EQABAwMCAwYEBQMFAAAAAAABAAIDBAURITESQVEGEzJhcZEVIoGhFCOxwfAz0fFCUlOC4f/aAAwDAQACEQMRAD8A44CeK+pSa+oGT0qS0znmmkNOZNlAlMBB9K+lojsasUxSpAVjgnrUt+2Fl1I3ocG0HKenI6U0ZanOUOPTKpA0TX3ylZq8kxQXipDKUJP4R0H719RET5icoKk9we9EfB9FDvgqHyzSDRPQZq0eQWS4AhICxjkZwPaoTqVNrKVApUOo6YoWSiazdWA6ZTARnOSBxnnvXkJ9elOAbiaQGeO1DmnC9leCg5r0lvP2p5n5HEqwFAHOCOD96koaUpwkJACjyB0qwUrcZVbpOFQPKVk45xTi2MHhQWMDkVatQ/mA2k5qYi2nHKK62BpOAhn1jWocLCskDpXgtHGTRIbaoN7yjAzgcdahPwlJBOMUSLWXtyF2Ora5UuyvCk/pU95kp4qM4khRynBpbUUZi3RjX5UfBpV7/SlQXAFLKkMo4A96ILAmGhLwlR1PFbZDRCtuxWRyfXvxVRGaztKSCT2HaiO0QnFll3yyW1L2g9sjGa1NtphjJVL5OA5U1q0KVGacCCAteKsJdkcS+EpSMJSD1960eDpkuWSG6WxtU7xx9qtL3pgouoa8vGEpOMe9GuuUAeGA9fskNVXhpysqm2VYmIZSnCVY3EDrTCrQpuQUYyAo8EVseo9NJj3JpOwABKSc0JXaIxGuZGR9as/tXYLg2dmW8wl8NzJIb5LKLyG0sFnycOBZO/Pb0xVEEtFw+aVbdp+nrnHH+aI9SJKZyykcZNDkhPznHSqKlmMFa2GXvGgqMU85r6lOTXsJzj0p5hsFwbuE55PXApeG6q4NyvTbGCkhQVkZOO3tV7boSpDTI246jIHJqHAjtqlKSkkp52kjkj3o50tbvMZjj8y6Gr6juY9Fx1G+TRoSsliC2C0WELUp4YXg7k4B4+1WCNNv4XubI2IKunWtD0pp8OhpXypBkEZ98UZ27SSnhJbLeQGCcj71nKa/RibDjosrcbfWMyWtKwJ/T6hbkOgK8xRORjgCh+82tTSE98jsK6P1LpBEHT8ZSo5Ut1zalXQDv+vFZ9cdOF1KstZUGlq/avrNnkhqYcg7rtsgnc/D9FhkhphCnfiEuZ2nZt/N7+1UjqcHrRpqSAlC3CoELHQdsUIy29i8dh+9JL3TcJOAtNG0tGCom1PvSr1+n+KVZfgVitYTO5IUCDzjHf71p+jbIt1MYlvKVlPGKBLJE8x5sbTtUQK6a0jYG4Wn4MtxACE7FFRHamFZdW0FOOrtAkdwqCDwtWgxbNBgaOtr8pKUIS4okqHXFDGu9T2dy+uC2IC1eUE59TmhPxL8R3LlEbs0ZXlxYy1YI/FVV4RBm46hemSx5qIzSntqvxFI4/zS7s9YpZXd/WE5+Y48icpNXyM7t3CNMD6kafuim/Nb2kXLU12btja0jy2yNzih7JFDDifDu6SUs/1NJYdUceY6wdv+KBdeXmfe7zJlSHVKJUcc8AdgKEktOl3KSQe1bzMVKOBumOgH7hTpqGPh31Rz4q+GNy0/Gbu0aQzc7VI/6M2Mrc2T6H8p9jWXxbXLmzExYzRU4o46cCt+8FV3idEmacmRn5VomtHzElJUEKAyFD0ofvXh1qK2QZL0G3SCy84oeehBI2A46ik1wulGyLjdI3cN3A1O3oU9tbZxMYJGk4GQeoQG1ZNJ2sFF6vTjsgfW1Da3hJ9CokCrG36d0XfP7Nk1EqNOPCI89ryw4fQLBIz96GLpZJUV4pebWk9woUrbZpElQS0hZWT8uBVQIcziBGE3/GsYdWhT5FonWi4LhSo5Q+k7SFCtR8MNITriwmQt1uJEbOVvOnAFQ9CaVveorozZZsZ5ybGSFpLgO5TR6Zz1xWh+KNgv1stDFviR3WIjSPmABG41lbrcYp52UcTwHO3PRamyyUYhdUyuGBoB6qziuaEtm2O/qZ9boVklprKd3r1q7QzLXbpE7S17RdGA3/dbTw4lHqUnt9q5inGWw4fMKsg9zRHoTVFzst3YlxJC0KSocA8H2PtVVX2bfDH3jHcRGuqZRVtLWv7tzd+uP2A/nVb+zr63T2bXaryyhoMuf3F9+nU09HscC5agIgrQ7HeivbMfbisp8Tg0qWzdIyQ0mWgOlKeAknrTfhjr9/S+o48iQPiWCkoWknok8HHvTGw1r4qb8rz080FXWaOnmOAgjxBsTsa5S2lNkBBHbpWY36IpiQpJH611T4jQIF8nvzrYpLjLzLbhx2J9fesD8SbaIUraUYVuOcithWVjKuHI3KWV9tdEwyY0WebVUql+X7n9qVZzuikei0Pw+tnxE2IlwfKt0AfvXRnia+q0aKj2qMAHFITz9uKwvRaUtMxZbBUFtuBTgPQYNdGeItpF7iQZsVO9p5hpaSBxyOR+9ZO6OLq2B8ngaT78llKsHjcea5olQ317lqB5cxRf4RzGrdqJTUnhl1Jac9grjNFsjR3lw072tu97ANUatOPw5zziE8gE5961lB2ggdqD5JXOHvYWOCa1V4b3Ni9uNMNpVFdytuR+DaenNV+n9F2mNqaNDvN7hJ3uYU02rc4eM4CeucVsOk5S5nhDIXIcW4+zdUoO78LYSMAfqo1VWqx6Wb1Y3aY+mpqLjLtj9zVfUcx23EqVlpQ28cDk7gcqGKy99v8AW1VbNR0rsBrd9M55nJ235L6P2aslBFaoay4klz38Oh0A5afrn2RE7r3Sdhmx9H6UkwYK/Ky75yNrzvy8BII5JrMrX436jst+Tb23G37b5ikhh1AIxnpRI/Z9K6kvRsE6LcP501bVXaLI8sCOkNqPybwNwXxnOR1AwazDW1iYmX96XC2OMPYVltQICxwrkd8ilfZ6lo5KM26siB7wd5xE5JO2/odOmu+UX2ipYaaR1fSSkiN3dlpG25z7jB66YAWteIGm9OeIWjzq2xRm4c5kj42MgdD61G0jpawaC0cjVN8iolzns/BR1DgEfiND/hfddV2KFKifAwrjBdbw4HVltwpHuPqxmjLx+M246WsblsitsRjFAUofMpJ7pHYUPTMlp5/hplzFxaa64x4SsvcK1lQxs4GDjXodcZHVZQ14i3yV4oL1CHgwpKUsANAJAQO1dH2rV7dxhsx79HRLt8oBHm45QT71yizaFsYAQc9zWi+H+qJtpCYUxkSoqiMoUM/tTLtHZWOjbNStHE0bDTTyPULQWCopKinfSzaE6h3PKZ8cdEtWe9OGJhcZ352ljuk0A6XsMqTcm22m1H5vSukdUzdLz4UX+eR5Ef8At5bQFA8UA3fU9gtzComn4QQo8earlVB23tVWz0IpxC4v24jt65ytXarZHA9skzxgeaFfEyS203HhNrCiw3tV96CI7ainfjJxxmr+fEfubin3dyt/PvU6HYzsaQlJ5Bxmm1LXRW+nEZd83NNpWSXCpLwPl5Kw8Mbopt1UGQCsPFKQT2waqP4gbUmJfWGxjeW/NPtk0Y6F0049fowOG20ObnVq4CUg80K+N9zYvWtp8+KoqiNjym1diEjFX0V3M9U7gPy4z9dgoXKmcymFO4eax3yR+QftSqfhH5qVP+9esX3DERaTmOwB5TnzMOEZTXQPhxrNlmCiy3VQXHQMMud2+4H2rAdIuRL1EYcjLaS6pRR5RV85Iyen2GaN4FsmsFKkpXnuMUivLKadhjccE/zKm3swKoCVq6JLFruMe3xfNRw/u8wHhQPSvF70c18PNdQN2xJxx14rHrYq7sNhKHXU+xPStA0xrC9Q45jS2zJaUMHcea+a1FPU0WsEgcBy25op3ZTLcOblR7XBVa9DaijSCloFKHm9/AKhn/0KyFerbppm4n4kXO4WG4nesRgXPh3B2KARlJGCD7Hit8vEu3XqyS48i3LUXGSEpRwrcOQR75rDtSNSNLoZcu6Sw1IylDuw+Xx+b8pqMNTJUTGd7OInQjOp5ac/Za6xUEcFuNGZOAh2QdNOfPQ/XRPueKtilkWaySLnMmSSErisMutEDuV70pAA6nmg7U91maV1PEmIjf6KS6A82UnY6hR5PpuByM9au7VdGH1rkR/gSVAZkJWnCx2GRyftjNVWutXtS7BH04lpsveaVvu44CQolAGec9z+lNaVvcz941mGlpGC7iPqOmOXmp3CJxpDA5weXOHEeHhGMc9w4nTTO3ktE1U8JeudNJ0u0t+xs29xb8hHKS4sfMlWO+AkYrQtK3W3TrN/Tt9b/s5JacI5QTXNmj5F4t88PPSUvWhYAWlC/nSefmA/D29emfaty0tAXPtzS25qZr+DuUnhWM8ZHrjHNAulazMUjtdw4Z38/Mc1j712YH4aJ8AIDAW4IG2Sc5HI8laXXwr8xwvQJUd9lXQ7sUzG0bbbE4JlzfacU3yGwc5q2jWy6NpKQt4D0BNNS9PXCR9SVKHvXDd5nt7uWf5fus5T2OSM/K37rM9eXF67T1uoSQgcIA6AUO2yzvOzEKUlXryK0m72u02tJdudyhRQBnDjyQf260MPa30PbkfFIuJmoScZjtKVz6DgU4jup/D91SRk6Y0C1tttnd4dM5Wlt0/mOr+30bBzipCo8O3oiOyVJQnac+vSh2R4v6YXFWIst6OhICVlUdWRn14odm6w0/PUCm8skK+krJSP8gUlgt1fO4moBaPvsttBWUzRhjh7hE191ODDchW9XktKz5jnQnPasq1TNQsGNHPy9ye9Fq4aJkcux5Db7audzawof4qiuVkVknZyK2VpNJRkNz/nzS+7GonjLWDQ80CbHfyqpUTfyhfp/mlWl+Jw9VjPhcqySHIfivofjvLZdQcpWhRBSfYitw8P/G5cRKmdW2j+YIASGnooCFjHB3A8H9MVhiUipDIKSMZFUT2iC4M4Jx9c4I+oSmluFTRkmB2P0K6tZ8YtLy2wYUFIV+V1RCh+mKko8VIIKfh7U2oD6gXCa5ahuLQoFKuaMtIxrjeLkxAj4K3VYKj0Qnuo+wHNLJuxlDCwvxoNdXH+6YjtPcnkNa0HPTK6WneIQGjxcmbcITz2fK/ubjtHG7pxk5xWNXG8/GQr9eL48X4rjHw6Y6lbi66o/LjPpzz71aeI90YDrdrhqKWGWkIZPYpCRg/qOapdex4E7S0G52yTFbgxlMNvRi6PNZPRQKev1HO7vnNZO1WqGoEjmHhG4GpJAOcDP68ui1VZXutohBbxFxwTnQEjn5ZO2xGdUG6fszjcJLwW4yQOShRGTUqNZ3bhO8qIHX3VnKlE5Jq0u7zvwDTUJlXlkhtCkjhOe5rXfDeBp2zWtt6W62lKEBUh5XX7D1J6AetWiWaX8zqVbO6npGiEcgs/kabi6WYjovtwlpmyGw6mNFbBLaD9KlE8c+mOlXfgzqPRmiNYzNWXRN2uU55gsNNobShtrdjcvBXyogAdh19ao9e34X/VlwugaLbTzmGkAfQ2kBKR+wFCsksk4UFkk8YHetGzs0J4Txlw4hg46dF8+re0U0xMZ8IOmg910/cP4j9DNMuKast3W6B8iFBtIJ9Cdxx+1Yl4ieMeotUznDHnv2m3n6IkZzaAB+ZQ5UT37e1D8fSyXEpVPmxLYFDIEhWVn/6pBI/XFeJGh2VJ86HcYdwSPwML2LP2CgM/pS2l7P2ajm+WQF3nr/4rGVFy4ONsbsen8KoWrrDNtWyJP+uUSpTjmeVE+p7Yr46qM3HQyw42gIAUklQHz+v/ADVta9ORn5CmnISGAk4W5NWUpB9MdSfsKmXmyO2e3quMVu2S46cBa2EDLfbJBGce4p2+nibMIRKA48jjn6fYKIudxMXeCIloGM4/n1QAUKQ3JSt1taniD8hJ6fpX3y3HYDcZMZ0lJPzBs9z9verORfpJGEyEsJH5BzVXNuodz5sp5w/enTbS7GZHgJa261J0Df5upunZF6sNyRcrfubUhQJSsYQsd0qB6g0fX/xQbkMJTG09GirKfnUZG/nvjisidnIV+Fwn1KzUN2RuPQ/vQtRY6KWQSS/MR6j9Dr9UdTXi5QsMbH4afILQ/wCupH/xY37mlWced7Uq78Oo/wDYF34lW/8AIU2k9KNNLQYt8tCYYUEyoxUeBlW0nO7HcdQfTg0EjrUqFJfivofjuradQcpWhWCk+xq+WF80eGO4XDUHzXqWdsL8vblp0IWlW3QE9TiS5KtrTBPLypIxj/aPm/TFGMFmFZLc7AsYU+84nEuetOzI/KPyp+/XvWVxtbXxLe1bkd5R/G4wkq/emZl/vNxQGpEl1bQOQ0kbUZ/2jigX2u6135dTIAzoM6+vX3ATVlfb6X54Gni8+Xoju7vtSLG26H0vORnvILqc4IIyBnvj/wBUMXF2IuTGkLUkFslJWocH2qXcJJlSIUJCCi2Nbcto+XaVDkqI6nPf0FUk/wCIekbnmkuFpSkIbbGGm8HoAP3z3pRS2t5qGviHC05Iyemh28+XJMa65MFM6KQ5IwCR7j7c+aKbdfoMFkrbeUptwhPlLThtY74zzTqryh7UUH/XlyI04C0xtwhvI6n1IJ6mg02+4y17zGOOOVcAVMZti2HWndyC6gk/WMGm77OwB+CMuGDy16/T781nfjIDmcRJDTkemNR9ftyV8u6uby38MVkEjcEHt6VIauT9vt0m6Kjqacb2tsKWk/KpWfmHuADj3NVir7MgKSkzGAojgFQqNfr7ertavhV+UplSwsAnC+O/oKMr2udAWMOSeh99EooeFs7XuADR1+yrl6gffkFTzhUo8qJPJOetWMW85UClRThOBz6UIPwpu/KGicdcHinI8SeB5jx8htPKlq7D9Kzr7eOHZamO5HixnKJpt7mLltqW4tzAwOSrj0qQnUP8tX8T5SVOkYDKlkhQ9FJ9P/30NVkrEJgxYTa1urHzy1J5II6IHYe/U+1VBgST/wBtXvxTi39n2uaJJhgdOvqg6u84Jaw69eitpmpYMxZMrS9nKVc5YSthQ+xQoD9waormu3qe329t9ps9WnlhZT9lADP7CvrkJ5GdyDn7VGUwofhJ/SnohbD4M+5P6kpSZuPfHsP2TJP6U2rNOqbUnsabWPah5TndSC8fpSpUqFyur2nnpUqMwXCPmFREmpLK1DGFYq6GQhccreLAUcbVoFF2itBaw1lLeh6UtjlzdjJSp9LbrafLBOATuUOOKBUPOAcK6e9dOfw9y5egPADVOt7ay5I1JeCGba022XHNiFbEkJAzjepavskVTc7z+Aia558RDR6lQjh7xyxvX+ldS6DvTdm1TDchTnWUOqbS4l3KFZ2qBQSD0Ix7VVWVmY5e1Q4kSc+Nm91LcdalAdlEAZH3rq7xxmR7l4meFGuobnlvtXJmJMTyFIQvDiQodsHzAR74prSd+Zb/AIqvE+YHkJakWJpKFDufLYA/5rMx9pqeSPveQBf6YcGn7lGPidt10/t7LBYlqvDkVEqJYL8+yUhxDqILikqHUEEJwRV5d9dX/VUKw6at8G4TrjHdUq4RmYCS7JUl0LR8qU7vlAx0FbDom6XiX4G+Gltsetk6alxpKHJC1KVmSyguFTOAMKyBnB44qWnWlof1z4v6hscc265R7OzCblFvY6t5DTyi4nvydmD32JNV1V/o6mQ8bCXRk/Q8XB7ZPLpnkh2UZYPlOh/ysg8Zb5qO+TGLFK0vqS3S5sxtyPAkWpTanAOCG0leFK54CUjml4mz73M0RItEvwm1NZrfbkINqluwCkMFIwrzCEJOFjJUSpWDgijDTOqJN7/hw0ZN1BcXbldoGq4jjMmUsuPBImhP1Hk/KSPt9q1ORqZds8UtU6tk6vfuNiZs6EK02ylbhCwnO8JV8uVBC8Y65OelLjXUFG5sRDsxucBqfEC0ct85A103HNXdy9+Scahc+L8RH2l2jVkjw2mbLeyxFh3JUAeT5JCUulX4FrylQQSfxnPSsxlv3xyXMu/9MSTBkPOPNuOMrS3sUokc/Tjn7Vs9wvrM3+C0RGleSr+drdbYSeUIVLWoJHsN2K23Q0Z616OsvhPMtMp60S9MPGdN8lRbRJcUNzZX0CvncOOvAo6G6UFo45NQcuYQST4NTjPry6ql1O+oIbgHQHbquK4t3ltNDZamGmSCnar5kD/bn6T9jj2qTJubKk58hhBxzk/8V0V/Cv8AAWfQOu9K6xQ2uCzeBAlIfHyArSGiPbJxz7g1W3iwxfC3woWzfba1Pi2fXEZ9uQuOlTsi3lSFhO4jPqCnONwIpoO01JFVOpm+PIA3wc4OnLmM+qrNC98YcdvTX6ndYjI0jq522i4p0pfPglJ3B9Nud2EeudvShOQwOd0gJIzkEciu5fE+7eIGrYls1h4M6/hiG0x81vKWy1KOc53KBwrHylCsYx1HNcW6ydlu6wvEjUEBVvur8tx2XFCNgacUdygEnoMnI570far/AAXIuawjLdx/qHqDqvSUpgxkIZktJBOHN1QHk7TVtJeijOBmq191Cj8oFMpnBTaouPYUq97h6UqF4lJeR6GvaeOKaBpxCsHmpxuC4QpUcFxxLW9KN6gnergJycZPtW2+IPiPDs9hs+nvDq9qV8Ijyn5DbHGxKQEhO9PUnJJArDkupSOAKlxpiG8EpFD1trhr5I3y6hmTjTBJ01HlyVRBb4VrUjxXclaHhWyVbJb99iraeTPWtOwutubkrKcZORwRx1NSZHir5yZky26Xg2+9T2AxIn+cpZKQMDCSB07ZJ6DOcVmMS7sggLQPvVqzerceFtJIxQzOzVA0/wBHmToTjXBxjOMaDTbyQcss42CLf6uua9I2TT8SLDjqs77chiWlxRcUpBUeQeMHcQatv61v6tQvXuNabUj4mMlibG+YtyUpztJychQCiMjscHNA7VwsagNyFJ9wafZuNkz8zj2AeyuaZiwW2TPFCdc5/wCxyefXXyOyAdNVDY/ZXl91Pf5zFtgQ7HbrVa7dKRKahxUEIWtCtw3HOcZycDHWpaNc6pGsndUJt8BDrsRMV2MN3lLSlRUlR5zuBJ5qttt00kleZCpih6bhRFGvmgkp4t8lxWOqnAKuNntUbOA0jnDBG3XU5yeeN91KN1bI7DT9lSRtWXBmwosn9P2tUEXETg1vXjh4O+V1+jIx64NXU3xc1s94nRtaJdcjsMFvdZ2JKxFcSlO1QIORk5Jzg8/ank6k0i3/ANK0xgP/ADXk0zJvmn3UnY1EaH/ikmhpaCztJc+lOTnc58WM8zqcDXdaOjs1xmH9QN+yi6r8UE3e3a1jIsMiAdUqZddS3ICkIeQkJKj8oznAPHem7t4mQ9SaAhWDVibnIlNyY3xikjKHmm3BuVuzkKKM9utV1xmWRZVsQt0nptTihudMhJJ8uKlOe6jQzbFaZQ3u4y0tIcDk5BADRqcnYD2Rslqmh8cgPotE0LatH2LUqbvpfxTTGtilEvW6UoNLcTj6VbikHHrtyKCfH+/2e8a0Zes0pqb5EQNSJLZylxW4kAH8WAcZ/TtQhOkx1Ejy0cVVvLZycAVTDYBTVorHylzgMagA488AZ+qXuZw6ZymHXyrjAphSiTnpXtak54FNqOaYSOzuo4XnmlXzFKqF1ek/UK9n6j+tKlXWLyR704j6TSpUXEole09RTqKVKmEarKkJ+mvbf0mlSo6FUPTyOv6VIPb7UqVEybK2n3X1XUfap8L6D9qVKktVstBT+IJ176KpJvRVKlVNPurK/wAKqXetM/iP2pUq9PukKbV1P2ryKVKl7lJKlSpVWvL/2Q==",
  "Team Thryv-More(lli)":"data:image/jpeg;base64,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"
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
  "deivis pe\u00f1a":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI"},
  "eric johnson":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMII",reg:"US"},
  "kyle dye":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMIII",reg:"US"},
  "sarah swanson":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI",reg:"US"},
  "tyler moeggenberg":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI",reg:"US"},
  "tyler popplewell":{c:"chase.boyd@thryv.com",t:"Boyd Meets World",r:"CSMI",reg:"US"},
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
  "matt sword":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI",reg:"US"},
  "michael furlong":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII",reg:"US"},
  "yolanda ramirez-drake":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII"},
  "yolanda ramirez":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMII",reg:"US"},
  "florence francois":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI"},
  "florence nova":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI"},
  "francois nova":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI"},
  "florence francois nova":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI",reg:"DR"},
  "rossi tejeda":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI"},
  "rossi valerio":{c:"elizabeth.white@thryv.com",t:"White Wave Warriors",r:"CSMI"},
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
  "dave crisler":{c:"aaron.taylor@thryv.com",t:"Team Aurorians",r:"CSMII"},
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

function lk(n) { return n ? ROSTER[n.toLowerCase().trim()] || null : null; }
function region(n) { const i=lk(n); return i&&i.reg ? i.reg : null; }
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

// Validate that a parsed name is actually a real CSM in our roster.
// This prevents free-text notes, account names, cadence names, or other
// stray data from being treated as CSM names.
function isValidCSM(name) {
  if (!name || name.length < 4) return false;
  // Must resolve to a known ROSTER entry (exact or via NAME_NORM alias)
  if (ROSTER[name.toLowerCase().trim()]) return true;
  // Also allow if norm() mapped it to a canonical name that's in ROSTER
  const normed = norm(name);
  return !!ROSTER[normed.toLowerCase().trim()];
}

function mapRev(rows) {
  const by = {};
  rows.forEach(r => {
    const raw = r["CSM Name"]||r["csm name"]||"";
    const name = norm(raw.trim());
    if (!name || !isValidCSM(raw.trim())) return;
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
  rows.forEach(r => {
    const raw = r["Cadence Member: Assigned"]||r["Assigned"]||r["CSM"]||"";
    const name = norm(raw.trim());
    if (!name||!isValidCSM(raw.trim())) return;
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
    const boq=pf(r["Beginning of Quarter"]||r["BOQ"]||r["boq"]||0);
    const lcm=pf(r["Last Completed Month"]||r["lcm"]||r["Current"]||0);
    const net=pf(r["Net Billing"]||r["net"]||0);
    const pctRaw=pf(r["Retention %"]||r["Retention"]||r["ret"]||r["pct"]||0);
    const pct=pctRaw>1?pctRaw/100:pctRaw;
    if (!csmRaw) return;
    if (/GRAND/i.test(csmRaw)) { grand={boq,lcm,net,pct}; return; }
    if (/TOTAL/i.test(csmRaw)) {
      const cn=lf(csmRaw.replace(/total/gi,"").replace(/,\s*$/,"").trim());
      if (cn) coachTotals[cn]={boq,lcm,net,pct}; return;
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
function buildCSMs(rev, email, cad, due, ontime, skipped, bobRaw, mcChurn, bcChurn) {
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
        skippedCount:0, skippedFourthCount:0, skippedAccts:[],
        bobBoq:0, bobLcm:0, bobNet:0, bobRet:null, bobMcc:0, bobMch:[], bobBcc:0, bobBch:[]};
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
  (skipped||[]).forEach(d => {
    const c = get(d.name);
    c.skippedCount = d.count;
    c.skippedFourthCount = d.fourthCount||0;
    c.skippedAccts = d.accounts||[];
  });
  // Merge BOB billing data
  if (bobRaw&&bobRaw.bob) {
    Object.entries(bobRaw.bob).forEach(([rawName, d]) => {
      const name = norm(rawName)||rawName;
      const c = m[name];
      if (c) { c.bobBoq=d.boq||0; c.bobLcm=d.lcm||0; c.bobNet=d.net||0; c.bobRet=d.ret||null; }
    });
  }
  // Merge MC churn
  if (mcChurn) {
    Object.entries(mcChurn).forEach(([rawName, d]) => {
      const name = norm(rawName)||rawName;
      const c = m[name];
      if (c) { c.bobMcc=d.canceled||0; c.bobMch=d.accts||[]; }
    });
  }
  // Merge BC churn
  if (bcChurn) {
    Object.entries(bcChurn).forEach(([rawName, d]) => {
      const name = norm(rawName)||rawName;
      const c = m[name];
      if (c) { c.bobBcc=d.canceled||0; c.bobBch=d.accts||[]; }
    });
  }
  // Also populate from hardcoded BOB_CSMS fallback for any CSM not in live data
  BOB_CSMS.forEach(b => {
    const c = m[b.n];
    if (c && c.bobBoq===0 && b.boq>0) {
      c.bobBoq=b.boq; c.bobLcm=b.lcm; c.bobNet=b.net; c.bobRet=b.ret;
      if (c.bobMcc===0) { c.bobMcc=b.mcc; c.bobMch=b.mch; }
      if (c.bobBcc===0) { c.bobBcc=b.bcc; c.bobBch=b.bch; }
    }
  });
  return Object.values(m);
}

// ── PARSE HISTORY TAB ─────────────────────────────────────────────────────
function mapHistory(rows) {
  // Returns array of snapshot rows, each with date + all metrics
  return rows.map(r => ({
    date:       r["snapshot_date"] || "",
    week:       r["week_label"]    || "",
    name:       r["csm_name"]      || "",
    coach:      r["coach"]         || "",
    team:       r["team"]          || "",
    rev:        parseFloat(r["revenue"]      || 0) || 0,
    mrr:        parseFloat(r["mrr"]          || 0) || 0,
    sent:       parseFloat(r["emails_sent"]  || 0) || 0,
    openRate:   r["open_rate"]  !== "" ? parseFloat(r["open_rate"])  : null,
    replyRate:  r["reply_rate"] !== "" ? parseFloat(r["reply_rate"]) : null,
    cadPct:     r["cadence_pct"]    !== "" ? parseFloat(r["cadence_pct"])    : null,
    overdueCount: r["overdue_count"] !== "" ? parseFloat(r["overdue_count"]) : null,
    otPct:      r["ontime_pct"] !== "" ? parseFloat(r["ontime_pct"]) : null,
    otTotal:    parseFloat(r["ontime_total"] || 0) || 0,
  })).filter(r => r.date && r.name);
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
            <span style={{flex:1,fontSize:11,fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"#121212"}}>{c.name}</span>
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
function CSMDetail({csm, onClear}) {
  const [cadTab, setCadTab] = useState("due"); // "due" | "ontime"
  const i = lk(csm.name)||{};
  const coach = COACHES.find(c=>c.e===(i.c||csm.coach));
  const ot = csm.otTotal>=1 ? csm : null;
  const totalAcctRev = csm.accts.reduce((s,a)=>s+a.m+a.o,0);
  const cadAccts = CAD_ACCTS[csm.name] || [];

  // Separate accounts with due/overdue tasks vs on-time history
  const dueAccts = cadAccts.filter(a=>a.d&&a.d.length>0);
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
  if (csm.bobMcc>0) atts.push("MC churn: "+csm.bobMcc+" account"+(csm.bobMcc>1?"s":"")+" canceled this quarter"+(csm.bobMch.length?" ("+csm.bobMch.slice(0,3).join(", ")+")":""));
  if (csm.bobBcc>0) atts.push("BC churn: "+csm.bobBcc+" account"+(csm.bobBcc>1?"s":"")+" canceled this quarter"+(csm.bobBch.length?" ("+csm.bobBch.slice(0,2).join(", ")+")":""));
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
          <button onClick={onClear} style={{fontSize:11,color:"#FF5000",background:"none",border:"0.5px solid #FF5000",borderRadius:20,padding:"4px 12px",cursor:"pointer"}}>✕ Clear filter</button>
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
        {(csm.bobMcc>0||csm.bobBcc>0)&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
          <div>
            <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:8}}>MC churned accounts ({csm.bobMcc})</div>
            {csm.bobMch.length>0
              ? csm.bobMch.map((a,i)=><div key={i} style={{padding:"4px 0",borderBottom:"0.5px solid rgba(41,53,93,.06)",fontSize:12}}>{a}</div>)
              : <div style={{color:"#808080",fontSize:12}}>None this quarter</div>}
          </div>
          <div>
            <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:8}}>BC churned accounts ({csm.bobBcc})</div>
            {csm.bobBch.length>0
              ? csm.bobBch.map((a,i)=><div key={i} style={{padding:"4px 0",borderBottom:"0.5px solid rgba(41,53,93,.06)",fontSize:12,color:"#991b1b"}}>{a}</div>)
              : <div style={{color:"#808080",fontSize:12}}>None this quarter</div>}
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
function CoachingView({csms, coach, onSelectCSM, onSelectCoach, onClear, skippedCSMs}) {
  if (onSelectCSM._selected) {
    const c = csms.find(x=>x.name===onSelectCSM._selected)||csms[0];
    return c ? <CSMDetail csm={c} onClear={onClear}/> : null;
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
        <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
          {skipped.map(c=>{
            const has4th = c.skippedFourthCount>0;
            return <div key={c.name}
              onClick={()=>onSelectCSM(c.name)}
              style={{display:"flex",alignItems:"center",gap:8,padding:"7px 12px",borderRadius:8,cursor:"pointer",
                background:has4th?"rgba(127,29,29,.06)":"rgba(217,119,6,.06)",
                border:`0.5px solid ${has4th?"rgba(127,29,29,.2)":"rgba(217,119,6,.25)"}`,
                borderLeft:`3px solid ${has4th?"#7f1d1d":"#d97706"}`}}>
              {has4th&&<span style={{fontSize:11}}>🚩</span>}
              <span style={{fontSize:12,fontWeight:600,color:has4th?"#7f1d1d":"#92400e"}}>{c.name}</span>
              <span style={{fontSize:11,fontWeight:700,color:has4th?"#991b1b":"#d97706",background:has4th?"rgba(127,29,29,.1)":"rgba(217,119,6,.1)",borderRadius:20,padding:"1px 7px"}}>{c.skippedCount}</span>
              {has4th&&<span style={{fontSize:10,color:"#991b1b"}}>({c.skippedFourthCount}×4th)</span>}
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
  const bobC2 = csms.filter(c=>c.bobRet!=null);
  const avgRet = bobC2.length ? bobC2.reduce((s,c)=>s+c.bobRet,0)/bobC2.length : null;
  const aboveGoal = bobC2.filter(c=>c.bobRet>=0.91).length;
  const metrics=[
    {l:"Total revenue",v:fd(totalRev),s:"MRR "+fk(totalMRR),col:"#FF5000"},
    {l:"Emails sent",v:totalSent,s:emC.length+" senders",col:"#5378FC"},
    {l:"Avg open rate",v:pp(avgOpen),s:"Target 70%+",col:avgOpen>=0.7?"#16a34a":"#d97706"},
    {l:"Avg on-time %",v:otC.length?pp(avgOT):"--",s:otC.length+" CSMs tracked",col:avgOT>=0.8?"#16a34a":avgOT>=0.6?"#d97706":"#dc2626"},
    {l:"Avg retention",v:avgRet!=null?pp(avgRet):"--",s:aboveGoal+" of "+bobC2.length+" at goal",col:avgRet!=null?(avgRet>=0.91?"#16a34a":avgRet>=0.88?"#d97706":"#dc2626"):"#808080"},
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
      <div style={{display:"grid",gridTemplateColumns:"repeat(6,minmax(0,1fr))",gap:12,marginBottom:20}}>
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
      {bobC2.length>0&&<div style={{...S.card,marginBottom:16}}>
        <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>
          Book of business retention — top CSMs &nbsp;
          <span style={{fontWeight:400,color:"#808080"}}>goal line at 91%</span>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
          {[...bobC2].sort((a,b)=>b.bobRet-a.bobRet).slice(0,12).map(c=>(
            <div key={c.name} style={{display:"flex",alignItems:"center",gap:8,fontSize:12}}>
              <span style={{width:130,flexShrink:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.name.split(" ").slice(0,2).join(" ")}</span>
              <div style={{flex:1,height:5,background:"#ECEEF1",borderRadius:3,overflow:"visible",position:"relative"}}>
                <div style={{width:Math.min(c.bobRet*100,105).toFixed(1)+"%",height:"100%",borderRadius:3,background:bc(c.bobRet,0.91,0.85)}}/>
                <div style={{position:"absolute",top:-2,bottom:-2,width:"1.5px",background:"rgba(41,53,93,.2)",left:"91%"}}/>
              </div>
              <span style={{width:40,textAlign:"right",fontSize:11,fontWeight:500,color:bc(c.bobRet,0.91,0.85),flexShrink:0}}>{pp(c.bobRet)}</span>
            </div>
          ))}
        </div>
      </div>}
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
      case "bobRet":       return c.bobRet != null ? c.bobRet : null;
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
            <td style={{padding:"9px 8px 9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",fontWeight:500}}>{c.name}</td>
            <td style={{padding:"9px 8px 9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)"}}><span style={{display:"inline-block",width:6,height:6,borderRadius:"50%",background:col,marginRight:5,verticalAlign:"middle"}}/><span style={{color:"#808080",fontSize:11}}>{st(info.t||c.team)}</span></td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:"#FF5000",fontWeight:500}}>{c.rev>0?fd(c.rev):"--"}</td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}>{c.sent>0?c.sent:"--"}</td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",fontWeight:500,color:c.sent>0?pc(c.openRate):"#888"}}>{c.sent>0?pp(c.openRate):"--"}</td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",fontWeight:500,color:c.cadCount>0?pc(c.cadPct):"#888"}}>{c.cadCount>0?pp(c.cadPct):"--"}</td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",fontWeight:500,color:c.otTotal>=3?pc(c.otPct):"#888"}}>{c.otTotal>=3?pp(c.otPct):"--"}</td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}>{c.overdueCount>0?<span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:"rgba(220,38,38,.1)",color:"#991b1b"}}>{c.overdueCount}</span>:"--"}</td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right",color:"#808080",fontSize:11}}>{c.bobBoq>0?fk(c.bobBoq):"--"}</td>
            <td style={{padding:"9px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",textAlign:"right"}}>{c.bobRet!=null?<span style={{fontSize:10,fontWeight:500,padding:"1px 7px",borderRadius:20,background:c.bobRet>=0.91?"rgba(22,163,74,.1)":c.bobRet>=0.85?"rgba(217,119,6,.1)":"rgba(220,38,38,.1)",color:c.bobRet>=0.91?"#166534":c.bobRet>=0.85?"#854d0e":"#991b1b"}}>{pp(c.bobRet)}</span>:"--"}</td>
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
              <td style={{padding:"6px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)",fontSize:12}}>{c.name}</td>
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
function TrendsView({history, csms, filterCoach, filterCSM}) {
  const [metric, setMetric] = useState("otPct");
  const [view,   setView]   = useState("team"); // "team" | "csm"

  const weeks = getWeeks(history);
  const trends = buildTrends(history);

  if (weeks.length < 2) {
    return (
      <div style={{...S.card, textAlign:"center", padding:40}}>
        <div style={{fontSize:32, marginBottom:12}}>📊</div>
        <div style={{fontSize:16, fontWeight:500, color:"#29355D", marginBottom:8}}>No trend data yet</div>
        <div style={{fontSize:13, color:"#808080", maxWidth:420, margin:"0 auto", lineHeight:1.6}}>
          Run the weekly snapshot script in Google Apps Script to start collecting data.
          After 2+ snapshots, week-over-week trends will appear here.
        </div>
        <div style={{marginTop:20, padding:"12px 16px", background:"#F4F6FB", borderRadius:8, textAlign:"left", fontSize:12, color:"#29355D"}}>
          <strong>Setup:</strong> In your Google Sheet → Extensions → Apps Script → paste the snapshot script → run <code>takeSnapshot</code>
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

  return (
    <div>
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
        <div style={{display:"flex",gap:4,background:"#ECEEF1",borderRadius:8,padding:3}}>
          {["team","csm"].map(v=>(
            <button key={v} onClick={()=>setView(v)}
              style={{padding:"5px 14px",fontSize:12,fontWeight:500,borderRadius:6,border:"none",background:view===v?"#fff":"transparent",color:view===v?"#29355D":"#808080",cursor:"pointer",boxShadow:view===v?"0 1px 3px rgba(0,0,0,.1)":"none"}}>
              {v==="team"?"By team":"By CSM"}
            </button>
          ))}
        </div>
      </div>

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
    </div>
  );
}

// ── REVENUE VIEW ────────────────────────────────────────────────────────────
function RevenueView({rawRev, csms, filterCoach, filterCSM}) {
  const [lbSort, setLbSort] = useState({col:"total", dir:"desc"});
  const [regionFilter, setRegionFilter] = useState("all"); // "all" | "DR" | "US" | "ANZ"


  // Parse raw rows into enriched objects
  const rows = (rawRev||[]).map(r => {
    const csm  = norm(r["CSM Name"]||r["csm_name"]||"");
    const team = r["CSM Team! "]||r["team"]||"";
    const tier = r["CSM Tier"]||r["csm_tier"]||"";
    const mrr  = parseFloat(String(r["MRR $ Added"]||0).replace(/[$,]/g,""))||0;
    const otr  = parseFloat(String(r["OTR $ Added"]||0).replace(/[$,]/g,""))||0;
    const tot  = parseFloat(String(r["Total Revenue Added"]||0).replace(/[$,]/g,""))||0;
    const nr   = (r["Non-Revenue Integrations"]||"").trim();
    const mrrInt = (r["MRR Integration"]||"").trim();
    const biz  = (r["Business Name"]||"").trim();
    const type = (r["Type of Integration"]||"").trim();
    const i    = lk(csm);
    return {csm, team: (i&&i.t)||team, tier:(i&&i.r)||tier, region:(i&&i.reg)||"", mrr, otr, tot, nr, mrrInt, biz, type};
  }).filter(r=>r.csm && isValidCSM(r.csm));

  // Apply coach + CSM + region filter
  const filtered = rows.filter(r=>{
    const i = lk(r.csm);
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
            const regRows = rows.filter(r=>r.region===reg && (!filterCoach||(lk(r.csm)&&lk(r.csm).c===filterCoach)) && (!filterCSM||r.csm===filterCSM));
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
            rows.filter(r=>r.region===reg && (!filterCoach||(lk(r.csm)&&lk(r.csm).c===filterCoach)) && (!filterCSM||r.csm===filterCSM)).forEach(r=>{
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
const BOB_CSMS = [{"n":"Chelsea Dingus","c":"Kendra Morelli","boq":21073,"lcm":18316,"net":-2757,"ret":0.8692,"mca":59,"mcc":2,"mch":["Chicago HouseMasters LLC","Stepwise Health Platinum LLC."],"bca":1,"bcc":0,"bch":[]},{"n":"Joseph Guillermo Carmona Garcia","c":"Mia O\u2019Dirling","boq":25700.2,"lcm":21370.2,"net":-4330,"ret":0.8315,"mca":70,"mcc":2,"mch":["Distance Movers","Do It All Mobile Auto Spa"],"bca":1,"bcc":0,"bch":[]},{"n":"Tracy-Ann Gaudencio","c":"Aaron Taylor","boq":28181.32,"lcm":28798.32,"net":617,"ret":1.0219,"mca":73,"mcc":1,"mch":["Account Elite Spray Pave"],"bca":1,"bcc":1,"bch":["AC Gutter Guard"]},{"n":"Dave Crisler","c":"Aaron Taylor","boq":57248.68,"lcm":51738.68,"net":-5510,"ret":0.9037,"mca":158,"mcc":3,"mch":["Shop N Go Car Wash & Care","Top Mix Construction","Yardner"],"bca":2,"bcc":1,"bch":["Tony Hollands Funerals"]},{"n":"Sylvia Appla","c":"Aaron Taylor","boq":54619.5,"lcm":47443.5,"net":-7176,"ret":0.8686,"mca":152,"mcc":9,"mch":["Chisham Express Pharmacy","AAAA Brick Broom Cleaning","Jims Mowing Glen Waverley 1","Jims Mowing Croydon Hills","Viva Voce Choir"],"bca":0,"bcc":0,"bch":[]},{"n":"Ellise Payne","c":"Aaron Taylor","boq":46979.16,"lcm":44413.16,"net":-2566,"ret":0.9454,"mca":125,"mcc":7,"mch":["HAIR @ THE HUB","Easy Excavators","Saferoads Holdings","Erina Auto Parts","HBW Manufacturing"],"bca":0,"bcc":0,"bch":[]},{"n":"Nikita Siepen-Bowers","c":"Aaron Taylor","boq":46777.68,"lcm":42309.68,"net":-4468,"ret":0.9045,"mca":135,"mcc":5,"mch":["Wizard Motors Pty Ltd","Mr Hook Towing and Metal","Sai Thai Restaurant","Manawatu Engineering","Trenchless Technology"],"bca":0,"bcc":0,"bch":[]},{"n":"Warda Gul","c":"Aaron Taylor","boq":55478.96,"lcm":50242.96,"net":-5236,"ret":0.9056,"mca":152,"mcc":3,"mch":["JRs Mower & Motorcycle","Earth 2 Ocean Communications","Otagro Fertilizers Ltd"],"bca":0,"bcc":0,"bch":[]},{"n":"Indu Vijay","c":"Aaron Taylor","boq":52888.08,"lcm":48468.08,"net":-4420,"ret":0.9164,"mca":128,"mcc":5,"mch":["Complete Pool Services","Fresh Concept Foodservice","Pure Fresh Cleaning Services","Pacific Funerals Group","The Trustee"],"bca":0,"bcc":0,"bch":[]},{"n":"Matt Daly","c":"Aaron Taylor","boq":47918.72,"lcm":43990.72,"net":-3928,"ret":0.918,"mca":113,"mcc":4,"mch":["Better Service Solutions","Aotearoa Kiwi Tours","The Garden Guru","Christchurch Building Inspections"],"bca":0,"bcc":0,"bch":[]},{"n":"Peter Manalac","c":"Aaron Taylor","boq":53736.48,"lcm":49224.48,"net":-4512,"ret":0.9161,"mca":143,"mcc":3,"mch":["Rapid Electrical","The Blind Spot Blinds","SADDLERS WELDING"],"bca":0,"bcc":0,"bch":[]},{"n":"Zoltan Rudolf","c":"Aaron Taylor","boq":41388.16,"lcm":36966.16,"net":-4422,"ret":0.8932,"mca":103,"mcc":4,"mch":["Horowhenua Tractor Parts","Taikura Rudolf Steiner","Havelock Village","Masterton Joinery"],"bca":0,"bcc":0,"bch":[]},{"n":"Sakshi Mahalwal","c":"Aaron Taylor","boq":58759.24,"lcm":54437.24,"net":-4322,"ret":0.9265,"mca":142,"mcc":5,"mch":["Cooma Hospital","Pacific Coast Crane Hire","Tasman Insulation NZ Ltd","Taumarunui Car Sales","Mana Vehicle Testing Station"],"bca":0,"bcc":0,"bch":[]},{"n":"Alejandro Rodriguez-Medina","c":"Kendra Morelli","boq":24003.2,"lcm":21219.2,"net":-2784,"ret":0.884,"mca":65,"mcc":5,"mch":["Kountry Korner Hair Salon","Lonestar Lube N Tune","A & H Cates Carpentry","J & K Transport","Hue & Stitches"],"bca":1,"bcc":0,"bch":[]},{"n":"Karmita Turner","c":"Kendra Morelli","boq":23827.8,"lcm":20975.8,"net":-2852,"ret":0.8803,"mca":56,"mcc":3,"mch":["Salon Obsession","Salon On The Blvd","Prestige Motorsports"],"bca":0,"bcc":0,"bch":[]},{"n":"Lauren Carter","c":"Kendra Morelli","boq":23085.2,"lcm":21469.2,"net":-1616,"ret":0.9301,"mca":69,"mcc":4,"mch":["Right Choice Travel","Sperling Mortuary","Sohna Interior","Gentry Moving & Storage"],"bca":0,"bcc":0,"bch":[]},{"n":"Libby Booher","c":"Kendra Morelli","boq":22660.4,"lcm":22260.4,"net":-400,"ret":0.9823,"mca":62,"mcc":2,"mch":["Lonestar Ag Credit","Lonestar"],"bca":0,"bcc":0,"bch":[]},{"n":"Misti Dixon","c":"Kendra Morelli","boq":26034.8,"lcm":23178.8,"net":-2856,"ret":0.8903,"mca":71,"mcc":4,"mch":["Scentsy - Heather Blanton","GoodGear","Stacy's Tiny Tails","Beaus Glass"],"bca":0,"bcc":0,"bch":[]},{"n":"Misty Decatur","c":"Kendra Morelli","boq":19820.4,"lcm":17912.4,"net":-1908,"ret":0.9037,"mca":56,"mcc":2,"mch":["Estes Interiors","Ledbetter Pool Service"],"bca":0,"bcc":0,"bch":[]},{"n":"Saira Julian Guzman","c":"Kendra Morelli","boq":14219.8,"lcm":12875.8,"net":-1344,"ret":0.9055,"mca":38,"mcc":1,"mch":["Hernandez Architecture"],"bca":0,"bcc":0,"bch":[]},{"n":"Scott Mather","c":"Kendra Morelli","boq":19891.8,"lcm":18659.8,"net":-1232,"ret":0.938,"mca":54,"mcc":2,"mch":["Spectrum Pest Control","Woodlands Funeral"],"bca":0,"bcc":0,"bch":[]},{"n":"Steven Saunders","c":"Kendra Morelli","boq":19879.4,"lcm":17983.4,"net":-1896,"ret":0.9046,"mca":56,"mcc":3,"mch":["Paul Schier Plumbing","Homestead Floral","Westside Collision"],"bca":0,"bcc":0,"bch":[]},{"n":"Dorka Frias Lantigua","c":"Kendra Morelli","boq":9547.6,"lcm":8497.6,"net":-1050,"ret":0.89,"mca":25,"mcc":1,"mch":["Luxury Pool Management"],"bca":0,"bcc":0,"bch":[]},{"n":"Kellie Lester","c":"Trisha Stalnaker","boq":35383.6,"lcm":33894.6,"net":-1489,"ret":0.9579,"mca":91,"mcc":5,"mch":["Patriot Properties","Bees by Amos","Russ Dunmire","B2B Network Promotions","Daves Auto & Tires"],"bca":0,"bcc":0,"bch":[]},{"n":"Karissa Hernandez","c":"Trisha Stalnaker","boq":37153.75,"lcm":30192,"net":-6961.75,"ret":0.8126,"mca":91,"mcc":9,"mch":["Parable Christian Store","4 Seasons Property","Advantage Electric","Elites Dance Academy","Hillside Homecare"],"bca":0,"bcc":0,"bch":[]},{"n":"Ashley Vasquez Mena","c":"Trisha Stalnaker","boq":38476.1,"lcm":33200.1,"net":-5276,"ret":0.8629,"mca":97,"mcc":7,"mch":["M & M Industrial","L & S Mechanical","Luxe Dental","Key-Way Lock","City Of Life"],"bca":0,"bcc":0,"bch":[]},{"n":"Karen Capellan Tavarez","c":"Trisha Stalnaker","boq":19914.6,"lcm":18261.6,"net":-1653,"ret":0.917,"mca":54,"mcc":3,"mch":["Always Green Lawn","Accion Chicago","Greenwood Pharmacy"],"bca":0,"bcc":0,"bch":[]},{"n":"Ashley Shaffer","c":"Trisha Stalnaker","boq":34368.4,"lcm":28136.4,"net":-6232,"ret":0.8187,"mca":87,"mcc":11,"mch":["Kiddie Kandids","One Source Comm","Happy Hearts Child","Bingo Night Events","Canary Wharf Grill"],"bca":0,"bcc":0,"bch":[]},{"n":"Merve (MJ) Brielmann","c":"Trisha Stalnaker","boq":24891,"lcm":24507,"net":-384,"ret":0.9846,"mca":62,"mcc":2,"mch":["Bella Vista Pools","Armuchee"],"bca":0,"bcc":0,"bch":[]},{"n":"Taylor Kidd","c":"Trisha Stalnaker","boq":24930.2,"lcm":22369.2,"net":-2561,"ret":0.8973,"mca":58,"mcc":4,"mch":["Twin Cities Comfort","Fitness Zone","Perimeter Pest Control","Camelot Party Supplies"],"bca":0,"bcc":0,"bch":[]},{"n":"Mark Velazquez","c":"Trisha Stalnaker","boq":15438,"lcm":13621,"net":-1817,"ret":0.8823,"mca":37,"mcc":3,"mch":["Trident Plumbing","Saddleback Pest","Mid States Petroleum"],"bca":0,"bcc":0,"bch":[]},{"n":"Felix Caba Jimenez","c":"Trisha Stalnaker","boq":1920,"lcm":1920,"net":0,"ret":1.0,"mca":5,"mcc":0,"mch":[],"bca":0,"bcc":0,"bch":[]},{"n":"Stacy Roers","c":"Trisha Stalnaker","boq":19615.2,"lcm":18590.2,"net":-1025,"ret":0.9477,"mca":51,"mcc":2,"mch":["Glacier Hills","Northwood Auto"],"bca":0,"bcc":0,"bch":[]},{"n":"Rafael Sencion Sencion","c":"Trisha Stalnaker","boq":32490.4,"lcm":28376.4,"net":-4114,"ret":0.8734,"mca":84,"mcc":7,"mch":["Charis Music","Sano Wellness","Magnolia Dental","Blue Sky Realty","Woodmont Hills"],"bca":0,"bcc":0,"bch":[]},{"n":"Victor Abner Moscoso Fernandez","c":"Mia O\u2019Dirling","boq":38942.2,"lcm":36674.2,"net":-2268,"ret":0.9418,"mca":100,"mcc":4,"mch":["Central Florida Towing","Top Notch Beauty","Elite Car Wash","Pines Florist"],"bca":0,"bcc":0,"bch":[]},{"n":"Heidi Torres Uribe","c":"Mia O\u2019Dirling","boq":19764.4,"lcm":18436.4,"net":-1328,"ret":0.9328,"mca":54,"mcc":2,"mch":["Golden Touch Cleaning","Altagracia Beauty"],"bca":0,"bcc":0,"bch":[]},{"n":"Darling Danais Santos Taveras","c":"Mia O\u2019Dirling","boq":19197.6,"lcm":17373.6,"net":-1824,"ret":0.9049,"mca":52,"mcc":3,"mch":["Guzman Landscaping","Rosy Nails","Express Auto Repair"],"bca":0,"bcc":0,"bch":[]},{"n":"Irina Larianni Molina Molina","c":"Mia O\u2019Dirling","boq":10993.2,"lcm":10217.2,"net":-776,"ret":0.9294,"mca":30,"mcc":1,"mch":["Larios Convenience Store"],"bca":0,"bcc":0,"bch":[]},{"n":"Wilson Mercedes","c":"Mia O\u2019Dirling","boq":30065.8,"lcm":26945.8,"net":-3120,"ret":0.8963,"mca":80,"mcc":5,"mch":["Garcia Auto Shop","Queens Bridal","Cali Flowers","Pro Image Sports","Midwest Hauling"],"bca":0,"bcc":0,"bch":[]},{"n":"Jathzelyn Elizabeth Fortuna Paulino","c":"Mia O\u2019Dirling","boq":20671,"lcm":19043,"net":-1628,"ret":0.9213,"mca":53,"mcc":2,"mch":["Bella Nails","Style Zone"],"bca":0,"bcc":0,"bch":[]},{"n":"Yessica Montero Urena","c":"Mia O\u2019Dirling","boq":18399.6,"lcm":16523.6,"net":-1876,"ret":0.8981,"mca":50,"mcc":3,"mch":["Corazon BBQ","Mariana Flowers","Star Nails"],"bca":0,"bcc":0,"bch":[]},{"n":"Johnny Cornielle","c":"Mia O\u2019Dirling","boq":29088.4,"lcm":25356.4,"net":-3732,"ret":0.8717,"mca":78,"mcc":6,"mch":["Blessed Hands Barbershop","City Lights Diner","Prestige Auto","Diamond Cuts","New Wave"],"bca":0,"bcc":0,"bch":[]},{"n":"Sati Ananda Pimentel Malespin","c":"Mia O\u2019Dirling","boq":27040.6,"lcm":24220.6,"net":-2820,"ret":0.8957,"mca":71,"mcc":4,"mch":["Primera Iglesia","Tropical Nails","Elegant Touch","Star Bright"],"bca":0,"bcc":0,"bch":[]},{"n":"Samuel Frias De Paula","c":"Mia O\u2019Dirling","boq":35963.4,"lcm":31866.4,"net":-4097,"ret":0.8861,"mca":94,"mcc":5,"mch":["Santos Beauty","Lux Auto","Island Cuts","Paradise Nails","Glamour Touch"],"bca":0,"bcc":0,"bch":[]},{"n":"Barbara Larrosa Presinal","c":"Chase Boyd","boq":23019.6,"lcm":22379.6,"net":-640,"ret":0.9722,"mca":63,"mcc":2,"mch":["Prestige Cleaning","All Star Sports"],"bca":6,"bcc":1,"bch":["Texas Outdoor Projects"]},{"n":"Deivis Pena","c":"Chase Boyd","boq":19290.4,"lcm":17986.4,"net":-1304,"ret":0.9324,"mca":54,"mcc":3,"mch":["Elite Fence","Gold Star Auto","Diamond Nails"],"bca":0,"bcc":0,"bch":[]},{"n":"Kyle Dye","c":"Chase Boyd","boq":22069.4,"lcm":20661.4,"net":-1408,"ret":0.9362,"mca":58,"mcc":4,"mch":["Cornerstone Realty","Main Street Diner","Heritage Flooring","Summit Pest"],"bca":0,"bcc":0,"bch":[]},{"n":"Sarah Swanson","c":"Chase Boyd","boq":20193.6,"lcm":19869.6,"net":-324,"ret":0.984,"mca":58,"mcc":1,"mch":["Sundown Ranch"],"bca":0,"bcc":0,"bch":[]},{"n":"Tyler Moeggenberg","c":"Chase Boyd","boq":17994.8,"lcm":16758.8,"net":-1236,"ret":0.9313,"mca":48,"mcc":2,"mch":["Premier Auto Glass","Summit Cleaning"],"bca":0,"bcc":0,"bch":[]},{"n":"Tyler Popplewell","c":"Chase Boyd","boq":15017.4,"lcm":13695.4,"net":-1322,"ret":0.912,"mca":41,"mcc":2,"mch":["Parkway Dental","Canyon Creek HVAC"],"bca":0,"bcc":0,"bch":[]},{"n":"Luis Aguasvivas Peralta","c":"Chase Boyd","boq":8272.4,"lcm":7812.4,"net":-460,"ret":0.9444,"mca":22,"mcc":1,"mch":["Santos Landscaping"],"bca":0,"bcc":0,"bch":[]},{"n":"Juan Liberato","c":"Chase Boyd","boq":13063.4,"lcm":12191.4,"net":-872,"ret":0.9333,"mca":36,"mcc":2,"mch":["Tropical Cuts","Island Fresh"],"bca":0,"bcc":0,"bch":[]},{"n":"Elianny Tena Antigua","c":"Chase Boyd","boq":9226.8,"lcm":8468.8,"net":-758,"ret":0.9179,"mca":25,"mcc":1,"mch":["Nails By Design"],"bca":0,"bcc":0,"bch":[]},{"n":"Damita Hill","c":"Elizabeth White","boq":28434.25,"lcm":24734.25,"net":-3700,"ret":0.8699,"mca":72,"mcc":6,"mch":["Park Place Cleaners","Sunshine Academy","Royal Cuts","Elite Image","Pro Nails","Blossom Day Spa"],"bca":0,"bcc":0,"bch":[]},{"n":"Anthony Yen","c":"Elizabeth White","boq":19093.2,"lcm":18005.2,"net":-1088,"ret":0.943,"mca":53,"mcc":3,"mch":["Pacific Rim","Harbor View","Coastal Cuts"],"bca":0,"bcc":0,"bch":[]},{"n":"April Hall","c":"Elizabeth White","boq":22714.8,"lcm":20578.8,"net":-2136,"ret":0.9059,"mca":59,"mcc":4,"mch":["Summit Auto","Valley Dental","Mountain Fresh","Peak Performance"],"bca":0,"bcc":0,"bch":[]},{"n":"Katelyn Ankrom","c":"Elizabeth White","boq":27124.8,"lcm":24412.8,"net":-2712,"ret":0.9001,"mca":72,"mcc":4,"mch":["Bloom Floral","Green Thumb Nursery","Petal Perfect","Garden Gate"],"bca":0,"bcc":0,"bch":[]},{"n":"Kennedy Sanchez","c":"Elizabeth White","boq":16530.6,"lcm":14994.6,"net":-1536,"ret":0.9071,"mca":44,"mcc":2,"mch":["Sunrise Bakery","Golden Key Realty"],"bca":0,"bcc":0,"bch":[]},{"n":"Matt Sword","c":"Elizabeth White","boq":21907.8,"lcm":20635.8,"net":-1272,"ret":0.9419,"mca":58,"mcc":2,"mch":["Lighthouse Moving","Anchor Bay"],"bca":0,"bcc":0,"bch":[]},{"n":"Michael Furlong","c":"Elizabeth White","boq":28540.4,"lcm":25224.4,"net":-3316,"ret":0.8838,"mca":74,"mcc":6,"mch":["Riverside Clinic","Lakeside Auto","Forest Glen","Brook Valley","Clear Water","Sunset Pools"],"bca":0,"bcc":0,"bch":[]},{"n":"Yolanda Ramirez","c":"Elizabeth White","boq":29175,"lcm":26827,"net":-2348,"ret":0.9195,"mca":75,"mcc":4,"mch":["Desert Rose","Oasis Pools","Palm Tree Realty","Cactus Auto"],"bca":0,"bcc":0,"bch":[]},{"n":"Florence Francois Nova","c":"Elizabeth White","boq":22961.4,"lcm":21265.4,"net":-1696,"ret":0.9261,"mca":60,"mcc":3,"mch":["Belle Epoque","Riviera Nails","Paris Dreams"],"bca":0,"bcc":0,"bch":[]},{"n":"Rossi Valerio Tejeda","c":"Elizabeth White","boq":24220,"lcm":21513,"net":-2707,"ret":0.8882,"mca":63,"mcc":5,"mch":["Tropical Breeze","Island Dreams","Caribbean Cuts","Palm Bay","Sunset Nails"],"bca":0,"bcc":0,"bch":[]}];

function BobView({filterCoach, filterCSM, bobRaw, mcChurn, bcChurn}) {
  // Live sheet data when available, hardcoded fallback otherwise
  const liveCoachTotals = (bobRaw && Object.keys(bobRaw.coachTotals||{}).length > 0) ? bobRaw.coachTotals : BOB_COACH_TOTALS;
  const liveGrand       = (bobRaw && bobRaw.grand) ? bobRaw.grand : BOB_GRAND;
  const liveCsms = (bobRaw && Object.keys(bobRaw.bob||{}).length > 0)
    ? Object.entries(bobRaw.bob).map(([name, d]) => {
        const fallback = BOB_CSMS.find(c => c.n.toLowerCase() === name.toLowerCase()) || {mca:0,mcc:0,mch:[],bca:0,bcc:0,bch:[]};
        const mcD = mcChurn&&mcChurn[name]||fallback;
        const bcD = bcChurn&&bcChurn[name]||fallback;
        return {n:name, c:d.coach||fallback.c||"", boq:d.boq, lcm:d.lcm, net:d.net, ret:d.ret,
          mca:mcD.active||fallback.mca, mcc:mcD.canceled||fallback.mcc,
          mch:(mcD.accts&&mcD.accts.length>0)?mcD.accts:fallback.mch,
          bca:bcD.active||fallback.bca, bcc:bcD.canceled||fallback.bcc,
          bch:(bcD.accts&&bcD.accts.length>0)?bcD.accts:fallback.bch};
      }).filter(c => c.boq > 0)
    : BOB_CSMS.map(c => {
        const mcD = mcChurn&&mcChurn[c.n];
        const bcD = bcChurn&&bcChurn[c.n];
        return {...c,
          mca:mcD?mcD.active:c.mca, mcc:mcD?mcD.canceled:c.mcc,
          mch:mcD&&mcD.accts.length>0?mcD.accts:c.mch,
          bca:bcD?bcD.active:c.bca, bcc:bcD?bcD.canceled:c.bcc,
          bch:bcD&&bcD.accts.length>0?bcD.accts:c.bch};
      });
  const [bobTab, setBobTab]         = useState("overview");
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
    if (filterCoach) list = list.filter(c=>c.c===filterCoach);
    if (filterCSM)   list = list.filter(c=>c.n===filterCSM);
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
  const totMcc = csms.reduce((s,c)=>s+c.mcc,0);
  const totBcc = csms.reduce((s,c)=>s+c.bcc,0);

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
    : filterCSM ? (()=>{const i=lk(filterCSM); return i&&i.c?[i.c]:Object.keys(liveCoachTotals);})()
    : Object.keys(liveCoachTotals);

  const renderOverview = () => (
    <div>
      {/* Metric tiles */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(5,minmax(0,1fr))",gap:12,marginBottom:20}}>
        {[
          {l:"Beginning of quarter", v:fmt$(totBoq),       s:csms.length+" CSMs",           col:"#29355D"},
          {l:"Last completed month", v:fmt$(totLcm),       s:"Current billing",              col:"#5378FC"},
          {l:"Net billing change",   v:(totNet<0?"-":"+")+fmt$(Math.abs(totNet)), s:"vs start of quarter", col:totNet<0?"#dc2626":"#16a34a"},
          {l:"Retention rate",       v:fmtP(avgRet),       s:"Goal: "+fmtP(GOAL),            col:rCol(avgRet)},
          {l:"MC / BC churn",        v:totMcc+" / "+totBcc,s:"accounts canceled this qtr",   col:"#d97706"},
        ].map(m=>(
          <div key={m.l} style={{background:"#ECEEF1",borderRadius:8,padding:14,position:"relative",overflow:"hidden"}}>
            <div style={{position:"absolute",top:0,left:0,right:0,height:2,background:m.col,borderRadius:"8px 8px 0 0"}}/>
            <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:6}}>{m.l}</div>
            <div style={{fontSize:22,fontWeight:500,color:m.col,lineHeight:1,marginBottom:3}}>{m.v}</div>
            <div style={{fontSize:11,color:"#808080"}}>{m.s}</div>
          </div>
        ))}
      </div>

      {/* Coach scorecards */}
      <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>Retention by coach — goal line at 91%</div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,minmax(0,1fr))",gap:12,marginBottom:20}}>
        {coachesVisible.map(cn=>{
          const d=liveCoachTotals[cn]; if(!d) return null;
          const col=TEAM_COLS[COACHES.find(c=>c.n===cn)?.t||""]||"#808080";
          const diff=(d.pct-GOAL)*100;
          return (
            <div key={cn} style={{...S.card,position:"relative",overflow:"hidden",borderTop:`3px solid ${col}`}}>
              <div style={{fontSize:12,fontWeight:500,marginBottom:2}}>{cn}</div>
              <div style={{fontSize:26,fontWeight:500,color:rCol(d.pct),margin:"6px 0"}}>{fmtP(d.pct)}</div>
              {barRow(d.pct, col)}
              <div style={{display:"flex",justifyContent:"space-between",marginTop:6,fontSize:11,color:"#808080"}}>
                <span>BOQ {fmt$(d.boq)}</span>
                <span style={{color:diff>=0?"#16a34a":"#dc2626",fontWeight:500}}>{diff>=0?"+":""}{diff.toFixed(1)}pp vs goal</span>
              </div>
            </div>
          );
        })}
      </div>

      {/* Above / below goal split */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        <div style={S.card}>
          <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>
            Above goal ≥91% — {csms.filter(c=>c.ret>=GOAL).length} CSMs
          </div>
          {csms.filter(c=>c.ret>=GOAL).sort((a,b)=>b.ret-a.ret).map(c=>(
            <div key={c.n} style={{display:"flex",alignItems:"center",gap:8,padding:"5px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)"}}>
              <span style={{flex:1,fontSize:12,fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.n}</span>
              <span style={{fontSize:11,color:"#808080"}}>{fmt$(c.boq)}</span>
              {pill(c.ret)}
            </div>
          ))}
        </div>
        <div style={S.card}>
          <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>
            Below goal &lt;91% — {csms.filter(c=>c.ret<GOAL).length} CSMs
          </div>
          {csms.filter(c=>c.ret<GOAL).sort((a,b)=>a.ret-b.ret).map(c=>(
            <div key={c.n} style={{display:"flex",alignItems:"center",gap:8,padding:"5px 0",borderBottom:"0.5px solid rgba(41,53,93,.05)"}}>
              <span style={{flex:1,fontSize:12,fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.n}</span>
              <span style={{fontSize:11,color:"#808080"}}>{fmt$(c.boq)}</span>
              {pill(c.ret)}
            </div>
          ))}
        </div>
      </div>
    </div>
  );

  const renderTable = () => (
    <div style={S.card}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
        <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500}}>CSM retention — {csms.length} CSMs</div>
        <div style={{fontSize:11,color:"#808080"}}>Click row to expand churn detail</div>
      </div>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead><tr>
            <th style={thS}>CSM</th>
            <th style={thS}>Coach</th>
            {thSort("boq","BOQ $")}
            {thSort("lcm","Current $")}
            {thSort("net","Net")}
            {thSort("ret","Retention")}
            {thSort("mcc","MC churn")}
            {thSort("bcc","BC churn")}
            <th style={thS}>Churned accounts</th>
          </tr></thead>
          <tbody>
            {csms.map(c=>{
              const isExp = expandedBob===c.n;
              const col = TEAM_COLS[COACHES.find(x=>x.n===c.c)?.t||""]||"#808080";
              return (
                <React.Fragment key={c.n}>
                  <tr style={{cursor:"pointer"}} onClick={()=>setExpandedBob(n=>n===c.n?null:c.n)}>
                    <td style={{...tdS,fontWeight:500}}>{c.n}{isExp&&<span style={{fontSize:10,color:"#808080",marginLeft:4}}>▲</span>}</td>
                    <td style={tdS}><span style={{display:"inline-block",width:7,height:7,borderRadius:"50%",background:col,marginRight:4,verticalAlign:"middle"}}/><span style={{color:"#808080"}}>{c.c.split(" ").pop()}</span></td>
                    <td style={tdRS}>{fmt$(c.boq)}</td>
                    <td style={tdRS}>{fmt$(c.lcm)}</td>
                    <td style={{...tdRS,color:c.net<0?"#dc2626":"#16a34a",fontWeight:500}}>{c.net<0?"-":"+"}{fmt$(Math.abs(c.net))}</td>
                    <td style={tdRS}>{pill(c.ret)}</td>
                    <td style={tdRS}>{c.mcc>0?<span style={{fontSize:11,fontWeight:500,color:"#dc2626"}}>{c.mcc}</span>:"—"}</td>
                    <td style={tdRS}>{c.bcc>0?<span style={{fontSize:11,fontWeight:500,color:"#dc2626"}}>{c.bcc}</span>:"—"}</td>
                    <td style={tdS}>
                      {c.mch.slice(0,2).map((a,i)=><span key={i} style={{display:"inline-block",fontSize:10,padding:"1px 6px",borderRadius:10,background:"#F4F6FB",color:"#808080",border:"0.5px solid rgba(41,53,93,.1)",margin:"1px 2px 1px 0"}}>MC: {a}</span>)}
                      {c.bch.slice(0,1).map((a,i)=><span key={i} style={{display:"inline-block",fontSize:10,padding:"1px 6px",borderRadius:10,background:"rgba(220,38,38,.06)",color:"#991b1b",border:"0.5px solid rgba(220,38,38,.2)",margin:"1px 2px 1px 0"}}>BC: {a}</span>)}
                    </td>
                  </tr>
                  {isExp&&(
                    <tr>
                      <td colSpan={9} style={{padding:"10px 12px",background:"#F4F6FB",borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>
                        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,fontSize:12}}>
                          <div>
                            <div style={{fontWeight:500,marginBottom:8,fontSize:11,textTransform:"uppercase",color:"#808080"}}>MC churned ({c.mcc} of {c.mca} active)</div>
                            {c.mch.length ? c.mch.map((a,i)=><div key={i} style={{padding:"4px 0",borderBottom:"0.5px solid rgba(41,53,93,.06)"}}>{a}</div>) : <span style={{color:"#808080"}}>None this quarter</span>}
                          </div>
                          <div>
                            <div style={{fontWeight:500,marginBottom:8,fontSize:11,textTransform:"uppercase",color:"#808080"}}>BC churned ({c.bcc} of {c.bca} active)</div>
                            {c.bch.length ? c.bch.map((a,i)=><div key={i} style={{padding:"4px 0",borderBottom:"0.5px solid rgba(41,53,93,.06)",color:"#991b1b"}}>{a}</div>) : <span style={{color:"#808080"}}>None this quarter</span>}
                          </div>
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );

  const renderChurn = () => {
    const churnCSMs = csms.filter(c=>c.mcc>0||c.bcc>0).sort((a,b)=>(b.mcc+b.bcc)-(a.mcc+a.bcc));
    const allMcc = csms.reduce((s,c)=>s+c.mcc,0);
    const allBcc = csms.reduce((s,c)=>s+c.bcc,0);
    const allMca = csms.reduce((s,c)=>s+c.mca,0);
    const allBca = csms.reduce((s,c)=>s+c.bca,0);
    return (
      <div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,minmax(0,1fr))",gap:12,marginBottom:20}}>
          {[
            {l:"MC accounts churned",v:allMcc,s:`${allMca.toLocaleString()} active — ${(allMcc/Math.max(allMca,1)*100).toFixed(1)}% rate`,col:"#dc2626"},
            {l:"BC accounts churned",v:allBcc,s:`${allBca.toLocaleString()} active — ${allBca>0?(allBcc/allBca*100).toFixed(1)+"% rate":"n/a"}`,col:"#dc2626"},
            {l:"CSMs with churn",v:churnCSMs.length,s:`of ${csms.length} total`,col:"#d97706"},
          ].map(m=>(
            <div key={m.l} style={{background:"#ECEEF1",borderRadius:8,padding:14,position:"relative",overflow:"hidden"}}>
              <div style={{position:"absolute",top:0,left:0,right:0,height:2,background:m.col,borderRadius:"8px 8px 0 0"}}/>
              <div style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:6}}>{m.l}</div>
              <div style={{fontSize:22,fontWeight:500,color:m.col,lineHeight:1,marginBottom:3}}>{m.v}</div>
              <div style={{fontSize:11,color:"#808080"}}>{m.s}</div>
            </div>
          ))}
        </div>
        <div style={S.card}>
          <div style={{fontSize:11,textTransform:"uppercase",color:"#808080",fontWeight:500,marginBottom:12}}>Churned accounts by CSM — click row to expand</div>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr>
              <th style={thS}>CSM</th><th style={thS}>Coach</th>
              <th style={thRS}>MC churn</th><th style={thRS}>BC churn</th>
              <th style={thRS}>MC active</th><th style={thRS}>MC rate</th>
              <th style={thS}>Accounts</th>
            </tr></thead>
            <tbody>
              {churnCSMs.map(c=>{
                const isExp=expandedBob===c.n+"_ch";
                const col=TEAM_COLS[COACHES.find(x=>x.n===c.c)?.t||""]||"#808080";
                const mcRate=c.mca>0?c.mcc/c.mca:0;
                return (
                  <React.Fragment key={c.n}>
                    <tr style={{cursor:"pointer"}} onClick={()=>setExpandedBob(n=>n===c.n+"_ch"?null:c.n+"_ch")}>
                      <td style={{...tdS,fontWeight:500}}>{c.n}</td>
                      <td style={tdS}><span style={{display:"inline-block",width:7,height:7,borderRadius:"50%",background:col,marginRight:4,verticalAlign:"middle"}}/><span style={{color:"#808080"}}>{c.c.split(" ").pop()}</span></td>
                      <td style={tdRS}>{c.mcc>0?<span style={{fontWeight:500,color:"#dc2626"}}>{c.mcc}</span>:"—"}</td>
                      <td style={tdRS}>{c.bcc>0?<span style={{fontWeight:500,color:"#dc2626"}}>{c.bcc}</span>:"—"}</td>
                      <td style={{...tdRS,color:"#808080"}}>{c.mca}</td>
                      <td style={tdRS}>
                        <span style={{fontSize:10,fontWeight:500,padding:"2px 8px",borderRadius:20,
                          background:mcRate<0.05?"rgba(22,163,74,.1)":mcRate<0.1?"rgba(217,119,6,.1)":"rgba(220,38,38,.1)",
                          color:mcRate<0.05?"#166534":mcRate<0.1?"#854d0e":"#991b1b"}}>
                          {fmtP(mcRate)}
                        </span>
                      </td>
                      <td style={tdS}>
                        {c.mch.slice(0,2).map((a,i)=><span key={i} style={{display:"inline-block",fontSize:10,padding:"1px 6px",borderRadius:10,background:"#F4F6FB",color:"#808080",border:"0.5px solid rgba(41,53,93,.1)",margin:"1px 2px 1px 0"}}>{a}</span>)}
                        {c.bch.slice(0,1).map((a,i)=><span key={i} style={{display:"inline-block",fontSize:10,padding:"1px 6px",borderRadius:10,background:"rgba(220,38,38,.06)",color:"#991b1b",border:"0.5px solid rgba(220,38,38,.2)",margin:"1px 2px 1px 0"}}>{a}</span>)}
                      </td>
                    </tr>
                    {isExp&&(
                      <tr>
                        <td colSpan={7} style={{padding:"10px 12px",background:"#F4F6FB",borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>
                          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,fontSize:12}}>
                            <div>
                              <div style={{fontWeight:500,marginBottom:8,fontSize:11,textTransform:"uppercase",color:"#808080"}}>MC churned ({c.mcc} of {c.mca})</div>
                              {c.mch.map((a,i)=><div key={i} style={{padding:"4px 0",borderBottom:"0.5px solid rgba(41,53,93,.06)"}}>{i+1}. {a}</div>)}
                            </div>
                            <div>
                              <div style={{fontWeight:500,marginBottom:8,fontSize:11,textTransform:"uppercase",color:"#808080"}}>BC churned ({c.bcc} of {c.bca})</div>
                              {c.bch.length ? c.bch.map((a,i)=><div key={i} style={{padding:"4px 0",color:"#991b1b",borderBottom:"0.5px solid rgba(41,53,93,.06)"}}>{i+1}. {a}</div>) : <span style={{color:"#808080"}}>None</span>}
                            </div>
                          </div>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    );
  };

  return (
    <div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16}}>
        <div style={{display:"flex",gap:2,background:"#ECEEF1",borderRadius:8,padding:3}}>
          {[["overview","Overview"],["table","CSM table"],["churn","Churn detail"]].map(([t,l])=>(
            <button key={t} onClick={()=>setBobTab(t)}
              style={{padding:"5px 14px",fontSize:12,fontWeight:500,border:"none",borderRadius:6,cursor:"pointer",
                background:bobTab===t?"#fff":"transparent",color:bobTab===t?"#29355D":"#808080",
                boxShadow:bobTab===t?"0 1px 3px rgba(0,0,0,.08)":"none"}}>
              {l}
            </button>
          ))}
        </div>
        <span style={{fontSize:12,color:"#808080"}}>
          Monthly data · Overall: <strong style={{color:rCol(liveGrand.pct)}}>{fmtP(liveGrand.pct)}</strong> vs <strong>91%</strong> goal
        </span>
      </div>
      {bobTab==="overview" && renderOverview()}
      {bobTab==="table"    && renderTable()}
      {bobTab==="churn"    && renderChurn()}
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
  const [fontScale, setFontScale] = useState(()=>{try{return parseFloat(sessionStorage.getItem(FONT_KEY)||"1");}catch(e){return 1;}});
  const changeFontScale = (delta) => setFontScale(s=>{
    const next = Math.min(1.4, Math.max(0.8, Math.round((s+delta)*10)/10));
    try{sessionStorage.setItem(FONT_KEY, String(next));}catch(e){}
    return next;
  });
  const [csms, setCSMs] = useState([]);
  const [tab, setTab] = useState("coaching");
  const [filterCoach, setFilterCoach] = useState("");
  const [filterCSM, setFilterCSM] = useState("");
  const [status, setStatus] = useState("loading");
  const [updatedAt, setUpdatedAt] = useState(null);
  const [history, setHistory] = useState([]);
  const [skippedCSMs, setSkippedCSMs] = useState([]);
  const [rawRev, setRawRev] = useState([]);
  const [bobData, setBobData] = useState(null);
  const [bobRaw,  setBobRaw]  = useState({bob:{},coachTotals:{},grand:null});
  const [mcChurn, setMcChurn] = useState({});
  const [bcChurn, setBcChurn] = useState({});

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

  // Full load on mount (all 6 tabs)
  useEffect(()=>{
    if (!unlocked) return;
    setStatus("loading");

    // Store non-revenue data so revenue polls can reuse it
    let latestEmail=[], latestCad=[], latestDue=[], latestOntime=[], latestHistory=[], latestSkipped=[], latestBob=[], latestMcChurn=[], latestBcChurn=[];

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
        fetchCSV(CSV_MC_CHURN).catch(()=>[]),
        fetchCSV(CSV_BC_CHURN).catch(()=>[]),
      ]).then(([revRows, emailRows, cadRows, dueRows, ontimeRows, historyRows, skippedRows, bobRows, mcRows, bcRows]) => {
        latestEmail   = emailRows;
        latestCad     = cadRows;
        latestDue     = dueRows;
        latestOntime  = ontimeRows;
        latestHistory = historyRows;
        latestSkipped = skippedRows;
        latestBob      = bobRows;
        latestMcChurn  = mcRows;
        latestBcChurn  = bcRows;
        setRawRev(revRows);
        const rev     = mapRev(revRows);
        const email   = mapEmail(emailRows);
        const cad     = mapCadence(cadRows);
        const due     = mapDue(dueRows);
        const ontime  = mapOnTime(ontimeRows);
        const skipped = mapSkipped(skippedRows);
        const built   = buildCSMs(rev, email, cad, due, ontime, skipped, bobRaw, mapChurn(mcRows), mapChurn(bcRows));
        setCSMs(built);
        setSkippedCSMs(built.filter(c=>c.skippedCount>0).sort((a,b)=>b.skippedCount-a.skippedCount));
        setBobRaw(mapBob(bobRows));
        setMcChurn(mapChurn(mcRows));
        setBcChurn(mapChurn(bcRows));
        setHistory(mapHistory(historyRows));
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
        const built   = buildCSMs(rev, email, cad, due, ontime, skipped, bobRaw, mapChurn(latestMcChurn), mapChurn(latestBcChurn));
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
          {["coaching","overview","revenue","bob","leaderboard","activity","trends"].map(t=>(
            <button key={t} onClick={()=>setTab(t)}
              style={{padding:"10px 18px",fontSize:13,fontWeight:500,color:tab===t?"#fff":"rgba(255,255,255,.55)",background:"transparent",border:"none",cursor:"pointer",borderBottom:tab===t?"3px solid #FF5000":"3px solid transparent",whiteSpace:"nowrap"}}>
              {t==="coaching"?"Coaching":t==="trends"?"📈 Trends":t==="revenue"?"💰 Revenue":t==="bob"?"📋 Book of Business":t.charAt(0).toUpperCase()+t.slice(1)}
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
        <div style={{padding:"20px 24px",zoom:fontScale}}>
          {tab==="coaching"&&<CoachingView csms={filteredCSMs} coach={filterCoach} onSelectCSM={selectCSMFn} onSelectCoach={e=>{setFilterCoach(e);setFilterCSM("");}} onClear={()=>{setFilterCoach("");setFilterCSM("");}} skippedCSMs={skippedCSMs.filter(c=>{const i=lk(c.name);if(filterCoach&&(i&&i.c)!==filterCoach)return false;if(filterCSM&&c.name!==filterCSM)return false;return true;})}/>}
          {tab==="overview"&&<OverviewView csms={filteredCSMs} allCSMs={csms}/>}
          {tab==="leaderboard"&&<LeaderboardView csms={filteredCSMs}/>}
          {tab==="activity"&&<ActivityView csms={filteredCSMs}/>}
          {tab==="revenue"&&<RevenueView rawRev={rawRev} csms={filteredCSMs} filterCoach={filterCoach} filterCSM={filterCSM}/>}
          {tab==="bob"&&<BobView filterCoach={filterCoach} filterCSM={filterCSM} bobRaw={bobRaw} mcChurn={mcChurn} bcChurn={bcChurn}/>}
          {tab==="trends"&&<TrendsView history={history} csms={filteredCSMs} filterCoach={filterCoach} filterCSM={filterCSM}/>}
        </div>
      )}
    </div>
  );
}
