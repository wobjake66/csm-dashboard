import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";

const PIN = "thryv2025";
const PIN_KEY = "csm_pin_v1";

const CSV_REV     = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=1721544342&single=true&output=csv"; // live JotForm sync
const CSV_EMAIL   = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=0&single=true&output=csv";
const CSV_CAD     = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=1973544046&single=true&output=csv";
const CSV_DUE     = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=341836664&single=true&output=csv";
const CSV_ONTIME  = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=459845057&single=true&output=csv";
const CSV_SKIPPED = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=1238903633&single=true&output=csv"; // prior-day skipped cadences
const CSV_HISTORY = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRiYN66PuGwyOhd2jC1gHVv5Zv1ub5vxTZU8uCQ5k1OXNbYL8NFHdonbmb7zzHpWkAooXv9P8LoCufo/pub?gid=162960918&single=true&output=csv";

const CAD_ACCTS={"Taylor Kidd":[{"n":"F Lees Tax & Accounting Services","ott":3,"oto":2},{"n":"Liz Bienstock, Realtor","ott":1,"oto":0},{"n":"Nixdorf Tree Service","ott":1,"oto":0},{"n":"Torrey Tucker Electric LLC","ott":1,"oto":1},{"n":"Uniquely Yours Screen Printing","ott":3,"oto":3}],"April Hall":[{"n":"JaysonRachel Carter","ott":1,"oto":0},{"n":"M & C Homecare","d":[{"t":"Highlight Video","due":"5/29/2026","ov":true,"nw":false}]},{"n":"TAG Construction","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]}],"Zoltan Rudolf":[{"n":"Budget Screens & Awnings","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Classic Steamboat Cruises","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Clean Planet West Auckand","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"North East Survey Design","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Perfect Fit Interior","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"VISION INSTALLATIONS LIMITED","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]}],"Libby Booher":[{"n":"Advanced Acupuncture And Chinese Herbal Clinic","ott":1,"oto":0},{"n":"Battle Ground Personal Training","ott":3,"oto":1},{"n":"CAMPBELL REMODELING AND MAINTENANCE LLC","ott":3,"oto":1},{"n":"Discount Auto Glass","ott":3,"oto":1},{"n":"Hawaii Food Products Inc","ott":1,"oto":0},{"n":"Motorcycle Rider Training","ott":3,"oto":1},{"n":"Zebra Stripes Child Care & Preschool","ott":1,"oto":1}],"Indu Vijay":[{"n":"Blinds Hub","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Bold Insurance","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Byte Size Labs","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Oz Seals","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Tranquility Pools","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Luis Aguasvivas Peralta":[{"n":"Mission Success Solutions LLC","ott":1,"oto":1}],"Mark Velazquez":[{"n":"Antesberger Plumbing","ott":1,"oto":1},{"n":"CJ Construction","ott":1,"oto":1},{"n":"CJS Heating and Air","ott":1,"oto":0},{"n":"FYZICAL Therapy & Balance Center of Little Silver","ott":1,"oto":0},{"n":"Home Pro Masters","ott":1,"oto":0},{"n":"Korte Does It All, Inc.","ott":1,"oto":0},{"n":"Legacy Heating and Air, Inc.","ott":1,"oto":0},{"n":"Newtown Heating & Air Conditioning Inc","ott":3,"oto":3},{"n":"Precision Today Plumbing Heating Cooling Electrical","ott":1,"oto":0},{"n":"Stitches and Screens","ott":2,"oto":1}],"Sylvia Appla":[{"n":"A F Isaac Suveying Pty Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"AESTHETIC DENTAL LIMITED","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"AK Painting","ott":1,"oto":0},{"n":"ASAP Security and Protection","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"ATOM TREE SPECIALISTS LIMITED","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Acland Street Physiotherapy","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"All Four x 4 Spares","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Allen's Interiors & Building","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"BYC Dental Pty Ltd ATF BYC Dental Trust","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Bayview Denture Care","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Cavana Services","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Country Tyres Yass","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Darwin Family Law","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Dons Mobile Marine","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"EK GLASS LIMITED","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Earth 2 Ocean Communications","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Exacte Advisors","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Exo Graphics","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Fencemen Fencing Contractors Ltd.","d":[{"t":"Email","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Flash Roofing Supplies","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Gold Coast Hypnotherapy","ott":1,"oto":0},{"n":"HERRON TODD WHITE","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Hyperspace Architecture","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Informed Decisions Consultancy Limited","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"J S De Rooy","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"KITCHEN & APPLIANCES SOUTHLAND LIMITED","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"MURRAY BROWN ELECTRICAL (1996)","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Nerang Stockfeed","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Northlane Welding Services","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Nudge Osteopathy","d":[{"t":"Email","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Oxen Lawyers","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"PNJ Hire","ott":1,"oto":0},{"n":"Pearla Plumbing & Electrical","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Penrith Engine Services","d":[{"t":"Email","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Perth Taxi Booking","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"PowerSmart Heat Pumps","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Quickturn Pty Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Red Roo Solutions","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"SHRI METALS MERCHANT PTY LTD","ott":1,"oto":0},{"n":"Scott Electrics","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Shubbs International Pty Ltd","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Smart Thread Solutions","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Straightline Guttering Pty Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Streamline Drains & Pipelines (NSW) PL","d":[{"t":"Email","due":"6/1/2026","ov":true,"nw":true}]},{"n":"The Baker's Den Bakery Cafe","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Village Carpet Care","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Voodoo Rod And Custom","d":[{"t":"Email","due":"6/1/2026","ov":false,"nw":true}]},{"n":"WORKSTORE SA LIMITED","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Wagga Scrap Metals","d":[{"t":"Email","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Wide Bay Memorials","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Wollongong Auto Excellence","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]}],"Rafael Sencion Sencion":[{"n":"5 Star Fence","ott":1,"oto":1},{"n":"D & D Overhead Door LLC","ott":1,"oto":1},{"n":"Grow With Us Learning Centers","ott":1,"oto":1},{"n":"HVAC Pros Los Angeles","ott":1,"oto":1},{"n":"Pampered Paws Pet Grooming","ott":1,"oto":1},{"n":"Patrick Alley Handyman","ott":1,"oto":1},{"n":"Technology Networks","ott":1,"oto":1},{"n":"Uptown General Aesthetic Dentistry P A","ott":1,"oto":1},{"n":"Woodforest Family Chiropractic Clinic","ott":1,"oto":1}],"Katelyn Ankrom":[{"n":"Aberdeen Cemetery Assoc","ott":1,"oto":0},{"n":"Gleason's Salinas RV","ott":1,"oto":0},{"n":"Neilson Roy Plumbing","ott":1,"oto":0}],"Kellie Lester":[{"n":"Chatham Laschober","ott":1,"oto":1},{"n":"D &T Power Washing","ott":3,"oto":3},{"n":"East Cooper Lock & Safe","ott":1,"oto":1},{"n":"Lewis Brothers Inc","ott":1,"oto":1},{"n":"Mansfield Well Drilling Inc","ott":1,"oto":0},{"n":"Martin Landscaping","ott":1,"oto":1},{"n":"Metro Appliance Repair","ott":3,"oto":3},{"n":"Mikes Seal Coating & Services","ott":1,"oto":0},{"n":"Pleasure Pools","ott":2,"oto":2},{"n":"WBC Carpentry Corp","ott":1,"oto":0}],"Ashley Shaffer":[{"n":"A Dion & Son Floor Contractors","ott":1,"oto":0},{"n":"Absolute Roofing & Remodeling","ott":3,"oto":3},{"n":"Alaska Urgent Care LLC","ott":1,"oto":1},{"n":"Benny Electric Inc Benny Electric","ott":3,"oto":3},{"n":"Body Sculpt Skin Laser","ott":3,"oto":3},{"n":"Cement Hill Storage","ott":1,"oto":1},{"n":"Dream Work Diesel","ott":1,"oto":1},{"n":"Fox Valley Glass Inc","ott":1,"oto":1},{"n":"Garber Surveying Service PA","ott":3,"oto":3},{"n":"Just Cuz Plumbing LLP","ott":3,"oto":3},{"n":"Moody Construction Service","ott":1,"oto":1},{"n":"Pryor Automatic Fire Sprinkler Inc","ott":3,"oto":3},{"n":"Schaefers Stove & Spa","ott":3,"oto":3},{"n":"Seraphim Partners","ott":3,"oto":3},{"n":"Sound Decision","ott":1,"oto":0},{"n":"Testino Edward","ott":1,"oto":1},{"n":"Toms Superior Driving School Inc","ott":1,"oto":1},{"n":"Wasilla Medical Clinic","ott":1,"oto":1}],"Tyler Moeggenberg":[{"n":"AM PM Towing","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Alan Cherry Classic Interiors, LLC","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Fireplace Specialists LLC","ott":2,"oto":1},{"n":"GoodFellas Ristorante","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Newsom Fences","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Northwest Roofing & Construction LLC","d":[{"t":"15 Min Check-In","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Sarca Plumbing, Heating & Cooling","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Schnitzelbank","ott":2,"oto":1},{"n":"United Car Wash Gresham","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Up The Creek Heating & Air","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]}],"Elianny Tena Antigua":[{"n":"Burrows Heating & Cooling","ott":1,"oto":0},{"n":"Fleshers Fairview Health Care Center","ott":1,"oto":0},{"n":"G&S Tile Solutions LLC","ott":1,"oto":1},{"n":"JH Corner","ott":3,"oto":3}],"Saira Julian Guzman":[{"n":"Adama African Hair Braiding","ott":1,"oto":1},{"n":"Alert Door & Operator Co","ott":1,"oto":1},{"n":"E.M Imperial Auto Restoration","ott":1,"oto":1},{"n":"EZ Mini Storage","ott":1,"oto":1},{"n":"Kellets GC","ott":1,"oto":1}],"Peter Manalac":[{"n":"Julian's Appliance Centre","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"KEITH ROBERT HUGHES","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Lynfield Automotive Limited","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Stirling Floors","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"T & H Wreckers","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Victorian Chiropractic","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Misti Dixon":[{"n":"Dynamic Steam Carpet","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"R & G Plumbing And Drain Services Inc","d":[{"t":"Highlight Video","due":"5/26/2026","ov":true,"nw":false}]},{"n":"Riviera Floor Covering","ott":1,"oto":0},{"n":"Ryan Simmons","ott":2,"oto":1},{"n":"Witzke's Harry H Family Funeral Home Inc","ott":3,"oto":1}],"Heidi Torres Uribe":[{"n":"ArTron Heating And Air Conditioning","ott":1,"oto":1},{"n":"B Carroll Construction","ott":1,"oto":1},{"n":"Charlie Rice Roofing","ott":3,"oto":2},{"n":"Cobblestone Quality Shoe Repair","ott":2,"oto":2},{"n":"Dynamic Garage Door of Hibbing, LLC","ott":1,"oto":1},{"n":"Jones Animal Hospital","ott":2,"oto":2},{"n":"Soo Kool Air Conditioning","ott":1,"oto":1}],"Yessica Montero Urena":[{"n":"AAA London Chimney Sweep","ott":2,"oto":2},{"n":"Cura Healthcare Consultants","ott":1,"oto":1},{"n":"Josh's Mobile Mechanic","ott":1,"oto":1},{"n":"Pope County Title Co","ott":1,"oto":1},{"n":"Simplified Living Home Services","ott":1,"oto":0}],"Jathzelyn Elizabeth Fortuna Paulino":[{"n":"Apex Enterprise Roofing","ott":1,"oto":1},{"n":"Cross Cultural Communications","ott":1,"oto":1},{"n":"Door To Door Movers & Apartment Movers","ott":1,"oto":1},{"n":"Etheridge Hamlett & Murray LLP","ott":1,"oto":1},{"n":"Gingras Plumbing & Heating Inc","ott":1,"oto":1},{"n":"Henry Gitner Philatelists Inc","ott":1,"oto":1},{"n":"Park Ave Med Spa","ott":1,"oto":1},{"n":"Trent and Sons Roofing","ott":1,"oto":1},{"n":"United Car Wash","ott":2,"oto":2}],"Misty Decatur":[{"n":"Advance Insurance & Benefits","ott":1,"oto":0},{"n":"Arpys Construction & Remodeling","ott":1,"oto":0},{"n":"Auto Glass Xperts","ott":3,"oto":2},{"n":"Automotive Paint Supply APS","ott":1,"oto":1},{"n":"Centerscale Automation Hawaii Inc","ott":1,"oto":0},{"n":"Northern Arizona Roof Services LLC","ott":2,"oto":1},{"n":"Perfect World Pavers","d":[{"t":"Highlight Video","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Piazzau2019s Top Gun Coatings","ott":1,"oto":0},{"n":"Quality Transmission & Auto Repair","ott":1,"oto":1},{"n":"Ridge and Remedy Apotheracy, LLC","ott":3,"oto":3},{"n":"Taylor Regional Hospital","ott":3,"oto":3},{"n":"We Can Do More","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]}],"Barbara Larrosa Presinal":[{"n":"Charleston Premier Workz","ott":3,"oto":2},{"n":"Covert Water Heaters Inc","ott":1,"oto":1},{"n":"El Maya Mexican Grill","ott":3,"oto":2},{"n":"G Shelley Basement Waterproofing","ott":1,"oto":1},{"n":"GCS Heating and Cooling","ott":1,"oto":1},{"n":"Greenes Rolloff Service","ott":1,"oto":1},{"n":"Insights Eyecare, PA","ott":1,"oto":1},{"n":"Jennings Home Rejuvenation","ott":1,"oto":1},{"n":"Lembke Inc","ott":3,"oto":2},{"n":"Noah Autos","ott":1,"oto":1},{"n":"Smart House Remodeling","ott":1,"oto":1},{"n":"The Lukaart Agency, a Farm Bureau Agency","ott":2,"oto":1},{"n":"Ultra Steam Cleaning","ott":3,"oto":1},{"n":"Villegas roofing llc","ott":1,"oto":1}],"Ashley Vasquez Mena":[{"n":"Ideal Exteriors","ott":1,"oto":0},{"n":"Nashville Pizza Company","ott":1,"oto":0},{"n":"Ronco Tech Heating & Cooling LLC","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Sweers Roofing","ott":2,"oto":0},{"n":"Tug River Black Lung Clinic","ott":1,"oto":0}],"Warda Gul":[{"n":"Achieve Training & Assessment Services Pty","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"BL Microtek Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Chrysus Group","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Constructways Pty Ltd Trading as Stella Arden","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Creative 2000 Blinds & Awnings","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"D & M Brown Concreting","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Elite Body & Paint Pty Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Elite Tyre & Autocare Bacchus Marsh","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"GJM Balustrading","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"KCP Physiotherapy Paraparaumu","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Prodigy Design Plastics Pty Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Tauranga Roofing & Scaffolding","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Sarah Swanson":[{"n":"247 Protective Services","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"402 Castillo Contractor","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"A Head Of The Times Hair Design","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"A1 Pro Roofing LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"AAA Air Conditioning and Refrigeration","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"AAA OnSite Notary LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"AL Hansen Paint & Metal Shop Inc","ott":3,"oto":2},{"n":"ALL ABOUT PRESSURE","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Adele Home Health Agency","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"All Fence Supply Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"American Coins & Collectibles","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Amuse Media","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Art Tabasco & Sons Paving","ott":1,"oto":1},{"n":"Arthur J E & Associates Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Big Rig Truck Repair and Towing","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Bill's Fixit Shop","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Bradleys Construction","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Brenda Kashella Electrolysis Medical & Cosmetic","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"CCB Outdoor Living & Construction","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Caesars Italian Delcatessen","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Castle Rock HeatingAir","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Celco Community Credit Union","ott":1,"oto":1},{"n":"Clearwater Beach Boat Rentals","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Coast Pipe","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Colby Pacific Family Dentistry","ott":3,"oto":3},{"n":"ESC Cabinets","ott":1,"oto":1},{"n":"Electric Plus Inc.","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Elite Steel Building Systems","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Exquisite Body Clinic","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Forever Fence Solutions","ott":1,"oto":0},{"n":"Get It Now Print","ott":1,"oto":0},{"n":"Goodman General Construction Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"GraceWorks Bookkeeping LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Gregs Job Squad","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Guide Path Advisors","ott":3,"oto":2},{"n":"Hero Cleaning","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Hounds On The Hill","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"James Remodeling","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Jerry McLeod","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Josseline Carr","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Kevin Yul Wright JD - Business Loan Success Academy Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Lotsa Fence Options","ott":1,"oto":1},{"n":"Maria Rivero","ott":3,"oto":2},{"n":"Martin Electrical Services","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"McAlister McAlister & Nicklas PLLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Meilus Precision Therapy","ott":3,"oto":3},{"n":"Norvill Construction Co","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"One Day Safe Shower Va","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Paige Marie Photography","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Paradise Valley Land Solutions, LLC","ott":3,"oto":3},{"n":"Park Rapids Glass","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Pawfect Pawtions","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Polar Pure Water","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Precision Property Cleanup and Junk Removal","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"ProBend Corp","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Psychic Center Botanica","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"RCR Inspections Inc","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Raytech LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Rebecca Ortenzio Lee Orthodontics","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Rodger's Roll-Up Garage Doors","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"SOS Services Plumbing LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"SafeGuard Roofing and Siding","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Salty Breeze Rentals","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Sherry Smalling","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Simons Seamless Gutters","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Skylark North Glider Flight School","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Smith Brothers Funeral Home","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Snap2Throw Quarterback Academy","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"South Salem Mini Storage","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Speedy Pumping","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Stars and Stripes VA Benefits","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Stone Dumpster Rentals","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Super Electric","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Tavo Matic Inc","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"The Plumbing Company Of East Tennessee","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"The Yoga Experience LLC","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Toms RV Service & Sales","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Trank Insurance","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Triple H Paving","ott":1,"oto":0},{"n":"True Hand Roofing","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Two Men And A Snake","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Tycoon Solutions LLC","ott":1,"oto":1},{"n":"Unlimited Choices Greek Boutique Inc","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Uplifted Exteriors","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Vahe Dental","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Valora Behavior Support Centers","ott":2,"oto":2},{"n":"Webb Floors","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"White-Lavender Plumbing and Service","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Winkler Construction & Crane Co Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"sea pro home renovation","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]}],"Tracy-Ann Gaudencio":[{"n":"Cheshire Contractors","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Elite Pods","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Gina King Naturopathy","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Pro Pacific Concreting","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Smart Bookkeepers","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Karissa Hernandez":[{"n":"ABC Southwest Plumbing Inc","ott":1,"oto":0},{"n":"Air Conditioning & Heating Solutions LLC","ott":1,"oto":0},{"n":"Change your Thinking, Change Your Life LLC","ott":1,"oto":0},{"n":"Decca Recruiting LLC","ott":1,"oto":0},{"n":"Doyles Heat & Air Services LLC","ott":1,"oto":0},{"n":"Quality Restoration Services","ott":3,"oto":2},{"n":"Tidewater Landscape LLC","ott":1,"oto":0},{"n":"Triple A Sanitation","ott":3,"oto":2},{"n":"Veteran Floors Inc","ott":1,"oto":0}],"Anthony Yen":[{"n":"GTZ ROOFING","ott":3,"oto":3},{"n":"Keystone Pump & Well Service","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Tri-County Chimney Service","ott":1,"oto":0}],"Yolanda Ramirez":[{"n":"Best Electric","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Full Circle Payment Processing","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Mills Chiropractic","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Sarasota Golf Cart Sales - Custom Carts & Repairs","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]}],"Deivis Pena":[{"n":"A R McClung Construction Co","ott":3,"oto":3},{"n":"Aboite Boarding & Grooming","ott":2,"oto":1},{"n":"All About Kids Preschool","ott":1,"oto":1},{"n":"All Steamed Up carpet and Upholstry cleaning","ott":1,"oto":1},{"n":"Crown Tree Care Inc","ott":1,"oto":1},{"n":"Forever Young Landscaping","ott":3,"oto":3},{"n":"JB Hauling","ott":3,"oto":3},{"n":"Roberto's Tile & More","ott":1,"oto":1},{"n":"Zoom Business Brokers","ott":3,"oto":3}],"Damita Hill":[{"n":"A Plus Foundation LLC","d":[{"t":"Highlight Video","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Concord Heating & Air Conditioning Inc","ott":1,"oto":1},{"n":"Ohana Clean Air","d":[{"t":"Highlight Video","due":"6/1/2026","ov":false,"nw":true}]},{"n":"Top Notch Moving Company","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]},{"n":"Williams Auto Parts Inc","d":[{"t":"Highlight Video","due":"6/1/2026","ov":true,"nw":true}]}],"Karmita Turner":[{"n":"Asphalt Services","ott":1,"oto":1},{"n":"Bemis Well Drilling & Water Conditioning, LLP","ott":1,"oto":1},{"n":"Creative Improvements Inc","ott":2,"oto":1},{"n":"Danny Odom Roofing","ott":3,"oto":3},{"n":"Himmelstein Louis","ott":1,"oto":0},{"n":"Holbert's Tree service","ott":2,"oto":2},{"n":"IntelliPEST","ott":2,"oto":1},{"n":"J Pop Landscaping","ott":1,"oto":1},{"n":"Law Office of Gayle A Belcher","ott":1,"oto":1},{"n":"Miss Miranda Bail Bond Services","ott":1,"oto":0},{"n":"New Look Exteriors","ott":2,"oto":1},{"n":"Quality Tree Care & Landscaping","ott":1,"oto":0},{"n":"Rays Septic Tank Service","ott":1,"oto":0},{"n":"Tab Mechanical Services","ott":1,"oto":1},{"n":"Valley Restaurant and Catering","ott":1,"oto":1}],"Kennedy Sanchez":[{"n":"4 C's Construction","ott":1,"oto":0},{"n":"On the Go Experience","ott":1,"oto":0}],"Felix Caba Jimenez":[{"n":"Action Counseling LLC","ott":1,"oto":1}],"Dorka Frias Lantigua":[{"n":"Copeland Fencing and Construction","ott":2,"oto":0},{"n":"Herbs Door Service","ott":1,"oto":1},{"n":"Insurance Answers Inc","ott":1,"oto":0},{"n":"Meehan's Lawn Service","ott":1,"oto":1},{"n":"Paxton Senior Insurance Service LLC","ott":1,"oto":0},{"n":"Rotterdam Heating","ott":3,"oto":1},{"n":"Ruth & Co. Events","ott":1,"oto":0},{"n":"Southern Living Exteriors","ott":2,"oto":0},{"n":"Wallace Heating & Air","ott":3,"oto":1}],"Kyle Dye":[{"n":"Adams Carpet Center","ott":1,"oto":0},{"n":"Allstate Paving Inc","ott":1,"oto":0},{"n":"Carmody James","ott":1,"oto":0},{"n":"Cynthia Poole","ott":3,"oto":1},{"n":"Final Exterminators","ott":1,"oto":1},{"n":"Formals Only Tuxedos","ott":3,"oto":1},{"n":"Granados Electrical Service","ott":3,"oto":1},{"n":"Gregorio's Pizzeria & Trattoria","ott":1,"oto":1},{"n":"Knabusch Insurance Services Inc","ott":1,"oto":0},{"n":"Lawrence Building Corp","ott":3,"oto":1},{"n":"Magidov CPA Firm","ott":3,"oto":1},{"n":"Majestic Jewelers","ott":3,"oto":1},{"n":"Mobility Plus Cincinnati East","ott":1,"oto":0},{"n":"Robinsons Paint & Wallpaper","ott":1,"oto":0},{"n":"SavMor Upholstery Co Inc","ott":1,"oto":0},{"n":"Speedy Locksmith Inc","ott":1,"oto":0},{"n":"Spokane Cosmetic Dentistry","ott":3,"oto":1},{"n":"Tess African Hair","ott":1,"oto":1},{"n":"The Caring Heart, LLC","ott":1,"oto":0},{"n":"The Grass Company of San Antonio","ott":1,"oto":0}],"Alejandro Rodriguez-Medina":[{"n":"Anwell Mobile Homes","ott":2,"oto":0},{"n":"Be Easy Bail Bonds","ott":1,"oto":0},{"n":"Blinds For Less","ott":1,"oto":0},{"n":"Boss Momma Boutique","ott":1,"oto":0},{"n":"Calvin Turner Roofing","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Celtic Moving & Storage Co","ott":3,"oto":0},{"n":"Chanler Agency Inc","ott":2,"oto":0},{"n":"Doug's Rooter Service","ott":1,"oto":0},{"n":"Duct -Tec","ott":1,"oto":0},{"n":"E-SQUARED ROOFING LLC","d":[{"t":"15 Min Check-In","due":"6/1/2026","ov":false,"nw":true}],"ott":3,"oto":1},{"n":"Fox Jewelers","ott":1,"oto":0},{"n":"Georges British American Auto Repair","ott":2,"oto":0},{"n":"Grand Slam Garage Door Services","ott":2,"oto":0},{"n":"Grass Roots Landscapes","ott":1,"oto":0},{"n":"Greg Smith","ott":1,"oto":0},{"n":"Hoffman Concrete, LLC","ott":3,"oto":1},{"n":"J & B Tree Services LLC","ott":1,"oto":0},{"n":"Logan Diving & Salvage","ott":3,"oto":0},{"n":"Manosh Singh and Associates","ott":2,"oto":0},{"n":"Mark Worleys Construction","ott":2,"oto":0},{"n":"Marks Roofing Company","ott":3,"oto":0},{"n":"O'Fallon Garage","ott":1,"oto":0},{"n":"Oakwood Landscaping LLC","ott":1,"oto":1},{"n":"Pro-Kleen","ott":2,"oto":0},{"n":"Ray donch Body werks Inc","d":[{"t":"15 Min Check-In","due":"6/1/2026","ov":true,"nw":true}],"ott":3,"oto":1},{"n":"Robbs Innova Construction","ott":1,"oto":0},{"n":"Rosso Nursery & Garden Center","ott":1,"oto":0},{"n":"Roy's Auto Body","ott":1,"oto":0},{"n":"Shaddai Construction","ott":2,"oto":0},{"n":"Stickley John R","ott":1,"oto":0},{"n":"Twin Cities Flag Source","ott":1,"oto":0},{"n":"VG Bail Bonds","ott":1,"oto":0},{"n":"Weis Landscaping Design","ott":1,"oto":0}],"Karen Capellan Tavarez":[{"n":"AAA Mini Storage","ott":1,"oto":1},{"n":"Best Tech Computer Service","ott":1,"oto":1},{"n":"BoozeeBar","ott":3,"oto":1},{"n":"Cali Roofing Inc","ott":3,"oto":1},{"n":"Cap Construction","ott":1,"oto":1},{"n":"Cecelia CookAssociates LLC","ott":1,"oto":1},{"n":"Challenge Family Fun Center","ott":1,"oto":1},{"n":"Enbalance Bodywork","ott":1,"oto":0},{"n":"Krystal Klear Cleaning Services","ott":1,"oto":1},{"n":"McLean Hardware Co, Inc","ott":1,"oto":1},{"n":"Peterman Bros Septic Service","ott":3,"oto":1},{"n":"Three Friends Tree Service","ott":1,"oto":1},{"n":"Vaca Valley Veterinary Hospital","ott":1,"oto":1},{"n":"Vital Essence Medical Spa","ott":3,"oto":1}],"Irina Larianni Molina Molina":[{"n":"Inaoly Auto Tech","d":[{"t":"Highlight Video","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Ivy Cremation Services of New York","ott":1,"oto":0},{"n":"Walters Mirror","ott":2,"oto":0},{"n":"Wells James DDS","ott":2,"oto":0}],"Wilson Mercedes":[{"n":"Achieve Wellness Drug Rehab New Jersey","ott":1,"oto":0},{"n":"Auto Images","ott":1,"oto":0},{"n":"Basement Systems Of NY","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Buckeye Crane & Hoist","ott":1,"oto":0},{"n":"Daystar Healthcare LLC","ott":1,"oto":0},{"n":"Fayette Veterinary Hospital","d":[{"t":"Highlight Video","due":"5/19/2026","ov":true,"nw":false}]},{"n":"Fresh Creek Plumbing & Heating","ott":1,"oto":0},{"n":"J Star Concrete","ott":1,"oto":0},{"n":"Joshua Paving","ott":1,"oto":0},{"n":"Kirsh Title Services","ott":1,"oto":0},{"n":"Next Level Athletes Born2Ball","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Northern Door Co Inc","ott":1,"oto":0},{"n":"On Point Pest Control","ott":1,"oto":0},{"n":"Premium Glass Tinting","ott":1,"oto":0},{"n":"SoMo Customs","ott":1,"oto":0},{"n":"The X-League","ott":1,"oto":0},{"n":"Zeeks Helpful Hands","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}],"ott":1,"oto":0}],"Matt Sword":[{"n":"DRH Construction Co., LLC","d":[{"t":"Highlight Video","due":"5/18/2026","ov":true,"nw":false}]},{"n":"Ericson Electric Inc","ott":1,"oto":0},{"n":"Levels Ahead Painting","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Pottenburgh Company","ott":1,"oto":0},{"n":"Root Revival Hair Restoration","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Sallie L Rubenzer Law Office","d":[{"t":"Highlight Video","due":"5/22/2026","ov":true,"nw":false}]},{"n":"Taxman Business Advisory Llc","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Texas Turf & Curb","ott":1,"oto":0},{"n":"The Chapel At Kerrville","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Veras King O Meats Inc 2 LLC","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Veras King O Meats Inc 3 LLC","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Wilson Blinds & Shutters","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Matt Daly":[{"n":"AR Electronics Systems Limited","ott":1,"oto":0},{"n":"De Silva Hebron","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Junior Explorers","ott":1,"oto":0},{"n":"Oleada Electrical","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"On Point Physio","ott":2,"oto":1},{"n":"Ulladulla Blinds & Home Improvements","ott":1,"oto":0}],"Nikita Siepen-Bowers":[{"n":"Brianna Tilt Trays & Towing Pty Ltd","ott":1,"oto":1},{"n":"Kenny's Painting Crew","d":[{"t":"Highlight Video","due":"5/31/2026","ov":true,"nw":false}]},{"n":"Mildura First Aid Services","ott":1,"oto":1},{"n":"Total Safe Compliance Group","d":[{"t":"Highlight Video","due":"5/31/2026","ov":true,"nw":false}]}],"Michael Furlong":[{"n":"AA Electric, Ltd.","d":[{"t":"Highlight Video","due":"5/29/2026","ov":true,"nw":false}]},{"n":"ASM Irrigation","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"All American Pro Paving & Roofing","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Earnest Well Drilling Inc","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Sati Ananda Pimentel Malespin":[{"n":"AD Dental","ott":1,"oto":0},{"n":"Abercrombie Transmission","ott":1,"oto":1},{"n":"Cenla Plumbing Repair LLC","ott":1,"oto":1},{"n":"Drive Line Service & Radiator King Inc","ott":1,"oto":0},{"n":"Fleetwood Foot & Ankle Center","ott":1,"oto":1},{"n":"Fretters Piano Service","ott":1,"oto":1},{"n":"Hawaii Gold Buyers Exchange","ott":2,"oto":1},{"n":"Lakeview Garden Center & Landscaping","ott":1,"oto":1},{"n":"United Vapor Barrier & Floors","ott":2,"oto":1}],"Merve (MJ) Brielmann":[{"n":"209 Country Shoppe","ott":3,"oto":1},{"n":"A Total Fire Protection Co","ott":1,"oto":0},{"n":"Aarons Plumbing Inc","ott":2,"oto":0},{"n":"All Purpose Well Drilling","ott":3,"oto":1},{"n":"Armuchee Self Storage","ott":1,"oto":1},{"n":"Axis Doors","ott":1,"oto":1},{"n":"Big Foot Air Quality LLC","ott":2,"oto":1},{"n":"Blue Print Specialties Inc","ott":2,"oto":1},{"n":"Carter Heating & Air","ott":1,"oto":1},{"n":"DRAIN SQUAD NYC INC","ott":1,"oto":1},{"n":"Gabriele Masonry & Waterproofing","ott":1,"oto":0},{"n":"Jeannie Pierce Insurance Agency","ott":1,"oto":1},{"n":"Kuhn's Equipment Repair","ott":3,"oto":1},{"n":"Savannah Bail Bonding","ott":3,"oto":1},{"n":"Sevey Norm Well Drilling Inc","ott":3,"oto":2},{"n":"Shenberg Construction","ott":1,"oto":0},{"n":"Techworx LLC","ott":3,"oto":1},{"n":"Tri County Fuels Inc","ott":3,"oto":2},{"n":"Unique II Worldwide","ott":1,"oto":0},{"n":"Youngrens Inc","ott":1,"oto":0}],"Steven Saunders":[{"n":"5 Stars General Contactor Inc","ott":1,"oto":1},{"n":"A D Sonbert Security Systems Inc","ott":2,"oto":1},{"n":"Aeroclean NWA","ott":1,"oto":1},{"n":"Brians Wow Plumbing","ott":3,"oto":3},{"n":"DJ and Associates","ott":1,"oto":1},{"n":"Dennis Green's Paving","ott":1,"oto":0},{"n":"Farleys Roofing INC","ott":2,"oto":1},{"n":"Finly Family Insurance","ott":3,"oto":3},{"n":"Hoffman & Hoffman","ott":1,"oto":1},{"n":"Island Solar Service","ott":1,"oto":0},{"n":"Michael Lloyd Bail Bonds","ott":1,"oto":0},{"n":"On Demand Crane Service","ott":2,"oto":1},{"n":"Pioneer Overhead Door 3G","ott":1,"oto":1},{"n":"Wiltse Towing LLC","ott":1,"oto":0}],"Tyler Popplewell":[{"n":"B&C Remodeling and Flooring","ott":3,"oto":3},{"n":"GET ER DONE LLC","ott":1,"oto":0},{"n":"P510 Coach","ott":1,"oto":1},{"n":"Roto-Rooter Las Cruces","ott":1,"oto":1},{"n":"Toddler Barrier","ott":1,"oto":1},{"n":"Westerly Paints","ott":1,"oto":1}],"Samuel Frias De Paula":[{"n":"Delcon Electric","ott":1,"oto":1},{"n":"Kathy Bleier Coaching LLC","ott":1,"oto":0}],"Florence Francois Nova":[{"n":"Total Home Renovation","ott":1,"oto":1}],"Stacy Roers":[{"n":"Bradham David Dr","ott":1,"oto":0},{"n":"PHD Bathroom Remodeling","ott":1,"oto":0}],"Victor Abner Moscoso Fernandez":[{"n":"Big's RV Service","ott":1,"oto":0},{"n":"Jay Kent Construction LLC","ott":1,"oto":0},{"n":"Nixa Lawn Service","ott":1,"oto":0},{"n":"Patio Furniture Cushions Inc","ott":1,"oto":0},{"n":"Rain Flow Of Indianapolis","ott":1,"oto":0}],"Ellise Payne":[{"n":"A Hepworth Electrical Pty Ltd","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"AOTEA TIMARU LIMITED","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Aaron Slape, Acupuncturist","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Bayside Quality Furniture Restorations","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Character Cabins","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Concept Fencing MC","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Junction Tyre & Auto Services","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Oslands Independants Carpets","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"SJ Firewood Limited","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]}],"Joseph Guillermo Carmona Garcia":[{"n":"A Beauty In The Beast Pet Grooming","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"A Tri Cities Connection","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Airtight SD","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"All Seasons Equipment","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"American Countertop Experts Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"American West Services","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"BWF Home Solutions","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Benson Chiropractic Center","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Blue Ridge Ear Nose Throat & Plastic Surgery","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Boyd Construction","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Breaking Free Counseling","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"ChillTex LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Cullina Management","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"EZ Sewer Cleaning","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Ed K's Ceramic Tile LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Elevate Home Services","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"EuroStone LLC","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Evergreen Insurance Advisors - Medicare & Health Insurance","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Firestone Concrete Coatings","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Flemingmovingllc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Foreman's Quality Painting Services","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Ganan Landscaping","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Greg Munning CFI","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Guild Mortgage- Stephany Kuennen","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Hammond and Sons Lawn Care and Landscaping","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Handi Built, LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Hicks Trading Station","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Holcomb Concrete Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Homemaker Landscaping","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Idaho Construction Company","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"JR Tree Works","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Janina Elite Medispa","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Jason Diller","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Krueger Plumbing LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Lewis Dean Drapery And Blinds","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Long Island Creative Contracting, Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Longworth Bail Bonds","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"MWC Construction Inc","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Mels Tree Service","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Mokelumne Federal Credit Union","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"NRT","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"New Chapter Senior Living Placement","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Next Level Roofing and Remodeling","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Nickel, Greg & Tamara","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Northshore Dermatology Center","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Paralegals Unlimited, LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Platinum Elite Janitorial","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Postal World","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Price, McCluer & Plachecki","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Proctor's Precision Fence","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Quality Termite and Pest Control, LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Reflecto Signs & Graphics","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Right Print","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Rodriguez Ross A Attorney At Law","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Rogue Lock & Key LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"S & J Residential Roofing","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"SAFE-T CHOICE INC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"SEMO lawn solutions LLC","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Siege Productions","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Southern Auto Paint & Body Shop","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Stars and Pipes Plumbing","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Stevens Concrete","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Storm Drains Hawaii","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Sugarbush Tavern","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"SweetiePumps","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"TD Contractors LLC","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"The Bodhi tree Holistic Health Solutions","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"The Childcare Concierge Agency","d":[{"t":"Email","due":"5/29/2026","ov":true,"nw":false}]},{"n":"Wagner Kuntz & Grabouski","d":[{"t":"Email","due":"5/27/2026","ov":true,"nw":false}]},{"n":"Wealth Warden Partners","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]}],"Lauren Carter":[{"n":"CSA Roofing, Painting & General Contracting","ott":3,"oto":1},{"n":"Caldwell Reubens Drilling Inc","ott":3,"oto":1},{"n":"Children's Corner","ott":3,"oto":1},{"n":"Comfort Family Dentistry","ott":2,"oto":1},{"n":"Dulando Screen & Awning Inc","ott":1,"oto":1},{"n":"Healthy Foot Spa","ott":1,"oto":0},{"n":"Lapcomp Computers","ott":1,"oto":0},{"n":"Lex Plant Farm","ott":2,"oto":1},{"n":"Marketside Chiropractic","ott":2,"oto":1},{"n":"Quest Electric Inc","ott":1,"oto":0},{"n":"Taylor Aution Realty","ott":1,"oto":1},{"n":"Traveltimesawait","d":[{"t":"Highlight Video","due":"5/27/2026","ov":true,"nw":false}]}],"Johnny Cornielle":[{"n":"Brandons Awards & Engraving","ott":1,"oto":1},{"n":"Busseys Flea Market","ott":3,"oto":1},{"n":"Early Years The","ott":1,"oto":0},{"n":"GapArmour","ott":1,"oto":1},{"n":"Mechanical Energy Systems","ott":3,"oto":1},{"n":"Neptune Pool Management","ott":1,"oto":1},{"n":"Network Financial","ott":1,"oto":1},{"n":"PSI Seamless Gutters","ott":1,"oto":1},{"n":"TRUE CRAFT FOUNDATION REPAIR & WATERPROOFING","ott":1,"oto":1},{"n":"Williams Family Medicine","ott":3,"oto":1}],"David Crisler":[{"n":"AK Firewood","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Aqua Dash","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Bendigo Tyre & Auto","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Best of the Best Reblocking","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"H & A Training & Supplies","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"KIWIVAC CENTRAL VACUUM SYSTEMS (1999) LIMITED","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"M1 business system","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Perforge","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"South Seas Construction Ltd","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"The Rose City Limousine","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]},{"n":"Top Mix Construction","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"Tree Limits Pty Ltd","d":[{"t":"Highlight Video","due":"5/28/2026","ov":true,"nw":false}]},{"n":"oxen limited","d":[{"t":"Highlight Video","due":"5/25/2026","ov":true,"nw":false}]}],"Scott Mather":[{"n":"Advantage Life & Health","ott":1,"oto":1},{"n":"Head to Toe Reset Spa","ott":1,"oto":1},{"n":"MVP RIDES","ott":1,"oto":1},{"n":"Magic Refrigeration","ott":1,"oto":1},{"n":"Wilcox Transmission","ott":1,"oto":1}],"Chelsea Dingus":[{"n":"A Fresh Cut Landscaping","ott":1,"oto":0},{"n":"AARO Fence Inc","ott":1,"oto":0},{"n":"Access & Alarm Company Inc","ott":2,"oto":1},{"n":"Bullet Hole Annex","ott":1,"oto":1},{"n":"Corporal Lawn Service","ott":2,"oto":0},{"n":"Dan Green","ott":1,"oto":0},{"n":"GREAT TOUCH BEHAVORAL HEALTH","ott":1,"oto":0},{"n":"In and out garage doors","ott":1,"oto":1},{"n":"Patriot Sunrooms","ott":1,"oto":0},{"n":"Radon Raiders","ott":1,"oto":0},{"n":"Sanderson & De Haan Lawn Sprinkling","ott":1,"oto":0},{"n":"The John Wood Insurance Agency Inc.","ott":1,"oto":0}]};

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
  rows.forEach(r => {
    const raw = r["Touchpoint: Owner Name \u2191"]||r["Touchpoint: Owner Name"]||r["name"]||r["Name"]||"";
    const name = norm(raw.trim());
    const status = (r["Cadence Member: Status"]||r["Status"]||"").trim();
    if (!name||name==="Total"||!isValidCSM(raw.trim())) return;
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
function mapSkipped(rows) {
  const SKIP_OUTCOME = "continued after 4th reschedule";
  const by = {};
  rows.forEach(r => {
    const outcome = (r["Outcome"]||"").trim().toLowerCase();
    if (outcome !== SKIP_OUTCOME) return;
    // Column G in the sheet
    const raw = r["Touchpoint: Owner Name"]||r["Touchpoint: Owner Name \u2191"]||r["Owner Name"]||"";
    const name = norm(raw.trim());
    if (!name || !isValidCSM(raw.trim())) return;
    if (!by[name]) by[name] = {name, count:0, accounts:[]};
    by[name].count++;
    const acct = (r["Account"]||"").trim();
    if (acct && !by[name].accounts.find(a=>a.n===acct)) {
      by[name].accounts.push({n:acct});
    }
  });
  return Object.values(by);
}

// ── BUILD UNIFIED CSM LIST ─────────────────────────────────────────────────
function buildCSMs(rev, email, cad, due, ontime, skipped) {
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
        skippedCount:0, skippedAccts:[]};
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
    c.skippedAccts = d.accounts||[];
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
  let sc=0,sf=0;
  if(cadC.length>0&&avgCad!=null){sc+=Math.min(avgCad/0.9,1)*30;sf++;}
  if(avgOpen!=null){sc+=Math.min(avgOpen/0.7,1)*25;sf++;}
  if(avgOT!=null){sc+=Math.min(avgOT/0.8,1)*20;sf++;}
  if(revPct!=null){sc+=Math.min(revPct/0.7,1)*25;sf++;}
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
  if(csm.skippedCount>0) atts.unshift("🚩 "+csm.skippedCount+" account"+(csm.skippedCount>1?"s":"")+" with Continued After 4th Reschedule: "+csm.skippedAccts.map(a=>a.n).join(", "));
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
  const attn = csms.filter(c=>(c.cadCount>0&&c.cadPct<0.9)||c.skippedCount>0).sort((a,b)=>{
    if(b.skippedCount>0&&!a.skippedCount) return 1;
    if(a.skippedCount>0&&!b.skippedCount) return -1;
    return a.cadPct-b.cadPct;
  });
  const wins = csms.filter(c=>c.cadCount>0&&c.cadPct>=0.9&&c.skippedCount===0).sort((a,b)=>b.cadPct-a.cadPct);
  return (
    <div>
      {/* 🚩 Red flag: Skipped cadences (Continued After 4th Reschedule) */}
      {skipped.length>0&&<div style={{background:"rgba(127,29,29,.06)",border:"1px solid rgba(127,29,29,.3)",borderRadius:12,padding:16,marginBottom:16}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
          <span style={{fontSize:16,fontWeight:600,color:"#7f1d1d"}}>🚩 Skipped Cadences — Continued After 4th Reschedule</span>
          <span style={{fontSize:10,fontWeight:600,padding:"2px 8px",borderRadius:20,background:"rgba(127,29,29,.12)",color:"#7f1d1d"}}>{skipped.length} CSMs · Prior day</span>
        </div>
        <div style={{fontSize:11,color:"#991b1b",marginBottom:10,fontStyle:"italic"}}>High negative impact on score — these accounts have been rescheduled 4+ times without completion</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>
          {skipped.map(c=>{
            const i=lk(c.name)||{};
            return <div key={c.name} style={{background:"#fff",border:"0.5px solid rgba(127,29,29,.2)",borderRadius:10,padding:12,display:"flex",alignItems:"flex-start",gap:10,cursor:"pointer"}} onClick={()=>onSelectCSM(c.name)}>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:12,fontWeight:600,color:"#7f1d1d",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.name}</div>
                <div style={{fontSize:11,color:"#808080",marginTop:2}}>{st(i.t||"")}</div>
                <div style={{marginTop:6,display:"flex",flexWrap:"wrap",gap:4}}>
                  {(c.skippedAccts||[]).map((a,ai)=>(
                    <span key={ai} style={{fontSize:10,padding:"1px 6px",borderRadius:10,background:"rgba(127,29,29,.08)",color:"#7f1d1d",border:"0.5px solid rgba(127,29,29,.15)"}}>{a.n}</span>
                  ))}
                </div>
              </div>
              <div style={{textAlign:"center",flexShrink:0}}>
                <div style={{fontSize:20,fontWeight:700,color:"#7f1d1d"}}>{c.skippedCount}</div>
                <div style={{fontSize:10,color:"#991b1b"}}>skipped</div>
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
              <thead><tr>{["CSM","Team","Skipped","Cadence","On-time","Overdue","Gap"].map((h,j)=><th key={h} style={{fontSize:10,textTransform:"uppercase",color:"#808080",fontWeight:500,padding:"0 8px 8px 0",textAlign:j>=2?"right":"left",borderBottom:"0.5px solid rgba(41,53,93,.08)"}}>{h}</th>)}</tr></thead>
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
function RevenueView({rawRev, csms, filterCoach}) {
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

  // Apply coach + region filter
  const filtered = rows.filter(r=>{
    const i = lk(r.csm);
    if (filterCoach && !(i&&i.c===filterCoach)) return false;
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
            const regRows = rows.filter(r=>r.region===reg && (!filterCoach||(lk(r.csm)&&lk(r.csm).c===filterCoach)));
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
            rows.filter(r=>r.region===reg && (!filterCoach||(lk(r.csm)&&lk(r.csm).c===filterCoach))).forEach(r=>{
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
  const [history, setHistory] = useState([]);
  const [skippedCSMs, setSkippedCSMs] = useState([]);
  const [rawRev, setRawRev] = useState([]);

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
    let latestEmail=[], latestCad=[], latestDue=[], latestOntime=[], latestHistory=[], latestSkipped=[];

    function loadAll() {
      return Promise.all([
        fetchCSV(CSV_REV),
        fetchCSV(CSV_EMAIL),
        fetchCSV(CSV_CAD),
        fetchCSV(CSV_DUE),
        fetchCSV(CSV_ONTIME),
        fetchCSV(CSV_HISTORY).catch(()=>[]),
        fetchCSV(CSV_SKIPPED).catch(()=>[]),
      ]).then(([revRows, emailRows, cadRows, dueRows, ontimeRows, historyRows, skippedRows]) => {
        latestEmail   = emailRows;
        latestCad     = cadRows;
        latestDue     = dueRows;
        latestOntime  = ontimeRows;
        latestHistory = historyRows;
        latestSkipped = skippedRows;
        setRawRev(revRows);
        const rev     = mapRev(revRows);
        const email   = mapEmail(emailRows);
        const cad     = mapCadence(cadRows);
        const due     = mapDue(dueRows);
        const ontime  = mapOnTime(ontimeRows);
        const skipped = mapSkipped(skippedRows);
        const built   = buildCSMs(rev, email, cad, due, ontime, skipped);
        setCSMs(built);
        setSkippedCSMs(built.filter(c=>c.skippedCount>0).sort((a,b)=>b.skippedCount-a.skippedCount));
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
        const built   = buildCSMs(rev, email, cad, due, ontime, skipped);
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
          {["coaching","overview","revenue","leaderboard","activity","trends"].map(t=>(
            <button key={t} onClick={()=>setTab(t)}
              style={{padding:"10px 18px",fontSize:13,fontWeight:500,color:tab===t?"#fff":"rgba(255,255,255,.55)",background:"transparent",border:"none",cursor:"pointer",borderBottom:tab===t?"3px solid #FF5000":"3px solid transparent",whiteSpace:"nowrap"}}>
              {t==="coaching"?"Coaching":t==="trends"?"📈 Trends":t==="revenue"?"💰 Revenue":t.charAt(0).toUpperCase()+t.slice(1)}
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
          {tab==="coaching"&&<CoachingView csms={filteredCSMs} coach={filterCoach} onSelectCSM={selectCSMFn} onSelectCoach={e=>{setFilterCoach(e);setFilterCSM("");}} onClear={()=>{setFilterCoach("");setFilterCSM("");}} skippedCSMs={skippedCSMs.filter(c=>{const i=lk(c.name);if(filterCoach&&(i&&i.c)!==filterCoach)return false;if(filterCSM&&c.name!==filterCSM)return false;return true;})}/>}
          {tab==="overview"&&<OverviewView csms={filteredCSMs} allCSMs={csms}/>}
          {tab==="leaderboard"&&<LeaderboardView csms={filteredCSMs}/>}
          {tab==="activity"&&<ActivityView csms={filteredCSMs}/>}
          {tab==="revenue"&&<RevenueView rawRev={rawRev} csms={filteredCSMs} filterCoach={filterCoach}/>}
          {tab==="trends"&&<TrendsView history={history} csms={csms} filterCoach={filterCoach} filterCSM={filterCSM}/>}
        </div>
      )}
    </div>
  );
}
