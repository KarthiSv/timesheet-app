import React, { useState, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, Legend } from "recharts";

const CSS_LOGIN = "@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;600&family=IBM+Plex+Mono:wght@400;600&display=swap');*{box-sizing:border-box;margin:0;padding:0}";
const CSS_MANAGER = `@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;600&family=IBM+Plex+Mono:wght@400;600&display=swap');
*{box-sizing:border-box;margin:0;padding:0}
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-thumb{background:#8d8d8d}
.emp-row:hover td{background:#dce9ff!important;cursor:pointer}
input[type=checkbox]{accent-color:#0f62fe;width:14px;height:14px;cursor:pointer}

/* ── MOBILE RESPONSIVE ────────────────────────────────────────── */
.mgr-nav-links { display:flex; align-items:center; gap:12px; }
.mgr-nav-menu  { display:none; }
.mgr-sidebar   { display:none; }

.col-hide-mobile { }
@media (max-width: 768px) {
  .col-hide-mobile { display:none !important; }
  /* Nav: hide text labels, show hamburger */
  .mgr-nav-links  { display:none !important; }
  .mgr-nav-menu   { display:flex !important; align-items:center; gap:8px; }
  .mgr-nav-mobile { display:flex !important; flex-direction:column; position:fixed;
    top:48px; left:0; right:0; background:#161616; z-index:300; padding:12px 0;
    border-bottom:1px solid #393939; }
  .mgr-nav-mobile a, .mgr-nav-mobile button.nav-item {
    padding:13px 24px; color:#f4f4f4; font-size:14px; text-align:left;
    background:none; border:none; border-bottom:1px solid #262626; cursor:pointer;
    font-family:inherit; width:100%; }

  /* Header band: stack vertically */
  .hdr-band { padding:14px 16px !important; }
  .hdr-band h1 { font-size:18px !important; }
  .hdr-selectors { flex-direction:column !important; gap:8px !important; width:100%; }
  .hdr-selectors > div { width:100% !important; }
  .hdr-selectors select { width:100% !important; }

  /* Dashboard: single column */
  .dash-charts  { flex-direction:column !important; padding:12px 16px !important; }
  .dash-charts > div { flex:none !important; width:100% !important; }
  .dash-stat-grid { grid-template-columns: repeat(2, 1fr) !important; }
  .dash-variance-bar { padding:12px 16px !important; flex-wrap:wrap; gap:12px !important; }

  /* Records: horizontal scroll with sticky first col */
  .records-table-wrap { margin:0 8px !important; }
  .records-table-wrap table { font-size:12px !important; }
  .records-table-wrap th,
  .records-table-wrap td { padding:8px 7px !important; }

  /* Filter bar: stack */
  .filter-row1 { flex-direction:column !important; gap:8px !important; }
  .filter-row2 { flex-direction:column !important; gap:6px !important; }
  .filter-row2 > div { width:100% !important; }

  /* Bulk bar: wrap */
  .bulk-bar { padding:8px 12px !important; }

  /* Panels: full width */
  .panel-slide { width:100vw !important; }

  /* Import modal: full screen */
  .import-modal { width:100vw !important; height:100vh !important; max-height:100vh !important; border-radius:0 !important; }

  /* General padding reduction */
  .page-pad { padding:12px 16px !important; }
}

@media (max-width: 480px) {
  .dash-stat-grid { grid-template-columns: repeat(2, 1fr) !important; }
  .hdr-band h1 { font-size:16px !important; }
  .records-table-wrap th, .records-table-wrap td { padding:6px 5px !important; font-size:11px !important; }
}
`;
const CSS_USER = "@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;600&family=IBM+Plex+Mono:wght@400;600&display=swap');*{box-sizing:border-box;margin:0;padding:0}::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-thumb{background:#8d8d8d}";
const FF_SANS = "IBM Plex Sans, Helvetica Neue, Arial, sans-serif";
const FF_MONO = "IBM Plex Mono, monospace";

// ─── CONSTANTS ────────────────────────────────────────────────────────────────
const MONTH_NAMES = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const YEARS = [2024,2025,2026];
const PERIODS = [{ label:"Whole Month", value:"WM" },{ label:"Period 1 (1–15)", value:"P1" },{ label:"Period 2 (16–30/31)", value:"P2" }];
const IBM = {
  blue60:"#0f62fe", blue70:"#0043ce", blue10:"#edf5ff", blue20:"#d0e2ff",
  gray100:"#161616", gray90:"#262626", gray80:"#393939", gray70:"#525252",
  gray60:"#6f6f6f", gray50:"#8d8d8d", gray30:"#c6c6c6", gray20:"#e0e0e0", gray10:"#f4f4f4",
  green50:"#24a148", green10:"#defbe6", green20:"#a7f0ba",
  yellow30:"#f1c21b", yellow10:"#fdf6dd", yellow20:"#fcdc00",
  red60:"#da1e28", red10:"#fff1f1", red20:"#ffb3b8",
  orange40:"#ff832b", orange10:"#fff2e8",
  purple60:"#6929c4", purple10:"#f6f2ff",
  teal50:"#009d9a",
};
const CAL_EVENT_TYPES = {
  holiday:  { label:"Holiday",          color:"#6929c4", bg:"#f6f2ff", icon:"🎉" },
  offshore: { label:"Offshore Holiday",  color:"#0043ce", bg:"#d0e2ff", icon:"🌏" },
  shutdown: { label:"Shutdown",          color:"#da1e28", bg:"#fff1f1", icon:"🔒" },
  deadline: { label:"Deadline",          color:"#ff832b", bg:"#fff2e8", icon:"⚠️" },
  event:    { label:"Org Event",         color:"#009d9a", bg:"#d9fbfb", icon:"📅" },
  note:     { label:"Note",              color:"#525252", bg:"#f4f4f4", icon:"📝" },
};
const DAY_TYPE_COLORS = {
  work:    { bg:"#fff",     label:"Work",    color:IBM.blue60 },
  leave:   { bg:"#fff2e8", label:"Leave",   color:IBM.orange40 },
  holiday: { bg:"#f6f2ff", label:"Holiday", color:IBM.purple60 },
  sick:    { bg:"#fff1f1", label:"Sick",    color:IBM.red60 },
  wfh:     { bg:"#defbe6", label:"WFH",     color:IBM.green50 },
};
// Per-project color palette
const PROJ_COLORS = [
  { bg:"#edf5ff", color:"#0043ce", border:"#d0e2ff" },
  { bg:"#defbe6", color:"#0e6027", border:"#a7f0ba" },
  { bg:"#fff2e8", color:"#8a3800", border:"#ffd9bb" },
  { bg:"#f6f2ff", color:"#491d8b", border:"#d4bbff" },
  { bg:"#d9fbfb", color:"#005d5d", border:"#9ef0f0" },
  { bg:"#fdf6dd", color:"#6e4a00", border:"#fcdc00" },
];

// ─── MICROSOFT AZURE AD / MSAL CONFIG ────────────────────────────────────────
// Fill these in after registering your app in Azure AD:
//   https://portal.azure.com → Azure Active Directory → App registrations → New
// CLIENT_ID  : the "Application (client) ID" from your app registration
// TENANT_ID  : your IBM/org "Directory (tenant) ID" — or use "common" for any Microsoft account
// REDIRECT_URI: the URL your app is deployed at (e.g. https://your-app.vercel.app)
var MSAL_CLIENT_ID   = (typeof window!=="undefined"&&window.MSAL_CLIENT_ID)   || "YOUR_CLIENT_ID_HERE";
var MSAL_TENANT_ID   = (typeof window!=="undefined"&&window.MSAL_TENANT_ID)   || "YOUR_TENANT_ID_HERE";
var MSAL_REDIRECT_URI= (typeof window!=="undefined"&&window.MSAL_REDIRECT_URI)|| window.location.origin;

// ─── SUPABASE CONFIG (optional — for shared user role storage) ────────────────
var SUPABASE_URL      = (typeof window!=="undefined"&&window.SUPABASE_URL)      || "";
var SUPABASE_ANON_KEY = (typeof window!=="undefined"&&window.SUPABASE_ANON_KEY) || "";

function isMSALConfigured(){
  return MSAL_CLIENT_ID && MSAL_CLIENT_ID !== "YOUR_CLIENT_ID_HERE";
}

// ─── LOCAL USER STORE (fallback when Supabase not configured) ─────────────────
// Stored in localStorage as "tsm_users" JSON array
function getLocalUsers() {
  try {
    var raw = localStorage.getItem("tsm_users");
    if (raw) return JSON.parse(raw);
  } catch(e) {}
  // Default seed: just the manager account
  return [{
    id: "local-manager",
    username: "manager",
    password_hash: "86829726faca89db8d78e8b072ca7da43b1c301c49bb3ceb10460a52a80fb868",
    role: "manager",
    full_name: "Admin Manager",
    email: "",
    dept: "",
    emp_id: "",
    is_active: true,
    created_at: new Date().toISOString(),
  }];
}
function saveLocalUsers(users) {
  try { localStorage.setItem("tsm_users", JSON.stringify(users)); } catch(e) {}
}
function isSupabaseConfigured() {
  return SUPABASE_URL && SUPABASE_URL.startsWith("https://") && SUPABASE_ANON_KEY.length > 20;
}

// ─── SUPABASE API HELPERS ──────────────────────────────────────────────────────
async function supabaseFetch(path, method, body, token) {
  var headers = {
    "Content-Type": "application/json",
    "apikey": SUPABASE_ANON_KEY,
    "Authorization": "Bearer " + (token || SUPABASE_ANON_KEY),
  };
  var res = await fetch(SUPABASE_URL + "/rest/v1/" + path, {
    method: method || "GET",
    headers: headers,
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    var err = await res.json().catch(function(){ return {}; });
    throw new Error(err.message || err.hint || ("HTTP " + res.status));
  }
  var text = await res.text();
  return text ? JSON.parse(text) : [];
}

// Lookup user by username from Supabase OR localStorage
async function findUser(username) {
  if (isSupabaseConfigured()) {
    var rows = await supabaseFetch(
      "app_users?username=eq." + encodeURIComponent(username.toLowerCase()) + "&is_active=eq.true&select=*",
      "GET"
    );
    return rows[0] || null;
  } else {
    var users = getLocalUsers();
    return users.find(function(u){ return u.username === username.toLowerCase() && u.is_active !== false; }) || null;
  }
}

// Get all users (manager only)
async function getAllUsers() {
  if (isSupabaseConfigured()) {
    return await supabaseFetch("app_users?select=*&order=role.desc,username.asc", "GET");
  } else {
    return getLocalUsers();
  }
}

// Create user
async function createUser(userData) {
  if (isSupabaseConfigured()) {
    return await supabaseFetch("app_users", "POST", userData);
  } else {
    var users = getLocalUsers();
    if (users.find(function(u){ return u.username === userData.username; })) {
      throw new Error("Username already exists");
    }
    var newUser = Object.assign({ id: "local-" + Date.now(), is_active: true, created_at: new Date().toISOString() }, userData);
    users.push(newUser);
    saveLocalUsers(users);
    return newUser;
  }
}

// Update user
async function updateUser(id, updates) {
  if (isSupabaseConfigured()) {
    return await supabaseFetch("app_users?id=eq." + id, "PATCH", updates);
  } else {
    var users = getLocalUsers();
    var idx = users.findIndex(function(u){ return u.id === id; });
    if (idx === -1) throw new Error("User not found");
    users[idx] = Object.assign({}, users[idx], updates, { updated_at: new Date().toISOString() });
    saveLocalUsers(users);
    return users[idx];
  }
}

// Delete (deactivate) user
async function deactivateUser(id) {
  return updateUser(id, { is_active: false });
}


async function hashPassword(username, password) {
  var data = new TextEncoder().encode(username.toLowerCase() + ":" + password);
  var hashBuffer = await window.crypto.subtle.digest("SHA-256", data);
  var hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(function(b){ return b.toString(16).padStart(2,"0"); }).join("");
}
// Rate limiter for login attempts
var loginAttempts = {};
function checkLoginRateLimit(username) {
  var now = Date.now();
  var key = username.toLowerCase();
  if (!loginAttempts[key]) loginAttempts[key] = { count: 0, resetAt: now + 60000 };
  if (now > loginAttempts[key].resetAt) { loginAttempts[key] = { count: 0, resetAt: now + 60000 }; }
  loginAttempts[key].count++;
  return loginAttempts[key].count <= 5; // max 5 attempts per minute
}
function resetLoginAttempts(username) {
  delete loginAttempts[username.toLowerCase()];
}

// ─── HELPERS ─────────────────────────────────────────────────────────────────
function getSeverity(u){const e=Number(u.entered)||0,s=Number(u.scheduled)||0;if(s===0||e===s)return 0;if(e===0)return 4;const g=((s-e)/s)*100;if(g<=10)return 1;if(g<=30)return 2;if(g<=60)return 3;return 4;}
const SEV={0:{label:"Complete",color:IBM.green50,bg:IBM.green10,glow:"#24a14855"},1:{label:"Low",color:"#0e6027",bg:"#d1f5d9",glow:"#24a14833"},2:{label:"Medium",color:"#8e6a00",bg:IBM.yellow10,glow:"#f1c21b44"},3:{label:"High",color:IBM.orange40,bg:IBM.orange10,glow:"#ff832b44"},4:{label:"Critical",color:IBM.red60,bg:IBM.red10,glow:"#da1e2855"}};
function getStatus(u){
  var e=Number(u.entered)||0, s=Number(u.scheduled)||0;
  if(u.dataSource==="Clarity only") return "purple"; // separate category
  if(e===0) return "red";
  if(e<s) return "yellow";
  return "green";
}
const STATUS_META={green:{label:"Complete",color:IBM.green50,bg:IBM.green10,glow:"#24a14855"},yellow:{label:"Mismatch",color:"#8e6a00",bg:IBM.yellow10,glow:"#f1c21b55"},red:{label:"Missing",color:IBM.red60,bg:IBM.red10,glow:"#da1e2855"},purple:{label:"No IBM Schedule",color:IBM.purple60,bg:IBM.purple10,glow:"#6929c433"}};

function getDatesForPeriod(monthName,year,period){
  const mIdx=MONTH_NAMES.indexOf(monthName);
  const dim=new Date(year,mIdx+1,0).getDate();
  let start=1,end=dim;
  if(period==="P1"){start=1;end=15;}else if(period==="P2"){start=16;end=dim;}
  const dates=[];
  for(let d=start;d<=end;d++){const dt=new Date(year,mIdx,d);const dow=dt.getDay();dates.push({day:d,dow,isWeekend:dow===0||dow===6,label:["Sun","Mon","Tue","Wed","Thu","Fri","Sat"][dow]});}
  return dates;
}

// monthKey: unique key per month/year combination
function monthKey(monthName,year){return `${monthName}-${year}`;}

// Empty month entries: { P1:{1:[],2:[],...15:[]}, P2:{16:[],...dim:[]}, periodNotes:{P1:"",P2:""} }
// Each day holds an array of { projectCode, hours, type }
function makeEmptyMonthEntries(monthName,year){
  const mIdx=MONTH_NAMES.indexOf(monthName);
  const dim=new Date(year,mIdx+1,0).getDate();
  const p1={};for(let d=1;d<=15;d++)p1[d]=[];
  const p2={};for(let d=16;d<=dim;d++)p2[d]=[];
  return{P1:p1,P2:p2,periodNotes:{P1:"",P2:""}};
}

// ─── MOCK DATA ────────────────────────────────────────────────────────────────
function makeHistory(projList,totalSched,totalEntered){
  const months=MONTH_NAMES.slice(0,6);const rows=[];
  projList.forEach(p=>{
    const ps=Math.round(totalSched/projList.length),pe=Math.round(totalEntered/projList.length);
    months.forEach((m,mi)=>{["P1","P2"].forEach(period=>{
      const isCurrent=mi===4&&period==="P2";const sch=Math.round(ps/2);let ent;
      if(isCurrent){ent=Math.round(pe/2);}else if(totalEntered===0){ent=Math.random()>0.3?0:Math.round(sch*0.4);}
      else{const v=(Math.random()-0.5)*0.2;ent=Math.max(0,Math.round(sch*(1+v)));if(ent>sch)ent=sch;}
      rows.push({month:m,period,periodLabel:`${m} ${period}`,projectCode:p.code,projectName:p.name,scheduled:sch,entered:ent,diff:sch-ent,isCurrent});
    });});
  });
  return rows;
}

const now_g=new Date();
const BASE_USERS=[
  {id:"E001",name:"Alice Johnson", email:"alice.j@co.com",  dept:"Engineering",resourceManager:"Rachel Green", scheduled:80,entered:80,lastEntry:"2025-06-14",projects:[{code:"PRJ-001",name:"Horizon Platform"},{code:"PRJ-004",name:"Data Lake"}]},
  {id:"E002",name:"Bob Martinez",  email:"bob.m@co.com",    dept:"Design",      resourceManager:"Steve Rogers", scheduled:80,entered:72,lastEntry:"2025-06-13",projects:[{code:"PRJ-002",name:"UX Revamp"}]},
  {id:"E003",name:"Carol White",   email:"carol.w@co.com",  dept:"Engineering", resourceManager:"Rachel Green", scheduled:80,entered:0, lastEntry:null,        projects:[{code:"PRJ-001",name:"Horizon Platform"},{code:"PRJ-005",name:"API Gateway"}]},
  {id:"E004",name:"David Kim",     email:"david.k@co.com",  dept:"HR",          resourceManager:"Nina Patel",   scheduled:80,entered:80,lastEntry:"2025-06-15",projects:[{code:"PRJ-006",name:"HR Portal"}]},
  {id:"E005",name:"Eva Chen",      email:"eva.c@co.com",    dept:"Finance",     resourceManager:"Marcus Webb",  scheduled:80,entered:64,lastEntry:"2025-06-12",projects:[{code:"PRJ-003",name:"Finance Suite"},{code:"PRJ-007",name:"Compliance Engine"}]},
  {id:"E006",name:"Frank Patel",   email:"frank.p@co.com",  dept:"Engineering", resourceManager:"Rachel Green", scheduled:80,entered:0, lastEntry:null,        projects:[{code:"PRJ-001",name:"Horizon Platform"}]},
  {id:"E007",name:"Grace Lee",     email:"grace.l@co.com",  dept:"Marketing",   resourceManager:"Steve Rogers", scheduled:80,entered:80,lastEntry:"2025-06-15",projects:[{code:"PRJ-008",name:"Brand Refresh"},{code:"PRJ-009",name:"Campaign Analytics"}]},
  {id:"E008",name:"Henry Brown",   email:"henry.b@co.com",  dept:"Finance",     resourceManager:"Marcus Webb",  scheduled:80,entered:76,lastEntry:"2025-06-14",projects:[{code:"PRJ-003",name:"Finance Suite"}]},
  {id:"E009",name:"Iris Nakamura", email:"iris.n@co.com",   dept:"Design",      resourceManager:"Steve Rogers", scheduled:80,entered:80,lastEntry:"2025-06-15",projects:[{code:"PRJ-002",name:"UX Revamp"},{code:"PRJ-010",name:"Design System"}]},
  {id:"E010",name:"James Wilson",  email:"james.w@co.com",  dept:"HR",          resourceManager:"Nina Patel",   scheduled:80,entered:48,lastEntry:"2025-06-10",projects:[{code:"PRJ-006",name:"HR Portal"}]},
  {id:"E011",name:"Karen Thomas",  email:"karen.t@co.com",  dept:"Engineering", resourceManager:"Rachel Green", scheduled:80,entered:80,lastEntry:"2025-06-15",projects:[{code:"PRJ-004",name:"Data Lake"},{code:"PRJ-005",name:"API Gateway"}]},
  {id:"E012",name:"Leo Garcia",    email:"leo.g@co.com",    dept:"Marketing",   resourceManager:"Steve Rogers", scheduled:80,entered:0, lastEntry:null,        projects:[{code:"PRJ-008",name:"Brand Refresh"}]},
];
const MOCK_USERS=BASE_USERS.map(u=>({...u,history:makeHistory(u.projects,u.scheduled,u.entered),monthlyEntries:{}}));

function parseExcelRows(rows){const byId=Object.create(null);rows.forEach(r=>{if(!r||typeof r!=="object")return;const id=r.empid||r.id||"";if(!byId[id])byId[id]={id,name:r.name||"",entered:0,scheduled:0};byId[id].entered+=Number(r.entered)||0;byId[id].scheduled+=Number(r.scheduled)||0;});return Object.values(byId);}
function downloadTemplate(){const h=["EmpID","Name","Email","Department","ResourceManager","ProjectCode","ProjectName","ScheduledHours","EnteredHours","LastEntryDate"];const s=[["E001","Alice Johnson","alice.j@co.com","Engineering","Rachel Green","PRJ-001","Horizon Platform",80,80,"2025-06-14"]];const ws=XLSX.utils.aoa_to_sheet([h,...s]);ws["!cols"]=h.map(()=>({wch:20}));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Timesheet_Template");XLSX.writeFile(wb,"Timesheet_Import_Template.xlsx");}
function handleExportFull(records, monthFilter) {
  var month = monthFilter === "All" ? "All Months" : (monthFilter||"All").replace("-"," ");
  var headers = ["IBM Name","BMO/Clarity Name","Data Source","Talent ID","Serial ID","WBS ID","Billing Code","Country","Resource Manager","Workitems","Claim Months","Activity Code","IBM Scheduled Hrs","Clarity Actual Hrs","Variance","Variance %","Timesheet Status","Approved By","Resource Active","Reporting Periods","Notes"];
  var rows = records.map(function(u){
    var sched = Number(u.scheduled)||0;
    var actual = monthFilter==="All" ? Number(u.entered)||0 : ((u.monthlyHours||{})[monthFilter]||Number(u.entered)||0);
    var variance = sched - actual;
    var variancePct = sched > 0 ? Math.round((variance/sched)*100) : (actual>0?-100:0);
    return [
      u.name||"",
      u.clarityName&&u.clarityName!==u.name?u.clarityName:"",
      u.dataSource||"",
      u.talentId||"",
      u.serialId||"",
      u.wbsId||"",
      u.billingCode||"",
      u.country||"",
      u.resourceManager||"",
      (u.projects||[]).map(function(p){return p.name;}).join("; ")||"",
      (u.claimMonths||[]).join(", ")||"",
      u.activityCode||"",
      sched,
      actual,
      variance,
      variancePct+"%",
      u.timesheetStatus||"",
      u.approvedBy||"",
      u.resourceActive||"",
      (u.clarityPeriods||u.periods||[]).join("; ")||"",
      ""
    ];
  });
  var ws = XLSX.utils.aoa_to_sheet([headers].concat(rows));
  // Column widths
  ws["!cols"] = [22,22,14,12,12,16,14,14,20,28,18,14,14,14,10,10,16,18,14,24,20].map(function(w){return{wch:w};});
  // Style header row - make it bold (basic)
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Timesheet_"+month.replace(" ","_"));
  // Sanitize filename - remove any characters that could cause path issues
  var safeMonth = (month||"All").replace(/[^a-zA-Z0-9_\-]/g, "_").slice(0, 50);
  XLSX.writeFile(wb, "Timesheet_Export_" + safeMonth + ".xlsx");
}
function downloadConsolidated(users,mL,pL){const rows=[];users.forEach(u=>{const st=getStatus(u);const base=[u.id,u.name,u.email,u.dept,u.resourceManager,u.scheduled,u.entered,Number(u.scheduled)-Number(u.entered),u.entered>0?"Yes":"No",st==="green"?"Complete":st==="yellow"?"Mismatch":"Missing",SEV[getSeverity(u)].label,u.lastEntry||"N/A"];if(u.projects && projects.length){u.projects.forEach(p=>{rows.push([...base,p.code||"—",p.name||"—"]);});}else rows.push([...base,"—","—"]);});const h=["EmpID","Name","Email","Dept","Resource Mgr","Scheduled Hrs","Entered Hrs","Diff","Data Entered","Status","Severity","Last Entry","Project Code","Project Name"];const ws=XLSX.utils.aoa_to_sheet([h,...rows]);ws["!cols"]=h.map(()=>({wch:18}));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,`${mL} ${pL}`);XLSX.writeFile(wb,`Timesheet_${mL.replace(" ","_")}_Consolidated.xlsx`);}

// ─── NOTIFICATIONS ────────────────────────────────────────────────────────────
function buildNotifTemplate(user,status,mL,pL){
  const today=new Date().toLocaleDateString("en-GB",{day:"2-digit",month:"long",year:"numeric"});
  const projects=(user.projects||[]).map(p=>p.code+" - "+p.name).join(", ")||"-";
  const sch=Number(user.scheduled)||0,ent=Number(user.entered)||0;
  const gap=sch-ent,sev=SEV[getSeverity(user)].label.toUpperCase();
  const div="--------------------------------------";
  const mismatches=(user.history||[]).filter(r=>r.diff>0).sort((a,b)=>b.diff-a.diff).slice(0,8);
  const entStr=ent===0?"Not submitted":ent+"h",gapStr=gap>0?"-"+gap+"h":"None";
  let mmBlock="";
  if(mismatches.length>0){
    const rows=mismatches.map(m=>"  "+(m.periodLabel||"").padEnd(20)+"| "+(m.projectCode||"").padEnd(15)+"| "+(m.scheduled+"h").padEnd(10)+"| "+(m.entered===0?"Not submitted":m.entered+"h").padEnd(9)+"| -"+m.diff+"h").join("\n");
    mmBlock=[div,"  MISMATCH DETAILS",div,"  Period              | Project        | Scheduled | Entered | Gap","  "+"-".repeat(64),rows,div].join("\n")+"\n";
  }
  const action=status==="red"?"Your timesheet has NOT been submitted. Please log in immediately.":"Your timesheet is incomplete. Please review and update your entries.";
  const subject="[Timesheet Reminder] Action Required - "+user.name+" | "+mL+" | "+pL;
  const body=["Dear "+user.name+",","","This is an automated reminder from the Timesheet Management System.","Period: "+mL+" - "+pL,div,"  Employee    : "+user.name,"  Employee ID : "+user.id,"  Severity    : "+sev,div,"  Scheduled   : "+sch+"h","  Entered     : "+entStr,"  Gap         : "+gapStr,"  Projects    : "+projects,div,mmBlock,"ACTION REQUIRED:",action,"","Please resolve by end of business today ("+today+").",div,"Timesheet Management System"].join("\n");
  return{subject,body};
}
function genNotifForUser(u,mL,pL){return buildNotifTemplate(u,getStatus(u),mL,pL);}
function genBulkNotifs(targets,mL,pL,onP){const r={};targets.forEach((u,i)=>{r[u.id]=buildNotifTemplate(u,getStatus(u),mL,pL);onP(i+1,targets.length,u.name);});return r;}

// ─── UI ATOMS ─────────────────────────────────────────────────────────────────
function StatusDot({status,size=12}){const m=STATUS_META[status]||STATUS_META.red;return <span style={{display:"inline-block",width:size,height:size,borderRadius:"50%",background:status==="yellow"?IBM.yellow30:m.color,boxShadow:"0 0 "+Math.round(size/1.4)+"px 2px "+m.glow,flexShrink:0}}/>;}
function SevBadge({sev}){const s=SEV[sev];return <span style={{background:s.bg,color:s.color,padding:"2px 9px",fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.05em",borderRadius:2,display:"inline-flex",alignItems:"center",gap:5,border:`1px solid ${s.color}`}}>{sev===4?"●":sev===3?"▲":sev===2?"◆":sev===1?"▼":"✓"} {s.label}</span>;}
function ProjectChips({projects}){const[exp,setExp]=useState(false);if(!projects && projects.length)return <span style={{color:IBM.gray50,fontSize:12}}>—</span>;const vis=exp?projects:projects.slice(0,1);return <div style={{display:"flex",flexWrap:"wrap",gap:4}}>{vis.map((p,i)=><span key={i} style={{background:"#dde1e7",color:IBM.gray90,fontSize:11,padding:"2px 7px",borderRadius:2,whiteSpace:"nowrap"}}><b>{p.code}</b>{p.name?` · ${p.name}`:""}</span>)}{projects.length>1&&<button onClick={()=>setExp(v=>!v)} style={{fontSize:11,color:IBM.blue60,background:"none",border:"none",cursor:"pointer",padding:"2px 4px"}}>{exp?`▲ less`:`+${projects.length-1} more`}</button>}</div>;}

// ★ FIX: Sel component — always white bg options, no global CSS bleed
function Sel({value,onChange,options,dark=false,style={}}){
  return(
    <select value={value} onChange={onChange}
      style={{background:dark?"#0043ce":"#fff",border:dark?"1px solid #4589ff":`1px solid ${IBM.gray30}`,color:dark?"#fff":IBM.gray100,padding:"7px 30px 7px 10px",fontSize:13,cursor:"pointer",outline:"none",appearance:"none",fontFamily:"inherit",backgroundImage:`url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='10' viewBox='0 0 12 12'%3E%3Cpath fill='${dark?"%23ffffff":"%23525252"}' d='M6 8L1 3h10z'/%3E%3C/svg%3E")`,backgroundRepeat:"no-repeat",backgroundPosition:"right 8px center",...style}}>
      {options.map(o=>{const v=typeof o==="object"?o.value:String(o),l=typeof o==="object"?o.label:String(o);return <option key={v} value={v} style={{background:"#ffffff",color:"#161616"}}>{l}</option>;})}
    </select>
  );
}

// ─── LOGIN SCREEN ─────────────────────────────────────────────────────────────
// ─── MSAL / MICROSOFT SSO ─────────────────────────────────────────────────────
// Uses the MSAL Browser SDK loaded from CDN (added to index.html)
function getMSALApp() {
  if (!isMSALConfigured()) return null;
  if (typeof window.msal === "undefined") return null;
  if (window._msalAppInstance) return window._msalAppInstance;
  var config = {
    auth: {
      clientId:    MSAL_CLIENT_ID,
      authority:   "https://login.microsoftonline.com/" + MSAL_TENANT_ID,
      redirectUri: MSAL_REDIRECT_URI,
    },
    cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false },
  };
  window._msalAppInstance = new window.msal.PublicClientApplication(config);
  return window._msalAppInstance;
}

async function msalLogin() {
  var app = getMSALApp();
  if (!app) throw new Error("MSAL not configured");
  // Try silent first (already signed in)
  var accounts = app.getAllAccounts();
  var scopes = ["openid", "profile", "email", "User.Read"];
  if (accounts.length > 0) {
    try {
      var silent = await app.acquireTokenSilent({ scopes: scopes, account: accounts[0] });
      return silent;
    } catch(e) { /* fall through to popup */ }
  }
  // Popup sign-in
  return await app.loginPopup({ scopes: scopes });
}

async function msalLogout() {
  var app = getMSALApp();
  if (app) {
    var accounts = app.getAllAccounts();
    if (accounts.length > 0) {
      await app.logoutPopup({ account: accounts[0] });
    }
  }
}

// Map Microsoft account to app session
// Roles: look up email in Supabase/localStorage to get manager vs user role
async function resolveRoleForEmail(email) {
  var normalizedEmail = (email||"").toLowerCase().trim();
  // Check local/Supabase user store first
  try {
    var users = await getAllUsers();
    var match = users.find(function(u){
      return (u.email||"").toLowerCase() === normalizedEmail && u.is_active !== false;
    });
    if (match) return { role: match.role, empId: match.emp_id||"", dept: match.dept||"", userId: match.id };
  } catch(e) {}
  // Default: first sign-in = employee role (manager must grant manager role via Users tab)
  return { role: "user", empId: "", dept: "" };
}

function LoginScreen({onLogin}){
  const[un,setUn]=useState("");
  const[pw,setPw]=useState("");
  const[show,setShow]=useState(false);
  const[err,setErr]=useState("");
  const[loading,setLoading]=useState(false);
  const[msalReady,setMsalReady]=useState(false);
  const[mode,setMode]=useState("sso"); // "sso" | "password"

  // Check if MSAL SDK is loaded
  React.useEffect(function(){
    var tries = 0;
    var t = setInterval(function(){
      if (typeof window.msal !== "undefined") { setMsalReady(true); clearInterval(t); }
      if (++tries > 20) clearInterval(t);
    }, 300);
    return function(){ clearInterval(t); };
  }, []);

  // Handle Microsoft SSO login
  function handleMSAL() {
    setErr(""); setLoading(true);
    msalLogin().then(function(result){
      var acc = result.account || (result.idTokenClaims ? { name: result.idTokenClaims.name, username: result.idTokenClaims.preferred_username||result.idTokenClaims.email } : null);
      if (!acc) { setErr("Could not retrieve account info."); setLoading(false); return; }
      var email = acc.username || acc.idTokenClaims && acc.idTokenClaims.email || "";
      var name  = acc.name || email;
      return resolveRoleForEmail(email).then(function(roleInfo){
        setLoading(false);
        onLogin({
          username: email.split("@")[0].toLowerCase(),
          name:     name,
          email:    email,
          role:     roleInfo.role,
          empId:    roleInfo.empId,
          dept:     roleInfo.dept,
          userId:   roleInfo.userId||null,
          msalAccount: acc,
        });
      });
    }).catch(function(ex){
      setLoading(false);
      var msg = ex.message||"";
      if (msg.indexOf("user_cancelled")!==-1||msg.indexOf("cancel")!==-1) { setErr(""); return; }
      setErr("Microsoft login failed: " + (msg||"Please try again."));
    });
  }

  // Handle username/password login (fallback)
  function handlePassword() {
    var username = un.trim().toLowerCase();
    if(!username||!pw){setErr("Please enter username and password.");return;}
    if(!checkLoginRateLimit(username)){setErr("Too many attempts. Please wait 1 minute.");return;}
    setLoading(true);
    hashPassword(username,pw).then(function(hashed){
      return findUser(username).then(function(user){
        setLoading(false);
        if(!user||user.password_hash!==hashed){setErr("Invalid username or password.");return;}
        if(user.is_active===false){setErr("Account is disabled. Contact your manager.");return;}
        resetLoginAttempts(username);
        setErr("");
        onLogin({username:username,role:user.role,empId:user.emp_id||null,name:user.full_name||null,email:user.email||null,dept:user.dept||null,userId:user.id});
      });
    }).catch(function(ex){
      setLoading(false);
      setErr("Login error: "+(ex.message||"Please try again."));
    });
  }

  var msalConfigured = isMSALConfigured();

  return(
    <div style={{minHeight:"100vh",background:"#161616",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:FF_SANS}}>
      <style>{CSS_LOGIN}</style>
      <div style={{width:"min(440px,96vw)"}}>
        {/* Logo */}
        <div style={{textAlign:"center",marginBottom:28}}>
          <div style={{fontSize:36,fontWeight:700,color:IBM.blue60,fontFamily:FF_MONO,letterSpacing:"-2px"}}>IBM</div>
          <div style={{fontSize:13,color:IBM.gray30,marginTop:6,letterSpacing:"0.15em",textTransform:"uppercase"}}>Timesheet Management</div>
        </div>

        <div style={{background:"#262626",border:"1px solid #393939"}}>
          <div style={{background:IBM.blue60,padding:"18px 28px"}}>
            <div style={{fontSize:17,fontWeight:600,color:"#fff"}}>Sign In</div>
            <div style={{fontSize:12,color:"#a6c8ff",marginTop:3}}>
              {msalConfigured ? "Use your IBM Microsoft account" : "Enter your credentials"}
            </div>
          </div>

          <div style={{padding:"28px"}}>
            {/* Microsoft SSO Button */}
            {msalConfigured && mode==="sso" && (
              <React.Fragment>
                <button onClick={handleMSAL} disabled={loading||!msalReady}
                  style={{width:"100%",padding:"13px 16px",background:loading?"#444":"#2f2f2f",color:"#fff",border:"1px solid #555",cursor:loading||!msalReady?"not-allowed":"pointer",fontSize:14,fontWeight:600,display:"flex",alignItems:"center",justifyContent:"center",gap:12,marginBottom:16,fontFamily:FF_SANS}}>
                  {/* Microsoft logo SVG */}
                  {!loading && (
                    <svg width="20" height="20" viewBox="0 0 21 21" xmlns="http://www.w3.org/2000/svg">
                      <rect x="1" y="1" width="9" height="9" fill="#f25022"/>
                      <rect x="11" y="1" width="9" height="9" fill="#7fba00"/>
                      <rect x="1" y="11" width="9" height="9" fill="#00a4ef"/>
                      <rect x="11" y="11" width="9" height="9" fill="#ffb900"/>
                    </svg>
                  )}
                  {loading ? "Signing in…" : !msalReady ? "Loading…" : "Sign in with Microsoft"}
                </button>

                <div style={{textAlign:"center",marginBottom:16}}>
                  <span style={{fontSize:11,color:IBM.gray60}}>or </span>
                  <button onClick={function(){setMode("password");setErr("");}} style={{background:"none",border:"none",color:IBM.blue60,cursor:"pointer",fontSize:11,textDecoration:"underline",padding:0}}>
                    use username &amp; password
                  </button>
                </div>
              </React.Fragment>
            )}

            {/* Username/password form */}
            {(!msalConfigured || mode==="password") && (
              <React.Fragment>
                {msalConfigured && (
                  <div style={{marginBottom:16,display:"flex",alignItems:"center",gap:8}}>
                    <button onClick={function(){setMode("sso");setErr("");}} style={{background:"none",border:"none",color:IBM.blue60,cursor:"pointer",fontSize:12,textDecoration:"underline",padding:0}}>
                      &#8592; Back to Microsoft login
                    </button>
                  </div>
                )}
                <div style={{marginBottom:14}}>
                  <label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.07em",color:IBM.gray30,display:"block",marginBottom:6}}>Username</label>
                  <input value={un} onChange={function(e){setUn(e.target.value);setErr("");}} onKeyDown={function(e){if(e.key==="Enter")handlePassword();}}
                    placeholder="e.g. alice.j or manager"
                    style={{width:"100%",padding:"10px 12px",background:"#393939",border:"1px solid "+(err?"#da1e28":"#525252"),color:"#fff",fontSize:14,outline:"none",fontFamily:"inherit"}}/>
                </div>
                <div style={{marginBottom:20}}>
                  <label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.07em",color:IBM.gray30,display:"block",marginBottom:6}}>Password</label>
                  <div style={{position:"relative"}}>
                    <input type={show?"text":"password"} value={pw} onChange={function(e){setPw(e.target.value);setErr("");}} onKeyDown={function(e){if(e.key==="Enter")handlePassword();}}
                      placeholder="Enter password"
                      style={{width:"100%",padding:"10px 40px 10px 12px",background:"#393939",border:"1px solid "+(err?"#da1e28":"#525252"),color:"#fff",fontSize:14,outline:"none",fontFamily:"inherit"}}/>
                    <button onClick={function(){setShow(function(v){return !v;});}} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",color:IBM.gray50,cursor:"pointer",fontSize:12}}>{show?"Hide":"Show"}</button>
                  </div>
                </div>
                <button onClick={handlePassword} disabled={loading}
                  style={{width:"100%",padding:"12px",background:loading?IBM.gray50:IBM.blue60,color:"#fff",border:"none",cursor:loading?"not-allowed":"pointer",fontSize:14,fontWeight:600}}>
                  {loading?"Signing in…":"Sign In →"}
                </button>
              </React.Fragment>
            )}

            {err&&<div style={{background:"#3d1a1a",border:"1px solid #da1e28",color:"#ff8389",padding:"10px 14px",fontSize:13,marginTop:14}}>&#9888; {err}</div>}
          </div>
        </div>

        {/* Setup notice */}
        {!msalConfigured && (
          <div style={{marginTop:14,background:"#1c2a1c",border:"1px solid #393939",padding:"12px 16px",fontSize:11,color:IBM.gray50,lineHeight:1.6}}>
            <b style={{color:IBM.orange40}}>&#9888; Microsoft SSO not configured.</b> Using username/password mode.
            See SECURITY.md for how to connect Azure AD.
          </div>
        )}
        {!msalConfigured && (
          <div style={{marginTop:10,background:"#1c1c1c",border:"1px solid #393939",padding:"12px 16px"}}>
            <div style={{fontSize:10,fontWeight:600,color:IBM.gray50,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:8}}>Demo Accounts</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:5}}>
              {[["manager","manager123","Manager"],["alice.j","alice123","Employee"]].map(function(item){
                return (
                  <button key={item[0]} onClick={function(){setUn(item[0]);setPw(item[1]);setErr("");setMode("password");}}
                    style={{background:"#262626",border:"1px solid #393939",color:IBM.gray30,padding:"7px 10px",cursor:"pointer",textAlign:"left",fontSize:11}}>
                    <b style={{color:item[2].startsWith("M")?IBM.blue60:IBM.gray20}}>{item[0]}</b>
                    <div style={{color:IBM.gray60,fontSize:10,marginTop:2}}>{item[2]} · {item[1]}</div>
                  </button>
                );
              })}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── TIMESHEET ENTRY GRID ─────────────────────────────────────────────────────
// Each day holds an array of rows: [{projectCode, hours, type}]
// Employee can add multiple project rows per day
function TimesheetEntryGrid({userProjects,period,monthName,year,periodData,onUpdate,periodNote,onUpdateNote,calendarEvents,isManager}){
  const dates=useMemo(()=>getDatesForPeriod(monthName,year,period),[monthName,year,period]);

  const projColorMap={};
  userProjects.forEach((p,i)=>{projColorMap[p.code]=PROJ_COLORS[i%PROJ_COLORS.length];});

  const getCalEvt=day=>{
    const key=`${year}-${String(MONTH_NAMES.indexOf(monthName)+1).padStart(2,"0")}-${String(day).padStart(2,"0")}`;
    return(calendarEvents||[]).find(e=>e.date===key)||null;
  };

  const dayTotal=day=>(periodData[day]||[]).reduce((s,r)=>s+(parseFloat(r.hours)||0),0);
  const grandTotal=dates.reduce((s,d)=>s+dayTotal(d.day),0);
  const workDays=dates.filter(d=>!d.isWeekend).length;

  const updateRow=(day,ri,field,val)=>{
    const rows=(periodData[day]||[]).map((r,i)=>i===ri?{...r,[field]:val}:r);
    onUpdate({...periodData,[day]:rows});
  };
  const addRow=day=>{
    const existing=periodData[day]||[];
    const used=new Set(existing.map(r=>r.projectCode));
    const next=userProjects.find(p=>!used.has(p.code))||userProjects[0];
    if(!next)return;
    onUpdate({...periodData,[day]:[...existing,{projectCode:next.code,hours:"",type:"work"}]});
  };
  const removeRow=(day,ri)=>{
    const rows=(periodData[day]||[]).filter((_,i)=>i!==ri);
    onUpdate({...periodData,[day]:rows});
  };

  return(
    <div>
      {/* Summary bar */}
      <div style={{display:"flex",gap:14,alignItems:"center",flexWrap:"wrap",padding:"10px 16px",background:"#fff",border:`1px solid ${IBM.gray20}`,marginBottom:10}}>
        <span style={{fontSize:12,fontWeight:600,color:IBM.gray70}}>Period: <b style={{color:IBM.gray100}}>{period==="P1"?"1 – 15":"16 – End"} {monthName} {year}</b></span>
        <span style={{width:1,height:14,background:IBM.gray20,flexShrink:0}}/>
        <span style={{fontSize:12,color:IBM.gray70}}>Total: <b style={{color:IBM.blue60}}>{grandTotal}h</b></span>
        <span style={{width:1,height:14,background:IBM.gray20,flexShrink:0}}/>
        <span style={{fontSize:12,color:IBM.gray70}}>Work days: <b>{workDays}</b></span>
        <span style={{width:1,height:14,background:IBM.gray20,flexShrink:0}}/>
        <span style={{fontSize:12,color:IBM.gray70}}>Expected: <b>{workDays*8}h</b></span>
        <span style={{marginLeft:"auto",fontSize:12,fontWeight:700,color:grandTotal===0?IBM.gray50:grandTotal>=(workDays*8)?IBM.green50:IBM.orange40}}>
          {grandTotal===0?"Not started":grandTotal>=(workDays*8)?"✓ Complete":"⚠ Incomplete"}
        </span>
      </div>

      {/* Project colour legend */}
      <div style={{display:"flex",gap:6,marginBottom:10,flexWrap:"wrap",alignItems:"center"}}>
        {userProjects.map((p,i)=>{const c=PROJ_COLORS[i%PROJ_COLORS.length];return(
          <span key={p.code} style={{fontSize:11,padding:"3px 10px",background:c.bg,color:c.color,border:`1px solid ${c.border}`,fontWeight:700,borderRadius:2}}>
            {p.code} <span style={{fontWeight:400,opacity:.8}}>— {p.name}</span>
          </span>
        );})}
        {!isManager&&userProjects.length>1&&<span style={{fontSize:11,color:IBM.gray50}}>· Click <b>+ Add</b> to split hours across projects on the same day</span>}
      </div>

      {/* GRID */}
      <div style={{overflowX:"auto",border:`1px solid ${IBM.gray20}`}}>
        <table style={{borderCollapse:"collapse",fontSize:12,minWidth:dates.length*84+170}}>
          <thead>
            <tr>
              <th style={{padding:"8px 14px",background:IBM.gray100,color:"#fff",textAlign:"left",fontWeight:500,fontSize:11,minWidth:170,position:"sticky",left:0,zIndex:3,borderRight:`1px solid ${IBM.gray80}`}}>DATE</th>
              {dates.map(d=>{
                const evt=getCalEvt(d.day);const es=evt?CAL_EVENT_TYPES[evt.type]:null;
                return(
                  <th key={d.day} style={{padding:"5px 3px",textAlign:"center",minWidth:84,
                    background:es?es.bg+"55":d.isWeekend?"#1a1a1a":IBM.gray100,
                    borderLeft:`1px solid ${IBM.gray80}`,borderBottom:`3px solid ${es?es.color:d.isWeekend?IBM.gray80:IBM.blue60}`}}>
                    <div style={{fontSize:10,color:d.isWeekend?"#666":"#a6c8ff"}}>{d.label}</div>
                    <div style={{fontSize:15,fontWeight:700,color:d.isWeekend?"#666":es?es.color:"#fff",lineHeight:1.2}}>{d.day}</div>
                    {es&&<div style={{fontSize:11,marginTop:1}}>{es.icon}</div>}
                    {!es&&d.isWeekend&&<div style={{fontSize:9,color:"#555"}}>OFF</div>}
                  </th>
                );
              })}
              <th style={{padding:"8px 10px",background:IBM.gray100,color:"#fff",textAlign:"center",fontWeight:600,fontSize:12,minWidth:64,borderLeft:`1px solid ${IBM.gray80}`}}>TOTAL</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td style={{padding:"10px 14px",background:IBM.blue10,fontWeight:600,fontSize:11,color:IBM.blue70,position:"sticky",left:0,zIndex:1,borderRight:`1px solid ${IBM.blue20}`,borderBottom:`1px solid ${IBM.blue20}`,verticalAlign:"top"}}>
                HOURS / PROJECT
              </td>
              {dates.map(d=>{
                const rows=periodData[d.day]||[];
                const dt=dayTotal(d.day);
                const evt=getCalEvt(d.day);const es=evt?CAL_EVENT_TYPES[evt.type]:null;
                const cellBg=es?es.bg:d.isWeekend?IBM.gray10:"#fff";

                if(d.isWeekend){
                  return <td key={d.day} style={{background:cellBg,borderLeft:`1px solid ${IBM.gray20}`,borderBottom:`1px solid ${IBM.gray20}`,textAlign:"center",verticalAlign:"middle",padding:8}}>
                    <span style={{fontSize:11,color:IBM.gray30}}>—</span>
                  </td>;
                }
                return(
                  <td key={d.day} style={{background:cellBg,borderLeft:`1px solid ${IBM.gray20}`,borderBottom:`1px solid ${IBM.gray20}`,padding:4,verticalAlign:"top",minWidth:84}}>
                    {/* Empty state */}
                    {rows.length===0&&!isManager&&(
                      <div style={{textAlign:"center",padding:"8px 0"}}>
                        <button onClick={()=>addRow(d.day)} style={{background:"none",border:`1px dashed ${IBM.gray30}`,color:IBM.gray50,cursor:"pointer",fontSize:11,padding:"5px 8px",width:"100%",borderRadius:2}}>+ Add</button>
                      </div>
                    )}
                    {rows.length===0&&isManager&&(
                      <div style={{textAlign:"center",padding:10,color:IBM.gray30,fontSize:12}}>—</div>
                    )}
                    {/* Project rows */}
                    {rows.map((row,ri)=>{
                      const c=projColorMap[row.projectCode]||PROJ_COLORS[0];
                      const tc=DAY_TYPE_COLORS[row.type]||DAY_TYPE_COLORS.work;
                      return(
                        <div key={ri} style={{marginBottom:ri<rows.length-1?3:0,background:c.bg,border:`1px solid ${c.border}`,padding:"3px 4px",borderRadius:2}}>
                          {/* Project selector */}
                          {!isManager?(
                            <select value={row.projectCode} onChange={e=>updateRow(d.day,ri,"projectCode",e.target.value)}
                              style={{width:"100%",padding:"2px 3px",border:"none",background:"transparent",fontSize:10,color:c.color,fontWeight:700,outline:"none",cursor:"pointer",fontFamily:"inherit",marginBottom:2}}>
                              {userProjects.map(p=><option key={p.code} value={p.code} style={{background:"#fff",color:"#161616"}}>{p.code}</option>)}
                            </select>
                          ):(
                            <div style={{fontSize:10,fontWeight:700,color:c.color,marginBottom:2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{row.projectCode}</div>
                          )}
                          {/* Hours */}
                          {isManager?(
                            <div style={{fontSize:14,fontWeight:700,color:IBM.gray100,textAlign:"center",padding:"2px 0"}}>{row.hours||"—"}</div>
                          ):(
                            <input type="number" min="0" max="24" step="0.5" value={row.hours}
                              onChange={e=>updateRow(d.day,ri,"hours",e.target.value)}
                              placeholder="0h"
                              style={{width:"100%",padding:"3px 4px",border:`1px solid ${row.hours?c.color:IBM.gray20}`,textAlign:"center",fontSize:13,outline:"none",fontWeight:row.hours?"700":"400",background:"#fff",color:IBM.gray100,fontFamily:"inherit",borderRadius:1}}/>
                          )}
                          {/* Day type */}
                          {!isManager?(
                            <select value={row.type} onChange={e=>updateRow(d.day,ri,"type",e.target.value)}
                              style={{width:"100%",padding:"2px 3px",border:"none",background:"transparent",fontSize:9,color:tc.color,fontWeight:700,outline:"none",cursor:"pointer",marginTop:2,fontFamily:"inherit"}}>
                              {Object.entries(DAY_TYPE_COLORS).map(([k,v])=><option key={k} value={k} style={{background:"#fff",color:"#161616"}}>{v.label}</option>)}
                            </select>
                          ):(
                            <div style={{fontSize:9,color:tc.color,fontWeight:700,marginTop:2}}>{tc.label}</div>
                          )}
                          {/* Remove row */}
                          {!isManager&&rows.length>1&&(
                            <button onClick={()=>removeRow(d.day,ri)} style={{width:"100%",background:"none",border:"none",color:IBM.red60,cursor:"pointer",fontSize:9,padding:"1px 0",marginTop:1}}>✕ remove</button>
                          )}
                        </div>
                      );
                    })}
                    {/* Add another project */}
                    {!isManager&&rows.length>0&&rows.length<userProjects.length&&(
                      <button onClick={()=>addRow(d.day)} style={{width:"100%",background:"none",border:`1px dashed ${IBM.gray30}`,color:IBM.gray50,cursor:"pointer",fontSize:9,padding:"2px 0",marginTop:3,borderRadius:2}}>+ project</button>
                    )}
                    {/* Day total */}
                    {dt>0&&(
                      <div style={{textAlign:"center",fontSize:11,fontWeight:700,color:IBM.blue70,marginTop:3,borderTop:`1px solid ${IBM.blue20}`,paddingTop:2}}>{dt}h</div>
                    )}
                  </td>
                );
              })}
              <td style={{padding:"10px 8px",background:IBM.blue10,fontWeight:700,fontSize:16,color:IBM.blue60,borderLeft:`1px solid ${IBM.blue20}`,borderBottom:`1px solid ${IBM.blue20}`,textAlign:"center",verticalAlign:"top"}}>{grandTotal}h</td>
            </tr>

            {/* Org events info row */}
            {dates.some(d=>getCalEvt(d.day))&&(
              <tr>
                <td style={{padding:"6px 14px",background:"#f6f2ff",fontWeight:600,fontSize:11,color:IBM.purple60,position:"sticky",left:0,zIndex:1,borderRight:`1px solid #d4bbff`,borderBottom:`1px solid #d4bbff`}}>ORG EVENTS</td>
                {dates.map(d=>{const evt=getCalEvt(d.day);const es=evt?CAL_EVENT_TYPES[evt.type]:null;return(
                  <td key={d.day} style={{padding:"4px 3px",textAlign:"center",background:es?es.bg:d.isWeekend?IBM.gray10:"#fff",borderLeft:`1px solid #d4bbff`,borderBottom:`1px solid #d4bbff`}}>
                    {evt&&<span style={{fontSize:9,color:es.color,fontWeight:700,display:"block",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:80,margin:"0 auto"}} title={evt.label}>{es.icon} {evt.label.length>9?evt.label.slice(0,9)+"…":evt.label}</span>}
                  </td>
                );})}
                <td style={{background:"#f6f2ff",borderLeft:`1px solid #d4bbff`,borderBottom:`1px solid #d4bbff`}}/>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* PERIOD NOTES — single textarea at bottom */}
      <div style={{marginTop:12,background:"#fff",border:`1px solid ${IBM.gray20}`}}>
        <div style={{padding:"9px 16px",background:IBM.gray10,borderBottom:`1px solid ${IBM.gray20}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span style={{fontSize:12,fontWeight:600,color:IBM.gray70}}>
            💬 Period Notes
            {!isManager&&<span style={{fontSize:11,fontWeight:400,color:IBM.gray50,marginLeft:8}}>e.g. "Leave on days 5 and 6", "WFH on 10, 12"</span>}
          </span>
          {periodNote&&<span style={{fontSize:11,color:IBM.green50,fontWeight:600}}>✓ Noted</span>}
        </div>
        <div style={{padding:"10px 16px"}}>
          {isManager?(
            <div style={{fontSize:13,color:periodNote?IBM.gray80:IBM.gray40,fontStyle:periodNote?"normal":"italic",lineHeight:1.6,minHeight:30}}>
              {periodNote||"No notes from employee for this period."}
            </div>
          ):(
            <textarea value={periodNote||""} onChange={e=>onUpdateNote&&onUpdateNote(e.target.value)} rows={2}
              placeholder={`Add any notes for your manager — e.g. "Leave on days 5 and 6", "WFH on 10 and 11 — client week"`}
              style={{width:"100%",padding:"9px 12px",border:`1px solid ${IBM.gray30}`,fontSize:13,resize:"vertical",fontFamily:"inherit",outline:"none",color:IBM.gray100,lineHeight:1.6}}/>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── SUBMISSION HISTORY ───────────────────────────────────────────────────────
// Grouped by month → period → projects, card layout, easy to read
function SubmissionHistory({history}){
  const[expanded,setExpanded]=useState(new Set(["January","February","March","April","May","June"].slice(-2)));
  if(!history && history.length)return null;

  const byMonth={};
  history.forEach(r=>{
    if(!byMonth[r.month])byMonth[r.month]={month:r.month,periods:{}};
    if(!byMonth[r.month].periods[r.period])byMonth[r.month].periods[r.period]={period:r.period,projects:[]};
    byMonth[r.month].periods[r.period].projects.push(r);
  });
  const months=Object.values(byMonth).slice(0,6);

  return(
    <div style={{padding:"0 28px 40px"}}>
      <div style={{fontSize:12,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.07em",color:IBM.gray70,borderBottom:`2px solid ${IBM.blue60}`,paddingBottom:6,marginBottom:14}}>Submission History</div>
      <div style={{display:"flex",flexDirection:"column",gap:8}}>
        {months.map(m=>{
          const periods=Object.values(m.periods);
          const allOk=periods.every(p=>p.projects.every(r=>r.entered>0&&r.diff===0));
          const anyMiss=periods.some(p=>p.projects.some(r=>r.entered===0));
          const sc=allOk?IBM.green50:anyMiss?IBM.red60:IBM.orange40;
          const isOpen=expanded.has(m.month);
          return(
            <div key={m.month} style={{background:"#fff",border:`1px solid ${IBM.gray20}`,borderLeft:`4px solid ${sc}`}}>
              {/* Header row — clickable to expand */}
              <div style={{padding:"10px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",cursor:"pointer",userSelect:"none"}}
                onClick={()=>setExpanded(prev=>{const n=new Set(prev);n.has(m.month)?n.delete(m.month):n.add(m.month);return n;})}>
                <div style={{display:"flex",alignItems:"center",gap:14}}>
                  <span style={{fontSize:14,fontWeight:700,color:IBM.gray100}}>{m.month}</span>
                  {periods.map(p=>{
                    const tot=p.projects.reduce((s,r)=>s+r.entered,0);
                    const sch=p.projects.reduce((s,r)=>s+r.scheduled,0);
                    const ok=tot===sch&&sch>0;
                    return <span key={p.period} style={{fontSize:11,padding:"2px 8px",background:ok?IBM.green10:tot===0?IBM.red10:IBM.yellow10,color:ok?IBM.green50:tot===0?IBM.red60:"#8e6a00",border:`1px solid ${ok?IBM.green20:tot===0?IBM.red20:IBM.yellow20}`,fontWeight:600,borderRadius:2}}>{p.period==="P1"?"P1":"P2"}: {tot===0?"—":tot+"h"}/{sch+"h"}</span>;
                  })}
                </div>
                <div style={{display:"flex",alignItems:"center",gap:10}}>
                  <span style={{fontSize:11,fontWeight:700,color:sc}}>{allOk?"✓ Complete":anyMiss?"⚠ Missing":"⚠ Partial"}</span>
                  <span style={{fontSize:13,color:IBM.gray50}}>{isOpen?"▲":"▼"}</span>
                </div>
              </div>

              {/* Detail rows */}
              {isOpen&&(
                <div style={{borderTop:`1px solid ${IBM.gray20}`,display:"flex",flexWrap:"wrap"}}>
                  {periods.map((p,pi)=>(
                    <div key={p.period} style={{flex:"1 1 50%",padding:"12px 16px",borderRight:pi<periods.length-1?`1px solid ${IBM.gray20}`:"none",minWidth:220}}>
                      <div style={{fontSize:11,fontWeight:700,background:p.period==="P1"?IBM.blue10:"#f6f2ff",color:p.period==="P1"?IBM.blue60:IBM.purple60,padding:"3px 10px",display:"inline-block",borderRadius:2,marginBottom:8}}>{p.period==="P1"?"Period 1 — Days 1–15":"Period 2 — Days 16–End"}</div>
                      <div style={{display:"flex",flexDirection:"column",gap:5}}>
                        {p.projects.map((r,ri)=>{
                          const c=PROJ_COLORS[ri%PROJ_COLORS.length];
                          const ok=r.entered>0&&r.diff===0;
                          const miss=r.entered===0;
                          return(
                            <div key={ri} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 10px",background:ok?IBM.green10:miss?IBM.red10:IBM.yellow10,border:`1px solid ${ok?IBM.green20:miss?IBM.red20:IBM.yellow20}`,borderRadius:2}}>
                              <span style={{fontSize:10,fontWeight:700,background:c.bg,color:c.color,padding:"1px 6px",border:`1px solid ${c.border}`,borderRadius:2,whiteSpace:"nowrap"}}>{r.projectCode}</span>
                              <span style={{fontSize:11,color:IBM.gray60,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.projectName}</span>
                              <span style={{fontSize:12,fontWeight:700,color:ok?IBM.green50:miss?IBM.red60:IBM.orange40,whiteSpace:"nowrap"}}>{miss?"—":r.entered+"h"}</span>
                              {r.diff>0&&<span style={{fontSize:10,color:IBM.red60,whiteSpace:"nowrap"}}>({r.scheduled}h sched.)</span>}
                              {ok&&<span style={{fontSize:11,color:IBM.green50}}>✓</span>}
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}
// ─── USER TIMESHEET VIEW ──────────────────────────────────────────────────────
function UserTimesheetView({session,users,setUsers,calendarEvents,showToast}){
  const now=new Date();
  const[selMonth,setSelMonth]=useState(MONTH_NAMES[now.getMonth()]);
  const[selYear,setSelYear]=useState(now.getFullYear());
  const[activePeriod,setActivePeriod]=useState("P1");
  const[saved,setSaved]=useState(false);

  const currentUser=users.find(u=>u.id===session.empId)||{};
  const mk=monthKey(selMonth,selYear);

  // Always read live from users state
  const monthData=(currentUser.monthlyEntries||{})[mk]||makeEmptyMonthEntries(selMonth,selYear);

  const updateMonthData=patch=>{
    setUsers(prev=>prev.map(u=>{
      if(u.id!==currentUser.id)return u;
      const existing=(u.monthlyEntries||{})[mk]||makeEmptyMonthEntries(selMonth,selYear);
      return{...u,monthlyEntries:{...(u.monthlyEntries||{}),[mk]:{...existing,...patch}}};
    }));
    setSaved(false);
  };

  const handleUpdatePeriod=(period,newData)=>updateMonthData({[period]:newData});

  const handleUpdateNote=(period,val)=>{
    setUsers(prev=>prev.map(u=>{
      if(u.id!==currentUser.id)return u;
      const existing=(u.monthlyEntries||{})[mk]||makeEmptyMonthEntries(selMonth,selYear);
      return{...u,monthlyEntries:{...(u.monthlyEntries||{}),[mk]:{...existing,periodNotes:{...(existing.periodNotes||{P1:"",P2:""}),[period]:val}}}};
    }));
    setSaved(false);
  };

  const handleSave=()=>{
    let total=0;
    ["P1","P2"].forEach(p=>{
      const pd=monthData[p]||{};
      Object.values(pd).forEach(rows=>rows.forEach(r=>total+=parseFloat(r.hours)||0));
    });
    setUsers(prev=>prev.map(u=>u.id!==currentUser.id?u:{...u,entered:total,lastEntry:new Date().toISOString().split("T")[0]}));
    setSaved(true);
    showToast(`✓ Timesheet saved — ${total}h recorded`);
  };

  const totalFor=period=>{
    const pd=monthData[period]||{};
    return Object.values(pd).reduce((s,rows)=>s+rows.reduce((ss,r)=>ss+(parseFloat(r.hours)||0),0),0);
  };
  const totalP1=useMemo(()=>totalFor("P1"),[monthData]);
  const totalP2=useMemo(()=>totalFor("P2"),[monthData]);

  return(
    <div>
      {/* Blue header */}
      <div style={{background:IBM.blue60,padding:"20px 28px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",flexWrap:"wrap",gap:12}}>
          <div>
            <h1 style={{fontSize:22,fontWeight:300,color:"#fff",margin:0}}>My Timesheet</h1>
            <p style={{fontSize:13,color:"#a6c8ff",marginTop:4}}>{currentUser.name} · {currentUser.id} · {currentUser.dept}</p>
          </div>
          <div className="hdr-selectors" style={{display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap"}}>
            <div><label style={{fontSize:10,color:"#a6c8ff",textTransform:"uppercase",letterSpacing:"0.07em",display:"block",marginBottom:4}}>Month</label><Sel dark value={selMonth} onChange={e=>setSelMonth(e.target.value)} options={MONTH_NAMES}/></div>
            <div><label style={{fontSize:10,color:"#a6c8ff",textTransform:"uppercase",letterSpacing:"0.07em",display:"block",marginBottom:4}}>Year</label><Sel dark value={selYear} onChange={e=>setSelYear(Number(e.target.value))} options={YEARS}/></div>
          </div>
        </div>
      </div>

      {/* Stat cards */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:1,background:IBM.gray20,border:`1px solid ${IBM.gray20}`}}>
        {[{l:"Period 1",v:`${totalP1}h`,c:IBM.blue60},{l:"Period 2",v:`${totalP2}h`,c:IBM.blue60},{l:"Total Entered",v:`${totalP1+totalP2}h`,c:IBM.green50},{l:"Scheduled",v:`${currentUser.scheduled||0}h`,c:IBM.gray70}].map(({l,v,c})=>(
          <div key={l} style={{background:"#fff",padding:"14px 18px",borderTop:`3px solid ${c}`}}>
            <div style={{fontSize:26,fontWeight:300,color:c}}>{v}</div>
            <div style={{fontSize:11,color:IBM.gray70,marginTop:4,textTransform:"uppercase",letterSpacing:"0.07em"}}>{l}</div>
          </div>
        ))}
      </div>

      {/* Period tabs + save */}
      <div style={{background:"#fff",borderBottom:`1px solid ${IBM.gray20}`,padding:"0 28px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
        <div style={{display:"flex"}}>
          {[["P1","Period 1  (1 – 15)"],["P2","Period 2  (16 – End)"]].map(([v,l])=>(
            <button key={v} onClick={()=>setActivePeriod(v)}
              style={{padding:"13px 20px",background:"none",border:"none",borderBottom:activePeriod===v?`3px solid ${IBM.blue60}`:"3px solid transparent",color:activePeriod===v?IBM.blue60:IBM.gray60,fontWeight:activePeriod===v?600:400,cursor:"pointer",fontSize:13,fontFamily:"inherit"}}>
              {l} <span style={{fontSize:11,color:IBM.gray50}}>({v==="P1"?totalP1:totalP2}h)</span>
            </button>
          ))}
        </div>
        <div style={{display:"flex",gap:10,alignItems:"center",padding:"8px 0"}}>
          {saved&&<span style={{fontSize:12,color:IBM.green50,fontWeight:600}}>✓ Saved</span>}
          <button onClick={handleSave} style={{padding:"9px 22px",background:IBM.blue60,color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:600}}>💾 Save Timesheet</button>
        </div>
      </div>

      {/* Grid */}
      <div style={{padding:"18px 28px"}}>
        <TimesheetEntryGrid
          userProjects={currentUser.projects||[]}
          period={activePeriod}
          monthName={selMonth}
          year={selYear}
          periodData={monthData[activePeriod]||{}}
          onUpdate={d=>handleUpdatePeriod(activePeriod,d)}
          periodNote={(monthData.periodNotes||{})[activePeriod]||""}
          onUpdateNote={v=>handleUpdateNote(activePeriod,v)}
          calendarEvents={calendarEvents}
          isManager={false}
        />
      </div>

      <SubmissionHistory history={currentUser.history}/>
    </div>
  );
}

// ─── CALENDAR EVENTS TAB ──────────────────────────────────────────────────────
function CalDayCell({d, selYear, mIdx, es, isWk, dk, setForm}) {
  var bg = es ? es.bg : (isWk ? "#fafafa" : "#fff");
  var cursor = (d && !isWk) ? "pointer" : "default";
  var numColor = es ? es.color : (isWk ? IBM.gray30 : IBM.gray100);
  function handleClick() { if (d && !isWk) setForm(function(f){ return Object.assign({}, f, {date: dk(d)}); }); }
  return (
    <div style={{minHeight:54,padding:"3px 4px",background:bg,borderRight:"1px solid "+IBM.gray20,borderBottom:"1px solid "+IBM.gray20,cursor:cursor}} onClick={handleClick}>
      {d && (
        <div>
          <div style={{fontSize:12,fontWeight:600,color:numColor}}>{d}</div>
          {es && <div style={{fontSize:9,color:es.color,fontWeight:600,lineHeight:1.3,overflow:"hidden"}}>{es.icon}</div>}
        </div>
      )}
    </div>
  );
}

function CalEventRow({e, es, selYear, mIdx, selMonth, calendarEvents, handleEdit, handleDelete}) {
  var d = parseInt(e.date.split("-")[2]);
  var days = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
  var dl = days[new Date(selYear, mIdx, d).getDay()];
  var gi = (calendarEvents||[]).findIndex(function(x){ return x.id===e.id; });
  return (
    <div key={e.id} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 12px",background:es.bg,border:"1px solid "+es.color,marginBottom:4,borderRadius:2}}>
      <span style={{fontSize:14}}>{es.icon}</span>
      <div style={{flex:1}}>
        <b style={{fontSize:13,color:es.color}}>{e.label}</b>
        <span style={{fontSize:11,color:IBM.gray60,marginLeft:8}}>{dl}, {selMonth} {d}</span>
      </div>
      <button onClick={function(){ handleEdit(gi); }} style={{background:"none",border:"1px solid "+IBM.gray30,color:IBM.gray70,padding:"2px 9px",cursor:"pointer",fontSize:11}}>Edit</button>
      <button onClick={function(){ handleDelete(gi); }} style={{background:"none",border:"1px solid "+IBM.red60,color:IBM.red60,padding:"2px 9px",cursor:"pointer",fontSize:11}}>&#x2715;</button>
    </div>
  );
}

function CalEventTableRow({e, i, calendarEvents, handleEdit, handleDelete, setSelMonth, setSelYear}) {
  var es = CAL_EVENT_TYPES[e.type];
  var dt = new Date(e.date);
  var days = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
  var dl = days[dt.getDay()];
  var gi = (calendarEvents||[]).findIndex(function(x){ return x.id===e.id; });
  var bg = i%2 ? IBM.gray10 : "#fff";
  function onEdit() {
    handleEdit(gi);
    setSelMonth(MONTH_NAMES[new Date(e.date).getMonth()]);
    setSelYear(new Date(e.date).getFullYear());
  }
  return (
    <tr style={{background:bg}}>
      <td style={{padding:"8px 12px",borderBottom:"1px solid "+IBM.gray20,fontSize:11}}>{e.date}</td>
      <td style={{padding:"8px 12px",borderBottom:"1px solid "+IBM.gray20,fontSize:11,color:IBM.gray60}}>{dl}</td>
      <td style={{padding:"8px 12px",borderBottom:"1px solid "+IBM.gray20}}>
        <span style={{fontSize:10,padding:"2px 7px",background:es.bg,color:es.color,border:"1px solid "+es.color,fontWeight:600}}>{es.icon} {es.label}</span>
      </td>
      <td style={{padding:"8px 12px",borderBottom:"1px solid "+IBM.gray20,fontWeight:600}}>{e.label}</td>
      <td style={{padding:"8px 12px",borderBottom:"1px solid "+IBM.gray20}}>
        <div style={{display:"flex",gap:5}}>
          <button onClick={onEdit} style={{background:"none",border:"1px solid "+IBM.gray30,color:IBM.gray70,padding:"2px 8px",cursor:"pointer",fontSize:11}}>Edit</button>
          <button onClick={function(){ handleDelete(gi); }} style={{background:"none",border:"1px solid "+IBM.red60,color:IBM.red60,padding:"2px 8px",cursor:"pointer",fontSize:11}}>&#x2715;</button>
        </div>
      </td>
    </tr>
  );
}

function CalendarEventsTab({calendarEvents,setCalendarEvents,showToast}){
  var now=new Date();
  const[selMonth,setSelMonth]=useState(MONTH_NAMES[now.getMonth()]);
  const[selYear,setSelYear]=useState(now.getFullYear());
  const[form,setForm]=useState({date:"",type:"holiday",label:""});
  const[editIdx,setEditIdx]=useState(null);
  var mIdx=MONTH_NAMES.indexOf(selMonth);
  var dim=new Date(selYear,mIdx+1,0).getDate();
  var firstDow=new Date(selYear,mIdx,1).getDay();
  var calDays=[];
  for(var ci=0;ci<firstDow;ci++) calDays.push(null);
  for(var cd=1;cd<=dim;cd++) calDays.push(cd);
  function dk(d){ return selYear+"-"+String(mIdx+1).padStart(2,"0")+"-"+String(d).padStart(2,"0"); }
  var mk2=selYear+"-"+String(mIdx+1).padStart(2,"0");
  var monthEvts=(calendarEvents||[]).filter(function(e){ return e.date.startsWith(mk2); });
  function dayEvts(d){ return (calendarEvents||[]).filter(function(e){ return e.date===dk(d); }); }

  function handleAdd(){
    if(!form.date||!form.label.trim()){showToast("Fill in date and label","error");return;}
    var nv={id:Date.now(),date:form.date,type:form.type,label:form.label.trim()};
    if(editIdx!==null){setCalendarEvents(function(p){ return p.map(function(e,i){ return i===editIdx?nv:e; }); });setEditIdx(null);showToast("Event updated");}
    else{setCalendarEvents(function(p){ return [...(p||[]),nv]; });showToast("Event added");}
    setForm({date:"",type:"holiday",label:""});
  }
  function handleEdit(idx){var e=calendarEvents[idx];setForm({date:e.date,type:e.type,label:e.label});setEditIdx(idx);}
  function handleDelete(idx){
    setCalendarEvents(function(p){ return p.filter(function(_,i){ return i!==idx; }); });
    if(editIdx===idx){setEditIdx(null);setForm({date:"",type:"holiday",label:""});}
    showToast("Event removed");
  }
  function prevMonth(){var d=new Date(selYear,mIdx-1,1);setSelMonth(MONTH_NAMES[d.getMonth()]);setSelYear(d.getFullYear());}
  function nextMonth(){var d=new Date(selYear,mIdx+1,1);setSelMonth(MONTH_NAMES[d.getMonth()]);setSelYear(d.getFullYear());}

  var sortedEvents = [...(calendarEvents||[])].sort(function(a,b){ return a.date.localeCompare(b.date); });
  var evtTypeColor = editIdx!==null ? IBM.teal50 : IBM.blue60;
  var evtTypeLabel = editIdx!==null ? "Edit Event" : "Add Calendar Event";
  var previewBg = CAL_EVENT_TYPES[form.type].bg;
  var previewColor = CAL_EVENT_TYPES[form.type].color;
  var previewIcon = CAL_EVENT_TYPES[form.type].icon;

  return (
    <div style={{padding:"24px 28px",maxWidth:980}}>
      <div style={{background:IBM.blue10,border:"1px solid "+IBM.blue20,padding:"12px 18px",marginBottom:20,fontSize:13,color:IBM.blue70,lineHeight:1.6}}>
        <b>Organisation Calendar Events</b> — Add org-wide dates that appear as highlighted bands on every employee timesheet automatically.
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 310px",gap:24,alignItems:"start"}}>
        <div>
          <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:14}}>
            <button onClick={prevMonth} style={{padding:"6px 11px",background:"#fff",border:"1px solid "+IBM.gray30,cursor:"pointer",fontSize:13}}>&#8249;</button>
            <Sel value={selMonth} onChange={function(e){setSelMonth(e.target.value);}} options={MONTH_NAMES}/>
            <Sel value={selYear} onChange={function(e){setSelYear(Number(e.target.value));}} options={YEARS}/>
            <button onClick={nextMonth} style={{padding:"6px 11px",background:"#fff",border:"1px solid "+IBM.gray30,cursor:"pointer",fontSize:13}}>&#8250;</button>
          </div>
          <div style={{background:"#fff",border:"1px solid "+IBM.gray20}}>
            <div style={{background:IBM.gray100,color:"#fff",padding:"10px 14px",fontSize:13,fontWeight:600}}>
              {selMonth} {selYear}
              <span style={{float:"right",fontSize:11,color:IBM.gray30,fontWeight:400}}>{monthEvts.length} events</span>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",borderBottom:"1px solid "+IBM.gray20}}>
              {["Su","Mo","Tu","We","Th","Fr","Sa"].map(function(day){
                return <div key={day} style={{padding:"5px 0",textAlign:"center",fontSize:11,fontWeight:600,color:IBM.gray60,borderRight:"1px solid "+IBM.gray20}}>{day}</div>;
              })}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)"}}>
              {calDays.map(function(d,i){
                var evts=d?dayEvts(d):[];
                var evt=evts[0];
                var es=evt?CAL_EVENT_TYPES[evt.type]:null;
                var isWk=d&&(new Date(selYear,mIdx,d).getDay()===0||new Date(selYear,mIdx,d).getDay()===6);
                return <CalDayCell key={i} d={d} selYear={selYear} mIdx={mIdx} es={es} isWk={isWk} dk={dk} setForm={setForm}/>;
              })}
            </div>
          </div>
          {monthEvts.length>0 && (
            <div style={{marginTop:10}}>
              {monthEvts.sort(function(a,b){return a.date.localeCompare(b.date);}).map(function(e){
                var es=CAL_EVENT_TYPES[e.type];
                return <CalEventRow key={e.id} e={e} es={es} selYear={selYear} mIdx={mIdx} selMonth={selMonth} calendarEvents={calendarEvents} handleEdit={handleEdit} handleDelete={handleDelete}/>;
              })}
            </div>
          )}
        </div>
        <div style={{background:"#fff",border:"1px solid "+IBM.gray20,position:"sticky",top:70}}>
          <div style={{background:evtTypeColor,color:"#fff",padding:"12px 18px",fontSize:13,fontWeight:600}}>{evtTypeLabel}</div>
          <div style={{padding:"16px"}}>
            <div style={{marginBottom:12}}>
              <label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",color:IBM.gray70,display:"block",marginBottom:5}}>Date</label>
              <input type="date" value={form.date} onChange={function(e){setForm(function(f){return Object.assign({},f,{date:e.target.value});});}} style={{width:"100%",padding:"8px 10px",border:"1px solid "+IBM.gray30,fontSize:13,outline:"none"}}/>
              <div style={{fontSize:10,color:IBM.gray50,marginTop:3}}>Or click any date on the calendar</div>
            </div>
            <div style={{marginBottom:12}}>
              <label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",color:IBM.gray70,display:"block",marginBottom:5}}>Event Type</label>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:4}}>
                {Object.entries(CAL_EVENT_TYPES).map(function(entry){
                  var k=entry[0],v=entry[1];
                  var isSelected=form.type===k;
                  return <button key={k} onClick={function(){setForm(function(f){return Object.assign({},f,{type:k});});}} style={{padding:"6px 7px",background:isSelected?v.bg:"#fff",border:"1px solid "+(isSelected?v.color:IBM.gray30),color:isSelected?v.color:IBM.gray70,cursor:"pointer",fontSize:10,fontWeight:600,textAlign:"left"}}>{v.icon} {v.label}</button>;
                })}
              </div>
            </div>
            <div style={{marginBottom:14}}>
              <label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",color:IBM.gray70,display:"block",marginBottom:5}}>Label</label>
              <input type="text" value={form.label} onChange={function(e){setForm(function(f){return Object.assign({},f,{label:e.target.value});});}} placeholder="e.g. Offshore Holiday" onKeyDown={function(e){if(e.key==="Enter")handleAdd();}} style={{width:"100%",padding:"8px 10px",border:"1px solid "+IBM.gray30,fontSize:13,outline:"none"}}/>
            </div>
            {form.date&&form.label&&(
              <div style={{marginBottom:12,padding:"7px 10px",background:previewBg,border:"1px solid "+previewColor,fontSize:12,color:previewColor,fontWeight:600,borderRadius:2}}>
                {previewIcon} {form.label} — {form.date}
              </div>
            )}
            <div style={{display:"flex",gap:7}}>
              <button onClick={handleAdd} style={{flex:1,padding:"9px",background:IBM.blue60,color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:600}}>{editIdx!==null?"Update":"Add Event"}</button>
              {editIdx!==null&&<button onClick={function(){setEditIdx(null);setForm({date:"",type:"holiday",label:""}); }} style={{padding:"9px 12px",background:"none",border:"1px solid "+IBM.gray30,color:IBM.gray70,cursor:"pointer",fontSize:13}}>Cancel</button>}
            </div>
            <div style={{marginTop:12,padding:"9px 10px",background:IBM.gray10,border:"1px solid "+IBM.gray20,fontSize:11,color:IBM.gray60,lineHeight:1.6}}>
              Events appear on all employee timesheets as coloured date highlights.
            </div>
          </div>
        </div>
      </div>
      {(calendarEvents||[]).length>0&&(
        <div style={{marginTop:24}}>
          <div style={{fontSize:11,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.07em",color:IBM.gray60,marginBottom:8}}>All Events ({calendarEvents.length})</div>
          <div style={{overflowX:"auto",border:"1px solid "+IBM.gray20}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead>
                <tr style={{background:IBM.gray100,color:"#fff"}}>
                  {["Date","Day","Type","Label",""].map(function(h){
                    return <th key={h} style={{padding:"7px 12px",textAlign:"left",fontWeight:400,fontSize:11,textTransform:"uppercase",borderRight:"1px solid "+IBM.gray80}}>{h}</th>;
                  })}
                </tr>
              </thead>
              <tbody>
                {sortedEvents.map(function(e,i){
                  return <CalEventTableRow key={e.id} e={e} i={i} calendarEvents={calendarEvents} handleEdit={handleEdit} handleDelete={handleDelete} setSelMonth={setSelMonth} setSelYear={setSelYear}/>;
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── EMPLOYEE DETAIL PANEL (Manager) ─────────────────────────────────────────
// KEY FIX: receives userId and looks up LIVE user from users array every render
function EmployeeDetailPanel({userId,users,monthLabel,periodLabel,onFixEntry,onSendEmail,onSendTeams,onClose,calendarEvents}){
  const user=users.find(u=>u.id===userId);
  if(!user)return null;

  const[panelTab,setPanelTab]=useState("overview");
  const[mgrMonth,setMgrMonth]=useState(MONTH_NAMES[new Date().getMonth()]);
  const[mgrYear,setMgrYear]=useState(new Date().getFullYear());
  const[mgrPeriod,setMgrPeriod]=useState("P1");
  const[notifText,setNotifText]=useState(null);
  const[sentEmail,setSentEmail]=useState(false);
  const[sentTeams,setSentTeams]=useState(false);
  const[showFix,setShowFix]=useState(false);
  const[fixHours,setFixHours]=useState("");
  const[fixNote,setFixNote]=useState("");

  const sev=getSeverity(user),status=getStatus(user);
  const diff=Number(user.scheduled)-Number(user.entered);
  const pct=user.scheduled?Math.round((user.entered/user.scheduled)*100):0;

  // Live monthly data
  const mk=monthKey(mgrMonth,mgrYear);
  const monthData=(user.monthlyEntries||{})[mk]||makeEmptyMonthEntries(mgrMonth,mgrYear);

  // Compute live monthly total
  const liveMonthTotal=useMemo(()=>{
    let t=0;
    ["P1","P2"].forEach(p=>{
      const pd=monthData[p]||{};
      Object.values(pd).forEach(rows=>rows.forEach(r=>t+=parseFloat(r.hours)||0));
    });
    return t;
  },[monthData]);

  const projectSummary=useMemo(()=>{const map={};(user.history||[]).forEach(r=>{const k=r.projectCode;if(!map[k])map[k]={code:r.projectCode,name:r.projectName,scheduled:0,entered:0,diff:0,periods:0,missPeriods:0};map[k].scheduled+=r.scheduled;map[k].entered+=r.entered;map[k].diff+=r.diff;map[k].periods++;if(r.diff>0)map[k].missPeriods++;});return Object.values(map);},[user]);
  const timelineData=useMemo(()=>{const bp={};(user.history||[]).forEach(r=>{const k=r.periodLabel;if(!bp[k])bp[k]={period:k,scheduled:0,entered:0};bp[k].scheduled+=r.scheduled;bp[k].entered+=r.entered;});return Object.values(bp);},[user]);

  const mTh={background:IBM.gray100,color:"#fff",padding:"8px 12px",textAlign:"left",fontWeight:400,fontSize:11,textTransform:"uppercase",letterSpacing:"0.05em",borderRight:`1px solid ${IBM.gray80}`,whiteSpace:"nowrap"};
  const mTd=(alt,hi)=>({padding:"9px 12px",borderBottom:`1px solid ${IBM.gray20}`,background:hi?IBM.red10:alt?IBM.gray10:"#fff",fontSize:13});

  return(
    <div style={{position:"fixed",inset:0,background:"rgba(22,22,22,.55)",zIndex:400,display:"flex",justifyContent:"flex-end"}} onClick={onClose}>
      <div className="panel-slide" style={{background:"#fff",width:"min(960px,100vw)",height:"100vh",overflowY:"auto",boxShadow:"-8px 0 32px rgba(0,0,0,.18)",display:"flex",flexDirection:"column",fontFamily:FF_SANS}} onClick={e=>e.stopPropagation()}>
        {/* Header */}
        <div style={{background:IBM.gray100,color:"#fff",padding:"18px 28px",flexShrink:0}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div style={{display:"flex",alignItems:"center",gap:12}}>
              <div style={{width:42,height:42,borderRadius:"50%",background:IBM.blue60,display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,fontWeight:700,color:"#fff",flexShrink:0}}>{user.name.split(" ").map(n=>n[0]).join("")}</div>
              <div><div style={{fontSize:19,fontWeight:600}}>{user.name}</div>
              {user.clarityName && user.clarityName !== user.name && (
                <div style={{display:"flex",alignItems:"center",gap:6,marginTop:3}}>
                  <span style={{fontSize:10,background:"rgba(255,255,255,0.15)",border:"1px solid rgba(255,255,255,0.4)",color:"#fff",padding:"1px 6px",fontWeight:700,letterSpacing:"0.04em"}}>BMO</span>
                  <span style={{fontSize:13,color:IBM.gray30}}>{user.clarityName}</span>
                </div>
              )}
              <div style={{fontSize:12,color:IBM.gray30,marginTop:2}}>{user.email} · {user.dept} · {user.id}</div></div>
            </div>
            <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:8}}>
              <button onClick={onClose} style={{background:"none",border:`1px solid ${IBM.gray70}`,color:IBM.gray30,fontSize:13,cursor:"pointer",padding:"5px 14px"}}>✕ Close</button>
              <SevBadge sev={sev}/>
            </div>
          </div>
        </div>

        {/* Panel tabs */}
        <div style={{background:IBM.gray90,padding:"0 28px",display:"flex",flexShrink:0}}>
          {[["overview","📊 Overview"],["timesheet","📅 Timesheet"],["notify","📋 Notifications"]].map(([v,l])=>(
            <button key={v} onClick={()=>setPanelTab(v)} style={{padding:"10px 18px",background:"none",border:"none",borderBottom:panelTab===v?`2px solid ${IBM.blue60}`:"2px solid transparent",color:panelTab===v?"#fff":IBM.gray50,cursor:"pointer",fontSize:13,fontFamily:"inherit",fontWeight:panelTab===v?600:400}}>{l}</button>
          ))}
        </div>

        <div style={{flex:1,overflowY:"auto",padding:"20px 28px 32px"}}>

          {/* OVERVIEW */}
          {panelTab==="overview"&&(
            <React.Fragment>
              <div style={{fontSize:11,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.07em",color:IBM.gray70,marginBottom:10,borderBottom:`1px solid ${IBM.gray20}`,paddingBottom:5}}>Current Period Summary</div>
              {/* IBM vs Clarity Comparison Summary */}
              <div style={{marginBottom:12}}>
                {/* Top bar: Variance callout */}
                <div style={{background:diff===0?IBM.green10:diff>0?IBM.red10:IBM.yellow10, border:"1px solid "+(diff===0?IBM.green20:diff>0?IBM.red20:IBM.yellow20), padding:"12px 18px", display:"flex", alignItems:"center", gap:16, flexWrap:"wrap", marginBottom:1}}>
                  <div>
                    <div style={{fontSize:10,color:IBM.gray60,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Variance (IBM Scheduled vs Clarity Actual)</div>
                    <div style={{fontSize:28,fontWeight:300,color:diff===0?IBM.green50:diff>0?IBM.red60:IBM.orange40}}>
                      {diff===0?"0h — On Track":diff>0?"-"+diff+"h under-reported":"+"+Math.abs(diff)+"h over-reported"}
                    </div>
                  </div>
                  {user.timesheetStatus&&(
                    <div style={{marginLeft:"auto",textAlign:"right"}}>
                      <div style={{fontSize:10,color:IBM.gray50,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Clarity Status</div>
                      <span style={{fontSize:13,fontWeight:700,
                        color:user.timesheetStatus==="Approved"?IBM.green50:user.timesheetStatus==="Posted"?IBM.blue60:user.timesheetStatus==="Not in Clarity"?IBM.gray50:"#8e6a00",
                        background:user.timesheetStatus==="Approved"?IBM.green10:user.timesheetStatus==="Posted"?IBM.blue10:user.timesheetStatus==="Not in Clarity"?IBM.gray10:IBM.yellow10,
                        padding:"4px 12px",border:"1px solid "+(user.timesheetStatus==="Approved"?IBM.green20:user.timesheetStatus==="Posted"?IBM.blue20:user.timesheetStatus==="Not in Clarity"?IBM.gray20:IBM.yellow20)}}>
                        {user.timesheetStatus}
                      </span>
                    </div>
                  )}
                </div>

                {/* Side-by-side IBM vs Clarity cards */}
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:1,background:IBM.gray20,marginBottom:1}}>
                  {/* IBM side */}
                  <div style={{background:"#fff",padding:"14px 18px",borderTop:"3px solid "+IBM.blue60}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                      <span style={{fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.07em",color:IBM.blue60}}>IBM Scheduled</span>
                      {user.dataSource==="Both"&&<span style={{fontSize:9,background:IBM.blue10,color:IBM.blue60,padding:"1px 5px",border:"1px solid "+IBM.blue20}}>matched</span>}
                    </div>
                    <div style={{fontSize:32,fontWeight:300,color:IBM.blue60,marginBottom:6}}>{user.scheduled||0}h</div>
                    <div style={{fontSize:11,color:IBM.gray60,marginBottom:4}}>
                      {user.claimMonths&&user.claimMonths.length>0&&<div><b>Claim months:</b> {user.claimMonths.join(", ")}</div>}
                      {user.wbsId&&<div><b>WBS:</b> {user.wbsId}</div>}
                      {user.billingCode&&<div><b>Billing:</b> {user.billingCode}</div>}
                    </div>
                    {/* Weekly bar chart */}
                    {user.weeklyBreakdown&&user.weeklyBreakdown.length>0&&(
                      <div style={{marginTop:8}}>
                        <div style={{fontSize:10,color:IBM.gray50,marginBottom:4}}>{user.weeklyBreakdown.length} week{user.weeklyBreakdown.length>1?"s":""} of data</div>
                        <div style={{display:"flex",gap:3,alignItems:"flex-end",height:36}}>
                          {user.weeklyBreakdown.slice().reverse().map(function(w,wi){
                            var maxH = Math.max.apply(null, user.weeklyBreakdown.map(function(x){return x.total||0;}));
                            var barH = maxH > 0 ? Math.round(((w.total||0)/maxH)*32) : 0;
                            return (
                              <div key={wi} title={"W/E "+w.weekEnd+": "+w.total+"h"} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:2}}>
                                <div style={{fontSize:8,color:IBM.blue60,fontWeight:700}}>{w.total||0}</div>
                                <div style={{width:"100%",background:IBM.blue60,height:barH+"px",minHeight:2}}/>
                                <div style={{fontSize:7,color:IBM.gray50,textAlign:"center",lineHeight:1}}>{w.weekEnd?w.weekEnd.slice(0,5):""}</div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    )}
                  </div>

                  {/* Clarity side */}
                  <div style={{background:"#fff",padding:"14px 18px",borderTop:"3px solid "+IBM.purple60}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                      <span style={{fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.07em",color:IBM.purple60}}>Clarity Actual</span>
                      {user.dataSource==="Clarity only"&&<span style={{fontSize:9,background:IBM.purple10,color:IBM.purple60,padding:"1px 5px",border:"1px solid #d4bbff"}}>clarity only</span>}
                    </div>
                    <div style={{fontSize:32,fontWeight:300,color:user.entered>0?IBM.purple60:IBM.gray30,marginBottom:6}}>
                      {user.entered>0?user.entered+"h":"—"}
                    </div>
                    <div style={{fontSize:11,color:IBM.gray60,marginBottom:4}}>
                      {user.resourceManager&&user.resourceManager!=="—"&&<div><b>Manager:</b> {user.resourceManager}</div>}
                      {user.approvedBy&&<div><b>Approved by:</b> {user.approvedBy}</div>}
                      {user.resourceActive&&<div><b>Active:</b> {user.resourceActive}</div>}
                    </div>
                    {/* Clarity periods breakdown */}
                    {user.clarityPeriods&&user.clarityPeriods.length>0&&(
                      <div style={{marginTop:8}}>
                        <div style={{fontSize:10,color:IBM.gray50,marginBottom:4}}>{user.clarityPeriods.length} reporting period{user.clarityPeriods.length>1?"s":""}</div>
                        <div style={{display:"flex",flexDirection:"column",gap:3}}>
                          {(function(){
                            var perPeriod = user.entered>0&&user.clarityPeriods.length>0 ? Math.round(user.entered/user.clarityPeriods.length) : 0;
                            return user.clarityPeriods.map(function(p,pi){
                              return (
                                <div key={pi} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"3px 6px",background:IBM.purple10,border:"1px solid #d4bbff"}}>
                                  <span style={{fontSize:10,color:IBM.purple60,fontWeight:600,flex:1,marginRight:6}}>{p}</span>
                                  <span style={{fontSize:11,fontWeight:700,color:IBM.purple60,whiteSpace:"nowrap"}}>{perPeriod}h</span>
                                </div>
                              );
                            });
                          })()}
                        </div>
                      </div>
                    )}
                    {(!user.clarityPeriods||user.clarityPeriods.length===0)&&user.dataSource!=="Both"&&(
                      <div style={{marginTop:8,padding:"6px 8px",background:IBM.red10,border:"1px solid "+IBM.red20,fontSize:11,color:IBM.red60}}>Not found in Clarity</div>
                    )}
                  </div>
                </div>

                {/* Completion progress bar */}
                <div style={{background:"#fff",padding:"10px 18px",border:"1px solid "+IBM.gray20}}>
                  <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:IBM.gray60,marginBottom:5}}>
                    <span>Clarity actual vs IBM scheduled</span>
                    <span style={{fontWeight:700,color:diff===0?IBM.green50:diff>0?IBM.red60:IBM.orange40}}>{pct}% complete</span>
                  </div>
                  <div style={{height:8,background:IBM.gray20,overflow:"hidden",borderRadius:1}}>
                    <div style={{height:"100%",width:Math.min(pct,100)+"%",background:pct>=100?IBM.green50:pct<50?IBM.red60:IBM.orange40,transition:"width 0.3s"}}/>
                  </div>
                </div>
              </div>
              {/* IBM + Clarity Fields */}
              {(user.wbsId||user.talentId||user.billingCode||user.approvedBy||user.periods.length>0)&&(
                <div style={{background:IBM.gray10,border:"1px solid "+IBM.gray20,padding:"10px 16px",marginBottom:14,display:"flex",gap:20,flexWrap:"wrap"}}>
                  {user.wbsId&&<div><div style={{fontSize:10,color:IBM.gray50,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:2}}>WBS ID</div><div style={{fontSize:12,fontWeight:600}}>{user.wbsId}</div></div>}
                  {user.talentId&&<div><div style={{fontSize:10,color:IBM.gray50,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:2}}>Talent ID</div><div style={{fontSize:12,fontWeight:600}}>{user.talentId}</div></div>}
                  {user.serialId&&<div><div style={{fontSize:10,color:IBM.gray50,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:2}}>Serial</div><div style={{fontSize:12,fontWeight:600}}>{user.serialId}</div></div>}
                  {user.billingCode&&<div><div style={{fontSize:10,color:IBM.gray50,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:2}}>Billing Code</div><div style={{fontSize:12,fontWeight:600}}>{user.billingCode}</div></div>}
                  {user.country&&<div><div style={{fontSize:10,color:IBM.gray50,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:2}}>Country</div><div style={{fontSize:12,fontWeight:600}}>{user.country}</div></div>}
                  {user.approvedBy&&<div><div style={{fontSize:10,color:IBM.gray50,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:2}}>Approved By</div><div style={{fontSize:12,fontWeight:600}}>{user.approvedBy}</div></div>}
                  {user.periods&&user.periods.length>0&&<div><div style={{fontSize:10,color:IBM.gray50,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:2}}>Reporting Period(s)</div><div style={{fontSize:12,fontWeight:600}}>{user.periods.join(", ")}</div></div>}
                </div>
              )}
              <div style={{display:"flex",justifyContent:"space-between",fontSize:12,color:IBM.gray60,marginBottom:4}}><span>Completion</span><span>{user.entered}h / {user.scheduled}h</span></div>
              <div style={{height:8,background:IBM.gray20,overflow:"hidden",marginBottom:18}}><div style={{height:"100%",width:`${Math.min(pct,100)}%`,background:pct===100?IBM.green50:pct<40?IBM.red60:IBM.orange40}}/></div>

              {status!=="green"&&(
                <div style={{background:IBM.gray10,border:`1px solid ${IBM.gray20}`,padding:"12px 16px",marginBottom:18}}>
                  <div style={{fontSize:11,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.07em",color:IBM.gray60,marginBottom:10}}>Manager Actions</div>
                  <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                    <button onClick={()=>{setNotifText(buildNotifTemplate(user,status,monthLabel,periodLabel));setPanelTab("notify");}} style={{padding:"9px 18px",background:IBM.blue60,color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:600}}>📋 Draft Notification</button>
                    <button onClick={()=>{onSendTeams&&onSendTeams(user);setSentTeams(true);}} style={{padding:"9px 18px",background:sentTeams?"#464775":"#5b5ea6",color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:600}}>{sentTeams?"✓ Teams Sent":"💬 Teams"}</button>
                    <button onClick={()=>{onSendEmail&&onSendEmail(user);setSentEmail(true);}} style={{padding:"9px 18px",background:sentEmail?IBM.green50:IBM.orange40,color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:600}}>{sentEmail?"✓ Email Sent":"✉ Email"}</button>
                    <button onClick={()=>setShowFix(true)} style={{padding:"9px 18px",background:"#fff",color:IBM.teal50,border:`2px solid ${IBM.teal50}`,cursor:"pointer",fontSize:13,fontWeight:600}}>🔧 Fix Entry</button>
                  </div>
                </div>
              )}

              <div style={{fontSize:11,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.07em",color:IBM.gray70,marginBottom:8,borderBottom:`1px solid ${IBM.gray20}`,paddingBottom:5}}>Project Breakdown</div>
              <div style={{overflowX:"auto",border:`1px solid ${IBM.gray20}`,marginBottom:18}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                  <thead><tr>{["Project","Name","Scheduled","Entered","Gap","Miss Rate"].map(h=><th key={h} style={mTh}>{h}</th>)}</tr></thead>
                  <tbody>{projectSummary.map((p,i)=>(
                    <tr key={p.code}>
                      <td style={mTd(i%2,p.diff>20)}><code style={{fontSize:12,background:"#dde1e7",padding:"1px 5px"}}>{p.code}</code></td>
                      <td style={mTd(i%2,p.diff>20)}>{p.name}</td>
                      <td style={mTd(i%2,p.diff>20)}>{p.scheduled}h</td>
                      <td style={mTd(i%2,p.diff>20)}>{p.entered}h</td>
                      <td style={mTd(i%2,p.diff>20)}>{p.diff===0?<span style={{color:IBM.green50,fontWeight:600}}>✓</span>:<span style={{color:IBM.red60,fontWeight:600}}>-{p.diff}h</span>}</td>
                      <td style={mTd(i%2,p.diff>20)}>{p.missPeriods}/{p.periods} periods</td>
                    </tr>
                  ))}</tbody>
                </table>
              </div>

              <div style={{fontSize:11,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.07em",color:IBM.gray70,marginBottom:8,borderBottom:`1px solid ${IBM.gray20}`,paddingBottom:5}}>Hours Timeline</div>
              <div style={{background:"#fff",border:`1px solid ${IBM.gray20}`,padding:"12px"}}>
                <ResponsiveContainer width="100%" height={140}><BarChart data={timelineData} margin={{top:0,right:0,bottom:0,left:-18}}><CartesianGrid strokeDasharray="3 3" stroke={IBM.gray20}/><XAxis dataKey="period" tick={{fontSize:9}} angle={-30} textAnchor="end" height={46}/><YAxis tick={{fontSize:10}}/><Tooltip formatter={(v,n)=>[`${v}h`,n]}/><Legend/><Bar dataKey="scheduled" name="Scheduled" fill={IBM.blue20}/><Bar dataKey="entered" name="Entered" fill={IBM.blue60}/></BarChart></ResponsiveContainer>
              </div>
            </React.Fragment>
          )}

          {/* TIMESHEET — live data */}
          {panelTab==="timesheet"&&(
            <React.Fragment>
              {/* Employee notes banner */}
              {((monthData.periodNotes && monthData.periodNotes.P1)||(monthData.periodNotes && monthData.periodNotes.P2))&&(
                <div style={{background:IBM.yellow10,border:`1px solid ${IBM.yellow20}`,padding:"10px 16px",marginBottom:14,borderRadius:2}}>
                  <div style={{fontSize:11,fontWeight:600,color:"#8e6a00",marginBottom:5}}>💬 Employee Notes — {mgrMonth} {mgrYear}</div>
                  {monthData.periodNotes.P1&&<div style={{fontSize:13,color:IBM.gray80,marginBottom:3}}><b style={{fontSize:11,color:IBM.gray60}}>Period 1:</b> {monthData.periodNotes.P1}</div>}
                  {monthData.periodNotes.P2&&<div style={{fontSize:13,color:IBM.gray80}}><b style={{fontSize:11,color:IBM.gray60}}>Period 2:</b> {monthData.periodNotes.P2}</div>}
                </div>
              )}

              {/* Live total indicator */}
              {liveMonthTotal>0&&(
                <div style={{background:IBM.green10,border:`1px solid ${IBM.green20}`,padding:"8px 14px",marginBottom:12,fontSize:13,fontWeight:600,color:IBM.green50,borderRadius:2}}>
                  ✓ {user.name} has entered <b>{liveMonthTotal}h</b> for {mgrMonth} {mgrYear}
                </div>
              )}
              {liveMonthTotal===0&&(
                <div style={{background:IBM.red10,border:`1px solid ${IBM.red20}`,padding:"8px 14px",marginBottom:12,fontSize:13,color:IBM.red60,borderRadius:2}}>
                  ⚠ No timesheet entries found for {mgrMonth} {mgrYear}
                </div>
              )}

              {/* Month/Period controls */}
              <div style={{display:"flex",gap:10,marginBottom:14,alignItems:"flex-end",flexWrap:"wrap"}}>
                <div><label style={{fontSize:11,color:IBM.gray60,display:"block",marginBottom:4}}>Month</label><Sel value={mgrMonth} onChange={e=>setMgrMonth(e.target.value)} options={MONTH_NAMES}/></div>
                <div><label style={{fontSize:11,color:IBM.gray60,display:"block",marginBottom:4}}>Year</label><Sel value={mgrYear} onChange={e=>setMgrYear(Number(e.target.value))} options={YEARS}/></div>
                <div style={{display:"flex",gap:6}}>
                  {[["P1","Period 1"],["P2","Period 2"]].map(([v,l])=>(
                    <button key={v} onClick={()=>setMgrPeriod(v)} style={{padding:"8px 16px",background:mgrPeriod===v?IBM.blue60:"#fff",color:mgrPeriod===v?"#fff":IBM.gray70,border:`1px solid ${mgrPeriod===v?IBM.blue60:IBM.gray30}`,cursor:"pointer",fontSize:12,fontWeight:600,fontFamily:"inherit"}}>{l}</button>
                  ))}
                </div>
              </div>

              {user.weeklyBreakdown && user.weeklyBreakdown.length > 0 ? (
                <div>
                  {user.clarityPeriods && user.clarityPeriods.length > 0 && (
                    <div style={{marginBottom:16}}>
                      <div style={{fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.07em",color:IBM.purple60,marginBottom:8,borderBottom:"2px solid "+IBM.purple60,paddingBottom:4}}>Clarity — Actual Hours by Period</div>
                      <div style={{display:"flex",gap:1,background:IBM.gray20}}>
                        {user.clarityPeriods.map(function(period, pi) {
                          var totalActual = user.entered || 0;
                          var split = Math.round(totalActual / (user.clarityPeriods.length||1));
                          return (
                            <div key={pi} style={{flex:1,background:"#fff",padding:"10px 14px",borderTop:"3px solid "+IBM.purple60}}>
                              <div style={{fontSize:10,color:IBM.gray50,textTransform:"uppercase",marginBottom:3}}>{period}</div>
                              <div style={{fontSize:18,fontWeight:300,color:IBM.purple60}}>{split}h</div>
                              <div style={{fontSize:10,color:IBM.gray50,marginTop:2}}>Status: <b style={{color:user.timesheetStatus==="Posted"||user.timesheetStatus==="Approved"?IBM.green50:"#8e6a00"}}>{user.timesheetStatus||"—"}</b></div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  )}
                  <div style={{fontSize:11,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.07em",color:IBM.blue60,marginBottom:8,borderBottom:"2px solid "+IBM.blue60,paddingBottom:4}}>IBM — Weekly Scheduled Hours</div>
                  <div style={{overflowX:"auto",border:"1px solid "+IBM.gray20}}>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                      <thead>
                        <tr style={{background:IBM.gray100,color:"#fff"}}>
                          {["W/E Date","Month","Mon","Tue","Wed","Thu","Fri","Sat","Sun","Total","Workitem","Activity"].map(function(h){
                            return <th key={h} style={{padding:"7px 9px",textAlign:"center",fontWeight:400,fontSize:10,textTransform:"uppercase",borderRight:"1px solid "+IBM.gray80,whiteSpace:"nowrap"}}>{h}</th>;
                          })}
                        </tr>
                      </thead>
                      <tbody>
                        {user.weeklyBreakdown.map(function(w, wi) {
                          var rowBg = wi%2 ? IBM.gray10 : "#fff";
                          function dc(v){ return {padding:"7px 9px",borderBottom:"1px solid "+IBM.gray20,textAlign:"center",fontWeight:v>0?700:400,color:v>0?IBM.gray100:IBM.gray30,background:rowBg}; }
                          return (
                            <tr key={wi}>
                              <td style={{padding:"7px 9px",borderBottom:"1px solid "+IBM.gray20,fontWeight:600,whiteSpace:"nowrap",background:rowBg}}>{w.weekEnd||"—"}</td>
                              <td style={{padding:"7px 9px",borderBottom:"1px solid "+IBM.gray20,fontSize:11,color:IBM.gray60,background:rowBg}}>{w.claimMonth||"—"}</td>
                              <td style={dc(w.mon)}>{w.mon||"—"}</td>
                              <td style={dc(w.tue)}>{w.tue||"—"}</td>
                              <td style={dc(w.wed)}>{w.wed||"—"}</td>
                              <td style={dc(w.thu)}>{w.thu||"—"}</td>
                              <td style={dc(w.fri)}>{w.fri||"—"}</td>
                              <td style={dc(w.sat)}>{w.sat||"—"}</td>
                              <td style={dc(w.sun)}>{w.sun||"—"}</td>
                              <td style={{padding:"7px 9px",borderBottom:"1px solid "+IBM.gray20,textAlign:"center",fontWeight:700,color:IBM.blue60,background:rowBg}}>{w.total}h</td>
                              <td style={{padding:"7px 9px",borderBottom:"1px solid "+IBM.gray20,fontSize:11,color:IBM.gray60,maxWidth:140,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",background:rowBg}} title={w.workitem}>{w.workitem||"—"}</td>
                              <td style={{padding:"7px 9px",borderBottom:"1px solid "+IBM.gray20,fontSize:11,color:IBM.gray60,background:rowBg}}>{w.activityCode||"—"}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                      <tfoot>
                        <tr style={{background:IBM.blue10}}>
                          <td colSpan={2} style={{padding:"8px 9px",fontWeight:700,fontSize:12,color:IBM.blue70}}>Total</td>
                          {(function(){
                            var t={mon:0,tue:0,wed:0,thu:0,fri:0,sat:0,sun:0,total:0};
                            user.weeklyBreakdown.forEach(function(w){t.mon+=w.mon||0;t.tue+=w.tue||0;t.wed+=w.wed||0;t.thu+=w.thu||0;t.fri+=w.fri||0;t.sat+=w.sat||0;t.sun+=w.sun||0;t.total+=w.total||0;});
                            return ["mon","tue","wed","thu","fri","sat","sun","total"].map(function(k){
                              return <td key={k} style={{padding:"8px 9px",fontWeight:700,textAlign:"center",color:k==="total"?IBM.blue60:IBM.blue70,fontSize:k==="total"?13:12}}>{t[k]||"—"}{k==="total"?"h":""}</td>;
                            }).concat([<td key="x" colSpan={2}/>]);
                          })()}
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                </div>
              ) : (
                <TimesheetEntryGrid
                  userProjects={user.projects||[]}
                  period={mgrPeriod}
                  monthName={mgrMonth}
                  year={mgrYear}
                  periodData={monthData[mgrPeriod]||{}}
                  onUpdate={function(){}}
                  periodNote={(monthData.periodNotes||{})[mgrPeriod]||""}
                  onUpdateNote={null}
                  calendarEvents={calendarEvents}
                  isManager={true}
                />
              )}
            </React.Fragment>
          )}

          {/* NOTIFY */}
          {panelTab==="notify"&&(
            <React.Fragment>
              <div style={{fontSize:11,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.07em",color:IBM.gray70,marginBottom:12,borderBottom:`1px solid ${IBM.gray20}`,paddingBottom:5}}>Notification Template</div>
              {!notifText?(
                <div>
                  <button onClick={()=>setNotifText(buildNotifTemplate(user,status,monthLabel,periodLabel))} style={{padding:"10px 24px",background:IBM.blue60,color:"#fff",border:"none",cursor:"pointer",fontSize:14,fontWeight:600,marginBottom:10}}>📋 Generate Notification</button>
                  <p style={{fontSize:12,color:IBM.gray60}}>Auto-fills with mismatch details, dates & hours</p>
                </div>
              ):(
                <div>
                  <div style={{background:IBM.gray10,padding:"10px 14px",marginBottom:10,fontSize:13}}><b>Subject:</b> {notifText.subject}</div>
                  <textarea value={notifText.body} readOnly rows={16} style={{width:"100%",padding:"12px",border:`1px solid ${IBM.gray30}`,fontSize:12,fontFamily:FF_MONO,resize:"vertical",outline:"none"}}/>
                  <div style={{display:"flex",gap:10,marginTop:10}}>
                    <button onClick={()=>{onSendEmail(user);setSentEmail(true);}} style={{padding:"9px 18px",background:IBM.orange40,color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:600}}>✉ Email</button>
                    <button onClick={()=>{onSendTeams(user);setSentTeams(true);}} style={{padding:"9px 18px",background:"#5b5ea6",color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:600}}>💬 Teams</button>
                  </div>
                </div>
              )}
            </React.Fragment>
          )}
        </div>

        {/* Fix modal */}
        {showFix&&(
          <div style={{position:"fixed",inset:0,background:"rgba(22,22,22,.7)",zIndex:600,display:"flex",alignItems:"center",justifyContent:"center"}} onClick={()=>setShowFix(false)}>
            <div style={{background:"#fff",width:"min(460px,96vw)",border:"1px solid #e0e0e0",fontFamily:FF_SANS}} onClick={e=>e.stopPropagation()}>
              <div style={{background:IBM.teal50,color:"#fff",padding:"14px 20px",display:"flex",justifyContent:"space-between"}}><b style={{fontSize:15}}>🔧 Fix Entry</b><button onClick={()=>setShowFix(false)} style={{background:"none",border:"none",color:"#fff",fontSize:22,cursor:"pointer"}}>×</button></div>
              <div style={{padding:"20px"}}>
                <div style={{marginBottom:14}}><label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",color:IBM.gray70,display:"block",marginBottom:6}}>New Scheduled Hours</label><input type="number" min="0" max="200" value={fixHours} onChange={e=>setFixHours(e.target.value)} style={{width:"100%",padding:"9px 12px",border:`1px solid ${IBM.gray30}`,fontSize:14,outline:"none"}}/></div>
                <div style={{marginBottom:16}}><label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",color:IBM.gray70,display:"block",marginBottom:6}}>Manager Note</label><textarea value={fixNote} onChange={e=>setFixNote(e.target.value)} rows={3} style={{width:"100%",padding:"9px 12px",border:`1px solid ${IBM.gray30}`,fontSize:13,outline:"none",resize:"vertical",fontFamily:"inherit"}}/></div>
                <div style={{display:"flex",gap:8}}>
                  <button disabled={!fixHours} onClick={()=>{if(!fixHours)return;onFixEntry(user.id,Number(fixHours),fixNote);setShowFix(false);}} style={{padding:"10px 22px",background:!fixHours?IBM.gray30:IBM.teal50,color:!fixHours?IBM.gray60:"#fff",border:"none",cursor:!fixHours?"not-allowed":"pointer",fontSize:13,fontWeight:600}}>✓ Apply</button>
                  <button onClick={()=>setShowFix(false)} style={{padding:"10px 14px",background:"none",border:`1px solid ${IBM.gray30}`,color:IBM.gray70,cursor:"pointer",fontSize:13}}>Cancel</button>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
// ─── IMPORT PARSING HELPERS ─────────────────────────────────────────────────

function normalizeName(raw) {
  if (!raw) return "";
  var s = String(raw).trim().toLowerCase();
  // Remove punctuation except spaces (dots, commas, apostrophes etc.)
  s = s.replace(/[^a-z ]/g, " ");
  var tokens = s.split(/\s+/).filter(function(t){ return t.length > 0; });
  // Merge consecutive single-letter tokens: ["p","m"] -> "pm", ["a","b","c"] -> "abc"
  var merged = [];
  var i = 0;
  while (i < tokens.length) {
    if (tokens[i].length === 1) {
      var group = "";
      while (i < tokens.length && tokens[i].length === 1) { group += tokens[i]; i++; }
      merged.push(group);
    } else {
      merged.push(tokens[i]);
      i++;
    }
  }
  // Drop any leftover single-char tokens (safety), then sort alphabetically
  var parts = merged.filter(function(p){ return p.length > 1; });
  parts.sort();
  return parts.join(" ");
}

function findSheet(wb, searchStr) {
  var lower = searchStr.toLowerCase();
  for (var i = 0; i < wb.SheetNames.length; i++) {
    if (wb.SheetNames[i].toLowerCase().indexOf(lower) !== -1) return wb.SheetNames[i];
  }
  return null;
}

function getCol(row, candidates) {
  var keys = Object.keys(row);
  for (var ci = 0; ci < candidates.length; ci++) {
    var cand = candidates[ci].toLowerCase();
    for (var ki = 0; ki < keys.length; ki++) {
      if (keys[ki].toLowerCase().indexOf(cand) !== -1) return row[keys[ki]];
    }
  }
  return "";
}

function parseIBMFile(wb) {
  var sheetName = findSheet(wb, "labor claim only details");
  if (!sheetName) {
    // Fallback: scan each sheet's header row for IBM-specific column signatures
    var IBM_SIGNALS = ["hours performed", "billing code", "claim month", "workitem"];
    for (var si = 0; si < wb.SheetNames.length; si++) {
      var ws_try = wb.Sheets[wb.SheetNames[si]];
      var rows_try = XLSX.utils.sheet_to_json(ws_try, { defval: "", header: 1 });
      if (!rows_try.length) continue;
      var headerRow = rows_try[0].map(function(h){ return String(h).toLowerCase(); });
      var matchCount = IBM_SIGNALS.filter(function(sig){
        return headerRow.some(function(h){ return h.indexOf(sig) !== -1; });
      }).length;
      if (matchCount >= 2) { sheetName = wb.SheetNames[si]; break; }
    }
  }
  if (!sheetName) return { error: "Could not find sheet containing 'Labor claim only details'. Sheets: " + wb.SheetNames.join(", ") };
  var ws = wb.Sheets[sheetName];
  var rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  if (!rows.length) return { error: "Sheet '" + sheetName + "' is empty." };

  function claimMonthLabel(cm) {
    var n = parseInt(cm);
    var MN = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    if (!isNaN(n) && n >= 1 && n <= 12) return MN[n-1];
    return String(cm);
  }

  // Keywords that identify summary/total rows — not real people
  var SKIP_NAME_PATTERNS = ["grand total", "total", "sub total", "subtotal", "summary"];
  var byName = {};
  rows.forEach(function(row) {
    var name = String(getCol(row, ["name"]) || "").trim();
    if (!name) return;
    // Skip summary/total rows
    var nameLower = name.toLowerCase();
    if (SKIP_NAME_PATTERNS.some(function(p){ return nameLower === p || nameLower.indexOf(p) === 0; })) return;
    var key = normalizeName(name);
    if (!byName[key]) {
      byName[key] = {
        rawName: name, normalizedName: key,
        email: String(getCol(row, ["internet address"]) || ""),
        talentId: String(getCol(row, ["talentid (cnum)","cnum"]) || ""),
        serialId: String(getCol(row, ["resource talentid (serial)","serial"]) || ""),
        country: String(getCol(row, ["resource country name","country"]) || ""),
        billingCode: String(getCol(row, ["billing code"]) || ""),
        wbsId: String(getCol(row, ["ippf customer wbs id","wbs"]) || ""),
        activityCode: String(getCol(row, ["activity code"]) || ""),
        scheduledHours: 0, workitems: [], claimMonths: [],
        satHrs:0, sunHrs:0, monHrs:0, tueHrs:0, wedHrs:0, thuHrs:0, friHrs:0,
        weeklyBreakdown: [],
      };
    }
    var hrs = parseFloat(getCol(row, ["hours performed"]) || 0) || 0;
    var sat = parseFloat(getCol(row, ["sat hours","sat"]) || 0) || 0;
    var sun = parseFloat(getCol(row, ["sun hours","sun"]) || 0) || 0;
    var mon = parseFloat(getCol(row, ["mon hours","mon"]) || 0) || 0;
    var tue = parseFloat(getCol(row, ["tue hours","tue"]) || 0) || 0;
    var wed = parseFloat(getCol(row, ["wed hours","wed"]) || 0) || 0;
    var thu = parseFloat(getCol(row, ["thu hours","thu"]) || 0) || 0;
    var fri = parseFloat(getCol(row, ["fri hours","fri"]) || 0) || 0;
    var we  = String(getCol(row, ["hours performed for w/e","w/e","week end","week ending"]) || "").trim();
    var cm  = String(getCol(row, ["claim month"]) || "").trim();
    var wi  = String(getCol(row, ["workitem title","workitem"]) || "").trim();
    var ac  = String(getCol(row, ["activity code"]) || "").trim();
    byName[key].scheduledHours += hrs;
    byName[key].satHrs += sat;
    byName[key].sunHrs += sun;
    byName[key].monHrs += mon;
    byName[key].tueHrs += tue;
    byName[key].wedHrs += wed;
    byName[key].thuHrs += thu;
    byName[key].friHrs += fri;
    // Extract year from weekEnd date (format "DD-MM-YYYY" or "YYYY-MM-DD")
    var weYear = null;
    if (we) {
      var weMatch = we.match(/(\d{4})/);
      if (weMatch) weYear = parseInt(weMatch[1]);
    }
    // Also try year from claim month field if it contains a year (e.g. "2026-02" or "Feb-2026")
    if (!weYear && cm) {
      var cmYearMatch = String(cm).match(/(\d{4})/);
      if (cmYearMatch) weYear = parseInt(cmYearMatch[1]);
    }
    var cml = claimMonthLabel(cm);
    var MFULL_MAP = {Jan:"January",Feb:"February",Mar:"March",Apr:"April",May:"May",Jun:"June",Jul:"July",Aug:"August",Sep:"September",Oct:"October",Nov:"November",Dec:"December"};
    var cmlFull = MFULL_MAP[cml] || cml;
    // Build month-year key like "February-2026"
    // If cml is empty, try extracting month from weekEnd date
    if (!cml && we) {
      var weParts = we.replace(/[^0-9\-\/]/g,"").split(/[\-\/]/);
      // Try both DD-MM-YYYY and YYYY-MM-DD
      var weMonthNum = null;
      if (weParts.length >= 3) {
        if (weParts[0].length === 4) weMonthNum = parseInt(weParts[1]); // YYYY-MM-DD
        else weMonthNum = parseInt(weParts[1]);                          // DD-MM-YYYY
      }
      if (weMonthNum && weMonthNum >= 1 && weMonthNum <= 12) {
        var MFULL2 = ["January","February","March","April","May","June","July","August","September","October","November","December"];
        cmlFull = MFULL2[weMonthNum - 1];
      }
    }
    var monthYearKey = (cmlFull && weYear) ? (cmlFull + "-" + weYear) : (cmlFull || "");
    byName[key].weeklyBreakdown.push({ weekEnd:we, total:hrs, sat:sat, sun:sun, mon:mon, tue:tue, wed:wed, thu:thu, fri:fri, claimMonth:cml, claimYear:weYear, monthYearKey:monthYearKey, workitem:wi, activityCode:ac });
    if (wi && byName[key].workitems.indexOf(wi) === -1) byName[key].workitems.push(wi);
    // Store month-year keyed scheduled hours (like Clarity's monthlyHours)
    if (monthYearKey) {
      byName[key].ibmMonthlyHours = byName[key].ibmMonthlyHours || {};
      byName[key].ibmMonthlyHours[monthYearKey] = (byName[key].ibmMonthlyHours[monthYearKey] || 0) + hrs;
    }
    if (cm) { if (byName[key].claimMonths.indexOf(monthYearKey) === -1) byName[key].claimMonths.push(monthYearKey); }
    if (!byName[key].email) byName[key].email = String(getCol(row, ["internet address"]) || "");
    if (!byName[key].wbsId) byName[key].wbsId = String(getCol(row, ["ippf customer wbs id","wbs"]) || "");
    if (!byName[key].billingCode) byName[key].billingCode = String(getCol(row, ["billing code"]) || "");
    if (!byName[key].country) byName[key].country = String(getCol(row, ["resource country name","country"]) || "");
    if (!byName[key].talentId) byName[key].talentId = String(getCol(row, ["talentid (cnum)","cnum"]) || "");
    if (!byName[key].serialId) byName[key].serialId = String(getCol(row, ["resource talentid (serial)","serial"]) || "");
    if (!byName[key].activityCode && ac) byName[key].activityCode = ac;
  });
  Object.values(byName).forEach(function(r) {
    r.weeklyBreakdown.sort(function(a,b){ return String(b.weekEnd).localeCompare(String(a.weekEnd)); });
  });
  return { sheetName: sheetName, rowCount: rows.length, records: Object.values(byName), error: null };
}
function extractMonthFromSheetName(sheetName) {
  // e.g. "CORP_AML_FCU_Feb2026_Actual hrs" -> {month:"February", year:2026, label:"February 2026"}
  var MONTHS = [
    ["jan","January"],["feb","February"],["mar","March"],["apr","April"],
    ["may","May"],["jun","June"],["jul","July"],["aug","August"],
    ["sep","September"],["oct","October"],["nov","November"],["dec","December"]
  ];
  var s = sheetName.toLowerCase();
  var monthName = null, year = null;
  for (var mi = 0; mi < MONTHS.length; mi++) {
    if (s.indexOf(MONTHS[mi][0]) !== -1) { monthName = MONTHS[mi][1]; break; }
  }
  var yearMatch = sheetName.match(/20\d{2}/);
  if (yearMatch) year = parseInt(yearMatch[0]);
  if (!monthName && !year) return null;
  return {
    month: monthName || "Unknown",
    year:  year || new Date().getFullYear(),
    label: (monthName || "Unknown") + " " + (year || new Date().getFullYear()),
    key:   (monthName || "Unknown") + "-" + (year || new Date().getFullYear()),
  };
}

function parseClarityFile(wb) {
  var sheetName = findSheet(wb, "actual hr");
  if (!sheetName) sheetName = findSheet(wb, "actual");
  if (!sheetName) sheetName = wb.SheetNames[0];
  var ws = wb.Sheets[sheetName];

  // Read raw rows to detect dual-period layout (two sets of columns side by side)
  var rawRows = XLSX.utils.sheet_to_json(ws, { defval: "", header: 1 });
  if (!rawRows.length) return { error: "Sheet '" + sheetName + "' is empty." };

  // Find header row (first row containing "resource name")
  var headerRowIdx = 0;
  for (var ri = 0; ri < Math.min(rawRows.length, 5); ri++) {
    var rh = rawRows[ri].map(function(c){ return String(c).toLowerCase(); });
    if (rh.some(function(c){ return c.indexOf("resource name") !== -1; })) { headerRowIdx = ri; break; }
  }
  var headers = rawRows[headerRowIdx].map(function(c){ return String(c).toLowerCase().trim(); });

  // Detect all column indices for key fields (handles duplicate headers = multiple periods)
  function allIndicesOf(headers, keyword) {
    var idxs = [];
    headers.forEach(function(h, i){ if (h.indexOf(keyword) !== -1) idxs.push(i); });
    return idxs;
  }
  var nameIdxs       = allIndicesOf(headers, "resource name");
  var periodIdxs     = allIndicesOf(headers, "time reporting period");
  var hoursIdxs      = allIndicesOf(headers, "total actual hours").concat(allIndicesOf(headers, "total hours"));
  // Deduplicate hoursIdxs
  hoursIdxs = hoursIdxs.filter(function(v,i,a){ return a.indexOf(v)===i; });
  var statusIdxs     = allIndicesOf(headers, "timesheet status");
  var mgrIdxs        = allIndicesOf(headers, "resource manager");
  var approvedIdxs   = allIndicesOf(headers, "approved by");
  var activeIdxs     = allIndicesOf(headers, "resource active");

  // Pair up period columns with their corresponding hours columns by proximity
  // Each "block" = { periodIdx, hoursIdx }
  var blocks = [];
  periodIdxs.forEach(function(pIdx) {
    // Find the nearest hours column to the right of this period column
    var nearest = null, nearestDist = 999;
    hoursIdxs.forEach(function(hIdx) {
      var dist = hIdx - pIdx;
      if (dist > 0 && dist < nearestDist) { nearest = hIdx; nearestDist = dist; }
    });
    if (nearest !== null) blocks.push({ periodIdx: pIdx, hoursIdx: nearest });
  });
  // If no blocks found fall back to first period+hours pair
  if (!blocks.length && periodIdxs.length && hoursIdxs.length) {
    blocks.push({ periodIdx: periodIdxs[0], hoursIdx: hoursIdxs[0] });
  }

  var monthInfo = extractMonthFromSheetName(sheetName);
  var MFULL_P = ["January","February","March","April","May","June","July","August","September","October","November","December"];

  function periodToMonthKey(periodStr) {
    if (!periodStr) return null;
    var dateMatch = periodStr.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (dateMatch) {
      var pMonth = parseInt(dateMatch[1]), pYear = parseInt(dateMatch[3]);
      if (pMonth >= 1 && pMonth <= 12) return MFULL_P[pMonth-1] + "-" + pYear;
    }
    var yearM = periodStr.match(/(\d{4})/);
    if (yearM) {
      var py = parseInt(yearM[1]), pl = periodStr.toLowerCase();
      for (var mi = 0; mi < MFULL_P.length; mi++) {
        if (pl.indexOf(MFULL_P[mi].toLowerCase()) !== -1) return MFULL_P[mi] + "-" + py;
      }
    }
    return null;
  }

  var byName = {};
  var nameIdx = nameIdxs.length ? nameIdxs[0] : -1;
  var mgrIdx     = mgrIdxs.length ? mgrIdxs[0] : -1;
  var statusIdx  = statusIdxs.length ? statusIdxs[statusIdxs.length-1] : -1; // last status = most recent
  var approvedIdx= approvedIdxs.length ? approvedIdxs[0] : -1;
  var activeIdx  = activeIdxs.length ? activeIdxs[0] : -1;

  for (var rowi = headerRowIdx + 1; rowi < rawRows.length; rowi++) {
    var row = rawRows[rowi];
    var name = nameIdx >= 0 ? String(row[nameIdx] || "").trim() : "";
    if (!name) continue;
    var key = normalizeName(name);
    if (!byName[key]) {
      byName[key] = {
        rawName: name, normalizedName: key,
        resourceManager: mgrIdx >= 0 ? String(row[mgrIdx] || "") : "",
        timesheetStatus: statusIdx >= 0 ? String(row[statusIdx] || "") : "",
        approvedBy: approvedIdx >= 0 ? String(row[approvedIdx] || "") : "",
        resourceActive: activeIdx >= 0 ? String(row[activeIdx] || "") : "",
        actualHours: 0, periods: [], monthlyHours: {},
      };
    }
    // Accumulate hours from ALL period blocks (handles dual half-month layout)
    blocks.forEach(function(blk) {
      var hrs = parseFloat(row[blk.hoursIdx] || 0) || 0;
      if (!hrs) return;
      var periodStr = String(row[blk.periodIdx] || "").trim();
      byName[key].actualHours += hrs;
      if (periodStr && byName[key].periods.indexOf(periodStr) === -1) byName[key].periods.push(periodStr);
      var mkey = periodToMonthKey(periodStr) || (monthInfo ? monthInfo.key : "Unknown");
      byName[key].monthlyHours[mkey] = (byName[key].monthlyHours[mkey] || 0) + hrs;
    });
    // If no blocks matched, fall back to plain sheet_to_json reading
    if (!blocks.length) {
      var hrs = parseFloat(String(row[hoursIdxs[0]] || 0)) || 0;
      byName[key].actualHours += hrs;
    }
    // Update latest status/manager/approver if present in this row
    if (statusIdx >= 0 && row[statusIdx]) byName[key].timesheetStatus = String(row[statusIdx]).trim();
    if (mgrIdx >= 0 && row[mgrIdx]) byName[key].resourceManager = String(row[mgrIdx]).trim();
    if (approvedIdx >= 0 && row[approvedIdx]) byName[key].approvedBy = String(row[approvedIdx]).trim();
    if (activeIdx >= 0 && row[activeIdx]) byName[key].resourceActive = String(row[activeIdx]).trim();
  }
  return {
    sheetName: sheetName,
    rowCount: rawRows.length - headerRowIdx - 1,
    records: Object.values(byName),
    monthInfo: monthInfo,
    error: null,
  };
}

// ─── FUZZY MATCHING ENGINE ────────────────────────────────────────────────────
function editDistance(a, b) {
  if (a === b) return 0;
  var la = a.length, lb = b.length;
  if (la === 0) return lb;
  if (lb === 0) return la;
  var dp = [];
  for (var j = 0; j <= lb; j++) dp[j] = j;
  for (var i = 1; i <= la; i++) {
    var ndp = [i];
    for (var j2 = 1; j2 <= lb; j2++) {
      var cost = a[i-1] === b[j2-1] ? 0 : 1;
      ndp[j2] = Math.min(ndp[j2-1]+1, dp[j2]+1, dp[j2-1]+cost);
    }
    dp = ndp;
  }
  return dp[lb];
}

// Check if all tokens of a spaced name form (concatenate into) a compound word
// e.g. ["vijaya","lakshmi"] -> "vijayalakshmi" covers >=85% of compound chars
function tokensFormCompound(tokens, compound) {
  if (!tokens.every(function(t){ return compound.indexOf(t) !== -1; })) return false;
  var temp = compound;
  var sortedTokens = tokens.slice().sort(function(a,b){
    return compound.indexOf(a) - compound.indexOf(b);
  });
  for (var i = 0; i < sortedTokens.length; i++) {
    var t = sortedTokens[i];
    var idx = temp.indexOf(t);
    if (idx === -1) return false;
    temp = temp.slice(0, idx) + Array(t.length+1).join("_") + temp.slice(idx+t.length);
  }
  var covered = tokens.reduce(function(s,t){ return s+t.length; }, 0);
  return covered >= compound.length * 0.85;
}

function fuzzyMatchScore(nameA, nameB) {
  var na = normalizeName(nameA);
  var nb = normalizeName(nameB);
  if (!na || !nb) return 0;
  if (na === nb) return 1.0;

  var ta = na.split(" "); var tb = nb.split(" ");
  var setA = {}; ta.forEach(function(t){ setA[t]=1; });
  var setB = {}; tb.forEach(function(t){ setB[t]=1; });
  var scores = [];

  // 1. Token subset: all tokens of shorter appear in longer
  var inAB = ta.filter(function(t){ return setB[t]; }).length;
  var inBA = tb.filter(function(t){ return setA[t]; }).length;
  if (inAB === ta.length || inBA === tb.length) {
    var ratio = Math.min(ta.length, tb.length) / Math.max(ta.length, tb.length);
    scores.push(0.85 + ratio * 0.1);
  }

  var nans = na.replace(/ /g, "");
  var nbns = nb.replace(/ /g, "");
  if (nans === nbns) return 0.95;

  // 2. Compound name formation (Vijaya Lakshmi <-> Vijayalakshmi)
  if (ta.length > 1 && tokensFormCompound(ta, nbns)) scores.push(0.92);
  if (tb.length > 1 && tokensFormCompound(tb, nans)) scores.push(0.92);
  // Reverse: compound splits into tokens (Radhakrishnan <-> Radha Krishnan)
  if (ta.length === 1 && tb.length > 1 && tb.every(function(t){ return nans.indexOf(t) !== -1; })) scores.push(0.90);
  if (tb.length === 1 && ta.length > 1 && ta.every(function(t){ return nbns.indexOf(t) !== -1; })) scores.push(0.90);

  // 3. Token overlap (Jaccard)
  var inter = ta.filter(function(t){ return setB[t]; });
  if (inter.length > 0) {
    var unionLen = ta.length + tb.length - inter.length;
    var jaccard = inter.length / unionLen;
    var sub = inter.length / Math.min(ta.length, tb.length);
    scores.push(Math.max(jaccard, sub * 0.85));
  }

  // 4. Edit distance on full normalized string (with spaces) — handles spelling variants
  var ed1 = editDistance(na, nb);
  scores.push((1 - ed1 / Math.max(na.length, nb.length, 1)) * 0.95);

  // 5. Edit distance on no-space concatenated string
  var ed2 = editDistance(nans, nbns);
  scores.push((1 - ed2 / Math.max(nans.length, nbns.length, 1)) * 0.95);

  var best = scores.length > 0 ? Math.max.apply(null, scores) : 0;

  // Penalty: single very short shared token between different names (false positive guard)
  if (best < 0.7 && inter.length === 1 && inter[0].length <= 3) {
    best = Math.min(best, 0.45);
  }

  return best;
}


// mergeRecords: returns { matched, autoMatched, ibmOnly, clarityOnly, suggestions }
// suggestions: ibmOnly records with their best clarity candidates ranked by score
function mergeRecords(ibmRecords, clarityRecords, manualMatches) {
  // manualMatches: { [ibmNormName]: clarityNormName }
  var manualMap = manualMatches || {};
  var clarityMap = {};
  clarityRecords.forEach(function(r) { clarityMap[r.normalizedName] = r; });
  var ibmMap = {};
  ibmRecords.forEach(function(r) { ibmMap[r.normalizedName] = r; });

  var matched = [];
  var ibmOnly = [];
  var clarityOnly = [];
  var usedClarityKeys = {};

  function buildRecord(ibm, c) {
    return {
      id: ibm.normalizedName,
      name: ibm.rawName,
      normalizedName: ibm.normalizedName,
      clarityName: c ? c.rawName : null,
      email: ibm.email,
      talentId: ibm.talentId,
      serialId: ibm.serialId,
      country: ibm.country,
      billingCode: ibm.billingCode,
      wbsId: ibm.wbsId,
      activityCode: ibm.activityCode || "",
      workitems: ibm.workitems,
      claimMonths: ibm.claimMonths,
      scheduledHours: ibm.scheduledHours,
      dayHours: { sat:ibm.satHrs||0, sun:ibm.sunHrs||0, mon:ibm.monHrs||0, tue:ibm.tueHrs||0, wed:ibm.wedHrs||0, thu:ibm.thuHrs||0, fri:ibm.friHrs||0 },
      weeklyBreakdown: ibm.weeklyBreakdown || [],
      ibmMonthlyHours: ibm.ibmMonthlyHours || {},
      monthlyHours: c ? (c.monthlyHours || {}) : {},
      actualHours: c ? c.actualHours : 0,
      resourceManager: c ? c.resourceManager : "",
      timesheetStatus: c ? c.timesheetStatus : "Not in Clarity",
      approvedBy: c ? c.approvedBy : "",
      resourceActive: c ? c.resourceActive : "",
      periods: c ? c.periods : [],
      clarityPeriods: c ? c.periods : [],
      matched: !!c,
      dataSource: c ? "Both" : "IBM only",
    };
  }

  ibmRecords.forEach(function(ibm) {
    // 1. Exact match
    var c = clarityMap[ibm.normalizedName];
    if (c) { usedClarityKeys[ibm.normalizedName] = true; matched.push(buildRecord(ibm, c)); return; }

    // 2. Manual match override
    var manualKey = manualMap[ibm.normalizedName];
    if (manualKey && clarityMap[manualKey]) {
      c = clarityMap[manualKey];
      usedClarityKeys[manualKey] = true;
      matched.push(buildRecord(ibm, c));
      return;
    }

    // 3. Auto fuzzy match (score >= 0.85)
    var bestScore = 0, bestClarity = null, bestKey = null;
    clarityRecords.forEach(function(cr) {
      if (usedClarityKeys[cr.normalizedName]) return;
      var score = fuzzyMatchScore(ibm.rawName, cr.rawName);
      if (score > bestScore) { bestScore = score; bestClarity = cr; bestKey = cr.normalizedName; }
    });
    if (bestScore >= 0.85 && bestClarity) {
      usedClarityKeys[bestKey] = true;
      var rec = buildRecord(ibm, bestClarity);
      rec.autoMatchScore = bestScore;
      rec.dataSource = "Both";
      matched.push(rec);
      return;
    }

    ibmOnly.push(buildRecord(ibm, null));
  });

  // Clarity records not matched to any IBM
  clarityRecords.forEach(function(c) {
    if (!usedClarityKeys[c.normalizedName]) {
      clarityOnly.push({
        id: c.normalizedName, name: c.rawName, normalizedName: c.normalizedName,
        clarityName: c.rawName, email:"", talentId:"", serialId:"", country:"",
        billingCode:"", wbsId:"", activityCode:"", workitems:[], claimMonths:[],
        scheduledHours:0, dayHours:{sat:0,sun:0,mon:0,tue:0,wed:0,thu:0,fri:0},
        weeklyBreakdown:[], monthlyHours:c.monthlyHours||{}, actualHours:c.actualHours, resourceManager:c.resourceManager,
        timesheetStatus:c.timesheetStatus, approvedBy:c.approvedBy,
        resourceActive:c.resourceActive, periods:c.periods, clarityPeriods:c.periods,
        matched:false, dataSource:"Clarity only",
      });
    }
  });

  // Build fuzzy suggestions for unmatched IBM records
  var suggestions = {};
  ibmOnly.forEach(function(ibmRec) {
    var candidates = [];
    clarityOnly.forEach(function(cRec) {
      var score = fuzzyMatchScore(ibmRec.name, cRec.name);
      if (score >= 0.4) candidates.push({ clarityName: cRec.name, clarityKey: cRec.normalizedName, score: score });
    });
    candidates.sort(function(a,b){ return b.score - a.score; });
    if (candidates.length > 0) suggestions[ibmRec.normalizedName] = candidates.slice(0, 5);
  });

  return { matched: matched, ibmOnly: ibmOnly, clarityOnly: clarityOnly, suggestions: suggestions };
}


// ─── IMPORT MODAL ─────────────────────────────────────────────────────────────
function ImportModal({onImport, onClose}) {
  // IBM: array of {file, result, error, id}  — supports multiple files
  const[ibmFiles, setIbmFiles] = useState([]);
  const[clarityFiles, setClarityFiles] = useState([]); // array of {id,file,result,error,loading}
  const[dragOver, setDragOver] = useState(null);
  const[manualMatches, setManualMatches] = useState({}); // {ibmNormName: clarityNormName}
  const[showManualUI, setShowManualUI] = useState(false);
  const ibmRef = useRef();
  const clarityRef = useRef();

  var MAX_FILE_SIZE_MB = 50;
  function readIBMFile(file) {
    if (file.size > MAX_FILE_SIZE_MB * 1024 * 1024) {
      alert("File too large. Maximum size is " + MAX_FILE_SIZE_MB + "MB.");
      return;
    }
    // Validate file extension
    var allowedExts = [".xlsx",".xls",".xlsm",".xlsb",".xltx",".xltm",".xlt",".csv",".ods",".tsv"];
    var fname = file.name.toLowerCase();
    var isAllowed = allowedExts.some(function(ext){ return fname.endsWith(ext); });
    if (!isAllowed) { alert("File type not allowed. Please upload a spreadsheet file."); return; }
    var fileId = Date.now() + Math.random();
    // Add placeholder immediately
    setIbmFiles(function(prev) { return prev.concat([{id:fileId, file:file, result:null, error:"", loading:true}]); });
    var reader = new FileReader();
    reader.onload = function(ev) {
      try {
        var wb = XLSX.read(ev.target.result, {type:"array"});
        var result = parseIBMFile(wb);
        setIbmFiles(function(prev) {
          return prev.map(function(f) {
            return f.id === fileId ? {id:fileId, file:file, result:result.error?null:result, error:result.error||"", loading:false} : f;
          });
        });
      } catch(ex) {
        setIbmFiles(function(prev) {
          return prev.map(function(f) {
            return f.id === fileId ? {id:fileId, file:file, result:null, error:"Could not parse: "+(ex.message||ex), loading:false} : f;
          });
        });
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function readClarityFile(file) {
    var fileId = Date.now() + Math.random();
    setClarityFiles(function(prev){ return prev.concat([{id:fileId, file:file, result:null, error:"", loading:true}]); });
    var reader = new FileReader();
    reader.onload = function(ev) {
      try {
        var wb = XLSX.read(ev.target.result, {type:"array"});
        var result = parseClarityFile(wb);
        setClarityFiles(function(prev){
          return prev.map(function(f){
            if (f.id !== fileId) return f;
            return result.error
              ? {id:fileId, file:file, result:null, error:result.error, loading:false}
              : {id:fileId, file:file, result:result, error:"", loading:false};
          });
        });
      } catch(ex) {
        setClarityFiles(function(prev){
          return prev.map(function(f){
            return f.id===fileId ? {id:fileId, file:file, result:null, error:"Could not parse: "+(ex.message||ex), loading:false} : f;
          });
        });
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function removeClarityFile(id) {
    setClarityFiles(function(prev){ return prev.filter(function(f){ return f.id!==id; }); });
  }

  function mergeAllClarityRecords(filesList) {
    var combined = {};
    filesList.forEach(function(cf) {
      if (!cf.result) return;
      cf.result.records.forEach(function(r) {
        var key = r.normalizedName;
        if (!combined[key]) {
          combined[key] = JSON.parse(JSON.stringify(r));
          combined[key].sourceMonths = [];
        } else {
          combined[key].actualHours += r.actualHours;
          r.periods.forEach(function(p){ if(combined[key].periods.indexOf(p)===-1) combined[key].periods.push(p); });
          Object.keys(r.monthlyHours||{}).forEach(function(mk){
            combined[key].monthlyHours = combined[key].monthlyHours || {};
            combined[key].monthlyHours[mk] = (combined[key].monthlyHours[mk]||0) + r.monthlyHours[mk];
          });
          if (!combined[key].resourceManager && r.resourceManager) combined[key].resourceManager = r.resourceManager;
          if (!combined[key].approvedBy && r.approvedBy) combined[key].approvedBy = r.approvedBy;
          if (r.timesheetStatus) combined[key].timesheetStatus = r.timesheetStatus;
        }
        var mLabel = cf.result.monthInfo ? cf.result.monthInfo.label : "Unknown";
        if (combined[key].sourceMonths.indexOf(mLabel)===-1) combined[key].sourceMonths.push(mLabel);
      });
    });
    return Object.values(combined);
  }

  function removeIBMFile(id) {
    setIbmFiles(function(prev) { return prev.filter(function(f) { return f.id !== id; }); });
  }

  // Merge all IBM file records together (union by person, summing hours across WBS files)
  function mergeAllIBMRecords(ibmFilesList) {
    var combined = {};
    ibmFilesList.forEach(function(ibmFile) {
      if (!ibmFile.result) return;
      ibmFile.result.records.forEach(function(r) {
        var key = r.normalizedName;
        if (!combined[key]) {
          combined[key] = JSON.parse(JSON.stringify(r)); // deep copy
          combined[key].sourceFiles = [ibmFile.file.name];
        } else {
          // Merge: sum hours, combine workitems/claimMonths/weeklyBreakdown
          combined[key].scheduledHours += r.scheduledHours;
          combined[key].satHrs = (combined[key].satHrs||0) + (r.satHrs||0);
          combined[key].sunHrs = (combined[key].sunHrs||0) + (r.sunHrs||0);
          combined[key].monHrs = (combined[key].monHrs||0) + (r.monHrs||0);
          combined[key].tueHrs = (combined[key].tueHrs||0) + (r.tueHrs||0);
          combined[key].wedHrs = (combined[key].wedHrs||0) + (r.wedHrs||0);
          combined[key].thuHrs = (combined[key].thuHrs||0) + (r.thuHrs||0);
          combined[key].friHrs = (combined[key].friHrs||0) + (r.friHrs||0);
          r.workitems.forEach(function(w) { if (combined[key].workitems.indexOf(w) === -1) combined[key].workitems.push(w); });
          r.claimMonths.forEach(function(m) { if (combined[key].claimMonths.indexOf(m) === -1) combined[key].claimMonths.push(m); });
          combined[key].weeklyBreakdown = combined[key].weeklyBreakdown.concat(r.weeklyBreakdown);
          // Keep most complete metadata
          if (!combined[key].email && r.email) combined[key].email = r.email;
          if (!combined[key].talentId && r.talentId) combined[key].talentId = r.talentId;
          if (!combined[key].country && r.country) combined[key].country = r.country;
          combined[key].sourceFiles.push(ibmFile.file.name);
          // Combine WBS IDs if different
          if (r.wbsId && combined[key].wbsId && r.wbsId !== combined[key].wbsId) {
            combined[key].wbsId = combined[key].wbsId + ", " + r.wbsId;
          } else if (r.wbsId && !combined[key].wbsId) {
            combined[key].wbsId = r.wbsId;
          }
        }
      });
    });
    // Re-sort weekly breakdown by date desc
    Object.values(combined).forEach(function(r) {
      r.weeklyBreakdown.sort(function(a,b){ return String(b.weekEnd).localeCompare(String(a.weekEnd)); });
    });
    return Object.values(combined);
  }

  var loadedIBMFiles = ibmFiles.filter(function(f) { return f.result !== null; });
  var allIBMRecords = mergeAllIBMRecords(loadedIBMFiles);
  var loadedClarityFiles = clarityFiles.filter(function(f){ return f.result !== null; });
  var clarityRecs = mergeAllClarityRecords(loadedClarityFiles);
  var merged = (allIBMRecords.length > 0 || clarityRecs.length > 0) ? mergeRecords(allIBMRecords, clarityRecs, manualMatches) : null;
  var totalImport = merged ? (merged.matched.length + merged.ibmOnly.length + merged.clarityOnly.length) : 0;
  var availableMonths = [];
  loadedClarityFiles.forEach(function(cf){
    if (cf.result && cf.result.monthInfo) {
      var lbl = cf.result.monthInfo.label;
      if (availableMonths.indexOf(lbl)===-1) availableMonths.push(lbl);
    }
  });

  // Total scheduled across all IBM files
  var totalIBMHours = allIBMRecords.reduce(function(s,r){ return s + (r.scheduledHours||0); }, 0);
  var totalIBMPeople = allIBMRecords.length;
  var wbsList = [];
  loadedIBMFiles.forEach(function(f) {
    if (f.result) f.result.records.forEach(function(r) {
      if (r.wbsId) r.wbsId.split(", ").forEach(function(w) { if (w && wbsList.indexOf(w) === -1) wbsList.push(w); });
    });
  });

  function handleImportClick() {
    if (!merged) return;
    var all = merged.matched.concat(merged.ibmOnly).concat(merged.clarityOnly);
    onImport(all, merged);
    onClose();
  }

  function renderIBMZone() {
    var isDrag = dragOver === "ibm";
    return (
      <div style={{border:"1px solid "+IBM.gray20, marginBottom:16}}>
        {/* Header */}
        <div style={{background:loadedIBMFiles.length>0?IBM.blue70:IBM.blue60, color:"#fff", padding:"10px 16px", display:"flex", justifyContent:"space-between", alignItems:"center"}}>
          <div>
            <div style={{fontSize:13, fontWeight:700}}>
              IBM Scheduled Data
              {loadedIBMFiles.length > 0 && <span style={{marginLeft:8, background:"rgba(255,255,255,0.2)", padding:"1px 8px", fontSize:11, borderRadius:10}}>{loadedIBMFiles.length} file{loadedIBMFiles.length>1?"s":""}</span>}
            </div>
            <div style={{fontSize:10, opacity:0.85, marginTop:2}}>Sheet: "Labor claim only details" (fuzzy match) — upload multiple files for different WBS IDs</div>
          </div>
          <button
            onClick={function(){ ibmRef.current.click(); }}
            style={{background:"rgba(255,255,255,0.15)", border:"1px solid rgba(255,255,255,0.4)", color:"#fff", padding:"5px 12px", cursor:"pointer", fontSize:11, fontWeight:600, whiteSpace:"nowrap"}}>
            + Add File
          </button>
        </div>
        <input ref={ibmRef} type="file" accept=".xlsx,.xls,.xlsm,.xlsb,.xltx,.xltm,.xlt,.csv,.ods,.tsv" multiple style={{display:"none"}}
          onChange={function(e){ Array.from(e.target.files).forEach(function(f){ readIBMFile(f); }); e.target.value=""; }}/>

        {/* Loaded files list */}
        {ibmFiles.length > 0 && (
          <div style={{background:"#fafafa", borderBottom:"1px solid "+IBM.gray20}}>
            {ibmFiles.map(function(f) {
              return (
                <div key={f.id} style={{display:"flex", alignItems:"center", gap:10, padding:"8px 14px", borderBottom:"1px solid "+IBM.gray20}}>
                  <span style={{fontSize:14}}>{f.loading ? "⏳" : f.error ? "⚠️" : "✅"}</span>
                  <div style={{flex:1, minWidth:0}}>
                    <div style={{fontSize:12, fontWeight:600, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap"}}>{f.file.name}</div>
                    {f.result && (
                      <div style={{fontSize:11, color:IBM.gray60, marginTop:1}}>
                        Sheet: {f.result.sheetName} &nbsp;&#8226;&nbsp; {f.result.rowCount} rows &nbsp;&#8226;&nbsp; {f.result.records.length} people &nbsp;&#8226;&nbsp;
                        <span style={{color:IBM.blue60, fontWeight:600}}>
                          {f.result.records.reduce(function(s,r){ return s+(r.scheduledHours||0); },0)}h total
                        </span>
                        {(function(){
                          var wbs = [];
                          f.result.records.forEach(function(r){ if(r.wbsId) r.wbsId.split(", ").forEach(function(w){ if(w&&wbs.indexOf(w)===-1)wbs.push(w); }); });
                          return wbs.length > 0 ? <span style={{marginLeft:6, color:IBM.gray50}}> WBS: {wbs.join(", ")}</span> : null;
                        })()}
                      </div>
                    )}
                    {f.error && <div style={{fontSize:11, color:IBM.red60, marginTop:1}}>{f.error}</div>}
                    {f.loading && <div style={{fontSize:11, color:IBM.gray50, marginTop:1}}>Parsing...</div>}
                  </div>
                  <button onClick={function(){ removeIBMFile(f.id); }}
                    style={{background:"none", border:"1px solid "+IBM.gray30, color:IBM.gray60, padding:"2px 8px", cursor:"pointer", fontSize:11, flexShrink:0}}>
                    Remove
                  </button>
                </div>
              );
            })}
          </div>
        )}

        {/* Drop zone */}
        <div
          onDragOver={function(e){e.preventDefault();setDragOver("ibm");}}
          onDragLeave={function(){setDragOver(null);}}
          onDrop={function(e){e.preventDefault();setDragOver(null);Array.from(e.dataTransfer.files).forEach(function(f){readIBMFile(f);});}}
          onClick={function(){ibmRef.current.click();}}
          style={{padding:ibmFiles.length>0?"12px 20px":"20px", background:isDrag?IBM.blue10:"#fff", cursor:"pointer", textAlign:"center", borderTop:ibmFiles.length>0?"1px dashed "+IBM.gray30:"none"}}>
          <div style={{fontSize:isDrag?28:20, marginBottom:4}}>📊</div>
          <div style={{fontSize:12, color:isDrag?IBM.blue60:IBM.gray60}}>
            {ibmFiles.length > 0 ? "Drop another IBM file here or click to add more" : "Drag and drop IBM file(s) here or click to browse"}
          </div>
          <div style={{fontSize:10, color:IBM.gray40, marginTop:3}}>.xlsx, .xls, .xlsm, .xlsb, .csv, .ods and more &nbsp;&#8226;&nbsp; Multiple files allowed</div>
        </div>

        {/* Merged summary across all loaded files */}
        {loadedIBMFiles.length > 0 && (
          <div style={{padding:"10px 14px", background:IBM.blue10, borderTop:"1px solid "+IBM.blue20, display:"flex", gap:16, flexWrap:"wrap", alignItems:"center"}}>
            <span style={{fontSize:11, color:IBM.blue70}}><b>Combined:</b></span>
            <span style={{fontSize:11, background:"#fff", color:IBM.blue60, padding:"2px 8px", border:"1px solid "+IBM.blue20, fontWeight:600}}>{totalIBMPeople} people</span>
            <span style={{fontSize:11, background:"#fff", color:IBM.blue60, padding:"2px 8px", border:"1px solid "+IBM.blue20, fontWeight:600}}>{totalIBMHours}h scheduled</span>
            {wbsList.length > 0 && wbsList.map(function(w) {
              return <span key={w} style={{fontSize:11, background:"#fff", color:IBM.gray70, padding:"2px 8px", border:"1px solid "+IBM.gray20}}>WBS: {w}</span>;
            })}
          </div>
        )}
      </div>
    );
  }

  function renderClarityZone() {
    var isDrag = dragOver === "clarity";
    var loadedCount = clarityFiles.filter(function(f){ return f.result; }).length;
    var totalClarityPeople = clarityRecs.length;
    var totalClarityHours = clarityRecs.reduce(function(s,r){ return s+(r.actualHours||0); }, 0);

    return (
      <div style={{border:"1px solid "+(loadedCount>0?IBM.green20:IBM.gray20), marginBottom:16}}>
        <div style={{background:loadedCount>0?IBM.green50:IBM.purple60, color:"#fff", padding:"10px 16px", display:"flex", justifyContent:"space-between", alignItems:"center"}}>
          <div>
            <div style={{fontSize:13, fontWeight:700}}>
              Clarity Actual Hours
              {loadedCount>0 && <span style={{marginLeft:8, background:"rgba(255,255,255,0.2)", padding:"1px 8px", fontSize:11, borderRadius:10}}>{loadedCount} file{loadedCount>1?"s":""}</span>}
            </div>
            <div style={{fontSize:10, opacity:0.85, marginTop:2}}>Sheet contains month name (e.g. CORP_AML_FCU_Feb2026_Actual hrs) — upload one file per month</div>
          </div>
          <button onClick={function(){ clarityRef.current.click(); }}
            style={{background:"rgba(255,255,255,0.15)", border:"1px solid rgba(255,255,255,0.4)", color:"#fff", padding:"5px 12px", cursor:"pointer", fontSize:11, fontWeight:600, whiteSpace:"nowrap"}}>
            + Add File
          </button>
        </div>
        <input ref={clarityRef} type="file" accept=".xlsx,.xls,.xlsm,.xlsb,.xltx,.xltm,.xlt,.csv,.ods,.tsv" multiple style={{display:"none"}}
          onChange={function(e){ Array.from(e.target.files).forEach(function(f){ readClarityFile(f); }); e.target.value=""; }}/>

        {clarityFiles.length > 0 && (
          <div style={{background:"#fafafa", borderBottom:"1px solid "+IBM.gray20}}>
            {clarityFiles.map(function(f) {
              var mi = f.result && f.result.monthInfo;
              return (
                <div key={f.id} style={{display:"flex", alignItems:"center", gap:10, padding:"8px 14px", borderBottom:"1px solid "+IBM.gray20}}>
                  <span style={{fontSize:14}}>{f.loading?"⏳":f.error?"⚠️":"✅"}</span>
                  <div style={{flex:1, minWidth:0}}>
                    <div style={{fontSize:12, fontWeight:600, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap"}}>{f.file.name}</div>
                    {f.result && (
                      <div style={{fontSize:11, color:IBM.gray60, marginTop:1}}>
                        Sheet: {f.result.sheetName} &nbsp;&#8226;&nbsp; {f.result.rowCount} rows &nbsp;&#8226;&nbsp; {f.result.records.length} people
                        {mi && <span style={{marginLeft:6, background:IBM.purple10, color:IBM.purple60, padding:"1px 6px", border:"1px solid #d4bbff", fontWeight:700, fontSize:10}}>{mi.label}</span>}
                        <span style={{marginLeft:6, color:IBM.purple60, fontWeight:600}}>
                          {f.result.records.reduce(function(s,r){return s+(r.actualHours||0);},0)}h actual
                        </span>
                      </div>
                    )}
                    {f.error && <div style={{fontSize:11, color:IBM.red60, marginTop:1}}>{f.error}</div>}
                    {f.loading && <div style={{fontSize:11, color:IBM.gray50, marginTop:1}}>Parsing...</div>}
                  </div>
                  <button onClick={function(){ removeClarityFile(f.id); }}
                    style={{background:"none", border:"1px solid "+IBM.gray30, color:IBM.gray60, padding:"2px 8px", cursor:"pointer", fontSize:11, flexShrink:0}}>
                    Remove
                  </button>
                </div>
              );
            })}
          </div>
        )}

        <div
          onDragOver={function(e){e.preventDefault();setDragOver("clarity");}}
          onDragLeave={function(){setDragOver(null);}}
          onDrop={function(e){e.preventDefault();setDragOver(null);Array.from(e.dataTransfer.files).forEach(function(f){readClarityFile(f);});}}
          onClick={function(){clarityRef.current.click();}}
          style={{padding:clarityFiles.length>0?"12px 20px":"20px", background:isDrag?IBM.purple10:"#fff", cursor:"pointer", textAlign:"center", borderTop:clarityFiles.length>0?"1px dashed "+IBM.gray30:"none"}}>
          <div style={{fontSize:isDrag?28:20, marginBottom:4}}>📋</div>
          <div style={{fontSize:12, color:isDrag?IBM.purple60:IBM.gray60}}>
            {clarityFiles.length>0 ? "Drop another Clarity file or click to add more" : "Drag and drop Clarity file(s) here or click to browse"}
          </div>
          <div style={{fontSize:10, color:IBM.gray40, marginTop:3}}>.xlsx, .xls, .xlsm, .xlsb, .csv, .ods &nbsp;&#8226;&nbsp; Multiple files allowed (one per month)</div>
        </div>

        {loadedCount > 0 && (
          <div style={{padding:"10px 14px", background:IBM.purple10, borderTop:"1px solid #d4bbff", display:"flex", gap:16, flexWrap:"wrap", alignItems:"center"}}>
            <span style={{fontSize:11, color:IBM.purple60, fontWeight:700}}>Combined:</span>
            <span style={{fontSize:11, background:"#fff", color:IBM.purple60, padding:"2px 8px", border:"1px solid #d4bbff", fontWeight:600}}>{totalClarityPeople} people</span>
            <span style={{fontSize:11, background:"#fff", color:IBM.purple60, padding:"2px 8px", border:"1px solid #d4bbff", fontWeight:600}}>{totalClarityHours}h actual</span>
            {availableMonths.map(function(m){
              return <span key={m} style={{fontSize:11, background:"#fff", color:IBM.gray70, padding:"2px 8px", border:"1px solid "+IBM.gray20}}>{m}</span>;
            })}
          </div>
        )}
      </div>
    );
  }


  return (
    <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(22,22,22,.82)",zIndex:300,display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div className="import-modal" style={{background:"#fff",width:"min(900px,97vw)",maxHeight:"92vh",overflowY:"auto",fontFamily:FF_SANS,border:"1px solid #e0e0e0",display:"flex",flexDirection:"column"}}>
        {/* Header */}
        <div style={{background:IBM.gray100,color:"#fff",padding:"16px 24px",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
          <div>
            <div style={{fontSize:16,fontWeight:600}}>Import Data</div>
            <div style={{fontSize:11,color:IBM.gray30,marginTop:3}}>Upload multiple IBM files (one per WBS ID) and one Clarity file — people are matched by name automatically</div>
          </div>
          <button onClick={onClose} style={{background:"none",border:"1px solid "+IBM.gray70,color:IBM.gray30,fontSize:13,cursor:"pointer",padding:"5px 14px"}}>&#x2715; Close</button>
        </div>

        <div style={{padding:"20px 24px", flex:1, overflowY:"auto"}}>
          {renderIBMZone()}
          {renderClarityZone()}

          {/* Match preview */}
          {merged && (
            <div>
              <div style={{fontSize:12,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.07em",color:IBM.gray70,marginBottom:10,borderBottom:"2px solid "+IBM.blue60,paddingBottom:5}}>
                Match Preview
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:1,background:IBM.gray20,marginBottom:14}}>
                <div style={{background:"#fff",padding:"12px 16px",borderTop:"3px solid "+IBM.green50}}>
                  <div style={{fontSize:24,fontWeight:300,color:IBM.green50}}>{merged.matched.length}</div>
                  <div style={{fontSize:11,color:IBM.gray70,marginTop:3,textTransform:"uppercase"}}>Matched</div>
                  <div style={{fontSize:10,color:IBM.gray50,marginTop:2}}>In both IBM and Clarity</div>
                </div>
                <div style={{background:"#fff",padding:"12px 16px",borderTop:"3px solid "+IBM.blue60}}>
                  <div style={{fontSize:24,fontWeight:300,color:IBM.blue60}}>{merged.ibmOnly.length}</div>
                  <div style={{fontSize:11,color:IBM.gray70,marginTop:3,textTransform:"uppercase"}}>IBM Only</div>
                  <div style={{fontSize:10,color:IBM.gray50,marginTop:2}}>Not found in Clarity</div>
                </div>
                <div style={{background:"#fff",padding:"12px 16px",borderTop:"3px solid "+IBM.purple60}}>
                  <div style={{fontSize:24,fontWeight:300,color:IBM.purple60}}>{merged.clarityOnly.length}</div>
                  <div style={{fontSize:11,color:IBM.gray70,marginTop:3,textTransform:"uppercase"}}>Clarity Only</div>
                  <div style={{fontSize:10,color:IBM.gray50,marginTop:2}}>Not found in IBM</div>
                </div>
              </div>

              {merged.ibmOnly.length > 0 && (
                <div style={{marginBottom:10,padding:"9px 14px",background:IBM.blue10,border:"1px solid "+IBM.blue20,fontSize:12}}>
                  <b style={{color:IBM.blue60}}>IBM Only ({merged.ibmOnly.length}):</b>
                  <span style={{color:IBM.gray70,marginLeft:6}}>{merged.ibmOnly.slice(0,8).map(function(r){return r.name;}).join(", ")}{merged.ibmOnly.length>8?" ...":""}</span>
                </div>
              )}
              {merged.clarityOnly.length > 0 && (
                <div style={{marginBottom:10,padding:"9px 14px",background:IBM.purple10,border:"1px solid #d4bbff",fontSize:12}}>
                  <b style={{color:IBM.purple60}}>Clarity Only ({merged.clarityOnly.length}):</b>
                  <span style={{color:IBM.gray70,marginLeft:6}}>{merged.clarityOnly.slice(0,8).map(function(r){return r.name;}).join(", ")}{merged.clarityOnly.length>8?" ...":""}</span>
                </div>
              )}

              {/* ── Manual Match UI ─────────────────────────────────── */}
              {(merged.ibmOnly.length > 0 || merged.clarityOnly.length > 0) && (
                <div style={{marginBottom:14}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
                    <div style={{fontSize:12,fontWeight:700,color:IBM.orange40}}>
                      &#9888; {merged.ibmOnly.length + merged.clarityOnly.length} unmatched records
                      {Object.keys(manualMatches).length > 0 && <span style={{color:IBM.green50,marginLeft:8}}>&#10003; {Object.keys(manualMatches).length} manually linked</span>}
                    </div>
                    <button onClick={function(){ setShowManualUI(function(v){ return !v; }); }}
                      style={{padding:"5px 14px",background:showManualUI?IBM.orange40:"#fff",color:showManualUI?"#fff":IBM.orange40,border:"1px solid "+IBM.orange40,cursor:"pointer",fontSize:12,fontWeight:600}}>
                      {showManualUI ? "Hide Manual Match" : "Fix Name Mismatches"}
                    </button>
                  </div>

                  {showManualUI && (
                    <div style={{border:"1px solid "+IBM.orange40,background:"#fff"}}>
                      <div style={{background:IBM.orange40,color:"#fff",padding:"10px 16px",fontSize:12,fontWeight:600}}>
                        Manual Name Matching — select IBM and Clarity names that belong to the same person
                      </div>
                      <div style={{padding:"14px 16px"}}>
                        {/* Quick-link: show fuzzy suggestions first */}
                        {merged && Object.keys(merged.suggestions||{}).length > 0 && (
                          <div style={{marginBottom:14}}>
                            <div style={{fontSize:11,fontWeight:700,color:IBM.gray70,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:8}}>
                              Suggested matches (click to accept)
                            </div>
                            {Object.keys(merged.suggestions||{}).map(function(ibmKey) {
                              var suggestions = merged.suggestions[ibmKey];
                              var ibmRec = merged.ibmOnly.find(function(r){ return r.normalizedName===ibmKey; });
                              if (!ibmRec) return null;
                              var alreadyLinked = manualMatches[ibmKey];
                              return (
                                <div key={ibmKey} style={{marginBottom:10,padding:"10px 12px",background:alreadyLinked?IBM.green10:IBM.gray10,border:"1px solid "+(alreadyLinked?IBM.green20:IBM.gray20)}}>
                                  <div style={{display:"flex",alignItems:"flex-start",gap:12,flexWrap:"wrap"}}>
                                    <div style={{minWidth:160,flex:1}}>
                                      <div style={{fontSize:10,color:IBM.blue60,fontWeight:700,textTransform:"uppercase",marginBottom:3}}>IBM</div>
                                      <div style={{fontSize:13,fontWeight:600,color:IBM.gray100}}>{ibmRec.name}</div>
                                      <div style={{fontSize:10,color:IBM.gray50}}>{ibmRec.scheduledHours}h scheduled</div>
                                    </div>
                                    <div style={{display:"flex",flexDirection:"column",gap:4,flex:2}}>
                                      <div style={{fontSize:10,color:IBM.purple60,fontWeight:700,textTransform:"uppercase",marginBottom:3}}>Clarity suggestions</div>
                                      {suggestions.map(function(sug) {
                                        var isSelected = manualMatches[ibmKey] === sug.clarityKey;
                                        var pct = Math.round(sug.score*100);
                                        var barColor = pct >= 80 ? IBM.green50 : pct >= 60 ? IBM.orange40 : IBM.gray50;
                                        return (
                                          <div key={sug.clarityKey}
                                            onClick={function(){
                                              setManualMatches(function(prev){
                                                var next = Object.assign({}, prev);
                                                if (isSelected) { delete next[ibmKey]; }
                                                else { next[ibmKey] = sug.clarityKey; }
                                                return next;
                                              });
                                            }}
                                            style={{display:"flex",alignItems:"center",gap:8,padding:"6px 10px",background:isSelected?IBM.green10:"#fff",border:"1px solid "+(isSelected?IBM.green50:IBM.gray20),cursor:"pointer"}}>
                                            <div style={{flex:1}}>
                                              <div style={{fontSize:12,fontWeight:isSelected?700:400,color:isSelected?IBM.green50:IBM.gray100}}>{sug.clarityName}</div>
                                              <div style={{display:"flex",alignItems:"center",gap:6,marginTop:3}}>
                                                <div style={{flex:1,height:4,background:IBM.gray20,overflow:"hidden"}}>
                                                  <div style={{height:"100%",width:pct+"%",background:barColor}}/>
                                                </div>
                                                <span style={{fontSize:9,color:barColor,fontWeight:700,whiteSpace:"nowrap"}}>{pct}% match</span>
                                              </div>
                                            </div>
                                            {isSelected && <span style={{fontSize:16,color:IBM.green50}}>&#10003;</span>}
                                          </div>
                                        );
                                      })}
                                    </div>
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        )}

                        {/* Manual dropdowns for all unmatched IBM */}
                        <div>
                          <div style={{fontSize:11,fontWeight:700,color:IBM.gray70,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:8}}>
                            All unmatched IBM records — assign Clarity name manually
                          </div>
                          <div style={{overflowX:"auto"}}>
                            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                              <thead>
                                <tr style={{background:IBM.gray90,color:"#fff"}}>
                                  {["IBM Name","Sched Hrs","Assign Clarity Name",""].map(function(h){
                                    return <th key={h} style={{padding:"7px 10px",textAlign:"left",fontWeight:400,fontSize:10,textTransform:"uppercase",borderRight:"1px solid "+IBM.gray80}}>{h}</th>;
                                  })}
                                </tr>
                              </thead>
                              <tbody>
                                {merged.ibmOnly.map(function(r, ri) {
                                  var linked = manualMatches[r.normalizedName];
                                  var linkedRec = linked ? merged.clarityOnly.find(function(c){ return c.normalizedName===linked; }) : null;
                                  return (
                                    <tr key={r.id} style={{background:linked?IBM.green10:ri%2?IBM.gray10:"#fff"}}>
                                      <td style={{padding:"8px 10px",borderBottom:"1px solid "+IBM.gray20,fontWeight:600}}>{r.name}</td>
                                      <td style={{padding:"8px 10px",borderBottom:"1px solid "+IBM.gray20,textAlign:"right"}}>{r.scheduledHours}h</td>
                                      <td style={{padding:"8px 10px",borderBottom:"1px solid "+IBM.gray20}}>
                                        <select
                                          value={linked||""}
                                          onChange={function(e){
                                            var val = e.target.value;
                                            setManualMatches(function(prev){
                                              var next = Object.assign({}, prev);
                                              if (!val) { delete next[r.normalizedName]; }
                                              else { next[r.normalizedName] = val; }
                                              return next;
                                            });
                                          }}
                                          style={{width:"100%",padding:"5px 8px",border:"1px solid "+(linked?IBM.green50:IBM.gray30),fontSize:12,background:"#fff",color:IBM.gray100,outline:"none"}}>
                                          <option value="">-- Select Clarity person --</option>
                                          {merged.clarityOnly.map(function(c){
                                            var score = fuzzyMatchScore(r.name, c.name);
                                            var label = score >= 0.4 ? c.name + " (" + Math.round(score*100) + "% match)" : c.name;
                                            return <option key={c.normalizedName} value={c.normalizedName} style={{background:"#fff",color:IBM.gray100}}>{label}</option>;
                                          })}
                                        </select>
                                      </td>
                                      <td style={{padding:"8px 10px",borderBottom:"1px solid "+IBM.gray20}}>
                                        {linked && <button onClick={function(){ setManualMatches(function(prev){ var next=Object.assign({},prev); delete next[r.normalizedName]; return next; }); }} style={{background:"none",border:"1px solid "+IBM.red60,color:IBM.red60,padding:"2px 7px",cursor:"pointer",fontSize:10}}>&#x2715;</button>}
                                      </td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        {Object.keys(manualMatches).length > 0 && (
                          <div style={{marginTop:10,padding:"8px 12px",background:IBM.green10,border:"1px solid "+IBM.green20,fontSize:12,color:IBM.green50,fontWeight:600}}>
                            &#10003; {Object.keys(manualMatches).length} manual link{Object.keys(manualMatches).length>1?"s":""} added — match preview above will update automatically
                          </div>
                        )}
                      </div>
                    </div>
                  )}
                </div>
              )}

              {merged.matched.length > 0 && (
                <div style={{overflowX:"auto",border:"1px solid "+IBM.gray20,marginBottom:12}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                    <thead>
                      <tr style={{background:IBM.gray100,color:"#fff"}}>
                        {["Name","WBS ID","Talent ID","Country","Scheduled","Actual","Variance","Status","Manager","Period(s)"].map(function(h){
                          return <th key={h} style={{padding:"7px 10px",textAlign:"left",fontWeight:400,fontSize:10,textTransform:"uppercase",borderRight:"1px solid "+IBM.gray80,whiteSpace:"nowrap"}}>{h}</th>;
                        })}
                      </tr>
                    </thead>
                    <tbody>
                      {merged.matched.slice(0,15).map(function(r,i){
                        var variance = r.scheduledHours - r.actualHours;
                        var varColor = variance===0?IBM.green50:variance>0?IBM.red60:IBM.orange40;
                        var varStr = variance===0?"✓":(variance>0?"-"+variance+"h":"+"+Math.abs(variance)+"h");
                        return (
                          <tr key={r.id} style={{background:i%2?IBM.gray10:"#fff"}}>
                            <td style={{padding:"7px 10px",borderBottom:"1px solid "+IBM.gray20,fontWeight:600}}>{r.name}</td>
                            <td style={{padding:"7px 10px",borderBottom:"1px solid "+IBM.gray20,fontSize:11,color:IBM.gray60,maxWidth:120,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={r.wbsId}>{r.wbsId||"—"}</td>
                            <td style={{padding:"7px 10px",borderBottom:"1px solid "+IBM.gray20,fontSize:11}}>{r.talentId||"—"}</td>
                            <td style={{padding:"7px 10px",borderBottom:"1px solid "+IBM.gray20,fontSize:11}}>{r.country||"—"}</td>
                            <td style={{padding:"7px 10px",borderBottom:"1px solid "+IBM.gray20,textAlign:"right",fontWeight:600}}>{r.scheduledHours}h</td>
                            <td style={{padding:"7px 10px",borderBottom:"1px solid "+IBM.gray20,textAlign:"right",fontWeight:600,color:r.actualHours>0?IBM.green50:IBM.red60}}>{r.actualHours}h</td>
                            <td style={{padding:"7px 10px",borderBottom:"1px solid "+IBM.gray20,textAlign:"right",fontWeight:700,color:varColor}}>{varStr}</td>
                            <td style={{padding:"7px 10px",borderBottom:"1px solid "+IBM.gray20,fontSize:11}}>{r.timesheetStatus||"—"}</td>
                            <td style={{padding:"7px 10px",borderBottom:"1px solid "+IBM.gray20,fontSize:11}}>{r.resourceManager||"—"}</td>
                            <td style={{padding:"7px 10px",borderBottom:"1px solid "+IBM.gray20,fontSize:11}}>{r.periods.join(", ")||"—"}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                  {merged.matched.length > 15 && (
                    <div style={{padding:"7px 12px",background:IBM.gray10,fontSize:11,color:IBM.gray60,borderTop:"1px solid "+IBM.gray20}}>
                      Showing 15 of {merged.matched.length} matched records. All will be imported.
                    </div>
                  )}
                </div>
              )}
            </div>
          )}
        </div>

        {/* Footer */}
        <div style={{padding:"14px 24px",borderTop:"1px solid "+IBM.gray20,display:"flex",gap:12,alignItems:"center",flexShrink:0,background:IBM.gray10}}>
          {merged ? (
            <button onClick={handleImportClick} style={{padding:"10px 28px",background:IBM.blue60,color:"#fff",border:"none",cursor:"pointer",fontSize:14,fontWeight:600}}>
              &#10003; Import {totalImport} Records
            </button>
          ) : (
            <div style={{fontSize:13,color:IBM.gray50}}>Upload at least one file above to continue.</div>
          )}
          <button onClick={onClose} style={{padding:"10px 18px",background:"none",border:"1px solid "+IBM.gray30,color:IBM.gray70,cursor:"pointer",fontSize:13}}>Cancel</button>
          {merged && (
            <span style={{fontSize:12,color:IBM.gray50}}>
              Matched: {merged.matched.length} &#8226; IBM only: {merged.ibmOnly.length} &#8226; Clarity only: {merged.clarityOnly.length}
            </span>
          )}
          {loadedIBMFiles.length > 1 && (
            <span style={{fontSize:11,color:IBM.blue60,fontWeight:600}}>
              {loadedIBMFiles.length} IBM files merged &#8226; {wbsList.length} WBS IDs
            </span>
          )}
        </div>
      </div>
    </div>
  );
}
// ─── BULK BAR ─────────────────────────────────────────────────────────────────
function BulkBar({selected,total,onSelectAll,onClearAll,onBulkNotif,onSendAll,onBulkTeams,onBulkEmail,bulkLoading,bulkProgress,notifications,bulkTeamsSent,bulkEmailSent}){
  const hasGen=selected.some(id=>notifications[id]);const hasSel=selected.length>0;
  const bBtn=(active,bg)=>({padding:"6px 12px",background:active?bg:IBM.gray70,color:active?"#fff":IBM.gray50,border:"none",cursor:active?"pointer":"not-allowed",fontSize:12,fontWeight:600});
  return(
    <div style={{background:IBM.gray90,color:"#fff",padding:"10px 20px",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",borderBottom:`2px solid ${IBM.blue60}`}}>
      <span style={{fontSize:13,fontWeight:600}}>{selected.length} selected <span style={{color:IBM.gray30,fontWeight:400}}>of {total}</span></span>
      <button onClick={onSelectAll} style={{background:"none",border:`1px solid ${IBM.gray60}`,color:IBM.gray30,padding:"4px 10px",cursor:"pointer",fontSize:12}}>All</button>
      <button onClick={onClearAll} style={{background:"none",border:`1px solid ${IBM.gray60}`,color:IBM.gray30,padding:"4px 10px",cursor:"pointer",fontSize:12}}>Clear</button>
      <span style={{width:1,height:18,background:IBM.gray70,flexShrink:0}}/>
      <button onClick={onBulkNotif} disabled={!hasSel||bulkLoading} style={bBtn(hasSel&&!bulkLoading,IBM.blue60)}>{bulkLoading?`Generating… ${bulkProgress}`:`📋 Draft (${selected.length})`}</button>
      {hasGen&&<button onClick={onSendAll} style={bBtn(true,IBM.purple60)}>✉ Send All</button>}
      <span style={{width:1,height:18,background:IBM.gray70,flexShrink:0}}/>
      <button onClick={onBulkTeams} disabled={!hasSel} style={bBtn(hasSel,bulkTeamsSent?"#464775":"#5b5ea6")}>{bulkTeamsSent?`✓ Teams (${selected.length})`:`💬 Teams (${selected.length})`}</button>
      <button onClick={onBulkEmail} disabled={!hasSel} style={bBtn(hasSel,bulkEmailSent?IBM.green50:IBM.orange40)}>{bulkEmailSent?`✓ Email (${selected.length})`:`✉ Email (${selected.length})`}</button>
    </div>
  );
}

// ─── NAME MATCH PANEL (post-import fix) ──────────────────────────────────────
function NameMatchPanel({users, setUsers, onClose}) {
  var ibmOnly     = users.filter(function(u){ return u.dataSource === "IBM only"; });
  var clarityOnly = users.filter(function(u){ return u.dataSource === "Clarity only"; });

  const[search,    setSearch]    = useState("");
  const[selIBM,    setSelIBM]    = useState(null);
  const[selClarity,setSelClarity]= useState(null);
  const[linked,    setLinked]    = useState([]);
  const[applyDone, setApplyDone] = useState(false);

  // Always guard: ibmRec/clarityRec could be undefined if id no longer in list
  var ibmRec     = selIBM     ? (ibmOnly.find(function(u){ return u.id === selIBM; })     || null) : null;
  var clarityRec = selClarity ? (clarityOnly.find(function(u){ return u.id === selClarity; }) || null) : null;

  // If selection became stale (record removed after apply), clear it
  if (selIBM && !ibmRec)     { setTimeout(function(){ setSelIBM(null); }, 0); }
  if (selClarity && !clarityRec) { setTimeout(function(){ setSelClarity(null); }, 0); }

  // Compute fuzzy candidates for selected IBM record — guard ibmRec
  var candidates = useMemo(function(){
    if (!ibmRec || !ibmRec.name) {
      return clarityOnly.map(function(c){ return { u:c, score:0 }; });
    }
    var scored = clarityOnly.map(function(c){
      return { u:c, score: fuzzyMatchScore(ibmRec.name, c.name) };
    });
    scored.sort(function(a,b){ return b.score - a.score; });
    return scored;
  }, [selIBM, users.length, clarityOnly.length]);

  // Filter clarity list
  var filteredClarity = useMemo(function(){
    var q = search.toLowerCase();
    if (!q) return candidates;
    return candidates.filter(function(x){
      return x.u && x.u.name && x.u.name.toLowerCase().indexOf(q) !== -1;
    });
  }, [candidates, search]);

  // Filter IBM list
  var ibmFiltered = useMemo(function(){
    var q = search.toLowerCase();
    var linkedIds = {};
    linked.forEach(function(l){ linkedIds[l.ibmId] = true; });
    return ibmOnly.filter(function(u){
      if (!u || !u.id) return false;
      if (linkedIds[u.id]) return false;
      if (!q) return true;
      return u.name && u.name.toLowerCase().indexOf(q) !== -1;
    });
  }, [ibmOnly.length, users.length, linked.length, search]);

  function addLink() {
    if (!selIBM || !selClarity) return;
    var newLinks = linked.filter(function(l){ return l.ibmId !== selIBM && l.clarityId !== selClarity; });
    newLinks.push({ ibmId: selIBM, clarityId: selClarity });
    setLinked(newLinks);
    setSelIBM(null);
    setSelClarity(null);
    setSearch("");
  }

  function removeLink(ibmId) {
    setLinked(linked.filter(function(l){ return l.ibmId !== ibmId; }));
  }

  function applyLinks() {
    if (!linked.length) return;
    setUsers(function(prev){
      var next = prev.slice();
      linked.forEach(function(link){
        var ibmIdx     = next.findIndex(function(u){ return u && u.id === link.ibmId; });
        var clarityIdx = next.findIndex(function(u){ return u && u.id === link.clarityId; });
        if (ibmIdx === -1 || clarityIdx === -1) return;
        var ibm     = next[ibmIdx];
        var clarity = next[clarityIdx];
        if (!ibm || !clarity) return;
        // Merge IBM scheduled data + Clarity actual data
        next[ibmIdx] = Object.assign({}, ibm, {
          entered:         clarity.entered     || 0,
          actualHours:     clarity.entered     || 0,
          resourceManager: clarity.resourceManager || ibm.resourceManager,
          timesheetStatus: clarity.timesheetStatus || ibm.timesheetStatus,
          approvedBy:      clarity.approvedBy  || ibm.approvedBy,
          resourceActive:  clarity.resourceActive || ibm.resourceActive,
          periods:         clarity.periods     || ibm.periods || [],
          clarityPeriods:  clarity.clarityPeriods || clarity.periods || [],
          monthlyHours:    clarity.monthlyHours || ibm.monthlyHours || {},
          clarityName:     clarity.name,
          dataSource:      "Both",
          lastEntry:       (clarity.periods && clarity.periods.length)
                            ? clarity.periods[clarity.periods.length-1]
                            : ibm.lastEntry,
        });
        // Remove the now-merged clarity record
        var removeIdx = next.findIndex(function(u){ return u && u.id === link.clarityId; });
        if (removeIdx !== -1) next.splice(removeIdx, 1);
      });
      return next;
    });
    setLinked([]);
    setApplyDone(true);
    setTimeout(function(){ onClose(); }, 900);
  }

  var pendingCount = linked.length;

  return (
    <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(22,22,22,.6)",zIndex:500,display:"flex",justifyContent:"flex-end"}} onClick={onClose}>
      <div className="panel-slide" style={{background:"#fff",width:"min(900px,100vw)",height:"100vh",overflowY:"hidden",boxShadow:"-8px 0 32px rgba(0,0,0,.2)",fontFamily:FF_SANS,display:"flex",flexDirection:"column"}} onClick={function(e){e.stopPropagation();}}>

        {/* Header */}
        <div style={{background:IBM.orange40,color:"#fff",padding:"14px 24px",flexShrink:0}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div>
              <div style={{fontSize:16,fontWeight:700}}>Fix Name Mismatches</div>
              <div style={{fontSize:12,opacity:0.9,marginTop:3}}>
                {ibmOnly.length} IBM-only &nbsp;&#8226;&nbsp; {clarityOnly.length} Clarity-only
                {pendingCount > 0 && <span style={{marginLeft:8,background:"rgba(255,255,255,0.25)",padding:"1px 8px",fontWeight:700}}>{pendingCount} linked</span>}
              </div>
            </div>
            <button onClick={onClose} style={{background:"none",border:"1px solid rgba(255,255,255,0.5)",color:"#fff",padding:"5px 14px",cursor:"pointer",fontSize:13}}>&#x2715; Close</button>
          </div>
        </div>

        {/* Pending links banner */}
        {pendingCount > 0 && (
          <div style={{background:IBM.green10,borderBottom:"1px solid "+IBM.green20,padding:"10px 24px",flexShrink:0,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
            <span style={{fontSize:12,fontWeight:700,color:IBM.green50}}>{pendingCount} match{pendingCount>1?"es":""} queued</span>
            <div style={{display:"flex",gap:6,flexWrap:"wrap",flex:1}}>
              {linked.map(function(l){
                var ib = users.find(function(u){ return u && u.id===l.ibmId; });
                var cl = users.find(function(u){ return u && u.id===l.clarityId; });
                if (!ib || !cl) return null;
                return (
                  <span key={l.ibmId} style={{fontSize:11,background:"#fff",border:"1px solid "+IBM.green20,padding:"2px 8px",color:IBM.gray80,display:"inline-flex",alignItems:"center",gap:6}}>
                    <b style={{color:IBM.blue60}}>{ib.name}</b>
                    <span style={{color:IBM.gray40}}>&#8596;</span>
                    <b style={{color:IBM.purple60}}>{cl.name}</b>
                    <button onClick={function(){ removeLink(l.ibmId); }} style={{background:"none",border:"none",color:IBM.red60,cursor:"pointer",fontSize:13,lineHeight:1,padding:"0 1px"}}>&#x2715;</button>
                  </span>
                );
              })}
            </div>
            <button onClick={applyLinks}
              style={{padding:"8px 20px",background:applyDone?IBM.green50:IBM.green50,color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:700,whiteSpace:"nowrap",flexShrink:0}}>
              {applyDone ? "✓ Applied!" : "Apply " + pendingCount + " Match" + (pendingCount>1?"es":"")}
            </button>
          </div>
        )}

        {/* Instructions */}
        <div style={{background:"#fffbf5",borderBottom:"1px solid "+IBM.yellow20,padding:"8px 24px",flexShrink:0,fontSize:12,color:"#6e4a00"}}>
          <b>How to use:</b> Click an IBM name on the left &rarr; Clarity names sort by match score on the right &rarr; click the Clarity name &rarr; confirm the link &rarr; Apply when done.
        </div>

        {/* Split pane */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",flex:1,overflow:"hidden",minHeight:0}}>

          {/* LEFT: IBM Only */}
          <div style={{borderRight:"1px solid "+IBM.gray20,display:"flex",flexDirection:"column",overflow:"hidden"}}>
            <div style={{background:IBM.blue10,borderBottom:"1px solid "+IBM.blue20,padding:"10px 16px",flexShrink:0}}>
              <div style={{fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.07em",color:IBM.blue60,marginBottom:6}}>
                IBM Only — {ibmOnly.length} record{ibmOnly.length!==1?"s":""}
              </div>
              <input value={search} onChange={function(e){ setSearch(e.target.value); setSelClarity(null); }}
                placeholder="Search by name…"
                style={{width:"100%",padding:"6px 10px",border:"1px solid "+IBM.blue20,fontSize:12,outline:"none",background:"#fff"}}/>
            </div>
            <div style={{flex:1,overflowY:"auto"}}>
              {ibmOnly.length === 0 && (
                <div style={{padding:"32px",textAlign:"center",color:IBM.green50,fontSize:13,fontWeight:600}}>&#10003; All IBM records matched</div>
              )}
              {ibmFiltered.map(function(u){
                if (!u || !u.id) return null;
                var isSelected = selIBM === u.id;
                return (
                  <div key={u.id}
                    onClick={function(){ setSelIBM(isSelected ? null : u.id); setSelClarity(null); }}
                    style={{padding:"11px 16px",borderBottom:"1px solid "+IBM.gray20,cursor:"pointer",
                      background:isSelected ? IBM.blue10 : "#fff",
                      borderLeft: isSelected ? "3px solid "+IBM.blue60 : "3px solid transparent",
                      transition:"background 0.1s"}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontSize:13,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:IBM.gray100}}>{u.name}</div>
                        <div style={{fontSize:11,color:IBM.gray60,marginTop:2,display:"flex",gap:8,flexWrap:"wrap"}}>
                          {u.scheduled > 0 && <span>{u.scheduled}h sched</span>}
                          {u.wbsId && <span>{u.wbsId}</span>}
                          {u.country && <span>{u.country}</span>}
                        </div>
                      </div>
                      <span style={{fontSize:9,background:IBM.blue10,color:IBM.blue60,padding:"2px 6px",border:"1px solid "+IBM.blue20,fontWeight:700,flexShrink:0,marginLeft:6}}>IBM</span>
                    </div>
                    {isSelected && (
                      <div style={{fontSize:11,color:IBM.blue60,marginTop:5,fontWeight:600}}>
                        &#8594; Now select the matching Clarity name on the right
                      </div>
                    )}
                  </div>
                );
              })}
              {/* Linked rows shown at bottom */}
              {linked.map(function(l){
                var ib = users.find(function(u){ return u && u.id===l.ibmId; });
                var cl = users.find(function(u){ return u && u.id===l.clarityId; });
                if (!ib || !cl) return null;
                return (
                  <div key={l.ibmId} style={{padding:"10px 16px",borderBottom:"1px solid "+IBM.gray20,background:IBM.green10,borderLeft:"3px solid "+IBM.green50}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                      <div>
                        <div style={{fontSize:12,fontWeight:700,color:IBM.green50}}>&#10003; {ib.name}</div>
                        <div style={{fontSize:11,color:IBM.gray60,marginTop:1}}>Linked to: <b style={{color:IBM.purple60}}>{cl.name}</b></div>
                      </div>
                      <button onClick={function(e){ e.stopPropagation(); removeLink(l.ibmId); }}
                        style={{background:"none",border:"1px solid "+IBM.red60,color:IBM.red60,padding:"2px 8px",cursor:"pointer",fontSize:10}}>Undo</button>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* RIGHT: Clarity Only */}
          <div style={{display:"flex",flexDirection:"column",overflow:"hidden"}}>
            <div style={{background:IBM.purple10,borderBottom:"1px solid #d4bbff",padding:"10px 16px",flexShrink:0}}>
              <div style={{fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.07em",color:IBM.purple60,marginBottom:2}}>
                Clarity Only — {clarityOnly.length} record{clarityOnly.length!==1?"s":""}
              </div>
              {!selIBM
                ? <div style={{fontSize:11,color:IBM.gray50}}>&#8592; Select an IBM record first</div>
                : <div style={{fontSize:11,color:IBM.purple60}}>Showing best matches for <b>{ibmRec ? ibmRec.name : ""}</b></div>
              }
            </div>
            <div style={{flex:1,overflowY:"auto"}}>
              {clarityOnly.length === 0 && (
                <div style={{padding:"32px",textAlign:"center",color:IBM.green50,fontSize:13,fontWeight:600}}>&#10003; All Clarity records matched</div>
              )}
              {filteredClarity.map(function(item){
                if (!item || !item.u || !item.u.id) return null;
                var c        = item.u;
                var score    = item.score || 0;
                var pct      = Math.round(score * 100);
                var isSelected   = selClarity === c.id;
                var alreadyLinked = linked.some(function(l){ return l.clarityId === c.id; });
                var barColor = pct >= 85 ? IBM.green50 : pct >= 55 ? IBM.orange40 : IBM.gray40;
                return (
                  <div key={c.id}
                    onClick={function(){ if (!alreadyLinked && selIBM) setSelClarity(isSelected ? null : c.id); }}
                    style={{padding:"11px 16px",borderBottom:"1px solid "+IBM.gray20,
                      cursor: alreadyLinked || !selIBM ? "default" : "pointer",
                      background: alreadyLinked ? IBM.gray10 : isSelected ? IBM.purple10 : "#fff",
                      borderLeft: isSelected ? "3px solid "+IBM.purple60 : alreadyLinked ? "3px solid "+IBM.gray30 : "3px solid transparent",
                      opacity: alreadyLinked ? 0.5 : 1}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontSize:13,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:IBM.gray100}}>{c.name}</div>
                        <div style={{fontSize:11,color:IBM.gray60,marginTop:2,display:"flex",gap:8,flexWrap:"wrap"}}>
                          {c.entered > 0 && <span>{c.entered}h actual</span>}
                          {c.resourceManager && c.resourceManager !== "—" && <span>{c.resourceManager}</span>}
                          {c.timesheetStatus && <span style={{color:c.timesheetStatus==="Posted"?IBM.blue60:"#8e6a00"}}>{c.timesheetStatus}</span>}
                        </div>
                        {selIBM && (
                          <div style={{display:"flex",alignItems:"center",gap:6,marginTop:5}}>
                            <div style={{flex:1,height:5,background:IBM.gray20,overflow:"hidden",borderRadius:2}}>
                              <div style={{height:"100%",width:pct+"%",background:barColor,transition:"width 0.2s"}}/>
                            </div>
                            <span style={{fontSize:10,color:barColor,fontWeight:700,width:44,textAlign:"right",whiteSpace:"nowrap"}}>
                              {pct > 0 ? pct+"%" : "—"}
                            </span>
                          </div>
                        )}
                      </div>
                      <span style={{fontSize:9,background:IBM.purple10,color:IBM.purple60,padding:"2px 6px",border:"1px solid #d4bbff",fontWeight:700,flexShrink:0}}>Clarity</span>
                    </div>
                    {alreadyLinked && <div style={{fontSize:10,color:IBM.gray50,marginTop:3}}>Already linked to another IBM record</div>}
                  </div>
                );
              })}
            </div>

            {/* Confirm link button */}
            {selIBM && selClarity && ibmRec && clarityRec && (
              <div style={{padding:"12px 16px",borderTop:"1px solid "+IBM.gray20,background:"#fff",flexShrink:0}}>
                <div style={{fontSize:12,color:IBM.gray70,marginBottom:8}}>
                  Link <b style={{color:IBM.blue60}}>{ibmRec.name}</b> &#8596; <b style={{color:IBM.purple60}}>{clarityRec.name}</b>
                </div>
                <button onClick={addLink}
                  style={{width:"100%",padding:"10px",background:IBM.orange40,color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:700}}>
                  &#10003; Confirm Link
                </button>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}


// ─── USER MANAGEMENT TAB ─────────────────────────────────────────────────────
function UserManagementTab({session, showToast}) {
  const[users,    setUsers]    = useState([]);
  const[loading,  setLoading]  = useState(true);
  const[showForm, setShowForm] = useState(false);
  const[editUser, setEditUser] = useState(null);  // null = new, object = editing
  const[search,   setSearch]   = useState("");
  const[confirmDel, setConfirmDel] = useState(null);
  const[saving,   setSaving]   = useState(false);
  const[changePwFor, setChangePwFor] = useState(null); // user id for pw change
  const[newPw,    setNewPw]    = useState("");
  const[newPw2,   setNewPw2]   = useState("");
  const[form, setForm] = useState({
    username:"", full_name:"", email:"", dept:"", emp_id:"", role:"user", password:"", password2:""
  });
  const[formErr, setFormErr] = useState("");

  // Load users on mount
  React.useEffect(function(){
    setLoading(true);
    getAllUsers().then(function(rows){
      setUsers(rows||[]); setLoading(false);
    }).catch(function(ex){
      showToast("Failed to load users: "+(ex.message||ex), "error"); setLoading(false);
    });
  }, []);

  function reload() {
    getAllUsers().then(function(rows){ setUsers(rows||[]); }).catch(function(){});
  }

  function openNew() {
    setForm({username:"",full_name:"",email:"",dept:"",emp_id:"",role:"user",password:"",password2:""});
    setFormErr(""); setEditUser(null); setShowForm(true);
  }

  function openEdit(u) {
    setForm({username:u.username,full_name:u.full_name||"",email:u.email||"",dept:u.dept||"",emp_id:u.emp_id||"",role:u.role,password:"",password2:""});
    setFormErr(""); setEditUser(u); setShowForm(true);
  }

  async function handleSave() {
    setFormErr("");
    var uname = form.username.trim().toLowerCase();
    if (!uname) { setFormErr("Username is required."); return; }
    if (!/^[a-z0-9._-]+$/.test(uname)) { setFormErr("Username: only letters, numbers, . _ - allowed."); return; }
    if (!editUser && !form.password) { setFormErr("Password is required for new users."); return; }
    if (form.password && form.password !== form.password2) { setFormErr("Passwords do not match."); return; }
    if (form.password && form.password.length < 8) { setFormErr("Password must be at least 8 characters."); return; }
    setSaving(true);
    try {
      var payload = {
        username: uname,
        full_name: form.full_name.trim(),
        email: form.email.trim(),
        dept: form.dept.trim(),
        emp_id: form.emp_id.trim(),
        role: form.role,
        is_active: true,
        created_by: session.username,
      };
      if (form.password) {
        payload.password_hash = await hashPassword(uname, form.password);
      }
      if (editUser) {
        await updateUser(editUser.id, payload);
        showToast("✓ User updated: " + uname);
      } else {
        await createUser(payload);
        showToast("✓ User created: " + uname);
      }
      setShowForm(false); reload();
    } catch(ex) {
      setFormErr(ex.message||"Save failed. Please try again.");
    }
    setSaving(false);
  }

  async function handleDelete(u) {
    try {
      await deactivateUser(u.id);
      showToast("User deactivated: " + u.username);
      setConfirmDel(null); reload();
    } catch(ex) {
      showToast("Failed: "+(ex.message||ex), "error");
    }
  }

  async function handleChangePw() {
    if (!newPw) return;
    if (newPw !== newPw2) { showToast("Passwords do not match", "error"); return; }
    if (newPw.length < 8) { showToast("Password must be at least 8 characters", "error"); return; }
    var u = users.find(function(x){ return x.id === changePwFor; });
    if (!u) return;
    setSaving(true);
    try {
      var hashed = await hashPassword(u.username, newPw);
      await updateUser(u.id, { password_hash: hashed });
      showToast("✓ Password changed for " + u.username);
      setChangePwFor(null); setNewPw(""); setNewPw2("");
    } catch(ex) {
      showToast("Failed: "+(ex.message||ex), "error");
    }
    setSaving(false);
  }

  var filtered = users.filter(function(u){
    if (!u.is_active) return false;
    var q = search.toLowerCase();
    return !q || (u.username&&u.username.indexOf(q)!==-1)
               || (u.full_name&&u.full_name.toLowerCase().indexOf(q)!==-1)
               || (u.email&&u.email.toLowerCase().indexOf(q)!==-1)
               || (u.dept&&u.dept.toLowerCase().indexOf(q)!==-1);
  });

  var supOK = isSupabaseConfigured();

  // ── FORM ──────────────────────────────────────────────────────────────────
  if (showForm) return (
    <div style={{padding:"24px 28px",maxWidth:580,fontFamily:FF_SANS}}>
      <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:20}}>
        <button onClick={function(){setShowForm(false);}} style={{background:"none",border:"1px solid "+IBM.gray30,color:IBM.gray70,padding:"5px 12px",cursor:"pointer",fontSize:12}}>&#8592; Back</button>
        <h2 style={{fontSize:18,fontWeight:600,color:IBM.gray100,margin:0}}>{editUser?"Edit User":"Add New User"}</h2>
      </div>
      <div style={{background:"#fff",border:"1px solid "+IBM.gray20}}>
        <div style={{background:editUser?IBM.blue60:IBM.green50,color:"#fff",padding:"12px 20px",fontSize:13,fontWeight:600}}>
          {editUser ? "Editing: "+editUser.username : "Create New Account"}
        </div>
        <div style={{padding:"22px 24px",display:"grid",gridTemplateColumns:"1fr 1fr",gap:"14px 20px"}}>
          {/* Username */}
          <div style={{gridColumn:"1 / -1"}}>
            <label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",color:IBM.gray70,display:"block",marginBottom:5,letterSpacing:"0.07em"}}>Username *</label>
            <input value={form.username} onChange={function(e){setForm(function(f){return Object.assign({},f,{username:e.target.value});});}}
              disabled={!!editUser}
              placeholder="e.g. john.smith"
              style={{width:"100%",padding:"9px 12px",border:"1px solid "+IBM.gray30,fontSize:13,outline:"none",background:editUser?"#f4f4f4":"#fff",color:IBM.gray100}}/>
            {!editUser&&<div style={{fontSize:11,color:IBM.gray50,marginTop:3}}>Lowercase letters, numbers, dots, hyphens only</div>}
          </div>
          {/* Full name */}
          <div>
            <label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",color:IBM.gray70,display:"block",marginBottom:5,letterSpacing:"0.07em"}}>Full Name</label>
            <input value={form.full_name} onChange={function(e){setForm(function(f){return Object.assign({},f,{full_name:e.target.value});});}}
              placeholder="First Last" style={{width:"100%",padding:"9px 12px",border:"1px solid "+IBM.gray30,fontSize:13,outline:"none"}}/>
          </div>
          {/* Email */}
          <div>
            <label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",color:IBM.gray70,display:"block",marginBottom:5,letterSpacing:"0.07em"}}>Email</label>
            <input value={form.email} onChange={function(e){setForm(function(f){return Object.assign({},f,{email:e.target.value});});}}
              placeholder="user@company.com" type="email" style={{width:"100%",padding:"9px 12px",border:"1px solid "+IBM.gray30,fontSize:13,outline:"none"}}/>
          </div>
          {/* Dept */}
          <div>
            <label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",color:IBM.gray70,display:"block",marginBottom:5,letterSpacing:"0.07em"}}>Department</label>
            <input value={form.dept} onChange={function(e){setForm(function(f){return Object.assign({},f,{dept:e.target.value});});}}
              placeholder="e.g. Engineering" style={{width:"100%",padding:"9px 12px",border:"1px solid "+IBM.gray30,fontSize:13,outline:"none"}}/>
          </div>
          {/* Emp ID */}
          <div>
            <label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",color:IBM.gray70,display:"block",marginBottom:5,letterSpacing:"0.07em"}}>Employee ID</label>
            <input value={form.emp_id} onChange={function(e){setForm(function(f){return Object.assign({},f,{emp_id:e.target.value});});}}
              placeholder="e.g. E001" style={{width:"100%",padding:"9px 12px",border:"1px solid "+IBM.gray30,fontSize:13,outline:"none"}}/>
          </div>
          {/* Role */}
          <div>
            <label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",color:IBM.gray70,display:"block",marginBottom:5,letterSpacing:"0.07em"}}>Role</label>
            <select value={form.role} onChange={function(e){setForm(function(f){return Object.assign({},f,{role:e.target.value});});}}
              style={{width:"100%",padding:"9px 12px",border:"1px solid "+IBM.gray30,fontSize:13,outline:"none",background:"#fff"}}>
              <option value="user">Employee</option>
              <option value="manager">Manager</option>
            </select>
          </div>
          {/* Password fields */}
          <div>
            <label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",color:IBM.gray70,display:"block",marginBottom:5,letterSpacing:"0.07em"}}>{editUser?"New Password (leave blank to keep)":"Password *"}</label>
            <input type="password" value={form.password} onChange={function(e){setForm(function(f){return Object.assign({},f,{password:e.target.value});});}}
              placeholder={editUser?"Leave blank to keep current":"Min 8 characters"}
              style={{width:"100%",padding:"9px 12px",border:"1px solid "+IBM.gray30,fontSize:13,outline:"none"}}/>
          </div>
          <div>
            <label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",color:IBM.gray70,display:"block",marginBottom:5,letterSpacing:"0.07em"}}>Confirm Password</label>
            <input type="password" value={form.password2} onChange={function(e){setForm(function(f){return Object.assign({},f,{password2:e.target.value});});}}
              placeholder="Re-enter password"
              style={{width:"100%",padding:"9px 12px",border:"1px solid "+IBM.gray30,fontSize:13,outline:"none"}}/>
          </div>
        </div>
        {formErr&&<div style={{margin:"0 24px 16px",padding:"10px 14px",background:"#fff1f1",border:"1px solid #ffb3b8",color:IBM.red60,fontSize:13}}>&#9888; {formErr}</div>}
        <div style={{padding:"14px 24px",borderTop:"1px solid "+IBM.gray20,display:"flex",gap:10}}>
          <button onClick={handleSave} disabled={saving}
            style={{padding:"10px 24px",background:saving?IBM.gray30:IBM.blue60,color:"#fff",border:"none",cursor:saving?"not-allowed":"pointer",fontSize:13,fontWeight:600}}>
            {saving ? "Saving…" : (editUser ? "Save Changes" : "Create User")}
          </button>
          <button onClick={function(){setShowForm(false);}}
            style={{padding:"10px 18px",background:"none",border:"1px solid "+IBM.gray30,color:IBM.gray70,cursor:"pointer",fontSize:13}}>Cancel</button>
        </div>
      </div>
    </div>
  );

  // ── CHANGE PASSWORD MODAL ──────────────────────────────────────────────────
  if (changePwFor) {
    var cpUser = users.find(function(u){ return u.id===changePwFor; });
    return (
      <div style={{padding:"24px 28px",maxWidth:460,fontFamily:FF_SANS}}>
        <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:20}}>
          <button onClick={function(){setChangePwFor(null);setNewPw("");setNewPw2("");}} style={{background:"none",border:"1px solid "+IBM.gray30,color:IBM.gray70,padding:"5px 12px",cursor:"pointer",fontSize:12}}>&#8592; Back</button>
          <h2 style={{fontSize:18,fontWeight:600,margin:0}}>Change Password</h2>
        </div>
        <div style={{background:"#fff",border:"1px solid "+IBM.gray20}}>
          <div style={{background:IBM.orange40,color:"#fff",padding:"12px 20px",fontSize:13,fontWeight:600}}>Changing password for: <b>{cpUser&&cpUser.username}</b></div>
          <div style={{padding:"22px 24px"}}>
            <div style={{marginBottom:16}}>
              <label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",color:IBM.gray70,display:"block",marginBottom:6,letterSpacing:"0.07em"}}>New Password</label>
              <input type="password" value={newPw} onChange={function(e){setNewPw(e.target.value);}} placeholder="Min 8 characters"
                style={{width:"100%",padding:"10px 12px",border:"1px solid "+IBM.gray30,fontSize:13,outline:"none"}}/>
            </div>
            <div style={{marginBottom:20}}>
              <label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",color:IBM.gray70,display:"block",marginBottom:6,letterSpacing:"0.07em"}}>Confirm New Password</label>
              <input type="password" value={newPw2} onChange={function(e){setNewPw2(e.target.value);}} placeholder="Re-enter password"
                style={{width:"100%",padding:"10px 12px",border:"1px solid "+IBM.gray30,fontSize:13,outline:"none"}}/>
            </div>
            <div style={{display:"flex",gap:10}}>
              <button onClick={handleChangePw} disabled={saving||!newPw||newPw!==newPw2}
                style={{padding:"10px 22px",background:saving||!newPw||newPw!==newPw2?IBM.gray30:IBM.orange40,color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:600}}>
                {saving?"Saving…":"Update Password"}
              </button>
              <button onClick={function(){setChangePwFor(null);setNewPw("");setNewPw2("");}}
                style={{padding:"10px 16px",background:"none",border:"1px solid "+IBM.gray30,color:IBM.gray70,cursor:"pointer",fontSize:13}}>Cancel</button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ── MAIN LIST ──────────────────────────────────────────────────────────────
  return (
    <div style={{padding:"20px 28px",fontFamily:FF_SANS}}>
      {/* Header */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16,flexWrap:"wrap",gap:10}}>
        <div>
          <h2 style={{fontSize:18,fontWeight:600,color:IBM.gray100,margin:0}}>User Management</h2>
          <div style={{fontSize:12,color:IBM.gray60,marginTop:3}}>
            {supOK
              ? <span style={{color:IBM.green50}}>&#9679; Connected to Supabase — users shared across all devices</span>
              : <span style={{color:IBM.orange40}}>&#9679; Local mode — users stored in this browser only. <a href="#" onClick={function(e){e.preventDefault();alert("Set SUPABASE_URL and SUPABASE_ANON_KEY in your deployment environment variables.");}} style={{color:IBM.blue60}}>Configure Supabase</a></span>
            }
          </div>
        </div>
        <button onClick={openNew} style={{padding:"9px 20px",background:IBM.blue60,color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:600}}>
          + Add User
        </button>
      </div>

      {/* Search */}
      <div style={{marginBottom:14,display:"flex",alignItems:"center",gap:8,background:IBM.gray10,border:"1px solid "+IBM.gray20,padding:"6px 12px",maxWidth:340}}>
        <span style={{color:IBM.gray50,fontSize:12}}>&#128269;</span>
        <input value={search} onChange={function(e){setSearch(e.target.value);}} placeholder="Search name, username, dept…"
          style={{border:"none",background:"transparent",fontSize:13,outline:"none",flex:1,color:IBM.gray100}}/>
        {search&&<button onClick={function(){setSearch("");}} style={{background:"none",border:"none",color:IBM.gray50,cursor:"pointer",fontSize:13}}>&#x2715;</button>}
      </div>

      {/* Table */}
      {loading ? (
        <div style={{textAlign:"center",padding:"48px",color:IBM.gray50,fontSize:13}}>Loading users…</div>
      ) : (
        <div style={{overflowX:"auto",border:"1px solid "+IBM.gray20}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead>
              <tr style={{background:IBM.gray100,color:"#fff"}}>
                {["","Username","Full Name","Email","Dept","Emp ID","Role","Created","Actions"].map(function(h){
                  return <th key={h} style={{padding:"9px 12px",textAlign:"left",fontWeight:400,fontSize:11,textTransform:"uppercase",letterSpacing:"0.06em",borderRight:"1px solid "+IBM.gray80,whiteSpace:"nowrap"}}>{h}</th>;
                })}
              </tr>
            </thead>
            <tbody>
              {filtered.length===0&&(
                <tr><td colSpan={9} style={{padding:"32px",textAlign:"center",color:IBM.gray50}}>No users found</td></tr>
              )}
              {filtered.map(function(u,i){
                var isManager = u.role==="manager";
                var isSelf    = u.username===session.username;
                var alt       = i%2;
                return (
                  <tr key={u.id} style={{background:alt?IBM.gray10:"#fff"}}>
                    <td style={{padding:"10px 12px",borderBottom:"1px solid "+IBM.gray20}}>
                      <div style={{width:30,height:30,borderRadius:"50%",background:isManager?"#5b5ea6":IBM.blue60,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,color:"#fff"}}>
                        {(u.full_name||u.username).charAt(0).toUpperCase()}
                      </div>
                    </td>
                    <td style={{padding:"10px 12px",borderBottom:"1px solid "+IBM.gray20,fontWeight:600,whiteSpace:"nowrap"}}>
                      {u.username}
                      {isSelf&&<span style={{marginLeft:6,fontSize:9,background:IBM.blue10,color:IBM.blue60,padding:"1px 5px",border:"1px solid "+IBM.blue20,fontWeight:700}}>YOU</span>}
                    </td>
                    <td style={{padding:"10px 12px",borderBottom:"1px solid "+IBM.gray20}}>{u.full_name||"—"}</td>
                    <td style={{padding:"10px 12px",borderBottom:"1px solid "+IBM.gray20,color:IBM.gray60,fontSize:12}}>{u.email||"—"}</td>
                    <td style={{padding:"10px 12px",borderBottom:"1px solid "+IBM.gray20}}>{u.dept||"—"}</td>
                    <td style={{padding:"10px 12px",borderBottom:"1px solid "+IBM.gray20,fontFamily:FF_MONO,fontSize:12}}>{u.emp_id||"—"}</td>
                    <td style={{padding:"10px 12px",borderBottom:"1px solid "+IBM.gray20}}>
                      <span style={{fontSize:10,padding:"2px 8px",fontWeight:700,background:isManager?"#f6f2ff":IBM.blue10,color:isManager?IBM.purple60:IBM.blue60,border:"1px solid "+(isManager?"#d4bbff":IBM.blue20)}}>
                        {isManager?"Manager":"Employee"}
                      </span>
                    </td>
                    <td style={{padding:"10px 12px",borderBottom:"1px solid "+IBM.gray20,color:IBM.gray50,fontSize:11,whiteSpace:"nowrap"}}>
                      {u.created_at ? new Date(u.created_at).toLocaleDateString() : "—"}
                    </td>
                    <td style={{padding:"10px 12px",borderBottom:"1px solid "+IBM.gray20,whiteSpace:"nowrap"}}>
                      <div style={{display:"flex",gap:5}}>
                        <button onClick={function(){openEdit(u);}} title="Edit user"
                          style={{padding:"4px 10px",background:"none",border:"1px solid "+IBM.blue60,color:IBM.blue60,cursor:"pointer",fontSize:11,fontWeight:600}}>Edit</button>
                        <button onClick={function(){setChangePwFor(u.id);}} title="Change password"
                          style={{padding:"4px 10px",background:"none",border:"1px solid "+IBM.orange40,color:IBM.orange40,cursor:"pointer",fontSize:11,fontWeight:600}}>&#128274;</button>
                        {!isSelf&&(
                          <button onClick={function(){setConfirmDel(u);}} title="Deactivate user"
                            style={{padding:"4px 8px",background:"none",border:"1px solid "+IBM.red60,color:IBM.red60,cursor:"pointer",fontSize:12}}>&#x2715;</button>
                        )}
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
      <div style={{marginTop:10,fontSize:11,color:IBM.gray50}}>{filtered.length} of {users.filter(function(u){return u.is_active!==false;}).length} active users</div>

      {/* Delete confirm */}
      {confirmDel&&(
        <div style={{position:"fixed",inset:0,background:"rgba(22,22,22,.7)",zIndex:600,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:"#fff",width:"min(400px,96vw)",border:"1px solid "+IBM.gray20,fontFamily:FF_SANS}} onClick={function(e){e.stopPropagation();}}>
            <div style={{background:IBM.red60,color:"#fff",padding:"14px 20px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <b style={{fontSize:14}}>Deactivate User</b>
              <button onClick={function(){setConfirmDel(null);}} style={{background:"none",border:"none",color:"#fff",fontSize:20,cursor:"pointer"}}>&#x2715;</button>
            </div>
            <div style={{padding:"20px"}}>
              <p style={{fontSize:13,color:IBM.gray80,marginBottom:18}}>
                Deactivate <b>{confirmDel.full_name||confirmDel.username}</b>? They will not be able to log in. You can re-enable them later by editing their account.
              </p>
              <div style={{display:"flex",gap:10}}>
                <button onClick={function(){handleDelete(confirmDel);}}
                  style={{padding:"9px 20px",background:IBM.red60,color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:600}}>Deactivate</button>
                <button onClick={function(){setConfirmDel(null);}}
                  style={{padding:"9px 16px",background:"none",border:"1px solid "+IBM.gray30,color:IBM.gray70,cursor:"pointer",fontSize:13}}>Cancel</button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}


// ─── MANAGER APP ──────────────────────────────────────────────────────────────
function ManagerApp({session,onLogout,users,setUsers,calendarEvents,setCalendarEvents}){
  const now=new Date();
  const[activeTab,setActiveTab]=useState("dashboard");
  const[selMonth,setSelMonth]=useState(MONTH_NAMES[now.getMonth()]);
  const[selYear,setSelYear]=useState(now.getFullYear());
  const[selPeriod,setSelPeriod]=useState("WM");
  const[isImported,setIsImported]=useState(false);
  const[importedMonths,setImportedMonths]=useState([]); // e.g. ["February-2026","March-2026"]
  const[showAllMonths,setShowAllMonths]=useState(false); // false=filter by selMonth+selYear, true=all months
  const[showImport,setShowImport]=useState(false);
  const[filterStatus,setFilterStatus]=useState("all");
  const[sortMode,setSortMode]=useState("severity-desc");
  const[search,setSearch]=useState("");
  const[notifications,setNotifications]=useState({});
  const[selected,setSelected]=useState([]);
  const[bulkLoading,setBulkLoading]=useState(false);
  const[bulkProgress,setBulkProgress]=useState("");
  const[bulkTeamsSent,setBulkTeamsSent]=useState(false);
  const[bulkEmailSent,setBulkEmailSent]=useState(false);
  const[toast,setToast]=useState(null);
  const[detailUserId,setDetailUserId]=useState(null);  // store ID only — panel reads live from users
  const[showNameMatch,setShowNameMatch]=useState(false);
  const[filterRM,setFilterRM]=useState("");        // resource manager filter
  const[filterWBS,setFilterWBS]=useState("");      // WBS/project filter
  const[filterSource,setFilterSource]=useState("all"); // all|Both|IBM only|Clarity only
  const[confirmDelete,setConfirmDelete]=useState(null); // user id to delete
  const[emailLog,setEmailLog]=useState([]);
  const[mgrName,setMgrName]=useState(session.name||"");
  const[mgrEmail,setMgrEmail]=useState(session.email||"");
  const[mgrDept,setMgrDept]=useState(session.dept||"");
  const[mgrPhone,setMgrPhone]=useState("");
  const[mgrSaved,setMgrSaved]=useState(false);

  const periodLabel=(PERIODS.find(p=>p.value===selPeriod)||{label:selPeriod}).label;
  const monthLabel=`${selMonth} ${selYear}`;
  const showToast=(msg,type="success")=>{setToast({msg,type});setTimeout(()=>setToast(null),4000);};
  const handleImport=function(data, mergedInfo) {
    // Map imported records to user shape expected by the app
    var mapped = data.map(function(r) {
      return {
        id: r.id || r.normalizedName,
        name: r.name,
        clarityName: r.clarityName || null,
        email: r.email || "",
        dept: r.country || "Imported",
        resourceManager: r.resourceManager || "—",
        scheduled: r.scheduledHours || 0,
        entered: r.actualHours || 0,
        lastEntry: r.periods && r.periods.length ? r.periods[r.periods.length-1] : null,
        projects: r.workitems ? r.workitems.slice(0,3).map(function(w,i){ return {code:"WI-"+(i+1), name:w}; }) : [],
        weeklyBreakdown: r.weeklyBreakdown || [],
        clarityPeriods: r.clarityPeriods || r.periods || [],
        activityCode: r.activityCode || "",
        monthlyHours: r.monthlyHours || {},
        ibmMonthlyHours: r.ibmMonthlyHours || {},
        sourceMonths: r.sourceMonths || [],
        // Extended IBM/Clarity fields
        talentId: r.talentId || "",
        serialId: r.serialId || "",
        country: r.country || "",
        billingCode: r.billingCode || "",
        wbsId: r.wbsId || "",
        claimMonths: r.claimMonths || [],
        timesheetStatus: r.timesheetStatus || "",
        approvedBy: r.approvedBy || "",
        resourceActive: r.resourceActive || "",
        periods: r.periods || [],
        dayHours: r.dayHours || {},
        dataSource: r.matched ? "Both" : (r.scheduledHours > 0 ? "IBM only" : "Clarity only"),
        history: [],
        monthlyEntries: {},
      };
    });
    setUsers(mapped);
    setIsImported(true);
    setNotifications({});
    setSelected([]);
    // Collect all months from the imported data
    var allMonths = {};
    data.forEach(function(r){
      Object.keys(r.monthlyHours||{}).forEach(function(mk){ allMonths[mk]=true; });
    });
    setImportedMonths(Object.keys(allMonths).sort());
    var matchCount = mergedInfo ? mergedInfo.matched.length : data.length;
    showToast("✓ Imported " + data.length + " records (" + matchCount + " matched)");
  };

  const total=users.length,
    complete=users.filter(u=>getStatus(u)==="green").length,
    mismatch=users.filter(u=>getStatus(u)==="yellow").length,
    missing=users.filter(u=>getStatus(u)==="red").length,
    clarityNoSched=users.filter(u=>getStatus(u)==="purple").length,
    pct=total?Math.round(((complete+clarityNoSched)/total)*100):0;
  var clarityOnlyCount2=users.filter(u=>u.dataSource==="Clarity only").length;
  const pieData=[{name:"Complete",value:complete,color:IBM.green50},{name:"Mismatch",value:mismatch,color:IBM.yellow30},{name:"Missing",value:missing,color:IBM.red60},{name:"No IBM Sched",value:clarityOnlyCount2,color:IBM.purple60}];
  const deptMap={};users.forEach(u=>{if(!deptMap[u.dept])deptMap[u.dept]={dept:u.dept,complete:0,mismatch:0,missing:0,noibm:0,scheduled:0,actual:0};const st=getStatus(u);deptMap[u.dept][st==="green"?"complete":st==="yellow"?"mismatch":st==="purple"?"noibm":"missing"]++;deptMap[u.dept].scheduled+=Number(u.scheduled)||0;var _amkDept=selMonth+"-"+selYear; deptMap[u.dept].actual+=showAllMonths?Number(u.entered)||0:((u.monthlyHours||{})[_amkDept]||Number(u.entered)||0);});
  const barData=Object.values(deptMap);
  const sevDist=useMemo(()=>{const d={0:0,1:0,2:0,3:0,4:0};users.forEach(u=>d[getSeverity(u)]++);return[0,1,2,3,4].map(k=>({label:SEV[k].label,value:d[k],color:k===0?IBM.green50:k===1?"#0e6027":k===2?"#8e6a00":k===3?IBM.orange40:IBM.red60}));},[users]);
  const filtered=useMemo(function(){
    var q=search.toLowerCase();
    var qrm=filterRM.toLowerCase();
    var qwbs=filterWBS.toLowerCase();
    var l=users.filter(function(u){
      var st=getStatus(u);
      // Status filter
      if(filterStatus!=="all"&&st!==filterStatus) return false;
      // Source filter
      if(filterSource!=="all"&&u.dataSource!==filterSource) return false;
      // Text search (name, email, dept, project)
      if(q&&!(
        (u.name&&u.name.toLowerCase().indexOf(q)!==-1)||
        (u.clarityName&&u.clarityName.toLowerCase().indexOf(q)!==-1)||
        (u.email&&u.email.toLowerCase().indexOf(q)!==-1)||
        (u.dept&&u.dept.toLowerCase().indexOf(q)!==-1)||
        (u.wbsId&&u.wbsId.toLowerCase().indexOf(q)!==-1)||
        (u.projects&&u.projects.some(function(p){return (p.code+p.name).toLowerCase().indexOf(q)!==-1;}))
      )) return false;
      // Resource manager filter
      if(qrm&&!(u.resourceManager&&u.resourceManager.toLowerCase().indexOf(qrm)!==-1)) return false;
      // WBS / Project filter
      if(qwbs&&!(
        (u.wbsId&&u.wbsId.toLowerCase().indexOf(qwbs)!==-1)||
        (u.projects&&u.projects.some(function(p){return (p.code+p.name).toLowerCase().indexOf(qwbs)!==-1;}))
      )) return false;
      // Month filter: filter by selMonth+selYear unless showAllMonths=true
      // IBM-only: always shown (scheduled data, no monthly breakdown)
      // Clarity-only: always shown
      // Both: show if monthlyHours has data for selected month
      if(!showAllMonths){
        var activeMonthKey = selMonth + "-" + selYear;
        var clarityMH = u.monthlyHours || {};
        var ibmMH = u.ibmMonthlyHours || {};
        var clarityKeys = Object.keys(clarityMH);
        var ibmKeys = Object.keys(ibmMH);
        var hasAnyClarityMonths = clarityKeys.length > 0;
        var hasAnyIBMMonths = ibmKeys.length > 0;
        // If this record has NO monthly breakdown at all, it was imported with old parser.
        // Show it for any month selection — user must re-import to get month filtering.
        if(!hasAnyClarityMonths && !hasAnyIBMMonths){
          // Still try claimMonths as best-effort
          var claims = u.claimMonths || [];
          var MABBR = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
          var MFULL = ["January","February","March","April","May","June","July","August","September","October","November","December"];
          var abbr = MABBR[MFULL.indexOf(selMonth)] || selMonth.slice(0,3);
          var matchesClaim = claims.length === 0 || claims.some(function(cm){
            return cm === activeMonthKey || cm === selMonth || cm === abbr
                || (cm.indexOf(String(selYear)) !== -1 && (cm.indexOf(selMonth) !== -1 || cm.indexOf(abbr) !== -1));
          });
          if(!matchesClaim) return false;
          // If no claimMonths stored at all, show the record (no data to filter by)
          return true;
        }
        // Record has monthly data — filter strictly by selected month+year
        var hasClarityData = (clarityMH[activeMonthKey] || 0) > 0;
        var hasIBMData     = (ibmMH[activeMonthKey] || 0) > 0;
        if(!hasClarityData && !hasIBMData) return false;
      }
      return true;
    });
    if(sortMode==="severity-desc") l=l.slice().sort(function(a,b){return getSeverity(b)-getSeverity(a);});
    else if(sortMode==="severity-asc") l=l.slice().sort(function(a,b){return getSeverity(a)-getSeverity(b);});
    else if(sortMode==="name") l=l.slice().sort(function(a,b){return a.name.localeCompare(b.name);});
    else if(sortMode==="hours-gap") l=l.slice().sort(function(a,b){return (Number(b.scheduled)-Number(b.entered))-(Number(a.scheduled)-Number(a.entered));});
    return l;
  },[users,filterStatus,filterSource,search,filterRM,filterWBS,sortMode,showAllMonths,selMonth,selYear]);

  const handleGenNotif=u=>{if(getStatus(u)==="green")return;setNotifications(p=>({...p,[u.id]:genNotifForUser(u,monthLabel,periodLabel)}));showToast(`✓ Ready for ${u.name}`);};
  const handleSendEmail=u=>{const n=notifications[u.id]||genNotifForUser(u,monthLabel,periodLabel);setEmailLog(p=>[{id:Date.now(),type:"email",to:u.email,toName:u.name,subject:n.subject,sentAt:new Date().toLocaleString()},...p]);setNotifications(p=>({...p,[u.id]:n}));showToast("✉ Email sent to "+u.email);};
  const handleSendTeams=u=>{const n=notifications[u.id]||genNotifForUser(u,monthLabel,periodLabel);setEmailLog(p=>[{id:Date.now(),type:"teams",to:u.name,toName:u.name,subject:"Teams: "+n.subject,sentAt:new Date().toLocaleString()},...p]);setNotifications(p=>({...p,[u.id]:n}));showToast("💬 Teams sent to "+u.name);};
  const handleFixEntry=(uid,newSch,note)=>{setUsers(prev=>prev.map(u=>u.id!==uid?u:{...u,scheduled:newSch}));showToast(`✓ Updated to ${newSch}h`);};
  const handleBulkNotif=()=>{const targets=users.filter(u=>selected.includes(u.id)&&getStatus(u)!=="green");if(!targets.length){showToast("All selected are complete","error");return;}const results=genBulkNotifs(targets,monthLabel,periodLabel,(d,t,n)=>setBulkProgress(`${d}/${t}`));setNotifications(p=>({...p,...results}));showToast(`✓ ${Object.keys(results).length} notifications ready`);setBulkProgress("");};
  const handleBulkSendAll=()=>{users.filter(u=>selected.includes(u.id)&&notifications[u.id]).forEach(u=>handleSendEmail(u));showToast("Emails sent");};

  const navBtn=a=>({background:"none",border:"none",color:a?"#fff":IBM.gray50,cursor:"pointer",fontSize:13,padding:"0 4px",height:48,borderBottom:a?`2px solid ${IBM.blue60}`:"2px solid transparent",fontFamily:"inherit"});
  const TH={background:IBM.gray100,color:"#fff",padding:"9px 11px",textAlign:"left",fontWeight:400,fontSize:11,textTransform:"uppercase",letterSpacing:"0.06em",borderRight:`1px solid ${IBM.gray80}`,whiteSpace:"nowrap"};
  const TD=alt=>({padding:"10px 11px",borderBottom:`1px solid ${IBM.gray20}`,verticalAlign:"middle",background:alt?IBM.gray10:"#fff"});

  return(
    <div style={{fontFamily:FF_SANS,background:IBM.gray10,minHeight:"100vh",color:IBM.gray100}}>
      <style>{CSS_MANAGER}</style>

      {/* NAV */}
      {(function(){
        var[mobileMenuOpen,setMobileMenuOpen]=React.useState(false);
        return (
          <React.Fragment>
            <nav style={{background:IBM.gray100,padding:"0 16px 0 20px",display:"flex",alignItems:"center",height:48,position:"sticky",top:0,zIndex:200,borderBottom:"1px solid "+IBM.gray80}}>
              <div style={{display:"flex",alignItems:"center",gap:8,flex:1,minWidth:0}}>
                <span style={{fontSize:20,fontWeight:700,color:IBM.blue60,fontFamily:FF_MONO,letterSpacing:"-1px",flexShrink:0}}>IBM</span>
                <span style={{width:1,height:18,background:IBM.gray70,margin:"0 10px",flexShrink:0}}/>
                <span style={{fontSize:13,color:"#f4f4f4",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>Timesheet Manager</span>
                <span style={{background:"#5b5ea6",color:"#fff",padding:"2px 6px",fontSize:10,fontWeight:600,flexShrink:0}}>MGR</span>
                {isImported&&<span style={{background:IBM.green10,border:"1px solid "+IBM.green20,color:"#0e6027",padding:"2px 6px",fontSize:10,fontWeight:600,flexShrink:0}}>● LIVE</span>}
              </div>
              {/* Desktop nav */}
              <div className="mgr-nav-links" style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:10,flexShrink:0}}>
                {[["dashboard","Dashboard"],["records","Records"],["calendar","📅 Calendar"],["users","👥 Users"],["profile","Profile"]].map(function(item){
                  return <button key={item[0]} style={navBtn(activeTab===item[0])} onClick={function(){setActiveTab(item[0]);}}>{item[1]}</button>;
                })}
                <button style={{padding:"5px 12px",background:IBM.blue60,color:"#fff",border:"none",cursor:"pointer",fontSize:12}} onClick={function(){setShowImport(true);}}>↑ Import</button>
                <div style={{width:28,height:28,borderRadius:"50%",background:"#5b5ea6",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,color:"#fff"}}>{(session.name||"M").split(" ").map(function(n){return n[0];}).join("").slice(0,2)}</div>
                <button onClick={onLogout} style={{background:"none",border:"1px solid "+IBM.gray70,color:IBM.gray30,padding:"4px 9px",cursor:"pointer",fontSize:12}}>Sign Out</button>
              </div>
              {/* Mobile hamburger */}
              <div className="mgr-nav-menu" style={{marginLeft:"auto",display:"none",alignItems:"center",gap:8}}>
                <button style={{padding:"5px 10px",background:IBM.blue60,color:"#fff",border:"none",cursor:"pointer",fontSize:11}} onClick={function(){setShowImport(true);}}>↑</button>
                <div style={{width:26,height:26,borderRadius:"50%",background:"#5b5ea6",display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:700,color:"#fff"}}>{(session.name||"M").split(" ").map(function(n){return n[0];}).join("").slice(0,2)}</div>
                <button onClick={function(){setMobileMenuOpen(function(v){return !v;});}}
                  style={{background:"none",border:"1px solid "+IBM.gray60,color:"#fff",padding:"5px 10px",cursor:"pointer",fontSize:16,lineHeight:1}}>
                  {mobileMenuOpen ? "✕" : "☰"}
                </button>
              </div>
            </nav>
            {/* Mobile dropdown menu */}
            {mobileMenuOpen&&(
              <div className="mgr-nav-mobile" style={{display:"none"}}>
                {[["dashboard","📊 Dashboard"],["records","📋 Records"],["calendar","📅 Calendar"],["users","👥 Users"],["profile","👤 Profile"]].map(function(item){
                  return (
                    <button key={item[0]} className="nav-item"
                      style={{padding:"14px 24px",color:activeTab===item[0]?"#0f62fe":"#f4f4f4",fontSize:15,textAlign:"left",background:"none",border:"none",borderBottom:"1px solid #262626",cursor:"pointer",fontFamily:FF_SANS,fontWeight:activeTab===item[0]?600:400,width:"100%"}}
                      onClick={function(){setActiveTab(item[0]);setMobileMenuOpen(false);}}>
                      {item[1]}
                    </button>
                  );
                })}
                <button onClick={function(){onLogout();}} style={{padding:"14px 24px",color:IBM.red60,fontSize:15,textAlign:"left",background:"none",border:"none",cursor:"pointer",fontFamily:FF_SANS,width:"100%"}}>Sign Out</button>
              </div>
            )}
          </React.Fragment>
        );
      })()}

      {/* HEADER BAND */}
      {(activeTab==="dashboard"||activeTab==="records")&&(
        <div className="hdr-band" style={{background:IBM.blue60,color:"#fff",padding:"18px 28px"}}>
          <div style={{display:"flex",alignItems:"flex-end",justifyContent:"space-between",flexWrap:"wrap",gap:12}}>
            <div>
              <h1 style={{fontSize:22,fontWeight:300,margin:0}}>{activeTab==="dashboard"?"Overview":"Employee Records"}</h1>
              <p style={{fontSize:13,color:"#a6c8ff",marginTop:4}}>
                {showAllMonths ? "All Months Combined" : selMonth + " " + selYear}
                {importedMonths.length>0 && (
                  <span style={{marginLeft:6,opacity:0.7,fontSize:12}}>
                    {importedMonths.length} imported month{importedMonths.length>1?"s":""} loaded
                  </span>
                )}
              </p>
            </div>
            <div style={{display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap"}}>
              <div>
                <label style={{fontSize:10,color:"#a6c8ff",textTransform:"uppercase",letterSpacing:"0.07em",display:"block",marginBottom:4}}>Year</label>
                <Sel dark value={selYear} onChange={function(e){
                  setSelYear(Number(e.target.value));
                }} options={YEARS}/>
              </div>
              <div>
                <label style={{fontSize:10,color:"#a6c8ff",textTransform:"uppercase",letterSpacing:"0.07em",display:"block",marginBottom:4}}>Month</label>
                <Sel dark value={selMonth} onChange={function(e){
                  setSelMonth(e.target.value);
                }} options={MONTH_NAMES}/>
              </div>
              <div><label style={{fontSize:10,color:"#a6c8ff",textTransform:"uppercase",letterSpacing:"0.07em",display:"block",marginBottom:4}}>Period</label><Sel dark value={selPeriod} onChange={e=>setSelPeriod(e.target.value)} options={PERIODS.map(p=>({label:p.label,value:p.value}))}/></div>
              <div style={{display:"flex",flexDirection:"column",justifyContent:"flex-end"}}>
                <button onClick={function(){setShowAllMonths(function(v){return !v;});}}
                  style={{padding:"7px 14px",background:showAllMonths?"rgba(255,255,255,0.25)":"rgba(255,255,255,0.1)",border:"1px solid rgba(255,255,255,0.4)",color:"#fff",cursor:"pointer",fontSize:11,fontWeight:600,whiteSpace:"nowrap"}}>
                  {showAllMonths ? "📅 Showing All" : "Show All Months"}
                </button>
              </div>
              {importedMonths.length>0&&(
                <div>
                  <label style={{fontSize:10,color:"#a6c8ff",textTransform:"uppercase",letterSpacing:"0.07em",display:"block",marginBottom:4}}>Imported Data</label>
                  <Sel dark value={selMonth+"-"+selYear} onChange={function(e){
                      if(e.target.value==="All"){ setShowAllMonths(true); }
                      else {
                        var p=e.target.value.split("-");
                        if(MONTH_NAMES.indexOf(p[0])!==-1) setSelMonth(p[0]);
                        if(p[1]) setSelYear(Number(p[1]));
                        setShowAllMonths(false);
                      }
                    }}
                    options={[{label:"All Months",value:"All"}].concat(importedMonths.map(function(m){ var p=m.split("-"); return {label:(p[0]||m)+(p[1]?" "+p[1]:""),value:m}; }))}/>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* DASHBOARD */}
      {activeTab==="dashboard"&&(
        <React.Fragment>
          <div className="dash-variance-bar" style={{margin:"0 28px 1px",background:IBM.gray100,padding:"14px 20px",display:"flex",alignItems:"center",gap:20,flexWrap:"wrap"}}>
            {(function(){
              var totalSched=users.reduce(function(s,u){return s+(Number(u.scheduled)||0);},0);
              var _amk = selMonth + "-" + selYear;
              var totalActual = showAllMonths
                ? users.reduce(function(s,u){return s+(Number(u.entered)||0);},0)
                : users.reduce(function(s,u){
                    var mh = u.monthlyHours || {};
                    return s + (mh[_amk] || Number(u.entered) || 0);
                  },0);
              var totalVar=totalSched-totalActual;
              var varPct=totalSched>0?Math.min(Math.round((totalActual/totalSched)*100),100):0;
              var varColor=totalVar===0?IBM.green50:totalVar>0?IBM.red60:IBM.orange40;
              var varLabel=totalVar===0?"On Track":totalVar>0?"Under-reported":"Over-reported";
              var varDisplay=totalVar===0?"0h":(totalVar>0?"-":"+")+(Math.abs(totalVar))+"h";
              return [
                <div key="v" style={{display:"flex",alignItems:"baseline",gap:8}}>
                  <span style={{fontSize:38,fontWeight:300,color:varColor}}>{varDisplay}</span>
                  <span style={{fontSize:13,color:"#a6c8ff"}}>total variance</span>
                </div>,
                <div key="d1" style={{width:1,height:36,background:IBM.gray80}}/>,
                <div key="s">
                  <div style={{fontSize:10,color:IBM.gray30,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Status</div>
                  <span style={{fontSize:13,fontWeight:700,color:varColor}}>{varLabel}</span>
                </div>,
                <div key="d2" style={{width:1,height:36,background:IBM.gray80}}/>,
                <div key="h">
                  <div style={{fontSize:10,color:IBM.gray30,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Scheduled vs Actual</div>
                  <div style={{fontSize:13,color:"#fff"}}>{totalSched}h scheduled &nbsp;&#8226;&nbsp; {totalActual}h actual</div>
                </div>,
                <div key="bar" style={{flex:1,minWidth:160}}>
                  <div style={{display:"flex",justifyContent:"space-between",fontSize:10,color:IBM.gray50,marginBottom:4}}>
                    <span>Actual vs Scheduled</span><span>{varPct}%</span>
                  </div>
                  <div style={{height:6,background:IBM.gray80,overflow:"hidden"}}>
                    <div style={{height:"100%",width:varPct+"%",background:totalActual>=totalSched?IBM.green50:IBM.orange40}}/>
                  </div>
                </div>
              ];
            })()}
          </div>
          <div className="dash-stat-grid" style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(110px,1fr))",gap:1,background:IBM.gray20,margin:"0 28px",border:"1px solid "+IBM.gray20}}>
            {(function(){
              var matchedCount=users.filter(function(u){return u.dataSource==="Both";}).length;
              var ibmOnlyCount=users.filter(function(u){return u.dataSource==="IBM only";}).length;
              var clarityOnlyCount=users.filter(function(u){return u.dataSource==="Clarity only";}).length;
              // Month-aware actual hours
              var _amk2 = selMonth + "-" + selYear;
              var monthActual = showAllMonths
                ? users.reduce(function(s,u){return s+(Number(u.entered)||0);},0)
                : users.reduce(function(s,u){return s+((u.monthlyHours||{})[_amk2] || Number(u.entered) || 0);},0);
              var monthLabel2 = showAllMonths ? "All Months" : selMonth;
              var cards=[
                {l:"Total People",v:total,c:IBM.blue60},
                {l:"Matched",v:matchedCount,c:IBM.green50},
                {l:"IBM Scheduled",v:users.reduce(function(s,u){return s+(Number(u.scheduled)||0);},0)+"h",c:IBM.blue60},
                {l:monthLabel2+" Actual",v:monthActual+"h",c:monthActual>0?IBM.purple60:IBM.gray50},
                {l:"IBM Only",v:ibmOnlyCount,c:IBM.orange40},
                {l:"Clarity Only",v:clarityOnlyCount,c:IBM.purple60}
              ];
              return cards.map(function(card){
                return <div key={card.l} style={{background:"#fff",padding:"12px 16px",borderTop:"3px solid "+card.c}}><div style={{fontSize:22,fontWeight:300,color:card.c}}>{card.v}</div><div style={{fontSize:10,color:IBM.gray70,marginTop:3,textTransform:"uppercase",letterSpacing:"0.08em"}}>{card.l}</div></div>;
              });
            })()}
          </div>
          <div className="dash-charts" style={{padding:"20px 28px",display:"flex",gap:20,flexWrap:"wrap"}}>
            <div style={{background:"#fff",padding:"18px",flex:"1 1 240px",border:`1px solid ${IBM.gray20}`}}>
              <div style={{fontSize:11,fontWeight:600,marginBottom:10,textTransform:"uppercase",letterSpacing:"0.07em",borderBottom:`2px solid ${IBM.blue60}`,paddingBottom:5}}>Submission Status</div>
              <ResponsiveContainer width="100%" height={200}>
                <PieChart>
                  <Pie data={pieData.filter(function(d){return d.value>0;})}
                    cx="50%" cy="45%" innerRadius={38} outerRadius={65}
                    dataKey="value" paddingAngle={3}
                    labelLine={false}
                    label={function(props){
                      var RADIAN=Math.PI/180;
                      var radius=props.innerRadius+(props.outerRadius-props.innerRadius)*0.5;
                      var x=props.cx+radius*Math.cos(-props.midAngle*RADIAN);
                      var y=props.cy+radius*Math.sin(-props.midAngle*RADIAN);
                      if(props.percent<0.06) return null;
                      return <text x={x} y={y} fill="#fff" textAnchor="middle" dominantBaseline="central" fontSize={11} fontWeight={700}>{props.value}</text>;
                    }}>
                    {pieData.filter(function(d){return d.value>0;}).map(function(d,i){return <Cell key={i} fill={d.color}/>;})}</Pie>
                  <Tooltip formatter={function(value,name){return [value+" people",name];}}/>
                  <Legend iconSize={10} wrapperStyle={{fontSize:11,paddingTop:4}}/>
                </PieChart>
              </ResponsiveContainer>
            </div>
            <div style={{background:"#fff",padding:"18px",flex:"1 1 200px",border:`1px solid ${IBM.gray20}`}}>
              <div style={{fontSize:11,fontWeight:600,marginBottom:10,textTransform:"uppercase",letterSpacing:"0.07em",borderBottom:`2px solid ${IBM.blue60}`,paddingBottom:5}}>Severity</div>
              {sevDist.map(s=><div key={s.label} style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}><span style={{fontSize:11,width:58,color:s.color,fontWeight:600}}>{s.label}</span><div style={{flex:1,height:10,background:IBM.gray10,overflow:"hidden"}}><div style={{height:"100%",width:`${total?Math.round((s.value/total)*100):0}%`,background:s.color}}/></div><span style={{fontSize:12,fontWeight:600,color:s.color,width:22,textAlign:"right"}}>{s.value}</span></div>)}
            </div>
            <div style={{background:"#fff",padding:"18px",flex:"1 1 320px",border:`1px solid ${IBM.gray20}`}}>
              <div style={{fontSize:11,fontWeight:600,marginBottom:10,textTransform:"uppercase",letterSpacing:"0.07em",borderBottom:`2px solid ${IBM.blue60}`,paddingBottom:5}}>By Department</div>
              <ResponsiveContainer width="100%" height={170}><BarChart data={barData} margin={{top:0,right:0,bottom:0,left:-18}}><CartesianGrid strokeDasharray="3 3" stroke={IBM.gray20}/><XAxis dataKey="dept" tick={{fontSize:9}} interval={0} angle={-20} textAnchor="end" height={40}/><YAxis tick={{fontSize:10}}/><Tooltip/><Legend/><Bar dataKey="complete" name="Complete" fill={IBM.green50}/><Bar dataKey="mismatch" name="Mismatch" fill={IBM.yellow30}/><Bar dataKey="missing" name="Missing" fill={IBM.red60}/><Bar dataKey="noibm" name="No IBM Sched" fill={IBM.purple60}/></BarChart></ResponsiveContainer>
            </div>
            <div style={{background:"#fff",padding:"18px",flex:"1 1 200px",border:`1px solid ${IBM.gray20}`}}>
              <div style={{fontSize:11,fontWeight:600,marginBottom:10,textTransform:"uppercase",letterSpacing:"0.07em",borderBottom:`2px solid ${IBM.blue60}`,paddingBottom:5}}>Action Required</div>
              {users.filter(function(u){ var st=getStatus(u); return st!=="green"; }).slice(0,6).map(function(u){
                var st=getStatus(u);
                var issueLabel = st==="purple"?"No IBM Schedule":st==="red"?"Missing Entry":st==="yellow"?"Mismatch":"";
                var issueColor = st==="purple"?IBM.purple60:st==="red"?IBM.red60:IBM.orange40;
                return (
                  <div key={u.id} style={{display:"flex",alignItems:"center",gap:10,marginBottom:8,padding:"7px",background:IBM.gray10,cursor:"pointer",borderLeft:"3px solid "+issueColor}} onClick={function(){setDetailUserId(u.id);}}>
                    <StatusDot status={st}/>
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{u.name}</div>
                      <div style={{fontSize:10,color:issueColor,fontWeight:600}}>{issueLabel}</div>
                    </div>
                    {st!=="purple"&&<SevBadge sev={getSeverity(u)}/>}
                    {st==="purple"&&<span style={{fontSize:9,background:IBM.purple10,color:IBM.purple60,padding:"2px 5px",border:"1px solid #d4bbff",fontWeight:700}}>Clarity Only</span>}
                  </div>
                );
              })}
              {!users.some(function(u){ return getStatus(u)!=="green"; })&&<div style={{color:IBM.green50,fontSize:13,fontWeight:600}}>✓ All complete!</div>}
              {(users.some(function(u){return u.dataSource==="IBM only";})||users.some(function(u){return u.dataSource==="Clarity only";}))&&(
                <button onClick={function(){setShowNameMatch(true);}} style={{marginTop:8,width:"100%",padding:"7px",background:IBM.orange40,color:"#fff",border:"none",cursor:"pointer",fontSize:12,fontWeight:600}}>
                  &#9888; Fix Name Mismatches
                </button>
              )}
            </div>
          </div>
        </React.Fragment>
      )}

      {/* RECORDS */}
      {activeTab==="records"&&(
        <div>
          <BulkBar selected={selected} total={filtered.length} onSelectAll={()=>setSelected(filtered.map(u=>u.id))} onClearAll={()=>setSelected([])} onBulkNotif={handleBulkNotif} onSendAll={handleBulkSendAll} onBulkTeams={()=>{users.filter(u=>selected.includes(u.id)).forEach(u=>handleSendTeams(u));setBulkTeamsSent(true);}} onBulkEmail={()=>{users.filter(u=>selected.includes(u.id)).forEach(u=>handleSendEmail(u));setBulkEmailSent(true);}} bulkLoading={bulkLoading} bulkProgress={bulkProgress} notifications={notifications} bulkTeamsSent={bulkTeamsSent} bulkEmailSent={bulkEmailSent}/>
          {/* Filter bar */}
          <div style={{padding:"12px 28px 0",background:"#fff",borderBottom:"1px solid "+IBM.gray20}}>
            {/* Row 1: Status chips + search + sort + actions */}
            <div className="filter-row1" style={{display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:8,marginBottom:8,alignItems:"center"}}>
              <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
                {[["all","All",IBM.gray80],["green","Complete",IBM.green50],["yellow","Mismatch",IBM.yellow30],["red","Missing",IBM.red60],["purple","No IBM Sched",IBM.purple60]].map(function(item){
                  var v=item[0],l=item[1],c=item[2];
                  return <button key={v} onClick={function(){setFilterStatus(v);}} style={{padding:"4px 10px",border:"1px solid "+(filterStatus===v?c:IBM.gray20),background:filterStatus===v?c:"#fff",color:filterStatus===v?(c===IBM.yellow30?IBM.gray100:"#fff"):IBM.gray70,cursor:"pointer",fontSize:11,fontWeight:600}}>{l}</button>;
                })}
                <span style={{width:1,height:16,background:IBM.gray20,flexShrink:0}}/>
                {/* Source filter */}
                {[["all","All Sources"],["Both","Matched"],["IBM only","IBM Only"],["Clarity only","Clarity Only"]].map(function(item){
                  var v=item[0],l=item[1];
                  var c=v==="Both"?IBM.green50:v==="IBM only"?IBM.blue60:v==="Clarity only"?IBM.purple60:IBM.gray70;
                  return <button key={v} onClick={function(){setFilterSource(v);}} style={{padding:"4px 10px",border:"1px solid "+(filterSource===v?c:IBM.gray20),background:filterSource===v?c:"#fff",color:filterSource===v?"#fff":IBM.gray70,cursor:"pointer",fontSize:11,fontWeight:600}}>{l}</button>;
                })}
                <span style={{width:1,height:16,background:IBM.gray20,flexShrink:0}}/>
                <Sel value={sortMode} onChange={function(e){setSortMode(e.target.value);}} options={[{label:"Severity ↑",value:"severity-desc"},{label:"Severity ↓",value:"severity-asc"},{label:"Name A–Z",value:"name"},{label:"Hours Gap",value:"hours-gap"}]}/>
              </div>
              <div style={{display:"flex",gap:6,alignItems:"center"}}>
                {(function(){
                  var ibmOnlyCount=users.filter(function(u){return u.dataSource==="IBM only";}).length;
                  var clarityOnlyCount=users.filter(function(u){return u.dataSource==="Clarity only";}).length;
                  return (ibmOnlyCount>0||clarityOnlyCount>0)?(
                    <button onClick={function(){setShowNameMatch(true);}} style={{padding:"5px 12px",background:IBM.orange40,border:"none",color:"#fff",cursor:"pointer",fontSize:11,fontWeight:700}}>&#9888; Fix ({ibmOnlyCount+clarityOnlyCount})</button>
                  ):null;
                })()}
                <button onClick={function(){handleExportFull(filtered, showAllMonths ? "All" : selMonth+"-"+selYear);}} style={{padding:"5px 12px",background:IBM.blue60,color:"#fff",border:"none",cursor:"pointer",fontSize:11,fontWeight:600}}>&#8595; Export ({filtered.length})</button>
                <button onClick={function(){downloadConsolidated(users,monthLabel,periodLabel);}} style={{padding:"5px 12px",background:"#fff",border:"1px solid "+IBM.gray30,color:IBM.gray70,cursor:"pointer",fontSize:11}}>Legacy Export</button>
              </div>
            </div>
            {/* Row 2: Column search inputs */}
            <div className="filter-row2" style={{display:"flex",gap:8,flexWrap:"wrap",paddingBottom:10,alignItems:"center"}}>
              <div style={{display:"flex",alignItems:"center",gap:4,background:IBM.gray10,border:"1px solid "+IBM.gray20,padding:"4px 8px"}}>
                <span style={{fontSize:10,color:IBM.gray50,whiteSpace:"nowrap"}}>Name / BMO</span>
                <input value={search} onChange={function(e){setSearch(e.target.value);}} placeholder="Search…"
                  style={{padding:"3px 6px",border:"none",background:"transparent",fontSize:12,outline:"none",width:140}}/>
                {search&&<button onClick={function(){setSearch("");}} style={{background:"none",border:"none",color:IBM.gray50,cursor:"pointer",fontSize:12,padding:"0 2px"}}>&#x2715;</button>}
              </div>
              <div style={{display:"flex",alignItems:"center",gap:4,background:IBM.gray10,border:"1px solid "+IBM.gray20,padding:"4px 8px"}}>
                <span style={{fontSize:10,color:IBM.gray50,whiteSpace:"nowrap"}}>Resource Mgr</span>
                <input value={filterRM} onChange={function(e){setFilterRM(e.target.value);}} placeholder="Filter…"
                  style={{padding:"3px 6px",border:"none",background:"transparent",fontSize:12,outline:"none",width:120}}/>
                {filterRM&&<button onClick={function(){setFilterRM("");}} style={{background:"none",border:"none",color:IBM.gray50,cursor:"pointer",fontSize:12,padding:"0 2px"}}>&#x2715;</button>}
              </div>
              <div style={{display:"flex",alignItems:"center",gap:4,background:IBM.gray10,border:"1px solid "+IBM.gray20,padding:"4px 8px"}}>
                <span style={{fontSize:10,color:IBM.gray50,whiteSpace:"nowrap"}}>WBS / Project</span>
                <input value={filterWBS} onChange={function(e){setFilterWBS(e.target.value);}} placeholder="Filter…"
                  style={{padding:"3px 6px",border:"none",background:"transparent",fontSize:12,outline:"none",width:120}}/>
                {filterWBS&&<button onClick={function(){setFilterWBS("");}} style={{background:"none",border:"none",color:IBM.gray50,cursor:"pointer",fontSize:12,padding:"0 2px"}}>&#x2715;</button>}
              </div>
              {(search||filterRM||filterWBS||filterStatus!=="all"||filterSource!=="all")&&(
                <button onClick={function(){setSearch("");setFilterRM("");setFilterWBS("");setFilterStatus("all");setFilterSource("all");}}
                  style={{padding:"4px 10px",background:"none",border:"1px solid "+IBM.red60,color:IBM.red60,cursor:"pointer",fontSize:11}}>
                  Clear All Filters
                </button>
              )}
              <span style={{marginLeft:"auto",fontSize:11,color:IBM.gray60}}>{filtered.length} of {users.length} records</span>
            </div>
          </div>
          {/* Clarity-only records alert section */}
          {(function(){
            var cOnly = users.filter(function(u){ return u.dataSource==="Clarity only"; });
            if (!cOnly.length) return null;
            return (
              <div style={{margin:"0 28px 8px",background:IBM.purple10,border:"1px solid #d4bbff"}}>
                <div style={{padding:"10px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
                  <div>
                    <span style={{fontSize:12,fontWeight:700,color:IBM.purple60}}>&#9888; {cOnly.length} Clarity record{cOnly.length>1?"s":""} with no IBM scheduled data</span>
                    <span style={{fontSize:11,color:IBM.gray60,marginLeft:8}}>These have actual hours in Clarity but no matching IBM schedule entry</span>
                  </div>
                  <div style={{display:"flex",gap:6}}>
                    <button onClick={function(){setFilterSource("Clarity only");setFilterStatus("all");}} style={{padding:"4px 10px",background:IBM.purple60,color:"#fff",border:"none",cursor:"pointer",fontSize:11,fontWeight:600}}>View All</button>
                    <button onClick={function(){setShowNameMatch(true);}} style={{padding:"4px 10px",background:"none",color:IBM.purple60,border:"1px solid "+IBM.purple60,cursor:"pointer",fontSize:11,fontWeight:600}}>Fix Matches</button>
                  </div>
                </div>
                <div style={{borderTop:"1px solid #d4bbff",padding:"8px 16px",display:"flex",gap:8,flexWrap:"wrap"}}>
                  {cOnly.slice(0,8).map(function(u){
                    return (
                      <span key={u.id} onClick={function(){setDetailUserId(u.id);}} style={{fontSize:11,background:"#fff",border:"1px solid #d4bbff",padding:"2px 8px",color:IBM.purple60,cursor:"pointer",fontWeight:600}}>
                        {u.name} <span style={{color:IBM.gray50,fontWeight:400}}>({u.entered}h)</span>
                      </span>
                    );
                  })}
                  {cOnly.length>8&&<span style={{fontSize:11,color:IBM.gray50}}>+{cOnly.length-8} more</span>}
                </div>
              </div>
            );
          })()}
          <div className="records-table-wrap" style={{overflowX:"auto",margin:"0 28px"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead><tr>{["","","Name","Source","WBS / Talent ID","Dept / Country","Resource Mgr","Workitems","Scheduled","Actual Hrs","Variance","Status","Billing Code","Actions"].map(function(h){ return <th key={h} className={(h==="Billing Code"||h==="Workitems")?"col-hide-mobile":""} style={TH}>{h}</th>; })}</tr></thead>
              <tbody>
                {filtered.map((u,idx)=>{
                  const alt=idx%2,st=getStatus(u),diff=Number(u.scheduled)-Number(u.entered);
                  const mk2=monthKey(selMonth,selYear);
                  const hasNotes=(u.monthlyEntries && u.monthlyEntries[mk2] && u.monthlyEntries[mk2].periodNotes && u.monthlyEntries[mk2].periodNotes.P1)||(u.monthlyEntries && u.monthlyEntries[mk2] && u.monthlyEntries[mk2].periodNotes && u.monthlyEntries[mk2].periodNotes.P2);
                  // Show live entered hours for current month if available
                  var _amkRow = selMonth + "-" + selYear;
                  var liveTotal = showAllMonths
                    ? (Number(u.entered)||0)
                    : ((u.monthlyHours||{})[_amkRow] || Number(u.entered) || 0);
                  return(
                    <tr key={u.id} className="emp-row" onClick={function(){setDetailUserId(u.id);}}>
                      <td style={TD(alt)} onClick={function(e){e.stopPropagation();}}><input type="checkbox" checked={selected.includes(u.id)} onChange={function(){setSelected(function(p){return p.includes(u.id)?p.filter(function(x){return x!==u.id;}):[...p,u.id];});}}/></td>
                      <td style={TD(alt)}><StatusDot status={st}/></td>
                      <td style={Object.assign({},TD(alt),{fontWeight:600})}>
                        <div style={{fontWeight:600,color:IBM.gray100}}>{u.name}</div>
                        {u.clarityName && u.clarityName !== u.name && (
                          <div style={{fontSize:10,color:IBM.purple60,marginTop:2,display:"flex",alignItems:"center",gap:4}}>
                            <span style={{background:IBM.purple10,border:"1px solid #d4bbff",padding:"1px 5px",fontWeight:700,letterSpacing:"0.03em",flexShrink:0}}>BMO</span>
                            <span style={{color:IBM.gray60,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{u.clarityName}</span>
                          </div>
                        )}
                      </td>
                      <td style={TD(alt)}>
                        <span style={{fontSize:10,padding:"2px 7px",fontWeight:700,
                          background:u.dataSource==="Both"?IBM.green10:u.dataSource==="IBM only"?IBM.blue10:IBM.purple10,
                          color:u.dataSource==="Both"?IBM.green50:u.dataSource==="IBM only"?IBM.blue60:IBM.purple60,
                          border:"1px solid "+(u.dataSource==="Both"?IBM.green20:u.dataSource==="IBM only"?IBM.blue20:"#d4bbff")}}>
                          {u.dataSource||"—"}
                        </span>
                      </td>
                      <td style={TD(alt)}>
                        {u.wbsId&&<div style={{fontSize:11,color:IBM.gray70,marginBottom:2}}><b>WBS:</b> {u.wbsId}</div>}
                        {u.talentId&&<div style={{fontSize:11,color:IBM.gray60}}><b>TID:</b> {u.talentId}</div>}
                        {!u.wbsId&&!u.talentId&&<span style={{color:IBM.gray30,fontSize:12}}>—</span>}
                      </td>
                      <td style={TD(alt)}>
                        <div style={{fontSize:12}}>{u.dept}</div>
                        {u.country&&u.country!==u.dept&&<div style={{fontSize:11,color:IBM.gray60}}>{u.country}</div>}
                      </td>
                      <td style={Object.assign({},TD(alt),{fontSize:12,color:IBM.gray60})}>{u.resourceManager}</td>
                      <td style={TD(alt)}>
                        {u.projects&&u.projects.length>0
                          ?<div style={{fontSize:11,color:IBM.gray70}}>{u.projects.slice(0,2).map(function(p){return p.name;}).join("; ")}{u.projects.length>2?" +"+(u.projects.length-2)+" more":""}</div>
                          :<span style={{color:IBM.gray30,fontSize:12}}>—</span>
                        }
                      </td>
                      <td style={Object.assign({},TD(alt),{textAlign:"right"})}>{u.scheduled>0?u.scheduled+"h":"—"}</td>
                      <td style={Object.assign({},TD(alt),{textAlign:"right"})}>
                        {liveTotal>0
                          ?<span style={{color:IBM.green50,fontWeight:700}}>{liveTotal}h</span>
                          :u.entered===0?<span style={{color:IBM.red60,fontWeight:700}}>—</span>:u.entered+"h"
                        }
                      </td>
                      <td style={Object.assign({},TD(alt),{textAlign:"right",color:diff===0?IBM.green50:diff>20?IBM.red60:IBM.orange40,fontWeight:diff>0?700:400})}>{diff===0?"✓":"-"+diff+"h"}</td>
                      <td style={TD(alt)}>
                        {u.timesheetStatus
                          ?<span style={{fontSize:10,padding:"2px 6px",fontWeight:600,
                              background:u.timesheetStatus==="Approved"?IBM.green10:u.timesheetStatus==="Not in Clarity"?IBM.gray10:IBM.yellow10,
                              color:u.timesheetStatus==="Approved"?IBM.green50:u.timesheetStatus==="Not in Clarity"?IBM.gray50:"#8e6a00"}}>{u.timesheetStatus}</span>
                          :<span style={{fontSize:11,color:IBM.gray30}}>—</span>
                        }
                      </td>
                      <td className="col-hide-mobile" style={TD(alt)}><span style={{fontSize:11,color:IBM.gray60}}>{u.billingCode||"—"}</span></td>
                      <td style={Object.assign({},TD(alt),{whiteSpace:"nowrap"})} onClick={function(e){e.stopPropagation();}}>
                        <div style={{display:"flex",gap:4,alignItems:"center"}}>
                          {st!=="green"&&st!=="purple"&&<React.Fragment>
                            <button onClick={function(){handleGenNotif(u);}} style={{padding:"3px 7px",background:notifications[u.id]?IBM.green10:"#fff",color:notifications[u.id]?"#0e6027":IBM.blue60,border:"1px solid "+(notifications[u.id]?IBM.green50:IBM.blue60),cursor:"pointer",fontSize:10,fontWeight:600}}>{notifications[u.id]?"✓":"Draft"}</button>
                            {notifications[u.id]&&<button onClick={function(){handleSendEmail(u);}} style={{padding:"3px 7px",background:IBM.orange40,color:"#fff",border:"none",cursor:"pointer",fontSize:10}}>&#x2709;</button>}
                          </React.Fragment>}
                          <button onClick={function(){setConfirmDelete(u.id);}} title="Delete record"
                            style={{padding:"3px 6px",background:"none",border:"1px solid "+IBM.red60,color:IBM.red60,cursor:"pointer",fontSize:13,lineHeight:1}}>&#x1F5D1;</button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          {filtered.length===0&&<div style={{textAlign:"center",padding:"48px",color:IBM.gray50,fontSize:13}}>No records match your filter.</div>}
        </div>
      )}

      {activeTab==="users"&&<UserManagementTab session={session} showToast={showToast}/>}
      {activeTab==="calendar"&&<CalendarEventsTab calendarEvents={calendarEvents} setCalendarEvents={setCalendarEvents} showToast={showToast}/>}

      {activeTab==="profile"&&(
        <div style={{padding:"28px",maxWidth:800}}>
          {/* Change own password */}
          {(function(){
            var[ownPw,setOwnPw]=React.useState("");
            var[ownPw2,setOwnPw2]=React.useState("");
            var[ownPwErr,setOwnPwErr]=React.useState("");
            var[ownPwOk,setOwnPwOk]=React.useState(false);
            var[ownPwSaving,setOwnPwSaving]=React.useState(false);
            async function handleOwnPw(){
              setOwnPwErr("");
              if(!ownPw){setOwnPwErr("Enter a new password.");return;}
              if(ownPw.length<8){setOwnPwErr("Password must be at least 8 characters.");return;}
              if(ownPw!==ownPw2){setOwnPwErr("Passwords do not match.");return;}
              setOwnPwSaving(true);
              try {
                var users=await getAllUsers();
                var me=users.find(function(u){return u.username===session.username;});
                if(!me){setOwnPwErr("User not found.");setOwnPwSaving(false);return;}
                var hashed=await hashPassword(session.username,ownPw);
                await updateUser(me.id,{password_hash:hashed});
                setOwnPwOk(true);setOwnPw("");setOwnPw2("");
                setTimeout(function(){setOwnPwOk(false);},3000);
              } catch(ex){ setOwnPwErr(ex.message||"Failed"); }
              setOwnPwSaving(false);
            }
            return (
              <div style={{background:"#fff",border:"1px solid "+IBM.gray20,marginBottom:20}}>
                <div style={{background:IBM.orange40,color:"#fff",padding:"12px 20px",fontSize:13,fontWeight:600}}>&#128274; Change Your Password</div>
                <div style={{padding:"18px 20px",display:"flex",gap:14,flexWrap:"wrap",alignItems:"flex-end"}}>
                  <div>
                    <label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",color:IBM.gray70,display:"block",marginBottom:5,letterSpacing:"0.07em"}}>New Password</label>
                    <input type="password" value={ownPw} onChange={function(e){setOwnPw(e.target.value);setOwnPwErr("");}} placeholder="Min 8 characters"
                      style={{padding:"8px 12px",border:"1px solid "+IBM.gray30,fontSize:13,outline:"none",width:200}}/>
                  </div>
                  <div>
                    <label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",color:IBM.gray70,display:"block",marginBottom:5,letterSpacing:"0.07em"}}>Confirm</label>
                    <input type="password" value={ownPw2} onChange={function(e){setOwnPw2(e.target.value);setOwnPwErr("");}} placeholder="Re-enter"
                      style={{padding:"8px 12px",border:"1px solid "+IBM.gray30,fontSize:13,outline:"none",width:200}}/>
                  </div>
                  <button onClick={handleOwnPw} disabled={ownPwSaving}
                    style={{padding:"9px 18px",background:ownPwSaving?IBM.gray30:IBM.orange40,color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:600,flexShrink:0}}>
                    {ownPwSaving?"Saving…":"Update Password"}
                  </button>
                </div>
                {ownPwErr&&<div style={{margin:"0 20px 14px",padding:"8px 12px",background:"#fff1f1",border:"1px solid #ffb3b8",color:IBM.red60,fontSize:12}}>&#9888; {ownPwErr}</div>}
                {ownPwOk&&<div style={{margin:"0 20px 14px",padding:"8px 12px",background:IBM.green10,border:"1px solid "+IBM.green20,color:IBM.green50,fontSize:12,fontWeight:600}}>&#10003; Password updated successfully!</div>}
              </div>
            );
          })()}
          <div style={{background:"#fff",border:`1px solid ${IBM.gray20}`,marginBottom:20}}>
            <div style={{background:IBM.gray100,color:"#fff",padding:"14px 20px",display:"flex",alignItems:"center",gap:12}}>
              <div style={{width:36,height:36,borderRadius:"50%",background:"#5b5ea6",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:700,color:"#fff"}}>{mgrName?mgrName.split(" ").map(n=>n[0]).join("").slice(0,2):"M"}</div>
              <div><div style={{fontSize:14,fontWeight:600}}>{mgrName||"Manager Profile"}</div><div style={{fontSize:12,color:IBM.gray30,marginTop:1}}>{mgrEmail}</div></div>
              {mgrSaved&&<span style={{marginLeft:"auto",background:IBM.green10,color:"#0e6027",border:`1px solid ${IBM.green20}`,padding:"3px 10px",fontSize:11,fontWeight:700}}>✓ SAVED</span>}
            </div>
            <div style={{padding:"22px",display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
              {[{label:"Full Name",val:mgrName,set:setMgrName,ph:"e.g. John Smith",type:"text"},{label:"FROM Email",val:mgrEmail,set:setMgrEmail,ph:"e.g. mgr@company.com",type:"email"},{label:"Department",val:mgrDept,set:setMgrDept,ph:"e.g. Engineering",type:"text"},{label:"Phone / Ext",val:mgrPhone,set:setMgrPhone,ph:"e.g. +1 555 123 4567",type:"text"}].map(({label,val,set,ph,type})=>(
                <div key={label}><label style={{fontSize:11,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.07em",color:IBM.gray70,display:"block",marginBottom:6}}>{label}</label><input type={type} value={val} onChange={e=>set(e.target.value)} placeholder={ph} style={{width:"100%",padding:"9px 12px",border:`1px solid ${IBM.gray30}`,fontSize:13,outline:"none",fontFamily:"inherit"}}/></div>
              ))}
            </div>
            <div style={{padding:"0 22px 20px"}}>
              <button onClick={()=>{setMgrSaved(true);showToast("✓ Profile saved");}} style={{padding:"10px 28px",background:IBM.blue60,color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:600}}>Save Profile</button>
            </div>
          </div>
          {/* Email log */}
          <div style={{background:"#fff",border:`1px solid ${IBM.gray20}`}}>
            <div style={{background:IBM.gray90,color:"#fff",padding:"12px 18px",display:"flex",justifyContent:"space-between"}}><span style={{fontSize:13,fontWeight:600}}>📬 Sent Log</span><span style={{fontSize:12,color:IBM.gray30}}>{emailLog.length} sent</span></div>
            {emailLog.length===0?<div style={{padding:"28px",textAlign:"center",color:IBM.gray50,fontSize:13}}>No notifications sent yet.</div>:(
              <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}><thead><tr style={{background:IBM.gray100,color:"#fff"}}>{["Type","To","Subject","Sent At"].map(h=><th key={h} style={{padding:"8px 14px",textAlign:"left",fontWeight:400,fontSize:11,textTransform:"uppercase",borderRight:`1px solid ${IBM.gray80}`}}>{h}</th>)}</tr></thead><tbody>{emailLog.map((e,i)=><tr key={e.id} style={{background:i%2?IBM.gray10:"#fff"}}><td style={{padding:"9px 14px",borderBottom:`1px solid ${IBM.gray20}`}}><span style={{background:e.type==="email"?IBM.orange40:"#5b5ea6",color:"#fff",padding:"2px 7px",fontSize:11,fontWeight:600}}>{e.type==="email"?"✉ EMAIL":"💬 TEAMS"}</span></td><td style={{padding:"9px 14px",borderBottom:`1px solid ${IBM.gray20}`,fontWeight:600}}>{e.toName}</td><td style={{padding:"9px 14px",borderBottom:`1px solid ${IBM.gray20}`,color:IBM.gray70,maxWidth:260,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{e.subject}</td><td style={{padding:"9px 14px",borderBottom:`1px solid ${IBM.gray20}`,color:IBM.gray60,fontSize:12,whiteSpace:"nowrap"}}>{e.sentAt}</td></tr>)}</tbody></table></div>
            )}
          </div>
        </div>
      )}

      {showImport&&<ImportModal onImport={handleImport} onClose={()=>setShowImport(false)}/>}

      {/* Employee detail panel — receives userId so it always reads live data */}
      {detailUserId&&(
        <EmployeeDetailPanel
          userId={detailUserId}
          users={users}
          monthLabel={monthLabel}
          periodLabel={periodLabel}
          onFixEntry={handleFixEntry}
          onSendEmail={handleSendEmail}
          onSendTeams={handleSendTeams}
          onClose={()=>setDetailUserId(null)}
          calendarEvents={calendarEvents}
        />
      )}



      {/* Delete confirmation */}
      {confirmDelete&&(
        <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(22,22,22,.7)",zIndex:600,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:"#fff",width:"min(420px,96vw)",border:"1px solid "+IBM.gray20,fontFamily:FF_SANS}} onClick={function(e){e.stopPropagation();}}>
            <div style={{background:IBM.red60,color:"#fff",padding:"14px 20px",display:"flex",justifyContent:"space-between"}}>
              <b style={{fontSize:14}}>Delete Record</b>
              <button onClick={function(){setConfirmDelete(null);}} style={{background:"none",border:"none",color:"#fff",fontSize:20,cursor:"pointer"}}>&#x2715;</button>
            </div>
            <div style={{padding:"20px"}}>
              {(function(){
                var u=users.find(function(x){return x.id===confirmDelete;});
                return u?(
                  <div>
                    <p style={{fontSize:13,color:IBM.gray80,marginBottom:16}}>Are you sure you want to delete <b>{u.name}</b>{u.clarityName&&u.clarityName!==u.name?" (BMO: "+u.clarityName+")":""}? This cannot be undone.</p>
                    <div style={{display:"flex",gap:10}}>
                      <button onClick={function(){setUsers(function(prev){return prev.filter(function(x){return x.id!==confirmDelete;});});setConfirmDelete(null);showToast("Record deleted");}} style={{padding:"9px 22px",background:IBM.red60,color:"#fff",border:"none",cursor:"pointer",fontSize:13,fontWeight:600}}>Delete</button>
                      <button onClick={function(){setConfirmDelete(null);}} style={{padding:"9px 18px",background:"none",border:"1px solid "+IBM.gray30,color:IBM.gray70,cursor:"pointer",fontSize:13}}>Cancel</button>
                    </div>
                  </div>
                ):<div style={{color:IBM.gray50,fontSize:13}}>Record not found.</div>;
              })()}
            </div>
          </div>
        </div>
      )}

      {showNameMatch&&(
        <NameMatchPanel
          users={users}
          setUsers={setUsers}
          onClose={function(){setShowNameMatch(false);}}
        />
      )}
      {toast&&<div style={{position:"fixed",bottom:22,right:22,zIndex:9999,background:toast.type==="error"?IBM.red60:IBM.green50,color:"#fff",padding:"12px 22px",fontSize:13,boxShadow:"0 4px 16px rgba(0,0,0,.2)",maxWidth:360}}>{toast.msg}</div>}

      <div style={{background:IBM.gray100,color:IBM.gray60,padding:"12px 28px",fontSize:12,display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
        <span>© {selYear} IBM Corporation · Timesheet Management System</span>
        <span>{monthLabel} · {periodLabel}</span>
      </div>
    </div>
  );
}

// ─── USER APP ─────────────────────────────────────────────────────────────────
function UserApp({session,onLogout,users,setUsers,calendarEvents}){
  const[toast,setToast]=useState(null);
  const showToast=(msg,type="success")=>{setToast({msg,type});setTimeout(()=>setToast(null),4000);};
  const empUser=users.find(u=>u.id===session.empId);
  return(
    <div style={{fontFamily:FF_SANS,background:IBM.gray10,minHeight:"100vh"}}>
      <style>{CSS_USER}</style>
      <nav style={{background:IBM.gray100,padding:"0 28px",display:"flex",alignItems:"center",height:48,position:"sticky",top:0,zIndex:200,borderBottom:`1px solid ${IBM.gray80}`}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <span style={{fontSize:21,fontWeight:700,color:IBM.blue60,fontFamily:FF_MONO,letterSpacing:"-1px"}}>IBM</span>
          <span style={{width:1,height:20,background:IBM.gray70,margin:"0 14px"}}/>
          <span style={{fontSize:14,color:"#f4f4f4"}}>Timesheet Manager</span>
          <span style={{background:IBM.blue60,color:"#fff",padding:"2px 8px",fontSize:11,fontWeight:600}}>EMPLOYEE</span>
        </div>
        <div style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:12}}>
          <span style={{fontSize:13,color:IBM.gray30}}>{(empUser && empUser.name)||session.username}</span>
          <div style={{width:28,height:28,borderRadius:"50%",background:IBM.blue60,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,color:"#fff"}}>{((empUser && empUser.name)||"?").split(" ").map(n=>n[0]).join("")}</div>
          <button onClick={onLogout} style={{background:"none",border:`1px solid ${IBM.gray70}`,color:IBM.gray30,padding:"4px 10px",cursor:"pointer",fontSize:12}}>Sign Out</button>
        </div>
      </nav>
      {empUser
        ?<UserTimesheetView session={session} users={users} setUsers={setUsers} calendarEvents={calendarEvents} showToast={showToast}/>
        :<div style={{padding:40,textAlign:"center",color:IBM.gray60}}>Employee record not found for ID: {session.empId}</div>
      }
      {toast&&<div style={{position:"fixed",bottom:22,right:22,zIndex:9999,background:toast.type==="error"?IBM.red60:IBM.green50,color:"#fff",padding:"12px 22px",fontSize:13,boxShadow:"0 4px 16px rgba(0,0,0,.2)",maxWidth:360}}>{toast.msg}</div>}
    </div>
  );
}

// ─── ROOT APP ─────────────────────────────────────────────────────────────────
export default function App(){
  const[session,setSession]=useState(null);
  const[users,setUsers]=useState(MOCK_USERS);
  const[calendarEvents,setCalendarEvents]=useState([
    {id:1,date:`${new Date().getFullYear()}-${String(new Date().getMonth()+1).padStart(2,"0")}-05`,type:"offshore",label:"Offshore Holiday – Mumbai"},
    {id:2,date:`${new Date().getFullYear()}-${String(new Date().getMonth()+1).padStart(2,"0")}-15`,type:"deadline",label:"Q2 Timesheet Deadline"},
  ]);
  if(!session)return <LoginScreen onLogin={setSession}/>;
  if(session.role==="manager")return <ManagerApp session={session} onLogout={()=>{if(session.msalAccount){msalLogout().catch(function(){});}setSession(null);}} users={users} setUsers={setUsers} calendarEvents={calendarEvents} setCalendarEvents={setCalendarEvents}/>;
  return <UserApp session={session} onLogout={()=>{if(session.msalAccount){msalLogout().catch(function(){});}setSession(null);}} users={users} setUsers={setUsers} calendarEvents={calendarEvents}/>;
}
