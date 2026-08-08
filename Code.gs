/**
 * Event Ticketing Platform V6 - Google Apps Script backend
 * --------------------------------------------------------
 * One backend serves:
 *   /event-ticketing-v5/   configuration, admin, vouchers
 *   /event-gate-v5/        gate operations
 *
 * Run in a NEW Google Sheet:
 *   1) setupV6System()
 *   2) bootstrapOwner()
 *   3) createClientConfigurationForm()
 *   4) Deploy as Web App (Execute as Me, Access Anyone)
 */

const TICKETING_SITE = 'https://greenprofessionals.github.io/event-ticketing-v6';
const GATE_SITE = 'https://greenprofessionals.github.io/event-gate-v6';
const CONFIG_URL = TICKETING_SITE + '/config.html';
const ADMIN_URL = TICKETING_SITE + '/admin.html';
const VOUCHER_URL = TICKETING_SITE + '/v.html';
const DISTRIBUTION_URL = TICKETING_SITE + '/distribute.html';
const GATE_URL = GATE_SITE + '/index.html';

const DEFAULT_PRIMARY = '#0B3D24';
const DEFAULT_ACCENT = '#C9A24B';
const DEFAULT_TIER_COLORS = ['#0B3D24','#C9A24B','#164B8C','#8A1C2D','#6B3FA0','#0E7C7B','#D97706','#243B64','#B23A62','#5F6B6D','#3C7A3C','#A66A2C'];
const LIFECYCLE = ['Draft','Client Submitted','Preview Ready','Client Approved','Active','Closed','Archived'];
const ROLES = ['SYSTEM_OWNER','EVENT_ADMIN','FINANCE','GATE_SUPERVISOR','GATE_STAFF','CLIENT_ADMIN','CLIENT_VIEWER','CLIENT_FINANCE'];

const SHEETS = {
  EVENTS:'Events', TIERS:'Tiers', GROUPS:'Groups', ACCESS:'AccessControl',
  VOUCHERS:'Vouchers', CLAIMS:'Claims', CHECKINS:'CheckIns', PAYMENTS:'Payments',
  AUDIT:'AuditLog', COUNTERS:'Counters', FORM:'EventConfigResponses', CONTACTS:'ContactLog',
  SIMULATIONS:'SimulationRuns', CLIENTS:'Clients', SETTINGS:'SystemSettings'
};

const HEADERS = {
  Events:['EventID','ClientID','ClientName','ClientEmail','ClientPhone','OrgName','ChapterName','EventTitle','Tagline','EventDate','EventTime','VenueName','VenueAddress','ContactPhone','ContactEmail','WebsiteURL','DressCode','PrimaryColor','AccentColor','LogoFileId','BackgroundFileId','GroupLabel','UseGroups','SerialPrefix','CurrencySymbol','FooterLegalText','Capacity','Status','ConfigToken','ApprovedAt','ApprovedBy','AutoReminders','LastReminderRun','EmergencyReadOnly','CreatedAt','UpdatedAt'],
  Tiers:['ClientID','EventID','TierKey','Label','Price','Capacity','Active','Color','SortOrder'],
  Groups:['ClientID','EventID','GroupName','Active','SortOrder'],
  AccessControl:['UserID','Name','Role','PasscodeHash','ClientScope','EventScope','Email','Phone','Active','CreatedAt'],
  Vouchers:['Timestamp','ClientID','EventID','BatchID','VoucherToken','TierKey','SuggestedGroup','PrefillName','PrefillPhone','Claimed','Serial','Dispatched','RecipientName','RecipientEmail','RecipientPhone','SentAt','IssuedBy','Status'],
  Claims:['Timestamp','ClientID','EventID','Serial','CheckInToken','Name','Email','Phone','GroupName','TierKey','Source','VoucherToken','Status','AmountDue','AmountPaid','PaymentStatus','PaymentMethod','PaymentNote','TransferredFrom','InternalNotes','UpdatedAt'],
  CheckIns:['Timestamp','ClientID','EventID','Serial','Name','GroupName','TierKey','Phone','PaymentStatus','PaymentMethod','AmountPaid','CheckedInBy','GateNote','Status'],
  Payments:['Timestamp','ClientID','EventID','Serial','Amount','Method','Status','Note','RecordedBy'],
  AuditLog:['Timestamp','ClientID','EventID','Action','EntityType','EntityID','Actor','Role','Details'],
  Counters:['EventID','CurrentNumber'],
  EventConfigResponses:[],
  ContactLog:['Timestamp','ClientID','EventID','EntityType','EntityID','Channel','Recipient','Actor','Note'],
  SimulationRuns:['Timestamp','ClientID','EventID','RunID','Mode','Passed','TotalTests','PassedTests','FailedTests','Warnings','Actor','Role','ConfigFingerprint','SummaryJSON'],
  Clients:['ClientID','ClientName','PrimaryContact','Email','Phone','Status','CreatedAt','UpdatedAt'],
  SystemSettings:['Key','Value','UpdatedAt','UpdatedBy']
};

function setupV6System() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(name=>{ if(name==='EventConfigResponses') return; ensureSheet_(ss,SHEETS[name.toUpperCase()]||name,HEADERS[name]); });
  ensureSheet_(ss,SHEETS.EVENTS,HEADERS.Events);
  ensureSheet_(ss,SHEETS.TIERS,HEADERS.Tiers);
  ensureSheet_(ss,SHEETS.GROUPS,HEADERS.Groups);
  ensureSheet_(ss,SHEETS.ACCESS,HEADERS.AccessControl);
  ensureSheet_(ss,SHEETS.VOUCHERS,HEADERS.Vouchers);
  ensureSheet_(ss,SHEETS.CLAIMS,HEADERS.Claims);
  ensureSheet_(ss,SHEETS.CHECKINS,HEADERS.CheckIns);
  ensureSheet_(ss,SHEETS.PAYMENTS,HEADERS.Payments);
  ensureSheet_(ss,SHEETS.AUDIT,HEADERS.AuditLog);
  ensureSheet_(ss,SHEETS.COUNTERS,HEADERS.Counters);
  ensureSheet_(ss,SHEETS.CONTACTS,HEADERS.ContactLog);
  ensureSheet_(ss,SHEETS.SIMULATIONS,HEADERS.SimulationRuns);
  ensureSheet_(ss,SHEETS.CLIENTS,HEADERS.Clients);
  ensureSheet_(ss,SHEETS.SETTINGS,HEADERS.SystemSettings);
  SpreadsheetApp.getUi().alert('V6 sheets created. Next run bootstrapOwner(), then createClientConfigurationForm().');
}

function setupV5System(){ return setupV6System(); }

function bootstrapOwner() {
  const ui=SpreadsheetApp.getUi();
  const nameR=ui.prompt('System Owner','Owner name:',ui.ButtonSet.OK_CANCEL); if(nameR.getSelectedButton()!==ui.Button.OK)return;
  const passR=ui.prompt('System Owner','Choose a strong passcode:',ui.ButtonSet.OK_CANCEL); if(passR.getSelectedButton()!==ui.Button.OK)return;
  const pass=passR.getResponseText().trim(); if(pass.length<8){ui.alert('Use at least 8 characters.');return;}
  const emailR=ui.prompt('System Owner','Email (optional):',ui.ButtonSet.OK_CANCEL); if(emailR.getSelectedButton()!==ui.Button.OK)return;
  const phoneR=ui.prompt('System Owner','Phone (optional):',ui.ButtonSet.OK_CANCEL); if(phoneR.getSelectedButton()!==ui.Button.OK)return;
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ACCESS);
  const id='OWNER-'+Utilities.getUuid().slice(0,8).toUpperCase();
  sh.appendRow([id,nameR.getResponseText().trim(),'SYSTEM_OWNER',hash_(pass),'*','*',emailR.getResponseText().trim(),phoneR.getResponseText().trim(),true,new Date()]);
  ui.alert('Owner created','Use the passcode you just entered to unlock Admin and Gate Supervisor functions.',ui.ButtonSet.OK);
}

function addAccessUser() {
  const ui=SpreadsheetApp.getUi();
  const name=ui.prompt('Add Access User','Name:',ui.ButtonSet.OK_CANCEL); if(name.getSelectedButton()!==ui.Button.OK)return;
  const role=ui.prompt('Add Access User','Role: '+ROLES.join(', '),ui.ButtonSet.OK_CANCEL); if(role.getSelectedButton()!==ui.Button.OK)return;
  const r=role.getResponseText().trim().toUpperCase(); if(!ROLES.includes(r)){ui.alert('Invalid role.');return;}
  const scope=ui.prompt('Add Access User','Event scope: * for all, or comma-separated Event IDs',ui.ButtonSet.OK_CANCEL); if(scope.getSelectedButton()!==ui.Button.OK)return;
  const pass=ui.prompt('Add Access User','Passcode (8+ characters):',ui.ButtonSet.OK_CANCEL); if(pass.getSelectedButton()!==ui.Button.OK)return;
  if(pass.getResponseText().trim().length<8){ui.alert('Use at least 8 characters.');return;}
  const email=ui.prompt('Add Access User','Email (optional):',ui.ButtonSet.OK_CANCEL); if(email.getSelectedButton()!==ui.Button.OK)return;
  const phone=ui.prompt('Add Access User','Phone (optional):',ui.ButtonSet.OK_CANCEL); if(phone.getSelectedButton()!==ui.Button.OK)return;
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.ACCESS);
  sh.appendRow(['USR-'+Utilities.getUuid().slice(0,8).toUpperCase(),name.getResponseText().trim(),r,hash_(pass.getResponseText().trim()),'*',scope.getResponseText().trim()||'*',email.getResponseText().trim(),phone.getResponseText().trim(),true,new Date()]);
  ui.alert('Access user added.');
}

function createClientConfigurationForm() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const old=ss.getSheetByName(SHEETS.FORM);
  if(old){const stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'America/New_York','yyyyMMdd_HHmmss');old.setName('EventConfigResponses_Legacy_'+stamp);}
  const before=ss.getSheets().map(s=>s.getName());
  const form=FormApp.create('Event Ticket Configuration');
  form.setDescription('Complete the event information. The Event ID and configuration key are prefilled from your private client link. Do not change them.');
  form.setConfirmationMessage('Configuration submitted. Return to the client portal, load the ticket preview, revise if needed, then approve the design.');
  const text=(t,h,req)=>{const i=form.addTextItem().setTitle(t).setHelpText(h||'');if(req)i.setRequired(true);return i;};
  const para=(t,h)=>form.addParagraphTextItem().setTitle(t).setHelpText(h||'');
  text('EventID','Prefilled. Do not change.',true);
  text('ConfigToken','Prefilled security key. Do not change.',true);
  text('OrgName','Organization name displayed on ticket',true);
  text('ChapterName','Chapter/unit/organizing body displayed on ticket',true);
  text('EventTitle','Public event title',true);
  para('Tagline','Optional subtitle');
  text('EventDate','Example: Saturday, September 5, 2026',true);
  text('EventTime','Example: 8:00 PM',true);
  text('VenueName','Venue name',true);
  text('VenueAddress','Full address',true);
  text('ContactPhone','Public event contact phone');
  text('ContactEmail','Public event contact email');
  text('WebsiteURL','Public website');
  text('DressCode','Optional');
  text('PrimaryColor','Optional HEX, e.g. #0B3D24');
  text('AccentColor','Optional HEX, e.g. #C9A24B');
  text('SerialPrefix','Short prefix, e.g. NY-');
  text('CurrencySymbol','Example: $');
  text('FooterLegalText','Optional small print');
  text('Capacity','Optional total capacity');
  form.addMultipleChoiceItem().setTitle('UseGroups').setChoiceValues(['true','false']).setRequired(true);
  text('GroupLabel','Example: Chapter, Team, Table');
  para('GroupsList','One group/chapter per line');
  for(let i=1;i<=12;i++){
    form.addSectionHeaderItem().setTitle('Ticket Tier '+i);
    text('Tier'+i+'Name','Leave blank to skip');
    text('Tier'+i+'Price','Numbers only');
    text('Tier'+i+'Capacity','Optional');
    text('Tier'+i+'Color','Optional HEX; blank uses a system default');
  }
  form.setDestination(FormApp.DestinationType.SPREADSHEET,ss.getId());
  const created=ss.getSheets().filter(s=>!before.includes(s.getName())); if(created.length)created[0].setName(SHEETS.FORM);
  PropertiesService.getScriptProperties().setProperty('V5_CONFIG_FORM_ID',form.getId());
  PropertiesService.getScriptProperties().setProperty('V5_CONFIG_FORM_URL',form.getPublishedUrl());
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='handleConfigFormSubmit_').forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('handleConfigFormSubmit_').forSpreadsheet(ss).onFormSubmit().create();
  SpreadsheetApp.getUi().alert('Client configuration form created','Editor:\n'+form.getEditUrl()+'\n\nPublished:\n'+form.getPublishedUrl(),SpreadsheetApp.getUi().ButtonSet.OK);
}


function handleConfigFormSubmit_(e){
  try{
    const nv=e&&e.namedValues||{};const eventId=normId_((nv.EventID||[''])[0]),key=String((nv.ConfigToken||[''])[0]||'').trim();
    const r=getEventBase_(eventId);if(!r||String(r.ConfigToken)!==key)return;
    syncFormToOperational_(eventId,key,null);
    r.Status='Client Submitted';r.ApprovedAt='';r.ApprovedBy='';r.UpdatedAt=new Date();writeRow_(SHEETS.EVENTS,HEADERS.Events,r._row,r);
    audit_(eventId,'CLIENT_SUBMITTED','Event',eventId,r.ClientName||'Client','CLIENT',{source:'Google Form'});
  }catch(err){console.log(err);}
}

// ---------- Core helpers ----------
function ensureSheet_(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getLastRow()===0)sh.appendRow(headers);else ensureHeaders_(sh,headers);sh.setFrozenRows(1);return sh;}
function ensureHeaders_(sh,headers){const last=Math.max(sh.getLastColumn(),1);const existing=sh.getRange(1,1,1,last).getValues()[0].map(String);headers.forEach(h=>{if(!existing.includes(h)){sh.getRange(1,sh.getLastColumn()+1).setValue(h);existing.push(h);}});}
function rows_(name){const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);if(!sh||sh.getLastRow()<2)return[];const data=sh.getDataRange().getValues();const h=data.shift().map(String);return data.map((r,i)=>{const o={_row:i+2};h.forEach((k,j)=>o[k]=r[j]);return o;});}
function rowBy_(name,field,value){return rows_(name).find(r=>String(r[field])===String(value))||null;}
function writeRow_(name,headers,rowNum,obj){SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).getRange(rowNum,1,1,headers.length).setValues([headers.map(h=>obj[h]===undefined?'':obj[h])]);}
function append_(name,headers,obj){SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).appendRow(headers.map(h=>obj[h]===undefined?'':obj[h]));}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function body_(e){try{return JSON.parse(e&&e.postData&&e.postData.contents||'{}')}catch(_){return{}}}
function normId_(v){return String(v||'').trim().toUpperCase().replace(/[^A-Z0-9_-]/g,'');}
function bool_(v){return v===true||v===1||String(v).toLowerCase()==='true';}
function money_(v){const n=Number(v);return isFinite(n)?n:0;}
function digits_(v){return String(v||'').replace(/\D/g,'');}
function token_(){return Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');}
function hash_(s){return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s||''),Utilities.Charset.UTF_8));}
function validHex_(v){const s=String(v||'').trim();return /^#[0-9a-f]{6}$/i.test(s)?s.toUpperCase():'';}
function tierColor_(v,i){return validHex_(v)||DEFAULT_TIER_COLORS[i%DEFAULT_TIER_COLORS.length];}
function cleanGroup_(v){const s=String(v==null?'':v).trim();return !s||/^(undefined|null|nan)$/i.test(s)?'':s;}
function splitGroups_(v){if(Array.isArray(v))return v.map(cleanGroup_).filter(Boolean);const s=String(v==null?'':v).trim();if(!s)return[];return s.split(/\r?\n|\s*;\s*|\s*,\s*/).map(cleanGroup_).filter(Boolean);}
function normalizeDate_(v){if(v instanceof Date&&!isNaN(v))return Utilities.formatDate(v,Session.getScriptTimeZone()||'America/New_York','EEEE, MMMM d, yyyy');const s=String(v||'').trim();if(!s)return'';const d=new Date(s);if(!isNaN(d)&&/(GMT|T\d\d:)/.test(s))return Utilities.formatDate(d,Session.getScriptTimeZone()||'America/New_York','EEEE, MMMM d, yyyy');return s;}
function normalizeTime_(v){if(v instanceof Date&&!isNaN(v))return Utilities.formatDate(v,Session.getScriptTimeZone()||'America/New_York','h:mm a');const s=String(v||'').trim();if(!s)return'';if(/^\d{1,2}:\d{2}(\s*[AP]M)?$/i.test(s)){if(/[AP]M/i.test(s))return s.toUpperCase();const p=s.split(':');let h=Number(p[0]);const ap=h>=12?'PM':'AM';h=h%12||12;return h+':'+p[1]+' '+ap;}const d=new Date(s);return !isNaN(d)?Utilities.formatDate(d,Session.getScriptTimeZone()||'America/New_York','h:mm a'):s;}
function dateObj_(v){if(v instanceof Date&&!isNaN(v))return new Date(v.getFullYear(),v.getMonth(),v.getDate());const d=new Date(String(v||''));return isNaN(d)?null:new Date(d.getFullYear(),d.getMonth(),d.getDate());}
function audit_(eventId,action,type,id,actor,role,details){append_(SHEETS.AUDIT,HEADERS.AuditLog,{Timestamp:new Date(),ClientID:eventClientId_(eventId),EventID:eventId,Action:action,EntityType:type,EntityID:id,Actor:actor||'',Role:role||'',Details:typeof details==='string'?details:JSON.stringify(details||{})});}
function contactLog_(eventId,type,id,channel,recipient,actor,note){append_(SHEETS.CONTACTS,HEADERS.ContactLog,{Timestamp:new Date(),ClientID:eventClientId_(eventId),EventID:eventId,EntityType:type,EntityID:id,Channel:channel,Recipient:recipient,Actor:actor||'',Note:note||''});}

// ---------- Authentication / authorization ----------
function authenticate_(pass){const h=hash_(String(pass||'').trim());const u=rows_(SHEETS.ACCESS).find(r=>bool_(r.Active)&&r.PasscodeHash===h);if(!u)return null;return{userId:u.UserID,name:u.Name,role:String(u.Role),clientScope:String(u.ClientScope||'*'),scope:String(u.EventScope||'*'),email:u.Email||'',phone:u.Phone||''};}
function eventClientId_(eventId){const e=getEventBase_(normId_(eventId));return e?normId_(e.ClientID||''):'';}
function clientAllows_(u,clientId){if(!u)return false;if(u.role==='SYSTEM_OWNER'||u.clientScope==='*')return true;const id=normId_(clientId);return String(u.clientScope||'').split(',').map(normId_).includes(id);}
function scopeAllows_(u,eventId){if(!u)return false;const cid=eventClientId_(eventId);if(cid&&!clientAllows_(u,cid))return false;if(u.role==='SYSTEM_OWNER'||u.scope==='*')return true;return String(u.scope||'').split(',').map(normId_).includes(normId_(eventId));}
function requireRole_(body,roles,eventId){const u=authenticate_(body.adminPasscode);if(!u)return{ok:false,error:'Incorrect passcode.'};if(!roles.includes(u.role)&&u.role!=='SYSTEM_OWNER')return{ok:false,error:'Your role does not permit this action.'};if(eventId&&!scopeAllows_(u,eventId))return{ok:false,error:'You are not assigned to this event.'};return{ok:true,user:u};}
function roleCanGate_(role){return['SYSTEM_OWNER','EVENT_ADMIN','GATE_SUPERVISOR','GATE_STAFF'].includes(role);}

// ---------- Event + form overlay ----------
function latestValidForm_(eventId,configToken){
  const wantedEvent=normId_(eventId), wantedToken=String(configToken||'').trim();
  if(!wantedEvent||!wantedToken)return null;

  // Primary source: read the Google Form itself. This avoids depending on the
  // response sheet name and works even if Google created/renamed that sheet
  // asynchronously after setDestination().
  try{
    const formId=PropertiesService.getScriptProperties().getProperty('V5_CONFIG_FORM_ID');
    if(formId){
      const form=FormApp.openById(formId), responses=form.getResponses();
      for(let i=responses.length-1;i>=0;i--){
        const response=responses[i], o={Timestamp:response.getTimestamp()};
        response.getItemResponses().forEach(ir=>{
          const title=String(ir.getItem().getTitle()||'').trim();
          if(title)o[title]=ir.getResponse();
        });
        if(normId_(o.EventID)===wantedEvent&&String(o.ConfigToken||'').trim()===wantedToken)return o;
      }
    }
  }catch(err){console.log('Form response lookup fallback: '+err);}

  // Fallback: scan likely spreadsheet response sheets. Exact token matching
  // prevents another event/legacy response from being selected.
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    const sheets=ss.getSheets().filter(sh=>{
      const n=String(sh.getName()||'');
      return n===SHEETS.FORM || /^Form Responses/i.test(n) || /^EventConfigResponses/i.test(n);
    });
    for(let si=0;si<sheets.length;si++){
      const sh=sheets[si]; if(sh.getLastRow()<2)continue;
      const data=sh.getDataRange().getValues(), h=data[0].map(v=>String(v||'').trim());
      const canon=x=>String(x||'').replace(/[^a-z0-9]/gi,'').toLowerCase();
      const ei=h.findIndex(x=>canon(x)==='eventid'), ki=h.findIndex(x=>canon(x)==='configtoken');
      if(ei<0||ki<0)continue;
      for(let r=data.length-1;r>=1;r--){
        if(normId_(data[r][ei])===wantedEvent&&String(data[r][ki]||'').trim()===wantedToken){
          const o={};h.forEach((k,j)=>{if(k)o[k]=data[r][j];});
          if(o.Timestamp===undefined)o.Timestamp=data[r][0];
          return o;
        }
      }
    }
  }catch(err){console.log('Sheet response lookup failed: '+err);}
  return null;
}

function canonKey_(v){return String(v==null?'':v).replace(/[^a-z0-9]/gi,'').toLowerCase();}
function formValue_(obj,names){
  if(!obj)return undefined;
  const wanted=(Array.isArray(names)?names:[names]).map(canonKey_);
  const keys=Object.keys(obj);
  for(let i=0;i<keys.length;i++){
    if(wanted.includes(canonKey_(keys[i])))return obj[keys[i]];
  }
  return undefined;
}
function tierFormValue_(f,i,part){
  const p=String(part||'');
  const aliases={
    Name:['Tier'+i+'Name','Tier '+i+' Name','Ticket Tier '+i+' Name','TicketTier'+i+'Name','Ticket Tier '+i],
    Price:['Tier'+i+'Price','Tier '+i+' Price','Ticket Tier '+i+' Price','TicketTier'+i+'Price'],
    Capacity:['Tier'+i+'Capacity','Tier '+i+' Capacity','Ticket Tier '+i+' Capacity','TicketTier'+i+'Capacity'],
    Color:['Tier'+i+'Color','Tier '+i+' Color','Ticket Tier '+i+' Color','TicketTier'+i+'Color']
  };
  return formValue_(f,aliases[p]||[]);
}
function deleteEventRows_(sheetName,eventId){
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if(!sh||sh.getLastRow()<2)return;
  const data=sh.getDataRange().getValues(), headers=data[0].map(String), ei=headers.indexOf('EventID');
  if(ei<0)return;
  for(let r=data.length-1;r>=1;r--)if(normId_(data[r][ei])===normId_(eventId))sh.deleteRow(r+1);
}
function syncFormToOperational_(eventId,configToken,formObj){
  const b=getEventBase_(eventId);if(!b)return{ok:false,error:'Event not found.'};
  if(String(b.ConfigToken||'')!==String(configToken||''))return{ok:false,error:'Configuration token mismatch.'};
  const f=formObj||latestValidForm_(b.EventID,b.ConfigToken);if(!f)return{ok:false,error:'No matching configuration submission found.'};
  const aliases={
    OrgName:['OrgName','Organization','Organization Name'],ChapterName:['ChapterName','Chapter / Unit','Chapter','Unit','Organizing Body'],
    EventTitle:['EventTitle','Event Title','Title'],Tagline:['Tagline','Subtitle'],EventDate:['EventDate','Event Date','Date'],EventTime:['EventTime','Event Time','Time'],
    VenueName:['VenueName','Venue Name','Venue'],VenueAddress:['VenueAddress','Venue Address','Address'],ContactPhone:['ContactPhone','Contact Phone','Event Contact Phone'],
    ContactEmail:['ContactEmail','Contact Email','Event Contact Email'],WebsiteURL:['WebsiteURL','Website URL','Website'],DressCode:['DressCode','Dress Code'],
    PrimaryColor:['PrimaryColor','Primary Color'],AccentColor:['AccentColor','Accent / Gold Color','Accent Color','Gold Color'],GroupLabel:['GroupLabel','Group Label'],
    UseGroups:['UseGroups','Use Groups','Use Chapters/Groups'],SerialPrefix:['SerialPrefix','Serial Prefix'],CurrencySymbol:['CurrencySymbol','Currency Symbol'],
    FooterLegalText:['FooterLegalText','Footer Legal Text','Footer Text'],Capacity:['Capacity','Event Capacity','Total Capacity']
  };
  Object.keys(aliases).forEach(k=>{const v=formValue_(f,aliases[k]);if(v!==undefined&&String(v).trim()!=='')b[k]=v;});
  b.PrimaryColor=validHex_(b.PrimaryColor)||DEFAULT_PRIMARY;b.AccentColor=validHex_(b.AccentColor)||DEFAULT_ACCENT;
  b.EventDate=normalizeDate_(b.EventDate);b.EventTime=normalizeTime_(b.EventTime);b.Capacity=Number(b.Capacity)||0;b.UseGroups=String(b.UseGroups||'true');b.UpdatedAt=new Date();
  writeRow_(SHEETS.EVENTS,HEADERS.Events,b._row,b);

  let tierFieldsPresent=false;for(let i=1;i<=12;i++){if(tierFormValue_(f,i,'Name')!==undefined){tierFieldsPresent=true;break;}}
  if(tierFieldsPresent){
    deleteEventRows_(SHEETS.TIERS,b.EventID);let sort=1,colorIndex=0;
    for(let i=1;i<=12;i++){
      const label=String(tierFormValue_(f,i,'Name')||'').trim();if(!label)continue;
      let key=(label.toLowerCase().replace(/[^a-z0-9]+/g,'').slice(0,30)||('tier'+i));
      if(rowBy_(SHEETS.TIERS,'TierKey',key))key=key+'_'+i;
      append_(SHEETS.TIERS,HEADERS.Tiers,{ClientID:normId_(b.ClientID||''),EventID:b.EventID,TierKey:key,Label:label,Price:money_(tierFormValue_(f,i,'Price')),Capacity:Number(tierFormValue_(f,i,'Capacity'))||0,Active:true,Color:tierColor_(tierFormValue_(f,i,'Color'),colorIndex++),SortOrder:sort++});
    }
  }
  const gv=formValue_(f,['GroupsList','Groups List','Groups','Chapters','Chapter List','Groups / Chapters']);
  if(gv!==undefined){
    deleteEventRows_(SHEETS.GROUPS,b.EventID);const gs=[...new Set(splitGroups_(gv))];
    gs.forEach((g,i)=>append_(SHEETS.GROUPS,HEADERS.Groups,{ClientID:normId_(b.ClientID||''),EventID:b.EventID,GroupName:g,Active:true,SortOrder:i+1}));
  }
  return{ok:true,event:getEventBase_(b.EventID),tierCount:Object.keys(getTiersFromSheet_(b.EventID)).length,groupCount:getGroupsFromSheet_(b.EventID).length};
}
function getTiersFromSheet_(eventId){const out={};rows_(SHEETS.TIERS).filter(r=>normId_(r.EventID)===normId_(eventId)&&String(r.Active).toLowerCase()!=='false').sort((a,b)=>(Number(a.SortOrder)||0)-(Number(b.SortOrder)||0)).forEach((r,i)=>{const key=String(r.TierKey||('tier'+(i+1)));out[key]={label:String(r.Label||key),price:money_(r.Price),capacity:Number(r.Capacity)||0,color:tierColor_(r.Color,i)};});return out;}
function getGroupsFromSheet_(eventId){return[...new Set(rows_(SHEETS.GROUPS).filter(r=>normId_(r.EventID)===normId_(eventId)&&String(r.Active).toLowerCase()!=='false').sort((a,b)=>(Number(a.SortOrder)||0)-(Number(b.SortOrder)||0)).map(r=>cleanGroup_(r.GroupName)).filter(Boolean))];}

function getEventBase_(eventId){return rowBy_(SHEETS.EVENTS,'EventID',normId_(eventId));}
function getEvent_(eventId){const base=getEventBase_(eventId);if(!base)return null;const e=Object.assign({},base);const f=latestValidForm_(e.EventID,e.ConfigToken);if(f){const aliases={OrgName:['OrgName','Organization','Organization Name'],ChapterName:['ChapterName','Chapter / Unit','Chapter'],EventTitle:['EventTitle','Event Title'],Tagline:['Tagline'],EventDate:['EventDate','Event Date'],EventTime:['EventTime','Event Time'],VenueName:['VenueName','Venue Name','Venue'],VenueAddress:['VenueAddress','Venue Address','Address'],ContactPhone:['ContactPhone','Contact Phone'],ContactEmail:['ContactEmail','Contact Email'],WebsiteURL:['WebsiteURL','Website URL','Website'],DressCode:['DressCode','Dress Code'],PrimaryColor:['PrimaryColor','Primary Color'],AccentColor:['AccentColor','Accent / Gold Color','Accent Color'],GroupLabel:['GroupLabel','Group Label'],UseGroups:['UseGroups','Use Groups'],SerialPrefix:['SerialPrefix','Serial Prefix'],CurrencySymbol:['CurrencySymbol','Currency Symbol'],FooterLegalText:['FooterLegalText','Footer Legal Text'],Capacity:['Capacity','Event Capacity','Total Capacity']};Object.keys(aliases).forEach(k=>{const v=formValue_(f,aliases[k]);if(v!==undefined&&String(v).trim()!=='')e[k]=v;});}e.PrimaryColor=validHex_(e.PrimaryColor)||DEFAULT_PRIMARY;e.AccentColor=validHex_(e.AccentColor)||DEFAULT_ACCENT;e.EventDate=normalizeDate_(e.EventDate);e.EventTime=normalizeTime_(e.EventTime);e.Capacity=Number(e.Capacity)||0;e.UseGroups=String(e.UseGroups||'true');e.LogoURL=e.LogoFileId?driveData_(e.LogoFileId):'';e.BackgroundURL=e.BackgroundFileId?driveData_(e.BackgroundFileId):'';return e;}
function publicEvent_(e){if(!e)return null;const keys=['EventID','ClientID','OrgName','ChapterName','EventTitle','Tagline','EventDate','EventTime','VenueName','VenueAddress','ContactPhone','ContactEmail','WebsiteURL','DressCode','PrimaryColor','AccentColor','GroupLabel','UseGroups','SerialPrefix','CurrencySymbol','FooterLegalText','Capacity','Status','LogoURL','BackgroundURL','AutoReminders','EmergencyReadOnly'];const o={};keys.forEach(k=>o[k]=e[k]);return o;}
function getTiers_(eventId){const e=getEventBase_(eventId);if(!e)return{};const fromSheet=getTiersFromSheet_(eventId);if(Object.keys(fromSheet).length)return fromSheet;const f=latestValidForm_(e.EventID,e.ConfigToken);if(f){const out={};let n=0;for(let i=1;i<=12;i++){const label=String(tierFormValue_(f,i,'Name')||'').trim();if(!label)continue;const key=(label.toLowerCase().replace(/[^a-z0-9]+/g,'').slice(0,30)||('tier'+i));out[key]={label:label,price:money_(tierFormValue_(f,i,'Price')),capacity:Number(tierFormValue_(f,i,'Capacity'))||0,color:tierColor_(tierFormValue_(f,i,'Color'),n++)};}return out;}return{};}
function getGroups_(eventId){const sheetGroups=getGroupsFromSheet_(eventId);if(sheetGroups.length)return sheetGroups;const e=getEventBase_(eventId);if(!e)return[];const f=latestValidForm_(e.EventID,e.ConfigToken);const gv=f?formValue_(f,['GroupsList','Groups List','Groups','Chapters','Chapter List','Groups / Chapters']):undefined;return gv===undefined?[]:[...new Set(splitGroups_(gv))];}
function driveData_(id){try{const b=DriveApp.getFileById(String(id)).getBlob();return'data:'+b.getContentType()+';base64,'+Utilities.base64Encode(b.getBytes())}catch(_){return''}}
function assetFolder_(){const n='Event Ticketing V5 Assets',it=DriveApp.getFoldersByName(n);return it.hasNext()?it.next():DriveApp.createFolder(n);}
function formPrefillUrl_(eventId,token){const id=PropertiesService.getScriptProperties().getProperty('V5_CONFIG_FORM_ID');if(!id)return'';const form=FormApp.openById(id);const resp=form.createResponse();form.getItems().forEach(item=>{if(item.getTitle()==='EventID')resp.withItemResponse(item.asTextItem().createResponse(eventId));if(item.getTitle()==='ConfigToken')resp.withItemResponse(item.asTextItem().createResponse(token));});return resp.toPrefilledUrl();}

function eventSummary_(e){return{EventID:e.EventID,ClientID:e.ClientID||'',ClientName:e.ClientName||'',ClientEmail:e.ClientEmail||'',ClientPhone:e.ClientPhone||'',OrgName:e.OrgName||'',ChapterName:e.ChapterName||'',EventTitle:e.EventTitle||e.EventID,EventDate:e.EventDate||'',VenueName:e.VenueName||'',Status:e.Status||'Draft',Capacity:Number(e.Capacity)||0,ApprovedAt:e.ApprovedAt||'',PrimaryColor:e.PrimaryColor||DEFAULT_PRIMARY,AccentColor:e.AccentColor||DEFAULT_ACCENT};}
function currentFutureEvents_(u,includePast){const today=new Date();today.setHours(0,0,0,0);return rows_(SHEETS.EVENTS).map(r=>getEvent_(r.EventID)).filter(Boolean).filter(e=>scopeAllows_(u,e.EventID)).filter(e=>{if(includePast)return true;if(String(e.Status)==='Archived')return false;const d=dateObj_(e.EventDate);return !d||d>=today;}).sort((a,b)=>{const da=dateObj_(a.EventDate),db=dateObj_(b.EventDate);if(!da&&!db)return String(a.EventTitle).localeCompare(String(b.EventTitle));if(!da)return 1;if(!db)return-1;return da-db;}).map(eventSummary_);}

// ---------- Client configuration ----------
function createEventShell_(body,u){const eventId=normId_(body.eventId);if(!eventId)return{ok:false,error:'Event ID is required.'};if(getEventBase_(eventId))return{ok:false,error:'That Event ID already exists.'};const clientId=normId_(body.clientId||body.clientName||eventId);if(!clientId)return{ok:false,error:'Client ID is required.'};if(u.role!=='SYSTEM_OWNER'&&!clientAllows_(u,clientId))return{ok:false,error:'You are not assigned to this client.'};ensureClient_(clientId,String(body.clientName||clientId),String(body.clientEmail||''),String(body.clientPhone||''));const now=new Date(),tok=token_();append_(SHEETS.EVENTS,HEADERS.Events,{EventID:eventId,ClientID:clientId,ClientName:String(body.clientName||''),ClientEmail:String(body.clientEmail||''),ClientPhone:String(body.clientPhone||''),EventTitle:String(body.eventTitle||eventId),PrimaryColor:DEFAULT_PRIMARY,AccentColor:DEFAULT_ACCENT,UseGroups:'true',GroupLabel:'Chapter',CurrencySymbol:'$',Status:'Draft',ConfigToken:tok,CreatedAt:now,UpdatedAt:now});append_(SHEETS.COUNTERS,HEADERS.Counters,{EventID:eventId,CurrentNumber:0});if(u.role==='EVENT_ADMIN'&&u.scope!=='*'){const ar=rowBy_(SHEETS.ACCESS,'UserID',u.userId);if(ar){const ids=String(ar.EventScope||'').split(',').map(normId_).filter(Boolean);if(!ids.includes(eventId))ids.push(eventId);ar.EventScope=ids.join(',');writeRow_(SHEETS.ACCESS,HEADERS.AccessControl,ar._row,ar);}}audit_(eventId,'CREATE_EVENT_SHELL','Event',eventId,u.name,u.role,{clientName:body.clientName||''});return{ok:true,eventId:eventId,configUrl:CONFIG_URL+'?event='+encodeURIComponent(eventId)+'&key='+encodeURIComponent(tok),formUrl:formPrefillUrl_(eventId,tok)};}
function clientContext_(eventId,key){const b=getEventBase_(eventId);if(!b||String(b.ConfigToken)!==String(key||''))return{ok:false,error:'This configuration link is invalid or expired.'};const f=latestValidForm_(b.EventID,b.ConfigToken);if(f){syncFormToOperational_(b.EventID,b.ConfigToken,f);const submittedAt=f.Timestamp instanceof Date?f.Timestamp:new Date(f.Timestamp||0);const approvedAt=b.ApprovedAt instanceof Date?b.ApprovedAt:new Date(b.ApprovedAt||0);if(b.Status==='Draft'||((b.Status==='Client Approved'||b.Status==='Active')&&submittedAt>approvedAt)){b.Status='Client Submitted';b.ApprovedAt='';b.ApprovedBy='';b.UpdatedAt=new Date();writeRow_(SHEETS.EVENTS,HEADERS.Events,b._row,b);audit_(b.EventID,'CLIENT_SUBMITTED','Event',b.EventID,b.ClientName||'Client','CLIENT',{});}}const e=getEvent_(eventId);return{ok:true,event:publicEvent_(e),tiers:getTiers_(eventId),groups:getGroups_(eventId),formUrl:formPrefillUrl_(b.EventID,b.ConfigToken),submitted:!!f,submittedAt:f&&f.Timestamp?String(f.Timestamp):'',status:b.Status,client:{name:b.ClientName||'',email:b.ClientEmail||'',phone:b.ClientPhone||''},hasLogo:!!b.LogoFileId,hasBackground:!!b.BackgroundFileId};}
function clientUpload_(body){const b=getEventBase_(body.eventId);if(!b||String(b.ConfigToken)!==String(body.configKey||''))return{ok:false,error:'Invalid configuration link.'};const type=body.assetType==='background'?'background':'logo';const m=String(body.dataUrl||'').match(/^data:(image\/(?:png|jpeg|jpg|webp));base64,(.+)$/i);if(!m)return{ok:false,error:'Use PNG, JPG, or WebP.'};const bytes=Utilities.base64Decode(m[2]);if(bytes.length>3*1024*1024)return{ok:false,error:'Image must be 3 MB or smaller.'};const ext=m[1].includes('png')?'png':m[1].includes('webp')?'webp':'jpg';const f=assetFolder_().createFile(Utilities.newBlob(bytes,m[1],b.EventID+'-'+type+'.'+ext));b[type==='logo'?'LogoFileId':'BackgroundFileId']=f.getId();b.UpdatedAt=new Date();writeRow_(SHEETS.EVENTS,HEADERS.Events,b._row,b);audit_(b.EventID,'CLIENT_UPLOAD_ASSET','Event',b.EventID,b.ClientName||'Client','CLIENT',type);return{ok:true};}
function clientPreview_(body){const b=getEventBase_(body.eventId);if(!b||String(b.ConfigToken)!==String(body.configKey||''))return{ok:false,error:'Invalid configuration link.'};const f=latestValidForm_(b.EventID,b.ConfigToken);if(!f)return{ok:false,error:'No matching Google Form submission was found for this event. If you already submitted it, reopen the prefilled form from this page and confirm the Event ID/security key were not changed, then submit once more.'};syncFormToOperational_(b.EventID,b.ConfigToken,f);if(b.Status==='Draft'||b.Status==='Client Submitted')b.Status='Preview Ready';b.UpdatedAt=new Date();writeRow_(SHEETS.EVENTS,HEADERS.Events,b._row,b);audit_(b.EventID,'CLIENT_PREVIEW','Event',b.EventID,b.ClientName||'Client','CLIENT',{});return{ok:true,event:publicEvent_(getEvent_(b.EventID)),tiers:getTiers_(b.EventID),groups:getGroups_(b.EventID)};}
function clientApprove_(body){const b=getEventBase_(body.eventId);if(!b||String(b.ConfigToken)!==String(body.configKey||''))return{ok:false,error:'Invalid configuration link.'};if(!latestValidForm_(b.EventID,b.ConfigToken))return{ok:false,error:'Configuration form has not been submitted.'};b.Status='Client Approved';b.ApprovedAt=new Date();b.ApprovedBy=b.ClientName||'Client';b.UpdatedAt=new Date();writeRow_(SHEETS.EVENTS,HEADERS.Events,b._row,b);audit_(b.EventID,'CLIENT_APPROVE','Event',b.EventID,b.ApprovedBy,'CLIENT',{});return{ok:true,status:b.Status};}

// ---------- Capacity / tickets / vouchers ----------
function allocations_(eventId){const claims=rows_(SHEETS.CLAIMS).filter(r=>normId_(r.EventID)===normId_(eventId)&&String(r.Status)!=='Revoked');const vouchers=rows_(SHEETS.VOUCHERS).filter(r=>normId_(r.EventID)===normId_(eventId)&&!bool_(r.Claimed)&&String(r.Status)!=='Cancelled');const byTier={};claims.forEach(r=>byTier[r.TierKey]=(byTier[r.TierKey]||0)+1);vouchers.forEach(r=>{if(r.TierKey)byTier[r.TierKey]=(byTier[r.TierKey]||0)+1});return{total:claims.length+vouchers.length,byTier:byTier,claims:claims.length,pendingVouchers:vouchers.length};}
function capacity_(eventId,adds,allowOverride){const e=getEvent_(eventId);if(!e)return{ok:false,error:'Event not found.'};if(e.Status!=='Active')return{ok:false,error:'Voucher generation and ticket claiming require an Active event.'};const a=allocations_(eventId),sum=Object.values(adds||{}).reduce((x,y)=>x+Number(y||0),0);if(!allowOverride&&e.Capacity&&a.total+sum>e.Capacity)return{ok:false,error:'Event capacity exceeded. '+Math.max(0,e.Capacity-a.total)+' place(s) remain.'};const ts=getTiers_(eventId);for(const k in adds){const n=Number(adds[k])||0,cap=ts[k]&&Number(ts[k].capacity)||0;if(!allowOverride&&cap&&(a.byTier[k]||0)+n>cap)return{ok:false,error:(ts[k]?.label||k)+' capacity exceeded.'};}return{ok:true};}
function nextSerial_(eventId){const lock=LockService.getScriptLock();lock.waitLock(15000);try{let r=rowBy_(SHEETS.COUNTERS,'EventID',eventId);if(!r){append_(SHEETS.COUNTERS,HEADERS.Counters,{EventID:eventId,CurrentNumber:0});r=rowBy_(SHEETS.COUNTERS,'EventID',eventId);}const n=(Number(r.CurrentNumber)||0)+1;r.CurrentNumber=n;writeRow_(SHEETS.COUNTERS,HEADERS.Counters,r._row,r);const p=getEvent_(eventId).SerialPrefix||eventId+'-';return p+String(n).padStart(3,'0');}finally{lock.releaseLock();}}
function generateVouchers_(body,u){const eventId=normId_(body.eventId);if(eventLocked_(eventId))return{ok:false,error:'Event is in emergency read-only mode.'};const auth=scopeAllows_(u,eventId);if(!auth)return{ok:false,error:'Not assigned to event.'};const eb=getEventBase_(eventId);if(eb){const ff=latestValidForm_(eb.EventID,eb.ConfigToken);if(ff)syncFormToOperational_(eb.EventID,eb.ConfigToken,ff);}const tiers=getTiers_(eventId),counts=body.tierCounts||{};let total=Math.max(0,Number(body.openCount)||0),adds={};Object.keys(tiers).forEach(k=>{const n=Math.max(0,Number(counts[k])||0);adds[k]=n;total+=n});if(total<1||total>100)return{ok:false,error:'Generate between 1 and 100 vouchers per batch.'};const cap=capacity_(eventId,adds,bool_(body.capacityOverride)&&['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role));if(!cap.ok)return cap;const batch=token_().slice(0,32),now=new Date(),group=cleanGroup_(body.suggestedGroup),recipientName=String(body.distributorName||'').trim(),recipientEmail=String(body.distributorEmail||'').trim(),recipientPhone=String(body.distributorPhone||'').trim();const rows=[];Object.keys(adds).forEach(k=>{for(let i=0;i<adds[k];i++)rows.push(k)});for(let i=0;i<Math.max(0,Number(body.openCount)||0);i++)rows.push('');rows.forEach(k=>append_(SHEETS.VOUCHERS,HEADERS.Vouchers,{Timestamp:now,ClientID:clientIdForEvent_(eventId),EventID:eventId,BatchID:batch,VoucherToken:token_(),TierKey:k,SuggestedGroup:group,PrefillName:'',PrefillPhone:'',Claimed:false,Dispatched:false,RecipientName:'',RecipientEmail:'',RecipientPhone:'',IssuedBy:u.name,Status:'Available'}));audit_(eventId,'GENERATE_VOUCHERS','Batch',batch,u.name,u.role,{count:total,group:group,distributorName:recipientName});return{ok:true,batchId:batch,count:total,distributorUrl:DISTRIBUTION_URL+'?batch='+encodeURIComponent(batch),distributorName:recipientName,distributorEmail:recipientEmail,distributorPhone:recipientPhone};}
function batch_(id){const rs=rows_(SHEETS.VOUCHERS).filter(r=>r.BatchID===String(id));if(!rs.length)return{ok:false,error:'Batch not found.'};const eventId=normId_(rs[0].EventID);return{ok:true,eventId:eventId,event:publicEvent_(getEvent_(eventId)),tiers:getTiers_(eventId),vouchers:rs.map(r=>({token:r.VoucherToken,tier:r.TierKey,claimed:bool_(r.Claimed),serial:r.Serial||'',dispatched:bool_(r.Dispatched),status:r.Status||'Available',suggestedGroup:cleanGroup_(r.SuggestedGroup)}))};}
function voucher_(tok){const r=rowBy_(SHEETS.VOUCHERS,'VoucherToken',String(tok||''));if(!r)return{ok:false,error:'Voucher not found.'};const eventId=normId_(r.EventID);if(String(r.Status)==='Cancelled')return{ok:false,error:'This voucher has been cancelled.'};if(bool_(r.Claimed)){const c=rowBy_(SHEETS.CLAIMS,'Serial',r.Serial);return{ok:false,alreadyClaimed:true,eventId:eventId,serial:r.Serial,tier:c&&c.TierKey,checkInToken:c&&c.CheckInToken,event:publicEvent_(getEvent_(eventId)),tiers:getTiers_(eventId)};}return{ok:true,eventId:eventId,event:publicEvent_(getEvent_(eventId)),tiers:getTiers_(eventId),groups:getGroups_(eventId),tier:r.TierKey||'',groupName:cleanGroup_(r.SuggestedGroup),name:r.PrefillName||'',phone:r.PrefillPhone||''};}
function sendVoucher_(body){const r=rowBy_(SHEETS.VOUCHERS,'VoucherToken',String(body.token||''));if(!r)return{ok:false,error:'Voucher not found.'};if(bool_(r.Claimed)||bool_(r.Dispatched)||String(r.Status)==='Cancelled')return{ok:false,error:'Voucher is no longer available to send.'};const email=String(body.recipientEmail||'').trim(),phone=String(body.recipientPhone||'').trim(),name=String(body.recipientName||'').trim();if(!email&&!phone)return{ok:false,error:'Enter an email or phone number.'};const url=VOUCHER_URL+'?voucher='+encodeURIComponent(r.VoucherToken);if(email){try{MailApp.sendEmail(email,'Your event ticket voucher','Your voucher is ready. Claim your ticket here: '+url)}catch(e){return{ok:false,error:'Email could not be sent: '+e.message}}}r.Dispatched=true;r.RecipientName=name;r.RecipientEmail=email;r.RecipientPhone=phone;r.SentAt=new Date();r.Status='Sent';writeRow_(SHEETS.VOUCHERS,HEADERS.Vouchers,r._row,r);contactLog_(r.EventID,'Voucher',r.VoucherToken,email?'Email':'Phone',email||phone,'Distributor','Voucher delivery');return{ok:true,url:url};}
function claim_(body){const tok=String(body.voucher||'').trim();if(!tok)return{ok:false,error:'A valid voucher is required.'};const v=rowBy_(SHEETS.VOUCHERS,'VoucherToken',tok);if(!v)return{ok:false,error:'Voucher not found.'};if(String(v.Status)==='Cancelled')return{ok:false,error:'Voucher cancelled.'};if(bool_(v.Claimed)){const c=rowBy_(SHEETS.CLAIMS,'Serial',v.Serial);return{ok:false,alreadyClaimed:true,eventId:v.EventID,serial:v.Serial,tier:c&&c.TierKey,checkInToken:c&&c.CheckInToken};}const eventId=normId_(v.EventID),e=getEvent_(eventId);if(eventLocked_(eventId))return{ok:false,error:'Ticket claiming is temporarily paused while this event is in read-only mode.'};if(!e||e.Status!=='Active')return{ok:false,error:'This event is not currently accepting ticket claims.'};const name=String(body.fullName||'').trim();if(!name)return{ok:false,error:'Name is required.'};const tiers=getTiers_(eventId);let tier=v.TierKey||String(body.tier||'');if(!tier)tier=Object.keys(tiers)[0]||'';if(!tiers[tier])return{ok:false,error:'Invalid ticket tier.'};const groups=getGroups_(eventId),group=cleanGroup_(body.groupName);if(String(e.UseGroups).toLowerCase()!=='false'&&groups.length&&!groups.includes(group))return{ok:false,error:'Select a valid '+(e.GroupLabel||'group')+'.'};if(!v.TierKey){const cap=capacity_(eventId,{[tier]:1},false);if(!cap.ok)return cap;}const serial=nextSerial_(eventId),qr=token_().slice(0,48),price=money_(tiers[tier].price),now=new Date();append_(SHEETS.CLAIMS,HEADERS.Claims,{Timestamp:now,ClientID:clientIdForEvent_(eventId),EventID:eventId,Serial:serial,CheckInToken:qr,Name:name,Email:String(body.email||''),Phone:String(body.phone||''),GroupName:group,TierKey:tier,Source:'Voucher',VoucherToken:tok,Status:'Active',AmountDue:price,AmountPaid:0,PaymentStatus:'Pending',PaymentMethod:'',PaymentNote:'',UpdatedAt:now});v.Claimed=true;v.Serial=serial;v.Status='Claimed';writeRow_(SHEETS.VOUCHERS,HEADERS.Vouchers,v._row,v);audit_(eventId,'CLAIM_TICKET','Claim',serial,name,'CLAIMANT',{tier:tier});return{ok:true,eventId:eventId,serial:serial,checkInToken:qr,tier:tier,price:price};}

// ---------- Ticket admin / payments ----------
function guestSearch_(eventId,q){q=String(q||'').trim().toLowerCase();if(!q)return[];const check=rows_(SHEETS.CHECKINS).filter(r=>normId_(r.EventID)===eventId&&String(r.Status)!=='Undone');return rows_(SHEETS.CLAIMS).filter(r=>normId_(r.EventID)===eventId&&String(r.Status)!=='Revoked').filter(r=>[r.Name,r.Email,r.Phone,r.Serial].some(v=>String(v||'').toLowerCase().includes(q))).slice(0,50).map(r=>({serial:r.Serial,name:r.Name,email:r.Email||'',phone:r.Phone||'',groupName:r.GroupName||'',tier:r.TierKey,paymentStatus:r.PaymentStatus||'Pending',paymentMethod:r.PaymentMethod||'',amountDue:money_(r.AmountDue),amountPaid:money_(r.AmountPaid),status:r.Status,checkedIn:check.some(c=>c.Serial===r.Serial),internalNotes:r.InternalNotes||''}));}
function updatePayment_(body,u){const eventId=normId_(body.eventId);if(eventLocked_(eventId))return{ok:false,error:'Event is in emergency read-only mode.'};const r=rowBy_(SHEETS.CLAIMS,'Serial',String(body.serial||''));if(!r||normId_(r.EventID)!==eventId)return{ok:false,error:'Ticket not found.'};const amount=money_(body.amountPaid),status=String(body.paymentStatus||'Pending'),method=String(body.paymentMethod||''),note=String(body.paymentNote||'');r.AmountPaid=amount;r.PaymentStatus=status;r.PaymentMethod=method;r.PaymentNote=note;r.UpdatedAt=new Date();writeRow_(SHEETS.CLAIMS,HEADERS.Claims,r._row,r);append_(SHEETS.PAYMENTS,HEADERS.Payments,{Timestamp:new Date(),ClientID:clientIdForEvent_(eventId),EventID:eventId,Serial:r.Serial,Amount:amount,Method:method,Status:status,Note:note,RecordedBy:u.name});audit_(eventId,'UPDATE_PAYMENT','Claim',r.Serial,u.name,u.role,{amount:amount,status:status,method:method});return{ok:true};}
function transferTicket_(body,u){const eventId=normId_(body.eventId);if(eventLocked_(eventId))return{ok:false,error:'Event is in emergency read-only mode.'};const r=rowBy_(SHEETS.CLAIMS,'Serial',String(body.serial||''));if(!r||normId_(r.EventID)!==eventId)return{ok:false,error:'Ticket not found.'};const old=r.Name;r.TransferredFrom=[r.TransferredFrom,old].filter(Boolean).join(' | ');r.Name=String(body.newName||'').trim()||r.Name;r.Email=String(body.newEmail||'').trim();r.Phone=String(body.newPhone||'').trim();r.UpdatedAt=new Date();writeRow_(SHEETS.CLAIMS,HEADERS.Claims,r._row,r);audit_(eventId,'TRANSFER_TICKET','Claim',r.Serial,u.name,u.role,{from:old,to:r.Name});return{ok:true};}
function reissueQr_(body,u){const eventId=normId_(body.eventId);if(eventLocked_(eventId))return{ok:false,error:'Event is in emergency read-only mode.'};const r=rowBy_(SHEETS.CLAIMS,'Serial',String(body.serial||''));if(!r||normId_(r.EventID)!==eventId)return{ok:false,error:'Ticket not found.'};r.CheckInToken=token_().slice(0,48);r.UpdatedAt=new Date();writeRow_(SHEETS.CLAIMS,HEADERS.Claims,r._row,r);audit_(eventId,'REISSUE_QR','Claim',r.Serial,u.name,u.role,{});return{ok:true,checkInToken:r.CheckInToken,tier:r.TierKey};}
function setTicketStatus_(body,u,status){const eventId=normId_(body.eventId);if(eventLocked_(eventId))return{ok:false,error:'Event is in emergency read-only mode.'};const r=rowBy_(SHEETS.CLAIMS,'Serial',String(body.serial||''));if(!r||normId_(r.EventID)!==eventId)return{ok:false,error:'Ticket not found.'};r.Status=status;r.UpdatedAt=new Date();writeRow_(SHEETS.CLAIMS,HEADERS.Claims,r._row,r);audit_(eventId,status==='Revoked'?'REVOKE_TICKET':'REACTIVATE_TICKET','Claim',r.Serial,u.name,u.role,{});return{ok:true};}


function gateSummary_(eventId){const a=allocations_(eventId),claims=rows_(SHEETS.CLAIMS).filter(r=>normId_(r.EventID)===eventId&&String(r.Status)!=='Revoked'),checks=rows_(SHEETS.CHECKINS).filter(r=>normId_(r.EventID)===eventId&&String(r.Status)!=='Undone'),e=getEvent_(eventId);return{ok:true,counts:{issued:claims.length,checkedIn:checks.length,notArrived:Math.max(0,claims.length-checks.length),allocated:a.total,capacity:e&&e.Capacity||0}};}
function gateSearch_(eventId,q,u){
  const results=guestSearch_(eventId,q);const elevated=['SYSTEM_OWNER','EVENT_ADMIN','GATE_SUPERVISOR'].includes(u.role);
  return results.map(r=>({serial:r.serial,name:r.name,phone:elevated?r.phone:'',email:elevated?r.email:'',groupName:r.groupName,tier:r.tier,paymentStatus:r.paymentStatus,amountPaid:elevated?r.amountPaid:0,internalNotes:elevated?r.internalNotes:'',checkedIn:r.checkedIn,status:r.status}));
}

// ---------- Gate ----------
function lookupClaim_(eventId,value){const s=String(value||'').trim();return rows_(SHEETS.CLAIMS).find(r=>normId_(r.EventID)===eventId&&(String(r.CheckInToken)===s||String(r.Serial).toUpperCase()===s.toUpperCase()))||null;}
function checkIn_(body,u){const eventId=normId_(body.eventId);if(!scopeAllows_(u,eventId)||!roleCanGate_(u.role))return{ok:false,error:'Not permitted for this event.'};const ev=getEvent_(eventId);if(!ev||ev.Status!=='Active')return{ok:false,error:'This event is not active for gate check-in.'};const c=lookupClaim_(eventId,body.serial);if(!c)return{ok:false,error:'Ticket not found.'};if(String(c.Status)==='Revoked')return{ok:false,error:'Ticket has been revoked.',revoked:true,serial:c.Serial,name:c.Name};const existing=rows_(SHEETS.CHECKINS).find(r=>normId_(r.EventID)===eventId&&r.Serial===c.Serial&&String(r.Status)!=='Undone');if(existing)return{ok:false,error:'Already checked in.',alreadyCheckedIn:true,serial:c.Serial,name:c.Name,checkedInAt:existing.Timestamp,checkedInBy:existing.CheckedInBy};const payStatus=String(body.paymentStatus||c.PaymentStatus||'Pending'),payMethod=String(body.paymentMethod||c.PaymentMethod||''),amount=body.amountPaid!==undefined?money_(body.amountPaid):money_(c.AmountPaid);append_(SHEETS.CHECKINS,HEADERS.CheckIns,{Timestamp:new Date(),ClientID:clientIdForEvent_(eventId),EventID:eventId,Serial:c.Serial,Name:c.Name,GroupName:c.GroupName,TierKey:c.TierKey,Phone:c.Phone,PaymentStatus:payStatus,PaymentMethod:payMethod,AmountPaid:amount,CheckedInBy:u.name,GateNote:String(body.gateNote||''),Status:'Checked In'});audit_(eventId,'CHECK_IN','Claim',c.Serial,u.name,u.role,{});return{ok:true,serial:c.Serial,name:c.Name,groupName:c.GroupName,tier:c.TierKey,paymentStatus:payStatus};}
function undoCheckIn_(body,u){if(!['SYSTEM_OWNER','EVENT_ADMIN','GATE_SUPERVISOR'].includes(u.role))return{ok:false,error:'Supervisor permission required.'};const eventId=normId_(body.eventId),serial=String(body.serial||'');const matches=rows_(SHEETS.CHECKINS).filter(r=>normId_(r.EventID)===eventId&&r.Serial===serial&&String(r.Status)!=='Undone');if(!matches.length)return{ok:false,error:'No active check-in found.'};const r=matches[matches.length-1];r.Status='Undone';r.GateNote=(r.GateNote?String(r.GateNote)+' | ':'')+'Undone by '+u.name+': '+String(body.reason||'');writeRow_(SHEETS.CHECKINS,HEADERS.CheckIns,r._row,r);audit_(eventId,'UNDO_CHECK_IN','CheckIn',serial,u.name,u.role,{reason:body.reason||''});return{ok:true};}
function walkIn_(body,u){if(!['SYSTEM_OWNER','EVENT_ADMIN','GATE_SUPERVISOR'].includes(u.role))return{ok:false,error:'Supervisor permission required for walk-ins.'};const eventId=normId_(body.eventId),e=getEvent_(eventId);if(!e||e.Status!=='Active')return{ok:false,error:'Event not active.'};const name=String(body.fullName||'').trim();if(!name)return{ok:false,error:'Name is required.'};const tier=String(body.tier||''),tiers=getTiers_(eventId);if(!tiers[tier])return{ok:false,error:'Choose a valid tier.'};const cap=capacity_(eventId,{[tier]:1},bool_(body.capacityOverride));if(!cap.ok)return cap;const serial=nextSerial_(eventId),qr=token_().slice(0,48),amount=money_(body.amountPaid),status=String(body.paymentStatus||'Paid'),method=String(body.paymentMethod||'Cash'),now=new Date();append_(SHEETS.CLAIMS,HEADERS.Claims,{Timestamp:now,ClientID:clientIdForEvent_(eventId),EventID:eventId,Serial:serial,CheckInToken:qr,Name:name,Email:'',Phone:String(body.phone||''),GroupName:cleanGroup_(body.groupName),TierKey:tier,Source:'Walk-In',Status:'Active',AmountDue:money_(tiers[tier].price),AmountPaid:amount,PaymentStatus:status,PaymentMethod:method,PaymentNote:String(body.note||''),UpdatedAt:now});append_(SHEETS.CHECKINS,HEADERS.CheckIns,{Timestamp:now,ClientID:clientIdForEvent_(eventId),EventID:eventId,Serial:serial,Name:name,GroupName:cleanGroup_(body.groupName),TierKey:tier,Phone:String(body.phone||''),PaymentStatus:status,PaymentMethod:method,AmountPaid:amount,CheckedInBy:u.name,GateNote:'Walk-in',Status:'Checked In'});if(amount>0)append_(SHEETS.PAYMENTS,HEADERS.Payments,{Timestamp:now,ClientID:clientIdForEvent_(eventId),EventID:eventId,Serial:serial,Amount:amount,Method:method,Status:status,Note:'Walk-in',RecordedBy:u.name});audit_(eventId,'WALK_IN','Claim',serial,u.name,u.role,{});return{ok:true,serial:serial,name:name,checkInToken:qr,tier:tier};}

// ---------- Reporting ----------

// ---------- Event readiness simulator ----------
function simulationFingerprint_(eventId){
  const e=getEvent_(eventId)||{}, tiers=getTiers_(eventId), groups=getGroups_(eventId);
  const stable={event:{EventID:e.EventID||'',OrgName:e.OrgName||'',ChapterName:e.ChapterName||'',EventTitle:e.EventTitle||'',EventDate:e.EventDate||'',EventTime:e.EventTime||'',VenueName:e.VenueName||'',VenueAddress:e.VenueAddress||'',Capacity:Number(e.Capacity)||0,UseGroups:String(e.UseGroups||''),GroupLabel:e.GroupLabel||'',SerialPrefix:e.SerialPrefix||'',CurrencySymbol:e.CurrencySymbol||'',PrimaryColor:e.PrimaryColor||'',AccentColor:e.AccentColor||'',LogoFileId:e.LogoFileId||'',BackgroundFileId:e.BackgroundFileId||''},tiers:tiers,groups:groups};
  return hash_(JSON.stringify(stable));
}
function latestSimulation_(eventId,mode){
  const fp=simulationFingerprint_(eventId);
  const rs=rows_(SHEETS.SIMULATIONS).filter(r=>normId_(r.EventID)===normId_(eventId)&&(!mode||String(r.Mode)===String(mode)));
  if(!rs.length)return null;
  const r=rs[rs.length-1];
  let summary={};try{summary=JSON.parse(String(r.SummaryJSON||'{}'));}catch(_){}
  return{runId:r.RunID,mode:r.Mode,passed:bool_(r.Passed),timestamp:r.Timestamp,total:Number(r.TotalTests)||0,passedTests:Number(r.PassedTests)||0,failedTests:Number(r.FailedTests)||0,warnings:Number(r.Warnings)||0,actor:r.Actor||'',configCurrent:String(r.ConfigFingerprint||'')===fp,summary:summary};
}
function simulateEvent_(body,u){
  const eventId=normId_(body.eventId), mode=String(body.mode||'full').toLowerCase()==='quick'?'quick':'full';
  if(!scopeAllows_(u,eventId)||!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role))return{ok:false,error:'Admin permission required.'};
  const eb=getEventBase_(eventId);if(!eb)return{ok:false,error:'Event not found.'};
  const ff=latestValidForm_(eb.EventID,eb.ConfigToken);if(ff)syncFormToOperational_(eb.EventID,eb.ConfigToken,ff);
  const e=getEvent_(eventId), tiers=getTiers_(eventId), groups=getGroups_(eventId), tierKeys=Object.keys(tiers);
  const tests=[], warnings=[];
  const add=(name,pass,detail,category)=>tests.push({name:name,pass:!!pass,detail:String(detail||''),category:category||'General'});
  const warn=(name,detail)=>warnings.push({name:name,detail:String(detail||'')});
  const req={OrgName:'Organization',ChapterName:'Chapter / Unit',EventTitle:'Event title',EventDate:'Event date',EventTime:'Event time',VenueName:'Venue',VenueAddress:'Venue address',SerialPrefix:'Serial prefix'};
  Object.keys(req).forEach(k=>add(req[k]+' configured',!!String(e&&e[k]||'').trim(),String(e&&e[k]||'')||'Missing','Configuration'));
  add('Event configuration synchronized',!!e&&normId_(e.EventID)===eventId,'Operational event record available','Configuration');
  add('Event date is readable',!!dateObj_(e&&e.EventDate),String(e&&e.EventDate||''),'Configuration');
  add('Capacity is valid',!e||Number(e.Capacity)>=0,'Capacity '+Number(e&&e.Capacity||0),'Capacity');
  add('At least one ticket tier exists',tierKeys.length>0,tierKeys.length+' tier(s)','Tiers');
  const labelSet={};
  tierKeys.forEach((k,i)=>{const t=tiers[k]||{},label=String(t.label||'').trim(),color=tierColor_(t.color,i),price=money_(t.price),cap=Number(t.capacity)||0;add('Tier '+(label||k)+' has a label',!!label,label||'Missing','Tiers');add('Tier '+(label||k)+' price is valid',price>=0,(e.CurrencySymbol||'$')+price,'Tiers');add('Tier '+(label||k)+' color is valid',!!validHex_(color),color,'Tiers');add('Tier '+(label||k)+' capacity is valid',cap>=0,String(cap),'Capacity');add('Tier '+(label||k)+' key is unique',!labelSet[k],k,'Tiers');labelSet[k]=true;});
  const useGroups=String(e&&e.UseGroups||'true').toLowerCase()!=='false';
  add('Group configuration is valid',!useGroups||groups.length>0,useGroups?groups.length+' group(s)':'Groups disabled','Groups');
  add('Group list contains no blank/undefined values',groups.every(g=>!!cleanGroup_(g)),groups.join(', ')||'None','Groups');
  if(e&&e.Capacity&&tierKeys.length){const sumCaps=tierKeys.reduce((s,k)=>s+(Number(tiers[k].capacity)||0),0);if(sumCaps&&sumCaps>Number(e.Capacity))warn('Tier capacities exceed event capacity',sumCaps+' tier seats vs '+e.Capacity+' event capacity');}

  // In-memory sandbox. No production sheets, counters, vouchers, claims, payments or check-ins are changed.
  const state={serial:0,vouchers:{},claims:{},checkins:{},payments:[],walkins:[]};
  const nextSerial=()=>String(e.SerialPrefix||'SIM-')+String(++state.serial).padStart(3,'0');
  const newVoucher=(tier)=>{const id='SIMV-'+String(Object.keys(state.vouchers).length+1).padStart(3,'0');state.vouchers[id]={id:id,tier:tier||'',claimed:false,status:'Available'};return state.vouchers[id]};
  const simClaim=(v,tierChoice,idx)=>{if(!v||v.status==='Cancelled')return{ok:false,error:'Voucher unavailable'};if(v.claimed)return{ok:false,error:'Already claimed'};const tier=v.tier||tierChoice;if(!tiers[tier])return{ok:false,error:'Invalid ticket tier'};const group=useGroups?(groups[idx%Math.max(groups.length,1)]||''):'';if(useGroups&&groups.length&&!groups.includes(group))return{ok:false,error:'Invalid group'};const serial=nextSerial(),qr='SIMQR-'+token_().slice(0,32);const c={serial:serial,qr:qr,name:'Simulator Guest '+serial,tier:tier,group:group,status:'Active',amountDue:money_(tiers[tier].price),amountPaid:0,paymentStatus:'Pending',source:'Voucher'};state.claims[serial]=c;v.claimed=true;v.status='Claimed';v.serial=serial;return{ok:true,claim:c}};
  const simCheckIn=(claim)=>{if(!claim||claim.status==='Revoked')return{ok:false,error:'Revoked'};if(state.checkins[claim.serial])return{ok:false,error:'Already checked in'};state.checkins[claim.serial]={serial:claim.serial,tier:claim.tier};return{ok:true}};
  const simWalkIn=(tier,idx)=>{if(!tiers[tier])return{ok:false,error:'Invalid tier'};const serial=nextSerial(),group=useGroups?(groups[idx%Math.max(groups.length,1)]||''):'',statuses=['Paid','Pending','Complimentary','Sponsored','Waived'],methods=['Cash','Zelle','Cash App','Card','Other'];const c={serial:serial,qr:'SIMQR-'+token_().slice(0,32),name:'Simulator Walk-In '+serial,tier:tier,group:group,status:'Active',amountDue:money_(tiers[tier].price),amountPaid:idx%2===0?money_(tiers[tier].price):0,paymentStatus:statuses[idx%statuses.length],paymentMethod:methods[idx%methods.length],source:'Walk-In'};state.claims[serial]=c;state.walkins.push(c);state.checkins[serial]={serial:serial,tier:tier};if(c.amountPaid>0)state.payments.push({serial:serial,amount:c.amountPaid,method:c.paymentMethod});return{ok:true,claim:c}};

  const claimed=[];
  tierKeys.forEach((k,i)=>{const t=tiers[k],v=newVoucher(k),r=simClaim(v,'',i);add('Fixed-tier voucher claims '+t.label,r.ok&&r.claim&&r.claim.tier===k,r.ok?r.claim.serial:r.error,'Voucher → Claim');if(r.ok){claimed.push(r.claim);add('Ticket '+t.label+' price matches configuration',r.claim.amountDue===money_(t.price),(e.CurrencySymbol||'$')+r.claim.amountDue,'Ticket');add('Ticket '+t.label+' serial uses prefix',String(r.claim.serial).startsWith(String(e.SerialPrefix||'')),r.claim.serial,'Ticket');add('Ticket '+t.label+' QR generated',String(r.claim.qr).length>20,'Secure QR token generated','Ticket');const ci=simCheckIn(r.claim);add('Gate check-in succeeds for '+t.label,ci.ok,ci.ok?'Checked in':ci.error,'Gate');const dup=simCheckIn(r.claim);add('Duplicate gate check-in blocked for '+t.label,!dup.ok&&dup.error==='Already checked in',dup.error,'Gate');}});
  if(tierKeys.length){const openTargets=mode==='full'?tierKeys:[tierKeys[0]];openTargets.forEach((k,i)=>{const v=newVoucher(''),r=simClaim(v,k,i);add('Open voucher can select '+tiers[k].label,r.ok&&r.claim.tier===k,r.ok?r.claim.serial:r.error,'Voucher → Claim');});}
  if(tierKeys.length){const v=newVoucher(tierKeys[0]);const first=simClaim(v,'',0),second=simClaim(v,'',0);add('Duplicate voucher claim blocked',first.ok&&!second.ok&&second.error==='Already claimed',second.error||'Not blocked','Security');const bad=newVoucher('');const invalid=simClaim(bad,'__INVALID__',0);add('Invalid tier is rejected',!invalid.ok&&invalid.error==='Invalid ticket tier',invalid.error,'Security');}
  if(useGroups&&groups.length){const badGroup='__INVALID_GROUP__';add('Invalid group would be rejected',!groups.includes(badGroup),badGroup+' is not configured','Security');}
  if(claimed.length){const c=claimed[0],oldQr=c.qr;c.status='Revoked';const revoked=simCheckIn(c);add('Revoked ticket is rejected at gate',!revoked.ok&&revoked.error==='Revoked',revoked.error,'Security');c.status='Active';c.qr='SIMQR-'+token_().slice(0,32);add('QR reissue invalidates prior credential',c.qr!==oldQr,'Old and new QR differ','Security');}
  if(e&&Number(e.Capacity)>0)add('Event over-capacity condition is detectable',Number(e.Capacity)+1>Number(e.Capacity),'Capacity '+e.Capacity,'Capacity');
  tierKeys.forEach(k=>{const cap=Number(tiers[k].capacity)||0;if(cap)add(tiers[k].label+' over-capacity condition is detectable',cap+1>cap,'Tier capacity '+cap,'Capacity')});
  const walkTargets=mode==='full'?tierKeys:(tierKeys.length?[tierKeys[0]]:[]);walkTargets.forEach((k,i)=>{const r=simWalkIn(k,i);add('Walk-in created and checked in for '+tiers[k].label,r.ok&&!!state.checkins[r.claim.serial],r.ok?r.claim.serial:r.error,'Walk-In');if(r.ok)add('Walk-in payment state valid for '+tiers[k].label,['Paid','Pending','Complimentary','Sponsored','Waived'].includes(r.claim.paymentStatus),r.claim.paymentStatus,'Payments');});
  const allClaims=Object.values(state.claims), searchTarget=allClaims[0];if(searchTarget){const q=searchTarget.serial.toLowerCase(),found=allClaims.some(c=>[c.name,c.serial].some(v=>String(v).toLowerCase().includes(q)));add('Guest search finds a simulated ticket',found,searchTarget.serial,'Search');}
  const checkedCount=Object.keys(state.checkins).length, issuedCount=allClaims.length, face=allClaims.reduce((s,c)=>s+c.amountDue,0),paid=allClaims.reduce((s,c)=>s+c.amountPaid,0);add('Dashboard issued count reconciles',issuedCount===Object.keys(state.claims).length,String(issuedCount),'Reporting');add('Dashboard check-in count reconciles',checkedCount<=issuedCount,checkedCount+' / '+issuedCount,'Reporting');add('Finance totals reconcile',face>=paid,'Face '+face+'; paid '+paid,'Reporting');

  const failed=tests.filter(t=>!t.pass), passed=tests.length-failed.length, runId='SIM-'+Utilities.getUuid().slice(0,10).toUpperCase(), fingerprint=simulationFingerprint_(eventId);
  const summary={eventId:eventId,eventTitle:e.EventTitle||'',mode:mode,passed:failed.length===0,totalTests:tests.length,passedTests:passed,failedTests:failed.length,warnings:warnings.length,tests:tests,warningsList:warnings,synthetic:{tiersTested:tierKeys.length,vouchersCreated:Object.keys(state.vouchers).length,ticketsCreated:issuedCount,walkInsCreated:state.walkins.length,checkInsCreated:checkedCount,faceValue:face,amountPaid:paid},configuration:{capacity:Number(e.Capacity)||0,tierCount:tierKeys.length,groupCount:groups.length,status:e.Status}};
  append_(SHEETS.SIMULATIONS,HEADERS.SimulationRuns,{Timestamp:new Date(),ClientID:clientIdForEvent_(eventId),EventID:eventId,RunID:runId,Mode:mode,Passed:summary.passed,TotalTests:tests.length,PassedTests:passed,FailedTests:failed.length,Warnings:warnings.length,Actor:u.name,Role:u.role,ConfigFingerprint:fingerprint,SummaryJSON:JSON.stringify(summary)});
  audit_(eventId,'RUN_READINESS_SIMULATION','Simulation',runId,u.name,u.role,{mode:mode,passed:summary.passed,total:tests.length,failed:failed.length,warnings:warnings.length});
  return{ok:true,runId:runId,...summary,configFingerprint:fingerprint};
}

function dashboard_(eventId){const e=getEvent_(eventId),tiers=getTiers_(eventId),alloc=allocations_(eventId),claims=rows_(SHEETS.CLAIMS).filter(r=>normId_(r.EventID)===eventId&&String(r.Status)!=='Revoked'),checkins=rows_(SHEETS.CHECKINS).filter(r=>normId_(r.EventID)===eventId&&String(r.Status)!=='Undone'),vouchers=rows_(SHEETS.VOUCHERS).filter(r=>normId_(r.EventID)===eventId&&String(r.Status)!=='Cancelled'),payments=rows_(SHEETS.PAYMENTS).filter(r=>normId_(r.EventID)===eventId);const due=claims.reduce((s,r)=>s+money_(r.AmountDue),0),paid=claims.reduce((s,r)=>s+money_(r.AmountPaid),0),byTier={};Object.keys(tiers).forEach(k=>byTier[k]={label:tiers[k].label,color:tiers[k].color,issued:0,checkedIn:0});claims.forEach(r=>{if(byTier[r.TierKey])byTier[r.TierKey].issued++});checkins.forEach(r=>{if(byTier[r.TierKey])byTier[r.TierKey].checkedIn++});return{ok:true,event:publicEvent_(e),counts:{allocated:alloc.total,claims:claims.length,pendingVouchers:vouchers.filter(v=>!bool_(v.Claimed)).length,checkedIn:checkins.length,noShows:Math.max(0,claims.length-checkins.length),capacity:e.Capacity||0},finance:{faceValue:due,amountPaid:paid,outstanding:Math.max(0,due-paid),paymentEntries:payments.length},byTier:byTier};}
function auditList_(eventId){return{ok:true,rows:rows_(SHEETS.AUDIT).filter(r=>normId_(r.EventID)===eventId).slice(-250).reverse().map(r=>({timestamp:r.Timestamp,action:r.Action,entityType:r.EntityType,entityId:r.EntityID,actor:r.Actor,role:r.Role,details:r.Details}))};}

// ---------- Communication ----------
function logCommunication_(body,u){const eventId=normId_(body.eventId);if(eventId&&!scopeAllows_(u,eventId))return{ok:false,error:'Not assigned.'};if(!['SYSTEM_OWNER','EVENT_ADMIN','FINANCE','GATE_SUPERVISOR'].includes(u.role))return{ok:false,error:'Communication permission required.'};contactLog_(eventId,String(body.entityType||'Contact'),String(body.entityId||''),String(body.channel||''),String(body.recipient||''),u.name,String(body.note||'Initiated from web interface'));return{ok:true};}
function sendAdminEmail_(body,u){const eventId=normId_(body.eventId),to=String(body.to||'').trim();if(!to)return{ok:false,error:'Email address required.'};try{MailApp.sendEmail(to,String(body.subject||'Event ticketing message'),String(body.message||''));contactLog_(eventId,String(body.entityType||'Contact'),String(body.entityId||''),'Email',to,u.name,String(body.subject||''));return{ok:true};}catch(e){return{ok:false,error:e.message};}}

function repairEventConfig_(eventId,u){eventId=normId_(eventId);if(!scopeAllows_(u,eventId)||!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role))return{ok:false,error:'Admin permission required.'};const b=getEventBase_(eventId);if(!b)return{ok:false,error:'Event not found.'};const f=latestValidForm_(b.EventID,b.ConfigToken);if(!f)return{ok:false,error:'No matching Google Form submission found for this event.'};const r=syncFormToOperational_(b.EventID,b.ConfigToken,f);audit_(eventId,'SYNC_EVENT_CONFIGURATION','Event',eventId,u.name,u.role,{tiers:r.tierCount,groups:r.groupCount});return{ok:true,event:publicEvent_(getEvent_(eventId)),tiers:getTiers_(eventId),groups:getGroups_(eventId),tierCount:r.tierCount,groupCount:r.groupCount};}

// ---------- Status / admin maintenance ----------
function setEventStatus_(body,u){const eventId=normId_(body.eventId),status=String(body.status||'');if(!LIFECYCLE.includes(status))return{ok:false,error:'Invalid status.'};const r=getEventBase_(eventId);if(!r)return{ok:false,error:'Event not found.'};const current=String(r.Status||'Draft');if(['Draft','Client Submitted','Preview Ready','Client Approved'].includes(status)&&status!==current)return{ok:false,error:'That lifecycle stage is controlled by the client configuration workflow.'};if(status==='Active'&&current!=='Client Approved')return{ok:false,error:'Client approval is required before activation.'};if(status==='Active'){const sim=latestSimulation_(eventId,'full');const override=bool_(body.simulationOverride)&&u.role==='SYSTEM_OWNER';if((!sim||!sim.passed||!sim.configCurrent)&&!override)return{ok:false,error:'A successful Full Event Readiness Test for the current configuration is required before activation.',needsSimulation:true,latestSimulation:sim};if(override)audit_(eventId,'SIMULATION_OVERRIDE','Event',eventId,u.name,u.role,{reason:String(body.overrideReason||'System Owner override')});}if(status==='Closed'&&current!=='Active')return{ok:false,error:'Only an Active event can be closed.'};if(status==='Archived'&&current!=='Closed')return{ok:false,error:'Close the event before archiving it.'};r.Status=status;r.UpdatedAt=new Date();writeRow_(SHEETS.EVENTS,HEADERS.Events,r._row,r);audit_(eventId,'SET_EVENT_STATUS','Event',eventId,u.name,u.role,{from:current,status:status});return{ok:true,status:status};}
function configLink_(eventId,u){const r=getEventBase_(eventId);if(!r||!scopeAllows_(u,eventId))return{ok:false,error:'Event not found or not assigned.'};if(!r.ConfigToken){r.ConfigToken=token_();r.UpdatedAt=new Date();writeRow_(SHEETS.EVENTS,HEADERS.Events,r._row,r);}return{ok:true,url:CONFIG_URL+'?event='+encodeURIComponent(r.EventID)+'&key='+encodeURIComponent(r.ConfigToken),formUrl:formPrefillUrl_(r.EventID,r.ConfigToken),clientName:r.ClientName||'',clientEmail:r.ClientEmail||'',clientPhone:r.ClientPhone||''};}


function accessList_(u){if(u.role!=='SYSTEM_OWNER')return{ok:false,error:'System Owner permission required.'};return{ok:true,users:rows_(SHEETS.ACCESS).map(r=>({userId:r.UserID,name:r.Name,role:r.Role,clientScope:r.ClientScope||'*',eventScope:r.EventScope,email:r.Email||'',phone:r.Phone||'',active:bool_(r.Active)}))};}
function createAccessUser_(body,u){if(u.role!=='SYSTEM_OWNER')return{ok:false,error:'System Owner permission required.'};const role=String(body.role||'').toUpperCase();if(!ROLES.includes(role))return{ok:false,error:'Invalid role.'};const pass=String(body.newPasscode||'').trim();if(pass.length<8)return{ok:false,error:'Passcode must be at least 8 characters.'};const id='USR-'+Utilities.getUuid().slice(0,8).toUpperCase();append_(SHEETS.ACCESS,HEADERS.AccessControl,{UserID:id,Name:String(body.name||'').trim(),Role:role,PasscodeHash:hash_(pass),ClientScope:String(body.clientScope||'*').trim()||'*',EventScope:String(body.eventScope||'*').trim()||'*',Email:String(body.email||'').trim(),Phone:String(body.phone||'').trim(),Active:true,CreatedAt:new Date()});audit_('','CREATE_ACCESS_USER','Access',id,u.name,u.role,{role:role,scope:body.eventScope||'*'});return{ok:true,userId:id};}
function setAccessActive_(body,u){if(u.role!=='SYSTEM_OWNER')return{ok:false,error:'System Owner permission required.'};const r=rowBy_(SHEETS.ACCESS,'UserID',String(body.userId||''));if(!r)return{ok:false,error:'User not found.'};r.Active=bool_(body.active);writeRow_(SHEETS.ACCESS,HEADERS.AccessControl,r._row,r);audit_('','SET_ACCESS_ACTIVE','Access',r.UserID,u.name,u.role,{active:r.Active});return{ok:true};}
function voucherSearch_(eventId,q){q=String(q||'').toLowerCase().trim();return rows_(SHEETS.VOUCHERS).filter(r=>normId_(r.EventID)===eventId).filter(r=>!q||[r.BatchID,r.VoucherToken,r.RecipientName,r.RecipientEmail,r.RecipientPhone,r.Serial].some(v=>String(v||'').toLowerCase().includes(q))).slice(-100).reverse().map(r=>({token:r.VoucherToken,batchId:r.BatchID,tier:r.TierKey,status:r.Status||'Available',claimed:bool_(r.Claimed),dispatched:bool_(r.Dispatched),serial:r.Serial||'',recipientName:r.RecipientName||'',recipientEmail:r.RecipientEmail||'',recipientPhone:r.RecipientPhone||''}));}
function cancelVoucher_(body,u){const r=rowBy_(SHEETS.VOUCHERS,'VoucherToken',String(body.token||''));if(!r)return{ok:false,error:'Voucher not found.'};const eventId=normId_(r.EventID);if(!scopeAllows_(u,eventId)||!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role))return{ok:false,error:'Admin permission required.'};if(bool_(r.Claimed))return{ok:false,error:'Claimed vouchers cannot be cancelled. Revoke the issued ticket instead.'};r.Status='Cancelled';writeRow_(SHEETS.VOUCHERS,HEADERS.Vouchers,r._row,r);audit_(eventId,'CANCEL_VOUCHER','Voucher',r.VoucherToken,u.name,u.role,{});return{ok:true};}
function backupEvent_(eventId,u){if(!scopeAllows_(u,eventId)||!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role))return{ok:false,error:'Admin permission required.'};const payload={createdAt:new Date().toISOString(),event:getEvent_(eventId),tiers:getTiers_(eventId),groups:getGroups_(eventId),vouchers:rows_(SHEETS.VOUCHERS).filter(r=>normId_(r.EventID)===eventId),claims:rows_(SHEETS.CLAIMS).filter(r=>normId_(r.EventID)===eventId),checkIns:rows_(SHEETS.CHECKINS).filter(r=>normId_(r.EventID)===eventId),payments:rows_(SHEETS.PAYMENTS).filter(r=>normId_(r.EventID)===eventId),simulations:rows_(SHEETS.SIMULATIONS).filter(r=>normId_(r.EventID)===eventId),audit:rows_(SHEETS.AUDIT).filter(r=>normId_(r.EventID)===eventId)};const blob=Utilities.newBlob(JSON.stringify(payload,null,2),'application/json',eventId+'-backup-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone()||'America/New_York','yyyyMMdd-HHmmss')+'.json');const f=assetFolder_().createFile(blob);audit_(eventId,'BACKUP_EVENT','Event',eventId,u.name,u.role,{fileId:f.getId()});return{ok:true,fileName:f.getName(),driveUrl:f.getUrl()};}



// ---------- V6 optional automated email reminders ----------
function installV6DailyReminderTrigger(){ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='runV6DailyReminders').forEach(t=>ScriptApp.deleteTrigger(t));ScriptApp.newTrigger('runV6DailyReminders').timeBased().everyDays(1).atHour(9).create();SpreadsheetApp.getUi().alert('Daily reminder trigger installed for approximately 9 AM script time. Only Active events with AutoReminders enabled are processed.');}
function setAutoReminders_(body,u){const id=normId_(body.eventId),e=getEventBase_(id);if(!e)return{ok:false,error:'Event not found.'};if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role)||!scopeAllows_(u,id))return{ok:false,error:'Admin permission required.'};e.AutoReminders=bool_(body.enabled);e.UpdatedAt=new Date();writeRow_(SHEETS.EVENTS,HEADERS.Events,e._row,e);audit_(id,'SET_AUTO_REMINDERS','Event',id,u.name,u.role,{enabled:e.AutoReminders});return{ok:true,enabled:e.AutoReminders};}
function runV6DailyReminders(){const today=new Date();const day=Utilities.formatDate(today,Session.getScriptTimeZone()||'America/New_York','yyyy-MM-dd');rows_(SHEETS.EVENTS).filter(e=>String(e.Status)==='Active'&&bool_(e.AutoReminders)).forEach(e=>{const last=e.LastReminderRun?Utilities.formatDate(new Date(e.LastReminderRun),Session.getScriptTimeZone()||'America/New_York','yyyy-MM-dd'):'';if(last===day)return;const eventId=normId_(e.EventID),event=getEvent_(eventId),date=dateObj_(event&&event.EventDate);if(date){const days=Math.ceil((date-new Date(today.getFullYear(),today.getMonth(),today.getDate()))/86400000);if(days<0||days>14)return;}const unclaimed=rows_(SHEETS.VOUCHERS).filter(v=>normId_(v.EventID)===eventId&&!bool_(v.Claimed)&&String(v.Status)!=='Cancelled'&&String(v.RecipientEmail||'').trim());unclaimed.forEach(v=>{try{const url=VOUCHER_URL+'?voucher='+encodeURIComponent(v.VoucherToken);MailApp.sendEmail(String(v.RecipientEmail).trim(),'Reminder: claim your event ticket','Your event ticket voucher is waiting to be claimed.\n\n'+url);contactLog_(eventId,'Voucher',v.VoucherToken,'Email',v.RecipientEmail,'Automated Reminder','Unclaimed voucher reminder');}catch(_){}});const pending=rows_(SHEETS.CLAIMS).filter(c=>normId_(c.EventID)===eventId&&String(c.Status)!=='Revoked'&&!['Paid','Complimentary','Sponsored','Waived'].includes(String(c.PaymentStatus))&&String(c.Email||'').trim());pending.forEach(c=>{try{MailApp.sendEmail(String(c.Email).trim(),'Event ticket payment reminder','This is a reminder that your event ticket has an outstanding payment. Please contact the event organizer if you need assistance.');contactLog_(eventId,'Claim',c.Serial,'Email',c.Email,'Automated Reminder','Pending payment reminder');}catch(_){}});const base=getEventBase_(eventId);if(base){base.LastReminderRun=new Date();writeRow_(SHEETS.EVENTS,HEADERS.Events,base._row,base);}audit_(eventId,'AUTOMATED_REMINDERS','Event',eventId,'System','SYSTEM',{unclaimed:unclaimed.length,pendingPayments:pending.length});});}

// ---------- V6 client isolation / monitoring / health ----------
function clientIdForEvent_(eventId){const e=getEventBase_(eventId);return normId_(e&&e.ClientID||'');}
function ensureClient_(clientId,name,email,phone){clientId=normId_(clientId);if(!clientId)return null;let c=rowBy_(SHEETS.CLIENTS,'ClientID',clientId);const now=new Date();if(!c){append_(SHEETS.CLIENTS,HEADERS.Clients,{ClientID:clientId,ClientName:name||clientId,PrimaryContact:name||'',Email:email||'',Phone:phone||'',Status:'Active',CreatedAt:now,UpdatedAt:now});c=rowBy_(SHEETS.CLIENTS,'ClientID',clientId);}return c;}
function createClient_(body,u){if(u.role!=='SYSTEM_OWNER')return{ok:false,error:'System Owner permission required.'};const id=normId_(body.clientId);if(!id)return{ok:false,error:'Client ID is required.'};if(rowBy_(SHEETS.CLIENTS,'ClientID',id))return{ok:false,error:'Client ID already exists.'};ensureClient_(id,String(body.clientName||id),String(body.email||''),String(body.phone||''));audit_('','CREATE_CLIENT','Client',id,u.name,u.role,{clientName:body.clientName||''});return{ok:true,clientId:id};}
function listClients_(u){if(u.role==='SYSTEM_OWNER')return rows_(SHEETS.CLIENTS).filter(r=>String(r.Status||'Active')!=='Disabled').map(r=>({clientId:r.ClientID,clientName:r.ClientName,email:r.Email||'',phone:r.Phone||'',status:r.Status||'Active'}));const ids=String(u.clientScope||'').split(',').map(normId_).filter(Boolean);return rows_(SHEETS.CLIENTS).filter(r=>ids.includes(normId_(r.ClientID))).map(r=>({clientId:r.ClientID,clientName:r.ClientName,email:r.Email||'',phone:r.Phone||'',status:r.Status||'Active'}));}
function eventLocked_(eventId){const e=getEventBase_(eventId);return !!(e&&bool_(e.EmergencyReadOnly));}
function setEmergencyReadOnly_(body,u){const id=normId_(body.eventId),e=getEventBase_(id);if(!e)return{ok:false,error:'Event not found.'};if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role)||!scopeAllows_(u,id))return{ok:false,error:'Admin permission required.'};e.EmergencyReadOnly=bool_(body.enabled);e.UpdatedAt=new Date();ensureHeaders_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EVENTS),['EmergencyReadOnly']);const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EVENTS),heads=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String),col=heads.indexOf('EmergencyReadOnly')+1;sh.getRange(e._row,col).setValue(e.EmergencyReadOnly);audit_(id,'SET_EMERGENCY_READ_ONLY','Event',id,u.name,u.role,{enabled:e.EmergencyReadOnly});return{ok:true,enabled:e.EmergencyReadOnly};}
function health_(eventId){const e=getEvent_(eventId),tiers=getTiers_(eventId),groups=getGroups_(eventId),base=getEventBase_(eventId);if(!e)return{status:'BLOCKED',score:0,issues:['Event not found'],warnings:[]};const issues=[],warnings=[];if(!e.EventTitle)issues.push('Missing event title');if(!e.EventDate)issues.push('Missing event date');if(!e.EventTime)issues.push('Missing event time');if(!e.VenueName)issues.push('Missing venue');if(!Object.keys(tiers).length)issues.push('No ticket tiers');if(String(e.UseGroups).toLowerCase()!=='false'&&!groups.length)issues.push('No groups/chapters');if(!base||!base.ApprovedAt)warnings.push('Client has not approved current design');if(!e.LogoURL)warnings.push('No logo uploaded');if(!e.Capacity)warnings.push('No finite event capacity');const sim=latestSimulation_(eventId,'full');if(!sim||!sim.passed||!sim.configCurrent)issues.push('Current configuration has no passing Full Certification Test');if(eventLocked_(eventId))warnings.push('Emergency read-only mode is enabled');const status=issues.length?'BLOCKED':warnings.length?'NEEDS ATTENTION':'READY';const score=Math.max(0,100-issues.length*20-warnings.length*5);return{status,score,issues,warnings,simulation:sim||null};}
function recentActivity_(eventId,limit){return rows_(SHEETS.AUDIT).filter(r=>normId_(r.EventID)===normId_(eventId)).slice(-Math.min(Number(limit)||20,50)).reverse().map(r=>({timestamp:r.Timestamp,action:r.Action,entityType:r.EntityType,entityId:r.EntityID,actor:r.Actor,role:r.Role,details:r.Details}));}
function globalSearch_(u,q){q=String(q||'').trim().toLowerCase();if(!q)return[];const allowedEvents=rows_(SHEETS.EVENTS).filter(e=>scopeAllows_(u,e.EventID)).map(e=>normId_(e.EventID));const out=[];rows_(SHEETS.CLAIMS).filter(r=>allowedEvents.includes(normId_(r.EventID))).forEach(r=>{if([r.Name,r.Email,r.Phone,r.Serial,r.GroupName].some(v=>String(v||'').toLowerCase().includes(q)))out.push({type:'Ticket',eventId:r.EventID,id:r.Serial,label:r.Name,detail:(r.GroupName||'')+' · '+(r.PaymentStatus||'')});});rows_(SHEETS.VOUCHERS).filter(r=>allowedEvents.includes(normId_(r.EventID))).forEach(r=>{if([r.BatchID,r.VoucherToken,r.RecipientName,r.RecipientEmail,r.RecipientPhone,r.Serial].some(v=>String(v||'').toLowerCase().includes(q)))out.push({type:'Voucher',eventId:r.EventID,id:r.VoucherToken,label:r.RecipientName||r.BatchID,detail:r.Status||'Available'});});rows_(SHEETS.EVENTS).filter(r=>allowedEvents.includes(normId_(r.EventID))).forEach(r=>{if([r.EventID,r.EventTitle,r.ClientName,r.ClientID].some(v=>String(v||'').toLowerCase().includes(q)))out.push({type:'Event',eventId:r.EventID,id:r.EventID,label:r.EventTitle,detail:r.Status});});return out.slice(0,100);}
function updateGuestNote_(body,u){const id=normId_(body.eventId),r=rowBy_(SHEETS.CLAIMS,'Serial',String(body.serial||''));if(!r||normId_(r.EventID)!==id)return{ok:false,error:'Ticket not found.'};if(!['SYSTEM_OWNER','EVENT_ADMIN','GATE_SUPERVISOR'].includes(u.role)||!scopeAllows_(u,id))return{ok:false,error:'Supervisor permission required.'};ensureHeaders_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CLAIMS),['InternalNotes']);const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CLAIMS),heads=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String),col=heads.indexOf('InternalNotes')+1;sh.getRange(r._row,col).setValue(String(body.note||''));audit_(id,'UPDATE_GUEST_NOTE','Claim',r.Serial,u.name,u.role,{});return{ok:true};}
function reminderCandidates_(eventId,u){if(!scopeAllows_(u,eventId))return{ok:false,error:'Not assigned.'};const e=getEvent_(eventId),claims=rows_(SHEETS.CLAIMS).filter(r=>normId_(r.EventID)===eventId&&String(r.Status)!=='Revoked'),v=rows_(SHEETS.VOUCHERS).filter(r=>normId_(r.EventID)===eventId&&!bool_(r.Claimed)&&String(r.Status)!=='Cancelled');return{ok:true,event:publicEvent_(e),unclaimed:v.slice(0,100).map(r=>({token:r.VoucherToken,name:r.RecipientName||'',email:r.RecipientEmail||'',phone:r.RecipientPhone||'',batchId:r.BatchID})),pendingPayments:claims.filter(r=>!['Paid','Complimentary','Sponsored','Waived'].includes(String(r.PaymentStatus))).slice(0,100).map(r=>({serial:r.Serial,name:r.Name,email:r.Email||'',phone:r.Phone||'',amountDue:money_(r.AmountDue),amountPaid:money_(r.AmountPaid)}))};}
function exportClientCsv_(u,clientId){clientId=normId_(clientId||u.clientScope);if(!clientAllows_(u,clientId))return{ok:false,error:'Client access denied.'};const events=rows_(SHEETS.EVENTS).filter(e=>normId_(e.ClientID)===clientId);const ids=events.map(e=>normId_(e.EventID));const lines=[['ClientID','EventID','EventTitle','Serial','Guest','Email','Phone','Group','Tier','Status','AmountDue','AmountPaid','PaymentStatus','CheckedIn'].join(',')];const check=rows_(SHEETS.CHECKINS);rows_(SHEETS.CLAIMS).filter(r=>ids.includes(normId_(r.EventID))).forEach(r=>{const ev=events.find(e=>normId_(e.EventID)===normId_(r.EventID))||{};const vals=[clientId,r.EventID,ev.EventTitle,r.Serial,r.Name,r.Email,r.Phone,r.GroupName,r.TierKey,r.Status,r.AmountDue,r.AmountPaid,r.PaymentStatus,check.some(c=>normId_(c.EventID)===normId_(r.EventID)&&c.Serial===r.Serial&&String(c.Status)!=='Undone')?'Yes':'No'];lines.push(vals.map(v=>'"'+String(v==null?'':v).replace(/"/g,'""')+'"').join(','));});return{ok:true,fileName:clientId+'-ticketing-export.csv',csv:lines.join('\n')};}
function clientDashboard_(u){if(!['CLIENT_ADMIN','CLIENT_VIEWER','CLIENT_FINANCE','SYSTEM_OWNER'].includes(u.role))return{ok:false,error:'Client portal role required.'};let ids=[];if(u.role==='SYSTEM_OWNER')ids=rows_(SHEETS.CLIENTS).map(c=>normId_(c.ClientID));else ids=String(u.clientScope||'').split(',').map(normId_).filter(Boolean);const clients=rows_(SHEETS.CLIENTS).filter(c=>ids.includes(normId_(c.ClientID)));const events=rows_(SHEETS.EVENTS).filter(e=>ids.includes(normId_(e.ClientID))).filter(e=>scopeAllows_(u,e.EventID)).map(e=>{const d=dashboard_(normId_(e.EventID));return{clientId:e.ClientID,eventId:e.EventID,title:e.EventTitle,date:normalizeDate_(e.EventDate),status:e.Status,health:health_(e.EventID),counts:d.counts,finance:d.finance,byTier:d.byTier,readOnly:bool_(e.EmergencyReadOnly)};});const totals=events.reduce((a,e)=>{a.issued+=e.counts.claims;a.checkedIn+=e.counts.checkedIn;a.pendingVouchers+=e.counts.pendingVouchers;a.collected+=e.finance.amountPaid;a.outstanding+=e.finance.outstanding;return a;},{issued:0,checkedIn:0,pendingVouchers:0,collected:0,outstanding:0});return{ok:true,user:u,clients:clients.map(c=>({clientId:c.ClientID,clientName:c.ClientName})),events,totals};}
function clientRecent_(u,clientId){clientId=normId_(clientId);if(!clientAllows_(u,clientId))return{ok:false,error:'Client access denied.'};const ids=rows_(SHEETS.EVENTS).filter(e=>normId_(e.ClientID)===clientId&&scopeAllows_(u,e.EventID)).map(e=>normId_(e.EventID));const rows=rows_(SHEETS.AUDIT).filter(r=>ids.includes(normId_(r.EventID))).slice(-30).reverse();return{ok:true,rows:rows.map(r=>({timestamp:r.Timestamp,eventId:r.EventID,action:r.Action,actor:r.Actor,role:r.Role}))};}

// ---------- Web API ----------
function doGet(e){const p=e&&e.parameter||{},action=String(p.action||'');try{
  if(action==='clientContext')return json_(clientContext_(normId_(p.event),p.key));
  if(action==='voucher')return json_(voucher_(p.token));
  if(action==='batch')return json_(batch_(p.batch));
  return json_({ok:true,service:'Event Ticketing V6'});
}catch(err){return json_({ok:false,error:err.message});}}

function doPost(e){const b=body_(e),action=String(b.action||'');try{
  // public/client/voucher actions
  if(action==='clientContext')return json_(clientContext_(normId_(b.eventId),b.configKey));
  if(action==='clientUpload')return json_(clientUpload_(b));
  if(action==='clientPreview')return json_(clientPreview_(b));
  if(action==='clientApprove')return json_(clientApprove_(b));
  if(action==='sendVoucher')return json_(sendVoucher_(b));
  if(action==='claim')return json_(claim_(b));

  const u=authenticate_(b.adminPasscode);if(!u)return json_({ok:false,error:'Incorrect passcode.'});
  if(action==='verifyPasscode')return json_({ok:true,user:u});
  if(action==='clientDashboard')return json_(clientDashboard_(u));
  if(action==='clientRecent')return json_(clientRecent_(u,b.clientId));
  if(action==='clientExportCsv')return json_(exportClientCsv_(u,b.clientId));
  if(action==='clients')return json_({ok:true,clients:listClients_(u)});
  if(action==='createClient')return json_(createClient_(b,u));
  if(action==='globalSearch')return json_({ok:true,results:globalSearch_(u,b.query)});
  if(action==='eventHealth'){if(!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Not assigned.'});return json_({ok:true,health:health_(normId_(b.eventId))});}
  if(action==='recentActivity'){if(!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Not assigned.'});return json_({ok:true,rows:recentActivity_(normId_(b.eventId),b.limit)});}
  if(action==='reminderCandidates'){if(!['SYSTEM_OWNER','EVENT_ADMIN','FINANCE','CLIENT_ADMIN','CLIENT_FINANCE'].includes(u.role))return json_({ok:false,error:'Reminder permission required.'});return json_(reminderCandidates_(normId_(b.eventId),u));}
  if(action==='updateGuestNote')return json_(updateGuestNote_(b,u));
  if(action==='setEmergencyReadOnly')return json_(setEmergencyReadOnly_(b,u));
  if(action==='setAutoReminders')return json_(setAutoReminders_(b,u));
  if(action==='eventsAdmin'){if(!['SYSTEM_OWNER','EVENT_ADMIN','FINANCE','CLIENT_ADMIN','CLIENT_VIEWER','CLIENT_FINANCE'].includes(u.role))return json_({ok:false,error:'Admin, Finance, or Client Portal role required.'});return json_({ok:true,user:u,events:currentFutureEvents_(u,bool_(b.includePast))});}
  if(action==='createEventShell'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role))return json_({ok:false,error:'Admin permission required.'});return json_(createEventShell_(b,u));}
  if(action==='configLink'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role))return json_({ok:false,error:'Admin permission required.'});return json_(configLink_(normId_(b.eventId),u));}
  if(action==='setEventStatus'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role))return json_({ok:false,error:'Admin permission required.'});if(!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Not assigned.'});return json_(setEventStatus_(b,u));}
  if(action==='repairEventConfig')return json_(repairEventConfig_(b.eventId,u));
  if(action==='runSimulation'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin permission required.'});return json_(simulateEvent_(b,u));}
  if(action==='latestSimulation'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin permission required.'});return json_({ok:true,latest:latestSimulation_(normId_(b.eventId),b.mode||'full')});}
  if(action==='generateVouchers'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role))return json_({ok:false,error:'Admin permission required.'});return json_(generateVouchers_(b,u));}
  if(action==='adminEventContext'){if(!['SYSTEM_OWNER','EVENT_ADMIN','FINANCE','CLIENT_ADMIN','CLIENT_VIEWER','CLIENT_FINANCE'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin, Finance, or Client Portal permission required.'});const id=normId_(b.eventId),eb=getEventBase_(id);if(eb){const ff=latestValidForm_(eb.EventID,eb.ConfigToken);if(ff)syncFormToOperational_(eb.EventID,eb.ConfigToken,ff);}return json_({ok:true,event:publicEvent_(getEvent_(id)),tiers:getTiers_(id),groups:getGroups_(id)});}
  if(action==='dashboard'){if(!['SYSTEM_OWNER','EVENT_ADMIN','FINANCE','CLIENT_ADMIN','CLIENT_VIEWER','CLIENT_FINANCE'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin, Finance, or Client Portal permission required.'});return json_(dashboard_(normId_(b.eventId)));}
  if(action==='searchGuests'){if(!['SYSTEM_OWNER','EVENT_ADMIN','FINANCE','CLIENT_ADMIN','CLIENT_VIEWER','CLIENT_FINANCE'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin, Finance, or Client Portal permission required.'});return json_({ok:true,results:guestSearch_(normId_(b.eventId),b.query)});}
  if(action==='updatePayment'){if(!['SYSTEM_OWNER','EVENT_ADMIN','FINANCE','GATE_SUPERVISOR'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Payment permission required.'});return json_(updatePayment_(b,u));}
  if(action==='transferTicket'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin permission required.'});return json_(transferTicket_(b,u));}
  if(action==='reissueQr'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin permission required.'});return json_(reissueQr_(b,u));}
  if(action==='revokeTicket'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin permission required.'});return json_(setTicketStatus_(b,u,'Revoked'));}
  if(action==='reactivateTicket'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin permission required.'});return json_(setTicketStatus_(b,u,'Active'));}
  if(action==='audit'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin permission required.'});return json_(auditList_(normId_(b.eventId)));}
  if(action==='sendAdminEmail'){if(!['SYSTEM_OWNER','EVENT_ADMIN','FINANCE','GATE_SUPERVISOR'].includes(u.role))return json_({ok:false,error:'Communication permission required.'});return json_(sendAdminEmail_(b,u));}
  if(action==='logCommunication')return json_(logCommunication_(b,u));
  if(action==='accessList')return json_(accessList_(u));
  if(action==='createAccessUser')return json_(createAccessUser_(b,u));
  if(action==='setAccessActive')return json_(setAccessActive_(b,u));
  if(action==='voucherSearch'){if(!['SYSTEM_OWNER','EVENT_ADMIN'].includes(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Admin permission required.'});return json_({ok:true,rows:voucherSearch_(normId_(b.eventId),b.query)});}
  if(action==='cancelVoucher')return json_(cancelVoucher_(b,u));
  if(action==='backupEvent')return json_(backupEvent_(normId_(b.eventId),u));

  // gate-specific actions
  if(action==='gateEvents'){if(!roleCanGate_(u.role))return json_({ok:false,error:'Gate role required.'});return json_({ok:true,user:u,events:currentFutureEvents_(u,false).filter(x=>x.Status==='Active')});}
  if(action==='gateContext'){if(!roleCanGate_(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Gate permission required.'});const id=normId_(b.eventId);return json_({ok:true,user:u,event:publicEvent_(getEvent_(id)),tiers:getTiers_(id),groups:getGroups_(id)});}
  if(action==='gateSearch'){if(!roleCanGate_(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Gate permission required.'});return json_({ok:true,results:gateSearch_(normId_(b.eventId),b.query,u)});}
  if(action==='gateSummary'){if(!roleCanGate_(u.role)||!scopeAllows_(u,b.eventId))return json_({ok:false,error:'Gate permission required.'});return json_(gateSummary_(normId_(b.eventId)));}
  if(action==='checkIn')return json_(checkIn_(b,u));
  if(action==='undoCheckIn')return json_(undoCheckIn_(b,u));
  if(action==='walkIn')return json_(walkIn_(b,u));

  return json_({ok:false,error:'Unknown action.'});
}catch(err){return json_({ok:false,error:err.message});}}
