// Championship branding + first-launch setup
const FEDERATION_NAME = "FÉDÉRATION TUNISIENNE DE TAEKWONDO";
const RIGHTS_TEXT = "ALL RIGHTS RESERVED • FÉDÉRATION TUNISIENNE DE TAEKWONDO";
function federationLogoSVG(){return `<img class="fedLogo officialFedLogo" src="federation_logo_transparent.png" alt="Fédération Tunisienne de Taekwondo — logo officiel" referrerpolicy="no-referrer">`}
function landingScreen(role="operator"){
 const isPublic=role==="public";
 return `<div class="landing ${isPublic?"landingPublic":"landingOperator"}">
   <div class="landingGlow glowA"></div><div class="landingGlow glowB"></div>
   <div class="landingArt leftArt" aria-hidden="true"><svg viewBox="0 0 260 300"><circle cx="130" cy="40" r="18"/><path d="M130 63 L130 150 M130 88 L72 120 M130 88 L190 112 M130 150 L76 224 M130 150 L177 194 L236 151"/><path d="M72 120 L40 91 M190 112 L222 80"/></svg></div>
   <div class="landingArt rightArt" aria-hidden="true"><svg viewBox="0 0 260 300"><circle cx="130" cy="40" r="18"/><path d="M130 63 L130 150 M130 88 L72 120 M130 88 L190 112 M130 150 L76 224 M130 150 L177 194 L236 151"/><path d="M72 120 L40 91 M190 112 L222 80"/></svg></div>
   <div class="landingCard">
     <div class="logoRing">${federationLogoSVG()}</div>
     <div class="eyebrow">WORLD TAEKWONDO • POOMSAE</div>
     <h1>${esc((STATE.setupDraft?.competitionName)||STATE.competitionName||"POOMSAE CHAMPIONSHIP")}</h1>
     <h2>${FEDERATION_NAME}</h2>
     <div class="landingRule"></div>
     <div class="landingSubtitle">OFFICIAL POOMSAE SCORING SYSTEM</div>
     <div class="landingStatus">${isPublic?"PUBLIC DISPLAY":"OPERATOR CONSOLE"}</div>
     ${isPublic?(()=>{const d=STATE.setupDraft||{};const has=d.category||d.athleteCount||d.competitionName;return `<div class="waitingText">${has?"COMPETITION SETUP IN PROGRESS…":"WAITING FOR OPERATOR CONFIGURATION…"}</div>${has?`<div class="publicSetupLive"><div><b>${esc(d.classCode||"—")}</b> · ${esc(d.category||"—")} · ${esc(d.gender||"—")}</div><div>${d.athleteCount||0} ${d.participation==="DUO"?"DUOS":d.participation==="TEAM"?"TEAMS":"ATHLETES"} · ${d.judgeCount||3} JUDGES · TAPIS ${d.matNumber||1}</div><div>${(Number(d.athleteCount)>(d.participation==="INDIVIDUAL"?8:4))?`DEMI-FINALE: ${esc(WT_DATA.poomsae[d.poomsaeSemi1]||"—")} · ${esc(WT_DATA.poomsae[d.poomsaeSemi2]||"—")}<br>`:""}FINALE: ${esc(WT_DATA.poomsae[d.poomsaeFinal1]||"—")} · ${esc(WT_DATA.poomsae[d.poomsaeFinal2]||"—")}</div>${(d.athletes||[]).length?`<div class="liveAthleteList">${d.athletes.map((a,i)=>`<span>${i+1}. ${esc(a.name||"—")}</span>`).join("")}</div>`:""}</div>`:""}`})():`<div class="landingButtons"><button class="btn goldBtn launchBtn" onclick="openCompetitionSetup()">COMMENCER LA COMPÉTITION</button><button class="btn primary launchBtn" onclick="openPublicDisplay()">OPEN PUBLIC DISPLAY</button></div>`}
     <div class="rights">${RIGHTS_TEXT}</div>
   </div>
 </div>`
}
function broadcastSetupDraft(){
 const get=id=>document.getElementById(id);
 if(!get("setupClass"))return;
 STATE.setupDraft={competitionName:get("setupCompetition")?.value.trim()||"POOMSAE CHAMPIONSHIP",classCode:get("setupClass")?.value||"",category:get("setupCategory")?.value||"",gender:get("setupGender")?.value||"",participation:get("setupParticipation")?.value||"INDIVIDUAL",athleteCount:Number(get("setupAthletes")?.value)||0,judgeCount:Number(get("setupJudges")?.value)||3,matNumber:Number(get("setupMat")?.value)||1,poomsaeSemi1:get("setupSemiPoomsae1")?.value||"",poomsaeSemi2:get("setupSemiPoomsae2")?.value||"",poomsaeFinal1:get("setupFinalPoomsae1")?.value||"",poomsaeFinal2:get("setupFinalPoomsae2")?.value||"",athletes:[...document.querySelectorAll("#setupAthleteFields .setupAthleteItem")].map(x=>{const part=get("setupParticipation")?.value||"INDIVIDUAL";const e=buildEntryFromRow(x,part);return {name:e.name||"",club:e.club||"",members:e.members||null}})};
 saveState(); broadcast();
}
function participationMemberCount(participation){return participation==="DUO"?2:participation==="TEAM"?3:1;}
function entryGenderLabel(entry){
  if(entry?.members?.length){
    const gs=entry.members.map(m=>m.gender).filter(Boolean);
    if(STATE.participation==="DUO") return "Mixed";
    if(gs.length&&gs.every(g=>g===gs[0])) return gs[0];
  }
  return entry?.gender||STATE.gender||"—";
}
function entryDisplayName(entry){
  if(entry?.members?.length) return entry.members.map(m=>m.name).filter(Boolean).join(" / ");
  return entry?.name||"";
}
function buildEntryFromRow(row, participation){
  const memberCount=participationMemberCount(participation);
  if(participation==="INDIVIDUAL"){
    return {name:row.querySelector(".athName")?.value.trim()||"",club:row.querySelector(".athClub")?.value.trim()||"",gender:STATE.gender||""};
  }
  const members=Array.from({length:memberCount},(_,i)=>({name:row.querySelector(`.memberName${i}`)?.value.trim()||"",gender:row.querySelector(`.memberGender${i}`)?.value||""}));
  return {members,club:row.querySelector(".athClub")?.value.trim()||"",name:members.map(m=>m.name).filter(Boolean).join(" / "),gender:participation==="DUO"?"Mixed":(members[0]?.gender||"")};
}

function normalizeExcelHeader(v){return String(v??"").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^a-z0-9]+/g,"");}
function importAthletesFromExcel(file){
 if(!file)return;
 if(typeof XLSX==="undefined")return alert("Le module Excel n'est pas disponible. Vérifiez la connexion Internet puis rechargez la page.");
 const reader=new FileReader();
 reader.onload=e=>{
   try{
     const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});
     const ws=wb.Sheets[wb.SheetNames[0]];
     const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:false});
     if(!rows.length)throw new Error("Fichier vide");
     const first=rows[0].map(normalizeExcelHeader);
     const nameKeys=["nom","nomprenom","nometprenom","athlete","athletenom","name","fullname","participant","participantnom"];
     const clubKeys=["club","clubnom","association","team","equipe","structure"];
     let nameIdx=first.findIndex(x=>nameKeys.includes(x));
     let clubIdx=first.findIndex(x=>clubKeys.includes(x));
     let start=1;
     if(nameIdx<0){
       nameIdx=0; clubIdx=rows[0].length>1?1:-1; start=0;
     }
     const imported=[];
     for(let i=start;i<rows.length;i++){
       const r=rows[i]||[];
       const name=String(r[nameIdx]??"").trim();
       const club=clubIdx>=0?String(r[clubIdx]??"").trim():"";
       if(name)imported.push({name,club});
     }
     if(!imported.length)throw new Error("Aucun athlète trouvé. Utilisez des colonnes Nom et Club.");
     if(imported.length>50)throw new Error("Le fichier contient plus de 50 athlètes. La limite est de 50.");
     const part=document.getElementById("setupParticipation");
     if(part&&part.value!=="INDIVIDUAL"){
       part.value="INDIVIDUAL";
       part.dispatchEvent(new Event("change",{bubbles:true}));
     }
     const count=document.getElementById("setupAthletes");
     if(count){count.value=String(imported.length);count.dispatchEvent(new Event("change",{bubbles:true}));}
     const fields=document.getElementById("setupAthleteFields");
     if(!fields)throw new Error("La liste des athlètes n'est pas disponible.");
     const items=[...fields.querySelectorAll(".setupAthleteItem")];
     imported.forEach((a,i)=>{
       const row=items[i]; if(!row)return;
       const n=row.querySelector(".athName"),c=row.querySelector(".athClub");
       if(n)n.value=a.name;
       if(c)c.value=a.club;
     });
     fields.dispatchEvent(new Event("input",{bubbles:true}));
     const status=document.getElementById("excelImportStatus");
     if(status){status.textContent=`${imported.length} ATHLÈTES IMPORTÉS`;status.classList.add("successImport");}
     if(typeof validateSetupAthletes==="function")validateSetupAthletes();
     if(typeof broadcastSetupDraft==="function")broadcastSetupDraft();
   }catch(err){alert(`Import Excel impossible : ${err.message||err}`);}
   const input=document.getElementById("setupExcelFile");if(input)input.value="";
 };
 reader.onerror=()=>alert("Impossible de lire le fichier Excel.");
 reader.readAsArrayBuffer(file);
}
function handleExcelFileInput(input){importAthletesFromExcel(input?.files?.[0]);}


const TUNISIA_GOVERNORATES_AR = [
 {id:"TN11",name:"تونس"},{id:"TN12",name:"أريانة"},{id:"TN13",name:"بن عروس"},{id:"TN14",name:"منوبة"},
 {id:"TN15",name:"نابل"},{id:"TN16",name:"زغوان"},{id:"TN17",name:"بنزرت"},
 {id:"TN21",name:"باجة"},{id:"TN22",name:"جندوبة"},{id:"TN23",name:"الكاف"},{id:"TN24",name:"سليانة"},
 {id:"TN31",name:"سوسة"},{id:"TN32",name:"المنستير"},{id:"TN33",name:"المهدية"},{id:"TN34",name:"صفاقس"},
 {id:"TN41",name:"القيروان"},{id:"TN42",name:"القصرين"},{id:"TN43",name:"سيدي بوزيد"},
 {id:"TN51",name:"قابس"},{id:"TN52",name:"مدنين"},{id:"TN53",name:"تطاوين"},
 {id:"TN61",name:"قفصة"},{id:"TN62",name:"توزر"},{id:"TN63",name:"قبلي"}
];
const COMPETITION_TYPES_AR = [
 {key:"regional",label:"البطولة الجهوية"},
 {key:"territorial",label:"البطولة الإقليمية"},
 {key:"national",label:"البطولة الوطنية"}
];
function generatedArabicCompetitionName(type, governorate){
 const t=COMPETITION_TYPES_AR.find(x=>x.key===type)?.label||"البطولة الجهوية";
 return type === "national" ? `${t} للبومزا` : `${t} للبومزا - ${governorate||"تونس"}`;
}
function openArabicCompetitionNameBuilder(){
 const old=document.getElementById("arabicCompetitionNameModal"); if(old){old.remove();return;}
 document.body.insertAdjacentHTML("beforeend",`<div id="arabicCompetitionNameModal" class="setupOverlay open" dir="rtl">
   <div class="arabicNameModal">
     <div class="arabicNameHeader">
       <div><div class="phaseTag">تسمية المسابقة</div><h2>إنشاء اسم المسابقة</h2><p>اختر نوع البطولة ثم مرّر المؤشر فوق الولاية المطلوبة واضغط عليها.</p></div>
       <button class="btn" onclick="closeArabicCompetitionNameBuilder()">إغلاق</button>
     </div>
     <div class="arabicNameTypes" role="group" aria-label="نوع البطولة">
       ${COMPETITION_TYPES_AR.map(t=>`<button type="button" class="arabicNameTypeBtn ${STATE.competitionType===t.key?"selected":""}" data-type="${t.key}" onclick="selectArabicCompetitionType('${t.key}')">${t.label}</button>`).join("")}
     </div>
     <div class="arabicMapLayout" id="arabicMapLayout">
       <div class="tunisiaMapCard" id="tunisiaMapCard">
         <div class="tunisiaMapTitle">خريطة الجمهورية التونسية</div>
         <div class="tunisiaMapHint">مرّر المؤشر فوق الولاية لمعرفة اسمها، ثم اضغط لاختيارها.</div>
         <div class="govHoverTooltip" id="govHoverTooltip">حرّك المؤشر فوق ولاية</div>
         <div class="tunisiaMapWrap" id="tunisiaGovernorateMap">
           <div id="tunisiaRealMap" class="tunisiaRealMap" aria-label="الخريطة التفاعلية لولايات تونس"></div>
         </div>
       </div>
       <div class="arabicNamePreviewCard">
         <div class="phaseTag">المعاينة</div>
         <div class="arabicNamePreviewLabel">اسم المسابقة النهائي</div>
         <div id="arabicCompetitionNamePreview" class="arabicNamePreview">${esc(generatedArabicCompetitionName(STATE.competitionType,STATE.competitionGovernorate))}</div>
         <div class="arabicSelectedInfo"><span>نوع البطولة</span><strong id="arabicSelectedType">${esc(COMPETITION_TYPES_AR.find(x=>x.key===STATE.competitionType)?.label||"—")}</strong><span>الولاية</span><strong id="arabicSelectedGovernorate">${esc(STATE.competitionGovernorate||"—")}</strong></div>
         <button class="btn goldBtn arabicNameApply" onclick="applyArabicCompetitionName()">استخدام هذا الاسم</button>
       </div>
     </div>
   </div>
 </div>`);
 updateArabicCompetitionBuilderUI();
 loadRealTunisiaGovernorateMap();
}

async function loadRealTunisiaGovernorateMap(){
 const host=document.getElementById("tunisiaRealMap");
 if(!host)return;
 if(host.dataset.loaded==="1")return;
 try{
   // Use the governorate topology from the Tunisia administrative map project.
   // The previous riateLab GeoJSON contained repeated governorate IDs/geometries,
   // which could place labels in the wrong governorate.
   const res=await fetch("https://raw.githubusercontent.com/mtimet/tnacmaps/master/topojson/tunisia.json",{cache:"force-cache"});
   if(!res.ok)throw new Error("map fetch failed");
   const topo=await res.json();
   renderRealTunisiaMap(host,topo);
   host.dataset.loaded="1";
 }catch(e){
   renderReferenceTunisiaMap(host);
 }
}

const TUNISIA_GOV_AR_BY_ID={
 TN11:"تونس",TN12:"أريانة",TN13:"بن عروس",TN14:"منوبة",TN15:"نابل",TN16:"زغوان",TN17:"بنزرت",TN21:"باجة",TN22:"جندوبة",TN23:"الكاف",TN24:"سليانة",TN31:"سوسة",TN32:"المنستير",TN33:"المهدية",TN34:"صفاقس",TN41:"القيروان",TN42:"القصرين",TN43:"سيدي بوزيد",TN51:"قابس",TN52:"مدنين",TN53:"تطاوين",TN61:"قفصة",TN62:"توزر",TN63:"قبلي"
};
const TUNISIA_GOV_AR_BY_FR={
 tunis:"تونس",tunisgovernorate:"تونس",ariana:"أريانة",benarous:"بن عروس",manouba:"منوبة",mannouba:"منوبة",nabeul:"نابل",zaghouan:"زغوان",bizerte:"بنزرت",beja:"باجة",jendouba:"جندوبة",jendouba:"جندوبة",kef:"الكاف",elkef:"الكاف",siliana:"سليانة",sousse:"سوسة",monastir:"المنستير",mahdia:"المهدية",sfax:"صفاقس",kairouan:"القيروان",kasserine:"القصرين",sidibouzid:"سيدي بوزيد",gabes:"قابس",medenine:"مدنين",tataouine:"تطاوين",gafsa:"قفصة",tozeur:"توزر",kebili:"قبلي"
};
function normalizeGovKey(v){return String(v||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().replace(/[^a-z0-9\u0600-\u06ff]+/g,"");}
function topoDecodeArcs(topo){
 const tr=topo.transform||{},sc=tr.scale||[1,1],tl=tr.translate||[0,0],cache=[];
 return (topo.arcs||[]).map((arc,idx)=>{
   if(cache[idx])return cache[idx];let x=0,y=0;const out=[];
   arc.forEach(pt=>{x+=pt[0];y+=pt[1];out.push([x*sc[0]+tl[0],y*sc[1]+tl[1]])});
   cache[idx]=out;return out;
 });
}
function topoArcRing(ids,arcs){
 const out=[];ids.forEach((id,k)=>{const rev=id<0,arc=arcs[rev?~id:id]||[],pts=rev?arc.slice().reverse():arc.slice();if(k&&pts.length)pts.shift();out.push(...pts)});return out;
}
function topoGeometryToCoordinates(g,arcs){
 if(!g)return null;
 if(g.type==="Polygon")return g.arcs.map(r=>topoArcRing(r,arcs));
 if(g.type==="MultiPolygon")return g.arcs.map(poly=>poly.map(r=>topoArcRing(r,arcs)));
 return null;
}
function topoObjectToFeatures(obj,arcs){
 const geoms=obj?.geometries||[];
 return geoms.map(g=>({properties:g.properties||{},geometry:{type:g.type,coordinates:topoGeometryToCoordinates(g,arcs)}})).filter(f=>f.geometry.coordinates);
}
function canonicalGovernorateName(props){
 const id=String(props.gov_id||props.gouv_id||props.governorate_id||props.govid||props.id||"").toUpperCase();
 if(TUNISIA_GOV_AR_BY_ID[id])return TUNISIA_GOV_AR_BY_ID[id];
 const candidates=[props.gov_name_f,props.gouv_fr,props.gov_name,props.name,props.name_fr,props.governorate,props.gov_name_a,props.gouv_ar,props.gov_ar];
 for(const c of candidates){const k=normalizeGovKey(c);if(TUNISIA_GOV_AR_BY_FR[k])return TUNISIA_GOV_AR_BY_FR[k];}
 return String(props.gov_name_a||props.gouv_ar||props.gov_ar||props.name||"").trim();
}
function ringAreaCentroid(ring){
 if(!ring?.length)return [0,0,0];let a=0,cx=0,cy=0;
 for(let i=0,j=ring.length-1;i<ring.length;j=i++){
   const [x0,y0]=ring[j],[x1,y1]=ring[i],cross=x0*y1-x1*y0;a+=cross;cx+=(x0+x1)*cross;cy+=(y0+y1)*cross;
 }
 a/=2;if(Math.abs(a)<1e-12){const p=ring[Math.floor(ring.length/2)]||[0,0];return [p[0],p[1],0];}
 return [cx/(6*a),cy/(6*a),Math.abs(a)];
}
function featureLabelPoint(geometry){
 const polys=geometry.type==="Polygon"?[geometry.coordinates]:geometry.coordinates;
 let best=null;
 (polys||[]).forEach(poly=>{const c=ringAreaCentroid(poly?.[0]||[]);if(!best||c[2]>best[2])best=c;});
 return best?[best[0],best[1]]:[0,0];
}
function renderRealTunisiaMap(host,topo){
 const obj=topo?.objects?.governorates||topo?.objects?.Governorates||Object.values(topo?.objects||{}).find(o=>/govern/i.test(String(o?.name||"")))||Object.values(topo?.objects||{})[0];
 if(!obj)throw new Error("Governorate topology not found");
 const arcs=topoDecodeArcs(topo),raw=topoObjectToFeatures(obj,arcs);
 const feats=raw.map(f=>({...f,govName:canonicalGovernorateName(f.properties)})).filter(f=>f.govName&&f.geometry);
 const uniq=new Map();feats.forEach(f=>{if(!uniq.has(f.govName))uniq.set(f.govName,f)});
 const finalFeats=[...uniq.values()];
 if(finalFeats.length<20)throw new Error("Incomplete governorate topology");
 const palette=["#E85D75","#F39C5A","#F5C84C","#48B7A8","#8ACB5A","#A875E8","#5C9FEA","#7E64D9","#49B86B","#E978A7","#4E86D8","#9A66C7","#3BAE98","#5C8FE8","#E78C3D","#F0D15A","#D96A9E","#5B8FE8","#77BF61","#E66363","#EBAA45","#38AFA4","#8D68D4","#4DA3D8"];
 let minX=Infinity,minY=Infinity,maxX=-Infinity,maxY=-Infinity;
 const visit=c=>{if(Array.isArray(c?.[0]))c.forEach(visit);else if(Array.isArray(c)){minX=Math.min(minX,c[0]);maxX=Math.max(maxX,c[0]);minY=Math.min(minY,c[1]);maxY=Math.max(maxY,c[1]);}};
 finalFeats.forEach(f=>visit(f.geometry.coordinates));
 const pad=18,W=760,H=900,sc=Math.min((W-2*pad)/(maxX-minX),(H-2*pad)/(maxY-minY));
 const px=x=>pad+(x-minX)*sc+(W-(maxX-minX)*sc-2*pad)/2;
 const py=y=>H-pad-(y-minY)*sc-(H-(maxY-minY)*sc-2*pad)/2;
 const pathD=ring=>ring.map((p,i)=>`${i?'L':'M'} ${px(p[0]).toFixed(2)} ${py(p[1]).toFixed(2)}`).join(" ")+" Z";
 const geomPath=geom=>geom.type==="Polygon"?geom.coordinates.map(pathD).join(" "):geom.type==="MultiPolygon"?geom.coordinates.flat().map(pathD).join(" "):"";
 const safe=v=>String(v).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/\"/g,"&quot;").replace(/'/g,"&#39;");
 const js=v=>String(v).replace(/\\/g,"\\\\").replace(/'/g,"\\'");
 const regions=finalFeats.map((f,i)=>{const name=f.govName,d=geomPath(f.geometry),c=featureLabelPoint(f.geometry),color=palette[i%palette.length];return `<g class="govMapRegion" data-gov="${safe(name)}" style="--gov-color:${color}" onmouseenter="previewArabicGovernorate('${js(name)}');showArabicGovernorateTooltip('${js(name)}')" onfocus="showArabicGovernorateTooltip('${js(name)}')" onclick="selectArabicGovernorate('${js(name)}')" tabindex="0" role="button" aria-label="${safe(name)}"><path d="${d}"/><text x="${px(c[0]).toFixed(1)}" y="${py(c[1]).toFixed(1)}">${safe(name)}</text></g>`}).join("");
 const legend=finalFeats.map((f,i)=>`<button type="button" class="govLegendItem" data-gov="${safe(f.govName)}" onclick="selectArabicGovernorate('${js(f.govName)}')"><span class="govLegendSwatch" style="--gov-color:${palette[i%palette.length]}"></span><span>${safe(f.govName)}</span></button>`).join("");
 host.innerHTML=`<div class="realMapProgramGrid"><div class="realMapSvgBox"><svg class="tunisiaActualMap realGeoMap" viewBox="0 0 ${W} ${H}" role="img" aria-label="خريطة ولايات تونس"><defs><filter id="tunisiaMapShadow" x="-20%" y="-20%" width="140%" height="140%"><feDropShadow dx="0" dy="8" stdDeviation="7" flood-opacity=".20"/></filter></defs><g filter="url(#tunisiaMapShadow)">${regions}</g></svg></div><div class="govLegend"><div class="govLegendTitle">ولايات تونس (24)</div><div class="govLegendGrid">${legend}</div></div></div><div class="mapSourceNote">خريطة تفاعلية مبرمجة بحدود الولايات التونسية، مع أسماء الولايات العربية الصحيحة.</div>`;
 updateArabicCompetitionBuilderUI();
}
function renderReferenceTunisiaMap(host){
 const pos={TN17:[58,3],TN21:[41,11],TN22:[31,14],TN23:[29,18],TN24:[46,20],TN11:[65,10],TN12:[66,8],TN14:[63,10],TN13:[67,11],TN15:[79,15],TN16:[64,16],TN31:[76,23],TN32:[81,24],TN33:[87,28],TN41:[63,25],TN42:[32,32],TN43:[48,34],TN34:[80,38],TN61:[31,42],TN51:[63,50],TN62:[15,49],TN63:[36,52],TN52:[73,57],TN53:[72,63]};
 host.innerHTML=`<div class="referenceMapFallback"><div class="referenceMapVisual"><img src="tunisia_governorates_reference.png" alt="خريطة ولايات تونس"><div class="referenceMapLabels">${TUNISIA_GOVERNORATES_AR.map(g=>{const p=pos[g.id]||[50,50];return `<button type="button" title="${g.name}" style="left:${p[0]}%;top:${p[1]}%" onclick="selectArabicGovernorate('${g.name.replace(/'/g,"\'")}')">${g.name}</button>`}).join("")}</div></div><div class="realMapFallbackHint">الخريطة de secours affiche les 24 gouvernorats. Vous pouvez aussi les choisir dans la liste ci-dessous.</div><div class="govFallbackGrid">${TUNISIA_GOVERNORATES_AR.map(g=>`<button type="button" onclick="selectArabicGovernorate('${g.name.replace(/'/g,"\'")}')">${g.name}</button>`).join("")}</div></div>`;
 updateArabicCompetitionBuilderUI();
}

function updateArabicCompetitionBuilderUI(){
 const national=STATE.competitionType==="national";
 const card=document.getElementById("tunisiaMapCard");
 if(card)card.classList.toggle("nationalLocked",national);
 const wrap=document.getElementById("tunisiaGovernorateMap");
 if(wrap)wrap.setAttribute("aria-hidden",national?"true":"false");
 document.querySelectorAll(".govMapRegion").forEach(b=>{b.style.pointerEvents=national?"none":"";b.style.opacity=national?".32":"";});
 const tooltip=document.getElementById("govHoverTooltip");
 if(tooltip)tooltip.textContent=national?"البطولة الوطنية — لا توجد ولاية":"حرّك المؤشر فوق ولاية";
}
function closeArabicCompetitionNameBuilder(){document.getElementById("arabicCompetitionNameModal")?.remove()}
function selectArabicCompetitionType(type){
 STATE.competitionType=type;
 if(type==="national") STATE.competitionGovernorate="";
 document.querySelectorAll(".arabicNameTypeBtn").forEach(b=>b.classList.toggle("selected",b.dataset.type===type));
 updateArabicCompetitionBuilderUI();
 updateArabicCompetitionNamePreview();
}
function showArabicGovernorateTooltip(name){const t=document.getElementById("govHoverTooltip");if(t)t.textContent=`الولاية: ${name}`;}
function previewArabicGovernorate(name){
 const p=document.getElementById("arabicCompetitionNamePreview"); if(p && STATE.competitionType)p.textContent=generatedArabicCompetitionName(STATE.competitionType,name);
}
function selectArabicGovernorate(name){
 STATE.competitionGovernorate=name;
 document.querySelectorAll(".govMapRegion,.govLegendItem").forEach(b=>b.classList.toggle("selected",b.dataset.gov===name));
 updateArabicCompetitionNamePreview();
}
function updateArabicCompetitionNamePreview(){
 const typeLabel=COMPETITION_TYPES_AR.find(x=>x.key===STATE.competitionType)?.label||"—";
 const national=STATE.competitionType==="national";
 const gov=national?"غير مطلوبة":(STATE.competitionGovernorate||"—");
 const p=document.getElementById("arabicCompetitionNamePreview"); if(p)p.textContent=STATE.competitionType?generatedArabicCompetitionName(STATE.competitionType,STATE.competitionGovernorate):"اختر نوع البطولة";
 const t=document.getElementById("arabicSelectedType");if(t)t.textContent=typeLabel;
 const g=document.getElementById("arabicSelectedGovernorate");if(g)g.textContent=gov;
 updateArabicCompetitionBuilderUI();
}
function applyArabicCompetitionName(){
 if(!STATE.competitionType)return alert("يرجى اختيار نوع البطولة.");
 if(STATE.competitionType!=="national" && !STATE.competitionGovernorate)return alert("يرجى اختيار الولاية من الخريطة.");
 const name=generatedArabicCompetitionName(STATE.competitionType,STATE.competitionGovernorate);
 const input=document.getElementById("setupCompetition"); if(input){input.value=name;input.dataset.generated="true";input.dispatchEvent(new Event("input",{bubbles:true}));}
 STATE.competitionName=name;STATE.competitionNameMode="generated";
 closeArabicCompetitionNameBuilder();
 if(typeof broadcastSetupDraft==="function")broadcastSetupDraft();
}

function openCompetitionSetup(){
 const modal=document.getElementById("setupModal"); if(modal){modal.classList.add("open");return}
 document.body.insertAdjacentHTML("beforeend",`<div id="setupModal" class="setupOverlay open"><div class="setupModal">
   <div class="setupTop"><div>${federationLogoSVG()}</div><div><div class="phaseTag">COMPETITION SETUP</div><h2>INITIALISATION DE LA CATÉGORIE</h2><p>Le système choisira automatiquement la phase.</p></div></div>
   <div class="cardMini excelImportCard">
     <div class="phaseTag">IMPORTATION DE LA LISTE DES ATHLÈTES</div>
     <div class="excelImportRow">
       <div><strong>Importer depuis Excel</strong><p class="muted">Colonnes recommandées : <b>Nom</b> et <b>Club</b>. Maximum 50 athlètes.</p></div>
       <label class="btn primary excelFileBtn">CHOISIR UN FICHIER EXCEL<input id="setupExcelFile" type="file" accept=".xlsx,.xls,.csv" onchange="handleExcelFileInput(this)"></label>
     </div>
     <div id="excelImportStatus" class="excelImportStatus">AUCUN FICHIER IMPORTÉ</div>
   </div>
   <div class="setupGrid">
     <label class="competitionNameField">Nom de la compétition
       <div class="competitionNameInputRow"><input id="setupCompetition" value="${esc(STATE.competitionName||"POOMSAE CHAMPIONSHIP")}" placeholder="Nom de la compétition"><button type="button" class="btn arabicNameGeneratorBtn" onclick="openArabicCompetitionNameBuilder()">🇹🇳 توليد اسم المسابقة</button></div>
       <small class="competitionNameHint">يمكنك كتابة الاسم يدويًا أو استعمال مولّد الاسم العربي.</small>
     </label>
     <label>Classe<select id="setupClass"><option value="A">A</option><option value="B">B</option></select></label>
     <label>Catégorie<select id="setupCategory"></select></label>
     <label>Sexe<select id="setupGender"><option value="Male">Male</option><option value="Female">Female</option></select></label>
     <label>Participation<select id="setupParticipation"><option value="INDIVIDUAL">Indiv</option><option value="DUO">Duo</option><option value="TEAM">Team</option></select></label>
     <label>Nombre d'athlètes<select id="setupAthletes">${Array.from({length:50},(_,i)=>`<option value="${i+1}">${i+1}</option>`).join("")}</select></label>
     <label>Nombre de juges<select id="setupJudges">${WT_DATA.judges.map(n=>`<option value="${n}">${n}</option>`).join("")}</select></label>
     <label>Numéro du tapis<select id="setupMat">${WT_DATA.mats.map(n=>`<option value="${n}">Tapis ${n}</option>`).join("")}</select></label>
   </div>
   <div class="poomsaeSetup cardMini">
     <div class="phaseTag">POOMSAE OFFICIELLES — FIXATION DE LA COMPÉTITION</div>
     <p class="muted">Après le nombre d’athlètes, si un demi-final est requis, choisissez les 2 Poomsae du demi-final et les 2 Poomsae de la finale. Elles seront verrouillées avant le début de la compétition.</p>
     <div id="semiPoomsaeSetup">
       <div class="phaseTag">DEMI-FINALE — 2 POOMSAE</div>
       <div class="setupGrid twoCols">
         <label>Round 1 — Demi-finale<select id="setupSemiPoomsae1"></select></label>
         <label>Round 2 — Demi-finale<select id="setupSemiPoomsae2"></select></label>
       </div>
     </div>
     <div class="phaseTag" style="margin-top:16px">FINALE — 2 POOMSAE</div>
     <div class="setupGrid twoCols">
       <label>Round 1 — Finale<select id="setupFinalPoomsae1"></select></label>
       <label>Round 2 — Finale<select id="setupFinalPoomsae2"></select></label>
     </div>
     <div class="lockNotice">🔒 POOMSAE FIXÉES — L'opérateur ne pourra plus les modifier pendant la compétition.</div>
   </div>
   <div id="athleteSetup" class="athleteSetup cardMini">
     <div class="phaseTag">ATHLETES — FIXATION DE LA LISTE</div>
     <p class="muted">Saisissez les noms avant le début de la compétition. La liste sera verrouillée après confirmation.</p>
     <div id="setupAthleteFields" class="setupAthleteFields"></div>
   </div>
   <div id="phasePreview" class="phasePreview"></div>
   <div class="setupActions"><button class="btn" onclick="closeCompetitionSetup()">ANNULER</button><button class="btn goldBtn" onclick="confirmCompetitionSetup()">CONFIRMER & COMMENCER</button></div>
 </div></div>`);
 const cls=document.getElementById("setupClass"),cat=document.getElementById("setupCategory"),count=document.getElementById("setupAthletes"),judges=document.getElementById("setupJudges"),pSemi1=document.getElementById("setupSemiPoomsae1"),pSemi2=document.getElementById("setupSemiPoomsae2"),pFinal1=document.getElementById("setupFinalPoomsae1"),pFinal2=document.getElementById("setupFinalPoomsae2"), athleteFields=document.getElementById("setupAthleteFields");
 cls.value=STATE.classCode||"B"; cat.innerHTML=WT_DATA.categories.map(x=>`<option>${esc(x)}</option>`).join(""); cat.value=STATE.category||"Cadet"; document.getElementById("setupGender").value=STATE.gender||""; document.getElementById("setupParticipation").value=STATE.participation||"INDIVIDUAL"; count.value=STATE.athleteCount||8; judges.value=STATE.judgeCount||3; document.getElementById("setupMat").value=STATE.matNumber||1;
 const refreshPoomsae=()=>{
   const allowed=allowedPoomsae(cls.value,cat.value);
   const opts=allowed.map(k=>`<option value="${k}">${esc(WT_DATA.poomsae[k])}</option>`).join("");
   [pSemi1,pSemi2,pFinal1,pFinal2].forEach(x=>x.innerHTML=opts);
   const fixed=STATE.fixedPoomsae||{};
   const semi=fixed.semi||{}; const fin=fixed.final||{};
   pSemi1.value=allowed.includes(semi.round1)?semi.round1:(fixed.round1&&allowed.includes(fixed.round1)?fixed.round1:(allowed[0]||""));
   pSemi2.value=allowed.includes(semi.round2)?semi.round2:(fixed.round2&&allowed.includes(fixed.round2)?fixed.round2:(allowed[1]||allowed[0]||""));
   pFinal1.value=allowed.includes(fin.round1)?fin.round1:(allowed[1]||allowed[0]||"");
   pFinal2.value=allowed.includes(fin.round2)?fin.round2:(allowed[2]||allowed[1]||allowed[0]||"");
   const semiWrap=document.getElementById("semiPoomsaeSetup"); const partNow=document.getElementById("setupParticipation")?.value||"INDIVIDUAL"; const thresholdNow=partNow==="INDIVIDUAL"?8:4; if(semiWrap)semiWrap.style.display=(Number(count.value)>thresholdNow?"block":"none");
 };
 const renderAthleteFields=()=>{
   const n=Number(count.value)||1;
   const participation=document.getElementById("setupParticipation")?.value||"INDIVIDUAL";
   const memberCount=participationMemberCount(participation);
   const old=[...athleteFields.querySelectorAll(".setupAthleteItem")].map(row=>buildEntryFromRow(row, participation));
   athleteFields.innerHTML=Array.from({length:n},(_,i)=>{
     const v=old[i]||{};
     if(participation==="INDIVIDUAL") return `<div class="setupAthleteItem"><span class="athNo">${i+1}</span><input class="athName" placeholder="Nom et prénom de l'athlète ${i+1}" value="${esc(v.name||"")}"><input class="athClub" placeholder="Club ${i+1}" value="${esc(v.club||"")}"></div>`;
     const members=Array.from({length:memberCount},(_,m)=>v.members?.[m]||{});
     const memberInputs=members.map((m,mi)=>`<div class="memberCell"><input class="memberName${mi}" placeholder="${participation==="DUO"?`Athlète ${mi+1}`:`Membre ${mi+1}`} — nom et prénom" value="${esc(m.name||"")}">${participation==="DUO"?`<select class="memberGender${mi}"><option value="">Sexe</option><option value="Male" ${m.gender==="Male"?"selected":""}>Male</option><option value="Female" ${m.gender==="Female"?"selected":""}>Female</option></select>`:""}</div>`).join("");
     return `<div class="setupAthleteItem groupEntry"><span class="athNo">${i+1}</span><div class="groupMembers">${memberInputs}</div><input class="athClub" placeholder="Club / Équipe ${i+1}" value="${esc(v.club||"")}"><small class="groupHint">${participation==="DUO"?"DUO : 1 homme + 1 femme · même club":"TEAM : 3 membres · même sexe · même club"}</small></div>`;
   }).join("");
 };
 const update=()=>{const n=Number(count.value);const part=document.getElementById("setupParticipation")?.value||"INDIVIDUAL";const threshold=part==="INDIVIDUAL"?8:4;const semiWrap=document.getElementById("semiPoomsaeSetup");if(semiWrap)semiWrap.style.display=n>threshold?"block":"none";const qualifier=part==="INDIVIDUAL"?8:4;const label=part==="DUO"?"DUOS":part==="TEAM"?"ÉQUIPES":"ATHLÈTES";document.getElementById("phasePreview").innerHTML=n<=threshold?`<span class="direct">DIRECT FINAL</span><small>${n} ${label} • ACCÈS DIRECT À LA FINALE</small>`:`<span class="semi">SEMI-FINAL</span><small>${n} ${label} • TOP ${qualifier} POUR LA FINALE</small>`;renderAthleteFields()};
 cls.addEventListener("change",()=>{refreshPoomsae();broadcastSetupDraft()});
 cat.addEventListener("change",()=>{refreshPoomsae();broadcastSetupDraft()});
 const participationEl=document.getElementById("setupParticipation");
 participationEl?.addEventListener("change",()=>{update();broadcastSetupDraft()});
 count.addEventListener("change",()=>{update();broadcastSetupDraft()});
 judges.addEventListener("change",broadcastSetupDraft);
 document.getElementById("setupCompetition")?.addEventListener("input",function(){if(!this.dataset.generated)STATE.competitionNameMode="manual";this.dataset.generated="";broadcastSetupDraft();});
 ["setupGender","setupParticipation","setupSemiPoomsae1","setupSemiPoomsae2","setupFinalPoomsae1","setupFinalPoomsae2","setupJudges","setupMat"].forEach(id=>document.getElementById(id)?.addEventListener("input",broadcastSetupDraft));
 ["setupGender","setupParticipation","setupSemiPoomsae1","setupSemiPoomsae2","setupFinalPoomsae1","setupFinalPoomsae2","setupJudges","setupMat"].forEach(id=>document.getElementById(id)?.addEventListener("change",broadcastSetupDraft));
 athleteFields.addEventListener("input",()=>{broadcastSetupDraft();validateSetupAthletes();});
 athleteFields.addEventListener("change",()=>{broadcastSetupDraft();validateSetupAthletes();});
 refreshPoomsae();update();renderJudgeNameFields();broadcastSetupDraft();validateSetupAthletes();
}
function closeCompetitionSetup(){document.getElementById("setupModal")?.remove()}
function validateSetupAthletes(showErrors=false){
 const rows=[...document.querySelectorAll("#setupAthleteFields .setupAthleteItem")];
 const participation=document.getElementById("setupParticipation")?.value||STATE.participation||"INDIVIDUAL";
 let ok=true; const names=[];
 rows.forEach(row=>{
   if(participation==="INDIVIDUAL"){
     const nameEl=row.querySelector(".athName"),clubEl=row.querySelector(".athClub");
     const name=nameEl?.value.trim()||"",club=clubEl?.value.trim()||"";
     nameEl?.classList.toggle("invalid-athlete",!name);clubEl?.classList.toggle("invalid-athlete",!club);if(!name||!club)ok=false;names.push(name.toLowerCase());return;
   }
   const clubEl=row.querySelector(".athClub");const club=clubEl?.value.trim()||"";clubEl?.classList.toggle("invalid-athlete",!club);if(!club)ok=false;
   const members=[...row.querySelectorAll(".groupMembers .memberCell")].map(cell=>({name:cell.querySelector("input")?.value.trim()||"",gender:cell.querySelector("select")?.value||"",nameEl:cell.querySelector("input"),genderEl:cell.querySelector("select")}));
   members.forEach(m=>{m.nameEl?.classList.toggle("invalid-athlete",!m.name);if(participation==="DUO")m.genderEl?.classList.toggle("invalid-athlete",!m.gender);if(!m.name||(participation==="DUO"&&!m.gender))ok=false;names.push(m.name.toLowerCase())});
   if(participation==="DUO" && members.length===2 && members[0].gender===members[1].gender){members.forEach(m=>m.genderEl?.classList.add("invalid-athlete"));ok=false;}
   if(participation==="TEAM" && !document.getElementById("setupGender")?.value){ok=false;}
 });
 const seen=new Map();names.filter(Boolean).forEach(n=>seen.set(n,(seen.get(n)||0)+1));
 rows.forEach(row=>row.querySelectorAll(".athName,.groupMembers input").forEach(el=>{const n=el.value.trim().toLowerCase();if(n&&seen.get(n)>1){el.classList.add("invalid-athlete");ok=false;}}));
 const btn=document.querySelector(".setupActions .goldBtn");if(btn){btn.disabled=!ok;btn.title=ok?"":(participation==="DUO"?"Chaque duo doit avoir 2 athlètes du même club : 1 homme + 1 femme.":participation==="TEAM"?"Chaque équipe doit avoir 3 athlètes du même club et du même sexe.":"Complétez les noms et clubs des athlètes.");}
 return ok;
}
function confirmCompetitionSetup(){
 const name=document.getElementById("setupCompetition")?.value.trim()||"POOMSAE CHAMPIONSHIP";
 const count=Number(document.getElementById("setupAthletes")?.value)||1;
 const previousMat=Number(STATE.matNumber)||1; STATE.competitionName=name;STATE.competitionNameMode=document.getElementById("setupCompetition")?.dataset.generated==="true"?"generated":(STATE.competitionNameMode||"manual");STATE.participation=document.getElementById("setupParticipation").value||"INDIVIDUAL";STATE.matNumber=Number(document.getElementById("setupMat").value)||1; if(typeof ensureStationId==="function" && (!STATE.stationId || previousMat!==STATE.matNumber)){STATE.stationId=stationNewId(STATE.matNumber);}STATE.classCode=document.getElementById("setupClass").value;STATE.category=document.getElementById("setupCategory").value;STATE.gender=document.getElementById("setupGender").value;STATE.athleteCount=count;STATE.judgeCount=Number(document.getElementById("setupJudges").value)||3;STATE.matNumber=Number(document.getElementById("setupMat").value)||1;STATE.judgeNames=STATE.judgeNames||{};for(let j=1;j<=STATE.judgeCount;j++){if(typeof STATE.judgeNames[j]!=="string")STATE.judgeNames[j]="";}const phaseThreshold=STATE.participation==="INDIVIDUAL"?8:4;STATE.phase=count<=phaseThreshold?"DIRECT FINAL":"SEMI-FINAL";
 STATE.fixedPoomsae={semi:{round1:document.getElementById("setupSemiPoomsae1").value,round2:document.getElementById("setupSemiPoomsae2").value},final:{round1:document.getElementById("setupFinalPoomsae1").value,round2:document.getElementById("setupFinalPoomsae2").value}};
 STATE.fixedPoomsae.round1=STATE.fixedPoomsae.final.round1; STATE.fixedPoomsae.round2=STATE.fixedPoomsae.final.round2;
 const semiThreshold=STATE.participation==="INDIVIDUAL"?8:4; if(count>semiThreshold && (!STATE.fixedPoomsae.semi.round1||!STATE.fixedPoomsae.semi.round2))return alert("Sélectionnez les 2 Poomsae du demi-final.");
 if(!STATE.fixedPoomsae.final.round1||!STATE.fixedPoomsae.final.round2)return alert("Sélectionnez les 2 Poomsae de la finale.");
 const rows=[...document.querySelectorAll("#setupAthleteFields .setupAthleteItem")];
 if(rows.length!==count)return alert("Le nombre d'athlètes ne correspond pas à la liste.");
 const names=rows.map(x=>x.querySelector(".athName")?.value.trim()||"");
 const clubs=rows.map(x=>x.querySelector(".athClub")?.value.trim()||"");
 if(!validateSetupAthletes(true))return false;
 // Keep existing QR tokens/Peer connection across category changes. Only create tokens for newly added judges.
 STATE.judgeQr=STATE.judgeQr||{peerId:"",tokens:{},deleted:{},sessionId:"",scores:{}};
 STATE.collectorQr=STATE.collectorQr||{peerId:"",token:"",deleted:false};
 if(typeof judgeConnections!=="undefined"){
   Object.keys(judgeConnections).forEach(k=>{if(Number(k)>STATE.judgeCount){try{judgeConnections[k]?.close()}catch(_){} delete judgeConnections[k];}});
 }
 for(let j=1;j<=STATE.judgeCount;j++){if(!STATE.judgeQr.tokens[j])STATE.judgeQr.tokens[j]=judgeRandomToken();STATE.judgeQr.deleted[j]=false;}
 Object.keys(STATE.judgeQr.tokens).forEach(k=>{if(Number(k)>STATE.judgeCount)delete STATE.judgeQr.tokens[k];});
 STATE.judgeQr.sessionId="";STATE.judgeQr.scores={};
 if(typeof judgePeer!=="undefined" && judgePeer?.open){STATE.judgeQr.peerId=judgePeer.id;STATE.collectorQr.peerId=judgePeer.id;}
 saveState();
 STATE.setupComplete=true;STATE.competitionStarted=false;const builtEntries=rows.map(row=>buildEntryFromRow(row,STATE.participation));if(STATE.participation==="DUO")STATE.gender="Mixed";if(STATE.participation==="TEAM")builtEntries.forEach(entry=>{entry.gender=STATE.gender;entry.members?.forEach(m=>{m.gender=STATE.gender;});}); STATE.athletes=builtEntries.map((entry)=>({id:crypto.randomUUID?.()||String(Date.now()+Math.random()),name:entry.name,club:entry.club||"—",gender:entry.gender||STATE.gender,participation:STATE.participation,members:entry.members||null,matNumber:STATE.matNumber,absent:false,rounds:[],finalScore:null}));STATE.currentIndex=0;STATE.currentRound=1;STATE.rankings=[];STATE.publishedResult=null;STATE.tieResolutions={};STATE.state=STATES.IDLE;STATE.publicMode="ready";STATE.timer={elapsed:0,running:false,expired:false};saveState();closeCompetitionSetup();broadcast();render();
 if(typeof ensureJudgePeer==="function"){ensureJudgePeer(()=>{try{judgeSendContext({scoringOpen:false});}catch(_){} });}
}
