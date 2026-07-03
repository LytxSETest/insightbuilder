(function(){
  "use strict";

  // Must match the "id" in the add-in configuration installed in MyGeotab.
  // Saved dashboards (AddInData) are keyed on this value.
  var ADDIN_ID = "aZdataInsightBuilder0001";

  var PALETTE = ["#009CDE","#1F1446","#12A594","#E08735","#C2453E","#6E56CF","#3E7BB6","#8A9099"];
  var TOPN = 12, TOPN_SERIES = 6, MAX_MEASURES = 4, MAX_DIMS = 2;
  var SVGNS = "http://www.w3.org/2000/svg";

  // Each catalog item carries a type signature. The auto-viz engine and the
  // resolver read only the signature, never the specific data point, so
  // standard and "non-standard" combinations run through identical code.
  var MEASURES = [
    { id:"harsh_braking", label:"Harsh braking", cat:"Safety", source:"exception", agg:"count", unit:"events", vtype:"int", match:["harsh brak","hard brak"] },
    { id:"harsh_accel",   label:"Harsh acceleration", cat:"Safety", source:"exception", agg:"count", unit:"events", vtype:"int", match:["harsh accel","hard accel"] },
    { id:"harsh_corner",  label:"Harsh cornering", cat:"Safety", source:"exception", agg:"count", unit:"events", vtype:"int", match:["corner"] },
    { id:"speeding",      label:"Speeding", cat:"Safety", source:"exception", agg:"count", unit:"events", vtype:"int", match:["speed"] },
    { id:"seatbelt",      label:"Seatbelt", cat:"Safety", source:"exception", agg:"count", unit:"events", vtype:"int", match:["seat belt","seatbelt"] },
    { id:"idle_events",   label:"Idling events", cat:"Safety", source:"exception", agg:"count", unit:"events", vtype:"int", match:["idl"] },
    { id:"total_events",  label:"All safety events", cat:"Safety", source:"exception", agg:"count", unit:"events", vtype:"int", matchAll:true },

    { id:"distance",      label:"Distance driven", cat:"Utilization", source:"trip", agg:"sum", field:"distance", unit:"km", vtype:"dist" },
    { id:"driving_time",  label:"Driving time", cat:"Utilization", source:"trip", agg:"sum", field:"drivingDuration", durField:true, unit:"h", vtype:"dur" },
    { id:"idle_time",     label:"Idle time", cat:"Utilization", source:"trip", agg:"sum", field:"idlingDuration", durField:true, unit:"h", vtype:"dur" },
    { id:"stop_time",     label:"Stopped time", cat:"Utilization", source:"trip", agg:"sum", field:"stopDuration", durField:true, unit:"h", vtype:"dur" },
    { id:"trips",         label:"Trips", cat:"Utilization", source:"trip", agg:"count", unit:"trips", vtype:"int" },
    { id:"avg_speed",     label:"Average speed", cat:"Utilization", source:"trip", agg:"avg", field:"averageSpeed", unit:"km/h", vtype:"speed" },
    { id:"top_speed",     label:"Top speed", cat:"Utilization", source:"trip", agg:"max", field:"maximumSpeed", unit:"km/h", vtype:"speed" },
    { id:"idle_pct",      label:"Idle %", cat:"Utilization", source:"trip", agg:"derived", unit:"%", vtype:"pct", components:["idlingDuration","drivingDuration"] },

    { id:"eng_rpm_avg",  label:"Engine RPM (avg)", cat:"Engine", source:"status", diag:"DiagnosticEngineSpeedId", agg:"avg", unit:"rpm", vtype:"rpm" },
    { id:"eng_rpm_peak", label:"Engine RPM (peak)", cat:"Engine", source:"status", diag:"DiagnosticEngineSpeedId", agg:"max", unit:"rpm", vtype:"rpm" },
    { id:"eng_coolant",  label:"Coolant temp (avg)", cat:"Engine", source:"status", diag:"DiagnosticEngineCoolantTemperatureId", agg:"avg", unit:"\u00B0C", vtype:"tempC" },
    { id:"eng_roadspeed_peak", label:"Engine road speed (peak)", cat:"Engine", source:"status", diag:"DiagnosticEngineRoadSpeedId", agg:"max", unit:"km/h", vtype:"speedKph" },
    { id:"eng_hours",    label:"Engine hours (latest)", cat:"Engine", source:"status", diag:"DiagnosticEngineHoursId", agg:"last", unit:"h", vtype:"hoursSec" },
    { id:"odometer",     label:"Odometer (latest)", cat:"Engine", source:"status", diag:"DiagnosticOdometerId", agg:"last", unit:"km", vtype:"distM" },
    { id:"odo_distance", label:"Distance driven (odometer)", cat:"Engine", source:"status", diag:"DiagnosticOdometerId", agg:"delta", unit:"km", vtype:"distM" },

    { id:"fuel_level",   label:"Fuel level (avg %)", cat:"Fuel & efficiency", source:"status", diag:"DiagnosticFuelLevelId", agg:"avg", unit:"%", vtype:"pct" },
    { id:"fuel_tank",    label:"Fuel in tank (avg)", cat:"Fuel & efficiency", source:"status", diag:"DiagnosticFuelUnitsId", agg:"avg", unit:"L", vtype:"volumeL" },
    { id:"fuel_used",    label:"Fuel used (period)", cat:"Fuel & efficiency", source:"status", diag:"DiagnosticTotalFuelUsedId", agg:"delta", unit:"L", vtype:"volumeL" },
    { id:"idle_fuel",    label:"Idle fuel used (period)", cat:"Fuel & efficiency", source:"status", diag:"DiagnosticDeviceTotalIdleFuelId", agg:"delta", unit:"L", vtype:"volumeL" },

    { id:"crank_v_min",  label:"Cranking voltage (lowest)", cat:"Battery & electrical", source:"status", diag:"DiagnosticCrankingVoltageId", agg:"min", unit:"V", vtype:"volts" },
    { id:"crank_v_avg",  label:"Cranking voltage (avg)", cat:"Battery & electrical", source:"status", diag:"DiagnosticCrankingVoltageId", agg:"avg", unit:"V", vtype:"volts" },
    { id:"dev_v_min",    label:"Device voltage (lowest)", cat:"Battery & electrical", source:"status", diag:"DiagnosticGoDeviceVoltageId", agg:"min", unit:"V", vtype:"volts" },

    { id:"tire_fl",      label:"Tire pressure FL (avg)", cat:"Tires", source:"status", diag:"DiagnosticTirePressureFrontLeftId", agg:"avg", unit:"kPa", vtype:"pressurePa" },
    { id:"tire_fr",      label:"Tire pressure FR (avg)", cat:"Tires", source:"status", diag:"DiagnosticTirePressureFrontRightId", agg:"avg", unit:"kPa", vtype:"pressurePa" },
    { id:"tire_rl",      label:"Tire pressure RL (avg)", cat:"Tires", source:"status", diag:"DiagnosticTirePressureRearLeftId", agg:"avg", unit:"kPa", vtype:"pressurePa" },
    { id:"tire_rr",      label:"Tire pressure RR (avg)", cat:"Tires", source:"status", diag:"DiagnosticTirePressureRearRightId", agg:"avg", unit:"kPa", vtype:"pressurePa" },

    { id:"outside_temp", label:"Outside temp (avg)", cat:"Environment", source:"status", diag:"DiagnosticOutsideTemperatureId", agg:"avg", unit:"\u00B0C", vtype:"tempC" },

    { id:"brake_g",      label:"Hardest braking (g)", cat:"Motion", source:"status", diag:"DiagnosticAccelerationForwardBrakingId", agg:"min", unit:"g", vtype:"gforce" },
    { id:"corner_g",     label:"Hardest cornering (g)", cat:"Motion", source:"status", diag:"DiagnosticAccelerationSideToSideId", agg:"max", unit:"g", vtype:"gforce" },

    { id:"breakdown_risk", label:"Predicted breakdown risk (peak)", cat:"Predictive", source:"status", diag:"DiagnosticPredictedRiskOfBreakdownId", agg:"max", unit:"%", vtype:"pct" },
    { id:"elec_rating",  label:"Electrical system rating (latest)", cat:"Predictive", source:"status", diag:"DiagnosticElectricalSystemRatingId", agg:"last", unit:"", vtype:"num" },

    { id:"faults_all",   label:"Faults logged", cat:"Faults", source:"fault", agg:"count", unit:"faults", vtype:"int" },
    { id:"faults_mil",   label:"Check-engine (MIL) faults", cat:"Faults", source:"fault", agg:"count", unit:"faults", vtype:"int", lamp:"malfunctionLamp" },
    { id:"faults_battery", label:"Low-battery faults", cat:"Faults", source:"fault", agg:"count", unit:"faults", vtype:"int", faultDiag:"DiagnosticVehicleBatteryLowVoltageId" },
    { id:"faults_accident", label:"Accident-level accel events", cat:"Faults", source:"fault", agg:"count", unit:"faults", vtype:"int", faultDiag:"DiagnosticAccidentLevelAccelerationEventId" }
  ];
  var DIMENSIONS = [
    { id:"vehicle", label:"Vehicle", dkey:"vehicle", time:false, card:"high" },
    { id:"driver",  label:"Driver",  dkey:"driver",  time:false, card:"high" },
    { id:"group",   label:"Group",   dkey:"group",   time:false, card:"low" },
    { id:"day",     label:"Day",     dkey:"day",     time:true },
    { id:"week",    label:"Week",    dkey:"week",     time:true }
  ];
  var MBYID = {}; MEASURES.forEach(function(m){ MBYID[m.id]=m; });
  var DBYID = {}; DIMENSIONS.forEach(function(d){ DBYID[d.id]=d; });
  var SOURCES = { exception:{ typeName:"ExceptionEvent", dateField:"activeFrom" }, trip:{ typeName:"Trip", dateField:"start" } };
  function dateFieldFor(src){ if(src.indexOf("status:")===0) return "dateTime"; if(src==="fault") return "dateTime"; return (SOURCES[src]||{}).dateField || "dateTime"; }

  var S = {
    api:null, live:false, ready:false,
    days:30, customFrom:null, customTo:null,
    groupSel:[], vehSel:[], units:"metric",
    measures:[], dims:[], chartType:"auto",
    title:"Untitled chart", titleEdited:false,
    meta:{ devices:[], deviceById:{}, groups:[], groupById:{}, driverById:{}, rules:[], matchedRuleIds:{} },
    cache:{}, lastResult:null, lastChartFn:null,
    tiles:[], dashName:"Untitled dashboard", savedList:[], tab:"build",
    dynLoaded:false, showAllDiag:false
  };

  function $(id){ return document.getElementById(id); }
  function ce(t,c,txt){ var e=document.createElement(t); if(c) e.className=c; if(txt!=null) e.textContent=txt; return e; }
  function svg(tag,attrs){ var e=document.createElementNS(SVGNS,tag); if(attrs) for(var k in attrs) e.setAttribute(k,attrs[k]); return e; }
  function clear(el){ while(el.firstChild) el.removeChild(el.firstChild); }
  function debounce(fn,ms){ var t; return function(){ var a=arguments,c=this; clearTimeout(t); t=setTimeout(function(){ fn.apply(c,a); },ms); }; }
  function trim(s,n){ s=String(s); return s.length>n? s.slice(0,n-1)+"\u2026":s; }

  function toast(msg){ var t=$("vizToast"); t.textContent=msg; t.classList.add("on"); clearTimeout(toast._t); toast._t=setTimeout(function(){ t.classList.remove("on"); },2200); }
  function log(msg,kind){ var p=$("logPanel"); var ln=ce("div","ln"+(kind?(" "+kind):"")); var ts=new Date().toTimeString().slice(0,8); ln.appendChild(ce("span","t",ts+"  ")); ln.appendChild(document.createTextNode(msg)); p.appendChild(ln); p.scrollTop=p.scrollHeight; }

  var KM2MI=0.621371;
  function fmtNum(n,dp){ if(n==null||isNaN(n)) return "0"; return Number(n).toLocaleString(undefined,{minimumFractionDigits:dp||0,maximumFractionDigits:dp||0}); }
  function convVal(vtype,v){
    var imp=S.units==="imperial";
    if(vtype==="dist"||vtype==="speed"||vtype==="speedKph") return imp? v*KM2MI : v;
    if(vtype==="distM") return imp? (v/1000)*KM2MI : v/1000;
    if(vtype==="tempC") return imp? v*9/5+32 : v;
    if(vtype==="pressurePa") return imp? v*0.000145038 : v/1000;
    if(vtype==="volumeL") return imp? v*0.264172 : v;
    if(vtype==="hoursSec") return v/3600;
    if(vtype==="gforce") return v/9.81;
    return v;
  }
  function unitLabel(m){
    var imp=S.units==="imperial", t=m.vtype;
    if(t==="dist"||t==="distM") return imp?"mi":"km";
    if(t==="speed"||t==="speedKph") return imp?"mph":"km/h";
    if(t==="tempC") return imp?"\u00B0F":"\u00B0C";
    if(t==="pressurePa") return imp?"psi":"kPa";
    if(t==="volumeL") return imp?"gal":"L";
    if(t==="hoursSec") return "h";
    if(t==="gforce") return "g";
    if(t==="volts") return "V";
    if(t==="rpm") return "rpm";
    return m.unit;
  }
  function fmtMeasure(m,v,longForm){
    v=convVal(m.vtype,v);
    if(m.vtype==="int") return fmtNum(v,0);
    if(m.vtype==="pct") return fmtNum(v,1);
    if(m.vtype==="dist"||m.vtype==="distM") return fmtNum(v, v>=100?0:1);
    if(m.vtype==="speed"||m.vtype==="speedKph") return fmtNum(v,1);
    if(m.vtype==="tempC") return fmtNum(v,1);
    if(m.vtype==="pressurePa") return fmtNum(v, v>=100?0:1);
    if(m.vtype==="volumeL") return fmtNum(v, v>=100?0:1);
    if(m.vtype==="volts") return fmtNum(v,2);
    if(m.vtype==="rpm") return fmtNum(v,0);
    if(m.vtype==="gforce") return fmtNum(v,2);
    if(m.vtype==="hoursSec") return fmtNum(v, v>=100?0:1);
    if(m.vtype==="dur"){ if(longForm){ var h=Math.floor(v); var mi=Math.round((v-h)*60); return h+"h "+(mi<10?"0":"")+mi+"m"; } return fmtNum(v, v>=100?0:1); }
    return fmtNum(v,1);
  }

  function durHours(s){
    if(s==null) return 0;
    if(typeof s==="number") return s;
    s=String(s);
    var m=s.match(/^(?:(\d+)\.)?(\d+):(\d+):(\d+)(?:\.\d+)?$/);
    if(m){ return (+(m[1]||0))*24 + (+m[2]) + (+m[3])/60 + (+m[4])/3600; }
    var iso=s.match(/^P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?$/);
    if(iso){ return (+(iso[1]||0))*24 + (+(iso[2]||0)) + (+(iso[3]||0))/60 + (parseFloat(iso[4]||0))/3600; }
    var f=parseFloat(s); return isNaN(f)?0:f;
  }

  function pad2(n){ return (n<10?"0":"")+n; }
  function localDayKey(d){ return d.getFullYear()+"-"+pad2(d.getMonth()+1)+"-"+pad2(d.getDate()); }
  function weekStartKey(d){ var x=new Date(d.getFullYear(),d.getMonth(),d.getDate()); var dow=(x.getDay()+6)%7; x.setDate(x.getDate()-dow); return localDayKey(x); }
  function rangeISO(){
    var to,from;
    if(S.days==="custom" && S.customFrom && S.customTo){ from=new Date(S.customFrom+"T00:00:00"); to=new Date(S.customTo+"T23:59:59"); }
    else { to=new Date(); from=new Date(); from.setDate(from.getDate()-(+S.days)); from.setHours(0,0,0,0); }
    return { fromIso:from.toISOString(), toIso:to.toISOString(), from:from, to:to };
  }
  function eachDay(from,to){ var out=[],d=new Date(from.getFullYear(),from.getMonth(),from.getDate()); var end=new Date(to.getFullYear(),to.getMonth(),to.getDate()); while(d<=end){ out.push(localDayKey(d)); d.setDate(d.getDate()+1); } return out; }
  function eachWeek(from,to){ var out=[],seen={}; var d=new Date(from.getFullYear(),from.getMonth(),from.getDate()); var end=new Date(to.getFullYear(),to.getMonth(),to.getDate()); while(d<=end){ var k=weekStartKey(d); if(!seen[k]){ seen[k]=1; out.push(k); } d.setDate(d.getDate()+1); } return out; }
  function shortDay(k){ var p=k.split("-"); return (+p[1])+"/"+(+p[2]); }

  function gcall(method,params){ return new Promise(function(res,rej){ S.api.call(method,params,res,rej); }); }
  function gmulti(calls){ return new Promise(function(res,rej){ S.api.multiCall(calls,res,rej); }); }

  function deviceInGroups(dv,groupIds){ var gs=dv.groups||[]; for(var i=0;i<gs.length;i++){ if(groupIds.indexOf(gs[i].id)>=0) return true; } return false; }
  function scopeDeviceIds(){
    if(S.vehSel.length){ var set={}; S.vehSel.forEach(function(id){ set[id]=1; }); if(S.groupSel.length){ S.meta.devices.forEach(function(dv){ if(deviceInGroups(dv,S.groupSel)) set[dv.id]=1; }); } return Object.keys(set); }
    if(S.groupSel.length){ return S.meta.devices.filter(function(dv){ return deviceInGroups(dv,S.groupSel); }).map(function(dv){ return dv.id; }); }
    return S.meta.devices.map(function(dv){ return dv.id; });
  }
  function scopeHash(ids){ return ids.slice().sort().join(",").length+":"+ids.length; }
  function cacheKey(source,r,ids){ return source+"|"+r.fromIso+"|"+r.toIso+"|"+scopeHash(ids); }

  function fetchRows(source){
    var r=rangeISO(); var ids=scopeDeviceIds(); var key=cacheKey(source,r,ids);
    if(S.cache[key]) return Promise.resolve(S.cache[key]);
    if(!S.live){ var rows=mockRows(source,r,ids); S.cache[key]=rows; log("demo: "+rows.length+" "+source+" rows for "+ids.length+" vehicles","ok"); return Promise.resolve(rows); }
    if(!ids.length) return Promise.resolve([]);
    var diag=null, src=source; if(source.indexOf("status:")===0){ src="status"; diag=source.slice(7); }
    var tn = src==="status"?"StatusData" : src==="fault"?"FaultData" : SOURCES[src].typeName;
    var calls=ids.map(function(id){ var search={ deviceSearch:{id:id}, fromDate:r.fromIso, toDate:r.toIso }; if(diag) search.diagnosticSearch={id:diag}; return ["Get",{ typeName:tn, search:search, resultsLimit:50000 }]; });
    return runChunked(calls).then(function(rows){ S.cache[key]=rows; log(tn+": loaded "+rows.length+" rows across "+ids.length+" vehicles","ok"); return rows; });
  }
  function runChunked(calls){
    var size=40, out=[], i=0, total=calls.length;
    function step(){
      if(i>=calls.length) return Promise.resolve(out);
      var batch=calls.slice(i,i+size); i+=size;
      setMode("live","Loading "+Math.min(i,total)+"/"+total+" vehicles\u2026");
      return gmulti(batch).then(function(res){ res.forEach(function(arr){ if(arr&&arr.length) out=out.concat(arr); }); return step(); });
    }
    return step().then(function(rows){ setMode("live","Connected"); return rows; });
  }

  // ---- aggregation: raw rows -> measure values keyed by dimension --------
  function dimKeyOf(dim,row){
    if(dim.dkey==="vehicle"){ var dev=row.device&&row.device.id; var o=S.meta.deviceById[dev]; return dev?[{k:dev,label:o?o.name:dev}]:[{k:"?",label:"Unknown vehicle"}]; }
    if(dim.dkey==="driver"){ var dr=row.driver&&row.driver.id; if(!dr||dr==="UnknownDriverId"||dr==="NoDriverId") return [{k:"_none",label:"Unassigned"}]; var u=S.meta.driverById[dr]; return [{k:dr,label:u?u.name:dr}]; }
    if(dim.dkey==="group"){ var dev2=row.device&&row.device.id; var o2=S.meta.deviceById[dev2]; var gs=(o2&&o2.groups)||[]; var picks=[]; gs.forEach(function(g){ var gi=S.meta.groupById[g.id]; if(gi&&!gi.isRoot) picks.push({k:g.id,label:gi.name}); }); return picks.length?picks:[{k:"_ungrouped",label:"Ungrouped"}]; }
    var ds=row.__date; if(dim.dkey==="day") return [{k:localDayKey(ds),label:null}];
    if(dim.dkey==="week") return [{k:weekStartKey(ds),label:null}];
    return [{k:"_all",label:"All"}];
  }
  function measureRows(m,rowsBySource){
    if(m.source==="status") return rowsBySource["status:"+m.diag]||[];
    var rows=rowsBySource[m.source]||[];
    if(m.source==="exception"){ if(m.ruleId){ rows=rows.filter(function(r){ return r.rule && r.rule.id===m.ruleId; }); } else if(!m.matchAll){ var ids=S.meta.matchedRuleIds[m.id]||{}; rows=rows.filter(function(r){ return r.rule && ids[r.rule.id]; }); } }
    if(m.source==="fault"){ if(m.faultDiag) rows=rows.filter(function(r){ return r.diagnostic && r.diagnostic.id===m.faultDiag; }); else if(m.lamp) rows=rows.filter(function(r){ return r[m.lamp]===true; }); }
    return rows;
  }
  function valOf(m,r){ if(m.source==="status"){ var d=r.data; return (d==null||isNaN(d))?0:+d; } if(m.durField) return durHours(r[m.field]); var v=r[m.field]; return (v==null||isNaN(v))?0:+v; }
  function finalize(m,a){
    if(m.agg==="count") return a.cnt;
    if(m.agg==="sum") return a.sum;
    if(m.agg==="avg") return a.cnt? a.sum/a.cnt : 0;
    if(m.agg==="max") return a.max===-Infinity?0:a.max;
    if(m.agg==="min") return a.min===Infinity?0:a.min;
    if(m.agg==="last") return a.lVal||0;
    if(m.agg==="delta") return Math.max(0,(a.lVal||0)-(a.fVal||0));
    if(m.agg==="derived"){ var den=a.c1+a.c2; return den? (a.c1/den)*100 : 0; }
    return 0;
  }
  function aggregateOne(m,rows,keyFn){
    if(m.agg==="delta"||m.agg==="last"){
      var accn={};
      rows.forEach(function(r){ var dev=(r.device&&r.device.id)||"_"; var t=r.__date?r.__date.getTime():0; var v=valOf(m,r); var keys=keyFn(r);
        for(var j=0;j<keys.length;j++){ var dk=keys[j].k; var g=accn[dk]||(accn[dk]={}); var d=g[dev]||(g[dev]={fAt:null,fVal:0,lAt:null,lVal:0}); if(d.fAt===null||t<d.fAt){ d.fAt=t; d.fVal=v; } if(d.lAt===null||t>=d.lAt){ d.lAt=t; d.lVal=v; } }
      });
      var outn={}; Object.keys(accn).forEach(function(dk){ var devs=accn[dk], s=0, c=0; Object.keys(devs).forEach(function(dv){ var d=devs[dv]; if(m.agg==="delta") s+=Math.max(0,d.lVal-d.fVal); else { s+=d.lVal; c++; } }); outn[dk]= m.agg==="delta"? s : (c? s/c : 0); }); return outn;
    }
    var acc={};
    function bump(k){ if(!acc[k]) acc[k]={sum:0,cnt:0,max:-Infinity,min:Infinity,c1:0,c2:0}; return acc[k]; }
    rows.forEach(function(r){ var keys=keyFn(r); var v=(m.agg==="count")?1:valOf(m,r);
      for(var j=0;j<keys.length;j++){ var a=bump(keys[j].k);
        if(m.agg==="count"){ a.cnt+=1; }
        else if(m.agg==="sum"){ a.sum+=v; }
        else if(m.agg==="avg"){ a.sum+=v; a.cnt+=1; }
        else if(m.agg==="max"){ if(v>a.max) a.max=v; }
        else if(m.agg==="min"){ if(v<a.min) a.min=v; }
        else if(m.agg==="derived"){ a.c1+=durHours(r[m.components[0]]); a.c2+=durHours(r[m.components[1]]); }
      } });
    var out={}; Object.keys(acc).forEach(function(k){ out[k]=finalize(m,acc[k]); }); return out;
  }

  function buildResult(rowsBySource){
    var dims=S.dims.map(function(id){ return DBYID[id]; });
    var measures=S.measures.map(function(id){ return MBYID[id]; });
    var r=rangeISO();
    for(var src in rowsBySource){ var df=dateFieldFor(src); rowsBySource[src].forEach(function(row){ if(!row.__date) row.__date=new Date(row[df]); }); }

    var res={ dims:dims, measures:measures };
    if(dims.length===0){
      res.kind="scalar"; res.scalars={};
      measures.forEach(function(m){ var rows=measureRows(m,rowsBySource); var agg=aggregateOne(m,rows,function(){return [{k:"_all"}];}); res.scalars[m.id]=agg["_all"]||0; });
      return res;
    }
    if(dims.length===1){
      var dim=dims[0]; res.kind="one"; res.dim=dim; res.byMeasure={};
      var keyset={}, labels={};
      measures.forEach(function(m){ var rows=measureRows(m,rowsBySource); var agg=aggregateOne(m,rows,function(row){ return dimKeyOf(dim,row); }); res.byMeasure[m.id]=agg; Object.keys(agg).forEach(function(k){ keyset[k]=1; });
        if(dim.time===false){ rows.forEach(function(row){ dimKeyOf(dim,row).forEach(function(p){ labels[p.k]=p.label; }); }); } });
      if(dim.time){ res.keys = dim.dkey==="day"? eachDay(r.from,r.to) : eachWeek(r.from,r.to); res.labelFor=function(k){ return shortDay(k); }; }
      else { var primary=measures[0].id; res.keys=Object.keys(keyset).sort(function(a,b){ return (res.byMeasure[primary][b]||0)-(res.byMeasure[primary][a]||0); }); res.fullCount=res.keys.length; res.keys=res.keys.slice(0,TOPN); res.labelFor=function(k){ return labels[k]||k; }; }
      return res;
    }
    // two dims
    res.kind="two"; res.dimA=dims[0]; res.dimB=dims[1]; res.measure=measures[0];
    var m0=measures[0]; var rows0=measureRows(m0,rowsBySource);
    var grid={}, ak={}, bk={}, alab={}, blab={};
    rows0.forEach(function(row){ var ka=dimKeyOf(dims[0],row), kb=dimKeyOf(dims[1],row);
      ka.forEach(function(pa){ kb.forEach(function(pb){ ak[pa.k]=1; bk[pb.k]=1; if(pa.label) alab[pa.k]=pa.label; if(pb.label) blab[pb.k]=pb.label;
        var key=pa.k+"||"+pb.k; if(!grid[key]) grid[key]={sum:0,cnt:0,max:-Infinity,min:Infinity,c1:0,c2:0,fAt:null,fVal:0,lAt:null,lVal:0}; var a=grid[key];
        var v=(m0.agg==="count")?1:valOf(m0,row); var tt=row.__date?row.__date.getTime():0;
        if(m0.agg==="count") a.cnt+=1; else if(m0.agg==="sum") a.sum+=v; else if(m0.agg==="avg"){ a.sum+=v; a.cnt+=1; }
        else if(m0.agg==="max"){ if(v>a.max)a.max=v; } else if(m0.agg==="min"){ if(v<a.min)a.min=v; }
        else if(m0.agg==="last"){ if(a.lAt===null||tt>=a.lAt){ a.lAt=tt; a.lVal=v; } }
        else if(m0.agg==="delta"){ if(a.fAt===null||tt<a.fAt){ a.fAt=tt; a.fVal=v; } if(a.lAt===null||tt>=a.lAt){ a.lAt=tt; a.lVal=v; } }
        else if(m0.agg==="derived"){ a.c1+=durHours(row[m0.components[0]]); a.c2+=durHours(row[m0.components[1]]); }
      }); });
    });
    function tot(k){ var s=0; Object.keys(bk).forEach(function(b){ var c=grid[k+"||"+b]; if(c) s+=finalize(m0,c); }); return s; }
    var keysA=dims[0].time? (dims[0].dkey==="day"?eachDay(r.from,r.to):eachWeek(r.from,r.to)) : Object.keys(ak).sort(function(a,b){ return tot(b)-tot(a); }).slice(0,TOPN);
    var keysB=Object.keys(bk).sort(function(a,b){ var sa=0,sb=0; keysA.forEach(function(k){ var ca=grid[k+"||"+a], cb=grid[k+"||"+b]; if(ca)sa+=finalize(m0,ca); if(cb)sb+=finalize(m0,cb); }); return sb-sa; }).slice(0,TOPN_SERIES);
    res.keysA=keysA; res.keysB=keysB;
    res.valAt=function(a,b){ var c=grid[a+"||"+b]; return c?finalize(m0,c):0; };
    res.labelA=function(k){ return dims[0].time? shortDay(k) : (alab[k]||k); };
    res.labelB=function(k){ return blab[k]||k; };
    return res;
  }

  // ---- auto-viz chooser: shape of the selection -> chart -----------------
  function shapeOf(){
    var ms=S.measures.map(function(id){return MBYID[id];});
    var ds=S.dims.map(function(id){return DBYID[id];});
    var timeDims=ds.filter(function(d){return d.time;});
    var catDims=ds.filter(function(d){return !d.time;});
    return { M:ms.length, D:ds.length, hasTime:timeDims.length>0, timeDim:timeDims[0], catDims:catDims, ms:ms, ds:ds };
  }
  function validCharts(sp){
    var v=[]; if(sp.M===0) return v;
    if(sp.D===0){ v.push("kpi"); return v; }
    if(sp.D===1){ if(sp.hasTime){ v.push("line"); } else { v.push("bar"); if(sp.M===2) v.push("scatter"); } v.push("table"); return v; }
    if(sp.hasTime && sp.catDims.length===1) v.push("line"); else v.push("bar");
    v.push("table"); return v;
  }
  function autoChart(sp){
    if(sp.M===0) return null;
    if(sp.D===0) return "kpi";
    if(sp.D===1) return sp.hasTime?"line":"bar";
    if(sp.hasTime && sp.catDims.length===1) return "line";
    return "bar";
  }
  function chartLabel(t){ return {kpi:"Number tiles",line:"Line",bar:"Bars",scatter:"Scatter",table:"Table"}[t]||t; }

  // ---- SVG chart renderers ----------------------------------------------
  function mWidth(host){ return Math.max(320, host.clientWidth-32); }
  function niceMax(v){ if(v<=0) return 1; var p=Math.pow(10,Math.floor(Math.log10(v))); var f=v/p; var n=f<=1?1:f<=2?2:f<=5?5:10; return n*p; }
  function longestLabel(groups){ var m=4; groups.forEach(function(g){ m=Math.max(m,String(g.label).length); }); return Math.min(26,m); }
  function noteEl(t){ var d=ce("div"); d.style.cssText="color:var(--faint);font-size:11px;margin:-4px 0 10px;"; d.textContent=t; return d; }
  function legend(host,items){ var l=ce("div","legend"); items.forEach(function(it){ var x=ce("div","it"); var sw=ce("span","sw"); sw.style.background=it.color; x.appendChild(sw); x.appendChild(document.createTextNode(it.label)); l.appendChild(x); }); host.appendChild(l); }
  function tipBind(node,html){ node.addEventListener("mousemove",function(e){ var t=$("vizTip"); t.innerHTML=html; t.classList.add("on"); t.style.left=(e.clientX+12)+"px"; t.style.top=(e.clientY+12)+"px"; }); node.addEventListener("mouseleave",function(){ $("vizTip").classList.remove("on"); }); }

  function renderKPI(host,res){
    var wrap=ce("div","kpis");
    res.measures.forEach(function(m,i){ var v=res.scalars[m.id]||0; var card=ce("div","kpi"); card.appendChild(ce("div","kl",m.label));
      var kv=ce("div","kv"); kv.style.color=PALETTE[i%PALETTE.length]; kv.appendChild(document.createTextNode(fmtMeasure(m,v,true))); kv.appendChild(ce("span","ku",m.vtype==="dur"?"":unitLabel(m))); card.appendChild(kv); wrap.appendChild(card); });
    host.appendChild(wrap);
  }

  function renderLine(host,res){
    var xkeys, xlabel, series;
    if(res.kind==="one"){ xkeys=res.keys; xlabel=function(k){return res.labelFor(k);};
      series=res.measures.map(function(m,i){ return { name:m.label, m:m, color:PALETTE[i%PALETTE.length], pts:xkeys.map(function(k){ return res.byMeasure[m.id][k]||0; }) }; }); }
    else { var timeIsA=res.dimA.time; var xs=timeIsA?res.keysA:res.keysB; var cats=timeIsA?res.keysB:res.keysA; xkeys=xs; xlabel=function(k){ return timeIsA?res.labelA(k):res.labelB(k); };
      series=cats.map(function(c,i){ return { name:(timeIsA?res.labelB(c):res.labelA(c)), m:res.measure, color:PALETTE[i%PALETTE.length], pts:xs.map(function(x){ return timeIsA?res.valAt(x,c):res.valAt(c,x); }) }; }); }
    var dual=false, axisOf=function(){return 0;};
    if(res.kind==="one" && series.length===2 && series[0].m.vtype!==series[1].m.vtype){ dual=true; axisOf=function(i){return i===1?1:0;}; }
    var W=mWidth(host), H=340, padL=52, padR=dual?56:18, padT=14, padB=42;
    var max0=0,max1=0; series.forEach(function(s,i){ s.pts.forEach(function(p){ var c=convVal(s.m.vtype,p); if(axisOf(i)===1){ if(c>max1)max1=c; } else if(c>max0)max0=c; }); }); max0=niceMax(max0); max1=niceMax(max1||1);
    legend(host,series.map(function(s){ return {label:s.name+(dual?" ("+unitLabel(s.m)+")":""), color:s.color}; }));
    var s1=svg("svg",{class:"chart",viewBox:"0 0 "+W+" "+H,width:"100%",height:H});
    var plotW=W-padL-padR, plotH=H-padT-padB; var xstep=xkeys.length>1?plotW/(xkeys.length-1):0;
    var g=svg("g",{class:"grid"});
    for(var t=0;t<=4;t++){ var yy=padT+plotH*(t/4); g.appendChild(svg("line",{x1:padL,y1:yy,x2:padL+plotW,y2:yy}));
      var lab=svg("text",{x:padL-8,y:yy+3,"text-anchor":"end","font-size":"10",class:"axlab"}); lab.textContent=fmtNum(max0*(1-t/4), max0>=100?0:1); g.appendChild(lab);
      if(dual){ var lab2=svg("text",{x:padL+plotW+8,y:yy+3,"text-anchor":"start","font-size":"10"}); lab2.textContent=fmtNum(max1*(1-t/4),max1>=100?0:1); lab2.setAttribute("fill",series[1].color); g.appendChild(lab2); } }
    s1.appendChild(g);
    var every=Math.ceil(xkeys.length/10);
    xkeys.forEach(function(k,i){ if(i%every!==0 && i!==xkeys.length-1) return; var xx=padL+xstep*i; var lab=svg("text",{x:xx,y:H-padB+16,"text-anchor":"middle","font-size":"10",class:"axlab"}); lab.textContent=xlabel(k); s1.appendChild(lab); });
    series.forEach(function(s,si){ var mx=axisOf(si)===1?max1:max0; var d="", coords=[];
      s.pts.forEach(function(p,i){ var c=convVal(s.m.vtype,p); var xx=padL+xstep*i; var yy=padT+plotH*(1-(mx?c/mx:0)); coords.push([xx,yy,p]); d+=(i?"L":"M")+xx.toFixed(1)+" "+yy.toFixed(1)+" "; });
      s1.appendChild(svg("path",{d:d,fill:"none",stroke:s.color,"stroke-width":"2.4","stroke-linejoin":"round","stroke-linecap":"round"}));
      coords.forEach(function(c,i){ if(xkeys.length>60) return; var dot=svg("circle",{class:"dot",cx:c[0],cy:c[1],r:"3",fill:"#fff",stroke:s.color,"stroke-width":"2"}); tipBind(dot,"<b>"+xlabel(xkeys[i])+"</b><br>"+s.name+': <span class="tk">'+fmtMeasure(s.m,c[2],true)+" "+(s.m.vtype==="dur"?"":unitLabel(s.m))+"</span>"); s1.appendChild(dot); });
    });
    host.appendChild(s1);
  }

  function renderBar(host,res){
    var groups, seriesNames, colorFor, stacked=false, getVal, yMeasure;
    if(res.kind==="one"){
      groups=res.keys.map(function(k){ return {key:k,label:res.labelFor(k)}; });
      seriesNames=res.measures.map(function(m){ return m.label; }); yMeasure=res.measures[0];
      getVal=function(gk,si){ return res.byMeasure[res.measures[si].id][gk]||0; }; colorFor=function(si){ return PALETTE[si%PALETTE.length]; };
      if(res.fullCount>res.keys.length) host.appendChild(noteEl("Showing top "+res.keys.length+" of "+res.fullCount+" "+res.dim.label.toLowerCase()+"s by "+res.measures[0].label.toLowerCase()+"."));
    } else {
      stacked=true; yMeasure=res.measure; groups=res.keysA.map(function(k){ return {key:k,label:res.labelA(k)}; });
      if(res.dimA.time===false && res.keysA.length>=TOPN) host.appendChild(noteEl("Showing top "+res.keysA.length+" "+res.dimA.label.toLowerCase()+"s; stacked by "+res.dimB.label.toLowerCase()+"."));
      seriesNames=res.keysB.map(function(b){ return res.labelB(b); }); getVal=function(gk,si){ return res.valAt(gk,res.keysB[si]); }; colorFor=function(si){ return PALETTE[si%PALETTE.length]; };
    }
    var nS=seriesNames.length;
    legend(host,seriesNames.map(function(n,i){ return {label:n,color:colorFor(i)}; }));
    var W=mWidth(host), rowH=30, padL=Math.min(180, Math.max(90, longestLabel(groups)*7)), padR=20, padT=6, padB=24;
    var H=padT+padB+groups.length*rowH+6; var s1=svg("svg",{class:"chart",viewBox:"0 0 "+W+" "+H,width:"100%",height:H}); var plotW=W-padL-padR;
    var max=0; groups.forEach(function(g){ if(stacked){ var s=0; for(var i=0;i<nS;i++) s+=convVal(yMeasure.vtype,getVal(g.key,i)); if(s>max)max=s; } else { for(var i2=0;i2<nS;i2++){ var v=convVal(yMeasure.vtype,getVal(g.key,i2)); if(v>max)max=v; } } }); max=niceMax(max);
    var gr=svg("g",{class:"grid"});
    for(var t=0;t<=4;t++){ var xx=padL+plotW*(t/4); gr.appendChild(svg("line",{x1:xx,y1:padT,x2:xx,y2:padT+groups.length*rowH})); var lab=svg("text",{x:xx,y:H-8,"text-anchor":"middle","font-size":"10",class:"axlab"}); lab.textContent=fmtNum(max*(t/4),max>=100?0:1); gr.appendChild(lab); }
    s1.appendChild(gr);
    groups.forEach(function(g,gi){ var y0=padT+gi*rowH; var lab=svg("text",{x:padL-8,y:y0+rowH/2+3,"text-anchor":"end","font-size":"11",fill:"var(--ink-2)"}); lab.textContent=trim(g.label,26); s1.appendChild(lab);
      if(stacked){ var xacc=padL; for(var i=0;i<nS;i++){ var v=convVal(yMeasure.vtype,getVal(g.key,i)); var w=max?(v/max)*plotW:0; if(w<=0) continue; var bh=rowH-10; var rect=svg("rect",{class:"bar",x:xacc,y:y0+5,width:Math.max(0,w),height:bh,fill:colorFor(i),rx:i===0||i===nS-1?2:0}); tipBind(rect,"<b>"+g.label+"</b><br>"+seriesNames[i]+': <span class="tk">'+fmtMeasure(yMeasure,getVal(g.key,i),true)+"</span>"); s1.appendChild(rect); xacc+=w; } }
      else { var bh2=(rowH-8)/nS; for(var j=0;j<nS;j++){ var v2=convVal(yMeasure.vtype,getVal(g.key,j)); var w2=max?(v2/max)*plotW:0; var yy=y0+4+j*bh2; var mj=res.kind==="one"?res.measures[j]:yMeasure; var r2=svg("rect",{class:"bar",x:padL,y:yy,width:Math.max(0,w2),height:Math.max(2,bh2-2),fill:colorFor(j),rx:2}); tipBind(r2,"<b>"+g.label+"</b><br>"+seriesNames[j]+': <span class="tk">'+fmtMeasure(mj,getVal(g.key,j),true)+" "+(mj.vtype==="dur"?"":unitLabel(mj))+"</span>"); s1.appendChild(r2); if(nS===1){ var vt=svg("text",{x:padL+w2+6,y:yy+bh2/2+3,"font-size":"10.5",fill:"var(--muted)"}); vt.textContent=fmtMeasure(mj,getVal(g.key,j),false); s1.appendChild(vt); } } }
    });
    host.appendChild(s1);
  }

  function renderScatter(host,res){
    var mx=res.measures[0], my=res.measures[1];
    var pts=res.keys.map(function(k){ return { label:res.labelFor(k), x:res.byMeasure[mx.id][k]||0, y:res.byMeasure[my.id][k]||0 }; });
    var W=mWidth(host),H=360,padL=58,padR=18,padT=14,padB=46;
    var maxX=niceMax(Math.max.apply(null,pts.map(function(p){return convVal(mx.vtype,p.x);}).concat([1])));
    var maxY=niceMax(Math.max.apply(null,pts.map(function(p){return convVal(my.vtype,p.y);}).concat([1])));
    var s1=svg("svg",{class:"chart",viewBox:"0 0 "+W+" "+H,width:"100%",height:H}); var plotW=W-padL-padR, plotH=H-padT-padB; var gr=svg("g",{class:"grid"});
    for(var t=0;t<=4;t++){ var yy=padT+plotH*(t/4); gr.appendChild(svg("line",{x1:padL,y1:yy,x2:padL+plotW,y2:yy})); var l=svg("text",{x:padL-8,y:yy+3,"text-anchor":"end","font-size":"10",class:"axlab"}); l.textContent=fmtNum(maxY*(1-t/4),maxY>=100?0:1); gr.appendChild(l);
      var xx=padL+plotW*(t/4); gr.appendChild(svg("line",{x1:xx,y1:padT,x2:xx,y2:padT+plotH})); var l2=svg("text",{x:xx,y:H-padB+16,"text-anchor":"middle","font-size":"10",class:"axlab"}); l2.textContent=fmtNum(maxX*(t/4),maxX>=100?0:1); gr.appendChild(l2); }
    s1.appendChild(gr);
    var ax=svg("text",{x:padL+plotW/2,y:H-8,"text-anchor":"middle","font-size":"11",fill:"var(--muted)"}); ax.textContent=mx.label+" ("+(mx.vtype==="dur"?"h":unitLabel(mx))+") \u2192"; s1.appendChild(ax);
    var ay=svg("text",{x:14,y:padT+plotH/2,"text-anchor":"middle","font-size":"11",fill:"var(--muted)",transform:"rotate(-90 14 "+(padT+plotH/2)+")"}); ay.textContent=my.label+" ("+(my.vtype==="dur"?"h":unitLabel(my))+") \u2192"; s1.appendChild(ay);
    pts.forEach(function(p){ var cx=padL+(maxX?convVal(mx.vtype,p.x)/maxX:0)*plotW; var cy=padT+plotH*(1-(maxY?convVal(my.vtype,p.y)/maxY:0)); var c=svg("circle",{class:"pt",cx:cx,cy:cy,r:"6",fill:PALETTE[0],"fill-opacity":"0.72",stroke:"#fff","stroke-width":"1.5"}); tipBind(c,"<b>"+p.label+"</b><br>"+mx.label+': <span class="tk">'+fmtMeasure(mx,p.x,true)+"</span><br>"+my.label+': <span class="tk">'+fmtMeasure(my,p.y,true)+"</span>"); s1.appendChild(c); });
    host.appendChild(s1);
  }

  function renderTable(host,res){
    var tbl=ce("table","dtable"), thead=ce("thead"), tr=ce("tr"); var rows=[], cols=[];
    if(res.kind==="scalar"){ cols=[{t:"Measure"},{t:"Value",n:true}]; res.measures.forEach(function(m){ rows.push([m.label, fmtMeasure(m,res.scalars[m.id],true)+" "+(m.vtype==="dur"?"":unitLabel(m))]); }); }
    else if(res.kind==="one"){ cols=[{t:res.dim.label}]; res.measures.forEach(function(m){ cols.push({t:m.label,n:true}); }); res.keys.forEach(function(k){ var row=[res.labelFor(k)]; res.measures.forEach(function(m){ row.push(fmtMeasure(m,res.byMeasure[m.id][k]||0,false)); }); rows.push(row); }); }
    else { cols=[{t:res.dimA.label}]; res.keysB.forEach(function(b){ cols.push({t:res.labelB(b),n:true}); }); res.keysA.forEach(function(a){ var row=[res.labelA(a)]; res.keysB.forEach(function(b){ row.push(fmtMeasure(res.measure,res.valAt(a,b),false)); }); rows.push(row); }); }
    cols.forEach(function(c){ tr.appendChild(ce("th",c.n?"n":null,c.t)); }); thead.appendChild(tr); tbl.appendChild(thead);
    var tb=ce("tbody"); rows.forEach(function(r){ var trr=ce("tr"); r.forEach(function(cell,ci){ trr.appendChild(ce("td",cols[ci].n?"n":null,cell)); }); tb.appendChild(trr); }); tbl.appendChild(tb); host.appendChild(tbl);
  }

  function renderChartInto(host,res,type){ clear(host); if(type==="kpi") renderKPI(host,res); else if(type==="line") renderLine(host,res); else if(type==="bar") renderBar(host,res); else if(type==="scatter") renderScatter(host,res); else renderTable(host,res); }

  // ---- pipeline: gather sources -> aggregate -> choose -> render ---------
  function neededSources(){ var set={}; S.measures.forEach(function(id){ var m=MBYID[id]; set[m.source==="status"?("status:"+m.diag):m.source]=1; }); return Object.keys(set); }
  function ensureProgress(host){ if(!host.querySelector(".progress")){ var p=ce("div","progress"); p.appendChild(ce("i")); host.insertBefore(p,host.firstChild); } return host.querySelector(".progress"); }
  function showProgress(on){ var host=$("chartHost"); var p=host.querySelector(".progress"); if(p) p.classList.toggle("on",!!on); }
  function emptyState(ic,big,sub,isErr){ var d=ce("div","empty"+(isErr?" err":"")); d.appendChild(ce("div","ic",ic)); d.appendChild(ce("div","big",big)); d.appendChild(ce("div",null,sub)); return d; }
  function showError(big,sub){ var host=$("chartHost"); clear(host); host.appendChild(emptyState("\u26A0",big,sub,true)); $("addTile").disabled=true; }
  function errMsg(e){ if(!e) return "unknown error"; if(typeof e==="string") return e; if(e.message) return e.message; if(e.name) return e.name; try{ return JSON.stringify(e).slice(0,160); }catch(x){ return "error"; } }
  function liveErr(e){ var m=errMsg(e); if(/permission|forbidden|clearance/i.test(m)) return "Your access may not cover this data. "+m; if(/timeout|time out/i.test(m)) return "The query took too long \u2014 try a shorter range or a single group."; return "Try a shorter date range or a smaller group. ("+m+")"; }
  function updateStatusRows(by){ var n=0; for(var k in by) n+=by[k].length; $("stRows").innerHTML="<b>"+fmtNum(n)+"</b> rows loaded"; $("stScope").innerHTML="<b>"+fmtNum(scopeDeviceIds().length)+"</b> vehicles in scope"; }

  var recompute=debounce(function(){ runPipeline(); },140);
  function runPipeline(){
    refreshChartTypeOptions(); syncTitle();
    var host=$("chartHost");
    if(S.measures.length===0){ clear(host); host.appendChild(emptyState("\u25E7","Build a chart","Add a measure to begin. Add more \u2014 or a slice \u2014 and the chart picks itself.")); $("addTile").disabled=true; return; }
    ensureProgress(host); showProgress(true);
    var srcs=neededSources();
    Promise.all(srcs.map(function(s){ return fetchRows(s); })).then(function(arrs){
      showProgress(false);
      var by={}; srcs.forEach(function(s,i){ by[s]=arrs[i]; });
      var res; try{ res=buildResult(by); }catch(e){ showError("Couldn't build the chart","Something went wrong while combining the data. "+(e&&e.message||"")); log("aggregate error: "+errMsg(e),"err"); return; }
      S.lastResult=res;
      var sp=shapeOf();
      var type=S.chartType==="auto"? autoChart(sp) : (validCharts(sp).indexOf(S.chartType)>=0? S.chartType : autoChart(sp));
      clear(host); ensureProgress(host); var area=ce("div"); host.appendChild(area);
      try{ renderChartInto(area,res,type); }catch(e2){ showError("Couldn't draw the chart","Rendering failed: "+(e2&&e2.message||"")); log("render error: "+errMsg(e2),"err"); return; }
      S.lastChartFn=function(h){ renderChartInto(h,res,type); };
      $("addTile").disabled=false; updateStatusRows(by);
    }).catch(function(e){ showProgress(false); showError("Couldn't load data", liveErr(e)); log("load error: "+errMsg(e),"err"); setMode(S.live?"live":"demo", S.live?"Connected":"Demo data"); });
  }

  function autoTitle(){
    if(S.measures.length===0) return "Untitled chart";
    var ms=S.measures.map(function(id){return MBYID[id].label;});
    var base = ms.length<=2? ms.join(" and ") : (ms.slice(0,2).join(", ")+" +"+(ms.length-2));
    if(S.dims.length){ base += " by " + S.dims.map(function(id){return DBYID[id].label.toLowerCase();}).join(" and "); }
    return base.charAt(0).toUpperCase()+base.slice(1);
  }
  function syncTitle(){ if(!S.titleEdited){ S.title=autoTitle(); $("chartTitle").value=S.title; } }

  function refreshChartTypeOptions(){
    var sp=shapeOf(); var sel=$("chartType"); var valid=validCharts(sp);
    clear(sel); var optA=ce("option",null,"Auto"+(sp.M?" \u2014 "+chartLabel(autoChart(sp)):"")); optA.value="auto"; sel.appendChild(optA);
    valid.forEach(function(t){ var o=ce("option",null,chartLabel(t)); o.value=t; sel.appendChild(o); });
    if(S.chartType!=="auto" && valid.indexOf(S.chartType)<0) S.chartType="auto"; sel.value=S.chartType;
  }

  // ---- build strip + catalog --------------------------------------------
  function addMeasure(id){ if(S.measures.indexOf(id)>=0) return; if(S.measures.length>=MAX_MEASURES){ toast("Up to "+MAX_MEASURES+" measures per chart."); return; } S.measures.push(id); S.titleEdited=false; renderStrip(); renderCatalog(); recompute(); }
  function addDim(id){ if(S.dims.indexOf(id)>=0) return; var d=DBYID[id]; if(d.time){ var ex=S.dims.filter(function(x){return DBYID[x].time;})[0]; if(ex){ S.dims=S.dims.filter(function(x){return x!==ex;}); } } if(S.dims.length>=MAX_DIMS){ toast("Up to "+MAX_DIMS+" slices. Remove one first."); return; } S.dims.push(id); S.titleEdited=false; renderStrip(); renderCatalog(); recompute(); }
  function removeItem(kind,id){ if(kind==="measure") S.measures=S.measures.filter(function(x){return x!==id;}); else S.dims=S.dims.filter(function(x){return x!==id;}); S.titleEdited=false; renderStrip(); renderCatalog(); recompute(); }

  function phEl(t){ return ce("span","ph",t); }
  function token(kind,id){ var item=kind==="measure"?MBYID[id]:DBYID[id]; var t=ce("span","tok "+kind); t.appendChild(document.createTextNode(item.label)); var x=ce("button","x","\u00D7"); x.setAttribute("aria-label","Remove "+item.label); x.onclick=function(){ removeItem(kind,id); }; t.appendChild(x); return t; }
  function renderStrip(){
    var zm=$("zMeasures"); clear(zm); if(S.measures.length===0){ zm.appendChild(phEl("Drag measures here, or click one on the left.")); } else S.measures.forEach(function(id){ zm.appendChild(token("measure",id)); });
    var zd=$("zDims"); clear(zd); if(S.dims.length===0){ zd.appendChild(phEl("Optional. Add up to two.")); } else S.dims.forEach(function(id){ zd.appendChild(token("dimension",id)); });
  }
  function catItem(item,kind,disabled){
    var b=ce("button","chip"); b.setAttribute("draggable","true"); b.dataset.kind=kind; b.dataset.id=item.id;
    b.appendChild(ce("span","mk "+kind)); b.appendChild(ce("span","nm",item.label)); if(kind==="measure") b.appendChild(ce("span","un",unitLabel(item)));
    if(disabled) b.setAttribute("aria-disabled","true");
    b.onclick=function(){ if(disabled) return; kind==="measure"?addMeasure(item.id):addDim(item.id); };
    b.addEventListener("dragstart",function(e){ e.dataTransfer.setData("text/plain",kind+":"+item.id); e.dataTransfer.effectAllowed="copy"; });
    return b;
  }
  function renderCatalog(){
    var host=$("catScroll"); clear(host); var q=($("catSearch").value||"").toLowerCase().trim();
    var cats={}; var order=[]; var extras=0;
    function push(cn,node){ if(!cats[cn]){ cats[cn]=[]; order.push(cn); } cats[cn].push(node); }
    MEASURES.forEach(function(m){
      if(q && m.label.toLowerCase().indexOf(q)<0 && m.cat.toLowerCase().indexOf(q)<0) return;
      if(S.live && S.dynLoaded){ if(m.source==="exception" && m.match && !m.ruleId) return; } else { if(m.dynamic) return; }
      if(m.cat==="All reported diagnostics"){ extras++; if(!S.showAllDiag && !q) return; }
      push(m.cat, catItem(m,"measure",S.measures.indexOf(m.id)>=0));
    });
    DIMENSIONS.forEach(function(d){ if(q && d.label.toLowerCase().indexOf(q)<0 && "break down by".indexOf(q)<0) return; push("Break down by", catItem(d,"dimension",S.dims.indexOf(d.id)>=0)); });
    order.forEach(function(cn){ var g=ce("div","cat-group"); g.appendChild(ce("h3",null,cn)); cats[cn].forEach(function(n){ g.appendChild(n); }); host.appendChild(g); });
    if(extras>0 && !q){ var t=ce("button","btn ghost"); t.style.cssText="margin:8px;font-size:12px;"; t.textContent=S.showAllDiag? "Hide extra diagnostics" : ("Show all "+extras+" reported diagnostics (incl. custom)"); t.onclick=function(){ S.showAllDiag=!S.showAllDiag; renderCatalog(); }; host.appendChild(t); }
    if(order.length===0){ var e=ce("div"); e.style.cssText="padding:18px 8px;color:var(--faint);font-size:12px;text-align:center;"; e.textContent="No data points match \u201C"+q+"\u201D."; host.appendChild(e); }
  }
  function wireDropzones(){
    ["zMeasures","zDims"].forEach(function(zid){ var z=$(zid); var want=z.dataset.kind;
      z.addEventListener("dragover",function(e){ e.preventDefault(); z.classList.add("hot"); e.dataTransfer.dropEffect="copy"; });
      z.addEventListener("dragleave",function(){ z.classList.remove("hot"); });
      z.addEventListener("drop",function(e){ e.preventDefault(); z.classList.remove("hot"); var data=e.dataTransfer.getData("text/plain")||""; var parts=data.split(":"); var kind=parts[0], id=parts[1]; if(!id) return; if(kind!==want){ toast(kind==="measure"?"That's a measure \u2014 drop it in Measures.":"That's a slice \u2014 drop it in Slice by."); return; } kind==="measure"?addMeasure(id):addDim(id); });
    });
  }

  // ---- group / vehicle multiselect popover ------------------------------
  function closePopovers(){ var p=$("activePop"); if(p) p.parentNode.removeChild(p); }
  function pickLabel(btn,sel,items,allText){ var lab=btn.querySelector(".lab"); if(!sel.length){ lab.textContent=allText; } else { var names=sel.map(function(id){ var f=items.filter(function(x){return x.id===id;})[0]; return f?f.name:id; }); lab.textContent = sel.length===1? names[0] : sel.length+" selected"; } }
  function openPicker(anchorBtn, items, selected, labelAll, onApply){
    closePopovers();
    var pop=ce("div","popover on"); pop.id="activePop";
    var search=ce("input","psearch"); search.placeholder="Search\u2026"; pop.appendChild(search);
    var list=ce("div","plist"); pop.appendChild(list); var local=selected.slice();
    function draw(){ clear(list); var q=(search.value||"").toLowerCase(); items.filter(function(it){ return it.name.toLowerCase().indexOf(q)>=0; }).slice(0,400).forEach(function(it){ var row=ce("label","opt"); var cb=ce("input"); cb.type="checkbox"; cb.checked=local.indexOf(it.id)>=0; cb.onchange=function(){ if(cb.checked){ if(local.indexOf(it.id)<0) local.push(it.id); } else local=local.filter(function(x){return x!==it.id;}); }; row.appendChild(cb); row.appendChild(document.createTextNode(it.name)); list.appendChild(row); }); }
    search.oninput=draw; draw();
    var foot=ce("div","pfoot"); var clr=ce("button",null,"Clear"); clr.onclick=function(){ local=[]; draw(); }; var apply=ce("button",null,"Apply"); apply.onclick=function(){ onApply(local); closePopovers(); }; foot.appendChild(clr); foot.appendChild(apply); pop.appendChild(foot);
    $("viz-app").appendChild(pop); var r=anchorBtn.getBoundingClientRect(); pop.style.left=r.left+"px"; pop.style.top=(r.bottom+6)+"px"; setTimeout(function(){ search.focus(); },10);
  }

  // ---- dashboard --------------------------------------------------------
  function emptyTiny(t){ var d=ce("div"); d.style.cssText="padding:24px;text-align:center;color:var(--faint);font-size:12px;"; d.textContent=t; return d; }
  function addTile(){
    if(!S.lastResult) return;
    S.tiles.push({ title:$("chartTitle").value||autoTitle(), measures:S.measures.slice(), dims:S.dims.slice(), chartType:S.chartType, days:S.days, customFrom:S.customFrom, customTo:S.customTo, groupSel:S.groupSel.slice(), vehSel:S.vehSel.slice(), units:S.units });
    toast("Added to dashboard ("+S.tiles.length+")"); updateTabs(); if(S.tab==="dash") renderDashboard();
  }
  function renderDashboard(){
    $("dashName").textContent=S.dashName; $("dashCount").textContent=S.tiles.length?("\u00B7 "+S.tiles.length+" chart"+(S.tiles.length>1?"s":"")):"";
    var grid=$("dashGrid"); clear(grid);
    if(S.tiles.length===0){ var e=emptyState("\u25A6","No charts yet","Build a chart, then choose \u201CAdd to dashboard\u201D. Save the layout to share it."); e.classList.add("dash-empty"); grid.appendChild(e); return; }
    S.tiles.forEach(function(tile,idx){
      var card=ce("div","dash-tile"); var th=ce("div","th"); th.appendChild(ce("div","tt",tile.title));
      var del=ce("button","btn ghost","Remove"); var armed=false; del.onclick=function(){ if(!armed){ armed=true; del.textContent="Confirm"; del.classList.add("danger"); setTimeout(function(){ if(armed){ armed=false; del.textContent="Remove"; del.classList.remove("danger"); } },2500); return; } S.tiles.splice(idx,1); updateTabs(); renderDashboard(); };
      th.appendChild(del); card.appendChild(th); var body=ce("div","body"); card.appendChild(body); grid.appendChild(card); renderTileBody(body,tile);
    });
  }
  function renderTileBody(body,tile){
    body.appendChild(emptyTiny("Loading\u2026"));
    var saved={days:S.days,customFrom:S.customFrom,customTo:S.customTo,groupSel:S.groupSel,vehSel:S.vehSel,units:S.units,measures:S.measures,dims:S.dims};
    S.days=tile.days; S.customFrom=tile.customFrom; S.customTo=tile.customTo; S.groupSel=tile.groupSel; S.vehSel=tile.vehSel; S.units=tile.units;
    var ms=tile.measures, ds=tile.dims; var srcs={}; ms.forEach(function(id){ srcs[MBYID[id].source]=1; }); var srcList=Object.keys(srcs);
    Promise.all(srcList.map(function(s){ return fetchRows(s); })).then(function(arrs){
      var by={}; srcList.forEach(function(s,i){ by[s]=arrs[i]; });
      S.measures=ms; S.dims=ds;
      var res=buildResult(by);
      var sp={M:ms.length,D:ds.length,hasTime:ds.some(function(id){return DBYID[id].time;}),catDims:ds.filter(function(id){return !DBYID[id].time;})};
      var type=tile.chartType==="auto"?autoChart(sp):tile.chartType;
      clear(body); renderChartInto(body,res,type);
    }).catch(function(e){ clear(body); body.appendChild(emptyTiny("Couldn't load")); log("tile load error: "+errMsg(e),"err"); })
    .then(function(){ S.days=saved.days; S.customFrom=saved.customFrom; S.customTo=saved.customTo; S.groupSel=saved.groupSel; S.vehSel=saved.vehSel; S.units=saved.units; S.measures=saved.measures; S.dims=saved.dims; });
  }

  // ---- save / load (AddInData live, localStorage demo) -------------------
  function dashPayload(){ return { name:S.dashName, tiles:S.tiles, savedAt:new Date().toISOString() }; }
  function closeMenu(){ $("savedMenu").classList.remove("on"); }
  function saveDashboard(name, shared){
    S.dashName=name||S.dashName; var payload=dashPayload(); payload.name=S.dashName; var json=JSON.stringify(payload);
    if(json.length>9800){ toast("Dashboard too large to save \u2014 remove a chart."); return; }
    if(!S.live){ try{ var all=JSON.parse(localStorage.getItem("vizDashboards")||"[]"); all=all.filter(function(d){return d.name!==S.dashName;}); all.push(payload); localStorage.setItem("vizDashboards",JSON.stringify(all)); }catch(e){} toast("Saved (demo, this browser)"); loadSavedList(); return; }
    var entity={ addInId:ADDIN_ID, details:json };
    if(shared){ var topGroup=S.meta.groups && S.meta.groups[0]; if(topGroup) entity.groups=[{id:topGroup.id}]; }
    gcall("Add",{ typeName:"AddInData", entity:entity }).then(function(){ toast(shared?"Saved and shared with your group":"Saved"); loadSavedList(); }).catch(function(e){ toast("Save failed"); log("AddInData save error: "+errMsg(e),"err"); });
  }
  function loadSavedList(){
    if(!S.live){ try{ S.savedList=JSON.parse(localStorage.getItem("vizDashboards")||"[]").map(function(d,i){ return {id:"local"+i, name:d.name, payload:d}; }); }catch(e){ S.savedList=[]; } renderSavedMenu(); return; }
    gcall("Get",{ typeName:"AddInData", search:{ addInId:ADDIN_ID } }).then(function(rows){ S.savedList=(rows||[]).map(function(r){ var p; try{ p=JSON.parse(r.details); }catch(e){ p=null; } return p?{id:r.id, name:p.name||"Untitled", payload:p}:null; }).filter(Boolean); renderSavedMenu(); }).catch(function(e){ log("AddInData get error: "+errMsg(e),"err"); });
  }
  function renderSavedMenu(){
    var m=$("savedMenu"); clear(m);
    var sr=ce("div","saverow"); var inp=ce("input"); inp.placeholder="Dashboard name"; inp.value=S.dashName; var sv=ce("button","btn primary","Save"); sv.onclick=function(){ saveDashboard(inp.value, S.live); }; sr.appendChild(inp); sr.appendChild(sv); m.appendChild(sr); m.appendChild(ce("div","seph"));
    if(S.savedList.length===0){ m.appendChild(ce("div","empty-m","No saved dashboards yet.")); return; }
    S.savedList.forEach(function(item){ var row=ce("div","row"); row.appendChild(ce("div","nm",item.name));
      var open=ce("button","btn ghost","Open"); open.onclick=function(){ S.tiles=(item.payload.tiles||[]).slice(); S.dashName=item.payload.name||"Dashboard"; closeMenu(); switchTab("dash"); toast("Opened "+S.dashName); }; row.appendChild(open);
      var del=ce("button","btn ghost","Delete"); var armed=false; del.onclick=function(ev){ ev.stopPropagation(); if(!armed){ armed=true; del.textContent="Confirm"; del.classList.add("danger"); setTimeout(function(){ armed=false; del.textContent="Delete"; del.classList.remove("danger"); },2500); return; } deleteDashboard(item); }; row.appendChild(del); m.appendChild(row); });
  }
  function deleteDashboard(item){
    if(!S.live){ try{ var all=JSON.parse(localStorage.getItem("vizDashboards")||"[]").filter(function(d){return d.name!==item.name;}); localStorage.setItem("vizDashboards",JSON.stringify(all)); }catch(e){} loadSavedList(); toast("Deleted"); return; }
    gcall("Remove",{ typeName:"AddInData", entity:{ id:item.id, addInId:ADDIN_ID } }).then(function(){ toast("Deleted"); loadSavedList(); }).catch(function(e){ toast("Delete failed"); log("AddInData remove error: "+errMsg(e),"err"); });
  }

  // ---- tabs / mode ------------------------------------------------------
  function updateTabs(){ $("tabDash").textContent="Dashboard"+(S.tiles.length?(" ("+S.tiles.length+")"):""); }
  function switchTab(which){ S.tab=which; $("tabBuild").setAttribute("aria-selected", which==="build"?"true":"false"); $("tabDash").setAttribute("aria-selected", which==="dash"?"true":"false"); $("vizBuildView").hidden = which!=="build"; $("vizDashView").hidden = which!=="dash"; if(which==="dash"){ renderDashboard(); loadSavedList(); } }
  function setMode(kind,text){ var el=$("vizMode"); var dot=el.querySelector(".viz-dot"); dot.className="viz-dot"+(kind==="live"?" live":kind==="demo"?" demo":""); if(el.lastChild && el.lastChild.nodeType===3) el.removeChild(el.lastChild); el.appendChild(document.createTextNode(" "+text)); }

  // ---- controls wiring --------------------------------------------------
  function invalidateCacheAndRecompute(){ S.cache={}; recompute(); if(S.tab==="dash") renderDashboard(); }
  function wireControls(){
    $("rangeSeg").addEventListener("click",function(e){ var b=e.target.closest("button"); if(!b) return; var v=b.dataset.days; Array.prototype.forEach.call(this.children,function(c){ c.setAttribute("aria-pressed", c===b?"true":"false"); }); $("customDates").classList.toggle("on", v==="custom"); S.days = v==="custom"?"custom":(+v); if(v==="custom" && !S.customFrom){ var t=new Date(),f=new Date(); f.setDate(f.getDate()-30); $("dFrom").value=localDayKey(f); $("dTo").value=localDayKey(t); S.customFrom=$("dFrom").value; S.customTo=$("dTo").value; } invalidateCacheAndRecompute(); });
    $("dFrom").addEventListener("change",function(){ S.customFrom=this.value; invalidateCacheAndRecompute(); });
    $("dTo").addEventListener("change",function(){ S.customTo=this.value; invalidateCacheAndRecompute(); });
    $("unitSeg").addEventListener("click",function(e){ var b=e.target.closest("button"); if(!b) return; Array.prototype.forEach.call(this.children,function(c){ c.setAttribute("aria-pressed", c===b?"true":"false"); }); S.units=b.dataset.unit; renderCatalog(); recompute(); if(S.tab==="dash") renderDashboard(); });
    $("groupPick").addEventListener("click",function(){ openPicker(this, S.meta.groups||[], S.groupSel, "All groups", function(sel){ S.groupSel=sel; pickLabel($("groupPick"),sel,S.meta.groups||[],"All groups"); invalidateCacheAndRecompute(); }); });
    $("vehPick").addEventListener("click",function(){ openPicker(this, S.meta.devices||[], S.vehSel, "All vehicles", function(sel){ S.vehSel=sel; pickLabel($("vehPick"),sel,S.meta.devices||[],"All vehicles"); invalidateCacheAndRecompute(); }); });
    $("catSearch").addEventListener("input",debounce(renderCatalog,80));
    $("chartType").addEventListener("change",function(){ S.chartType=this.value; recompute(); });
    $("chartTitle").addEventListener("input",function(){ S.title=this.value; S.titleEdited=true; });
    $("addTile").addEventListener("click",addTile);
    $("tabBuild").addEventListener("click",function(){ switchTab("build"); });
    $("tabDash").addEventListener("click",function(){ switchTab("dash"); });
    $("clearDash").addEventListener("click",function(){ if(!S.tiles.length) return; var btn=this; if(btn.dataset.armed!=="1"){ btn.dataset.armed="1"; btn.textContent="Confirm clear"; setTimeout(function(){ btn.dataset.armed="0"; btn.textContent="Clear"; },2500); return; } S.tiles=[]; btn.dataset.armed="0"; btn.textContent="Clear"; updateTabs(); renderDashboard(); });
    $("savedBtn").addEventListener("click",function(e){ e.stopPropagation(); $("savedMenu").classList.toggle("on"); if($("savedMenu").classList.contains("on")) loadSavedList(); });
    $("logToggle").addEventListener("click",function(){ var p=$("logPanel"); p.classList.toggle("on"); this.textContent=(p.classList.contains("on")?"Data log \u25B4":"Data log \u25BE"); });
    $("catToggle").addEventListener("click",function(){ $("vizCatalog").classList.toggle("open"); });
    document.addEventListener("click",function(e){ if(!e.target.closest("#savedBtn")&&!e.target.closest("#savedMenu")) closeMenu(); if(!e.target.closest(".popover")&&!e.target.closest(".pick-btn")) closePopovers(); });
    window.addEventListener("resize",debounce(function(){ if(S.lastChartFn){ var host=$("chartHost"); clear(host); ensureProgress(host); var a=ce("div"); host.appendChild(a); try{ S.lastChartFn(a); }catch(e){} } if(S.tab==="dash") renderDashboard(); },180));
  }

  // ---- metadata ---------------------------------------------------------
  function loadMetadata(){
    if(!S.live){ applyMeta(mockMeta()); return Promise.resolve(); }
    setMode("live","Loading fleet\u2026");
    return Promise.all([
      gcall("Get",{typeName:"Device", search:{ fromDate:new Date().toISOString() }, resultsLimit:50000}),
      gcall("Get",{typeName:"Group", resultsLimit:5000}),
      gcall("Get",{typeName:"Rule", resultsLimit:5000}),
      gcall("Get",{typeName:"User", search:{ isDriver:true }, resultsLimit:50000})
    ]).then(function(r){ var devices=r[0]||[], groups=r[1]||[], rules=r[2]||[], drivers=r[3]||[]; applyMeta({devices:devices,groups:groups,rules:rules,drivers:drivers}); log("metadata: "+devices.length+" vehicles, "+groups.length+" groups, "+rules.length+" rules, "+drivers.length+" drivers","ok"); setMode("live","Connected"); })
    .catch(function(e){ log("metadata load failed: "+errMsg(e)+" \u2014 falling back to demo data","err"); toast("Couldn't reach MyGeotab \u2014 showing demo data"); S.live=false; applyMeta(mockMeta()); setMode("demo","Demo data (load failed)"); });
  }
  function applyMeta(mk){
    S.meta.devices=mk.devices.map(function(d){ return { id:d.id, name:d.name||d.serialNumber||d.id, groups:d.groups||[], __driver:d.__driver }; });
    S.meta.deviceById={}; S.meta.devices.forEach(function(d){ S.meta.deviceById[d.id]=d; });
    S.meta.groups=mk.groups.map(function(g){ return { id:g.id, name:(g.name||g.id), isRoot:(g.id==="GroupCompanyId"||/entire organization|company group/i.test(g.name||"")) }; });
    S.meta.groupById={}; S.meta.groups.forEach(function(g){ S.meta.groupById[g.id]=g; });
    S.meta.driverById={}; (mk.drivers||[]).forEach(function(u){ S.meta.driverById[u.id]={ id:u.id, name:((u.firstName||u.lastName)?((u.firstName||"")+" "+(u.lastName||"")).trim():(u.name||u.id)) }; });
    S.meta.rules=mk.rules||[]; matchRules();
    pickLabel($("groupPick"),S.groupSel,S.meta.groups,"All groups"); pickLabel($("vehPick"),S.vehSel,S.meta.devices,"All vehicles");
    $("stScope").innerHTML="<b>"+fmtNum(S.meta.devices.length)+"</b> vehicles in scope";
  }
  function matchRules(){
    S.meta.matchedRuleIds={};
    MEASURES.forEach(function(m){ if(m.source!=="exception"||m.matchAll) return; var ids={}; S.meta.rules.forEach(function(r){ var nm=(r.name||"").toLowerCase(); for(var i=0;i<m.match.length;i++){ if(nm.indexOf(m.match[i])>=0){ ids[r.id]=1; break; } } }); S.meta.matchedRuleIds[m.id]=ids; });
    var seen={}; for(var mid in S.meta.matchedRuleIds){ for(var rid in S.meta.matchedRuleIds[mid]) seen[rid]=1; }
    var matchedCount=Object.keys(seen).length; $("stRules").innerHTML="<b>"+matchedCount+"</b> safety rules matched";
    if(S.live && matchedCount===0) log("no rules matched safety measures by name \u2014 safety event metrics will read 0. Check rule names in this database.","err");
  }

  // ---- runtime discovery: surface every rule + diagnostic this fleet has --
  function guessVtype(name){ var n=(name||"").toLowerCase();
    if(/voltage|volts?\b/.test(n)) return "volts";
    if(/temperature|temp\b/.test(n)) return "tempC";
    if(/pressure/.test(n)) return "pressurePa";
    if(/\brpm\b|engine speed/.test(n)) return "rpm";
    if(/odometer|distance/.test(n)) return "distM";
    if(/road speed|\bspeed\b/.test(n)) return "speedKph";
    if(/hours\b/.test(n)) return "hoursSec";
    if(/percent|state of charge|level \(perc/.test(n)) return "pct";
    if(/fuel|volume|litre|liter|gallon/.test(n)) return "volumeL";
    return "num";
  }
  function guessAgg(name){ var n=(name||"").toLowerCase();
    if(/odometer|engine hours|\bhours\b|\btotal\b|counts?\b|distance|fuel used/.test(n)) return "last";
    return "avg";
  }
  function discoverCatalog(){
    if(!S.live || S.dynLoaded) return Promise.resolve();
    setMode("live","Discovering data points\u2026");
    try{ (S.meta.rules||[]).forEach(function(rl){ if(!rl||!rl.id) return; var mid="rule_"+rl.id; if(MBYID[mid]) return; var m={ id:mid, label:(rl.name||rl.id), cat:"Rules & exceptions", source:"exception", agg:"count", unit:"events", vtype:"int", ruleId:rl.id, dynamic:true }; MEASURES.push(m); MBYID[mid]=m; }); }catch(e){ log("rule discovery: "+errMsg(e),"err"); }
    var now=new Date(); var from=new Date(now.getTime()-3*24*3600*1000);
    var ids=S.meta.devices.map(function(d){return d.id;}).slice(0,12);
    if(!ids.length){ finishDiscover(); return Promise.resolve(); }
    var sdCalls=ids.map(function(id){ return ["Get",{ typeName:"StatusData", search:{ fromDate:from.toISOString(), toDate:now.toISOString(), deviceSearch:{id:id} }, resultsLimit:3000 }]; });
    var curated={}; MEASURES.forEach(function(m){ if(m.source==="status"&&m.diag) curated[m.diag]=1; });
    return gmulti(sdCalls).then(function(res){
      var seen={}; res.forEach(function(a){ if(a&&a.length) a.forEach(function(r){ var d=r.diagnostic&&r.diagnostic.id; if(d && !seen[d]) seen[d]={id:d,val:r.data}; }); });
      var newIds=Object.keys(seen).filter(function(d){ return !curated[d] && !MBYID["diag_"+d]; }).slice(0,250);
      if(!newIds.length) return;
      var dCalls=newIds.map(function(d){ return ["Get",{ typeName:"Diagnostic", search:{ id:d } }]; });
      return gmulti(dCalls).then(function(dres){
        var names={}; dres.forEach(function(a){ if(a&&a[0]) names[a[0].id]=a[0].name; });
        newIds.forEach(function(d){ var nm=names[d]||d; var mid="diag_"+d; if(MBYID[mid]) return; var m={ id:mid, label:nm, cat:"All reported diagnostics", source:"status", diag:d, agg:guessAgg(nm), unit:"", vtype:guessVtype(nm), dynamic:true }; MEASURES.push(m); MBYID[mid]=m; });
      });
    }).then(finishDiscover).catch(function(e){ log("diagnostic discovery: "+errMsg(e),"err"); finishDiscover(); });
  }
  function finishDiscover(){ S.dynLoaded=true; var extra=0,rules=0; MEASURES.forEach(function(m){ if(m.dynamic){ if(m.source==="status")extra++; else rules++; } }); log("discovery: "+rules+" rules and "+extra+" additional diagnostics found (incl. any custom)","ok"); renderCatalog(); setMode("live","Connected"); }

  // ---- demo data --------------------------------------------------------
  function mulberry32(a){ return function(){ a|=0; a=a+0x6D2B79F5|0; var t=Math.imul(a^a>>>15,1|a); t=t+Math.imul(t^t>>>7,61|t)^t; return ((t^t>>>14)>>>0)/4294967296; }; }
  function hashStr(s){ var h=2166136261; for(var i=0;i<s.length;i++){ h^=s.charCodeAt(i); h=Math.imul(h,16777619); } return h>>>0; }
  function secToHMS(sec){ sec=Math.max(0,Math.round(sec)); var h=Math.floor(sec/3600), m=Math.floor((sec%3600)/60), s=sec%60; return pad2(h)+":"+pad2(m)+":"+pad2(s); }
  function poisson(rnd,lambda){ if(lambda<=0) return 0; var L=Math.exp(-lambda), k=0, p=1; do{ k++; p*=rnd(); }while(p>L); return k-1; }
  var RULE_IDS={ harsh_braking:"r_brake", harsh_accel:"r_accel", harsh_corner:"r_corner", speeding:"r_speed", idle_events:"r_idle", seatbelt:"r_seat" };
  function mockMeta(){
    var groups=[ {id:"GroupCompanyId",name:"Company"}, {id:"g_north",name:"North Region"}, {id:"g_south",name:"South Region"}, {id:"g_west",name:"West Region"} ];
    var regionIds=["g_north","g_south","g_west"];
    var rules=[ {id:"r_brake",name:"Harsh braking"},{id:"r_accel",name:"Harsh acceleration"},{id:"r_corner",name:"Harsh cornering"},{id:"r_speed",name:"Speeding"},{id:"r_idle",name:"Idling"},{id:"r_seat",name:"Seatbelt"} ];
    var dn=["Alex Rivera","Jordan Lee","Sam Patel","Casey Brooks","Morgan Diaz","Taylor Nguyen"]; var drivers=[]; for(var i=0;i<6;i++){ var parts=dn[i].split(" "); drivers.push({id:"driver_"+i, firstName:parts[0], lastName:parts[1]}); }
    var devices=[]; for(var v=1;v<=14;v++){ var rg=regionIds[v%3]; devices.push({ id:"dev_"+v, name:"Truck "+(v<10?"0":"")+v, groups:[{id:"GroupCompanyId"},{id:rg}], __driver:"driver_"+(v%6) }); }
    return { groups:groups, rules:rules, drivers:drivers, devices:devices };
  }
  function mockDiagValue(diag,rnd,di,n){
    function around(x,p){ return x*(1-p+rnd()*2*p); }
    var f=di/Math.max(1,n);
    if(/Voltage/i.test(diag)) return around(12.4,0.12);
    if(/CoolantTemperature/i.test(diag)) return around(85,0.15);
    if(/OutsideTemperature/i.test(diag)) return around(18,0.5);
    if(/EngineSpeed/i.test(diag)) return around(1500,0.4);
    if(/RoadSpeed/i.test(diag)) return around(60,0.6);
    if(/FuelLevelId/i.test(diag)) return around(55,0.5);
    if(/FuelUnits/i.test(diag)) return around(60,0.35);
    if(/TirePressure/i.test(diag)) return around(275000,0.06);
    if(/PredictedRiskOfBreakdown/i.test(diag)) return Math.max(0,around(12,1));
    if(/ElectricalSystemRating/i.test(diag)) return around(80,0.12);
    if(/Acceleration/i.test(diag)) return (rnd()-0.5)*6;
    if(/Odometer/i.test(diag)) return 35000000 + f*1500000 + rnd()*3000;
    if(/EngineHours/i.test(diag)) return 3200000 + f*80000;
    if(/TotalFuelUsed|DeviceTotalFuel/i.test(diag)) return 4000 + f*600 + rnd()*4;
    if(/IdleFuel/i.test(diag)) return 70 + f*20;
    return around(50,0.4);
  }
  function mockRows(source,r,ids){
    var out=[]; var days=eachDay(r.from,r.to); var devs=ids.length?ids:S.meta.devices.map(function(d){return d.id;});
    if(source.indexOf("status:")===0){ var diag=source.slice(7); devs.forEach(function(id){ var rnd=mulberry32(hashStr(id+diag)); days.forEach(function(dayKey,di){ var dt=new Date(dayKey+"T09:00:00"); var n=3+Math.floor(rnd()*4); for(var i=0;i<n;i++){ var when=new Date(dt.getTime()+i*2*3600*1000); out.push({ device:{id:id}, dateTime:when.toISOString(), data:mockDiagValue(diag,rnd,di,days.length) }); } }); }); return out; }
    if(source==="fault"){ devs.forEach(function(id){ var rnd=mulberry32(hashStr(id+"flt")); days.forEach(function(dayKey){ if(rnd()<0.07){ var dt=new Date(dayKey+"T12:00:00"); var isBatt=rnd()<0.5; out.push({ device:{id:id}, dateTime:dt.toISOString(), diagnostic:{id:isBatt?"DiagnosticVehicleBatteryLowVoltageId":"DiagnosticAccidentLevelAccelerationEventId"}, malfunctionLamp:rnd()<0.3, count:1 }); } }); }); return out; }
    devs.forEach(function(id){
      var seed=hashStr(id), rnd=mulberry32(seed); var dev=S.meta.deviceById[id]||{}; var driverId=dev.__driver||("driver_"+(seed%6));
      var risk=0.5+rnd()*1.6, activity=0.6+rnd()*1.3, avgSpeedBase=40+rnd()*22;
      days.forEach(function(dayKey,di){
        var dt=new Date(dayKey+"T08:00:00"); var dow=dt.getDay(); var wk=(dow===0||dow===6)?0.35:1; var trend=1+(di/Math.max(1,days.length))*0.25;
        if(source==="trip"){
          var nTrips=Math.max(0,Math.round(activity*wk*(1.4+rnd()*3)));
          for(var t=0;t<nTrips;t++){ var dist=4+rnd()*48; var avgSp=avgSpeedBase*(0.8+rnd()*0.4); var driveH=dist/Math.max(15,avgSp); var idleH=driveH*(0.05+rnd()*0.3); var stopH=rnd()*1.2; var maxSp=avgSp+10+rnd()*30; var start=new Date(dt.getTime()+(t*1.3+rnd())*3600*1000);
            out.push({ device:{id:id}, driver:{id:driverId}, start:start.toISOString(), distance:dist, drivingDuration:secToHMS(driveH*3600), idlingDuration:secToHMS(idleH*3600), stopDuration:secToHMS(stopH*3600), averageSpeed:avgSp, maximumSpeed:maxSp }); }
        } else {
          var bases={ harsh_braking:1.0, harsh_accel:0.8, harsh_corner:0.6, speeding:1.1, idle_events:1.4, seatbelt:0.3 };
          Object.keys(RULE_IDS).forEach(function(mid){ var lambda=risk*wk*trend*bases[mid]; var n=poisson(rnd,lambda); for(var e=0;e<n;e++){ var when=new Date(dt.getTime()+rnd()*10*3600*1000); out.push({ device:{id:id}, driver:{id:driverId}, rule:{id:RULE_IDS[mid]}, activeFrom:when.toISOString(), duration:secToHMS(2+rnd()*8), distance:rnd()*0.3 }); } });
        }
      });
    });
    return out;
  }

  // ---- discovery probe (temporary: run once to inventory the live database) ----
  function injectDiscoverButton(){
    var bar=document.querySelector("#viz-app .status-bar"); if(!bar||document.getElementById("discoverBtn")) return;
    var b=ce("button","toglog"); b.id="discoverBtn"; b.textContent="Discover data points"; b.style.marginRight="10px";
    b.onclick=function(){ runDiscovery(b); };
    bar.insertBefore(b, $("logToggle"));
  }
  function runDiscovery(btn){
    if(!S.live){ toast("Discovery needs a live MyGeotab connection."); return; }
    btn.disabled=true; var label=btn.textContent; btn.textContent="Discovering\u2026";
    var now=new Date(); var dayAgo=new Date(now.getTime()-24*3600*1000); var weekAgo=new Date(now.getTime()-7*24*3600*1000);
    var ids=scopeDeviceIds(); var sampleDevs=ids.slice(0,8);
    var report={ generatedAt:now.toISOString(), scope:{ vehiclesInScope:ids.length, sampledVehicles:sampleDevs.length }, probes:{} };
    function isoBack(days){ return new Date(now.getTime()-days*24*3600*1000).toISOString(); }
    function probe(name, promise, shaper){ return promise.then(function(rows){ report.probes[name]={ ok:true, count:(rows&&rows.length)||0, sample:(shaper?shaper(rows):(rows||[]).slice(0,2)) }; }).catch(function(e){ report.probes[name]={ ok:false, error:errMsg(e) }; }); }
    function sample1(typeName, key, days){ var s={ fromDate:isoBack(days||30), toDate:now.toISOString() }; if(sampleDevs[0]) s.deviceSearch={id:sampleDevs[0]}; return probe(key||typeName, gcall("Get",{ typeName:typeName, search:s, resultsLimit:2 })); }
    var tasks=[];

    if(sampleDevs.length){
      var sdCalls=sampleDevs.map(function(id){ return ["Get",{ typeName:"StatusData", search:{ fromDate:dayAgo.toISOString(), toDate:now.toISOString(), deviceSearch:{id:id} }, resultsLimit:2000 }]; });
      tasks.push( gmulti(sdCalls).then(function(res){
        var rows=[]; res.forEach(function(a){ if(a&&a.length) rows=rows.concat(a); });
        var byDiag={}; rows.forEach(function(r){ var d=r.diagnostic&&r.diagnostic.id; if(!d) return; if(!byDiag[d]) byDiag[d]={ id:d, count:0, sampleValue:r.data, sampleAt:r.dateTime }; byDiag[d].count++; });
        var list=Object.keys(byDiag).map(function(k){ return byDiag[k]; });
        report.probes.statusData={ ok:true, rowsSampled:rows.length, distinctDiagnostics:list.length, diagnostics:list };
        var dCalls=list.slice(0,150).map(function(d){ return ["Get",{ typeName:"Diagnostic", search:{ id:d.id } }]; });
        return gmulti(dCalls).then(function(dres){ var info={}; dres.forEach(function(a){ if(a&&a[0]){ var d=a[0]; info[d.id]={ name:d.name, unitId:(d.unitOfMeasure&&d.unitOfMeasure.id)||null, type:d.diagnosticType||null }; } }); list.forEach(function(dg){ var i=info[dg.id]; if(i){ dg.name=i.name; dg.unitId=i.unitId; dg.type=i.type; } }); }).catch(function(e){ report.probes.statusData.nameResolveError=errMsg(e); });
      }).catch(function(e){ report.probes.statusData={ ok:false, error:errMsg(e) }; }) );

      var fdCalls=sampleDevs.map(function(id){ return ["Get",{ typeName:"FaultData", search:{ fromDate:weekAgo.toISOString(), toDate:now.toISOString(), deviceSearch:{id:id} }, resultsLimit:500 }]; });
      tasks.push( gmulti(fdCalls).then(function(res){ var rows=[]; res.forEach(function(a){ if(a&&a.length) rows=rows.concat(a); }); report.probes.faultData={ ok:true, rowsSampled:rows.length, sample:rows.slice(0,3) }; }).catch(function(e){ report.probes.faultData={ ok:false, error:errMsg(e) }; }) );
    } else {
      report.probes.statusData={ ok:false, error:"no vehicles in scope" }; report.probes.faultData={ ok:false, error:"no vehicles in scope" };
    }

    tasks.push( probe("unitsOfMeasure", gcall("Get",{ typeName:"UnitOfMeasure", resultsLimit:500 }), function(rows){ return (rows||[]).map(function(u){ return { id:u.id, name:u.name }; }); }) );
    tasks.push( sample1("Trip","trip",30) );
    tasks.push( sample1("ExceptionEvent","exceptionEvent",30) );
    tasks.push( sample1("LogRecord","logRecord_gps",2) );
    tasks.push( sample1("DVIRLog","dvir",90) );
    tasks.push( sample1("DutyStatusLog","hos",14) );
    tasks.push( probe("zones", gcall("Get",{ typeName:"Zone", resultsLimit:3 })) );
    tasks.push( sample1("FuelUsed","fuelUsed",30) );
    tasks.push( sample1("FuelTaxDetail","ifta",30) );
    tasks.push( sample1("FuelTransaction","fuelTransaction",60) );
    tasks.push( sample1("ChargeEvent","ev_chargeEvent",90) );

    Promise.all(tasks).then(function(){ finishDiscovery(btn,label,report); }).catch(function(){ finishDiscovery(btn,label,report); });
  }
  function finishDiscovery(btn,label,report){ btn.disabled=false; btn.textContent=label; log("discovery complete \u2014 copy the inventory and paste it back to Claude","ok"); showJsonModal(report); }
  function showJsonModal(obj){
    var ov=ce("div"); ov.style.cssText="position:fixed;inset:0;background:rgba(16,20,40,.45);z-index:200;display:flex;align-items:center;justify-content:center;padding:24px;";
    var box=ce("div"); box.style.cssText="background:#fff;border-radius:12px;max-width:860px;width:100%;max-height:84vh;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 20px 60px rgba(16,20,40,.4);";
    var head=ce("div"); head.style.cssText="padding:14px 18px;border-bottom:1px solid #E6E8EC;display:flex;align-items:center;gap:12px;";
    var h=ce("div",null,"Data-point inventory"); h.style.cssText="font-weight:700;font-size:15px;flex:1 1 auto;color:#1F1446;";
    var copy=ce("button","btn primary","Copy"); var close=ce("button","btn","Close"); head.appendChild(h); head.appendChild(copy); head.appendChild(close); box.appendChild(head);
    var sub=ce("div",null,"Copy this and paste it back in the chat. It is the catalog of data points your fleet reports, plus one or two sample rows per type so the field names are exact \u2014 scrub anything you would rather not share."); sub.style.cssText="padding:9px 18px 0;color:#6B7280;font-size:12px;line-height:1.5;"; box.appendChild(sub);
    var ta=document.createElement("textarea"); ta.value=JSON.stringify(obj,null,2); ta.readOnly=true; ta.style.cssText="flex:1 1 auto;margin:12px 18px 18px;min-height:320px;border:1px solid #E6E8EC;border-radius:8px;padding:12px;font-family:ui-monospace,Menlo,monospace;font-size:11.5px;resize:none;white-space:pre;overflow:auto;color:#1F1446;"; box.appendChild(ta);
    copy.onclick=function(){ ta.focus(); ta.select(); var done=false; try{ if(navigator.clipboard&&navigator.clipboard.writeText){ navigator.clipboard.writeText(ta.value); done=true; } }catch(e){} if(!done){ try{ document.execCommand("copy"); }catch(e2){} } copy.textContent="Copied"; setTimeout(function(){ copy.textContent="Copy"; },1500); };
    close.onclick=function(){ if(ov.parentNode) ov.parentNode.removeChild(ov); };
    ov.addEventListener("click",function(e){ if(e.target===ov && ov.parentNode) ov.parentNode.removeChild(ov); });
    ov.appendChild(box); document.getElementById("viz-app").appendChild(ov);
  }

  // ---- boot -------------------------------------------------------------
  function boot(api,live){
    S.api=api; S.live=live;
    renderStrip(); wireDropzones(); wireControls(); renderCatalog(); updateTabs(); injectDiscoverButton();
    setMode(live?"live":"demo", live?"Connecting\u2026":"Demo data");
    loadMetadata().then(function(){ S.ready=true; recompute(); discoverCatalog(); });
  }
  if (typeof geotab!=="undefined" && geotab.addin){
    geotab.addin.insightBuilder=function(){ return { initialize:function(api,state,cb){ boot(api,true); if(cb) cb(); }, focus:function(){}, blur:function(){} }; };
  } else {
    if(document.readyState==="loading"){ document.addEventListener("DOMContentLoaded",function(){ boot(null,false); }); } else { boot(null,false); }
  }

})();
