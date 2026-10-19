// Test-only. Routes Arbor's GitHub calls to the local fake Gist, tagged with
// which "device" (origin) made them; can background the page the instant boot
// starts fetching, which is what a phone does when you switch away mid-launch;
// and carries a soak driver that survives relaunches via localStorage.
(function(){
  const dev=location.hostname==="localhost"?"A":"B";
  const real=window.fetch.bind(window);
  window.__dev=dev;
  window.fetch=function(url,opts){
    if(typeof url==="string"&&url.startsWith("https://api.github.com")){
      const o=Object.assign({},opts||{}); o.headers=Object.assign({},o.headers||{},{"X-Device":dev});
      const path=url.slice("https://api.github.com".length);
      const isWorkspaceGet=(!o.method||o.method==="GET")&&/^\/gists\/[^/]+$/.test(path.split("?")[0]);
      const p=real(location.origin+"/fakegist"+path,o);
      if(isWorkspaceGet&&!window.__hidOnce&&localStorage.getItem("__hideOnBootFetch")==="1"){
        window.__hidOnce=true;
        Object.defineProperty(document,"visibilityState",{configurable:true,get:()=>"hidden"});
        Object.defineProperty(document,"hidden",{configurable:true,get:()=>true});
        document.dispatchEvent(new Event("visibilitychange"));
        window.__backgroundedDuringBoot=Date.now();
        p.then(()=>{ Object.defineProperty(document,"visibilityState",{configurable:true,get:()=>"visible"});
                     Object.defineProperty(document,"hidden",{configurable:true,get:()=>false}); },()=>{});
      }
      return p;
    }
    return real(url,opts);
  };
  Object.defineProperty(document,"visibilityState",{configurable:true,get:()=>"visible"});
  Object.defineProperty(document,"hidden",{configurable:true,get:()=>false});
  if(navigator.serviceWorker){ navigator.serviceWorker.register=()=>Promise.reject(new Error("off")); }

  window.__soak={
    log(){ try{ return JSON.parse(localStorage.getItem("__soak_log")||'{"ops":[]}'); }catch(e){ return {ops:[]}; } },
    save(l){ localStorage.setItem("__soak_log",JSON.stringify(l)); },
    sleep(ms){ return real(location.origin+"/fakegist/__sleep?ms="+ms).catch(()=>{}); },
    lastOps(L){ const m=new Map(); L.ops.forEach(o=>m.set(o.id,o)); return m; },
    find(coll,id){
      if(coll==="diary"){ for(const a of DB.animals) for(const e of (a.diary||[])) if(e.id===id) return {a,e}; return null; }
      return (DB[coll]||[]).find(x=>x.id===id)||null;
    },
    duringBoot(){
      const L=this.log(), id=uid("n_");
      (DB.notes=DB.notes||[]).push({id,title:"soak "+window.__dev+" during boot",body:"",created:Date.now(),links:[]});
      persist(); L.ops.push({op:"create",coll:"notes",id,duringBoot:true}); this.save(L); return id;
    },
    async step(n,seedStr){
      let s=0; for(const c of seedStr) s=(s*31+c.charCodeAt(0))>>>0;
      const rnd=()=>{ s=(s*1664525+1013904223)>>>0; return s/4294967296; };
      const pick=a=>a[Math.floor(rnd()*a.length)];
      const L=this.log(); const did=[];
      for(let i=0;i<n;i++){
        const last=this.lastOps(L), own=[...last.values()];
        const liveOwn=own.filter(o=>o.op!=="delete"&&this.find(o.coll,o.id));
        const deadOwn=own.filter(o=>o.op==="delete");
        const r=rnd();
        if(r<0.17){ const id=uid("n_"); DB.notes.push({id,title:"soak "+dev,body:"",created:Date.now(),links:[]}); persist(); L.ops.push({op:"create",coll:"notes",id}); did.push("note+"); }
        else if(r<0.30){ const id=uid("tk_"); taskList().push({id,title:"soak "+dev,when:null,created:Date.now()}); persist(); L.ops.push({op:"create",coll:"schedule",id}); did.push("task+"); }
        else if(r<0.43){ const a=pick(DB.animals); const id=uid("d_"); (a.diary=a.diary||[]).push({id,type:"comment",date:"2026-10-18",data:{note:"soak "+dev}}); persist(); L.ops.push({op:"create",coll:"diary",id,animal:a.id}); did.push("diary+"); }
        else if(r<0.53){ const xs=DB.notes.filter(x=>/^soak/.test(x.title||"")); if(xs.length){ const x=pick(xs); x.title="soak "+dev+" edited "+(Date.now()%100000); persist(); did.push("edit"); } }
        else if(r<0.68){ if(liveOwn.length){ const o=pick(liveOwn); let obj;
            if(o.coll==="diary"){ const f=this.find("diary",o.id); obj=f.e; f.a.diary=f.a.diary.filter(e=>e.id!==o.id); o.animal=f.a.id; }
            else { obj=this.find(o.coll,o.id); DB[o.coll]=DB[o.coll].filter(x=>x.id!==o.id); }
            persist(); L.ops.push({op:"delete",coll:o.coll,id:o.id,animal:o.animal,obj:JSON.parse(JSON.stringify(obj))}); did.push("del"); } }
        else if(r<0.76){ if(deadOwn.length){ const o=pick(deadOwn);
            if(o.coll==="diary"){ const a=DB.animals.find(x=>x.id===o.animal); if(!a) continue; (a.diary=a.diary||[]).push(o.obj); }
            else DB[o.coll].push(o.obj);
            persist(); L.ops.push({op:"revive",coll:o.coll,id:o.id,animal:o.animal}); did.push("undo"); } }
        else if(r<0.88){ await runSync(); did.push("sync"); }
        else { await autoPull(); did.push("pull"); }
        this.save(L);
        await this.sleep(Math.floor(rnd()*400));
      }
      return did.join(" ");
    },
    expected(){ const out={present:[],absent:[]}; this.lastOps(this.log()).forEach(o=>{ (o.op==="delete"?out.absent:out.present).push({coll:o.coll,id:o.id}); }); return out; }
  };
})();
