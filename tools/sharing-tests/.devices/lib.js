/* Sharing test suite — runs inside a live Arbor, driving the app's own
   functions. Loaded into both devices; each test says which device acts.
   Everything asserts against the repository or the other device, never
   against the function that was just called. */
window.T = (function(){
  const results=[];
  const sleep=ms=>new Promise(r=>setTimeout(r,ms));
  async function repo(){ return (await fetch('/fakegist/__repo')).json(); }
  async function file(pid){ return (await repo()).files['projects/'+pid+'.json']||null; }
  async function index(){ return (await repo()).files['space.json']||null; }
  // wait until a predicate holds, so timing never decides a verdict
  async function until(fn,ms=20000,step=500){
    const t0=Date.now();
    for(;;){ let v=null; try{ v=await fn(); }catch(e){}
      if(v) return v;
      if(Date.now()-t0>ms) return null;
      await sleep(step); }
  }
  async function push(){ await sleep(300); noteSharedChanges(); await runSharedPush(); await sleep(600); }
  async function pull(){ await sharedPoll(true); await sleep(400); }
  function ok(name,pass,detail){ results.push({name,pass:!!pass,detail:detail===undefined?"":String(detail)}); return !!pass; }
  function eq(name,got,want){ return ok(name, JSON.stringify(got)===JSON.stringify(want), `got ${JSON.stringify(got)} want ${JSON.stringify(want)}`); }
  function report(){ return {pass:results.filter(r=>r.pass).length, fail:results.filter(r=>!r.pass).length,
                             failures:results.filter(r=>!r.pass), all:results}; }
  function reset(){ results.length=0; }

  // ---- fixtures -----------------------------------------------------------
  function mkProject(name,opts){
    const pr=addProject(); pr.name=name; pr.shared=true;
    pr.shareDefaults=Object.assign({timeline:"shared",diary:"shared"},(opts||{}).defaults);
    if((opts||{}).libShare) pr.libShare=opts.libShare;
    persist(); return pr;
  }
  function mkTask(pr,title,when,extra){
    const t=Object.assign({id:uid("tk_"),title,when:when||null,created:Date.now(),
                           projectId:pr?pr.id:null,notes:""},extra||{});
    taskList().push(t); persist(); return t;
  }
  function mkAnimal(pr,name,extra){
    const a=Object.assign({id:uid("a_"),name,alive:true,sex:"M",diary:[],created:Date.now()},extra||{});
    (DB.animals=DB.animals||[]).push(a);
    if(pr){ (pr.links=pr.links||[]).push({kind:"animal",id:a.id}); }
    persist(); return a;
  }
  function mkDiary(a,type,date,time,extra){
    const e=Object.assign({id:uid("d_"),type,date,time:time||null,endTime:null,
                           data:{note:""},created:Date.now()},extra||{});
    (a.diary=a.diary||[]).push(e); persist(); return e;
  }
  function mkExperiment(pr,title){
    const ex=addExperiment(); ex.title=title; ex.projectId=pr?pr.id:null; persist(); return ex;
  }
  function findAnimal(name){ return (DB.animals||[]).find(a=>a.name===name)||null; }
  function findTask(title){ return (DB.schedule||[]).find(t=>t.title===title)||null; }
  function findDiary(id){ for(const a of (DB.animals||[])) for(const e of (a.diary||[])) if(e.id===id) return {a,e}; return null; }
  function findProject(name){ return (DB.projects||[]).find(p=>p.name===name)||null; }

  return {results,sleep,repo,file,index,until,push,pull,ok,eq,report,reset,
          mkProject,mkTask,mkAnimal,mkDiary,mkExperiment,
          findAnimal,findTask,findDiary,findProject};
})();
"suite loaded"
