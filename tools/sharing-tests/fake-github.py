# Static server + an in-memory fake of the GitHub Gist API. Two origins served
# from here (localhost and 127.0.0.1) are two "devices" with separate
# localStorage and one shared Gist between them.
import json, os, sys, time, threading, hashlib, socket
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
ROOT=sys.argv[1]; PORT=int(sys.argv[2])
GISTS={}
STATE={"latency":{},"log":[],"reports":{},"repo":{},"repolog":[],"conflicts":0}
def G(gid):
    return GISTS.setdefault(gid,{"content":"{}","history":[],"id":gid})
LOCK=threading.Lock()
def snap(gid,content):
    g=G(gid)
    v=hashlib.sha1((content+str(time.time())+str(len(g["history"]))).encode()).hexdigest()
    g["history"].insert(0,{"version":v,"committed_at":time.strftime("%Y-%m-%dT%H:%M:%SZ",time.gmtime()),"content":content})
    g["content"]=content
def gist_obj(gid,content=None):
    g=G(gid); c=g["content"] if content is None else content
    return {"id":gid,"updated_at":g["history"][0]["committed_at"] if g["history"] else None,
            "history":[{"version":h["version"],"committed_at":h["committed_at"]} for h in g["history"]],
            "files":{"arbor_workspace.json":{"content":c,"truncated":False}}}
class H(SimpleHTTPRequestHandler):
    def __init__(s,*a,**k): super().__init__(*a,directory=ROOT,**k)
    def log_message(s,*a): pass
    def end_headers(s):
        s.send_header("Cache-Control","no-store"); s.send_header("Access-Control-Allow-Origin","*")
        s.send_header("Access-Control-Allow-Headers","*"); s.send_header("Access-Control-Allow-Methods","GET,POST,PATCH,OPTIONS")
        super().end_headers()
    def j(s,code,obj):
        b=json.dumps(obj).encode(); s.send_response(code); s.send_header("Content-Type","application/json")
        s.send_header("Content-Length",str(len(b))); s.end_headers(); s.wfile.write(b)
    def body(s):
        n=int(s.headers.get("Content-Length") or 0); return json.loads(s.rfile.read(n) or b"{}")
    def do_OPTIONS(s): s.send_response(204); s.end_headers()
    def route(s,method):
        p=s.path.split("?")[0]
        if not p.startswith("/fakegist"): return False
        p=p[len("/fakegist"):]; dev=s.headers.get("X-Device","?")
        if p=="/__reset" and method=="POST":
            b=s.body()
            with LOCK:
                STATE["log"]=[];STATE["latency"]={};STATE["reports"]={}
                for gid in (b.get("gists") or ["g1"]):
                    GISTS.pop(gid,None); snap(gid,b["content"])
            return s.j(200,{"ok":True}) or True
        if p=="/__latency" and method=="POST":
            with LOCK: STATE["latency"]=s.body()
            return s.j(200,{"ok":True}) or True
        if p=="/__sleep":
            import urllib.parse
            q=urllib.parse.parse_qs(s.path.split("?",1)[1] if "?" in s.path else "")
            time.sleep(min(int((q.get("ms") or ["0"])[0]),5000)/1000); return s.j(200,{}) or True
        if p=="/__report" and method=="POST":
            b=s.body()
            with LOCK: STATE["reports"][b["dev"]]=b
            return s.j(200,{"ok":True}) or True
        if p=="/__reports":
            with LOCK: return s.j(200,STATE["reports"]) or True
        if p=="/__state":
            import urllib.parse as _u
            q=_u.parse_qs(s.path.split("?",1)[1] if "?" in s.path else "")
            gid=(q.get("gist") or ["g1"])[0]
            with LOCK: return s.j(200,{"content":G(gid)["content"],"revisions":len(G(gid)["history"]),"log":STATE["log"][-200:]}) or True
        if p=="/__history":
            import urllib.parse as _u
            q=_u.parse_qs(s.path.split("?",1)[1] if "?" in s.path else "")
            gid=(q.get("gist") or ["g1"])[0]
            with LOCK: return s.j(200,[{"version":h["version"],"committed_at":h["committed_at"],"content":h["content"]} for h in G(gid)["history"]]) or True
        # ── fake GitHub repository contents API, with real compare-and-swap
        if p=="/__repo":
            with LOCK: return s.j(200,{"files":{k:json.loads(v["content"]) for k,v in STATE["repo"].items()},
                                       "shas":{k:v["sha"] for k,v in STATE["repo"].items()},
                                       "log":STATE["repolog"][-100:],"conflicts":STATE["conflicts"]}) or True
        if p=="/__repo_reset" and method=="POST":
            with LOCK: STATE["repo"]={}; STATE["repolog"]=[]; STATE["conflicts"]=0
            return s.j(200,{"ok":True}) or True
        if p=="/installation/repositories":
            return s.j(200,{"total_count":1,"repositories":[{"full_name":"lab/arbor-shared"}]}) or True
        # head commit: one sha that moves whenever any file in the repo does,
        # which is what the app polls before fetching anything else
        if p.startswith("/repos/") and "/commits" in p:
            import hashlib as _h2
            with LOCK:
                stamp="|".join(sorted(k+":"+v["sha"] for k,v in STATE["repo"].items())) or "empty"
                return s.j(200,[{"sha":_h2.sha1(stamp.encode()).hexdigest()}]) or True
        if p.startswith("/repos/") and "/contents/" in p:
            owner_repo,path=p[len("/repos/"):].split("/contents/",1)
            import base64 as _b64, hashlib as _h
            if method=="GET":
                with LOCK:
                    f=STATE["repo"].get(path.split("?")[0])
                    if not f: return s.j(404,{"message":"Not Found"}) or True
                    return s.j(200,{"content":_b64.b64encode(f["content"].encode()).decode(),
                                    "sha":f["sha"],"encoding":"base64"}) or True
            if method=="PUT":
                b=s.body(); key=path.split("?")[0]
                content=_b64.b64decode(b["content"]).decode()
                with LOCK:
                    cur=STATE["repo"].get(key)
                    # this is the whole reason for a repo over a gist: a write
                    # that cites a stale version is refused, not applied
                    if cur and b.get("sha")!=cur["sha"]:
                        STATE["conflicts"]+=1
                        return s.j(409,{"message":"is at %s but expected %s"%(cur["sha"],b.get("sha"))}) or True
                    if not cur and b.get("sha"):
                        return s.j(409,{"message":"sha given for a file that does not exist"}) or True
                    sha=_h.sha1((content+str(time.time())).encode()).hexdigest()
                    STATE["repo"][key]={"content":content,"sha":sha}
                    STATE["repolog"].append([time.time(),dev,key,len(content)])
                    return s.j(200,{"content":{"sha":sha}}) or True
        if p.startswith("/repos/") and p.count("/")==3:
            name=p[len("/repos/"):]
            # only one repository exists here, so a wrong name fails the way
            # GitHub would and the app's rollback can actually be tested
            if name!="lab/arbor-shared": return s.j(404,{"message":"Not Found"}) or True
            return s.j(200,{"full_name":name,"private":True,
                            "permissions":{"push":True,"admin":False}}) or True
        lat=STATE["latency"].get(dev,0)
        if lat and method=="GET" and p.startswith("/gists/") and p.count("/")==2: time.sleep(lat/1000)
        if p=="/user": return s.j(200,{"login":"fake"}) or True
        if p in ("/gists","/gists/") and method=="GET": return s.j(200,[gist_obj("g1")]) or True
        parts=p.strip("/").split("/")
        if parts[0]=="gists" and len(parts)==2 and method=="GET":
            with LOCK: STATE["log"].append([time.time(),dev,"GET"]); return s.j(200,gist_obj(parts[1])) or True
        if parts[0]=="gists" and len(parts)==2 and method=="PATCH":
            c=s.body()["files"]["arbor_workspace.json"]["content"]
            with LOCK:
                snap(parts[1],c); d=json.loads(c)
                STATE["log"].append([time.time(),dev,"PATCH",{k:len(d.get(k) or []) for k in
                    ("animals","projects","experiments","schedule","notes","readings","domains","tasks")}])
                return s.j(200,gist_obj(parts[1])) or True
        if parts[0]=="gists" and len(parts)==3 and method=="GET":
            with LOCK:
                h=[x for x in G(parts[1])["history"] if x["version"]==parts[2]]
                return s.j(200 if h else 404,gist_obj(parts[1],h[0]["content"]) if h else {"message":"Not Found"}) or True
        if parts[0]=="gists" and method=="POST":
            with LOCK: snap("g1",s.body()["files"]["arbor_workspace.json"]["content"]); return s.j(201,gist_obj("g1")) or True
        return s.j(404,{"message":"fake: unhandled "+method+" "+p}) or True
    def do_GET(s):
        if not s.route("GET"): super().do_GET()
    def do_POST(s): s.route("POST")
    def do_PATCH(s): s.route("PATCH")
    def do_PUT(s): s.route("PUT")
class Dual(ThreadingHTTPServer):
    address_family=socket.AF_INET6
    def server_bind(s):
        s.socket.setsockopt(socket.IPPROTO_IPV6,socket.IPV6_V6ONLY,0); super().server_bind()
Dual(("::",PORT),H).serve_forever()
