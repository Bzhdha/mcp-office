"""MCP Office: convertit les réponses IA en DOCX/XLSX/PPTX à partir du dernier modèle d'entreprise (SharePoint)."""
from copy import deepcopy
import io,os,re,time,secrets,zipfile,hmac,httpx,uvicorn
from pathlib import Path
from urllib.parse import quote
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import load_workbook,Workbook
from openpyxl.worksheet.table import Table,TableStyleInfo
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.util import Emu
from pptx.oxml.ns import qn as pq
from mcp.server.fastmcp import FastMCP
from starlette.responses import Response,JSONResponse

E=os.environ.get
TENANT,CID,CSEC,SITE,FOLDER=E("SP_TENANT_ID"),E("SP_CLIENT_ID"),E("SP_CLIENT_SECRET"),E("SP_SITE_ID"),E("SP_FOLDER","Modeles")
LOCAL=E("TEMPLATE_DIR");API_KEY=E("MCP_API_KEY","");BASE=E("PUBLIC_BASE_URL","http://localhost:8000").rstrip("/")
OUT=Path(E("OUTPUT_DIR","/tmp/mcp-office"));OUT.mkdir(parents=True,exist_ok=True);TTL=int(E("FILE_TTL","3600"));MAXIN=int(E("MAX_INPUT","500000"))
LANG=E("DOC_LANG","fr-FR");CACHE=int(E("TEMPLATE_CACHE","300"))
EXT={"docx":(".dotx",".docx"),"xlsx":(".xltx",".xlsx"),"pptx":(".potx",".pptx")}
MIME={"docx":"application/vnd.openxmlformats-officedocument.wordprocessingml.document","xlsx":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","pptx":"application/vnd.openxmlformats-officedocument.presentationml.presentation"}
G="https://graph.microsoft.com/v1.0";_tok=[None,0];_tpl={}

# ---------- Modèles (SharePoint via Graph, ou dossier local synchronisé) ----------
def _gtoken():
 if _tok[1]>time.time()+60:return _tok[0]
 r=httpx.post(f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token",data={"grant_type":"client_credentials","client_id":CID,"client_secret":CSEC,"scope":"https://graph.microsoft.com/.default"},timeout=20);r.raise_for_status();j=r.json()
 _tok[:]=[j["access_token"],time.time()+j["expires_in"]];return _tok[0]
def _list():
 if LOCAL:return[{"name":p.name,"modified":p.stat().st_mtime,"id":str(p),"etag":str(p.stat().st_mtime)}for p in Path(LOCAL).iterdir()if p.is_file()]
 h={"Authorization":"Bearer "+_gtoken()};u=f"{G}/sites/{SITE}/drive/root:/{FOLDER.strip('/')}:/children?$select=id,name,lastModifiedDateTime,eTag,file&$top=999";items=[]
 while u:r=httpx.get(u,headers=h,timeout=20);r.raise_for_status();j=r.json();items+=j["value"];u=j.get("@odata.nextLink")
 return[{"name":i["name"],"modified":i["lastModifiedDateTime"],"id":i["id"],"etag":i["eTag"]}for i in items if"file"in i]
def latest(kind,prefix=""):
 """Dernier modèle modifié pour le type (préfixe de nom optionnel)."""
 c=[f for f in _list()if f["name"].lower().endswith(EXT[kind])and not f["name"].startswith("~$")and f["name"].lower().startswith(prefix.lower())]
 if not c:raise ValueError(f"Aucun modèle {kind} trouvé")
 f=max(c,key=lambda x:x["modified"]);k=(kind,prefix);hit=_tpl.get(k)
 if hit and hit[0]==f["etag"]and hit[2]>time.time():return f["name"],hit[1]
 if LOCAL:b=Path(f["id"]).read_bytes()
 else:r=httpx.get(f"{G}/sites/{SITE}/drive/items/{f['id']}/content",headers={"Authorization":"Bearer "+_gtoken()},follow_redirects=True,timeout=60);r.raise_for_status();b=r.content
 _tpl[k]=(f["etag"],b,time.time()+CACHE);return f["name"],b
def _untemplate(b):
 """.dotx/.potx -> document ouvrable (content-type template -> document)."""
 zi,o=zipfile.ZipFile(io.BytesIO(b)),io.BytesIO()
 with zipfile.ZipFile(o,"w",zipfile.ZIP_DEFLATED)as z:
  for it in zi.infolist():
   d=zi.read(it)
   if it.filename=="[Content_Types].xml":d=d.replace(b"wordprocessingml.template.main",b"wordprocessingml.document.main").replace(b"presentationml.template.main",b"presentationml.presentation.main")
   z.writestr(it,d)
 o.seek(0);return o

# ---------- Markdown minimal ----------
def blocks(md):
 L=md.replace("\r","").split("\n");i=0;out=[]
 while i<len(L):
  l=L[i];s=l.strip()
  if not s:i+=1;continue
  if s.startswith("```"):
   i+=1;c=[]
   while i<len(L)and not L[i].strip().startswith("```"):c.append(L[i]);i+=1
   out.append(("code","\n".join(c)));i+=1;continue
  if m:=re.match(r"(#{1,6})\s+(.*)",s):out.append(("h",len(m[1]),m[2]));i+=1;continue
  if s.startswith("|"):
   rows=[]
   while i<len(L)and L[i].strip().startswith("|"):
    r=[c.strip()for c in L[i].strip().strip("|").split("|")]
    if not all(re.fullmatch(r":?-{2,}:?",c)for c in r if c):rows.append(r)
    i+=1
   out.append(("table",rows));continue
  if m:=re.match(r"(\s*)([-*+]|\d+[.)])\s+(.*)",l):out.append(("num"if m[2][0].isdigit()else"ul",min(len(m[1].replace("\t","  "))//2,2),m[3]));i+=1;continue
  p=[s];i+=1
  while i<len(L)and L[i].strip()and not re.match(r"\s*(#|\||```|[-*+]\s|\d+[.)]\s)",L[i]):p.append(L[i].strip());i+=1
  out.append(("p"," ".join(p)))
 return out
def runs(t):return[(x.strip("*`"),x.startswith("**"),x.startswith("*")and not x.startswith("**"),x.startswith("`"))for x in re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)",t)if x]
def plain(t):return"".join(r[0]for r in runs(t))

# ---------- Génération ----------
def _save(kind,name,obj):
 for f in OUT.iterdir():
  if f.stat().st_mtime<time.time()-TTL:f.unlink(missing_ok=True)
 tok=secrets.token_urlsafe(24);safe=re.sub(r"[^\w\-. ]","_",name)[:80].strip()or"document"
 d=OUT/tok;d.mkdir();obj.save(d/f"{safe}.{kind}");return f"{BASE}/files/{tok}/{quote(safe)}.{kind}"

def make_docx(title,markdown,template_prefix=""):
 tn,b=latest("docx",template_prefix);doc=Document(_untemplate(b));body=doc.element.body
 for e in list(body):
  if e.tag!=qn("w:sectPr"):body.remove(e)
 names={s.name for s in doc.styles};st=lambda*n:next((x for x in n if x in names),None)
 doc.core_properties.title=title;doc.core_properties.language=LANG
 def para(t,style=None,pre=""):
  p=doc.add_paragraph(style=style);pre and p.add_run(pre)
  for x,b_,i_,c in runs(t):r=p.add_run(x);r.bold=b_ or None;r.italic=i_ or None;c and setattr(r.font,"name","Consolas")
  return p
 if title:doc.add_heading(title,0)if st("Title")else para(title,st("Heading 1"))
 for bl in blocks(markdown):
  k=bl[0]
  if k=="h":doc.add_heading(plain(bl[2]),min(bl[1],9))
  elif k in("ul","num"):
   base="List Bullet"if k=="ul"else"List Number";s=st(base+(f" {bl[1]+1}"if bl[1]else""),base,"List Paragraph")
   para(bl[2],s,""if s and s.startswith(base)else("• "if k=="ul"else"- "))
  elif k=="code":para(bl[1],st("Code","HTML Preformatted"))
  elif k=="table":
   rows=bl[1];nc=max(map(len,rows));t=doc.add_table(rows=0,cols=nc);t.style=st("Table Grid","Light Grid Accent 1")
   for ri,r in enumerate(rows):
    cells=t.add_row().cells
    for ci in range(nc):
     c=cells[ci];c.text=plain(r[ci])if ci<len(r)else""
     if ri==0:
      for rn in c.paragraphs[0].runs:rn.bold=True
   h=t.rows[0]._tr.get_or_add_trPr();e=OxmlElement("w:tblHeader");e.set(qn("w:val"),"true");h.append(e)  # en-tête répété (accessibilité)
  else:para(bl[1])
 return tn,doc

def _num(v):
 if isinstance(v,str):
  s=v.strip().replace(" ","").replace(" ","")
  if re.fullmatch(r"-?\d+([.,]\d+)?",s):return float(s.replace(",","."))if re.search(r"[.,]",s)else int(s)
 return v
def make_xlsx(title,sheets,template_prefix=""):
 tn,b=latest("xlsx",template_prefix);wb=load_workbook(io.BytesIO(b));wb.template=False;base=wb.worksheets[0]
 wb.properties.title=title;wb.properties.language=LANG;used=set()
 for n,sh in enumerate(sheets):
  ws=base if n==0 else wb.copy_worksheet(base);nm=re.sub(r"[\[\]:*?/\\]","",str(sh.get("name")or f"Feuille{n+1}"))[:31]or f"Feuille{n+1}"
  while nm in used:nm=nm[:28]+f"_{n}"
  used.add(nm);ws.title=nm;rows=sh.get("rows")or[];start=ws.max_row+2 if ws.max_row>1 or ws["A1"].value is not None else 1
  for ri,r in enumerate(rows):
   for ci,v in enumerate(r):ws.cell(start+ri,ci+1,_num(v))
  if len(rows)>1:
   nc=max(map(len,rows));ref=f"A{start}:{get_column_letter(nc)}{start+len(rows)-1}"
   hdr=[str(rows[0][i])if i<len(rows[0])and rows[0][i]not in(None,"")else f"Col{i+1}"for i in range(nc)]
   for i,h in enumerate(hdr):ws.cell(start,i+1,h)
   t=Table(displayName=re.sub(r"\W","_",f"T_{nm}")[:250],ref=ref);t.tableStyleInfo=TableStyleInfo(name=E("XLSX_TABLE_STYLE","TableStyleMedium2"),showRowStripes=True);ws.add_table(t)
   ws.freeze_panes=ws.cell(start+1,1)
   for i in range(nc):ws.column_dimensions[get_column_letter(i+1)].width=min(60,max(10,*(len(str(r[i]))+2 for r in rows if i<len(r))))
 for ws in wb.worksheets[len(sheets):]:wb.remove(ws)
 return tn,wb

# PowerPoint : le modèle contient des diapos types préparées -> on les duplique et on remplit leurs zones.
RNS="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
def _stitle(s):
 t=s.shapes.title
 return(t.text_frame.text.strip()if t is not None and t.has_text_frame else"")or s.slide_layout.name
def slide_types(p):
 """Diapos types du modèle : n°, libellé, zones remplissables (nom de forme PowerPoint)."""
 r=[]
 for i,s in enumerate(p.slides,1):
  z=[]
  for sh in s.shapes:
   k="table"if sh.has_table else"title"if sh.is_placeholder and int(sh.placeholder_format.type)in(1,3)else"text"if sh.has_text_frame else None
   if k:z.append({"zone":sh.name,"kind":k,"sample":(sh.text_frame.text[:60]if sh.has_text_frame else f"{len(sh.table.rows)}x{len(sh.table.columns)}")})
  r.append({"type":i,"label":_stitle(s),"zones":z})
 return r
def _pick(p,t,has_table=False):
 S=list(p.slides)
 if isinstance(t,int)or str(t).isdigit():return S[int(t)-1]
 if t:
  for s in S:
   if str(t).lower()in _stitle(s).lower():return s
  raise ValueError(f"Diapo type « {t} » introuvable (voir list_slide_types)")
 for s in S[1:]or S:  # auto : 1re diapo (hors couverture) avec tableau ou corps de texte
  if has_table and any(sh.has_table for sh in s.shapes):return s
  if not has_table and any(sh.is_placeholder and int(sh.placeholder_format.type)in(2,7)for sh in s.shapes):return s
 return S[-1]
def _dup(p,src):
 s=p.slides.add_slide(src.slide_layout);tree=s.shapes._spTree
 for sh in list(s.shapes):tree.remove(sh._element)
 m={}
 for rid,rel in src.part.rels.items():
  if rel.reltype.endswith(("/slideLayout","/notesSlide")):continue
  m[rid]=s.part.rels.get_or_add_ext_rel(rel.reltype,rel.target_ref)if rel.is_external else s.part.relate_to(rel.target_part,rel.reltype)
 bg=src._element.cSld.bg
 if bg is not None:s._element.cSld.insert(0,deepcopy(bg))
 for el in list(src.shapes._spTree)[2:]:
  e=deepcopy(el)
  for x in e.iter():
   for a in("embed","id","link"):
    v=x.get(f"{{{RNS}}}{a}")
    if v in m:x.set(f"{{{RNS}}}{a}",m[v])
  tree.append(e)
 return s
def _fill(tf,lines):
 """Remplace le texte en gardant la mise en forme de la 1re ligne (et le niveau)."""
 ps=tf._txBody.findall(pq("a:p"));p0=ps[0]
 r0=p0.find(pq("a:r"));rpr=deepcopy(r0.find(pq("a:rPr")))if r0 is not None and r0.find(pq("a:rPr"))is not None else None
 for x in ps:tf._txBody.remove(x)
 for ln in lines or[""]:
  p=deepcopy(p0)
  for c in list(p):
   if c.tag!=pq("a:pPr"):p.remove(c)
  tf._txBody.append(p);pa=tf.paragraphs[-1];lvl=(len(ln)-len(ln.lstrip()))//2
  if lvl:pa.level=min(lvl,8)
  for x,b_,i_,_ in runs(re.sub(r"^[-*•]\s+","",ln.strip())):
   r=pa.add_run()
   if rpr is not None:r._r.insert(0,deepcopy(rpr))
   r.text=x;b_ and setattr(r.font,"bold",True);i_ and setattr(r.font,"italic",True)
def _table(t,rows):
 tr=t._tbl.tr_lst;nc=len(t.columns)
 while len(t._tbl.tr_lst)<len(rows):t._tbl.append(deepcopy(t._tbl.tr_lst[-1]))
 for x in t._tbl.tr_lst[len(rows):]:t._tbl.remove(x)
 for ri,r in enumerate(rows):
  for ci in range(nc):_fill(t.cell(ri,ci).text_frame,[str(r[ci])if ci<len(r)else""])
def make_pptx(title,slides,subtitle="",template_prefix=""):
 tn,b=latest("pptx",template_prefix);p=Presentation(_untemplate(b));orig=list(p.slides)
 if not orig:raise ValueError("Le modèle PowerPoint ne contient aucune diapo type")
 p.core_properties.title=title;p.core_properties.language=LANG
 for d in[{"type":1,"title":title,"bullets":[subtitle]if subtitle else[]}]+list(slides):
  tb=d.get("table");s=_dup(p,_pick(p,d.get("type"),bool(tb)));zones=d.get("zones")or{};body=[d.get("bullets")or[],d.get("bullets2")or[]];bi=0
  for sh in list(s.shapes):
   ph=sh.is_placeholder and int(sh.placeholder_format.type)
   if sh.name in zones:
    v=zones[sh.name]
    if sh.has_table and isinstance(v,list):_table(sh.table,v)
    elif sh.has_text_frame:_fill(sh.text_frame,v if isinstance(v,list)else str(v).split("\n"))
   elif sh.has_table and tb:_table(sh.table,tb);tb=None
   elif ph in(1,3):_fill(sh.text_frame,[d.get("title","")])
   elif ph in(2,4,7)and sh.has_text_frame:
    if bi<2 and body[bi]:_fill(sh.text_frame,body[bi]);bi+=1
    else:sh._element.getparent().remove(sh._element)  # zone vide -> supprimée (pas de texte d'exemple)
  if tb:W,H=p.slide_width,p.slide_height;_table(s.shapes.add_table(len(tb),max(map(len,tb)),Emu(W//20),Emu(H//4),Emu(W*9//10),Emu(H//2)).table,tb)
  if d.get("notes"):s.notes_slide.notes_text_frame.text=d["notes"]
 sl=p.slides._sldIdLst
 for s in orig:
  for x in list(sl):
   if p.part.related_part(x.rId)is s.part:p.part.drop_rel(x.rId);sl.remove(x)
 return tn,p

# ---------- Serveur MCP ----------
mcp=FastMCP("office-templates",host=E("HOST","0.0.0.0"),port=int(E("PORT","8000")),stateless_http=True)
def _chk(*a):
 if len(repr(a))>MAXIN:raise ValueError("Contenu trop volumineux")
def _ret(kind,name,tn,obj):u=_save(kind,name,obj);return f"Document généré avec le modèle « {tn} » : [{name}.{kind}]({u}) (lien valable {TTL//60} min)"

@mcp.tool()
def list_templates()->str:
 """Liste les modèles d'entreprise disponibles et indique le plus récent par type."""
 f=_list();r=[]
 for k,e in EXT.items():
  c=sorted([x for x in f if x["name"].lower().endswith(e)],key=lambda x:x["modified"],reverse=True)
  r.append(f"{k}: "+(", ".join(x["name"]for x in c)or"aucun")+(f" (utilisé: {c[0]['name']})"if c else""))
 return"\n".join(r)
@mcp.tool()
def create_word(title:str,markdown:str,filename:str="",template_prefix:str="")->str:
 """Crée un document Word (.docx) avec le dernier modèle d'entreprise. `markdown`: contenu (titres #, listes -, 1., tableaux |, **gras**, *italique*, blocs ```). `template_prefix`: filtre optionnel sur le nom du modèle."""
 _chk(title,markdown);tn,d=make_docx(title,markdown,template_prefix);return _ret("docx",filename or title,tn,d)
@mcp.tool()
def create_excel(title:str,sheets:list[dict],filename:str="",template_prefix:str="")->str:
 """Crée un classeur Excel (.xlsx) avec le dernier modèle. `sheets`: [{"name":"Ventes","rows":[["Col1","Col2"],[1,2]]}], 1re ligne = en-têtes (mise en tableau Excel)."""
 _chk(title,sheets);tn,w=make_xlsx(title,sheets,template_prefix);return _ret("xlsx",filename or title,tn,w)
@mcp.tool()
def list_slide_types(template_prefix:str="")->str:
 """Liste les diapos types du dernier modèle PowerPoint (n°, libellé, zones nommées). À appeler avant create_powerpoint."""
 tn,b=latest("pptx",template_prefix);return f"Modèle {tn}\n"+"\n".join(f"{t['type']}. {t['label']} | zones: "+"; ".join(f"{z['zone']} ({z['kind']}: {z['sample']!r})"for z in t["zones"])for t in slide_types(Presentation(_untemplate(b))))
@mcp.tool()
def create_powerpoint(title:str,slides:list[dict],subtitle:str="",filename:str="",template_prefix:str="")->str:
 """Crée une présentation (.pptx) en dupliquant les diapos types du dernier modèle (couverture = diapo 1 avec title/subtitle). `slides`: [{"type":3 ou "Sommaire" (n° ou libellé, cf. list_slide_types; omis = auto),"title":"...","bullets":["point","  sous-point"],"bullets2":[2e colonne],"table":[["A","B"],["1","2"]],"zones":{"NomForme":"texte" ou [lignes] ou [[tableau]]},"notes":"..."}]. Les zones de contenu non remplies sont supprimées ; logos, images et décor du modèle sont conservés."""
 _chk(title,slides);tn,p=make_pptx(title,slides,subtitle,template_prefix);return _ret("pptx",filename or title,tn,p)

@mcp.custom_route("/files/{tok}/{name}",methods=["GET"])
async def dl(req):
 tok,name=req.path_params["tok"],req.path_params["name"]
 if not re.fullmatch(r"[\w-]{20,64}",tok)or"/"in name or".."in name:return Response(status_code=404)
 f=OUT/tok/name;kind=f.suffix[1:]
 if not f.is_file()or kind not in MIME or f.stat().st_mtime<time.time()-TTL:return Response(status_code=404)
 return Response(f.read_bytes(),media_type=MIME[kind],headers={"Content-Disposition":f"attachment; filename*=UTF-8''{quote(name)}","X-Content-Type-Options":"nosniff","Cache-Control":"no-store"})
@mcp.custom_route("/health",methods=["GET"])
async def health(_):return JSONResponse({"ok":True})

def app():
 a=mcp.streamable_http_app()
 async def auth(scope,rcv,snd):
  if scope["type"]=="http"and scope["path"].startswith("/mcp")and API_KEY:
   h=dict(scope["headers"]).get(b"authorization",b"").decode()
   if not hmac.compare_digest(h,"Bearer "+API_KEY):return await JSONResponse({"error":"unauthorized"},401)(scope,rcv,snd)
  await a(scope,rcv,snd)
 return auth
if __name__=="__main__":
 if not API_KEY:print("⚠ MCP_API_KEY non défini : endpoint /mcp non protégé")
 uvicorn.run(app(),host=E("HOST","0.0.0.0"),port=int(E("PORT","8000")),proxy_headers=True)
