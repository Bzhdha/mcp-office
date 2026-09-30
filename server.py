"""MCP Office: convertit les réponses IA en DOCX/XLSX/PPTX à partir du dernier modèle d'entreprise (SharePoint)."""
from copy import deepcopy
import io,os,re,json,time,secrets,zipfile,hmac,httpx,uvicorn
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
def _ver(n,mod):
 """Clé « plus récent » : année, puis version (v3.2), puis date de modification."""
 y=re.search(r"(?<!\d)(20\d\d)(?!\d)",n);v=re.search(r"[vV](\d+(?:\.\d+)*)",n)
 return(int(y[1])if y else 0,tuple(map(int,v[1].split(".")))if v else(),str(mod))
def latest(kind,prefix=""):
 """Dernier modèle du type. `prefix` : partie du nom (sinon DEFAULT_TEMPLATE_<TYPE>)."""
 a=[f for f in _list()if f["name"].lower().endswith(EXT[kind])and not f["name"].startswith("~$")]
 if not prefix:  # défaut par type (famille de modèles), repli sur tous
  prefix=E(f"DEFAULT_TEMPLATE_{kind.upper()}","theme"if kind=="pptx"else"")
  if not any(prefix.lower()in f["name"].lower()for f in a):prefix=""
 c=[f for f in a if prefix.lower()in f["name"].lower()]
 if not c:raise ValueError(f"Aucun modèle {kind} correspondant à « {prefix} » (voir list_templates)")
 f=max(c,key=lambda x:_ver(x["name"],x["modified"]));k=(kind,prefix);hit=_tpl.get(k)
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

def make_docx(title,markdown,template=""):
 tn,b=latest("docx",template);doc=Document(_untemplate(b));body=doc.element.body
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
def make_xlsx(title,sheets,template=""):
 tn,b=latest("xlsx",template);wb=load_workbook(io.BytesIO(b));wb.template=False;base=wb.worksheets[0]
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

# ---------- PowerPoint : construction stricte à partir des dispositions (layouts) du thème ----------
# Chaque diapo = une disposition du thème (L n°) ou, si le modèle en contient, une diapo préparée dupliquée (n°).
# Aucune forme libre ni mise en forme manuelle : seules les zones prévues par le thème sont remplies.
RNS="http://schemas.openxmlformats.org/officeDocument/2006/relationships";BODY=(2,4,7);TITLE=(1,3);PIC=18
PREP=E("PPTX_PREPARED_SLIDES","auto").lower()  # on / off / auto (= sauf modèles dont le nom contient « theme »)
GUIDE=json.loads(Path(E("LAYOUT_GUIDE")).read_text("utf-8"))if E("LAYOUT_GUIDE")else{}  # {"nom disposition":"quand l'utiliser"}
def _walk(shapes):
 for sh in shapes:
  if sh.shape_type==6:yield from _walk(sh.shapes)  # groupe
  else:yield sh
def _pht(sh):return int(sh.placeholder_format.type)if sh.is_placeholder else None
def _txt(t):return re.sub(r"\s+"," ",t).strip()
def _stitle(s):
 t=s.shapes.title;return _txt(t.text_frame.text)if t is not None and t.has_text_frame and t.text_frame.text.strip()else s.slide_layout.name
def _prep(tn):return PREP=="on"or(PREP=="auto"and"theme"not in tn.lower())
def _sz(sh,p):
 """Taille de police effective (pt) : texte, puis héritage disposition -> masque."""
 for x in _chain(sh):
  v=x._element.xpath(".//a:rPr/@sz|.//a:lvl1pPr/a:defRPr/@sz|.//a:endParaRPr/@sz")
  if v:return int(v[0])/100
 v=p.slide_master._element.xpath(f"./p:txStyles/p:{'titleStyle'if _pht(sh)in TITLE else'bodyStyle'}/a:lvl1pPr/a:defRPr/@sz")
 return int(v[0])/100 if v else 14
def _chain(sh):
 c=[sh]
 try:
  b=sh._base_placeholder
  while b is not None:c.append(b);b=getattr(b,"_base_placeholder",None)
 except Exception:pass
 return c
def _cap(sh,p,sib=()):
 """(caractères par ligne, lignes) estimés pour une zone (zone extensible : jusqu'à la zone suivante en dessous)."""
 sz=_sz(sh,p);w=(sh.width or 0)/12700-14;h=(sh.height or 0)/12700-7
 if next((x._element.xpath(".//a:bodyPr/a:spAutoFit")for x in _chain(sh)if x._element.xpath(".//a:bodyPr/*[self::a:spAutoFit or self::a:noAutofit or self::a:normAutofit]")),None):
  t,l,r=sh.top or 0,sh.left or 0,(sh.left or 0)+(sh.width or 0)
  lim=min([x.top for x in sib if x is not sh and(x.top or 0)>t+(sh.height or 0)//2 and(x.left or 0)<r and(x.left or 0)+(x.width or 0)>l]+[int(p.slide_height*0.9)])
  h=max(h,(lim-t)/12700-7)
 return max(1,int(w/(sz*(0.62 if _pht(sh)in TITLE else 0.52)))),max(1,round(h/sz))
def _kind(sh):
 pt=_pht(sh);return"table"if sh.has_table else"titre"if pt in TITLE else"image"if pt==PIC else"texte"if sh.has_text_frame and(sh.is_placeholder or sh.text_frame.text.strip())else None
def _zdesc(sh,p,key,k,sib=()):
 W,H=p.slide_width,p.slide_height;pos=f"@{100*(sh.left or 0)//W},{100*(sh.top or 0)//H} {100*(sh.width or 0)//W}x{100*(sh.height or 0)//H}%"
 if k=="table":return f"  #{key} table {pos} {len(sh.table.rows)}x{len(sh.table.columns)} en-tête={[_txt(c.text)for c in sh.table.rows[0].cells][:8]}"
 if k=="image":return f"  #{key} image {pos} (à insérer par l'utilisateur)"
 cpl,nl=_cap(sh,p,sib);t=_txt(sh.text_frame.text)[:60]
 return f"  #{key} {k} {pos} ≤{nl}l×{cpl}c : {t!r}"
def _yx(sh):return((sh.top or 0)//200000,sh.left or 0)
def slide_types(p,tn,detail=None):
 """Dispositions (L n°) et diapos préparées (n°). `detail` : clés à détailler ({"L12","4"})."""
 o=[]
 for i,l in enumerate(p.slide_layouts,1):
  if detail and f"L{i}"not in detail:continue
  z=[(ph,_kind(ph))for ph in sorted(l.placeholders,key=_yx)if int(ph.placeholder_format.type)not in(13,15,16)]
  g=f" — {GUIDE[l.name]}"if l.name in GUIDE else""
  if not detail:c={};[c.__setitem__(k,c.get(k,0)+1)for _,k in z];o.append(f"L{i}. {l.name}{g} ("+", ".join(f"{v} {k}"for k,v in c.items())+")");continue
  o.append(f"L{i}. {l.name}{g}\n"+"\n".join(_zdesc(ph,p,ph.placeholder_format.idx,k,[x for x,_ in z])for ph,k in z))
 if _prep(tn):
  for i,s in enumerate(p.slides,1):
   if detail and str(i)not in detail:continue
   z=[(sh,_kind(sh))for sh in sorted(_walk(s.shapes),key=_yx)];z=[x for x in z if x[1]]
   if not detail:o.append(f"{i}. [diapo préparée] {_stitle(s)} ({len(z)} zones)");continue
   o.append(f"{i}. [diapo préparée] {_stitle(s)}\n"+"\n".join(_zdesc(sh,p,sh.shape_id,k,[x for x,_ in z])for sh,k in z))
 return"\n".join(o)or"Aucune disposition correspondante"
def _pick(p,S,t,tn,has_table=False):
 """-> ("layout", disposition) ou ("slide", diapo préparée)."""
 L=list(p.slide_layouts);t=""if t is None else str(t).strip()
 if m:=re.fullmatch(r"[Ll]\s*(\d+)",t):
  if not 1<=int(m[1])<=len(L):raise ValueError(f"Disposition {t} inexistante (L1..L{len(L)})")
  return"layout",L[int(m[1])-1]
 if t.isdigit():
  if not _prep(tn)or not 1<=int(t)<=len(S):raise ValueError(f"Diapo préparée {t} indisponible : utiliser une disposition « L n° » (voir list_slide_types)")
  return"slide",S[int(t)-1]
 if t:
  for l in L:
   if t.lower()==l.name.lower():return"layout",l
  for l in L:
   if t.lower()in l.name.lower():return"layout",l
  raise ValueError(f"Disposition « {t} » introuvable (voir list_slide_types)")
 for l in L:  # auto : 1re disposition titre + contenu
  ty={int(ph.placeholder_format.type)for ph in l.placeholders}
  if ty&set(TITLE)and ty&{2,7}:return"layout",l
 return"layout",L[0]
def _remap(el,m):
 for x in el.iter():
  for a in("embed","id","link","pict"):
   k=f"{{{RNS}}}{a}";v=x.get(k)
   if v in m:x.set(k,m[v])
def _clone(part,pkg):
 """Copie une part (graphique, classeur embarqué...) et ses dépendances pour ne pas la partager entre diapos."""
 t=re.sub(r"\d+(?=\.\w+$)","%d",str(part.partname));t=t if"%d"in t else re.sub(r"(?=\.\w+$)","%d",t)
 n=type(part).load(pkg.next_partname(t),part.content_type,pkg,part.blob);m={}
 for rid,rel in part.rels.items():m[rid]=n.rels.get_or_add_ext_rel(rel.reltype,rel.target_ref)if rel.is_external else n.relate_to(_clone(rel.target_part,pkg),rel.reltype)
 if hasattr(n,"_element"):_remap(n._element,m)
 return n
def _dup(p,src):
 s=p.slides.add_slide(src.slide_layout);tree=s.shapes._spTree
 for sh in list(s.shapes):tree.remove(sh._element)
 m={}
 for rid,rel in src.part.rels.items():
  if rel.reltype.endswith(("/slideLayout","/notesSlide")):continue
  if rel.is_external:m[rid]=s.part.rels.get_or_add_ext_rel(rel.reltype,rel.target_ref)
  else:m[rid]=s.part.relate_to(_clone(rel.target_part,p.part.package)if rel.reltype.endswith(("/chart","/oleObject","/package","/diagramData","/diagramDrawing"))else rel.target_part,rel.reltype)
 bg=src._element.cSld.bg
 if bg is not None:s._element.cSld.insert(0,deepcopy(bg))
 for el in list(src.shapes._spTree)[2:]:
  e=deepcopy(el);_remap(e,m);tree.append(e)
 return s
def _fill(tf,lines):
 """Remplace le texte en gardant la mise en forme du thème (1re ligne existante) ; seuls gras/italique ajoutés."""
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
def _lines(v):return[str(x)for x in v]if isinstance(v,list)else str(v).split("\n")
def _over(sh,p,lines,sib=()):
 cpl,nl=_cap(sh,p,sib);need=sum(max(1,-(-len(plain(l.strip()))//cpl))for l in lines);return need>nl and f"{need} lignes pour {nl} disponibles"
def _table(t,rows):
 nc=len(t.columns);w=max(map(len,rows))
 if 0<w<nc and not t._tbl.xpath(".//a:tc[@gridSpan or @hMerge]"):  # colonnes en trop supprimées, largeur totale conservée
  g=t._tbl.tblGrid;cols=g.findall(pq("a:gridCol"));tot=sum(int(c.get("w"))for c in cols);keep=cols[:w];kw=sum(int(c.get("w"))for c in keep)
  for c in cols[w:]:g.remove(c)
  for c in keep:c.set("w",str(int(c.get("w"))*tot//kw))
  for tr in t._tbl.tr_lst:
   for tc in tr.findall(pq("a:tc"))[w:]:tr.remove(tc)
  nc=w
 while len(t._tbl.tr_lst)<len(rows):t._tbl.append(deepcopy(t._tbl.tr_lst[-1]))
 for x in t._tbl.tr_lst[len(rows):]:t._tbl.remove(x)
 for ri,r in enumerate(rows):
  for ci in range(nc):_fill(t.cell(ri,ci).text_frame,[str(r[ci])if ci<len(r)else""])
def make_pptx(title,slides,template=""):
 tn,b=latest("pptx",template);p=Presentation(_untemplate(b));S=list(p.slides);warn=[]
 if not slides:raise ValueError("Aucune diapo demandée")
 p.core_properties.title=title;p.core_properties.language=LANG;drop=re.compile(E("PPTX_DROP",r"^\s*(EXEMPLES?\b|[AÀ] REMPLIR|REPRENDRE LES SLIDES)"),re.I)
 for n,d in enumerate(slides,1):
  tb=d.get("table");kind,src=_pick(p,S,d.get("type"),tn,bool(tb));s=_dup(p,src)if kind=="slide"else p.slides.add_slide(src);w=[]
  Z={str(k).lstrip("#"):v for k,v in(d.get("zones")or{}).items()};done=set();shs=list(_walk(s.shapes))
  key=(lambda x:str(x.shape_id))if kind=="slide"else(lambda x:str(x.placeholder_format.idx)if x.is_placeholder else"")
  if bad:=set(Z)-{key(x)for x in shs}-{x.name for x in shs}:w.append("zones inconnues "+", ".join(f"#{x}"for x in sorted(bad)))
  def put(sh,v):
   if sh.has_table and isinstance(v,list)and v and isinstance(v[0],list):_table(sh.table,v)
   elif sh.has_text_frame:
    L=_lines(v);_fill(sh.text_frame,L)
    if o:=_over(sh,p,L,[x for x in shs if x.is_placeholder or x.has_text_frame]):w.append(f"#{key(sh)} trop long ({o})")
   done.add(sh.shape_id)
  for sh in shs:
   if not sh.is_placeholder and sh.has_text_frame and drop.search(sh.text_frame.text):sh._element.getparent().remove(sh._element);done.add(sh.shape_id);continue  # post-it de consigne
   v=Z.get(key(sh),Z.get(sh.name))
   if v is not None:put(sh,v)
  free=lambda f:[sh for sh in shs if sh.shape_id not in done and f(sh)]
  if d.get("title")is not None:
   t=free(lambda x:_pht(x)in TITLE)
   if t:put(t[0],d["title"])
   else:w.append("pas de zone titre sur cette disposition (titre ignoré)")
  big=lambda:sorted(free(lambda x:_pht(x)in BODY and x.has_text_frame),key=lambda x:-(x.width or 0)*(x.height or 0))
  if tb:
   t=free(lambda x:x.has_table)
   if t:_table(t[0].table,tb);done.add(t[0].shape_id)
   elif z:=big():  # tableau posé dans la plus grande zone de contenu (style de tableau par défaut du thème)
    z=z[0];_table(s.shapes.add_table(len(tb),max(map(len,tb)),z.left,z.top,z.width,z.height).table,tb);z._element.getparent().remove(z._element);done.add(z.shape_id)
   else:w.append("aucune zone pour le tableau (ignoré)")
  for v in(d.get("bullets"),d.get("bullets2")):
   if v:
    z=big()
    if z:put(z[0],v)
    else:w.append("pas de zone de contenu libre pour bullets")
  kept=[]
  for sh in free(lambda x:x.has_text_frame and _pht(x)in BODY):
   if not sh.text_frame.text.strip():sh._element.getparent().remove(sh._element)  # zone vide -> supprimée (aucun texte d'invite)
   elif len(_txt(sh.text_frame.text))>3:kept.append(f"#{key(sh)} {_txt(sh.text_frame.text)[:30]!r}")
  if kept and len(d)>1:w.append("texte d'exemple conservé "+", ".join(kept[:6]))
  if t:=free(lambda x:_pht(x)in TITLE and not x.text_frame.text.strip()):w.append("titre vide (accessibilité : chaque diapo doit avoir un titre)")
  if im:=free(lambda x:_pht(x)==PIC):w.append("image(s) à insérer : "+", ".join(f"#{key(x)}"for x in im))
  if d.get("notes"):s.notes_slide.notes_text_frame.text=d["notes"]
  if w:warn.append(f"diapo {n} ({s.slide_layout.name if kind=='layout'else _stitle(s)[:30]}) : "+" ; ".join(w))
 sid=p.slides._sldIdLst
 for s in S:
  for x in list(sid):
   if p.part.related_part(x.rId)is s.part:p.part.drop_rel(x.rId);sid.remove(x)
 p.warnings=warn;return tn,p

# ---------- Serveur MCP ----------
INSTR="""Serveur de documents conformes aux modèles d'entreprise.
PowerPoint — démarche à suivre :
1. Avec l'utilisateur, définir l'objectif, le public et le plan (couverture/ouverture, sommaire, intercalaires de section, contenus, conclusion).
2. Appeler list_slide_types (liste compacte) puis proposer pour chaque diapo la disposition (L n°) la plus adaptée à son contenu (chiffres clés, colonnes, équipe, citation, référence client…) ; faire valider le plan.
3. Appeler list_slide_types avec les dispositions retenues ("L4,L13,…") pour connaître leurs zones (#id, position, capacité ≤lignes×caractères, texte d'invite = rôle attendu).
4. Remplir chaque zone via `zones` en respectant rôle et capacité ; textes courts ; n'utiliser `bullets` que sur les dispositions titre + contenu.
5. create_powerpoint avec dry_run=true pour vérifier, corriger les alertes (texte trop long, titre vide, zones inconnues), puis générer.
Ne jamais inventer de mise en forme : le thème impose polices, couleurs et positions."""
mcp=FastMCP("office-templates",instructions=INSTR,host=E("HOST","0.0.0.0"),port=int(E("PORT","8000")),stateless_http=True)
def _chk(*a):
 if len(repr(a))>MAXIN:raise ValueError("Contenu trop volumineux")
def _ret(kind,name,tn,obj):u=_save(kind,name,obj);return f"Document généré avec le modèle « {tn} » : [{name}.{kind}]({u}) (lien valable {TTL//60} min)"

@mcp.tool()
def list_templates()->str:
 """Liste les modèles d'entreprise disponibles et indique le plus récent par type."""
 f=_list();r=[]
 for k,e in EXT.items():
  c=sorted([x for x in f if x["name"].lower().endswith(e)],key=lambda x:_ver(x["name"],x["modified"]),reverse=True)
  try:u=latest(k)[0]
  except ValueError:u="aucun"
  r.append(f"{k}: "+(", ".join(x["name"]for x in c)or"aucun")+f" (par défaut: {u})")
 return"\n".join(r)
@mcp.tool()
def create_word(title:str,markdown:str,filename:str="",template:str="")->str:
 """Crée un document Word (.docx) avec le dernier modèle d'entreprise. `markdown`: contenu (titres #, listes -, 1., tableaux |, **gras**, *italique*, blocs ```). `template`: filtre optionnel (partie du nom du modèle)."""
 _chk(title,markdown);tn,d=make_docx(title,markdown,template);return _ret("docx",filename or title,tn,d)
@mcp.tool()
def create_excel(title:str,sheets:list[dict],filename:str="",template:str="")->str:
 """Crée un classeur Excel (.xlsx) avec le dernier modèle. `sheets`: [{"name":"Ventes","rows":[["Col1","Col2"],[1,2]]}], 1re ligne = en-têtes (mise en tableau Excel)."""
 _chk(title,sheets);tn,w=make_xlsx(title,sheets,template);return _ret("xlsx",filename or title,tn,w)
@mcp.tool()
def list_slide_types(slides:str="",template:str="")->str:
 """Dispositions du thème PowerPoint (L n°) [+ diapos préparées (n°) pour les modèles qui en ont]. Sans `slides` : catalogue compact pour choisir la disposition de chaque diapo. Avec `slides`="L4,L13" : zones de chaque disposition triées haut->bas, gauche->droite (#id, type, @x,y l×h % de la diapo, capacité ≤lignes×caractères, texte d'invite = rôle attendu). `template` : partie du nom du modèle."""
 tn,b=latest("pptx",template);det={x.upper()for x in re.findall(r"[Ll]?\d+",slides)}or None
 return f"Modèle {tn}\n"+slide_types(Presentation(_untemplate(b)),tn,det)
@mcp.tool()
def create_powerpoint(title:str,slides:list[dict],filename:str="",template:str="",dry_run:bool=False)->str:
 """Crée une présentation (.pptx) strictement conforme au thème : chaque diapo utilise une disposition du thème, dans l'ordre donné. `slides`: [{"type":"L13" (ou nom exact de disposition ; n° = diapo préparée si le modèle en a), "title":"...", "zones":{"#12":"texte" | ["ligne","  sous-ligne"] | [[tableau]]} (ids via list_slide_types), "bullets":[...] (dispositions titre + contenu), "table":[["En-tête","..."],["..."]], "notes":"notes orateur"}]. Zones non remplies supprimées ; **gras** / *italique* acceptés. `dry_run`=true : vérifie le plan sans générer (alertes : texte trop long, zone inconnue, titre vide, image à insérer). `template` : partie du nom du modèle."""
 _chk(title,slides);tn,p=make_pptx(title,slides,template);wr="\nÀ vérifier : "+" | ".join(p.warnings)if p.warnings else""
 if dry_run:return f"Plan vérifié avec « {tn} » ({len(slides)} diapos)."+(wr or" Aucune alerte.")
 return _ret("pptx",filename or title,tn,p)+wr
@mcp.tool()
def layout_catalog(template:str="")->str:
 """Génère un catalogue visuel (.pptx) : une diapo par disposition du thème, chaque zone étiquetée « #id ». Utile pour choisir les dispositions avec l'utilisateur."""
 tn,b=latest("pptx",template);p=Presentation(_untemplate(b))
 sl=[{"type":f"L{i}","zones":{str(ph.placeholder_format.idx):("L%d · %s"%(i,l.name)if int(ph.placeholder_format.type)in TITLE else f"#{ph.placeholder_format.idx}")for ph in l.placeholders if int(ph.placeholder_format.type)not in(PIC,13,15,16)}}for i,l in enumerate(p.slide_layouts,1)]
 tn,p=make_pptx("Catalogue des dispositions",sl,template);return _ret("pptx","Catalogue dispositions",tn,p)

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
