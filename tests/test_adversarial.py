import os, sys, json, tempfile, subprocess
import os as _os
CONVERT=_os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "..", "scripts", "convert.py")
res=[]; blobs=[]
def cli(a,cwd=None):
    p=subprocess.run([sys.executable,CONVERT,*a],capture_output=True,text=True,cwd=cwd); blobs.append(p.stdout+p.stderr); return p.returncode,p.stdout,p.stderr
def ck(n,c,d=""): res.append((n,bool(c),d))
D=tempfile.mkdtemp()
def P(n): return os.path.join(D,n)

# samples
open(P("a.csv"),"w").write("x,y\n1,2\n")
open(P("b.csv"),"w").write("p,q\n3,4\n")
os.makedirs(P("d1")); os.makedirs(P("d2"))
open(P("d1/report.pdf.src"),"w").write("dummy")  # not real pdf; use csv for collision
open(P("d1/same.csv"),"w").write("a\n1\n"); open(P("d2/same.csv"),"w").write("b\n2\n")
# no-extension text file
open(P("READMEnoext"),"w").write("k,v\n1,2\n")
# cp1251 (non-utf8) csv
open(P("win.csv"),"wb").write("город,страна\nМосква,Россия\n".encode("cp1251"))
# big utf-8 text (~2MB) with cyrillic near start
open(P("big.csv"),"w",encoding="utf-8").write("колонка1,колонка2\n"+("значение,второе\n"*120000))
# nested dirs for skip/recursive
os.makedirs(P("tree/sub")); open(P("tree/top.csv"),"w").write("t\n1\n"); open(P("tree/sub/deep.csv"),"w").write("d\n9\n")
# symlink
try:
    os.symlink(P("a.csv"), P("link.csv")); have_link=True
except Exception: have_link=False
# existing file to block -o dir
open(P("blocker"),"w").write("x")
open(P("blocker2"),"w").write("x")

# 1 multiple inputs + -o existing FILE -> clean exit1, no traceback
rc,_,er=cli([P("a.csv"),P("b.csv"),"-o",P("blocker")]); ck("01 multi + -o file -> clean exit1", rc==1 and "ERROR:" in er and "Traceback" not in er, er.strip()[-80:])
# 2 -o dir/ where a FILE exists -> clean exit1
rc,_,er=cli([P("a.csv"),"-o",P("blocker2")+"/"]); ck("02 -o dir/ over existing file -> clean exit1", rc==1 and "ERROR:" in er and "Traceback" not in er)
# 3 collision across dirs: d1/same.csv + d2/same.csv -> two outputs, no overwrite
o=P("oc"); rc,_,_=cli([P("d1/same.csv"),P("d2/same.csv"),"-o",o+"/"]); fs=sorted(os.listdir(o)); ck("03 cross-dir collision deduped", rc==0 and len(fs)==2 and fs[0]!=fs[1], str(fs))
#   verify contents differ (no clobber)
contents=set(open(os.path.join(o,f)).read() for f in fs); ck("03b both contents preserved", len(contents)==2, str(contents))
# 4 no-extension file -> converts (treated as text/plain by markitdown)
rc,so,er=cli([P("READMEnoext")]); ck("04 no-extension file", rc==0 and "Traceback" not in er)
# 5 cp1251 file -> detected, cyrillic correct
rc,_,_=cli([P("win.csv"),"-o",P("win.md")]); txt=open(P("win.md"),encoding="utf-8").read(); ck("05 cp1251 decoded", "Москва" in txt, repr(txt[:40]))
# 5b short, mostly-ASCII cp1251 file: charset detection must not mis-fire to a CJK codec (Big5 etc.)
open(P("short.json"),"wb").write('{"product": "Виджет", "price": 42}'.encode("cp1251"))
rc,_,_=cli([P("short.json"),"-o",P("short.md")]); st=open(P("short.md"),encoding="utf-8").read(); ck("05b short cp1251 not mis-detected as CJK", "Виджет" in st, repr(st[:50]))
# 6 big utf-8 file -> bounded charset read, correct, fast, no MemoryError
rc,_,er=cli([P("big.csv"),"-o",P("big.md")]); head=open(P("big.md"),encoding="utf-8").read(200); ck("06 big utf-8 file ok", rc==0 and "значение" in head and "Traceback" not in er)
# 7 subdir skip note (default)
o=P("ot"); rc,_,er=cli([P("tree"),"-o",o+"/"]); ck("07 subdir skipped w/ note", rc==0 and "skipped" in er.lower() and len(os.listdir(o))==1, er.strip()[-60:])
# 8 recursive includes deep file
o=P("otr"); rc,_,er=cli([P("tree"),"--recursive","-o",o+"/"]); ck("08 --recursive includes deep", rc==0 and len(os.listdir(o))==2, str(os.listdir(o)))
# 9 symlink to a file converts
if have_link:
    rc,_,er=cli([P("link.csv"),"-o",P("lnk.md")]); ck("09 symlink converts", rc==0 and "Traceback" not in er)
else:
    ck("09 symlink converts", True, "skipped (no symlink)")
# 10 nested non-existent output dir for -o file -> parent created
rc,_,er=cli([P("a.csv"),"-o",P("ne/sub/out.md")]); ck("10 nested -o parent created", rc==0 and os.path.exists(P("ne/sub/out.md")))
# 11 directory(2 files) + -o FILE -> multiple => clean error
rc,_,er=cli([P("d1"),P("d2"),"-o",P("blocker")]); ck("11 dir expand multi + -o file -> clean", rc==1 and "Traceback" not in er)
# 12 empty directory -> "No inputs" exit1
os.makedirs(P("emptydir")); rc,_,er=cli([P("emptydir")]); ck("12 empty dir -> exit1 clean", rc==1 and "Traceback" not in er)
# 13 stdout not contaminated: single file to stdout has only markdown (no NOTE/pip)
rc,so,er=cli([P("a.csv")], cwd=P("")); ck("13 stdout clean markdown", rc==0 and so.startswith("|") and "NOTE" not in so and "Traceback" not in er, repr(so[:30]))
# 14 weird filename: leading dash protected via -- ? argparse: use ./ -name. test spaces+unicode+symbols
weird=P("a (1) — копия [v2].csv"); open(weird,"w",encoding="utf-8").write("к,в\n1,2\n")
rc,_,er=cli([weird,"-o",P("weird.md")]); ck("14 weird filename", rc==0 and "Traceback" not in er and os.path.exists(P("weird.md")))
# 14b stdout echo of non-ASCII (CJK+emoji) on the no -o path must not crash a successful
#     conversion (exit 0) and a piping caller must receive faithful UTF-8 bytes.
open(P("echo.csv"),"w",encoding="utf-8").write("项目,emoji\n值,😀\n")
_pe=subprocess.run([sys.executable,CONVERT,P("echo.csv")],capture_output=True,cwd=D); blobs.append(_pe.stderr.decode("utf-8","replace"))
_so=_pe.stdout.decode("utf-8","replace"); ck("14b non-ASCII stdout echo no-crash", _pe.returncode==0 and "项目" in _so and "😀" in _so, "rc=%d so=%r"%(_pe.returncode,_so[:40]))
# 15 global: zero tracebacks anywhere
tb=sum(b.count("Traceback (most recent call last)") for b in blobs); ck("15 ZERO tracebacks", tb==0, "tb=%d"%tb)

p=sum(1 for _,ok,_ in res if ok); f=len(res)-p; w=max(len(n) for n,_,_ in res)
print("="*70)
for n,ok,d in res: print(f"  {'PASS' if ok else 'FAIL'}  {n:<{w}}  {'' if ok else d}")
print("="*70); print(f"ADVERSARIAL: {p}/{len(res)} passed, {f} failed | tracebacks={tb}")
sys.exit(1 if f else 0)
