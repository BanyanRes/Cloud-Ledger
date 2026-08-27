path = r"C:\Users\JimmyYun\Cloud-Ledger\server\requisition_draft.js"
raw=open(path,"rb").read(); nl="\r\n" if b"\r\n" in raw else "\n"
s=raw.decode("utf-8"); orig=s
def rep(old,new,n=1):
    global s
    old=old.replace("\n",nl); new=new.replace("\n",nl)
    c=s.count(old); assert c==n,"exp %d of %r got %d"%(n,old[:55],c)
    s=s.replace(old,new,n)

# 1) Add countStreams + cleanReportBaseName helpers before phasedFilename.
anchor = "function phasedFilename(name, phase) {"
helpers = (
"// Count how many requisition streams an entity runs (distinct phases across all\n"
"// its drafts, open or finalized). Used to decide whether a phase label belongs\n"
"// in the filename: with a single stream there is nothing to disambiguate, so\n"
"// the phase is omitted; with two (e.g. 2 and 2a) each name carries its phase.\n"
"function countStreams(db, eid) {\n"
"  try {\n"
"    const rows = db.prepare(\"SELECT DISTINCT IFNULL(phase,'') AS ph FROM requisition_draft WHERE entity_id=?\").all(eid);\n"
"    return rows.length;\n"
"  } catch (_) { return 1; }\n"
"}\n\n"
"// Tidy the report base name: drop a leading code prefix like \"0005 B1 \" and\n"
"// shorten \"County Line SRN\" to \"SRN\". Cosmetic only; the \"Requisition Report\"\n"
"// token (which seeding matches on) is preserved.\n"
"function cleanReportBaseName(name) {\n"
"  let s = String(name || '');\n"
"  s = s.replace(/^\\s*\\d{2,}\\s+[A-Za-z]\\d+\\s+/, '');   // \"0005 B1 \" -> \"\"\n"
"  s = s.replace(/County Line SRN/gi, 'SRN');\n"
"  return s.replace(/\\s{2,}/g, ' ').trim();\n"
"}\n\n"
)
rep(anchor, helpers + anchor)

# 2) phasedFilename: accept opts.multiStream; only add the phase label when multi.
rep(
"function phasedFilename(name, phase) {\n"
"  const lbl = phaseLabel(phase);\n"
"  if (!lbl) return name;",
"function phasedFilename(name, phase, opts) {\n"
"  const multiStream = !!(opts && opts.multiStream);\n"
"  const lbl = phaseLabel(phase);\n"
"  if (!lbl || !multiStream) return name; // single stream -> no phase in the name"
)

# 3) phaseMatchesName: accept opts.singleStream; a single-stream entity's phase-less
#    name matches any requested phase (there is only one report to seed from).
rep(
"function phaseMatchesName(name, phase) {\n"
"  const ph = normPhase(phase);\n"
"  const s = String(name || '');\n"
"  const anyPhase = /\\bphase\\s+[0-9a-z]+/i.test(s);\n"
"  if (!ph) return !anyPhase;",
"function phaseMatchesName(name, phase, opts) {\n"
"  const singleStream = !!(opts && opts.singleStream);\n"
"  const ph = normPhase(phase);\n"
"  const s = String(name || '');\n"
"  const anyPhase = /\\bphase\\s+[0-9a-z]+/i.test(s);\n"
"  if (singleStream && !anyPhase) return true; // one stream, phase-less name is the report\n"
"  if (!ph) return !anyPhase;"
)

# 4) Export countStreams + cleanReportBaseName.
rep(
"  phasedFilename,\n"
"  sha256,\n"
"};",
"  phasedFilename,\n"
"  countStreams,\n"
"  cleanReportBaseName,\n"
"  sha256,\n"
"};"
)

assert s!=orig
open(path,"w",encoding="utf-8",newline="").write(s)
print("draft module: countStreams, cleanReportBaseName, multiStream/singleStream added")
