path = r"C:\Users\JimmyYun\Cloud-Ledger\client\src\App.jsx"
raw=open(path,"rb").read(); nl=b"\r\n" if b"\r\n" in raw else b"\n"
s=raw
def rep(old,new,n=1):
    global s
    old=old.replace(b"\n",nl); new=new.replace(b"\n",nl)
    c=s.count(old); assert c==n, b"exp %d got %d for %r"%(n,c,old[:50])
    s=s.replace(old,new,n)

# Replace the code-first table with a plain-English list (title + help + numbers).
rep(
b"        <table style={S.table}><thead><tr><th style={S.th}>Check</th><th style={S.th}>Level</th><th style={S.thR}>Expected</th><th style={S.thR}>Actual</th><th style={S.th}>Detail</th></tr></thead>\n"
b"          <tbody>{rfDetail.checks.filter(c=>!c.pass).map((c,i)=><tr key={i}>\n"
b"            <td style={{...S.td,fontWeight:600,color:T.textBright}}>{c.id}</td>\n"
b"            <td style={S.td}>{c.level}</td>\n"
b"            <td style={S.tdR}>{c.expected!=null?Number(c.expected).toLocaleString(undefined,{maximumFractionDigits:2}):'\\u2014'}</td>\n"
b"            <td style={S.tdR}>{c.actual!=null?Number(c.actual).toLocaleString(undefined,{maximumFractionDigits:2}):'\\u2014'}</td>\n"
b"            <td style={{...S.td,fontSize:11,color:T.textMuted}}>{c.detail}</td></tr>)}</tbody></table>",
b"        <div>{rfDetail.checks.filter(c=>!c.pass).map((c,i)=>(\n"
b"          <div key={i} style={{padding:'8px 10px',background:T.bgElevated,borderRadius:6,border:'1px solid '+T.red+'25',marginBottom:6}}>\n"
b"            <div style={{fontWeight:600,color:T.text}}>{c.title||'Check failed'}</div>\n"
b"            {c.help?<div style={{fontSize:12,color:T.textMuted,marginTop:2}}>{c.help}</div>:null}\n"
b"            {(c.expected!=null&&c.actual!=null)?(<div style={{fontSize:11,color:T.textMuted,marginTop:2}}>{'Expected '+Number(c.expected).toLocaleString(undefined,{maximumFractionDigits:2})+', found '+Number(c.actual).toLocaleString(undefined,{maximumFractionDigits:2})+'.'}</div>):null}\n"
b"          </div>))}</div>"
)

assert s!=raw
open(path,"wb").write(s)
print("rfDetail checks table -> plain-English list")
