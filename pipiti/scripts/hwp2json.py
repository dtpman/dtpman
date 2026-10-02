import olefile, zlib, struct, sys
f = olefile.OleFileIO(sys.argv[1])
comp = f.openstream('FileHeader').read()[36] & 1
def recs(d):
    i=0
    while i < len(d):
        h = struct.unpack('<I', d[i:i+4])[0]; i+=4
        tag=h&0x3ff; lvl=(h>>10)&0x3ff; size=h>>20
        if size==0xfff: size=struct.unpack('<I', d[i:i+4])[0]; i+=4
        yield tag,lvl,d[i:i+size]; i+=size
di=f.openstream('DocInfo').read()
if comp: di=zlib.decompress(di,-15)
shapes=[]
for tag,lvl,r in recs(di):
    if tag==21:
        attr=struct.unpack('<I',r[46:50])[0]
        ul=(attr>>2)&3; bold=(attr>>1)&1
        color=struct.unpack('<I',r[52:56])[0] if len(r)>=56 else 0
        shapes.append((ul,bold,color))
secs=sorted([e for e in f.listdir() if e[0]=='BodyText'], key=lambda e:int(e[1][7:]))
out=[]
for s in secs:
    d=f.openstream(s).read()
    if comp: d=zlib.decompress(d,-15)
    text=None
    for tag,lvl,r in recs(d):
        if tag==66: ps=struct.unpack('<H',r[8:10])[0]; sty=r[10]
        if tag==67: text=r
        elif tag==71 and r[:4][::-1] in (b'tbl ',):
            out.append(('TABLE',0,0,''))
        elif tag==71 and r[:4][::-1]==b'gso ':
            out.append(('OBJ',0,0,''))
        elif tag==68:
            runs=[struct.unpack('<II',r[k:k+8]) for k in range(0,len(r),8)]
            if text is None: out.append(('P',sty,lvl,'')); continue
            res=[]; j=0; pos=0; cur=None
            def style(p):
                sid=0
                for a,b in runs:
                    if a<=p: sid=b
                return shapes[sid] if sid<len(shapes) else (0,0,0)
            ul_on=False
            while j<len(text):
                c=struct.unpack('<H',text[j:j+2])[0]
                st=style(pos)
                if c>=32:
                    u=st[0]==1
                    if u and not ul_on: res.append('[['); ul_on=True
                    if not u and ul_on: res.append(']]'); ul_on=False
                    res.append(chr(c)); j+=2; pos+=1
                elif c==10: res.append('\n'); j+=2; pos+=1
                elif c in (13,0): j+=2; pos+=1
                elif c==9: res.append('\t'); j+=16; pos+=8
                elif c in (1,2,3,11,12,14,15,16,17,18,19,20,21,22,23): j+=16; pos+=8
                else: j+=2; pos+=1
            if ul_on: res.append(']]')
            col=style(0)[2]
            out.append(('P',sty,lvl,''.join(res).replace(']][[','')))
            text=None
import json; json.dump(out,open('paras.json','w'),ensure_ascii=False)
