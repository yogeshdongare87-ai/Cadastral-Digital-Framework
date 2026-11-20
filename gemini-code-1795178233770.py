import io, os, re, json, base64, webbrowser
import tkinter as tk
from tkinter import filedialog
from IPython.display import display, HTML
from pptx import Presentation
from PIL import Image

QC_12_TITLES = [
    "फोल्डर स्ट्रक्चर व SOP फाइल नामकरण",
    "वाडी विभाजन व उपविभाग पडताळणी (Wadi Vibhajan)",
    "UTM प्रोजेक्शन, GCP वितरण व Affine RMSE (≤ 0.30 m)",
    "ट्रेसिंग नकाशा व HRSI बांध जुळणी (≥ 90% निकष)",
    "क्षेत्रफळ तुलना अहवाल (7/12 ROR vs GIS Area ±3%)",
    "मानक ॲट्रिब्यूट टेबल व अनिवार्य फील्ड्स तपासणी",
    "टोपोलॉजी एरर व OGC जॉमेट्री शुद्धता तपासणी",
    "मिसिंग गट क्रमांक तपासणी (Missing Gat Numbers)",
    "बेस्ट फिट जंक्शन जुळणी (Bi-, Tri- & Quad Junctions)",
    "लगतच्या गावांची सीमा जुळणी (Edge Matching & Seamless Boundary)",
    "स्कॅन नकाशाशी आकार न जुळणारे गट (< 80% जुळणी / मोजणी यादी)",
    "अंतिम Vendor Self-Audit व DSLR सादरीकरण स्वाक्षरी"
]

def pick_pptx():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    p = filedialog.askopenfilename(title="PPT निवडा", filetypes=[("PPTX", "*.pptx")])
    root.destroy()
    return p

def get_slide_title(lines, slide_idx):
    m_idx = slide_idx - 1
    for ln in lines:
        s = re.sub(r"\s+", " ", ln).strip()
        if not s or "गाव:" in s or "LGD:" in s:
            continue
        if re.match(r"^(चेकलिस्ट\s*)?मुद्दा\s*[०-९0-9]+\s*/\s*[०-९0-9]+", s):
            if 1 <= m_idx <= len(QC_12_TITLES):
                return QC_12_TITLES[m_idx - 1]
            continue
        cleaned = re.sub(r"^[0-9०-९]+[\.\)\-:]\s*", "", s).strip()
        if cleaned:
            return cleaned
    if 1 <= m_idx <= len(QC_12_TITLES):
        return QC_12_TITLES[m_idx - 1]
    return f"स्लाइड {slide_idx:02d}"

def read_ppt(path):
    prs = Presentation(path)
    data = []
    for i, s in enumerate(prs.slides, 1):
        txt_list, imgs = [], []
        for sh in s.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip():
                txt_list.append(sh.text_frame.text.strip())
            if sh.shape_type == 13:
                im = Image.open(io.BytesIO(sh.image.blob))
                im.thumbnail((1400, 900))
                b = io.BytesIO()
                im.save(b, "PNG")
                imgs.append(base64.b64encode(b.getvalue()).decode())
        lines = [x.strip() for x in "\n".join(txt_list).splitlines() if x.strip()]
        title = get_slide_title(lines, i)
        body_text = "\n".join(lines[1:]) if len(lines) > 1 else ("\n".join(lines) or "तपशील नाही.")
        data.append({
            "num": f"{i:02d}",
            "title": title,
            "body": body_text,
            "images": imgs
        })
    return data

def make_html(items, fname):
    payload = json.dumps(items, ensure_ascii=False)
    return (
        "<!DOCTYPE html><html lang='mr'><head><meta charset='UTF-8'>"
        "<meta name='viewport' content='width=device-width,initial-scale=1'>"
        "<title>GIS QC डॅशबोर्ड</title><style>"
        "*{box-sizing:border-box}html,body{margin:0;padding:0;width:100%;min-height:100vh;font-family:Segoe UI,Arial,sans-serif;background:#eef2f6;color:#0f172a}"
        ".shell{width:100vw;min-height:100vh;display:flex;flex-direction:column;background:#fff}"
        ".hdr{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;padding:12px 22px;background:#0b192c;color:#fff}"
        ".hdr h1{margin:0;font-size:19px}.sub{font-size:12px;color:#cbd5e1}"
        ".badge{background:rgba(255,255,255,.14);padding:6px 12px;border-radius:6px;font-size:12.5px}"
        ".btn-fs{background:#0f766e;color:#fff;border:0;padding:7px 12px;border-radius:6px;cursor:pointer;font-weight:600}"
        ".grid{flex:1;display:grid;grid-template-columns:410px 1fr;min-height:calc(100vh - 60px)}"
        "@media(max-width:900px){.grid{grid-template-columns:1fr}}"
        ".left{background:#f8fafc;border-right:1px solid #cbd5e1;padding:14px;overflow-y:auto}"
        ".pbtn{width:100%;text-align:left;background:#fff;border:1px solid #cbd5e1;border-radius:8px;padding:10px 12px;margin-bottom:8px;cursor:pointer;display:flex;gap:10px;align-items:center;font-size:13.5px;font-weight:600;color:#0f172a}"
        ".pbtn.active{background:#0f766e;color:#fff;border-color:#0f766e}"
        ".pnum{background:#e2e8f0;color:#0f172a;padding:2px 7px;border-radius:5px;font-size:12px;font-weight:700;flex-shrink:0}"
        ".pbtn.active .pnum{background:rgba(255,255,255,.22);color:#fff}"
        ".right{padding:22px 26px;overflow-y:auto}"
        ".empty{border:1.5px dashed #cbd5e1;border-radius:10px;padding:40px;text-align:center;color:#475569;background:#f8fafc}"
        ".card{display:none}.card.active{display:block}"
        ".chead{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;border-bottom:2px solid #e2e8f0;padding-bottom:12px;margin-bottom:18px}"
        ".chead h2{margin:0;font-size:21px}"
        ".navs button{padding:7px 14px;border:1px solid #cbd5e1;background:#f8fafc;border-radius:6px;cursor:pointer;font-weight:600}"
        ".navs button:disabled{opacity:.4;cursor:not-allowed}"
        ".split{display:grid;grid-template-columns:1fr 1.25fr;gap:22px;align-items:start}"
        "@media(max-width:1100px){.split{grid-template-columns:1fr}}"
        ".box{background:#f8fafc;border:1px solid #cbd5e1;border-left:4px solid #0f766e;border-radius:8px;padding:16px}"
        ".lbl{font-size:12.5px;font-weight:700;color:#0f766e;margin-bottom:8px}"
        ".txt{font-size:15px;line-height:1.7;white-space:pre-line}"
        ".mbox{background:#f8fafc;border:1px solid #cbd5e1;border-radius:8px;padding:14px;text-align:center}"
        ".mbox img{width:100%;max-height:68vh;object-fit:contain;background:#fff;border:1px solid #cbd5e1;border-radius:6px;margin-top:8px;cursor:zoom-in}"
        ".modal{display:none;position:fixed;inset:0;background:rgba(15,23,42,.88);z-index:9999;align-items:center;justify-content:center;padding:20px}"
        ".modal.open{display:flex}.modal img{max-width:96vw;max-height:92vh;border-radius:8px;cursor:zoom-out}"
        "</style></head><body><div class='shell'><div class='hdr'><div><h1>GIS कॅडस्ट्रल QC रिपोर्ट डॅशबोर्ड</h1>"
        "<div class='sub'>डाव्या बाजूच्या टॉपिकवर क्लिक करा — उजव्या बाजूला पूर्ण स्क्रीनमध्ये रिपोर्ट दिसेल</div></div>"
        f"<div style='display:flex;gap:8px;align-items:center'><div class='badge'>{fname} | मुद्दे: {len(items)}</div>"
        "<button class='btn-fs' onclick=\"document.fullscreenElement?document.exitFullscreen():document.documentElement.requestFullscreen()\">⛶ फुल स्क्रीन</button></div></div>"
        "<div class='grid'><aside class='left' id='leftNav'></aside><main class='right'>"
        "<div class='empty' id='emptyBox'>डाव्या बाजूवरून कोणत्याही टॉपिकवर क्लिक करा.</div><div id='rightCards'></div></main></div></div>"
        "<div class='modal' id='modal' onclick=\"this.classList.remove('open')\"><img id='mImg'></div>"
        f"<script>const items={payload};const leftNav=document.getElementById('leftNav'),rightCards=document.getElementById('rightCards'),emptyBox=document.getElementById('emptyBox'),modal=document.getElementById('modal'),mImg=document.getElementById('mImg');"
        "function showIdx(i){emptyBox.style.display='none';document.querySelectorAll('.pbtn').forEach((b,k)=>b.classList.toggle('active',k===i));document.querySelectorAll('.card').forEach((c,k)=>c.classList.toggle('active',k===i))}"
        "items.forEach((it,i)=>{const b=document.createElement('button');b.className='pbtn';b.innerHTML='<span class=\"pnum\">'+it.num+'</span><span>'+it.title+'</span>';b.onclick=()=>showIdx(i);leftNav.appendChild(b);"
        "const c=document.createElement('section');c.className='card';"
        "let imgs='<div style=\"color:#64748b;padding:24px\">[इमेज नाही]</div>';"
        "if(it.images&&it.images.length){imgs=it.images.map(x=>{const u='data:image/png;base64,'+x;return '<img src=\"'+u+'\" onclick=\"mImg.src=\\''+u+'\\';modal.classList.add(\\'open\\')\">'}).join('')+'<div style=\"font-size:12px;color:#64748b;margin-top:6px\">मोठे पाहण्यासाठी इमेजवर क्लिक करा</div>'}"
        "c.innerHTML='<div class=\"chead\"><h2>'+it.num+'. '+it.title+'</h2><div class=\"navs\"><button '+(i===0?'disabled':'')+' onclick=\"showIdx('+(i-1)+')\">◀ मागील</button> <button '+(i===items.length-1?'disabled':'')+' onclick=\"showIdx('+(i+1)+')\">पुढील ▶</button></div></div>"
        "<div class=\"split\"><div class=\"box\"><div class=\"lbl\">अहवाल / मुख्य मुद्दे</div><div class=\"txt\">'+it.body+'</div></div>"
        "<div class=\"mbox\"><div class=\"lbl\" style=\"text-align:left\">स्क्रीनशॉट ('+it.num+')</div>'+imgs+'</div></div>';rightCards.appendChild(c)})</script></body></html>"
    )

p = pick_pptx()
if p and os.path.exists(p):
    items = read_ppt(p)
    out_file = os.path.abspath(os.path.splitext(os.path.basename(p))[0] + "_QC_Dashboard.html")
    with open(out_file, "w", encoding="utf-8") as f:
        f.write(make_html(items, os.path.basename(p)))
    url = "file:///" + out_file.replace("\\", "/")
    webbrowser.open_new_tab(url)
    display(HTML(f"<a href='{url}' target='_blank' style='font-family:Segoe UI;font-weight:700;color:#0f766e'>रिपोर्ट नवीन टॅबमध्ये उघडा ↗</a>"))