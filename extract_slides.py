from pptx import Presentation
from pathlib import Path
from zipfile import ZipFile
import json
src=Path(r'C:\Users\ml\Documents\xwechat_files\wxid_br1bv0s9chxo22_2a69\msg\file\2026-11\项目申报(2).pptx')
out=Path('ppt_assets'); out.mkdir(exist_ok=True)
prs=Presentation(str(src))
with ZipFile(src) as archive:
    for item in archive.namelist():
        if item.startswith('ppt/media/media') and item.lower().endswith('.mp4'):
            (out/'robomaster-demo.mp4').write_bytes(archive.read(item))
slides=[]
for i,slide in enumerate(prs.slides,1):
    blocks=[]
    for shape in slide.shapes:
        if shape.has_text_frame:
            text='\n'.join(p.text for p in shape.text_frame.paragraphs).strip()
            if text: blocks.append({'type':'text','text':text,'x':round(shape.left/prs.slide_width,4),'y':round(shape.top/prs.slide_height,4),'w':round(shape.width/prs.slide_width,4),'h':round(shape.height/prs.slide_height,4)})
        if shape.shape_type==13:
            try:
                image=shape.image
                name=f'slide-{i}-image-{len([b for b in blocks if b["type"]=="image"])+1}.{image.ext}'
                (out/name).write_bytes(image.blob)
                blocks.append({'type':'image','src':'ppt_assets/'+name,'x':round(shape.left/prs.slide_width,4),'y':round(shape.top/prs.slide_height,4),'w':round(shape.width/prs.slide_width,4),'h':round(shape.height/prs.slide_height,4)})
            except ValueError:
                pass
    slides.append({'number':i,'blocks':blocks})
Path('slides.json').write_text(json.dumps({'width':prs.slide_width,'height':prs.slide_height,'slides':slides},ensure_ascii=False,indent=2),encoding='utf-8')
print('slides',len(slides),'texts',sum(sum(b['type']=='text' for b in s['blocks']) for s in slides),'images',sum(sum(b['type']=='image' for b in s['blocks']) for s in slides))
