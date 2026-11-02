from pptx import Presentation
from pathlib import Path
from zipfile import ZipFile
p=Presentation(r'C:\Users\ml\Documents\xwechat_files\wxid_br1bv0s9chxo22_2a69\msg\file\2026-11\项目申报(2).pptx')
for i,s in enumerate(p.slides,1):
 for sh in s.shapes:
  if sh.shape_type==13:
   try:
    im=sh.image
    print(i,sh.name,im.ext,round(sh.left/p.slide_width,2),round(sh.top/p.slide_height,2),round(sh.width/p.slide_width,2),round(sh.height/p.slide_height,2))
   except Exception: pass
with ZipFile(r'C:\Users\ml\Documents\xwechat_files\wxid_br1bv0s9chxo22_2a69\msg\file\2026-11\项目申报(2).pptx') as z:
 for n in z.namelist():
  if n.startswith('ppt/media/'):
   inf=z.getinfo(n); print('MEDIA',n,inf.file_size)
