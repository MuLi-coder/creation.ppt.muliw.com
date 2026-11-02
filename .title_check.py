from pptx import Presentation
p=Presentation(r'C:\Users\ml\Documents\xwechat_files\wxid_br1bv0s9chxo22_2a69\msg\file\2026-11\项目申报(2).pptx')
for i in range(2,6):
 print('slide',i+1)
 for s in p.slides[i].shapes:
  if s.has_text_frame and s.text.strip(): print(repr(s.text[:50]),*[round(v/1e6,3) for v in (s.left,s.top,s.width,s.height)])
