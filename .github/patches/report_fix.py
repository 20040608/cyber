from pathlib import Path
base=Path('Tekshiruv_API_tayyor')
p=base/'desktop_webview.py'
s=p.read_text(encoding='utf-8')
bridge='''class DesktopReportBridge:
    """Native Save As dialog for Word reports."""
    def save_report(self, check_id):
        import re
        from inspector.store import get
        from inspector.report import generate_docx
        from inspector.download_naming import report_filename
        try:
            if not isinstance(check_id,str) or not re.fullmatch(r'[a-f0-9]{32}',check_id):
                return {'ok':False,'error':'Tekshiruv ID noto‘g‘ri.'}
            record=get(check_id)
            if not record:return {'ok':False,'error':'Natija tarixdan topilmadi.'}
            data=generate_docx(record,mode='automatic')
            if not data.startswith(b'PK'):return {'ok':False,'error':'Word hujjat yaratilmagan.'}
            import webview
            selected=webview.windows[0].create_file_dialog(webview.SAVE_DIALOG,
                save_filename=report_filename(record,'docx'),
                file_types=('Word documents (*.docx)',))
            if not selected:return {'ok':False,'cancelled':True}
            target=Path(str(selected[0] if isinstance(selected,(tuple,list)) else selected))
            if target.suffix.lower()!='.docx':target=target.with_name(target.name+'.docx')
            target.write_bytes(data)
            return {'ok':True,'path':str(target)}
        except Exception as exc:
            write_error('Word saqlash xatosi',exc)
            return {'ok':False,'error':'Word saqlanmadi. Boshqa papkani tanlang.'}


'''
assert 'class DesktopReportBridge:' not in s
s=s.replace('def main():\n',bridge+'def main():\n',1)
assert 'min_size=(1000, 660), text_select=True)' in s
s=s.replace('min_size=(1000, 660), text_select=True)','min_size=(1000, 660), text_select=True, js_api=DesktopReportBridge())')
p.write_text(s,encoding='utf-8')
p=base/'static/app.js'
s=p.read_text(encoding='utf-8')
start=s.index('async function downloadExport(')
end=s.index("$('word-btn').onclick",start)
s=s[:start]+'''async function downloadExport(url,ext,record=latest){
 const id=record?.id;if(!id){alert('Avval natijani tanlang.');return}
 try{
  if(window.pywebview?.api?.save_report){
   const result=await window.pywebview.api.save_report(id);
   if(result?.cancelled)return;
   if(!result?.ok)throw Error(result?.error||'Word xulosasi saqlanmadi');
   alert('Word xulosasi saqlandi: '+result.path);return;
  }
  const response=await fetch(url,{cache:'no-store'});
  if(!response.ok)throw Error((await response.text()).slice(0,160));
  const blob=await response.blob();if(!blob.size)throw Error('Hisobot bo‘sh.');
  const a=document.createElement('a'),u=URL.createObjectURL(blob);
  a.href=u;a.download=serverDownloadName(response,ext,record);
  document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(u),60000);
 }catch(err){alert('Xulosani yuklab olishda xato: '+err.message)}
}
'''+s[end:]
start=s.index("const r=await fetch(",s.index("[data-h-word]"))
end=s.index("setTimeout(()=>URL.revokeObjectURL(u),2000)",start)+len("setTimeout(()=>URL.revokeObjectURL(u),2000)")
assert 'report' in s[start:end] and 'blob' in s[start:end]
replacement="await downloadExport("+chr(96)+"/api/checks/"+chr(36)+"{encodeURIComponent(id)}/report"+chr(96)+",'docx',detail)"
s=s[:start]+replacement+s[end:]
p.write_text(s,encoding='utf-8')
print('Applied report saving fix')
