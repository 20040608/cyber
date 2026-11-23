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

# Show evidence-based findings in original result card, instead of only counters.
html=base/'static/index.html'
layout=html.read_text(encoding='utf-8')
anchor='    <div class="action-row">'
panel='''    <div id="findings-panel" aria-live="polite" style="padding:14px 18px;line-height:1.5"></div>
'''
assert anchor in layout
layout=layout.replace(anchor,panel+anchor,1)
html.write_text(layout,encoding='utf-8')
p=base/'static/app.js'
js=p.read_text(encoding='utf-8')
assert "function renderFileRows(){" in js
js=js.replace("function renderFileRows(){",'''function renderFindings(r){
 const panel=$('findings-panel');if(!panel)return;
 const overview=r.automated_explanation||{}, defects=overview.defects||[];
 const unverified=(r.issues||[]).filter(x=>x.level==='review'&&!x.resident_actionable);
 const head='<h3 style="margin:0 0 8px">Aniqlangan kamchiliklar va tavsiyalar</h3>';
 const items=defects.map((x,i)=>'<li style="margin:0 0 12px"><strong>'+esc(x.title||'Kamchilik')+'</strong><div>'+esc(x.evidence||'Dalil keltirilmagan')+'</div><div><b>Tuzatish:</b> '+esc(x.fix||'Qo‘shimcha tekshirish kerak')+'</div>'+(x.file?'<small>Fayl: '+esc(x.file)+'</small>':'')+'</li>').join('');
 const caution='<p style="margin:8px 0;font-size:13px">Qo‘lda tekshirish talab qilinadigan holatlar: '+unverified.length+'. Texnik tekshiruv rezidentlikni tasdiqlamaydi.</p>';
 panel.innerHTML=head+(items?'<ol style="margin:0;padding-left:24px">'+items+'</ol>':'<p style="margin:0">Hozircha dalil bilan tasdiqlangan kamchilik aniqlanmadi.</p>')+caution;
}
function renderFileRows(){''',1)
needle="$('word-btn').disabled=false;renderFileRows()}"
assert needle in js
js=js.replace(needle,"$('word-btn').disabled=false;renderFileRows();renderFindings(r)}",1)
p.write_text(js,encoding='utf-8')
# Add native-bridge regression tests to the standard suite.
test=base/'tests/test_native_word_export.py'
test.write_text('''import io
from pathlib import Path
from unittest.mock import MagicMock, patch
from zipfile import ZipFile
from docx import Document
from desktop_webview import DesktopReportBridge


def test_native_report_save_and_cancel(tmp_path):
    import webview
    record = {
        'id': 'a' * 32, 'company_name': 'Cyber Park Demo',
        'language': 'uz_lat', 'files_count': 1, 'passed_files': 0,
        'issues': [], 'documents': [], 'status': 'review',
    }
    output = tmp_path / 'sample.docx'
    window = MagicMock()
    window.create_file_dialog.return_value = str(output)
    with patch('inspector.store.get', return_value=record), patch.object(webview,'windows',[window]):
        good = DesktopReportBridge().save_report(record['id'])
        assert good['ok'] is True and output.exists()
        assert ZipFile(output).testzip() is None
        doc = Document(output)
        assert 'Cyber Park Demo' in '\\\\n'.join(p.text for p in doc.paragraphs)
        window.create_file_dialog.return_value = None
        assert DesktopReportBridge().save_report(record['id']).get('cancelled') is True
    assert DesktopReportBridge().save_report('../bad')['ok'] is False
''',encoding='utf-8')
