let songs=[], selectedId=null, saveTimer=null, saveChain=Promise.resolve(), previewUrl=null, dragId=null;
const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const current=()=>songs.find(s=>s.id===selectedId);

async function api(path,options={}){
  const response=await fetch(path,{cache:'no-store',...options});
  if(!response.ok){let message=`Request failed (${response.status})`;
    try{const data=await response.json();message=data.error||message}catch(_){ }
    throw new Error(message);
  }
  const type=response.headers.get('content-type')||'';
  return type.includes('application/json')?response.json():response.blob();
}
function toast(message,isError=false){
  const el=$('toast');el.textContent=message;el.className='visible'+(isError?' error':'');
  clearTimeout(el._timeout);el._timeout=setTimeout(()=>el.className='',4500);
}
function count(){const n=songs.filter(s=>s.included).length;
  $('songCount').textContent=String(songs.length);
  $('exportCount').textContent=`${n} ${n===1?'song':'songs'} selected`;
  $('exportPpt').disabled=$('exportPdf').disabled=n===0;
}
function renderList(){
  $('songList').innerHTML=songs.map((s,i)=>`<div class="songrow ${s.id===selectedId?'active':''}" data-id="${esc(s.id)}" draggable="true" title="Drag to reorder">
    <span class="grip" aria-hidden="true">⋮⋮</span>
    <input type="checkbox" class="included" ${s.included?'checked':''} aria-label="Include ${esc(s.title)} in setlist">
    <div class="song-main"><b>${esc(s.title)||'Untitled Song'}</b><small>${esc(s.artist)||'No artist yet'}</small></div>
    <div class="reorder"><button class="up" title="Move up" aria-label="Move up" ${i===0?'disabled':''}>▲</button><button class="down" title="Move down" aria-label="Move down" ${i===songs.length-1?'disabled':''}>▼</button></div>
  </div>`).join('')||'<p class="side-help">No songs yet. Click + to add one.</p>';
  count();
}
function fillEditor(){
  const s=current(),present=!!s;
  for(const id of ['titleInput','artistInput','contentInput','deleteBtn','duplicateBtn'])$(id).disabled=!present;
  $('editorHeading').textContent=present?s.title:'Select or add a song';
  $('titleInput').value=present?s.title:'';
  $('artistInput').value=present?s.artist:'';
  $('contentInput').value=present?s.content:'';
  $('saveState').textContent=present?'Saved locally':'';
  if(present)updatePreview();
  else{if(previewUrl)URL.revokeObjectURL(previewUrl);previewUrl=null;$('preview').removeAttribute('src');setPreviewMessage('Add a song to see its chart.');}
}
function setPreviewMessage(text,error=false){
  const el=$('previewMessage');
  el.textContent=text;el.className='preview-message'+(error?' error':'')+(text?'':' hidden');
}
async function updatePreview(){
  const id=selectedId;if(!id)return;
  setPreviewMessage('Rendering chart…');
  try{
    const blob=await api(`/api/preview/${encodeURIComponent(id)}.png?time=${Date.now()}`);
    if(selectedId!==id)return;
    const url=URL.createObjectURL(blob);
    if(previewUrl)URL.revokeObjectURL(previewUrl);
    previewUrl=url;$('preview').src=url;setPreviewMessage('');
  }catch(e){
    if(selectedId===id){$('preview').removeAttribute('src');setPreviewMessage(e.message,true)}
  }
}
function queueSave(){
  const s=current();if(!s)return;
  s.title=$('titleInput').value;s.artist=$('artistInput').value;s.content=$('contentInput').value;
  $('editorHeading').textContent=s.title||'Untitled Song';
  $('saveState').textContent='Saving…';
  renderList();
  clearTimeout(saveTimer);
  saveTimer=setTimeout(()=>persistNow(),550);
}
function persistNow(){
  clearTimeout(saveTimer);saveTimer=null;
  const s=current();if(!s)return saveChain;
  const snapshot=JSON.parse(JSON.stringify(s));
  saveChain=saveChain.catch(()=>{}).then(async()=>{
    await api(`/api/song/${encodeURIComponent(snapshot.id)}`,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(snapshot)});
    if(selectedId===snapshot.id){$('saveState').textContent='Saved locally';await updatePreview()}
  }).catch(e=>{toast(`Could not save: ${e.message}`,true);$('saveState').textContent='Save error';});
  return saveChain;
}
async function flushSave(){if(saveTimer!==null)await persistNow();else await saveChain;}
async function select(id){if(selectedId===id)return;await flushSave();selectedId=id;renderList();fillEditor();}
async function addSong(copy=null){
  await flushSave();
  try{
    const song=await api('/api/songs',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify(copy||{title:'New Song',artist:'',included:true,content:'[Verse 1]\nD5 | First line of the verse\n\n[Chorus]\nBOLD: Guitar cue goes here\nFirst line of the chorus\n\n[Outro]\nBOLD: Ending / stop cue'})});
    songs.push(song);await select(song.id);$('titleInput').focus();$('titleInput').select();
  }catch(e){toast(e.message,true)}
}
async function removeSong(){const s=current();if(!s)return;
  if(!confirm(`Delete "${s.title}" from your library? This cannot be undone (unless you saved a backup).`))return;
  await flushSave();
  try{await api(`/api/song/${encodeURIComponent(s.id)}`,{method:'DELETE'});
    const idx=songs.findIndex(x=>x.id===s.id);songs.splice(idx,1);
    selectedId=songs[Math.min(idx,songs.length-1)]?.id||null;renderList();fillEditor();
  }catch(e){toast(e.message,true)}
}
async function setOrder(){
  try{await api('/api/order',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({ids:songs.map(s=>s.id)})})}
  catch(e){toast(e.message,true)}
}
async function move(id,delta){await flushSave();let i=songs.findIndex(s=>s.id===id);
  if(i<0||i+delta<0||i+delta>=songs.length)return;
  [songs[i],songs[i+delta]]=[songs[i+delta],songs[i]];renderList();await setOrder();
}
async function download(format){
  try{
    await flushSave();
    const button=$(format==='pptx'?'exportPpt':'exportPdf');
    button.disabled=true;const original=button.textContent;button.textContent='Generating…';
    try{
      const blob=await api(`/api/export/${format}`);
      const url=URL.createObjectURL(blob);const a=document.createElement('a');
      a.href=url;a.download=`live-show-charts.${format}`;document.body.append(a);a.click();a.remove();
      setTimeout(()=>URL.revokeObjectURL(url),30000);
      toast(`${songs.filter(s=>s.included).length}-song ${format.toUpperCase()} exported.`);
    }finally{button.textContent=original;count()}
  }catch(e){toast(e.message,true)}
}
async function backup(){
  try{await flushSave();const blob=await api('/api/backup');
    const url=URL.createObjectURL(blob);const a=document.createElement('a');
    a.href=url;a.download='song-library-backup.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),30000);
    toast('Song library backup downloaded.');
  }catch(e){toast(e.message,true)}
}
async function importJson(file){if(!file)return;
  await flushSave();
  try{const data=JSON.parse(await file.text());
    if(!Array.isArray(data))throw new Error('Backup should contain a JSON array of songs.');
    if(!confirm(`Replace this entire library with ${data.length} imported songs? Export a backup first if needed.`))return;
    await api('/api/restore',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
    await initialize();toast(`Imported ${data.length} songs.`);
  }catch(e){toast(`Import error: ${e.message}`,true)}
  finally{$('importFile').value=''}
}
async function initialize(){
  try{songs=await api('/api/songs');selectedId=songs.find(s=>s.id===selectedId)?.id||songs[0]?.id||null;
    renderList();fillEditor();
  }catch(e){toast(`Unable to load library: ${e.message}`,true)}
}
$('addBtn').addEventListener('click',()=>addSong());
$('duplicateBtn').addEventListener('click',()=>{const s=current();if(s)addSong({...s,title:s.title+' (copy)'})});
$('deleteBtn').addEventListener('click',removeSong);
for(const id of ['titleInput','artistInput','contentInput'])$(id).addEventListener('input',queueSave);
$('exportPpt').addEventListener('click',()=>download('pptx'));
$('exportPdf').addEventListener('click',()=>download('pdf'));
$('backupBtn').addEventListener('click',backup);
$('importFile').addEventListener('change',e=>importJson(e.target.files[0]));
$('songList').addEventListener('click',async e=>{
  const row=e.target.closest('.songrow');if(!row)return;
  const id=row.dataset.id;
  if(e.target.matches('input.included')){
    const s=songs.find(s=>s.id===id);s.included=e.target.checked;count();
    try{await api(`/api/song/${encodeURIComponent(id)}`,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(s)})}
    catch(err){toast(err.message,true)}
  }else if(e.target.closest('.up'))await move(id,-1);
  else if(e.target.closest('.down'))await move(id,1);
  else await select(id);
});
$('songList').addEventListener('dragstart',e=>{
  const row=e.target.closest('.songrow');if(!row)return;
  dragId=row.dataset.id;row.classList.add('dragging');e.dataTransfer.effectAllowed='move';
});
$('songList').addEventListener('dragover',e=>{
  const row=e.target.closest('.songrow');if(!row||!dragId)return;
  e.preventDefault();document.querySelectorAll('.drop-target').forEach(el=>el.classList.remove('drop-target'));
  row.classList.add('drop-target');
});
$('songList').addEventListener('drop',async e=>{
  e.preventDefault();const row=e.target.closest('.songrow');
  document.querySelectorAll('.drop-target,.dragging').forEach(el=>el.classList.remove('drop-target','dragging'));
  if(!row||!dragId||row.dataset.id===dragId){dragId=null;return}
  await flushSave();const from=songs.findIndex(s=>s.id===dragId);
  let to=songs.findIndex(s=>s.id===row.dataset.id);
  if(from>=0&&to>=0){const [s]=songs.splice(from,1);to=songs.findIndex(s=>s.id===row.dataset.id);songs.splice(to,0,s);renderList();await setOrder()}
  dragId=null;
});
$('songList').addEventListener('dragend',()=>{dragId=null;document.querySelectorAll('.drop-target,.dragging').forEach(el=>el.classList.remove('drop-target','dragging'))});
document.addEventListener('keydown',e=>{
  if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='s'){e.preventDefault();persistNow();toast('Saving chart…')}
});
initialize();
