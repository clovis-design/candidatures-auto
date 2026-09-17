const $ = id => document.getElementById(id);
const escapeHtml = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const dropzone = $('dropzone');
const cvInput = $('cv');
const cvInfo = $('cvInfo');

dropzone.addEventListener('click', ()=> cvInput.click());
cvInput.addEventListener('click', e => e.stopPropagation());
dropzone.addEventListener('keydown', e => { if(e.key === 'Enter' || e.key === ' '){ e.preventDefault(); cvInput.click(); } });
dropzone.addEventListener('dragover', e=>{e.preventDefault(); dropzone.classList.add('drag')});
dropzone.addEventListener('dragleave', ()=> dropzone.classList.remove('drag'));
dropzone.addEventListener('drop', e=>{
  e.preventDefault(); dropzone.classList.remove('drag');
  if(e.dataTransfer.files[0]){ cvInput.files = e.dataTransfer.files; updateCVInfo(); }
});
cvInput.addEventListener('change', updateCVInfo);
function updateCVInfo(){
  if(cvInput.files[0]){
    const f=cvInput.files[0];
    if(!/\.(pdf|docx?)$/i.test(f.name) || f.size > 10 * 1024 * 1024){
      cvInput.value=''; cvInfo.textContent='Choisissez un PDF, DOC ou DOCX de moins de 10 Mo.'; return;
    }
    cvInfo.textContent = `✓ ${f.name} (${(f.size/1024).toFixed(1)} Ko)`;
    document.dispatchEvent(new Event('workspacechange'));
  }
}

// Emails comptage + Excel entries (tableau ligne 4+)
const emailsText = $('emails_text');
const emailCount = $('emailCount');
const invalidCount = $('invalidCount');
let debounce;
let excelEntries = null; // stocke les lignes du tableau Excel importé (ligne 5+ avec stage etc.)
emailsText.addEventListener('input', ()=>{
  // Si l'utilisateur édite manuellement après un import Excel, on invalide le tableau enrichi
  if(excelEntries){ excelEntries=null; $('excelInfo').textContent='ℹ️ Liste modifiée manuellement — le tableau Excel enrichi a été réinitialisé. Ré-importez ou régénérez le tableau.'; $('excelPreview').classList.add('hidden'); }
  clearTimeout(debounce);
  debounce=setTimeout(parseEmails, 400);
});
$('btnClean').addEventListener('click', ()=>{
  // Nettoie : une entrée par ligne, trim
  const lines = emailsText.value.split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
  const seen=new Set(); const out=[];
  const re=/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/;
  for(const l of lines){
    const m=l.match(re);
    const key=m?m[0].toLowerCase():l;
    if(!seen.has(key)){ seen.add(key); out.push(l); }
  }
  emailsText.value = out.join('\n');
  // le tableau Excel enrichi n'est plus valide après nettoyage manuel
  if(excelEntries){ excelEntries=null; $('excelPreview').classList.add('hidden'); $('excelInfo').textContent='ℹ️ Liste nettoyée — tableau réinitialisé.'; }
  parseEmails();
});
$('csvImport').addEventListener('change', async e=>{
  const file=e.target.files[0];
  if(!file) return;
  const ext=file.name.toLowerCase().split('.').pop();
  // Si Excel -> envoi au backend parse-excel (headers ligne 4 -> données ligne 5+)
  if(['xlsx','xlsm','csv'].includes(ext)){
    $('excelInfo').textContent='Lecture du tableur et détection des colonnes…';
    const fd=new FormData(); fd.append('file', file);
    try{
      let res;
      try{
        res=await fetch('/api/parse-excel',{method:'POST', body:fd});
      }catch(netErr){
        throw new Error('Serveur injoignable (Failed to fetch). Lance ./start.sh puis ouvre http://localhost:5000.');
      }
      let data;
      try{ data=await res.json(); }
      catch(e){ throw new Error('Réponse serveur invalide (HTTP '+res.status+'). Redémarre le serveur avec ./start.sh.'); }
      if(!res.ok) throw new Error(data.error||('Erreur '+res.status));
      // Remplit la liste d'emails à partir du tableau (le tableau se remplit depuis la liste d'emails)
      emailsText.value = data.emails_text || data.emails.join('\n');
      excelEntries = data.entries; // garde le tableau enrichi (stage etc.) pour l'envoi
      if(data.entries.some(en=>en.entreprise)) $('useCustomNames').checked=true;
      updateCustomNamesUI();
      parseEmails();
      // Affiche preview tableau
      renderExcelPreview(data.entries, data.invalid_rows);
      $('excelInfo').textContent=`✓ ${data.count} candidatures importées. Vous pouvez ajuster les entreprises et postes, ou décocher des destinataires. ` + data.invalid_rows.map(r => `Ligne ${r.row} : ${r.reason}`).join(' • ');
      document.dispatchEvent(new Event('workspacechange'));
    }catch(err){
      $('excelInfo').textContent='❌ '+err.message;
    }
    return;
  }
  emailsText.value = await file.text();
  emailsText.dispatchEvent(new Event('input', {bubbles:true}));
});

function renderExcelPreview(entries, invalid_rows){
  const wrap=$('excelPreview');
  const tbl=$('previewTable');
  const stats=$('previewStats');
  if(!entries.length){ wrap.classList.add('hidden'); return; }
  wrap.classList.remove('hidden');
  const headers=["Inclure","Entreprise","Email","Poste ciblé","Interlocuteur","Date d’envoi / suivi"];
  let html='<thead><tr>'+headers.map(h=>`<th>${h}</th>`).join('')+'</tr></thead><tbody>';
  entries.forEach((en, index)=>{
    html+=`<tr data-row="${index}" class="${en.excluded?'excluded':''}"><td><input type="checkbox" data-key="selected" aria-label="Inclure ${escapeHtml(en.email)}" ${en.excluded?'':'checked'}></td><td><input data-key="entreprise" aria-label="Entreprise ligne ${index+1}" value="${escapeHtml(en.entreprise)}"></td><td>${escapeHtml(en.email)}</td><td><input data-key="stage" aria-label="Poste ligne ${index+1}" value="${escapeHtml(en.stage)}" placeholder="Poste du profil"></td><td><input data-key="interlocuteur" aria-label="Interlocuteur ligne ${index+1}" value="${escapeHtml(en.interlocuteur)}"></td><td><input data-key="date_envoi" aria-label="Date d’envoi ligne ${index+1}" value="${escapeHtml(en.date_envoi)}" placeholder="Pas encore envoyée"><small>${escapeHtml(en.resultats||'À envoyer')}</small></td></tr>`;
  });
  html+=`</tbody>`;
  tbl.innerHTML=html;
  let s=`${entries.filter(e=>!e.excluded).length} candidature(s) sélectionnée(s) sur ${entries.length}`;
  if(invalid_rows && invalid_rows.length) s+=` • ${invalid_rows.length} ligne(s) invalide(s) (pas d'email en colonne Coordonnées)`;
  stats.textContent=s;
}

$('previewTable').addEventListener('input', event => {
  const input=event.target;
  if(!input.dataset.key || !excelEntries) return;
  const row=input.closest('tr');
  const entry=excelEntries[Number(row.dataset.row)];
  if(input.dataset.key==='selected') entry.excluded=!input.checked;
  else entry[input.dataset.key]=input.value;
  row.classList.toggle('excluded', !!entry.excluded);
  $('previewStats').textContent=`${excelEntries.filter(e=>!e.excluded).length} candidature(s) sélectionnée(s) sur ${excelEntries.length}`;
  document.dispatchEvent(new Event('workspacechange'));
});

// Boutons Excel : modèle vide + génération depuis liste d'emails (remplit le tableau ligne 5+)
$('btnTemplateExcel')?.addEventListener('click', ()=>{ window.location.href='/api/excel-template'; });
$('btnGenerateExcel')?.addEventListener('click', async ()=>{
  const text=emailsText.value.trim();
  if(!text){ alert('Ajoute d\'abord une liste d\'emails (une par ligne)'); return; }
  const poste=$('poste').value.trim();
  const btn=$('btnGenerateExcel');
  btn.disabled=true; btn.textContent='⏳ Génération...';
  $('excelInfo').textContent='⏳ Génération du tableau (headers ligne 4, données ligne 5+) à partir de ta liste d\'emails...';
  try{
    let res;
    try{
      res=await fetch('/api/generate-excel',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({emails_text:text, poste, ...(excelEntries ? {entries:excelEntries.filter(e=>!e.excluded)} : {})})});
    }catch(netErr){
      throw new Error('Serveur injoignable (Failed to fetch). Vérifie que le serveur tourne : lance ./start.sh puis ouvre http://localhost:5000 (ne pas ouvrir index.html en file://).');
    }
    if(!res.ok){
      let msg='Erreur '+res.status;
      try{ const j=await res.json(); msg=j.error||msg; }
      catch(e){ try{ msg=await res.text(); }catch(e2){} msg=msg.slice(0,300); }
      throw new Error(msg);
    }
    const blob=await res.blob();
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a'); a.href=url; a.download=`suivi_candidatures_${new Date().toISOString().slice(0,10)}.xlsx`; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
    $('excelInfo').textContent='✓ Liste exportée. Les dates d’envoi restent vides tant que les candidatures ne sont pas envoyées. Le suivi réel se télécharge dans « Mes candidatures ».';
  }catch(err){ $('excelInfo').textContent='❌ '+err.message; }
  finally{ btn.disabled=false; btn.textContent='Exporter cette liste en Excel'; }
});

async function parseEmails(){
  const text=emailsText.value;
  if(!text.trim()){ emailCount.textContent='0 email valide'; invalidCount.textContent=''; return;}
  try{
    const res=await fetch('/api/parse-emails',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({text})});
    const data=await res.json();
    const useCustom=$('useCustomNames')?.checked;
    let label = `${data.count} email${data.count>1?'s':''} valide${data.count>1?'s':''}`;
    if(useCustom && data.entries){
      const withName = data.entries.filter(e=>e.entreprise).length;
      if(withName) label += ` • ${withName} avec nom saisi`;
    }
    emailCount.textContent = label;
    document.dispatchEvent(new Event('workspacechange'));
    invalidCount.textContent = data.invalid.length? ` • ${data.invalid.length} invalide(s): ${data.invalid.slice(0,3).join(', ')}` : '';
  }catch(e){}
}
// Toggle noms personnalisés - UI améliorée
function updateCustomNamesUI(){
  const chk=$('useCustomNames');
  const status=$('customNamesStatus');
  const desc=$('customNamesDesc');
  const card=$('customNamesCard');
  const on=chk.checked;
  status.textContent=on?'ON':'OFF';
  status.className='toggle-status '+(on?'on':'off');
  desc.textContent=on?'Mode manuel : utilise ton nom saisi (ex: toto@gmail.com | SuperCorp)' : 'Mode auto : nom extrait du domaine';
  card.style.borderColor=on?'#c7d2fe':'var(--border)';
  card.style.background=on?'#eef2ff':'#fbfdff';
  // highlight example
  const exOn=$('exOn'), exOff=$('exOff');
  if(exOn) exOn.style.opacity=on?'1':'.55';
  if(exOff) exOff.style.opacity=on?'.55':'1';
}
document.getElementById('useCustomNames')?.addEventListener('change', ()=>{ updateCustomNamesUI(); parseEmails(); });
document.getElementById('btnHelpCustom')?.addEventListener('click', ()=>{
  $('customNamesHelp').classList.toggle('hidden');
});
// init
updateCustomNamesUI();

// --- Mode lettre : Auto vs Ma propre lettre ---
let letterMode = 'auto'; // 'auto' | 'custom'
function setLetterMode(mode){
  letterMode = mode;
  $('modeAuto').classList.toggle('active', mode==='auto');
  $('modeCustom').classList.toggle('active', mode==='custom');
  $('autoPanel').classList.toggle('hidden', mode!=='auto');
  $('customPanel').classList.toggle('hidden', mode!=='custom');
  $('modeLabel').textContent = mode==='auto' ? 'Auto' : 'Ma propre lettre';
  $('modeAuto').setAttribute('aria-pressed', mode==='auto');
  $('modeCustom').setAttribute('aria-pressed', mode==='custom');
  // Si on passe en custom et le textarea est vide, on affiche la zone de preview pour écrire
  if(mode==='custom'){
    $('letterPreview').classList.remove('hidden');
    if(!$('lettre_template').value.trim()){
      $('previewText').textContent='Écris ou colle ta lettre ci-dessous avec les annotations {ENTREPRISE_NOM}, {POSTE}... puis clique « Prévisualiser ».';
    }
  }
}
$('modeAuto')?.addEventListener('click', ()=> setLetterMode('auto'));
$('modeCustom')?.addEventListener('click', ()=> setLetterMode('custom'));

// Insertion placeholders dans la lettre au curseur
document.querySelectorAll('.chip[data-insert]')?.forEach(btn=>{
  btn.addEventListener('click', ()=>{
    const ta=$('lettre_template');
    const ins=btn.dataset.insert;
    const start=ta.selectionStart ?? ta.value.length;
    const end=ta.selectionEnd ?? ta.value.length;
    ta.value = ta.value.slice(0,start) + ins + ta.value.slice(end);
    ta.focus();
    ta.selectionStart = ta.selectionEnd = start + ins.length;
    $('letterPreview').classList.remove('hidden');
  });
});

// Import .txt dans lettre custom
$('customFile')?.addEventListener('change', e=>{
  const f=e.target.files[0];
  if(!f) return;
  const r=new FileReader();
  r.onload=()=>{
    $('lettre_template').value = r.result;
    $('letterPreview').classList.remove('hidden');
    setLetterMode('custom');
  };
  r.readAsText(f);
});

// Exemple de lettre perso avec annotations
$('btnUseExample')?.addEventListener('click', ()=>{
  const p=$('prenom').value.trim()||'[Prénom]';
  const n=$('nom').value.trim()||'[Nom]';
  const poste=$('poste').value.trim()||'le stage visé';
  $('lettre_template').value =
`${p} ${n}

À l'attention du service Recrutement
{ENTREPRISE_NOM}

Objet : Candidature au poste de ${poste}

Madame, Monsieur,

C'est avec un vif intérêt que je vous adresse ma candidature au poste de {POSTE} au sein de {ENTREPRISE_NOM}.

[Écris ici ton paragraphe d'expérience : diplôme, compétences, projets...]

Motivé(e) par les missions de {ENTREPRISE}, je serais ravi(e) d'échanger avec vous lors d'un entretien.

Je vous prie d'agréer, Madame, Monsieur, l'expression de mes salutations distinguées.

{PRENOM} {NOM}`;
  $('letterPreview').classList.remove('hidden');
});

// Prévisualisation lettre PERSO (sans régénérer) via /api/preview-letter
$('btnPreviewCustom')?.addEventListener('click', async ()=>{
  const text=$('lettre_template').value;
  if(!text.trim()){ alert('Écris d\'abord ta lettre dans la zone de texte.'); return; }
  const btn=$('btnPreviewCustom');
  btn.disabled=true; btn.textContent='⏳ Aperçu...';
  try{
    const res=await fetch('/api/preview-letter',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({
      text,
      prenom:$('prenom').value.trim(), nom:$('nom').value.trim(),
      ville:$('ville').value.trim(), telephone:$('telephone').value.trim(),
      emailCandidat:$('emailCandidat').value.trim(),
      poste:$('poste').value.trim()||'le poste visé',
      entrepriseExemple:'ExempleCorp'
    })});
    const j=await res.json();
    if(!res.ok) throw new Error(j.error||'Erreur');
    $('letterPreview').classList.remove('hidden');
    $('previewText').textContent=j.preview+'\n\n— (aperçu avec ExempleCorp, à l\'envoi chaque entreprise aura son nom)';
    $('previewMode').textContent='(ma lettre)';
    $('pdfLink').href=j.pdf_url;
    $('letterPreview').scrollIntoView({behavior:'smooth', block:'center'});
  }catch(err){ alert('Erreur aperçu: '+err.message); }
  finally{ btn.disabled=false; btn.textContent='👁️ Prévisualiser ma lettre (PDF)'; }
});

// Génération lettre AUTO (feature existante conservée)
$('btnGenerate').addEventListener('click', async ()=>{
  const data=collectFormData();
  if(!data.prenom || !data.nom || !data.poste){
    alert('Veuillez remplir Prénom, Nom et Poste visé.');
    return;
  }
  const btn=$('btnGenerate');
  btn.disabled=true; btn.textContent='⏳ Génération...';
  try{
    const res=await fetch('/api/generate-letter',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
    const j=await res.json();
    if(!res.ok) throw new Error(j.error||'Erreur');
    $('letterPreview').classList.remove('hidden');
    $('previewText').textContent=j.preview;
    $('lettre_template').value=j.template;
    document.dispatchEvent(new Event('workspacechange'));
    $('previewMode').textContent='(auto)';
    $('pdfLink').href=j.pdf_url;
    document.querySelector('.step[data-step="2"]').classList.add('active');
    // scroll
    $('letterPreview').scrollIntoView({behavior:'smooth', block:'center'});
  }catch(e){
    alert('Erreur génération: '+e.message);
  }finally{ btn.disabled=false; btn.textContent='⚡ Générer / Prévisualiser la lettre'; }
});

function collectFormData(){
  return {
    prenom: $('prenom').value.trim(),
    nom: $('nom').value.trim(),
    emailCandidat: $('emailCandidat').value.trim(),
    telephone: $('telephone').value.trim(),
    ville: $('ville').value.trim(),
    poste: $('poste').value.trim(),
    disponibilite: $('disponibilite').value.trim(),
    ton: $('ton').value,
    experience: $('experience').value.trim(),
    motivation: $('motivation').value.trim(),
    personnalise: $('personnalise').checked,
    entreprisePlaceholder: 'votre entreprise'
  };
}

// Envoi
$('btnSend').addEventListener('click', async ()=>{
  if($('btnSend').disabled) return;
  // validations
  if(!$('prenom').value.trim() || !$('nom').value.trim() || !$('poste').value.trim()){
    alert('Profil incomplet : prénom / nom / poste requis');
    return;
  }
  if(!cvInput.files[0]){ alert('Veuillez ajouter votre CV'); return; }
  if(!emailsText.value.trim()){ alert('Ajoutez au moins une adresse email'); return; }
  if(excelEntries && !excelEntries.some(e=>!e.excluded)){ alert('Sélectionnez au moins une candidature dans le tableau.'); return; }
  if(!$('smtp_host').value.trim() || !$('smtp_port').value.trim() || !$('smtp_user').value.trim()){
    alert('Configuration SMTP incomplète (hôte/port/utilisateur requis)');
    return;
  }
  const btn=$('btnSend');
  btn.disabled=true;
  // smtp_pass peut être vide si smtp_secure = none (test local)
  if(!$('lettre_template').value.trim()){
    if(letterMode==='auto'){
      // en mode auto on régénère si vide
      const data=collectFormData();
      try{
        const r=await fetch('/api/generate-letter',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
        const j=await r.json();
        if(!r.ok) throw new Error(j.error || 'Erreur de génération');
        $('lettre_template').value=j.template;
      }catch(e){ btn.disabled=false; alert('Impossible de générer la lettre : '+e.message); return; }
    } else {
      alert('En mode « Ma propre lettre », écris ou importe ta lettre avant d\'envoyer.');
      setLetterMode('custom');
      $('lettre_template').focus();
      btn.disabled=false;
      return;
    }
  }

  const originalText=btn.textContent;
  btn.textContent='⏳ Envoi en cours... ne fermez pas la page';

  const resultsDiv=$('results');
  resultsDiv.innerHTML='';
  $('progress').style.width='10%';
  $('progressText').textContent='Connexion SMTP...';
  $('summary').classList.add('hidden');

  const fd=new FormData();
  const campaignId=crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  fd.append('campaign_id', campaignId);
  fd.append('follow_up_days', $('follow_up_days').value);
  fd.append('cv', cvInput.files[0]);
  fd.append('smtp_host', $('smtp_host').value.trim());
  fd.append('smtp_port', $('smtp_port').value.trim());
  fd.append('smtp_user', $('smtp_user').value.trim());
  fd.append('smtp_pass', $('smtp_pass').value);
  fd.append('smtp_secure', $('smtp_secure').value);
  fd.append('from_name', $('from_name').value.trim());
  fd.append('subject', $('subject').value.trim());
  fd.append('message_body', $('message_body').value);
  fd.append('emails_text', emailsText.value);
  fd.append('lettre_template', $('lettre_template').value);
  fd.append('prenom', $('prenom').value.trim());
  fd.append('nom', $('nom').value.trim());
  fd.append('ville', $('ville').value.trim());
  fd.append('telephone', $('telephone').value.trim());
  fd.append('emailCandidat', $('emailCandidat').value.trim());
  fd.append('poste', $('poste').value.trim());
  fd.append('experience', $('experience').value.trim());
  fd.append('motivation', $('motivation').value.trim());
  fd.append('ton', $('ton').value);
  fd.append('disponibilite', $('disponibilite').value.trim());
  fd.append('personnalise', $('personnalise').checked);
  fd.append('useCustomNames', $('useCustomNames').checked);
  fd.append('delay', $('delay').value);
  if(excelEntries && excelEntries.length){
    fd.append('excel_entries', JSON.stringify(excelEntries.filter(e=>!e.excluded)));
  }

  // Pré-affichage pending - une entrée par ligne
  const pendingEntries = excelEntries ? excelEntries.filter(e=>!e.excluded).map(e=>`${e.entreprise||''} <${e.email}>`) : emailsText.value.split(/\r?\n/).map(s=>s.trim()).filter(s=>s.includes('@'));
  pendingEntries.forEach(entry=>{
    const d=document.createElement('div');
    d.className='result pending';
    // Affiche nom + email si présent
    const display = entry.length>60 ? entry.slice(0,60)+'…' : entry;
    d.innerHTML=`<span>${escapeHtml(display)}</span><span>⏳</span>`;
    resultsDiv.appendChild(d);
  });

  let polling=true;
  const poll=async ()=>{
    try{
      const response=await fetch(`/api/campaigns/${campaignId}`);
      const job=await response.json();
      if(!polling) return;
      if(job.total){
        $('progress').style.width=`${Math.round(job.results.length/job.total*100)}%`;
        $('progressText').textContent=`${job.results.length} / ${job.total} candidatures traitées`;
        resultsDiv.innerHTML=job.results.map(r=>`<div class="result ${escapeHtml(r.status)}"><span>${escapeHtml(r.email)}</span><span>${r.status==='success'?'Envoyée':r.status==='skipped'?'Ignorée':'Échec'}</span></div>`).join('');
      }
    }catch(error){ /* La réponse finale reste la source des résultats. */ }
    if(polling) setTimeout(poll, 1000);
  };
  const firstPoll=setTimeout(poll, 500);
  try{
    const res=await fetch('/api/send',{method:'POST', body:fd});
    const data=await res.json();
    if(!res.ok) throw new Error(data.error||'Erreur envoi');
    polling=false;

    // update results (affiche le vrai message d'erreur pour diagnostic)
    const esc = s => String(s??'').replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    resultsDiv.innerHTML='';
    data.results.forEach(r=>{
      const d=document.createElement('div');
      d.className=`result ${r.status==='success'?'success':r.status==='skipped'?'skipped':'error'}`;
      if(r.status==='success'){
        d.innerHTML=`<span>✓ ${esc(r.email)} <small>(${esc(r.entreprise)})</small></span><span>Envoyé</span>`;
      } else if(r.status==='skipped'){
        d.innerHTML=`<span>${esc(r.email)}<small> — ${esc(r.reason)}</small></span><span>Ignorée</span>`;
      } else {
        const shortErr = r.error && r.error.length>120 ? r.error.slice(0,120)+'…' : (r.error||'Erreur inconnue');
        d.style.flexDirection='column';
        d.style.alignItems='stretch';
        d.innerHTML=`<div style="display:flex;justify-content:space-between;gap:8px"><span>✗ ${esc(r.email)}</span><span>Erreur</span></div><div class="err-detail" title="${esc(r.error||'')}">${esc(shortErr)}</div>`;
      }
      resultsDiv.appendChild(d);
    });

    $('progress').style.width='100%';
    $('progressText').textContent=`${data.success}/${data.total} envoyés`;
    const summary=$('summary');
    summary.classList.remove('hidden');
    if(data.failed===0){
      summary.className='summary ok';
      summary.textContent=`🎉 ${data.success} candidatures envoyées avec succès !`;
    }else{
      summary.className='summary fail';
      summary.textContent=`⚠️ ${data.success} envoyés, ${data.failed} échec(s).`;
    }
    if(data.invalid_emails && data.invalid_emails.length){
      summary.textContent+=` • ${data.invalid_emails.length} email(s) invalide(s) ignoré(s)`;
    }
    if(data.skipped) summary.textContent+=` • ${data.skipped} déjà envoyée(s), ignorée(s).`;
    document.dispatchEvent(new Event('campaigncomplete'));
  }catch(e){
    $('progress').style.width='0%';
    $('progressText').textContent='Erreur';
    const summary=$('summary');
    summary.classList.remove('hidden');
    summary.className='summary fail';
    summary.textContent='❌ '+e.message;
    console.error(e);
  }finally{
    polling=false;
    clearTimeout(firstPoll);
    btn.disabled=false;
    btn.textContent=originalText;
  }
});

// Auto remplissage objet/message intelligents quand poste change
$('poste').addEventListener('blur', ()=>{
  if(!$('subject').value && $('poste').value){
    $('subject').placeholder = `Candidature - {ENTREPRISE} - ${$('poste').value}`;
  }
});
