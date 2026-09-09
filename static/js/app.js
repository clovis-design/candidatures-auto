const $ = id => document.getElementById(id);
const dropzone = $('dropzone');
const cvInput = $('cv');
const cvInfo = $('cvInfo');

dropzone.addEventListener('click', ()=> cvInput.click());
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
    cvInfo.textContent = `✓ ${f.name} (${(f.size/1024).toFixed(1)} Ko)`;
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
  if(['xlsx','xls','xlsm'].includes(ext)){
    $('excelInfo').textContent='⏳ Lecture Excel (ligne 4 = en-têtes, ligne 5+ = données)...';
    const fd=new FormData(); fd.append('file', file);
    try{
      const res=await fetch('/api/parse-excel',{method:'POST', body:fd});
      const data=await res.json();
      if(!res.ok) throw new Error(data.error||'Erreur');
      // Remplit la liste d'emails à partir du tableau (le tableau se remplit depuis la liste d'emails)
      emailsText.value = data.emails_text || data.emails.join('\n');
      excelEntries = data.entries; // garde le tableau enrichi (stage etc.) pour l'envoi
      if(data.entries.some(en=>en.entreprise)) $('useCustomNames').checked=true;
      parseEmails();
      // Affiche preview tableau
      renderExcelPreview(data.entries, data.invalid_rows);
      $('excelInfo').textContent=`✓ Excel importé : ${data.count} entreprises depuis ligne 5 (headers ligne 4) • ${data.invalid_rows.length} lignes invalides ignorées — le tableau est rempli, l'envoi utilisera ces infos (Entreprise, Stage ciblé, etc.)`;
    }catch(err){
      $('excelInfo').textContent='❌ '+err.message;
    }
    return;
  }
  // Sinon CSV/TXT
  const reader=new FileReader();
  reader.onload=()=>{
    const text=reader.result;
    const lines=text.split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
    const re=/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/;
    const out=[];
    for(const line of lines){
      if(!line) continue;
      let parts = line.split(/[\t;]+/);
      if(parts.length===1) parts=line.split(',');
      let email=null, nom=null;
      for(const p of parts){
        if(re.test(p.trim())) email=p.trim().match(re)[0];
        else if(p.trim()) nom=p.trim();
      }
      if(!email){
        const m=line.match(re);
        if(m) email=m[0];
      }
      if(email){
        if(nom) out.push(`${nom} <${email}>`);
        else {
          if(line.includes('<') || line.includes('|')) out.push(line);
          else out.push(email);
        }
      }
    }
    if(out.length){
      emailsText.value = out.join('\n');
      if(out.some(l=>l.includes('<') || l.includes('|'))) $('useCustomNames').checked=true;
      parseEmails();
    } else {
      const re2=/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
      const found=text.match(re2)||[];
      if(found.length){ emailsText.value = found.join('\n'); parseEmails(); }
    }
  };
  reader.readAsText(file);
});

function renderExcelPreview(entries, invalid_rows){
  const wrap=$('excelPreview');
  const tbl=$('previewTable');
  const stats=$('previewStats');
  if(!entries.length){ wrap.classList.add('hidden'); return; }
  wrap.classList.remove('hidden');
  const headers=["Entreprise","Email","Stage","Type","Interlocuteur","Résultats"];
  let html=`<thead><tr style="background:#f1f5f9">`+headers.map(h=>`<th style="padding:6px 8px;text-align:left;border-bottom:1px solid #e2e8f0;white-space:nowrap">${h}</th>`).join('')+`</tr></thead><tbody>`;
  entries.slice(0,20).forEach(en=>{
    html+=`<tr style="border-bottom:1px solid #f1f5f9"><td style="padding:6px 8px">${en.entreprise||''}</td><td style="padding:6px 8px;font-family:monospace">${en.email}</td><td style="padding:6px 8px">${en.stage||''}</td><td style="padding:6px 8px">${en.type_candidature||''}</td><td style="padding:6px 8px">${en.interlocuteur||''}</td><td style="padding:6px 8px">${en.resultats||''}</td></tr>`;
  });
  if(entries.length>20) html+=`<tr><td colspan="6" style="padding:6px 8px;text-align:center;color:#64748b">... et ${entries.length-20} autres</td></tr>`;
  html+=`</tbody>`;
  tbl.innerHTML=html;
  let s=`${entries.length} ligne(s) lues dès ligne 5 (headers ligne 4)`;
  if(invalid_rows && invalid_rows.length) s+=` • ${invalid_rows.length} ligne(s) invalide(s) (pas d'email en colonne Coordonnées)`;
  stats.textContent=s;
}

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
    const res=await fetch('/api/generate-excel',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({emails_text:text, poste})});
    if(!res.ok){ const j=await res.json(); throw new Error(j.error||'Erreur'); }
    const blob=await res.blob();
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a'); a.href=url; a.download=`suivi_candidatures_${new Date().toISOString().slice(0,10)}.xlsx`; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
    $('excelInfo').textContent=`✓ Excel généré : le tableau a été rempli à partir de ta liste d'emails (${text.split(/\n/).filter(s=>s.includes('@')).length} lignes dès ligne 5). Headers en ligne 4.`;
    // Optionnel : preview
    const entriesRes=await fetch('/api/parse-emails',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({text})});
    const data=await entriesRes.json();
    // Simule entries enrichies pour preview et garde pour envoi (stage par ligne) - defaults "aucun"
    const fakeEntries=data.entries.map(en=>({email:en.email, entreprise:en.entreprise, stage:poste, type_candidature:'Candidature spontanée', interlocuteur:'aucun', resultats:'en attente', date_entretien:'aucun', date_envoi:new Date().toLocaleDateString('fr-FR')}));
    excelEntries=fakeEntries;
    renderExcelPreview(fakeEntries, []);
  }catch(err){ $('excelInfo').textContent='❌ '+err.message; }
  finally{ btn.disabled=false; btn.textContent='📊 Générer Excel de suivi à partir de ma liste d\'emails → remplit le tableau ligne 5+'; }
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

// Génération lettre
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
  // validations
  if(!$('prenom').value.trim() || !$('nom').value.trim() || !$('poste').value.trim()){
    alert('Profil incomplet : prénom / nom / poste requis');
    return;
  }
  if(!cvInput.files[0]){ alert('Veuillez ajouter votre CV'); return; }
  if(!emailsText.value.trim()){ alert('Ajoutez au moins une adresse email'); return; }
  if(!$('smtp_host').value.trim() || !$('smtp_port').value.trim() || !$('smtp_user').value.trim()){
    alert('Configuration SMTP incomplète (hôte/port/utilisateur requis)');
    return;
  }
  // smtp_pass peut être vide si smtp_secure = none (test local)
  if(!$('lettre_template').value.trim()){
    // auto générer
    const data=collectFormData();
    try{
      const r=await fetch('/api/generate-letter',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
      const j=await r.json();
      $('lettre_template').value=j.template;
    }catch(e){ alert('Impossible de générer la lettre'); return; }
  }

  const btn=$('btnSend');
  btn.disabled=true;
  const originalText=btn.textContent;
  btn.textContent='⏳ Envoi en cours... ne fermez pas la page';

  const resultsDiv=$('results');
  resultsDiv.innerHTML='';
  $('progress').style.width='10%';
  $('progressText').textContent='Connexion SMTP...';
  $('summary').classList.add('hidden');

  const fd=new FormData();
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
    fd.append('excel_entries', JSON.stringify(excelEntries));
  }

  // Pré-affichage pending - une entrée par ligne
  const pendingEntries = emailsText.value.split(/\r?\n/).map(s=>s.trim()).filter(s=>s.includes('@'));
  pendingEntries.forEach(entry=>{
    const d=document.createElement('div');
    d.className='result pending';
    // Affiche nom + email si présent
    const display = entry.length>60 ? entry.slice(0,60)+'…' : entry;
    d.innerHTML=`<span>${display}</span><span>⏳</span>`;
    resultsDiv.appendChild(d);
  });

  try{
    const res=await fetch('/api/send',{method:'POST', body:fd});
    const data=await res.json();
    if(!res.ok) throw new Error(data.error||'Erreur envoi');

    // update results
    resultsDiv.innerHTML='';
    data.results.forEach(r=>{
      const d=document.createElement('div');
      d.className=`result ${r.status==='success'?'success':'error'}`;
      d.innerHTML = r.status==='success'
        ? `<span>✓ ${r.email} <small>(${r.entreprise})</small></span><span>Envoyé</span>`
        : `<span>✗ ${r.email}</span><span title="${r.error}">Erreur</span>`;
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
  }catch(e){
    $('progress').style.width='0%';
    $('progressText').textContent='Erreur';
    const summary=$('summary');
    summary.classList.remove('hidden');
    summary.className='summary fail';
    summary.textContent='❌ '+e.message;
    console.error(e);
  }finally{
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
