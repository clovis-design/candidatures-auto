// Tableau de bord et brouillon : le serveur est la source de vérité du suivi.
(() => {
  const draftKey='candidatures-auto.draft.v2';
  const draftFields=['prenom','nom','emailCandidat','telephone','ville','poste','disponibilite','ton',
    'experience','motivation','personnalise','useCustomNames','lettre_template','emails_text',
    'smtp_host','smtp_port','smtp_user','smtp_secure','from_name','subject','message_body','delay','follow_up_days'];
  let applications=[];
  let statuses={};
  let editingId=null;
  let draftTimer;

  function notify(message, error=false){
    $('workspaceMessage').textContent=message;
    $('workspaceMessage').className=`workspace-message${error?' error':''}`;
  }
  async function api(url, options={}){
    const response=await fetch(url, options);
    const data=await response.json();
    if(!response.ok) throw new Error(data.error || `Erreur HTTP ${response.status}`);
    return data;
  }
  function showPanel(tracking){
    $('composePanel').classList.toggle('hidden',tracking);
    $('trackingPanel').classList.toggle('hidden',!tracking);
    $('navCompose').classList.toggle('active',!tracking);
    $('navTracking').classList.toggle('active',tracking);
    $('navCompose').setAttribute('aria-pressed',!tracking);
    $('navTracking').setAttribute('aria-pressed',tracking);
    window.scrollTo({top:0,behavior:'instant'});
    if(tracking) loadHistory();
  }
  $('navCompose').addEventListener('click',()=>showPanel(false));
  $('navTracking').addEventListener('click',()=>showPanel(true));
  $('btnFirstCampaign').addEventListener('click',()=>showPanel(false));
  document.querySelectorAll('.stat-card').forEach(button=>button.addEventListener('click',()=>{
    $('historyFilter').value=button.dataset.filter;
    showPanel(true);
  }));

  function saveDraft(){
    const fields={};
    draftFields.forEach(id=>fields[id]=$(id).type==='checkbox'?$(id).checked:$(id).value);
    try{
      localStorage.setItem(draftKey,JSON.stringify({fields,excelEntries,letterMode}));
      $('draftStatus').textContent=`Brouillon sauvegardé à ${new Date().toLocaleTimeString('fr-FR',{hour:'2-digit',minute:'2-digit'})}`;
    }catch(error){ $('draftStatus').textContent='Sauvegarde locale indisponible dans ce navigateur'; }
  }
  function readiness(){
    const checks=[
      [$('prenom').value.trim() && $('nom').value.trim() && $('poste').value.trim(),'Profil renseigné'],
      [cvInput.files.length,'CV joint'],
      [$('lettre_template').value.trim() || (letterMode==='auto' && $('prenom').value.trim() && $('nom').value.trim() && $('poste').value.trim()),$('lettre_template').value.trim()?'Lettre préparée':'Lettre automatique à l’envoi'],
      [excelEntries ? excelEntries.some(e=>!e.excluded) : emailsText.value.includes('@'),'Destinataires ajoutés'],
      [$('smtp_host').value.trim() && $('smtp_port').value.trim() && $('smtp_user').value.trim(),'Configuration SMTP renseignée'],
    ];
    $('readiness').innerHTML=checks.map(([ready,label])=>`<li class="${ready?'ready':''}">${label}</li>`).join('');
    const steps=[checks[0][0] && checks[1][0],checks[2][0],checks[3][0],checks[4][0]];
    document.querySelectorAll('.step').forEach((step,i)=>step.classList.toggle('active',!!steps[i]));
  }
  function changed(){
    readiness(); clearTimeout(draftTimer); draftTimer=setTimeout(saveDraft,500);
  }
  try{
    const draft=JSON.parse(localStorage.getItem(draftKey)||'null');
    if(draft?.fields){
      draftFields.forEach(id=>{
        if(draft.fields[id]===undefined) return;
        if($(id).type==='checkbox') $(id).checked=!!draft.fields[id];
        else $(id).value=draft.fields[id];
      });
      excelEntries=Array.isArray(draft.excelEntries)?draft.excelEntries:null;
      setLetterMode(draft.letterMode==='custom'?'custom':'auto');
      if($('lettre_template').value.trim()) $('letterPreview').classList.remove('hidden');
      if(excelEntries) renderExcelPreview(excelEntries,[]);
      updateCustomNamesUI(); parseEmails();
      $('draftStatus').textContent='Brouillon restauré · Ajoutez à nouveau votre CV et le mot de passe SMTP';
    }
  }catch(error){ $('draftStatus').textContent='Nouveau brouillon'; }
  $('composePanel').addEventListener('input',changed);
  $('composePanel').addEventListener('change',changed);
  $('composePanel').addEventListener('click',()=>setTimeout(changed,0));
  document.addEventListener('workspacechange',changed);
  window.addEventListener('pagehide',saveDraft);
  readiness();

  // Associe aussi les libellés historiques aux champs pour clavier et lecteurs d'écran.
  document.querySelectorAll('.field').forEach(field=>{
    const label=field.querySelector('label');
    const control=field.querySelector('input[id],select[id],textarea[id]');
    if(label && control && !label.htmlFor) label.htmlFor=control.id;
  });
  $('smtpPreset').addEventListener('change',()=>{
    const preset={gmail:['smtp.gmail.com','587','tls'],outlook:['smtp.office365.com','587','tls'],
      ovh:['ssl0.ovh.net','465','ssl'],local:['localhost','1025','none']}[$('smtpPreset').value];
    if(!preset) return;
    ['smtp_host','smtp_port','smtp_secure'].forEach((id,i)=>$(id).value=preset[i]);
    changed();
  });
  $('btnTestSmtp').addEventListener('click',async()=>{
    const button=$('btnTestSmtp'); button.disabled=true;
    $('smtpTestResult').textContent='Connexion en cours…';
    try{
      const payload={};
      ['smtp_host','smtp_port','smtp_user','smtp_pass','smtp_secure'].forEach(id=>payload[id]=$(id).value);
      const data=await api('/api/test-smtp',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
      $('smtpTestResult').textContent=data.message;
    }catch(error){ $('smtpTestResult').textContent=error.message; }
    finally{button.disabled=false;}
  });

  function formatDate(value){
    if(!value) return '—';
    const date=new Date(value.length===10?`${value}T12:00:00`:value);
    return Number.isNaN(date.getTime())?value:date.toLocaleDateString('fr-FR');
  }
  async function loadHistory(){
    try{
      const data=await api('/api/applications');
      applications=data.applications; statuses=data.statuses;
      $('statSent').textContent=data.stats.sent;
      $('statInterviews').textContent=data.stats.interviews;
      $('statDue').textContent=data.stats.due;
      $('statFailed').textContent=data.stats.failed;
      $('historyCount').textContent=data.stats.total;
      renderHistory();
    }catch(error){notify(`Chargement du suivi impossible : ${error.message}`,true);}
  }
  function renderHistory(){
    const search=$('historySearch').value.toLocaleLowerCase('fr').trim();
    const filter=$('historyFilter').value;
    const rows=applications.filter(r=>(!filter || (filter==='due'?r.due:r.status===filter)) &&
      `${r.entreprise} ${r.email} ${r.stage} ${r.interlocuteur||''} ${r.notes}`.toLocaleLowerCase('fr').includes(search));
    $('historyBody').innerHTML=rows.map(r=>`<tr>
      <td><strong>${escapeHtml(r.entreprise||r.email)}</strong><small>${escapeHtml(r.email)}</small></td>
      <td>${escapeHtml(r.stage)}${r.date_entretien?`<small>Entretien : ${escapeHtml(formatDate(r.date_entretien))}</small>`:''}</td>
      <td>${escapeHtml(formatDate(r.sent_at))}</td>
      <td><span class="status-pill ${escapeHtml(r.status)}">${escapeHtml(statuses[r.status])}</span>${r.error?`<small>${escapeHtml(r.error)}</small>`:''}</td>
      <td><span class="${r.due?'due-date':''}">${escapeHtml(formatDate(r.follow_up_date))}</span></td>
      <td>${r.sent_at?`<button class="btn small ghost" data-edit="${r.id}">Mettre à jour</button>`:r.status==='error'?`<button class="btn small ghost" data-retry="${r.id}">Reprendre</button>`:'—'}</td>
    </tr>`).join('');
    $('historyEmpty').classList.toggle('hidden',applications.length>0);
    $('historyStats').textContent=applications.length?`${rows.length} candidature(s) affichée(s) sur ${applications.length}${!rows.length?' · Aucun résultat pour ces filtres.':''}`:'';
  }
  $('historySearch').addEventListener('input',renderHistory);
  $('historyFilter').addEventListener('change',renderHistory);
  $('btnRefreshHistory').addEventListener('click',loadHistory);
  document.addEventListener('campaigncomplete',()=>{loadHistory();saveDraft();});
  $('historyBody').addEventListener('click',event=>{
    const edit=event.target.closest('[data-edit]');
    const retry=event.target.closest('[data-retry]');
    const record=applications.find(r=>r.id===Number(edit?.dataset.edit || retry?.dataset.retry));
    if(!record) return;
    if(retry){
      // Une reprise prépare le destinataire, elle n'envoie rien immédiatement.
      if(emailsText.value.trim() && !confirm('Remplacer la liste de destinataires du brouillon par cette candidature à reprendre ?')) return;
      excelEntries=[{email:record.email,entreprise:record.entreprise,stage:record.stage,
        type_candidature:record.type_candidature,interlocuteur:record.interlocuteur,coordonnees:record.coordonnees}];
      emailsText.value=`${record.entreprise||''} <${record.email}>`;
      renderExcelPreview(excelEntries,[]); parseEmails(); showPanel(false); changed();
      notify('Candidature chargée pour reprise. Vérifiez votre CV et la configuration SMTP, puis lancez l’envoi.');
      return;
    }
    editingId=record.id;
    $('editTitle').textContent=record.entreprise||'Suivi de candidature';
    $('editContact').textContent=`${record.stage} · ${record.email}`;
    $('editStatus').value=record.status;
    $('editFollowUp').value=record.follow_up_date;
    $('editInterview').value=/^\d{4}-\d{2}-\d{2}$/.test(record.date_entretien||'')?record.date_entretien:'';
    $('editInterlocuteur').value=record.interlocuteur||'';
    $('editNotes').value=record.notes||'';
    $('editError').textContent='';
    $('editApplication').showModal();
  });
  $('closeEdit').addEventListener('click',()=>$('editApplication').close());
  $('trackingForm').addEventListener('submit',async event=>{
    event.preventDefault(); $('saveTracking').disabled=true;
    try{
      await api(`/api/applications/${editingId}`,{method:'PATCH',headers:{'Content-Type':'application/json'},body:JSON.stringify({
        status:$('editStatus').value,follow_up_date:$('editFollowUp').value,date_entretien:$('editInterview').value,
        interlocuteur:$('editInterlocuteur').value,notes:$('editNotes').value,
      })});
      $('editApplication').close(); await loadHistory();
      notify('Suivi enregistré. Les prochaines exportations Excel incluront vos modifications.');
    }catch(error){$('editError').textContent=error.message;}
    finally{$('saveTracking').disabled=false;}
  });
  $('updateWorkbook').addEventListener('change',async event=>{
    const file=event.target.files[0]; if(!file) return;
    const fd=new FormData(); fd.append('file',file);
    event.target.disabled=true;
    notify('Mise à jour de votre tableur…');
    try{
      const response=await fetch('/api/update-excel',{method:'POST',body:fd});
      if(!response.ok){const data=await response.json();throw new Error(data.error);}
      const url=URL.createObjectURL(await response.blob());
      const link=document.createElement('a');link.href=url;link.download=`suivi_${file.name}`;
      document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
      const count=Number(response.headers.get('X-Updated-Rows'));
      notify(count?`${count} ligne(s) complétée(s). Votre tableur mis à jour a été téléchargé.`:'Aucune correspondance unique trouvée. Vérifiez les colonnes Email et Poste de votre fichier.');
    }catch(error){notify(error.message,true);}
    finally{event.target.disabled=false;event.target.value='';}
  });
  loadHistory();
})();
