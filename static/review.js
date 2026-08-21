(() => {
  const rows = JSON.parse(document.getElementById('appRows').textContent || '[]');
  const refs = JSON.parse(document.getElementById('appRefs').textContent || '{}');
  const headers = JSON.parse(document.getElementById('appHeaders').textContent || '[]');
  const sources = JSON.parse(document.getElementById('appSources').textContent || '[]');
  const cards = document.getElementById('cards');
  const saveForm = document.getElementById('saveForm');
  const rowsPayload = document.getElementById('rowsPayload');
  const searchInput = document.getElementById('searchInput');
  const onlyNeeds = document.getElementById('onlyNeeds');
  const validationBanner = document.getElementById('validationBanner');
  const allFieldsDialog = document.getElementById('allFieldsDialog');
  const allFieldsGrid = document.getElementById('allFieldsGrid');
  const sourceDialog = document.getElementById('sourceDialog');
  const sourceGrid = document.getElementById('sourceGrid');
  const activeTabs = {};
  let editingIndex = null;
  let allFieldDraft = null;

  const esc = (s) => String(s ?? '').replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
  const norm = (s) => String(s ?? '').toLowerCase().replace(/[^a-z0-9]+/g, '');
  const clean = (s) => String(s ?? '').trim();
  const digits = (s) => String(s ?? '').replace(/\D/g, '');
  const storeList = Array.isArray(refs.stores) ? refs.stores : [];
  const dealers = Array.isArray(refs.dealers) ? refs.dealers : [];
  const zones = Array.isArray(refs.zones) ? refs.zones : [];
  const taxzones = Array.isArray(refs.taxzones) ? refs.taxzones : [];
  const agents = Array.isArray(refs.agents) ? refs.agents : [];
  const existingContracts = new Set((Array.isArray(refs.existing_contracts) ? refs.existing_contracts : []).map(x => clean(x).toUpperCase()).filter(Boolean));
  const usedSuffixes = ['U', ...'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('').filter(x => x !== 'U')];

  function key(obj, ...names) {
    for (const n of names) if (obj && obj[n] !== undefined) return obj[n];
    return '';
  }
  function storeId(s) { return clean(key(s, 'store', 'STORE')); }
  function storeTitle(s) { return clean(key(s, 'storetitle', 'STORETITLE')) || `Store ${storeId(s)}`; }
  function dealerId(d) { return clean(key(d, 'id', 'ID')); }
  function dealerName(d) { return clean(key(d, 'companydealername', 'COMPANYDEALERNAME')); }
  function dealerStore(d) { return clean(key(d, 'store', 'STORE')); }
  function zoneId(z) { return clean(key(z, 'zone', 'ZONE')); }
  function zoneSelector(z) { return clean(key(z, 'selector', 'SELECTOR')) || `${clean(key(z,'description','DESCRIPTION'))}, ${zoneId(z)}`; }
  function zoneStore(z) { return clean(key(z, 'store', 'STORE')); }
  function taxZoneId(t) { return clean(key(t, 'zone', 'ZONE')); }
  function taxState(t) { return clean(key(t, 'state', 'STATE')).toUpperCase(); }
  function taxCounty(t) { return clean(key(t, 'county', 'COUNTY')).toUpperCase(); }
  function taxCity(t) { return clean(key(t, 'city', 'CITY')).toUpperCase(); }
  function taxDesc(t) { return clean(key(t, 'description', 'DESCRIPTION')) || taxZoneId(t); }
  function taxRate(t) { return clean(key(t, 'taxrate', 'TAXRATE')); }
  function agentStore(a) { return clean(key(a, 'store', 'STORE')); }
  function agentName(a) { return clean(key(a, 'agent', 'AGENT')); }

  function findStoreTitle(id) {
    const s = storeList.find(x => storeId(x) === clean(id));
    return s ? storeTitle(s) : '';
  }
  function findDealerName(id) {
    const d = dealers.find(x => dealerId(x) === clean(id));
    return d ? dealerName(d) : '';
  }
  function findAgent(id) {
    const a = agents.find(x => agentStore(x) === clean(id));
    return a ? agentName(a) : '';
  }

  function storeOptions(current='') {
    const seen = new Set();
    let html = '<option value="">— Select store —</option>';
    for (const s of storeList) {
      const id = storeId(s); if (!id || seen.has(id)) continue; seen.add(id);
      html += `<option value="${esc(id)}" ${id===clean(current)?'selected':''}>${esc(storeTitle(s))} · ID ${esc(id)}</option>`;
    }
    if (current && !seen.has(clean(current))) html += `<option selected value="${esc(current)}">Custom store · ${esc(current)}</option>`;
    return html;
  }

  function dealerOptions(row) {
    const current = clean(row.DEALERID);
    const list = dealers.filter(d => !row.STORE || dealerStore(d) === clean(row.STORE));
    let html = '<option value="">— Select dealer —</option>';
    const seen = new Set();
    for (const d of list.sort((a,b)=>dealerName(a).localeCompare(dealerName(b)))) {
      const id = dealerId(d); if (!id || seen.has(id)) continue; seen.add(id);
      html += `<option value="${esc(id)}" ${id===current?'selected':''}>${esc(dealerName(d))} · ID ${esc(id)}</option>`;
    }
    if (current && !seen.has(current)) html += `<option selected value="${esc(current)}">Current/Manual · ID ${esc(current)}</option>`;
    return html;
  }

  function zoneOptions(row) {
    const current = clean(row.ZONE);
    const list = zones.filter(z => !row.STORE || zoneStore(z) === clean(row.STORE));
    let html = '<option value="">— Select zone —</option>';
    const seen = new Set();
    for (const z of list.sort((a,b)=>zoneSelector(a).localeCompare(zoneSelector(b)))) {
      const id = zoneId(z); if (!id || seen.has(id)) continue; seen.add(id);
      html += `<option value="${esc(id)}" ${id===current?'selected':''}>${esc(zoneSelector(z))}</option>`;
    }
    if (current && !seen.has(current)) html += `<option selected value="${esc(current)}">Current/Manual · ${esc(current)}</option>`;
    return html;
  }

  function taxOptions(row) {
    const current = clean(row.TAXZONE);
    const st = clean(row.del_state || row._source_state || row.STATE).toUpperCase();
    const county = clean(row._ziptax_county || row.COUNTY).toUpperCase();
    const list = taxzones.filter(t => !st || taxState(t) === st);
    let html = '<option value="">— Select tax zone —</option>';
    const seen = new Set();
    const sorted = list.sort((a,b) => {
      const ac = county && taxCounty(a) === county ? 0 : 1;
      const bc = county && taxCounty(b) === county ? 0 : 1;
      return ac - bc || taxDesc(a).localeCompare(taxDesc(b));
    });
    for (const t of sorted) {
      const id = taxZoneId(t); if (!id || seen.has(id)) continue; seen.add(id);
      const rate = taxRate(t);
      const where = [taxCounty(t), taxCity(t)].filter(Boolean).join(' · ');
      html += `<option data-rate="${esc(rate)}" value="${esc(id)}" ${id===current?'selected':''}>${esc(taxDesc(t))} · ${esc(id)}${rate?` · ${esc(rate)}`:''}${where?` · ${esc(where)}`:''}</option>`;
    }
    if (current && !seen.has(current)) html += `<option selected value="${esc(current)}">Current/Manual · ${esc(current)}</option>`;
    return html;
  }

  function boolish(v) {
    return v === true || ['1','true','yes','on'].includes(clean(v).toLowerCase());
  }

  function usedSuffixState(i, suffix) {
    const r=rows[i];
    const model=clean(r.MODEL1);
    suffix=clean(suffix).toUpperCase();
    if (!model) return {contract:'',suffix,error:'Used building needs a model number.',taken:false};
    if (!usedSuffixes.includes(suffix)) return {contract:'',suffix,error:'Choose a valid used-building suffix.',taken:false};
    const candidate=`${model}${suffix}`;
    if (candidate.length > 10) return {contract:candidate,suffix,error:'That suffix would exceed RTO Pro’s 10-character contract limit.',taken:false};
    const key=candidate.toUpperCase();
    let taken=existingContracts.has(key);
    if (!taken) {
      taken=rows.some((other,j)=>j!==i && clean(other.CONTRACT).toUpperCase()===key);
    }
    return {contract:candidate,suffix,error:taken?'That contract number is already in use. Choose another suffix.':'',taken};
  }

  function usedContractCandidate(i) {
    for (const suffix of usedSuffixes) {
      const state=usedSuffixState(i,suffix);
      if (state.contract && !state.error) return state;
    }
    return {contract:'',suffix:'',error:'No available suffix fits RTO Pro’s 10-character contract limit.'};
  }

  function usedSuffixOptions(i, current='') {
    current=clean(current).toUpperCase();
    return usedSuffixes.map(suffix=>{
      const state=usedSuffixState(i,suffix);
      const note=state.error ? (state.taken?' · already used':' · unavailable') : '';
      return `<option value="${esc(suffix)}" ${suffix===current?'selected':''} ${state.error?'disabled':''}>${esc(suffix+note)}</option>`;
    }).join('');
  }

  function setUsedSuffix(i, suffix) {
    const r=rows[i];
    const state=usedSuffixState(i,suffix);
    r._used_suffix_manual=true;
    r._used_contract_error=state.error;
    if (!state.error && state.contract) {
      r.CONTRACT=state.contract;
      r._used_contract_suffix=state.suffix;
      r._used_base_contract=clean(r.MODEL1);
    }
  }

  function applyUsedBuilding(i, enabled, manual=true) {
    const r=rows[i];
    const preferredSuffix=clean(r._used_contract_suffix).toUpperCase();
    const preserveManual=boolish(r._used_suffix_manual) && usedSuffixes.includes(preferredSuffix);
    r._used_building=!!enabled;
    if (manual) r._used_manual_override=true;
    r._used_base_contract=clean(r.MODEL1);
    r._used_contract_error='';
    if (enabled) {
      if (preserveManual) {
        const preferred=usedSuffixState(i,preferredSuffix);
        if (!preferred.error && preferred.contract) {
          r.CONTRACT=preferred.contract;
          r._used_contract_suffix=preferred.suffix;
          return;
        }
      }
      const choice=usedContractCandidate(i);
      if (choice.contract) {
        r.CONTRACT=choice.contract;
        r._used_contract_suffix=choice.suffix;
        r._used_suffix_manual=false;
      } else {
        r.CONTRACT='';
        r._used_contract_suffix='';
        r._used_suffix_manual=false;
        r._used_contract_error=choice.error;
      }
    } else {
      // Normal V3/V6 behavior: contract follows the building model.
      r.CONTRACT=clean(r.MODEL1);
      r._used_contract_suffix='';
      r._used_suffix_manual=false;
    }
  }

  function rowProblems(r) {
    const p=[];
    if (!clean(r.STORE)) p.push('Store');
    if (!clean(r.DEALERID)) p.push('Dealer ID');
    if (!clean(r.MODEL1)) p.push('Model #');
    if (!clean(r.CONTRACT)) p.push('Contract #');
    if (!clean(r.ZONE)) p.push('Zone');
    if (!clean(r.TAXZONE)) p.push('Tax zone');
    if (boolish(r._used_building)) {
      const model=clean(r.MODEL1).toUpperCase(), contract=clean(r.CONTRACT).toUpperCase();
      const suffix=(model && contract.startsWith(model)) ? contract.slice(model.length) : '';
      if (clean(r._used_contract_error) || !contract || contract===model || !contract.startsWith(model) || !usedSuffixes.includes(suffix) || contract.length>10) p.push('Used contract suffix');
    }
    return [...new Set(p)];
  }

  function sourceMatchText(r) {
    const selected = findDealerName(r.DEALERID) || clean(r._dealer_name);
    if (!clean(r._source_dealer)) return '';
    if (selected && norm(selected) === norm(r._source_dealer)) return '<span class="match good">✓ source dealer matched</span>';
    if (selected) return `<span class="match warn">Selected: ${esc(selected)}</span>`;
    return '<span class="match warn">Dealer needs mapping</span>';
  }

  function inputField(i, field, label, opts={}) {
    const r=rows[i]; const val=clean(r[field]);
    const type=opts.type || 'text';
    return `<label class="field ${opts.wide?'wide-field':''} ${opts.emphasis?'emphasis':''}"><span>${esc(label)}</span><input type="${type}" data-i="${i}" data-field="${field}" value="${esc(val)}" ${opts.placeholder?`placeholder="${esc(opts.placeholder)}"`:''}>${opts.small?`<small>${esc(opts.small)}</small>`:''}</label>`;
  }

  function helperField(i, helper, label, opts={}) {
    const val=clean(rows[i][helper]);
    return `<label class="field ${opts.wide?'wide-field':''}"><span>${esc(label)}</span><input data-i="${i}" data-helper="${esc(helper)}" value="${esc(val)}" ${opts.placeholder?`placeholder="${esc(opts.placeholder)}"`:''}>${opts.small?`<small>${esc(opts.small)}</small>`:''}</label>`;
  }

  function textareaField(i, field, label, opts={}) {
    return `<label class="field textarea-field ${opts.wide?'wide-field':''}"><span>${esc(label)}</span><textarea data-i="${i}" data-field="${field}" rows="${opts.rows || 3}">${esc(rows[i][field])}</textarea>${opts.small?`<small>${esc(opts.small)}</small>`:''}</label>`;
  }

  function mappingTab(r,i) {
    const inventorySuggestion=clean(r._source_inventory_id);
    const nextSuggestion=clean(r._next_model);
    const stockSuggestion=clean(r._stock_model);
    const contractSuggestion=clean(r._set_contract);
    const used=boolish(r._used_building);
    const detected=boolish(r._used_detected);
    const reason=clean(r._used_detection_reason);
    const usedError=clean(r._used_contract_error);
    const usedStatus = used
      ? `<span class="used-pill active">USED BUILDING${r._used_contract_suffix?` · ${esc(r._used_contract_suffix)}`:''}</span>`
      : (detected ? '<span class="used-pill detected">USED DETECTED</span>' : '');
    return `<div class="tab-section">
      <div class="used-building-box ${used?'active':''} ${detected?'detected':''}">
        <label class="used-checkbox"><input type="checkbox" class="used-building-toggle" data-i="${i}" ${used?'checked':''}><span><strong>Used building</strong><small>Keep the same Model # and give the RTO contract a unique suffix.</small></span></label>
        <div class="used-building-status">${usedStatus}${reason?`<small>${esc(reason)}</small>`:''}${usedError?`<small class="used-error">${esc(usedError)}</small>`:''}</div>
      </div>
      <div class="mapping-grid top-grid">
        <label class="field emphasis"><span>Store</span><select class="store-select" data-i="${i}">${storeOptions(r.STORE)}</select><small>${esc(r._store_title || findStoreTitle(r.STORE) || 'Choose per contract')}</small></label>
        <label class="field emphasis dealer-field"><span>Dealer</span><select class="dealer-select" data-i="${i}">${dealerOptions(r)}</select><small>Filtered by selected Store</small></label>
        ${inputField(i,'DEALERID','Dealer ID',{placeholder:'auto from Dealer'})}
        <div class="field suggestion-field"><span>Account</span><input data-i="${i}" data-field="ACCOUNT" value="${esc(r.ACCOUNT)}" placeholder="existing/new account"><div class="suggestions">${Array.isArray(r._account_suggestions)?r._account_suggestions.map(a=>`<button type="button" class="chip suggestion" data-i="${i}" data-field="ACCOUNT" data-value="${esc(a)}">Existing ${esc(a)}</button>`).join(''):''}</div></div>
      </div>
      <div class="mapping-grid model-grid">
        <div class="field suggestion-field"><span>Model #</span><input data-i="${i}" data-field="MODEL1" value="${esc(r.MODEL1)}"><div class="suggestions">${inventorySuggestion?`<button type="button" class="chip suggestion" data-i="${i}" data-field="MODEL1" data-value="${esc(inventorySuggestion)}">Inventory ${esc(inventorySuggestion)}</button>`:''}${stockSuggestion?`<button type="button" class="chip suggestion stock" data-i="${i}" data-field="MODEL1" data-value="${esc(stockSuggestion)}">Stock ${esc(stockSuggestion)}</button>`:''}${!used && nextSuggestion?`<button type="button" class="chip suggestion" data-i="${i}" data-field="MODEL1" data-value="${esc(nextSuggestion)}">Next ${esc(nextSuggestion)}</button>`:''}</div>${used?'<small class="used-model-note">Used: keep the original building model number.</small>':''}</div>
        <div class="field suggestion-field"><span>Contract #</span><input data-i="${i}" data-field="CONTRACT" value="${esc(r.CONTRACT)}" ${used?'readonly aria-readonly="true"':''}><div class="suggestions">${!used?`<button type="button" class="chip copy-model" data-i="${i}">Copy Model #</button>`:''}${!used && contractSuggestion?`<button type="button" class="chip suggestion" data-i="${i}" data-field="CONTRACT" data-value="${esc(contractSuggestion)}">Order ${esc(contractSuggestion)}</button>`:''}${used?`<label class="used-suffix-picker"><span>Suffix</span><select class="used-suffix-select" data-i="${i}">${usedSuffixOptions(i,r._used_contract_suffix)}</select>${boolish(r._used_suffix_manual)?'<small>Manual</small>':'<small>Auto</small>'}</label>`:''}</div></div>
        <label class="field"><span>Zone</span><select class="zone-select" data-i="${i}">${zoneOptions(r)}</select><small>${esc(r._zone_selector || '')}</small></label>
        <div class="field readout-field"><span>Tax Zone</span><strong>${esc(r.TAXZONE || 'Needs selection')}</strong><small>Change it in Tax & Address</small></div>
      </div>
    </div>`;
  }

  function inventoryTab(r,i) {
    return `<div class="tab-section">
      <div class="inventory-feature-row">
        ${inputField(i,'CATEGORY1','Category',{emphasis:true})}
        ${textareaField(i,'DESCRIPTION1','Description',{rows:2})}
      </div>
      <div class="source-hint">ShedSuite model variation: <strong>${esc(r._source_model_variation || '—')}</strong></div>
      <div class="mapping-grid inventory-grid">
        ${inputField(i,'MODEL1','Model #')}
        ${inputField(i,'SERIAL1','Serial #')}
        ${inputField(i,'CONDITION1','Condition')}
        ${inputField(i,'BRAND1','Brand')}
        ${inputField(i,'VENDOR1','Vendor')}
        ${inputField(i,'COST1','Cost')}
        ${inputField(i,'RETAIL1','Retail')}
        ${inputField(i,'RATE1','Rental rate')}
        ${inputField(i,'STOCK1','Stock')}
        ${inputField(i,'INVOICE1','Invoice #')}
      </div>
    </div>`;
  }

  function phoneSourceOptions(r) {
    const cur=clean(r._phone_other_source || 'ref1').toLowerCase();
    return [
      ['ref1','Ref 1'],['ref2','Ref 2'],['secondary','Secondary']
    ].map(([v,l])=>`<option value="${v}" ${cur===v?'selected':''}>${l}</option>`).join('');
  }

  function customerTab(r,i) {
    return `<div class="tab-section">
      <div class="mapping-grid customer-basic-grid">
        ${inputField(i,'NAME','Customer name')}
        ${inputField(i,'EMAIL','Email')}
        ${inputField(i,'DOB1','DOB')}
        ${inputField(i,'SS1','SSN')}
        ${inputField(i,'DL1','Driver license / ID')}
        ${inputField(i,'WORK1','Employer')}
        ${inputField(i,'WORKPH1','Employer phone')}
      </div>

      <div class="address-compare">
        <section class="address-box"><header><strong>Mailing address</strong><span>Customer / RTO mailing</span></header>
          <div class="mapping-grid address-grid">
            ${inputField(i,'ADDRESS','Street')}${inputField(i,'APT','Apt / Unit')}${inputField(i,'CITY','City')}${inputField(i,'STATE','State')}${inputField(i,'ZIP','ZIP')}
          </div>
        </section>
        <section class="address-box delivery"><header><strong>Delivery address</strong><span>Physical destination / tax address</span></header>
          <div class="mapping-grid address-grid">
            ${inputField(i,'del_address1','Street')}${inputField(i,'del_address2','Address 2')}${inputField(i,'del_city','City')}${inputField(i,'del_state','State')}${inputField(i,'del_zip','ZIP')}
          </div>
        </section>
      </div>

      <section class="phone-rule-box">
        <div class="section-heading-row"><div><strong>Phone rule</strong><span>Primary always remains Primary. “Other” is selected automatically from Secondary / Ref 1 / Ref 2.</span></div><button type="button" class="mini-btn auto-phone" data-i="${i}">Reapply auto rule</button></div>
        <div class="mapping-grid phone-grid">
          ${inputField(i,'CELL','Primary phone',{emphasis:true})}
          ${helperField(i,'_secondary_phone','Secondary phone')}
          <label class="field emphasis"><span>Selected “Other” source</span><select class="phone-source-select" data-i="${i}">${phoneSourceOptions(r)}</select><small>${r._phone_other_manual?'Manual override':'Automatic rule'}</small></label>
          ${inputField(i,'PHONE5','RTO Other phone',{emphasis:true})}
        </div>
        <div class="reference-grid">
          <div class="reference-card"><strong>Reference 1</strong><div class="mapping-grid ref-fields">${inputField(i,'REFERENCE1','Name')}${inputField(i,'REFRELATION1','Relationship')}${inputField(i,'REFPHONE1','Phone')}</div></div>
          <div class="reference-card"><strong>Reference 2</strong><div class="mapping-grid ref-fields">${inputField(i,'REFERENCE2','Name')}${inputField(i,'REFRELATION2','Relationship')}${inputField(i,'REFPHONE2','Phone')}</div></div>
        </div>
        ${textareaField(i,'COMMENTS','Comment box',{rows:2,wide:true,small:'Automatically says OTHER IS + relationship + name when Ref 1 or Ref 2 is selected.'})}
      </section>
    </div>`;
  }

  function incorporatedText(v) {
    if (v === true || v === 'true') return 'Yes';
    if (v === false || v === 'false') return 'No';
    return 'Unknown';
  }

  function taxTab(r,i) {
    const note=clean(r._ziptax_note);
    return `<div class="tab-section">
      <div class="tax-layout">
        <section class="address-box delivery"><header><strong>Delivery / tax address</strong><span>ZipTax uses this address</span></header>
          <div class="mapping-grid address-grid">
            ${inputField(i,'del_address1','Street')}${inputField(i,'del_address2','Address 2')}${inputField(i,'del_city','City')}${inputField(i,'del_state','State')}${inputField(i,'del_zip','ZIP')}
          </div>
          <button type="button" class="secondary redetect-tax" data-i="${i}">Detect tax zone with ZipTax</button>
        </section>
        <section class="ziptax-box">
          <header><strong>ZipTax result</strong><span class="zt-status ${r.TAXZONE?'good':'warn'}">${r.TAXZONE?'Detected / mapped':'Needs selection'}</span></header>
          <div class="ziptax-facts">
            <div><span>County</span><strong>${esc(r._ziptax_county || r.COUNTY || '—')}</strong></div>
            <div><span>City</span><strong>${esc(r._ziptax_city || '—')}</strong></div>
            <div><span>In city limits</span><strong>${esc(incorporatedText(r._ziptax_incorporated))}</strong></div>
            <div><span>ZipTax rate</span><strong>${esc(r._ziptax_total_rate || '—')}</strong></div>
          </div>
          ${r._ziptax_normalized_address?`<div class="normalized-address">${esc(r._ziptax_normalized_address)}</div>`:''}
          <div class="tax-note">${esc(note || 'ZipTax runs automatically during import. Use the dropdown below to override if needed.')}</div>
        </section>
      </div>
      <div class="mapping-grid tax-edit-grid">
        <label class="field emphasis tax-field"><span>RTO Tax Zone</span><select class="tax-select" data-i="${i}">${taxOptions(r)}</select><small>Auto-selected when a unique RTO zone can be identified; always editable.</small></label>
        ${inputField(i,'TAXRATE','Tax rate',{emphasis:true})}
        ${inputField(i,'COUNTY','County')}
        <label class="field"><span>RTO Zone</span><select class="zone-select" data-i="${i}">${zoneOptions(r)}</select><small>${esc(r._zone_selector || '')}</small></label>
      </div>
    </div>`;
  }

  function pdfTab(r,i) {
    const pdfLink=r._pdf_available ? `<a class="button ghost" target="_blank" href="/pdf/${encodeURIComponent(window.RTO_JOB)}/${i}">Open combined PDF</a>` : '<span class="mini-muted">No PDF loaded</span>';
    return `<div class="tab-section">
      <div class="pdf-summary">
        <div class="pdf-coordinates"><span>Coordinates extracted from PDF</span><strong>${esc(r._pdf_coordinates || 'Coordinates not found')}</strong>${r._pdf_coordinates?`<button type="button" class="mini-btn copy-coordinates" data-i="${i}">Put in Directions</button>`:''}</div>
        <div class="pdf-actions">${pdfLink}</div>
      </div>
      <div class="mapping-grid pdf-grid">
        ${inputField(i,'SALESMAN','PDF Salesman / Agent',{emphasis:true,small:r._pdf_salesman?`Extracted: ${r._pdf_salesman}`:'No PDF agent extracted'})}
        ${inputField(i,'AGENT1','RTO Agent (Firebird)',{small:'This is the RTO Pro agent for the selected Store.'})}
        ${inputField(i,'INVOICE1','Invoice #',{small:r._pdf_invoice?`Extracted: ${r._pdf_invoice}`:''})}
        ${inputField(i,'DUEDATE','Next due')}
        ${inputField(i,'SACDATE','90-day SAC')}
      </div>
      ${textareaField(i,'DIRECTIONS','Directions / coordinates',{rows:4,wide:true})}
      <div class="pdf-note">The Invoice PDF is now always included in the download set because that is where ShedSuite commonly stores the <strong>Agent:</strong> field.</div>
    </div>`;
  }

  function tabContent(r,i,tab) {
    if (tab==='inventory') return inventoryTab(r,i);
    if (tab==='customer') return customerTab(r,i);
    if (tab==='tax') return taxTab(r,i);
    if (tab==='pdf') return pdfTab(r,i);
    return mappingTab(r,i);
  }

  function buildCard(r,i) {
    const probs=rowProblems(r); const needs=probs.length>0;
    const tab=activeTabs[i] || 'mapping';
    const warningHtml=Array.isArray(r._warnings) && r._warnings.length ? `<details class="row-warnings"><summary>${r._warnings.length} original check${r._warnings.length===1?'':'s'}</summary>${r._warnings.map(x=>`<div>${esc(x)}</div>`).join('')}</details>` : '';
    return `<article class="contract-card ${needs?'needs-review':'ready'}" data-index="${i}">
      <header class="contract-header">
        <div class="contract-identity"><span class="status-dot"></span><div><h2>${esc(r.NAME || 'Unnamed')}</h2><div class="identity-meta"><span>Order <strong>${esc(r._source_order_id)}</strong></span><span>Serial <strong>${esc(r.SERIAL1)}</strong></span><span>${esc(r._source_company)}</span></div></div></div>
        <div class="contract-header-actions">${boolish(r._used_building)?'<span class="used-header-tag">USED BUILDING</span>':(boolish(r._used_detected)?'<span class="used-header-tag detected">USED DETECTED</span>':'')}<span class="state-chip">${esc(r.del_state || r._source_state)}</span><button type="button" class="mini-btn source-btn" data-i="${i}">Source row</button><button type="button" class="mini-btn all-btn" data-i="${i}">All RTO fields</button></div>
      </header>
      <div class="source-strip"><span class="source-label">ShedSuite dealer</span><strong>${esc(r._source_dealer || '—')}</strong>${sourceMatchText(r)}${warningHtml}</div>
      <nav class="card-tabs" aria-label="Contract sections">
        ${[['mapping','Mapping'],['inventory','Inventory'],['customer','Customer'],['tax','Tax & Address'],['pdf','PDF / Agent']].map(([v,l])=>`<button type="button" class="card-tab ${tab===v?'active':''}" data-i="${i}" data-tab="${v}">${l}</button>`).join('')}
      </nav>
      <div class="card-tab-content">${tabContent(r,i,tab)}</div>
      <div class="card-foot"><span class="problem-text">${needs?`Needs: ${esc(probs.join(', '))}`:'✓ Required mapping fields are filled'}</span></div>
    </article>`;
  }

  function visibleIndexes() {
    const q=clean(searchInput.value).toLowerCase(); const needsOnly=onlyNeeds.checked;
    const out=[];
    rows.forEach((r,i)=>{
      const hay=[r.NAME,r._source_order_id,r.SERIAL1,r._source_dealer,r._source_company,r.MODEL1,r.CONTRACT,r.DEALERID,r.ADDRESS,r.del_address1,r.CATEGORY1,r.DESCRIPTION1].join(' ').toLowerCase();
      if (q && !hay.includes(q)) return;
      if (needsOnly && rowProblems(r).length===0) return;
      out.push(i);
    });
    return out;
  }

  function updateStats() {
    const needs=rows.filter(r=>rowProblems(r).length).length;
    document.getElementById('needsCount').textContent=needs;
    document.getElementById('readyCount').textContent=rows.length-needs;
  }

  function render() {
    const idxs=visibleIndexes();
    cards.innerHTML=idxs.length ? idxs.map(i=>buildCard(rows[i],i)).join('') : '<div class="empty-state">No contracts match the current filter.</div>';
    updateStats();
    bindCardEvents();
  }

  function setField(i, field, value, rerender=false) {
    rows[i][field]=value;
    if (field==='STORE') rows[i]._store_title=findStoreTitle(value);
    if (field==='DEALERID') rows[i]._dealer_name=findDealerName(value);
    if (rerender) render();
  }

  function remapDealerForStore(i) {
    const r=rows[i];
    const candidates=dealers.filter(d=>dealerStore(d)===clean(r.STORE));
    const exact=candidates.filter(d=>norm(dealerName(d))===norm(r._source_dealer));
    if (exact.length===1) {
      r.DEALERID=dealerId(exact[0]); r._dealer_name=dealerName(exact[0]);
    } else {
      r.DEALERID=''; r._dealer_name='';
    }
    const a=findAgent(r.STORE); if (a) r.AGENT1=a;
  }

  function phoneCommentFor(r, source) {
    if (source==='ref1') return `OTHER IS ${clean(r.REFRELATION1)} ${clean(r.REFERENCE1)}`.trim().toUpperCase();
    if (source==='ref2') return `OTHER IS ${clean(r.REFRELATION2)} ${clean(r.REFERENCE2)}`.trim().toUpperCase();
    if (source==='secondary' && clean(r._secondary_phone)) return 'OTHER IS SECONDARY PHONE';
    return '';
  }

  function applyPhoneSource(i, source, manual=false) {
    const r=rows[i];
    r._phone_other_source=source;
    r._phone_other_manual=manual;
    if (source==='ref1') r.PHONE5=clean(r.REFPHONE1);
    else if (source==='ref2') r.PHONE5=clean(r.REFPHONE2);
    else r.PHONE5=clean(r._secondary_phone);
    r.COMMENTS=phoneCommentFor(r,source);
  }

  function applyPhoneRule(i, forceAuto=false) {
    const r=rows[i];
    if (r._phone_other_manual && !forceAuto) {
      applyPhoneSource(i, clean(r._phone_other_source || 'ref1'), true);
      return;
    }
    const p=digits(r.CELL), s=digits(r._secondary_phone), r1=digits(r.REFPHONE1), r2=digits(r.REFPHONE2);
    let source='secondary';
    if (!s) source='ref1';
    else if (p && s===p) source='ref1';
    else if (r2 && s===r2) source='ref2';
    else if (r1 && s===r1) source='ref1';
    applyPhoneSource(i,source,false);
  }

  async function redetectTax(i, button) {
    const r=rows[i];
    const original=button.textContent;
    button.disabled=true; button.textContent='Checking ZipTax…';
    try {
      const resp=await fetch(`/api/ziptax/${encodeURIComponent(window.RTO_JOB)}/${i}`,{
        method:'POST', headers:{'Content-Type':'application/json'},
        body:JSON.stringify({street:r.del_address1,city:r.del_city,state:r.del_state,zip:r.del_zip,csv_rate:r.TAXRATE})
      });
      const data=await resp.json();
      r._ziptax_county=clean(data.county);
      r._ziptax_city=clean(data.city);
      r._ziptax_incorporated=data.incorporated;
      r._ziptax_normalized_address=clean(data.normalized_address);
      r._ziptax_total_rate=clean(data.total_rate);
      r._ziptax_response_code=clean(data.response_code);
      r._ziptax_note=clean(data.note || data.error);
      if (data.county) r.COUNTY=data.county;
      if (data.zone) r.TAXZONE=data.zone;
      if (data.rate) r.TAXRATE=data.rate;
      render();
    } catch (e) {
      r._ziptax_note=`ZipTax request failed: ${e}`;
      render();
    } finally {
      button.disabled=false; button.textContent=original;
    }
  }

  function bindCardEvents() {
    document.querySelectorAll('.card-tab').forEach(btn=>btn.addEventListener('click',()=>{
      activeTabs[+btn.dataset.i]=btn.dataset.tab; render();
    }));
    document.querySelectorAll('input[data-field], textarea[data-field]').forEach(el=>el.addEventListener('input',e=>{
      const i=+e.target.dataset.i; const field=e.target.dataset.field; setField(i,field,e.target.value,false);
      if (['CELL','REFPHONE1','REFPHONE2','REFERENCE1','REFRELATION1','REFERENCE2','REFRELATION2'].includes(field)) applyPhoneRule(i,false);
      if (['MODEL1','CONTRACT','DEALERID'].includes(field)) updateStats();
    }));
    document.querySelectorAll('input[data-field="MODEL1"]').forEach(el=>el.addEventListener('change',e=>{
      const i=+e.target.dataset.i; if (boolish(rows[i]._used_building)) { applyUsedBuilding(i,true,false); render(); }
    }));
    document.querySelectorAll('.used-building-toggle').forEach(el=>el.addEventListener('change',e=>{
      const i=+e.target.dataset.i; applyUsedBuilding(i,e.target.checked,true); render();
    }));
    document.querySelectorAll('.used-suffix-select').forEach(el=>el.addEventListener('change',e=>{
      const i=+e.target.dataset.i; setUsedSuffix(i,e.target.value); render();
    }));
    document.querySelectorAll('input[data-helper]').forEach(el=>el.addEventListener('input',e=>{
      const i=+e.target.dataset.i; rows[i][e.target.dataset.helper]=e.target.value;
      if (e.target.dataset.helper==='_secondary_phone') applyPhoneRule(i,false);
    }));
    document.querySelectorAll('.store-select').forEach(el=>el.addEventListener('change',e=>{
      const i=+e.target.dataset.i; rows[i].STORE=e.target.value; rows[i]._store_title=findStoreTitle(e.target.value); remapDealerForStore(i);
      const z=zones.find(z=>zoneStore(z)===clean(rows[i].STORE) && norm(zoneSelector(z))===norm(rows[i]._zone_selector));
      rows[i].ZONE=z?zoneId(z):''; render();
    }));
    document.querySelectorAll('.dealer-select').forEach(el=>el.addEventListener('change',e=>{
      const i=+e.target.dataset.i; rows[i].DEALERID=e.target.value; rows[i]._dealer_name=findDealerName(e.target.value); render();
    }));
    document.querySelectorAll('.zone-select').forEach(el=>el.addEventListener('change',e=>{ rows[+e.target.dataset.i].ZONE=e.target.value; updateStats(); }));
    document.querySelectorAll('.tax-select').forEach(el=>el.addEventListener('change',e=>{
      const i=+e.target.dataset.i; rows[i].TAXZONE=e.target.value; const opt=e.target.selectedOptions[0]; if (opt && opt.dataset.rate) rows[i].TAXRATE=opt.dataset.rate; render();
    }));
    document.querySelectorAll('.phone-source-select').forEach(el=>el.addEventListener('change',e=>{ applyPhoneSource(+e.target.dataset.i,e.target.value,true); render(); }));
    document.querySelectorAll('.auto-phone').forEach(btn=>btn.addEventListener('click',()=>{ applyPhoneRule(+btn.dataset.i,true); render(); }));
    document.querySelectorAll('.redetect-tax').forEach(btn=>btn.addEventListener('click',()=>redetectTax(+btn.dataset.i,btn)));
    document.querySelectorAll('.copy-coordinates').forEach(btn=>btn.addEventListener('click',()=>{
      const i=+btn.dataset.i; const c=clean(rows[i]._pdf_coordinates); const d=clean(rows[i].DIRECTIONS);
      if (c && !d.includes(c)) rows[i].DIRECTIONS=c+(d?`      ${d}`:''); render();
    }));
    document.querySelectorAll('.suggestion').forEach(btn=>btn.addEventListener('click',()=>{ const i=+btn.dataset.i; setField(i,btn.dataset.field,btn.dataset.value,false); if (btn.dataset.field==='MODEL1' && boolish(rows[i]._used_building)) applyUsedBuilding(i,true,false); render(); }));
    document.querySelectorAll('.copy-model').forEach(btn=>btn.addEventListener('click',()=>{ const i=+btn.dataset.i; rows[i].CONTRACT=rows[i].MODEL1; render(); }));
    document.querySelectorAll('.all-btn').forEach(btn=>btn.addEventListener('click',()=>openAllFields(+btn.dataset.i)));
    document.querySelectorAll('.source-btn').forEach(btn=>btn.addEventListener('click',()=>openSource(+btn.dataset.i)));
  }

  function validate(show=true) {
    const issues=[];
    rows.forEach((r,i)=>{
      const p=rowProblems(r);
      if (p.length) issues.push(`${r.NAME || `Row ${i+1}`}: ${p.join(', ')}`);
    });
    if (show) {
      validationBanner.classList.remove('hidden','good','bad');
      if (!issues.length) {
        validationBanner.classList.add('good');
        validationBanner.innerHTML='<strong>All contracts have the required mapping fields.</strong> Save edits before downloading the final import package.';
      } else {
        validationBanner.classList.add('bad');
        validationBanner.innerHTML=`<strong>${issues.length} contract${issues.length===1?'':'s'} need review.</strong><div>${issues.slice(0,12).map(esc).join('<br>')}${issues.length>12?'<br>…':''}</div>`;
      }
    }
    updateStats(); return issues;
  }

  function saveAll() {
    validate(true); rowsPayload.value=JSON.stringify(rows); saveForm.submit();
  }

  function openAllFields(i) {
    editingIndex=i; allFieldDraft={};
    document.getElementById('dialogTitle').textContent=`${rows[i].NAME} · all RTO fields`;
    allFieldsGrid.innerHTML=headers.map(h=>{
      const val=clean(rows[i][h]); allFieldDraft[h]=val;
      const important=['ACCOUNT','CONTRACT','MODEL1','STORE','DEALERID','ZONE','TAXZONE','TAXRATE','CATEGORY1','DESCRIPTION1'].includes(h);
      return `<label class="field ${important?'important-field':''}"><span>${esc(h)}</span><input data-all-field="${esc(h)}" value="${esc(val)}"></label>`;
    }).join('');
    allFieldsGrid.querySelectorAll('input').forEach(el=>el.addEventListener('input',()=>allFieldDraft[el.dataset.allField]=el.value));
    allFieldsDialog.showModal();
  }
  document.getElementById('applyAllFields').addEventListener('click',()=>{
    if (editingIndex===null || !allFieldDraft) return;
    headers.forEach(h=>rows[editingIndex][h]=clean(allFieldDraft[h]));
    rows[editingIndex]._store_title=findStoreTitle(rows[editingIndex].STORE) || rows[editingIndex]._store_title;
    rows[editingIndex]._dealer_name=findDealerName(rows[editingIndex].DEALERID) || rows[editingIndex]._dealer_name;
    if (boolish(rows[editingIndex]._used_building)) applyUsedBuilding(editingIndex,true,false);
    allFieldsDialog.close(); editingIndex=null; allFieldDraft=null; render();
  });

  function openSource(i) {
    const src=sources[i] || {};
    document.getElementById('sourceTitle').textContent=`${rows[i].NAME} · ShedSuite source`;
    sourceGrid.innerHTML=Object.entries(src).map(([k,v])=>`<div><span>${esc(k)}</span><strong>${esc(v || '—')}</strong></div>`).join('');
    sourceDialog.showModal();
  }

  function populateBulkStore() { document.getElementById('bulkStore').innerHTML=storeOptions(''); }
  document.getElementById('applyBulkStore').addEventListener('click',()=>{
    const val=document.getElementById('bulkStore').value; if (!val) return;
    visibleIndexes().forEach(i=>{ rows[i].STORE=val; rows[i]._store_title=findStoreTitle(val); remapDealerForStore(i); const z=zones.find(z=>zoneStore(z)===val && norm(zoneSelector(z))===norm(rows[i]._zone_selector)); rows[i].ZONE=z?zoneId(z):''; });
    render();
  });
  document.getElementById('validateBtn').addEventListener('click',()=>validate(true));
  document.getElementById('saveTopBtn').addEventListener('click',saveAll);
  searchInput.addEventListener('input',render); onlyNeeds.addEventListener('change',render);
  populateBulkStore(); render();
})();
