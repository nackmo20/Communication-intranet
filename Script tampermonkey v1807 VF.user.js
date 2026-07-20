// ==UserScript==
// @name         EAFC Drupal - Mise à jour parcours depuis Mapping JSON + Excel
// @namespace    https://eafc.local/tampermonkey
// @version      1.8.09
// @description  Met à jour des pages de parcours Drupal depuis un mapping JSON externe et un export Excel EAFC Formation Export plan et sessions.
// @author       EAFC
// @match        *://*/*
// @run-at       document-idle
// @grant        GM_addStyle
// @grant        GM_download
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_deleteValue
// @require      https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js
// ==/UserScript==

(function () {
  'use strict';

  const SCRIPT_VERSION = '1.8.09';

  /***************************************************************************
   * CONFIGURATION GLOBALE
   ***************************************************************************/
  const CONFIG = {
    DRY_RUN: true,
    AUTO_SAVE: false,
    SAFE_CLICK_MODE: true,
    DELAY_BETWEEN_ACTIONS: 500,
    MAX_WAIT_TIME: 15000,
    ADMIN_CONTENT_URL: '/admin/content',
    PREINSCRIPTION_BUTTON_TEXT: 'Préinscrivez vous',
    CALLOUT_BODY_TEXT: '<h2>\n    Préinscriptions\n</h2>',
    CALLOUT_COLOR_VALUE: 'blue-cumulus',
    ADVANCED_FORMAT_VALUE: 'formatage_avance',
    // Correspondance demandée : le mapping JSON donne l'identifiant dispositif,
    // l'Excel est retrouvé via la colonne exacte/souple "Dispositif : code".
    // Le module reste conservé pour les logs et diagnostics, mais n'est plus bloquant.
    MATCH_EXCEL_BY_DISPOSITIF_ONLY: true,
    THEMATIC_AUTOCOMPLETE_DELAY_MS: 800,
    THEMATIC_AUTOCOMPLETE_SELECTION_DELAY_MS: 600,
    // Thématiques obligatoires ajoutées à chaque fiche Drupal.
    REQUIRED_THEMATICS: ['EAFC poitiers', '6331'],
    // Mapping conservé du module isolé : alias source -> identifiants Drupal existants.
    THEMATIC_ALIAS_MAP: [
      [['Pédagogique', 'Pratique pédagogique transversales', 'Pratiques pédagogiques transversales'], ['6624']],
      [['Numérique', 'Compétences, culture et usages du numérique'], ['6362']],
      [['QVCT, SST', 'QVCT', 'SST', 'Qualité de vie au travail – Santé et sécurité au travail'], ['6385', '6404']],
      [['Valeurs de la république, citoyenneté', 'Valeurs de la République et citoyenneté', 'Valeurs de la république', 'citoyenneté'], ['6627', '6236']],
      [['Orientation', 'Orientation et parcours scolaire'], ['6364']],
      [['Ecole inclusive', 'École inclusive', 'Ecole inclusive, accessible et ouverte à tous'], ['6314']],
      [['Europe et international', 'Ouverture européenne et internationale'], ['6328']],
      [['Formateurs', 'Formation de formateurs et tuteurs'], ['6625']],
      [['Santé des élèves', 'Santé mentale et bien-être des élèves'], ['6403']],
      [['Encadrement'], ['6455']],
      [['Carrière', 'Carrière et évolution professionnelle'], ['6227']],
      [['2nd degré', 'Second degré'], ['6202']]
    ]
  };

  const SELECTORS = {
    adminMenu: [
      'button#toolbar-item-administration',
      'button[title="Menu d\'administration"]',
      '.toolbar-icon-menu.trigger.toolbar-item'
    ],
    contentLink: [
      'a[href="/admin/content"]',
      'a[data-drupal-link-system-path="admin/content"]'
    ],
    nidFilter: [
      'input[data-drupal-selector="edit-nid"]',
      'input#edit-nid',
      'input[name="nid"]'
    ],
    filterButton: [
      'input[data-drupal-selector="edit-submit-content"]',
      'input#edit-submit-content',
      'input[value="Filtrer"]'
    ],
    resultTitleLink: [
      'table.views-table tbody tr:first-child td[headers="view-title-table-column"].views-field-title > a[href]',
      'table.views-table tbody tr:first-child td[headers="view-title-table-column"] > a[href]',
      'td[headers="view-title-table-column"].views-field-title > a[href]',
      'td[headers="view-title-table-column"] > a[href]',
      'td.views-field-title > a[href]'
    ],
    editLinks: [
      'a[href$="/edit"]',
      'a[href*="/edit?"]',
      'a[href*="/node/"][href*="/edit"]',
      'a:contains("Modifier")'
    ],
    startDate: [
      'input[data-drupal-selector="edit-field-event-start-date-0-value-date"]',
      'input#edit-field-event-start-date-0-value-date',
      'input[name="field_event_start_date[0][value][date]"]'
    ],
    endDate: [
      'input[data-drupal-selector="edit-field-event-end-date-0-value-date"]',
      'input#edit-field-event-end-date-0-value-date',
      'input[name="field_event_end_date[0][value][date]"]'
    ],
    unpublishDate: [
      'input[data-drupal-selector="edit-unpublish-on-0-value-date"]',
      'input#edit-unpublish-on-0-value-date',
      'input[name="unpublish_on[0][value][date]"]'
    ],
    formationBodyEdit: [
      'input[name="field_formation_body_0_edit"][value="Modifier"]',
      'input[data-drupal-selector^="field-formation-body-0-edit"][value="Modifier"]',
      'input.paragraphs-icon-button-edit[title="Modifier"]',
      'input.paragraphs-icon-button-edit[value="Modifier"]'
    ],
    textFormat: [
      'select[data-drupal-selector*="field-richtext-content"][data-drupal-selector$="-format"]',
      'select[name*="[field_richtext_content]"][name$="[format]"]',
      'select.editor.form-select'
    ],
    ckSourceButton: [
      'button.ck-source-editing-button',
      'button[data-cke-tooltip-text="Source"]',
      'button[aria-labelledby*="ck-editor"]'
    ],
    ckSourceArea: [
      '.ck-source-editing-area textarea',
      'textarea.ck-source-editing-area',
      '.ck-source-editing-area [contenteditable="true"]',
      'textarea[data-drupal-selector*="field-richtext-content"]:not([aria-hidden="true"])',
      'textarea[name*="field_richtext_content"]'
    ],
    addParagraph: [
      'input.field-add-more-submit.paragraph-type-add-modal-button',
      'input[data-drupal-selector*="field-formation-body-add-more"]',
      'input[value="Ajouter Paragraphe"]'
    ],
    calloutOverlay: [
      '.paragraphs-button--icon_overlay--wrapper',
      '.paragraphs-button--icon_overlay'
    ],
    calloutIcon: [
      'button[name="field_formation_body_callout_add_more"] .paragraphs-button--icon[style*="callout-default-icon"]',
      'button[data-drupal-selector^="field-formation-body-callout-add-more"] .paragraphs-button--icon[style*="callout-default-icon"]',
      '.paragraphs-button--icon[style*="callout-default-icon"]'
    ],
    calloutButton: [
      'button[name="field_formation_body_callout_add_more"][value="Mise en avant"]',
      'button[data-drupal-selector^="field-formation-body-callout-add-more"]',
      'button.field-add-more-submit.paragraphs-button--add-more[value="Mise en avant"]',
      'button.paragraphs-button--add-more[value="Mise en avant"]',
      'button[value="Mise en avant"]',
      'input[type="submit"][value="Mise en avant"]',
      'button[data-drupal-selector*="mise-en-avant"]',
      'input[data-drupal-selector*="mise-en-avant"]'
    ],
    calloutUri: [
      'input[data-drupal-selector^="edit-field-formation-body-"][data-drupal-selector$="-subform-field-callout-cta-0-uri"]',
      'input[name^="field_formation_body["][name$="[subform][field_callout_cta][0][uri]"]',
      'input[name*="[field_callout_cta]"][name$="[uri]"]',
      'input[data-drupal-selector*="field-callout-cta"][data-drupal-selector$="-uri"]',
      'input.form-autocomplete[name*="field_callout_cta"]'
    ],
    calloutTitle: [
      'input[data-drupal-selector^="edit-field-formation-body-"][data-drupal-selector$="-subform-field-callout-cta-0-title"]',
      'input[name^="field_formation_body["][name$="[subform][field_callout_cta][0][title]"]',
      'input[name*="[field_callout_cta]"][name$="[title]"]',
      'input[data-drupal-selector*="field-callout-cta"][data-drupal-selector$="-title"]'
    ],
    calloutContentFormat: [
      'select[data-drupal-selector^="edit-field-formation-body-"][data-drupal-selector$="-subform-field-callout-content-0-format"]',
      'select[name^="field_formation_body["][name$="[subform][field_callout_content][0][format]"]'
    ],
    calloutSourceButton: [
      'button.ck-source-editing-button',
      'button[data-cke-tooltip-text="Source"]',
      'span.ck-button__label'
    ],
    calloutSourceTextarea: [
      'textarea[aria-label="Source code editing area"]',
      '.ck-source-editing-area textarea'
    ],
    calloutColor: [
      'select[data-drupal-selector^="edit-field-formation-body-"][data-drupal-selector$="-subform-field-callout-color"]',
      'select[name^="field_formation_body["][name$="[subform][field_callout_color]"]',
      'select[name*="[field_callout_color]"]',
      'select[data-drupal-selector*="field-callout-color"]'
    ],
    title: [
      'input[data-drupal-selector="edit-title-0-value"]',
      'input#edit-title-0-value',
      'input[name="title[0][value]"]'
    ],
    thematics: [
      'input[name^="field_formation_thematic["][name$="[target_id]"]',
      'input[data-drupal-selector^="edit-field-formation-thematic-"][data-autocomplete-path]',
      'input[data-drupal-selector^="edit-field-formation-thematic-"][data-drupal-selector$="-target-id"]',
      'input[name*="field_thematique"]',
      'input[name*="field_theme"]'
    ],
    thematicAddMore: [
      '[data-drupal-selector="edit-field-formation-thematic-add-more"]',
      'input[name="field_formation_thematic_add_more"]'
    ],
    publicTreeWrapper: ['#edit-field-metier-tags-wrapper'],
    save: [
      'input[data-drupal-selector="edit-submit"]',
      'input#edit-submit',
      'input[name="op"][value="Enregistrer"]',
      'input.button--primary[value="Enregistrer"]'
    ]
  };

  const COLUMN_NAMES = {
    dispositif: ['Dispositif : code', 'dispositif : code', 'dispositif code', 'dispositif', 'identifiant', 'code dispositif', 'dispositifCode', 'Dispositif', 'Identifiant'],
    module: ['module', 'code module', 'moduleCode', 'Code', 'Module', 'Code module'],
    rne: ['RNE', 'rne', 'UAI', 'uai', 'code RNE', 'code UAI', 'Établissement : RNE', 'Etablissement : RNE'],
    groupId: ['Groupe : identifiant', 'Groupe : code', 'groupe', 'groupId', 'Identifiant groupe', 'id groupe'],
    sessionLabel: ['Session : libellé', 'Session : code', 'session', 'sessionLabel', 'Code session', 'Identifiant session'],
    territoire: ['territoire', 'territoireCode', 'département', 'departement', 'Département', 'code département', 'code departement'],
    modalite: ['modalité', 'modalite', 'Modalité', 'Modalite', 'modalité session', 'modalite session', 'type formation'],
    sessionStart: ['Début', 'Debut', 'début', 'debut', 'Session : début', 'Session : debut', 'Session : date début', 'Session : date debut', 'Date début session', 'Date debut session', 'date et heure de début', 'date et heure de debut', 'Début session', 'Debut session'],
    sessionEnd: ['Fin', 'fin', 'Session : fin', 'Session : date fin', 'Date fin session', 'date et heure de fin', 'Fin session'],
    lieu: ['lieu', 'Lieu', 'Lieux', 'site', 'localisation'],
    start: ['Préinscriptions : Publication : début', 'Publication début', 'début publication', 'date début préinscription', 'date debut preinscription', 'startDate'],
    end: ['Préinscriptions : Publication : fin', 'Publication fin', 'fin publication', 'date fin préinscription', 'date fin preinscription', 'endDate'],
    link: ['Préinscriptions : Publication : lien préinscription', 'lien préinscription', 'lien de préinscription', 'url préinscription', 'url', 'link'],
    total: ['Préinscriptions : Préinscrits : nombre total', 'nombre total', 'préinscrits total', 'inscrits', 'total'],
    max: ['Préinscriptions : Préinscrits : nombre maximal', 'nombre maximal', 'capacité', 'capacité maximale', 'max', 'places']
  };

  const STORAGE_KEYS = {
    mappingItems: 'eafcDrupalUpdater.mappingItems.v1',
    excelRows: 'eafcDrupalUpdater.excelRows.v1',
    savedAt: 'eafcDrupalUpdater.savedAt.v1',
    batchState: 'eafcDrupalUpdater.batchState.v1',
    uiConfig: 'eafcDrupalUpdater.uiConfig.v1',
    updateReport: 'eafcDrupalUpdater.updateReport.v1',
    intranetPayload: 'eafcDrupalUpdater.intranetPayload.v1',
    intranetPayloadFileName: 'eafcDrupalUpdater.intranetPayloadFileName.v1',
    intranetBatchState: 'eafcDrupalUpdater.intranetBatchState.v1',
    intranetReport: 'eafcDrupalUpdater.intranetReport.v1'
  };

  const state = {
    mappingItems: [],
    excelRows: [],
    excelIndex: new Map(),
    enrichedItems: [],
    validationErrors: [],
    logs: [],
    updateReport: [],
    running: false,
    stopRequested: false,
    panelCollapsed: false,
    lastSummary: { mapping: 0, excel: 0, matches: 0, errors: 0 },
    activeWorkflow: 'sofia',
    intranetUpdate: {
      payload: null, payloadType: '', payloadFileName: '', normalizedTargets: [],
      validationErrors: [], validationWarnings: [], preview: [], reportItems: [],
      startedAt: '', finishedAt: '', backupItems: [], backupGenerated: false
    }
  };

  /***************************************************************************
   * INTERFACE TAMPERMONKEY
   ***************************************************************************/
  function initPanel() {
    GM_addStyle(`
      #eafc-updater-panel{position:fixed;right:18px;bottom:18px;z-index:2147483647;width:390px;max-height:88vh;background:#fff;border:2px solid #000091;border-radius:10px;box-shadow:0 8px 28px rgba(0,0,0,.25);font:13px/1.35 Arial,sans-serif;color:#161616;overflow:hidden}
      #eafc-updater-panel.eafc-collapsed{width:260px}.eafc-header{display:flex;align-items:center;justify-content:space-between;gap:8px;background:#000091;color:#fff;padding:10px 12px;font-weight:700}.eafc-toggle{background:#fff;color:#000091;border:0;border-radius:4px;padding:4px 8px;cursor:pointer}.eafc-body{padding:12px;max-height:calc(88vh - 45px);overflow:auto}.eafc-collapsed .eafc-body{display:none}.eafc-row{display:flex;gap:8px;flex-wrap:wrap;margin:8px 0}.eafc-btn{border:1px solid #000091;background:#000091;color:#fff;border-radius:4px;padding:7px 9px;cursor:pointer}.eafc-btn.secondary{background:#fff;color:#000091}.eafc-btn.danger{background:#ce0500;border-color:#ce0500}.eafc-btn:disabled{opacity:.55;cursor:not-allowed}.eafc-file{display:none}.eafc-summary{display:grid;grid-template-columns:1fr auto;gap:4px;background:#f6f6f6;border-radius:6px;padding:8px;margin:8px 0}.eafc-options label{display:block;margin:6px 0}.eafc-number{width:95px;max-width:100%;padding:3px 5px;border:1px solid #aaa;border-radius:4px}.eafc-alert{background:#fff4d6;border-left:4px solid #ffb000;padding:8px;margin:8px 0}.eafc-errors{max-height:130px;overflow:auto;border:1px solid #ddd;margin-top:8px}.eafc-errors table{border-collapse:collapse;width:100%;font-size:11px}.eafc-errors th,.eafc-errors td{border:1px solid #ddd;padding:3px;text-align:left}.eafc-logbox{height:180px;overflow:auto;background:#1e1e1e;color:#eee;border-radius:6px;padding:8px;font:11px/1.35 Consolas,monospace}.eafc-log-info{color:#d5e9ff}.eafc-log-success{color:#b8f5c2}.eafc-log-warning{color:#ffe08a}.eafc-log-error{color:#ffb3ad}.eafc-small{font-size:11px;color:#666}.eafc-version{font-size:11px;font-weight:700;color:#fff;background:rgba(255,255,255,.18);border:1px solid rgba(255,255,255,.45);border-radius:999px;padding:2px 7px;white-space:nowrap}.eafc-status{font-weight:700;margin-top:8px}
    `);

    const panel = document.createElement('div');
    panel.id = 'eafc-updater-panel';
    panel.innerHTML = `
      <div class="eafc-header"><span>Mise à jour parcours EAFC</span><span class="eafc-version">v${SCRIPT_VERSION}</span><button class="eafc-toggle" type="button">−</button></div>
      <div class="eafc-body">
        <input id="eafc-mapping-file" class="eafc-file" type="file" accept="application/json,.json">
        <input id="eafc-intranet-file" class="eafc-file" type="file" accept="application/json,.json">
        <input id="eafc-excel-file" class="eafc-file" type="file" accept=".xlsx,.xls,.xlsm,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel">
        <div class="eafc-row">
          <button id="eafc-import-mapping" class="eafc-btn secondary" type="button">Importer mapping JSON</button>
          <button id="eafc-import-excel" class="eafc-btn secondary" type="button">Importer Excel</button>
          <button id="eafc-import-intranet" class="eafc-btn secondary" type="button">Importer mises à jour intranet</button>
        </div>
        <div id="eafc-workflow-mode" class="eafc-alert"></div>
        <div id="eafc-intranet-preview" class="eafc-small"></div>
        <div class="eafc-summary" id="eafc-summary"></div>
        <div class="eafc-options">
          <label><input id="eafc-dry-run" type="checkbox" checked> Mode simulation</label>
          <label><input id="eafc-auto-save" type="checkbox"> Enregistrer automatiquement</label>
          <label><input id="eafc-safe-click" type="checkbox" checked> Mode sécurité : validation à chaque clic</label>
          <label>Vitesse des clics (ms) <input id="eafc-click-delay" class="eafc-number" type="number" min="0" step="50" value="500"></label>
        </div>
        <div class="eafc-alert" id="eafc-alert">Importez le mapping JSON et l'Excel, puis analysez les correspondances.</div>
        <div class="eafc-row">
          <button id="eafc-analyze" class="eafc-btn secondary" type="button">Analyser les correspondances</button>
          <button id="eafc-start" class="eafc-btn" type="button">Démarrer</button>
          <button id="eafc-stop" class="eafc-btn danger" type="button">Stop</button>
          <button id="eafc-export" class="eafc-btn secondary" type="button">Exporter les logs</button>
          <button id="eafc-export-report" class="eafc-btn secondary" type="button">Exporter bilan Excel</button>
          <button id="eafc-export-intranet-report" class="eafc-btn secondary" type="button">Exporter rapport JSON</button>
          <button id="eafc-reset-intranet" class="eafc-btn secondary" type="button">Reset JSON maj public/contenu</button>
          <button id="eafc-reset-excel" class="eafc-btn secondary" type="button">Retirer Excel</button>
          <button id="eafc-reset-imports" class="eafc-btn danger" type="button">Reset JSON/Excel</button>
        </div>
        <div id="eafc-errors" class="eafc-errors" hidden></div>
        <div class="eafc-status" id="eafc-status">Prêt.</div>
        <div class="eafc-logbox" id="eafc-logbox"></div>
        <div class="eafc-small">Les codes Sofia-FMO/GAIA servent uniquement à la correspondance et ne sont jamais injectés dans le HTML Drupal.</div>
      </div>`;
    document.body.appendChild(panel);
    loadPersistedUpdateReport();

    qs('#eafc-import-mapping').addEventListener('click', () => qs('#eafc-mapping-file').click());
    qs('#eafc-import-intranet').addEventListener('click', () => qs('#eafc-intranet-file').click());
    qs('#eafc-intranet-file').addEventListener('change', (e) => e.target.files[0] && handleIntranetUpdateFile(e.target.files[0]));
    qs('#eafc-import-excel').addEventListener('click', () => qs('#eafc-excel-file').click());
    qs('#eafc-mapping-file').addEventListener('change', (e) => e.target.files[0] && handleMappingFile(e.target.files[0]));
    qs('#eafc-excel-file').addEventListener('change', (e) => e.target.files[0] && handleExcelFile(e.target.files[0]));
    applyUiConfigToPanel();
    qs('#eafc-dry-run').addEventListener('change', (e) => { CONFIG.DRY_RUN = e.target.checked; persistUiConfig(); renderSummary(); });
    qs('#eafc-auto-save').addEventListener('change', (e) => { CONFIG.AUTO_SAVE = e.target.checked; persistUiConfig(); renderSummary(); });
    qs('#eafc-safe-click').addEventListener('change', (e) => { CONFIG.SAFE_CLICK_MODE = e.target.checked; persistUiConfig(); renderSummary(); });
    qs('#eafc-click-delay').addEventListener('change', (e) => { CONFIG.DELAY_BETWEEN_ACTIONS = sanitizeDelay(e.target.value); e.target.value = String(CONFIG.DELAY_BETWEEN_ACTIONS); persistUiConfig(); renderSummary(); });
    qs('#eafc-analyze').addEventListener('click', analyzeCorrespondences);
    qs('#eafc-start').addEventListener('click', runUpdateBatch);
    qs('#eafc-stop').addEventListener('click', () => { markRemainingIntranetTargetsCancelled(); state.stopRequested = true; clearBatchState(); log('warning', null, 'Arrêt demandé par l’utilisateur.'); setStatus('Arrêt demandé…'); });
    qs('#eafc-export').addEventListener('click', exportLogs);
    qs('#eafc-export-report').addEventListener('click', exportUpdateReportExcel);
    qs('#eafc-export-intranet-report').addEventListener('click', exportIntranetReportJson);
    qs('#eafc-reset-intranet').addEventListener('click', resetIntranetUpdateData);
    qs('#eafc-reset-excel').addEventListener('click', resetExcelData);
    qs('#eafc-reset-imports').addEventListener('click', resetImportedData);
    qs('.eafc-toggle').addEventListener('click', () => {
      state.panelCollapsed = !state.panelCollapsed;
      panel.classList.toggle('eafc-collapsed', state.panelCollapsed);
      qs('.eafc-toggle').textContent = state.panelCollapsed ? '+' : '−';
    });
    renderSummary();
  }

  function renderSummary() {
    const s = state.lastSummary;
    qs('#eafc-summary').innerHTML = `
      <span>Lignes mapping</span><strong>${s.mapping}</strong>
      <span>Lignes Excel</span><strong>${s.excel}</strong>
      <span>Correspondances trouvées</span><strong>${s.matches}</strong>
      <span>Erreurs</span><strong>${s.errors}</strong>
      <span>Simulation</span><strong>${CONFIG.DRY_RUN ? 'Oui' : 'Non'}</strong>
      <span>Auto-save</span><strong>${CONFIG.AUTO_SAVE ? 'Oui' : 'Non'}</strong>
      <span>Sécurité clic</span><strong>${CONFIG.SAFE_CLICK_MODE ? 'Oui' : 'Non'}</strong>
      <span>Vitesse clics</span><strong>${CONFIG.DELAY_BETWEEN_ACTIONS} ms</strong>`;
    renderWorkflowMode();
  }

  function renderErrors() {
    const box = qs('#eafc-errors');
    if (!state.validationErrors.length) { box.hidden = true; box.innerHTML = ''; return; }
    box.hidden = false;
    box.innerHTML = `<table><thead><tr><th>ID Drupal</th><th>Titre</th><th>Erreur</th></tr></thead><tbody>${state.validationErrors.map((e) => `<tr><td>${escapeHtml(e.item?.drupalParcoursId || '')}</td><td>${escapeHtml(e.item?.drupalTitle || '')}</td><td>${escapeHtml(e.message)}</td></tr>`).join('')}</tbody></table>`;
  }


  function sanitizeDelay(value) {
    const parsed = Number.parseInt(value, 10);
    if (!Number.isFinite(parsed) || parsed < 0) return 0;
    return Math.min(parsed, 10000);
  }

  /***************************************************************************
   * PERSISTANCE DES CHOIX UTILISATEUR
   ***************************************************************************/
  function loadPersistedUiConfig() {
    const saved = readStoredJson(STORAGE_KEYS.uiConfig, null);
    if (!saved) return;
    CONFIG.DRY_RUN = saved.DRY_RUN !== undefined ? Boolean(saved.DRY_RUN) : CONFIG.DRY_RUN;
    CONFIG.AUTO_SAVE = saved.AUTO_SAVE !== undefined ? Boolean(saved.AUTO_SAVE) : CONFIG.AUTO_SAVE;
    CONFIG.SAFE_CLICK_MODE = saved.SAFE_CLICK_MODE !== undefined ? Boolean(saved.SAFE_CLICK_MODE) : CONFIG.SAFE_CLICK_MODE;
    CONFIG.DELAY_BETWEEN_ACTIONS = saved.DELAY_BETWEEN_ACTIONS !== undefined ? sanitizeDelay(saved.DELAY_BETWEEN_ACTIONS) : CONFIG.DELAY_BETWEEN_ACTIONS;
  }

  function persistUiConfig() {
    writeStoredJson(STORAGE_KEYS.uiConfig, {
      DRY_RUN: CONFIG.DRY_RUN,
      AUTO_SAVE: CONFIG.AUTO_SAVE,
      SAFE_CLICK_MODE: CONFIG.SAFE_CLICK_MODE,
      DELAY_BETWEEN_ACTIONS: CONFIG.DELAY_BETWEEN_ACTIONS
    });
  }

  function applyUiConfigToPanel() {
    qs('#eafc-dry-run').checked = CONFIG.DRY_RUN;
    qs('#eafc-auto-save').checked = CONFIG.AUTO_SAVE;
    qs('#eafc-safe-click').checked = CONFIG.SAFE_CLICK_MODE;
    qs('#eafc-click-delay').value = String(CONFIG.DELAY_BETWEEN_ACTIONS);
  }


  /***************************************************************************
   * PERSISTANCE DES IMPORTS ENTRE LES PAGES DRUPAL
   ***************************************************************************/
  function loadPersistedImports() {
    const persistedMapping = readStoredJson(STORAGE_KEYS.mappingItems, []);
    const persistedExcelRows = readStoredJson(STORAGE_KEYS.excelRows, []);
    const savedAt = readStoredValue(STORAGE_KEYS.savedAt, '');
    restoreIntranetPayload();

    if (Array.isArray(persistedMapping) && persistedMapping.length) {
      state.mappingItems = persistedMapping;
      state.lastSummary.mapping = persistedMapping.length;
    }

    if (Array.isArray(persistedExcelRows) && persistedExcelRows.length) {
      state.excelRows = persistedExcelRows;
      state.excelIndex = buildExcelIndex(state.excelRows);
      state.lastSummary.excel = persistedExcelRows.length;
    }

    if (state.mappingItems.length || state.excelRows.length) {
      analyzeCorrespondences(false);
      const restored = `Imports restaurés automatiquement${savedAt ? ` depuis ${savedAt}` : ''} : ${state.mappingItems.length} ligne(s) mapping, ${state.excelRows.length} ligne(s) Excel.`;
      setAlert(restored);
      log('success', null, restored);
      scheduleAutoResumeBatch();
    } else {
      renderSummary();
      if (readBatchState().active) scheduleAutoResumeBatch();
    }
    renderIntranetPreview();
  }

  function persistImportedData(source) {
    writeStoredJson(STORAGE_KEYS.mappingItems, state.mappingItems);
    writeStoredJson(STORAGE_KEYS.excelRows, state.excelRows);
    writeStoredValue(STORAGE_KEYS.savedAt, new Date().toLocaleString('fr-FR'));
    log('info', null, `Import ${source} mémorisé : il sera restauré automatiquement après changement de page Drupal.`);
  }


  function resetIntranetUpdateData() {
    if (!state.intranetUpdate.payload && !readStoredJson(STORAGE_KEYS.intranetPayload, null)) {
      setAlert('Aucun JSON de mise à jour public/contenu n’est actuellement chargé.');
      setStatus('Aucun JSON de mise à jour à réinitialiser.');
      return;
    }
    if (!window.confirm('Réinitialiser uniquement le JSON de mise à jour public/contenu ? Le mapping JSON Sofia et l’Excel seront conservés.')) return;

    state.intranetUpdate = {
      payload: null,
      payloadType: '',
      payloadFileName: '',
      normalizedTargets: [],
      validationErrors: [],
      validationWarnings: [],
      preview: [],
      reportItems: [],
      startedAt: '',
      finishedAt: '',
      backupItems: [],
      backupGenerated: false
    };
    state.activeWorkflow = 'sofia';
    state.stopRequested = true;

    const intranetInput = qs('#eafc-intranet-file');
    if (intranetInput) intranetInput.value = '';

    deleteStoredValue(STORAGE_KEYS.intranetPayload);
    deleteStoredValue(STORAGE_KEYS.intranetPayloadFileName);
    deleteStoredValue(STORAGE_KEYS.intranetBatchState);
    deleteStoredValue(STORAGE_KEYS.intranetReport);
    clearBatchState();
    renderWorkflowMode();
    renderIntranetPreview();
    renderSummary();
    renderErrors();
    setAlert('JSON de mise à jour public/contenu réinitialisé. Le mapping JSON Sofia et l’Excel sont conservés.');
    setStatus('JSON de mise à jour réinitialisé.');
    log('warning', null, 'JSON de mise à jour public/contenu réinitialisé ; mapping JSON Sofia et Excel conservés.');
  }


  function resetExcelData() {
    if (!state.excelRows.length && !readStoredJson(STORAGE_KEYS.excelRows, []).length) {
      setAlert('Aucun fichier Excel de mise à jour n’est actuellement chargé.');
      setStatus('Aucun Excel à retirer.');
      return;
    }
    if (!window.confirm('Retirer uniquement le fichier Excel de mise à jour ? Le mapping JSON et les imports intranet seront conservés.')) return;

    state.excelRows = [];
    state.excelIndex = new Map();
    state.enrichedItems = [];
    state.validationErrors = [];
    state.lastSummary = {
      ...state.lastSummary,
      excel: 0,
      matches: 0,
      errors: 0
    };
    state.stopRequested = true;

    const excelInput = qs('#eafc-excel-file');
    if (excelInput) excelInput.value = '';

    writeStoredJson(STORAGE_KEYS.mappingItems, state.mappingItems);
    deleteStoredValue(STORAGE_KEYS.excelRows);
    writeStoredValue(STORAGE_KEYS.savedAt, new Date().toLocaleString('fr-FR'));
    clearBatchState();
    renderSummary();
    renderErrors();
    setAlert('Fichier Excel retiré. Le mapping JSON est conservé ; importez un nouvel Excel pour relancer l’analyse Sofia.');
    setStatus('Excel retiré.');
    log('warning', null, 'Fichier Excel de mise à jour retiré ; mapping JSON conservé.');
  }


  function resetImportedData() {
    if (!window.confirm('Réinitialiser le mapping JSON, l’Excel et le batch en cours ?')) return;
    state.mappingItems = [];
    state.excelRows = [];
    state.excelIndex = new Map();
    state.enrichedItems = [];
    state.validationErrors = [];
    state.lastSummary = { mapping: 0, excel: 0, matches: 0, errors: 0 };
    state.stopRequested = true;
    deleteStoredValue(STORAGE_KEYS.mappingItems);
    deleteStoredValue(STORAGE_KEYS.excelRows);
    deleteStoredValue(STORAGE_KEYS.savedAt);
    deleteStoredValue(STORAGE_KEYS.updateReport);
    state.updateReport = [];
    clearBatchState();
    renderSummary();
    renderErrors();
    setAlert('Mapping JSON et Excel réinitialisés. Importez de nouveaux fichiers pour relancer un batch.');
    setStatus('Imports réinitialisés.');
    log('warning', null, 'Mapping JSON, Excel et batch en cours réinitialisés par l’utilisateur.');
  }

  function readStoredJson(key, fallback) {
    try {
      const value = readStoredValue(key, null);
      if (!value) return fallback;
      return typeof value === 'string' ? JSON.parse(value) : value;
    } catch (error) {
      console.warn('[EAFC] Lecture stockage impossible', key, error);
      return fallback;
    }
  }

  function writeStoredJson(key, value) {
    writeStoredValue(key, JSON.stringify(value || []));
  }

  function readStoredValue(key, fallback) {
    try {
      if (typeof GM_getValue === 'function') return GM_getValue(key, fallback);
      const value = window.localStorage.getItem(key);
      return value == null ? fallback : value;
    } catch (error) {
      console.warn('[EAFC] GM_getValue/localStorage indisponible', key, error);
      return fallback;
    }
  }

  function writeStoredValue(key, value) {
    try {
      if (typeof GM_setValue === 'function') GM_setValue(key, value);
      else window.localStorage.setItem(key, value);
    } catch (error) {
      console.warn('[EAFC] GM_setValue/localStorage indisponible', key, error);
    }
  }


  function deleteStoredValue(key) {
    try {
      if (typeof GM_deleteValue === 'function') GM_deleteValue(key);
      else window.localStorage.removeItem(key);
    } catch (error) {
      console.warn('[EAFC] GM_deleteValue/localStorage indisponible', key, error);
    }
  }

  /***************************************************************************
   * IMPORT MAPPING JSON
   ***************************************************************************/
  async function handleMappingFile(file) {
    try {
      state.activeWorkflow = 'sofia';
      const text = await file.text();
      const parsed = parseMappingJson(text);
      state.mappingItems = normalizeMappingItems(parsed.items || parsed, parsed);
      log('success', null, `${state.mappingItems.length} ligne(s) de mapping importée(s).`);
      state.lastSummary.mapping = state.mappingItems.length;
      persistImportedData('mapping');
      analyzeCorrespondences(false);
    } catch (error) {
      log('error', null, `Import mapping impossible : ${error.message}`);
    }
    renderSummary();
  }

  function parseMappingJson(json) {
    const data = typeof json === 'string' ? JSON.parse(json) : json;
    if (!data || (!Array.isArray(data) && !Array.isArray(data.items))) throw new Error('JSON invalide : items absent.');
    return data;
  }

  function normalizeMappingItems(items, globalData = {}) {
    return (items || []).map((item) => {
      const planSession = item.planSession || item.planSessions?.[0] || {};
      const updateKey = item.tampermonkeyUpdateKeys?.[0] || planSession.tampermonkeyUpdateKey || {};
      return {
        drupalParcoursId: cleanCode(item.drupalParcoursId || item.drupalNodeId || item.nid || item.id),
        drupalNodeId: cleanCode(item.drupalNodeId || item.drupalParcoursId || item.nid || item.id),
        drupalTitle: String(item.drupalTitle || item.title || item.titre || ''),
        dispositifCode: cleanCode(item.sofia?.dispositifCode || item.dispositifCode || item.identifiant || item.identifiantDispositif || item.dispositif || item['Dispositif : code'] || item['dispositif : code'] || ''),
        dispositifLabel: String(item.sofia?.dispositifLabel || item.dispositifLabel || ''),
        moduleCode: cleanCode(item.sofia?.moduleCode || item.moduleCode || item.module || ''),
        moduleLabel: String(item.sofia?.moduleLabel || item.moduleLabel || ''),
        territoireCode: normalizeTerritoireCode(item.territoire?.code || item.territoireCode || planSession.territoireCode || updateKey.territoireCode || planSession.departmentFromRne || ''),
        territoireLabel: String(item.territoire?.label || item.territoireLabel || planSession.territoireLabel || ''),
        rne: normalizeRne(item.rne || item.uai || planSession.rne || updateKey.rne || extractRneFromText(planSession.lieu || '')),
        planSessionGroupId: cleanCode(planSession.groupId || item.groupId || ''),
        planSessionLabel: cleanCode(planSession.sessionLabel || item.sessionLabel || ''),
        theme: extractThematicSourceValue(item, globalData),
        publics: extractPublicSourceValue(item, globalData),
        mappingMode: String(item.mappingMode || 'single')
      };
    });
  }

  function extractThematicSourceValue(item = {}, globalData = {}) {
    return firstDefinedValue(
      item.theme,
      item.thematiques,
      item.thematic,
      item.themes,
      item.taxonomy?.theme,
      item.taxonomy?.thematiques,
      item.changes?.theme?.value,
      item.expectedAfter?.theme,
      globalData.theme,
      globalData.thematiques,
      globalData.thematic,
      globalData.themes,
      globalData.taxonomy?.theme,
      globalData.taxonomy?.thematiques,
      globalData.changes?.theme?.value,
      globalData.expectedAfter?.theme
    );
  }

  function extractPublicSourceValue(item = {}, globalData = {}) {
    return firstDefinedValue(
      item.publics,
      item.metiersPublics,
      item.publicsMetiers,
      item.taxonomy?.publics,
      item.taxonomy?.metiersPublics,
      item.changes?.publics?.value,
      item.changes?.metiersPublics?.value,
      item.expectedAfter?.publics,
      item.expectedAfter?.metiersPublics,
      globalData.publics,
      globalData.metiersPublics,
      globalData.publicsMetiers,
      globalData.taxonomy?.publics,
      globalData.taxonomy?.metiersPublics,
      globalData.changes?.publics?.value,
      globalData.changes?.metiersPublics?.value,
      globalData.expectedAfter?.publics,
      globalData.expectedAfter?.metiersPublics
    );
  }

  function firstDefinedValue(...values) {
    return values.find((value) => value !== undefined && value !== null && !(Array.isArray(value) && !value.length) && String(value).trim() !== '');
  }

  /***************************************************************************
   * IMPORT EXCEL VIA SHEETJS
   ***************************************************************************/
  async function handleExcelFile(file) {
    try {
      state.activeWorkflow = 'sofia';
      const rows = await readExcelWorkbook(file);
      state.excelRows = normalizeExcelRows(rows);
      state.excelIndex = buildExcelIndex(state.excelRows);
      log('success', null, `${state.excelRows.length} ligne(s) Excel importée(s).`);
      state.lastSummary.excel = state.excelRows.length;
      persistImportedData('excel');
      analyzeCorrespondences(false);
    } catch (error) {
      log('error', null, `Import Excel impossible : ${error.message}`);
    }
    renderSummary();
  }

  async function readExcelWorkbook(file) {
    if (!window.XLSX) throw new Error('SheetJS XLSX non chargé.');
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) throw new Error('Aucune feuille Excel trouvée.');
    return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '', raw: true });
  }

  function normalizeExcelRows(rows) {
    return (rows || []).map((row) => {
      const lieu = String(row[findColumn(row, COLUMN_NAMES.lieu)] || '').trim();
      const rne = normalizeRne(row[findColumn(row, COLUMN_NAMES.rne)] || extractRneFromText(lieu));
      return {
        dispositifCode: cleanCode(row[findColumn(row, COLUMN_NAMES.dispositif)]),
        moduleCode: cleanCode(row[findColumn(row, COLUMN_NAMES.module)]),
        lieu,
        rne,
        territoireCode: normalizeTerritoireCode(row[findColumn(row, COLUMN_NAMES.territoire)] || departmentFromRne(rne)),
        groupId: cleanCode(row[findColumn(row, COLUMN_NAMES.groupId)]),
        sessionLabel: cleanCode(row[findColumn(row, COLUMN_NAMES.sessionLabel)]),
        modalite: stringifyCell(row[findColumn(row, COLUMN_NAMES.modalite)]),
        sessionStartRaw: row[findColumn(row, COLUMN_NAMES.sessionStart)],
        sessionEndRaw: row[findColumn(row, COLUMN_NAMES.sessionEnd)],
        sessionStartDateTime: toDisplayDateTime(row[findColumn(row, COLUMN_NAMES.sessionStart)]),
        sessionEndDateTime: toDisplayDateTime(row[findColumn(row, COLUMN_NAMES.sessionEnd)]),
        publicationStartDate: toInputDate(row[findColumn(row, COLUMN_NAMES.start)]),
        publicationEndDate: toInputDate(row[findColumn(row, COLUMN_NAMES.end)]),
        preinscriptionLink: String(row[findColumn(row, COLUMN_NAMES.link)] || '').trim(),
        preinscritsTotal: stringifyCell(row[findColumn(row, COLUMN_NAMES.total)]),
        preinscritsMax: stringifyCell(row[findColumn(row, COLUMN_NAMES.max)]),
        raw: row
      };
    }).filter((row) => row.dispositifCode || row.moduleCode);
  }

  function findColumn(row, possibleNames) {
    const keys = Object.keys(row || {});
    const normalizedTargets = possibleNames.map(normalizeColumnName);
    let key = keys.find((k) => normalizedTargets.includes(normalizeColumnName(k)));
    if (key) return key;

    key = keys.find((k) => {
      const normalizedKey = normalizeColumnName(k);
      return normalizedTargets.some((target) => normalizedKey.includes(target) || (normalizedKey.length >= 8 && target.includes(normalizedKey)));
    });
    return key || possibleNames[0];
  }

  function buildExcelIndex(rows) {
    const index = new Map();
    for (const row of rows || []) {
      const key = makeMappingKey(row.dispositifCode, row.moduleCode);
      if (!key) continue;
      if (!index.has(key)) index.set(key, []);
      index.get(key).push(row);
    }
    for (const [key, matches] of index.entries()) {
      if (matches.length > 1) log('warning', null, `Doublon Excel pour ${key} : ${matches.length} lignes. La première sera utilisée par défaut.`);
    }
    return index;
  }

  /***************************************************************************
   * CORRESPONDANCES ET VALIDATION
   ***************************************************************************/
  function makeMappingKey(dispositifCode, moduleCode) {
    const d = cleanCode(dispositifCode);
    if (!d) return '';
    if (CONFIG.MATCH_EXCEL_BY_DISPOSITIF_ONLY) return d;
    const m = cleanCode(moduleCode);
    return d && m ? `${d}::${m}` : '';
  }

  function enrichMappingWithExcelData(mappingItems, excelIndex) {
    return (mappingItems || []).map((item) => {
      const key = makeMappingKey(item.dispositifCode, item.moduleCode);
      const matches = excelIndex.get(key) || [];
      const excelData = selectExcelMatchForMappingItem(item, matches);
      return { ...item, excelMatches: matches, excelData, valid: false, validationMessages: [] };
    });
  }

  function selectExcelMatchForMappingItem(item, matches) {
    if (!matches.length) return null;
    const moduleMatches = item.moduleCode ? matches.filter((row) => row.moduleCode === item.moduleCode) : [];
    const candidateMatches = moduleMatches.length ? moduleMatches : matches;
    const scored = candidateMatches.map((row, index) => ({ row, index, score: scoreExcelMatch(item, row) }))
      .sort((a, b) => b.score - a.score || a.index - b.index);
    const selected = scored[0]?.row || matches[0];
    return buildExcelDataWithGroupSessions(selected, candidateMatches, item);
  }

  function buildExcelDataWithGroupSessions(selected, matches, item) {
    const sessions = matches.filter((row) => isSamePlanSessionGroup(selected, row, item))
      .sort((a, b) => compareSessionRows(a, b));
    const orderedSessions = sessions.length ? sessions : [selected];
    const firstWithLink = orderedSessions.find((row) => row.preinscriptionLink) || selected;
    const firstWithPublicationStart = orderedSessions.find((row) => row.publicationStartDate) || selected;
    const firstWithPublicationEnd = orderedSessions.find((row) => row.publicationEndDate) || selected;
    const firstWithTotal = orderedSessions.find((row) => row.preinscritsTotal !== '') || selected;
    const firstWithMax = orderedSessions.find((row) => row.preinscritsMax !== '') || selected;

    return {
      ...selected,
      preinscriptionLink: firstWithLink.preinscriptionLink || selected.preinscriptionLink,
      publicationStartDate: firstWithPublicationStart.publicationStartDate || selected.publicationStartDate,
      publicationEndDate: firstWithPublicationEnd.publicationEndDate || selected.publicationEndDate,
      preinscritsTotal: firstWithTotal.preinscritsTotal,
      preinscritsMax: firstWithMax.preinscritsMax,
      sessions: orderedSessions
    };
  }

  function isSamePlanSessionGroup(selected, row, item) {
    if (selected.groupId && row.groupId) return selected.groupId === row.groupId;
    if (item.planSessionGroupId && row.groupId) return item.planSessionGroupId === row.groupId;
    if (item.rne && row.rne) return item.rne === row.rne;
    if (selected.rne && row.rne) return selected.rne === row.rne;
    if (item.territoireCode && row.territoireCode) return item.territoireCode === row.territoireCode;
    return row === selected;
  }

  function compareSessionRows(a, b) {
    const aTime = parseExcelDate(a.sessionStartRaw)?.getTime() || Number.MAX_SAFE_INTEGER;
    const bTime = parseExcelDate(b.sessionStartRaw)?.getTime() || Number.MAX_SAFE_INTEGER;
    if (aTime !== bTime) return aTime - bTime;
    return String(a.sessionLabel || '').localeCompare(String(b.sessionLabel || ''), 'fr', { numeric: true });
  }

  function scoreExcelMatch(item, row) {
    let score = 0;
    if (item.rne && row.rne && item.rne === row.rne) score += 100;
    if (item.territoireCode && row.territoireCode && item.territoireCode === row.territoireCode) score += 30;
    if (item.planSessionGroupId && row.groupId && item.planSessionGroupId === row.groupId) score += 20;
    if (item.planSessionLabel && row.sessionLabel && item.planSessionLabel === row.sessionLabel) score += 15;
    if (item.moduleCode && row.moduleCode && item.moduleCode === row.moduleCode) score += 5;
    return score;
  }

  function analyzeCorrespondences(verbose = true) {
    state.enrichedItems = enrichMappingWithExcelData(state.mappingItems, state.excelIndex);
    state.validationErrors = [];
    let matches = 0;

    for (const item of state.enrichedItems) {
      const errors = validateItem(item);
      item.validationMessages = errors;
      item.valid = errors.length === 0;
      if (item.excelData) matches += 1;
      errors.forEach((message) => state.validationErrors.push({ item, message }));
      if (!item.excelData && item.dispositifCode && item.moduleCode) log('error', item, 'Aucune ligne Excel correspondante : parcours ignoré.');
      if (item.excelMatches?.length > 1) log('warning', item, `Plusieurs lignes Excel (${item.excelMatches.length}) pour ce dispositif/module : sélection via RNE/territoire du mapping (${item.rne || item.territoireCode || 'aucune clé'}).`);
    }

    state.lastSummary = {
      mapping: state.mappingItems.length,
      excel: state.excelRows.length,
      matches,
      errors: state.validationErrors.length
    };
    renderSummary();
    renderErrors();
    if (verbose) {
      const validCount = state.enrichedItems.filter((item) => item.valid).length;
      const msg = `Analyse terminée : ${matches} correspondance(s) via Dispositif : code, ${validCount} ligne(s) valide(s), ${state.validationErrors.length} erreur(s).`;
      setAlert(msg);
      log(validCount ? 'success' : 'warning', null, msg);
    }
  }

  function validateItem(item) {
    const errors = [];
    const data = item.excelData;
    if (!item.drupalParcoursId) errors.push('ID parcours Drupal manquant.');
    if (!item.dispositifCode) errors.push('Code dispositif manquant.');
    if (!CONFIG.MATCH_EXCEL_BY_DISPOSITIF_ONLY && !item.moduleCode) errors.push('Code module manquant.');
    if (!data) errors.push('Correspondance Excel introuvable via la colonne Excel "Dispositif : code".');
    if (data) {
      if (!data.preinscriptionLink) errors.push('Lien de préinscription manquant.');
      if (!data.publicationStartDate) errors.push('Date de début de publication manquante/invalide.');
      if (!data.publicationEndDate) errors.push('Date de fin de publication manquante/invalide.');
      if (data.preinscritsTotal === '') errors.push('Nombre total de préinscrits manquant.');
      if (data.preinscritsMax === '') errors.push('Nombre maximal de préinscrits manquant.');
    }
    return errors;
  }

  /***************************************************************************
   * DATES
   ***************************************************************************/
  function parseExcelDate(value) {
    if (!value && value !== 0) return null;
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    if (typeof value === 'number') {
      const parsed = XLSX?.SSF?.parse_date_code ? XLSX.SSF.parse_date_code(value) : null;
      if (parsed) return new Date(Date.UTC(parsed.y, parsed.m - 1, parsed.d, parsed.H || 0, parsed.M || 0, Math.floor(parsed.S || 0)));
      return new Date(Math.round((value - 25569) * 86400 * 1000));
    }
    const text = String(value).trim().replace(/\s+/g, ' ');
    if (!text) return null;
    let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T]+(\d{1,2})[:h](\d{2})(?::(\d{2}))?)?/);
    if (match) return new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3]), Number(match[4] || 0), Number(match[5] || 0), Number(match[6] || 0)));
    match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})(?:[ T]+(\d{1,2})[:h](\d{2})(?::(\d{2}))?)?/);
    if (match) return new Date(Date.UTC(Number(match[3]), Number(match[2]) - 1, Number(match[1]), Number(match[4] || 0), Number(match[5] || 0), Number(match[6] || 0)));
    const date = new Date(text);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  function toInputDate(value) {
    const date = parseExcelDate(value);
    if (!date) return '';
    const y = date.getUTCFullYear();
    const m = String(date.getUTCMonth() + 1).padStart(2, '0');
    const d = String(date.getUTCDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  function toDisplayDateTime(value) {
    if (!value && value !== 0) return '';
    if (typeof value === 'string') {
      const trimmed = value.trim().replace(/\s+/g, ' ');
      if (!trimmed) return '';
      const fr = trimmed.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})(?:[ T]+(\d{1,2})[:h](\d{2})(?::(\d{2}))?)?/);
      if (fr) return `${fr[1].padStart(2, '0')}/${fr[2].padStart(2, '0')}/${fr[3]}${fr[4] ? ` à ${fr[4].padStart(2, '0')}:${fr[5]}${fr[6] ? `:${fr[6].padStart(2, '0')}` : ''}` : ''}`;
    }
    const date = parseExcelDate(value);
    if (!date) return stringifyCell(value);
    const day = String(date.getUTCDate()).padStart(2, '0');
    const month = String(date.getUTCMonth() + 1).padStart(2, '0');
    const year = date.getUTCFullYear();
    const hours = String(date.getUTCHours()).padStart(2, '0');
    const minutes = String(date.getUTCMinutes()).padStart(2, '0');
    const seconds = String(date.getUTCSeconds()).padStart(2, '0');
    const hasTime = date.getUTCHours() || date.getUTCMinutes() || date.getUTCSeconds();
    return `${day}/${month}/${year}${hasTime ? ` à ${hours}:${minutes}${date.getUTCSeconds() ? `:${seconds}` : ''}` : ''}`;
  }

  function addDaysToInputDate(date, days) {
    const parsed = parseExcelDate(date);
    if (!parsed) return '';
    parsed.setUTCDate(parsed.getUTCDate() + Number(days || 0));
    return toInputDate(parsed);
  }

  /***************************************************************************
   * BOUCLE DE TRAITEMENT
   ***************************************************************************/
  async function runUpdateBatch() {
    if (state.running) return;
    if (state.activeWorkflow !== 'sofia') return startIntranetBatch();
    const existingBatch = readBatchState();
    if (existingBatch.active) {
      log('info', null, `Reprise du batch existant à l'étape ${existingBatch.phase}, élément ${Number(existingBatch.index || 0) + 1}.`);
      await runPersistedBatchStep();
      return;
    }
    analyzeCorrespondences(false);
    const validItems = state.enrichedItems.filter((item) => item.valid);
    if (!state.mappingItems.length || !state.excelRows.length || !validItems.length) {
      setAlert('Impossible de démarrer : importez les fichiers et conservez au moins une ligne valide.');
      log('error', null, 'Contrôles avant lancement non satisfaits.');
      return;
    }

    const summary = `Résumé avant lancement : ${validItems.length} parcours valides sur ${state.enrichedItems.length}. Correspondance Excel par colonne "Dispositif : code". Simulation=${CONFIG.DRY_RUN ? 'oui' : 'non'}, auto-save=${CONFIG.AUTO_SAVE ? 'oui' : 'non'}, validation clic=${CONFIG.SAFE_CLICK_MODE ? 'oui' : 'non'}. Continuer ?`;
    if (!window.confirm(summary)) return;

    state.updateReport = [];
    writeStoredJson(STORAGE_KEYS.updateReport, state.updateReport);
    writeBatchState({ active: true, index: 0, phase: 'goToAdmin', workflowType: 'sofia', startedAt: new Date().toISOString() });
    state.stopRequested = false;
    log('info', null, `${summary} Le batch reprendra automatiquement après chaque changement de page Drupal.`);
    await runPersistedBatchStep();
  }

  function scheduleAutoResumeBatch() {
    const batch = readBatchState();
    if (!batch.active) return;

    const workflowType = batch.workflowType || 'sofia';
    if (workflowType !== 'sofia') {
      const payload = state.intranetUpdate.payload;
      if (!payload) {
        log('warning', null, `Batch ${workflowType} abandonné : aucun payload intranet correspondant n’est chargé.`);
        clearBatchState();
        return;
      }
      if (workflowType !== state.activeWorkflow) {
        log('warning', null, `Batch ${workflowType} abandonné : le workflow actif est ${state.activeWorkflow}.`);
        clearBatchState();
        return;
      }
    }

    setStatus('Batch détecté : reprise automatique…');
    setAlert('Batch compatible restauré : reprise automatique.');
    window.setTimeout(() => runPersistedBatchStep(), 900);
  }

  async function runPersistedBatchStep() {
    if (state.running) return;
    state.running = true;
    state.stopRequested = false;
    setStatus('Traitement en cours…');

    try {
      while (!state.stopRequested) {
        const batch = readBatchState();
        const workflowType = batch.workflowType || 'sofia';
        let validItems;
        if (workflowType === 'sofia') {
          analyzeCorrespondences(false);
          validItems = state.enrichedItems.filter((entry) => entry.valid);
        } else {
          validItems = state.intranetUpdate.normalizedTargets;
        }

        if (!batch.active) break;
        if (batch.index >= validItems.length) {
          clearBatchState();
          setStatus('Terminé.');
          log('success', null, 'Batch terminé.');
          break;
        }

        const item = validItems[batch.index];
        const itemId = workflowType === 'sofia' ? item.drupalParcoursId : item.nodeId;
        const phase = batch.phase || 'goToAdmin';
        log('info', item, `Reprise batch automatique : élément ${batch.index + 1}/${validItems.length}, étape ${phase}.`);

        try {
          if (phase === 'goToAdmin') {
            writeBatchState({ ...batch, phase: 'filter' });
            await goToAdminContent();
            continue;
          }

          if (phase === 'filter') {
            // Important : le clic sur "Filtrer" peut recharger immédiatement la page.
            // On mémorise donc l'étape suivante AVANT le clic pour éviter de recliquer
            // en boucle sur Filtrer après le rechargement de /admin/content.
            writeBatchState({ ...batch, phase: 'openResult' });
            await filterByDrupalId(itemId);
            continue;
          }

          if (phase === 'openResult') {
            writeBatchState({ ...batch, phase: 'updatePage' });
            await openFirstResult(item);
            continue;
          }

          if (phase === 'updatePage') {
            if (location.pathname.startsWith(CONFIG.ADMIN_CONTENT_URL)) {
              // Après un rechargement, on peut revenir sur /admin/content avec l'étape updatePage déjà mémorisée.
              // Dans ce cas il faut d'abord recliquer sur le lien du parcours, pas chercher un lien Modifier dans l'admin.
              writeBatchState({ ...batch, phase: 'openResult' });
              continue;
            }
            const readyToEdit = await ensureEditPage();
            if (!readyToEdit) return;
            let intranetResult = null;
            if (workflowType === 'sofia') {
              if (state.activeWorkflow !== 'sofia') throw typedError('workflow_mismatch', `Refus du workflow Sofia : le workflow actif est ${state.activeWorkflow}.`);
              await updateCurrentDrupalPage(item, item.excelData);
            } else {
              if (!state.intranetUpdate.payload) throw typedError('workflow_mismatch', 'Aucun JSON intranet chargé pour ce batch.');
              if (workflowType !== state.activeWorkflow) throw typedError('workflow_mismatch', `Workflow du batch incompatible : ${workflowType} au lieu de ${state.activeWorkflow}.`);
              if (item.payloadType !== state.intranetUpdate.payload.type) throw typedError('workflow_mismatch', `Type de cible incompatible : ${item.payloadType} au lieu de ${state.intranetUpdate.payload.type}.`);
              intranetResult = await updateCurrentDrupalPageFromIntranetPayload(item);
            }
            const nextIndex = Number(batch.index || 0) + 1;
            const hasMoreItems = nextIndex < validItems.length;

            // Le clic Enregistrer peut recharger immédiatement la page : on persiste donc
            // l'état suivant AVANT le clic. S'il n'y a plus rien à traiter, le batch est
            // désactivé avant sauvegarde pour éviter toute reprise automatique après reload.
            if (hasMoreItems) writeBatchState({ ...batch, index: nextIndex, phase: 'goToAdmin' });
            else clearBatchState();

            let saved = false;
            if (workflowType === 'sofia') saved = await saveCurrentPageIfAllowed(item);
            else {
              const saveOptions = buildIntranetSaveOptions(intranetResult, workflowType);
              const saveExpected = !CONFIG.DRY_RUN && CONFIG.AUTO_SAVE && !saveOptions.forbidden && !saveOptions.conflict && saveOptions.fieldsValid !== false && saveOptions.verified !== false && saveOptions.backupGenerated;
              // Persistance avant le clic : une navigation Drupal peut détruire le contexte JavaScript immédiatement.
              finalizeIntranetReportItem(intranetResult, saveExpected);
              saved = await saveCurrentPageIfAllowed(item, saveOptions);
              if (!saveExpected || saved === false) finalizeIntranetReportItem(intranetResult, Boolean(saved));
            }
            log('success', item, 'Traitement terminé pour ce parcours.');

            if (!hasMoreItems) {
              setStatus('Terminé.');
              setAlert('Dernier parcours traité : le batch est terminé et ne reprendra pas après l’enregistrement.');
              log('success', null, 'Batch terminé : aucun autre parcours à mettre à jour.');
              break;
            }

            await sleep(CONFIG.DELAY_BETWEEN_ACTIONS);
            continue;
          }

          writeBatchState({ ...batch, phase: 'goToAdmin' });
        } catch (error) {
          if (isClickCancelledError(error)) {
            setStatus('En pause : clic annulé.');
            setAlert('Clic annulé : le batch reste positionné sur cette étape. Relancez avec Démarrer ou rechargez la page pour reprendre.');
            log('warning', item, `Batch mis en pause sans passer au parcours suivant : ${error.message}`);
            return;
          }
          log('error', item, `Erreur parcours : ${error.message}. Contexte : ${captureTextContext()}`);
          if (workflowType !== 'sofia') {
            const failed = createIntranetReportItem(item);
            failed.status = /accès refusé|connexion/i.test(error.message) ? 'access_denied' : /résultat|premier résultat/i.test(error.message) ? 'page_not_found' : 'field_not_found';
            failed.failedFields = getRequestedFields(item);
            failed.errors = [{ code: failed.status, message: error.message }];
            finalizeIntranetReportItem(failed, false);
          }
          writeBatchState({ ...batch, index: batch.index + 1, phase: 'goToAdmin' });
          await sleep(CONFIG.DELAY_BETWEEN_ACTIONS);
        }
      }
    } finally {
      state.running = false;
      if (state.stopRequested) {
        clearBatchState();
        setStatus('Arrêté.');
        log('warning', null, 'Batch interrompu proprement.');
      }
    }
  }

  function readBatchState() {
    const batch = readStoredJson(STORAGE_KEYS.batchState, { active: false, index: 0, phase: 'goToAdmin', workflowType: 'sofia' }) || { active: false, index: 0, phase: 'goToAdmin' };
    return { ...batch, workflowType: batch.workflowType || 'sofia' };
  }

  function writeBatchState(batchState) {
    writeStoredJson(STORAGE_KEYS.batchState, batchState);
  }

  function clearBatchState() {
    writeBatchState({ active: false, index: 0, phase: 'goToAdmin', workflowType: 'sofia' });
  }

  function isClickCancelledError(error) {
    return /Clic annulé par l’utilisateur/i.test(String(error?.message || error || ''));
  }

  async function processOneMappingItem(item) {
    log('info', item, `Démarrage parcours Drupal ${item.drupalParcoursId}.`);
    await goToAdminContent();
    await filterByDrupalId(item.drupalParcoursId);
    await openFirstResult(item);
    await updateCurrentDrupalPage(item, item.excelData);
    await saveCurrentPageIfAllowed(item);
  }

  async function goToAdminContent() {
    if (!location.pathname.startsWith(CONFIG.ADMIN_CONTENT_URL)) {
      const menu = await waitForAnySelector(SELECTORS.adminMenu, 3000).catch(() => null);
      if (menu) await clickElement(menu, 'Ouvrir le menu d’administration');
      const link = await waitForAnySelector(SELECTORS.contentLink, 3000).catch(() => null);
      if (link) {
        await clickElement(link, 'Ouvrir la page Contenu');
        await waitForLocation(CONFIG.ADMIN_CONTENT_URL);
      } else {
        location.href = CONFIG.ADMIN_CONTENT_URL;
        await waitForLocation(CONFIG.ADMIN_CONTENT_URL);
      }
    }
    await waitForAnySelector(SELECTORS.nidFilter, CONFIG.MAX_WAIT_TIME);
  }

  async function filterByDrupalId(drupalParcoursId) {
    const input = await waitForAnySelector(SELECTORS.nidFilter, CONFIG.MAX_WAIT_TIME);
    setInputValue(input, drupalParcoursId);
    const button = await waitForAnySelector(SELECTORS.filterButton, CONFIG.MAX_WAIT_TIME);
    await clickElement(button, `Filtrer par ID Drupal ${drupalParcoursId}`);
    await sleep(1000);
  }

  function getFirstResultTitleLink() {
    // Ciblage explicite demandé : cliquer sur le premier <a href="/formation/...">
    // contenu dans la cellule titre Drupal :
    // <td headers="view-title-table-column" class="views-field views-field-title">...
    const exactFirstRowLink = document.querySelector('table.views-table tbody tr:first-child td[headers="view-title-table-column"].views-field-title > a[href^="/formation/"]')
      || document.querySelector('table.views-table tbody tr:first-child td[headers="view-title-table-column"].views-field-title > a[href]');
    if (exactFirstRowLink) return exactFirstRowLink;

    const titleCells = Array.from(document.querySelectorAll('td[headers="view-title-table-column"].views-field-title, td[headers="view-title-table-column"]'));
    for (const cell of titleCells) {
      const firstLink = cell.querySelector('a[href^="/formation/"]') || cell.querySelector('a[href]');
      if (firstLink) return firstLink;
    }

    return findElementBySelectors(SELECTORS.resultTitleLink);
  }

  async function openFirstResult(item) {
    const link = await waitForFunction(getFirstResultTitleLink, CONFIG.MAX_WAIT_TIME, 'Lien du premier résultat dans la colonne titre introuvable.').catch(() => null);
    if (!link && item?.pageUrl) {
      log('warning', item, `Résultat admin introuvable : navigation de secours vers ${item.pageUrl}.`);
      location.href = item.pageUrl;
      await sleep(1000);
      return;
    }
    if (!link) throw new Error('Lien du premier résultat dans la colonne titre introuvable.');

    const row = link.closest('tr');
    const expectedId = item.drupalParcoursId || item.nodeId || '';
    if (row && expectedId && !row.textContent.includes(expectedId)) log('warning', item, 'Le premier résultat ne contient pas clairement l’ID filtré dans la ligne.');
    log('info', item, `Ouverture du résultat : ${normalizeText(link.textContent).slice(0, 120)}`);
    await clickElement(link, 'Ouvrir le premier résultat Drupal');
    await sleep(1000);

    const editLink = findElementBySelectors(SELECTORS.editLinks);
    if (editLink && !/\/edit(?:$|[?#])/.test(location.pathname + location.search)) {
      await clickElement(editLink, 'Ouvrir la page de modification');
      await sleep(1000);
    }
  }

  async function ensureEditPage() {
    if (/\/edit(?:$|[?#])/.test(location.pathname + location.search)) return true;
    if (findElementBySelectors(SELECTORS.startDate)) return true;
    if (location.pathname.startsWith(CONFIG.ADMIN_CONTENT_URL)) return false;

    const editLink = await waitForFunction(() => findElementBySelectors(SELECTORS.editLinks), 4000, 'Lien Modifier introuvable sur la page formation.').catch(() => null);
    if (!editLink) throw new Error('Page de modification introuvable : aucun lien Modifier ni champ date détecté.');
    await clickElement(editLink, 'Ouvrir la page de modification');
    await sleep(1000);
    return /\/edit(?:$|[?#])/.test(location.pathname + location.search) || Boolean(findElementBySelectors(SELECTORS.startDate));
  }

  async function clickFormationBodyEditBeforeRichText(item) {
    const editButton = await waitForAnySelector(SELECTORS.formationBodyEdit, 5000).catch(() => null);
    if (!editButton) {
      log('warning', item, 'Bouton Modifier du paragraphe formation body introuvable avant modification du texte riche.');
      return;
    }

    await clickElement(editButton, 'Cliquer sur Modifier du paragraphe formation body avant modification du texte riche');
    await sleep(1500);
    log('success', item, 'Bouton Modifier du paragraphe formation body cliqué avant modification du texte riche.');
  }

  async function updateCurrentDrupalPage(item, data) {
    const startDate = data.publicationStartDate;
    const endDate = data.publicationEndDate;
    const unpublishDate = addDaysToInputDate(endDate, 1);

    setInputValue(await waitForAnySelector(SELECTORS.startDate, CONFIG.MAX_WAIT_TIME), startDate);
    setInputValue(await waitForAnySelector(SELECTORS.endDate, CONFIG.MAX_WAIT_TIME), endDate);
    setInputValue(await waitForAnySelector(SELECTORS.unpublishDate, CONFIG.MAX_WAIT_TIME), unpublishDate);
    log('success', item, `Dates mises à jour : début=${startDate}, fin=${endDate}, dépublication=${unpublishDate}.`);

    await applyPublicsAndThematicsFromSource(item, data);

    await clickFormationBodyEditBeforeRichText(item);

    const formatSelect = await waitForAnySelector(SELECTORS.textFormat, CONFIG.MAX_WAIT_TIME).catch(() => null);
    if (formatSelect) {
      setSelectValue(formatSelect, CONFIG.ADVANCED_FORMAT_VALUE);
      log('success', item, 'Formatage avancé sélectionné.');
      await sleep(1200);
    } else {
      log('warning', item, 'Select de format texte introuvable : poursuite du traitement.');
    }

    await openCkEditorSource();
    const currentHtml = await getCurrentHtmlFromEditor();
    const newHtml = updateFormationHtml(currentHtml, data);
    await setHtmlInEditor(newHtml);
    log('success', item, 'HTML éditorial mis à jour sans insertion de codes Sofia-FMO/GAIA.');

    const existingCallout = findExistingCalloutParagraph();
    if (existingCallout) {
      log('success', item, 'Mise en avant déjà présente : aucune nouvelle mise en avant ajoutée, contenu conservé.');
    } else {
      log('info', item, 'Aucune mise en avant existante détectée : création du paragraphe avant tout enregistrement.');
      await createCallout(data);
      // addCalloutParagraph() a déjà renseigné le CTA, le format avancé, la source et la couleur.
      // Ne pas rappeler setCalloutSourceToBr() ici : un second clic sur Source peut refermer l'éditeur
      // et empêcher le clic final sur Enregistrer.
      if (!findExistingCalloutParagraph()) {
        throw new Error('La mise en avant n’a pas été détectée après création : enregistrement annulé pour éviter une page incomplète.');
      }
      log('success', item, 'Mise en avant de préinscription créée et vérifiée.');
    }

    recordUpdateReport(item, data, { startDate, endDate, unpublishDate });
  }

  async function saveCurrentPageIfAllowed(item, options = {}) {
    if (options.forbidden || options.conflict || options.fieldsValid === false || options.verified === false || (options.workflowType && options.workflowType !== 'sofia' && !options.backupGenerated)) {
      const msg = options.reason || 'Enregistrement intranet interdit par les contrôles de sécurité.';
      setAlert(msg); log('warning', item, msg); return false;
    }
    if (CONFIG.DRY_RUN || !CONFIG.AUTO_SAVE) {
      const msg = 'Page modifiée en mode simulation ou sans sauvegarde automatique. Vérifiez puis enregistrez manuellement.';
      setAlert(msg);
      log('warning', item, msg);
      return false;
    }

    // Sécurité de séquence : le bouton Drupal "Enregistrer" ne doit être cliqué
    // qu'après toutes les modifications (dates, texte riche, mise en avant, couleur).
    await waitForAjax();
    await sleep(CONFIG.DELAY_BETWEEN_ACTIONS);
    log('info', item, 'Toutes les modifications sont terminées : clic final sur Enregistrer autorisé.');

    const save = await waitForFunction(findFinalSaveButton, CONFIG.MAX_WAIT_TIME, 'Bouton final Enregistrer introuvable.');
    await clickElement(save, 'Enregistrer la page Drupal en tout dernier');
    await sleep(1500);
    return true;
  }

  function findFinalSaveButton() {
    return findElementBySelectors(SELECTORS.save);
  }

  async function goNext() {
    await sleep(CONFIG.DELAY_BETWEEN_ACTIONS);
  }

  /***************************************************************************
   * CKEDITOR 5 / SOURCE
   ***************************************************************************/
  async function openCkEditorSource() {
    const existing = getCkEditorSourceElement();
    if (existing) return existing;
    const sourceButton = await waitForAnySelector(SELECTORS.ckSourceButton, CONFIG.MAX_WAIT_TIME);
    await clickElement(sourceButton, 'Ouvrir la source CKEditor');
    return waitForFunction(getCkEditorSourceElement, CONFIG.MAX_WAIT_TIME, 'Source CKEditor introuvable.');
  }

  function getCkEditorSourceElement() {
    return SELECTORS.ckSourceArea.map((selector) => Array.from(document.querySelectorAll(selector)).find(isVisible)).find(Boolean) || null;
  }

  async function getCurrentHtmlFromEditor() {
    const source = getCkEditorSourceElement() || await openCkEditorSource();
    if ('value' in source) return source.value || '';
    return source.textContent || source.innerHTML || '';
  }

  async function setHtmlInEditor(html) {
    const source = getCkEditorSourceElement() || await openCkEditorSource();
    if ('value' in source) source.value = html;
    else source.textContent = html;
    dispatchAllEvents(source);
  }

  /***************************************************************************
   * MISE À JOUR HTML ÉDITORIAL
   ***************************************************************************/
  function updateFormationHtml(html, data) {
    let output = String(html || '');
    output = removeExistingLieuSection(output);
    output = removeExistingPlacesLine(output);
    output = removeExistingSessionParagraphs(output);
    output = upsertMiseEnPlaceFormationDetails(output, data);
    return cleanHtml(output);
  }

  function upsertMiseEnPlaceFormationDetails(html, data) {
    const sessions = data.sessions?.length ? data.sessions : [data];
    const sessionParagraphs = buildSessionParagraphs(sessions);
    const placesItem = buildPlacesListItem(data.preinscritsTotal, data.preinscritsMax);
    const lieuxItem = buildLieuxDatesListItem();
    const generatedListItems = `${placesItem}\n    ${lieuxItem}`;
    const miseEnPlaceRegex = /(<h3[^>]*>\s*Mise en place de la formation(?:&nbsp;|\s)*<\/h3>\s*<ul[^>]*>)([\s\S]*?)(<\/ul>)/i;

    if (miseEnPlaceRegex.test(html)) {
      return html.replace(miseEnPlaceRegex, (match, beforeList, listContent, afterList) => {
        const cleanedList = removeGeneratedMiseEnPlaceItems(listContent).trimEnd();
        return `${beforeList}${cleanedList}\n    ${generatedListItems}\n${afterList}\n${sessionParagraphs}`;
      });
    }

    const section = `<h3>Mise en place de la formation&nbsp;</h3>\n<ul>\n    ${generatedListItems}\n</ul>\n${sessionParagraphs}`;
    const contactRegex = /<h3[^>]*>\s*Contact de la formation\s*<\/h3>/i;
    const preinscriptionRegex = /<h3[^>]*>\s*Préinscription\s*<\/h3>|<h3[^>]*>\s*Preinscription\s*<\/h3>/i;
    if (contactRegex.test(html)) return html.replace(contactRegex, `${section}\n$&`);
    if (preinscriptionRegex.test(html)) return html.replace(preinscriptionRegex, `${section}\n$&`);
    return `${html.trim()}\n${section}`;
  }

  function buildPlacesListItem(total, max) {
    const remaining = calculateRemainingPlaces(total, max);
    return `<li>
        <strong>Nombre de places restantes :</strong> ${escapeHtml(remaining)}
    </li>`;
  }

  function calculateRemainingPlaces(total, max) {
    const registered = parseNumberCell(total);
    const capacity = parseNumberCell(max);
    if (registered == null || capacity == null) return '';
    return String(Math.max(capacity - registered, 0));
  }

  function parseNumberCell(value) {
    if (value == null || value === '') return null;
    const normalized = String(value).replace(/\s+/g, '').replace(',', '.');
    const match = normalized.match(/-?\d+(?:\.\d+)?/);
    if (!match) return null;
    const parsed = Number(match[0]);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function buildLieuxDatesListItem() {
    return `<li class="ck-list-marker-bold">\n        <strong>Lieux et dates de la formation :&nbsp;</strong>\n    </li>`;
  }

  function buildSessionParagraphs(sessions) {
    return (sessions || [])
      .map((session, index) => formatSessionParagraph(session, index + 1))
      .filter(Boolean)
      .join('\n');
  }

  function formatSessionParagraph(session, number) {
    const modalite = escapeHtml(session.modalite || '');
    const start = escapeHtml(session.sessionStartDateTime || toDisplayDateTime(session.sessionStartRaw));
    const end = escapeHtml(session.sessionEndDateTime || toDisplayDateTime(session.sessionEndRaw));
    const lieu = escapeHtml(session.lieu || '');
    return `<p>\n    Session ${number}<br>\n    Modalité : ${modalite}<br>\n    Date et heure : Du ${start} au ${end}<br>\n    Lieu : ${lieu}\n</p>`;
  }

  function removeGeneratedMiseEnPlaceItems(listHtml) {
    return String(listHtml || '')
      .replace(/<li[^>]*>[\s\S]*?<strong>\s*(?:Nombre de place restante|Nombre de places restantes|Nombre de personne inscrite par rapport au nombre maximal de stagiaires)\s*:?[\s\S]*?<\/li>\s*/gi, '')
      .replace(/<li[^>]*>[\s\S]*?<strong>\s*(?:Lieux? et dates de la formation|Lieu et dates de la formation)\s*:?(?:&nbsp;)?[\s\S]*?<\/li>\s*/gi, '');
  }

  function removeExistingLieuSection(html) {
    let output = String(html || '');
    output = output.replace(/<h3[^>]*>\s*Lieu et dates de la formation\s*<\/h3>\s*(?:<p[^>]*>.*?<\/p>\s*)+/gis, '');
    output = output.replace(/<h3[^>]*>\s*Lieux et dates de la formation\s*<\/h3>\s*(?:<p[^>]*>.*?<\/p>\s*)+/gis, '');
    output = output.replace(/<h3[^>]*>\s*Lieu de la formation\s*<\/h3>\s*<p[^>]*>.*?<\/p>\s*/gis, '');
    output = output.replace(/<h3[^>]*>\s*Lieu\s*de\s*la\s*formation\s*<\/h3>\s*/gis, '');
    return output;
  }

  function removeExistingPlacesLine(html) {
    return String(html || '')
      .replace(/<p[^>]*>\s*Nombre de place restante\s*:\s*.*?<\/p>\s*/gis, '')
      .replace(/<p[^>]*>\s*Nombre de places restantes\s*:\s*.*?<\/p>\s*/gis, '')
      .replace(/<p[^>]*>\s*Nombre de personne inscrite par rapport au nombre maximal de stagiaires\s*:\s*.*?<\/p>\s*/gis, '');
  }

  function removeExistingSessionParagraphs(html) {
    return String(html || '').replace(/<p[^>]*>\s*Session\s+\d+\s*<br\s*\/?>(?:.|\n)*?<\/p>\s*/gis, '');
  }

  function cleanHtml(html) {
    return String(html || '').replace(/\n{3,}/g, '\n\n').trim();
  }

  /***************************************************************************
   * PARAGRAPHE CALLOUT
   ***************************************************************************/
  async function addCalloutParagraph(data) {
    // Flux demandé : après modification du texte riche principal, ouvrir l'ajout de paragraphe,
    // choisir "Mise en avant" via l'overlay, renseigner CTA, format avancé, source <h2>Préinscriptions</h2>, couleur.
    const addButton = await waitForAnySelector(SELECTORS.addParagraph, CONFIG.MAX_WAIT_TIME);
    await clickElement(addButton, 'Ajouter Paragraphe');
    await waitForAjax();

    const miseEnAvantChoice = await waitForFunction(findMiseEnAvantChoiceElement, 12000, 'Choix Mise en avant introuvable.');
    await clickElement(miseEnAvantChoice, 'Choisir Mise en avant');
    await waitForAjax();

    let titleReady = await waitForAnySelector(SELECTORS.calloutTitle, 3000).catch(() => null);
    if (!titleReady) {
      const iconFallback = findCalloutIcon();
      if (iconFallback && iconFallback !== miseEnAvantChoice) {
        await clickElement(iconFallback, 'Choisir Mise en avant via icône callout');
        await waitForAjax();
        titleReady = await waitForAnySelector(SELECTORS.calloutTitle, 3000).catch(() => null);
      }
    }

    if (!titleReady) await waitForAnySelector(SELECTORS.calloutTitle, 12000);
    fillCalloutFields(document, data);
    await setCalloutContentFormatAdvanced(document);
    await setCalloutSourceToBr(document);
    await setCalloutColor(document);
  }

  function findMiseEnAvantChoiceElement() {
    // Le bon choix de paragraphe est le bouton "Mise en avant" ajouté par Drupal.
    // Si Drupal ne déclenche pas le bouton parent, on accepte aussi l'icône callout fournie.
    return findCalloutButton() || findCalloutIcon();
  }

  function findCalloutButton() {
    const selectorCandidates = SELECTORS.calloutButton
      .flatMap((selector) => Array.from(document.querySelectorAll(selector)))
      .filter((element, index, list) => list.indexOf(element) === index)
      .filter((element) => isClickableVisible(element) && isCalloutButton(element));

    const exactDrupalButton = selectorCandidates.find((element) => element.matches('button[name="field_formation_body_callout_add_more"], button[data-drupal-selector^="field-formation-body-callout-add-more"]'));
    if (exactDrupalButton) return exactDrupalButton;
    if (selectorCandidates.length) return selectorCandidates[0];

    const labeledButton = Array.from(document.querySelectorAll('button, input[type="submit"], a, [role="button"]'))
      .find((el) => isClickableVisible(el) && isCalloutButton(el));
    if (labeledButton) return labeledButton;

    const icon = findCalloutIcon();
    const iconButton = icon?.closest('button, input[type="submit"], a, [role="button"]');
    if (iconButton && isClickableVisible(iconButton) && isCalloutButton(iconButton)) return iconButton;

    return icon;
  }

  function findCalloutIcon() {
    const icon = SELECTORS.calloutIcon
      .flatMap((selector) => Array.from(document.querySelectorAll(selector)))
      .find((element) => isClickableVisible(element));
    if (icon) return icon;

    const overlay = SELECTORS.calloutOverlay
      .flatMap((selector) => Array.from(document.querySelectorAll(selector)))
      .find((element) => isClickableVisible(element));
    const overlayIcon = overlay?.closest('.paragraphs-button--icon');
    if (overlayIcon && isClickableVisible(overlayIcon) && /callout-default-icon/i.test(overlayIcon.getAttribute('style') || '')) return overlayIcon;

    return null;
  }

  function isCalloutButton(element) {
    const descriptor = normalizeText(`${element.textContent || ''} ${element.value || ''} ${element.name || ''} ${element.title || ''} ${element.getAttribute('aria-label') || ''} ${element.getAttribute('data-drupal-selector') || ''} ${element.className || ''}`);
    const isAddMoreButton = /field[_-]formation[_-]body[_-]callout[_-]add[_-]more|paragraphs[_-]button[_-]*add[_-]more/i.test(descriptor);
    return descriptor.includes('mise en avant') && (descriptor.includes('callout') || isAddMoreButton || descriptor.includes('paragraphs-button--add-more'));
  }

  function findExistingCalloutParagraph() {
    // Important : ne jamais considérer le bouton "Ajouter Paragraphe" ou la carte de choix
    // "Mise en avant" comme un paragraphe déjà créé. On accepte uniquement :
    // 1) un vrai champ de sous-formulaire callout ouvert ;
    // 2) une ligne Drupal déjà créée de type paragraph-type--callout (même repliée).
    const existingCalloutRow = Array.from(document.querySelectorAll('tr.paragraph-type--callout, tr[class*="paragraph-type--callout"]'))
      .find((row) => {
        if (isParagraphTypeChooser(row)) return false;
        const hasFormationBodyWrapper = Boolean(row.querySelector('[data-drupal-selector^="edit-field-formation-body-"]'));
        const hasCalloutType = Boolean(row.querySelector('img.paragraph-type-icon[title="Mise en avant"], img[src*="callout-default-icon"], .paragraph-type-label'));
        const hasEditAction = Boolean(row.querySelector('input[name^="field_formation_body_"][name$="_edit"], [data-drupal-selector^="field-formation-body-"][data-drupal-selector*="-edit"]'));
        const label = normalizeText(row.querySelector('.paragraph-type-label')?.textContent || row.textContent || '');
        return hasFormationBodyWrapper && hasCalloutType && hasEditAction && (label.includes('mise en avant') || label.includes('callout'));
      });
    if (existingCalloutRow) return existingCalloutRow;

    const fieldSelectors = [
      ...SELECTORS.calloutColor,
      ...SELECTORS.calloutUri,
      ...SELECTORS.calloutTitle,
      ...SELECTORS.calloutContentFormat
    ];

    for (const selector of fieldSelectors) {
      const elements = Array.from(document.querySelectorAll(selector));
      for (const element of elements) {
        const signature = `${element.name || ''} ${element.getAttribute('data-drupal-selector') || ''}`;
        if (!/field[_-]formation[_-]body|edit-field-formation-body/i.test(signature)) continue;
        if (!/field[_-]callout|field-callout/i.test(signature)) continue;

        const container = element.closest('.paragraphs-subform, .paragraph-type, .field--type-entity-reference-revisions, fieldset, details, .form-wrapper') || element.parentElement;
        if (container && !isParagraphTypeChooser(container)) return container;
      }
    }

    return null;
  }

  function isParagraphTypeChooser(element) {
    const descriptor = `${element.textContent || ''} ${element.getAttribute('data-drupal-selector') || ''} ${element.className || ''}`;
    return /paragraph-type-add|paragraphs-add-dialog|paragraphs-button|ui-dialog|choix de paragraphe/i.test(descriptor);
  }

  async function createCallout(data) {
    await addCalloutParagraph(data);
  }

  async function ensureCalloutContentAndColor(item) {
    await setCalloutContentFormatAdvanced(document);
    await setCalloutSourceToBr(document);
    await setCalloutColor(document);
    log('success', item, 'Contenu source <h2>Préinscriptions</h2> de mise en avant et couleur Bleu cumulus vérifiés avant sauvegarde.');
  }


  async function setCalloutContentFormatAdvanced(scope = document) {
    const format = await waitForFunction(
      () => findElementBySelectors(SELECTORS.calloutContentFormat, scope) || findElementBySelectors(SELECTORS.calloutContentFormat, document),
      10000,
      'Select format du contenu de mise en avant introuvable.'
    );
    setSelectValue(format, CONFIG.ADVANCED_FORMAT_VALUE, 'Formatage avancé');
    await waitForAjax(5000);
  }

  async function setCalloutSourceToBr(scope = document) {
    let sourceArea = findCalloutSourceTextarea(scope);
    if (!sourceArea) {
      const sourceButton = await waitForFunction(
        () => findCalloutSourceButton(scope),
        10000,
        'Bouton Source du contenu de mise en avant introuvable.'
      );
      await clickElement(sourceButton, 'Ouvrir la source de la mise en avant');

      sourceArea = await waitForFunction(
        () => findCalloutSourceTextarea(scope),
        10000,
        'Zone source de la mise en avant introuvable.'
      );
    }

    setInputValue(sourceArea, CONFIG.CALLOUT_BODY_TEXT);
  }

  function findCalloutSourceButton(scope = document) {
    const candidates = Array.from((scope || document).querySelectorAll(SELECTORS.calloutSourceButton.join(',')))
      .map((element) => (element.matches('span.ck-button__label') ? element.closest('button') || element : element))
      .filter((element, index, list) => element && list.indexOf(element) === index && isVisible(element))
      .filter((element) => normalizeText(`${element.textContent || ''} ${element.getAttribute('aria-label') || ''} ${element.getAttribute('data-cke-tooltip-text') || ''}`).includes('source'));

    if (!candidates.length && scope !== document) return findCalloutSourceButton(document);

    const calloutScoped = candidates.filter((element) => /field[_-]callout|mise en avant|callout/i.test(`${element.closest('.paragraphs-subform, fieldset, details, .form-wrapper, .ck-editor')?.textContent || ''} ${element.closest('.paragraphs-subform, fieldset, details, .form-wrapper, .ck-editor')?.getAttribute('data-drupal-selector') || ''}`));
    const list = calloutScoped.length ? calloutScoped : candidates;
    return list[list.length - 1] || null;
  }

  function findCalloutSourceTextarea(scope = document) {
    const areas = Array.from((scope || document).querySelectorAll(SELECTORS.calloutSourceTextarea.join(','))).filter(isVisible);
    if (!areas.length && scope !== document) return findCalloutSourceTextarea(document);

    const calloutScoped = areas.filter((area) => /field[_-]callout|mise en avant|callout/i.test(`${area.closest('.paragraphs-subform, fieldset, details, .form-wrapper, .ck-editor')?.textContent || ''} ${area.closest('.paragraphs-subform, fieldset, details, .form-wrapper, .ck-editor')?.getAttribute('data-drupal-selector') || ''}`));
    const list = calloutScoped.length ? calloutScoped : areas;
    return list[list.length - 1] || null;
  }

  async function setCalloutRichTextBody(scope = document) {
    const editor = await waitForFunction(() => findCalloutRichTextEditor(scope), 8000, 'Éditeur riche de la mise en avant introuvable.').catch(() => null);
    if (!editor) throw new Error('Éditeur riche de la mise en avant introuvable.');
    setCkEditorEditableHtml(editor, `<p>${escapeHtml(CONFIG.CALLOUT_BODY_TEXT)}</p>`);
  }

  function findCalloutRichTextEditor(scope = document) {
    const editorSelector = `.ck.ck-content.ck-editor__editable_inline[contenteditable="true"], .ck-editor__editable_inline[contenteditable="true"], div[role="textbox"][contenteditable="true"][aria-label*="Éditeur de texte enrichi"], div[role="textbox"][contenteditable="true"][aria-label*="Zone d'édition"]`;
    const editors = Array.from(scope.querySelectorAll(editorSelector)).filter(isVisible);
    if (!editors.length && scope !== document) return findCalloutRichTextEditor(document);

    const calloutScopedEditor = editors.find((editor) => /callout|mise en avant|field_callout|field-callout/i.test(`${editor.closest('.paragraphs-subform, fieldset, details, .form-wrapper')?.textContent || ''} ${editor.closest('.paragraphs-subform, fieldset, details, .form-wrapper')?.getAttribute('data-drupal-selector') || ''}`));
    return editors.find((editor) => editor.classList.contains('ck-focused'))
      || calloutScopedEditor
      || editors.find((editor) => normalizeText(editor.textContent) === '' || editor.querySelector('[data-cke-filler]'))
      || editors[editors.length - 1]
      || null;
  }

  function setCkEditorEditableHtml(editor, html) {
    editor.scrollIntoView({ block: 'center', inline: 'center' });
    editor.focus();
    editor.innerHTML = html;
    editor.dispatchEvent(new InputEvent('beforeinput', { bubbles: true, cancelable: true, inputType: 'insertText', data: CONFIG.CALLOUT_BODY_TEXT }));
    editor.dispatchEvent(new InputEvent('input', { bubbles: true, cancelable: true, inputType: 'insertText', data: CONFIG.CALLOUT_BODY_TEXT }));
    editor.dispatchEvent(new Event('change', { bubbles: true, cancelable: true }));
    editor.dispatchEvent(new Event('blur', { bubbles: true, cancelable: true }));

    const editable = editor.ckeditorInstance || editor.closest('.ck-editor')?.ckeditorInstance;
    if (editable?.setData) editable.setData(html);
  }

  async function setCalloutColor(scope = document) {
    const color = await waitForFunction(() => findElementBySelectors(SELECTORS.calloutColor, scope) || findElementBySelectors(SELECTORS.calloutColor, document), 8000, 'Select couleur Mise en avant introuvable.');
    setSelectValue(color, CONFIG.CALLOUT_COLOR_VALUE, 'Bleu cumulus');
  }

  function fillCalloutFields(scope, data) {
    const uri = findElementBySelectors(SELECTORS.calloutUri, scope);
    const title = findElementBySelectors(SELECTORS.calloutTitle, scope);
    if (!uri) throw new Error('Champ lien callout introuvable.');
    if (!title) throw new Error('Champ titre callout introuvable.');
    setInputValue(uri, data.preinscriptionLink);
    setInputValue(title, CONFIG.PREINSCRIPTION_BUTTON_TEXT);
  }

  /***************************************************************************
   * MISES À JOUR INTRANET (CONTENU ET TAXONOMIES)
   ***************************************************************************/
  const INTRANET_PAYLOAD_TYPES = ['update_drupal_pages', 'bulk_drupal_theme_public_update'];
  const INTRANET_FIELDS = ['title', 'objectif', 'contenu', 'accessibilite', 'miseEnPlace', 'effectif', 'dureeTotale', 'presentielH', 'distancielH', 'publics', 'prerequis', 'contact', 'theme', 'metiersPublics', 'departements', 'unpublishDate', 'preinscriptionLinks'];
  const EDITORIAL_FIELDS = ['objectif', 'contenu', 'accessibilite', 'miseEnPlace', 'effectif', 'dureeTotale', 'presentielH', 'distancielH', 'publics', 'prerequis', 'contact'];
  const TAXONOMY_FIELDS = ['theme', 'metiersPublics', 'departements'];
  const CONTENT_OPERATIONS = ['add', 'replace', 'delete'];
  const THEME_PUBLIC_OPERATIONS = ['add', 'remove', 'replace', 'clear'];
  const REPORT_STATUSES = ['success', 'partial_success', 'already_up_to_date', 'conflict', 'page_not_found', 'field_not_found', 'save_failed', 'access_denied', 'workflow_mismatch', 'unauthorized_field_access', 'unauthorized_field_change', 'skipped', 'cancelled'];

  async function handleIntranetUpdateFile(file) {
    try {
      const payload = JSON.parse(await file.text());
      const validation = validateIntranetPayload(payload);
      state.intranetUpdate = {
        ...state.intranetUpdate,
        payload,
        payloadType: payload?.type || '',
        payloadFileName: file.name,
        normalizedTargets: validation.targets,
        validationErrors: validation.errors,
        validationWarnings: validation.warnings,
        preview: buildIntranetPreview(payload, validation.targets),
        reportItems: [], startedAt: '', finishedAt: '', backupItems: [], backupGenerated: false
      };
      if (validation.errors.length) {
        state.activeWorkflow = 'sofia';
        setAlert(`Payload intranet invalide : ${validation.errors.length} erreur(s). Aucune modification ne sera possible.`);
        validation.errors.forEach((message) => log('error', null, message));
      } else {
        const previousBatch = readBatchState();
        if (previousBatch.active) log('warning', null, `Batch précédent annulé lors de l’import du nouveau JSON : ${previousBatch.workflowType || 'sofia'}.`);
        clearBatchState();
        state.stopRequested = false;
        state.running = false;
        state.activeWorkflow = payload.type === 'update_drupal_pages' ? 'intranet_content' : 'intranet_taxonomy';
        writeStoredJson(STORAGE_KEYS.intranetPayload, payload);
        writeStoredValue(STORAGE_KEYS.intranetPayloadFileName, file.name);
        writeStoredJson(STORAGE_KEYS.intranetReport, buildIntranetReport());
        const safe = payload.settings?.safeMode === true ? ' Enregistrement interdit par le mode sécurisé du fichier importé.' : '';
        setAlert(`Payload ${payload.type} importé et validé. L’import ne déclenche aucune action Drupal.${safe}`);
        log('success', null, `${validation.targets.length} cible(s) intranet importée(s), sans démarrage automatique.${safe}`);
      }
      renderWorkflowMode();
      renderIntranetPreview();
    } catch (error) {
      state.intranetUpdate.validationErrors = [`JSON invalide : ${error.message}`];
      state.activeWorkflow = 'sofia';
      setAlert(`Import intranet refusé : JSON invalide (${error.message}).`);
      log('error', null, `Import intranet : ${error.message}`);
      renderWorkflowMode();
    }
  }

  function validateIntranetPayload(payload) {
    const errors = [], warnings = [];
    if (!payload || typeof payload !== 'object' || Array.isArray(payload)) return { errors: ['Payload vide ou non objet.'], warnings, targets: [] };
    if (!payload.type) errors.push('Propriété type absente.');
    else if (!INTRANET_PAYLOAD_TYPES.includes(payload.type)) errors.push(`Type inconnu : ${payload.type}.`);
    if (!String(payload.batchId || '').trim()) errors.push('Batch ID absent.');
    if (payload.type === 'update_drupal_pages') {
      if (payload.version !== '3.0') errors.push('Version update_drupal_pages non prise en charge.');
      if (!Array.isArray(payload.items)) errors.push('items[] est obligatoire.');
    }
    if (payload.type === 'bulk_drupal_theme_public_update') {
      if (payload.version !== '1.0') errors.push('Version bulk_drupal_theme_public_update non prise en charge.');
      if (!Array.isArray(payload.targets)) errors.push('targets[] est obligatoire.');
    }
    const targets = normalizeIntranetTargets(payload);
    if (!targets.length) errors.push('Aucune cible exploitable dans le payload.');
    const seen = new Map();
    targets.forEach((target, index) => {
      if (!String(target.nodeId || '').trim()) errors.push(`Cible ${index + 1} : Node ID absent.`);
      const requested = target.changes;
      if (!requested || typeof requested !== 'object' || Array.isArray(requested) || !Object.keys(requested).length) errors.push(`Cible ${target.nodeId || index + 1} vide : changes/expectedAfter absent.`);
      Object.entries(requested || {}).forEach(([field, request]) => {
        if (!INTRANET_FIELDS.includes(field)) errors.push(`Cible ${target.nodeId}: champ inconnu ${field}.`);
        if (payload.type === 'update_drupal_pages') {
          if (!request || typeof request !== 'object') errors.push(`Cible ${target.nodeId}, ${field}: demande invalide.`);
          else {
            if (!CONTENT_OPERATIONS.includes(request.operation)) errors.push(`Cible ${target.nodeId}, ${field}: opération inconnue ${request.operation}.`);
            if (!Object.prototype.hasOwnProperty.call(request, 'value')) errors.push(`Cible ${target.nodeId}, ${field}: value est obligatoire pour update_drupal_pages.`);
            if (Object.prototype.hasOwnProperty.call(request, 'values')) errors.push(`Cible ${target.nodeId}, ${field}: values est interdit pour update_drupal_pages.`);
          }
          if (field === 'title' && request?.operation === 'delete') errors.push(`Cible ${target.nodeId}, title: invalid_operation (le titre ne peut pas être vidé).`);
        }
        if (payload.type === 'bulk_drupal_theme_public_update') {
          if (!['theme', 'publics'].includes(field)) errors.push(`Cible ${target.nodeId}: le format thématiques/publics ne reconnaît que theme et publics.`);
          if (!request || typeof request !== 'object' || !Array.isArray(request.values)) errors.push(`Cible ${target.nodeId}, ${field}: values doit être un tableau.`);
          else {
            if (!THEME_PUBLIC_OPERATIONS.includes(request.operation)) errors.push(`Cible ${target.nodeId}, ${field}: opération inconnue ${request.operation}.`);
            if (Object.prototype.hasOwnProperty.call(request, 'value')) errors.push(`Cible ${target.nodeId}, ${field}: value est interdit pour bulk_drupal_theme_public_update.`);
          }
          if (!Array.isArray(target.expectedAfter?.[field])) errors.push(`Cible ${target.nodeId}, ${field}: expectedAfter doit être un tableau.`);
          else if (request?.values && !compareTaxonomySets(request.values, target.expectedAfter[field])) errors.push(`Cible ${target.nodeId}, ${field}: changes.${field}.values doit correspondre à expectedAfter.${field}.`);
        }
        const key = `${target.nodeId}|${field}`;
        const signature = JSON.stringify(request);
        if (seen.has(key) && seen.get(key) !== signature) errors.push(`Doublon contradictoire pour Node ID ${target.nodeId}, champ ${field}.`);
        seen.set(key, signature);
      });
    });
    if (payload.type === 'bulk_drupal_theme_public_update' && payload.selection?.selectedCount != null && Number(payload.selection.selectedCount) !== targets.length) warnings.push(`selection.selectedCount (${payload.selection.selectedCount}) diffère du nombre de cibles (${targets.length}).`);
    return { errors: [...new Set(errors)], warnings: [...new Set(warnings)], targets };
  }

  function normalizeIntranetTargets(payload) {
    if (payload?.type === 'update_drupal_pages') {
      return (payload.items || []).flatMap((item) => (item.targets || []).map((target) => ({
        batchId: payload.batchId, payloadType: payload.type, intentionId: item.intentionId || '', dispositifCode: item.dispositifCode || '',
        requestedTitle: item.requestedTitle || '', linkedCourses: Array.isArray(item.linkedCourses) ? item.linkedCourses : [],
        nodeId: String(target.nodeId || '').trim(), pageUrl: target.pageUrl || '', inventorySheet: target.inventorySheet || '',
        departmentCodes: Array.isArray(target.departmentCodes) ? target.departmentCodes : [], expectedBefore: target.expectedBefore || {}, changes: target.changes || {}
      })));
    }
    if (payload?.type === 'bulk_drupal_theme_public_update') return (payload.targets || []).map((target) => ({
      batchId: payload.batchId, payloadType: payload.type, nodeId: String(target.nodeId || '').trim(), pageUrl: target.pageUrl || '',
      title: target.title || '', requestedTitle: target.title || '', inventorySheet: target.inventorySheet || '', sourceOvp: target.sourceOvp || '',
      expectedBefore: target.expectedBefore || {}, changes: target.changes || {}, expectedAfter: target.expectedAfter || {}
    }));
    return [];
  }

  function buildIntranetPreview(payload, targets) {
    return targets.map((target) => {
      const desired = payload.type === 'update_drupal_pages'
        ? Object.fromEntries(Object.entries(target.changes).map(([field, change]) => [field, change.value]))
        : Object.fromEntries(Object.entries(target.changes).map(([field, change]) => [field, change.values]));
      const destructive = payload.type === 'update_drupal_pages'
        ? Object.values(target.changes).filter((change) => change.operation === 'delete').length
        : Object.values(target.changes).filter((change) => ['remove', 'replace', 'clear'].includes(change.operation)).length;
      return { nodeId: target.nodeId, title: target.requestedTitle || target.title || '', fields: Object.keys(desired), before: target.expectedBefore || {}, after: desired, destructive };
    });
  }

  function restoreIntranetPayload() {
    const payload = readStoredJson(STORAGE_KEYS.intranetPayload, null);
    if (!payload) return;
    const validation = validateIntranetPayload(payload);
    state.intranetUpdate.payload = payload;
    state.intranetUpdate.payloadType = payload.type || '';
    state.intranetUpdate.payloadFileName = readStoredValue(STORAGE_KEYS.intranetPayloadFileName, 'payload restauré');
    state.intranetUpdate.normalizedTargets = validation.targets;
    state.intranetUpdate.validationErrors = validation.errors;
    state.intranetUpdate.validationWarnings = validation.warnings;
    state.intranetUpdate.preview = buildIntranetPreview(payload, validation.targets);
    const persistedReport = readStoredJson(STORAGE_KEYS.intranetReport, null);
    state.intranetUpdate.reportItems = persistedReport?.items || [];
    state.intranetUpdate.startedAt = persistedReport?.startedAt || '';
    state.intranetUpdate.finishedAt = persistedReport?.finishedAt || '';
    if (!validation.errors.length) state.activeWorkflow = payload.type === 'update_drupal_pages' ? 'intranet_content' : 'intranet_taxonomy';
  }

  function renderWorkflowMode() {
    const element = qs('#eafc-workflow-mode');
    if (!element) return;
    const labels = { sofia: 'Sofia-FMO : dates, lieux et préinscriptions', intranet_content: 'Mise à jour des contenus intranet', intranet_taxonomy: 'Mise à jour massive des thématiques/publics' };
    const safe = state.intranetUpdate.payload?.settings?.safeMode === true && state.activeWorkflow !== 'sofia' ? '<br><strong>Enregistrement interdit par le mode sécurisé du fichier importé.</strong>' : '';
    element.innerHTML = `<strong>Mode actif : ${labels[state.activeWorkflow] || labels.sofia}</strong>${safe}`;
  }

  function renderIntranetPreview() {
    const box = qs('#eafc-intranet-preview');
    if (!box) return;
    const payload = state.intranetUpdate.payload;
    if (!payload) { box.innerHTML = ''; return; }
    const preview = state.intranetUpdate.preview;
    const fields = preview.reduce((sum, row) => sum + row.fields.length, 0);
    const destructive = preview.reduce((sum, row) => sum + row.destructive, 0);
    box.innerHTML = `<details open><summary><strong>Prévisualisation intranet</strong></summary>
      <div>Type : ${escapeHtml(payload.type)} | Batch ID : ${escapeHtml(payload.batchId || '')}</div>
      <div>Source : ${escapeHtml(state.intranetUpdate.payloadFileName)} | Cibles : ${preview.length} | Champs : ${fields} | Erreurs : ${state.intranetUpdate.validationErrors.length} | Opérations destructives : ${destructive}</div>
      ${preview.map((row) => `<div><strong>Node ${escapeHtml(row.nodeId)}</strong> ${escapeHtml(row.title)}<br>Champs : ${escapeHtml(row.fields.join(', '))}<br>Avant : ${escapeHtml(JSON.stringify(row.before))}<br>Cible : ${escapeHtml(JSON.stringify(row.after))}</div>`).join('')}
      ${state.intranetUpdate.validationWarnings.map((warning) => `<div>⚠ ${escapeHtml(warning)}</div>`).join('')}</details>`;
  }

  async function startIntranetBatch() {
    const intranet = state.intranetUpdate;
    if (!intranet.payload || intranet.validationErrors.length || !intranet.normalizedTargets.length) {
      setAlert('Impossible de démarrer : importez un payload intranet valide contenant au moins une cible.'); return;
    }
    const existing = readBatchState();
    if (existing.active) {
      const existingWorkflow = existing.workflowType || 'sofia';
      const requestedWorkflow = state.activeWorkflow;
      if (existingWorkflow !== requestedWorkflow) {
        log('warning', null, `Ancien batch ${existingWorkflow} annulé : le workflow demandé est ${requestedWorkflow}.`);
        clearBatchState();
      } else {
        return runPersistedBatchStep();
      }
    }
    const safe = intranet.payload.settings?.safeMode === true;
    const persistentAllowed = !CONFIG.DRY_RUN && CONFIG.AUTO_SAVE && !safe;
    const blockingReasons = [
      CONFIG.DRY_RUN ? 'mode simulation actif' : '',
      !CONFIG.AUTO_SAVE ? 'enregistrement automatique désactivé' : '',
      safe ? 'mode sécurisé du JSON actif' : ''
    ].filter(Boolean).join(', ') || 'aucun blocage';
    const summary = `${intranet.normalizedTargets.length} cible(s), workflow ${state.activeWorkflow}.\nSimulation : ${CONFIG.DRY_RUN ? 'oui' : 'non'}\nEnregistrement automatique : ${CONFIG.AUTO_SAVE ? 'oui' : 'non'}\nMode sécurisé du JSON : ${safe ? 'oui' : 'non'}\nModification persistante possible : ${persistentAllowed ? 'oui' : 'non'}\nBlocage éventuel : ${blockingReasons}\nContinuer ?`;
    if (!window.confirm(summary)) return;
    intranet.reportItems = []; intranet.startedAt = new Date().toISOString(); intranet.finishedAt = ''; intranet.backupItems = []; intranet.backupGenerated = false;
    persistIntranetReport();
    writeBatchState({ active: true, index: 0, phase: 'goToAdmin', workflowType: state.activeWorkflow, startedAt: intranet.startedAt });
    state.stopRequested = false;
    await runPersistedBatchStep();
  }

  function resolveDrupalTargetField(payloadType, fieldName) {
    if (fieldName === 'publics') return payloadType === 'bulk_drupal_theme_public_update' ? 'metiersPublics' : 'editorialPublics';
    if (fieldName === 'thematiques') return 'theme';
    return fieldName;
  }

  async function updateCurrentDrupalPageFromIntranetPayload(target) {
    const report = createIntranetReportItem(target);
    try {
      logIntranetTargetMode(target);
      if (isAccessDeniedPage()) throw typedError('access_denied', 'Accès Drupal refusé ou redirection vers la connexion.');
      const before = await readCurrentDrupalValues(target);
      report.before = before.values; report.title = before.title || target.requestedTitle || target.title || '';
      const conflicts = compareExpectedBefore(target, before.values);
      if (conflicts.length && state.intranetUpdate.payload.settings?.stopOnConflict === true) {
        report.status = 'conflict'; report.failedFields = conflicts.map((entry) => entry.field); report.errors = conflicts;
        return report;
      }
      await ensureIntranetBackup(target, before);
      if (hasRequestedEditorialChanges(target) || hasRequestedSimpleDrupalChanges(target)) {
        const editorial = await applyEditorialChanges(target, before.values);
        mergeApplyResult(report, editorial);
      }
      if (hasRequestedTaxonomyChanges(target)) {
        const taxonomy = await applyTaxonomyChanges(target, before.values);
        mergeApplyResult(report, taxonomy);
      }
      assertOnlyRequestedFieldsUpdated(target, report);
      report.after = { ...before.values, ...report.after };
      report.status = deriveReportStatus(report);
      return report;
    } catch (error) {
      const code = error.code || (/introuvable/i.test(error.message) ? 'field_not_found' : 'partial_success');
      report.status = REPORT_STATUSES.includes(code) ? code : 'partial_success';
      report.errors.push({ code, message: error.message });
      return report;
    }
  }

  function createIntranetReportItem(target) {
    return { nodeId: String(target.nodeId || ''), pageUrl: target.pageUrl || location.href, title: target.requestedTitle || target.title || '', status: 'skipped', updatedFields: [], unchangedFields: [], failedFields: [], before: {}, after: {}, saved: false, errors: [], _verified: true, _conflict: false };
  }

  function typedError(code, message) { const error = new Error(message); error.code = code; return error; }
  function isAccessDeniedPage() { return /\/user\/login/.test(location.pathname) || /accès refusé|access denied/i.test(document.title + ' ' + document.body.textContent.slice(0, 1000)); }

  async function readCurrentDrupalValues(target) {
    const requestedFields = getRequestedFields(target);
    const values = {};
    const result = { title: target.requestedTitle || target.title || '', values, fullHtml: '', thematics: [], metiersPublics: [], departements: [] };
    if (requestedFields.includes('title')) {
      const titleInput = findElementBySelectors(SELECTORS.title);
      result.title = titleInput?.value || '';
      values.title = result.title;
    }
    const htmlFields = requestedFields.filter((field) => {
      const resolved = resolveDrupalTargetField(target.payloadType, field);
      return resolved === 'editorialPublics' || EDITORIAL_FIELDS.includes(resolved);
    });
    if (htmlFields.length) {
      await prepareFormationEditor(target);
      result.fullHtml = await getCurrentHtmlFromEditor();
      Object.assign(values, extractRequestedFormationHtmlValues(result.fullHtml, htmlFields));
    }
    const taxonomyFieldsToRead = requestedFields.filter((field) => ['theme', 'thematiques', 'metiersPublics', 'departements'].includes(field) || (target.payloadType === 'bulk_drupal_theme_public_update' && field === 'publics'));
    for (const originalField of taxonomyFieldsToRead) {
      const field = resolveDrupalTargetField(target.payloadType, originalField);
      const current = readTaxonomyField(field);
      values[originalField] = current;
      if (field === 'theme') result.thematics = current;
      if (field === 'metiersPublics') result.metiersPublics = current;
      if (field === 'departements') result.departements = current;
    }
    readSimpleDrupalFields(target, values);
    return result;
  }

  function getRequestedFields(target) { return Object.keys(target.changes || {}); }

  function getResolvedRequestedFields(target) {
    return getRequestedFields(target).map((field) => resolveDrupalTargetField(target.payloadType, field));
  }

  function hasRequestedEditorialChanges(target) {
    return getResolvedRequestedFields(target).some((field) => field === 'editorialPublics' || EDITORIAL_FIELDS.includes(field));
  }

  function hasRequestedTaxonomyChanges(target) {
    return getResolvedRequestedFields(target).some((field) => TAXONOMY_FIELDS.includes(field));
  }

  function hasRequestedSimpleDrupalChanges(target) {
    return getResolvedRequestedFields(target).some((field) => ['title', 'publicationDate', 'preinscriptionStartDate', 'unpublishDate', 'preinscriptionLinks'].includes(field));
  }

  function isTaxonomyOnlyTarget(target) {
    return hasRequestedTaxonomyChanges(target) && !hasRequestedEditorialChanges(target) && !hasRequestedSimpleDrupalChanges(target);
  }

  function logIntranetTargetMode(target, workflowType = state.activeWorkflow) {
    log('info', target, `Node ${target.nodeId} — Payload : ${target.payloadType}; Workflow : ${workflowType}; Champs explicitement demandés : ${getRequestedFields(target).join(', ') || '(aucun)'}; Cible limitée aux taxonomies : ${isTaxonomyOnlyTarget(target) ? 'oui' : 'non'}; CKEditor autorisé : ${hasRequestedEditorialChanges(target) ? 'oui' : 'non'}; Workflow Sofia autorisé : ${workflowType === 'sofia' ? 'oui' : 'non'}.`);
  }

  async function prepareFormationEditor(target) {
    if (!hasRequestedEditorialChanges(target)) throw typedError('unauthorized_field_access', `ERREUR : tentative de modification d’un champ non demandé — ouverture de CKEditor refusée pour le Node ${target.nodeId}.`);
    await clickFormationBodyEditBeforeRichText(target);
    const format = await waitForAnySelector(SELECTORS.textFormat, CONFIG.MAX_WAIT_TIME).catch(() => null);
    if (format && format.value !== CONFIG.ADVANCED_FORMAT_VALUE) { setSelectValue(format, CONFIG.ADVANCED_FORMAT_VALUE); await waitForAjax(); }
    await openCkEditorSource();
  }

  function extractRequestedFormationHtmlValues(html, fields) {
    const documentHtml = parseHtmlFragment(html), values = {};
    fields.forEach((field) => {
      if (['objectif', 'contenu', 'accessibilite', 'contact'].includes(field)) values[field] = readSectionValue(documentHtml, editorialFieldDefinition(field).label);
      else if (field === 'miseEnPlace') values[field] = readMiseEnPlaceDescription(documentHtml);
      else values[field] = readStructuredLineValue(documentHtml, editorialFieldDefinition(field)?.label || '');
    });
    return values;
  }

  async function applyEditorialChanges(target, beforeValues = {}) {
    const result = emptyApplyResult();
    if (target.payloadType !== 'update_drupal_pages') return result;
    const changes = target.changes || {};
    if (changes.title) {
      const input = findElementBySelectors(SELECTORS.title);
      if (!input) failField(result, 'title', 'field_not_found', 'Champ titre Drupal introuvable.');
      else if (changes.title.operation === 'delete') failField(result, 'title', 'invalid_operation', 'Le titre Drupal ne peut pas être vidé.');
      else {
        const desired = applyScalarOperation(input.value, changes.title);
        if (compareTextValues(input.value, desired)) unchangedField(result, 'title', input.value);
        else { setInputValue(input, desired); updatedField(result, 'title', input.value, input.value); }
      }
    }
    const htmlChanges = Object.fromEntries(Object.entries(changes).filter(([field]) => EDITORIAL_FIELDS.includes(field)));
    if (Object.keys(htmlChanges).length) {
      await prepareFormationEditor(target);
      const currentHtml = await getCurrentHtmlFromEditor();
      const transformed = applyRequestedFormationHtmlChanges(currentHtml, htmlChanges);
      transformed.updatedFields.forEach((field) => updatedField(result, field, beforeValues[field], transformed.after[field]));
      transformed.unchangedFields.forEach((field) => unchangedField(result, field, transformed.after[field]));
      transformed.failedFields.forEach((field) => failField(result, field, 'field_not_found', `Section ou ligne HTML introuvable pour ${field}.`));
      if (transformed.html !== currentHtml) await setHtmlInEditor(transformed.html);
    }
    await applySimpleDrupalChanges(target, result, beforeValues);
    return result;
  }

  function applyRequestedFormationHtmlChanges(html, changes) {
    const documentHtml = parseHtmlFragment(html), result = { html: String(html || ''), updatedFields: [], unchangedFields: [], failedFields: [], after: {} };
    Object.entries(changes || {}).forEach(([field, change]) => {
      const definition = editorialFieldDefinition(field);
      if (!definition) return;
      let outcome;
      if (definition.kind === 'section') outcome = mutateSectionValue(documentHtml, definition.label, change);
      else if (definition.kind === 'description') outcome = mutateMiseEnPlaceDescription(documentHtml, change);
      else outcome = mutateStructuredLine(documentHtml, definition.label, change);
      if (!outcome.found) result.failedFields.push(field);
      else if (!outcome.changed) result.unchangedFields.push(field);
      else result.updatedFields.push(field);
      result.after[field] = outcome.value;
    });
    result.html = documentHtml.body.innerHTML.trim();
    return result;
  }

  function parseHtmlFragment(html) { return new DOMParser().parseFromString(`<body>${String(html || '')}</body>`, 'text/html'); }
  function normalizedHtmlLabel(element) { return normalizeText(element?.textContent || '').replace(/\s*:\s*$/, ''); }
  function editorialFieldDefinition(field) {
    return {
      objectif: { kind: 'section', label: 'Objectif général du parcours' }, contenu: { kind: 'section', label: 'Contenu de la formation' },
      accessibilite: { kind: 'section', label: 'Accessibilité de la formation' }, contact: { kind: 'section', label: 'Contact de la formation' },
      miseEnPlace: { kind: 'description', label: 'Mise en place de la formation' }, effectif: { kind: 'line', label: 'Effectif' },
      dureeTotale: { kind: 'line', label: 'Durée totale du parcours' }, presentielH: { kind: 'line', label: 'Durée du présentiel' },
      distancielH: { kind: 'line', label: 'Durée du distanciel' }, publics: { kind: 'line', label: 'Public concerné' }, prerequis: { kind: 'line', label: 'Prérequis du parcours' }
    }[field] || null;
  }
  function findHeading(documentHtml, label) { const wanted = normalizeText(label); return Array.from(documentHtml.querySelectorAll('h2,h3')).find((heading) => normalizedHtmlLabel(heading) === wanted) || null; }
  function sectionElements(heading) { const elements = []; if (!heading) return elements; for (let node = heading.nextElementSibling; node && !/^H[23]$/.test(node.tagName); node = node.nextElementSibling) elements.push(node); return elements; }
  function readSectionValue(documentHtml, label) { return sectionElements(findHeading(documentHtml, label)).map((element) => element.textContent.trim()).filter(Boolean).join('\n'); }
  function applyScalarOperation(current, change) {
    const value = String(change?.value ?? '');
    if (change.operation === 'delete') return '';
    if (change.operation === 'add') {
      if (!value || normalizeText(current).includes(normalizeText(value))) return String(current || '');
      return current ? `${current}\n${value}` : value;
    }
    return value;
  }
  function mutateSectionValue(documentHtml, label, change) {
    const heading = findHeading(documentHtml, label); if (!heading) return { found: false, changed: false, value: '' };
    const elements = sectionElements(heading), current = elements.map((element) => element.textContent.trim()).filter(Boolean).join('\n');
    const desired = applyScalarOperation(current, change);
    if (compareTextValues(current, desired)) return { found: true, changed: false, value: current };
    if (change.operation === 'add' && elements.length) { const paragraph = documentHtml.createElement('p'); paragraph.textContent = String(change.value ?? ''); elements.at(-1).after(paragraph); }
    else if (elements.length) { elements[0].innerHTML = escapeHtml(desired).replace(/\n/g, '<br>'); elements.slice(1).forEach((element) => element.remove()); }
    else { const paragraph = documentHtml.createElement('p'); paragraph.textContent = desired; heading.after(paragraph); }
    return { found: true, changed: true, value: desired };
  }
  function findStructuredLine(documentHtml, label) {
    const wanted = normalizeText(label);
    const heading = findHeading(documentHtml, 'Mise en place de la formation');
    return sectionElements(heading).flatMap((element) => [element, ...element.querySelectorAll('li,p')]).find((element) => {
      const strong = element.querySelector('strong,b');
      const candidate = normalizedHtmlLabel(strong || element).split(':')[0].trim();
      return candidate === wanted || normalizeText(element.textContent).replace(/\s*:.*/, '') === wanted;
    }) || null;
  }
  function readStructuredLineValue(documentHtml, label) { const line = findStructuredLine(documentHtml, label); if (!line) return ''; const clone = line.cloneNode(true); clone.querySelector('strong,b')?.remove(); return clone.textContent.replace(/^\s*:\s*/, '').trim(); }
  function mutateStructuredLine(documentHtml, label, change) {
    const line = findStructuredLine(documentHtml, label); if (!line) return { found: false, changed: false, value: '' };
    const current = readStructuredLineValue(documentHtml, label), desired = applyScalarOperation(current, change);
    if (compareTextValues(current, desired)) return { found: true, changed: false, value: current };
    const marker = line.querySelector('strong,b');
    if (marker) { while (marker.nextSibling) marker.nextSibling.remove(); marker.after(documentHtml.createTextNode(` : ${desired}`)); }
    else line.textContent = `${label} : ${desired}`;
    return { found: true, changed: true, value: desired };
  }
  function readMiseEnPlaceDescription(documentHtml) { const candidates = sectionElements(findHeading(documentHtml, 'Mise en place de la formation')).filter((element) => element.tagName === 'P' && !/:/.test(element.textContent) && !/^session\s+\d+/i.test(element.textContent.trim())); return candidates.length === 1 ? candidates[0].textContent.trim() : ''; }
  function mutateMiseEnPlaceDescription(documentHtml, change) {
    const candidates = sectionElements(findHeading(documentHtml, 'Mise en place de la formation')).filter((element) => element.tagName === 'P' && !/:/.test(element.textContent) && !/^session\s+\d+/i.test(element.textContent.trim()));
    if (candidates.length !== 1) return { found: false, changed: false, value: '' };
    const current = candidates[0].textContent.trim(), desired = applyScalarOperation(current, change);
    if (compareTextValues(current, desired)) return { found: true, changed: false, value: current };
    if (change.operation === 'add' && current) { const paragraph = documentHtml.createElement('p'); paragraph.textContent = String(change.value ?? ''); candidates[0].after(paragraph); }
    else candidates[0].textContent = desired;
    return { found: true, changed: true, value: desired };
  }
  function compareTextValues(a, b) { return String(a ?? '').replace(/[\s\u00a0]+/g, ' ').trim() === String(b ?? '').replace(/[\s\u00a0]+/g, ' ').trim(); }

  async function applyTaxonomyChanges(target) {
    const result = emptyApplyResult();
    const desiredEntries = Object.entries(target.changes || {}).filter(([field]) => TAXONOMY_FIELDS.includes(field) || (target.payloadType === 'bulk_drupal_theme_public_update' && field === 'publics'));
    if (desiredEntries.length) log('info', target, `Node ${target.nodeId} — taxonomies demandées : ${desiredEntries.map(([field]) => field).join(', ')}`);
    for (const [originalField, request] of desiredEntries) {
      const field = resolveDrupalTargetField(target.payloadType, originalField);
      if (!['theme', 'metiersPublics', 'departements'].includes(field)) continue;
      const desired = target.payloadType === 'bulk_drupal_theme_public_update' ? normalizeTerms(request.values) : normalizeTerms(request.value);
      try {
        const current = readTaxonomyField(field);
        if (field === 'theme') log('info', target, `Thématiques actuelles : ${JSON.stringify(current)}. Thématiques attendues : ${JSON.stringify(desired)}.`);
        if (field === 'metiersPublics') log('info', target, `Publics actuels : ${JSON.stringify(current)}. Publics attendus : ${JSON.stringify(desired)}.`);
        if (compareTaxonomySets(current, desired)) {
          unchangedField(result, originalField, current);
          log('success', target, `${field === 'theme' ? 'Thématiques' : 'Publics'} vérifiées avec succès.`);
          continue;
        }
        if (field === 'theme') await applyThematicsExpectedState(desired);
        else {
          const container = field === 'metiersPublics' ? findPublicTreeContainer() : findTaxonomyContainerByLabel(['Zone géographique', 'Départements', 'Département']);
          if (!container) throw typedError('field_not_found', field === 'metiersPublics' ? 'Champ publics introuvable : #edit-field-metier-tags-wrapper absent' : `Champ Drupal distinct ${field} introuvable.`);
          await applyFancyTreeExpectedState(container, desired);
        }
        const after = readTaxonomyField(field);
        if (!compareTaxonomySets(after, desired)) throw typedError('field_not_found', `Valeur finale incorrecte pour ${originalField}.`);
        updatedField(result, originalField, current, after);
        log('success', target, `${field === 'theme' ? 'Thématiques' : 'Publics'} vérifiées avec succès.`);
      } catch (error) {
        log('error', target, error.message);
        failField(result, originalField, error.code || 'field_not_found', error.message);
      }
    }
    return result;
  }

  function normalizeTerms(value) { return (Array.isArray(value) ? value : String(value ?? '').split(/[;,|\n]/)).map((term) => typeof term === 'object' ? String(term.label || term.name || '') : String(term)).map((term) => term.trim()).filter(Boolean); }
  function normalizedTaxonomySet(terms) { return [...new Set(normalizeTerms(terms).map(normalizeText))].sort(); }
  function compareTaxonomySets(a, b) { return JSON.stringify(normalizedTaxonomySet(a)) === JSON.stringify(normalizedTaxonomySet(b)); }


  /***************************************************************************
   * THÉMATIQUES / PUBLICS — import du module isolé
   *
   * Sélectionne uniquement des termes déjà existants dans Drupal : aucune
   * création de termes de taxonomie n'est effectuée ici.
   ***************************************************************************/
  async function applyPublicsAndThematicsFromSource(item = {}, exportData = {}) {
    await applyMappedThematics({ item, exportData });
    await applyMappedPublics({ item, exportData });
  }

  /***************************************************************************
   * THÉMATIQUES
   ***************************************************************************/
  async function applyMappedThematics({ item = {}, exportData = {} } = {}) {
    const thematics = resolveMappedThematics({ item, exportData });
    if (!thematics.length) { log('info', item, 'Aucune thématique à ajouter.'); return []; }
    log('info', item, `Thématiques à ajouter : ${thematics.join(', ')}.`);
    let current = readTaxonomyField('theme');
    for (const thematic of thematics) {
      if (current.some((value) => normalizeText(value) === normalizeText(thematic))) continue;
      await addAutocompleteTermExact(thematic);
      current = readTaxonomyField('theme');
    }
    return thematics;
  }

  function resolveMappedThematics({ item = {}, exportData = {} } = {}) {
    const rawTheme = [exportData.theme, item.theme].filter(Boolean).join(', ');
    const normalizedTheme = normalize(rawTheme);
    const values = [];
    CONFIG.THEMATIC_ALIAS_MAP.forEach(([aliases, thematicIds]) => {
      if (!aliases.some((alias) => normalizedTheme.includes(normalize(alias)))) return;
      thematicIds.forEach((thematicId) => {
        const value = String(thematicId).trim();
        if (value && !values.includes(value)) values.push(value);
      });
    });
    CONFIG.REQUIRED_THEMATICS.forEach((thematic) => {
      const value = String(thematic).trim();
      if (value && !values.includes(value)) values.push(value);
    });
    return values;
  }

  async function getThematicInputByIndex(index) {
    await waitForAnySelector(SELECTORS.thematics, CONFIG.MAX_WAIT_TIME);
    const inputs = Array.from(document.querySelectorAll(SELECTORS.thematics.join(','))).filter(isVisible);
    const input = inputs[index] || inputs.find((candidate) => !candidate.value) || inputs.at(-1);
    if (!input) throw new Error(`Champ thématique introuvable pour l'index ${index}.`);
    return input;
  }

  async function fillThematicAutocomplete(input, value) {
    if (!input) throw new Error(`Champ d'autocomplétion introuvable pour : ${value}.`);
    setValue(input, value);
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    await sleep(Math.max(CONFIG.DELAY_BETWEEN_ACTIONS, CONFIG.THEMATIC_AUTOCOMPLETE_DELAY_MS));
    await waitForFunction(() => visibleAutocompleteSuggestions().length > 0, 5000, `Suggestion thématique introuvable : ${value}.`);
    const normalizedValue = normalize(value);
    const suggestions = visibleAutocompleteSuggestions();
    const matches = suggestions.filter((element) => {
      const text = normalize(element.textContent);
      return text.includes(normalizedValue) || text.includes(`(${normalizedValue})`);
    });
    if (!matches.length) throw new Error(`Suggestion thématique introuvable pour la valeur recherchée : ${value}.`);
    await safeClick(matches[0], `Valider l'autocomplétion : ${value}`);
    await waitForFunction(() => normalize(input.value) !== normalizedValue || !isVisible(input), 3000, `Thématique non validée par Drupal : ${value}.`);
    await sleep(Math.max(CONFIG.DELAY_BETWEEN_ACTIONS, CONFIG.THEMATIC_AUTOCOMPLETE_SELECTION_DELAY_MS));
    await waitForAjax();
  }

  function visibleAutocompleteSuggestions() {
    return Array.from(document.querySelectorAll('.ui-autocomplete li, .ui-menu-item, [role="option"]')).filter(isVisible);
  }

  /***************************************************************************
   * PUBLICS
   ***************************************************************************/
  async function applyMappedPublics({ item = {}, exportData = {} } = {}) {
    const publics = uniquePublicEntries([...(Array.isArray(exportData.publics) ? exportData.publics : exportData.publics ? [exportData.publics] : []), ...(Array.isArray(item.publics) ? item.publics : item.publics ? [item.publics] : [])]);
    if (!publics.length) {
      await clickFancytreeTitle('Tous les personnels', 'Sélectionner le métier par défaut : Tous les personnels', { optional: true, scopeSelector: SELECTORS.publicTreeWrapper[0] });
      return ['Tous les personnels'];
    }
    const selectedPaths = [];
    for (const entry of publics) {
      const path = publicPath(entry);
      if (!path.length) continue;
      const selected = await clickFancytreePath(SELECTORS.publicTreeWrapper[0], path, `Sélectionner le métier : ${path.join(' > ')}`, { optional: true });
      if (selected) selectedPaths.push(path.join(' > '));
    }
    return selectedPaths;
  }

  async function clickFancytreePath(scopeSelector, path, message, options = {}) {
    const cleanPath = path.map(String).map((value) => value.trim()).filter(Boolean);
    for (const title of cleanPath.slice(0, -1)) await ensureFancytreeExpanded(title, { ...options, scopeSelector });
    return clickFancytreeTitle(cleanPath.at(-1), message, { ...options, scopeSelector });
  }

  async function ensureFancytreeExpanded(title, options = {}) {
    const node = findFancytreeNode(title, options);
    if (!node) { if (options.optional) { log('warning', null, `Nœud public optionnel introuvable : ${title}.`); return false; } throw new Error(`Nœud Fancytree introuvable : ${title}.`); }
    const treeItem = node.closest('[role="treeitem"],li');
    const isExpanded = treeItem?.getAttribute('aria-expanded') === 'true' || node.classList.contains('fancytree-expanded') || treeItem?.classList.contains('fancytree-expanded');
    if (!isExpanded) {
      const expander = node.querySelector('.fancytree-expander') || node.closest('.fancytree-node,li')?.querySelector('.fancytree-expander');
      if (expander) { await safeClick(expander, `Dérouler ${title}`); await waitForAjax(); }
    }
    return true;
  }

  async function clickFancytreeTitle(title, message, options = {}) {
    await waitForFunction(() => findFancytreeNode(title, options) || options.optional, 8000, `Public introuvable : ${title}.`);
    const node = findFancytreeNode(title, options);
    if (!node) { log('warning', null, `Champ public optionnel non trouvé : ${title}.`); return false; }
    const checkbox = node.querySelector('.fancytree-checkbox,input[type="checkbox"]') || node.closest('.fancytree-node,li')?.querySelector('.fancytree-checkbox,input[type="checkbox"]');
    if (!checkbox) { if (options.optional) { log('warning', null, `Checkbox publique optionnelle introuvable : ${title}.`); return false; } throw new Error(`Checkbox Fancytree introuvable : ${title}.`); }
    const treeItem = node.closest('[role="treeitem"],li');
    const isSelected = treeItem?.getAttribute('aria-selected') === 'true' || node.classList.contains('fancytree-selected') || node.classList.contains('fancytree-partsel') || checkbox.checked || checkbox.getAttribute('aria-checked') === 'true';
    if (!isSelected) { await safeClick(checkbox, message); await waitForAjax(); }
    else log('info', null, `Public déjà sélectionné : ${title}.`);
    return true;
  }

  function findFancytreeNode(title, options = {}) {
    const wanted = normalize(title), scope = options.scopeSelector ? document.querySelector(options.scopeSelector) : document;
    if (!scope) return null;
    return Array.from(scope.querySelectorAll('.fancytree-node,li')).find((node) => {
      const text = normalize(node.querySelector('.fancytree-title,label')?.textContent || '');
      return options.contains ? text.includes(wanted) : text === wanted;
    }) || null;
  }

  function uniquePublicEntries(values) {
    const seen = new Set();
    return values.filter((value) => {
      if (!value) return false;
      const key = typeof value === 'string' ? normalize(value) : normalize(value.id || value.path || value.label);
      if (!key || seen.has(key)) return false;
      seen.add(key); return true;
    });
  }

  function publicPath(value) {
    if (typeof value === 'string') return [value.trim()].filter(Boolean);
    if (value?.path) return String(value.path).split('>').map((part) => part.trim()).filter(Boolean);
    return [value?.label || value?.id].map((entry) => String(entry || '').trim()).filter(Boolean);
  }

  /***************************************************************************
   * DÉPENDANCES PARTAGÉES
   *
   * Ces alias réutilisent les utilitaires historiques du script cible pour
   * éviter toute deuxième implémentation conflictuelle.
   ***************************************************************************/
  const safeClick = clickElement;
  const setValue = setInputValue;
  const normalize = normalizeText;
  function findTaxonomyContainerByLabel(labels) {
    const wanted = labels.map(normalizeText), candidates = Array.from(document.querySelectorAll('fieldset,.form-item,.field--widget-fancytree,.js-form-wrapper'));
    return candidates.filter((container) => { const label = container.querySelector(':scope > legend, :scope > label, :scope > .fieldset-legend, :scope > .form-item__label'); return label && wanted.includes(normalizeText(label.textContent)); }).sort((a, b) => a.querySelectorAll('*').length - b.querySelectorAll('*').length)[0] || null;
  }
  function findPublicTreeContainer() {
    return document.querySelector(SELECTORS.publicTreeWrapper[0]) || findTaxonomyContainerByLabel(['Métiers', 'Métier', 'Publics', 'Public']);
  }

  function getSelectedFancyTreeTerms(container) {
    if (!container) throw typedError('field_not_found', 'Champ publics introuvable : #edit-field-metier-tags-wrapper absent');
    const values = [];
    const add = (value) => { const clean = String(value || '').trim(); if (clean && !values.some((entry) => normalizeText(entry) === normalizeText(clean))) values.push(clean); };
    Array.from(container.querySelectorAll('.fancytree-node.fancytree-selected')).forEach((node) => add(node.querySelector('.fancytree-title')?.textContent || ''));
    Array.from(container.querySelectorAll('.fancytree-node input[type="checkbox"]:checked')).forEach((checkbox) => add(checkbox.closest('.fancytree-node')?.querySelector('.fancytree-title')?.textContent || ''));
    return values;
  }

  function findFancyTreeNodeExact(container, label) {
    if (!container) throw typedError('field_not_found', 'Champ publics introuvable : #edit-field-metier-tags-wrapper absent');
    const wanted = normalizeText(label);
    const matches = Array.from(container.querySelectorAll('.fancytree-node')).filter((node) => normalizeText(node.querySelector('.fancytree-title')?.textContent || '') === wanted);
    if (matches.length > 1) throw typedError('ambiguous_term', `Plusieurs nœuds Fancytree portent exactement le libellé : ${label}`);
    return matches[0] || null;
  }

  async function expandFancyTreeAncestors(node) {
    const ancestors = [];
    for (let li = node?.closest('li')?.parentElement?.closest('li'); li; li = li.parentElement?.closest('li')) ancestors.unshift(li);
    for (const ancestor of ancestors) {
      const ancestorNode = ancestor.querySelector(':scope > .fancytree-node');
      const expander = ancestorNode?.querySelector(':scope .fancytree-expander');
      const expanded = ancestor.classList.contains('fancytree-expanded') || ancestorNode?.classList.contains('fancytree-expanded') || ancestor.getAttribute('aria-expanded') === 'true';
      if (expander && !expanded) { await clickElement(expander, 'Déplier la taxonomie'); await waitForAjax(); }
    }
  }

  async function setFancyTreeNodeChecked(node, checked) {
    const checkbox = node.querySelector('input[type="checkbox"],.fancytree-checkbox');
    const selected = node.classList.contains('fancytree-selected') || checkbox?.checked || checkbox?.getAttribute('aria-checked') === 'true';
    if (Boolean(selected) !== Boolean(checked)) { await clickElement(checkbox || node.querySelector('.fancytree-title'), `${checked ? 'Sélectionner' : 'Désélectionner'} le terme exact`); await waitForAjax(); }
  }

  async function applyFancyTreeExpectedState(container, expectedTerms) {
    const expected = normalizeTerms(expectedTerms);
    let selected = getSelectedFancyTreeTerms(container);
    for (const label of selected.filter((term) => !expected.some((wanted) => normalizeText(wanted) === normalizeText(term)))) {
      const node = findFancyTreeNodeExact(container, label);
      if (node) { await expandFancyTreeAncestors(node); await setFancyTreeNodeChecked(node, false); }
    }
    selected = getSelectedFancyTreeTerms(container);
    for (const label of expected.filter((wanted) => !selected.some((term) => normalizeText(wanted) === normalizeText(term)))) {
      const node = findFancyTreeNodeExact(container, label);
      if (!node) throw typedError('field_not_found', `Public introuvable dans Fancytree : ${label}`);
      await expandFancyTreeAncestors(node);
      await setFancyTreeNodeChecked(node, true);
    }
    const final = getSelectedFancyTreeTerms(container);
    if (!compareTaxonomySets(final, expected)) throw typedError('field_not_found', 'État final du FancyTree différent de expectedAfter.');
    return final;
  }

  function readTaxonomyField(field) {
    if (field === 'theme') {
      const inputs = Array.from(document.querySelectorAll(SELECTORS.thematics.join(','))).filter(isVisible);
      if (!inputs.length) throw typedError('field_not_found', 'Champ Thématiques Drupal introuvable.');
      return inputs.map((input) => input.value.replace(/\s*\([^)]*\)\s*$/, '').trim()).filter(Boolean);
    }
    const container = field === 'metiersPublics' ? findPublicTreeContainer() : findTaxonomyContainerByLabel(['Zone géographique', 'Départements', 'Département']);
    if (!container) throw typedError('field_not_found', field === 'metiersPublics' ? 'Champ publics introuvable : #edit-field-metier-tags-wrapper absent' : `Champ Drupal distinct ${field} introuvable.`);
    return getSelectedFancyTreeTerms(container);
  }
  async function applyThematicsExpectedState(expectedTerms) {
    const expected = normalizeTerms(expectedTerms);
    let current = readTaxonomyField('theme');
    for (const value of current.filter((entry) => !expected.some((term) => normalizeText(term) === normalizeText(entry)))) {
      const input = findThematicInputByCurrentValue(value);
      if (!input) throw typedError('field_not_found', `Champ thématique introuvable pour suppression : ${value}.`);
      const row = input.closest('tr') || input.closest('.field-multiple-table') || input.closest('.form-item');
      const remove = row?.querySelector('input[value="Retirer"],button[value="Retirer"],button[aria-label*="Retirer"],input[name*="remove_button"]');
      if (!remove) throw typedError('field_not_found', `Bouton de suppression thématique introuvable pour ${value}.`);
      await clickElement(remove, 'Retirer la thématique exacte');
      await waitForAjax();
      current = readTaxonomyField('theme');
    }
    current = readTaxonomyField('theme');
    for (const label of expected.filter((term) => !current.some((value) => normalizeText(value) === normalizeText(term)))) {
      await addAutocompleteTermExact(label);
      current = readTaxonomyField('theme');
    }
  }

  function findThematicInputByCurrentValue(value) {
    return Array.from(document.querySelectorAll(SELECTORS.thematics.join(','))).filter(isVisible).find((input) => normalizeText(input.value.replace(/\s*\([^)]*\)\s*$/, '')) === normalizeText(value)) || null;
  }

  async function getOrCreateEmptyThematicInput() {
    let inputs = Array.from(document.querySelectorAll(SELECTORS.thematics.join(','))).filter(isVisible);
    let emptyInput = inputs.find((input) => !input.value.trim());
    if (emptyInput) return emptyInput;

    const addButton = await waitForAnySelector(SELECTORS.thematicAddMore, CONFIG.MAX_WAIT_TIME);
    await clickElement(addButton, 'Ajouter un autre élément thématique');
    await waitForAjax();

    inputs = Array.from(document.querySelectorAll(SELECTORS.thematics.join(','))).filter(isVisible);
    emptyInput = inputs.find((input) => !input.value.trim());
    if (!emptyInput) throw typedError('field_not_found', 'Aucune nouvelle ligne thématique vide après le clic Ajouter un autre élément.');
    return emptyInput;
  }

  async function addAutocompleteTermExact(label) {
    const input = await getOrCreateEmptyThematicInput();
    setInputValue(input, label);
    const normalizedLabel = normalizeText(label);
    const suggestions = await waitForFunction(() => Array.from(document.querySelectorAll('.ui-autocomplete li,.ui-menu-item,[role="option"]')).filter(isVisible).filter((entry) => {
      const full = normalizeText(entry.textContent || '');
      const withoutId = normalizeText((entry.textContent || '').replace(/\s*\([^)]*\)\s*$/, ''));
      return withoutId === normalizedLabel || full.includes(`(${normalizedLabel})`);
    }), CONFIG.MAX_WAIT_TIME, `Suggestion thématique introuvable : ${label}`);
    if (suggestions.length !== 1) throw typedError(suggestions.length > 1 ? 'ambiguous_term' : 'field_not_found', `Suggestion thématique introuvable : ${label}`);
    await clickElement(suggestions[0], `Choisir la suggestion exacte ${label}`);
    await waitForAjax();
    await waitForFunction(() => normalizeText(input.value) !== normalizedLabel || !isVisible(input), 3000, `Thématique non validée par Drupal : ${label}`);
  }

  function readSimpleDrupalFields(target, values) {
    const map = { publicationDate: SELECTORS.startDate, preinscriptionStartDate: SELECTORS.startDate, unpublishDate: SELECTORS.unpublishDate };
    getRequestedFields(target).forEach((field) => { if (map[field]) values[field] = findElementBySelectors(map[field])?.value || ''; });
    if (getRequestedFields(target).includes('preinscriptionLinks')) values.preinscriptionLinks = Array.from(document.querySelectorAll(SELECTORS.calloutUri.join(','))).map((input) => input.value).filter(Boolean);
    if (getRequestedFields(target).includes('lieux') && !Object.prototype.hasOwnProperty.call(values, 'lieux')) values.lieux = [];
  }
  async function applySimpleDrupalChanges(target, result, beforeValues) {
    const map = { publicationDate: SELECTORS.startDate, preinscriptionStartDate: SELECTORS.startDate, unpublishDate: SELECTORS.unpublishDate };
    for (const [field, selectors] of Object.entries(map)) {
      const change = target.changes?.[field]; if (!change) continue;
      const input = findElementBySelectors(selectors); if (!input) { failField(result, field, 'field_not_found', `Champ ${field} introuvable.`); continue; }
      const desired = applyScalarOperation(input.value, change);
      if (compareTextValues(input.value, desired)) unchangedField(result, field, input.value); else { const old = input.value; setInputValue(input, desired); updatedField(result, field, old, input.value); }
    }
    if (target.changes?.preinscriptionLinks) {
      const desired = normalizeTerms(target.changes.preinscriptionLinks.value);
      let inputs = Array.from(document.querySelectorAll(SELECTORS.calloutUri.join(','))).filter(isVisible);
      if (!inputs.length && desired.length && target.changes.preinscriptionLinks.operation !== 'delete') { await createCallout({ preinscriptionLink: desired[0] }); inputs = Array.from(document.querySelectorAll(SELECTORS.calloutUri.join(','))).filter(isVisible); }
      if (!inputs.length) failField(result, 'preinscriptionLinks', 'field_not_found', 'Bloc/lien de préinscription introuvable.');
      else if (compareTaxonomySets(inputs.map((input) => input.value), desired)) unchangedField(result, 'preinscriptionLinks', desired);
      else { const old = inputs.map((input) => input.value); setInputValue(inputs[0], desired[0] || ''); updatedField(result, 'preinscriptionLinks', old, desired); }
    }
  }

  function compareExpectedBefore(target, current) {
    const differences = [];
    const requestedFields = new Set(getRequestedFields(target));
    Object.entries(target.expectedBefore || {}).forEach(([field, expected]) => {
      if (!requestedFields.has(field)) return;
      const actual = current[field];
      const resolvedField = resolveDrupalTargetField(target.payloadType, field);
      const taxonomy = Array.isArray(expected) || ['theme', 'metiersPublics', 'departements'].includes(resolvedField);
      const same = taxonomy ? compareTaxonomySets(actual, expected) : compareTextValues(actual, expected);
      if (!same) differences.push({ field, expected, actual, code: 'conflict', message: `expectedBefore différent pour ${field}.` });
    });
    return differences;
  }

  function assertOnlyRequestedFieldsUpdated(target, report) {
    const requested = new Set(getRequestedFields(target));
    report.updatedFields.forEach((field) => {
      if (!requested.has(field)) throw typedError('unauthorized_field_change', `ERREUR : tentative de modification d’un champ non demandé : ${field}.`);
    });
  }
  function emptyApplyResult() { return { updatedFields: [], unchangedFields: [], failedFields: [], before: {}, after: {}, errors: [], verified: true }; }
  function updatedField(result, field, before, after) { if (!result.updatedFields.includes(field)) result.updatedFields.push(field); result.before[field] = before; result.after[field] = after; }
  function unchangedField(result, field, value) { if (!result.unchangedFields.includes(field)) result.unchangedFields.push(field); result.before[field] = value; result.after[field] = value; }
  function failField(result, field, code, message) { if (!result.failedFields.includes(field)) result.failedFields.push(field); result.errors.push({ field, code, message }); result.verified = false; }
  function mergeApplyResult(report, result) { ['updatedFields', 'unchangedFields', 'failedFields'].forEach((key) => { report[key] = [...new Set([...report[key], ...result[key]])]; }); Object.assign(report.before, result.before); Object.assign(report.after, result.after); report.errors.push(...result.errors); report._verified = report._verified && result.verified; }
  function deriveReportStatus(report) { if (report.failedFields.length && report.updatedFields.length) return 'partial_success'; if (report.failedFields.length) return 'field_not_found'; if (!report.updatedFields.length) return 'already_up_to_date'; return 'success'; }

  async function ensureIntranetBackup(target, before) {
    const intranet = state.intranetUpdate;
    intranet.backupItems.push({ nodeId: target.nodeId, title: before.title, url: target.pageUrl || location.href, html: before.fullHtml, thematiques: before.thematics, metiersPublics: before.metiersPublics, departements: before.departements, requestedValues: before.values, readAt: new Date().toISOString() });
    const savingPossible = !CONFIG.DRY_RUN && CONFIG.AUTO_SAVE && intranet.payload.settings?.safeMode !== true;
    if (!savingPossible) return;
    const backup = { type: 'drupal_pre_update_backup', version: '1.0', batchId: intranet.payload.batchId, generatedAt: new Date().toISOString(), restorationAutomatic: false, items: intranet.backupItems };
    downloadFile(`eafc-backup-drupal-avant-maj-${safeFilename(intranet.payload.batchId)}.json`, JSON.stringify(backup, null, 2), 'application/json');
    intranet.backupGenerated = true;
  }
  function safeFilename(value) { return String(value || 'batch').replace(/[^a-z0-9._-]+/gi, '-'); }
  function buildIntranetSaveOptions(report, workflowType) {
    const safe = state.intranetUpdate.payload.settings?.safeMode === true;
    const unauthorized = report.status === 'unauthorized_field_change' || report.errors.some((error) => error.code === 'unauthorized_field_change' || error.code === 'unauthorized_field_access' || error.code === 'workflow_mismatch');
    return { workflowType, forbidden: safe || unauthorized || !report.updatedFields.length, conflict: report.status === 'conflict', fieldsValid: !report.failedFields.length && !unauthorized, verified: report._verified, backupGenerated: state.intranetUpdate.backupGenerated, reason: unauthorized ? 'Enregistrement interdit : tentative de modification ou accès à un champ non demandé.' : safe ? 'Enregistrement interdit par le mode sécurisé du fichier importé.' : !report.updatedFields.length ? 'Aucun changement : enregistrement inutile.' : report.failedFields.length ? 'Enregistrement interdit : certains champs demandés sont introuvables ou non vérifiés.' : '' };
  }
  function finalizeIntranetReportItem(report, saved) {
    report.saved = Boolean(saved);
    if (!saved && report.status === 'success' && !CONFIG.DRY_RUN && CONFIG.AUTO_SAVE && state.intranetUpdate.payload.settings?.safeMode !== true) report.status = 'save_failed';
    delete report._verified; delete report._conflict;
    const others = state.intranetUpdate.reportItems.filter((item) => String(item.nodeId) !== String(report.nodeId));
    state.intranetUpdate.reportItems = [...others, report];
    state.intranetUpdate.finishedAt = new Date().toISOString();
    recordIntranetExcelReport(report);
    persistIntranetReport();
  }
  function markRemainingIntranetTargetsCancelled() {
    const batch = readBatchState();
    if (!batch.active || (batch.workflowType || 'sofia') === 'sofia') return;
    const completed = new Set(state.intranetUpdate.reportItems.map((item) => String(item.nodeId)));
    state.intranetUpdate.normalizedTargets.slice(Number(batch.index || 0)).forEach((target) => {
      if (completed.has(String(target.nodeId))) return;
      const report = createIntranetReportItem(target); report.status = 'cancelled'; report.failedFields = getRequestedFields(target); report.errors = [{ code: 'cancelled', message: 'Arrêt demandé par l’utilisateur.' }];
      state.intranetUpdate.reportItems.push(report);
    });
    state.intranetUpdate.finishedAt = new Date().toISOString(); persistIntranetReport();
  }
  function persistIntranetReport() { writeStoredJson(STORAGE_KEYS.intranetReport, buildIntranetReport()); }
  function buildIntranetReport() { return { type: 'drupal_update_report', version: '1.0', batchId: state.intranetUpdate.payload?.batchId || '', sourcePayloadType: state.intranetUpdate.payloadType || '', generatedAt: new Date().toISOString(), startedAt: state.intranetUpdate.startedAt || '', finishedAt: state.intranetUpdate.finishedAt || '', items: state.intranetUpdate.reportItems.map(({ _verified, _conflict, ...item }) => item) }; }
  function exportIntranetReportJson() { if (!state.intranetUpdate.payload) { setAlert('Aucun payload intranet : aucun rapport JSON à exporter.'); return; } const report = buildIntranetReport(); downloadFile(`eafc-rapport-drupal-${safeFilename(report.batchId)}.json`, JSON.stringify(report, null, 2), 'application/json'); }
  function recordIntranetExcelReport(report) { loadPersistedUpdateReport(); const entry = { intranet: true, dateMiseAJour: new Date().toLocaleString('fr-FR'), batchId: state.intranetUpdate.payload?.batchId || '', payloadType: state.intranetUpdate.payloadType, drupalParcoursId: report.nodeId, drupalTitle: report.title, status: report.status, updatedFields: report.updatedFields.join(', '), unchangedFields: report.unchangedFields.join(', '), failedFields: report.failedFields.join(', '), before: JSON.stringify(report.before), after: JSON.stringify(report.after), saved: report.saved ? 'Oui' : 'Non', errorMessage: report.errors.map((error) => error.message || error.code || error).join(' | ') }; state.updateReport = [...state.updateReport.filter((row) => !(row.intranet && row.batchId === entry.batchId && String(row.drupalParcoursId) === String(entry.drupalParcoursId))), entry]; writeStoredJson(STORAGE_KEYS.updateReport, state.updateReport); }

  /***************************************************************************
   * BILAN EXCEL DES MISES À JOUR
   ***************************************************************************/
  function loadPersistedUpdateReport() {
    state.updateReport = readStoredJson(STORAGE_KEYS.updateReport, []) || [];
  }

  function recordUpdateReport(item, data, dates) {
    loadPersistedUpdateReport();
    const entry = {
      dateMiseAJour: new Date().toLocaleString('fr-FR'),
      drupalParcoursId: item?.drupalParcoursId || '',
      drupalTitle: item?.drupalTitle || '',
      dispositifCode: item?.dispositifCode || '',
      moduleCode: item?.moduleCode || '',
      placesRestantes: calculateRemainingPlaces(data?.preinscritsTotal, data?.preinscritsMax),
      preinscritsTotal: data?.preinscritsTotal || '',
      preinscritsMax: data?.preinscritsMax || '',
      lienPreinscription: data?.preinscriptionLink || '',
      dateDebutSaisie: dates?.startDate || data?.publicationStartDate || '',
      dateFinSaisie: dates?.endDate || data?.publicationEndDate || '',
      dateDepublicationSaisie: dates?.unpublishDate || '',
      sessionsSaisies: formatReportSessions(data?.sessions?.length ? data.sessions : [data])
    };
    const withoutSameItem = state.updateReport.filter((row) => row.drupalParcoursId !== entry.drupalParcoursId);
    state.updateReport = [...withoutSameItem, entry];
    writeStoredJson(STORAGE_KEYS.updateReport, state.updateReport);
    log('success', item, `Bilan Excel mémorisé : ${entry.placesRestantes} place(s) restante(s).`);
  }

  function formatReportSessions(sessions) {
    return (sessions || []).map((session, index) => {
      const start = session.sessionStartDateTime || toDisplayDateTime(session.sessionStartRaw);
      const end = session.sessionEndDateTime || toDisplayDateTime(session.sessionEndRaw);
      return `Session ${index + 1} | Modalité : ${session.modalite || ''} | Date et heure : Du ${start || ''} au ${end || ''} | Lieu : ${session.lieu || ''}`;
    }).join('\n');
  }

  function exportUpdateReportExcel() {
    loadPersistedUpdateReport();
    if (!state.updateReport.length) {
      setAlert('Aucun bilan à exporter pour le moment. Lancez au moins une mise à jour de parcours.');
      log('warning', null, 'Export bilan Excel demandé sans donnée disponible.');
      return;
    }
    if (!window.XLSX) throw new Error('SheetJS XLSX non chargé : export Excel impossible.');

    const rows = state.updateReport.map((row) => row.intranet ? {
      'Date de mise à jour': row.dateMiseAJour, 'ID parcours Drupal': row.drupalParcoursId, 'Titre Drupal': row.drupalTitle,
      'Batch ID': row.batchId, 'Type de payload': row.payloadType, 'Statut': row.status, 'Champs modifiés': row.updatedFields,
      'Champs inchangés': row.unchangedFields, 'Champs en erreur': row.failedFields, 'Ancienne valeur': row.before,
      'Nouvelle valeur': row.after, 'Enregistrement effectué': row.saved, 'Message d’erreur': row.errorMessage
    } : {
      'Date de mise à jour': row.dateMiseAJour, 'ID parcours Drupal': row.drupalParcoursId, 'Titre Drupal': row.drupalTitle,
      'Dispositif': row.dispositifCode, 'Module': row.moduleCode, 'Nombre de places restantes': row.placesRestantes,
      'Nombre inscrits': row.preinscritsTotal, 'Nombre maximal': row.preinscritsMax, 'Lien de préinscription': row.lienPreinscription,
      'Date début saisie': row.dateDebutSaisie, 'Date fin saisie': row.dateFinSaisie,
      'Date dépublication saisie': row.dateDepublicationSaisie, 'Sessions saisies': row.sessionsSaisies
    });

    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, sheet, 'Bilan mises à jour');
    const content = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    downloadFile(`eafc-bilan-mises-a-jour-${stamp}.xlsx`, content, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  }

  /***************************************************************************
   * LOGS
   ***************************************************************************/
  function log(level, item, message) {
    const entry = {
      timestamp: new Date().toISOString(),
      level,
      drupalParcoursId: item?.drupalParcoursId || '',
      drupalTitle: item?.drupalTitle || '',
      dispositifCode: item?.dispositifCode || '',
      moduleCode: item?.moduleCode || '',
      message
    };
    state.logs.push(entry);
    console[level === 'error' ? 'error' : level === 'warning' ? 'warn' : 'log']('[EAFC]', entry);
    renderLogs();
  }

  function renderLogs() {
    const box = qs('#eafc-logbox');
    if (!box) return;
    box.innerHTML = state.logs.slice(-250).map((entry) => `<div class="eafc-log-${entry.level}">[${entry.timestamp}] ${entry.level.toUpperCase()} ${entry.drupalParcoursId ? `#${escapeHtml(entry.drupalParcoursId)} ` : ''}${escapeHtml(entry.message)}</div>`).join('');
    box.scrollTop = box.scrollHeight;
    state.lastSummary.errors = state.logs.filter((entry) => entry.level === 'error').length + state.validationErrors.length;
    renderSummary();
  }

  function exportLogs() {
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    const json = JSON.stringify(state.logs, null, 2);
    const csv = toCsv(state.logs);
    downloadFile(`eafc-drupal-logs-${stamp}.json`, json, 'application/json');
    downloadFile(`eafc-drupal-logs-${stamp}.csv`, csv, 'text/csv;charset=utf-8');
  }

  function toCsv(rows) {
    const headers = ['timestamp', 'level', 'drupalParcoursId', 'drupalTitle', 'dispositifCode', 'moduleCode', 'message'];
    const escapeCsv = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`;
    return [headers.join(','), ...rows.map((row) => headers.map((h) => escapeCsv(row[h])).join(','))].join('\n');
  }

  /***************************************************************************
   * UTILITAIRES DOM / ACTIONS
   ***************************************************************************/
  function qs(selector, scope = document) { return scope.querySelector(selector); }

  function findElementBySelectors(selectors, scope = document) {
    for (const selector of selectors) {
      if (selector.includes(':contains(')) continue;
      const found = Array.from(scope.querySelectorAll(selector)).find(isVisible) || scope.querySelector(selector);
      if (found) return found;
    }
    if (selectors.some((s) => s.includes(':contains("Modifier")'))) {
      return Array.from(scope.querySelectorAll('a,button,input[type="submit"]')).find((el) => /modifier/i.test(el.textContent || el.value || '') && isVisible(el));
    }
    return null;
  }

  async function waitForSelector(selector, timeout = CONFIG.MAX_WAIT_TIME) {
    return waitForFunction(() => document.querySelector(selector), timeout, `Sélecteur introuvable : ${selector}`);
  }

  async function waitForAnySelector(selectors, timeout = CONFIG.MAX_WAIT_TIME) {
    return waitForFunction(() => findElementBySelectors(selectors), timeout, `Aucun sélecteur trouvé : ${selectors.join(' | ')}`);
  }

  async function waitForFunction(fn, timeout, errorMessage) {
    const start = Date.now();
    while (Date.now() - start < timeout) {
      const result = fn();
      if (result) return result;
      await sleep(150);
    }
    throw new Error(errorMessage || 'Attente expirée.');
  }

  async function waitForLocation(path, timeout = CONFIG.MAX_WAIT_TIME) {
    await waitForFunction(() => location.pathname.startsWith(path), timeout, `Navigation vers ${path} non confirmée.`);
    await sleep(500);
  }


  async function waitForAjax(timeout = CONFIG.MAX_WAIT_TIME) {
    const start = Date.now();
    await sleep(300);
    while (Date.now() - start < timeout) {
      const ajaxActive = Boolean(document.querySelector('.ajax-progress, .ajax-progress-throbber, .ui-dialog .ajax-progress'));
      const documentBusy = document.body.getAttribute('aria-busy') === 'true' || document.documentElement.getAttribute('aria-busy') === 'true';
      if (!ajaxActive && !documentBusy) return;
      await sleep(150);
    }
  }

  async function clickElement(element, label = 'clic') {
    if (!element) throw new Error(`Élément introuvable pour ${label}.`);
    element.scrollIntoView({ block: 'center', inline: 'center' });
    highlightElement(element);
    if (CONFIG.SAFE_CLICK_MODE) {
      const descriptor = `${label}\n\nÉlément : ${describeElement(element)}\n\nValider ce clic ?`;
      if (!window.confirm(descriptor)) throw new Error(`Clic annulé par l’utilisateur : ${label}.`);
    }
    element.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
    element.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
    element.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
    element.click();
    await sleep(CONFIG.DELAY_BETWEEN_ACTIONS);
  }

  function setInputValue(element, value) {
    if (!element) throw new Error('Champ introuvable.');
    element.scrollIntoView({ block: 'center', inline: 'center' });
    element.focus();
    element.value = value == null ? '' : String(value);
    dispatchAllEvents(element);
  }

  function setSelectValue(select, value, fallbackLabel = '') {
    if (!select) throw new Error('Select introuvable.');
    const wanted = String(value || '');
    const option = Array.from(select.options || []).find((opt) => opt.value === wanted)
      || Array.from(select.options || []).find((opt) => normalizeText(opt.textContent) === normalizeText(fallbackLabel || wanted));
    if (!option) throw new Error(`Option select introuvable : ${wanted}.`);
    select.value = option.value;
    dispatchAllEvents(select);
  }

  function dispatchAllEvents(element) {
    ['input', 'change', 'blur'].forEach((type) => element.dispatchEvent(new Event(type, { bubbles: true, cancelable: true })));
  }

  function sleep(ms) { return new Promise((resolve) => setTimeout(resolve, ms)); }

  function normalizeText(text) {
    return String(text || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  function normalizeColumnName(text) {
    return normalizeText(text).replace(/[_\-:]/g, ' ').replace(/\s+/g, ' ').trim();
  }

  function normalizeRne(value) {
    const text = String(value ?? '').trim().toUpperCase();
    const match = text.match(/\b0?[0-9]{2,3}[0-9A-Z]{4,6}[A-Z]\b/);
    return match ? match[0].padStart(8, '0') : text;
  }

  function extractRneFromText(text) {
    const match = String(text || '').toUpperCase().match(/\[?\b(0?[0-9]{2,3}[0-9A-Z]{4,6}[A-Z])\b\]?/);
    return match ? normalizeRne(match[1]) : '';
  }

  function departmentFromRne(rne) {
    const normalized = normalizeRne(rne);
    return normalized ? normalizeTerritoireCode(normalized.slice(0, 3)) : '';
  }

  function normalizeTerritoireCode(value) {
    const text = String(value ?? '').trim().toUpperCase();
    if (!text) return '';
    const digits = text.match(/\d{2,3}/)?.[0] || text;
    return digits.length === 3 && digits.startsWith('0') ? digits.slice(1) : digits;
  }

  function cleanCode(value) {
    return String(value ?? '').trim().replace(/\.0$/, '');
  }

  function stringifyCell(value) {
    if (value == null) return '';
    if (typeof value === 'number' && Number.isFinite(value)) return String(value);
    return String(value).trim();
  }

  function escapeHtml(text) {
    return String(text ?? '').replace(/[&<>'"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[c]));
  }

  function downloadFile(filename, content, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    if (typeof GM_download === 'function') {
      GM_download({ url, name: filename, saveAs: true, ontimeout: () => URL.revokeObjectURL(url), onerror: () => fallbackDownload(url, filename), onload: () => setTimeout(() => URL.revokeObjectURL(url), 1000) });
    } else {
      fallbackDownload(url, filename);
    }
  }

  function fallbackDownload(url, filename) {
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function isVisible(element) {
    if (!element) return false;
    const style = getComputedStyle(element);
    const rect = element.getBoundingClientRect();
    return style.display !== 'none' && style.visibility !== 'hidden' && rect.width >= 0 && rect.height >= 0;
  }

  function isClickableVisible(element) {
    if (!element) return false;
    const style = getComputedStyle(element);
    const rect = element.getBoundingClientRect();
    return style.display !== 'none'
      && style.visibility !== 'hidden'
      && style.pointerEvents !== 'none'
      && rect.width > 0
      && rect.height > 0;
  }

  function highlightElement(element) {
    const oldOutline = element.style.outline;
    element.style.outline = '3px solid #ffb000';
    setTimeout(() => { element.style.outline = oldOutline; }, 1200);
  }

  function describeElement(element) {
    const text = normalizeText(element.textContent || element.value || '').slice(0, 100);
    return `<${element.tagName.toLowerCase()} id="${element.id || ''}" name="${element.getAttribute('name') || ''}" data-drupal-selector="${element.getAttribute('data-drupal-selector') || ''}"> ${text}`;
  }

  function captureTextContext() {
    return normalizeText(document.title + ' ' + location.href + ' ' + document.body.textContent).slice(0, 500);
  }

  function setStatus(message) {
    const el = qs('#eafc-status');
    if (el) el.textContent = message;
  }

  function setAlert(message) {
    const el = qs('#eafc-alert');
    if (el) el.textContent = message;
  }

  loadPersistedUiConfig();
  initPanel();
  loadPersistedImports();
  log('info', null, 'Script EAFC Drupal chargé. Mode simulation et sécurité clic activés par défaut.');
})();
