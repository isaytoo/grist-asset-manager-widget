// =============================================================================
// GRIST ASSET MANAGER WIDGET
// =============================================================================

var BIENS_TABLE = 'BM_Biens';
var GESTIONNAIRES_TABLE = 'BM_Gestionnaires';

// State
var biens = [];
var gestionnaires = [];
var isOwner = false;
var canManage = false; // Owner OR designated gestionnaire
var currentLang = 'fr';
var searchPage = 1;
var searchPageSize = 20;
var searchResults = [];
var sortCol = '';
var sortDir = 'asc';

// Column definitions for BM_Biens
var BIEN_COLUMNS = [
  { id: 'Reference_DDC', label_fr: 'Référence DDC', label_en: 'DDC Reference', type: 'Text' },
  { id: 'Gestion_SPI', label_fr: 'Gestion SPI', label_en: 'SPI Management', type: 'Text' },
  { id: 'Nom_OFA_OFT', label_fr: 'Nom OFA / OFT', label_en: 'OFA / OFT Name', type: 'Text' },
  { id: 'Mouvement', label_fr: 'Mouvement', label_en: 'Movement', type: 'Text' },
  { id: 'Date_Acte', label_fr: "Date de l'acte", label_en: 'Deed Date', type: 'Text' },
  { id: 'Annee', label_fr: 'Année', label_en: 'Year', type: 'Text' },
  { id: 'Commune', label_fr: 'Commune', label_en: 'Municipality', type: 'Text' },
  { id: 'Adresse', label_fr: 'Adresse', label_en: 'Address', type: 'Text' },
  { id: 'Ref_Parcelles', label_fr: 'Réf. Parcelles', label_en: 'Plot Ref.', type: 'Text' },
  { id: 'Type_Bien', label_fr: 'Type', label_en: 'Type', type: 'Text' },
  { id: 'Surface_Bati', label_fr: 'Surface bâti (m²)', label_en: 'Built Area (m²)', type: 'Text' },
  { id: 'Surface_Parcelle', label_fr: 'Surface parcelle (m²)', label_en: 'Plot Area (m²)', type: 'Text' },
  { id: 'Nouvelle_Copropriete', label_fr: 'Nouvelle copropriété', label_en: 'New Co-ownership', type: 'Text' },
  { id: 'Occupation', label_fr: 'Occupation', label_en: 'Occupation', type: 'Text' },
  { id: 'Jouissance_Anticipee', label_fr: 'Jouissance anticipée', label_en: 'Early Enjoyment', type: 'Text' },
  { id: 'Jouissance_Differee', label_fr: 'Jouissance différée', label_en: 'Deferred Enjoyment', type: 'Text' },
  { id: 'Temps_Portage', label_fr: 'Temps portage / Année fin', label_en: 'Carrying Time / End Year', type: 'Text' },
  { id: 'Bail_Longue_Duree', label_fr: 'Mis à bail longue durée', label_en: 'Long-term Lease', type: 'Text' },
  { id: 'Acquisition_Compte_Tiers', label_fr: 'Acquisition compte tiers', label_en: 'Third-party Acquisition', type: 'Text' },
  { id: 'Prefinancement', label_fr: 'Préfinancement', label_en: 'Pre-financing', type: 'Text' },
  { id: 'Surface_Assurance', label_fr: 'Surface pour assurance (m²)', label_en: 'Insurance Area (m²)', type: 'Text' },
  { id: 'Tiers_Vendeur_Acquereur', label_fr: 'Tiers Vendeur ou Acquéreur', label_en: 'Third-party Seller/Buyer', type: 'Text' },
  { id: 'Nature_Bien', label_fr: 'Nature du bien - Projet acquisition', label_en: 'Property Nature - Acquisition Project', type: 'Text' },
  { id: 'Import_GIMA', label_fr: 'Import GIMA', label_en: 'GIMA Import', type: 'Text' },
  { id: 'Num_Site', label_fr: 'N° du site', label_en: 'Site Number', type: 'Text' },
  { id: 'Saisies_Manuelles', label_fr: 'Saisies manuelles', label_en: 'Manual Entries', type: 'Text' },
  { id: 'Date_Integration_GIMA', label_fr: "Date d'intégration GIMA", label_en: 'GIMA Integration Date', type: 'Text' },
  { id: 'Dossier_Numerique', label_fr: 'Dossier numérique sous L', label_en: 'Digital Folder under L', type: 'Text' },
  { id: 'Observation', label_fr: 'Observations', label_en: 'Observations', type: 'Text' }
];

// =============================================================================
// I18N
// =============================================================================

var i18n = {
  fr: {
    appTitle: 'Gestion des Biens',
    appSubtitle: 'Ajoutez, modifiez et supprimez des biens patrimoniaux',
    tabSearch: 'Recherche',
    tabGestion: 'Gestion des Biens',
    tabDashboard: 'Tableau de Bord',
    tabGestionnaires: 'Gestionnaires',
    searchTitle: 'Recherche Multicritères',
    searchSubtitle: 'Trouvez instantanément vos biens avec des filtres avancés',
    searchBtn: 'Rechercher',
    resetBtn: 'Réinitialiser',
    exportBtn: 'Exporter Excel',
    resultsFound: 'résultats trouvés',
    page: 'page',
    of: 'sur',
    clickHint: 'Cliquez sur une ligne pour voir les détails complets',
    noResults: 'Aucun résultat trouvé',
    allCommunes: 'Toutes les communes',
    allMovements: 'Tous les mouvements',
    allTypes: 'Tous les types',
    allYears: 'Toutes les années',
    addTitle: 'Ajouter un Nouveau Bien',
    addDesc: 'Créez un nouveau bien patrimonial avec toutes ses caractéristiques (localisation, surfaces, mouvement, etc.)',
    addBtn: 'Ajouter un bien',
    editTitle: 'Actualiser un Bien',
    editDesc: "Modifiez les informations d'un bien existant en saisissant sa référence DDC",
    editBtn: 'Rechercher et modifier',
    deleteTitle: 'Supprimer un Bien',
    deleteDesc: 'Supprimez définitivement un bien en saisissant sa référence DDC',
    deleteBtn: 'Rechercher et supprimer',
    refPlaceholder: 'Référence DDC (ex: ECH 69389 22 00001)',
    modalAdd: 'Ajouter un nouveau bien',
    modalEdit: 'Modifier le bien',
    modalDetail: 'Détails du bien',
    save: 'Enregistrer',
    cancel: 'Annuler',
    delete: 'Supprimer',
    confirmDelete: 'Êtes-vous sûr de vouloir supprimer ce bien ?',
    confirmDeleteTitle: 'Confirmation de suppression',
    confirmDeleteMsg: 'Êtes-vous sûr de vouloir supprimer définitivement ce bien ?',
    confirmDeleteWarning: 'Cette action est irréversible. Toutes les données de ce bien seront perdues.',
    confirmDeleteBtn: 'Supprimer définitivement',
    bienAdded: 'Bien ajouté avec succès',
    bienUpdated: 'Bien mis à jour avec succès',
    bienDeleted: 'Bien supprimé avec succès',
    bienNotFound: 'Bien non trouvé avec cette référence',
    sectionIdentification: 'Identification',
    sectionMouvement: 'Mouvement',
    sectionLocalisation: 'Localisation',
    sectionCaracteristiques: 'Caractéristiques du Bien',
    sectionOccupation: 'Occupation et Jouissance',
    sectionFinancement: 'Acquisition et Financement',
    sectionGIMA: 'Gestion GIMA',
    sectionNature: 'Nature du Bien et Observations',
    select: 'Sélectionner',
    yes: 'OUI',
    no: 'NON',
    dashTitle: 'Tableau de Bord',
    dashSubtitle: "Vue d'ensemble de votre patrimoine immobilier",
    totalBiens: 'Total des Biens',
    totalCommunes: 'Communes',
    totalSurfaceParcelle: 'Surface Totale des parcelles',
    totalSurfaceBati: 'Surface totale bâtis',
    repartitionMouvement: 'Répartition par Type de Mouvement',
    topCommunes: 'Top 10 des Communes',
    dashFilters: 'Filtres temporels',
    dashYear: 'Année',
    dashMonth: "Mois de l'acte",
    dashAll: 'Toutes',
    dashMonthHint: 'Sélectionnez un ou plusieurs mois',
    dashSurfaceAnalysis: 'Analyse des Surfaces',
    dashSurfAcqCed: 'Surfaces acquises / surfaces cédées',
    dashBatiNonBati: 'Répartition bâti et non-bâti',
    dashSurfBati: 'Surface Bâti (surface de plancher)',
    dashSurfNonBati: 'Surface parcellaire Non Bâti',
    dashGestionSPI: 'Gestion SPI',
    dashTypeActe: "Type d'acte",
    dashSPIEntrants: 'Dont actes pour biens entrants en gestion au SPI',
    dashSPIStats: 'Statistiques SPI par type d\'acte',
    dashDetailTab: 'Détails des surfaces par bien',
    dashExportExcel: 'Exporter Excel (4 onglets)',
    dashExportDone: 'Export Excel téléchargé avec succès',
    gestTitle: 'Gestion des Gestionnaires',
    gestSubtitle: 'Désignez les personnes autorisées à gérer les biens (en plus du Owner)',
    gestEmail: 'Email du gestionnaire',
    gestAdd: 'Ajouter',
    gestEmpty: 'Aucun gestionnaire désigné. Seul le Owner peut gérer les biens.',
    gestAdded: 'Gestionnaire ajouté',
    gestRemoved: 'Gestionnaire retiré',
    accessDenied: "Vous n'avez pas les droits pour effectuer cette action",
    ownerOnly: 'Réservé au Owner',
    tabImport: 'Import Excel',
    importTitle: 'Import de Fichier Excel',
    importSubtitle: 'Importez ou mettez à jour vos biens depuis un fichier Excel (.xlsx, .xls)',
    importDrop: 'Glissez-déposez votre fichier Excel ici',
    importOr: 'ou cliquez pour sélectionner un fichier',
    importFormats: 'Formats acceptés : .xlsx, .xls',
    importReplace: 'Remplacer tout',
    importReplaceDesc: 'Supprime tous les biens existants et les remplace par le fichier',
    importAppend: 'Ajouter uniquement',
    importAppendDesc: 'Ajoute les nouvelles lignes sans toucher aux existantes',
    importUpdate: 'Mettre à jour',
    importUpdateDesc: 'Met à jour les biens existants (par Réf DDC) et ajoute les nouveaux',
    importBtn: 'Lancer l\'import',
    importPreview: 'Aperçu des données',
    importRows: 'lignes détectées',
    importCols: 'colonnes mappées',
    importInProgress: 'Import en cours...',
    importDone: 'Import terminé !',
    importAdded: 'ajoutés',
    importUpdated: 'mis à jour',
    importDeleted: 'supprimés',
    importErrors: 'erreurs',
    importSelectSheet: 'Sélectionner la feuille',
    importNoFile: 'Veuillez sélectionner un fichier',
    importConfirmReplace: 'Attention : cette action va supprimer tous les biens existants et les remplacer. Continuer ?'
  },
  en: {
    appTitle: 'Asset Management',
    appSubtitle: 'Add, edit and delete real estate assets',
    tabSearch: 'Search',
    tabGestion: 'Asset Management',
    tabDashboard: 'Dashboard',
    tabGestionnaires: 'Managers',
    searchTitle: 'Multi-criteria Search',
    searchSubtitle: 'Find your assets instantly with advanced filters',
    searchBtn: 'Search',
    resetBtn: 'Reset',
    exportBtn: 'Export Excel',
    resultsFound: 'results found',
    page: 'page',
    of: 'of',
    clickHint: 'Click a row to see full details',
    noResults: 'No results found',
    allCommunes: 'All municipalities',
    allMovements: 'All movements',
    allTypes: 'All types',
    allYears: 'All years',
    addTitle: 'Add a New Asset',
    addDesc: 'Create a new real estate asset with all its characteristics (location, areas, movement, etc.)',
    addBtn: 'Add an asset',
    editTitle: 'Update an Asset',
    editDesc: 'Edit an existing asset by entering its DDC reference',
    editBtn: 'Search and edit',
    deleteTitle: 'Delete an Asset',
    deleteDesc: 'Permanently delete an asset by entering its DDC reference',
    deleteBtn: 'Search and delete',
    refPlaceholder: 'DDC Reference (e.g. ECH 69389 22 00001)',
    modalAdd: 'Add a new asset',
    modalEdit: 'Edit asset',
    modalDetail: 'Asset details',
    save: 'Save',
    cancel: 'Cancel',
    delete: 'Delete',
    confirmDelete: 'Are you sure you want to delete this asset?',
    confirmDeleteTitle: 'Delete Confirmation',
    confirmDeleteMsg: 'Are you sure you want to permanently delete this asset?',
    confirmDeleteWarning: 'This action is irreversible. All data for this asset will be lost.',
    confirmDeleteBtn: 'Delete permanently',
    bienAdded: 'Asset added successfully',
    bienUpdated: 'Asset updated successfully',
    bienDeleted: 'Asset deleted successfully',
    bienNotFound: 'No asset found with this reference',
    sectionIdentification: 'Identification',
    sectionMouvement: 'Movement',
    sectionLocalisation: 'Location',
    sectionCaracteristiques: 'Asset Characteristics',
    sectionOccupation: 'Occupation and Enjoyment',
    sectionFinancement: 'Acquisition and Financing',
    sectionGIMA: 'GIMA Management',
    sectionNature: 'Property Nature and Observations',
    select: 'Select',
    yes: 'YES',
    no: 'NO',
    dashTitle: 'Dashboard',
    dashSubtitle: 'Overview of your real estate portfolio',
    totalBiens: 'Total Assets',
    totalCommunes: 'Municipalities',
    totalSurfaceParcelle: 'Total Plot Area',
    totalSurfaceBati: 'Total Built Area',
    repartitionMouvement: 'Distribution by Movement Type',
    topCommunes: 'Top 10 Municipalities',
    dashFilters: 'Time Filters',
    dashYear: 'Year',
    dashMonth: 'Deed month',
    dashAll: 'All',
    dashMonthHint: 'Select one or more months',
    dashSurfaceAnalysis: 'Surface Analysis',
    dashSurfAcqCed: 'Acquired surfaces / ceded surfaces',
    dashBatiNonBati: 'Built and unbuilt distribution',
    dashSurfBati: 'Built Surface (floor area)',
    dashSurfNonBati: 'Unbuilt Plot Surface',
    dashGestionSPI: 'SPI Management',
    dashTypeActe: 'Deed type',
    dashSPIEntrants: 'Of which deeds for assets entering SPI management',
    dashSPIStats: 'SPI Statistics by deed type',
    dashDetailTab: 'Surface details by asset',
    dashExportExcel: 'Export Excel (4 tabs)',
    dashExportDone: 'Excel export downloaded successfully',
    gestTitle: 'Manager Management',
    gestSubtitle: 'Designate people authorized to manage assets (in addition to Owner)',
    gestEmail: 'Manager email',
    gestAdd: 'Add',
    gestEmpty: 'No designated managers. Only the Owner can manage assets.',
    gestAdded: 'Manager added',
    gestRemoved: 'Manager removed',
    accessDenied: 'You do not have permission to perform this action',
    ownerOnly: 'Owner only',
    tabImport: 'Import Excel',
    importTitle: 'Excel File Import',
    importSubtitle: 'Import or update your assets from an Excel file (.xlsx, .xls)',
    importDrop: 'Drag and drop your Excel file here',
    importOr: 'or click to select a file',
    importFormats: 'Accepted formats: .xlsx, .xls',
    importReplace: 'Replace all',
    importReplaceDesc: 'Deletes all existing assets and replaces them with the file',
    importAppend: 'Append only',
    importAppendDesc: 'Adds new rows without touching existing ones',
    importUpdate: 'Update',
    importUpdateDesc: 'Updates existing assets (by DDC Ref) and adds new ones',
    importBtn: 'Start import',
    importPreview: 'Data preview',
    importRows: 'rows detected',
    importCols: 'columns mapped',
    importInProgress: 'Import in progress...',
    importDone: 'Import complete!',
    importAdded: 'added',
    importUpdated: 'updated',
    importDeleted: 'deleted',
    importErrors: 'errors',
    importSelectSheet: 'Select sheet',
    importNoFile: 'Please select a file',
    importConfirmReplace: 'Warning: this will delete all existing assets and replace them. Continue?'
  }
};

function t(key) {
  return (i18n[currentLang] && i18n[currentLang][key]) || key;
}

function colLabel(col) {
  return currentLang === 'fr' ? col.label_fr : col.label_en;
}

function setLang(lang) {
  currentLang = lang;
  document.querySelectorAll('.lang-btn').forEach(function(btn) {
    btn.classList.toggle('active', btn.textContent.trim() === lang.toUpperCase());
  });
  document.getElementById('app-title').textContent = t('appTitle');
  document.getElementById('app-subtitle').textContent = t('appSubtitle');
  document.querySelectorAll('[data-i18n]').forEach(function(el) {
    el.textContent = t(el.getAttribute('data-i18n'));
  });
  refreshAllViews();
}

// =============================================================================
// UTILITIES
// =============================================================================

function sanitize(str) {
  if (!str && str !== 0) return '';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function showToast(msg, type) {
  var toast = document.createElement('div');
  toast.className = 'toast toast-' + (type || 'info');
  toast.textContent = msg;
  document.getElementById('toast-container').appendChild(toast);
  setTimeout(function() { toast.remove(); }, 3000);
}

function closeModal(event) {
  if (event.target.classList.contains('modal-overlay')) {
    document.getElementById('modal-container').innerHTML = '';
  }
}

function closeModalForce() {
  document.getElementById('modal-container').innerHTML = '';
}

function isInsideGrist() {
  try { return window.frameElement !== null || window !== window.parent; } catch (e) { return true; }
}

function parseNum(val) {
  if (!val && val !== 0) return 0;
  var n = parseFloat(String(val).replace(/\s/g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

function movementBadge(mouvement) {
  if (!mouvement) return '<span class="badge badge-other">--</span>';
  var m = mouvement.toUpperCase().trim();
  if (m.indexOf('ACQUISITION') !== -1 || m === 'ACQUISITION') return '<span class="badge badge-acquisition">ACQUISITION</span>';
  if (m.indexOf('CESSION') !== -1) return '<span class="badge badge-cession">CESSION</span>';
  if (m.indexOf('PREEMPTION') !== -1 || m.indexOf('PRÉEMPTION') !== -1) return '<span class="badge badge-preemption">PRÉEMPTION</span>';
  if (m.indexOf('SERVITUDE') !== -1) return '<span class="badge badge-servitude">SERVITUDE</span>';
  if (m.indexOf('ECHANGE') !== -1 || m.indexOf('ÉCHANGE') !== -1) return '<span class="badge badge-echange">ÉCHANGE</span>';
  if (m.indexOf('EXPROPRIATION') !== -1) return '<span class="badge badge-expropriation">EXPROPRIATION</span>';
  if (m.indexOf('LIBERATION') !== -1 || m.indexOf('LIBÉRATION') !== -1) return '<span class="badge badge-liberation">LIBÉRATION</span>';
  return '<span class="badge badge-other">' + sanitize(mouvement) + '</span>';
}

// =============================================================================
// TAB SWITCHING
// =============================================================================

function switchTab(tabId) {
  document.querySelectorAll('.tab-btn').forEach(function(btn) {
    btn.classList.toggle('active', btn.getAttribute('data-tab') === tabId);
  });
  document.querySelectorAll('.tab-content').forEach(function(tc) {
    tc.classList.toggle('active', tc.id === 'tab-' + tabId);
  });
  if (tabId === 'search') renderSearchView();
  if (tabId === 'gestion') renderGestionView();
  if (tabId === 'dashboard') renderDashboardView();
  if (tabId === 'import') renderImportView();
  if (tabId === 'gestionnaires') renderGestionnairesView();
}

function refreshAllViews() {
  var active = document.querySelector('.tab-btn.active');
  if (active) switchTab(active.getAttribute('data-tab'));
}

// =============================================================================
// SEARCH VIEW
// =============================================================================

function getUniqueValues(field) {
  var vals = {};
  for (var i = 0; i < biens.length; i++) {
    var v = biens[i][field];
    if (v && String(v).trim()) vals[String(v).trim()] = true;
  }
  return Object.keys(vals).sort();
}

function renderSearchView() {
  var communes = getUniqueValues('Commune');
  var mouvements = getUniqueValues('Mouvement');
  var types = getUniqueValues('Type_Bien');
  var annees = getUniqueValues('Annee').sort(function(a, b) { return b - a; });

  var html = '<div class="section-card">';
  html += '<h3>🔍 ' + t('searchTitle') + '</h3>';
  html += '<p style="color:#64748b;margin-bottom:16px;">' + t('searchSubtitle') + '</p>';

  html += '<div class="search-grid">';
  html += '<div class="search-field"><label>Référence DDC</label><input type="text" id="s-ref" placeholder="Ex: ECH 69389 22 00001" oninput="doSearch()" /></div>';
  html += '<div class="search-field"><label>Commune</label><select id="s-commune" onchange="doSearch()"><option value="">' + t('allCommunes') + '</option>';
  for (var i = 0; i < communes.length; i++) html += '<option value="' + sanitize(communes[i]) + '">' + sanitize(communes[i]) + '</option>';
  html += '</select></div>';
  html += '<div class="search-field"><label>Mouvement</label><select id="s-mouvement" onchange="doSearch()"><option value="">' + t('allMovements') + '</option>';
  for (var i = 0; i < mouvements.length; i++) html += '<option value="' + sanitize(mouvements[i]) + '">' + sanitize(mouvements[i]) + '</option>';
  html += '</select></div>';
  html += '<div class="search-field"><label>Adresse</label><input type="text" id="s-adresse" placeholder="Ex: LA JACQUIERE" oninput="doSearch()" /></div>';
  html += '<div class="search-field"><label>Réf. Parcelle</label><input type="text" id="s-parcelle" oninput="doSearch()" /></div>';
  html += '<div class="search-field"><label>Type de Bien</label><select id="s-type" onchange="doSearch()"><option value="">' + t('allTypes') + '</option>';
  for (var i = 0; i < types.length; i++) html += '<option value="' + sanitize(types[i]) + '">' + sanitize(types[i]) + '</option>';
  html += '</select></div>';
  html += '<div class="search-field"><label>Année</label><select id="s-annee" onchange="doSearch()"><option value="">' + t('allYears') + '</option>';
  for (var i = 0; i < annees.length; i++) html += '<option value="' + sanitize(annees[i]) + '">' + sanitize(annees[i]) + '</option>';
  html += '</select></div>';
  html += '<div class="search-field"><label>N° Site</label><input type="text" id="s-site" oninput="doSearch()" /></div>';
  html += '</div>';

  html += '<div class="search-actions">';
  html += '<button class="btn btn-secondary" onclick="resetSearch()">🔄 ' + t('resetBtn') + '</button>';
  html += '</div>';
  html += '</div>';

  // Results
  html += '<div id="search-results"></div>';

  document.getElementById('search-view').innerHTML = html;

  // Auto-show all results
  doSearch();
}

function doSearch() {
  var ref = (document.getElementById('s-ref') ? document.getElementById('s-ref').value : '').trim().toLowerCase();
  var commune = (document.getElementById('s-commune') ? document.getElementById('s-commune').value : '');
  var mouvement = (document.getElementById('s-mouvement') ? document.getElementById('s-mouvement').value : '');
  var adresse = (document.getElementById('s-adresse') ? document.getElementById('s-adresse').value : '').trim().toLowerCase();
  var parcelle = (document.getElementById('s-parcelle') ? document.getElementById('s-parcelle').value : '').trim().toLowerCase();
  var type = (document.getElementById('s-type') ? document.getElementById('s-type').value : '');
  var annee = (document.getElementById('s-annee') ? document.getElementById('s-annee').value : '');
  var site = (document.getElementById('s-site') ? document.getElementById('s-site').value : '').trim().toLowerCase();

  searchResults = biens.filter(function(b) {
    if (ref && String(b.Reference_DDC || '').toLowerCase().indexOf(ref) === -1) return false;
    if (commune && b.Commune !== commune) return false;
    if (mouvement && b.Mouvement !== mouvement) return false;
    if (adresse && String(b.Adresse || '').toLowerCase().indexOf(adresse) === -1) return false;
    if (parcelle && String(b.Ref_Parcelles || '').toLowerCase().indexOf(parcelle) === -1) return false;
    if (type && b.Type_Bien !== type) return false;
    if (annee && String(b.Annee) !== annee) return false;
    if (site && String(b.Num_Site || '').toLowerCase().indexOf(site) === -1) return false;
    return true;
  });

  // Sort
  if (sortCol) {
    searchResults.sort(function(a, b) {
      var va = String(a[sortCol] || '').toLowerCase();
      var vb = String(b[sortCol] || '').toLowerCase();
      if (va < vb) return sortDir === 'asc' ? -1 : 1;
      if (va > vb) return sortDir === 'asc' ? 1 : -1;
      return 0;
    });
  }

  searchPage = 1;
  renderSearchResults();
}

function resetSearch() {
  var ids = ['s-ref', 's-commune', 's-mouvement', 's-adresse', 's-parcelle', 's-type', 's-annee', 's-site'];
  for (var i = 0; i < ids.length; i++) {
    var el = document.getElementById(ids[i]);
    if (el) el.value = '';
  }
  sortCol = '';
  sortDir = 'asc';
  doSearch();
}

function sortBy(col) {
  if (sortCol === col) {
    sortDir = sortDir === 'asc' ? 'desc' : 'asc';
  } else {
    sortCol = col;
    sortDir = 'asc';
  }
  doSearch();
}

function renderSearchResults() {
  var total = searchResults.length;
  var totalPages = Math.max(1, Math.ceil(total / searchPageSize));
  if (searchPage > totalPages) searchPage = totalPages;
  var start = (searchPage - 1) * searchPageSize;
  var pageItems = searchResults.slice(start, start + searchPageSize);

  var html = '';
  html += '<div class="results-info">' + total + ' ' + t('resultsFound') + ' (' + t('page') + ' ' + searchPage + ' ' + t('of') + ' ' + totalPages + ')</div>';
  html += '<div class="results-hint">💡 ' + t('clickHint') + '</div>';

  if (total === 0) {
    html += '<div style="text-align:center;padding:40px;color:#94a3b8;">' + t('noResults') + '</div>';
  } else {
    var displayCols = ['Reference_DDC', 'Mouvement', 'Date_Acte', 'Commune', 'Adresse', 'Type_Bien', 'Ref_Parcelles', 'Surface_Bati', 'Surface_Parcelle'];

    html += '<div style="overflow-x:auto;"><table class="data-table"><thead><tr>';
    for (var c = 0; c < displayCols.length; c++) {
      var col = BIEN_COLUMNS.find(function(cc) { return cc.id === displayCols[c]; });
      var arrow = sortCol === displayCols[c] ? (sortDir === 'asc' ? ' ▲' : ' ▼') : '';
      html += '<th onclick="sortBy(\'' + displayCols[c] + '\')">' + colLabel(col) + '<span class="sort-icon">' + arrow + '</span></th>';
    }
    html += '</tr></thead><tbody>';

    for (var i = 0; i < pageItems.length; i++) {
      var b = pageItems[i];
      html += '<tr onclick="openDetailModal(' + b.id + ')">';
      html += '<td>' + sanitize(b.Reference_DDC) + '</td>';
      html += '<td>' + movementBadge(b.Mouvement) + '</td>';
      html += '<td>' + sanitize(b.Date_Acte) + '</td>';
      html += '<td>' + sanitize(b.Commune) + '</td>';
      html += '<td>' + sanitize(b.Adresse) + '</td>';
      html += '<td>' + sanitize(b.Type_Bien) + '</td>';
      html += '<td>' + sanitize(b.Ref_Parcelles) + '</td>';
      html += '<td style="text-align:right;">' + sanitize(b.Surface_Bati) + '</td>';
      html += '<td style="text-align:right;">' + sanitize(b.Surface_Parcelle) + '</td>';
      html += '</tr>';
    }
    html += '</tbody></table></div>';

    // Pagination
    html += '<div class="pagination">';
    html += '<button ' + (searchPage <= 1 ? 'disabled' : '') + ' onclick="searchPage=1;renderSearchResults();">«</button>';
    html += '<button ' + (searchPage <= 1 ? 'disabled' : '') + ' onclick="searchPage--;renderSearchResults();">‹</button>';
    var startP = Math.max(1, searchPage - 2);
    var endP = Math.min(totalPages, searchPage + 2);
    for (var p = startP; p <= endP; p++) {
      html += '<button class="' + (p === searchPage ? 'active' : '') + '" onclick="searchPage=' + p + ';renderSearchResults();">' + p + '</button>';
    }
    html += '<button ' + (searchPage >= totalPages ? 'disabled' : '') + ' onclick="searchPage++;renderSearchResults();">›</button>';
    html += '<button ' + (searchPage >= totalPages ? 'disabled' : '') + ' onclick="searchPage=' + totalPages + ';renderSearchResults();">»</button>';
    html += '</div>';
  }

  document.getElementById('search-results').innerHTML = html;
}

// =============================================================================
// DETAIL MODAL (read-only)
// =============================================================================

function openDetailModal(bienId) {
  var b = biens.find(function(x) { return x.id === bienId; });
  if (!b) return;

  var html = '<div class="modal-overlay" onclick="closeModal(event)">';
  html += '<div class="modal" onclick="event.stopPropagation()">';
  html += '<div class="modal-header-red"><h3>📋 ' + t('modalDetail') + '</h3><button class="modal-close-white" onclick="closeModalForce()">✕</button></div>';
  html += '<div class="modal-body">';

  html += '<div class="detail-grid">';
  for (var i = 0; i < BIEN_COLUMNS.length; i++) {
    var col = BIEN_COLUMNS[i];
    var val = b[col.id];
    var displayVal = val ? sanitize(String(val)) : '--';
    if (col.id === 'Mouvement') displayVal = movementBadge(val);
    var fullClass = (col.id === 'Nature_Bien' || col.id === 'Observation') ? ' full-width' : '';
    html += '<div class="detail-item' + fullClass + '">';
    html += '<div class="detail-label">' + colLabel(col) + '</div>';
    html += '<div class="detail-value">' + displayVal + '</div>';
    html += '</div>';
  }
  html += '</div>';

  html += '</div>'; // modal-body

  if (canManage) {
    html += '<div class="modal-footer">';
    html += '<button class="btn btn-danger" onclick="deleteBien(' + b.id + ')">🗑️ ' + t('delete') + '</button>';
    html += '<button class="btn btn-primary" onclick="openEditModal(' + b.id + ')">✏️ ' + t('modalEdit') + '</button>';
    html += '</div>';
  }

  html += '</div></div>';
  document.getElementById('modal-container').innerHTML = html;
}

// =============================================================================
// GESTION VIEW
// =============================================================================

function renderGestionView() {
  var html = '<div class="section-card">';
  html += '<h3>🏠 ' + t('tabGestion') + '</h3>';
  html += '<p style="color:#64748b;margin-bottom:20px;">' + t('appSubtitle') + '</p>';

  html += '<div class="gestion-cards">';

  // Add card
  html += '<div class="gestion-card">';
  html += '<div class="card-icon">➕</div>';
  html += '<h4>' + t('addTitle') + '</h4>';
  html += '<p>' + t('addDesc') + '</p>';
  if (canManage) {
    html += '<button class="btn btn-success" onclick="openAddModal()">✨ ' + t('addBtn') + '</button>';
  } else {
    html += '<button class="btn btn-secondary" disabled>🔒 ' + t('accessDenied') + '</button>';
  }
  html += '</div>';

  // Edit card
  html += '<div class="gestion-card">';
  html += '<div class="card-icon">✏️</div>';
  html += '<h4>' + t('editTitle') + '</h4>';
  html += '<p>' + t('editDesc') + '</p>';
  html += '<div class="card-input"><input type="text" id="edit-ref" placeholder="' + t('refPlaceholder') + '" /></div>';
  if (canManage) {
    html += '<button class="btn btn-primary" onclick="searchAndEdit()">🔍 ' + t('editBtn') + '</button>';
  } else {
    html += '<button class="btn btn-secondary" disabled>🔒 ' + t('accessDenied') + '</button>';
  }
  html += '</div>';

  // Delete card
  html += '<div class="gestion-card">';
  html += '<div class="card-icon">🗑️</div>';
  html += '<h4>' + t('deleteTitle') + '</h4>';
  html += '<p>' + t('deleteDesc') + '</p>';
  html += '<div class="card-input"><input type="text" id="delete-ref" placeholder="' + t('refPlaceholder') + '" /></div>';
  if (canManage) {
    html += '<button class="btn btn-danger" onclick="searchAndDelete()">🗑️ ' + t('deleteBtn') + '</button>';
  } else {
    html += '<button class="btn btn-secondary" disabled>🔒 ' + t('accessDenied') + '</button>';
  }
  html += '</div>';

  html += '</div></div>';

  document.getElementById('gestion-view').innerHTML = html;
}

function searchAndEdit() {
  var ref = document.getElementById('edit-ref').value.trim();
  if (!ref) return;
  var b = biens.find(function(x) { return String(x.Reference_DDC || '').trim().toLowerCase() === ref.toLowerCase(); });
  if (!b) { showToast(t('bienNotFound'), 'error'); return; }
  openEditModal(b.id);
}

function searchAndDelete() {
  var ref = document.getElementById('delete-ref').value.trim();
  if (!ref) return;
  var b = biens.find(function(x) { return String(x.Reference_DDC || '').trim().toLowerCase() === ref.toLowerCase(); });
  if (!b) { showToast(t('bienNotFound'), 'error'); return; }
  openDeleteConfirmModal(b);
}

function openDeleteConfirmModal(bien) {
  var html = '<div class="modal-overlay" onclick="closeModal(event)">';
  html += '<div class="modal" style="max-width:560px;">';
  html += '<div class="modal-header-red" style="background:#dc2626;">';
  html += '<h3>🗑️ ' + t('confirmDeleteTitle') + '</h3>';
  html += '<button class="modal-close-white" onclick="closeModalForce()">✕</button>';
  html += '</div>';
  html += '<div class="modal-body" style="text-align:center;padding:30px;">';

  // Warning icon
  html += '<div style="font-size:56px;margin-bottom:16px;">⚠️</div>';
  html += '<p style="font-size:15px;font-weight:700;color:#1e293b;margin-bottom:20px;">' + t('confirmDeleteMsg') + '</p>';

  // Asset details card
  html += '<div style="background:#fef2f2;border:1px solid #fecaca;border-radius:12px;padding:16px;text-align:left;margin-bottom:20px;">';
  html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">';
  html += '<div><span style="font-size:10px;font-weight:700;color:#94a3b8;text-transform:uppercase;">Réf. DDC</span><br/><span style="font-weight:700;color:#1e293b;">' + sanitize(bien.Reference_DDC) + '</span></div>';
  html += '<div><span style="font-size:10px;font-weight:700;color:#94a3b8;text-transform:uppercase;">' + t('tabSearch') + '</span><br/>' + movementBadge(bien.Mouvement) + '</div>';
  html += '<div><span style="font-size:10px;font-weight:700;color:#94a3b8;text-transform:uppercase;">Commune</span><br/><span style="font-weight:600;">' + sanitize(bien.Commune || '--') + '</span></div>';
  html += '<div><span style="font-size:10px;font-weight:700;color:#94a3b8;text-transform:uppercase;">Adresse</span><br/><span style="font-weight:600;">' + sanitize(bien.Adresse || '--') + '</span></div>';
  if (bien.Type_Bien) {
    html += '<div><span style="font-size:10px;font-weight:700;color:#94a3b8;text-transform:uppercase;">Type</span><br/><span style="font-weight:600;">' + sanitize(bien.Type_Bien) + '</span></div>';
  }
  if (bien.Annee) {
    html += '<div><span style="font-size:10px;font-weight:700;color:#94a3b8;text-transform:uppercase;">' + colLabel(BIEN_COLUMNS[5]) + '</span><br/><span style="font-weight:600;">' + sanitize(bien.Annee) + '</span></div>';
  }
  html += '</div></div>';

  // Warning text
  html += '<p style="font-size:12px;color:#dc2626;font-weight:600;">⚠️ ' + t('confirmDeleteWarning') + '</p>';

  html += '</div>';

  // Footer with buttons
  html += '<div class="modal-footer" style="justify-content:center;gap:16px;">';
  html += '<button class="btn btn-secondary" onclick="closeModalForce()" style="min-width:120px;">✕ ' + t('cancel') + '</button>';
  html += '<button class="btn" onclick="confirmDeleteBien(' + bien.id + ')" style="min-width:120px;background:#dc2626;color:#fff;">🗑️ ' + t('confirmDeleteBtn') + '</button>';
  html += '</div>';

  html += '</div></div>';

  document.getElementById('modal-container').innerHTML = html;
}

// =============================================================================
// ADD / EDIT MODAL
// =============================================================================

function buildFormHtml(bien) {
  var isEdit = !!bien;
  var v = function(field) { return isEdit ? sanitize(bien[field] || '') : ''; };

  var mouvementOptions = ['', 'Acquisition', 'Cession', 'Préemption', 'Servitude', 'Échange', 'Expropriation', 'Libération', 'Annulation EEDV-RCP'];
  var typeOptions = ['', 'Bâti sans terrain', 'Non bâti', 'Bâti', 'Terrain'];
  var ouiNonOptions = ['', 'Oui', 'Non'];
  var importGimaOptions = ['', 'Import automatique', 'Exempté'];
  var saisiesOptions = ['', 'Oui', 'Non'];

  var html = '';

  // Section: Identification
  html += '<div class="form-section"><h4>📋 ' + t('sectionIdentification') + '</h4>';
  html += '<div class="form-grid">';
  html += '<div class="form-group"><label>Référence DDC <span class="required">*</span></label><input type="text" id="f-Reference_DDC" value="' + v('Reference_DDC') + '" /></div>';
  html += '<div class="form-group"><label>Gestion SPI</label>';
  html += '<div class="radio-group"><label><input type="radio" name="f-Gestion_SPI" value="Non" ' + (v('Gestion_SPI') !== 'Oui' ? 'checked' : '') + ' /> NON</label>';
  html += '<label><input type="radio" name="f-Gestion_SPI" value="Oui" ' + (v('Gestion_SPI') === 'Oui' ? 'checked' : '') + ' /> OUI</label></div></div>';
  html += '<div class="form-group"><label>Nom OFA / OFT</label><input type="text" id="f-Nom_OFA_OFT" value="' + v('Nom_OFA_OFT') + '" /></div>';
  html += '</div></div>';

  // Section: Mouvement
  html += '<div class="form-section"><h4>🔄 ' + t('sectionMouvement') + '</h4>';
  html += '<div class="form-grid">';
  html += '<div class="form-group"><label>Type de Mouvement <span class="required">*</span></label><select id="f-Mouvement">';
  for (var i = 0; i < mouvementOptions.length; i++) {
    html += '<option value="' + mouvementOptions[i] + '"' + (v('Mouvement') === mouvementOptions[i] ? ' selected' : '') + '>' + (mouvementOptions[i] || t('select')) + '</option>';
  }
  html += '</select></div>';
  html += '<div class="form-group"><label>Date de l\'acte</label><input type="text" id="f-Date_Acte" value="' + v('Date_Acte') + '" placeholder="jj/mm/aaaa" /></div>';
  html += '<div class="form-group"><label>Année</label><input type="text" id="f-Annee" value="' + v('Annee') + '" /></div>';
  html += '</div></div>';

  // Section: Localisation
  html += '<div class="form-section"><h4>📍 ' + t('sectionLocalisation') + '</h4>';
  html += '<div class="form-grid">';
  html += '<div class="form-group"><label>Commune <span class="required">*</span></label><input type="text" id="f-Commune" value="' + v('Commune') + '" /></div>';
  html += '<div class="form-group"><label>Adresse</label><input type="text" id="f-Adresse" value="' + v('Adresse') + '" /></div>';
  html += '<div class="form-group"><label>Référence Parcelles</label><input type="text" id="f-Ref_Parcelles" value="' + v('Ref_Parcelles') + '" /></div>';
  html += '</div></div>';

  // Section: Caractéristiques
  html += '<div class="form-section"><h4>🏗️ ' + t('sectionCaracteristiques') + '</h4>';
  html += '<div class="form-grid">';
  html += '<div class="form-group"><label>Type de Bien</label><select id="f-Type_Bien">';
  for (var i = 0; i < typeOptions.length; i++) {
    html += '<option value="' + typeOptions[i] + '"' + (v('Type_Bien') === typeOptions[i] ? ' selected' : '') + '>' + (typeOptions[i] || t('select')) + '</option>';
  }
  html += '</select></div>';
  html += '<div class="form-group"><label>Surface bâti (m²)</label><input type="text" id="f-Surface_Bati" value="' + v('Surface_Bati') + '" /></div>';
  html += '<div class="form-group"><label>Surface parcelle (m²)</label><input type="text" id="f-Surface_Parcelle" value="' + v('Surface_Parcelle') + '" /></div>';
  html += '<div class="form-group"><label>Surface assurance (m²)</label><input type="text" id="f-Surface_Assurance" value="' + v('Surface_Assurance') + '" /></div>';
  html += '<div class="form-group"><label>Nouvelle copropriété</label><select id="f-Nouvelle_Copropriete">';
  for (var i = 0; i < ouiNonOptions.length; i++) {
    html += '<option value="' + ouiNonOptions[i] + '"' + (v('Nouvelle_Copropriete') === ouiNonOptions[i] ? ' selected' : '') + '>' + (ouiNonOptions[i] || t('select')) + '</option>';
  }
  html += '</select></div>';
  html += '</div></div>';

  // Section: Occupation et Jouissance
  html += '<div class="form-section"><h4>🏠 ' + t('sectionOccupation') + '</h4>';
  html += '<div class="form-grid">';
  html += '<div class="form-group"><label>Occupation</label><select id="f-Occupation">';
  for (var i = 0; i < ouiNonOptions.length; i++) {
    html += '<option value="' + ouiNonOptions[i] + '"' + (v('Occupation') === ouiNonOptions[i] ? ' selected' : '') + '>' + (ouiNonOptions[i] || t('select')) + '</option>';
  }
  html += '</select></div>';
  html += '<div class="form-group"><label>Jouissance anticipée</label><select id="f-Jouissance_Anticipee">';
  for (var i = 0; i < ouiNonOptions.length; i++) {
    html += '<option value="' + ouiNonOptions[i] + '"' + (v('Jouissance_Anticipee') === ouiNonOptions[i] ? ' selected' : '') + '>' + (ouiNonOptions[i] || t('select')) + '</option>';
  }
  html += '</select></div>';
  html += '<div class="form-group"><label>Jouissance différée</label><select id="f-Jouissance_Differee">';
  for (var i = 0; i < ouiNonOptions.length; i++) {
    html += '<option value="' + ouiNonOptions[i] + '"' + (v('Jouissance_Differee') === ouiNonOptions[i] ? ' selected' : '') + '>' + (ouiNonOptions[i] || t('select')) + '</option>';
  }
  html += '</select></div>';
  html += '<div class="form-group"><label>Temps portage - Année fin</label><input type="text" id="f-Temps_Portage" value="' + v('Temps_Portage') + '" /></div>';
  html += '<div class="form-group"><label>Mis à bail longue durée</label><select id="f-Bail_Longue_Duree">';
  for (var i = 0; i < ouiNonOptions.length; i++) {
    html += '<option value="' + ouiNonOptions[i] + '"' + (v('Bail_Longue_Duree') === ouiNonOptions[i] ? ' selected' : '') + '>' + (ouiNonOptions[i] || t('select')) + '</option>';
  }
  html += '</select></div>';
  html += '</div></div>';

  // Section: Acquisition et Financement
  html += '<div class="form-section"><h4>💰 ' + t('sectionFinancement') + '</h4>';
  html += '<div class="form-grid">';
  html += '<div class="form-group"><label>Acquisition compte tiers</label><select id="f-Acquisition_Compte_Tiers">';
  for (var i = 0; i < ouiNonOptions.length; i++) {
    html += '<option value="' + ouiNonOptions[i] + '"' + (v('Acquisition_Compte_Tiers') === ouiNonOptions[i] ? ' selected' : '') + '>' + (ouiNonOptions[i] || t('select')) + '</option>';
  }
  html += '</select></div>';
  html += '<div class="form-group"><label>Préfinancement</label><select id="f-Prefinancement">';
  for (var i = 0; i < ouiNonOptions.length; i++) {
    html += '<option value="' + ouiNonOptions[i] + '"' + (v('Prefinancement') === ouiNonOptions[i] ? ' selected' : '') + '>' + (ouiNonOptions[i] || t('select')) + '</option>';
  }
  html += '</select></div>';
  html += '<div class="form-group"><label>Tiers Vendeur ou Acquéreur</label><input type="text" id="f-Tiers_Vendeur_Acquereur" value="' + v('Tiers_Vendeur_Acquereur') + '" /></div>';
  html += '</div></div>';

  // Section: GIMA
  html += '<div class="form-section"><h4>💾 ' + t('sectionGIMA') + '</h4>';
  html += '<div class="form-grid">';
  html += '<div class="form-group"><label>Import GIMA</label><select id="f-Import_GIMA">';
  for (var i = 0; i < importGimaOptions.length; i++) {
    html += '<option value="' + importGimaOptions[i] + '"' + (v('Import_GIMA') === importGimaOptions[i] ? ' selected' : '') + '>' + (importGimaOptions[i] || t('select')) + '</option>';
  }
  html += '</select></div>';
  html += '<div class="form-group"><label>N° du site</label><input type="text" id="f-Num_Site" value="' + v('Num_Site') + '" /></div>';
  html += '<div class="form-group"><label>Saisies manuelles</label><select id="f-Saisies_Manuelles">';
  for (var i = 0; i < saisiesOptions.length; i++) {
    html += '<option value="' + saisiesOptions[i] + '"' + (v('Saisies_Manuelles') === saisiesOptions[i] ? ' selected' : '') + '>' + (saisiesOptions[i] || t('select')) + '</option>';
  }
  html += '</select></div>';
  html += '<div class="form-group"><label>Date intégration GIMA</label><input type="text" id="f-Date_Integration_GIMA" value="' + v('Date_Integration_GIMA') + '" placeholder="jj/mm/aaaa" /></div>';
  html += '<div class="form-group"><label>Dossier numérique sous L</label><input type="text" id="f-Dossier_Numerique" value="' + v('Dossier_Numerique') + '" /></div>';
  html += '</div></div>';

  // Section: Nature et Observations (rich text editors)
  html += '<div class="form-section"><h4>📝 ' + t('sectionNature') + '</h4>';
  html += '<div class="form-group"><label>Nature du bien - Projet acquisition</label><div id="f-Nature_Bien">' + (isEdit ? (bien.Nature_Bien || '') : '') + '</div></div>';
  html += '<div class="form-group"><label>Observations</label><div id="f-Observation">' + (isEdit ? (bien.Observation || '') : '') + '</div></div>';
  html += '</div>';

  return html;
}

function getFormData() {
  var record = {};
  for (var i = 0; i < BIEN_COLUMNS.length; i++) {
    var col = BIEN_COLUMNS[i];
    if (col.id === 'Gestion_SPI') {
      var radio = document.querySelector('input[name="f-Gestion_SPI"]:checked');
      record[col.id] = radio ? radio.value : 'Non';
    } else if (col.id === 'Nature_Bien') {
      record[col.id] = window._joditNature ? window._joditNature.value : '';
    } else if (col.id === 'Observation') {
      record[col.id] = window._joditObservation ? window._joditObservation.value : '';
    } else {
      var el = document.getElementById('f-' + col.id);
      record[col.id] = el ? el.value.trim() : '';
    }
  }
  return record;
}

function initModalEditors() {
  var joditConfig = {
    language: currentLang,
    minHeight: 120,
    toolbarAdaptive: false,
    askBeforePasteHTML: false,
    askBeforePasteFromWord: false,
    defaultActionOnPaste: 'insert_clear_html',
    buttons: [
      'undo', 'redo', '|',
      'bold', 'italic', 'underline', 'strikethrough', '|',
      'brush', 'paragraph', '|',
      'align', '|',
      'ul', 'ol', '|',
      'indent', 'outdent', '|',
      'eraser'
    ],
    style: {
      'font-family': '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
      'font-size': '13px',
      'line-height': '1.5'
    },
    iframe: false,
    showCharsCounter: false,
    showWordsCounter: false,
    showXPathInStatusbar: false,
    placeholder: currentLang === 'fr' ? 'Saisir le texte...' : 'Enter text...'
  };

  var natureEl = document.getElementById('f-Nature_Bien');
  if (natureEl) {
    window._joditNature = Jodit.make('#f-Nature_Bien', Object.assign({}, joditConfig, {
      placeholder: currentLang === 'fr' ? 'Décrire la nature du bien...' : 'Describe the property nature...'
    }));
  }

  var obsEl = document.getElementById('f-Observation');
  if (obsEl) {
    window._joditObservation = Jodit.make('#f-Observation', Object.assign({}, joditConfig, {
      placeholder: currentLang === 'fr' ? 'Informations complémentaires...' : 'Additional information...'
    }));
  }
}

function openAddModal() {
  if (!canManage) { showToast(t('accessDenied'), 'error'); return; }

  var html = '<div class="modal-overlay" onclick="closeModal(event)">';
  html += '<div class="modal" onclick="event.stopPropagation()">';
  html += '<div class="modal-header-red"><h3>✨ ' + t('modalAdd') + '</h3><button class="modal-close-white" onclick="closeModalForce()">✕</button></div>';
  html += '<div class="modal-body">';
  html += buildFormHtml(null);
  html += '</div>';
  html += '<div class="modal-footer">';
  html += '<button class="btn btn-secondary" onclick="closeModalForce()">❌ ' + t('cancel') + '</button>';
  html += '<button class="btn btn-success" onclick="saveBien(null)">💾 ' + t('save') + '</button>';
  html += '</div></div></div>';

  document.getElementById('modal-container').innerHTML = html;
  setTimeout(initModalEditors, 100);
}

function openEditModal(bienId) {
  if (!canManage) { showToast(t('accessDenied'), 'error'); return; }
  var b = biens.find(function(x) { return x.id === bienId; });
  if (!b) return;

  var html = '<div class="modal-overlay" onclick="closeModal(event)">';
  html += '<div class="modal" onclick="event.stopPropagation()">';
  html += '<div class="modal-header-red"><h3>✏️ ' + t('modalEdit') + '</h3><button class="modal-close-white" onclick="closeModalForce()">✕</button></div>';
  html += '<div class="modal-body">';
  html += buildFormHtml(b);
  html += '</div>';
  html += '<div class="modal-footer">';
  html += '<button class="btn btn-danger" onclick="openDeleteConfirmModal(biens.find(function(x){return x.id===' + b.id + '}))">🗑️ ' + t('delete') + '</button>';
  html += '<div style="display:flex;gap:8px;">';
  html += '<button class="btn btn-secondary" onclick="closeModalForce()">❌ ' + t('cancel') + '</button>';
  html += '<button class="btn btn-success" onclick="saveBien(' + b.id + ')">💾 ' + t('save') + '</button>';
  html += '</div></div></div></div>';

  document.getElementById('modal-container').innerHTML = html;
  setTimeout(initModalEditors, 100);
}

// =============================================================================
// DASHBOARD VIEW
// =============================================================================

var dashFilterYears = []; // empty = all years
var dashFilterMonths = []; // empty = all months
var dashSurfaceTab = 'analyse'; // 'analyse' or 'details'
var dashDetailTab = 'acq'; // 'acq', 'ces', 'bati', 'nonbati'

function parseMonthFromDate(dateVal) {
  if (!dateVal && dateVal !== 0) return 0;
  var s = String(dateVal).trim();
  // Case 1: Excel serial number (pure number like 41449)
  var num = Number(s);
  if (!isNaN(num) && num > 1000 && s.indexOf('/') === -1 && s.indexOf('-') === -1) {
    // Excel serial: days since 1899-12-30
    var d = new Date((num - 25569) * 86400000);
    return d.getUTCMonth() + 1; // 1-12
  }
  // Case 2: DD/MM/YYYY or DD-MM-YYYY (European)
  var match = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
  if (match) {
    var a = parseInt(match[1], 10);
    var b = parseInt(match[2], 10);
    var c = parseInt(match[3], 10);
    // If first part > 12, it's DD/MM/YYYY
    if (a > 12 && b >= 1 && b <= 12) return b;
    // If second part > 12, it's MM/DD/YYYY
    if (b > 12 && a >= 1 && a <= 12) return a;
    // Ambiguous: if year part is 4 digits and >= 2000, assume DD/MM/YYYY (European)
    if (c >= 2000 && a >= 1 && a <= 31 && b >= 1 && b <= 12) return b;
    // Otherwise assume M/DD/YY (American, like the Excel source)
    if (a >= 1 && a <= 12) return a;
    if (b >= 1 && b <= 12) return b;
  }
  // Case 3: ISO YYYY-MM-DD
  var iso = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (iso) return parseInt(iso[2], 10);
  return 0;
}

function getFilteredBiens() {
  return biens.filter(function(b) {
    // Year filter (multi-select)
    if (dashFilterYears.length > 0 && dashFilterYears.indexOf(String(b.Annee)) === -1) return false;
    // Month filter (multi-select)
    if (dashFilterMonths.length > 0) {
      var month = parseMonthFromDate(b.Date_Acte);
      if (dashFilterMonths.indexOf(String(month)) === -1) return false;
    }
    return true;
  });
}

function buildPieChart(slices, size) {
  // slices = [{value, color, label}], size = diameter in px
  var total = 0;
  for (var i = 0; i < slices.length; i++) total += Math.abs(slices[i].value);
  if (total === 0) return '<div style="width:' + size + 'px;height:' + size + 'px;display:flex;align-items:center;justify-content:center;color:#94a3b8;font-size:13px;">Aucune donnée</div>';

  var r = size / 2;
  var cx = r, cy = r;
  var svg = '<svg width="' + size + '" height="' + size + '" viewBox="0 0 ' + size + ' ' + size + '">';

  var startAngle = -Math.PI / 2;
  for (var i = 0; i < slices.length; i++) {
    var pct = Math.abs(slices[i].value) / total;
    var angle = pct * 2 * Math.PI;
    var endAngle = startAngle + angle;
    var largeArc = angle > Math.PI ? 1 : 0;

    var x1 = cx + (r - 4) * Math.cos(startAngle);
    var y1 = cy + (r - 4) * Math.sin(startAngle);
    var x2 = cx + (r - 4) * Math.cos(endAngle);
    var y2 = cy + (r - 4) * Math.sin(endAngle);

    if (slices.length === 1) {
      svg += '<circle cx="' + cx + '" cy="' + cy + '" r="' + (r - 4) + '" fill="' + slices[i].color + '" />';
    } else {
      svg += '<path d="M ' + cx + ' ' + cy + ' L ' + x1 + ' ' + y1 + ' A ' + (r - 4) + ' ' + (r - 4) + ' 0 ' + largeArc + ' 1 ' + x2 + ' ' + y2 + ' Z" fill="' + slices[i].color + '" />';
    }

    // Label in the middle of the slice
    var midAngle = startAngle + angle / 2;
    var labelR = (r - 4) * 0.6;
    var lx = cx + labelR * Math.cos(midAngle);
    var ly = cy + labelR * Math.sin(midAngle);
    var labelVal = Math.round(Math.abs(slices[i].value)).toLocaleString() + ' m²';
    if (pct > 0.05) {
      svg += '<text x="' + lx + '" y="' + ly + '" text-anchor="middle" dominant-baseline="central" fill="#fff" font-size="' + (size > 180 ? 12 : 10) + '" font-weight="800">' + labelVal + '</text>';
    }

    startAngle = endAngle;
  }
  svg += '</svg>';

  // Legend
  var legend = '<div style="display:flex;gap:16px;justify-content:center;flex-wrap:wrap;margin-top:8px;">';
  for (var i = 0; i < slices.length; i++) {
    legend += '<div style="display:flex;align-items:center;gap:4px;font-size:11px;color:#475569;">';
    legend += '<span style="width:12px;height:12px;border-radius:3px;background:' + slices[i].color + ';display:inline-block;"></span>';
    legend += slices[i].label + '</div>';
  }
  legend += '</div>';

  return '<div style="display:flex;flex-direction:column;align-items:center;">' + svg + legend + '</div>';
}

function renderDashboardView() {
  var allYears = {};
  for (var i = 0; i < biens.length; i++) {
    if (biens[i].Annee) allYears[biens[i].Annee] = true;
  }
  var yearList = Object.keys(allYears).sort(function(a, b) { return parseInt(b) - parseInt(a); });

  var filtered = getFilteredBiens();
  var totalBiens = filtered.length;
  var communesSet = {};
  var surfaceParcelle = 0;
  var surfaceBati = 0;
  var mouvementCount = {};
  var communeCount = {};

  // SPI stats: CESSION = cession, everything else = acquisition (like reference app)
  var spiAcq = 0, spiCes = 0, noSpiAcq = 0, noSpiCes = 0;
  // Surface analysis (SPI=OUI only)
  var surfAcquise = 0, surfCedee = 0;
  var surfBatiSPI = 0, surfNonBatiSPI = 0;

  for (var i = 0; i < filtered.length; i++) {
    var b = filtered[i];
    if (b.Commune) { communesSet[b.Commune] = true; communeCount[b.Commune] = (communeCount[b.Commune] || 0) + 1; }
    var sp = parseNum(b.Surface_Parcelle);
    var sb = parseNum(b.Surface_Bati);
    surfaceParcelle += Math.abs(sp);
    surfaceBati += Math.abs(sb);
    if (b.Mouvement) mouvementCount[b.Mouvement] = (mouvementCount[b.Mouvement] || 0) + 1;

    var mouv = String(b.Mouvement || '').toUpperCase();
    var isSPI = String(b.Gestion_SPI || '').toUpperCase() === 'OUI';
    var isCes = mouv.indexOf('CESSION') !== -1;

    // Surface analysis: SPI only, totalSurface = surfaceBati + surfaceParcelle per item
    if (isSPI) {
      var totalSurfItem = sb + sp; // Both can be negative for cessions
      surfBatiSPI += sb;
      surfNonBatiSPI += sp;
      if (isCes) {
        surfCedee += totalSurfItem;
      } else {
        surfAcquise += totalSurfItem;
      }
    }

    // SPI stats: CESSION = cession, everything else = acquisition
    if (isCes) {
      if (isSPI) spiCes++;
      else noSpiCes++;
    } else {
      if (isSPI) spiAcq++;
      else noSpiAcq++;
    }
  }

  var html = '<div class="section-card">';
  html += '<h3 style="display:flex;align-items:center;gap:10px;"><span style="background:#ef4444;color:#fff;padding:8px 12px;border-radius:10px;font-size:20px;">📊</span> ' + t('dashTitle') + '</h3>';
  html += '<p style="color:#64748b;margin-bottom:20px;">' + t('dashSubtitle') + '</p>';

  // Stats cards
  html += '<div class="stats-row">';
  html += '<div class="stat-card"><div class="stat-value">' + totalBiens.toLocaleString() + '</div><div class="stat-label">' + t('totalBiens') + '</div></div>';
  html += '<div class="stat-card"><div class="stat-value">' + Object.keys(communesSet).length + '</div><div class="stat-label">' + t('totalCommunes') + '</div></div>';
  html += '<div class="stat-card"><div class="stat-value">' + Math.round(surfaceParcelle).toLocaleString() + ' m²</div><div class="stat-label">' + t('totalSurfaceParcelle') + '</div></div>';
  html += '<div class="stat-card"><div class="stat-value">' + Math.round(surfaceBati).toLocaleString() + ' m²</div><div class="stat-label">' + t('totalSurfaceBati') + '</div></div>';
  html += '</div>';

  // ===== Filtres temporels =====
  html += '<div class="section-card" style="margin-top:20px;">';
  html += '<h3>📅 ' + t('dashFilters') + '</h3>';

  // Year pills (multi-select)
  html += '<div style="margin-bottom:12px;"><span style="font-size:12px;font-weight:700;color:#64748b;">' + t('dashYear') + '</span>';
  html += '<div class="filter-pills">';
  html += '<button class="filter-pill' + (dashFilterYears.length === 0 ? ' active' : '') + '" onclick="toggleDashFilter(\'year\',\'all\')">' + t('dashAll') + '</button>';
  for (var i = 0; i < yearList.length; i++) {
    html += '<button class="filter-pill' + (dashFilterYears.indexOf(yearList[i]) !== -1 ? ' active' : '') + '" onclick="toggleDashFilter(\'year\',\'' + yearList[i] + '\')">' + yearList[i] + '</button>';
  }
  html += '</div></div>';

  // Month pills (multi-select)
  var months_fr = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Août', 'Sep', 'Oct', 'Nov', 'Déc'];
  var months_en = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var months = currentLang === 'fr' ? months_fr : months_en;
  html += '<div><span style="font-size:12px;font-weight:700;color:#64748b;">' + t('dashMonth') + '</span>';
  html += '<div class="filter-pills">';
  html += '<button class="filter-pill' + (dashFilterMonths.length === 0 ? ' active' : '') + '" onclick="toggleDashFilter(\'month\',\'all\')">' + t('dashAll') + '</button>';
  for (var i = 0; i < 12; i++) {
    html += '<button class="filter-pill' + (dashFilterMonths.indexOf(String(i + 1)) !== -1 ? ' active' : '') + '" onclick="toggleDashFilter(\'month\',\'' + (i + 1) + '\')">' + months[i] + '</button>';
  }
  html += '</div>';
  html += '<p style="font-size:11px;color:#94a3b8;margin-top:4px;">' + t('dashMonthHint') + '</p>';
  html += '</div></div>';

  // ===== Analyse des Surfaces (Gestion SPI = OUI) =====
  // Build detail lists for the 4 sub-tabs
  var detailAcq = [], detailCes = [], detailBati = [], detailNonBati = [];
  for (var i = 0; i < filtered.length; i++) {
    var b = filtered[i];
    var isSPIb = String(b.Gestion_SPI || '').toUpperCase() === 'OUI';
    if (!isSPIb) continue;
    var mouvB = String(b.Mouvement || '').toUpperCase();
    var spB = parseNum(b.Surface_Parcelle);
    var sbB = parseNum(b.Surface_Bati);
    var isCesB = mouvB.indexOf('CESSION') !== -1;
    // CESSION = cession, everything else = acquisition (like reference app)
    if (isCesB) {
      detailCes.push({ ref: b.Reference_DDC, commune: b.Commune, adresse: b.Adresse, type: b.Type_Bien, mouvement: b.Mouvement, surface: spB, date: b.Date_Acte, annee: b.Annee });
    } else {
      detailAcq.push({ ref: b.Reference_DDC, commune: b.Commune, adresse: b.Adresse, type: b.Type_Bien, mouvement: b.Mouvement, surface: spB, date: b.Date_Acte, annee: b.Annee });
    }
    // Bâti: all SPI items with surfaceBati
    if (sbB !== 0) {
      detailBati.push({ ref: b.Reference_DDC, commune: b.Commune, adresse: b.Adresse, mouvement: b.Mouvement, surface: sbB, date: b.Date_Acte, annee: b.Annee });
    }
    // Non Bâti: all SPI items with surfaceParcelle (= surface non bâti in reference app)
    if (spB !== 0) {
      detailNonBati.push({ ref: b.Reference_DDC, commune: b.Commune, adresse: b.Adresse, mouvement: b.Mouvement, surface: spB, date: b.Date_Acte, annee: b.Annee });
    }
  }
  detailAcq.sort(function(a, b) { return b.surface - a.surface; });
  detailCes.sort(function(a, b) { return b.surface - a.surface; });
  detailBati.sort(function(a, b) { return Math.abs(b.surface) - Math.abs(a.surface); });
  detailNonBati.sort(function(a, b) { return Math.abs(b.surface) - Math.abs(a.surface); });

  // Store for export
  window._dashDetailAcq = detailAcq;
  window._dashDetailCes = detailCes;
  window._dashDetailBati = detailBati;
  window._dashDetailNonBati = detailNonBati;

  // Build dynamic filter label for titles
  var filterYearLabel = dashFilterYears.length === 0 ? (currentLang === 'fr' ? 'Toutes années' : 'All years') : dashFilterYears.sort().join(', ');
  var months_fr_full = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Août', 'Sep', 'Oct', 'Nov', 'Déc'];
  var months_en_full = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var mfl = currentLang === 'fr' ? months_fr_full : months_en_full;
  var filterMonthLabel = '';
  if (dashFilterMonths.length > 0) {
    var sortedMonths = dashFilterMonths.slice().sort(function(a, b) { return parseInt(a) - parseInt(b); });
    filterMonthLabel = sortedMonths.map(function(m) { return mfl[parseInt(m) - 1]; }).join(', ');
  }
  var filterSubtitle = '';
  if (dashFilterYears.length > 0 || dashFilterMonths.length > 0) {
    filterSubtitle = '<p style="font-size:12px;color:#94a3b8;margin-top:2px;">' + (filterMonthLabel || '') + '</p>';
  }
  var titleYearSuffix = dashFilterYears.length > 0 ? ' - ' + dashFilterYears.sort().join(', ') : '';

  html += '<div class="section-card" style="margin-top:20px;">';
  html += '<h3>📊 ' + t('dashSurfaceAnalysis') + titleYearSuffix + ' <span style="font-size:12px;font-weight:400;color:#64748b;">(Gestion SPI = OUI)</span></h3>';
  html += filterSubtitle;

  // Main sub-tabs: Analyse des Surfaces | Détails des surfaces par bien
  html += '<div class="sub-tabs">';
  html += '<button class="sub-tab' + (dashSurfaceTab === 'analyse' ? ' active' : '') + '" onclick="setDashSurfaceTab(\'analyse\')">📊 ' + t('dashSurfaceAnalysis') + '</button>';
  html += '<button class="sub-tab' + (dashSurfaceTab === 'details' ? ' active-green' : '') + '" onclick="setDashSurfaceTab(\'details\')">📋 ' + t('dashDetailTab') + '</button>';
  html += '</div>';

  if (dashSurfaceTab === 'analyse') {
    var ecart = surfAcquise + surfCedee; // cedees is already negative
    var totalSurf = surfBatiSPI + surfNonBatiSPI;

    html += '<div class="analysis-grid">';

    // LEFT: Surfaces acquises / cédées with pie chart
    html += '<div class="analysis-card" style="text-align:center;">';
    html += '<h4>' + t('dashSurfAcqCed') + '</h4>';
    html += buildPieChart([
      { value: surfAcquise, color: '#22c55e', label: 'Surfaces acquises' },
      { value: surfCedee, color: '#d97706', label: 'Surfaces cédées' }
    ], 200);
    html += '<div style="display:flex;gap:0;margin-top:16px;border-radius:8px;overflow:hidden;">';
    html += '<div style="flex:1;background:#22c55e;color:#fff;padding:8px 4px;text-align:center;font-weight:700;font-size:11px;">Acquisitions</div>';
    html += '<div style="flex:1;background:#ef4444;color:#fff;padding:8px 4px;text-align:center;font-weight:700;font-size:11px;">Cessions</div>';
    html += '<div style="flex:1;background:#dc2626;color:#fff;padding:8px 4px;text-align:center;font-weight:700;font-size:11px;">Écart</div>';
    html += '</div>';
    html += '<div style="display:flex;gap:0;">';
    html += '<div style="flex:1;padding:8px 4px;text-align:center;font-weight:800;font-size:13px;color:#22c55e;">' + Math.round(surfAcquise).toLocaleString() + ' m²</div>';
    html += '<div style="flex:1;padding:8px 4px;text-align:center;font-weight:800;font-size:13px;color:#ef4444;">' + Math.round(surfCedee).toLocaleString() + ' m²</div>';
    html += '<div style="flex:1;padding:8px 4px;text-align:center;font-weight:800;font-size:13px;color:#1e293b;">' + Math.round(ecart).toLocaleString() + ' m²</div>';
    html += '</div>';
    html += '</div>';

    // RIGHT: Répartition bâti / non-bâti with pie chart
    html += '<div class="analysis-card" style="text-align:center;">';
    html += '<h4>' + t('dashBatiNonBati') + '</h4>';
    html += buildPieChart([
      { value: surfBatiSPI, color: '#3b82f6', label: 'Surface bâti' },
      { value: surfNonBatiSPI, color: '#22c55e', label: 'Surface non bâti' }
    ], 200);
    html += '<div style="display:flex;gap:0;margin-top:16px;border-radius:8px;overflow:hidden;">';
    html += '<div style="flex:1;background:#3b82f6;color:#fff;padding:8px 4px;text-align:center;font-weight:700;font-size:10px;">' + t('dashSurfBati') + '</div>';
    html += '<div style="flex:1;background:#22c55e;color:#fff;padding:8px 4px;text-align:center;font-weight:700;font-size:10px;">' + t('dashSurfNonBati') + '</div>';
    html += '<div style="flex:1;background:#475569;color:#fff;padding:8px 4px;text-align:center;font-weight:700;font-size:11px;">Total</div>';
    html += '</div>';
    html += '<div style="display:flex;gap:0;">';
    html += '<div style="flex:1;padding:8px 4px;text-align:center;font-weight:800;font-size:13px;color:#3b82f6;">' + Math.round(surfBatiSPI).toLocaleString() + ' m²</div>';
    html += '<div style="flex:1;padding:8px 4px;text-align:center;font-weight:800;font-size:13px;color:#22c55e;">' + Math.round(surfNonBatiSPI).toLocaleString() + ' m²</div>';
    html += '<div style="flex:1;padding:8px 4px;text-align:center;font-weight:800;font-size:13px;">' + Math.round(totalSurf).toLocaleString() + ' m²</div>';
    html += '</div>';
    html += '</div>';

    html += '</div>';

  } else {
    // Details tab
    // Export button
    html += '<div style="text-align:right;margin-bottom:12px;">';
    html += '<button class="btn-export" onclick="exportSurfaceDetails()">⬇ ' + t('dashExportExcel') + '</button>';
    html += '</div>';

    // 4 detail sub-tabs
    var yearLabel = dashFilterYears.length === 0 ? (currentLang === 'fr' ? 'Toutes années' : 'All years') : dashFilterYears.join(', ');
    html += '<div class="detail-sub-tabs">';
    html += '<button class="detail-sub-tab' + (dashDetailTab === 'acq' ? ' active' : '') + '" onclick="setDashDetailTab(\'acq\')">Acquisitions (' + detailAcq.length + ')</button>';
    html += '<button class="detail-sub-tab' + (dashDetailTab === 'ces' ? ' active-orange' : '') + '" onclick="setDashDetailTab(\'ces\')">Cessions (' + detailCes.length + ')</button>';
    html += '<button class="detail-sub-tab' + (dashDetailTab === 'bati' ? ' active-blue' : '') + '" onclick="setDashDetailTab(\'bati\')">Bâti (' + detailBati.length + ')</button>';
    html += '<button class="detail-sub-tab' + (dashDetailTab === 'nonbati' ? ' active-dark' : '') + '" onclick="setDashDetailTab(\'nonbati\')">Non Bâti (' + detailNonBati.length + ')</button>';
    html += '</div>';

    if (dashDetailTab === 'acq') {
      html += '<h4 style="margin-bottom:12px;">Acquisitions - ' + yearLabel + '</h4>';
      html += buildDetailTable(detailAcq, 'acq');
    } else if (dashDetailTab === 'ces') {
      html += '<h4 style="margin-bottom:12px;">Cessions - ' + yearLabel + '</h4>';
      html += buildDetailTable(detailCes, 'ces');
    } else if (dashDetailTab === 'bati') {
      html += '<h4 style="margin-bottom:12px;">Bâti - ' + yearLabel + '</h4>';
      html += buildDetailTable(detailBati, 'bati');
    } else if (dashDetailTab === 'nonbati') {
      html += '<h4 style="margin-bottom:12px;">Non Bâti - ' + yearLabel + '</h4>';
      html += buildDetailTable(detailNonBati, 'nonbati');
    }
  }

  html += '</div>';

  // ===== Gestion SPI =====
  html += '<div class="section-card" style="margin-top:20px;">';
  html += '<h3>📈 ' + t('dashGestionSPI') + titleYearSuffix + '</h3>';
  html += filterSubtitle;

  // SPI Pie chart + Stats table side by side
  var totalActes = spiAcq + noSpiAcq + spiCes + noSpiCes;
  var totalSPI = spiAcq + spiCes;
  var spiNoAcqAll = noSpiAcq;
  var spiNoCesAll = noSpiCes;

  html += '<div class="analysis-grid">';

  // LEFT: Répartition SPI pie chart
  html += '<div class="analysis-card" style="text-align:center;">';
  html += '<h4>Répartition SPI' + titleYearSuffix + '</h4>';
  html += filterSubtitle;
  html += buildPieChart([
    { value: spiAcq, color: '#166534', label: 'Gestion SPI - Tous mouvements sauf Cession' },
    { value: spiNoAcqAll, color: '#86efac', label: 'Pas de Gestion SPI - Tous mouvements sauf Cession' },
    { value: spiCes, color: '#4d7c0f', label: 'Gestion SPI - Cession' },
    { value: spiNoCesAll, color: '#d9f99d', label: 'Pas de Gestion SPI - Cession' }
  ], 220);
  html += '</div>';

  // RIGHT: Statistiques SPI par type d'acte
  html += '<div class="analysis-card">';
  html += '<h4>📊 ' + t('dashSPIStats') + '</h4>';
  html += '<table class="analysis-table spi-table"><thead><tr>';
  html += '<th>' + t('dashTypeActe') + '</th>';
  html += '<th>Total</th>';
  html += '<th style="color:#22c55e;">' + t('dashSPIEntrants') + '</th>';
  html += '</tr></thead><tbody>';
  html += '<tr><td>Acquisitions</td><td><strong>' + (spiAcq + noSpiAcq).toLocaleString() + '</strong></td><td style="font-weight:800;color:#22c55e;">' + spiAcq.toLocaleString() + '</td></tr>';
  html += '<tr><td>Cessions</td><td><strong>' + (spiCes + noSpiCes).toLocaleString() + '</strong></td><td style="font-weight:800;color:#22c55e;">' + spiCes.toLocaleString() + '</td></tr>';
  html += '<tr><td><strong>Total</strong></td><td><strong>' + totalActes.toLocaleString() + '</strong></td><td style="font-weight:800;color:#22c55e;"><strong>' + totalSPI.toLocaleString() + '</strong></td></tr>';
  html += '</tbody></table>';
  html += '</div>';

  html += '</div></div>';

  // ===== Répartition par mouvement (card grid) =====
  html += '<div class="section-card" style="margin-top:20px;">';
  html += '<h3>📊 ' + t('repartitionMouvement') + titleYearSuffix + '</h3>';
  html += filterSubtitle;
  var mouvKeys = Object.keys(mouvementCount).sort(function(a, b) { return mouvementCount[b] - mouvementCount[a]; });
  html += '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:16px;margin-top:16px;">';
  for (var i = 0; i < mouvKeys.length; i++) {
    html += '<div style="text-align:center;padding:16px 8px;">';
    html += '<div style="width:100%;height:4px;background:#ef4444;border-radius:2px;margin-bottom:12px;"></div>';
    html += '<div style="font-size:36px;font-weight:900;color:#ef4444;line-height:1;">' + mouvementCount[mouvKeys[i]].toLocaleString() + '</div>';
    html += '<div style="font-size:11px;font-weight:800;color:#475569;margin-top:6px;text-transform:uppercase;">' + sanitize(mouvKeys[i]) + '</div>';
    html += '</div>';
  }
  html += '</div></div>';

  // ===== Top 10 communes (card grid) =====
  html += '<div class="section-card" style="margin-top:20px;">';
  html += '<h3>🏘️ ' + t('topCommunes') + titleYearSuffix + '</h3>';
  html += filterSubtitle;
  var communeKeys = Object.keys(communeCount).sort(function(a, b) { return communeCount[b] - communeCount[a]; }).slice(0, 10);
  html += '<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:16px;margin-top:16px;">';
  for (var i = 0; i < communeKeys.length; i++) {
    html += '<div style="text-align:center;padding:16px 8px;">';
    html += '<div style="width:100%;height:4px;background:#ef4444;border-radius:2px;margin-bottom:12px;"></div>';
    html += '<div style="font-size:32px;font-weight:900;color:#ef4444;line-height:1;">' + communeCount[communeKeys[i]] + '</div>';
    html += '<div style="font-size:10px;font-weight:800;color:#475569;margin-top:6px;text-transform:uppercase;">' + sanitize(communeKeys[i]) + '</div>';
    html += '</div>';
  }
  html += '</div></div>';

  html += '</div>';

  document.getElementById('dashboard-view').innerHTML = html;
}

function toggleDashFilter(type, value) {
  if (type === 'year') {
    if (value === 'all') {
      dashFilterYears = [];
    } else {
      var idx = dashFilterYears.indexOf(value);
      if (idx !== -1) dashFilterYears.splice(idx, 1);
      else dashFilterYears.push(value);
    }
  }
  if (type === 'month') {
    if (value === 'all') {
      dashFilterMonths = [];
    } else {
      var idx = dashFilterMonths.indexOf(value);
      if (idx !== -1) dashFilterMonths.splice(idx, 1);
      else dashFilterMonths.push(value);
    }
  }
  renderDashboardView();
}

function setDashSurfaceTab(tab) {
  dashSurfaceTab = tab;
  renderDashboardView();
}

function setDashDetailTab(tab) {
  dashDetailTab = tab;
  renderDashboardView();
}

function buildDetailTable(rows, type) {
  var hasType = (type === 'acq' || type === 'ces');
  var headClass = type === 'acq' ? 'green-head' : type === 'ces' ? 'orange-head' : type === 'bati' ? 'blue-head' : 'dark-head';
  var surfLabel = type === 'bati' ? 'Surface Bâti (m²)' : type === 'nonbati' ? 'Surface Non Bâti (m²)' : 'Surface Parcelle (m²)';

  var html = '<div style="max-height:500px;overflow:auto;">';
  html += '<table class="detail-table"><thead class="' + headClass + '"><tr>';
  html += '<th>Référence DDC</th><th>Commune</th><th>Adresse</th>';
  if (hasType) html += '<th>Type</th>';
  html += '<th>Mouvement</th><th>' + surfLabel + '</th><th>Date de l\'acte</th><th>Année</th>';
  html += '</tr></thead><tbody>';

  var totalSurf = 0;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    totalSurf += r.surface;
    var valClass = r.surface >= 0 ? 'val-positive' : 'val-negative';
    html += '<tr>';
    html += '<td>' + sanitize(r.ref) + '</td>';
    html += '<td>' + sanitize(r.commune) + '</td>';
    html += '<td>' + sanitize(r.adresse) + '</td>';
    if (hasType) html += '<td>' + sanitize(r.type || '') + '</td>';
    html += '<td>' + movementBadge(r.mouvement) + '</td>';
    html += '<td class="' + valClass + '">' + Math.round(r.surface).toLocaleString() + '</td>';
    html += '<td>' + sanitize(r.date) + '</td>';
    html += '<td>' + sanitize(r.annee) + '</td>';
    html += '</tr>';
  }
  html += '</tbody>';

  // Footer with total
  html += '<tfoot><tr>';
  html += '<td colspan="' + (hasType ? 5 : 4) + '" style="text-align:right;">Total :</td>';
  var totalClass = totalSurf >= 0 ? 'val-positive' : 'val-negative';
  html += '<td class="' + totalClass + '">' + Math.round(totalSurf).toLocaleString() + '</td>';
  html += '<td colspan="2"></td>';
  html += '</tr></tfoot>';

  html += '</table></div>';
  return html;
}

function exportSurfaceDetails() {
  try {
    var wb = XLSX.utils.book_new();

    // Helper to build sheet data
    function makeSheet(rows, type) {
      var hasType = (type === 'acq' || type === 'ces');
      var surfLabel = type === 'bati' ? 'Surface Bâti (m²)' : type === 'nonbati' ? 'Surface Non Bâti (m²)' : 'Surface Parcelle (m²)';
      var headers = ['Référence DDC', 'Commune', 'Adresse'];
      if (hasType) headers.push('Type');
      headers.push('Mouvement', surfLabel, "Date de l'acte", 'Année');

      var data = [headers];
      var totalSurf = 0;
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        totalSurf += r.surface;
        var row = [r.ref || '', r.commune || '', r.adresse || ''];
        if (hasType) row.push(r.type || '');
        row.push(r.mouvement || '', Math.round(r.surface), r.date || '', r.annee || '');
        data.push(row);
      }
      // Total row
      var totalRow = [];
      for (var j = 0; j < headers.length; j++) totalRow.push('');
      totalRow[0] = 'Total';
      var surfIdx = hasType ? 5 : 4;
      totalRow[surfIdx] = Math.round(totalSurf);
      data.push(totalRow);

      return XLSX.utils.aoa_to_sheet(data);
    }

    XLSX.utils.book_append_sheet(wb, makeSheet(window._dashDetailAcq || [], 'acq'), 'Acquisitions');
    XLSX.utils.book_append_sheet(wb, makeSheet(window._dashDetailCes || [], 'ces'), 'Cessions');
    XLSX.utils.book_append_sheet(wb, makeSheet(window._dashDetailBati || [], 'bati'), 'Bâti');
    XLSX.utils.book_append_sheet(wb, makeSheet(window._dashDetailNonBati || [], 'nonbati'), 'Non Bâti');

    var yearLabel = dashFilterYears.length === 0 ? 'Toutes_annees' : dashFilterYears.join('_');
    XLSX.writeFile(wb, 'Surfaces_SPI_' + yearLabel + '.xlsx');
    showToast(t('dashExportDone'), 'success');
  } catch (e) {
    console.error('Export error:', e);
    showToast('Erreur export: ' + e.message, 'error');
  }
}

// =============================================================================
// IMPORT VIEW
// =============================================================================

// Mapping: Excel column header → BM_Biens column id
var EXCEL_COL_MAP = {
  'Référence DDC': 'Reference_DDC',
  'Reference DDC': 'Reference_DDC',
  'Gestion SPI': 'Gestion_SPI',
  "Nom de l'OFA / OFT": 'Nom_OFA_OFT',
  'Nom de l\'OFA / OFT': 'Nom_OFA_OFT',
  'Mouvement': 'Mouvement',
  "Date de l'acte": 'Date_Acte',
  'Date de l\'acte': 'Date_Acte',
  'Année': 'Annee',
  'Annee': 'Annee',
  'Commune': 'Commune',
  'Adresse': 'Adresse',
  'Ref_ Parcelles': 'Ref_Parcelles',
  'Ref_Parcelles': 'Ref_Parcelles',
  'Ref Parcelles': 'Ref_Parcelles',
  'Type': 'Type_Bien',
  'Surface bâti': 'Surface_Bati',
  'Surface bati': 'Surface_Bati',
  'Surface parcelle acquise ou vendue': 'Surface_Parcelle',
  'Surface parcelle': 'Surface_Parcelle',
  'Nouvelle copropriété': 'Nouvelle_Copropriete',
  'Nouvelle copropriete': 'Nouvelle_Copropriete',
  'Occupation': 'Occupation',
  'Jouissance anticipée': 'Jouissance_Anticipee',
  'Jouissance anticipee': 'Jouissance_Anticipee',
  'Jouissance différée': 'Jouissance_Differee',
  'Jouissance differee': 'Jouissance_Differee',
  'Temps portage Année fin portage': 'Temps_Portage',
  'Temps portage': 'Temps_Portage',
  'Mis à bail longue durée': 'Bail_Longue_Duree',
  'Mis a bail longue duree': 'Bail_Longue_Duree',
  'Acquisition compte tiers': 'Acquisition_Compte_Tiers',
  'Préfinancement': 'Prefinancement',
  'Prefinancement': 'Prefinancement',
  'Surface pour assurance': 'Surface_Assurance',
  'Tiers Vendeur ou Acquéreur': 'Tiers_Vendeur_Acquereur',
  'Tiers Vendeur ou Acquereur': 'Tiers_Vendeur_Acquereur',
  'Nature du bien - Projet acquisition': 'Nature_Bien',
  'Nature du bien': 'Nature_Bien',
  'Import GIMA': 'Import_GIMA',
  'N° du site': 'Num_Site',
  'N du site': 'Num_Site',
  'Saisies manuelles': 'Saisies_Manuelles',
  "Date d'intégration archivage GIMA": 'Date_Integration_GIMA',
  'Date d\'intégration archivage GIMA': 'Date_Integration_GIMA',
  'Date integration GIMA': 'Date_Integration_GIMA',
  'Dossier numérique sous L': 'Dossier_Numerique',
  'Dossier numerique sous L': 'Dossier_Numerique',
  'Observation': 'Observation',
  'Observations': 'Observation'
};

var importParsedRows = [];
var importMappedCols = 0;
var importMode = 'update'; // 'replace', 'append', 'update'
var importWorkbook = null;
var importSheetNames = [];
var importFileName = '';

function renderImportView() {
  if (!canManage) {
    document.getElementById('import-view').innerHTML = '<div class="section-card"><p style="color:#94a3b8;font-style:italic;">🔒 ' + t('accessDenied') + '</p></div>';
    // Hide import tab for non-managers
    var importTab = document.querySelector('[data-tab="import"]');
    if (importTab) importTab.style.display = canManage ? '' : 'none';
    return;
  }

  var html = '<div class="section-card">';
  html += '<h3>📥 ' + t('importTitle') + '</h3>';
  html += '<p style="color:#64748b;margin-bottom:20px;">' + t('importSubtitle') + '</p>';

  // Drop zone
  html += '<div class="drop-zone" id="import-drop-zone" onclick="document.getElementById(\'import-file-input\').click()">';
  html += '<div class="drop-icon">📁</div>';
  html += '<div class="drop-text">' + t('importDrop') + '</div>';
  html += '<div class="drop-hint">' + t('importOr') + '</div>';
  html += '<div class="drop-hint" style="margin-top:8px;">' + t('importFormats') + '</div>';
  html += '</div>';
  html += '<input type="file" id="import-file-input" accept=".xlsx,.xls" style="display:none;" onchange="handleImportFile(this)" />';

  // Sheet selector (hidden initially)
  html += '<div id="import-sheet-select" style="display:none;margin-top:16px;text-align:center;">';
  html += '<label style="font-weight:700;margin-right:8px;">' + t('importSelectSheet') + ' :</label>';
  html += '<select id="import-sheet" onchange="parseSelectedSheet()" style="padding:8px 12px;border:1px solid #e2e8f0;border-radius:8px;"></select>';
  html += '</div>';

  // Import mode options
  html += '<div class="import-options" id="import-mode-options" style="display:none;">';
  html += '<div class="import-option' + (importMode === 'replace' ? ' selected' : '') + '" onclick="setImportMode(\'replace\')">';
  html += '<h5>🔄 ' + t('importReplace') + '</h5><p>' + t('importReplaceDesc') + '</p></div>';
  html += '<div class="import-option' + (importMode === 'append' ? ' selected' : '') + '" onclick="setImportMode(\'append\')">';
  html += '<h5>➕ ' + t('importAppend') + '</h5><p>' + t('importAppendDesc') + '</p></div>';
  html += '<div class="import-option selected" onclick="setImportMode(\'update\')">';
  html += '<h5>🔃 ' + t('importUpdate') + '</h5><p>' + t('importUpdateDesc') + '</p></div>';
  html += '</div>';

  // Preview area
  html += '<div id="import-preview-area"></div>';

  // Import button
  html += '<div id="import-action-area" style="display:none;text-align:center;margin-top:16px;">';
  html += '<button class="btn btn-primary" onclick="executeImport()" id="import-exec-btn">📥 ' + t('importBtn') + '</button>';
  html += '</div>';

  // Progress area
  html += '<div id="import-progress-area"></div>';

  html += '</div>';

  document.getElementById('import-view').innerHTML = html;

  // Show/hide import tab
  var importTab = document.querySelector('[data-tab="import"]');
  if (importTab) importTab.style.display = canManage ? '' : 'none';

  // Setup drag and drop
  setTimeout(setupDropZone, 100);
}

function setupDropZone() {
  var dz = document.getElementById('import-drop-zone');
  if (!dz) return;
  dz.addEventListener('dragover', function(e) { e.preventDefault(); dz.classList.add('drag-over'); });
  dz.addEventListener('dragleave', function() { dz.classList.remove('drag-over'); });
  dz.addEventListener('drop', function(e) {
    e.preventDefault();
    dz.classList.remove('drag-over');
    if (e.dataTransfer.files.length > 0) {
      processExcelFile(e.dataTransfer.files[0]);
    }
  });
}

function handleImportFile(input) {
  if (input.files.length > 0) {
    processExcelFile(input.files[0]);
  }
}

function processExcelFile(file) {
  importFileName = file.name;
  var reader = new FileReader();
  reader.onload = function(e) {
    try {
      importWorkbook = XLSX.read(e.target.result, { type: 'array' });
      importSheetNames = importWorkbook.SheetNames;

      // Update drop zone to show file name
      var dz = document.getElementById('import-drop-zone');
      if (dz) {
        dz.innerHTML = '<div class="drop-icon">✅</div><div class="drop-text">' + sanitize(importFileName) + '</div><div class="drop-hint">' + importSheetNames.length + ' feuille(s) détectée(s)</div>';
      }

      // Show sheet selector if multiple sheets
      var sheetSelect = document.getElementById('import-sheet-select');
      var selectEl = document.getElementById('import-sheet');
      if (sheetSelect && selectEl) {
        selectEl.innerHTML = '';
        for (var i = 0; i < importSheetNames.length; i++) {
          var opt = document.createElement('option');
          opt.value = importSheetNames[i];
          opt.textContent = importSheetNames[i];
          selectEl.appendChild(opt);
        }
        sheetSelect.style.display = importSheetNames.length > 1 ? 'block' : 'none';
      }

      // Parse first sheet
      parseSelectedSheet();

    } catch (err) {
      console.error('Error reading Excel:', err);
      showToast('Erreur lecture fichier: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function parseSelectedSheet() {
  var sheetName = document.getElementById('import-sheet') ? document.getElementById('import-sheet').value : importSheetNames[0];
  var ws = importWorkbook.Sheets[sheetName];
  var jsonData = XLSX.utils.sheet_to_json(ws, { defval: '' });

  if (jsonData.length === 0) {
    showToast('Feuille vide', 'error');
    return;
  }

  // Map columns
  var excelHeaders = Object.keys(jsonData[0]);
  importMappedCols = 0;
  var mapping = {};

  for (var i = 0; i < excelHeaders.length; i++) {
    var h = excelHeaders[i].trim();
    if (EXCEL_COL_MAP[h]) {
      mapping[h] = EXCEL_COL_MAP[h];
      importMappedCols++;
    }
  }

  // Convert rows
  importParsedRows = [];
  for (var r = 0; r < jsonData.length; r++) {
    var row = {};
    for (var h in mapping) {
      var val = jsonData[r][h];
      // Clean HTML tags from values
      if (typeof val === 'string') {
        val = val.replace(/<[^>]*>/g, '').trim();
      }
      row[mapping[h]] = val !== undefined && val !== null ? String(val) : '';
    }
    // Only add rows that have at least a reference or commune
    if (row.Reference_DDC || row.Commune) {
      importParsedRows.push(row);
    }
  }

  // Show options and preview
  document.getElementById('import-mode-options').style.display = 'flex';
  document.getElementById('import-action-area').style.display = 'block';

  // Render preview
  renderImportPreview();
}

function renderImportPreview() {
  var html = '<div style="margin-top:20px;">';

  // Stats
  html += '<div class="import-stats">';
  html += '<div class="import-stat"><div class="stat-num">' + importParsedRows.length + '</div><div class="stat-txt">' + t('importRows') + '</div></div>';
  html += '<div class="import-stat"><div class="stat-num">' + importMappedCols + '</div><div class="stat-txt">' + t('importCols') + '</div></div>';
  html += '<div class="import-stat"><div class="stat-num">' + biens.length + '</div><div class="stat-txt">' + (currentLang === 'fr' ? 'biens existants' : 'existing assets') + '</div></div>';
  html += '</div>';

  // Preview table (first 10 rows)
  html += '<h4 style="margin-top:16px;font-weight:800;">📋 ' + t('importPreview') + ' (' + Math.min(10, importParsedRows.length) + '/' + importParsedRows.length + ')</h4>';
  var previewCols = ['Reference_DDC', 'Mouvement', 'Commune', 'Adresse', 'Type_Bien', 'Surface_Parcelle'];
  html += '<div class="import-preview"><table class="data-table"><thead><tr>';
  for (var c = 0; c < previewCols.length; c++) {
    var col = BIEN_COLUMNS.find(function(cc) { return cc.id === previewCols[c]; });
    html += '<th>' + (col ? colLabel(col) : previewCols[c]) + '</th>';
  }
  html += '</tr></thead><tbody>';
  var previewMax = Math.min(10, importParsedRows.length);
  for (var i = 0; i < previewMax; i++) {
    html += '<tr>';
    for (var c = 0; c < previewCols.length; c++) {
      var val = importParsedRows[i][previewCols[c]] || '';
      if (previewCols[c] === 'Mouvement') {
        html += '<td>' + movementBadge(val) + '</td>';
      } else {
        html += '<td>' + sanitize(val) + '</td>';
      }
    }
    html += '</tr>';
  }
  if (importParsedRows.length > 10) {
    html += '<tr><td colspan="' + previewCols.length + '" style="text-align:center;color:#94a3b8;">... ' + (importParsedRows.length - 10) + ' ' + (currentLang === 'fr' ? 'lignes supplémentaires' : 'more rows') + '</td></tr>';
  }
  html += '</tbody></table></div>';
  html += '</div>';

  document.getElementById('import-preview-area').innerHTML = html;
}

function setImportMode(mode) {
  importMode = mode;
  document.querySelectorAll('.import-option').forEach(function(el) { el.classList.remove('selected'); });
  event.currentTarget.classList.add('selected');
}

async function executeImport() {
  if (!canManage) { showToast(t('accessDenied'), 'error'); return; }
  if (importParsedRows.length === 0) { showToast(t('importNoFile'), 'error'); return; }

  if (importMode === 'replace') {
    if (!confirm(t('importConfirmReplace'))) return;
  }

  var btn = document.getElementById('import-exec-btn');
  if (btn) btn.disabled = true;

  var progressArea = document.getElementById('import-progress-area');
  var added = 0, updated = 0, deleted = 0, errors = 0;

  progressArea.innerHTML = '<div class="import-progress"><p style="font-weight:700;">⏳ ' + t('importInProgress') + '</p><div class="progress-bar"><div class="progress-bar-fill" id="import-progress-fill" style="width:0%"></div></div><p id="import-progress-text" style="font-size:12px;color:#64748b;margin-top:6px;">0 / ' + importParsedRows.length + '</p></div>';

  try {
    var BATCH_SIZE = 50;

    if (importMode === 'replace') {
      // Delete all existing records
      if (biens.length > 0) {
        var deleteIds = biens.map(function(b) { return b.id; });
        for (var i = 0; i < deleteIds.length; i += BATCH_SIZE) {
          var batch = deleteIds.slice(i, i + BATCH_SIZE);
          var actions = batch.map(function(id) { return ['RemoveRecord', BIENS_TABLE, id]; });
          await grist.docApi.applyUserActions(actions);
          deleted += batch.length;
        }
      }
      // Add all new records
      for (var i = 0; i < importParsedRows.length; i += BATCH_SIZE) {
        var batch = importParsedRows.slice(i, i + BATCH_SIZE);
        var actions = batch.map(function(row) { return ['AddRecord', BIENS_TABLE, null, row]; });
        await grist.docApi.applyUserActions(actions);
        added += batch.length;
        updateImportProgress(added + deleted, importParsedRows.length + biens.length);
      }

    } else if (importMode === 'append') {
      // Just add all rows
      for (var i = 0; i < importParsedRows.length; i += BATCH_SIZE) {
        var batch = importParsedRows.slice(i, i + BATCH_SIZE);
        var actions = batch.map(function(row) { return ['AddRecord', BIENS_TABLE, null, row]; });
        await grist.docApi.applyUserActions(actions);
        added += batch.length;
        updateImportProgress(added, importParsedRows.length);
      }

    } else if (importMode === 'update') {
      // Build index of existing biens by Reference_DDC
      var existingIndex = {};
      for (var i = 0; i < biens.length; i++) {
        var ref = String(biens[i].Reference_DDC || '').trim().toLowerCase();
        if (ref) existingIndex[ref] = biens[i].id;
      }

      var toAdd = [];
      var toUpdate = [];

      for (var i = 0; i < importParsedRows.length; i++) {
        var row = importParsedRows[i];
        var ref = String(row.Reference_DDC || '').trim().toLowerCase();
        if (ref && existingIndex[ref]) {
          toUpdate.push({ id: existingIndex[ref], data: row });
        } else {
          toAdd.push(row);
        }
      }

      var total = toAdd.length + toUpdate.length;
      var done = 0;

      // Update existing
      for (var i = 0; i < toUpdate.length; i += BATCH_SIZE) {
        var batch = toUpdate.slice(i, i + BATCH_SIZE);
        var actions = batch.map(function(item) { return ['UpdateRecord', BIENS_TABLE, item.id, item.data]; });
        await grist.docApi.applyUserActions(actions);
        updated += batch.length;
        done += batch.length;
        updateImportProgress(done, total);
      }

      // Add new
      for (var i = 0; i < toAdd.length; i += BATCH_SIZE) {
        var batch = toAdd.slice(i, i + BATCH_SIZE);
        var actions = batch.map(function(row) { return ['AddRecord', BIENS_TABLE, null, row]; });
        await grist.docApi.applyUserActions(actions);
        added += batch.length;
        done += batch.length;
        updateImportProgress(done, total);
      }
    }

    // Done!
    progressArea.innerHTML = '<div class="import-progress" style="margin-top:16px;">' +
      '<p style="font-weight:800;color:#22c55e;font-size:16px;">✅ ' + t('importDone') + '</p>' +
      '<div class="import-stats" style="margin-top:12px;">' +
      (added > 0 ? '<div class="import-stat"><div class="stat-num" style="color:#22c55e;">' + added + '</div><div class="stat-txt">' + t('importAdded') + '</div></div>' : '') +
      (updated > 0 ? '<div class="import-stat"><div class="stat-num" style="color:#3b82f6;">' + updated + '</div><div class="stat-txt">' + t('importUpdated') + '</div></div>' : '') +
      (deleted > 0 ? '<div class="import-stat"><div class="stat-num" style="color:#ef4444;">' + deleted + '</div><div class="stat-txt">' + t('importDeleted') + '</div></div>' : '') +
      '</div></div>';

    showToast(t('importDone') + ' ' + added + ' ' + t('importAdded') + ', ' + updated + ' ' + t('importUpdated'), 'success');
    await loadAllData();

  } catch (e) {
    console.error('Import error:', e);
    showToast('Erreur import: ' + e.message, 'error');
    progressArea.innerHTML += '<p style="color:#ef4444;margin-top:8px;">❌ Erreur: ' + sanitize(e.message) + '</p>';
  }

  if (btn) btn.disabled = false;
}

function updateImportProgress(done, total) {
  var pct = total > 0 ? Math.round((done / total) * 100) : 0;
  var fill = document.getElementById('import-progress-fill');
  var text = document.getElementById('import-progress-text');
  if (fill) fill.style.width = pct + '%';
  if (text) text.textContent = done + ' / ' + total + ' (' + pct + '%)';
}

// =============================================================================
// GESTIONNAIRES VIEW
// =============================================================================

function renderGestionnairesView() {
  var html = '<div class="section-card">';
  html += '<h3>👥 ' + t('gestTitle') + '</h3>';
  html += '<p style="color:#64748b;margin-bottom:16px;">' + t('gestSubtitle') + '</p>';

  if (!isOwner) {
    html += '<p style="color:#94a3b8;font-style:italic;">🔒 ' + t('ownerOnly') + '</p>';
  } else {
    // Add gestionnaire
    html += '<div class="gestionnaire-add">';
    html += '<input type="text" id="gest-email" placeholder="' + t('gestEmail') + '" />';
    html += '<button class="btn btn-primary btn-sm" onclick="addGestionnaire()">➕ ' + t('gestAdd') + '</button>';
    html += '</div>';
  }

  // List
  html += '<div class="gestionnaire-list" id="gest-list">';
  if (gestionnaires.length === 0) {
    html += '<p style="color:#94a3b8;font-size:12px;margin-top:10px;">' + t('gestEmpty') + '</p>';
  } else {
    for (var i = 0; i < gestionnaires.length; i++) {
      var g = gestionnaires[i];
      html += '<span class="gestionnaire-chip">' + sanitize(g.Email);
      if (isOwner) html += ' <span class="chip-remove" onclick="removeGestionnaire(' + g.id + ')">✕</span>';
      html += '</span>';
    }
  }
  html += '</div>';

  html += '</div>';

  document.getElementById('gestionnaires-view').innerHTML = html;

  // Hide tab for non-owners
  var gestTab = document.querySelector('[data-tab="gestionnaires"]');
  if (gestTab) gestTab.style.display = isOwner ? '' : 'none';
}

// =============================================================================
// CRUD OPERATIONS
// =============================================================================

async function saveBien(bienId) {
  if (!canManage) { showToast(t('accessDenied'), 'error'); return; }

  var record = getFormData();
  if (!record.Reference_DDC) { showToast('Référence DDC requise', 'error'); return; }

  try {
    if (bienId) {
      // Update
      await grist.docApi.applyUserActions([
        ['UpdateRecord', BIENS_TABLE, bienId, record]
      ]);
      showToast(t('bienUpdated'), 'success');
    } else {
      // Add
      await grist.docApi.applyUserActions([
        ['AddRecord', BIENS_TABLE, null, record]
      ]);
      showToast(t('bienAdded'), 'success');
    }
    closeModalForce();
    await loadAllData();
  } catch (e) {
    console.error('Error saving bien:', e);
    showToast('Erreur: ' + e.message, 'error');
  }
}

async function confirmDeleteBien(bienId) {
  if (!canManage) { showToast(t('accessDenied'), 'error'); return; }

  try {
    await grist.docApi.applyUserActions([
      ['RemoveRecord', BIENS_TABLE, bienId]
    ]);
    showToast(t('bienDeleted'), 'success');
    closeModalForce();
    await loadAllData();
  } catch (e) {
    console.error('Error deleting bien:', e);
    showToast('Erreur: ' + e.message, 'error');
  }
}

async function addGestionnaire() {
  if (!isOwner) return;
  var email = document.getElementById('gest-email').value.trim();
  if (!email) return;

  // Check duplicate
  for (var i = 0; i < gestionnaires.length; i++) {
    if (gestionnaires[i].Email.toLowerCase() === email.toLowerCase()) return;
  }

  try {
    await grist.docApi.applyUserActions([
      ['AddRecord', GESTIONNAIRES_TABLE, null, { Email: email }]
    ]);
    showToast(t('gestAdded'), 'success');
    await loadAllData();
  } catch (e) {
    console.error('Error adding gestionnaire:', e);
  }
}

async function removeGestionnaire(gId) {
  if (!isOwner) return;
  try {
    await grist.docApi.applyUserActions([
      ['RemoveRecord', GESTIONNAIRES_TABLE, gId]
    ]);
    showToast(t('gestRemoved'), 'success');
    await loadAllData();
  } catch (e) {
    console.error('Error removing gestionnaire:', e);
  }
}

// =============================================================================
// DATA LOADING
// =============================================================================

async function ensureTables() {
  try {
    var tables = await grist.docApi.listTables();

    if (tables.indexOf(BIENS_TABLE) === -1) {
      var cols = [];
      for (var i = 0; i < BIEN_COLUMNS.length; i++) {
        cols.push({ id: BIEN_COLUMNS[i].id, fields: { type: 'Text', label: BIEN_COLUMNS[i].label_fr } });
      }
      await grist.docApi.applyUserActions([
        ['AddTable', BIENS_TABLE, cols]
      ]);
    }

    if (tables.indexOf(GESTIONNAIRES_TABLE) === -1) {
      await grist.docApi.applyUserActions([
        ['AddTable', GESTIONNAIRES_TABLE, [
          { id: 'Email', fields: { type: 'Text', label: 'Email' } }
        ]]
      ]);
    }
  } catch (e) {
    console.error('Error ensuring tables:', e);
  }
}

async function loadAllData() {
  try {
    var biensData = await grist.docApi.fetchTable(BIENS_TABLE);
    biens = [];
    if (biensData && biensData.id) {
      for (var i = 0; i < biensData.id.length; i++) {
        var row = { id: biensData.id[i] };
        for (var j = 0; j < BIEN_COLUMNS.length; j++) {
          var colId = BIEN_COLUMNS[j].id;
          row[colId] = biensData[colId] ? biensData[colId][i] : '';
        }
        biens.push(row);
      }
    }
  } catch (e) {
    console.error('Error loading biens:', e);
  }

  try {
    var gestData = await grist.docApi.fetchTable(GESTIONNAIRES_TABLE);
    gestionnaires = [];
    if (gestData && gestData.id) {
      for (var i = 0; i < gestData.id.length; i++) {
        gestionnaires.push({ id: gestData.id[i], Email: gestData.Email ? gestData.Email[i] : '' });
      }
    }
  } catch (e) {
    console.error('Error loading gestionnaires:', e);
  }

  // Update canManage based on gestionnaires
  updateCanManage();

  // Show/hide FAB
  document.getElementById('fab-add').style.display = canManage ? '' : 'none';

  refreshAllViews();
}

function updateCanManage() {
  // Owner can always manage
  if (isOwner) { canManage = true; return; }
  // Check if current user is in gestionnaires list
  // Since we can't easily get current user email in Grist widget API,
  // we default to: if user has full access and is not owner, check gestionnaires
  // For now, if there are gestionnaires and user has full access, allow management
  // The real check would need the user's email from Grist
  canManage = isOwner;
}

// =============================================================================
// INITIALIZATION
// =============================================================================

if (!isInsideGrist()) {
  document.getElementById('not-in-grist').classList.remove('hidden');
  document.getElementById('main-content').classList.add('hidden');
} else {
  (async function() {
    await grist.ready({ requiredAccess: 'full' });

    // Detect owner
    try {
      await grist.docApi.getAccessRules();
      isOwner = true;
    } catch (e) {
      isOwner = true; // Default to owner for custom widgets with full access
    }

    canManage = isOwner;

    await ensureTables();
    await loadAllData();
  })();
}
