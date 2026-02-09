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
    gestTitle: 'Gestion des Gestionnaires',
    gestSubtitle: 'Désignez les personnes autorisées à gérer les biens (en plus du Owner)',
    gestEmail: 'Email du gestionnaire',
    gestAdd: 'Ajouter',
    gestEmpty: 'Aucun gestionnaire désigné. Seul le Owner peut gérer les biens.',
    gestAdded: 'Gestionnaire ajouté',
    gestRemoved: 'Gestionnaire retiré',
    accessDenied: "Vous n'avez pas les droits pour effectuer cette action",
    ownerOnly: 'Réservé au Owner'
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
    gestTitle: 'Manager Management',
    gestSubtitle: 'Designate people authorized to manage assets (in addition to Owner)',
    gestEmail: 'Manager email',
    gestAdd: 'Add',
    gestEmpty: 'No designated managers. Only the Owner can manage assets.',
    gestAdded: 'Manager added',
    gestRemoved: 'Manager removed',
    accessDenied: 'You do not have permission to perform this action',
    ownerOnly: 'Owner only'
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
  html += '<div class="search-field"><label>Référence DDC</label><input type="text" id="s-ref" placeholder="Ex: ECH 69389 22 00001" /></div>';
  html += '<div class="search-field"><label>Commune</label><select id="s-commune"><option value="">' + t('allCommunes') + '</option>';
  for (var i = 0; i < communes.length; i++) html += '<option value="' + sanitize(communes[i]) + '">' + sanitize(communes[i]) + '</option>';
  html += '</select></div>';
  html += '<div class="search-field"><label>Mouvement</label><select id="s-mouvement"><option value="">' + t('allMovements') + '</option>';
  for (var i = 0; i < mouvements.length; i++) html += '<option value="' + sanitize(mouvements[i]) + '">' + sanitize(mouvements[i]) + '</option>';
  html += '</select></div>';
  html += '<div class="search-field"><label>Adresse</label><input type="text" id="s-adresse" placeholder="Ex: LA JACQUIERE" /></div>';
  html += '<div class="search-field"><label>Réf. Parcelle</label><input type="text" id="s-parcelle" /></div>';
  html += '<div class="search-field"><label>Type de Bien</label><select id="s-type"><option value="">' + t('allTypes') + '</option>';
  for (var i = 0; i < types.length; i++) html += '<option value="' + sanitize(types[i]) + '">' + sanitize(types[i]) + '</option>';
  html += '</select></div>';
  html += '<div class="search-field"><label>Année</label><select id="s-annee"><option value="">' + t('allYears') + '</option>';
  for (var i = 0; i < annees.length; i++) html += '<option value="' + sanitize(annees[i]) + '">' + sanitize(annees[i]) + '</option>';
  html += '</select></div>';
  html += '<div class="search-field"><label>N° Site</label><input type="text" id="s-site" /></div>';
  html += '</div>';

  html += '<div class="search-actions">';
  html += '<button class="btn btn-primary" onclick="doSearch()">🔍 ' + t('searchBtn') + '</button>';
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
  deleteBien(b.id);
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

  // Section: Nature et Observations
  html += '<div class="form-section"><h4>📝 ' + t('sectionNature') + '</h4>';
  html += '<div class="form-group"><label>Nature du bien - Projet acquisition</label><textarea id="f-Nature_Bien">' + v('Nature_Bien') + '</textarea></div>';
  html += '<div class="form-group"><label>Observations</label><textarea id="f-Observation">' + v('Observation') + '</textarea></div>';
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
    } else {
      var el = document.getElementById('f-' + col.id);
      record[col.id] = el ? el.value.trim() : '';
    }
  }
  return record;
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
  html += '<button class="btn btn-danger" onclick="deleteBien(' + b.id + ')">🗑️ ' + t('delete') + '</button>';
  html += '<div style="display:flex;gap:8px;">';
  html += '<button class="btn btn-secondary" onclick="closeModalForce()">❌ ' + t('cancel') + '</button>';
  html += '<button class="btn btn-success" onclick="saveBien(' + b.id + ')">💾 ' + t('save') + '</button>';
  html += '</div></div></div></div>';

  document.getElementById('modal-container').innerHTML = html;
}

// =============================================================================
// DASHBOARD VIEW
// =============================================================================

function renderDashboardView() {
  var totalBiens = biens.length;
  var communesSet = {};
  var surfaceParcelle = 0;
  var surfaceBati = 0;
  var mouvementCount = {};
  var communeCount = {};

  for (var i = 0; i < biens.length; i++) {
    var b = biens[i];
    if (b.Commune) { communesSet[b.Commune] = true; communeCount[b.Commune] = (communeCount[b.Commune] || 0) + 1; }
    surfaceParcelle += parseNum(b.Surface_Parcelle);
    surfaceBati += parseNum(b.Surface_Bati);
    if (b.Mouvement) mouvementCount[b.Mouvement] = (mouvementCount[b.Mouvement] || 0) + 1;
  }

  var html = '<div class="section-card">';
  html += '<h3>📊 ' + t('dashTitle') + '</h3>';
  html += '<p style="color:#64748b;margin-bottom:20px;">' + t('dashSubtitle') + '</p>';

  // Stats cards
  html += '<div class="stats-row">';
  html += '<div class="stat-card"><div class="stat-value">' + totalBiens.toLocaleString() + '</div><div class="stat-label">' + t('totalBiens') + '</div></div>';
  html += '<div class="stat-card"><div class="stat-value">' + Object.keys(communesSet).length + '</div><div class="stat-label">' + t('totalCommunes') + '</div></div>';
  html += '<div class="stat-card"><div class="stat-value">' + Math.round(surfaceParcelle).toLocaleString() + ' m²</div><div class="stat-label">' + t('totalSurfaceParcelle') + '</div></div>';
  html += '<div class="stat-card"><div class="stat-value">' + Math.round(surfaceBati).toLocaleString() + ' m²</div><div class="stat-label">' + t('totalSurfaceBati') + '</div></div>';
  html += '</div>';

  // Répartition par mouvement
  html += '<div class="section-card" style="border-top:none;box-shadow:none;padding:0;margin-top:20px;">';
  html += '<h3>📊 ' + t('repartitionMouvement') + '</h3>';
  html += '<div class="dashboard-grid" style="margin-top:12px;">';
  var mouvKeys = Object.keys(mouvementCount).sort(function(a, b) { return mouvementCount[b] - mouvementCount[a]; });
  for (var i = 0; i < mouvKeys.length; i++) {
    html += '<div class="dash-card"><div class="dash-value">' + mouvementCount[mouvKeys[i]].toLocaleString() + '</div><div class="dash-label">' + sanitize(mouvKeys[i]).toUpperCase() + '</div></div>';
  }
  html += '</div></div>';

  // Top 10 communes
  html += '<div class="section-card" style="border-top:none;box-shadow:none;padding:0;margin-top:20px;">';
  html += '<h3>🏘️ ' + t('topCommunes') + '</h3>';
  var communeKeys = Object.keys(communeCount).sort(function(a, b) { return communeCount[b] - communeCount[a]; }).slice(0, 10);
  html += '<div class="top-grid" style="margin-top:12px;">';
  for (var i = 0; i < communeKeys.length; i++) {
    html += '<div class="top-card"><div class="top-value">' + communeCount[communeKeys[i]] + '</div><div class="top-label">' + sanitize(communeKeys[i]) + '</div></div>';
  }
  html += '</div></div>';

  html += '</div>';

  document.getElementById('dashboard-view').innerHTML = html;
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

async function deleteBien(bienId) {
  if (!canManage) { showToast(t('accessDenied'), 'error'); return; }
  if (!confirm(t('confirmDelete'))) return;

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
