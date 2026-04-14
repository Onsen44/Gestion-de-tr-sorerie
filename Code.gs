const APP_TITLE = 'Gestion de tresorerie ONSEN';
const SPREADSHEET_ID = '1ay6HV2kqnyaxpLHw0UlwOfTtKDFptFgi7Qy6fpivzAU';

const DRIVE_FACTURES_ROOTS = {
  'ALTERACIG': '1UQHgnj0PKRKNhW4Mjykb8tDC6wxcz2Sj',
  'OSLO SHOP': '1phsyiWf81dhA04D0SnmDl0jqCIBbKGwt',
  'LCDV': '1mG0gk_4KOZ_CaSM0DkwXpLLhjuS72urq',
  'VAPOTE & COMPAGNIE': '1yWlrtgiihmck8Z1MlmZhdpeD75mrQ4jv',
  'GROUPE ONSEN': '17bGzBjyEm-QUnLgQgpXw6VsXaLjFPZe8'
};

const DOCUMENT_AI = {
  projectId: '',
  location: 'eu',
  processorId: ''
};

const DEFAULT_SETTINGS = {
  alertLowBalance: 5000,
  alertNegativeBalance: 0,
  alertVarianceAmount: 500,
  alertVariancePercent: 20,
  classificationConfidenceThreshold: 70,
  reconciliationConfidenceThreshold: 75
};

const SHEETS = {
  MOUVEMENTS: 'Mouvements',
  COMPTES: 'Comptes',
  CATEGORIES: 'Categories',
  REGLES: 'Regles',
  A_CLASSER: 'Transactions_A_Classer',
  FACTURES: 'Factures',
  FACTURES_A_VERIFIER: 'Factures_A_Verifier',
  PREVISIONS: 'Previsions',
  SCENARIOS: 'Scenarios',
  RAPPROCHEMENTS: 'Rapprochements',
  LOGS: 'Logs_Imports'
};

const HEADERS = {
  Mouvements: [
    'Date', 'Date_valeur', 'Libelle', 'Libelle_normalise', 'Entite', 'Groupe', 'Compte', 'Compte_id', 'Banque',
    'Categorie', 'Sous_categorie', 'Mode_reglement', 'Sens', 'Encaissement', 'Decaissement', 'Montant', 'Devise',
    'Solde_cumule', 'Reference_bancaire', 'Source_import', 'Import_id', 'Transaction_uid', 'Statut_classement',
    'Statut_rapprochement', 'Commentaire', 'Date_creation', 'Date_maj'
  ],
  Comptes: [
    'Compte_id', 'Entite', 'Groupe', 'Nom_compte', 'Banque', 'RIB_IBAN', 'Devise', 'Actif', 'Type_entite', 'Commentaire'
  ],
  Categories: [
    'Categorie', 'Sous_categorie', 'Type_flux', 'Frequence', 'Active', 'Commentaire'
  ],
  Regles: [
    'Regle_id', 'Active', 'Entite', 'Compte_id', 'Champ_source', 'Operateur', 'Valeur', 'Categorie', 'Sous_categorie',
    'Sens', 'Priorite', 'Frequence', 'Derniere_utilisation', 'Commentaire'
  ],
  Transactions_A_Classer: [
    'Transaction_uid', 'Date', 'Libelle', 'Entite', 'Compte', 'Montant', 'Sens', 'Categorie_proposee',
    'Sous_categorie_proposee', 'Raison', 'Statut', 'Commentaire'
  ],
  Factures: [
    'Facture_id', 'Entite', 'Groupe', 'Annee', 'Mois', 'Fournisseur', 'Numero_facture', 'Date_facture',
    'Date_echeance', 'Montant_TTC', 'Devise', 'Categorie', 'Sous_categorie', 'Lien_drive', 'Nom_fichier',
    'Statut_lecture', 'Statut_rapprochement', 'Statut_paiement', 'Date_paiement', 'Mode_paiement',
    'Transaction_uid_liee', 'Commentaire'
  ],
  Factures_A_Verifier: [
    'Facture_id', 'Entite', 'Nom_fichier', 'Lien_drive', 'Probleme', 'Statut', 'Commentaire'
  ],
  Previsions: [
    'Prevision_id', 'Entite', 'Compte_id', 'Categorie', 'Sous_categorie', 'Date_prevue', 'Periode', 'Frequence',
    'Sens', 'Montant_prevu', 'Mode_calcul', 'Base_reference', 'Active', 'Commentaire'
  ],
  Scenarios: [
    'Scenario_id', 'Nom_scenario', 'Entite', 'Categorie', 'Sous_categorie', 'Type_ajustement', 'Valeur_ajustement',
    'Montant_cible', 'Date_effet', 'Active', 'Commentaire'
  ],
  Rapprochements: [
    'Rapprochement_id', 'Transaction_uid', 'Facture_id', 'Entite', 'Montant_transaction', 'Montant_facture',
    'Ecart', 'Score_confiance', 'Statut', 'Commentaire', 'Date_creation'
  ],
  Logs_Imports: [
    'Import_id', 'Date_import', 'Type_import', 'Entite', 'Compte_id', 'Nom_fichier', 'Nombre_lignes',
    'Nouvelles_lignes', 'Doublons', 'A_classer', 'Statut', 'Commentaire'
  ]
};

const DEFAULT_COMPTES = [
  ['GROUPE-ONSEN-001', 'GROUPE ONSEN', 'GROUPE ONSEN', 'Compte principal Groupe Onsen', '', '', 'EUR', 'Oui', 'groupe', 'A completer'],
  ['ALTERACIG-001', 'ALTERACIG', 'GROUPE ONSEN', 'Compte principal ALTERACIG', '', '', 'EUR', 'Oui', 'societe', 'A completer'],
  ['OSLO-001', 'OSLO SHOP', 'GROUPE ONSEN', 'Compte principal OSLO SHOP', '', '', 'EUR', 'Oui', 'societe', 'A completer'],
  ['VAPOTE-001', 'VAPOTE & COMPAGNIE', 'GROUPE ONSEN', 'Compte principal VAPOTE & COMPAGNIE', '', '', 'EUR', 'Oui', 'societe', 'A completer'],
  ['LCDV-001', 'LCDV', 'GROUPE ONSEN', 'Compte principal LCDV', '', '', 'EUR', 'Oui', 'societe', 'A completer']
];

const DEFAULT_CATEGORIES = [
  ['Securite', 'Kiwatch', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Securite', 'Verisure', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Securite', 'Delta Security', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Securite', 'Homiris', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Securite', 'EPS', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Electricite', 'EDF', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Electricite', 'Engie', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Electricite', 'Vattenfall', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Electricite', 'TotalEnergies', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Electricite', 'Octopus', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Internet', 'Bouygues', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Internet', 'Orange', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Internet', 'SFR', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Internet', 'Free', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Logiciels / SaaS', 'Adobe', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Logiciels / SaaS', 'Skello', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Logiciels / SaaS', 'Google Cloud', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Honoraires', 'VIMA', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Honoraires', 'GESBERT SARL', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Honoraires', 'Avocat', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Honoraires', 'Cabinet', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Location materiel', 'Ingenico', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Frais bancaires', 'Commissions CB', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Frais bancaires', 'Services bancaires', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Frais bancaires', 'Cotisations', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale'],
  ['Frais bancaires', 'Interets bancaires', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Assurances', 'AXA', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Assurances', 'MAAF', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Assurances', 'MMA', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Assurances', 'Allianz', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Assurances', 'Groupama', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Eau', 'Veolia', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Eau', 'Suez', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Eau', 'Saur', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Eau', 'Regie des eaux', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Eau', 'Service des eaux', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Eau', 'Syndicat des eaux', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Eau', 'Agglomeration des eaux', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Eau', 'Communaute de communes eau', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Eau', 'Eau et assainissement', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Eau', 'A definir', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Carburant', 'Total', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Carburant', 'Esso', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Carburant', 'Shell', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Carburant', 'Auchan Carburant', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Carburant', 'A definir', 'Decaissement', 'Variable', 'Oui', 'Base initiale'],
  ['Loyer', 'A definir', 'Decaissement', 'Mensuelle', 'Oui', 'Base initiale']
];


const DEFAULT_REGLES = [
  ['REG-BASE-001', 'Oui', '', '', 'Libelle_normalise', 'contains', 'KIWATCH', 'Securite', 'Kiwatch', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-002', 'Oui', '', '', 'Libelle_normalise', 'contains', 'VERISURE', 'Securite', 'Verisure', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-003', 'Oui', '', '', 'Libelle_normalise', 'contains', 'DELTA SECURITY', 'Securite', 'Delta Security', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-004', 'Oui', '', '', 'Libelle_normalise', 'contains', 'HOMIRIS', 'Securite', 'Homiris', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-005', 'Oui', '', '', 'Libelle_normalise', 'contains', 'EPS', 'Securite', 'EPS', 'Decaissement', 90, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-006', 'Oui', '', '', 'Libelle_normalise', 'contains', 'EDF', 'Electricite', 'EDF', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-007', 'Oui', '', '', 'Libelle_normalise', 'contains', 'ENGIE', 'Electricite', 'Engie', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-008', 'Oui', '', '', 'Libelle_normalise', 'contains', 'VATTENFALL', 'Electricite', 'Vattenfall', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-009', 'Oui', '', '', 'Libelle_normalise', 'contains', 'TOTALENERGIES', 'Electricite', 'TotalEnergies', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-010', 'Oui', '', '', 'Libelle_normalise', 'contains', 'OCTOPUS', 'Electricite', 'Octopus', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-011', 'Oui', '', '', 'Libelle_normalise', 'contains', 'BOUYGUES', 'Internet', 'Bouygues', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-012', 'Oui', '', '', 'Libelle_normalise', 'contains', 'ORANGE', 'Internet', 'Orange', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-013', 'Oui', '', '', 'Libelle_normalise', 'contains', 'SFR', 'Internet', 'SFR', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-014', 'Oui', '', '', 'Libelle_normalise', 'contains', 'FREE', 'Internet', 'Free', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-015', 'Oui', '', '', 'Libelle_normalise', 'contains', 'ADOBE', 'Logiciels / SaaS', 'Adobe', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-016', 'Oui', '', '', 'Libelle_normalise', 'contains', 'SKELLO', 'Logiciels / SaaS', 'Skello', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-017', 'Oui', '', '', 'Libelle_normalise', 'contains', 'GOOGLE CLOUD', 'Logiciels / SaaS', 'Google Cloud', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-018', 'Oui', '', '', 'Libelle_normalise', 'contains', 'GOOGLE', 'Logiciels / SaaS', 'Google Cloud', 'Decaissement', 80, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-019', 'Oui', '', '', 'Libelle_normalise', 'contains', 'VIMA', 'Honoraires', 'VIMA', 'Decaissement', 100, 'Variable', '', 'Base initiale'],
  ['REG-BASE-020', 'Oui', '', '', 'Libelle_normalise', 'contains', 'GESBERT', 'Honoraires', 'GESBERT SARL', 'Decaissement', 100, 'Variable', '', 'Base initiale'],
  ['REG-BASE-021', 'Oui', '', '', 'Libelle_normalise', 'contains', 'AVOCAT', 'Honoraires', 'Avocat', 'Decaissement', 90, 'Variable', '', 'Base initiale'],
  ['REG-BASE-022', 'Oui', '', '', 'Libelle_normalise', 'contains', 'CABINET', 'Honoraires', 'Cabinet', 'Decaissement', 70, 'Variable', '', 'Base initiale'],
  ['REG-BASE-023', 'Oui', '', '', 'Libelle_normalise', 'contains', 'INGENICO', 'Location materiel', 'Ingenico', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-024', 'Oui', '', '', 'Libelle_normalise', 'contains', 'SUMUP', 'Frais bancaires', 'Commissions CB', 'Decaissement', 100, 'Variable', '', 'Base initiale'],
  ['REG-BASE-025', 'Oui', '', '', 'Libelle_normalise', 'contains', 'WORLDLINE', 'Frais bancaires', 'Services bancaires', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-026', 'Oui', '', '', 'Libelle_normalise', 'contains', 'MYPOS', 'Frais bancaires', 'Services bancaires', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-027', 'Oui', '', '', 'Libelle_normalise', 'contains', 'AXA', 'Assurances', 'AXA', 'Decaissement', 100, 'Variable', '', 'Base initiale'],
  ['REG-BASE-028', 'Oui', '', '', 'Libelle_normalise', 'contains', 'MAAF', 'Assurances', 'MAAF', 'Decaissement', 100, 'Variable', '', 'Base initiale'],
  ['REG-BASE-029', 'Oui', '', '', 'Libelle_normalise', 'contains', 'MMA', 'Assurances', 'MMA', 'Decaissement', 100, 'Variable', '', 'Base initiale'],
  ['REG-BASE-030', 'Oui', '', '', 'Libelle_normalise', 'contains', 'ALLIANZ', 'Assurances', 'Allianz', 'Decaissement', 100, 'Variable', '', 'Base initiale'],
  ['REG-BASE-031', 'Oui', '', '', 'Libelle_normalise', 'contains', 'GROUPAMA', 'Assurances', 'Groupama', 'Decaissement', 100, 'Variable', '', 'Base initiale'],
  ['REG-BASE-032', 'Oui', '', '', 'Libelle_normalise', 'contains', 'VEOLIA', 'Eau', 'Veolia', 'Decaissement', 100, 'Variable', '', 'Base initiale'],
  ['REG-BASE-033', 'Oui', '', '', 'Libelle_normalise', 'contains', 'SUEZ', 'Eau', 'Suez', 'Decaissement', 100, 'Variable', '', 'Base initiale'],
  ['REG-BASE-034', 'Oui', '', '', 'Libelle_normalise', 'contains', 'SAUR', 'Eau', 'Saur', 'Decaissement', 100, 'Variable', '', 'Base initiale'],
  ['REG-BASE-035', 'Oui', '', '', 'Libelle_normalise', 'contains', 'REGIE DES EAUX', 'Eau', 'Regie des eaux', 'Decaissement', 95, 'Variable', '', 'Base initiale'],
  ['REG-BASE-036', 'Oui', '', '', 'Libelle_normalise', 'contains', 'SERVICE DES EAUX', 'Eau', 'Service des eaux', 'Decaissement', 95, 'Variable', '', 'Base initiale'],
  ['REG-BASE-037', 'Oui', '', '', 'Libelle_normalise', 'contains', 'SYNDICAT DES EAUX', 'Eau', 'Syndicat des eaux', 'Decaissement', 95, 'Variable', '', 'Base initiale'],
  ['REG-BASE-038', 'Oui', '', '', 'Libelle_normalise', 'contains', 'AGGLOMERATION DES EAUX', 'Eau', 'Agglomeration des eaux', 'Decaissement', 90, 'Variable', '', 'Base initiale'],
  ['REG-BASE-039', 'Oui', '', '', 'Libelle_normalise', 'contains', 'COMMUNAUTE DE COMMUNES EAU', 'Eau', 'Communaute de communes eau', 'Decaissement', 90, 'Variable', '', 'Base initiale'],
  ['REG-BASE-040', 'Oui', '', '', 'Libelle_normalise', 'contains', 'EAU ET ASSAINISSEMENT', 'Eau', 'Eau et assainissement', 'Decaissement', 90, 'Variable', '', 'Base initiale'],
  ['REG-BASE-041', 'Oui', '', '', 'Libelle_normalise', 'contains', 'FRAIS BANCAIRES', 'Frais bancaires', 'Services bancaires', 'Decaissement', 100, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-042', 'Oui', '', '', 'Libelle_normalise', 'contains', 'COM CB', 'Frais bancaires', 'Commissions CB', 'Decaissement', 100, 'Variable', '', 'Base initiale'],
  ['REG-BASE-043', 'Oui', '', '', 'Libelle_normalise', 'contains', 'COTISATION', 'Frais bancaires', 'Cotisations', 'Decaissement', 80, 'Mensuelle', '', 'Base initiale'],
  ['REG-BASE-044', 'Oui', '', '', 'Libelle_normalise', 'contains', 'INTERETS BANCAIRES', 'Frais bancaires', 'Interets bancaires', 'Decaissement', 90, 'Variable', '', 'Base initiale'],
  ['REG-BASE-045', 'Oui', '', '', 'Libelle_normalise', 'contains', 'TOTAL', 'Carburant', 'Total', 'Decaissement', 70, 'Variable', '', 'Base initiale'],
  ['REG-BASE-046', 'Oui', '', '', 'Libelle_normalise', 'contains', 'ESSO', 'Carburant', 'Esso', 'Decaissement', 100, 'Variable', '', 'Base initiale'],
  ['REG-BASE-047', 'Oui', '', '', 'Libelle_normalise', 'contains', 'SHELL', 'Carburant', 'Shell', 'Decaissement', 100, 'Variable', '', 'Base initiale'],
  ['REG-BASE-048', 'Oui', '', '', 'Libelle_normalise', 'contains', 'CARBURANT', 'Carburant', 'A definir', 'Decaissement', 60, 'Variable', '', 'Base initiale'],
  ['REG-BASE-049', 'Oui', '', '', 'Libelle_normalise', 'contains', 'AUCHAN CARBURANT', 'Carburant', 'Auchan Carburant', 'Decaissement', 100, 'Variable', '', 'Base initiale'],
  ['REG-BASE-050', 'Oui', '', '', 'Libelle_normalise', 'contains', 'CARTE CARBURANT', 'Carburant', 'A definir', 'Decaissement', 70, 'Variable', '', 'Base initiale']
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAppBootstrap() {
  installOrUpdateWorkbook();
  const ss = getSpreadsheet_();
  return {
    appTitle: APP_TITLE,
    spreadsheetId: SPREADSHEET_ID,
    spreadsheetUrl: ss.getUrl(),
    counts: getSheetCounts_(),
    dashboard: buildDashboardSummary_(),
    menu: ['Accueil', 'Gestion de tresorerie', 'Graphiques', 'Previsionnel', 'Scenarios', 'Rapprochement bancaire', 'Parametres'],
    config: buildConfigPayload_(),
    aClasser: getTransactionsToClassify(50),
    recentMovements: getRecentMovements(20),
    recentPrevisions: getRecentPrevisions(50),
    recentScenarios: getRecentScenarios(50),
    recentRapprochements: getRecentRapprochements(50),
      facturesToVerify: getFacturesToVerify(50),
      transactionsWithoutInvoice: getTransactionsWithoutInvoice(50),
      alerts: getAlerts_(12),
      forecastSummary: buildForecastSummary_(),
      scenarioSummary: buildScenarioSummary_(),
      periodAnalytics: buildPeriodAnalytics_(),
    graphData: getGraphData_(),
    driveConfig: getDriveRoots_(),
    facturesToPay: getFacturesToPay(100),
    automaticSync: getAutomaticSyncStatus(),
    exportData: getExportPayload_(),
    settings: getSettingsConfig_(),
    ocrReady: isOcrConfigured_(),
    ocrConfig: getDocumentAiConfig_()
  };
}

function installOrUpdateWorkbook() {
  // Initialiser les IDs sécurisés
  initializeSecureIds_();

  const ss = getSpreadsheet_();
  Object.keys(SHEETS).forEach(function(key) {
    ensureSheetWithHeaders_(ss, SHEETS[key], HEADERS[SHEETS[key]]);
  });
  seedDefaults_();
  formatWorkbook_(ss);
  return { success: true, message: 'Google Sheet installe et synchronise.', spreadsheetUrl: ss.getUrl() };
}

function importBankRows(payload) {
  installOrUpdateWorkbook();
  payload = payload || {};
  const rows = payload.rows || [];
  const mapping = payload.mapping || {};
  const compteId = String(payload.compteId || '').trim();
  const entite = String(payload.entite || '').trim();
  const sourceName = String(payload.sourceName || 'import').trim();

  if (!rows.length) throw new Error('Aucune ligne a importer.');
  if (!compteId) throw new Error('Compte bancaire manquant.');
  if (!entite) throw new Error('Entite manquante.');
  if (!mapping || !Object.keys(mapping).length) throw new Error('Mapping vide - configure les colonnes SVP.');
  if (!mapping.date || !mapping.libelle) throw new Error('Mapping incomplet: date et libelle sont obligatoires.');
  if (!mapping.montant && !mapping.debit && !mapping.credit) throw new Error('Mapping incomplet: montant OU (debit ET credit) sont obligatoires.');

  const ss = getSpreadsheet_();
  const compte = getComptesMap_()[compteId];
  if (!compte) throw new Error('Compte introuvable dans l\'onglet Comptes: ' + compteId);

  const mouvementsSheet = ss.getSheetByName(SHEETS.MOUVEMENTS);
  const aClasserSheet = ss.getSheetByName(SHEETS.A_CLASSER);
  const logSheet = ss.getSheetByName(SHEETS.LOGS);
  const existingUids = getExistingUids_();
  const rules = getActiveRules_();
  const importId = 'IMP-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Europe/Paris', 'yyyyMMdd-HHmmss');
  const now = new Date();

  const mouvementRows = [];
  const aClasserRows = [];
  let duplicates = 0;
  let inserted = 0;

  let errors = 0;
  rows.forEach(function(sourceRow, index) {
    try {
      const tx = normalizeImportedRow_(sourceRow, mapping, compte, entite, sourceName, importId);
      if (!tx || !tx.Date || !tx.Libelle || tx.Montant === 0) return;
      if (existingUids[tx.Transaction_uid]) {
        duplicates++;
        return;
      }
      existingUids[tx.Transaction_uid] = true;

      const ruleMatch = findMatchingRule_(tx, rules);
      if (ruleMatch) {
        tx.Categorie = ruleMatch.Categorie || '';
        tx.Sous_categorie = ruleMatch.Sous_categorie || '';
        tx.Statut_classement = 'CLASSE';
      } else {
        tx.Statut_classement = 'A_CLASSER';
        aClasserRows.push([
          tx.Transaction_uid, tx.Date, tx.Libelle, tx.Entite, tx.Compte, Math.abs(tx.Montant), tx.Sens,
          '', '', 'Aucune regle fiable trouvee', 'A_TRAITER', ''
        ]);
      }

      mouvementRows.push([
        tx.Date, tx.Date_valeur, tx.Libelle, tx.Libelle_normalise, tx.Entite, tx.Groupe, tx.Compte, tx.Compte_id, tx.Banque,
        tx.Categorie, tx.Sous_categorie, tx.Mode_reglement, tx.Sens, tx.Encaissement, tx.Decaissement, tx.Montant, tx.Devise,
        tx.Solde_cumule, tx.Reference_bancaire, tx.Source_import, tx.Import_id, tx.Transaction_uid, tx.Statut_classement,
        tx.Statut_rapprochement, tx.Commentaire, now, now
      ]);
      inserted++;
    } catch (error) {
      errors++;
      Logger.log('Erreur ligne ' + (index + 1) + ': ' + error.message);
    }
  });

  if (mouvementRows.length) {
    mouvementsSheet.getRange(mouvementsSheet.getLastRow() + 1, 1, mouvementRows.length, mouvementRows[0].length).setValues(mouvementRows);
  }
  if (aClasserRows.length) {
    aClasserSheet.getRange(aClasserSheet.getLastRow() + 1, 1, aClasserRows.length, aClasserRows[0].length).setValues(aClasserRows);
  }

  Logger.log('Import bancaire: ' + rows.length + ' lignes, ' + inserted + ' insérées, ' + duplicates + ' doublons, ' + errors + ' erreurs');

  logSheet.appendRow([
    importId, now, 'BANQUE', entite, compteId, sourceName, rows.length, inserted, duplicates, aClasserRows.length, errors > 0 ? 'ERREURS' : 'OK', errors + ' erreurs'
  ]);

  const reconciliation = runAutomaticReconciliation({ entites: [entite], source: 'IMPORT_BANQUE' });

  return {
    success: true,
    importId: importId,
    totalRows: rows.length,
    inserted: inserted,
    duplicates: duplicates,
    toClassify: aClasserRows.length,
    reconciliation: reconciliation,
    dashboard: buildDashboardSummary_(),
    aClasser: getTransactionsToClassify(50),
    recentMovements: getRecentMovements(20),
    recentPrevisions: getRecentPrevisions(50),
    recentRapprochements: getRecentRapprochements(50),
    facturesToVerify: getFacturesToVerify(50),
    transactionsWithoutInvoice: getTransactionsWithoutInvoice(50),
    alerts: getAlerts_(12),
    forecastSummary: buildForecastSummary_(),
    scenarioSummary: buildScenarioSummary_(),
    periodAnalytics: buildPeriodAnalytics_(),
    graphData: getGraphData_(),
    driveConfig: getDriveRoots_(),
    facturesToPay: getFacturesToPay(100),
    automaticSync: getAutomaticSyncStatus(),
    counts: getSheetCounts_()
  };
}

function classifyTransaction(payload) {
  installOrUpdateWorkbook();
  payload = payload || {};
  const transactionUid = String(payload.transactionUid || '').trim();
  const categorie = String(payload.categorie || '').trim();
  const sousCategorie = String(payload.sousCategorie || '').trim();
  const createRule = payload.createRule === true;
  const ruleValue = String(payload.ruleValue || '').trim();
  const frequence = String(payload.frequence || '').trim() || 'Variable';

  if (!transactionUid) throw new Error('Transaction manquante.');
  if (!categorie) throw new Error('Categorie manquante.');

  const ss = getSpreadsheet_();
  const mouvementsSheet = ss.getSheetByName(SHEETS.MOUVEMENTS);
  const aClasserSheet = ss.getSheetByName(SHEETS.A_CLASSER);
  const reglesSheet = ss.getSheetByName(SHEETS.REGLES);

  const mInfo = findRowByValue_(mouvementsSheet, 'Transaction_uid', transactionUid);
  if (!mInfo.rowNumber) throw new Error('Transaction introuvable dans Mouvements.');

  const headers = getHeaders_(mouvementsSheet);
  const rowValues = mouvementsSheet.getRange(mInfo.rowNumber, 1, 1, headers.length).getValues()[0];
  const rowObj = rowToObject_(headers, rowValues);

  setCellByHeader_(mouvementsSheet, headers, mInfo.rowNumber, 'Categorie', categorie);
  setCellByHeader_(mouvementsSheet, headers, mInfo.rowNumber, 'Sous_categorie', sousCategorie);
  setCellByHeader_(mouvementsSheet, headers, mInfo.rowNumber, 'Statut_classement', 'CLASSE');
  setCellByHeader_(mouvementsSheet, headers, mInfo.rowNumber, 'Date_maj', new Date());

  const aHeaders = getHeaders_(aClasserSheet);
  const aInfo = findRowByValue_(aClasserSheet, 'Transaction_uid', transactionUid);
  if (aInfo.rowNumber) {
    setCellByHeader_(aClasserSheet, aHeaders, aInfo.rowNumber, 'Statut', 'TRAITE');
    setCellByHeader_(aClasserSheet, aHeaders, aInfo.rowNumber, 'Categorie_proposee', categorie);
    setCellByHeader_(aClasserSheet, aHeaders, aInfo.rowNumber, 'Sous_categorie_proposee', sousCategorie);
  }

  if (createRule && ruleValue) {
    reglesSheet.appendRow([
      'REG-' + Utilities.getUuid().slice(0, 8), 'Oui', rowObj.Entite || '', rowObj.Compte_id || '', 'Libelle_normalise', 'contains',
      normalizeText_(ruleValue), categorie, sousCategorie, rowObj.Sens || '', 100, frequence, '', 'Regle creee depuis A classer'
    ]);
  }

  return {
    success: true,
    aClasser: getTransactionsToClassify(50),
    recentMovements: getRecentMovements(20),
    recentPrevisions: getRecentPrevisions(50),
    counts: getSheetCounts_(),
    dashboard: buildDashboardSummary_(),
    recentRapprochements: getRecentRapprochements(50),
    facturesToVerify: getFacturesToVerify(50),
    transactionsWithoutInvoice: getTransactionsWithoutInvoice(50),
    alerts: getAlerts_(12),
    forecastSummary: buildForecastSummary_(),
    scenarioSummary: buildScenarioSummary_(),
    periodAnalytics: buildPeriodAnalytics_(),
    graphData: getGraphData_(),
    driveConfig: getDriveRoots_(),
    facturesToPay: getFacturesToPay(100),
    automaticSync: getAutomaticSyncStatus()
  };
}

function syncDriveInvoices(payload) {
  installOrUpdateWorkbook();
  payload = payload || {};
  const rootMap = getDriveRootMap_();
  const entites = (payload.entites && payload.entites.length ? payload.entites : Object.keys(rootMap))
    .filter(function(entite) { return rootMap[entite]; });

  const ss = getSpreadsheet_();
  const facturesSheet = ss.getSheetByName(SHEETS.FACTURES);
  const facturesVerifierSheet = ss.getSheetByName(SHEETS.FACTURES_A_VERIFIER);
  const logsSheet = ss.getSheetByName(SHEETS.LOGS);
  const now = new Date();

  let inserted = 0;
  let verifier = 0;
  let scannedFiles = 0;

  entites.forEach(function(entite) {
    const rootId = rootMap[entite];
    if (!rootId) return;

    const existingFactureIds = getExistingFactureIds_(entite);
    const rootFolder = DriveApp.getFolderById(rootId);
    const yearFolders = rootFolder.getFolders();
    let entiteScanned = 0;
    let entiteInserted = 0;
    let entiteVerifier = 0;

    while (yearFolders.hasNext()) {
      const yearFolder = yearFolders.next();
      const yearName = String(yearFolder.getName() || '').trim();
      if (!/^\d{4}$/.test(yearName)) continue;

      const monthFolders = yearFolder.getFolders();
      while (monthFolders.hasNext()) {
        const monthFolder = monthFolders.next();
        const monthName = String(monthFolder.getName() || '').trim();
        const monthLower = normalizeText_(monthName);
        const validMonths = ['janvier', 'fevrier', 'mars', 'avril', 'mai', 'juin', 'juillet', 'aout', 'septembre', 'octobre', 'novembre', 'decembre'];
        const isValidMonth = validMonths.some(function(m) { return monthLower.indexOf(m) !== -1; }) || /^\d{1,2}$/.test(monthName);
        if (!isValidMonth) continue;
        const files = monthFolder.getFiles();

        while (files.hasNext()) {
          const file = files.next();
          entiteScanned++;
          scannedFiles++;
          const fileId = file.getId();
          if (existingFactureIds[fileId]) continue;
          existingFactureIds[fileId] = true;

          const fileName = file.getName();
          const parsed = parseInvoiceFileName_(fileName);
          const factureId = 'FAC-' + fileId;

          facturesSheet.appendRow([
            factureId,
            entite,
            'GROUPE ONSEN',
            yearName,
            monthName,
            parsed.fournisseur,
            parsed.numero,
            '',
            '',
            '',
            'EUR',
            parsed.categorie,
            parsed.sousCategorie,
            file.getUrl(),
            fileName,
            parsed.statutLecture,
            'NON_RAPPROCHE',
            'A_REGLER',
            '',
            '',
            '',
            'Indexe depuis Google Drive en lecture seule'
          ]);
          inserted++;
          entiteInserted++;

          if (parsed.statutLecture !== 'OK') {
            facturesVerifierSheet.appendRow([
              factureId,
              entite,
              fileName,
              file.getUrl(),
              parsed.probleme,
              'A_VERIFIER',
              ''
            ]);
            verifier++;
            entiteVerifier++;
          }
        }
      }
    }

    logsSheet.appendRow([
      'DRV-' + Utilities.getUuid().slice(0, 8),
      now,
      'DRIVE_FACTURES',
      entite,
      '',
      rootFolder.getName(),
      entiteScanned,
      entiteInserted,
      0,
      entiteVerifier,
      'OK',
      'Indexation Drive en lecture seule'
    ]);
  });

  const reconciliation = runAutomaticReconciliation({ entites: entites, source: 'SYNC_DRIVE' });

  return {
    success: true,
    scannedFiles: scannedFiles,
    inserted: inserted,
    toVerify: verifier,
    reconciliation: reconciliation,
    counts: getSheetCounts_(),
    facturesToPay: getFacturesToPay(100),
    automaticSync: getAutomaticSyncStatus(),
    recentRapprochements: getRecentRapprochements(50),
    facturesToVerify: getFacturesToVerify(50),
    transactionsWithoutInvoice: getTransactionsWithoutInvoice(50),
    alerts: getAlerts_(12),
    forecastSummary: buildForecastSummary_(),
    scenarioSummary: buildScenarioSummary_(),
    periodAnalytics: buildPeriodAnalytics_()
  };
}

function updateInvoicePaymentStatus(payload) {
  installOrUpdateWorkbook();
  payload = payload || {};
  const factureId = String(payload.factureId || '').trim();
  const statutPaiement = String(payload.statutPaiement || '').trim();
  const modePaiement = String(payload.modePaiement || '').trim();
  const datePaiement = String(payload.datePaiement || '').trim();
  const transactionUid = String(payload.transactionUid || '').trim();

  if (!factureId) throw new Error('Facture manquante.');
  if (!statutPaiement) throw new Error('Statut de paiement manquant.');

  const sheet = getSpreadsheet_().getSheetByName(SHEETS.FACTURES);
  const info = findRowByValue_(sheet, 'Facture_id', factureId);
  if (!info.rowNumber) throw new Error('Facture introuvable.');

  const headers = getHeaders_(sheet);
  setCellByHeader_(sheet, headers, info.rowNumber, 'Statut_paiement', statutPaiement);
  setCellByHeader_(sheet, headers, info.rowNumber, 'Mode_paiement', modePaiement);
  setCellByHeader_(sheet, headers, info.rowNumber, 'Date_paiement', datePaiement);
  setCellByHeader_(sheet, headers, info.rowNumber, 'Transaction_uid_liee', transactionUid);
  if (transactionUid && statutPaiement === 'REGLEE') {
    setCellByHeader_(sheet, headers, info.rowNumber, 'Statut_rapprochement', 'RAPPROCHE');
    updateMovementReconciliationStatus_(transactionUid, 'RAPPROCHE');
  }

  return {
    success: true,
    facturesARegler: getFacturesToPay(100),
    recentRapprochements: getRecentRapprochements(50),
    transactionsWithoutInvoice: getTransactionsWithoutInvoice(50),
    alerts: getAlerts_(12),
    forecastSummary: buildForecastSummary_(),
    scenarioSummary: buildScenarioSummary_(),
    periodAnalytics: buildPeriodAnalytics_(),
    counts: getSheetCounts_()
  };
}

function setDocumentAiConfig(payload) {
  payload = payload || {};

  // Vérifier que seul un administrateur peut configurer Document AI
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const spreadsheet = getSpreadsheet_();
    const editors = spreadsheet.getEditors();
    const isEditor = editors.some(function(editor) { return editor.getEmail() === userEmail; });
    if (!isEditor) {
      throw new Error('Vous n\'avez pas les droits pour modifier la configuration OCR.');
    }
  } catch (error) {
    throw new Error('Configuration OCR impossible - Erreur: ' + error.message);
  }

  const props = PropertiesService.getScriptProperties();
  const config = {
    DOCUMENT_AI_PROJECT_ID: String(payload.projectId || '').trim(),
    DOCUMENT_AI_LOCATION: String(payload.location || 'eu').trim(),
    DOCUMENT_AI_PROCESSOR_ID: String(payload.processorId || '').trim()
  };

  // Valider que les IDs ne sont pas vides
  if (!config.DOCUMENT_AI_PROJECT_ID || !config.DOCUMENT_AI_PROCESSOR_ID) {
    throw new Error('Project ID et Processor ID sont obligatoires pour configurer l\'OCR.');
  }

  props.setProperties(config, true);
  Logger.log('Configuration OCR mise à jour avec Project ID: ' + config.DOCUMENT_AI_PROJECT_ID);

  return {
    success: true,
    ocrReady: isOcrConfigured_(),
    ocrConfig: getDocumentAiConfig_()
  };
}

function runInvoiceOcr(payload) {
  installOrUpdateWorkbook();

  payload = payload || {};
  const limit = Math.max(Number(payload.limit || 10), 1);
  const factures = getRows_(getSpreadsheet_().getSheetByName(SHEETS.FACTURES))
    .filter(function(row) {
      if (!row.Facture_id || String(row.Facture_id).indexOf('FAC-') !== 0) return false;
      if (payload.entites && payload.entites.length && payload.entites.indexOf(String(row.Entite || '').trim()) === -1) return false;
      const status = String(row.Statut_lecture || '').trim();
      return status !== 'OCR_OK';
    })
    .slice(0, limit);

  let processed = 0;
  let ok = 0;
  let failed = 0;

  factures.forEach(function(facture) {
    processed++;
    try {
      const fileId = String(facture.Facture_id).replace(/^FAC-/, '');
      const file = DriveApp.getFileById(fileId);
      const blob = file.getBlob();
      const parsed = processInvoiceBlobWithOcrSpace_(blob, file.getMimeType());
      applyInvoiceOcrResult_(facture.Facture_id, parsed);
      ok++;
    } catch (error) {
      failed++;
      markInvoiceForVerification_(facture.Facture_id, facture.Entite, facture.Nom_fichier, facture.Lien_drive, error.message || String(error));
    }
  });

  const reconciliation = runAutomaticReconciliation({ source: 'OCR' });

  return {
    success: true,
    processed: processed,
    ok: ok,
    failed: failed,
    reconciliation: reconciliation,
    facturesToPay: getFacturesToPay(100),
    automaticSync: getAutomaticSyncStatus(),
    recentRapprochements: getRecentRapprochements(50),
    facturesToVerify: getFacturesToVerify(50),
    transactionsWithoutInvoice: getTransactionsWithoutInvoice(50),
    alerts: getAlerts_(12),
    forecastSummary: buildForecastSummary_(),
    scenarioSummary: buildScenarioSummary_(),
    periodAnalytics: buildPeriodAnalytics_(),
    counts: getSheetCounts_(),
    ocrReady: true,
    ocrConfig: { projectId: '', location: 'eu', processorId: '' }
  };
}

function processInvoiceBlobWithOcrSpace_(blob, mimeType) {
  const base64Data = Utilities.base64Encode(blob.getBytes());
  const url = 'https://api.ocr.space/parse/image';

  const payload = {
    base64Image: 'data:' + (mimeType || blob.getContentType() || 'application/pdf') + ';base64,' + base64Data,
    language: 'fra',
    isOverlayRequired: false
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    timeout: 30000
  });

  const code = response.getResponseCode();
  const text = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('OCR.space a repondu ' + code + ' : ' + text);
  }

  const data = JSON.parse(text);
  if (data.isErroredOnProcessing) {
    throw new Error('OCR.space erreur: ' + (data.errorMessage || 'erreur inconnue'));
  }

  return parseOcrSpaceInvoice_(data.parsedText || '');
}

function parseDocumentAiInvoice_(document) {
  const entities = document.entities || [];
  const result = {
    fournisseur: '',
    numero: '',
    dateFacture: '',
    dateEcheance: '',
    montantTTC: '',
    devise: 'EUR',
    categorie: '',
    sousCategorie: '',
    statutLecture: 'OCR_OK',
    commentaire: 'Lecture OCR Document AI'
  };

  entities.forEach(function(entity) {
    const type = String(entity.type || '').toLowerCase();
    const mention = String(entity.mentionText || '').trim();
    const normalized = normalizeText_(mention);
    if (!mention) return;

    if (type === 'supplier_name' && !result.fournisseur) result.fournisseur = mention;
    if ((type === 'invoice_id' || type === 'invoice_number') && !result.numero) result.numero = mention;
    if (type === 'invoice_date' && !result.dateFacture) result.dateFacture = normalizeDocumentAiDate_(entity.normalizedValue, mention);
    if (type === 'due_date' && !result.dateEcheance) result.dateEcheance = normalizeDocumentAiDate_(entity.normalizedValue, mention);
    if ((type === 'total_amount' || type === 'amount_due') && !result.montantTTC) result.montantTTC = normalizeDocumentAiAmount_(entity.normalizedValue, mention);
    if ((type === 'currency' || type === 'currency_code') && !result.devise) result.devise = mention;

    if (!result.categorie) {
      const inferred = inferCategoryFromText_(normalized);
      if (inferred.categorie) {
        result.categorie = inferred.categorie;
        result.sousCategorie = inferred.sousCategorie;
      }
    }
  });

  if (!result.fournisseur && !result.numero && !result.montantTTC) {
    throw new Error('Aucune information exploitable extraite par OCR.');
  }

  if (!result.categorie && result.fournisseur) {
    const inferred = inferCategoryFromText_(normalizeText_(result.fournisseur));
    result.categorie = inferred.categorie || 'A verifier';
    result.sousCategorie = inferred.sousCategorie || ('Fournisseur: ' + result.fournisseur);
  }

  if (!result.sousCategorie && result.fournisseur) {
    result.sousCategorie = result.fournisseur;
  }

  return result;
}

function parseOcrSpaceInvoice_(text) {
  const result = {
    fournisseur: '',
    numero: '',
    dateFacture: '',
    dateEcheance: '',
    montantTTC: '',
    devise: 'EUR',
    categorie: '',
    sousCategorie: '',
    statutLecture: 'OCR_OK',
    commentaire: 'Lecture OCR OCR.space'
  };

  if (!text || String(text).trim().length === 0) {
    throw new Error('Aucun texte extrait par OCR.');
  }

  const lines = text.split('\n');
  const textLower = text.toLowerCase();

  // Chercher le montant (TOTAL, AMOUNT, etc.)
  const montantPatterns = [
    /(?:total|montant|amount|ttc|ht|net)\s*(?:\(|:|=)?\s*[\s\S]*?([\d\s.,]+)/i,
    /([\d.,]+)\s*(?:€|eur|dollars?|usd)/i
  ];

  for (let i = 0; i < montantPatterns.length; i++) {
    const match = text.match(montantPatterns[i]);
    if (match && match[1]) {
      const cleaned = String(match[1]).replace(/\s/g, '').replace(',', '.');
      if (/^\d+\.?\d*$/.test(cleaned) && parseFloat(cleaned) > 0) {
        result.montantTTC = cleaned;
        break;
      }
    }
  }

  // Chercher le numéro de facture
  const numeroPatterns = [
    /(?:facture|invoice|n°|no\.?|#)\s*:?\s*([^\s\n,]+)/i,
    /(?:ref|reference)\s*:?\s*([^\s\n,]+)/i
  ];

  for (let i = 0; i < numeroPatterns.length; i++) {
    const match = text.match(numeroPatterns[i]);
    if (match && match[1]) {
      const num = String(match[1]).trim();
      if (num.length > 0 && num.length < 50) {
        result.numero = num;
        break;
      }
    }
  }

  // Chercher les dates (dd/mm/yyyy ou dd-mm-yyyy)
  const datePattern = /(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/g;
  const dates = [];
  let dateMatch;
  while ((dateMatch = datePattern.exec(text)) !== null) {
    dates.push(dateMatch[0]);
  }

  if (dates.length > 0) {
    result.dateFacture = dates[0];
    if (dates.length > 1) {
      result.dateEcheance = dates[dates.length - 1];
    }
  }

  // Fournisseur: première ligne ou mots-clés
  for (let i = 0; i < Math.min(lines.length, 5); i++) {
    const line = String(lines[i]).trim();
    if (line.length > 3 && line.length < 100 &&
        !line.match(/^\d/) &&
        !line.match(/facture|invoice|date|amount|total|montant/i)) {
      result.fournisseur = line;
      break;
    }
  }

  // Inférer la catégorie à partir du texte
  const inferred = inferCategoryFromText_(normalizeText_(text));
  if (inferred.categorie) {
    result.categorie = inferred.categorie;
    result.sousCategorie = inferred.sousCategorie;
  }

  if (!result.categorie && result.fournisseur) {
    const inferred2 = inferCategoryFromText_(normalizeText_(result.fournisseur));
    result.categorie = inferred2.categorie || 'A verifier';
    result.sousCategorie = inferred2.sousCategorie || ('Fournisseur: ' + result.fournisseur);
  }

  if (!result.montantTTC && !result.numero && !result.fournisseur) {
    throw new Error('Aucune information exploitable extraite par OCR.');
  }

  if (!result.sousCategorie && result.fournisseur) {
    result.sousCategorie = result.fournisseur;
  }

  return result;
}

function normalizeDocumentAiDate_(normalizedValue, fallback) {
  if (normalizedValue && normalizedValue.text) {
    const parsed = parseDate_(normalizedValue.text);
    if (parsed) return formatDateIso_(parsed);
  }
  const parsedFallback = parseDate_(fallback);
  return parsedFallback ? formatDateIso_(parsedFallback) : '';
}

function normalizeDocumentAiAmount_(normalizedValue, fallback) {
  if (normalizedValue && normalizedValue.moneyValue) {
    const units = Number(normalizedValue.moneyValue.units || 0);
    const nanos = Number(normalizedValue.moneyValue.nanos || 0) / 1000000000;
    return units + nanos;
  }
  return toNumber_(fallback);
}

function inferCategoryFromText_(normalizedText) {
  /**
   * Infère la catégorie et sous-catégorie d'une facture à partir du texte (fournisseur, etc.)
   * @param normalizedText {string} Texte normalisé (fournisseur, libellé, etc.)
   * @returns {object} { categorie: string, sousCategorie: string }
   */
  const inferred = { categorie: '', sousCategorie: '' };
  const providerHints = [
    ['EDF', 'Electricite', 'EDF'],
    ['ENGIE', 'Electricite', 'Engie'],
    ['VATTENFALL', 'Electricite', 'Vattenfall'],
    ['TOTALENERGIES', 'Electricite', 'TotalEnergies'],
    ['AUCHAN CARBURANT', 'Carburant', 'Auchan Carburant'],
    ['CARTE CARBURANT', 'Carburant', 'A definir'],
    ['BOUYGUES', 'Internet', 'Bouygues'],
    ['ORANGE', 'Internet', 'Orange'],
    ['SFR', 'Internet', 'SFR'],
    ['FREE', 'Internet', 'Free'],
    ['VERISURE', 'Securite', 'Verisure'],
    ['DELTA', 'Securite', 'Delta Security'],
    ['KIWATCH', 'Securite', 'Kiwatch'],
    ['SKELLO', 'Logiciels / SaaS', 'Skello'],
    ['ADOBE', 'Logiciels / SaaS', 'Adobe'],
    ['GOOGLE', 'Logiciels / SaaS', 'Google Cloud'],
    ['INGENICO', 'Location materiel', 'Ingenico'],
    ['VIMA', 'Honoraires', 'VIMA'],
    ['GESBERT', 'Honoraires', 'GESBERT SARL'],
    ['VEOLIA', 'Eau', 'Veolia'],
    ['SUEZ', 'Eau', 'Suez'],
    ['SAUR', 'Eau', 'Saur'],
    ['REGIE DES EAUX', 'Eau', 'Regie des eaux'],
    ['SERVICE DES EAUX', 'Eau', 'Service des eaux'],
    ['SYNDICAT DES EAUX', 'Eau', 'Syndicat des eaux'],
    ['AGGLOMERATION DES EAUX', 'Eau', 'Agglomeration des eaux'],
    ['COMMUNAUTE DE COMMUNES EAU', 'Eau', 'Communaute de communes eau'],
    ['EAU ET ASSAINISSEMENT', 'Eau', 'Eau et assainissement']
  ];

  for (var i = 0; i < providerHints.length; i++) {
    if (normalizedText.indexOf(providerHints[i][0]) !== -1) {
      inferred.categorie = providerHints[i][1];
      inferred.sousCategorie = providerHints[i][2];
      return inferred;
    }
  }

  if (/LOYER|QUITTANCE|LOCATION/.test(normalizedText)) {
    inferred.categorie = 'Loyer';
    inferred.sousCategorie = 'A definir';
  }
  return inferred;
}

function applyInvoiceOcrResult_(factureId, parsed) {
  const sheet = getSpreadsheet_().getSheetByName(SHEETS.FACTURES);
  const info = findRowByValue_(sheet, 'Facture_id', factureId);
  if (!info.rowNumber) throw new Error('Facture introuvable pour OCR : ' + factureId);
  const headers = getHeaders_(sheet);

  if (parsed.fournisseur) setCellByHeader_(sheet, headers, info.rowNumber, 'Fournisseur', parsed.fournisseur);
  if (parsed.numero) setCellByHeader_(sheet, headers, info.rowNumber, 'Numero_facture', parsed.numero);
  if (parsed.dateFacture) setCellByHeader_(sheet, headers, info.rowNumber, 'Date_facture', parsed.dateFacture);
  if (parsed.dateEcheance) setCellByHeader_(sheet, headers, info.rowNumber, 'Date_echeance', parsed.dateEcheance);
  if (parsed.montantTTC !== null && parsed.montantTTC !== undefined && String(parsed.montantTTC).trim() !== '') {
    setCellByHeader_(sheet, headers, info.rowNumber, 'Montant_TTC', parsed.montantTTC);
  }
  if (parsed.devise) setCellByHeader_(sheet, headers, info.rowNumber, 'Devise', parsed.devise);
  if (parsed.categorie) setCellByHeader_(sheet, headers, info.rowNumber, 'Categorie', parsed.categorie);
  if (parsed.sousCategorie) setCellByHeader_(sheet, headers, info.rowNumber, 'Sous_categorie', parsed.sousCategorie);
  setCellByHeader_(sheet, headers, info.rowNumber, 'Statut_lecture', parsed.statutLecture || 'OCR_OK');
  setCellByHeader_(sheet, headers, info.rowNumber, 'Commentaire', parsed.commentaire || 'Lecture OCR Document AI');
}

function markInvoiceForVerification_(factureId, entite, nomFichier, lienDrive, probleme) {
  const verifierSheet = getSpreadsheet_().getSheetByName(SHEETS.FACTURES_A_VERIFIER);
  const existing = getRows_(verifierSheet).some(function(row) {
    return String(row.Facture_id || '') === String(factureId || '') && String(row.Probleme || '') === String(probleme || '');
  });
  if (!existing) {
    verifierSheet.appendRow([
      factureId,
      entite || '',
      nomFichier || '',
      lienDrive || '',
      probleme || 'A verifier',
      'A_VERIFIER',
      'Genere par OCR'
    ]);
  }
}
function upsertPrevision(payload) {
  installOrUpdateWorkbook();
  payload = payload || {};
  const previsionId = String(payload.Prevision_id || '').trim() || ('PREV-' + Utilities.getUuid().slice(0, 8));
  const entite = String(payload.Entite || '').trim();
  const categorie = String(payload.Categorie || '').trim();
  if (!entite) throw new Error('Entite manquante.');
  if (!categorie) throw new Error('Categorie manquante.');

  const sheet = getSpreadsheet_().getSheetByName(SHEETS.PREVISIONS);
  const headers = getHeaders_(sheet);
  const info = findRowByValue_(sheet, 'Prevision_id', previsionId);
  const merged = Object.assign({
    Prevision_id: previsionId,
    Entite: entite,
    Compte_id: '',
    Categorie: categorie,
    Sous_categorie: '',
    Date_prevue: '',
    Periode: 'Mensuel',
    Frequence: 'Mensuelle',
    Sens: 'Decaissement',
    Montant_prevu: 0,
    Mode_calcul: 'Manuel',
    Base_reference: '',
    Active: 'Oui',
    Commentaire: ''
  }, payload);
  const values = headers.map(function(header) {
    return merged[header] !== undefined ? merged[header] : '';
  });

  if (info.rowNumber) {
    sheet.getRange(info.rowNumber, 1, 1, headers.length).setValues([values]);
  } else {
    sheet.appendRow(values);
  }

  return {
    success: true,
    recentPrevisions: getRecentPrevisions(50),
    forecastSummary: buildForecastSummary_(),
    alerts: getAlerts_(12),
    periodAnalytics: buildPeriodAnalytics_(),
    counts: getSheetCounts_()
  };
}

function getRecentPrevisions(limit, entite) {
  let rows = getRows_(getSpreadsheet_().getSheetByName(SHEETS.PREVISIONS));
  if (entite && entite !== 'TOUTES') {
    rows = rows.filter(function(row) { return String(row.Entite || '').trim() === entite; });
  }
  return rows.slice(-1 * Math.max(limit || 50, 1)).reverse();
}

function upsertScenario(payload) {
  installOrUpdateWorkbook();
  payload = payload || {};
  const scenarioId = String(payload.Scenario_id || '').trim() || ('SCN-' + Utilities.getUuid().slice(0, 8));
  const nomScenario = String(payload.Nom_scenario || '').trim();
  const entite = String(payload.Entite || '').trim();
  const categorie = String(payload.Categorie || '').trim();
  if (!nomScenario) throw new Error('Nom du scenario manquant.');
  if (!entite) throw new Error('Entite manquante.');
  if (!categorie) throw new Error('Categorie manquante.');

  const sheet = getSpreadsheet_().getSheetByName(SHEETS.SCENARIOS);
  const headers = getHeaders_(sheet);
  const info = findRowByValue_(sheet, 'Scenario_id', scenarioId);
  const merged = Object.assign({
    Scenario_id: scenarioId,
    Nom_scenario: nomScenario,
    Entite: entite,
    Categorie: categorie,
    Sous_categorie: '',
    Type_ajustement: 'pourcentage',
    Valeur_ajustement: 0,
    Montant_cible: '',
    Date_effet: '',
    Active: 'Oui',
    Commentaire: ''
  }, payload);
  const values = headers.map(function(header) {
    return merged[header] !== undefined ? merged[header] : '';
  });

  if (info.rowNumber) sheet.getRange(info.rowNumber, 1, 1, headers.length).setValues([values]);
  else sheet.appendRow(values);

  return {
    success: true,
    recentScenarios: getRecentScenarios(50),
    scenarioSummary: buildScenarioSummary_(),
    counts: getSheetCounts_()
  };
}

function getRecentScenarios(limit, entite) {
  let rows = getRows_(getSpreadsheet_().getSheetByName(SHEETS.SCENARIOS));
  if (entite && entite !== 'TOUTES') {
    rows = rows.filter(function(row) { return String(row.Entite || '').trim() === entite; });
  }
  return rows.slice(-1 * Math.max(limit || 50, 1)).reverse();
}

function getRecentRapprochements(limit, entite) {
  let rows = getRows_(getSpreadsheet_().getSheetByName(SHEETS.RAPPROCHEMENTS));
  if (entite && entite !== 'TOUTES') {
    rows = rows.filter(function(row) { return String(row.Entite || '').trim() === entite; });
  }
  return rows.slice(-1 * Math.max(limit || 50, 1)).reverse();
}

function getFacturesToVerify(limit, entite) {
  let rows = getRows_(getSpreadsheet_().getSheetByName(SHEETS.FACTURES_A_VERIFIER));
  if (entite && entite !== 'TOUTES') {
    rows = rows.filter(function(row) { return String(row.Entite || '').trim() === entite; });
  }
  return rows.filter(function(row) {
      return String(row.Statut || '').trim() !== 'TRAITE';
    })
    .slice(0, limit || 50);
}

function getTransactionsToClassify(limit, entite) {
  let rows = getRows_(getSpreadsheet_().getSheetByName(SHEETS.A_CLASSER));
  if (entite && entite !== 'TOUTES') {
    rows = rows.filter(function(row) { return String(row.Entite || '').trim() === entite; });
  }
  return rows.filter(function(row) {
      return String(row.Statut || '').trim() !== 'TRAITE';
    })
    .slice(0, limit || 50);
}

function getRecentMovements(limit, entite) {
  /**
   * Retourne les mouvements récents (les plus nouveaux en premier)
   * @param limit {number} Nombre de mouvements à retourner
   * @param entite {string} Entité à filtrer (optionnel)
   * @returns {array} Mouvements triés par date décroissante (plus récents d'abord)
   */
  let rows = getRows_(getSpreadsheet_().getSheetByName(SHEETS.MOUVEMENTS));
  if (entite && entite !== 'TOUTES') {
    rows = rows.filter(function(row) { return String(row.Entite || '').trim() === entite; });
  }
  return rows.slice(-1 * Math.max(limit || 20, 1)).reverse();
}

function getFacturesToPay(limit, entite) {
  let rows = getRows_(getSpreadsheet_().getSheetByName(SHEETS.FACTURES));
  if (entite && entite !== 'TOUTES') {
    rows = rows.filter(function(row) { return String(row.Entite || '').trim() === entite; });
  }
  return rows.filter(function(row) {
      const status = String(row.Statut_paiement || '').trim();
      return status !== 'REGLEE' && status !== 'ANNULEE';
    })
    .slice(0, limit || 100);
}

function getTransactionsWithoutInvoice(limit, entite) {
  let rows = getRows_(getSpreadsheet_().getSheetByName(SHEETS.MOUVEMENTS));
  if (entite && entite !== 'TOUTES') {
    rows = rows.filter(function(row) { return String(row.Entite || '').trim() === entite; });
  }
  return rows.filter(function(row) {
      return String(row.Sens || '').trim() === 'Decaissement' &&
        String(row.Statut_rapprochement || '').trim() !== 'RAPPROCHE';
    })
    .slice(-1 * Math.max(limit || 50, 1))
    .reverse();
}

function getDriveRoots_() {
  const roots = getDriveRootMap_();
  return Object.keys(roots).map(function(entite) {
    return { entite: entite, folderId: roots[entite] };
  });
}

function getExistingFactureIds_(entite) {
  const map = {};
  getRows_(getSpreadsheet_().getSheetByName(SHEETS.FACTURES)).forEach(function(row) {
    if (entite && String(row.Entite || '').trim() !== entite) return;
    const factureId = String(row.Facture_id || '');
    if (factureId.indexOf('FAC-') === 0) {
      map[factureId.replace(/^FAC-/, '')] = true;
    }
  });
  return map;
}

function parseInvoiceFileName_(fileName) {
  const cleaned = String(fileName || '').trim().replace(/\.pdf$/i, '');
  const normalized = normalizeText_(cleaned);
  let fournisseur = cleaned;
  let numero = '';
  let categorie = '';
  let sousCategorie = '';
  let statutLecture = 'A_VERIFIER';
  let probleme = 'Lecture automatique a confirmer';

  const numberMatch = cleaned.match(/([A-Z]{0,4}\d{3,})/i);
  if (numberMatch) numero = numberMatch[1];

  const providerHints = [
    ['EDF', 'Electricite', 'EDF'],
    ['ENGIE', 'Electricite', 'Engie'],
    ['VATTENFALL', 'Electricite', 'Vattenfall'],
    ['TOTALENERGIES', 'Electricite', 'TotalEnergies'],
    ['AUCHAN CARBURANT', 'Carburant', 'Auchan Carburant'],
    ['CARTE CARBURANT', 'Carburant', 'A definir'],
    ['BOUYGUES', 'Internet', 'Bouygues'],
    ['ORANGE', 'Internet', 'Orange'],
    ['SFR', 'Internet', 'SFR'],
    ['FREE', 'Internet', 'Free'],
    ['VERISURE', 'Securite', 'Verisure'],
    ['DELTA', 'Securite', 'Delta Security'],
    ['KIWATCH', 'Securite', 'Kiwatch'],
    ['SKELLO', 'Logiciels / SaaS', 'Skello'],
    ['ADOBE', 'Logiciels / SaaS', 'Adobe'],
    ['GOOGLE', 'Logiciels / SaaS', 'Google Cloud'],
    ['INGENICO', 'Location materiel', 'Ingenico'],
    ['VIMA', 'Honoraires', 'VIMA'],
    ['GESBERT', 'Honoraires', 'GESBERT SARL'],
    ['VEOLIA', 'Eau', 'Veolia'],
    ['SUEZ', 'Eau', 'Suez'],
    ['SAUR', 'Eau', 'Saur'],
    ['REGIE DES EAUX', 'Eau', 'Regie des eaux'],
    ['SERVICE DES EAUX', 'Eau', 'Service des eaux'],
    ['SYNDICAT DES EAUX', 'Eau', 'Syndicat des eaux'],
    ['AGGLOMERATION DES EAUX', 'Eau', 'Agglomeration des eaux'],
    ['COMMUNAUTE DE COMMUNES EAU', 'Eau', 'Communaute de communes eau'],
    ['EAU ET ASSAINISSEMENT', 'Eau', 'Eau et assainissement']
  ];

  for (var i = 0; i < providerHints.length; i++) {
    if (normalized.indexOf(providerHints[i][0]) !== -1) {
      fournisseur = providerHints[i][2];
      categorie = providerHints[i][1];
      sousCategorie = providerHints[i][2];
      statutLecture = 'OK';
      probleme = '';
      break;
    }
  }

  if (!categorie && /LOYER|QUITTANCE|LOCATION/.test(normalized)) {
    categorie = 'Loyer';
    sousCategorie = 'A definir';
  }

  return {
    fournisseur: fournisseur,
    numero: numero,
    categorie: categorie,
    sousCategorie: sousCategorie,
    statutLecture: statutLecture,
    probleme: probleme
  };
}

function runAutomaticReconciliation(payload) {
  installOrUpdateWorkbook();
  payload = payload || {};
  const entitesFilter = payload.entites && payload.entites.length ? payload.entites : [];
  const rapprochementsSheet = getSpreadsheet_().getSheetByName(SHEETS.RAPPROCHEMENTS);
  const mouvements = getRows_(getSpreadsheet_().getSheetByName(SHEETS.MOUVEMENTS)).filter(function(row) {
    if (!row.Transaction_uid) return false;
    if (entitesFilter.length && entitesFilter.indexOf(String(row.Entite || '').trim()) === -1) return false;
    if (String(row.Sens || '').trim() !== 'Decaissement') return false;
    return String(row.Statut_rapprochement || '').trim() !== 'RAPPROCHE';
  });
  const factures = getRows_(getSpreadsheet_().getSheetByName(SHEETS.FACTURES)).filter(function(row) {
    if (!row.Facture_id) return false;
    if (entitesFilter.length && entitesFilter.indexOf(String(row.Entite || '').trim()) === -1) return false;
    if (String(row.Statut_paiement || '').trim() === 'ANNULEE') return false;
    return String(row.Statut_rapprochement || '').trim() !== 'RAPPROCHE';
  });

  const existing = getExistingRapprochementIndex_();
  const usedTransactions = {};
  let autoMatched = 0;
  let toVerify = 0;

  factures.forEach(function(facture) {
    if (existing.factures[facture.Facture_id]) return;

    let best = null;
    mouvements.forEach(function(tx) {
      if (usedTransactions[tx.Transaction_uid]) return;
      if (existing.transactions[tx.Transaction_uid]) return;
      const score = scoreReconciliationPair_(facture, tx);
      if (score.total < 55) return;
      if (!best || score.total > best.score.total) {
        best = { facture: facture, transaction: tx, score: score };
      }
    });

    if (!best) return;

    usedTransactions[best.transaction.Transaction_uid] = true;
    const threshold = getSettingsConfig_().reconciliationConfidenceThreshold || 75;
    const statut = best.score.total >= threshold ? 'AUTO_RAPPROCHE' : 'A_VERIFIER';
    rapprochementsSheet.appendRow([
      'RAPP-' + Utilities.getUuid().slice(0, 8),
      best.transaction.Transaction_uid,
      best.facture.Facture_id,
      best.facture.Entite,
      Math.abs(toNumber_(best.transaction.Montant)),
      Math.abs(toNumber_(best.facture.Montant_TTC)),
      best.score.ecart,
      best.score.total,
      statut,
      best.score.commentaire,
      new Date()
    ]);

    if (statut === 'AUTO_RAPPROCHE') {
      autoMatched++;
      updateMovementReconciliationStatus_(best.transaction.Transaction_uid, 'RAPPROCHE');
      updateInvoiceReconciliationStatus_(best.facture.Facture_id, 'RAPPROCHE', 'REGLEE', best.transaction.Transaction_uid, best.transaction.Date);
    } else {
      toVerify++;
      updateMovementReconciliationStatus_(best.transaction.Transaction_uid, 'A_VERIFIER');
      updateInvoiceReconciliationStatus_(best.facture.Facture_id, 'A_VERIFIER', best.facture.Statut_paiement || 'A_REGLER', '', best.facture.Date_paiement || '');
    }
  });

  return {
    success: true,
    autoMatched: autoMatched,
    toVerify: toVerify
  };
}

function scoreReconciliationPair_(facture, tx) {
  /**
   * Score la correspondance entre une facture et une transaction
   * @param facture {object} Ligne du sheet Factures
   * @param tx {object} Transaction (ligne du sheet Mouvements)
   * @returns {object} { total: number, ecart: number, commentaire: string, raisons: array }
   */
  if (String(facture.Entite || '').trim() !== String(tx.Entite || '').trim()) {
    return { total: 0, ecart: 0, commentaire: 'Entite differente' };
  }

  const montantFacture = Math.abs(toNumber_(facture.Montant_TTC));
  const montantTx = Math.abs(toNumber_(tx.Montant));
  const ecart = Math.abs(montantFacture - montantTx);
  const fournisseur = normalizeText_(facture.Fournisseur || '');
  const numero = normalizeText_(facture.Numero_facture || '');
  const libelle = normalizeText_(tx.Libelle_normalise || tx.Libelle || '');
  const factureDate = parseDate_(facture.Date_facture);
  const txDate = parseDate_(tx.Date);
  let total = 30;
  const raisons = [];

  if (montantFacture > 0) {
    if (ecart <= 0.01) {
      total += 45;
      raisons.push('montant exact');
    } else if (ecart <= 1) {
      total += 35;
      raisons.push('montant tres proche');
    } else if (ecart <= 5) {
      total += 20;
      raisons.push('montant proche');
    } else if (ecart <= 20) {
      total += 8;
      raisons.push('montant approximatif');
    }
  }

  if (String(numero || '').trim().length > 0 && libelle.indexOf(numero) !== -1) {
    total += 25;
    raisons.push('numero facture trouve');
  }

  if (String(fournisseur || '').trim().length >= 4 && libelle.indexOf(fournisseur) !== -1) {
    total += 20;
    raisons.push('fournisseur trouve');
  }

  if (factureDate && txDate) {
    const diffDays = Math.abs(Math.round((txDate.getTime() - factureDate.getTime()) / 86400000));
    if (diffDays <= 3) {
      total += 20;
      raisons.push('date tres proche');
    } else if (diffDays <= 10) {
      total += 12;
      raisons.push('date proche');
    } else if (diffDays <= 30) {
      total += 5;
      raisons.push('date compatible');
    }
  } else if (monthFolderMatchesDate_(facture.Mois, facture.Annee, txDate)) {
    total += 5;
    raisons.push('mois compatible');
  }

  if (facture.Categorie && tx.Categorie && String(facture.Categorie) === String(tx.Categorie)) {
    total += 5;
    raisons.push('categorie proche');
  }

  return {
    total: total,
    ecart: ecart,
    commentaire: raisons.join(', ') || 'Aucune correspondance fiable'
  };
}

function monthFolderMatchesDate_(moisValue, anneeValue, txDate) {
  if (!txDate) return false;
  const year = String(anneeValue || '').trim();
  if (year && String(txDate.getFullYear()) !== year) return false;
  const normalized = normalizeText_(moisValue || '');
  const monthMap = {
    'JANVIER': 0, 'FEVRIER': 1, 'MARS': 2, 'AVRIL': 3, 'MAI': 4, 'JUIN': 5,
    'JUILLET': 6, 'AOUT': 7, 'SEPTEMBRE': 8, 'OCTOBRE': 9, 'NOVEMBRE': 10, 'DECEMBRE': 11
  };
  const keys = Object.keys(monthMap);
  for (var i = 0; i < keys.length; i++) {
    if (normalized.indexOf(keys[i]) !== -1) {
      return txDate.getMonth() === monthMap[keys[i]];
    }
  }
  return false;
}

function getExistingRapprochementIndex_() {
  const index = { factures: {}, transactions: {} };
  getRows_(getSpreadsheet_().getSheetByName(SHEETS.RAPPROCHEMENTS)).forEach(function(row) {
    if (row.Facture_id) index.factures[row.Facture_id] = true;
    if (row.Transaction_uid) index.transactions[row.Transaction_uid] = true;
  });
  return index;
}

function updateMovementReconciliationStatus_(transactionUid, status) {
  const sheet = getSpreadsheet_().getSheetByName(SHEETS.MOUVEMENTS);
  const info = findRowByValue_(sheet, 'Transaction_uid', transactionUid);
  if (!info.rowNumber) return;
  const headers = getHeaders_(sheet);
  setCellByHeader_(sheet, headers, info.rowNumber, 'Statut_rapprochement', status);
  setCellByHeader_(sheet, headers, info.rowNumber, 'Date_maj', new Date());
}

function updateInvoiceReconciliationStatus_(factureId, rapprochementStatus, paiementStatus, transactionUid, datePaiement) {
  const sheet = getSpreadsheet_().getSheetByName(SHEETS.FACTURES);
  const info = findRowByValue_(sheet, 'Facture_id', factureId);
  if (!info.rowNumber) return;
  const headers = getHeaders_(sheet);
  setCellByHeader_(sheet, headers, info.rowNumber, 'Statut_rapprochement', rapprochementStatus);
  if (paiementStatus) setCellByHeader_(sheet, headers, info.rowNumber, 'Statut_paiement', paiementStatus);
  if (transactionUid !== undefined) setCellByHeader_(sheet, headers, info.rowNumber, 'Transaction_uid_liee', transactionUid);
  if (datePaiement) setCellByHeader_(sheet, headers, info.rowNumber, 'Date_paiement', datePaiement);
}
function buildForecastSummary_(entite) {
  const now = new Date();
  const currentKey = Utilities.formatDate(now, Session.getScriptTimeZone() || 'Europe/Paris', 'yyyy-MM');
  const previousDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const previousKey = Utilities.formatDate(previousDate, Session.getScriptTimeZone() || 'Europe/Paris', 'yyyy-MM');
  let mouvements = getRows_(getSpreadsheet_().getSheetByName(SHEETS.MOUVEMENTS));
  let previsions = getRows_(getSpreadsheet_().getSheetByName(SHEETS.PREVISIONS)).filter(function(row) {
    return String(row.Active || '').toLowerCase() !== 'non';
  });

  // Filtrer par entité si spécifiée
  if (entite && entite !== 'TOUTES') {
    mouvements = mouvements.filter(function(row) { return String(row.Entite || '').trim() === entite; });
    previsions = previsions.filter(function(row) { return String(row.Entite || '').trim() === entite; });
  }

  const monthlyByCategory = {};
  mouvements.forEach(function(row) {
    const date = parseDate_(row.Date);
    if (!date) return;
    const monthKey = Utilities.formatDate(date, Session.getScriptTimeZone() || 'Europe/Paris', 'yyyy-MM');
    const categoryKey = buildForecastKey_(row.Entite, row.Categorie, row.Sous_categorie);
    if (!monthlyByCategory[monthKey]) monthlyByCategory[monthKey] = {};
    monthlyByCategory[monthKey][categoryKey] = (monthlyByCategory[monthKey][categoryKey] || 0) + toNumber_(row.Montant);
  });

  const currentActualMap = monthlyByCategory[currentKey] || {};
  const previousActualMap = monthlyByCategory[previousKey] || {};
  const forecastMap = {};

  previsions.forEach(function(row) {
    const date = parseDate_(row.Date_prevue);
    const monthKey = date ? Utilities.formatDate(date, Session.getScriptTimeZone() || 'Europe/Paris', 'yyyy-MM') : currentKey;
    if (monthKey !== currentKey) return;
    const key = buildForecastKey_(row.Entite, row.Categorie, row.Sous_categorie);
    const sign = String(row.Sens || '').trim() === 'Encaissement' ? 1 : -1;
    forecastMap[key] = (forecastMap[key] || 0) + sign * Math.abs(toNumber_(row.Montant_prevu));
  });

  if (!Object.keys(forecastMap).length) {
    forecastFromHistory_(mouvements).forEach(function(item) {
      forecastMap[item.key] = item.montant;
    });
  }

  const unionKeys = {};
  [forecastMap, currentActualMap, previousActualMap].forEach(function(map) {
    Object.keys(map).forEach(function(key) { unionKeys[key] = true; });
  });

  const lines = Object.keys(unionKeys).map(function(key) {
    const parsed = parseForecastKey_(key);
    const forecast = toNumber_(forecastMap[key]);
    const actual = toNumber_(currentActualMap[key]);
    const previous = toNumber_(previousActualMap[key]);
    return {
      entite: parsed.entite,
      categorie: parsed.categorie,
      sousCategorie: parsed.sousCategorie,
      forecast: forecast,
      actual: actual,
      previous: previous,
      variance: actual - forecast,
      variancePrevious: actual - previous
    };
  }).sort(function(a, b) {
    return Math.abs(b.variance) - Math.abs(a.variance);
  });

  return {
    currentMonth: currentKey,
    previousMonth: previousKey,
    totals: {
      forecast: sumByKey_(lines, 'forecast'),
      actual: sumByKey_(lines, 'actual'),
      previous: sumByKey_(lines, 'previous')
    },
    lines: lines.slice(0, 50)
  };
}

function buildScenarioSummary_(entite) {
  const base = buildForecastSummary_(entite);
  let scenarios = getRows_(getSpreadsheet_().getSheetByName(SHEETS.SCENARIOS)).filter(function(row) {
    return String(row.Active || '').toLowerCase() !== 'non';
  });

  // Filtrer par entité si spécifiée
  if (entite && entite !== 'TOUTES') {
    scenarios = scenarios.filter(function(row) { return String(row.Entite || '').trim() === entite; });
  }

  const scenarioLines = scenarios.map(function(row) {
    const match = (base.lines || []).find(function(line) {
      return String(line.entite || '').trim() === String(row.Entite || '').trim() &&
        String(line.categorie || '').trim() === String(row.Categorie || '').trim() &&
        String(line.sousCategorie || '').trim() === String(row.Sous_categorie || '').trim();
    });

    const baseAmount = match ? toNumber_(match.forecast) : 0;
    const adjustmentType = String(row.Type_ajustement || 'pourcentage').trim().toLowerCase();
    const adjustmentValue = toNumber_(row.Valeur_ajustement);
    const targetAmount = toNumber_(row.Montant_cible);
    let projected = baseAmount;

    if (adjustmentType === 'pourcentage') projected = baseAmount * (1 + adjustmentValue / 100);
    else if (adjustmentType === 'montant') projected = baseAmount + adjustmentValue;
    else if (adjustmentType === 'cible') projected = targetAmount;

    return {
      scenarioId: row.Scenario_id || '',
      nom: row.Nom_scenario || '',
      entite: row.Entite || '',
      categorie: row.Categorie || '',
      sousCategorie: row.Sous_categorie || '',
      baseAmount: baseAmount,
      projectedAmount: projected,
      impact: projected - baseAmount,
      typeAjustement: row.Type_ajustement || '',
      valeurAjustement: row.Valeur_ajustement || '',
      montantCible: row.Montant_cible || '',
      dateEffet: row.Date_effet || ''
    };
  });

  return {
    baseForecastTotal: sumByKey_(base.lines || [], 'forecast'),
    scenarioForecastTotal: scenarioLines.reduce(function(total, row) { return total + row.projectedAmount; }, 0),
    impactTotal: scenarioLines.reduce(function(total, row) { return total + row.impact; }, 0),
    lines: scenarioLines.slice(0, 50)
  };
}

function buildPeriodAnalytics_(entite) {
  let rows = getRows_(getSpreadsheet_().getSheetByName(SHEETS.MOUVEMENTS));

  // Filtrer par entité si spécifiée
  if (entite && entite !== 'TOUTES') {
    rows = rows.filter(function(row) { return String(row.Entite || '').trim() === entite; });
  }

  const now = new Date();
  return {
    day: computePeriodAnalytics_(rows, now, 'day'),
    week: computePeriodAnalytics_(rows, now, 'week'),
    month: computePeriodAnalytics_(rows, now, 'month'),
    quarter: computePeriodAnalytics_(rows, now, 'quarter'),
    semester: computePeriodAnalytics_(rows, now, 'semester'),
    year: computePeriodAnalytics_(rows, now, 'year')
  };
}

function computePeriodAnalytics_(rows, now, periodType) {
  const range = getPeriodRange_(now, periodType, 0);
  const previousRange = getPeriodRange_(now, periodType, -1);
  const current = aggregateRowsForRange_(rows, range.start, range.end);
  const previous = aggregateRowsForRange_(rows, previousRange.start, previousRange.end);
  return {
    currentLabel: range.label,
    previousLabel: previousRange.label,
    current: current,
    previous: previous
  };
}

function getPeriodRange_(baseDate, periodType, offset) {
  let start;
  let end;
  let label;
  const date = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate());

  if (periodType === 'day') {
    start = new Date(date.getFullYear(), date.getMonth(), date.getDate() + offset);
    end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 1);
    label = Utilities.formatDate(start, Session.getScriptTimeZone() || 'Europe/Paris', 'yyyy-MM-dd');
  } else if (periodType === 'week') {
    const day = (date.getDay() + 6) % 7;
    start = new Date(date.getFullYear(), date.getMonth(), date.getDate() - day + (offset * 7));
    end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 7);
    label = 'Semaine du ' + Utilities.formatDate(start, Session.getScriptTimeZone() || 'Europe/Paris', 'yyyy-MM-dd');
  } else if (periodType === 'month') {
    start = new Date(date.getFullYear(), date.getMonth() + offset, 1);
    end = new Date(start.getFullYear(), start.getMonth() + 1, 1);
    label = Utilities.formatDate(start, Session.getScriptTimeZone() || 'Europe/Paris', 'yyyy-MM');
  } else if (periodType === 'quarter') {
    const quarterStartMonth = Math.floor(date.getMonth() / 3) * 3 + (offset * 3);
    start = new Date(date.getFullYear(), quarterStartMonth, 1);
    end = new Date(start.getFullYear(), start.getMonth() + 3, 1);
    label = 'T' + (Math.floor(start.getMonth() / 3) + 1) + ' ' + start.getFullYear();
  } else if (periodType === 'semester') {
    const semesterStartMonth = (date.getMonth() < 6 ? 0 : 6) + (offset * 6);
    start = new Date(date.getFullYear(), semesterStartMonth, 1);
    end = new Date(start.getFullYear(), start.getMonth() + 6, 1);
    label = (start.getMonth() < 6 ? 'S1 ' : 'S2 ') + start.getFullYear();
  } else {
    start = new Date(date.getFullYear() + offset, 0, 1);
    end = new Date(start.getFullYear() + 1, 0, 1);
    label = String(start.getFullYear());
  }

  return { start: start, end: end, label: label };
}

function aggregateRowsForRange_(rows, start, end) {
  return rows.reduce(function(acc, row) {
    const date = parseDate_(row.Date);
    if (!date || date < start || date >= end) return acc;
    acc.encaissements += toNumber_(row.Encaissement);
    acc.decaissements += toNumber_(row.Decaissement);
    acc.net += toNumber_(row.Montant);
    acc.charges += Math.abs(toNumber_(row.Decaissement));
    return acc;
  }, { encaissements: 0, decaissements: 0, net: 0, charges: 0 });
}

function forecastFromHistory_(mouvements) {
  const now = new Date();
  const history = {};
  mouvements.forEach(function(row) {
    const date = parseDate_(row.Date);
    if (!date) return;
    const diffMonths = (now.getFullYear() - date.getFullYear()) * 12 + (now.getMonth() - date.getMonth());
    if (diffMonths < 1 || diffMonths > 3) return;
    const key = buildForecastKey_(row.Entite, row.Categorie, row.Sous_categorie);
    if (!history[key]) history[key] = [];
    history[key].push(toNumber_(row.Montant));
  });

  return Object.keys(history).map(function(key) {
    const values = history[key];
    const avg = values.reduce(function(total, value) { return total + value; }, 0) / values.length;
    return { key: key, montant: avg };
  });
}

function buildForecastKey_(entite, categorie, sousCategorie) {
  return [String(entite || '').trim(), String(categorie || '').trim(), String(sousCategorie || '').trim()].join('|');
}

function parseForecastKey_(key) {
  const parts = String(key || '').split('|');
  return {
    entite: parts[0] || '',
    categorie: parts[1] || '',
    sousCategorie: parts[2] || ''
  };
}

function sumByKey_(rows, field) {
  return rows.reduce(function(total, row) {
    return total + toNumber_(row[field]);
  }, 0);
}

function getAlerts_(limit, entite) {
  const alerts = [];
  const dashboard = buildDashboardSummary_(entite);
  const forecast = buildForecastSummary_(entite);
  const settings = getSettingsConfig_();
  const aClasserCount = getTransactionsToClassify(1000, entite).length;
  const facturesToPayCount = getFacturesToPay(1000, entite).length;
  const facturesToVerifyCount = getFacturesToVerify(1000, entite).length;
  const transactionsWithoutInvoiceCount = getTransactionsWithoutInvoice(1000, entite).length;

  dashboard.comptes.forEach(function(compte) {
    if (compte.value < settings.alertNegativeBalance) {
      alerts.push({ level: 'danger', title: 'Solde negatif', detail: compte.label + ' a un solde de ' + formatMoneyText_(compte.value) });
    } else if (compte.value < settings.alertLowBalance) {
      alerts.push({ level: 'warn', title: 'Solde faible', detail: compte.label + ' est sous ' + formatMoneyText_(settings.alertLowBalance) + ' (' + formatMoneyText_(compte.value) + ')' });
    }
  });

  if (aClasserCount) {
    alerts.push({ level: 'warn', title: 'Transactions a classer', detail: aClasserCount + ' transaction(s) attendent une categorie.' });
  }
  if (facturesToVerifyCount) {
    alerts.push({ level: 'warn', title: 'Factures a verifier', detail: facturesToVerifyCount + ' facture(s) demandent un controle.' });
  }
  if (facturesToPayCount) {
    alerts.push({ level: 'warn', title: 'Factures a regler', detail: facturesToPayCount + ' facture(s) restent ouvertes au paiement.' });
  }
  if (transactionsWithoutInvoiceCount) {
    alerts.push({ level: 'warn', title: 'Transactions sans facture', detail: transactionsWithoutInvoiceCount + ' sortie(s) bancaires restent non rapprochees.' });
  }

  forecast.lines.slice(0, 8).forEach(function(line) {
    const variancePercent = line.forecast ? Math.abs((line.variance / line.forecast) * 100) : 0;
    if (Math.abs(line.variance) >= settings.alertVarianceAmount && (variancePercent >= settings.alertVariancePercent || Math.abs(line.forecast) === 0)) {
      alerts.push({
        level: line.variance > 0 ? 'danger' : 'warn',
        title: 'Ecart previsionnel',
        detail: [line.entite, line.categorie, line.sousCategorie].filter(Boolean).join(' / ') + ' : ecart de ' + formatMoneyText_(line.variance)
      });
    }
    const previousVariancePercent = line.previous ? Math.abs((line.variancePrevious / line.previous) * 100) : 0;
    if (Math.abs(line.variancePrevious) >= settings.alertVarianceAmount && (previousVariancePercent >= settings.alertVariancePercent || Math.abs(line.previous) === 0)) {
      alerts.push({
        level: line.variancePrevious > 0 ? 'danger' : 'warn',
        title: 'Variation vs mois precedent',
        detail: [line.entite, line.categorie, line.sousCategorie].filter(Boolean).join(' / ') + ' : variation de ' + formatMoneyText_(line.variancePrevious)
      });
    }
  });

  return alerts.slice(0, limit || 12);
}

function formatMoneyText_(value) {
  return Utilities.formatString('%.2f EUR', toNumber_(value));
}

function getSettingsConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    alertLowBalance: toNumber_(props.getProperty('ALERT_LOW_BALANCE') || DEFAULT_SETTINGS.alertLowBalance),
    alertNegativeBalance: toNumber_(props.getProperty('ALERT_NEGATIVE_BALANCE') || DEFAULT_SETTINGS.alertNegativeBalance),
    alertVarianceAmount: toNumber_(props.getProperty('ALERT_VARIANCE_AMOUNT') || DEFAULT_SETTINGS.alertVarianceAmount),
    alertVariancePercent: toNumber_(props.getProperty('ALERT_VARIANCE_PERCENT') || DEFAULT_SETTINGS.alertVariancePercent)
  };
}

function getDriveRootMap_() {
  const props = PropertiesService.getScriptProperties();
  const override = JSON.parse(props.getProperty('DRIVE_FACTURES_ROOTS_OVERRIDE') || '{}');
  return Object.assign({}, DRIVE_FACTURES_ROOTS, override);
}
function getExportPayload_() {
  return {
    mouvements: getRows_(getSpreadsheet_().getSheetByName(SHEETS.MOUVEMENTS)),
    factures: getRows_(getSpreadsheet_().getSheetByName(SHEETS.FACTURES)),
    rapprochements: getRows_(getSpreadsheet_().getSheetByName(SHEETS.RAPPROCHEMENTS)),
    previsions: getRows_(getSpreadsheet_().getSheetByName(SHEETS.PREVISIONS)),
    scenarios: getRows_(getSpreadsheet_().getSheetByName(SHEETS.SCENARIOS))
  };
}

function upsertCompte(payload) {
  installOrUpdateWorkbook();
  payload = payload || {};
  const compteId = String(payload.Compte_id || '').trim();
  if (!compteId) throw new Error('Compte_id manquant.');

  const sheet = getSpreadsheet_().getSheetByName(SHEETS.COMPTES);
  const headers = getHeaders_(sheet);
  const info = findRowByValue_(sheet, 'Compte_id', compteId);
  const values = headers.map(function(header) {
    return payload[header] !== undefined ? payload[header] : '';
  });

  if (info.rowNumber) {
    sheet.getRange(info.rowNumber, 1, 1, headers.length).setValues([values]);
  } else {
    sheet.appendRow(values);
  }

  return {
    success: true,
    config: buildConfigPayload_(),
    counts: getSheetCounts_()
  };
}

function upsertCategorie(payload) {
  installOrUpdateWorkbook();
  payload = payload || {};
  const categorie = String(payload.Categorie || '').trim();
  const sousCategorie = String(payload.Sous_categorie || '').trim();
  if (!categorie) throw new Error('Categorie manquante.');

  const sheet = getSpreadsheet_().getSheetByName(SHEETS.CATEGORIES);
  const headers = getHeaders_(sheet);
  const rows = getRows_(sheet);
  let rowNumber = 0;
  rows.forEach(function(row, index) {
    if (String(row.Categorie || '').trim() === categorie && String(row.Sous_categorie || '').trim() === sousCategorie) {
      rowNumber = index + 2;
    }
  });

  const values = headers.map(function(header) {
    return payload[header] !== undefined ? payload[header] : '';
  });

  if (rowNumber) {
    sheet.getRange(rowNumber, 1, 1, headers.length).setValues([values]);
  } else {
    sheet.appendRow(values);
  }

  return {
    success: true,
    config: buildConfigPayload_(),
    counts: getSheetCounts_()
  };
}

function setAlertSettings(payload) {
  payload = payload || {};
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    ALERT_LOW_BALANCE: String(payload.alertLowBalance != null ? payload.alertLowBalance : DEFAULT_SETTINGS.alertLowBalance),
    ALERT_NEGATIVE_BALANCE: String(payload.alertNegativeBalance != null ? payload.alertNegativeBalance : DEFAULT_SETTINGS.alertNegativeBalance),
    ALERT_VARIANCE_AMOUNT: String(payload.alertVarianceAmount != null ? payload.alertVarianceAmount : DEFAULT_SETTINGS.alertVarianceAmount),
    ALERT_VARIANCE_PERCENT: String(payload.alertVariancePercent != null ? payload.alertVariancePercent : DEFAULT_SETTINGS.alertVariancePercent)
  }, true);
  return {
    success: true,
    settings: getSettingsConfig_(),
    alerts: getAlerts_(12)
  };
}

function setDriveRoot(payload) {
  payload = payload || {};
  const entite = String(payload.entite || '').trim();
  const folderId = String(payload.folderId || '').trim();
  if (!entite) throw new Error('Entite manquante.');

  const props = PropertiesService.getScriptProperties();
  const current = JSON.parse(props.getProperty('DRIVE_FACTURES_ROOTS_OVERRIDE') || '{}');
  if (folderId) current[entite] = folderId;
  else delete current[entite];
  props.setProperty('DRIVE_FACTURES_ROOTS_OVERRIDE', JSON.stringify(current));
  return {
    success: true,
    driveConfig: getDriveRoots_()
  };
}
function buildConfigPayload_() {
  const ss = getSpreadsheet_();
  const comptes = getRows_(ss.getSheetByName(SHEETS.COMPTES)).filter(function(row) {
    return String(row.Actif || '').toLowerCase() !== 'non';
  });
  const categories = getRows_(ss.getSheetByName(SHEETS.CATEGORIES)).filter(function(row) {
    return String(row.Active || '').toLowerCase() !== 'non';
  });
  const regles = getRows_(ss.getSheetByName(SHEETS.REGLES)).filter(function(row) {
    return String(row.Active || '').toLowerCase() !== 'non';
  });
  const entites = uniqueValues_(comptes.map(function(row) { return row.Entite; }));

  return {
    entites: entites,
    comptes: comptes,
    categories: categories,
    regles: regles,
    categoryMap: buildCategoryMap_(categories),
    driveRoots: getDriveRoots_()
  };
}

function getGraphData_(entite) {
  const rows = getRows_(getSpreadsheet_().getSheetByName(SHEETS.MOUVEMENTS));
  const monthly = {};
  rows.forEach(function(row) {
    // Filtrer par entité si spécifiée
    if (entite && entite !== 'TOUTES' && String(row.Entite || '').trim() !== entite) return;

    const date = parseDate_(row.Date);
    if (!date) return;
    const key = Utilities.formatDate(date, Session.getScriptTimeZone() || 'Europe/Paris', 'yyyy-MM');
    if (!monthly[key]) monthly[key] = { encaissements: 0, decaissements: 0, net: 0 };
    monthly[key].encaissements += toNumber_(row.Encaissement);
    monthly[key].decaissements += toNumber_(row.Decaissement);
    monthly[key].net += toNumber_(row.Montant);
  });
  return Object.keys(monthly).sort().slice(-8).map(function(key) {
    return {
      label: key,
      encaissements: monthly[key].encaissements,
      decaissements: monthly[key].decaissements,
      net: monthly[key].net
    };
  });
}

function buildDashboardSummary_(entite) {
  const rows = getRows_(getSpreadsheet_().getSheetByName(SHEETS.MOUVEMENTS));
  const comptesNet = {};
  const latestByCompte = {};
  let encaissements = 0;
  let decaissements = 0;

  rows.forEach(function(row) {
    // Filtrer par entité si spécifiée
    if (entite && entite !== 'TOUTES' && String(row.Entite || '').trim() !== entite) return;

    const enc = toNumber_(row.Encaissement);
    const dec = toNumber_(row.Decaissement);
    const net = toNumber_(row.Montant) || (enc - dec);
    const compte = row.Compte || 'Non renseigne';
    const rowEntite = row.Entite || 'Non renseignee';
    const dt = parseDate_(row.Date) || new Date(0);
    const solde = row.Solde_cumule === '' ? null : toNumber_(row.Solde_cumule);

    encaissements += enc;
    decaissements += dec;
    comptesNet[compte] = (comptesNet[compte] || 0) + net;

    if (solde !== null) {
      if (!latestByCompte[compte] || dt.getTime() >= latestByCompte[compte].time) {
        latestByCompte[compte] = { time: dt.getTime(), solde: solde, entite: rowEntite };
      }
    }
  });

  const compteBalances = {};
  const entiteBalances = {};
  Object.keys(comptesNet).forEach(function(compte) {
    const value = latestByCompte[compte] ? latestByCompte[compte].solde : comptesNet[compte];
    compteBalances[compte] = value;
    const entite = latestByCompte[compte] ? latestByCompte[compte].entite : 'Non renseignee';
    entiteBalances[entite] = (entiteBalances[entite] || 0) + value;
  });

  return {
    tresorerieActuelle: sumMapValues_(compteBalances),
    encaissements: encaissements,
    decaissements: decaissements,
    entites: toSortedArray_(entiteBalances),
    comptes: toSortedArray_(compteBalances)
  };
}

function getComptesMap_() {
  const map = {};
  getRows_(getSpreadsheet_().getSheetByName(SHEETS.COMPTES)).forEach(function(row) {
    map[row.Compte_id] = row;
  });
  return map;
}

function getActiveRules_() {
  return getRows_(getSpreadsheet_().getSheetByName(SHEETS.REGLES)).filter(function(row) {
    return String(row.Active || '').toLowerCase() !== 'non';
  });
}

function getExistingUids_() {
  const map = {};
  getRows_(getSpreadsheet_().getSheetByName(SHEETS.MOUVEMENTS)).forEach(function(row) {
    if (row.Transaction_uid) map[row.Transaction_uid] = true;
  });
  return map;
}

function seedDefaults_() {
  const ss = getSpreadsheet_();
  const comptesSheet = ss.getSheetByName(SHEETS.COMPTES);
  const categoriesSheet = ss.getSheetByName(SHEETS.CATEGORIES);
  const reglesSheet = ss.getSheetByName(SHEETS.REGLES);
  if (comptesSheet.getLastRow() <= 1) {
    comptesSheet.getRange(2, 1, DEFAULT_COMPTES.length, DEFAULT_COMPTES[0].length).setValues(DEFAULT_COMPTES);
  }
  if (categoriesSheet.getLastRow() <= 1) {
    categoriesSheet.getRange(2, 1, DEFAULT_CATEGORIES.length, DEFAULT_CATEGORIES[0].length).setValues(DEFAULT_CATEGORIES);
  }
  mergeDefaultRules_(reglesSheet);
}

function formatWorkbook_(ss) {
  Object.keys(SHEETS).forEach(function(key) {
    const sheet = ss.getSheetByName(SHEETS[key]);
    if (!sheet) return;
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#10284d').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    try { sheet.autoResizeColumns(1, lastCol); } catch (e) {}
  });
}

function ensureSheetWithHeaders_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  const existing = sheet.getLastRow() > 0 ? sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0] : [];
  const cleaned = existing.map(function(value) { return String(value || '').trim(); });
  if (!cleaned.some(String)) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sheet;
  }
  headers.forEach(function(header) {
    if (cleaned.indexOf(header) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      cleaned.push(header);
    }
  });
  return sheet;
}

function getSheetCounts_(entite) {
  const ss = getSpreadsheet_();
  const counts = {};
  Object.keys(SHEETS).forEach(function(key) {
    const sheetName = SHEETS[key];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      counts[sheetName] = 0;
      return;
    }

    // Si pas de filtre entite, retourner le total
    if (!entite || entite === 'TOUTES') {
      counts[sheetName] = Math.max(sheet.getLastRow() - 1, 0);
    } else {
      // Sinon, compter les lignes filtrées par entité
      const rows = getRows_(sheet);
      const filtered = rows.filter(function(row) {
        return String(row.Entite || '').trim() === entite;
      });
      counts[sheetName] = filtered.length;
    }
  });
  return counts;
}

function getRows_(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values.filter(function(row) {
    return row.some(function(cell) { return String(cell || '').trim() !== ''; });
  }).map(function(row) {
    return rowToObject_(headers, row);
  });
}

function getHeaders_(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach(function(header, index) {
    obj[header] = row[index];
  });
  return obj;
}

function findRowByValue_(sheet, headerName, value) {
  const headers = getHeaders_(sheet);
  const idx = headers.indexOf(headerName);
  if (idx === -1 || sheet.getLastRow() <= 1) return { rowNumber: 0, headers: headers };
  const values = sheet.getRange(2, idx + 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(value)) return { rowNumber: i + 2, headers: headers };
  }
  return { rowNumber: 0, headers: headers };
}

function setCellByHeader_(sheet, headers, rowNumber, headerName, value) {
  const idx = headers.indexOf(headerName);
  if (idx !== -1) sheet.getRange(rowNumber, idx + 1).setValue(value);
}

function buildCategoryMap_(rows) {
  const map = {};
  rows.forEach(function(row) {
    const category = row.Categorie;
    if (!map[category]) map[category] = [];
    map[category].push(row.Sous_categorie || '');
  });
  return map;
}

function uniqueValues_(values) {
  const out = [];
  const seen = {};
  values.forEach(function(value) {
    const key = String(value || '').trim();
    if (!key || seen[key]) return;
    seen[key] = true;
    out.push(key);
  });
  return out;
}

function normalizeText_(value) {
  return String(value || '')
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseDate_(value) {
  if (value === '' || value === null || value === undefined) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value;
  if (typeof value === 'number') {
    if (value < -1000 || value > 60000) return null;
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const date = new Date(epoch.getTime() + value * 86400000);
    return isNaN(date.getTime()) ? null : date;
  }
  const raw = String(value).trim();
  if (!raw) return null;
  const parts = raw.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (parts) {
    const year = parts[3].length === 2 ? Number('20' + parts[3]) : Number(parts[3]);
    const date = new Date(year, Number(parts[2]) - 1, Number(parts[1]));
    return isNaN(date.getTime()) ? null : date;
  }
  const parsed = new Date(raw);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function formatDateIso_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone() || 'Europe/Paris', 'yyyy-MM-dd');
}

function toNumber_(value) {
  if (value === '' || value === null || value === undefined) return 0;
  if (typeof value === 'number') return value;
  const normalized = String(value)
    .replace(/\s/g, '')
    .replace(/EUR/gi, '')
    .replace(/â‚¬/g, '')
    .replace(/\.(?=\d{3}(\D|$))/g, '')
    .replace(',', '.');
  const parsed = parseFloat(normalized);
  return isNaN(parsed) ? 0 : parsed;
}

function toSortedArray_(map) {
  return Object.keys(map).map(function(key) {
    return { label: key, value: map[key] };
  }).sort(function(a, b) {
    return Math.abs(b.value) - Math.abs(a.value);
  });
}

function sumMapValues_(map) {
  return Object.keys(map).reduce(function(total, key) {
    return total + toNumber_(map[key]);
  }, 0);
}

function getDocumentAiConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    projectId: String(props.getProperty('DOCUMENT_AI_PROJECT_ID') || DOCUMENT_AI.projectId || '').trim(),
    location: String(props.getProperty('DOCUMENT_AI_LOCATION') || DOCUMENT_AI.location || 'eu').trim(),
    processorId: String(props.getProperty('DOCUMENT_AI_PROCESSOR_ID') || DOCUMENT_AI.processorId || '').trim()
  };
}

function isOcrConfigured_() {
  const config = getDocumentAiConfig_();
  return !!(config.projectId && config.location && config.processorId);
}

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty('SPREADSHEET_ID') || SPREADSHEET_ID;
  return SpreadsheetApp.openById(ssId);
}

function initializeSecureIds_() {
  /**
   * Initialise les IDs sécurisés dans Properties
   * À appeler au premier démarrage ou après migration
   */
  const props = PropertiesService.getScriptProperties();

  // Initialiser le SPREADSHEET_ID si pas encore fait
  if (!props.getProperty('SPREADSHEET_ID')) {
    props.setProperty('SPREADSHEET_ID', SPREADSHEET_ID);
  }

  // Initialiser les Drive roots si pas encore fait
  if (!props.getProperty('DRIVE_FACTURES_ROOTS_DEFAULT')) {
    props.setProperty('DRIVE_FACTURES_ROOTS_DEFAULT', JSON.stringify(DRIVE_FACTURES_ROOTS));
  }

  return {
    success: true,
    message: 'IDs sécurisés initialisés dans Google Properties Service'
  };
}









function installAutomaticSyncTrigger() {
  removeAutomaticSyncTriggers();
  ScriptApp.newTrigger('runScheduledMaintenance')
    .timeBased()
    .everyMinutes(5)
    .create();
  return {
    success: true,
    message: 'Declencheur automatique active toutes les 5 minutes.',
    automaticSync: getAutomaticSyncStatus()
  };
}

function removeAutomaticSyncTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    const handler = trigger.getHandlerFunction();
    if (handler === 'runScheduledMaintenance') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  return {
    success: true,
    message: 'Declencheurs automatiques supprimes.',
    automaticSync: getAutomaticSyncStatus()
  };
}

function getAutomaticSyncStatus() {
  const triggers = ScriptApp.getProjectTriggers().filter(function(trigger) {
    return trigger.getHandlerFunction() === 'runScheduledMaintenance';
  });
  return {
    enabled: triggers.length > 0,
    count: triggers.length,
    frequency: triggers.length > 0 ? 'Toutes les 5 minutes' : 'Desactive'
  };
}

function runScheduledMaintenance() {
  installOrUpdateWorkbook();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return {
      success: false,
      skipped: true,
      message: 'Un traitement automatique est deja en cours.'
    };
  }

  try {
    const driveResult = syncDriveInvoices({ source: 'AUTO_TRIGGER' });
    let ocrResult = {
      success: false,
      skipped: true,
      message: 'OCR non configure.'
    };

    if (isOcrConfigured_()) {
      ocrResult = runInvoiceOcr({ source: 'AUTO_TRIGGER' });
    }

    const reconciliationResult = runAutomaticReconciliation({ source: 'AUTO_TRIGGER' });

    return {
      success: true,
      drive: driveResult,
      ocr: ocrResult,
      reconciliation: reconciliationResult,
      automaticSync: getAutomaticSyncStatus(),
      executedAt: new Date()
    };
  } finally {
    lock.releaseLock();
  }
}





function mergeDefaultRules_(reglesSheet) {
  const existingRows = getRows_(reglesSheet);
  const existingKeys = {};
  existingRows.forEach(function(row) {
    const key = [
      normalizeText_(row.Champ_source || ''),
      normalizeText_(row.Operateur || ''),
      normalizeText_(row.Valeur || ''),
      normalizeText_(row.Categorie || ''),
      normalizeText_(row.Sous_categorie || ''),
      normalizeText_(row.Sens || '')
    ].join('|');
    existingKeys[key] = true;
  });

  const missing = DEFAULT_REGLES.filter(function(rule) {
    const key = [
      normalizeText_(rule[4] || ''),
      normalizeText_(rule[5] || ''),
      normalizeText_(rule[6] || ''),
      normalizeText_(rule[7] || ''),
      normalizeText_(rule[8] || ''),
      normalizeText_(rule[9] || '')
    ].join('|');
    return !existingKeys[key];
  });

  if (missing.length) {
    reglesSheet.getRange(reglesSheet.getLastRow() + 1, 1, missing.length, missing[0].length).setValues(missing);
  }
}

// ============================================================================
// CORRECTION PHASE 1 - Fonctions manquantes pour l'import bancaire
// ============================================================================

function normalizeImportedRow_(sourceRow, mapping, compte, entite, sourceName, importId) {
  /**
   * Normalise une ligne importée depuis un fichier bancaire
   * @param sourceRow {object} Ligne brute du fichier importé
   * @param mapping {object} Mapping des colonnes (colonnes importées -> colonnes cibles)
   * @param compte {object} Infos du compte (depuis HEADERS.Comptes)
   * @param entite {string} Entité (ex: "ALTERACIG")
   * @param sourceName {string} Nom de la source (ex: "BNP_EXPORT")
   * @param importId {string} ID unique de cet import
   * @returns {object} Transaction normalisée ou null si invalide
   */
  if (!sourceRow || !mapping) return null;

  // Mapper les colonnes selon le mapping fourni
  const getField = function(mappingKey) {
    const fieldName = mapping[mappingKey];
    return fieldName ? sourceRow[fieldName] : '';
  };

  // Extraire les données brutes
  const dateRaw = getField('Date') || getField('DateOperation') || getField('date');
  const montantRaw = getField('Montant') || getField('amount') || getField('Amount');
  const libelleRaw = getField('Libelle') || getField('Label') || getField('label');
  const dateDiffereeRaw = getField('DateValeur') || getField('ValueDate') || '';
  const refBancRaw = getField('Reference') || getField('RefBanc') || '';

  // Parser et normaliser
  const date = parseDate_(dateRaw);
  if (!date) return null; // Date obligatoire

  const dateValeur = parseDate_(dateDiffereeRaw) || date;
  const montant = toNumber_(montantRaw);
  const libelle = String(libelleRaw || '').trim();

  if (!libelle || montant === 0) return null; // Libellé et montant obligatoires

  // Déterminer le sens (Encaissement ou Décaissement)
  const sens = montant > 0 ? 'Encaissement' : 'Décaissement';
  const montantAbs = Math.abs(montant);

  // Construire Transaction_uid unique pour dédoublonnage
  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone() || 'Europe/Paris', 'yyyyMMdd');
  const transactionUid = [
    entite,
    compte.Compte_id || '',
    dateStr,
    normalizeText_(libelle),
    montantAbs.toFixed(2)
  ].join('|');

  // Retourner transaction normalisée
  return {
    Date: date,
    Date_valeur: dateValeur,
    Libelle: libelle,
    Libelle_normalise: normalizeText_(libelle),
    Entite: entite,
    Groupe: compte.Groupe || '',
    Compte: compte.Nom_compte || '',
    Compte_id: compte.Compte_id || '',
    Banque: compte.Banque || '',
    Categorie: '',
    Sous_categorie: '',
    Mode_reglement: 'Virement',
    Sens: sens,
    Encaissement: sens === 'Encaissement' ? montantAbs : 0,
    Decaissement: sens === 'Décaissement' ? montantAbs : 0,
    Montant: montant,
    Devise: 'EUR',
    Solde_cumule: 0, // Sera calculé après tri
    Reference_bancaire: String(refBancRaw || '').trim(),
    Source_import: sourceName,
    Import_id: importId,
    Transaction_uid: transactionUid,
    Statut_classement: 'A_CLASSER',
    Statut_rapprochement: 'NON_RAPPROCHE',
    Commentaire: 'Importé depuis ' + sourceName
  };
}

function findMatchingRule_(tx, rules) {
  /**
   * Cherche une règle de classement automatique pour une transaction
   * @param tx {object} Transaction normalisée (résultat de normalizeImportedRow_)
   * @param rules {array} Liste des règles actives (depuis SHEETS.REGLES)
   * @returns {object|null} Règle matchante ou null
   */
  if (!tx || !rules || !Array.isArray(rules)) return null;

  let bestMatch = null;
  let bestScore = 0;

  rules.forEach(function(rule) {
    const score = scoreRule_(tx, rule);
    if (score > bestScore) {
      bestScore = score;
      bestMatch = rule;
    }
  });

  const threshold = getSettingsConfig_().classificationConfidenceThreshold || 70;
  return bestScore >= threshold ? bestMatch : null;
}

function scoreRule_(tx, rule) {
  /**
   * Score une règle par rapport à une transaction
   * Plus la priorité est haute, meilleur le score
   * @returns {number} Score 0-100
   */
  if (!rule || !tx) return 0;

  // La règle doit s'appliquer à l'entité ou être générique
  if (rule.Entite && rule.Entite !== '' && rule.Entite !== tx.Entite) {
    return 0;
  }

  // La règle doit s'appliquer au compte ou être générique
  if (rule.Compte_id && rule.Compte_id !== '' && rule.Compte_id !== tx.Compte_id) {
    return 0;
  }

  // Vérifier la condition de la règle
  const champ = rule.Champ_source || 'Libelle_normalise';
  const valeur = String(tx[champ] || '').trim();
  const operateur = (rule.Operateur || 'contains').toLowerCase();
  const criterium = normalizeText_(rule.Valeur || '');

  let conditionMatches = false;
  if (operateur === 'contains') {
    conditionMatches = normalizeText_(valeur).indexOf(criterium) !== -1;
  } else if (operateur === 'equals') {
    conditionMatches = normalizeText_(valeur) === criterium;
  } else if (operateur === 'startswith') {
    conditionMatches = normalizeText_(valeur).indexOf(criterium) === 0;
  }

  if (!conditionMatches) return 0;

  // Si la condition matche, le score est la priorité (50-100)
  const priorite = Number(rule.Priorite || 50);
  return Math.max(Math.min(priorite, 100), 50);
}

// ============================================================================
// OPTION C - Fonction pour charger les données filtrées par entité
// ============================================================================

function getDataForEntite(entite) {
  /**
   * Récupère les données filtrées pour une entité spécifique
   * Appelée par le frontend quand on change d'entité
   * @param entite {string} Entité sélectionnée (ou "TOUTES")
   * @returns {object} Données complètes filtrées pour cette entité
   */
  return {
    dashboard: buildDashboardSummary_(entite),
    aClasser: getTransactionsToClassify(50, entite),
    recentMovements: getRecentMovements(20, entite),
    recentPrevisions: getRecentPrevisions(50, entite),
    recentScenarios: getRecentScenarios(50, entite),
    recentRapprochements: getRecentRapprochements(50, entite),
    facturesToVerify: getFacturesToVerify(50, entite),
    transactionsWithoutInvoice: getTransactionsWithoutInvoice(50, entite),
    alerts: getAlerts_(12, entite),
    forecastSummary: buildForecastSummary_(entite),
    scenarioSummary: buildScenarioSummary_(entite),
    periodAnalytics: buildPeriodAnalytics_(entite),
    graphData: getGraphData_(entite),
    facturesToPay: getFacturesToPay(100, entite),
    counts: getSheetCounts_(entite)
  };
}
