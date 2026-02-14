# Grist Asset Manager Widget

> **Author:** Said Hamadou
> **License:** Apache-2.0

---

*[English](#english) | [Français](#français)*

---

<a id="english"></a>

## 🇬🇧 English

Real estate asset management widget for Grist. Full CRUD interface with advanced search, dashboard, and role-based access control.

**Widget URL:** `https://grist-asset-manager-widget.vercel.app/index.html`

### 🚀 Quick Start

1. In Grist, click **"Add widget to page"**
2. Select **"Custom"** as the widget type
3. Enter the custom widget URL:
   ```
   https://grist-asset-manager-widget.vercel.app/index.html
   ```
4. Set the access level to **"Full document access"**
5. Done! Start managing your assets.

### 📋 Features

- **Complete CRUD** for real estate assets (create, read, update, delete)
- **Advanced search**: Classic search and multi-criteria filtering
- **Sortable columns** with pagination
- **Dashboard** with statistics and key metrics
- **Role-based access**: Owner, editor, and manager roles
- **Manager management**: Assign and manage property managers
- **30+ property fields**: Reference, address, municipality, area, type, occupation, etc.
- **Excel/CSV export** capabilities
- **Bilingual interface** (French / English)

### 🔒 Security

- Role-based access control (owner/editor/manager)
- XSS protection on all user inputs
- Identifier sanitization for Grist compatibility

### 🛠️ Local Development

```bash
git clone https://github.com/isaytoo/grist-asset-manager-widget.git
cd grist-asset-manager-widget
python3 -m http.server 8585
```

Then in Grist, use: `http://localhost:8585/index.html`

### ⚙️ Required Configuration

The widget requires **Full document access** to:
- Manage asset tables (`BM_Biens`, `BM_Gestionnaires`)
- Read and write property data
- Manage user roles and permissions

### 📁 File Structure

```
grist-asset-manager-widget/
├── index.html       # Widget UI (HTML + CSS)
├── widget.js        # JavaScript logic (CRUD, search, dashboard, i18n)
├── standalone/      # Standalone version
├── package.json     # Metadata
├── vercel.json      # Vercel config (iframe headers)
├── .gitignore
└── README.md
```

---

<a id="français"></a>

## 🇫🇷 Français

Widget de gestion de biens immobiliers pour Grist. Interface CRUD complète avec recherche avancée, tableau de bord et gestion des droits par rôle.

**URL du widget :** `https://grist-asset-manager-widget.vercel.app/index.html`

### 🚀 Utilisation rapide

1. Dans Grist, cliquez sur **"Ajouter un widget à la page"**
2. Sélectionnez **"Personnalisé"** comme type de widget
3. Entrez l'URL :
   ```
   https://grist-asset-manager-widget.vercel.app/index.html
   ```
4. Définissez le niveau d'accès sur **"Full document access"**
5. C'est prêt ! Commencez à gérer vos biens.

### 📋 Fonctionnalités

- **CRUD complet** pour les biens immobiliers (créer, lire, modifier, supprimer)
- **Recherche avancée** : recherche classique et filtrage multi-critères
- **Colonnes triables** avec pagination
- **Tableau de bord** avec statistiques et indicateurs clés
- **Gestion des rôles** : propriétaire, éditeur et gestionnaire
- **Gestion des gestionnaires** : assignation et gestion des gestionnaires de biens
- **30+ champs** : référence, adresse, commune, surface, type, occupation, etc.
- **Export Excel/CSV**
- **Interface bilingue** (Français / Anglais)

### 🔒 Sécurité

- Contrôle d'accès par rôle (propriétaire/éditeur/gestionnaire)
- Protection XSS sur toutes les entrées utilisateur
- Sanitization des identifiants pour compatibilité Grist

### 🛠️ Développement local

```bash
git clone https://github.com/isaytoo/grist-asset-manager-widget.git
cd grist-asset-manager-widget
python3 -m http.server 8585
```

Puis dans Grist, utilisez : `http://localhost:8585/index.html`

### ⚙️ Configuration requise

Le widget nécessite un **accès complet au document** pour :
- Gérer les tables de biens (`BM_Biens`, `BM_Gestionnaires`)
- Lire et écrire les données des biens
- Gérer les rôles et permissions utilisateurs

### 📁 Structure des fichiers

```
grist-asset-manager-widget/
├── index.html       # Interface HTML + CSS du widget
├── widget.js        # Logique JavaScript (CRUD, recherche, dashboard, i18n)
├── standalone/      # Version autonome
├── package.json     # Métadonnées
├── vercel.json      # Configuration Vercel (headers iframe)
├── .gitignore
└── README.md
```

---

## 🔗 Resources / Ressources

- [Grist Custom Widgets Documentation](https://support.getgrist.com/widget-custom/)
- [Grist Plugin API](https://support.getgrist.com/code/modules/grist_plugin_api/)
- [GristUp Widget Marketplace](https://www.gristup.fr)
