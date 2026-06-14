#!/usr/bin/env python3
"""
Create a professional PowerPoint presentation for PFA internship defense.
Uses raw Office Open XML since python-pptx is not available.
"""

import zipfile
import os
import shutil
from copy import deepcopy

# ============================================================================
# CONFIGURATION
# ============================================================================

OUTPUT_FILE = "/projects/sandbox/Presentation_PFA/Presentation_Stage_PFA.pptx"

# EMU (English Metric Units): 1 inch = 914400 EMU
# Slide dimensions: 10 inches x 7.5 inches (standard 4:3) -> use 16:9 widescreen
SLIDE_WIDTH = 12192000   # 10 inches in EMU (will override to widescreen)
SLIDE_HEIGHT = 6858000   # 7.5 inches in EMU

# Colors (in RRGGBB format for OOXML)
PRIMARY = "1E40AF"       # Blue
SECONDARY = "6B7280"     # Gray
ACCENT = "059669"        # Green
WHITE = "FFFFFF"
DARK = "1F2937"
LIGHT_BG = "F8FAFC"

# ============================================================================
# SLIDE CONTENT DEFINITIONS
# ============================================================================

slides_content = [
    # Slide 1: Title
    {
        "type": "title",
        "title": "IA Pratique pour une Plateforme\nE-Commerce d'Affiliation",
        "subtitle": "Recommandation de Produits, Intelligence d'Approvisionnement\net Analytics Metier",
        "footer": "CHELLAK Aymane | Encadrant: MOUSSAID Ahmed / ECHCHAHID Khalid\nFST Mohammedia | Maroc Affiliate Network | 2025-2026"
    },
    # Slide 2: Plan
    {
        "type": "content",
        "title": "Plan de la Presentation",
        "bullets": [
            "1. Contexte General et Problematique",
            "2. Presentation de l'Entreprise",
            "3. Methodologie de Travail",
            "4. Architecture Technique",
            "5. Module 1 : Analytics Dashboard",
            "6. Module 2 : Systeme de Recommandation TrackA",
            "7. Module 3 : Enrichissement de Stock",
            "8. Difficultes et Solutions",
            "9. Bilan et Perspectives",
        ]
    },
    # Slide 3: Contexte
    {
        "type": "content",
        "title": "Contexte General",
        "bullets": [
            "E-commerce en croissance exponentielle au Maroc",
            "Modele d'affiliation : vendeurs independants sans risque d'inventaire",
            "Plateforme multi-roles : Admin, Vendeur, Entrepot, Livreur",
            "Gestion complete : catalogue, stock, commandes, livraison, facturation",
        ]
    },
    # Slide 4: Problematique
    {
        "type": "content",
        "title": "Problematique",
        "bullets": [
            "Comment fournir une vision claire des performances ? (Analytics)",
            "Comment aider les vendeurs a decouvrir les bons produits ? (Recommandation)",
            "Comment optimiser l'approvisionnement ? (Stock)",
            "",
            "3 missions principales :",
            "  -> Systeme d'analytics avance (pilotage operationnel)",
            "  -> Moteur de recommandation intelligent (personnalisation)",
            "  -> Module d'enrichissement de stock (approvisionnement)",
        ]
    },
    # Slide 5: Entreprise
    {
        "type": "content",
        "title": "Maroc Affiliate Network",
        "bullets": [
            "Plateforme e-commerce par affiliation au Maroc",
            "Modele economique :",
            "  - Admin acquiert produits aupres de fournisseurs",
            "  - Stockage dans entrepots centraux et regionaux",
            "  - Vendeurs affilies selectionnent et vendent les produits",
            "  - Plateforme gere la logistique de livraison",
            "  - Vendeur percoit un benefice sur chaque vente",
            "",
            "4 acteurs : Admin, Vendeur, Entrepot, Livreur",
        ]
    },
    # Slide 6: Existant
    {
        "type": "content",
        "title": "Existant Avant le Stage",
        "bullets": [
            "Systeme d'authentification et autorisation (RBAC)",
            "Gestion du catalogue produits et variantes",
            "Systeme de commandes basique",
            "Integration societes de livraison externes",
            "Point de vente (POS) avec support hors-ligne",
            "",
            "Besoins identifies :",
            "  -> Pilotage operationnel (pas de dashboards)",
            "  -> Decouverte produit (catalogue en croissance)",
            "  -> Processus d'approvisionnement (manuel et lent)",
        ]
    },
    # Slide 7: Methodologie
    {
        "type": "content",
        "title": "Methodologie de Travail",
        "bullets": [
            "Approche Agile iterative (inspiree Scrum)",
            "Sprints de 2 semaines avec objectifs definis",
            "Daily standups + Sprint review + Retrospective",
            "",
            "Outils :",
            "  - Git + GitHub (versionnement, PR, code review)",
            "  - Vercel (deploiement continu frontend)",
            "  - Railway (deploiement microservices)",
            "  - Discord (communication asynchrone)",
            "",
            "Workflow : branche feature -> PR -> review -> deploy preview -> merge",
        ]
    },
    # Slide 8: Architecture Globale
    {
        "type": "content",
        "title": "Architecture Globale",
        "bullets": [
            "Architecture hybride : monolithe modulaire + microservices",
            "",
            "Next.js 15 (App Router) -> Frontend + Backend (Vercel)",
            "FastAPI (Python) -> Service de recommandation (Railway)",
            "PostgreSQL 16 (Neon) -> Base de donnees principale",
            "Redis 7.4 -> Cache + Queue de jobs",
            "BullMQ Workers -> Traitements asynchrones",
            "",
            "Justifications :",
            "  - Server Components + streaming SSR",
            "  - PostgreSQL pour agregations analytiques avancees",
            "  - Microservice Python pour isolation et scaling",
        ]
    },
    # Slide 9: Stack Technique
    {
        "type": "content",
        "title": "Stack Technologique",
        "bullets": [
            "Frontend : Next.js 15, TypeScript, Tailwind CSS 4, Radix UI",
            "Backend : Server Actions, Prisma 7, PostgreSQL 16",
            "Auth : NextAuth v5 (RBAC granulaire)",
            "Charts : Recharts 2 | Tables : TanStack Table",
            "Recommandation : FastAPI, SQLAlchemy, Redis",
            "Queue/Jobs : BullMQ + Redis",
            "Monitoring : Prometheus + Grafana",
            "Deploy : Vercel (frontend) + Railway (services)",
            "i18n : 3 langues (fr, en, ar) avec support RTL",
        ]
    },
    # Slide 10: Module Analytics - Intro
    {
        "type": "section",
        "title": "Module 1\nAnalytics Dashboard",
        "subtitle": "Pilotage operationnel de la plateforme"
    },
    # Slide 11: Analytics - Architecture
    {
        "type": "content",
        "title": "Analytics : Architecture",
        "bullets": [
            "8 tableaux de bord : Executif, Commandes, Livraison,",
            "  Produits, Variantes, Vendeurs, Clients, Entrepots",
            "",
            "Architecture en couches :",
            "  1. Presentation (React Server/Client Components)",
            "  2. Page (page.tsx avec Suspense)",
            "  3. Donnees (queries.ts + unstable_cache)",
            "  4. Persistance (Prisma $queryRaw -> PostgreSQL)",
            "",
            "Pattern identique pour chaque sous-page (maintenabilite)",
        ]
    },
    # Slide 12: Analytics - Fonctionnalites
    {
        "type": "content",
        "title": "Analytics : Fonctionnalites",
        "bullets": [
            "50+ indicateurs KPI avec comparaison periode-sur-periode",
            "",
            "Commandes : funnel de conversion, performance par ville",
            "Livraison : comparaison transporteurs, heatmap societe x ville",
            "Produits : quadrant performance vs stock",
            "Variantes : heatmap couleur x taille (5 modes)",
            "Vendeurs : leaderboard, scoring churn, risque",
            "Clients : segmentation LTV (1, 2-3, 4-6, 7+)",
            "Entrepots : comparaison multi-criteres, rotation stock",
        ]
    },
    # Slide 13: Analytics - Optimisations
    {
        "type": "content",
        "title": "Analytics : Optimisations",
        "bullets": [
            "Parallelisation : Promise.all() -> 11 requetes en parallele",
            "  Temps = max(requetes) au lieu de somme",
            "",
            "Cache intelligent : TTL 300s + invalidation on-demand",
            "  revalidateTag() dans les server actions de mutation",
            "",
            "SQL optimise : COUNT(*) FILTER + CTE en une seule passe",
            "",
            "UX : Suspense + skeleton anime + streaming SSR",
            "",
            "Resultats : < 1s (cache hit), < 3s (cache miss)",
        ]
    },
    # Slide 14: Module Recommandation - Intro
    {
        "type": "section",
        "title": "Module 2\nSysteme de Recommandation TrackA",
        "subtitle": "Microservice autonome de recommandation personnalisee"
    },
    # Slide 15: TrackA - Architecture
    {
        "type": "content",
        "title": "TrackA : Architecture Microservice",
        "bullets": [
            "Microservice independant (FastAPI + Python 3.11)",
            "Deploye sur Railway : 3 processus depuis 1 image Docker",
            "",
            "Composants :",
            "  - API (FastAPI + Uvicorn) : sert les recommandations REST",
            "  - Worker (ThreadPool x 4) : calcule et prechauffe le cache",
            "  - Scheduler (APScheduler) : enfile les vendeurs toutes les 30min",
            "",
            "Infrastructure : PostgreSQL + Redis (LRU 256MB)",
            "Observabilite : Prometheus + Grafana (7 metriques)",
        ]
    },
    # Slide 16: TrackA - Algorithme
    {
        "type": "content",
        "title": "TrackA : Algorithme a 5 Signaux",
        "bullets": [
            "Ensemble pondere de 5 signaux positifs + penalite :",
            "",
            "  - Popularite (25%) : volume commandes recentes [global]",
            "  - Historique (35%) : commandes passees + affinite [vendeur]",
            "  - Recence (20%) : courbe en U inverse [vendeur]",
            "  - Nouveaute (10%) : age du produit [global]",
            "  - Engagement (10%) : likes + commentaires [global]",
            "",
            "Score final = sum(wi * Si) - penalite(annulations)",
            "",
            "Diversite categorielle + exploration aleatoire (5%)",
        ]
    },
    # Slide 17: TrackA - Adaptation Cold Start
    {
        "type": "content",
        "title": "TrackA : Adaptation Dynamique des Poids",
        "bullets": [
            "Poids adaptes selon le profil vendeur :",
            "",
            "Nouveau (0 cmd)      : Pop=52.5% | Hist=0%    | Nouv=37.5%",
            "Demarrage froid (<5)  : Pop=33.4% | Hist=14%   | Nouv=22.6%",
            "Standard (5-50)      : Pop=25%   | Hist=35%   | Rec=20%",
            "Experimente (>50)    : Pop=12.5% | Hist=43.8% | Rec=23.8%",
            "",
            "Courbe de recence en U inverse :",
            "  Zone optimale = 14-40 jours (cycle reapprovisionnement)",
        ]
    },
    # Slide 18: TrackA - Pipeline
    {
        "type": "content",
        "title": "TrackA : Pipeline de Serving",
        "bullets": [
            "Strategie multi-couches (5 fallbacks) :",
            "",
            "  1. Redis cache hit -> < 5ms",
            "  2. Snapshot frais -> warm Redis + retourner",
            "  3. Snapshot perime -> retourner stale + queue refresh",
            "  4. Aucun snapshot -> queue refresh + produits populaires",
            "  5. Erreur totale -> fallback popularite globale",
            "",
            "Evenements declencheurs :",
            "  order-placed, product-engaged, product-updated, seller-created",
            "  -> Invalidation immediate + refresh asynchrone",
        ]
    },
    # Slide 19: TrackA - Performance
    {
        "type": "content",
        "title": "TrackA : Resultats en Production",
        "bullets": [
            "Performance mesuree (150+ produits, 80+ vendeurs, 3000+ commandes) :",
            "",
            "  - Temps reponse (cache hit) : < 5 ms",
            "  - Temps reponse (calcul complet) : 150-300 ms",
            "  - Debit worker (4 threads) : ~40 refresh/min",
            "  - Taux de cache hit : ~85%",
            "  - Requetes SQL par calcul : 8 (batch, sans N+1)",
            "",
            "Qualite :",
            "  - Couverture catalogue : 92%",
            "  - Diversite : 4.2 categories / set de 30",
            "  - Personnalisation : 68% pour vendeurs 5+ commandes",
        ]
    },
    # Slide 20: Module Stock - Intro
    {
        "type": "section",
        "title": "Module 3\nEnrichissement de Stock",
        "subtitle": "Optimisation du processus d'approvisionnement"
    },
    # Slide 21: Stock - Architecture
    {
        "type": "content",
        "title": "Stock : Architecture et Modele de Donnees",
        "bullets": [
            "Pipeline : Fournisseur -> Scan/Selection -> PurchaseBatch",
            "           -> Main Warehouse -> Transfert -> Entrepot destination",
            "",
            "Modeles Prisma :",
            "  - PurchaseBatch : lot d'achat (adminId, totalPrice, notes)",
            "  - Purchase : ligne d'achat (variant, qty, remainingQty, status)",
            "  - StockMovement : tracabilite (type: PURCHASE)",
            "  - WarehouseStock : stock par variante par entrepot",
            "",
            "Transaction atomique (timeout 60s) pour chaque lot",
        ]
    },
    # Slide 22: Stock - Fonctionnalites
    {
        "type": "content",
        "title": "Stock : Deux Modes d'Entree",
        "bullets": [
            "Mode Scanner QR (pistolet barcode USB) :",
            "  - Auto-focus permanent + file d'attente client",
            "  - Traitement sequentiel (0% perte de scans)",
            "  - Format QR : {productCode}-{variantAssignmentId}-{variantId}",
            "  - Scan rapide : support 3 scans/seconde",
            "",
            "Mode Manuel :",
            "  - Recherche textuelle (debounced 300ms)",
            "  - Filtres : categorie, fournisseur, statut",
            "  - Modal de selection variantes (quantite individuelle)",
            "  - Fonction 'Appliquer a toutes'",
        ]
    },
    # Slide 23: Stock - Transfert
    {
        "type": "content",
        "title": "Stock : Panier et Transfert",
        "bullets": [
            "Panier d'achat :",
            "  - Position sticky, controles quantite inline",
            "  - Badge methode (Scanner/Manuel), resume temps reel",
            "  - Sauvegarde sessionStorage (protection contre perte)",
            "",
            "Transfert Purchase -> Entrepot :",
            "  - Selection lignes + quantite ajustable",
            "  - Choix entrepot destination",
            "  - Reutilisation pipeline executeTransfer existant",
            "  - Transfert partiel supporte (remainingQuantity)",
            "  - 0 modification du code existant (1200+ lignes)",
        ]
    },
    # Slide 24: Difficultes
    {
        "type": "content",
        "title": "Difficultes et Solutions",
        "bullets": [
            "Analytics :",
            "  - Requetes lentes -> Promise.all() + cache TTL 300s",
            "  - Staleness cache -> revalidateTag on-demand",
            "",
            "Recommandation :",
            "  - Cold-start -> adaptation dynamique des poids",
            "  - N+1 queries (150 req/calcul) -> batch GROUP BY (8 req fixes)",
            "  - Pannes -> 5 niveaux de fallback (degradation gracieuse)",
            "",
            "Stock :",
            "  - Scans perdus (15-20%) -> file d'attente client (0%)",
            "  - Atomicite -> transaction Prisma unique 60s",
            "  - Integration -> adaptateur (Open/Closed principle)",
        ]
    },
    # Slide 25: Bilan
    {
        "type": "content",
        "title": "Bilan des Realisations",
        "bullets": [
            "Analytics Dashboard :",
            "  -> 8 dashboards, 50+ KPIs, < 1s en cache",
            "",
            "Systeme de Recommandation TrackA :",
            "  -> Microservice autonome, < 5ms, 85% cache hit",
            "  -> Algorithme 5 signaux adaptatif",
            "",
            "Enrichissement de Stock :",
            "  -> Dual-mode (QR + manuel), 0% perte",
            "  -> Temps reception : 2h -> 15min",
            "",
            "Tous les objectifs atteints et en production",
        ]
    },
    # Slide 26: Perspectives
    {
        "type": "content",
        "title": "Perspectives",
        "bullets": [
            "Modele ML supervise pour recommandations (A/B testing)",
            "Vues materialisees PostgreSQL pour agregations lourdes",
            "Import CSV/Excel pour achats en gros volume",
            "Application mobile dediee pour scan d'inventaire",
            "Dashboard temps reel avec WebSockets",
            "",
            "Apports personnels :",
            "  - Conception architectures distribuees",
            "  - Compromis complexite technique / valeur metier",
            "  - Validation par mesures en production",
            "  - Pragmatisme dans les decisions d'ingenierie",
        ]
    },
    # Slide 27: Merci
    {
        "type": "title",
        "title": "Merci pour votre attention",
        "subtitle": "Questions ?",
        "footer": "CHELLAK Aymane | Stage PFA 2025-2026\nMaroc Affiliate Network | FST Mohammedia"
    },
]

# ============================================================================
# OOXML TEMPLATE COMPONENTS
# ============================================================================

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  {slide_overrides}
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"""

def make_presentation_xml(num_slides):
    slide_list = "\n    ".join(
        f'<p:sldId id="{256+i}" r:id="rId{i+4}"/>' for i in range(num_slides)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                saveSubsetFonts="1">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    {slide_list}
  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"""

def make_presentation_rels(num_slides):
    slide_rels = "\n  ".join(
        f'<Relationship Id="rId{i+4}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i+1}.xml"/>'
        for i in range(num_slides)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>
  {slide_rels}
</Relationships>"""

PRES_PROPS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"""

THEME = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="PFA Theme">
  <a:themeElements>
    <a:clrScheme name="PFA Colors">
      <a:dk1><a:srgbClr val="1F2937"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1E40AF"/></a:dk2>
      <a:lt2><a:srgbClr val="F8FAFC"/></a:lt2>
      <a:accent1><a:srgbClr val="1E40AF"/></a:accent1>
      <a:accent2><a:srgbClr val="059669"/></a:accent2>
      <a:accent3><a:srgbClr val="6B7280"/></a:accent3>
      <a:accent4><a:srgbClr val="DC2626"/></a:accent4>
      <a:accent5><a:srgbClr val="7C3AED"/></a:accent5>
      <a:accent6><a:srgbClr val="EA580C"/></a:accent6>
      <a:hlink><a:srgbClr val="1E40AF"/></a:hlink>
      <a:folHlink><a:srgbClr val="6B7280"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="PFA Fonts">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>"""

SLIDE_MASTER = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
    <p:sldLayoutId id="2147483650" r:id="rId2"/>
  </p:sldLayoutIdLst>
</p:sldMaster>"""

SLIDE_MASTER_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"""

SLIDE_LAYOUT_1 = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             type="title" preserve="1">
  <p:cSld name="Title Slide">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"""

SLIDE_LAYOUT_2 = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             type="obj" preserve="1">
  <p:cSld name="Title and Content">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"""

SLIDE_LAYOUT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"""

SLIDE_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout{layout_id}.xml"/>
</Relationships>"""

# ============================================================================
# SLIDE XML GENERATORS
# ============================================================================

def make_text_run(text, font_size, bold=False, color=None, font_name="Calibri"):
    """Create a text run XML fragment."""
    bold_xml = '<a:rPr lang="fr-FR" dirty="0"'
    if bold:
        bold_xml += ' b="1"'
    if font_size:
        bold_xml += f' sz="{font_size}"'
    bold_xml += '>'
    
    color_xml = ""
    if color:
        color_xml = f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
    
    font_xml = f'<a:latin typeface="{font_name}"/><a:cs typeface="{font_name}"/>'
    
    bold_xml += color_xml + font_xml + '</a:rPr>'
    
    # Escape XML special characters
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    
    return f'<a:r>{bold_xml}<a:t>{text}</a:t></a:r>'

def make_paragraph(text, font_size, bold=False, color=None, align="l", spc_after=0, font_name="Calibri"):
    """Create a paragraph XML fragment."""
    align_xml = f' algn="{align}"' if align else ''
    spc_xml = f'<a:spcAft><a:spcPts val="{spc_after}"/></a:spcAft>' if spc_after else ''
    
    return f"""<a:p>
      <a:pPr{align_xml}>{spc_xml}</a:pPr>
      {make_text_run(text, font_size, bold, color, font_name)}
    </a:p>"""

def make_textbox(x, y, cx, cy, paragraphs_xml):
    """Create a textbox shape."""
    return f"""<p:sp>
      <p:nvSpPr>
        <p:cNvPr id="0" name="TextBox"/>
        <p:cNvSpPr txBox="1"/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="{x}" y="{y}"/>
          <a:ext cx="{cx}" cy="{cy}"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
      </p:spPr>
      <p:txBody>
        <a:bodyPr wrap="square" rtlCol="0" anchor="t">
          <a:spAutoFit/>
        </a:bodyPr>
        <a:lstStyle/>
        {paragraphs_xml}
      </p:txBody>
    </p:sp>"""

def make_rect(x, y, cx, cy, fill_color, border_color=None):
    """Create a filled rectangle shape."""
    border_xml = '<a:ln><a:noFill/></a:ln>' if not border_color else f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{border_color}"/></a:solidFill></a:ln>'
    return f"""<p:sp>
      <p:nvSpPr>
        <p:cNvPr id="0" name="Rectangle"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="{x}" y="{y}"/>
          <a:ext cx="{cx}" cy="{cy}"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="{fill_color}"/></a:solidFill>
        {border_xml}
      </p:spPr>
      <p:txBody>
        <a:bodyPr/>
        <a:lstStyle/>
        <a:p><a:endParaRPr lang="fr-FR"/></a:p>
      </p:txBody>
    </p:sp>"""

def make_title_slide(slide_data):
    """Generate a title slide XML."""
    title = slide_data["title"]
    subtitle = slide_data.get("subtitle", "")
    footer = slide_data.get("footer", "")
    
    # Dark blue background bar at top
    top_bar = make_rect(0, 0, 12192000, 1200000, PRIMARY)
    
    # Bottom accent bar
    bottom_bar = make_rect(0, 6400000, 12192000, 458000, "F1F5F9")
    
    # Title text
    title_lines = title.split("\n")
    title_paragraphs = ""
    for line in title_lines:
        title_paragraphs += make_paragraph(line, 3600, bold=True, color=PRIMARY, align="ctr", spc_after=200)
    
    title_box = make_textbox(800000, 1600000, 10500000, 2500000, title_paragraphs)
    
    # Subtitle
    subtitle_lines = subtitle.split("\n") if subtitle else []
    sub_paragraphs = ""
    for line in subtitle_lines:
        sub_paragraphs += make_paragraph(line, 2000, bold=False, color=SECONDARY, align="ctr", spc_after=100)
    
    subtitle_box = make_textbox(1500000, 3800000, 9200000, 1500000, sub_paragraphs)
    
    # Footer
    footer_lines = footer.split("\n") if footer else []
    footer_paragraphs = ""
    for line in footer_lines:
        footer_paragraphs += make_paragraph(line, 1200, bold=False, color=SECONDARY, align="ctr", spc_after=50)
    
    footer_box = make_textbox(1500000, 5600000, 9200000, 1000000, footer_paragraphs)
    
    # Small decorative line under title
    accent_line = make_rect(5000000, 3500000, 2200000, 40000, ACCENT)
    
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      {top_bar}
      {bottom_bar}
      {title_box}
      {accent_line}
      {subtitle_box}
      {footer_box}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""

def make_section_slide(slide_data):
    """Generate a section divider slide."""
    title = slide_data["title"]
    subtitle = slide_data.get("subtitle", "")
    
    # Full blue background
    bg_rect = make_rect(0, 0, 12192000, 6858000, PRIMARY)
    
    # Decorative accent bar
    accent_bar = make_rect(5400000, 3000000, 1400000, 40000, ACCENT)
    
    # Title
    title_lines = title.split("\n")
    title_paragraphs = ""
    for line in title_lines:
        title_paragraphs += make_paragraph(line, 4000, bold=True, color=WHITE, align="ctr", spc_after=200)
    
    title_box = make_textbox(1000000, 1500000, 10200000, 2500000, title_paragraphs)
    
    # Subtitle
    sub_paragraphs = make_paragraph(subtitle, 2200, bold=False, color="BDD5F5", align="ctr")
    subtitle_box = make_textbox(2000000, 3500000, 8200000, 1200000, sub_paragraphs)
    
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      {bg_rect}
      {title_box}
      {accent_bar}
      {subtitle_box}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""

def make_content_slide(slide_data, slide_num):
    """Generate a content slide with title and bullet points."""
    title = slide_data["title"]
    bullets = slide_data.get("bullets", [])
    
    # Title bar background
    title_bg = make_rect(0, 0, 12192000, 900000, "F1F5F9")
    
    # Left accent bar
    left_bar = make_rect(0, 0, 60000, 900000, PRIMARY)
    
    # Title text
    title_para = make_paragraph(title, 2400, bold=True, color=PRIMARY, align="l")
    title_box = make_textbox(500000, 150000, 11000000, 700000, title_para)
    
    # Slide number
    num_para = make_paragraph(str(slide_num), 1000, bold=False, color=SECONDARY, align="r")
    num_box = make_textbox(11200000, 6400000, 700000, 400000, num_para)
    
    # Bottom line
    bottom_line = make_rect(500000, 6550000, 11200000, 15000, "E5E7EB")
    
    # Content bullets
    bullet_paragraphs = ""
    for bullet in bullets:
        if bullet == "":
            bullet_paragraphs += '<a:p><a:pPr><a:spcAft><a:spcPts val="200"/></a:spcAft></a:pPr><a:endParaRPr lang="fr-FR" sz="800"/></a:p>'
        elif bullet.startswith("  "):
            # Sub-bullet (indented)
            clean = bullet.strip()
            bullet_paragraphs += f"""<a:p>
              <a:pPr marL="457200" indent="-228600">
                <a:spcAft><a:spcPts val="300"/></a:spcAft>
                <a:buFont typeface="Arial"/>
                <a:buChar char="&#8226;"/>
              </a:pPr>
              {make_text_run(clean, 1500, False, SECONDARY)}
            </a:p>"""
        else:
            # Main bullet
            bullet_paragraphs += f"""<a:p>
              <a:pPr marL="228600" indent="-228600">
                <a:spcAft><a:spcPts val="400"/></a:spcAft>
                <a:buFont typeface="Arial"/>
                <a:buChar char="&#8226;"/>
              </a:pPr>
              {make_text_run(bullet, 1700, False, DARK)}
            </a:p>"""
    
    content_box = make_textbox(500000, 1050000, 11200000, 5200000, bullet_paragraphs)
    
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      {title_bg}
      {left_bar}
      {title_box}
      {content_box}
      {bottom_line}
      {num_box}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""

# ============================================================================
# BUILD THE PPTX
# ============================================================================

def build_pptx():
    num_slides = len(slides_content)
    
    # Generate slide overrides for [Content_Types].xml
    slide_overrides = "\n  ".join(
        f'<Override PartName="/ppt/slides/slide{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(num_slides)
    )
    
    with zipfile.ZipFile(OUTPUT_FILE, 'w', zipfile.ZIP_DEFLATED) as zf:
        # Core structure
        zf.writestr("[Content_Types].xml", CONTENT_TYPES.format(slide_overrides=slide_overrides))
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("ppt/presentation.xml", make_presentation_xml(num_slides))
        zf.writestr("ppt/_rels/presentation.xml.rels", make_presentation_rels(num_slides))
        zf.writestr("ppt/presProps.xml", PRES_PROPS)
        zf.writestr("ppt/theme/theme1.xml", THEME)
        
        # Slide master and layouts
        zf.writestr("ppt/slideMasters/slideMaster1.xml", SLIDE_MASTER)
        zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", SLIDE_MASTER_RELS)
        zf.writestr("ppt/slideLayouts/slideLayout1.xml", SLIDE_LAYOUT_1)
        zf.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", SLIDE_LAYOUT_RELS)
        zf.writestr("ppt/slideLayouts/slideLayout2.xml", SLIDE_LAYOUT_2)
        zf.writestr("ppt/slideLayouts/_rels/slideLayout2.xml.rels", SLIDE_LAYOUT_RELS)
        
        # Generate each slide
        for i, slide_data in enumerate(slides_content):
            slide_type = slide_data["type"]
            
            if slide_type == "title":
                slide_xml = make_title_slide(slide_data)
                layout_id = 1
            elif slide_type == "section":
                slide_xml = make_section_slide(slide_data)
                layout_id = 1
            else:  # content
                slide_xml = make_content_slide(slide_data, i + 1)
                layout_id = 2
            
            zf.writestr(f"ppt/slides/slide{i+1}.xml", slide_xml)
            zf.writestr(f"ppt/slides/_rels/slide{i+1}.xml.rels", 
                       SLIDE_RELS.format(layout_id=layout_id))
    
    print(f"Presentation created: {OUTPUT_FILE}")
    print(f"Total slides: {num_slides}")

if __name__ == "__main__":
    build_pptx()
