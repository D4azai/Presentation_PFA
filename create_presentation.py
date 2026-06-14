#!/usr/bin/env python3
"""
Create a professional PowerPoint presentation for PFA internship defense.
Uses raw Office Open XML with embedded images (no external dependencies).
"""

import zipfile
import os
import struct

# ============================================================================
# CONFIGURATION
# ============================================================================

OUTPUT_FILE = "/projects/sandbox/Presentation_PFA/Presentation_Stage_PFA.pptx"
IMAGES_DIR = "/projects/sandbox/Presentation_PFA/PFAimages"

# EMU (English Metric Units): 1 inch = 914400 EMU
# Widescreen 16:9 dimensions
SLIDE_CX = 12192000  # width
SLIDE_CY = 6858000   # height

# Colors (RRGGBB)
PRIMARY = "1E40AF"
SECONDARY = "6B7280"
ACCENT = "059669"
WHITE = "FFFFFF"
DARK = "1F2937"
LIGHT_BG = "F1F5F9"

# ============================================================================
# IMAGE UTILITIES
# ============================================================================

def get_png_dimensions(filepath):
    """Read PNG width/height from header."""
    with open(filepath, 'rb') as f:
        f.read(8)   # PNG signature
        f.read(4)   # chunk length
        f.read(4)   # IHDR
        w = struct.unpack('>I', f.read(4))[0]
        h = struct.unpack('>I', f.read(4))[0]
    return w, h

def fit_image(img_w, img_h, box_cx, box_cy):
    """Scale image to fit within box preserving aspect ratio. Returns (cx, cy) in EMU."""
    ratio_w = box_cx / img_w
    ratio_h = box_cy / img_h
    ratio = min(ratio_w, ratio_h)
    return int(img_w * ratio), int(img_h * ratio)

# ============================================================================
# SLIDE DEFINITIONS
# ============================================================================

# Image mapping (filename -> description for reference)
IMG = {
    "arch_tracka": "tracka-architecture.png",
    "refresh_seq": "tracka-refresh-sequence.png",
    "tracka_logs": "trackA.png",
    "railway": "Screenshot from 2026-06-05 18-56-12.png",
    # Analytics screenshots (05-30)
    "analytics_exec": "Screenshot from 2026-05-30 13-18-27.png",
    "analytics_orders_funnel": "Screenshot from 2026-05-30 13-19-26.png",
    "analytics_orders_city": "Screenshot from 2026-05-30 13-19-42.png",
    "analytics_delivery1": "Screenshot from 2026-05-30 13-19-51.png",
    "analytics_delivery2": "Screenshot from 2026-05-30 13-19-54.png",
    "analytics_products": "Screenshot from 2026-05-30 13-20-12.png",
    "analytics_variants": "Screenshot from 2026-05-30 13-20-19.png",
    "analytics_sellers": "Screenshot from 2026-05-30 13-20-28.png",
    "analytics_customers": "Screenshot from 2026-05-30 13-21-59.png",
    "analytics_warehouses": "Screenshot from 2026-05-30 13-22-22.png",
    "admin_datatable": "Screenshot from 2026-05-30 13-22-30.png",
    "admin_detail": "Screenshot from 2026-05-30 13-22-38.png",
    "admin_filters": "Screenshot from 2026-05-30 13-23-04.png",
    "admin_stock": "Screenshot from 2026-05-30 13-23-11.png",
    "admin_movements": "Screenshot from 2026-05-30 13-23-19.png",
    # Stock screenshots (06-05)
    "stock_scanner1": "Screenshot from 2026-06-05 19-24-31.png",
    "stock_scanner2": "Screenshot from 2026-06-05 19-24-44.png",
    "stock_manual1": "Screenshot from 2026-06-05 19-25-08.png",
    "stock_manual2": "Screenshot from 2026-06-05 19-27-18.png",
    "stock_list": "Screenshot from 2026-06-05 19-34-28.png",
    "stock_detail": "Screenshot from 2026-06-05 19-34-35.png",
    "stock_transfer": "Screenshot from 2026-06-05 19-34-47.png",
    "stock_extra": "Screenshot from 2026-06-05 19-35-38.png",
}

# Slides definition: each is a dict with type and content
# Types: "title", "section", "content", "image", "split" (text left + image right)
slides_content = [
    # ===== INTRO =====
    {
        "type": "title",
        "title": "IA Pratique pour une Plateforme\nE-Commerce d'Affiliation",
        "subtitle": "Recommandation de Produits, Intelligence d'Approvisionnement\net Analytics Metier",
        "footer": "CHELLAK Aymane | Encadrants: MOUSSAID Ahmed / ECHCHAHID Khalid\nFST Mohammedia | Maroc Affiliate Network | 2025-2026"
    },
    {
        "type": "content",
        "title": "Plan de la Presentation",
        "bullets": [
            "1. Contexte General et Problematique",
            "2. Presentation de l'Entreprise et Existant",
            "3. Methodologie et Architecture Technique",
            "4. Module Analytics Dashboard (demos)",
            "5. Systeme de Recommandation TrackA (architecture + resultats)",
            "6. Module Enrichissement de Stock (demos)",
            "7. Difficultes Rencontrees et Solutions",
            "8. Bilan et Perspectives",
        ]
    },
    # ===== CONTEXTE =====
    {
        "type": "content",
        "title": "Contexte et Problematique",
        "bullets": [
            "E-commerce en croissance exponentielle au Maroc",
            "Modele d'affiliation : vendeurs independants, infrastructure centralisee",
            "Plateforme multi-roles : Admin, Vendeur, Entrepot, Livreur",
            "",
            "3 problematiques identifiees :",
            "  -> Pas de vision claire des performances (Analytics)",
            "  -> Difficulte de decouverte produit (Recommandation)",
            "  -> Approvisionnement manuel et lent (Stock)",
        ]
    },
    {
        "type": "content",
        "title": "Maroc Affiliate Network",
        "bullets": [
            "Plateforme e-commerce par affiliation au Maroc",
            "",
            "Chaine de valeur integree :",
            "  1. Admin acquiert produits aupres de fournisseurs",
            "  2. Stockage dans entrepots centraux et regionaux",
            "  3. Vendeurs affilies selectionnent et vendent",
            "  4. Plateforme gere la logistique de livraison",
            "  5. Vendeur percoit un benefice sur chaque vente",
            "",
            "Existant : Auth RBAC, Catalogue, Commandes, Livraison, POS",
        ]
    },
    # ===== METHODOLOGIE + ARCHI =====
    {
        "type": "content",
        "title": "Methodologie et Outils",
        "bullets": [
            "Approche Agile iterative (Scrum adapte)",
            "  - Sprints 2 semaines, daily standups, sprint review",
            "  - User stories + criteres d'acceptation",
            "",
            "Workflow de developpement :",
            "  branche feature -> PR -> code review -> preview Vercel -> merge",
            "",
            "Outils : Git/GitHub, Vercel, Railway, Discord",
            "Temps moyen PR -> merge : 1-2 jours",
        ]
    },
    {
        "type": "content",
        "title": "Architecture Globale",
        "bullets": [
            "Architecture hybride : monolithe modulaire + microservices",
            "",
            "Frontend + Backend : Next.js 15 App Router (Vercel)",
            "Recommandation : FastAPI Python (Railway)",
            "Base de donnees : PostgreSQL 16 (Neon)",
            "Cache : Redis 7.4 (Railway)",
            "Jobs async : BullMQ Workers",
            "Monitoring : Prometheus + Grafana",
            "",
            "70+ modeles Prisma, 34 modules admin, 3 langues (fr/en/ar RTL)",
        ]
    },
    {
        "type": "content",
        "title": "Stack Technologique",
        "bullets": [
            "Frontend : Next.js 15, TypeScript strict, Tailwind CSS 4",
            "UI : Radix UI + shadcn/ui, Recharts, TanStack Table",
            "Backend : Server Actions, Prisma 7, PostgreSQL 16",
            "Auth : NextAuth v5, RBAC granulaire par permission",
            "Recommandation : FastAPI 0.104, SQLAlchemy 2, Redis",
            "Monitoring : Prometheus Client, APScheduler",
            "Deploy : Vercel (CI/CD) + Railway (microservices)",
            "Conteneurisation : Docker (image unique multi-process)",
        ]
    },
    # ===== MODULE ANALYTICS =====
    {
        "type": "section",
        "title": "Module 1\nAnalytics Dashboard",
        "subtitle": "8 tableaux de bord | 50+ KPIs | Pilotage operationnel"
    },
    {
        "type": "content",
        "title": "Analytics : Architecture et Conception",
        "bullets": [
            "8 sous-tableaux : Executif, Commandes, Livraison,",
            "  Produits, Variantes, Vendeurs, Clients, Entrepots",
            "",
            "Architecture en couches (pattern repetable) :",
            "  1. React Server Components (Suspense + streaming)",
            "  2. queries.ts (unstable_cache + requirePermission)",
            "  3. Prisma $queryRaw -> PostgreSQL (SQL analytique)",
            "  4. Cache TTL 300s + revalidateTag on-demand",
            "",
            "Parallelisation Promise.all() : 11 requetes en parallele",
        ]
    },
    # Analytics demo slides with images
    {
        "type": "image",
        "title": "Vue Executive - KPIs Agreges",
        "image": "analytics_exec",
        "caption": "Cartes KPI avec indicateurs de tendance periode-sur-periode"
    },
    {
        "type": "image",
        "title": "Analytics Commandes - Funnel",
        "image": "analytics_orders_funnel",
        "caption": "Funnel de conversion et tendance de volume temporelle"
    },
    {
        "type": "image",
        "title": "Analytics Commandes - Performance par Ville",
        "image": "analytics_orders_city",
        "caption": "Taux de confirmation et retour par ville"
    },
    {
        "type": "image",
        "title": "Analytics Livraison",
        "image": "analytics_delivery1",
        "caption": "Comparaison des transporteurs par taux de succes et delai"
    },
    {
        "type": "image",
        "title": "Analytics Produits - Quadrant",
        "image": "analytics_products",
        "caption": "Quadrant performance vs stock pour prioriser les actions"
    },
    {
        "type": "image",
        "title": "Analytics Variantes - Heatmap",
        "image": "analytics_variants",
        "caption": "Heatmap couleur x taille avec 5 modes d'affichage"
    },
    {
        "type": "image",
        "title": "Analytics Vendeurs",
        "image": "analytics_sellers",
        "caption": "Leaderboard par profit, tendance d'activite, scoring churn"
    },
    {
        "type": "image",
        "title": "Analytics Clients & Entrepots",
        "image": "analytics_customers",
        "caption": "Segmentation LTV, analyse de repetition, demande geographique"
    },
    # ===== MODULE TRACKA =====
    {
        "type": "section",
        "title": "Module 2\nSysteme de Recommandation TrackA",
        "subtitle": "Microservice autonome | Algorithme 5 signaux | < 5ms"
    },
    {
        "type": "image",
        "title": "TrackA : Architecture du Microservice",
        "image": "arch_tracka",
        "caption": "Couches applicatives, services internes, infrastructure de donnees"
    },
    {
        "type": "content",
        "title": "TrackA : Algorithme a 5 Signaux",
        "bullets": [
            "Ensemble pondere + penalite annulations :",
            "",
            "  Popularite (25%) : volume commandes recentes [global]",
            "  Historique (35%) : commandes passees + affinite cat. [vendeur]",
            "  Recence (20%) : courbe U inverse - reapprovisionnement [vendeur]",
            "  Nouveaute (10%) : age produit, decroissance 180j [global]",
            "  Engagement (10%) : likes + commentaires [global]",
            "",
            "Score = sum(wi * Si) - penalite",
            "Diversite categorielle (min 4-5 cat/30) + exploration 5%",
        ]
    },
    {
        "type": "content",
        "title": "TrackA : Adaptation Cold-Start",
        "bullets": [
            "Poids adaptes dynamiquement selon le profil :",
            "",
            "  Nouveau (0 cmd)    : Pop 52.5% | Nouv 37.5% | Eng 10%",
            "  Cold-start (<5)    : Pop 33.4% | Hist 14%  | Nouv 22.6%",
            "  Standard (5-50)    : Pop 25%  | Hist 35%  | Rec 20%",
            "  Experimente (>50)  : Pop 12.5% | Hist 43.8% | Rec 23.8%",
            "",
            "Courbe de recence en U inverse :",
            "  Zone optimale = 14-40 jours (cycle reapprovisionnement B2B)",
            "  Valide empiriquement sur 200+ vendeurs actifs",
        ]
    },
    {
        "type": "image",
        "title": "TrackA : Pipeline de Rafraichissement",
        "image": "refresh_seq",
        "caption": "Diagramme de sequence - Pipeline de rafraichissement des recommandations"
    },
    {
        "type": "content",
        "title": "TrackA : Serving Multi-Couches",
        "bullets": [
            "5 niveaux de fallback (degradation gracieuse) :",
            "",
            "  1. Redis cache hit           -> retour < 5ms",
            "  2. Snapshot frais            -> warm Redis + retourner",
            "  3. Snapshot perime           -> retourner stale + queue refresh",
            "  4. Pas de snapshot           -> queue refresh + populaires",
            "  5. Erreur totale             -> fallback popularite globale",
            "",
            "Evenements : order-placed, product-engaged,",
            "  product-updated, seller-created",
            "  -> Invalidation immediate + refresh asynchrone",
        ]
    },
    {
        "type": "image",
        "title": "TrackA : Deploiement Railway",
        "image": "railway",
        "caption": "Dashboard Railway - Services TrackA en production (API, Worker, Scheduler)"
    },
    {
        "type": "image",
        "title": "TrackA : Reponse Temps Reel",
        "image": "tracka_logs",
        "caption": "Logs Next.js : 100 recommandations en 4.42ms (cache hit), decomposition des scores"
    },
    {
        "type": "content",
        "title": "TrackA : Resultats en Production",
        "bullets": [
            "Environnement : 150+ produits, 80+ vendeurs, 3000+ commandes",
            "",
            "  Temps reponse (cache hit) :     < 5 ms",
            "  Temps reponse (calcul complet) : 150-300 ms",
            "  Debit worker (4 threads) :       ~40 refresh/min",
            "  Taux de cache hit :              ~85%",
            "  Requetes SQL par calcul :        8 (batch, 0 N+1)",
            "",
            "Qualite :",
            "  Couverture catalogue : 92%",
            "  Diversite : 4.2 categories / set de 30 (vs 2.1 sans filtre)",
            "  Personnalisation : 68% pour vendeurs avec 5+ commandes",
        ]
    },
    # ===== MODULE STOCK =====
    {
        "type": "section",
        "title": "Module 3\nEnrichissement de Stock",
        "subtitle": "Dual-mode QR + Manuel | Transactions atomiques | Transfert partiel"
    },
    {
        "type": "content",
        "title": "Stock : Architecture et Pipeline",
        "bullets": [
            "Pipeline complet :",
            "  Fournisseur -> Scan/Selection -> PurchaseBatch",
            "  -> Main Warehouse -> Transfert -> Entrepot destination",
            "",
            "Transaction atomique (Prisma, timeout 60s) :",
            "  1. Verifier permission purchases:create",
            "  2. Creer PurchaseBatch + N lignes Purchase",
            "  3. Creer StockMovement par ligne",
            "  4. Upsert WarehouseStock (increment quantite)",
            "  5. Update barcode + audit BarcodeScans",
            "  6. Invalider caches (6 tags)",
        ]
    },
    {
        "type": "image",
        "title": "Stock : Mode Scanner QR",
        "image": "stock_scanner1",
        "caption": "Champ auto-focus avec indicateur de queue et historique d'activite"
    },
    {
        "type": "image",
        "title": "Stock : Panier d'Achat (Scanner)",
        "image": "stock_scanner2",
        "caption": "Panier rempli avec resume : total articles, valeur MAD, fournisseurs"
    },
    {
        "type": "image",
        "title": "Stock : Mode Manuel",
        "image": "stock_manual1",
        "caption": "Grille de produits avec filtres par categorie et fournisseur"
    },
    {
        "type": "image",
        "title": "Stock : Selection de Variantes",
        "image": "stock_manual2",
        "caption": "Modal de selection - Quantite par variante avec 'Appliquer a toutes'"
    },
    {
        "type": "image",
        "title": "Stock : Liste et Detail des Lots",
        "image": "stock_detail",
        "caption": "Detail d'un lot - Selection multi-lignes avec badges de statut de transfert"
    },
    {
        "type": "image",
        "title": "Stock : Transfert vers Entrepot",
        "image": "stock_transfer",
        "caption": "Dialog de transfert - Choix de l'entrepot destination avec resume des quantites"
    },
    # ===== DIFFICULTES & BILAN =====
    {
        "type": "content",
        "title": "Difficultes et Solutions",
        "bullets": [
            "Analytics :",
            "  - 11 requetes lentes -> Promise.all() (11x -> 1x la plus lente)",
            "  - Staleness 5min -> revalidateTag() dans 12 server actions",
            "",
            "Recommandation :",
            "  - Cold-start -> adaptation dynamique des poids (0% -> 52.5% pop)",
            "  - N+1 (150 req, 2-4s) -> batch GROUP BY (8 req, 150-300ms)",
            "  - Pannes Redis/PG -> 5 fallbacks (toujours une reponse)",
            "",
            "Stock :",
            "  - Scans perdus 15-20% -> file d'attente client (0% perte)",
            "  - Integration pipeline -> adaptateur (Open/Closed, 0 modif)",
        ]
    },
    {
        "type": "content",
        "title": "Bilan et Perspectives",
        "bullets": [
            "Objectifs atteints - tous en production :",
            "  -> Analytics : 8 dashboards, 50+ KPIs, < 1s cache",
            "  -> TrackA : microservice autonome, < 5ms, 85% cache hit",
            "  -> Stock : dual-mode, 0% perte, 2h -> 15min",
            "",
            "Perspectives :",
            "  - ML supervise + A/B testing (recommandations)",
            "  - Vues materialisees PostgreSQL (analytics)",
            "  - Import CSV/Excel + app mobile (stock)",
            "  - Dashboard temps reel WebSockets",
        ]
    },
    {
        "type": "title",
        "title": "Merci pour votre attention",
        "subtitle": "Questions ?",
        "footer": "CHELLAK Aymane | Stage PFA 2025-2026\nMaroc Affiliate Network | FST Mohammedia"
    },
]

# ============================================================================
# OOXML GENERATION
# ============================================================================

def xml_escape(text):
    """Escape XML special characters."""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

def make_text_run(text, font_size, bold=False, color=None):
    """Create a text run XML."""
    rpr = f'<a:rPr lang="fr-FR" dirty="0"'
    if bold:
        rpr += ' b="1"'
    if font_size:
        rpr += f' sz="{font_size}"'
    rpr += '>'
    if color:
        rpr += f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
    rpr += '<a:latin typeface="Calibri"/><a:cs typeface="Calibri"/></a:rPr>'
    return f'<a:r>{rpr}<a:t>{xml_escape(text)}</a:t></a:r>'

def make_paragraph(text, font_size, bold=False, color=None, align="l", spc_after=0):
    """Create a paragraph XML."""
    algn = f' algn="{align}"' if align else ''
    spc = f'<a:spcAft><a:spcPts val="{spc_after}"/></a:spcAft>' if spc_after else ''
    return f'<a:p><a:pPr{algn}>{spc}</a:pPr>{make_text_run(text, font_size, bold, color)}</a:p>'

def make_textbox(x, y, cx, cy, paragraphs_xml, anchor="t"):
    """Create a textbox shape."""
    return f"""<p:sp>
  <p:nvSpPr><p:cNvPr id="0" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" rtlCol="0" anchor="{anchor}"><a:spAutoFit/></a:bodyPr>
    <a:lstStyle/>{paragraphs_xml}
  </p:txBody>
</p:sp>"""

def make_rect(x, y, cx, cy, fill_color):
    """Create a filled rectangle."""
    return f"""<p:sp>
  <p:nvSpPr><p:cNvPr id="0" name="Rect"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{fill_color}"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="fr-FR"/></a:p></p:txBody>
</p:sp>"""

def make_image_shape(rId, x, y, cx, cy):
    """Create a picture shape referencing an image relationship."""
    return f"""<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="0" name="Picture"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="{rId}"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>"""

def wrap_slide(shapes_xml, bg_color="FFFFFF"):
    """Wrap shapes in a complete slide XML."""
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="{bg_color}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      {shapes_xml}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""

# ============================================================================
# SLIDE GENERATORS
# ============================================================================

def gen_title_slide(data):
    """Title/ending slide."""
    shapes = ""
    # Top gradient bar
    shapes += make_rect(0, 0, SLIDE_CX, 1200000, PRIMARY)
    # Bottom bar
    shapes += make_rect(0, 6400000, SLIDE_CX, 458000, LIGHT_BG)
    # Title
    title_lines = data["title"].split("\n")
    tp = "".join(make_paragraph(l, 3600, True, PRIMARY, "ctr", 200) for l in title_lines)
    shapes += make_textbox(600000, 1600000, 11000000, 2200000, tp, "ctr")
    # Accent line
    shapes += make_rect(5000000, 3500000, 2200000, 40000, ACCENT)
    # Subtitle
    if data.get("subtitle"):
        sub_lines = data["subtitle"].split("\n")
        sp = "".join(make_paragraph(l, 2000, False, SECONDARY, "ctr", 100) for l in sub_lines)
        shapes += make_textbox(1500000, 3800000, 9200000, 1200000, sp, "ctr")
    # Footer
    if data.get("footer"):
        fp = "".join(make_paragraph(l, 1200, False, SECONDARY, "ctr", 50) for l in data["footer"].split("\n"))
        shapes += make_textbox(1500000, 5600000, 9200000, 800000, fp, "ctr")
    return wrap_slide(shapes), None  # no image needed

def gen_section_slide(data):
    """Section divider with blue background."""
    shapes = ""
    shapes += make_rect(0, 0, SLIDE_CX, SLIDE_CY, PRIMARY)
    # Title
    title_lines = data["title"].split("\n")
    tp = "".join(make_paragraph(l, 4000, True, WHITE, "ctr", 200) for l in title_lines)
    shapes += make_textbox(1000000, 1800000, 10200000, 2200000, tp, "ctr")
    # Accent bar
    shapes += make_rect(5400000, 3700000, 1400000, 40000, ACCENT)
    # Subtitle
    if data.get("subtitle"):
        sp = make_paragraph(data["subtitle"], 2000, False, "BDD5F5", "ctr")
        shapes += make_textbox(1500000, 4000000, 9200000, 1000000, sp, "ctr")
    return wrap_slide(shapes, PRIMARY), None

def gen_content_slide(data, slide_num):
    """Content slide with title bar and bullets."""
    shapes = ""
    # Title background bar
    shapes += make_rect(0, 0, SLIDE_CX, 860000, LIGHT_BG)
    # Left accent
    shapes += make_rect(0, 0, 55000, 860000, PRIMARY)
    # Title
    tp = make_paragraph(data["title"], 2200, True, PRIMARY, "l")
    shapes += make_textbox(450000, 180000, 11000000, 650000, tp)
    # Slide number
    np = make_paragraph(str(slide_num), 1000, False, SECONDARY, "r")
    shapes += make_textbox(11200000, 6400000, 700000, 350000, np)
    # Bottom line
    shapes += make_rect(450000, 6500000, 11300000, 12000, "E2E8F0")
    # Bullets
    bullets_xml = ""
    for bullet in data.get("bullets", []):
        if bullet == "":
            bullets_xml += '<a:p><a:pPr><a:spcAft><a:spcPts val="200"/></a:spcAft></a:pPr><a:endParaRPr lang="fr-FR" sz="600"/></a:p>'
        elif bullet.startswith("  "):
            clean = bullet.strip()
            bullets_xml += f'<a:p><a:pPr marL="457200" indent="-228600"><a:spcAft><a:spcPts val="250"/></a:spcAft><a:buFont typeface="Arial"/><a:buChar char="&#8226;"/></a:pPr>{make_text_run(clean, 1500, False, SECONDARY)}</a:p>'
        else:
            bullets_xml += f'<a:p><a:pPr marL="228600" indent="-228600"><a:spcAft><a:spcPts val="350"/></a:spcAft><a:buFont typeface="Arial"/><a:buChar char="&#8226;"/></a:pPr>{make_text_run(bullet, 1600, False, DARK)}</a:p>'
    shapes += make_textbox(450000, 1000000, 11300000, 5300000, bullets_xml)
    return wrap_slide(shapes), None

def gen_image_slide(data, slide_num, img_key):
    """Slide with title bar, large image, and caption."""
    img_filename = IMG[img_key]
    img_path = os.path.join(IMAGES_DIR, img_filename)
    img_w, img_h = get_png_dimensions(img_path)
    
    # Image area: below title bar, above caption
    # Title bar: y=0 to 860000
    # Caption: last 500000
    # Image area: y=920000, height=5100000, width with margins
    box_x = 300000
    box_y = 920000
    box_cx = SLIDE_CX - 600000  # 11592000
    box_cy = 5100000
    
    # Fit image
    fit_cx, fit_cy = fit_image(img_w, img_h, box_cx, box_cy)
    # Center image in the box
    img_x = box_x + (box_cx - fit_cx) // 2
    img_y = box_y + (box_cy - fit_cy) // 2
    
    shapes = ""
    # Title bar
    shapes += make_rect(0, 0, SLIDE_CX, 860000, LIGHT_BG)
    shapes += make_rect(0, 0, 55000, 860000, PRIMARY)
    tp = make_paragraph(data["title"], 2000, True, PRIMARY, "l")
    shapes += make_textbox(450000, 200000, 11000000, 600000, tp)
    # Slide number
    np = make_paragraph(str(slide_num), 1000, False, SECONDARY, "r")
    shapes += make_textbox(11200000, 6400000, 700000, 350000, np)
    # Image (will reference rId2)
    shapes += make_image_shape("rId2", img_x, img_y, fit_cx, fit_cy)
    # Caption
    if data.get("caption"):
        cp = make_paragraph(data["caption"], 1100, False, SECONDARY, "ctr")
        shapes += make_textbox(1000000, 6200000, 10200000, 400000, cp)
    # Thin border around image area
    shapes += make_rect(img_x - 10000, img_y - 10000, fit_cx + 20000, 12000, "E2E8F0")  # top border
    shapes += make_rect(img_x - 10000, img_y + fit_cy, fit_cx + 20000, 12000, "E2E8F0")  # bottom border
    
    return wrap_slide(shapes), img_filename

# ============================================================================
# PACKAGE STRUCTURE FILES
# ============================================================================

def make_content_types(num_slides, image_slides):
    """Generate [Content_Types].xml with proper image type."""
    overrides = "\n".join(
        f'  <Override PartName="/ppt/slides/slide{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(num_slides)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
{overrides}
</Types>"""

RELS_ROOT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"""

def make_presentation_xml(num_slides):
    sld_list = "\n    ".join(f'<p:sldId id="{256+i}" r:id="rId{i+4}"/>' for i in range(num_slides))
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>
    {sld_list}
  </p:sldIdLst>
  <p:sldSz cx="{SLIDE_CX}" cy="{SLIDE_CY}"/>
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
    <a:clrScheme name="PFA">
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
    <a:fontScheme name="PFA">
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>"""

SLIDE_MASTER = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
    <p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>
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

SLIDE_LAYOUT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="{ltype}" preserve="1">
  <p:cSld name="{name}"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"""

SLIDE_LAYOUT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"""

def make_slide_rels(layout_id, image_filename=None):
    """Generate slide .rels file. If image_filename is provided, add image relationship."""
    rels = [f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout{layout_id}.xml"/>']
    if image_filename:
        rels.append(f'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{image_filename}"/>')
    rels_str = "\n  ".join(rels)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {rels_str}
</Relationships>"""

# ============================================================================
# BUILD PPTX
# ============================================================================

def build_pptx():
    num_slides = len(slides_content)
    
    # Track which images are used (to add to media/)
    used_images = {}  # media_filename -> source_path
    # Track image slides for their rels
    slide_images = {}  # slide_index -> media_filename
    
    # First pass: determine which slides need images
    for i, slide_data in enumerate(slides_content):
        if slide_data["type"] == "image":
            img_key = slide_data["image"]
            src_filename = IMG[img_key]
            # Use a sanitized name for media
            media_name = f"image{i+1}.png"
            used_images[media_name] = os.path.join(IMAGES_DIR, src_filename)
            slide_images[i] = media_name
    
    with zipfile.ZipFile(OUTPUT_FILE, 'w', zipfile.ZIP_DEFLATED) as zf:
        # Package structure
        zf.writestr("[Content_Types].xml", make_content_types(num_slides, slide_images))
        zf.writestr("_rels/.rels", RELS_ROOT)
        zf.writestr("ppt/presentation.xml", make_presentation_xml(num_slides))
        zf.writestr("ppt/_rels/presentation.xml.rels", make_presentation_rels(num_slides))
        zf.writestr("ppt/presProps.xml", PRES_PROPS)
        zf.writestr("ppt/theme/theme1.xml", THEME)
        zf.writestr("ppt/slideMasters/slideMaster1.xml", SLIDE_MASTER)
        zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", SLIDE_MASTER_RELS)
        zf.writestr("ppt/slideLayouts/slideLayout1.xml", SLIDE_LAYOUT.format(ltype="title", name="Title"))
        zf.writestr("ppt/slideLayouts/slideLayout2.xml", SLIDE_LAYOUT.format(ltype="obj", name="Content"))
        zf.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", SLIDE_LAYOUT_RELS)
        zf.writestr("ppt/slideLayouts/_rels/slideLayout2.xml.rels", SLIDE_LAYOUT_RELS)
        
        # Add media images
        for media_name, src_path in used_images.items():
            zf.write(src_path, f"ppt/media/{media_name}")
        
        # Generate slides
        for i, slide_data in enumerate(slides_content):
            stype = slide_data["type"]
            slide_num = i + 1
            
            if stype == "title":
                slide_xml, _ = gen_title_slide(slide_data)
                layout_id = 1
                img_media = None
            elif stype == "section":
                slide_xml, _ = gen_section_slide(slide_data)
                layout_id = 1
                img_media = None
            elif stype == "image":
                slide_xml, _ = gen_image_slide(slide_data, slide_num, slide_data["image"])
                layout_id = 2
                img_media = slide_images[i]
            else:  # content
                slide_xml, _ = gen_content_slide(slide_data, slide_num)
                layout_id = 2
                img_media = None
            
            zf.writestr(f"ppt/slides/slide{slide_num}.xml", slide_xml)
            zf.writestr(f"ppt/slides/_rels/slide{slide_num}.xml.rels", 
                       make_slide_rels(layout_id, img_media))
    
    # Print summary
    file_size = os.path.getsize(OUTPUT_FILE) / (1024*1024)
    print(f"{'='*60}")
    print(f"  Presentation creee : {OUTPUT_FILE}")
    print(f"  Nombre de slides   : {num_slides}")
    print(f"  Images integrees   : {len(used_images)}")
    print(f"  Taille fichier     : {file_size:.1f} MB")
    print(f"{'='*60}")
    print(f"\n  Slides avec images :")
    for idx, media in sorted(slide_images.items()):
        print(f"    Slide {idx+1:2d} : {slides_content[idx]['title']}")

if __name__ == "__main__":
    build_pptx()
