import templateDocxSrc from './Fiche vérification de prestation Nettoyage Saclay V2 (1).docx?url'

export const LOGO_SRC = '/cea-logo-original.png'
export const SCALIAN_LOGO_SRC = '/scalian-logo.jpg'
export const TEMPLATE_DOCX_SRC = templateDocxSrc
export const HEADER_TITLE = 'Fiche de vérification de prestation'
export const HEADER_SUBTITLE = 'Contrat Nettoyage des locaux du centre de Paris-Saclay site de Saclay et ses annexes'
export const DRAFT_STORAGE_KEY = 'cleaning-inspection-saclay-draft-v1'
export const ARCHIVE_STORAGE_KEY = 'cleaning-inspection-saclay-archive-v1'
export const ARCHIVE_DB_NAME = 'cleaning-inspection-saclay-db'
export const ARCHIVE_FILE_STORE = 'archive-files'
export const DOCX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
export const PHOTO_MAX_WIDTH = 210
export const PHOTO_MAX_HEIGHT = 280
export const EMUS_PER_PIXEL = 9525
export const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
export const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml'
export const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
export const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
export const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
export const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
export const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
export const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
export const SUPPORTED_PHOTO_TYPES = {
  'image/jpeg': 'jpeg',
  'image/jpg': 'jpg',
  'image/png': 'png',
}

export const GENERAL_FIELDS = [
  ['batiments', 'Bâtiment(s)'],
  ['dateHeure', 'Date et heure du constat'],
  ['localNumero', 'Numéro du local de nettoyage'],
  ['installation', 'Installation'],
  ['redacteur', 'Rédacteur'],
  ['intervenantPresent', 'Intervenant(s) ATALIAN présent (oui/non)'],
  ['visa', 'Visa'],
]

export const SECTION_1 = [
  ['1.1', 'Accueil sécurité installation à jour ?', ['OUI', 'NON', 'Sans objet']],
  ['1.2', 'L’intervenant dispose-t-il d’une habilitation électrique à jour ?', ['OUI', 'NON', 'Sans objet']],
  ['1.3', 'Autorisations spécifiques (PIRL, escabeau, etc.) ? si concerné', ['OUI', 'NON', 'Sans objet']],
]

export const SECTION_2 = [
  ['2.1', 'Chaussures de sécurité (classique fermé)', ['OUI', 'NON', 'Sans objet']],
  ['2.2', 'Gants de protection adaptés', ['OUI', 'NON', 'Sans objet']],
  ['2.3', 'Protection auditive (si besoin)', ['OUI', 'NON', 'Sans objet']],
  ['2.4', 'Lunettes de sécurité', ['OUI', 'NON', 'Sans objet']],
  ['2.5', 'Tenue de travail logotée', ['OUI', 'NON', 'Sans objet']],
  ['2.6', 'EPI en bon état', ['OUI', 'NON', 'Sans objet']],
  ['2.7', 'Port effectif des EPI observé', ['OUI', 'NON', 'Sans objet']],
]

export const SECTION_3 = [
  ['3.1', 'Détartrant sanitaire (Rouge)'],
  ['3.2', 'Ultracid'],
  ['3.3', 'Nettoyant sol (INOV’R Floral Vert)'],
  ['3.4', 'Nettoyant surfaces (INOV’R Surface Bleu)'],
  ['3.5', 'Identification des produits sur bidon (étiquetage)'],
  ['3.6', 'Vaporisateurs bleu / rouge'],
  ['3.7', 'Local ménage : Stockage propre et organisé'],
  ['3.8', 'Disponibilité FDS des produits'],
]

export const SECTION_4 = [
  ['4.1', 'Chariot de lavage bi-bac', true],
  ['4.2', 'Chariot de lavage mono-bac', true],
  ['4.3', 'Sacs poubelles', false],
  ['4.4', 'Aspirateur', true],
  ['4.5', 'Autolaveuse', true],
  ['4.6', 'Lavettes microfibres rouge', false],
  ['4.7', 'Lavettes microfibres bleues', false],
  ['4.8', 'Lavettes microfibres vertes', false],
  ['4.9', 'Lavettes microfibres jaunes', false],
  ['4.10', 'Balai plat + bandeau de sol', true],
  ['4.11', 'Balai classique', true],
  ['4.12', 'Signalétique sol glissant', true],
  ['4.13', 'Stocks suffisants', false],
  ['4.14', 'Disponibilité des modes opératoires', false],
]

export const QUALITY_OPTIONS = [
  ['Satisfaisant', '▲', '#00b050'],
  ['Acceptable', '►', '#ffc000'],
  ['Remise en état à prévoir', '▼', '#ff0000'],
]

export const IMPROVEMENTS = [
  'Sanitaires',
  'Circulations',
  'Coin cafétéria',
  'Autre (précisez en commentaires)',
]

export const TOTAL_BUILDINGS_TARGET = 170
