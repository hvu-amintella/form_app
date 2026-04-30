import { useEffect, useMemo, useRef } from 'react'
import { useState } from 'react'
import JSZip from 'jszip'
import { clearArchiveFiles, getArchiveFile, saveArchiveFile } from './archiveStorage'
import { buildGenericDocxBlob, createInitialGenericValues, parseGenericDocxTemplate } from './genericDocx'
import { buildGenericPdfBlob, createInitialGenericPdfValues, parseGenericPdfTemplate } from './genericPdf'
import {
  A_NS,
  CONTENT_TYPES_NS,
  DOCX_MIME_TYPE,
  DRAFT_STORAGE_KEY,
  EMUS_PER_PIXEL,
  GENERAL_FIELDS,
  HEADER_SUBTITLE,
  HEADER_TITLE,
  IMPROVEMENTS,
  LOGO_SRC,
  PIC_NS,
  PHOTO_MAX_HEIGHT,
  PHOTO_MAX_WIDTH,
  QUALITY_OPTIONS,
  RELS_NS,
  R_NS,
  SCALIAN_LOGO_SRC,
  SECTION_1,
  SECTION_2,
  SECTION_3,
  SECTION_4,
  SUPPORTED_PHOTO_TYPES,
  TEMPLATE_DOCX_SRC,
  TOTAL_BUILDINGS_TARGET,
  W14_NS,
  WP_NS,
  W_NS,
  ARCHIVE_STORAGE_KEY,
} from './constants'
import { buildStatisticsSnapshot } from './statistics'
import './App.css'

function createInitialData() {
  const data = {
    batiments: '',
    dateHeure: '',
    localNumero: '',
    installation: '',
    redacteur: '',
    intervenantPresent: '',
    visa: '',
    comments1: '',
    comments2: '',
    comments3: '',
    comments4: '',
    quality: '',
    qualityComments: '',
    improve: [],
    improveComments: '',
    otherRemarks: '',
  }

  for (const [key] of SECTION_1) data[key] = ''
  for (const [key] of SECTION_2) data[key] = ''
  for (const [key] of SECTION_3) data[key] = { status: '', name: '' }
  for (const [key, , withState] of SECTION_4) data[key] = withState ? { status: '', state: '' } : { status: '' }
  return data
}

function createStoredDraft() {
  if (typeof window === 'undefined') return createInitialData()

  const fallback = createInitialData()
  const raw = window.localStorage.getItem(DRAFT_STORAGE_KEY)

  if (!raw) return fallback

  try {
    const parsed = JSON.parse(raw)
    return { ...fallback, ...parsed }
  } catch {
    return fallback
  }
}

function createStoredArchive() {
  if (typeof window === 'undefined') return []

  const raw = window.localStorage.getItem(ARCHIVE_STORAGE_KEY)
  if (!raw) return []

  try {
    const parsed = JSON.parse(raw)
    return Array.isArray(parsed) ? parsed : []
  } catch {
    return []
  }
}

function completionCount(data) {
  const keys = [
    ...GENERAL_FIELDS.map(([key]) => key),
    ...SECTION_1.map(([key]) => key),
    ...SECTION_2.map(([key]) => key),
    ...SECTION_3.map(([key]) => key),
    ...SECTION_4.map(([key]) => key),
    'quality',
    'improve',
  ]

  const done = keys.filter((key) => {
    const value = data[key]
    if (Array.isArray(value)) return value.length > 0
    if (value && typeof value === 'object') return Object.values(value).some(Boolean)
    return Boolean(value)
  }).length

  return Math.round((done / keys.length) * 100)
}

function displayValue(value) {
  if (Array.isArray(value)) return value.length ? value.join(', ') : ''
  if (value && typeof value === 'object') return Object.values(value).filter(Boolean).join(' ')
  return value || ''
}

function deepClone(value) {
  return JSON.parse(JSON.stringify(value))
}

function checkboxMark(current, expected) {
  return current === expected ? '☒' : '☐'
}

function stateMark(current, expected) {
  return current === expected ? '☒' : '☐'
}

function getWordChildren(node, localName) {
  return Array.from(node?.children || []).filter(
    (child) => child.namespaceURI === W_NS && child.localName === localName,
  )
}

function getWordDescendants(node, localName) {
  return Array.from(node?.getElementsByTagNameNS(W_NS, localName) || [])
}

function normalizeWordText(value) {
  return value.replaceAll('\u00a0', ' ').replace(/\s+/g, ' ').trim()
}

function getNodeText(node) {
  return normalizeWordText(
    getWordDescendants(node, 't')
      .map((textNode) => textNode.textContent || '')
      .join(''),
  )
}

function createWordElement(doc, localName) {
  return doc.createElementNS(W_NS, `w:${localName}`)
}

function clearParagraph(paragraph) {
  const properties = getWordChildren(paragraph, 'pPr')[0] || null
  paragraph.replaceChildren()
  if (properties) paragraph.append(properties)
}

function setParagraphText(paragraph, value) {
  clearParagraph(paragraph)

  const lines = String(value || '').split('\n')
  if (lines.length === 1 && !lines[0]) return

  const doc = paragraph.ownerDocument
  lines.forEach((line, index) => {
    const run = createWordElement(doc, 'r')
    if (index > 0) run.append(createWordElement(doc, 'br'))
    if (line) {
      const textNode = createWordElement(doc, 't')
      if (/^\s|\s$/.test(line)) {
        textNode.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve')
      }
      textNode.textContent = line
      run.append(textNode)
    }
    paragraph.append(run)
  })
}

function setCellText(cell, value) {
  const paragraphs = getWordChildren(cell, 'p')
  const target = paragraphs[0] || createWordElement(cell.ownerDocument, 'p')

  if (!paragraphs.length) cell.append(target)
  setParagraphText(target, value)

  paragraphs.slice(1).forEach((paragraph) => clearParagraph(paragraph))
}

function setCheckboxState(sdt, checked) {
  const checkedNode = sdt.getElementsByTagNameNS(W14_NS, 'checked')[0]
  if (checkedNode) checkedNode.setAttributeNS(W14_NS, 'w14:val', checked ? '1' : '0')

  const textNode = sdt.getElementsByTagNameNS(W_NS, 't')[0]
  if (textNode) textNode.textContent = checked ? '☒' : '☐'
}

function setFirstCheckboxInNode(node, checked) {
  const checkbox = getWordDescendants(node, 'sdt')[0]
  if (checkbox) setCheckboxState(checkbox, checked)
}

function findTableByText(doc, text) {
  return getWordDescendants(doc, 'tbl').find((table) => getNodeText(table).includes(text))
}

function getTableRows(table) {
  return getWordChildren(table, 'tr')
}

function getRowCells(row) {
  return getWordChildren(row, 'tc')
}

function findParagraphByText(doc, text) {
  return getWordDescendants(doc, 'p').find((paragraph) => getNodeText(paragraph).includes(text))
}

function getNextWordParagraph(paragraph) {
  let current = paragraph?.nextElementSibling || null
  while (current) {
    if (current.namespaceURI === W_NS && current.localName === 'p') return current
    current = current.nextElementSibling
  }
  return null
}

function isCheckboxChecked(checkbox) {
  if (!checkbox) return false

  const checkedNode = checkbox.getElementsByTagNameNS(W14_NS, 'checked')[0]
  const checkedValue = checkedNode?.getAttributeNS(W14_NS, 'val')
    || checkedNode?.getAttribute('w14:val')
    || checkedNode?.getAttribute('val')

  if (checkedValue) return checkedValue === '1' || checkedValue === 'true'

  const textNode = checkbox.getElementsByTagNameNS(W_NS, 't')[0]
  return textNode?.textContent === '☒'
}

function getParagraphTextAfterLabel(doc, label, occurrence = 0) {
  const matches = getWordDescendants(doc, 'p').filter((paragraph) => getNodeText(paragraph).includes(label))
  const target = getNextWordParagraph(matches[occurrence])
  return target ? getNodeText(target) : ''
}

function inferChoiceValue(cells, options) {
  if (isCheckboxChecked(getWordDescendants(cells[2], 'sdt')[0])) return 'OUI'
  if (isCheckboxChecked(getWordDescendants(cells[3], 'sdt')[0])) return 'NON'
  if (options.includes('Sans objet') && isCheckboxChecked(getWordDescendants(cells[4], 'sdt')[0])) {
    return 'Sans objet'
  }
  return ''
}

function setParagraphAfterLabel(doc, label, value, occurrence = 0) {
  const matches = getWordDescendants(doc, 'p').filter((paragraph) => getNodeText(paragraph).includes(label))
  const target = getNextWordParagraph(matches[occurrence])
  if (target) setParagraphText(target, value)
}

function getExportFileName(data) {
  const building = String(data.batiments || '')
    .trim()
    .replace(/[\\/:*?"<>|]/g, ' ')
    .replace(/\s+/g, ' ')

  return `${building || 'Document'} - Fiche vérification de prestation Nettoyage Saclay.docx`
}

function getImportedMergedFileName(fileName, data) {
  const baseName = stripFileExtension(fileName || getExportFileName(data))
    .trim()
    .replace(/[\\/:*?"<>|]/g, ' ')
    .replace(/\s+/g, ' ')

  return `${baseName || 'Document'} - avec photos.docx`
}

function getStatisticsExportFileName() {
  const stamp = new Date().toISOString().slice(0, 10)
  return `${stamp} - Statistiques Fiche vérification de prestation Nettoyage Saclay.docx`
}

function getArchiveZipFileName() {
  const stamp = new Date().toISOString().slice(0, 10)
  return `${stamp} - Archives inspections.docx.zip`
}

function getUniqueArchiveEntryName(fileName, usedNames) {
  const safeName = String(fileName || 'archive.docx')
    .trim()
    .replace(/[\\/:*?"<>|]/g, ' ')
    .replace(/\s+/g, ' ')

  const normalizedName = safeName || 'archive.docx'
  const extensionIndex = normalizedName.lastIndexOf('.')
  const hasExtension = extensionIndex > 0
  const baseName = hasExtension ? normalizedName.slice(0, extensionIndex) : normalizedName
  const extension = hasExtension ? normalizedName.slice(extensionIndex) : '.docx'

  let candidate = `${baseName}${extension}`
  let suffix = 2

  while (usedNames.has(candidate)) {
    candidate = `${baseName} (${suffix})${extension}`
    suffix += 1
  }

  usedNames.add(candidate)
  return candidate
}

function createArchiveRecord({ data, photos = [], sourceType, sourceName = '', storedFileName = '' }) {
  return {
    id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    sourceType,
    sourceName,
    storedFileName,
    createdAt: new Date().toISOString(),
    building: data.batiments || '',
    inspectionDate: data.dateHeure || '',
    installation: data.installation || '',
    quality: data.quality || '',
    photoCount: photos.length,
    photos: photos.map((photo) => ({
      name: photo.name || '',
      caption: photo.caption || '',
    })),
    data: deepClone(data),
  }
}

function createGeneratedFormArchiveRecord({ schema, values, storedFileName }) {
  return {
    id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    sourceType: 'generated_form',
    sourceName: schema.fileName || '',
    storedFileName,
    createdAt: new Date().toISOString(),
    formKind: schema.kind || 'form',
    templateId: `${schema.kind || 'form'}:${schema.fileName || 'template'}`,
    schema: deepClone(schema),
    values: deepClone(values),
  }
}

async function parseImportedDocxFile(file) {
  const templateBuffer = await file.arrayBuffer()
  const zip = await JSZip.loadAsync(templateBuffer)
  const documentXml = await zip.file('word/document.xml')?.async('string')

  if (!documentXml) {
    throw new Error('Le fichier Word ne contient pas de document exploitable.')
  }

  const xmlDoc = new DOMParser().parseFromString(documentXml, 'application/xml')
  const parsed = createInitialData()

  const generalTable = findTableByText(xmlDoc, 'Bâtiment(s)')
  if (!generalTable) {
    throw new Error('Le modele Word importe ne correspond pas a la fiche attendue.')
  }

  const generalRows = getTableRows(generalTable)
  parsed.batiments = getNodeText(getRowCells(generalRows[0])[1])
  parsed.dateHeure = getNodeText(getRowCells(generalRows[0])[4])
  parsed.localNumero = getNodeText(getRowCells(generalRows[1])[1])
  parsed.installation = getNodeText(getRowCells(generalRows[2])[1])
  parsed.redacteur = getNodeText(getRowCells(generalRows[2])[4])
  parsed.intervenantPresent = getNodeText(getRowCells(generalRows[3])[1])
  parsed.visa = getNodeText(getRowCells(generalRows[3])[4])

  const sectionOneTable = findTableByText(xmlDoc, 'Formations et Habilitations')
  const sectionOneRows = getTableRows(sectionOneTable)
  SECTION_1.forEach(([key, , options], index) => {
    parsed[key] = inferChoiceValue(getRowCells(sectionOneRows[index + 1]), options)
  })
  parsed.comments1 = getParagraphTextAfterLabel(xmlDoc, 'Commentaires', 0)

  const sectionTwoTable = findTableByText(xmlDoc, 'Equipements des intervenants')
  const sectionTwoRows = getTableRows(sectionTwoTable)
  SECTION_2.forEach(([key, , options], index) => {
    parsed[key] = inferChoiceValue(getRowCells(sectionTwoRows[index + 1]), options)
  })
  parsed.comments2 = getParagraphTextAfterLabel(xmlDoc, 'Commentaires', 1)

  const productsTable = findTableByText(xmlDoc, 'Produits de nettoyage')
  const productRows = getTableRows(productsTable)
  SECTION_3.forEach(([key], index) => {
    const cells = getRowCells(productRows[index + 1])
    parsed[key] = {
      status: inferChoiceValue(cells, ['OUI', 'NON']),
      name: getNodeText(cells[4]),
    }
  })
  parsed.comments3 = getParagraphTextAfterLabel(xmlDoc, 'Commentaires', 2)

  const materialsTable = findTableByText(xmlDoc, 'Matériels / documents')
  const materialRows = getTableRows(materialsTable)
  SECTION_4.forEach(([key, , withState], index) => {
    const cells = getRowCells(materialRows[index + 1])
    const nextValue = {
      status: inferChoiceValue(cells, ['OUI', 'NON']),
    }

    if (withState) {
      const stateCheckboxes = getWordDescendants(cells[4], 'sdt')
      nextValue.state = isCheckboxChecked(stateCheckboxes[0])
        ? 'Bon état'
        : isCheckboxChecked(stateCheckboxes[1])
          ? 'Etat d’usage'
          : isCheckboxChecked(stateCheckboxes[2])
            ? 'Vétuste'
            : ''
    }

    parsed[key] = nextValue
  })
  parsed.comments4 = getParagraphTextAfterLabel(xmlDoc, 'Commentaires', 3)

  QUALITY_OPTIONS.forEach(([label]) => {
    const paragraph = findParagraphByText(xmlDoc, label)
    if (isCheckboxChecked(getWordDescendants(paragraph, 'sdt')[0])) parsed.quality = label
  })

  IMPROVEMENTS.forEach((label) => {
    const paragraph = findParagraphByText(xmlDoc, label)
    if (isCheckboxChecked(getWordDescendants(paragraph, 'sdt')[0])) {
      parsed.improve = [...parsed.improve, label]
    }
  })

  parsed.qualityComments = getParagraphTextAfterLabel(xmlDoc, 'Commentaires', 4)
  parsed.improveComments = getParagraphTextAfterLabel(xmlDoc, 'Commentaires', 5)
  parsed.otherRemarks = getParagraphTextAfterLabel(xmlDoc, 'Autres Remarques')

  return parsed
}

function createPhotoId() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

function stripFileExtension(fileName) {
  return fileName.replace(/\.[^.]+$/, '')
}

function createImportedPhotoEntries(files) {
  return files.map((file, index) => ({
    id: `import-photo-${Date.now()}-${index}-${Math.random().toString(36).slice(2, 8)}`,
    file,
    name: `Photo ${index + 1}`,
    caption: stripFileExtension(file.name),
  }))
}

function revokePhotoUrls(photos) {
  photos.forEach((photo) => {
    if (photo.previewUrl) URL.revokeObjectURL(photo.previewUrl)
  })
}

function scaleToFit(width, height, maxWidth, maxHeight) {
  if (!width || !height) return { width: maxWidth, height: Math.round(maxWidth * 0.75) }

  const ratio = Math.min(maxWidth / width, maxHeight / height, 1)
  return {
    width: Math.max(1, Math.round(width * ratio)),
    height: Math.max(1, Math.round(height * ratio)),
  }
}

function readImageDimensions(file) {
  return new Promise((resolve, reject) => {
    const imageUrl = URL.createObjectURL(file)
    const image = new Image()

    image.onload = () => {
      resolve({
        width: image.naturalWidth || image.width,
        height: image.naturalHeight || image.height,
      })
      URL.revokeObjectURL(imageUrl)
    }

    image.onerror = () => {
      URL.revokeObjectURL(imageUrl)
      reject(new Error(`Impossible de lire l'image ${file.name}`))
    }

    image.src = imageUrl
  })
}

async function preparePhotoForDocx(photo) {
  const [buffer, dimensions] = await Promise.all([
    photo.file.arrayBuffer(),
    readImageDimensions(photo.file),
  ])

  const scaled = scaleToFit(dimensions.width, dimensions.height, PHOTO_MAX_WIDTH, PHOTO_MAX_HEIGHT)

  return {
    ...photo,
    buffer,
    width: scaled.width,
    height: scaled.height,
  }
}

function ensureContentTypeDefault(contentTypesDoc, extension, contentType) {
  const defaults = Array.from(contentTypesDoc.getElementsByTagNameNS(CONTENT_TYPES_NS, 'Default'))
  const exists = defaults.some((node) => node.getAttribute('Extension') === extension)
  if (exists) return

  const typesNode = contentTypesDoc.documentElement
  const defaultNode = contentTypesDoc.createElementNS(CONTENT_TYPES_NS, 'Default')
  defaultNode.setAttribute('Extension', extension)
  defaultNode.setAttribute('ContentType', contentType)
  typesNode.append(defaultNode)
}

function appendRunText(paragraph, text, options = {}) {
  const { bold = false, fontSize = null } = options
  const run = createWordElement(paragraph.ownerDocument, 'r')

  if (bold || fontSize) {
    const runProperties = createWordElement(paragraph.ownerDocument, 'rPr')
    if (bold) runProperties.append(createWordElement(paragraph.ownerDocument, 'b'))
    if (fontSize) {
      const sizeNode = createWordElement(paragraph.ownerDocument, 'sz')
      sizeNode.setAttributeNS(W_NS, 'w:val', String(fontSize))
      runProperties.append(sizeNode)
    }
    run.append(runProperties)
  }

  const textNode = createWordElement(paragraph.ownerDocument, 't')
  if (/^\s|\s$/.test(text)) {
    textNode.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve')
  }
  textNode.textContent = text
  run.append(textNode)
  paragraph.append(run)
}

function createStyledParagraph(doc, text, options = {}) {
  const {
    bold = false,
    fontSize = null,
    centered = false,
    pageBreakBefore = false,
    spacingAfter = null,
  } = options

  const paragraph = createWordElement(doc, 'p')
  const properties = createWordElement(doc, 'pPr')

  if (centered) {
    const alignment = createWordElement(doc, 'jc')
    alignment.setAttributeNS(W_NS, 'w:val', 'center')
    properties.append(alignment)
  }

  if (pageBreakBefore) properties.append(createWordElement(doc, 'pageBreakBefore'))

  if (spacingAfter !== null) {
    const spacing = createWordElement(doc, 'spacing')
    spacing.setAttributeNS(W_NS, 'w:after', String(spacingAfter))
    properties.append(spacing)
  }

  if (properties.children.length) paragraph.append(properties)
  appendRunText(paragraph, text, { bold, fontSize })
  return paragraph
}

function createTableCell(doc, children, options = {}) {
  const { width = null, shaded = false, centered = false } = options
  const cell = createWordElement(doc, 'tc')
  const properties = createWordElement(doc, 'tcPr')

  if (width !== null) {
    const widthNode = createWordElement(doc, 'tcW')
    widthNode.setAttributeNS(W_NS, 'w:w', String(width))
    widthNode.setAttributeNS(W_NS, 'w:type', 'dxa')
    properties.append(widthNode)
  }

  if (shaded) {
    const shading = createWordElement(doc, 'shd')
    shading.setAttributeNS(W_NS, 'w:val', 'clear')
    shading.setAttributeNS(W_NS, 'w:color', 'auto')
    shading.setAttributeNS(W_NS, 'w:fill', 'D9E7F5')
    properties.append(shading)
  }

  if (centered) {
    const verticalAlign = createWordElement(doc, 'vAlign')
    verticalAlign.setAttributeNS(W_NS, 'w:val', 'center')
    properties.append(verticalAlign)
  }

  cell.append(properties)
  children.forEach((child) => cell.append(child))
  return cell
}

function createPhotoParagraph(doc, photo, relationshipId, drawingId) {
  const paragraph = createWordElement(doc, 'p')
  const properties = createWordElement(doc, 'pPr')
  const alignment = createWordElement(doc, 'jc')
  alignment.setAttributeNS(W_NS, 'w:val', 'center')
  properties.append(alignment)
  const spacing = createWordElement(doc, 'spacing')
  spacing.setAttributeNS(W_NS, 'w:after', '160')
  properties.append(spacing)
  paragraph.append(properties)

  const run = createWordElement(doc, 'r')
  const drawing = doc.createElementNS(W_NS, 'w:drawing')
  const inline = doc.createElementNS(WP_NS, 'wp:inline')
  const extent = doc.createElementNS(WP_NS, 'wp:extent')
  extent.setAttribute('cx', String(photo.width * EMUS_PER_PIXEL))
  extent.setAttribute('cy', String(photo.height * EMUS_PER_PIXEL))
  inline.append(extent)

  const effectExtent = doc.createElementNS(WP_NS, 'wp:effectExtent')
  effectExtent.setAttribute('l', '0')
  effectExtent.setAttribute('t', '0')
  effectExtent.setAttribute('r', '0')
  effectExtent.setAttribute('b', '0')
  inline.append(effectExtent)

  const docPr = doc.createElementNS(WP_NS, 'wp:docPr')
  docPr.setAttribute('id', String(drawingId))
  docPr.setAttribute('name', `Photo ${drawingId}`)
  inline.append(docPr)

  const cNvGraphicFramePr = doc.createElementNS(WP_NS, 'wp:cNvGraphicFramePr')
  const graphicFrameLocks = doc.createElementNS(A_NS, 'a:graphicFrameLocks')
  graphicFrameLocks.setAttribute('noChangeAspect', '1')
  cNvGraphicFramePr.append(graphicFrameLocks)
  inline.append(cNvGraphicFramePr)

  const graphic = doc.createElementNS(A_NS, 'a:graphic')
  const graphicData = doc.createElementNS(A_NS, 'a:graphicData')
  graphicData.setAttribute('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture')

  const picture = doc.createElementNS(PIC_NS, 'pic:pic')
  const nvPicPr = doc.createElementNS(PIC_NS, 'pic:nvPicPr')
  const cNvPr = doc.createElementNS(PIC_NS, 'pic:cNvPr')
  cNvPr.setAttribute('id', '0')
  cNvPr.setAttribute('name', photo.file.name)
  nvPicPr.append(cNvPr)
  nvPicPr.append(doc.createElementNS(PIC_NS, 'pic:cNvPicPr'))
  picture.append(nvPicPr)

  const blipFill = doc.createElementNS(PIC_NS, 'pic:blipFill')
  const blip = doc.createElementNS(A_NS, 'a:blip')
  blip.setAttributeNS(R_NS, 'r:embed', relationshipId)
  blipFill.append(blip)
  const stretch = doc.createElementNS(A_NS, 'a:stretch')
  stretch.append(doc.createElementNS(A_NS, 'a:fillRect'))
  blipFill.append(stretch)
  picture.append(blipFill)

  const spPr = doc.createElementNS(PIC_NS, 'pic:spPr')
  const transform2d = doc.createElementNS(A_NS, 'a:xfrm')
  const offset = doc.createElementNS(A_NS, 'a:off')
  offset.setAttribute('x', '0')
  offset.setAttribute('y', '0')
  const extents = doc.createElementNS(A_NS, 'a:ext')
  extents.setAttribute('cx', String(photo.width * EMUS_PER_PIXEL))
  extents.setAttribute('cy', String(photo.height * EMUS_PER_PIXEL))
  transform2d.append(offset, extents)
  spPr.append(transform2d)
  const presetGeometry = doc.createElementNS(A_NS, 'a:prstGeom')
  presetGeometry.setAttribute('prst', 'rect')
  presetGeometry.append(doc.createElementNS(A_NS, 'a:avLst'))
  spPr.append(presetGeometry)
  picture.append(spPr)

  graphicData.append(picture)
  graphic.append(graphicData)
  inline.append(graphic)
  drawing.append(inline)
  run.append(drawing)
  paragraph.append(run)

  return paragraph
}

function createPhotoTable(doc, entries) {
  const table = createWordElement(doc, 'tbl')
  const tableProperties = createWordElement(doc, 'tblPr')
  const tableWidth = createWordElement(doc, 'tblW')
  tableWidth.setAttributeNS(W_NS, 'w:w', '9600')
  tableWidth.setAttributeNS(W_NS, 'w:type', 'dxa')
  tableProperties.append(tableWidth)

  const justification = createWordElement(doc, 'jc')
  justification.setAttributeNS(W_NS, 'w:val', 'center')
  tableProperties.append(justification)

  const borders = createWordElement(doc, 'tblBorders')
  ;['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].forEach((side) => {
    const border = createWordElement(doc, side)
    border.setAttributeNS(W_NS, 'w:val', 'single')
    border.setAttributeNS(W_NS, 'w:sz', '8')
    border.setAttributeNS(W_NS, 'w:space', '0')
    border.setAttributeNS(W_NS, 'w:color', 'A9BDD2')
    borders.append(border)
  })
  tableProperties.append(borders)

  const cellMargins = createWordElement(doc, 'tblCellMar')
  ;[
    ['top', '120'],
    ['left', '120'],
    ['bottom', '120'],
    ['right', '120'],
  ].forEach(([side, value]) => {
    const margin = createWordElement(doc, side)
    margin.setAttributeNS(W_NS, 'w:w', value)
    margin.setAttributeNS(W_NS, 'w:type', 'dxa')
    cellMargins.append(margin)
  })
  tableProperties.append(cellMargins)
  table.append(tableProperties)

  const grid = createWordElement(doc, 'tblGrid')
  ;['2200', '3600', '3800'].forEach((width) => {
    const gridCol = createWordElement(doc, 'gridCol')
    gridCol.setAttributeNS(W_NS, 'w:w', width)
    grid.append(gridCol)
  })
  table.append(grid)

  const headerRow = createWordElement(doc, 'tr')
  headerRow.append(
    createTableCell(doc, [createStyledParagraph(doc, 'Nom photo', { bold: true, fontSize: 24 })], {
      width: 2200,
      shaded: true,
    }),
    createTableCell(doc, [createStyledParagraph(doc, 'Photo', { bold: true, fontSize: 24, centered: true })], {
      width: 3600,
      shaded: true,
    }),
    createTableCell(doc, [createStyledParagraph(doc, 'Commentaire', { bold: true, fontSize: 24 })], {
      width: 3800,
      shaded: true,
    }),
  )
  table.append(headerRow)

  entries.forEach(({ photo, relationshipId, drawingId, index }) => {
    const contentRow = createWordElement(doc, 'tr')
    contentRow.append(
      createTableCell(doc, [createStyledParagraph(doc, photo.name || `Photo ${index + 1}`, { fontSize: 22 })], {
        width: 2200,
        centered: true,
      }),
      createTableCell(doc, [createPhotoParagraph(doc, photo, relationshipId, drawingId)], {
        width: 3600,
        centered: true,
      }),
      createTableCell(doc, [createStyledParagraph(doc, photo.caption || '', { fontSize: 22 })], {
        width: 3800,
      }),
    )
    table.append(contentRow)
  })

  return table
}

function appendPhotoAppendix(documentDoc, relationshipsDoc, contentTypesDoc, preparedPhotos) {
  if (!preparedPhotos.length) return

  const body = getWordDescendants(documentDoc, 'body')[0]
  const sectionProperties = getWordChildren(body, 'sectPr')[0] || null
  const insertBefore = sectionProperties
  const relationshipNodes = Array.from(relationshipsDoc.getElementsByTagNameNS(RELS_NS, 'Relationship'))
  let nextRelationshipId = relationshipNodes.reduce((maxId, node) => {
    const value = Number((node.getAttribute('Id') || '').replace('rId', ''))
    return Number.isFinite(value) ? Math.max(maxId, value) : maxId
  }, 0) + 1

  const appendNode = (node) => {
    if (insertBefore) body.insertBefore(node, insertBefore)
    else body.append(node)
  }

  appendNode(createStyledParagraph(documentDoc, '', {
    pageBreakBefore: true,
  }))
  appendNode(createStyledParagraph(documentDoc, 'Annexe photos', {
    bold: true,
    fontSize: 32,
    centered: true,
    spacingAfter: 240,
  }))

  const photoEntries = preparedPhotos.map((photo, index) => {
    const extension = SUPPORTED_PHOTO_TYPES[photo.file.type]
    const relationshipId = `rId${nextRelationshipId}`
    nextRelationshipId += 1

    ensureContentTypeDefault(contentTypesDoc, extension, photo.file.type)

    const relationship = relationshipsDoc.createElementNS(RELS_NS, 'Relationship')
    relationship.setAttribute('Id', relationshipId)
    relationship.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
    relationship.setAttribute('Target', `media/photo-${index + 1}.${extension}`)
    relationshipsDoc.documentElement.append(relationship)
    return {
      photo,
      relationshipId,
      drawingId: 1000 + index,
      index,
    }
  })

  appendNode(createPhotoTable(documentDoc, photoEntries))
}

function PhotoSection({ photos, photoMessage, onAddPhotos, onRemovePhoto, onCaptionChange, onNameChange }) {
  return (
    <section className="photo-section">
      <div className="photo-section-header">
        <div>
          <p className="eyebrow">Photos</p>
          <h2>Constats photo</h2>
          <p className="photo-copy">
            Ajoutez des photos depuis le telephone ou l appareil photo. Elles seront ajoutees
            a la fin du meme fichier Word que la fiche principale.
          </p>
        </div>

        <div className="photo-actions">
          <label className="photo-picker">
            <span>Prendre une photo</span>
            <input type="file" accept="image/jpeg,image/png" capture="environment" multiple onChange={onAddPhotos} />
          </label>
          <label className="photo-picker">
            <span>Choisir des photos</span>
            <input type="file" accept="image/jpeg,image/png" multiple onChange={onAddPhotos} />
          </label>
        </div>
      </div>

      <div className="draft-note">Les photos restent uniquement dans cette session et ne sont pas sauvegardees dans le brouillon local.</div>
      {photoMessage ? <div className="share-note">{photoMessage}</div> : null}

      {photos.length ? (
        <div className="photo-grid">
          {photos.map((photo, index) => (
            <article key={photo.id} className="photo-card">
              <img src={photo.previewUrl} alt={photo.caption || `Photo ${index + 1}`} className="photo-preview" />
              <div className="photo-card-body">
                <div className="photo-index">Photo {index + 1}</div>
                <TextInput
                  value={photo.name}
                  placeholder="Nom de la photo"
                  onChange={(value) => onNameChange(photo.id, value)}
                />
                <TextInput
                  value={photo.caption}
                  placeholder="Commentaire"
                  onChange={(value) => onCaptionChange(photo.id, value)}
                />
                <button type="button" className="text-action" onClick={() => onRemovePhoto(photo.id)}>
                  Retirer cette photo
                </button>
              </div>
            </article>
          ))}
        </div>
      ) : (
        <div className="photo-empty">Aucune photo ajoutee pour le moment.</div>
      )}
    </section>
  )
}

function ScalianSignature() {
  return (
    <footer className="scalian-signature" aria-label="Copyright Scalian">
      <a
        className="scalian-logo-link"
        href="https://www.scalian.com/"
        target="_blank"
        rel="noreferrer"
        aria-label="Open Scalian website"
      >
        <img src={SCALIAN_LOGO_SRC} alt="Scalian" className="scalian-logo" />
      </a>
      <div className="scalian-copy">
        <div className="scalian-text">Copyright © Scalian DS</div>
        <div className="scalian-developer">Developed by Hai Dang VU - haidang.vu@scalian.com</div>
      </div>
    </footer>
  )
}

function AppTabs({ activeTab, onChange }) {
  const tabs = [
    ['inspection', 'Inspection'],
    ['generateur', 'Generateur'],
    ['historique', 'Historique'],
    ['statistiques', 'Statistiques'],
  ]

  return (
    <nav className="app-tabs" aria-label="Navigation principale">
      {tabs.map(([id, label]) => (
        <button
          key={id}
          type="button"
          className={`app-tab${activeTab === id ? ' active' : ''}`}
          onClick={() => onChange(id)}
        >
          {label}
        </button>
      ))}
    </nav>
  )
}

function downloadBrowserBlob(blob, fileName) {
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')

  link.href = url
  link.download = fileName
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
  URL.revokeObjectURL(url)
}

function getGeneratedOutputFileName(schema) {
  const baseName = String(schema?.fileName || 'formulaire')
    .replace(/\.[^.]+$/, '')
    .replace(/[\\/:*?"<>|]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
  const extension = schema?.kind === 'pdf' ? 'pdf' : 'docx'

  return `${baseName || 'formulaire'} - rempli depuis app.${extension}`
}

function isPdfFile(file) {
  return file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')
}

function isDocxFile(file) {
  return file.name.toLowerCase().endsWith('.docx')
}

function GenericFormGenerator({ onArchiveRecord, onArchiveMessage }) {
  const [templateFile, setTemplateFile] = useState(null)
  const [schema, setSchema] = useState(null)
  const [values, setValues] = useState({})
  const [message, setMessage] = useState('')
  const [isWorking, setIsWorking] = useState(false)

  async function handleTemplateSelection(event) {
    const file = event.target.files?.[0] || null
    if (!file) return

    setIsWorking(true)
    setMessage(isPdfFile(file) ? 'Analyse du PDF en cours...' : 'Analyse du document Word en cours...')

    try {
      if (!isPdfFile(file) && !isDocxFile(file)) {
        throw new Error('Format non pris en charge. Importez un fichier .docx ou un PDF remplissable.')
      }

      const nextSchema = isPdfFile(file)
        ? await parseGenericPdfTemplate(file)
        : await parseGenericDocxTemplate(file)
      setTemplateFile(file)
      setSchema(nextSchema)
      setValues(
        nextSchema.kind === 'pdf'
          ? createInitialGenericPdfValues(nextSchema)
          : createInitialGenericValues(nextSchema),
      )
      setMessage(
        nextSchema.fieldCount
          ? nextSchema.mode === 'flat'
            ? `${nextSchema.fieldCount} champ(s) proposes depuis le texte du PDF plat. L export ajoutera une page de reponses au PDF.`
            : `${nextSchema.fieldCount} champ(s) detecte(s). L app correspondante est prete.`
          : nextSchema.kind === 'pdf'
            ? 'Aucun champ PDF remplissable ni texte exploitable detecte. Un PDF scanne demandera une etape OCR/layout.'
            : 'Aucun champ modifiable detecte automatiquement dans ce document.',
      )
    } catch (error) {
      setTemplateFile(null)
      setSchema(null)
      setValues({})
      setMessage(error instanceof Error ? error.message : 'Impossible d analyser ce document Word.')
    } finally {
      setIsWorking(false)
      event.target.value = ''
    }
  }

  function setGenericValue(fieldId, value) {
    setValues((current) => ({ ...current, [fieldId]: value }))
  }

  async function exportGeneratedDocx() {
    if (!templateFile || !schema) return

    setIsWorking(true)
    setMessage('Generation du document Word...')

    try {
      const blob = schema.kind === 'pdf'
        ? await buildGenericPdfBlob(templateFile, schema, values)
        : await buildGenericDocxBlob(templateFile, schema, values)
      const fileName = getGeneratedOutputFileName(schema)
      const archiveFile = new File([blob], fileName, { type: schema.kind === 'pdf' ? 'application/pdf' : DOCX_MIME_TYPE })
      const record = createGeneratedFormArchiveRecord({ schema, values, storedFileName: fileName })

      downloadBrowserBlob(blob, fileName)
      await saveArchiveFile(record.id, archiveFile)
      onArchiveRecord(record)
      onArchiveMessage('Le formulaire genere a ete ajoute a l historique local.')
      setMessage(schema.kind === 'pdf' ? 'PDF genere et archive depuis l app dynamique.' : 'Document Word genere et archive depuis l app dynamique.')
    } catch (error) {
      setMessage(error instanceof Error ? error.message : 'Impossible de generer le document Word.')
    } finally {
      setIsWorking(false)
    }
  }

  return (
    <section className="placeholder-panel generator-panel">
      <div className="history-header">
        <div>
          <p className="eyebrow">Generateur de formulaires</p>
          <h2>Creer une app depuis un formulaire</h2>
          <p className="placeholder-copy">
            Importez une fiche Word ou un PDF remplissable. L outil detecte les champs disponibles, puis construit un formulaire web capable de regenerer le fichier rempli.
          </p>
        </div>

        <div className="photo-actions">
          <label className="photo-picker">
            <span>Importer .docx ou .pdf</span>
            <input type="file" accept=".docx,.pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/pdf" onChange={handleTemplateSelection} />
          </label>
          <button type="button" className="primary-action" onClick={exportGeneratedDocx} disabled={!schema?.fieldCount || isWorking}>
            Generer le fichier rempli
          </button>
        </div>
      </div>

      {message ? <div className="share-note">{message}</div> : null}

      {schema ? (
        <div className="generator-summary">
          <div className="stats-overview-card">
            <div className="stats-overview-label">Modele</div>
            <div className="stats-overview-title">{schema.fileName}</div>
            <div className="stats-overview-copy">{schema.kind?.toUpperCase()} - {schema.sections.length} section(s) detectee(s)</div>
          </div>
          <div className="stats-overview-card">
            <div className="stats-overview-label">Champs</div>
            <div className="stats-overview-title">{schema.fieldCount}</div>
            <div className="stats-overview-copy">Textes, choix et cases</div>
          </div>
        </div>
      ) : null}

      {schema?.sections.length ? (
        <div className="generated-form">
          {schema.sections.map((section) => (
            <section key={section.id} className="generated-section">
              <h3>{section.title}</h3>
              <div className="generated-fields">
                {section.fields.map((field) => (
                  <GeneratedField
                    key={field.id}
                    field={field}
                    value={values[field.id]}
                    onChange={(value) => setGenericValue(field.id, value)}
                  />
                ))}
              </div>
            </section>
          ))}
        </div>
      ) : (
        <div className="photo-empty">Importez un formulaire Word ou PDF pour voir l app generee ici.</div>
      )}
    </section>
  )
}

function GeneratedField({ field, value, onChange }) {
  if (field.type === 'text') {
    return (
      <label className="generated-field">
        <span>{field.label}</span>
        <TextInput value={value || ''} onChange={onChange} />
      </label>
    )
  }

  if (field.type === 'checkbox') {
    return (
      <label className="generated-check">
        <input
          className="native-check"
          type="checkbox"
          checked={Boolean(value)}
          onChange={(event) => onChange(event.target.checked)}
        />
        <span>{field.label}</span>
      </label>
    )
  }

  return (
    <div className="generated-field">
      <span>{field.label}</span>
      <div className="generated-choice-row">
        {field.options.map((option) => (
          <button
            key={option.label}
            type="button"
            className={`choice-pill${value === option.label ? ' selected' : ''}`}
            onClick={() => onChange(value === option.label ? '' : option.label)}
          >
            {option.label}
          </button>
        ))}
      </div>
    </div>
  )
}

function getHistoryBadge(record) {
  if (record.sourceType === 'generated_form') return `${record.formKind || 'form'} genere`
  return record.sourceType === 'manual_docx' ? 'Import DOCX' : 'Depuis l app'
}

function getHistoryTitle(record) {
  if (record.sourceType === 'generated_form') return record.sourceName || 'Formulaire genere'
  return record.building || 'Batiment non renseigne'
}

function getHistoryPrimaryCopy(record) {
  if (record.sourceType === 'generated_form') return `${record.schema?.fieldCount || 0} champ(s) detecte(s)`
  return record.installation || 'Installation non renseignee'
}

function HistoryPanel({ records, importMessage, onImportDocx, onClearArchive, onDownloadDocx, onDownloadAllDocxZip }) {
  return (
    <section className="placeholder-panel">
      <div className="history-header">
        <div>
          <p className="eyebrow">Archive locale</p>
          <h2>Historique des inspections</h2>
          <p className="placeholder-copy">
            Importez d anciennes fiches Word deja remplies et consolidez-les avec les fiches creees depuis l application. Vous pouvez aussi selectionner un fichier .docx avec ses photos pour generer une seule archive Word fusionnee.
          </p>
        </div>

        <div className="photo-actions">
          <label className="photo-picker">
            <span>Importer .docx et photos</span>
            <input type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document,image/jpeg,image/png" multiple onChange={onImportDocx} />
          </label>
          <button type="button" className="secondary-action" onClick={onDownloadAllDocxZip} disabled={!records.length}>
            Telecharger tous les fichiers en .zip
          </button>
          <button type="button" className="secondary-action" onClick={onClearArchive} disabled={!records.length}>
            Vider l historique
          </button>
        </div>
      </div>

      <div className="draft-note">Les inspections archivees ici sont stockees localement sur cet appareil tant qu aucune base de donnees n est connectee.</div>
      {importMessage ? <div className="share-note">{importMessage}</div> : null}

      {records.length ? (
        <div className="history-list">
          {records.map((record) => (
            <article key={record.id} className="history-card">
              <div className="history-meta">
                <span className="history-badge">{getHistoryBadge(record)}</span>
                <span className="history-date">{new Date(record.createdAt).toLocaleString('fr-FR')}</span>
              </div>
              <div className="history-title">{getHistoryTitle(record)}</div>
              <div className="history-copy">{getHistoryPrimaryCopy(record)}</div>
              {record.sourceType !== 'generated_form' ? (
                <>
                  <div className="history-copy">{record.inspectionDate || 'Date non renseignee'}</div>
                  <div className="history-copy">{record.quality || 'Qualite non renseignee'}</div>
                </>
              ) : null}
              {record.sourceName ? <div className="history-copy">Source: {record.sourceName}</div> : null}
              {record.storedFileName ? <div className="history-copy">Fichier: {record.storedFileName}</div> : null}
              <button type="button" className="secondary-action history-action" onClick={() => onDownloadDocx(record)}>
                Telecharger le fichier
              </button>
            </article>
          ))}
        </div>
      ) : (
        <div className="photo-empty">Aucune fiche archivee pour le moment. Importez un ancien .docx ou generez une nouvelle fiche depuis l application.</div>
      )}
    </section>
  )
}

function formatResponseStats(item) {
  return item.responseStats
    .map((stat) => `${stat.label} ${stat.count}/${item.totalCount} (${stat.percentage}%)`)
    .join(' | ')
}

function formatMaterialStats(item) {
  const presentStats = `OUI ${item.presentCount}/${item.totalCount} (${item.presentPercentage}%)`
  if (!item.materialStats.length) return presentStats

  return `${presentStats} | Etat du materiel: ${item.materialStats
    .map((stat) => `${stat.label} ${stat.count}/${item.presentCount} (${stat.percentage}%)`)
    .join(' | ')}`
}

function StatChip({ label, count, total, percentage, tone = 'default' }) {
  return (
    <span className={`stats-chip stats-chip-${tone}`}>
      <strong>{percentage}%</strong>
      <span>{label}</span>
      <small>{count}/{total}</small>
    </span>
  )
}

function ResponseStatsChips({ item }) {
  return (
    <div className="stats-chip-row">
      {item.responseStats.map((stat) => (
        <StatChip
          key={stat.label}
          label={stat.label}
          count={stat.count}
          total={item.totalCount}
          percentage={stat.percentage}
          tone={stat.label === 'NON' ? 'warning' : stat.label === 'Sans objet' ? 'muted' : 'success'}
        />
      ))}
    </div>
  )
}

function MaterialStatsChips({ item }) {
  return (
    <div className="stats-chip-row">
      <StatChip
        label="Present"
        count={item.presentCount}
        total={item.totalCount}
        percentage={item.presentPercentage}
        tone="success"
      />
      {item.materialStats.map((stat) => (
        <StatChip
          key={stat.label}
          label={stat.label}
          count={stat.count}
          total={item.presentCount}
          percentage={stat.percentage}
          tone={stat.label === 'Vetuste' ? 'warning' : stat.label === "Etat d'usage" ? 'muted' : 'success'}
        />
      ))}
    </div>
  )
}

function GenericTemplateStats({ templates }) {
  if (!templates.length) {
    return (
      <section className="placeholder-panel">
        <p className="eyebrow">Formulaires generes</p>
        <h2>Aucune statistique generique</h2>
        <p className="placeholder-copy">
          Les formulaires generes apparaitront ici apres export depuis l onglet Generateur.
        </p>
      </section>
    )
  }

  return (
    <>
      {templates.map((template) => (
        <section key={template.key} className="placeholder-panel">
          <p className="eyebrow">{template.kind?.toUpperCase()} genere</p>
          <h2>{template.title}</h2>
          <p className="placeholder-copy">
            {template.recordCount} archive(s) pour ce modele. Les statistiques suivent les champs detectes dans le formulaire source.
          </p>
          <div className="placeholder-points">
            {template.sections.flatMap((section) => section.fields.map((field) => (
              <GenericFieldStat key={`${section.id}-${field.id}`} section={section} field={field} />
            )))}
          </div>
        </section>
      ))}
    </>
  )
}

function GenericFieldStat({ section, field }) {
  return (
    <div className="placeholder-point stats-item">
      <div className="stats-item-head">
        <span className="stats-item-id">{section.title}</span>
        <span className="stats-item-title">{field.label}</span>
      </div>
      <div className="stats-chip-row">
        <StatChip
          label={field.type === 'checkbox' ? 'Coche' : 'Renseigne'}
          count={field.filledCount}
          total={field.totalCount}
          percentage={field.filledPercentage}
          tone="success"
        />
        {field.valueStats.map((stat) => (
          <StatChip
            key={stat.label}
            label={stat.label}
            count={stat.count}
            total={field.totalCount}
            percentage={stat.percentage}
            tone={stat.label === 'Vide' || stat.label === 'Non coche' ? 'muted' : 'default'}
          />
        ))}
      </div>
      <div className="stats-item-note">
        Type: {field.type}
      </div>
    </div>
  )
}

function StatisticsPanel({ records, onExportStats }) {
  const {
    importedCount,
    appCount,
    doneBuildingsCount,
    genericCount,
    saclayCount,
    totalArchives,
    sectionBuildingGroups,
    groupedRemarks,
    genericTemplates,
  } = buildStatisticsSnapshot(records)

  return (
    <section className="stats-panel">
      <div className="stats-toolbar">
        <div className="draft-note">Vue simplifiee des archives: statistiques Saclay et champs des formulaires generes.</div>
        <button type="button" className="secondary-action" onClick={onExportStats}>
          Generer statistiques .docx
        </button>
      </div>

      <section className="stats-group">
        <div className="stats-group-head">
          <p className="eyebrow">Tendances archivees</p>
          <h2>Archives par workflow</h2>
          <p className="placeholder-copy">
            Les formulaires generes sont groupes par modele source. Les fiches Saclay gardent leurs statistiques historiques par batiment, section et remarques.
          </p>
        </div>

        <div className="stats-grid">
          <article className="stat-card">
            <div className="stat-label">Archives</div>
            <div className="stat-value">{totalArchives}</div>
            <div className="stat-copy">Fiches Saclay et formulaires generes archives localement</div>
          </article>
          <article className="stat-card">
            <div className="stat-label">Formulaires generes</div>
            <div className="stat-value">{genericCount}</div>
            <div className="stat-copy">Archives issues du generateur DOCX/PDF</div>
          </article>
          <article className="stat-card">
            <div className="stat-label">Fiches Saclay</div>
            <div className="stat-value">{saclayCount}</div>
            <div className="stat-copy">Archives analysees avec les statistiques historiques Saclay</div>
          </article>
          <article className="stat-card">
            <div className="stat-label">Imports DOCX</div>
            <div className="stat-value">{importedCount}</div>
            <div className="stat-copy">Anciennes fiches manuelles reintegrees dans l historique</div>
          </article>
          <article className="stat-card">
            <div className="stat-label">Fiches app</div>
            <div className="stat-value">{appCount}</div>
            <div className="stat-copy">Fiches nouvelles archivees depuis cette application</div>
          </article>
          <article className="stat-card">
            <div className="stat-label">Batiments faits</div>
            <div className="stat-value">{doneBuildingsCount}/{TOTAL_BUILDINGS_TARGET}</div>
            <div className="stat-copy">Nombre de batiments distincts deja couverts par les fiches archivees</div>
          </article>
        </div>

        <div className="stats-sections">
          <GenericTemplateStats templates={genericTemplates} />

          {sectionBuildingGroups.map((section) => (
            <section key={section.id} className="placeholder-panel">
              <p className="eyebrow">Section {section.id}</p>
              <h2>{section.title}</h2>
              <p className="placeholder-copy">{section.description}</p>
              {section.id === '5' ? (
                <>
                  <div className="chart-list">
                    {section.qualityStats.map((item) => (
                      <div key={item.label} className="chart-row">
                        <div className="chart-head">
                          <span className="chart-label">{item.symbol} {item.label}</span>
                          <span className="chart-count">
                            <strong>{item.percentage}%</strong>
                            <span>{item.count}/{item.totalCount}</span>
                          </span>
                        </div>
                        <div className="chart-track">
                          <div
                            className="chart-bar"
                            style={{
                              width: `${item.percentage}%`,
                              background: item.color,
                            }}
                          />
                        </div>
                        <div className="chart-buildings">
                          {item.buildings.length ? `Batiments: ${item.buildings.join(', ')}` : 'Aucun batiment'}
                        </div>
                      </div>
                    ))}
                  </div>

                  <div className="placeholder-copy" style={{ marginTop: '1rem' }}>
                    Remise a l etat a prevoir: details des batiments et remarques.
                  </div>
                  <div className="placeholder-points">
                    {section.remiseDetails.length ? section.remiseDetails.map((item) => (
                      <div key={item.building} className="placeholder-point">
                        {item.building}: {item.remarks || 'Aucune remarque'}
                      </div>
                    )) : (
                      <div className="placeholder-point">Aucun batiment en remise a l etat a prevoir.</div>
                    )}
                  </div>
                </>
              ) : (
                <div className="placeholder-points">
                  {section.items.map((item) => (
                    <div key={item.key} className="placeholder-point stats-item">
                      <div className="stats-item-head">
                        <span className="stats-item-id">{item.key}</span>
                        <span className="stats-item-title">{item.label}</span>
                      </div>
                      {['1', '2', '3'].includes(section.id) ? (
                        <>
                          <ResponseStatsChips item={item} />
                          <div className="stats-item-note">
                            Batiments NON: {item.buildings.length ? item.buildings.join(', ') : 'Aucun batiment'}
                          </div>
                        </>
                      ) : section.id === '4' ? (
                        <MaterialStatsChips item={item} />
                      ) : section.id === '6' ? (
                        <>
                          <div className="stats-chip-row">
                            <StatChip
                              label="Coche"
                              count={item.count}
                              total={item.totalCount}
                              percentage={item.percentage}
                              tone="success"
                            />
                          </div>
                          <div className="stats-item-note">
                            Batiments: {item.buildings.length ? item.buildings.join(', ') : 'Aucun batiment'}
                          </div>
                        </>
                      ) : (
                        item.buildings.length ? item.buildings.join(', ') : 'Aucun batiment'
                      )}
                    </div>
                  ))}
                </div>
              )}
            </section>
          ))}

          <section className="placeholder-panel">
            <p className="eyebrow">Remarques</p>
            <h2>Batiments regroupes par remarques similaires</h2>
            <p className="placeholder-copy">
              Les remarques sont consolidees a partir des commentaires de qualite, des points d amelioration et des autres remarques.
            </p>
            <div className="placeholder-points">
              {groupedRemarks.length ? groupedRemarks.map((group) => (
                <div key={group.remark} className="placeholder-point">
                  <div className="stats-item-head">
                    <span className="stats-item-title">{group.remark}</span>
                  </div>
                  <div className="stats-item-note">
                    Batiments: {group.buildings.join(', ')}
                  </div>
                  {group.examples?.length ? (
                    <div className="stats-item-note">
                      Formulations: {group.examples.join(' ; ')}
                    </div>
                  ) : null}
                </div>
              )) : (
                <div className="placeholder-point">Aucune remarque archivee pour le moment.</div>
              )}
            </div>
          </section>
        </div>
      </section>
    </section>
  )
}

function fillChoiceTable(table, rows, data, commentKey) {
  const tableRows = getTableRows(table)
  rows.forEach(([key, , options], index) => {
    const cells = getRowCells(tableRows[index + 1])
    setFirstCheckboxInNode(cells[2], data[key] === 'OUI')
    setFirstCheckboxInNode(cells[3], data[key] === 'NON')
    if (options.includes('Sans objet')) {
      setFirstCheckboxInNode(cells[4], data[key] === 'Sans objet')
    }
  })

  const commentRow = tableRows[rows.length + 1]
  const commentCell = getRowCells(commentRow)[4] || getRowCells(commentRow)[0]
  const commentParagraph = getNextWordParagraph(getWordChildren(commentCell, 'p')[0])
  if (commentParagraph) setParagraphText(commentParagraph, data[commentKey] || '')
}

function fillTemplateDoc(doc, data) {
  const generalTable = findTableByText(doc, 'Bâtiment(s)')
  const generalRows = getTableRows(generalTable)
  setCellText(getRowCells(generalRows[0])[1], data.batiments)
  setCellText(getRowCells(generalRows[0])[4], data.dateHeure)
  setCellText(getRowCells(generalRows[1])[1], data.localNumero)
  setCellText(getRowCells(generalRows[2])[1], data.installation)
  setCellText(getRowCells(generalRows[2])[4], data.redacteur)
  setCellText(getRowCells(generalRows[3])[1], data.intervenantPresent)
  setCellText(getRowCells(generalRows[3])[4], data.visa)

  fillChoiceTable(findTableByText(doc, 'Formations et Habilitations'), SECTION_1, data, 'comments1')
  fillChoiceTable(findTableByText(doc, 'Equipements des intervenants'), SECTION_2, data, 'comments2')

  const productsTable = findTableByText(doc, 'Produits de nettoyage')
  const productRows = getTableRows(productsTable)
  SECTION_3.forEach(([key], index) => {
    const cells = getRowCells(productRows[index + 1])
    setFirstCheckboxInNode(cells[2], data[key].status === 'OUI')
    setFirstCheckboxInNode(cells[3], data[key].status === 'NON')
    setCellText(cells[4], data[key].name || '')
  })
  setParagraphAfterLabel(doc, 'Commentaires', data.comments3, 2)

  const materialsTable = findTableByText(doc, 'Matériels / documents')
  const materialRows = getTableRows(materialsTable)
  SECTION_4.forEach(([key, , withState], index) => {
    const cells = getRowCells(materialRows[index + 1])
    setFirstCheckboxInNode(cells[2], data[key].status === 'OUI')
    setFirstCheckboxInNode(cells[3], data[key].status === 'NON')

    if (withState) {
      const stateCheckboxes = getWordDescendants(cells[4], 'sdt')
      setCheckboxState(stateCheckboxes[0], data[key].state === 'Bon état')
      setCheckboxState(stateCheckboxes[1], data[key].state === 'Etat d’usage')
      setCheckboxState(stateCheckboxes[2], data[key].state === 'Vétuste')
    }
  })
  setParagraphAfterLabel(doc, 'Commentaires', data.comments4, 3)

  QUALITY_OPTIONS.forEach(([label]) => {
    const paragraph = findParagraphByText(doc, label)
    if (paragraph) setFirstCheckboxInNode(paragraph, data.quality === label)
  })

  IMPROVEMENTS.forEach((label) => {
    const paragraph = findParagraphByText(doc, label)
    if (paragraph) setFirstCheckboxInNode(paragraph, data.improve.includes(label))
  })

  setParagraphAfterLabel(doc, 'Commentaires', data.qualityComments, 4)
  setParagraphAfterLabel(doc, 'Commentaires', data.improveComments, 5)
  setParagraphAfterLabel(doc, 'Autres Remarques', data.otherRemarks)
}

function ChoiceControl({ value, currentValue, onChange }) {
  return (
    <button
      type="button"
      className={`mark-button${currentValue === value ? ' selected' : ''}`}
      onClick={() => onChange(value)}
      aria-pressed={currentValue === value}
    >
      {currentValue === value ? '☒' : '☐'}
    </button>
  )
}

function TextInput({ value, onChange, placeholder }) {
  return (
    <input
      className="cell-input"
      value={value}
      placeholder={placeholder}
      onChange={(event) => onChange(event.target.value)}
    />
  )
}

function TextArea({ value, onChange, placeholder, rows = 4 }) {
  return (
    <textarea
      className="comments-input"
      rows={rows}
      value={value}
      placeholder={placeholder}
      onChange={(event) => onChange(event.target.value)}
    />
  )
}

function PageHeader({ pageNumber }) {
  return (
    <table className="doc-header-table" aria-label={`En-tete page ${pageNumber}`}>
      <tbody>
        <tr>
          <td rowSpan="2" className="doc-logo-cell">
            <img src={LOGO_SRC} alt="CEA" className="doc-logo" />
          </td>
          <td rowSpan="2" className="doc-org-cell">
            <div>Direction Générale</div>
            <div>Département de Soutien Scientifique et Technique</div>
          </td>
          <td className="doc-title-cell">{HEADER_TITLE}</td>
          <td className="doc-page-cell">
            Page
            <br />
            {pageNumber}/2
          </td>
        </tr>
        <tr>
          <td colSpan="2" className="doc-subtitle-cell">{HEADER_SUBTITLE}</td>
        </tr>
      </tbody>
    </table>
  )
}

function PrintChoiceTable({ titleNumber, title, rows, data, commentKey }) {
  return (
    <table className="print-table">
      <tbody>
        <tr className="print-section-row">
          <td className="print-number-col">{titleNumber}</td>
          <td>{title}</td>
          <td className="print-choice-col">OUI</td>
          <td className="print-choice-col">NON</td>
          <td className="print-choice-col">Sans objet</td>
        </tr>
        {rows.map(([key, label, options]) => (
          <tr key={key}>
            <td className="print-number-col">{key}</td>
            <td>{label}</td>
            <td className="print-center-cell">{checkboxMark(data[key], 'OUI')}</td>
            <td className="print-center-cell">{checkboxMark(data[key], 'NON')}</td>
            <td className="print-center-cell">{options.includes('Sans objet') ? checkboxMark(data[key], 'Sans objet') : ''}</td>
          </tr>
        ))}
        <tr>
          <td colSpan="5" className="print-comments-cell">
            <span className="print-comments-label">Commentaires :</span>
            <div className="print-comment-box">{displayValue(data[commentKey])}</div>
          </td>
        </tr>
      </tbody>
    </table>
  )
}

function PrintProductsTable({ data }) {
  return (
    <table className="print-table">
      <tbody>
        <tr className="print-section-row">
          <td className="print-number-col">3</td>
          <td>Produits de nettoyage</td>
          <td className="print-choice-col">OUI</td>
          <td className="print-choice-col">NON</td>
          <td>Nom du produit</td>
        </tr>
        {SECTION_3.map(([key, label]) => (
          <tr key={key}>
            <td className="print-number-col">{key}</td>
            <td>{label}</td>
            <td className="print-center-cell">{checkboxMark(data[key].status, 'OUI')}</td>
            <td className="print-center-cell">{checkboxMark(data[key].status, 'NON')}</td>
            <td>{displayValue(data[key].name)}</td>
          </tr>
        ))}
        <tr>
          <td colSpan="5" className="print-comments-cell">
            <span className="print-comments-label">Commentaires :</span>
            <div className="print-comment-box">{displayValue(data.comments3)}</div>
          </td>
        </tr>
      </tbody>
    </table>
  )
}

function PrintMaterialsTable({ data }) {
  return (
    <table className="print-table">
      <tbody>
        <tr className="print-section-row">
          <td className="print-number-col">4</td>
          <td>Matériels / documents</td>
          <td className="print-choice-col">OUI</td>
          <td className="print-choice-col">NON</td>
          <td>Etat du matériel</td>
        </tr>
        {SECTION_4.map(([key, label, withState]) => (
          <tr key={key}>
            <td className="print-number-col">{key}</td>
            <td>{label}</td>
            <td className="print-center-cell">{checkboxMark(data[key].status, 'OUI')}</td>
            <td className="print-center-cell">{checkboxMark(data[key].status, 'NON')}</td>
            <td className={!withState ? 'print-disabled-cell' : ''}>
              {withState ? (
                <div className="print-state-line">
                  <span>{stateMark(data[key].state, 'Bon état')} Bon état</span>
                  <span>{stateMark(data[key].state, 'Etat d’usage')} Etat d’usage</span>
                  <span>{stateMark(data[key].state, 'Vétuste')} Vétuste</span>
                </div>
              ) : null}
            </td>
          </tr>
        ))}
        <tr>
          <td colSpan="5" className="print-comments-cell">
            <span className="print-comments-label">Commentaires :</span>
            <div className="print-comment-box">{displayValue(data.comments4)}</div>
          </td>
        </tr>
      </tbody>
    </table>
  )
}

function ScreenPageOne({ data, setField, setNestedField }) {
  return (
    <div className="paper-page">
      <PageHeader pageNumber={1} />
      <table className="inspection-table">
        <tbody>
          <tr>
            <td className="label-cell">Bâtiment(s) :</td>
            <td><TextInput value={data.batiments} onChange={(value) => setField('batiments', value)} /></td>
            <td className="spacer-cell" />
            <td rowSpan="2" className="label-cell">Date et heure du constat :</td>
            <td rowSpan="2"><TextInput value={data.dateHeure} onChange={(value) => setField('dateHeure', value)} /></td>
          </tr>
          <tr>
            <td className="label-cell">Numéro du local de nettoyage :</td>
            <td><TextInput value={data.localNumero} onChange={(value) => setField('localNumero', value)} /></td>
            <td className="spacer-cell" />
          </tr>
          <tr>
            <td className="label-cell">Installation :</td>
            <td><TextInput value={data.installation} onChange={(value) => setField('installation', value)} /></td>
            <td className="spacer-cell" />
            <td className="label-cell">Rédacteur :</td>
            <td><TextInput value={data.redacteur} onChange={(value) => setField('redacteur', value)} /></td>
          </tr>
          <tr>
            <td className="label-cell">Intervenant(s) ATALIAN présent (oui/non) :</td>
            <td><TextInput value={data.intervenantPresent} onChange={(value) => setField('intervenantPresent', value)} /></td>
            <td className="spacer-cell" />
            <td className="label-cell">Visa :</td>
            <td><TextInput value={data.visa} onChange={(value) => setField('visa', value)} /></td>
          </tr>
        </tbody>
      </table>

      <SectionChoiceTable
        titleNumber="1"
        title="Formations et Habilitations"
        rows={SECTION_1}
        data={data}
        commentKey="comments1"
        setField={setField}
      />

      <SectionChoiceTable
        titleNumber="2"
        title="Equipements des intervenants"
        rows={SECTION_2}
        data={data}
        commentKey="comments2"
        setField={setField}
      />

      <table className="inspection-table">
        <tbody>
          <tr className="section-row">
            <td className="number-col">3</td>
            <td>Produits de nettoyage</td>
            <td className="choice-heading">OUI</td>
            <td className="choice-heading">NON</td>
            <td className="wide-col">Nom du produit</td>
          </tr>
          {SECTION_3.map(([key, label]) => (
            <tr key={key}>
              <td className="number-col">{key}</td>
              <td>{label}</td>
              <td className="center-cell">
                <ChoiceControl
                  value="OUI"
                  currentValue={data[key].status}
                  onChange={(value) => setNestedField(key, { status: value })}
                />
              </td>
              <td className="center-cell">
                <ChoiceControl
                  value="NON"
                  currentValue={data[key].status}
                  onChange={(value) => setNestedField(key, { status: value })}
                />
              </td>
              <td>
                <TextInput
                  value={data[key].name}
                  placeholder="Nom du produit"
                  onChange={(value) => setNestedField(key, { name: value })}
                />
              </td>
            </tr>
          ))}
          <tr>
            <td colSpan="5" className="comments-cell">
              <span className="comments-label">Commentaires :</span>
              <TextArea value={data.comments3} onChange={(value) => setField('comments3', value)} rows={4} />
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  )
}

function ScreenPageTwo({ data, setField, setNestedField, toggleImprove }) {
  return (
    <div className="paper-page">
      <PageHeader pageNumber={2} />
      <table className="inspection-table">
        <tbody>
          <tr className="section-row">
            <td className="number-col">4</td>
            <td>Matériels / documents</td>
            <td className="choice-heading">OUI</td>
            <td className="choice-heading">NON</td>
            <td className="wide-col">Etat du matériel</td>
          </tr>
          {SECTION_4.map(([key, label, withState]) => (
            <tr key={key}>
              <td className="number-col">{key}</td>
              <td>{label}</td>
              <td className="center-cell">
                <ChoiceControl
                  value="OUI"
                  currentValue={data[key].status}
                  onChange={(value) => setNestedField(key, { status: value })}
                />
              </td>
              <td className="center-cell">
                <ChoiceControl
                  value="NON"
                  currentValue={data[key].status}
                  onChange={(value) => setNestedField(key, { status: value })}
                />
              </td>
              <td className={withState ? '' : 'disabled-cell'}>
                {withState ? (
                  <div className="state-stack">
                    {['Bon état', 'Etat d’usage', 'Vétuste'].map((state) => (
                      <label key={state} className="state-option">
                        <ChoiceControl
                          value={state}
                          currentValue={data[key].state}
                          onChange={(value) => setNestedField(key, { state: value })}
                        />
                        <span>{state}</span>
                      </label>
                    ))}
                  </div>
                ) : null}
              </td>
            </tr>
          ))}
          <tr>
            <td colSpan="5" className="comments-cell">
              <span className="comments-label">Commentaires :</span>
              <TextArea value={data.comments4} onChange={(value) => setField('comments4', value)} rows={4} />
            </td>
          </tr>
        </tbody>
      </table>

      <div className="quality-section">
        <div className="quality-title">5</div>
        <div className="quality-heading">Qualité de la prestation</div>
      </div>

      <div className="quality-grid">
        <section className="quality-card">
          <div className="quality-label">5.1</div>
          <p className="quality-copy">Etat général du bâtiment :</p>
          {QUALITY_OPTIONS.map(([label, symbol, color]) => (
            <label key={label} className="quality-option">
              <ChoiceControl value={label} currentValue={data.quality} onChange={(value) => setField('quality', value)} />
              <span className="quality-symbol" style={{ color }}>
                {symbol}
              </span>
              <span>{label}</span>
            </label>
          ))}
          <div className="quality-comments">
            <span className="comments-label">Commentaires :</span>
            <TextArea value={data.qualityComments} onChange={(value) => setField('qualityComments', value)} rows={4} />
          </div>
        </section>

        <section className="quality-card">
          <div className="quality-label">5.2</div>
          <p className="quality-copy">Points d’améliorations</p>
          {IMPROVEMENTS.map((item) => (
            <label key={item} className="quality-option">
              <input
                className="native-check"
                type="checkbox"
                checked={data.improve.includes(item)}
                onChange={() => toggleImprove(item)}
              />
              <span>{item}</span>
            </label>
          ))}
          <div className="quality-comments">
            <span className="comments-label">Commentaires :</span>
            <TextArea value={data.improveComments} onChange={(value) => setField('improveComments', value)} rows={4} />
          </div>
        </section>
      </div>

      <section className="remarks-box">
        <span className="comments-label">Autres Remarques :</span>
        <TextArea value={data.otherRemarks} onChange={(value) => setField('otherRemarks', value)} rows={6} />
      </section>
    </div>
  )
}

function PrintPageOne({ data }) {
  return (
    <div className="print-page">
      <PageHeader pageNumber={1} />
      <table className="print-table">
        <tbody>
          <tr>
            <td className="print-label-cell">Bâtiment(s) :</td>
            <td>{displayValue(data.batiments)}</td>
            <td className="print-spacer-cell" />
            <td rowSpan="2" className="print-label-cell">Date et heure du constat :</td>
            <td rowSpan="2">{displayValue(data.dateHeure)}</td>
          </tr>
          <tr>
            <td className="print-label-cell">Numéro du local de nettoyage :</td>
            <td>{displayValue(data.localNumero)}</td>
            <td className="print-spacer-cell" />
          </tr>
          <tr>
            <td className="print-label-cell">Installation :</td>
            <td>{displayValue(data.installation)}</td>
            <td className="print-spacer-cell" />
            <td className="print-label-cell">Rédacteur :</td>
            <td>{displayValue(data.redacteur)}</td>
          </tr>
          <tr>
            <td className="print-label-cell">Intervenant(s) ATALIAN présent (oui/non) :</td>
            <td>{displayValue(data.intervenantPresent)}</td>
            <td className="print-spacer-cell" />
            <td className="print-label-cell">Visa :</td>
            <td>{displayValue(data.visa)}</td>
          </tr>
        </tbody>
      </table>

      <PrintChoiceTable titleNumber="1" title="Formations et Habilitations" rows={SECTION_1} data={data} commentKey="comments1" />
      <PrintChoiceTable titleNumber="2" title="Equipements des intervenants" rows={SECTION_2} data={data} commentKey="comments2" />
      <PrintProductsTable data={data} />
    </div>
  )
}

function PrintPageTwo({ data }) {
  return (
    <div className="print-page">
      <PageHeader pageNumber={2} />
      <PrintMaterialsTable data={data} />

      <div className="print-quality-section">
        <div className="print-quality-title">5</div>
        <div className="print-quality-heading">Qualité de la prestation</div>
      </div>

      <div className="print-quality-grid">
        <section className="print-quality-card">
          <div className="print-quality-label">5.1</div>
          <p className="print-quality-copy">Etat général du bâtiment :</p>
          {QUALITY_OPTIONS.map(([label, symbol, color]) => (
            <div key={label} className="print-quality-option">
              <span>{checkboxMark(data.quality, label)}</span>
              <span className="print-quality-symbol" style={{ color }}>{symbol}</span>
              <span>{label}</span>
            </div>
          ))}
          <div className="print-quality-comments">
            <span className="print-comments-label">Commentaires :</span>
            <div className="print-comment-box">{displayValue(data.qualityComments)}</div>
          </div>
        </section>

        <section className="print-quality-card">
          <div className="print-quality-label">5.2</div>
          <p className="print-quality-copy">Points d’améliorations</p>
          {IMPROVEMENTS.map((item) => (
            <div key={item} className="print-quality-option">
              <span>{data.improve.includes(item) ? '☒' : '☐'}</span>
              <span>{item}</span>
            </div>
          ))}
          <div className="print-quality-comments">
            <span className="print-comments-label">Commentaires :</span>
            <div className="print-comment-box">{displayValue(data.improveComments)}</div>
          </div>
        </section>
      </div>

      <section className="print-remarks-box">
        <span className="print-comments-label">Autres Remarques :</span>
        <div className="print-remarks-content">{displayValue(data.otherRemarks)}</div>
      </section>
    </div>
  )
}

function PrintDocument({ data }) {
  return (
    <section className="print-document" aria-hidden="true">
      <PrintPageOne data={data} />
      <PrintPageTwo data={data} />
    </section>
  )
}

export default function App() {
  const [data, setData] = useState(() => createStoredDraft())
  const [photos, setPhotos] = useState([])
  const [photoMessage, setPhotoMessage] = useState('')
  const [importMessage, setImportMessage] = useState('')
  const [archiveRecords, setArchiveRecords] = useState(() => createStoredArchive())
  const [activeTab, setActiveTab] = useState('inspection')
  const photosRef = useRef([])
  const progress = useMemo(() => completionCount(data), [data])

  useEffect(() => {
    window.localStorage.setItem(DRAFT_STORAGE_KEY, JSON.stringify(data))
  }, [data])

  useEffect(() => {
    photosRef.current = photos
  }, [photos])

  useEffect(() => () => {
    revokePhotoUrls(photosRef.current)
  }, [])

  useEffect(() => {
    window.localStorage.setItem(ARCHIVE_STORAGE_KEY, JSON.stringify(archiveRecords))
  }, [archiveRecords])

  function setField(key, value) {
    setData((current) => ({ ...current, [key]: value }))
  }

  function setNestedField(key, patch) {
    setData((current) => ({ ...current, [key]: { ...current[key], ...patch } }))
  }

  function toggleImprove(item) {
    setData((current) => ({
      ...current,
      improve: current.improve.includes(item)
        ? current.improve.filter((entry) => entry !== item)
        : [...current.improve, item],
    }))
  }

  function handlePhotoSelection(event) {
    const selectedFiles = Array.from(event.target.files || [])
    const incomingFiles = selectedFiles.filter((file) => Object.hasOwn(SUPPORTED_PHOTO_TYPES, file.type))
    const rejectedCount = selectedFiles.length - incomingFiles.length

    if (!incomingFiles.length) {
      setPhotoMessage('Aucune photo exploitable n a ete selectionnee.')
      event.target.value = ''
      return
    }

    const nextPhotos = incomingFiles.map((file, index) => ({
      id: createPhotoId(),
      file,
      name: `Photo ${photos.length + index + 1}`,
      caption: stripFileExtension(file.name),
      previewUrl: URL.createObjectURL(file),
    }))

    setPhotos((current) => [...current, ...nextPhotos])
    setPhotoMessage(
      rejectedCount
        ? `${nextPhotos.length} photo${nextPhotos.length > 1 ? 's ajoutees' : ' ajoutee'} au document. ${rejectedCount} fichier${rejectedCount > 1 ? 's sont incompatibles' : ' est incompatible'} avec Word.`
        : `${nextPhotos.length} photo${nextPhotos.length > 1 ? 's ajoutees' : ' ajoutee'} au document.`,
    )
    event.target.value = ''
  }

  function removePhoto(photoId) {
    setPhotos((current) => {
      const target = current.find((photo) => photo.id === photoId)
      if (target?.previewUrl) URL.revokeObjectURL(target.previewUrl)
      return current.filter((photo) => photo.id !== photoId)
    })
    setPhotoMessage('Photo retiree du rapport photo.')
  }

  function updatePhotoCaption(photoId, caption) {
    setPhotos((current) => current.map((photo) => (
      photo.id === photoId ? { ...photo, caption } : photo
    )))
  }

  function updatePhotoName(photoId, name) {
    setPhotos((current) => current.map((photo) => (
      photo.id === photoId ? { ...photo, name } : photo
    )))
  }

  function resetForm() {
    revokePhotoUrls(photos)
    setPhotos([])
    setData(createInitialData())
    window.localStorage.removeItem(DRAFT_STORAGE_KEY)
    setPhotoMessage('')
  }

  async function handleImportDocx(event) {
    const files = Array.from(event.target.files || [])
    if (!files.length) return

    const imported = []
    const failures = []
    const docxFiles = files.filter((file) => file.name.toLowerCase().endsWith('.docx'))
    const imageFiles = files.filter((file) => Object.hasOwn(SUPPORTED_PHOTO_TYPES, file.type))
    const unsupportedFiles = files.filter((file) => !docxFiles.includes(file) && !imageFiles.includes(file))

    if (unsupportedFiles.length) {
      failures.push(`${unsupportedFiles.length} fichier(s) non pris en charge`)
    }

    if (imageFiles.length) {
      if (docxFiles.length !== 1) {
        failures.push('Pour fusionner des photos, importez exactement un fichier .docx avec ses images dans la meme selection')
      } else {
        try {
          const sourceDocx = docxFiles[0]
          const importedData = await parseImportedDocxFile(sourceDocx)
          const importedPhotos = createImportedPhotoEntries(imageFiles)
          const mergedBlob = await buildDocxBlobForInspection(importedData, importedPhotos)
          const mergedFileName = getImportedMergedFileName(sourceDocx.name, importedData)
          const mergedFile = new File([mergedBlob], mergedFileName, { type: DOCX_MIME_TYPE })
          const record = createArchiveRecord({
            data: importedData,
            photos: importedPhotos,
            sourceType: 'manual_docx',
            sourceName: `${sourceDocx.name} + ${imageFiles.length} photo${imageFiles.length > 1 ? 's' : ''}`,
            storedFileName: mergedFileName,
          })

          await saveArchiveFile(record.id, mergedFile)
          imported.push(record)
        } catch (error) {
          failures.push(`${docxFiles[0]?.name || 'Import'} (${error instanceof Error ? error.message : 'erreur inconnue'})`)
        }
      }
    }

    const plainDocxFiles = imageFiles.length ? [] : docxFiles

    for (const file of plainDocxFiles) {
      try {
        const importedData = await parseImportedDocxFile(file)
        const record = createArchiveRecord({
          data: importedData,
          sourceType: 'manual_docx',
          sourceName: file.name,
          storedFileName: file.name,
        })
        await saveArchiveFile(record.id, file)
        imported.push(record)
      } catch (error) {
        failures.push(`${file.name} (${error instanceof Error ? error.message : 'erreur inconnue'})`)
      }
    }

    if (imported.length) {
      setArchiveRecords((current) => [...imported, ...current])
    }

    if (imported.length && failures.length) {
      setImportMessage(`${imported.length} fiche(s) importee(s). Echecs: ${failures.join(' ; ')}`)
    } else if (imported.length) {
      setImportMessage(
        imageFiles.length
          ? `${imported.length} fiche fusionnee avec ${imageFiles.length} photo${imageFiles.length > 1 ? 's' : ''} dans l historique.`
          : `${imported.length} fiche(s) importee(s) dans l historique.`,
      )
    } else if (failures.length) {
      setImportMessage(`Import impossible: ${failures.join(' ; ')}`)
    }

    event.target.value = ''
  }

  async function clearArchive() {
    setArchiveRecords([])
    setImportMessage('Historique local vide.')
    window.localStorage.removeItem(ARCHIVE_STORAGE_KEY)
    try {
      await clearArchiveFiles()
    } catch {
      setImportMessage('Historique efface mais certains fichiers locaux n ont pas pu etre supprimes.')
    }
  }

  async function downloadArchivedDocx(record) {
    try {
      const storedFile = await getArchiveFile(record.id)
      if (!storedFile) {
        setImportMessage('Le fichier archive correspondant est introuvable sur cet appareil.')
        return
      }

      downloadBlob(storedFile, record.storedFileName || record.sourceName || 'archive.docx')
    } catch {
      setImportMessage('Impossible de telecharger ce fichier archive pour le moment.')
    }
  }

  async function downloadAllArchivedDocxAsZip() {
    try {
      const zip = new JSZip()
      const usedNames = new Set()
      const missingFiles = []

      await Promise.all(archiveRecords.map(async (record) => {
        const storedFile = await getArchiveFile(record.id)

        if (!storedFile) {
          missingFiles.push(record.storedFileName || record.sourceName || record.id)
          return
        }

        const entryName = getUniqueArchiveEntryName(
          record.storedFileName || record.sourceName || `${record.building || 'inspection'}.docx`,
          usedNames,
        )

        zip.file(entryName, storedFile)
      }))

      if (!Object.keys(zip.files).length) {
        setImportMessage('Aucun fichier archive n a pu etre retrouve sur cet appareil.')
        return
      }

      const archiveBlob = await zip.generateAsync({ type: 'blob' })
      downloadBlob(archiveBlob, getArchiveZipFileName())

      if (missingFiles.length) {
        setImportMessage(`Archive .zip telechargee, mais ${missingFiles.length} fichier(s) etai(en)t introuvable(s) localement.`)
      } else {
        setImportMessage(`Archive .zip telechargee avec ${archiveRecords.length} fichier(s).`)
      }
    } catch {
      setImportMessage('Impossible de preparer l archive .zip pour le moment.')
    }
  }

  async function buildDocxBlobForInspection(inspectionData, inspectionPhotos = []) {
    const response = await fetch(TEMPLATE_DOCX_SRC)
    const templateBuffer = await response.arrayBuffer()
    const zip = await JSZip.loadAsync(templateBuffer)
    const documentXml = await zip.file('word/document.xml').async('string')
    const relationshipsXml = await zip.file('word/_rels/document.xml.rels').async('string')
    const contentTypesXml = await zip.file('[Content_Types].xml').async('string')
    const xmlDoc = new DOMParser().parseFromString(documentXml, 'application/xml')
    const relationshipsDoc = new DOMParser().parseFromString(relationshipsXml, 'application/xml')
    const contentTypesDoc = new DOMParser().parseFromString(contentTypesXml, 'application/xml')

    fillTemplateDoc(xmlDoc, inspectionData)

    if (inspectionPhotos.length) {
      const preparedPhotos = await Promise.all(inspectionPhotos.map((photo) => preparePhotoForDocx(photo)))
      appendPhotoAppendix(xmlDoc, relationshipsDoc, contentTypesDoc, preparedPhotos)
      preparedPhotos.forEach((photo, index) => {
        const extension = SUPPORTED_PHOTO_TYPES[photo.file.type]
        zip.file(`word/media/photo-${index + 1}.${extension}`, photo.buffer)
      })
    }

    const updatedXml = new XMLSerializer().serializeToString(xmlDoc)
    const updatedRelationshipsXml = new XMLSerializer().serializeToString(relationshipsDoc)
    const updatedContentTypesXml = new XMLSerializer().serializeToString(contentTypesDoc)
    zip.file('word/document.xml', updatedXml)
    zip.file('word/_rels/document.xml.rels', updatedRelationshipsXml)
    zip.file('[Content_Types].xml', updatedContentTypesXml)

    return zip.generateAsync({
      type: 'blob',
      mimeType: DOCX_MIME_TYPE,
    })
  }

  async function buildDocxBlob() {
    return buildDocxBlobForInspection(data, photos)
  }

  function downloadBlob(blob, fileName) {
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')

    link.href = url
    link.download = fileName
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
    URL.revokeObjectURL(url)
  }

  async function exportAsDocx() {
    const blob = await buildDocxBlob()
    const fileName = getExportFileName(data)
    const archiveFile = new File([blob], fileName, { type: DOCX_MIME_TYPE })
    const record = createArchiveRecord({
      data,
      photos,
      sourceType: 'app',
      storedFileName: fileName,
    })

    downloadBlob(blob, fileName)
    await saveArchiveFile(record.id, archiveFile)
    setArchiveRecords((current) => [record, ...current])
    setImportMessage('La fiche generee a ete ajoutee a l historique local avec son fichier Word.')
  }

  async function exportStatisticsAsDocx() {
    const snapshot = buildStatisticsSnapshot(archiveRecords)
    const { AlignmentType, Document, HeadingLevel, Packer, Paragraph, TextRun } = await import('docx')

    const children = [
      new Paragraph({
        text: 'Statistiques - Archives formulaires',
        heading: HeadingLevel.TITLE,
        alignment: AlignmentType.CENTER,
        spacing: { after: 240 },
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'Date du rapport : ', bold: true }),
          new TextRun(new Date().toLocaleString('fr-FR')),
        ],
        spacing: { after: 120 },
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'Archives consolidées : ', bold: true }),
          new TextRun(String(snapshot.totalArchives)),
        ],
        spacing: { after: 120 },
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'Formulaires generes : ', bold: true }),
          new TextRun(String(snapshot.genericCount)),
        ],
        spacing: { after: 120 },
      }),
      new Paragraph({
        children: [
          new TextRun({ text: 'Bâtiments faits : ', bold: true }),
          new TextRun(`${snapshot.doneBuildingsCount}/${TOTAL_BUILDINGS_TARGET}`),
        ],
        spacing: { after: 240 },
      }),
      new Paragraph({
        text: 'Batiments par section et par item',
        heading: HeadingLevel.HEADING_1,
        spacing: { after: 140 },
      }),
      ...snapshot.sectionBuildingGroups.flatMap((section) => {
        const sectionHeader = [
          new Paragraph({
            text: `Section ${section.id} - ${section.title}`,
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 220, after: 100 },
          }),
          new Paragraph({
            text: section.description,
            spacing: { after: 100 },
          }),
        ]

        if (section.id === '5') {
          return [
            ...sectionHeader,
            ...section.qualityStats.map((item) => new Paragraph({
              text: `${item.label} : ${item.count}/${item.totalCount} (${item.percentage}%) | Batiments: ${item.buildings.length ? item.buildings.join(', ') : 'Aucun bâtiment'}`,
              bullet: { level: 0 },
            })),
            new Paragraph({
              text: 'Remise a l etat a prevoir - details',
              spacing: { before: 120, after: 80 },
            }),
            ...(section.remiseDetails.length
              ? section.remiseDetails.map((item) => new Paragraph({
                  text: `${item.building} : ${item.remarks || 'Aucune remarque'}`,
                  bullet: { level: 0 },
                }))
              : [new Paragraph({ text: 'Aucun bâtiment en remise à l état à prévoir.' })]),
          ]
        }

        return [
          ...sectionHeader,
          ...section.items.map((item) => new Paragraph({
            text: ['1', '2', '3'].includes(section.id)
              ? `${item.key} - ${item.label} : ${formatResponseStats(item)} | Batiments NON: ${item.buildings.length ? item.buildings.join(', ') : 'Aucun bâtiment'}`
              : section.id === '4'
                ? `${item.key} - ${item.label} : ${formatMaterialStats(item)}`
              : section.id === '6'
                ? `${item.key} - ${item.label} : Coche ${item.count}/${item.totalCount} (${item.percentage}%) | Batiments: ${item.buildings.length ? item.buildings.join(', ') : 'Aucun bâtiment'}`
              : `${item.key} - ${item.label} : ${item.buildings.length ? item.buildings.join(', ') : 'Aucun bâtiment'}`,
            bullet: { level: 0 },
          })),
        ]
      }),
      new Paragraph({
        text: 'Statistiques des formulaires generes',
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 220, after: 120 },
      }),
      ...(snapshot.genericTemplates.length
        ? snapshot.genericTemplates.flatMap((template) => [
            new Paragraph({
              text: `${template.title} (${template.kind}) - ${template.recordCount} archive(s)`,
              heading: HeadingLevel.HEADING_2,
              spacing: { before: 180, after: 100 },
            }),
            ...template.sections.flatMap((section) => [
              new Paragraph({
                text: section.title,
                spacing: { before: 120, after: 80 },
              }),
              ...section.fields.map((field) => new Paragraph({
                text: `${field.label} : ${field.filledCount}/${field.totalCount} renseigne(s) (${field.filledPercentage}%) | ${field.valueStats.map((stat) => `${stat.label} ${stat.count}/${field.totalCount} (${stat.percentage}%)`).join(' ; ')}`,
                bullet: { level: 0 },
              })),
            ]),
          ])
        : [new Paragraph({ text: 'Aucun formulaire genere archive pour le moment.' })]),
      new Paragraph({
        text: 'Remarques regroupées',
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 220, after: 120 },
      }),
      ...(snapshot.groupedRemarks.length
        ? snapshot.groupedRemarks.map((group) => new Paragraph({
            text: `${group.remark} : ${group.buildings.join(', ')}${group.examples?.length ? ` | Formulations: ${group.examples.join(' ; ')}` : ''}`,
            bullet: { level: 0 },
          }))
        : [new Paragraph({ text: 'Aucune remarque archivée pour le moment.' })]),
    ]

    const document = new Document({
      sections: [{ children }],
    })

    const blob = await Packer.toBlob(document)
    downloadBlob(blob, getStatisticsExportFileName())
  }

  return (
    <div className="page-shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">Modele Saclay</p>
          <h1>Fiche de verification de prestation</h1>
          <p className="subtitle">
            Cette interface suit la structure du document Word fourni, puis exporte un fichier
            Word-compatible avec le meme decoupage.
          </p>
        </div>

        <div className="actions">
          <div className="progress-badge">{progress}% complete</div>
          <div className="draft-note">Brouillon enregistre automatiquement sur cet appareil</div>
          <div className="draft-note">Export .docx base sur le modele Word original.</div>
          <button type="button" className="primary-action" onClick={exportAsDocx}>
            Generer le .docx
          </button>
          <button type="button" className="secondary-action" onClick={resetForm}>
            Reinitialiser
          </button>
        </div>
      </header>

      <main className="document">
        <AppTabs activeTab={activeTab} onChange={setActiveTab} />

        {activeTab === 'inspection' ? (
          <>
            <section className="paper">
              <ScreenPageOne data={data} setField={setField} setNestedField={setNestedField} />
              <div className="page-separator" />
              <ScreenPageTwo data={data} setField={setField} setNestedField={setNestedField} toggleImprove={toggleImprove} />
            </section>
            <PhotoSection
              photos={photos}
              photoMessage={photoMessage}
              onAddPhotos={handlePhotoSelection}
              onRemovePhoto={removePhoto}
              onCaptionChange={updatePhotoCaption}
              onNameChange={updatePhotoName}
            />
          </>
        ) : null}

        {activeTab === 'generateur' ? (
          <GenericFormGenerator
            onArchiveRecord={(record) => setArchiveRecords((current) => [record, ...current])}
            onArchiveMessage={setImportMessage}
          />
        ) : null}

        {activeTab === 'historique' ? (
          <HistoryPanel
            records={archiveRecords}
            importMessage={importMessage}
            onImportDocx={handleImportDocx}
            onClearArchive={clearArchive}
            onDownloadDocx={downloadArchivedDocx}
            onDownloadAllDocxZip={downloadAllArchivedDocxAsZip}
          />
        ) : null}

        {activeTab === 'statistiques' ? (
          <StatisticsPanel
            records={archiveRecords}
            onExportStats={exportStatisticsAsDocx}
          />
        ) : null}
      </main>

      <PrintDocument data={data} />
      <ScalianSignature />
    </div>
  )
}

function SectionChoiceTable({ titleNumber, title, rows, data, commentKey, setField }) {
  return (
    <table className="inspection-table">
      <tbody>
        <tr className="section-row">
          <td className="number-col">{titleNumber}</td>
          <td>{title}</td>
          <td className="choice-heading">OUI</td>
          <td className="choice-heading">NON</td>
          <td className="choice-heading">Sans objet</td>
        </tr>
        {rows.map(([key, label, options]) => (
          <tr key={key}>
            <td className="number-col">{key}</td>
            <td>{label}</td>
            {['OUI', 'NON', 'Sans objet'].map((option) => (
              <td key={option} className="center-cell">
                {options.includes(option) ? (
                  <ChoiceControl
                    value={option}
                    currentValue={data[key]}
                    onChange={(value) => setField(key, value)}
                  />
                ) : null}
              </td>
            ))}
          </tr>
        ))}
        <tr>
          <td colSpan="5" className="comments-cell">
            <span className="comments-label">Commentaires :</span>
            <TextArea value={data[commentKey]} onChange={(value) => setField(commentKey, value)} rows={4} />
          </td>
        </tr>
      </tbody>
    </table>
  )
}
