import JSZip from 'jszip'
import {
  DOCX_MIME_TYPE,
  W14_NS,
  W_NS,
} from './constants'

const KNOWN_CHOICE_LABELS = ['OUI', 'NON', 'Sans objet', 'NA', 'N/A']

function getWordChildren(node, localName) {
  return Array.from(node?.children || []).filter(
    (child) => child.namespaceURI === W_NS && child.localName === localName,
  )
}

function getWordDescendants(node, localName) {
  return Array.from(node?.getElementsByTagNameNS(W_NS, localName) || [])
}

function normalizeWordText(value) {
  return String(value || '').replaceAll('\u00a0', ' ').replace(/\s+/g, ' ').trim()
}

function getNodeText(node) {
  return normalizeWordText(
    getWordDescendants(node, 't')
      .map((textNode) => textNode.textContent || '')
      .join(' '),
  )
}

function getTableRows(table) {
  return getWordChildren(table, 'tr')
}

function getRowCells(row) {
  return getWordChildren(row, 'tc')
}

function getCellText(cell) {
  return getNodeText(cell)
}

function getCheckboxes(node) {
  return getWordDescendants(node, 'sdt').filter((sdt) => sdt.getElementsByTagNameNS(W14_NS, 'checkbox').length)
}

function hasCheckbox(node) {
  return getCheckboxes(node).length > 0
}

function normalizeComparableText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim()
}

function looksLikeChoiceLabel(value) {
  const normalized = normalizeComparableText(value)
  return KNOWN_CHOICE_LABELS.some((label) => normalizeComparableText(label) === normalized)
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

  const doc = paragraph.ownerDocument
  const lines = String(value || '').split('\n')
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
  const checkedNode = sdt?.getElementsByTagNameNS(W14_NS, 'checked')[0]
  if (checkedNode) setAttributeWithFallback(checkedNode, W14_NS, 'w14:val', checked ? '1' : '0')

  const textNode = sdt?.getElementsByTagNameNS(W_NS, 't')[0]
  if (textNode) textNode.textContent = checked ? '\u2612' : '\u2610'
}

function setAttributeWithFallback(node, namespace, name, value) {
  try {
    node.setAttributeNS(namespace, name, value)
  } catch {
    node.setAttribute(name, value)
  }
}

function cleanFieldLabel(value) {
  return normalizeWordText(value)
    .replace(/:+$/g, '')
    .replace(/\s+/g, ' ')
}

function makeFieldId(parts) {
  return parts.join('-')
}

function getHeaderOptions(rows, rowIndex) {
  const current = getRowCells(rows[rowIndex])
  const previous = rowIndex > 0 ? getRowCells(rows[rowIndex - 1]) : []
  const candidates = [current, previous]

  return candidates.reduce((options, cells) => {
    cells.forEach((cell, cellIndex) => {
      const text = cleanFieldLabel(getCellText(cell))
      if (looksLikeChoiceLabel(text)) options[cellIndex] = text
    })
    return options
  }, {})
}

function getLabelFromRow(cells, checkboxCellIndexes) {
  const firstChoiceIndex = Math.min(...checkboxCellIndexes)
  const labelCells = cells
    .slice(0, firstChoiceIndex)
    .map(getCellText)
    .map(cleanFieldLabel)
    .filter((text) => text && !looksLikeChoiceLabel(text))

  return labelCells.at(-1) || labelCells[0] || `Question ${firstChoiceIndex}`
}

function getTableTitle(rows, rowIndex, tableIndex) {
  for (let index = rowIndex - 1; index >= 0; index -= 1) {
    const text = getCellText(rows[index])
    if (text && !KNOWN_CHOICE_LABELS.some((label) => normalizeComparableText(text).includes(normalizeComparableText(label)))) {
      return cleanFieldLabel(text)
    }
  }

  return `Table ${tableIndex + 1}`
}

function inferChoiceField(tableIndex, rowIndex, rows) {
  const cells = getRowCells(rows[rowIndex])
  const checkboxCellIndexes = cells
    .map((cell, cellIndex) => (hasCheckbox(cell) ? cellIndex : null))
    .filter((cellIndex) => cellIndex !== null)

  if (!checkboxCellIndexes.length) return null

  const headerOptions = getHeaderOptions(rows, rowIndex)
  const options = checkboxCellIndexes.map((cellIndex, optionIndex) => ({
    label: headerOptions[cellIndex] || cleanFieldLabel(getCellText(cells[cellIndex])) || `Option ${optionIndex + 1}`,
    locator: {
      tableIndex,
      rowIndex,
      cellIndex,
      checkboxIndex: 0,
    },
  }))

  return {
    id: makeFieldId(['choice', tableIndex, rowIndex]),
    type: checkboxCellIndexes.length === 1 ? 'checkbox' : 'choice',
    title: getTableTitle(rows, rowIndex, tableIndex),
    label: getLabelFromRow(cells, checkboxCellIndexes),
    options,
  }
}

function inferTextFields(tableIndex, rowIndex, rows) {
  const cells = getRowCells(rows[rowIndex])
  const fields = []

  cells.forEach((cell, cellIndex) => {
    if (hasCheckbox(cell) || getCellText(cell)) return

    const previousText = cleanFieldLabel(getCellText(cells[cellIndex - 1]))
    const nextText = cleanFieldLabel(getCellText(cells[cellIndex + 1]))
    const label = previousText || nextText

    if (!label || looksLikeChoiceLabel(label)) return

    fields.push({
      id: makeFieldId(['text', tableIndex, rowIndex, cellIndex]),
      type: 'text',
      title: getTableTitle(rows, rowIndex, tableIndex),
      label,
      locator: {
        tableIndex,
        rowIndex,
        cellIndex,
      },
    })
  })

  return fields
}

function dedupeFields(fields) {
  const seen = new Set()
  return fields.filter((field) => {
    if (seen.has(field.id)) return false
    seen.add(field.id)
    return true
  })
}

function groupFieldsByTitle(fields) {
  return fields.reduce((sections, field) => {
    const existing = sections.find((section) => section.title === field.title)
    if (existing) {
      existing.fields.push(field)
      return sections
    }

    sections.push({
      id: `section-${sections.length + 1}`,
      title: field.title,
      fields: [field],
    })
    return sections
  }, [])
}

export async function parseGenericDocxTemplate(file) {
  const buffer = await file.arrayBuffer()
  const zip = await JSZip.loadAsync(buffer)
  const documentXml = await zip.file('word/document.xml')?.async('string')

  if (!documentXml) {
    throw new Error('Le fichier Word ne contient pas de document exploitable.')
  }

  const xmlDoc = new DOMParser().parseFromString(documentXml, 'application/xml')
  const tables = getWordDescendants(xmlDoc, 'tbl')
  const fields = dedupeFields(tables.flatMap((table, tableIndex) => {
    const rows = getTableRows(table)
    return rows.flatMap((_, rowIndex) => {
      const choiceField = inferChoiceField(tableIndex, rowIndex, rows)
      if (choiceField) return [choiceField]
      return inferTextFields(tableIndex, rowIndex, rows)
    })
  }))

  return {
    kind: 'docx',
    fileName: file.name,
    fieldCount: fields.length,
    sections: groupFieldsByTitle(fields),
  }
}

export function createInitialGenericValues(schema) {
  return Object.fromEntries(
    schema.sections.flatMap((section) => section.fields.map((field) => [field.id, field.type === 'checkbox' ? false : ''])),
  )
}

export async function buildGenericDocxBlob(templateFile, schema, values) {
  const buffer = await templateFile.arrayBuffer()
  const zip = await JSZip.loadAsync(buffer)
  const documentXml = await zip.file('word/document.xml')?.async('string')

  if (!documentXml) {
    throw new Error('Le fichier Word ne contient pas de document exploitable.')
  }

  const xmlDoc = new DOMParser().parseFromString(documentXml, 'application/xml')
  const tables = getWordDescendants(xmlDoc, 'tbl')

  schema.sections.flatMap((section) => section.fields).forEach((field) => {
    if (field.type === 'text') {
      const row = getTableRows(tables[field.locator.tableIndex])?.[field.locator.rowIndex]
      const cell = getRowCells(row)?.[field.locator.cellIndex]
      if (cell) setCellText(cell, values[field.id])
      return
    }

    field.options.forEach((option) => {
      const row = getTableRows(tables[option.locator.tableIndex])?.[option.locator.rowIndex]
      const cell = getRowCells(row)?.[option.locator.cellIndex]
      const checkbox = getCheckboxes(cell)[option.locator.checkboxIndex]
      const checked = field.type === 'checkbox'
        ? Boolean(values[field.id])
        : values[field.id] === option.label
      setCheckboxState(checkbox, checked)
    })
  })

  zip.file('word/document.xml', new XMLSerializer().serializeToString(xmlDoc))

  return zip.generateAsync({
    type: 'blob',
    mimeType: DOCX_MIME_TYPE,
  })
}
