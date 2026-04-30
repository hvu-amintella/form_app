import {
  PDFCheckBox,
  PDFDocument,
  PDFDropdown,
  PDFOptionList,
  PDFRadioGroup,
  StandardFonts,
  PDFTextField,
} from 'pdf-lib'
import {
  GENERAL_FIELDS,
  IMPROVEMENTS,
  QUALITY_OPTIONS,
  SECTION_1,
  SECTION_2,
  SECTION_3,
  SECTION_4,
} from './constants'

const PDF_MIME_TYPE = 'application/pdf'
const FLAT_PDF_SECTION_TITLE = 'Champs detectes dans le PDF plat'
const SCANNED_PDF_SECTION_PREFIX = 'PDF scanne - '
const TEXT_DECODER = new TextDecoder('latin1')
const UTF8_DECODER = new TextDecoder('utf-8')
const MAX_FLAT_FIELDS = 80
const KNOWN_CHOICE_LABELS = ['OUI', 'NON', 'Sans objet', 'NA', 'N/A']

function getPdfFieldType(field) {
  if (field instanceof PDFTextField) return 'text'
  if (field instanceof PDFCheckBox) return 'checkbox'
  if (field instanceof PDFRadioGroup || field instanceof PDFDropdown || field instanceof PDFOptionList) return 'choice'
  return 'unsupported'
}

function getPdfFieldOptions(field) {
  if (typeof field.getOptions !== 'function') return []
  return field.getOptions().map((label) => ({ label }))
}

function getPdfFieldLabel(name) {
  return String(name || '')
    .split(/[./]/)
    .filter(Boolean)
    .at(-1) || name
}

function getPdfSectionTitle(name) {
  const parts = String(name || '').split(/[./]/).filter(Boolean)
  return parts.length > 1 ? parts.slice(0, -1).join(' / ') : 'Champs PDF remplissables'
}

function createPdfField(field) {
  const name = field.getName()
  const type = getPdfFieldType(field)

  if (type === 'unsupported') return null

  return {
    id: `pdf-${name}`,
    type,
    title: getPdfSectionTitle(name),
    label: getPdfFieldLabel(name),
    pdfName: name,
    options: type === 'choice' ? getPdfFieldOptions(field) : [],
  }
}

function groupFieldsByTitle(fields) {
  return fields.reduce((sections, field) => {
    const existing = sections.find((section) => section.title === field.title)
    if (existing) {
      existing.fields.push(field)
      return sections
    }

    sections.push({
      id: `pdf-section-${sections.length + 1}`,
      title: field.title,
      fields: [field],
    })
    return sections
  }, [])
}

function normalizePdfText(value) {
  return String(value || '')
    .replaceAll('\u00a0', ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

function normalizeComparableText(value) {
  return normalizePdfText(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
}

function looksLikeChoiceLabel(value) {
  const normalized = normalizeComparableText(value)
  return KNOWN_CHOICE_LABELS.some((label) => normalizeComparableText(label) === normalized)
}

function cleanFlatPdfLabel(value) {
  return normalizePdfText(value)
    .replace(/^[\d\s.)-]+/, '')
    .replace(/[:.\s]+$/g, '')
}

function toPdfDrawText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[’‘]/g, "'")
    .replace(/[“”]/g, '"')
    .replace(/[–—]/g, '-')
    .replace(/[^\x20-\x7e]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

function isUsefulFlatPdfLabel(value) {
  const label = cleanFlatPdfLabel(value)
  const normalized = normalizeComparableText(label)

  if (label.length < 3 || label.length > 90) return false
  if (looksLikeChoiceLabel(label)) return false
  if (/^[\d\s./-]+$/.test(label)) return false
  if (!/[a-zA-ZÀ-ÿ]/.test(label)) return false
  if (normalized.includes('page ') || normalized.includes('fiche de verification')) return false

  return true
}

function decodePdfLiteralString(rawValue) {
  let output = ''

  for (let index = 0; index < rawValue.length; index += 1) {
    const current = rawValue[index]
    if (current !== '\\') {
      output += current
      continue
    }

    const next = rawValue[index + 1]
    if (!next) continue

    if (next === 'n') output += '\n'
    else if (next === 'r') output += '\r'
    else if (next === 't') output += '\t'
    else if (next === 'b') output += '\b'
    else if (next === 'f') output += '\f'
    else if (next === '\r' || next === '\n') {
      if (next === '\r' && rawValue[index + 2] === '\n') index += 1
    } else if (/[0-7]/.test(next)) {
      const octal = rawValue.slice(index + 1, index + 4).match(/^[0-7]{1,3}/)?.[0] || next
      output += String.fromCharCode(parseInt(octal, 8))
      index += octal.length - 1
    } else {
      output += next
    }

    index += 1
  }

  return output
}

function decodePdfHexString(rawValue) {
  const clean = rawValue.replace(/\s+/g, '')
  const bytes = []

  for (let index = 0; index < clean.length; index += 2) {
    bytes.push(parseInt(clean.slice(index, index + 2).padEnd(2, '0'), 16))
  }

  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    let output = ''
    for (let index = 2; index < bytes.length; index += 2) {
      output += String.fromCharCode((bytes[index] << 8) + (bytes[index + 1] || 0))
    }
    return output
  }

  return TEXT_DECODER.decode(new Uint8Array(bytes))
}

async function inflatePdfStream(bytes) {
  if (typeof DecompressionStream === 'undefined') return null

  try {
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate'))
    return new Uint8Array(await new Response(stream).arrayBuffer())
  } catch {
    return null
  }
}

function getRawPdfStreams(bytes) {
  const text = TEXT_DECODER.decode(bytes)
  const streams = []
  let searchIndex = 0

  while (searchIndex < text.length) {
    const streamIndex = text.indexOf('stream', searchIndex)
    if (streamIndex === -1) break

    const dictionaryStart = text.lastIndexOf('<<', streamIndex)
    const dictionaryEnd = text.lastIndexOf('>>', streamIndex)
    const endIndex = text.indexOf('endstream', streamIndex)

    if (dictionaryStart === -1 || dictionaryEnd === -1 || endIndex === -1) break

    let dataStart = streamIndex + 'stream'.length
    if (text[dataStart] === '\r' && text[dataStart + 1] === '\n') dataStart += 2
    else if (text[dataStart] === '\n' || text[dataStart] === '\r') dataStart += 1

    let dataEnd = endIndex
    if (text[dataEnd - 2] === '\r' && text[dataEnd - 1] === '\n') dataEnd -= 2
    else if (text[dataEnd - 1] === '\n' || text[dataEnd - 1] === '\r') dataEnd -= 1

    streams.push({
      dictionary: text.slice(dictionaryStart, dictionaryEnd + 2),
      bytes: bytes.slice(dataStart, dataEnd),
    })

    searchIndex = endIndex + 'endstream'.length
  }

  return streams
}

function extractPdfTextFragments(content) {
  const fragments = []

  content.replace(/\((?:\\.|[^\\)])*\)\s*(?:Tj|'|")/g, (match) => {
    const rawValue = match.slice(1, match.lastIndexOf(')'))
    fragments.push(decodePdfLiteralString(rawValue))
    return match
  })

  content.replace(/<([0-9a-fA-F\s]+)>\s*Tj/g, (match, rawValue) => {
    fragments.push(decodePdfHexString(rawValue))
    return match
  })

  content.replace(/\[((?:.|\n|\r)*?)\]\s*TJ/g, (match, arrayBody) => {
    const parts = []

    arrayBody.replace(/\((?:\\.|[^\\)])*\)|<([0-9a-fA-F\s]+)>/g, (part, hexValue) => {
      parts.push(hexValue ? decodePdfHexString(hexValue) : decodePdfLiteralString(part.slice(1, -1)))
      return part
    })

    if (parts.length) fragments.push(parts.join(''))
    return match
  })

  return fragments.map(normalizePdfText).filter(Boolean)
}

async function extractFlatPdfText(file) {
  const bytes = new Uint8Array(await file.arrayBuffer())
  const streams = getRawPdfStreams(bytes)
  const contents = []

  for (const stream of streams) {
    if (stream.dictionary.includes('/FlateDecode')) {
      const inflated = await inflatePdfStream(stream.bytes)
      if (inflated) contents.push(UTF8_DECODER.decode(inflated))
    } else {
      contents.push(TEXT_DECODER.decode(stream.bytes))
    }
  }

  return contents.flatMap(extractPdfTextFragments)
}

function createFlatPdfFields(textFragments) {
  const candidates = textFragments
    .flatMap((fragment) => fragment.split(/\s{2,}|[|]/g))
    .map(cleanFlatPdfLabel)
    .filter(isUsefulFlatPdfLabel)

  const seen = new Set()
  const fields = []

  for (const label of candidates) {
    const key = normalizeComparableText(label)
    if (seen.has(key)) continue
    seen.add(key)

    fields.push({
      id: `flat-pdf-${fields.length + 1}`,
      type: 'text',
      title: FLAT_PDF_SECTION_TITLE,
      label,
      flatPdf: true,
    })

    if (fields.length >= MAX_FLAT_FIELDS) break
  }

  return fields
}

function createSaclayScannedPdfSections() {
  const createChoiceField = (id, title, label, options) => ({
    id,
    type: 'choice',
    title,
    label,
    options: options.map((option) => ({ label: option })),
    flatPdf: true,
  })

  const createTextField = (id, title, label) => ({
    id,
    type: 'text',
    title,
    label,
    flatPdf: true,
  })

  const createCheckboxField = (id, title, label) => ({
    id,
    type: 'checkbox',
    title,
    label,
    flatPdf: true,
  })

  return [
    {
      id: 'scanned-pdf-general',
      title: `${SCANNED_PDF_SECTION_PREFIX}Informations generales`,
      fields: GENERAL_FIELDS.map(([key, label]) => createTextField(`scanned-general-${key}`, `${SCANNED_PDF_SECTION_PREFIX}Informations generales`, label)),
    },
    {
      id: 'scanned-pdf-section-1',
      title: `${SCANNED_PDF_SECTION_PREFIX}Formations et habilitations`,
      fields: [
        ...SECTION_1.map(([key, label, options]) => createChoiceField(`scanned-section-1-${key}`, `${SCANNED_PDF_SECTION_PREFIX}Formations et habilitations`, label, options)),
        createTextField('scanned-comments-1', `${SCANNED_PDF_SECTION_PREFIX}Formations et habilitations`, 'Commentaires'),
      ],
    },
    {
      id: 'scanned-pdf-section-2',
      title: `${SCANNED_PDF_SECTION_PREFIX}Equipements des intervenants`,
      fields: [
        ...SECTION_2.map(([key, label, options]) => createChoiceField(`scanned-section-2-${key}`, `${SCANNED_PDF_SECTION_PREFIX}Equipements des intervenants`, label, options)),
        createTextField('scanned-comments-2', `${SCANNED_PDF_SECTION_PREFIX}Equipements des intervenants`, 'Commentaires'),
      ],
    },
    {
      id: 'scanned-pdf-section-3',
      title: `${SCANNED_PDF_SECTION_PREFIX}Produits de nettoyage`,
      fields: [
        ...SECTION_3.flatMap(([key, label]) => [
          createChoiceField(`scanned-section-3-${key}-status`, `${SCANNED_PDF_SECTION_PREFIX}Produits de nettoyage`, label, ['OUI', 'NON']),
          createTextField(`scanned-section-3-${key}-name`, `${SCANNED_PDF_SECTION_PREFIX}Produits de nettoyage`, `${label} - nom du produit`),
        ]),
        createTextField('scanned-comments-3', `${SCANNED_PDF_SECTION_PREFIX}Produits de nettoyage`, 'Commentaires'),
      ],
    },
    {
      id: 'scanned-pdf-section-4',
      title: `${SCANNED_PDF_SECTION_PREFIX}Materiels / documents`,
      fields: [
        ...SECTION_4.flatMap(([key, label, withState]) => [
          createChoiceField(`scanned-section-4-${key}-status`, `${SCANNED_PDF_SECTION_PREFIX}Materiels / documents`, label, ['OUI', 'NON']),
          ...(withState ? [createChoiceField(`scanned-section-4-${key}-state`, `${SCANNED_PDF_SECTION_PREFIX}Materiels / documents`, `${label} - etat`, ['Bon etat', 'Etat d usage', 'Vetuste'])] : []),
        ]),
        createTextField('scanned-comments-4', `${SCANNED_PDF_SECTION_PREFIX}Materiels / documents`, 'Commentaires'),
      ],
    },
    {
      id: 'scanned-pdf-section-5',
      title: `${SCANNED_PDF_SECTION_PREFIX}Evaluation qualite`,
      fields: [
        createChoiceField('scanned-quality', `${SCANNED_PDF_SECTION_PREFIX}Evaluation qualite`, 'Niveau de qualite', QUALITY_OPTIONS.map(([label]) => label)),
        createTextField('scanned-quality-comments', `${SCANNED_PDF_SECTION_PREFIX}Evaluation qualite`, 'Commentaires'),
      ],
    },
    {
      id: 'scanned-pdf-section-6',
      title: `${SCANNED_PDF_SECTION_PREFIX}Remise en etat a prevoir`,
      fields: [
        ...IMPROVEMENTS.map((label) => createCheckboxField(`scanned-improve-${normalizeComparableText(label).replace(/[^a-z0-9]+/g, '-')}`, `${SCANNED_PDF_SECTION_PREFIX}Remise en etat a prevoir`, label)),
        createTextField('scanned-improve-comments', `${SCANNED_PDF_SECTION_PREFIX}Remise en etat a prevoir`, 'Commentaires'),
        createTextField('scanned-other-remarks', `${SCANNED_PDF_SECTION_PREFIX}Remise en etat a prevoir`, 'Autres remarques'),
      ],
    },
  ]
}

function createSaclayScannedPdfSchema(fileName) {
  const sections = createSaclayScannedPdfSections()

  return {
    kind: 'pdf',
    fileName,
    fieldCount: sections.reduce((total, section) => total + section.fields.length, 0),
    sections,
    mode: 'scanned-saclay',
  }
}

export async function parseGenericPdfTemplate(file) {
  const buffer = await file.arrayBuffer()
  const pdfDoc = await PDFDocument.load(buffer)
  const form = pdfDoc.getForm()
  const fields = form.getFields().map(createPdfField).filter(Boolean)

  if (fields.length) {
    return {
      kind: 'pdf',
      fileName: file.name,
      fieldCount: fields.length,
      sections: groupFieldsByTitle(fields),
      mode: 'fillable',
    }
  }

  const flatFields = createFlatPdfFields(await extractFlatPdfText(file))

  if (!flatFields.length) {
    return createSaclayScannedPdfSchema(file.name)
  }

  return {
    kind: 'pdf',
    fileName: file.name,
    fieldCount: flatFields.length,
    sections: groupFieldsByTitle(flatFields),
    mode: 'flat',
  }
}

export function createInitialGenericPdfValues(schema) {
  return Object.fromEntries(
    schema.sections.flatMap((section) => section.fields.map((field) => [field.id, field.type === 'checkbox' ? false : ''])),
  )
}

export async function buildGenericPdfBlob(templateFile, schema, values) {
  const buffer = await templateFile.arrayBuffer()
  const pdfDoc = await PDFDocument.load(buffer)

  if (schema.mode === 'flat' || schema.mode === 'scanned-saclay') {
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica)
    const boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold)
    const fields = schema.sections.flatMap((section) => section.fields)
    let page = pdfDoc.addPage()
    let y = page.getHeight() - 56

    page.drawText('Reponses saisies depuis l app', { x: 48, y, size: 15, font: boldFont })
    y -= 28
    page.drawText(toPdfDrawText(schema.fileName || 'PDF plat'), { x: 48, y, size: 10, font })
    y -= 28

    fields.forEach((field) => {
      const rawValue = field.type === 'checkbox'
        ? (values[field.id] ? 'Oui' : '')
        : values[field.id]
      const value = toPdfDrawText(rawValue)
      if (!value) return

      if (y < 72) {
        page = pdfDoc.addPage()
        y = page.getHeight() - 56
      }

      page.drawText(`${toPdfDrawText(field.label)}:`, { x: 48, y, size: 10, font: boldFont })
      y -= 14
      value.match(/.{1,95}(\s|$)|\S+/g)?.forEach((line) => {
        if (y < 72) {
          page = pdfDoc.addPage()
          y = page.getHeight() - 56
        }
        page.drawText(line.trim(), { x: 64, y, size: 10, font })
        y -= 13
      })
      y -= 8
    })

    const bytes = await pdfDoc.save()
    return new Blob([bytes], { type: PDF_MIME_TYPE })
  }

  const form = pdfDoc.getForm()

  schema.sections.flatMap((section) => section.fields).forEach((field) => {
    const pdfField = form.getField(field.pdfName)
    const value = values[field.id]

    if (pdfField instanceof PDFTextField) {
      pdfField.setText(String(value || ''))
      return
    }

    if (pdfField instanceof PDFCheckBox) {
      if (value) pdfField.check()
      else pdfField.uncheck()
      return
    }

    if (pdfField instanceof PDFRadioGroup || pdfField instanceof PDFDropdown || pdfField instanceof PDFOptionList) {
      if (value) pdfField.select(value)
    }
  })

  const bytes = await pdfDoc.save()
  return new Blob([bytes], { type: PDF_MIME_TYPE })
}
