import {
  PDFCheckBox,
  PDFDocument,
  PDFDropdown,
  PDFOptionList,
  PDFRadioGroup,
  StandardFonts,
  PDFTextField,
} from 'pdf-lib'

const PDF_MIME_TYPE = 'application/pdf'
const FLAT_PDF_SECTION_TITLE = 'Champs detectes dans le PDF plat'
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

  if (schema.mode === 'flat') {
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica)
    const boldFont = await pdfDoc.embedFont(StandardFonts.HelveticaBold)
    const fields = schema.sections.flatMap((section) => section.fields)
    let page = pdfDoc.addPage()
    let y = page.getHeight() - 56

    page.drawText('Reponses saisies depuis l app', { x: 48, y, size: 15, font: boldFont })
    y -= 28
    page.drawText(schema.fileName || 'PDF plat', { x: 48, y, size: 10, font })
    y -= 28

    fields.forEach((field) => {
      const value = String(values[field.id] || '').trim()
      if (!value) return

      if (y < 72) {
        page = pdfDoc.addPage()
        y = page.getHeight() - 56
      }

      page.drawText(`${field.label}:`, { x: 48, y, size: 10, font: boldFont })
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
