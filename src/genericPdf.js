import {
  PDFCheckBox,
  PDFDocument,
  PDFDropdown,
  PDFOptionList,
  PDFRadioGroup,
  PDFTextField,
} from 'pdf-lib'

const PDF_MIME_TYPE = 'application/pdf'

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

export async function parseGenericPdfTemplate(file) {
  const buffer = await file.arrayBuffer()
  const pdfDoc = await PDFDocument.load(buffer)
  const form = pdfDoc.getForm()
  const fields = form.getFields().map(createPdfField).filter(Boolean)

  return {
    kind: 'pdf',
    fileName: file.name,
    fieldCount: fields.length,
    sections: groupFieldsByTitle(fields),
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
