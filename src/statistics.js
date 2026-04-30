import {
  IMPROVEMENTS,
  QUALITY_OPTIONS,
  SECTION_1,
  SECTION_2,
  SECTION_3,
  SECTION_4,
} from './constants'

function getRecordBuildingName(record) {
  const buildingName = String(record.building || '').trim()
  return buildingName || 'Batiment non renseigne'
}

function sortBuildings(buildings) {
  return [...new Set(buildings)]
    .filter(Boolean)
    .sort((left, right) => left.localeCompare(right, 'fr', { numeric: true, sensitivity: 'base' }))
}

function normalizeComparableText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[’']/g, "'")
    .toLowerCase()
    .trim()
}

const REMARK_STOP_WORDS = new Set([
  'avec',
  'aux',
  'ce',
  'ces',
  'dans',
  'des',
  'est',
  'les',
  'non',
  'pas',
  'par',
  'pour',
  'que',
  'qui',
  'sur',
  'une',
])

const REMARK_WORD_ALIASES = {
  cafet: 'cafeteria',
  cafe: 'cafeteria',
  calcaire: 'detartrage',
  corbeille: 'poubelle',
  corbeilles: 'poubelle',
  couloir: 'circulation',
  couloirs: 'circulation',
  dechet: 'poubelle',
  dechets: 'poubelle',
  detartrant: 'detartrage',
  detartre: 'detartrage',
  lavabo: 'sanitaire',
  lavabos: 'sanitaire',
  odeurs: 'odeur',
  papier: 'consommable',
  papiers: 'consommable',
  plein: 'poubelle',
  pleine: 'poubelle',
  pleines: 'poubelle',
  poussiereux: 'poussiere',
  poubelles: 'poubelle',
  sanitaires: 'sanitaire',
  savon: 'consommable',
  sols: 'sol',
  toilette: 'sanitaire',
  toilettes: 'sanitaire',
  traces: 'trace',
  vitre: 'vitre',
  vitres: 'vitre',
  wc: 'sanitaire',
}

const REMARK_THEMES = [
  ['Sanitaires / WC', ['sanitaire', 'wc', 'toilette', 'lavabo', 'urinoir']],
  ['Poubelles / déchets', ['poubelle', 'corbeille', 'dechet', 'ordure']],
  ['Sols / traces', ['sol', 'trace', 'tache', 'lavage', 'balayage']],
  ['Poussière / surfaces', ['poussiere', 'surface', 'bureau', 'etagere']],
  ['Vitrerie', ['vitre', 'vitrage', 'fenetre']],
  ['Circulations / couloirs', ['circulation', 'couloir', 'hall']],
  ['Coin cafétéria', ['cafeteria', 'cafet', 'cafe', 'kitchenette', 'cuisine']],
  ['Consommables', ['savon', 'papier', 'essuie', 'consommable']],
  ['Odeurs', ['odeur', 'odorant', 'malodorant']],
  ['Détartrage / calcaire', ['calcaire', 'detartrage', 'detartrant']],
  ['Matériel / chariot', ['materiel', 'chariot', 'aspirateur', 'autolaveuse', 'balai']],
]

function splitRemarkText(value) {
  return String(value || '')
    .split(/\n+|[.;•]+|\s+\|\s+/)
    .map((part) => part.trim())
    .filter(Boolean)
}

function tokenizeRemark(value) {
  return normalizeComparableText(value)
    .replace(/[^a-z0-9' ]+/g, ' ')
    .split(/\s+/)
    .map((word) => word.replace(/^l'|^d'|^qu'/, ''))
    .map((word) => REMARK_WORD_ALIASES[word] || word)
    .filter((word) => word.length > 2 && !REMARK_STOP_WORDS.has(word))
}

function getRemarkThemes(tokens) {
  const tokenSet = new Set(tokens)
  return REMARK_THEMES
    .filter(([, keywords]) => keywords.some((keyword) => tokenSet.has(REMARK_WORD_ALIASES[keyword] || keyword)))
    .map(([label]) => label)
}

function jaccardSimilarity(leftTokens, rightTokens) {
  const left = new Set(leftTokens)
  const right = new Set(rightTokens)
  const intersection = [...left].filter((token) => right.has(token)).length
  const union = new Set([...left, ...right]).size

  return union ? intersection / union : 0
}

function hasEnoughSharedMeaning(left, right) {
  if (left.normalized === right.normalized) return true
  if (left.themes.some((theme) => right.themes.includes(theme))) return true

  const sharedTokens = left.tokens.filter((token) => right.tokens.includes(token))
  const shortestTokenCount = Math.min(left.tokens.length, right.tokens.length)
  return jaccardSimilarity(left.tokens, right.tokens) >= 0.45
    || (shortestTokenCount <= 4 && sharedTokens.length >= 2)
}

function createRemarkEntry(record, remark) {
  const tokens = tokenizeRemark(remark)
  return {
    building: getRecordBuildingName(record),
    normalized: normalizeComparableText(remark),
    remark,
    themes: getRemarkThemes(tokens),
    tokens,
  }
}

function isYesText(value) {
  const normalizedValue = normalizeComparableText(value)
  return normalizedValue === 'oui' || normalizedValue.startsWith('oui ')
}

function buildResponseItemEntries(definitions, records, getResponseValue, eligibleRecords = records) {
  return definitions.map(([key, label, options = ['OUI', 'NON']]) => {
    const issueBuildings = sortBuildings(
      eligibleRecords
        .filter((record) => getResponseValue(record, key) === 'NON')
        .map(getRecordBuildingName),
    )
    const responseStats = options.map((option) => {
      const matchingRecords = eligibleRecords.filter((record) => getResponseValue(record, key) === option)
      const count = matchingRecords.length

      return {
        label: option,
        count,
        percentage: eligibleRecords.length ? Math.round((count / eligibleRecords.length) * 100) : 0,
        buildings: sortBuildings(matchingRecords.map(getRecordBuildingName)),
      }
    })

    return {
      key,
      label,
      buildings: issueBuildings,
      responseStats,
      totalCount: eligibleRecords.length,
    }
  })
}

function getPercentage(count, total) {
  return total ? Math.round((count / total) * 100) : 0
}

function buildMaterialItemEntries(records) {
  const materialStates = [
    ['Neuf', (state) => normalizeComparableText(state) === 'bon etat' || normalizeComparableText(state) === 'neuf'],
    ["Etat d'usage", (state) => normalizeComparableText(state) === "etat d'usage"],
    ['Vetuste', (state) => normalizeComparableText(state) === 'vetuste'],
  ]

  return SECTION_4.map(([key, label, withState]) => {
    const presentRecords = records.filter((record) => record.data?.[key]?.status === 'OUI')
    const presentCount = presentRecords.length

    return {
      key,
      label,
      withState,
      presentCount,
      presentPercentage: getPercentage(presentCount, records.length),
      totalCount: records.length,
      materialStats: withState
        ? materialStates.map(([stateLabel, matchesState]) => {
            const matchingRecords = presentRecords.filter((record) => matchesState(record.data?.[key]?.state))
            const count = matchingRecords.length

            return {
              label: stateLabel,
              count,
              percentage: getPercentage(count, presentCount),
              buildings: sortBuildings(matchingRecords.map(getRecordBuildingName)),
            }
          })
        : [],
    }
  })
}

function buildRecordRemarks(record) {
  const remarks = [
    String(record.data?.qualityComments || '').trim(),
    String(record.data?.improveComments || '').trim(),
    String(record.data?.otherRemarks || '').trim(),
  ].filter(Boolean)

  return remarks.join(' | ')
}

function buildRecordRemarkEntries(record) {
  return [
    record.data?.qualityComments,
    record.data?.improveComments,
    record.data?.otherRemarks,
  ]
    .flatMap(splitRemarkText)
    .map((remark) => createRemarkEntry(record, remark))
    .filter((entry) => entry.tokens.length || entry.themes.length)
}

function getGenericRecords(records) {
  return records.filter((record) => record.sourceType === 'generated_form')
}

function getSaclayRecords(records) {
  return records.filter((record) => record.sourceType !== 'generated_form')
}

function getRecordDisplayName(record) {
  return record.sourceName || record.storedFileName || getRecordBuildingName(record)
}

function getGenericValueLabel(value) {
  if (Array.isArray(value)) return value.length ? value.join(', ') : 'Vide'
  if (typeof value === 'boolean') return value ? 'Coche' : 'Non coche'
  return String(value || '').trim() || 'Vide'
}

function buildGenericFieldStats(field, records) {
  const totalCount = records.length
  const values = records.map((record) => record.values?.[field.id])
  const filledCount = values.filter((value) => (
    Array.isArray(value) ? value.length : typeof value === 'boolean' ? value : Boolean(String(value || '').trim())
  )).length

  if (field.type === 'choice') {
    return {
      id: field.id,
      label: field.label,
      type: field.type,
      totalCount,
      filledCount,
      filledPercentage: getPercentage(filledCount, totalCount),
      valueStats: field.options.map((option) => {
        const matchingRecords = records.filter((record) => record.values?.[field.id] === option.label)
        const count = matchingRecords.length

        return {
          label: option.label,
          count,
          percentage: getPercentage(count, totalCount),
          records: matchingRecords.map(getRecordDisplayName),
        }
      }),
    }
  }

  if (field.type === 'checkbox') {
    const checkedRecords = records.filter((record) => Boolean(record.values?.[field.id]))
    const count = checkedRecords.length

    return {
      id: field.id,
      label: field.label,
      type: field.type,
      totalCount,
      filledCount: count,
      filledPercentage: getPercentage(count, totalCount),
      valueStats: [
        {
          label: 'Coche',
          count,
          percentage: getPercentage(count, totalCount),
          records: checkedRecords.map(getRecordDisplayName),
        },
      ],
    }
  }

  const valueStats = Object.entries(values.reduce((counts, value, index) => {
    const label = getGenericValueLabel(value)
    if (!counts[label]) counts[label] = { count: 0, records: [] }
    counts[label].count += 1
    counts[label].records.push(getRecordDisplayName(records[index]))
    return counts
  }, {}))
    .map(([label, stat]) => ({
      label,
      count: stat.count,
      percentage: getPercentage(stat.count, totalCount),
      records: sortBuildings(stat.records),
    }))
    .sort((left, right) => right.count - left.count || left.label.localeCompare(right.label, 'fr', { sensitivity: 'base' }))
    .slice(0, 5)

  return {
    id: field.id,
    label: field.label,
    type: field.type,
    totalCount,
    filledCount,
    filledPercentage: getPercentage(filledCount, totalCount),
    valueStats,
  }
}

function buildGenericTemplateStatistics(records) {
  const groups = records.reduce((templates, record) => {
    const key = record.templateId || record.sourceName || record.storedFileName || 'Modele genere'
    if (!templates[key]) {
      templates[key] = {
        key,
        title: record.sourceName || 'Modele genere',
        kind: record.formKind || 'form',
        schema: record.schema,
        records: [],
      }
    }
    templates[key].records.push(record)
    return templates
  }, {})

  return Object.values(groups).map((template) => ({
    key: template.key,
    title: template.title,
    kind: template.kind,
    recordCount: template.records.length,
    sections: (template.schema?.sections || []).map((section) => ({
      id: section.id,
      title: section.title,
      fields: section.fields.map((field) => buildGenericFieldStats(field, template.records)),
    })),
  }))
}

function groupRecordsBySimilarRemark(records) {
  const groups = []

  records.forEach((record) => {
    buildRecordRemarkEntries(record).forEach((entry) => {
      const existingGroup = groups.find((group) => hasEnoughSharedMeaning(entry, group.reference))

      if (existingGroup) {
        existingGroup.buildings.push(entry.building)
        if (!existingGroup.remarks.includes(entry.remark)) existingGroup.remarks.push(entry.remark)
        existingGroup.tokens = [...new Set([...existingGroup.tokens, ...entry.tokens])]
        existingGroup.reference = {
          ...existingGroup.reference,
          tokens: existingGroup.tokens,
        }
        return
      }

      groups.push({
        reference: entry,
        remark: entry.themes[0] || entry.remark,
        remarks: [entry.remark],
        tokens: entry.tokens,
        buildings: [entry.building],
      })
    })
  })

  return groups
    .map((group) => ({
      remark: group.remark,
      examples: group.remarks.slice(0, 4),
      buildings: sortBuildings(group.buildings),
    }))
    .sort((left, right) => {
      if (right.buildings.length !== left.buildings.length) return right.buildings.length - left.buildings.length
      return left.remark.localeCompare(right.remark, 'fr', { sensitivity: 'base' })
    })
}

export function buildStatisticsSnapshot(records) {
  const saclayRecords = getSaclayRecords(records)
  const genericRecords = getGenericRecords(records)
  const importedCount = saclayRecords.filter((record) => record.sourceType === 'manual_docx').length
  const appCount = saclayRecords.filter((record) => record.sourceType === 'app').length
  const uniqueBuildings = sortBuildings(
    saclayRecords
      .map((record) => String(record.building || '').trim())
      .filter(Boolean),
  )
  const doneBuildingsCount = uniqueBuildings.length
  const qualityStats = QUALITY_OPTIONS.map(([label, symbol, color]) => {
    const matchingRecords = saclayRecords.filter((record) => record.data?.quality === label)
    const count = matchingRecords.length

    return {
      label,
      symbol,
      color,
      count,
      percentage: getPercentage(count, saclayRecords.length),
      totalCount: saclayRecords.length,
      buildings: sortBuildings(matchingRecords.map(getRecordBuildingName)),
    }
  })
  const remiseRecords = saclayRecords.filter((record) => record.data?.quality === 'Remise en état à prévoir')
  const remiseDetails = remiseRecords
    .map((record) => ({
      building: getRecordBuildingName(record),
      remarks: buildRecordRemarks(record),
    }))
    .sort((left, right) => left.building.localeCompare(right.building, 'fr', { numeric: true, sensitivity: 'base' }))
  const groupedRemarks = groupRecordsBySimilarRemark(saclayRecords)
  const recordsWithAtalianPresent = saclayRecords.filter((record) => isYesText(record.data?.intervenantPresent))
  const sectionBuildingGroups = [
    {
      id: '0',
      title: 'Presence des intervenants ATALIAN',
      description: 'Batiments dans lesquels les intervenants ATALIAN sont presents (Oui).',
      items: [
        {
          key: '0.1',
          label: 'Intervenant(s) ATALIAN present (Oui)',
          buildings: sortBuildings(recordsWithAtalianPresent.map(getRecordBuildingName)),
        },
      ],
    },
    {
      id: '1',
      title: 'Formations et habilitations',
      description: 'Pourcentage des reponses OUI, NON et Sans objet calcule uniquement si Intervenant(s) ATALIAN present = Oui. Batiments listes si la reponse est NON.',
      items: buildResponseItemEntries(
        SECTION_1,
        saclayRecords,
        (record, key) => record.data?.[key],
        recordsWithAtalianPresent,
      ),
    },
    {
      id: '2',
      title: 'Equipements des intervenants',
      description: 'Pourcentage des reponses OUI, NON et Sans objet. Batiments listes si la reponse est NON.',
      items: buildResponseItemEntries(
        SECTION_2,
        saclayRecords,
        (record, key) => record.data?.[key],
      ),
    },
    {
      id: '3',
      title: 'Produits de nettoyage',
      description: 'Pourcentage des reponses OUI et NON. Batiments listes si la reponse est NON.',
      items: buildResponseItemEntries(
        SECTION_3,
        saclayRecords,
        (record, key) => record.data?.[key]?.status,
      ),
    },
    {
      id: '4',
      title: 'Materiels et documents',
      description: 'Pourcentage des items presents (OUI), avec repartition de l etat du materiel quand applicable.',
      items: buildMaterialItemEntries(saclayRecords),
    },
    {
      id: '5',
      title: 'Etat general du batiment',
      description: 'Pourcentage des batiments par niveau de qualite, avec les remarques detaillees pour les remises a l etat a prevoir.',
      qualityStats,
      remiseDetails,
    },
    {
      id: '6',
      title: 'Points d amelioration',
      description: 'Pourcentage des batiments pour lesquels chaque point d amelioration est coche.',
      items: IMPROVEMENTS.map((label, index) => {
        const matchingRecords = saclayRecords.filter((record) => (
          Array.isArray(record.data?.improve) && record.data.improve.includes(label)
        ))
        const count = matchingRecords.length

        return {
          key: `6.${index + 1}`,
          label,
          count,
          percentage: getPercentage(count, saclayRecords.length),
          totalCount: saclayRecords.length,
          buildings: sortBuildings(matchingRecords.map(getRecordBuildingName)),
        }
      }),
    },
  ]

  return {
    importedCount,
    appCount,
    doneBuildingsCount,
    genericCount: genericRecords.length,
    saclayCount: saclayRecords.length,
    totalArchives: records.length,
    sectionBuildingGroups,
    groupedRemarks,
    genericTemplates: buildGenericTemplateStatistics(genericRecords),
  }
}
