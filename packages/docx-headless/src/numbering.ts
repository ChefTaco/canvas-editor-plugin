import type { Element, Node } from '@xmldom/xmldom'
import { parseDocumentXml } from './packageValidation'

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
type Loss = (code: string, message: string) => void
const isElement = (node: Node): node is Element => node.nodeType === 1
const children = (node: Node, name?: string): Element[] =>
  Array.from(node.childNodes).filter(isElement).filter(element =>
    element.namespaceURI === w &&
    (name === undefined || element.localName === name))
const child = (node: Node, name: string): Element | undefined => children(node, name)[0]
const val = (node: Element | undefined, name = 'val'): string | null => node?.getAttributeNS(w, name) ?? null
function byId(elements: Element[], attribute: string): Map<string | null, Element[]> {
  const groups = new Map<string | null, Element[]>()
  for (const element of elements) {
    const id = val(element, attribute)
    const group = groups.get(id) ?? []
    group.push(element)
    groups.set(id, group)
  }
  return groups
}

/** The exchange model stores only a flat bullet/number flag, not Word list
 * identities or arbitrary level definitions. Audit definitions actually used
 * by document/stories/styles/review; dormant numbering parts do not affect
 * rendered content and remain in the immutable source. */
export function auditNumbering(files: Map<string, Buffer>, lose: Loss): void {
  const references = new Map<string, Set<string>>()
  for (const [name, bytes] of files) {
    if (!name.startsWith('word/') || !name.endsWith('.xml') ||
      name === 'word/numbering.xml') continue
    const part = parseDocumentXml(bytes)
    for (const properties of Array.from(part.getElementsByTagNameNS(w, 'numPr'))) {
      const id = val(child(properties, 'numId'))
      if (!id || id === '0') {
        lose('SOURCE_NUMBERING_REFERENCE', 'Missing or disabled numbering references cannot be retained as independent Word numbering semantics.')
        continue
      }
      const owners = references.get(id) ?? new Set<string>()
      owners.add(name)
      references.set(id, owners)
      if (children(properties).some(property =>
        !['ilvl', 'numId'].includes(property.localName ?? ''))) {
        lose('SOURCE_NUMBERING_PROPERTIES', 'Additional numbering properties or revisions remain only in the original.')
      }
      let ancestor: Node | null = properties.parentNode
      while (ancestor) {
        if (ancestor.namespaceURI === w && ancestor.localName === 'style') {
          lose('SOURCE_NUMBERING_STYLE', 'List identity inherited from a named style is not retained by flat list conversion.')
          break
        }
        ancestor = ancestor.parentNode
      }
    }
    // A repeated Word list resumes its counter across intervening prose. The
    // flat exporter instead creates a fresh list at each contiguous group.
    const seen = new Set<string>()
    let previous: string | null = null
    for (const paragraph of Array.from(part.getElementsByTagNameNS(w, 'p'))) {
      const properties = child(paragraph, 'pPr')
      const numbering = properties ? child(properties, 'numPr') : undefined
      const id = numbering ? val(child(numbering, 'numId')) : null
      if (id && id !== '0' && seen.has(id) && id !== previous) {
        lose('SOURCE_NUMBERING_CONTINUATION', 'A Word list resumes after other content; flat list conversion cannot retain its counter continuation.')
      }
      if (id) seen.add(id)
      previous = id
    }
  }
  if (!references.size) return
  if (references.size > 1 ||
    Array.from(references.values()).some(owners => owners.size > 1)) {
    lose('SOURCE_NUMBERING_IDENTITY', 'Independent or cross-story Word list identities cannot be reconstructed from flat list flags.')
  }
  const bytes = files.get('word/numbering.xml')
  if (!bytes) {
    lose('SOURCE_NUMBERING_REFERENCE', 'Referenced Word numbering definitions are missing.')
    return
  }
  const root = parseDocumentXml(bytes)
  const nums = byId(Array.from(root.getElementsByTagNameNS(w, 'num')), 'numId')
  const abstracts = byId(Array.from(root.getElementsByTagNameNS(w, 'abstractNum')), 'abstractNumId')
  for (const id of references.keys()) {
    const instances = nums.get(id) ?? []
    const instance = instances[0]
    const abstractId = instance ? val(child(instance, 'abstractNumId')) : null
    const definitions = abstracts.get(abstractId) ?? []
    const definition = definitions[0]
    if (instances.length !== 1 || !instance || definitions.length !== 1 || !definition || abstractId === null) {
      lose('SOURCE_NUMBERING_REFERENCE', 'A referenced numbering instance or abstract definition is missing or ambiguous.')
      continue
    }
    if (children(instance, 'lvlOverride').length) {
      lose('SOURCE_NUMBERING_OVERRIDE', 'Numbering level/start overrides remain only in the original.')
    }
    const levels = children(definition, 'lvl')
    const level = levels.find(item => val(item, 'ilvl') === '0')
    const multilevel = val(child(definition, 'multiLevelType'))
    if (levels.length !== 1 || !level ||
      multilevel !== null && multilevel !== 'singleLevel') {
      lose('SOURCE_NUMBERING_LEVELS', 'Multilevel Word numbering definitions cannot be retained by flat list conversion.')
    }
    if (!level) continue
    const format = val(child(level, 'numFmt'))
    if (format !== 'decimal' && format !== 'bullet') {
      lose('SOURCE_NUMBERING_FORMAT', 'Only default decimal and bullet list formats are represented; custom formats remain only in the original.')
    }
    if (val(child(level, 'start')) !== '1') {
      lose('SOURCE_NUMBERING_START', 'Custom or unspecified numbering starts are not preserved by default-start flat lists.')
    }
    if (val(child(level, 'lvlText')) !== (format === 'bullet' ? '•' : '%1.')) {
      lose('SOURCE_NUMBERING_LABEL', 'Custom list labels, bullet glyphs or punctuation remain only in the original.')
    }
    const properties = child(level, 'pPr')
    const indent = properties ? child(properties, 'ind') : undefined
    const customIndent = indent && Array.from(indent.attributes).some(attribute =>
      attribute.namespaceURI !== 'http://www.w3.org/2000/xmlns/' &&
      (attribute.namespaceURI !== w ||
        !['left', 'hanging'].includes(attribute.localName ?? '') ||
        attribute.value !== (attribute.localName === 'left' ? '720' : '360')))
    const unsupportedLevel = children(level).some(property =>
      !['start', 'numFmt', 'lvlText', 'lvlJc', 'pPr'].includes(property.localName ?? ''))
    const unsupportedDefinition = children(definition).some(property =>
      !['multiLevelType', 'lvl', 'nsid', 'tmpl'].includes(property.localName ?? ''))
    const unsupportedParagraph = properties && children(properties).some(property =>
      !['tabs', 'ind'].includes(property.localName ?? ''))
    const alignment = val(child(level, 'lvlJc'))
    if (unsupportedLevel || unsupportedDefinition || unsupportedParagraph ||
      customIndent || alignment !== null && alignment !== 'left') {
      lose('SOURCE_NUMBERING_PROPERTIES', 'Numbering-level fonts, positioning, restart rules, links or additional properties remain only in the original.')
    }
  }
}
