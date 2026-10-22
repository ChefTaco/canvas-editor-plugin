import test from 'node:test'
import assert from 'node:assert/strict'
import { defaultLayout, exportDocx, importDocx, paragraphsOf, validateDocxPackage, type WordDocument } from '../src/index'

void test('explicit heading levels survive named paragraph styles and all review projections', async () => {
  const document: WordDocument = {
    schemaVersion: 1, layout: { ...defaultLayout(), pageNumbers: false },
    blocks: [{ type: 'paragraph', id: 'p', style: { name: 'Custom', heading: 2 }, runs: [{ id: 'r', text: 'A named heading', style: {} }] }],
    changes: [{ id: 'format', kind: 'paragraphFormat', authorName: 'Reviewer', createdAt: '2026-10-22T00:00:00Z', status: 'pending', targetRunIds: [], targetBlockIds: ['p'], beforeParagraphStyle: { name: 'Custom', heading: 1 }, afterParagraphStyle: { name: 'Custom', heading: 2 } }], comments: []
  }
  const definitions = { Custom: { paragraph: { alignment: 'center' as const }, run: { bold: true } } }
  const before = structuredClone(document)
  for (const mode of ['retain-review', 'original', 'final'] as const) {
    const bytes = await exportDocx(document, mode, new Map(), undefined, definitions)
    const imported = await importDocx(bytes)
    assert.deepEqual(imported.losses, [])
    assert.equal(paragraphsOf(imported.document)[0]?.style.heading, mode === 'original' ? 1 : 2)
    if (mode === 'retain-review') {
      const change = imported.document.changes.find(change => change.kind === 'paragraphFormat')
      assert.equal(change?.beforeParagraphStyle?.heading, 1)
      assert.equal(change?.afterParagraphStyle?.heading, 2)
    }
    const xml = (await validateDocxPackage(bytes)).get('word/document.xml')?.toString('utf8') ?? ''
    assert.match(xml, /<w:outlineLvl w:val="[01]"\/>/u)
  }
  assert.deepEqual(document, before)
})
