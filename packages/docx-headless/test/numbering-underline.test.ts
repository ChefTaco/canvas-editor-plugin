import test from 'node:test'
import assert from 'node:assert/strict'
import { createZipBuffer, importDocx, exportDocx, paragraphsOf, validateDocxPackage } from '../src/index'

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const rel = 'http://schemas.openxmlformats.org/package/2006/relationships'
const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
async function source(body: string, numbering = '', styles = '', header = ''): Promise<Buffer> {
  return createZipBuffer([
    { name: '[Content_Types].xml', content: '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>' },
    { name: '_rels/.rels', content: `<Relationships xmlns="${rel}"><Relationship Id="root" Type="${office}/officeDocument" Target="word/document.xml"/></Relationships>` },
    { name: 'word/_rels/document.xml.rels', content: `<Relationships xmlns="${rel}">${header ? `<Relationship Id="header" Type="${office}/header" Target="header1.xml"/>` : ''}</Relationships>` },
    { name: 'word/document.xml', content: `<w:document xmlns:w="${w}" xmlns:r="${office}"><w:body>${body}<w:sectPr>${header ? '<w:headerReference w:type="default" r:id="header"/>' : ''}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/></w:sectPr></w:body></w:document>` },
    ...(numbering ? [{ name: 'word/numbering.xml', content: `<w:numbering xmlns:w="${w}">${numbering}</w:numbering>` }] : []),
    ...(styles ? [{ name: 'word/styles.xml', content: `<w:styles xmlns:w="${w}">${styles}</w:styles>` }] : []),
    ...(header ? [{ name: 'word/header1.xml', content: `<w:hdr xmlns:w="${w}">${header}</w:hdr>` }] : [])
  ])
}
const paragraph = (id = '1') => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${id}"/></w:numPr></w:pPr><w:r><w:t>Clue</w:t></w:r></w:p>`
const definition = (properties = '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>', override = '') => `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0">${properties}</w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/>${override}</w:num>`

void test('Roman numbering, custom starts/labels and overrides cannot report lossless import', async () => {
  const cases = [
    { properties: '<w:start w:val="7"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="Clue %1:"/>', override: '', codes: ['SOURCE_NUMBERING_FORMAT', 'SOURCE_NUMBERING_START', 'SOURCE_NUMBERING_LABEL'] },
    { properties: '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>', override: '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="7"/></w:lvlOverride>', codes: ['SOURCE_NUMBERING_OVERRIDE'] }
  ]
  for (const item of cases) {
    const imported = await importDocx(await source(paragraph(), definition(item.properties, item.override)))
    for (const code of item.codes) assert.ok(imported.losses.some(loss => loss.code === code), code)
    assert.equal(paragraphsOf(imported.document)[0]?.runs[0]?.text, 'Clue')
  }
})
void test('independent list identities, resumed numbering and missing definitions are diagnosed', async () => {
  const independent = await importDocx(await source(paragraph() + paragraph('2'), definition() + '<w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num>'))
  assert.ok(independent.losses.some(loss => loss.code === 'SOURCE_NUMBERING_IDENTITY'))
  const resumed = await importDocx(await source(paragraph() + '<w:p><w:r><w:t>Interlude</w:t></w:r></w:p>' + paragraph(), definition()))
  assert.ok(resumed.losses.some(loss => loss.code === 'SOURCE_NUMBERING_CONTINUATION'))
  const missing = await importDocx(await source(paragraph(), ''))
  assert.ok(missing.losses.some(loss => loss.code === 'SOURCE_NUMBERING_REFERENCE'))
})
void test('default flat numbering and single/none underline remain supported on round trip', async () => {
  const imported = await importDocx(await source(paragraph() + '<w:p><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>Single</w:t></w:r><w:r><w:rPr><w:u w:val="none"/></w:rPr><w:t>None</w:t></w:r></w:p>', definition()))
  assert.deepEqual(imported.losses, [])
  const before = structuredClone(imported.document)
  const reopened = await importDocx(await exportDocx(imported.document, 'retain-review'))
  assert.deepEqual(reopened.losses, [])
  assert.equal(paragraphsOf(reopened.document)[0]?.style.list, 'number')
  assert.deepEqual(paragraphsOf(reopened.document)[1]?.runs.map(run => run.style.underline), [true, false])
  assert.deepEqual(imported.document, before)
})
void test('unused numbering definitions do not disqualify documents with no lists', async () => {
  const imported = await importDocx(await source('<w:p><w:r><w:t>Plain prose</w:t></w:r></w:p>', definition('<w:start w:val="7"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="Clue %1:"/>')))
  assert.deepEqual(imported.losses, [])
})
void test('default bullets remain supported and custom numbering is audited in stories, styles, cells and previous snapshots', async () => {
  const bullet = await importDocx(await source(paragraph(), definition('<w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>')))
  assert.deepEqual(bullet.losses, [])
  assert.deepEqual((await importDocx(await exportDocx(bullet.document, 'retain-review'))).losses, [])
  const properties = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
  const custom = definition('<w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/>')
  const cases = [
    { body: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>', styles: '', header: paragraph() },
    { body: '<w:p><w:pPr><w:pStyle w:val="List"/></w:pPr><w:r><w:t>Style</w:t></w:r></w:p>', styles: `<w:style w:type="paragraph" w:styleId="List"><w:pPr>${properties}</w:pPr></w:style>`, header: '' },
    { body: `<w:p><w:pPr><w:pPrChange w:id="1" w:author="Reviewer" w:date="2026-10-22T00:00:00Z"><w:pPr>${properties}</w:pPr></w:pPrChange></w:pPr><w:r><w:t>Changed</w:t></w:r></w:p>`, styles: '', header: '' },
    { body: `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${paragraph()}</w:tc></w:tr></w:tbl>`, styles: '', header: '' }
  ]
  for (const item of cases) {
    const imported = await importDocx(await source(item.body, custom, item.styles, item.header))
    assert.ok(imported.losses.some(loss => loss.code === 'SOURCE_NUMBERING_FORMAT'))
  }
})
void test('underline variants and underline-specific attributes report losses in all property sources', async () => {
  const revision = '<w:rPrChange w:id="1" w:author="Reviewer" w:date="2026-10-22T00:00:00Z"><w:rPr><w:u w:val="double"/></w:rPr></w:rPrChange>'
  const imported = await importDocx(await source(
    `<w:p><w:r><w:rPr><w:u w:val="single"/>${revision}</w:rPr><w:t>Changed</w:t></w:r></w:p><w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:rPr><w:u w:val="dotted"/></w:rPr><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`, '',
    '<w:docDefaults><w:rPrDefault><w:rPr><w:u w:val="double"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="character" w:styleId="Colored"><w:rPr><w:u w:val="single" w:color="FF0000"/></w:rPr></w:style>',
    '<w:p><w:r><w:rPr><w:u w:val="single" w:themeColor="accent1"/></w:rPr><w:t>Header</w:t></w:r></w:p>'
  ))
  assert.equal(imported.losses.find(loss => loss.code === 'SOURCE_RUN_UNDERLINE_UNSUPPORTED')?.count, 5)
  const change = imported.document.changes.find(item => item.kind === 'format')
  assert.ok(change); assert.equal(change.beforeStyle?.underline, true); assert.equal(change.afterStyle?.underline, true)
})
void test('double underline from an independent source can never silently become single', async () => {
  const bytes = await source('<w:p><w:r><w:rPr><w:u w:val="double"/></w:rPr><w:t>Clue</w:t></w:r></w:p>')
  const imported = await importDocx(bytes)
  assert.ok(imported.losses.some(loss => loss.code === 'SOURCE_RUN_UNDERLINE_UNSUPPORTED'))
  const emitted = await validateDocxPackage(await exportDocx(imported.document, 'retain-review'))
  assert.match(emitted.get('word/document.xml')?.toString('utf8') ?? '', /w:u w:val="single"/u)
})
