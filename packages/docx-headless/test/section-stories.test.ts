import test from 'node:test'
import assert from 'node:assert/strict'
import {defaultLayout,exportDocxSections,importDocx,validateDocxPackage,type WordDocument,type WordPackageInput} from '../src/index'

function document(text:string):WordDocument {
  return {schemaVersion:1,layout:{...defaultLayout(),pageNumbers:false,header:[],footer:[]},blocks:[{type:'paragraph',id:'p',style:{},runs:[{id:'r',text,style:{}}]}],changes:[],comments:[]}
}
function book(documents:WordDocument[]):WordPackageInput {
  return {sections:documents.map((document,index)=>({title:'Section '+index,suppressTitle:true,breakAfter:'continuous',document})),assets:new Map()}
}
void test('section export does not manufacture globally absent header/footer paragraphs in any review view',async()=>{
  for(const count of [1,3])for(const mode of ['retain-review','original','final'] as const){
    const input=book(Array.from({length:count},(_,i)=>document('Body '+i)));const before=structuredClone(input)
    const bytes=await exportDocxSections(input,mode);const files=await validateDocxPackage(bytes,true)
    assert.ok(!Array.from(files.keys()).some(name=>/\/(?:header|footer)\d+\.xml/u.test(name)))
    assert.doesNotMatch(files.get('word/document.xml')?.toString()??'',/headerReference|footerReference/u)
    assert.doesNotMatch(files.get('word/_rels/document.xml.rels')?.toString()??'',/relationships\/(?:header|footer)/u)
    assert.doesNotMatch(files.get('[Content_Types].xml')?.toString()??'',/wordprocessingml\.(?:header|footer)/u)
    const reopened=await importDocx(bytes);assert.deepEqual(reopened.losses,[]);assert.equal(reopened.sections.length,count)
    for(const section of reopened.sections){assert.deepEqual(section.document.layout.header,[]);assert.deepEqual(section.document.layout.footer,[]);assert.equal(section.stories?.header,undefined);assert.equal(section.stories?.footer,undefined)}
    assert.deepEqual(reopened.document.blocks.filter(block=>block.type==='paragraph').flatMap(block=>block.runs.map(run=>run.text)),Array.from({length:count},(_,i)=>'Body '+i))
    assert.deepEqual(input,before)
  }
})
void test('an empty middle story still clears an earlier header while the absent footer is omitted',async()=>{
  const first=document('First'),middle=document('Middle'),last=document('Last')
  first.layout.header=[{type:'paragraph',id:'h',style:{},runs:[{id:'hr',text:'First header',style:{}}]}]
  last.layout.header=[{type:'paragraph',id:'h',style:{},runs:[{id:'hr',text:'Last header',style:{}}]}]
  const bytes=await exportDocxSections(book([first,middle,last]),'retain-review');const files=await validateDocxPackage(bytes,true)
  assert.equal((files.get('word/document.xml')?.toString().match(/headerReference/gu)??[]).length,3)
  assert.doesNotMatch(files.get('word/document.xml')?.toString()??'',/footerReference/u)
  const result=await importDocx(bytes);assert.deepEqual(result.losses,[])
  assert.deepEqual(result.sections.map(section=>section.document.layout.header.flatMap(p=>p.runs.map(run=>run.text)).join('')),['First header','','Last header'])
  assert.equal(result.sections[1]?.document.layout.header.length,1)
})
void test('explicit blank paragraphs and generated page-number footers remain authored stories',async()=>{
  const blank=document('Blank story'),numbered=document('Numbered')
  blank.layout.header=[{type:'paragraph',id:'h',style:{before:12,after:8},runs:[{id:'hr',text:'',style:{}}]}];numbered.layout.pageNumbers=true
  const bytes=await exportDocxSections(book([blank,numbered]),'retain-review');const files=await validateDocxPackage(bytes,true)
  assert.match(files.get('word/document.xml')?.toString()??'',/headerReference/u)
  assert.match(files.get('word/document.xml')?.toString()??'',/footerReference/u)
  const result=await importDocx(bytes);assert.deepEqual(result.losses,[])
  assert.equal(result.sections[0]?.document.layout.header[0]?.style.before,12)
  assert.equal(result.sections[1]?.document.layout.pageNumbers,true)
})
void test('owned story documents take precedence, including non-paragraph content',async()=>{
  const main=document('Body'),story=document('Story');main.layout.header=[{type:'paragraph',id:'unused',style:{},runs:[{id:'unused-r',text:'',style:{}}]}]
  const input=book([main]);const section=input.sections[0];assert.ok(section)
  const header=document('Header cell');const paragraph=header.blocks[0];assert.ok(paragraph?.type==='paragraph')
  header.blocks=[{type:'table',id:'table',widths:[300],rows:[{id:'row',cells:[{id:'cell',paragraphs:[paragraph]}]}]}]
  section.stories={header,footer:story}
  const bytes=await exportDocxSections(input,'retain-review');const files=await validateDocxPackage(bytes,true)
  assert.match(files.get('word/document.xml')?.toString()??'',/headerReference/u)
  assert.match(files.get('word/document.xml')?.toString()??'',/footerReference/u)
  const reopened=await importDocx(bytes);assert.deepEqual(reopened.losses,[])
  assert.equal(reopened.sections[0]?.stories?.header?.blocks[0]?.type,'table')
  assert.equal(reopened.sections[0]?.document.layout.footer[0]?.runs[0]?.text,'Story')
})
