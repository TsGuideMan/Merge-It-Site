'use strict';

const assert = require('node:assert/strict');
const test = require('node:test');
const JSZip = require('../app/vendor/jszip.min.js');
const engine = require('../app/merge-engine.js');

async function makeTemplate() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/></p:sldIdLst></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="theme" Target="theme/theme1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', '<p:sld><a:t>{{Name}}</a:t></p:sld>');
  zip.file('ppt/slides/slide2.xml', '<p:sld><a:t>{{City}}</a:t></p:sld>');
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships><Relationship Id="rId1" Type="layout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>');
  zip.file('ppt/slides/_rels/slide2.xml.rels', '<Relationships><Relationship Id="rId1" Type="layout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>');
  zip.file('ppt/slideLayouts/slideLayout1.xml', '<layout/>');
  zip.file('ppt/theme/theme1.xml', '<theme/>');
  return zip;
}

test('escapes values and replaces exact spreadsheet placeholders', () => {
  const result = engine.replaceTemplateValues('<a>{{Name}}</a><a>{{City}}</a>', ['Name', 'City'], ['A & B <C>', null]);
  assert.equal(result, '<a>A &amp; B &lt;C&gt;</a><a></a>');
});

test('creates a per-row presentation while preserving template parts', async () => {
  const template = await makeTemplate();
  const output = await engine.buildPerRowPresentation(JSZip, template, ['Name', 'City'], ['Ada', 'Oslo']);
  assert.equal(await output.file('ppt/slides/slide1.xml').async('string'), '<p:sld><a:t>Ada</a:t></p:sld>');
  assert.equal(await output.file('ppt/slides/slide2.xml').async('string'), '<p:sld><a:t>Oslo</a:t></p:sld>');
  assert.ok(output.file('ppt/theme/theme1.xml'));
});

test('creates a combined presentation with matching slides, relationships, and content types', async () => {
  const template = await makeTemplate();
  const output = await engine.buildSinglePresentation(JSZip, template, ['Name', 'City'], [['Ada', 'Oslo'], ['Lin', 'Lima']]);
  assert.deepEqual(engine.getSlidePaths(output).map(path => path.replace(/^ppt\/slides\//, '')), [
    'slide1.xml', 'slide2.xml', 'slide3.xml', 'slide4.xml'
  ]);
  assert.equal(await output.file('ppt/slides/slide3.xml').async('string'), '<p:sld><a:t>Lin</a:t></p:sld>');
  const relationships = await output.file('ppt/_rels/presentation.xml.rels').async('string');
  assert.equal((relationships.match(/relationships\/slide"/g) || []).length, 4);
  assert.match(relationships, /Id="rId1" Type="theme"/);
  assert.match(relationships, /Target="slides\/slide4\.xml"/);
  const contentTypes = await output.file('[Content_Types].xml').async('string');
  assert.equal((contentTypes.match(/PartName="\/ppt\/slides\/slide\d+\.xml"/g) || []).length, 4);
  const presentation = await output.file('ppt/presentation.xml').async('string');
  assert.equal((presentation.match(/<p:sldId\b/g) || []).length, 4);
  assert.match(presentation, /r:id="rId4"/);
  assert.match(presentation, /r:id="rId7"/);
  assert.ok(output.file('ppt/slides/_rels/slide4.xml.rels'));
});
