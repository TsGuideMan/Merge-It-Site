(function (root, factory) {
  const engine = factory();
  if (typeof module === 'object' && module.exports) module.exports = engine;
  if (root) root.MergeItEngine = engine;
})(typeof globalThis === 'undefined' ? this : globalThis, function () {
  'use strict';

  const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
  const SLIDE_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';

  function escapeXml(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function replaceTemplateValues(xml, headers, row) {
    return headers.reduce((result, header, index) => {
      const token = `{{${header}}}`;
      return result.split(token).join(escapeXml(row[index]));
    }, xml);
  }

  function getSlidePaths(zip) {
    return Object.keys(zip.files)
      .filter(path => /^ppt\/slides\/slide\d+\.xml$/.test(path))
      .sort((a, b) => Number(a.match(/slide(\d+)\.xml$/)[1]) - Number(b.match(/slide(\d+)\.xml$/)[1]));
  }

  function readText(bytes) {
    return new TextDecoder().decode(bytes);
  }

  async function buildPerRowPresentation(JSZip, templateZip, headers, row) {
    const outputZip = new JSZip();
    for (const path of Object.keys(templateZip.files)) {
      const file = templateZip.files[path];
      if (file.dir) continue;
      const content = await file.async('uint8array');
      if (/^ppt\/slides\/slide\d+\.xml$/.test(path)) {
        outputZip.file(path, replaceTemplateValues(readText(content), headers, row));
      } else {
        outputZip.file(path, content);
      }
    }
    return outputZip;
  }

  function getNextRelationshipId(xml) {
    const ids = [...xml.matchAll(/\bId="rId(\d+)"/g)].map(match => Number(match[1]));
    return Math.max(0, ...ids) + 1;
  }

  function updatePresentationRelationships(xml, slideCount, firstNewId) {
    const relationshipMatches = [...xml.matchAll(/<Relationship\b[^>]*\/>/g)];
    if (!relationshipMatches.length && !/<Relationships\b[^>]*>\s*<\/Relationships>/.test(xml)) {
      throw new Error('Template presentation relationships are missing or invalid.');
    }
    const nonSlideRelationships = relationshipMatches
      .map(match => match[0])
      .filter(tag => !new RegExp(`\\bType="${SLIDE_RELATIONSHIP}"`).test(tag));
    const slideRelationships = Array.from({ length: slideCount }, (_, index) =>
      `<Relationship Id="rId${firstNewId + index}" Type="${SLIDE_RELATIONSHIP}" Target="slides/slide${index + 1}.xml"/>`
    );
    const openingTag = xml.match(/<Relationships\b[^>]*>/);
    if (!openingTag) throw new Error('Template presentation relationships are missing or invalid.');
    return `${openingTag[0]}${nonSlideRelationships.join('')}${slideRelationships.join('')}</Relationships>`;
  }

  function updateContentTypes(xml, slideCount) {
    const openingTag = xml.match(/<Types\b[^>]*>/);
    if (!openingTag) throw new Error('Template content types are missing or invalid.');
    const defaults = [...xml.matchAll(/<Default\b[^>]*\/>/g)].map(match => match[0]);
    const overrides = [...xml.matchAll(/<Override\b[^>]*\/>/g)]
      .map(match => match[0])
      .filter(tag => !/\bPartName="\/ppt\/slides\/slide\d+\.xml"/.test(tag));
    const slideOverrides = Array.from({ length: slideCount }, (_, index) =>
      `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
    );
    return `${openingTag[0]}${defaults.join('')}${overrides.join('')}${slideOverrides.join('')}</Types>`;
  }

  async function buildSinglePresentation(JSZip, templateZip, headers, rows, onSlide) {
    const slidePaths = getSlidePaths(templateZip);
    if (!slidePaths.length) throw new Error('Template has no slide content.');

    const outputZip = new JSZip();
    const presentationPath = 'ppt/presentation.xml';
    const presentationRelsPath = 'ppt/_rels/presentation.xml.rels';
    const contentTypesPath = '[Content_Types].xml';
    const presentationFile = templateZip.file(presentationPath);
    const presentationRelsFile = templateZip.file(presentationRelsPath);
    const contentTypesFile = templateZip.file(contentTypesPath);
    if (!presentationFile || !presentationRelsFile || !contentTypesFile) {
      throw new Error('This PowerPoint template is missing required presentation files.');
    }
    const presentationXml = readText(await presentationFile.async('uint8array'));
    const presentationRelsXml = readText(await presentationRelsFile.async('uint8array'));
    const contentTypesXml = readText(await contentTypesFile.async('uint8array'));
    const firstSlideRelationshipId = getNextRelationshipId(presentationRelsXml);
    const slideEntries = [];
    let slideNumber = 0;

    for (const row of rows) {
      for (const slidePath of slidePaths) {
        slideNumber += 1;
        const slideXml = await templateZip.file(slidePath).async('uint8array');
        outputZip.file(`ppt/slides/slide${slideNumber}.xml`, replaceTemplateValues(readText(slideXml), headers, row));

        const sourceRelPath = slidePath.replace(/^ppt\/slides\//, 'ppt/slides/_rels/').replace(/\.xml$/, '.xml.rels');
        const slideRels = templateZip.file(sourceRelPath);
        if (slideRels) {
          outputZip.file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`, await slideRels.async('uint8array'));
        }

        slideEntries.push(`<p:sldId id="${255 + slideNumber}" r:id="rId${firstSlideRelationshipId + slideNumber - 1}"/>`);
        if (onSlide) onSlide(slideNumber, rows.length * slidePaths.length);
      }
    }

    for (const path of Object.keys(templateZip.files)) {
      const file = templateZip.files[path];
      if (file.dir || path === presentationPath || path === presentationRelsPath || path === contentTypesPath) continue;
      if (/^ppt\/slides\/slide\d+\.xml$/.test(path) || /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(path)) continue;
      outputZip.file(path, await file.async('uint8array'));
    }

    if (!/<p:sldIdLst\b[^>]*>[\s\S]*?<\/p:sldIdLst>/.test(presentationXml)) {
      throw new Error('Template presentation has no slide list to update.');
    }
    outputZip.file(presentationPath, presentationXml.replace(/<p:sldIdLst\b[^>]*>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${slideEntries.join('')}</p:sldIdLst>`));
    outputZip.file(presentationRelsPath, updatePresentationRelationships(presentationRelsXml, slideNumber, firstSlideRelationshipId));
    outputZip.file(contentTypesPath, updateContentTypes(contentTypesXml, slideNumber));
    return outputZip;
  }

  return {
    PPTX_MIME,
    buildPerRowPresentation,
    buildSinglePresentation,
    escapeXml,
    getSlidePaths,
    getNextRelationshipId,
    replaceTemplateValues
  };
});
