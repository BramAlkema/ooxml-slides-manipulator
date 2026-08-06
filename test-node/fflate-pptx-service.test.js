const assert = require('node:assert/strict');
const { after, before, test } = require('node:test');
const fflate = require('fflate');

const FFlatePPTXService = require('../lib/FFlatePPTXService');

function gasBlob(bytes, contentType, name) {
  let blobContentType = contentType;
  let blobName = name;
  const data = Uint8Array.from(bytes);

  return {
    getBytes: () => Array.from(data),
    getContentType: () => blobContentType,
    getName: () => blobName,
    setContentType(value) {
      blobContentType = value;
      return this;
    },
    setName(value) {
      blobName = value;
      return this;
    },
  };
}

before(() => {
  global.fflate = fflate;
  global.Utilities = { newBlob: gasBlob };
});

after(() => {
  delete global.fflate;
  delete global.Utilities;
});

test('compresses and extracts OOXML text and binary parts locally', () => {
  const files = {
    '[Content_Types].xml': '<Types/>',
    '_rels/.rels': '<Relationships/>',
    'ppt/slides/slide1.xml': '<p:sld>hello</p:sld>',
    'ppt/media/image1.png': Uint8Array.from([0, 127, 128, 255]),
  };

  const blob = FFlatePPTXService.compressFiles(files);
  const extracted = FFlatePPTXService.extractFiles(blob);

  assert.equal(
    blob.getContentType(),
    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  );
  assert.equal(blob.getName(), 'presentation.pptx');
  assert.equal(extracted['[Content_Types].xml'], files['[Content_Types].xml']);
  assert.equal(extracted['_rels/.rels'], files['_rels/.rels']);
  assert.equal(extracted['ppt/slides/slide1.xml'], files['ppt/slides/slide1.xml']);
  assert.deepEqual(extracted['ppt/media/image1.png'], files['ppt/media/image1.png']);
});

test('rejects archives without the required OOXML content types part', () => {
  assert.throws(
    () => FFlatePPTXService.compressFiles({ '_rels/.rels': '<Relationships/>' }),
    /FFLATE_COMPRESS_ERROR: Missing required file: \[Content_Types\]\.xml/,
  );
});
