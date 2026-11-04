(function () {
  'use strict';

  const state = { workbook: null, headers: [], rows: [], valid: false };
  const $ = id => document.getElementById(id);
  const dataInput = $('dataFile');
  const pptxInput = $('pptxFile');
  const mergeButton = $('mergeBtn');
  const statusTitle = $('statusTitle');
  const statusMessage = $('statusMessage');
  const preview = $('preview');
  const selectedMode = () => document.querySelector('input[name="mode"]:checked').value;

  function setStatus(title, message, percent, isError) {
    statusTitle.textContent = title;
    statusMessage.textContent = message;
    statusMessage.classList.toggle('error', Boolean(isError));
    $('statusPercent').textContent = `${Math.max(0, Math.min(100, Math.round(percent)))}%`;
    $('progressBar').style.width = `${Math.max(0, Math.min(100, percent))}%`;
  }

  function checkReady() {
    const ready = state.valid && Boolean(pptxInput.files[0]);
    $('previewBtn').disabled = !ready;
    mergeButton.disabled = !ready;
  }

  function loadSheet(name) {
    state.valid = false;
    const sheet = state.workbook.Sheets[name];
    const values = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    state.headers = (values[0] || []).map(value => String(value).trim());
    state.rows = values.slice(1);
    const blankHeader = state.headers.findIndex(header => !header);
    const duplicateHeader = state.headers.find((header, index) => state.headers.indexOf(header) !== index);
    if (!state.headers.length || blankHeader >= 0) {
      setStatus('Check your spreadsheet', 'The first row must contain a non-empty header for every column.', 0, true);
    } else if (duplicateHeader) {
      setStatus('Check your spreadsheet', `The header "${duplicateHeader}" is repeated. Column headers must be unique.`, 0, true);
    } else if (!state.rows.length) {
      setStatus('No data rows found', 'Add at least one data row beneath the header row.', 0, true);
    } else {
      state.valid = true;
      setStatus('Spreadsheet ready', `${state.headers.length} columns and ${state.rows.length} data rows found in "${name}".`, 0, false);
    }
    checkReady();
  }

  dataInput.addEventListener('change', async () => {
    const file = dataInput.files[0];
    if (!file) return;
    $('dataFileName').textContent = file.name;
    setStatus('Reading spreadsheet', 'The selected file is being read locally in your browser.', 0, false);
    try {
      state.workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' });
      const selector = $('sheetSelect');
      selector.replaceChildren(...state.workbook.SheetNames.map(name => new Option(name, name)));
      selector.disabled = false;
      loadSheet(selector.value);
    } catch (error) {
      state.workbook = null;
      state.headers = [];
      state.rows = [];
      state.valid = false;
      $('sheetSelect').replaceChildren(new Option('Could not read spreadsheet', ''));
      $('sheetSelect').disabled = true;
      setStatus('Spreadsheet error', error.message || 'This file could not be read.', 0, true);
      checkReady();
    }
  });

  $('sheetSelect').addEventListener('change', event => loadSheet(event.target.value));
  pptxInput.addEventListener('change', () => {
    $('pptxFileName').textContent = pptxInput.files[0] ? pptxInput.files[0].name : 'Choose a .pptx template';
    checkReady();
  });

  $('previewBtn').addEventListener('click', () => {
    const startRow = Number($('startRow').value);
    const row = state.rows[startRow - 2];
    if (!Number.isInteger(startRow) || startRow < 2 || !row) {
      setStatus('Preview unavailable', 'Choose a start row that contains data.', 0, true);
      return;
    }
    const heading = document.createElement('h3');
    heading.textContent = `Preview · spreadsheet row ${startRow}`;
    const table = document.createElement('table');
    const body = document.createElement('tbody');
    state.headers.forEach((header, index) => {
      const tr = document.createElement('tr');
      const token = document.createElement('th');
      const value = document.createElement('td');
      token.textContent = `{{${header}}}`;
      value.textContent = String(row[index] == null ? '' : row[index]);
      tr.append(token, value);
      body.append(tr);
    });
    table.append(body);
    preview.replaceChildren(heading, table);
    preview.hidden = false;
  });

  function download(blob, name) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = name;
    document.body.append(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function safeBaseName(value) {
    return value.trim().replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_').replace(/[. ]+$/g, '') || 'Report';
  }

  function selectedRows() {
    const start = Number($('startRow').value);
    const endValue = $('endRow').value.trim();
    const end = endValue ? Number(endValue) : state.rows.length + 1;
    if (!Number.isInteger(start) || start < 2 || !Number.isInteger(end) || end < start) {
      throw new Error('Enter a valid row range. The start row must be 2 or later, and the end row must not come before it.');
    }
    if (start > state.rows.length + 1) throw new Error('The start row is past the end of the data.');
    return state.rows.slice(start - 2, Math.min(end - 1, state.rows.length));
  }

  mergeButton.addEventListener('click', async () => {
    mergeButton.disabled = true;
    $('previewBtn').disabled = true;
    preview.hidden = true;
    setStatus('Preparing files', 'Validating the selected row range and template.', 0, false);
    try {
      const rows = selectedRows();
      if (!rows.length) throw new Error('There are no data rows in the selected range.');
      const baseName = safeBaseName($('outName').value);
      const template = await JSZip.loadAsync(await pptxInput.files[0].arrayBuffer());
      if (!template.file('[Content_Types].xml') || !template.file('ppt/presentation.xml') ||
          !template.file('ppt/_rels/presentation.xml.rels') || !MergeItEngine.getSlidePaths(template).length) {
        throw new Error('The selected file is not a valid PowerPoint template with at least one slide.');
      }
      const mode = selectedMode();
      if (mode === 'single') {
        const output = await MergeItEngine.buildSinglePresentation(JSZip, template, state.headers, rows, (complete, total) => {
          setStatus('Building presentation', `Prepared ${complete} of ${total} slides.`, total ? complete / total * 75 : 0, false);
        });
        const blob = await output.generateAsync({ type: 'blob', mimeType: MergeItEngine.PPTX_MIME }, metadata => {
          setStatus('Preparing download', 'Packaging the combined PowerPoint file.', 75 + metadata.percent * .25, false);
        });
        download(blob, `${baseName}.pptx`);
        setStatus('Done', `Downloaded ${baseName}.pptx with ${rows.length} row${rows.length === 1 ? '' : 's'}.`, 100, false);
      } else {
        const archive = new JSZip();
        for (let index = 0; index < rows.length; index += 1) {
          const output = await MergeItEngine.buildPerRowPresentation(JSZip, template, state.headers, rows[index]);
          const fileName = `${baseName}_${String(index + 1).padStart(3, '0')}.pptx`;
          archive.file(fileName, await output.generateAsync({ type: 'uint8array', mimeType: MergeItEngine.PPTX_MIME }));
          setStatus('Building presentations', `Prepared ${index + 1} of ${rows.length} presentations.`, (index + 1) / rows.length * 80, false);
        }
        const blob = await archive.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 4 } }, metadata => {
          setStatus('Preparing download', 'Packaging your presentations in one ZIP file.', 80 + metadata.percent * .2, false);
        });
        const archiveName = rows.length === 1 ? `${baseName}.zip` : `${baseName}_presentations.zip`;
        download(blob, archiveName);
        setStatus('Done', `Downloaded ${archiveName} containing ${rows.length} PowerPoint presentation${rows.length === 1 ? '' : 's'}.`, 100, false);
      }
    } catch (error) {
      setStatus('Could not create files', error.message || 'The selected files could not be processed.', 0, true);
      console.error(error);
    } finally {
      checkReady();
    }
  });

  $('clearBtn').addEventListener('click', () => {
    dataInput.value = '';
    pptxInput.value = '';
    $('dataFileName').textContent = 'Choose Excel or CSV file';
    $('pptxFileName').textContent = 'Choose a .pptx template';
    $('sheetSelect').replaceChildren(new Option('Choose a spreadsheet first', ''));
    $('sheetSelect').disabled = true;
    $('outName').value = 'Report';
    $('startRow').value = '2';
    $('endRow').value = '';
    document.querySelector('input[name="mode"][value="perRow"]').checked = true;
    state.workbook = null;
    state.headers = [];
    state.rows = [];
    state.valid = false;
    preview.replaceChildren();
    preview.hidden = true;
    setStatus('Ready when you are', 'Select a spreadsheet and a PowerPoint template to begin.', 0, false);
    checkReady();
  });
})();
