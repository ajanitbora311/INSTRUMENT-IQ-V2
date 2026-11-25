const byId = id => document.getElementById(id);
const input = byId('workbook');
const download = byId('download');
const drop = byId('drop');
let selectedFile;
let catalogRequest;
let estimate;

function announce(text, error = false) {
  const status = byId('message');
  status.hidden = !text;
  status.textContent = text;
  status.classList.toggle('error', error);
}

function selectFile(file) {
  selectedFile = undefined;
  estimate = undefined;
  byId('estimate-actions').hidden = byId('dashboard').hidden = true;
  announce('');
  if (file && !/\.(xlsx|xls|xlsm)$/i.test(file.name)) {
    announce('Choose an XLSX, XLS, or XLSM workbook.', true);
  } else if (file && file.size > 30 * 1024 * 1024) {
    announce('Choose a workbook under 30 MB.', true);
  } else {
    selectedFile = file;
  }
  byId('file-label').textContent = selectedFile?.name || 'Choose Excel file';
  download.disabled = !selectedFile;
}

input.addEventListener('change', () => selectFile(input.files[0]));
drop.addEventListener('dragover', event => {
  event.preventDefault();
  drop.classList.add('active');
});
drop.addEventListener('dragleave', () => drop.classList.remove('active'));
drop.addEventListener('drop', event => {
  event.preventDefault();
  drop.classList.remove('active');
  if (download.dataset.busy) return;
  if (event.dataTransfer.files.length !== 1) {
    input.value = '';
    selectFile(undefined);
    announce('Drop one Excel workbook at a time.', true);
    return;
  }
  input.value = '';
  selectFile(event.dataTransfer.files[0]);
});

function fetchCatalog() {
  catalogRequest ??= fetch('./catalog.json')
    .then(async response => {
      if (!response.ok) throw new Error();
      const catalog = await response.json();
      for (const name of ['sor', 'po', 'analysis', 'material']) {
        if (!Array.isArray(catalog[name])) throw new Error();
      }
      return catalog;
    })
    .catch(() => {
      catalogRequest = undefined;
      throw new Error('The rate catalog is unavailable. Contact the site administrator.');
    });
  return catalogRequest;
}

download.addEventListener('click', async () => {
  if (!selectedFile) return;
  const file = selectedFile;
  estimate = undefined;
  byId('estimate-actions').hidden = byId('dashboard').hidden = true;
  download.disabled = input.disabled = true;
  download.dataset.busy = 'true';
  announce('Preparing your estimate…');
  try {
    const workbook = XLSX.read(await file.arrayBuffer(), {type:'array'});
    const takeoff = readBook(workbook, 'boq', file.name);
    const catalog = await fetchCatalog();
    const rows = priceRows(takeoff, catalog.sor, catalog.po, catalog.analysis);
    estimate = {rows, catalog, name:file.name};
    byId('estimate-actions').hidden = false;
    const review = rows.filter(row => row.status === 'Review').length;
    announce('Estimate ready. ' + (rows.length - review) + ' priced · ' + review + ' for review.');
  } catch (error) {
    announce(error.message || 'The estimate could not be generated.', true);
  } finally {
    download.disabled = input.disabled = false;
    delete download.dataset.busy;
  }
});

byId('view').addEventListener('click', () => {
  if (estimate) showDashboard(estimate);
});

byId('export').addEventListener('click', () => {
  if (!estimate) return;
  try {
    XLSX.writeFile(exportRows(estimate.rows, estimate.catalog.material),
      estimate.name.replace(/\.[^.]+$/, '') + '-estimate.xlsx');
  } catch (error) {
    announce(error.message || 'The estimate could not be exported.', true);
  }
});

byId('template').addEventListener('click', () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([[...HEAD.boq, ...OUTPUT.slice(0, 5)]]);
  sheet['!cols'] = [10, 32, 60, 12, 14, 16, 18, 28, 28, 28].map(wch => ({wch}));
  XLSX.utils.book_append_sheet(workbook, sheet, 'Take-off');
  XLSX.writeFile(workbook, 'take-off-template.xlsx');
});
