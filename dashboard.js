const dashboardId = id => document.getElementById(id);
const sourceLabels = {analysis:'Rate analysis workbook', sor:'NRL SOR', po:'Service PO', material:'Material PO'};
const inr = new Intl.NumberFormat('en-IN', {style:'currency', currency:'INR'});
const rateDisplay = new Intl.NumberFormat('en-IN', {maximumFractionDigits:8});
const pageSize = 25;
let dashboardState;

function appendText(parent, tag, text) {
  const element = document.createElement(tag);
  element.textContent = text ?? '';
  parent.append(element);
  return element;
}

function showDashboard(estimate) {
  if (dashboardState?.estimate !== estimate) {
    const {rows, catalog, name} = estimate;
    const priced = rows.filter(row => row.source).length;
    dashboardState = {estimate, page:0};
    dashboardState.createSearch = () => createCandidateSearch(
      [...catalog.sor, ...catalog.po, ...catalog.analysis, ...catalog.material]);
    dashboardId('estimate-name').textContent = name;
    dashboardId('dash-total').textContent = inr.format(estimateTotal(rows));
    dashboardId('dash-priced').textContent = priced;
    dashboardId('dash-review').textContent = rows.length - priced;
    dashboardId('source-filter').value = 'all';
    renderDashboard();
  }
  dashboardId('dashboard').hidden = false;
  dashboardId('dashboard-title').focus();
  dashboardId('dashboard').scrollIntoView({block:'start'});
}

function renderDashboard() {
  const {estimate, page} = dashboardState;
  const filter = dashboardId('source-filter').value;
  const rows = estimate.rows.filter(row => filter === 'all' || (row.source?.kind ?? row.status) === filter);
  const tbody = dashboardId('dash-rows');
  tbody.replaceChildren();
  const pages = Math.max(1, Math.ceil(rows.length / pageSize));
  dashboardId('dash-page').textContent = (page + 1) + ' of ' + pages;
  dashboardId('dash-prev').disabled = page === 0;
  dashboardId('dash-next').disabled = page + 1 >= pages;
  if (!rows.length) appendText(appendText(tbody, 'tr'), 'td', 'No rows match this filter.').colSpan = 6;
  for (const row of rows.slice(page * pageSize, (page + 1) * pageSize)) {
    const tr = appendText(tbody, 'tr');
    const [serial, short, long, unit, quantity] = row.input;
    const input = appendText(tr, 'td', serial);
    if (row.sheet && row.row) appendText(input, 'p', row.sheet + ' · row ' + row.row);
    const description = appendText(tr, 'td', short || long);
    if (short && long && short !== long) {
      const details = appendText(description, 'details');
      appendText(details, 'summary', 'Take-off description');
      appendText(details, 'p', long);
    }
    appendText(tr, 'td', quantity + ' ' + unit);
    appendText(tr, 'td', row.source ? '₹' + rateDisplay.format(row.rate) : '—');
    appendText(tr, 'td', row.source ? inr.format(row.amount) : '—');
    const provenance = appendText(tr, 'td');
    if (row.source) {
      const source = row.source;
      appendText(provenance, 'strong', sourceLabels[source.kind] + ' · ' + source.reference);
      appendText(provenance, 'p', source.name + ' · ' + source.sheet + ' · row ' + source.row);
      const details = appendText(provenance, 'details');
      appendText(details, 'summary', 'Rate details');
      appendText(details, 'p', row.method + ' · ₹' + rateDisplay.format(source.rate) + ' per ' + source.unit);
      appendText(details, 'p', source.description);
      if (row.audit) appendText(details, 'p', row.audit);
    } else {
      appendText(provenance, 'strong', 'Review');
      appendText(provenance, 'p', row.reason);
      const details = appendText(provenance, 'details');
      appendText(details, 'summary', 'Related source records');
      details.addEventListener('toggle', () => {
        if (!details.open || details.dataset.loaded) return;
        details.dataset.loaded = 'true';
        const search = dashboardState.lookup ??= dashboardState.createSearch();
        const matches = search(long || short, unit);
        if (!matches.length) {
          appendText(details, 'p', 'No related records found with a compatible unit.');
          return;
        }
        const list = appendText(details, 'ul');
        for (const {source} of matches) {
          const item = appendText(list, 'li', sourceLabels[source.kind] + ' · ' + source.reference);
          appendText(item, 'p', source.short || source.description);
          if (source.kind === 'material' && Number.isFinite(source.rate)) {
            const value = source.currency === 'INR'
              ? inr.format(source.rate) : [source.currency, rateDisplay.format(source.rate)].filter(Boolean).join(' ');
            appendText(item, 'small', 'Reported PO unit price ' + value + (source.unit ? ' per ' + source.unit : ''));
          }
          appendText(item, 'small', source.name + ' · ' + source.sheet + ' · row ' + source.row);
        }
      });
    }
  }
}

dashboardId('source-filter').addEventListener('change', () => {
  dashboardState.page = 0;
  renderDashboard();
});
dashboardId('dash-prev').addEventListener('click', () => { dashboardState.page--; renderDashboard(); });
dashboardId('dash-next').addEventListener('click', () => { dashboardState.page++; renderDashboard(); });
