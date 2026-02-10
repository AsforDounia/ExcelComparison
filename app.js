const file1Input = document.getElementById('file1');
const file2Input = document.getElementById('file2');
const file1Status = document.getElementById('file1-status');
const file2Status = document.getElementById('file2-status');
const compareBtn = document.getElementById('compare-btn');
const summaryDiv = document.getElementById('summary');
const resultsDiv = document.getElementById('results');
const notFoundSection = document.getElementById('not-found-section');
const diffTbody = document.getElementById('diff-tbody');
const notfoundTbody = document.getElementById('notfound-tbody');
const noDiffMsg = document.getElementById('no-diff');
const totalCountEl = document.getElementById('total-count');
const matchCountEl = document.getElementById('match-count');
const diffCountEl = document.getElementById('diff-count');
const notfoundCountEl = document.getElementById('notfound-count');

let data1 = null; // Fichier 1: Produits, Camions, Qté facturées...
let data2 = null; // Fichier 2: Référence interne, Total poids net...

file1Input.addEventListener('change', () => handleFile(file1Input, 1));
file2Input.addEventListener('change', () => handleFile(file2Input, 2));
compareBtn.addEventListener('click', compareFiles);

function handleFile(input, fileNum) {
  const file = input.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const workbook = XLSX.read(e.target.result, { type: 'binary' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });

    if (fileNum === 1) {
      data1 = json;
      file1Status.textContent = `✅ ${json.length} lignes chargées`;
    } else {
      data2 = json;
      file2Status.textContent = `✅ ${json.length} lignes chargées`;
    }

    compareBtn.disabled = !(data1 && data2);
  };
  reader.readAsBinaryString(file);
}

function findColumn(headers, candidates) {
  for (const candidate of candidates) {

    const exact = headers.find(h => h === candidate);
    if (exact) return exact;

    const trimmed = headers.find(h => h.trim().toLowerCase() === candidate.toLowerCase());
    if (trimmed) return trimmed;

    const partial = headers.find(h => h.toLowerCase().includes(candidate.toLowerCase()));
    if (partial) return partial;
  }
  return null;
}

function parseNumber(val) {
  if (val === '' || val === null || val === undefined) return 0;
  if (typeof val === 'number') return val;
  const cleaned = String(val).replace(/\s/g, '').replace(',', '.');
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

function compareFiles() {
  if (!data1 || !data2) {
    alert('Veuillez charger les deux fichiers.');
    return;
  }

  const headers1 = Object.keys(data1[0] || {});
  const headers2 = Object.keys(data2[0] || {});

  // Find columns in File 1
  const colCamions = findColumn(headers1, ['Camions', 'Camion']);
  const colQte = findColumn(headers1, ['Qté facturées', 'Qte facturees', 'Qté', 'Qte facturées']);
  const colProduit = findColumn(headers1, ['Produits', 'Produit']);

  // Find columns in File 2
  const colRef = findColumn(headers2, ['Ressource','Ressource']);
  const colPoidsNet = findColumn(headers2, ['Total poids net', 'Poids net', 'Total Poids Net']);

  // Validate
  const missing = [];
  if (!colCamions) missing.push('Camions (Fichier 1)');
  if (!colQte) missing.push('Qté facturées (Fichier 1)');
  if (!colRef) missing.push('Ressource (Fichier 2)');
  if (!colPoidsNet) missing.push('Total poids net (Fichier 2)');

  if (missing.length > 0) {
    alert('Colonnes introuvables :\n- ' + missing.join('\n- ') +
      '\n\nColonnes Fichier 1 : ' + headers1.join(', ') +
      '\n\nColonnes Fichier 2 : ' + headers2.join(', '));
    return;
  }

  const refMap = new Map();
  data2.forEach((row) => {
    const ref = String(row[colRef] ?? '').trim().replace(/\s/g, '');
    if (ref) {
      if (refMap.has(ref)) {
        const existing = refMap.get(ref);
        existing.poidsNet += parseNumber(row[colPoidsNet]);
        existing.rows.push(row);
      } else {
        refMap.set(ref, {
          poidsNet: parseNumber(row[colPoidsNet]),
          rows: [row]
        });
      }
    }
  });

  const differences = [];
  const notFound = [];
  let matchCount = 0;

  data1.forEach((row1) => {
    const camion = String(row1[colCamions] ?? '').trim().replace(/\s/g, '');
    if (!camion) return;

    const qte = parseNumber(row1[colQte]);
    const produit = row1[colProduit] || '';

    if (refMap.has(camion)) {
      const refData = refMap.get(camion);
      const poidsNet = refData.poidsNet;
      const ecart = qte - poidsNet;

      if (Math.abs(ecart) > 0.01) {
        differences.push({ camion, produit, qte, poidsNet, ecart });
      } else {
        matchCount++;
      }
    } else {
      notFound.push({ camion, produit, qte });
    }
  });

  renderResults(differences, notFound, matchCount);
}

function renderResults(differences, notFound, matchCount) {
  const totalMatched = differences.length + matchCount;

  // Summary
  summaryDiv.classList.remove('hidden');
  totalCountEl.textContent = totalMatched + notFound.length;
  matchCountEl.textContent = matchCount;
  diffCountEl.textContent = differences.length;
  notfoundCountEl.textContent = notFound.length;

  // Clear tables
  diffTbody.innerHTML = '';
  notfoundTbody.innerHTML = '';

  if (differences.length === 0 && notFound.length === 0) {
    noDiffMsg.classList.remove('hidden');
    resultsDiv.classList.add('hidden');
    notFoundSection.classList.add('hidden');
    return;
  }

  noDiffMsg.classList.add('hidden');

  // Differences table
  if (differences.length > 0) {
    resultsDiv.classList.remove('hidden');
    differences.forEach((d) => {
      const tr = document.createElement('tr');

      const tdCamion = document.createElement('td');
      tdCamion.textContent = d.camion;
      tr.appendChild(tdCamion);

      const tdProduit = document.createElement('td');
      tdProduit.textContent = d.produit;
      tr.appendChild(tdProduit);

      const tdQte = document.createElement('td');
      tdQte.textContent = d.qte.toLocaleString('fr-FR');
      tdQte.classList.add('val-diff');
      tr.appendChild(tdQte);

      const tdPoids = document.createElement('td');
      tdPoids.textContent = d.poidsNet.toLocaleString('fr-FR');
      tdPoids.classList.add('val-diff');
      tr.appendChild(tdPoids);

      const tdEcart = document.createElement('td');
      tdEcart.textContent = d.ecart.toLocaleString('fr-FR');
      tdEcart.classList.add(d.ecart > 0 ? 'ecart-pos' : 'ecart-neg');
      tr.appendChild(tdEcart);

      const tdStatus = document.createElement('td');
      tdStatus.textContent = '❌ Différent';
      tdStatus.classList.add('status-diff');
      tr.appendChild(tdStatus);

      diffTbody.appendChild(tr);
    });
  } else {
    resultsDiv.classList.add('hidden');
  }

  // Not found table
  if (notFound.length > 0) {
    notFoundSection.classList.remove('hidden');
    notFound.forEach((nf) => {
      const tr = document.createElement('tr');

      const tdCamion = document.createElement('td');
      tdCamion.textContent = nf.camion;
      tr.appendChild(tdCamion);

      const tdProduit = document.createElement('td');
      tdProduit.textContent = nf.produit;
      tr.appendChild(tdProduit);

      const tdQte = document.createElement('td');
      tdQte.textContent = nf.qte.toLocaleString('fr-FR');
      tr.appendChild(tdQte);

      notfoundTbody.appendChild(tr);
    });
  } else {
    notFoundSection.classList.add('hidden');
  }
}