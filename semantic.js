// Only interchangeable instrumentation terms are normalized for automatic pricing.
// Quantities, punctuation, ratings, materials and scope qualifiers remain in the key.
const equivalentTerms = [
  [/\b(?:instal\.|installing|installation|install)\s*(?:of\b)?/g, 'install '],
  [/\b(?:calibrating|calibration|calibrate)\s*(?:of\b)?/g, 'calibrate '],
  [/\b(?:testing|test)\s*(?:of\b)?/g, 'test '],
  [/\b(?:laying of|laying|lay)\b/g, 'lay'],
  [/\b(?:cables|cable)\b/g, 'cable'],
  [/\b(?:junction boxes|junction box|jbs|jb)\b/g, 'junction box'],
  [/\b(?:weather proof|weather-proof|weatherproof|wp)\b/g, 'weatherproof'],
  [/\b(?:ex-proof|ex proof|exproof)\b/g, 'ex-proof'],
  [/\b(?:up to|upto)\b/g, 'up to'],
  [/\b(?:supply and installation|supply & installation|supply and install|supply & install)\b/g, 'supply and install'],
  [/\b(?:stainless steel|ss)\b/g, 'stainless steel'],
  [/\b(?:outer diameter|od)\b/g, 'outer diameter']
];

function semanticKey(value) {
  let phrase = norm(value);
  for (const [pattern, term] of equivalentTerms) phrase = phrase.replace(pattern, term);
  return phrase.replace(/\s+/g, ' ').trim();
}

const commonWords = new Set(['a', 'an', 'and', 'for', 'from', 'in', 'of', 'on', 'the', 'to', 'with']);
const termsFor = text => new Set((semanticKey(text).match(/[\p{L}\p{N}]+(?:[.-][\p{L}\p{N}]+)*/gu) || [])
  .filter(word => !commonWords.has(word)));

function createCandidateSearch(sources) {
  const documents = sources.map(source => ({source, terms:termsFor([source.short, source.description].filter(present).join(' '))}));
  const counts = new Map();
  for (const {terms} of documents) for (const word of terms) counts.set(word, (counts.get(word) || 0) + 1);
  const weight = word => 1 + Math.log((documents.length + 1) / ((counts.get(word) || 0) + 1));
  const magnitude = terms => Math.sqrt([...terms].reduce((sum, word) => sum + weight(word) ** 2, 0));
  for (const document of documents) document.length = magnitude(document.terms);
  return (description, unit) => {
    const query = termsFor(description);
    if (!query.size) return [];
    const length = magnitude(query);
    return documents.flatMap(({source, terms, length: docLength}) => {
      if (!docLength || conversion(source.unit, unit) === null) return [];
      const shared = [...query].filter(word => terms.has(word));
      const score = shared.reduce((sum, word) => sum + weight(word) ** 2, 0) / (length * docLength);
      return shared.length >= 2 && score >= .2 ? [{source, score}] : [];
    }).sort((a, b) => b.score - a.score).slice(0, 5);
  };
}
