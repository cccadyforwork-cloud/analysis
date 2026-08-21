import fs from 'node:fs';

const html = fs.readFileSync(process.argv[2], 'utf8');
const scripts = [...html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)]
  .filter((match) => !match[0].includes('application/json'))
  .map((match) => match[1])
  .join('\n');
new Function(scripts);
const payload = html.match(/<script id="wk34-embedded-data" type="application\/json">([\s\S]*?)<\/script>/);
if (!payload) throw new Error('embedded data block not found');
const data = JSON.parse(payload[1]);
const columns = Object.keys(data.masterRows?.[0] || {});
console.log(JSON.stringify({
  syntax: 'ok',
  masterRows: data.masterRows?.length || 0,
  hasOwner: columns.includes('负责人1（业绩归属人）'),
  hasFba: columns.includes('FBA可售'),
  has30DaySales: columns.includes('30日销量'),
  weekInputs: (html.match(/id="weekFiles"/g) || []).length,
  currentWeekInputs: (html.match(/id="currentWeekFile"/g) || []).length,
  amazonWeekInputs: (html.match(/id="amazonWeekFiles"/g) || []).length,
  internationalInputs: (html.match(/id="intlFile"/g) || []).length,
}));
