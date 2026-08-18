import fs from 'node:fs/promises';

const path = '/product/1945180142/';
const endpoint = `https://api.ozon.ru/composer-api.bx/page/json/v2?url=${encodeURIComponent(path)}`;
const headers = {
  'Accept': 'application/json; charset=utf-8',
  'Accept-Language': 'ru-RU,ru;q=0.9',
  'User-Agent': 'ozonapp_android/17.48.0+2528',
  'X-O3-App-Name': 'ozonapp_android',
  'X-O3-App-Version': '17.48.0(2528)',
  'X-O3-Device-Type': 'mobile',
  'X-O3-Sample-Trace': 'false',
  'Mobile-Lat': '0'
};

await fs.mkdir('output', { recursive: true });
const response = await fetch(endpoint, { headers });
const body = await response.text();
const result = {
  status: response.status,
  contentType: response.headers.get('content-type'),
  length: body.length,
  preview: body.slice(0, 1000)
};
console.log(JSON.stringify(result));
await fs.writeFile('output/mobile-probe.json', JSON.stringify(result, null, 2));
if (response.status !== 200) process.exitCode = 1;
