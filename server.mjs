import https from 'node:https';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { randomBytes, timingSafeEqual } from 'node:crypto';
import { homedir } from 'node:os';
import { rewrite } from './model.mjs';
const port = Number(process.env.PORT || 3443);
const origin = `https://localhost:${port}`;
const token = randomBytes(32).toString('hex');
const config = { baseUrl: process.env.MODEL_BASE_URL || 'https://api.openai.com/v1', model: process.env.MODEL_NAME || '', key: process.env.MODEL_API_KEY || '' };
const routes = { '/icon-32.png':['public/icon-32.png','image/png'], '/icon-64.png':['public/icon-64.png','image/png'], '/': ['public/index.html','text/html'], '/app.mjs': ['public/app.mjs','text/javascript'], '/word.mjs': ['public/word.mjs','text/javascript'], '/engine.mjs': ['public/engine.mjs','text/javascript'], '/style.css': ['public/style.css','text/css'], '/vendor/diff.js': ['node_modules/diff/libesm/index.js','text/javascript'] };
const certDir = `${homedir()}/.office-addin-dev-certs`;
let tls;
try { tls = { key: await readFile(`${certDir}/localhost.key`), cert: await readFile(`${certDir}/localhost.crt`) }; }
catch { console.error('缺少本地 HTTPS 证书，请先运行 npm run setup。'); process.exit(1); }
https.createServer(tls, async (req, res) => {
  const json = (status, body) => { res.writeHead(status, {'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'}); res.end(JSON.stringify(body)); };
  res.setHeader('X-Content-Type-Options','nosniff');
  res.setHeader('Cache-Control','no-store');
  // localhost Host 校验防止 DNS rebinding；敏感操作仅允许同源页面加一次性进程令牌。
  if (req.headers.host !== `localhost:${port}`) return json(403, {error:'不允许的 Host'});
  if (req.headers.origin && req.headers.origin !== origin) return json(403, {error:'不允许的来源'});
  try {
    if (req.method === 'GET' && req.url === '/api/config') return json(200, { token, configured: !!config.model, model: config.model, endpoint: new URL(config.baseUrl).origin });
    if (req.method === 'POST' && req.url === '/api/rewrite') {
      const supplied = String(req.headers['x-review-token'] || '');
      if (req.headers.origin !== origin || supplied.length !== token.length || !timingSafeEqual(Buffer.from(supplied),Buffer.from(token))) return json(403,{error:'请重新打开插件'});
      if (!req.headers['content-type']?.startsWith('application/json')) return json(415,{error:'仅支持 JSON'});
      let body = ''; let size = 0;
      for await (const chunk of req) { size += chunk.length; if (size > 150000) { json(413,{error:'请求过大'}); return; } body += chunk; }
      const { paragraphs, instruction } = JSON.parse(body);
      return json(200, {paragraphs: await rewrite(paragraphs, instruction, config)});
    }
    let route = routes[req.url];
    // diff 的 ESM 内部模块；只开放 npm 包内的 JS 文件。
    if (!route && /^\/vendor\/[a-zA-Z0-9_/-]+\.js$/.test(req.url)) route = ['node_modules/diff/libesm/' + req.url.slice('/vendor/'.length), 'text/javascript'];
    if (req.method !== 'GET' || !route) return json(404,{error:'未找到'});
    const content = await readFile(fileURLToPath(new URL(route[0], import.meta.url)));
    res.writeHead(200,{'Content-Type':route[1]+'; charset=utf-8'}); res.end(content);
  } catch (error) { json(400,{error:error.name === 'TimeoutError' ? '模型请求超时，请重试。' : error.message}); }
}).listen(port, '127.0.0.1', () => console.log(`Word GPT 审阅已启动：${origin}\n模型：${config.model || '尚未配置；可先打开演示预览'}`));
