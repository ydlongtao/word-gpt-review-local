import { changes, validateResult } from './engine.mjs';
import { capture, apply, release } from './word.mjs';
const $ = id => document.getElementById(id);
let config, ready = false, snapshot, result, busy = false;
function status(text, error = false) { $('status').textContent = text; $('status').classList.toggle('error',error); }
function lock(value) { busy = value; $('generate').disabled = value || !ready || !config?.configured; $('apply').disabled = value || !snapshot || !result; $('demo').disabled = value; $('scope').disabled = value; }
function preview(source, output) {
  $('preview').replaceChildren(); let count = 0;
  source.forEach((text,i) => {
    const { parts, edits } = changes(text,output[i]); count += edits.length;
    const p = document.createElement('p');
    for (const part of parts) { const node = document.createElement(part.added ? 'ins' : part.removed ? 'del' : 'span'); node.textContent = part.value; p.append(node); }
    $('preview').append(p);
  });
  $('count').textContent = `${count} 处修改`; $('preview-section').hidden = false;
  return count;
}
async function clear() { const previous = snapshot; snapshot = null; result = null; if(previous) await release(previous); }
$('generate').onclick = async () => {
  if (busy) return; lock(true);
  try {
    await clear(); $('preview-section').hidden = true;
    status('正在读取 Word 正文…'); snapshot = await capture($('scope').value);
    status(`正在润色 ${snapshot.texts.length} 段文字，请稍候…`);
    const response = await fetch('/api/rewrite',{method:'POST',headers:{'Content-Type':'application/json','X-Review-Token':config.token},body:JSON.stringify({paragraphs:snapshot.texts,instruction:$('instruction').value}),signal:AbortSignal.timeout(130000)});
    const data = await response.json(); if(!response.ok) throw new Error(data.error);
    result = validateResult(snapshot.texts,data.paragraphs);
    const count = preview(snapshot.texts,result);
    status(count ? '建议已生成。检查差异后，点击下方按钮写入 Word。' : '这次没有需要修改的文字。');
    if(!count) result = null;
  } catch(error) { status(error.message,true); result = null; }
  finally {lock(false);}
};
$('apply').onclick = async () => {
  if(busy || !snapshot || !result) return; lock(true);
  try { status('正在以 Word 原生修订写入…'); const count = await apply(snapshot,result); status(`已写入 ${count} 处修改。请在 Word“审阅 → 所有标记”中查看、接受或拒绝。`); }
  catch(error) { status(error.message,true); }
  finally { try { await clear(); } catch {} lock(false); }
};
$('demo').onclick = async () => {
  if(busy) return; lock(true);
  try { await clear(); preview(['本研究的目的是为了探索这一问题。我们对数据进行了分析。'],['本研究旨在探索这一问题。我们分析了数据。']); status('这是内置示例，没有调用模型，也没有修改 Word 文档。'); }
  catch(error) { status(error.message,true); } finally {lock(false);}
};
document.querySelectorAll('[data-prompt]').forEach(button => button.onclick = () => {$('instruction').value = button.dataset.prompt;});
try { const response = await fetch('/api/config'); if(!response.ok) throw new Error('本地服务连接失败'); config = await response.json(); $('connection').textContent = config.configured ? `${config.model} · ${config.endpoint}` : '本地服务已连接 · 尚未配置模型'; $('light').classList.add('ready'); }
catch(error) {status(error.message,true);}
if (window.Office) Office.onReady(info => {
  ready = info.host === Office.HostType.Word && Office.context.requirements.isSetSupported('WordApi','1.4');
  if(!ready) status('请在支持 WordApi 1.4 的 Word 中打开插件。浏览器中可查看修订示例。');
  else status(config?.configured ? '在 Word 中选择正文，然后生成修改建议。' : 'Word 已连接。请先配置 .env 中的模型信息，或查看修订示例。');
  lock(false);
});
