import { diffChars } from 'diff';

export function changes(before, after) {
  const parts = diffChars(before, after);
  const edits = [];
  let offset = 0;
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    if (!part.added && !part.removed) { offset += part.value.length; continue; }
    const start = offset;
    let removed = '', inserted = '';
    while (i < parts.length && (parts[i].added || parts[i].removed)) {
      if (parts[i].removed) { removed += parts[i].value; offset += parts[i].value.length; }
      else inserted += parts[i].value;
      i++;
    }
    i--;
    edits.push({ start, end: offset, removed, inserted });
  }
  return { parts, edits };
}

export function validateResult(source, result) {
  if (!Array.isArray(result) || result.length !== source.length) throw new Error('模型返回的段落数量不一致，请重新生成。');
  return result.map((text, i) => {
    if (typeof text !== 'string' || /[\r\n\v\f\u0007]/u.test(text)) throw new Error('模型改变了段落结构，请重新生成。');
    if (source[i].trim() && !text.trim()) throw new Error('模型返回了空段落，已阻止清空原文。');
    if (text.length > Math.max(1000, source[i].length * 4)) throw new Error('模型扩写过多，请缩小范围后重试。');
    return text;
  });
}

export function characterMap(items, expected) {
  if (items.map(x => x.text).join('') !== expected) throw new Error('Word 字符定位与原文不一致，已停止写入。请改选普通正文。');
  const map = new Map();
  let offset = 0;
  for (const item of items) { map.set(offset, item); offset += item.text.length; }
  map.set(offset, null);
  return map;
}
