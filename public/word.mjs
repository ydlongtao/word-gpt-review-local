import { changes, characterMap } from './engine.mjs';
const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function inspectXml(value) {
  const xml = new DOMParser().parseFromString(value, 'application/xml');
  if (xml.querySelector('parsererror')) throw new Error('无法解析 Word 内容。');
  const unsafe = ['ins','del','moveFrom','moveTo','fldChar','fldSimple','drawing','pict','footnoteReference','endnoteReference','sdt','hyperlink','commentRangeStart','commentRangeEnd'];
  if (unsafe.some(tag => xml.getElementsByTagNameNS(NS,tag).length)) throw new Error('范围内含已有修订、批注、链接、域或嵌入对象。请先选择普通正文，或处理已有修订。');
}
export async function release(snapshot) {
  if (!snapshot) return;
  await Word.run(snapshot.ranges, async context => {
    snapshot.ranges.forEach(r => r.untrack()); await context.sync();
  });
}
export async function capture(scope) {
  return Word.run(async context => {
    const root = scope === 'body' ? context.document.body.getRange() : context.document.getSelection();
    const paragraphs = root.paragraphs;
    root.load('text'); paragraphs.load('items'); await context.sync();
    if (!root.text.trim()) throw new Error('请先在 Word 中选中文字。');
    if (root.text.length > 12000 || paragraphs.items.length > 80) throw new Error('每次最多 12000 字符、80 段，请缩小选区。');
    const ranges = paragraphs.items.map(p => p.getRange('Content').intersectWith(root));
    const xmls = ranges.map(r => { r.load('text'); return r.getOoxml(); });
    await context.sync();
    xmls.forEach(x => inspectXml(x.value));
    const selected = ranges.filter(r => r.text.trim());
    if (!selected.length) throw new Error('没有可润色的正文。');
    for (const r of selected) {
      if (/[\r\n\v\f\u0007]/u.test(r.text)) throw new Error('暂不支持段内手动换行或特殊分隔符，请改选普通段落。');
      r.track();
    }
    await context.sync();
    return { ranges: selected, texts: selected.map(r => r.text) };
  });
}
export async function apply(snapshot, result) {
  return Word.run(snapshot.ranges, async context => {
    snapshot.ranges.forEach(r => r.load('text'));
    context.document.load('changeTrackingMode');
    const xmls = snapshot.ranges.map(r => r.getOoxml());
    await context.sync();
    snapshot.ranges.forEach((r,i) => {
      if (r.text !== snapshot.texts[i]) throw new Error('原文已发生变化，请重新生成建议。');
      inspectXml(xmls[i].value);
    });
    const plans = snapshot.ranges.map((r,i) => ({ range:r, ...changes(snapshot.texts[i],result[i]) }));
    const changed = plans.filter(p => p.edits.length);
    if (!changed.length) return 0;
    for (const p of changed) { p.characters = p.range.search('?',{matchWildcards:true}); p.characters.load('items/text'); }
    await context.sync();
    const operations = [];
    for (const p of changed) {
      const map = characterMap(p.characters.items,p.range.text);
      for (const edit of p.edits) {
        if (!map.has(edit.start) || !map.has(edit.end)) throw new Error('修改位置跨越特殊字符边界，已停止写入。');
        const start = map.get(edit.start)?.getRange('Start') || p.range.getRange('End');
        const end = map.get(edit.end)?.getRange('Start') || p.range.getRange('End');
        operations.push({range: start.expandTo(end), text: edit.inserted});
      }
    }
    // 在启用修订之前解析全部范围，尽量提前捕获定位错误。
    operations.forEach(op => op.range.load('text'));
    await context.sync();
    const previous = context.document.changeTrackingMode;
    let count = 0;
    let failure;
    try {
      context.document.changeTrackingMode = 'TrackAll';
      await context.sync();
      context.document.load('changeTrackingMode'); await context.sync();
      if (context.document.changeTrackingMode !== 'TrackAll') throw new Error('无法启用 Word 修订。');
      // 倒序执行，避免前面的编辑改变后面原文的位置。
      for (const op of operations.reverse()) { op.range.insertText(op.text,'Replace'); }
      await context.sync(); count = operations.length;
    } catch (error) {
      failure = new Error(`写入未完整完成：${error.message}。Word 可能已写入部分修订，请检查“审阅”窗格，必要时撤销；不要直接重试。`);
    } finally {
      try { context.document.changeTrackingMode = previous; await context.sync(); }
      catch { failure = new Error('无法确认修订模式已恢复。请在 Word“审阅”中检查修订内容和开关状态。'); }
    }
    if (failure) throw failure;
    return count;
  });
}
