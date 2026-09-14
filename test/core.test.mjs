import { test } from 'node:test';
import assert from 'node:assert/strict';
import { changes, characterMap, validateResult } from '../public/engine.mjs';
import { rewrite } from '../model.mjs';
function execute(before, edits) { for (const e of [...edits].reverse()) { assert.equal(before.slice(e.start,e.end),e.removed); before = before.slice(0,e.start)+e.inserted+before.slice(e.end); } return before; }
for (const [name,a,b] of [
  ['未改变','保持原样。','保持原样。'], ['中文','本研究的目的是为了探索这一问题。','本研究旨在探索这一问题。'],
  ['重复文字','数据数据，数据。','数据，新的数据。'], ['首尾插入','正文','前言正文结尾'],
  ['纯删除','太过冗余的文字','文字'], ['emoji','你好😀 世界','你好🙂，世界'],
  ['组合字符','Cafe\u0301','Café'], ['英文空格','This is  an test.','This is a test.'], ['XML字符','a < b & c','a ≤ b & c']
]) test(name,() => { const plan = changes(a,b); assert.equal(execute(a,plan.edits),b); assert.equal(plan.parts.filter(p=>!p.added).map(p=>p.value).join(''),a); assert.equal(plan.parts.filter(p=>!p.removed).map(p=>p.value).join(''),b); if(a===b) assert.equal(plan.edits.length,0); });
test('随机差异可逆，保证偏移正确',()=>{let seed=123; const rand=()=>((seed=(seed*16807)%2147483647)%8); const chars=['中','文','a',' ','😀','。','?','b']; for(let i=0;i<200;i++){const a=Array.from({length:20},()=>chars[rand()]).join('');const b=Array.from({length:20},()=>chars[rand()]).join('');assert.equal(execute(a,changes(a,b).edits),b);}});
test('模型结构校验',()=>{for(const out of [[],[''],['a\nb'],[null],['a','b']]) assert.throws(()=>validateResult(['原文'],out));});
test('定位偏移按 UTF-16，防止 emoji 后错位',()=>{const m=characterMap([{text:'😀'},{text:'中'}],'😀中');assert.ok(m.has(2));assert.ok(!m.has(1));assert.throws(()=>characterMap([{text:'中'}],'中文'));});
const cfg={model:'test',baseUrl:'https://example.com/v1',key:'test'};
test('模型接口正确构造请求并验证返回',async()=>{let sent;const output=await rewrite(['原文'],'纠错',cfg,async(url,options)=>{sent={url,options};return {ok:true,json:async()=>({choices:[{message:{content:'{"paragraphs":["润色"]}'},finish_reason:'stop'}]})};});assert.deepEqual(output,['润色']);assert.equal(sent.url,'https://example.com/v1/chat/completions');assert.equal(sent.options.headers.Authorization,'Bearer test');assert.equal(JSON.parse(sent.options.body).messages.length,2);});
test('拒绝接口错误、截断、错误 JSON',async()=>{for(const response of [{ok:false,status:401},{ok:true,json:async()=>({choices:[{finish_reason:'length'}]})},{ok:true,json:async()=>({choices:[{message:{content:'hello'}}]})}])await assert.rejects(rewrite(['正文'],'润色',cfg,async()=>response));});
test('拒绝无模型、超限、远程明文传输',async()=>{await assert.rejects(rewrite(['正文'],'润色',{...cfg,model:''}));await assert.rejects(rewrite(['a'.repeat(12001)],'润色',cfg));await assert.rejects(rewrite(['正文'],'润色',{...cfg,baseUrl:'http://example.com/v1'}));});
