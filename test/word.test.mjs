import {test} from 'node:test';
import assert from 'node:assert/strict';
import {apply} from '../public/word.mjs';
globalThis.DOMParser=class { parseFromString(){return {querySelector:()=>null,getElementsByTagNameNS:()=>[]};} };
function host(text, fail=false) {
 const state={text,mode:'Off',writes:[],pending:[],fail};
 class Range {
  constructor(start,end){this.start=start;this.end=end;}
  get text(){return state.text.slice(this.start,this.end);}
  load(){return this;}
  getOoxml(){return {value:'<xml/>'};}
  getRange(where){return new Range(where==='Start'?this.start:this.end,where==='Start'?this.start:this.end);}
  expandTo(other){return new Range(Math.min(this.start,other.start),Math.max(this.end,other.end));}
  insertText(text,mode){assert.equal(mode,'Replace');assert.equal(state.mode,'TrackAll');state.pending.push({start:this.start,end:this.end,text});}
  search(){let offset=this.start;return {items:Array.from(this.text).map(ch=>{const r=new Range(offset,offset+ch.length);offset+=ch.length;return r;}),load(){}};}
 }
 const document={load(){},get changeTrackingMode(){return state.mode;},set changeTrackingMode(mode){state.mode=mode;}};
 const context={document,async sync(){if(state.fail&&state.pending.length){state.pending=[];throw new Error('host failure');}for(const op of state.pending){state.text=state.text.slice(0,op.start)+op.text+state.text.slice(op.end);state.writes.push(op);}state.pending=[];}};
 globalThis.Word={run:async(ranges,callback)=>callback(context)};
 return {state,snapshot:{ranges:[new Range(0,text.length)],texts:[text]}};
}
test('Word 适配层仅替换差异，倒序写入并恢复修订状态',async()=>{const {state,snapshot}=host('前面的文字，后面的文字。');assert.equal(await apply(snapshot,['前面文字，后面文字！']),3);assert.equal(state.text,'前面文字，后面文字！');assert.equal(state.mode,'Off');assert.ok(state.writes[0].start>state.writes[1].start);});
test('Word 适配层定位重复字与 emoji 后的插入',async()=>{const {state,snapshot}=host('😀数据数据');await apply(snapshot,['😀新数据数据！']);assert.equal(state.text,'😀新数据数据！');});
test('无差异不写入',async()=>{const {state,snapshot}=host('保持');assert.equal(await apply(snapshot,['保持']),0);assert.equal(state.writes.length,0);assert.equal(state.mode,'Off');});
test('原文变化后阻止写入',async()=>{const {state,snapshot}=host('原文');state.text='已改';await assert.rejects(apply(snapshot,['润色']),/原文已发生变化/);assert.equal(state.writes.length,0);});
test('宿主写入失败仍尝试恢复开关，并提示部分写入风险',async()=>{const {state,snapshot}=host('原文',true);await assert.rejects(apply(snapshot,['新文']),/可能已写入部分修订/);assert.equal(state.mode,'Off');});
