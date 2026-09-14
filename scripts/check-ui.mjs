import { chromium } from '@playwright/test';
import assert from 'node:assert/strict';
import { mkdir } from 'node:fs/promises';
const browser = await chromium.launch({channel:'chrome',headless:true});
try {
 const page = await browser.newPage({ignoreHTTPSErrors:true,viewport:{width:390,height:1000}});
 const errors=[];page.on('pageerror',e=>errors.push(e.message));
 await page.goto('https://localhost:3443/',{waitUntil:'domcontentloaded'});
 await page.getByRole('button',{name:'先看一个修订示例'}).click();
 await page.locator('#preview-section').waitFor({state:'visible'});
 assert.ok(await page.locator('#preview ins').count());assert.ok(await page.locator('#preview del').count());assert.ok(await page.locator('#preview span').count());
 assert.equal(await page.locator('#apply').isDisabled(),true);
 assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),true);
 await page.getByRole('button',{name:'仅纠错',exact:true}).click();
 assert.match(await page.locator('#instruction').inputValue(),/错别字/);
 assert.deepEqual(errors,[]);
 await mkdir('artifacts',{recursive:true});await page.screenshot({path:'artifacts/preview.png',fullPage:true});
 console.log('PASS：侧边栏、真实 diff 模块加载、三类差异、安全示例模式、窄屏布局与快捷要求；未模拟真实 Word 宿主。');
} finally {await browser.close();}
