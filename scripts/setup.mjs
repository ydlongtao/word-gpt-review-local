import { mkdir, copyFile, access, chmod } from 'node:fs/promises';
import { homedir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { ensureCertificatesAreInstalled } from 'office-addin-dev-certs';
const root = fileURLToPath(new URL('../',import.meta.url));
process.chdir(root);
try { await access('.env'); } catch { await copyFile('.env.example','.env'); await chmod('.env',0o600); }
if (process.platform === 'darwin') {
  const target = `${homedir()}/Library/Containers/com.microsoft.Word/Data/Documents/wef`;
  await mkdir(target,{recursive:true});
  await copyFile('manifest.xml',`${target}/word-gpt-review.xml`);
  console.log('插件清单已安装。填写 .env 后运行 npm start，再在 Word“主页 → 加载项”中打开 GPT 本地审阅。若未出现，请保存文档后重新打开 Word。');
} else console.log('证书已配置。请按照 README 中的 Windows / 网页版说明旁加载 manifest.xml。');

await ensureCertificatesAreInstalled();
