import { execSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const PROJECT_ROOT = path.join(__dirname, '../..');

const CERT_DIR = path.join(PROJECT_ROOT, 'certs');
const DOMAIN = 'localhost';

function run(cmd: string): string {
  try {
    return execSync(cmd, { encoding: 'utf-8', stdio: 'pipe' });
  } catch (e: any) {
    return e.stderr || e.stdout || '';
  }
}

function main() {
  console.log('=== Office Add-in 证书生成工具 ===\n');

  if (!fs.existsSync(CERT_DIR)) {
    fs.mkdirSync(CERT_DIR, { recursive: true });
  }

  const keyPath = path.join(CERT_DIR, 'localhost.key');
  const certPath = path.join(CERT_DIR, 'localhost.crt');

  if (fs.existsSync(keyPath) && fs.existsSync(certPath)) {
    console.log('✓ 证书已存在，跳过生成');
    console.log(`  密钥: ${keyPath}`);
    console.log(`  证书: ${certPath}`);
    return;
  }

  console.log('生成自签名证书...');
  
  const opensslCmd = `openssl req -x509 -newkey rsa:2048 -keyout "${keyPath}" -out "${certPath}" -days 365 -nodes -subj "/CN=${DOMAIN}"`;
  
  const result = run(opensslCmd);
  
  if (fs.existsSync(keyPath) && fs.existsSync(certPath)) {
    console.log('✓ 证书生成成功');
    console.log(`  密钥: ${keyPath}`);
    console.log(`  证书: ${certPath}`);
    console.log('\n下一步:');
    console.log('1. 双击 certs/localhost.crt 安装证书');
    console.log('2. 选择"将所有证书放入下列存储"');
    console.log('3. 选择"受信任的根证书颁发机构"');
    console.log('4. 运行 npm run dev:office-https');
  } else {
    console.log('✗ 证书生成失败');
    console.log(result);
    console.log('\n请确保已安装 OpenSSL:');
    console.log('  macOS: brew install openssl');
    console.log('  Windows: https://slproweb.com/products/Win32OpenSSL.html');
  }
}

main();
