import { spawn } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { createDocxWithComment } from './fixtures/createDocx.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const sourceRoot = path.join(__dirname, '..');
const distRoot = path.join(__dirname, '../..');
const PROJECT_ROOT = existsSync(path.join(sourceRoot, 'package.json')) ? sourceRoot : distRoot;

const SERVER_PORT = 3456;
const SERVER_URL = `http://localhost:${SERVER_PORT}`;

async function sleep(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function main() {
  console.log('Starting Office Agent server in Office mode...');
  
  const serverProcess = spawn(
    'node',
    ['dist/src/api/server.js', '--office'],
    {
      cwd: PROJECT_ROOT,
      env: { ...process.env, PORT: String(SERVER_PORT), HOST: 'localhost' },
      stdio: 'pipe',
    }
  );

  console.log('Waiting for server to start...');
  await sleep(3000);

  try {
    console.log('\nTest 1: Health check...');
    const healthResponse = await fetch(`${SERVER_URL}/api/health`);
    const healthData = await healthResponse.json() as { status: string };
    console.log('Health status:', healthData.status);
    assert.strictEqual(healthData.status, 'ok');
    console.log('✓ Health check passed');

    console.log('\nTest 2: Task Pane HTML...');
    const taskpaneResponse = await fetch(`${SERVER_URL}/addin/taskpane.html`);
    const taskpaneContent = await taskpaneResponse.text();
    console.log('Contains Office Agent title:', taskpaneContent.includes('<title>Office Agent</title>'));
    assert.ok(taskpaneContent.includes('<title>Office Agent</title>'));
    console.log('✓ Task Pane HTML loaded');

    console.log('\nTest 3: Task Pane CSS...');
    const cssResponse = await fetch(`${SERVER_URL}/addin/taskpane.css`);
    const cssContent = await cssResponse.text();
    console.log('Contains CSS variables:', cssContent.includes('--color-primary'));
    assert.ok(cssContent.includes('--color-primary'));
    console.log('✓ Task Pane CSS loaded');

    console.log('\nTest 4: Task Pane JS...');
    const jsResponse = await fetch(`${SERVER_URL}/addin/taskpane.js`);
    const jsContent = await jsResponse.text();
    console.log('Contains Office.onReady:', jsContent.includes('Office.onReady'));
    assert.ok(jsContent.includes('Office.onReady'));
    console.log('✓ Task Pane JS loaded');

    console.log('\nTest 5: Commands HTML...');
    const commandsResponse = await fetch(`${SERVER_URL}/addin/commands.html`);
    console.log('Commands page status:', commandsResponse.status);
    assert.strictEqual(commandsResponse.status, 200);
    console.log('✓ Commands HTML loaded');

    console.log('\nTest 6: Manifest XML...');
    const manifestResponse = await fetch(`${SERVER_URL}/addin/manifest.xml`);
    const manifestContent = await manifestResponse.text();
    console.log('Contains OfficeApp:', manifestContent.includes('OfficeApp'));
    assert.ok(manifestContent.includes('OfficeApp'));
    console.log('✓ Manifest XML loaded');

    console.log('\nTest 7: Document inspect API...');
    const docxBuffer = await createDocxWithComment();
    const form = new FormData();
    form.append(
      'file',
      new Blob([new Uint8Array(docxBuffer)], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
      'å¹´åº¦æŠ¥å‘Š.docx'
    );
    const inspectResponse = await fetch(`${SERVER_URL}/api/document/inspect`, {
      method: 'POST',
      body: form,
    });
    const inspectData = await inspectResponse.json() as {
      fileType: string;
      filename: string;
      summary: { comments: number };
      comments: Array<{ textPreview: string }>;
    };
    console.log('Inspect status:', inspectResponse.status);
    assert.strictEqual(inspectResponse.status, 200);
    assert.strictEqual(inspectData.fileType, 'docx');
    assert.strictEqual(inspectData.filename, '年度报告.docx');
    assert.strictEqual(inspectData.summary.comments, 1);
    assert.ok(inspectData.comments[0].textPreview.includes('risk specific'));
    console.log('✓ Document inspect API loaded');

    console.log('\n✓ All tests passed!');

  } catch (error) {
    console.error('\n✗ Test failed:', error);
    process.exit(1);
  } finally {
    serverProcess.kill('SIGTERM');
  }
}

main().catch(console.error);
