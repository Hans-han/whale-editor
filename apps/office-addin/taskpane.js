Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initializeApp();
  }
});

function initializeApp() {
  const taskInput = document.getElementById('task-input');
  const btnRun = document.getElementById('btn-run');
  const btnDownload = document.getElementById('btn-download');
  const status = document.getElementById('status');
  const manifestSection = document.querySelector('.manifest-section');
  const patchSection = document.querySelector('.patch-section');
  const validationSection = document.querySelector('.validation-section');
  const metricsSection = document.querySelector('.metrics-section');

  let modifiedFileBlob = null;

  taskInput.addEventListener('input', () => {
    btnRun.disabled = !taskInput.value.trim();
  });

  btnRun.addEventListener('click', async () => {
    const task = taskInput.value.trim();
    if (!task) return;

    await runAgent(task);
  });

  btnDownload.addEventListener('click', () => {
    if (modifiedFileBlob) {
      downloadFile(modifiedFileBlob, 'modified_document.docx');
    }
  });

  async function runAgent(task) {
    try {
      showStatus('loading', 'Extracting document content...');
      btnRun.disabled = true;
      btnRun.querySelector('.btn-text').hidden = true;
      btnRun.querySelector('.btn-loading').hidden = false;

      const documentBlob = await getDocumentAsBlob();
      
      showStatus('loading', 'Sending to Office Agent...');

      const formData = new FormData();
      formData.append('file', documentBlob, 'document.docx');
      formData.append('task', task);

      const response = await fetch('/api/agent/run', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error || 'Agent execution failed');
      }

      const contentType = response.headers.get('content-type');
      if (contentType && contentType.includes('application/json')) {
        const result = await response.json();
        showStatus('error', `Agent failed: ${result.summary}`);
        showManifest(result.manifest);
        showValidation(result.iterations[result.iterations.length - 1]?.validationReport);
        showMetrics(result.usage);
      } else {
        modifiedFileBlob = await response.blob();
        showStatus('success', 'Document modified successfully!');
        btnDownload.hidden = false;
      }

    } catch (error) {
      showStatus('error', `Error: ${error.message}`);
    } finally {
      btnRun.disabled = false;
      btnRun.querySelector('.btn-text').hidden = false;
      btnRun.querySelector('.btn-loading').hidden = true;
    }
  }

  async function getDocumentAsBlob() {
    return new Promise((resolve, reject) => {
      Office.context.document.getFileAsync(
        Office.FileType.Compressed,
        { sliceSize: 65536 },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            reject(new Error('Failed to get document'));
            return;
          }

          const file = result.value;
          const slices = [];
          let slicesReceived = 0;

          function getSlice(index) {
            file.getSliceAsync(index, (sliceResult) => {
              if (sliceResult.status === Office.AsyncResultStatus.Failed) {
                reject(new Error('Failed to get slice'));
                return;
              }

              const data = sliceResult.value.data;
              const bytes = data instanceof ArrayBuffer
                ? new Uint8Array(data)
                : Array.isArray(data)
                  ? new Uint8Array(data)
                  : typeof data === 'string'
                    ? Uint8Array.from(atob(data), (c) => c.charCodeAt(0))
                    : new Uint8Array(data);
              slices.push(bytes);
              slicesReceived++;

              if (slicesReceived === file.sliceCount) {
                file.closeAsync();
                const blob = new Blob(slices, {
                  type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                });
                resolve(blob);
              } else {
                getSlice(slicesReceived);
              }
            });
          }

          getSlice(0);
        }
      );
    });
  }

  function showStatus(type, message) {
    status.className = `status ${type}`;
    status.textContent = message;
  }

  function showManifest(manifest) {
    if (!manifest) return;
    manifestSection.hidden = false;
    document.getElementById('manifest-output').textContent = JSON.stringify(manifest, null, 2);
  }

  function showValidation(validation) {
    if (!validation) return;
    validationSection.hidden = false;
    const output = document.getElementById('validation-output');
    output.innerHTML = '';

    if (validation.isValid) {
      const item = document.createElement('div');
      item.className = 'validation-item success';
      item.textContent = '✓ Document is valid';
      output.appendChild(item);
    } else {
      validation.errors.forEach((error) => {
        const item = document.createElement('div');
        item.className = 'validation-item error';
        item.textContent = `✗ ${error}`;
        output.appendChild(item);
      });
    }
  }

  function showMetrics(usage) {
    if (!usage) return;
    metricsSection.hidden = false;
    const output = document.getElementById('metrics-output');
    output.innerHTML = '';

    const metrics = [
      { label: 'Input Tokens', value: usage.inputTokens },
      { label: 'Output Tokens', value: usage.outputTokens },
      { label: 'Cache Hit', value: usage.cacheHitTokens ?? 'N/A' },
      { label: 'Cache Miss', value: usage.cacheMissTokens ?? 'N/A' },
    ];

    metrics.forEach((metric) => {
      const item = document.createElement('div');
      item.className = 'metric-item';
      item.innerHTML = `
        <div class="metric-label">${metric.label}</div>
        <div class="metric-value">${metric.value}</div>
      `;
      output.appendChild(item);
    });
  }

  function downloadFile(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }
}
