/* ═══════════════════════════════════════════════════════════
   Advanced RAG Pipeline Explorer — Frontend Logic
   ═══════════════════════════════════════════════════════════ */

// ── State ──
let currentStep = 1;
let highestCompletedStep = 0;

// ── Navigation ──
function goToStep(step) {
  if (step > highestCompletedStep + 1 && step !== 1) return; // Prevent skipping ahead
  
  // Update stepper UI
  document.querySelectorAll('.step').forEach(s => {
    s.classList.remove('active');
    if (parseInt(s.dataset.step) <= highestCompletedStep) {
      s.classList.add('completed');
    }
  });
  const targetStep = document.getElementById(`step-${step}`);
  if (targetStep) targetStep.classList.add('active');

  // Update content UI
  document.querySelectorAll('.stage').forEach(s => s.classList.remove('active'));
  const targetStage = document.getElementById(`stage-${step}`);
  if (targetStage) targetStage.classList.add('active');
  
  currentStep = step;
}

function completeStep(step) {
  if (step > highestCompletedStep) {
    highestCompletedStep = step;
  }
  document.getElementById(`step-${step}`).classList.add('completed');
  const btnNext = document.getElementById(`btnNext${step}`);
  if (btnNext) btnNext.style.display = 'inline-flex';
}

// ── Helpers ──
function esc(t) { return (t || '').toString().replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function catCls(c) { return c === 'Positive' ? 'b-pos' : c === 'Negative' ? 'b-neg' : c === 'Edge Case' ? 'b-edge' : 'b-bnd'; }
function priCls(p) { return p === 'Critical' ? 'b-crit' : p === 'High' ? 'b-high' : p === 'Medium' ? 'b-med' : 'b-low'; }

// ── STEP 1: Upload ──
const uploadZone = document.getElementById('uploadZone');
const fileInput = document.getElementById('fileInput');

uploadZone.addEventListener('click', () => fileInput.click());
uploadZone.addEventListener('dragover', (e) => { e.preventDefault(); uploadZone.classList.add('dragover'); });
uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('dragover'));
uploadZone.addEventListener('drop', (e) => {
  e.preventDefault();
  uploadZone.classList.remove('dragover');
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', (e) => {
  if (e.target.files.length) handleFile(e.target.files[0]);
});

async function handleFile(file) {
  if (!file.name.match(/\.(xlsx|xls)$/)) {
    alert("Please upload an Excel file (.xlsx or .xls)");
    return;
  }

  document.getElementById('uploadInner').style.display = 'none';
  document.getElementById('uploadProgress').style.display = 'flex';
  
  const formData = new FormData();
  formData.append('file', file);

  try {
    const res = await fetch('/api/upload', { method: 'POST', body: formData });
    const data = await res.json();
    
    document.getElementById('uploadProgress').style.display = 'none';
    
    if (data.error) {
      alert(data.error);
      document.getElementById('uploadInner').style.display = 'block';
      return;
    }
    
    // Success
    uploadZone.style.display = 'none';
    document.getElementById('uploadResult').style.display = 'block';
    document.getElementById('uploadFileName').textContent = data.filename;
    
    document.getElementById('uploadStats').innerHTML = `
      <div class="stats-grid">
        <div class="stat-box">
          <div class="stat-val">${data.total_cases}</div>
          <div class="stat-lbl">Test Cases Parsed</div>
        </div>
        <div class="stat-box">
          <div class="stat-val">${Object.keys(data.modules).length}</div>
          <div class="stat-lbl">Modules Found</div>
        </div>
      </div>
    `;
    
    completeStep(1);
    
    // Pre-fetch chunking data for step 2
    fetchChunks();
    
  } catch (e) {
    alert("Upload failed: " + e.message);
    document.getElementById('uploadProgress').style.display = 'none';
    document.getElementById('uploadInner').style.display = 'block';
  }
}

// ── STEP 2: Chunking ──
async function fetchChunks() {
  try {
    const res = await fetch('/api/chunks');
    const data = await res.json();
    
    if (data.error) return;
    
    const summary = data.all_chunks_summary;
    document.getElementById('chunkingOverview').innerHTML = `
      <div class="result-card" style="margin-top:0">
        <div class="result-icon">✂️</div>
        <div class="result-body">
          <h3>Structured Chunking Applied</h3>
          <p style="font-size:13px;color:var(--text2);margin-bottom:12px;max-width:800px;">
            Instead of arbitrarily splitting text by character count, we mapped each Excel row to exactly <strong>1 test case chunk</strong>.
            This preserves the natural boundary and context of every test case.
          </p>
          <div class="stats-grid">
            <div class="stat-box"><div class="stat-val">${data.total}</div><div class="stat-lbl">Total Chunks</div></div>
            <div class="stat-box"><div class="stat-val">${summary.avg_length}</div><div class="stat-lbl">Avg Chars / Chunk</div></div>
            <div class="stat-box"><div class="stat-val">~${summary.total_tokens_approx}</div><div class="stat-lbl">Est. Total Tokens</div></div>
          </div>
        </div>
      </div>
    `;
    
    let gridHtml = '';
    data.chunks.forEach(c => {
      gridHtml += `
        <div class="tc-card">
          <div class="tc-top">
            <span class="badge b-id">${esc(c.tc_id)}</span>
            <span class="badge b-mod">${esc(c.module)}</span>
          </div>
          <div class="tc-desc">${esc(c.description)}</div>
          <div class="tc-text">${esc(c.text)}</div>
          <div class="tc-meta">
            <span>📏 ${c.length} chars</span>
            <span><span class="badge b-cat ${catCls(c.category)}">${esc(c.category)}</span></span>
          </div>
        </div>
      `;
    });
    
    document.getElementById('chunkGrid').innerHTML = gridHtml;
    completeStep(2);
    
  } catch (e) {
    console.error("Failed to fetch chunks:", e);
  }
}

// ── STEP 3: Embed & Store ──
async function embedAndStore() {
  const btn = document.getElementById('btnEmbed');
  btn.disabled = true;
  document.getElementById('embedAction').style.display = 'none';
  document.getElementById('embedProgress').style.display = 'flex';
  
  try {
    const res = await fetch('/api/embed-and-store', { method: 'POST' });
    const data = await res.json();
    
    document.getElementById('embedProgress').style.display = 'none';
    
    if (data.error) {
      alert(data.error);
      document.getElementById('embedAction').style.display = 'block';
      btn.disabled = false;
      return;
    }
    
    const vizHtml = data.sample_embedding ? `
      <div class="vector-viz">
        <strong>Sample Vector (first 10 dims):</strong><br>
        [${data.sample_embedding.first_10_dims.join(', ')} ...]<br>
        <br>
        <strong>Dimension:</strong> ${data.sample_embedding.dimension} | <strong>Norm:</strong> ${data.sample_embedding.norm}
      </div>
    ` : '';
    
    document.getElementById('embedResult').style.display = 'block';
    document.getElementById('embedResult').innerHTML = `
      <div class="result-card success">
        <div class="result-icon">✅</div>
        <div class="result-body" style="width:100%">
          <h3>Vector Database Ready</h3>
          <p style="font-size:13px;color:var(--text2);margin-bottom:12px;">
            Successfully generated ${data.points_count} dense vectors using ${data.embed_model} and stored them in ${data.vector_db}.
          </p>
          <div class="stats-grid" style="grid-template-columns: repeat(3, 1fr);">
            <div class="stat-box"><div class="stat-val">${data.points_count}</div><div class="stat-lbl">Vectors Stored</div></div>
            <div class="stat-box"><div class="stat-val">${data.embed_dim}</div><div class="stat-lbl">Dimensions</div></div>
            <div class="stat-box"><div class="stat-val">Cosine</div><div class="stat-lbl">Distance Metric</div></div>
          </div>
          ${vizHtml}
        </div>
      </div>
    `;
    
    completeStep(3);
    
  } catch (e) {
    alert("Embedding failed: " + e.message);
    document.getElementById('embedProgress').style.display = 'none';
    document.getElementById('embedAction').style.display = 'block';
    btn.disabled = false;
  }
}

// ── STEP 4: Search (Top 10) ──
async function performSearch() {
  const query = document.getElementById('searchQuery').value.trim();
  if (!query) { alert("Please enter a search query."); return; }
  
  const btn = document.getElementById('btnSearch');
  const ogText = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'Searching...';
  
  document.getElementById('searchResults').innerHTML = `
    <div class="embed-progress" style="padding:32px;">
      <div class="loader-ring"></div>
      <p style="margin-top:12px;color:var(--text2);font-size:13px;">Embedding query and searching Qdrant...</p>
    </div>
  `;
  
  try {
    const res = await fetch('/api/search', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ question: query })
    });
    const data = await res.json();
    
    if (data.error) {
      document.getElementById('searchResults').innerHTML = `<p style="color:var(--rose)">${esc(data.error)}</p>`;
      btn.disabled = false; btn.textContent = ogText;
      return;
    }
    
    let html = `
      <div style="margin-bottom:16px;font-size:14px;color:var(--text2);">
        Found <strong>${data.results.length}</strong> results using dense vector similarity (Cosine).
      </div>
      <div class="chunk-grid" style="grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));">
    `;
    
    data.results.forEach(r => {
      html += `
        <div class="tc-card">
          <div class="tc-top">
            <div style="display:flex;align-items:center;gap:8px;">
              <span class="badge b-rank">#${r.rank}</span>
              <span class="badge b-id">${esc(r.tc_id)}</span>
            </div>
            <span class="sim-badge sim-qdrant">${(r.similarity * 100).toFixed(1)}% Sim</span>
          </div>
          <div class="tc-desc" style="font-size:12px;">${esc(r.description || 'No description')}</div>
          <div class="tc-text" style="max-height:80px;">${esc(r.text)}</div>
        </div>
      `;
    });
    
    html += '</div>';
    document.getElementById('searchResults').innerHTML = html;
    
    completeStep(4);
    
  } catch (e) {
    document.getElementById('searchResults').innerHTML = `<p style="color:var(--rose)">Search failed: ${e.message}</p>`;
  }
  
  btn.disabled = false;
  btn.textContent = ogText;
}

// ── STEP 5: Re-rank (Top 5) ──
async function performRerank() {
  const btn = document.getElementById('btnRerank');
  btn.disabled = true;
  document.getElementById('rerankAction').style.display = 'none';
  
  document.getElementById('rerankResults').innerHTML = `
    <div class="embed-progress" style="padding:32px;">
      <div class="loader-ring"></div>
      <p style="margin-top:12px;color:var(--text2);font-size:13px;">Scoring (query, document) pairs with Cross-Encoder...</p>
    </div>
  `;
  
  try {
    const res = await fetch('/api/rerank', { method: 'POST' });
    const data = await res.json();
    
    if (data.error) {
      document.getElementById('rerankResults').innerHTML = `<p style="color:var(--rose)">${esc(data.error)}</p>`;
      document.getElementById('rerankAction').style.display = 'block';
      btn.disabled = false;
      return;
    }
    
    // Find rank changes
    const beforeIds = data.before.map(r => r.tc_id);
    
    let html = `
      <div class="result-card success" style="margin-top:0; margin-bottom:24px;">
        <div class="result-icon">🔄</div>
        <div class="result-body" style="width:100%">
          <h3>Re-ranking Complete</h3>
          <p style="font-size:13px;color:var(--text2);">
            The cross-encoder evaluated the true semantic relationship between your query and the top-10 retrieved docs, refining the list to the highly-relevant top-5.
          </p>
        </div>
      </div>
      
      <table class="comparison-table">
        <thead>
          <tr>
            <th style="width:50px">Shift</th>
            <th style="width:100px">Rank</th>
            <th style="width:120px">Test Case</th>
            <th>Snippet</th>
            <th style="width:120px">Vector Sim</th>
            <th style="width:120px">Rerank Score</th>
          </tr>
        </thead>
        <tbody>
    `;
    
    data.after.forEach(r => {
      const oldRank = beforeIds.indexOf(r.tc_id) + 1;
      const rankDiff = oldRank - r.rank; // Positive = went up
      
      let shiftHtml = '<span class="rank-shift shift-same">-</span>';
      if (rankDiff > 0) shiftHtml = '<span class="rank-shift shift-up">↑</span>';
      if (rankDiff < 0) shiftHtml = '<span class="rank-shift shift-down">↓</span>';
      
      const textPreview = r.text.length > 100 ? r.text.substring(0, 100) + '...' : r.text;
      
      html += `
        <tr>
          <td>${shiftHtml}</td>
          <td><span class="badge b-rank">#${r.rank}</span> <span style="font-size:10px;color:var(--text3)">(was #${oldRank})</span></td>
          <td><span class="badge b-id">${esc(r.tc_id)}</span></td>
          <td style="color:var(--text2);font-family:monospace;font-size:11px;">${esc(textPreview)}</td>
          <td><span class="sim-badge sim-qdrant">${(r.similarity_before * 100).toFixed(1)}%</span></td>
          <td><span class="sim-badge sim-rerank">${(r.reranker_score * 100).toFixed(1)}%</span></td>
        </tr>
      `;
    });
    
    html += `</tbody></table>`;
    document.getElementById('rerankResults').innerHTML = html;
    
    completeStep(5);
    
  } catch (e) {
    document.getElementById('rerankResults').innerHTML = `<p style="color:var(--rose)">Reranking failed: ${e.message}</p>`;
    document.getElementById('rerankAction').style.display = 'block';
    btn.disabled = false;
  }
}

// ── STEP 6: Ask Questions (Full Q&A) ──
async function askQuestion() {
  const query = document.getElementById('qaInput').value.trim();
  if (!query) return;
  
  const btn = document.getElementById('btnAsk');
  const resArea = document.getElementById('qaResults');
  btn.disabled = true;
  
  resArea.innerHTML = `
    <div class="qa-empty">
      <div class="loader-ring large"></div>
      <p>Running full pipeline: Search → Rerank → LLM generation...</p>
    </div>
  `;
  
  try {
    const res = await fetch('/api/ask', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ question: query })
    });
    const data = await res.json();
    
    if (data.error) {
      resArea.innerHTML = `<div class="qa-empty"><p style="color:var(--rose)">${esc(data.error)}</p></div>`;
      btn.disabled = false;
      return;
    }
    
    let html = `
      <div class="answer-box">
        <div class="answer-label">Groq Llama-3.3 Answer</div>
        <div class="answer-content">${formatAnswer(data.answer)}</div>
      </div>
      
      <div class="context-header">
        <span>Supporting Context (Top-5 Re-ranked)</span>
      </div>
      
      <div class="chunk-grid">
    `;
    
    data.retrieved_after.forEach(r => {
      html += `
        <div class="tc-card">
          <div class="tc-top">
            <span class="badge b-id">${esc(r.tc_id)}</span>
            <span class="badge b-mod">${esc(r.module)}</span>
          </div>
          <div class="tc-meta" style="margin-top:0;margin-bottom:8px;">
            <span>Sim: ${(r.similarity_before * 100).toFixed(1)}%</span>
            <span>Rerank: ${(r.similarity_after * 100).toFixed(1)}%</span>
          </div>
          <div class="tc-text" style="max-height:150px;">${esc(r.text)}</div>
        </div>
      `;
    });
    
    html += '</div>';
    resArea.innerHTML = html;
    
  } catch (e) {
    resArea.innerHTML = `<div class="qa-empty"><p style="color:var(--rose)">Query failed: ${e.message}</p></div>`;
  }
  
  btn.disabled = false;
}

function formatAnswer(ans) {
  // Simple markdown to HTML conversion for the answer text
  let html = esc(ans);
  // Bold
  html = html.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
  // Bullets
  html = html.replace(/^\s*-\s+(.*)$/gm, '• $1');
  return html;
}
