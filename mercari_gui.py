from flask import Flask, request, jsonify
import sys
import os
import io
import requests as http_requests
import openpyxl
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'temp_download', 'mercari-2.2.1'))
from mercari import search, getItemInfo, MercariSort, MercariOrder, MercariSearchStatus

app = Flask(__name__)

@app.after_request
def add_cors_headers(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Headers'] = '*'
    return response

HTML_PAGE = u'''<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Mercari JP Search</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; }
  body { background: #f0f2f5; display: flex; justify-content: center; min-height: 100vh; }
  .container { max-width: 1000px; width: 100%; padding: 24px 16px; }
  h1 { font-size: 24px; color: #ea352d; margin-bottom: 20px; display: flex; align-items: center; gap: 8px; }
  h1::before { content: "\\01F4F1"; font-size: 24px; }
  .card { background: white; border-radius: 12px; padding: 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin-bottom: 16px; }
  .card h2 { font-size: 16px; color: #333; margin-bottom: 16px; }
  .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 12px; }
  .form-group { display: flex; flex-direction: column; }
  .form-group.full { grid-column: 1 / -1; }
  label { font-size: 12px; color: #666; margin-bottom: 4px; font-weight: 500; }
  input, select { padding: 10px 12px; border: 1px solid #ddd; border-radius: 8px; font-size: 14px; outline: none; transition: border-color 0.2s; background: #fafafa; }
  input:focus, select:focus { border-color: #ea352d; background: white; }
  .btn-search { width: 100%; padding: 12px; background: #ea352d; color: white; border: none; border-radius: 8px; font-size: 16px; font-weight: 600; cursor: pointer; transition: background 0.2s; display: flex; align-items: center; justify-content: center; gap: 8px; }
  .btn-search:hover { background: #d42c25; }
  .btn-search:disabled { background: #ccc; cursor: not-allowed; }
  .btn-search .spinner { display: none; width: 18px; height: 18px; border: 2px solid rgba(255,255,255,0.3); border-top-color: white; border-radius: 50%; animation: spin 0.6s linear infinite; }
  .btn-search.loading .spinner { display: inline-block; }
  .btn-search.loading .btn-text { display: none; }
  @keyframes spin { to { transform: rotate(360deg); } }
  .results { display: grid; gap: 12px; }
  .result-card { display: flex; gap: 16px; padding: 16px; background: white; border: 1px solid #eee; border-radius: 10px; transition: box-shadow 0.2s; align-items: center; overflow: hidden; }
  .result-card:hover { box-shadow: 0 2px 12px rgba(0,0,0,0.1); }
  .result-icon { width: 48px; height: 48px; background: #f5f5f5; border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 24px; flex-shrink: 0; }
  .result-info { flex: 1; min-width: 0; }
  .result-name { font-size: 14px; font-weight: 500; color: #333; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; min-width: 0; }
  .result-price { font-size: 18px; font-weight: 700; color: #ea352d; margin-top: 4px; }
  .result-price span { font-size: 12px; font-weight: 400; color: #999; }
  .result-link { padding: 8px 16px; background: #f0f2f5; border-radius: 6px; text-decoration: none; color: #333; font-size: 13px; font-weight: 500; transition: background 0.2s; white-space: nowrap; }
  .result-link:hover { background: #ea352d; color: white; }
  .status { text-align: center; padding: 32px; color: #999; }
  .status .emoji { font-size: 48px; margin-bottom: 12px; }
  .status p { font-size: 14px; }
  .error-msg { background: #fff0f0; border: 1px solid #ffd4d4; color: #d42c25; padding: 12px 16px; border-radius: 8px; font-size: 14px; margin-bottom: 12px; display: none; }
  .count-badge { font-size: 13px; color: #666; margin-bottom: 12px; }
  .count-badge strong { color: #333; }
  .modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 1000; justify-content: center; align-items: center; }
  .modal-overlay.active { display: flex; }
  .modal { background: white; border-radius: 12px; padding: 24px; max-width: 600px; width: 90%; max-height: 80vh; overflow-y: auto; box-shadow: 0 4px 24px rgba(0,0,0,0.2); position: relative; }
  .modal h3 { font-size: 18px; color: #333; margin-bottom: 8px; padding-right: 32px; }
  .modal .modal-meta { font-size: 12px; color: #999; margin-bottom: 16px; }
  .modal .modal-desc { font-size: 14px; color: #444; line-height: 1.7; white-space: pre-wrap; word-break: break-word; }
  .modal .modal-close { position: absolute; top: 16px; right: 16px; background: none; border: none; font-size: 24px; cursor: pointer; color: #999; padding: 4px 8px; border-radius: 4px; }
  .modal .modal-close:hover { background: #f0f2f5; color: #333; }
  .modal-loading { text-align: center; padding: 32px; color: #999; }
  @keyframes modalFadeIn { from { opacity: 0; transform: scale(0.95); } to { opacity: 1; transform: scale(1); } }
  .modal { animation: modalFadeIn 0.2s ease-out; }
  .result-link.desc-btn { background: #ea352d; color: white; margin-right: 6px; }
  .result-link.desc-btn:hover { background: #d42c25; }
  .result-actions { display: flex; gap: 6px; flex-shrink: 0; }
  .result-tag { display: inline-block; font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 4px; margin-top: 4px; text-transform: uppercase; letter-spacing: 0.5px; }
  .result-tag.sold { background: #ea352d; color: white; }
  .result-tag.trading { background: #f5a623; color: white; }
  .result-tag.auction { background: #9b59b6; color: white; }
  .result-tag.other { background: #999; color: white; }
  .result-updated { font-size: 11px; color: #888; margin-top: 2px; }
  .result-condition { font-size: 11px; color: #555; margin-top: 1px; font-style: italic; }
  .result-seller { font-size: 11px; color: #888; margin-top: 1px; }
  .result-info { display: flex; flex-direction: column; gap: 2px; }
  .tabs { display: flex; gap: 0; margin-bottom: 16px; border-bottom: 2px solid #ddd; }
  .tab-btn { padding: 10px 20px; background: none; border: none; border-bottom: 2px solid transparent; margin-bottom: -2px; font-size: 14px; font-weight: 500; color: #999; cursor: pointer; transition: all 0.2s; }
  .tab-btn:hover { color: #333; }
  .tab-btn.active { color: #ea352d; border-bottom-color: #ea352d; }
  .tab-content { display: none; }
  .tab-content.active { display: block; }
  .bulk-area { display: flex; flex-direction: column; gap: 12px; }
  .bulk-area textarea { width: 100%; min-height: 120px; padding: 10px 12px; border: 1px solid #ddd; border-radius: 8px; font-size: 13px; outline: none; resize: vertical; font-family: inherit; }
  .bulk-area textarea:focus { border-color: #ea352d; }
  .file-upload { border: 2px dashed #ddd; border-radius: 8px; padding: 20px; text-align: center; cursor: pointer; transition: all 0.2s; position: relative; }
  .file-upload:hover { border-color: #ea352d; background: #fff5f5; }
  .file-upload input[type="file"] { position: absolute; top: 0; left: 0; width: 100%; height: 100%; opacity: 0; cursor: pointer; }
  .file-upload .file-label { font-size: 13px; color: #999; pointer-events: none; }
  .file-upload .file-label strong { color: #ea352d; }
  .bulk-keyword-count { font-size: 12px; color: #999; }
  .bulk-result-group { margin-bottom: 16px; border: 1px solid #eee; border-radius: 10px; overflow: hidden; }
  .bulk-result-header { padding: 10px 14px; background: #f9f9f9; font-size: 14px; font-weight: 600; color: #333; cursor: pointer; display: flex; justify-content: space-between; align-items: center; }
  .bulk-result-header:hover { background: #f0f0f0; }
  .bulk-result-header .count { font-size: 12px; font-weight: 400; color: #999; }
  .bulk-result-body { padding: 8px 14px 14px; }
  .bulk-result-body .result-card { margin-bottom: 8px; }
  .bulk-result-body .result-card:last-child { margin-bottom: 0; }
  .bulk-result-body .result-card .result-link { font-size: 12px; padding: 6px 12px; }
  .bulk-result-body .result-card .result-name { font-size: 13px; }
  .dl-btn { padding: 8px 16px; background: #333; color: white; border: none; border-radius: 6px; font-size: 13px; cursor: pointer; margin-top: 8px; }
  .dl-btn:hover { background: #555; }
  @media (max-width: 600px) { .form-row { grid-template-columns: 1fr; } }
</style>
</head>
<body>
<div class="container">
  <h1>Mercari JP Search</h1>
  <div id="errorMsg" class="error-msg"></div>
  <div class="tabs">
    <button class="tab-btn active" onclick="switchTab('simple')">Simple Search</button>
    <button class="tab-btn" onclick="switchTab('bulk')">Bulk Search</button>
  </div>
  <div class="tab-content active" id="tabSimple">
  <div class="card">
    <div class="form-row">
      <div class="form-group full">
        <label for="keyword">Keyword *</label>
        <input type="text" id="keyword" placeholder="e.g. iPhone 15 Pro 256GB">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label for="exclude">Exclude Keywords</label>
        <input type="text" id="exclude" placeholder="Space-separated, e.g. \u30b8\u30e3\u30f3\u30af \u30b1\u30fc\u30b9 max">
      </div>
      <div class="form-group">
        <label for="limit">Max Results</label>
        <input type="number" id="limit" value="20" min="1" max="100">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label for="minPrice">Min Price (JPY)</label>
        <input type="number" id="minPrice" placeholder="e.g. 10000">
      </div>
      <div class="form-group">
        <label for="maxPrice">Max Price (JPY)</label>
        <input type="number" id="maxPrice" placeholder="e.g. 200000">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label for="statusFilter">Status</label>
        <select id="statusFilter">
          <option value="all">All Items</option>
          <option value="available">Available Only</option>
          <option value="sold">Sold Out Only</option>
        </select>
      </div>
      <div class="form-group">
        <label for="conditionFilter">Condition</label>
        <select id="conditionFilter">
          <option value="">All</option>
          <option value="新品、未使用">New / Unused</option>
          <option value="未使用に近い">Like New</option>
          <option value="目立った傷や汚れなし">Excellent (No Scratches)</option>
          <option value="やや傷や汚れあり">Good (Slight Scratches)</option>
          <option value="傷や汚れあり">Fair (Scratches/Dirt)</option>
          <option value="全体的に状態が悪い">Poor</option>
        </select>
      </div>
    </div>
    <button class="btn-search" id="searchBtn" onclick="doSearch()">
      <span class="spinner"></span>
      <span class="btn-text">Search Mercari</span>
    </button>
  </div>
  <div id="resultsContainer">
    <div class="status" id="initialStatus">
      <div class="emoji">&#x1F50D;</div>
      <p>Enter a keyword and click Search</p>
    </div>
  </div>
</div>
<div class="tab-content" id="tabBulk">
  <div class="card">
    <h2>Bulk Search</h2>
    <div class="bulk-area">
      <div class="file-upload" id="fileUpload">
        <input type="file" accept=".txt,.csv,.xlsx" id="fileInput" onchange="handleFile(event)">
        <div class="file-label">Drop a <strong>.txt</strong>, <strong>.csv</strong>, or <strong>.xlsx</strong> file here or click to browse</div>
      </div>
      <div style="text-align:center;color:#999;font-size:12px;">&mdash; OR &mdash;</div>
      <textarea id="bulkKeywords" placeholder="Paste keywords here, one per line&#10;e.g.&#10;iPhone 15 Pro&#10;Canon PowerShot&#10;Roland SC-88"></textarea>
      <div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap;">
        <div class="form-group" style="flex:1;min-width:120px;">
          <label for="bulkLimit">Per Keyword</label>
          <select id="bulkLimit"><option value="1">1</option><option value="2">2</option><option value="3" selected>3</option><option value="5">5</option></select>
        </div>
        <div class="form-group" style="flex:1;min-width:120px;">
          <label for="bulkStatus">Status</label>
          <select id="bulkStatus"><option value="all">All</option><option value="available">Available</option><option value="sold">Sold Out</option></select>
        </div>
        <div class="form-group" style="flex:1;min-width:120px;">
          <label for="bulkCondition">Condition</label>
          <select id="bulkCondition">
            <option value="">All</option>
            <option value="新品、未使用">New / Unused</option>
            <option value="未使用に近い">Like New</option>
            <option value="目立った傷や汚れなし">Excellent (No Scratches)</option>
            <option value="やや傷や汚れあり">Good (Slight Scratches)</option>
            <option value="傷や汚れあり">Fair (Scratches/Dirt)</option>
            <option value="全体的に状態が悪い">Poor</option>
          </select>
        </div>
        <div class="form-group" style="flex:1;min-width:80px;">
          <label for="bulkMinReviews">Min Reviews</label>
          <input type="number" id="bulkMinReviews" placeholder="0" min="0">
        </div>
        <div class="form-group" style="flex:1;min-width:80px;">
          <label for="bulkMaxReviews">Max Reviews</label>
          <input type="number" id="bulkMaxReviews" placeholder="Any" min="0">
        </div>
        <button class="btn-search" id="bulkBtn" onclick="doBulkSearch()" style="flex:1;min-width:150px;margin-top:18px;">
          <span class="spinner"></span>
          <span class="btn-text">Search Bulk</span>
        </button>
      </div>
      <div class="bulk-keyword-count" id="bulkCount"></div>
    </div>
  </div>
  <div id="bulkResults"></div>
</div>
<div class="modal-overlay" id="modalOverlay" onclick="closeModal(event)">
  <div class="modal" id="modalContent" onclick="event.stopPropagation()">
    <button class="modal-close" onclick="closeModal()">&times;</button>
    <div id="modalBody"><div class="modal-loading">Loading...</div></div>
  </div>
</div>
<script>
async function doSearch() {
  const btn = document.getElementById('searchBtn');
  const errDiv = document.getElementById('errorMsg');
  const resultsContainer = document.getElementById('resultsContainer');
  errDiv.style.display = 'none';
  const keyword = document.getElementById('keyword').value.trim();
  if (!keyword) { showError('Please enter a keyword'); return; }
  btn.classList.add('loading'); btn.disabled = true;
  try {
    const params = new URLSearchParams({
      keyword: keyword,
      exclude_keywords: document.getElementById('exclude').value.trim(),
      min_price: document.getElementById('minPrice').value || '',
      max_price: document.getElementById('maxPrice').value || '',
      limit: document.getElementById('limit').value || '20',
      status_filter: document.getElementById('statusFilter').value,
      condition: document.getElementById('conditionFilter').value
    });
    const resp = await fetch('/api/search?'+params.toString());
    if (!resp.ok) { const e = await resp.json(); throw new Error(e.error || 'Search failed'); }
    const data = await resp.json();
    renderResults(data, keyword);
  } catch (e) { showError(e.message); }
  finally { btn.classList.remove('loading'); btn.disabled = false; }
}
function showError(msg) {
  const e = document.getElementById('errorMsg');
  e.textContent = msg; e.style.display = 'block';
}
function renderResults(items, keyword) {
  const container = document.getElementById('resultsContainer');
  if (!items.length) {
    container.innerHTML = '<div class="status"><div class="emoji">&#x1F622;</div><p>No results found for "'+keyword+'"</p></div>';
    return;
  }
  let html = '<div class="count-badge">Found <strong>'+items.length+'</strong> item'+(items.length>1?'s':'')+'</div><div class="results">';
  for (const item of items) {
    html += '<div class="result-card">'
      + '<div class="result-icon">&#x1F4E6;</div>'
      + '<div class="result-info"><div class="result-name">'+(item.name_en ? escapeHtml(item.name_en) : escapeHtml(item.name))+'</div>'
      + '<div class="result-price">&yen;'+Number(item.price).toLocaleString()+' <span>JPY</span></div>'
        + (item.updated ? '<div class="result-updated">Updated: '+timeAgo(item.updated)+'</div>' : '')
        + (item.condition_en ? '<div class="result-condition">'+escapeHtml(item.condition_en)+'</div>' : '')
        + (item.seller_name ? '<div class="result-seller">Seller: '+escapeHtml(item.seller_name)+' ('+item.seller_reviews+' reviews)</div>' : '')
        + (item.status === 'ITEM_STATUS_SOLD_OUT' ? '<div class="result-tag sold">SOLD OUT</div>' : '')
      + (item.status === 'ITEM_STATUS_TRADING' ? '<div class="result-tag trading">TRADING</div>' : '')
      + (item.auction ? '<div class="result-tag auction">AUCTION</div>' : '')
      + '</div>'
      + '<div class="result-actions">'
      + '<button class="result-link desc-btn" onclick="fetchDescription(\\''+escapeHtml(item.id)+'\\', this)">Description</button>'
      + '<a class="result-link" href="'+escapeHtml(item.url)+'" target="_blank" rel="noopener">View &rarr;</a>'
      + '</div>'
      + '</div>';
  }
  html += '</div>';
  container.innerHTML = html;
}
function closeModal(e) {
  if (e && e.target !== e.currentTarget) return;
  document.getElementById('modalOverlay').classList.remove('active');
}
async function fetchDescription(itemId, btn) {
  const overlay = document.getElementById('modalOverlay');
  const body = document.getElementById('modalBody');
  body.innerHTML = '<div class="modal-loading">Loading & translating...</div>';
  overlay.classList.add('active');
  try {
    const resp = await fetch('/api/item-info/'+encodeURIComponent(itemId));
    if (!resp.ok) throw new Error('Failed to fetch');
    const data = await resp.json();
    const displayDesc = data.description_en || data.description || 'No description available';
    body.innerHTML = '<h3>'+escapeHtml(data.name)+'</h3>'
      + '<div class="modal-meta">'
      + (data.condition ? 'Condition: '+escapeHtml(data.condition)+' &middot; ' : '')
      + 'Seller: '+escapeHtml(data.seller)
      + '</div>'
      + '<div class="modal-desc">'+escapeHtml(displayDesc)+'</div>';
  } catch(e) {
    body.innerHTML = '<div class="modal-loading" style="color:#d42c25;">Error: '+e.message+'</div>';
  }
}
function timeAgo(ts) {
  const sec = Math.floor((Date.now() - ts*1000) / 1000);
  const intervals = [
    [31536000, 'year'], [2592000, 'month'], [604800, 'week'],
    [86400, 'day'], [3600, 'hour'], [60, 'minute']
  ];
  for (const [s, label] of intervals) {
    const n = Math.floor(sec / s);
    if (n >= 1) return n+' '+label+(n>1?'s':'')+' ago';
  }
  return 'just now';
}
function escapeHtml(str) { 
  const d = document.createElement('div'); d.textContent = str; return d.innerHTML; 
}
document.getElementById('keyword').addEventListener('keydown', function(e) {
  if (e.key === 'Enter') doSearch();
});
function switchTab(name) {
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
  document.querySelector('.tab-btn[onclick*="'+name+'"]').classList.add('active');
  document.getElementById('tab'+name.charAt(0).toUpperCase()+name.slice(1)).classList.add('active');
}
async function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'txt' || ext === 'csv') {
    const text = await file.text();
    document.getElementById('bulkKeywords').value = text;
  } else {
    const formData = new FormData();
    formData.append('file', file);
    try {
      const resp = await fetch('/api/upload-keywords', { method: 'POST', body: formData });
      if (!resp.ok) throw new Error('Upload failed');
      const data = await resp.json();
      if (data.keywords) document.getElementById('bulkKeywords').value = data.keywords.join('\\n');
    } catch(e) { showError(e.message); }
  }
  updateBulkCount();
}
function updateBulkCount() {
  const v = document.getElementById('bulkKeywords').value.trim();
  const n = v ? v.split('\\n').filter(l => l.trim()).length : 0;
  document.getElementById('bulkCount').textContent = n ? n+' keyword(s) loaded' : '';
}
document.getElementById('bulkKeywords').addEventListener('input', updateBulkCount);
async function doBulkSearch() {
  const btn = document.getElementById('bulkBtn');
  const errDiv = document.getElementById('errorMsg');
  errDiv.style.display = 'none';
  const text = document.getElementById('bulkKeywords').value.trim();
  if (!text) { showError('Enter keywords or upload a file'); return; }
  const keywords = text.split('\\n').map(l => l.trim()).filter(l => l);
  if (!keywords.length) { showError('No valid keywords found'); return; }
  if (keywords.length > 100) { showError('Maximum 100 keywords allowed'); return; }
  btn.classList.add('loading'); btn.disabled = true;
  try {
    const resp = await fetch('/api/bulk-search', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        keywords: keywords,
        per_keyword: parseInt(document.getElementById('bulkLimit').value) || 3,
        status_filter: document.getElementById('bulkStatus').value,
        condition: document.getElementById('bulkCondition').value,
        min_reviews: parseInt(document.getElementById('bulkMinReviews').value) || '',
        max_reviews: parseInt(document.getElementById('bulkMaxReviews').value) || ''
      })
    });
    if (!resp.ok) { const e = await resp.json(); throw new Error(e.error || 'Search failed'); }
    const data = await resp.json();
    renderBulkResults(data.results);
  } catch(e) { showError(e.message); }
  finally { btn.classList.remove('loading'); btn.disabled = false; }
}
function renderBulkResults(results) {
  const container = document.getElementById('bulkResults');
  const keys = Object.keys(results);
  if (!keys.length) {
    container.innerHTML = '<div class="status"><div class="emoji">&#x1F622;</div><p>No results found</p></div>';
    return;
  }
  let html = '<div class="count-badge">Results for <strong>'+keys.length+'</strong> keyword(s)</div>';
  html += '<button class="dl-btn" onclick="downloadCSV()">Download CSV</button>';
  let csvData = [];
  for (const kw of keys) {
    const items = results[kw] || [];
    html += '<div class="bulk-result-group"><div class="bulk-result-header" onclick="this.nextElementSibling.style.display=this.nextElementSibling.style.display===\\'none\\'?\\'block\\':\\'none\\'">'
      + escapeHtml(kw) + ' <span class="count">'+items.length+' item(s)</span></div>'
      + '<div class="bulk-result-body"><div class="results">';
    for (const item of items) {
      const name = item.name_en || item.name;
      html += '<div class="result-card">'
        + '<div class="result-icon">&#x1F4E6;</div>'
        + '<div class="result-info"><div class="result-name">'+escapeHtml(name)+'</div>'
        + '<div class="result-price">&yen;'+Number(item.price).toLocaleString()+' <span>JPY</span></div>'
      + (item.updated ? '<div class="result-updated">Updated: '+timeAgo(item.updated)+'</div>' : '')
      + (item.condition_en ? '<div class="result-condition">'+escapeHtml(item.condition_en)+'</div>' : '')
      + (item.seller_name ? '<div class="result-seller">Seller: '+escapeHtml(item.seller_name)+' ('+item.seller_reviews+' reviews)</div>' : '')
        + (item.status === 'ITEM_STATUS_SOLD_OUT' ? '<div class="result-tag sold">SOLD OUT</div>' : '')
        + (item.auction ? '<div class="result-tag auction">AUCTION</div>' : '')
        + '</div>'
        + '<div class="result-actions">'
        + '<button class="result-link desc-btn" onclick="fetchDescription(\\''+escapeHtml(item.id)+'\\', this)">Desc</button>'
        + '<a class="result-link" href="'+escapeHtml(item.url)+'" target="_blank" rel="noopener">View</a>'
        + '</div></div>';
      csvData.push({keyword: kw, name: name, price: item.price, url: item.url, status: item.status, condition_en: item.condition_en, seller_name: item.seller_name, seller_reviews: item.seller_reviews, updated: item.updated});
    }
    html += '</div></div></div>';
  }
  window._bulkCSV = csvData;
  html += '<button class="dl-btn" onclick="downloadCSV()" style="margin-top:12px;">Download CSV</button>';
  container.innerHTML = html;
}
function downloadCSV() {
  const data = window._bulkCSV || [];
  if (!data.length) return;
  let csv = '\\uFEFFKeyword,Name,Price JPY,URL,Status,Condition,Seller,Reviews,Updated\\n';
  for (const r of data) {
    csv += '"'+r.keyword+'","'+(r.name||'').replace(/"/g,'""')+'",'+r.price+',"'+r.url+'","'+r.status+'","'+(r.condition_en||'')+'","'+(r.seller_name||'')+'","'+(r.seller_reviews||'')+'","'+(r.updated?timeAgo(r.updated):'')+'"\\n';
  }
  const blob = new Blob([csv], {type: 'text/csv;charset=utf-8;'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'mercari_bulk_results.csv'; a.click();
}
</script>
</body>
</html>'''

@app.route('/')
def index():
    return HTML_PAGE

@app.route('/api/search')
def api_search():
    keyword = request.args.get('keyword', '').strip()
    if not keyword:
        return jsonify({'error': 'Keyword is required'}), 400
    
    exclude_keywords = request.args.get('exclude_keywords', '').strip()
    min_price = request.args.get('min_price', '').strip()
    max_price = request.args.get('max_price', '').strip()
    limit = request.args.get('limit', '20').strip()
    status_filter = request.args.get('status_filter', 'all').strip()
    condition_filter = request.args.get('condition', '').strip()
    
    try:
        limit = int(limit) if limit else 20
    except ValueError:
        limit = 20
    
    try:
        min_price_val = int(min_price) if min_price else None
    except ValueError:
        min_price_val = None
    
    try:
        max_price_val = int(max_price) if max_price else None
    except ValueError:
        max_price_val = None
    
    try:
        search_results = search(
            keyword,
            sort=MercariSort.SORT_SCORE,
            order=MercariOrder.ORDER_DESC,
            status=MercariSearchStatus.DEFAULT,
            exclude_keywords=exclude_keywords
        )
        
        COND_IDS = {'1': '新品、未使用', '2': '未使用に近い', '3': '目立った傷や汚れなし', '4': 'やや傷や汚れあり', '5': '傷や汚れあり', '6': '全体的に状態が悪い'}
        COND_EN = {
            '新品、未使用': 'NEW, UNUSED',
            '未使用に近い': 'LIKE NEW, UNUSED',
            '目立った傷や汚れなし': 'NO NOTICEABLE SCRATCHES OR DIRT',
            'やや傷や汚れあり': 'SLIGHT SCRATCHES OR DIRT',
            '傷や汚れあり': 'SCRATCHES OR DIRT',
            '全体的に状態が悪い': 'POOR CONDITION'
        }
        
        items = []
        required_terms = [t.lower() for t in keyword.split()]
        unwanted_terms = [t.lower() for t in exclude_keywords.split()] if exclude_keywords else []
        all_unwanted = list(set(unwanted_terms))
        
        for item in search_results:
            try:
                name = getattr(item, 'productName', None)
                price = getattr(item, 'price', None)
                url = getattr(item, 'productURL', '')
                if not name or price is None:
                    continue
                
                price = float(price)
                lower_name = name.lower()
                
                name_ok = all(t in lower_name for t in required_terms)
                name_bad = any(t in lower_name for t in all_unwanted) if all_unwanted else False
                
                if not name_ok or name_bad:
                    continue
                
                if min_price_val is not None and price < min_price_val:
                    continue
                if max_price_val is not None and price > max_price_val:
                    continue
                
                item_id = getattr(item, 'id', '')
                item_status = getattr(item, 'status', '')
                is_auction = getattr(item, 'auction', None) is not None
                cond_id = str(getattr(item, 'itemConditionId', None) or '')
                cond_name = COND_IDS.get(cond_id, '')
                
                if status_filter == 'available' and item_status != 'ITEM_STATUS_ON_SALE':
                    continue
                if status_filter == 'sold' and item_status != 'ITEM_STATUS_SOLD_OUT':
                    continue
                if condition_filter and cond_name != condition_filter:
                    continue
                
                updated_ts = getattr(item, 'updated', 0)

                # Detect shop item from search result (no extra API call)
                if getattr(item, 'isShopItem', False):
                    url = 'https://jp.mercari.com/shops/product/' + item_id

                full_info = None
                seller_name = ''
                seller_reviews = 0
                try:
                    full_info = getItemInfo(item_id)
                    if full_info is not None:
                        if hasattr(full_info, 'seller') and full_info.seller is not None:
                            seller_name = full_info.seller.name or ''
                            seller_reviews = full_info.seller.num_ratings or 0
                        if hasattr(full_info, 'item_condition') and full_info.item_condition is not None:
                            cond_name = full_info.item_condition.name or cond_name
                        if hasattr(full_info, 'is_shop_item') and str(full_info.is_shop_item).lower() in ('true', '1'):
                            url = 'https://jp.mercari.com/shops/product/' + item_id
                except Exception:
                    pass

                items.append({
                    'id': item_id,
                    'name': name,
                    'price': price,
                    'url': url,
                    'status': item_status,
                    'auction': is_auction,
                    'condition': cond_name,
                    'condition_en': COND_EN.get(cond_name, cond_name),
                    'updated': updated_ts,
                    'seller_name': seller_name,
                    'seller_reviews': seller_reviews
                })
                
                if len(items) >= limit:
                    break
            except Exception:
                continue
        
        names = [it['name'] for it in items]
        translations = batch_translate(names)
        for i, it in enumerate(items):
            it['name_en'] = translations[i] if i < len(translations) else ''
        
        return jsonify(items)
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def translate_ja_to_en(text):
    if not text:
        return ''
    try:
        resp = http_requests.post(
            'https://translate.googleapis.com/translate_a/single',
            params={'client': 'gtx', 'sl': 'ja', 'tl': 'en', 'dt': 't'},
            data={'q': text},
            timeout=10
        )
        resp.raise_for_status()
        data = resp.json()
        if data and data[0]:
            return ''.join(part[0] for part in data[0] if part[0])
    except Exception:
        pass
    return ''

def batch_translate(texts):
    if not texts:
        return []
    joined = '\n'.join(texts)
    try:
        resp = http_requests.post(
            'https://translate.googleapis.com/translate_a/single',
            params={'client': 'gtx', 'sl': 'ja', 'tl': 'en', 'dt': 't'},
            data={'q': joined},
            timeout=15
        )
        resp.raise_for_status()
        data = resp.json()
        if data and data[0]:
            result = []
            for part in data[0]:
                if part and part[0]:
                    result.append(part[0])
                else:
                    result.append('')
            while len(result) < len(texts):
                result.append('')
            return result[:len(texts)]
    except Exception:
        pass
    return [''] * len(texts)

@app.route('/api/item-info/<item_id>')
def api_item_info(item_id):
    try:
        item = getItemInfo(item_id)
        desc = item.description or ''
        en_desc = translate_ja_to_en(desc)
        return jsonify({
            'id': item.id,
            'name': item.name,
            'description': desc,
            'description_en': en_desc,
            'price': item.price,
            'condition': item.item_condition.name if hasattr(item, 'item_condition') else '',
            'seller': item.seller.name if hasattr(item, 'seller') else ''
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/upload-keywords', methods=['POST'])
def api_upload_keywords():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        f = request.files['file']
        if f.filename == '':
            return jsonify({'error': 'Empty filename'}), 400
        ext = os.path.splitext(f.filename)[1].lower()
        keywords = []
        if ext in ('.txt', '.csv'):
            content = f.read().decode('utf-8', errors='ignore')
            for line in content.splitlines():
                line = line.strip()
                if line:
                    keywords.append(line)
        elif ext in ('.xlsx', '.xls'):
            wb = openpyxl.load_workbook(io.BytesIO(f.read()))
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                if row and row[0] is not None:
                    kw = str(row[0]).strip()
                    if kw:
                        keywords.append(kw)
        else:
            return jsonify({'error': 'Unsupported file format. Use .txt, .csv, or .xlsx'}), 400
        return jsonify({'keywords': keywords, 'count': len(keywords)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/bulk-search', methods=['POST'])
def api_bulk_search():
    try:
        data = request.get_json()
        if not data or 'keywords' not in data:
            return jsonify({'error': 'Keywords required'}), 400
        keywords = data['keywords']
        if not isinstance(keywords, list) or not keywords:
            return jsonify({'error': 'Keywords must be a non-empty array'}), 400
        per_keyword = int(data.get('per_keyword', 3))
        per_keyword = max(1, min(per_keyword, 10))
        status_filter = data.get('status_filter', 'all')
        condition_filter = data.get('condition', '')
        min_reviews = int(data.get('min_reviews', 0)) if data.get('min_reviews') else None
        max_reviews = int(data.get('max_reviews', 0)) if data.get('max_reviews') else None
        COND_IDS = {'1': '新品、未使用', '2': '未使用に近い', '3': '目立った傷や汚れなし', '4': 'やや傷や汚れあり', '5': '傷や汚れあり', '6': '全体的に状態が悪い'}
        COND_EN = {
            '新品、未使用': 'NEW, UNUSED',
            '未使用に近い': 'LIKE NEW, UNUSED',
            '目立った傷や汚れなし': 'NO NOTICEABLE SCRATCHES OR DIRT',
            'やや傷や汚れあり': 'SLIGHT SCRATCHES OR DIRT',
            '傷や汚れあり': 'SCRATCHES OR DIRT',
            '全体的に状態が悪い': 'POOR CONDITION'
        }
        
        results = {}
        for kw in keywords[:100]:
            kw = kw.strip()
            if not kw:
                continue
            try:
                search_results = search(kw, sort=MercariSort.SORT_SCORE, order=MercariOrder.ORDER_DESC, status=MercariSearchStatus.ON_SALE)
                items = []
                for item in search_results:
                    name = getattr(item, 'productName', None)
                    price = getattr(item, 'price', None)
                    if not name or price is None:
                        continue
                    item_status = getattr(item, 'status', '')
                    if status_filter == 'available' and item_status != 'ITEM_STATUS_ON_SALE':
                        continue
                    if status_filter == 'sold' and item_status != 'ITEM_STATUS_SOLD_OUT':
                        continue
                    cond_id = str(getattr(item, 'itemConditionId', None) or '')
                    cond_name = COND_IDS.get(cond_id, '')
                    if condition_filter and cond_name != condition_filter:
                        continue
                    updated_ts = getattr(item, 'updated', 0)
                    item_id = getattr(item, 'id', '')
                    full_info = None
                    seller_name = ''
                    seller_reviews = 0
                    item_url = getattr(item, 'productURL', '')
                    # Detect shop item from search result (no extra API call)
                    if getattr(item, 'isShopItem', False):
                        item_url = 'https://jp.mercari.com/shops/product/' + item_id
                    try:
                        full_info = getItemInfo(item_id)
                        if full_info is not None:
                            if hasattr(full_info, 'seller') and full_info.seller is not None:
                                seller_name = full_info.seller.name or ''
                                seller_reviews = full_info.seller.num_ratings or 0
                            if hasattr(full_info, 'item_condition') and full_info.item_condition is not None:
                                cond_name = full_info.item_condition.name or cond_name
                            if hasattr(full_info, 'is_shop_item') and str(full_info.is_shop_item).lower() in ('true', '1'):
                                item_url = 'https://jp.mercari.com/shops/product/' + item_id
                    except Exception:
                        pass
                    if min_reviews is not None and seller_reviews < min_reviews:
                        continue
                    if max_reviews is not None and seller_reviews > max_reviews:
                        continue
                    items.append({
                        'id': item_id,
                        'name': name,
                        'price': float(price),
                        'url': item_url,
                        'status': item_status,
                        'auction': getattr(item, 'auction', None) is not None,
                        'condition': cond_name,
                        'condition_en': COND_EN.get(cond_name, cond_name),
                        'updated': updated_ts,
                        'seller_name': seller_name,
                        'seller_reviews': seller_reviews
                    })
                    if len(items) >= per_keyword:
                        break
                if items:
                    names = [it['name'] for it in items]
                    translations = batch_translate(names)
                    for i, it in enumerate(items):
                        it['name_en'] = translations[i] if i < len(translations) else ''
                results[kw] = items
            except Exception:
                results[kw] = []
        return jsonify({'results': results})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    print('='*60)
    print(' Mercari JP Search GUI')
    print('='*60)
    print(' Open http://127.0.0.1:5000 in your browser')
    print(' Press CTRL+C to stop the server')
    print('='*60)
    app.run(debug=True, host='127.0.0.1', port=5000)
