from flask import Flask, request, jsonify, Response, stream_with_context
import json
import sys
import os
import io
import requests as http_requests
import openpyxl
from concurrent.futures import ThreadPoolExecutor, as_completed
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'temp_download', 'mercari-2.2.1'))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'temp_download'))
from mercari import search, getItemInfo, getShopProductInfo, ShopProductInfo, MercariSort, MercariOrder, MercariSearchStatus
from yahoo_auctions import search_yahoo, get_item_detail, parse_search_results, build_search_url
from ebay_search import search_ebay, EbayItem

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
<title>Mercari JP + Yahoo Auctions Search</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; }
  body { background: #f0f2f5; display: flex; justify-content: center; min-height: 100vh; }
  .container { max-width: 1000px; width: 100%; padding: 24px 16px; }
  h1 { font-size: 24px; color: #ea352d; margin-bottom: 20px; display: flex; align-items: center; gap: 8px; }
  h1::before { content: "📱"; font-size: 24px; }
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
  .result-thumb { width: 80px; height: 80px; border-radius: 8px; overflow: hidden; flex-shrink: 0; background: #f5f5f5; display: flex; align-items: center; justify-content: center; }
  .result-thumb img { width: 100%; height: 100%; object-fit: cover; display: block; }
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
  .count-badge.streaming { background: #fff3cd; color: #856404; font-size: 13px; }
  .bulk-result-group { margin-bottom: 16px; border: 1px solid #eee; border-radius: 10px; overflow: hidden; }
  .bulk-result-header { padding: 10px 14px; background: #f9f9f9; font-size: 14px; font-weight: 600; color: #333; cursor: pointer; display: flex; justify-content: space-between; align-items: center; }
  .bulk-result-header:hover { background: #f0f0f0; }
  .bulk-result-header .count { font-size: 12px; font-weight: 400; color: #999; }
  .keyword-badge { font-size: 10px; font-weight: 600; color: #666; background: #f0f0f0; padding: 2px 8px; border-radius: 4px; display: inline-block; margin-bottom: 6px; text-transform: uppercase; letter-spacing: 0.3px; }
  .bulk-result-body { padding: 8px 14px 14px; }
  .bulk-result-body .result-card { margin-bottom: 8px; }
  .bulk-result-body .result-card:last-child { margin-bottom: 0; }
  .bulk-result-body .result-card .result-link { font-size: 12px; padding: 6px 12px; }
  .ebay-search-btn { border: 1px solid #0064d2 !important; color: #0064d2 !important; background: #fff !important; cursor: pointer; }
  .ebay-search-btn:hover { background: #0064d2 !important; color: #fff !important; }
  .bulk-result-body .result-card .result-name { font-size: 13px; }
  .dl-btn { padding: 8px 16px; background: #333; color: white; border: none; border-radius: 6px; font-size: 13px; cursor: pointer; margin-top: 8px; }
  .dl-btn:hover { background: #555; }
  .yahoo-result-card { display: flex; gap: 12px; padding: 12px; background: white; border: 1px solid #eee; border-radius: 10px; transition: box-shadow 0.2s; align-items: stretch; overflow: hidden; }
  .yahoo-result-card:hover { box-shadow: 0 2px 12px rgba(0,0,0,0.1); }
  .yahoo-thumb { width: 90px; min-height: 90px; border-radius: 6px; overflow: hidden; flex-shrink: 0; background: #f5f5f5; }
  .yahoo-thumb img { width: 100%; height: 100%; object-fit: cover; display: block; }
  .yahoo-info { flex: 1; min-width: 0; display: flex; flex-direction: column; gap: 3px; }
  .yahoo-title { font-size: 14px; font-weight: 500; color: #333; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .yahoo-price-row { display: flex; align-items: baseline; gap: 8px; flex-wrap: wrap; }
  .yahoo-price { font-size: 18px; font-weight: 700; color: #222; }
  .yahoo-price span { font-size: 11px; font-weight: 400; color: #999; }
  .yahoo-buynow { font-size: 13px; color: #888; }
  .yahoo-meta { display: flex; gap: 12px; font-size: 11px; color: #777; flex-wrap: wrap; }
  .yahoo-meta span { display: inline-flex; align-items: center; gap: 3px; }
  .yahoo-badge { display: inline-block; font-size: 10px; font-weight: 700; padding: 1px 5px; border-radius: 3px; margin-right: 4px; }
  .yahoo-badge.freeship { background: #e8f5e9; color: #2e7d32; }
  .yahoo-badge.unused { background: #e3f2fd; color: #1565c0; }
  .yahoo-badge.newitem { background: #fce4ec; color: #c62828; }
  .yahoo-badge.sold { background: #ea352d; color: white; }
  .yahoo-actions { display: flex; flex-direction: column; gap: 4px; flex-shrink: 0; justify-content: center; }
  .yahoo-actions .result-link { text-align: center; }
  .yahoo-thumb-placeholder { width: 100%; height: 100%; display: flex; align-items: center; justify-content: center; font-size: 28px; color: #ccc; }
  .combined-grid { display: flex; gap: 16px; }
  .combined-grid .combined-col { flex: 1; min-width: 0; width: 50%; }
  .combined-grid .combined-col h3 { font-size: 14px; color: #333; margin-bottom: 12px; padding-bottom: 6px; border-bottom: 2px solid #eee; }
  .combined-grid .combined-col h3.mercari-header { border-bottom-color: #ea352d; }
  .combined-grid .combined-col h3.yahoo-header { border-bottom-color: #222; }
  .combined-grid .result-name, .combined-grid .yahoo-title { white-space: normal; overflow: visible; text-overflow: clip; word-break: break-word; line-height: 1.4; }
  .combined-grid .result-actions, .combined-grid .yahoo-actions { flex-wrap: wrap; justify-content: flex-end; max-width: 130px; }
  .combined-grid .result-actions .result-link, .combined-grid .yahoo-actions .result-link { margin-bottom: 4px; }
  .combined-keyword-title { font-size: 13px; font-weight: 600; color: #555; margin: 12px 0 8px; padding: 6px 10px; background: #f5f5f5; border-radius: 6px; }
  .stock-card { display: flex; gap: 12px; padding: 12px 16px; background: white; border: 1px solid #eee; border-radius: 10px; align-items: center; margin-bottom: 8px; transition: box-shadow 0.2s; }
  .stock-card:hover { box-shadow: 0 2px 12px rgba(0,0,0,0.1); }
  .stock-status { padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; white-space: nowrap; flex-shrink: 0; }
  .stock-status.available { background: #e8f5e9; color: #2e7d32; }
  .stock-status.sold { background: #fce4ec; color: #c62828; }
  .stock-status.error { background: #fff3e0; color: #e65100; }
  .stock-link-text { font-size: 12px; color: #888; max-width: 280px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex-shrink: 0; }
  .stock-platform { font-size: 11px; font-weight: 600; padding: 2px 8px; border-radius: 4px; flex-shrink: 0; }
  .stock-platform.mercari { background: #fce4ec; color: #c62828; }
  .stock-platform.yahoo { background: #e3f2fd; color: #1565c0; }
  .stock-platform.yahoo_shopping { background: #fff3e0; color: #e65100; }
  .stock-price { font-size: 15px; font-weight: 700; color: #333; white-space: nowrap; flex-shrink: 0; }
  .stock-name { flex: 1; min-width: 0; font-size: 13px; font-weight: 500; color: #333; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .stock-keyword { font-size: 11px; font-weight: 600; color: #555; background: #f0f0f0; padding: 2px 8px; border-radius: 4px; white-space: nowrap; flex-shrink: 0; max-width: 140px; overflow: hidden; text-overflow: ellipsis; }
  .stock-sno { font-size: 12px; font-weight: 700; color: #999; min-width: 28px; text-align: center; flex-shrink: 0; }
  .ebay-result-card { display: flex; gap: 12px; padding: 12px; background: white; border: 1px solid #eee; border-radius: 10px; transition: box-shadow 0.2s; align-items: stretch; overflow: hidden; }
  .ebay-result-card:hover { box-shadow: 0 2px 12px rgba(0,0,0,0.1); }
  .ebay-thumb { width: 90px; min-height: 90px; border-radius: 6px; overflow: hidden; flex-shrink: 0; background: #f5f5f5; }
  .ebay-thumb img { width: 100%; height: 100%; object-fit: cover; display: block; }
  .ebay-info { flex: 1; min-width: 0; display: flex; flex-direction: column; gap: 3px; }
  .ebay-title { font-size: 14px; font-weight: 500; color: #333; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .ebay-price-row { display: flex; align-items: baseline; gap: 8px; flex-wrap: wrap; }
  .ebay-price { font-size: 18px; font-weight: 700; color: #0064d2; }
  .ebay-price span { font-size: 11px; font-weight: 400; color: #999; }
  .ebay-bin-price { font-size: 13px; color: #888; }
  .ebay-meta { display: flex; gap: 12px; font-size: 11px; color: #777; flex-wrap: wrap; }
  .ebay-meta span { display: inline-flex; align-items: center; gap: 3px; }
  .ebay-badge { display: inline-block; font-size: 10px; font-weight: 700; padding: 1px 5px; border-radius: 3px; margin-right: 4px; }
  .ebay-badge.freeship { background: #e8f5e9; color: #2e7d32; }
  .ebay-badge.toprated { background: #fff3e0; color: #e65100; }
  .ebay-badge.auction { background: #e3f2fd; color: #1565c0; }
  .ebay-badge.bin { background: #f3e5f5; color: #7b1fa2; }
  .ebay-actions { display: flex; flex-direction: column; gap: 4px; flex-shrink: 0; justify-content: center; }
  .ebay-actions .result-link { text-align: center; }
  .ebay-thumb-placeholder { width: 100%; height: 100%; display: flex; align-items: center; justify-content: center; font-size: 28px; color: #ccc; }
  .settings-section { background: #f9f9f9; border: 1px solid #eee; border-radius: 8px; padding: 12px 16px; margin-bottom: 16px; display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
  .settings-section label { font-size: 12px; color: #666; margin-bottom: 2px; font-weight: 500; }
  .settings-section input { padding: 8px 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 13px; outline: none; flex: 1; min-width: 200px; background: white; }
  .settings-section input:focus { border-color: #0064d2; }
  .settings-section .btn-save { padding: 8px 16px; background: #0064d2; color: white; border: none; border-radius: 6px; font-size: 13px; cursor: pointer; font-weight: 500; }
  .settings-section .btn-save:hover { background: #0053b0; }
  .settings-section .status-msg { font-size: 12px; color: #2e7d32; display: none; }
  @media (max-width: 768px) { .combined-grid { flex-direction: column; } .combined-grid .combined-col { width: 100%; } }
  @media (max-width: 600px) { .form-row { grid-template-columns: 1fr; } }
</style>
</head>
<body>
<div class="container">
  <h1>Mercari JP Search</h1>
  <div id="errorMsg" class="error-msg"></div>
  <div class="settings-section" id="ebaySettings">
    <label for="ebayAppId" style="white-space:nowrap;">eBay App ID:</label>
    <input type="text" id="ebayAppId" placeholder="Paste your eBay App ID (Client ID)" value="" style="min-width:250px;">
    <label for="ebayCertId" style="white-space:nowrap;">Cert ID:</label>
    <input type="text" id="ebayCertId" placeholder="Paste your eBay Cert ID (Client Secret)" value="" style="min-width:250px;">
    <button class="btn-save" onclick="saveEbayAppId()">Save</button>
    <span class="status-msg" id="ebaySettingsStatus">Saved</span>
  </div>
  <div class="tabs">
    <button class="tab-btn active" data-tab="simple" onclick="switchTab('simple')">Simple Search</button>
    <button class="tab-btn" data-tab="bulk" onclick="switchTab('bulk')">Mercari Bulk Search</button>
    <button class="tab-btn" data-tab="yahoo" onclick="switchTab('yahoo')">Yahoo Auctions</button>
    <button class="tab-btn" data-tab="yahooBulk" onclick="switchTab('yahooBulk')">Yahoo Bulk</button>
    <button class="tab-btn" data-tab="combinedBulk" onclick="switchTab('combinedBulk')">Mercari + Yahoo Bulk</button>
    <button class="tab-btn" data-tab="stockCheck" onclick="switchTab('stockCheck')">Check Out Of Stock</button>
    <button class="tab-btn" data-tab="ebay" onclick="switchTab('ebay')">eBay Search</button>
    <button class="tab-btn" data-tab="ebayBulk" onclick="switchTab('ebayBulk')">eBay Bulk</button>
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
    <h2>Mercari Bulk Search</h2>
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
<div class="tab-content" id="tabYahoo">
  <div class="card">
    <div class="form-row">
      <div class="form-group full">
        <label for="yahooKeyword">Keyword *</label>
        <input type="text" id="yahooKeyword" placeholder="e.g. iPhone 15 Pro 256GB">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label for="yahooMinPrice">Min Price (JPY)</label>
        <input type="number" id="yahooMinPrice" placeholder="e.g. 10000">
      </div>
      <div class="form-group">
        <label for="yahooMaxPrice">Max Price (JPY)</label>
        <input type="number" id="yahooMaxPrice" placeholder="e.g. 200000">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label for="yahooStatus">Status</label>
        <select id="yahooStatus">
          <option value="live">Live Auctions</option>
          <option value="sold">Sold / Ended</option>
        </select>
      </div>
      <div class="form-group">
        <label for="yahooSort">Sort By</label>
        <select id="yahooSort">
          <option value="new">Newly Listed</option>
          <option value="end">Ending Soon</option>
          <option value="price">Price: Low to High</option>
          <option value="-price">Price: High to Low</option>
          <option value="bid">Most Bids</option>
        </select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label for="yahooCondition">Condition</label>
        <select id="yahooCondition">
          <option value="">All</option>
          <option value="1">Unused</option>
          <option value="2">Used</option>
          <option value="3">Near Unused</option>
          <option value="4">No Scratches/Dirt</option>
        </select>
      </div>
      <div class="form-group">
        <label for="yahooBIN">BIN Filter</label>
        <select id="yahooBIN">
          <option value="all">All Items</option>
          <option value="bin_only">BIN Only</option>
        </select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label for="yahooLimit">Max Results</label>
        <input type="number" id="yahooLimit" value="20" min="1" max="100">
      </div>
      <div class="form-group"></div>
    </div>
    <button class="btn-search" id="yahooSearchBtn" onclick="doYahooSearch()">
      <span class="spinner"></span>
      <span class="btn-text">Search Yahoo Auctions</span>
    </button>
  </div>
  <div id="yahooResultsContainer">
    <div class="status" id="yahooInitialStatus">
      <div class="emoji">&#x1F50D;</div>
      <p>Enter a keyword and click Search</p>
    </div>
  </div>
  <button class="dl-btn" id="yahooDlBtn" onclick="downloadYahooCSV()" style="display:none;margin-top:8px;">Download CSV</button>
</div>
<div class="tab-content" id="tabYahooBulk">
  <div class="card">
    <h2>Yahoo Auctions Bulk Search</h2>
    <div class="bulk-area">
      <div class="file-upload" id="yahooBulkFileUpload">
        <input type="file" accept=".txt,.csv,.xlsx" id="yahooBulkFileInput" onchange="handleYahooBulkFile(event)">
        <div class="file-label">Drop a <strong>.txt</strong>, <strong>.csv</strong>, or <strong>.xlsx</strong> file here or click to browse</div>
      </div>
      <div style="text-align:center;color:#999;font-size:12px;">&mdash; OR &mdash;</div>
      <textarea id="yahooBulkKeywords" placeholder="Paste keywords here, one per line&#10;e.g.&#10;iPhone 15 Pro&#10;Canon PowerShot&#10;Roland SC-88"></textarea>
      <div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap;">
        <div class="form-group" style="flex:1;min-width:100px;">
          <label for="yahooBulkLimit">Per Keyword</label>
          <select id="yahooBulkLimit"><option value="1">1</option><option value="2">2</option><option value="3" selected>3</option><option value="5">5</option><option value="10">10</option></select>
        </div>
        <div class="form-group" style="flex:1;min-width:100px;">
          <label for="yahooBulkStatus">Status</label>
          <select id="yahooBulkStatus"><option value="live">Live Auctions</option><option value="sold">Sold / Ended</option></select>
        </div>
        <div class="form-group" style="flex:1;min-width:100px;">
          <label for="yahooBulkCondition">Condition</label>
          <select id="yahooBulkCondition">
            <option value="">All</option>
            <option value="1">Unused</option>
            <option value="2">Used</option>
            <option value="3">Near Unused</option>
            <option value="4">No Scratches/Dirt</option>
          </select>
        </div>
        <div class="form-group" style="flex:1;min-width:100px;">
          <label for="yahooBulkSort">Sort By</label>
          <select id="yahooBulkSort">
            <option value="new">Newly Listed</option>
            <option value="end">Ending Soon</option>
            <option value="price">Price: Low to High</option>
            <option value="-price">Price: High to Low</option>
            <option value="bid">Most Bids</option>
          </select>
        </div>
        <div class="form-group" style="flex:1;min-width:100px;">
          <label for="yahooBulkBIN">BIN Filter</label>
          <select id="yahooBulkBIN">
            <option value="all">All Items</option>
            <option value="bin_only">BIN Only</option>
          </select>
        </div>
      </div>
      <div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap;">
        <div class="form-group" style="flex:1;min-width:100px;">
          <label for="yahooBulkMinPrice">Min Price (JPY)</label>
          <input type="number" id="yahooBulkMinPrice" placeholder="e.g. 10000">
        </div>
        <div class="form-group" style="flex:1;min-width:100px;">
          <label for="yahooBulkMaxPrice">Max Price (JPY)</label>
          <input type="number" id="yahooBulkMaxPrice" placeholder="e.g. 200000">
        </div>
        <button class="btn-search" id="yahooBulkBtn" onclick="doYahooBulkSearch()" style="flex:1;min-width:150px;margin-top:18px;">
          <span class="spinner"></span>
          <span class="btn-text">Search Yahoo Bulk</span>
        </button>
      </div>
      <div class="bulk-keyword-count" id="yahooBulkCount"></div>
    </div>
  </div>
  <div id="yahooBulkResults"></div>
</div>
<div class="tab-content" id="tabCombinedBulk">
  <div class="card">
    <h2>Mercari + Yahoo Bulk Search</h2>
    <div class="bulk-area">
      <div class="file-upload" id="combinedFileUpload">
        <input type="file" accept=".txt,.csv,.xlsx" id="combinedFileInput" onchange="handleCombinedFile(event)">
        <div class="file-label">Drop a <strong>.txt</strong>, <strong>.csv</strong>, or <strong>.xlsx</strong> file here or click to browse</div>
      </div>
      <div style="text-align:center;color:#999;font-size:12px;">&mdash; OR &mdash;</div>
      <textarea id="combinedKeywords" placeholder="Paste keywords here, one per line&#10;e.g.&#10;iPhone 15 Pro&#10;Canon PowerShot&#10;Roland SC-88"></textarea>
      <div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap;">
        <div class="form-group" style="flex:1;min-width:100px;">
          <label for="combinedLimit">Per Keyword</label>
          <select id="combinedLimit"><option value="1">1</option><option value="2">2</option><option value="3" selected>3</option><option value="5">5</option></select>
        </div>
        <div class="form-group" style="flex:1;min-width:120px;">
          <label for="combinedStatus">Mercari Status</label>
          <select id="combinedStatus"><option value="all">All</option><option value="available">Available</option><option value="sold">Sold Out</option></select>
        </div>
        <div class="form-group" style="flex:1;min-width:120px;">
          <label for="combinedCondition">Mercari Condition</label>
          <select id="combinedCondition">
            <option value="">All</option>
            <option value="新品、未使用">New / Unused</option>
            <option value="未使用に近い">Like New</option>
            <option value="目立った傷や汚れなし">Excellent (No Scratches)</option>
            <option value="やや傷や汚れあり">Good (Slight Scratches)</option>
            <option value="傷や汚れあり">Fair (Scratches/Dirt)</option>
            <option value="全体的に状態が悪い">Poor</option>
          </select>
        </div>
        <div class="form-group" style="flex:1;min-width:100px;">
          <label for="combinedMinReviews">Min Reviews</label>
          <input type="number" id="combinedMinReviews" placeholder="e.g. 50" min="0">
        </div>
        <div class="form-group" style="flex:0;min-width:auto;display:flex;align-items:center;gap:6px;margin-top:18px;flex-direction:row;">
          <input type="checkbox" id="combinedExcludeAuctions" style="width:16px;height:16px;cursor:pointer;">
          <label for="combinedExcludeAuctions" style="margin:0;cursor:pointer;font-size:12px;white-space:nowrap;">Exclude Mercari Auctions</label>
        </div>
        <div style="width:1px;height:36px;background:#ddd;align-self:center;margin-top:18px;"></div>
        <div class="form-group" style="flex:1;min-width:100px;">
          <label for="combinedYahooStatus">Yahoo Status</label>
          <select id="combinedYahooStatus"><option value="live">Live Auctions</option><option value="sold">Sold / Ended</option></select>
        </div>
        <div class="form-group" style="flex:1;min-width:100px;">
          <label for="combinedYahooCondition">Yahoo Condition</label>
          <select id="combinedYahooCondition">
            <option value="">All</option>
            <option value="1">Unused</option>
            <option value="2">Used</option>
            <option value="3">Near Unused</option>
            <option value="4">No Scratches/Dirt</option>
          </select>
        </div>
        <div class="form-group" style="flex:1;min-width:100px;">
          <label for="combinedYahooBIN">BIN Filter</label>
          <select id="combinedYahooBIN">
            <option value="all">All Items</option>
            <option value="bin_only">BIN Only</option>
          </select>
        </div>
        <button class="btn-search" id="combinedBtn" onclick="doCombinedBulkSearch()" style="flex:1;min-width:150px;margin-top:18px;">
          <span class="spinner"></span>
          <span class="btn-text">Search Both</span>
        </button>
      </div>
      <div class="bulk-keyword-count" id="combinedCount"></div>
    </div>
  </div>
  <div id="combinedResults">
    <div class="status"><div class="emoji">&#x1F50D;</div><p>Enter keywords and click Search Both</p></div>
  </div>
</div>
<div class="tab-content" id="tabEbay">
  <div class="card">
    <div class="form-row">
      <div class="form-group full">
        <label for="ebayKeyword">Keyword *</label>
        <input type="text" id="ebayKeyword" placeholder="e.g. iPhone 15 Pro 256GB">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label for="ebayMinPrice">Min Price (USD)</label>
        <input type="number" id="ebayMinPrice" placeholder="e.g. 100" step="0.01">
      </div>
      <div class="form-group">
        <label for="ebayMaxPrice">Max Price (USD)</label>
        <input type="number" id="ebayMaxPrice" placeholder="e.g. 2000" step="0.01">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label for="ebayCondition">Condition</label>
        <select id="ebayCondition">
          <option value="">All</option>
          <option value="new">New</option>
          <option value="open_box">Open Box</option>
          <option value="refurbished">Refurbished</option>
          <option value="used">Used</option>
          <option value="for_parts">For Parts / Not Working</option>
        </select>
      </div>
      <div class="form-group">
        <label for="ebaySort">Sort By</label>
        <select id="ebaySort">
          <option value="best_match">Best Match</option>
          <option value="price_asc" selected>Price: Low to High</option>
          <option value="price_desc">Price: High to Low</option>
        </select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label for="ebaySite">eBay Site</label>
        <select id="ebaySite">
          <option value="EBAY-US" selected>United States</option>
          <option value="EBAY-GB">United Kingdom</option>
          <option value="EBAY-DE">Germany</option>
          <option value="EBAY-FR">France</option>
          <option value="EBAY-JAPAN">Japan</option>
          <option value="EBAY-AU">Australia</option>
          <option value="EBAY-CA">Canada</option>
          <option value="EBAY-IT">Italy</option>
          <option value="EBAY-ES">Spain</option>
          <option value="EBAY-HK">Hong Kong</option>
          <option value="EBAY-SG">Singapore</option>
        </select>
      </div>
      <div class="form-group">
        <label for="ebayItemLocation">Seller Location</label>
        <select id="ebayItemLocation">
          <option value="">Anywhere</option>
          <option value="US">United States</option>
          <option value="JP" selected>Japan</option>
          <option value="PK">Pakistan</option>
          <option value="GB">United Kingdom</option>
          <option value="DE">Germany</option>
          <option value="FR">France</option>
          <option value="CA">Canada</option>
          <option value="AU">Australia</option>
          <option value="CN">China</option>
          <option value="HK">Hong Kong</option>
          <option value="SG">Singapore</option>
          <option value="KR">South Korea</option>
          <option value="TW">Taiwan</option>
          <option value="IN">India</option>
          <option value="AE">UAE</option>
        </select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group" style="display:flex;align-items:flex-end;">
        <label style="opacity:0;">BIN Only</label>
        <div style="display:flex;align-items:center;gap:8px;padding:8px 0;">
          <input type="checkbox" id="ebayBINOnly" style="width:16px;height:16px;cursor:pointer;">
          <label for="ebayBINOnly" style="margin:0;cursor:pointer;font-size:13px;">Buy It Now only</label>
        </div>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label for="ebayLimit">Max Results</label>
        <input type="number" id="ebayLimit" value="20" min="1" max="100">
      </div>
      <div class="form-group"></div>
    </div>
    <button class="btn-search" id="ebaySearchBtn" onclick="doEbaySearch()" style="background:#0064d2;">
      <span class="spinner"></span>
      <span class="btn-text">Search eBay</span>
    </button>
  </div>
  <div id="ebayResultsContainer">
    <div class="status" id="ebayInitialStatus">
      <div class="emoji">&#x1F50D;</div>
      <p>Enter a keyword and click Search. Set your eBay App ID in the settings bar above.</p>
    </div>
  </div>
  <button class="dl-btn" id="ebayDlBtn" onclick="downloadEbayCSV()" style="display:none;margin-top:8px;background:#0064d2;">Download CSV</button>
</div>
<div class="tab-content" id="tabEbayBulk">
  <div class="card">
    <h2>eBay Bulk Search</h2>
    <div class="bulk-area">
      <div class="file-upload" id="ebayBulkFileUpload">
        <input type="file" accept=".txt,.csv,.xlsx" id="ebayBulkFileInput" onchange="handleEbayBulkFile(event)">
        <div class="file-label">Drop a <strong>.txt</strong>, <strong>.csv</strong>, or <strong>.xlsx</strong> file here or click to browse</div>
      </div>
      <div style="text-align:center;color:#999;font-size:12px;">&mdash; OR &mdash;</div>
      <textarea id="ebayBulkKeywords" placeholder="Paste keywords here, one per line&#10;e.g.&#10;iPhone 15 Pro&#10;Canon PowerShot&#10;Roland SC-88"></textarea>
      <div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap;">
        <div class="form-group" style="flex:1;min-width:100px;">
          <label for="ebayBulkLimit">Per Keyword</label>
          <select id="ebayBulkLimit"><option value="1">1</option><option value="2">2</option><option value="3" selected>3</option><option value="5">5</option><option value="10">10</option></select>
        </div>
        <div class="form-group" style="flex:1;min-width:120px;">
          <label for="ebayBulkCondition">Condition</label>
          <select id="ebayBulkCondition">
            <option value="">All</option>
            <option value="new">New</option>
            <option value="open_box">Open Box</option>
            <option value="refurbished">Refurbished</option>
            <option value="used">Used</option>
            <option value="for_parts">For Parts</option>
          </select>
        </div>
        <div class="form-group" style="flex:1;min-width:120px;">
          <label for="ebayBulkSort">Sort By</label>
          <select id="ebayBulkSort">
            <option value="best_match">Best Match</option>
            <option value="price_asc">Price: Low to High</option>
            <option value="price_desc">Price: High to Low</option>
          </select>
        </div>
        <div class="form-group" style="flex:1;min-width:100px;">
          <label for="ebayBulkItemLocation">Seller Location</label>
          <select id="ebayBulkItemLocation">
            <option value="">Anywhere</option>
            <option value="US">United States</option>
            <option value="JP">Japan</option>
            <option value="PK">Pakistan</option>
            <option value="GB">United Kingdom</option>
            <option value="DE">Germany</option>
            <option value="CA">Canada</option>
            <option value="AU">Australia</option>
            <option value="CN">China</option>
            <option value="SG">Singapore</option>
            <option value="KR">South Korea</option>
          </select>
        </div>
        <div class="form-group" style="flex:1;min-width:120px;">
          <label for="ebayBulkSite">eBay Site</label>
          <select id="ebayBulkSite">
            <option value="EBAY-US">United States</option>
            <option value="EBAY-GB">United Kingdom</option>
            <option value="EBAY-DE">Germany</option>
            <option value="EBAY-JAPAN">Japan</option>
            <option value="EBAY-AU">Australia</option>
            <option value="EBAY-CA">Canada</option>
          </select>
        </div>
        <div class="form-group" style="flex:1;min-width:80px;">
          <label for="ebayBulkMinPrice">Min $</label>
          <input type="number" id="ebayBulkMinPrice" placeholder="e.g. 50" step="0.01">
        </div>
        <div class="form-group" style="flex:1;min-width:80px;">
          <label for="ebayBulkMaxPrice">Max $</label>
          <input type="number" id="ebayBulkMaxPrice" placeholder="e.g. 2000" step="0.01">
        </div>
        <div class="form-group" style="flex:0;align-self:flex-end;margin-top:18px;display:flex;align-items:center;gap:6px;">
          <input type="checkbox" id="ebayBulkBINOnly" style="width:16px;height:16px;cursor:pointer;">
          <label for="ebayBulkBINOnly" style="margin:0;cursor:pointer;font-size:12px;white-space:nowrap;">BIN Only</label>
        </div>
        <button class="btn-search" id="ebayBulkBtn" onclick="doEbayBulkSearch()" style="flex:1;min-width:150px;margin-top:18px;background:#0064d2;">
          <span class="spinner"></span>
          <span class="btn-text">Search eBay Bulk</span>
        </button>
      </div>
      <div class="bulk-keyword-count" id="ebayBulkCount"></div>
    </div>
  </div>
  <div id="ebayBulkResults"></div>
</div>
<div class="tab-content" id="tabStockCheck">
  <div class="card">
    <h2>Check Out Of Stock</h2>
    <div class="bulk-area">
      <div class="form-row">
        <div class="form-group">
          <label for="stockKeywords">Keywords (one per line, paired by index)</label>
          <textarea id="stockKeywords" placeholder="Keyword for each link&#10;e.g.&#10;iPhone 15 Pro&#10;Canon PowerShot&#10;PS5" style="min-height:120px;"></textarea>
        </div>
        <div class="form-group">
          <label for="stockLinks">Product Links (one per line)</label>
          <textarea id="stockLinks" placeholder="Paste Yahoo / Mercari product links here&#10;e.g.&#10;https://auctions.yahoo.co.jp/jp/auction/1237842851&#10;https://jp.mercari.com/item/m37921031300" style="min-height:120px;"></textarea>
        </div>
      </div>
      <button class="btn-search" id="stockBtn" onclick="doStockCheck()" style="margin-top:4px;">
        <span class="spinner"></span>
        <span class="btn-text">Check Stock</span>
      </button>
      <div class="bulk-keyword-count" id="stockCount"></div>
    </div>
  </div>
  <div id="stockResults">
    <div class="status"><div class="emoji">&#x1F50D;</div><p>Enter keywords &amp; product links and click Check Stock</p></div>
  </div>
  <button class="dl-btn" id="stockDlBtn" onclick="downloadStockCSV()" style="display:none;margin-top:8px;">Download CSV</button>
</div>
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
      + ''+(item.image ? '<div class="result-thumb"><img src="'+escapeHtml(item.image)+'" alt="" loading="lazy" onerror="handleImageError(this)"></div>' : '<div class="result-icon">&#x1F4E6;</div>')
      +'<div class="result-info"><div class="result-name">'+(item.name_en ? escapeHtml(item.name_en) : escapeHtml(item.name))+'</div>'
      + '<div class="result-price">&yen;'+Number(item.price).toLocaleString()+' <span>JPY</span></div>'
        + (item.updated ? '<div class="result-updated">Updated: '+timeAgo(item.updated)+'</div>' : '')
        + (item.condition_en ? '<div class="result-condition">'+escapeHtml(item.condition_en)+'</div>' : '')
        + (item.seller_name ? '<div class="result-seller">Seller: '+escapeHtml(item.seller_name)+' ('+item.seller_reviews+' reviews)</div>' : '')
        + (item.status === 'ITEM_STATUS_SOLD_OUT' ? '<div class="result-tag sold">SOLD OUT</div>' : '')
      + (item.status === 'ITEM_STATUS_TRADING' ? '<div class="result-tag trading">TRADING</div>' : '')
      + (item.auction ? '<div class="result-tag auction">AUCTION</div>' : '')
      + '</div>'
      + '<div class="result-actions">'
      + '<button class="result-link desc-btn" data-item-id="'+escapeHtml(item.id)+'">Description</button>'
      + '<a class="result-link" href="'+escapeHtml(item.url)+'" target="_blank" rel="noopener">View &rarr;</a>'
      + '</div>'
      + '</div>';
  }
  html += '</div>';
  container.innerHTML = html;
}

function toggleBulkGroup(el) {
  const body = el.nextElementSibling;
  if (body) {
    body.style.display = body.style.display === 'none' ? 'block' : 'none';
  }
}


var JPY_RATE = 163.53;
function toJpyDisplay(usd) {
  var jpy = Math.round(Number(usd) * JPY_RATE);
  return '&nbsp;≈&nbsp;&yen;' + jpy.toLocaleString();
}

function handleImageError(img) {
  if (img && img.parentElement) {
    img.parentElement.className = 'result-icon';
    img.parentElement.innerHTML = '&#x1F4E6;';
  }
}

function closeModal(e) {
  if (e && e.target !== e.currentTarget) return;
  document.getElementById('modalOverlay').classList.remove('active');
}
async function fetchDescription(itemId) {
  const overlay = document.getElementById('modalOverlay');
  const body = document.getElementById('modalBody');
  body.innerHTML = '<div class="modal-loading">Loading &amp; translating...</div>';
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
  const targetId = 'tab' + name.charAt(0).toUpperCase() + name.slice(1);
  document.querySelectorAll('.tab-btn').forEach(function(btn) {
    btn.classList.remove('active');
    if (btn.getAttribute('data-tab') === name) btn.classList.add('active');
  });
  document.querySelectorAll('.tab-content').forEach(function(content) {
    content.classList.remove('active');
    if (content.id === targetId) content.classList.add('active');
  });
}

function searchEbayForKeyword(el) {
  const kw = el.getAttribute('data-kw');
  if (!kw) return;
  document.getElementById('ebayKeyword').value = kw;
  switchTab('ebay');
  setTimeout(doEbaySearch, 100);
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
  const n = v ? v.split('\\n').filter(function(l){ return l.trim(); }).length : 0;
  document.getElementById('bulkCount').textContent = n ? n+' keyword(s) loaded' : '';
}
document.getElementById('bulkKeywords').addEventListener('input', updateBulkCount);

async function streamBulkSearch(url, body, onResult, onComplete, onError) {
  const response = await fetch(url, {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(body)
  });
  if (!response.ok) {
    const e = await response.json().catch(function(){ return {}; });
    throw new Error(e.error || 'Search failed');
  }
  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';
  while (true) {
    const res = await reader.read();
    if (res.done) break;
    buffer += decoder.decode(res.value, {stream: true});
    const parts = buffer.split('\\n\\n');
    buffer = parts.pop();
    for (const part of parts) {
      for (const line of part.split('\\n')) {
        if (line.startsWith('data: ')) {
          try {
            const data = JSON.parse(line.slice(6));
            if (data.error) { if (onError) onError(data.error); return; }
            if (data.complete) { if (onComplete) onComplete(data); return; }
            if (onResult) onResult(data);
          } catch(e) {}
        }
      }
    }
  }
}

async function doBulkSearch() {
  const btn = document.getElementById('bulkBtn');
  const errDiv = document.getElementById('errorMsg');
  errDiv.style.display = 'none';
  const text = document.getElementById('bulkKeywords').value.trim();
  if (!text) { showError('Enter keywords or upload a file'); return; }
  const keywords = text.split('\\n').map(function(l){ return l.trim(); }).filter(function(l){ return l; });
  if (!keywords.length) { showError('No valid keywords found'); return; }
  if (keywords.length > 100) { showError('Maximum 100 keywords allowed'); return; }
  btn.classList.add('loading'); btn.disabled = true;
  const total = keywords.length;
  const container = document.getElementById('bulkResults');
  container.innerHTML = '<div class="count-badge streaming">Searching... 0/'+total+'</div><div id="mercariBulkGroups"></div>';
  window._bulkCSV = [];
  let doneCount = 0;
  try {
    await streamBulkSearch('/api/bulk-search', {
      keywords: keywords,
      per_keyword: parseInt(document.getElementById('bulkLimit').value) || 3,
      status_filter: document.getElementById('bulkStatus').value,
      condition: document.getElementById('bulkCondition').value,
      min_reviews: parseInt(document.getElementById('bulkMinReviews').value) || '',
      max_reviews: parseInt(document.getElementById('bulkMaxReviews').value) || '',
      exclude_auctions: document.getElementById('bulkExcludeAuctions').checked ? '1' : ''
    }, function(data) {
      appendBulkResult(data.keyword, data.items);
      doneCount++;
      const badge = container.querySelector('.count-badge');
      if (badge) badge.innerHTML = 'Searching... '+doneCount+'/'+total;
    }, function() {
      const badge = container.querySelector('.count-badge');
      if (badge) badge.innerHTML = 'Results for <strong>'+doneCount+'</strong> keyword(s)';
      const groups = document.getElementById('mercariBulkGroups');
      if (groups) {
        const dl = document.createElement('button');
        dl.className = 'dl-btn'; dl.style.marginTop = '12px';
        dl.textContent = 'Download CSV'; dl.onclick = downloadCSV;
        groups.appendChild(dl);
      }
    });
  } catch(e) { showError(e.message); }
  finally { btn.classList.remove('loading'); btn.disabled = false; }
}
function appendBulkResult(kw, items) {
  const groups = document.getElementById('mercariBulkGroups');
  if (!groups) return;
  if (!items || !items.length) return;
  let html = '<div class="bulk-result-group"><div class="bulk-result-header" onclick="toggleBulkGroup(this)">'
    + escapeHtml(kw) + ' <span class="count">'+items.length+' item(s)</span></div>'
    + '<div class="bulk-result-body"><div class="results">';
  for (const item of items) {
    const name = item.name_en || item.name;
    html += '<div class="result-card">'
      + (item.image ? '<div class="result-thumb"><img src="'+escapeHtml(item.image)+'" alt="" loading="lazy" onerror="handleImageError(this)"></div>' : '<div class="result-icon">&#x1F4E6;</div>')
      + '<div class="result-info"><div class="result-name">'+escapeHtml(name)+'</div>'
      + '<div class="result-price">&yen;'+Number(item.price).toLocaleString()+' <span>JPY</span></div>'
    + (item.updated ? '<div class="result-updated">Updated: '+timeAgo(item.updated)+'</div>' : '')
    + (item.condition_en ? '<div class="result-condition">'+escapeHtml(item.condition_en)+'</div>' : '')
    + (item.seller_name ? '<div class="result-seller">Seller: '+escapeHtml(item.seller_name)+' ('+item.seller_reviews+' reviews)</div>' : '')
      + (item.status === 'ITEM_STATUS_SOLD_OUT' ? '<div class="result-tag sold">SOLD OUT</div>' : '')
      + (item.auction ? '<div class="result-tag auction">AUCTION</div>' : '')
      + '</div>'
      + '<div class="result-actions">'
      + '<button class="result-link desc-btn" data-item-id="'+escapeHtml(item.id)+'">Desc</button>'
      + '<a class="result-link" href="'+escapeHtml(item.url)+'" target="_blank" rel="noopener">View</a>'
      + '<button class="result-link ebay-search-btn" data-kw="'+escapeHtml(name)+'" onclick="searchEbayForKeyword(this)">Search eBay</button>'
      + '</div></div>';
    window._bulkCSV.push({keyword: kw, name: name, price: item.price, url: item.url, status: item.status, condition_en: item.condition_en, seller_name: item.seller_name, seller_reviews: item.seller_reviews, updated: item.updated});
  }
  html += '</div></div></div>';
  groups.insertAdjacentHTML('beforeend', html);
}

function downloadCSV() {
  const data = window._bulkCSV || [];
  if (!data.length) return;
  let csv = '\uFEFFKeyword,Name,Price JPY,URL,Status,Condition,Seller,Reviews,Updated\\n';
  for (const r of data) {
    csv += '"'+r.keyword+'","'+(r.name||'').replace(/"/g,'""')+'",'+r.price+',"'+r.url+'","'+r.status+'","'+(r.condition_en||'')+'","'+(r.seller_name||'')+'","'+(r.seller_reviews||'')+'","'+(r.updated?timeAgo(r.updated):'')+'"\\n';
  }
  const blob = new Blob([csv], {type: 'text/csv;charset=utf-8;'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'mercari_bulk_results.csv'; a.click();
}

// Yahoo Auctions
document.getElementById('yahooKeyword').addEventListener('keydown', function(e) {
  if (e.key === 'Enter') doYahooSearch();
});

let _yahooItems = [];

async function doYahooSearch() {
  const btn = document.getElementById('yahooSearchBtn');
  const errDiv = document.getElementById('errorMsg');
  errDiv.style.display = 'none';
  const keyword = document.getElementById('yahooKeyword').value.trim();
  if (!keyword) { showError('Please enter a keyword'); return; }
  btn.classList.add('loading'); btn.disabled = true;
  try {
    const params = new URLSearchParams({
      keyword: keyword,
      min_price: document.getElementById('yahooMinPrice').value || '',
      max_price: document.getElementById('yahooMaxPrice').value || '',
      status: document.getElementById('yahooStatus').value,
      sort: document.getElementById('yahooSort').value,
      condition: document.getElementById('yahooCondition').value,
      bin_filter: document.getElementById('yahooBIN').value,
      limit: document.getElementById('yahooLimit').value || '20',
      page: '1'
    });
    const resp = await fetch('/api/yahoo-search?'+params.toString());
    if (!resp.ok) { const e = await resp.json(); throw new Error(e.error || 'Search failed'); }
    const data = await resp.json();
    renderYahooResults(data.items, keyword);
    _yahooItems = data.items || [];
    document.getElementById('yahooDlBtn').style.display = data.items && data.items.length ? 'inline-block' : 'none';
  } catch (e) { showError(e.message); }
  finally { btn.classList.remove('loading'); btn.disabled = false; }
}

function renderYahooResults(items, keyword) {
  const container = document.getElementById('yahooResultsContainer');
  if (!items || !items.length) {
    container.innerHTML = '<div class="status"><div class="emoji">&#x1F622;</div><p>No results found for "'+escapeHtml(keyword)+'"</p></div>';
    return;
  }
  let html = '<div class="count-badge">Found <strong>'+items.length+'</strong> item'+(items.length>1?'s':'')+'</div><div class="results">';
  for (const item of items) {
    const badges = [];
    if (item.status === 'sold') badges.push('<span class="yahoo-badge sold">SOLD</span>');
    if (item.free_shipping) badges.push('<span class="yahoo-badge freeship">FREE SHIP</span>');
    if (item.unused) badges.push('<span class="yahoo-badge unused">UNUSED</span>');
    if (item.is_new) badges.push('<span class="yahoo-badge newitem">NEW</span>');

    html += '<div class="yahoo-result-card">'
      + '<div class="yahoo-thumb">'
      + (item.thumbnail ? '<img src="'+escapeHtml(item.thumbnail)+'" alt="" loading="lazy" onerror="handleImageError(this)">' : '<div class="yahoo-thumb-placeholder">&#x1F4E6;</div>')
      + '</div>'
      + '<div class="yahoo-info">'
      + '<div class="yahoo-title">'+escapeHtml(item.title)+'</div>'
      + '<div class="yahoo-price-row">'
      + '<div class="yahoo-price">&yen;'+Number(item.price).toLocaleString()+' <span>JPY</span></div>'
      + (item.buy_now_price ? '<div class="yahoo-buynow">BIN &yen;'+Number(item.buy_now_price).toLocaleString()+'</div>' : '')
      + '</div>'
      + '<div class="yahoo-meta">'
      + '<span>&#x1F3F7; '+item.bid_count+' bid'+(item.bid_count!==1?'s':'')+'</span>'
      + (item.time_remaining ? '<span>&#x23F3; '+escapeHtml(item.time_remaining)+'</span>' : '')
      + (item.seller_name ? '<span>&#x1F464; '+escapeHtml(item.seller_name)+'</span>' : '<span>&#x1F464; '+escapeHtml(item.seller_id||'Unknown')+'</span>')
      + '</div>'
      + '<div>'+badges.join(' ')+'</div>'
      + '</div>'
      + '<div class="yahoo-actions">'
      + '<button class="result-link desc-btn" data-yahoo-id="'+escapeHtml(item.id)+'">Detail</button>'
      + '<a class="result-link" href="'+escapeHtml(item.url)+'" target="_blank" rel="noopener">View &rarr;</a>'
      + '</div>'
      + '</div>';
  }
  html += '</div>';
  container.innerHTML = html;
}

async function fetchYahooDescription(itemId) {
  const overlay = document.getElementById('modalOverlay');
  const body = document.getElementById('modalBody');
  body.innerHTML = '<div class="modal-loading">Loading item detail...</div>';
  overlay.classList.add('active');
  try {
    const resp = await fetch('/api/yahoo-item/'+encodeURIComponent(itemId));
    if (!resp.ok) throw new Error('Failed to fetch');
    const data = await resp.json();
    const displayDesc = data.description || 'No description available';
    const priceStr = data.price ? '&yen;'+Number(data.price).toLocaleString() : 'N/A';
    const bnStr = data.buy_now_price ? ' (BIN &yen;'+Number(data.buy_now_price).toLocaleString()+')' : '';
    body.innerHTML = '<h3>'+escapeHtml(data.title || data.id)+'</h3>'
      + '<div class="modal-meta">'
      + 'Price: '+priceStr+bnStr
      + (data.condition ? ' &middot; Condition: '+escapeHtml(data.condition) : '')
      + (data.seller_name ? ' &middot; Seller: '+escapeHtml(data.seller_name) : '')
      + (data.bid_count ? ' &middot; Bids: '+data.bid_count : '')
      + (data.end_time ? ' &middot; Ends: '+escapeHtml(data.end_time) : '')
      + '</div>'
      + '<div class="modal-desc">'+escapeHtml(displayDesc)+'</div>';
  } catch(e) {
    body.innerHTML = '<div class="modal-loading" style="color:#d42c25;">Error: '+e.message+'</div>';
  }
}

function downloadYahooCSV() {
  const data = _yahooItems;
  if (!data.length) return;
  let csv = '\uFEFFTitle,Price JPY,BuyNow JPY,Bids,Time Remaining,Seller,Status,Free Shipping,URL\\n';
  for (const r of data) {
    csv += '"'+(r.title||'').replace(/"/g,'""')+'",'+r.price+','+(r.buy_now_price||'')+','+r.bid_count+',"'+escapeHtml(r.time_remaining)+'","'+(r.seller_name||r.seller_id||'')+'","'+r.status+'","'+(r.free_shipping?'Yes':'No')+'","'+r.url+'"\\n';
  }
  const blob = new Blob([csv], {type: 'text/csv;charset=utf-8;'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'yahoo_auctions_results.csv'; a.click();
}

// Yahoo Bulk Search
document.getElementById('yahooBulkKeywords').addEventListener('input', updateYahooBulkCount);

function updateYahooBulkCount() {
  const v = document.getElementById('yahooBulkKeywords').value.trim();
  const n = v ? v.split('\\n').filter(function(l){ return l.trim(); }).length : 0;
  document.getElementById('yahooBulkCount').textContent = n ? n+' keyword(s) loaded' : '';
}

async function handleYahooBulkFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'txt' || ext === 'csv') {
    const text = await file.text();
    document.getElementById('yahooBulkKeywords').value = text;
  } else {
    const formData = new FormData();
    formData.append('file', file);
    try {
      const resp = await fetch('/api/upload-keywords', { method: 'POST', body: formData });
      if (!resp.ok) throw new Error('Upload failed');
      const data = await resp.json();
      if (data.keywords) document.getElementById('yahooBulkKeywords').value = data.keywords.join('\\n');
    } catch(e) { showError(e.message); }
  }
  updateYahooBulkCount();
}

async function doYahooBulkSearch() {
  const btn = document.getElementById('yahooBulkBtn');
  const errDiv = document.getElementById('errorMsg');
  errDiv.style.display = 'none';
  const text = document.getElementById('yahooBulkKeywords').value.trim();
  if (!text) { showError('Enter keywords or upload a file'); return; }
  const keywords = text.split('\\n').map(function(l){ return l.trim(); }).filter(function(l){ return l; });
  if (!keywords.length) { showError('No valid keywords found'); return; }
  if (keywords.length > 100) { showError('Maximum 100 keywords allowed'); return; }
  btn.classList.add('loading'); btn.disabled = true;
  const total = keywords.length;
  const container = document.getElementById('yahooBulkResults');
  container.innerHTML = '<div class="count-badge streaming">Searching... 0/'+total+'</div><div id="yahooBulkGroups"></div>';
  window._yahooBulkCSV = [];
  let doneCount = 0;
  try {
    await streamBulkSearch('/api/yahoo-bulk-search', {
      keywords: keywords,
      per_keyword: parseInt(document.getElementById('yahooBulkLimit').value) || 3,
      status: document.getElementById('yahooBulkStatus').value,
      condition: document.getElementById('yahooBulkCondition').value,
      sort: document.getElementById('yahooBulkSort').value,
      bin_filter: document.getElementById('yahooBulkBIN').value,
      min_price: document.getElementById('yahooBulkMinPrice').value || '',
      max_price: document.getElementById('yahooBulkMaxPrice').value || ''
    }, function(data) {
      appendYahooBulkResult(data.keyword, data.items);
      doneCount++;
      const badge = container.querySelector('.count-badge');
      if (badge) badge.innerHTML = 'Searching... '+doneCount+'/'+total;
    }, function() {
      const badge = container.querySelector('.count-badge');
      if (badge) badge.innerHTML = 'Results for <strong>'+doneCount+'</strong> keyword(s)';
      const groups = document.getElementById('yahooBulkGroups');
      if (groups) {
        const dl = document.createElement('button');
        dl.className = 'dl-btn'; dl.style.marginTop = '12px';
        dl.textContent = 'Download CSV'; dl.onclick = downloadYahooBulkCSV;
        groups.appendChild(dl);
      }
    });
  } catch(e) { showError(e.message); }
  finally { btn.classList.remove('loading'); btn.disabled = false; }
}

function appendYahooBulkResult(kw, items) {
  const groups = document.getElementById('yahooBulkGroups');
  if (!groups) return;
  if (!items || !items.length) return;
  let html = '<div class="bulk-result-group"><div class="bulk-result-header" onclick="toggleBulkGroup(this)">'
    + escapeHtml(kw) + ' <span class="count">'+items.length+' item(s)</span></div>'
    + '<div class="bulk-result-body"><div class="results">';
  for (const item of items) {
    const badges = [];
    if (item.status === 'sold') badges.push('<span class="yahoo-badge sold">SOLD</span>');
    if (item.free_shipping) badges.push('<span class="yahoo-badge freeship">FREE SHIP</span>');
    if (item.unused) badges.push('<span class="yahoo-badge unused">UNUSED</span>');
    if (item.is_new) badges.push('<span class="yahoo-badge newitem">NEW</span>');

    html += '<div class="yahoo-result-card">'
      + '<div class="yahoo-thumb">'
      + (item.thumbnail ? '<img src="'+escapeHtml(item.thumbnail)+'" alt="" loading="lazy" onerror="handleImageError(this)">' : '<div class="yahoo-thumb-placeholder">&#x1F4E6;</div>')
      + '</div>'
      + '<div class="yahoo-info">'
      + '<div class="yahoo-title">'+escapeHtml(item.title)+'</div>'
      + '<div class="yahoo-price-row">'
      + '<div class="yahoo-price">&yen;'+Number(item.price).toLocaleString()+' <span>JPY</span></div>'
      + (item.buy_now_price ? '<div class="yahoo-buynow">BIN &yen;'+Number(item.buy_now_price).toLocaleString()+'</div>' : '')
      + '</div>'
      + '<div class="yahoo-meta">'
      + '<span>&#x1F3F7; '+item.bid_count+' bid'+(item.bid_count!==1?'s':'')+'</span>'
      + (item.time_remaining ? '<span>&#x23F3; '+escapeHtml(item.time_remaining)+'</span>' : '')
      + '<span>&#x1F464; '+escapeHtml(item.seller_name || item.seller_id || 'Unknown')+'</span>'
      + '</div>'
      + '<div>'+badges.join(' ')+'</div>'
      + '</div>'
      + '<div class="yahoo-actions">'
      + '<a class="result-link" href="'+escapeHtml(item.url)+'" target="_blank" rel="noopener">View &rarr;</a>'
      + '</div>'
      + '</div>';
    window._yahooBulkCSV.push({keyword: kw, title: item.title, price: item.price, buy_now_price: item.buy_now_price, bid_count: item.bid_count, time_remaining: item.time_remaining, seller: item.seller_name || item.seller_id, status: item.status, free_shipping: item.free_shipping, url: item.url});
  }
  html += '</div></div></div>';
  groups.insertAdjacentHTML('beforeend', html);
}

function downloadYahooBulkCSV() {
  const data = window._yahooBulkCSV || [];
  if (!data.length) return;
  let csv = '\uFEFFKeyword,Title,Price JPY,BuyNow JPY,Bids,Time Remaining,Seller,Status,Free Shipping,URL\\n';
  for (const r of data) {
    csv += '"'+r.keyword+'","'+(r.title||'').replace(/"/g,'""')+'",'+r.price+','+(r.buy_now_price||'')+','+r.bid_count+',"'+escapeHtml(r.time_remaining)+'","'+(r.seller||'')+'","'+r.status+'","'+(r.free_shipping?'Yes':'No')+'","'+r.url+'"\\n';
  }
  const blob = new Blob([csv], {type: 'text/csv;charset=utf-8;'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'yahoo_bulk_results.csv'; a.click();
}

// Combined Mercari + Yahoo Bulk Search
document.getElementById('combinedKeywords').addEventListener('input', updateCombinedCount);

function updateCombinedCount() {
  const v = document.getElementById('combinedKeywords').value.trim();
  const n = v ? v.split('\\n').filter(function(l){ return l.trim(); }).length : 0;
  document.getElementById('combinedCount').textContent = n ? n+' keyword(s) loaded' : '';
}

async function handleCombinedFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'txt' || ext === 'csv') {
    const text = await file.text();
    document.getElementById('combinedKeywords').value = text;
  } else {
    const formData = new FormData();
    formData.append('file', file);
    try {
      const resp = await fetch('/api/upload-keywords', { method: 'POST', body: formData });
      if (!resp.ok) throw new Error('Upload failed');
      const data = await resp.json();
      if (data.keywords) document.getElementById('combinedKeywords').value = data.keywords.join('\\n');
    } catch(e) { showError(e.message); }
  }
  updateCombinedCount();
}

async function doCombinedBulkSearch() {
  const btn = document.getElementById('combinedBtn');
  const errDiv = document.getElementById('errorMsg');
  errDiv.style.display = 'none';
  const text = document.getElementById('combinedKeywords').value.trim();
  if (!text) { showError('Enter keywords or upload a file'); return; }
  const keywords = text.split('\\n').map(function(l){ return l.trim(); }).filter(function(l){ return l; });
  if (!keywords.length) { showError('No valid keywords found'); return; }
  if (keywords.length > 50) { showError('Maximum 50 keywords allowed'); return; }
  btn.classList.add('loading'); btn.disabled = true;
  const total = keywords.length;
  const container = document.getElementById('combinedResults');
  container.innerHTML = '<div class="count-badge streaming">Searching... 0/'+total+'</div><div id="combinedGroups"></div>';
  let doneCount = 0;
  try {
    await streamBulkSearch('/api/combined-bulk-search', {
      keywords: keywords,
      per_keyword: parseInt(document.getElementById('combinedLimit').value) || 3,
      mercari_status: document.getElementById('combinedStatus').value,
      mercari_condition: document.getElementById('combinedCondition').value,
      mercari_min_reviews: parseInt(document.getElementById('combinedMinReviews').value) || 0,
      mercari_exclude_auctions: document.getElementById('combinedExcludeAuctions').checked ? '1' : '',
      yahoo_status: document.getElementById('combinedYahooStatus').value,
      yahoo_condition: document.getElementById('combinedYahooCondition').value,
      yahoo_bin_filter: document.getElementById('combinedYahooBIN').value
    }, function(data) {
      appendCombinedResult(data.keyword, data.mercari_items || [], data.yahoo_items || []);
      doneCount++;
      const badge = container.querySelector('.count-badge');
      if (badge) badge.innerHTML = 'Searching... '+doneCount+'/'+total;
    }, function() {
      const badge = container.querySelector('.count-badge');
      if (badge) badge.innerHTML = 'Results for <strong>'+doneCount+'</strong> keyword(s)';
    });
  } catch(e) { showError(e.message); }
  finally { btn.classList.remove('loading'); btn.disabled = false; }
}

function appendCombinedResult(kw, mercariItems, yahooItems) {
  const groups = document.getElementById('combinedGroups');
  if (!groups) return;
  if (!mercariItems.length && !yahooItems.length) return;

  let html = '<div class="bulk-result-group">'
    + '<div class="bulk-result-header" onclick="toggleBulkGroup(this)">'
    + escapeHtml(kw) + ' <span class="count">Mercari: '+mercariItems.length+' &middot; Yahoo: '+yahooItems.length+'</span></div>'
    + '<div class="bulk-result-body"><div class="combined-grid">'
    + '<div class="combined-col"><h3 class="mercari-header">&#x1F4E6; Mercari</h3>';

  if (mercariItems.length) {
    for (const item of mercariItems) {
      const name = item.name_en || item.name;
      html += '<div class="result-card" style="margin-bottom:8px;">'
        + (item.image ? '<div class="result-thumb"><img src="'+escapeHtml(item.image)+'" alt="" loading="lazy" onerror="handleImageError(this)"></div>' : '<div class="result-icon">&#x1F4E6;</div>')
        + '<div class="result-info"><div class="result-name">'+escapeHtml(name)+'</div>'
        + '<div class="result-price">&yen;'+Number(item.price).toLocaleString()+' <span>JPY</span></div>'
        + (item.condition_en ? '<div class="result-condition">'+escapeHtml(item.condition_en)+'</div>' : '')
        + (item.seller_name ? '<div class="result-seller">'+escapeHtml(item.seller_name)+' ('+item.seller_reviews+' reviews)</div>' : '')
        + (item.status === 'ITEM_STATUS_SOLD_OUT' ? '<div class="result-tag sold">SOLD OUT</div>' : '')
        + '</div>'
        + '<div class="result-actions">'
        + '<button class="result-link desc-btn" data-item-id="'+escapeHtml(item.id)+'">Description</button>'
        + '<a class="result-link" href="'+escapeHtml(item.url)+'" target="_blank" rel="noopener">View</a>'
        + '<button class="result-link ebay-search-btn" data-kw="'+escapeHtml(kw)+'" onclick="searchEbayForKeyword(this)">Search eBay</button>'
        + '</div></div>';
    }
  } else {
    html += '<div class="status" style="padding:16px;font-size:12px;">&#x1F622; No Mercari results</div>';
  }

  html += '</div><div class="combined-col"><h3 class="yahoo-header">&#x1F4E6; Yahoo Auctions</h3>';

  if (yahooItems.length) {
    for (const item of yahooItems) {
      const badges = [];
      if (item.status === 'sold') badges.push('<span class="yahoo-badge sold">SOLD</span>');
      if (item.free_shipping) badges.push('<span class="yahoo-badge freeship">FREE SHIP</span>');

      html += '<div class="yahoo-result-card" style="margin-bottom:8px;">'
        + '<div class="yahoo-thumb" style="width:60px;min-height:60px;">'
        + (item.thumbnail ? '<img src="'+escapeHtml(item.thumbnail)+'" alt="" loading="lazy" onerror="handleImageError(this)">' : '<div class="yahoo-thumb-placeholder">&#x1F4E6;</div>')
        + '</div>'
        + '<div class="yahoo-info">'
        + '<div class="yahoo-title">'+escapeHtml(item.title)+'</div>'
        + '<div class="yahoo-price-row">'
        + '<div class="yahoo-price">&yen;'+Number(item.price).toLocaleString()+' <span>JPY</span></div>'
        + (item.buy_now_price ? '<div class="yahoo-buynow">BIN &yen;'+Number(item.buy_now_price).toLocaleString()+'</div>' : '')
        + '</div>'
        + '<div class="yahoo-meta">'
        + '<span>'+item.bid_count+' bid'+(item.bid_count!==1?'s':'')+'</span>'
        + (item.time_remaining ? '<span>'+escapeHtml(item.time_remaining)+'</span>' : '')
        + '</div>'
        + '<div>'+badges.join(' ')+'</div>'
        + '</div>'
        + '<div class="yahoo-actions">'
        + '<button class="result-link desc-btn" data-yahoo-id="'+escapeHtml(item.id)+'">Detail</button>'
        + '<a class="result-link" href="'+escapeHtml(item.url)+'" target="_blank" rel="noopener">View</a>'
        + '<button class="result-link ebay-search-btn" data-kw="'+escapeHtml(kw)+'" onclick="searchEbayForKeyword(this)">Search eBay</button>'
        + '</div></div>';
    }
  } else {
    html += '<div class="status" style="padding:16px;font-size:12px;">&#x1F622; No Yahoo results</div>';
  }

  html += '</div></div></div></div>';
  groups.insertAdjacentHTML('beforeend', html);
}

// Stock Check
document.getElementById('stockLinks').addEventListener('input', updateStockCount);

function updateStockCount() {
  const v = document.getElementById('stockLinks').value.trim();
  const n = v ? v.split('\\n').filter(function(l){ return l.trim(); }).length : 0;
  document.getElementById('stockCount').textContent = n ? n+' link(s) loaded' : '';
}

function showModalError(msg) {
  const overlay = document.getElementById('modalOverlay');
  const body = document.getElementById('modalBody');
  body.innerHTML = '<div style="text-align:center;padding:16px;">'
    + '<div style="font-size:48px;margin-bottom:12px;">&#x26A0;</div>'
    + '<h3 style="color:#d42c25;margin-bottom:12px;font-size:16px;">Error</h3>'
    + '<p style="color:#555;font-size:14px;line-height:1.6;white-space:pre-wrap;">'+escapeHtml(msg)+'</p>'
    + '</div>';
  overlay.classList.add('active');
}

async function doStockCheck() {
  const btn = document.getElementById('stockBtn');
  const errDiv = document.getElementById('errorMsg');
  errDiv.style.display = 'none';
  const kwText = document.getElementById('stockKeywords').value.trim();
  const linkText = document.getElementById('stockLinks').value.trim();
  if (!linkText) { showModalError('Paste product links first'); return; }
  const links = linkText.split('\\n').map(function(l){ return l.trim(); }).filter(function(l){ return l; });
  const keywords = kwText ? kwText.split('\\n').map(function(l){ return l.trim(); }).filter(function(l){ return l; }) : [];
  if (!links.length) { showModalError('No valid links found'); return; }
  if (links.length > 200) { showModalError('Maximum 200 links allowed'); return; }
  btn.classList.add('loading'); btn.disabled = true;
  const total = links.length;
  const container = document.getElementById('stockResults');
  container.innerHTML = '<div class="count-badge streaming">Checking... 0/'+total+'</div><div id="stockResultsList"></div>';
  window._stockCSV = [];
  window._stockErrors = [];
  _stockSno = 0;
  let doneCount = 0;
  try {
    await streamBulkSearch('/api/check-stock', { links: links, keywords: keywords }, function(data) {
      appendStockResult(data);
      doneCount++;
      const badge = container.querySelector('.count-badge');
      if (badge) badge.innerHTML = 'Checking... '+doneCount+'/'+total;
    }, function() {
      const badge = container.querySelector('.count-badge');
      if (badge) badge.innerHTML = 'Results for <strong>'+doneCount+'</strong> item(s)';
      document.getElementById('stockDlBtn').style.display = doneCount ? 'inline-block' : 'none';
      if (window._stockErrors.length) {
        showModalError('Completed with ' + window._stockErrors.length + ' error(s):\\n\\n' + window._stockErrors.join('\\n'));
      }
    });
  } catch(e) { showModalError(e.message); }
  finally { btn.classList.remove('loading'); btn.disabled = false; }
}

let _stockSno = 0;
function appendStockResult(data) {
  const list = document.getElementById('stockResultsList');
  if (!list) return;
  if (data.complete) return;
  _stockSno++;
  if (data.error) {
    const linkShort = data.link ? (data.link.length > 60 ? data.link.slice(0, 60) + '...' : data.link) : 'Unknown';
    window._stockErrors.push(_stockSno + '. ' + (data.keyword||'') + ' -> ' + linkShort + ' -> ' + data.error);
    list.insertAdjacentHTML('beforeend', '<div class="stock-card">'
      + '<span class="stock-sno">'+_stockSno+'</span>'
      + (data.keyword ? '<span class="stock-keyword" title="Keyword">'+escapeHtml(data.keyword)+'</span>' : '')
      + '<span class="stock-platform '+(data.platform || '')+'">'+(data.platform || '?').toUpperCase()+'</span>'
      + '<span class="stock-link-text">'+escapeHtml(data.link)+'</span>'
      + '<span class="stock-status error">Error: '+escapeHtml(data.error)+'</span>'
      + '</div>');
    return;
  }
  data.s_no = _stockSno;
  const statusClass = data.available ? 'available' : 'sold';
  const name = data.name || 'Unknown item';
  const priceStr = data.price ? '&yen;'+Number(data.price).toLocaleString() : '';
  const statusLabel = data.available_text || 'Unknown';
  list.insertAdjacentHTML('beforeend', '<div class="stock-card">'
    + '<span class="stock-sno">'+_stockSno+'</span>'
    + (data.keyword ? '<span class="stock-keyword" title="Keyword">'+escapeHtml(data.keyword)+'</span>' : '')
    + '<span class="stock-platform '+data.platform+'">'+data.platform.toUpperCase()+'</span>'
    + '<span class="stock-name" title="'+escapeHtml(data.status||'')+'">'+escapeHtml(name)+'</span>'
    + (priceStr ? '<span class="stock-price">'+priceStr+'</span>' : '')
    + '<span class="stock-status '+statusClass+'">'+statusLabel+'</span>'
    + '<a class="result-link" href="'+escapeHtml(data.link)+'" target="_blank" rel="noopener" style="font-size:12px;">View</a>'
    + '</div>');
  window._stockCSV.push(data);
}

function downloadStockCSV() {
  const data = window._stockCSV || [];
  if (!data.length) return;
  let csv = '\uFEFFS.No,Keyword,Link,Platform,ID,Name,Price JPY,Status,Available\\n';
  for (const r of data) {
    if (r.complete) continue;
    csv += ''+(r.s_no||'')+',"'+(r.keyword||'').replace(/"/g,'""')+'","'+(r.link||'').replace(/"/g,'""')+'","'+(r.platform||'')+'","'+(r.id||'')+'","'+(r.name||'').replace(/"/g,'""')+'",'+(r.price||'')+',"'+escapeHtml(r.status||'')+'","'+(r.available_text||'')+'"\\n';
  }
  const blob = new Blob([csv], {type: 'text/csv;charset=utf-8;'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'stock_check_results.csv'; a.click();
}

// eBay Settings
function saveEbayAppId() {
  const appId = document.getElementById('ebayAppId').value.trim();
  const certId = document.getElementById('ebayCertId').value.trim();
  localStorage.setItem('ebay_app_id', appId);
  localStorage.setItem('ebay_cert_id', certId);
  const status = document.getElementById('ebaySettingsStatus');
  status.style.display = 'inline';
  status.textContent = 'Saved' + (appId && certId ? '' : ' (incomplete)');
  setTimeout(function(){ status.style.display = 'none'; }, 2000);
}

function getEbayCredentials() {
  let appId = document.getElementById('ebayAppId').value.trim();
  let certId = document.getElementById('ebayCertId').value.trim();
  if (!appId) {
    const storedApp = localStorage.getItem('ebay_app_id');
    if (storedApp) { appId = storedApp; document.getElementById('ebayAppId').value = storedApp; }
  }
  if (!certId) {
    const storedCert = localStorage.getItem('ebay_cert_id');
    if (storedCert) { certId = storedCert; document.getElementById('ebayCertId').value = storedCert; }
  }
  return { appId: appId, certId: certId };
}

(function loadEbayCredentials() {
  const storedApp = localStorage.getItem('ebay_app_id');
  const storedCert = localStorage.getItem('ebay_cert_id');
  if (storedApp) document.getElementById('ebayAppId').value = storedApp;
  if (storedCert) document.getElementById('ebayCertId').value = storedCert;
})();

// eBay Simple Search
document.getElementById('ebayKeyword').addEventListener('keydown', function(e) {
  if (e.key === 'Enter') doEbaySearch();
});

let _ebayItems = [];

async function doEbaySearch() {
  const btn = document.getElementById('ebaySearchBtn');
  const errDiv = document.getElementById('errorMsg');
  errDiv.style.display = 'none';
  const keyword = document.getElementById('ebayKeyword').value.trim();
  if (!keyword) { showError('Please enter a keyword'); return; }
  const creds = getEbayCredentials();
  if (!creds.appId) { showError('Please set your eBay App ID in the settings bar above'); return; }
  if (!creds.certId) { showError('Please set your eBay Cert ID in the settings bar above'); return; }
  btn.classList.add('loading'); btn.disabled = true;
  try {
    const params = new URLSearchParams({
      keyword: keyword,
      app_id: creds.appId,
      cert_id: creds.certId,
      min_price: document.getElementById('ebayMinPrice').value || '',
      max_price: document.getElementById('ebayMaxPrice').value || '',
      condition: document.getElementById('ebayCondition').value,
      sort: document.getElementById('ebaySort').value,
      global_id: document.getElementById('ebaySite').value,
      item_location: document.getElementById('ebayItemLocation').value,
      bin_only: document.getElementById('ebayBINOnly').checked ? '1' : '',
      limit: document.getElementById('ebayLimit').value || '20'
    });
    const resp = await fetch('/api/ebay-search?'+params.toString());
    if (!resp.ok) { const e = await resp.json(); throw new Error(e.error || 'Search failed'); }
    const data = await resp.json();
    renderEbayResults(data.items, keyword);
    _ebayItems = data.items || [];
    document.getElementById('ebayDlBtn').style.display = data.items && data.items.length ? 'inline-block' : 'none';
  } catch (e) { showError(e.message); }
  finally { btn.classList.remove('loading'); btn.disabled = false; }
}

function getShippingHtml(item) {
  if (item.free_shipping) return '<span>&#x1F69A; Free Shipping</span>';
  if (item.shipping_type === 'CALCULATED') return '<span>&#x1F69A; Shipping: Calculated</span>';
  if (item.shipping_cost !== null && item.shipping_cost !== undefined) {
    const sym = item.currency === 'USD' ? '$' : (item.currency === 'JPY' ? '&yen;' : item.currency+' ');
    return '<span>&#x1F69A; Shipping: '+sym+Number(item.shipping_cost).toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})+'</span>';
  }
  return '';
}

function renderEbayResults(items, keyword) {
  const container = document.getElementById('ebayResultsContainer');
  if (!items || !items.length) {
    container.innerHTML = '<div class="status"><div class="emoji">&#x1F622;</div><p>No results found for "'+escapeHtml(keyword)+'"</p></div>';
    return;
  }
  let html = '<div class="count-badge">Found <strong>'+items.length+'</strong> item'+(items.length>1?'s':'')+'</div><div class="results">';
  for (const item of items) {
    const badges = [];
    if (item.free_shipping) badges.push('<span class="ebay-badge freeship">FREE SHIP</span>');
    if (item.top_rated_seller) badges.push('<span class="ebay-badge toprated">TOP RATED</span>');
    if (item.is_auction) badges.push('<span class="ebay-badge auction">AUCTION</span>');
    else badges.push('<span class="ebay-badge bin">BUY IT NOW</span>');

    const priceSymbol = item.currency === 'USD' ? '$' : (item.currency === 'JPY' ? '&yen;' : item.currency+' ');
    let totalHtml = '';
    if (item.free_shipping) {
      var t = Number(item.price);
      totalHtml = '<div class="result-price" style="font-size:13px;color:#2e7d32;">Total: '+priceSymbol+t.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})+' <span style="font-size:10px;">(incl. free ship)</span>'+(item.currency==='USD'?' '+toJpyDisplay(t):'')+'</div>';
    } else if (item.shipping_cost !== null && item.shipping_cost !== undefined) {
      var t = Number(item.price) + Number(item.shipping_cost);
      totalHtml = '<div class="result-price" style="font-size:13px;color:#555;">Total: '+priceSymbol+t.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})+' <span style="font-size:10px;">(incl. shipping)</span>'+(item.currency==='USD'?' '+toJpyDisplay(t):'')+'</div>';
    }
    html += '<div class="ebay-result-card">'
      + '<div class="ebay-thumb">'
      + (item.thumbnail ? '<img src="'+escapeHtml(item.thumbnail)+'" alt="" loading="lazy" onerror="handleImageError(this)">' : '<div class="ebay-thumb-placeholder">&#x1F4E6;</div>')
      + '</div>'
      + '<div class="ebay-info">'
      + '<div class="ebay-title">'+escapeHtml(item.title)+'</div>'
      + '<div class="ebay-price-row">'
      + '<div class="ebay-price">'+priceSymbol+Number(item.price).toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})+' <span>'+item.currency+'</span>'
      + (item.currency === 'USD' ? '<span style="font-size:11px;color:#888;">'+toJpyDisplay(item.price)+'</span>' : '')
      + '</div>'
      + (item.buy_it_now_price ? '<div class="ebay-bin-price">BIN '+priceSymbol+Number(item.buy_it_now_price).toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})+'</div>' : '')
      + '</div>'
      + totalHtml
      + '<div class="ebay-meta">'
      + (item.condition ? '<span>'+escapeHtml(item.condition)+'</span>' : '')
      + (item.bid_count !== null && item.bid_count !== undefined ? '<span>&#x1F3F7; '+item.bid_count+' bid'+(item.bid_count!==1?'s':'')+'</span>' : '')
      + (item.seller_name ? '<span>&#x1F464; '+escapeHtml(item.seller_name)+'</span>' : '')
      + (item.seller_rating ? '<span>&#x2B50; '+item.seller_rating+'</span>' : '')
      + (item.location ? '<span>&#x1F4CD; '+escapeHtml(item.location)+'</span>' : '')
      + getShippingHtml(item)
      + '</div>'
      + '<div>'+badges.join(' ')+'</div>'
      + '</div>'
      + '<div class="ebay-actions">'
      + '<a class="result-link" href="'+escapeHtml(item.url)+'" target="_blank" rel="noopener" style="border:1px solid #0064d2;color:#0064d2;background:white;">View &rarr;</a>'
      + '</div>'
      + '</div>';
  }
  html += '</div>';
  container.innerHTML = html;
}

function downloadEbayCSV() {
  const data = _ebayItems;
  if (!data.length) return;
  let csv = '\uFEFFTitle,Price,Currency,BuyNow,Bids,Condition,Seller,Rating,Location,Free Shipping,Shipping Cost,Total,Type,URL\\n';
  for (const r of data) {
    const total = r.free_shipping ? r.price : (r.shipping_cost != null ? Number(r.price) + Number(r.shipping_cost) : '');
    csv += '"'+(r.title||'').replace(/"/g,'""')+'",'+r.price+','+r.currency+','+(r.buy_it_now_price||'')+','+(r.bid_count!==null&&r.bid_count!==undefined?r.bid_count:'')+',"'+escapeHtml(r.condition)+'","'+(r.seller_name||'')+'","'+(r.seller_rating||'')+'","'+(r.location||'')+'","'+(r.free_shipping?'Yes':'No')+'","'+(r.shipping_cost||'')+'","'+total+'","'+(r.listing_type||'')+'","'+r.url+'"\\n';
  }
  const blob = new Blob([csv], {type: 'text/csv;charset=utf-8;'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'ebay_results.csv'; a.click();
}

// eBay Bulk Search
document.getElementById('ebayBulkKeywords').addEventListener('input', updateEbayBulkCount);

function updateEbayBulkCount() {
  const v = document.getElementById('ebayBulkKeywords').value.trim();
  const n = v ? v.split('\\n').filter(function(l){ return l.trim(); }).length : 0;
  document.getElementById('ebayBulkCount').textContent = n ? n+' keyword(s) loaded' : '';
}

async function handleEbayBulkFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'txt' || ext === 'csv') {
    const text = await file.text();
    document.getElementById('ebayBulkKeywords').value = text;
  } else {
    const formData = new FormData();
    formData.append('file', file);
    try {
      const resp = await fetch('/api/upload-keywords', { method: 'POST', body: formData });
      if (!resp.ok) throw new Error('Upload failed');
      const data = await resp.json();
      if (data.keywords) document.getElementById('ebayBulkKeywords').value = data.keywords.join('\\n');
    } catch(e) { showError(e.message); }
  }
  updateEbayBulkCount();
}

async function doEbayBulkSearch() {
  const btn = document.getElementById('ebayBulkBtn');
  const errDiv = document.getElementById('errorMsg');
  errDiv.style.display = 'none';
  const creds = getEbayCredentials();
  if (!creds.appId) { showError('Please set your eBay App ID in the settings bar above'); return; }
  if (!creds.certId) { showError('Please set your eBay Cert ID in the settings bar above'); return; }
  const text = document.getElementById('ebayBulkKeywords').value.trim();
  if (!text) { showError('Enter keywords or upload a file'); return; }
  const keywords = text.split('\\n').map(function(l){ return l.trim(); }).filter(function(l){ return l; });
  if (!keywords.length) { showError('No valid keywords found'); return; }
  if (keywords.length > 100) { showError('Maximum 100 keywords allowed'); return; }
  btn.classList.add('loading'); btn.disabled = true;
  const total = keywords.length;
  const container = document.getElementById('ebayBulkResults');
  container.innerHTML = '<div class="count-badge streaming">Searching... 0/'+total+'</div><div id="ebayBulkGroups"></div>';
  window._ebayBulkCSV = [];
  let doneCount = 0;
  try {
    await streamBulkSearch('/api/ebay-bulk-search', {
      keywords: keywords,
      app_id: creds.appId,
      cert_id: creds.certId,
      per_keyword: parseInt(document.getElementById('ebayBulkLimit').value) || 3,
      condition: document.getElementById('ebayBulkCondition').value,
      sort: document.getElementById('ebayBulkSort').value,
      global_id: document.getElementById('ebayBulkSite').value,
      item_location: document.getElementById('ebayBulkItemLocation').value,
      min_price: document.getElementById('ebayBulkMinPrice').value || '',
      max_price: document.getElementById('ebayBulkMaxPrice').value || '',
      bin_only: document.getElementById('ebayBulkBINOnly').checked ? '1' : ''
    }, function(data) {
      appendEbayBulkResult(data.keyword, data.items);
      doneCount++;
      const badge = container.querySelector('.count-badge');
      if (badge) badge.innerHTML = 'Searching... '+doneCount+'/'+total;
    }, function() {
      const badge = container.querySelector('.count-badge');
      if (badge) badge.innerHTML = 'Results for <strong>'+doneCount+'</strong> keyword(s)';
      const groups = document.getElementById('ebayBulkGroups');
      if (groups) {
        const dl = document.createElement('button');
        dl.className = 'dl-btn'; dl.style.marginTop = '12px'; dl.style.background = '#0064d2';
        dl.textContent = 'Download CSV'; dl.onclick = downloadEbayBulkCSV;
        groups.appendChild(dl);
      }
    });
  } catch(e) { showError(e.message); }
  finally { btn.classList.remove('loading'); btn.disabled = false; }
}

function appendEbayBulkResult(kw, items) {
  const groups = document.getElementById('ebayBulkGroups');
  if (!groups) return;
  if (!items || !items.length) return;
  let html = '<div class="bulk-result-group"><div class="bulk-result-header" onclick="toggleBulkGroup(this)">'
    + escapeHtml(kw) + ' <span class="count">'+items.length+' item(s)</span></div>'
    + '<div class="bulk-result-body"><div class="results">';
  for (const item of items) {
    const badges = [];
    if (item.free_shipping) badges.push('<span class="ebay-badge freeship">FREE SHIP</span>');
    if (item.top_rated_seller) badges.push('<span class="ebay-badge toprated">TOP RATED</span>');
    if (item.is_auction) badges.push('<span class="ebay-badge auction">AUCTION</span>');
    else badges.push('<span class="ebay-badge bin">BIN</span>');

    const priceSymbol = item.currency === 'USD' ? '$' : (item.currency === 'JPY' ? '&yen;' : item.currency+' ');
    let totalHtml = '';
    if (item.free_shipping) {
      var t = Number(item.price);
      totalHtml = '<div style="font-size:12px;color:#2e7d32;">Total: '+priceSymbol+t.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})+' <span style="font-size:10px;">(incl. free ship)</span>'+(item.currency==='USD'?' '+toJpyDisplay(t):'')+'</div>';
    } else if (item.shipping_cost !== null && item.shipping_cost !== undefined) {
      var t = Number(item.price) + Number(item.shipping_cost);
      totalHtml = '<div style="font-size:12px;color:#555;">Total: '+priceSymbol+t.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})+' <span style="font-size:10px;">(incl. shipping)</span>'+(item.currency==='USD'?' '+toJpyDisplay(t):'')+'</div>';
    }
    html += '<div class="ebay-result-card" style="margin-bottom:8px;">'
      + '<div class="ebay-thumb" style="width:60px;min-height:60px;">'
      + (item.thumbnail ? '<img src="'+escapeHtml(item.thumbnail)+'" alt="" loading="lazy" onerror="handleImageError(this)">' : '<div class="ebay-thumb-placeholder">&#x1F4E6;</div>')
      + '</div>'
      + '<div class="ebay-info">'
      + '<div class="ebay-title">'+escapeHtml(item.title)+'</div>'
      + '<div class="ebay-price-row">'
      + '<div class="ebay-price" style="font-size:15px;">'+priceSymbol+Number(item.price).toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})+' <span style="font-size:11px;color:#888;">'+(item.currency==='USD'?toJpyDisplay(item.price):'')+'</span></div>'
      + (item.buy_it_now_price ? '<div class="ebay-bin-price">BIN '+priceSymbol+Number(item.buy_it_now_price).toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})+'</div>' : '')
      + '</div>'
      + totalHtml
      + '<div class="ebay-meta">'
      + (item.condition ? '<span>'+escapeHtml(item.condition)+'</span>' : '')
      + (item.bid_count !== null && item.bid_count !== undefined ? '<span>'+item.bid_count+' bids</span>' : '')
      + getShippingHtml(item)
      + '</div>'
      + '<div>'+badges.join(' ')+'</div>'
      + '</div>'
      + '<div class="ebay-actions">'
      + '<a class="result-link" href="'+escapeHtml(item.url)+'" target="_blank" rel="noopener" style="border:1px solid #0064d2;color:#0064d2;background:white;font-size:12px;padding:6px 10px;">View</a>'
      + '</div></div>';
    window._ebayBulkCSV.push({keyword: kw, title: item.title, price: item.price, currency: item.currency, buy_it_now_price: item.buy_it_now_price, bid_count: item.bid_count, condition: item.condition, seller_name: item.seller_name, location: item.location, free_shipping: item.free_shipping, shipping_cost: item.shipping_cost, listing_type: item.listing_type, url: item.url});
  }
  html += '</div></div></div>';
  groups.insertAdjacentHTML('beforeend', html);
}

function downloadEbayBulkCSV() {
  const data = window._ebayBulkCSV || [];
  if (!data.length) return;
  let csv = '\uFEFFKeyword,Title,Price,Currency,BuyNow,Bids,Condition,Seller,Location,Free Shipping,Shipping Cost,Total,Type,URL\\n';
  for (const r of data) {
    const total = r.free_shipping ? r.price : (r.shipping_cost != null ? Number(r.price) + Number(r.shipping_cost) : '');
    csv += '"'+r.keyword+'","'+(r.title||'').replace(/"/g,'""')+'",'+r.price+','+r.currency+','+(r.buy_it_now_price||'')+','+(r.bid_count!==null&&r.bid_count!==undefined?r.bid_count:'')+',"'+escapeHtml(r.condition)+'","'+(r.seller_name||'')+'","'+(r.location||'')+'","'+(r.free_shipping?'Yes':'No')+'","'+(r.shipping_cost||'')+'","'+total+'","'+(r.listing_type||'')+'","'+r.url+'"\\n';
  }
  const blob = new Blob([csv], {type: 'text/csv;charset=utf-8;'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'ebay_bulk_results.csv'; a.click();
}

// Event delegation for description buttons
document.addEventListener('click', function(e) {
  const mercariBtn = e.target.closest('[data-item-id]');
  if (mercariBtn) {
    fetchDescription(mercariBtn.getAttribute('data-item-id'));
    return;
  }
  const yahooBtn = e.target.closest('[data-yahoo-id]');
  if (yahooBtn) {
    fetchYahooDescription(yahooBtn.getAttribute('data-yahoo-id'));
    return;
  }
});
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
            exclude_keywords=exclude_keywords,
            max_items=limit * 3
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
                image_url = getattr(item, 'imageURL', '') or (item.thumbnails[0] if getattr(item, 'thumbnails', None) else '')

                # Detect shop item from search result (no extra API call)
                if getattr(item, 'isShopItem', False):
                    url = 'https://jp.mercari.com/shops/product/' + item_id

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
                    'seller_name': '',
                    'seller_reviews': 0,
                    'image': image_url
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
        exclude_auctions = data.get('exclude_auctions', '') in ('1', 'true', True)
        COND_IDS = {'1': '新品、未使用', '2': '未使用に近い', '3': '目立った傷や汚れなし', '4': 'やや傷や汚れあり', '5': '傷や汚れあり', '6': '全体的に状態が悪い'}
        COND_EN = {
            '新品、未使用': 'NEW, UNUSED',
            '未使用に近い': 'LIKE NEW, UNUSED',
            '目立った傷や汚れなし': 'NO NOTICEABLE SCRATCHES OR DIRT',
            'やや傷や汚れあり': 'SLIGHT SCRATCHES OR DIRT',
            '傷や汚れあり': 'SCRATCHES OR DIRT',
            '全体的に状態が悪い': 'POOR CONDITION'
        }

        def search_keyword(kw):
            try:
                needs_review_filter = min_reviews is not None or max_reviews is not None
                search_results = search(kw, sort=MercariSort.SORT_SCORE, order=MercariOrder.ORDER_DESC, status=MercariSearchStatus.ON_SALE, max_items=per_keyword * 3)
                items = []
                for item in search_results:
                    name = getattr(item, 'productName', None)
                    price = getattr(item, 'price', None)
                    if not name or price is None:
                        continue
                    if exclude_auctions and getattr(item, 'auction', None) is not None:
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
                    item_id = getattr(item, 'id', '')
                    item_url = getattr(item, 'productURL', '')
                    image_url = getattr(item, 'imageURL', '') or (item.thumbnails[0] if getattr(item, 'thumbnails', None) else '')
                    if getattr(item, 'isShopItem', False):
                        item_url = 'https://jp.mercari.com/shops/product/' + item_id
                    items.append({
                        'id': item_id,
                        'name': name,
                        'price': float(price),
                        'image': image_url,
                        'url': item_url,
                        'status': item_status,
                        'auction': getattr(item, 'auction', None) is not None,
                        'condition': cond_name,
                        'updated': getattr(item, 'updated', 0)
                    })
                    if len(items) >= per_keyword:
                        break

                if not items:
                    return kw, []

                if needs_review_filter:
                    with ThreadPoolExecutor(max_workers=5) as info_pool:
                        def enrich(item):
                            try:
                                full_info = getItemInfo(item['id'])
                                if full_info is not None:
                                    if hasattr(full_info, 'seller') and full_info.seller is not None:
                                        item['seller_name'] = full_info.seller.name or ''
                                        item['seller_reviews'] = full_info.seller.num_ratings or 0
                                    if hasattr(full_info, 'item_condition') and full_info.item_condition is not None:
                                        item['condition'] = full_info.item_condition.name or item['condition']
                                    if hasattr(full_info, 'is_shop_item') and str(full_info.is_shop_item).lower() in ('true', '1'):
                                        item['url'] = 'https://jp.mercari.com/shops/product/' + item['id']
                            except Exception:
                                item.setdefault('seller_name', '')
                                item.setdefault('seller_reviews', 0)
                            return item
                        futures = [info_pool.submit(enrich, it) for it in items]
                        items = [f.result() for f in as_completed(futures)]

                    if min_reviews is not None:
                        items = [e for e in items if e.get('seller_reviews', 0) >= min_reviews]
                    if max_reviews is not None:
                        items = [e for e in items if e.get('seller_reviews', 0) <= max_reviews]

                for it in items:
                    it.setdefault('seller_name', '')
                    it.setdefault('seller_reviews', 0)

                names = [it['name'] for it in items]
                translations = batch_translate(names)
                for i, it in enumerate(items):
                    it['name_en'] = translations[i] if i < len(translations) else ''
                    it['condition_en'] = COND_EN.get(it['condition'], it['condition'])

                return kw, items
            except Exception:
                return kw, []

        keywords_list = [k.strip() for k in keywords[:100] if k.strip()]

        def generate():
            with ThreadPoolExecutor(max_workers=5) as kw_pool:
                futures = {kw_pool.submit(search_keyword, kw): kw for kw in keywords_list}
                for f in as_completed(futures):
                    kw, items = f.result()
                    yield 'data: ' + json.dumps({'keyword': kw, 'items': items}, ensure_ascii=False) + '\n\n'
            yield 'data: ' + json.dumps({'complete': True, 'keywords': keywords_list}, ensure_ascii=False) + '\n\n'

        return Response(stream_with_context(generate()), mimetype='text/event-stream')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/yahoo-search')
def api_yahoo_search():
    keyword = request.args.get('keyword', '').strip()
    if not keyword:
        return jsonify({'error': 'Keyword is required'}), 400

    min_price = request.args.get('min_price', '').strip()
    max_price = request.args.get('max_price', '').strip()
    status = request.args.get('status', 'live').strip()
    condition = request.args.get('condition', '').strip()
    sort = request.args.get('sort', 'new').strip()
    bin_filter = request.args.get('bin_filter', 'all').strip()
    page = request.args.get('page', '1').strip()
    limit = request.args.get('limit', '20').strip()

    try:
        limit = int(limit) if limit else 20
    except ValueError:
        limit = 20
    try:
        page = int(page) if page else 1
    except ValueError:
        page = 1

    min_price_val = None
    max_price_val = None
    try:
        min_price_val = int(min_price) if min_price else None
    except ValueError:
        pass
    try:
        max_price_val = int(max_price) if max_price else None
    except ValueError:
        pass

    try:
        results = search_yahoo(
            keyword=keyword,
            min_price=min_price_val,
            max_price=max_price_val,
            status=status,
            condition=condition,
            sort=sort,
            page=page,
            limit=limit,
        )

        items = []
        for item in results:
            if bin_filter == 'bin_only' and not item.buy_now_price:
                continue
            items.append({
                'id': item.id,
                'title': item.title,
                'price': item.price,
                'buy_now_price': item.buy_now_price,
                'bid_count': item.bid_count,
                'time_remaining': item.time_remaining,
                'end_timestamp': item.end_timestamp,
                'seller_id': item.seller_id,
                'thumbnail': item.thumbnail,
                'url': item.url,
                'free_shipping': item.free_shipping,
                'unused': item.unused,
                'is_new': item.is_new,
                'status': item.status,
                'seller_name': item.seller_name,
            })

        return jsonify({'items': items, 'count': len(items)})

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/yahoo-bulk-search', methods=['POST'])
def api_yahoo_bulk_search():
    try:
        data = request.get_json()
        if not data or 'keywords' not in data:
            return jsonify({'error': 'Keywords required'}), 400
        keywords = data['keywords']
        if not isinstance(keywords, list) or not keywords:
            return jsonify({'error': 'Keywords must be a non-empty array'}), 400
        per_keyword = int(data.get('per_keyword', 3))
        per_keyword = max(1, min(per_keyword, 20))
        status = data.get('status', 'live')
        condition = data.get('condition', '')
        sort = data.get('sort', 'new')
        bin_filter = data.get('bin_filter', 'all')
        min_price = data.get('min_price', '')
        max_price = data.get('max_price', '')

        min_price_val = int(min_price) if min_price else None
        max_price_val = int(max_price) if max_price else None

        def search_keyword(kw):
            try:
                results = search_yahoo(
                    keyword=kw,
                    min_price=min_price_val,
                    max_price=max_price_val,
                    status=status,
                    condition=condition,
                    sort=sort,
                    page=1,
                    limit=per_keyword,
                )
                items = []
                for item in results[:per_keyword]:
                    if bin_filter == 'bin_only' and not item.buy_now_price:
                        continue
                    items.append({
                        'id': item.id,
                        'title': item.title,
                        'price': item.price,
                        'buy_now_price': item.buy_now_price,
                        'bid_count': item.bid_count,
                        'time_remaining': item.time_remaining,
                        'seller_id': item.seller_id,
                        'seller_name': item.seller_name,
                        'thumbnail': item.thumbnail,
                        'url': item.url,
                        'free_shipping': item.free_shipping,
                        'unused': item.unused,
                        'is_new': item.is_new,
                        'status': item.status,
                    })
                return kw, items
            except Exception:
                return kw, []

        keywords_list = [k.strip() for k in keywords[:100] if k.strip()]

        def generate():
            with ThreadPoolExecutor(max_workers=5) as pool:
                futures = {pool.submit(search_keyword, kw): kw for kw in keywords_list}
                for f in as_completed(futures):
                    kw, items = f.result()
                    yield 'data: ' + json.dumps({'keyword': kw, 'items': items}, ensure_ascii=False) + '\n\n'
            yield 'data: ' + json.dumps({'complete': True, 'keywords': keywords_list}, ensure_ascii=False) + '\n\n'

        return Response(stream_with_context(generate()), mimetype='text/event-stream')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/yahoo-item/<item_id>')
def api_yahoo_item_detail(item_id):
    try:
        detail = get_item_detail(item_id)
        return jsonify(detail)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/combined-bulk-search', methods=['POST'])
def api_combined_bulk_search():
    try:
        data = request.get_json()
        if not data or 'keywords' not in data:
            return jsonify({'error': 'Keywords required'}), 400
        keywords = data['keywords']
        if not isinstance(keywords, list) or not keywords:
            return jsonify({'error': 'Keywords must be a non-empty array'}), 400
        per_keyword = int(data.get('per_keyword', 3))
        per_keyword = max(1, min(per_keyword, 10))
        mercari_status = data.get('mercari_status', 'all')
        mercari_condition = data.get('mercari_condition', '')
        yahoo_status = data.get('yahoo_status', 'live')
        yahoo_condition = data.get('yahoo_condition', '')
        yahoo_bin_filter = data.get('yahoo_bin_filter', 'all')
        mercari_min_reviews = int(data.get('mercari_min_reviews', 0)) or 0
        mercari_exclude_auctions = data.get('mercari_exclude_auctions', '') in ('1', 'true', True)

        keywords_list = [k.strip() for k in keywords[:50] if k.strip()]

        COND_EN = {
            '新品、未使用': 'NEW, UNUSED',
            '未使用に近い': 'LIKE NEW, UNUSED',
            '目立った傷や汚れなし': 'NO NOTICEABLE SCRATCHES OR DIRT',
            'やや傷や汚れあり': 'SLIGHT SCRATCHES OR DIRT',
            '傷や汚れあり': 'SCRATCHES OR DIRT',
            '全体的に状態が悪い': 'POOR CONDITION'
        }
        COND_IDS = {'1': '新品、未使用', '2': '未使用に近い', '3': '目立った傷や汚れなし', '4': 'やや傷や汚れあり', '5': '傷や汚れあり', '6': '全体的に状態が悪い'}

        def search_mercari_kw(kw):
            try:
                search_results = search(kw, sort=MercariSort.SORT_SCORE, order=MercariOrder.ORDER_DESC, status=MercariSearchStatus.ON_SALE, max_items=per_keyword * 3)
                items = []
                for item in search_results:
                    name = getattr(item, 'productName', None)
                    price = getattr(item, 'price', None)
                    if not name or price is None:
                        continue
                    if mercari_exclude_auctions and getattr(item, 'auction', None) is not None:
                        continue
                    item_status = getattr(item, 'status', '')
                    if mercari_status == 'available' and item_status != 'ITEM_STATUS_ON_SALE':
                        continue
                    if mercari_status == 'sold' and item_status != 'ITEM_STATUS_SOLD_OUT':
                        continue
                    cond_id = str(getattr(item, 'itemConditionId', None) or '')
                    cond_name = COND_IDS.get(cond_id, '')
                    if mercari_condition and cond_name != mercari_condition:
                        continue
                    item_id = getattr(item, 'id', '')
                    item_url = getattr(item, 'productURL', '')
                    if getattr(item, 'isShopItem', False):
                        item_url = 'https://jp.mercari.com/shops/product/' + item_id
                    image_url = getattr(item, 'imageURL', '') or (item.thumbnails[0] if getattr(item, 'thumbnails', None) else '')
                    items.append({
                        'id': item_id,
                        'name': name,
                        'price': float(price),
                        'image': image_url,
                        'url': item_url,
                        'status': item_status,
                        'auction': getattr(item, 'auction', None) is not None,
                        'condition': cond_name,
                        'condition_en': COND_EN.get(cond_name, cond_name),
                        'updated': getattr(item, 'updated', 0)
                    })
                    if len(items) >= per_keyword:
                        break

                if items:
                    with ThreadPoolExecutor(max_workers=5) as info_pool:
                        def enrich(item):
                            try:
                                full_info = getItemInfo(item['id'])
                                if full_info is not None:
                                    if hasattr(full_info, 'seller') and full_info.seller is not None:
                                        item['seller_name'] = full_info.seller.name or ''
                                        item['seller_reviews'] = full_info.seller.num_ratings or 0
                                    if hasattr(full_info, 'item_condition') and full_info.item_condition is not None:
                                        item['condition'] = full_info.item_condition.name or item['condition']
                                        item['condition_en'] = COND_EN.get(item['condition'], item['condition'])
                                    if hasattr(full_info, 'is_shop_item') and str(full_info.is_shop_item).lower() in ('true', '1'):
                                        item['url'] = 'https://jp.mercari.com/shops/product/' + item['id']
                            except Exception:
                                item.setdefault('seller_name', '')
                                item.setdefault('seller_reviews', 0)
                            return item
                        futures = [info_pool.submit(enrich, it) for it in items]
                        items = [f.result() for f in as_completed(futures)]

                    if mercari_min_reviews > 0:
                        items = [it for it in items if it.get('seller_reviews', 0) >= mercari_min_reviews]

                    names = [it['name'] for it in items]
                    translations = batch_translate(names)
                    for i, it in enumerate(items):
                        it['name_en'] = translations[i] if i < len(translations) else ''
                return items
            except Exception:
                return []

        def search_yahoo_kw(kw):
            try:
                results = search_yahoo(
                    keyword=kw,
                    status=yahoo_status,
                    condition=yahoo_condition,
                    sort='new',
                    page=1,
                    limit=per_keyword,
                )
                yahoo_items = []
                for item in results[:per_keyword]:
                    if yahoo_bin_filter == 'bin_only' and not item.buy_now_price:
                        continue
                    yahoo_items.append({
                        'id': item.id,
                        'title': item.title,
                        'price': item.price,
                        'buy_now_price': item.buy_now_price,
                        'bid_count': item.bid_count,
                        'time_remaining': item.time_remaining,
                        'seller_id': item.seller_id,
                        'seller_name': item.seller_name,
                        'thumbnail': item.thumbnail,
                        'url': item.url,
                        'free_shipping': item.free_shipping,
                        'unused': item.unused,
                        'is_new': item.is_new,
                        'status': item.status,
                    })
                return yahoo_items
            except Exception:
                return []

        keywords_done = set()
        mercari_temp = {}
        yahoo_temp = {}

        def generate():
            nonlocal mercari_temp, yahoo_temp
            with ThreadPoolExecutor(max_workers=10) as pool:
                mercari_futures = {pool.submit(search_mercari_kw, kw): kw for kw in keywords_list}
                yahoo_futures = {pool.submit(search_yahoo_kw, kw): kw for kw in keywords_list}
                all_futures = list(mercari_futures.keys()) + list(yahoo_futures.keys())
                for f in as_completed(all_futures):
                    kw = mercari_futures.get(f) or yahoo_futures.get(f)
                    if f in mercari_futures:
                        mercari_temp[kw] = f.result()
                    else:
                        yahoo_temp[kw] = f.result()
                    if kw not in keywords_done and kw in mercari_temp and kw in yahoo_temp:
                        keywords_done.add(kw)
                        yield 'data: ' + json.dumps({
                            'keyword': kw,
                            'mercari_items': mercari_temp.get(kw, []),
                            'yahoo_items': yahoo_temp.get(kw, [])
                        }, ensure_ascii=False) + '\n\n'
                for kw in keywords_list:
                    if kw not in keywords_done:
                        yield 'data: ' + json.dumps({
                            'keyword': kw,
                            'mercari_items': mercari_temp.get(kw, []),
                            'yahoo_items': yahoo_temp.get(kw, [])
                        }, ensure_ascii=False) + '\n\n'
            yield 'data: ' + json.dumps({'complete': True, 'keywords': keywords_list}, ensure_ascii=False) + '\n\n'

        return Response(stream_with_context(generate()), mimetype='text/event-stream')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/check-stock', methods=['POST'])
def api_check_stock():
    try:
        data = request.get_json()
        if not data or 'links' not in data:
            return jsonify({'error': 'Links required'}), 400
        links = data['links']
        if not isinstance(links, list) or not links:
            return jsonify({'error': 'Links must be a non-empty array'}), 400

        raw_keywords = data.get('keywords', [])
        if not isinstance(raw_keywords, list):
            raw_keywords = []

        links = [l.strip() for l in links[:200] if l.strip()]
        keywords = [k.strip() for k in raw_keywords[:200] if k.strip()]
        import re

        def parse_link(link):
            yahoo_match = re.search(r'auctions\.yahoo\.co\.jp/(?:jp/)?auction/([a-zA-Z0-9_]+)', link)
            if yahoo_match:
                return 'yahoo', yahoo_match.group(1)
            yahoo_match2 = re.search(r'page\.auctions\.yahoo\.co\.jp/auction/([a-zA-Z0-9_]+)', link)
            if yahoo_match2:
                return 'yahoo', yahoo_match2.group(1)
            paypay_match = re.search(r'paypayfleamarket\.yahoo\.co\.jp/item/([a-zA-Z0-9_]+)', link)
            if paypay_match:
                return 'paypay', paypay_match.group(1)
            mercari_shop_match = re.search(r'(?:jp\.)?mercari\.com/shops/product/([a-zA-Z0-9_]+)', link)
            if mercari_shop_match:
                return 'mercari_shop', mercari_shop_match.group(1)
            mercari_match = re.search(r'(?:jp\.)?mercari\.com/item/([a-zA-Z0-9_]+)', link)
            if mercari_match:
                return 'mercari', mercari_match.group(1)
            netmall_match = re.search(r'netmall\.hardoff\.co\.jp/product/(\d+)', link)
            if netmall_match:
                return 'netmall', netmall_match.group(1)
            rakuten_match = re.search(r'item\.rakuten\.co\.jp/([^/]+)/([a-zA-Z0-9_-]+)', link)
            if rakuten_match:
                return 'rakuten', rakuten_match.group(1) + '/' + rakuten_match.group(2)
            yahoo_shopping_match = re.search(r'store\.shopping\.yahoo\.co\.jp/([^/]+)/([a-zA-Z0-9_-]+)\.html', link)
            if yahoo_shopping_match:
                return 'yahoo_shopping', yahoo_shopping_match.group(1) + '/' + yahoo_shopping_match.group(2)
            return None, None

        def fetch_paypay_item(item_id):
            url = 'https://paypayfleamarket.yahoo.co.jp/item/' + item_id
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml',
                'Accept-Language': 'ja-JP,ja;q=0.9,en-US;q=0.8,en;q=0.7',
            }
            resp = http_requests.get(url, headers=headers, timeout=15)
            resp.raise_for_status()
            html = resp.text
            nd_match = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', html, re.DOTALL)
            if not nd_match:
                title_match2 = re.search(r'<title[^>]*>(.*?)</title>', html)
                title = title_match2.group(1) if title_match2 else ''
                return {'title': title, 'status': '', 'price': 0}
            raw = nd_match.group(1)
            raw_clean = re.sub(r'[\x00-\x1f]', '', raw)
            try:
                import json
                parsed = json.loads(raw_clean)
                init = parsed.get('props', {}).get('initialState', {})
                item_data = init.get('itemsState', {}).get('items', {}).get('item', {})
                status = item_data.get('status', '')
                title = item_data.get('title', '')
                price = item_data.get('price', 0)
                return {'title': title, 'status': status, 'price': price, 'is_purchased': False}
            except Exception:
                pass
            status_m = re.search(r'"status"\s*:\s*"([^"]+)"', raw)
            title_m = re.search(r'<title[^>]*>(.*?)</title>', html)
            price_m = re.search(r'"price"\s*:\s*(\d+)', raw)
            return {
                'title': title_m.group(1) if title_m else '',
                'status': status_m.group(1) if status_m else '',
                'price': int(price_m.group(1)) if price_m else 0,
                'is_purchased': False
            }

        def fetch_netmall_item(item_id):
            url = 'https://netmall.hardoff.co.jp/product/' + item_id + '/'
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml',
                'Accept-Language': 'ja-JP,ja;q=0.9,en-US;q=0.8,en;q=0.7',
            }
            try:
                resp = http_requests.get(url, headers=headers, timeout=15)
            except http_requests.exceptions.HTTPError as e:
                if e.response is not None and e.response.status_code == 404:
                    return {'title': '', 'availability': 'sold_out', 'price': 0, 'status_code': 404}
                raise
            if resp.status_code == 404:
                return {'title': '', 'availability': 'sold_out', 'price': 0, 'status_code': 404}
            html = resp.text
            ld_matches = re.findall(r'<script type="application/ld\+json"[^>]*>(.*?)</script>', html, re.DOTALL)
            availability = 'unknown'
            price = 0
            for raw in ld_matches:
                try:
                    import json
                    parsed = json.loads(raw)
                    if parsed.get('@type') == 'Product':
                        offers = parsed.get('offers', {})
                        availability = offers.get('availability', '')
                        price_str = offers.get('price', '0')
                        price = int(float(price_str)) if price_str else 0
                except Exception:
                    pass
            title_m = re.search(r'<title[^>]*>(.*?)</title>', html)
            raw_title = title_m.group(1) if title_m else ''
            title = re.sub(r'\|【ハードオフ公式通販】オフモール\|.*$', '', raw_title)
            return {'title': title, 'availability': availability, 'price': price}

        def fetch_rakuten_item(item_id):
            shop_item = item_id
            url = 'https://item.rakuten.co.jp/' + shop_item + '/'
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml',
                'Accept-Language': 'ja-JP,ja;q=0.9,en-US;q=0.8,en;q=0.7',
            }
            resp = http_requests.get(url, headers=headers, timeout=15)
            resp.encoding = 'EUC-JP'
            html = resp.text
            if len(html) < 100 and 'Reference' in html:
                return {'title': '', 'availability': 'unknown', 'price': 0, 'status_code': 403, 'error': 'Rakuten blocked the request (Akamai)'}
            title_m = re.search(r'<title[^>]*>(.*?)</title>', html)
            title = ''
            if title_m:
                raw = title_m.group(1)
                title = re.sub(r'^【楽天市場】', '', raw)
                title = re.sub(r'｜.*$', '', title)
            has_soldout_popup = 'irc="SoldOutPopup"' in html
            availability_m = re.search(r'itemprop="availability"\s+content="([^"]+)"', html)
            availability = ''
            if availability_m:
                availability = availability_m.group(1)
            is_available = (not has_soldout_popup) and ('InStock' in availability or 'instock' in availability.lower())
            price_m = re.search(r'itemprop="price"\s+content="(\d+)"', html)
            price = int(price_m.group(1)) if price_m else 0
            return {'title': title, 'availability': availability, 'price': price, 'soldout_popup': has_soldout_popup}

        def fetch_mercari_shop_item(item_id):
            try:
                info = getShopProductInfo(item_id)
                return {'title': info.name, 'price': info.price, 'status': info.status}
            except Exception as e:
                print(f'[shop api error] {e}', file=sys.stderr)
                return {'title': '', 'price': 0, 'status': 'on_sale'}

        def fetch_yahoo_shopping_item(item_id):
            url = 'https://store.shopping.yahoo.co.jp/' + item_id + '.html'
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml',
                'Accept-Language': 'ja-JP,ja;q=0.9,en-US;q=0.8,en;q=0.7',
            }
            try:
                resp = http_requests.get(url, headers=headers, timeout=15)
            except http_requests.exceptions.HTTPError as e:
                if e.response is not None and e.response.status_code == 404:
                    return {'title': '', 'price': 0, 'available': False}
                raise
            if resp.status_code == 404:
                return {'title': '', 'price': 0, 'available': False}
            html = resp.text
            nd_match = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', html, re.DOTALL)
            if nd_match:
                try:
                    nd = json.loads(nd_match.group(1))
                    page_props = nd.get('props', {}).get('pageProps', {})
                    item = page_props.get('item', {})
                    stock = item.get('stock', {})
                    is_available = stock.get('isAvailable', False)
                    title = item.get('name', '')
                    price = item.get('applicablePrice', 0)
                    return {'title': title, 'price': price, 'available': is_available}
                except Exception:
                    pass
            ld_matches = re.findall(r'<script type="application/ld\+json"[^>]*>(.*?)</script>', html, re.DOTALL)
            for raw in ld_matches:
                try:
                    parsed = json.loads(raw)
                    if parsed.get('@type') == 'Product':
                        offers = parsed.get('offers', {})
                        avail = offers.get('availability', '')
                        is_available = 'InStock' in avail
                        title_m = re.search(r'<title[^>]*>(.*?)</title>', html)
                        title = re.sub(r'\s*-\s*通販\s*-\s*Yahoo!ショッピング\s*$', '', title_m.group(1)) if title_m else ''
                        price_str = offers.get('price', '0')
                        price = int(float(price_str)) if price_str else 0
                        return {'title': title, 'price': price, 'available': is_available}
                except Exception:
                    pass
            title_m = re.search(r'<title[^>]*>(.*?)</title>', html)
            title = re.sub(r'\s*-\s*通販\s*-\s*Yahoo!ショッピング\s*$', '', title_m.group(1)) if title_m else ''
            has_sold = '売り切れ' in html or '在庫なし' in html or 'sold out' in html.lower()
            return {'title': title, 'price': 0, 'available': not has_sold}

        def check_link(link, keyword=''):
            base_result = {'link': link, 'keyword': keyword}
            platform, item_id = parse_link(link)
            if not platform:
                return {**base_result, 'platform': 'unknown', 'error': 'Could not parse URL'}

            base_result['platform'] = platform
            base_result['id'] = item_id

            try:
                if platform == 'paypay':
                    detail = fetch_paypay_item(item_id)
                    status = detail.get('status', '') or ''
                    is_purchased = detail.get('is_purchased', False)
                    is_available = status.upper() == 'OPEN' and not is_purchased
                    return {**base_result,
                        'name': detail.get('title', ''),
                        'price': detail.get('price', 0),
                        'status': status,
                        'available': is_available,
                        'available_text': 'In Stock' if is_available else 'Out of Stock',
                        'url': link
                    }
                elif platform == 'mercari':
                    item = getItemInfo(item_id)
                    status = getattr(item, 'status', '') or ''
                    name = getattr(item, 'name', '')
                    price = getattr(item, 'price', 0)
                    sold_statuses = ('item_status_sold_out', 'sold_out', 'item_status_stop', 'stop', 'item_status_cancel', 'cancel', 'item_status_admin_cancel')
                    is_available = status.lower() not in sold_statuses
                    return {**base_result,
                        'name': name,
                        'price': price,
                        'status': status,
                        'available': is_available,
                        'available_text': 'In Stock' if is_available else 'Out of Stock',
                        'url': link
                    }
                elif platform == 'mercari_shop':
                    detail = fetch_mercari_shop_item(item_id)
                    is_available = detail.get('status') == 'on_sale'
                    return {**base_result,
                        'name': detail.get('title', ''),
                        'price': detail.get('price', 0),
                        'status': detail.get('status', ''),
                        'available': is_available,
                        'available_text': 'In Stock' if is_available else 'Out of Stock',
                        'url': link
                    }
                elif platform == 'yahoo':
                    detail = get_item_detail(item_id)
                    status = detail.get('status', '') or ''
                    is_available = status.lower() not in ('sold', 'closed', 'ended')
                    return {**base_result,
                        'name': detail.get('title', ''),
                        'price': detail.get('price', 0),
                        'status': status,
                        'available': is_available,
                        'available_text': 'In Stock' if is_available else 'Out of Stock',
                        'url': link
                    }
                elif platform == 'netmall':
                    detail = fetch_netmall_item(item_id)
                    status_code = detail.get('status_code', 200)
                    availability = detail.get('availability', '')
                    is_sold_page = status_code == 404
                    is_out_of_stock = availability in ('https://schema.org/OutOfStock', 'https://schema.org/Discontinued', 'SoldOut')
                    is_available = not is_sold_page and not is_out_of_stock
                    return {**base_result,
                        'name': detail.get('title', ''),
                        'price': detail.get('price', 0),
                        'status': availability,
                        'available': is_available,
                        'available_text': 'In Stock' if is_available else 'Out of Stock',
                        'url': link
                    }
                elif platform == 'rakuten':
                    detail = fetch_rakuten_item(item_id)
                    if detail.get('error'):
                        return {**base_result, 'error': detail['error']}
                    is_available = detail.get('availability') == 'http://schema.org/InStock' and not detail.get('soldout_popup', False)
                    return {**base_result,
                        'name': detail.get('title', ''),
                        'price': detail.get('price', 0),
                        'status': detail.get('availability', ''),
                        'available': is_available,
                        'available_text': 'In Stock' if is_available else 'Out of Stock',
                        'url': link
                    }
                elif platform == 'yahoo_shopping':
                    detail = fetch_yahoo_shopping_item(item_id)
                    is_available = detail.get('available', False)
                    return {**base_result,
                        'name': detail.get('title', ''),
                        'price': detail.get('price', 0),
                        'status': 'on_sale' if is_available else 'sold_out',
                        'available': is_available,
                        'available_text': 'In Stock' if is_available else 'Out of Stock',
                        'url': link
                    }
            except Exception as e:
                return {**base_result, 'error': str(e)}

        def generate():
            from concurrent.futures import as_completed, TimeoutError as _TimeoutError
            with ThreadPoolExecutor(max_workers=3) as pool:
                fut_to_key = {}
                for idx, link in enumerate(links):
                    kw = keywords[idx] if idx < len(keywords) else ''
                    fut = pool.submit(check_link, link, kw)
                    fut_to_key[fut] = kw
                try:
                    for f in as_completed(fut_to_key, timeout=300):
                        try:
                            result = f.result(timeout=5)
                        except Exception as exc:
                            result = {'link': '', 'keyword': fut_to_key.get(f, ''), 'error': str(exc)}
                        yield 'data: ' + json.dumps(result, ensure_ascii=False) + '\n\n'
                except _TimeoutError:
                    for f in list(fut_to_key):
                        if not f.done():
                            f.cancel()
            yield 'data: ' + json.dumps({'complete': True, 'total': len(links)}, ensure_ascii=False) + '\n\n'

        return Response(stream_with_context(generate()), mimetype='text/event-stream')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/ebay-search')
def api_ebay_search():
    keyword = request.args.get('keyword', '').strip()
    if not keyword:
        return jsonify({'error': 'Keyword is required'}), 400

    app_id = request.args.get('app_id', '').strip()
    cert_id = request.args.get('cert_id', '').strip()
    if not app_id:
        app_id = os.environ.get('EBAY_APP_ID', '')
    if not cert_id:
        cert_id = os.environ.get('EBAY_CERT_ID', '')
    if not app_id:
        return jsonify({'error': 'eBay App ID is required.'}), 400
    if not cert_id:
        return jsonify({'error': 'eBay Cert ID is required.'}), 400

    min_price = request.args.get('min_price', '').strip()
    max_price = request.args.get('max_price', '').strip()
    condition = request.args.get('condition', '').strip()
    sort = request.args.get('sort', 'best_match').strip()
    global_id = request.args.get('global_id', 'EBAY-US').strip()
    item_location = request.args.get('item_location', '').strip()
    bin_only = request.args.get('bin_only', '') == '1'
    limit = request.args.get('limit', '20').strip()

    try:
        limit = int(limit) if limit else 20
    except ValueError:
        limit = 20
    limit = max(1, min(limit, 100))

    min_price_val = None
    max_price_val = None
    try:
        min_price_val = float(min_price) if min_price else None
    except ValueError:
        pass
    try:
        max_price_val = float(max_price) if max_price else None
    except ValueError:
        pass

    try:
        results = search_ebay(
            app_id=app_id,
            cert_id=cert_id,
            keyword=keyword,
            min_price=min_price_val,
            max_price=max_price_val,
            condition=condition,
            sort=sort,
            limit=limit,
            global_id=global_id,
            bin_only=bin_only,
            item_location=item_location,
        )

        items = []
        for item in results:
            items.append({
                'id': item.id,
                'title': item.title,
                'price': item.price,
                'currency': item.currency,
                'condition': item.condition,
                'condition_id': item.condition_id,
                'thumbnail': item.thumbnail,
                'url': item.url,
                'listing_type': item.listing_type,
                'bid_count': item.bid_count,
                'buy_it_now_price': item.buy_it_now_price,
                'seller_name': item.seller_name,
                'seller_rating': item.seller_rating,
                'seller_feedback_percent': item.seller_feedback_percent,
                'location': item.location,
                'shipping_cost': item.shipping_cost,
                'shipping_type': item.shipping_type,
                'free_shipping': item.free_shipping,
                'end_time': item.end_time,
                'is_auction': item.is_auction,
                'top_rated_seller': item.top_rated_seller,
                'category_name': item.category_name,
            })

        return jsonify({'items': items, 'count': len(items)})

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/ebay-bulk-search', methods=['POST'])
def api_ebay_bulk_search():
    try:
        data = request.get_json()
        if not data or 'keywords' not in data:
            return jsonify({'error': 'Keywords required'}), 400
        keywords = data['keywords']
        if not isinstance(keywords, list) or not keywords:
            return jsonify({'error': 'Keywords must be a non-empty array'}), 400

        app_id = data.get('app_id', '').strip()
        cert_id = data.get('cert_id', '').strip()
        if not app_id:
            app_id = os.environ.get('EBAY_APP_ID', '')
        if not cert_id:
            cert_id = os.environ.get('EBAY_CERT_ID', '')
        if not app_id:
            return jsonify({'error': 'eBay App ID is required'}), 400
        if not cert_id:
            return jsonify({'error': 'eBay Cert ID is required'}), 400

        per_keyword = int(data.get('per_keyword', 3))
        per_keyword = max(1, min(per_keyword, 20))
        condition = data.get('condition', '')
        sort = data.get('sort', 'best_match')
        global_id = data.get('global_id', 'EBAY-US')
        min_price = data.get('min_price', '')
        max_price = data.get('max_price', '')
        bin_only = data.get('bin_only', '') in ('1', 'true', True)
        item_location = data.get('item_location', '')

        min_price_val = float(min_price) if min_price else None
        max_price_val = float(max_price) if max_price else None

        def search_keyword(kw):
            try:
                results = search_ebay(
                    app_id=app_id,
                    cert_id=cert_id,
                    keyword=kw,
                    min_price=min_price_val,
                    max_price=max_price_val,
                    condition=condition,
                    sort=sort,
                    limit=per_keyword,
                    global_id=global_id,
                    bin_only=bin_only,
                    item_location=item_location,
                )
                items = []
                for item in results[:per_keyword]:
                    items.append({
                        'id': item.id,
                        'title': item.title,
                        'price': item.price,
                        'currency': item.currency,
                        'condition': item.condition,
                        'thumbnail': item.thumbnail,
                        'url': item.url,
                        'listing_type': item.listing_type,
                        'bid_count': item.bid_count,
                        'buy_it_now_price': item.buy_it_now_price,
                        'seller_name': item.seller_name,
                        'seller_rating': item.seller_rating,
                        'location': item.location,
                        'shipping_cost': item.shipping_cost,
                        'shipping_type': item.shipping_type,
                        'free_shipping': item.free_shipping,
                        'is_auction': item.is_auction,
                        'top_rated_seller': item.top_rated_seller,
                    })
                return kw, items
            except Exception:
                return kw, []

        keywords_list = [k.strip() for k in keywords[:100] if k.strip()]

        def generate():
            with ThreadPoolExecutor(max_workers=5) as pool:
                futures = {pool.submit(search_keyword, kw): kw for kw in keywords_list}
                for f in as_completed(futures):
                    kw, items = f.result()
                    yield 'data: ' + json.dumps({'keyword': kw, 'items': items}, ensure_ascii=False) + '\n\n'
            yield 'data: ' + json.dumps({'complete': True, 'keywords': keywords_list}, ensure_ascii=False) + '\n\n'

        return Response(stream_with_context(generate()), mimetype='text/event-stream')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    print('='*60)
    print(' Mercari JP Search + Yahoo Auctions + eBay GUI')
    print('='*60)
    print(' Open http://127.0.0.1:5000 in your browser')
    print(' Press CTRL+C to stop the server')
    print('='*60)
    app.run(debug=True, host='127.0.0.1', port=5000)
