import React, { useState, useEffect, useMemo, useRef } from 'react';
import { Search, FileJson, AlertCircle, ChevronLeft, ChevronRight, Database, FileText, Info, CheckCircle2, Download, Sparkles, X, Loader2, KeyRound, Eye, EyeOff, RotateCcw } from 'lucide-react';

// ─── Constants ───────────────────────────────────────────────────────────────

const DATA_URL = "https://g100timess.github.io/iea-clean-tech-zh-tw_new/rebuilt_dataset_zh_en.json";

// The AI key is supplied by whoever is using the app (stored in their own
// browser's localStorage, one slot per provider) rather than baked into the
// bundle, so no single key gets shipped to every visitor or billed on the
// developer's account. Switching providers keeps each provider's previously
// entered key intact.
const AI_PROVIDER_STORAGE_KEY = "iea_dashboard_ai_provider";
const AI_API_KEY_STORAGE_PREFIX = "iea_dashboard_api_key_";
const OPENROUTER_MODEL_STORAGE_KEY = "iea_dashboard_openrouter_model";

// Each provider's default model and how to reach it. OpenRouter's whole
// purpose is picking an arbitrary underlying model, so it gets a free-text
// model field in the UI; the other three use a fixed default and aren't
// user-editable (same as this app's previous single-provider design).
const AI_PROVIDERS = {
  gemini: { label: "Google Gemini", defaultModel: "gemini-2.5-flash" },
  openai: { label: "OpenAI", defaultModel: "gpt-4o-mini" },
  claude: { label: "Anthropic Claude", defaultModel: "claude-haiku-4-5-20251001" },
  openrouter: { label: "OpenRouter", defaultModel: "openai/gpt-4o-mini" }
};
const AI_PROVIDER_ORDER = ["gemini", "openai", "claude", "openrouter"];

const DEFAULT_JSON = {
  dataset_version: "IEA ETP Clean Energy Technology Guide",
  generated_at: "2026-05-21T12:10:09.402Z",
  source_stats: { technology_rows: 5, initiative_rows: 3 },
  technologies: {
    "tech_h2_1": {
      technology_name: "Hydrogen electrolysers",
      technology_name_zh: "氫能電解槽",
      sector_zh: "能源轉換",
      latest_trl: "9",
      technology_status_summary_zh: "利用電力將水分解為氫氣與氧氣的技術，為淨零排放提供無碳氫能。電解槽是整合再生能源與難脫碳產業的核心節點。",
      market_dynamics_summary_zh: "全球產能快速擴張中，政策支持與企業投資皆達歷史新高，重點在於降低設備成本與提升效率。",
      review_status: "reviewed",
      linked_records_count: 45,
      initiatives: [
        { country: "United States", year: "2024", type: "Policy", description: "National Clean Hydrogen Strategy", read_more: "https://example.com" }
      ]
    },
    "tech_h2_2": {
      technology_name: "Hydrogen storage in salt caverns",
      technology_name_zh: "鹽穴氫能儲存",
      sector_zh: "基礎設施",
      latest_trl: "8",
      technology_status_summary_zh: "利用地下鹽穴進行大規模、高壓的氫氣儲存。解決再生能源季節性波動，提供長期穩定的氫氣供應。",
      market_dynamics_summary_zh: "",
      review_status: "reviewed",
      linked_records_count: 12,
      initiatives: []
    }
  }
};

const FIXED_SECTORS = {
  "Industry": ["Aluminium", "Metallic products", "Cement and concrete", "Chemicals and plastics", "Industrial heating", "Iron and steel", "Pulp and paper"],
  "Buildings": ["Cooking technologies", "Operations and equipment", "Design and envelope"],
  "Transport": ["Rail", "Aviation", "Road transport", "Shipping"],
  "Fossil fuels": ["Coal", "Oil and gas"],
  "Renewables": ["Hydropower", "Ocean", "Wind", "Bioenergy", "Geothermal", "Solar"],
  "Nuclear": ["Fission", "Fusion"],
  "Hydrogen": ["Sythetic fuels production", "H2 infrastructure", "Power generation", "Production"],
  "Energy networks and storage": ["Thermal storage for buildings and district heating", "Mechanical storage", "Electrochemical storage", "Thermal storage", "Physical grids", "Smart grids"],
  "Carbon Capture and Storage": ["CO2 capture", "CO2 storage", "CO2 transport"],
  "Critical minerals": ["Mineral processing and refining", "Mining and extraction", "E-waste recycling"]
};

const SECTOR_TRANSLATIONS = {
  "Industry": "工業", "Aluminium": "鋁", "Metallic products": "金屬產品",
  "Cement and concrete": "水泥與混凝土", "Chemicals and plastics": "化學品與塑膠",
  "Industrial heating": "工業供熱", "Iron and steel": "鋼鐵", "Pulp and paper": "紙漿與造紙",
  "Buildings": "建築", "Cooking technologies": "烹飪技術", "Operations and equipment": "營運與設備",
  "Design and envelope": "設計與外殼", "Transport": "運輸", "Rail": "鐵路", "Aviation": "航空",
  "Road transport": "公路運輸", "Shipping": "航運", "Fossil fuels": "化石燃料", "Coal": "煤炭",
  "Oil and gas": "石油與天然氣", "Renewables": "再生能源", "Hydropower": "水力發電",
  "Ocean": "海洋能", "Wind": "風力", "Bioenergy": "生質能", "Geothermal": "地熱",
  "Solar": "太陽能", "Nuclear": "核能", "Fission": "核分裂", "Fusion": "核融合",
  "Hydrogen": "氫能", "Sythetic fuels production": "合成燃料生產", "H2 infrastructure": "氫能基礎設施",
  "Power generation": "發電", "Production": "生產", "Energy networks and storage": "能源網路與儲能",
  "Thermal storage for buildings and district heating": "建築與區域供熱儲能",
  "Mechanical storage": "機械儲能", "Electrochemical storage": "電化學儲能",
  "Thermal storage": "熱儲存", "Physical grids": "實體電網", "Smart grids": "智慧電網",
  "Carbon Capture and Storage": "碳捕捉與封存", "CO2 capture": "二氧化碳捕捉",
  "CO2 storage": "二氧化碳封存", "CO2 transport": "二氧化碳運輸",
  "Critical minerals": "關鍵礦物", "Mineral processing and refining": "礦物加工與精煉",
  "Mining and extraction": "採礦與開採", "E-waste recycling": "電子廢棄物回收"
};

const SEARCH_ALIASES = {
  '氫能': ['氫能', 'hydrogen', 'h2'], 'hydrogen': ['氫能', 'hydrogen', 'h2'], 'h2': ['氫能', 'hydrogen', 'h2'],
  '太陽能': ['太陽能', '太陽光電', 'solar', 'solar pv', 'photovoltaic', 'pv'],
  '太陽光電': ['太陽能', '太陽光電', 'solar', 'solar pv', 'photovoltaic', 'pv'],
  'solar': ['太陽能', '太陽光電', 'solar', 'solar pv', 'photovoltaic', 'pv'],
  '地熱': ['地熱', 'geothermal'], 'geothermal': ['地熱', 'geothermal'],
  '碳捕捉': ['碳捕捉', 'ccs', 'ccus'], 'ccs': ['碳捕捉', 'ccs', 'ccus'], 'ccus': ['碳捕捉', 'ccs', 'ccus'],
};

const TRL_STAGE_LABELS = ['原型研發 TRL 1-3', '技術示範 TRL 4-6', '商業初期 TRL 7-8', '成熟推廣 TRL 9+'];
const TRL_STAGE_LABELS_EN = ['Prototype TRL 1-3', 'Demonstration TRL 4-6', 'Early Commercial TRL 7-8', 'Mature/Deployed TRL 9+'];

// computeTechStats always keys trlBreakdown by the zh stage names above (used
// internally as a stable index); this maps to the English label for display
// when the platform is in English mode, without changing computeTechStats's
// shared contract.
function trlStageLabel(stage, uiLang) {
  if (uiLang !== 'en') return stage;
  const idx = TRL_STAGE_LABELS.indexOf(stage);
  return idx >= 0 ? TRL_STAGE_LABELS_EN[idx] : stage;
}

// In-slide microcopy (hints, prefixes, empty states) that should follow the
// same language/wording the user asks Gemini for, same as tech names. These
// are defaults used when geminiData.ui_labels is missing a key (e.g. the
// rule-based fallback, or a partial AI response) — NOT the surrounding modal
// chrome (pagination, download button, composer), which stays fixed since
// it's the tool's own interface, not part of the slide content.
const DEFAULT_UI_LABELS = {
  zh: {
    trlBarHint: '點擊長條可查看該區間的技術清單',
    trlListHeading: '技術清單',
    countUnit: '項',
    trlListEmpty: '此區間沒有符合的技術',
    trlTechHint: '點擊技術名稱可查看該技術的 TRL 年度趨勢',
    collapseButton: '收起',
    relatedPrefix: '依據',
    caseCountSuffix: '筆案例',
    trendChartSuffix: 'TRL 年度趨勢',
    trendInsufficientData: '此技術缺乏足夠的年度 TRL 資料，無法繪製趨勢'
  },
  en: {
    trlBarHint: 'Click a bar to see the technologies in that range',
    trlListHeading: 'Technology list',
    countUnit: '',
    trlListEmpty: 'No matching technologies in this range',
    trlTechHint: 'Click a technology name to see its TRL trend over time',
    collapseButton: 'Collapse',
    relatedPrefix: 'Based on',
    caseCountSuffix: 'cases',
    trendChartSuffix: 'TRL trend',
    trendInsufficientData: 'Not enough yearly TRL data to plot a trend for this technology'
  }
};

// The AI modal's own operating chrome (nav buttons, composer, loading states)
// — unlike DEFAULT_UI_LABELS above, this is NOT AI-controlled. It's the
// tool's own interface, so it gets a plain manual zh/en toggle instead of
// asking Gemini to translate it.
const MODAL_CHROME_LABELS = {
  zh: {
    title: 'AI 策略簡報預覽',
    subtitle: (n) => `共分析 ${n} 筆技術資料・點擊帶底線的技術名稱可查看詳情`,
    download: '下載 PPTX',
    downloading: '生成中...',
    basicBadge: '基本統計',
    aiBadge: 'AI 洞察',
    upgradeToAI: '升級為 AI 洞察',
    upgrading: 'AI 生成中...',
    needApiKeyForUpgrade: '請先在右上角選擇 AI 供應商並輸入對應的 API 金鑰，才能升級為 AI 洞察。',
    prevSlide: '上一頁',
    nextSlide: '下一頁',
    prevVersion: '上一版',
    nextVersion: '下一版',
    versionLabel: (i, n) => `版本 ${i} / ${n}`,
    followUpPlaceholder: '追問調整這份簡報，例如：請更聚焦氫能 / 幫我用英文重寫',
    revising: '調整中...',
    send: '送出',
    noPreview: '無法生成預覽，請重試'
  },
  en: {
    title: 'AI Strategy Slide Preview',
    subtitle: (n) => `Analyzed ${n} technologies · click underlined tech names for details`,
    download: 'Download PPTX',
    downloading: 'Generating...',
    basicBadge: 'Basic Stats',
    aiBadge: 'AI Insights',
    upgradeToAI: 'Upgrade to AI Insights',
    upgrading: 'Generating with AI...',
    needApiKeyForUpgrade: 'Choose an AI provider and enter its API key in the top-right corner to upgrade to AI insights.',
    prevSlide: 'Previous',
    nextSlide: 'Next',
    prevVersion: 'Previous version',
    nextVersion: 'Next version',
    versionLabel: (i, n) => `Version ${i} / ${n}`,
    followUpPlaceholder: 'Ask for a change, e.g. "Focus more on hydrogen" / "Rewrite in English"',
    revising: 'Revising...',
    send: 'Send',
    noPreview: 'Could not generate a preview, please try again'
  }
};

// Platform-wide UI chrome (header, search panel, tabs, footer). Same zh/en
// toggle idea as MODAL_CHROME_LABELS above, driven by the same global uiLang
// state — this is the app's own interface, not the underlying dataset.
// IMPORTANT: this does NOT translate cached data fields (technology_name_zh,
// summaries, descriptions) — those are static pre-computed Chinese text in
// the loaded JSON, not something a UI toggle can transform without calling
// an LLM per field. Only the labels/buttons/hints around them switch.
const APP_LABELS = {
  zh: {
    appTitle: 'IEA 淨零技術智慧分析平台',
    datasetVersion: (v) => `資料版本: ${v}`,
    sampleData: '測試資料',
    fallbackBadge: '目前使用測試資料',
    loadingBadge: '正在讀取 JSON 資料...',
    successBadge: (t, i) => `完整資料已載入：${t} 項技術、${i} 筆案例`,
    aiProviderLabel: 'AI 供應商',
    apiKeyPlaceholder: (p) => `貼上你的 ${p} API 金鑰`,
    openRouterModelPlaceholder: '模型 ID，例如 openai/gpt-4o-mini',
    hideKey: '隱藏金鑰',
    showKey: '顯示金鑰',
    uploadJson: '選擇 JSON 檔案',
    rawFormatDetected: '偵測到原始英文技術資料（無中文摘要），已自動切換為英文顯示',
    restoreDefaultData: '還原為預設資料',
    subjectSearch: '技術主題查詢',
    fulltextSearch: 'IEA 全文查詢',
    allSectors: '所有官方分類 (All Sectors)',
    searchPlaceholder: '輸入關鍵字 (例: hydrogen)...',
    searchHint: '技術主題查詢以技術名稱與分類為主；IEA 全文查詢另納入描述中提及的應用、能源來源與供應鏈關聯。',
    resultsCount: (n, page, totalPages) => `共 ${n} 筆｜第 ${page} / ${totalPages} 頁`,
    noResultsShort: '無結果',
    exportAll: '匯出所有完整資料',
    exportSingle: '匯出單一資料',
    exportStrategyTitle: '匯出完整查詢結果',
    exportStrategyDescPrefix: '將目前符合條件的全部',
    exportStrategyDescSuffix: '筆技術資料匯出為 CSV 檔案。',
    emptyPrompt: '請輸入關鍵字或選擇官方分類開始查詢',
    noMatch: '找不到符合條件的技術資料',
    directHit: '直接相關',
    descHit: '描述提及',
    unlabeled: '未標示',
    caseCountInline: (n) => `案例: ${n}`,
    prevPage: '上一頁',
    nextPage: '下一頁',
    detailTab: '技術詳情',
    strategyTab: '查詢策略總覽',
    selectPrompt: '請從左側列表選取一項技術以檢視詳情',
    fieldSector: (s) => `領域: ${s}`,
    fieldTrl: (v) => `TRL: ${v}`,
    fieldCaseCount: (n) => `相關案例: ${n} 筆`,
    summaryTitle: '中文摘要',
    marketTitle: '市場與示範動態',
    aiDisclaimer: '提醒：中文摘要由 AI 輔助生成，正式引用前請核對原始 IEA 資料。',
    noSummary: '快取資料未提供技術中文摘要',
    noMarket: '快取資料未提供市場動態中文摘要',
    initiativesTitle: '相關案例',
    totalCount: (n) => `共 ${n} 筆`,
    noInitiatives: '快取表未連結案例',
    sourceLink: '來源連結',
    collapseCases: '收合案例',
    expandCases: (n) => `展開全部 ${n} 筆`,
    strategyTitle: '查詢策略總覽',
    strategySubtitle: '根據目前查詢結果產生的整體分析，與左側清單、單一技術詳情是平行的檢視角度，不會隨你點選哪一項技術而改變。',
    strategyCardTitle: '策略總覽與 AI 洞察',
    strategyCardDescPrefix: '立即查看目前符合條件的全部',
    strategyCardDescSuffix: '筆技術的整體分析，之後可選擇升級為 AI 深度洞察。',
    openStrategyBtn: '開啟策略總覽',
    footerDisclaimer: '免責文字：根據 IEA ETP Clean Energy Technology Guide 快取資料生成，正式引用前請核對原始 IEA 資料。',
    loadingFromGithub: '正在載入資料...',
    githubLoadSuccess: '資料載入成功',
    remoteFormatError: '遠端檔案格式錯誤：找不到有效的技術陣列資料。',
    networkLoadFailed: (msg) => `網路載入失敗: ${msg}`,
    autoLoadFailed: '自動載入失敗',
    fileLoadFailed: '檔案載入失敗：解析後找不到有效的技術資料。',
    jsonParseFailed: (msg) => `JSON 解析失敗: ${msg}`,
    localFileReadError: '讀取本機檔案時發生系統錯誤。',
    backToAiSlides: '返回 AI 簡報預覽'
  },
  en: {
    appTitle: 'IEA Net-Zero Technology Intelligence Platform',
    datasetVersion: (v) => `Dataset version: ${v}`,
    sampleData: 'Sample data',
    fallbackBadge: 'Currently showing sample data',
    loadingBadge: 'Loading JSON data...',
    successBadge: (t, i) => `Data loaded: ${t} technologies, ${i} case records`,
    aiProviderLabel: 'AI Provider',
    apiKeyPlaceholder: (p) => `Paste your ${p} API key`,
    openRouterModelPlaceholder: 'Model ID, e.g. openai/gpt-4o-mini',
    hideKey: 'Hide key',
    showKey: 'Show key',
    uploadJson: 'Choose JSON file',
    rawFormatDetected: 'Detected the raw English source (no Chinese summaries) — switched to English display automatically',
    restoreDefaultData: 'Restore default data',
    subjectSearch: 'Subject search',
    fulltextSearch: 'IEA full-text search',
    allSectors: 'All Sectors',
    searchPlaceholder: 'Enter a keyword (e.g. hydrogen)...',
    searchHint: 'Subject search matches technology names and categories; IEA full-text search also matches applications, energy sources and supply chains mentioned in the descriptions.',
    resultsCount: (n, page, totalPages) => `${n} results | page ${page} / ${totalPages}`,
    noResultsShort: 'No results',
    exportAll: 'Export all data',
    exportSingle: 'Export this item',
    exportStrategyTitle: 'Export Full Query Results',
    exportStrategyDescPrefix: 'Export all',
    exportStrategyDescSuffix: 'matching technologies as a CSV file.',
    emptyPrompt: 'Enter a keyword or choose an official category to start searching',
    noMatch: 'No matching technologies found',
    directHit: 'Direct match',
    descHit: 'Mentioned in description',
    unlabeled: 'Unlabeled',
    caseCountInline: (n) => `Cases: ${n}`,
    prevPage: 'Previous',
    nextPage: 'Next',
    detailTab: 'Technology Detail',
    strategyTab: 'Query Strategy Overview',
    selectPrompt: 'Select a technology from the list on the left to see its details',
    fieldSector: (s) => `Sector: ${s}`,
    fieldTrl: (v) => `TRL: ${v}`,
    fieldCaseCount: (n) => `Related cases: ${n}`,
    summaryTitle: 'Summary',
    marketTitle: 'Market & Deployment Dynamics',
    aiDisclaimer: 'Note: this summary was AI-assisted and cached from the source data; verify against the original IEA data before formal citation.',
    noSummary: 'No cached summary available for this technology',
    noMarket: 'No cached market dynamics summary available',
    initiativesTitle: 'Related Cases',
    totalCount: (n) => `${n} total`,
    noInitiatives: 'No linked cases in the cache',
    sourceLink: 'Source link',
    collapseCases: 'Collapse',
    expandCases: (n) => `Show all ${n}`,
    strategyTitle: 'Query Strategy Overview',
    strategySubtitle: 'An aggregate analysis of the current query results — a view parallel to the list and the single-technology detail, unaffected by whichever technology you select.',
    strategyCardTitle: 'Strategy Overview & AI Insight',
    strategyCardDescPrefix: 'Instantly view an overview of all',
    strategyCardDescSuffix: 'matching technologies, with the option to upgrade to AI-generated insight afterwards.',
    openStrategyBtn: 'Open Strategy Overview',
    footerDisclaimer: 'Disclaimer: generated from cached IEA ETP Clean Energy Technology Guide data — verify against the original IEA source before formal citation.',
    loadingFromGithub: 'Loading data...',
    githubLoadSuccess: 'Data loaded successfully',
    remoteFormatError: 'Remote file format error: no valid technology array found.',
    networkLoadFailed: (msg) => `Network load failed: ${msg}`,
    autoLoadFailed: 'Auto-load failed',
    fileLoadFailed: 'File load failed: no valid technology data found after parsing.',
    jsonParseFailed: (msg) => `JSON parse failed: ${msg}`,
    localFileReadError: 'A system error occurred while reading the local file.',
    backToAiSlides: 'Back to AI slide preview'
  }
};

// ─── Shared helpers ──────────────────────────────────────────────────────────

function countOccurrences(text, q) {
  if (!text || !q) return 0;
  if (/^[a-z0-9]+$/i.test(q)) {
    try {
      const regex = new RegExp(`\\b${q}\\b`, 'gi');
      return (text.match(regex) || []).length;
    } catch (e) { return 0; }
  }
  let count = 0, pos = text.indexOf(q);
  while (pos !== -1) { count++; pos = text.indexOf(q, pos + q.length); }
  return count;
}

function escapeCsv(val) {
  if (val === null || val === undefined) return '""';
  return `"${String(val).replace(/"/g, '""')}"`;
}

// Normalizes any of the accepted JSON shapes into a flat row array where every
// row carries a stable `_id` (the original dataset key, when available). This
// id is what lets an AI-generated slide point back to an exact technology
// instead of matching on display name.
// The IEA's own raw ETP source is a plain array of tech objects — no bilingual
// summaries, no case records — shaped like {name, sector[], description, trl[],
// NZErationale, supplyChain[], theme[], keyCountries[], read_more}. This lets a
// user upload that file directly (instead of only our pre-translated cache)
// and still get a working, English-only view rather than a broken/blank one.
function isRawEtpFormat(parsed) {
  if (!Array.isArray(parsed) || parsed.length === 0) return false;
  const sample = parsed[0];
  return typeof sample === 'object' && sample !== null
    && typeof sample.name === 'string' && Array.isArray(sample.trl)
    && !('technology_name_zh' in sample) && !('technology_status_summary_zh' in sample);
}

const RAW_ETP_TRL_YEARS = ['2020', '2021', '2022', '2023', '2024', '2025'];

function slugifyRawEtpName(name, idx) {
  const s = (name || 'tech').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
  return `tech_${String(idx).padStart(4, '0')}_${s.slice(0, 40)}`;
}

// Maps each raw entry onto the same field shape our own bilingual cache uses
// (technology_name, trl_2020.., technology_status_summary_en, etc.) so every
// other part of the app (search, TRL stats, detail view) works unmodified.
// Chinese fields are left blank rather than guessed — sector_zh gets the raw
// English sector name instead of staying empty, so sector filters/labels
// show real text instead of the "未標示" placeholder in either language.
function normalizeRawEtpRows(parsed) {
  return parsed.map((t, i) => {
    const trlArr = Array.isArray(t.trl) ? t.trl : [];
    const trlFields = {};
    let lastKnown = '';
    RAW_ETP_TRL_YEARS.forEach((y, j) => {
      const v = trlArr[j];
      trlFields[`trl_${y}`] = v != null ? v : '';
      if (v != null) lastKnown = v;
    });
    const sectorArr = Array.isArray(t.sector) ? t.sector : [];
    const sectorLeaf = sectorArr[sectorArr.length - 1] || '';
    return {
      _id: slugifyRawEtpName(t.name, i),
      technology_name: t.name || '',
      technology_name_zh: '',
      sector_zh: sectorLeaf,
      sector: sectorArr,
      sector_en: sectorLeaf,
      ...trlFields,
      latest_trl: lastKnown,
      technology_status_summary_en: t.description || '',
      technology_status_summary_zh: '',
      market_dynamics_summary_en: t.NZErationale || '',
      market_dynamics_summary_zh: '',
      description: t.description || '',
      NZErationale: t.NZErationale || '',
      supplyChain: t.supplyChain || [],
      theme: t.theme || [],
      keyCountries: t.keyCountries || [],
      read_more: t.read_more || null,
      linked_records_count: 0,
      initiatives: []
    };
  });
}

function extractRowsWithIds(parsed) {
  if (isRawEtpFormat(parsed)) return normalizeRawEtpRows(parsed);

  let rawRows = [];
  if (Array.isArray(parsed.rows)) rawRows = parsed.rows;
  else if (Array.isArray(parsed)) rawRows = parsed;
  else if (parsed.technologies && typeof parsed.technologies === 'object') {
    rawRows = Object.entries(parsed.technologies).map(([id, tech]) => ({ ...tech, _id: tech._id ?? id }));
  } else if (typeof parsed === 'object' && parsed !== null) {
    rawRows = Object.entries(parsed)
      .filter(([, v]) => typeof v === 'object' && v !== null)
      .map(([id, v]) => ({ ...v, _id: v._id ?? id }));
  }
  return rawRows.map((r, i) => (r._id != null ? r : { ...r, _id: `row_${i}` }));
}

// Shared stats used by both the rule-based strategy overlay and the AI
// analysis fallback, so the two no longer compute TRL/sector/country
// breakdowns with slightly different logic.
function techLabel(tech) {
  return tech.technology_name_zh || tech.technology_name || '未標示';
}

function computeTechStats(techs) {
  const totalFound = techs.length;

  const sectorCounts = {};
  techs.forEach(r => { const s = r.sector_zh || '未標示'; sectorCounts[s] = (sectorCounts[s] || 0) + 1; });
  const topSector = Object.entries(sectorCounts).sort((a, b) => b[1] - a[1])[0]?.[0] || '未標示';

  // { id, name } per bucket (not just ids) so the UI can drill down from a
  // TRL bar into the exact technologies behind it, with a display name that
  // matches slide4's shape — id for lookup/navigation, name for display.
  const trlBreakdown = TRL_STAGE_LABELS.map(stage => ({ stage, count: 0, techs: [] }));
  techs.forEach(r => {
    const trl = parseInt(r.latest_trl) || 0;
    let bucket = -1;
    if (trl > 0 && trl <= 3) bucket = 0;
    else if (trl <= 6) bucket = 1;
    else if (trl <= 8) bucket = 2;
    else if (trl > 8) bucket = 3;
    if (bucket >= 0) { trlBreakdown[bucket].count++; trlBreakdown[bucket].techs.push({ id: r._id, name: techLabel(r) }); }
  });

  const allInitiatives = techs.flatMap(r => r.initiatives || []);
  const countryCounts = {};
  allInitiatives.forEach(i => { if (i.country) countryCounts[i.country] = (countryCounts[i.country] || 0) + 1; });
  const topCountries = Object.entries(countryCounts).sort((a, b) => b[1] - a[1]).map(c => c[0]);

  const topTechs = [...techs].sort((a, b) => (b.linked_records_count || 0) - (a.linked_records_count || 0));

  return { totalFound, topSector, trlBreakdown, allInitiatives, topCountries, topTechs };
}

// Scoped to exactly the two slides AI actually adds value on. slide1-4 are
// deterministic aggregates (counts, rankings) already computed correctly by
// buildFallbackAnalysis — asking the AI to redo them just produces the same
// numbers in different words. This prompt instead pushes it to read the two
// free-text columns (technology_status_summary_zh / market_dynamics_summary_zh)
// and synthesize things counting literally cannot: recurring themes, risks,
// and relationships across technologies.
function buildTrendsAndRecommendationsPrompt(csvStr, uiLang = 'zh') {
  const languageInstruction = uiLang === 'en'
    ? '整個平台目前設定為英文介面。請直接用英文輸出所有內容：slide5、slide6 的 title/subtitle/summary/points/recommendations，以及所有 "name" 顯示用技術名稱，全部都要是英文，不要輸出中文，也不需要使用者另外追問才翻譯。CSV 裡的 "_id" 仍必須照抄原始字串，不受語言影響。'
    : '請用繁體中文輸出所有內容。';

  return `你是內建於「淨零碳排技術查詢平台」的策略分析引擎。以下 CSV 是本次查詢結果的技術清單。

前四頁統計摘要（查詢總覽、技術成熟度分佈、專案活動熱區、關鍵技術清單）已經由本地程式正確算好，不需要你重做，也不在你的輸出範圍內。你的任務是仔細閱讀每一列的 "technology_status_summary_zh"（技術狀態摘要）與 "market_dynamics_summary_zh"（市場動態摘要）這兩欄文字內容，找出單靠數字統計看不出來的東西：多筆技術裡反覆出現的共通主題、風險或機會、技術之間可能的關聯性（互補或競爭）、容易被忽略但值得注意的異常或矛盾之處。不要只是重新條列 TRL 分佈或案例數量這類已經算過的統計數字，那不是你的任務。

${languageInstruction}

嚴格輸出 JSON 格式（不要包含任何 markdown 符號或說明文字，直接輸出純 JSON），只需要 slide5 與 slide6 兩個頂層欄位。CSV 第一欄 "_id" 是每項技術在資料庫中的唯一識別碼，"points"/"recommendations" 裡只要提到具體技術支持你的論點，就要在 "related_techs" 用 {"id": "...", "name": "..."} 的形式列出："id" 必須直接照抄 CSV 裡對應那一列的 "_id" 原始字串，絕對不可自行編造或省略；"name" 是顯示用的技術名稱。

【輸入資料】
${csvStr}

【輸出格式】
{
  "slide5": {
    "title": "趨勢、風險與關聯性",
    "subtitle": "一句話說明本頁的分析角度",
    "summary": "1-2句總體觀察",
    "points": [
      {"text": "具體洞察，引用摘要文字中的具體內容作為依據，不是重述統計數字", "related_techs": [{"id": "CSV裡的_id原值", "name": "技術名稱"}, "..."]},
      {"text": "第二點洞察", "related_techs": [{"id": "CSV裡的_id原值", "name": "技術名稱"}, "..."]},
      {"text": "第三點洞察", "related_techs": [{"id": "CSV裡的_id原值", "name": "技術名稱"}, "..."]}
    ]
  },
  "slide6": {
    "title": "綜合策略建議",
    "subtitle": "基於前一頁的趨勢分析提出的行動方向",
    "recommendations": [
      {"text": "建議一：具體可行的策略建議，呼應 slide5 的某個洞察", "related_techs": [{"id": "CSV裡支持此建議的_id原值", "name": "技術名稱"}, "..."]},
      {"text": "建議二：具體可行的策略建議", "related_techs": [{"id": "CSV裡支持此建議的_id原值", "name": "技術名稱"}, "..."]},
      {"text": "建議三：具體可行的策略建議", "related_techs": [{"id": "CSV裡支持此建議的_id原值", "name": "技術名稱"}, "..."]}
    ]
  }
}`;
}

// Produces only slide1-4 (overview / TRL / hotspots / top techs) — purely
// computed from real data, so this is the "instant, no-wait" baseline every
// strategy overview opens with. Recommendations used to live here as slide5,
// but that's now exclusively an AI-generated addition (see
// buildTrendsAndRecommendationsPrompt) so the two tiers are structurally
// different, not just differently-worded restatements of the same stats.
function buildFallbackAnalysis(techs, uiLang = 'zh') {
  const { totalFound, topSector, trlBreakdown, allInitiatives, topCountries, topTechs } = computeTechStats(techs);
  const top5Countries = topCountries.slice(0, 5);
  const nameFor = (t) => (uiLang === 'en' ? (t.technology_name || techLabel(t)) : techLabel(t));
  const localizedTrlBreakdown = trlBreakdown.map(b => ({ ...b, stage: trlStageLabel(b.stage, uiLang) }));

  const top5Techs = topTechs.slice(0, 5).map(t => ({
    id: t._id,
    name: nameFor(t),
    count: t.linked_records_count || 0
  }));

  if (uiLang === 'en') {
    return {
      slide1: {
        title: "Query Overview",
        subtitle: `Analyzed ${totalFound} net-zero technologies`,
        summary: `This query covers ${totalFound} technologies, mostly concentrated in the "${topSector}" sector, linked to ${allInitiatives.length} policy cases.`,
        kpis: [
          { value: totalFound, label: "Technologies" },
          { value: allInitiatives.length, label: "Related cases" },
          { value: topCountries.length, label: "Countries covered" }
        ]
      },
      slide2: {
        title: "Technology Maturity Analysis",
        subtitle: "TRL distribution",
        summary: `${totalFound} technologies span multiple maturity stages, reflecting a diverse technology ecosystem.`,
        total: totalFound,
        trl_breakdown: localizedTrlBreakdown
      },
      slide3: {
        title: "Project Activity Hotspots",
        subtitle: "Global geographic distribution",
        summary: `Key active regions identified from ${allInitiatives.length} cases, reflecting policy and market momentum.`,
        countries: top5Countries.length > 0 ? top5Countries : ['Insufficient data']
      },
      slide4: {
        title: "Key Technology List",
        subtitle: "Ranked by linked case count",
        top_techs: top5Techs
      },
      ui_labels: DEFAULT_UI_LABELS.en
    };
  }

  return {
    slide1: {
      title: "查詢總覽",
      subtitle: `共分析 ${totalFound} 項淨零技術`,
      summary: `本次查詢涵蓋 ${totalFound} 項技術，主要集中於「${topSector}」領域，共連結 ${allInitiatives.length} 筆政策案例。`,
      kpis: [
        { value: totalFound, label: "技術項目" },
        { value: allInitiatives.length, label: "相關案例" },
        { value: topCountries.length, label: "涵蓋國家" }
      ]
    },
    slide2: {
      title: "技術成熟度分析",
      subtitle: "TRL 分佈狀況",
      summary: `${totalFound} 項技術橫跨各成熟度階段，反映出整體技術生態系的多元發展。`,
      total: totalFound,
      trl_breakdown: localizedTrlBreakdown
    },
    slide3: {
      title: "專案活動熱區",
      subtitle: "全球地理分佈",
      summary: `從 ${allInitiatives.length} 筆案例中識別出關鍵活躍地區，反映政策與市場推動力道。`,
      countries: top5Countries.length > 0 ? top5Countries : ['資料不足']
    },
    slide4: {
      title: "關鍵技術清單",
      subtitle: "依案例連結數排序",
      top_techs: top5Techs
    },
    ui_labels: DEFAULT_UI_LABELS.zh
  };
}

// ─── AI Analysis Modal ──────────────────────────────────────────────────────

function AIAnalysisModal({ geminiData, isAI, isUpgrading, hasApiKey, onUpgrade, onClose, onDownload, isDownloading, totalMatches, techById, onJumpToTech, isRevising, onRefine, versionCount, versionIndex, onGoToVersion, uiLang }) {
  const [currentSlide, setCurrentSlide] = useState(0);
  const [followUpText, setFollowUpText] = useState('');
  const t = MODAL_CHROME_LABELS[uiLang];

  const submitFollowUp = () => {
    if (!followUpText.trim() || isRevising) return;
    onRefine(followUpText.trim());
    setFollowUpText('');
  };

  const slides = geminiData ? [
    geminiData.slide1,
    geminiData.slide2,
    geminiData.slide3,
    geminiData.slide4,
    geminiData.slide5,
    geminiData.slide6,
  ].filter(Boolean) : [];

  const slideCount = slides.length || 4;

  const slideColors = ['#2563EB', '#059669', '#7C3AED', '#D97706', '#0D9488', '#DC2626'];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4" onClick={onClose}>
      <div
        className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl max-h-[90vh] flex flex-col overflow-hidden"
        onClick={e => e.stopPropagation()}
      >
        {/* Modal Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-200 bg-slate-50">
          <div className="flex items-center gap-3">
            <div className="bg-blue-600 text-white rounded-lg p-1.5">
              <Sparkles size={16} />
            </div>
            <div>
              <h2 className="text-base font-bold text-slate-800 flex items-center gap-2">
                {t.title}
                <span className={`text-[10px] px-1.5 py-0.5 rounded-full font-bold ${isAI ? 'bg-purple-100 text-purple-700' : 'bg-slate-100 text-slate-600'}`}>
                  {isAI ? t.aiBadge : t.basicBadge}
                </span>
              </h2>
              <p className="text-xs text-slate-500">{t.subtitle(totalMatches)}</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            {geminiData && (
              <button
                onClick={onDownload}
                disabled={isDownloading}
                className="flex items-center gap-1.5 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-300 text-white text-xs font-medium px-3 py-1.5 rounded-lg transition-colors"
              >
                {isDownloading ? <Loader2 size={14} className="animate-spin" /> : <Download size={14} />}
                {isDownloading ? t.downloading : t.download}
              </button>
            )}
            <button
              onClick={onClose}
              className="text-slate-400 hover:text-slate-600 hover:bg-slate-200 p-1.5 rounded-lg transition-colors"
            >
              <X size={18} />
            </button>
          </div>
        </div>

        {/* Slide Area */}
        <div className="flex-1 overflow-hidden flex flex-col min-h-0">
          {slides.length > 0 ? (
            <>
              {/* Slide Preview */}
              <div className="flex-1 p-6 overflow-auto bg-slate-100">
                <SlidePreview slide={slides[currentSlide]} index={currentSlide} totalSlides={slides.length} accentColor={slideColors[currentSlide]} techById={techById} onJumpToTech={onJumpToTech} uiLabels={geminiData?.ui_labels} uiLang={uiLang} />
              </div>

              {/* Slide Navigation */}
              <div className="border-t border-slate-200 px-6 py-3 bg-white flex items-center justify-between gap-4">
                <button
                  disabled={currentSlide === 0}
                  onClick={() => setCurrentSlide(p => Math.max(0, p - 1))}
                  className="flex items-center gap-1 text-xs font-medium px-3 py-1.5 rounded border border-slate-300 text-slate-600 hover:bg-slate-50 disabled:opacity-30 transition-colors"
                >
                  <ChevronLeft size={14} /> {t.prevSlide}
                </button>

                <div className="flex gap-2">
                  {Array.from({ length: slideCount }).map((_, i) => (
                    <button
                      key={i}
                      onClick={() => setCurrentSlide(i)}
                      className={`rounded-full transition-all ${i === currentSlide
                        ? 'w-6 h-2.5 bg-blue-600'
                        : 'w-2.5 h-2.5 bg-slate-300 hover:bg-slate-400'
                      }`}
                    />
                  ))}
                </div>

                <button
                  disabled={currentSlide === slides.length - 1}
                  onClick={() => setCurrentSlide(p => Math.min(slides.length - 1, p + 1))}
                  className="flex items-center gap-1 text-xs font-medium px-3 py-1.5 rounded border border-slate-300 text-slate-600 hover:bg-slate-50 disabled:opacity-30 transition-colors"
                >
                  {t.nextSlide} <ChevronRight size={14} />
                </button>
              </div>

              {/* Below the slides: either the "upgrade to AI" step (while
                  still on the instant rule-based baseline) or the follow-up
                  refinement composer (once a real AI turn exists) — never
                  both, so this reads as one progressive flow, not two
                  competing controls. */}
              <div className="border-t border-slate-200 px-6 py-3 bg-slate-50">
                {isAI ? (
                  <>
                    {versionCount > 1 && (
                      <div className="flex items-center justify-center gap-3 mb-2">
                        <button
                          onClick={() => onGoToVersion(versionIndex - 1)}
                          disabled={versionIndex <= 0 || isRevising}
                          title={t.prevVersion}
                          className="p-1 rounded hover:bg-slate-200 disabled:opacity-30 disabled:cursor-not-allowed text-slate-500 transition-colors"
                        >
                          <ChevronLeft size={16} />
                        </button>
                        <span className="text-xs text-slate-500 font-medium flex items-center gap-1.5">
                          <RotateCcw size={12} /> {t.versionLabel(versionIndex + 1, versionCount)}
                        </span>
                        <button
                          onClick={() => onGoToVersion(versionIndex + 1)}
                          disabled={versionIndex >= versionCount - 1 || isRevising}
                          title={t.nextVersion}
                          className="p-1 rounded hover:bg-slate-200 disabled:opacity-30 disabled:cursor-not-allowed text-slate-500 transition-colors"
                        >
                          <ChevronRight size={16} />
                        </button>
                      </div>
                    )}
                    <div className="flex gap-2 items-center">
                    <input
                      type="text"
                      value={followUpText}
                      onChange={e => setFollowUpText(e.target.value)}
                      onKeyDown={e => { if (e.key === 'Enter') submitFollowUp(); }}
                      placeholder={t.followUpPlaceholder}
                      disabled={isRevising}
                      className="flex-1 text-sm px-3 py-2 rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-purple-400 disabled:bg-slate-100"
                    />
                    <button
                      onClick={submitFollowUp}
                      disabled={isRevising || !followUpText.trim()}
                      className="flex items-center gap-1.5 bg-gradient-to-r from-purple-600 to-blue-600 hover:from-purple-700 hover:to-blue-700 disabled:from-slate-300 disabled:to-slate-300 disabled:cursor-not-allowed text-white text-sm font-medium px-4 py-2 rounded-lg transition-all flex-shrink-0"
                    >
                      {isRevising ? <Loader2 size={14} className="animate-spin" /> : <Sparkles size={14} />}
                      {isRevising ? t.revising : t.send}
                    </button>
                    </div>
                  </>
                ) : hasApiKey ? (
                  <div className="flex items-center justify-center">
                    <button
                      onClick={onUpgrade}
                      disabled={isUpgrading}
                      className="flex items-center gap-1.5 bg-gradient-to-r from-purple-600 to-blue-600 hover:from-purple-700 hover:to-blue-700 disabled:from-slate-300 disabled:to-slate-300 disabled:cursor-not-allowed text-white text-sm font-medium px-4 py-2 rounded-lg transition-all"
                    >
                      {isUpgrading ? <Loader2 size={14} className="animate-spin" /> : <Sparkles size={14} />}
                      {isUpgrading ? t.upgrading : t.upgradeToAI}
                    </button>
                  </div>
                ) : (
                  <p className="text-xs text-amber-600 text-center">{t.needApiKeyForUpgrade}</p>
                )}
              </div>
            </>
          ) : (
            <div className="flex-1 flex items-center justify-center text-slate-400 text-sm">
              {t.noPreview}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── Slide Preview Component (White/Light style) ─────────────────────────────

function SlidePreview({ slide, index, totalSlides, accentColor, techById, onJumpToTech, uiLabels, uiLang }) {
  const [expandedTrlBucket, setExpandedTrlBucket] = useState(null);
  const [trendTechId, setTrendTechId] = useState(null);
  const labels = { ...DEFAULT_UI_LABELS[uiLang || 'zh'], ...(uiLabels || {}) };

  // Drilling into a TRL bucket only makes sense on the slide you drilled into.
  useEffect(() => {
    setExpandedTrlBucket(null);
    setTrendTechId(null);
  }, [index]);

  if (!slide) return null;

  const titles = ['查詢總覽', '技術成熟度分析', '專案活動熱區', '關鍵技術清單', '趨勢、風險與關聯性', '綜合策略建議'];
  const icons = ['📊', '🔬', '🌍', '🏆', '🔍', '💡'];

  const activeBucket = slide.trl_breakdown && expandedTrlBucket !== null ? slide.trl_breakdown[expandedTrlBucket] : null;
  // Display name comes from Gemini's own "techs" field (so it follows
  // whatever language/rewrite the user asked for) — id still resolves the
  // real record for navigation/trend-chart data, which can't be translated.
  const activeBucketTechs = activeBucket ? (activeBucket.techs || []).filter(t => techById?.has(t.id)) : [];

  return (
    <>
    <div
      className="w-full mx-auto rounded-xl shadow-lg border border-slate-200"
      style={{ minHeight: '393px', maxWidth: '700px', background: '#fff', position: 'relative' }}
    >
      {/* Top accent bar */}
      <div style={{ height: '6px', background: accentColor, width: '100%', borderRadius: '12px 12px 0 0' }} />

      {/* No fixed height / overflow-hidden here anymore — a hard-clipped
          16:9 box kept cutting off slide2/slide5 whenever their content
          (TRL drill-down, longer recommendations + reference chips) grew
          past a "real slide" size. The card now just grows with its
          content, and the modal's own overflow-auto wrapper scrolls it. */}
      <div className="p-8 flex flex-col">
        {/* Slide number badge */}
        <div className="flex items-start justify-between mb-5">
          <div>
            <div
              className="text-xs font-bold px-2 py-0.5 rounded-full inline-block mb-2"
              style={{ background: accentColor + '18', color: accentColor }}
            >
              第 {index + 1} / {totalSlides || 4} 頁
            </div>
            <h2
              className="text-2xl font-bold text-slate-900 leading-tight"
            >
              {icons[index]} {slide.title || titles[index]}
            </h2>
            {slide.subtitle && (
              <p className="text-sm text-slate-500 mt-1">{slide.subtitle}</p>
            )}
          </div>
          <div className="text-right flex-shrink-0">
            <div className="text-xs text-slate-300 font-mono">IEA NET ZERO</div>
            <div className="text-[10px] text-slate-300 mt-0.5 whitespace-nowrap">AI Generated · {new Date().getFullYear()}</div>
          </div>
        </div>

        {/* Divider */}
        <div style={{ width: '48px', height: '3px', background: accentColor, borderRadius: '2px', marginBottom: '16px' }} />

        {/* Content area */}
        <div className="flex-1 flex gap-5 min-h-0">
          {/* Main content */}
          <div className="flex-1 min-w-0">
            {slide.summary && (
              <p className="text-sm text-slate-700 leading-relaxed mb-4">{slide.summary}</p>
            )}

            {/* KPIs (slide 1) */}
            {slide.kpis && Array.isArray(slide.kpis) && (
              <div className="grid grid-cols-3 gap-3 mt-2">
                {slide.kpis.map((kpi, i) => (
                  <div key={i} className="rounded-lg p-3 border" style={{ borderColor: accentColor + '30', background: accentColor + '08' }}>
                    <div className="text-xl font-black" style={{ color: accentColor }}>{kpi.value}</div>
                    <div className="text-xs text-slate-500 mt-0.5">{kpi.label}</div>
                  </div>
                ))}
              </div>
            )}

            {/* Bullet points (slide 5: trends/risks/correlations) — same
                {text, related_techs} shape as recommendations below, just a
                lighter bulleted style since these are observations, not
                action items. Falls back to a plain string for safety if a
                point ever arrives un-objectified. */}
            {slide.points && Array.isArray(slide.points) && (
              <ul className="space-y-2.5">
                {slide.points.map((point, i) => {
                  const text = typeof point === 'string' ? point : point.text;
                  const relatedTechs = ((typeof point === 'string' ? [] : point.related_techs) || []).filter(t => techById?.has(t.id));
                  return (
                    <li key={i} className="flex items-start gap-2 text-sm text-slate-700">
                      <span style={{ color: accentColor, marginTop: '2px', flexShrink: 0 }}>▸</span>
                      <div className="flex-1">
                        <span className="leading-relaxed">{text}</span>
                        {relatedTechs.length > 0 && (
                          <div className="flex flex-wrap gap-1.5 mt-1.5">
                            {relatedTechs.map(t => (
                              <button
                                key={t.id}
                                type="button"
                                onClick={() => onJumpToTech(t.id)}
                                title={uiLang === 'en' ? 'Click to view details' : '點擊查看此技術詳情'}
                                className="text-[11px] px-2 py-0.5 rounded-full font-medium transition-colors hover:brightness-95"
                                style={{ background: '#fff', color: accentColor, border: `1px solid ${accentColor}40` }}
                              >
                                {labels.relatedPrefix}{uiLang === 'en' ? ': ' : '：'}{t.name}
                              </button>
                            ))}
                          </div>
                        )}
                      </div>
                    </li>
                  );
                })}
              </ul>
            )}

            {/* TRL breakdown (slide 2) — bars only; drill-down list & trend
                chart render in a separate panel below the slide (see return),
                since this card is a fixed-aspect-ratio "slide" that clips
                overflow and can't scroll to reveal growing content. */}
            {slide.trl_breakdown && (
              <div className="space-y-2">
                {slide.trl_breakdown.map((item, i) => {
                  const isActive = expandedTrlBucket === i;
                  return (
                    <button
                      key={i}
                      type="button"
                      onClick={() => { setExpandedTrlBucket(isActive ? null : i); setTrendTechId(null); }}
                      disabled={item.count === 0}
                      className={`w-full flex items-center gap-3 text-left disabled:cursor-default rounded-lg -mx-1 px-1 py-0.5 transition-colors ${isActive ? 'bg-slate-50' : ''}`}
                    >
                      <div className="text-xs w-28 flex-shrink-0" style={{ color: isActive ? accentColor : '#64748b', fontWeight: isActive ? 700 : 400 }}>{item.stage}</div>
                      <div className="flex-1 h-5 rounded-full bg-slate-100 overflow-hidden">
                        <div
                          className="h-full rounded-full transition-all duration-700"
                          style={{ width: `${Math.min(100, (item.count / (slide.total || 1)) * 100)}%`, background: item.count > 0 ? accentColor : accentColor + '80' }}
                        />
                      </div>
                      <div className="text-xs font-bold text-slate-700 w-8 text-right">{item.count}</div>
                    </button>
                  );
                })}
                {slide.trl_breakdown.some(b => b.count > 0) && (
                  <p className="text-[11px] text-slate-400 pt-1">{labels.trlBarHint}</p>
                )}
              </div>
            )}

            {/* Countries (slide 3) */}
            {slide.countries && Array.isArray(slide.countries) && (
              <div className="flex flex-wrap gap-2">
                {slide.countries.map((c, i) => (
                  <span
                    key={i}
                    className="text-xs px-2.5 py-1 rounded-full font-medium"
                    style={{ background: accentColor + '15', color: accentColor, border: `1px solid ${accentColor}30` }}
                  >
                    {c}
                  </span>
                ))}
              </div>
            )}

            {/* Ranked list (slide 4) */}
            {slide.top_techs && Array.isArray(slide.top_techs) && (
              <ol className="space-y-2">
                {slide.top_techs.map((tech, i) => {
                  const jumpable = !!(tech.id && techById?.has(tech.id) && onJumpToTech);
                  return (
                    <li key={i}>
                      <button
                        type="button"
                        onClick={jumpable ? () => onJumpToTech(tech.id) : undefined}
                        disabled={!jumpable}
                        title={jumpable ? (uiLang === 'en' ? 'Click to view details' : '點擊查看此技術詳情') : undefined}
                        className={`w-full flex items-center gap-3 text-sm text-left rounded-lg -mx-1 px-1 py-0.5 transition-colors ${jumpable ? 'hover:bg-slate-50 cursor-pointer' : 'cursor-default'}`}
                      >
                        <span
                          className="w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold flex-shrink-0 text-white"
                          style={{ background: accentColor }}
                        >
                          {i + 1}
                        </span>
                        <span className={`text-slate-800 font-medium flex-1 truncate ${jumpable ? 'underline decoration-dotted underline-offset-2' : ''}`}>{tech.name}</span>
                        <span className="text-xs text-slate-400">{tech.count} {labels.caseCountSuffix}</span>
                      </button>
                    </li>
                  );
                })}
              </ol>
            )}

            {/* Recommendations (slide 6) */}
            {slide.recommendations && Array.isArray(slide.recommendations) && (
              <div className="space-y-2.5">
                {slide.recommendations.map((rec, i) => {
                  const text = typeof rec === 'string' ? rec : rec.text;
                  // Display name comes from Gemini's own "related_techs" field
                  // (follows whatever language/rewrite was asked for); id
                  // still has to resolve a real record to stay clickable.
                  const relatedTechs = ((typeof rec === 'string' ? [] : rec.related_techs) || []).filter(t => techById?.has(t.id));
                  return (
                    <div
                      key={i}
                      className="rounded-lg p-3 text-sm text-slate-700 leading-relaxed"
                      style={{ background: accentColor + '08', borderLeft: `3px solid ${accentColor}` }}
                    >
                      <p>{text}</p>
                      {relatedTechs.length > 0 && (
                        <div className="flex flex-wrap gap-1.5 mt-2">
                          {relatedTechs.map(t => (
                            <button
                              key={t.id}
                              type="button"
                              onClick={() => onJumpToTech(t.id)}
                              title={uiLang === 'en' ? 'Click to view details' : '點擊查看此技術詳情'}
                              className="text-[11px] px-2 py-0.5 rounded-full font-medium transition-colors hover:brightness-95"
                              style={{ background: '#fff', color: accentColor, border: `1px solid ${accentColor}40` }}
                            >
                              {labels.relatedPrefix}{uiLang === 'en' ? ': ' : '：'}{t.name}
                            </button>
                          ))}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>

    {/* TRL drill-down panel — deliberately outside the fixed-aspect slide
        card above so it can grow freely and the modal's own scroll area
        (overflow-auto) can reveal it, instead of being clipped. */}
    {activeBucket && (
      <div className="w-full mx-auto mt-4" style={{ maxWidth: '700px' }}>
        <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-4">
          <div className="flex items-center justify-between mb-3">
            <h4 className="text-sm font-bold text-slate-800">
              {activeBucket.stage}　{labels.trlListHeading}（共 {activeBucket.count} {labels.countUnit}）
            </h4>
            <button
              type="button"
              onClick={() => { setExpandedTrlBucket(null); setTrendTechId(null); }}
              className="text-xs text-slate-400 hover:text-slate-600 px-2 py-1 rounded hover:bg-slate-100 transition-colors"
            >
              {labels.collapseButton}
            </button>
          </div>

          {activeBucketTechs.length === 0 ? (
            <p className="text-xs text-slate-400 py-2">{labels.trlListEmpty}</p>
          ) : (
            <>
              <p className="text-[11px] text-slate-400 mb-2">{labels.trlTechHint}</p>
              <div className="flex flex-wrap gap-2">
              {activeBucketTechs.map(t => {
                const isSelected = trendTechId === t.id;
                return (
                  <button
                    key={t.id}
                    type="button"
                    onClick={() => setTrendTechId(isSelected ? null : t.id)}
                    className="text-xs px-3 py-1.5 rounded-full border font-medium transition-colors"
                    style={{
                      background: isSelected ? accentColor : '#fff',
                      color: isSelected ? '#fff' : '#475569',
                      borderColor: isSelected ? accentColor : '#e2e8f0'
                    }}
                  >
                    {t.name}
                  </button>
                );
              })}
              </div>
            </>
          )}

          {trendTechId && activeBucketTechs.some(t => t.id === trendTechId) && (
            <TrlTrendChart
              tech={techById.get(trendTechId)}
              displayName={activeBucketTechs.find(t => t.id === trendTechId)?.name}
              accentColor={accentColor}
              chartSuffix={labels.trendChartSuffix}
              insufficientDataText={labels.trendInsufficientData}
            />
          )}
        </div>
      </div>
    )}
    </>
  );
}

const TRL_TREND_YEARS = ['2020', '2021', '2022', '2023', '2024', '2025'];

// Minimal hand-rolled line chart (no charting library) for a single
// technology's TRL across trl_2020..trl_2025 — kept lightweight and visually
// consistent with the rest of the app rather than pulling in a dependency.
function TrlTrendChart({ tech, displayName, accentColor, chartSuffix, insufficientDataText }) {
  const label = displayName || techLabel(tech);
  const suffix = chartSuffix || DEFAULT_UI_LABELS.zh.trendChartSuffix;
  const points = TRL_TREND_YEARS.map(year => {
    const raw = tech[`trl_${year}`];
    const val = raw !== undefined && raw !== null && raw !== '' ? parseFloat(raw) : null;
    return { year, val: Number.isFinite(val) ? val : null };
  });
  const validCount = points.filter(p => p.val !== null).length;

  if (validCount < 2) {
    return <p className="text-xs text-slate-400 mt-2 py-1">{insufficientDataText || DEFAULT_UI_LABELS.zh.trendInsufficientData}</p>;
  }

  const w = 320, h = 132, padX = 24, padTop = 26, padBottom = 20;
  const minTrl = 1, maxTrl = 9;
  const xStep = (w - padX * 2) / (TRL_TREND_YEARS.length - 1);
  const xFor = (i) => padX + i * xStep;
  const yFor = (v) => h - padBottom - ((v - minTrl) / (maxTrl - minTrl)) * (h - padTop - padBottom);

  const linePath = points
    .map((p, i) => (p.val !== null ? `${xFor(i).toFixed(1)},${yFor(p.val).toFixed(1)}` : null))
    .filter(Boolean)
    .join(' L ');

  return (
    <div className="mt-2 mb-1 p-3 rounded-lg border" style={{ borderColor: accentColor + '30', background: accentColor + '06' }}>
      <p className="text-xs font-semibold mb-1" style={{ color: accentColor }}>
        {label} — {suffix}
      </p>
      <svg viewBox={`0 0 ${w} ${h}`} width="100%" height={h} role="img">
        <title>{`${label} 2020 至 2025 年 TRL 變化`}</title>
        {[1, 5, 9].map(v => (
          <text key={v} x={2} y={yFor(v) + 3} fontSize="8" fill="#94a3b8">{v}</text>
        ))}
        <path d={`M ${linePath}`} fill="none" stroke={accentColor} strokeWidth="2" />
        {points.map((p, i) => p.val !== null ? (
          <g key={i}>
            <circle cx={xFor(i)} cy={yFor(p.val)} r="3.5" fill={accentColor} stroke="#fff" strokeWidth="1" />
            {/* Explicit TRL value above each point, not just the line shape */}
            <text x={xFor(i)} y={yFor(p.val) - 8} fontSize="10" fontWeight="700" fill={accentColor} textAnchor="middle">
              {p.val}
            </text>
            <text x={xFor(i)} y={h - 2} fontSize="8" fill="#64748b" textAnchor="middle">{p.year}</text>
          </g>
        ) : null)}
      </svg>
    </div>
  );
}

// ─── Main App ────────────────────────────────────────────────────────────────

export default function App() {
  // Global platform language — drives APP_LABELS (Dashboard chrome),
  // MODAL_CHROME_LABELS (AI modal chrome), and tells Gemini to generate
  // slide content directly in this language rather than requiring a
  // follow-up "translate to English" request.
  const [uiLang, setUiLang] = useState('zh');
  const [dataset, setDataset] = useState(DEFAULT_JSON);
  const [rows, setRows] = useState(() => extractRowsWithIds(DEFAULT_JSON));
  const [isSampleData, setIsSampleData] = useState(true);
  // True once the user has loaded their own file via handleFileUpload, so
  // the header can offer a way back to the app's own bilingual DATA_URL
  // source without requiring a full page refresh.
  const [isCustomUpload, setIsCustomUpload] = useState(false);

  const [appState, setAppState] = useState({
    status: 'fallback',
    errorMsg: '',
    fileName: '',
    isRawFormat: false
  });

  const [query, setQuery] = useState('');
  const [searchMode, setSearchMode] = useState('subject');
  const [selectedTech, setSelectedTech] = useState(null);
  const [selectedSector, setSelectedSector] = useState('');
  // Lifted out of Dashboard so handleJumpToTech (below) can switch back to
  // the "技術詳情" tab when a slide reference is clicked from inside the AI
  // modal — otherwise selectedTech updates invisibly behind whichever tab
  // the user was already on.
  const [rightPanelTab, setRightPanelTab] = useState('detail');

  // AI Modal state
  const [showAIModal, setShowAIModal] = useState(false);
  // Renamed from the old isGenerating: opening the overview is now instant
  // (rule-based, no API wait), so this only covers the optional in-place
  // "upgrade to AI" step — never a full-modal blocking spinner anymore.
  const [isUpgrading, setIsUpgrading] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);
  const [analysisTotal, setAnalysisTotal] = useState(0);
  // The exact tech list backing the currently open overview, kept around so
  // "upgrade to AI" (triggered from inside the modal, without a fresh
  // searchData argument) can rebuild the same CSV the rule-based pass used.
  const [analysisTechs, setAnalysisTechs] = useState([]);
  const [isRevising, setIsRevising] = useState(false);
  // Every generation/refinement is appended here rather than overwriting a
  // single slot, so the user can navigate back to ANY earlier version (e.g.
  // the one right after their first follow-up), not just undo the last step.
  // geminiData/chatHistory/isAI are derived from whichever version is selected.
  const [versions, setVersions] = useState({ list: [], index: -1 });
  const geminiData = versions.list[versions.index]?.geminiData ?? null;
  const chatHistory = versions.list[versions.index]?.chatHistory ?? [];
  // Whether the currently displayed version is the AI-generated one (vs. the
  // instant rule-based baseline every overview opens with) — drives the
  // modal's badge and whether the composer shows "upgrade" or "follow-up".
  const isAI = !!versions.list[versions.index]?.isAI;
  const [aiProvider, setAiProvider] = useState(() => {
    try { return localStorage.getItem(AI_PROVIDER_STORAGE_KEY) || 'gemini'; }
    catch (e) { return 'gemini'; }
  });
  // One key per provider so switching providers doesn't clobber a key you
  // already typed in for a different one.
  const [apiKeys, setApiKeys] = useState(() => {
    const initial = {};
    try {
      AI_PROVIDER_ORDER.forEach(p => { initial[p] = localStorage.getItem(AI_API_KEY_STORAGE_PREFIX + p) || ''; });
    } catch (e) { /* localStorage unavailable (e.g. private browsing) */ }
    return initial;
  });
  const [openRouterModel, setOpenRouterModel] = useState(() => {
    try { return localStorage.getItem(OPENROUTER_MODEL_STORAGE_KEY) || AI_PROVIDERS.openrouter.defaultModel; }
    catch (e) { return AI_PROVIDERS.openrouter.defaultModel; }
  });
  const apiKey = apiKeys[aiProvider] || '';

  const handleProviderChange = (provider) => {
    setAiProvider(provider);
    try { localStorage.setItem(AI_PROVIDER_STORAGE_KEY, provider); }
    catch (e) { /* localStorage unavailable (e.g. private browsing) */ }
  };

  const handleApiKeyChange = (val) => {
    setApiKeys(prev => ({ ...prev, [aiProvider]: val }));
    try { localStorage.setItem(AI_API_KEY_STORAGE_PREFIX + aiProvider, val); }
    catch (e) { /* localStorage unavailable (e.g. private browsing) — key still works for this session */ }
  };

  const handleOpenRouterModelChange = (val) => {
    setOpenRouterModel(val);
    try { localStorage.setItem(OPENROUTER_MODEL_STORAGE_KEY, val); }
    catch (e) { /* localStorage unavailable (e.g. private browsing) */ }
  };

  // Pulled out of the mount-only effect so a "restore default data" button
  // can re-run the exact same load later, after a user has switched to their
  // own uploaded file and wants to go back to the app's own bilingual source.
  const fetchJsonData = async () => {
    const T = APP_LABELS[uiLang];
    setAppState({ status: 'loading', errorMsg: '', fileName: T.loadingFromGithub });
    try {
      const response = await fetch(DATA_URL);
      if (!response.ok) throw new Error(`狀態碼: ${response.status}`);
      const parsed = await response.json();
      const rawRows = extractRowsWithIds(parsed);

      if (rawRows.length > 0) {
        setDataset(parsed);
        setRows(rawRows);
        setIsSampleData(false);
        setIsCustomUpload(false);
        setAppState({ status: 'success', errorMsg: '', fileName: T.githubLoadSuccess, isRawFormat: false });
        setQuery('');
        setSelectedSector('');
      } else {
        throw new Error(T.remoteFormatError);
      }
    } catch (err) {
      setAppState({ status: 'error', errorMsg: APP_LABELS[uiLang].networkLoadFailed(err.message), fileName: APP_LABELS[uiLang].autoLoadFailed });
    }
  };

  useEffect(() => {
    fetchJsonData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setAppState({ status: 'loading', errorMsg: '', fileName: file.name });
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        if (!evt.target?.result) throw new Error("File is empty");
        const parsed = JSON.parse(evt.target.result);
        const isRawUpload = isRawEtpFormat(parsed);
        const rawRows = extractRowsWithIds(parsed);
        if (rawRows.length > 0) {
          setDataset(parsed);
          setRows(rawRows);
          setIsSampleData(false);
          setIsCustomUpload(true);
          // The raw IEA source has no Chinese fields at all — staying in
          // Chinese mode would just show English fallback text under Chinese
          // chrome, so switch the whole platform to English automatically.
          if (isRawUpload) setUiLang('en');
          setAppState({ status: 'success', errorMsg: '', fileName: file.name, isRawFormat: isRawUpload });
          setQuery('');
          setSelectedSector('');
        } else {
          setAppState({ status: 'error', errorMsg: APP_LABELS[uiLang].fileLoadFailed, fileName: file.name });
        }
      } catch (err) {
        setAppState({ status: 'error', errorMsg: APP_LABELS[uiLang].jsonParseFailed(err.message), fileName: file.name });
      }
    };
    reader.onerror = () => setAppState({ status: 'error', errorMsg: APP_LABELS[uiLang].localFileReadError, fileName: file.name });
    reader.readAsText(file);
  };

  // Lets an AI-generated slide reference ("id": "tech_h2_1") resolve back to
  // the actual technology object so the UI can jump to it.
  const techById = useMemo(() => new Map(rows.map(r => [r._id, r])), [rows]);

  const handleJumpToTech = (id) => {
    const tech = techById.get(id);
    if (tech) {
      setSelectedTech(tech);
      setRightPanelTab('detail');
      setShowAIModal(false);
    }
  };

  // ── AI Analysis via the user's chosen provider ────────────────────────────
  // `contents` is always kept in Gemini's own shape ({role:'user'|'model',
  // parts:[{text}]}) internally — that's the format buildAnalysisPrompt/
  // handleRefineAnalysis/chatHistory already use — and gets adapted to each
  // provider's actual wire format only at call time, right here.
  const toSimpleMessages = (contents) => contents.map(c => ({
    role: c.role === 'model' ? 'assistant' : 'user',
    content: (c.parts || []).map(p => p.text).join('\n')
  }));

  const callAI = async (contents) => {
    const key = apiKey.trim();
    const model = aiProvider === 'openrouter' ? openRouterModel.trim() : AI_PROVIDERS[aiProvider].defaultModel;

    if (aiProvider === 'gemini') {
      const response = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ contents, generationConfig: { responseMimeType: "application/json" } })
        }
      );
      if (response.status === 400 || response.status === 403) throw new Error('invalid-key');
      if (!response.ok) throw new Error(`API Error ${response.status}`);
      const data = await response.json();
      return data.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
    }

    if (aiProvider === 'claude') {
      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": key,
          "anthropic-version": "2023-06-01",
          // Anthropic blocks direct browser-origin calls by default; this
          // header opts back in. The key is user-supplied and only ever
          // stored in the user's own browser, same trust model as the other
          // three providers here — not a new exposure, just Claude's own
          // explicit opt-in flag for it.
          "anthropic-dangerous-direct-browser-access": "true"
        },
        body: JSON.stringify({ model, max_tokens: 8192, messages: toSimpleMessages(contents) })
      });
      if (response.status === 401 || response.status === 403) throw new Error('invalid-key');
      if (!response.ok) throw new Error(`API Error ${response.status}`);
      const data = await response.json();
      return data.content?.[0]?.text || '{}';
    }

    // openai and openrouter both speak the same OpenAI-compatible
    // chat-completions shape.
    const endpoint = aiProvider === 'openai'
      ? 'https://api.openai.com/v1/chat/completions'
      : 'https://openrouter.ai/api/v1/chat/completions';
    const headers = { "Content-Type": "application/json", "Authorization": `Bearer ${key}` };
    if (aiProvider === 'openrouter') {
      headers["HTTP-Referer"] = window.location.origin;
      headers["X-Title"] = "IEA Net-Zero Technology Intelligence Platform";
    }
    const response = await fetch(endpoint, {
      method: "POST",
      headers,
      body: JSON.stringify({ model, messages: toSimpleMessages(contents), response_format: { type: "json_object" } })
    });
    if (response.status === 401 || response.status === 403) throw new Error('invalid-key');
    if (!response.ok) throw new Error(`API Error ${response.status}`);
    const data = await response.json();
    return data.choices?.[0]?.message?.content || '{}';
  };

  const parseAIJson = (rawText) => JSON.parse(rawText.replace(/```json/gi, '').replace(/```/g, '').trim());

  // Opens the strategy overview with an instant, no-wait rule-based summary
  // — step one of a progressive experience. Upgrading to AI (below) is a
  // separate, optional step from inside the modal, not a second competing
  // entry point.
  const handleOpenStrategyOverview = (searchData) => {
    if (!searchData.results || searchData.results.length === 0) return;
    const exportTechs = searchData.results.map(r => r.tech);
    setShowAIModal(true);
    setAnalysisTotal(exportTechs.length);
    setAnalysisTechs(exportTechs);
    setVersions({
      list: [{ label: '基本統計', geminiData: buildFallbackAnalysis(exportTechs, uiLang), chatHistory: [], isAI: false }],
      index: 0
    });
  };

  // slide1-4 always come from the rule-based baseline (version 0) — AI only
  // ever owns slide5/slide6, so every AI turn (upgrade or follow-up) merges
  // its partial response back onto that unchanging foundation rather than
  // being trusted to reproduce the stats slides itself.
  const mergeWithBaseSlides = (aiPartial) => {
    const base = versions.list[0]?.geminiData || {};
    return {
      slide1: base.slide1, slide2: base.slide2, slide3: base.slide3, slide4: base.slide4,
      slide5: aiPartial.slide5, slide6: aiPartial.slide6,
      ui_labels: base.ui_labels
    };
  };

  // Step two, triggered from inside the modal: adds the AI-only slide5/6
  // (trends/risks/correlations, then recommendations) alongside the
  // untouched rule-based slide1-4, in place, without closing or re-opening
  // anything.
  const handleUpgradeToAI = async () => {
    if (!apiKey.trim() || analysisTechs.length === 0) return;

    setIsUpgrading(true);

    const requiredKeys = [
      '_id', 'technology_name_zh', 'sector_zh', 'latest_trl',
      'technology_status_summary_zh', 'market_dynamics_summary_zh', 'linked_records_count'
    ];
    let csvStr = requiredKeys.join(",") + "\n";
    analysisTechs.forEach(tech => {
      const row = requiredKeys.map(k => {
        let val = tech[k] !== undefined && tech[k] !== null ? String(tech[k]) : "";
        return `"${val.replace(/"/g, '""')}"`;
      });
      csvStr += row.join(",") + "\n";
    });

    const promptText = buildTrendsAndRecommendationsPrompt(csvStr, uiLang);

    try {
      const extractedText = await callAI([{ role: 'user', parts: [{ text: promptText }] }]);
      // Only a genuine AI success seeds chat history — follow-up refinement
      // needs a real prior AI turn to build on, not the rule-based baseline.
      const newVersion = {
        label: 'AI 洞察',
        geminiData: mergeWithBaseSlides(parseAIJson(extractedText)),
        chatHistory: [
          { role: 'user', parts: [{ text: promptText }] },
          { role: 'model', parts: [{ text: extractedText }] }
        ],
        isAI: true
      };
      setVersions(v => ({ list: [...v.list, newVersion], index: v.list.length }));
    } catch (err) {
      if (err.message === 'invalid-key') {
        alert('API 金鑰無效或未授權，仍可繼續使用基本統計版本，請確認金鑰是否正確。');
      } else {
        alert('升級為 AI 洞察失敗，請稍後再試。');
      }
    } finally {
      setIsUpgrading(false);
    }
  };

  // Follow-up refinement — reuses the running conversation so only the new
  // instruction needs sending; the model already has the original CSV and
  // its own prior slide5/6 answer in context. Scoped to those two slides
  // only, same as the initial upgrade.
  const handleRefineAnalysis = async (instruction) => {
    if (!geminiData || chatHistory.length === 0 || !instruction.trim() || !apiKey.trim()) return;

    setIsRevising(true);
    const currentLanguageNote = uiLang === 'en'
      ? '目前平台介面是英文模式，除非這個指示明確要求換成別的語言，否則請繼續全部用英文輸出。'
      : '目前平台介面是繁體中文模式，除非這個指示明確要求換成別的語言，否則請繼續全部用繁體中文輸出。';
    const newTurn = {
      role: 'user',
      parts: [{ text: `請根據以下指示調整剛才產生的 slide5、slide6 JSON，維持完全相同的結構，只輸出純 JSON，不要 markdown 或說明文字，也不要輸出 slide1~slide4（那些不歸你管）。${currentLanguageNote}：\n${instruction}` }]
    };

    try {
      const extractedText = await callAI([...chatHistory, newTurn]);
      const newVersion = {
        label: instruction.length > 24 ? `${instruction.slice(0, 24)}…` : instruction,
        geminiData: mergeWithBaseSlides(parseAIJson(extractedText)),
        chatHistory: [...chatHistory, newTurn, { role: 'model', parts: [{ text: extractedText }] }],
        isAI: true
      };
      // Always appended at the end — even if the user had navigated back to
      // an earlier version and asked from there, so nothing already created
      // in this session is ever discarded, only added to.
      setVersions(v => ({ list: [...v.list, newVersion], index: v.list.length }));
    } catch (err) {
      alert('追問失敗，簡報維持原樣，請確認 API 金鑰或稍後再試。');
    } finally {
      setIsRevising(false);
    }
  };

  // Lets the user step through every version created this session (initial +
  // every follow-up), not just undo the single most recent change.
  const handleGoToVersion = (idx) => {
    setVersions(v => ({ ...v, index: Math.max(0, Math.min(v.list.length - 1, idx)) }));
  };

  // ── PPTX Download ─────────────────────────────────────────────────────────
  const handleDownloadPPTX = async () => {
    if (!geminiData) return;
    setIsDownloading(true);
    try {
      if (!(window).PptxGenJS) {
        await new Promise((resolve, reject) => {
          const script = document.createElement('script');
          script.src = 'https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
          script.onload = resolve;
          script.onerror = reject;
          document.head.appendChild(script);
        });
      }

      const PptxGenJS = (window).PptxGenJS;
      const pptx = new PptxGenJS();
      pptx.defineLayout({ name: 'CUSTOM_WIDE', width: 13.333, height: 7.5 });
      pptx.layout = 'CUSTOM_WIDE';
      pptx.title = 'IEA 技術發展策略摘要';

      const C = {
        white: 'FFFFFF', bg: 'F8FAFC', slate50: 'F8FAFC', slate100: 'F1F5F9',
        slate700: '374151', slate500: '64748B', slate400: '94A3B8', slate300: 'CBD5E1',
        accent: [
          ['2563EB', 'EFF6FF'],
          ['059669', 'ECFDF5'],
          ['7C3AED', 'F5F3FF'],
          ['D97706', 'FFFBEB'],
          ['0D9488', 'F0FDFA'],
          ['DC2626', 'FEF2F2'],
        ]
      };
      const FONT = 'Noto Sans';

      // Filtered (not just left sparse) so slides.length reflects what's
      // actually being exported — slide5/6 only exist once upgraded to AI,
      // and this same array also drives the "X / Y" page label per slide.
      const slides = [geminiData.slide1, geminiData.slide2, geminiData.slide3, geminiData.slide4, geminiData.slide5, geminiData.slide6].filter(Boolean);
      const slideIcons = ['📊', '🔬', '🌍', '🏆', '🔍', '💡'];

      // Shared so slide 5 can spawn continuation slides (same header/style)
      // when its recommendations don't fit on one page, without duplicating
      // the header-drawing code for every branch.
      const addSlideHeader = (idx, title, subtitle, pageLabel) => {
        const [accent, accentLight] = C.accent[idx % C.accent.length];
        const slide = pptx.addSlide();
        slide.background = { color: C.white };

        slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.08, fill: { color: accent } });
        slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.08, w: 13.333, h: 1.6, fill: { color: C.bg }, line: { color: C.slate300, width: 0.5 } });
        slide.addText(`${slideIcons[idx]} ${title || ''}`, { x: 0.5, y: 0.2, w: 9.4, h: 0.65, fontFace: FONT, fontSize: 26, bold: true, color: C.slate700 });
        if (subtitle) {
          slide.addText(subtitle, { x: 0.5, y: 0.9, w: 9.4, h: 0.3, fontFace: FONT, fontSize: 11, color: C.slate500 });
        }
        // Footer info now lives up here (small, right-aligned) instead of at
        // the bottom, where it used to sit on top of long slide content.
        slide.addText(pageLabel, { x: 10.2, y: 0.35, w: 2.6, h: 0.3, fontFace: FONT, fontSize: 10, color: C.slate400, align: 'right' });
        slide.addText('IEA ETP Clean Energy Technology Guide · AI Generated', {
          x: 10.2, y: 0.62, w: 2.6, h: 0.5, fontFace: FONT, fontSize: 7.5, color: C.slate400, align: 'right', wrap: true
        });
        slide.addShape(pptx.ShapeType.rect, { x: 0.5, y: 1.55, w: 0.5, h: 0.04, fill: { color: accent } });

        return { slide, accent, accentLight };
      };

      slides.forEach((slideData, idx) => {
        if (!slideData) return;
        const { slide, accent, accentLight } = addSlideHeader(idx, slideData.title, slideData.subtitle, `${idx + 1} / ${slides.length}`);

        // Content
        const contentY = 1.8;

        if (idx === 0 && slideData.summary) {
          slide.addText(slideData.summary, { x: 0.5, y: contentY, w: 12.3, h: 0.7, fontFace: FONT, fontSize: 13, color: C.slate700 });
          if (slideData.kpis) {
            slideData.kpis.forEach((kpi, i) => {
              const x = 0.5 + i * 4.2;
              slide.addShape(pptx.ShapeType.roundRect, { x, y: contentY + 0.9, w: 3.8, h: 1.5, rectRadius: 0.1, fill: { color: accentLight }, line: { color: accent, width: 0.8 } });
              slide.addText(String(kpi.value), { x, y: contentY + 1.05, w: 3.8, h: 0.65, fontFace: FONT, fontSize: 32, bold: true, color: accent, align: 'center' });
              slide.addText(String(kpi.label), { x, y: contentY + 1.75, w: 3.8, h: 0.3, fontFace: FONT, fontSize: 11, color: C.slate500, align: 'center' });
            });
          }
        }

        if (idx === 1 && slideData.trl_breakdown) {
          if (slideData.summary) {
            slide.addText(slideData.summary, { x: 0.5, y: contentY, w: 12.3, h: 0.5, fontFace: FONT, fontSize: 12, color: C.slate700 });
          }
          const total = slideData.total || 1;
          slideData.trl_breakdown.forEach((item, i) => {
            const y = contentY + 0.7 + i * 1.1;
            const barW = Math.max(0.1, (item.count / total) * 9);
            slide.addText(item.stage, { x: 0.5, y, w: 3, h: 0.35, fontFace: FONT, fontSize: 11, color: C.slate500 });
            slide.addShape(pptx.ShapeType.rect, { x: 3.7, y: y + 0.05, w: 9, h: 0.35, fill: { color: C.slate100 } });
            slide.addShape(pptx.ShapeType.rect, { x: 3.7, y: y + 0.05, w: barW, h: 0.35, fill: { color: accent } });
            slide.addText(String(item.count), { x: 12.8, y, w: 0.5, h: 0.35, fontFace: FONT, fontSize: 11, bold: true, color: C.slate700, align: 'right' });
          });
        }

        if (idx === 2) {
          if (slideData.summary) {
            slide.addText(slideData.summary, { x: 0.5, y: contentY, w: 12.3, h: 0.6, fontFace: FONT, fontSize: 12, color: C.slate700 });
          }
          if (slideData.countries) {
            slideData.countries.forEach((country, i) => {
              const col = i % 3, row = Math.floor(i / 3);
              const x = 0.5 + col * 4.2;
              const y = contentY + 0.8 + row * 0.8;
              slide.addShape(pptx.ShapeType.roundRect, { x, y, w: 3.8, h: 0.55, rectRadius: 0.08, fill: { color: accentLight }, line: { color: accent, width: 0.6 } });
              slide.addText(country, { x, y, w: 3.8, h: 0.55, fontFace: FONT, fontSize: 13, bold: true, color: accent, align: 'center', valign: 'middle' });
            });
          }
        }

        if (idx === 3 && slideData.top_techs) {
          slideData.top_techs.forEach((tech, i) => {
            const y = contentY + i * 0.95;
            slide.addShape(pptx.ShapeType.ellipse, { x: 0.5, y: y + 0.05, w: 0.45, h: 0.45, fill: { color: accent } });
            slide.addText(String(i + 1), { x: 0.5, y: y + 0.05, w: 0.45, h: 0.45, fontFace: FONT, fontSize: 12, bold: true, color: C.white, align: 'center', valign: 'middle' });
            slide.addText(tech.name, { x: 1.15, y: y + 0.05, w: 9.5, h: 0.45, fontFace: FONT, fontSize: 13, color: C.slate700, valign: 'middle' });
            slide.addText(`${tech.count} 筆`, { x: 10.8, y: y + 0.05, w: 2, h: 0.45, fontFace: FONT, fontSize: 11, color: C.slate500, align: 'right', valign: 'middle' });
          });
        }

        const boxItems = slideData.points || slideData.recommendations;
        if ((idx === 4 || idx === 5) && boxItems) {
          // slide5 (points: trends/risks/correlations) and slide6
          // (recommendations) share the same {text, related_techs} shape, so
          // one rendering pass handles both — just fed from whichever field
          // the slide actually has.
          //
          // A fixed box height per item guessed wrong whenever the AI's text
          // ran longer than expected. Instead, estimate line count from the
          // actual text length and stack items with a running cursor. If an
          // item would run past the bottom margin, spill onto a fresh
          // continuation slide (same header) rather than clipping it.
          const estimateLines = (str) => Math.max(1, Math.ceil((str || '').length / 46));
          const contentBottomLimit = 7.15;
          let activeSlide = slide;
          let activeAccent = accent;
          let activeAccentLight = accentLight;
          let cursorY = contentY;
          let continuationCount = 0;

          boxItems.forEach((rec) => {
            const recText = typeof rec === 'string' ? rec : rec.text;
            const relatedTechs = typeof rec === 'string' ? [] : (rec.related_techs || []);
            const relatedNames = relatedTechs.filter(t => techById.has(t.id)).map(t => t.name);

            const textH = estimateLines(recText) * 0.26 + 0.15;
            const relatedH = relatedNames.length > 0 ? 0.32 : 0;
            const boxH = textH + relatedH + 0.15;

            if (cursorY + boxH > contentBottomLimit) {
              continuationCount += 1;
              const created = addSlideHeader(idx, `${slideData.title || ''}（續）`, slideData.subtitle, `${idx + 1} / ${slides.length}・續${continuationCount}`);
              activeSlide = created.slide;
              activeAccent = created.accent;
              activeAccentLight = created.accentLight;
              cursorY = contentY;
            }

            const y = cursorY;
            activeSlide.addShape(pptx.ShapeType.roundRect, { x: 0.5, y, w: 12.3, h: boxH, rectRadius: 0.1, fill: { color: activeAccentLight }, line: { color: activeAccent, width: 0.8 } });
            activeSlide.addShape(pptx.ShapeType.rect, { x: 0.5, y, w: 0.08, h: boxH, fill: { color: activeAccent } });
            activeSlide.addText(recText, {
              x: 0.8, y: y + 0.1, w: 11.8, h: textH,
              fontFace: FONT, fontSize: 12, color: C.slate700, valign: 'top', wrap: true, autoFit: true
            });
            if (relatedNames.length > 0) {
              const relatedPrefix = geminiData?.ui_labels?.relatedPrefix || DEFAULT_UI_LABELS[uiLang]?.relatedPrefix || DEFAULT_UI_LABELS.zh.relatedPrefix;
              const separator = uiLang === 'en' ? ': ' : '：';
              const joiner = uiLang === 'en' ? ', ' : '、';
              activeSlide.addText(`${relatedPrefix}${separator}${relatedNames.join(joiner)}`, {
                x: 0.8, y: y + 0.1 + textH, w: 11.8, h: relatedH,
                fontFace: FONT, fontSize: 9.5, italic: true, color: C.slate500, valign: 'top', autoFit: true
              });
            }

            cursorY += boxH + 0.25;
          });
        }
      });

      await pptx.writeFile({ fileName: `IEA_AI_策略簡報_${new Date().getTime()}.pptx` });
    } catch (error) {
      alert("下載簡報時發生錯誤: " + error.message);
    } finally {
      setIsDownloading(false);
    }
  };

  return (
    <>
      <Dashboard
        dataset={dataset}
        rows={rows}
        isSampleData={isSampleData}
        appState={appState}
        onUpload={handleFileUpload}
        isCustomUpload={isCustomUpload}
        onRestoreDefault={fetchJsonData}
        query={query}
        setQuery={setQuery}
        searchMode={searchMode}
        setSearchMode={setSearchMode}
        selectedTech={selectedTech}
        setSelectedTech={setSelectedTech}
        selectedSector={selectedSector}
        setSelectedSector={setSelectedSector}
        onOpenStrategyOverview={handleOpenStrategyOverview}
        apiKey={apiKey}
        onApiKeyChange={handleApiKeyChange}
        aiProvider={aiProvider}
        onProviderChange={handleProviderChange}
        openRouterModel={openRouterModel}
        onOpenRouterModelChange={handleOpenRouterModelChange}
        rightPanelTab={rightPanelTab}
        setRightPanelTab={setRightPanelTab}
        uiLang={uiLang}
        setUiLang={setUiLang}
      />

      {showAIModal && (
        <AIAnalysisModal
          geminiData={geminiData}
          isAI={isAI}
          isUpgrading={isUpgrading}
          hasApiKey={!!apiKey.trim()}
          onUpgrade={handleUpgradeToAI}
          onClose={() => setShowAIModal(false)}
          onDownload={handleDownloadPPTX}
          isDownloading={isDownloading}
          totalMatches={analysisTotal}
          techById={techById}
          onJumpToTech={handleJumpToTech}
          isRevising={isRevising}
          onRefine={handleRefineAnalysis}
          versionCount={versions.list.length}
          versionIndex={versions.index}
          onGoToVersion={handleGoToVersion}
          uiLang={uiLang}
        />
      )}

      {/* Jumping to a technology closes the modal; this lets you get back to the
          already-generated slides instantly without re-calling Gemini. */}
      {!showAIModal && geminiData && (
        <button
          onClick={() => setShowAIModal(true)}
          className="fixed bottom-5 right-5 z-40 flex items-center gap-2 bg-gradient-to-r from-purple-600 to-blue-600 hover:from-purple-700 hover:to-blue-700 text-white text-sm font-medium pl-4 pr-5 py-2.5 rounded-full shadow-lg transition-all"
        >
          <Sparkles size={16} />
          {APP_LABELS[uiLang].backToAiSlides}
        </button>
      )}
    </>
  );
}

// ─── Dashboard ───────────────────────────────────────────────────────────────

function Dashboard({
  dataset, rows, isSampleData, appState, onUpload,
  isCustomUpload, onRestoreDefault,
  query, setQuery, searchMode, setSearchMode,
  selectedTech, setSelectedTech,
  selectedSector, setSelectedSector,
  onOpenStrategyOverview,
  apiKey, onApiKeyChange,
  aiProvider, onProviderChange,
  openRouterModel, onOpenRouterModelChange,
  rightPanelTab, setRightPanelTab,
  uiLang, setUiLang
}) {
  const L = APP_LABELS[uiLang];
  const [showAllInitiatives, setShowAllInitiatives] = useState(false);
  const [showApiKey, setShowApiKey] = useState(false);
  const [currentPage, setCurrentPage] = useState(1);
  const listContainerRef = useRef(null);

  const techCount = rows.length;
  const initiativeCount = dataset.source_stats?.initiative_rows ??
    rows.reduce((sum, tech) => sum + (Array.isArray(tech.initiatives) ? tech.initiatives.length : 0), 0);

  const datasetVersion = (dataset.dataset_version || "IEA ETP Clean Energy Tech").replace(/ 2026/g, '');

  const searchData = useMemo(() => {
    if (!query.trim() && !selectedSector) return { results: [], totalMatches: 0 };
    const lowerQuery = query.trim().toLowerCase();
    const lowerSector = selectedSector.trim().toLowerCase();

    const expandedQueries = lowerQuery ? (SEARCH_ALIASES[lowerQuery] || [lowerQuery]) : [];

    let results = rows.map(tech => {
      let score = 0, isDirectHit = false, isDescHit = false;
      const sectorTexts = [tech.sector_en, tech.sector_en_path, tech.sector_zh, tech.sector_zh_path, tech.breadcrumb, ...(Array.isArray(tech.sector) ? tech.sector : [])].filter(Boolean).map(s => String(s).toLowerCase());
      const isSectorMatch = selectedSector ? sectorTexts.some(t => t.includes(lowerSector)) : true;
      const subjectTexts = [tech.technology_name, tech.technology_name_zh].filter(Boolean).map(s => String(s).toLowerCase());
      const descTexts = [tech.description, tech.description_en, tech.description_zh, tech.technology_status_summary_zh, tech.market_dynamics_summary_zh, tech.NZErationale, tech.NZErationale_zh, tech.nze_rationale, tech.supplyChain, tech.supply_chain, tech.theme].filter(Boolean).map(s => String(s).toLowerCase());

      let passesSectorFilter = true, passesQueryFilter = true;

      if (searchMode === 'subject') {
        if (selectedSector) { if (!isSectorMatch) passesSectorFilter = false; else { score += 1; isDirectHit = true; } }
        if (expandedQueries.length > 0) {
          let queryHit = false;
          for (const q of expandedQueries) {
            for (const text of subjectTexts) { const hits = countOccurrences(text, q); if (hits > 0) { score += hits * 5; queryHit = true; isDirectHit = true; } }
          }
          if (!queryHit) passesQueryFilter = false;
        }
      } else if (searchMode === 'fulltext') {
        if (selectedSector) {
          if (isSectorMatch) { score += 1; isDirectHit = true; }
          else {
            let sectorDescHit = false;
            for (const text of descTexts) { const hits = countOccurrences(text, lowerSector); if (hits > 0) { score += hits; sectorDescHit = true; isDescHit = true; } }
            if (!sectorDescHit) passesSectorFilter = false;
          }
        }
        if (expandedQueries.length > 0) {
          let queryHit = false;
          for (const q of expandedQueries) {
            for (const text of subjectTexts) { const hits = countOccurrences(text, q); if (hits > 0) { score += hits * 5; queryHit = true; isDirectHit = true; } }
            for (const text of descTexts) { const hits = countOccurrences(text, q); if (hits > 0) { score += hits; queryHit = true; isDescHit = true; } }
          }
          if (!queryHit) passesQueryFilter = false;
        }
      }

      if (passesSectorFilter && passesQueryFilter && score > 0) return { tech, score, isDirectHit, isDescHit };
      return null;
    }).filter(Boolean);

    results.sort((a, b) => {
      if (a.isDirectHit !== b.isDirectHit) return a.isDirectHit ? -1 : 1;
      if (b.score !== a.score) return b.score - a.score;
      return (b.tech.linked_records_count || 0) - (a.tech.linked_records_count || 0);
    });

    return { results, totalMatches: results.length };
  }, [query, rows, searchMode, selectedSector]);

  const totalMatches = searchData.totalMatches;
  const totalPages = Math.max(1, Math.ceil(totalMatches / 20));

  const paginatedResults = useMemo(() => {
    const startIndex = (currentPage - 1) * 20;
    return searchData.results.slice(startIndex, startIndex + 20);
  }, [searchData.results, currentPage]);

  useEffect(() => {
    setCurrentPage(1);
    setSelectedTech(searchData.results[0]?.tech || null);
  }, [searchData.results, searchMode]);

  useEffect(() => {
    if (listContainerRef.current) listContainerRef.current.scrollTop = 0;
  }, [currentPage]);

  useEffect(() => { setShowAllInitiatives(false); }, [selectedTech]);

  const sortedInitiatives = useMemo(() => {
    if (!selectedTech || !selectedTech.initiatives) return [];
    return [...selectedTech.initiatives].sort((a, b) => (parseInt(b.year) || 0) - (parseInt(a.year) || 0));
  }, [selectedTech]);

  const displayedInitiatives = showAllInitiatives ? sortedInitiatives : sortedInitiatives.slice(0, 3);

  // Shared by both the "export all matched results" card (strategy tab) and
  // the "export just this technology" button (left panel) — same CSV shape,
  // just a different tech list and filename going in.
  const handleExportTechsCsv = (techs, filenamePrefix) => {
    if (!techs || techs.length === 0) return;
    const allKeysSet = new Set();
    techs.forEach(tech => Object.keys(tech).forEach(key => { if (key !== '_id') allKeysSet.add(key); }));
    const preferredOrder = ['technology_name_zh', 'technology_name', 'sector_zh', 'latest_trl', 'technology_status_summary_zh', 'market_dynamics_summary_zh', 'linked_records_count'];
    const otherKeys = Array.from(allKeysSet).filter(k => !preferredOrder.includes(k));
    const allKeys = [...preferredOrder.filter(k => allKeysSet.has(k)), ...otherKeys];
    let csvContent = "﻿";
    csvContent += allKeys.map(key => escapeCsv(key)).join(",") + "\n";
    techs.forEach(tech => {
      const row = allKeys.map(key => {
        let val = tech[key];
        if (typeof val === 'object' && val !== null) val = JSON.stringify(val);
        return escapeCsv(val);
      });
      csvContent += row.join(",") + "\n";
    });
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `${filenamePrefix}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleExportListCsv = () => {
    const exportTechs = (searchData.results || []).map(r => r.tech);
    handleExportTechsCsv(exportTechs, `IEA_完整查詢結果_${exportTechs.length}筆`);
  };

  const handleExportSelectedTechCsv = () => {
    if (!selectedTech) return;
    const rawName = (uiLang === 'en' ? selectedTech.technology_name : selectedTech.technology_name_zh) || selectedTech.technology_name || 'tech';
    const safeName = String(rawName).replace(/[\\/:*?"<>|]/g, '').slice(0, 40);
    handleExportTechsCsv([selectedTech], `IEA_${safeName}`);
  };

  return (
    <div className="h-screen bg-slate-50 flex flex-col font-sans overflow-hidden">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 px-6 py-4 flex flex-col md:flex-row justify-between items-center shadow-sm z-10 flex-shrink-0">
        <div className="flex-1">
          <h1 className="text-xl font-bold text-slate-800">{L.appTitle}</h1>
          <div className="text-sm mt-1.5 flex items-center gap-2 flex-wrap">
            <span className="text-slate-500 font-medium">{L.datasetVersion(isSampleData ? L.sampleData : datasetVersion)}</span>
            {appState.status === 'fallback' && <span className="bg-amber-100 text-amber-800 px-2 py-0.5 rounded text-xs border border-amber-200">{L.fallbackBadge}</span>}
            {appState.status === 'loading' && <span className="bg-blue-100 text-blue-800 px-2 py-0.5 rounded text-xs border border-blue-200 animate-pulse">{L.loadingBadge}</span>}
            {appState.status === 'success' && (
              <span className="bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded text-xs border border-emerald-200 flex items-center gap-1">
                <CheckCircle2 size={12} /> {L.successBadge(techCount, initiativeCount)}
              </span>
            )}
            {appState.status === 'success' && appState.isRawFormat && (
              <span className="bg-sky-100 text-sky-800 px-2 py-0.5 rounded text-xs border border-sky-200">{L.rawFormatDetected}</span>
            )}
            {appState.status === 'error' && (
              <span className="bg-red-100 text-red-800 px-2 py-0.5 rounded text-xs border border-red-200 flex items-center gap-1">
                <AlertCircle size={12} /> {appState.errorMsg}
              </span>
            )}
          </div>
        </div>
        <div className="mt-4 md:mt-0 flex flex-wrap items-center gap-4 text-sm">

          <button
            type="button"
            onClick={() => setUiLang(l => (l === 'zh' ? 'en' : 'zh'))}
            title="Switch platform language / 切換平台語言"
            className="text-xs font-medium px-3 py-1.5 rounded-full border border-slate-300 text-slate-600 hover:bg-slate-100 transition-colors flex-shrink-0"
          >
            {uiLang === 'zh' ? 'EN' : '中文'}
          </button>

          <select
            value={aiProvider}
            onChange={e => onProviderChange(e.target.value)}
            title={L.aiProviderLabel}
            className="text-xs font-medium px-2 py-1.5 rounded-lg border border-slate-300 text-slate-700 bg-white focus:outline-none focus:ring-2 focus:ring-blue-500"
          >
            {AI_PROVIDER_ORDER.map(p => <option key={p} value={p}>{AI_PROVIDERS[p].label}</option>)}
          </select>

          {aiProvider === 'openrouter' && (
            <input
              type="text"
              value={openRouterModel}
              onChange={e => onOpenRouterModelChange(e.target.value)}
              placeholder={L.openRouterModelPlaceholder}
              className="text-xs text-slate-700 placeholder:text-slate-400 border border-slate-300 rounded-lg px-3 py-1.5 focus:outline-none focus:ring-2 focus:ring-blue-500 w-44"
              autoComplete="off"
              spellCheck={false}
            />
          )}

          <div className="flex items-center gap-1.5 bg-white border border-slate-300 rounded-lg pl-3 pr-1.5 py-1.5">
            <KeyRound size={14} className={apiKey.trim() ? 'text-emerald-500' : 'text-slate-400'} />
            <input
              type={showApiKey ? 'text' : 'password'}
              value={apiKey}
              onChange={e => onApiKeyChange(e.target.value)}
              placeholder={L.apiKeyPlaceholder(AI_PROVIDERS[aiProvider].label)}
              className="text-xs text-slate-700 placeholder:text-slate-400 focus:outline-none w-40"
              autoComplete="off"
              spellCheck={false}
            />
            <button
              type="button"
              onClick={() => setShowApiKey(v => !v)}
              className="text-slate-400 hover:text-slate-600 p-1 rounded transition-colors"
              title={showApiKey ? L.hideKey : L.showKey}
            >
              {showApiKey ? <EyeOff size={14} /> : <Eye size={14} />}
            </button>
          </div>

          <label className="cursor-pointer bg-slate-800 hover:bg-slate-700 text-white font-medium py-1.5 px-4 rounded-lg transition-colors flex items-center gap-2">
            <FileJson size={16} />
            <span>{L.uploadJson}</span>
            <input type="file" accept=".json" className="hidden" onChange={onUpload} />
          </label>

          {isCustomUpload && (
            <button
              type="button"
              onClick={onRestoreDefault}
              title={L.restoreDefaultData}
              className="flex items-center gap-1.5 text-xs font-medium px-3 py-1.5 rounded-lg border border-slate-300 text-slate-600 hover:bg-slate-100 transition-colors"
            >
              <RotateCcw size={14} /> {L.restoreDefaultData}
            </button>
          )}
        </div>
      </header>

      {/* Main layout */}
      <main className="flex-1 flex flex-col md:flex-row overflow-hidden min-h-0">

        {/* Left panel */}
        <section className="w-full md:w-[320px] lg:w-[340px] flex-shrink-0 border-b md:border-b-0 md:border-r border-slate-200 bg-white flex flex-col h-1/2 md:h-full min-h-0">
          <div className="p-4 border-b border-slate-100 bg-slate-50 flex-shrink-0">
            <div className="flex gap-2 mb-3">
              <button onClick={() => setSearchMode('subject')} className={`flex-1 text-xs py-1.5 rounded-md font-medium transition-colors border ${searchMode === 'subject' ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-50'}`}>{L.subjectSearch}</button>
              <button onClick={() => setSearchMode('fulltext')} className={`flex-1 text-xs py-1.5 rounded-md font-medium transition-colors border ${searchMode === 'fulltext' ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-50'}`}>{L.fulltextSearch}</button>
            </div>

            <div className="relative mb-3">
              <select value={selectedSector} onChange={e => setSelectedSector(e.target.value)} className="w-full pl-3 pr-8 py-2 rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm appearance-none bg-white font-medium text-slate-700">
                <option value="">{L.allSectors}</option>
                {Object.entries(FIXED_SECTORS).map(([sector, subsectors]) => (
                  <optgroup key={sector} label={`${sector} (${SECTOR_TRANSLATIONS[sector] || ''})`}>
                    <option value={sector}>{uiLang === 'en' ? sector : `全部 ${sector} (${SECTOR_TRANSLATIONS[sector] || ''})`}</option>
                    {subsectors.map(sub => <option key={sub} value={sub}>↳ {uiLang === 'en' ? sub : `${sub} (${SECTOR_TRANSLATIONS[sub] || ''})`}</option>)}
                  </optgroup>
                ))}
              </select>
              <div className="pointer-events-none absolute inset-y-0 right-0 flex items-center px-2 text-slate-500">
                <svg className="fill-current h-4 w-4" viewBox="0 0 20 20"><path d="M9.293 12.95l.707.707L15.657 8l-1.414-1.414L10 10.828 5.757 6.586 4.343 8z" /></svg>
              </div>
            </div>

            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
              <input type="text" placeholder={L.searchPlaceholder} className="w-full pl-10 pr-4 py-2.5 rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm transition-shadow" value={query} onChange={e => setQuery(e.target.value)} />
            </div>
            <div className="text-[10.5px] text-slate-500 mt-2 leading-relaxed">{L.searchHint}</div>

            <div className="mt-3 flex justify-between items-center">
              <div className="text-xs text-slate-600 font-medium">{totalMatches > 0 ? L.resultsCount(totalMatches, currentPage, totalPages) : L.noResultsShort}</div>
              {selectedTech && (
                <button onClick={handleExportSelectedTechCsv} className="flex items-center gap-1 text-[11px] font-medium text-emerald-600 hover:text-emerald-700 bg-emerald-50 hover:bg-emerald-100 px-2 py-1 rounded transition-colors">
                  <Download size={12} /> {L.exportSingle}
                </button>
              )}
            </div>
          </div>

          <div ref={listContainerRef} className="flex-1 overflow-y-auto p-2 bg-white">
            {!query.trim() && !selectedSector ? (
              <div className="h-full flex flex-col items-center justify-center text-slate-400 p-6 text-center">
                <Search size={32} className="mb-2 opacity-50" />
                <p>{L.emptyPrompt}</p>
              </div>
            ) : totalMatches === 0 ? (
              <div className="h-full flex flex-col items-center justify-center text-slate-400 p-6 text-center">
                <AlertCircle size={32} className="mb-2 opacity-50" />
                <p>{L.noMatch}</p>
              </div>
            ) : (
              <ul className="space-y-1">
                {paginatedResults.map((item, idx) => {
                  const tech = item.tech;
                  const isSelected = selectedTech === tech;
                  const primaryName = uiLang === 'en' ? (tech.technology_name || techLabel(tech)) : (tech.technology_name_zh || tech.technology_name || L.unlabeled);
                  const secondaryName = uiLang === 'en' ? tech.technology_name_zh : tech.technology_name;
                  return (
                    <li key={idx}>
                      <button onClick={() => { setSelectedTech(tech); setRightPanelTab('detail'); }} className={`w-full text-left p-2.5 rounded-lg transition-colors flex flex-col gap-1 ${isSelected ? 'bg-blue-50 border border-blue-200' : 'hover:bg-slate-50 border border-transparent'}`}>
                        <div className="flex justify-between items-start gap-2">
                          <span className="font-semibold text-slate-800 text-sm line-clamp-1">{primaryName}</span>
                          <span className={`flex-shrink-0 text-[10px] px-1.5 py-0.5 rounded font-medium ${item.isDirectHit ? 'bg-blue-100 text-blue-700' : 'bg-orange-100 text-orange-700'}`}>{item.isDirectHit ? L.directHit : L.descHit}</span>
                        </div>
                        <span className="text-xs text-slate-500 line-clamp-1">{secondaryName || L.unlabeled}</span>
                        <div className="flex justify-between items-center mt-0.5">
                          <span className="text-[10px] bg-slate-100 text-slate-600 px-2 py-0.5 rounded border border-slate-200">{tech.sector_zh || L.unlabeled}</span>
                          <span className="text-[10px] text-blue-600 font-medium">{L.caseCountInline(tech.linked_records_count || 0)}</span>
                        </div>
                      </button>
                    </li>
                  );
                })}
              </ul>
            )}
          </div>

          {totalMatches > 0 && (
            <div className="p-3 border-t border-slate-200 bg-white flex justify-between items-center flex-shrink-0">
              <button onClick={() => setCurrentPage(p => Math.max(1, p - 1))} disabled={currentPage === 1} className="text-xs font-medium px-3 py-1.5 rounded border border-slate-300 text-slate-600 hover:bg-slate-50 disabled:opacity-30 transition-colors flex items-center gap-1"><ChevronLeft size={14} />{L.prevPage}</button>
              <button onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))} disabled={currentPage === totalPages} className="text-xs font-medium px-3 py-1.5 rounded border border-slate-300 text-slate-600 hover:bg-slate-50 disabled:opacity-30 transition-colors flex items-center gap-1">{L.nextPage}<ChevronRight size={14} /></button>
            </div>
          )}
        </section>

        {/* Right panels */}
        <section className="flex-1 flex flex-col bg-slate-50 relative overflow-hidden h-1/2 md:h-full min-h-0">

          {/* Tab bar — detail view vs. aggregate query strategy are parallel,
              not nested: the strategy overview acts on the whole result set
              (searchData), same as the export button on the left, not on
              whichever single technology happens to be selected. */}
          <div className="flex border-b border-slate-200 bg-white px-4 md:px-6 flex-shrink-0">
            <button
              onClick={() => setRightPanelTab('detail')}
              className={`px-4 py-3 text-sm font-medium border-b-2 -mb-px transition-colors flex items-center gap-1.5 ${rightPanelTab === 'detail' ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-500 hover:text-slate-700'}`}
            >
              <FileText size={14} /> {L.detailTab}
            </button>
            <button
              onClick={() => setRightPanelTab('strategy')}
              className={`px-4 py-3 text-sm font-medium border-b-2 -mb-px transition-colors flex items-center gap-1.5 ${rightPanelTab === 'strategy' ? 'border-purple-600 text-purple-600' : 'border-transparent text-slate-500 hover:text-slate-700'}`}
            >
              <Sparkles size={14} /> {L.strategyTab}
            </button>
          </div>

          <div className="flex-1 flex flex-col lg:flex-row overflow-hidden min-h-0">
            {rightPanelTab === 'detail' && (
              <div className="flex-1 flex flex-col overflow-y-auto p-4 md:p-6 w-full">
                {!selectedTech ? (
                  <div className="m-auto text-center text-slate-400">
                    <FileText size={48} className="mx-auto mb-4 opacity-50" />
                    <p>{L.selectPrompt}</p>
                  </div>
                ) : (
                  <div className="max-w-3xl mx-auto w-full space-y-5 pb-20">
                    {/* Tech header */}
                    <div className="bg-white p-5 md:p-6 rounded-xl shadow-sm border border-slate-200">
                      <h2 className="text-xl md:text-2xl font-bold text-slate-900 mb-1">{uiLang === 'en' ? (selectedTech.technology_name || techLabel(selectedTech)) : (selectedTech.technology_name_zh || selectedTech.technology_name || L.unlabeled)}</h2>
                      <p className="text-xs md:text-sm text-slate-500 font-mono mb-4">{(uiLang === 'en' ? selectedTech.technology_name_zh : selectedTech.technology_name) || L.unlabeled}</p>
                      <div className="flex flex-wrap gap-2 text-xs font-medium">
                        <span className="bg-blue-50 text-blue-700 border border-blue-100 px-2 py-1 rounded">{L.fieldSector((uiLang === 'en' ? selectedTech.sector_en : null) || selectedTech.sector_zh || L.unlabeled)}</span>
                        <span className="bg-slate-100 text-slate-700 border border-slate-200 px-2 py-1 rounded">{L.fieldTrl(selectedTech.latest_trl || L.unlabeled)}</span>
                        <span className="bg-emerald-50 text-emerald-700 border border-emerald-200 px-2 py-1 rounded">{L.fieldCaseCount(selectedTech.linked_records_count || 0)}</span>
                      </div>
                    </div>

                    {/* Summary */}
                    <div className="bg-white p-5 md:p-6 rounded-xl shadow-sm border border-slate-200">
                      <div className="border-b border-slate-200 pb-4 mb-4">
                        <h3 className="text-lg font-semibold text-slate-800 mb-3">{L.summaryTitle}</h3>
                        <div className="bg-amber-50 text-amber-700 px-3 py-2 rounded text-xs font-medium border border-amber-200 inline-flex items-start gap-1.5">
                          <Info size={14} className="mt-0.5 flex-shrink-0" />
                          <span>{L.aiDisclaimer}</span>
                        </div>
                      </div>
                      <div className="space-y-5">
                        <div className="space-y-1.5">
                          <h4 className="text-sm font-bold text-slate-600">{L.summaryTitle}</h4>
                          {(uiLang === 'en'
                            ? (selectedTech.technology_status_summary_en || selectedTech.description_en || selectedTech.description || selectedTech.technology_status_summary_zh)
                            : (selectedTech.technology_status_summary_zh || selectedTech.description_zh || selectedTech.description))
                            ? <p className="text-sm text-slate-700 leading-relaxed">{uiLang === 'en'
                                ? (selectedTech.technology_status_summary_en || selectedTech.description_en || selectedTech.description || selectedTech.technology_status_summary_zh)
                                : (selectedTech.technology_status_summary_zh || selectedTech.description_zh || selectedTech.description)}</p>
                            : <p className="text-sm text-slate-400 italic">{L.noSummary}</p>}
                        </div>
                        <div className="space-y-1.5 pt-4 border-t border-slate-100">
                          <h4 className="text-sm font-bold text-slate-600">{L.marketTitle}</h4>
                          {(uiLang === 'en'
                            ? (selectedTech.market_dynamics_summary_en || selectedTech.NZErationale_en || selectedTech.NZErationale || selectedTech.market_dynamics_summary_zh)
                            : (selectedTech.market_dynamics_summary_zh || selectedTech.NZErationale_zh || selectedTech.NZErationale))
                            ? <p className="text-sm text-slate-700 leading-relaxed">{uiLang === 'en'
                                ? (selectedTech.market_dynamics_summary_en || selectedTech.NZErationale_en || selectedTech.NZErationale || selectedTech.market_dynamics_summary_zh)
                                : (selectedTech.market_dynamics_summary_zh || selectedTech.NZErationale_zh || selectedTech.NZErationale)}</p>
                            : <p className="text-sm text-slate-400 italic">{L.noMarket}</p>}
                        </div>
                      </div>
                    </div>

                    {/* Initiatives */}
                    <div className="bg-white p-5 md:p-6 rounded-xl shadow-sm border border-slate-200">
                      <div className="flex flex-wrap justify-between items-center mb-4 gap-3">
                        <h3 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
                          <Database size={18} className="text-emerald-500" /> {L.initiativesTitle}
                        </h3>
                        <span className="text-sm bg-slate-100 text-slate-600 px-2 py-1 rounded-full border border-slate-200">{L.totalCount(sortedInitiatives.length)}</span>
                      </div>
                      {sortedInitiatives.length === 0 ? (
                        <div className="text-center text-slate-500 py-8 bg-slate-50 rounded-lg border border-slate-100 border-dashed">{L.noInitiatives}</div>
                      ) : (
                        <div className="space-y-4">
                          {displayedInitiatives.map((init, idx) => (
                            <div key={idx} className="border-l-2 border-emerald-400 pl-4 py-1">
                              <div className="flex flex-wrap gap-2 mb-1 text-xs font-medium text-slate-500">
                                {init.year && <span className="bg-slate-100 px-1.5 py-0.5 rounded border border-slate-200">{init.year}</span>}
                                {init.country && <span className="bg-emerald-50 text-emerald-700 px-1.5 py-0.5 rounded border border-emerald-100">{init.country}</span>}
                                {init.type && <span className="text-blue-600">{init.type}</span>}
                              </div>
                              <p className="text-sm text-slate-700 leading-relaxed mt-1">
                                {init.description || L.unlabeled}
                                {init.read_more && <a href={init.read_more} target="_blank" rel="noopener noreferrer" className="text-blue-600 hover:text-blue-800 underline ml-2">{L.sourceLink}</a>}
                              </p>
                            </div>
                          ))}
                          {sortedInitiatives.length > 3 && (
                            <div className="text-center pt-3 border-t border-slate-100">
                              <button onClick={() => setShowAllInitiatives(!showAllInitiatives)} className="text-xs text-blue-600 hover:text-blue-800 font-medium px-4 py-1.5 bg-blue-50 hover:bg-blue-100 rounded transition-colors">
                                {showAllInitiatives ? L.collapseCases : L.expandCases(sortedInitiatives.length)}
                              </button>
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  </div>
                )}
              </div>
            )}

            {rightPanelTab === 'strategy' && (
              <div className="flex-1 flex flex-col overflow-y-auto p-4 md:p-8 w-full">
                <div className="max-w-2xl mx-auto w-full space-y-6">
                  <div>
                    <h2 className="text-lg font-bold text-slate-800">{L.strategyTitle}</h2>
                    <p className="text-sm text-slate-500 mt-1">{L.strategySubtitle}</p>
                  </div>

                  {/* Strategy overview card — opens instantly with a rule-based
                      summary; the option to upgrade to an AI-generated version
                      lives inside that same modal, so this is one progressive
                      entry point rather than two separate, overlapping ones. */}
                  <div className="bg-white p-6 md:p-8 rounded-2xl shadow-sm border border-slate-200">
                    <div className="flex items-start gap-4 mb-5">
                      <div className="bg-gradient-to-br from-purple-500 to-blue-500 text-white rounded-xl p-3 flex-shrink-0">
                        <Sparkles size={22} />
                      </div>
                      <div>
                        <h3 className="text-base font-bold text-slate-800">{L.strategyCardTitle}</h3>
                        <p className="text-sm text-slate-500 mt-1 leading-relaxed">
                          {L.strategyCardDescPrefix} <span className="font-semibold text-slate-700">{totalMatches}</span> {L.strategyCardDescSuffix}
                        </p>
                      </div>
                    </div>
                    <button
                      onClick={() => onOpenStrategyOverview(searchData)}
                      disabled={totalMatches === 0}
                      className="w-full sm:w-auto flex items-center justify-center gap-2 bg-gradient-to-r from-purple-600 to-blue-600 hover:from-purple-700 hover:to-blue-700 disabled:from-slate-300 disabled:to-slate-300 disabled:cursor-not-allowed text-white shadow-sm rounded-lg px-6 py-3 text-sm font-semibold transition-all"
                    >
                      <Sparkles size={18} /> {L.openStrategyBtn}
                    </button>
                  </div>

                  {/* Export card */}
                  <div className="bg-white p-6 md:p-8 rounded-2xl shadow-sm border border-slate-200">
                    <div className="flex items-start gap-4 mb-5">
                      <div className="bg-emerald-50 text-emerald-600 rounded-xl p-3 flex-shrink-0">
                        <Download size={22} />
                      </div>
                      <div>
                        <h3 className="text-base font-bold text-slate-800">{L.exportStrategyTitle}</h3>
                        <p className="text-sm text-slate-500 mt-1 leading-relaxed">
                          {L.exportStrategyDescPrefix} <span className="font-semibold text-slate-700">{totalMatches}</span> {L.exportStrategyDescSuffix}
                        </p>
                      </div>
                    </div>
                    <button
                      onClick={handleExportListCsv}
                      disabled={totalMatches === 0}
                      className="w-full sm:w-auto flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-300 disabled:cursor-not-allowed text-white shadow-sm rounded-lg px-6 py-3 text-sm font-semibold transition-colors"
                    >
                      <Download size={18} /> {L.exportAll}
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>
        </section>
      </main>

      <footer className="bg-slate-800 text-slate-400 text-xs py-3 px-6 text-center z-30 flex-shrink-0">
        {L.footerDisclaimer}
      </footer>
    </div>
  );
}
