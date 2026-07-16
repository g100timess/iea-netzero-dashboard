import React, { useState, useEffect, useMemo, useRef } from 'react';
import { Search, FileJson, AlertCircle, ChevronLeft, ChevronRight, Database, FileText, Info, CheckCircle2, Download, Sparkles, X, Loader2, KeyRound, Eye, EyeOff, RotateCcw, Copy, ExternalLink, ChevronDown, ChevronUp, Globe, HelpCircle, Presentation } from 'lucide-react';
import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType } from 'docx';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas-pro';

// ─── Constants ───────────────────────────────────────────────────────────────

const DATA_URL = "https://g100timess.github.io/iea-netzero-dashboard/rebuilt_dataset_zh_en.json";

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
  "Fossil fuels": ["Coal", "Oil and Gas"],
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
  "Renewables": "再生能源", "Hydropower": "水力發電",
  "Ocean": "海洋能", "Wind": "風力", "Bioenergy": "生質能", "Geothermal": "地熱",
  "Solar": "太陽能", "Nuclear": "核能", "Fission": "核分裂", "Fusion": "核融合",
  "Hydrogen": "氫能", "Sythetic fuels production": "合成燃料生產", "H2 infrastructure": "氫能基礎設施",
  "Power generation": "發電", "Production": "生產", "Energy networks and storage": "能源網路與儲能",
  "Thermal storage for buildings and district heating": "建築與區域供熱熱能儲存",
  "Mechanical storage": "機械儲能", "Electrochemical storage": "電化學儲能",
  "Thermal storage": "熱能儲存", "Physical grids": "實體電網", "Smart grids": "智慧電網",
  "Carbon Capture and Storage": "碳捕捉與封存", "CO2 capture": "二氧化碳捕捉",
  "CO2 storage": "二氧化碳封存", "CO2 transport": "二氧化碳運輸",
  "Critical minerals": "關鍵礦物", "Mineral processing and refining": "礦物加工與精煉",
  "Mining and extraction": "採礦與開採", "E-waste recycling": "電子廢棄物回收",
  "Combustion and conversion": "燃燒與轉化", "Components": "航空器零組件",
  "Other components": "其他零組件", "Thermal processes": "熱製程",
  "Nuclear fuel cycles": "核燃料循環", "EV charging equipment": "電動車充電設備",
  "Building-level thermal storage": "建築端熱能儲存",
  "Ammonia": "氨", "Battery recycling": "電池回收",
  "Biogases and biomethane production": "沼氣與生質甲烷生產",
  "Building envelope materials": "建築外殼材料", "Building envelope technologies": "建築外殼技術",
  "Cross-cutting recycling": "跨領域回收技術",
  "Electric vehicle battery (electrical storage)": "電動車電池（電能儲存）",
  "Electrolysis": "電解", "H2 storage": "氫氣儲存",
  "Heating, cooling and ventilation technologies": "供暖、冷房與通風技術",
  "High temperature": "高溫",
  "Hydrogen use in road transport": "氫能於公路運輸之應用",
  "Liquid biofuels production": "液態生質燃料生產",
  "Low to medium temperature": "中低溫",
  "Oil and Gas": "石油與天然氣", "Power generation from biomass": "生質能發電",
  "Propulsion": "推進系統",
  "Subsurface operations for geothermal energy": "地熱能地下作業",
  "digital and tools": "數位技術與工具", "large-scale or industrial": "大型或工業規模技術",
  "mass-manufactured": "量產型技術"
};

const SEARCH_ALIASES = {
  '氫能': ['氫能', 'hydrogen', 'h2'], 'hydrogen': ['氫能', 'hydrogen', 'h2'], 'h2': ['氫能', 'hydrogen', 'h2'],
  '太陽能': ['太陽能', '太陽光電', 'solar', 'solar pv', 'photovoltaic', 'pv'],
  '太陽光電': ['太陽能', '太陽光電', 'solar', 'solar pv', 'photovoltaic', 'pv'],
  'solar': ['太陽能', '太陽光電', 'solar', 'solar pv', 'photovoltaic', 'pv'],
  '地熱': ['地熱', 'geothermal'], 'geothermal': ['地熱', 'geothermal'],
  '碳捕捉': ['碳捕捉', 'ccs', 'ccus'], 'ccs': ['碳捕捉', 'ccs', 'ccus'], 'ccus': ['碳捕捉', 'ccs', 'ccus'],
};

// Chinese label for one segment of a technology's sector[] breadcrumb path.
// The leaf (last) segment is tech-specific — tech.sector_zh is authoritative
// there, not a shared lookup table, since sector_en/sector_zh are computed
// per-technology. Every other position (the manufacturing-scale tag at index
// 0, and all intermediate category segments) comes from SECTOR_TRANSLATIONS,
// which the id/id_zh audit confirmed has zero gaps against the real dataset.
function sectorSegmentLabel(tech, index) {
  const segments = Array.isArray(tech.sector) ? tech.sector : [];
  const seg = segments[index];
  if (seg === undefined) return '';
  if (index === segments.length - 1) return tech.sector_zh || seg;
  return SECTOR_TRANSLATIONS[seg] || seg;
}

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

// Compact "TRL X-X" form (no 原型研發/技術示範/... prefix) for the one-line
// sector summary sentence under the donut chart — same 4 buckets/order as
// TRL_STAGE_LABELS above, just terser wording for an inline sentence.
const TRL_COMPACT_LABELS_ZH = ['TRL 1-3', 'TRL 4-6', 'TRL 7-8', 'TRL 9+'];

// adoption_stage is a market/commercial-diffusion axis, distinct from TRL
// (technical maturity) — restored from the original raw IEA source (it had
// been dropped from this dataset's rebuild). Only 3 raw English values ever
// appear across all 639 technologies; empty string means IEA didn't rate
// that technology's adoption stage, not a data error, so the UI hides the
// badge entirely rather than showing an "unrated" placeholder.
const ADOPTION_STAGE_LABELS_ZH = {
  'Awaiting adoption': '尚待導入',
  'Building momentum': '應用擴大中',
  'Commercial in many markets': '廣泛商業化',
};
function adoptionStageLabel(stage, uiLang) {
  if (!stage) return '';
  if (uiLang === 'en') return stage;
  return ADOPTION_STAGE_LABELS_ZH[stage] || stage;
}

// One-line, hand-written definitions for every sector_zh value in the
// cached dataset (132 total) — keyed by the exact sector_zh string. Used
// only to add a short "what is this" clause to the sector summary sentence
// when a donut segment is locked; not proofread by a domain expert, so
// treat as a first pass and flag anything that reads wrong. English mode
// intentionally has no equivalent map yet (skips the clause rather than
// guessing a translation) — the definitions were reviewed in Chinese only.
const SECTOR_DEFINITIONS = {
  '其他生產技術': '未歸入其他既有分類的生產方法。',
  '能源來源轉換與電氣化': '以較潔淨能源及電力取代既有燃料。',
  '牆體、屋頂與建築立面': '構成建築物非透明外殼的外部構件。',
  '燃燒與轉化': '將燃料轉化為熱能或其他能源的製程。',
  '生質柴油與生質航空燃油': '以生質原料製成的柴油與航空燃油。',
  '回收再利用': '回收並重新處理材料以供再次利用。',
  '電力加熱': '將電能轉換為熱能的加熱方式。',
  '採用 CCUS 技術': '採用碳捕捉、再利用與封存技術。',
  '採用碳捕捉技術': '配備捕捉碳排放的相關技術。',
  '燃料供熱': '透過燃料燃燒產生熱能的供熱方式。',
  '加工': '將材料處理成可用產品或中間產物。',
  '建築能源管理系統': '監測並最佳化建築能源使用的系統。',
  '甲烷排放監測與減量': '偵測、量測並減少甲烷排放的技術。',
  '生質甲烷': '以生質原料製成的富甲烷氣體。',
  '太陽光電': '將太陽光直接轉換為電力的技術。',
  '直接還原鐵（DRI）': '在不熔融礦石下還原製成的鐵。',
  '電動車': '完全或部分以電力驅動的車輛。',
  '其他零組件': '未歸入其他既有分類的零組件。',
  '磁約束核融合': '利用磁場約束高溫電漿的核融合技術。',
  '潛熱（相變材料）': '材料相變時吸收或釋放的熱能。',
  '容量管理': '依需求變化管理系統可用容量。',
  '其他熱泵技術': '未歸入其他既有分類的熱泵技術。',
  '鋰基電池': '電化學系統中使用鋰的電池。',
  '空氣源熱泵': '與室外空氣交換熱能的熱泵。',
  '鐵路': '利用鐵路運送旅客或貨物。',
  '海洋能': '從波浪、潮汐、海流或海洋熱能取得的能源。',
  '地熱鑽井': '鑽設井孔以探勘及取得地熱資源。',
  '小型模組化反應爐': '單機容量不超過 300 MWe 的模組化核子反應爐。',
  '氫氣輸配': '將氫氣輸送及配送至終端用戶。',
  '區域供熱熱能儲存': '整合於區域供熱管網的熱能儲存系統。',
  '機械儲能': '以機械能或位能形式儲存能量。',
  '建築端熱能儲存': '設置於建築端的熱能儲存系統。',
  '控制與監測系統': '控制作業並監測設備效能的系統。',
  '濕法冶金': '利用水溶液提取及回收金屬的製程。',
  '高價值化學品': '具較高經濟價值或特殊用途的化學品。',
  '甲醇': '用作化工原料、燃料或能源載體的化學品。',
  '採用碳捕捉與利用技術': '捕捉二氧化碳並將其作為原料利用。',
  '照明技術與控制系統': '高效率照明設備及其運轉控制系統。',
  '減少材料損耗': '減少材料在生產及使用過程中的耗損。',
  '固態冷卻設備': '利用固態材料熱效應進行冷卻的設備。',
  '航空器': '用於空中飛行的運具。',
  '航空器零組件': '構成航空器的零件與子系統。',
  '其他電池技術': '未歸入其他電池分類的技術。',
  '氫燃料電池電動車（FCEV）': '以氫燃料電池供電的電動車。',
  '先進燃燒引擎與動力系統': '提升效率並降低排放的引擎與動力系統。',
  '風力': '將風的動能轉換而取得的能源。',
  '生質乙醇': '以生質原料製成的乙醇燃料。',
  '地熱能地面作業': '利用地面設備處理地熱流體並產生能源的作業。',
  '地下儲氫': '將大量氫氣儲存於地下地質構造。',
  '氫氣直接利用': '直接將氫氣用作燃料或工業原料。',
  '熱裂解': '在缺氧或無氧條件下加熱分解材料。',
  '電化學儲能': '利用可逆電化學反應儲存能量。',
  '電纜回收': '從廢電纜回收金屬及其他材料。',
  '製造': '生產材料、零組件或成品的作業。',
  '替代製程路徑': '採用替代製程或原料的生產路徑。',
  '熱催化': '結合熱能與觸媒進行化學轉化。',
  '資源循環（將廢棄物轉化為化學品與生質能）': '將廢棄物轉化為化學品、燃料或生質能。',
  '門窗系統': '建築物的窗戶、玻璃門及天窗等開口系統。',
  '熱驅動式熱泵': '主要以熱能驅動的熱泵。',
  '氫動力航空器零組件': '為氫動力航空器設計的零組件。',
  '氫燃料車輛': '以燃料電池或燃燒方式使用氫氣驅動的車輛。',
  '充電與燃料加注': '為運具補充電力或燃料。',
  '煉製、運輸與儲存': '煉製、運輸及儲存燃料或能源載體。',
  '太陽熱能': '將太陽輻射轉換為可利用熱能的技術。',
  '核燃料循環前端': '核燃料進入反應爐前的製備階段。',
  '替代方案': '不同於傳統技術的替代方法。',
  '液態燃料': '用於產生熱能或動力的液態物質。',
  '氫氣調質製程': '調整氫氣純度、壓力或狀態以利輸儲及使用。',
  '氫基燃料應用': '使用以氫氣製成的氨或合成碳氫燃料。',
  '電解槽設計': '以電力製氫之電解槽與系統設計。',
  '其他': '未歸入其他既有分類的技術。',
  '電纜與導體': '用於輸送及配電的電纜與導電元件。',
  '其他封存技術': '用於監測或將二氧化碳封存於地質構造的其他技術。',
  '前處理與拆解': '在材料回收前進行產品處理及拆解。',
  '生質燃料': '直接或間接以生質原料製成的燃料。',
  '其他能源': '未歸入其他既有分類的能源來源。',
  '建築設計工具': '用於設計及評估建築效能的工具。',
  '輕量化': '在維持必要性能下減輕產品重量。',
  '牆面、屋頂與建築立面太陽熱能技術': '整合於牆面、屋頂或建築立面的太陽熱能系統。',
  '空調設備': '調節室內溫度與空氣狀態的設備。',
  '蒸發冷卻': '利用水分蒸發吸熱的冷卻方式。',
  '地源與水源熱泵': '與土壤、地下水或水體交換熱能的熱泵。',
  '其他建築冷暖技術': '未歸入其他分類的建築供暖與冷房技術。',
  '多聯供系統': '整合生產多種能源產品的系統。',
  '次世代金屬電池': '採用新興金屬電化學體系的先進電池。',
  '電動車充電設備': '為電動車供應電力的充電設備。',
  '加氫與基礎設施': '為車輛加注氫氣及其相關基礎設施。',
  '引擎': '將能源轉換為機械動力的設備。',
  '船舶營運': '船舶航行、操作及維護的相關作業。',
  '搭配二氧化碳捕捉': '配備捕捉二氧化碳排放的技術。',
  '次世代地熱能源系統': '突破傳統資源限制以取得地熱的先進系統。',
  '大型反應爐': '電力容量超過 700 MW 的核子反應爐。',
  '核融合': '藉由輕原子核融合反應產生能量。',
  '慣性約束核融合': '藉由快速壓縮燃料靶丸實現核融合。',
  '合成燃料生產': '透過化學合成製造燃料。',
  '地上儲存': '將能源載體儲存於地面設施。',
  '氫氣海運': '以船舶運輸氫氣或氫載體。',
  '地質氫': '從天然聚積中開採或於地質構造內生成的氫氣。',
  '潛熱儲存': '利用材料相變儲存及釋放熱能。',
  '顯熱儲存': '藉由改變材料溫度來儲存熱能。',
  '熱化學儲熱': '利用可逆化學反應儲存及釋放熱能。',
  '變壓與電力轉換': '執行電壓變換、電力轉換及電力潮流控制的技術。',
  '直接空氣捕捉': '直接從環境空氣中捕捉二氧化碳。',
  '礦物封存': '將二氧化碳轉化為穩定碳酸鹽礦物的永久封存方式。',
  '二氧化碳運輸': '將捕捉的二氧化碳運往利用或封存地點。',
  '採礦與開採': '從地層中取得礦產資源的作業。',
  '生物技術': '利用生物、細胞或酵素進行的技術。',
  '苯、甲苯及二甲苯': '用作燃料成分及化工原料的芳香族烴。',
  '乙烯': '主要用於製造塑膠的基礎化學品。',
  '電力': '藉由電荷流動傳輸的能源。',
  '建築生質能供暖': '利用生質能供應建築物熱能。',
  '電池交換': '將低電量電池更換為已充電電池。',
  '需量反應': '依電網訊號調整用電量或用電時間。',
  '動態充電或電動道路系統': '在車輛行駛期間供應電力的道路系統。',
  '燃料電池電動車用氫燃料電池': '專為電動車供電設計的氫燃料電池。',
  '氫基技術': '生產、處理或使用氫氣的相關技術。',
  '清潔燃料船用內燃機': '使用低排放或清潔燃料的船用內燃機。',
  '燃料電池電動船舶': '以燃料電池發電並由電力推進的船舶。',
  '船舶': '用於水上運送旅客或貨物的運具。',
  '天然氣發電（搭配 CCUS）': '採用碳捕捉、再利用與封存的天然氣發電。',
  '油氣探勘與生產': '探尋、評估及開採石油與天然氣的作業。',
  '核燃料循環': '從核燃料製備至最終廢棄物管理的完整流程。',
  '核燃料循環後端': '核燃料使用後的管理階段。',
  '材料型儲氫': '利用氫化物、吸附材料或化學載體儲存氫氣。',
  '新型原料': '用於生產的新興或替代性原料。',
  '熱化學水分解': '利用高溫及化學反應分解水以製取氫氣。',
  '其他精煉方法': '未歸入其他既有分類的精煉方法。',
  '其他供熱方式': '未歸入其他既有分類的供熱方式。',
  '航空營運': '航空器運航及地面移動的相關作業。',
  '其他車輛': '未歸入其他既有分類的車輛。',
  '水力發電': '利用水流能量產生電力。',
  '熱製程': '主要利用熱能進行的工業製程。'
};

// English-mode counterpart to SECTOR_DEFINITIONS above — keyed by the exact
// sector_en string (not sector_zh), since the locked-segment summary looks
// up by whichever field the current uiLang displays. Being filled in
// gradually in batches (unlike SECTOR_DEFINITIONS, which was written in one
// pass) — sector_zh values with no entry here yet simply show no definition
// clause in English mode, same graceful-omission behavior as an unrated
// adoption_stage. IMPORTANT: keys must match the dataset's actual sector_en
// spelling, including its existing typos (e.g. "powetrains", not
// "powertrains") — a corrected spelling here would silently never match.
const SECTOR_DEFINITIONS_EN = {
  'Aboveground storage': 'Storage of energy carriers in aboveground facilities.',
  'Advanced combustion engines and powetrains': 'Efficient, low-emission engine and powertrain technologies.',
  'Aircrafts': 'Vehicles designed for flight.',
  'Alternative approaches': 'Methods different from conventional technical solutions.',
  'Alternative routes': 'Production pathways using alternative processes or inputs.',
  'Battery swapping': 'Replacing a depleted battery with a charged one.',
  'Benzene, toluene and xylenes': 'Aromatic hydrocarbons used as fuels and chemical feedstocks.',
  'Bioethanol': 'Ethanol fuel produced from biomass.',
  'Biofuels': 'Fuels produced directly or indirectly from biomass.',
  'Biological': 'Processes using organisms, cells or enzymes.',
  'Biomethane': 'Methane-rich gas produced from biomass.',
  'Building design tools': 'Tools for designing and evaluating building performance.',
  'Building energy management systems': 'Systems that monitor and optimise building energy use.',
  'Building heating from bioenergy': 'Building heating supplied by bioenergy.',
  'Building-level thermal storage': 'Thermal energy storage installed within a building.',
  'Air conditioners': 'Equipment used to cool and condition indoor air.',
  'Back-end of the fuel cycle': 'Management of spent nuclear fuel after reactor use.',
  'Biodiesel and biokerosene': 'Diesel and aviation fuels produced from biomass.',
  'CCUS-enabled': 'Equipped with carbon capture, utilisation and storage.',
  'CO2 transport': 'Transport of captured CO2 to utilisation or storage sites.',
  'Cable recycling': 'Recovery of metals and materials from discarded cables.',
  'Cables and conductors': 'Components used to transmit and distribute electricity.',
  'Capacity management': 'Management of system capacity to meet changing demand.',
  'Clean-fuel internal combustion ship engines': 'Ship engines powered by low-emission or clean fuels.',
  'Combustion and conversion': 'Processes that convert fuels into heat or other energy forms.',
  'Components': 'Parts and subsystems used in aircraft.',
  'Charging and refuelling': 'Supplying electricity or fuel to vehicles.',
  'Control systems and monitoring': 'Systems that control operations and monitor performance.',
  'Conditioning processes': 'Processes that prepare hydrogen for transport, storage or use.',
  'Demand response': 'Adjustment of electricity demand in response to grid signals.',
  'Direct air capture': 'Capture of carbon dioxide directly from ambient air.',
  'Direct reduced iron (DRI)': 'Iron produced by reducing iron ore without melting it.',
  'Direct use of hydrogen': 'Direct use of hydrogen as a fuel or feedstock.',
  'District heating thermal storage': 'Thermal storage integrated with a district heating network.',
  'Dynamic charging or electric road system': 'Road systems that supply electricity to moving vehicles.',
  'EV charging equipment': 'Equipment that supplies electricity to electric vehicles.',
  'Electric vehicle (EV)': 'A vehicle powered partly or entirely by electricity.',
  'Electrical heating': 'Production of heat using electrical energy.',
  'Drilling': 'Drilling wells to explore and access geothermal resources.',
  'Electricity': 'Energy transferred through the flow of electric charge.',
  'Electrochemical storage': 'Energy storage based on reversible electrochemical reactions.',
  'Electrolyser design': 'Design of systems that use electricity to produce hydrogen.',
  'Engines': 'Machines that convert energy into mechanical power.',
  'Ethylene': 'A basic chemical used mainly to produce plastics.',
  'Evaporative cooling': 'Cooling achieved through the evaporation of water.',
  'Fuel cell electric ship': 'A ship propelled by electricity generated from fuel cells.',
  'Exploration and production': 'Activities for locating, assessing and extracting oil and natural gas.',
  'Fenestration': 'Building openings such as windows, glazed doors and skylights.',
  'Front-end of the fuel cycle': 'Processes that prepare nuclear fuel for reactor use.',
  'Fuel-based heating': 'Heating produced through the combustion of fuels.',
  'Fusion': 'Energy production by combining light atomic nuclei.',
  'Geologic hydrogen': 'Hydrogen extracted from natural accumulations or generated within geological formations.',
  'Ground-source and water-source heat pumps': 'Heat pumps that exchange heat with the ground or water.',
  'H2 shipping': 'Maritime transport of hydrogen or hydrogen carriers.',
  'H2 transmission and distribution': 'Transmission and distribution of hydrogen to end users.',
  'High value chemicals': 'Chemicals with high economic value or specialised uses.',
  'Hydrogen based technologies': 'Technologies that produce, handle or use hydrogen.',
  'Hydrogen fuel cell electric vehicles (FCEVs)': 'Electric vehicles powered by hydrogen fuel cells.',
  'Hydrogen fuel cell for FCEVs': 'Hydrogen fuel cells designed to power electric vehicles.',
  'Hydrogen-powered aircraft components': 'Aircraft components designed for hydrogen-powered flight.',
  'Hydrometallurgy': 'Metal extraction and recovery using aqueous solutions.',
  'Hydropower': 'Electricity generation using the energy of flowing water.',
  'Inertial confinement fusion': 'Fusion achieved by rapidly compressing a fuel target.',
  'Large reactors': 'Nuclear reactors with an electrical capacity above 700 MW.',
  'Latent heat (phase change material)': 'Heat absorbed or released during a material phase change.',
  'Latent heat storage': 'Thermal storage using the phase change of a material.',
  'Lighting technologies and control systems': 'Efficient lighting equipment and systems that control its operation.',
  'Lightweighting': 'Reducing product weight while maintaining required performance.',
  'Liquid fuels': 'Liquid materials used to produce heat or power.',
  'Lithium based batteries': 'Batteries that use lithium in their electrochemical system.',
  'Magnetic confinement fusion': 'Fusion using magnetic fields to confine hot plasma.',
  'Manufacturing': 'Production of materials, components or finished products.',
  'Materials-based storage': 'Hydrogen storage using hydrides, sorbents or chemical carriers.',
  'Mechanical storage': 'Energy storage using mechanical or potential energy.',
  'Methanol': 'A chemical used as a feedstock, fuel or energy carrier.',
  'Mineral storage': 'Permanent CO2 storage through formation of stable carbonate minerals.',
  'Mining and extraction': 'Activities that remove mineral resources from the earth.',
  'Next-generation geothermal energy systems': 'Advanced systems that access geothermal heat beyond conventional resources.',
  'Next-generation metal batteries': 'Advanced batteries using emerging metal-based chemistries.',
  'Novel feedstock': 'New or alternative raw materials used in production.',
  'Nuclear fuel cycles': 'Processes from nuclear fuel preparation to final waste management.',
  'Ocean': 'Energy obtained from waves, tides, currents or ocean heat.',
  'Other': 'Technologies not included in another defined category.',
  'Other batteries technologies': 'Battery technologies not included in other battery categories.',
  'Other building heating and cooling technologies': 'Building heating and cooling technologies not otherwise classified.',
  'Other components': 'Components not included in another defined category.',
  'Other heat pumps technologies': 'Heat pump technologies not included in another category.',
  'Other heating': 'Heating methods not included in another category.',
  'Other production techniques': 'Production methods not included in another category.',
  'Other refining methods': 'Refining methods not included in another category.',
  'Other sources': 'Energy sources not included in another category.',
  'Other storage technologies': 'Additional technologies for monitoring or storing CO2 in geological formations.',
  'Other vehicles': 'Vehicle types not included in another category.',
  'Photovoltaic': 'Technology that converts sunlight directly into electricity.',
  'Polygeneration systems': 'Integrated systems that produce multiple energy products.',
  'Power generation from natural gas, with CCUS': 'Natural gas power generation equipped with CCUS.',
  'Pre-processing and disassembly': 'Preparation and dismantling of products before material recovery.',
  'Processing': 'Treatment of materials into usable products or intermediates.',
  'Pyrolysis': 'Thermal decomposition of materials with little or no oxygen.',
  'Rail': 'Transport of passengers or freight by rail.',
  'Recycling': 'Recovery and reprocessing of materials for further use.',
  'Recycling (waste product conversion to chemicals and bioenergy)': 'Conversion of waste into chemicals, fuels or bioenergy.',
  'Reducing material losses': 'Measures that minimise material waste during production and use.',
  'Refining, transport and storage': 'Processing, transporting and storing fuels or energy carriers.',
  'Refueling and infrastructure': 'Hydrogen refuelling systems and their supporting infrastructure.',
  'Sensible heat storage': 'Thermal storage achieved by changing a material’s temperature.',
  'Shift in energy sources and electrification': 'Replacement of existing fuels through cleaner energy sources and electricity.',
  'Small modular reactors': 'Factory-built modular reactors with capacity up to 300 MWe.',
  'Solar thermal': 'Technology that converts solar radiation into usable heat.',
  'Solar thermal technologies for wall, roof & façades': 'Solar thermal systems integrated into building walls, roofs or façades.',
  'Surface operations for geothermal energy': 'Surface processes and equipment used to produce geothermal energy.',
  'Sythetic fuels production': 'Production of fuels through chemical synthesis.',
  'Thermal processes': 'Industrial processes driven primarily by heat.',
  'Thermally-driven heat pump': 'A heat pump driven mainly by thermal energy.',
  'Thermocatalytic': 'Chemical conversion using heat and a catalyst.',
  'Thermochemical heat storage': 'Thermal storage using reversible chemical reactions.',
  'Thermochemical water splitting': 'Hydrogen production by splitting water with heat and chemical reactions.',
  'Transformation and conversion': 'Technologies that transform voltage, convert power and control electricity flows.',
  'Use of hydrogen-based fuels': 'Use of ammonia or synthetic hydrocarbons produced from hydrogen.',
  'Vessel operations': 'Activities involved in operating and maintaining ships.',
  'Vessels': 'Waterborne vehicles used to transport passengers or cargo.',
  'Wall, roof & façade': 'External building components forming the opaque envelope.',
  'Wind': 'Energy obtained by converting the motion of wind.',
  'With CO2 capture': 'Equipped to capture carbon dioxide emissions.',
  'With carbon capture': 'Equipped with technology that captures carbon emissions.',
  'With carbon capture and usage': 'Equipped to capture carbon dioxide and use it as an input.',
  'Underground storage': 'Large-scale hydrogen storage in underground geological formations.',
  'Solid-state equipment cooling': 'Cooling equipment based on solid-state caloric effects.',
  'Operations': 'Activities involved in aircraft operation and ground movement.',
  'Methane emissions monitoring and abatement': 'Technologies that detect, measure and reduce methane emissions.',
  'Hydrogen-fuelled vehicules': 'Vehicles powered by hydrogen through fuel cells or combustion.',
  'Air-source heat pumps': 'Heat pumps that exchange heat with outdoor air.',
};

// One-line definitions for intermediate breadcrumb keywords (path segments
// above the leaf sector, e.g. "Industry > Aluminium > ..." — "Aluminium" is
// intermediate here, not a sector_zh leaf). Keyed by the English keyword
// (not zh) because these terms come from SECTOR_TRANSLATIONS, which is also
// English-keyed, and don't carry the "always-unique leaf" guarantee that
// sector_zh has. Not yet wired into any UI surface.
const SECTOR_PATH_DEFINITIONS_ZH = {
  'Aluminium': '鋁及其合金的生產與加工。',
  'Ammonia': '用作化工原料、肥料及能源載體的化學品。',
  'Aviation': '使用航空器運送旅客或貨物的運輸活動。',
  'Battery recycling': '從生產廢料及廢電池中回收可用材料。',
  'Bioenergy': '由生質原料及有機廢棄物產生的能源。',
  'Biogases and biomethane production': '利用有機原料生產沼氣及生質甲烷。',
  'Building envelope materials': '控制建築外殼熱量、空氣及濕氣傳遞的材料。',
  'Building envelope technologies': '改善建築外殼能源效能的技術。',
  'Buildings': '涵蓋住宅及服務業建築能源使用與技術的部門。',
  'CO2 capture': '從排放源或環境空氣中分離二氧化碳。',
  'CO2 storage': '將捕捉的二氧化碳永久封存於適當地質構造。',
  'Carbon Capture and Storage': '捕捉、運輸並永久封存二氧化碳的技術鏈。',
  'Cement and concrete': '水泥、混凝土及相關營建材料的生產。',
  'Chemicals and plastics': '化學品、高分子材料及塑膠產品的生產。',
  'Coal': '從地質礦床開採的富碳化石燃料。',
  'Cooking technologies': '為食物烹調提供能源的設備與系統。',
  'Critical minerals': '能源技術不可或缺且具有供應風險的礦產資源。',
  'Cross-cutting recycling': '可應用於多種產品與材料的回收技術。',
  'Design and envelope': '降低建築能源需求的設計與外殼措施。',
  'E-waste recycling': '從廢棄電氣與電子設備中回收可用材料。',
  'Electric vehicle battery (electrical storage)': '儲存電力並為電動車提供動力的充電式電池。',
  'Electrolysis': '利用電力驅動化學反應的製程。',
  'Energy networks and storage': '輸送、管理及儲存能源的基礎設施。',
  'Fission': '重原子核分裂並釋放能量的反應。',
  'Fossil fuels': '由古代有機物形成的煤炭、石油及天然氣。',
  'Geothermal': '從地球內部熱能取得的能源。',
  'H2 infrastructure': '用於氫氣運輸、配送及供應的基礎設施。',
  'H2 storage': '將氫氣保存至後續使用的儲存技術。',
  'Heating, cooling and ventilation technologies': '調節室內溫度、通風及空氣品質的技術。',
  'High temperature': '需要高溫熱能的工業製程。',
  'Hydrogen': '涵蓋氫氣生產、基礎設施及應用的能源領域。',
  'Hydrogen use in road transport': '使用氫氣為公路車輛提供動力。',
  'Industrial heating': '為工業生產製程供應熱能。',
  'Industry': '製造材料、中間產品及成品的產業部門。',
  'Iron and steel': '鐵與鋼材的生產及加工。',
  'Liquid biofuels production': '利用生質原料或廢棄物生產液態燃料。',
  'Low to medium temperature': '需要中低溫熱能的工業製程。',
  'Metallic products': '使用金屬製造產品及零組件。',
  'Mineral processing and refining': '將礦石加工成純化礦物、金屬或材料。',
  'Nuclear': '透過受控核反應產生的能源。',
  'Oil and Gas': '涵蓋石油與天然氣探勘、生產及處理的產業。',
  'Operations and equipment': '石油與天然氣設施的營運作業及設備。',
  'Physical grids': '用於輸電及配電的實體基礎設施。',
  'Power generation': '使用氫氣或氫基燃料產生電力。',
  'Power generation from biomass': '使用生質原料或有機廢棄物產生電力。',
  'Production': '利用水、燃料或其他資源生產氫氣的製程。',
  'Propulsion': '產生動力或推力以驅動船舶的系統。',
  'Pulp and paper': '紙漿、紙張及相關產品的生產。',
  'Renewables': '來自可自然持續補充來源的能源。',
  'Road transport': '使用道路車輛運送旅客或貨物。',
  'Shipping': '使用船舶經海洋或內陸水道運送旅客或貨物。',
  'Smart grids': '運用數位監測、通訊與控制技術管理電力的電網。',
  'Solar': '從太陽輻射取得的能源。',
  'Subsurface operations for geothermal energy': '探勘、開發及管理地下地熱資源的作業。',
  'Transport': '涵蓋旅客與貨物運送的部門。',
  'digital and tools': '支援能源系統運作的數位技術與分析工具。',
  'large-scale or industrial': '為大容量或工業應用所設計的技術。',
  'mass-manufactured': '以標準化單元進行大量製造的技術。',
  'Thermal storage': '以熱或冷的形式儲存能量，供後續使用。',
  'Thermal storage for buildings and district heating': '為建築或區域供熱管網提供的熱能儲存技術。',
  'Building-level thermal storage': '設置於建築端的熱能儲存系統。',
  'Combustion and conversion': '將燃料轉化為熱能或其他能源的製程。',
  'Components': '構成航空器的零件與子系統。',
  'EV charging equipment': '為電動車供應電力的充電設備。',
  'Nuclear fuel cycles': '從核燃料製備至最終廢棄物管理的完整流程。',
  'Other components': '未歸入其他既有分類的零組件。',
  'Thermal processes': '主要利用熱能進行的工業製程。',
  'Fusion': '藉由輕原子核融合反應產生能量。',
  'Sythetic fuels production': '透過化學合成製造燃料。',
};

const SECTOR_PATH_DEFINITIONS_EN = {
  'Aluminium': 'Production and processing of aluminium and its alloys.',
  'Ammonia': 'A chemical used as a feedstock, fertiliser and energy carrier.',
  'Aviation': 'Transport of passengers or cargo by aircraft.',
  'Battery recycling': 'Recovery of materials from production scrap and used batteries.',
  'Bioenergy': 'Energy produced from biomass and organic waste.',
  'Biogases and biomethane production': 'Production of biogases and biomethane from organic feedstocks.',
  'Building envelope materials': 'Materials that control heat, air and moisture through building envelopes.',
  'Building envelope technologies': 'Technologies that improve building-envelope energy performance.',
  'Buildings': 'The sector covering energy use and technologies in residential and service buildings.',
  'CO2 capture': 'Separation of CO2 from emission sources or ambient air.',
  'CO2 storage': 'Permanent storage of captured CO2 in suitable geological formations.',
  'Carbon Capture and Storage': 'Capture, transport and permanent storage of carbon dioxide.',
  'Cement and concrete': 'Production of cement, concrete and related construction materials.',
  'Chemicals and plastics': 'Production of chemicals, polymers and plastic products.',
  'Coal': 'A carbon-rich fossil fuel extracted from geological deposits.',
  'Cooking technologies': 'Equipment and systems that provide energy for cooking.',
  'Critical minerals': 'Minerals essential to energy technologies and exposed to supply risks.',
  'Cross-cutting recycling': 'Recycling technologies applicable across multiple products and materials.',
  'Design and envelope': 'Building design and envelope measures that reduce energy demand.',
  'E-waste recycling': 'Recovery of materials from discarded electrical and electronic equipment.',
  'Electric vehicle battery (electrical storage)': 'Rechargeable batteries that store electricity to power electric vehicles.',
  'Electrolysis': 'A process that uses electricity to drive a chemical reaction.',
  'Energy networks and storage': 'Infrastructure that transports, manages and stores energy.',
  'Fission': 'Splitting of heavy atomic nuclei to release energy.',
  'Fossil fuels': 'Coal, oil and natural gas formed from ancient organic matter.',
  'Geothermal': 'Energy obtained from heat within the Earth.',
  'H2 infrastructure': 'Infrastructure for transporting, distributing and supplying hydrogen.',
  'H2 storage': 'Technologies that store hydrogen for later use.',
  'Heating, cooling and ventilation technologies': 'Technologies that control indoor temperature, ventilation and air quality.',
  'High temperature': 'Industrial processes that require high-temperature heat.',
  'Hydrogen': 'The energy field covering hydrogen production, infrastructure and use.',
  'Hydrogen use in road transport': 'Use of hydrogen to power road vehicles.',
  'Industrial heating': 'Supply of thermal energy for industrial production processes.',
  'Industry': 'The sector that manufactures materials and finished products.',
  'Iron and steel': 'Production and processing of iron and steel.',
  'Liquid biofuels production': 'Production of liquid fuels from biomass or waste.',
  'Low to medium temperature': 'Industrial processes requiring low- to medium-temperature heat.',
  'Metallic products': 'Manufacture of products and components from metals.',
  'Mineral processing and refining': 'Processing ores into purified minerals, metals or materials.',
  'Nuclear': 'Energy produced through controlled nuclear reactions.',
  'Oil and Gas': 'The sector covering exploration, production and processing of oil and gas.',
  'Operations and equipment': 'Operations and equipment used in oil and gas facilities.',
  'Physical grids': 'Physical infrastructure for electricity transmission and distribution.',
  'Power generation': 'Electricity generation using hydrogen or hydrogen-based fuels.',
  'Power generation from biomass': 'Electricity generation using biomass or organic waste.',
  'Production': 'Processes that produce hydrogen from water, fuels or other resources.',
  'Propulsion': 'Systems that generate power or thrust to propel a vessel.',
  'Pulp and paper': 'Production of pulp, paper and related products.',
  'Renewables': 'Energy from naturally replenished sources.',
  'Road transport': 'Transport by road vehicles for passengers or freight.',
  'Shipping': 'Transport of passengers or goods by sea or inland waterways.',
  'Smart grids': 'Electricity networks using digital monitoring, communication and control.',
  'Solar': 'Energy obtained from solar radiation.',
  'Subsurface operations for geothermal energy': 'Underground activities for exploring, developing and managing geothermal resources.',
  'Transport': 'The sector covering the movement of passengers and freight.',
  'digital and tools': 'Digital technologies and analytical tools supporting energy systems.',
  'large-scale or industrial': 'Technologies designed for large-capacity or industrial applications.',
  'mass-manufactured': 'Technologies produced in high volumes as standardised units.',
  'Thermal storage': 'Storage of energy as heat or cold for later use.',
  'Thermal storage for buildings and district heating': 'Thermal energy storage serving buildings or district heating networks.',
  'Building-level thermal storage': 'Thermal energy storage installed within a building.',
  'Combustion and conversion': 'Processes that convert fuels into heat or other energy forms.',
  'Components': 'Parts and subsystems used in aircraft.',
  'EV charging equipment': 'Equipment that supplies electricity to electric vehicles.',
  'Nuclear fuel cycles': 'Processes from nuclear fuel preparation to final waste management.',
  'Other components': 'Components not included in another defined category.',
  'Thermal processes': 'Industrial processes driven primarily by heat.',
  'Fusion': 'Energy production by combining light atomic nuclei.',
  'Sythetic fuels production': 'Production of fuels through chemical synthesis.',
};

// initiative.country in the cached dataset is only ever stored in English
// (no country_zh field exists at all, unlike sector_zh/sector_en which both
// exist) — 134 distinct raw values including multi-country combos ("China,
// Germany") and a few source-data typos ("Irland", "Netherland",
// "Philipines", "New Zeland"), mapped here exactly as they appear rather
// than parsed/split, to sidestep ambiguity between combo separators (comma
// vs slash vs "and") and place names that themselves contain a hyphen
// ("Asia-Pacific") vs combo hyphens ("Argentina-Uruguay").
const COUNTRY_NAMES_ZH = {
  'Argentina': '阿根廷',
  'Argentina-Uruguay': '阿根廷、烏拉圭',
  'Asia-Pacific': '亞太地區',
  'Australia': '澳洲',
  'Australia, United Kingdom': '澳洲、英國',
  'Austria': '奧地利',
  'Bahrain': '巴林',
  'Belgium': '比利時',
  'Belgium,Germany': '比利時、德國',
  'Brazil': '巴西',
  'Canada': '加拿大',
  'Canada, United States': '加拿大、美國',
  'Chile': '智利',
  'China': '中國',
  'China, Europe': '中國、歐洲',
  'China, Germany': '中國、德國',
  'Colombia': '哥倫比亞',
  'Croatia': '克羅埃西亞',
  'Cuba': '古巴',
  'Czech Republic': '捷克',
  'Denmark': '丹麥',
  'Denmark, India': '丹麥、印度',
  'Denmark, Sweden': '丹麥、瑞典',
  'Dubai': '杜拜',
  'EU': '歐盟',
  'Egypt': '埃及',
  'Estonia, Netherlands': '愛沙尼亞、荷蘭',
  'Europe': '歐洲',
  'Europe, United States': '歐洲、美國',
  'European Union': '歐盟',
  'European Union and United States': '歐盟、美國',
  'Finland': '芬蘭',
  'Finland, United Kingdom': '芬蘭、英國',
  'France': '法國',
  'France, Germany': '法國、德國',
  'France, Japan': '法國、日本',
  'France, Netherlands': '法國、荷蘭',
  'France, United States': '法國、美國',
  'France/Europe': '法國、歐洲',
  'France/Italy': '法國、義大利',
  'Germany': '德國',
  'Germany, France, Poland': '德國、法國、波蘭',
  'Germany/India': '德國、印度',
  'Ghana': '迦納',
  'Global': '全球',
  'Greece': '希臘',
  'Iceland': '冰島',
  'India': '印度',
  'India, Europe': '印度、歐洲',
  'India, Mexico, Burkina Faso': '印度、墨西哥、布吉納法索',
  'India, Singapore': '印度、新加坡',
  'Indonesia': '印尼',
  'International': '國際',
  'Iran, Ireland': '伊朗、愛爾蘭',
  'Ireland': '愛爾蘭',
  'Irland': '愛爾蘭',
  'Israel': '以色列',
  'Italy': '義大利',
  'Italy, France': '義大利、法國',
  'Italy, Switzerland': '義大利、瑞士',
  'Japan': '日本',
  'Japan, France': '日本、法國',
  'Japan, United States': '日本、美國',
  'Kazakhstan': '哈薩克',
  'Kenya': '肯亞',
  'Korea': '韓國',
  'Korea, Bolivia': '韓國、玻利維亞',
  'Luxembourg': '盧森堡',
  'Malaysia': '馬來西亞',
  'Mali': '馬利',
  'Mexico': '墨西哥',
  'Mix': '多國混合',
  'Morocco': '摩洛哥',
  'Namibia': '納米比亞',
  'Namibia, Belgium': '納米比亞、比利時',
  'Namibia, Germany': '納米比亞、德國',
  'Nauru, China': '諾魯、中國',
  'Netherland': '荷蘭',
  'Netherlands': '荷蘭',
  'New Zealand': '紐西蘭',
  'New Zeland, United States': '紐西蘭、美國',
  'Norway': '挪威',
  'Oman': '阿曼',
  'Oman / United Arab Emirates (UAE)': '阿曼、阿拉伯聯合大公國',
  'Philipines': '菲律賓',
  'Philippines': '菲律賓',
  'Poland': '波蘭',
  'Portugal': '葡萄牙',
  'Romania': '羅馬尼亞',
  'Rotterdam': '鹿特丹',
  'Russia': '俄羅斯',
  'Saudi Arabia': '沙烏地阿拉伯',
  'Scotland': '蘇格蘭',
  'Singapore': '新加坡',
  'Slovakia/United States': '斯洛伐克、美國',
  'Slovenia': '斯洛維尼亞',
  'South Africa': '南非',
  'South Korea': '韓國',
  'Spain': '西班牙',
  'Sweden': '瑞典',
  'Switzerland': '瑞士',
  'Switzerland / Iceland': '瑞士、冰島',
  'Switzerland, Korea': '瑞士、韓國',
  'Taiwan': '台灣',
  'Thailand': '泰國',
  'The Netherlands': '荷蘭',
  'Turkey': '土耳其',
  'U.S.': '美國',
  'UAE': '阿拉伯聯合大公國',
  'UK': '英國',
  'UK, Germany': '英國、德國',
  'UK, US': '英國、美國',
  'UK, US, Netherlands': '英國、美國、荷蘭',
  'US': '美國',
  'US, France': '美國、法國',
  'USA': '美國',
  'Uganda': '烏干達',
  'Ukraine': '烏克蘭',
  'United Arab Emirates': '阿拉伯聯合大公國',
  'United Kingdom': '英國',
  'United States': '美國',
  'United States and Belgium': '美國、比利時',
  'United States, Albania': '美國、阿爾巴尼亞',
  'United States, Europe': '美國、歐洲',
  'United States, Germany': '美國、德國',
  'United States, Japan': '美國、日本',
  'United States, Korea': '美國、韓國',
  'United States, Russia': '美國、俄羅斯',
  'United States, United Kingdom': '美國、英國',
  'United States, global': '美國、全球',
  'Uruguay, Argentina': '烏拉圭、阿根廷',
  'Venezuela': '委內瑞拉',
  'Vietnam': '越南',
  'World': '全球'
};

function translateCountry(country, uiLang) {
  if (!country) return country;
  return uiLang === 'en' ? country : (COUNTRY_NAMES_ZH[country] || country);
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
    trendInsufficientData: '此技術缺乏足夠的年度 TRL 資料，無法繪製趨勢',
    countryBarHint: '點擊長條可查看該國家的相關案例',
    countryListHeading: '相關案例',
    countryListEmpty: '此國家沒有可連結的案例',
    countryCaseHint: '點擊案例可跳轉至對應技術詳情'
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
    trendInsufficientData: 'Not enough yearly TRL data to plot a trend for this technology',
    countryBarHint: 'Click a bar to see the cases from that country',
    countryListHeading: 'Related cases',
    countryListEmpty: 'No linkable cases for this country',
    countryCaseHint: "Click a case to jump to its technology's details"
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
    download: '下載報告',
    downloadWord: '下載為 Word (.docx)',
    downloadPdf: '下載為 PDF (.pdf)',
    downloading: '生成中...',
    basicBadge: '基本統計',
    aiBadge: 'AI 洞察',
    upgradeToAI: '生成 AI 洞察報告',
    upgrading: 'AI 生成中...',
    needApiKeyForUpgrade: '請先在右上角選擇 AI 供應商並輸入對應的 API 金鑰，才能生成 AI 洞察報告。',
    prevSlide: '上一頁',
    nextSlide: '下一頁',
    prevVersion: '上一版',
    nextVersion: '下一版',
    versionLabel: (i, n) => `版本 ${i} / ${n}`,
    focusTechPlaceholder: '選擇一項技術聚焦說明...',
    focusTechButton: '聚焦此技術',
    moreConciseButton: '整體更精簡',
    rewriteOtherLangButton: '改用英文重寫',
    followUpPlaceholder: '或自行輸入其他調整指示，例如：加強說明政策風險 / 補充跟氫能的關聯',
    revising: '調整中...',
    send: '送出',
    noPreview: '無法生成預覽，請重試'
  },
  en: {
    title: 'AI Strategy Slide Preview',
    subtitle: (n) => `Analyzed ${n} technologies · click underlined tech names for details`,
    download: 'Download Report',
    downloadWord: 'Download as Word (.docx)',
    downloadPdf: 'Download as PDF (.pdf)',
    downloading: 'Generating...',
    basicBadge: 'Basic Stats',
    aiBadge: 'AI Insights',
    upgradeToAI: 'Generate AI Insight Report',
    upgrading: 'Generating with AI...',
    needApiKeyForUpgrade: 'Choose an AI provider and enter its API key in the top-right corner to generate the AI insight report.',
    prevSlide: 'Previous',
    nextSlide: 'Next',
    prevVersion: 'Previous version',
    nextVersion: 'Next version',
    versionLabel: (i, n) => `Version ${i} / ${n}`,
    focusTechPlaceholder: 'Choose a technology to focus on...',
    focusTechButton: 'Focus on this',
    moreConciseButton: 'Make it more concise',
    rewriteOtherLangButton: 'Rewrite in Traditional Chinese',
    followUpPlaceholder: 'Or type another instruction, e.g. "Say more about policy risk" / "Add hydrogen linkages"',
    revising: 'Revising...',
    send: 'Send',
    noPreview: 'Could not generate a preview, please try again'
  }
};

// Fixed prompt text for NotebookLM (or any chat-based tool) to turn a
// pasted "NotebookLM article format" source into a 6-page, 16:9 technical
// slide outline. Provided verbatim by the user — not generated or edited
// by the app — and always copied as-is regardless of uiLang, since the
// prompt itself hard-requires Traditional Chinese (Taiwan) output.
const NOTEBOOKLM_PPT_PROMPT = `請只根據本筆記本已匯入的來源資料，製作一份6頁、16:9的技術分析簡報，供技術團隊及主管報告使用。

【分析與製作原則】
1. 全程使用台灣慣用繁體中文；技術英文名稱或縮寫可於第一次出現時放在括號內。
2. 所有技術名稱、TRL、數字、年份、國家、案例、應用潛力及主要限制，只能使用本筆記本來源，不得自行補充或推測。
3. 先通讀全部資料，提煉3至5項跨技術發現，不依資料順序逐項介紹。
4. 準確區分研究、測試、示範、建置、運轉及商業化，不得提高案例的實際進展程度。
5. 每頁只傳達一項主要訊息，正文原則上最多3項重點；第2頁可使用4個精簡摘要區塊。
6. 投影片正文只呈現技術結論、數據、案例、潛力、限制與下一步，不放置分析方法、免責文字、資料限制或防錯說明。
7. 六頁維持一致的專業配色與版面，優先使用矩陣、時間軸、案例卡片及簡潔圖示，避免3D圖表、裝飾性圖片、警語及長段文字。
8. 矩陣須清楚標示座標軸、分類與技術位置；案例比較須使用一致的資訊結構。
9. 生成前檢查數據、案例狀態、分類、繁體中文、圖表標示及版面，確認均與來源一致且清楚可讀。

【六頁架構】
第1頁｜封面
- 依資料主題撰寫精簡且具分析性的標題。
- 副標題說明技術範圍與分析目的。
- 顯示資料集或平台名稱。
- 不使用來源無法支持的預測或宣稱。

第2頁｜技術現況一頁掌握
本頁直接先給答案，不介紹分析方法。以4個精簡摘要區塊呈現：
1. 技術總數與領域分布；
2. TRL成熟度重點；
3. 案例數量與主要進展；
4. 應用潛力與主要限制關鍵字。
頁面下方以一句話呈現最重要的整體結論，且該結論須能由第3至5頁的資料支持。

第3頁｜技術成熟度分布矩陣
- 標題直接呈現最重要的TRL分布結論。
- 橫軸使用TRL 1至9。
- 縱軸使用來源中的技術領域或應用類別。
- 將全部技術放入矩陣，標示技術名稱及實際TRL。
- 右側列出2至3項具體發現，每項只呈現結論、代表性技術及TRL。
- 不加入案例細節或長段文字。

第4頁｜關鍵案例：應用潛力與主要限制
選擇3個資料完整且最具分析價值的案例。每個案例呈現：
1. 技術名稱及TRL；
2. 年份、國家及案例名稱；
3. 實際進展或具體成果；
4. 一項應用潛力；
5. 一項主要限制。
應用潛力與主要限制須從該技術的現況摘要、市場動態及案例內容歸納，不得自行評分或加入來源未提及的內容。
來源未明確提及主要限制時，直接省略，不要顯示「資料未提供」等提醒。
使用三欄案例比較或三張案例卡片呈現，每個案例使用相同結構並避免長段文字。

第5頁｜商業化進展與後續重點
- 綜合全部技術及案例，製作「TRL成熟度 × 案例實證階段」矩陣。
- 橫軸：TRL成熟度。
- 縱軸：研究、測試、示範、建置、運轉或商業化。
- 只放入案例資訊足以判斷實證階段的代表性技術。
- 每項技術只標示技術名稱、TRL及案例階段，不重複第4頁的潛力與限制。
- 右側呈現2項整體發現及1項後續追蹤方向。
- 若來源不足以建立案例實證階段矩陣，改用「TRL × 案例數」呈現，不得強行分類。

第6頁｜感謝聆聽
- 顯示「Q&A」及資料集或平台名稱。
- 不自行增加網址、聯絡人或電子郵件。
- 版面簡潔並保留充足留白。

請直接生成最終6頁簡報，不必先提供大綱、分析過程或設計方案。`;

// Step-by-step walkthrough shown in PptGuideModal, built from real product
// screenshots (public/guide/guide-N.png) rather than icons, since the
// previous icon-only version couldn't show *where* on NotebookLM's actual
// UI to click. Each `arrows` entry is a normalized (0-100%) point on the
// screenshot plus an approach direction ('nw' = arrow tail up-left of the
// target, 'ne' = tail up-right — pick whichever keeps the tail from running
// off the image edge) — see GuideArrow for how these render.
//
// Steps 5 and 6 are ordered prompt-copy-before-reopening-the-slide-panel
// (not the order the screenshots were captured in) — copying the prompt
// first means the user only has to re-enter NotebookLM once to both open
// the slide-customization panel and immediately paste into it, instead of
// opening the panel, leaving to fetch the prompt, then coming back.
const PPT_GUIDE_CONTENT = {
  zh: {
    title: '簡報生成使用指南',
    stepLabel: (i, n) => `步驟 ${i} / ${n}`,
    prev: '上一步',
    next: '下一步',
    done: '關閉',
    steps: [
      {
        image: './guide/guide-1.png',
        title: '複製文章並前往 NotebookLM',
        desc: '在平台的查詢策略總覽頁，點擊「複製 NotebookLM 文章格式」，再點擊「前往 NotebookLM」開啟新分頁。',
        arrows: [{ x: 27.7, y: 71.6, dir: 'nw' }, { x: 66.8, y: 71.6, dir: 'nw' }],
      },
      {
        image: './guide/guide-2.png',
        title: '建立新的筆記本',
        desc: '在 NotebookLM 首頁點擊「+」，建立一個新筆記本。',
        arrows: [{ x: 62, y: 64, dir: 'nw' }],
      },
      {
        image: './guide/guide-3.png',
        title: '選擇「複製的文字」來源',
        desc: '在新增來源畫面中，點擊「複製的文字」。',
        arrows: [{ x: 80, y: 86, dir: 'nw' }],
      },
      {
        image: './guide/guide-4.png',
        title: '貼上文章並插入',
        desc: '貼上剛才複製的文章內容，點擊「插入」。',
        arrows: [{ x: 87, y: 92, dir: 'nw' }],
      },
      {
        image: './guide/guide-5.png',
        title: '回到平台，複製簡報生成 Prompt',
        desc: '回到平台，點擊「複製簡報生成 Prompt」——先把這段文字準備好，等一下回到 NotebookLM 就能一次貼上。',
        arrows: [{ x: 23.8, y: 84.7, dir: 'nw' }],
      },
      {
        image: './guide/guide-6.png',
        title: '開啟簡報功能',
        desc: '回到 NotebookLM，在「工作室」面板點擊「簡報」右邊的「〉」。',
        arrows: [{ x: 84, y: 32, dir: 'nw' }],
      },
      {
        image: './guide/guide-7.png',
        title: '選擇語言並貼上 Prompt',
        desc: '將語言設定為希望簡報呈現的語言，在「說明要建立的簡報」欄位貼上剛複製的 Prompt，再點擊「生成」。',
        arrows: [{ x: 15, y: 56, dir: 'ne' }, { x: 49, y: 76, dir: 'nw' }, { x: 90, y: 93, dir: 'nw' }],
      },
      {
        image: null,
        title: '等待簡報生成',
        desc: 'NotebookLM 正在讀完你匯入的所有資料——好的簡報值得等一下，稍後回來查看成果吧。',
        arrows: [],
      },
    ],
  },
  en: {
    title: 'Slide Generation Guide',
    stepLabel: (i, n) => `Step ${i} of ${n}`,
    prev: 'Previous',
    next: 'Next',
    done: 'Close',
    steps: [
      {
        image: './guide/guide-1.png',
        title: 'Copy the article and open NotebookLM',
        desc: 'On the platform\'s Strategy Overview tab, click "Copy NotebookLM article format", then click "Open NotebookLM" to open a new tab.',
        arrows: [{ x: 27.7, y: 71.6, dir: 'nw' }, { x: 66.8, y: 71.6, dir: 'nw' }],
      },
      {
        image: './guide/guide-2.png',
        title: 'Create a new notebook',
        desc: 'On the NotebookLM home page, click "+" to create a new notebook.',
        arrows: [{ x: 62, y: 64, dir: 'nw' }],
      },
      {
        image: './guide/guide-3.png',
        title: 'Choose "Copied text" as the source',
        desc: 'On the add-source screen, click "Copied text".',
        arrows: [{ x: 80, y: 86, dir: 'nw' }],
      },
      {
        image: './guide/guide-4.png',
        title: 'Paste the article and insert it',
        desc: 'Paste what you copied, then click "Insert".',
        arrows: [{ x: 87, y: 92, dir: 'nw' }],
      },
      {
        image: './guide/guide-5.png',
        title: 'Back on the platform, copy the prompt',
        desc: 'Back on the platform, click "Copy slide-generation prompt" — getting it ready now means you only need one more trip back into NotebookLM.',
        arrows: [{ x: 23.8, y: 84.7, dir: 'nw' }],
      },
      {
        image: './guide/guide-6.png',
        title: 'Open the slide deck feature',
        desc: 'Back in NotebookLM, in the "Studio" panel, click the ">" next to "Slide deck".',
        arrows: [{ x: 84, y: 32, dir: 'nw' }],
      },
      {
        image: './guide/guide-7.png',
        title: 'Choose a language and paste the prompt',
        desc: 'Set the language, paste the prompt you copied into "Describe the slide deck that you want to create", then click "Generate".',
        arrows: [{ x: 15, y: 56, dir: 'ne' }, { x: 49, y: 76, dir: 'nw' }, { x: 90, y: 93, dir: 'nw' }],
      },
      {
        image: null,
        title: 'Wait for the slides to generate',
        desc: "NotebookLM is working through everything you gave it — a good deck is worth the wait. Check back in a bit.",
        arrows: [],
      },
    ],
  },
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
    appSubtitle: '技術資料庫查詢 & AI 洞察分析',
    datasetVersion: (v) => `資料版本: ${v}`,
    datasetTimeLabel: '資料版本時間',
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
    copyArticle: '複製 NotebookLM 文章格式',
    copied: '已複製！',
    openNotebookLM: '前往 NotebookLM',
    copyPptPrompt: '複製簡報生成 Prompt',
    openPptGuide: '使用指南',
    pptCardTitle: '用 NotebookLM 製作簡報',
    pptCardDescPrefix: '透過 NotebookLM 的 AI 功能，把這次查詢結果的',
    pptCardDescCheckedPrefix: '透過 NotebookLM 的 AI 功能，把你勾選的',
    pptCardDescSuffix: '筆技術轉換成一份簡報！',
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
    fieldAdoptionStage: (s) => `採用階段: ${s}`,
    fieldCaseCount: (n) => `相關案例: ${n} 筆`,
    pathBreadcrumbLabel: '分類路徑（點擊可篩選同路徑的技術）',
    pathFilterActive: '路徑篩選中',
    clearPathFilter: '清除路徑篩選',
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
    sectorDistributionTitle: '產業別分佈',
    sectorDistributionCenterLabel: '項技術',
    sectorDistributionHint: '滑鼠移到色塊查看細節',
    strategyCardTitle: '策略總覽與 AI 洞察',
    strategyCardDescPrefix: '立即查看目前符合條件的全部',
    strategyCardDescCheckedPrefix: '以你勾選的',
    strategyCardDescSuffix: '筆技術的整體分析，之後可選擇升級為 AI 深度洞察。',
    openStrategyBtn: '開啟策略總覽',
    checkedSubsetBadge: (n) => `已依勾選項目篩選：${n} 項技術`,
    checkedCountBar: (n) => `已勾選 ${n} 項技術（策略總覽將以勾選項目為準）`,
    clearChecked: '清除勾選',
    collapseSearch: '收合搜尋列',
    expandSearch: '展開搜尋列',
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
    appSubtitle: 'Technology Database Search & AI Insight Analysis',
    datasetVersion: (v) => `Dataset version: ${v}`,
    datasetTimeLabel: 'Dataset Generated',
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
    copyArticle: 'Copy NotebookLM article format',
    copied: 'Copied!',
    openNotebookLM: 'Open NotebookLM',
    copyPptPrompt: 'Copy slide-generation prompt',
    openPptGuide: 'How to use',
    pptCardTitle: 'Build Slides with NotebookLM',
    pptCardDescPrefix: "Using NotebookLM's AI, turn the",
    pptCardDescCheckedPrefix: "Using NotebookLM's AI, turn your checked",
    pptCardDescSuffix: 'matching technologies into a slide deck!',
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
    fieldAdoptionStage: (s) => `Adoption stage: ${s}`,
    fieldCaseCount: (n) => `Related cases: ${n}`,
    pathBreadcrumbLabel: 'Classification path (click a segment to filter matching technologies)',
    pathFilterActive: 'Path filter active',
    clearPathFilter: 'Clear path filter',
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
    sectorDistributionTitle: 'Sector Distribution',
    sectorDistributionCenterLabel: 'technologies',
    sectorDistributionHint: 'Hover a segment for details',
    strategyCardTitle: 'Strategy Overview & AI Insight',
    strategyCardDescPrefix: 'Instantly view an overview of all',
    strategyCardDescCheckedPrefix: 'Run on your',
    strategyCardDescSuffix: 'matching technologies, with the option to upgrade to AI-generated insight afterwards.',
    openStrategyBtn: 'Open Strategy Overview',
    checkedSubsetBadge: (n) => `Filtered to your checked selection: ${n} technologies`,
    checkedCountBar: (n) => `${n} technologies checked (the strategy overview will use only these)`,
    clearChecked: 'Clear selection',
    collapseSearch: 'Collapse search bar',
    expandSearch: 'Expand search bar',
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
  let str = String(val);
  // Neutralize spreadsheet formula injection: Excel treats cells starting
  // with = + - @ (or tab/CR) as formulas even inside a quoted CSV field,
  // which malicious cell content (e.g. from an untrusted uploaded dataset)
  // could abuse to execute commands on open. A leading apostrophe forces
  // Excel to read the cell as plain text.
  if (/^[=+\-@\t\r]/.test(str)) str = "'" + str;
  return `"${str.replace(/"/g, '""')}"`;
}

// Only http/https URLs may be rendered as clickable links — a dataset is
// user-uploadable JSON, so a malicious file could otherwise smuggle a
// javascript: URL into a link's href (React only warns on those, it does
// not block them), which on click would run script in the app's origin
// and could e.g. read the stored API keys.
function safeLinkUrl(url) {
  if (typeof url !== 'string') return null;
  return /^https?:\/\//i.test(url.trim()) ? url.trim() : null;
}

// Shared by both the user-facing CSV export and the AI briefing-report
// prompt (see buildBriefingReportPrompt) — curated to exactly the fields
// worth a column, not every field on the object. Checked against the real
// dataset before trimming: "theme"/"keyCountries"/the technology-level
// "read_more" are empty on every one of the 639 cached technologies;
// "description"/"NZErationale" are byte-for-byte duplicates of
// technology_status_summary_en/market_dynamics_summary_en; those _en
// summary fields themselves are never read by the prompt (it always
// synthesizes from the _zh columns, regardless of output language); the
// "sector" breadcrumb array duplicates sector_zh/sector_en; and the
// trl_2020..trl_2025 per-year fields only feed the in-app trend chart, not
// this report (which only cites latest_trl). "adoption_stage" (Awaiting
// adoption / Building momentum / Commercial in many markets) exists in the
// original raw IEA source but was dropped somewhere in this dataset's
// rebuild — restored here since it's a genuinely different axis from TRL
// (market/commercial diffusion, not technical maturity), populated on 227
// of the 639 technologies, empty string on the rest. Includes "_id" as the
// first column so report content stays traceable back to real technology
// records, and the full "initiatives" array (serialized as JSON per cell)
// so the report can cite real project names/figures instead of inventing
// plausible-sounding ones.
function buildFullDataCsv(techs) {
  const allKeys = [
    '_id', 'technology_name', 'technology_name_zh', 'sector_zh', 'sector_en',
    'latest_trl', 'adoption_stage', 'technology_status_summary_zh', 'market_dynamics_summary_zh',
    'linked_records_count', 'initiatives'
  ];
  let csvContent = allKeys.map(key => escapeCsv(key)).join(",") + "\n";
  techs.forEach(tech => {
    const row = allKeys.map(key => {
      let val = tech[key];
      if (typeof val === 'object' && val !== null) val = JSON.stringify(val);
      return escapeCsv(val);
    });
    csvContent += row.join(",") + "\n";
  });
  return csvContent;
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

function computeTechStats(techs, uiLang = 'zh') {
  const totalFound = techs.length;

  const sectorCounts = {};
  techs.forEach(r => {
    const s = (uiLang === 'en' ? r.sector_en : r.sector_zh) || (uiLang === 'en' ? 'Unlabeled' : '未標示');
    sectorCounts[s] = (sectorCounts[s] || 0) + 1;
  });
  const topSector = Object.entries(sectorCounts).sort((a, b) => b[1] - a[1])[0]?.[0] || (uiLang === 'en' ? 'Unlabeled' : '未標示');

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

  // techId/techName carried on each initiative (not just the raw record)
  // so slide3's country drill-down can list which case came from which
  // technology and jump there — initiatives have no page of their own,
  // the technology they're attached to is the only thing that's navigable.
  const allInitiatives = techs.flatMap(r => (r.initiatives || []).map(init => ({ ...init, techId: r._id, techName: techLabel(r) })));
  const countryGroups = {};
  allInitiatives.forEach(i => {
    if (!i.country) return;
    const country = translateCountry(i.country, uiLang);
    (countryGroups[country] = countryGroups[country] || []).push(i);
  });
  // {name, count, initiatives} per country (not just a name) so slide3 can
  // render an actual bar chart proportional to case volume, and drill down
  // into the exact cases behind it — same shape/pattern as trlBreakdown's
  // {stage, count, techs} above.
  const topCountries = Object.entries(countryGroups)
    .sort((a, b) => b[1].length - a[1].length)
    .map(([name, initiatives]) => ({ name, count: initiatives.length, initiatives }));

  const topTechs = [...techs].sort((a, b) => (b.linked_records_count || 0) - (a.linked_records_count || 0));

  return { totalFound, topSector, trlBreakdown, allInitiatives, topCountries, topTechs };
}

// Full document mode: an executive summary, a full overview table (one row
// per matched technology), a dedicated deep-dive section for every single
// matched technology (heading + intro + labeled sub-bullets, e.g. "技術特徵"
// / "市場進展"), and a closing cross-technology trends section. Deliberately
// no cap on how many tech_sections get generated — every matched technology
// gets one, however many that is; the caller is expected to feed the full
// per-technology field set (see buildFullDataCsv), including every linked
// case record, not just the two summary paragraphs, so this can cite real
// project names/capacities instead of inventing plausible-sounding ones.
function buildBriefingReportPrompt(csvStr, uiLang = 'zh') {
  const languageInstruction = uiLang === 'en'
    ? '整個平台目前設定為英文介面。請直接用英文輸出所有內容：report 底下每一個欄位（title、executive_summary、overview_table、tech_sections、trends_section 裡的所有文字），以及所有 "name" 顯示用技術名稱，全部都要是英文，不要輸出中文，也不需要使用者另外追問才翻譯。CSV 裡的 "_id" 仍必須照抄原始字串，不受語言影響。'
    : '請用繁體中文輸出所有內容。';

  return `你是內建於「淨零碳排技術查詢平台」的策略分析引擎。以下 CSV 是本次查詢結果的技術清單，每一列除了名稱、產業別、TRL、兩段摘要文字之外，"adoption_stage" 欄位是這項技術的市場採用階段（Awaiting adoption／Building momentum／Commercial in many markets 三種之一，空字串代表 IEA 未評級，不代表資料缺漏），這是跟 TRL（技術本身成熟度）不同的獨立維度，代表市場／商業化普及程度，若某項技術有這個欄位、請在分析中明確運用它（尤其是市場進展、商業化挑戰相關的段落），"initiatives" 欄位是這項技術連結的完整案例清單（JSON 陣列，每筆案例可能包含年份、國家、類型、描述、來源連結），其餘欄位是原始資料庫裡的其他欄位。

前四頁統計摘要（查詢總覽、技術成熟度分佈、專案活動熱區、關鍵技術清單）已經由本地程式正確算好，不需要你重做，也不在你的輸出範圍內。你的任務是撰寫一份完整、詳盡的「簡要報告」（briefing document），結構如下：

1. **摘要（Executive Summary）**：開門見山，一段精簡但完整地點出本次查詢範圍內最關鍵的結論與重點。
2. **技術現況總覽表**：針對 CSV 裡「每一項」技術，各生成一列，欄位為技術類別（名稱）、能源來源／技術原理（一句話）、TRL 等級、主要市場動態（一句話摘要）。這是表格，不是段落，每項技術都要有自己的一列，不可省略或合併。
3. **核心技術深度分析**：針對 CSV 裡「每一項」技術，各生成一個獨立小節，包含：一段簡介（說明技術原理與定位）、以及依內容需要命名的子項目（例如「技術特徵」「市場進展」「發展限制」「關鍵動態」「挑戰」等，項目名稱依實際內容彈性決定，不必每項技術都用同樣的子項目名稱），子項目底下用條列方式呈現重點，盡量引用 "initiatives" 案例資料裡的具體專案名稱、地點、數字（容量、金額、年份等）作為佐證——只能使用 CSV 資料裡實際出現的具體數字與名稱，絕對不可自行編造。如果某項技術的案例或摘要資料很少，仍要盡力根據既有資料寫出簡要但真實的內容，不要跳過任何一項技術。
4. **關鍵趨勢與戰略洞察**：綜合跨技術的觀察，依主題分成幾個子項目（例如「系統性效益」「經濟與環境價值」「商業化挑戰」等），每個子項目底下條列重點，並在提及具體技術時附上追溯依據。

語氣必須客觀、犀利，不浮誇、不重述已經算好的統計數字（TRL 分佈、案例數量等），聚焦在單靠數字看不出來的洞察：技術之間的關聯性、風險、機會、容易被忽略但值得注意的異常或矛盾之處。

${languageInstruction}

嚴格輸出 JSON 格式（不要包含任何 markdown 符號或說明文字，直接輸出純 JSON），只需要 "report" 這一個頂層欄位，結構完全比照下方【輸出格式】。CSV 第一欄 "_id" 是每項技術在資料庫中的唯一識別碼："overview_table.rows"、"tech_sections" 都要用 "id" 欄位照抄該技術在 CSV 裡對應的 "_id" 原始字串，絕對不可自行編造或省略；"trends_section" 裡提及具體技術支持論點時，也要在對應的 "related_techs" 用 {"id": "...", "name": "..."} 的形式列出，規則相同。

【輸入資料】
${csvStr}

【輸出格式】
{
  "report": {
    "title": "簡要報告標題",
    "executive_summary": "一段精簡但完整的摘要，開門見山列出最關鍵的結論，讓讀者一眼掌握重點",
    "overview_table": {
      "rows": [
        {"id": "CSV裡的_id原值", "name": "技術名稱", "energy_source": "能源來源／技術原理，一句話", "trl": "TRL 等級", "market_dynamics": "主要市場動態，一句話摘要"},
        "...CSV 裡每一項技術都要有一列，不可省略..."
      ]
    },
    "tech_sections": [
      {
        "id": "CSV裡的_id原值",
        "heading": "技術名稱",
        "intro": "一段簡介，說明技術原理與定位",
        "subsections": [
          {"label": "依內容命名的子項目標題，例如：技術特徵", "bullets": ["具體重點一，盡量引用案例資料裡的真實數字與專案名稱", "具體重點二"]},
          {"label": "依內容命名的子項目標題，例如：市場進展", "bullets": ["具體重點一", "具體重點二"]}
        ]
      },
      "...CSV 裡每一項技術都要有自己的一個 tech_sections 項目，不可省略..."
    ],
    "trends_section": {
      "heading": "關鍵趨勢與戰略洞察",
      "subsections": [
        {
          "label": "子項目標題，例如：系統性效益",
          "bullets": [
            {"text": "具體洞察，引用資料中的具體內容作為依據", "related_techs": [{"id": "CSV裡的_id原值", "name": "技術名稱"}, "..."]}
          ]
        },
        "...依內容可以有多個子項目，例如經濟與環境價值、商業化挑戰等..."
      ]
    }
  }
}`;
}

// Produces only slide1-4 (overview / TRL / hotspots / top techs) — purely
// computed from real data, so this is the "instant, no-wait" baseline every
// strategy overview opens with. The narrative briefing report is exclusively
// an AI-generated addition (see buildBriefingReportPrompt) so the two tiers
// are structurally different, not just differently-worded restatements of
// the same stats.
function buildFallbackAnalysis(techs, uiLang = 'zh') {
  const { totalFound, topSector, trlBreakdown, allInitiatives, topCountries, topTechs } = computeTechStats(techs, uiLang);
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
          { value: totalFound, label: "Technologies", caption: "Matched to this query" },
          { value: allInitiatives.length, label: "Related cases", caption: "Linked case records" },
          { value: topCountries.length, label: "Countries covered", caption: "Regions represented" }
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
        countries: top5Countries
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
        { value: totalFound, label: "技術項目", caption: "符合本次查詢條件" },
        { value: allInitiatives.length, label: "相關案例", caption: "已連結案例紀錄" },
        { value: topCountries.length, label: "涵蓋國家", caption: "案例分布地區數" }
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
      countries: top5Countries
    },
    slide4: {
      title: "關鍵技術清單",
      subtitle: "依案例連結數排序",
      top_techs: top5Techs
    },
    ui_labels: DEFAULT_UI_LABELS.zh
  };
}

// Same field set as the CSV export (every key on the tech object, not a
// curated subset) — just laid out as a headed Markdown article instead of a
// spreadsheet row, so tools like NotebookLM that read prose have enough
// narrative context to work with. Shared by both the file-download export
// and the clipboard-copy export.
function buildArticleMarkdown(techs, uiLang = 'zh') {
  const isEn = uiLang === 'en';
  const nameOf = (tech) => (isEn ? (tech.technology_name || tech.technology_name_zh) : (tech.technology_name_zh || tech.technology_name)) || (isEn ? 'Unlabeled' : '未標示');
  const structuredKeys = new Set([
    'technology_name_zh', 'technology_name', 'sector_zh', 'sector_en',
    'latest_trl', 'linked_records_count',
    'technology_status_summary_zh', 'technology_status_summary_en',
    'market_dynamics_summary_zh', 'market_dynamics_summary_en',
    'initiatives', '_id', 'id'
  ]);

  const formatInitiatives = (initiatives) => {
    if (!Array.isArray(initiatives) || initiatives.length === 0) {
      return isEn ? '(No linked cases)' : '（無相關案例資料）';
    }
    return initiatives.map((init, i) => {
      const meta = [init.year, translateCountry(init.country, uiLang), init.type].filter(Boolean).join(' / ');
      let line = `${i + 1}. ${meta ? `[${meta}] ` : ''}${init.description || (isEn ? 'Unlabeled' : '未標示')}`;
      if (init.read_more) line += ` (${init.read_more})`;
      return line;
    }).join('\n');
  };

  let md = isEn ? `# IEA Net-Zero Technology Intelligence Platform — Full Export\n\n` : `# IEA 淨零技術智慧分析平台 — 完整匯出報告\n\n`;
  md += isEn ? `Exported: ${new Date().toLocaleString('en-US')}\n` : `匯出時間：${new Date().toLocaleString('zh-TW')}\n`;
  md += isEn ? `Technologies: ${techs.length}\n\n---\n\n` : `技術總數：${techs.length}\n\n---\n\n`;

  techs.forEach((tech, idx) => {
    md += `## ${idx + 1}. ${nameOf(tech)}\n\n`;
    const altName = isEn ? tech.technology_name_zh : tech.technology_name;
    if (altName) md += `*${altName}*\n\n`;
    md += isEn
      ? `- Sector: ${tech.sector_en || tech.sector_zh || 'Unlabeled'}\n- TRL: ${tech.latest_trl || 'Unlabeled'}\n- Related cases: ${tech.linked_records_count || 0}\n\n`
      : `- 產業別：${tech.sector_zh || tech.sector_en || '未標示'}\n- TRL 等級：${tech.latest_trl || '未標示'}\n- 相關案例數：${tech.linked_records_count || 0}\n\n`;

    const statusSummary = isEn
      ? (tech.technology_status_summary_en || tech.technology_status_summary_zh)
      : tech.technology_status_summary_zh;
    if (statusSummary) md += `### ${isEn ? 'Technology Status Summary' : '技術現況摘要'}\n${statusSummary}\n\n`;

    const marketSummary = isEn
      ? (tech.market_dynamics_summary_en || tech.market_dynamics_summary_zh)
      : tech.market_dynamics_summary_zh;
    if (marketSummary) md += `### ${isEn ? 'Market Dynamics' : '市場動態'}\n${marketSummary}\n\n`;

    // Every remaining field on the object, so nothing gets dropped
    // relative to the CSV export's full column set.
    const otherKeys = Object.keys(tech).filter(k => !structuredKeys.has(k));
    if (otherKeys.length > 0) {
      md += `### ${isEn ? 'Other Fields' : '其他欄位資料'}\n`;
      otherKeys.forEach(key => {
        let val = tech[key];
        if (val === null || val === undefined || val === '') return;
        if (typeof val === 'object') val = JSON.stringify(val);
        md += `- **${key}**: ${val}\n`;
      });
      md += `\n`;
    }

    md += `### ${isEn ? 'Related Cases (Initiatives)' : '相關案例（Initiatives）'}\n${formatInitiatives(tech.initiatives)}\n\n---\n\n`;
  });

  return md;
}

function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Renders the report as a plain HTML string for the PDF export path. This
// gets mounted off-screen and rasterized (see handleDownloadReportPdf)
// rather than drawn with jsPDF's native text APIs, specifically to sidestep
// jsPDF's lack of built-in Traditional Chinese glyph support — embedding a
// CJK font in jsPDF requires converting a whole .ttf to base64 ahead of
// time, which isn't practical here. Rasterizing whatever font the user's
// own browser already renders Chinese with (the same one the on-screen UI
// uses) sidesteps that entirely, at the cost of the PDF's text not being
// selectable.
function buildReportHtml(report, uiLang, techById, relatedPrefix) {
  const isEn = uiLang === 'en';
  let html = `<div style="font-family: 'Noto Sans TC', 'Noto Sans', sans-serif; color:#1e293b; padding:32px; width:800px; background:#ffffff;">`;
  html += `<h1 style="font-size:26px; font-weight:800; margin:0 0 16px; color:#0f172a;">${escapeHtml(report.title || (isEn ? 'AI Insight Report' : 'AI 洞察報告'))}</h1>`;

  if (report.executive_summary) {
    html += `<div style="background:#eff6ff; border-left:4px solid #2563eb; padding:16px; margin-bottom:24px; border-radius:6px;">`;
    html += `<div style="font-size:13px; font-weight:700; color:#2563eb; margin-bottom:4px;">${isEn ? 'Executive Summary' : '摘要'}</div>`;
    html += `<div style="font-size:14px; line-height:1.7;">${escapeHtml(report.executive_summary)}</div>`;
    html += `</div>`;
  }

  const tableRows = report.overview_table?.rows;
  if (Array.isArray(tableRows) && tableRows.length > 0) {
    html += `<h2 style="font-size:19px; font-weight:700; margin:24px 0 10px; color:#0f172a;">${isEn ? 'Technology Overview' : '技術現況總覽表'}</h2>`;
    const headerLabels = isEn
      ? ['Technology', 'Energy Source', 'TRL', 'Market Dynamics']
      : ['技術類別', '能源來源', 'TRL', '主要市場動態'];
    html += `<table style="width:100%; border-collapse:collapse; font-size:12px; margin-bottom:16px;"><thead><tr>`;
    html += headerLabels.map(h => `<th style="text-align:left; padding:6px 8px; background:#1f2937; color:#ffffff; border:1px solid #cbd5e1;">${escapeHtml(h)}</th>`).join('');
    html += `</tr></thead><tbody>`;
    tableRows.forEach(row => {
      html += `<tr>`
        + `<td style="padding:6px 8px; border:1px solid #e2e8f0; font-weight:600; vertical-align:top;">${escapeHtml(row.name)}</td>`
        + `<td style="padding:6px 8px; border:1px solid #e2e8f0; vertical-align:top;">${escapeHtml(row.energy_source)}</td>`
        + `<td style="padding:6px 8px; border:1px solid #e2e8f0; vertical-align:top; text-align:center;">${escapeHtml(row.trl)}</td>`
        + `<td style="padding:6px 8px; border:1px solid #e2e8f0; vertical-align:top;">${escapeHtml(row.market_dynamics)}</td>`
        + `</tr>`;
    });
    html += `</tbody></table>`;
  }

  const techSections = report.tech_sections;
  if (Array.isArray(techSections) && techSections.length > 0) {
    html += `<h2 style="font-size:19px; font-weight:700; margin:24px 0 10px; color:#0f172a;">${isEn ? 'In-Depth Technology Analysis' : '核心技術深度分析'}</h2>`;
    techSections.forEach((section, i) => {
      html += `<h3 style="font-size:16px; font-weight:700; margin:16px 0 6px; color:#1e293b;">${i + 1}. ${escapeHtml(section.heading)}</h3>`;
      if (section.intro) {
        html += `<div style="font-size:13px; font-style:italic; color:#64748b; margin-bottom:8px; line-height:1.6;">${escapeHtml(section.intro)}</div>`;
      }
      (section.subsections || []).forEach(sub => {
        html += `<div style="font-size:13px; font-weight:700; color:#475569; margin:10px 0 4px;">${escapeHtml(sub.label)}</div>`;
        html += `<ul style="margin:0 0 8px; padding-left:20px;">`;
        (sub.bullets || []).forEach(bullet => {
          const text = typeof bullet === 'string' ? bullet : bullet.text;
          html += `<li style="font-size:13px; line-height:1.6; margin-bottom:4px;">${escapeHtml(text)}</li>`;
        });
        html += `</ul>`;
      });
    });
  }

  if (report.trends_section) {
    html += `<h2 style="font-size:19px; font-weight:700; margin:24px 0 10px; color:#0f172a;">${escapeHtml(report.trends_section.heading || (isEn ? 'Key Trends & Strategic Insights' : '關鍵趨勢與戰略洞察'))}</h2>`;
    (report.trends_section.subsections || []).forEach(sub => {
      html += `<div style="font-size:13px; font-weight:700; color:#475569; margin:10px 0 4px;">${escapeHtml(sub.label)}</div>`;
      html += `<ul style="margin:0 0 8px; padding-left:20px;">`;
      (sub.bullets || []).forEach(bullet => {
        const text = typeof bullet === 'string' ? bullet : bullet.text;
        const relatedTechs = (typeof bullet === 'string' ? [] : (bullet.related_techs || [])).filter(t => techById.has(t.id));
        const relatedNames = relatedTechs.map(t => t.name);
        let line = escapeHtml(text);
        if (relatedNames.length > 0) {
          const separator = isEn ? ': ' : '：';
          const joiner = isEn ? ', ' : '、';
          line += ` <span style="font-style:italic; color:#94a3b8;">(${escapeHtml(relatedPrefix)}${escapeHtml(separator)}${escapeHtml(relatedNames.join(joiner))})</span>`;
        }
        html += `<li style="font-size:13px; line-height:1.6; margin-bottom:4px;">${line}</li>`;
      });
      html += `</ul>`;
    });
  }

  html += `</div>`;
  return html;
}

// Shared by both the AI-report PDF export and the no-AI-needed stats-only
// PDF export below: mounts an HTML string off-screen, rasterizes it
// (html2canvas), and slices it across as many A4 pages as needed. Page
// breaks are computed from real element positions (getBoundingClientRect,
// not offsetTop — table rows can become their own offsetParent, which
// throws off offsetTop-based math), so a break only ever lands in the
// whitespace gap between two blocks/table rows/list items, never through
// the middle of one — a naive fixed-height slice cut text off mid-line
// whenever a line happened to straddle a page boundary.
async function renderHtmlToPdf(html, filename) {
  const container = document.createElement('div');
  container.style.position = 'fixed';
  container.style.left = '-99999px';
  container.style.top = '0';
  container.innerHTML = html;
  document.body.appendChild(container);

  try {
    // reportRoot is the single outer <div> buildReportHtml wraps everything
    // in. Table rows and list items get expanded into their own breakable
    // units, since either can independently grow past a page's height as
    // the number of matched technologies grows.
    const reportRoot = container.firstElementChild;
    const collectUnits = (root) => {
      const units = [];
      Array.from(root.children).forEach(child => {
        if (child.tagName === 'TABLE') {
          const rows = Array.from(child.querySelectorAll('tr'));
          (rows.length > 0 ? rows : [child]).forEach(row => units.push(row));
        } else if (child.tagName === 'UL' || child.tagName === 'OL') {
          const items = Array.from(child.children);
          (items.length > 0 ? items : [child]).forEach(item => units.push(item));
        } else {
          units.push(child);
        }
      });
      return units;
    };
    const units = collectUnits(reportRoot);
    const rootRect = reportRoot.getBoundingClientRect();
    const unitPositions = units.map(u => {
      const r = u.getBoundingClientRect();
      return { top: r.top - rootRect.top, bottom: r.bottom - rootRect.top };
    });
    const contentWidthPx = reportRoot.offsetWidth;
    const totalHeightPx = rootRect.height;

    const canvas = await html2canvas(container, { scale: 2, backgroundColor: '#ffffff' });
    const scaleFactor = canvas.width / contentWidthPx;

    const pdf = new jsPDF({ unit: 'mm', format: 'a4' });
    const pageWidthMm = pdf.internal.pageSize.getWidth();
    const pageHeightMm = pdf.internal.pageSize.getHeight();
    const marginMm = 10;
    const usableHeightMm = pageHeightMm - marginMm * 2;
    const imgWidthMm = pageWidthMm;
    const pxToMm = imgWidthMm / contentWidthPx;
    const usableHeightPx = usableHeightMm / pxToMm;

    // Greedily pack units onto pages: whenever the next unit would push
    // the running page past its height budget, break before that unit
    // instead of mid-way through it.
    const breakpointsPx = [0];
    let pageStart = 0;
    unitPositions.forEach(pos => {
      if (pos.bottom - pageStart > usableHeightPx && pos.top > pageStart) {
        breakpointsPx.push(pos.top);
        pageStart = pos.top;
      }
    });
    breakpointsPx.push(totalHeightPx);

    for (let p = 0; p < breakpointsPx.length - 1; p++) {
      const sliceTopPx = breakpointsPx[p] * scaleFactor;
      const sliceBottomPx = breakpointsPx[p + 1] * scaleFactor;
      const sliceHeightPx = sliceBottomPx - sliceTopPx;
      if (sliceHeightPx <= 0) continue;

      const sliceCanvas = document.createElement('canvas');
      sliceCanvas.width = canvas.width;
      sliceCanvas.height = sliceHeightPx;
      const ctx = sliceCanvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, sliceCanvas.width, sliceCanvas.height);
      ctx.drawImage(canvas, 0, sliceTopPx, canvas.width, sliceHeightPx, 0, 0, canvas.width, sliceHeightPx);

      const sliceHeightMm = (sliceHeightPx / scaleFactor) * pxToMm;
      if (p > 0) pdf.addPage();
      pdf.addImage(sliceCanvas.toDataURL('image/jpeg', 0.92), 'JPEG', 0, marginMm, imgWidthMm, sliceHeightMm);
    }

    pdf.save(filename);
  } finally {
    if (container.parentNode) container.parentNode.removeChild(container);
  }
}

// Shared tail of both docx download handlers — packs a docx Document into
// a blob and triggers a browser download for it.
async function downloadDocxBlob(doc, filename) {
  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.setAttribute('href', url);
  link.setAttribute('download', filename);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

// Shared by the AI-report download in the modal header below — button+menu
// shape factored out on its own in case another download control needs it.
function DownloadDropdown({ label, isDownloading, downloadingLabel, wordLabel, pdfLabel, onWord, onPdf, open, onToggle, onClose }) {
  return (
    <div className="relative">
      <button
        onClick={onToggle}
        disabled={isDownloading}
        className="flex items-center gap-1.5 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-300 text-white text-sm font-medium px-3 py-1.5 rounded-lg transition-colors"
      >
        {isDownloading ? <Loader2 size={14} className="animate-spin" /> : <Download size={14} />}
        {isDownloading ? downloadingLabel : label}
        {!isDownloading && <ChevronDown size={14} />}
      </button>
      {open && !isDownloading && (
        <>
          <div className="fixed inset-0 z-10" onClick={onClose} />
          <div className="absolute right-0 mt-1 w-48 bg-white rounded-lg shadow-lg border border-slate-200 py-1 z-20">
            <button
              onClick={() => { onClose(); onWord(); }}
              className="w-full text-left px-3 py-2 text-sm text-slate-700 hover:bg-slate-50 transition-colors"
            >
              {wordLabel}
            </button>
            <button
              onClick={() => { onClose(); onPdf(); }}
              className="w-full text-left px-3 py-2 text-sm text-slate-700 hover:bg-slate-50 transition-colors"
            >
              {pdfLabel}
            </button>
          </div>
        </>
      )}
    </div>
  );
}

// ─── AI Analysis Modal ──────────────────────────────────────────────────────

function AIAnalysisModal({ geminiData, isAI, isUpgrading, hasApiKey, onUpgrade, onClose, onDownloadWord, onDownloadPdf, isDownloading, totalMatches, techById, onJumpToTech, isRevising, onRefine, versionCount, versionIndex, onGoToVersion, uiLang, analysisTechs }) {
  const [currentSlide, setCurrentSlide] = useState(0);
  const [showDownloadMenu, setShowDownloadMenu] = useState(false);
  const [focusTechId, setFocusTechId] = useState('');
  const [followUpText, setFollowUpText] = useState('');
  const t = MODAL_CHROME_LABELS[uiLang];

  const techDisplayName = (tech) => (uiLang === 'en' ? (tech.technology_name || tech.technology_name_zh) : (tech.technology_name_zh || tech.technology_name));

  const submitFollowUp = () => {
    if (!followUpText.trim() || isRevising) return;
    onRefine(followUpText.trim());
    setFollowUpText('');
  };

  // Guided quick-actions for the common cases (each sends a fully-formed
  // instruction straight to onRefine) sit alongside the free-text box below
  // rather than replacing it — the buttons cover the common cases without
  // typing, but anything more specific still needs the open input.
  const handleMoreConcise = () => {
    if (isRevising) return;
    onRefine(uiLang === 'en'
      ? 'Keep the same overall section structure, but make the whole report more concise — trim less-essential detail.'
      : '請將整份報告維持相同的章節結構，但內容寫得更精簡扼要，去除較次要的細節。');
  };

  const handleRewriteOtherLanguage = () => {
    if (isRevising) return;
    onRefine(uiLang === 'en'
      ? '請將整份報告改用繁體中文重寫，其餘結構保持不變。'
      : 'Please rewrite the whole report in English, keeping the same overall section structure.');
  };

  const handleFocusTech = () => {
    if (!focusTechId || isRevising) return;
    const tech = (analysisTechs || []).find(t2 => t2._id === focusTechId);
    if (!tech) return;
    const name = techDisplayName(tech);
    onRefine(uiLang === 'en'
      ? `Focus more deeply on "${name}" — give it a more thorough analysis, and you can shorten the coverage of the other technologies. Keep the same overall section structure.`
      : `請更聚焦在「${name}」，針對這項技術做更深入的分析與說明，其他技術的內容可以精簡帶過，但整體章節結構維持不變。`);
  };

  const slides = geminiData ? [
    geminiData.slide1,
    geminiData.slide2,
    geminiData.slide3,
    geminiData.slide4,
    geminiData.report,
  ].filter(Boolean) : [];

  const slideCount = slides.length || 4;

  const slideColors = ['#2563EB', '#059669', '#7C3AED', '#D97706', '#0D9488'];

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
              <h2 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                {t.title}
                <span className={`text-xs px-1.5 py-0.5 rounded-full font-bold ${isAI ? 'bg-purple-100 text-purple-700' : 'bg-slate-100 text-slate-600'}`}>
                  {isAI ? t.aiBadge : t.basicBadge}
                </span>
              </h2>
              <p className="text-sm text-slate-500">{t.subtitle(totalMatches)}</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            {/* Only shown while actually viewing the report page itself —
                the download always exports the report regardless of which
                slide the export/download button was clicked from, so
                showing it while browsing the rule-based slide1-4 stats
                would wrongly imply those pages are downloadable too. */}
            {geminiData?.report && currentSlide === slides.length - 1 && (
              <DownloadDropdown
                label={t.download}
                isDownloading={isDownloading}
                downloadingLabel={t.downloading}
                wordLabel={t.downloadWord}
                pdfLabel={t.downloadPdf}
                onWord={onDownloadWord}
                onPdf={onDownloadPdf}
                open={showDownloadMenu}
                onToggle={() => setShowDownloadMenu(v => !v)}
                onClose={() => setShowDownloadMenu(false)}
              />
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
                  className="flex items-center gap-1 text-sm font-medium px-3 py-1.5 rounded border border-slate-300 text-slate-600 hover:bg-slate-50 disabled:opacity-30 transition-colors"
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
                  className="flex items-center gap-1 text-sm font-medium px-3 py-1.5 rounded border border-slate-300 text-slate-600 hover:bg-slate-50 disabled:opacity-30 transition-colors"
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
                        <span className="text-sm text-slate-500 font-medium flex items-center gap-1.5">
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
                    <div className="space-y-2">
                      <div className="flex gap-2 items-center">
                        <select
                          value={focusTechId}
                          onChange={e => setFocusTechId(e.target.value)}
                          disabled={isRevising || !analysisTechs || analysisTechs.length === 0}
                          className="flex-1 text-base px-3 py-2 rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-purple-400 disabled:bg-slate-100 bg-white"
                        >
                          <option value="">{t.focusTechPlaceholder}</option>
                          {(analysisTechs || []).map(tech => (
                            <option key={tech._id} value={tech._id}>{techDisplayName(tech)}</option>
                          ))}
                        </select>
                        <button
                          onClick={handleFocusTech}
                          disabled={isRevising || !focusTechId}
                          className="flex items-center gap-1.5 bg-gradient-to-r from-purple-600 to-blue-600 hover:from-purple-700 hover:to-blue-700 disabled:from-slate-300 disabled:to-slate-300 disabled:cursor-not-allowed text-white text-base font-medium px-4 py-2 rounded-lg transition-all flex-shrink-0"
                        >
                          {isRevising ? <Loader2 size={14} className="animate-spin" /> : <Sparkles size={14} />}
                          {isRevising ? t.revising : t.focusTechButton}
                        </button>
                      </div>
                      <div className="flex gap-2 flex-wrap">
                        <button
                          onClick={handleMoreConcise}
                          disabled={isRevising}
                          className="text-sm font-medium px-3 py-1.5 rounded-full border border-purple-300 text-purple-700 bg-white hover:bg-purple-50 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
                        >
                          {t.moreConciseButton}
                        </button>
                        <button
                          onClick={handleRewriteOtherLanguage}
                          disabled={isRevising}
                          className="text-sm font-medium px-3 py-1.5 rounded-full border border-purple-300 text-purple-700 bg-white hover:bg-purple-50 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
                        >
                          {t.rewriteOtherLangButton}
                        </button>
                      </div>
                      <div className="flex gap-2 items-center">
                        <input
                          type="text"
                          value={followUpText}
                          onChange={e => setFollowUpText(e.target.value)}
                          onKeyDown={e => { if (e.key === 'Enter') submitFollowUp(); }}
                          placeholder={t.followUpPlaceholder}
                          disabled={isRevising}
                          className="flex-1 text-base px-3 py-2 rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-purple-400 disabled:bg-slate-100"
                        />
                        <button
                          onClick={submitFollowUp}
                          disabled={isRevising || !followUpText.trim()}
                          className="flex items-center gap-1.5 bg-gradient-to-r from-purple-600 to-blue-600 hover:from-purple-700 hover:to-blue-700 disabled:from-slate-300 disabled:to-slate-300 disabled:cursor-not-allowed text-white text-base font-medium px-4 py-2 rounded-lg transition-all flex-shrink-0"
                        >
                          {isRevising ? <Loader2 size={14} className="animate-spin" /> : <Sparkles size={14} />}
                          {isRevising ? t.revising : t.send}
                        </button>
                      </div>
                    </div>
                  </>
                ) : hasApiKey ? (
                  <div className="flex items-center justify-center">
                    <button
                      onClick={onUpgrade}
                      disabled={isUpgrading}
                      className="flex items-center gap-1.5 bg-gradient-to-r from-purple-600 to-blue-600 hover:from-purple-700 hover:to-blue-700 disabled:from-slate-300 disabled:to-slate-300 disabled:cursor-not-allowed text-white text-base font-medium px-4 py-2 rounded-lg transition-all"
                    >
                      {isUpgrading ? <Loader2 size={14} className="animate-spin" /> : <Sparkles size={14} />}
                      {isUpgrading ? t.upgrading : t.upgradeToAI}
                    </button>
                  </div>
                ) : (
                  <p className="text-sm text-amber-600 text-center">{t.needApiKeyForUpgrade}</p>
                )}
              </div>
            </>
          ) : (
            <div className="flex-1 flex items-center justify-center text-slate-400 text-base">
              {t.noPreview}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// One color+icon per slide1 KPI card (technologies / cases / countries, in
// that fixed order from buildFallbackAnalysis) — distinct per-card accents
// rather than all three sharing slide1's single accentColor, matching the
// "colored left border + watermark icon" stat-card look.
const KPI_CARD_STYLES = [
  { color: '#2563EB', Icon: Database },
  { color: '#059669', Icon: FileText },
  { color: '#7C3AED', Icon: Globe },
];

// ─── PPT Generation Guide Modal ───────────────────────────────────────────────

// A red pointer arrow drawn at a normalized (x%, y%) point over a
// screenshot, its tip landing exactly on that point regardless of the
// image's rendered size. `dir` picks which corner of the arrow's own
// bounding box is pinned to (x,y) — 'nw' anchors the box's bottom-right
// corner there (tail extends up-left), 'ne' anchors the bottom-left corner
// there (tail extends up-right) — pick whichever keeps the tail from
// running past the image edge for a given target.
//
// Deliberately short (an 18-unit stroke, not a full-box diagonal) and has
// no circle/ring marker at the tip — both were covering nearby button
// labels on the denser screenshots. The tip's position in the 52x52
// viewBox (42,42 for 'nw'; 10,42 for 'ne') is unchanged from a longer
// earlier version, only the tail moved closer to it, so the translate
// percentages that pin the tip to (x%, y%) — 80.8%/19.2% of the box's own
// size — still apply; don't recompute those if you resize the box itself.
//
// GUIDE_ARROW_COLOR (#FBEC31) is sampled directly from den-mascot.png's
// fur so the arrow reads as "part of DEN's palette" rather than a generic
// UI red. A pure bright yellow has weak contrast on the mostly-white
// screenshots though, so each stroke is drawn twice — a wider dark-slate
// pass underneath for a thin outline, then the gold pass on top — instead
// of just picking a darker gold that would drift away from the sampled hue.
const GUIDE_ARROW_COLOR = '#FBEC31';
const GUIDE_ARROW_OUTLINE = '#1E293B';
function GuideArrow({ x, y, dir = 'nw' }) {
  const isNe = dir === 'ne';
  const xShift = isNe ? -19.2 : -80.8;
  const geometry = isNe
    ? { line: { x1: 28, y1: 24, x2: 10, y2: 42 }, head: 'M10,42 L20,39 M10,42 L13,32' }
    : { line: { x1: 24, y1: 24, x2: 42, y2: 42 }, head: 'M42,42 L32,39 M42,42 L39,32' };
  return (
    <div
      className="absolute pointer-events-none"
      style={{ left: `${x}%`, top: `${y}%`, transform: `translate(${xShift}%, -80.8%)` }}
    >
      <svg width="52" height="52" viewBox="0 0 52 52" style={{ overflow: 'visible', display: 'block' }}>
        <g stroke={GUIDE_ARROW_OUTLINE} strokeWidth="5.5" strokeLinecap="round" strokeLinejoin="round" fill="none">
          <line {...geometry.line} />
          <path d={geometry.head} />
        </g>
        <g stroke={GUIDE_ARROW_COLOR} strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" fill="none">
          <line {...geometry.line} />
          <path d={geometry.head} />
        </g>
      </svg>
    </div>
  );
}

// Paginated walkthrough for the copy-article → NotebookLM source →
// copy-prompt → generate flow, built from real screenshots (see
// PPT_GUIDE_CONTENT) with arrows pointing at exactly what to click. The
// final step has no screenshot (nothing to point at while waiting).
function PptGuideModal({ isOpen, onClose, uiLang }) {
  const [step, setStep] = useState(0);
  if (!isOpen) return null;

  const t = PPT_GUIDE_CONTENT[uiLang];
  const steps = t.steps;
  const total = steps.length;
  const current = steps[step];

  const handleClose = () => {
    setStep(0);
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4" onClick={handleClose}>
      <div className="relative w-full max-w-xl" onClick={e => e.stopPropagation()}>
        <div className="bg-white rounded-2xl shadow-2xl overflow-hidden">
          <div className="flex items-center justify-between px-5 py-4 border-b border-slate-200 bg-slate-50">
            <h3 className="text-base font-bold text-slate-800">{t.title}</h3>
            <button onClick={handleClose} className="text-slate-400 hover:text-slate-600 transition-colors">
              <X size={20} />
            </button>
          </div>

          <div className="p-6 flex flex-col items-center text-center gap-3">
            {/* DEN rides along next to the step label/title on every
                screenshot step. The final (no-image) step already shows
                DEN front-and-center as the working animation below, so it
                doesn't get a second copy up here too. */}
            <div className="flex items-center justify-center gap-3">
              {current.image && (
                <img src="./guide/den-mascot.png" alt="DEN" className="w-11 h-11 object-contain flex-shrink-0" />
              )}
              <div className="text-left">
                <div className="text-xs font-semibold tracking-wide text-purple-600">{t.stepLabel(step + 1, total)}</div>
                <h4 className="text-lg font-bold text-slate-800">{current.title}</h4>
              </div>
            </div>
            <p className="text-base text-slate-600 leading-relaxed max-w-lg">{current.desc}</p>

            {current.image ? (
              <div className="relative inline-block mt-1 rounded-xl overflow-hidden border border-slate-200 shadow-sm max-h-80">
                <img src={current.image} alt={current.title} className="block max-h-80 w-auto" />
                {current.arrows.map((a, i) => (
                  <GuideArrow key={i} x={a.x} y={a.y} dir={a.dir} />
                ))}
              </div>
            ) : (
              <div className="flex items-center justify-center py-8">
                <img src="./guide/den-mascot.png" alt="DEN" className="w-24 h-24 object-contain animate-den-work" />
              </div>
            )}
          </div>

          <div className="flex items-center justify-between px-5 py-4 border-t border-slate-100">
            <button
              onClick={() => setStep(s => Math.max(0, s - 1))}
              disabled={step === 0}
              className="flex items-center gap-1 text-sm font-medium text-slate-500 disabled:opacity-30 disabled:cursor-not-allowed hover:text-slate-700 transition-colors"
            >
              <ChevronLeft size={16} /> {t.prev}
            </button>
            <div className="flex gap-1.5">
              {steps.map((_, i) => (
                <span key={i} className={`w-1.5 h-1.5 rounded-full transition-colors ${i === step ? 'bg-blue-600' : 'bg-slate-200'}`} />
              ))}
            </div>
            {step === total - 1 ? (
              <button onClick={handleClose} className="text-sm font-semibold text-blue-600 hover:text-blue-800 transition-colors">
                {t.done}
              </button>
            ) : (
              <button
                onClick={() => setStep(s => Math.min(total - 1, s + 1))}
                className="flex items-center gap-1 text-sm font-medium text-blue-600 hover:text-blue-800 transition-colors"
              >
                {t.next} <ChevronRight size={16} />
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── Slide Preview Component (White/Light style) ─────────────────────────────

function SlidePreview({ slide, index, totalSlides, accentColor, techById, onJumpToTech, uiLabels, uiLang }) {
  const [expandedTrlBucket, setExpandedTrlBucket] = useState(null);
  const [trendTechId, setTrendTechId] = useState(null);
  const [expandedCountry, setExpandedCountry] = useState(null);
  const labels = { ...DEFAULT_UI_LABELS[uiLang || 'zh'], ...(uiLabels || {}) };

  // Drilling into a TRL bucket/country only makes sense on the slide you
  // drilled into.
  useEffect(() => {
    setExpandedTrlBucket(null);
    setTrendTechId(null);
    setExpandedCountry(null);
  }, [index]);

  if (!slide) return null;

  const titles = ['查詢總覽', '技術成熟度分析', '專案活動熱區', '關鍵技術清單', '簡要報告'];
  const icons = ['📊', '🔬', '🌍', '🏆', '📝'];

  const activeBucket = slide.trl_breakdown && expandedTrlBucket !== null ? slide.trl_breakdown[expandedTrlBucket] : null;
  // Display name comes from Gemini's own "techs" field (so it follows
  // whatever language/rewrite the user asked for) — id still resolves the
  // real record for navigation/trend-chart data, which can't be translated.
  const activeBucketTechs = activeBucket ? (activeBucket.techs || []).filter(t => techById?.has(t.id)) : [];

  // Same drill-down pattern as the TRL bucket above — click a country bar
  // to reveal the cases behind it, same {name, count, initiatives} shape
  // computeTechStats produces.
  const activeCountry = slide.countries && expandedCountry !== null ? slide.countries[expandedCountry] : null;
  const activeCountryInitiatives = activeCountry ? (activeCountry.initiatives || []).filter(i => techById?.has(i.techId)) : [];

  return (
    <>
    <div
      className="w-full mx-auto rounded-xl shadow-lg border border-slate-200"
      style={{ minHeight: '393px', maxWidth: '700px', background: '#fff', position: 'relative' }}
    >
      {/* Top accent bar */}
      <div style={{ height: '6px', background: accentColor, width: '100%', borderRadius: '12px 12px 0 0' }} />

      {/* No fixed height / overflow-hidden here anymore — a hard-clipped
          16:9 box kept cutting off slide2/the report slide whenever their
          content (TRL drill-down, longer report sections + reference chips)
          grew past a "real slide" size. The card now just grows with its
          content, and the modal's own overflow-auto wrapper scrolls it. */}
      <div className="p-8 flex flex-col">
        {/* Slide number badge */}
        <div className="flex items-start justify-between mb-5">
          <div>
            <div
              className="text-sm font-bold px-2 py-0.5 rounded-full inline-block mb-2"
              style={{ background: accentColor + '18', color: accentColor }}
            >
              第 {index + 1} / {totalSlides || 4} 頁
            </div>
            <h2
              className="text-3xl font-bold text-slate-900 leading-tight"
            >
              {icons[index]} {slide.title || titles[index]}
            </h2>
            {slide.subtitle && (
              <p className="text-base text-slate-500 mt-1">{slide.subtitle}</p>
            )}
          </div>
          <div className="text-right flex-shrink-0">
            <div className="text-sm text-slate-300 font-mono">IEA NET ZERO</div>
            <div className="text-xs text-slate-300 mt-0.5 whitespace-nowrap">AI Generated · {new Date().getFullYear()}</div>
          </div>
        </div>

        {/* Divider */}
        <div style={{ width: '48px', height: '3px', background: accentColor, borderRadius: '2px', marginBottom: '16px' }} />

        {/* Content area */}
        <div className="flex-1 flex gap-5 min-h-0">
          {/* Main content */}
          <div className="flex-1 min-w-0">
            {slide.summary && (
              <p className="text-base text-slate-700 leading-relaxed mb-4">{slide.summary}</p>
            )}

            {/* KPIs (slide 1) — colored left border + big number + faint
                watermark icon + bottom caption, one accent per card. */}
            {slide.kpis && Array.isArray(slide.kpis) && (
              <div className="grid grid-cols-3 gap-3 mt-2">
                {slide.kpis.map((kpi, i) => {
                  const style = KPI_CARD_STYLES[i % KPI_CARD_STYLES.length];
                  const CardIcon = style.Icon;
                  return (
                    <div
                      key={i}
                      className="relative overflow-hidden rounded-lg p-3 bg-white border border-slate-200"
                      style={{ borderLeftWidth: '4px', borderLeftColor: style.color }}
                    >
                      <CardIcon size={56} className="absolute -right-3 -bottom-3 opacity-[0.08] pointer-events-none" style={{ color: style.color }} />
                      <div className="relative text-2xl font-black" style={{ color: style.color }}>{kpi.value}</div>
                      <div className="relative text-sm text-slate-500 mt-0.5">{kpi.label}</div>
                      {kpi.caption && (
                        <div className="relative text-xs text-slate-400 mt-1.5 pt-1.5 border-t border-slate-100">{kpi.caption}</div>
                      )}
                    </div>
                  );
                })}
              </div>
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
                      <div className="text-sm w-28 flex-shrink-0" style={{ color: isActive ? accentColor : '#64748b', fontWeight: isActive ? 700 : 400 }}>{item.stage}</div>
                      <div className="flex-1 h-5 rounded-full bg-slate-100 overflow-hidden">
                        <div
                          className="h-full rounded-full transition-all duration-700"
                          style={{ width: `${Math.min(100, (item.count / (slide.total || 1)) * 100)}%`, background: item.count > 0 ? accentColor : accentColor + '80' }}
                        />
                      </div>
                      <div className="text-sm font-bold text-slate-700 w-8 text-right">{item.count}</div>
                    </button>
                  );
                })}
                {slide.trl_breakdown.some(b => b.count > 0) && (
                  <p className="text-xs text-slate-400 pt-1">{labels.trlBarHint}</p>
                )}
              </div>
            )}

            {/* Countries (slide 3) — bar chart proportional to case count,
                same visual language and click-to-drill-down pattern as the
                TRL bars above; drill-down list & jump-to-tech render in a
                separate panel below the slide (see return), same reason
                the TRL drill-down does. */}
            {slide.countries && Array.isArray(slide.countries) && (
              slide.countries.length === 0 ? (
                <p className="text-sm text-slate-400">{uiLang === 'en' ? 'Not enough case data to identify hotspots.' : '案例資料不足，無法識別活躍地區。'}</p>
              ) : (() => {
                const maxCount = Math.max(...slide.countries.map(x => x.count || 0), 1);
                return (
                  <div className="space-y-2">
                    {slide.countries.map((c, i) => {
                      const isActive = expandedCountry === i;
                      return (
                        <button
                          key={i}
                          type="button"
                          onClick={() => { setExpandedCountry(isActive ? null : i); }}
                          disabled={!c.count}
                          className={`w-full flex items-center gap-3 text-left disabled:cursor-default rounded-lg -mx-1 px-1 py-0.5 transition-colors ${isActive ? 'bg-slate-50' : ''}`}
                        >
                          <div className="text-sm w-28 flex-shrink-0 truncate" style={{ color: isActive ? accentColor : '#64748b', fontWeight: isActive ? 700 : 400 }}>{c.name}</div>
                          <div className="flex-1 h-5 rounded-full bg-slate-100 overflow-hidden">
                            <div
                              className="h-full rounded-full transition-all duration-700"
                              style={{ width: `${Math.min(100, (c.count / maxCount) * 100)}%`, background: accentColor }}
                            />
                          </div>
                          <div className="text-sm font-bold text-slate-700 w-8 text-right">{c.count}</div>
                        </button>
                      );
                    })}
                    <p className="text-xs text-slate-400 pt-1">{labels.countryBarHint}</p>
                  </div>
                );
              })()
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
                        className={`w-full flex items-center gap-3 text-base text-left rounded-lg -mx-1 px-1 py-0.5 transition-colors ${jumpable ? 'hover:bg-slate-50 cursor-pointer' : 'cursor-default'}`}
                      >
                        <span
                          className="w-6 h-6 rounded-full flex items-center justify-center text-sm font-bold flex-shrink-0 text-white"
                          style={{ background: accentColor }}
                        >
                          {i + 1}
                        </span>
                        <span className={`text-slate-800 font-medium flex-1 truncate ${jumpable ? 'underline decoration-dotted underline-offset-2' : ''}`}>{tech.name}</span>
                        <span className="text-sm text-slate-400">{tech.count} {labels.caseCountSuffix}</span>
                      </button>
                    </li>
                  );
                })}
              </ol>
            )}

            {/* Briefing report (slide "report") — a NotebookLM-style briefing
                document: executive summary, a full one-row-per-technology
                overview table, a deep-dive section for every matched
                technology, and a closing cross-technology trends section. */}
            {slide.executive_summary && (
              <div
                className="rounded-lg p-4 mb-5 text-base text-slate-700 leading-relaxed"
                style={{ background: accentColor + '0c', borderLeft: `3px solid ${accentColor}` }}
              >
                <div className="text-sm font-bold mb-1" style={{ color: accentColor }}>
                  {uiLang === 'en' ? 'Executive Summary' : '摘要'}
                </div>
                <p>{slide.executive_summary}</p>
              </div>
            )}

            {slide.overview_table?.rows?.length > 0 && (
              <div className="mb-6 overflow-x-auto">
                <table className="w-full text-sm border-collapse">
                  <thead>
                    <tr className="border-b-2" style={{ borderColor: accentColor }}>
                      <th className="text-left py-2 pr-3 font-bold text-slate-700">{uiLang === 'en' ? 'Technology' : '技術類別'}</th>
                      <th className="text-left py-2 pr-3 font-bold text-slate-700">{uiLang === 'en' ? 'Energy Source' : '能源來源'}</th>
                      <th className="text-left py-2 pr-3 font-bold text-slate-700">TRL</th>
                      <th className="text-left py-2 font-bold text-slate-700">{uiLang === 'en' ? 'Market Dynamics' : '主要市場動態'}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {slide.overview_table.rows.map((row, i) => {
                      const jumpable = !!(row.id && techById?.has(row.id) && onJumpToTech);
                      return (
                        <tr key={i} className="border-b border-slate-100 align-top">
                          <td className="py-2 pr-3 font-medium text-slate-800">
                            {jumpable ? (
                              <button type="button" onClick={() => onJumpToTech(row.id)} title={uiLang === 'en' ? 'Click to view details' : '點擊查看此技術詳情'} className="text-left underline decoration-dotted underline-offset-2 hover:brightness-90" style={{ color: accentColor }}>
                                {row.name}
                              </button>
                            ) : row.name}
                          </td>
                          <td className="py-2 pr-3 text-slate-600">{row.energy_source}</td>
                          <td className="py-2 pr-3 text-slate-600 whitespace-nowrap">{row.trl}</td>
                          <td className="py-2 text-slate-600">{row.market_dynamics}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}

            {slide.tech_sections && Array.isArray(slide.tech_sections) && slide.tech_sections.length > 0 && (
              <div className="space-y-6 mb-6">
                {slide.tech_sections.map((section, i) => {
                  const jumpable = !!(section.id && techById?.has(section.id) && onJumpToTech);
                  return (
                    <div key={i}>
                      <h3 className="text-lg font-bold text-slate-800 mb-1">
                        {i + 1}.{' '}
                        {jumpable ? (
                          <button type="button" onClick={() => onJumpToTech(section.id)} title={uiLang === 'en' ? 'Click to view details' : '點擊查看此技術詳情'} className="underline decoration-dotted underline-offset-2 hover:brightness-90">
                            {section.heading}
                          </button>
                        ) : section.heading}
                      </h3>
                      {section.intro && <p className="text-base text-slate-700 leading-relaxed mb-2">{section.intro}</p>}
                      {(section.subsections || []).map((sub, j) => (
                        <div key={j} className="mt-2">
                          <div className="text-sm font-bold" style={{ color: accentColor }}>{sub.label}</div>
                          <ul className="mt-1 space-y-1">
                            {(sub.bullets || []).map((bullet, k) => (
                              <li key={k} className="flex items-start gap-2 text-base text-slate-700">
                                <span style={{ color: accentColor, marginTop: '2px', flexShrink: 0 }}>▸</span>
                                <span className="leading-relaxed">{typeof bullet === 'string' ? bullet : bullet.text}</span>
                              </li>
                            ))}
                          </ul>
                        </div>
                      ))}
                    </div>
                  );
                })}
              </div>
            )}

            {slide.trends_section && (
              <div>
                {slide.trends_section.heading && <h3 className="text-lg font-bold text-slate-800 mb-2">{slide.trends_section.heading}</h3>}
                {(slide.trends_section.subsections || []).map((sub, i) => (
                  <div key={i} className="mt-3">
                    <div className="text-sm font-bold" style={{ color: accentColor }}>{sub.label}</div>
                    <ul className="mt-1 space-y-2">
                      {(sub.bullets || []).map((bullet, j) => {
                        const text = typeof bullet === 'string' ? bullet : bullet.text;
                        const relatedTechs = ((typeof bullet === 'string' ? [] : bullet.related_techs) || []).filter(t => techById?.has(t.id));
                        return (
                          <li key={j} className="flex items-start gap-2 text-base text-slate-700">
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
                                      className="text-xs px-2 py-0.5 rounded-full font-medium transition-colors hover:brightness-95"
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
                  </div>
                ))}
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
            <h4 className="text-base font-bold text-slate-800">
              {activeBucket.stage}　{labels.trlListHeading}（共 {activeBucket.count} {labels.countUnit}）
            </h4>
            <button
              type="button"
              onClick={() => { setExpandedTrlBucket(null); setTrendTechId(null); }}
              className="text-sm text-slate-400 hover:text-slate-600 px-2 py-1 rounded hover:bg-slate-100 transition-colors"
            >
              {labels.collapseButton}
            </button>
          </div>

          {activeBucketTechs.length === 0 ? (
            <p className="text-sm text-slate-400 py-2">{labels.trlListEmpty}</p>
          ) : (
            <>
              <p className="text-xs text-slate-400 mb-2">{labels.trlTechHint}</p>
              <div className="flex flex-wrap gap-2">
              {activeBucketTechs.map(t => {
                const isSelected = trendTechId === t.id;
                return (
                  <button
                    key={t.id}
                    type="button"
                    onClick={() => setTrendTechId(isSelected ? null : t.id)}
                    className="text-sm px-3 py-1.5 rounded-full border font-medium transition-colors"
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

    {/* Country drill-down panel — same reasoning as the TRL panel above:
        outside the fixed-aspect slide card so it can grow freely. Each
        case jumps straight to its technology (via onJumpToTech) since
        initiatives don't have a page of their own to link to. */}
    {activeCountry && (
      <div className="w-full mx-auto mt-4" style={{ maxWidth: '700px' }}>
        <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-4">
          <div className="flex items-center justify-between mb-3">
            <h4 className="text-base font-bold text-slate-800">
              {activeCountry.name}　{labels.countryListHeading}（共 {activeCountry.count} {labels.countUnit}）
            </h4>
            <button
              type="button"
              onClick={() => setExpandedCountry(null)}
              className="text-sm text-slate-400 hover:text-slate-600 px-2 py-1 rounded hover:bg-slate-100 transition-colors"
            >
              {labels.collapseButton}
            </button>
          </div>

          {activeCountryInitiatives.length === 0 ? (
            <p className="text-sm text-slate-400 py-2">{labels.countryListEmpty}</p>
          ) : (
            <>
              <p className="text-xs text-slate-400 mb-2">{labels.countryCaseHint}</p>
              <div className="space-y-2">
                {activeCountryInitiatives.map((init, idx) => (
                  <button
                    key={idx}
                    type="button"
                    onClick={() => onJumpToTech(init.techId)}
                    className="w-full text-left border-l-2 pl-4 py-2 rounded-r-lg hover:bg-slate-50 transition-colors"
                    style={{ borderColor: accentColor }}
                  >
                    <div className="flex flex-wrap items-center gap-2 mb-1 text-sm font-medium text-slate-500">
                      {init.year && <span className="bg-slate-100 px-1.5 py-0.5 rounded border border-slate-200">{init.year}</span>}
                      {init.type && <span style={{ color: accentColor }}>{init.type}</span>}
                      <span className="font-semibold underline decoration-dotted underline-offset-2" style={{ color: accentColor }}>{init.techName}</span>
                    </div>
                    <p className="text-sm text-slate-700 leading-relaxed">{init.description || (uiLang === 'en' ? 'Unlabeled' : '未標示')}</p>
                  </button>
                ))}
              </div>
            </>
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
    return <p className="text-sm text-slate-400 mt-2 py-1">{insufficientDataText || DEFAULT_UI_LABELS.zh.trendInsufficientData}</p>;
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
      <p className="text-sm font-semibold mb-1" style={{ color: accentColor }}>
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
  // Set when the user clicks a segment of a technology's full sector-path
  // breadcrumb (in the detail view) — an array of the *English* sector[]
  // values up to and including the clicked segment, e.g. ['large-scale or
  // industrial', 'Industry', 'Aluminium']. Kept separate from selectedSector
  // (which does a loose substring match against a flattened bag of sector
  // text) because a path click needs an exact, ordered prefix match against
  // tech.sector — clicking "Aluminium" must not also pull in an unrelated
  // "Aluminium" appearing at a different depth under a different top
  // category. See searchData below for the matching logic.
  const [pathFilter, setPathFilter] = useState(null);
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
      // Key goes in the x-goog-api-key header (officially supported), not a
      // ?key= query param — URLs get recorded in proxy logs and browser
      // history, headers don't.
      const response = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json", "x-goog-api-key": key },
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
  // ever owns the report, so every AI turn (upgrade or follow-up) merges
  // its partial response back onto that unchanging foundation rather than
  // being trusted to reproduce the stats slides itself.
  const mergeWithBaseSlides = (aiPartial) => {
    const base = versions.list[0]?.geminiData || {};
    return {
      slide1: base.slide1, slide2: base.slide2, slide3: base.slide3, slide4: base.slide4,
      report: aiPartial.report,
      ui_labels: base.ui_labels
    };
  };

  // Step two, triggered from inside the modal: adds the AI-only briefing
  // report (executive summary + themed sections) alongside the untouched
  // rule-based slide1-4, in place, without closing or re-opening anything.
  const handleUpgradeToAI = async () => {
    if (!apiKey.trim() || analysisTechs.length === 0) return;

    setIsUpgrading(true);

    // See buildFullDataCsv for exactly which fields this includes and why
    // — every linked case record (for real project names/figures to cite),
    // trimmed of the columns the prompt never actually reads.
    const csvStr = buildFullDataCsv(analysisTechs);

    const promptText = buildBriefingReportPrompt(csvStr, uiLang);

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
  // its own prior report answer in context. Scoped to the report only, same
  // as the initial upgrade.
  const handleRefineAnalysis = async (instruction) => {
    if (!geminiData || chatHistory.length === 0 || !instruction.trim() || !apiKey.trim()) return;

    setIsRevising(true);
    const currentLanguageNote = uiLang === 'en'
      ? '目前平台介面是英文模式，除非這個指示明確要求換成別的語言，否則請繼續全部用英文輸出。'
      : '目前平台介面是繁體中文模式，除非這個指示明確要求換成別的語言，否則請繼續全部用繁體中文輸出。';
    const newTurn = {
      role: 'user',
      parts: [{ text: `請根據以下指示調整剛才產生的 report JSON，維持完全相同的結構，只輸出純 JSON，不要 markdown 或說明文字，也不要輸出 slide1~slide4（那些不歸你管）。${currentLanguageNote}：\n${instruction}` }]
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

  // ── Report Download (Word) ────────────────────────────────────────────────
  // Deliberately scoped to ONLY the AI-generated report (executive summary +
  // overview table + tech deep-dives + trends) — the rule-based slide1-4
  // stats are for on-screen viewing only, not part of this export. Renders
  // as a real Word document via the "docx" library rather than a slide
  // deck, since a multi-section written report reads better as a document.
  const handleDownloadReportDocx = async () => {
    const report = geminiData?.report;
    if (!report) return;
    setIsDownloading(true);
    try {
      const isEn = uiLang === 'en';
      const children = [];

      children.push(new Paragraph({
        text: report.title || (isEn ? 'AI Insight Report' : 'AI 洞察報告'),
        heading: HeadingLevel.TITLE
      }));

      if (report.executive_summary) {
        children.push(new Paragraph({ text: isEn ? 'Executive Summary' : '摘要', heading: HeadingLevel.HEADING_1, spacing: { before: 300 } }));
        children.push(new Paragraph({ text: report.executive_summary, spacing: { after: 200 } }));
      }

      const tableRows = report.overview_table?.rows;
      if (Array.isArray(tableRows) && tableRows.length > 0) {
        children.push(new Paragraph({ text: isEn ? 'Technology Overview' : '技術現況總覽表', heading: HeadingLevel.HEADING_1, spacing: { before: 300 } }));
        const headerLabels = isEn
          ? ['Technology', 'Energy Source', 'TRL', 'Market Dynamics']
          : ['技術類別', '能源來源', 'TRL', '主要市場動態'];
        const headerRow = new TableRow({
          tableHeader: true,
          children: headerLabels.map(h => new TableCell({
            width: { size: 25, type: WidthType.PERCENTAGE },
            shading: { fill: '1F2937' },
            children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, color: 'FFFFFF' })] })]
          }))
        });
        const bodyRows = tableRows.map(row => new TableRow({
          children: [row.name, row.energy_source, row.trl, row.market_dynamics].map(val => new TableCell({
            width: { size: 25, type: WidthType.PERCENTAGE },
            children: [new Paragraph({ text: String(val || '') })]
          }))
        }));
        children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow, ...bodyRows] }));
        // A blank paragraph after the table — Word renders content
        // immediately butted against a table's bottom edge otherwise.
        children.push(new Paragraph({ text: '' }));
      }

      const techSections = report.tech_sections;
      if (Array.isArray(techSections) && techSections.length > 0) {
        children.push(new Paragraph({ text: isEn ? 'In-Depth Technology Analysis' : '核心技術深度分析', heading: HeadingLevel.HEADING_1, spacing: { before: 400 } }));
        techSections.forEach((section, i) => {
          children.push(new Paragraph({ text: `${i + 1}. ${section.heading || ''}`, heading: HeadingLevel.HEADING_2, spacing: { before: 300 } }));
          if (section.intro) {
            children.push(new Paragraph({ children: [new TextRun({ text: section.intro, italics: true })], spacing: { after: 100 } }));
          }
          (section.subsections || []).forEach(sub => {
            children.push(new Paragraph({ children: [new TextRun({ text: sub.label || '', bold: true })], spacing: { before: 150 } }));
            (sub.bullets || []).forEach(bullet => {
              const text = typeof bullet === 'string' ? bullet : bullet.text;
              children.push(new Paragraph({ text, bullet: { level: 0 } }));
            });
          });
        });
      }

      if (report.trends_section) {
        children.push(new Paragraph({
          text: report.trends_section.heading || (isEn ? 'Key Trends & Strategic Insights' : '關鍵趨勢與戰略洞察'),
          heading: HeadingLevel.HEADING_1, spacing: { before: 400 }
        }));
        (report.trends_section.subsections || []).forEach(sub => {
          children.push(new Paragraph({ children: [new TextRun({ text: sub.label || '', bold: true })], spacing: { before: 150 } }));
          (sub.bullets || []).forEach(bullet => {
            const text = typeof bullet === 'string' ? bullet : bullet.text;
            const relatedTechs = (typeof bullet === 'string' ? [] : (bullet.related_techs || [])).filter(t => techById.has(t.id));
            const relatedNames = relatedTechs.map(t => t.name);
            const runs = [new TextRun({ text })];
            if (relatedNames.length > 0) {
              const relatedPrefix = geminiData?.ui_labels?.relatedPrefix || DEFAULT_UI_LABELS[uiLang]?.relatedPrefix || DEFAULT_UI_LABELS.zh.relatedPrefix;
              const separator = isEn ? ': ' : '：';
              const joiner = isEn ? ', ' : '、';
              runs.push(new TextRun({ text: `  (${relatedPrefix}${separator}${relatedNames.join(joiner)})`, italics: true, color: '64748B' }));
            }
            children.push(new Paragraph({ children: runs, bullet: { level: 0 } }));
          });
        });
      }

      const doc = new Document({ sections: [{ children }] });
      const safeTitle = String(report.title || (isEn ? 'AI_Insight_Report' : 'AI_洞察報告')).replace(/[\\/:*?"<>|]/g, '').slice(0, 60);
      await downloadDocxBlob(doc, `${safeTitle}.docx`);
    } catch (error) {
      alert("下載報告時發生錯誤: " + error.message);
    } finally {
      setIsDownloading(false);
    }
  };

  // ── Report Download (PDF) ─────────────────────────────────────────────────
  // Same report-only scope as the Word export above, but rendered by
  // mounting buildReportHtml off-screen and rasterizing it (html2canvas)
  // into a jsPDF document via the shared renderHtmlToPdf, instead of
  // drawing text with jsPDF's native APIs — see buildReportHtml's comment
  // for why (no practical way to embed a Traditional Chinese font in
  // jsPDF here).
  const handleDownloadReportPdf = async () => {
    const report = geminiData?.report;
    if (!report) return;
    setIsDownloading(true);
    try {
      const isEn = uiLang === 'en';
      const relatedPrefix = geminiData?.ui_labels?.relatedPrefix || DEFAULT_UI_LABELS[uiLang]?.relatedPrefix || DEFAULT_UI_LABELS.zh.relatedPrefix;
      const html = buildReportHtml(report, uiLang, techById, relatedPrefix);
      const safeTitle = String(report.title || (isEn ? 'AI_Insight_Report' : 'AI_洞察報告')).replace(/[\\/:*?"<>|]/g, '').slice(0, 60);
      await renderHtmlToPdf(html, `${safeTitle}.pdf`);
    } catch (error) {
      alert("下載報告時發生錯誤: " + error.message);
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
        pathFilter={pathFilter}
        setPathFilter={setPathFilter}
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
          onDownloadWord={handleDownloadReportDocx}
          onDownloadPdf={handleDownloadReportPdf}
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
          analysisTechs={analysisTechs}
        />
      )}

      {/* Jumping to a technology closes the modal; this lets you get back to the
          already-generated slides instantly without re-calling Gemini. */}
      {!showAIModal && geminiData && (
        <button
          onClick={() => setShowAIModal(true)}
          className="fixed bottom-5 right-5 z-40 flex items-center gap-2 bg-gradient-to-r from-purple-600 to-blue-600 hover:from-purple-700 hover:to-blue-700 text-white text-base font-medium pl-4 pr-5 py-2.5 rounded-full shadow-lg transition-all"
        >
          <Sparkles size={16} />
          {APP_LABELS[uiLang].backToAiSlides}
        </button>
      )}
    </>
  );
}

// ─── Loading Skeleton ────────────────────────────────────────────────────────
// Shown in place of the real left/right panels only during the very first
// background fetch (isSampleData still true, status still 'loading') — the
// app initializes rows from a 5-item embedded sample dataset, so without
// this the user would otherwise see a real, clickable search UI whose
// results come from those 5 fake rows for the few seconds it takes to pull
// down the real ~4.7MB dataset, then have it all pop/swap out from under
// them. Static, non-interactive placeholders sidestep both problems.

function SkeletonBlock({ className }) {
  return <div className={`bg-slate-200 rounded animate-pulse ${className || ''}`} />;
}

function SkeletonListItem({ nameWidth }) {
  return (
    <div className="p-2.5 flex flex-col gap-2">
      <div className="flex justify-between items-start gap-2">
        <SkeletonBlock className={`h-4 ${nameWidth}`} />
        <SkeletonBlock className="h-4 w-12 flex-shrink-0" />
      </div>
      <SkeletonBlock className="h-3 w-2/3" />
      <div className="flex justify-between items-center mt-0.5">
        <SkeletonBlock className="h-3 w-16" />
        <SkeletonBlock className="h-3 w-10" />
      </div>
    </div>
  );
}

const SKELETON_LIST_WIDTHS = ['w-3/4', 'w-2/3', 'w-4/5', 'w-1/2', 'w-3/5', 'w-2/3', 'w-3/4', 'w-1/2'];

function SkeletonMain() {
  return (
    <main className="flex-1 flex flex-col md:flex-row overflow-hidden min-h-0" aria-busy="true">
      {/* Left panel skeleton — same widths as the real left panel section
          so nothing visibly shifts when the real content swaps in. */}
      <section className="w-full md:w-[400px] lg:w-[460px] flex-shrink-0 border-b md:border-b-0 md:border-r border-slate-200 bg-white flex flex-col h-1/2 md:h-full min-h-0">
        <div className="p-4 border-b border-slate-100 bg-slate-50 flex-shrink-0 space-y-3">
          <div className="flex gap-2">
            <SkeletonBlock className="h-8 flex-1 rounded-md" />
            <SkeletonBlock className="h-8 flex-1 rounded-md" />
          </div>
          <SkeletonBlock className="h-9 w-full rounded-lg" />
          <SkeletonBlock className="h-10 w-full rounded-lg" />
          <SkeletonBlock className="h-3 w-4/5" />
        </div>
        <div className="flex-1 overflow-hidden p-2 space-y-1">
          {SKELETON_LIST_WIDTHS.map((w, i) => <SkeletonListItem key={i} nameWidth={w} />)}
        </div>
      </section>

      {/* Right panel skeleton */}
      <section className="flex-1 flex flex-col bg-slate-50 overflow-hidden h-1/2 md:h-full min-h-0">
        <div className="flex items-center gap-6 border-b border-slate-200 bg-white px-4 md:px-6 py-3 flex-shrink-0">
          <SkeletonBlock className="h-5 w-20" />
          <SkeletonBlock className="h-5 w-28" />
        </div>
        <div className="flex-1 flex items-center justify-center p-6">
          <div className="w-full max-w-md space-y-3">
            <SkeletonBlock className="h-6 w-2/3 mx-auto" />
            <SkeletonBlock className="h-4 w-1/2 mx-auto" />
          </div>
        </div>
      </section>
    </main>
  );
}

// Hand-rolled SVG donut (no charting library) — stacks colored arcs via
// strokeDasharray/strokeDashoffset on concentric circles of the same
// radius, rotated -90deg so the first segment starts at 12 o'clock. Used in
// the Strategy tab's sector-distribution card: no static legend — hovering
// a segment previews it (thicker ring + center label swap); clicking a
// segment locks that preview in place until either that same segment is
// clicked again or the click lands elsewhere within the chart (the empty
// hole, not a segment). While locked, hovering a *different* segment still
// shows a temporary preview of that one and reverts back to the lock on
// mouse-leave, not to the total. pointerEvents:'stroke' is set explicitly
// because a fill="none" ring otherwise only reports pointer events
// inconsistently across browsers on its empty interior. onLockChange (if
// given) fires with the locked segment or null, so the parent card can
// show a summary underneath while a segment is locked.
function DonutChart({ segments, centerValue, centerLabel, size = 132, strokeWidth = 18, onLockChange }) {
  const [hoveredIndex, setHoveredIndex] = useState(null);
  const [lockedIndex, setLockedIndex] = useState(null);
  const radius = (size - strokeWidth) / 2;
  const circumference = 2 * Math.PI * radius;
  const total = segments.reduce((sum, s) => sum + s.value, 0) || 1;
  let cumulative = 0;

  const displayIndex = hoveredIndex !== null ? hoveredIndex : lockedIndex;
  const displayed = displayIndex !== null ? segments[displayIndex] : null;

  const applyLock = (idx) => {
    setLockedIndex(idx);
    if (onLockChange) onLockChange(idx !== null ? segments[idx] : null);
  };

  return (
    <div
      className="relative flex-shrink-0"
      style={{ width: size, height: size }}
      onClick={() => applyLock(null)}
    >
      <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} style={{ transform: 'rotate(-90deg)', overflow: 'visible' }}>
        <circle cx={size / 2} cy={size / 2} r={radius} fill="none" stroke="#F1F5F9" strokeWidth={strokeWidth} />
        {segments.map((seg, i) => {
          const fraction = seg.value / total;
          const dash = fraction * circumference;
          const offset = -cumulative * circumference;
          cumulative += fraction;
          const isActive = displayIndex === i;
          const isDimmed = displayIndex !== null && !isActive;
          return (
            <circle
              key={i}
              cx={size / 2} cy={size / 2} r={radius}
              fill="none"
              stroke={seg.color}
              strokeWidth={isActive ? strokeWidth + 4 : strokeWidth}
              strokeDasharray={`${dash} ${circumference - dash}`}
              strokeDashoffset={offset}
              style={{
                pointerEvents: 'stroke',
                cursor: 'pointer',
                opacity: isDimmed ? 0.4 : 1,
                transition: 'stroke-width 0.15s ease, opacity 0.15s ease, filter 0.15s ease',
                filter: isActive ? `drop-shadow(0 0 5px ${seg.color}) drop-shadow(0 0 10px ${seg.color}80)` : 'none'
              }}
              onMouseEnter={() => setHoveredIndex(i)}
              onMouseLeave={() => setHoveredIndex(null)}
              onClick={(e) => { e.stopPropagation(); applyLock(lockedIndex === i ? null : i); }}
            />
          );
        })}
      </svg>
      <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none px-4 text-center">
        {displayed ? (
          <>
            <div className="text-3xl font-black text-slate-800">{displayed.value}</div>
            <div className="text-base text-slate-500 leading-snug line-clamp-2">{displayed.label}</div>
          </>
        ) : (
          <>
            <div className="text-4xl font-black text-slate-800">{centerValue}</div>
            <div className="text-sm text-slate-400">{centerLabel}</div>
          </>
        )}
      </div>
    </div>
  );
}

// Purely decorative header background — a slowly spinning wind turbine
// silhouette and (used alongside it in the header) a scrolling wave strip,
// both on-theme for a clean-energy platform rather than generic tech
// particle effects. Rendered at low opacity, pointer-events-none, and
// behind the real header content via z-index — decoration, not content.
function WindTurbineIcon({ className, style, duration }) {
  return (
    <svg viewBox="0 0 100 140" className={className} style={style} aria-hidden="true">
      <line x1="50" y1="60" x2="50" y2="140" stroke="white" strokeWidth="3" />
      <g className="animate-spin-slow" style={{ animationDuration: duration || '9s' }}>
        <ellipse cx="50" cy="32" rx="6" ry="28" fill="white" />
        <ellipse cx="50" cy="32" rx="6" ry="28" fill="white" transform="rotate(120 50 60)" />
        <ellipse cx="50" cy="32" rx="6" ry="28" fill="white" transform="rotate(240 50 60)" />
      </g>
      <circle cx="50" cy="60" r="4" fill="white" />
    </svg>
  );
}

// ─── Dashboard ───────────────────────────────────────────────────────────────

function Dashboard({
  dataset, rows, isSampleData, appState, onUpload,
  isCustomUpload, onRestoreDefault,
  query, setQuery, searchMode, setSearchMode,
  selectedTech, setSelectedTech,
  selectedSector, setSelectedSector,
  pathFilter, setPathFilter,
  onOpenStrategyOverview,
  apiKey, onApiKeyChange,
  aiProvider, onProviderChange,
  openRouterModel, onOpenRouterModelChange,
  rightPanelTab, setRightPanelTab,
  uiLang, setUiLang
}) {
  const L = APP_LABELS[uiLang];
  // True only for the very first background fetch — see SkeletonMain's
  // comment for why this specific condition (not just appState.status)
  // matters.
  const isInitialLoading = isSampleData && appState.status === 'loading';
  const [showAllInitiatives, setShowAllInitiatives] = useState(false);
  const [showApiKey, setShowApiKey] = useState(false);
  const [currentPage, setCurrentPage] = useState(1);
  const [copiedArticleKey, setCopiedArticleKey] = useState(null);
  const [showPptGuide, setShowPptGuide] = useState(false);
  // Search header collapse — folds the mode toggle / sector dropdown /
  // search input away so the technology list below gets the vertical space.
  const [searchPanelCollapsed, setSearchPanelCollapsed] = useState(false);
  // Per-technology checkbox selection (by _id). When any matched techs are
  // checked, the Strategy tab (overview modal, donut, exports) runs on just
  // the checked subset instead of every matched result.
  const [checkedTechIds, setCheckedTechIds] = useState(() => new Set());
  const listContainerRef = useRef(null);
  const copyResetTimerRef = useRef(null);

  const toggleTechChecked = (id) => {
    setCheckedTechIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };

  const techCount = rows.length;
  const initiativeCount = dataset.source_stats?.initiative_rows ??
    rows.reduce((sum, tech) => sum + (Array.isArray(tech.initiatives) ? tech.initiatives.length : 0), 0);

  // Global sector_en -> sector_zh leaf lookup, for translating a pathFilter
  // segment when displaying the active-filter chip trail — pathFilter is
  // just an array of English strings (no longer tied to the tech it was
  // clicked from), and its last segment may be a leaf or an intermediate
  // category depending on which breadcrumb chip the user clicked.
  const sectorLeafZhMap = useMemo(() => {
    const m = {};
    rows.forEach(t => { if (t.sector_en && t.sector_zh) m[t.sector_en] = t.sector_zh; });
    return m;
  }, [rows]);
  const pathSegmentZh = (seg) => sectorLeafZhMap[seg] || SECTOR_TRANSLATIONS[seg] || seg;

  // Category label used by the Strategy tab's sector-distribution donut
  // (and the locked-segment stats/definition below it). Once pathFilter
  // narrows the result set past the top level, every remaining technology
  // already shares that same prefix — grouping by each tech's own final
  // leaf would still work, but grouping by the *next* segment past the
  // clicked prefix is what actually reveals a breakdown within the current
  // filter instead of a single 100%-share slice repeating the filter
  // itself. Falls back to the true leaf (sector_zh/sector_en, unchanged
  // from before pathFilter existed) whenever pathFilter is inactive, only
  // one level deep (nothing meaningful to unify past yet), or a given tech
  // simply has no path left beyond the clicked prefix.
  const dynamicCategoryFor = (tech) => {
    if (pathFilter && pathFilter.length > 1 && Array.isArray(tech.sector) && tech.sector.length > pathFilter.length) {
      const segEn = tech.sector[pathFilter.length];
      if (segEn !== undefined) return { en: segEn, zh: pathSegmentZh(segEn) };
    }
    return { en: tech.sector_en || 'Unlabeled', zh: tech.sector_zh || '未標示' };
  };

  const datasetVersion = (dataset.dataset_version || "IEA ETP Clean Energy Tech").replace(/ 2026/g, '');

  // Header's "dataset generated at" box — real data (dataset.generated_at
  // from the cached JSON), never a fake live-refresh countdown, since this
  // dataset is a static snapshot, not a streaming feed.
  const generatedAtLabel = (() => {
    if (isSampleData) return L.sampleData;
    const raw = dataset?.generated_at;
    const d = raw ? new Date(raw) : null;
    if (!d || isNaN(d.getTime())) return '—';
    return d.toLocaleString(uiLang === 'en' ? 'en-US' : 'zh-TW', { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' });
  })();

  const searchData = useMemo(() => {
    if (!query.trim() && !selectedSector && !pathFilter) return { results: [], totalMatches: 0 };
    const lowerQuery = query.trim().toLowerCase();
    const lowerSector = selectedSector.trim().toLowerCase();

    const expandedQueries = lowerQuery ? (SEARCH_ALIASES[lowerQuery] || [lowerQuery]) : [];

    let results = rows.map(tech => {
      let score = 0, isDirectHit = false, isDescHit = false;
      const sectorTexts = [tech.sector_en, tech.sector_en_path, tech.sector_zh, tech.sector_zh_path, tech.breadcrumb, ...(Array.isArray(tech.sector) ? tech.sector : [])].filter(Boolean).map(s => String(s).toLowerCase());
      const isSectorMatch = selectedSector ? sectorTexts.some(t => t.includes(lowerSector)) : true;
      const subjectTexts = [tech.technology_name, tech.technology_name_zh].filter(Boolean).map(s => String(s).toLowerCase());
      const descTexts = [tech.description, tech.description_en, tech.description_zh, tech.technology_status_summary_zh, tech.market_dynamics_summary_zh, tech.NZErationale, tech.NZErationale_zh, tech.nze_rationale, tech.supplyChain, tech.supply_chain, tech.theme].filter(Boolean).map(s => String(s).toLowerCase());

      let passesSectorFilter = true, passesQueryFilter = true, passesPathFilter = true;

      // pathFilter (set by clicking a breadcrumb segment in the detail view)
      // is an ordered, exact prefix of tech.sector — deliberately NOT a loose
      // substring match like selectedSector, so clicking "Aluminium" only
      // pulls in technologies actually nested under that same path, not any
      // technology whose sector text happens to contain "Aluminium".
      if (pathFilter && pathFilter.length > 0) {
        const isPathMatch = Array.isArray(tech.sector) && pathFilter.every((seg, i) => tech.sector[i] === seg);
        if (!isPathMatch) passesPathFilter = false;
        else { score += 1; isDirectHit = true; }
      }

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

      if (passesSectorFilter && passesQueryFilter && passesPathFilter && score > 0) return { tech, score, isDirectHit, isDescHit };
      return null;
    }).filter(Boolean);

    results.sort((a, b) => {
      if (a.isDirectHit !== b.isDirectHit) return a.isDirectHit ? -1 : 1;
      if (b.score !== a.score) return b.score - a.score;
      return (b.tech.linked_records_count || 0) - (a.tech.linked_records_count || 0);
    });

    return { results, totalMatches: results.length };
  }, [query, rows, searchMode, selectedSector, pathFilter]);

  const totalMatches = searchData.totalMatches;
  const totalPages = Math.max(1, Math.ceil(totalMatches / 20));

  // The subset the Strategy tab actually runs on: the checked technologies
  // (if any of them are among the current matches), otherwise every match.
  // Intersecting with current results — rather than trusting the checked
  // set as-is — means stale checkmarks left over from a previous search
  // can't silently drag unrelated technologies into the analysis.
  const strategyResults = useMemo(() => {
    if (checkedTechIds.size === 0) return searchData.results;
    const checked = searchData.results.filter(r => checkedTechIds.has(r.tech._id));
    return checked.length > 0 ? checked : searchData.results;
  }, [searchData.results, checkedTechIds]);
  const checkedMatchCount = useMemo(
    () => (checkedTechIds.size === 0 ? 0 : searchData.results.filter(r => checkedTechIds.has(r.tech._id)).length),
    [searchData.results, checkedTechIds]
  );
  const isCheckedSubsetActive = checkedMatchCount > 0;
  const strategyCount = strategyResults.length;

  // Sector-distribution donut for the Strategy tab landing page — live off
  // the strategy subset (checked techs when any are ticked, else the whole
  // match set), so it always previews exactly what the overview would run on.
  const sectorDonutData = useMemo(() => {
    const counts = {};
    strategyResults.forEach(r => {
      const cat = dynamicCategoryFor(r.tech);
      const s = (uiLang === 'en' ? cat.en : cat.zh) || (uiLang === 'en' ? 'Unlabeled' : '未標示');
      counts[s] = (counts[s] || 0) + 1;
    });
    const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
    // Back to the original sky/violet/pink/amber/emerald hue set (not the
    // page-accent blue/emerald/purple/amber/teal reuse), but one shade
    // lighter across the board for a brighter, less saturated "tech" look
    // — the blue tone in particular read too saturated at full strength.
    const palette = ['#7DD3FC', '#C4B5FD', '#F9A8D4', '#FCD34D', '#6EE7B7'];
    const top = sorted.slice(0, 5).map(([name, count], i) => ({ label: name, value: count, color: palette[i % palette.length] }));
    const restCount = sorted.slice(5).reduce((sum, [, c]) => sum + c, 0);
    if (restCount > 0) top.push({ label: uiLang === 'en' ? 'Other' : '其他', value: restCount, color: '#94A3B8' });
    return top;
  }, [strategyResults, uiLang, pathFilter]);

  // Locked-segment summary sentence — rule-based, not AI-generated. Shows a
  // one-line definition clause when one exists for the current uiLang.
  // sectorName may now be a true leaf (SECTOR_DEFINITIONS/_EN, keyed by
  // zh/en respectively) or a dynamic-depth intermediate category from
  // dynamicCategoryFor above (SECTOR_PATH_DEFINITIONS_ZH/_EN, both keyed by
  // the English form) — sectorLeafZhMap membership tells them apart, since
  // it only ever contains genuine sector_en leaves.
  const [lockedSector, setLockedSector] = useState(null);

  const lockedSectorStats = useMemo(() => {
    if (!lockedSector) return null;
    const sectorName = lockedSector.label;
    const techsInSector = strategyResults
      .map(r => r.tech)
      .filter(t => {
        const cat = dynamicCategoryFor(t);
        return (uiLang === 'en' ? cat.en : cat.zh) === sectorName;
      });
    if (techsInSector.length === 0) return null;

    const trlCounts = [0, 0, 0, 0];
    techsInSector.forEach(t => {
      const trl = parseInt(t.latest_trl) || 0;
      let bucket = -1;
      if (trl > 0 && trl <= 3) bucket = 0;
      else if (trl <= 6) bucket = 1;
      else if (trl <= 8) bucket = 2;
      else if (trl > 8) bucket = 3;
      if (bucket >= 0) trlCounts[bucket]++;
    });
    let topBucket = 0;
    trlCounts.forEach((c, i) => { if (c > trlCounts[topBucket]) topBucket = i; });

    const initiativesInSector = techsInSector.flatMap(t => t.initiatives || []);
    const caseCount = initiativesInSector.length;
    const countryCounts = {};
    initiativesInSector.forEach(i => {
      if (!i.country) return;
      const country = translateCountry(i.country, uiLang);
      countryCounts[country] = (countryCounts[country] || 0) + 1;
    });
    const topCountryEntry = Object.entries(countryCounts).sort((a, b) => b[1] - a[1])[0];

    const sectorNameEn = dynamicCategoryFor(techsInSector[0]).en;
    const isLeaf = !!sectorLeafZhMap[sectorNameEn];
    const definition = (uiLang === 'en'
      ? (isLeaf ? SECTOR_DEFINITIONS_EN[sectorNameEn] : SECTOR_PATH_DEFINITIONS_EN[sectorNameEn])
      : (isLeaf ? SECTOR_DEFINITIONS[sectorName] : SECTOR_PATH_DEFINITIONS_ZH[sectorNameEn])) || null;

    return {
      count: techsInSector.length,
      topStage: TRL_COMPACT_LABELS_ZH[topBucket],
      caseCount,
      topCountry: topCountryEntry ? topCountryEntry[0] : null,
      definition
    };
  }, [lockedSector, strategyResults, uiLang, pathFilter]);

  const lockedSectorSummary = useMemo(() => {
    if (!lockedSector || !lockedSectorStats) return null;
    const { count, topStage, caseCount, topCountry, definition } = lockedSectorStats;
    if (uiLang === 'en') {
      const defClauseEn = definition ? ` (${definition})` : '';
      const countryClause = topCountry ? `, most active in ${topCountry}` : '';
      return `"${lockedSector.label}"${defClauseEn} currently has ${count} technologies, mostly at ${topStage}; ${caseCount} related cases in total${countryClause}.`;
    }
    const defClause = definition ? `（${definition}）` : '';
    const countryClause = topCountry ? `，以${topCountry}最為活躍` : '';
    return `「${lockedSector.label}」${defClause}目前有 ${count} 項技術，成熟度以 ${topStage} 階段為主；共連結 ${caseCount} 筆案例${countryClause}。`;
  }, [lockedSector, lockedSectorStats, uiLang]);

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
    const exportTechs = strategyResults.map(r => r.tech);
    handleExportTechsCsv(exportTechs, isCheckedSubsetActive ? `IEA_勾選技術_${exportTechs.length}筆` : `IEA_完整查詢結果_${exportTechs.length}筆`);
  };

  const handleExportSelectedTechCsv = () => {
    if (!selectedTech) return;
    const rawName = (uiLang === 'en' ? selectedTech.technology_name : selectedTech.technology_name_zh) || selectedTech.technology_name || 'tech';
    const safeName = String(rawName).replace(/[\\/:*?"<>|]/g, '').slice(0, 40);
    handleExportTechsCsv([selectedTech], `IEA_${safeName}`);
  };

  // Copy-to-clipboard companion to buildArticleMarkdown — lets the user jump
  // straight to notebooklm.google.com and paste, since NotebookLM has no
  // public API for us to push content into a notebook automatically.
  const handleCopyTechsArticle = async (techs, copyKey) => {
    if (!techs || techs.length === 0) return;
    const md = buildArticleMarkdown(techs, uiLang);
    try {
      await navigator.clipboard.writeText(md);
      setCopiedArticleKey(copyKey);
      window.clearTimeout(copyResetTimerRef.current);
      copyResetTimerRef.current = window.setTimeout(() => setCopiedArticleKey(null), 2000);
    } catch (err) {
      console.error('Clipboard copy failed', err);
    }
  };

  const handleCopyListArticle = () => {
    const exportTechs = strategyResults.map(r => r.tech);
    handleCopyTechsArticle(exportTechs, 'list');
  };

  // NOTEBOOKLM_PPT_PROMPT is fixed text (not per-technology), so this just
  // copies it as-is — same clipboard/feedback mechanism as the article copy
  // above, keyed separately so the two buttons' "copied!" states don't
  // collide.
  const handleCopyPptPrompt = async () => {
    try {
      await navigator.clipboard.writeText(NOTEBOOKLM_PPT_PROMPT);
      setCopiedArticleKey('pptPrompt');
      window.clearTimeout(copyResetTimerRef.current);
      copyResetTimerRef.current = window.setTimeout(() => setCopiedArticleKey(null), 2000);
    } catch (err) {
      console.error('Clipboard copy failed', err);
    }
  };

  return (
    <div className="h-screen bg-slate-50 flex flex-col font-sans overflow-hidden">
      {/* Header — vibrant blue gradient with a decorative, on-theme
          animated background (slow-spinning wind turbines + a scrolling
          wave strip) instead of generic tech particle effects, since this
          is a clean-energy platform. The dataset-time box on the right
          always shows dataset.generated_at (a real, static timestamp from
          the cached JSON), never a live-refresh countdown — this dataset
          doesn't auto-update, so implying it does would be misleading. */}
      <header className="relative overflow-hidden bg-gradient-to-br from-cyan-500 via-blue-600 to-blue-800 px-6 py-3 2xl:py-5 shadow-lg z-10 flex-shrink-0">
        {/* Decorative background — pointer-events-none, low opacity, sits
            behind the real content (z-0) so it never interferes with
            reading or clicking anything above it (z-10). */}
        <div className="absolute inset-0 z-0 pointer-events-none overflow-hidden opacity-20">
          {/* Positioned in the header's open middle zone — the left edge
              is title text, the right edge is the dataset-time box and
              reload button, both of which these must stay clear of. */}
          <WindTurbineIcon className="absolute -top-4 left-[62%] w-16 h-24" duration="10s" />
          <WindTurbineIcon className="absolute -top-2 left-[70%] w-11 h-16" duration="8s" />
          <WindTurbineIcon className="absolute top-1 left-[77%] w-9 h-14" duration="11s" />
          <div className="absolute bottom-0 left-0 w-full h-6 overflow-hidden">
            <div className="flex animate-wave-scroll" style={{ width: '200%' }}>
              <svg viewBox="0 0 1200 40" className="w-1/2 flex-shrink-0" preserveAspectRatio="none">
                <path d="M0,20 Q150,0 300,20 T600,20 T900,20 T1200,20 L1200,40 L0,40 Z" fill="white" />
              </svg>
              <svg viewBox="0 0 1200 40" className="w-1/2 flex-shrink-0" preserveAspectRatio="none">
                <path d="M0,20 Q150,0 300,20 T600,20 T900,20 T1200,20 L1200,40 L0,40 Z" fill="white" />
              </svg>
            </div>
          </div>
        </div>

        <div className="relative z-10 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
          {/* Left: icon badge + title + status */}
          <div className="flex items-center gap-3 min-w-0">
            <div className="bg-white/15 border border-white/20 text-white rounded-xl p-2.5 flex-shrink-0">
              <Sparkles size={22} />
            </div>
            <div className="min-w-0">
              <h1 className="text-xl md:text-2xl font-bold text-white truncate">{L.appTitle}</h1>
              <div className="text-sm mt-1 flex items-center gap-2 flex-wrap">
                <span className="text-blue-100 font-medium">{L.appSubtitle}</span>
                <span className="text-blue-300">·</span>
                <span className="text-blue-100 font-medium">{L.datasetVersion(isSampleData ? L.sampleData : datasetVersion)}</span>
                {appState.status === 'fallback' && <span className="bg-white/15 text-white border border-white/20 px-2 py-0.5 rounded text-sm">{L.fallbackBadge}</span>}
                {appState.status === 'loading' && <span className="bg-white/15 text-white border border-white/20 px-2 py-0.5 rounded text-sm animate-pulse">{L.loadingBadge}</span>}
                {appState.status === 'success' && (
                  <span className="bg-white/15 text-white border border-white/20 px-2 py-0.5 rounded text-sm flex items-center gap-1">
                    <CheckCircle2 size={12} /> {L.successBadge(techCount, initiativeCount)}
                  </span>
                )}
                {appState.status === 'success' && appState.isRawFormat && (
                  <span className="bg-white/15 text-white border border-white/20 px-2 py-0.5 rounded text-sm">{L.rawFormatDetected}</span>
                )}
                {appState.status === 'error' && (
                  <span className="bg-red-500/20 text-red-200 border border-red-300/40 px-2 py-0.5 rounded text-sm flex items-center gap-1">
                    <AlertCircle size={12} /> {appState.errorMsg}
                  </span>
                )}
              </div>
            </div>
          </div>

          {/* Right: dataset-generated-at box + restore-default ("reload") button */}
          <div className="flex items-center gap-3 flex-shrink-0">
            <div className="bg-white/15 border border-white/20 rounded-xl px-4 py-2 text-right">
              <div className="text-[11px] uppercase tracking-wider text-blue-100 font-semibold">{L.datasetTimeLabel}</div>
              <div className="text-sm font-bold text-white mt-0.5 whitespace-nowrap">{generatedAtLabel}</div>
            </div>
            {isCustomUpload && (
              <button
                type="button"
                onClick={onRestoreDefault}
                title={L.restoreDefaultData}
                className="flex items-center gap-1.5 bg-white/10 hover:bg-white/20 border border-white/10 text-white text-sm font-medium px-4 py-2.5 rounded-xl transition-colors flex-shrink-0"
              >
                <RotateCcw size={14} /> {L.restoreDefaultData}
              </button>
            )}
          </div>
        </div>

        {/* Secondary row: platform controls (language, AI provider, API key, upload) */}
        <div className="mt-3 pt-3 2xl:mt-4 2xl:pt-4 border-t border-white/10 flex flex-wrap items-center gap-3 text-base">

          <button
            type="button"
            onClick={() => setUiLang(l => (l === 'zh' ? 'en' : 'zh'))}
            title="Switch platform language / 切換平台語言"
            className="text-sm font-medium px-3 py-1.5 rounded-full border border-white/20 text-slate-200 hover:bg-white/10 transition-colors flex-shrink-0"
          >
            {uiLang === 'zh' ? 'EN' : '中文'}
          </button>

          <select
            value={aiProvider}
            onChange={e => onProviderChange(e.target.value)}
            title={L.aiProviderLabel}
            className="text-sm font-medium px-2 py-1.5 rounded-lg border border-white/20 text-slate-100 bg-white/10 focus:outline-none focus:ring-2 focus:ring-blue-300"
          >
            {AI_PROVIDER_ORDER.map(p => <option key={p} value={p} className="text-slate-800">{AI_PROVIDERS[p].label}</option>)}
          </select>

          {aiProvider === 'openrouter' && (
            <input
              type="text"
              value={openRouterModel}
              onChange={e => onOpenRouterModelChange(e.target.value)}
              placeholder={L.openRouterModelPlaceholder}
              className="text-sm text-slate-100 placeholder:text-slate-400 bg-white/10 border border-white/20 rounded-lg px-3 py-1.5 focus:outline-none focus:ring-2 focus:ring-blue-300 w-44"
              autoComplete="off"
              spellCheck={false}
            />
          )}

          <div className="flex items-center gap-1.5 bg-white/10 border border-white/20 rounded-lg pl-3 pr-1.5 py-1.5">
            <KeyRound size={14} className={apiKey.trim() ? 'text-emerald-400' : 'text-slate-400'} />
            <input
              type={showApiKey ? 'text' : 'password'}
              value={apiKey}
              onChange={e => onApiKeyChange(e.target.value)}
              placeholder={L.apiKeyPlaceholder(AI_PROVIDERS[aiProvider].label)}
              className="text-sm text-slate-100 placeholder:text-slate-400 bg-transparent focus:outline-none w-40"
              autoComplete="off"
              spellCheck={false}
            />
            <button
              type="button"
              onClick={() => setShowApiKey(v => !v)}
              className="text-slate-400 hover:text-slate-200 p-1 rounded transition-colors"
              title={showApiKey ? L.hideKey : L.showKey}
            >
              {showApiKey ? <EyeOff size={14} /> : <Eye size={14} />}
            </button>
          </div>

          <label className="cursor-pointer bg-white text-blue-800 hover:bg-blue-50 font-medium py-1.5 px-4 rounded-lg transition-colors flex items-center gap-2">
            <FileJson size={16} />
            <span>{L.uploadJson}</span>
            <input type="file" accept=".json" className="hidden" onChange={onUpload} />
          </label>
        </div>
      </header>

      {/* Main layout */}
      {isInitialLoading ? <SkeletonMain /> : (
      <main className="flex-1 flex flex-col md:flex-row overflow-hidden min-h-0">

        {/* Left panel */}
        <section className="w-full md:w-[400px] lg:w-[460px] flex-shrink-0 border-b md:border-b-0 md:border-r border-slate-200 bg-white flex flex-col h-1/2 md:h-full min-h-0">
          <div className="p-3 2xl:p-4 border-b border-slate-100 bg-slate-50 flex-shrink-0">
            {/* Collapsible search controls — folding these away hands their
                vertical space to the technology list below, which is where
                the checkbox-selection work actually happens. */}
            {!searchPanelCollapsed && (
              <>
                <div className="flex gap-2 mb-2 2xl:mb-3">
                  <button onClick={() => setSearchMode('subject')} className={`flex-1 text-sm py-1.5 rounded-md font-medium transition-colors border ${searchMode === 'subject' ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-50'}`}>{L.subjectSearch}</button>
                  <button onClick={() => setSearchMode('fulltext')} className={`flex-1 text-sm py-1.5 rounded-md font-medium transition-colors border ${searchMode === 'fulltext' ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-50'}`}>{L.fulltextSearch}</button>
                </div>

                <div className="relative mb-2 2xl:mb-3">
                  <select value={selectedSector} onChange={e => { setSelectedSector(e.target.value); setPathFilter(null); }} className="w-full pl-3 pr-8 py-2 rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500 text-base appearance-none bg-white font-medium text-slate-700">
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
                  <input type="text" placeholder={L.searchPlaceholder} className="w-full pl-10 pr-4 py-2.5 rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-base transition-shadow" value={query} onChange={e => setQuery(e.target.value)} />
                </div>
                <div className="text-xs text-slate-500 mt-2 leading-relaxed">{L.searchHint}</div>
                {pathFilter && pathFilter.length > 0 && (
                  <div className="flex items-center justify-between gap-2 bg-blue-50 border border-blue-100 rounded px-2 py-1.5 mt-2">
                    <span className="text-xs text-blue-700 font-medium truncate">
                      {L.pathFilterActive}: {pathFilter.map(seg => uiLang === 'en' ? seg : pathSegmentZh(seg)).join(' > ')}
                    </span>
                    <button onClick={() => setPathFilter(null)} className="text-xs font-medium text-blue-600 hover:text-blue-800 underline flex-shrink-0">
                      {L.clearPathFilter}
                    </button>
                  </div>
                )}
              </>
            )}

            <div className={`${searchPanelCollapsed ? '' : 'mt-2 2xl:mt-3'} flex flex-col gap-2`}>
              <div className="flex items-center justify-between gap-2">
                <div className="text-sm text-slate-600 font-medium">{totalMatches > 0 ? L.resultsCount(totalMatches, currentPage, totalPages) : L.noResultsShort}</div>
                <button
                  onClick={() => setSearchPanelCollapsed(v => !v)}
                  title={searchPanelCollapsed ? L.expandSearch : L.collapseSearch}
                  className="flex items-center gap-1 text-xs font-medium text-slate-500 hover:text-slate-700 bg-white hover:bg-slate-100 border border-slate-200 px-2 py-1 rounded transition-colors flex-shrink-0"
                >
                  {searchPanelCollapsed ? <ChevronDown size={12} /> : <ChevronUp size={12} />}
                  {searchPanelCollapsed ? L.expandSearch : L.collapseSearch}
                </button>
              </div>
              {checkedMatchCount > 0 && (
                <div className="flex items-center justify-between gap-2 bg-purple-50 border border-purple-100 rounded px-2 py-1.5">
                  <span className="text-xs text-purple-700 font-medium">{L.checkedCountBar(checkedMatchCount)}</span>
                  <button onClick={() => setCheckedTechIds(new Set())} className="text-xs font-medium text-purple-600 hover:text-purple-800 underline flex-shrink-0">
                    {L.clearChecked}
                  </button>
                </div>
              )}
              {selectedTech && (
                <div className="flex flex-wrap items-center gap-1.5">
                  <button onClick={handleExportSelectedTechCsv} className="flex items-center gap-1 text-xs font-medium text-emerald-600 hover:text-emerald-700 bg-emerald-50 hover:bg-emerald-100 px-2 py-1 rounded transition-colors">
                    <Download size={12} /> {L.exportSingle}
                  </button>
                </div>
              )}
            </div>
          </div>

          <div ref={listContainerRef} className="flex-1 overflow-y-auto p-2 bg-white">
            {!query.trim() && !selectedSector && !pathFilter ? (
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
                      {/* Checkbox sits beside (not inside) the row button —
                          an input nested in a button is invalid HTML and
                          the two click targets would fight each other. */}
                      <div className={`w-full flex items-start gap-2 p-2.5 rounded-lg transition-colors ${isSelected ? 'bg-blue-50 border border-blue-200' : 'hover:bg-slate-50 border border-transparent'}`}>
                        <input
                          type="checkbox"
                          checked={checkedTechIds.has(tech._id)}
                          onChange={() => toggleTechChecked(tech._id)}
                          className="mt-1 w-4 h-4 accent-purple-600 flex-shrink-0 cursor-pointer"
                        />
                        <button onClick={() => { setSelectedTech(tech); setRightPanelTab('detail'); }} className="flex-1 min-w-0 text-left flex flex-col gap-1">
                          <div className="flex justify-between items-start gap-2">
                            <span className="font-semibold text-slate-800 text-base line-clamp-1">{primaryName}</span>
                            <span className={`flex-shrink-0 text-xs px-1.5 py-0.5 rounded font-medium ${item.isDirectHit ? 'bg-blue-100 text-blue-700' : 'bg-orange-100 text-orange-700'}`}>{item.isDirectHit ? L.directHit : L.descHit}</span>
                          </div>
                          <span className="text-sm text-slate-500 line-clamp-1">{secondaryName || L.unlabeled}</span>
                          <div className="flex justify-between items-center mt-0.5">
                            <span className="text-xs bg-slate-100 text-slate-600 px-2 py-0.5 rounded border border-slate-200">{tech.sector_zh || L.unlabeled}</span>
                            <span className="text-xs text-blue-600 font-medium">{L.caseCountInline(tech.linked_records_count || 0)}</span>
                          </div>
                        </button>
                      </div>
                    </li>
                  );
                })}
              </ul>
            )}
          </div>

          {totalMatches > 0 && (
            <div className="p-3 border-t border-slate-200 bg-white flex justify-between items-center flex-shrink-0">
              <button onClick={() => setCurrentPage(p => Math.max(1, p - 1))} disabled={currentPage === 1} className="text-sm font-medium px-3 py-1.5 rounded border border-slate-300 text-slate-600 hover:bg-slate-50 disabled:opacity-30 transition-colors flex items-center gap-1"><ChevronLeft size={14} />{L.prevPage}</button>
              <button onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))} disabled={currentPage === totalPages} className="text-sm font-medium px-3 py-1.5 rounded border border-slate-300 text-slate-600 hover:bg-slate-50 disabled:opacity-30 transition-colors flex items-center gap-1">{L.nextPage}<ChevronRight size={14} /></button>
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
              className={`px-4 py-3 text-base font-medium border-b-2 -mb-px transition-colors flex items-center gap-1.5 ${rightPanelTab === 'detail' ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-500 hover:text-slate-700'}`}
            >
              <FileText size={14} /> {L.detailTab}
            </button>
            <button
              onClick={() => setRightPanelTab('strategy')}
              className={`px-4 py-3 text-base font-medium border-b-2 -mb-px transition-colors flex items-center gap-1.5 ${rightPanelTab === 'strategy' ? 'border-purple-600 text-purple-600' : 'border-transparent text-slate-500 hover:text-slate-700'}`}
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
                      <h2 className="text-2xl md:text-3xl font-bold text-slate-900 mb-1">{uiLang === 'en' ? (selectedTech.technology_name || techLabel(selectedTech)) : (selectedTech.technology_name_zh || selectedTech.technology_name || L.unlabeled)}</h2>
                      <p className="text-sm md:text-base text-slate-500 font-mono mb-4">{(uiLang === 'en' ? selectedTech.technology_name_zh : selectedTech.technology_name) || L.unlabeled}</p>
                      <div className="flex flex-wrap gap-2 text-sm font-medium">
                        <span className="bg-blue-50 text-blue-700 border border-blue-100 px-2 py-1 rounded">{L.fieldSector((uiLang === 'en' ? selectedTech.sector_en : null) || selectedTech.sector_zh || L.unlabeled)}</span>
                        <span className="bg-slate-100 text-slate-700 border border-slate-200 px-2 py-1 rounded">{L.fieldTrl(selectedTech.latest_trl || L.unlabeled)}</span>
                        {selectedTech.adoption_stage && (
                          <span className="bg-purple-50 text-purple-700 border border-purple-100 px-2 py-1 rounded">{L.fieldAdoptionStage(adoptionStageLabel(selectedTech.adoption_stage, uiLang))}</span>
                        )}
                        <span className="bg-emerald-50 text-emerald-700 border border-emerald-200 px-2 py-1 rounded">{L.fieldCaseCount(selectedTech.linked_records_count || 0)}</span>
                      </div>
                      {Array.isArray(selectedTech.sector) && selectedTech.sector.length > 0 && (
                        <div className="mt-3 pt-3 border-t border-slate-100">
                          <div className="text-xs text-slate-400 mb-1.5">{L.pathBreadcrumbLabel}</div>
                          <div className="flex flex-wrap items-center gap-1 text-sm">
                            {selectedTech.sector.map((seg, i) => (
                              <React.Fragment key={i}>
                                {i > 0 && <ChevronRight size={14} className="text-slate-300 flex-shrink-0" />}
                                <button
                                  onClick={() => {
                                    setPathFilter(selectedTech.sector.slice(0, i + 1));
                                    setSelectedSector('');
                                    setQuery('');
                                  }}
                                  className="text-blue-600 hover:text-blue-800 hover:underline font-medium px-1 py-0.5 rounded transition-colors"
                                >
                                  {uiLang === 'en' ? seg : sectorSegmentLabel(selectedTech, i)}
                                </button>
                              </React.Fragment>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>

                    {/* Summary */}
                    <div className="bg-white p-5 md:p-6 rounded-xl shadow-sm border border-slate-200">
                      <div className="border-b border-slate-200 pb-4 mb-4">
                        <h3 className="text-xl font-semibold text-slate-800 mb-3">{L.summaryTitle}</h3>
                        <div className="bg-amber-50 text-amber-700 px-3 py-2 rounded text-sm font-medium border border-amber-200 inline-flex items-start gap-1.5">
                          <Info size={14} className="mt-0.5 flex-shrink-0" />
                          <span>{L.aiDisclaimer}</span>
                        </div>
                      </div>
                      <div className="space-y-5">
                        <div className="space-y-1.5">
                          <h4 className="text-base font-bold text-slate-600">{L.summaryTitle}</h4>
                          {(uiLang === 'en'
                            ? (selectedTech.technology_status_summary_en || selectedTech.description_en || selectedTech.description || selectedTech.technology_status_summary_zh)
                            : (selectedTech.technology_status_summary_zh || selectedTech.description_zh || selectedTech.description))
                            ? <p className="text-base text-slate-700 leading-relaxed">{uiLang === 'en'
                                ? (selectedTech.technology_status_summary_en || selectedTech.description_en || selectedTech.description || selectedTech.technology_status_summary_zh)
                                : (selectedTech.technology_status_summary_zh || selectedTech.description_zh || selectedTech.description)}</p>
                            : <p className="text-base text-slate-400 italic">{L.noSummary}</p>}
                        </div>
                        <div className="space-y-1.5 pt-4 border-t border-slate-100">
                          <h4 className="text-base font-bold text-slate-600">{L.marketTitle}</h4>
                          {(uiLang === 'en'
                            ? (selectedTech.market_dynamics_summary_en || selectedTech.NZErationale_en || selectedTech.NZErationale || selectedTech.market_dynamics_summary_zh)
                            : (selectedTech.market_dynamics_summary_zh || selectedTech.NZErationale_zh || selectedTech.NZErationale))
                            ? <p className="text-base text-slate-700 leading-relaxed">{uiLang === 'en'
                                ? (selectedTech.market_dynamics_summary_en || selectedTech.NZErationale_en || selectedTech.NZErationale || selectedTech.market_dynamics_summary_zh)
                                : (selectedTech.market_dynamics_summary_zh || selectedTech.NZErationale_zh || selectedTech.NZErationale)}</p>
                            : <p className="text-base text-slate-400 italic">{L.noMarket}</p>}
                        </div>
                      </div>
                    </div>

                    {/* Initiatives */}
                    <div className="bg-white p-5 md:p-6 rounded-xl shadow-sm border border-slate-200">
                      <div className="flex flex-wrap justify-between items-center mb-4 gap-3">
                        <h3 className="text-xl font-semibold text-slate-800 flex items-center gap-2">
                          <Database size={18} className="text-emerald-500" /> {L.initiativesTitle}
                        </h3>
                        <span className="text-base bg-slate-100 text-slate-600 px-2 py-1 rounded-full border border-slate-200">{L.totalCount(sortedInitiatives.length)}</span>
                      </div>
                      {sortedInitiatives.length === 0 ? (
                        <div className="text-center text-slate-500 py-8 bg-slate-50 rounded-lg border border-slate-100 border-dashed">{L.noInitiatives}</div>
                      ) : (
                        <div className="space-y-4">
                          {displayedInitiatives.map((init, idx) => (
                            <div key={idx} className="border-l-2 border-emerald-400 pl-4 py-1">
                              <div className="flex flex-wrap gap-2 mb-1 text-sm font-medium text-slate-500">
                                {init.year && <span className="bg-slate-100 px-1.5 py-0.5 rounded border border-slate-200">{init.year}</span>}
                                {init.country && <span className="bg-emerald-50 text-emerald-700 px-1.5 py-0.5 rounded border border-emerald-100">{translateCountry(init.country, uiLang)}</span>}
                                {init.type && <span className="text-blue-600">{init.type}</span>}
                              </div>
                              <p className="text-base text-slate-700 leading-relaxed mt-1">
                                {init.description || L.unlabeled}
                                {safeLinkUrl(init.read_more) && <a href={safeLinkUrl(init.read_more)} target="_blank" rel="noopener noreferrer" className="text-blue-600 hover:text-blue-800 underline ml-2">{L.sourceLink}</a>}
                              </p>
                            </div>
                          ))}
                          {sortedInitiatives.length > 3 && (
                            <div className="text-center pt-3 border-t border-slate-100">
                              <button onClick={() => setShowAllInitiatives(!showAllInitiatives)} className="text-sm text-blue-600 hover:text-blue-800 font-medium px-4 py-1.5 bg-blue-50 hover:bg-blue-100 rounded transition-colors">
                                {showAllInitiatives ? L.collapseCases : L.expandCases(sortedInitiatives.length)}
                              </button>
                            </div>
                          )}
                        </div>
                      )}
                    </div>

                    {/* Disclaimer — moved here from a dark footer bar so it
                        blends into the page bottom as muted gray text. */}
                    <p className="text-xs text-slate-400 text-center leading-relaxed pt-2">{L.footerDisclaimer}</p>
                  </div>
                )}
              </div>
            )}

            {rightPanelTab === 'strategy' && (
              <div className="flex-1 flex flex-col overflow-y-auto p-4 md:p-8 w-full">
                <div className="max-w-6xl mx-auto w-full">
                  <div className="mb-6">
                    <h2 className="text-xl font-bold text-slate-800">{L.strategyTitle}</h2>
                    <p className="text-base text-slate-500 mt-1">{L.strategySubtitle}</p>
                  </div>

                  <div className="flex flex-col lg:flex-row gap-6 items-stretch">
                    {/* Sector-distribution donut — a purely informational
                        preview (no click/navigation of its own) of the
                        current search results, so it doesn't compete with
                        the two action cards as a third entry point. Same
                        plain-white-card look as the two beside it, just
                        with a blue accent, for a cleaner and more cohesive
                        page instead of a separate dark "control room" card. */}
                    {totalMatches > 0 && (
                      <div className="lg:w-[340px] flex-shrink-0 bg-white rounded-2xl shadow-sm border border-slate-200 border-t-4 border-t-blue-500 p-6 md:p-8 flex flex-col items-center">
                        <h3 className="text-lg font-bold text-slate-800 mb-1 self-start">{L.sectorDistributionTitle}</h3>
                        <p className="text-xs text-slate-400 mb-5 self-start">{L.sectorDistributionHint}</p>
                        {isCheckedSubsetActive && (
                          <p className="text-sm text-purple-600 bg-purple-50 border border-purple-100 rounded px-2 py-1 mb-4 self-start">{L.checkedSubsetBadge(strategyCount)}</p>
                        )}
                        <DonutChart segments={sectorDonutData} centerValue={strategyCount} centerLabel={L.sectorDistributionCenterLabel} size={208} strokeWidth={22} onLockChange={setLockedSector} />
                        {lockedSectorSummary && (
                          <p className="text-base text-slate-600 leading-relaxed mt-5 pt-4 border-t border-slate-100 self-stretch">
                            {lockedSectorSummary}
                          </p>
                        )}
                      </div>
                    )}

                    <div className="flex-1 space-y-6 min-w-0">
                      {/* Strategy overview card — opens instantly with a rule-based
                          summary; the option to upgrade to an AI-generated version
                          lives inside that same modal, so this is one progressive
                          entry point rather than two separate, overlapping ones. */}
                      <div className="bg-white p-6 md:p-8 rounded-2xl shadow-sm border border-slate-200 border-t-4 border-t-purple-500">
                        <div className="flex items-start gap-4 mb-5">
                          <div className="bg-gradient-to-br from-purple-500 to-blue-500 text-white rounded-xl p-3 flex-shrink-0">
                            <Sparkles size={22} />
                          </div>
                          <div>
                            <h3 className="text-lg font-bold text-slate-800">{L.strategyCardTitle}</h3>
                            <p className="text-base text-slate-500 mt-1 leading-relaxed">
                              {isCheckedSubsetActive ? L.strategyCardDescCheckedPrefix : L.strategyCardDescPrefix} <span className="font-semibold text-slate-700">{strategyCount}</span> {L.strategyCardDescSuffix}
                            </p>
                          </div>
                        </div>
                        <button
                          onClick={() => onOpenStrategyOverview({ results: strategyResults, totalMatches: strategyCount })}
                          disabled={strategyCount === 0}
                          className="w-full sm:w-auto flex items-center justify-center gap-2 bg-gradient-to-r from-purple-600 to-blue-600 hover:from-purple-700 hover:to-blue-700 disabled:from-slate-300 disabled:to-slate-300 disabled:cursor-not-allowed text-white shadow-sm rounded-lg px-6 py-3 text-base font-semibold transition-all"
                        >
                          <Sparkles size={18} /> {L.openStrategyBtn}
                        </button>
                      </div>

                      {/* Export card — CSV download only; the NotebookLM/PPT
                          workflow used to be crammed into this same card
                          (5 buttons, 2 divider rows) but the two features
                          serve different intents (raw data vs. an external
                          slide-generation workflow), so they're now split
                          into separate cards. */}
                      <div className="bg-white p-6 md:p-8 rounded-2xl shadow-sm border border-slate-200 border-t-4 border-t-emerald-500">
                        <div className="flex items-start gap-4 mb-5">
                          <div className="bg-emerald-50 text-emerald-600 rounded-xl p-3 flex-shrink-0">
                            <Download size={22} />
                          </div>
                          <div>
                            <h3 className="text-lg font-bold text-slate-800">{L.exportStrategyTitle}</h3>
                            <p className="text-base text-slate-500 mt-1 leading-relaxed">
                              {isCheckedSubsetActive ? L.strategyCardDescCheckedPrefix : L.exportStrategyDescPrefix} <span className="font-semibold text-slate-700">{strategyCount}</span> {L.exportStrategyDescSuffix}
                            </p>
                          </div>
                        </div>
                        <button
                          onClick={handleExportListCsv}
                          disabled={strategyCount === 0}
                          className="w-full sm:w-auto flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-300 disabled:cursor-not-allowed text-white shadow-sm rounded-lg px-6 py-3 text-base font-semibold transition-colors"
                        >
                          <Download size={18} /> {L.exportAll}
                        </button>
                      </div>

                      {/* NotebookLM / slide-generation card — the step-by-step
                          "how do I actually use this" detail lives in
                          PptGuideModal now (with real screenshots and
                          arrows), so this card doesn't need to spell out the
                          paste-here/paste-there sequence in its own text
                          anymore; the guide button is the primary action,
                          and the three raw actions it walks through sit
                          below as lightweight secondary buttons for anyone
                          who already knows the flow. */}
                      <div className="bg-white p-6 md:p-8 rounded-2xl shadow-sm border border-slate-200 border-t-4 border-t-blue-500">
                        <div className="flex items-start gap-4 mb-5">
                          <div className="bg-gradient-to-br from-blue-500 to-cyan-500 text-white rounded-xl p-3 flex-shrink-0">
                            <Presentation size={22} />
                          </div>
                          <div>
                            <h3 className="text-lg font-bold text-slate-800">{L.pptCardTitle}</h3>
                            <p className="text-base text-slate-500 mt-1 leading-relaxed">
                              {isCheckedSubsetActive ? L.pptCardDescCheckedPrefix : L.pptCardDescPrefix} <span className="font-semibold text-slate-700">{strategyCount}</span> {L.pptCardDescSuffix}
                            </p>
                          </div>
                        </div>

                        <button
                          onClick={() => setShowPptGuide(true)}
                          className="w-full sm:w-auto flex items-center justify-center gap-2 bg-gradient-to-r from-blue-600 to-cyan-600 hover:from-blue-700 hover:to-cyan-700 text-white shadow-sm rounded-lg px-6 py-3 text-base font-semibold transition-all"
                        >
                          <HelpCircle size={18} /> {L.openPptGuide}
                        </button>

                        <div className="mt-4 pt-4 border-t border-slate-100 flex flex-wrap gap-2">
                          <button
                            onClick={handleCopyListArticle}
                            disabled={strategyCount === 0}
                            className="flex items-center gap-1.5 text-sm font-medium text-slate-600 hover:text-slate-800 disabled:text-slate-300 disabled:cursor-not-allowed bg-slate-50 hover:bg-slate-100 disabled:hover:bg-slate-50 rounded-lg px-4 py-2 transition-colors"
                          >
                            <Copy size={15} /> {copiedArticleKey === 'list' ? L.copied : L.copyArticle}
                          </button>
                          <a
                            href="https://notebooklm.google.com"
                            target="_blank"
                            rel="noopener noreferrer"
                            className="flex items-center gap-1.5 text-sm font-medium text-slate-600 hover:text-slate-800 bg-slate-50 hover:bg-slate-100 rounded-lg px-4 py-2 transition-colors"
                          >
                            <ExternalLink size={15} /> {L.openNotebookLM}
                          </a>
                          <button
                            onClick={handleCopyPptPrompt}
                            disabled={strategyCount === 0}
                            className="flex items-center gap-1.5 text-sm font-medium text-slate-600 hover:text-slate-800 disabled:text-slate-300 disabled:cursor-not-allowed bg-slate-50 hover:bg-slate-100 disabled:hover:bg-slate-50 rounded-lg px-4 py-2 transition-colors"
                          >
                            <Copy size={15} /> {copiedArticleKey === 'pptPrompt' ? L.copied : L.copyPptPrompt}
                          </button>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Disclaimer — moved here from a dark footer bar so it
                      blends into the page bottom as muted gray text. */}
                  <p className="text-xs text-slate-400 text-center leading-relaxed mt-8">{L.footerDisclaimer}</p>
                </div>
              </div>
            )}
          </div>
        </section>
      </main>
      )}
      <PptGuideModal isOpen={showPptGuide} onClose={() => setShowPptGuide(false)} uiLang={uiLang} />
    </div>
  );
}
