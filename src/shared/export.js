export function createLibraryJson(library) {
  return JSON.stringify({
    format: "english-reader-library",
    schemaVersion: 1,
    exportedAt: new Date().toISOString(),
    vocabulary: library.vocabulary ?? [],
    sentences: library.sentences ?? []
  }, null, 2);
}

export function createExcelXml(vocabulary, sentences) {
  const sheets = [
    ["生词库", ["词语", "音标", "释义", "词义", "语言", "来源"], vocabulary.map((item) => [
      item.text ?? "",
      item.ipa ?? "",
      item.chineseDefinition ?? "",
      (item.meanings ?? []).map((meaning) => `${partName(meaning.partOfSpeech)} ${meaning.definitionZh ?? meaning.definitionEn ?? ""}`).join("\n"),
      item.sourceLanguage ?? "",
      sourceOrigin(item.source?.pageUrl)
    ])],
    ["句子库", ["句子", "翻译", "固定搭配", "语言", "来源"], sentences.map((item) => [
      item.text ?? "",
      item.translationZh ?? "",
      (item.collocations ?? []).map(({ phrase, meaningZh }) => `${phrase} — ${meaningZh}`).join("\n"),
      item.sourceLanguage ?? "",
      safeHttpUrl(item.source?.pageUrl)
    ])]
  ];
  const worksheets = sheets.map(([name, headers, rows]) => {
    const table = [headers, ...rows].map((row, rowIndex) => `<Row>${row.map((cell) => `<Cell${rowIndex === 0 ? ' ss:StyleID="Header"' : ""}><Data ss:Type="String">${escapeXml(cell)}</Data></Cell>`).join("")}</Row>`).join("");
    return `<Worksheet ss:Name="${name}"><Table>${table}</Table><WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane></WorksheetOptions></Worksheet>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8"?><?mso-application progid="Excel.Sheet"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"><Styles><Style ss:ID="Default"><Alignment ss:Vertical="Top" ss:WrapText="1"/></Style><Style ss:ID="Header"><Font ss:Bold="1"/><Interior ss:Color="#D8E6DF" ss:Pattern="Solid"/></Style></Styles>${worksheets}</Workbook>`;
}

export function createHtml(items, title = "Poke Poke 学习库", initialSort = "newest", showTypeTabs = false) {
  const languageNames = { en: "英语", fr: "法语", de: "德语", ko: "韩语", es: "西班牙语", ja: "日语", it: "意大利语", pt: "葡萄牙语", ru: "俄语" };
  const cards = items.map((item) => {
    const language = languageNames[item.sourceLanguage] ?? item.sourceLanguage ?? "英语";
    const detail = item.meanings?.length
      ? `<ol class="meanings translation">${item.meanings.map((meaning) => `<li><strong>${escapeHtml(partName(meaning.partOfSpeech))}</strong> ${escapeHtml(meaning.definitionZh ?? meaning.definitionEn ?? "")}</li>`).join("")}</ol>`
      : `<p class="translation">${escapeHtml(item.kind === "sentence" ? item.translationZh ?? "" : item.chineseDefinition ?? "")}</p>`;
    const ipa = item.ipa ? `<p class="ipa">${escapeHtml(item.ipa)}</p>` : "";
    const collocations = item.collocations?.length ? `<div class="collocations"><strong>固定搭配</strong><ul>${item.collocations.map(({ phrase, meaningZh }) => `<li><strong>${escapeHtml(phrase)}</strong><span class="translation"> — ${escapeHtml(meaningZh)}</span></li>`).join("")}</ul></div>` : "";
    const date = item.createdAt ? new Date(item.createdAt).toLocaleString("zh-CN") : "";
    const sourceUrl = safeHttpUrl(item.source?.pageUrl);
    const sourceLabel = shortSourceLabel(item.source?.pageTitle);
    const sourceTitle = sourceLabel ? sourceUrl ? `<a href="${escapeHtml(sourceUrl)}" target="_blank" rel="noopener noreferrer">${escapeHtml(sourceLabel)}</a>` : escapeHtml(sourceLabel) : "";
    const meta = [escapeHtml(language), sourceTitle, escapeHtml(date)].filter(Boolean).join(" · ");
    const kind = item.kind === "sentence" ? "sentences" : "vocabulary";
    return `<article class="item" data-kind="${kind}" data-text="${escapeHtml(String(item.text ?? "").toLocaleLowerCase())}" data-created-at="${escapeHtml(item.createdAt ?? "")}"><h2>${escapeHtml(item.text ?? "")}</h2>${detail}${ipa}${collocations}<p class="meta">${meta}</p></article>`;
  }).join("\n");
  const sortOptions = [["newest", "添加时间倒序"], ["oldest", "添加时间正序"], ["alphabetical", "字母顺序 A–Z"]]
    .map(([value, label]) => `<option value="${value}"${value === initialSort ? " selected" : ""}>${label}</option>`).join("");
  const typeTabs = showTypeTabs ? `<div class="type-tabs"><button class="active" data-view="vocabulary">生词库</button><button data-view="sentences">句子库</button></div>` : "";
  const initialView = showTypeTabs ? "vocabulary" : "all";
  return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${escapeHtml(title)}</title><style>:root{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}*{box-sizing:border-box}body{max-width:860px;margin:0 auto;padding:24px;background:#f5f1e8;color:#25231e}header{margin:-24px -24px 20px;padding:24px;background:#173f35;color:white}.header-controls,.controls,.type-tabs{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.header-controls{justify-content:space-between}.controls{margin-left:auto}button,select{margin-top:12px;border:0;border-radius:8px;padding:8px 12px;background:#e4ded0;cursor:pointer}.type-tabs button.active{background:#fff;color:#173f35;font-weight:700}.item{margin:10px 0;padding:16px;border:1px solid #ddd5c6;border-radius:12px;background:#fffdf8}.item h2{margin:0 0 8px;font-size:18px}.item p{color:#625d53}.ipa{color:#355f50;font-family:Georgia,serif}.meanings{padding-left:24px}.collocations{margin-top:12px;color:#355f50}.collocations ul{margin:5px 0;padding-left:22px}.meta{color:#938b7d!important;font-size:11px}.translations-hidden .translation{display:none}</style></head><body><header><h1>${escapeHtml(title)}</h1><div class="header-controls">${typeTabs}<div class="controls"><select id="sort" aria-label="学习记录排序">${sortOptions}</select><button id="toggle">隐藏翻译</button></div></div></header><main>${cards || "<p>没有学习记录。</p>"}</main><script>var main=document.querySelector("main"),sort=document.getElementById("sort"),view="${initialView}";function applyView(){main.querySelectorAll(".item").forEach(function(card){card.hidden=view!=="all"&&card.dataset.kind!==view})}function reorder(){var cards=Array.from(main.querySelectorAll(".item"));cards.sort(function(a,b){if(sort.value==="alphabetical")return a.dataset.text.localeCompare(b.dataset.text);var order=(a.dataset.createdAt||"").localeCompare(b.dataset.createdAt||"");return sort.value==="oldest"?order:-order});cards.forEach(function(card){main.appendChild(card)});applyView()}document.querySelectorAll("[data-view]").forEach(function(button){button.onclick=function(){view=this.dataset.view;document.querySelectorAll("[data-view]").forEach(function(item){item.classList.toggle("active",item===button)});applyView()}});sort.onchange=reorder;reorder();document.getElementById("toggle").onclick=function(){var hidden=document.body.classList.toggle("translations-hidden");this.textContent=hidden?"显示翻译":"隐藏翻译"}</script></body></html>`;
}

function shortSourceLabel(value) {
  return String(value ?? "").split(/\s[-–—|]\s|-/)[0].trim();
}

function sourceOrigin(value) {
  try {
    return new URL(value).origin;
  } catch (_error) {
    return "";
  }
}

function safeHttpUrl(value) {
  try {
    const url = new URL(value);
    return ["http:", "https:"].includes(url.protocol) ? url.href : "";
  } catch (_error) {
    return "";
  }
}

function partName(value) {
  return ({ preferred: "首选释义", contextPhrase: "语境短语", noun: "名词", verb: "动词", adjective: "形容词", adverb: "副词", phrase: "短语" })[value] ?? value ?? "";
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[character]);
}

function escapeXml(value) {
  return String(value).replace(/[&<>]/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" })[character]);
}
