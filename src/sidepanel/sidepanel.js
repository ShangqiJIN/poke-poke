import { createExcelXml, createHtml } from "../shared/export.js";
import { groupItemsBySource, normalizeSourceUrl, removeCollectionItems } from "../shared/collections.js";

let library = { vocabulary: [], sentences: [] };
let activeTab = "vocabulary";
let searchQuery = "";
let activeLanguage = "all";
let visibleItems = [];
let translationsVisible = true;
let sortOrder = "newest";
let collections = [];
let activeCollection = "default";
let editorCollectionId = null;
let editorCheckedIds = new Set();
let editorSourceUrls = new Set();
let managerSelectedIds = new Set();
let returnToManager = false;
const selectedIds = new Set();
const languageNames = { en: "英语", fr: "法语", de: "德语", ko: "韩语", es: "西班牙语", ja: "日语", it: "意大利语", pt: "葡萄牙语", ru: "俄语" };

document.querySelectorAll("nav button").forEach((button) => {
  button.addEventListener("click", () => {
    setActiveTab(button.dataset.tab);
    render();
  });
});

document.querySelector("#search").addEventListener("input", (event) => {
  searchQuery = event.target.value.trim().toLocaleLowerCase();
  selectedIds.clear();
  render();
});

document.querySelector("#language-filter").addEventListener("change", (event) => {
  activeLanguage = event.target.value;
  selectedIds.clear();
  render();
});

document.querySelector("#sort-order").addEventListener("change", (event) => {
  sortOrder = event.target.value;
  render();
});

document.querySelector("#collection-filter").addEventListener("change", async (event) => {
  if (event.target.value === "manage") {
    event.target.value = activeCollection;
    openCollectionManager();
    return;
  }
  activeCollection = event.target.value;
  selectedIds.clear();
  ensureCollectionTab();
  updateCollectionFilter();
  render();
});

document.querySelector("#close-collection-dialog").addEventListener("click", closeCollectionEditor);
document.querySelector("#close-collection-manager").addEventListener("click", closeCollectionEditor);
document.querySelector("#new-collection").addEventListener("click", () => openCollectionEditor(null, true));
document.querySelector("#delete-collections").addEventListener("click", deleteSelectedCollections);
document.querySelector("#cancel-collection").addEventListener("click", () => returnToManager ? openCollectionManager() : closeCollectionEditor());
document.querySelector("#collection-select-all").addEventListener("click", () => {
  itemsForSources(editorSourceUrls).forEach((item) => editorCheckedIds.add(item.id));
  renderCollectionItems();
});
document.querySelector("#collection-select-none").addEventListener("click", () => {
  editorCheckedIds.clear();
  renderCollectionItems();
});
document.querySelector("#save-collection").addEventListener("click", saveCollection);

document.querySelector("#select-all").addEventListener("click", () => {
  const allSelected = visibleItems.length > 0 && visibleItems.every((item) => selectedIds.has(item.id));
  visibleItems.forEach((item) => allSelected ? selectedIds.delete(item.id) : selectedIds.add(item.id));
  render();
});

document.querySelector("#toggle-translations").addEventListener("click", () => {
  translationsVisible = !translationsVisible;
  document.querySelector("#list").classList.toggle("translations-hidden", !translationsVisible);
  document.querySelector("#toggle-translations").textContent = translationsVisible ? "隐藏翻译" : "显示翻译";
});

document.querySelector("#export-excel").addEventListener("click", () => {
  const items = exportItems();
  const vocabulary = items.filter((item) => item.kind !== "sentence");
  const sentences = items.filter((item) => item.kind === "sentence");
  downloadText(createExcelXml(vocabulary, sentences), exportFilename("xls"), "application/vnd.ms-excel;charset=utf-8");
  setNotice(`已导出 Excel：${vocabulary.length} 个词语 · ${sentences.length} 个句子。`);
});

document.querySelector("#export-html").addEventListener("click", () => {
  const items = exportItems();
  const kinds = new Set(items.map((item) => item.kind === "sentence" ? "sentences" : "vocabulary"));
  downloadText(createHtml(items, exportTitle(kinds), sortOrder, kinds.size > 1), exportFilename("html", kinds), "text/html;charset=utf-8");
  setNotice(`已导出 ${items.length} 条 HTML 记录。`);
});

document.querySelector("#delete-selected").addEventListener("click", async () => {
  const selectedCurrent = visibleItems.filter((item) => selectedIds.has(item.id));
  if (!selectedCurrent.length) return;
  const count = selectedCurrent.length;
  if (activeCollection !== "default") {
    const collection = collections.find((item) => item.id === activeCollection);
    if (!collection || !window.confirm(`确定将选中的 ${count} 条记录移出子库“${collection.name}”吗？原学习记录不会被删除。`)) return;
    collection.itemIds = removeCollectionItems(collection, selectedCurrent.map((item) => item.id));
    collection.updatedAt = new Date().toISOString();
    await chrome.storage.local.set({ collections });
    selectedCurrent.forEach((item) => selectedIds.delete(item.id));
    setNotice(`已从子库移出 ${count} 条记录。`);
    render();
    return;
  }
  if (!window.confirm(`确定从学习库删除选中的 ${count} 条记录吗？请先导出备份；此操作无法在插件内撤销。`)) return;
  const response = await chrome.runtime.sendMessage({
    type: "delete-library-items",
    payload: {
      kind: activeTab === "sentences" ? "sentence" : "vocabulary",
      ids: selectedCurrent.map((item) => item.id)
    }
  });
  if (!response?.ok) {
    setNotice(response?.error ?? "删除失败。", true);
    return;
  }
  selectedCurrent.forEach((item) => selectedIds.delete(item.id));
  setNotice(`已删除 ${response.summary.deleted} 条记录。`);
  await load();
});

chrome.storage.onChanged.addListener(load);
load();

async function load() {
  const response = await chrome.runtime.sendMessage({ type: "get-library" });
  if (!response?.ok) {
    setNotice(response?.error ?? "无法读取学习库。", true);
    return;
  }
  library = response.library;
  const stored = await chrome.storage.local.get("collections");
  collections = Array.isArray(stored.collections) ? stored.collections : [];
  const existingIds = new Set([...library.vocabulary, ...library.sentences].map((item) => item.id));
  [...selectedIds].forEach((id) => {
    if (!existingIds.has(id)) selectedIds.delete(id);
  });
  document.querySelector("#summary").textContent =
    `${library.vocabulary.length} 个词语 · ${library.sentences.length} 个句子`;
  updateLanguageFilter();
  ensureCollectionTab();
  updateCollectionFilter();
  render();
}

function itemLanguage(item) {
  return item.sourceLanguage || "en";
}

function languageLabel(code) {
  return languageNames[code] ?? `其他语言（${code}）`;
}

function updateLanguageFilter() {
  const select = document.querySelector("#language-filter");
  const counts = new Map();
  [...library.vocabulary, ...library.sentences].forEach((item) => {
    const language = itemLanguage(item);
    counts.set(language, (counts.get(language) || 0) + 1);
  });
  const available = [...counts.keys()].sort((a, b) => languageLabel(a).localeCompare(languageLabel(b), "zh-CN"));
  if (activeLanguage !== "all" && !counts.has(activeLanguage)) activeLanguage = "all";
  select.replaceChildren(new Option("全部语言", "all"));
  available.forEach((language) => select.appendChild(new Option(`${languageLabel(language)}（${counts.get(language)}）`, language)));
  select.value = activeLanguage;
}

function render() {
  const list = document.querySelector("#list");
  visibleItems = filteredItems(activeTab);

  list.replaceChildren();
  updateSelectionControls();

  if (!visibleItems.length) {
    const empty = document.createElement("p");
    empty.className = "empty";
    empty.textContent = searchQuery
      ? "没有找到匹配的学习记录。"
      : activeTab === "vocabulary"
        ? "在网页中选择英文词语后，它会出现在这里。"
        : "选择一个英文句子开始分析。";
    list.appendChild(empty);
    return;
  }

  let renderedLanguage = "";
  const groupedItems = orderedVisibleItems();
  groupedItems.forEach((item) => {
    renderedLanguage = renderItem(item, groupedItems, renderedLanguage);
  });
}

function filteredItems(tab) {
  const collectionIds = activeCollection === "default"
    ? null
    : new Set(collections.find((collection) => collection.id === activeCollection)?.itemIds ?? []);
  return (library[tab] ?? []).filter((item) => {
    if (collectionIds && !collectionIds.has(item.id)) return false;
    if (activeLanguage !== "all" && itemLanguage(item) !== activeLanguage) return false;
    if (!searchQuery) return true;
    const searchable = [
      item.text,
      item.chineseDefinition,
      item.translationZh,
      item.sourceLanguage,
      item.source?.pageTitle,
      ...(item.collocations ?? []).flatMap((entry) => [entry.phrase, entry.meaningZh])
    ].filter(Boolean).join(" ").toLocaleLowerCase();
    return searchable.includes(searchQuery);
  });
}

function renderItem(item, groupedItems, renderedLanguage) {
    const list = document.querySelector("#list");
    const itemLanguageCode = itemLanguage(item);
    if (itemLanguageCode !== renderedLanguage) {
      renderedLanguage = itemLanguageCode;
      const group = document.createElement("h2");
      group.className = "language-group";
      const count = groupedItems.filter((candidate) => itemLanguage(candidate) === itemLanguageCode).length;
      group.textContent = `${languageLabel(itemLanguageCode)} · ${count}`;
      list.appendChild(group);
    }
    const article = document.createElement("article");
    article.className = "item";
    const head = document.createElement("div");
    head.className = "item-head";
    const selection = document.createElement("label");
    selection.className = "item-select";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = selectedIds.has(item.id);
    checkbox.setAttribute("aria-label", `选择 ${item.text}`);
    checkbox.addEventListener("change", () => {
      checkbox.checked ? selectedIds.add(item.id) : selectedIds.delete(item.id);
      updateSelectionControls();
    });
    const title = document.createElement("h2");
    title.textContent = item.text;
    selection.append(checkbox, title);
    const speak = document.createElement("button");
    speak.className = "speak";
    speak.appendChild(createSpeakerIcon());
    speak.title = "朗读原文";
    speak.setAttribute("aria-label", "朗读原文");
    speak.addEventListener("click", () => speakText(item.text, item.sourceLanguage || "en"));
    head.append(selection, speak);
    const detail = document.createElement("p");
    detail.className = "translation";
    detail.textContent = item.kind === "sentence" ? item.translationZh : item.chineseDefinition;
    article.append(head, detail);
    if (item.ipa) {
      const ipa = document.createElement("p");
      ipa.className = "ipa";
      ipa.textContent = item.ipa;
      article.appendChild(ipa);
    }
    if (item.meanings?.length) {
      detail.remove();
      const meanings = document.createElement("ol");
      meanings.className = "meanings translation";
      item.meanings.forEach((meaning) => {
        const row = document.createElement("li");
        const part = document.createElement("strong");
        part.textContent = partOfSpeechName(meaning.partOfSpeech);
        row.append(part, document.createTextNode(` ${meaning.definitionZh || meaning.definitionEn}`));
        meanings.appendChild(row);
      });
      article.appendChild(meanings);
    }
    if (item.collocations?.length) {
      const collocations = document.createElement("ul");
      collocations.className = "collocations";
      item.collocations.forEach(({ phrase, meaningZh }) => {
        const row = document.createElement("li");
        const strong = document.createElement("strong");
        strong.textContent = phrase;
        const translation = document.createElement("span");
        translation.className = "translation";
        translation.textContent = " — " + meaningZh;
        row.append(strong, translation);
        collocations.appendChild(row);
      });
      article.appendChild(collocations);
    }
    const meta = document.createElement("p");
    meta.className = "meta";
    const date = item.createdAt ? new Date(item.createdAt).toLocaleString("zh-CN") : "时间未知";
    const language = languageLabel(itemLanguage(item));
    meta.append(document.createTextNode(language));
    if (item.source?.pageTitle) {
      meta.append(document.createTextNode(" · "));
      if (/^https?:\/\//.test(item.source?.pageUrl ?? "")) {
        const source = document.createElement("a");
        source.href = item.source.pageUrl;
        source.target = "_blank";
        source.rel = "noopener noreferrer";
        source.textContent = sourceLabel(item.source.pageTitle);
        meta.appendChild(source);
      } else {
        meta.append(document.createTextNode(sourceLabel(item.source.pageTitle)));
      }
    }
    if (date) meta.append(document.createTextNode(` · ${date}`));
    article.appendChild(meta);
    list.appendChild(article);
    return itemLanguageCode;
}

function orderedVisibleItems() {
  return orderedItems(visibleItems);
}

function orderedItems(items) {
  return [...items].sort((a, b) => {
    const languageOrder = languageLabel(itemLanguage(a)).localeCompare(languageLabel(itemLanguage(b)), "zh-CN");
    if (languageOrder) return languageOrder;
    if (sortOrder === "alphabetical") return String(a.text ?? "").localeCompare(String(b.text ?? ""), itemLanguage(a), { sensitivity: "base" });
    const dateOrder = String(a.createdAt ?? "").localeCompare(String(b.createdAt ?? ""));
    return sortOrder === "oldest" ? dateOrder : -dateOrder;
  });
}

function exportItems() {
  const all = orderedItems(["vocabulary", "sentences"].flatMap((tab) => filteredItems(tab)));
  const selected = all.filter((item) => selectedIds.has(item.id));
  return selected.length ? selected : orderedVisibleItems();
}

function exportTitle(kinds) {
  const collection = collections.find((item) => item.id === activeCollection)?.name;
  if (collection) return `Poke Poke · ${collection}`;
  const scopeName = kinds.size > 1 ? "学习库" : kinds.has("sentences") ? "句子库" : "生词库";
  return `Poke Poke ${scopeName}`;
}

function partOfSpeechName(partOfSpeech) {
  const names = { preferred: "首选释义", contextPhrase: "语境短语", phrase: "短语", noun: "名词", verb: "动词", adjective: "形容词", adverb: "副词", pronoun: "代词", preposition: "介词", conjunction: "连词", interjection: "感叹词" };
  return names[partOfSpeech] ?? partOfSpeech;
}

function updateSelectionControls() {
  const selectedVisible = visibleItems.filter((item) => selectedIds.has(item.id)).length;
  const selectedCurrent = selectedVisible;
  const selectAll = document.querySelector("#select-all");
  selectAll.disabled = visibleItems.length === 0;
  selectAll.textContent = selectedVisible === visibleItems.length && visibleItems.length
    ? "取消全选"
    : "全选当前结果";
  const deleteButton = document.querySelector("#delete-selected");
  deleteButton.disabled = selectedCurrent === 0;
  const deleteLabel = activeCollection === "default" ? "删除选中项" : "移出子库";
  deleteButton.textContent = selectedCurrent ? `${deleteLabel}（${selectedCurrent}）` : deleteLabel;
  const words = collectionItems("vocabulary");
  const sentences = collectionItems("sentences");
  const selectedWords = words.filter((item) => selectedIds.has(item.id)).length;
  const selectedSentences = sentences.filter((item) => selectedIds.has(item.id)).length;
  document.querySelector('[data-tab="vocabulary"]').textContent = `生词库 · ${words.length}${selectedWords ? ` (${selectedWords})` : ""}`;
  document.querySelector('[data-tab="sentences"]').textContent = `句子库 · ${sentences.length}${selectedSentences ? ` (${selectedSentences})` : ""}`;
  document.querySelector("#export-excel").textContent = "导出 Excel";
  document.querySelector("#export-html").textContent = "导出 HTML";
}

function updateCollectionFilter() {
  const select = document.querySelector("#collection-filter");
  if (activeCollection !== "default" && !collections.some((collection) => collection.id === activeCollection)) activeCollection = "default";
  select.replaceChildren(new Option("默认", "default"));
  collections.forEach((collection) => select.appendChild(new Option(collection.name, collection.id)));
  select.appendChild(new Option("编辑子库…", "manage"));
  select.value = activeCollection;
}

function collectionItems(tab) {
  if (activeCollection === "default") return library[tab] ?? [];
  const itemIds = new Set(collections.find((collection) => collection.id === activeCollection)?.itemIds ?? []);
  return (library[tab] ?? []).filter((item) => itemIds.has(item.id));
}

function setActiveTab(tab) {
  activeTab = tab;
  document.querySelectorAll("nav button").forEach((button) => button.classList.toggle("active", button.dataset.tab === tab));
}

function ensureCollectionTab() {
  if (activeCollection === "default") return;
  const itemIds = new Set(collections.find((collection) => collection.id === activeCollection)?.itemIds ?? []);
  if ((library[activeTab] ?? []).some((item) => itemIds.has(item.id))) return;
  const alternative = activeTab === "vocabulary" ? "sentences" : "vocabulary";
  if ((library[alternative] ?? []).some((item) => itemIds.has(item.id))) setActiveTab(alternative);
}

function sourceGroups() {
  const groups = new Map();
  groupItemsBySource([...library.vocabulary, ...library.sentences]).forEach((items, url) => {
    groups.set(url, { title: sourceLabel(items[0]?.source?.pageTitle) || url, items });
  });
  return groups;
}

function itemsForSource(url) {
  return sourceGroups().get(url)?.items ?? [];
}

function itemsForSources(urls) {
  return [...urls].flatMap((url) => itemsForSource(url));
}

function openCollectionManager() {
  returnToManager = false;
  editorCollectionId = null;
  editorCheckedIds.clear();
  editorSourceUrls.clear();
  managerSelectedIds.clear();
  document.querySelector("#collection-dialog-title").textContent = "编辑子库";
  document.querySelector("#collection-manager").hidden = false;
  document.querySelector("#collection-editor").hidden = true;
  renderCollectionManager();
  const dialog = document.querySelector("#collection-dialog");
  if (!dialog.open) dialog.showModal();
}

function renderCollectionManager() {
  const container = document.querySelector("#collection-list");
  container.replaceChildren();
  collections.forEach((collection) => {
    const row = document.createElement("div");
    row.className = "collection-row";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.setAttribute("aria-label", `选择删除 ${collection.name}`);
    checkbox.checked = managerSelectedIds.has(collection.id);
    checkbox.addEventListener("change", () => {
      checkbox.checked ? managerSelectedIds.add(collection.id) : managerSelectedIds.delete(collection.id);
      document.querySelector("#delete-collections").disabled = managerSelectedIds.size === 0;
    });
    const name = document.createElement("span");
    name.textContent = collection.name;
    const open = document.createElement("button");
    open.className = "open-collection";
    open.textContent = "›";
    open.setAttribute("aria-label", `编辑 ${collection.name}`);
    open.addEventListener("click", () => openCollectionEditor(collection.id, true));
    row.append(checkbox, name, open);
    container.appendChild(row);
  });
  if (!collections.length) {
    const empty = document.createElement("p");
    empty.className = "manager-empty";
    empty.textContent = "还没有自定义子库。";
    container.appendChild(empty);
  }
  document.querySelector("#delete-collections").disabled = managerSelectedIds.size === 0;
}

function openCollectionEditor(collectionId = null, fromManager = false) {
  const groups = sourceGroups();
  if (!groups.size) {
    setNotice("当前学习记录中没有可用的来源 URL。", true);
    return;
  }
  const collection = collections.find((item) => item.id === collectionId);
  returnToManager = fromManager;
  const selected = [...library.vocabulary, ...library.sentences].filter((item) => selectedIds.has(item.id));
  const selectedUrls = [...new Set(selected.map((item) => normalizeSourceUrl(item.source?.pageUrl)).filter(Boolean))];
  const savedUrls = collection?.sourceUrls?.length ? collection.sourceUrls : collection?.sourceUrl ? [collection.sourceUrl] : [];
  const initialUrls = savedUrls.length ? savedUrls : selectedUrls.length ? selectedUrls : [groups.keys().next().value];
  editorCollectionId = collection?.id ?? null;
  editorSourceUrls = new Set(initialUrls.filter((url) => groups.has(url)));
  editorCheckedIds = new Set(collection?.itemIds ?? (selectedUrls.length ? selected.map((item) => item.id) : itemsForSources(editorSourceUrls).map((item) => item.id)));
  document.querySelector("#collection-dialog-title").textContent = collection ? "编辑子库" : "新建子库";
  document.querySelector("#collection-manager").hidden = true;
  document.querySelector("#collection-editor").hidden = false;
  document.querySelector("#collection-name").value = collection?.name || sourceLabel(itemsForSources(editorSourceUrls)[0]?.source?.pageTitle) || "新子库";
  renderCollectionSources(groups);
  renderCollectionItems();
  const dialog = document.querySelector("#collection-dialog");
  if (!dialog.open) dialog.showModal();
}

function renderCollectionSources(groups = sourceGroups()) {
  const container = document.querySelector("#collection-sources");
  container.replaceChildren();
  groups.forEach((group, url) => {
    const label = document.createElement("label");
    label.className = "collection-source";
    label.title = url;
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = editorSourceUrls.has(url);
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) {
        editorSourceUrls.add(url);
        group.items.forEach((item) => editorCheckedIds.add(item.id));
      } else {
        editorSourceUrls.delete(url);
        group.items.forEach((item) => editorCheckedIds.delete(item.id));
      }
      renderCollectionItems();
    });
    label.append(checkbox, document.createTextNode(`${group.title}（${group.items.length}）`));
    container.appendChild(label);
  });
}

function renderCollectionItems() {
  const container = document.querySelector("#collection-items");
  container.replaceChildren();
  itemsForSources(editorSourceUrls).forEach((item) => {
    const label = document.createElement("label");
    label.className = "collection-item";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = editorCheckedIds.has(item.id);
    checkbox.addEventListener("change", () => checkbox.checked ? editorCheckedIds.add(item.id) : editorCheckedIds.delete(item.id));
    const text = document.createElement("span");
    const kind = document.createElement("span");
    kind.className = "collection-kind";
    kind.textContent = item.kind === "sentence" ? "句子 · " : "词语 · ";
    text.append(kind, document.createTextNode(item.text));
    label.append(checkbox, text);
    container.appendChild(label);
  });
}

async function saveCollection() {
  const name = document.querySelector("#collection-name").value.trim();
  const sourceUrls = [...editorSourceUrls];
  const validIds = new Set(itemsForSources(sourceUrls).map((item) => item.id));
  const itemIds = [...editorCheckedIds].filter((id) => validIds.has(id));
  if (!name || !sourceUrls.length || !itemIds.length) {
    setNotice(!name ? "请输入子库名称。" : !sourceUrls.length ? "请至少选择一个来源 URL。" : "请至少纳入一条记录。", true);
    return;
  }
  const existing = collections.find((item) => item.id === editorCollectionId);
  if (existing) {
    Object.assign(existing, { name, sourceUrl: sourceUrls[0], sourceUrls, itemIds, updatedAt: new Date().toISOString() });
  } else {
    const collection = { id: `collection-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`, name, sourceUrl: sourceUrls[0], sourceUrls, itemIds, createdAt: new Date().toISOString() };
    collections.push(collection);
    activeCollection = collection.id;
  }
  await chrome.storage.local.set({ collections });
  selectedIds.clear();
  ensureCollectionTab();
  updateCollectionFilter();
  render();
  setNotice(`已保存子库“${name}”，包含 ${itemIds.length} 条记录。`);
  if (returnToManager) openCollectionManager();
  else closeCollectionEditor();
}

async function deleteSelectedCollections() {
  if (!managerSelectedIds.size || !window.confirm(`确定删除选中的 ${managerSelectedIds.size} 个子库吗？学习记录不会被删除。`)) return;
  const deletingActive = managerSelectedIds.has(activeCollection);
  collections = collections.filter((item) => !managerSelectedIds.has(item.id));
  if (deletingActive) {
    activeCollection = "default";
    setActiveTab("vocabulary");
  }
  await chrome.storage.local.set({ collections });
  managerSelectedIds.clear();
  updateCollectionFilter();
  render();
  renderCollectionManager();
}

function closeCollectionEditor() {
  document.querySelector("#collection-dialog").close();
  editorCollectionId = null;
  editorCheckedIds.clear();
  editorSourceUrls.clear();
  managerSelectedIds.clear();
  returnToManager = false;
}

function sourceLabel(value) {
  return String(value ?? "").split(/\s[-–—|]\s|-/)[0].trim();
}

function createSpeakerIcon() {
  const namespace = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(namespace, "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("aria-hidden", "true");
  const speaker = document.createElementNS(namespace, "path");
  speaker.setAttribute("d", "M3 9v6h4l5 4V5L7 9H3z");
  const waves = document.createElementNS(namespace, "path");
  waves.setAttribute("d", "M14 8.5v2.1a2 2 0 0 1 0 2.8v2.1a5 5 0 0 0 0-7zm0-4v2.05a7 7 0 0 1 0 10.9v2.05a9 9 0 0 0 0-15z");
  svg.append(speaker, waves);
  return svg;
}

function downloadText(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function exportFilename(extension, kinds = new Set(exportItems().map((item) => item.kind === "sentence" ? "sentences" : "vocabulary"))) {
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const scope = kinds.size > 1 ? "library" : kinds.has("sentences") ? "sentences" : "vocabulary";
  return `poke-poke-${scope}-${stamp}.${extension}`;
}

function setNotice(message, isError = false) {
  const notice = document.querySelector("#notice");
  notice.textContent = message;
  notice.style.color = isError ? "#8d3b32" : "#347453";
}

function speakText(text, language = "en") {
  if (!("speechSynthesis" in window)) return;
  window.speechSynthesis.cancel();
  const utterance = new SpeechSynthesisUtterance(text);
  const voiceLocales = { en: "en-US", fr: "fr-FR", de: "de-DE", ko: "ko-KR", es: "es-ES", ja: "ja-JP", it: "it-IT", pt: "pt-PT", ru: "ru-RU" };
  utterance.lang = voiceLocales[language] ?? language;
  const sentenceRates = { ko: 0.92, ja: 0.9, de: 0.9, fr: 0.92, en: 0.88 };
  const wordRates = { ko: 0.82, ja: 0.82, de: 0.8, fr: 0.82, en: 0.8 };
  utterance.rate = text.includes(" ") ? (sentenceRates[language] ?? 0.9) : (wordRates[language] ?? 0.82);
  utterance.pitch = 1;
  const voices = window.speechSynthesis.getVoices();
  const requestedLocale = utterance.lang.toLocaleLowerCase();
  const matchingVoices = voices.filter((voice) => voice.lang.toLocaleLowerCase().startsWith(language));
  utterance.voice = matchingVoices.sort((a, b) => voiceScore(b, requestedLocale, language) - voiceScore(a, requestedLocale, language))[0] ?? null;
  window.speechSynthesis.speak(utterance);
}

function voiceScore(voice, requestedLocale, language) {
  const locale = voice.lang.toLocaleLowerCase();
  const preferredNames = { ko: ["yuna"], ja: ["kyoko", "otoya"], de: ["anna"], fr: ["amelie", "thomas"], en: ["samantha", "alex"] };
  const preferred = (preferredNames[language] || []).some((name) => voice.name.toLocaleLowerCase().includes(name));
  return (locale === requestedLocale ? 100 : 0) + (preferred ? 60 : 0) + (voice.localService ? 20 : 0) + (voice.default ? 5 : 0);
}
