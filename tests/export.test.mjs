import assert from "node:assert/strict";
import test from "node:test";
import { createExcelXml, createHtml, createLibraryJson } from "../src/shared/export.js";

test("exports a versioned library json document", () => {
  const parsed = JSON.parse(createLibraryJson({ vocabulary: [{ id: "1" }], sentences: [] }));
  assert.equal(parsed.format, "english-reader-library");
  assert.equal(parsed.schemaVersion, 1);
  assert.equal(parsed.vocabulary[0].id, "1");
});

test("creates a standalone html view with escaped cards and collocations", () => {
  const html = createHtml([{
    kind: "sentence",
    text: "A < B",
    translationZh: "A 小于 B",
    collocations: [{ phrase: "pass by", meaningZh: "经过" }],
    createdAt: "2026-09-01T00:00:00.000Z",
    source: { pageTitle: "Reading-a-b-root", pageUrl: "https://example.com/story" }
  }], "句子库", "alphabetical", true);
  assert.match(html, /<!doctype html>/);
  assert.match(html, /A &lt; B/);
  assert.match(html, /固定搭配/);
  assert.match(html, /pass by/);
  assert.match(html, /id="toggle">隐藏翻译/);
  assert.match(html, /translations-hidden/);
  assert.match(html, /href="https:\/\/example\.com\/story"/);
  assert.match(html, />Reading<\/a>/);
  assert.doesNotMatch(html, /Reading-a-b-root/);
  assert.match(html, /id="sort"/);
  assert.match(html, /value="alphabetical" selected/);
  assert.match(html, /data-created-at="2026-09-01T00:00:00.000Z"/);
  assert.match(html, /data-kind="sentences"/);
  assert.match(html, /data-view="vocabulary">生词库/);
  assert.match(html, /data-view="sentences">句子库/);
  assert.doesNotMatch(html, /data-view="all">全部/);
  assert.match(html, /class="header-controls"/);
  assert.doesNotMatch(html, /条当前筛选记录/);
});

test("creates an Excel workbook with separate vocabulary and sentence sheets", () => {
  const xml = createExcelXml(
    [{ kind: "vocabulary", text: "organic", chineseDefinition: "有机的" }],
    [{ kind: "sentence", text: "It works.", translationZh: "它有效。", source: { pageUrl: "https://example.com/story/chapter-1" } }]
  );
  assert.match(xml, /<Worksheet ss:Name="生词库">/);
  assert.match(xml, /<Worksheet ss:Name="句子库">/);
  assert.match(xml, /organic/);
  assert.match(xml, /It works\./);
  assert.match(xml, /https:\/\/example\.com\/story\/chapter-1/);
});
