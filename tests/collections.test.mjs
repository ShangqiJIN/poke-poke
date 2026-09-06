import assert from "node:assert/strict";
import test from "node:test";
import { groupItemsBySource, normalizeSourceUrl, removeCollectionItems } from "../src/shared/collections.js";

test("groups words and sentences from the same article and ignores URL fragments", () => {
  const items = [
    { id: "word", source: { pageUrl: "https://example.com/story#word" } },
    { id: "sentence", source: { pageUrl: "https://example.com/story#sentence" } },
    { id: "other", source: { pageUrl: "https://example.com/other" } },
    { id: "missing" }
  ];
  const url = normalizeSourceUrl("https://example.com/story#top");
  const groups = groupItemsBySource(items);
  assert.deepEqual(groups.get(url).map((item) => item.id), ["word", "sentence"]);
  assert.equal(groups.size, 2);
});

test("removing a card from a sublibrary does not mutate the source card", () => {
  const card = { id: "word", text: "organic" };
  const collection = { itemIds: [card.id, "sentence"] };
  assert.deepEqual(removeCollectionItems(collection, [card.id]), ["sentence"]);
  assert.equal(card.text, "organic");
});
