export function normalizeSourceUrl(value) {
  try {
    const url = new URL(value);
    url.hash = "";
    return url.href;
  } catch (_error) {
    return "";
  }
}

export function groupItemsBySource(items) {
  const groups = new Map();
  items.forEach((item) => {
    const url = normalizeSourceUrl(item.source?.pageUrl);
    if (!url) return;
    if (!groups.has(url)) groups.set(url, []);
    groups.get(url).push(item);
  });
  return groups;
}

export function removeCollectionItems(collection, ids) {
  const removed = new Set(ids);
  return collection.itemIds.filter((id) => !removed.has(id));
}
