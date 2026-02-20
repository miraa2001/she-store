export function normalizeText(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[أإآ]/g, "ا")
    .replace(/ى/g, "ي")
    .replace(/\s+/g, " ");
}

function includesNormalized(value, normalizedQuery) {
  if (!normalizedQuery) return true;
  return normalizeText(value).includes(normalizedQuery);
}

export function searchByName(list, query, accessor) {
  const normalizedQuery = normalizeText(query);
  if (!normalizedQuery) return list || [];

  return (list || []).filter((item) => {
    const values = accessor(item);
    if (Array.isArray(values)) {
      return values.some((value) => includesNormalized(value, normalizedQuery));
    }
    return includesNormalized(values, normalizedQuery);
  });
}

