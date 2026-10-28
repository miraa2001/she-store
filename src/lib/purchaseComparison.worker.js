import { generateComparisons } from "./purchaseComparison.js";

self.onmessage = ({ data }) => {
  try {
    self.postMessage({ plans: generateComparisons(data.purchases, data.groups) });
  } catch (error) {
    self.postMessage({ error: error.message || "تعذرت المقارنة" });
  }
};
