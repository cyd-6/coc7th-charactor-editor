/* Exact decimal multiplication, shared by the browser and Node verification. */
(function (root) {
  "use strict";
  function decimal(value) {
    const text = String(value).replace(/\$/g, "").replace(/,/g, "").replace(/USD/g, "").trim();
    const match = /^\+?(\d+(?:\.\d*)?|\.\d+)(?:e([+-]?\d+))?$/i.exec(text);
    if (!match || text.length > 400) throw new Error("美元金额请输入非负数");
    const exponent = Number(match[2] || 0);
    if (!Number.isInteger(exponent) || Math.abs(exponent) > 300) throw new Error("金额超出可计算范围");
    const [whole, fraction = ""] = match[1].split(".");
    let n = BigInt((whole || "0") + fraction), scale = fraction.length - exponent;
    if (scale < 0) { n *= 10n ** BigInt(-scale); scale = 0; }
    return { n, scale };
  }
  function formatAmount(amount, rate) {
    const a = decimal(amount), r = decimal(rate);
    if (a.n > 10n ** BigInt(15 + a.scale)) throw new Error("美元金额不能超过 1,000,000,000,000,000");
    const denominator = 10n ** BigInt(a.scale + r.scale);
    const cents = (a.n * r.n * 100n + denominator / 2n) / denominator;
    return `${(cents / 100n).toLocaleString("en-US")}.${String(cents % 100n).padStart(2, "0")}`;
  }
  function migrateCode(code, year) {
    if (code === "FRF" && year >= 1920 && year <= 1959) return "FRF_OLD";
    if (code === "GRR" && year >= 1925 && year <= 1947) return "RM";
    return code;
  }
  const api = { formatAmount, migrateCode };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Coc7Currency = api;
})(globalThis);
