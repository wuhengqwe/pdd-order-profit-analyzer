const assert = require("node:assert/strict");
const test = require("node:test");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const site = path.join(__dirname, "..", "docs");
const XLSX = require(path.join(site, "vendor/xlsx.full.min.js"));
const fflate = require(path.join(site, "vendor/fflate.js"));
const scope = { XLSX, fflate, TextEncoder, TextDecoder, File, AbortSignal };
const messages = [];
let onMessage;
scope.self = {
  addEventListener(type, callback) {
    if (type === "message") onMessage = callback;
  },
  postMessage(message) { messages.push(message); },
};
scope.importScripts = () => {};
const context = vm.createContext(scope);
for (const name of ["analysis.js", "reports.js", "worker.js"]) {
  vm.runInContext(fs.readFileSync(path.join(site, name), "utf8"), context, { filename: name });
}
const { analyze, DEFAULTS } = scope.PddAnalysis;
function order({ order, id = "900000000001", code = "00101", type = "已收货", service = "", quantity = 1, receipt = 20, date = "2026-10-11" }) {
  return { order, id, code, date, state: type, service, quantity, receipt, buyer: receipt, title: "示例", variant: code, spec: "", warehouse: "" };
}
function run({ orders, promos = [], costs = [], options = {}, info = {} }) {
  return analyze({ orders, promos, costs, start: "2026-08-03", end: "2026-10-31", options, info });
}
const promo = (id, spend) => ({ id, spend, orders: 0, revenue: 0, title: "示例" });

test("按订单号统计，退款与未付过滤，多件保留且重复导出行去重", () => {
  const kept = order({ order: "sample-1", quantity: 3, receipt: 60 });
  const result = run({ orders: [kept, kept, order({ order: "sample-2", type: "已发货，退款成功" }),
    order({ order: "sample-3", type: "待付款", date: "" }), order({ order: "sample-4", type: "已取消" })],
    costs: [{ code: "00101", price: 4, warehouse: "义乌" }] });
  assert.equal(result.stats.raw, 4);
  assert.equal(result.stats.retained, 1);
  assert.equal(result.stats.units, 3);
  assert.equal(result.stats.refunded, 1);
  assert.equal(result.stats.repeated, 1);
  assert.equal(result.specs[0].code, "00101");
  assert.equal(result.specs[0].price, 20);
});

test("先在莆田完成套数匹配，找不到再用义乌，并为缺成本保留空利润", () => {
  const result = run({ orders: [order({ order: "a", code: "00101X2" }),
    order({ order: "b", id: "900000000002", code: "00921" }), order({ order: "c", id: "900000000003", code: "NO-COST" })],
    costs: [{ code: "00101", price: 2.5, warehouse: "莆田" }, { code: "00101X2", price: 8, warehouse: "义乌" },
      { code: "921", price: 3, warehouse: "义乌" }] });
  const find = id => result.links.find(x => x.id.endsWith(id));
  assert.equal(find("001").cost, 5);
  assert.equal(find("001").costWarehouse, "莆田");
  assert.equal(find("002").cost, 3);
  assert.equal(find("002").costWarehouse, "义乌");
  assert.equal(find("003").base, null);
});

test("同仓冲突价格待核对，组合编码可使用跨仓部件", () => {
  const result = run({ orders: [order({ order: "a", code: "A+B" }), order({ order: "b", code: "C", id: "900000000002" })],
    costs: [{ code: "A", price: 2, warehouse: "莆田" }, { code: "B", price: 3, warehouse: "义乌" },
      { code: "C", price: 5, warehouse: "莆田" }, { code: "C", price: 8, warehouse: "莆田" }, { code: "C", price: 2, warehouse: "义乌" }] });
  assert.equal(result.links[0].cost, 5);
  assert.equal(result.links[0].costWarehouse, "莆田 + 义乌");
  assert.equal(result.links[1].base, null);
});

test("税费按实收 5%，推广费全额入账，缺花费不估算", () => {
  const find = (result, id) => result.links.find(x => x.id.endsWith(id));
  const inputs = { orders: [order({ order: "sample", receipt: 100 }), order({ order: "missing", id: "900000000003" })],
    promos: [promo("900000000001", 10), promo("900000000002", 7), promo("900000000003", null)],
    costs: [{ code: "00101", price: 40, warehouse: "莆田" }] };
  const result = run(inputs);
  assert.equal(DEFAULTS.tax, 5);
  assert.equal(find(result, "001").base, 34.9);
  assert.equal(find(result, "001").stress, 34.4);
  assert.equal(find(result, "002").base, -7);
  assert.equal(find(result, "003").base, null);
  assert.equal(result.stats.spend, 17);
  assert.equal(run({ ...inputs, options: { tax: 1.5 } }).links.find(x => x.id === "900000000001").base, 38.4);
});

test("同码同价归并，部分同规格交集保留为待核对", () => {
  const base = [
    order({ order: "1", code: "A", receipt: 20 }), order({ order: "2", code: "B", receipt: 32 }),
    order({ order: "3", id: "900000000002", code: "B", receipt: 32 }), order({ order: "4", id: "900000000002", code: "A", receipt: 20 }),
    order({ order: "5", id: "900000000003", code: "A", receipt: 20 }), order({ order: "6", id: "900000000003", code: "B", receipt: 35 }),
  ];
  const result = run({ orders: base });
  assert.equal(result.templates.length, 2);
  assert.equal(result.templates.find(x => x.linked === 2).orders, 4);
  assert.equal(result.overlaps[0].code, "A");
});

test("同一订单多商品时整店去重并平摊运费，不重复扣费", () => {
  const result = run({ orders: [order({ order: "multi" }), order({ order: "multi", id: "900000000002" })] });
  assert.equal(result.stats.retained, 1);
  assert.equal(result.links[0].orders, 1);
  assert.equal(result.links[1].orders, 1);
  assert.equal(result.links[0].freightUnits + result.links[1].freightUnits, 1);
});

test("演示通过真实 ZIP、XLSX 解析与导出，且导出编码为文本", async () => {
  await onMessage({ data: { kind: "demo" } });
  const completed = messages.find(x => x.kind === "result");
  assert.ok(completed, String(messages.at(-1).text));
  assert.equal(completed.result.stats.retained, 171);
  assert.equal(completed.result.stats.units, 180);
  assert.equal(completed.result.stats.promoGap, 0);
  assert.equal(completed.result.templates[0].action, "优先复发");
  await onMessage({ data: { kind: "export" } });
  const download = messages.at(-1);
  assert.equal(download.kind, "download", download.text);
  const book = XLSX.read(download.binary, { type: "array" });
  assert.equal(book.Sheets["规格成交证据"].D2.t, "s");
  assert.ok(book.Sheets["规格成交证据"].D2.v.startsWith("00"));
  assert.equal(book.SheetNames.length, 7);
});
