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
function order({ order, id = "900000000001", code = "00101", type = "已收货", service = "", quantity = 1, receipt = 20, date = "2026-10-11", warehouse = "" }) {
  return { order, id, code, date, state: type, service, quantity, receipt, buyer: receipt, title: "示例", variant: code, spec: "", warehouse };
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

test("按莆田 1.5% 和义乌 5% 分开计税，推广费全额入账", () => {
  const find = (result, id) => result.links.find(x => x.id.endsWith(id));
  const inputs = { orders: [order({ order: "sample", receipt: 100 }), order({ order: "missing", id: "900000000003" }),
    order({ order: "yiwu", id: "900000000004", code: "00102", receipt: 100 })],
    promos: [promo("900000000001", 10), promo("900000000002", 7), promo("900000000003", null), promo("900000000004", 10)],
    costs: [{ code: "00101", price: 40, warehouse: "莆田" }, { code: "00102", price: 40, warehouse: "义乌" }] };
  const result = run(inputs);
  assert.equal(DEFAULTS.putianTax, 1.5);
  assert.equal(DEFAULTS.yiwuTax, 5);
  assert.equal(find(result, "001").base, 38.4);
  assert.equal(find(result, "001").stress, 37.9);
  assert.equal(find(result, "001").tax, 1.5);
  assert.equal(find(result, "004").base, 34.9);
  assert.equal(find(result, "004").tax, 5);
  assert.equal(find(result, "002").base, -7);
  assert.equal(find(result, "003").base, null);
  assert.equal(result.stats.spend, 27);
  assert.equal(run({ ...inputs, options: { putianTax: 5 } }).links.find(x => x.id === "900000000001").base, 34.9);
});

test("明确发货仓优先计税，缺仓的组合按两仓成本分摊，未知仓留空", () => {
  const result = run({
    orders: [order({ order: "mix", code: "A+B", receipt: 100 }),
      order({ order: "shipped", id: "900000000002", code: "A", receipt: 100, warehouse: "义乌发货仓" }),
      order({ order: "unknown", id: "900000000003", code: "C", receipt: 100 })],
    costs: [{ code: "A", price: 2, warehouse: "莆田" }, { code: "B", price: 3, warehouse: "义乌" },
      { code: "C", price: 1, warehouse: "" }],
  });
  const [mix, shipped, unknown] = result.links;
  assert.equal(mix.putianReceipt, 40);
  assert.equal(mix.yiwuReceipt, 60);
  assert.equal(mix.tax, 3.6);
  assert.equal(shipped.costWarehouse, "莆田");
  assert.equal(shipped.yiwuReceipt, 100);
  assert.equal(shipped.tax, 5);
  assert.equal(unknown.base, null);
  assert.equal(unknown.taxMissing, 1);
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

test("每个模板取保留单最多的商品 ID，独立于盈利和近期代表链接", () => {
  const high = "900000000001", profitable = "900000000002";
  const orders = [
    ...Array.from({ length: 8 }, (_, i) => order({ order: "old-" + i, id: high, date: "2026-08-19" })),
    ...Array.from({ length: 4 }, (_, i) => order({ order: "new-" + i, id: profitable })),
  ];
  const result = run({ orders, promos: [promo(high, 500)], costs: [{ code: "00101", price: 4, warehouse: "莆田" }] });
  assert.equal(result.templates[0].rep, profitable);
  assert.equal(result.templates[0].orders, 12);
  assert.equal(result.leaders.length, 1);
  assert.equal(result.leaders[0].id, high);
  assert.equal(result.leaders[0].orders, 8);
  assert.equal(result.leaders[0].recent, 0);
  assert.equal(result.leaders[0].spend, 500);
  assert.ok(result.leaders[0].base < 0);
});

test("总花费缺列和单行缺值时回退成交花费，已有零值仍取总花费", async () => {
  const csv = (lines, name) => new File([lines.join("\r\n")],name);
  const orderFile=csv([
    "订单号,商品ID,订单状态,订单成交时间,商家实收,商品数量,规格编码",
    "f1,900000000001,已收货,2026-10-01,50,1,0010",
    "f2,900000000002,已收货,2026-10-01,50,1,0010",
  ],"order.csv");
  const make = promoCsv => scope.PddReportFiles.importFiles({order:[orderFile],promo:[promoCsv],cost:[]},
    {start:"2026-10-01",end:"2026-10-01",options:{}});
  let result=await make(csv(["商品ID,总花费(元),成交花费(元)","900000000001,0,9","900000000002,--,7","总计,7,16"],"promo.csv"));
  assert.equal(result.info.transactionSpendRows,1);
  assert.equal(result.links.find(x=>x.id==="900000000001").spend,0);
  assert.equal(result.stats.spend,7);
  result=await make(csv(["商品ID,成交花费(元)","900000000001,6","900000000002,7","总计,13"],"promo.csv"));
  assert.equal(result.info.transactionSpendRows,2);
  assert.equal(result.stats.spend,13);
  assert.equal(result.stats.promoGap,0);
});

test("演示通过真实 ZIP、XLSX 解析与导出，且导出编码为文本", async () => {
  await onMessage({ data: { kind: "demo" } });
  const completed = messages.find(x => x.kind === "result");
  assert.ok(completed, String(messages.at(-1).text));
  assert.equal(completed.result.stats.retained, 171);
  assert.equal(completed.result.stats.units, 180);
  assert.equal(completed.result.stats.promoGap, 0);
  assert.equal(completed.result.templates[0].action, "优先复发");
  assert.equal(completed.result.templates[0].orders, 67);
  assert.equal(completed.result.templates[0].linked, 2);
  assert.equal(Object.hasOwn(completed.result.templates[0], "orderRate"), false);
  await onMessage({ data: { kind: "export" } });
  const download = messages.at(-1);
  assert.equal(download.kind, "download", download.text);
  const book = XLSX.read(download.binary, { type: "array" });
  assert.equal(book.Sheets["规格成交证据"].D2.t, "s");
  assert.ok(book.Sheets["规格成交证据"].D2.v.startsWith("00"));
  assert.equal(book.SheetNames.length, 8);
  const topLink = XLSX.utils.sheet_to_json(book.Sheets["模板最高单量链接"])[0];
  assert.equal(topLink["最高单量商品ID"], "900000000001");
  assert.equal(topLink["保留单"], 40);
  const template = XLSX.utils.sheet_to_json(book.Sheets["模板清单"])[0];
  assert.equal(template["保留单"], 67);
  assert.deepEqual(template["同组商品ID"].split(", "), ["900000000001", "900000000002"]);
  assert.equal(topLink["同组其他商品ID"], "900000000002");
  assert.equal(topLink["同组保留单"], 67);
  assert.equal(topLink["同组其他链接出单明细"], "900000000002（27单）");
  assert.equal(template["同组链接出单明细"], "900000000001（40单）, 900000000002（27单）");
  const links = XLSX.utils.sheet_to_json(book.Sheets["链接盈亏"]);
  assert.equal(links.find(x=>x["商品ID"]==="900000000001")["同组其他商品ID"], "900000000002");
  assert.equal(links.find(x=>x["商品ID"]==="900000000002")["同组其他商品ID"], "900000000001");
  assert.equal(links.find(x=>x["商品ID"]==="900000000004")["同组其他商品ID"], undefined);
  assert.equal(links.find(x=>x["商品ID"]==="900000000001")["同组其他链接出单明细"], "900000000002（27单）");
  assert.equal(links.find(x=>x["商品ID"]==="900000000002")["同组其他链接出单明细"], "900000000001（40单）");
  assert.equal(links.find(x=>x["商品ID"]==="900000000002")["同组保留单"], 67);
  assert.equal(Object.hasOwn(template, "出单率（单/链接）"), false);
  assert.ok(!book.SheetNames.some(name=>name.includes("出单率")));
});
