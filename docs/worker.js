"use strict";

importScripts("./vendor/xlsx.full.min.js", "./vendor/fflate.js", "./analysis.js?v=1.4.0", "./reports.js?v=1.4.0");

let result = null;
const snapshotUrl = "https://wuhengqwe.github.io/profit-analysis-tool/cost_data.json";
function makeDemo() {
  const orderHeaders = ["订单号", "商品id", "订单状态", "售后状态", "订单成交时间", "商家实收金额(元)", "商品数量(件)",
    "商家编码-规格维度", "样式ID", "商品名称", "商品规格", "用户实付金额(元)"];
  const rows = [orderHeaders];
  let no = 0;
  const product = (id, title, tiers) => {
    for (const [code, price, count, qty] of tiers) {
      for (let i = 0; i < count; i++) {
        const d = PddAnalysis.shiftDay("2026-10-31", -(i % 36));
        rows.push(["DEMO-" + (++no), id, "已收货", "无售后或售后取消", d + " 13:12:05",
          price * qty, qty, code, code, title, qty === 1 ? "标准规格" : "订单购买 " + qty + " 件", price * qty]);
      }
    }
  };
  product("900000000001", "演示款 A｜常规版", [["00101", 18.9, 24, 1], ["00101X2", 29.9, 16, 1]]);
  product("900000000002", "演示款 A｜同价链接", [["00101", 18.9, 18, 1], ["00101X2", 29.9, 9, 2]]);
  product("900000000003", "演示款 A｜其他价位", [["00101", 16.9, 13, 1], ["00101X2", 29.9, 7, 1]]);
  product("900000000004", "演示款 B｜高出单", [["00102", 13.9, 62, 1]]);
  product("900000000005", "演示款 C｜缺成本", [["00103", 21.9, 8, 1]]);
  product("900000000007", "演示款 D｜低运费才盈利", [["00104", 15, 14, 1]]);
  for (let i = 0; i < 4; i++) rows.push(["DEMO-" + (++no), "900000000001", "已发货，退款成功", "退款成功",
    "2026-10-12 10:00:00", 18.9, 1, "00101", "00101", "演示款 A｜常规版", "标准规格", 18.9]);
  rows.push(["DEMO-" + (++no), "900000000004", "待付款", "无售后或售后取消", "", 0, 1,
    "00102", "00102", "演示款 B｜高出单", "标准规格", 0]);
  const csv = rows.map(row => row.map(v => '"' + String(v).replaceAll('"', '""') + '"').join(",")).join("\r\n");
  const zip = fflate.zipSync({ "订单对账报表.csv": new TextEncoder().encode("\uFEFF" + csv) });
  const promos = [
    ["商品ID", "商品名称", "总花费(元)", "净成交笔数", "净交易额(元)"],
    ["900000000001", "演示款 A｜常规版", 199.5, 23, 524.7],
    ["900000000002", "演示款 A｜同价链接", 175.2, 17, 490.3],
    ["900000000003", "演示款 A｜其他价位", 141, 12, 221.3],
    ["900000000004", "演示款 B｜高出单", 420, 52, 752.9],
    ["900000000005", "演示款 C｜缺成本", 52, 6, 131.4],
    ["900000000006", "演示款 E｜有推广无单", 38.8, 0, 0],
    ["900000000007", "演示款 D｜低运费才盈利", 57, 10, 150],
    ["900000000008", "演示款 A｜常规版", 0, 0, 0],
    ["总计", "-", 1083.5, 120, 2270.6],
  ];
  const costs = [
    ["规格编码", "单位成本", "仓库"], ["00101", 4.8, "义乌"], ["00102", 4.6, "义乌"], ["00104", 6.1, "义乌"],
  ];
  function excel(rows, name) {
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), "报表");
    return new File([XLSX.write(book, { bookType: "xlsx", type: "array" })], name);
  }
  return {
    order: [new File([zip], "演示_订单对账报表.zip")],
    promo: [excel(promos, "演示_推广商品_20260803至20261031.xlsx")],
    cost: [excel(costs, "演示_商品成本表.xlsx")],
  };
}
function workbookFor(r) {
  const workbook = XLSX.utils.book_new();
  const text = v => v == null ? "" : v;
  const p = n => n == null ? "" : n;
  const clip = v => typeof v === "string" ? v.slice(0, 32000) : v;
  const fmt = (val, type) => val === "" || val == null ? undefined : type === "percent" ? { t: "n", v: val, z: "0.0%" } :
    type === "currency" ? { t: "n", v: val, z: "#,##0.00;[Red]-#,##0.00" } : clip(val);
  function add(name, headers, rows, moneyCols = [], pctCols = [], widths = []) {
    const sheet = XLSX.utils.aoa_to_sheet([headers, ...rows.map(row => row.map((v, i) =>
      fmt(text(v), pctCols.includes(i) ? "percent" : moneyCols.includes(i) ? "currency" : "plain")))]);
    sheet["!cols"] = headers.map((_, i) => ({ wch: widths[i] || 16 }));
    sheet["!autofilter"] = { ref: XLSX.utils.encode_range({ r: 0, c: 0 }, { r: rows.length, c: headers.length - 1 }) };
    XLSX.utils.book_append_sheet(workbook, sheet, name);
  }
  const s = r.stats, o = r.opt;
  add("总览与口径", ["项目", "值", "说明"], [
    ["分析起始日", r.start], ["分析结束日", r.end], ["原始订单数", s.raw], ["保留订单数", s.retained,
      "按订单号去重，排除未付、取消、退款成功和区间外记录"],
    ["保留商品件数", s.units], ["近30天保留订单数", s.recent], ["退款成功订单数", s.refunded],
    ["售后处理中（保留）", s.pendingAfterSales], ["有估算利润的链接", s.known], ["利润待补链接", s.pending],
    ["商家实收（元）", s.revenue], ["商品推广花费（元）", s.spend], ["推广总计与明细差额", p(s.promoGap)],
    ["基础运费下可估算链接利润（元）", p(s.base)], ["对照运费下可估算链接利润（元）", p(s.stress)],
    ["优先复发模板组", s.priority], ["同码同价多链接组", s.multi], ["无保留单推广花费（元）", s.noOrdersAdSpend],
    ["清单中上架链接", s.listed], ["可计算出单率的模板", s.rateTemplates],
    ["出单率（单/链接）", "模板保留订单去重后，除以推广报表中可归到同模板的去重商品 ID 数，含零单链接。可选的上架清单优先。"],
    ["推广报表商品 ID", s.promoIds], ["未归组推广 ID", s.promoUnresolved], ["名称估算归组的 ID", s.promoInferred],
    ["推广表未覆盖的出单 ID", s.promoMissingSold], ["推广费取成交花费作为补充的行", r.info.transactionSpendRows],
    ["模板最高单量链接", "每组按保留订单数选最高的一条商品链接；订单相同时优先近30天订单数较多的链接。"],
    ["同模板上架清单", r.info.listingName || "未导入"],
    ["莆田税点", o.putianTax + "%", "对莆田保留单实收计提"],
    ["义乌税点", o.yiwuTax + "%", "对义乌保留单实收计提"],
    ["莆田计税实收", s.putianReceipt], ["义乌计税实收", s.yiwuReceipt], ["计税仓待核对订单", s.taxPending],
    ["分仓计税", "优先按明确发货仓；缺发货仓按成本匹配。跨仓组合按仓库成本占比分摊实收。"], ["包装损耗率", o.packaging / 100, "按商品成本计提"],
    ["技术服务费率", o.service / 100, "按保留订单商家实收计提"], ["退款损耗预留率", o.reserve / 100, "按保留订单商家实收估提"],
    ["基础运费（元/单）", o.shipping], ["对照运费（元/单）", o.stressShipping],
    ["成本匹配顺序", "先匹配莆田，未匹配时查义乌；包含套数和组合规格"],
    ["成本快照时间", r.info.snapshotDate || "自有成本表或演示数据"],
    ["估算公式", "实收 − 商品成本 − 包装损耗 − 推广费用 − 税费 − 技术服务费 − 运费 − 退款损耗预留"],
    ["模板归并", "按上架清单模板标识归并，未关联到清单时按全部成交规格编码与常见实收单价归并。"],
    ["推荐口径", "对照运费下利润为正、至少两条链接各有五单且盈利，组累计≥20单，近期≥10单且占比≥12%，各链接主价比例≥50%。"],
    ["利润边界", "以本次导入成本表为准。缺成本、推广花费或计税仓的链接估算利润留空。"],
    ["文件边界", "请确保各表属于同一店铺。订单以成交日筛选，订单总数不用于跨商品 ID 的简单相加。"],
    ["去重与运费", "同订单内的多个有效商品链接各计一单，整店只计一单，运费在链接间平摊。"],
    ["订单文件", r.info.orderNames.join(" / ")], ["推广文件", r.info.promoName], ["成本文件", r.info.costNames.join(" / ") || "未导入"],
  ], [], [], [34, 76, 75]);
  add("模板清单", ["模板号", "发布建议", "商品标题", "代表商品ID", "同组链接", "保留单", "近30天单", "对照盈利链接",
    "基础利润", "对照利润", "基础利润率", "商家实收", "推广花费", "规格与常见实收单价", "组内最低主价占比", "清单模板标识", "出单率分母链接数", "出单率（单/链接）", "出单率来源", "名称推断链接数", "未见推广表的出单链接数"],
  r.templates.map(x => [x.id, x.action, x.title, x.rep, x.linked, x.orders, x.recent, x.solid, x.base, x.stress,
    x.margin, x.revenue, x.spend, x.signatureText, x.share, x.listingTemplate, x.listedLinks, x.orderRate, x.rateSource, x.estimatedLinks, x.promoMissing]), [8,9,11,12,17], [10,14], [12,18,48,20,12,12,14,18,15,15,14,16,16,74,20,22,18,20]);
  add("模板最高单量链接", ["模板号", "最高单量商品ID", "商品标题", "保留单", "近30天单", "商品件数", "退款单",
    "商家实收", "推广花费", "估算税费", "基础利润", "对照利润", "基础利润率", "主销规格编码", "主销实收单价",
    "匹配成本仓", "上架状态", "利润状态"],
  r.leaders.map(x => [x.template, x.id, x.title, x.orders, x.recent, x.units, x.refunded, x.revenue, x.spend,
    x.tax, x.base, x.stress, x.margin, x.primaryCode, x.price, x.costWarehouse,
    x.listed === true ? "在售" : x.listed === false ? "已下架" : "", x.status]),
  [7,8,9,10,11,14], [12], [12,20,48,12,14,12,12,16,16,15,15,15,14,20,15]);
  add("链接盈亏", ["排名", "商品ID", "商品标题", "模板号", "保留单", "近30天单", "商品件数", "已付单", "退款单", "售后处理中",
    "商家实收", "推广花费", "匹配商品成本", "基础利润", "对照利润", "基础利润率", "主销规格编码", "主销实收单价",
    "缺成本单", "状态", "推广净成交笔数", "推广净交易额", "运费计费单数", "无规格编码行", "发布建议", "匹配成本仓", "莆田计税实收", "义乌计税实收", "估算税费", "计税仓待核对单", "上架状态"],
  r.links.map(x => [x.rank, x.id, x.title, x.template, x.orders, x.recent, x.units, x.paid, x.refunded, x.pending, x.revenue,
    x.spend, x.partialCost, x.base, x.stress, x.margin, x.primaryCode, x.price, x.missing, x.status, x.adOrders, x.adRevenue,
    x.freightUnits, x.noCode, x.action, x.costWarehouse, x.putianReceipt, x.yiwuReceipt, x.tax, x.taxMissing, x.listed === true ? "在售" : x.listed === false ? "已下架" : ""]), [10,11,12,13,14,17,21,26,27,28], [15], [9,20,48,12]);
  add("规格成交证据", ["模板号", "商品ID", "商品标题", "规格编码", "规格描述", "保留订单", "件数", "常见实收单价",
    "常见买家实付单价", "最低实收单价", "最高实收单价", "主价行占比", "缺成本单", "匹配成本仓"],
  r.specs.map(x => [x.template, x.id, x.title, x.code, x.description, x.orders, x.units, x.price, x.buyer, x.lower, x.upper, x.share, x.missing, x.costWarehouse]),
  [7,8,9,10], [11], [12,20,48,22,45]);
  add("重叠规格待核对", ["规格编码", "常见实收单价", "同价链接数", "涉及模板组", "规格保留单", "商品ID示例", "模板号示例"],
    r.overlaps.map(x => [x.code, x.price, x.links, x.groups, x.orders, x.ids.slice(0, 25).join(" / "), x.templates.slice(0, 25).join(" / ")]),
    [1], [], [22,18,16,16,16,72,72]);
  add("出单率链接核对", ["商品 ID", "商品名称", "推广名称", "模板号", "归组依据", "计入分母", "数据来源"],
    r.rateDetails.map(x => [x.id, x.title, x.promotionName, x.template, x.evidence, x.counted ? "是" : "否", x.source]),
    [], [], [20,54,51,14,27,14,18]);
  add("缺成本编码", ["规格编码", "涉及保留订单"], r.missingCosts.map(x => [x.code, x.orders]), [], [], [27,20]);
  add("每日订单", ["成交日", "保留订单"], r.days, [], [], [18,18]);
  return XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true });
}

self.addEventListener("message", async ({ data }) => {
  try {
    if (data.kind === "export") {
      if (!result) throw new Error("先完成分析，再导出结果。");
      const binary = workbookFor(result);
      self.postMessage({ kind: "download", binary, filename: "拼多多出单分析_" + result.start + "_" + result.end + ".xlsx" }, [binary]);
      return;
    }
    result = null;
    const demo = data.kind === "demo";
    const files = demo ? makeDemo() : data.files;
    const params = demo ? { start: "", end: "", options: PddAnalysis.DEFAULTS } : data.params;
    const progress = text => self.postMessage({ kind: "progress", text });
    let snapshotDate = "";
    if (!demo && params.costSource === "public") {
      progress("正在加载莆田与义乌的成本快照…");
      let snapshot;
      try {
        const response = await fetch(snapshotUrl, { cache: "no-store", credentials: "omit", signal: AbortSignal.timeout(14000) });
        if (!response.ok) throw new Error();
        snapshot = await response.json();
        if (!Array.isArray(snapshot.p) || !Array.isArray(snapshot.y) || !snapshot.p.length || !snapshot.y.length) throw new Error();
      } catch {
        throw new Error("现有成本快照读取失败。请重试，或选择导入自己的成本表。");
      }
      snapshotDate = snapshot._sync?.checkedAt || "";
      files.cost = [new File([JSON.stringify(snapshot)], "公开成本快照（莆田＋义乌）.json")];
    }
    result = await PddReportFiles.importFiles(files, params, progress);
    result.info.snapshotDate = snapshotDate;
    self.postMessage({ kind: "result", result, demo });
  } catch (error) {
    if (data.kind !== "export") result = null;
    self.postMessage({ kind: "error", text: error?.message || "读取报表失败。请检查文件格式后重试。" });
  }
});
