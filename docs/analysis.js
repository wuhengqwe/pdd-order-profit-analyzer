(() => {
  "use strict";

  const DEFAULTS = Object.freeze({ tax: 5, packaging: 5, service: 0.6, reserve: 5, shipping: 2.5, stressShipping: 3 });
  const ACTIONS = ["优先复发", "小批复测", "运费敏感", "样本待积累", "待补成本或规格", "暂缓复发"];
  const num = new Intl.Collator("zh-CN", { numeric: true });
  const clean = value => String(value ?? "").normalize("NFKC").trim();
  const codeKey = value => clean(value).toUpperCase();
  const selectedText = value => codeKey(value).replace(/\s+/g, "");
  const cents = amount => Math.sign(amount) * Math.round(Number((Math.abs(amount) * 100).toFixed(7)));
  const money = amount => amount === null ? null : amount / 100;
  const add = (m, key, value = 1) => m.set(key, (m.get(key) || 0) + value);
  const mode = m => [...m].sort(([a, ac], [b, bc]) => bc - ac || (typeof a === "number" ? a - b : num.compare(a, b)))[0]?.[0] ?? null;
  const shiftDay = (date, days) => {
    const d = new Date(date + "T00:00:00Z");
    d.setUTCDate(d.getUTCDate() + days);
    return d.toISOString().slice(0, 10);
  };
  function isoDay(value) {
    const m = clean(value).match(/^(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})(?=\D|$)/);
    if (!m) return "";
    const date = m[1] + "-" + m[2].padStart(2, "0") + "-" + m[3].padStart(2, "0");
    const d = new Date(date + "T00:00:00Z");
    return Number.isFinite(d.getTime()) && d.toISOString().slice(0, 10) === date ? date : "";
  }
  function optionsFrom(input) {
    const result = { ...DEFAULTS, ...input };
    for (const key of ["tax", "packaging", "service", "reserve", "shipping", "stressShipping"]) {
      const n = result[key];
      if (typeof n !== "number" || !Number.isFinite(n) || n < 0 || n > (key.toLowerCase().includes("shipping") ? 1000 : 100)) {
        throw new Error("费用参数超出范围，请核对税率、损耗比例和运费。");
      }
    }
    if (result.tax + result.service + result.reserve > 100) throw new Error("税率、技术服务费率和退款损耗率合计不能超过 100%。");
    if (result.stressShipping < result.shipping) throw new Error("对照运费不能低于基础运费。");
    return result;
  }

  function costMatcher(costs) {
    const indexes = { 莆田: new Map(), 义乌: new Map(), 其他: new Map() };
    const conflict = Symbol("conflicting price");
    const unitSuffixes = ["*1盒", "*1包", "*1袋", "*1片", "*1支", "*1瓶", "*1条", "*1箱", "*1"];
    function insert(map, key, entry) {
      const previous = map.get(key);
      if (previous === undefined) map.set(key, entry);
      else if (previous === conflict || previous.price !== entry.price) map.set(key, conflict);
    }
    const numericYiwu = new Map();
    for (const cost of costs) {
      const code = codeKey(cost.code);
      const tag = clean(cost.warehouse);
      const warehouse = tag.includes("莆田") ? "莆田" : tag.includes("义乌") ? "义乌" : "其他";
      const entry = { price: cost.price, warehouse: tag || "自有成本表" };
      insert(indexes[warehouse], code, entry);
      if (warehouse === "义乌" && /^\d+$/.test(code)) {
        insert(numericYiwu, code.replace(/^0+(?=\d)/, ""), entry);
      }
    }
    function lookup(key, warehouse) {
      return indexes[warehouse].get(key) ?? (warehouse === "义乌" && /^\d+$/.test(key) ?
        numericYiwu.get(key.replace(/^0+(?=\d)/, "")) : undefined);
    }
    function combine(parts) {
      if (parts.includes(conflict)) return conflict;
      if (!parts.every(Boolean)) return undefined;
      return {
        price: parts.reduce((sum, item) => sum + item.price, 0),
        warehouse: [...new Set(parts.map(item => item.warehouse))].join(" + "),
      };
    }
    function singleWarehouse(key, warehouse) {
      if (!key) return undefined;
      const hit = lookup(key, warehouse);
      if (hit !== undefined) return hit;
      if (key.includes("+")) return combine(key.split("+").map(part => singleWarehouse(part.trim(), warehouse)));
      const pack = key.match(/^(.*?)[*X×](\d+)[盒包袋片支瓶条箱]?$/);
      const base = (pack ? pack[1] : key).trim();
      const qty = pack ? Number(pack[2]) : 1;
      if (!base || !Number.isSafeInteger(qty) || qty < 1 || qty > 100000) return undefined;
      for (const candidate of new Set([base, ...unitSuffixes.map(s => base + s)])) {
        const item = lookup(candidate, warehouse);
        if (item === conflict) return conflict;
        if (item !== undefined) return { ...item, price: item.price * qty };
      }
      return undefined;
    }
    const cache = new Map();
    function find(code) {
      if (!cache.has(code)) {
        const result = singleWarehouse(code, "莆田") ?? singleWarehouse(code, "义乌") ?? singleWarehouse(code, "其他") ??
          (code.includes("+") ? combine(code.split("+").map(part => find(part.trim()))) : undefined);
        cache.set(code, result);
      }
      return cache.get(code);
    }
    return code => {
      const match = find(code);
      return match === conflict ? null : match;
    };
  }

  function newLink(id, title) {
    return { id, title, paid: new Set(), refunded: new Set(), kept: new Set(), recent: new Set(), pending: new Set(),
      units: 0, revenue: 0, cost: 0, spend: 0, spendMissing: false, adOrders: 0, adRevenue: 0, noCode: 0, warehouses: new Set(),
      missing: new Set(), freightUnits: 0, skus: new Map() };
  }

  function resultForLink(a, opt) {
    const ratio = (opt.tax + opt.service + opt.reserve) / 100;
    const netBeforeShipping = a.revenue * (1 - ratio) - a.cost * (1 + opt.packaging / 100) - a.spend;
    const known = !a.missing.size && !a.spendMissing;
    const base = known ? money(cents(netBeforeShipping - a.freightUnits * opt.shipping)) : null;
    const stress = known ? money(cents(netBeforeShipping - a.freightUnits * opt.stressShipping)) : null;
    const specs = [...a.skus].map(([code, s]) => {
      const p = mode(s.receipt);
      return {
        id: a.id, code, description: mode(s.descriptions) || "", price: money(p), buyer: money(mode(s.buyer)),
        lower: money([...s.receipt.keys()].reduce((x,y)=>Math.min(x,y))), upper: money([...s.receipt.keys()].reduce((x,y)=>Math.max(x,y))),
        orders: s.orders.size, units: s.units, costWarehouse: mode(s.sources) || "", atPrice: s.receipt.get(p), lineCount: s.lines,
        share: s.receipt.get(p) / s.lines, missing: s.missing.size,
      };
    }).sort((left, right) => right.orders - left.orders || num.compare(left.code, right.code));
    const share = specs.length ? specs.reduce((s, x) => s + x.atPrice, 0) / (a.noCode + specs.reduce((s, x) => s + x.lineCount, 0)) : 0;
    const status = [a.missing.size ? "缺少成本" : "", a.spendMissing ? "缺少推广花费" : ""].filter(Boolean).join("、") ||
      (a.kept.size ? "可估算" : "无保留订单");
    return { id: a.id, title: a.title || a.id, orders: a.kept.size, recent: a.recent.size, paid: a.paid.size,
      refunded: [...a.refunded].filter(x => !a.kept.has(x)).length, pending: a.pending.size, units: a.units,
      revenue: money(cents(a.revenue)), spend: money(cents(a.spend)), cost: a.missing.size ? null : money(cents(a.cost)),
      partialCost: money(cents(a.cost)), base, stress, margin: base !== null && a.revenue ? base / a.revenue : null, adOrders: a.adOrders,
      adRevenue: money(cents(a.adRevenue)), freightUnits: a.freightUnits, missing: a.missing.size, spendMissing: a.spendMissing,
      noCode: a.noCode, costWarehouse: [...a.warehouses].sort().join(" / "), primaryCode: specs[0]?.code ?? "", price: specs[0]?.price ?? null, share, status, specs };
  }

  function templatesFor(links) {
    const groups = new Map();
    for (const row of links) {
      if (!row.orders) continue;
      const signature = row.noCode ? [] : row.specs.map(s => [s.code, s.price]).sort(([a], [b]) => num.compare(a, b) || (a < b ? -1 : a > b ? 1 : 0));
      const key = signature.length ? JSON.stringify(signature) : "missing:" + row.id;
      if (!groups.has(key)) groups.set(key, { signature, members: [] });
      groups.get(key).members.push(row);
    }
    const templates = [...groups.values()].map(({ signature, members }) => {
      const total = field => members.reduce((sum, link) => sum + link[field], 0);
      const orders = total("orders");
      const recent = total("recent");
      const revenue = money(cents(total("revenue")));
      const spend = money(cents(total("spend")));
      const complete = members.every(link => link.base !== null);
      const base = complete ? money(cents(total("base"))) : null;
      const stress = complete ? money(cents(total("stress"))) : null;
      const solid = members.filter(x => x.stress > 0 && x.orders >= 5).length;
      const share = members.reduce((min, x) => Math.min(min, x.share), 1);
      const rep = [...members].sort((a, b) => Number(b.stress > 0) - Number(a.stress > 0) || b.recent - a.recent ||
        b.orders - a.orders || (b.stress || 0) - (a.stress || 0) || num.compare(a.id, b.id))[0];
      let action = "样本待积累";
      if (!complete || !signature.length) action = "待补成本或规格";
      else if (base <= 0) action = "暂缓复发";
      else if (stress <= 0) action = "运费敏感";
      else if (solid >= 2 && orders >= 20 && recent >= Math.max(10, orders * .12) && share >= .5) action = "优先复发";
      else if (rep.orders >= 10 && recent >= 3 && share >= .5) action = "小批复测";
      return { action, rep: rep.id, title: rep.title, linked: members.length, solid, orders, recent, base, stress,
        margin: base !== null && revenue ? base / revenue : null, revenue, spend, share, signature,
        signatureText: signature.map(([code, p]) => code + " ¥" + p.toFixed(2)).join(" / "), members };
    });
    templates.sort((a, b) => ACTIONS.indexOf(a.action) - ACTIONS.indexOf(b.action) || (b.stress || 0) - (a.stress || 0) ||
      b.recent - a.recent || b.orders - a.orders || num.compare(a.rep, b.rep));
    templates.forEach((row, i) => {
      row.id = "T" + String(i + 1).padStart(4, "0");
      for (const link of row.members) {
        link.template = row.id;
        link.action = row.action;
      }
      row.memberIds = row.members.map(link => link.id);
      delete row.members;
    });
    return templates;
  }

  function analyze({ orders, promos, costs, options, start, end, info = {} }) {
    const opt = optionsFrom(options);
    if (!isoDay(start) || !isoDay(end) || start > end) throw new Error("请核对分析开始日期与结束日期。");
    const recentStart = shiftDay(end, -29);
    const map = new Map();
    const retainedGoods = new Map();
    const paid = new Set();
    const refunded = new Set();
    const presale = new Set();
    const raw = new Set();
    const uniqueLines = new Map();
    const days = new Map();
    const missingCodes = new Map();
    const matcher = costMatcher(costs);
    let repeated = 0, omitted = 0, missingSpendRows = 0;
    function link(id, title) {
      if (!map.has(id)) map.set(id, newLink(id, title));
      const result = map.get(id);
      if (!result.title && title) result.title = title;
      return result;
    }
    for (const row of orders) {
      raw.add(row.order);
      const code = codeKey(row.code);
      const key = JSON.stringify([row.order, row.id, row.variant || code || row.spec]);
      const snapshot = JSON.stringify([row.date, row.state, row.service, row.quantity, row.receipt, row.buyer, code,
        row.warehouse, row.spec]);
      if (uniqueLines.has(key)) {
        if (uniqueLines.get(key) !== snapshot) throw new Error("订单报表存在同订单、同商品、同规格的冲突行。请重新导出同一时点且日期不重叠的报表。");
        repeated++;
        continue;
      }
      uniqueLines.set(key, snapshot);
      const isRefunded = /退款成功|全额退款|退款完成/.test(row.state + " " + row.service);
      const isUnpaid = /待付款|未付款|未支付|待支付/.test(row.state);
      const isCanceled = /已取消|交易取消|交易关闭|已关闭/.test(row.state);
      if ((!isRefunded && (isUnpaid || isCanceled)) || !row.date || row.date < start || row.date > end) {
        omitted++;
        continue;
      }
      const a = link(row.id, row.title);
      a.paid.add(row.order);
      paid.add(row.order);
      if (isRefunded) {
        a.refunded.add(row.order);
        refunded.add(row.order);
        continue;
      }
      if (!Number.isSafeInteger(row.quantity) || row.quantity < 1 || row.receipt === null) {
        throw new Error("有效订单存在商品件数或商家实收缺失。请查看订单对账报表中的数量和金额列。");
      }
      const goods = retainedGoods.get(row.order) || new Set();
      if (!goods.size) add(days, row.date);
      goods.add(row.id);
      retainedGoods.set(row.order, goods);
      a.kept.add(row.order);
      if (row.date >= recentStart) a.recent.add(row.order);
      if (/售后处理中|退款中|退款处理中/.test(row.service + " " + row.state)) {
        a.pending.add(row.order);
        presale.add(row.order);
      }
      a.units += row.quantity;
      a.revenue += row.receipt;
      const hit = matcher(code);
      if (hit) { a.cost += hit.price * row.quantity; a.warehouses.add(hit.warehouse); }
      else {
        a.missing.add(row.order);
        if (!missingCodes.has(code)) missingCodes.set(code, new Set());
        missingCodes.get(code).add(row.order);
      }
      if (!code) {
        a.noCode++;
        continue;
      }
      if (!a.skus.has(code)) a.skus.set(code, {
        receipt: new Map(), buyer: new Map(), descriptions: new Map(), orders: new Set(), missing: new Set(), sources: new Map(), units: 0, lines: 0,
      });
      const s = a.skus.get(code);
      add(s.receipt, cents(row.receipt / row.quantity));
      if (row.buyer !== null) add(s.buyer, cents(row.buyer / row.quantity));
      if (row.spec) add(s.descriptions, row.spec);
      s.lines++;
      s.orders.add(row.order);
      s.units += row.quantity;
      if (hit) add(s.sources, hit.warehouse);
      else s.missing.add(row.order);
    }
    for (const goods of retainedGoods.values()) {
      for (const id of goods) map.get(id).freightUnits += 1 / goods.size;
    }
    for (const promo of promos) {
      if (!map.has(promo.id) && !promo.spend && promo.spend !== null && !promo.orders) continue;
      const a = link(promo.id, promo.title);
      a.spend += promo.spend ?? 0;
      a.adOrders += promo.orders || 0;
      a.adRevenue += promo.revenue || 0;
      if (promo.spend === null) {
        a.spendMissing = true;
        missingSpendRows++;
      }
    }
    const links = [...map.values()].map(a => resultForLink(a, opt));
    links.sort((a, b) => b.orders - a.orders || b.recent - a.recent || b.revenue - a.revenue || b.spend - a.spend || num.compare(a.id, b.id));
    links.forEach((r, i) => { r.rank = i + 1; });
    const templates = templatesFor(links);
    const specs = [];
    const similar = new Map();
    for (const a of links) {
      for (const s of a.specs) {
        const row = { ...s, template: a.template, title: a.title };
        specs.push(row);
        const key = JSON.stringify([s.code, s.price]);
        if (!similar.has(key)) similar.set(key, []);
        similar.get(key).push(row);
      }
      delete a.specs;
    }
    const overlaps = [...similar.values()].flatMap(rows => {
      const groups = [...new Set(rows.map(x => x.template))];
      if (groups.length < 2) return [];
      return [{ code: rows[0].code, price: rows[0].price, links: rows.length, groups: groups.length,
        orders: rows.reduce((sum, s) => sum + s.orders, 0), ids: rows.map(s => s.id), templates: groups }];
    }).sort((a, b) => b.orders - a.orders || num.compare(a.code, b.code));
    const total = (rows, field) => money(cents(rows.reduce((sum, x) => sum + (x[field] ?? 0), 0)));
    const known = links.filter(x => x.base !== null);
    const pending = links.length - known.length;
    const spend = total(links, "spend");
    const promoGap = info.promoTotal == null ? null : money(cents(info.promoTotal - spend));
    const stats = {
      raw: raw.size, retained: retainedGoods.size, paid: paid.size, refunded: [...refunded].filter(x => !retainedGoods.has(x)).length,
      pendingAfterSales: presale.size, units: links.reduce((sum, x) => sum + x.units, 0), recent: [...days].filter(([d]) => d >= recentStart)
        .reduce((sum, [, count]) => sum + count, 0), links: links.length,
      revenue: total(links, "revenue"), spend, base: known.length ? total(known, "base") : null,
      stress: known.length ? total(known, "stress") : null, known: known.length, pending, repeated, omitted,
      noOrdersAdSpend: total(links.filter(x => !x.orders), "spend"), promoGap, missingSpendRows,
      priority: templates.filter(x => x.action === "优先复发").length, multi: templates.filter(x => x.linked > 1).length,
    };
    if (!links.length) throw new Error("所选日期内没有可分析的成交或商品推广明细。请检查分析区间和报表内容。");
    return { start, end, info, opt, stats, templates, links, specs, overlaps, missingCosts: [...missingCodes]
      .map(([code, ids]) => ({ code: code || "（空编码）", orders: ids.size })).sort((a, b) => b.orders - a.orders || num.compare(a.code, b.code)),
      days: [...days].sort(([a], [b]) => num.compare(a, b)) };
  }
  globalThis.PddAnalysis = Object.freeze({ analyze, DEFAULTS, ACTIONS, isoDay, clean, selectedText, cents, shiftDay });
})();
