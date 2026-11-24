(() => {
  "use strict";

  const { clean, isoDay, selectedText } = PddAnalysis;
  const limitBytes = 120 * 1024 * 1024;
  const maxExpanded = 280 * 1024 * 1024;
  const fields = {
    order: {
      order: ["订单号", "订单编号", "子订单号"],
      id: ["商品id", "商品编号", "商品链接id", "goodsid"],
      state: ["订单状态", "状态"],
      service: ["售后状态", "售后进度"],
      date: ["订单成交时间", "成交时间", "拼单成功时间", "支付时间", "订单支付时间", "成交日期"],
      receipt: ["商家实收金额", "商家实收", "订单实收金额"],
      quantity: ["商品数量", "商品件数", "购买数量", "件数", "数量"],
      title: ["商品", "商品名称", "商品标题"],
      code: ["商家编码-规格维度", "商家编码(规格维度)", "商品规格编码", "规格维度", "SKU编码", "规格编码"],
      variant: ["样式ID", "规格ID", "sku_id"],
      spec: ["商品规格", "规格名称", "SKU规格"],
      buyer: ["用户实付金额", "买家实付金额", "用户实付", "买家实付"],
      warehouse: ["发货仓库", "订单仓库", "仓库", "发货仓"],
    },
    promo: {
      id: ["商品ID", "商品编号", "商品链接ID"],
      title: ["商品名称", "商品标题", "商品"],
      spend: ["总花费", "推广总花费", "推广费", "推广花费", "实际消耗", "消耗"],
      orders: ["净成交笔数", "净成交订单数", "净成交订单"],
      revenue: ["净交易额", "净成交金额"],
      date: ["推广日期", "日期", "统计日期", "数据日期"],
    },
    listing: {
      id: ["商品ID", "商品编号", "商品链接ID"],
      template: ["模板标识", "模板编号", "模板号", "模板ID", "模板名称"],
      active: ["上架状态", "商品状态", "是否上架"],
      title: ["商品名称", "商品标题", "商品"],
    },
    cost: {
      code: ["规格编码", "SKU编码", "产品编码", "商品编码", "商家编码-规格维度", "商家编码", "货品编码", "SKU"],
      price: ["单位成本", "单件成本", "商品成本", "成本价", "成本", "采购价", "结算价"],
      warehouse: ["仓库", "所属仓库", "成本仓", "发货仓"],
    },
  };
  const required = { order: ["order","id","state","date","receipt","quantity","code"], promo: ["id","spend"], cost: ["code","price"], listing: ["id","template"] };
  const types = { order: "订单对账报表", promo: "商品推广报表", cost: "商品成本表", listing: "同模板上架清单" };

  const heading = value => selectedText(value).replace(/\((元|人民币|件|个|单|笔|%|￥)\)/g, "").replace(/[-_()/:：]/g, "");
  const blank = value => ["", "-", "--", "—", "暂无", "/"].includes(clean(value));
  const describe = (name, row) => name + " 第 " + row + " 行";
  function amount(value, label, fallback = null) {
    if (blank(value)) return fallback;
    const text = clean(value).replace(/[,￥¥\s]/g, "");
    if (!/^\+?\d+(?:\.\d+)?$/.test(text) || !Number.isFinite(Number(text))) throw new Error(label + "包含无效或负数数值。");
    return Number(text);
  }
  function dateCell(value) {
    if (typeof value !== "number") return isoDay(value);
    const d = XLSX.SSF.parse_date_code(value);
    return d ? isoDay(d.y + "-" + d.m + "-" + d.d) : "";
  }
  function formattedIdentifier(sheet, row, col, value) {
    if (typeof value !== "number") return clean(value);
    if (!Number.isSafeInteger(value)) return "";
    const cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
    return /^\d+$/.test(cell?.w || "") ? cell.w : String(value);
  }
  function findHeader(sheet, kind) {
    if (!sheet["!ref"]) return null;
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    if (range.e.r - range.s.r > 300000 || range.e.c - range.s.c > 250) {
      throw new Error("表格使用范围过大，请删除空白区域或把报表按日期拆分后再导入。");
    }
    for (let row = range.s.r; row <= Math.min(range.e.r, range.s.r + 18); row++) {
      const header = new Map();
      for (let col = range.s.c; col <= range.e.c; col++) {
        const label = heading(sheet[XLSX.utils.encode_cell({ r: row, c: col })]?.v);
        if (label && !header.has(label)) header.set(label, col);
      }
      const columns = Object.fromEntries(Object.entries(fields[kind]).map(([key, aliases]) => [
        key, aliases.map(heading).map(a => header.get(a)).find(col => col !== undefined),
      ]));
      if (required[kind].every(key => columns[key] !== undefined)) return { row, last: range.e.r, columns };
    }
    return null;
  }
  function decodeCsv(bytes) {
    if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder("utf-16le").decode(bytes);
    if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder("utf-16be").decode(bytes);
    try { return new TextDecoder("utf-8", { fatal: true }).decode(bytes); }
    catch { return new TextDecoder("gb18030", { fatal: true }).decode(bytes); }
  }
  function checkZip(bytes) {
    if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) return;
    let expanded = 0, count = 0;
    fflate.unzipSync(bytes, { filter(entry) {
      expanded += entry.originalSize;
      if (++count > 500 || expanded > maxExpanded) throw new Error("压缩文件解包后的体积过大，请按时间分批导出后分析。");
      return false;
    }});
  }
  function readBook(bytes, name) {
    checkZip(bytes);
    const readOptions = { raw: true, nodim: true, cellNF: true, cellText: true };
    if (/\.(?:csv|tsv)$/i.test(name)) {
      return XLSX.read(decodeCsv(bytes), { ...readOptions, type: "string", FS: name.toLowerCase().endsWith(".tsv") ? "\t" : "," });
    }
    return XLSX.read(bytes, { ...readOptions, type: "array" });
  }
  async function entries(file, kind) {
    if (!file?.size || file.size > limitBytes) throw new Error(types[kind] + "为空或超过 120 MB。可缩小报表区间后重试。");
    const bytes = new Uint8Array(await file.arrayBuffer());
    if (/\.zip$/i.test(file.name)) {
      if (kind !== "order") throw new Error("请先解压并选择" + types[kind] + "中的表格。");
      checkZip(bytes);
      const files = fflate.unzipSync(bytes, { filter: item =>
        !item.name.includes("__MACOSX/") && !item.name.split("/").at(-1).startsWith("~$") && /\.(?:csv|tsv|xlsx|xls)$/i.test(item.name),
      });
      if (!Object.keys(files).length) throw new Error("订单 ZIP 中没有 CSV、XLSX 或 XLS 表格。");
      return Object.entries(files).map(([name, data]) => ({ name, data }));
    }
    if (!/\.(?:csv|tsv|xlsx|xls|json)$/i.test(file.name)) throw new Error("暂不支持文件“" + file.name + "”的格式。");
    return [{ name: file.name, data: bytes }];
  }
  function loadRows(book, name, kind) {
    const results = [];
    for (const sheetName of book.SheetNames) {
      const sheet = book.Sheets[sheetName];
      const header = findHeader(sheet, kind);
      if (!header) continue;
      const { row, last, columns } = header;
      const get = (r, name, isId) => {
        const col = columns[name];
        if (col === undefined) return "";
        const v = sheet[XLSX.utils.encode_cell({ r, c: col })]?.v;
        return isId ? formattedIdentifier(sheet, r, col, v) : v ?? "";
      };
      const rows = [];
      for (let r = row + 1; r <= last; r++) {
        if (!required[kind].some(key => !blank(get(r, key)))) continue;
        rows.push({ n: r + 1, get: (name, isId) => get(r, name, isId) });
      }
      results.push({ name, sheetName, rows, dated: columns.date !== undefined, hasStatus: columns.active !== undefined });
    }
    return results;
  }
  function codeList(kind) {
    const important = { order: "订单号、商品 ID、订单状态、成交时间、商家实收、商品数量和规格编码",
      promo: "商品 ID、总花费（或推广总花费）", cost: "规格编码、单位成本", listing: "商品 ID、模板标识" };
    return "需要包含" + important[kind] + "。";
  }
  async function tables(files, kind) {
    let all = [];
    for (const file of files) {
      let found = 0;
      for (const { data, name } of await entries(file, kind)) {
        if (kind === "cost" && /\.json$/i.test(name)) {
          const obj = JSON.parse(decodeCsv(data));
          if (!(Array.isArray(obj?.p) || Array.isArray(obj?.y))) throw new Error("成本 JSON 需要包含 p（莆田）或 y（义乌）数组。");
          all.push({ json: obj, name });
          found++;
        } else {
          const matched = loadRows(readBook(data, name), name, kind);
          found += matched.length;
          all = all.concat(matched);
        }
      }
      if (!found) throw new Error("文件“" + file.name + "”未识别到" + types[kind] + "表头。" + codeList(kind));
    }
    return all;
  }
  function orderRows(tables, allowEmpty) {
    const rows = [];
    for (const { name, rows: lines } of tables) {
      for (const { n, get } of lines) {
        const order = get("order", true), id = get("id", true);
        if (["合计","总计"].includes(order)) continue;
        const state = clean(get("state")), service = clean(get("service"));
        const date = dateCell(get("date"));
        if (!order || !/^\d{6,}$/.test(id) || !state) {
          throw new Error(describe(name, n) + "订单号、商品 ID 或订单状态格式无法识别。");
        }
        const closed = /待付款|未付款|未支付|待支付|已取消|交易取消|交易关闭|已关闭/.test(state);
        if (!date && !closed) throw new Error(describe(name, n) + "缺有效成交时间，请导出含成交时间的对账报表。");
        const refunded = /退款成功|全额退款|退款完成/.test(state + " " + service);
        const label = describe(name, n);
        rows.push({ order, id, state, service, date, code: get("code", true), variant: get("variant", true),
          title: clean(get("title")), spec: clean(get("spec")), warehouse: clean(get("warehouse")),
          receipt: closed || refunded ? null : amount(get("receipt"), label + "商家实收"),
          quantity: closed || refunded ? null : amount(get("quantity"), label + "数量"),
          buyer: closed || refunded ? null : amount(get("buyer"), label + "买家实付"),
        });
      }
    }
    if (!rows.length && !allowEmpty) throw new Error("订单文件没有明细行，请检查下载的报表。");
    return rows;
  }
  function promoRows(tables) {
    if (tables.length !== 1) throw new Error("推广文件包含多张商品数据表。请导出单一「推广商品」汇总表或分天商品表后再导入。");
    const rows = [];
    const { name, sheetName, rows: lines, dated } = tables[0];
    let total = null;
    for (const { n, get } of lines) {
      const id = get("id", true);
      if (/^(总计|合计|汇总)$/.test(id)) {
        const one = amount(get("spend"), describe(name, n) + "推广总花费");
        if (total !== null && one !== total) throw new Error("推广报表含多个不同的总计金额，请导入单一商品汇总报表。");
        total = one;
        continue;
      }
      if (/^(注[:：]?|备注[:：]?|仅下载)/.test(id) || (!id && blank(get("spend")))) continue;
      if (!/^\d{6,}$/.test(id)) throw new Error(describe(name, n) + "商品 ID 无法识别，需导出到商品维度。");
      const date = dateCell(get("date"));
      if (dated && !date) throw new Error(describe(name, n) + "缺有效推广日期。");
      const label = describe(name, n);
      rows.push({ id, date, title: clean(get("title")), spend: amount(get("spend"), label + "推广总花费"),
        orders: amount(get("orders"), label + "推广净成交笔数", 0), revenue: amount(get("revenue"), label + "推广净交易额", 0) });
    }
    return { rows, total, namedPeriod: spanFromName(name) ?? spanFromName(sheetName), dated };
  }
  function costRows(tables) {
    const rows = [];
    for (const table of tables) {
      const put = (code, value, warehouse, where) => {
        if (!code && blank(value)) return;
        if (/^(总计|合计|汇总)$/.test(code)) return;
        const price = amount(value, where + "单位成本");
        if (!code || price === null) throw new Error(where + "缺少规格编码或单位成本。");
        rows.push({ code, price, warehouse });
      };
      if (table.json) {
        for (const [key, warehouse] of [["p", "莆田"], ["y", "义乌"]]) {
          for (const item of table.json[key] || []) {
            if (!Array.isArray(item) || item.length < 2) throw new Error("成本 JSON 的每一行需要 [规格编码, 单位成本]。");
            put(clean(item[0]), item[1], warehouse, table.name);
          }
        }
      } else {
        const inferredWarehouse = /义乌|莆田/.exec(table.sheetName)?.[0] || "";
        for (const line of table.rows) put(line.get("code", true), line.get("price"),
          clean(line.get("warehouse")) || inferredWarehouse, describe(table.name, line.n));
      }
    }
    if (tables.length && !rows.length) throw new Error("成本表没有有效的编码和成本行。可下载成本表模板核对。");
    return rows;
  }

  function listingRows(tables) {
    if (tables.length > 1) throw new Error("上架清单包含多张商品表，请整理成一张完整清单。");
    const byId = new Map();
    for (const { name, rows, hasStatus } of tables) {
      for (const { n, get } of rows) {
        const id = get("id", true);
        if (/^(总计|合计|汇总)$/.test(id)) continue;
        const template = get("template", true);
        const status = clean(get("active"));
        if (!/^\d{6,}$/.test(id) || !template) throw new Error(describe(name, n) + "需要有效的商品 ID 和模板标识。");
        const active = !hasStatus || /^(已上架|上架|在售|出售中|销售中|1)$/.test(status);
        if (!active && !/^(已下架|下架|停售|已停售|未上架|草稿|已删除|0)$/.test(status)) {
          throw new Error(describe(name, n) + "上架状态无法识别，请填写「在售」或「已下架」。仅包含上架商品时可移除该列。");
        }
        const previous = byId.get(id);
        if (previous && (previous.template !== template || previous.active !== active)) {
          throw new Error(describe(name, n) + "的商品 ID 对应了冲突的模板或上架状态。请核对清单。");
        }
        byId.set(id, { id, template, active, title: clean(get("title")) || previous?.title || "" });
      }
    }
    if (tables.length && !byId.size) throw new Error("上架清单没有有效的商品与模板记录。");
    return byId;
  }

  function spanFromName(name) {
    const found = [...name.matchAll(/(?<!\d)(20\d{2})[-_./年]?(0[1-9]|1[0-2])[-_./月]?(0[1-9]|[12]\d|3[01])(?!\d)/g)]
      .map(m => isoDay(m[1] + "-" + m[2] + "-" + m[3]));
    return found.length === 2 && found[0] && found[1] && found[0] <= found[1] ? found : null;
  }
  function extents(rows) {
    let min = "", max = "";
    for (const { date } of rows) {
      if (!date) continue;
      if (!min || date < min) min = date;
      if (date > max) max = date;
    }
    return [min, max];
  }
  async function importFiles(files, params, progress = () => {}) {
    progress("正在读取订单报表…");
    const orders = orderRows(await tables(files.order, "order"), Boolean(files.listing?.length));
    progress("正在读取商品推广报表…");
    const promo = promoRows(await tables(files.promo, "promo"));
    progress("正在读取成本表…");
    const costs = costRows(await tables(files.cost, "cost"));
    progress("正在读取同模板上架清单…");
    const listing = listingRows(await tables(files.listing || [], "listing"));
    const observed = extents(orders);
    const promoDays = promo.dated ? extents(promo.rows) : ["", ""];
    const inferred = promo.namedPeriod ?? (promoDays[0] ? promoDays : observed);
    const start = params.start || inferred[0], end = params.end || inferred[1];
    if (!start || !end) throw new Error("无法从报表识别分析时期。请手动填写开始和结束日期。");
    if (promo.namedPeriod && (promo.dated ? start < promo.namedPeriod[0] || end > promo.namedPeriod[1] :
      start !== promo.namedPeriod[0] || end !== promo.namedPeriod[1])) {
      throw new Error("分析日期与推广报表日期（" + promo.namedPeriod.join(" 至 ") + "）不一致。请重新选择同期报表或调整分析区间。");
    }
    const sliced = promo.dated && (start > (promo.namedPeriod?.[0] || promoDays[0]) || end < (promo.namedPeriod?.[1] || promoDays[1]));
    const promos = promo.dated ? promo.rows.filter(row => row.date >= start && row.date <= end) : promo.rows;
    const info = {
      orderNames: files.order.map(file => file.name), promoName: files.promo[0].name, costNames: files.cost.map(file => file.name),
      promoTotal: sliced ? null : promo.total, observed, namedPeriod: promo.namedPeriod, datedPromo: promo.dated, listingName: files.listing?.[0]?.name || "",
      promoRows: promos.length, costRows: costs.length, hasCosts: costs.length > 0,
    };
    progress("正在按商品、规格和成交价计算…");
    return PddAnalysis.analyze({ orders, promos, costs, listing, start, end, options: params.options, info });
  }

  globalThis.PddReportFiles = Object.freeze({ importFiles });
})();
