(() => {
  "use strict";

  const $ = id => document.getElementById(id);
  const normalize = value => String(value ?? "").normalize("NFKC").trim().toLocaleLowerCase("zh-CN");
  const fmtCount = new Intl.NumberFormat("zh-CN");
  const fmtRate = new Intl.NumberFormat("zh-CN", { maximumFractionDigits: 2, minimumFractionDigits: 2 });
  const fmtMoney = new Intl.NumberFormat("zh-CN", { style: "currency", currency: "CNY" });
  const compare = new Intl.Collator("zh-CN", { numeric: true, sensitivity: "base" });
  const pageSize = 25;
  const options = ["putianTax", "yiwuTax", "shipping", "stressShipping", "packaging", "service", "reserve"];
  const labels = { templates: "模板清单", leaders: "模板最高单量链接", links: "链接盈亏", specs: "规格与价格", overlaps: "重叠规格", rateDetails: "出单率核对" };
  const headers = {
    templates: [["id","模板号"],["title","代表商品"],["action","发布建议"],["linked","出单链接", "count"],["listedLinks","分母链接数","count"],["orders","保留单","count"],["orderRate","出单率（单/链接）","rate"],
      ["recent","近 30 天","count"],["base","基础利润","money"],["stress","对照利润","money"],["spend","推广费","money"]],
    leaders: [["template","模板号"],["title","最高单量链接"],["orders","保留单","count"],["recent","近 30 天","count"],
      ["units","商品件数","count"],["refunded","退款单","count"],["revenue","商家实收","money"],["spend","推广费","money"],
      ["tax","估算税费","money"],["base","基础利润","money"],["stress","对照利润","money"],["costWarehouse","匹配成本仓"]],
    links: [["rank","出单排名","count"],["title","商品链接"],["template","模板号"],["orders","保留单","count"],["units","商品件数","count"],
      ["recent","近 30 天","count"],["base","基础利润","money"],["stress","对照利润","money"],["spend","推广费","money"],
      ["costWarehouse","匹配成本仓"],["tax","估算税费","money"],["listed","上架状态"]],
    specs: [["code","规格编码"],["title","商品链接"],["template","模板号"],["orders","保留单","count"],["units","商品件数","count"],
      ["price","最常见实收单价","money"],["lower","最低实收单价","money"],["upper","最高实收单价","money"],["costWarehouse","匹配成本仓"]],
    overlaps: [["code","规格编码"],["price","最常见实收单价","money"],["links","涉及链接","count"],["groups","模板组","count"],
      ["orders","规格保留单","count"],["templates","模板组编号"]],
    rateDetails: [["id","商品 ID"],["title","商品"],["template","模板号"],["evidence","归组依据"],
      ["counted","计入分母"],["source","数据来源"]],
  };

  const state = { result: null, view: "templates", renderedView: "", sort: {}, page: 0, worker: null, waiting: false };
  const displayMoney = value => value == null ? "待补" : fmtMoney.format(value);
  function el(tag, text, className) {
    const node = document.createElement(tag);
    if (text != null) node.textContent = text;
    if (className) node.className = className;
    return node;
  }
  function signed(node, value) {
    if (value > 0) node.classList.add("gain");
    if (value < 0) node.classList.add("loss");
  }
  function status(text) { $("status").textContent = text; }
  function busy(waiting) {
    state.waiting = waiting;
    $("run").disabled = $("demo").disabled = waiting;
    $("download").disabled = waiting || !state.result;
    $("setup").setAttribute("aria-busy", String(waiting));
  }
  function clearError() {
    $("error").hidden = true;
    $("error").textContent = "";
    $("setup").querySelectorAll('[aria-invalid="true"]').forEach(x => x.removeAttribute("aria-invalid"));
  }
  function showError(text, focusNode) {
    $("error").hidden = false;
    $("error").textContent = text;
    if (focusNode) {
      focusNode.setAttribute("aria-invalid", "true");
      focusNode.focus();
    } else $("error").focus();
    status("请核对报表或费用设置后重试。");
  }
  function stopWorker() {
    state.worker?.terminate();
    state.worker = null;
  }
  function invalidate() {
    stopWorker();
    state.result = null;
    $("results").hidden = true;
    $("rows").replaceChildren();
    busy(false);
    clearError();
    status("分析完成后可下载 Excel 明细。");
  }
  function fileLabel(input, target, placeholder) {
    const label = $(target);
    const files = [...$(input).files];
    label.textContent = files.length ? files.map(x => x.name).join("、") : placeholder;
    label.classList.toggle("loaded", files.length > 0);
  }
  function switchCost() {
    const mode = $("setup").querySelector('input[name="costSource"]:checked').value;
    $("custom-cost").hidden = mode !== "custom";
    $("cost-file").disabled = mode !== "custom";
    $("cost-file").required = mode === "custom";
    $("cost-help").textContent = {
      public: "按规格编码先查莆田，未匹配再查义乌。快照从原亏损分析网站读取。",
      custom: "按「仓库」列区分莆田和义乌。未标仓库的成本表在两仓都匹配不到时使用。",
      none: "按订单数和成交价归并模板。成交链接的利润会显示待补。",
    }[mode];
  }
  function start(message) {
    invalidate();
    try {
      const worker = new Worker("./worker.js?v=1.4.0");
      state.worker = worker;
      worker.addEventListener("message", ({ data }) => {
        if (worker !== state.worker) return;
        if (data.kind === "progress") {
          status(data.text);
        } else if (data.kind === "error") {
          busy(false);
          showError(data.text);
        } else if (data.kind === "result") {
          state.result = data.result;
          state.sort = { links: { field: "orders", direction: "descending" } };
          busy(false);
          showResults(data.demo);
          status(data.demo ? "示例分析完成。" : "分析完成，可导出 Excel。");
          $("results").scrollIntoView({ block: "start" });
          $("result-title").focus({ preventScroll: true });
        } else if (data.kind === "download") {
          busy(false);
          const url = URL.createObjectURL(new Blob([data.binary], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          }));
          const a = el("a");
          a.href = url;
          a.download = data.filename;
          document.body.append(a);
          a.click();
          a.remove();
          setTimeout(() => URL.revokeObjectURL(url), 30000);
          status("Excel 已生成。");
        }
      });
      worker.addEventListener("error", () => {
        if (worker !== state.worker) return;
        stopWorker();
        busy(false);
        showError("分析组件未能启动。请刷新网页并重新选择报表。");
      });
      busy(true);
      status("正在准备分析…");
      worker.postMessage(message);
    } catch {
      stopWorker();
      busy(false);
      showError("此浏览器无法启动分析组件，请使用当前版本的主流浏览器。");
    }
  }
  function metric(label, value, sub, amount) {
    const group = el("dl", null, "metric");
    const v = el("dd", value);
    if (amount !== undefined) signed(v, amount);
    group.append(el("dt", label), v, el("small", sub));
    return group;
  }
  function showResults(demo) {
    const r = state.result, s = r.stats, o = r.opt;
    $("results").hidden = false;
    $("result-title").textContent = demo ? "示例分析结果" : "分析结果";
    const snapshot = r.info.snapshotDate ? " · 成本快照 " + r.info.snapshotDate.slice(0,10) : "";
    $("result-scope").textContent = r.start + " 至 " + r.end + " · 莆田税点 " + o.putianTax + "% · 义乌税点 " + o.yiwuTax + "% · 运费 " +
      fmtMoney.format(o.shipping) + " / 对照 " + fmtMoney.format(o.stressShipping) + snapshot;
    $("metrics").replaceChildren(
      metric("保留订单", fmtCount.format(s.retained), "原始订单 " + fmtCount.format(s.raw) + " 单"),
      metric("保留商品件数", fmtCount.format(s.units), "近 30 天 " + fmtCount.format(s.recent) + " 单"),
      metric("商品推广费用", fmtMoney.format(s.spend), "商家实收 " + fmtMoney.format(s.revenue)),
      metric("可估算链接利润", displayMoney(s.base), "对照运费 " + displayMoney(s.stress), s.base),
      metric("优先复发模板", fmtCount.format(s.priority), "同码同价多链接组 " + fmtCount.format(s.multi)),
    );
    const warnings = [];
    if (!r.info.hasCosts) warnings.push("本次未导入成本表。成交链接的利润和盈利模板建议需补齐成本后再看。");
    else if (s.pending) warnings.push(s.pending + " 条链接缺成本、推广花费或计税仓，利润留空。");
    if (s.taxPending && r.info.hasCosts) warnings.push(s.taxPending + " 笔保留订单的计税仓待核对。自有成本表可在「仓库」列填写莆田或义乌。");
    if (s.promoGap !== null && Math.abs(s.promoGap) > .02) warnings.push("推广报表总计与商品明细相差 " + fmtMoney.format(s.promoGap) + "。差额未分摊到链接，请核对推广报表。");
    if (!r.info.namedPeriod && !r.info.datedPromo) warnings.push("推广文件未识别出日期，请核对其与分析区间和店铺一致。");
    if (!demo && r.info.observed[0] && (r.info.observed[0] > r.start || r.info.observed[1] < r.end)) {
      warnings.push("订单文件中有成交记录的日期为 " + r.info.observed.join(" 至 ") + "，请确认分析区间内的订单报表均已导入。");
    }
    if (r.info.transactionSpendRows) warnings.push(fmtCount.format(r.info.transactionSpendRows) + " 行推广明细缺总花费，已取成交花费。");
    if (s.promoUnresolved) warnings.push("推广报表中有 " + fmtCount.format(s.promoUnresolved) + " 个商品 ID 待确认为同模板，分母采用可归组的商品，可在「出单率核对」查看依据。");
    if (s.promoMissingSold && !r.info.listingName) warnings.push(fmtCount.format(s.promoMissingSold) + " 个出过单的商品 ID 未在推广表出现，相关模板的分母仅覆盖推广报表范围。");
    if (s.repeated) warnings.push("已合并 " + fmtCount.format(s.repeated) + " 行重复订单明细。");
    if (s.pendingAfterSales) warnings.push(s.pendingAfterSales + " 笔售后处理中订单暂计入保留单。");
    $("warnings").replaceChildren(...warnings.map(value => el("li", value)));
    $("warnings").hidden = !warnings.length;
    $("filter").replaceChildren(new Option("全部建议", ""),
      ...[...new Set(r.templates.map(x => x.action))].map(value => new Option(value, value)));
    setView(r.templates.length ? "templates" : "links");
  }

  function setView(view, query = "") {
    state.view = view;
    state.page = 0;
    $("query").value = query;
    $("filter").value = "";
    $("filter").hidden = view !== "templates" && view !== "links";
    $("views").querySelectorAll("button").forEach(btn => btn.setAttribute("aria-pressed", String(btn.dataset.view === view)));
    render();
  }
  function sortableHeader() {
    if (state.renderedView === state.view) return;
    state.renderedView = state.view;
    $("table-title").textContent = labels[state.view];
    const row = el("tr");
    for (const [field, name, type] of headers[state.view]) {
      const th = el("th", null, type ? "num" : "");
      th.scope = "col";
      if (type) {
        const button = el("button", name, "sort");
        button.type = "button";
        button.dataset.sort = field;
        const marker = el("span", "↕");
        marker.setAttribute("aria-hidden", "true");
        button.append(marker);
        th.append(button);
      } else th.textContent = name;
      row.append(th);
    }
    $("head").replaceChildren(row);
    $("rows").replaceChildren();
  }
  function rowKey(item) {
    if (state.view === "specs") return JSON.stringify([item.id, item.code]);
    if (state.view === "overlaps") return JSON.stringify([item.code, item.price]);
    if (state.view === "rateDetails") return JSON.stringify([item.id, item.source]);
    return item.id;
  }
  function rowElement(item) {
    const row = el("tr");
    row.dataset.key = rowKey(item);
    for (const [idx,[field,,type]] of headers[state.view].entries()) {
      const v = item[field], cell = el(idx === 0 ? "th" : "td", null, type ? "num" : "");
      if (idx === 0) cell.scope = "row";
      if (state.view === "templates" && field === "id") {
        const button = el("button", v, "row-action");
        button.type = "button";
        button.dataset.members = v;
        button.setAttribute("aria-label", "查看模板 " + v + " 的成员链接");
        cell.append(button);
      } else if (field === "title") {
        cell.append(el("span", v));
        if (state.view === "rateDetails") {
          if (item.promotionName) cell.append(el("small", item.promotionName));
        } else cell.append(el("small", "商品 ID " + (item.rep ?? item.id)));
        if (state.view === "templates") {
          if (item.listingTemplate) cell.append(el("small", "清单模板 " + item.listingTemplate));
          if (item.signatureText) cell.append(el("small", item.signatureText, "sku-prices"));
        }
      } else if (field === "action") {
        cell.append(el("span", v, "tag " + (v === "优先复发" ? "priority" : v === "暂缓复发" ? "loss" : "")));
      } else if (field === "templates") {
        cell.textContent = v.slice(0, 14).join(" / ");
      } else if (type === "money") {
        cell.textContent = displayMoney(v);
        if (field === "base" || field === "stress") signed(cell, v);
      } else if (type === "rate") {
        if (v == null) cell.textContent = "—";
        else {
          cell.append(el("span", fmtRate.format(v)));
          if (item.rateSource) cell.append(el("small", item.estimatedLinks ? "含 " + item.estimatedLinks + " 个推断 ID" : item.rateSource));
        }
      } else if (field === "counted") {
        cell.textContent = v ? "是" : "待核对或由清单覆盖";
      } else if (field === "listed") {
        cell.textContent = v === true ? "在售" : v === false ? "已下架" : "—";
      } else if (type === "count") {
        cell.textContent = v == null ? "—" : fmtCount.format(v);
      } else {
        cell.textContent = v || "—";
      }
      row.append(cell);
    }
    return row;
  }
  function render() {
    if (!state.result) return;
    sortableHeader();
    const query = normalize($("query").value);
    const filter = $("filter").value;
    const matching = state.result[state.view].filter(item => {
      if (filter && item.action !== filter) return false;
      if (!query) return true;
      return [item.id, item.rep, item.title, item.code, item.primaryCode, item.signatureText, item.template,
        item.listingTemplate, item.promotionName, item.evidence, item.ids?.join(" "), item.templates?.join(" ")].some(field => normalize(field).includes(query));
    });
    const sort = state.sort[state.view];
    if (sort) {
      if (state.view !== "links" || sort.field !== "orders" || sort.direction !== "descending") {
        matching.sort((a, b) => {
          const left = a[sort.field], right = b[sort.field];
          if (left == null) return right == null ? 0 : 1;
          if (right == null) return -1;
          return (typeof left === "number" ? left - right : compare.compare(left, right)) * (sort.direction === "ascending" ? 1 : -1);
        });
      }
    }
    for (const button of $("head").querySelectorAll("[data-sort]")) {
      const th = button.parentElement;
      if (sort?.field === button.dataset.sort) th.setAttribute("aria-sort", sort.direction);
      else th.removeAttribute("aria-sort");
      button.lastElementChild.textContent = sort?.field === button.dataset.sort ? (sort.direction === "ascending" ? "↑" : "↓") : "↕";
    }
    const maxPage = Math.max(0, Math.ceil(matching.length / pageSize) - 1);
    state.page = Math.min(maxPage, Math.max(state.page, 0));
    const page = matching.slice(state.page * pageSize, (state.page + 1) * pageSize);
    const keys = new Set(page.map(rowKey));
    const tbody = $("rows");
    const existing = new Map();
    for (const row of [...tbody.children]) {
      if (!keys.has(row.dataset.key)) row.remove();
      else existing.set(row.dataset.key, row);
    }
    for (let i = 0; i < page.length; i++) {
      const item = page[i];
      const row = existing.get(rowKey(item)) || rowElement(item);
      if (row === tbody.children[i]) continue;
      if (row.parentNode === tbody && typeof tbody.moveBefore === "function") tbody.moveBefore(row, tbody.children[i] || null);
      else tbody.insertBefore(row, tbody.children[i] || null);
    }
    if (!page.length) {
      const tr = el("tr"), td = el("td", "当前筛选条件下没有记录。", "empty");
      td.colSpan = headers[state.view].length;
      tr.append(td);
      tbody.append(tr);
    }
    $("result-count").textContent = fmtCount.format(matching.length) + " 条 · 第 " + (state.page + 1) + " / " + (maxPage + 1) + " 页";
    $("previous").disabled = state.page === 0;
    $("next").disabled = state.page >= maxPage;
  }

  $("setup").addEventListener("change", event => {
    invalidate();
    if (event.target.name === "costSource") switchCost();
    fileLabel("orders-file", "orders-name", "支持多个 CSV、Excel 或 ZIP");
    fileLabel("promo-file", "promo-name", "支持 XLS、XLSX、CSV");
    fileLabel("listing-file", "listing-name", "CSV 或 Excel · 商品 ID、模板标识、上架状态");
    fileLabel("cost-file", "cost-name", "可选择多份表格，按「仓库」列区分莆田和义乌");
  });
  $("setup").addEventListener("invalid", event => {
    event.target.setAttribute("aria-invalid", "true");
    if (event.target.closest("details")) event.target.closest("details").open = true;
  }, true);
  $("setup").addEventListener("submit", event => {
    event.preventDefault();
    if (state.waiting) return;
    const params = {
      start: $("start").value, end: $("end").value, options: Object.fromEntries(options.map(key => [key, $(key).valueAsNumber])),
      costSource: $("setup").querySelector('input[name="costSource"]:checked').value,
    };
    if (params.start && params.end && params.start > params.end) return showError("结束日期须等于或晚于开始日期。", $("end"));
    if (params.options.stressShipping < params.options.shipping) return showError("对照运费须等于或高于基础运费。", $("stressShipping"));
    start({ kind: "analyze", params, files: {
      order: [...$("orders-file").files], promo: [...$("promo-file").files], listing: [...$("listing-file").files],
      cost: params.costSource === "custom" ? [...$("cost-file").files] : [],
    } });
  });
  $("demo").addEventListener("click", () => start({ kind: "demo" }));
  $("download").addEventListener("click", () => {
    if (!state.result || !state.worker || state.waiting) return;
    busy(true);
    status("正在生成 Excel…");
    state.worker.postMessage({ kind: "export" });
  });
  $("views").addEventListener("click", event => {
    const view = event.target.closest("[data-view]")?.dataset.view;
    if (view) setView(view);
  });
  $("head").addEventListener("click", event => {
    const button = event.target.closest("[data-sort]");
    if (!button) return;
    const field = button.dataset.sort;
    const previous = state.sort[state.view];
    state.sort[state.view] = { field, direction: previous?.field === field && previous.direction === "ascending" ? "descending" : "ascending" };
    state.page = 0;
    render();
  });
  $("rows").addEventListener("click", event => {
    const template = event.target.closest("[data-members]")?.dataset.members;
    if (template) {
      setView("links", template);
      $("query").focus();
    }
  });
  for (const name of ["query", "filter"]) $(name).addEventListener(name === "query" ? "input" : "change", () => { state.page = 0; render(); });
  $("previous").addEventListener("click", () => { state.page--; render(); });
  $("next").addEventListener("click", () => { state.page++; render(); });
  document.querySelector('a[href="#how"]').addEventListener("click", () => { $("how").open = true; });
  switchCost();
})();
