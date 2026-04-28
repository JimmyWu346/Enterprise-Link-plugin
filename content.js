(() => {
  const API_BASE = "https://zsk-cmis.cicpa.org.cn/open/enterprise_info_api/v3";
  const DEFAULT_SETTINGS = {
    timeoutSeconds: 60,
    investmentLevel: 6,
    shareholderLevel: 3,
    investmentConcurrency: 6,
    requestTimeoutMs: 10000
  };

  let isRunning = false;

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (!message || !message.type) {
      return false;
    }

    if (message.type === "GET_COOKIE") {
      sendResponse({ ok: true, cookie: document.cookie || "" });
      return false;
    }

    if (message.type === "EXTRACT_EXCEL") {
      if (isRunning) {
        sendResponse({ ok: false, error: "当前页面已有导出任务在执行。" });
        return false;
      }

      isRunning = true;
      runExport(message.settings)
        .then((result) => sendResponse({ ok: true, ...result }))
        .catch((error) => sendResponse({ ok: false, error: error.message }))
        .finally(() => {
          isRunning = false;
        });
      return true;
    }

    return false;
  });

  async function runExport(rawSettings) {
    const settings = normalizeSettings(rawSettings);
    const orgid = readUrlParam("orgid");
    if (!orgid) {
      throw new Error("地址栏未找到 orgid 参数。");
    }

    const collector = createCollector(settings);
    await crawlRoot({ orgid, collector });
    const exportDebug = downloadWorkbook(collector);
    // eslint-disable-next-line no-console
    console.log("[export-debug]", exportDebug);
    return {
      timedOut: collector.timeoutReached,
      personnelCount: collector.personnelRows.length,
      relationCount: collector.investmentRows.length,
      shareholderSheetCount: collector.shareholderSheets.length,
      exportDebug
    };
  }

  function createCollector(settings) {
    return {
      rootCompanyName: "",
      personnelRows: [],
      investmentRows: [],
      sourceCompanies: [],
      shareholderSheets: [],
      settings,
      stopAt: Date.now() + settings.timeoutSeconds * 1000,
      timeoutReached: false
    };
  }

  async function crawlRoot(params) {
    const { orgid, collector } = params;
    if (shouldStop(collector)) {
      return;
    }
    const rootData = await fetchEnterprisePlate(orgid, "", collector);
    const rootName = rootData.orgname || orgid;
    collector.rootCompanyName = rootName;
    collector.sourceCompanies.push({ id: orgid, name: rootName, level: 1 });
    collectPersonnel(rootData, collector, rootName);

    await crawlInvestmentNode({
      orgid,
      companyName: rootName,
      level: 1,
      collector,
      seenCompanyNames: new Set([rootName])
    });

    await buildShareholderSheets(collector);
  }

  async function crawlInvestmentNode(params) {
    const { orgid, companyName, level, collector, seenCompanyNames } = params;
    if (shouldStop(collector) || level >= collector.settings.investmentLevel) {
      return;
    }
    const data = await fetchEnterprisePlate(orgid, "", collector);
    const nextLevel = level + 1;
    if (nextLevel > collector.settings.investmentLevel) {
      return;
    }
    await crawlMainInvestmentBranch({
      data,
      companyName,
      nextLevel,
      collector,
      seenCompanyNames
    });
  }

  async function crawlMainInvestmentBranch(params) {
    const { data, companyName, nextLevel, collector, seenCompanyNames } = params;
    if (shouldStop(collector)) {
      return;
    }
    const investments = findChildrenByNodeNameFromData(data, "对外投资");
    const taskFactories = [];
    for (const company of investments) {
      if (shouldStop(collector)) {
        break;
      }
      if (String(company.type) !== "1") {
        continue;
      }
      if (seenCompanyNames.has(company.name)) {
        continue;
      }
      collector.investmentRows.push({
        level: toLevelText(nextLevel),
        parentName: companyName,
        ratio: normalizeRatio(company.info),
        subjectName: company.name
      });
      seenCompanyNames.add(company.name);
      collector.sourceCompanies.push({ id: company.id, name: company.name, level: nextLevel });
      taskFactories.push(async () =>
        safeCrawlInvestmentNode({
          orgid: company.id,
          companyName: company.name,
          level: nextLevel,
          collector,
          seenCompanyNames
        })
      );
    }
    if (taskFactories.length) {
      await runWithConcurrency(taskFactories, collector.settings.investmentConcurrency);
    }
  }

  async function safeCrawlInvestmentNode(params) {
    try {
      if (!shouldStop(params.collector)) {
        await crawlInvestmentNode(params);
      }
    } catch (error) {
      // ignore node failure
    }
  }

  async function buildShareholderSheets(collector) {
    if (shouldStop(collector)) {
      return;
    }
    const existingCompanyNames = new Set(collector.sourceCompanies.map((item) => item.name));
    const existingHolderNames = new Set();
    const existingAllNames = new Set(collector.sourceCompanies.map((item) => item.name));
    const pendingCompanies = dedupeCompaniesByName(collector.sourceCompanies).filter(
      (item) => item.level <= collector.settings.shareholderLevel
    );
    const pendingNameSet = new Set(pendingCompanies.map((item) => item.name));
    const processedSourceNames = new Set();

    while (pendingCompanies.length) {
      if (shouldStop(collector)) {
        break;
      }
      const company = pendingCompanies.shift();
      pendingNameSet.delete(company.name);
      if (processedSourceNames.has(company.name)) {
        continue;
      }
      processedSourceNames.add(company.name);

      const sheet = await buildSingleShareholderSheet({
        orgid: company.id,
        companyName: company.name,
        existingCompanyNames,
        existingHolderNames,
        existingAllNames,
        collector
      });
      if (!sheet) {
        continue;
      }

      collector.shareholderSheets.push(sheet);
      existingAllNames.add(sheet.holderName);
      for (const discovered of sheet.discoveredCompanies) {
        existingAllNames.add(discovered.name);
      }
      for (const discovered of sheet.discoveredCompanies) {
        if (discovered.level > collector.settings.shareholderLevel) {
          continue;
        }
        if (processedSourceNames.has(discovered.name) || pendingNameSet.has(discovered.name)) {
          continue;
        }
        pendingCompanies.push(discovered);
        pendingNameSet.add(discovered.name);
      }
    }
  }

  async function buildSingleShareholderSheet(params) {
    const { orgid, companyName, existingCompanyNames, existingHolderNames, existingAllNames, collector } = params;
    if (shouldStop(collector)) {
      return null;
    }
    const data = await safeFetchEnterprisePlate(orgid, "left", collector);
    if (!data) {
      return null;
    }
    const shareholders = findChildrenByNodeNameFromData(data, "股东");
    const maxResult = pickSingleMaxByPercent(shareholders);
    if (!maxResult) {
      return null;
    }

    const holder = maxResult.item;
    const holderRatio = normalizeRatio(holder.info);
    if (
      existingCompanyNames.has(holder.name) ||
      existingHolderNames.has(holder.name) ||
      existingAllNames.has(holder.name)
    ) {
      return null;
    }
    existingHolderNames.add(holder.name);

    const context = {
      rows: [["一级", "", "", holder.name]],
      seenCompanyNames: new Set([companyName]),
      discoveredCompanies: new Map(),
      collector
    };

    if (String(holder.type) === "1") {
      context.seenCompanyNames.add(holder.name);
      recordDiscoveredCompany(context, holder.id, holder.name, 1);
      await appendCompanyInvestmentRows({
        orgid: holder.id,
        companyName: holder.name,
        level: 2,
        context
      });
    } else if (String(holder.type) === "2") {
      await appendPersonHoldingRows({
        personId: holder.id,
        parentName: holder.name,
        level: 2,
        fallbackRatio: holderRatio,
        context
      });
    } else {
      return null;
    }

    return {
      sheetName: `${companyName} - （股东${holderRatio}）${holder.name}`,
      holderName: holder.name,
      headers: ["层级", "上级公司/上级自然人", "持股比例", "公司/自然人"],
      rows: context.rows,
      discoveredCompanies: [...context.discoveredCompanies.values()]
    };
  }

  async function appendCompanyInvestmentRows(params) {
    const { orgid, companyName, level, context } = params;
    if (shouldStop(context.collector) || level > context.collector.settings.shareholderLevel) {
      return;
    }
    const data = await safeFetchEnterprisePlate(orgid, "left", context.collector);
    if (!data) {
      return;
    }
    const investments = findChildrenByNodeNameFromData(data, "对外投资");
    const taskFactories = [];
    for (const company of investments) {
      if (shouldStop(context.collector)) {
        break;
      }
      if (String(company.type) !== "1" || context.seenCompanyNames.has(company.name)) {
        continue;
      }
      context.rows.push([toLevelText(level), companyName, normalizeRatio(company.info), company.name]);
      recordDiscoveredCompany(context, company.id, company.name, level);
      context.seenCompanyNames.add(company.name);
      if (level + 1 > context.collector.settings.shareholderLevel) {
        continue;
      }
      taskFactories.push(async () =>
        appendCompanyInvestmentRows({
          orgid: company.id,
          companyName: company.name,
          level: level + 1,
          context
        })
      );
    }
    if (taskFactories.length) {
      await runWithConcurrency(taskFactories, context.collector.settings.investmentConcurrency);
    }
  }

  async function appendPersonHoldingRows(params) {
    const { personId, parentName, level, fallbackRatio, context } = params;
    if (shouldStop(context.collector) || level > context.collector.settings.shareholderLevel) {
      return;
    }
    const personData = await safeFetchPersonExtend(personId, context.collector);
    if (!personData) {
      return;
    }
    const companies = findChildrenByNodeNameFromData(personData, "担任股东");
    for (const company of companies) {
      if (shouldStop(context.collector)) {
        break;
      }
      if (String(company.type) !== "1" || context.seenCompanyNames.has(company.name)) {
        continue;
      }
      context.rows.push([
        toLevelText(level),
        parentName,
        normalizeRatio(company.info) || fallbackRatio,
        company.name
      ]);
      recordDiscoveredCompany(context, company.id, company.name, level);
      context.seenCompanyNames.add(company.name);
      if (level + 1 > context.collector.settings.shareholderLevel) {
        continue;
      }
      await appendCompanyInvestmentRows({
        orgid: company.id,
        companyName: company.name,
        level: level + 1,
        context
      });
    }
  }

  async function safeFetchEnterprisePlate(orgid, position, collector) {
    try {
      return await fetchEnterprisePlate(orgid, position, collector);
    } catch (error) {
      return null;
    }
  }

  async function safeFetchPersonExtend(rwid, collector) {
    try {
      return await fetchPersonExtend(rwid, collector);
    } catch (error) {
      return null;
    }
  }

  function collectPersonnel(data, collector, companyName) {
    const rightNodes = Array.isArray(data.rightNode) ? data.rightNode : [];
    const personnelNode = rightNodes.find((item) => item && item.name === "主要人员");
    const personnel = personnelNode && Array.isArray(personnelNode.children) ? personnelNode.children : [];
    for (const person of personnel) {
      collector.personnelRows.push({
        companyName,
        personName: person.name || "",
        position: person.info || ""
      });
    }
  }

  function pickSingleMaxByPercent(items) {
    const validItems = items
      .map((item) => ({ item, percent: parsePercent(item.info) }))
      .filter((entry) => Number.isFinite(entry.percent));
    if (!validItems.length) {
      return null;
    }
    validItems.sort((a, b) => b.percent - a.percent);
    const top = validItems[0];
    const second = validItems[1];
    if (second && top.percent === second.percent) {
      return null;
    }
    return top;
  }

  function parsePercent(raw) {
    if (typeof raw !== "string") {
      return Number.NaN;
    }
    const match = raw.match(/(\d+(\.\d+)?)\s*%/);
    return match ? Number(match[1]) : Number.NaN;
  }

  function normalizeRatio(raw) {
    return typeof raw === "string" ? raw.trim() : "";
  }

  function normalizeSettings(rawSettings) {
    const settings = rawSettings || {};
    return {
      timeoutSeconds: toPositiveInt(settings.timeoutSeconds, DEFAULT_SETTINGS.timeoutSeconds),
      investmentLevel: toPositiveInt(settings.investmentLevel, DEFAULT_SETTINGS.investmentLevel),
      shareholderLevel: toPositiveInt(settings.shareholderLevel, DEFAULT_SETTINGS.shareholderLevel),
      investmentConcurrency: toPositiveInt(settings.investmentConcurrency, DEFAULT_SETTINGS.investmentConcurrency),
      requestTimeoutMs: toPositiveInt(settings.requestTimeoutMs, DEFAULT_SETTINGS.requestTimeoutMs)
    };
  }

  function toPositiveInt(value, fallback) {
    const num = Number(value);
    if (!Number.isFinite(num) || num <= 0) {
      return fallback;
    }
    return Math.floor(num);
  }

  function shouldStop(collector) {
    if (collector.timeoutReached) {
      return true;
    }
    if (Date.now() >= collector.stopAt) {
      collector.timeoutReached = true;
      return true;
    }
    return false;
  }

  function findChildrenByNodeName(nodes, nodeName) {
    const safeNodes = Array.isArray(nodes) ? nodes : [];
    const target = safeNodes.find((item) => item && item.name === nodeName);
    return target && Array.isArray(target.children) ? target.children : [];
  }

  function findChildrenByNodeNameFromData(data, nodeName) {
    const left = findChildrenByNodeName(data && data.leftNode, nodeName);
    return left.length ? left : findChildrenByNodeName(data && data.rightNode, nodeName);
  }

  function recordDiscoveredCompany(context, id, name, levelInSheet) {
    if (!context.discoveredCompanies.has(name)) {
      context.discoveredCompanies.set(name, { id, name, level: levelInSheet });
    }
  }

  function dedupeCompaniesByName(companies) {
    const map = new Map();
    for (const item of companies) {
      if (!map.has(item.name)) {
        map.set(item.name, item);
        continue;
      }
      const existing = map.get(item.name);
      if ((item.level || 999) < (existing.level || 999)) {
        map.set(item.name, item);
      }
    }
    return [...map.values()];
  }

  async function fetchEnterprisePlate(orgid, position, collector) {
    const params = new URLSearchParams({
      orgid,
      v: "v1",
      position: position || ""
    });
    const url = `${API_BASE}/enterprise_plates?${params.toString()}`;
    const body = await fetchJsonWithRetry(url, 3, collector.settings.requestTimeoutMs);
    if (!body || body.status_code !== 0 || !body.data) {
      throw new Error(`企业接口返回异常：${body && body.status_msg ? body.status_msg : "未知错误"}`);
    }
    return body.data;
  }

  async function fetchPersonExtend(rwid, collector) {
    const params = new URLSearchParams({
      rwid,
      v: "v1",
      position: "left"
    });
    const url = `${API_BASE}/enterprise_extends?${params.toString()}`;
    const body = await fetchJsonWithRetry(url, 3, collector.settings.requestTimeoutMs);
    if (!body || body.status_code !== 0 || !body.data) {
      throw new Error(`股东接口返回异常：${body && body.status_msg ? body.status_msg : "未知错误"}`);
    }
    return body.data;
  }

  async function fetchJsonWithRetry(url, maxAttempts, timeoutMs) {
    let lastError = null;
    for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), timeoutMs);
      try {
        const response = await fetch(url, {
          method: "GET",
          credentials: "include",
          signal: controller.signal
        });
        clearTimeout(timer);
        if (!response.ok) {
          throw new Error(`接口HTTP状态异常：${response.status}`);
        }
        return await response.json();
      } catch (error) {
        clearTimeout(timer);
        lastError = error;
        if (attempt < maxAttempts) {
          await sleep(attempt * 300);
        }
      }
    }
    throw lastError || new Error("请求失败");
  }

  function sleep(ms) {
    return new Promise((resolve) => {
      setTimeout(resolve, ms);
    });
  }

  async function runWithConcurrency(taskFactories, limit) {
    const safeLimit = Math.max(1, limit || 1);
    let cursor = 0;
    async function worker() {
      while (true) {
        const current = cursor;
        cursor += 1;
        if (current >= taskFactories.length) {
          return;
        }
        await taskFactories[current]();
      }
    }
    const workerCount = Math.min(safeLimit, taskFactories.length);
    const workers = Array.from({ length: workerCount }, () => worker());
    await Promise.allSettled(workers);
  }

  function toLevelText(level) {
    const map = {
      1: "一级",
      2: "二级",
      3: "三级",
      4: "四级",
      5: "五级",
      6: "六级",
      7: "七级",
      8: "八级",
      9: "九级",
      10: "十级"
    };
    return map[level] || `第${level}级`;
  }

  function readUrlParam(key) {
    const searchParams = new URLSearchParams(window.location.search || "");
    const fromSearch = searchParams.get(key);
    if (fromSearch) {
      return fromSearch;
    }
    const hash = window.location.hash || "";
    const queryIndex = hash.indexOf("?");
    if (queryIndex >= 0) {
      const hashQuery = hash.slice(queryIndex + 1);
      const hashParams = new URLSearchParams(hashQuery);
      const fromHash = hashParams.get(key);
      if (fromHash) {
        return fromHash;
      }
    }
    return "";
  }

  function downloadWorkbook(collector) {
    const personnelHeaders = ["姓名", "职务"];
    const personnelRows = collector.personnelRows.map((row) => [row.personName, row.position]);

    const relationHeaders = ["层级", "上级公司/上级自然人", "持股比例", "公司/自然人"];
    const relationRows = [
      ["一级", "", "", collector.rootCompanyName],
      ...sortRowsByLevel(collector.investmentRows).map((row) => [row.level, row.parentName, row.ratio, row.subjectName])
    ];

    const rawSheets = [
      {
        sheetName: "一级公司人员信息",
        headers: personnelHeaders,
        rows: personnelRows
      },
      {
        sheetName: "一级公司对外投资链路",
        headers: relationHeaders,
        rows: relationRows
      },
      ...collector.shareholderSheets.map((sheet) => ({
        sheetName: sheet.sheetName,
        headers: sheet.headers,
        rows: sortSheetRowsByLevel(sheet.rows)
      }))
    ];

    const normalizedNames = makeUniqueSheetNames(rawSheets.map((sheet) => normalizeSheetName(sheet.sheetName)));
    const duplicateNames = findDuplicateNames(rawSheets.map((sheet) => normalizeSheetName(sheet.sheetName)));
    const finalSheets = rawSheets.map((sheet, index) => ({
      ...sheet,
      sheetName: normalizedNames[index]
    }));

    const workbookXml = buildWorkbookXml(finalSheets);
    const blob = new Blob([workbookXml], {
      type: "application/vnd.ms-excel;charset=utf-8"
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const safeName = (collector.rootCompanyName || "企业关系导出").replace(/[\\/:*?"<>|]/g, "_");
    link.href = url;
    link.download = `${safeName}_关系链路导出.xls`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    return {
      rawSheetCount: rawSheets.length,
      normalizedSheetCount: finalSheets.length,
      normalizedNames,
      duplicateNames
    };
  }

  function buildWorkbookXml(sheets) {
    const worksheetXml = sheets.map((sheet) => buildWorksheetXml(sheet)).join("");
    return [
      "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
      "<?mso-application progid=\"Excel.Sheet\"?>",
      "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"",
      " xmlns:o=\"urn:schemas-microsoft-com:office:office\"",
      " xmlns:x=\"urn:schemas-microsoft-com:office:excel\"",
      " xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"",
      " xmlns:html=\"http://www.w3.org/TR/REC-html40\">",
      worksheetXml,
      "</Workbook>"
    ].join("");
  }

  function buildWorksheetXml(sheet) {
    const rows = [sheet.headers, ...sheet.rows];
    const rowXml = rows
      .map((row) => {
        const cellXml = row
          .map((cell) => `<Cell><Data ss:Type="String">${escapeXml(cell == null ? "" : String(cell))}</Data></Cell>`)
          .join("");
        return `<Row>${cellXml}</Row>`;
      })
      .join("");
    return `<Worksheet ss:Name="${escapeXml(sheet.sheetName)}"><Table>${rowXml}</Table></Worksheet>`;
  }

  function sortRowsByLevel(rows) {
    return [...rows].sort((a, b) => toLevelNumber(a.level) - toLevelNumber(b.level));
  }

  function sortSheetRowsByLevel(rows) {
    const firstLevelRows = rows.filter((row) => row[0] === "一级");
    const otherRows = rows.filter((row) => row[0] !== "一级");
    otherRows.sort((a, b) => toLevelNumber(a[0]) - toLevelNumber(b[0]));
    return [...firstLevelRows, ...otherRows];
  }

  function toLevelNumber(levelText) {
    const map = {
      一级: 1,
      二级: 2,
      三级: 3,
      四级: 4,
      五级: 5,
      六级: 6,
      七级: 7,
      八级: 8,
      九级: 9,
      十级: 10
    };
    return map[levelText] || 999;
  }

  function normalizeSheetName(name) {
    const safe = name.replace(/[\\/?*:[\]]/g, "_");
    return safe.length > 31 ? safe.slice(0, 31) : safe;
  }

  function makeUniqueSheetNames(names) {
    const used = new Map();
    const result = [];
    for (const name of names) {
      const base = name || "Sheet";
      const count = used.get(base) || 0;
      if (count === 0) {
        result.push(base);
        used.set(base, 1);
        continue;
      }
      let next = count;
      let candidate = buildSheetNameWithSuffix(base, next);
      while (used.has(candidate)) {
        next += 1;
        candidate = buildSheetNameWithSuffix(base, next);
      }
      result.push(candidate);
      used.set(base, next + 1);
      used.set(candidate, 1);
    }
    return result;
  }

  function buildSheetNameWithSuffix(base, index) {
    const suffix = `_${index}`;
    const maxBaseLength = Math.max(1, 31 - suffix.length);
    const trimmedBase = base.slice(0, maxBaseLength);
    return `${trimmedBase}${suffix}`;
  }

  function findDuplicateNames(names) {
    const countMap = new Map();
    for (const name of names) {
      countMap.set(name, (countMap.get(name) || 0) + 1);
    }
    return [...countMap.entries()]
      .filter((entry) => entry[1] > 1)
      .map((entry) => entry[0]);
  }

  function escapeXml(value) {
    return value
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }
})();
