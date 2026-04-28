"use strict";

const fs = require("fs/promises");
const https = require("https");
const path = require("path");

const API_BASE = "https://zsk-cmis.cicpa.org.cn/open/enterprise_info_api/v3";
const DEBUG_ORGID = "Q01B4DE61A";
const DEBUG_COOKIE = "XSRF-TOKEN=1bfc37fb-4a73-441e-a0cd-7494fd950d8f; yuqing_whole_jsessionid=85D1CA0091D7B03FEC339E6DEEC0F69F; cicpa_token=fc5fb1aafb424febae9de3062b205275; cicpa_ticket=7d0dffad1fad4fcbac15e6a6ba403e1a; companyVerifyCode=15bdd65567a2958bcacab860c5671f64; userid=891344617; u_name=5&410300060008";
const DEBUG_LOG = false;
const DEBUG_PRINT_API_BODY = false;
const INVESTMENT_MAX_LEVEL = 6;  // 投资公司
const SHAREHOLDER_MAX_LEVEL = 3; // 股东
const INVESTMENT_CONCURRENCY = 6;
const RUN_TIMEOUT_MS = 60 * 1000; 
const REQUEST_TIMEOUT_MS = 10 * 1000;

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const orgid = args.orgid || args.orgId || DEBUG_ORGID;
  const cookie = args.cookie || DEBUG_COOKIE;

  if (!orgid || orgid === "请替换为一级公司orgid") {
    throw new Error("请先在 local-debug.js 顶部设置 DEBUG_ORGID。");
  }
  if (!cookie || cookie === "请替换为完整cookie") {
    throw new Error("请先在 local-debug.js 顶部设置 DEBUG_COOKIE。");
  }

  const collector = createCollector();
  debugLog("开始执行抓取", { orgid });
  startRunTimer(collector);
  await crawlRoot({ orgid, cookie, collector });

  const fileName = `${sanitizeFileName(collector.rootCompanyName || orgid)}_关系链路导出.xls`;
  const outputPath = path.resolve(process.cwd(), fileName);
  const workbookXml = buildWorkbookXml(collector);
  await fs.writeFile(outputPath, workbookXml, "utf8");

  console.log(`导出完成：${outputPath}`);
  console.log(`一级人员条数：${collector.personnelRows.length}`);
  console.log(`对外投资链路条数：${collector.investmentRows.length}`);
  console.log(`最大股东分sheet数：${collector.shareholderSheets.length}`);
  if (collector.timeoutReached) {
    console.log("已达到60秒时限，导出当前已爬取结果。");
  }
}

function createCollector() {
  return {
    rootCompanyName: "",
    personnelRows: [],
    investmentRows: [],
    sourceCompanies: [],
    shareholderSheets: [],
    stopAt: 0,
    timeoutReached: false
  };
}

async function crawlRoot(params) {
  const { orgid, cookie, collector } = params;
  if (shouldStop(collector)) {
    return;
  }
  const rootData = await fetchEnterprisePlate(orgid, cookie);
  const rootName = rootData.orgname || orgid;
  debugLog("一级公司信息", { orgid, rootName });
  collector.rootCompanyName = rootName;
  collector.sourceCompanies.push({ id: orgid, name: rootName, level: 1 });
  collectPersonnel(rootData, collector, rootName);

  await crawlInvestmentNode({
    orgid,
    companyName: rootName,
    level: 1,
    cookie,
    collector,
    seenCompanyNames: new Set([rootName])
  });

  await buildShareholderSheets(cookie, collector);
}

async function crawlInvestmentNode(params) {
  const { orgid, companyName, level, cookie, collector, seenCompanyNames } = params;
  if (shouldStop(collector)) {
    return;
  }
  if (level >= INVESTMENT_MAX_LEVEL) {
    return;
  }
  const data = await fetchEnterprisePlate(orgid, cookie);
  const nextLevel = level + 1;
  if (nextLevel > INVESTMENT_MAX_LEVEL) {
    return;
  }
  await crawlMainInvestmentBranch({
    data,
    companyName,
    nextLevel,
    cookie,
    collector,
    seenCompanyNames
  });
}

async function crawlMainInvestmentBranch(params) {
  const { data, companyName, nextLevel, cookie, collector, seenCompanyNames } = params;
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
      debugLog("跳过非企业对外投资节点", { parent: companyName, nodeName: company.name, type: company.type });
      continue;
    }
    if (seenCompanyNames.has(company.name)) {
      debugLog("跳过重复公司（对外投资）", { parent: companyName, companyName: company.name });
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
    taskFactories.push(async () => safeCrawlInvestmentNode({
      orgid: company.id,
      companyName: company.name,
      level: nextLevel,
      cookie,
      collector,
      seenCompanyNames
    }));
  }
  if (taskFactories.length) {
    await runWithConcurrency(taskFactories, INVESTMENT_CONCURRENCY);
  }
}

async function safeCrawlInvestmentNode(params) {
  try {
    if (shouldStop(params.collector)) {
      return;
    }
    await crawlInvestmentNode(params);
  } catch (error) {
    debugLog("跳过异常公司节点", { companyName: params.companyName, reason: error.message });
  }
}

async function safeFetchPersonExtend(rwid, cookie) {
  try {
    return await fetchPersonExtend(rwid, cookie);
  } catch (error) {
    debugLog("跳过异常自然人节点", { rwid, reason: error.message });
    return null;
  }
}

async function buildShareholderSheets(cookie, collector) {
  if (shouldStop(collector)) {
    return;
  }
  const existingCompanyNames = new Set(collector.sourceCompanies.map((item) => item.name));
  const existingHolderNames = new Set();
  const existingAllNames = new Set(collector.sourceCompanies.map((item) => item.name));
  const pendingCompanies = dedupeCompaniesByName(collector.sourceCompanies).filter((item) => item.level <= SHAREHOLDER_MAX_LEVEL);
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

    debugDivider(`开始处理Sheet源公司：${company.name}`);
    debugLog("开始构建最大股东sheet", { companyName: company.name, orgid: company.id });
    const sheet = await buildSingleShareholderSheet({
      orgid: company.id,
      companyName: company.name,
      sourceLevel: company.level,
      cookie,
      existingCompanyNames,
      existingHolderNames,
      existingAllNames,
      collector
    });
    if (sheet) {
      collector.shareholderSheets.push(sheet);
      existingAllNames.add(sheet.holderName);
      for (const discovered of sheet.discoveredCompanies) {
        existingAllNames.add(discovered.name);
      }
      debugLog("完成最大股东sheet", { sheetName: sheet.sheetName, rowCount: sheet.rows.length });
      for (const discovered of sheet.discoveredCompanies) {
        if (discovered.level > SHAREHOLDER_MAX_LEVEL) {
          continue;
        }
        if (processedSourceNames.has(discovered.name) || pendingNameSet.has(discovered.name)) {
          continue;
        }
        pendingCompanies.push(discovered);
        pendingNameSet.add(discovered.name);
      }
    } else {
      debugLog("未生成最大股东sheet", { companyName: company.name, orgid: company.id });
    }
    debugDivider(`结束处理Sheet源公司：${company.name}`);
  }
}

async function buildSingleShareholderSheet(params) {
  const { orgid, companyName, cookie, existingCompanyNames, existingHolderNames, existingAllNames, collector } = params;
  if (shouldStop(collector)) {
    return null;
  }
  const data = await safeFetchEnterprisePlate(orgid, cookie, "left");
  if (!data) {
    return null;
  }
  const shareholders = findChildrenByNodeNameFromData(data, "股东");
  const maxResult = pickSingleMaxByPercent(shareholders);
  if (!maxResult) {
    debugLog("未命中唯一最大股东", { companyName, orgid, shareholderCount: shareholders.length });
    return null;
  }

  const holder = maxResult.item;
  const holderType = String(holder.type);
  const holderRatio = normalizeRatio(holder.info);
  if (existingCompanyNames.has(holder.name) || existingHolderNames.has(holder.name) || existingAllNames.has(holder.name)) {
    debugLog("跳过创建sheet：最大股东名称已存在", {
      companyName,
      holderName: holder.name,
      holderType
    });
    return null;
  }
  existingHolderNames.add(holder.name);
  debugLog("命中最大股东", {
    companyName,
    holderName: holder.name,
    holderType,
    ratio: holderRatio
  });
  const context = {
    rows: [["一级", "", "", holder.name]],
    seenCompanyNames: new Set([companyName]),
    seenPersonNames: new Set(),
    discoveredCompanies: new Map(),
    collector
  };

  if (holderType === "1") {
    context.seenCompanyNames.add(holder.name);
    debugLog("最大股东是公司，进入公司对外投资递归", { companyName, holderName: holder.name, holderOrgid: holder.id });
    await appendCompanyInvestmentRows({
      orgid: holder.id,
      companyName: holder.name,
      level: 2,
      cookie,
      context
    });
  } else if (holderType === "2") {
    context.seenPersonNames.add(holder.name);
    debugLog("最大股东是自然人，进入自然人股东递归", { companyName, holderName: holder.name, holderRwid: holder.id });
    await appendPersonHoldingRows({
      personId: holder.id,
      parentName: holder.name,
      level: 2,
      fallbackRatio: holderRatio,
      cookie,
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
  const { orgid, companyName, level, cookie, context } = params;
  if (shouldStop(context.collector)) {
    return;
  }
  if (level > SHAREHOLDER_MAX_LEVEL) {
    return;
  }
  const data = await safeFetchEnterprisePlate(orgid, cookie, "left");
  if (!data) {
    return;
  }
  const investments = findChildrenByNodeNameFromData(data, "对外投资");
  debugLog("公司对外投资节点（最大股东sheet）", {
    companyName,
    orgid,
    level: toLevelText(level),
    total: investments.length
  });

  const taskFactories = [];
  for (const company of investments) {
    if (shouldStop(context.collector)) {
      break;
    }
    if (String(company.type) !== "1") {
      debugLog("跳过非企业对外投资节点（最大股东sheet）", {
        parent: companyName,
        nodeName: company.name,
        type: company.type
      });
      continue;
    }
    if (context.seenCompanyNames.has(company.name)) {
      debugLog("跳过重复公司（最大股东sheet对外投资）", {
        parent: companyName,
        companyName: company.name
      });
      continue;
    }

    context.rows.push([
      toLevelText(level),
      companyName,
      normalizeRatio(company.info),
      company.name
    ]);
    recordDiscoveredCompany(context, company.id, company.name, level);
    context.seenCompanyNames.add(company.name);

    if (level + 1 > SHAREHOLDER_MAX_LEVEL) {
      continue;
    }
    taskFactories.push(async () => appendCompanyInvestmentRows({
      orgid: company.id,
      companyName: company.name,
      level: level + 1,
      cookie,
      context
    }));
  }
  if (taskFactories.length) {
    await runWithConcurrency(taskFactories, INVESTMENT_CONCURRENCY);
  }
}

async function appendPersonHoldingRows(params) {
  const { personId, parentName, level, fallbackRatio, cookie, context } = params;
  if (shouldStop(context.collector)) {
    return;
  }
  if (level > SHAREHOLDER_MAX_LEVEL) {
    return;
  }
  const personData = await safeFetchPersonExtend(personId, cookie);
  if (!personData) {
    return;
  }
  const companies = findChildrenByNodeNameFromData(personData, "担任股东");
  debugLog("自然人担任股东节点", {
    parentName,
    personId,
    level: toLevelText(level),
    companyCount: companies.length
  });
  for (const company of companies) {
    if (shouldStop(context.collector)) {
      break;
    }
    if (String(company.type) !== "1") {
      debugLog("跳过自然人担任股东中的非企业节点", { person: parentName, nodeName: company.name, type: company.type });
      continue;
    }
    if (context.seenCompanyNames.has(company.name)) {
      debugLog("跳过重复公司（自然人担任股东）", { person: parentName, companyName: company.name });
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
    if (level + 1 > SHAREHOLDER_MAX_LEVEL) {
      continue;
    }
    await appendCompanyInvestmentRows({
      orgid: company.id,
      companyName: company.name,
      level: level + 1,
      cookie,
      context
    });
  }
}

async function safeFetchEnterprisePlate(orgid, cookie, position) {
  try {
    return await fetchEnterprisePlate(orgid, cookie, position);
  } catch (error) {
    debugLog("跳过异常企业接口", { orgid, reason: error.message });
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

function findChildrenByNodeName(leftNodes, nodeName) {
  const safeNodes = Array.isArray(leftNodes) ? leftNodes : [];
  const target = safeNodes.find((item) => item && item.name === nodeName);
  return target && Array.isArray(target.children) ? target.children : [];
}

function findChildrenByNodeNameFromData(data, nodeName) {
  const leftResult = findChildrenByNodeName(data && data.leftNode, nodeName);
  if (leftResult.length) {
    return leftResult;
  }
  return findChildrenByNodeName(data && data.rightNode, nodeName);
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

function startRunTimer(collector) {
  collector.stopAt = Date.now() + RUN_TIMEOUT_MS;
}

function shouldStop(collector) {
  if (!collector) {
    return false;
  }
  if (collector.timeoutReached) {
    return true;
  }
  if (collector.stopAt && Date.now() >= collector.stopAt) {
    collector.timeoutReached = true;
    return true;
  }
  return false;
}

function pickSingleMaxByPercent(items) {
  const validItems = items
    .map((item) => ({
      item,
      percent: parsePercent(item.info)
    }))
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
  if (!match) {
    return Number.NaN;
  }
  return Number(match[1]);
}

function normalizeRatio(raw) {
  return typeof raw === "string" ? raw.trim() : "";
}

async function fetchEnterprisePlate(orgid, cookie, position) {
  const params = new URLSearchParams({
    orgid,
    v: "v1",
    position: position || ""
  });
  const url = `${API_BASE}/enterprise_plates?${params.toString()}`;
  const body = await fetchJson(url, cookie);
  if (!body || body.status_code !== 0 || !body.data) {
    throw new Error(`企业接口返回异常：${body && body.status_msg ? body.status_msg : "未知错误"}`);
  }
  return body.data;
}

async function fetchPersonExtend(rwid, cookie) {
  const params = new URLSearchParams({
    rwid,
    v: "v1",
    position: "left"
  });
  const url = `${API_BASE}/enterprise_extends?${params.toString()}`;
  const body = await fetchJson(url, cookie);
  if (!body || body.status_code !== 0 || !body.data) {
    throw new Error(`股东接口返回异常：${body && body.status_msg ? body.status_msg : "未知错误"}`);
  }
  return body.data;
}

async function fetchJson(url, cookie) {
  const bodyText = await requestWithRetry(url, cookie, 3);
  if (DEBUG_PRINT_API_BODY) {
    debugLog("接口返回原文", { url, bodyPreview: bodyText.slice(0, 800) });
  }
  try {
    const json = JSON.parse(bodyText);
    debugLog("接口返回摘要", {
      url,
      status_code: json && json.status_code,
      status_msg: json && json.status_msg
    });
    return json;
  } catch (error) {
    throw new Error(`接口返回非JSON：${url}`);
  }
}

async function requestWithRetry(url, cookie, maxAttempts) {
  let lastError = null;
  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      return await httpGet(url, cookie);
    } catch (error) {
      lastError = error;
      if (attempt < maxAttempts) {
        await sleep(attempt * 300);
      }
    }
  }
  throw lastError || new Error("请求失败");
}

function httpGet(url, cookie) {
  return new Promise((resolve, reject) => {
    const request = https.get(
      url,
      {
        headers: {
          Cookie: cookie
        }
      },
      (response) => {
        const { statusCode } = response;
        let raw = "";

        response.setEncoding("utf8");
        response.on("data", (chunk) => {
          raw += chunk;
        });
        response.on("end", () => {
          if (statusCode < 200 || statusCode >= 300) {
            reject(new Error(`请求失败：${statusCode} ${url}`));
            return;
          }
          resolve(raw);
        });
      }
    );

    request.on("error", (error) => {
      reject(new Error(`网络请求异常：${error.message}`));
    });
    request.setTimeout(REQUEST_TIMEOUT_MS, () => {
      request.destroy(new Error(`请求超时：${REQUEST_TIMEOUT_MS}ms`));
    });
  });
}

function sleep(ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

function buildWorkbookXml(collector) {
  const personnelHeaders = ["姓名", "职务"];
  const personnelRows = collector.personnelRows.map((row) => [row.personName, row.position]);

  const relationHeaders = ["层级", "上级公司/上级自然人", "持股比例", "公司/自然人"];
  const relationRows = [
    ["一级", "", "", collector.rootCompanyName],
    ...sortRowsByLevel(collector.investmentRows).map((row) => [row.level, row.parentName, row.ratio, row.subjectName])
  ];

  const sheets = [
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
      sheetName: normalizeSheetName(sheet.sheetName),
      headers: sheet.headers,
      rows: sortSheetRowsByLevel(sheet.rows)
    }))
  ];

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

function escapeXml(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function sanitizeFileName(name) {
  return name.replace(/[\\/:*?"<>|]/g, "_");
}

function normalizeSheetName(name) {
  const safe = name.replace(/[\\/?*:[\]]/g, "_");
  return safe.length > 31 ? safe.slice(0, 31) : safe;
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

function debugLog(message, payload) {
  if (!DEBUG_LOG) {
    return;
  }
  if (payload === undefined) {
    console.log(`[DEBUG] ${message}`);
    return;
  }
  console.log(`[DEBUG] ${message} -> ${safeStringify(payload)}`);
}

function debugDivider(title) {
  if (!DEBUG_LOG) {
    return;
  }
  const text = `[DEBUG] ======== ${title} ========`;
  console.log(text);
}

function safeStringify(value) {
  try {
    return JSON.stringify(value);
  } catch (error) {
    return "[unserializable]";
  }
}

function parseArgs(argv) {
  const result = {};
  for (const arg of argv) {
    if (!arg.startsWith("--")) {
      continue;
    }
    const pureArg = arg.slice(2);
    const idx = pureArg.indexOf("=");
    if (idx < 0) {
      result[pureArg] = "true";
      continue;
    }
    const key = pureArg.slice(0, idx);
    const value = pureArg.slice(idx + 1);
    result[key] = value;
  }
  return result;
}

main().catch((error) => {
  console.error(`执行失败：${error.message}`);
  process.exitCode = 1;
});
