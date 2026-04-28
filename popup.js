const DEFAULT_SETTINGS = {
  timeoutSeconds: 60,
  investmentLevel: 6,
  shareholderLevel: 3
};

const timeoutInput = document.querySelector("#timeoutSeconds");
const investmentInput = document.querySelector("#investmentLevel");
const shareholderInput = document.querySelector("#shareholderLevel");
const getCookieBtn = document.querySelector("#getCookieBtn");
const extractBtn = document.querySelector("#extractBtn");
const resetBtn = document.querySelector("#resetBtn");
const statusEl = document.querySelector("#status");

init().catch((error) => {
  setStatus(`初始化失败：${error.message}`);
});

async function init() {
  const settings = await loadSettings();
  applySettingsToForm(settings);

  getCookieBtn.addEventListener("click", async () => {
    setStatus("读取cookie中...");
    try {
      const response = await sendToActiveTab({ type: "GET_COOKIE" });
      if (!response || !response.ok) {
        throw new Error((response && response.error) || "获取cookie失败");
      }
      await chrome.storage.local.set({ extractedCookie: response.cookie || "" });
      setStatus(`cookie已获取，长度：${(response.cookie || "").length}`);
    } catch (error) {
      setStatus(`获取cookie失败：${error.message}`);
    }
  });

  extractBtn.addEventListener("click", async () => {
    const settingsFromForm = readSettingsFromForm();
    await chrome.storage.local.set({ exportSettings: settingsFromForm });

    setStatus("开始提取，请稍候...");
    try {
      const response = await sendToActiveTab({
        type: "EXTRACT_EXCEL",
        settings: settingsFromForm
      });
      if (!response || !response.ok) {
        throw new Error((response && response.error) || "提取失败");
      }
      const lines = [];
      lines.push(response.timedOut ? "已超时导出部分结果。" : "提取完成并已导出。");
      if (response.exportDebug) {
        lines.push(`sheet数量(原始/归一化)：${response.exportDebug.rawSheetCount}/${response.exportDebug.normalizedSheetCount}`);
        lines.push(`sheet名：${(response.exportDebug.normalizedNames || []).join(" | ")}`);
        lines.push(
          `重名sheet：${
            (response.exportDebug.duplicateNames || []).length
              ? response.exportDebug.duplicateNames.join(" | ")
              : "无"
          }`
        );
      }
      setStatus(lines.join("\n"));
    } catch (error) {
      setStatus(`提取失败：${error.message}`);
    }
  });

  resetBtn.addEventListener("click", async () => {
    await chrome.storage.local.set({
      exportSettings: DEFAULT_SETTINGS,
      extractedCookie: ""
    });
    applySettingsToForm(DEFAULT_SETTINGS);
    setStatus("设置已重置。");
  });
}

function applySettingsToForm(settings) {
  timeoutInput.value = settings.timeoutSeconds;
  investmentInput.value = settings.investmentLevel;
  shareholderInput.value = settings.shareholderLevel;
}

function readSettingsFromForm() {
  return {
    timeoutSeconds: normalizePositiveInt(timeoutInput.value, DEFAULT_SETTINGS.timeoutSeconds),
    investmentLevel: normalizePositiveInt(investmentInput.value, DEFAULT_SETTINGS.investmentLevel),
    shareholderLevel: normalizePositiveInt(shareholderInput.value, DEFAULT_SETTINGS.shareholderLevel)
  };
}

async function loadSettings() {
  const result = await chrome.storage.local.get("exportSettings");
  const stored = result.exportSettings || {};
  return {
    timeoutSeconds: normalizePositiveInt(stored.timeoutSeconds, DEFAULT_SETTINGS.timeoutSeconds),
    investmentLevel: normalizePositiveInt(stored.investmentLevel, DEFAULT_SETTINGS.investmentLevel),
    shareholderLevel: normalizePositiveInt(stored.shareholderLevel, DEFAULT_SETTINGS.shareholderLevel)
  };
}

function normalizePositiveInt(value, fallback) {
  const num = Number(value);
  if (!Number.isFinite(num) || num <= 0) {
    return fallback;
  }
  return Math.floor(num);
}

async function sendToActiveTab(message) {
  const [tab] = await chrome.tabs.query({
    active: true,
    currentWindow: true
  });
  if (!tab || !tab.id) {
    throw new Error("未找到当前活动标签页");
  }
  if (!isSupportedPage(tab.url || "")) {
    throw new Error("请先切换到目标网站页面后再操作");
  }
  try {
    return await chrome.tabs.sendMessage(tab.id, message);
  } catch (error) {
    if (!String(error.message || "").includes("Receiving end does not exist")) {
      throw error;
    }
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files: ["content.js"]
    });
    return chrome.tabs.sendMessage(tab.id, message);
  }
}

function setStatus(text) {
  statusEl.textContent = text;
}

function isSupportedPage(url) {
  return url.startsWith("https://zsk-cmis.cicpa.org.cn/");
}
