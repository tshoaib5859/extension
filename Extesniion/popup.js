/**
 * Popup: CSV/template input, controls, progress, logs; syncs with chrome.storage.
 */

(function () {
  "use strict";

  const U = self.TitanUtils;

  const el = {
    file: document.getElementById("file"),
    csv: document.getElementById("csv"),
    scriptUrl: document.getElementById("scriptUrl"),
    subject: document.getElementById("subject"),
    template: document.getElementById("template"),
    delay: document.getElementById("delay"),
    jitter: document.getElementById("jitter"),
    max: document.getElementById("max"),
    parallelTabs: document.getElementById("parallelTabs"),
    btnStart: document.getElementById("btnStart"),
    btnPause: document.getElementById("btnPause"),
    btnResume: document.getElementById("btnResume"),
    btnStop: document.getElementById("btnStop"),
    btnClearSent: document.getElementById("btnClearSent"),
    btnClearLogs: document.getElementById("btnClearLogs"),
    statusDot: document.getElementById("statusDot"),
    statusText: document.getElementById("statusText"),
    statusSub: document.getElementById("statusSub"),
    progress: document.getElementById("progress"),
    progressLabel: document.getElementById("progressLabel"),
    logs: document.getElementById("logs"),
    // Template notice
    templateNotice: document.getElementById("templateNotice"),
  };

  function send(type, payload) {
    return new Promise((resolve) => {
      chrome.runtime.sendMessage({ type, ...payload }, (res) => {
        const err = chrome.runtime.lastError;
        if (err) resolve({ ok: false, error: err.message });
        else resolve(res || { ok: false });
      });
    });
  }

  function formatStatus(data) {
    const rs = data[U.STORAGE_KEYS.RUN_STATUS] || U.RUN_STATUS.IDLE;
    const pauseReq = !!data[U.STORAGE_KEYS.PAUSE_REQUESTED];
    el.statusDot.className = "dot";
    if (rs === U.RUN_STATUS.RUNNING && pauseReq) {
      el.statusText.textContent = "Running — pausing after current";
      el.statusSub.textContent = "Finishing the in-flight message, then pauses.";
      el.statusDot.classList.add("run");
      return;
    }
    if (rs === U.RUN_STATUS.RUNNING) {
      el.statusText.textContent = "Running";
      el.statusSub.textContent = "Sending through the Titan tab.";
      el.statusDot.classList.add("run");
      return;
    }
    if (rs === U.RUN_STATUS.PAUSED) {
      el.statusText.textContent = "Paused";
      el.statusSub.textContent = "Resume when Titan compose is ready.";
      el.statusDot.classList.add("pause");
      return;
    }
    if (rs === U.RUN_STATUS.STOPPED) {
      el.statusText.textContent = "Stopped";
      el.statusSub.textContent = "Queue halted. Start again with new data if needed.";
      el.statusDot.classList.add("stop");
      return;
    }
    el.statusText.textContent = "Idle";
    el.statusSub.textContent = "Configure template and Start.";
    el.statusDot.classList.add("idle");
  }

  function renderLogs(logs) {
    el.logs.innerHTML = "";
    const list = Array.isArray(logs) ? logs.slice(-40) : [];
    for (const entry of list) {
      const p = document.createElement("p");
      p.className = "log-line " + (entry.level || "info");
      p.textContent = `[${entry.ts || ""}] ${entry.message || ""}`;
      el.logs.appendChild(p);
    }
    el.logs.scrollTop = el.logs.scrollHeight;
  }

  function renderProgress(data) {
    const q = data[U.STORAGE_KEYS.QUEUE] || [];
    const idx = data[U.STORAGE_KEYS.CURRENT_INDEX] | 0;
    const total = q.length;
    const done = Math.min(idx, total);
    el.progress.max = Math.max(1, total);
    el.progress.value = total ? done : 0;
    el.progressLabel.textContent = total ? `${done} / ${total}` : "0 / 0";
  }

  async function refresh() {
    const res = await send("GET_STATE");
    if (!res || !res.ok || !res.data) return;
    const data = res.data;
    formatStatus(data);
    renderLogs(data[U.STORAGE_KEYS.LOGS]);
    renderProgress(data);

    if (!el.template.value && data[U.STORAGE_KEYS.TEMPLATE]) {
      el.template.value = data[U.STORAGE_KEYS.TEMPLATE];
    }
    // Pre-fill with default template if still empty after storage check.
    if (!el.template.value.trim()) {
      el.template.value = U.DEFAULT_BODY_TEMPLATE;
    }

    if (!el.subject.value && data[U.STORAGE_KEYS.SUBJECT_TEMPLATE]) {
      el.subject.value = data[U.STORAGE_KEYS.SUBJECT_TEMPLATE];
    }
    // Pre-fill subject with default if empty.
    if (!el.subject.value.trim()) {
      el.subject.value = U.DEFAULT_SUBJECT;
    }

    // Restore saved script URL.
    if (!el.scriptUrl.value && data[U.STORAGE_KEYS.SCRIPT_URL]) {
      el.scriptUrl.value = data[U.STORAGE_KEYS.SCRIPT_URL];
    }
    const d = data[U.STORAGE_KEYS.DELAY_MS];
    const j = data[U.STORAGE_KEYS.JITTER_MS];
    const m = data[U.STORAGE_KEYS.MAX_PER_RUN];
    const pt = data[U.STORAGE_KEYS.PARALLEL_TABS];
    if (d != null) el.delay.value = String(d);
    if (j != null) el.jitter.value = String(j);
    if (m != null) el.max.value = String(m);
    if (pt != null) el.parallelTabs.value = String(pt);
  }

  // ────────────── Template prompt ──────────────

  /**
   * Highlight the template field and show the notice asking the user
   * to enter their email body template.
   */
  function promptForTemplate() {
    el.template.classList.add("field-error");
    el.templateNotice.classList.add("visible");
    el.template.scrollIntoView({ behavior: "smooth", block: "center" });
    el.template.focus();
  }

  function clearTemplatePrompt() {
    el.template.classList.remove("field-error");
    el.templateNotice.classList.remove("visible");
  }

  // Remove error state once the user starts typing
  el.template.addEventListener("input", () => {
    if (el.template.value.trim()) clearTemplatePrompt();
  });

  el.file.addEventListener("change", (ev) => {
    const f = ev.target.files && ev.target.files[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = () => {
      el.csv.value = String(reader.result || "");
    };
    reader.readAsText(f);
  });

  el.btnStart.addEventListener("click", async () => {
    try {
      const csvText        = el.csv.value || "";
      const template       = el.template.value || "";
      const subjectTemplate = el.subject.value || "";
      const scriptUrl      = (el.scriptUrl.value || "").trim();
      const delayMs        = parseInt(el.delay.value, 10);
      const jitterMs       = parseInt(el.jitter.value, 10);
      const maxPerRun      = parseInt(el.max.value, 10);
      const parallelTabs   = parseInt(el.parallelTabs.value, 10);

      // Require template — show inline prompt instead of bare alert
      if (!template.trim()) {
        promptForTemplate();
        return;
      }
      clearTemplatePrompt();

      // Persist the script URL so it survives popup close.
      if (scriptUrl) {
        chrome.storage.local.set({ [U.STORAGE_KEYS.SCRIPT_URL]: scriptUrl });
      }

      el.btnStart.disabled = true;
      const res = await send("START", {
        csvText,
        template,
        subjectTemplate,
        scriptUrl,
        delayMs:  Number.isFinite(delayMs)  ? delayMs  : U.DEFAULT_DELAY_MS,
        jitterMs: Number.isFinite(jitterMs) ? jitterMs : U.DEFAULT_JITTER_MS,
        maxPerRun: Number.isFinite(maxPerRun) ? maxPerRun : U.DEFAULT_MAX_PER_RUN,
        parallelTabs: Number.isFinite(parallelTabs) ? Math.max(1, Math.min(parallelTabs, 10)) : U.DEFAULT_PARALLEL_TABS,
      });

      if (!res.ok) {
        let msg = res.error || "Start failed.";
        if (Array.isArray(res.skipped) && res.skipped.length) {
          const examples = res.skipped
            .slice(0, 5)
            .map((s) => `Line ${s.line}: ${s.reason}`)
            .join("\n");
          msg += `\n\nFirst skipped rows:\n${examples}`;
        }
        alert(msg);
        return;
      }

      // Show how many rows are pending in the sheet (sheet mode only).
      if (res.pending != null) {
        await appendLocalLog(`Sheet mode: ${res.queued} loaded, ${res.pending} total pending in sheet.`);
      }

      await refresh();
    } catch (e) {
      alert(`Start failed: ${String((e && e.message) || e)}`);
    } finally {
      el.btnStart.disabled = false;
    }
  });

  function appendLocalLog(msg) {
    // Just refresh — background already logged it.
    return refresh();
  }

  el.btnPause.addEventListener("click", async () => {
    await send("PAUSE");
    await refresh();
  });

  el.btnResume.addEventListener("click", async () => {
    await send("RESUME");
    await refresh();
  });

  el.btnStop.addEventListener("click", async () => {
    await send("STOP");
    await refresh();
  });

  el.btnClearSent.addEventListener("click", async () => {
    await send("CLEAR_SENT_HISTORY");
    await refresh();
  });

  el.btnClearLogs.addEventListener("click", async () => {
    await send("CLEAR_LOGS");
    await refresh();
  });

  chrome.storage.onChanged.addListener((changes, area) => {
    if (area === "local") refresh();
  });

  refresh();
})();
