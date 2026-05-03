/**
 * Content script: Titan Webmail DOM automation with fallbacks and safe queries.
 * Exposes checkReady + sendEmail; never throws uncaught errors to the page.
 */

(function () {
  "use strict";

  const SEND_CONFIRM_MS = 12000;
  const POLL_MS = 150;
  const MAX_WAIT_MS = 20000;

  // Simple flag to prevent simultaneous compose operations on THIS tab
  let composeOperationInProgress = false;

  // User-provided Titan XPath fallbacks (used when CSS strategies fail).
  const XPATHS = {
    compose:      "/html/body/div[1]/div/div[5]/div[2]/div[2]/div/div[1]/div[1]/div[1]/div/div[1]/div/div[1]/button/span/div",
    recipient:    "/html/body/div[1]/div/div[5]/div[3]/div/div/div/div/div/div/div[2]/div/div/div/div[1]/div/div[1]/div[2]/div/div/div[1]/div/div[2]",
    body:         "/html/body/div[1]/div/div[5]/div[3]/div/div/div/div/div/div/div[2]/div/div/div/div[1]/div/div[3]/div/div/div/div/div/div[1]/div[1]",
    send:         "/html/body/div[1]/div/div[5]/div[2]/div/div[1]/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/div[1]/div[1]",
    insertHtml:   "/html/body/div[1]/div/div[5]/div[3]/div/div/div/div/div/div/div[2]/div/div/div/div[2]/div/div[1]/div[6]/button",
    closeCompose: "/html/body/div[1]/div/div[5]/div[3]/div/div/div/div/div/div/div[1]/span[2]/span[3]",
  };

  function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }

  function isVisible(el) {
    if (!el || !(el instanceof Element)) return false;
    const st = window.getComputedStyle(el);
    if (st.display === "none" || st.visibility === "hidden" || st.opacity === "0") return false;
    const r = el.getBoundingClientRect();
    return r.width > 0 && r.height > 0;
  }

  function queryXPath(path) {
    try {
      const n = document.evaluate(path, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
      return n instanceof Element ? n : null;
    } catch {
      return null;
    }
  }

  function closestClickable(el) {
    if (!el) return null;
    if (el.matches && el.matches("button, [role='button'], a")) return el;
    return el.closest ? el.closest("button, [role='button'], a") : null;
  }

  /**
   * Multiple strategies to find "To" input (Titan UI may change).
   */
  function findToField() {
    const xTo = queryXPath(XPATHS.recipient);
    if (xTo && isVisible(xTo)) {
      const nested = xTo.matches && xTo.matches("input, textarea, [contenteditable='true']")
        ? xTo
        : xTo.querySelector("input, textarea, [contenteditable='true'], [role='textbox'], [role='combobox']");
      if (nested && isVisible(nested)) return nested;
      return xTo;
    }

    const selectors = [
      'input[placeholder="To"]',
      'input[aria-label="To field"]',
      'input[type="email"]',
      'input[placeholder*="To" i]',
      'input[aria-label*="To" i]',
      'input[name*="to" i]',
      'input[id*="to" i]',
      '[data-testid*="to" i]',
      '[contenteditable="true"][aria-label*="to" i]',
      'textarea[placeholder*="To" i]',
    ];
    for (const sel of selectors) {
      const nodes = Array.from(document.querySelectorAll(sel));
      const hit = nodes.find(isVisible);
      if (hit) return hit;
    }
    const all = Array.from(document.querySelectorAll("input, textarea, [contenteditable='true']"));
    return (
      all.find(
        (el) =>
          isVisible(el) &&
          (/recipient|to\b/i.test(el.getAttribute("placeholder") || "") ||
            /to/i.test(el.getAttribute("aria-label") || ""))
      ) || null
    );
  }

  function findSubjectField() {
    const selectors = [
      'input[placeholder*="Subject" i]',
      'input[aria-label*="Subject" i]',
      'input[name*="subject" i]',
      'input[id*="subject" i]',
      '[data-testid*="subject" i]',
      '[contenteditable="true"][aria-label*="Subject" i]',
      '[contenteditable="true"][placeholder*="Subject" i]',
      '[role="textbox"][aria-label*="Subject" i]',
      '[role="textbox"][placeholder*="Subject" i]',
    ];
    for (const sel of selectors) {
      const nodes = Array.from(document.querySelectorAll(sel));
      const hit = nodes.find((n) => isVisible(n) && (n.matches('input, textarea') || n.getAttribute('contenteditable') === 'true'));
      if (hit) return hit;
    }
    const candidates = Array.from(document.querySelectorAll('input[type="text"], input:not([type]), textarea, [contenteditable="true"], [role="textbox"]'));
    return (
      candidates.find(
        (el) =>
          isVisible(el) &&
          (/subject/i.test(el.getAttribute("placeholder") || "") ||
            /subject/i.test(el.getAttribute("aria-label") || "") ||
            /subject/i.test(el.getAttribute("role") || ""))
      ) || null
    );
  }

  /**
   * Prefer contenteditable inside compose; fallback to iframe body if present.
   */
  function findBodyEditor() {
    const xBody = queryXPath(XPATHS.body);
    if (xBody && isVisible(xBody)) return xBody;

    const selectors = [
      '[role="textbox"][contenteditable="true"]',
      '[contenteditable="true"]',
      ".ProseMirror",
      '[aria-label*="Message" i][contenteditable]',
      "iframe.compose-body",
    ];
    for (const sel of selectors) {
      const el = document.querySelector(sel);
      if (sel.startsWith("iframe") && el && el.tagName === "IFRAME") {
        try {
          const doc = el.contentDocument || el.contentWindow.document;
          const ce = doc.querySelector('[contenteditable="true"]');
          if (ce) return ce;
        } catch {
          /* cross-origin */
        }
      } else if (el && el.getAttribute("contenteditable") === "true" && isVisible(el)) {
        return el;
      }
    }
    const editable = Array.from(document.querySelectorAll('[contenteditable="true"]'));
    const big = editable.filter(isVisible).sort((a, b) => b.innerText.length - a.innerText.length);
    return big[0] || null;
  }

  function findSendButton() {
    // Tier 1: semantic XPath — targets the send button by its known attributes,
    // explicitly excluded buttons inside dialogs (e.g. "Mail subject missing").
    const xSemantic = queryXPath(
      '//button[@data-testid="send-action-btn" or contains(@class,"btn-send") or contains(@class,"send-btn")]' +
      '[not(ancestor::*[@role="dialog"])]'
    );
    if (xSemantic && isVisible(xSemantic) && !xSemantic.disabled) return xSemantic;

    // Tier 2: CSS by data-testid and class, skipping any dialog ancestors.
    for (const sel of ['[data-testid="send-action-btn"]', '.btn-send', '.btn-primary.btn-send']) {
      const el = document.querySelector(sel);
      if (el && isVisible(el) && !el.disabled && !el.closest('[role="dialog"]')) return el;
    }

    // Tier 3: absolute XPath container — look inside for the button.
    const xEl = queryXPath(XPATHS.send);
    if (xEl) {
      const xSend = xEl.querySelector('[data-testid="send-action-btn"], .btn-send, button') || xEl;
      if (xSend && isVisible(xSend) && !xSend.disabled) return xSend;
    }

    // Tier 4: coordinate fallback — element at Titan's known Send button position.
    try {
      const atCoords = document.elementFromPoint(607, 807);
      if (atCoords) {
        const btn = (atCoords.matches && atCoords.matches("button, [role='button']"))
          ? atCoords
          : (atCoords.closest ? atCoords.closest("button, [role='button']") : null);
        if (btn && isVisible(btn) && !btn.disabled && !btn.closest('[role="dialog"]')) return btn;
      }
    } catch { /* ignore */ }

    // Tier 5: scan all visible buttons whose text is exactly "Send" (outside dialogs).
    const buttons = Array.from(document.querySelectorAll("button, [role='button']"));
    return (
      buttons.find(
        (b) => isVisible(b) && !b.disabled &&
               !b.closest('[role="dialog"]') &&
               /^\s*send\s*$/i.test((b.textContent || "").trim())
      ) || null
    );
  }

  function findComposeButton() {
    const xCompose = closestClickable(queryXPath(XPATHS.compose)) || queryXPath(XPATHS.compose);
    if (xCompose && isVisible(xCompose) && !xCompose.disabled) return xCompose;

    const selectors = [
      'button[aria-label*="compose" i]',
      'button[aria-label*="new message" i]',
      'button[data-testid*="compose" i]',
      '[role="button"][aria-label*="compose" i]',
      '[role="button"][data-testid*="compose" i]',
    ];
    for (const sel of selectors) {
      const nodes = Array.from(document.querySelectorAll(sel));
      const hit = nodes.find((el) => isVisible(el) && !el.disabled);
      if (hit) return hit;
    }

    const buttons = Array.from(document.querySelectorAll("button, [role='button'], a"));
    return (
      buttons.find((el) => {
        if (!isVisible(el) || el.disabled) return false;
        const text = `${el.textContent || ""} ${el.getAttribute("aria-label") || ""}`.toLowerCase();
        return /compose|new message|new mail|write|create/i.test(text);
      }) || null
    );
  }

  function detectLoginPage() {
    const pw = document.querySelector('input[type="password"]');
    const loginBtn = Array.from(document.querySelectorAll("button, a")).some((el) =>
      /sign\s*in|log\s*in/i.test(el.textContent || "")
    );
    const url = location.href.toLowerCase();
    if (url.includes("login") || url.includes("signin")) return true;
    if (pw && isVisible(pw) && loginBtn) return true;
    return false;
  }

  function composeFieldsPresent() {
    const to = findToField();
    const sub = findSubjectField();
    const body = findBodyEditor();
    return !!(to && sub && body);
  }

  async function ensureComposeOpen() {
    if (composeFieldsPresent()) return true;

    // Prevent multiple simultaneous compose opens on this tab
    while (composeOperationInProgress) {
      await new Promise(r => setTimeout(r, 100));
    }

    const composeButton = findComposeButton();
    if (!composeButton) return false;

    clickEl(composeButton);
    return waitFor(() => composeFieldsPresent(), MAX_WAIT_MS);
  }

  /**
   * Wait until predicate or timeout (polling + one MutationObserver pass).
   */
  function waitFor(predicate, timeoutMs) {
    const deadline = Date.now() + (timeoutMs || MAX_WAIT_MS);
    return new Promise((resolve) => {
      const obs = new MutationObserver(() => check());
      obs.observe(document.documentElement, { childList: true, subtree: true, attributes: true });
      function check() {
        try {
          if (predicate()) {
            obs.disconnect();
            resolve(true);
            return;
          }
        } catch {
          /* ignore */
        }
        if (Date.now() >= deadline) {
          obs.disconnect();
          resolve(false);
          return;
        }
        setTimeout(check, POLL_MS);
      }
      check();
    });
  }

  function recipientEditableTarget(node) {
    if (!node) return null;
    if (
      node.matches &&
      node.matches("input, textarea, [contenteditable='true'], [role='textbox'], [role='combobox']")
    ) {
      return node;
    }
    return (
      node.querySelector &&
      node.querySelector("input, textarea, [contenteditable='true'], [role='textbox'], [role='combobox']")
    );
  }

  function recipientLooksCommitted(email, toEl) {
    const needle = String(email || "").trim().toLowerCase();
    if (!needle) return false;

    const target = recipientEditableTarget(toEl) || toEl;
    const currentText = fieldText(target).toLowerCase().trim();

    // If the input has been cleared, Titan likely created a chip/token.
    if (currentText === "") return true;

    // Walk up the DOM to find the recipient container around the To field.
    const container = toEl
      ? (toEl.closest('[class*="to"], [class*="recipient"], [class*="compose"], [class*="To"], [class*="Recipient"]') ||
         toEl.parentElement?.parentElement?.parentElement ||
         toEl.parentElement?.parentElement ||
         toEl.parentElement)
      : document.body;

    function matchesCommittedNode(node) {
      if (!node || !isVisible(node)) return false;
      const text = (node.textContent || "").toLowerCase();
      if (!text) return false;
      if (text.includes(needle) && node !== target && !node.contains(target) && node !== toEl) {
        return true;
      }
      const title = (node.getAttribute && node.getAttribute("title")) || "";
      if (title.toLowerCase().includes(needle)) return true;
      const dataEmail = (node.getAttribute && node.getAttribute("data-email")) || "";
      if (dataEmail.toLowerCase().includes(needle)) return true;
      return false;
    }

    if (container) {
      const nodes = Array.from(container.querySelectorAll("*"));
      if (nodes.some(matchesCommittedNode)) return true;
    }

    // Additional global fallback: sometimes the chip exists outside the immediate container.
    const globalCandidates = Array.from(document.querySelectorAll("*[title], *[data-email], span, div, button, a, li"));
    if (globalCandidates.some(matchesCommittedNode)) return true;

    return false;
  }

  function findAutocompleteSuggestion(emailLower) {
    // Broad set of selectors covering common dropdown patterns.
    const candidates = Array.from(document.querySelectorAll(
      '[role="option"], [role="listitem"], [role="menuitem"], ' +
      'ul[role="listbox"] li, ul[role="listbox"] [role="option"], ' +
      '[class*="suggestion"], [class*="Suggestion"], ' +
      '[class*="autocomplete"] li, [class*="Autocomplete"] li, ' +
      '[class*="dropdown"] li, [class*="Dropdown"] li, ' +
      '[class*="contact-list"] li, [class*="ContactList"] li, ' +
      '[class*="contact"] li, [class*="Contact"] li, ' +
      '[class*="result"] li, [class*="Result"] li, ' +
      '[class*="option"], [class*="Option"]'
    ));
    const hit = candidates.find(
      (el) => isVisible(el) && (el.textContent || "").toLowerCase().includes(emailLower)
    );
    if (hit) return hit;

    // Broader last-resort scan over visible dropdown-like nodes.
    const toField = document.querySelector(
      'input[type="email"], input[placeholder*="To" i], [contenteditable="true"][aria-label*="to" i], [role="combobox"]'
    );
    const toRect = toField ? toField.getBoundingClientRect() : null;
    const broader = Array.from(document.querySelectorAll("li, div, span, button, [data-value], [data-email], [title]"));
    return (
      broader.find((el) => {
        if (!isVisible(el)) return false;
        const text = (el.textContent || "").toLowerCase();
        if (!text.includes(emailLower)) return false;
        if (!toRect) return true;
        const r = el.getBoundingClientRect();
        return r.top >= toRect.top && r.left < toRect.right && r.right > toRect.left;
      }) || null
    );
  }

  async function setRecipientValue(toEl, email) {
    const target = recipientEditableTarget(toEl) || toEl;
    const val = String(email == null ? "" : email).trim();
    if (!target || !val) return false;

    const emailLower = val.toLowerCase();

    // Poll until committed or timeout expires.
    async function waitCommitted(ms) {
      const deadline = Date.now() + ms;
      while (Date.now() < deadline) {
        if (recipientLooksCommitted(val, toEl)) return true;
        await sleep(150);
      }
      return recipientLooksCommitted(val, toEl);
    }

    // Tab and non-Enter keys are safe to fire fully; Enter keypress/keyup
    // bubble to Titan's global compose-close handler so we use keydown only.
    function dispatchTab(el) {
      const opts = { key: "Tab", code: "Tab", keyCode: 9, which: 9, bubbles: true, composed: true, cancelable: true };
      el.dispatchEvent(new KeyboardEvent("keydown",  opts));
      el.dispatchEvent(new KeyboardEvent("keypress", opts));
      el.dispatchEvent(new KeyboardEvent("keyup",    opts));
    }

    function dispatchEnterKeydown(el) {
      // keypress + keyup for Enter bubble to Titan's global close handler — keydown only.
      el.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", code: "Enter", keyCode: 13, which: 13, bubbles: true, composed: true, cancelable: true }));
    }

    // Step 1: focus and clear any existing content.
    target.focus();
    await sleep(80);
    try {
      document.execCommand("selectAll", false, null);
      document.execCommand("delete", false, null);
    } catch {
      if (target.tagName === "INPUT" || target.tagName === "TEXTAREA") target.value = "";
      else target.textContent = "";
    }

    // Step 2: insert text via execCommand (React-compatible).
    const inserted = document.execCommand("insertText", false, val);
    if (!inserted) {
      if (target.tagName === "INPUT" || target.tagName === "TEXTAREA") {
        const proto = Object.getPrototypeOf(target);
        const desc = Object.getOwnPropertyDescriptor(proto, "value");
        if (desc && typeof desc.set === "function") desc.set.call(target, val);
        else target.value = val;
        target.dispatchEvent(new Event("input",  { bubbles: true }));
        target.dispatchEvent(new Event("change", { bubbles: true }));
      } else {
        target.textContent = val;
        target.dispatchEvent(new Event("input", { bubbles: true }));
      }
    }

    // Verify text landed; retry via native setter if not.
    await sleep(120);
    if (!fieldText(target).toLowerCase().includes(emailLower)) {
      if (target.tagName === "INPUT" || target.tagName === "TEXTAREA") {
        const proto = Object.getPrototypeOf(target);
        const desc = Object.getOwnPropertyDescriptor(proto, "value");
        if (desc && typeof desc.set === "function") desc.set.call(target, val);
        else target.value = val;
        target.dispatchEvent(new Event("input",  { bubbles: true }));
        target.dispatchEvent(new Event("change", { bubbles: true }));
      } else if (target.getAttribute("contenteditable") === "true") {
        target.textContent = val;
        target.dispatchEvent(new Event("input", { bubbles: true }));
      }
    }

    // Step 3: wait for autocomplete dropdown, then click matching suggestion.
    await sleep(800);
    const suggestion = findAutocompleteSuggestion(emailLower);
    if (suggestion) {
      clickEl(suggestion);
      if (await waitCommitted(800)) return true;
    }

    // Step 4: Enter keydown only — keypress/keyup bubble to Titan's compose-close handler.
    target.focus();
    await sleep(60);
    dispatchEnterKeydown(target);
    if (await waitCommitted(1200)) return true;

    // Step 5: Tab key — moves focus to next field, forcing chip commit on blur.
    target.focus();
    await sleep(60);
    dispatchTab(target);
    if (await waitCommitted(1200)) return true;

    // Step 6: Click the Subject field — most reliable blur trigger for Titan.
    const subjectField = findSubjectField();
    if (subjectField && subjectField !== target) {
      clickEl(subjectField);
      subjectField.focus();
      if (await waitCommitted(1800)) return true;
    }

    // Step 7: Also send Enter keydown to document.activeElement if different from target.
    target.focus();
    await sleep(60);
    const active = document.activeElement;
    dispatchEnterKeydown(target);
    if (active && active !== target) {
      dispatchEnterKeydown(active);
    }
    if (await waitCommitted(1000)) return true;

    // Step 8: comma then Enter keydown (some email UIs use comma as separator).
    target.focus();
    await sleep(60);
    document.execCommand("insertText", false, ",");
    target.dispatchEvent(new Event("input", { bubbles: true }));
    await sleep(200);
    dispatchEnterKeydown(target);
    if (await waitCommitted(800)) return true;

    // Step 9: final blur + click-away.
    target.blur();
    document.body.click();
    return await waitCommitted(1000);
  }

  function setInputValue(el, value) {
    if (!el) return false;
    const text = String(value == null ? "" : value);

    function resolveEditableTarget(node) {
      if (!node) return null;
      if (
        node.matches &&
        node.matches("input, textarea, [contenteditable='true'], [role='textbox'], [role='combobox']")
      ) {
        return node;
      }
      const found = node.querySelector(
        "input, textarea, [contenteditable='true'], [role='textbox'], [role='combobox']"
      );
      return found || null;
    }

    const target = resolveEditableTarget(el) || el;
    const targetTag = target.tagName;

    function setNativeInputValue(input, nextValue) {
      const proto = Object.getPrototypeOf(input);
      const desc = Object.getOwnPropertyDescriptor(proto, "value");
      if (desc && typeof desc.set === "function") {
        desc.set.call(input, nextValue);
      } else {
        input.value = nextValue;
      }
    }

    if (targetTag === "INPUT" || targetTag === "TEXTAREA") {
      target.focus();
      setNativeInputValue(target, "");
      target.dispatchEvent(new Event("input", { bubbles: true }));
      setNativeInputValue(target, text);
      target.dispatchEvent(new Event("input", { bubbles: true }));
      target.dispatchEvent(new Event("change", { bubbles: true }));
      return true;
    }

    if (
      target.getAttribute("contenteditable") === "true" ||
      target.getAttribute("role") === "textbox" ||
      target.getAttribute("role") === "combobox"
    ) {
      target.focus();
      try {
        document.execCommand("selectAll", false, null);
        document.execCommand("delete", false, null);
      } catch {
        target.textContent = "";
      }
      const ok = document.execCommand("insertText", false, text);
      if (!ok) {
        try {
          target.dispatchEvent(
            new InputEvent("beforeinput", {
              bubbles: true,
              cancelable: true,
              inputType: "insertText",
              data: text,
            })
          );
        } catch {}
        target.textContent = text;
      }
      target.dispatchEvent(new Event("input", { bubbles: true }));
      return true;
    }
    return false;
  }

  function findCloseComposeButton() {
    const xClose = closestClickable(queryXPath(XPATHS.closeCompose)) || queryXPath(XPATHS.closeCompose);
    if (xClose && isVisible(xClose)) return xClose;
    const selectors = [
      '[aria-label*="close" i][role="button"]',
      '[aria-label*="discard" i][role="button"]',
      '[title*="close" i]',
      '[title*="discard" i]',
    ];
    for (const sel of selectors) {
      const el = document.querySelector(sel);
      if (el && isVisible(el)) return el;
    }
    return null;
  }

  /**
   * Dismiss any "Discard changes?" confirmation Titan shows after close.
   */
  async function dismissDiscardDialog() {
    await sleep(350);
    const discardBtn = Array.from(document.querySelectorAll("button, [role='button']")).find(
      (b) =>
        isVisible(b) &&
        /^\s*(discard|yes|confirm|don.t save|delete|remove)\s*$/i.test((b.textContent || "").trim())
    );
    if (discardBtn) {
      clickEl(discardBtn);
      await sleep(300);
    }
  }

  /**
   * Close the compose panel and wait until it disappears.
   * Handles the optional "Discard changes?" dialog Titan may show.
   * Safe to call even if compose is already closed.
   */
  async function closeComposeWindow() {
    if (!composeFieldsPresent()) return;

    const btn = findCloseComposeButton();
    if (btn) {
      clickEl(btn);
      await dismissDiscardDialog();
      await waitFor(() => !composeFieldsPresent(), 4000);
    }

    // Fallback: Escape key.
    if (composeFieldsPresent()) {
      document.dispatchEvent(
        new KeyboardEvent("keydown", { key: "Escape", code: "Escape", keyCode: 27, bubbles: true })
      );
      await dismissDiscardDialog();
      await waitFor(() => !composeFieldsPresent(), 3000);
    }
  }

  function findInsertHtmlButton() {
    // Primary: Titan's own data-testid and class.
    const byTestId = document.querySelector('[data-testid="insert-html-button"]');
    if (byTestId && isVisible(byTestId) && !byTestId.disabled) return byTestId;

    const byClass = document.querySelector('.insert-html-button-pendo, [class*="insert-html"]');
    if (byClass && isVisible(byClass) && !byClass.disabled) return byClass;

    // Titan toolbar button: title="Source Code" or data-action="insertHTML".
    const byAttr = document.querySelector('[title="Source Code"], [data-action="insertHTML"]');
    if (byAttr && isVisible(byAttr) && !byAttr.disabled) return closestClickable(byAttr) || byAttr;

    const xBtn = closestClickable(queryXPath(XPATHS.insertHtml)) || queryXPath(XPATHS.insertHtml);
    if (xBtn && isVisible(xBtn)) return xBtn;

    const buttons = Array.from(document.querySelectorAll("button, [role='button']"));
    return (
      buttons.find(
        (b) =>
          isVisible(b) &&
          !b.disabled &&
          /source\s*code|insert\s*html|html\s*template|<>|\(⌘⇧E\)|\(Ctrl\+Shift\+E\)/i.test(
            (b.title || b.getAttribute("aria-label") || b.textContent || "").trim()
          )
      ) || null
    );
  }

  /**
   * Click the "Insert HTML" toolbar button, wait for the dialog, fill the
   * textarea with the rendered HTML, then confirm.  Returns null on success
   * or an error string on failure.
   */
  async function insertBodyViaHtmlDialog(html) {
    // First try keyboard shortcut to open HTML dialog (Cmd+Shift+E / Ctrl+Shift+E)
    const isMac = navigator.platform.toUpperCase().indexOf('MAC') >= 0;
    const shortcutEvent = new KeyboardEvent('keydown', {
      key: 'E',
      code: 'KeyE',
      ctrlKey: !isMac,
      metaKey: isMac,
      shiftKey: true,
      bubbles: true
    });
    document.dispatchEvent(shortcutEvent);
    await sleep(600);

    // Check if dialog opened
    let dialog = document.querySelector(
      "[role='dialog'], [class*='modal'], [class*='dialog'], [class*='Modal'], [class*='Dialog']"
    );
    if (!dialog || !isVisible(dialog)) {
      // Fallback: find and click the button
      const btn = findInsertHtmlButton();
      if (!btn) return "Insert HTML button not found in Titan toolbar";

      clickEl(btn);
      await sleep(600);

      dialog = document.querySelector(
        "[role='dialog'], [class*='modal'], [class*='dialog'], [class*='Modal'], [class*='Dialog']"
      );
      if (!dialog || !isVisible(dialog)) return "Insert HTML dialog did not open";
    }

    const textarea =
      dialog.querySelector("textarea") ||
      dialog.querySelector("input[type='text']") ||
      dialog.querySelector("[contenteditable='true']");
    if (!textarea) return "HTML textarea not found in Insert HTML dialog";

    textarea.focus();
    await sleep(150);
    const proto = Object.getPrototypeOf(textarea);
    const desc =
      Object.getOwnPropertyDescriptor(proto, "value") ||
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, "value") ||
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value");
    if (desc && typeof desc.set === "function") {
      desc.set.call(textarea, html);
    } else {
      textarea.value = html;
    }
    textarea.dispatchEvent(new Event("input",  { bubbles: true }));
    textarea.dispatchEvent(new Event("change", { bubbles: true }));
    await sleep(300);

    const okBtn =
      Array.from(dialog.querySelectorAll("button, [role='button']")).find(
        (b) =>
          !b.disabled &&
          isVisible(b) &&
          /^\s*(ok|insert|apply|confirm|done|save)\s*$/i.test((b.textContent || "").trim())
      ) ||
      Array.from(dialog.querySelectorAll("button, [role='button']")).find(
        (b) => !b.disabled && isVisible(b) && /ok|insert|apply|confirm/i.test(b.textContent || "")
      );

    if (!okBtn) return "Confirm button not found in Insert HTML dialog";

    clickEl(okBtn);
    await waitFor(() => {
      const d = document.querySelector(
        "[role='dialog'], [class*='modal'], [class*='dialog'], [class*='Modal'], [class*='Dialog']"
      );
      return !d || !isVisible(d);
    }, 4000);
    await sleep(400);
    return null;
  }

  function setBodyHtml(el, html) {
    if (!el) return false;
    el.focus();
    if (el.getAttribute("contenteditable") === "true") {
      try {
        document.execCommand("selectAll", false, null);
        document.execCommand("delete", false, null);
        const ok = document.execCommand("insertHTML", false, html);
        if (!ok) {
          el.innerHTML = html;
        }
      } catch {
        el.innerHTML = html;
      }
      el.dispatchEvent(new Event("input", { bubbles: true }));
      return true;
    }
    return false;
  }

  function clickEl(el) {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    const cx = Math.round(rect.left + rect.width / 2);
    const cy = Math.round(rect.top + rect.height / 2);
    // detail:1 is required — el.click() produces detail:0 which React onClick handlers
    // treat as a programmatic (non-user) click and ignore for sensitive actions like Send.
    const downInit = { bubbles: true, cancelable: true, clientX: cx, clientY: cy, button: 0, buttons: 1, view: window };
    const upInit   = { ...downInit, buttons: 0, detail: 1 };
    el.dispatchEvent(new PointerEvent("pointerdown", { ...downInit, isPrimary: true, pointerId: 1 }));
    el.dispatchEvent(new MouseEvent("mousedown",     downInit));
    el.dispatchEvent(new PointerEvent("pointerup",   { ...upInit,  isPrimary: true, pointerId: 1 }));
    el.dispatchEvent(new MouseEvent("mouseup",       upInit));
    el.dispatchEvent(new MouseEvent("click",         upInit));
    return true;
  }

  function fieldText(el) {
    if (!el) return "";
    if (el.value != null) return String(el.value || "").trim();
    return String(el.innerText || el.textContent || "").trim();
  }

  function captureComposeSnapshot(payload) {
    const toEl = findToField();
    const bodyEl = findBodyEditor();
    return {
      payloadTo: String((payload && payload.to) || "").trim().toLowerCase(),
      toBefore: fieldText(toEl).toLowerCase(),
      bodyBeforeLen: fieldText(bodyEl).length,
      hadCompose: composeFieldsPresent(),
    };
  }

  function inferSentFromUiChange(snapshot) {
    if (snapshot && snapshot.hadCompose && !composeFieldsPresent()) return true;

    const toEl = findToField();
    const bodyEl = findBodyEditor();
    const toAfter = fieldText(toEl).toLowerCase();
    const bodyAfterLen = fieldText(bodyEl).length;

    if (snapshot && snapshot.payloadTo) {
      const hadRecipient = snapshot.toBefore.includes(snapshot.payloadTo);
      const hasRecipientNow = toAfter.includes(snapshot.payloadTo);
      if (hadRecipient && !hasRecipientNow && bodyAfterLen <= Math.max(0, snapshot.bodyBeforeLen * 0.2)) {
        return true;
      }
    }

    return false;
  }

  /**
   * Observe toasts / aria-live for success after send.
   * Also handles the "Mail subject missing" confirmation dialog.
   */
  function watchSendConfirmation(snapshot) {
    return new Promise((resolve) => {
      let noSubjectHandled = false;

      const done = (ok) => {
        obs.disconnect();
        clearTimeout(t);
        resolve(ok);
      };
      const t = setTimeout(() => done(inferSentFromUiChange(snapshot)), SEND_CONFIRM_MS);

      function scan() {
        if (snapshot && snapshot.hadCompose && !composeFieldsPresent()) return done(true);

        const live = Array.from(document.querySelectorAll('[role="status"], [role="alert"], [aria-live]'));
        for (const n of live) {
          const tx = (n.textContent || "").toLowerCase();
          if (/sent|message sent|success|delivered/.test(tx)) return done(true);
          if (/fail|error|could not send/.test(tx)) return done(false);
        }

        const dialogs = Array.from(document.querySelectorAll(
          "[role='dialog'], [class*='modal'], [class*='dialog'], [class*='Modal'], [class*='Dialog']"
        ));
        for (const d of dialogs) {
          if (!isVisible(d)) continue;
          const tx = (d.textContent || "").toLowerCase();

          // "Mail subject missing" — click the Send button inside the dialog to confirm.
          if (/subject.*missing|no subject|send without subject|missing.*subject|mail subject/.test(tx)) {
            if (!noSubjectHandled) {
              noSubjectHandled = true;
              const confirmBtn = Array.from(d.querySelectorAll("button, [role='button']")).find(
                (b) => isVisible(b) && !b.disabled && /^\s*send\s*$/i.test((b.textContent || "").trim())
              );
              if (confirmBtn) { confirmBtn.focus(); clickEl(confirmBtn); }
            }
            return; // keep watching — compose will close after confirm
          }

          if (/recipient missing|provide one or more recipients|no recipients|add a recipient/.test(tx)) return done(false);
          if (/please add a message|message is empty|add.*message|write.*message/.test(tx)) return done(false);
          if (/could not send|failed|error|something went wrong/.test(tx)) return done(false);
        }

        const toast = document.querySelector(".toast, .snackbar, [class*='toast'], [class*='notification']");
        if (toast && isVisible(toast)) {
          const tx = (toast.textContent || "").toLowerCase();
          if (/sent|success/.test(tx)) return done(true);
        }
      }

      const obs = new MutationObserver(scan);
      obs.observe(document.documentElement, { childList: true, subtree: true, characterData: true });
      scan();
    });
  }

  async function findComposeFieldsWithRetry() {
    for (let attempt = 0; attempt < 3; attempt++) {
      await ensureComposeOpen();
      await waitFor(() => composeFieldsPresent(), MAX_WAIT_MS);
      const to = findToField();
      const sub = findSubjectField();
      const body = findBodyEditor();
      if (to && sub && body) return { to, sub, body };
      await sleep(400 * (attempt + 1));
    }
    return null;
  }

  async function applyComposePayload(fields, payload) {
    const { to, sub } = fields;

    if (!(await setRecipientValue(to, payload.to))) return "Could not set To field";
    await sleep(700);

    sub.focus();
    await sleep(200);
    if (!setInputValue(sub, payload.subject)) return "Could not set Subject";
    await sleep(600);

    const bodyErr = await insertBodyViaHtmlDialog(payload.bodyHtml);
    if (bodyErr) {
      const bodyEl = findBodyEditor();
      if (!bodyEl || !setBodyHtml(bodyEl, payload.bodyHtml)) {
        return "Could not set body (dialog: " + bodyErr + ")";
      }
    }
    await sleep(500);

    return null;
  }

  async function clickSend(_fields, payload) {
    const btn = findSendButton();
    if (!btn) return "Send button not found";
    const snap = captureComposeSnapshot(payload);
    clickEl(btn);
    const ok = await watchSendConfirmation(snap);
    return ok ? null : "Send confirmation not detected or message not sent";
  }

  async function sendEmailAction(payload) {
    if (detectLoginPage()) {
      return { ok: false, error: "Login page — please sign in to Titan." };
    }

    // Prevent simultaneous compose operations on this tab
    while (composeOperationInProgress) {
      await new Promise(r => setTimeout(r, 50));
    }
    composeOperationInProgress = true;

    try {
      // ── STEP 0: Always close any existing compose, then open a fresh one ─────
      await closeComposeWindow();
      await sleep(400);

      const composeBtn = findComposeButton();
      if (!composeBtn) {
        return { ok: false, error: "Compose button not found on Titan page." };
      }
      clickEl(composeBtn);
      const composeReady = await waitFor(() => composeFieldsPresent(), MAX_WAIT_MS);
      if (!composeReady) {
        return { ok: false, error: "Compose window did not open in time." };
      }
      await sleep(600); // let React finish rendering all sub-fields

      // ── STEP 1: Recipient ────────────────────────────────────────────────────
      const toEl = findToField();
      if (!toEl) return { ok: false, error: "To field not found after compose opened." };

      const recipOk = await setRecipientValue(toEl, payload.to);
      if (!recipOk) return { ok: false, error: `Could not commit recipient: ${payload.to}` };
      await sleep(500);

      // ── STEP 2: Subject ──────────────────────────────────────────────────────
      const subEl = findSubjectField();
      if (!subEl) return { ok: false, error: "Subject field not found." };
      subEl.focus();
      await sleep(250);
      if (!setInputValue(subEl, payload.subject)) return { ok: false, error: "Could not set Subject." };
      await sleep(400);
      subEl.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", code: "Tab", keyCode: 9, bubbles: true }));
      subEl.dispatchEvent(new KeyboardEvent("keyup",   { key: "Tab", code: "Tab", keyCode: 9, bubbles: true }));
      subEl.blur();
      await sleep(500);

      // ── STEP 3: HTML Body ────────────────────────────────────────────────────
      const bodyErr = await insertBodyViaHtmlDialog(payload.bodyHtml);
      if (bodyErr) {
        const bodyElFb = findBodyEditor();
        if (!bodyElFb || !setBodyHtml(bodyElFb, payload.bodyHtml)) {
          return { ok: false, error: "Could not set body — dialog: " + bodyErr };
        }
      }
      await sleep(500);

      // Verify body actually has content.
      const bodyEl = findBodyEditor();
      if (!bodyEl || !(bodyEl.innerText || "").trim()) {
        return { ok: false, error: "Body is empty after HTML insert — aborting send." };
      }

      // ── STEP 4: Send ─────────────────────────────────────────────────────────
      const sendBtn = findSendButton();
      if (!sendBtn) return { ok: false, error: "Send button not found." };

      // Focus the button so it's the active element, then click.
      sendBtn.focus();
      await sleep(150);

      const snap = captureComposeSnapshot(payload);
      clickEl(sendBtn);

      // Retry once after 2 s if Titan hasn't responded (slow render or missed click).
      const retryTimer = setTimeout(() => {
        if (composeFieldsPresent()) {
          const btn = findSendButton();
          if (btn) { btn.focus(); clickEl(btn); }
        }
      }, 2000);

      const sentOk = await watchSendConfirmation(snap);
      clearTimeout(retryTimer);
      if (!sentOk) return { ok: false, error: "Send did not complete — Titan may have shown an error." };

      return { ok: true };
    } finally {
      try {
        if (composeFieldsPresent()) {
          await closeComposeWindow();
          await sleep(300);
        }
      } catch {
        // Ignore cleanup failures so the worker can continue.
      } finally {
        composeOperationInProgress = false;
      }
    }
  }

  async function checkReadyAction() {
    if (detectLoginPage()) {
      return { ready: false, reason: "login" };
    }
    // Don't open compose here — sendEmailAction always closes and reopens it.
    // Just confirm the page is in a state where compose can be opened.
    if (composeFieldsPresent()) {
      return { ready: true, reason: "ok" };
    }
    const canOpenCompose = findComposeButton();
    return canOpenCompose
      ? { ready: true, reason: "ok" }
      : { ready: false, reason: "compose" };
  }

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (!msg || !msg.action) return;

    if (msg.action === "checkReady") {
      (async () => {
        try {
          sendResponse(await checkReadyAction());
        } catch (e) {
          sendResponse({ ready: false, reason: "error", error: String(e && e.message) });
        }
      })();
      return true;
    }

    if (msg.action === "sendEmail") {
      (async () => {
        try {
          const res = await sendEmailAction(msg.payload || {});
          sendResponse(res);
        } catch (e) {
          sendResponse({ ok: false, error: String(e && e.message) });
        }
      })();
      return true;
    }

    return false;
  });
})();
