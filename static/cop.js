(function () {
  function digitsOnly(text) {
    if (text === null || text === undefined) return "";
    return String(text).replace(/\D/g, "");
  }

  function normalizeIntPart(text) {
    if (text === null || text === undefined) return "";
    let s = String(text).trim();
    if (s === "") return "";
    s = s.replace(/\$/g, "").replace(/\s/g, "");
    const parts = s.split(",");
    let intPart = parts[0] || "";
    intPart = intPart.replace(/\./g, "");
    intPart = intPart.replace(/\D/g, "");
    return intPart;
  }

  function formatCOP(value) {
    const num = parseCOPToNumber(value);
    const fixed = num.toFixed(2);
    const parts = fixed.split(".");
    const intPart = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".");
    return "$" + intPart + "," + (parts[1] || "00");
  }

  function formatCOPFromDigits(digits) {
    return formatCOP(digits);
  }

  function parseCOPToNumber(text) {
    if (text === null || text === undefined) return 0;
    let s = String(text).trim();
    if (s === "") return 0;
    s = s.replace(/\$/g, "").replace(/\s/g, "");
    if (s.includes(",")) {
      s = s.replace(/\./g, "");
      s = s.replace(/,/g, ".");
    } else if ((s.match(/\./g) || []).length > 1) {
      s = s.replace(/\./g, "");
    } else if (s.includes(".")) {
      const parts = s.split(".");
      const dec = parts[1] || "";
      if (dec.length > 2) {
        s = parts.join("");
      }
    }
    s = s.replace(/[^0-9.]/g, "");
    const num = parseFloat(s);
    return Number.isNaN(num) ? 0 : num;
  }

  function formatIdCO(value) {
    const digits = digitsOnly(value);
    if (!digits) return "";
    return digits.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  }

  function calcNITDV(nitDigits) {
    const digits = digitsOnly(nitDigits);
    if (!digits) return "";
    const weights = [3, 7, 13, 17, 19, 23, 29, 37, 41, 43, 47, 53, 59, 67, 71];
    let sum = 0;
    let idx = 0;
    for (let i = digits.length - 1; i >= 0; i--) {
      const d = parseInt(digits[i], 10);
      const w = weights[idx] || 0;
      sum += d * w;
      idx += 1;
    }
    const mod = sum % 11;
    if (mod === 0 || mod === 1) return String(mod);
    return String(11 - mod);
  }

  function attachCOPInputs() {
    const inputs = document.querySelectorAll(".cop-input");
    inputs.forEach((input) => {
      if (input.value) {
        input.value = formatCOP(input.value);
      }
      input.addEventListener("focus", () => {
        input.value = normalizeIntPart(input.value);
      });
      input.addEventListener("blur", () => {
        input.value = formatCOP(input.value);
      });
      input.addEventListener("input", () => {
        input.value = normalizeIntPart(input.value);
      });
    });

    document.querySelectorAll("form").forEach((form) => {
      form.addEventListener("submit", () => {
        form.querySelectorAll(".cop-input").forEach((input) => {
          input.value = String(parseCOPToNumber(input.value));
        });
      });
    });
  }

  function attachCopItemDisplay() {
    const inputs = document.querySelectorAll(".cop-item-display");
    inputs.forEach((input) => {
      const hiddenName = input.dataset.hiddenTarget;
      const hidden = hiddenName ? input.form?.querySelector(`input[name="${hiddenName}"]`) : null;

      function updateHidden() {
        if (!hidden) return;
        const intPart = normalizeIntPart(input.value);
        hidden.value = intPart || "0";
      }

      if (input.value) {
        input.value = formatCOP(input.value);
      }
      updateHidden();

      input.addEventListener("focus", () => {
        input.value = normalizeIntPart(input.value);
        updateHidden();
      });
      input.addEventListener("input", () => {
        input.value = normalizeIntPart(input.value);
        updateHidden();
      });
      input.addEventListener("blur", () => {
        input.value = formatCOP(input.value);
        updateHidden();
      });
    });
  }

  function attachIdInputs() {
    const inputs = document.querySelectorAll(".id-co");
    inputs.forEach((input) => {
      if (input.value) {
        input.value = formatIdCO(input.value);
      }
      input.addEventListener("input", () => {
        input.value = formatIdCO(input.value);
      });
      input.addEventListener("blur", () => {
        input.value = formatIdCO(input.value);
      });
    });

    document.querySelectorAll("form").forEach((form) => {
      form.addEventListener("submit", () => {
        form.querySelectorAll(".id-co").forEach((input) => {
          input.value = digitsOnly(input.value);
        });
      });
    });
  }

  function attachNitDv() {
    const forms = document.querySelectorAll("form");
    forms.forEach((form) => {
      const idInput = form.querySelector(".nit-id");
      const typeSelect = form.querySelector(".tipo-id");
      const dvBadge = form.querySelector(".nit-dv");
      if (!idInput || !typeSelect || !dvBadge) return;

      function updateDv() {
        const tipo = (typeSelect.value || "").toUpperCase();
        if (tipo !== "NIT") {
          dvBadge.textContent = "DV: -";
          dvBadge.classList.add("d-none");
          return;
        }
        const dv = calcNITDV(idInput.value);
        dvBadge.textContent = dv ? "DV: " + dv : "DV: -";
        dvBadge.classList.remove("d-none");
      }

      idInput.addEventListener("input", updateDv);
      idInput.addEventListener("blur", updateDv);
      typeSelect.addEventListener("change", updateDv);
      updateDv();
    });
  }

  window.formatCOP = formatCOP;
  window.parseCOP = parseCOPToNumber;
  window.formatCOPFromDigits = formatCOPFromDigits;
  window.parseCOPToNumber = parseCOPToNumber;
  window.formatIdCO = formatIdCO;
  window.digitsOnly = digitsOnly;
  window.calcNITDV = calcNITDV;

  document.addEventListener("DOMContentLoaded", () => {
    attachCOPInputs();
    attachCopItemDisplay();
    attachIdInputs();
    attachNitDv();
  });
})();
