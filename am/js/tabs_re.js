(function () {
    "use strict";

    /* ==========================
     Helpers seguros
  ========================== */
    const qs = (sel, ctx = document) => ctx.querySelector(sel);
    const qsa = (sel, ctx = document) => [...ctx.querySelectorAll(sel)];

    let currentTab = 0;
    let showTab = () => {}; // dummy para que nunca sea undefined

    /* ==========================
     Lógica padres / hijos
  ========================== */
    function initPadres() {
        const padres = qsa(".padre");
        const botones = qsa(".tab-button-curso");

        if (!padres.length || !botones.length) return;

        botones.forEach((button) => {
            button.addEventListener("click", function () {
                if (!this.classList.contains("padre") && !this.classList.contains("hijo")) {
                    qsa(".hijo").forEach((h) => (h.style.display = "none"));
                }

                if (this.classList.contains("padre")) {
                    qsa(".padre").forEach((p) => p.classList.remove("active"));
                    qsa(".hijo").forEach((h) => (h.style.display = "none"));

                    this.classList.add("active");

                    let sibling = this.nextElementSibling;
                    while (sibling && sibling.classList.contains("hijo")) {
                        sibling.style.display = "block";
                        sibling = sibling.nextElementSibling;
                    }
                }
            });
        });

        // Auto-click primer padre si existe
        padres[0]?.click();
    }

    /* ==========================
     Tabs responsive
  ========================== */
    function checkWindowSize() {
        const tabs = qsa(".tab-content-curso");
        const buttons = qsa(".tab-button-curso");

        if (!tabs.length || !buttons.length) return;

        updateButtonVisibility();

        /* ===== MOBILE ===== */
        if (window.innerWidth < 801) {
            showTab = function (index = 0) {
                if (!tabs[index]) return;

                tabs.forEach((t) => (t.style.display = "none"));
                buttons.forEach((b) => b.classList.remove("active"));

                tabs[index].style.display = "block";
                buttons[index]?.classList.add("active");
                currentTab = index;

                updateButtonVisibility();
            };
        } else {
            /* ===== DESKTOP ===== */
            showTab = function (index = 0) {
                if (!tabs[index]) return;

                tabs.forEach((t) => {
                    t.style.display = "none";
                    t.classList.remove("active-tab-content");
                });
                buttons.forEach((b) => b.classList.remove("active"));

                tabs[index].style.display = "block";
                tabs[index].classList.add("active-tab-content");
                buttons[index]?.classList.add("active");

                currentTab = index;
                updateButtonVisibility();
            };
        }
    }

    /* ==========================
     Prev / Next seguros
  ========================== */
    function updateButtonVisibility() {
        const prev = qs(".pre");
        const next = qs(".next");
        const total = qsa(".tab-button-curso").length;

        if (!prev || !next) return;

        prev.style.visibility = currentTab === 0 ? "hidden" : "visible";
        next.style.visibility = currentTab >= total - 1 ? "hidden" : "visible";
    }

    window.showNextTab = () => showTab(currentTab + 1);
    window.showPreviousTab = () => showTab(currentTab - 1);

    /* ==========================
     Init GLOBAL (DOM READY)
  ========================== */
    document.addEventListener("DOMContentLoaded", () => {
        console.log("Script Moodle seguro cargado ✅");

        initPadres();
        checkWindowSize();

        // listeners tabs
        qsa(".tab-button-curso").forEach((btn, i) => {
            btn.addEventListener("click", () => showTab(i));
        });

        // tab inicial
        showTab(0);
    });

    window.addEventListener("resize", checkWindowSize);
})();
