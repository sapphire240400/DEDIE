(function () {
    "use strict";

    const qs = (s, c = document) => c.querySelector(s);
    const qsa = (s, c = document) => [...c.querySelectorAll(s)];

    let currentTab = 0;
    let showTab = () => {};

    /* =========================
     PADRES / HIJOS
  ========================= */
    function initPadres() {
        const padres = qsa(".padre");
        if (!padres.length) return;

        qsa(".tab-button-curso").forEach((btn) => {
            btn.addEventListener("click", function () {
                if (!this.classList.contains("padre") && !this.classList.contains("hijo")) {
                    qsa(".hijo").forEach((h) => (h.style.display = "none"));
                }

                if (this.classList.contains("padre")) {
                    qsa(".padre").forEach((p) => p.classList.remove("active"));
                    qsa(".hijo").forEach((h) => (h.style.display = "none"));

                    this.classList.add("active");

                    let sib = this.nextElementSibling;
                    while (sib && sib.classList.contains("hijo")) {
                        sib.style.display = "block";
                        sib = sib.nextElementSibling;
                    }
                }
            });
        });

        padres[0]?.click();
    }

    /* =========================
     SHOW TAB (RESPONSIVE)
  ========================= */
    function buildShowTab() {
        const tabs = qsa(".tab-content-curso");
        const buttons = qsa(".tab-button-curso");
        if (!tabs.length || !buttons.length) return;

        /* ===== MOBILE ===== */
        if (window.innerWidth < 801) {
            showTab = function (index = 0) {
                if (!tabs[index]) return;

                tabs.forEach((t) => (t.style.display = "none"));
                buttons.forEach((b) => b.classList.remove("active"));

                tabs[index].style.display = "block";
                buttons[index].classList.add("active");
                currentTab = index;

                // ----- SUBTABS (MOBILE) -----
                qsa(".button-container-subtabs", tabs[index]).forEach((container) => {
                    const btns = qsa(".toggle-button-subtabs", container);
                    btns.forEach((b, i) => b.classList.toggle("active-button-subtabs", i === 0));
                });

                qsa(".content-subtabs", tabs[index]).forEach((c, i) =>
                    c.classList.toggle("active-content-subtabs", i === 0)
                );

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
                buttons[index].classList.add("active");
                currentTab = index;

                // ----- SUBTABS (DESKTOP) -----
                qsa(".button-container-subtabs").forEach((c) => {
                    qsa(".toggle-button-subtabs", c).forEach((b, i) =>
                        b.classList.toggle(
                            "active-button-subtabs",
                            i === 0 && c.closest(".tab-content-curso") === tabs[index]
                        )
                    );
                });

                qsa(".content-subtabs").forEach((c) => c.classList.remove("active-content-subtabs"));

                qs(".content-subtabs", tabs[index])?.classList.add("active-content-subtabs");

                updateButtonVisibility();
            };
        }
    }

    /* =========================
     PREV / NEXT
  ========================= */
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

    /* =========================
     INIT SEGURO
  ========================= */
    document.addEventListener("DOMContentLoaded", () => {
        /* console.log("JS Moodle móvil + desktop listo");
         */
        initPadres();
        buildShowTab();

        qsa(".tab-button-curso").forEach((b, i) => b.addEventListener("click", () => showTab(i)));

        showTab(0);
    });

    window.addEventListener("resize", () => {
        buildShowTab();
        showTab(currentTab);
    });
})();
