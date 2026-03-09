(function () {
  "use strict";

  document.addEventListener("DOMContentLoaded", function () {

    const frames = document.querySelectorAll('iframe[src*="youtube.com/embed"]');

    if (!frames.length) return;

    frames.forEach((frame) => {

      try {

        const currentSrc = frame.getAttribute("src");
        if (!currentSrc) return;

        const url = new URL(currentSrc);

        // parámetros de tracking que conviene eliminar
        const removeParams = [
          "si",
          "feature",
          "utm_source",
          "utm_medium",
          "utm_campaign"
        ];

        removeParams.forEach(param => url.searchParams.delete(param));

        const cleanedSrc =
          url.origin +
          url.pathname +
          (url.searchParams.toString()
            ? "?" + url.searchParams.toString()
            : "");

        // evitar recargar iframe si no cambió
        if (cleanedSrc !== currentSrc) {
          frame.src = cleanedSrc;
        }

        // asegurar permisos correctos
        frame.setAttribute(
          "allow",
          "accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture"
        );

      } catch (err) {
        console.warn("YouTube iframe cleanup error:", err);
      }

    });

  });

})();
