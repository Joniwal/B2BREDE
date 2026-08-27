/* Rótulos fixos: carregue após Chart.js e antes de main.js / analises.js. */
(() => {
  const formatter = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 1 });

  const chartValueLabels = {
    id: "valueLabels",
    defaults: {
      color: "#333",
      backgroundColor: "rgba(226, 226, 226, 0.95)",
      fontSize: 11,
      offset: 4,
      showZero: true,
    },
    afterDatasetsDraw(chart, _args, options) {
      if (!chart.chartArea) return;
      const ctx = chart.ctx;
      const fontSize = options.fontSize;
      const labelHeight = fontSize + 6;
      ctx.save();
      ctx.font = `600 ${fontSize}px 'Segoe UI', Arial, sans-serif`;
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";

      chart.data.datasets.forEach((_dataset, datasetIndex) => {
        if (!chart.isDatasetVisible(datasetIndex)) return;
        const meta = chart.getDatasetMeta(datasetIndex);
        if (meta.type !== "bar" && meta.type !== "line") return;

        meta.data.forEach((element, index) => {
          if (element.skip || element.hidden || !chart.getDataVisibility(index)) return;
          const parsed = meta.controller.getParsed(index);
          const valueAxis = meta.vScale?.axis || "y";
          const rawValue = parsed?.[valueAxis];
          if (rawValue === null || rawValue === undefined) return;
          const value = Number(rawValue);
          if (!Number.isFinite(value) || (value === 0 && !options.showZero)) return;
          if (!Number.isFinite(element.x) || !Number.isFinite(element.y)) return;

          const isPositive = value >= 0;
          const isHorizontalBar = meta.type === "bar" && valueAxis === "x";
          const label = formatter.format(value);
          const labelWidth = ctx.measureText(label).width + 8;
          let x = element.x - labelWidth / 2;
          let y = isPositive
            ? element.y - options.offset - labelHeight
            : element.y + options.offset;

          if (isHorizontalBar) {
            x = isPositive
              ? element.x + options.offset
              : element.x - options.offset - labelWidth;
            y = element.y - labelHeight / 2;
          }

          // Mantém a caixa inteira dentro do canvas, inclusive nos extremos.
          x = Math.max(2, Math.min(x, chart.width - labelWidth - 2));
          y = Math.max(2, Math.min(y, chart.height - labelHeight - 2));
          ctx.fillStyle = options.backgroundColor;
          ctx.fillRect(x, y, labelWidth, labelHeight);
          ctx.fillStyle = options.color;
          ctx.fillText(label, x + labelWidth / 2, y + labelHeight / 2);
        });
      });

      ctx.restore();
    },
  };

  Chart.register(chartValueLabels);
})();
