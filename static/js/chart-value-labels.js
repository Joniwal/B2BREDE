/* Rótulos numéricos sobre barras, colunas e pontos dos gráficos. */
(() => {
  const formatter = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 1 });

  const chartValueLabels = {
    id: "valueLabels",
    afterDatasetsDraw(chart, _args, options) {
      const ctx = chart.ctx;
      ctx.save();
      ctx.font = "600 9px Arial, sans-serif";
      ctx.textAlign = "center";
      ctx.lineJoin = "round";
      ctx.lineWidth = 3;
      ctx.strokeStyle = "rgba(255, 255, 255, 0.9)";
      ctx.fillStyle = options.color || "#39424e";

      chart.data.datasets.forEach((_dataset, datasetIndex) => {
        const meta = chart.getDatasetMeta(datasetIndex);
        if (meta.hidden) return;

        meta.data.forEach((element, index) => {
          const parsed = meta.controller.getParsed(index);
          const valueAxis = meta.vScale?.axis || "y";
          const value = Number(parsed[valueAxis]);
          if (!Number.isFinite(value) || (value === 0 && !options.showZero)) return;

          const isPositive = value >= 0;
          const isBar = meta.type === "bar";
          const isHorizontalBar = isBar && valueAxis === "x";
          let x = element.x;
          let y = element.y;
          let textAlign = "center";
          let baseline = isPositive ? "bottom" : "top";

          if (isHorizontalBar) {
            // Nas barras horizontais, o valor fica além da ponta direita/esquerda.
            x += isPositive ? 7 : -7;
            textAlign = isPositive ? "left" : "right";
            baseline = "middle";
          } else {
            // Colunas e pontos: coloca o rótulo acima/abaixo da extremidade,
            // nunca no centro ou dentro do elemento gráfico.
            y += isPositive ? -7 : 7;
          }

          const label = formatter.format(value);
          ctx.textAlign = textAlign;
          ctx.textBaseline = baseline;
          ctx.strokeText(label, x, y);
          ctx.fillText(label, x, y);
        });
      });

      ctx.restore();
    },
  };

  Chart.register(chartValueLabels);
})();
