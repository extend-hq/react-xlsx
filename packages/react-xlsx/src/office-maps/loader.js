async function embeddedModule(loadPayload) {
  const { default: base64 } = await loadPayload();
  const binary = atob(base64);
  const compressed = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) compressed[i] = binary.charCodeAt(i);
  const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream("gzip"));
  return new Response(stream, { headers: { "Content-Type": "application/wasm" } });
}

export function createLoader(init, renderChartSvg, loadPayload) {
  const renderer = {
    renderChart(xml, options = {}) {
      return renderChartSvg(xml, JSON.stringify(options)) ?? null;
    },
  };

  let pending;

  return function loadOfficeCharts(source) {
    if (!pending) {
      const input = source ?? (typeof DecompressionStream === "function" ? embeddedModule(loadPayload) : undefined);
      pending = Promise.resolve(input)
        .then((module) => init(module === undefined ? undefined : { module_or_path: module }))
        .then(() => renderer)
        .catch((error) => {
          pending = undefined;
          throw error;
        });
    }
    return pending;
  };
}
