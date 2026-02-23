self.onmessage = (event) => {
  const { candidates = [], target = 0, tolerance = 0.01, timeLimit = 3000 } = event.data || {};
  const start = Date.now();
  const out = [];

  const cents = candidates.map((c) => ({ ...c, cent: Math.round(Number(c.amount || 0) * 100) }));
  const targetCent = Math.round(Number(target || 0) * 100);
  const tolCent = Math.round(Number(tolerance || 0) * 100);

  let interrupted = false;

  function pushResult(picks, total) {
    out.push({
      id: Math.random().toString(36).slice(2, 10),
      transactionIds: picks.map((p) => p.id),
      total: total / 100,
      delta: (total - targetCent) / 100,
    });
  }

  function dfs(idx, picks, total) {
    if (out.length >= 50 || interrupted) return;
    if (Date.now() - start > timeLimit) {
      interrupted = true;
      return;
    }
    if (Math.abs(total - targetCent) <= tolCent && picks.length > 0) {
      pushResult(picks, total);
      if (out.length >= 50) return;
    }
    if (idx >= cents.length) return;

    for (let i = idx; i < cents.length; i += 1) {
      picks.push(cents[i]);
      dfs(i + 1, picks, total + cents[i].cent);
      picks.pop();
      if (out.length >= 50 || interrupted) return;
    }
  }

  dfs(0, [], 0);

  self.postMessage({
    results: out,
    interrupted,
    elapsed: Date.now() - start,
  });
};
