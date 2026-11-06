/**
 * 集計・レポート生成ロジック(スプレッドシートに依存しない純粋関数)。
 * Node.js でもテストできるよう、GAS の API はここでは使いません。
 */

/**
 * 指定日と前日の売上を集計する。
 * @param {Array<{date:string, product:string, category:string, quantity:number, amount:number}>} rows
 *   date は 'yyyy-MM-dd' 形式
 * @param {string} targetKey 集計対象日 'yyyy-MM-dd'
 * @param {string} prevKey   比較対象日(前日) 'yyyy-MM-dd'
 */
function buildSummary(rows, targetKey, prevKey) {
  const today = rows.filter(function (r) { return r.date === targetKey; });
  const prev = rows.filter(function (r) { return r.date === prevKey; });

  const total = sumAmount(today);
  const prevTotal = sumAmount(prev);

  // カテゴリ別(前日比つき)
  const catToday = groupSum(today, 'category');
  const catPrev = groupSum(prev, 'category');
  const categoryNames = uniq(Object.keys(catToday).concat(Object.keys(catPrev)));
  const byCategory = categoryNames
    .map(function (name) {
      const t = catToday[name] || 0;
      const p = catPrev[name] || 0;
      return { name: name, total: t, prev: p, diff: t - p, rate: changeRate(t, p) };
    })
    .sort(function (a, b) { return b.total - a.total; });

  // 商品別トップ5
  const prodToday = groupSum(today, 'product');
  const topProducts = Object.keys(prodToday)
    .map(function (name) { return { name: name, total: prodToday[name] }; })
    .sort(function (a, b) { return b.total - a.total; })
    .slice(0, 5);

  return {
    date: targetKey,
    prevDate: prevKey,
    total: total,
    prevTotal: prevTotal,
    diff: total - prevTotal,
    rate: changeRate(total, prevTotal),
    count: today.length,
    quantity: today.reduce(function (s, r) { return s + r.quantity; }, 0),
    byCategory: byCategory,
    topProducts: topProducts,
  };
}

/** 前日比(%)。前日が0のときは null(比較不能) */
function changeRate(current, previous) {
  if (!previous) return null;
  return Math.round(((current - previous) / previous) * 1000) / 10;
}

function sumAmount(rows) {
  return rows.reduce(function (s, r) { return s + r.amount; }, 0);
}

function groupSum(rows, key) {
  const result = {};
  rows.forEach(function (r) {
    const k = r[key] || '(未設定)';
    result[k] = (result[k] || 0) + r.amount;
  });
  return result;
}

function uniq(arr) {
  return arr.filter(function (v, i) { return arr.indexOf(v) === i; });
}

// ---------------------------------------------------------------------------
// 表示用フォーマット
// ---------------------------------------------------------------------------
function formatYen(n) {
  const sign = n < 0 ? '-' : '';
  return sign + '¥' + Math.abs(Math.round(n)).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function formatDiff(n) {
  return (n > 0 ? '+' : '') + formatYen(n);
}

function formatRate(rate) {
  if (rate === null) return '—';
  return (rate > 0 ? '+' : '') + rate.toFixed(1) + '%';
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** メール件名 */
function buildSubject(summary, prefix) {
  return (prefix || '【日次売上レポート】') + summary.date + ' 売上 ' + formatYen(summary.total) +
    '(前日比 ' + formatRate(summary.rate) + ')';
}

/** テキストメール本文(HTML非対応のメーラー向け) */
function buildTextBody(s) {
  const lines = [
    s.date + ' の売上レポート',
    '',
    '売上合計: ' + formatYen(s.total) + '(前日 ' + formatYen(s.prevTotal) + ' / ' + formatDiff(s.diff) + ' / ' + formatRate(s.rate) + ')',
    '取引件数: ' + s.count + ' 件 / 販売数量: ' + s.quantity + ' 個',
    '',
    '■ カテゴリ別',
  ];
  s.byCategory.forEach(function (c) {
    lines.push('・' + c.name + ': ' + formatYen(c.total) + '(前日比 ' + formatRate(c.rate) + ')');
  });
  lines.push('', '■ 売上トップ5');
  s.topProducts.forEach(function (p, i) {
    lines.push((i + 1) + '. ' + p.name + ': ' + formatYen(p.total));
  });
  return lines.join('\n');
}

/** HTMLメール本文 */
function buildHtmlBody(s, sheetUrl) {
  const color = function (n) { return n > 0 ? '#1a7f37' : n < 0 ? '#cf222e' : '#57606a'; };
  const td = 'padding:6px 10px;border-bottom:1px solid #e5e7eb;';
  const tdR = td + 'text-align:right;';

  const catRows = s.byCategory.map(function (c) {
    return '<tr><td style="' + td + '">' + escapeHtml(c.name) + '</td>' +
      '<td style="' + tdR + '">' + formatYen(c.total) + '</td>' +
      '<td style="' + tdR + 'color:' + color(c.diff) + '">' + formatDiff(c.diff) + '</td>' +
      '<td style="' + tdR + 'color:' + color(c.diff) + '">' + formatRate(c.rate) + '</td></tr>';
  }).join('');

  const topRows = s.topProducts.map(function (p, i) {
    return '<tr><td style="' + td + '">' + (i + 1) + '</td>' +
      '<td style="' + td + '">' + escapeHtml(p.name) + '</td>' +
      '<td style="' + tdR + '">' + formatYen(p.total) + '</td></tr>';
  }).join('');

  const th = 'padding:6px 10px;background:#f3f4f6;text-align:left;font-weight:600;';
  return '' +
    '<div style="font-family:sans-serif;color:#1f2328;max-width:560px;">' +
    '<h2 style="margin:0 0 4px;">日次売上レポート</h2>' +
    '<div style="color:#57606a;margin-bottom:16px;">' + s.date + '(比較: ' + s.prevDate + ')</div>' +
    '<div style="border:1px solid #d0d7de;border-radius:8px;padding:16px;margin-bottom:20px;">' +
    '<div style="color:#57606a;font-size:13px;">売上合計</div>' +
    '<div style="font-size:28px;font-weight:700;">' + formatYen(s.total) + '</div>' +
    '<div style="color:' + color(s.diff) + ';font-weight:600;">前日比 ' + formatDiff(s.diff) + '(' + formatRate(s.rate) + ')</div>' +
    '<div style="color:#57606a;font-size:13px;margin-top:8px;">取引 ' + s.count + ' 件 / 数量 ' + s.quantity + ' 個 / 前日売上 ' + formatYen(s.prevTotal) + '</div>' +
    '</div>' +
    '<h3 style="margin:0 0 8px;">カテゴリ別</h3>' +
    '<table style="border-collapse:collapse;width:100%;margin-bottom:20px;font-size:14px;">' +
    '<tr><th style="' + th + '">カテゴリ</th><th style="' + th + 'text-align:right;">売上</th><th style="' + th + 'text-align:right;">前日差</th><th style="' + th + 'text-align:right;">前日比</th></tr>' +
    catRows + '</table>' +
    '<h3 style="margin:0 0 8px;">売上トップ5</h3>' +
    '<table style="border-collapse:collapse;width:100%;margin-bottom:20px;font-size:14px;">' +
    '<tr><th style="' + th + '">#</th><th style="' + th + '">商品</th><th style="' + th + 'text-align:right;">売上</th></tr>' +
    topRows + '</table>' +
    (sheetUrl ? '<a href="' + escapeHtml(sheetUrl) + '" style="color:#0969da;">スプレッドシートを開く</a>' : '') +
    '</div>';
}

// Node.js でのテスト用(GAS では module が存在しないため無視される)
if (typeof module !== 'undefined') {
  module.exports = {
    buildSummary: buildSummary, changeRate: changeRate, formatYen: formatYen, formatDiff: formatDiff,
    formatRate: formatRate, escapeHtml: escapeHtml, buildSubject: buildSubject,
    buildTextBody: buildTextBody, buildHtmlBody: buildHtmlBody,
  };
}
