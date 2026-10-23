// Hleður gögnum mælaborðs fyrir valinn mánuð, t.d. fjarhagur.html?m=2026-09.
// Án ?m= er nýjasti mánuðurinn í data/manudir.js notaður.
// Notar <script>-hleðslu (ekki fetch) svo þetta virki líka þegar skráin er opnuð beint af disk.
(function () {
  var key = document.currentScript.getAttribute('data-key');
  var listi = window.MANUDIR || [];
  var q = new URLSearchParams(location.search);
  var m = q.get('m');
  var v = q.get('v');
  if (!listi.some(function (x) { return x.id === m; })) {
    m = listi.length ? listi[listi.length - 1].id : '';
  }
  window.VALINN_MANUDUR = m;
  document.write('<script src="../data/' + m + '/' + key + '.js' + (v ? '?v=' + v : '') + '"><\/script>');
  // Valfrjálsar aukaskrár (t.d. data-extra="skyringar"); mælaborðið þarf að þola að þær vanti fyrir suma mánuði
  var extra = document.currentScript.getAttribute('data-extra');
  if (extra) extra.split(',').forEach(function (k) {
    document.write('<script src="../data/' + m + '/' + k.trim() + '.js' + (v ? '?v=' + v : '') + '"><\/script>');
  });
})();
