// Hleður gögnum mælaborðs fyrir valinn mánuð, t.d. fjarhagur.html?m=2026-09.
// Án ?m= er nýjasti mánuðurinn í data/manudir.js notaður.
// Notar <script>-hleðslu (ekki fetch) svo þetta virki líka þegar skráin er opnuð beint af disk.
(function () {
  var key = document.currentScript.getAttribute('data-key');
  var listi = window.MANUDIR || [];
  var m = new URLSearchParams(location.search).get('m');
  if (!listi.some(function (x) { return x.id === m; })) {
    m = listi.length ? listi[listi.length - 1].id : '';
  }
  window.VALINN_MANUDUR = m;
  document.write('<script src="../data/' + m + '/' + key + '.js"><\/script>');
})();
