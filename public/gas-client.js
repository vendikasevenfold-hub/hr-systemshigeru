/**
 * Pengganti google.script.run di browser. index.html tetap memakai pola yang sama:
 *   google.script.run.withSuccessHandler(fn).withFailureHandler(fn).namaFungsi(arg1, arg2)
 * Setiap panggilan dikirim ke server lewat POST /api/run/namaFungsi, membawa token sesi
 * (header X-HR-Token) yang disimpan index.html di sessionStorage setelah login.
 */
(function () {
  var PUBLIC_FUNCTIONS = { loginUser: 1, validateSession: 1, logoutSession: 1 };
  var sesiBerakhirDitangani = false;

  function token() {
    try { return sessionStorage.getItem('hr_system_token') || ''; } catch (e) { return ''; }
  }

  // Token kedaluwarsa (mis. server di-restart): hapus sesi tersimpan lalu kembali ke halaman login.
  function sesiBerakhir() {
    if (sesiBerakhirDitangani) return;
    sesiBerakhirDitangani = true;
    try { sessionStorage.removeItem('hr_system_user'); sessionStorage.removeItem('hr_system_token'); } catch (e) { /* abaikan */ }
    alert('Sesi login Anda sudah berakhir. Silakan login ulang.');
    window.location.reload();
  }

  function makeRunner(opts) {
    return new Proxy({}, {
      get: function (_t, prop) {
        if (prop === 'withSuccessHandler') return function (fn) { return makeRunner(Object.assign({}, opts, { ok: fn })); };
        if (prop === 'withFailureHandler') return function (fn) { return makeRunner(Object.assign({}, opts, { fail: fn })); };
        if (prop === 'withUserObject') return function (obj) { return makeRunner(Object.assign({}, opts, { user: obj, hasUser: true })); };
        if (typeof prop !== 'string' || prop === 'then') return undefined;
        return function () {
          var args = Array.prototype.slice.call(arguments);
          var tok = token();
          if (!tok && !PUBLIC_FUNCTIONS[prop]) {
            // Belum login: jangan panggil server (pasti ditolak).
            if (opts.fail) opts.fail(new Error('Belum login.'), opts.user);
            return;
          }
          fetch('/api/run/' + encodeURIComponent(prop), {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'X-HR-Token': tok },
            body: JSON.stringify({ args: args })
          })
            .then(function (res) {
              return res.json().catch(function () { return { ok: false, error: 'Respons server tidak valid (HTTP ' + res.status + ').' }; });
            })
            .then(function (data) {
              if (data && data.ok) {
                if (opts.ok) opts.ok(data.result, opts.user);
              } else if (data && data.sessionExpired) {
                sesiBerakhir();
              } else {
                var err = new Error((data && data.error) || 'Terjadi kesalahan di server.');
                if (opts.fail) opts.fail(err, opts.user); else console.error(err);
              }
            }, function (netErr) {
              var err = new Error('Tidak dapat terhubung ke server: ' + netErr.message);
              if (opts.fail) opts.fail(err, opts.user); else console.error(err);
            });
        };
      }
    });
  }
  window.google = window.google || {};
  window.google.script = window.google.script || {};
  Object.defineProperty(window.google.script, 'run', { get: function () { return makeRunner({}); } });
})();
