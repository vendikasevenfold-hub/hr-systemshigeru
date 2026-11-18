/**
 * Pengganti google.script.run di browser. index.html tetap memakai pola yang sama:
 *   google.script.run.withSuccessHandler(fn).withFailureHandler(fn).namaFungsi(arg1, arg2)
 * Setiap panggilan dikirim ke server lewat POST /api/run/namaFungsi.
 */
(function () {
  function makeRunner(opts) {
    return new Proxy({}, {
      get: function (_t, prop) {
        if (prop === 'withSuccessHandler') return function (fn) { return makeRunner(Object.assign({}, opts, { ok: fn })); };
        if (prop === 'withFailureHandler') return function (fn) { return makeRunner(Object.assign({}, opts, { fail: fn })); };
        if (prop === 'withUserObject') return function (obj) { return makeRunner(Object.assign({}, opts, { user: obj, hasUser: true })); };
        if (typeof prop !== 'string' || prop === 'then') return undefined;
        return function () {
          var args = Array.prototype.slice.call(arguments);
          fetch('/api/run/' + encodeURIComponent(prop), {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ args: args })
          })
            .then(function (res) {
              return res.json().catch(function () { return { ok: false, error: 'Respons server tidak valid (HTTP ' + res.status + ').' }; });
            })
            .then(function (data) {
              if (data && data.ok) {
                if (opts.ok) opts.ok(data.result, opts.user);
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
