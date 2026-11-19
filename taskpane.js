/* Outlook 대용량 첨부 (Google Drive) - 작업창 로직 */
(function () {
  'use strict';

  var CFG = window.ADDIN_CONFIG || {};
  var BASE = new URL('.', location.href).href;
  var TOKEN_KEY = 'gdrive_token_v1';
  var DRIVE = 'https://www.googleapis.com/drive/v3';
  var UPLOAD = 'https://www.googleapis.com/upload/drive/v3';
  var CHUNK = Math.max(1, (CFG.CHUNK_MB | 0) || 8) * 1024 * 1024; // 256KB의 배수
  var busy = false;
  var memToken = null;
  var folderIdCache = null;

  function $(id) { return document.getElementById(id); }
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  // ---------- UI ----------
  Office.onReady(function () {
    bindUI();
    refreshAccount();
  });

  function bindUI() {
    $('pick').addEventListener('click', function () { $('file').click(); });
    $('file').addEventListener('change', function (e) {
      var files = Array.prototype.slice.call(e.target.files);
      e.target.value = '';
      handleFiles(files);
    });
    var dz = $('drop');
    ['dragenter', 'dragover'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('over'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove('over'); });
    });
    dz.addEventListener('drop', function (e) {
      handleFiles(Array.prototype.slice.call(e.dataTransfer.files || []));
    });
    $('signin').addEventListener('click', function () {
      hideError();
      getToken(true).then(refreshAccount).catch(showError);
    });
    $('signout').addEventListener('click', signOut);
  }

  function setStatus(t) { $('status').textContent = t || ''; }
  function showError(e) { var el = $('error'); el.textContent = (e && e.message) || String(e); el.hidden = false; }
  function hideError() { $('error').hidden = true; }

  function fmtSize(n) {
    if (n >= 1073741824) return (n / 1073741824).toFixed(2) + ' GB';
    if (n >= 1048576) return (n / 1048576).toFixed(1) + ' MB';
    return Math.max(1, Math.round(n / 1024)) + ' KB';
  }
  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function addRow(file) {
    var li = document.createElement('li');
    li.innerHTML = '<div class="name"></div><div class="meta"></div><div class="bar"><i></i></div>';
    li.querySelector('.name').textContent = file.name;
    var meta = li.querySelector('.meta');
    var bar = li.querySelector('.bar > i');
    meta.textContent = fmtSize(file.size) + ' · 대기 중';
    $('list').insertBefore(li, $('list').firstChild);
    return {
      progress: function (sent) {
        var pct = file.size ? Math.min(100, Math.floor(sent / file.size * 100)) : 100;
        bar.style.width = pct + '%';
        meta.textContent = fmtSize(file.size) + ' · 업로드 ' + pct + '%';
      },
      stage: function (t) { meta.textContent = fmtSize(file.size) + ' · ' + t; },
      done: function (scope) {
        li.classList.add('ok'); bar.style.width = '100%';
        var label = scope === 'anyone' ? '완료 · 링크가 있는 모든 사용자'
          : scope === 'domain' ? '완료 · 사내 계정만 열람 가능(외부 공유 차단됨)'
          : '완료 · 공유 설정 실패(본인만 열람 가능)';
        if (scope !== 'anyone') li.classList.add('warn');
        meta.textContent = fmtSize(file.size) + ' · ' + label;
      },
      fail: function (msg) { li.classList.add('fail'); meta.textContent = '실패: ' + msg; }
    };
  }

  // ---------- 인증 ----------
  function loadToken() {
    if (memToken && memToken.exp > Date.now() + 120000) return memToken;
    try {
      var t = JSON.parse(localStorage.getItem(TOKEN_KEY));
      if (t && t.exp > Date.now() + 120000) { memToken = t; return t; }
    } catch (e) { /* storage 불가 환경 */ }
    return null;
  }
  function saveToken(token, expiresIn) {
    memToken = { token: token, exp: Date.now() + (Number(expiresIn) || 3600) * 1000 };
    try { localStorage.setItem(TOKEN_KEY, JSON.stringify(memToken)); } catch (e) { /* 무시 */ }
    return memToken;
  }
  function clearToken() {
    memToken = null;
    try { localStorage.removeItem(TOKEN_KEY); } catch (e) { /* 무시 */ }
  }

  function getToken(interactive) {
    var t = loadToken();
    if (t) return Promise.resolve(t.token);
    if (!interactive) return Promise.resolve(null);
    return openAuthDialog().then(function (r) { return saveToken(r.access_token, r.expires_in).token; });
  }

  function openAuthDialog() {
    return new Promise(function (resolve, reject) {
      if (!Office.context.ui || !Office.context.ui.displayDialogAsync) {
        return reject(new Error('이 Outlook 버전은 로그인 창을 지원하지 않습니다. Outlook을 업데이트하거나 Outlook 웹에서 사용해 주세요.'));
      }
      Office.context.ui.displayDialogAsync(BASE + 'auth.html', { height: 70, width: 35 }, function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          return reject(new Error('로그인 창을 열지 못했습니다: ' + res.error.message));
        }
        var dlg = res.value, finished = false;
        dlg.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          finished = true;
          dlg.close();
          var m;
          try { m = JSON.parse(arg.message); } catch (e) { return reject(new Error('로그인 응답을 읽지 못했습니다.')); }
          if (m.error) return reject(new Error('Google 로그인 실패: ' + m.error));
          resolve(m);
        });
        dlg.addEventHandler(Office.EventType.DialogEventReceived, function () {
          if (!finished) reject(new Error('로그인 창이 닫혔습니다.'));
        });
      });
    });
  }

  function signOut() {
    var t = loadToken();
    clearToken();
    if (t) {
      fetch('https://oauth2.googleapis.com/revoke?token=' + encodeURIComponent(t.token), { method: 'POST' })
        .catch(function () { /* 무시 */ });
    }
    folderIdCache = null;
    refreshAccount();
  }

  function refreshAccount() {
    return getToken(false).then(function (t) {
      if (!t) return setAccount(null);
      return apiJson(DRIVE + '/about?fields=user(emailAddress)').then(function (a) {
        var email = a.user && a.user.emailAddress;
        if (CFG.HOSTED_DOMAIN && email && !email.toLowerCase().endsWith('@' + CFG.HOSTED_DOMAIN.toLowerCase())) {
          signOut();
          showError('회사 계정(@' + CFG.HOSTED_DOMAIN + ')으로 로그인해 주세요.');
          return;
        }
        setAccount(email || '로그인됨');
      }).catch(function () { setAccount(null); });
    });
  }
  function setAccount(email) {
    $('acct').textContent = email ? email : 'Google 로그인 필요';
    $('signin').hidden = !!email;
    $('signout').hidden = !email;
  }

  // ---------- Drive API ----------
  function api(url, opts, retried) {
    opts = opts || {};
    return getToken(true).then(function (token) {
      var headers = Object.assign({}, opts.headers || {}, { Authorization: 'Bearer ' + token });
      return fetch(url, Object.assign({}, opts, { headers: headers }));
    }).then(function (res) {
      if (res.status === 401 && !retried) { clearToken(); return api(url, opts, true); }
      return res;
    });
  }
  function apiJson(url, opts) {
    return api(url, opts).then(function (r) {
      return r.text().then(function (txt) {
        var j = {};
        try { j = txt ? JSON.parse(txt) : {}; } catch (e) { j = { raw: txt }; }
        if (!r.ok) {
          var err = new Error((j.error && j.error.message) || ('HTTP ' + r.status));
          err.status = r.status;
          throw err;
        }
        return j;
      });
    });
  }

  function getParentId() {
    var name = (CFG.FOLDER_NAME || '').trim();
    if (!name) return Promise.resolve(null); // 내 드라이브 최상위
    if (folderIdCache) return Promise.resolve(folderIdCache);
    var q = "name='" + name.replace(/\\/g, '\\\\').replace(/'/g, "\\'") +
      "' and mimeType='application/vnd.google-apps.folder' and 'root' in parents and trashed=false";
    return apiJson(DRIVE + '/files?fields=files(id)&q=' + encodeURIComponent(q)).then(function (r) {
      if (r.files && r.files.length) return (folderIdCache = r.files[0].id);
      return apiJson(DRIVE + '/files?fields=id', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json; charset=UTF-8' },
        body: JSON.stringify({ name: name, mimeType: 'application/vnd.google-apps.folder', parents: ['root'] })
      }).then(function (f) { return (folderIdCache = f.id); });
    });
  }

  function createSession(file, parentId) {
    var meta = { name: file.name };
    if (file.type) meta.mimeType = file.type;
    if (parentId) meta.parents = [parentId];
    return api(UPLOAD + '/files?uploadType=resumable&fields=id,name,size', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=UTF-8',
        'X-Upload-Content-Type': file.type || 'application/octet-stream'
      },
      body: JSON.stringify(meta)
    }).then(function (r) {
      if (!r.ok) return r.text().then(function (t) { throw new Error('업로드 시작 실패 (' + r.status + ') ' + t.slice(0, 200)); });
      var loc = r.headers.get('Location');
      if (!loc) throw new Error('업로드 세션 주소를 받지 못했습니다.');
      return loc;
    });
  }

  // blob이 null이면 진행 상태 조회(또는 0바이트 파일 완료) 요청
  function putChunk(url, blob, start, total, onProgress) {
    return new Promise(function (resolve) {
      var x = new XMLHttpRequest();
      x.open('PUT', url, true);
      if (blob) x.setRequestHeader('Content-Range', 'bytes ' + start + '-' + (start + blob.size - 1) + '/' + total);
      else x.setRequestHeader('Content-Range', 'bytes */' + total);
      if (onProgress && blob) {
        x.upload.onprogress = function (e) { if (e.lengthComputable) onProgress(start + e.loaded); };
      }
      x.onload = function () { resolve({ status: x.status, range: x.getResponseHeader('Range'), body: x.responseText }); };
      x.onerror = function () { resolve({ status: 0 }); };
      x.send(blob || null);
    });
  }
  function nextOffset(range) { return range ? parseInt(range.split('-')[1], 10) + 1 : 0; }

  function uploadFile(file, parentId, onProgress) {
    var total = file.size;
    return createSession(file, parentId).then(function (session) {
      var offset = 0, fails = 0;
      function step() {
        var blob = total === 0 ? null : file.slice(offset, Math.min(offset + CHUNK, total));
        return putChunk(session, blob, offset, total, onProgress).then(handle);
      }
      function handle(r) {
        if (r.status === 200 || r.status === 201) { onProgress(total); return JSON.parse(r.body); }
        if (r.status === 308) { fails = 0; offset = nextOffset(r.range); onProgress(offset); return step(); }
        if (r.status === 0 || r.status === 429 || r.status >= 500) {
          if (++fails > 6) throw new Error('네트워크 오류로 업로드에 실패했습니다.');
          return sleep(Math.min(32000, 1000 * Math.pow(2, fails))).then(function () {
            return putChunk(session, null, 0, total); // 서버에 받은 위치 확인 후 이어서
          }).then(function (s) {
            if (s.status === 200 || s.status === 201) { onProgress(total); return JSON.parse(s.body); }
            if (s.status === 308) offset = nextOffset(s.range);
            return step();
          });
        }
        if (r.status === 404 || r.status === 410) throw new Error('업로드 세션이 만료되었습니다. 다시 시도해 주세요.');
        throw new Error('업로드 실패 (' + r.status + ') ' + String(r.body || '').slice(0, 200));
      }
      return step();
    });
  }

  function share(id, perm) {
    return apiJson(DRIVE + '/files/' + encodeURIComponent(id) + '/permissions?supportsAllDrives=true&fields=id', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json; charset=UTF-8' },
      body: JSON.stringify(perm)
    });
  }
  function shareFile(id) {
    return share(id, { role: 'reader', type: 'anyone', allowFileDiscovery: false })
      .then(function () { return 'anyone'; })
      .catch(function (e) {
        if (!CFG.HOSTED_DOMAIN) return 'none';
        return share(id, { role: 'reader', type: 'domain', domain: CFG.HOSTED_DOMAIN, allowFileDiscovery: false })
          .then(function () { return 'domain'; })
          .catch(function () { return 'none'; });
      });
  }

  // ---------- 메일 본문 ----------
  function viewUrl(id) { return 'https://drive.google.com/file/d/' + id + '/view?usp=sharing'; }
  function dlUrl(id) { return 'https://drive.google.com/uc?export=download&id=' + id; }

  function buildHtml(items) {
    var allPublic = items.every(function (i) { return i.scope === 'anyone'; });
    var rows = items.map(function (i) {
      var note = i.scope === 'domain' ? ' <span style="color:#9a6700">(사내 계정만 열람 가능)</span>'
        : i.scope === 'none' ? ' <span style="color:#cf222e">(공유 설정 안 됨)</span>' : '';
      return '<div style="margin:3px 0">• <a href="' + viewUrl(i.id) + '">' + esc(i.name) + '</a> ' +
        '<span style="color:#57606a">(' + fmtSize(i.size) + ')</span> &nbsp;' +
        '<a href="' + dlUrl(i.id) + '">[다운로드]</a>' + note + '</div>';
    }).join('');
    return '<div style="margin:12px 0;padding:10px 14px;border:1px solid #d0d7de;border-radius:6px;' +
      'background:#f6f8fa;font-family:\'Malgun Gothic\',\'맑은 고딕\',sans-serif;font-size:10pt">' +
      '<div style="font-weight:bold;margin-bottom:6px">대용량 첨부파일 ' + items.length + '개 (Google Drive)</div>' +
      rows +
      (allPublic ? '<div style="color:#6a737d;font-size:9pt;margin-top:6px">링크를 클릭하면 로그인 없이 내려받을 수 있습니다.</div>' : '') +
      '</div><br>';
  }
  function buildText(items) {
    return '\n[대용량 첨부파일 ' + items.length + '개 (Google Drive)]\n' + items.map(function (i) {
      return '- ' + i.name + ' (' + fmtSize(i.size) + ')\n  보기: ' + viewUrl(i.id) + '\n  다운로드: ' + dlUrl(i.id);
    }).join('\n') + '\n';
  }

  function insertIntoBody(items) {
    return new Promise(function (resolve, reject) {
      var item = Office.context.mailbox.item;
      item.body.getTypeAsync(function (t) {
        var isHtml = t.status === Office.AsyncResultStatus.Succeeded && t.value === Office.CoercionType.Html;
        var data = isHtml ? buildHtml(items) : buildText(items);
        var opt = { coercionType: isHtml ? Office.CoercionType.Html : Office.CoercionType.Text };
        item.body.setSelectedDataAsync(data, opt, function (r) {
          if (r.status === Office.AsyncResultStatus.Succeeded) return resolve();
          item.body.prependAsync(data, opt, function (r2) {
            if (r2.status === Office.AsyncResultStatus.Succeeded) resolve();
            else reject(new Error('본문에 링크를 넣지 못했습니다: ' + r2.error.message));
          });
        });
      });
    });
  }

  // ---------- 메인 흐름 ----------
  function handleFiles(files) {
    if (!files.length) return;
    if (busy) { setStatus('진행 중인 업로드가 끝난 뒤 다시 시도해 주세요.'); return; }
    busy = true; hideError(); setStatus('');
    var results = [];

    getToken(true)
      .then(function () { refreshAccount(); return getParentId(); })
      .then(function (parentId) {
        return files.reduce(function (p, f) {
          return p.then(function () {
            var row = addRow(f);
            return uploadFile(f, parentId, row.progress)
              .then(function (meta) {
                row.stage('공유 설정 중…');
                return shareFile(meta.id).then(function (scope) {
                  results.push({ id: meta.id, name: f.name, size: f.size, scope: scope });
                  row.done(scope);
                });
              })
              .catch(function (e) { row.fail(e.message); });
          });
        }, Promise.resolve());
      })
      .then(function () {
        if (!results.length) return;
        return insertIntoBody(results).then(function () {
          setStatus(results.length + '개 파일의 링크를 본문에 넣었습니다.');
        });
      })
      .catch(showError)
      .then(function () { busy = false; });
  }
})();
