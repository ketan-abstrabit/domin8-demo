/**
 * DOMIN8 reporting — the client's Run now button.
 *
 * A standalone Apps Script deployed as a web app. The client opens one URL,
 * presses one button, and the GitHub Actions workflow starts within seconds.
 * They never see GitHub, never see the code, and never handle a credential.
 *
 * WHY THE SCRIPT OWNS THE TOKEN
 * -----------------------------
 * Triggering a workflow needs a GitHub token. If this script were bound to a
 * Sheet the client can edit, they could open Extensions > Apps Script and read
 * it. So this is a STANDALONE script, owned by Abstrabit, deployed to
 * "Execute as: me". Every function below runs on the server under the owner's
 * identity; google.script.run never ships the token to the browser. The client
 * gets the button, not the key.
 *
 * SETUP  (see README.md in this folder for the full walk-through)
 *   Project Settings > Script Properties:
 *     GITHUB_REPO     ketan-abstrabit/domin8-demo
 *     GITHUB_TOKEN    fine-grained PAT — Contents: read+write, Actions: read
 *     DRIVE_ROOT_ID   the shared-drive folder id
 *   Then run selftest() once from the editor and read the log.
 */

/**
 * Bump this whenever Code.gs or Index.html changes.
 *
 * Editing the files does not change the live web app — only Deploy > Manage
 * deployments > New version does. That distinction has cost real debugging
 * time twice, with no way to tell a stale deployment from a broken one. The
 * page prints this in its footer and selftest reports it, so "which code is
 * actually running" is a five-second question instead of an argument.
 */
var BUILD = '2026-09-01c  (fetch button)';

var EVENT_TYPE = 'run-report';
var COOLDOWN_SECONDS = 120;
var API = 'https://api.github.com';


/** What the deployed code can do. The page uses this to prove it is current. */
function buildInfo() {
  return {
    build: BUILD,
    has_fetch: (typeof triggerFetch === 'function'),
    has_status: (typeof fetchStatus === 'function')
  };
}


// ---------------------------------------------------------------------------
// configuration
// ---------------------------------------------------------------------------

function prop_(key, required) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  v = v ? v.trim() : '';
  if (!v && required) {
    throw new Error('Script property ' + key + ' is not set. ' +
                    'Project Settings > Script Properties.');
  }
  return v;
}

function ghHeaders_() {
  return {
    Authorization: 'Bearer ' + prop_('GITHUB_TOKEN', true),
    Accept: 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28'
  };
}

/** Every GitHub call goes through here so failures explain themselves. */
function gh_(method, path, payload) {
  var opts = {
    method: method,
    headers: ghHeaders_(),
    muteHttpExceptions: true,
    contentType: 'application/json'
  };
  if (payload) opts.payload = JSON.stringify(payload);

  var res = UrlFetchApp.fetch(API + path, opts);
  var code = res.getResponseCode();
  var body = res.getContentText();

  if (code >= 200 && code < 300) {
    return body ? JSON.parse(body) : {};
  }

  // The setup mistakes here are all token-shaped, and GitHub's own message is
  // more accurate than anything guessed from documentation. Surface it.
  var detail = '';
  try { detail = JSON.parse(body).message || ''; } catch (e) { detail = body; }

  if (code === 401) {
    throw new Error('GitHub rejected the token (401). It is wrong, expired, ' +
                    'or was revoked. Regenerate it and update GITHUB_TOKEN.');
  }
  if (code === 403 || code === 404) {
    throw new Error(
      'GitHub returned ' + code + ': ' + detail + '\n\n' +
      'For a fine-grained token this almost always means a missing ' +
      'permission or the repo not being selected on the token. Triggering ' +
      'needs Contents: read and write; reading run status needs Actions: ' +
      'read. Check both, and that the token lists this exact repository.');
  }
  throw new Error('GitHub ' + code + ': ' + detail);
}


// ---------------------------------------------------------------------------
// the button
// ---------------------------------------------------------------------------

/**
 * Start a run. Called from the page via google.script.run.
 *
 * Returns a plain object rather than throwing, so the page can show a sentence
 * instead of a stack trace.
 */
function triggerRun(force) {
  var cache = CacheService.getScriptCache();
  if (cache.get('cooldown')) {
    return {
      ok: false,
      message: 'A run was started less than two minutes ago. Give it a ' +
               'moment — pressing again would only queue a duplicate.'
    };
  }

  var who = '';
  try { who = Session.getActiveUser().getEmail() || ''; } catch (e) { who = ''; }

  try {
    gh_('post', '/repos/' + prop_('GITHUB_REPO', true) + '/dispatches', {
      event_type: EVENT_TYPE,
      client_payload: { force: force === true, requested_by: who }
    });
  } catch (err) {
    return { ok: false, message: String(err.message || err) };
  }

  cache.put('cooldown', '1', COOLDOWN_SECONDS);
  console.log('run requested by ' + (who || 'unknown') + ' force=' + (force === true));
  return {
    ok: true,
    message: force ? 'Rebuilding from scratch.' : 'Started.'
  };
}


/**
 * Pull the Uniware reports into Drive, and stop.
 *
 * Deliberately a separate button from Run. If Uniware is down, this fails and
 * the report is untouched — the previous pull is still in input/uniware/ and
 * Run still works. The client can also ignore this button entirely and upload
 * the five exports by hand, exactly as they do today.
 */
function triggerFetch(days) {
  var cache = CacheService.getScriptCache();
  if (cache.get('fetch_cooldown')) {
    return {
      ok: false,
      message: 'A pull was started less than two minutes ago. Uniware takes a ' +
               'few minutes to build its exports.'
    };
  }

  var who = '';
  try { who = Session.getActiveUser().getEmail() || ''; } catch (e) { who = ''; }

  try {
    gh_('post', '/repos/' + prop_('GITHUB_REPO', true) + '/dispatches', {
      event_type: 'fetch-uniware',
      client_payload: { days: days || 90, requested_by: who }
    });
  } catch (err) {
    return { ok: false, message: String(err.message || err) };
  }

  cache.put('fetch_cooldown', '1', COOLDOWN_SECONDS);
  console.log('uniware fetch requested by ' + (who || 'unknown'));
  return { ok: true, message: 'Pulling the last ' + (days || 90) + ' days.' };
}

/** How the last Uniware pull went, from _state/last_fetch.json. */
function fetchStatus() {
  try {
    var root = DriveApp.getFolderById(prop_('DRIVE_ROOT_ID', true));
    var st = root.getFoldersByName('_state');
    if (!st.hasNext()) return { known: false };
    var f = st.next().getFilesByName('last_fetch.json');
    if (!f.hasNext()) return { known: false };
    var j = JSON.parse(f.next().getBlob().getDataAsString());
    return {
      known: true, ok: j.ok === true, run_id: j.run_id || '',
      files: j.files || 0, error: j.error || '', by: j.requested_by || ''
    };
  } catch (err) {
    return { known: false };
  }
}


// ---------------------------------------------------------------------------
// status
// ---------------------------------------------------------------------------

/**
 * What is happening right now, and how the last run ended.
 *
 * GitHub knows whether a run is queued or building; STATUS.txt in Drive is the
 * only thing that knows what the run actually concluded. The page needs both,
 * so one call returns both.
 */
function runStatus() {
  var out = { phase: 'unknown', label: 'Unknown', detail: '', status_text: '' };

  try {
    var runs = gh_('get', '/repos/' + prop_('GITHUB_REPO', true) +
                          '/actions/runs?per_page=5');
    var list = (runs && runs.workflow_runs) || [];
    if (list.length) {
      var r = list[0];
      // A fetch and a build look identical here unless the event type is
      // read. Telling someone "building the report" while it is actually
      // talking to Uniware is worse than saying nothing.
      var isFetch = (r.display_title || '').indexOf('fetch-uniware') > -1;
      out.kind = isFetch ? 'fetch' : 'report';
      if (r.status === 'queued' || r.status === 'in_progress' ||
          r.status === 'waiting' || r.status === 'requested') {
        out.phase = 'running';
        out.label = (r.status === 'queued') ? 'Queued'
                  : isFetch ? 'Pulling from Uniware' : 'Building the report';
        out.detail = 'Started ' + ago_(r.run_started_at || r.created_at) +
                     (isFetch ? '. Uniware builds each export on its side, '
                              + 'so this can take several minutes.'
                              : '. This usually takes three to four minutes.');
      } else if (r.conclusion === 'success') {
        out.phase = 'ok';
        out.label = 'Last run finished';
        out.detail = 'Completed ' + ago_(r.updated_at) + '.';
      } else {
        out.phase = 'failed';
        out.label = 'Last run failed';
        out.detail = 'It ended ' + ago_(r.updated_at) + ' (' +
                     (r.conclusion || r.status) + '). Abstrabit has been ' +
                     'emailed. The previous reports are still in output/latest.';
      }
    }
  } catch (err) {
    out.phase = 'error';
    out.label = 'Cannot reach GitHub';
    out.detail = String(err.message || err);
  }

  out.status_text = statusFile_();
  return out;
}

/** STATUS.txt from the Drive folder — the run's own account of itself. */
function statusFile_() {
  try {
    var it = DriveApp.getFolderById(prop_('DRIVE_ROOT_ID', true))
                     .getFilesByName('STATUS.txt');
    if (!it.hasNext()) return '';
    return it.next().getBlob().getDataAsString();
  } catch (err) {
    return '';
  }
}

function ago_(iso) {
  if (!iso) return 'a moment ago';
  var secs = Math.max(0, (new Date().getTime() - new Date(iso).getTime()) / 1000);
  if (secs < 90) return Math.round(secs) + ' seconds ago';
  if (secs < 5400) return Math.round(secs / 60) + ' minutes ago';
  if (secs < 172800) return Math.round(secs / 3600) + ' hours ago';
  return Math.round(secs / 86400) + ' days ago';
}

/** Deep link to the reports, so the page can hand them somewhere to go. */
function folderLinks() {
  var out = { root: '', latest: '' };
  try {
    var root = DriveApp.getFolderById(prop_('DRIVE_ROOT_ID', true));
    out.root = root.getUrl();
    var outs = root.getFoldersByName('output');
    if (outs.hasNext()) {
      var latest = outs.next().getFoldersByName('latest');
      if (latest.hasNext()) out.latest = latest.next().getUrl();
    }
  } catch (err) { /* the page copes with empty links */ }
  return out;
}


// ---------------------------------------------------------------------------
// serving
// ---------------------------------------------------------------------------

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || 'run';
  if (page === 'dashboard') return serveDashboard_();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('DOMIN8 reporting')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Serve the dashboard the pipeline already builds.
 *
 * Drive will not render an HTML file — it offers a download. Reading it here
 * and returning it as the response gives the client a normal web page at a URL
 * that never changes, behind their Google sign-in, always showing the newest
 * run. No hosting, no public bucket, no third party.
 */
function serveDashboard_() {
  var html;
  try {
    var root = DriveApp.getFolderById(prop_('DRIVE_ROOT_ID', true));
    var latest = root.getFoldersByName('output').next()
                     .getFoldersByName('latest').next();
    var files = latest.getFilesByName('dashboard.html');
    if (!files.hasNext()) throw new Error('dashboard.html is not in output/latest yet');
    html = files.next().getBlob().getDataAsString();
  } catch (err) {
    html = '<div style="font:15px/1.6 system-ui;padding:40px;max-width:640px">' +
           '<h2>The dashboard is not available yet</h2><p>' +
           escapeHtml_(String(err.message || err)) +
           '</p><p>Run a report first, then reload this page.</p></div>';
  }
  return HtmlService.createHtmlOutput(html)
    .setTitle('DOMIN8 dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
                  .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function webAppUrl() {
  return ScriptApp.getService().getUrl();
}


// ---------------------------------------------------------------------------
// selftest
//
// Apps Script cannot be tested from outside Google, so this is the substitute:
// run it once from the editor after deploying and it checks every assumption
// the web app makes, reporting exactly which one is wrong. It never triggers a
// run, so it is safe to re-run whenever something looks off.
// ---------------------------------------------------------------------------

function selftest() {
  var lines = ['DOMIN8 web app selftest', '========================'];
  var ok = true;

  function check(name, fn) {
    try {
      var detail = fn();
      lines.push('  PASS  ' + name + (detail ? '  — ' + detail : ''));
    } catch (err) {
      ok = false;
      lines.push('  FAIL  ' + name + '\n          ' + String(err.message || err));
    }
  }

  check('deployed build', function () {
    var b = buildInfo();
    if (!b.has_fetch) {
      throw new Error('this Code.gs has no triggerFetch — the Fetch button ' +
                      'cannot work. Paste the current apps_script/Code.gs.');
    }
    return b.build;
  });

  check('GITHUB_REPO is set', function () {
    var r = prop_('GITHUB_REPO', true);
    if (r.split('/').length !== 2) throw new Error('expected owner/repo, got "' + r + '"');
    return r;
  });

  check('GITHUB_TOKEN is set', function () {
    var t = prop_('GITHUB_TOKEN', true);
    return t.length + ' characters, starts "' + t.substring(0, 7) + '..."';
  });

  check('token can read the repo', function () {
    var r = gh_('get', '/repos/' + prop_('GITHUB_REPO', true));
    return r.full_name + (r.private ? ' (private)' : ' (public)');
  });

  check('token can read workflow runs  [Actions: read]', function () {
    var r = gh_('get', '/repos/' + prop_('GITHUB_REPO', true) +
                       '/actions/runs?per_page=1');
    return (r.total_count || 0) + ' run(s) in history';
  });

  check('workflow accepts repository_dispatch', function () {
    var wfs = gh_('get', '/repos/' + prop_('GITHUB_REPO', true) + '/actions/workflows');
    var names = (wfs.workflows || []).map(function (w) { return w.name; });
    if (!names.length) throw new Error('no workflows found — is the file pushed?');
    return names.join(', ');
  });

  check('DRIVE_ROOT_ID resolves', function () {
    return DriveApp.getFolderById(prop_('DRIVE_ROOT_ID', true)).getName();
  });

  check('output/latest exists', function () {
    var root = DriveApp.getFolderById(prop_('DRIVE_ROOT_ID', true));
    var o = root.getFoldersByName('output');
    if (!o.hasNext()) {
      throw new Error('no output/ folder yet. Expected before the first ' +
                      'successful run, or if it was deleted during cleanup. ' +
                      'Press the button once and this clears.');
    }
    var dupes = 0, first = null;
    while (o.hasNext()) { var f = o.next(); dupes++; if (!first) first = f; }
    if (dupes > 1) {
      throw new Error(dupes + ' folders named "output" exist. Keep the one ' +
                      'containing latest/ and archive/, delete the others.');
    }
    var l = first.getFoldersByName('latest');
    if (!l.hasNext()) throw new Error('no output/latest/ folder yet');
    var n = 0, it = l.next().getFiles();
    while (it.hasNext()) { it.next(); n++; }
    // An empty output/latest is the failure this check exists to catch. It
    // reported "0 published file(s)" as a pass once; the folder existing is
    // not the point, the reports being in it is.
    if (n === 0) {
      throw new Error('output/latest/ is empty — no reports have been ' +
                      'published. Press Run with "rebuild even if nothing ' +
                      'has changed" ticked.');
    }
    return n + ' published file(s)';
  });

  check('STATUS.txt is readable', function () {
    var t = statusFile_();
    if (!t) throw new Error('not found — normal before the first run');
    // Report the RESULT, not just the timestamp. A stale STATUS.txt from a
    // failed run reads identically to a fresh one otherwise, which is exactly
    // the confusion this check exists to prevent.
    var when = '', result = '';
    t.split('\n').forEach(function (line) {
      if (line.indexOf('Last run') === 0) when = line.split(':').slice(1).join(':').trim();
      if (line.indexOf('Result') === 0) result = line.split(':').slice(1).join(':').trim();
    });
    // SKIPPED is a healthy outcome, not a failure: it means the inputs were
    // identical and last cycle's reports still stand. Only a real FAILED
    // deserves to fail this check — the empty-output case is caught above,
    // where it belongs.
    if (result && result.indexOf('FAILED') === 0) {
      throw new Error('the last run failed at ' + when +
                      '. Open STATUS.txt in the Drive folder for the reason.');
    }
    return (result || 'present') + ', ' + when;
  });

  lines.push('');
  lines.push(ok ? 'All checks pass. The button is ready to hand over.'
                : 'Fix the FAILs above, then run selftest() again.');
  lines.push('');
  lines.push('Note: this deliberately does not start a run. Use the button for that.');

  var report = lines.join('\n');
  console.log(report);
  return report;
}
