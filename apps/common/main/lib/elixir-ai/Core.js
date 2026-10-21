/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Elixir AI core, editor independent:
 *  - Settings: endpoint / key / model, stored per browser (localStorage). Nothing configured => no network call.
 *  - Journal: local log of every AI call (IndexedDB, in-memory fallback), viewable and exportable.
 *  - Connector: one OpenAI-compatible client (chat completions, tools, streaming SSE) using fetch.
 *  - Orchestrator: the tool loop (capped rounds, validated arguments, structured errors).
 *
 * The API key is only ever sent in the Authorization header of requests to the configured endpoint.
 * It is never logged, never written to the journal and never part of exported data.
 */

define([], function () {
    'use strict';

    var ElixirAI = (window.Common = window.Common || {}).ElixirAI = window.Common.ElixirAI || {};

    // ---------------------------------------------------------------------------------------------
    // Settings
    // ---------------------------------------------------------------------------------------------

    var SETTINGS_KEY = 'elixir-ai-settings';

    // Indicative prices in USD per million tokens (input, output), to be checked against the provider's
    // price list; overridable through the stored settings ("prices"). Used only for the journal estimate.
    var DEFAULT_PRICES = {
        'gpt-4.1-mini': [0.40, 1.60],
        'gpt-5.4-mini': [0.75, 4.50],
        'gpt-5.4': [2.50, 15.00]
    };

    var DEFAULTS = {
        endpoint: '',       // e.g. https://api.openai.com/v1 ; empty = not configured
        apiKey: '',
        model: 'gpt-4.1-mini',
        extraParams: '',    // optional JSON merged into the request body, e.g. {"reasoning_effort":"low"}
        maxRounds: 10,
        mode: 'scaffold',   // "scaffold" (plan, analysis tools, playbooks, verification) or "direct" (plain tool loop)
        earlyStop: true,    // end without a closing model call when a round only made successful changes
        verify: 'always',   // scaffolded mode: verify every change ("always") or only compound ones ("compound")
        sampling: 'deterministic', // "deterministic": temperature 0 where the model accepts it; "default": provider defaults
        prices: DEFAULT_PRICES
    };

    // Check raw authorities before URL normalization (127.1 and encoded hosts are not literals).
    function endpointAllowed(value, managed) {
        if (managed && /^\/elixir-ai\/v1\/?$/.test(value)) return true;
        if (typeof value !== 'string' || value.length > 2048 || /[\u0000-\u0020\u007f\\?#]/.test(value)) return false;
        var raw = /^https?:\/\/([^/]+)(?:\/|$)/i.exec(value), url;
        if (!raw) return false;
        try { url = new URL(value); } catch (e) { return false; }
        if (raw[1].indexOf('@') >= 0 || /:$/.test(raw[1]) || url.username || url.password || url.port === '0') return false;
        if (url.protocol === 'https:') return true;
        return url.protocol === 'http:' && /^(?:localhost|127\.0\.0\.1|\[::1\])(?::[0-9]+)?$/i.test(raw[1]);
    }

    function readJSON(key) {
        try { return JSON.parse(window.localStorage.getItem(key) || 'null'); } catch (e) { return null; }
    }

    ElixirAI.Settings = {
        defaults: DEFAULTS,
        // Server-managed mode: the document server may publish /elixir-ai/config.json
        // (endpoint on the same origin, key added server-side). Settings saved in the
        // browser with an endpoint take precedence.
        managed: null,
        managedReady: null,
        managedLoaded: false,
        // a configuration saved in this browser (own service, own key)
        hasUserConfig: function () { if (window.eoDesktopAI) return !!window.eoDesktopAI.get().endpoint; var u = readJSON(SETTINGS_KEY); return !!(u && u.endpoint); },
        // /elixir-ai/config.json: {"enabled":false} (or no file) = AI off on the server; otherwise the
        // same-origin endpoint, the model and, for information, the host the server forwards to.
        // Not requested at all when the user has their own configuration, which takes precedence.
        loadManaged: function (force) {
            var me = this;
            if (window.eoDesktopAI) return window.eoDesktopAI.ready();
            if (force) me.managedReady = null;
            if (!me.managedReady) {
                if (me.hasUserConfig()) { me.managedLoaded = true; return Promise.resolve(null); }
                me.managedReady = (window.fetch ? fetch('/elixir-ai/config.json', { cache: 'no-store', credentials: 'omit', redirect: 'error' }) : Promise.reject())
                    .then(function (r) { return r.ok ? r.json() : null; })
                    .then(function (c) {
                        me.managed = (c && c.enabled !== false && typeof c.endpoint === 'string' && endpointAllowed(c.endpoint, true) && c.endpoint.charAt(0) === '/') ? c : null;
                        me.managedLoaded = true;
                        return me.managed;
                    })
                    .catch(function () { me.managedLoaded = true; return null; });
            }
            return me.managedReady;
        },
        get: function () {
            var s = window.eoDesktopAI ? window.eoDesktopAI.get() : readJSON(SETTINGS_KEY) || {};
            if (!s.endpoint && this.managed && !this.hasUserConfig()) {
                s = Object.assign({}, s, { endpoint: this.managed.endpoint, apiKey: '', managed: true });
                if (!s.model && this.managed.model) s.model = this.managed.model;
            }
            var out = {};
            for (var k in DEFAULTS) out[k] = (s[k] !== undefined && s[k] !== null) ? s[k] : DEFAULTS[k];
            out.managed = !!s.managed;
            if (window.eoDesktopAI) out.hasKey = !!s.hasKey;
            out.prices = Object.assign({}, DEFAULT_PRICES, s.prices || {});
            return out;
        },
        endpointAllowed: endpointAllowed,
        save: function (s) {
            if (!s || !endpointAllowed(s.endpoint, s.managed)) return false;
            if (window.eoDesktopAI) return window.eoDesktopAI.configure(s);
            try {
                window.localStorage.setItem(SETTINGS_KEY, JSON.stringify(s));
                return true;
            } catch (e) { return false; }
        },
        clear: function () { if (window.eoDesktopAI) return window.eoDesktopAI.clear(); try { window.localStorage.removeItem(SETTINGS_KEY); } catch (e) {} },
        isConfigured: function (s) {
            s = s || this.get();
            if (window.eoDesktopAI) return window.eoDesktopAI.configured(s);
            return !!(endpointAllowed(s.endpoint, s.managed) && s.model);
        },
        // What may be shown or journaled: never the key.
        publicView: function (s) {
            var host = '';
            try { host = s.managed ? 'server' : new URL(s.endpoint).host; } catch (e) {}
            return { endpointHost: host, model: s.model, hasKey: !!(s.apiKey || (window.eoDesktopAI && s.hasKey)) };
        },
        estimateCost: function (s, model, usage) {
            var p = (s.prices || {})[model];
            if (!p || !usage) return null;
            return ((usage.prompt_tokens || 0) * p[0] + (usage.completion_tokens || 0) * p[1]) / 1e6;
        }
    };

    // ---------------------------------------------------------------------------------------------
    // Journal (IndexedDB)
    // ---------------------------------------------------------------------------------------------

    var DB_NAME = 'elixir-ai', STORE = 'calls';
    var memory = [];
    var dbPromise = null;

    function openDb() {
        if (dbPromise) return dbPromise;
        dbPromise = new Promise(function (resolve) {
            try {
                var req = window.indexedDB.open(DB_NAME, 1);
                req.onupgradeneeded = function () { req.result.createObjectStore(STORE, { keyPath: 'id', autoIncrement: true }); };
                req.onsuccess = function () { resolve(req.result); };
                req.onerror = function () { resolve(null); };
            } catch (e) { resolve(null); }
        });
        return dbPromise;
    }

    function tx(mode, fn) {
        return openDb().then(function (db) {
            if (!db) return fn(null);
            return new Promise(function (resolve, reject) {
                var t = db.transaction(STORE, mode), store = t.objectStore(STORE), result;
                Promise.resolve(fn(store)).then(function (r) { result = r; });
                t.oncomplete = function () { resolve(result); };
                t.onerror = function () { reject(t.error); };
            });
        });
    }

    function reqP(req) {
        return new Promise(function (resolve, reject) {
            req.onsuccess = function () { resolve(req.result); };
            req.onerror = function () { reject(req.error); };
        });
    }

    ElixirAI.Journal = {
        add: function (entry) {
            return tx('readwrite', function (store) {
                if (!store) { entry.id = memory.length + 1; memory.push(entry); return entry.id; }
                return reqP(store.add(entry));
            }).catch(function () { return null; });
        },
        update: function (entry) {
            return tx('readwrite', function (store) {
                if (!store) return entry.id;
                return reqP(store.put(entry));
            }).catch(function () { return null; });
        },
        list: function () {
            return tx('readonly', function (store) {
                if (!store) return memory.slice();
                return reqP(store.getAll());
            }).then(function (all) { return (all || []).sort(function (a, b) { return b.id - a.id; }); })
              .catch(function () { return memory.slice().reverse(); });
        },
        clear: function () {
            memory = [];
            return tx('readwrite', function (store) { return store ? reqP(store.clear()) : null; }).catch(function () {});
        },
        exportJSON: function () {
            return this.list().then(function (all) {
                return JSON.stringify({ exported: new Date().toISOString(), entries: all }, null, 2);
            });
        }
    };

    // ---------------------------------------------------------------------------------------------
    // Connector: OpenAI-compatible chat completions with tools, streamed (SSE)
    // ---------------------------------------------------------------------------------------------

    function ConnectorError(message, status) {
        var e = new Error(message);
        e.status = status;
        return e;
    }

    // Models that only accept the default temperature (OpenAI reasoning models)
    var REASONING_MODEL = /^(o\d|gpt-5)/i;
    // Models that refused a temperature in this session
    var SAMPLING_REFUSED = {};

    ElixirAI.Connector = {
        /**
         * opts: {settings, messages, tools, signal, onDelta(text)}
         * resolves {message:{role, content, tool_calls}, usage, finish_reason, ms, ttfbMs}
         */
        chat: function (opts) {
            var s = opts.settings;
            if (!ElixirAI.Settings.isConfigured(s))
                return Promise.reject(ConnectorError('Elixir AI is not configured.', 0));
            var url = s.endpoint.replace(/\/+$/, '') + '/chat/completions';
            var body = {
                model: s.model,
                messages: opts.messages,
                stream: true,
                stream_options: { include_usage: true }
            };
            if (opts.tools && opts.tools.length) { body.tools = opts.tools; body.tool_choice = opts.toolChoice || 'auto'; }
            // Deterministic sampling: the same request gives (nearly always) the same plan. Without it the
            // provider samples at temperature 1 and a task that passes 3/3 can fail 4 times out of 9.
            // Reasoning models refuse temperature: they are skipped, and a 400 naming it drops it for that
            // model for the rest of the session. No `seed`: with it OpenAI answered some requests in 45-55 s
            // instead of 1.5 s (docs/ai.md §6.6), for no gain over temperature 0.
            if (s.sampling !== 'default' && !REASONING_MODEL.test(s.model) && !SAMPLING_REFUSED[s.model]) body.temperature = 0;
            if (s.extraParams) {
                try {
                    var extra = JSON.parse(s.extraParams);
                    delete extra.model;
                    delete extra.messages;
                    Object.assign(body, extra);
                } catch (e) {}
            }
            var headers = { 'Content-Type': 'application/json' };
            if (s.apiKey) headers.Authorization = 'Bearer ' + s.apiKey;

            var t0 = performance.now(), ttfb = null;
            var send = function (b) {
                return (window.eoDesktopAI ? window.eoDesktopAI.fetch : fetch)(url, { method: 'POST', headers: headers, body: JSON.stringify(b), signal: opts.signal, credentials: 'omit', referrerPolicy: 'no-referrer', redirect: 'error' });
            };
            var retried = false;
            var retry400 = function (resp) {
                if (resp.status !== 400 || retried) return resp;
                return resp.clone().text().then(function (t) {
                    // some compatible servers reject stream_options or a sampling parameter: retry once without it
                    var drop = ['stream_options', 'temperature'].filter(function (k) { return body[k] !== undefined && new RegExp(k).test(t); });
                    if (!drop.length) return resp;
                    retried = true;
                    drop.forEach(function (k) {
                        delete body[k];
                        if (k === 'temperature') SAMPLING_REFUSED[s.model] = true;
                    });
                    return send(body);
                });
            };
            return send(body).then(retry400).then(function (resp) {
                if (!resp.ok) {
                    return resp.text().then(function (t) {
                        var msg = t;
                        try { msg = (JSON.parse(t).error || {}).message || t; } catch (e) {}
                        throw ConnectorError('HTTP ' + resp.status + ': ' + String(msg).slice(0, 300), resp.status);
                    });
                }
                var ctype = resp.headers.get('content-type') || '';
                if (ctype.indexOf('text/event-stream') < 0 || !resp.body) {
                    // non-streaming answer
                    return resp.json().then(function (j) {
                        var ch = (j.choices || [])[0] || {};
                        return { message: ch.message || { role: 'assistant', content: '' }, usage: j.usage || null, finish_reason: ch.finish_reason, ms: performance.now() - t0, ttfbMs: performance.now() - t0 };
                    });
                }
                var reader = resp.body.getReader(), decoder = new TextDecoder(), buf = '';
                var content = '', calls = [], usage = null, finish = null;
                var handle = function (data) {
                    if (data === '[DONE]') return;
                    var j;
                    try { j = JSON.parse(data); } catch (e) { return; }
                    if (j.usage) usage = j.usage;
                    var ch = (j.choices || [])[0];
                    if (!ch) return;
                    if (ch.finish_reason) finish = ch.finish_reason;
                    var d = ch.delta || {};
                    if (d.content) {
                        content += d.content;
                        opts.onDelta && opts.onDelta(d.content, content);
                    }
                    (d.tool_calls || []).forEach(function (tc) {
                        var i = tc.index !== undefined ? tc.index : calls.length;
                        var c = calls[i] = calls[i] || { id: '', type: 'function', 'function': { name: '', arguments: '' } };
                        if (tc.id) c.id = tc.id;
                        if (tc['function']) {
                            if (tc['function'].name) c['function'].name += tc['function'].name;
                            if (tc['function'].arguments) c['function'].arguments += tc['function'].arguments;
                        }
                    });
                };
                var pump = function () {
                    return reader.read().then(function (r) {
                        if (ttfb === null) ttfb = performance.now() - t0;
                        if (r.done) return;
                        buf += decoder.decode(r.value, { stream: true });
                        var lines = buf.split(/\r?\n/);
                        buf = lines.pop();
                        lines.forEach(function (line) {
                            if (line.indexOf('data:') === 0) handle(line.slice(5).trim());
                        });
                        return pump();
                    });
                };
                return pump().then(function () {
                    if (buf.indexOf('data:') === 0) handle(buf.slice(5).trim());
                    var message = { role: 'assistant', content: content || null };
                    calls = calls.filter(Boolean);
                    if (calls.length) message.tool_calls = calls;
                    return { message: message, usage: usage, finish_reason: finish, ms: performance.now() - t0, ttfbMs: ttfb };
                });
            });
        }
    };

    // ---------------------------------------------------------------------------------------------
    // Orchestrator: model <-> tools loop
    //
    // Two modes (settings.mode):
    //  - "scaffold" (default): the model first submits a structured plan (submit_plan: goal,
    //    preconditions, steps mapped to tools) that is validated; it can use analysis tools that
    //    pre-digest the document and playbooks for common compound requests; after execution the
    //    orchestrator shows it the document state and asks for verify_result, with one automatic
    //    repair round; plan steps left undone are pointed out once.
    //  - "direct": the plain tool loop (no plan, no analysis tools, no verification), kept to
    //    measure what the scaffolding brings.
    // In both modes a round whose calls all succeeded and all changed the document ends the action
    // without a closing model call (earlyStop), unless a verification is due.
    // ---------------------------------------------------------------------------------------------

    function addUsage(total, u) {
        if (!u) return;
        total.prompt_tokens += u.prompt_tokens || 0;
        total.completion_tokens += u.completion_tokens || 0;
        total.total_tokens += u.total_tokens || ((u.prompt_tokens || 0) + (u.completion_tokens || 0));
    }

    /**
     * opts: {instruction, toolset, settings, signal, lang, onEvent(type, data)}
     * Events: context, delta, round, tool:start, tool:end, verify, done, error.
     * Resolves {ok, ctx, summary, answer, done:[{name, args, result}], untracked:[keys], changed, plan,
     *           verification, cannotDo, question, entry}
     */
    ElixirAI.Orchestrator = {
        run: function (opts) {
            var ts = opts.toolset, s = opts.settings, emit = opts.onEvent || function () {};
            var mode = s.mode === 'direct' ? 'direct' : 'scaffold';
            var maxRounds = Math.max(1, Math.min(20, +s.maxRounds || 10));
            var earlyStop = s.earlyStop !== false;
            var tools = ts.toolsFor ? ts.toolsFor(mode) : ts.tools;
            var t0 = performance.now();
            var entry = {
                ts: new Date().toISOString(), editor: ts.id, prompt: opts.instruction, mode: mode,
                endpointHost: ElixirAI.Settings.publicView(s).endpointHost, model: s.model,
                extraParams: s.extraParams || '', sampling: s.sampling || 'deterministic', rounds: [], usage: { prompt_tokens: 0, completion_tokens: 0, total_tokens: 0 },
                outcome: 'running'
            };
            var ctx = null, messages = [], changed = false, untracked = [], groupStarted = false, finalText = '';
            var done = [], executed = {}, userNotes = [], answerText = '', laterText = '', nudged = false, verifyAsked = false, repairRounds = 0;

            var finish = function (outcome, err) {
                entry.outcome = outcome;
                entry.ms = Math.round(performance.now() - t0);
                entry.costUsd = ElixirAI.Settings.estimateCost(s, s.model, entry.usage);
                if (err) entry.error = String(err && err.message || err);
                entry.summary = finalText;
                entry.answer = answerText;
                entry.changed = changed;
                if (ctx) {
                    entry.plan = ctx.plan || null;
                    entry.verification = ctx.verification || null;
                    entry.cannotDo = ctx.cannotDo || null;
                    entry.question = ctx.question || null;
                }
                return ElixirAI.Journal.add(entry).then(function (id) { entry.id = id; return entry; });
            };

            // Plan steps whose tool never succeeded (mutating tools only)
            var missingSteps = function () {
                if (!ctx.plan) return [];
                return ctx.plan.steps.filter(function (st) { return ts.isMutating(st.tool) && !executed[st.tool]; });
            };
            var needVerify = function () {
                if (mode !== 'scaffold' || !changed || verifyAsked) return false;
                if (s.verify !== 'compound') return true;
                var mut = ctx.plan ? ctx.plan.steps.filter(function (st) { return ts.isMutating(st.tool); }).length : 0;
                return mut >= 2 || Object.keys(executed).some(function (n) { return ts.needsVerify && ts.needsVerify(n); });
            };

            // What to do after a round: 'continue' (call the model again) or 'stop'
            var afterRound = function (calls, results) {
                if (ctx.cannotDo || ctx.question) return Promise.resolve('stop');
                var allOk = results.every(function (r) { return r && r.ok; });
                var acted = calls.filter(function (c) { return ts.isMutating(c.name); });
                var onlyActions = acted.length > 0 && calls.every(function (c) { return ts.isMutating(c.name) || c.name === 'submit_plan'; });
                if (calls.length && !(allOk && onlyActions && earlyStop)) return Promise.resolve('continue');
                if (!calls.length && !changed) return Promise.resolve('stop');
                // the model stopped (or a clean action round): check the plan, then verify
                if (mode === 'scaffold') {
                    var miss = missingSteps();
                    if (miss.length && !nudged) {
                        nudged = true;
                        messages.push({ role: 'user', content: 'These steps of your plan were not done: ' + miss.map(function (m) { return m.tool + ' (' + m.purpose + ')'; }).join('; ') + '. Do them now, or call verify_result with the reason they are not needed.' });
                        return Promise.resolve('continue');
                    }
                    if (needVerify()) {
                        verifyAsked = true;
                        emit('verify', {});
                        return ts.verifyState(ctx).then(function (state) {
                            var auto = state && state.automatic_checks;
                            if (auto) { entry.automaticChecks = auto; delete state.automatic_checks; }
                            messages.push({
                                role: 'user', content: 'Verification. User request: ' + opts.instruction + (ctx.plan ? '\nYour goal: ' + ctx.plan.goal : '') +
                                    '\nDocument state now: ' + JSON.stringify(state) +
                                    (auto ? '\nAutomatic checks found these problems (they are certain, fix them): ' + auto.join(' ') : '') +
                                    '\nCheck every part of the request against this state (text, styles, alignment, sizes, lists, tables, page setup, pages). If everything is done, call verify_result {goal_met:true}. Otherwise call verify_result {goal_met:false, problems} AND, in the same turn, the tools that fix the problems.'
                            });
                            return 'continue';
                        });
                    }
                }
                if (!calls.length) return Promise.resolve('stop');
                return Promise.resolve(earlyStop ? 'stop' : 'continue');
            };

            return ts.init().then(function (c) {
                ctx = c;
                ctx.instruction = opts.instruction;
                entry.scope = { kind: ctx.scope, from: ctx.from, to: ctx.to, blocks: ctx.context.blocks.length };
                emit('context', ctx);
                messages = [
                    { role: 'system', content: ts.systemPrompt(ctx, { mode: mode, lang: opts.lang }) },
                    { role: 'user', content: ts.userMessage(opts.instruction, ctx) }
                ];
                var round = function (n) {
                    if (opts.signal && opts.signal.aborted) throw new Error('Cancelled');
                    emit('round', { n: n });
                    return ElixirAI.Connector.chat({
                        settings: s, messages: messages, tools: tools, signal: opts.signal,
                        // scaffolded mode: the first answer must go through the tools (the plan, cannot_do, ask_user or a reading tool)
                        toolChoice: mode === 'scaffold' && n === 1 ? { type: 'function', 'function': { name: 'submit_plan' } } : 'auto',
                        onDelta: function (d, all) { emit('delta', { text: all, round: n }); }
                    }).then(function (res) {
                        var rec = { n: n, ms: Math.round(res.ms), ttfbMs: Math.round(res.ttfbMs || 0), usage: res.usage, finish_reason: res.finish_reason, content: res.message.content || '', tool_calls: [] };
                        entry.rounds.push(rec);
                        addUsage(entry.usage, res.usage);
                        messages.push(res.message);
                        if (res.message.content) finalText = res.message.content;
                        var calls = res.message.tool_calls || [];
                        if (!calls.length && res.message.content) answerText = res.message.content;
                        if (n > 1 && res.message.content) laterText = res.message.content;
                        if (calls.length && n >= maxRounds) {
                            rec.capped = true;
                            throw new Error('Stopped after ' + maxRounds + ' tool rounds.');
                        }
                        // execute tool calls sequentially, in order
                        var chain = Promise.resolve(), results = [], named = [];
                        calls.forEach(function (call) {
                            chain = chain.then(function () {
                                var name = call['function'].name, args, t1 = performance.now();
                                try { args = call['function'].arguments ? JSON.parse(call['function'].arguments) : {}; }
                                catch (e) { args = undefined; }
                                var tc = { id: call.id, name: name, args: args !== undefined ? args : call['function'].arguments };
                                rec.tool_calls.push(tc);
                                named.push(tc);
                                emit('tool:start', tc);
                                var p;
                                if (args !== undefined && ts.normalize) args = ts.normalize(name, args);
                                var err = args === undefined ? { ok: false, error: { code: 'invalid_json', message: 'Arguments are not valid JSON.' } } : ts.validate(name, args, ctx);
                                if (!err && mode === 'direct' && tools.indexOf(tools.filter(function (t) { return t['function'].name === name; })[0]) < 0)
                                    err = { ok: false, error: { code: 'unknown_tool', message: 'Unknown tool "' + name + '".' } };
                                if (err) p = Promise.resolve(err);
                                else if (ts.isMutating(name) && !groupStarted) {
                                    groupStarted = true;
                                    p = ts.begin(ctx).then(function () { return ts.execute(name, args, ctx); });
                                } else p = ts.execute(name, args, ctx);
                                return p.then(function (result) {
                                    tc.result = result;
                                    tc.ms = Math.round(performance.now() - t1);
                                    results.push(result);
                                    if (result && result.ok && ts.isMutating(name)) {
                                        changed = true;
                                        executed[name] = (executed[name] || 0) + 1;
                                        done.push({ name: name, args: args, result: result });
                                        var note = ts.untrackedNote(name, args);
                                        if (note && untracked.indexOf(note) < 0) untracked.push(note);
                                        var un = ts.userNote && ts.userNote(name, args, result);
                                        if (un && userNotes.indexOf(un) < 0) userNotes.push(un);
                                    }
                                    emit('tool:end', tc);
                                    // the model gets the result without the UI-only summary data
                                    var forModel = result;
                                    if (result && result.done) { forModel = Object.assign({}, result); delete forModel.done; }
                                    messages.push({ role: 'tool', tool_call_id: call.id, content: JSON.stringify(forModel) });
                                });
                            });
                        });
                        return chain.then(function () {
                            if (ctx.verification && ctx.verification.goal_met === false) {
                                // the repair calls came with verify_result: allow one more look, then stop
                                if (repairRounds++ >= 1) return 'stop';
                            }
                            if (ctx.verification && ctx.verification.goal_met === true && !calls.some(function (c) { return ts.isMutating(c['function'].name); })) return 'stop';
                            return afterRound(named, results);
                        }).then(function (next) {
                            return next === 'continue' ? round(n + 1) : null;
                        });
                    });
                };
                return round(1);
            }).then(function () {
                return (groupStarted ? ts.end(ctx, true) : Promise.resolve()).then(function () {
                    entry.revisions = ctx.revisionsAfter || null;
                    return finish(changed ? 'applied' : (ctx.cannotDo ? 'cannot-do' : (ctx.question ? 'question' : 'no-change')));
                }).then(function () {
                    var res = {
                        ok: true, ctx: ctx, summary: finalText, answer: answerText || (ctx.verification && ctx.verification.message) || laterText, userNotes: userNotes, done: done, untracked: untracked, changed: changed,
                        plan: ctx.plan || null, verification: ctx.verification || null, cannotDo: ctx.cannotDo || null, question: ctx.question || null, entry: entry
                    };
                    emit('done', res);
                    return res;
                });
            }, function (err) {
                // Fatal error (network, cap, cancel): revert everything done in this action
                return (groupStarted && ctx ? ts.end(ctx, false) : Promise.resolve()).catch(function () {}).then(function () {
                    return finish(opts.signal && opts.signal.aborted ? 'cancelled' : 'error', err);
                }).then(function () {
                    emit('error', err);
                    return { ok: false, error: err, ctx: ctx, entry: entry, changed: false };
                });
            });
        }
    };

    return ElixirAI;
});
