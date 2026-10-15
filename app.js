(() => {
  "use strict";

  const cfg = window.APP_CONFIG;
  const STAGES = [
    ["new_lead", "New Lead"],
    ["quoted", "Quoted"],
    ["follow_up", "Follow-up"],
    ["booked", "Booked"],
    ["won", "Won"],
    ["lost", "Lost"],
  ];
  const STAGE_LABEL = Object.fromEntries(STAGES);
  const SEGMENT_LABEL = { residential: "Residential", commercial: "Commercial" };
  const GRAPH_SCOPES = ["Calendars.ReadWrite"];

  const state = {
    user: null,
    segment: "residential",
    stage: null, // null = show all stages
    search: "",
    customers: [],
    currentId: null,
  };

  const $ = (id) => document.getElementById(id);
  const esc = (s) =>
    String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  function toast(msg) {
    const t = $("toast");
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toast.timer);
    toast.timer = setTimeout(() => (t.hidden = true), 3500);
  }

  function showView(name) {
    $("login-view").classList.toggle("hidden", name !== "login");
    $("app-view").classList.toggle("hidden", name !== "app");
  }

  // Local time as "YYYY-MM-DDTHH:mm:00" (what Microsoft Graph expects with a timeZone)
  function localIso(d) {
    const p = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:00`;
  }

  function fmt(iso) {
    return new Date(iso).toLocaleString([], { dateStyle: "medium", timeStyle: "short" });
  }

  // ---------- Setup ----------
  const configured =
    !cfg.SUPABASE_URL.startsWith("PASTE") && !cfg.SUPABASE_ANON_KEY.startsWith("PASTE");
  const sb = configured ? window.supabase.createClient(cfg.SUPABASE_URL, cfg.SUPABASE_ANON_KEY) : null;

  // Microsoft sign-in. Resolves to the MSAL app, or null if Azure is not set up yet.
  const msalReady = (async () => {
    if (cfg.MS_CLIENT_ID.startsWith("PASTE")) return null;
    const redirectUri = window.location.origin + window.location.pathname.replace(/index\.html$/, "");
    const app = new msal.PublicClientApplication({
      auth: { clientId: cfg.MS_CLIENT_ID, authority: cfg.MS_AUTHORITY, redirectUri },
      cache: { cacheLocation: "localStorage" },
    });
    await app.initialize();
    const result = await app.handleRedirectPromise();
    if (result && result.account) app.setActiveAccount(result.account);
    return app;
  })().catch((err) => {
    console.error("Microsoft sign-in setup failed", err);
    return null;
  });

  async function getActiveAccount() {
    const app = await msalReady;
    if (!app) return null;
    return app.getActiveAccount() || app.getAllAccounts()[0] || null;
  }

  async function getGraphToken() {
    const app = await msalReady;
    const account = await getActiveAccount();
    if (!app || !account) throw new Error("Connect Outlook first (top right button).");
    try {
      const r = await app.acquireTokenSilent({ scopes: GRAPH_SCOPES, account });
      return r.accessToken;
    } catch (err) {
      if (err instanceof msal.InteractionRequiredAuthError) {
        await app.acquireTokenRedirect({ scopes: GRAPH_SCOPES, account });
      }
      throw err;
    }
  }

  async function createOutlookEvent({ subject, start, body }) {
    const token = await getGraphToken();
    const end = new Date(start.getTime() + 30 * 60 * 1000);
    const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
    const res = await fetch("https://graph.microsoft.com/v1.0/me/events", {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        subject,
        body: { contentType: "text", content: body },
        start: { dateTime: localIso(start), timeZone },
        end: { dateTime: localIso(end), timeZone },
        isReminderOn: true,
        reminderMinutesBeforeStart: 15,
      }),
    });
    if (!res.ok) throw new Error(`Outlook returned ${res.status}: ${await res.text()}`);
    return (await res.json()).id;
  }

  async function deleteOutlookEvent(eventId) {
    const token = await getGraphToken();
    const res = await fetch(`https://graph.microsoft.com/v1.0/me/events/${encodeURIComponent(eventId)}`, {
      method: "DELETE",
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok && res.status !== 404) throw new Error(`Outlook returned ${res.status}`);
  }

  async function refreshOutlookButton() {
    const btn = $("outlook-btn");
    const app = await msalReady;
    if (!app) {
      btn.textContent = "Outlook not set up";
      return;
    }
    const account = await getActiveAccount();
    btn.textContent = account ? "Outlook connected" : "Connect Outlook";
  }

  // ---------- Auth ----------
  function setSession(session) {
    state.user = session?.user ?? null;
    if (!state.user) {
      state.customers = [];
      showView("login");
      return;
    }
    showView("app");
    refreshOutlookButton();
    loadCustomers();
  }

  $("login-form").addEventListener("submit", async (e) => {
    e.preventDefault();
    const err = $("login-error");
    err.hidden = true;
    const { error } = await sb.auth.signInWithPassword({
      email: $("login-email").value.trim(),
      password: $("login-password").value,
    });
    $("login-password").value = "";
    if (error) {
      err.textContent = "Sign-in failed: " + error.message;
      err.hidden = false;
    }
  });

  $("signout-btn").addEventListener("click", () => sb.auth.signOut());

  $("outlook-btn").addEventListener("click", async () => {
    const app = await msalReady;
    if (!app) return toast("Outlook setup is not finished. See SETUP.md.");
    if (await getActiveAccount()) return toast("Outlook is already connected.");
    await app.loginRedirect({ scopes: GRAPH_SCOPES });
  });

  // ---------- Data ----------
  async function loadCustomers() {
    const { data, error } = await sb
      .from("customers")
      .select("*")
      .order("last_name")
      .order("first_name");
    if (error) return toast("Could not load customers: " + error.message);
    state.customers = data;
    render();
  }

  function render() {
    renderFunnel();
    renderList();
  }

  function inSegment(c) {
    return c.segment === state.segment;
  }

  function renderFunnel() {
    const segCustomers = state.customers.filter(inSegment);
    const chips = [
      { stage: null, label: "All", count: segCustomers.length },
      ...STAGES.map(([key, label]) => ({
        stage: key,
        label,
        count: segCustomers.filter((c) => c.stage === key).length,
      })),
    ];
    $("funnel").innerHTML = chips
      .map(
        (ch) => `
        <button class="stage-chip ${state.stage === ch.stage ? "active" : ""}" data-stage="${ch.stage ?? ""}">
          ${esc(ch.label)} <span class="count">${ch.count}</span>
        </button>`
      )
      .join("");
  }

  function renderList() {
    const q = state.search.trim().toLowerCase();
    const items = state.customers.filter(
      (c) =>
        inSegment(c) &&
        (!state.stage || c.stage === state.stage) &&
        (!q ||
          [c.first_name, c.last_name, c.phone, c.email, c.address].join(" ").toLowerCase().includes(q))
    );
    $("customer-list").innerHTML = items
      .map(
        (c) => `
        <li class="customer-item" data-id="${c.id}">
          <div class="name">${esc(c.first_name)} ${esc(c.last_name)}<span class="badge">${esc(STAGE_LABEL[c.stage])}</span></div>
          <div class="meta">${esc([c.phone, c.address].filter(Boolean).join(" · ") || "No contact details yet")}</div>
        </li>`
      )
      .join("");
    $("empty-state").hidden = items.length > 0;
  }

  function setSegment(seg) {
    state.segment = seg;
    state.stage = null;
    document.querySelectorAll(".segment").forEach((b) => b.classList.toggle("active", b.dataset.segment === seg));
    render();
  }

  document.querySelectorAll(".segment").forEach((btn) =>
    btn.addEventListener("click", () => setSegment(btn.dataset.segment))
  );

  $("funnel").addEventListener("click", (e) => {
    const chip = e.target.closest("[data-stage]");
    if (!chip) return;
    state.stage = chip.dataset.stage || null;
    render();
  });

  $("search").addEventListener("input", (e) => {
    state.search = e.target.value;
    renderList();
  });

  $("customer-list").addEventListener("click", (e) => {
    const li = e.target.closest("[data-id]");
    if (li) openDetail(li.dataset.id);
  });

  // ---------- Add / edit customer ----------
  const form = $("customer-form");
  const FIELDS = ["first_name", "last_name", "phone", "email", "address", "segment", "stage"];
  let editingId = null;

  $("stage-select").innerHTML = STAGES.map(([k, l]) => `<option value="${k}">${l}</option>`).join("");

  function openCustomerDialog(customer = null) {
    editingId = customer?.id ?? null;
    $("customer-dialog-title").textContent = customer ? "Edit customer" : "New customer";
    form.reset();
    const data = customer ?? { segment: state.segment, stage: "new_lead" };
    for (const field of FIELDS) form.elements[field].value = data[field] ?? "";
    $("customer-dialog").showModal();
  }

  $("add-customer-btn").addEventListener("click", () => openCustomerDialog());
  $("customer-cancel").addEventListener("click", () => $("customer-dialog").close());

  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    const f = form.elements;
    const payload = {
      first_name: f.first_name.value.trim(),
      last_name: f.last_name.value.trim(),
      phone: f.phone.value.trim(),
      email: f.email.value.trim(),
      address: f.address.value.trim(),
      segment: f.segment.value,
      stage: f.stage.value,
    };
    const wasEditing = editingId;
    const request = wasEditing
      ? sb.from("customers").update(payload).eq("id", wasEditing).select().single()
      : sb.from("customers").insert(payload).select().single();
    const { data, error } = await request;
    if (error) return toast("Save failed: " + error.message);

    $("customer-dialog").close();
    toast(wasEditing ? "Customer updated" : "Customer added");
    if (state.segment !== data.segment) setSegment(data.segment);
    await loadCustomers();
    openDetail(data.id);
  });

  // ---------- Customer detail: notes, follow-ups, stage ----------
  async function openDetail(id) {
    const c = state.customers.find((x) => x.id === id);
    if (!c) return;
    state.currentId = id;

    const [notesRes, fuRes] = await Promise.all([
      sb.from("notes").select("*").eq("customer_id", id).order("created_at", { ascending: false }),
      sb.from("follow_ups").select("*").eq("customer_id", id).order("due_at"),
    ]);
    if (notesRes.error || fuRes.error) return toast("Could not load customer details");

    const name = `${c.first_name} ${c.last_name}`.trim();
    const tomorrow9 = new Date();
    tomorrow9.setDate(tomorrow9.getDate() + 1);
    tomorrow9.setHours(9, 0, 0, 0);
    const defaultDue = localIso(tomorrow9).slice(0, 16);

    $("detail-body").innerHTML = `
      <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:8px">
        <div>
          <h2>${esc(name)}</h2>
          <span class="muted">${esc(SEGMENT_LABEL[c.segment])}</span>
        </div>
        <button class="ghost small" data-action="close" aria-label="Close">✕</button>
      </div>

      <div class="kv" style="margin-top:12px">
        <div><b>Phone</b> ${c.phone ? `<a href="tel:${esc(c.phone)}">${esc(c.phone)}</a>` : "-"}</div>
        <div><b>Email</b> ${c.email ? `<a href="mailto:${esc(c.email)}">${esc(c.email)}</a>` : "-"}</div>
        <div><b>Address</b> ${c.address ? esc(c.address) : "-"}</div>
      </div>

      <label style="margin-top:12px">Stage
        <select id="detail-stage">
          ${STAGES.map(([k, l]) => `<option value="${k}" ${k === c.stage ? "selected" : ""}>${l}</option>`).join("")}
        </select>
      </label>

      <div class="inline-actions">
        <button class="small" data-action="edit">Edit details</button>
        <button class="small danger" data-action="delete">Delete customer</button>
      </div>

      <div class="detail-section">
        <h3>Notes</h3>
        <div class="add-row">
          <textarea id="new-note" rows="3" placeholder="Add a note..."></textarea>
          <button class="primary small" data-action="add-note" style="justify-self:start">Add note</button>
        </div>
        <div style="margin-top:10px">
          ${notesRes.data.map((n) => `<div class="note"><span class="stamp">${fmt(n.created_at)}</span>${esc(n.body)}</div>`).join("") ||
            '<p class="muted">No notes yet.</p>'}
        </div>
      </div>

      <div class="detail-section">
        <h3>Follow-ups</h3>
        <div class="add-row">
          <input id="fu-title" value="Follow up with ${esc(name)}" />
          <input id="fu-due" type="datetime-local" value="${defaultDue}" />
          <button class="primary small" data-action="add-followup" style="justify-self:start">Add follow-up and send to Outlook</button>
        </div>
        <div style="margin-top:10px">
          ${fuRes.data.map((f) => `
            <div class="followup ${f.done ? "done" : ""}">
              <span class="stamp">${fmt(f.due_at)} · ${f.outlook_event_id ? "in Outlook" : "not in Outlook"}</span>
              <b>${esc(f.title)}</b>
              <div class="actions">
                <button class="small" data-action="toggle-fu" data-id="${f.id}" data-done="${f.done}">${f.done ? "Mark open" : "Mark done"}</button>
                <button class="small ghost" data-action="delete-fu" data-id="${f.id}" data-outlook="${esc(f.outlook_event_id ?? "")}">Delete</button>
              </div>
            </div>`).join("") || '<p class="muted">No follow-ups yet.</p>'}
        </div>
      </div>`;

    if (!$("detail-dialog").open) $("detail-dialog").showModal();
  }

  $("detail-body").addEventListener("change", async (e) => {
    if (e.target.id !== "detail-stage") return;
    const { error } = await sb.from("customers").update({ stage: e.target.value }).eq("id", state.currentId);
    if (error) return toast("Could not update stage: " + error.message);
    toast("Stage updated");
    await loadCustomers();
  });

  $("detail-body").addEventListener("click", async (e) => {
    const btn = e.target.closest("[data-action]");
    if (!btn) return;
    const id = state.currentId;
    const c = state.customers.find((x) => x.id === id);
    const action = btn.dataset.action;

    if (action === "close") return $("detail-dialog").close();

    if (action === "edit") {
      $("detail-dialog").close();
      return openCustomerDialog(c);
    }

    if (action === "delete") {
      if (!confirm("Delete this customer and all their notes and follow-ups? This cannot be undone.")) return;
      const { error } = await sb.from("customers").delete().eq("id", id);
      if (error) return toast("Delete failed: " + error.message);
      state.currentId = null;
      $("detail-dialog").close();
      toast("Customer deleted");
      await loadCustomers();
      return;
    }

    if (action === "add-note") {
      const body = $("new-note").value.trim();
      if (!body) return toast("Write a note first.");
      const { error } = await sb.from("notes").insert({ customer_id: id, body });
      if (error) return toast("Could not save note: " + error.message);
      toast("Note added");
      return openDetail(id);
    }

    if (action === "add-followup") {
      const title = $("fu-title").value.trim();
      const dueValue = $("fu-due").value;
      if (!title || !dueValue) return toast("Add a title and a date and time.");
      const dueAt = new Date(dueValue);
      const name = `${c.first_name} ${c.last_name}`.trim();
      btn.disabled = true;
      try {
        let outlookId = null;
        if (await getActiveAccount()) {
          outlookId = await createOutlookEvent({
            subject: title,
            start: dueAt,
            body: `Customer: ${name}\nPhone: ${c.phone || "-"}\nEmail: ${c.email || "-"}\nAddress: ${c.address || "-"}`,
          });
        }
        const { error } = await sb.from("follow_ups").insert({
          customer_id: id,
          title,
          due_at: dueAt.toISOString(),
          outlook_event_id: outlookId,
        });
        if (error) throw error;
        toast(outlookId ? "Follow-up added to Outlook" : "Follow-up saved. Connect Outlook to add it to your calendar.");
        return openDetail(id);
      } catch (err) {
        toast("Follow-up failed: " + err.message);
      } finally {
        btn.disabled = false;
      }
      return;
    }

    if (action === "toggle-fu") {
      const nowDone = btn.dataset.done !== "true";
      const { error } = await sb.from("follow_ups").update({ done: nowDone }).eq("id", btn.dataset.id);
      if (error) return toast("Could not update follow-up: " + error.message);
      return openDetail(id);
    }

    if (action === "delete-fu") {
      if (!confirm("Delete this follow-up?")) return;
      if (btn.dataset.outlook) {
        try {
          await deleteOutlookEvent(btn.dataset.outlook);
        } catch (err) {
          console.warn("Could not remove the Outlook event", err);
          toast("Follow-up deleted, but the Outlook event could not be removed.");
        }
      }
      const { error } = await sb.from("follow_ups").delete().eq("id", btn.dataset.id);
      if (error) return toast("Delete failed: " + error.message);
      return openDetail(id);
    }
  });

  // ---------- Start ----------
  if (!configured) {
    showView("login");
    const err = $("login-error");
    err.textContent = "Setup not finished: fill in config.js first (see SETUP.md).";
    err.hidden = false;
    $("login-form").querySelector("button").disabled = true;
  } else {
    sb.auth.onAuthStateChange((_event, session) => setSession(session));
  }
})();
