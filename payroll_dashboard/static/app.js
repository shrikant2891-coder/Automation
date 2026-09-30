const TABS = [
  { id: "generator", label: "Generator" },
  { id: "settings", label: "Settings" },
  { id: "salary", label: "Salary" },
  { id: "master", label: "Master Data" },
  { id: "attendance", label: "Attendance" },
  { id: "increments", label: "Increments" },
  { id: "tax", label: "Income Tax" },
  { id: "xml", label: "XML" },
];

let state = null;

function $(id) {
  return document.getElementById(id);
}

function fmt(n) {
  if (n == null || n === "") return "—";
  return Number(n).toLocaleString("en-IN");
}

function initNav() {
  const nav = $("nav");
  TABS.forEach((t, i) => {
    const b = document.createElement("button");
    b.textContent = t.label;
    b.dataset.tab = t.id;
    if (i === 0) b.classList.add("active");
    b.addEventListener("click", () => showTab(t.id));
    nav.appendChild(b);
  });
}

function showTab(id) {
  document.querySelectorAll("nav button").forEach((b) => {
    b.classList.toggle("active", b.dataset.tab === id);
  });
  document.querySelectorAll(".panel").forEach((p) => {
    p.classList.toggle("active", p.dataset.panel === id);
  });
}

async function loadPayroll() {
  const res = await fetch("/api/payroll");
  state = await res.json();
  renderAll();
}

async function loadXmlRaw() {
  const res = await fetch("/api/xml");
  $("xml-raw").textContent = await res.text();
}

function fillEmpSelect(sel, withCode = true) {
  sel.innerHTML = "";
  state.employees.forEach((e) => {
    const o = document.createElement("option");
    o.value = e.code;
    o.textContent = withCode ? `${e.code} — ${e.master.name || ""}` : e.code;
    sel.appendChild(o);
  });
}

function renderAll() {
  $("meta-line").textContent = `FY ${state.fy} · Updated ${state.updated || "—"} · ${state.employees.length} employees`;

  const stb = $("settings-table").querySelector("tbody");
  stb.innerHTML = "";
  Object.entries(state.settings).forEach(([k, v]) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${k}</td><td><strong>${v}</strong></td>`;
    stb.appendChild(tr);
  });

  const sal = $("salary-table").querySelector("tbody");
  sal.innerHTML = "";
  (state.computed?.salary_rows || []).forEach((r) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${r.code}</td>
      <td>${r.name || ""}</td>
      <td>${r.pf_status}</td>
      <td>${fmt(r.basic)}</td>
      <td>${fmt(r.pre_sep_gross)}</td>
      <td>${fmt(r.post_sep_gross)}</td>
      <td>${fmt(r.pf_pre)}</td>
      <td>${fmt(r.pf_post)}</td>`;
    sal.appendChild(tr);
  });

  fillEmpSelect($("gen-emp"));
  fillEmpSelect($("master-emp"));
  fillEmpSelect($("att-emp"));
  fillEmpSelect($("tax-emp"));

  const gm = $("gen-month");
  gm.innerHTML = "";
  state.months.forEach((m) => {
    const o = document.createElement("option");
    o.value = m.label;
    o.textContent = m.label;
    gm.appendChild(o);
  });

  renderMaster();
  renderAttendance();
  renderIncrements();
  loadXmlRaw();
}

function renderMaster() {
  const code = $("master-emp").value;
  const emp = state.employees.find((e) => e.code === code);
  $("master-detail").textContent = JSON.stringify(emp?.master || {}, null, 2);
}

function renderAttendance() {
  const code = $("att-emp").value;
  const emp = state.employees.find((e) => e.code === code);
  const tb = $("att-table").querySelector("tbody");
  tb.innerHTML = "";
  (emp?.attendance || []).forEach((a) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${a.label}</td><td>${a.days}</td>`;
    tb.appendChild(tr);
  });
}

function renderIncrements() {
  const tb = $("inc-table").querySelector("tbody");
  tb.innerHTML = "";
  (state.increments || []).forEach((i) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${i.emp_code}</td>
      <td>${i.effective_month}</td>
      <td>${i.new_basic}</td>
      <td>${i.old_basic}</td>
      <td>${i.arrear_months}</td>
      <td>${i.notes || ""}</td>`;
    tb.appendChild(tr);
  });
}

async function previewSlip() {
  const emp = $("gen-emp").value;
  const month = $("gen-month").value;
  const res = await fetch(`/api/slip?emp=${encodeURIComponent(emp)}&month=${encodeURIComponent(month)}`);
  const slip = await res.json();
  const el = $("slip-preview");
  if (slip.skip_pdf) {
    el.innerHTML = `<p class="muted"><strong>Skipped:</strong> zero attendance for ${month}.</p>`;
    return;
  }
  el.innerHTML = `
    <h3>${slip.emp_name} (${slip.emp_code}) — ${slip.month}</h3>
    <p class="muted">Attendance: ${slip.days_attended} / ${slip.days_in_month}</p>
    <div class="kv">
      <span>Basic</span><span class="v">₹ ${fmt(slip.earnings.basic)}</span>
      <span>HRA</span><span class="v">₹ ${fmt(slip.earnings.hra)}</span>
      <span>Special Allowance</span><span class="v">₹ ${fmt(slip.earnings.special_allowance)}</span>
      <span>Increment arrear</span><span class="v">₹ ${fmt(slip.earnings.increment_arrear)}</span>
      <span><strong>Gross</strong></span><span class="v"><strong>₹ ${fmt(slip.earnings.gross)}</strong></span>
      <span>PF (employee)</span><span class="v">₹ ${fmt(slip.deductions.pf_employee)}</span>
      <span>PF arrear</span><span class="v">₹ ${fmt(slip.deductions.pf_arrear)}</span>
      <span>TDS</span><span class="v">₹ ${fmt(slip.deductions.tds)}</span>
      <span><strong>Net pay</strong></span><span class="v"><strong>₹ ${fmt(slip.net_pay)}</strong></span>
    </div>
    <p class="muted" style="margin-top:0.75rem">Full-month PF: ₹ ${fmt(slip.components_full_month.pf_employee)} · Gross (full): ₹ ${fmt(slip.components_full_month.gross)}</p>`;
}

async function previewTax() {
  const emp = $("tax-emp").value;
  const month = state.months[0]?.label || "Apr'26";
  const res = await fetch(`/api/slip?emp=${encodeURIComponent(emp)}&month=${encodeURIComponent(month)}`);
  const slip = await res.json();
  $("tax-detail").innerHTML = `
    <div class="kv">
      <span>Projected annual taxable</span><span class="v">₹ ${fmt(slip.tax.annual_taxable)}</span>
      <span>Annual tax (est.)</span><span class="v">₹ ${fmt(slip.tax.annual_tax)}</span>
      <span>Monthly TDS (est.)</span><span class="v">₹ ${fmt(slip.deductions.tds)}</span>
    </div>`;
}

async function runCommand() {
  const command = $("command-input").value.trim();
  if (!command) return;
  $("command-result").textContent = "Applying…";
  const res = await fetch("/api/command", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ command }),
  });
  const data = await res.json();
  if (!res.ok) {
    $("command-result").textContent = data.detail || "Command failed";
    $("command-result").style.color = "#f87171";
    return;
  }
  $("command-result").style.color = "";
  $("command-result").textContent = data.message;
  state = data.payroll;
  renderAll();
}

function bind() {
  $("command-run").addEventListener("click", runCommand);
  $("command-input").addEventListener("keydown", (e) => {
    if (e.key === "Enter") runCommand();
  });
  $("refresh-btn").addEventListener("click", loadPayroll);
  $("gen-preview").addEventListener("click", previewSlip);
  $("master-emp").addEventListener("change", renderMaster);
  $("att-emp").addEventListener("change", renderAttendance);
  $("tax-emp").addEventListener("change", previewTax);
}

initNav();
bind();
loadPayroll().then(() => previewTax());
