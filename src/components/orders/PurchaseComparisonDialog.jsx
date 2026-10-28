import { useEffect, useMemo, useRef, useState } from "react";
import { ArrowLeft, Check, LoaderCircle, RefreshCw, Scale, X } from "lucide-react";
import { formatILS } from "../../lib/orders";
import { PURCHASE_CATEGORIES, normalizePurchaseCategories } from "../../lib/purchaseCategories";
import { buildComparisonPayload, comparisonCategory, selectedComparisonTotal, summarizeComparison, validateComparison } from "../../lib/purchaseComparison";
import { confirmPurchaseComparison, fetchEmptyComparisonOrders } from "../../lib/purchaseComparisonApi";
import PurchaseCategoryLabels from "./PurchaseCategoryLabels";
import "./purchase-comparison.css";

const categoryName = (value) => PURCHASE_CATEGORIES.find((c) => c.value === value)?.label || value;
const makeGroup = (index) => ({ key: crypto.randomUUID(), category: "fast", target: "", mode: "new", orderId: "", name: `طلب مقارنة ${index + 1}` });

export default function PurchaseComparisonDialog({ purchases, onClose, onSaved }) {
  const [source] = useState(() => purchases.map((purchase) => ({ ...purchase })));
  const [selected, setSelected] = useState(() => new Set(source.map((p) => p.id)));
  const [choices, setChoices] = useState({});
  const [search, setSearch] = useState("");
  const [count, setCount] = useState("2");
  const [groups, setGroups] = useState(() => [makeGroup(0), makeGroup(1)]);
  const [step, setStep] = useState(0);
  const [plans, setPlans] = useState([]);
  const [activePlan, setActivePlan] = useState(0);
  const [assignments, setAssignments] = useState({});
  const [destinations, setDestinations] = useState([]);
  const [destinationsError, setDestinationsError] = useState("");
  const [destinationRevision, setDestinationRevision] = useState(0);
  const [error, setError] = useState("");
  const [running, setRunning] = useState(false);
  const [saving, setSaving] = useState(false);
  const [retrying, setRetrying] = useState(false);
  const worker = useRef(null);
  const deadline = useRef(null);
  const request = useRef(null);
  const root = useRef(null);
  const busy = running || saving || retrying;
  const rows = useMemo(() => source.filter((p) => selected.has(p.id)).map((p) => ({ ...p,
    comparison_category: comparisonCategory(p, choices[p.id]) })), [source, selected, choices]);
  const selectedTotal = useMemo(() => {
    try { return selectedComparisonTotal(rows); } catch { return null; }
  }, [rows]);
  const summary = useMemo(() => step === 2 ? summarizeComparison(rows, groups, assignments)
    : { groups: [], leftovers: [], included: 0, deviation: 0, valid: false }, [rows, groups, assignments, step]);
  const displayed = source.filter((p) => p.customer_name.toLowerCase().includes(search.trim().toLowerCase()));
  const edited = plans[activePlan] && rows.some((p) => assignments[p.id] !== plans[activePlan].assignments[p.id]);
  const groupName = (group, index) => group.mode === "existing"
    ? destinations.find((d) => d.id === group.orderId)?.order_name || `طلب ${index + 1}`
    : group.name || `طلب ${index + 1}`;

  function stopWorker() {
    worker.current?.terminate();
    worker.current = null;
    clearTimeout(deadline.current);
    setRunning(false);
  }

  useEffect(() => {
    const previous = document.activeElement;
    root.current?.focus();
    const overflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => {
      worker.current?.terminate();
      clearTimeout(deadline.current);
      document.body.style.overflow = overflow;
      previous?.focus?.();
    };
  }, []);

  useEffect(() => {
    let cancelled = false;
    setDestinationsError("");
    fetchEmptyComparisonOrders().then((data) => { if (!cancelled) setDestinations(data); })
      .catch((failure) => { if (!cancelled) setDestinationsError(failure.message || "تعذر تحميل الطلبات الفارغة"); });
    return () => { cancelled = true; };
  }, [destinationRevision]);

  function changeGroup(key, patch) {
    setGroups((current) => current.map((group) => group.key === key ? { ...group, ...patch } : group));
    setPlans([]);
  }

  function resizeGroups(value) {
    setCount(value);
    setPlans([]);
    const number = Number(value);
    if (!Number.isInteger(number) || number < 1 || number > 20) return;
    setGroups((current) => Array.from({ length: number }, (_, index) => current[index] || makeGroup(index)));
  }

  function close() {
    if (saving) return;
    const uncertain = !!request.current;
    if (uncertain && !window.confirm("نتيجة الحفظ غير مؤكدة. إغلاق المقارنة وتحديث المشتريات؟")) return;
    stopWorker();
    onClose(uncertain);
  }

  function runComparison() {
    setError("");
    try {
      if (Number(count) !== groups.length || !Number.isInteger(Number(count))) throw new Error("أدخلي عدد الطلبات من 1 إلى20");
      validateComparison(rows, groups);
      const seen = new Set();
      groups.forEach((group) => {
        if (group.mode === "existing") {
          const destination = destinations.find((d) => d.id === group.orderId);
          if (!destination || destination.shein_type !== (group.category === "fast" ? "normal" : group.category) || seen.has(group.orderId)) {
            throw new Error("اختاري طلباً فارغاً مختلفاً من نفس التصنيف لكل وجهة");
          }
          seen.add(group.orderId);
        } else if (!group.name.trim()) throw new Error("أدخلي أسماء الطلبات الجديدة");
      });
      setRunning(true);
      const current = new Worker(new URL("../../lib/purchaseComparison.worker.js", import.meta.url), { type: "module" });
      worker.current = current;
      current.onmessage = ({ data }) => {
        if (worker.current !== current) return;
        stopWorker();
        if (data.error) { setError(data.error); return; }
        setPlans(data.plans);
        setActivePlan(0);
        setAssignments(data.plans[0].assignments);
        setStep(2);
      };
      current.onerror = () => { if (worker.current === current) { stopWorker(); setError("تعذرت المقارنة. أعيدي المحاولة أو اختاري مشتريات أقل."); } };
      deadline.current = setTimeout(() => { stopWorker(); setError("استغرقت المقارنة وقتاً طويلاً. اختاري مشتريات أقل أو أعيدي المحاولة."); }, 45000);
      current.postMessage({ purchases: rows, groups });
    } catch (failure) {
      stopWorker();
      setError(failure.message);
    }
  }

  function movePurchase(purchase, target) {
    const next = { ...assignments };
    if (target) next[purchase.id] = target;
    else delete next[purchase.id];
    const group = groups.find((item) => item.key === target);
    const qty = rows.filter((p) => next[p.id] === target).reduce((sum, p) => sum + Number(p.qty), 0);
    if (group && (group.category !== purchase.comparison_category || qty > 100)) {
      setError("لا يمكن خلط التصنيفات أو تجاوز 100 قطعة في الطلب");
      return;
    }
    setError("");
    setAssignments(next);
  }

  async function confirm() {
    if (saving) return;
    setError("");
    try {
      if (!request.current) {
        const payload = buildComparisonPayload(rows, groups, assignments);
        if (!window.confirm(`تأكيد توزيع ${summary.included} مشترى على ${groups.length} طلب؟`)) return;
        request.current = { id: crypto.randomUUID(), groups: payload };
      }
      setSaving(true);
      const result = await confirmPurchaseComparison(request.current.id, request.current.groups);
      request.current = null;
      setRetrying(false);
      await onSaved(result);
    } catch (failure) {
      const code = String(failure.code || "");
      const definitive = code === "P0001" || code.startsWith("42") || code.startsWith("22") || code.startsWith("23") || code.startsWith("PGRST");
      if (definitive) request.current = null;
      setRetrying(!!request.current);
      setError(failure.message || "تعذر حفظ التقسيم");
    } finally { setSaving(false); }
  }

  function renderItems(items) {
    return <div className="comparison-items">{items.map((purchase) => (
      <div className="comparison-item" key={purchase.id}>
        <div><strong>{purchase.customer_name}</strong><small>{purchase.qty} قطع · {formatILS(purchase.price)} ₪</small></div>
        <select aria-label={`وجهة ${purchase.customer_name}`} value={assignments[purchase.id] || ""} disabled={busy}
          onChange={(event) => movePurchase(purchase, event.target.value)}>
          <option value="">غير معيّن</option>
          {groups.filter((group) => group.category === purchase.comparison_category).map((group, index) => (
            <option key={group.key} value={group.key}>{groupName(group, index)} · {categoryName(group.category)}</option>
          ))}
        </select>
      </div>
    ))}</div>;
  }

  return <div className="comparison-backdrop">
    <section className="comparison-workspace" role="dialog" aria-modal="true" aria-labelledby="comparison-title" tabIndex={-1} ref={root}
      onKeyDown={(event) => {
        if (event.key === "Escape") close();
        if (event.key !== "Tab") return;
        const controls = [...root.current.querySelectorAll("button:not(:disabled), input:not(:disabled), select:not(:disabled)")].filter((node) => node.getClientRects().length);
        const first = controls[0], last = controls.at(-1);
        if (!first) { event.preventDefault(); return; }
        if (event.shiftKey && [first, root.current].includes(document.activeElement)) { event.preventDefault(); last.focus(); }
        else if (!event.shiftKey && [last, root.current].includes(document.activeElement)) { event.preventDefault(); first.focus(); }
      }}>
      <header className="comparison-heading"><h2 id="comparison-title"><Scale size={22} /> مقارنة التقسيم</h2>
        <button type="button" className="icon-btn" title="إغلاق" aria-label="إغلاق المقارنة" disabled={saving} onClick={close}><X size={20} /></button>
      </header>
      <nav className="comparison-steps" aria-label="خطوات المقارنة">
        {["المشتريات", "الطلبات والأهداف", "مقارنة النتائج"].map((label, index) => <button type="button" key={label} aria-current={step === index ? "step" : undefined}
          disabled={busy || (index === 2 && !plans.length)} onClick={() => { setError(""); setStep(index); }}>{index + 1}. {label}</button>)}
      </nav>
      <div className="comparison-content">
        {error ? <p role="alert" className="comparison-error">{error}</p> : null}
        {step === 0 ? <>
          <div className="comparison-toolbar">
            <label className="comparison-check"><input type="checkbox" checked={source.length > 0 && selected.size === source.length}
              ref={(node) => { if (node) node.indeterminate = selected.size > 0 && selected.size < source.length; }}
              onChange={(event) => { setSelected(event.target.checked ? new Set(source.map((p) => p.id)) : new Set()); setPlans([]); }} /> تحديد الكل ({selected.size}/{source.length})</label>
            <input type="search" aria-label="بحث في مشتريات المقارنة" placeholder="بحث باسم الزبون..." value={search} onChange={(event) => setSearch(event.target.value)} />
          </div>
          <div className="comparison-selection">{displayed.map((purchase) => {
            const labels = normalizePurchaseCategories(purchase.purchase_categories);
            return <div key={purchase.id} className="comparison-selection-row">
              <label className="comparison-check"><input type="checkbox" checked={selected.has(purchase.id)} onChange={(event) => {
                setSelected((current) => { const next = new Set(current); if (event.target.checked) next.add(purchase.id); else next.delete(purchase.id); return next; }); setPlans([]);
              }} /><span><strong>{purchase.customer_name}</strong><small>{purchase.qty} قطع · السعر الأصلي {formatILS(purchase.price)} ₪</small></span></label>
              {labels.length === 1 ? <PurchaseCategoryLabels categories={labels} /> : <select aria-label={`تصنيف ${purchase.customer_name}`} value={choices[purchase.id] || ""} disabled={!selected.has(purchase.id)} onChange={(event) => { setChoices((current) => ({ ...current, [purchase.id]: event.target.value })); setPlans([]); }}>
                <option value="">اختاري تصنيف المقارنة</option>
                {PURCHASE_CATEGORIES.filter((c) => !labels.length || labels.includes(c.value)).map((c) => <option key={c.value} value={c.value}>{c.label}</option>)}
              </select>}
            </div>;
          })}</div>
          {!displayed.length ? <p className="comparison-muted">لا توجد مشتريات مطابقة</p> : null}
        </> : null}
        {step === 1 ? <>
          <div className="comparison-toolbar"><label>عدد الطلبات<input type="number" min="1" max="20" step="1" value={count} disabled={busy} onChange={(event) => resizeGroups(event.target.value)} /></label>
            <button type="button" className="icon-btn" title="تحديث الطلبات الفارغة" aria-label="تحديث الطلبات الفارغة" disabled={busy} onClick={() => setDestinationRevision((v) => v + 1)}><RefreshCw size={18} /></button></div>
          {destinationsError ? <p role="alert" className="comparison-error">{destinationsError}</p> : null}
          <div className="comparison-order-configs">{groups.map((group, index) => {
            const options = destinations.filter((order) => order.shein_type === (group.category === "fast" ? "normal" : group.category));
            return <fieldset key={group.key} className="comparison-order-config" disabled={busy}>
              <legend>طلب {index + 1}</legend>
              <label>التصنيف<select aria-label="التصنيف" value={group.category} onChange={(event) => changeGroup(group.key, { category: event.target.value, orderId: "" })}>{PURCHASE_CATEGORIES.map((c) => <option key={c.value} value={c.value}>{c.label}</option>)}</select></label>
              <label>السعر المستهدف (₪)<input type="number" min="0.01" step="0.01" max="1000000000" value={group.target} onChange={(event) => changeGroup(group.key, { target: event.target.value })} /></label>
              <label>الوجهة<select aria-label="الوجهة" value={group.mode} onChange={(event) => changeGroup(group.key, { mode: event.target.value, orderId: "" })}><option value="new">طلب جديد</option><option value="existing">طلب فارغ قيد الانتظار</option></select></label>
              {group.mode === "new" ? <label>اسم الطلب<input maxLength={200} value={group.name} onChange={(event) => changeGroup(group.key, { name: event.target.value })} /></label>
                : <label>الطلب<select aria-label="الطلب" value={group.orderId} onChange={(event) => changeGroup(group.key, { orderId: event.target.value })}><option value="">اختاري طلباً فارغاً</option>{options.map((order) => <option key={order.id} value={order.id}>{order.order_name}</option>)}</select></label>}
            </fieldset>;
          })}</div>
        </> : null}
        {step === 2 ? <>
          <div className="comparison-alternatives" role="tablist" aria-label="اقتراحات التقسيم">{plans.map((plan, index) => (
            <button key={index} type="button" role="tab" aria-selected={activePlan === index} disabled={busy} onClick={() => { setActivePlan(index); setAssignments({ ...plan.assignments }); setError(""); }}>
              <strong>اقتراح {index + 1}</strong><span>{plan.included} مشترى · الفرق {formatILS(plan.deviation / 100)} ₪</span>
            </button>
          ))}</div>
          <div className="comparison-summary"><span>موزّع: <strong>{summary.included}</strong></span><span>غير معيّن: <strong>{summary.leftovers.length}</strong></span><span>مجموع الفروق: <strong>{formatILS(summary.deviation / 100)} ₪</strong></span>{edited ? <span>تعديل يدوي</span> : null}</div>
          <div className="comparison-results">{summary.groups.map((group, index) => <section key={group.key} className="comparison-result-group">
            <header><h3>{groupName(group, index)}</h3><PurchaseCategoryLabels categories={[group.category]} /></header>
            <dl><div><dt>السعر الأصلي</dt><dd>{formatILS(group.cents / 100)} ₪</dd></div><div><dt>المستهدف</dt><dd>{formatILS(group.target)} ₪</dd></div><div><dt>الفرق</dt><dd>{group.difference > 0 ? "+" : ""}{formatILS(group.difference / 100)} ₪</dd></div><div><dt>القطع</dt><dd>{group.qty}/100</dd></div></dl>
            {!group.valid ? <p className="comparison-error">الطلب فارغ أو يتجاوز الحد أو يحتوي تصنيفاً مختلفاً</p> : null}
            {renderItems(group.items)}
          </section>)}</div>
          <section className="comparison-leftovers"><h3>مشتريات تبقى غير معيّنة ({summary.leftovers.length})</h3>{renderItems(summary.leftovers)}</section>
        </> : null}
      </div>
      <footer className="comparison-footer">
        {step > 0 ? <button type="button" className="btn-ghost" disabled={busy} onClick={() => { setStep((current) => current - 1); setError(""); }}><ArrowLeft size={18} /> السابق</button> : <span />}
        {step === 0 ? <div className="comparison-footer-actions comparison-next">
          <output aria-label="مجموع الأسعار المحددة">المجموع: <strong><bdi dir="ltr">{selectedTotal === null ? "—" : formatILS(selectedTotal)} ₪</bdi></strong></output>
          <button type="button" className="btn-primary" disabled={!rows.length} onClick={() => { setStep(1); setError(""); }}>التالي ({rows.length})</button>
        </div> : null}
        {step === 1 ? <div className="comparison-footer-actions">{running ? <button type="button" className="btn-ghost" onClick={stopWorker}>إلغاء المقارنة</button> : null}<button type="button" className="btn-primary" disabled={running || saving} onClick={runComparison}>{running ? <LoaderCircle size={18} className="comparison-spinner" /> : <Scale size={18} />} {running ? "جارٍ المقارنة..." : "مقارنة التقسيم"}</button></div> : null}
        {step === 2 ? <button type="button" className="btn-primary" disabled={saving || (!retrying && !summary.valid)} onClick={confirm}>{saving ? <LoaderCircle size={18} className="comparison-spinner" /> : <Check size={18} />} {retrying ? "إعادة محاولة الحفظ" : "تأكيد التوزيع"}</button> : null}
      </footer>
    </section>
  </div>;
}
