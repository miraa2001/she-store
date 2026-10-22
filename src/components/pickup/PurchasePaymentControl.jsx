import { useEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import { Banknote, History, X } from "lucide-react";
import { sb } from "../../lib/supabaseClient";
import { formatILS } from "../../lib/orders";
import { formatDateTime } from "../../lib/dateFormat";
import { agreedAmount, paidAmount, remainingAmount, collectibleAmount, canRecordPayment, canConfirmStaffPickup, confirmStaffPickup, paymentRpc, recordPayment } from "../../lib/payments";
import { buildPickupStatusMessage, notifyPickupStatus } from "../../lib/pickupNotifications";
import { setBodyScrollLock } from "../../lib/bodyScrollLock";
import PickupAnimatedCheckbox from "../common/PickupAnimatedCheckbox";
import "./purchase-payments.css";

export default function PurchasePaymentControl({ purchase, role, instant = false, onSaved, compact = false }) {
  const [open, setOpen] = useState(false);
  const [state, setState] = useState(null);
  const [amount, setAmount] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [busy, setBusy] = useState(false);
  const [reason, setReason] = useState("");
  const [cancelId, setCancelId] = useState("");
  const requestRef = useRef("");
  const dialogRef = useRef(null);
  const busyRef = useRef(false);
  const active = state?.purchase || purchase;
  const canPay = canRecordPayment(role);
  const canConfirmPickup = canConfirmStaffPickup(role);

  useEffect(() => {
    if (!open) return undefined;
    let cancelled = false;
    setBodyScrollLock(`payment-${purchase.id}`, true);
    setLoading(true); setError(""); setState(null); setCancelId(""); setReason("");
    paymentRpc(sb, "get_purchase_payment_state", { p_id: purchase.id, p_instant: instant }).then((result) => {
      if (cancelled) return;
      setState(result); setAmount(remainingAmount(result.purchase).toFixed(2)); requestRef.current = crypto.randomUUID();
    }).catch((e) => { if (!cancelled) setError(e.message || "تعذر تحميل الدفعات."); })
      .finally(() => { if (!cancelled) setLoading(false); });
    return () => { cancelled = true; setBodyScrollLock(`payment-${purchase.id}`, false); };
  }, [open, purchase.id, instant]);

  useEffect(() => {
    if (!open) return undefined;
    const previousFocus = document.activeElement;
    const frame = requestAnimationFrame(() => dialogRef.current?.querySelector("button")?.focus());
    const escape = (event) => {
      if (event.key === "Escape" && !busyRef.current) setOpen(false);
      if (event.key !== "Tab") return;
      const items = [...(dialogRef.current?.querySelectorAll("button:not([disabled]), input:not([disabled])") || [])];
      const first = items[0], last = items[items.length - 1];
      if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last?.focus(); }
      else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first?.focus(); }
    };
    document.addEventListener("keydown", escape);
    return () => {
      cancelAnimationFrame(frame);
      document.removeEventListener("keydown", escape);
      if (previousFocus?.isConnected) previousFocus.focus();
    };
  }, [open]);

  async function save(event) {
    event.preventDefault();
    if (!state || busyRef.current || !canPay) return;
    busyRef.current = true; setBusy(true); setError("");
    try {
      const result = await recordPayment(sb, { requestId: requestRef.current, purchase: active, amount, instant, receive: !active.picked_up });
      if (result.changed_receipt) await notifyPickupStatus(buildPickupStatusMessage({
        picked: true, customerName: active.customer_name, price: agreedAmount(active), pickupLabel: active.pickup_point
      }));
      setOpen(false);
      await onSaved?.(result.purchase);
    } catch (e) { setError(e.message || "تعذر حفظ الدفعة."); }
    finally { busyRef.current = false; setBusy(false); }
  }

  async function saveStaffPickup() {
    if (!state || busyRef.current || !canConfirmPickup || active.picked_up) return;
    busyRef.current = true; setBusy(true); setError("");
    try {
      const result = await confirmStaffPickup(sb, { requestId: requestRef.current, purchase: active, instant });
      if (result.changed_receipt) await notifyPickupStatus(buildPickupStatusMessage({
        picked: true, customerName: active.customer_name, price: agreedAmount(result.purchase), pickupLabel: active.pickup_point
      }));
      setOpen(false);
      await onSaved?.(result.purchase);
    } catch (e) { setError(e.message || "تعذر تأكيد الاستلام."); }
    finally { busyRef.current = false; setBusy(false); }
  }

  async function cancelPayment() {
    if (!cancelId || !reason.trim() || busyRef.current) return;
    busyRef.current = true; setBusy(true); setError("");
    try {
      const result = await paymentRpc(sb, "cancel_purchase_payment", { p_payment_id: cancelId, p_reason: reason.trim() });
      if (result?.id !== purchase.id) throw new Error("لم يتم تأكيد إلغاء الدفعة.");
      setOpen(false);
      await onSaved?.(result);
    } catch (e) { setError(e.message || "تعذر إلغاء الدفعة."); }
    finally { busyRef.current = false; setBusy(false); }
  }

  return <div className="purchase-payment-control">
    {!compact ? <PickupAnimatedCheckbox checked={purchase.picked_up} disabled={busy || (!canPay && (!canConfirmPickup || purchase.picked_up))}
      ariaLabel={`استلام ${purchase.customer_name}`} onChange={() => setOpen(true)} /> : null}
    <div className="purchase-payment-summary">
      <span>المدفوع {formatILS(paidAmount(purchase))} ₪ · المتبقي {formatILS(remainingAmount(purchase))} ₪</span>
      {collectibleAmount(purchase) > 0 ? <small>بانتظار التحصيل {formatILS(collectibleAmount(purchase))} ₪</small> : null}
      <div className="purchase-payment-actions">
        {canPay && remainingAmount(purchase) > 0 ? <button type="button" onClick={() => setOpen(true)}><Banknote size={15} />{purchase.picked_up ? "إضافة دفعة" : "استلام ودفعة"}</button> : null}
        <button type="button" title="سجل الدفعات" aria-label={`سجل دفعات ${purchase.customer_name}`} onClick={() => setOpen(true)}><History size={16} /></button>
      </div>
    </div>
    {open ? createPortal(<div className="purchase-payment-overlay" onClick={() => { if (!busy) setOpen(false); }}>
      <section ref={dialogRef} className="purchase-payment-dialog" role="dialog" aria-modal="true" aria-label={`دفعات ${purchase.customer_name}`} dir="rtl" onClick={(event) => event.stopPropagation()}>
        <header><h3>{purchase.customer_name}</h3><button type="button" disabled={busy} title="إغلاق" aria-label="إغلاق الدفعات" onClick={() => setOpen(false)}><X size={20} /></button></header>
        <p>{active.pickup_point}</p>
        <dl><div><dt>السعر المتفق عليه</dt><dd>{formatILS(agreedAmount(active))} ₪</dd></div><div><dt>المدفوع</dt><dd>{formatILS(paidAmount(active))} ₪</dd></div><div><dt>المتبقي</dt><dd>{formatILS(remainingAmount(active))} ₪</dd></div></dl>
        {error ? <div className="purchase-payment-error" role="alert">{error}</div> : null}
        {loading ? <p role="status">جاري التحميل...</p> : null}
        {state && canPay && remainingAmount(active) > 0 ? <form onSubmit={save}>
          <label>مبلغ الدفعة (₪)<input autoFocus type="number" inputMode="decimal" min="0.01" max={remainingAmount(active)} step="0.01" value={amount} disabled={busy} required
            onChange={(event) => { setAmount(event.target.value); requestRef.current = crypto.randomUUID(); }} /></label>
          <button className="payment-primary" type="submit" disabled={busy}>{busy ? "جاري الحفظ..." : active.picked_up ? "حفظ الدفعة" : "تأكيد الاستلام والدفعة"}</button>
        </form> : null}
        {state && canConfirmPickup && !active.picked_up ? <button className="payment-primary" type="button" disabled={busy} onClick={saveStaffPickup}>
          {busy ? "جاري الحفظ..." : `تأكيد الاستلام واستلام كامل المبلغ (${formatILS(remainingAmount(active))} ₪)`}
        </button> : null}
        <h4>سجل الدفعات</h4>
        <div className="purchase-payment-history">{state?.payments.map((entry) => <div key={entry.id} className={`purchase-payment-entry${entry.cancelled_at ? " cancelled" : ""}`}>
          <div><strong>{formatILS(entry.amount)} ₪</strong><small>{formatDateTime(entry.created_at)}</small></div>
          <span>{entry.cancelled_at ? "ملغاة" : entry.remitted_at ? "في مصاري البيت" : "بانتظار التحصيل"}{entry.legacy ? " · رصيد سابق" : ""}</span>
          {entry.remitted_at ? <small>التحصيل: {formatDateTime(entry.remitted_at)}</small> : null}
          {entry.cancelled_at ? <small>{entry.cancellation_reason} · {formatDateTime(entry.cancelled_at)}</small> : null}
          {role === "rahaf" && !entry.cancelled_at ? <button type="button" disabled={busy} onClick={() => { setCancelId(entry.id); setReason(""); }}>إلغاء الدفعة</button> : null}
        </div>)}</div>
        {state && !state.payments.length ? <p>لا توجد دفعات</p> : null}
        {cancelId ? <div className="purchase-payment-cancel"><label>سبب الإلغاء<input maxLength={500} value={reason} disabled={busy} onChange={(e) => setReason(e.target.value)} /></label>
          <button type="button" disabled={busy || !reason.trim()} onClick={cancelPayment}>تأكيد إلغاء الدفعة</button><button type="button" disabled={busy} onClick={() => setCancelId("")}>تراجع</button></div> : null}
      </section>
    </div>, document.body) : null}
  </div>;
}

export function OutstandingPaymentsSection({ rows, role, instant = false, onSaved }) {
  const pending = rows.filter((row) => row.picked_up && (remainingAmount(row) > 0 || collectibleAmount(row) > 0));
  if (!pending.length) return null;
  return <section className="outstanding-payments" dir="rtl"><h3>متبقي عليه مبلغ</h3>
    {pending.map((purchase) => <div className="outstanding-payment-row" key={purchase.id}><strong>{purchase.customer_name}</strong>
      <PurchasePaymentControl compact purchase={purchase} role={role} instant={instant} onSaved={onSaved} /></div>)}
  </section>;
}
