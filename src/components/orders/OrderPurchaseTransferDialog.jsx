import { useEffect, useRef, useState } from "react";
import { ArrowRightLeft, X } from "lucide-react";
import { formatILS } from "../../lib/orders";
import { transferPurchasesToOrder } from "../../lib/purchases";
import PurchaseCategoryLabels from "./PurchaseCategoryLabels";

export default function OrderPurchaseTransferDialog({ order, orders, purchases, onClose, onSaved }) {
  const [selected, setSelected] = useState([]);
  const [target, setTarget] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const dialogRef = useRef(null);
  useEffect(() => {
    const previous = document.activeElement;
    dialogRef.current?.focus();
    return () => previous?.focus?.();
  }, []);
  const destinations = orders.filter((item) => item.id !== order.id && !item.arrived && !item.placedAtPickup && !item.allCollected);
  const eligible = purchases.filter((item) => !item.ready_for_pickup && !item.picked_up && !item.collected);

  async function submit(event) {
    event.preventDefault();
    if (busy || !target || !selected.length) return;
    setBusy(true);
    setError("");
    try {
      await transferPurchasesToOrder(order.id, target === "unassigned" ? null : target, selected);
      await onSaved();
    } catch (failure) {
      setError(failure?.message || "تعذر نقل المشتريات");
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="purchase-modal-backdrop" onClick={() => { if (!busy) onClose(); }}>
      <form ref={dialogRef} tabIndex={-1} className="purchase-modal-card order-transfer-modal" role="dialog" aria-modal="true" aria-labelledby="order-transfer-title" onClick={(event) => event.stopPropagation()} onSubmit={submit} onKeyDown={(event) => {
        if (event.key === "Escape" && !busy) onClose();
        if (event.key !== "Tab") return;
        const controls = [...event.currentTarget.querySelectorAll("button:not(:disabled), input:not(:disabled), select:not(:disabled)")];
        const first = controls[0];
        const last = controls[controls.length - 1];
        if (!first) { event.preventDefault(); return; }
        if (event.shiftKey && (document.activeElement === first || document.activeElement === dialogRef.current)) {
          event.preventDefault(); last.focus();
        } else if (!event.shiftKey && (document.activeElement === last || document.activeElement === dialogRef.current)) {
          event.preventDefault(); first.focus();
        }
      }}>
        <div className="purchase-modal-head">
          <h3 id="order-transfer-title">نقل المشتريات</h3>
          <button type="button" className="icon-btn" aria-label="إغلاق" title="إغلاق" disabled={busy} onClick={onClose}><X size={20} /></button>
        </div>
        <div className="order-transfer-body">
          <strong>{order.name}</strong>
          <div className="order-transfer-list">
            {eligible.map((purchase) => (
              <label key={purchase.id} className="order-transfer-row">
                <input type="checkbox" disabled={busy} checked={selected.includes(purchase.id)} onChange={(event) => setSelected((ids) => event.target.checked ? [...ids, purchase.id] : ids.filter((id) => id !== purchase.id))} />
                <span className="order-transfer-customer">{purchase.customer_name}<PurchaseCategoryLabels categories={purchase.purchase_categories} /></span>
                <span>{formatILS(purchase.paid_price ?? purchase.price)} ₪</span>
              </label>
            ))}
            {!eligible.length ? <p>لا توجد مشتريات متاحة للنقل.</p> : null}
          </div>
          <label htmlFor="transfer-target-order">الطلب الجديد</label>
          <select id="transfer-target-order" value={target} onChange={(event) => setTarget(event.target.value)} disabled={busy} required>
            <option value="">اختاري الطلب</option>
            {order.id ? <option value="unassigned">مشتريات غير معيّنة</option> : null}
            {destinations.map((item) => <option key={item.id} value={item.id}>{item.name}</option>)}
          </select>
          {error ? <p role="alert" className="workspace-error">{error}</p> : null}
          <button type="submit" className="btn-primary" disabled={busy || !target || !selected.length}><ArrowRightLeft size={18} /> {busy ? "جارٍ النقل..." : `نقل (${selected.length})`}</button>
        </div>
      </form>
    </div>
  );
}
