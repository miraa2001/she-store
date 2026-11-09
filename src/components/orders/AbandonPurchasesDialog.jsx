import { useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import { Archive, LoaderCircle, X } from "lucide-react";
import { archiveAbandonedPurchases, canArchivePurchase } from "../../lib/abandonedPurchases";
import { formatILS } from "../../lib/orders";
import { setBodyScrollLock } from "../../lib/bodyScrollLock";
import "./abandoned-purchases.css";

export default function AbandonPurchasesDialog({ purchases, onClose, onSaved }) {
  const eligible = useMemo(() => purchases.filter(canArchivePurchase), [purchases]);
  const [selected, setSelected] = useState(() => eligible.length === 1 ? [eligible[0].id] : []);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const dialog = useRef(null);
  const request = useRef(null);
  useEffect(() => {
    const node = dialog.current; node.showModal(); setBodyScrollLock("abandon-purchases", true);
    return () => { node.close(); setBodyScrollLock("abandon-purchases", false); };
  }, []);
  function select(ids) { if (request.current || busy) return; setSelected(ids); }
  async function submit(event) {
    event.preventDefault();
    if (busy || !selected.length) return;
    setBusy(true); setError("");
    try {
      request.current ||= crypto.randomUUID();
      await archiveAbandonedPurchases(request.current, selected);
      await onSaved?.(); onClose();
    } catch (e) {
      if (e?.code === "P0001" || e?.code === "42501") request.current = null;
      setError(e?.message || "تعذر نقل المشتريات للأرشيف.");
    } finally { setBusy(false); }
  }
  return createPortal(<dialog ref={dialog} className="abandoned-dialog" dir="rtl" aria-labelledby="abandon-title"
    onCancel={event => { event.preventDefault(); if (!busy) onClose(); }}>
    <header><h2 id="abandon-title">نقل إلى أرشيف المتروكات</h2><button type="button" title="إغلاق" aria-label="إغلاق الأرشيف" disabled={busy} onClick={onClose}><X size={20} /></button></header>
    <form onSubmit={submit}>
      <fieldset disabled={busy || !!request.current}>
        <label className="abandoned-select-all"><input type="checkbox" checked={eligible.length > 0 && selected.length === eligible.length}
          onChange={event => select(event.target.checked ? eligible.map(p => p.id) : [])} />تحديد الكل <b>{selected.length} / {eligible.length}</b></label>
        <div className="abandoned-select-list">{eligible.map(purchase => <label key={purchase.id}>
          <input type="checkbox" checked={selected.includes(purchase.id)} onChange={event => select(event.target.checked ? [...selected, purchase.id] : selected.filter(id => id !== purchase.id))} />
          <span>{purchase.customer_name}</span><b><bdi>{formatILS(purchase.paid_price ?? purchase.price)} ₪</bdi></b>
        </label>)}{!eligible.length ? <p>لا توجد مشتريات غير مستلمة وبدون سجل دفعات.</p> : null}</div>
      </fieldset>
      {error ? <p className="abandoned-error" role="alert">{error}</p> : null}
      <footer><button type="submit" className="primary" disabled={busy || !selected.length}>{busy ? <LoaderCircle size={18} className="instant-pickup-spinner" /> : <Archive size={18} />}{request.current && !busy ? "إعادة المحاولة" : "نقل إلى الأرشيف"}</button><button type="button" disabled={busy} onClick={onClose}>إلغاء</button></footer>
    </form>
  </dialog>, document.body);
}
