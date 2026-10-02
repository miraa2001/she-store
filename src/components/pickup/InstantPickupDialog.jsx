import { useEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import { LoaderCircle, UserPlus, X } from "lucide-react";
import { setBodyScrollLock } from "../../lib/bodyScrollLock";
import { createInstantPickup, INSTANT_PICKUP_LOCATIONS } from "../../lib/instantPickups";
import { PICKUP_HOME } from "../../lib/pickup";
import "./instant-pickups.css";

export default function InstantPickupDialog({ defaultLocation = PICKUP_HOME, onClose, onSaved }) {
  const dialogRef = useRef(null);
  const requestId = useRef(crypto.randomUUID());
  const [customerName, setCustomerName] = useState("");
  const [price, setPrice] = useState("");
  const [pickupPoint, setPickupPoint] = useState(defaultLocation);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    const dialog = dialogRef.current;
    dialog.showModal();
    dialog.querySelector("input")?.focus();
    setBodyScrollLock("instant-pickup-dialog", true);
    return () => {
      dialog.close();
      setBodyScrollLock("instant-pickup-dialog", false);
    };
  }, []);

  async function save(event) {
    event.preventDefault();
    if (saving) return;
    setSaving(true);
    setError("");
    try {
      const entry = await createInstantPickup({ requestId: requestId.current, customerName, price, pickupPoint });
      onSaved?.(entry);
      onClose();
    } catch (saveError) {
      setError(saveError?.message || "تعذر حفظ المستلم.");
    } finally {
      setSaving(false);
    }
  }

  return createPortal(
    <dialog ref={dialogRef} className="instant-pickup-dialog" dir="rtl" aria-labelledby="instant-pickup-title"
      onCancel={(event) => { event.preventDefault(); if (!saving) onClose(); }}
      onClick={(event) => {
        if (saving || event.target !== event.currentTarget) return;
        const rect = event.currentTarget.getBoundingClientRect();
        if (event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom) onClose();
      }}>
      <div className="instant-pickup-dialog-head">
        <h2 id="instant-pickup-title">اضافة مستلم فوري</h2>
        <button type="button" className="instant-pickup-icon-btn" onClick={onClose} disabled={saving} title="إغلاق" aria-label="إغلاق"><X size={18} /></button>
      </div>
      <form onSubmit={save}>
        <fieldset disabled={saving}>
          <label>اسم الزبون<input autoFocus type="text" maxLength={200} required value={customerName} onChange={(event) => setCustomerName(event.target.value)} /></label>
          <label>السعر (₪)<input type="number" min="0" step="0.01" inputMode="decimal" required value={price} onChange={(event) => setPrice(event.target.value)} /></label>
          <label>نقطة الاستلام<select value={pickupPoint} onChange={(event) => setPickupPoint(event.target.value)}>
            {INSTANT_PICKUP_LOCATIONS.map((location) => <option key={location.value} value={location.value}>{location.label}</option>)}
          </select></label>
          {error ? <div className="instant-pickup-error" role="alert">{error}</div> : null}
          <div className="instant-pickup-dialog-actions">
            <button type="submit" className="instant-pickup-command primary">{saving ? <LoaderCircle size={18} className="instant-pickup-spinner" /> : <UserPlus size={18} />}حفظ</button>
            <button type="button" className="instant-pickup-command" onClick={onClose}>إلغاء</button>
          </div>
        </fieldset>
      </form>
    </dialog>, document.body
  );
}
