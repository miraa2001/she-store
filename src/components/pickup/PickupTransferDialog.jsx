import { formatPickupDisplayLabel } from "../../lib/pickup";
import { getTransferPickupOptions, normalizeOperationalPickupPoint } from "../../lib/purchases";

export default function PickupTransferDialog({
  purchase,
  saving = false,
  onClose,
  onTransfer
}) {
  if (!purchase) return null;

  const currentPickup = String(purchase.pickup_point || "").trim();
  const normalizedCurrentPickup = normalizeOperationalPickupPoint(currentPickup);
  const options = getTransferPickupOptions(currentPickup);

  return (
    <div className="purchase-modal-backdrop" onClick={onClose}>
      <div className="purchase-modal-card pickup-transfer-modal" onClick={(event) => event.stopPropagation()}>
        <div className="purchase-modal-head">
          <h3>نقل المشترى</h3>
          <button type="button" className="icon-btn tiny" onClick={onClose} disabled={saving} aria-label="إغلاق">
            ✕
          </button>
        </div>

        <div className="purchase-modal-body">
          <div className="pickup-transfer-summary">
            <strong>{purchase.customer_name || "بدون اسم"}</strong>
            <span>المكان الحالي: {formatPickupDisplayLabel(currentPickup, "غير محدد")}</span>
          </div>

          <div className="pickup-transfer-options">
            {options.map((option) => {
              const isCurrent = option.value === normalizedCurrentPickup;
              return (
                <button
                  key={option.value}
                  type="button"
                  className={`pickup-transfer-option ${isCurrent ? "is-current" : ""}`}
                  onClick={() => onTransfer?.(option.value)}
                  disabled={saving || isCurrent}
                >
                  <span>{option.label}</span>
                  {isCurrent ? <small>المكان الحالي</small> : null}
                </button>
              );
            })}
          </div>

          <div className="purchase-modal-foot">
            <button type="button" className="btn-ghost-light" onClick={onClose} disabled={saving}>
              إلغاء
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
