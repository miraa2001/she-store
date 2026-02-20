export default function CustomerQuickAddModal({
  open,
  form,
  saving,
  status,
  cityOptions,
  pickupOptions,
  onClose,
  onSubmit,
  onChange,
  Icon
}) {
  if (!open) return null;

  const closeSafely = () => {
    if (saving) return;
    onClose();
  };

  return (
    <div className="purchase-modal-backdrop" onClick={closeSafely}>
      <div className="purchase-modal-card purchase-modal-card-customer" onClick={(event) => event.stopPropagation()}>
        <div className="purchase-modal-head">
          <h3>إضافة زبون جديد</h3>
          <button type="button" className="icon-btn tiny" onClick={closeSafely}>
            <Icon name="close" className="icon" />
          </button>
        </div>

        <form className="purchase-modal-body" onSubmit={onSubmit}>
          <div className="modal-grid-two">
            <label className="customer-field-full">
              <span>اسم الزبون</span>
              <input
                type="text"
                value={form.name}
                onChange={(event) => onChange({ name: event.target.value })}
                placeholder="مثال: سارة أحمد"
                disabled={saving}
                autoFocus
              />
            </label>

            <label>
              <span>رقم الهاتف</span>
              <input
                type="tel"
                value={form.phone}
                onChange={(event) => onChange({ phone: event.target.value })}
                placeholder="9705xxxxxxx"
                disabled={saving}
              />
            </label>

            <label>
              <span>المدينة</span>
              <select
                value={form.city}
                onChange={(event) => onChange({ city: event.target.value })}
                disabled={saving}
              >
                {cityOptions.map((city) => (
                  <option key={city} value={city}>
                    {city}
                  </option>
                ))}
              </select>
            </label>

            <label className="customer-field-full">
              <span>مكان الاستلام</span>
              <select
                value={form.pickup}
                onChange={(event) => onChange({ pickup: event.target.value })}
                disabled={saving}
              >
                {pickupOptions.map((option) => (
                  <option key={option} value={option}>
                    {option}
                  </option>
                ))}
              </select>
            </label>
          </div>

          {status.text ? (
            <div className={`modal-help ${status.isError ? "modal-help-error" : ""}`}>{status.text}</div>
          ) : null}

          <div className="purchase-modal-foot">
            <button type="submit" className="btn-primary" disabled={saving}>
              {saving ? "جاري الحفظ..." : "حفظ الزبون"}
            </button>
            <button type="button" className="btn-ghost-light" onClick={closeSafely} disabled={saving}>
              إلغاء
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
