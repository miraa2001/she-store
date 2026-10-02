import { useEffect, useMemo, useState } from "react";
import { Check, LoaderCircle, RotateCcw, Save, Search } from "lucide-react";
import { formatDMY } from "../../lib/dateFormat";
import { formatILS } from "../../lib/orders";
import { calculateOrderFinanceProfit, parseFinanceAmount, saveOrderFinance } from "../../lib/finance";

function OrderFinanceRow({ order, postalFeeAvailable, onSaved }) {
  const [spent, setSpent] = useState(String(order.spent));
  const [postalFee, setPostalFee] = useState(String(order.postalFee ?? ""));
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState("");
  const [failed, setFailed] = useState(false);
  const dirty = spent !== String(order.spent) || postalFee !== String(order.postalFee ?? "");

  useEffect(() => {
    setSpent(String(order.spent));
    setPostalFee(String(order.postalFee ?? ""));
  }, [order.spent, order.postalFee]);

  let profit = null;
  if (postalFeeAvailable) {
    try {
      profit = calculateOrderFinanceProfit(order.purchaseValue, parseFinanceAmount(postalFee), parseFinanceAmount(spent));
    } catch {
      profit = null;
    }
  }

  function changeValue(setter, value) {
    setter(value);
    setMessage("");
    setFailed(false);
  }

  function reset() {
    setSpent(String(order.spent));
    setPostalFee(String(order.postalFee ?? ""));
    setMessage("");
    setFailed(false);
  }

  async function save() {
    if (saving || !postalFeeAvailable) return;
    setSaving(true);
    setMessage("");
    setFailed(false);
    try {
      const saved = await saveOrderFinance(order.id, { spent, postalFee });
      onSaved(saved);
      setMessage("تم الحفظ");
    } catch (error) {
      setFailed(true);
      setMessage(error?.message || "تعذر الحفظ. حاولي مرة أخرى.");
    } finally {
      setSaving(false);
    }
  }

  return (
    <tr data-order-id={order.id} className={dirty ? "finance-ledger-row-dirty" : ""}>
      <th scope="row" className="finance-ledger-order">{order.name}</th>
      <td className="finance-ledger-date">{formatDMY(order.orderDate || order.createdAt)}</td>
      <td>
        <input
          type="number"
          min="0"
          step="0.01"
          inputMode="decimal"
          className="finance-ledger-input"
          aria-label={`المصروف - ${order.name}`}
          value={spent}
          onChange={(event) => changeValue(setSpent, event.target.value)}
          disabled={saving || !postalFeeAvailable}
        />
      </td>
      <td className="finance-ledger-money">{formatILS(order.purchaseValue)}</td>
      <td>
        <input
          type="number"
          min="0"
          step="0.01"
          inputMode="decimal"
          className="finance-ledger-input"
          aria-label={`رسوم البريد - ${order.name}`}
          value={postalFee}
          onChange={(event) => changeValue(setPostalFee, event.target.value)}
          disabled={saving || !postalFeeAvailable}
        />
      </td>
      <td className={`finance-ledger-money finance-ledger-profit ${profit < 0 ? "is-negative" : ""}`}>
        {profit === null ? "—" : formatILS(profit)}
      </td>
      <td className="finance-ledger-money">{formatILS(order.marketingFee)}</td>
      <td className="finance-ledger-money">{profit === null ? "—" : formatILS(Math.max(0, Math.round((profit - order.marketingFee) * 100) / 100))}</td>
      <td className="finance-ledger-actions-cell">
        <div className="finance-ledger-actions">
          <button
            type="button"
            className="finance-ledger-icon-btn"
            onClick={save}
            disabled={saving || !dirty || !postalFeeAvailable}
            aria-label={`حفظ - ${order.name}`}
            title="حفظ"
          >
            {saving ? <LoaderCircle className="finance-ledger-spinner" size={18} /> : <Save size={18} />}
          </button>
          <button
            type="button"
            className="finance-ledger-icon-btn"
            onClick={reset}
            disabled={saving || !dirty}
            aria-label={`تراجع - ${order.name}`}
            title="تراجع"
          >
            <RotateCcw size={18} />
          </button>
        </div>
        {message ? (
          <div className={`finance-ledger-message ${failed ? "is-error" : ""}`} role={failed ? "alert" : "status"}>
            {!failed ? <Check size={14} aria-hidden="true" /> : null}
            <span>{message}</span>
          </div>
        ) : null}
      </td>
    </tr>
  );
}

export default function OrderFinanceTable({ orders, postalFeeAvailable, onSaved }) {
  const [search, setSearch] = useState("");
  const visibleOrders = useMemo(() => {
    const query = search.trim().toLowerCase();
    return query ? orders.filter((order) => order.name.toLowerCase().includes(query)) : orders;
  }, [orders, search]);
  const totals = useMemo(() => visibleOrders.reduce((sum, order) => ({
    spent: sum.spent + order.spent,
    purchaseValue: sum.purchaseValue + order.purchaseValue,
    postalFee: sum.postalFee + (order.postalFee ?? 0),
    marketingFee: sum.marketingFee + order.marketingFee,
    distributable: sum.distributable + Math.max(0, calculateOrderFinanceProfit(order.purchaseValue, order.postalFee ?? 0, order.spent) - order.marketingFee)
  }), { spent: 0, purchaseValue: 0, postalFee: 0, marketingFee: 0, distributable: 0 }), [visibleOrders]);
  const totalProfit = calculateOrderFinanceProfit(totals.purchaseValue, totals.postalFee, totals.spent);

  return (
    <section className="finance-ledger" aria-label="مالية الطلبيات">
      <div className="finance-ledger-toolbar">
        <div className="finance-ledger-search">
          <Search size={18} aria-hidden="true" />
          <input
            type="search"
            value={search}
            onChange={(event) => setSearch(event.target.value)}
            placeholder="بحث عن طلبية..."
            aria-label="بحث عن طلبية"
          />
        </div>
        <span className="finance-muted">{visibleOrders.length} طلبيات · ₪</span>
      </div>
      {!postalFeeAvailable ? (
        <div className="finance-error" role="alert">رسوم البريد غير متاحة حالياً. تعذر تعديل سجل المالية.</div>
      ) : null}
      <div className="finance-ledger-scroll" tabIndex={0} role="region" aria-label="جدول مالية الطلبيات">
        <table className="finance-ledger-table">
          <thead>
            <tr>
              <th scope="col" className="finance-ledger-order">الطلبية</th>
              <th scope="col">التاريخ</th>
              <th scope="col">المصروف</th>
              <th scope="col">قيمة المشتريات</th>
              <th scope="col">رسوم البريد</th>
              <th scope="col">ربح الطلبية</th>
              <th scope="col">التسويق</th>
              <th scope="col">الربح للتوزيع</th>
              <th scope="col">حفظ</th>
            </tr>
          </thead>
          <tbody>
            {visibleOrders.map((order) => (
              <OrderFinanceRow key={order.id} order={order} postalFeeAvailable={postalFeeAvailable} onSaved={onSaved} />
            ))}
            {!visibleOrders.length ? (
              <tr><td colSpan={9} className="finance-muted">{search ? "لا توجد نتائج" : "لا توجد طلبيات"}</td></tr>
            ) : null}
          </tbody>
          {visibleOrders.length ? (
            <tfoot>
              <tr>
                <th scope="row" colSpan={2}>المجموع</th>
                <td className="finance-ledger-money">{formatILS(totals.spent)}</td>
                <td className="finance-ledger-money">{formatILS(totals.purchaseValue)}</td>
                <td className="finance-ledger-money">{postalFeeAvailable ? formatILS(totals.postalFee) : "—"}</td>
                <td className={`finance-ledger-money finance-ledger-profit ${totalProfit < 0 ? "is-negative" : ""}`}>
                  {postalFeeAvailable ? formatILS(totalProfit) : "—"}
                </td>
                <td className="finance-ledger-money">{formatILS(totals.marketingFee)}</td>
                <td className="finance-ledger-money">{postalFeeAvailable ? formatILS(totals.distributable) : "—"}</td>
                <td />
              </tr>
            </tfoot>
          ) : null}
        </table>
      </div>
    </section>
  );
}
