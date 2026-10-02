import { useEffect, useRef, useState } from "react";
import { ArrowLeft, ArrowRight, LoaderCircle, Plus, RefreshCw, Undo2, Wallet } from "lucide-react";
import { formatDateTime } from "../../lib/dateFormat";
import { formatILS } from "../../lib/orders";
import {
  HOME_CASH_EXPENSE_CATEGORIES, HOME_CASH_PAGE_SIZE, HOME_CASH_SOURCE_LABELS,
  cancelHomeCashExpense, fetchHomeCash, recordHomeCashExpense
} from "../../lib/homeCash";
import SessionLoader from "../common/SessionLoader";

const EMPTY_EXPENSE = { category: "bags", amount: "", note: "" };
const ACTION_LABELS = {
  tracking_started: "بدء التتبع", receipt: "دخل", receipt_adjustment: "تعديل دخل",
  expense: "مصروف / إيداع", expense_reversal: "إلغاء مصروف / إيداع", profit_payout: "تسليم أرباح"
};

export default function HomeCashPanel() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [formError, setFormError] = useState("");
  const [message, setMessage] = useState("");
  const [expense, setExpense] = useState(EMPTY_EXPENSE);
  const [saving, setSaving] = useState(false);
  const [cancellingId, setCancellingId] = useState("");
  const [offset, setOffset] = useState(0);
  const [revision, setRevision] = useState(0);
  const requestId = useRef(null);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError("");
    fetchHomeCash(offset).then((result) => {
      if (!cancelled) setData(result);
    }).catch((loadError) => {
      if (!cancelled) setError(loadError?.message || "تعذر تحميل مصاري البيت.");
    }).finally(() => {
      if (!cancelled) setLoading(false);
    });
    return () => { cancelled = true; };
  }, [offset, revision]);

  function reload() {
    setOffset(0);
    setRevision((value) => value + 1);
  }

  function changeField(field, value) {
    setExpense((previous) => ({ ...previous, [field]: value }));
    setFormError("");
    setMessage("");
  }

  async function saveExpense(event) {
    event.preventDefault();
    if (saving || loading || !data) return;
    setSaving(true);
    setFormError("");
    setMessage("");
    try {
      requestId.current ||= crypto.randomUUID();
      await recordHomeCashExpense({ ...expense, requestId: requestId.current });
      requestId.current = null;
      setExpense((previous) => ({ ...EMPTY_EXPENSE, category: previous.category }));
      setMessage("تم حفظ المصروف");
      reload();
    } catch (saveError) {
      setFormError(saveError?.message || "تعذر حفظ المصروف.");
    } finally {
      setSaving(false);
    }
  }

  async function cancelExpense(entry) {
    if (cancellingId || saving || loading) return;
    if (!window.confirm(`إلغاء هذا المصروف وإعادة ${formatILS(-Number(entry.amount))} ₪ إلى مصاري البيت؟`)) return;
    setCancellingId(entry.id);
    setFormError("");
    setMessage("");
    try {
      await cancelHomeCashExpense(entry.id);
      setMessage("تم إلغاء المصروف");
      reload();
    } catch (cancelError) {
      setFormError(cancelError?.message || "تعذر إلغاء المصروف.");
    } finally {
      setCancellingId("");
    }
  }

  return (
    <section className="home-cash" aria-label="مصاري البيت">
      <div className="home-cash-heading">
        <h2><Wallet size={20} aria-hidden="true" />مصاري البيت</h2>
        <button type="button" className="finance-ledger-icon-btn" onClick={reload}
          disabled={loading || saving || !!cancellingId} title="تحديث" aria-label="تحديث مصاري البيت">
          <RefreshCw size={18} />
        </button>
      </div>
      {error ? <div className="finance-error" role="alert">{error}</div> : null}
      {loading && !data ? <SessionLoader label="جاري تحميل مصاري البيت..." /> : null}
      {data ? (
        <>
          <dl className="home-cash-summary" aria-busy={loading}>
            <div className="home-cash-balance">
              <dt>الرصيد في البيت</dt>
              <dd className={data.summary.balance < 0 ? "is-negative" : ""}>{formatILS(data.summary.balance)} <small>₪</small></dd>
            </div>
          </dl>

          <form className="home-cash-expense-form" onSubmit={saveExpense}>
            <h3>مصروف جديد</h3>
            <fieldset disabled={saving || loading || !!cancellingId}>
              <label>النوع
                <select value={expense.category} onChange={(event) => changeField("category", event.target.value)}>
                  {HOME_CASH_EXPENSE_CATEGORIES.map((category) => <option key={category.value} value={category.value}>{category.label}</option>)}
                </select>
              </label>
              <label>المبلغ المدفوع (₪)
                <input type="number" min="0.01" step="0.01" inputMode="decimal" required value={expense.amount}
                  onChange={(event) => changeField("amount", event.target.value)} />
              </label>
              <label className="home-cash-note-field">ملاحظة
                <input type="text" maxLength={500} value={expense.note}
                  onChange={(event) => changeField("note", event.target.value)} />
              </label>
              <button type="submit" className="home-cash-save-btn">
                {saving ? <LoaderCircle size={18} className="finance-ledger-spinner" /> : <Plus size={18} />}
                حفظ المصروف
              </button>
            </fieldset>
            {formError ? <div className="home-cash-form-error" role="alert">{formError}</div> : null}
            {message ? <div className="home-cash-form-message" role="status">{message}</div> : null}
          </form>

          <div className="home-cash-log-heading">
            <h3>سجل الحركات</h3>
            <div className="home-cash-pagination">
              <button type="button" className="finance-ledger-icon-btn" title="الأحدث" aria-label="الحركات الأحدث"
                disabled={offset === 0 || loading} onClick={() => setOffset((value) => Math.max(0, value - HOME_CASH_PAGE_SIZE))}>
                <ArrowRight size={18} />
              </button>
              <span>{Math.floor(offset / HOME_CASH_PAGE_SIZE) + 1}</span>
              <button type="button" className="finance-ledger-icon-btn" title="الأقدم" aria-label="الحركات الأقدم"
                disabled={!data.hasMore || loading} onClick={() => setOffset((value) => value + HOME_CASH_PAGE_SIZE)}>
                <ArrowLeft size={18} />
              </button>
            </div>
          </div>
          <div className="finance-ledger-scroll" tabIndex={0} role="region" aria-label="سجل حركات مصاري البيت" aria-busy={loading}>
            <table className="home-cash-log-table">
              <thead><tr><th>التاريخ</th><th>الحركة</th><th>المبلغ (₪)</th><th>الرصيد بعدها (₪)</th><th>بواسطة</th><th /></tr></thead>
              <tbody>
                {data.activity.map((entry) => {
                  const category = HOME_CASH_EXPENSE_CATEGORIES.find((item) => item.value === entry.category);
                  return (
                    <tr key={entry.id} className={entry.is_voided ? "home-cash-entry-voided" : ""}>
                      <td className="home-cash-log-date">{formatDateTime(entry.created_at)}</td>
                      <td>
                        <strong>{ACTION_LABELS[entry.kind] || entry.kind}{entry.is_voided ? " · ملغى" : ""}</strong>
                        <div className="finance-muted">{[category?.label, entry.quantity > 1 ? `الكمية: ${entry.quantity}` : "", HOME_CASH_SOURCE_LABELS[entry.source], entry.customer_name, entry.order_name].filter(Boolean).join(" · ")}</div>
                        {entry.note ? <div className="finance-muted">{entry.note}</div> : null}
                      </td>
                      <td className={`home-cash-log-amount ${Number(entry.amount) < 0 ? "is-negative" : ""}`}>
                        {Number(entry.amount) > 0 ? "+" : ""}{formatILS(entry.amount)}
                      </td>
                      <td className={`home-cash-log-amount ${Number(entry.balance_after) < 0 ? "is-negative" : ""}`}>{formatILS(entry.balance_after)}</td>
                      <td className="home-cash-log-actor">{entry.actor_email || "النظام"}</td>
                      <td>
                        {entry.kind === "expense" && !entry.is_voided ? (
                          <button type="button" className="finance-ledger-icon-btn" onClick={() => cancelExpense(entry)}
                            disabled={!!cancellingId || saving || loading} title="إلغاء المصروف" aria-label="إلغاء المصروف">
                            {cancellingId === entry.id ? <LoaderCircle size={18} className="finance-ledger-spinner" /> : <Undo2 size={18} />}
                          </button>
                        ) : null}
                      </td>
                    </tr>
                  );
                })}
                {!data.activity.length ? <tr><td colSpan={6} className="finance-muted">لا توجد حركات</td></tr> : null}
              </tbody>
            </table>
          </div>
        </>
      ) : null}
    </section>
  );
}
