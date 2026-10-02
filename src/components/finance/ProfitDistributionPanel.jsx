import { useEffect, useMemo, useRef, useState } from "react";
import { ArrowLeft, ArrowRight, Banknote, Check, LoaderCircle, RefreshCw } from "lucide-react";
import { formatDateTime } from "../../lib/dateFormat";
import { formatILS } from "../../lib/orders";
import { parseFinanceAmount } from "../../lib/finance";
import { fetchProfitData, PROFIT_PAGE_SIZE, PROFIT_PARTIES, recordProfitPayout } from "../../lib/profits";

const partyLabel = (party) => PROFIT_PARTIES.find((item) => item.value === party)?.label || party;

export default function ProfitDistributionPanel({ orders = [] }) {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [message, setMessage] = useState("");
  const [saving, setSaving] = useState(false);
  const [form, setForm] = useState({ party: "home", amount: "", deductHomeCash: "" });
  const [revision, setRevision] = useState(0);
  const [offset, setOffset] = useState(0);
  const [search, setSearch] = useState("");
  const request = useRef(null);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError("");
    fetchProfitData(offset).then((result) => { if (!cancelled) setData(result); })
      .catch((err) => { if (!cancelled) setError(err?.message || "تعذر تحميل الأرباح."); })
      .finally(() => { if (!cancelled) setLoading(false); });
    return () => { cancelled = true; };
  }, [offset, revision]);

  function reload() { setOffset(0); setRevision((value) => value + 1); }
  const selectedBalance = data?.balances.find((item) => item.party === form.party);

  function change(field, value) {
    // A failed request retains its payload for a safe retry after a lost response.
    if (request.current) return;
    setForm((previous) => ({ ...previous, [field]: value }));
    setMessage("");
  }

  async function submit(event) {
    event.preventDefault();
    if (saving || loading || !data) return;
    if (!window.confirm(`تأكيد تسليم ${form.amount} ₪ إلى ${partyLabel(form.party)}؟`)) return;
    setSaving(true);
    setError("");
    setMessage("");
    try {
      const amount = parseFinanceAmount(form.amount);
      if (amount <= 0) throw new Error("أدخلي مبلغاً أكبر من صفر.");
      request.current ||= { requestId: crypto.randomUUID(), party: form.party, amount, deductHomeCash: form.deductHomeCash === "yes" };
      await recordProfitPayout(request.current);
      request.current = null;
      setForm((previous) => ({ ...previous, amount: "", deductHomeCash: "" }));
      setMessage("تم تسجيل تسليم الأرباح");
      reload();
    } catch (err) {
      // Explicit database validation errors cannot have committed a payment.
      if (err?.code === "P0001" || err?.code === "42501" || String(err?.code || "").startsWith("22")) request.current = null;
      setError(err?.message || "تعذر تسجيل التسليم. أعيدي المحاولة.");
    } finally { setSaving(false); }
  }

  const visibleOrders = useMemo(() => {
    const dates = new Map(orders.map((order) => {
      const created = Date.parse(order.createdAt) || 0;
      const date = Date.parse(order.orderDate);
      return [order.id, { date: Number.isFinite(date) ? date : created, created }];
    }));
    return (data?.orders || [])
      .filter((order) => dates.has(order.order_id) && order.order_name.toLowerCase().includes(search.trim().toLowerCase()))
      .sort((a, b) => {
        const left = dates.get(a.order_id);
        const right = dates.get(b.order_id);
        return right.date - left.date || right.created - left.created || b.order_id.localeCompare(a.order_id);
      });
  }, [data?.orders, orders, search]);
  return <section className="profit-distribution" aria-label="توزيع الأرباح" aria-busy={loading}>
    <div className="home-cash-heading"><h2><Banknote size={20} />توزيع الأرباح</h2>
      <button type="button" className="finance-ledger-icon-btn" title="تحديث" aria-label="تحديث الأرباح" disabled={loading || saving} onClick={reload}><RefreshCw size={18} /></button>
    </div>
    {error ? <div className="finance-error" role="alert">{error}</div> : null}
    {loading && !data ? <div className="finance-loading"><LoaderCircle size={22} className="finance-ledger-spinner" /></div> : null}
    {data ? <>
      <div className="profit-balances">{PROFIT_PARTIES.map((party) => {
        const balance = data.balances.find((item) => item.party === party.value);
        return <section key={party.value} aria-label={`رصيد ${party.label}`}>
          <h3>{party.label}</h3><dl>
            <div><dt>الأرباح المتراكمة (₪)</dt><dd className={Number(balance?.balance) < 0 ? "is-negative" : ""}>{formatILS(balance?.balance)}</dd></div>
            <div><dt>إجمالي الأرباح</dt><dd>{formatILS(balance?.earned)}</dd></div>
            <div><dt>تم تسليمه</dt><dd>{formatILS(balance?.paid)}</dd></div>
          </dl>
        </section>;
      })}</div>
      <form className="profit-payout-form" onSubmit={submit}>
        <h3>تسليم أرباح</h3>
        <fieldset disabled={saving || loading || !!request.current}>
          <label>الطرف<select value={form.party} onChange={(event) => change("party", event.target.value)}>{PROFIT_PARTIES.map((party) => <option key={party.value} value={party.value}>{party.label}</option>)}</select></label>
          <label>المبلغ المسلّم (₪)<input type="number" min="0.01" step="0.01" required value={form.amount} onChange={(event) => change("amount", event.target.value)} /></label>
          <label>خصم من مصاري البيت<select required value={form.deductHomeCash} onChange={(event) => change("deductHomeCash", event.target.value)}><option value="" disabled>اختاري</option><option value="yes">نعم</option><option value="no">لا</option></select></label>
        </fieldset>
        <div className="profit-payout-submit"><span className="finance-muted">الرصيد المتاح: {formatILS(selectedBalance?.balance)} ₪</span>
          <button type="submit" className="finance-btn primary" disabled={saving || loading || !form.amount || !form.deductHomeCash}>{saving ? <LoaderCircle size={18} className="finance-ledger-spinner" /> : <Check size={18} />}{request.current ? "إعادة المحاولة" : "تم التسليم"}</button>
        </div>
        {message ? <div className="finance-ledger-message" role="status">{message}</div> : null}
      </form>
      <div className="profit-orders-heading"><h3>أرباح الطلبيات</h3><input type="search" aria-label="بحث في أرباح الطلبيات" placeholder="بحث عن طلبية..." value={search} onChange={(event) => setSearch(event.target.value)} /></div>
      <div className="finance-ledger-scroll" tabIndex={0} role="region" aria-label="أرباح الطلبيات">
        <table className="profit-table"><thead><tr><th>الطلبية</th><th>التسويق (₪)</th><th>الربح للتوزيع (₪)</th>{PROFIT_PARTIES.map((party) => <th key={party.value}>{party.label} (₪)</th>)}</tr></thead>
          <tbody>{visibleOrders.map((order) => <tr key={order.order_id}><th scope="row">{order.order_name}</th><td>{formatILS(order.marketing_fee)}</td><td>{formatILS(order.distributable_profit)}</td>{PROFIT_PARTIES.map((party) => <td key={party.value}>{order.configured ? <>{formatILS(order[`${party.value}_profit`])}<small>{formatILS(order[`${party.value}_percent`])}%</small>{Number(order[`${party.value}_deduction`]) > 0 ? <small>مخصومات: {formatILS(order[`${party.value}_deduction`])} ₪</small> : null}</> : "غير محدد"}</td>)}</tr>)}
            {!visibleOrders.length ? <tr><td colSpan={6} className="finance-muted">لا توجد طلبيات</td></tr> : null}</tbody>
        </table>
      </div>
      <div className="home-cash-log-heading"><h3>سجل تسليم الأرباح</h3><div className="home-cash-pagination">
        <button className="finance-ledger-icon-btn" title="الأحدث" aria-label="التسليمات الأحدث" disabled={!offset || loading || saving} onClick={() => setOffset((value) => Math.max(0, value - PROFIT_PAGE_SIZE))}><ArrowRight size={18} /></button>
        <span>{Math.floor(offset / PROFIT_PAGE_SIZE) + 1}</span>
        <button className="finance-ledger-icon-btn" title="الأقدم" aria-label="التسليمات الأقدم" disabled={!data.hasMore || loading || saving} onClick={() => setOffset((value) => value + PROFIT_PAGE_SIZE)}><ArrowLeft size={18} /></button>
      </div></div>
      <div className="finance-ledger-scroll" tabIndex={0} role="region" aria-label="سجل تسليم الأرباح">
        <table className="profit-table"><thead><tr><th>التاريخ</th><th>الطرف</th><th>المبلغ (₪)</th><th>خُصم من مصاري البيت</th><th>بواسطة</th></tr></thead><tbody>
          {data.payouts.map((payout) => <tr key={payout.id}><td>{formatDateTime(payout.created_at)}</td><td>{partyLabel(payout.party)}</td><td>{formatILS(payout.amount)}</td><td>{payout.deduct_home_cash ? "نعم" : "لا"}</td><td>{payout.actor_email || "النظام"}</td></tr>)}
          {!data.payouts.length ? <tr><td colSpan={5} className="finance-muted">لا توجد تسليمات</td></tr> : null}
        </tbody></table>
      </div>
    </> : null}
  </section>;
}
