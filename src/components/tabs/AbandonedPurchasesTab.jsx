import { useEffect, useMemo, useRef, useState } from "react";
import { LoaderCircle, RefreshCw, ShoppingBag, Undo2 } from "lucide-react";
import { ABANDONED_PURCHASES_CHANGED, fetchAbandonedPurchases, resellAbandonedPurchase, restoreAbandonedPurchases } from "../../lib/abandonedPurchases";
import { INSTANT_PICKUP_LOCATIONS, INSTANT_PICKUPS_CHANGED } from "../../lib/instantPickups";
import { formatILS } from "../../lib/orders";
import { toPublicImageUrl } from "../../lib/purchases";
import { formatDateTime } from "../../lib/dateFormat";
import { formatPickupFormValue, PICKUP_HOME } from "../../lib/pickup";
import { searchByName } from "../../lib/search";
import InstantPickupDialog from "../pickup/InstantPickupDialog";
import "../orders/abandoned-purchases.css";

const STATES = [{ value: "available", label: "متاح للاستلام الفوري" }, { value: "resold", label: "أعيد بيعه" }, { value: "restored", label: "معاد للطلب" }, { value: "all", label: "الكل" }];
export default function AbandonedPurchasesTab({ search, onDialogChange, onOpenLightbox, onChanged }) {
  const [rows, setRows] = useState([]), [loading, setLoading] = useState(true), [error, setError] = useState("");
  const [status, setStatus] = useState("available"), [revision, setRevision] = useState(0), [selected, setSelected] = useState([]);
  const [busy, setBusy] = useState(false), [reselling, setReselling] = useState(null);
  const restoreRequest = useRef(null);
  const reload = () => setRevision(value => value + 1);
  useEffect(() => {
    let cancelled = false; setLoading(true); setError("");
    fetchAbandonedPurchases().then(data => { if (!cancelled) setRows(data); })
      .catch(e => { if (!cancelled) setError(e?.message || "تعذر تحميل الأرشيف."); })
      .finally(() => { if (!cancelled) setLoading(false); });
    return () => { cancelled = true; };
  }, [revision]);
  useEffect(() => {
    const refresh = () => setRevision(value => value + 1);
    for (const name of [ABANDONED_PURCHASES_CHANGED, INSTANT_PICKUPS_CHANGED, "focus"]) window.addEventListener(name, refresh);
    return () => { for (const name of [ABANDONED_PURCHASES_CHANGED, INSTANT_PICKUPS_CHANGED, "focus"]) window.removeEventListener(name, refresh); };
  }, []);
  useEffect(() => { if (!reselling) return undefined; onDialogChange?.(true); return () => onDialogChange?.(false); }, [reselling, onDialogChange]);
  const visible = useMemo(() => searchByName(rows.filter(row => status === "all" || row.status === status), search,
    row => `${row.purchase_snapshot.customer_name} ${row.order_snapshot.order_name} ${row.resale?.customer_name || ""}`), [rows, status, search]);
  useEffect(() => { setSelected(previous => previous.filter(id => visible.some(row => row.id === id && row.status === "available"))); }, [visible]);
  async function restore(ids) {
    const effectiveIds = restoreRequest.current?.ids || ids;
    if (busy || loading || !effectiveIds.length) return;
    if (!window.confirm(`إعادة ${effectiveIds.length} مشترى إلى الطلب الأصلي؟`)) return;
    setBusy(true); setError("");
    try {
      restoreRequest.current ||= { id: crypto.randomUUID(), ids: effectiveIds };
      await restoreAbandonedPurchases(restoreRequest.current.id, restoreRequest.current.ids);
      restoreRequest.current = null; setSelected([]); reload(); await onChanged?.();
    } catch (e) {
      if (e?.code === "P0001" || e?.code === "42501") restoreRequest.current = null;
      setError(e?.message || "تعذر إعادة المشتريات.");
    } finally { setBusy(false); }
  }
  return <section className="abandoned-stock" aria-label="أرشيف المتروكات" aria-busy={loading}>
    <header className="abandoned-stock-heading"><h2>أرشيف المتروكات</h2><button type="button" title="تحديث الأرشيف" aria-label="تحديث الأرشيف" disabled={loading || busy} onClick={reload}><RefreshCw size={18} /></button></header>
    <div className="abandoned-stock-toolbar"><div className="abandoned-stock-modes" role="tablist" aria-label="حالة المتروكات">
      {STATES.map(item => <button key={item.value} type="button" role="tab" aria-selected={status === item.value} disabled={busy || !!restoreRequest.current}
        onClick={() => { setStatus(item.value); setSelected([]); }}>{item.label}</button>)}
    </div><span>{visible.length} مشترى</span></div>
    {status === "available" ? <div className="abandoned-stock-bulk"><label><input type="checkbox" disabled={loading || busy || !!restoreRequest.current} checked={visible.length > 0 && selected.length === visible.length}
      onChange={event => setSelected(event.target.checked ? visible.map(row => row.id) : [])} />تحديد الكل</label>
      <button type="button" disabled={loading || busy || (!selected.length && !restoreRequest.current)} onClick={() => restore(selected)}><Undo2 size={16} />{restoreRequest.current ? "إعادة المحاولة" : `إعادة المحدد للطلب (${selected.length})`}</button></div> : null}
    {error ? <p className="abandoned-error" role="alert">{error}</p> : null}
    {loading && !rows.length ? <p role="status"><LoaderCircle size={20} className="instant-pickup-spinner" />جاري التحميل...</p> : null}
    <div className="abandoned-stock-list">{visible.map(row => {
      const source = row.purchase_snapshot;
      const images = row.images.map(image => ({ ...image, url: toPublicImageUrl(image.storage_path) }));
      return <article key={row.id} className="abandoned-stock-item" data-archive-id={row.id}>
        <div>{status === "available" ? <input type="checkbox" aria-label={`تحديد ${source.customer_name}`} checked={selected.includes(row.id)} disabled={loading || busy || !!restoreRequest.current}
          onChange={event => setSelected(event.target.checked ? [...selected, row.id] : selected.filter(id => id !== row.id))} /> : null}</div>
        <div className="abandoned-stock-images">{images.length ? images.slice(0, 4).map((image, index) => <button key={image.id} type="button" aria-label={`صورة ${source.customer_name} ${index + 1}`} onClick={() => onOpenLightbox?.(images, index, source.customer_name)}><img src={image.url} alt={`مشترى ${source.customer_name}`} loading="lazy" /></button>) : <span>لا توجد صور</span>}</div>
        <div className="abandoned-stock-details"><h3>{source.customer_name}</h3><span className="abandoned-stock-status">{STATES.find(item => item.value === row.status)?.label}</span>
          <p>{row.order_snapshot.order_name}</p><p>{source.qty || 1} قطع · السعر الأصلي <bdi>{formatILS(source.price)} ₪</bdi> · المتفق عليه <bdi>{formatILS(source.paid_price ?? source.price)} ₪</bdi></p>
          {source.note ? <p>{source.note}</p> : null}
          {row.links.length ? <div className="abandoned-stock-links">{row.links.map((link, index) => <a key={link.id} href={link.url} target="_blank" rel="noreferrer">رابط {index + 1}</a>)}</div> : null}
          <small>الأرشفة: {formatDateTime(row.archived_at)} · <bdi>{row.actor_email || "غير معروف"}</bdi></small>
          {row.resale ? <p className="abandoned-stock-resale">المستلم الجديد: {row.resale.customer_name} · <bdi>{formatILS(row.resale.price)} ₪</bdi> · {row.resale.pickup_point}</p> : null}
        </div>
        {row.status === "available" ? <div className="abandoned-stock-actions"><button type="button" disabled={loading || busy || !!restoreRequest.current} onClick={() => setReselling(row)}><ShoppingBag size={17} />بيع كاستلام فوري</button><button type="button" title="إعادة للطلب الأصلي" aria-label={`إعادة ${source.customer_name} للطلب`} disabled={loading || busy || !!restoreRequest.current} onClick={() => restore([row.id])}><Undo2 size={17} /></button></div> : null}
      </article>;
    })}</div>
    {!loading && !visible.length ? <p className="abandoned-stock-empty">لا توجد مشتريات مطابقة.</p> : null}
    {reselling ? <InstantPickupDialog title="إعادة بيع كاستلام فوري" defaultLocation={INSTANT_PICKUP_LOCATIONS.some(location => location.value === formatPickupFormValue(reselling.purchase_snapshot.pickup_point)) ? formatPickupFormValue(reselling.purchase_snapshot.pickup_point) : PICKUP_HOME}
      onCreate={payload => resellAbandonedPurchase(reselling.id, payload)} onClose={() => setReselling(null)} onSaved={async () => { reload(); await onChanged?.(); }} /> : null}
  </section>;
}
