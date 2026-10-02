import { useEffect, useMemo, useState } from "react";
import { ArrowRightLeft, Banknote, LoaderCircle, RefreshCw, Trash2, UserPlus } from "lucide-react";
import { formatDateTime } from "../../lib/dateFormat";
import { formatILS } from "../../lib/orders";
import { isPickupPointRole } from "../../lib/pickup";
import { searchByName } from "../../lib/search";
import { collectInstantPickups, deleteInstantPickup, fetchInstantPickups, INSTANT_PICKUPS_CHANGED, updateInstantPickup } from "../../lib/instantPickups";
import PickupAnimatedCheckbox from "../common/PickupAnimatedCheckbox";
import PickupTransferDialog from "./PickupTransferDialog";
import InstantPickupDialog from "./InstantPickupDialog";
import "./instant-pickups.css";

const pickupDate = (value) => new Intl.DateTimeFormat("en-GB", {
  timeZone: "Asia/Hebron", day: "numeric", month: "numeric", year: "numeric"
}).format(new Date(value));

export default function InstantPickupSection({ pickupPoint, role, title = "", search = "", showCreateAction = true, onDialogChange }) {
  const isRahaf = role === "rahaf";
  const namedSection = isRahaf || role === "reem";
  const canReceive = isRahaf || isPickupPointRole(role);
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [revision, setRevision] = useState(0);
  const [showCollected, setShowCollected] = useState(false);
  const [creating, setCreating] = useState(false);
  const [busy, setBusy] = useState("");
  const [transfer, setTransfer] = useState(null);
  const dialogOpen = creating || !!transfer;

  useEffect(() => {
    if (!dialogOpen) return undefined;
    onDialogChange?.(true);
    return () => onDialogChange?.(false);
  }, [dialogOpen, onDialogChange]);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError("");
    fetchInstantPickups(pickupPoint, showCollected).then((result) => {
      if (!cancelled) setRows(result);
    }).catch((loadError) => {
      if (!cancelled) setError(namedSection ? loadError?.message || "تعذر تحميل المشتريات." : "تعذر تحميل المشتريات.");
    }).finally(() => { if (!cancelled) setLoading(false); });
    return () => { cancelled = true; };
  }, [pickupPoint, showCollected, revision, namedSection]);

  useEffect(() => {
    const refresh = () => setRevision((value) => value + 1);
    window.addEventListener(INSTANT_PICKUPS_CHANGED, refresh);
    window.addEventListener("focus", refresh);
    return () => {
      window.removeEventListener(INSTANT_PICKUPS_CHANGED, refresh);
      window.removeEventListener("focus", refresh);
    };
  }, []);

  const groups = useMemo(() => {
    const filteredRows = searchByName(rows, search, (row) => row.customer_name);
    if (namedSection) return [{ id: "instant", label: title || "استلام فوري", rows: filteredRows }];
    const byDate = new Map();
    filteredRows.forEach((row) => {
      const date = pickupDate(row.created_at);
      if (!byDate.has(date)) byDate.set(date, { id: date, label: date, rows: [] });
      byDate.get(date).rows.push(row);
    });
    return [...byDate.values()];
  }, [namedSection, rows, search, title]);

  async function act(id, action, destination) {
    if (busy || loading) return;
    const previous = rows.find((row) => row.id === id);
    setBusy(id);
    setError("");
    if (action === "receive" || action === "undo_receive") {
      setRows((items) => items.map((row) => row.id === id ? {
        ...row, picked_up: action === "receive",
        picked_up_at: action === "receive" ? new Date().toISOString() : null
      } : row));
    }
    try {
      await updateInstantPickup(id, action, destination);
      setTransfer(null);
    } catch (actionError) {
      if (previous) setRows((items) => items.map((row) => row.id === id ? previous : row));
      setError(actionError?.message || "تعذر حفظ التغيير.");
    } finally { setBusy(""); }
  }

  async function collect(items) {
    const pending = items.filter((row) => row.picked_up && !row.collected);
    if (!pending.length || busy || loading) return;
    const total = pending.reduce((sum, row) => sum + Number(row.price), 0);
    if (!window.confirm(`تأكيد تحصيل ${formatILS(total)} ₪؟`)) return;
    setBusy("collect");
    setError("");
    try { await collectInstantPickups(pending.map((row) => row.id)); }
    catch (collectError) { setError(collectError?.message || "تعذر تحصيل المبلغ."); }
    finally { setBusy(""); }
  }

  async function remove(row) {
    if (!isRahaf || row.picked_up || row.collected || busy || loading) return;
    if (!window.confirm(`حذف المستلم الفوري ${row.customer_name}؟ لا يمكن التراجع عن الحذف.`)) return;
    setBusy(row.id);
    setError("");
    try { await deleteInstantPickup(row.id); }
    catch (deleteError) { setError(deleteError?.message || "تعذر حذف المستلم."); }
    finally { setBusy(""); }
  }

  if (!namedSection && !rows.length && !error) return null;
  return (
    <section className="instant-pickup-section" aria-label={title || (namedSection ? "استلام فوري" : "المشتريات")} aria-busy={loading}>
      {error ? <div className="instant-pickup-error" role="alert">{error}</div> : null}
      {groups.map((group) => (
        <div key={group.id} className="instant-pickup-group">
          <div className="instant-pickup-section-head">
            <h3>{group.label} <small>{group.rows.length}</small></h3>
            <div className="instant-pickup-tools">
              {namedSection ? <div className="instant-pickup-modes" role="group" aria-label="حالة التحصيل">
                <button type="button" aria-pressed={!showCollected} onClick={() => setShowCollected(false)} disabled={!!busy}>بانتظار التحصيل</button>
                <button type="button" aria-pressed={showCollected} onClick={() => setShowCollected(true)} disabled={!!busy}>محصّل</button>
              </div> : null}
              {isRahaf && showCreateAction ? <button type="button" className="instant-pickup-command" onClick={() => setCreating(true)} disabled={!!busy}><UserPlus size={16} />اضافة مستلم فوري</button> : null}
              <button type="button" className="instant-pickup-icon-btn" title="تحديث" aria-label="تحديث المشتريات" disabled={loading || !!busy} onClick={() => setRevision((value) => value + 1)}>
                {loading ? <LoaderCircle size={18} className="instant-pickup-spinner" /> : <RefreshCw size={18} />}
              </button>
              {isRahaf && !showCollected ? <button type="button" className="instant-pickup-command" onClick={() => collect(group.rows)} disabled={loading || !!busy || !group.rows.some((row) => row.picked_up)}><Banknote size={18} />تحصيل الكل</button> : null}
            </div>
          </div>
          <div className="instant-pickup-total">المبلغ للتحصيل: <b>{formatILS(group.rows.filter((row) => row.picked_up && !row.collected).reduce((sum, row) => sum + Number(row.price), 0))} ₪</b></div>
          <div className="instant-pickup-scroll" tabIndex={0} role="region" aria-label={group.label}>
            <table className={`instant-pickup-table ${isRahaf ? "has-actions" : ""}`}>
              <thead><tr><th>الزبون</th><th>السعر (₪)</th><th>الاستلام</th><th>التاريخ</th>{isRahaf ? <th>الإجراءات</th> : null}</tr></thead>
              <tbody>{group.rows.map((row) => <tr key={row.id} data-instant-pickup-id={row.id}>
                <td>{row.customer_name}</td><td className="instant-pickup-money">{formatILS(row.price)}</td>
                <td><div className="instant-pickup-received"><PickupAnimatedCheckbox checked={row.picked_up} disabled={!canReceive || row.collected || !!busy || loading}
                  ariaLabel={`استلام ${row.customer_name}`} onChange={(event) => act(row.id, event.target.checked ? "receive" : "undo_receive")} />
                  <span>{row.collected ? "محصّل" : row.picked_up ? "تم الاستلام" : "غير مستلم"}</span></div></td>
                <td className="instant-pickup-date">{formatDateTime(row.picked_up_at || row.created_at)}</td>
                {isRahaf ? <td><div className="instant-pickup-row-actions">
                  <button type="button" className="instant-pickup-icon-btn danger" title={row.collected ? "لا يمكن حذف مستلم تم تحصيله" : row.picked_up ? "ألغِ الاستلام قبل الحذف" : "حذف المستلم"} aria-label={`حذف ${row.customer_name}`} disabled={row.picked_up || row.collected || !!busy || loading} onClick={() => remove(row)}><Trash2 size={18} /></button>
                  <button type="button" className="instant-pickup-icon-btn" title="تحصيل" aria-label={`تحصيل ${row.customer_name}`} disabled={!row.picked_up || row.collected || !!busy || loading} onClick={() => collect([row])}><Banknote size={18} /></button>
                  <button type="button" className="instant-pickup-icon-btn" title="نقل المشترى" aria-label={`نقل ${row.customer_name}`} disabled={row.collected || !!busy || loading} onClick={() => setTransfer(row)}><ArrowRightLeft size={18} /></button>
                </div></td> : null}
              </tr>)}
              {!group.rows.length ? <tr><td colSpan={isRahaf ? 5 : 4}>{loading ? "جاري التحميل..." : "لا توجد مشتريات"}</td></tr> : null}</tbody>
            </table>
          </div>
        </div>
      ))}
      {creating ? <InstantPickupDialog defaultLocation={pickupPoint} onClose={() => setCreating(false)} /> : null}
      <PickupTransferDialog purchase={transfer} saving={!!busy} onClose={() => { if (!busy) setTransfer(null); }} onTransfer={(destination) => act(transfer.id, "transfer", destination)} />
    </section>
  );
}
