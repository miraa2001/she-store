import { useCallback, useEffect, useMemo, useState } from "react";
import { useLocation } from "react-router-dom";
import { useAuthProfile } from "../hooks/useAuthProfile";
import { formatDMY } from "../lib/dateFormat";
import { formatILS, isOlderThanCurrentMonth, parsePrice } from "../lib/orders";
import { getOrdersNavItems, isNavHrefActive } from "../lib/navigation";
import { setBodyScrollLock } from "../lib/bodyScrollLock";
import { signOutAndRedirect } from "../lib/session";
import { sb } from "../lib/supabaseClient";
import SessionLoader from "../components/common/SessionLoader";
import AppNavIcon from "../components/common/AppNavIcon";
import SheStoreLogo from "../components/common/SheStoreLogo";
import customerHeaderIcon from "../assets/icons/pickup/customer.png";
import placeHeaderIcon from "../assets/icons/finance/place.png";
import amountHeaderIcon from "../assets/icons/pickup/price-ils.png";
import "./pickup-common.css";
import "./archive-page.css";

const IMAGE_BUCKET = "purchase-images";

function cleanupMessage(kind, count = 0) {
  if (kind === "none") return "لا يوجد صور للحذف.";
  if (kind === "start") return `جاري حذف ${count} صورة من التخزين...`;
  if (kind === "done") return `تم حذف ${count} صورة من التخزين.`;
  if (kind === "storage-error") return "تعذر حذف بعض الصور من التخزين.";
  if (kind === "db-error") return "تم حذف الصور من التخزين، لكن حدث خطأ في قاعدة البيانات.";
  return "";
}

function formatOrderDate(iso) {
  if (!iso) return "—";
  const parsed = new Date(iso);
  if (Number.isNaN(parsed.getTime())) return "—";
  return parsed.toLocaleDateString("ar", {
    year: "numeric",
    month: "long",
    day: "numeric"
  });
}

function getOrderDateKey(order) {
  return formatDMY(order?.createdAt);
}

function buildOrderGroups(orderList) {
  const groups = [];
  const map = new Map();

  orderList.forEach((order) => {
    const dateKey = getOrderDateKey(order) || "بدون تاريخ";
    if (!map.has(dateKey)) {
      const group = { id: `group-${dateKey}`, dateKey, label: dateKey, orders: [] };
      map.set(dateKey, group);
      groups.push(group);
    }
    map.get(dateKey).orders.push(order);
  });

  return groups;
}

export default function ArchivePage() {
  const { profile } = useAuthProfile();
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [ordersMenuOpen, setOrdersMenuOpen] = useState(false);
  const [orders, setOrders] = useState([]);
  const [selectedOrderId, setSelectedOrderId] = useState("");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [cleanupMsg, setCleanupMsg] = useState("");
  const location = useLocation();
  const sidebarLinks = useMemo(() => getOrdersNavItems(profile.role), [profile.role]);

  const selectedOrder = useMemo(
    () => orders.find((order) => String(order.id) === String(selectedOrderId)) || null,
    [orders, selectedOrderId]
  );
  const groupedOrders = useMemo(() => buildOrderGroups(orders), [orders]);

  const totalArchivedPaid = useMemo(
    () => orders.reduce((sum, order) => sum + (order.totalPaid || 0), 0),
    [orders]
  );

  useEffect(() => {
    const onKeyDown = (event) => {
      if (event.key === "Escape") {
        setSidebarOpen(false);
        setOrdersMenuOpen(false);
      }
    };
    document.addEventListener("keydown", onKeyDown);
    return () => document.removeEventListener("keydown", onKeyDown);
  }, []);

  useEffect(() => {
    const shouldLockBody = sidebarOpen || ordersMenuOpen;
    setBodyScrollLock("archive-page", shouldLockBody);
    return () => {
      setBodyScrollLock("archive-page", false);
    };
  }, [ordersMenuOpen, sidebarOpen]);

  useEffect(() => {
    setSidebarOpen(false);
    setOrdersMenuOpen(false);
  }, [location.pathname, location.search, location.hash]);

  const cleanupArchiveImages = useCallback(async (collectedOrders) => {
    const allPaths = [];
    collectedOrders.forEach((order) => {
      (order.purchases || []).forEach((purchase) => {
        (purchase.purchase_images || []).forEach((img) => {
          if (img?.storage_path) allPaths.push(img.storage_path);
        });
      });
    });

    const uniquePaths = Array.from(new Set(allPaths));
    if (!uniquePaths.length) {
      setCleanupMsg(cleanupMessage("none"));
      return;
    }

    setCleanupMsg(cleanupMessage("start", uniquePaths.length));

    const storageChunk = 100;
    for (let index = 0; index < uniquePaths.length; index += storageChunk) {
      const chunk = uniquePaths.slice(index, index + storageChunk);
      const { error: storageError } = await sb.storage.from(IMAGE_BUCKET).remove(chunk);
      if (storageError) {
        console.error(storageError);
        setCleanupMsg(cleanupMessage("storage-error"));
        return;
      }
    }

    const dbChunk = 500;
    for (let index = 0; index < uniquePaths.length; index += dbChunk) {
      const chunk = uniquePaths.slice(index, index + dbChunk);
      const { error: dbError } = await sb.from("purchase_images").delete().in("storage_path", chunk);
      if (dbError) {
        console.error(dbError);
        setCleanupMsg(cleanupMessage("db-error"));
        return;
      }
    }

    setCleanupMsg(cleanupMessage("done", uniquePaths.length));
  }, []);

  const loadArchive = useCallback(async () => {
    setLoading(true);
    setError("");
    setCleanupMsg("");

    try {
      const { data: purchaseRows, error: purchaseError } = await sb
        .from("purchases")
        .select(
          "id, order_id, customer_name, price, paid_price, pickup_point, collected, purchase_images(id, storage_path)"
        )
        .order("created_at", { ascending: false });

      if (purchaseError) throw purchaseError;

      const purchasesByOrder = new Map();
      (purchaseRows || []).forEach((purchase) => {
        if (!purchasesByOrder.has(purchase.order_id)) {
          purchasesByOrder.set(purchase.order_id, { allCollected: true, purchases: [] });
        }
        const entry = purchasesByOrder.get(purchase.order_id);
        if (!purchase.collected) entry.allCollected = false;
        entry.purchases.push(purchase);
      });

      const { data: orderRows, error: orderError } = await sb
        .from("orders")
        .select("id, order_name, created_at")
        .order("created_at", { ascending: false });

      if (orderError) throw orderError;

      const nextOrders = [];
      const collectedForCleanup = [];

      (orderRows || []).forEach((order) => {
        const entry = purchasesByOrder.get(order.id);
        const purchases = entry?.purchases || [];
        const allCollected = purchases.length > 0 && entry?.allCollected;
        const previousMonth = isOlderThanCurrentMonth(order.created_at);
        if (!allCollected && !previousMonth) return;

        const normalized = {
          id: order.id,
          orderName: String(order.order_name || "").trim() || "طلب بدون اسم",
          createdAt: order.created_at,
          archiveType: allCollected ? "تم تحصيلها" : "طلبات سابقة",
          purchases,
          totalPaid: purchases.reduce(
            (sum, purchase) => sum + parsePrice(purchase.paid_price ?? purchase.price),
            0
          )
        };
        nextOrders.push(normalized);
        if (allCollected) collectedForCleanup.push(normalized);
      });

      setOrders(nextOrders);
      setSelectedOrderId((prev) => {
        if (prev && nextOrders.some((order) => String(order.id) === String(prev))) return prev;
        return nextOrders[0]?.id || "";
      });

      if (!nextOrders.length) {
        setCleanupMsg("لا يوجد أرشيف بعد.");
      } else if (collectedForCleanup.length) {
        await cleanupArchiveImages(collectedForCleanup);
      } else {
        setCleanupMsg("");
      }
    } catch (err) {
      console.error(err);
      setError("تعذر تحميل الأرشيف.");
    } finally {
      setLoading(false);
    }
  }, [cleanupArchiveImages]);

  useEffect(() => {
    if (profile.loading || !profile.authenticated) return;
    if (profile.role !== "rahaf") return;
    loadArchive();
  }, [loadArchive, profile.authenticated, profile.loading, profile.role]);

  async function signOut() {
    await signOutAndRedirect();
  }

  if (profile.loading) {
    return (
      <div className="archive-page archive-state" dir="rtl">
        <SessionLoader />
      </div>
    );
  }

  if (!profile.authenticated) {
    return (
      <div className="archive-page archive-state" dir="rtl">
        <div className="archive-note archive-note-danger">
          <h2>لا توجد جلسة نشطة</h2>
          <p>يلزم تسجيل الدخول أولًا.</p>
          <a href="#/login" className="archive-link">
            فتح تسجيل الدخول
          </a>
        </div>
      </div>
    );
  }

  if (profile.role !== "rahaf") {
    return (
      <div className="archive-page archive-state" dir="rtl">
        <div className="archive-note archive-note-danger">
          <h2>لا توجد صلاحية</h2>
          <p>هذه الصفحة متاحة لحساب رهف فقط.</p>
          <a href="#/orders" className="archive-link">
            العودة للطلبيات
          </a>
        </div>
      </div>
    );
  }

  return (
    <div className="archive-page" dir="rtl">
      <div
        className={`archive-overlay app-sidebar-overlay ${sidebarOpen ? "open" : ""}`}
        onClick={() => setSidebarOpen(false)}
      />
      <aside className={`archive-sidebar app-sidebar-drawer ${sidebarOpen ? "open" : ""}`}>
        <div className="archive-sidebar-head app-sidebar-head">
          <div className="app-sidebar-brand">
            <SheStoreLogo className="app-sidebar-logo-link" imageClassName="app-sidebar-logo-img" />
            <b>القائمة</b>
          </div>
          <button
            type="button"
            className="archive-menu-btn danger app-sidebar-close"
            onClick={() => setSidebarOpen(false)}
          >
            ✕
          </button>
        </div>
        <div className="archive-sidebar-content app-sidebar-content">
          {sidebarLinks.map((item) => (
            <a
              key={item.href}
              href={item.href}
              className={`app-sidebar-link ${isNavHrefActive(item.href, location) ? "active" : ""}`}
              onClick={() => setSidebarOpen(false)}
            >
              <AppNavIcon name={item.icon} className="icon" />
              <span>{item.label}</span>
            </a>
          ))}
          <button type="button" className="danger app-sidebar-link app-sidebar-danger" onClick={signOut}>
            تسجيل خروج
          </button>
        </div>
      </aside>

      <div className="archive-wrap">
        <div className="archive-topbar">
          <div className="topbar-brand-with-logo">
            <SheStoreLogo className="topbar-logo-link" imageClassName="topbar-logo-img" />
            <div className="archive-brand">
              <b>الأرشيف</b>
              <div className="archive-muted">طلبات تم تحصيلها + طلبات أقدم من الشهر الحالي</div>
            </div>
          </div>
          <button type="button" className="archive-menu-btn" onClick={() => setSidebarOpen(true)}>
            ☰
          </button>
        </div>

        <div className="archive-orders-menu-row">
          <button
            type="button"
            className="pickup-orders-menu-trigger"
            onClick={() => setOrdersMenuOpen(true)}
            aria-label="فتح قائمة الطلبيات"
          >
            <AppNavIcon name="package" className="icon" />
            <span>الطلبيات</span>
            <b>{orders.length}</b>
          </button>
          <span className="archive-pill">
            {selectedOrder ? `الطلبية المختارة: ${selectedOrder.orderName}` : "اختر طلبية"}
          </span>
          <span className="archive-pill">إجمالي التحصيل: {formatILS(totalArchivedPaid)} ₪</span>
        </div>

        {cleanupMsg ? <div className="archive-cleanup-msg">{cleanupMsg}</div> : null}

        <div
          className={`pickup-orders-menu-overlay ${ordersMenuOpen ? "open" : ""}`}
          onClick={() => setOrdersMenuOpen(false)}
        >
          <aside className="pickup-orders-menu-panel" onClick={(event) => event.stopPropagation()}>
            <div className="pickup-orders-menu-head">
              <div className="pickup-orders-menu-title">
                <AppNavIcon name="package" className="icon" />
                <strong>الطلبيات</strong>
                <b>{orders.length}</b>
              </div>
              <button
                type="button"
                className="pickup-orders-menu-close"
                onClick={() => setOrdersMenuOpen(false)}
                aria-label="إغلاق قائمة الطلبيات"
              >
                ✕
              </button>
            </div>

            <div className="pickup-orders-menu-list">
              {!loading && !error && !groupedOrders.length ? (
                <div className="archive-muted archive-spacer">
                  لا يوجد بيانات
                  <div className="archive-refresh-row">
                    <button className="archive-btn" type="button" onClick={loadArchive}>
                      تحديث
                    </button>
                  </div>
                </div>
              ) : (
                groupedOrders.map((group) => (
                  <section key={group.id} className="group-block">
                    <div className="month-chip">
                      <AppNavIcon name="calendar" className="icon" />
                      <span>{group.label}</span>
                      <b>({group.orders.length})</b>
                    </div>

                    <div className="group-orders">
                      {group.orders.map((order) => {
                        const active = String(selectedOrderId) === String(order.id);
                        return (
                          <button
                            key={order.id}
                            type="button"
                            className={`order-row order-row-btn ${active ? "selected" : ""}`}
                            onClick={() => {
                              setSelectedOrderId(order.id);
                              setOrdersMenuOpen(false);
                            }}
                          >
                            <div className="order-main">
                              <strong>{order.orderName}</strong>
                              <span>{formatOrderDate(order.createdAt)}</span>
                            </div>
                            <div className="order-meta">
                              <small className="status at_pickup">{order.archiveType}</small>
                              <small className="status at_pickup">{formatILS(order.totalPaid)} ₪</small>
                            </div>
                          </button>
                        );
                      })}
                    </div>
                  </section>
                ))
              )}
            </div>
          </aside>
        </div>

        <div className="archive-main-shell">
          {loading ? (
            <div className="archive-card archive-spacer">
              <SessionLoader label="جاري تحميل البيانات..." />
            </div>
          ) : null}
          {error ? <div className="archive-error archive-spacer">{error}</div> : null}

          {!loading && !error ? (
            <main className="archive-card">
              {!selectedOrder ? (
                <div className="archive-muted archive-spacer">
                  لا يوجد بيانات
                  <div className="archive-refresh-row">
                    <button className="archive-btn" type="button" onClick={loadArchive}>
                      تحديث
                    </button>
                  </div>
                </div>
              ) : (
                <>
                  <div className="archive-row">
                    <div>
                      <b>{selectedOrder.orderName}</b>
                      <div className="archive-muted">{formatOrderDate(selectedOrder.createdAt)}</div>
                    </div>
                    <div className="archive-row">
                      <span className="archive-pill">{selectedOrder.archiveType}</span>
                      <span className="archive-pill">عدد المشتريات: {selectedOrder.purchases.length}</span>
                      <span className="archive-pill">المجموع الكلي: {formatILS(selectedOrder.totalPaid)} ₪</span>
                    </div>
                  </div>

                  <hr className="archive-divider" />

                  <div className="archive-table-shell">
                    <div className="archive-table-wrap">
                      <table className="archive-table">
                        <thead>
                          <tr>
                            <th>
                              <span className="archive-th-label">
                                <img src={customerHeaderIcon} alt="" className="archive-th-icon" aria-hidden="true" />
                                <span>الزبون</span>
                              </span>
                            </th>
                            <th>
                              <span className="archive-th-label">
                                <img src={placeHeaderIcon} alt="" className="archive-th-icon" aria-hidden="true" />
                                <span>المكان</span>
                              </span>
                            </th>
                            <th>
                              <span className="archive-th-label">
                                <img src={amountHeaderIcon} alt="" className="archive-th-icon" aria-hidden="true" />
                                <span>السعر</span>
                              </span>
                            </th>
                            <th>
                              <span className="archive-th-label">
                                <img src={amountHeaderIcon} alt="" className="archive-th-icon" aria-hidden="true" />
                                <span>المدفوع</span>
                              </span>
                            </th>
                          </tr>
                        </thead>
                        <tbody>
                          {selectedOrder.purchases.length ? (
                            selectedOrder.purchases.map((purchase) => (
                              <tr key={purchase.id}>
                                <td>{purchase.customer_name || ""}</td>
                                <td>{purchase.pickup_point || ""}</td>
                                <td>{formatILS(parsePrice(purchase.price))} ₪</td>
                                <td>
                                  {purchase.paid_price === null || purchase.paid_price === undefined || purchase.paid_price === ""
                                    ? "—"
                                    : `${formatILS(parsePrice(purchase.paid_price))} ₪`}
                                </td>
                              </tr>
                            ))
                          ) : (
                            <tr>
                              <td colSpan={4} className="archive-muted">
                                لا يوجد مشتريات
                              </td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </>
              )}
            </main>
          ) : null}
        </div>
      </div>
    </div>
  );
}
