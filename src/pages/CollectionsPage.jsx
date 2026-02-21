import { useCallback, useEffect, useMemo, useState } from "react";
import { createPortal } from "react-dom";
import { useLocation } from "react-router-dom";
import { useAuthProfile } from "../hooks/useAuthProfile";
import { formatDMY } from "../lib/dateFormat";
import { getOrdersNavItems, isNavHrefActive } from "../lib/navigation";
import { formatILS, parsePrice } from "../lib/orders";
import { isAuraPickup, PICKUP_HOME } from "../lib/pickup";
import { setBodyScrollLock } from "../lib/bodyScrollLock";
import { signOutAndRedirect } from "../lib/session";
import { sb } from "../lib/supabaseClient";
import SessionLoader from "../components/common/SessionLoader";
import AppNavIcon from "../components/common/AppNavIcon";
import SheStoreLogo from "../components/common/SheStoreLogo";
import customerHeaderIcon from "../assets/icons/pickup/customer.png";
import priceHeaderIcon from "../assets/icons/pickup/price-ils.png";
import pickedHeaderIcon from "../assets/icons/pickup/picked-up.png";
import "./pickup-common.css";
import "./collections-page.css";

function getOrderDateKey(order) {
  return formatDMY(order?.createdAt);
}

function buildOrderGroups(orderList) {
  const groups = [];
  const map = new Map();

  orderList.forEach((order) => {
    const dateKey = getOrderDateKey(order) || "غير محدد";
    if (!map.has(dateKey)) {
      const group = {
        id: `group-${dateKey}`,
        dateKey,
        label: dateKey,
        orders: []
      };
      map.set(dateKey, group);
      groups.push(group);
    }
    map.get(dateKey).orders.push(order);
  });

  return groups;
}

export default function CollectionsPage({ embedded = false }) {
  const { profile } = useAuthProfile();
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [orders, setOrders] = useState([]);
  const [selectedOrderId, setSelectedOrderId] = useState("");
  const [ordersMenuOpen, setOrdersMenuOpen] = useState(false);
  const [loadingOrders, setLoadingOrders] = useState(true);
  const [loadingDetails, setLoadingDetails] = useState(false);
  const [error, setError] = useState("");
  const [homeList, setHomeList] = useState([]);
  const [pickupList, setPickupList] = useState([]);
  const location = useLocation();
  const sidebarLinks = useMemo(() => getOrdersNavItems(profile.role), [profile.role]);
  const ordersMenuPortalTarget = typeof document !== "undefined" ? document.body : null;

  const selectedOrder = useMemo(
    () => orders.find((order) => String(order.id) === String(selectedOrderId)) || null,
    [orders, selectedOrderId]
  );
  const groupedOrders = useMemo(() => buildOrderGroups(orders), [orders]);

  const overallCollected = useMemo(
    () => orders.reduce((sum, order) => sum + order.collectedTotal, 0),
    [orders]
  );

  const homeTotal = useMemo(
    () => homeList.reduce((sum, item) => sum + parsePrice(item.paid_price), 0),
    [homeList]
  );

  const pickupTotal = useMemo(
    () => pickupList.reduce((sum, item) => sum + parsePrice(item.paid_price), 0),
    [pickupList]
  );

  const grandTotal = homeTotal + pickupTotal;
  const homePicked = useMemo(() => homeList.filter((purchase) => !!purchase.picked_up), [homeList]);
  const pickupPicked = useMemo(() => pickupList.filter((purchase) => !!purchase.picked_up), [pickupList]);

  const summarizeMethod = useCallback((list) => {
    const pickedCollected = list.filter((purchase) => !!purchase.picked_up && !!purchase.collected);
    const pickedNotCollected = list.filter((purchase) => !!purchase.picked_up && !purchase.collected);
    const notPickedNotCollected = list.filter((purchase) => !purchase.picked_up && !purchase.collected);
    const sumAmount = (rows) =>
      rows.reduce((sum, purchase) => sum + parsePrice(purchase.paid_price ?? purchase.price), 0);

    return {
      pickedCollectedCount: pickedCollected.length,
      pickedCollectedSum: sumAmount(pickedCollected),
      pickedNotCollectedCount: pickedNotCollected.length,
      pickedNotCollectedSum: sumAmount(pickedNotCollected),
      notPickedNotCollectedCount: notPickedNotCollected.length,
      notPickedNotCollectedSum: sumAmount(notPickedNotCollected)
    };
  }, []);

  const homeSummary = useMemo(() => summarizeMethod(homeList), [homeList, summarizeMethod]);
  const pickupSummary = useMemo(() => summarizeMethod(pickupList), [pickupList, summarizeMethod]);

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
    const shouldLockBody = ordersMenuOpen || (!embedded && sidebarOpen);
    setBodyScrollLock(`collections-${embedded ? "embedded" : "page"}`, shouldLockBody);
    return () => {
      setBodyScrollLock(`collections-${embedded ? "embedded" : "page"}`, false);
    };
  }, [embedded, ordersMenuOpen, sidebarOpen]);

  useEffect(() => {
    setSidebarOpen(false);
    setOrdersMenuOpen(false);
  }, [location.pathname, location.search, location.hash]);

  const loadOrders = useCallback(async () => {
    setLoadingOrders(true);
    setError("");
    try {
      const { data, error: ordersError } = await sb
        .from("orders")
        .select(
          "id, order_name, created_at, purchases!inner(id, pickup_point, collected, paid_price, price)"
        )
        .order("created_at", { ascending: false });

      if (ordersError) throw ordersError;

      const filtered = (data || [])
        .filter((order) => {
          const list = order.purchases || [];
          if (!list.length) return false;

          return list.some((p) => p.pickup_point === PICKUP_HOME || isAuraPickup(p.pickup_point));
        })
        .map((order) => {
          const allCollected = (order.purchases || []).every((purchase) => !!purchase.collected);
          const collectedTotal = (order.purchases || []).reduce((sum, purchase) => {
            if (!purchase.collected) return sum;
            return sum + parsePrice(purchase.paid_price);
          }, 0);
          return {
            id: order.id,
            orderName: order.order_name || "",
            createdAt: order.created_at,
            allCollected,
            collectedTotal
          };
        });

      setOrders(filtered);
      setSelectedOrderId((prev) => {
        if (prev && filtered.some((order) => String(order.id) === String(prev))) return prev;
        return filtered[0]?.id || "";
      });
    } catch (err) {
      console.error(err);
      setError("تعذر تحميل الطلبات.");
      setOrders([]);
      setSelectedOrderId("");
    } finally {
      setLoadingOrders(false);
    }
  }, []);

  const loadOrderCollections = useCallback(async (orderId) => {
    if (!orderId) {
      setHomeList([]);
      setPickupList([]);
      return;
    }

    setLoadingDetails(true);
    setError("");
    try {
      const { data, error: purchasesError } = await sb
        .from("purchases")
        .select("id, customer_name, price, paid_price, pickup_point, collected, picked_up")
        .eq("order_id", orderId)
        .order("created_at", { ascending: true });

      if (purchasesError) throw purchasesError;

      const list = data || [];
      setHomeList(list.filter((purchase) => purchase.pickup_point === PICKUP_HOME));
      setPickupList(list.filter((purchase) => isAuraPickup(purchase.pickup_point)));
    } catch (err) {
      console.error(err);
      setError("تعذر تحميل بيانات التحصيل.");
      setHomeList([]);
      setPickupList([]);
    } finally {
      setLoadingDetails(false);
    }
  }, []);

  useEffect(() => {
    if (profile.loading || !profile.authenticated) return;
    if (profile.role !== "rahaf") return;
    loadOrders();
  }, [loadOrders, profile.authenticated, profile.loading, profile.role]);

  useEffect(() => {
    if (!selectedOrderId || profile.role !== "rahaf") return;
    loadOrderCollections(selectedOrderId);
  }, [loadOrderCollections, profile.role, selectedOrderId]);

  async function signOut() {
    await signOutAndRedirect();
  }

  if (profile.loading) {
    return (
      <div className="collections-page collections-state" dir="rtl">
        <SessionLoader />
      </div>
    );
  }

  if (!profile.authenticated) {
    return (
      <div className="collections-page collections-state" dir="rtl">
        <div className="collections-note collections-note-danger">
          <h2>لا توجد جلسة نشطة</h2>
          <p>يلزم تسجيل الدخول أولًا.</p>
          <a href="#/login" className="collections-link">
            فتح تسجيل الدخول
          </a>
        </div>
      </div>
    );
  }

  if (profile.role !== "rahaf") {
    return (
      <div className="collections-page collections-state" dir="rtl">
        <div className="collections-note collections-note-danger">
          <h2>لا توجد صلاحية</h2>
          <p>هذه الصفحة متاحة لحساب رهف فقط.</p>
          <a href="#/orders" className="collections-link">
            العودة للطلبات
          </a>
        </div>
      </div>
    );
  }

  return (
    <div className={`collections-page ${embedded ? "embedded" : ""}`} dir="rtl">
      {!embedded ? (
        <>
          <div
            className={`collections-overlay app-sidebar-overlay ${sidebarOpen ? "open" : ""}`}
            onClick={() => setSidebarOpen(false)}
          />
          <aside className={`collections-sidebar app-sidebar-drawer ${sidebarOpen ? "open" : ""}`}>
            <div className="collections-sidebar-head app-sidebar-head">
              <div className="app-sidebar-brand">
                <SheStoreLogo className="app-sidebar-logo-link" imageClassName="app-sidebar-logo-img" />
                <b>القائمة</b>
              </div>
              <button
                type="button"
                className="collections-menu-btn danger app-sidebar-close"
                onClick={() => setSidebarOpen(false)}
              >
                ✕
              </button>
            </div>
            <div className="collections-sidebar-content app-sidebar-content">
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
        </>
      ) : null}
      <div className="collections-wrap">
        {!embedded ? (
          <div className="collections-topbar">
            <div className="topbar-brand-with-logo">
              <SheStoreLogo className="topbar-logo-link" imageClassName="topbar-logo-img" />
              <div className="collections-brand">
                <b>تحصيل المبالغ</b>
                <div className="collections-muted">طلبات البيت ونقطة الاستلام</div>
              </div>
            </div>
            <button type="button" className="collections-menu-btn" onClick={() => setSidebarOpen(true)}>
              ☰
            </button>
          </div>
        ) : null}

        <div className="collections-orders-menu-row">
          <button
            type="button"
            className="pickup-orders-menu-trigger"
            onClick={() => setOrdersMenuOpen(true)}
            aria-label="فتح قائمة الطلبات"
          >
            <AppNavIcon name="package" className="icon" />
            <span>الطلبات</span>
            <b>{orders.length}</b>
          </button>
          <span className="collections-pill">إجمالي التحصيل: {formatILS(overallCollected)} ₪</span>
        </div>

        <div className="collections-grid collections-grid--single">

          <main className="collections-card">
            {!selectedOrder ? (
              <div className="collections-muted collections-spacer">
                لا يوجد بيانات
                <div className="collections-refresh-row">
                  <button className="collections-btn" type="button" onClick={loadOrders}>
                    تحديث
                  </button>
                </div>
              </div>
            ) : (
              <>
                <div className="collections-row">
                  <div>
                    <b>{selectedOrder.orderName}</b>
                    <span className="collections-status-badge">
                      {selectedOrder.allCollected ? "تم التحصيل" : "قيد التحصيل"}
                    </span>
                  </div>
                  <div className="collections-row collections-top-pills-row">
                    <span className="collections-pill">عدد المشتريات: {homeList.length + pickupList.length}</span>
                    <span className="collections-pill">مجموع البيت: {formatILS(homeTotal)} ₪</span>
                    <span className="collections-pill">مجموع نقطة الاستلام: {formatILS(pickupTotal)} ₪</span>
                    <span className="collections-pill">المجموع الكلي: {formatILS(grandTotal)} ₪</span>
                  </div>
                </div>

                {loadingDetails ? (
                  <div className="collections-spacer">
                    <SessionLoader label="جاري تحميل تفاصيل التحصيل..." />
                  </div>
                ) : null}

                {!loadingDetails ? (
                  <>
                    <section className="collections-method-section">
                      <div className="collections-section-title collections-section-title--with-icon">
                        <AppNavIcon name="home" className="icon" />
                        <span>استلام البيت</span>
                      </div>
                      <div className="collections-method-summary">
                        <div className="collections-method-summary-row">
                          <span className="collections-pill">محصّل ومستلم: {homeSummary.pickedCollectedCount}</span>
                          <span className="collections-pill">المجموع: {formatILS(homeSummary.pickedCollectedSum)} ₪</span>
                        </div>
                        <div className="collections-method-summary-row">
                          <span className="collections-pill">مستلم وغير محصّل: {homeSummary.pickedNotCollectedCount}</span>
                          <span className="collections-pill">المجموع: {formatILS(homeSummary.pickedNotCollectedSum)} ₪</span>
                        </div>
                        <div className="collections-method-summary-row">
                          <span className="collections-pill">غير مستلم وغير محصّل: {homeSummary.notPickedNotCollectedCount}</span>
                          <span className="collections-pill">المجموع: {formatILS(homeSummary.notPickedNotCollectedSum)} ₪</span>
                        </div>
                      </div>
                      <div className="collections-table-wrap">
                        <table className="collections-table">
                          <thead>
                            <tr>
                              <th>
                                <span className="collections-th-label">
                                  <img
                                    src={customerHeaderIcon}
                                    alt=""
                                    aria-hidden="true"
                                    className="collections-th-icon"
                                  />
                                  الزبون
                                </span>
                              </th>
                              <th>
                                <span className="collections-th-label">
                                  <img src={priceHeaderIcon} alt="" aria-hidden="true" className="collections-th-icon" />
                                  المدفوع
                                </span>
                              </th>
                              <th>
                                <span className="collections-th-label">
                                  <img src={pickedHeaderIcon} alt="" aria-hidden="true" className="collections-th-icon" />
                                  حالة التحصيل
                                </span>
                              </th>
                            </tr>
                          </thead>
                          <tbody>
                            {homePicked.length ? (
                              homePicked.map((purchase) => (
                                <tr key={purchase.id}>
                                  <td>{purchase.customer_name || ""}</td>
                                  <td>{formatILS(parsePrice(purchase.paid_price ?? purchase.price))}</td>
                                  <td>{purchase.collected ? "محصّل" : "بانتظار التحصيل"}</td>
                                </tr>
                              ))
                            ) : (
                              <tr>
                                <td colSpan={3} className="collections-muted">
                                  لا يوجد مشتريات مستلمة
                                </td>
                              </tr>
                            )}
                          </tbody>
                        </table>
                      </div>
                    </section>

                    <section className="collections-method-section">
                      <div className="collections-section-title collections-section-title--with-icon">
                        <AppNavIcon name="map" className="icon" />
                        <span>نقطة الاستلام - La Aura</span>
                      </div>
                      <div className="collections-method-summary">
                        <div className="collections-method-summary-row">
                          <span className="collections-pill">محصّل ومستلم: {pickupSummary.pickedCollectedCount}</span>
                          <span className="collections-pill">المجموع: {formatILS(pickupSummary.pickedCollectedSum)} ₪</span>
                        </div>
                        <div className="collections-method-summary-row">
                          <span className="collections-pill">مستلم وغير محصّل: {pickupSummary.pickedNotCollectedCount}</span>
                          <span className="collections-pill">المجموع: {formatILS(pickupSummary.pickedNotCollectedSum)} ₪</span>
                        </div>
                        <div className="collections-method-summary-row">
                          <span className="collections-pill">غير مستلم وغير محصّل: {pickupSummary.notPickedNotCollectedCount}</span>
                          <span className="collections-pill">المجموع: {formatILS(pickupSummary.notPickedNotCollectedSum)} ₪</span>
                        </div>
                      </div>
                      <div className="collections-table-wrap">
                        <table className="collections-table">
                          <thead>
                            <tr>
                              <th>
                                <span className="collections-th-label">
                                  <img
                                    src={customerHeaderIcon}
                                    alt=""
                                    aria-hidden="true"
                                    className="collections-th-icon"
                                  />
                                  الزبون
                                </span>
                              </th>
                              <th>
                                <span className="collections-th-label">
                                  <img src={priceHeaderIcon} alt="" aria-hidden="true" className="collections-th-icon" />
                                  المدفوع
                                </span>
                              </th>
                              <th>
                                <span className="collections-th-label">
                                  <img src={pickedHeaderIcon} alt="" aria-hidden="true" className="collections-th-icon" />
                                  حالة التحصيل
                                </span>
                              </th>
                            </tr>
                          </thead>
                          <tbody>
                            {pickupPicked.length ? (
                              pickupPicked.map((purchase) => (
                                <tr key={purchase.id}>
                                  <td>{purchase.customer_name || ""}</td>
                                  <td>{formatILS(parsePrice(purchase.paid_price ?? purchase.price))}</td>
                                  <td>{purchase.collected ? "محصّل" : "بانتظار التحصيل"}</td>
                                </tr>
                              ))
                            ) : (
                              <tr>
                                <td colSpan={3} className="collections-muted">
                                  لا يوجد مشتريات مستلمة
                                </td>
                              </tr>
                            )}
                          </tbody>
                        </table>
                      </div>
                    </section>
                  </>
                ) : null}
              </>
            )}
          </main>
        </div>
      </div>

      {ordersMenuPortalTarget
        ? createPortal(
            <div
              className={`pickup-orders-menu-overlay ${ordersMenuOpen ? "open" : ""}`}
              onClick={() => setOrdersMenuOpen(false)}
            >
              <aside className="pickup-orders-menu-panel" onClick={(event) => event.stopPropagation()}>
                <div className="pickup-orders-menu-head">
                  <div className="pickup-orders-menu-title">
                    <AppNavIcon name="package" className="icon" />
                    <strong>الطلبات</strong>
                    <b>{orders.length}</b>
                  </div>
                  <button
                    type="button"
                    className="pickup-orders-menu-close"
                    onClick={() => setOrdersMenuOpen(false)}
                    aria-label="إغلاق قائمة الطلبات"
                  >
                    ✕
                  </button>
                </div>

                <div className="pickup-orders-menu-list">
                  {loadingOrders ? (
                    <div className="collections-spacer">
                      <SessionLoader label="جاري تحميل البيانات..." />
                    </div>
                  ) : null}
                  {!loadingOrders && error ? <div className="collections-error collections-spacer">{error}</div> : null}

                  {!loadingOrders && !error && !groupedOrders.length ? (
                    <div className="collections-muted collections-spacer">
                      لا يوجد بيانات
                      <div className="collections-refresh-row">
                        <button className="collections-btn" type="button" onClick={loadOrders}>
                          تحديث
                        </button>
                      </div>
                    </div>
                  ) : null}

                  {!loadingOrders && !error
                    ? groupedOrders.map((group) => (
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
                                    <strong>{order.orderName || "طلبية"}</strong>
                                    <span>{getOrderDateKey(order) || "—"}</span>
                                  </div>
                                  <div className="order-meta">
                                    <b>{formatILS(order.collectedTotal)} ₪</b>
                                  </div>
                                </button>
                              );
                            })}
                          </div>
                        </section>
                      ))
                    : null}
                </div>
              </aside>
            </div>,
            ordersMenuPortalTarget
          )
        : null}

    </div>
  );
}
