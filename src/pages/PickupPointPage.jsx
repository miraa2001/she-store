import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import { useLocation } from "react-router-dom";
import { formatDMY, formatDateTime } from "../lib/dateFormat";
import { useAuthProfile } from "../hooks/useAuthProfile";
import { usePurchaseCustomerSearch } from "../hooks/usePurchaseCustomerSearch";
import { getOrdersNavItems, isNavHrefActive } from "../lib/navigation";
import { formatILS, parsePrice } from "../lib/orders";
import { buildCollectedMoneyMessage, buildPickupStatusMessage, notifyPickupStatus } from "../lib/pickupNotifications";
import { isAuraPickup, PICKUP_POINT } from "../lib/pickup";
import { setBodyScrollLock } from "../lib/bodyScrollLock";
import { signOutAndRedirect } from "../lib/session";
import { sb } from "../lib/supabaseClient";
import SessionLoader from "../components/common/SessionLoader";
import AppNavIcon from "../components/common/AppNavIcon";
import PickupAnimatedCheckbox from "../components/common/PickupAnimatedCheckbox";
import SheStoreLogo from "../components/common/SheStoreLogo";
import customerHeaderIcon from "../assets/icons/pickup/customer.png";
import priceHeaderIcon from "../assets/icons/pickup/price-ils.png";
import bagHeaderIcon from "../assets/icons/pickup/bag-size.png";
import pickedUpHeaderIcon from "../assets/icons/pickup/picked-up.png";
import pickupTimeHeaderIcon from "../assets/icons/pickup/pickup-time.png";
import "./pickup-common.css";
import "./pickuppoint-page.css";

const AURA_PICKUP_LABEL = `${PICKUP_POINT} (La Aura)`;

function getOrderDateKey(order) {
  return formatDMY(order?.placedAtPickupAt || order?.orderDate || order?.createdAt);
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

export default function PickupPointPage({ embedded = false }) {
  const { profile } = useAuthProfile();
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [orders, setOrders] = useState([]);
  const [selectedItemId, setSelectedItemId] = useState("");
  const [purchases, setPurchases] = useState([]);
  const [loadingOrders, setLoadingOrders] = useState(true);
  const [loadingPurchases, setLoadingPurchases] = useState(false);
  const [error, setError] = useState("");
  const [ordersMenuOpen, setOrdersMenuOpen] = useState(false);
  const [collecting, setCollecting] = useState(false);
  const [collectingAll, setCollectingAll] = useState(false);
  const [allOrdersTotal, setAllOrdersTotal] = useState(0);
  const [search, setSearch] = useState("");
  const [highlightPurchaseId, setHighlightPurchaseId] = useState("");
  const [paidEditor, setPaidEditor] = useState({ id: "", value: "", saving: false });
  const location = useLocation();
  const highlightTimeoutRef = useRef(null);
  const lastLoadedOrderKeyRef = useRef("");

  const isRahaf = profile.role === "rahaf";
  const isLaaura = profile.role === "laaura";
  const canAccess = isRahaf || isLaaura;
  const sidebarLinks = useMemo(
    () => (isRahaf ? getOrdersNavItems(profile.role) : []),
    [isRahaf, profile.role]
  );
  const ordersMenuPortalTarget = typeof document !== "undefined" ? document.body : null;
  const auraSearchPostFilter = useCallback(
    (purchase) => isAuraPickup(purchase.pickup_point),
    []
  );
  const { searchResults, searchLoading, clearSearchResults } = usePurchaseCustomerSearch({
    search,
    orders,
    postFilter: auraSearchPostFilter
  });

  const groupedOrders = useMemo(() => buildOrderGroups(orders), [orders]);
  const selectedOrder = useMemo(() => {
    return orders.find((order) => String(order.id) === String(selectedItemId)) || null;
  }, [orders, selectedItemId]);

  const selectedOrderIds = useMemo(() => {
    if (!selectedOrder) return [];
    return [selectedOrder.id];
  }, [selectedOrder]);

  const visiblePurchases = useMemo(() => purchases.filter((purchase) => !purchase.collected), [purchases]);
  const pickedTotal = useMemo(
    () =>
      visiblePurchases
        .filter((purchase) => purchase.picked_up)
        .reduce((sum, purchase) => sum + parsePrice(purchase.paid_price ?? purchase.price), 0),
    [visiblePurchases]
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
    return () => {
      if (highlightTimeoutRef.current) window.clearTimeout(highlightTimeoutRef.current);
    };
  }, []);

  useEffect(() => {
    const shouldLockBody = ordersMenuOpen || (!embedded && sidebarOpen);
    setBodyScrollLock(`pickuppoint-${embedded ? "embedded" : "page"}`, shouldLockBody);
    return () => {
      setBodyScrollLock(`pickuppoint-${embedded ? "embedded" : "page"}`, false);
    };
  }, [embedded, ordersMenuOpen, sidebarOpen]);

  useEffect(() => {
    setSidebarOpen(false);
    setOrdersMenuOpen(false);
  }, [location.pathname, location.search, location.hash]);

  useEffect(() => {
    setSelectedItemId((prev) => {
      if (prev && orders.some((order) => String(order.id) === String(prev))) return prev;
      return groupedOrders[0]?.orders?.[0]?.id || "";
    });
    if (!orders.length) setPurchases([]);
  }, [groupedOrders, orders]);

  const loadOrders = useCallback(async () => {
    setLoadingOrders(true);
    setError("");
    try {
      const { data: purchaseRows, error: purchasesError } = await sb
        .from("purchases")
        .select("order_id, pickup_point, collected")
        .eq("collected", false);

      if (purchasesError) throw purchasesError;

      const pendingAura = (purchaseRows || []).filter((purchase) => isAuraPickup(purchase.pickup_point));
      const orderIds = Array.from(new Set(pendingAura.map((purchase) => purchase.order_id)));

      if (!orderIds.length) {
        setOrders([]);
        return;
      }

      const { data: orderRows, error: ordersError } = await sb
        .from("orders")
        .select("id, order_name, created_at, order_date, placed_at_pickup_at")
        .in("id", orderIds)
        .eq("arrived", true)
        .order("created_at", { ascending: false });

      if (ordersError) throw ordersError;

      setOrders(
        (orderRows || []).map((order) => ({
          id: order.id,
          orderName: order.order_name || "",
          createdAt: order.created_at,
          orderDate: order.order_date,
          placedAtPickupAt: order.placed_at_pickup_at
        }))
      );
    } catch (err) {
      console.error(err);
      setOrders([]);
      setError("تعذر تحميل الطلبات.");
    } finally {
      setLoadingOrders(false);
    }
  }, []);

  const loadPurchases = useCallback(async (orderIds) => {
    if (!orderIds?.length) {
      setPurchases([]);
      return;
    }

    setLoadingPurchases(true);
    setError("");
    try {
      const { data, error: purchasesError } = await sb
        .from("purchases")
        .select("id, order_id, customer_name, price, paid_price, bag_size, picked_up, picked_up_at, pickup_point, collected")
        .in("order_id", orderIds)
        .eq("collected", false)
        .order("created_at", { ascending: true });

      if (purchasesError) throw purchasesError;

      setPurchases((data || []).filter((purchase) => isAuraPickup(purchase.pickup_point)));
      setPaidEditor({ id: "", value: "", saving: false });
    } catch (err) {
      console.error(err);
      setError("تعذر تحميل مشتريات الطلب.");
      setPurchases([]);
    } finally {
      setLoadingPurchases(false);
    }
  }, []);

  const loadAllOrdersTotal = useCallback(async () => {
    if (!isRahaf) {
      setAllOrdersTotal(0);
      return;
    }

    const { data, error: totalError } = await sb
      .from("purchases")
      .select("paid_price, price, pickup_point")
      .eq("picked_up", true)
      .eq("collected", false);

    if (totalError) {
      console.error(totalError);
      setAllOrdersTotal(0);
      return;
    }

    const total = (data || [])
      .filter((purchase) => isAuraPickup(purchase.pickup_point))
      .reduce((sum, purchase) => sum + parsePrice(purchase.paid_price ?? purchase.price), 0);

    setAllOrdersTotal(total);
  }, [isRahaf]);

  useEffect(() => {
    if (profile.loading || !profile.authenticated || !canAccess) return;
    loadOrders();
  }, [canAccess, loadOrders, profile.authenticated, profile.loading]);

  useEffect(() => {
    if (profile.loading || !profile.authenticated || !isRahaf) return;
    loadAllOrdersTotal();
  }, [isRahaf, loadAllOrdersTotal, profile.authenticated, profile.loading]);

  useEffect(() => {
    const loadKey = selectedOrderIds.map((id) => String(id)).sort().join(",");
    if (!canAccess || !loadKey) {
      lastLoadedOrderKeyRef.current = "";
      return;
    }
    if (lastLoadedOrderKeyRef.current === loadKey) return;
    lastLoadedOrderKeyRef.current = loadKey;
    loadPurchases(selectedOrderIds);
  }, [canAccess, loadPurchases, selectedOrderIds]);

  async function signOut() {
    await signOutAndRedirect();
  }

  async function togglePicked(purchaseId, checked) {
    const target = purchases.find((purchase) => String(purchase.id) === String(purchaseId));
    const payload = checked
      ? { picked_up: true, picked_up_at: new Date().toISOString() }
      : { picked_up: false, picked_up_at: null };

    setPurchases((prev) =>
      prev.map((purchase) => (String(purchase.id) === String(purchaseId) ? { ...purchase, ...payload } : purchase))
    );

    const { error: updateError } = await sb.from("purchases").update(payload).eq("id", purchaseId);
    if (updateError) {
      console.error(updateError);
      await loadPurchases(selectedOrderIds);
      return;
    }

    if (target) {
      await notifyPickupStatus(
        buildPickupStatusMessage({
          picked: payload.picked_up,
          customerName: target.customer_name,
          price: target.price,
          pickupLabel: AURA_PICKUP_LABEL
        })
      );
    }

    await loadAllOrdersTotal();
  }

  async function collectCurrentOrder() {
    if (!isRahaf || !selectedItemId) return;
    const pending = visiblePurchases.filter((purchase) => purchase.picked_up && !purchase.collected);
    if (!pending.length) return;

    const pendingTotal = pending.reduce(
      (sum, purchase) => sum + parsePrice(purchase.paid_price ?? purchase.price),
      0
    );
    const pendingText = formatILS(pendingTotal);
    const ok = window.confirm(`تأكيد تحصيل ${pending.length} مشتريات بمبلغ ${pendingText} ₪؟`);
    if (!ok) return;

    setCollecting(true);
    const ids = pending.map((purchase) => purchase.id);
    const { error: collectError } = await sb
      .from("purchases")
      .update({ collected: true, collected_at: new Date().toISOString() })
      .in("id", ids);

    if (collectError) {
      console.error(collectError);
      setCollecting(false);
      return;
    }
    await notifyPickupStatus(
      buildCollectedMoneyMessage({ pickupLabel: AURA_PICKUP_LABEL, amountText: pendingText })
    );
    await loadOrders();
    await loadAllOrdersTotal();
    setCollecting(false);
  }

  async function collectAllOrders() {
    if (!isRahaf) return;

    const { data, error: totalError } = await sb
      .from("purchases")
      .select("id, paid_price, price, pickup_point")
      .eq("picked_up", true)
      .eq("collected", false);

    if (totalError) {
      console.error(totalError);
      return;
    }

    const auraList = (data || []).filter((purchase) => isAuraPickup(purchase.pickup_point));
    const total = auraList.reduce((sum, purchase) => sum + parsePrice(purchase.paid_price ?? purchase.price), 0);
    if (total <= 0) return;

    const ok = window.confirm(`تأكيد تحصيل كل المبالغ (${auraList.length} مشتريات) بمبلغ ${formatILS(total)} ₪؟`);
    if (!ok) return;

    setCollectingAll(true);
    const ids = auraList.map((purchase) => purchase.id);
    const { error: collectError } = await sb
      .from("purchases")
      .update({ collected: true, collected_at: new Date().toISOString() })
      .in("id", ids);

    if (collectError) {
      console.error(collectError);
      setCollectingAll(false);
      return;
    }
    await notifyPickupStatus(
      buildCollectedMoneyMessage({ pickupLabel: AURA_PICKUP_LABEL, amountText: formatILS(total) })
    );
    await loadOrders();
    await loadAllOrdersTotal();
    setCollectingAll(false);
  }

  function startEditPaid(purchase) {
    if (!isRahaf) return;
    setPaidEditor({
      id: purchase.id,
      value: purchase.paid_price ?? "",
      saving: false
    });
  }

  function cancelEditPaid() {
    setPaidEditor({ id: "", value: "", saving: false });
  }

  async function savePaidPrice() {
    if (!isRahaf || !paidEditor.id) return;
    const raw = String(paidEditor.value ?? "").trim();
    const nextVal = raw === "" ? null : Number(raw);
    if (raw !== "" && (!Number.isFinite(nextVal) || nextVal < 0)) {
      window.alert("السعر المدفوع غير صحيح.");
      return;
    }

    setPaidEditor((prev) => ({ ...prev, saving: true }));
    const { error: updateError } = await sb
      .from("purchases")
      .update({ paid_price: nextVal })
      .eq("id", paidEditor.id);

    if (updateError) {
      console.error(updateError);
      window.alert("فشل حفظ المدفوع.");
      setPaidEditor((prev) => ({ ...prev, saving: false }));
      return;
    }

    setPurchases((prev) =>
      prev.map((purchase) =>
        String(purchase.id) === String(paidEditor.id) ? { ...purchase, paid_price: nextVal } : purchase
      )
    );
    await loadAllOrdersTotal();
    cancelEditPaid();
  }

  function openSearchResult(result) {
    const targetOrder = orders.find((order) => String(order.id) === String(result.order_id));
    if (!targetOrder) return;
    setSelectedItemId(targetOrder.id);

    clearSearchResults();
    setHighlightPurchaseId(result.id);

    if (highlightTimeoutRef.current) window.clearTimeout(highlightTimeoutRef.current);
    highlightTimeoutRef.current = window.setTimeout(() => {
      setHighlightPurchaseId("");
    }, 2500);
  }

  if (profile.loading) {
    return (
      <div className="pickuppoint-page pickuppoint-state" dir="rtl">
        <SessionLoader />
      </div>
    );
  }

  if (!profile.authenticated) {
    return (
      <div className="pickuppoint-page pickuppoint-state" dir="rtl">
        <div className="pickuppoint-note pickuppoint-note-danger">
          <h2>لا توجد جلسة نشطة</h2>
          <p>يلزم تسجيل الدخول أولًا.</p>
          <a href="#/login" className="pickuppoint-link">
            فتح تسجيل الدخول
          </a>
        </div>
      </div>
    );
  }

  if (!canAccess) {
    return (
      <div className="pickuppoint-page pickuppoint-state" dir="rtl">
        <div className="pickuppoint-note pickuppoint-note-danger">
          <h2>لا توجد صلاحية</h2>
          <p>هذه الصفحة متاحة لرهف أو لارا فقط.</p>
          <a href="#/pickup-dashboard" className="pickuppoint-link">
            العودة
          </a>
        </div>
      </div>
    );
  }

  return (
    <div className={`pickuppoint-page ${embedded ? "embedded" : ""}`} dir="rtl">
      {!embedded ? (
        <>
          <div
            className={`pickuppoint-overlay app-sidebar-overlay ${sidebarOpen ? "open" : ""}`}
            onClick={() => setSidebarOpen(false)}
          />
          <aside className={`pickuppoint-sidebar app-sidebar-drawer ${sidebarOpen ? "open" : ""}`}>
            <div className="pickuppoint-sidebar-head app-sidebar-head">
              <div className="app-sidebar-brand">
                <SheStoreLogo className="app-sidebar-logo-link" imageClassName="app-sidebar-logo-img" />
                <b>القائمة</b>
              </div>
              <button
                type="button"
                className="pickuppoint-menu-btn danger app-sidebar-close"
                onClick={() => setSidebarOpen(false)}
              >
                ✕
              </button>
            </div>
            <div className="pickuppoint-sidebar-content app-sidebar-content">
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
      <div className={`pickuppoint-wrap ${embedded ? "pickup-embedded-container" : ""}`}>
        {!embedded ? (
          <div className="pickuppoint-topbar">
            <div className="topbar-brand-with-logo">
              <SheStoreLogo className="topbar-logo-link" imageClassName="topbar-logo-img" />
              <div className="pickuppoint-brand">
                <b>نقطة الاستلام - La Aura</b>
                <div className="pickuppoint-muted">طلبات الاستلام من نقطة الاستلام</div>
              </div>
            </div>
            <button type="button" className="pickuppoint-menu-btn" onClick={() => setSidebarOpen(true)}>
              ☰
            </button>
          </div>
        ) : null}

        <div className="pickuppoint-search-row pickup-section-header">
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
          <input
            value={search}
            onChange={(event) => setSearch(event.target.value)}
            className="pickuppoint-search-box pickup-search-input"
            placeholder="بحث باسم الزبون..."
          />
          {search.trim().length >= 2 ? (
            <span className="pickuppoint-pill pickuppoint-pill--results">
              {searchLoading ? "جاري البحث..." : `${searchResults.length} نتيجة`}
            </span>
          ) : null}
        </div>

        {search.trim().length >= 2 && searchResults.length ? (
          <div className="pickuppoint-search-results">
            {searchResults.map((result) => (
              <button key={result.id} type="button" onClick={() => openSearchResult(result)}>
                <b>{result.customer_name || ""}</b>
                <div className="pickuppoint-muted">
                  {result.orderName}{" "}
                  — السعر: {formatILS(result.price)}
                </div>
              </button>
            ))}
          </div>
        ) : null}

        <div className="pickuppoint-grid pickuppoint-grid--single pickup-two-col-layout">

          <main className="pickuppoint-card pickup-main-pane">
            {isRahaf ? (
              <div className="pickuppoint-global-summary">
                <div className="pickuppoint-global-collect-row">
                  <span className="pickuppoint-pill">اجمالي المبلغ للتحصيل: {formatILS(allOrdersTotal)} ₪</span>
                  <button
                    className="pickuppoint-btn"
                    type="button"
                    disabled={collectingAll || allOrdersTotal <= 0}
                    onClick={collectAllOrders}
                  >
                    {collectingAll ? "جاري التحصيل..." : "تم استلام تحصيل الكل"}
                  </button>
                </div>
              </div>
            ) : null}

            {!selectedOrder ? (
              <div className="pickuppoint-muted pickuppoint-spacer">
                لا يوجد بيانات
                <div className="pickuppoint-refresh-row">
                  <button className="pickuppoint-btn" type="button" onClick={loadOrders}>
                    تحديث
                  </button>
                </div>
              </div>
            ) : (
              <>
                <div className="pickuppoint-row">
                  <div>
                    <b>{selectedOrder.orderName || "طلبية"}</b>
                  </div>
                </div>

                <div className="pickuppoint-order-summary-card">
                  <div className="pickuppoint-order-pills-row">
                    <span className="pickuppoint-pill">عدد المشتريات: {visiblePurchases.length}</span>
                    <span className="pickuppoint-pill">مجموع المستلم: {formatILS(pickedTotal)} ₪</span>
                  </div>
                  {isRahaf ? (
                    <button
                      type="button"
                      className="pickuppoint-btn pickuppoint-order-collect-btn"
                      onClick={collectCurrentOrder}
                      disabled={collecting}
                    >
                      {collecting ? "جاري التحصيل..." : "تم استلام تحصيل المستلمين"}
                    </button>
                  ) : null}
                </div>

                {loadingPurchases ? (
                  <div className="pickuppoint-spacer">
                    <SessionLoader label="جاري تحميل المشتريات..." />
                  </div>
                ) : null}

                {!loadingPurchases ? (
                  <div className="pickuppoint-table-wrap pickup-table-wrap">
                    <table className="pickuppoint-table pickup-table">
                      <thead>
                        <tr>
                          <th>#</th>
                          <th>
                            <span className="pickuppoint-th-label">
                              <img src={customerHeaderIcon} alt="" className="pickuppoint-th-icon" aria-hidden="true" />
                              <span>الزبون</span>
                            </span>
                          </th>
                          <th>
                            <span className="pickuppoint-th-label">
                              <img src={priceHeaderIcon} alt="" className="pickuppoint-th-icon" aria-hidden="true" />
                              <span>السعر</span>
                            </span>
                          </th>
                          {isRahaf ? (
                            <th>
                              <span className="pickuppoint-th-label">
                                <img src={priceHeaderIcon} alt="" className="pickuppoint-th-icon" aria-hidden="true" />
                                <span>المدفوع</span>
                              </span>
                            </th>
                          ) : null}
                          {isRahaf ? <th className="pickuppoint-edit-col" /> : null}
                          <th>
                            <span className="pickuppoint-th-label">
                              <img src={bagHeaderIcon} alt="" className="pickuppoint-th-icon" aria-hidden="true" />
                              <span>حجم الكيس</span>
                            </span>
                          </th>
                          <th>
                            <span className="pickuppoint-th-label">
                              <img src={pickedUpHeaderIcon} alt="" className="pickuppoint-th-icon" aria-hidden="true" />
                              <span>تم الاستلام</span>
                            </span>
                          </th>
                          {isRahaf ? (
                            <th>
                              <span className="pickuppoint-th-label">
                                <img src={pickupTimeHeaderIcon} alt="" className="pickuppoint-th-icon" aria-hidden="true" />
                                <span>تاريخ الاستلام</span>
                              </span>
                            </th>
                          ) : null}
                        </tr>
                      </thead>
                      <tbody>
                        {visiblePurchases.length ? (
                          visiblePurchases.map((purchase, index) => {
                            const isHighlight = highlightPurchaseId && String(highlightPurchaseId) === String(purchase.id);
                            const isEditing = String(paidEditor.id) === String(purchase.id);
                            return (
                              <tr key={purchase.id} className={isHighlight ? "highlight" : ""}>
                                <td>{index + 1}</td>
                                <td>{purchase.customer_name || ""}</td>
                                <td>{formatILS(purchase.paid_price ?? purchase.price)}</td>
                                {isRahaf ? (
                                  <>
                                    <td>
                                      {isEditing ? (
                                        <input
                                          type="number"
                                          min="0"
                                          step="0.01"
                                          value={paidEditor.value}
                                          onChange={(event) =>
                                            setPaidEditor((prev) => ({ ...prev, value: event.target.value }))
                                          }
                                          className="pickuppoint-paid-input pickup-input mini"
                                          onKeyDown={(event) => {
                                            if (event.key === "Enter") savePaidPrice();
                                            if (event.key === "Escape") cancelEditPaid();
                                          }}
                                        />
                                      ) : (
                                        purchase.paid_price ?? "—"
                                      )}
                                    </td>
                                    <td className="pickuppoint-edit-col">
                                      {isEditing ? (
                                        <div className="pickuppoint-edit-actions pickup-edit-actions">
                                          <button
                                            type="button"
                                            className="pickuppoint-btn mini"
                                            onClick={savePaidPrice}
                                            disabled={paidEditor.saving}
                                          >
                                            ✅
                                          </button>
                                          <button
                                            type="button"
                                            className="pickuppoint-btn mini"
                                            onClick={cancelEditPaid}
                                            disabled={paidEditor.saving}
                                          >
                                            ✖
                                          </button>
                                        </div>
                                      ) : (
                                        <button
                                          type="button"
                                          className="pickuppoint-btn mini"
                                          onClick={() => startEditPaid(purchase)}
                                        >
                                          ✏️
                                        </button>
                                      )}
                                    </td>
                                  </>
                                ) : null}
                                <td>{purchase.bag_size || "—"}</td>
                                <td>
                                  <div className="pickuppoint-pick-row pickup-checkbox-wrap">
                                    <PickupAnimatedCheckbox
                                      checked={!!purchase.picked_up}
                                      onChange={(event) => togglePicked(purchase.id, event.target.checked)}
                                      ariaLabel={purchase.picked_up ? "تم الاستلام" : "غير مستلم"}
                                    />
                                    <span className={`pickuppoint-status ${purchase.picked_up ? "success" : "neutral"}`}>
                                      {purchase.picked_up ? "تم الاستلام" : "غير مستلم"}
                                    </span>
                                  </div>
                                </td>
                                {isRahaf ? <td>{formatDateTime(purchase.picked_up_at)}</td> : null}
                              </tr>
                            );
                          })
                        ) : (
                          <tr>
                            <td colSpan={isRahaf ? 8 : 5} className="pickuppoint-muted">
                              لا يوجد مشتريات
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
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
                    <div className="pickuppoint-spacer">
                      <SessionLoader label="جاري تحميل البيانات..." />
                    </div>
                  ) : null}
                  {!loadingOrders && error ? <div className="pickuppoint-error pickuppoint-spacer">{error}</div> : null}

                  {!loadingOrders && !error && !groupedOrders.length ? (
                    <div className="pickuppoint-muted pickuppoint-spacer">
                      لا يوجد بيانات
                      <div className="pickuppoint-refresh-row">
                        <button className="pickuppoint-btn" type="button" onClick={loadOrders}>
                          تحديث
                        </button>
                      </div>
                    </div>
                  ) : null}

                  {!loadingOrders && !error && groupedOrders.length ? (
                    <div className="workspace-list pickuppoint-orders-list pickup-orders-list">
                      {groupedOrders.map((group) => (
                        <section key={group.id} className="group-block">
                          <div className="month-chip">
                            <AppNavIcon name="calendar" className="icon" />
                            <span>{group.label}</span>
                            <b>({group.orders.length})</b>
                          </div>

                          <div className="group-orders">
                            {group.orders.map((order) => {
                              const active = String(order.id) === String(selectedItemId);
                              return (
                                <button
                                  key={order.id}
                                  type="button"
                                  className={`order-row order-row-btn ${active ? "selected" : ""}`}
                                  onClick={() => {
                                    setSelectedItemId(order.id);
                                    setOrdersMenuOpen(false);
                                  }}
                                >
                                  <div className="order-main">
                                    <strong>{order.orderName || "طلبية"}</strong>
                                    <span>{getOrderDateKey(order) || "—"}</span>
                                  </div>

                                  <div className="order-meta">
                                    <small className="status at_pickup">نقطة الاستلام</small>
                                  </div>
                                </button>
                              );
                            })}
                          </div>
                        </section>
                      ))}
                    </div>
                  ) : null}
                </div>
              </aside>
            </div>,
            ordersMenuPortalTarget
          )
        : null}

    </div>
  );
}
