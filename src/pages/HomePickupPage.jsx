import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useLocation } from "react-router-dom";
import { formatDMY, formatDateTime } from "../lib/dateFormat";
import { useAuthProfile } from "../hooks/useAuthProfile";
import { usePurchaseCustomerSearch } from "../hooks/usePurchaseCustomerSearch";
import { getOrdersNavItems, isNavHrefActive } from "../lib/navigation";
import { formatILS, parsePrice } from "../lib/orders";
import { buildCollectedMoneyMessage, buildPickupStatusMessage, notifyPickupStatus } from "../lib/pickupNotifications";
import { PICKUP_HOME } from "../lib/pickup";
import { setBodyScrollLock } from "../lib/bodyScrollLock";
import { signOutAndRedirect } from "../lib/session";
import { sb } from "../lib/supabaseClient";
import SessionLoader from "../components/common/SessionLoader";
import AppNavIcon from "../components/common/AppNavIcon";
import PickupAnimatedCheckbox from "../components/common/PickupAnimatedCheckbox";
import SheStoreLogo from "../components/common/SheStoreLogo";
import imagesHeaderIcon from "../assets/icons/pickup/images.png";
import customerHeaderIcon from "../assets/icons/pickup/customer.png";
import priceHeaderIcon from "../assets/icons/pickup/price-ils.png";
import pickedHeaderIcon from "../assets/icons/pickup/picked-up.png";
import pickupTimeHeaderIcon from "../assets/icons/pickup/pickup-time.png";
import "./pickup-common.css";
import "./homepickup-page.css";

const BUCKET = "purchase-images";

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

export default function HomePickupPage({ embedded = false }) {
  const { profile } = useAuthProfile();
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [viewMode, setViewMode] = useState("table");
  const [isDesktop, setIsDesktop] = useState(() =>
    typeof window === "undefined" ? true : window.matchMedia("(min-width: 1024px)").matches
  );
  const [orders, setOrders] = useState([]);
  const [selectedOrderId, setSelectedOrderId] = useState("");
  const [purchases, setPurchases] = useState([]);
  const [loadingOrders, setLoadingOrders] = useState(true);
  const [loadingPurchases, setLoadingPurchases] = useState(false);
  const [error, setError] = useState("");
  const [collecting, setCollecting] = useState(false);
  const [ordersMenuOpen, setOrdersMenuOpen] = useState(false);
  const [search, setSearch] = useState("");
  const [highlightPurchaseId, setHighlightPurchaseId] = useState("");
  const [paidEditor, setPaidEditor] = useState({ id: "", value: "", saving: false });
  const [lightbox, setLightbox] = useState({ open: false, images: [], index: 0, label: "" });
  const location = useLocation();
  const sidebarLinks = useMemo(() => getOrdersNavItems(profile.role), [profile.role]);
  const highlightTimeoutRef = useRef(null);
  const homeSearchQueryBuilder = useCallback(
    (request) => request.eq("pickup_point", PICKUP_HOME),
    []
  );
  const { searchResults, searchLoading, clearSearchResults } = usePurchaseCustomerSearch({
    search,
    orders,
    queryBuilder: homeSearchQueryBuilder
  });

  const isRahaf = profile.role === "rahaf";
  const isReemOrRawand = profile.role === "reem" || profile.role === "rawand";

  const selectedOrder = useMemo(
    () => orders.find((order) => String(order.id) === String(selectedOrderId)) || null,
    [orders, selectedOrderId]
  );
  const groupedOrders = useMemo(() => buildOrderGroups(orders), [orders]);

  const visiblePurchases = useMemo(() => purchases.filter((purchase) => !purchase.collected), [purchases]);
  const pickedTotal = useMemo(
    () =>
      visiblePurchases
        .filter((purchase) => purchase.picked_up)
        .reduce((sum, purchase) => sum + parsePrice(purchase.paid_price ?? purchase.price), 0),
    [visiblePurchases]
  );
  const amountToCollect = useMemo(() => pickedTotal, [pickedTotal]);
  const kanbanNotPicked = useMemo(
    () => visiblePurchases.filter((purchase) => !purchase.picked_up),
    [visiblePurchases]
  );
  const kanbanPicked = useMemo(
    () => visiblePurchases.filter((purchase) => purchase.picked_up),
    [visiblePurchases]
  );
  const activeViewMode = isDesktop ? viewMode : "table";

  useEffect(() => {
    const onKeyDown = (event) => {
      if (event.key === "Escape") {
        setSidebarOpen(false);
        setOrdersMenuOpen(false);
        setLightbox((prev) => ({ ...prev, open: false }));
      }
      if (!lightbox.open || !lightbox.images.length) return;
      if (event.key === "ArrowLeft") {
        setLightbox((prev) => ({
          ...prev,
          index: (prev.index - 1 + prev.images.length) % prev.images.length
        }));
      }
      if (event.key === "ArrowRight") {
        setLightbox((prev) => ({
          ...prev,
          index: (prev.index + 1) % prev.images.length
        }));
      }
    };

    document.addEventListener("keydown", onKeyDown);
    return () => document.removeEventListener("keydown", onKeyDown);
  }, [lightbox.images.length, lightbox.open]);

  useEffect(() => {
    if (typeof window === "undefined") return undefined;
    const media = window.matchMedia("(min-width: 1024px)");
    const onChange = (event) => setIsDesktop(event.matches);
    setIsDesktop(media.matches);

    if (typeof media.addEventListener === "function") {
      media.addEventListener("change", onChange);
      return () => media.removeEventListener("change", onChange);
    }

    media.addListener(onChange);
    return () => media.removeListener(onChange);
  }, []);

  useEffect(() => {
    if (!isDesktop) {
      setViewMode("table");
    }
  }, [isDesktop]);

  useEffect(() => {
    return () => {
      if (highlightTimeoutRef.current) window.clearTimeout(highlightTimeoutRef.current);
    };
  }, []);

  useEffect(() => {
    const shouldLockBody = ordersMenuOpen || (!embedded && sidebarOpen);
    setBodyScrollLock(`homepickup-${embedded ? "embedded" : "page"}`, shouldLockBody);
    return () => {
      setBodyScrollLock(`homepickup-${embedded ? "embedded" : "page"}`, false);
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
      const { data: pickupRows, error: pickupError } = await sb
        .from("purchases")
        .select("order_id, pickup_point, collected")
        .eq("pickup_point", PICKUP_HOME)
        .eq("collected", false);

      if (pickupError) throw pickupError;

      const orderIds = Array.from(new Set((pickupRows || []).map((row) => row.order_id)));
      if (!orderIds.length) {
        setOrders([]);
        setSelectedOrderId("");
        setPurchases([]);
        return;
      }

      const { data: orderRows, error: orderError } = await sb
        .from("orders")
        .select("id, order_name, created_at")
        .in("id", orderIds)
        .eq("arrived", true)
        .order("created_at", { ascending: false });

      if (orderError) throw orderError;

      const nextOrders = (orderRows || []).map((order) => ({
        id: order.id,
        orderName: order.order_name || "",
        createdAt: order.created_at
      }));

      setOrders(nextOrders);
      setSelectedOrderId((prev) => {
        if (prev && nextOrders.some((order) => String(order.id) === String(prev))) return prev;
        return nextOrders[0]?.id || "";
      });
    } catch (err) {
      console.error(err);
      setError("تعذر تحميل الطلبات.");
      setOrders([]);
      setSelectedOrderId("");
      setPurchases([]);
    } finally {
      setLoadingOrders(false);
    }
  }, []);

  const loadPurchases = useCallback(async (orderId) => {
    if (!orderId) {
      setPurchases([]);
      return;
    }

    setLoadingPurchases(true);
    setError("");
    try {
      const { data, error: purchasesError } = await sb
        .from("purchases")
        .select(
          "id, customer_name, price, paid_price, picked_up, picked_up_at, pickup_point, collected, purchase_images(storage_path)"
        )
        .eq("order_id", orderId)
        .eq("pickup_point", PICKUP_HOME)
        .eq("collected", false)
        .order("created_at", { ascending: true });

      if (purchasesError) throw purchasesError;

      const mapped = (data || []).map((purchase) => ({
        ...purchase,
        images: (purchase.purchase_images || []).map((img) => {
          const { data: urlData } = sb.storage.from(BUCKET).getPublicUrl(img.storage_path);
          return urlData?.publicUrl || "";
        })
      }));

      setPurchases(mapped);
      setPaidEditor({ id: "", value: "", saving: false });
    } catch (err) {
      console.error(err);
      setError("تعذر تحميل مشتريات الطلب.");
      setPurchases([]);
    } finally {
      setLoadingPurchases(false);
    }
  }, []);

  useEffect(() => {
    if (profile.loading || !profile.authenticated) return;
    if (!isRahaf && !isReemOrRawand) return;
    loadOrders();
  }, [isRahaf, isReemOrRawand, loadOrders, profile.authenticated, profile.loading]);

  useEffect(() => {
    if (!selectedOrderId || (!isRahaf && !isReemOrRawand)) return;
    loadPurchases(selectedOrderId);
  }, [isRahaf, isReemOrRawand, loadPurchases, selectedOrderId]);

  async function signOut() {
    await signOutAndRedirect();
  }

  async function togglePicked(purchaseId, checked) {
    const target = purchases.find((item) => String(item.id) === String(purchaseId));
    const payload = checked
      ? { picked_up: true, picked_up_at: new Date().toISOString() }
      : { picked_up: false, picked_up_at: null };

    setPurchases((prev) =>
      prev.map((item) => (String(item.id) === String(purchaseId) ? { ...item, ...payload } : item))
    );

    const { error: updateError } = await sb.from("purchases").update(payload).eq("id", purchaseId);
    if (updateError) {
      console.error(updateError);
      await loadPurchases(selectedOrderId);
      return;
    }

    if (target) {
      await notifyPickupStatus(
        buildPickupStatusMessage({
          picked: payload.picked_up,
          customerName: target.customer_name,
          price: target.price,
          pickupLabel: PICKUP_HOME
        })
      );
    }
  }

  async function collectHomeMoney() {
    if (!isRahaf || !selectedOrderId) return;
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
    const { error: collectError } = await sb
      .from("purchases")
      .update({ collected: true, collected_at: new Date().toISOString() })
      .eq("order_id", selectedOrderId)
      .eq("pickup_point", PICKUP_HOME)
      .eq("picked_up", true);

    if (collectError) {
      console.error(collectError);
      setCollecting(false);
      return;
    }
    await notifyPickupStatus(
      buildCollectedMoneyMessage({ pickupLabel: PICKUP_HOME, amountText: pendingText })
    );
    await loadPurchases(selectedOrderId);
    await loadOrders();
    setCollecting(false);
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
    cancelEditPaid();
  }

  function openSearchResult(result) {
    clearSearchResults();
    setSelectedOrderId(result.order_id);
    setHighlightPurchaseId(result.id);
    if (highlightTimeoutRef.current) window.clearTimeout(highlightTimeoutRef.current);
    highlightTimeoutRef.current = window.setTimeout(() => {
      setHighlightPurchaseId("");
    }, 2500);
  }

  function openLightbox(images, index, label) {
    if (!images.length) return;
    setLightbox({
      open: true,
      images,
      index,
      label: label || ""
    });
  }

  function renderKanbanPurchaseCard(purchase) {
    const isHighlight = highlightPurchaseId && String(highlightPurchaseId) === String(purchase.id);
    const isEditing = String(paidEditor.id) === String(purchase.id);
    return (
      <article key={purchase.id} className={`homepickup-kanban-card ${isHighlight ? "is-highlight" : ""}`}>
        <header className="homepickup-kanban-card-head">
          <strong>{purchase.customer_name || ""}</strong>
          {isRahaf ? (
            isEditing ? (
              <div className="homepickup-edit-actions pickup-edit-actions">
                <button
                  type="button"
                  className="homepickup-btn mini"
                  onClick={savePaidPrice}
                  disabled={paidEditor.saving}
                >
                  ✓
                </button>
                <button
                  type="button"
                  className="homepickup-btn mini"
                  onClick={cancelEditPaid}
                  disabled={paidEditor.saving}
                >
                  ×
                </button>
              </div>
            ) : (
              <button type="button" className="homepickup-btn mini" onClick={() => startEditPaid(purchase)}>
                ✏️
              </button>
            )
          ) : null}
        </header>

        <div className="homepickup-kanban-meta">
          <span>السعر: {formatILS(parsePrice(purchase.price))} ₪</span>
          {isRahaf ? (
            isEditing ? (
              <input
                type="number"
                min="0"
                step="0.01"
                value={paidEditor.value}
                onChange={(event) => setPaidEditor((prev) => ({ ...prev, value: event.target.value }))}
                className="homepickup-paid-input pickup-input mini"
              />
            ) : (
              <span>المدفوع: {purchase.paid_price === null || purchase.paid_price === undefined || purchase.paid_price === "" ? "—" : `${formatILS(parsePrice(purchase.paid_price))} ₪`}</span>
            )
          ) : null}
        </div>

        {purchase.images?.length ? (
          <div className="homepickup-thumbs homepickup-thumbs-kanban">
            {purchase.images.map((url, index) => (
              <img
                key={`${purchase.id}-kanban-img-${index}`}
                src={url}
                alt="صورة"
                onClick={() => openLightbox(purchase.images, index, purchase.customer_name || "")}
              />
            ))}
          </div>
        ) : null}

        <footer className="homepickup-kanban-actions">
          <div className="homepickup-pick-row pickup-checkbox-wrap">
              <PickupAnimatedCheckbox
                checked={!!purchase.picked_up}
                onChange={(event) => togglePicked(purchase.id, event.target.checked)}
                ariaLabel={purchase.picked_up ? "تم الاستلام" : "غير مستلم"}
              />
              <span>{purchase.picked_up ? "تم الاستلام" : "غير مستلم"}</span>
            </div>
          <small>{formatDateTime(purchase.picked_up_at)}</small>
        </footer>
      </article>
    );
  }

  if (profile.loading) {
    return (
      <div className="homepickup-page homepickup-state" dir="rtl">
        <SessionLoader />
      </div>
    );
  }

  if (!profile.authenticated) {
    return (
      <div className="homepickup-page homepickup-state" dir="rtl">
        <div className="homepickup-note homepickup-note-danger">
          <h2>لا توجد جلسة نشطة</h2>
          <p>يلزم تسجيل الدخول أولًا.</p>
          <a href="#/login" className="homepickup-link">
            فتح تسجيل الدخول
          </a>
        </div>
      </div>
    );
  }

  if (!isRahaf && !isReemOrRawand) {
    return (
      <div className="homepickup-page homepickup-state" dir="rtl">
        <div className="homepickup-note homepickup-note-danger">
          <h2>لا توجد صلاحية</h2>
          <p>هذه الصفحة متاحة لرهف، ريم، وروند فقط.</p>
          <a href="#/pickup-dashboard" className="homepickup-link">
            العودة
          </a>
        </div>
      </div>
    );
  }

  return (
    <div className={`homepickup-page ${embedded ? "embedded" : ""}`} dir="rtl">
      {!embedded ? (
        <>
          <div
            className={`homepickup-overlay app-sidebar-overlay ${sidebarOpen ? "open" : ""}`}
            onClick={() => setSidebarOpen(false)}
          />
          <aside className={`homepickup-sidebar app-sidebar-drawer ${sidebarOpen ? "open" : ""}`}>
            <div className="homepickup-sidebar-head app-sidebar-head">
              <div className="app-sidebar-brand">
                <SheStoreLogo className="app-sidebar-logo-link" imageClassName="app-sidebar-logo-img" />
                <b>القائمة</b>
              </div>
              <button
                type="button"
                className="homepickup-menu-btn danger app-sidebar-close"
                onClick={() => setSidebarOpen(false)}
              >
                ✕
              </button>
            </div>
            <div className="homepickup-sidebar-content app-sidebar-content">
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
      <div className={`homepickup-wrap ${embedded ? "pickup-embedded-container" : ""}`}>
        {!embedded ? (
          <div className="homepickup-topbar">
            <div className="topbar-brand-with-logo">
              <SheStoreLogo className="topbar-logo-link" imageClassName="topbar-logo-img" />
              <div className="homepickup-brand">
                <b>مستلمو البيت</b>
                <div className="homepickup-muted">طلبات الاستلام من البيت</div>
              </div>
            </div>
            <button type="button" className="homepickup-menu-btn" onClick={() => setSidebarOpen(true)}>
              ☰
            </button>
          </div>
        ) : null}

        <div className="homepickup-search-row pickup-section-header">
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
            className="homepickup-search-box pickup-search-input"
            placeholder="بحث باسم الزبون..."
          />
          {search.trim().length >= 2 ? (
            <span className="homepickup-pill">
              {searchLoading ? "..." : `${searchResults.length} نتيجة`}
            </span>
          ) : null}
        </div>

        {search.trim().length >= 2 && searchResults.length ? (
          <div className="homepickup-search-results">
            {searchResults.map((result) => (
              <button key={result.id} type="button" onClick={() => openSearchResult(result)}>
                <b>{result.customer_name || ""}</b>
                <div className="homepickup-muted">
                  {result.orderName} — السعر: {result.price ?? ""}
                </div>
              </button>
            ))}
          </div>
        ) : null}

        <div className="homepickup-grid homepickup-grid--single pickup-two-col-layout">

          <main className="homepickup-card pickup-main-pane">
            {!selectedOrder ? (
              <div className="homepickup-muted homepickup-spacer">
                لا يوجد بيانات
                <div className="homepickup-refresh-row">
                  <button className="homepickup-btn" type="button" onClick={loadOrders}>
                    تحديث
                  </button>
                </div>
              </div>
            ) : (
              <>
                <div className="homepickup-view-controls">
                  {isDesktop ? (
                    <div className="homepickup-view-toggle">
                      <button
                        type="button"
                        className={`homepickup-view-btn ${viewMode === "kanban" ? "active" : ""}`}
                        onClick={() => setViewMode("kanban")}
                      >
                        كانبان
                      </button>
                      <button
                        type="button"
                        className={`homepickup-view-btn ${viewMode === "table" ? "active" : ""}`}
                        onClick={() => setViewMode("table")}
                      >
                        جدول
                      </button>
                    </div>
                  ) : null}
                  <div className="homepickup-amount-display">
                    <span className="homepickup-amount-label">إجمالي المبلغ للتحصيل</span>
                    <strong className="homepickup-amount-value">{formatILS(amountToCollect)} ₪</strong>
                  </div>
                </div>

                <div className="homepickup-row pickup-main-header">
                  <div>
                    <b>{selectedOrder.orderName}</b>
                  </div>
                  <div className="homepickup-row pickup-main-actions">
                    <span className="homepickup-pill">عدد المشتريات: {visiblePurchases.length}</span>
                    {isRahaf ? (
                      <button
                        type="button"
                        className="homepickup-btn"
                        onClick={collectHomeMoney}
                        disabled={collecting}
                      >
                        {collecting ? "جارٍ التحصيل..." : "تم استلام تحصيل الكل"}
                      </button>
                    ) : null}
                  </div>
                </div>

                {loadingPurchases ? (
                  <div className="homepickup-spacer">
                    <SessionLoader label="جاري تحميل المشتريات..." />
                  </div>
                ) : null}

                {!loadingPurchases ? (
                  activeViewMode === "kanban" ? (
                    <div className="homepickup-kanban-grid">
                      <section className="homepickup-kanban-column">
                        <div className="homepickup-kanban-header">
                          <h3>غير مستلمة</h3>
                          <span>{kanbanNotPicked.length}</span>
                        </div>
                        <div className="homepickup-kanban-list">
                          {kanbanNotPicked.length ? (
                            kanbanNotPicked.map((purchase) => renderKanbanPurchaseCard(purchase))
                          ) : (
                            <div className="homepickup-muted">لا توجد مشتريات</div>
                          )}
                        </div>
                      </section>

                      <section className="homepickup-kanban-column homepickup-kanban-column-picked">
                        <div className="homepickup-kanban-header">
                          <h3>تم الاستلام</h3>
                          <span>{kanbanPicked.length}</span>
                        </div>
                        <div className="homepickup-kanban-list">
                          {kanbanPicked.length ? (
                            kanbanPicked.map((purchase) => renderKanbanPurchaseCard(purchase))
                          ) : (
                            <div className="homepickup-muted">لا توجد مشتريات</div>
                          )}
                        </div>
                      </section>
                    </div>
                  ) : (
                    <div className="homepickup-table-wrap pickup-table-wrap">
                      <table className="homepickup-table pickup-table">
                        <thead>
                          <tr>
                            <th>
                              <span className="homepickup-th-label">
                                <img src={customerHeaderIcon} alt="" className="homepickup-th-icon" aria-hidden="true" />
                                <span>الزبون</span>
                              </span>
                            </th>
                            <th>
                              <span className="homepickup-th-label">
                                <img src={priceHeaderIcon} alt="" className="homepickup-th-icon" aria-hidden="true" />
                                <span>السعر</span>
                              </span>
                            </th>
                            {isRahaf ? (
                              <th>
                                <span className="homepickup-th-label">
                                  <img src={priceHeaderIcon} alt="" className="homepickup-th-icon" aria-hidden="true" />
                                  <span>المدفوع</span>
                                </span>
                              </th>
                            ) : null}
                            {isRahaf ? <th className="homepickup-edit-col" /> : null}
                            <th>
                              <span className="homepickup-th-label">
                                <img src={imagesHeaderIcon} alt="" className="homepickup-th-icon" aria-hidden="true" />
                                <span>الصور</span>
                              </span>
                            </th>
                            <th>
                              <span className="homepickup-th-label">
                                <img src={pickedHeaderIcon} alt="" className="homepickup-th-icon" aria-hidden="true" />
                                <span>تم الاستلام</span>
                              </span>
                            </th>
                            <th>
                              <span className="homepickup-th-label">
                                <img src={pickupTimeHeaderIcon} alt="" className="homepickup-th-icon" aria-hidden="true" />
                                <span>وقت الاستلام</span>
                              </span>
                            </th>
                          </tr>
                        </thead>
                        <tbody>
                          {visiblePurchases.length ? (
                            visiblePurchases.map((purchase) => {
                              const isHighlight = highlightPurchaseId && String(highlightPurchaseId) === String(purchase.id);
                              const isEditing = String(paidEditor.id) === String(purchase.id);
                              return (
                                <tr key={purchase.id}>
                                  <td>
                                    {isHighlight ? <div className="homepickup-highlight">✅ نتيجة البحث</div> : null}
                                    {purchase.customer_name || ""}
                                  </td>
                                  <td>{purchase.paid_price ?? purchase.price ?? ""}</td>

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
                                            className="homepickup-paid-input pickup-input mini"
                                          />
                                        ) : (
                                          purchase.paid_price ?? "—"
                                        )}
                                      </td>
                                      <td className="homepickup-edit-col">
                                        {isEditing ? (
                                          <div className="homepickup-edit-actions pickup-edit-actions">
                                            <button
                                              type="button"
                                              className="homepickup-btn mini"
                                              onClick={savePaidPrice}
                                              disabled={paidEditor.saving}
                                            >
                                              حفظ
                                            </button>
                                            <button
                                              type="button"
                                              className="homepickup-btn mini"
                                              onClick={cancelEditPaid}
                                              disabled={paidEditor.saving}
                                            >
                                              إلغاء
                                            </button>
                                          </div>
                                        ) : (
                                          <button
                                            type="button"
                                            className="homepickup-btn mini"
                                            onClick={() => startEditPaid(purchase)}
                                          >
                                            ✏️
                                          </button>
                                        )}
                                      </td>
                                    </>
                                  ) : null}

                                  <td>
                                    {purchase.images?.length ? (
                                      <div className="homepickup-thumbs">
                                        {purchase.images.map((url, index) => (
                                          <img
                                            key={`${purchase.id}-img-${index}`}
                                            src={url}
                                            alt="صورة"
                                            onClick={() => openLightbox(purchase.images, index, purchase.customer_name || "")}
                                          />
                                        ))}
                                      </div>
                                    ) : (
                                      "—"
                                    )}
                                  </td>

                                  <td>
                                    <div className="homepickup-pick-row pickup-checkbox-wrap">
                                      <PickupAnimatedCheckbox
                                        checked={!!purchase.picked_up}
                                        onChange={(event) => togglePicked(purchase.id, event.target.checked)}
                                        ariaLabel={purchase.picked_up ? "تم الاستلام" : "غير مستلم"}
                                      />
                                      <span>{purchase.picked_up ? "تم الاستلام" : "غير مستلم"}</span>
                                    </div>
                                  </td>

                                  <td>{formatDateTime(purchase.picked_up_at)}</td>
                                </tr>
                              );
                            })
                          ) : (
                            <tr>
                              <td colSpan={isRahaf ? 7 : 5} className="homepickup-muted">
                                لا توجد مشتريات
                              </td>
                            </tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  )
                ) : null}
              </>
            )}
          </main>
        </div>
      </div>

      <div className={`pickup-orders-menu-overlay ${ordersMenuOpen ? "open" : ""}`} onClick={() => setOrdersMenuOpen(false)}>
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
              <div className="homepickup-spacer">
                <SessionLoader label="جاري تحميل البيانات..." />
              </div>
            ) : null}
            {!loadingOrders && error ? <div className="homepickup-error homepickup-spacer">{error}</div> : null}

            {!loadingOrders && !error && !groupedOrders.length ? (
              <div className="homepickup-muted homepickup-spacer">
                لا يوجد بيانات
                <div className="homepickup-refresh-row">
                  <button className="homepickup-btn" type="button" onClick={loadOrders}>
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
                          </button>
                        );
                      })}
                    </div>
                  </section>
                ))
              : null}
          </div>
        </aside>
      </div>

      {lightbox.open ? (
        <div className="homepickup-lightbox" onClick={() => setLightbox((prev) => ({ ...prev, open: false }))}>
          <button
            type="button"
            className="homepickup-lightbox-btn close"
            onClick={(event) => {
              event.stopPropagation();
              setLightbox((prev) => ({ ...prev, open: false }));
            }}
          >
            ✕
          </button>
          <div className="homepickup-lightbox-count">
            {lightbox.images.length ? `${lightbox.index + 1}/${lightbox.images.length}` : ""}
            {lightbox.label ? ` — ${lightbox.label}` : ""}
          </div>
          {lightbox.images.length > 1 ? (
            <button
              type="button"
              className="homepickup-lightbox-btn prev"
              onClick={(event) => {
                event.stopPropagation();
                setLightbox((prev) => ({
                  ...prev,
                  index: (prev.index - 1 + prev.images.length) % prev.images.length
                }));
              }}
            >
              ‹
            </button>
          ) : null}
          <img src={lightbox.images[lightbox.index]} alt="صورة كبيرة" />
          {lightbox.images.length > 1 ? (
            <button
              type="button"
              className="homepickup-lightbox-btn next"
              onClick={(event) => {
                event.stopPropagation();
                setLightbox((prev) => ({
                  ...prev,
                  index: (prev.index + 1) % prev.images.length
                }));
              }}
            >
              ›
            </button>
          ) : null}
        </div>
      ) : null}
    </div>
  );
}
