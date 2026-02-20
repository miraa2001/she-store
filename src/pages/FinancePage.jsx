import { useCallback, useEffect, useMemo, useState } from "react";
import { useLocation } from "react-router-dom";
import { useAuthProfile } from "../hooks/useAuthProfile";
import { formatDMY } from "../lib/dateFormat";
import { getOrdersNavItems, isNavHrefActive } from "../lib/navigation";
import { setBodyScrollLock } from "../lib/bodyScrollLock";
import { formatILS, parsePrice } from "../lib/orders";
import { signOutAndRedirect } from "../lib/session";
import { sb } from "../lib/supabaseClient";
import SessionLoader from "../components/common/SessionLoader";
import AppNavIcon from "../components/common/AppNavIcon";
import SheStoreLogo from "../components/common/SheStoreLogo";
import "./pickup-common.css";
import "./finance-page.css";

function getOrderDate(order) {
  if (order?.order_date) {
    const d = new Date(`${order.order_date}T00:00:00`);
    if (!Number.isNaN(d.getTime())) return d;
  }
  const fallback = new Date(order?.created_at || Date.now());
  return Number.isNaN(fallback.getTime()) ? new Date() : fallback;
}

function getOrderDateKey(order) {
  return formatDMY(order?.createdAt || order?.created_at);
}

function buildOrderGroups(orderList) {
  const groups = [];
  const map = new Map();

  orderList.forEach((order) => {
    const dateKey = getOrderDateKey(order) || "بدون تاريخ";
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

function monthKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function monthLabel(date) {
  return date.toLocaleDateString("ar", { month: "long", year: "numeric" });
}

function monthNames() {
  return ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"];
}

function topPickupPoint(map) {
  let top = "—";
  let max = 0;
  map.forEach((value, key) => {
    if (value > max) {
      max = value;
      top = key;
    }
  });
  return top;
}

export default function FinancePage({ embedded = false }) {
  const { profile } = useAuthProfile();
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [ordersMenuOpen, setOrdersMenuOpen] = useState(false);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [orders, setOrders] = useState([]);
  const [orderStatsMap, setOrderStatsMap] = useState(new Map());

  const [activeTab, setActiveTab] = useState("orders");
  const [selectedOrderId, setSelectedOrderId] = useState("");
  const [spentInput, setSpentInput] = useState("");
  const [spentMessage, setSpentMessage] = useState("");
  const [savingSpent, setSavingSpent] = useState(false);

  const [selectedYear, setSelectedYear] = useState(null);
  const [selectedMonthKey, setSelectedMonthKey] = useState("");
  const location = useLocation();
  const sidebarLinks = useMemo(() => getOrdersNavItems(profile.role), [profile.role]);

  const loadData = useCallback(async () => {
    setLoading(true);
    setError("");

    try {
      const [ordersRes, purchasesRes] = await Promise.all([
        sb
          .from("orders")
          .select("id, order_name, order_date, created_at, spent_amount")
          .order("created_at", { ascending: false }),
        sb
          .from("purchases")
          .select("order_id, price, paid_price, pickup_point, collected, picked_up")
      ]);

      if (ordersRes.error) throw ordersRes.error;
      if (purchasesRes.error) throw purchasesRes.error;

      const orderRows = (ordersRes.data || []).map((order) => ({
        ...order,
        spent_amount: parsePrice(order.spent_amount)
      }));

      const stats = new Map();

      (purchasesRes.data || []).forEach((purchase) => {
        const orderId = purchase.order_id;
        if (!orderId) return;

        if (!stats.has(orderId)) {
          stats.set(orderId, {
            collected: 0,
            expected: 0,
            purchaseCount: 0,
            pickupTotals: new Map(),
            pickupCollectedTotals: new Map(),
            pickupCounts: new Map()
          });
        }

        const stat = stats.get(orderId);
        const value = parsePrice(purchase.paid_price ?? purchase.price);
        const pickup = String(purchase.pickup_point || "").trim() || "بدون نقطة";

        stat.expected += value;
        stat.purchaseCount += 1;

        stat.pickupTotals.set(pickup, (stat.pickupTotals.get(pickup) || 0) + value);
        stat.pickupCounts.set(pickup, (stat.pickupCounts.get(pickup) || 0) + 1);

        if (purchase.collected) {
          stat.collected += value;
          stat.pickupCollectedTotals.set(pickup, (stat.pickupCollectedTotals.get(pickup) || 0) + value);
        }
      });

      setOrders(orderRows);
      setOrderStatsMap(stats);
    } catch (err) {
      console.error(err);
      setOrders([]);
      setOrderStatsMap(new Map());
      setError("تعذر تحميل بيانات المالية.");
    } finally {
      setLoading(false);
    }
  }, []);

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
    setBodyScrollLock("finance-page", shouldLockBody);
    return () => {
      setBodyScrollLock("finance-page", false);
    };
  }, [ordersMenuOpen, sidebarOpen]);

  useEffect(() => {
    setSidebarOpen(false);
    setOrdersMenuOpen(false);
  }, [location.pathname, location.search, location.hash]);

  useEffect(() => {
    if (profile.loading || !profile.authenticated || profile.role !== "rahaf") return;
    loadData();
  }, [loadData, profile.authenticated, profile.loading, profile.role]);

  const orderRows = useMemo(() => {
    return orders.map((order) => {
      const stats = orderStatsMap.get(order.id) || {
        collected: 0,
        expected: 0,
        purchaseCount: 0,
        pickupTotals: new Map(),
        pickupCollectedTotals: new Map(),
        pickupCounts: new Map()
      };
      const spent = parsePrice(order.spent_amount);
      const pending = Math.max(0, stats.expected - stats.collected);
      return {
        id: order.id,
        name: order.order_name || "طلبية",
        createdAt: order.created_at,
        orderDate: order.order_date,
        spent,
        collected: stats.collected,
        expected: stats.expected,
        pending,
        purchaseCount: stats.purchaseCount,
        pickupTotals: stats.pickupTotals,
        pickupCollectedTotals: stats.pickupCollectedTotals,
        pickupCounts: stats.pickupCounts
      };
    });
  }, [orderStatsMap, orders]);

  const groupedOrders = useMemo(() => buildOrderGroups(orderRows), [orderRows]);

  const selectedOrder = useMemo(
    () => orderRows.find((order) => String(order.id) === String(selectedOrderId)) || null,
    [orderRows, selectedOrderId]
  );

  useEffect(() => {
    if (!orderRows.length) {
      setSelectedOrderId("");
      return;
    }
    setSelectedOrderId((prev) =>
      prev && orderRows.some((order) => String(order.id) === String(prev)) ? prev : orderRows[0].id
    );
  }, [orderRows]);

  useEffect(() => {
    if (!selectedOrder) {
      setSpentInput("");
      return;
    }
    setSpentInput(selectedOrder.spent ? String(selectedOrder.spent) : "");
    setSpentMessage("");
  }, [selectedOrder]);

  const monthSummary = useMemo(() => {
    const map = new Map();

    orderRows.forEach((order) => {
      const date = getOrderDate(order);
      const key = monthKey(date);
      if (!map.has(key)) {
        map.set(key, {
          key,
          label: monthLabel(date),
          year: date.getFullYear(),
          monthIndex: date.getMonth(),
          orders: 0,
          purchases: 0,
          collected: 0,
          expected: 0,
          spent: 0,
          pickupTotals: new Map()
        });
      }

      const bucket = map.get(key);
      bucket.orders += 1;
      bucket.purchases += order.purchaseCount;
      bucket.collected += order.collected;
      bucket.expected += order.expected;
      bucket.spent += order.spent;

      order.pickupTotals.forEach((value, pickup) => {
        bucket.pickupTotals.set(pickup, (bucket.pickupTotals.get(pickup) || 0) + value);
      });
    });

    const years = Array.from(new Set(Array.from(map.values()).map((month) => month.year))).sort((a, b) => a - b);
    return { map, years };
  }, [orderRows]);

  useEffect(() => {
    if (!monthSummary.years.length) {
      setSelectedYear(null);
      setSelectedMonthKey("");
      return;
    }

    setSelectedYear((prev) =>
      prev !== null && monthSummary.years.includes(prev)
        ? prev
        : monthSummary.years[monthSummary.years.length - 1]
    );
  }, [monthSummary.years]);

  useEffect(() => {
    if (selectedYear === null) return;

    const monthKeys = Array.from(monthSummary.map.values())
      .filter((month) => month.year === selectedYear)
      .map((month) => month.key)
      .sort();

    if (!monthKeys.length) {
      setSelectedMonthKey("");
      return;
    }

    setSelectedMonthKey((prev) => (prev && monthKeys.includes(prev) ? prev : monthKeys[monthKeys.length - 1]));
  }, [monthSummary.map, selectedYear]);

  const selectedMonth = useMemo(
    () => (selectedMonthKey ? monthSummary.map.get(selectedMonthKey) || null : null),
    [monthSummary.map, selectedMonthKey]
  );

  async function saveSpent() {
    if (!selectedOrder || savingSpent) return;

    const value = spentInput.trim() === "" ? 0 : Number(spentInput);
    if (!Number.isFinite(value) || value < 0) {
      setSpentMessage("المصروف غير صحيح.");
      return;
    }

    setSavingSpent(true);
    setSpentMessage("جاري الحفظ...");

    const { error: updateError } = await sb
      .from("orders")
      .update({ spent_amount: value })
      .eq("id", selectedOrder.id);

    if (updateError) {
      console.error(updateError);
      setSpentMessage("فشل الحفظ.");
      setSavingSpent(false);
      return;
    }

    setOrders((prev) =>
      prev.map((order) =>
        String(order.id) === String(selectedOrder.id) ? { ...order, spent_amount: value } : order
      )
    );

    setSavingSpent(false);
    setSpentMessage("تم ✅");
    window.setTimeout(() => setSpentMessage(""), 1500);
  }

  async function signOut() {
    await signOutAndRedirect();
  }

  if (profile.loading) {
    return (
      <div className="finance-page finance-state" dir="rtl">
        <SessionLoader />
      </div>
    );
  }

  if (!profile.authenticated) {
    return (
      <div className="finance-page finance-state" dir="rtl">
        <div className="finance-note finance-note-danger">
          <h2>لا توجد جلسة نشطة</h2>
          <p>يلزم تسجيل الدخول أولًا.</p>
          <a href="#/login" className="finance-link">فتح تسجيل الدخول</a>
        </div>
      </div>
    );
  }

  if (profile.role !== "rahaf") {
    return (
      <div className="finance-page finance-state" dir="rtl">
        <div className="finance-note finance-note-danger">
          <h2>لا توجد صلاحية</h2>
          <p>هذه الصفحة متاحة لحساب رهف فقط.</p>
          <a href="#/orders" className="finance-link">العودة للطلبيات</a>
        </div>
      </div>
    );
  }

  return (
    <div className={`finance-page ${embedded ? "embedded" : ""}`} dir="rtl">
      {!embedded ? (
        <>
          <div
            className={`finance-overlay app-sidebar-overlay ${sidebarOpen ? "open" : ""}`}
            onClick={() => setSidebarOpen(false)}
          />

          <aside className={`finance-sidebar app-sidebar-drawer ${sidebarOpen ? "open" : ""}`}>
            <div className="finance-sidebar-head app-sidebar-head">
              <div className="app-sidebar-brand">
                <SheStoreLogo className="app-sidebar-logo-link" imageClassName="app-sidebar-logo-img" />
                <b>شي ستور</b>
              </div>
              <button
                type="button"
                className="finance-menu-btn danger app-sidebar-close"
                onClick={() => setSidebarOpen(false)}
              >
                ✕
              </button>
            </div>

            <div className="finance-sidebar-content app-sidebar-content">
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
                تسجيل الخروج
              </button>
            </div>
          </aside>
        </>
      ) : null}

      <div className="finance-wrap">
        {!embedded ? (
          <div className="finance-topbar">
            <div className="topbar-brand-with-logo">
              <SheStoreLogo className="topbar-logo-link" imageClassName="topbar-logo-img" />
              <div className="finance-brand">
                <b>المالية</b>
                <div className="finance-muted">متابعة التحصيل والمصاريف</div>
              </div>
            </div>

            <button type="button" className="finance-menu-btn" onClick={() => setSidebarOpen(true)}>
              ☰
            </button>
          </div>
        ) : null}

        <div className="finance-tabs">
          <button
            type="button"
            className={`finance-tab-btn ${activeTab === "orders" ? "active" : ""}`}
            onClick={() => setActiveTab("orders")}
          >
            حسب الطلبية
          </button>
          <button
            type="button"
            className={`finance-tab-btn ${activeTab === "months" ? "active" : ""}`}
            onClick={() => setActiveTab("months")}
          >
            حسب الشهر
          </button>
        </div>

        {error ? <div className="finance-error">{error}</div> : null}

        {loading ? (
          <div className="finance-loading">
            <SessionLoader label="جاري تحميل البيانات..." />
          </div>
        ) : null}

        {!loading && activeTab === "orders" ? (
          <>
            <div className="finance-orders-menu-row">
              <button
                type="button"
                className="pickup-orders-menu-trigger"
                onClick={() => setOrdersMenuOpen(true)}
                aria-label="فتح قائمة الطلبيات"
              >
                <AppNavIcon name="package" className="icon" />
                <span>الطلبيات</span>
                <b>{orderRows.length}</b>
              </button>

              <span className="finance-pill">
                {selectedOrder ? `الطلبية المختارة: ${selectedOrder.name}` : "اختر طلبية"}
              </span>
            </div>

            <div
              className={`pickup-orders-menu-overlay ${ordersMenuOpen ? "open" : ""}`}
              onClick={() => setOrdersMenuOpen(false)}
            >
              <aside className="pickup-orders-menu-panel" onClick={(event) => event.stopPropagation()}>
                <div className="pickup-orders-menu-head">
                  <div className="pickup-orders-menu-title">
                    <AppNavIcon name="package" className="icon" />
                    <strong>الطلبيات</strong>
                    <b>{orderRows.length}</b>
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
                  {!groupedOrders.length ? (
                    <div className="finance-empty">
                      لا توجد طلبيات
                      <div className="finance-refresh-row">
                        <button type="button" className="finance-btn" onClick={loadData}>إعادة تحميل</button>
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
                            const active = String(order.id) === String(selectedOrderId);
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
                                  <strong>{order.name || "طلبية"}</strong>
                                  <span>{getOrderDateKey(order) || "—"}</span>
                                </div>
                                <div className="order-meta">
                                  <small className="status at_pickup">{formatILS(order.collected)} ₪</small>
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

            <div className="finance-grid finance-grid--single">
              <main className="finance-card">
                {!selectedOrder ? (
                  <div className="finance-empty">اختر طلبية من القائمة</div>
                ) : (
                  <>
                    <div className="finance-row">
                      <b>{selectedOrder.name}</b>
                      <span
                        className={`finance-status ${
                          selectedOrder.pending === 0 && selectedOrder.expected > 0
                            ? "success"
                            : selectedOrder.expected === 0
                              ? "neutral"
                              : "warning"
                        }`}
                      >
                        {selectedOrder.expected === 0
                          ? "لا يوجد تحصيل"
                          : selectedOrder.pending === 0
                            ? "مكتمل"
                            : "قيد التحصيل"}
                      </span>
                    </div>

                    <div className="finance-kpi-grid">
                      <div className="finance-kpi">
                        <div className="label">تم تحصيله</div>
                        <div className="value">{formatILS(selectedOrder.collected)}</div>
                      </div>
                      <div className="finance-kpi">
                        <div className="label">متبقي للتحصيل</div>
                        <div className="value">{formatILS(selectedOrder.pending)}</div>
                      </div>
                      <div className="finance-kpi">
                        <div className="label">إجمالي المطلوب</div>
                        <div className="value">{formatILS(selectedOrder.expected)}</div>
                      </div>
                    </div>

                    <div className="finance-kpi-grid compact">
                      <div className="finance-kpi">
                        <div className="label">صافي المحصل</div>
                        <div className={`value ${selectedOrder.collected - selectedOrder.spent < 0 ? "neg" : ""}`}>
                          {formatILS(selectedOrder.collected - selectedOrder.spent)}
                        </div>
                      </div>
                      <div className="finance-kpi">
                        <div className="label">صافي المتوقع</div>
                        <div className={`value ${selectedOrder.expected - selectedOrder.spent < 0 ? "neg" : ""}`}>
                          {formatILS(selectedOrder.expected - selectedOrder.spent)}
                        </div>
                      </div>
                      <div className="finance-kpi">
                        <div className="label">المصروف</div>
                        <div className="value">{formatILS(selectedOrder.spent)}</div>
                      </div>
                    </div>

                    <div className="finance-spent-row">
                      <label htmlFor="financeSpentInput">المصروف:</label>
                      <input
                        id="financeSpentInput"
                        type="number"
                        step="0.01"
                        min="0"
                        value={spentInput}
                        onChange={(event) => setSpentInput(event.target.value)}
                        placeholder="اكتب قيمة المصروف"
                      />
                      <button type="button" className="finance-btn primary" onClick={saveSpent} disabled={savingSpent}>
                        {savingSpent ? "جاري الحفظ..." : "حفظ المصروف"}
                      </button>
                      {spentMessage ? <span className="finance-muted">{spentMessage}</span> : null}
                    </div>
                  </>
                )}
              </main>
            </div>
          </>
        ) : null}

        {!loading && activeTab === "months" ? (
          <main className="finance-card">
            <div className="finance-row center">
              <div className="finance-section-title">ملخص شهري</div>
            </div>

            <div className="finance-month-picker">
              <div className="finance-month-header">
                <button
                  type="button"
                  className="finance-btn mini"
                  disabled={!monthSummary.years.length || selectedYear === monthSummary.years[0]}
                  onClick={() =>
                    setSelectedYear((prev) =>
                      prev === null
                        ? prev
                        : monthSummary.years[Math.max(0, monthSummary.years.indexOf(prev) - 1)]
                    )
                  }
                >
                  ‹
                </button>
                <b>{selectedYear ?? "—"}</b>
                <button
                  type="button"
                  className="finance-btn mini"
                  disabled={!monthSummary.years.length || selectedYear === monthSummary.years[monthSummary.years.length - 1]}
                  onClick={() =>
                    setSelectedYear((prev) =>
                      prev === null
                        ? prev
                        : monthSummary.years[Math.min(monthSummary.years.length - 1, monthSummary.years.indexOf(prev) + 1)]
                    )
                  }
                >
                  ›
                </button>
              </div>

              <div className="finance-month-grid">
                {monthNames().map((name, index) => {
                  const key = `${selectedYear}-${String(index + 1).padStart(2, "0")}`;
                  const hasData = monthSummary.map.has(key);
                  return (
                    <button
                      key={name}
                      type="button"
                      className={`finance-month-btn ${selectedMonthKey === key ? "active" : ""}`}
                      disabled={!hasData}
                      onClick={() => setSelectedMonthKey(key)}
                    >
                      {name}
                    </button>
                  );
                })}
              </div>
            </div>

            {!selectedMonth ? (
              <div className="finance-empty">لا توجد بيانات للشهر</div>
            ) : (
              <div className="finance-month-shell">
                <div className="finance-row">
                  <b>{selectedMonth.label}</b>
                  <span
                    className={`finance-status ${
                      selectedMonth.expected > 0 && selectedMonth.collected >= selectedMonth.expected
                        ? "success"
                        : "warning"
                    }`}
                  >
                    {selectedMonth.expected > 0 && selectedMonth.collected >= selectedMonth.expected
                      ? "مكتمل"
                      : "قيد التحصيل"}
                  </span>
                </div>

                <div className="finance-health-row">
                  <span className="finance-pill">الطلبيات: {selectedMonth.orders}</span>
                  <span className="finance-pill">العمليات: {selectedMonth.purchases}</span>
                  <span className="finance-pill">
                    نسبة التحصيل: {selectedMonth.expected > 0 ? Math.round((selectedMonth.collected / selectedMonth.expected) * 100) : 0}%
                  </span>
                  <span className="finance-pill">أعلى نقطة: {topPickupPoint(selectedMonth.pickupTotals)}</span>
                </div>

                <div className="finance-kpi-grid">
                  <div className="finance-kpi">
                    <div className="label">تم تحصيله</div>
                    <div className="value">{formatILS(selectedMonth.collected)}</div>
                  </div>
                  <div className="finance-kpi">
                    <div className="label">متبقي للتحصيل</div>
                    <div className="value">{formatILS(Math.max(0, selectedMonth.expected - selectedMonth.collected))}</div>
                  </div>
                  <div className="finance-kpi">
                    <div className="label">إجمالي المطلوب</div>
                    <div className="value">{formatILS(selectedMonth.expected)}</div>
                  </div>
                </div>

                <div className="finance-kpi-grid compact">
                  <div className="finance-kpi">
                    <div className="label">صافي المحصل</div>
                    <div className={`value ${selectedMonth.collected - selectedMonth.spent < 0 ? "neg" : ""}`}>
                      {formatILS(selectedMonth.collected - selectedMonth.spent)}
                    </div>
                  </div>
                  <div className="finance-kpi">
                    <div className="label">صافي المتوقع</div>
                    <div className={`value ${selectedMonth.expected - selectedMonth.spent < 0 ? "neg" : ""}`}>
                      {formatILS(selectedMonth.expected - selectedMonth.spent)}
                    </div>
                  </div>
                  <div className="finance-kpi">
                    <div className="label">المصروف</div>
                    <div className="value">{formatILS(selectedMonth.spent)}</div>
                  </div>
                </div>
              </div>
            )}
          </main>
        ) : null}
      </div>
    </div>
  );
}
