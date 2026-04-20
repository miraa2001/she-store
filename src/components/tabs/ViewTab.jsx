import { useCallback, useEffect, useMemo, useState } from "react";
import { fetchArrivedOrders, formatILS, updateOrderPlacedAtPickup } from "../../lib/orders";
import { searchByName } from "../../lib/search";
import {
  fetchPurchasesByOrder,
  searchPurchasesByCustomerName,
  updatePurchaseBagSize
} from "../../lib/purchases";
import {
  buildArrivalNotifyMessage,
  buildWhatsappUrl,
  buildPickupInquiryMessage,
  resolvePurchaseWhatsappTarget
} from "../../lib/whatsapp";
import SessionLoader from "../common/SessionLoader";

const CAN_EDIT_BAG_ROLES = new Set(["rahaf", "reem", "rawand"]);
const BAG_OPTIONS = ["كيس كبير", "كيس صغير"];

export default function ViewTab({ role, onOpenLightbox, onToast }) {
  const [orders, setOrders] = useState([]);
  const [ordersLoading, setOrdersLoading] = useState(true);
  const [ordersError, setOrdersError] = useState("");
  const [selectedOrderId, setSelectedOrderId] = useState("");

  const [purchases, setPurchases] = useState([]);
  const [purchasesLoading, setPurchasesLoading] = useState(false);
  const [purchasesError, setPurchasesError] = useState("");
  const [viewIndex, setViewIndex] = useState(0);
  const [viewMode, setViewMode] = useState("card");
  const [highlightPurchaseId, setHighlightPurchaseId] = useState("");

  const [searchText, setSearchText] = useState("");
  const [searchResults, setSearchResults] = useState([]);
  const [searching, setSearching] = useState(false);

  const [placedAtPickupSaving, setPlacedAtPickupSaving] = useState(false);
  const [bagSavingId, setBagSavingId] = useState("");

  const selectedOrder = useMemo(
    () => orders.find((order) => String(order.id) === String(selectedOrderId)) || null,
    [orders, selectedOrderId]
  );

  const selectedPurchase = useMemo(
    () => (purchases.length ? purchases[viewIndex] : null),
    [purchases, viewIndex]
  );

  const canEditBag = CAN_EDIT_BAG_ROLES.has(role);
  const canShowWhatsapp = role === "rahaf" && !!selectedOrder?.arrived;

  const purchaseCountLabel = useMemo(() => {
    const totalItems = purchases.reduce((sum, item) => sum + Number(item.qty || 0), 0);
    return `عدد المشتريات: ${purchases.length} — مجموع القطع: ${totalItems}`;
  }, [purchases]);

  const loadOrders = useCallback(async (preferredOrderId = "") => {
    setOrdersLoading(true);
    setOrdersError("");

    try {
      const list = await fetchArrivedOrders();
      setOrders(list);
      setSelectedOrderId((prev) => {
        const candidate = preferredOrderId || prev;
        if (candidate && list.some((order) => String(order.id) === String(candidate))) {
          return candidate;
        }
        return list[0]?.id || "";
      });
    } catch (error) {
      console.error(error);
      setOrdersError("فشل تحميل طلبيات العرض.");
    } finally {
      setOrdersLoading(false);
    }
  }, []);

  const loadPurchases = useCallback(async (orderId) => {
    if (!orderId) {
      setPurchases([]);
      setPurchasesLoading(false);
      return;
    }

    setPurchasesLoading(true);
    setPurchasesError("");

    try {
      const list = await fetchPurchasesByOrder(orderId);
      setPurchases(list || []);
      setViewIndex(0);
    } catch (error) {
      console.error(error);
      setPurchasesError("فشل تحميل مشتريات الطلبية.");
    } finally {
      setPurchasesLoading(false);
    }
  }, []);

  useEffect(() => {
    loadOrders();
  }, [loadOrders]);

  useEffect(() => {
    loadPurchases(selectedOrderId);
  }, [loadPurchases, selectedOrderId]);

  useEffect(() => {
    if (!purchases.length) return;
    setViewIndex((prev) => Math.min(prev, purchases.length - 1));
  }, [purchases]);

  useEffect(() => {
    if (!highlightPurchaseId || !purchases.length) return;

    const index = purchases.findIndex((item) => String(item.id) === String(highlightPurchaseId));
    if (index >= 0) setViewIndex(index);

    const timer = setTimeout(() => setHighlightPurchaseId(""), 2500);
    return () => clearTimeout(timer);
  }, [highlightPurchaseId, purchases]);

  useEffect(() => {
    const query = String(searchText || "").trim();
    if (query.length < 2) {
      setSearchResults([]);
      return;
    }

    let mounted = true;
    const timer = setTimeout(async () => {
      setSearching(true);
      try {
        const list = await searchPurchasesByCustomerName(query, 50);
        if (!mounted) return;

        const arrivedOrderIds = new Set(orders.map((order) => String(order.id)));
        const filteredByOrder = list.filter((item) => arrivedOrderIds.has(String(item.order_id)));
        const filtered = searchByName(filteredByOrder, query, (item) => [item.customer_name]);
        setSearchResults(filtered);
      } catch (error) {
        console.error(error);
        if (mounted) setSearchResults([]);
      } finally {
        if (mounted) setSearching(false);
      }
    }, 250);

    return () => {
      mounted = false;
      clearTimeout(timer);
    };
  }, [orders, searchText]);

  const handleSelectOrder = (order) => {
    setSelectedOrderId(order.id);
    setViewIndex(0);
  };

  const handleResultClick = (result) => {
    const order = orders.find((item) => String(item.id) === String(result.order_id));
    if (!order) return;
    handleSelectOrder(order);
    setHighlightPurchaseId(result.id);
    setSearchResults([]);
  };

  const updateBagSize = async (purchaseId, nextValue) => {
    setBagSavingId(String(purchaseId));
    try {
      await updatePurchaseBagSize(purchaseId, nextValue);
      setPurchases((prev) =>
        prev.map((item) =>
          String(item.id) === String(purchaseId) ? { ...item, bag_size: nextValue } : item
        )
      );
      onToast?.({ type: "success", text: "تم تحديث حجم الكيس." });
    } catch (error) {
      console.error(error);
      onToast?.({ type: "danger", text: "فشل تحديث حجم الكيس." });
    } finally {
      setBagSavingId("");
    }
  };

  const handlePlacedAtPickup = async (checked) => {
    if (!selectedOrder || role !== "rahaf") return;
    setPlacedAtPickupSaving(true);

    try {
      const payload = await updateOrderPlacedAtPickup(selectedOrder.id, checked);
      setOrders((prev) =>
        prev.map((item) =>
          String(item.id) === String(selectedOrder.id)
            ? {
                ...item,
                placedAtPickup: payload.placed_at_pickup,
                placedAtPickupAt: payload.placed_at_pickup_at
              }
            : item
        )
      );
      onToast?.({ type: "success", text: "تم حفظ حالة نقطة الاستلام." });
    } catch (error) {
      console.error(error);
      onToast?.({ type: "danger", text: "فشل حفظ الحالة." });
    } finally {
      setPlacedAtPickupSaving(false);
    }
  };

  const openNotifyWhatsapp = async (purchase) => {
    try {
      const target = await resolvePurchaseWhatsappTarget(purchase);
      const message = buildArrivalNotifyMessage({
        pickupPoint: purchase.pickup_point,
        price: purchase.paid_price ?? purchase.price,
        customerName: target.customerName
      });
      const url = buildWhatsappUrl(target.phone, message);
      if (url.startsWith("whatsapp://")) {
        window.location.href = url;
      } else {
        window.open(url, "_blank", "noopener,noreferrer");
      }
    } catch (error) {
      console.error(error);
      onToast?.({ type: "danger", text: error?.message || "تعذر فتح واتساب." });
    }
  };

  const openInquiryWhatsapp = async (purchase) => {
    try {
      const target = await resolvePurchaseWhatsappTarget(purchase);
      const message = buildPickupInquiryMessage();
      const url = buildWhatsappUrl(target.phone, message);
      if (url.startsWith("whatsapp://")) {
        window.location.href = url;
      } else {
        window.open(url, "_blank", "noopener,noreferrer");
      }
    } catch (error) {
      console.error(error);
      onToast?.({ type: "danger", text: error?.message || "تعذر فتح واتساب." });
    }
  };

  const renderBagControl = (purchase) => {
    if (!canEditBag) return <span>{purchase.bag_size || "كيس صغير"}</span>;

    return (
      <select
        value={purchase.bag_size || "كيس صغير"}
        onChange={(event) => updateBagSize(purchase.id, event.target.value)}
        disabled={bagSavingId === String(purchase.id)}
      >
        {BAG_OPTIONS.map((option) => (
          <option key={option} value={option}>
            {option}
          </option>
        ))}
      </select>
    );
  };

  return (
    <>
      <div className="order-detail-header">
        <div>
          <h2>العرض</h2>
          <p>عرض الطلبات الواصلة مع تنقل سريع بين المشتريات.</p>
        </div>
        <div className="order-detail-actions">
          <input
            className="purchase-search"
            value={searchText}
            onChange={(event) => setSearchText(event.target.value)}
            placeholder="بحث باسم الزبون..."
          />
          <span className="status-chip pending">
            {searchText.trim().length >= 2 ? `${searchResults.length} نتيجة` : `${orders.length} طلبية`}
          </span>
        </div>
      </div>

      {searchText.trim().length >= 2 ? (
        <div className="view-search-results">
          {searching ? (
            <div className="workspace-empty workspace-loader">
              <SessionLoader label="جاري البحث..." />
            </div>
          ) : searchResults.length ? (
            searchResults.map((result) => {
              const order = orders.find((item) => String(item.id) === String(result.order_id));
              return (
                <button
                  key={`${result.id}-${result.order_id}`}
                  type="button"
                  onClick={() => handleResultClick(result)}
                >
                  <b>{result.customer_name || "—"}</b>
                  <span>
                    {order?.name || "طلبية"} — العدد: {result.qty ?? 0} — المدفوع: {formatILS(result.paid_price ?? result.price)} ₪
                  </span>
                </button>
              );
            })
          ) : (
            <div className="workspace-empty">لا توجد نتائج مطابقة.</div>
          )}
        </div>
      ) : null}

      <div className="view-workspace">
        <aside className="view-orders-pane">
          <div className="view-pane-head">
            <b>الطلبيات</b>
            <span className="status-chip completed">{orders.length}</span>
          </div>

          {ordersLoading ? (
            <div className="workspace-empty workspace-loader">
              <SessionLoader label="جاري تحميل الطلبيات..." />
            </div>
          ) : null}
          {ordersError ? <div className="workspace-empty workspace-error">{ordersError}</div> : null}
          {!ordersLoading && !ordersError && !orders.length ? (
            <div className="workspace-empty">لا يوجد طلبيات واصلة.</div>
          ) : null}

          {!ordersLoading && !ordersError && orders.length ? (
            <div className="view-orders-list">
              {orders.map((order) => (
                <button
                  key={order.id}
                  type="button"
                  className={`view-order-item ${String(order.id) === String(selectedOrderId) ? "active" : ""}`}
                  onClick={() => handleSelectOrder(order)}
                >
                  <span>{order.name}</span>
                  <small>عرض</small>
                </button>
              ))}
            </div>
          ) : null}
        </aside>

        <main className="view-main-pane">
          {!selectedOrder ? <div className="workspace-empty">اختاري طلبية من القائمة.</div> : null}

          {selectedOrder ? (
            <>
              <div className="view-main-head">
                <div>
                  <h3>{selectedOrder.name}</h3>
                  <p>{purchaseCountLabel}</p>
                </div>

                <div className="view-head-actions">
                  {role === "rahaf" ? (
                    <label className="arrived-toggle-chip">
                      <input
                        type="checkbox"
                        checked={!!selectedOrder.placedAtPickup}
                        onChange={(event) => handlePlacedAtPickup(event.target.checked)}
                        disabled={placedAtPickupSaving}
                      />
                      <span>تم وضع الطلب في نقطة الاستلام</span>
                    </label>
                  ) : null}

                  <button
                    type="button"
                    className="btn-ghost-light"
                    onClick={() => setViewMode((prev) => (prev === "card" ? "list" : "card"))}
                  >
                    {viewMode === "card" ? "عرض قائمة" : "عرض بطاقة"}
                  </button>
                </div>
              </div>

              {purchasesLoading ? (
                <div className="workspace-empty workspace-loader">
                  <SessionLoader label="جاري تحميل المشتريات..." />
                </div>
              ) : null}
              {purchasesError ? <div className="workspace-empty workspace-error">{purchasesError}</div> : null}

              {!purchasesLoading && !purchasesError && !purchases.length ? (
                <div className="workspace-empty">لا يوجد مشتريات في هذه الطلبية.</div>
              ) : null}

              {!purchasesLoading && !purchasesError && purchases.length ? (
                <>
                  {viewMode === "card" ? (
                    <>
                      <div className="view-card-nav">
                        <button
                          type="button"
                          className="btn-ghost-light"
                          onClick={() => setViewIndex((prev) => (prev - 1 + purchases.length) % purchases.length)}
                        >
                          →
                        </button>
                        <span className="status-chip pending">
                          {viewIndex + 1}/{purchases.length}
                        </span>
                        <button
                          type="button"
                          className="btn-ghost-light"
                          onClick={() => setViewIndex((prev) => (prev + 1) % purchases.length)}
                        >
                          ←
                        </button>
                      </div>

                      {selectedPurchase ? (
                        <article
                          className={`view-purchase-card ${
                            highlightPurchaseId && String(highlightPurchaseId) === String(selectedPurchase.id)
                              ? "highlight"
                              : ""
                          }`}
                        >
                          <div><b>الزبون:</b> {selectedPurchase.customer_name || "—"}</div>
                          <div><b>العدد:</b> {selectedPurchase.qty ?? 0}</div>
                          <div><b>المدفوع:</b> {formatILS(selectedPurchase.paid_price ?? selectedPurchase.price)} ₪</div>
                          <div>
                            <b>حجم الكيس:</b> {renderBagControl(selectedPurchase)}
                          </div>
                          <div>
                            <b>روابط:</b>{" "}
                            {selectedPurchase.links?.length
                              ? selectedPurchase.links.map((link, index) => (
                                  <a key={`${selectedPurchase.id}-link-${index}`} href={link} target="_blank" rel="noreferrer">
                                    رابط {index + 1}
                                  </a>
                                ))
                              : "—"}
                          </div>
                          <div><b>مكان الاستلام:</b> {selectedPurchase.pickup_point || "—"}</div>
                          <div className="view-card-images">
                            {selectedPurchase.images?.length
                              ? selectedPurchase.images.map((img, index) => (
                                  <button
                                    key={img.id || `${selectedPurchase.id}-${index}`}
                                    type="button"
                                    className="purchase-image-thumb"
                                    onClick={() =>
                                      onOpenLightbox?.(
                                        selectedPurchase.images,
                                        index,
                                        selectedPurchase.customer_name || "صورة"
                                      )
                                    }
                                  >
                                    <img src={img.url} alt="صورة" />
                                  </button>
                                ))
                              : <span className="muted">لا توجد صور</span>}
                          </div>
                          <div><b>ملاحظة:</b> {selectedPurchase.note?.trim() || "—"}</div>
                          {canShowWhatsapp ? (
                            <div className="wa-actions-row">
                              <button
                                type="button"
                                className="wa-btn wa-btn-inquiry"
                                onClick={() => openInquiryWhatsapp(selectedPurchase)}
                              >
                                استعلام عن نقطة الاستلام❓
                              </button>
                              <button
                                type="button"
                                className="wa-btn wa-btn-notify"
                                onClick={() => openNotifyWhatsapp(selectedPurchase)}
                              >
                                اعلام بوصول الطلب🔔
                              </button>
                            </div>
                          ) : null}
                        </article>
                      ) : null}
                    </>
                  ) : (
                    <div className="view-list-wrap">
                      <table className="view-list-table">
                        <thead>
                          <tr>
                            <th>#</th>
                            <th>الزبون</th>
                            <th>العدد</th>
                            <th>المدفوع</th>
                            <th>مكان الاستلام</th>
                            <th>حجم الكيس</th>
                            <th>روابط</th>
                            <th>صور</th>
                            <th>ملاحظة</th>
                            <th>استعلام نقطة الاستلام</th>
                            <th>واتساب</th>
                          </tr>
                        </thead>
                        <tbody>
                          {purchases.map((purchase, index) => (
                            <tr
                              key={purchase.id}
                              className={
                                highlightPurchaseId && String(highlightPurchaseId) === String(purchase.id)
                                  ? "highlight"
                                  : ""
                              }
                            >
                              <td>{index + 1}</td>
                              <td>{purchase.customer_name || "—"}</td>
                              <td>{purchase.qty ?? 0}</td>
                              <td>{formatILS(purchase.paid_price ?? purchase.price)} ₪</td>
                              <td>{purchase.pickup_point || "—"}</td>
                              <td>{renderBagControl(purchase)}</td>
                              <td>
                                {purchase.links?.length
                                  ? purchase.links.map((link, i) => (
                                      <a key={`${purchase.id}-list-link-${i}`} href={link} target="_blank" rel="noreferrer">
                                        رابط {i + 1}
                                      </a>
                                    ))
                                  : "—"}
                              </td>
                              <td>
                                <div className="view-list-images">
                                  {purchase.images?.length
                                    ? purchase.images.map((img, imgIndex) => (
                                        <button
                                          key={img.id || `${purchase.id}-list-img-${imgIndex}`}
                                          type="button"
                                          className="purchase-image-thumb"
                                          onClick={() =>
                                            onOpenLightbox?.(
                                              purchase.images,
                                              imgIndex,
                                              purchase.customer_name || "صورة"
                                            )
                                          }
                                        >
                                          <img src={img.url} alt="صورة" />
                                        </button>
                                      ))
                                    : "—"}
                                </div>
                              </td>
                              <td>{purchase.note?.trim() || "—"}</td>
                              <td>
                                {canShowWhatsapp ? (
                                  <button
                                    type="button"
                                    className="wa-btn wa-btn-inquiry"
                                    onClick={() => openInquiryWhatsapp(purchase)}
                                  >
                                    استعلام
                                  </button>
                                ) : (
                                  "—"
                                )}
                              </td>
                              <td>
                                {canShowWhatsapp ? (
                                  <button
                                    type="button"
                                    className="wa-btn wa-btn-notify"
                                    onClick={() => openNotifyWhatsapp(purchase)}
                                  >
                                    واتساب
                                  </button>
                                ) : (
                                  "—"
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </>
              ) : null}
            </>
          ) : null}
        </main>
      </div>
    </>
  );
}
