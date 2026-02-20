import { useEffect, useMemo, useRef, useState } from "react";
import { formatILS } from "../../lib/orders";
import SessionLoader from "../common/SessionLoader";
import OrderStatusDropdown from "./OrderStatusDropdown";
import pdfExportIconWeb from "../../assets/icons/documents/pdf-web.png";
import pdfExportIconAndroid from "../../assets/icons/documents/pdf-android.png";
import pdfExportIconIos from "../../assets/icons/documents/pdf-ios.png";
import actionsMenuIcon from "../../assets/icons/actions/menu-vertical.png";
import editIcon from "../../assets/icons/actions/edit.png";
import editPriceIcon from "../../assets/icons/actions/edit-price.png";
import deleteIcon from "../../assets/icons/actions/delete.png";

function normalizeSlideIndex(index, total) {
  if (!total) return 0;
  return ((Number(index) || 0) % total + total) % total;
}

export default function OrdersTab({
  selectedOrder,
  selectedOrderStatus = "pending",
  orderStatusLocked = false,
  orderStatusSaving = false,
  purchaseStats,
  isMobile = false,
  isRahaf,
  editMode,
  onUpdateOrderStatus,
  onOpenAddModal,
  onExportPdf,
  canExportPdf = true,
  pdfExporting,
  customersError,
  purchasesLoading,
  purchasesError,
  filteredPurchases,
  menuPurchaseId,
  onTogglePurchaseMenu,
  onEditPurchase,
  onMarkPaid,
  onDeletePurchase,
  onOpenLightbox,
  onInquireWhatsapp,
  onNotifyWhatsapp,
  highlightPurchaseId = "",
  hidePurchaseGrid = false
}) {
  const [cardSlideIndexes, setCardSlideIndexes] = useState({});
  const highlightRef = useRef(null);
  const canEditOrderStatus = isRahaf && editMode && !!selectedOrder;
  const pdfExportIcon = useMemo(() => {
    if (typeof navigator === "undefined") return pdfExportIconWeb;
    const ua = String(navigator.userAgent || "").toLowerCase();
    if (ua.includes("android")) return pdfExportIconAndroid;
    if (ua.includes("iphone") || ua.includes("ipad") || ua.includes("ipod")) return pdfExportIconIos;
    return pdfExportIconWeb;
  }, []);

  useEffect(() => {
    if (!highlightPurchaseId) return;
    if (!highlightRef.current) return;
    highlightRef.current.scrollIntoView({ behavior: "smooth", block: "center" });
  }, [highlightPurchaseId, filteredPurchases.length]);

  const setSlideIndex = (purchaseId, total, nextIndex) => {
    const normalized = normalizeSlideIndex(nextIndex, total);
    setCardSlideIndexes((prev) => {
      if (prev[purchaseId] === normalized) return prev;
      return { ...prev, [purchaseId]: normalized };
    });
  };

  return (
    <>
      <div className="order-detail-header">
        <div>
          <h2>{selectedOrder?.name || "اختاري طلبًا"}</h2>
          <p>
            عدد المشتريات: {purchaseStats.count} — مجموع القطع: {purchaseStats.totalQty} — مجموع الأسعار:{" "}
            {formatILS(purchaseStats.totalPrice)} ₪
          </p>
        </div>

        <div className="order-detail-actions">

          {canEditOrderStatus && !isMobile ? (
            <div className="order-status-inline">
              <span className="order-status-inline-label">حالة الطلب</span>
              <OrderStatusDropdown
                value={selectedOrderStatus}
                onChange={(nextValue) => onUpdateOrderStatus?.(nextValue)}
                disabled={orderStatusSaving || orderStatusLocked}
                lockCollected={orderStatusLocked}
              />
            </div>
          ) : null}

          {!isMobile && isRahaf && editMode ? (
            <button className="btn-primary" type="button" onClick={onOpenAddModal}>
              + إضافة مشترى
            </button>
          ) : null}

          {canExportPdf && !isMobile ? (
            <button
              className="btn-ghost-light"
              type="button"
              onClick={onExportPdf}
              disabled={pdfExporting}
              aria-label={pdfExporting ? "جاري تصدير PDF" : "تصدير PDF"}
            >
              <img
                src={pdfExportIcon}
                alt="pdf--v3"
                style={{ width: 28, height: 28, objectFit: "contain", display: "block" }}
              />
            </button>
          ) : null}
        </div>
      </div>

      {isMobile && canEditOrderStatus ? (
        <div className="mobile-order-controls">
          <div className="order-status-inline">
            <span className="order-status-inline-label">حالة الطلب</span>
            <OrderStatusDropdown
              value={selectedOrderStatus}
              onChange={(nextValue) => onUpdateOrderStatus?.(nextValue)}
              disabled={orderStatusSaving || orderStatusLocked}
              lockCollected={orderStatusLocked}
            />
          </div>

          <button className="btn-primary mobile-add-purchase-btn" type="button" onClick={onOpenAddModal}>
            + إضافة مشترى
          </button>
        </div>
      ) : null}

      {customersError ? <div className="workspace-empty workspace-error">{customersError}</div> : null}
      {purchasesLoading ? (
        <div className="workspace-empty workspace-loader">
          <SessionLoader />
        </div>
      ) : null}
      {purchasesError ? <div className="workspace-empty workspace-error">{purchasesError}</div> : null}

      {!purchasesLoading && !purchasesError && !filteredPurchases.length ? (
        <div className="workspace-empty">لا توجد مشتريات مطابقة.</div>
      ) : null}

      {!purchasesLoading && !purchasesError && filteredPurchases.length && !hidePurchaseGrid ? (
        <div className="purchase-cards-grid">
          {filteredPurchases.map((purchase) => {
            const canShowWhatsapp = !!selectedOrder?.arrived;
            const imageList = Array.isArray(purchase.images)
              ? purchase.images.filter((img) => img?.url)
              : [];
            const totalImages = imageList.length;
            const currentSlide = normalizeSlideIndex(cardSlideIndexes[purchase.id] || 0, totalImages);
            const menuNode = isRahaf && editMode ? (
              <div className="purchase-menu-wrap" data-menu-root>
                <button
                  type="button"
                  className="icon-btn menu-dots menu-dots-trigger"
                  onClick={(event) => {
                    event.stopPropagation();
                    onTogglePurchaseMenu(purchase.id);
                  }}
                  aria-label="إجراءات"
                  aria-haspopup="menu"
                  aria-expanded={String(menuPurchaseId) === String(purchase.id)}
                >
                  <img src={actionsMenuIcon} alt="" aria-hidden="true" className="menu-dots-icon" />
                </button>

                {String(menuPurchaseId) === String(purchase.id) ? (
                  <div className="purchase-menu-pop" role="menu">
                    <button type="button" className="value" role="menuitem" onClick={() => onEditPurchase(purchase)}>
                      <img src={editIcon} alt="" aria-hidden="true" className="value-icon" />
                      <span>تعديل</span>
                    </button>
                    <button type="button" className="value" role="menuitem" onClick={() => onMarkPaid(purchase)}>
                      <img src={editPriceIcon} alt="" aria-hidden="true" className="value-icon" />
                      <span>تعديل المدفوع</span>
                    </button>
                    <button
                      type="button"
                      className="value danger"
                      role="menuitem"
                      onClick={() => onDeletePurchase(purchase)}
                    >
                      <img src={deleteIcon} alt="" aria-hidden="true" className="value-icon" />
                      <span>حذف</span>
                    </button>
                  </div>
                ) : null}
              </div>
            ) : null;

            return (
              <article
                key={purchase.id}
                className={`purchase-card ${String(highlightPurchaseId) === String(purchase.id) ? "purchase-highlight" : ""}`}
                ref={String(highlightPurchaseId) === String(purchase.id) ? highlightRef : null}
              >
                <div className="purchase-desktop-shell">
                  <article className="purchaseVCard">
                    <div className="purchaseVMedia" dir="ltr">
                      {totalImages ? (
                        <>
                          <div className="purchaseVTrack" style={{ transform: `translateX(-${currentSlide * 100}%)` }}>
                            {imageList.map((img, index) => (
                              <button
                                key={img.id || `${purchase.id}-slide-${index}`}
                                type="button"
                                className="purchaseVSlide"
                                onClick={() =>
                                  onOpenLightbox(imageList, index, purchase.customer_name || "صورة المشترى")
                                }
                                aria-label={`عرض الصورة ${index + 1}`}
                              >
                                <img src={img.url} alt={`صورة ${index + 1}`} loading="lazy" />
                              </button>
                            ))}
                          </div>

                          {totalImages > 1 ? (
                            <>
                              <button
                                className="purchaseVNav prev"
                                type="button"
                                aria-label="الصورة السابقة"
                                onClick={() => setSlideIndex(purchase.id, totalImages, currentSlide - 1)}
                              >
                                ‹
                              </button>
                              <button
                                className="purchaseVNav next"
                                type="button"
                                aria-label="الصورة التالية"
                                onClick={() => setSlideIndex(purchase.id, totalImages, currentSlide + 1)}
                              >
                                ›
                              </button>
                              <div className="purchaseVDots">
                                {imageList.map((img, index) => (
                                  <button
                                    key={img.id || `${purchase.id}-dot-${index}`}
                                    type="button"
                                    className={`purchaseVDot ${index === currentSlide ? "is-active" : ""}`}
                                    aria-label={`الانتقال للصورة ${index + 1}`}
                                    onClick={() => setSlideIndex(purchase.id, totalImages, index)}
                                  />
                                ))}
                              </div>
                            </>
                          ) : null}

                          <span className="purchaseVCount">
                            {currentSlide + 1}/{totalImages}
                          </span>
                        </>
                      ) : (
                        <div className="purchaseVPlaceholder">لا توجد صور</div>
                      )}

                    </div>

                    <div className="purchaseVBody" dir="rtl">
                      <div className="purchaseVField purchaseVField-primary">
                        <p className="purchaseVLabel">الاسم</p>
                        <p className="purchaseVValue purchaseVValue-primary">{purchase.customer_name || "—"}</p>
                      </div>
                      <div className="purchaseVField">
                        <p className="purchaseVLabel">عدد القطع</p>
                        <p className="purchaseVValue">{purchase.qty || 0}</p>
                      </div>
                      <div className="purchaseVField">
                        <p className="purchaseVLabel">السعر</p>
                        <p className="purchaseVValue">{formatILS(purchase.price)} ₪</p>
                      </div>
                      <div className="purchaseVField">
                        <p className="purchaseVLabel">مكان الاستلام</p>
                        <p className="purchaseVValue">{purchase.pickup_point || "—"}</p>
                      </div>
                    </div>

                    {purchase.links?.length ? (
                      <div className="purchaseVLinks">
                        {purchase.links.map((link, index) => (
                          <a key={`${purchase.id}-v-link-${index}`} href={link} target="_blank" rel="noreferrer">
                            رابط {index + 1}
                          </a>
                        ))}
                      </div>
                    ) : null}

                    {canShowWhatsapp || menuNode ? (
                      <div className="wa-actions-row purchaseVWa">
                        {canShowWhatsapp ? (
                          <>
                            <button type="button" className="wa-btn wa-btn-inquiry" onClick={() => onInquireWhatsapp(purchase)}>
                              استعلام عن نقطة الاستلام ❓
                            </button>
                            <button type="button" className="wa-btn wa-btn-notify" onClick={() => onNotifyWhatsapp(purchase)}>
                              إعلام بوصول الطلب 🔔
                            </button>
                          </>
                        ) : null}
                        {menuNode ? <div className="purchaseVInlineMenu">{menuNode}</div> : null}
                      </div>
                    ) : null}
                  </article>
                </div>

                <div className="purchase-mobile-shell">
                  <div className="purchase-mobile-head">
                    <strong>{purchase.customer_name || "—"}</strong>
                  </div>

                  <div className="purchase-mobile-bubble">
                    <div className="purchase-mobile-images" dir="ltr">
                      {imageList.length ? (
                        imageList.map((img, index) => (
                          <button
                            key={img.id || `${purchase.id}-mobile-${index}`}
                            type="button"
                            className="purchase-mobile-image"
                            onClick={() => onOpenLightbox(imageList, index, purchase.customer_name || "صورة المشترى")}
                            aria-label={`عرض الصورة ${index + 1}`}
                          >
                            <img src={img.url} alt={`صورة ${index + 1}`} loading="lazy" />
                          </button>
                        ))
                      ) : (
                        <div className="purchase-mobile-empty-image">لا توجد صور</div>
                      )}
                    </div>

                    <div className="purchase-mobile-summary">
                      {purchase.qty || 0} قطع • {formatILS(purchase.price)} ₪
                    </div>

                    {purchase.links?.length ? (
                      <div className="purchase-mobile-links">
                        {purchase.links.map((link, index) => (
                          <a key={`${purchase.id}-m-link-${index}`} href={link} target="_blank" rel="noreferrer">
                            رابط {index + 1}
                          </a>
                        ))}
                      </div>
                    ) : null}

                    {canShowWhatsapp ? (
                      <div className="wa-actions-row purchase-mobile-wa">
                        <button type="button" className="wa-btn wa-btn-inquiry" onClick={() => onInquireWhatsapp(purchase)}>
                          استعلام عن نقطة الاستلام ❓
                        </button>
                        <button type="button" className="wa-btn wa-btn-notify" onClick={() => onNotifyWhatsapp(purchase)}>
                          إعلام بوصول الطلب 🔔
                        </button>
                      </div>
                    ) : null}
                  </div>

                  {menuNode ? (
                    <div className="purchase-mobile-actions">
                      <div className="purchase-head-actions">
                        {menuNode}
                      </div>
                    </div>
                  ) : null}
                </div>
              </article>
            );
          })}
        </div>
      ) : null}
    </>
  );
}
