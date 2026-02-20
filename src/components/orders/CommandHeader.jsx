import { useEffect, useState } from "react";
import SheStoreLogo from "../common/SheStoreLogo";

const MOBILE_BREAKPOINT = 768;
const DESKTOP_BREAKPOINT = 1024;

function getViewport() {
  if (typeof window === "undefined") return "desktop";
  if (window.innerWidth < MOBILE_BREAKPOINT) return "mobile";
  if (window.innerWidth < DESKTOP_BREAKPOINT) return "tablet";
  return "desktop";
}

function ListIcon({ className = "" }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <line x1="8" y1="6" x2="21" y2="6" />
      <line x1="8" y1="12" x2="21" y2="12" />
      <line x1="8" y1="18" x2="21" y2="18" />
      <circle cx="4" cy="6" r="1" />
      <circle cx="4" cy="12" r="1" />
      <circle cx="4" cy="18" r="1" />
    </svg>
  );
}

function ColumnsIcon({ className = "" }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <rect x="3" y="4" width="5" height="16" rx="1.5" />
      <rect x="10" y="4" width="5" height="16" rx="1.5" />
      <rect x="17" y="4" width="4" height="16" rx="1.5" />
    </svg>
  );
}

export default function CommandHeader({
  isRahaf,
  canAccessCustomers = false,
  activeTab,
  onActiveTabChange,
  search,
  onSearchChange,
  searchCount,
  onOpenSidebar,
  showOrdersMenuTrigger = false,
  onOpenOrdersMenu,
  totalOrders,
  showDesktopOrdersViewToggle = false,
  desktopOrdersView = "list",
  onDesktopOrdersViewChange,
  Icon
}) {
  const [viewport, setViewport] = useState(() => getViewport());

  const isMobile = viewport === "mobile";
  const isTablet = viewport === "tablet";
  const showOrdersCustomersTabs = isRahaf || canAccessCustomers;

  useEffect(() => {
    const onResize = () => setViewport(getViewport());
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);

  const renderTabs = (className = "") => {
    if (!showOrdersCustomersTabs) {
      return null;
    }

    return (
      <div className={`tabs-shell ${className}`.trim()} role="tablist" aria-label="التبويبات">
        <button
          type="button"
          className={`tab ${activeTab === "orders" ? "active" : ""}`}
          onClick={() => onActiveTabChange("orders")}
        >
          الطلبات
        </button>

        {canAccessCustomers ? (
          <button
            type="button"
            className={`tab ${activeTab === "customers" ? "active" : ""}`}
            onClick={() => onActiveTabChange("customers")}
          >
            العملاء
          </button>
        ) : null}
      </div>
    );
  };

  if (isMobile) {
    return (
      <header className="command-header command-header-mobile">
        <div className="command-logo-row">
          <SheStoreLogo className="she-store-logo-link command-logo-link" imageClassName="she-store-logo-img command-logo-img" />
        </div>
        <div className="command-mobile-row">
          <button
            type="button"
            className="icon-btn command-mobile-icon"
            onClick={onOpenSidebar}
            aria-label="فتح القائمة الجانبية"
          >
            <Icon name="menu" className="icon" />
          </button>

          <div className="command-mobile-title">
            <strong>الطلبات</strong>
            <small>{totalOrders} طلب</small>
          </div>

          <div className="command-mobile-search-inline">
            <form className="search-pill-form" onSubmit={(event) => event.preventDefault()}>
              <label htmlFor="ordersMobileSearch">بحث</label>
              <Icon name="search" className="search-pill-icon" />
              <input
                id="ordersMobileSearch"
                className="search-pill-input"
                type="search"
                value={search}
                onChange={(event) => onSearchChange(event.target.value)}
                placeholder="بحث..."
              />
              {search ? (
                <>
                  <button
                    type="button"
                    className="search-pill-clear"
                    onClick={() => onSearchChange("")}
                    aria-label="مسح البحث"
                  >
                    <Icon name="close" className="icon-sm" />
                  </button>
                  <span className="search-pill-count">{searchCount}</span>
                </>
              ) : null}
            </form>

            {showOrdersMenuTrigger ? (
              <button
                type="button"
                className="icon-btn command-mobile-icon orders-menu-trigger-btn"
                aria-label="فتح قائمة الطلبات"
                onClick={onOpenOrdersMenu}
              >
                <Icon name="package" className="icon" />
              </button>
            ) : null}
          </div>
        </div>
      </header>
    );
  }

  if (isTablet) {
    return (
      <header className="command-header command-header-tablet">
        <div className="command-logo-row">
          <SheStoreLogo className="she-store-logo-link command-logo-link" imageClassName="she-store-logo-img command-logo-img" />
        </div>
        <div className="command-tablet-row">
          <button
            type="button"
            className="icon-btn command-mobile-icon"
            onClick={onOpenSidebar}
            aria-label="فتح القائمة الجانبية"
          >
            <Icon name="menu" className="icon" />
          </button>

          {renderTabs("command-tablet-tabs")}

          <div className="search-shell command-tablet-search">
            <Icon name="search" className="search-icon" />
            <input value={search} onChange={(event) => onSearchChange(event.target.value)} placeholder="بحث..." />
            {search ? (
              <span className="search-count">
                <b>{searchCount}</b>
                <small>نتيجة</small>
              </span>
            ) : null}
          </div>

          {showOrdersMenuTrigger ? (
            <button
              type="button"
              className="icon-btn orders-menu-trigger-btn"
              onClick={onOpenOrdersMenu}
              aria-label="فتح قائمة الطلبات"
            >
              <Icon name="package" className="icon" />
            </button>
          ) : null}
        </div>
      </header>
    );
  }

  return (
    <header className="command-header">
      <div className="command-logo-row">
        <SheStoreLogo className="she-store-logo-link command-logo-link" imageClassName="she-store-logo-img command-logo-img" />
      </div>
      <div className="command-main command-main-group">
        {renderTabs()}

        <div className="search-shell command-search-group">
          <Icon name="search" className="search-icon" />
          <input value={search} onChange={(event) => onSearchChange(event.target.value)} placeholder="بحث..." />
          {search ? (
            <span className="search-count">
              <b>{searchCount}</b>
              <small>نتيجة</small>
            </span>
          ) : null}
        </div>
      </div>

      <div className="command-actions command-action-group" role="toolbar" aria-label="شريط أدوات الطلبات">
        {showDesktopOrdersViewToggle ? (
          <div className="view-controls-group">
            <span className="control-label">العرض:</span>
            <div className="orders-view-toggle" role="tablist" aria-label="طريقة عرض المشتريات">
              <button
                type="button"
                className={`orders-view-btn ${desktopOrdersView === "list" ? "active" : ""}`}
                onClick={() => onDesktopOrdersViewChange?.("list")}
              >
                <ListIcon className="icon-sm" />
                قائمة
              </button>
              <button
                type="button"
                className={`orders-view-btn ${desktopOrdersView === "kanban" ? "active" : ""}`}
                onClick={() => onDesktopOrdersViewChange?.("kanban")}
              >
                <ColumnsIcon className="icon-sm" />
                كانبان
              </button>
            </div>
          </div>
        ) : null}

        <div className="quick-actions-group">
          {showOrdersMenuTrigger ? (
            <button type="button" className="icon-btn orders-menu-trigger-btn" onClick={onOpenOrdersMenu} aria-label="فتح قائمة الطلبات">
              <Icon name="package" className="icon" />
            </button>
          ) : null}

          <button type="button" className="icon-btn" onClick={onOpenSidebar} aria-label="فتح القائمة الجانبية">
            <Icon name="menu" className="icon" />
          </button>
        </div>
      </div>
    </header>
  );
}
