import { useEffect, useMemo, useState } from "react";
import { useLocation } from "react-router-dom";
import { useAuthProfile } from "../hooks/useAuthProfile";
import { getOrdersNavItems, getPickupDashboardTabs, getRoleLabel, isNavHrefActive } from "../lib/navigation";
import { setBodyScrollLock } from "../lib/bodyScrollLock";
import { signOutAndRedirect } from "../lib/session";
import AppNavIcon from "../components/common/AppNavIcon";
import SessionLoader from "../components/common/SessionLoader";
import HomePickupPage from "./HomePickupPage";
import PickupPointPage from "./PickupPointPage";
import CollectionsPage from "./CollectionsPage";
import "./pickup-dashboard-page.css";
import SheStoreLogo from "../components/common/SheStoreLogo";
import homePickupsIcon from "../assets/icons/pickup-dashboard/home-pickups.png";
import laauraPickupsIcon from "../assets/icons/pickup-dashboard/laaura-pickups.png";
import moneyCollectionsIcon from "../assets/icons/pickup-dashboard/money-collections.png";

const TAB_CONFIG = {
  home: { id: "home", label: "مستلمو البيت", icon: homePickupsIcon },
  aura: { id: "aura", label: "La Aura", icon: laauraPickupsIcon },
  collections: { id: "collections", label: "تحصيل المبالغ", icon: moneyCollectionsIcon }
};

export default function PickupDashboardPage() {
  const { profile } = useAuthProfile();
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [activeTab, setActiveTab] = useState("home");
  const location = useLocation();

  const roleTabs = useMemo(() => getPickupDashboardTabs(profile.role), [profile.role]);
  const sidebarLinks = useMemo(() => getOrdersNavItems(profile.role), [profile.role]);
  const showSidebar = profile.role !== "laaura";

  useEffect(() => {
    const firstTab = roleTabs[0];
    if (!firstTab) return;
    setActiveTab((prev) => (roleTabs.includes(prev) ? prev : firstTab));
  }, [roleTabs]);

  useEffect(() => {
    const onKeyDown = (event) => {
      if (event.key === "Escape") setSidebarOpen(false);
    };
    document.addEventListener("keydown", onKeyDown);
    return () => document.removeEventListener("keydown", onKeyDown);
  }, []);

  useEffect(() => {
    setBodyScrollLock("pickup-dashboard-sidebar", sidebarOpen);
    return () => {
      setBodyScrollLock("pickup-dashboard-sidebar", false);
    };
  }, [sidebarOpen]);

  useEffect(() => {
    setSidebarOpen(false);
  }, [location.pathname, location.search, location.hash]);

  async function signOut() {
    await signOutAndRedirect();
  }

  function renderPanel(tabId) {
    if (tabId === "home") return <HomePickupPage embedded />;
    if (tabId === "aura") return <PickupPointPage embedded />;
    if (tabId === "collections") return <CollectionsPage embedded />;
    return null;
  }

  if (profile.loading) {
    return (
      <div className="pickup-page pickup-state">
        <SessionLoader label="جاري التحميل..." />
      </div>
    );
  }

  if (!profile.authenticated) {
    return (
      <div className="pickup-page pickup-state">
        <div className="pickup-note pickup-note-danger">
          <h2>لا توجد جلسة نشطة</h2>
          <p>يلزم تسجيل الدخول أولًا للوصول إلى لوحة الاستلام.</p>
          <a href="#/login" className="pickup-btn pickup-btn-primary">
            تسجيل الدخول
          </a>
        </div>
      </div>
    );
  }

  if (!roleTabs.length) {
    return (
      <div className="pickup-page pickup-state">
        <div className="pickup-note pickup-note-danger">
          <h2>لا توجد صلاحية</h2>
          <p>هذا الحساب لا يملك صلاحية الوصول إلى لوحة الاستلام والتحصيل.</p>
          <a href="#/orders" className="pickup-btn pickup-btn-primary">
            العودة للطلبات
          </a>
        </div>
      </div>
    );
  }

  return (
    <div className="pickup-page" dir="rtl">
      {showSidebar ? (
        <>
          <div
            className={`pickup-sidebar-overlay app-sidebar-overlay ${sidebarOpen ? "open" : ""}`}
            onClick={() => setSidebarOpen(false)}
            aria-hidden="true"
          />

          <aside
            className={`pickup-sidebar app-sidebar-drawer ${sidebarOpen ? "open" : ""}`}
            role="navigation"
            aria-label="القائمة الجانبية"
          >
            <div className="app-sidebar-head">
              <div className="app-sidebar-brand">
                <SheStoreLogo className="app-sidebar-logo-link" imageClassName="app-sidebar-logo-img" />
                <b>القائمة</b>
              </div>

              <button
                type="button"
                className="app-sidebar-close"
                onClick={() => setSidebarOpen(false)}
                aria-label="إغلاق القائمة"
              >
                ✕
              </button>
            </div>

            <div className="app-sidebar-content">
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

              <button type="button" className="app-sidebar-link app-sidebar-danger" onClick={signOut}>
                تسجيل الخروج
              </button>
            </div>
          </aside>
        </>
      ) : null}

      <div className="pickup-wrap">
        <header className="pickup-command-header" role="banner">
          <div className="pickup-command-main">
            <div className="pickup-brand">
              <SheStoreLogo className="she-store-logo-link" imageClassName="she-store-logo-img" />
              <div>
                <b>لوحة الاستلام</b>
                <div className="pickup-muted">متابعة الاستلام والتحصيل</div>
              </div>
            </div>
          </div>

          <div className="pickup-command-meta">
            <span className="pickup-role-chip">{getRoleLabel(profile.role)}</span>
            <div className="pickup-command-actions">
              {showSidebar ? (
                <button
                  type="button"
                  className="pickup-btn pickup-btn-icon"
                  onClick={() => setSidebarOpen(true)}
                  aria-label="فتح القائمة"
                >
                  ☰
                </button>
              ) : (
                <button type="button" className="pickup-btn" onClick={signOut}>
                  تسجيل خروج
                </button>
              )}
            </div>
          </div>

          </header>

        <div className="pickup-tabs-container">
          <div className="pickup-tabs" role="tablist" aria-label="أقسام لوحة الاستلام">
            {roleTabs.map((tabId) => {
              const config = TAB_CONFIG[tabId];
              if (!config) return null;
              return (
                <button
                  key={tabId}
                  type="button"
                  role="tab"
                  aria-selected={activeTab === tabId}
                  aria-controls={`panel-${tabId}`}
                  className={`pickup-tab-btn ${activeTab === tabId ? "active" : ""}`}
                  onClick={() => setActiveTab(tabId)}
                >
                  <img src={config.icon} alt="" className="pickup-tab-icon" aria-hidden="true" />
                  <span>{config.label}</span>
                </button>
              );
            })}
          </div>

          <div
            id={`panel-${activeTab}`}
            role="tabpanel"
            aria-labelledby={`tab-${activeTab}`}
            className="pickup-panel active"
          >
            <div className="pickup-panel-content">
              <div className="pickup-panel-scroll-shell">{renderPanel(activeTab)}</div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
