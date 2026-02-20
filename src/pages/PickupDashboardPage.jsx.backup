import { useEffect, useMemo, useState } from "react";
import { useLocation } from "react-router-dom";
import { useAuthProfile } from "../hooks/useAuthProfile";
import { getOrdersNavItems, getPickupDashboardTabs, getRoleLabel, isNavHrefActive } from "../lib/navigation";
import { signOutAndRedirect } from "../lib/session";
import HomePickupPage from "./HomePickupPage";
import PickupPointPage from "./PickupPointPage";
import CollectionsPage from "./CollectionsPage";
import "./pickup-dashboard-page.css";
import SessionLoader from "../components/common/SessionLoader";
import SheStoreLogo from "../components/common/SheStoreLogo";

const TAB_CONFIG = {
  home: { id: "home", label: "مستلمو البيت" },
  aura: { id: "aura", label: "La Aura" },
  collections: { id: "collections", label: "تحصيل المبالغ" }
};

export default function PickupDashboardPage() {
  const { profile } = useAuthProfile();
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [activeTab, setActiveTab] = useState("home");
  const [loadedTabs, setLoadedTabs] = useState([]);
  const location = useLocation();

  const roleTabs = useMemo(() => getPickupDashboardTabs(profile.role), [profile.role]);
  const sidebarLinks = useMemo(() => getOrdersNavItems(profile.role), [profile.role]);
  const showSidebar = profile.role !== "laaura";

  useEffect(() => {
    const firstTab = roleTabs[0];
    if (!firstTab) return;
    setActiveTab((prev) => (roleTabs.includes(prev) ? prev : firstTab));
    setLoadedTabs((prev) => (prev.includes(firstTab) ? prev : [...prev, firstTab]));
  }, [roleTabs]);

  useEffect(() => {
    if (!activeTab) return;
    setLoadedTabs((prev) => (prev.includes(activeTab) ? prev : [...prev, activeTab]));
  }, [activeTab]);

  useEffect(() => {
    const onKeyDown = (event) => {
      if (event.key === "Escape") setSidebarOpen(false);
    };
    document.addEventListener("keydown", onKeyDown);
    return () => document.removeEventListener("keydown", onKeyDown);
  }, []);

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
        <SessionLoader />
      </div>
    );
  }

  if (!profile.authenticated) {
    return (
      <div className="pickup-page pickup-state">
        <div className="pickup-note pickup-note-danger">
          <h2>لا توجد جلسة نشطة</h2>
          <p>يلزم تسجيل الدخول أولًا.</p>
          <a href="#/login" className="pickup-btn">
            فتح تسجيل الدخول
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
          <p>هذا الحساب لا يملك صلاحية لوحة الاستلام والتحصيل.</p>
          <a href="#/orders" className="pickup-btn">
            العودة للطلبيات
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
          />
          <aside className={`pickup-sidebar app-sidebar-drawer ${sidebarOpen ? "open" : ""}`}>
            <div className="pickup-sidebar-head app-sidebar-head">
              <div className="app-sidebar-brand">
                <SheStoreLogo className="app-sidebar-logo-link" imageClassName="app-sidebar-logo-img" />
                <b>القائمة</b>
              </div>
              <button
                type="button"
                className="pickup-btn pickup-btn-icon danger app-sidebar-close"
                onClick={() => setSidebarOpen(false)}
              >
                ✕
              </button>
            </div>
            <div className="pickup-sidebar-content app-sidebar-content">
              {sidebarLinks.map((item) => (
                <a
                  key={item.href}
                  href={item.href}
                  className={`app-sidebar-link ${isNavHrefActive(item.href, location) ? "active" : ""}`}
                  onClick={() => setSidebarOpen(false)}
                >
                  {item.label}
                </a>
              ))}
              <button type="button" className="danger app-sidebar-link app-sidebar-danger" onClick={signOut}>
                تسجيل خروج
              </button>
            </div>
          </aside>
        </>
      ) : null}

      <div className="pickup-wrap">
        <div className="pickup-topbar">
          <div className="topbar-brand-with-logo">
            <SheStoreLogo className="topbar-logo-link" imageClassName="topbar-logo-img" />
            <div className="pickup-brand">
              <b>الاستلام والتحصيل</b>
              <div className="pickup-muted">
                {profile.role === "laaura" ? "نقطة La Aura" : "من البيت + La Aura + التحصيل"} — {getRoleLabel(profile.role)}
              </div>
            </div>
          </div>
          {showSidebar ? (
            <button type="button" className="pickup-btn pickup-btn-icon" onClick={() => setSidebarOpen(true)}>
              ☰
            </button>
          ) : (
            <button type="button" className="pickup-btn danger" onClick={signOut}>
              تسجيل خروج
            </button>
          )}
        </div>

        {roleTabs.length > 1 ? (
          <div className="pickup-tabs">
            {roleTabs.map((tabId) => (
              <button
                key={tabId}
                type="button"
                className={`pickup-tab-btn ${activeTab === tabId ? "active" : ""}`}
                onClick={() => setActiveTab(tabId)}
              >
                {TAB_CONFIG[tabId].label}
              </button>
            ))}
          </div>
        ) : null}

        {roleTabs.map((tabId) => (
          <div key={tabId} className={`pickup-panel ${activeTab === tabId ? "active" : ""}`}>
            {loadedTabs.includes(tabId) ? <div className="pickup-panel-content">{renderPanel(tabId)}</div> : null}
          </div>
        ))}
      </div>
    </div>
  );
}
