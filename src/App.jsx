import { Link, Navigate, Route, Routes, useLocation, useParams } from "react-router-dom";
import SessionTimeoutGuard from "./components/SessionTimeoutGuard";
import OrdersPage from "./pages/OrdersPage";
import PickupDashboardPage from "./pages/PickupDashboardPage";
import FinancePage from "./pages/FinancePage";
import ArchivePage from "./pages/ArchivePage";
import CollectionsPage from "./pages/CollectionsPage";
import HomePickupPage from "./pages/HomePickupPage";
import PickupPointPage from "./pages/PickupPointPage";
import LoginPage from "./pages/LoginPage";
import WelcomeEntryPage from "./pages/WelcomeEntryPage";

const pages = [
  { id: "index", title: "الطلبات", file: "index.html", note: "لوحة الطلبات الرئيسية" },
  { id: "pickup-dashboard", title: "لوحة الاستلام", file: "pickup-dashboard.html", note: "متابعة الاستلام والتحصيل" },
  { id: "pickuppoint", title: "نقطة الاستلام", file: "pickuppoint.html", note: "تشغيل نقطة الاستلام" },
  { id: "archive", title: "الأرشيف", file: "archive.html", note: "الطلبات المؤرشفة" },
  { id: "finance", title: "المالية", file: "finance.html", note: "التقارير والتحصيل" },
  { id: "collections", title: "المجموعات", file: "collections.html", note: "إدارة المجموعات" },
  { id: "homepickup", title: "استلام المنزل", file: "homepickup.html", note: "متابعة توصيل المنزل" },
  { id: "home", title: "الصفحة العامة", file: "home.html", note: "صفحة العملاء" },
  { id: "login", title: "تسجيل الدخول", file: "login.html", note: "بوابة الدخول" }
];

const byId = Object.fromEntries(pages.map((p) => [p.id, p]));

function Dashboard() {
  const targetRoute = (pageId) => {
    if (pageId === "index") return "/orders";
    if (pageId === "pickup-dashboard") return "/pickup-dashboard";
    if (pageId === "finance") return "/finance";
    if (pageId === "archive") return "/archive";
    if (pageId === "collections") return "/collections";
    if (pageId === "homepickup") return "/homepickup";
    if (pageId === "pickuppoint") return "/pickuppoint";
    if (pageId === "login") return "/login";
    return `/legacy/${pageId}`;
  };

  return (
    <main className="app-shell">
      <header className="hero">
        <h1>She-Store React Migration</h1>
        <p>
          المشروع الآن React (Vite). تم حفظ النظام الحالي داخل <code>/public/legacy</code> حتى نقدر
          ننقل الصفحات تدريجيًا بدون كسر المنطق.
        </p>
      </header>

      <section className="hero quick-links">
        <h2>النسخة الجديدة الجاري نقلها</h2>
        <div className="quick-actions">
          <Link to="/orders" className="btn">
            فتح الطلبات (React Native)
          </Link>
          <Link to="/pickup-dashboard" className="btn">
            فتح لوحة الاستلام (React Native)
          </Link>
          <Link to="/legacy/index" className="btn btn-ghost">
            فتح الصفحة القديمة
          </Link>
        </div>
      </section>

      <section className="grid">
        {pages.map((page) => (
          <article key={page.id} className="card">
            <h2>{page.title}</h2>
            <p>{page.note}</p>
            <Link to={targetRoute(page.id)} className="btn">
              فتح
            </Link>
          </article>
        ))}
      </section>
    </main>
  );
}

function LegacyFrame() {
  const { page } = useParams();
  const location = useLocation();
  const target = byId[page || ""];

  if (!target) {
    return <Navigate to="/" replace />;
  }

  const src = `legacy/${target.file}${location.search || ""}`;

  return (
    <main className="legacy-shell">
      <header className="legacy-topbar">
        <div>
          <h1>{target.title}</h1>
          <p>{target.note}</p>
        </div>
        <div className="legacy-actions">
          <Link to="/" className="btn btn-ghost">
            رجوع
          </Link>
          <a className="btn" href={src} target="_blank" rel="noreferrer">
            فتح بنافذة جديدة
          </a>
        </div>
      </header>

      <iframe title={target.title} src={src} className="legacy-frame" />
    </main>
  );
}

export default function App() {
  return (
    <>
      <SessionTimeoutGuard />
      <Routes>
        <Route path="/" element={<WelcomeEntryPage />} />
        <Route path="/migration" element={<Dashboard />} />
        <Route path="/orders" element={<OrdersPage />} />
        <Route path="/pickup-dashboard" element={<PickupDashboardPage />} />
        <Route path="/finance" element={<FinancePage />} />
        <Route path="/archive" element={<ArchivePage />} />
        <Route path="/collections" element={<CollectionsPage />} />
        <Route path="/homepickup" element={<HomePickupPage />} />
        <Route path="/pickuppoint" element={<PickupPointPage />} />
        <Route path="/login" element={<LoginPage />} />
        <Route path="/legacy/:page" element={<LegacyFrame />} />
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </>
  );
}
