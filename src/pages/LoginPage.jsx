import { useEffect, useMemo, useState } from "react";
import { useLocation, useNavigate } from "react-router-dom";
import { sb } from "../lib/supabaseClient";
import SessionLoader from "../components/common/SessionLoader";
import SheStoreLogo from "../components/common/SheStoreLogo";

const EMAIL_DOMAIN = "she-store.com";
const REMEMBER_ME_STORAGE_KEY = "she_store:remember_me";

const LEGACY_ROUTE_MAP = {
  "index.html": "/orders",
  "pickup-dashboard.html": "/pickup-dashboard",
  "pickuppoint.html": "/pickuppoint",
  "archive.html": "/archive",
  "finance.html": "/finance",
  "collections.html": "/collections",
  "homepickup.html": "/homepickup",
  "login.html": "/login",
  "home.html": "/legacy/home"
};

function usernameToEmail(username) {
  const value = String(username || "").trim().toLowerCase();
  if (value.includes("@")) return value;
  const safe = value.replace(/[^a-z0-9._-]/g, "");
  return `${safe}@${EMAIL_DOMAIN}`;
}

function normalizeNextRoute(rawNext) {
  const fallback = "/orders";
  if (!rawNext) return fallback;

  let target = String(rawNext).trim();
  if (!target) return fallback;

  if (target.startsWith("#/")) target = target.slice(1);

  if (!target.startsWith("/")) {
    const filePart = target.split(/[?#]/)[0].split("/").pop()?.toLowerCase();
    if (filePart && LEGACY_ROUTE_MAP[filePart]) return LEGACY_ROUTE_MAP[filePart];
    return fallback;
  }

  if (target.startsWith("/legacy/")) return target;

  const allowedRoutes = [
    "/",
    "/orders",
    "/pickup-dashboard",
    "/pickuppoint",
    "/archive",
    "/finance",
    "/collections",
    "/homepickup",
    "/login"
  ];

  if (allowedRoutes.some((route) => target === route || target.startsWith(`${route}?`))) {
    return target;
  }

  return fallback;
}

function readRememberPreference() {
  try {
    return window.localStorage.getItem(REMEMBER_ME_STORAGE_KEY) === "1";
  } catch {
    return false;
  }
}

function writeRememberPreference(enabled) {
  try {
    window.localStorage.setItem(REMEMBER_ME_STORAGE_KEY, enabled ? "1" : "0");
  } catch {
    // no-op
  }
}

export default function LoginPage() {
  const location = useLocation();
  const navigate = useNavigate();

  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [rememberMe, setRememberMe] = useState(() => readRememberPreference());
  const [showPassword, setShowPassword] = useState(false);
  const [focused, setFocused] = useState(null);
  const [loading, setLoading] = useState(false);
  const [success, setSuccess] = useState(false);
  const [error, setError] = useState("");
  const [checkingSession, setCheckingSession] = useState(true);

  const nextRoute = useMemo(() => {
    const params = new URLSearchParams(location.search || "");
    return normalizeNextRoute(params.get("next"));
  }, [location.search]);

  useEffect(() => {
    let mounted = true;

    async function checkSession() {
      try {
        const {
          data: { session }
        } = await sb.auth.getSession();

        if (!mounted) return;
        if (session) {
          navigate(nextRoute, { replace: true });
          return;
        }
      } catch (sessionError) {
        console.error(sessionError);
      } finally {
        if (mounted) setCheckingSession(false);
      }
    }

    checkSession();
    return () => {
      mounted = false;
    };
  }, [navigate, nextRoute]);

  const handleLogin = async () => {
    if (loading || success) return;

    if (!username || !password) {
      setError("Please enter your username and password");
      window.setTimeout(() => setError(""), 2500);
      return;
    }

    setError("");
    setLoading(true);

    writeRememberPreference(rememberMe);

    const email = usernameToEmail(username);
    const { error: loginError } = await sb.auth.signInWithPassword({ email, password });

    if (loginError) {
      setLoading(false);
      setError(`Failed to sign in: ${loginError.message}`);
      window.setTimeout(() => setError(""), 3500);
      return;
    }

    setLoading(false);
    setSuccess(true);
    window.setTimeout(() => {
      navigate(nextRoute, { replace: true });
    }, 2350);
  };

  if (checkingSession) {
    return (
      <div
        dir="ltr"
        style={{
          minHeight: "100vh",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          background: "linear-gradient(145deg, #f6f2f8 0%, #ede5f1 40%, #e7dfeb 100%)"
        }}
      >
        <SessionLoader />
      </div>
    );
  }

  return (
    <div
      dir="ltr"
      style={{
        minHeight: "100vh",
        background: "linear-gradient(145deg, #f6f2f8 0%, #ede5f1 40%, #e7dfeb 100%)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        direction: "ltr",
        fontFamily: "'DM Sans', sans-serif",
        padding: "20px",
        position: "relative",
        overflow: "hidden"
      }}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;0,9..40,600;1,9..40,400&family=Fraunces:ital,wght@0,300;0,400;0,600;1,300;1,400&display=swap');

        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

        ::placeholder { color: #a0b4ae !important; }

        @keyframes fadeUp {
          from { opacity: 0; transform: translateY(24px); }
          to   { opacity: 1; transform: translateY(0); }
        }
        @keyframes fadeIn {
          from { opacity: 0; } to { opacity: 1; }
        }
        @keyframes shake {
          0%, 100% { transform: translateX(0); }
          20% { transform: translateX(-6px); }
          40% { transform: translateX(6px); }
          60% { transform: translateX(-4px); }
          80% { transform: translateX(4px); }
        }
        @keyframes spin {
          to { transform: rotate(360deg); }
        }
        @keyframes successPop {
          0%   { transform: scale(0.8); opacity: 0; }
          60%  { transform: scale(1.08); }
          100% { transform: scale(1); opacity: 1; }
        }
        @keyframes blobFloat1 {
          0%,100% { transform: translate(0,0) scale(1); }
          50%     { transform: translate(30px,-20px) scale(1.04); }
        }
        @keyframes blobFloat2 {
          0%,100% { transform: translate(0,0) scale(1); }
          50%     { transform: translate(-20px,25px) scale(1.03); }
        }

        .login-card {
          animation: fadeUp 0.65s cubic-bezier(0.16,1,0.3,1) both;
        }
        .login-logo-wrap {
          margin: 8px 0 26px;
          display: flex;
          justify-content: center;
        }
        .login-logo-link {
          border-radius: 14px;
        }
        .login-logo-img {
          width: clamp(200px, 40vw, 250px);
          display: block;
          height: auto;
          object-fit: contain;
        }
        .field-wrap { animation: fadeUp 0.65s cubic-bezier(0.16,1,0.3,1) both; }
        .field-wrap:nth-child(1) { animation-delay: 0.1s; }
        .field-wrap:nth-child(2) { animation-delay: 0.18s; }
        .field-wrap:nth-child(3) { animation-delay: 0.26s; }
        .field-wrap:nth-child(4) { animation-delay: 0.34s; }

        .login-input {
          width: 100%;
          padding: 13px 16px 13px 44px;
          border-radius: 12px;
          border: 1.5px solid #e7dfeb;
          background: #ffffff;
          color: #111827;
          font-size: 15px;
          font-family: 'DM Sans', sans-serif;
          font-weight: 400;
          outline: none;
          transition: border-color 0.2s, box-shadow 0.2s, background 0.2s;
          -webkit-appearance: none;
        }
        .login-input:hover { border-color: #8F6E97; }
        .login-input:focus {
          border-color: #6D3F6B;
          box-shadow: 0 0 0 3.5px rgba(109, 63, 107, 0.12);
          background: #fff;
        }
        .login-input.error-state {
          border-color: #b42318 !important;
          box-shadow: 0 0 0 3px rgba(180,35,24,0.1) !important;
          animation: shake 0.4s ease;
        }

        .login-btn {
          width: 100%;
          padding: 14px;
          border-radius: 12px;
          border: none;
          background: #47143F;
          color: #fff;
          font-size: 15px;
          font-weight: 600;
          font-family: 'DM Sans', sans-serif;
          letter-spacing: 0.02em;
          cursor: pointer;
          display: flex;
          align-items: center;
          justify-content: center;
          gap: 8px;
          transition: transform 0.2s, box-shadow 0.2s, background 0.2s;
          box-shadow: 0 4px 16px rgba(71, 20, 63, 0.25);
          position: relative;
          overflow: hidden;
        }
        .login-btn:hover:not(:disabled) {
          transform: translateY(-2px);
          box-shadow: 0 8px 28px rgba(71, 20, 63, 0.32);
          background: #47143F;
        }
        .login-btn:active:not(:disabled) { transform: translateY(0); }
        .login-btn:disabled { opacity: 0.85; cursor: not-allowed; }

        .remember-check {
          width: 18px; height: 18px;
          border-radius: 5px;
          border: 1.5px solid #8F6E97;
          background: #fff;
          cursor: pointer;
          display: flex; align-items: center; justify-content: center;
          transition: all 0.18s;
          flex-shrink: 0;
        }
        .remember-check.checked {
          background: #6D3F6B;
          border-color: #6D3F6B;
        }
        .remember-row {
          display: flex; align-items: center; gap: 9px;
          cursor: pointer; user-select: none;
        }

        .eye-btn {
          position: absolute; right: 14px; top: 50%;
          transform: translateY(-50%);
          background: none; border: none;
          color: #a0b4ae; cursor: pointer;
          display: flex; align-items: center;
          padding: 2px;
          transition: color 0.2s;
        }
        .eye-btn:hover { color: #6D3F6B; }

        .icon-prefix {
          position: absolute; left: 14px; top: 50%;
          transform: translateY(-50%);
          pointer-events: none;
          transition: color 0.2s;
          display: flex; align-items: center;
        }

        .success-overlay {
          position: absolute; inset: 0;
          background: #fff;
          border-radius: 24px;
          display: flex; flex-direction: column;
          align-items: center; justify-content: center;
          gap: 14px;
          animation: fadeIn 0.3s ease;
          z-index: 10;
        }
        .success-icon {
          width: 64px; height: 64px;
          border-radius: 50%;
          background: #47143F;
          display: flex; align-items: center; justify-content: center;
          box-shadow: 0 8px 24px rgba(71, 20, 63, 0.3);
          animation: successPop 0.5s cubic-bezier(0.16,1,0.3,1) both;
        }

        .blob1 {
          position: absolute; border-radius: 50%;
          width: 420px; height: 420px;
          background: radial-gradient(circle, rgba(109, 63, 107, 0.13) 0%, transparent 70%);
          top: -80px; left: -100px;
          filter: blur(40px);
          animation: blobFloat1 10s ease-in-out infinite;
          pointer-events: none;
        }
        .blob2 {
          position: absolute; border-radius: 50%;
          width: 380px; height: 380px;
          background: radial-gradient(circle, rgba(71, 20, 63, 0.1) 0%, transparent 70%);
          bottom: -60px; right: -80px;
          filter: blur(50px);
          animation: blobFloat2 13s ease-in-out infinite;
          pointer-events: none;
        }

        .divider-dot {
          width: 4px; height: 4px;
          border-radius: 50%;
          background: #8F6E97;
          display: inline-block;
        }

        @media (max-width: 480px) {
          .login-card { padding: 32px 24px !important; }
        }
      `}</style>

      <div className="blob1" />
      <div className="blob2" />

      <div
        style={{
          position: "absolute",
          inset: 0,
          pointerEvents: "none",
          backgroundImage: "radial-gradient(circle, rgba(109, 63, 107, 0.08) 1px, transparent 1px)",
          backgroundSize: "28px 28px"
        }}
      />

      <div
        className="login-card"
        style={{
          width: "100%",
          maxWidth: 440,
          background: "#ffffff",
          borderRadius: 24,
          padding: "44px 40px",
          boxShadow: "0 2px 8px rgba(71, 20, 63, 0.06), 0 16px 48px rgba(71, 20, 63, 0.1), 0 1px 0 rgba(255,255,255,0.8) inset",
          border: "1px solid rgba(221,231,227,0.8)",
          position: "relative",
          overflow: "hidden"
        }}
      >
        {success && (
          <div className="success-overlay">
            <div className="success-icon">
              <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="20 6 9 17 4 12" />
              </svg>
            </div>
            <p style={{ fontFamily: "'Fraunces', serif", fontSize: 20, color: "#111827", fontWeight: 600 }}>Welcome back!</p>
            <p style={{ color: "#64748b", fontSize: 14 }}>Redirecting to dashboard…</p>
          </div>
        )}

        <div
          style={{
            position: "absolute",
            top: 0,
            left: 0,
            right: 0,
            height: 4,
            background: "#47143F",
            borderRadius: "24px 24px 0 0"
          }}
        />

        <div className="login-logo-wrap">
          <SheStoreLogo
            href="#/"
            className="she-store-logo-link login-logo-link"
            imageClassName="she-store-logo-img login-logo-img"
            ariaLabel="She-Store Home"
          />
        </div>

        <div style={{ marginBottom: 36, animation: "fadeUp 0.5s both" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 28 }}>
            <div
              style={{
                width: 38,
                height: 38,
                background: "#47143F",
                borderRadius: 10,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                boxShadow: "0 4px 12px rgba(71, 20, 63, 0.25)"
              }}
            >
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M6 2L3 6v14a2 2 0 002 2h14a2 2 0 002-2V6l-3-4z" />
                <line x1="3" y1="6" x2="21" y2="6" />
                <path d="M16 10a4 4 0 01-8 0" />
              </svg>
            </div>
            <div>
              <p style={{ fontSize: 11, fontWeight: 600, color: "#6D3F6B", letterSpacing: "0.1em", textTransform: "uppercase" }}>Order Management</p>
              <p style={{ fontSize: 13, color: "#64748b", fontWeight: 400, marginTop: 1 }}>Shein Agent Portal</p>
            </div>
          </div>

          <h1
            style={{
              fontFamily: "'Fraunces', serif",
              fontSize: 30,
              fontWeight: 600,
              color: "#111827",
              letterSpacing: "-0.03em",
              lineHeight: 1.15,
              marginBottom: 6
            }}
          >
            Sign in to your
            <br />
            <span style={{ color: "#6D3F6B", fontStyle: "italic" }}>workspace</span>
          </h1>
          <p style={{ color: "#64748b", fontSize: 14, fontWeight: 400 }}>Manage orders and pickup points in one place</p>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
          <div className="field-wrap">
            <label
              style={{
                display: "block",
                marginBottom: 7,
                fontSize: 13,
                fontWeight: 500,
                color: "#374151",
                letterSpacing: "0.01em"
              }}
            >
              Username
            </label>
            <div style={{ position: "relative" }}>
              <span className="icon-prefix" style={{ color: focused === "username" ? "#6D3F6B" : "#a0b4ae" }}>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2" />
                  <circle cx="12" cy="7" r="4" />
                </svg>
              </span>
              <input
                className={`login-input${error && !username ? " error-state" : ""}`}
                type="text"
                value={username}
                onChange={(e) => setUsername(e.target.value)}
                onFocus={() => setFocused("username")}
                onBlur={() => setFocused(null)}
                onKeyDown={(e) => e.key === "Enter" && handleLogin()}
                placeholder="Enter your username"
                autoComplete="username"
              />
            </div>
          </div>

          <div className="field-wrap">
            <label
              style={{
                display: "block",
                marginBottom: 7,
                fontSize: 13,
                fontWeight: 500,
                color: "#374151"
              }}
            >
              Password
            </label>
            <div style={{ position: "relative" }}>
              <span className="icon-prefix" style={{ color: focused === "password" ? "#6D3F6B" : "#a0b4ae" }}>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                  <rect x="5" y="11" width="14" height="10" rx="2" ry="2" />
                  <path d="M8 11V7a4 4 0 018 0v4" />
                </svg>
              </span>
              <input
                className={`login-input${error && !password ? " error-state" : ""}`}
                type={showPassword ? "text" : "password"}
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                onFocus={() => setFocused("password")}
                onBlur={() => setFocused(null)}
                onKeyDown={(e) => e.key === "Enter" && handleLogin()}
                placeholder="Enter your password"
                autoComplete="current-password"
                style={{ paddingRight: 44 }}
              />
              <button className="eye-btn" type="button" onClick={() => setShowPassword(!showPassword)} tabIndex={-1}>
                {showPassword ? (
                  <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M17.94 17.94A10.07 10.07 0 0112 20c-7 0-11-8-11-8a18.45 18.45 0 015.06-5.94M9.9 4.24A9.12 9.12 0 0112 4c7 0 11 8 11 8a18.5 18.5 0 01-2.16 3.19m-6.72-1.07a3 3 0 11-4.24-4.24" />
                    <line x1="1" y1="1" x2="23" y2="23" />
                  </svg>
                ) : (
                  <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" />
                    <circle cx="12" cy="12" r="3" />
                  </svg>
                )}
              </button>
            </div>
          </div>

          <div className="field-wrap" style={{ marginTop: -4 }}>
            <div
              className="remember-row"
              onClick={() => {
                const next = !rememberMe;
                setRememberMe(next);
                writeRememberPreference(next);
              }}
            >
              <div className={`remember-check${rememberMe ? " checked" : ""}`}>
                {rememberMe && (
                  <svg width="11" height="11" viewBox="0 0 12 12" fill="none" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <polyline points="2 6 5 9 10 3" />
                  </svg>
                )}
              </div>
              <span style={{ fontSize: 13, color: "#374151", fontWeight: 400 }}>Remember me on this device</span>
            </div>
          </div>

          {error && (
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 8,
                padding: "10px 14px",
                background: "#fff6f6",
                border: "1px solid #f4bcbc",
                borderRadius: 10,
                animation: "fadeIn 0.2s ease"
              }}
            >
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#b42318" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0 }}>
                <circle cx="12" cy="12" r="10" />
                <line x1="12" y1="8" x2="12" y2="12" />
                <line x1="12" y1="16" x2="12.01" y2="16" />
              </svg>
              <span style={{ fontSize: 13, color: "#b42318", fontWeight: 500 }}>{error}</span>
            </div>
          )}

          <div className="field-wrap" style={{ marginTop: 4 }}>
            <button className="login-btn" type="button" onClick={handleLogin} disabled={loading || success}>
              {loading ? (
                <>
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" style={{ animation: "spin 0.75s linear infinite" }}>
                    <path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83" />
                  </svg>
                  Signing in…
                </>
              ) : (
                <>
                  Sign in
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M5 12h14M12 5l7 7-7 7" />
                  </svg>
                </>
              )}
            </button>
          </div>
        </div>

        <div
          style={{
            marginTop: 28,
            paddingTop: 20,
            borderTop: "1px solid #f6f2f8",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            gap: 10
          }}
        >
          <span className="divider-dot" />
          <span style={{ fontSize: 12, color: "#a0b4ae", textAlign: "center" }}>Having trouble? Contact your administrator</span>
          <span className="divider-dot" />
        </div>
      </div>
    </div>
  );
}
