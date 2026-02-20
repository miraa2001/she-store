import sheStoreLogoSrc from "../assets/She Store Logo - Upscaled.png";
import TypewriterEffect from "../components/common/TypewriterEffect";

export default function WelcomePage({ onContinue }) {
  return (
    <div
      style={{
        minHeight: "100vh",
        background: "linear-gradient(145deg, #f6f2f8 0%, #ede5f1 50%, #e7dfeb 100%)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        fontFamily: "'Nunito', sans-serif",
        padding: "20px",
        position: "relative",
        overflow: "hidden"
      }}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Nunito:wght@300;400;600;700;800&family=Quicksand:wght@400;500;600;700&display=swap');

        * { box-sizing: border-box; margin: 0; padding: 0; }

        @keyframes gentleFloat {
          0%, 100% { transform: translateY(0px); }
          50% { transform: translateY(-12px); }
        }
        @keyframes fadeIn {
          from { opacity: 0; transform: translateY(20px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes scaleIn {
          from { opacity: 0; transform: scale(0.92); }
          to { opacity: 1; transform: scale(1); }
        }
        @keyframes shimmer {
          0% { background-position: -200% center; }
          100% { background-position: 200% center; }
        }
        @keyframes blobPulse {
          0%, 100% { transform: scale(1) translate(0, 0); opacity: 0.15; }
          50% { transform: scale(1.1) translate(10px, -10px); opacity: 0.25; }
        }

        .welcome-container {
          animation: scaleIn 0.65s cubic-bezier(0.16, 1, 0.3, 1) forwards;
        }

        .hero-image {
          animation: gentleFloat 5s ease-in-out infinite;
          filter: drop-shadow(0 16px 32px rgba(109, 63, 107, 0.2));
        }

        .feature-card {
          animation: fadeIn 0.6s cubic-bezier(0.16, 1, 0.3, 1) both;
        }
        .feature-card:nth-child(1) { animation-delay: 0.15s; }
        .feature-card:nth-child(2) { animation-delay: 0.25s; }
        .feature-card:nth-child(3) { animation-delay: 0.35s; }

        .welcome-title {
          font-family: 'Quicksand', sans-serif;
          font-size: 34px;
          font-weight: 700;
          color: #111827;
          margin-bottom: 10px;
          line-height: 1.2;
          letter-spacing: -0.02em;
        }

        .auth-welcome-typewriter {
          font-family: 'Quicksand', sans-serif;
          font-size: 34px;
          font-weight: 700;
          letter-spacing: -0.02em;
          margin-bottom: 10px;
        }

        .auth-welcome-word {
          color: #111827;
        }

        .auth-welcome-word--accent {
          color: #6D3F6B;
        }

        .auth-welcome-cursor {
          background: #6D3F6B;
          height: 1.05em;
        }

        .welcome-subtitle {
          font-size: 15px;
          color: #64748b;
          font-weight: 400;
          line-height: 1.6;
          max-width: 360px;
          margin: 0 auto;
        }

        .login-button {
          background: #47143F;
          color: #fff;
          border: none;
          border-radius: 14px;
          padding: 15px 40px;
          font-size: 16px;
          font-weight: 700;
          font-family: 'Nunito', sans-serif;
          cursor: pointer;
          box-shadow: 0 8px 24px rgba(109, 63, 107, 0.3);
          position: relative;
          overflow: hidden;
          transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }
        .login-button::before {
          content: '';
          position: absolute;
          inset: 0;
          background: rgba(255, 255, 255, 0.12);
          opacity: 0;
          transition: opacity 0.3s;
        }
        .login-button:hover {
          transform: translateY(-3px);
          box-shadow: 0 12px 32px rgba(109, 63, 107, 0.4);
        }
        .login-button:hover::before {
          opacity: 1;
        }
        .login-button:active {
          transform: translateY(-1px);
        }

        .blob {
          position: absolute;
          border-radius: 50%;
          filter: blur(50px);
          animation: blobPulse 10s ease-in-out infinite;
          pointer-events: none;
        }
        .blob1 {
          top: -60px; left: -80px;
          width: 380px; height: 380px;
          background: radial-gradient(circle, rgba(109, 63, 107, 0.2), transparent);
          animation-delay: 0s;
        }
        .blob2 {
          bottom: -80px; right: -60px;
          width: 420px; height: 420px;
          background: radial-gradient(circle, rgba(71, 20, 63, 0.18), transparent);
          animation-delay: 3s;
        }

        @media (max-width: 768px) {
          .welcome-container {
            padding: 32px 24px !important;
          }
          .hero-image {
            max-width: 200px !important;
          }
          .welcome-title {
            font-size: 30px;
          }
          .auth-welcome-typewriter {
            font-size: 30px;
          }
        }

        @media (max-width: 480px) {
          .welcome-container {
            padding: 24px 18px !important;
            border-radius: 18px !important;
          }
          .hero-image {
            max-width: 168px !important;
          }
          .welcome-title {
            font-size: 26px;
          }
          .auth-welcome-typewriter {
            font-size: 26px;
          }
          .welcome-subtitle {
            font-size: 14px;
          }
          .login-button {
            width: 100%;
            min-height: 50px;
          }
        }
      `}</style>

      <div className="blob blob1" />
      <div className="blob blob2" />

      <div
        style={{
          position: "absolute",
          inset: 0,
          backgroundImage: "radial-gradient(circle, rgba(109, 63, 107, 0.06) 1px, transparent 1px)",
          backgroundSize: "32px 32px",
          pointerEvents: "none"
        }}
      />

      <div
        className="welcome-container"
        style={{
          background: "rgba(255, 255, 255, 0.9)",
          backdropFilter: "blur(16px)",
          WebkitBackdropFilter: "blur(16px)",
          borderRadius: 24,
          padding: "48px 40px",
          maxWidth: 480,
          width: "100%",
          textAlign: "center",
          boxShadow: "0 20px 60px rgba(109, 63, 107, 0.12), 0 4px 12px rgba(0, 0, 0, 0.06), inset 0 1px 0 rgba(255,255,255,0.9)",
          border: "1px solid rgba(221, 231, 227, 0.6)",
          position: "relative",
          zIndex: 1
        }}
      >
        <div
          style={{
            position: "absolute",
            top: 0,
            left: "20%",
            right: "20%",
            height: 3,
            background: "linear-gradient(90deg, transparent, #6D3F6B, transparent)",
            borderRadius: "0 0 6px 6px"
          }}
        />

        <div style={{ marginBottom: 36 }}>
          <img
            src={sheStoreLogoSrc}
            alt="She-Store Logo"
            className="hero-image"
            style={{
              maxWidth: 280,
              width: "100%",
              height: "auto",
              margin: "0 auto",
              display: "block"
            }}
          />
        </div>

        <div style={{ marginBottom: 32 }}>
          <TypewriterEffect
            className="auth-welcome-typewriter"
            cursorClassName="auth-welcome-cursor"
            words={[
              { text: "Welcome", className: "auth-welcome-word" },
              { text: "to", className: "auth-welcome-word" },
              { text: "your", className: "auth-welcome-word" },
              { text: "Order", className: "auth-welcome-word--accent" },
              { text: "Hub", className: "auth-welcome-word--accent" }
            ]}
          />
          <p className="welcome-subtitle">
            Streamline Shein order management with easy tracking and organization
          </p>
        </div>

        <div
          style={{
            display: "flex",
            justifyContent: "center",
            gap: 16,
            marginBottom: 36,
            flexWrap: "wrap"
          }}
        >
          {[
            { icon: "📦", label: "Easy Management" },
            { icon: "✓", label: "Quick Tracking" }
          ].map((item, i) => (
            <div
              key={i}
              className="feature-card"
              style={{
                background: "rgba(109, 63, 107, 0.08)",
                borderRadius: 12,
                padding: "10px 18px",
                border: "1px solid rgba(109, 63, 107, 0.15)",
                display: "flex",
                alignItems: "center",
                gap: 8
              }}
            >
              <span style={{ fontSize: 20 }}>{item.icon}</span>
              <span style={{ fontSize: 13, color: "#334155", fontWeight: 600 }}>{item.label}</span>
            </div>
          ))}
        </div>

        <button
          className="login-button"
          type="button"
          onClick={onContinue}
        >
          Sign In to Continue
        </button>

        <p
          style={{
            marginTop: 28,
            fontSize: 12,
            color: "#94a3b8",
            fontWeight: 400
          }}
        >
          Shein Agent Portal
        </p>
      </div>
    </div>
  );
}
