import { Navigate, useNavigate } from "react-router-dom";
import SessionLoader from "../components/common/SessionLoader";
import { useAuthProfile } from "../hooks/useAuthProfile";
import AuthWelcomePage from "./AuthWelcomePage";

export default function WelcomeEntryPage() {
  const { profile } = useAuthProfile();
  const navigate = useNavigate();

  if (profile.loading) {
    return (
      <div
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

  if (profile.authenticated) {
    return <Navigate to="/orders" replace />;
  }

  return <AuthWelcomePage onContinue={() => navigate("/login")} />;
}
