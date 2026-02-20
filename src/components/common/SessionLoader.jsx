import "./session-loader.css";

export default function SessionLoader({ label = "" }) {
  return (
    <div className="session-loader" role="status" aria-live="polite">
      <div className="session-loader-wrapper" aria-hidden="true">
        <div className="session-loader-circle" />
        <div className="session-loader-circle" />
        <div className="session-loader-circle" />
        <div className="session-loader-shadow" />
        <div className="session-loader-shadow" />
        <div className="session-loader-shadow" />
      </div>
      {label ? <div className="session-loader-label">{label}</div> : null}
    </div>
  );
}
