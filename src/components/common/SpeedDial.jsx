import { useEffect, useMemo, useState } from "react";
import "./speed-dial.css";

function renderActionIcon(icon) {
  if (typeof icon === "string") {
    return <span className="speed-dial-icon-text">{icon}</span>;
  }
  return icon || <span className="speed-dial-icon-text">•</span>;
}

export default function SpeedDial({ actions = [], position = "bottom-right", size = "large" }) {
  const [isOpen, setIsOpen] = useState(false);

  const visibleActions = useMemo(
    () => actions.filter((action) => action && action.show !== false),
    [actions]
  );

  useEffect(() => {
    if (!isOpen) return undefined;
    const onKeyDown = (event) => {
      if (event.key === "Escape") setIsOpen(false);
    };
    document.addEventListener("keydown", onKeyDown);
    return () => document.removeEventListener("keydown", onKeyDown);
  }, [isOpen]);

  useEffect(() => {
    if (!visibleActions.length) setIsOpen(false);
  }, [visibleActions.length]);

  if (!visibleActions.length) return null;

  const toggleDial = () => {
    if ("vibrate" in navigator) navigator.vibrate(10);
    setIsOpen((prev) => !prev);
  };

  const onActionClick = (action) => {
    if (action.disabled) return;
    setIsOpen(false);
    action.onClick?.();
  };

  return (
    <>
      <div
        className={`speed-dial-backdrop ${isOpen ? "open" : ""}`}
        aria-hidden={!isOpen}
        onClick={() => setIsOpen(false)}
      />

      <div className={`speed-dial speed-dial-${position} speed-dial-${size} ${isOpen ? "open" : ""}`}>
        <div className="speed-dial-actions" role="menu" aria-hidden={!isOpen}>
          {visibleActions.map((action, index) => (
            <div
              key={action.id || index}
              className={`speed-dial-item ${isOpen ? "open" : ""}`}
              style={{ "--sd-delay": `${index * 50}ms` }}
            >
              <span className="speed-dial-label" aria-hidden={!isOpen}>
                {action.label}
              </span>
              <button
                type="button"
                className={`speed-dial-btn ${action.primary ? "primary" : ""}`}
                onClick={() => onActionClick(action)}
                disabled={!!action.disabled}
                aria-label={action.label}
                role="menuitem"
              >
                {renderActionIcon(action.icon)}
              </button>
            </div>
          ))}
        </div>

        <button
          type="button"
          className={`speed-dial-main ${isOpen ? "open" : ""}`}
          aria-label="فتح قائمة الإجراءات"
          aria-expanded={isOpen}
          onClick={toggleDial}
        >
          <span className="speed-dial-main-icon">+</span>
        </button>
      </div>
    </>
  );
}
