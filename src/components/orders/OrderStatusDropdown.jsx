import { useEffect, useMemo, useRef, useState } from "react";
import styled from "styled-components";
import { ChevronDown, Clock, Truck, MapPin, CheckCircle } from "lucide-react";

const STATUSES = [
  {
    id: "pending",
    label: "قيد الانتظار",
    description: "في انتظار وصول الطلب",
    icon: Clock,
    color: "#b45309",
    bgColor: "rgba(251, 191, 36, 0.10)",
    borderColor: "rgba(251, 191, 36, 0.34)"
  },
  {
    id: "arrived",
    label: "تم وصول الطلب",
    description: "يتم معالجة المشتريات",
    icon: Truck,
    color: "#2563eb",
    bgColor: "rgba(37, 99, 235, 0.10)",
    borderColor: "rgba(37, 99, 235, 0.32)"
  },
  {
    id: "at_pickup",
    label: "في نقطة الاستلام",
    description: "المشتريات قيد الاستلام",
    icon: MapPin,
    color: "#592955",
    bgColor: "rgba(109, 63, 107, 0.12)",
    borderColor: "rgba(109, 63, 107, 0.34)"
  },
  {
    id: "collected",
    label: "تم التحصيل",
    description: "تم معالجة الطلب كاملًا",
    icon: CheckCircle,
    color: "#47143F",
    bgColor: "rgba(71, 20, 63, 0.11)",
    borderColor: "rgba(71, 20, 63, 0.30)"
  }
];

function getStatusById(id) {
  return STATUSES.find((status) => status.id === id) || STATUSES[0];
}

export default function OrderStatusDropdown({
  value = "pending",
  disabled = false,
  lockCollected = false,
  onChange
}) {
  const [isOpen, setIsOpen] = useState(false);
  const dropdownRef = useRef(null);
  const currentStatus = useMemo(() => getStatusById(value), [value]);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const pickStatus = (statusId) => {
    if (disabled) return;
    if (statusId === "collected" && !lockCollected) return;
    onChange?.(statusId);
    setIsOpen(false);
  };

  return (
    <StyledWrapper dir="rtl" ref={dropdownRef}>
      <button
        type="button"
        className={`trigger ${isOpen ? "open" : ""}`}
        onClick={() => !disabled && setIsOpen((prev) => !prev)}
        disabled={disabled}
        aria-haspopup="listbox"
        aria-expanded={isOpen}
      >
        <span className="trigger-value">{currentStatus.label}</span>
        <ChevronDown className={`chevron ${isOpen ? "rotated" : ""}`} />
      </button>

      <div className={`menu ${isOpen ? "open" : ""}`} role="listbox" aria-label="قائمة حالات الطلب">
        <div className="menu-items webkit-scrollbar">
          {STATUSES.map((status) => {
            const selected = currentStatus.id === status.id;
            const isCollectedLocked = status.id === "collected" && !lockCollected;
            const StatusIcon = status.icon;

            return (
              <button
                key={status.id}
                type="button"
                className={`menu-item ${selected ? "selected" : ""}`}
                onClick={() => pickStatus(status.id)}
                disabled={disabled || isCollectedLocked}
                style={{
                  "--status-color": status.color,
                  "--status-bg": status.bgColor,
                  "--status-border": status.borderColor
                }}
              >
                <div className="menu-item-icon-wrap">
                  <StatusIcon className="status-icon" />
                </div>

                <div className="menu-item-text">
                  <p className="menu-item-title">{status.label}</p>
                  <p className="menu-item-desc">{status.description}</p>
                </div>

                {isCollectedLocked ? <span className="lock-hint">🔒</span> : null}
              </button>
            );
          })}
        </div>
      </div>
    </StyledWrapper>
  );
}

const StyledWrapper = styled.div`
  position: relative;
  display: inline-block;
  width: fit-content;
  min-width: 170px;

  .trigger {
    min-height: 40px;
    height: 40px;
    width: fit-content;
    min-width: 170px;
    max-width: 280px;
    display: inline-flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    border-radius: 12px;
    border: 1px solid rgba(109, 63, 107, 0.35);
    background: #fff;
    color: #1f2937;
    padding: 0 12px;
    cursor: pointer;
    transition: all 220ms ease;
    font-family: inherit;
    box-shadow: 0 1px 2px rgba(15, 23, 42, 0.08);
  }

  .trigger:hover:not(:disabled),
  .trigger.open {
    border-color: #6D3F6B;
    box-shadow: 0 0 0 3px rgba(109, 63, 107, 0.14);
  }

  .trigger:disabled {
    opacity: 0.72;
    cursor: not-allowed;
  }

  .trigger-value {
    font-size: 13px;
    font-weight: 700;
    line-height: 1.2;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  .chevron {
    width: 16px;
    height: 16px;
    color: #47143F;
    transition: transform 260ms ease;
    flex-shrink: 0;
  }

  .chevron.rotated {
    transform: rotate(180deg);
  }

  .menu {
    position: absolute;
    right: 0;
    top: calc(100% + 8px);
    min-width: 240px;
    width: max-content;
    max-width: min(92vw, 420px);
    border-radius: 18px;
    border: 1px solid rgba(203, 213, 225, 0.8);
    background: rgba(255, 255, 255, 0.98);
    backdrop-filter: blur(8px);
    box-shadow: 0 20px 46px -14px rgba(15, 23, 42, 0.24);
    overflow: hidden;
    z-index: 60;
    opacity: 0;
    transform: translateY(-8px) scale(0.98);
    visibility: hidden;
    pointer-events: none;
    transition: all 260ms ease;
  }

  .menu.open {
    opacity: 1;
    transform: translateY(0) scale(1);
    visibility: visible;
    pointer-events: auto;
  }

  .menu-items {
    max-height: 280px;
    overflow-y: auto;
    display: grid;
    gap: 6px;
    padding: 8px;
  }

  .menu-item {
    width: 100%;
    border: 1px solid transparent;
    border-radius: 12px;
    background: #fff;
    display: flex;
    align-items: flex-start;
    gap: 10px;
    padding: 10px;
    text-align: right;
    cursor: pointer;
    transition: all 180ms ease;
    font-family: inherit;
  }

  .menu-item:hover:not(:disabled) {
    background: rgba(109, 63, 107, 0.05);
    border-color: rgba(109, 63, 107, 0.2);
  }

  .menu-item.selected {
    background: var(--status-bg);
    border-color: var(--status-border);
  }

  .menu-item:disabled {
    opacity: 0.58;
    cursor: not-allowed;
  }

  .menu-item-icon-wrap {
    width: 28px;
    height: 28px;
    border-radius: 10px;
    background: #fff;
    border: 1px solid var(--status-border);
    color: var(--status-color);
    display: inline-flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
  }

  .status-icon {
    width: 15px;
    height: 15px;
  }

  .menu-item-text {
    min-width: 0;
  }

  .menu-item-title {
    margin: 0 0 2px;
    font-size: 13px;
    font-weight: 800;
    color: #0f172a;
  }

  .menu-item-desc {
    margin: 0;
    font-size: 11px;
    line-height: 1.4;
    color: #64748b;
    font-weight: 600;
  }

  .lock-hint {
    font-size: 12px;
    line-height: 1;
    align-self: center;
  }

  .webkit-scrollbar::-webkit-scrollbar {
    width: 8px;
    height: 8px;
  }

  .webkit-scrollbar::-webkit-scrollbar-track {
    background: transparent;
  }

  .webkit-scrollbar::-webkit-scrollbar-thumb {
    background: rgba(109, 63, 107, 0.25);
    border-radius: 999px;
  }

  .webkit-scrollbar:hover::-webkit-scrollbar-thumb {
    background: rgba(109, 63, 107, 0.45);
  }
`;
