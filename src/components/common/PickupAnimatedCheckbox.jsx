import { useId } from "react";

export default function PickupAnimatedCheckbox({
  checked,
  onChange,
  disabled = false,
  ariaLabel = "?? ????????"
}) {
  const reactId = useId();
  const safeToken = String(reactId).replace(/[^a-zA-Z0-9_-]/g, "");
  const inputId = `pickup-cbx-${safeToken}`;
  const filterId = `pickup-goo-${safeToken}`;

  return (
    <div className="pickup-checkbox-12">
      <div className="pickup-cbx" style={{ filter: `url(#${filterId})`, WebkitFilter: `url(#${filterId})` }}>
        <input
          type="checkbox"
          id={inputId}
          checked={!!checked}
          onChange={onChange}
          disabled={disabled}
          aria-label={ariaLabel}
        />
        <label htmlFor={inputId} />
        <svg fill="none" viewBox="0 0 15 14" height={14} width={15} aria-hidden="true">
          <path d="M2 8.36364L6.23077 12L13 2" />
        </svg>
      </div>

      <svg className="pickup-checkbox-defs" version="1.1" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
        <defs>
          <filter id={filterId}>
            <feGaussianBlur result="blur" stdDeviation={4} in="SourceGraphic" />
            <feColorMatrix
              result={filterId}
              values="1 0 0 0 0  0 1 0 0 0  0 0 1 0 0  0 0 0 22 -7"
              mode="matrix"
              in="blur"
            />
            <feBlend in2={filterId} in="SourceGraphic" />
          </filter>
        </defs>
      </svg>
    </div>
  );
}
