import { useId, useMemo, useState } from "react";
import styled from "styled-components";

export default function MinimalInput({
  type = "text",
  placeholder,
  value,
  onChange,
  name,
  disabled = false,
  autoComplete,
  dir = "rtl"
}) {
  const reactId = useId();
  const inputId = useMemo(() => `minimal-input-${name || reactId.replace(/[:]/g, "")}`, [name, reactId]);
  const [isFocused, setIsFocused] = useState(false);
  const hasValue = String(value || "").length > 0;
  const isActive = isFocused || hasValue;

  return (
    <StyledWrapper
      className={`minimal-input-wrapper ${isActive ? "is-active" : ""}`}
      $isRtl={dir === "rtl"}
    >
      <input
        id={inputId}
        type={type}
        value={value}
        onChange={onChange}
        onFocus={() => setIsFocused(true)}
        onBlur={() => setIsFocused(false)}
        name={name}
        className="minimal-input"
        required
        disabled={disabled}
        autoComplete={autoComplete}
        dir={dir}
      />
      <label className={`minimal-label ${isActive ? "active" : ""}`} htmlFor={inputId}>
        {placeholder}
      </label>
      <span className="minimal-border" aria-hidden="true" />
    </StyledWrapper>
  );
}

const StyledWrapper = styled.div`
  position: relative;
  margin-bottom: 32px;

  .minimal-input {
    width: 100%;
    padding: 16px 0 8px;
    font-size: 18px;
    color: #000;
    border: none;
    border-bottom: 2px solid rgba(0, 0, 0, 0.1);
    background: transparent;
    outline: none;
    transition: border-color 0.3s ease;
  }

  .minimal-input:disabled {
    opacity: 0.6;
    cursor: not-allowed;
  }

  .minimal-label {
    position: absolute;
    top: 16px;
    ${(props) => (props.$isRtl ? "right: 0;" : "left: 0;")}
    font-size: 18px;
    color: rgba(0, 0, 0, 0.4);
    pointer-events: none;
    transition: top 0.3s ease, font-size 0.3s ease, color 0.3s ease, font-weight 0.3s ease;
  }

  .minimal-label.active {
    top: -8px;
    font-size: 12px;
    color: #47143F;
    font-weight: 600;
  }

  .minimal-border {
    position: absolute;
    bottom: 0;
    ${(props) => (props.$isRtl ? "right: 0;" : "left: 0;")}
    width: 0;
    height: 2px;
    background: #47143F;
    transition: width 0.3s ease;
  }

  &.is-active .minimal-input {
    border-bottom-color: #47143F;
  }

  &.is-active .minimal-border {
    width: 100%;
  }
`;
