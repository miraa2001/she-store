import ordersMenuIcon from "../../assets/icons/navigation/orders.png";

export default function AppNavIcon({ name, className = "icon" }) {
  const common = {
    fill: "none",
    stroke: "currentColor",
    strokeWidth: "2",
    strokeLinecap: "round",
    strokeLinejoin: "round",
    viewBox: "0 0 24 24"
  };

  if (name === "package") {
    return <img src={ordersMenuIcon} alt="" aria-hidden="true" className={className} />;
  }

  if (name === "map") {
    return (
      <svg className={className} {...common}>
        <path d="M12 22s7-5.7 7-12a7 7 0 1 0-14 0c0 6.3 7 12 7 12z" />
        <circle cx="12" cy="10" r="2.5" />
      </svg>
    );
  }

  if (name === "home") {
    return (
      <svg className={className} {...common}>
        <path d="m3 10 9-7 9 7" />
        <path d="M9 22V12h6v10" />
        <path d="M3 10v12h18V10" />
      </svg>
    );
  }

  if (name === "archive") {
    return (
      <svg className={className} {...common}>
        <rect x="3" y="4" width="18" height="4" rx="1" />
        <path d="M5 8v11a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8" />
        <path d="M10 12h4" />
      </svg>
    );
  }

  if (name === "dollar") {
    return (
      <svg className={className} {...common}>
        <line x1="12" y1="2" x2="12" y2="22" />
        <path d="M17 5H9.5a3.5 3.5 0 0 0 0 7H14.5a3.5 3.5 0 0 1 0 7H7" />
      </svg>
    );
  }

  if (name === "users") {
    return (
      <svg className={className} {...common}>
        <path d="M16 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" />
        <circle cx="8.5" cy="7" r="4" />
        <path d="M20 8v6" />
        <path d="M23 11h-6" />
      </svg>
    );
  }

  if (name === "bag") {
    return (
      <svg className={className} {...common}>
        <path d="M6 8h12l-1 12H7L6 8Z" />
        <path d="M9 8a3 3 0 0 1 6 0" />
      </svg>
    );
  }

  if (name === "truck") {
    return (
      <svg className={className} {...common}>
        <path d="M10 17h4V5H2v12h3" />
        <path d="M14 9h4l4 4v4h-3" />
        <circle cx="7" cy="17" r="2" />
        <circle cx="17" cy="17" r="2" />
      </svg>
    );
  }

  if (name === "calendar") {
    return (
      <svg className={className} {...common}>
        <path d="M8 2v4" />
        <path d="M16 2v4" />
        <rect x="3" y="4" width="18" height="18" rx="2" />
        <path d="M3 10h18" />
      </svg>
    );
  }

  return null;
}
