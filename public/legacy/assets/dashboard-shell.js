(function(){
  "use strict";

  function svgIcon(name){
    const common = "viewBox='0 0 24 24' fill='none' stroke='currentColor' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'";
    const icons = {
      menu: `<svg ${common}><line x1='4' y1='12' x2='20' y2='12'/><line x1='4' y1='6' x2='20' y2='6'/><line x1='4' y1='18' x2='20' y2='18'/></svg>`,
      close: `<svg ${common}><line x1='18' y1='6' x2='6' y2='18'/><line x1='6' y1='6' x2='18' y2='18'/></svg>`,
      search: `<svg ${common}><circle cx='11' cy='11' r='8'/><path d='m21 21-4.3-4.3'/></svg>`,
      package: `<svg ${common}><path d='M16.5 9.4 7.55 4.24'/><path d='M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z'/><path d='m3.3 7 8.7 5 8.7-5'/><path d='M12 22V12'/></svg>`,
      mappin: `<svg ${common}><path d='M12 22s7-5.7 7-12a7 7 0 1 0-14 0c0 6.3 7 12 7 12z'/><circle cx='12' cy='10' r='2.5'/></svg>`,
      home: `<svg ${common}><path d='m3 10 9-7 9 7'/><path d='M9 22V12h6v10'/><path d='M3 10v12h18V10'/></svg>`,
      archive: `<svg ${common}><rect x='3' y='4' width='18' height='4' rx='1'/><path d='M5 8v11a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8'/><path d='M10 12h4'/></svg>`,
      dollar: `<svg ${common}><line x1='12' y1='2' x2='12' y2='22'/><path d='M17 5H9.5a3.5 3.5 0 0 0 0 7H14.5a3.5 3.5 0 0 1 0 7H7'/></svg>`,
      users: `<svg ${common}><path d='M16 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2'/><circle cx='8.5' cy='7' r='4'/><path d='M20 8v6'/><path d='M23 11h-6'/></svg>`,
      bag: `<svg ${common}><path d='M6 8h12l-1 12H7L6 8Z'/><path d='M9 8a3 3 0 0 1 6 0'/></svg>`,
      truck: `<svg ${common}><path d='M10 17h4V5H2v12h3'/><path d='M14 9h4l4 4v4h-3'/><circle cx='7' cy='17' r='2'/><circle cx='17' cy='17' r='2'/></svg>`,
      logout: `<svg ${common}><path d='M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4'/><path d='M16 17l5-5-5-5'/><path d='M21 12H9'/></svg>`
    };
    return icons[name] || icons.package;
  }

  const NAV_ICON_BY_KEY = {
    orders: "package",
    "pickup-dashboard": "mappin",
    pickuppoint: "home",
    archive: "archive",
    finance: "dollar",
    customers: "users",
    collections: "bag",
    homepickup: "truck"
  };

  function pageKey(){
    return String(document.body && document.body.dataset ? document.body.dataset.page : "").toLowerCase();
  }

  function navKeyFromHref(href){
    const value = String(href || "").toLowerCase();
    if(value.includes("tab=customers")) return "customers";
    if(value.includes("pickup-dashboard")) return "pickup-dashboard";
    if(value.includes("pickuppoint")) return "pickuppoint";
    if(value.includes("archive")) return "archive";
    if(value.includes("finance")) return "finance";
    if(value.includes("collections")) return "collections";
    if(value.includes("homepickup")) return "homepickup";
    return "orders";
  }

  function normalizePath(path){
    return String(path || "").replace(/\/+$/, "").toLowerCase();
  }

  function isCustomersTabActive(){
    const current = normalizePath(window.location.pathname || "");
    const params = new URLSearchParams(window.location.search || "");
    return current.endsWith("/index.html") && params.get("tab") === "customers";
  }

  function isNavLinkActive(link){
    const href = link.getAttribute("href");
    if(!href) return false;

    let url;
    try{
      url = new URL(href, window.location.href);
    }catch(_err){
      return false;
    }

    const key = navKeyFromHref(href);
    if(key === "customers") return isCustomersTabActive();

    return normalizePath(window.location.pathname) === normalizePath(url.pathname);
  }

  function setControlIcon(btn, icon){
    if(!btn || btn.dataset.uiIconReady === "1") return;
    btn.innerHTML = svgIcon(icon);
    btn.dataset.uiIconReady = "1";
  }

  function ensureCustomersShortcut(content){
    if(!content || pageKey() !== "index") return;

    const links = Array.from(content.querySelectorAll("a"));
    if(!links.length) return;

    const hasCustomers = links.some((link)=>{
      const href = String(link.getAttribute("href") || "").toLowerCase();
      return href.includes("tab=customers") || href.includes("customers");
    });
    if(hasCustomers) return;

    const orderLink = links.find((link)=>{
      const href = String(link.getAttribute("href") || "").toLowerCase();
      return href.includes("index.html") && !href.includes("tab=customers");
    });
    if(!orderLink) return;

    const customers = document.createElement("a");
    customers.href = "./index.html?tab=customers";
    customers.textContent = "العملاء";
    orderLink.insertAdjacentElement("afterend", customers);
  }

  function enhanceCommandBar(bar){
    if(!bar || bar.dataset.uiEnhanced === "1") return;

    bar.classList.add("ss-ui-command");

    const menuBtn = bar.querySelector("#openSidebarBtn, .menuBtn");
    const tabs = bar.querySelector("#navTabs");
    const modeSwitch = bar.querySelector("#modeSwitch");
    const searchRow = bar.querySelector(".ordersSearchRow, .commandSearchRow") ||
      (document.getElementById("searchInput") ? document.getElementById("searchInput").closest(".row") : null);
    const commandLabel = bar.querySelector(".commandLabel");

    if(commandLabel) commandLabel.style.display = "none";

    if(menuBtn){
      menuBtn.classList.add("ss-ui-menu-btn");
      setControlIcon(menuBtn, "menu");
    }

    if(tabs) tabs.classList.add("ss-ui-tabs");
    if(modeSwitch) modeSwitch.classList.add("ss-ui-mode");

    if(searchRow){
      searchRow.classList.add("ss-ui-search-row");

      const input = searchRow.querySelector("#searchInput, .searchBox, input");
      const count = document.getElementById("searchCount");

      if(input && !searchRow.querySelector(".ss-ui-search-shell")){
        const shell = document.createElement("div");
        shell.className = "ss-ui-search-shell";

        const wrap = document.createElement("div");
        wrap.className = "ss-ui-search-wrap";

        const icon = document.createElement("span");
        icon.className = "ss-ui-search-icon";
        icon.innerHTML = svgIcon("search");

        wrap.appendChild(icon);
        wrap.appendChild(input);
        shell.appendChild(wrap);

        searchRow.appendChild(shell);
      }

      if(count && count.parentElement !== searchRow){
        searchRow.appendChild(count);
      }
    }

    const row = document.createElement("div");
    row.className = "ss-ui-command-row";

    const right = document.createElement("div");
    right.className = "ss-ui-command-right";

    const left = document.createElement("div");
    left.className = "ss-ui-command-left";

    if(tabs) right.appendChild(tabs);
    if(searchRow) right.appendChild(searchRow);
    if(modeSwitch) left.appendChild(modeSwitch);
    if(menuBtn) left.appendChild(menuBtn);

    bar.innerHTML = "";
    row.appendChild(right);
    row.appendChild(left);
    bar.appendChild(row);

    bar.dataset.uiEnhanced = "1";
  }

  function enhanceSimpleTopbar(topbar){
    if(!topbar || topbar.dataset.uiEnhanced === "1") return;

    topbar.classList.add("ss-ui-simple");

    const brand = topbar.querySelector(".brand");
    const menuBtn = topbar.querySelector("#openSidebarBtn, .menuBtn");

    if(menuBtn){
      menuBtn.classList.add("ss-ui-menu-btn");
      setControlIcon(menuBtn, "menu");
    }

    let titleText = document.title || "She-Store";
    let subText = "";

    if(brand){
      const b = brand.querySelector("b, strong, h1, h2");
      const muted = brand.querySelector(".muted");
      if(b && String(b.textContent || "").trim()) titleText = String(b.textContent || "").trim();
      if(muted && String(muted.textContent || "").trim()) subText = String(muted.textContent || "").trim();
    }

    const row = document.createElement("div");
    row.className = "ss-ui-simple-row";

    const titleWrap = document.createElement("div");
    titleWrap.className = "ss-ui-simple-title";

    const dot = document.createElement("span");
    dot.className = "ss-ui-simple-dot";

    const textWrap = document.createElement("div");

    const h1 = document.createElement("h1");
    h1.className = "ss-ui-simple-main";
    h1.textContent = titleText;

    textWrap.appendChild(h1);

    if(subText){
      const p = document.createElement("p");
      p.className = "ss-ui-simple-sub";
      p.textContent = subText;
      textWrap.appendChild(p);
    }

    titleWrap.appendChild(dot);
    titleWrap.appendChild(textWrap);

    topbar.innerHTML = "";
    row.appendChild(titleWrap);
    if(menuBtn) row.appendChild(menuBtn);
    topbar.appendChild(row);

    topbar.dataset.uiEnhanced = "1";
  }

  function enhanceSidebarHeader(sidebar){
    if(!sidebar) return;

    sidebar.classList.add("ss-ui-sidebar");

    const header = sidebar.querySelector(".sidebarHeader");
    if(!header) return;

    const closeBtn = header.querySelector("#closeSidebarBtn, .menuBtn");
    if(closeBtn){
      closeBtn.classList.add("ss-ui-close-btn");
      setControlIcon(closeBtn, "close");
    }

    if(header.querySelector(".ss-ui-brand")) return;

    let titleText = "She-Store";
    const oldTitle = header.querySelector("b");
    if(oldTitle){
      const raw = String(oldTitle.textContent || "").trim();
      if(raw && raw !== "القائمة") titleText = raw;
      oldTitle.remove();
    }

    const brand = document.createElement("div");
    brand.className = "ss-ui-brand";

    const badge = document.createElement("span");
    badge.className = "ss-ui-brand-badge";
    badge.textContent = "SS";

    const textWrap = document.createElement("div");
    const h2 = document.createElement("h2");
    h2.className = "ss-ui-brand-title";
    h2.textContent = titleText;

    const sub = document.createElement("p");
    sub.className = "ss-ui-brand-sub";
    sub.textContent = "Dashboard";

    textWrap.appendChild(h2);
    textWrap.appendChild(sub);

    brand.appendChild(badge);
    brand.appendChild(textWrap);

    header.insertBefore(brand, closeBtn || null);
  }

  function decorateSidebarContent(){
    const content = document.getElementById("sidebarContent");
    if(!content) return;

    ensureCustomersShortcut(content);

    const links = Array.from(content.querySelectorAll("a"));
    links.forEach((link)=>{
      link.classList.add("ss-ui-nav-link");

      if(!link.dataset.uiLabel){
        link.dataset.uiLabel = String(link.textContent || "").trim();
      }

      const labelText = link.dataset.uiLabel || "";
      const key = navKeyFromHref(link.getAttribute("href"));
      const iconName = NAV_ICON_BY_KEY[key] || "package";

      link.textContent = "";

      const icon = document.createElement("span");
      icon.className = "ss-ui-nav-icon";
      icon.innerHTML = svgIcon(iconName);

      const label = document.createElement("span");
      label.className = "ss-ui-nav-label";
      label.textContent = labelText;

      link.appendChild(icon);
      link.appendChild(label);
      link.classList.toggle("is-active", isNavLinkActive(link));
    });

    const logout = content.querySelector("button.danger");
    if(logout){
      logout.classList.add("ss-ui-logout-btn");
      if(!logout.dataset.uiLabel){
        logout.dataset.uiLabel = String(logout.textContent || "").trim();
      }

      const labelText = logout.dataset.uiLabel || "";
      logout.textContent = "";

      const icon = document.createElement("span");
      icon.className = "ss-ui-nav-icon";
      icon.innerHTML = svgIcon("logout");

      const label = document.createElement("span");
      label.className = "ss-ui-nav-label";
      label.textContent = labelText;

      logout.appendChild(icon);
      logout.appendChild(label);
    }
  }

  function enhanceSidebar(){
    const overlay = document.getElementById("sidebarOverlay");
    if(overlay) overlay.classList.add("ss-ui-overlay");

    const sidebar = document.getElementById("sidebar");
    if(!sidebar) return;

    const syncSidebarOpenState = ()=>{
      const isOpen = sidebar.classList.contains("open") || (overlay && overlay.classList.contains("open"));
      document.body.classList.toggle("ss-ui-sidebar-open", !!isOpen);
    };

    enhanceSidebarHeader(sidebar);
    decorateSidebarContent();
    syncSidebarOpenState();

    if(sidebar.dataset.uiOpenObserved !== "1"){
      const openObserver = new MutationObserver(syncSidebarOpenState);
      openObserver.observe(sidebar, { attributes:true, attributeFilter:["class"] });
      if(overlay) openObserver.observe(overlay, { attributes:true, attributeFilter:["class"] });
      window.addEventListener("pageshow", syncSidebarOpenState);
      window.addEventListener("resize", syncSidebarOpenState);
      sidebar.dataset.uiOpenObserved = "1";
    }

    const content = document.getElementById("sidebarContent");
    if(content && content.dataset.uiObserved !== "1"){
      const observer = new MutationObserver(()=>{
        decorateSidebarContent();
      });
      observer.observe(content, { childList:true, subtree:false });
      content.dataset.uiObserved = "1";
    }
  }

  function boot(){
    const commandBar = document.querySelector(".commandBar");
    const topbar = document.querySelector(".topbar");

    if(commandBar) enhanceCommandBar(commandBar);
    if(topbar) enhanceSimpleTopbar(topbar);

    enhanceSidebar();

    window.addEventListener("popstate", decorateSidebarContent);
  }

  if(document.readyState === "loading"){
    document.addEventListener("DOMContentLoaded", boot);
  }else{
    boot();
  }
})();
