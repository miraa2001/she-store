import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useLocation, useNavigate } from "react-router-dom";
import "./orders-page.css";
import {
  ORDER_TYPES,
  ORDER_STATUS,
  ORDER_STATUS_LABELS,
  SHEIN_ORDER_TYPES,
  createOrder,
  deriveOrderStatus,
  deleteOrderById,
  fetchOrdersWithSummary,
  formatILS,
  groupOrdersByMonth,
  isOrderFullyCollected,
  parsePrice,
  updateOrderName,
  updateOrderWorkflowStatus
} from "../lib/orders";
import { useAuthProfile } from "../hooks/useAuthProfile";
import { sb } from "../lib/supabaseClient";
import {
  createPurchaseWithRelations,
  deletePurchaseById,
  fetchPurchasesByOrder,
  markPurchasePaidPrice,
  restoreDeletedPurchase,
  sanitizeLinks,
  searchPurchasesByCustomerName,
  updatePurchaseWithRelations
} from "../lib/purchases";
import { searchByName } from "../lib/search";
import {
  CUSTOMER_CITIES,
  CUSTOMER_PICKUP_OPTIONS,
  createCustomer,
  customerFriendlyError,
  deleteCustomer,
  fetchCustomers,
  isValidCustomerPhone,
  normalizePhone,
  updateCustomer
} from "../lib/customers";
import {
  buildArrivalNotifyMessage,
  buildPickupInquiryMessage,
  buildWhatsappUrl,
  resolvePurchaseWhatsappTarget
} from "../lib/whatsapp";
import { exportOrderPdf } from "../lib/pdfExport";
import { hasGeminiKey, resolveTotalFromGemini, runGeminiCartAnalysis } from "../lib/gemini";
import { getOrdersNavItems, isNavHrefActive } from "../lib/navigation";
import { signOutAndRedirect } from "../lib/session";
import CustomersTab from "../components/tabs/CustomersTab";
import CommandHeader from "../components/orders/CommandHeader";
import OrdersBottomSheet from "../components/orders/OrdersBottomSheet";
import OrdersDrawer from "../components/orders/OrdersDrawer";
import OrdersTab from "../components/orders/OrdersTab";
import KanbanView from "../components/orders/KanbanView";
import PurchaseFormModal from "../components/orders/PurchaseFormModal";
import CustomerQuickAddModal from "../components/orders/CustomerQuickAddModal";
import LightboxModal from "../components/orders/LightboxModal";
import SessionLoader from "../components/common/SessionLoader";
import SpeedDial from "../components/common/SpeedDial";
import SheStoreLogo from "../components/common/SheStoreLogo";
import AppNavIcon from "../components/common/AppNavIcon";
import ordersMenuIcon from "../assets/icons/navigation/orders.png";

const BAG_OPTIONS = ["كيس كبير", "كيس صغير"];
const MAX_IMAGES = 10;
const ORDER_TYPE_OPTIONS = [
  { value: ORDER_TYPES.IHERB, label: "iHerb" },
  { value: ORDER_TYPES.SHEIN, label: "Shein" }
];
const SHEIN_TYPE_OPTIONS = [
  { value: SHEIN_ORDER_TYPES.NORMAL, label: "Normal" },
  { value: SHEIN_ORDER_TYPES.MAKEUP, label: "Makeup" },
  { value: SHEIN_ORDER_TYPES.ELECTRONICS, label: "Electronics" }
];

function createEmptyForm(orderId, customers) {
  return {
    purchaseId: "",
    orderId: orderId || "",
    customerId: "",
    customerName: "",
    qty: "1",
    price: "",
    paidPrice: "",
    bagSize: "كيس صغير",
    pickupPoint: "",
    note: "",
    links: [""],
    newFiles: [],
    existingImages: [],
    removeImageIds: []
  };
}


function Icon({ name, className = "" }) {
  const common = {
    fill: "none",
    stroke: "currentColor",
    strokeWidth: "2",
    strokeLinecap: "round",
    strokeLinejoin: "round",
    viewBox: "0 0 24 24"
  };

  if (name === "menu") {
    return (
      <svg className={className} {...common}>
        <line x1="4" y1="12" x2="20" y2="12" />
        <line x1="4" y1="6" x2="20" y2="6" />
        <line x1="4" y1="18" x2="20" y2="18" />
      </svg>
    );
  }

  if (name === "close") {
    return (
      <svg className={className} {...common}>
        <line x1="18" y1="6" x2="6" y2="18" />
        <line x1="6" y1="6" x2="18" y2="18" />
      </svg>
    );
  }

  if (name === "search") {
    return (
      <svg className={className} {...common}>
        <circle cx="11" cy="11" r="8" />
        <path d="m21 21-4.3-4.3" />
      </svg>
    );
  }

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

  if (name === "logout") {
    return (
      <svg className={className} {...common}>
        <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" />
        <path d="M16 17l5-5-5-5" />
        <path d="M21 12H9" />
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

  if (name === "chevron-left") {
    return (
      <svg className={className} {...common}>
        <path d="m15 18-6-6 6-6" />
      </svg>
    );
  }

  if (name === "chevron-right") {
    return (
      <svg className={className} {...common}>
        <path d="m9 18 6-6-6-6" />
      </svg>
    );
  }

  return null;
}

function statusLabel(status) {
  return ORDER_STATUS_LABELS[status] || ORDER_STATUS_LABELS[ORDER_STATUS.PENDING];
}

export default function OrdersPage() {
  const [viewportWidth, setViewportWidth] = useState(() =>
    typeof window === "undefined" ? 1280 : window.innerWidth
  );
  const [globalOpen, setGlobalOpen] = useState(false);
  const [ordersMenuOpen, setOrdersMenuOpen] = useState(false);
  const [activeTab, setActiveTab] = useState("orders");
  const [search, setSearch] = useState("");
  const [editMode, setEditMode] = useState(true);
  const [desktopOrdersView, setDesktopOrdersView] = useState("list");
  const [selectedOrderId, setSelectedOrderId] = useState("");
  const { profile } = useAuthProfile();

  const [orders, setOrders] = useState([]);
  const [ordersLoading, setOrdersLoading] = useState(true);
  const [ordersError, setOrdersError] = useState("");
  const [headerSearchResults, setHeaderSearchResults] = useState([]);
  const [headerSearchLoading, setHeaderSearchLoading] = useState(false);
  const [headerSearchTouched, setHeaderSearchTouched] = useState(false);

  const [purchases, setPurchases] = useState([]);
  const [purchasesLoading, setPurchasesLoading] = useState(false);
  const [purchasesError, setPurchasesError] = useState("");

  const [customers, setCustomers] = useState([]);
  const [customersLoading, setCustomersLoading] = useState(true);
  const [customersError, setCustomersError] = useState("");
  const [customerSearch, setCustomerSearch] = useState("");
  const [customerForm, setCustomerForm] = useState({
    name: "",
    phone: "",
    city: CUSTOMER_CITIES[0],
    pickup: CUSTOMER_PICKUP_OPTIONS[2]
  });
  const [customerFormMessage, setCustomerFormMessage] = useState("");
  const [customerFormSaving, setCustomerFormSaving] = useState(false);
  const [editingCustomerId, setEditingCustomerId] = useState("");
  const [editingCustomerForm, setEditingCustomerForm] = useState(null);

  const [menuPurchaseId, setMenuPurchaseId] = useState("");

  const [formOpen, setFormOpen] = useState(false);
  const [formMode, setFormMode] = useState("add");
  const [formState, setFormState] = useState(() => createEmptyForm("", []));
  const [formSaving, setFormSaving] = useState(false);
  const [formError, setFormError] = useState("");
  const [formUploadProgress, setFormUploadProgress] = useState("");
  const [formAiStatus, setFormAiStatus] = useState({ text: "", isError: false });
  const [formAiRunning, setFormAiRunning] = useState(false);
  const [quickCustomerOpen, setQuickCustomerOpen] = useState(false);
  const [quickCustomerSaving, setQuickCustomerSaving] = useState(false);
  const [quickCustomerStatus, setQuickCustomerStatus] = useState({ text: "", isError: false });
  const [quickCustomerForm, setQuickCustomerForm] = useState({
    name: "",
    phone: "",
    city: CUSTOMER_CITIES[0],
    pickup: CUSTOMER_PICKUP_OPTIONS[2]
  });
  const [pdfExporting, setPdfExporting] = useState(false);
  const [orderStatusSaving, setOrderStatusSaving] = useState(false);
  const [kanbanMovingPurchaseId, setKanbanMovingPurchaseId] = useState("");
  const [newFilePreviews, setNewFilePreviews] = useState([]);

  const [toast, setToast] = useState(null);
  const [deleteSnapshot, setDeleteSnapshot] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [confirmDeleteBusy, setConfirmDeleteBusy] = useState(false);
  const [paidPriceDialog, setPaidPriceDialog] = useState(null);
  const [paidPriceDialogBusy, setPaidPriceDialogBusy] = useState(false);
  const [orderDialog, setOrderDialog] = useState(null);
  const [orderDialogBusy, setOrderDialogBusy] = useState(false);
  const [lightbox, setLightbox] = useState({ open: false, images: [], index: 0, title: "" });
  const [highlightPurchaseId, setHighlightPurchaseId] = useState("");
  const hasInitializedUrlState = useRef(false);

  const location = useLocation();
  const navigate = useNavigate();

  const isMobile = viewportWidth < 768;
  const isTablet = viewportWidth >= 768 && viewportWidth < 1024;
  const isDesktop = viewportWidth >= 1024;

  useEffect(() => {
    const onResize = () => setViewportWidth(window.innerWidth);
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);

  useEffect(() => {
    let cancelled = false;
    const needle = String(search || "").trim();

    if (!needle) {
      setHeaderSearchLoading(false);
      setHeaderSearchResults([]);
      return undefined;
    }

    setHeaderSearchTouched(true);
    const timer = window.setTimeout(async () => {
      try {
        setHeaderSearchLoading(true);
        const rows = await searchPurchasesByCustomerName(needle, 100);
        if (cancelled) return;
        setHeaderSearchResults(rows || []);
      } catch (error) {
        console.error(error);
        if (!cancelled) setHeaderSearchResults([]);
      } finally {
        if (!cancelled) setHeaderSearchLoading(false);
      }
    }, 220);

    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [search]);

  const groupedOrders = useMemo(() => groupOrdersByMonth(orders), [orders]);
  const searchCount = useMemo(() => headerSearchResults.length, [headerSearchResults]);

  const totalOrders = orders.length;
  const isRahaf = profile.role === "rahaf";
  const isReem = profile.role === "reem";
  const isViewOnlyRole = profile.role === "reem" || profile.role === "rawand";
  const canUseOrdersWorkbench = isRahaf || isViewOnlyRole;
  const allowedTabs = useMemo(
    () => (isReem ? ["orders"] : isRahaf || isViewOnlyRole ? ["orders", "customers"] : ["orders"]),
    [isRahaf, isReem, isViewOnlyRole]
  );

  const visibleNavItems = useMemo(() => getOrdersNavItems(profile.role), [profile.role]);
  const isSidebarItemActive = useCallback(
    (href) => {
      const customersTabActive = isNavHrefActive("#/orders?tab=customers", location);
      const isRootOrdersItem = href === "#/orders";
      if (isRootOrdersItem) {
        return isNavHrefActive(href, location) && !customersTabActive;
      }
      return isNavHrefActive(href, location);
    },
    [location]
  );

  const selectedOrder = useMemo(
    () => orders.find((order) => String(order.id) === String(selectedOrderId)) || null,
    [orders, selectedOrderId]
  );

  const selectedOrderIsFullyCollected = useMemo(
    () => isOrderFullyCollected(purchases),
    [purchases]
  );

  const selectedOrderStatus = useMemo(() => {
    if (!selectedOrder) return ORDER_STATUS.PENDING;
    return deriveOrderStatus({
      arrived: selectedOrder.arrived,
      placedAtPickup: selectedOrder.placedAtPickup,
      purchaseCount: purchases.length,
      allCollected: selectedOrderIsFullyCollected
    });
  }, [purchases.length, selectedOrder, selectedOrderIsFullyCollected]);

  const orderNameById = useMemo(() => {
    const map = new Map();
    orders.forEach((order) => map.set(String(order.id), order.name));
    return map;
  }, [orders]);

  const filteredPurchases = useMemo(() => purchases, [purchases]);

  const purchaseStats = useMemo(() => {
    const count = purchases.length;
    const totalQty = purchases.reduce((sum, item) => sum + Number(item.qty || 0), 0);
    const totalPrice = purchases.reduce((sum, item) => sum + parsePrice(item.price), 0);
    return { count, totalQty, totalPrice };
  }, [purchases]);

  const filteredCustomers = useMemo(() => {
    return searchByName(customers, customerSearch, (customer) => [
      customer.name,
      customer.phone,
      customer.city
    ]);
  }, [customerSearch, customers]);

  const refreshOrders = useCallback(async (preferredId = "") => {
    setOrdersLoading(true);
    setOrdersError("");

    try {
      const data = await fetchOrdersWithSummary();
      const allOrders = data || [];
      const visibleOrders = isReem ? allOrders.filter((order) => !!order.arrived) : allOrders;
      setOrders(visibleOrders);
      setSelectedOrderId((prev) => {
        const candidate = preferredId || prev;
        if (candidate && visibleOrders.some((order) => String(order.id) === String(candidate))) {
          return candidate;
        }
        return visibleOrders[0]?.id || "";
      });
    } catch (error) {
      console.error(error);
      setOrdersError("فشل تحميل الطلبات من قاعدة البيانات.");
    } finally {
      setOrdersLoading(false);
    }
  }, [isReem]);

  const refreshCustomers = useCallback(async () => {
    setCustomersLoading(true);
    setCustomersError("");

    try {
      const list = await fetchCustomers();
      setCustomers(list || []);
      setEditingCustomerId((prev) =>
        prev && (list || []).some((customer) => String(customer.id) === String(prev)) ? prev : ""
      );
      return list || [];
    } catch (error) {
      console.error(error);
      setCustomersError("فشل تحميل بيانات العملاء.");
      return [];
    } finally {
      setCustomersLoading(false);
    }
  }, []);

  const refreshPurchases = useCallback(async (orderId) => {
    if (!orderId) {
      setPurchases([]);
      setPurchasesLoading(false);
      return;
    }

    setPurchasesLoading(true);
    setPurchasesError("");

    try {
      const list = await fetchPurchasesByOrder(orderId);
      setPurchases(list || []);
    } catch (error) {
      console.error(error);
      setPurchasesError("فشل تحميل المشتريات.");
    } finally {
      setPurchasesLoading(false);
    }
  }, []);

  useEffect(() => {
    if (!profile.authenticated) {
      setOrdersLoading(false);
      setCustomersLoading(false);
      return;
    }

    if (!canUseOrdersWorkbench) {
      setOrders([]);
      setOrdersLoading(false);
      setCustomers([]);
      setCustomersLoading(false);
      setOrdersError("هذا الحساب لا يملك صلاحية صفحة الطلبات.");
      return;
    }

    refreshOrders();
    refreshCustomers();
  }, [canUseOrdersWorkbench, profile.authenticated, refreshCustomers, refreshOrders]);

  useEffect(() => {
    if (!profile.authenticated) return;
    if (!isViewOnlyRole) return;
    setEditMode(false);
  }, [isViewOnlyRole, profile.authenticated]);

  useEffect(() => {
    if (!profile.authenticated) return;
    if (!isRahaf) return;
    setEditMode(true);
  }, [isRahaf, profile.authenticated]);

  useEffect(() => {
    if (!isReem) return;
    setDesktopOrdersView("list");
  }, [isReem]);


  useEffect(() => {
    const params = new URLSearchParams(location.search);
    const tabFromUrl = params.get("tab");

    if (tabFromUrl && allowedTabs.includes(tabFromUrl)) {
      setActiveTab((prev) => (prev === tabFromUrl ? prev : tabFromUrl));
    }

    hasInitializedUrlState.current = true;
  }, [allowedTabs, location.search]);

  useEffect(() => {
    if (!profile.authenticated) return;
    if (!hasInitializedUrlState.current) return;

    const params = new URLSearchParams(location.search);
    let changed = false;

    if (params.get("tab") !== activeTab) {
      params.set("tab", activeTab);
      changed = true;
    }

    if (params.has("mode")) {
      params.delete("mode");
      changed = true;
    }

    if (!changed) return;
    const query = params.toString();
    navigate(query ? `${location.pathname}?${query}` : location.pathname, { replace: true });
  }, [activeTab, location.pathname, navigate, profile.authenticated]);

  useEffect(() => {
    if (allowedTabs.includes(activeTab)) return;
    setActiveTab("orders");
  }, [activeTab, allowedTabs]);

  useEffect(() => {
    if (!profile.authenticated) return;
    if (profile.role !== "laaura") return;
    window.location.hash = "#/pickuppoint";
  }, [profile.authenticated, profile.role]);

  useEffect(() => {
    if (activeTab !== "orders") return;
    refreshPurchases(selectedOrderId);
  }, [activeTab, refreshPurchases, selectedOrderId]);

  useEffect(() => {
    if (activeTab === "orders") return;
    setOrdersMenuOpen(false);
  }, [activeTab]);

  useEffect(() => {
    if (!toast) return;
    const timer = setTimeout(() => setToast(null), 5000);
    return () => clearTimeout(timer);
  }, [toast]);

  useEffect(() => {
    if (!highlightPurchaseId) return undefined;
    const timer = window.setTimeout(() => setHighlightPurchaseId(""), 2000);
    return () => window.clearTimeout(timer);
  }, [highlightPurchaseId]);

  useEffect(() => {
    const onDocPointerDown = (event) => {
      const target = event.target;
      if (!(target instanceof Element)) return;
      if (target.closest(".purchase-menu-wrap")) return;
      setMenuPurchaseId("");
    };

    document.addEventListener("pointerdown", onDocPointerDown);
    return () => {
      document.removeEventListener("pointerdown", onDocPointerDown);
    };
  }, []);

  useEffect(() => {
    const previews = formState.newFiles.map((file) => ({
      key: `${file.name}-${file.size}-${file.lastModified}`,
      url: URL.createObjectURL(file)
    }));

    setNewFilePreviews(previews);

    return () => {
      previews.forEach((item) => URL.revokeObjectURL(item.url));
    };
  }, [formState.newFiles]);

  const handleCreateCustomer = async (event) => {
    event.preventDefault();
    setCustomerFormMessage("");

    const name = String(customerForm.name || "").trim();
    const phone = normalizePhone(customerForm.phone);
    const city = String(customerForm.city || "").trim();
    const pickup = String(customerForm.pickup || "").trim();

    if (!name) {
      setCustomerFormMessage("الاسم مطلوب.");
      return;
    }
    if (!phone) {
      setCustomerFormMessage("رقم الهاتف مطلوب.");
      return;
    }
    if (!isValidCustomerPhone(phone)) {
      setCustomerFormMessage("رقم الهاتف يجب أن يبدأ بـ 970 أو 972.");
      return;
    }
    if (!city) {
      setCustomerFormMessage("اختاري المدينة.");
      return;
    }

    setCustomerFormSaving(true);
    setCustomerFormMessage("جاري الحفظ...");

    try {
      await createCustomer({
        name,
        phone,
        city,
        usual_pickup_point: pickup
      });

      setCustomerForm({
        name: "",
        phone: "",
        city: CUSTOMER_CITIES[0],
        pickup: CUSTOMER_PICKUP_OPTIONS[2]
      });
      setCustomerFormMessage("تم ✅");
      await refreshCustomers();
      setToast({ type: "success", text: "تم حفظ العميل." });
    } catch (error) {
      console.error(error);
      setCustomerFormMessage(customerFriendlyError(error, "فشل إضافة العميل."));
    } finally {
      setCustomerFormSaving(false);
    }
  };

  const openQuickCustomerModal = (prefillName = "") => {
    setQuickCustomerForm({
      name: String(prefillName || "").trim(),
      phone: "",
      city: CUSTOMER_CITIES[0],
      pickup: CUSTOMER_PICKUP_OPTIONS[2]
    });
    setQuickCustomerStatus({ text: "", isError: false });
    setQuickCustomerOpen(true);
  };

  const closeQuickCustomerModal = () => {
    if (quickCustomerSaving) return;
    setQuickCustomerOpen(false);
    setQuickCustomerStatus({ text: "", isError: false });
  };

  const handleQuickCustomerSubmit = async (event) => {
    event.preventDefault();
    if (quickCustomerSaving) return;

    const name = String(quickCustomerForm.name || "").trim();
    const phone = normalizePhone(quickCustomerForm.phone);
    const city = String(quickCustomerForm.city || "").trim();
    const pickup = String(quickCustomerForm.pickup || "").trim();

    if (!name) {
      setQuickCustomerStatus({ text: "الاسم مطلوب.", isError: true });
      return;
    }
    if (!phone) {
      setQuickCustomerStatus({ text: "رقم الهاتف مطلوب.", isError: true });
      return;
    }
    if (!isValidCustomerPhone(phone)) {
      setQuickCustomerStatus({ text: "رقم الهاتف يجب أن يبدأ بـ 970 أو 972.", isError: true });
      return;
    }
    if (!city) {
      setQuickCustomerStatus({ text: "اختاري المدينة.", isError: true });
      return;
    }

    setQuickCustomerSaving(true);
    setQuickCustomerStatus({ text: "جاري الحفظ...", isError: false });

    try {
      const createdCustomer = await createCustomer({
        name,
        phone,
        city,
        usual_pickup_point: pickup
      });

      await refreshCustomers();

      if (createdCustomer?.id) {
        setFormState((prev) => ({
          ...prev,
          customerId: createdCustomer.id,
          customerName: createdCustomer.name || prev.customerName,
          pickupPoint: createdCustomer.usual_pickup_point || prev.pickupPoint
        }));
      }

      setQuickCustomerOpen(false);
      setQuickCustomerStatus({ text: "", isError: false });
      setToast({ type: "success", text: "تم حفظ العميل." });
    } catch (error) {
      console.error(error);
      setQuickCustomerStatus({
        text: customerFriendlyError(error, "فشل إضافة العميل."),
        isError: true
      });
    } finally {
      setQuickCustomerSaving(false);
    }
  };

  const beginEditCustomer = (customer) => {
    setEditingCustomerId(customer.id);
    setEditingCustomerForm({
      name: customer.name || "",
      phone: customer.phone || "",
      city: customer.city || CUSTOMER_CITIES[0],
      pickup: customer.usual_pickup_point || CUSTOMER_PICKUP_OPTIONS[2]
    });
  };

  const cancelEditCustomer = () => {
    setEditingCustomerId("");
    setEditingCustomerForm(null);
  };

  const saveEditCustomer = async (customerId) => {
    if (!editingCustomerForm) return;

    const name = String(editingCustomerForm.name || "").trim();
    const phone = normalizePhone(editingCustomerForm.phone);
    const city = String(editingCustomerForm.city || "").trim();
    const pickup = String(editingCustomerForm.pickup || "").trim();

    if (!name) {
      setToast({ type: "warn", text: "الاسم مطلوب." });
      return;
    }
    if (!phone) {
      setToast({ type: "warn", text: "رقم الهاتف مطلوب." });
      return;
    }
    if (!isValidCustomerPhone(phone)) {
      setToast({ type: "warn", text: "رقم الهاتف يجب أن يبدأ بـ 970 أو 972." });
      return;
    }
    if (!city) {
      setToast({ type: "warn", text: "اختاري المدينة." });
      return;
    }

    try {
      await updateCustomer(customerId, {
        name,
        phone,
        city,
        usual_pickup_point: pickup
      });

      setToast({ type: "success", text: "تم تحديث العميل." });
      cancelEditCustomer();
      await refreshCustomers();
    } catch (error) {
      console.error(error);
      setToast({ type: "danger", text: customerFriendlyError(error, "فشل تحديث العميل.") });
    }
  };

  const handleDeleteCustomer = (customer) => {
    if (!customer) return;
    setConfirmDelete({ kind: "customer", item: customer });
  };

  const signOut = async () => {
    await signOutAndRedirect();
  };

  const clearFormAiStatus = useCallback(() => {
    setFormAiStatus({ text: "", isError: false });
  }, []);

  const openAddModal = () => {
    if (!selectedOrder) return;
    setFormMode("add");
    setFormState(createEmptyForm(selectedOrder.id, customers));
    setFormError("");
    setFormUploadProgress("");
    clearFormAiStatus();
    setFormOpen(true);
  };

  const openEditModal = (purchase) => {
    setFormMode("edit");
    setFormState({
      purchaseId: purchase.id,
      orderId: purchase.order_id,
      customerId: purchase.customer_id || "",
      customerName: purchase.customer_name || "",
      qty: purchase.qty ?? 1,
      price: purchase.price ?? "",
      paidPrice: purchase.paid_price ?? purchase.price ?? "",
      bagSize: purchase.bag_size || "كيس صغير",
      pickupPoint: purchase.pickup_point || CUSTOMER_PICKUP_OPTIONS[2],
      note: purchase.note || "",
      links: purchase.links?.length ? purchase.links : [""],
      newFiles: [],
      existingImages: purchase.images || [],
      removeImageIds: []
    });
    setFormError("");
    setFormUploadProgress("");
    clearFormAiStatus();
    setFormOpen(true);
    setMenuPurchaseId("");
  };

  const closeFormModal = () => {
    if (formSaving || formAiRunning) return;
    setFormOpen(false);
    setFormError("");
    setFormUploadProgress("");
    clearFormAiStatus();
  };

  const updateForm = (patch) => {
    setFormState((prev) => ({ ...prev, ...patch }));
  };

  const updateLinkValue = (index, value) => {
    setFormState((prev) => {
      const links = [...prev.links];
      links[index] = value;
      return { ...prev, links };
    });
  };

  const addLinkInput = () => {
    setFormState((prev) => ({ ...prev, links: [...prev.links, ""] }));
    clearFormAiStatus();
  };

  const removeLinkInput = () => {
    setFormState((prev) => {
      if (prev.links.length <= 1) return prev;
      return { ...prev, links: prev.links.slice(0, -1) };
    });
    clearFormAiStatus();
  };

  const addNewImages = (files) => {
    const picked = Array.from(files || []).filter(
      (file) => file && file.type && file.type.startsWith("image/")
    );
    if (!picked.length) return;

    setFormState((prev) => {
      const remainingExisting =
        prev.existingImages.filter((img) => !prev.removeImageIds.includes(img.id)).length;
      const allowed = Math.max(0, MAX_IMAGES - remainingExisting - prev.newFiles.length);
      const toAdd = picked.slice(0, allowed);

      if (!toAdd.length) {
        setFormError(`الحد الأقصى للصور هو ${MAX_IMAGES}.`);
        return prev;
      }

      return { ...prev, newFiles: [...prev.newFiles, ...toAdd] };
    });
    clearFormAiStatus();
  };

  const removeNewImage = (index) => {
    setFormState((prev) => ({
      ...prev,
      newFiles: prev.newFiles.filter((_, i) => i !== index)
    }));
    clearFormAiStatus();
  };

  const replaceNewImage = (index, nextFile) => {
    if (!nextFile) return;
    setFormState((prev) => {
      if (!Array.isArray(prev.newFiles) || index < 0 || index >= prev.newFiles.length) return prev;
      const nextFiles = [...prev.newFiles];
      nextFiles[index] = nextFile;
      return { ...prev, newFiles: nextFiles };
    });
    clearFormAiStatus();
  };

  const toggleExistingImageRemoval = (imageId) => {
    setFormState((prev) => {
      const ids = prev.removeImageIds.includes(imageId)
        ? prev.removeImageIds.filter((id) => id !== imageId)
        : [...prev.removeImageIds, imageId];
      return { ...prev, removeImageIds: ids };
    });
    clearFormAiStatus();
  };

  const onCustomerChange = (customerId) => {
    const customer = customers.find((item) => String(item.id) === String(customerId));
    setFormState((prev) => ({
      ...prev,
      customerId,
      customerName: customer?.name || "",
      pickupPoint: customer?.usual_pickup_point || prev.pickupPoint
    }));
  };

  const submitPurchaseForm = async (event) => {
    event?.preventDefault?.();

    if (!selectedOrder) {
      setFormError("اختاري طلبًا أولًا.");
      return false;
    }

    const qty = Number(formState.qty);
    const price = Number(formState.price);
    const paidPrice = Number(formState.paidPrice);

    if (!formState.customerId) {
      setFormError("اختاري الزبون.");
      return false;
    }
    if (String(formState.price).trim() === "") {
      setFormError("أدخلي السعر.");
      return false;
    }
    if (String(formState.paidPrice).trim() === "") {
      setFormError("أدخلي السعر المدفوع.");
      return false;
    }
    if (!Number.isInteger(qty) || qty < 1 || qty > 200) {
      setFormError("عدد القطع يجب أن يكون بين 1 و 200.");
      return false;
    }
    if (!Number.isFinite(price) || price < 0) {
      setFormError("السعر غير صالح.");
      return false;
    }
    if (!Number.isFinite(paidPrice) || paidPrice < 0) {
      setFormError("السعر المدفوع غير صالح.");
      return false;
    }
    if (!formState.bagSize) {
      setFormError("اختاري حجم الكيس.");
      return false;
    }
    if (!formState.pickupPoint) {
      setFormError("اختاري مكان الاستلام.");
      return false;
    }

    const normalizedLinks = sanitizeLinks(formState.links);
    setFormSaving(true);
    setFormError("");
    setFormUploadProgress("");

    try {
      if (formMode === "add") {
        const result = await createPurchaseWithRelations({
          purchase: {
            order_id: selectedOrder.id,
            customer_id: formState.customerId,
            customer_name: formState.customerName,
            qty,
            price,
            paid_price: paidPrice,
            bag_size: formState.bagSize,
            pickup_point: formState.pickupPoint,
            note: formState.note.trim()
          },
          links: normalizedLinks,
          files: formState.newFiles,
          onUploadProgress: (done, total) => {
            setFormUploadProgress(`جاري رفع الصور ${done}/${total}...`);
          }
        });

        if (result.uploadErrors.length) {
          setToast({ type: "warn", text: "تمت الإضافة مع أخطاء في بعض الصور." });
        } else {
          setToast({ type: "success", text: "تمت إضافة المشترى." });
        }
      } else {
        const removeImages = formState.existingImages.filter((img) =>
          formState.removeImageIds.includes(img.id)
        );

        await updatePurchaseWithRelations({
          purchaseId: formState.purchaseId,
          purchasePatch: {
            order_id: selectedOrder.id,
            customer_id: formState.customerId,
            customer_name: formState.customerName,
            qty,
            price,
            paid_price: paidPrice,
            bag_size: formState.bagSize,
            pickup_point: formState.pickupPoint,
            note: formState.note.trim()
          },
          links: normalizedLinks,
          removeImages,
          newFiles: formState.newFiles,
          onUploadProgress: (done, total) => {
            setFormUploadProgress(`جاري رفع الصور ${done}/${total}...`);
          }
        });

        setToast({ type: "success", text: "تم تحديث المشترى." });
      }

      await refreshPurchases(selectedOrder.id);
      await refreshOrders(selectedOrder.id);
      setFormOpen(false);
      setFormUploadProgress("");
      clearFormAiStatus();
      return true;
    } catch (error) {
      console.error(error);
      setFormError(error?.message || "تعذر حفظ البيانات.");
      return false;
    } finally {
      setFormSaving(false);
    }
  };

  const handleDeletePurchase = (purchase) => {
    if (!purchase) return;
    setMenuPurchaseId("");
    setConfirmDelete({ kind: "purchase", item: purchase });
  };

  const closeConfirmDelete = () => {
    if (confirmDeleteBusy) return;
    setConfirmDelete(null);
  };

  const confirmDeleteAction = async () => {
    if (!confirmDelete || !confirmDelete.item || confirmDeleteBusy) return;

    const target = confirmDelete.item;
    setConfirmDeleteBusy(true);

    try {
      if (confirmDelete.kind === "customer") {
        await deleteCustomer(target.id);
        setToast({ type: "success", text: "تم حذف العميل." });
        if (String(editingCustomerId) === String(target.id)) {
          cancelEditCustomer();
        }
        await refreshCustomers();
      } else if (confirmDelete.kind === "purchase") {
        setMenuPurchaseId("");
        await deletePurchaseById(target.id);

        setDeleteSnapshot({
          purchase: {
            id: target.id,
            order_id: target.order_id,
            customer_id: target.customer_id,
            customer_name: target.customer_name,
            qty: target.qty,
            price: target.price,
            paid_price: target.paid_price,
            bag_size: target.bag_size,
            pickup_point: target.pickup_point,
            note: target.note,
            created_at: target.created_at
          },
          links: target.links || [],
          images: (target.images || []).map((img) => ({
            id: img.id,
            storage_path: img.storage_path
          }))
        });

        setToast({ type: "info", text: "تم حذف المشترى.", action: "تراجع" });

        const orderIdToRefresh = target.order_id || selectedOrder?.id;
        await refreshPurchases(orderIdToRefresh);
        await refreshOrders(orderIdToRefresh);
      }

      setConfirmDelete(null);
    } catch (error) {
      console.error(error);
      if (confirmDelete.kind === "customer") {
        setToast({ type: "danger", text: customerFriendlyError(error, "فشل حذف العميل.") });
      } else {
        setToast({ type: "danger", text: "فشل حذف المشترى." });
      }
    } finally {
      setConfirmDeleteBusy(false);
    }
  };

  const openRenameOrderDialog = (order) => {
    if (!isRahaf || !order) return;
    setOrderDialog({
      kind: "rename",
      orderId: order.id,
      name: String(order.name || "")
    });
  };

  const openCreateOrderDialog = () => {
    if (!isRahaf) return;
    setOrderDialog({
      kind: "create",
      orderId: "",
      name: "",
      orderType: ORDER_TYPES.SHEIN,
      sheinType: SHEIN_ORDER_TYPES.NORMAL
    });
  };

  const openDeleteOrderDialog = (order) => {
    if (!isRahaf || !order) return;
    setOrderDialog({
      kind: "delete",
      orderId: order.id,
      name: String(order.name || "")
    });
  };

  const closeOrderDialog = () => {
    if (orderDialogBusy) return;
    setOrderDialog(null);
  };

  const submitRenameOrder = async () => {
    if (!orderDialog || orderDialog.kind !== "rename" || orderDialogBusy) return;
    const nextName = String(orderDialog.name || "").trim();
    if (!nextName) {
      setToast({ type: "warn", text: "اسم الطلب مطلوب." });
      return;
    }

    setOrderDialogBusy(true);
    try {
      await updateOrderName(orderDialog.orderId, nextName);
      setToast({ type: "success", text: "تم تعديل اسم الطلب." });
      await refreshOrders(orderDialog.orderId);
      setOrderDialog(null);
    } catch (error) {
      console.error(error);
      setToast({ type: "danger", text: error?.message || "فشل تعديل اسم الطلب." });
    } finally {
      setOrderDialogBusy(false);
    }
  };

  const confirmDeleteOrder = async () => {
    if (!orderDialog || orderDialog.kind !== "delete" || orderDialogBusy) return;

    setOrderDialogBusy(true);
    try {
      await deleteOrderById(orderDialog.orderId);
      setToast({ type: "success", text: "تم حذف الطلب." });
      await refreshOrders();
      setOrderDialog(null);
    } catch (error) {
      console.error(error);
      const message = String(error?.message || "");
      if (message.toLowerCase().includes("foreign key")) {
        setToast({ type: "danger", text: "لا يمكن حذف الطلب قبل حذف المشتريات المرتبطة به." });
      } else {
        setToast({ type: "danger", text: "فشل حذف الطلب." });
      }
    } finally {
      setOrderDialogBusy(false);
    }
  };

  const undoDeletePurchase = async () => {
    if (!deleteSnapshot || !selectedOrder) return;

    try {
      await restoreDeletedPurchase(deleteSnapshot);
      setDeleteSnapshot(null);
      setToast({ type: "success", text: "تم التراجع عن الحذف." });
      await refreshPurchases(selectedOrder.id);
      await refreshOrders(selectedOrder.id);
    } catch (error) {
      console.error(error);
      setToast({ type: "danger", text: "فشل التراجع عن الحذف." });
    }
  };

  const handleMarkPaid = (purchase) => {
    if (!purchase) return;
    setMenuPurchaseId("");
    setPaidPriceDialog({
      purchaseId: purchase.id,
      customerName: purchase.customer_name || "",
      price: String(purchase.price ?? ""),
      paidPrice: String(purchase.paid_price ?? purchase.price ?? "")
    });
  };

  const closePaidPriceDialog = () => {
    if (paidPriceDialogBusy) return;
    setPaidPriceDialog(null);
  };

  const submitPaidPriceDialog = async () => {
    if (!paidPriceDialog || paidPriceDialogBusy) return;
    const nextPaid = Number(paidPriceDialog.paidPrice);
    if (!Number.isFinite(nextPaid) || nextPaid < 0) {
      setToast({ type: "warn", text: "السعر المدفوع غير صالح." });
      return;
    }

    setPaidPriceDialogBusy(true);
    try {
      await markPurchasePaidPrice(paidPriceDialog.purchaseId, nextPaid);
      setToast({ type: "success", text: "تم تحديث السعر المدفوع." });
      setPaidPriceDialog(null);
      if (selectedOrder?.id) {
        await refreshPurchases(selectedOrder.id);
        await refreshOrders(selectedOrder.id);
      }
    } catch (error) {
      console.error(error);
      setToast({ type: "danger", text: "فشل تحديث السعر المدفوع." });
    } finally {
      setPaidPriceDialogBusy(false);
    }
  };

  const handleUpdateOrderStatus = async (nextStatus) => {
    if (!selectedOrder || !isRahaf || !editMode) return;
    if (!nextStatus) return;

    if (selectedOrderIsFullyCollected) {
      setToast({ type: "warn", text: "تم التحصيل: الحالة مقفلة تلقائيًا." });
      return;
    }

    const normalizedNext = String(nextStatus).trim();
    if (!Object.values(ORDER_STATUS).includes(normalizedNext)) {
      setToast({ type: "danger", text: "حالة الطلب غير صالحة." });
      return;
    }

    if (normalizedNext === ORDER_STATUS.COLLECTED && !selectedOrderIsFullyCollected) {
      setToast({ type: "danger", text: "لا يمكن اختيار تم التحصيل قبل اكتمال التحصيل لكل المشتريات." });
      return;
    }

    setOrderStatusSaving(true);
    try {
      const { status, payload } = await updateOrderWorkflowStatus(selectedOrder.id, normalizedNext);
      setOrders((prev) =>
        prev.map((order) =>
          String(order.id) === String(selectedOrder.id)
            ? {
                ...order,
                arrived: !!payload.arrived,
                placedAtPickup: !!payload.placed_at_pickup,
                status
              }
            : order
        )
      );
      setToast({ type: "success", text: "تم تحديث حالة الطلب." });
    } catch (error) {
      console.error(error);
      setToast({ type: "danger", text: "فشل تحديث حالة الطلب." });
    } finally {
      setOrderStatusSaving(false);
    }
  };

  const notifyViaWhatsapp = async (purchase) => {
    if (!selectedOrder?.arrived) {
      setToast({ type: "warn", text: "يجب أن تكون الطلبية واصلة أولًا." });
      return;
    }

    try {
      const target = await resolvePurchaseWhatsappTarget(purchase);
      const message = buildArrivalNotifyMessage({
        pickupPoint: purchase.pickup_point,
        price: purchase.price,
        customerName: target.customerName
      });
      const url = buildWhatsappUrl(target.phone, message);
      window.open(url, "_blank", "noopener,noreferrer");
    } catch (error) {
      console.error(error);
      setToast({ type: "danger", text: error?.message || "تعذر فتح واتساب." });
    }
  };

  const submitCreateOrder = async () => {
    if (!orderDialog || orderDialog.kind !== "create" || orderDialogBusy) return;
    const nextName = String(orderDialog.name || "").trim();
    const nextOrderType = String(orderDialog.orderType || "").trim().toLowerCase();
    const nextSheinType = String(orderDialog.sheinType || "").trim().toLowerCase();

    if (!nextName) {
      setToast({ type: "warn", text: "اسم الطلب مطلوب." });
      return;
    }
    if (!Object.values(ORDER_TYPES).includes(nextOrderType)) {
      setToast({ type: "warn", text: "حددي نوع الطلب." });
      return;
    }
    if (nextOrderType === ORDER_TYPES.SHEIN && !Object.values(SHEIN_ORDER_TYPES).includes(nextSheinType)) {
      setToast({ type: "warn", text: "حددي نوع طلب شي إن." });
      return;
    }

    setOrderDialogBusy(true);
    try {
      const created = await createOrder({
        orderName: nextName,
        orderType: nextOrderType,
        sheinType: nextOrderType === ORDER_TYPES.SHEIN ? nextSheinType : null
      });
      const createdId = created?.id ? String(created.id) : "";
      setToast({ type: "success", text: "تم إنشاء طلب جديد." });
      await refreshOrders(createdId);
      if (createdId) setSelectedOrderId(createdId);
      setActiveTab("orders");
      setOrderDialog(null);
    } catch (error) {
      console.error(error);
      setToast({ type: "danger", text: error?.message || "فشل إنشاء الطلب." });
    } finally {
      setOrderDialogBusy(false);
    }
  };

  const inquirePickupViaWhatsapp = async (purchase) => {
    if (!selectedOrder?.arrived) {
      setToast({ type: "warn", text: "يجب أن تكون الطلبية واصلة أولًا." });
      return;
    }

    try {
      const target = await resolvePurchaseWhatsappTarget(purchase);
      const message = buildPickupInquiryMessage();
      const url = buildWhatsappUrl(target.phone, message);
      window.open(url, "_blank", "noopener,noreferrer");
    } catch (error) {
      console.error(error);
      setToast({ type: "danger", text: error?.message || "تعذر فتح واتساب." });
    }
  };

  const analyzeFormImagesWithGemini = async () => {
    if (!formOpen) {
      setToast({ type: "warn", text: "افتحي نموذج الإضافة/التعديل أولًا." });
      return;
    }
    if (!hasGeminiKey()) {
      setFormAiStatus({ text: "مفتاح Gemini غير مضبوط.", isError: true });
      return;
    }

    const files = formState.newFiles || [];
    const existingUrls =
      formMode === "edit"
        ? (formState.existingImages || [])
            .filter((img) => !formState.removeImageIds.includes(img.id))
            .map((img) => img.url)
            .filter(Boolean)
        : [];

    if (!files.length && !existingUrls.length) {
      setFormAiStatus({ text: "أضيفي صورًا أولًا.", isError: true });
      return;
    }

    setFormAiRunning(true);
    setFormAiStatus({ text: "جاري التحليل...", isError: false });
    setFormError("");

    try {
      const result = await runGeminiCartAnalysis({
        files,
        urls: existingUrls,
        onProgress: (text) => setFormAiStatus({ text, isError: false })
      });

      const total = resolveTotalFromGemini(result);
      if (!Number.isFinite(total) || total <= 0) {
        throw new Error("لم يتم العثور على أسعار واضحة. جربي صورًا أوضح.");
      }

      const rounded = Math.round(total);
      setFormState((prev) => ({
        ...prev,
        price: String(rounded),
        paidPrice: String(rounded)
      }));
      setFormAiStatus({ text: `تم احتساب السعر: ${rounded}`, isError: false });
    } catch (error) {
      setFormAiStatus({ text: error?.message || "فشل التحليل.", isError: true });
    } finally {
      setFormAiRunning(false);
    }
  };

  const exportPdfNative = async () => {
    if (!selectedOrder) return;
    if (pdfExporting) return;

    setPdfExporting(true);
    try {
      await exportOrderPdf({
        order: selectedOrder,
        purchases
      });
      setToast({ type: "success", text: "تم تصدير ملف PDF." });
    } catch (error) {
      console.error(error);
      setToast({ type: "danger", text: error?.message || "فشل تصدير PDF." });
    } finally {
      setPdfExporting(false);
    }
  };

  const handleHeaderSearchPick = (row) => {
    if (!row?.order_id || !row?.id) return;
    setActiveTab("orders");
    setOrdersMenuOpen(false);
    setSelectedOrderId(row.order_id);
    setMenuPurchaseId("");
    setSearch("");
    setHeaderSearchTouched(false);
    setHeaderSearchResults([]);
    setHighlightPurchaseId(String(row.id));
  };

  const handleMovePurchaseKanban = async (purchase, targetColumn) => {
    if (!purchase?.id || !selectedOrder) return;

    let patch = {};
    let successText = "";

    if (targetColumn === "pending") {
      patch = { picked_up: false, collected: false };
      successText = "تم نقل المشترى إلى قيد الانتظار.";
    } else if (targetColumn === "received") {
      patch = { picked_up: true, collected: false };
      successText = "تم نقل المشترى إلى تم الاستلام.";
    } else if (targetColumn === "collected") {
      patch = { picked_up: true, collected: true };
      successText = "تم نقل المشترى إلى تم التحصيل.";
    } else {
      return;
    }

    setKanbanMovingPurchaseId(String(purchase.id));
    try {
      const { error } = await sb.from("purchases").update(patch).eq("id", purchase.id);
      if (error) throw error;

      setPurchases((prev) =>
        prev.map((item) =>
          String(item.id) === String(purchase.id)
            ? {
                ...item,
                ...patch
              }
            : item
        )
      );

      setToast({ type: "success", text: successText });
      await refreshOrders(selectedOrder.id);
    } catch (error) {
      console.error(error);
      setToast({ type: "danger", text: "فشل نقل المشترى بين الأعمدة." });
    } finally {
      setKanbanMovingPurchaseId("");
    }
  };

  const speedDialActions = useMemo(() => {
    if (!isMobile) return [];

    const actions = [];

    if (allowedTabs.includes("customers")) {
      if (activeTab === "customers") {
        actions.push({
          id: "tab-orders",
          label: "الطلبات",
          icon: "📦",
          onClick: () => setActiveTab("orders")
        });
      } else {
        actions.push({
          id: "tab-customers",
          label: "العملاء",
          icon: "👥",
          onClick: () => setActiveTab("customers")
        });
      }
    }

    return actions;
  }, [
    activeTab,
    allowedTabs,
    isMobile
  ]);

  if (profile.loading) {
    return (
      <div className="orders-page orders-loading-screen" dir="rtl">
        <SessionLoader />
      </div>
    );
  }

  if (!profile.authenticated) {
    return (
      <div className="orders-page orders-auth-screen" dir="rtl">
        <div className="orders-auth-empty">
          <div className="orders-auth-logo" aria-hidden="true">
            <Icon name="package" className="icon" />
          </div>
          <div className="orders-auth-brand">She Store</div>
          <h2>مرحباً بك</h2>
          <p>يرجى تسجيل الدخول للمتابعة</p>
          <a href="#/login" className="auth-primary-btn">
            تسجيل الدخول
          </a>
        </div>
      </div>
    );
  }

  if (!canUseOrdersWorkbench) {
    return (
      <div className="orders-page orders-loading-screen" dir="rtl">
        <div className="legacy-note legacy-note-danger">
          <h2>لا يمكن فتح صفحة الطلبات</h2>
          <p>
            هذا الحساب لا يملك صلاحية صفحة الطلبات. سيتم تحويلك للصفحة المناسبة حسب الدور.
          </p>
          <a
            href={profile.role === "laaura" ? "#/pickuppoint" : "#/login"}
            className="mode auth-link"
          >
            متابعة
          </a>
        </div>
      </div>
    );
  }

  return (
    <div className="orders-page" dir="rtl">
      <>
        <div
          className={`global-overlay app-sidebar-overlay ${globalOpen ? "open" : ""}`}
          onClick={() => setGlobalOpen(false)}
        />

        <aside className={`global-sidebar app-sidebar-drawer ${globalOpen ? "open" : ""}`}>
          <div className="global-sidebar-head app-sidebar-head">
            <div className="app-sidebar-brand">
              <SheStoreLogo className="app-sidebar-logo-link" imageClassName="app-sidebar-logo-img" />
              <b>القائمة</b>
            </div>
            <button type="button" className="app-sidebar-close" onClick={() => setGlobalOpen(false)}>
              ✕
            </button>
          </div>

          <nav className="global-sidebar-nav app-sidebar-content">
            {visibleNavItems.map((item) => (
              <a
                key={item.id}
                className={`app-sidebar-link ${isSidebarItemActive(item.href) ? "active" : ""}`}
                href={item.href}
                onClick={() => setGlobalOpen(false)}
              >
                <AppNavIcon name={item.icon} className="icon" />
                <span>{item.label}</span>
              </a>
            ))}
          </nav>

          <button type="button" className="app-sidebar-link app-sidebar-danger" onClick={signOut}>
            تسجيل الخروج
          </button>
        </aside>
      </>

      <CommandHeader
        isRahaf={isRahaf}
        canAccessCustomers={allowedTabs.includes("customers")}
        activeTab={activeTab}
        onActiveTabChange={setActiveTab}
        search={search}
        onSearchChange={setSearch}
        searchCount={searchCount}
        editMode={editMode}
        onEditModeChange={setEditMode}
        onOpenSidebar={() => setGlobalOpen(true)}
        showOrdersMenuTrigger={activeTab === "orders"}
        onOpenOrdersMenu={() => setOrdersMenuOpen(true)}
        totalOrders={totalOrders}
        showDesktopOrdersViewToggle={isDesktop && activeTab === "orders" && !isReem}
        desktopOrdersView={desktopOrdersView}
        onDesktopOrdersViewChange={setDesktopOrdersView}
        Icon={Icon}
      />

      {String(search || "").trim() ? (
        <div className="orders-header-search-results">
          {headerSearchLoading ? (
            <div className="orders-search-hint workspace-loader">
              <SessionLoader label="جاري البحث عن المشتريات..." />
            </div>
          ) : headerSearchResults.length ? (
            <div className="orders-search-list">
              {headerSearchResults.map((row) => (
                <button
                  key={row.id}
                  type="button"
                  className="orders-search-item"
                  onClick={() => handleHeaderSearchPick(row)}
                >
                  <strong>{row.customer_name || "بدون اسم"}</strong>
                  <span>
                    {orderNameById.get(String(row.order_id)) || "طلب"} • {row.qty || 0} قطع •{" "}
                    {formatILS(row.price)} ₪
                  </span>
                </button>
              ))}
            </div>
          ) : headerSearchTouched ? (
            <div className="orders-search-hint">لا توجد مشتريات مطابقة.</div>
          ) : null}
        </div>
      ) : null}

      {activeTab === "orders" && !isMobile ? (
        <OrdersDrawer
          open={ordersMenuOpen}
          onClose={() => setOrdersMenuOpen(false)}
          groupedOrders={groupedOrders}
          ordersLoading={ordersLoading}
          ordersError={ordersError}
          selectedOrderId={selectedOrderId}
          onSelectOrder={setSelectedOrderId}
          isRahaf={isRahaf}
          onForceOrdersTab={() => setActiveTab("orders")}
          onCreateOrder={openCreateOrderDialog}
          onRenameOrder={openRenameOrderDialog}
          onDeleteOrder={openDeleteOrderDialog}
          totalOrders={totalOrders}
          statusLabel={statusLabel}
          Icon={Icon}
        />
      ) : null}

      {activeTab === "orders" && isMobile ? (
        <OrdersBottomSheet
          open={ordersMenuOpen}
          onClose={() => setOrdersMenuOpen(false)}
          groupedOrders={groupedOrders}
          ordersLoading={ordersLoading}
          ordersError={ordersError}
          selectedOrderId={selectedOrderId}
          onSelectOrder={setSelectedOrderId}
          isRahaf={isRahaf}
          onForceOrdersTab={() => setActiveTab("orders")}
          onCreateOrder={openCreateOrderDialog}
          onRenameOrder={openRenameOrderDialog}
          onDeleteOrder={openDeleteOrderDialog}
        />
      ) : null}

      <div className="workspace">
        <section className="workspace-main">
          {activeTab === "orders" ? (
            <>
              <OrdersTab
                selectedOrder={selectedOrder}
                selectedOrderStatus={selectedOrderStatus}
                orderStatusLocked={selectedOrderIsFullyCollected}
                orderStatusSaving={orderStatusSaving}
                purchaseStats={purchaseStats}
                isMobile={isMobile}
                isRahaf={isRahaf}
                isReem={isReem}
                editMode={editMode}
                onUpdateOrderStatus={handleUpdateOrderStatus}
                onOpenAddModal={openAddModal}
                onExportPdf={exportPdfNative}
                canExportPdf={!isReem}
                pdfExporting={pdfExporting}
                customersError={customersError}
                purchasesLoading={purchasesLoading}
                purchasesError={purchasesError}
                filteredPurchases={filteredPurchases}
                menuPurchaseId={menuPurchaseId}
                onTogglePurchaseMenu={(purchaseId) =>
                  setMenuPurchaseId((prev) => (String(prev) === String(purchaseId) ? "" : purchaseId))
                }
                onEditPurchase={openEditModal}
                onMarkPaid={handleMarkPaid}
                onDeletePurchase={handleDeletePurchase}
                onOpenLightbox={(images, index, title) => setLightbox({ open: true, images, index, title })}
                onInquireWhatsapp={inquirePickupViaWhatsapp}
                onNotifyWhatsapp={notifyViaWhatsapp}
                highlightPurchaseId={highlightPurchaseId}
                hidePurchaseGrid={isDesktop && desktopOrdersView === "kanban" && !isReem}
              />

              {isDesktop && desktopOrdersView === "kanban" && !isReem ? (
                <KanbanView
                  purchases={filteredPurchases}
                  movingPurchaseId={kanbanMovingPurchaseId}
                  onMovePurchase={handleMovePurchaseKanban}
                  onOpenLightbox={(images, index, title) => setLightbox({ open: true, images, index, title })}
                />
              ) : null}
            </>
          ) : activeTab === "customers" ? (
            <CustomersTab
              customerSearch={customerSearch}
              setCustomerSearch={setCustomerSearch}
              customers={customers}
              customersLoading={customersLoading}
              customersError={customersError}
              filteredCustomers={filteredCustomers}
              customerForm={customerForm}
              setCustomerForm={setCustomerForm}
              customerFormMessage={customerFormMessage}
              customerFormSaving={customerFormSaving}
              handleCreateCustomer={handleCreateCustomer}
              isRahaf={isRahaf}
              editingCustomerId={editingCustomerId}
              editingCustomerForm={editingCustomerForm}
              setEditingCustomerForm={setEditingCustomerForm}
              beginEditCustomer={beginEditCustomer}
              saveEditCustomer={saveEditCustomer}
              cancelEditCustomer={cancelEditCustomer}
              handleDeleteCustomer={handleDeleteCustomer}
              cityOptions={CUSTOMER_CITIES}
              pickupOptions={CUSTOMER_PICKUP_OPTIONS}
            />
          ) : null}
        </section>
      </div>

      {(isMobile || isTablet) && !globalOpen && !formOpen && !lightbox.open && !ordersMenuOpen ? (
        <SpeedDial actions={speedDialActions} position="bottom-right" size="large" />
      ) : null}

      <PurchaseFormModal
        open={formOpen}
        formMode={formMode}
        formState={formState}
        customers={customers}
        customersLoading={customersLoading}
        formSaving={formSaving}
        formAiRunning={formAiRunning}
        formAiStatus={formAiStatus}
        formUploadProgress={formUploadProgress}
        formError={formError}
        newFilePreviews={newFilePreviews}
        maxImages={MAX_IMAGES}
        bagOptions={BAG_OPTIONS}
        pickupOptions={CUSTOMER_PICKUP_OPTIONS}
        onClose={closeFormModal}
        onSubmit={submitPurchaseForm}
        onCustomerChange={onCustomerChange}
        onUpdateForm={updateForm}
        onAddLinkInput={addLinkInput}
        onRemoveLinkInput={removeLinkInput}
        onUpdateLinkValue={updateLinkValue}
        onAddNewImages={addNewImages}
        onAnalyzeWithGemini={analyzeFormImagesWithGemini}
        onToggleExistingImageRemoval={toggleExistingImageRemoval}
        onRemoveNewImage={removeNewImage}
        onReplaceNewImage={replaceNewImage}
        onOpenAddCustomerModal={openQuickCustomerModal}
        Icon={Icon}
      />

      <CustomerQuickAddModal
        open={quickCustomerOpen}
        form={quickCustomerForm}
        saving={quickCustomerSaving}
        status={quickCustomerStatus}
        cityOptions={CUSTOMER_CITIES}
        pickupOptions={CUSTOMER_PICKUP_OPTIONS}
        onClose={closeQuickCustomerModal}
        onSubmit={handleQuickCustomerSubmit}
        onChange={(patch) => setQuickCustomerForm((prev) => ({ ...prev, ...patch }))}
        Icon={Icon}
      />

      {confirmDelete ? (
        <div className="purchase-modal-backdrop" onClick={closeConfirmDelete}>
          <div
            className="purchase-modal-card purchase-modal-card-customer purchase-modal-card-delete"
            onClick={(event) => event.stopPropagation()}
          >
            <div className="purchase-modal-head">
              <h3>{confirmDelete.kind === "customer" ? "حذف عميل" : "حذف مشترى"}</h3>
              <button
                type="button"
                className="icon-btn tiny"
                onClick={closeConfirmDelete}
                disabled={confirmDeleteBusy}
              >
                <Icon name="close" className="icon" />
              </button>
            </div>

            <div className="purchase-modal-body delete-confirm-body">
              <p className="delete-confirm-text">
                {confirmDelete.kind === "customer"
                  ? "هل تريدين حذف هذا العميل؟"
                  : "هل تريدين حذف هذا المشترى؟"}
              </p>

              {confirmDelete.item?.name || confirmDelete.item?.customer_name ? (
                <div className="delete-confirm-target">
                  {confirmDelete.item?.name || confirmDelete.item?.customer_name}
                </div>
              ) : null}

              <div className="purchase-modal-foot">
                <button
                  type="button"
                  className="btn-ghost-light danger-btn"
                  onClick={confirmDeleteAction}
                  disabled={confirmDeleteBusy}
                >
                  {confirmDeleteBusy ? "جاري الحذف..." : "تأكيد الحذف"}
                </button>
                <button
                  type="button"
                  className="btn-ghost-light"
                  onClick={closeConfirmDelete}
                  disabled={confirmDeleteBusy}
                >
                  إلغاء
                </button>
              </div>
            </div>
          </div>
        </div>
      ) : null}

      {orderDialog ? (
        <div className="purchase-modal-backdrop" onClick={closeOrderDialog}>
          <div
            className="purchase-modal-card purchase-modal-card-customer purchase-modal-card-delete"
            onClick={(event) => event.stopPropagation()}
          >
            <div className="purchase-modal-head">
              <h3>
                {orderDialog.kind === "create"
                  ? "إضافة طلب جديد"
                  : orderDialog.kind === "rename"
                    ? "تعديل اسم الطلب"
                    : "حذف الطلب"}
              </h3>
              <button
                type="button"
                className="icon-btn tiny"
                onClick={closeOrderDialog}
                disabled={orderDialogBusy}
              >
                <Icon name="close" className="icon" />
              </button>
            </div>

            <div className="purchase-modal-body delete-confirm-body">
              {orderDialog.kind === "rename" || orderDialog.kind === "create" ? (
                <>
                  <label>
                    <span>اسم الطلب</span>
                    <input
                      type="text"
                      value={orderDialog.name}
                      onChange={(event) =>
                        setOrderDialog((prev) => (prev ? { ...prev, name: event.target.value } : prev))
                      }
                      disabled={orderDialogBusy}
                      autoFocus
                    />
                  </label>
                  {orderDialog.kind === "create" ? (
                    <label>
                      <span>نوع الطلب</span>
                      <select
                        value={orderDialog.orderType || ""}
                        onChange={(event) =>
                          setOrderDialog((prev) =>
                            prev
                              ? {
                                  ...prev,
                                  orderType: event.target.value,
                                  sheinType:
                                    event.target.value === ORDER_TYPES.SHEIN
                                      ? prev.sheinType || SHEIN_ORDER_TYPES.NORMAL
                                      : ""
                                }
                              : prev
                          )
                        }
                        disabled={orderDialogBusy}
                      >
                        {ORDER_TYPE_OPTIONS.map((option) => (
                          <option key={option.value} value={option.value}>
                            {option.label}
                          </option>
                        ))}
                      </select>
                    </label>
                  ) : null}
                  {orderDialog.kind === "create" && orderDialog.orderType === ORDER_TYPES.SHEIN ? (
                    <label>
                      <span>نوع شي إن</span>
                      <select
                        value={orderDialog.sheinType || ""}
                        onChange={(event) =>
                          setOrderDialog((prev) => (prev ? { ...prev, sheinType: event.target.value } : prev))
                        }
                        disabled={orderDialogBusy}
                      >
                        <option value="">اختاري النوع</option>
                        {SHEIN_TYPE_OPTIONS.map((option) => (
                          <option key={option.value} value={option.value}>
                            {option.label}
                          </option>
                        ))}
                      </select>
                    </label>
                  ) : null}
                  <div className="purchase-modal-foot">
                    <button
                      type="button"
                      className="btn-primary"
                      onClick={orderDialog.kind === "create" ? submitCreateOrder : submitRenameOrder}
                      disabled={orderDialogBusy}
                    >
                      {orderDialogBusy
                        ? orderDialog.kind === "create"
                          ? "جاري الإنشاء..."
                          : "جاري الحفظ..."
                        : orderDialog.kind === "create"
                          ? "إنشاء"
                          : "حفظ"}
                    </button>
                    <button
                      type="button"
                      className="btn-ghost-light"
                      onClick={closeOrderDialog}
                      disabled={orderDialogBusy}
                    >
                      إلغاء
                    </button>
                  </div>
                </>
              ) : (
                <>
                  <p className="delete-confirm-text">هل تريدين حذف هذا الطلب؟</p>
                  {orderDialog.name ? <div className="delete-confirm-target">{orderDialog.name}</div> : null}
                  <div className="purchase-modal-foot">
                    <button
                      type="button"
                      className="btn-ghost-light danger-btn"
                      onClick={confirmDeleteOrder}
                      disabled={orderDialogBusy}
                    >
                      {orderDialogBusy ? "جاري الحذف..." : "تأكيد الحذف"}
                    </button>
                    <button
                      type="button"
                      className="btn-ghost-light"
                      onClick={closeOrderDialog}
                      disabled={orderDialogBusy}
                    >
                      إلغاء
                    </button>
                  </div>
                </>
              )}
            </div>
          </div>
        </div>
      ) : null}

      {paidPriceDialog ? (
        <div className="purchase-modal-backdrop" onClick={closePaidPriceDialog}>
          <div
            className="purchase-modal-card purchase-modal-card-customer purchase-modal-card-delete"
            onClick={(event) => event.stopPropagation()}
          >
            <div className="purchase-modal-head">
              <h3>تعديل المدفوع</h3>
              <button
                type="button"
                className="icon-btn tiny"
                onClick={closePaidPriceDialog}
                disabled={paidPriceDialogBusy}
              >
                <Icon name="close" className="icon" />
              </button>
            </div>

            <div className="purchase-modal-body delete-confirm-body">
              {paidPriceDialog.customerName ? (
                <div className="delete-confirm-target">{paidPriceDialog.customerName}</div>
              ) : null}

              <label>
                <span>السعر الأصلي</span>
                <input
                  type="number"
                  value={paidPriceDialog.price}
                  disabled
                  className="readonly-field"
                />
              </label>

              <label>
                <span>السعر المدفوع</span>
                <input
                  type="number"
                  step="0.01"
                  min="0"
                  value={paidPriceDialog.paidPrice}
                  onChange={(event) =>
                    setPaidPriceDialog((prev) => (prev ? { ...prev, paidPrice: event.target.value } : prev))
                  }
                  disabled={paidPriceDialogBusy}
                  autoFocus
                />
              </label>

              <div className="purchase-modal-foot">
                <button
                  type="button"
                  className="btn-primary"
                  onClick={submitPaidPriceDialog}
                  disabled={paidPriceDialogBusy}
                >
                  {paidPriceDialogBusy ? "جاري الحفظ..." : "حفظ"}
                </button>
                <button
                  type="button"
                  className="btn-ghost-light"
                  onClick={closePaidPriceDialog}
                  disabled={paidPriceDialogBusy}
                >
                  إلغاء
                </button>
              </div>
            </div>
          </div>
        </div>
      ) : null}

      <LightboxModal
        lightbox={lightbox}
        onClose={() => setLightbox({ open: false, images: [], index: 0, title: "" })}
        onPrev={() =>
          setLightbox((prev) => ({
            ...prev,
            index: (prev.index - 1 + prev.images.length) % prev.images.length
          }))
        }
        onNext={() =>
          setLightbox((prev) => ({
            ...prev,
            index: (prev.index + 1) % prev.images.length
          }))
        }
        Icon={Icon}
      />

      {toast ? (
        <div className={`toast toast-${toast.type || "info"}`}>
          <span>{toast.text}</span>
          {toast.action === "تراجع" ? (
            <button type="button" onClick={undoDeletePurchase}>
              تراجع
            </button>
          ) : null}
        </div>
      ) : null}
    </div>
  );
}
