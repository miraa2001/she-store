import { useEffect, useState } from "react";
import SessionLoader from "../common/SessionLoader";
import actionsMenuIcon from "../../assets/icons/actions/menu-vertical.png";
import editIcon from "../../assets/icons/actions/edit.png";
import deleteIcon from "../../assets/icons/actions/delete.png";

export default function OrdersBottomSheet({
  open = false,
  onClose,
  groupedOrders,
  ordersLoading,
  ordersError,
  selectedOrderId,
  onSelectOrder,
  isRahaf,
  onForceOrdersTab,
  onCreateOrder,
  onRenameOrder,
  onDeleteOrder
}) {
  const [openActionsOrderId, setOpenActionsOrderId] = useState("");

  useEffect(() => {
    if (!open) return undefined;

    const onKeyDown = (event) => {
      if (event.key !== "Escape") return;
      if (openActionsOrderId) {
        setOpenActionsOrderId("");
        return;
      }
      onClose?.();
    };

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [open, onClose, openActionsOrderId]);

  useEffect(() => {
    if (!open) {
      setOpenActionsOrderId("");
      return undefined;
    }

    const onDocClick = (event) => {
      if (!event.target.closest("[data-order-actions-root]")) {
        setOpenActionsOrderId("");
      }
    };

    document.addEventListener("click", onDocClick);
    return () => document.removeEventListener("click", onDocClick);
  }, [open]);

  if (!open) return null;

  return (
    <div className="orders-sheet-overlay" onClick={onClose}>
      <div className="orders-sheet" onClick={(event) => event.stopPropagation()}>
        <div className="orders-sheet-handle" />

        <div className="orders-sheet-head">
          <strong>اختيار الطلب</strong>
          <div className="orders-drawer-actions">
            {isRahaf ? (
              <button
                type="button"
                className="orders-menu-create-btn"
                onClick={() => {
                  setOpenActionsOrderId("");
                  onCreateOrder?.();
                  onClose?.();
                }}
              >
                + طلب جديد
              </button>
            ) : null}
            <button type="button" className="icon-btn tiny" onClick={onClose} aria-label="إغلاق قائمة الطلبات">
              ✕
            </button>
          </div>
        </div>

        <div className="orders-sheet-list">
          {ordersLoading ? (
            <div className="orders-horizontal-empty workspace-loader">
              <SessionLoader label="جاري تحميل الطلبات..." />
            </div>
          ) : null}

          {!ordersLoading && ordersError ? <div className="orders-horizontal-empty orders-horizontal-error">{ordersError}</div> : null}

          {!ordersLoading && !ordersError && !groupedOrders.length ? (
            <div className="orders-horizontal-empty">لا توجد طلبات مطابقة.</div>
          ) : null}

          {!ordersLoading && !ordersError
            ? groupedOrders.map((group) => (
                <div key={group.month} className="orders-sheet-group">
                  <span className="orders-sheet-month">{group.month}</span>

                  <div className="orders-sheet-group-list">
                    {group.orders.map((order) => {
                      const selected = String(selectedOrderId) === String(order.id);
                      const isActionsOpen = String(openActionsOrderId) === String(order.id);

                      return (
                        <div key={order.id} className="orders-sheet-item-wrap" data-order-row-wrap>
                          <button
                            type="button"
                            className={`orders-sheet-item ${selected ? "selected" : ""}`}
                            onClick={() => {
                              setOpenActionsOrderId("");
                              onSelectOrder(order.id);
                              if (isRahaf) onForceOrdersTab();
                              onClose?.();
                            }}
                          >
                            <span className="order-name">{order.name}</span>
                          </button>

                          {isRahaf ? (
                            <div className="order-row-actions" data-order-actions-root onClick={(event) => event.stopPropagation()}>
                              <button
                                type="button"
                                className="icon-btn tiny actions-menu-trigger"
                                aria-label={"\u0625\u062C\u0631\u0627\u0621\u0627\u062A \u0627\u0644\u0637\u0644\u0628"}
                                aria-haspopup="menu"
                                aria-expanded={isActionsOpen}
                                onClick={() => {
                                  setOpenActionsOrderId((prev) =>
                                    String(prev) === String(order.id) ? "" : order.id
                                  );
                                }}
                              >
                                <img src={actionsMenuIcon} alt="" aria-hidden="true" />
                              </button>

                              {isActionsOpen ? (
                                <div className="actions-menu-pop order-row-menu-pop" role="menu">
                                  <button
                                    type="button"
                                    className="actions-menu-item"
                                    role="menuitem"
                                    onClick={() => {
                                      onRenameOrder?.(order);
                                      setOpenActionsOrderId("");
                                      onClose?.();
                                    }}
                                  >
                                    <img src={editIcon} alt="" aria-hidden="true" className="actions-menu-item-icon" />
                                    <span>{"\u062A\u0639\u062F\u064A\u0644"}</span>
                                  </button>
                                  <button
                                    type="button"
                                    className="actions-menu-item danger"
                                    role="menuitem"
                                    onClick={() => {
                                      onDeleteOrder?.(order);
                                      setOpenActionsOrderId("");
                                      onClose?.();
                                    }}
                                  >
                                    <img src={deleteIcon} alt="" aria-hidden="true" className="actions-menu-item-icon" />
                                    <span>{"\u062D\u0630\u0641"}</span>
                                  </button>
                                </div>
                              ) : null}
                            </div>
                          ) : null}
                        </div>
                      );
                    })}
                  </div>
                </div>
              ))
            : null}
        </div>
      </div>
    </div>
  );
}
