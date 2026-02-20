import { useEffect, useState } from "react";
import actionsMenuIcon from "../../assets/icons/actions/menu-vertical.png";
import editIcon from "../../assets/icons/actions/edit.png";
import deleteIcon from "../../assets/icons/actions/delete.png";
import SessionLoader from "../common/SessionLoader";

export default function CustomersTab({
  customerSearch,
  setCustomerSearch,
  customers,
  customersLoading,
  customersError,
  filteredCustomers,
  customerForm,
  setCustomerForm,
  customerFormMessage,
  customerFormSaving,
  handleCreateCustomer,
  isRahaf,
  editingCustomerId,
  editingCustomerForm,
  setEditingCustomerForm,
  beginEditCustomer,
  saveEditCustomer,
  cancelEditCustomer,
  handleDeleteCustomer,
  cityOptions,
  pickupOptions
}) {
  const [openCustomerMenuId, setOpenCustomerMenuId] = useState("");

  useEffect(() => {
    const onDocClick = (event) => {
      if (!event.target.closest("[data-customer-menu-root]")) {
        setOpenCustomerMenuId("");
      }
    };

    document.addEventListener("click", onDocClick);
    return () => document.removeEventListener("click", onDocClick);
  }, []);

  useEffect(() => {
    if (editingCustomerId) setOpenCustomerMenuId("");
  }, [editingCustomerId]);

  return (
    <>
      <div className="order-detail-header">
        <div>
          <h2>العملاء</h2>
          <p>{isRahaf ? "إدارة بيانات العملاء (إضافة، تعديل، حذف)." : "عرض بيانات العملاء فقط."}</p>
        </div>

        <div className="order-detail-actions">
          <input
            className="purchase-search"
            value={customerSearch}
            onChange={(event) => setCustomerSearch(event.target.value)}
            placeholder="بحث بالاسم أو الهاتف أو المدينة..."
          />
          <span className="status-chip pending">الإجمالي: {customers.length}</span>
        </div>
      </div>

      <div className={`customer-layout ${isRahaf ? "" : "customer-layout-readonly"}`.trim()}>
        {isRahaf ? (
          <form className="customer-form-card" onSubmit={handleCreateCustomer}>
            <h3>إضافة عميل</h3>

            <label>
              <span>الاسم</span>
              <input
                value={customerForm.name}
                onChange={(event) =>
                  setCustomerForm((prev) => ({ ...prev, name: event.target.value }))
                }
                placeholder="اسم العميل"
                disabled={customerFormSaving}
              />
            </label>

            <label>
              <span>رقم الهاتف</span>
              <input
                value={customerForm.phone}
                onChange={(event) =>
                  setCustomerForm((prev) => ({ ...prev, phone: event.target.value }))
                }
                placeholder="970xxxxxxxxx أو 972xxxxxxxxx"
                disabled={customerFormSaving}
              />
            </label>

            <label>
              <span>المدينة</span>
              <select
                value={customerForm.city}
                onChange={(event) =>
                  setCustomerForm((prev) => ({ ...prev, city: event.target.value }))
                }
                disabled={customerFormSaving}
              >
                {cityOptions.map((city) => (
                  <option key={city} value={city}>
                    {city}
                  </option>
                ))}
              </select>
            </label>

            <label>
              <span>نقطة الاستلام المعتادة</span>
              <select
                value={customerForm.pickup}
                onChange={(event) =>
                  setCustomerForm((prev) => ({ ...prev, pickup: event.target.value }))
                }
                disabled={customerFormSaving}
              >
                {pickupOptions.map((pickup) => (
                  <option key={pickup} value={pickup}>
                    {pickup}
                  </option>
                ))}
              </select>
            </label>

            <div className="customer-form-actions">
              <button type="submit" className="btn-primary" disabled={customerFormSaving}>
                {customerFormSaving ? "جاري الحفظ..." : "حفظ العميل"}
              </button>
            </div>
            {customerFormMessage ? <div className="modal-help">{customerFormMessage}</div> : null}
          </form>
        ) : null}

        <div className="customer-list-card">
          {customersLoading ? (
            <div className="workspace-empty workspace-loader">
              <SessionLoader label="جاري تحميل العملاء..." />
            </div>
          ) : null}
          {customersError ? <div className="workspace-empty workspace-error">{customersError}</div> : null}

          {!customersLoading && !customersError && !filteredCustomers.length ? (
            <div className="workspace-empty">لا يوجد عملاء مطابقون.</div>
          ) : null}

          {!customersLoading && !customersError && filteredCustomers.length ? (
            <div className="customer-list-grid">
              {filteredCustomers.map((customer) => {
                const isEditing = String(editingCustomerId) === String(customer.id);
                const form = isEditing ? editingCustomerForm : null;
                const isMenuOpen = String(openCustomerMenuId) === String(customer.id);

                return (
                  <article key={customer.id} className="customer-item-card">
                    {isEditing && form && isRahaf ? (
                      <>
                        <div className="customer-item-head">
                          <h4>تعديل العميل</h4>
                        </div>

                        <div className="customer-edit-grid">
                          <label>
                            <span>الاسم</span>
                            <input
                              value={form.name}
                              onChange={(event) =>
                                setEditingCustomerForm((prev) => ({ ...prev, name: event.target.value }))
                              }
                            />
                          </label>

                          <label>
                            <span>الهاتف</span>
                            <input
                              value={form.phone}
                              onChange={(event) =>
                                setEditingCustomerForm((prev) => ({ ...prev, phone: event.target.value }))
                              }
                            />
                          </label>

                          <label>
                            <span>المدينة</span>
                            <select
                              value={form.city}
                              onChange={(event) =>
                                setEditingCustomerForm((prev) => ({ ...prev, city: event.target.value }))
                              }
                            >
                              {cityOptions.map((city) => (
                                <option key={city} value={city}>
                                  {city}
                                </option>
                              ))}
                            </select>
                          </label>

                          <label>
                            <span>نقطة الاستلام</span>
                            <select
                              value={form.pickup}
                              onChange={(event) =>
                                setEditingCustomerForm((prev) => ({ ...prev, pickup: event.target.value }))
                              }
                            >
                              {pickupOptions.map((pickup) => (
                                <option key={pickup} value={pickup}>
                                  {pickup}
                                </option>
                              ))}
                            </select>
                          </label>
                        </div>

                        <div className="customer-item-actions">
                          <button type="button" className="btn-primary" onClick={() => saveEditCustomer(customer.id)}>
                            حفظ
                          </button>
                          <button type="button" className="btn-ghost-light" onClick={cancelEditCustomer}>
                            إلغاء
                          </button>
                        </div>
                      </>
                    ) : (
                      <>
                        <div className="customer-item-head">
                          <div>
                            <h4>{customer.name || "—"}</h4>
                            <p>{customer.phone || "—"}</p>
                          </div>
                          <span className="status-chip completed">{customer.city || "—"}</span>
                        </div>

                        <div className="customer-item-meta">
                          <span>نقطة الاستلام: {customer.usual_pickup_point || "—"}</span>
                        </div>

                        {isRahaf ? (
                          <div className="customer-item-actions">
                            <div className="customer-actions-menu-wrap" data-customer-menu-root>
                              <button
                                type="button"
                                className="icon-btn tiny actions-menu-trigger"
                                aria-label={"\u0625\u062C\u0631\u0627\u0621\u0627\u062A \u0627\u0644\u0639\u0645\u064A\u0644"}
                                aria-haspopup="menu"
                                aria-expanded={isMenuOpen}
                                onClick={(event) => {
                                  event.stopPropagation();
                                  setOpenCustomerMenuId((prev) =>
                                    String(prev) === String(customer.id) ? "" : customer.id
                                  );
                                }}
                              >
                                <img src={actionsMenuIcon} alt="" aria-hidden="true" />
                              </button>

                              {isMenuOpen ? (
                                <div className="actions-menu-pop customer-actions-menu" role="menu">
                                  <button
                                    type="button"
                                    className="actions-menu-item"
                                    role="menuitem"
                                    onClick={() => {
                                      setOpenCustomerMenuId("");
                                      beginEditCustomer(customer);
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
                                      setOpenCustomerMenuId("");
                                      handleDeleteCustomer(customer);
                                    }}
                                  >
                                    <img src={deleteIcon} alt="" aria-hidden="true" className="actions-menu-item-icon" />
                                    <span>{"\u062D\u0630\u0641"}</span>
                                  </button>
                                </div>
                              ) : null}
                            </div>
                          </div>
                        ) : null}
                      </>
                    )}
                  </article>
                );
              })}
            </div>
          ) : null}
        </div>
      </div>
    </>
  );
}
