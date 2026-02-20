import { useMemo, useState } from "react";
import { formatILS } from "../../lib/orders";

const COL_PENDING = "pending";
const COL_RECEIVED = "received";
const COL_COLLECTED = "collected";

const COLUMNS = [
  { id: COL_PENDING, title: "قيد الانتظار" },
  { id: COL_RECEIVED, title: "تم الاستلام" },
  { id: COL_COLLECTED, title: "تم التحصيل" }
];

function resolveColumnId(purchase) {
  const pickedUp = !!purchase?.picked_up;
  const collected = !!purchase?.collected;
  if (pickedUp && collected) return COL_COLLECTED;
  if (pickedUp) return COL_RECEIVED;
  return COL_PENDING;
}

export default function KanbanView({ purchases, onMovePurchase, onOpenLightbox, movingPurchaseId }) {
  const [draggingId, setDraggingId] = useState("");
  const [dragOverColumn, setDragOverColumn] = useState("");

  const purchaseById = useMemo(() => {
    const map = new Map();
    (purchases || []).forEach((purchase) => map.set(String(purchase.id), purchase));
    return map;
  }, [purchases]);

  const columns = useMemo(() => {
    const seed = {
      [COL_PENDING]: [],
      [COL_RECEIVED]: [],
      [COL_COLLECTED]: []
    };

    (purchases || []).forEach((purchase) => {
      seed[resolveColumnId(purchase)].push(purchase);
    });

    return seed;
  }, [purchases]);

  const handleDrop = async (event, targetColumnId) => {
    event.preventDefault();
    const droppedId = event.dataTransfer.getData("text/plain") || draggingId;
    const purchase = purchaseById.get(String(droppedId));

    setDraggingId("");
    setDragOverColumn("");

    if (!purchase) return;
    if (resolveColumnId(purchase) === targetColumnId) return;
    await onMovePurchase?.(purchase, targetColumnId);
  };

  return (
    <section className="kanban-board" aria-label="عرض كانبان للمشتريات">
      {COLUMNS.map((column) => (
        <section
          key={column.id}
          className={`kanban-column ${dragOverColumn === column.id ? "drag-over" : ""}`}
          onDragOver={(event) => {
            event.preventDefault();
            setDragOverColumn(column.id);
          }}
          onDragLeave={() => setDragOverColumn((prev) => (prev === column.id ? "" : prev))}
          onDrop={(event) => handleDrop(event, column.id)}
        >
          <header className="kanban-column-head">
            <h3>{column.title}</h3>
            <span className="kanban-count">{columns[column.id].length}</span>
          </header>

          <div className="kanban-column-list">
            {columns[column.id].map((purchase) => {
              const images = Array.isArray(purchase.images) ? purchase.images.filter((img) => img?.url) : [];
              const isMoving = String(movingPurchaseId || "") === String(purchase.id);

              return (
                <article
                  key={purchase.id}
                  className={`kanban-card ${isMoving ? "is-moving" : ""}`}
                  draggable={!isMoving}
                  onDragStart={(event) => {
                    setDraggingId(String(purchase.id));
                    event.dataTransfer.effectAllowed = "move";
                    event.dataTransfer.setData("text/plain", String(purchase.id));
                  }}
                  onDragEnd={() => {
                    setDraggingId("");
                    setDragOverColumn("");
                  }}
                >
                  <div className="kanban-card-top">
                    <strong>{purchase.customer_name || "—"}</strong>
                    <span className="kanban-drag-handle" aria-hidden="true">
                      ⋮⋮
                    </span>
                  </div>

                  <div className="kanban-thumbs" dir="ltr">
                    {images.slice(0, 3).map((img, index) => (
                      <button
                        key={img.id || `${purchase.id}-thumb-${index}`}
                        type="button"
                        className="kanban-thumb"
                        onClick={() => onOpenLightbox?.(images, index, purchase.customer_name || "صورة")}
                      >
                        <img src={img.url} alt={`صورة ${index + 1}`} loading="lazy" />
                      </button>
                    ))}
                    {images.length > 3 ? (
                      <span className="kanban-thumb-more">+{images.length - 3}</span>
                    ) : null}
                  </div>

                  <div className="kanban-meta">
                    <span>{purchase.qty || 0} قطع</span>
                    <span>{formatILS(purchase.price)} ₪</span>
                  </div>
                </article>
              );
            })}
          </div>
        </section>
      ))}
    </section>
  );
}
