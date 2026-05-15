import { useCallback, useEffect, useMemo, useRef, useState } from "react";

const MIN_ZOOM = 1;
const MAX_ZOOM = 4;
const ZOOM_STEP = 0.25;

function toImageUrl(image) {
  if (!image) return "";
  if (typeof image === "string") return image;
  return image.url || "";
}

function getTouchDistance(touches) {
  if (!touches || touches.length < 2) return 0;
  const [firstTouch, secondTouch] = touches;
  const dx = secondTouch.clientX - firstTouch.clientX;
  const dy = secondTouch.clientY - firstTouch.clientY;
  return Math.hypot(dx, dy);
}

export default function LightboxModal({ lightbox, onClose, onPrev, onNext, Icon }) {
  const images = useMemo(() => {
    return (lightbox.images || []).map(toImageUrl).filter(Boolean);
  }, [lightbox.images]);

  const total = images.length;
  const index = total ? ((lightbox.index % total) + total) % total : 0;
  const touchStartX = useRef(null);
  const swipeResetTimerRef = useRef(null);
  const pinchStateRef = useRef({ active: false, distance: 0, zoom: MIN_ZOOM });
  const [swipeDirection, setSwipeDirection] = useState("");
  const [swipeKey, setSwipeKey] = useState(0);
  const [zoom, setZoom] = useState(MIN_ZOOM);

  const clampZoom = useCallback((value) => {
    if (!Number.isFinite(value)) return MIN_ZOOM;
    return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, value));
  }, []);

  const changeZoom = useCallback(
    (delta) => {
      setZoom((prev) => clampZoom(prev + delta));
    },
    [clampZoom]
  );

  const resetZoom = useCallback(() => {
    setZoom(MIN_ZOOM);
  }, []);

  const triggerSwipeAnimation = useCallback((direction) => {
    setSwipeDirection(direction);
    setSwipeKey((prev) => prev + 1);
    if (swipeResetTimerRef.current) window.clearTimeout(swipeResetTimerRef.current);
    swipeResetTimerRef.current = window.setTimeout(() => {
      setSwipeDirection("");
    }, 240);
  }, []);

  const handlePrev = useCallback(() => {
    resetZoom();
    triggerSwipeAnimation("right");
    onPrev();
  }, [onPrev, resetZoom, triggerSwipeAnimation]);

  const handleNext = useCallback(() => {
    resetZoom();
    triggerSwipeAnimation("left");
    onNext();
  }, [onNext, resetZoom, triggerSwipeAnimation]);

  useEffect(() => {
    if (!lightbox.open || total === 0) return;
    setZoom(MIN_ZOOM);
    touchStartX.current = null;
    pinchStateRef.current = { active: false, distance: 0, zoom: MIN_ZOOM };
  }, [index, lightbox.open, total]);

  useEffect(() => {
    if (!lightbox.open || total === 0) return undefined;

    const onKeyDown = (event) => {
      if (event.key === "Escape") {
        event.preventDefault();
        onClose();
        return;
      }

      if (event.ctrlKey || event.metaKey) {
        const isZoomIn =
          event.key === "+" ||
          event.key === "=" ||
          event.code === "Equal" ||
          event.code === "NumpadAdd";
        const isZoomOut =
          event.key === "-" ||
          event.key === "_" ||
          event.code === "Minus" ||
          event.code === "NumpadSubtract";
        const isZoomReset = event.key === "0" || event.code === "Digit0" || event.code === "Numpad0";

        if (isZoomIn) {
          event.preventDefault();
          changeZoom(ZOOM_STEP);
          return;
        }

        if (isZoomOut) {
          event.preventDefault();
          changeZoom(-ZOOM_STEP);
          return;
        }

        if (isZoomReset) {
          event.preventDefault();
          resetZoom();
          return;
        }
      }

      if (total < 2) return;
      if (event.key === "ArrowLeft") {
        event.preventDefault();
        handlePrev();
      }
      if (event.key === "ArrowRight") {
        event.preventDefault();
        handleNext();
      }
    };

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [changeZoom, handleNext, handlePrev, lightbox.open, onClose, resetZoom, total]);

  useEffect(() => {
    if (!lightbox.open || total === 0) return undefined;
    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => {
      document.body.style.overflow = previousOverflow;
    };
  }, [lightbox.open, total]);

  useEffect(
    () => () => {
      if (swipeResetTimerRef.current) window.clearTimeout(swipeResetTimerRef.current);
    },
    []
  );

  if (!lightbox.open || total === 0) return null;

  return (
    <div className="purchase-modal-backdrop lightbox-backdrop" onClick={onClose}>
      <div className="lightbox-card" onClick={(event) => event.stopPropagation()}>
        <div className="lightbox-head">
          <strong>{lightbox.title}</strong>
          <div className="lightbox-head-meta">
            {total > 1 ? (
              <span className="lightbox-count-badge">
                {index + 1}/{total}
              </span>
            ) : null}
            <div className="lightbox-zoom-controls" aria-label="التحكم بتكبير الصورة">
              <button type="button" className="lightbox-zoom-btn" onClick={() => changeZoom(-ZOOM_STEP)} disabled={zoom <= MIN_ZOOM}>
                -
              </button>
              <button type="button" className="lightbox-zoom-btn lightbox-zoom-btn-reset" onClick={resetZoom}>
                {Math.round(zoom * 100)}%
              </button>
              <button type="button" className="lightbox-zoom-btn" onClick={() => changeZoom(ZOOM_STEP)} disabled={zoom >= MAX_ZOOM}>
                +
              </button>
            </div>
            <button type="button" className="icon-btn tiny" onClick={onClose} aria-label="إغلاق">
              <Icon name="close" className="icon" />
            </button>
          </div>
        </div>

        <div
          className={`lightbox-body ${zoom > MIN_ZOOM ? "is-zoomed" : ""}`}
          onDoubleClick={() => {
            setZoom((prev) => (prev > MIN_ZOOM ? MIN_ZOOM : 2));
          }}
          onWheel={(event) => {
            if (!event.ctrlKey && !event.metaKey) return;
            event.preventDefault();
            changeZoom(event.deltaY < 0 ? ZOOM_STEP / 2 : -ZOOM_STEP / 2);
          }}
          onTouchStart={(event) => {
            if (event.touches.length >= 2) {
              const distance = getTouchDistance(event.touches);
              pinchStateRef.current = {
                active: distance > 0,
                distance,
                zoom
              };
              touchStartX.current = null;
              return;
            }

            if (zoom > MIN_ZOOM) {
              touchStartX.current = null;
              return;
            }

            touchStartX.current = event.changedTouches?.[0]?.clientX ?? null;
          }}
          onTouchMove={(event) => {
            if (event.touches.length < 2) return;
            const pinchState = pinchStateRef.current;
            const distance = getTouchDistance(event.touches);
            if (!pinchState.active || !distance) return;
            event.preventDefault();
            setZoom(clampZoom(pinchState.zoom * (distance / pinchState.distance)));
          }}
          onTouchEnd={(event) => {
            if (pinchStateRef.current.active) {
              if (event.touches.length < 2) {
                pinchStateRef.current = { active: false, distance: 0, zoom };
              }
              touchStartX.current = null;
              return;
            }

            if (total < 2 || zoom > MIN_ZOOM || touchStartX.current === null) return;

            const startX = touchStartX.current;
            const endX = event.changedTouches?.[0]?.clientX;
            touchStartX.current = null;
            if (typeof endX !== "number") return;
            const dx = endX - startX;
            if (Math.abs(dx) < 40) return;
            if (dx < 0) handleNext();
            else handlePrev();
          }}
        >
          <div className="lightbox-zoom-stage">
            <img
              key={`lightbox-image-${index}-${swipeKey}`}
              className={`lightbox-image ${swipeDirection ? `swipe-${swipeDirection}` : ""}`}
              src={images[index]}
              alt="صورة"
              style={{ "--lightbox-zoom": zoom }}
            />
          </div>

          {total > 1 ? (
            <>
              <button type="button" className="icon-btn lightbox-nav lightbox-nav-prev" onClick={handlePrev} aria-label="الصورة السابقة">
                <Icon name="chevron-right" className="icon" />
              </button>
              <button type="button" className="icon-btn lightbox-nav lightbox-nav-next" onClick={handleNext} aria-label="الصورة التالية">
                <Icon name="chevron-left" className="icon" />
              </button>
            </>
          ) : null}
        </div>
      </div>
    </div>
  );
}
