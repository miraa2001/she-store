const activeBodyLocks = new Set();

function applyBodyScrollState() {
  if (typeof document === "undefined") return;
  document.body.style.overflow = activeBodyLocks.size > 0 ? "hidden" : "";
}

export function setBodyScrollLock(lockKey, isLocked) {
  if (!lockKey) return;
  if (isLocked) {
    activeBodyLocks.add(lockKey);
  } else {
    activeBodyLocks.delete(lockKey);
  }
  applyBodyScrollState();
}

