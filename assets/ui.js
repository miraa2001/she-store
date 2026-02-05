(function(){
  const ensureContainers = ()=>{
    let toastContainer = document.getElementById("ui-toast-container");
    if(!toastContainer){
      toastContainer = document.createElement("div");
      toastContainer.id = "ui-toast-container";
      toastContainer.className = "toast-container";
      document.body.appendChild(toastContainer);
    }
    let modalRoot = document.getElementById("ui-modal-root");
    if(!modalRoot){
      modalRoot = document.createElement("div");
      modalRoot.id = "ui-modal-root";
      document.body.appendChild(modalRoot);
    }
    return { toastContainer, modalRoot };
  };

  function showToast(message, {type="info", actionText=null, onAction=null, duration=4500} = {}){
    const { toastContainer } = ensureContainers();
    const toast = document.createElement("div");
    toast.className = `toast ${type}`;
    toast.innerHTML = `<span>${message}</span>`;
    let timer = null;
    let remaining = duration;
    let start = Date.now();

    const close = ()=>{
      if(timer) clearTimeout(timer);
      toast.remove();
    };

    if(actionText){
      const btn = document.createElement("button");
      btn.className = "toast-action";
      btn.type = "button";
      btn.textContent = actionText;
      btn.onclick = ()=>{
        if(onAction) onAction();
        close();
      };
      toast.appendChild(btn);
    }

    const schedule = ()=>{
      start = Date.now();
      timer = setTimeout(close, remaining);
    };
    schedule();
    toast.addEventListener("mouseenter", ()=>{
      if(timer) clearTimeout(timer);
      remaining -= (Date.now() - start);
    });
    toast.addEventListener("mouseleave", schedule);

    toastContainer.appendChild(toast);
  }

  function confirmModal({title, message, confirmText="تأكيد", cancelText="إلغاء"} = {}){
    const { modalRoot } = ensureContainers();
    return new Promise(resolve=>{
      modalRoot.innerHTML = "";
      const backdrop = document.createElement("div");
      backdrop.className = "modal-backdrop";
      backdrop.tabIndex = -1;
      const modal = document.createElement("div");
      modal.className = "modal";
      modal.setAttribute("role","dialog");
      modal.setAttribute("aria-modal","true");
      const titleId = `modal-title-${Date.now()}`;
      const descId = `modal-desc-${Date.now()}`;
      modal.setAttribute("aria-labelledby", titleId);
      modal.setAttribute("aria-describedby", descId);
      modal.innerHTML = `
        <h3 id="${titleId}">${title || "تأكيد"}</h3>
        <p id="${descId}">${message || ""}</p>
        <div class="row">
          <button class="btn btn-secondary" type="button" data-cancel>${cancelText}</button>
          <button class="btn primary" type="button" data-confirm>${confirmText}</button>
        </div>
      `;

      const cancelBtn = modal.querySelector("[data-cancel]");
      const confirmBtn = modal.querySelector("[data-confirm]");
      const focusable = ()=> modal.querySelectorAll('button,[href],input,select,textarea,[tabindex]:not([tabindex="-1"])');
      let lastFocused = document.activeElement;

      const cleanup = (result)=>{
        document.removeEventListener("keydown", onKey);
        modalRoot.innerHTML = "";
        if(lastFocused && lastFocused.focus) lastFocused.focus();
        resolve(result);
      };

      const onKey = (e)=>{
        if(e.key === "Escape"){ e.preventDefault(); cleanup(false); }
        if(e.key === "Tab"){
          const items = Array.from(focusable());
          if(!items.length) return;
          const first = items[0];
          const last = items[items.length-1];
          if(e.shiftKey && document.activeElement === first){
            e.preventDefault(); last.focus();
          } else if(!e.shiftKey && document.activeElement === last){
            e.preventDefault(); first.focus();
          }
        }
        if(e.key === "Enter" && document.activeElement === confirmBtn){
          e.preventDefault(); cleanup(true);
        }
      };

      cancelBtn.onclick = ()=> cleanup(false);
      confirmBtn.onclick = ()=> cleanup(true);
      backdrop.onclick = (e)=>{ if(e.target === backdrop) cleanup(false); };
      document.addEventListener("keydown", onKey);

      backdrop.appendChild(modal);
      modalRoot.appendChild(backdrop);
      setTimeout(()=>{ confirmBtn.focus(); }, 0);
    });
  }

  function renderSkeletonRows(tbodyEl, rowCount, colCount){
    if(!tbodyEl) return;
    tbodyEl.innerHTML = "";
    for(let i=0;i<rowCount;i++){
      const tr = document.createElement("tr");
      for(let c=0;c<colCount;c++){
        const td = document.createElement("td");
        const sk = document.createElement("div");
        sk.className = "skeleton";
        td.appendChild(sk);
        tr.appendChild(td);
      }
      tbodyEl.appendChild(tr);
    }
  }

  function statusBadge(label, variant="neutral"){
    const safe = String(label || "");
    return `<span class="badge badge-${variant}">${safe}</span>`;
  }

  window.showToast = showToast;
  window.confirmModal = confirmModal;
  window.renderSkeletonRows = renderSkeletonRows;
  window.statusBadge = statusBadge;
})();
