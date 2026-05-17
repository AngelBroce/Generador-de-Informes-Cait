/**
 * UI Components Shared for CAIT Panama
 * Generates consistent sidebar and header.
 */

const sidebarLinks = [
  { href: "/presentacion/index.html", text: "Presentación del informe", icon: "description" },
  { href: "/resultados/index.html", text: "Resultados de las pruebas", icon: "science" },
  { href: "/conclusiones/index.html", text: "Conclusión", icon: "summarize" },
  { href: "/conclusiones/recomendaciones.html", text: "Recomendaciones", icon: "format_list_bulleted" },
  { href: "/certificados/calibracion.html", text: "Certificados de calibración", icon: "precision_manufacturing" },
  { href: "/certificados/protocolo.html", text: "Adjuntos de resultados", icon: "history_edu" },
  { href: "/certificados/asistencia.html", text: "Listados de asistencia", icon: "checklist" }
];

function initSidebar() {
  const container = document.getElementById('sidebar-container');
  if (!container) return;

  const currentPath = window.location.pathname;
  const currentHash = window.location.hash;

  let html = `
    <div class="px-6 mb-8">
      <h2 class="text-primary font-bold text-lg leading-tight cursor-pointer select-none" ondblclick="toggleConsole()">CAIT Panamá</h2>
      <p class="text-on-surface-variant text-sm">Generador de Informes <span class="bg-primary/10 text-primary text-[10px] px-1.5 py-0.5 rounded ml-1 font-bold cursor-pointer select-none" ondblclick="toggleConsole()">v2.2.6</span></p>
    </div>
    <nav class="flex-1 px-2 space-y-1">
  `;

  sidebarLinks.forEach(link => {
    let isActive = false;
    // Lógica para determinar si el link está activo
    if (link.href.includes('#')) {
      const [path, hash] = link.href.split('#');
      isActive = currentPath.includes(path) && currentHash === '#' + hash;
    } else {
      // Si el link actual tiene hash pero el link de la lista no, no es activo (ej: Presentacion base vs Empresa)
      isActive = currentPath.includes(link.href.split('.html')[0]) && !currentHash;
    }

    // Caso especial: si es la página base de presentación y no hay hash, activar el primero
    if (currentPath === '/presentacion/' && !currentHash && link.text === "Presentación del informe") isActive = true;

    html += `
      <a href="${link.href}" class="${isActive ? 'nav-active' : 'nav-item'} flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm transition-colors hover:bg-surface-container-high">
        <span class="material-symbols-outlined">${link.icon}</span><span>${link.text}</span>
      </a>
    `;
  });

  html += `</nav>`;
  container.innerHTML = html;
}

// También unificamos el Header para que sea igual en todos lados
function initHeader() {
  const header = document.getElementById('header-container');
  if (!header) return;

  header.className = "h-16 flex items-center justify-between px-lg bg-surface border-b border-outline-variant shrink-0";
  header.innerHTML = `
    <div class="flex items-center gap-3 text-primary font-bold text-lg cursor-pointer" onclick="location.href='/'">
      <img src="/static/logo.png" alt="Logo CAIT" class="h-10 w-auto"/>
      Generador de Informes CAIT
    </div>
    <div class="flex items-center gap-3">
      <button onclick="location.href='/config/evaluadores.html'" class="flex items-center gap-1 px-3 py-2 text-xs font-bold text-outline border border-outline-variant rounded-lg hover:bg-surface-container transition-colors">
        <span class="material-symbols-outlined" style="font-size:18px;">group</span> Evaluadores
      </button>
      <button onclick="location.href='/config/contrapartes.html'" class="flex items-center gap-1 px-3 py-2 text-xs font-bold text-outline border border-outline-variant rounded-lg hover:bg-surface-container transition-colors">
        <span class="material-symbols-outlined" style="font-size:18px;">badge</span> Contrapartes
      </button>
      <div class="w-px h-6 bg-outline-variant mx-1"></div>
      <button id="btn-load-global" class="flex items-center gap-1 px-4 py-2 text-sm font-bold text-primary border border-primary rounded-lg hover:bg-surface-container transition-colors">
        <span class="material-symbols-outlined" style="font-size:18px;">upload</span> Cargar Borrador
      </button>
      <button id="btn-save-global" class="flex items-center gap-1 px-4 py-2 text-sm font-bold text-on-primary bg-primary rounded-lg hover:opacity-90 transition-opacity">
        <span class="material-symbols-outlined" style="font-size:18px;">save</span> Guardar Borrador
      </button>
      <button id="btn-export-zip-global" class="flex items-center gap-1 px-4 py-2 text-sm font-bold text-on-secondary bg-secondary rounded-lg hover:opacity-90 transition-opacity">
        <span class="material-symbols-outlined" style="font-size:18px;">archive</span> Exportar ZIP
      </button>
      <div class="w-px h-6 bg-outline-variant mx-1"></div>
      <button id="btn-clear-all-global" class="flex items-center gap-1 px-3 py-2 text-xs font-bold text-error border border-error/30 rounded-lg hover:bg-error/10 transition-colors">
        <span class="material-symbols-outlined" style="font-size:18px;">delete_sweep</span> Limpiar Todo
      </button>
    </div>
  `;

  initGlobalButtons();
}

function initGlobalButtons() {
  const btnSave = document.getElementById('btn-save-global');
  const btnLoad = document.getElementById('btn-load-global');
  const btnZip = document.getElementById('btn-export-zip-global');
  const btnClear = document.getElementById('btn-clear-all-global');

  // Inyectar modales de Carga/Guardado
  if (!document.getElementById('modal-save-draft')) {
    const modalHtml = `
      <div id="modal-save-draft" class="fixed inset-0 bg-black/50 hidden items-center justify-center z-[9999] backdrop-blur-sm">
        <div class="bg-surface rounded-2xl shadow-2xl w-full max-w-md p-6 border border-outline-variant scale-95 transition-transform duration-200">
          <h3 class="text-xl font-bold mb-4 text-primary">Guardar Borrador</h3>
          <p class="text-sm text-outline mb-4">Introduce un nombre para identificar este informe:</p>
          <input type="text" id="input-save-draft-name" placeholder="Ej: Informe_Empresa_A" class="w-full bg-surface-container-highest border border-outline rounded-xl px-4 py-3 mb-6 focus:ring-2 focus:ring-primary outline-none"/>
          <div class="flex justify-end gap-3">
            <button id="btn-cancel-save" class="px-5 py-2.5 rounded-xl hover:bg-surface-container-high text-outline font-medium">Cancelar</button>
            <button id="btn-confirm-save" class="px-6 py-2.5 bg-primary text-on-primary rounded-xl font-bold hover:shadow-lg transition-all">Guardar</button>
          </div>
        </div>
      </div>
      <div id="modal-load-draft" class="fixed inset-0 bg-black/50 hidden items-center justify-center z-[9999] backdrop-blur-sm">
        <div class="bg-surface rounded-2xl shadow-2xl w-full max-w-2xl p-6 border border-outline-variant scale-95 transition-transform duration-200 flex flex-col max-h-[80vh]">
          <h3 class="text-xl font-bold mb-4 text-primary">Cargar Borrador</h3>
          <div id="drafts-list" class="flex-1 overflow-y-auto space-y-2 pr-2 mb-6">
            <!-- Cargando... -->
          </div>
          <div class="flex justify-end">
            <button id="btn-cancel-load" class="px-5 py-2.5 rounded-xl hover:bg-surface-container-high text-outline font-medium">Cerrar</button>
          </div>
        </div>
      </div>
    `;
    document.body.insertAdjacentHTML('beforeend', modalHtml);
  }

  const modalSave = document.getElementById('modal-save-draft');
  const modalLoad = document.getElementById('modal-load-draft');

  if (btnSave) btnSave.onclick = () => modalSave.classList.remove('hidden', 'flex'), modalSave.classList.add('flex');
  if (btnLoad) btnLoad.onclick = async () => {
    modalLoad.classList.remove('hidden', 'flex');
    modalLoad.classList.add('flex');
    await refreshDraftsList();
  };
  
  if (btnClear) btnClear.onclick = async () => {
    if (!confirm("¿Está seguro de que desea limpiar TODOS los datos del informe actual? Esta acción no se puede deshacer (a menos que haya guardado un borrador).")) {
      return;
    }
    
    try {
      if (window.showToast) window.showToast('Limpiando datos...');
      const res = await fetch('/api/report', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({})
      });
      
      if (res.ok) {
        if (window.showToast) window.showToast('Datos limpiados ✓');
        location.reload();
      } else {
        if (window.showToast) window.showToast('Error al limpiar datos', 'error');
      }
    } catch(e) {
      console.error(e);
      if (window.showToast) window.showToast('Error de conexión', 'error');
    }
  };

  document.getElementById('btn-cancel-save').onclick = () => modalSave.classList.add('hidden');
  document.getElementById('btn-cancel-load').onclick = () => modalLoad.classList.add('hidden');

  document.getElementById('btn-confirm-save').onclick = async () => {
    const name = document.getElementById('input-save-draft-name').value.trim();
    if (!name) return alert('Por favor, introduce un nombre.');
    
    if (window.showToast) window.showToast('Guardando borrador...');
    
    // Obtener datos actuales de la página
    const collectFn = window.collectFormData || (() => ({}));
    const pageData = collectFn();
    const existing = await fetch('/api/report').then(r => r.json()).catch(() => ({}));
    const merged = { ...existing, ...pageData, _draft_name: name };

    const res = await fetch('/api/report', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(merged)
    });

    if (res.ok) {
      if (window.showToast) window.showToast('¡Borrador guardado exitosamente! ✓');
      modalSave.classList.add('hidden');
      document.getElementById('input-save-draft-name').value = '';
    } else {
      if (window.showToast) window.showToast('Error al guardar', 'error');
    }
  };

  async function refreshDraftsList() {
    const list = document.getElementById('drafts-list');
    list.innerHTML = '<div class="py-10 text-center text-outline italic">Cargando borradores...</div>';
    try {
      const res = await fetch('/api/drafts');
      const drafts = await res.json();
      list.innerHTML = '';
      if (drafts.length === 0) {
        list.innerHTML = '<div class="py-10 text-center text-outline italic">No hay borradores guardados.</div>';
        return;
      }
      drafts.forEach(d => {
        const date = new Date(d.modified * 1000).toLocaleString();
        const item = document.createElement('div');
        item.className = 'p-4 border border-outline-variant rounded-xl hover:bg-primary-container/10 transition-all cursor-pointer flex justify-between items-center group';
        item.innerHTML = `
          <div>
            <div class="font-bold text-primary group-hover:text-primary-dark">${d.name}</div>
            <div class="text-xs text-outline">${date} • ${(d.size / 1024).toFixed(1)} KB</div>
          </div>
          <div class="flex gap-2">
            <button class="btn-delete p-2 hover:bg-error/10 text-error rounded-lg opacity-0 group-hover:opacity-100 transition-opacity">
              <span class="material-symbols-outlined">delete</span>
            </button>
            <button class="btn-open px-4 py-2 bg-primary text-on-primary rounded-lg font-bold text-sm">Abrir</button>
          </div>
        `;
        item.querySelector('.btn-open').onclick = async () => {
          if (window.showToast) window.showToast(`Cargando ${d.name}...`);
          const rRes = await fetch(`/api/report?name=${d.name}`);
          const data = await rRes.json();
          // Guardar como actual para que todas las páginas lo vean
          await fetch('/api/report', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
          });
          if (window.loadData) window.loadData();
          else location.reload(); // Recargar si no hay función de carga parcial
          modalLoad.classList.add('hidden');
        };
        item.querySelector('.btn-delete').onclick = async (e) => {
          e.stopPropagation();
          if (!confirm(`¿Eliminar el borrador "${d.name}"?`)) return;
          await fetch(`/api/drafts/${d.name}`, { method: 'DELETE' });
          refreshDraftsList();
        };
        list.appendChild(item);
      });
    } catch(e) {
      list.innerHTML = '<div class="py-10 text-center text-error italic">Error al cargar borradores.</div>';
    }
  }

  if (btnZip) btnZip.onclick = async () => {
    let targetFolder = null;
    try {
      if (window.pywebview && window.pywebview.api && window.pywebview.api.select_folder) {
        targetFolder = await window.pywebview.api.select_folder();
        if (!targetFolder) return; // Cancelado
      }
    } catch(e) { console.error('Error selecting folder', e); }

    if (window.showToast) window.showToast('Generando paquete ZIP...');
    try {
      const body = targetFolder ? JSON.stringify({ target_folder: targetFolder }) : '{}';
      const res = await fetch('/api/export-zip', { 
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: body
      });
      if (!res.ok) throw new Error('Servidor respondió con error');
      const data = await res.json();
      if (data.status === 'ok') {
        if(window.showToast) window.showToast('¡ZIP listo! Guardado en la ubicación seleccionada ✓');
        
        // Si no hubo carpeta seleccionada (ej: desarrollo en navegador), descargar normalmente
        if (!targetFolder) {
          const downloadUrl = `/api/download-zip/${data.filename}`;
          const iframe = document.createElement('iframe');
          iframe.style.display = 'none';
          iframe.src = downloadUrl;
          document.body.appendChild(iframe);
          setTimeout(() => document.body.removeChild(iframe), 3000);
        }
        
      } else {
        if(window.showToast) window.showToast('Error: ' + (data.message || 'No se pudo generar el ZIP'), 'error');
      }
    } catch(e) {
      console.error(e);
      if(window.showToast) window.showToast('Error de conexión con el servidor', 'error');
    }
  };
}

function injectStyles() {
  const style = document.createElement('style');
  style.innerHTML = `
    #header-container { height: 64px; min-height: 64px; display: flex; background: #fcf9f8; }
    #sidebar-container { width: 256px; min-width: 256px; display: flex; background: #f6f3f2; }
    main { 
      animation: fadeIn 0.15s ease-out;
    }
    @keyframes fadeIn {
      from { opacity: 0; }
      to { opacity: 1; }
    }
    .btn-loading { opacity: 0.7; pointer-events: none; }
  `;
  document.head.appendChild(style);
}

function initDebugModal() {
  if (document.getElementById('modal-debug')) return;
  const modal = document.createElement('div');
  modal.id = 'modal-debug';
  modal.className = 'hidden fixed inset-0 bg-black/60 flex items-center justify-center z-[9999] backdrop-blur-sm';
  modal.innerHTML = `
    <div class="bg-surface p-6 rounded-2xl shadow-2xl w-80 border border-outline-variant animate-in fade-in zoom-in duration-200">
      <h3 class="text-lg font-bold text-primary mb-2">Modo Desarrollador</h3>
      <p class="text-xs text-outline mb-4">Ingrese la contraseña para ver los logs.</p>
      <div class="relative mb-6">
        <input type="password" id="debug-pass" class="w-full px-4 py-3 rounded-xl border border-outline bg-surface-container-low text-on-surface focus:border-primary outline-none transition-all" placeholder="Contraseña">
        <button onclick="toggleDebugPassVisibility()" class="absolute right-3 top-2.5 text-outline hover:text-primary transition-colors">
          <span id="eye-icon" class="material-symbols-outlined" style="font-size:20px;">visibility</span>
        </button>
      </div>
      <div class="flex justify-end gap-3">
        <button onclick="closeDebugModal()" class="px-4 py-2 text-xs font-bold text-outline hover:bg-surface-container-high rounded-lg transition-colors">Cerrar</button>
        <button onclick="confirmDebug()" class="px-6 py-2 text-xs font-bold bg-primary text-on-primary rounded-lg shadow-md hover:scale-105 transition-all">Activar</button>
      </div>
    </div>
  `;
  document.body.appendChild(modal);
}

window.toggleDebugPassVisibility = () => {
  const input = document.getElementById('debug-pass');
  const icon = document.getElementById('eye-icon');
  if (input.type === 'password') {
    input.type = 'text';
    icon.innerText = 'visibility_off';
  } else {
    input.type = 'password';
    icon.innerText = 'visibility';
  }
};

window.closeDebugModal = () => document.getElementById('modal-debug').classList.add('hidden');

window.confirmDebug = async () => {
  const pass = document.getElementById('debug-pass').value;
  if (!pass) return;
  try {
    const res = await fetch('/api/debug/console', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ password: pass })
    });
    const data = await res.json();
    if (data.status === 'ok') {
      document.getElementById('modal-debug').classList.add('hidden');
      document.getElementById('debug-pass').value = '';
      alert(data.message);
    } else {
      alert('Acceso denegado');
    }
  } catch (e) { console.error(e); }
};

document.addEventListener('DOMContentLoaded', () => {
  injectStyles();
  initHeader();
  initSidebar();
});

// Escuchar cambios de hash para actualizar sidebar sin recargar
window.addEventListener('hashchange', initSidebar);

async function toggleConsole() {
  initDebugModal();
  document.getElementById('modal-debug').classList.remove('hidden');
  document.getElementById('debug-pass').focus();
}
window.toggleConsole = toggleConsole;
