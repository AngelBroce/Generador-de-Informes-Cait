/**
 * UI Components Shared for CAIT Panama
 * Generates consistent sidebar, header, and Data Exchange modal (Import/Export no-PDF).
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
      <p class="text-on-surface-variant text-sm">Generador de Informes <span class="bg-primary/10 text-primary text-[10px] px-1.5 py-0.5 rounded ml-1 font-bold cursor-pointer select-none" ondblclick="toggleConsole()">v2.3.0</span></p>
    </div>
    <nav class="flex-1 px-2 space-y-1">
  `;

  sidebarLinks.forEach(link => {
    let isActive = false;
    if (link.href.includes('#')) {
      const [path, hash] = link.href.split('#');
      isActive = currentPath.includes(path) && currentHash === '#' + hash;
    } else {
      isActive = currentPath.includes(link.href.split('.html')[0]) && !currentHash;
    }

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

function initHeader() {
  const header = document.getElementById('header-container');
  if (!header) return;

  header.className = "h-16 flex items-center justify-between px-lg bg-surface border-b border-outline-variant shrink-0";
  header.innerHTML = `
    <div class="flex items-center gap-3 text-primary font-bold text-lg cursor-pointer" onclick="location.href='/'">
      <img src="/static/logo.png" alt="Logo CAIT" class="h-10 w-auto"/>
      Generador de Informes CAIT
    </div>
    <div class="flex items-center gap-2.5">
      <button onclick="location.href='/config/evaluadores.html'" class="flex items-center gap-1 px-2.5 py-1.5 text-xs font-bold text-outline border border-outline-variant rounded-lg hover:bg-surface-container transition-colors" title="Catálogo de Evaluadores">
        <span class="material-symbols-outlined" style="font-size:17px;">group</span> Evaluadores
      </button>
      <button onclick="location.href='/config/contrapartes.html'" class="flex items-center gap-1 px-2.5 py-1.5 text-xs font-bold text-outline border border-outline-variant rounded-lg hover:bg-surface-container transition-colors" title="Catálogo de Contrapartes">
        <span class="material-symbols-outlined" style="font-size:17px;">badge</span> Contrapartes
      </button>
      
      <div class="w-px h-6 bg-outline-variant mx-0.5"></div>
      
      <button id="btn-data-exchange-global" class="flex items-center gap-1.5 px-3 py-2 text-xs font-bold text-primary bg-primary/10 border border-primary/30 rounded-lg hover:bg-primary/20 transition-all shadow-sm" title="Importar o Exportar datos en formatos no-PDF (Excel, CSV, CAIT)">
        <span class="material-symbols-outlined" style="font-size:18px;">sync_alt</span> Importar / Exportar Datos
      </button>

      <button id="btn-load-global" class="flex items-center gap-1 px-3 py-2 text-xs font-bold text-primary border border-primary rounded-lg hover:bg-surface-container transition-colors">
        <span class="material-symbols-outlined" style="font-size:17px;">folder_open</span> Borradores
      </button>
      <button id="btn-save-global" class="flex items-center gap-1 px-3 py-2 text-xs font-bold text-on-primary bg-primary rounded-lg hover:opacity-90 transition-opacity">
        <span class="material-symbols-outlined" style="font-size:17px;">save</span> Guardar
      </button>
      <button id="btn-export-zip-global" class="flex items-center gap-1 px-3 py-2 text-xs font-bold text-on-secondary bg-secondary rounded-lg hover:opacity-90 transition-opacity">
        <span class="material-symbols-outlined" style="font-size:17px;">picture_as_pdf</span> Generar PDF / ZIP
      </button>
      
      <div class="w-px h-6 bg-outline-variant mx-0.5"></div>
      <button id="btn-clear-all-global" class="flex items-center gap-1 px-2.5 py-1.5 text-xs font-bold text-error border border-error/30 rounded-lg hover:bg-error/10 transition-colors" title="Limpiar formulario">
        <span class="material-symbols-outlined" style="font-size:17px;">delete_sweep</span>
      </button>
    </div>
  `;

  initGlobalButtons();
  initDataExchangeModal();
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
          <h3 class="text-xl font-bold mb-4 text-primary">Cargar Borrador Guardado</h3>
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

  if (btnSave) btnSave.onclick = () => {
    modalSave.classList.remove('hidden', 'flex');
    modalSave.classList.add('flex');
    document.getElementById('input-save-draft-name').focus();
  };
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
          await fetch('/api/report', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
          });
          if (window.loadData) window.loadData();
          else location.reload();
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
        if (!targetFolder) return;
      }
    } catch(e) { console.error('Error selecting folder', e); }

    if (window.showToast) window.showToast('Generando paquete ZIP y PDF...');
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

// =========================================================================
// MODAL DE IMPORTACIÓN / EXPORTACIÓN DE DATOS (NO-PDF)
// =========================================================================
function initDataExchangeModal() {
  if (document.getElementById('modal-data-exchange')) return;

  const modalHtml = `
    <div id="modal-data-exchange" class="fixed inset-0 bg-black/60 hidden items-center justify-center z-[9999] backdrop-blur-sm p-4">
      <div class="bg-surface rounded-2xl shadow-2xl w-full max-w-3xl border border-outline-variant scale-95 transition-transform duration-200 flex flex-col max-h-[90vh] overflow-hidden text-on-surface">
        
        <!-- Header del Modal -->
        <div class="px-6 py-4 bg-surface-container-low border-b border-outline-variant flex items-center justify-between">
          <div class="flex items-center gap-3">
            <span class="material-symbols-outlined text-primary" style="font-size:28px;">swap_horiz</span>
            <div>
              <h2 class="text-lg font-bold text-primary">Intercambio y Portabilidad de Datos</h2>
              <p class="text-xs text-outline">Transfiere informes, pacientes y bases de datos a otra computadora o aplicación sin PDF.</p>
            </div>
          </div>
          <button id="btn-close-exchange-modal" class="p-2 hover:bg-surface-container-high rounded-full text-outline hover:text-primary transition-colors">
            <span class="material-symbols-outlined">close</span>
          </button>
        </div>

        <!-- Pestañas -->
        <div class="flex border-b border-outline-variant bg-surface-container-lowest px-6 pt-2">
          <button id="tab-btn-export" class="px-5 py-2.5 font-bold text-sm border-b-2 border-primary text-primary transition-colors flex items-center gap-2">
            <span class="material-symbols-outlined" style="font-size:18px;">upload</span> Exportar Datos (Sacar)
          </button>
          <button id="tab-btn-import" class="px-5 py-2.5 font-bold text-sm border-b-2 border-transparent text-outline hover:text-on-surface transition-colors flex items-center gap-2">
            <span class="material-symbols-outlined" style="font-size:18px;">download</span> Importar y Auto-Registrar (Meter)
          </button>
        </div>

        <!-- Contenido de las Pestañas -->
        <div class="flex-1 overflow-y-auto p-6 bg-surface">
          
          <!-- PANEL DE EXPORTACIÓN -->
          <div id="panel-export" class="space-y-4">
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
              
              <!-- Card: CAIT Portable Case -->
              <div class="p-4 border border-outline-variant rounded-xl bg-surface-container-low hover:border-primary transition-all flex flex-col justify-between">
                <div>
                  <div class="flex items-center gap-2 text-primary font-bold text-sm mb-1">
                    <span class="material-symbols-outlined" style="font-size:20px;">description</span>
                    Informe Portátil (.cait)
                  </div>
                  <p class="text-xs text-outline mb-3">Guarda el caso clínico completo y los pacientes asociados en un archivo ligero para abrir en otra computadora.</p>
                </div>
                <button onclick="downloadExchangeFile('/api/export/cait', 'informe.cait')" class="w-full py-2 px-3 bg-primary text-on-primary rounded-lg text-xs font-bold hover:opacity-90 flex items-center justify-center gap-1.5 shadow-sm">
                  <span class="material-symbols-outlined" style="font-size:16px;">download</span> Descargar .cait
                </button>
              </div>

              <!-- Card: Excel XLSX -->
              <div class="p-4 border border-outline-variant rounded-xl bg-surface-container-low hover:border-primary transition-all flex flex-col justify-between">
                <div>
                  <div class="flex items-center gap-2 text-secondary font-bold text-sm mb-1">
                    <span class="material-symbols-outlined" style="font-size:20px;">table_view</span>
                    Libro de Excel (.xlsx)
                  </div>
                  <p class="text-xs text-outline mb-3">Exporta tablas con hojas separadas de Audiometría, Espirometría, Resumen y Catálogo de Pacientes con formato profesional.</p>
                </div>
                <button onclick="downloadExchangeFile('/api/export/excel', 'informe.xlsx')" class="w-full py-2 px-3 bg-secondary text-on-secondary rounded-lg text-xs font-bold hover:opacity-90 flex items-center justify-center gap-1.5 shadow-sm">
                  <span class="material-symbols-outlined" style="font-size:16px;">table_chart</span> Descargar .xlsx
                </button>
              </div>

              <!-- Card: CSV Estándar -->
              <div class="p-4 border border-outline-variant rounded-xl bg-surface-container-low hover:border-primary transition-all flex flex-col justify-between">
                <div>
                  <div class="flex items-center gap-2 text-on-surface font-bold text-sm mb-1">
                    <span class="material-symbols-outlined" style="font-size:20px;">csv</span>
                    Resultados en CSV (.csv)
                  </div>
                  <p class="text-xs text-outline mb-3">Exporta el listado de personas y resultados en texto delimitado por comas con codificación UTF-8 para otras bases de datos.</p>
                </div>
                <button onclick="downloadExchangeFile('/api/export/csv', 'resultados.csv')" class="w-full py-2 px-3 bg-surface-container-highest text-on-surface border border-outline rounded-lg text-xs font-bold hover:bg-surface-container-high flex items-center justify-center gap-1.5">
                  <span class="material-symbols-outlined" style="font-size:16px;">file_download</span> Descargar .csv
                </button>
              </div>

              <!-- Card: Paquete con Adjuntos -->
              <div class="p-4 border border-outline-variant rounded-xl bg-surface-container-low hover:border-primary transition-all flex flex-col justify-between">
                <div>
                  <div class="flex items-center gap-2 text-primary font-bold text-sm mb-1">
                    <span class="material-symbols-outlined" style="font-size:20px;">folder_zip</span>
                    Paquete Completo (.caitpkg)
                  </div>
                  <p class="text-xs text-outline mb-3">Empaqueta el caso junto con todos los certificados, imágenes y firmas adjuntas en un archivo ZIP portable.</p>
                </div>
                <button onclick="downloadExchangeFile('/api/export/caitpkg', 'caso_completo.caitpkg')" class="w-full py-2 px-3 bg-primary text-on-primary rounded-lg text-xs font-bold hover:opacity-90 flex items-center justify-center gap-1.5 shadow-sm">
                  <span class="material-symbols-outlined" style="font-size:16px;">archive</span> Descargar .caitpkg
                </button>
              </div>

            </div>

            <!-- Copia de seguridad completa del sistema -->
            <div class="p-4 border border-primary/20 bg-primary/5 rounded-xl flex items-center justify-between mt-4">
              <div>
                <div class="font-bold text-primary text-sm flex items-center gap-1.5">
                  <span class="material-symbols-outlined" style="font-size:18px;">security</span>
                  Copia de Seguridad del Sistema Completo (.caitbackup)
                </div>
                <p class="text-xs text-outline mt-0.5">Exporta todas las bases de datos (personas, evaluadores, contrapartes, plantillas y borradores) para migrar de PC.</p>
              </div>
              <button onclick="downloadExchangeFile('/api/system/backup', 'backup.caitbackup')" class="px-4 py-2 bg-primary text-on-primary rounded-lg text-xs font-bold hover:opacity-90 transition-opacity shrink-0 ml-4">
                Exportar Backup
              </button>
            </div>
          </div>

          <!-- PANEL DE IMPORTACIÓN -->
          <div id="panel-import" class="hidden space-y-4">
            
            <!-- Zona Dropzone -->
            <div id="exchange-dropzone" class="border-2 border-dashed border-outline-variant rounded-2xl p-8 text-center bg-surface-container-lowest hover:bg-surface-container-low hover:border-primary transition-all cursor-pointer">
              <input type="file" id="exchange-file-input" class="hidden" accept=".cait,.caitpkg,.json,.xlsx,.xls,.csv,.caitbackup,.zip"/>
              <span class="material-symbols-outlined text-primary mb-2" style="font-size:48px;">cloud_upload</span>
              <h4 class="font-bold text-sm text-primary mb-1">Arrastra aquí tu archivo o haz clic para seleccionar</h4>
              <p class="text-xs text-outline max-w-md mx-auto">Soporta archivos <strong>.cait, .caitpkg, .json, .xlsx, .csv, .caitbackup</strong>. La aplicación leerá los datos y los registrará automáticamente en la base de datos.</p>
              <div id="exchange-file-selected" class="mt-4 hidden items-center justify-center gap-2 text-xs font-bold text-primary bg-primary/10 py-1.5 px-3 rounded-lg w-fit mx-auto">
                <span class="material-symbols-outlined" style="font-size:16px;">check_circle</span>
                <span id="exchange-file-name">archivo.xlsx</span>
              </div>
            </div>

            <!-- Opciones según el tipo de archivo -->
            <div id="exchange-tabular-options" class="p-4 border border-outline-variant rounded-xl bg-surface-container-low hidden">
              <label class="text-xs font-bold text-primary uppercase block mb-2">Importar filas de Excel / CSV hacia:</label>
              <div class="flex gap-4">
                <label class="flex items-center gap-2 text-xs font-medium cursor-pointer">
                  <input type="radio" name="import-target-type" value="audiometria" checked class="text-primary focus:ring-primary"/>
                  Audiometría
                </label>
                <label class="flex items-center gap-2 text-xs font-medium cursor-pointer">
                  <input type="radio" name="import-target-type" value="espirometria" class="text-primary focus:ring-primary"/>
                  Espirometría
                </label>
              </div>
            </div>

            <!-- Botón de acción de importación -->
            <div class="flex justify-end gap-3 pt-2">
              <button id="btn-process-import" disabled class="w-full py-3 bg-primary text-on-primary rounded-xl font-bold text-sm hover:opacity-90 disabled:opacity-40 disabled:cursor-not-allowed transition-all shadow-md flex items-center justify-center gap-2">
                <span class="material-symbols-outlined" style="font-size:18px;">cloud_done</span>
                Procesar e Importar al Sistema
              </button>
            </div>

            <!-- Resultado de importación -->
            <div id="exchange-import-result" class="hidden p-4 rounded-xl border border-primary/30 bg-primary/10 text-on-surface text-xs space-y-1.5 animate-in fade-in duration-200">
              <!-- Mensaje dinámico de auto-registro -->
            </div>

          </div>

        </div>

        <!-- Footer del Modal -->
        <div class="px-6 py-3 bg-surface-container-low border-t border-outline-variant flex justify-between items-center text-xs text-outline">
          <span>CAIT Informes v2.3.0 • Sistema de Migración e Integración</span>
          <button id="btn-cancel-exchange" class="px-4 py-2 rounded-lg hover:bg-surface-container-high text-outline font-bold">Cerrar</button>
        </div>

      </div>
    </div>
  `;

  document.body.insertAdjacentHTML('beforeend', modalHtml);

  // Bindings del modal
  const modal = document.getElementById('modal-data-exchange');
  const btnOpen = document.getElementById('btn-data-exchange-global');
  const btnClose = document.getElementById('btn-close-exchange-modal');
  const btnCancel = document.getElementById('btn-cancel-exchange');

  const tabExport = document.getElementById('tab-btn-export');
  const tabImport = document.getElementById('tab-btn-import');
  const panelExport = document.getElementById('panel-export');
  const panelImport = document.getElementById('panel-import');

  const dropzone = document.getElementById('exchange-dropzone');
  const fileInput = document.getElementById('exchange-file-input');
  const fileSelectedBadge = document.getElementById('exchange-file-selected');
  const fileNameDisplay = document.getElementById('exchange-file-name');
  const btnProcess = document.getElementById('btn-process-import');
  const tabularOptions = document.getElementById('exchange-tabular-options');
  const resultContainer = document.getElementById('exchange-import-result');

  let selectedFile = null;

  if (btnOpen) btnOpen.onclick = () => {
    modal.classList.remove('hidden');
    modal.classList.add('flex');
  };

  const closeModal = () => modal.classList.add('hidden');
  if (btnClose) btnClose.onclick = closeModal;
  if (btnCancel) btnCancel.onclick = closeModal;

  // Manejo de Pestañas
  tabExport.onclick = () => {
    tabExport.className = "px-5 py-2.5 font-bold text-sm border-b-2 border-primary text-primary transition-colors flex items-center gap-2";
    tabImport.className = "px-5 py-2.5 font-bold text-sm border-b-2 border-transparent text-outline hover:text-on-surface transition-colors flex items-center gap-2";
    panelExport.classList.remove('hidden');
    panelImport.classList.add('hidden');
  };

  tabImport.onclick = () => {
    tabImport.className = "px-5 py-2.5 font-bold text-sm border-b-2 border-primary text-primary transition-colors flex items-center gap-2";
    tabExport.className = "px-5 py-2.5 font-bold text-sm border-b-2 border-transparent text-outline hover:text-on-surface transition-colors flex items-center gap-2";
    panelImport.classList.remove('hidden');
    panelExport.classList.add('hidden');
  };

  // Drag and drop
  dropzone.onclick = () => fileInput.click();
  
  dropzone.ondragover = (e) => {
    e.preventDefault();
    dropzone.classList.add('border-primary', 'bg-primary/5');
  };

  dropzone.ondragleave = () => {
    dropzone.classList.remove('border-primary', 'bg-primary/5');
  };

  dropzone.ondrop = (e) => {
    e.preventDefault();
    dropzone.classList.remove('border-primary', 'bg-primary/5');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFileSelected(e.dataTransfer.files[0]);
    }
  };

  fileInput.onchange = () => {
    if (fileInput.files && fileInput.files.length > 0) {
      handleFileSelected(fileInput.files[0]);
    }
  };

  function handleFileSelected(file) {
    selectedFile = file;
    fileNameDisplay.textContent = `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
    fileSelectedBadge.classList.remove('hidden');
    fileSelectedBadge.classList.add('flex');
    btnProcess.disabled = false;
    resultContainer.classList.add('hidden');

    const ext = file.name.toLowerCase().split('.').pop();
    if (['xlsx', 'xls', 'csv'].includes(ext)) {
      tabularOptions.classList.remove('hidden');
    } else {
      tabularOptions.classList.add('hidden');
    }
  }

  // Procesar archivo de importación
  btnProcess.onclick = async () => {
    if (!selectedFile) return;

    btnProcess.disabled = true;
    btnProcess.innerHTML = '<span class="material-symbols-outlined animate-spin" style="font-size:18px;">progress_activity</span> Importando y Registrando...';

    const formData = new FormData();
    formData.append('file', selectedFile);

    const ext = selectedFile.name.toLowerCase().split('.').pop();
    let url = '/api/import/cait';

    if (['xlsx', 'xls', 'csv'].includes(ext)) {
      url = '/api/import/tabular';
      const targetType = document.querySelector('input[name="import-target-type"]:checked').value;
      formData.append('target_test_type', targetType);
    } else if (ext === 'caitbackup') {
      url = '/api/system/restore';
      formData.append('mode', 'merge');
    }

    try {
      const res = await fetch(url, {
        method: 'POST',
        body: formData
      });

      const data = await res.json();

      if (res.ok && data.status === 'ok') {
        resultContainer.className = "p-4 rounded-xl border border-primary/40 bg-primary/10 text-on-surface text-xs space-y-2 animate-in fade-in duration-200";
        
        let detailsHtml = '';
        if (data.details) {
          detailsHtml = `
            <div class="grid grid-cols-2 gap-2 mt-2 pt-2 border-t border-primary/20">
              <div><strong>Pacientes Registrados:</strong> ${data.details.persons_registered || 0}</div>
              <div><strong>Borrador Guardado:</strong> ${data.details.draft_filename || 'Sí'}</div>
              <div><strong>Audiometrías:</strong> ${data.details.total_audiometria || 0}</div>
              <div><strong>Espirometrías:</strong> ${data.details.total_espirometria || 0}</div>
            </div>
          `;
        } else if (data.persons_registered !== undefined) {
          detailsHtml = `
            <div class="mt-2 pt-2 border-t border-primary/20">
              <div><strong>Pacientes registrados en la base de datos:</strong> ${data.persons_registered}</div>
            </div>
          `;
        }

        resultContainer.innerHTML = `
          <div class="flex items-center gap-2 text-primary font-bold text-sm">
            <span class="material-symbols-outlined" style="font-size:20px;">check_circle</span>
            ${data.message || 'Importación exitosa'}
          </div>
          ${detailsHtml}
          <div class="pt-3">
            <button onclick="location.reload()" class="px-4 py-2 bg-primary text-on-primary rounded-lg font-bold hover:opacity-90">
              Actualizar pantalla y ver datos
            </button>
          </div>
        `;
        resultContainer.classList.remove('hidden');

        if (window.showToast) window.showToast('¡Datos importados y registrados correctamente! ✓');
        
        // Si hay función loadData en la página actual, refrescarla
        if (window.loadData) {
          window.loadData();
        }

      } else {
        resultContainer.className = "p-4 rounded-xl border border-error/40 bg-error/10 text-error text-xs space-y-1 animate-in fade-in duration-200";
        resultContainer.innerHTML = `
          <div class="flex items-center gap-2 font-bold text-sm">
            <span class="material-symbols-outlined">error</span>
            Error al importar
          </div>
          <p>${data.message || 'No se pudo procesar el archivo.'}</p>
        `;
        resultContainer.classList.remove('hidden');
      }

    } catch (e) {
      console.error(e);
      resultContainer.className = "p-4 rounded-xl border border-error/40 bg-error/10 text-error text-xs animate-in fade-in duration-200";
      resultContainer.innerHTML = `<strong>Error de conexión:</strong> ${e.message}`;
      resultContainer.classList.remove('hidden');
    } finally {
      btnProcess.disabled = false;
      btnProcess.innerHTML = '<span class="material-symbols-outlined" style="font-size:18px;">cloud_done</span> Procesar e Importar al Sistema';
    }
  };
}

// Función global de descarga de archivos de intercambio
window.downloadExchangeFile = function(endpointUrl, defaultFilename) {
  if (window.showToast) window.showToast('Generando archivo...');
  
  const iframe = document.createElement('iframe');
  iframe.style.display = 'none';
  iframe.src = endpointUrl;
  document.body.appendChild(iframe);
  setTimeout(() => {
    document.body.removeChild(iframe);
    if (window.showToast) window.showToast('Descarga iniciada ✓');
  }, 2000);
};

window.openDataExchangeModal = function() {
  const modal = document.getElementById('modal-data-exchange');
  if (modal) {
    modal.classList.remove('hidden');
    modal.classList.add('flex');
  }
};

function injectStyles() {
  const style = document.createElement('style');
  style.innerHTML = `
    #header-container { height: 64px; min-height: 64px; display: flex; background: #fcf9f8; }
    #sidebar-container { width: 256px; min-width: 256px; display: flex; background: #f6f3f2; }
    main { animation: fadeIn 0.15s ease-out; }
    @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
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

function initLogsModal() {
  if (document.getElementById('modal-logs')) return;
  const modal = document.createElement('div');
  modal.id = 'modal-logs';
  modal.className = 'hidden fixed inset-0 bg-black/70 flex items-center justify-center z-[9999] backdrop-blur-sm';
  modal.innerHTML = `
    <div class="bg-surface p-6 rounded-2xl shadow-2xl w-[90%] max-w-4xl border border-outline-variant animate-in fade-in zoom-in duration-200 flex flex-col h-[80vh] text-on-surface">
      <div class="flex items-center justify-between mb-4">
        <div>
          <h3 class="text-lg font-bold text-primary">Logs del Sistema (Modo Desarrollador)</h3>
          <p class="text-xs text-outline">Registro de eventos y errores en tiempo real.</p>
        </div>
        <div class="flex items-center gap-3">
          <label class="flex items-center gap-2 text-xs text-outline cursor-pointer select-none">
            <input type="checkbox" id="logs-auto-refresh" checked class="rounded border-outline bg-surface focus:ring-primary">
            Auto-actualizar
          </label>
          <button onclick="refreshLogs()" class="p-2 text-outline hover:text-primary hover:bg-surface-container-high rounded-lg transition-colors flex items-center justify-center" title="Actualizar">
            <span class="material-symbols-outlined" style="font-size:20px;">refresh</span>
          </button>
          <button onclick="copyLogsToClipboard()" class="p-2 text-outline hover:text-primary hover:bg-surface-container-high rounded-lg transition-colors flex items-center justify-center" title="Copiar al portapapeles">
            <span class="material-symbols-outlined" style="font-size:20px;">content_copy</span>
          </button>
        </div>
      </div>
      <div class="flex-1 min-h-0 bg-zinc-950 text-emerald-400 font-mono p-4 text-xs overflow-y-auto rounded-xl border border-zinc-800" id="logs-content-container">
        <pre id="logs-content" class="whitespace-pre-wrap break-all"></pre>
      </div>
      <div class="flex justify-end gap-3 mt-4">
        <button onclick="closeLogsModal()" class="px-6 py-2 text-xs font-bold bg-primary text-on-primary rounded-lg shadow-md hover:scale-105 transition-all">Cerrar</button>
      </div>
    </div>
  `;
  document.body.appendChild(modal);
}

let logsInterval = null;

window.closeLogsModal = () => {
  document.getElementById('modal-logs').classList.add('hidden');
  if (logsInterval) {
    clearInterval(logsInterval);
    logsInterval = null;
  }
};

window.copyLogsToClipboard = () => {
  const content = document.getElementById('logs-content').innerText;
  navigator.clipboard.writeText(content).then(() => {
    alert('Logs copiados al portapapeles');
  }).catch(err => {
    console.error('Error al copiar logs:', err);
  });
};

window.refreshLogs = async () => {
  try {
    const res = await fetch('/api/debug/logs');
    const data = await res.json();
    if (data.status === 'ok') {
      const container = document.getElementById('logs-content-container');
      const pre = document.getElementById('logs-content');
      const isScrolledToBottom = container.scrollHeight - container.clientHeight <= container.scrollTop + 60;
      pre.innerText = data.logs || 'No hay logs registrados.';
      if (isScrolledToBottom) {
        container.scrollTop = container.scrollHeight;
      }
    }
  } catch (e) {
    console.error('Error fetching logs:', e);
  }
};

window.startLogsPolling = () => {
  if (logsInterval) clearInterval(logsInterval);
  window.refreshLogs();
  logsInterval = setInterval(() => {
    const auto = document.getElementById('logs-auto-refresh');
    if (auto && auto.checked && !document.getElementById('modal-logs').classList.contains('hidden')) {
      window.refreshLogs();
    }
  }, 2000);
};

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
      
      if (window.pywebview && window.pywebview.api && window.pywebview.api.open_logs_window) {
        window.pywebview.api.open_logs_window();
      } else {
        initLogsModal();
        document.getElementById('modal-logs').classList.remove('hidden');
        window.startLogsPolling();
      }
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

window.addEventListener('hashchange', initSidebar);

async function toggleConsole() {
  initDebugModal();
  document.getElementById('modal-debug').classList.remove('hidden');
  document.getElementById('debug-pass').focus();
}
window.toggleConsole = toggleConsole;
