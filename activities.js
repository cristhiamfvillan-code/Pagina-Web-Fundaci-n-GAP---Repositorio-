// js/views/activities.js
import { State } from '../core/state.js';
import { apiGet, apiPost } from '../services/api.js';

export async function loadActivities() {
    const user = window.AppState ? window.AppState.currentUser : null;
    if (!user) return;
    
    try {
        const data = await apiGet('Actividades');
        State.activities = data;
        
        const container = document.getElementById('listaActividadesVoluntarios');
        if (!container) return;
        container.innerHTML = '';
        
        if (data.length === 0) {
            container.innerHTML = '<div style="grid-column: 1/-1; text-align: center; color: var(--txt3);">No hay actividades programadas.</div>';
            return;
        }
        
        data.forEach((row, index) => {
            const card = document.createElement('div');
            card.className = 'card';
            card.style.position = 'relative';
            card.style.display = 'flex';
            card.style.flexDirection = 'column';
            card.style.height = '100%';
            
            const isAdmin = window.AppState && (window.AppState.isAdmin || window.AppState.currentUser?.rol === 'lider' || window.AppState.currentUser?.rol === 'admin');
            
            let deleteBtn = '';
            if (isAdmin) {
                deleteBtn = `<button class="btn btn-sm btn-ghost" onclick="window.ModernApp.activities.deleteActividad(${index})" style="color:var(--danger); background:rgba(239,68,68,0.1); border-radius:var(--rf); padding:8px; display:flex; align-items:center; justify-content:center; flex-shrink:0; transition:all 0.2s;" onmouseover="this.style.background='rgba(239,68,68,0.2)'" onmouseout="this.style.background='rgba(239,68,68,0.1)'" title="Eliminar"><i data-lucide="trash-2" width="16" height="16"></i></button>`;
            }
            
            // Lógica de fechas
            const actDate = new Date(row[3]);
            const today = new Date();
            today.setHours(0,0,0,0);
            const actDay = new Date(actDate);
            actDay.setHours(0,0,0,0);
            const diffTime = actDay - today;
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

            let timeBadgeHtml = '';
            if (diffDays > 1) {
                timeBadgeHtml = `<span style="display:inline-flex; align-items:center; gap:4px; padding:4px 10px; background:rgba(249, 115, 22, 0.12); color:var(--a600); border-radius:var(--rf); font-size:0.75rem; font-weight:800;"><i data-lucide="clock" width="14" height="14"></i> Faltan ${diffDays} días</span>`;
            } else if (diffDays === 1) {
                timeBadgeHtml = `<span style="display:inline-flex; align-items:center; gap:4px; padding:4px 10px; background:rgba(249, 115, 22, 0.12); color:var(--a600); border-radius:var(--rf); font-size:0.75rem; font-weight:800;"><i data-lucide="clock" width="14" height="14"></i> Mañana</span>`;
            } else if (diffDays === 0) {
                timeBadgeHtml = `<span style="display:inline-flex; align-items:center; gap:4px; padding:4px 10px; background:var(--b1); color:var(--p700); border-radius:var(--rf); font-size:0.75rem; font-weight:800;"><i data-lucide="zap" width="14" height="14"></i> ¡Es hoy!</span>`;
            } else {
                const absDays = Math.abs(diffDays);
                timeBadgeHtml = `<span style="display:inline-flex; align-items:center; gap:4px; padding:4px 10px; background:rgba(0, 0, 0, 0.05); color:var(--txt3); border-radius:var(--rf); font-size:0.75rem; font-weight:700;"><i data-lucide="check-circle" width="14" height="14"></i> Finalizada hace ${absDays} día${absDays !== 1 ? 's' : ''}</span>`;
            }
            
            // Opciones de formato de fecha
            const dateOptions = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' };
            const formattedDate = actDate.toLocaleString('es-ES', dateOptions);
            const capitalizedDate = formattedDate.charAt(0).toUpperCase() + formattedDate.slice(1);
            
            // Verificar si ya asiste
            let yaConfirmado = false;
            try {
                const asistentesList = JSON.parse(row[4] || '[]');
                if (window.AppState && window.AppState.currentUser) {
                    const currentEmail = window.AppState.currentUser.email;
                    yaConfirmado = asistentesList.some(a => a.email === currentEmail);
                }
            } catch(e) {}

            let actionButton = '';
            if (diffDays < 0) {
                actionButton = `<button class="btn btn-sm" style="padding:10px 20px; border-radius:var(--rf); background:var(--bg2); color:var(--txt3); cursor:not-allowed; box-shadow:none; border:1px solid var(--b2);">Cerrada</button>`;
            } else if (yaConfirmado) {
                actionButton = `<button class="btn btn-sm" style="padding:10px 20px; border-radius:var(--rf); background:var(--b1); color:var(--p700); cursor:default; box-shadow:none; border:none; display:flex; align-items:center; gap:6px;">
                                    <i data-lucide="check-circle-2" width="16" height="16"></i> ¡Confirmado!
                                </button>`;
            } else {
                actionButton = `<button class="btn btn-sm btn-accent" style="padding:10px 20px; border-radius:var(--rf);" onclick="window.ModernApp.activities.asistir('${row[0]}')">
                                    Asistiré
                                </button>`;
            }
            
            card.innerHTML = `
                <div style="padding:24px; display:flex; flex-direction:column; height:100%; gap:20px;">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:12px;">
                        <div>
                            <div style="display:inline-block; padding:4px 10px; background:var(--b1); color:var(--p600); border-radius:var(--r2); font-size:0.7rem; font-weight:800; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:12px;">Voluntariado</div>
                            <h3 style="font-size:1.3rem; color:var(--txt1); line-height:1.3; margin-bottom:0;">${row[1]}</h3>
                        </div>
                        ${deleteBtn}
                    </div>
                    
                    <div style="display:flex; flex-direction:column; gap:10px; background:var(--bg2); padding:16px; border-radius:var(--r2); border:1px solid rgba(0,0,0,0.03);">
                        <div style="display:flex; align-items:flex-start; gap:8px; color:var(--txt2); font-size:0.9rem; font-weight:600; line-height:1.4;">
                            <i data-lucide="calendar-days" width="18" height="18" style="color:var(--p500); flex-shrink:0; margin-top:2px;"></i>
                            <span>${capitalizedDate}</span>
                        </div>
                        <div style="margin-top:4px;">${timeBadgeHtml}</div>
                    </div>

                    <p style="color:var(--txt2); font-size:0.95rem; line-height:1.6; flex-grow:1; margin-bottom:0;">${row[2]}</p>

                    <div style="display:flex; justify-content:space-between; align-items:center; margin-top:8px; padding-top:20px; border-top:1px solid rgba(0,0,0,0.06);">
                        <span class="badge ${row[5] === 'Activa' ? 'success' : 'primary'}" style="font-size:0.75rem; padding:4px 10px;">${row[5] || 'Pendiente'}</span>
                        ${actionButton}
                    </div>
                </div>
            `;
            container.appendChild(card);
        });
        
        if (window.lucide) {
            window.lucide.createIcons();
        }
        
    } catch (e) {
        console.error("Error al cargar actividades:", e);
    }
}

export async function crearActividad(form) {
    const user = window.AppState ? window.AppState.currentUser : null;
    if (!user) return;
    
    const titulo = form.titulo.value;
    const fecha = form.fecha.value;
    const desc = form.descripcion.value;
    const actId = 'ACT_' + Date.now();
    
    const submitBtn = form.querySelector('button[type="submit"]');
    const originalText = submitBtn ? submitBtn.innerHTML : 'Publicar';
    
    if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.innerHTML = '<i data-lucide="loader-2" class="spin-anim" style="display:inline-block; vertical-align:middle;"></i> Publicando...';
        if (window.lucide) window.lucide.createIcons();
    }
    
    try {
        const result = await apiPost({
            action: 'add',
            sheet: 'Actividades',
            row: [actId, titulo, desc, fecha, '', 'Activa', user.name]
        });
        
        if (result.success) {
            if (window.showToast) window.showToast('Actividad publicada', 'success');
            else alert('Actividad publicada exitosamente');
            form.closest('.modal').classList.remove('active');
            form.reset();
            loadActivities();
        } else {
            if (window.showToast) window.showToast('Error: ' + result.error, 'error');
            else alert('Error al publicar: ' + result.error);
        }
    } catch (e) {
        if (window.showToast) window.showToast('Error de red', 'error');
        else alert('Error: ' + e.message);
    } finally {
        if (submitBtn) {
            submitBtn.disabled = false;
            submitBtn.innerHTML = originalText;
            if (window.lucide) window.lucide.createIcons();
        }
    }
}

export async function asistir(actId) {
    const user = window.AppState ? window.AppState.currentUser : null;
    if (!user) {
        if (window.showToast) window.showToast('Debes iniciar sesión para asistir', 'error');
        else alert('Debes iniciar sesión para asistir');
        return;
    }
    
    if (window.showToast) window.showToast('Registrando tu asistencia...', 'info');
    
    try {
        const res = await apiPost({ 
            action: 'asistir', 
            actividadID: actId,
            userID: user.uid,
            userName: user.name,
            userEmail: user.email 
        });
        
        if (res.success) {
            if (window.showToast) {
                window.showToast('¡Gracias por registrarte para asistir a esta actividad!', 'success');
            } else {
                alert('Te has registrado para asistir a esta actividad. ¡Gracias por tu compromiso!');
            }
        } else {
            if (window.showToast) window.showToast(res.error || 'Error al registrar', 'error');
            else alert(res.error || 'Error al registrar');
        }
    } catch (e) {
        console.error(e);
        if (window.showToast) window.showToast('Error de conexión', 'error');
        else alert('Error de red');
    }
}

export async function deleteActividad(rowIndex) {
    var modal = document.getElementById('deleteConfirmModal');
    if (!modal) {
        if (confirm('¿ESTÁS COMPLETAMENTE SEGURO de eliminar esta actividad? Esta acción no se puede deshacer.')) {
            // fallback
        }
        return;
    }
    
    document.getElementById('deleteStateConfirm').classList.remove('hidden');
    document.getElementById('deleteStateLoading').classList.add('hidden');
    document.getElementById('deleteStateSuccess').classList.add('hidden');
    
    document.getElementById('deleteConfirmTitle').textContent = '¿Eliminar Actividad?';
    
    var btnConfirm = document.getElementById('btnConfirmDelete');
    btnConfirm.textContent = 'Sí, eliminar';
    
    var newBtnConfirm = btnConfirm.cloneNode(true);
    btnConfirm.parentNode.replaceChild(newBtnConfirm, btnConfirm);
    
    var btnCancel = newBtnConfirm.nextElementSibling;
    var newBtnCancel = btnCancel.cloneNode(true);
    btnCancel.parentNode.replaceChild(newBtnCancel, btnCancel);
    
    newBtnConfirm.onclick = async function() {
        document.getElementById('deleteStateConfirm').classList.add('hidden');
        document.getElementById('deleteStateLoading').classList.remove('hidden');
        try {
            const res = await apiPost({ action: 'delete', sheet: 'Actividades', rowIndex: rowIndex });
            if (res.success) {
                document.getElementById('deleteStateLoading').classList.add('hidden');
                document.getElementById('deleteStateSuccess').classList.remove('hidden');
                document.getElementById('deleteSuccessDesc').textContent = 'La actividad se eliminó correctamente.';
                setTimeout(() => {
                    if (window.closeDeleteModal) window.closeDeleteModal();
                    loadActivities();
                }, 1500);
            } else {
                alert('Error al eliminar: ' + res.error);
                if (window.closeDeleteModal) window.closeDeleteModal();
            }
        } catch (e) {
            console.error(e);
            alert('Error de red');
            if (window.closeDeleteModal) window.closeDeleteModal();
        }
    };
    
    newBtnCancel.onclick = function() {
        if (window.closeDeleteModal) window.closeDeleteModal();
    };
    
    modal.classList.remove('hidden');
    document.body.style.overflow = 'hidden';
    if (window.lucide) window.lucide.createIcons();
}
