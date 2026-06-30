// js/views/adminUsers.js
import { State } from '../core/state.js';
import { apiGet, apiPost } from '../services/api.js';

let usersData = [];

export async function loadUsersAdmin() {
    if (!window.AppState || !window.AppState.currentUser || !window.AppState.isAdmin) {
        return;
    }
    
    try {
        const tbody = document.getElementById('adminUsersTableBody');
        if (tbody) tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;">Cargando usuarios...</td></tr>';
        
        const data = await apiGet('Usuarios');
        usersData = data;
        renderUsersTable();
    } catch (error) {
        console.error("Error al cargar usuarios", error);
        if (window.showToast) window.showToast('Error al cargar la lista de usuarios', 'error');
    }
}

function renderUsersTable() {
    const tbody = document.getElementById('adminUsersTableBody');
    if (!tbody) return;
    
    tbody.innerHTML = '';
    
    if (!usersData || usersData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;">No hay usuarios registrados.</td></tr>';
        return;
    }
    
    // Columnas devueltas por GET Usuarios: [UserID, Nombre, Email, Rol, FechaRegistro]
    usersData.forEach((user, index) => {
        const tr = document.createElement('tr');
        
        const uid = user[0] || '';
        const name = user[1] || 'Sin Nombre';
        const email = user[2] || '';
        const rol = user[3] || 'usuario';
        const date = user[4] || '';
        
        // Bloquear edicion/eliminacion si es el propio admin principal para evitar bloqueos
        const isSelf = (window.AppState && window.AppState.currentUser && window.AppState.currentUser.email === email);
        const disabledAttr = isSelf ? 'disabled title="No puedes modificar tu propio rol"' : '';
        
        tr.innerHTML = `
            <td>
                <div style="font-weight: 500;">${window.escapeHtml ? window.escapeHtml(name) : name}</div>
                <div style="font-size: 0.8rem; color: var(--txt3);">${uid.substring(0, 8)}...</div>
            </td>
            <td>${window.escapeHtml ? window.escapeHtml(email) : email}</td>
            <td>${date}</td>
            <td>
                <select class="form-input" style="padding: 4px; width: 100%; max-width: 120px;" 
                        onchange="window.ModernApp.adminUsers.changeUserRole('${email}', this.value)" ${disabledAttr}>
                    <option value="usuario" ${rol === 'usuario' ? 'selected' : ''}>Usuario</option>
                    <option value="voluntario" ${rol === 'voluntario' ? 'selected' : ''}>Voluntario</option>
                    <option value="lider" ${rol === 'lider' ? 'selected' : ''}>Líder</option>
                    <option value="admin" ${rol === 'admin' ? 'selected' : ''}>Admin</option>
                </select>
            </td>
            <td>
                <button class="mod-btn delete" onclick="window.ModernApp.adminUsers.deleteUser(${index}, '${email}')" ${disabledAttr}>
                    Eliminar
                </button>
            </td>
        `;
        
        tbody.appendChild(tr);
    });
}

export async function changeUserRole(email, newRole) {
    var modal = document.getElementById('deleteConfirmModal');
    if (!modal) {
        if (confirm(`¿Estás seguro de cambiar el rol de ${email} a ${newRole}?`)) {
            // Fallback (should not happen)
        }
        return;
    }
    
    document.getElementById('deleteStateConfirm').classList.remove('hidden');
    document.getElementById('deleteStateLoading').classList.add('hidden');
    document.getElementById('deleteStateSuccess').classList.add('hidden');
    
    document.getElementById('deleteConfirmTitle').textContent = '¿Cambiar Rol?';
    
    var btnConfirm = document.getElementById('btnConfirmDelete');
    btnConfirm.textContent = 'Sí, cambiar';
    
    var newBtnConfirm = btnConfirm.cloneNode(true);
    btnConfirm.parentNode.replaceChild(newBtnConfirm, btnConfirm);
    
    var btnCancel = newBtnConfirm.nextElementSibling;
    var newBtnCancel = btnCancel.cloneNode(true);
    btnCancel.parentNode.replaceChild(newBtnCancel, btnCancel);
    
    newBtnConfirm.onclick = async function() {
        document.getElementById('deleteStateConfirm').classList.add('hidden');
        document.getElementById('deleteStateLoading').classList.remove('hidden');
        try {
            const res = await apiPost({ action: 'setRole', email: email, rol: newRole });
            if (res.success) {
                document.getElementById('deleteStateLoading').classList.add('hidden');
                document.getElementById('deleteStateSuccess').classList.remove('hidden');
                document.getElementById('deleteSuccessDesc').textContent = 'El rol se actualizó correctamente.';
                setTimeout(() => {
                    if (window.closeDeleteModal) window.closeDeleteModal();
                    loadUsersAdmin();
                }, 1500);
            } else {
                alert('Error: ' + res.error);
                if (window.closeDeleteModal) window.closeDeleteModal();
                loadUsersAdmin();
            }
        } catch (e) {
            console.error(e);
            alert('Error de red');
            if (window.closeDeleteModal) window.closeDeleteModal();
            loadUsersAdmin();
        }
    };
    
    newBtnCancel.onclick = function() {
        if (window.closeDeleteModal) window.closeDeleteModal();
        loadUsersAdmin(); // revert visuals
    };
    
    modal.classList.remove('hidden');
    document.body.style.overflow = 'hidden';
    if (window.lucide) window.lucide.createIcons();
}

export async function deleteUser(rowIndex, email) {
    var modal = document.getElementById('deleteConfirmModal');
    if (!modal) {
        if (confirm(`¿ESTÁS COMPLETAMENTE SEGURO de eliminar al usuario ${email}?`)) {
            // fallback
        }
        return;
    }
    
    document.getElementById('deleteStateConfirm').classList.remove('hidden');
    document.getElementById('deleteStateLoading').classList.add('hidden');
    document.getElementById('deleteStateSuccess').classList.add('hidden');
    
    document.getElementById('deleteConfirmTitle').textContent = '¿Eliminar Usuario?';
    
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
            const res = await apiPost({ action: 'delete', sheet: 'Usuarios', rowIndex: rowIndex });
            if (res.success) {
                document.getElementById('deleteStateLoading').classList.add('hidden');
                document.getElementById('deleteStateSuccess').classList.remove('hidden');
                document.getElementById('deleteSuccessDesc').textContent = 'El usuario se eliminó correctamente.';
                setTimeout(() => {
                    if (window.closeDeleteModal) window.closeDeleteModal();
                    loadUsersAdmin();
                }, 1500);
            } else {
                alert('Error: ' + res.error);
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
