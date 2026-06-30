// js/components/notifications.js
import { State } from '../core/state.js';
import { apiGet, apiPost } from '../services/api.js';

export async function loadNotifications() {
    if (!State.currentUser) return;
    
    try {
        const notifs = await apiGet('Notificaciones');
        // Filtrar notificaciones para el usuario actual que no han sido leidas
        const myNotifs = notifs.filter(row => row[1].toLowerCase() === State.currentUser.email.toLowerCase() && row[3] !== true && row[3] !== "TRUE");
        
        State.notifications = myNotifs;
        
        const badge = document.getElementById('notificationBadge');
        if (myNotifs.length > 0) {
            badge.textContent = myNotifs.length;
            badge.style.display = 'block';
        } else {
            badge.style.display = 'none';
        }
        
        const list = document.getElementById('notificationsList');
        list.innerHTML = '';
        
        if (myNotifs.length === 0) {
            list.innerHTML = '<div style="padding: 16px; text-align: center; color: var(--txt3); font-size: 0.9rem;">No hay notificaciones</div>';
        } else {
            myNotifs.forEach(row => {
                const item = document.createElement('div');
                item.style = 'padding: 12px 16px; border-bottom: 1px solid var(--border); font-size: 0.9rem; cursor: pointer; transition: 0.2s;';
                item.onmouseover = () => item.style.backgroundColor = 'var(--bg2)';
                item.onmouseout = () => item.style.backgroundColor = 'transparent';
                
                item.innerHTML = `
                    <div style="font-weight: 500; color: var(--txt1); margin-bottom: 4px;">${row[2]}</div>
                    <div style="font-size: 0.8rem; color: var(--txt3);">${new Date(row[4]).toLocaleString()}</div>
                `;
                
                item.onclick = async () => {
                    await apiPost({ action: 'marcarNotificacionLeida', notificacionID: row[0] });
                    loadNotifications(); // recargar
                    if (row[6]) {
                        // Si hay enlace
                        window.navigateTo(row[6]);
                    }
                };
                
                list.appendChild(item);
            });
        }
    } catch(e) {
        console.error("Error al cargar notificaciones:", e);
    }
}

export function toggleNotifications() {
    const panel = document.getElementById('notificationsPanel');
    if (panel.classList.contains('hidden')) {
        panel.classList.remove('hidden');
        loadNotifications();
    } else {
        panel.classList.add('hidden');
    }
}

// Cierra panel si se hace click fuera
document.addEventListener('click', (e) => {
    const btn = document.getElementById('topBarNotificationsBtn');
    const panel = document.getElementById('notificationsPanel');
    if (btn && panel && !btn.contains(e.target) && !panel.contains(e.target)) {
        panel.classList.add('hidden');
    }
});
