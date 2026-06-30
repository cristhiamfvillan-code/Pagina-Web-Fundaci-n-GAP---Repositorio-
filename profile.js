// js/views/profile.js
import { State } from '../core/state.js';
import { apiGet } from '../services/api.js';

export async function loadProfile() {
    const user = window.AppState ? window.AppState.currentUser : null;
    if (!user) return;
    
    // Set UI basic data
    document.getElementById('miRolTexto').textContent = (user.rol || 'Usuario').toUpperCase();
    
    try {
        const donaciones = await apiGet('Donaciones');
        const misDonaciones = donaciones.filter(row => row[6] === user.uid || row[1].toLowerCase() === user.email.toLowerCase());
        
        let total = 0;
        const tbody = document.getElementById('miHistorialDonaciones');
        tbody.innerHTML = '';
        
        if (misDonaciones.length === 0) {
            tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;">No has realizado donaciones aún.</td></tr>';
        } else {
            misDonaciones.forEach(row => {
                const monto = parseFloat(row[2]) || 0;
                total += monto;
                
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${new Date(row[4]).toLocaleDateString()}</td>
                    <td>$${monto.toLocaleString()}</td>
                    <td>${row[3]}</td>
                    <td><span class="badge ${row[7] === 'Aprobada' ? 'success' : 'warning'}">${row[7] || 'Pendiente'}</span></td>
                `;
                tbody.appendChild(tr);
            });
        }
        
        document.getElementById('miTotalDonado').textContent = `$${total.toLocaleString()}`;
    } catch (e) {
        console.error("Error al cargar perfil:", e);
    }
}
