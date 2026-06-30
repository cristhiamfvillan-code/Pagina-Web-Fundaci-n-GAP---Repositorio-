// js/services/auth.js
import { apiPost } from './api.js';
import { State } from '../core/state.js';

/**
 * Sincroniza un usuario autenticado por Firebase (Google Auth) con el backend (Google Sheets).
 * @param {Object} firebaseUser - Objeto de usuario proveniente de Firebase Auth
 */
export async function syncFirebaseWithBackend(firebaseUser) {
    console.log("Sincronizando usuario de Firebase con el Backend...", firebaseUser);
    const result = await apiPost({
        action: 'syncFirebaseUser',
        uid: firebaseUser.uid,
        name: firebaseUser.displayName || firebaseUser.name,
        email: firebaseUser.email
    });
    
    if (result.success) {
        // Actualizar el estado global con los datos obtenidos del backend (incluye el Rol oficial)
        State.currentUser = {
            uid: result.user.uid,
            name: result.user.name,
            email: result.user.email,
            rol: result.user.rol || 'usuario'
        };
        
        // Mantener retrocompatibilidad con app.js
        if (window.AppState) {
            window.AppState.currentUser = State.currentUser;
        }
        
        return State.currentUser;
    } else {
        console.error("Error al sincronizar con el backend:", result.error);
        throw new Error(result.error);
    }
}

/**
 * Cierra la sesión en el estado moderno
 */
export function clearAuthState() {
    State.currentUser = null;
    if (window.AppState) {
        window.AppState.currentUser = null;
    }
}
