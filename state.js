// js/core/state.js
// Manejo de estado global para las nuevas implementaciones.

export const State = {
    currentUser: null, // { uid, name, email, rol }
    notifications: [],
    activities: [],
    userDonations: [],
    
    // Función para sincronizar con el AppState antiguo de app.js si es necesario
    syncLegacy() {
        if (window.AppState) {
            this.currentUser = window.AppState.currentUser;
        }
    }
};

window.ModernState = State; // Exponer globalmente si se necesita desde código legacy
