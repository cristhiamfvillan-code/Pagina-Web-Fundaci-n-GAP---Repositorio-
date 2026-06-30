// js/main.js
import { syncFirebaseWithBackend, clearAuthState } from './services/auth.js?v=12';
import { State } from './core/state.js?v=12';
import { apiGet, apiPost } from './services/api.js?v=12';
import { loadProfile } from './views/profile.js?v=12';
import { loadActivities, crearActividad, asistir, deleteActividad } from './views/activities.js?v=12';
import { loadNotifications, toggleNotifications } from './components/notifications.js?v=12';
import { loadForumData, submitForumPost, submitForumReply } from './views/forum.js?v=12';
import { loadUsersAdmin, changeUserRole, deleteUser } from './views/adminUsers.js?v=12';

window.ModernApp = window.ModernApp || {};
window.ModernApp.auth = { syncFirebaseWithBackend, clearAuthState };
window.ModernApp.state = State;
window.ModernApp.api = { apiGet, apiPost };
window.ModernApp.profile = { loadProfile };
window.ModernApp.activities = { loadActivities, crearActividad, asistir, deleteActividad };
window.ModernApp.notifications = { loadNotifications, toggleNotifications };
window.ModernApp.forum = Object.assign(window.ModernApp.forum || {}, {
    loadForumData,
    submitForumPost,
    submitForumReply
});
window.ModernApp.adminUsers = { loadUsersAdmin, changeUserRole, deleteUser };

// Globalizar toggleNotifications para el HTML onclick
window.toggleNotifications = toggleNotifications;

console.log("ModernApp initialization complete.");
