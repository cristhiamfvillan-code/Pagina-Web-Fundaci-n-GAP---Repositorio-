
    // =====================================================================
    // CONFIGURACION
    // =====================================================================

    const GOOGLE_SHEETS_API_URL = 'https://script.google.com/macros/s/AKfycbx1AYeZGDxjF5LGKKFSC7EXmr0CwEpUQuT7gnEIIQF4MIIj9jhmoZUg15tjKd2vsfSJOg/exec';

    // Client ID de Google Cloud Console → APIs & Services → Credentials
    // Formato: "000000000000-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx.apps.googleusercontent.com"
    const GOOGLE_CLIENT_ID = 'TU_CLIENT_ID.apps.googleusercontent.com';

    const FIREBASE_CONFIG = {
        apiKey: "TU_API_KEY",
        authDomain: "TU_PROYECTO.firebaseapp.com",
        projectId: "TU_PROYECTO",
        storageBucket: "TU_PROYECTO.appspot.com",
        messagingSenderId: "TU_SENDER_ID",
        appId: "TU_APP_ID"
    };

    const MSAL_CONFIG = {
        auth: {
            clientId: "TU_CLIENT_ID_DE_AZURE",
            authority: "https://login.microsoftonline.com/common",
            redirectUri: window.location.origin
        },
        cache: {
            cacheLocation: "localStorage",
            storeAuthStateInCookie: false
        }
    };

    // =====================================================================
    // ESTADO GLOBAL DE LA APLICACION
    // =====================================================================
    const AppState = {
        currentUser: null,
        isAdmin: false,
        currentPage: 'inicio',
        feedData: [],
        forumData: [],
        donationsData: [],
        confirmationResult: null,
        currentModalPostIndex: -1,
        currentModalImageIndex: 0,
        pendingImageFiles: []
    };
    window.AppState = AppState; // EXPUESTO GLOBALMENTE PARA LOS NUEVOS MÓDULOS

    /**
     * Lista de correos de administradores.
     * Cuando un usuario inicia sesion con uno de estos correos,
     * se le otorga rol de Administrador automaticamente.
     * PERSONALIZA esta lista con los correos de tus admins.
     */
    const ADMIN_EMAILS = [
        'admin@fundacion.org',
        'tu-correo-admin@gmail.com'
    ];

    // =====================================================================
    // CROWDFUNDING
    // =====================================================================
    async function loadCrowdfundingProgress() {
        var response = await fetchFromSheets('GET', 'Donaciones');
        var total = 0;
        if (response && Array.isArray(response.data)) {
            for (var i = 1; i < response.data.length; i++) {
                var row = response.data[i];
                if (row[7] === 'Confirmada') { // Asumiendo que el Estado esta en la col 7
                    total += parseInt(row[2]) || 0;
                }
            }
        }
        
        var goal = 5000000;
        var percentage = Math.min((total / goal) * 100, 100);
        
        var progressEl = document.getElementById('crowdfundProgress');
        var amountEl = document.getElementById('crowdfundAmount');
        if(progressEl) progressEl.style.width = percentage + '%';
        if(amountEl) amountEl.textContent = '$' + total.toLocaleString('es-CO');
    }

    // =====================================================================
    // INICIALIZACION
    // =====================================================================
    let firebaseApp, firebaseAuth, googleProvider, msalInstance;

    /** Inicializa Firebase Auth */
    function initFirebase() {
        try {
            firebaseApp = firebase.initializeApp(FIREBASE_CONFIG);
            firebaseAuth = firebase.auth();
            googleProvider = new firebase.auth.GoogleAuthProvider();

            firebaseAuth.onAuthStateChanged(function(user) {
                if (user) {
                    handleUserLogin({
                        uid: user.uid,
                        name: user.displayName || 'Usuario',
                        email: user.email || '',
                        photo: user.photoURL || null,
                        provider: user.providerData[0] ? user.providerData[0].providerId : 'unknown'
                    });
                }
            });
            console.log('Firebase inicializado correctamente');
        } catch (error) {
            console.warn('Firebase no configurado. Usando modo demo.', error.message);
        }
    }

    /** Inicializa MSAL para autenticacion con Microsoft */
    function initMSAL() {
        try {
            if (typeof msal !== 'undefined') {
                msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);
                console.log('MSAL inicializado correctamente');
            }
        } catch (error) {
            console.warn('MSAL no configurado. Usando modo demo.', error.message);
        }
    }

    /** Inicializa Google Identity Services (GIS).
     *  Se llama automáticamente con onload del script GIS y también desde DOMContentLoaded. */
    function initGoogleSignIn() {
        if (typeof google === 'undefined' || !google.accounts || !google.accounts.id) return;
        if (!GOOGLE_CLIENT_ID || GOOGLE_CLIENT_ID.includes('TU_CLIENT_ID')) return;

        google.accounts.id.initialize({
            client_id: GOOGLE_CLIENT_ID,
            callback: handleGoogleCredential,
            auto_select: false,
            cancel_on_tap_outside: true,
            use_fedcm_for_prompt: true
        });
        console.log('Google Identity Services inicializado');
    }

    /** Decodifica el JWT que devuelve GIS y loga al usuario */
    function handleGoogleCredential(response) {
        try {
            // El credential es un JWT firmado; el payload está en la parte central (base64url)
            var b64 = response.credential.split('.')[1]
                .replace(/-/g, '+').replace(/_/g, '/');
            // Relleno de padding si es necesario
            while (b64.length % 4) b64 += '=';
            var payload = JSON.parse(atob(b64));

            handleUserLogin({
                uid:      payload.sub,
                name:     payload.name || payload.given_name || 'Usuario',
                email:    payload.email || '',
                photo:    payload.picture || null,
                provider: 'google'
            });
        } catch (err) {
            console.error('Error al decodificar credencial de Google:', err);
            showToast('Error al procesar la sesión de Google', 'error');
        }
    }

    /** Quita el loader de primera carga */
    function initLoader() {
        var loader = document.getElementById('page-loader');
        if (!loader) return;
        // Mínimo 2 s para que la animación sea visible; se espera también el load completo
        var minTime = new Promise(function(r) { setTimeout(r, 2000); });
        var loaded  = new Promise(function(r) { window.addEventListener('load', r); });
        Promise.all([minTime, loaded]).then(function() {
            loader.classList.add('loaded');
            setTimeout(function() { loader.remove(); }, 700);
        });
    }

    /** Scroll reveal con IntersectionObserver */
    function initScrollReveal() {
        var observer = new IntersectionObserver(function(entries) {
            entries.forEach(function(entry) {
                if (entry.isIntersecting) {
                    entry.target.classList.add('visible');
                    observer.unobserve(entry.target);
                }
            });
        }, { threshold: 0.12 });

        document.querySelectorAll('.reveal').forEach(function(el) {
            observer.observe(el);
        });
    }

    /** Re-aplica reveal cuando se navega a una sección */
    function revealSection(sectionId) {
        var section = document.getElementById('page-' + sectionId);
        if (!section) return;
        setTimeout(function() {
            section.querySelectorAll('.reveal:not(.visible)').forEach(function(el) {
                el.classList.add('visible');
            });
        }, 80);
    }

    /** Inicializacion principal al cargar la pagina */
    document.addEventListener('DOMContentLoaded', function() {
        if (typeof lucide !== 'undefined') {
            lucide.createIcons();
        }

        initLoader();
        initScrollReveal();
        initFirebase();
        initMSAL();
        initGoogleSignIn();

        var dateInput = document.getElementById('postDate');
        if (dateInput) {
            dateInput.value = new Date().toISOString().split('T')[0];
        }

        loadFeedData().then(function() {
            // Verificar si hay parametro URL ?post=X
            var urlParams = new URLSearchParams(window.location.search);
            var postId = urlParams.get('post');
            if (postId !== null) {
                var index = parseInt(postId);
                if (!isNaN(index) && index >= 0 && index < AppState.feedData.length) {
                    navigateTo('labores');
                    openPostModal(index);
                }
            }
        });

        loadBankInfo(); // <-- NUEVO: Cargar informacion bancaria dinámica
        loadCrowdfundingProgress();
        registerVisit();
        checkLocalSession();

        console.log('Fundacion GAP - Aplicacion cargada');
    });

    // =====================================================================
    // NAVEGACION
    // =====================================================================

    /** Navega a una seccion/pagina especifica */
    function navigateTo(pageId) {
        document.querySelectorAll('.page-section').forEach(function(section) {
            section.classList.remove('active');
        });

        var targetSection = document.getElementById('page-' + pageId);
        if (targetSection) {
            targetSection.classList.add('active');
        }

        document.querySelectorAll('.nav-item').forEach(function(item) {
            item.classList.remove('active');
        });
        var activeNavItem = document.querySelector('.nav-item[data-page="' + pageId + '"]');
        if (activeNavItem) {
            activeNavItem.classList.add('active');
        }

        var pageTitles = {
            'inicio': 'Inicio',
            'quienes-somos': 'Quienes Somos',
            'labores': 'Nuestras Labores',
            'donaciones': 'Donaciones',
            'foro': 'Comunidad',
            'voluntariado': 'Voluntariado',
            'mi-cuenta': 'Mi Cuenta',
            'actividades': 'Actividades',
            'admin-dashboard': 'Dashboard',
            'admin-donantes': 'Gestion de Donantes',
            'admin-contenido': 'Gestion de Contenido',
            'admin-voluntarios': 'Gestion de Voluntarios',
            'admin-usuarios': 'Gestion de Usuarios'
        };
        document.getElementById('pageTitle').textContent = pageTitles[pageId] || 'Inicio';

        AppState.currentPage = pageId;

        if (pageId === 'admin-dashboard' && AppState.isAdmin) loadAdminDashboard();
        if (pageId === 'admin-donantes' && AppState.isAdmin) loadDonorsTable();
        if (pageId === 'admin-contenido' && AppState.isAdmin) loadForumPostsAdmin();
        if (pageId === 'admin-voluntarios' && AppState.isAdmin) loadVolunteers();
        
        // Hooks modernos
        if (pageId === 'foro' && window.ModernApp) window.ModernApp.forum.loadForumData();
        if (pageId === 'mi-cuenta' && window.ModernApp) window.ModernApp.profile.loadProfile();
        if (pageId === 'actividades' && window.ModernApp) window.ModernApp.activities.loadActivities();
        if (pageId === 'admin-usuarios' && window.ModernApp) window.ModernApp.adminUsers.loadUsersAdmin();

        closeSidebar();
        window.scrollTo(0, 0);
        revealSection(pageId);

        setTimeout(function() {
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }, 100);
    }

    function toggleSidebar() {
        var sidebar = document.getElementById('sidebar');
        var overlay = document.getElementById('sidebarOverlay');
        sidebar.classList.toggle('open');
        overlay.classList.toggle('show');
    }

    function closeSidebar() {
        var sidebar = document.getElementById('sidebar');
        var overlay = document.getElementById('sidebarOverlay');
        sidebar.classList.remove('open');
        overlay.classList.remove('show');
    }

    function closeShareMenu() {
        document.getElementById('shareOverlay').classList.remove('active');
        document.getElementById('shareMenu').classList.remove('active');
    }

    window.openHelpModal = function() {
        document.getElementById('helpModal').classList.add('show');
    }
    
    window.closeHelpModal = function() {
        document.getElementById('helpModal').classList.remove('show');
    }

    // =====================================================================
    // AUTENTICACION
    // =====================================================================

    function openLoginModal() {
        document.getElementById('loginModal').classList.add('show');
        setTimeout(function() {
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }, 100);
    }

    function closeLoginModal() {
        document.getElementById('loginModal').classList.remove('show');
    }

    window.showAuthLoading = function(titleText) {
        let modal = document.getElementById('authLoadingModal');
        if (!modal) {
            modal = document.createElement('div');
            modal.id = 'authLoadingModal';
            modal.className = 'post-viewer-overlay hidden';
            modal.style.cssText = 'z-index: 10005; align-items: center; justify-content: center; display: flex;';
            modal.innerHTML = `
                <div class="delete-modal-content" style="background: var(--bg1); border-radius: var(--r4); padding: 40px 32px; width: 90%; max-width: 450px; text-align: center; box-shadow: var(--sh4); transform: scale(1) !important; opacity: 1 !important;">
                    <div class="delete-icon loading" style="margin-bottom: 20px; color: var(--primary-500);">
                        <i data-lucide="loader-2" class="spin-anim" width="48" height="48"></i>
                    </div>
                    <h2 id="authLoadingTitle" style="color: var(--txt1); font-size: 1.5rem; margin-bottom: 8px;"></h2>
                    <p style="color: var(--txt2);">Por favor, espera un momento.</p>
                </div>
            `;
            document.body.appendChild(modal);
        }
        document.getElementById('authLoadingTitle').textContent = titleText;
        setTimeout(function() {
            modal.classList.remove('hidden');
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }, 10);
    };

    window.hideAuthLoading = function() {
        var modal = document.getElementById('authLoadingModal');
        if (modal) {
            setTimeout(function() {
                modal.classList.add('hidden');
            }, 50);
        }
    };

    /** Login con Google — GIS (prioritario) con Firebase como respaldo */
    async function loginWithGoogle() {
        var gisReady = typeof google !== 'undefined'
            && google.accounts && google.accounts.id
            && GOOGLE_CLIENT_ID && !GOOGLE_CLIENT_ID.includes('TU_CLIENT_ID');

        if (gisReady) {
            google.accounts.id.prompt(function(notification) {
                if (notification.isNotDisplayed() || notification.isSkippedMoment()) {
                    // One Tap bloqueado por el navegador o sin sesión Google activa → Firebase
                    loginWithGoogleFirebase();
                }
                // Si se muestra correctamente, handleGoogleCredential() se ejecuta solo
            });
            return;
        }

        // GIS no configurado: usar Firebase
        await loginWithGoogleFirebase();
    }

    /** Login con Google vía Firebase (respaldo cuando GIS no está configurado) */
    async function loginWithGoogleFirebase() {
        try {
            if (firebaseAuth && googleProvider) {
                var result = await firebaseAuth.signInWithPopup(googleProvider);
                var user = result.user;
                
                showToast('Autenticado con Google, sincronizando...', 'info');
                window.showAuthLoading('Iniciando sesión...');
                
                // USAR ARQUITECTURA MODERNA
                if (window.ModernApp && window.ModernApp.auth) {
                    var modernUser = await window.ModernApp.auth.syncFirebaseWithBackend(user);
                    handleUserLogin({
                        uid:   modernUser.uid,
                        name:  modernUser.name,
                        email: modernUser.email,
                        photo: user.photoURL,
                        rol:   modernUser.rol,
                        provider: 'google'
                    });
                    window.hideAuthLoading();
                } else {
                    // Fallback
                    handleUserLogin({
                        uid:   user.uid,
                        name:  user.displayName,
                        email: user.email,
                        photo: user.photoURL,
                        provider: 'google'
                    });
                    window.hideAuthLoading();
                }
            } else {
                showToast('Google Sign-In no está configurado. Usa correo y contraseña.', 'error');
            }
        } catch (error) {
            console.error('Error en login con Google:', error);
            if (error.code === 'auth/popup-closed-by-user') {
                showToast('Login cancelado', 'info');
            } else {
                window.hideAuthLoading();
                showToast('Error al iniciar sesión con Google.', 'error');
            }
        }
    }

    /** Login con Microsoft/Outlook usando MSAL */
    async function loginWithMicrosoft() {
        try {
            if (msalInstance) {
                var loginResponse = await msalInstance.loginPopup({
                    scopes: ["openid", "profile", "email", "User.Read"]
                });
                if (loginResponse.account) {
                    handleUserLogin({
                        uid: loginResponse.account.localAccountId,
                        name: loginResponse.account.name,
                        email: loginResponse.account.username,
                        photo: null,
                        provider: 'microsoft'
                    });
                }
            } else {
                showToast('Microsoft Auth no está configurado. Contacta al administrador.', 'error');
            }
        } catch (error) {
            console.error('Error en login con Microsoft:', error);
            showToast('Error al iniciar sesión con Microsoft.', 'error');
        }
    }

    /** Login con telefono - Paso 1: Enviar SMS */
    async function loginWithPhone() {
        var countryCode = document.getElementById('countryCode').value.trim();
        var phoneNumber = document.getElementById('phoneNumber').value.trim();

        if (!phoneNumber) {
            showToast('Ingresa tu numero de telefono', 'error');
            return;
        }

        var fullNumber = countryCode + phoneNumber.replace(/\s/g, '');

        try {
            if (firebaseAuth) {
                if (!window.recaptchaVerifier) {
                    window.recaptchaVerifier = new firebase.auth.RecaptchaVerifier('recaptcha-container', {
                        size: 'normal',
                        callback: function() { console.log('reCAPTCHA resuelto'); }
                    });
                }

                var confirmationResult = await firebaseAuth.signInWithPhoneNumber(fullNumber, window.recaptchaVerifier);
                AppState.confirmationResult = confirmationResult;
                document.getElementById('otpSection').classList.remove('hidden');
                showToast('Codigo enviado a tu telefono', 'success');
            } else {
                showToast('Firebase no está configurado. Usa correo y contraseña.', 'error');
            }
        } catch (error) {
            console.error('Error enviando SMS:', error);
            showToast('Error al enviar el SMS. Verifica el número.', 'error');
        }
    }

    /** Verifica el codigo OTP */
    async function verifyOtp() {
        var otpCode = document.getElementById('otpCode').value.trim();
        if (!otpCode || otpCode.length !== 6) {
            showToast('Ingresa un codigo de 6 digitos', 'error');
            return;
        }
        try {
            if (AppState.confirmationResult) {
                var result = await AppState.confirmationResult.confirm(otpCode);
                var user = result.user;
                handleUserLogin({
                    uid: user.uid,
                    name: user.phoneNumber || 'Usuario',
                    email: '',
                    photo: null,
                    provider: 'phone'
                });
            }
        } catch (error) {
            showToast('Codigo incorrecto. Intenta de nuevo.', 'error');
        }
    }

    /** Login con correo/contraseña contra el backend de Google Sheets */
    async function loginWithCredentials() {
        var email = document.getElementById('usernameInput').value.trim();
        var pass = document.getElementById('passwordInput').value.trim();

        if (!email || !pass) {
            showToast('Ingresa tu correo y contraseña', 'error');
            return;
        }

        window.showAuthLoading('Iniciando sesión...');

        var result = await fetchFromSheets('POST', 'Usuarios', {
            action: 'loginUser',
            email: email,
            password: pass
        });

        if (result && result.success && result.user) {
            handleUserLogin({
                uid: result.user.uid,
                name: result.user.name,
                email: result.user.email,
                photo: null,
                provider: 'credentials',
                rol: result.user.rol
            });
        } else {
            var msg = (result && result.error === 'INVALID_CREDENTIALS')
                ? 'Correo o contraseña incorrectos'
                : 'Error al iniciar sesión. Intenta de nuevo.';
            window.hideAuthLoading();
            showToast(msg, 'error');
        }
    }

    /** Registrar una cuenta nueva con correo y contraseña */
    async function registerWithCredentials() {
        var name  = document.getElementById('registerName').value.trim();
        var email = document.getElementById('registerEmail').value.trim();
        var pass  = document.getElementById('registerPassword').value;
        var pass2 = document.getElementById('registerPasswordConfirm').value;

        if (!name || !email || !pass) {
            showToast('Completa todos los campos', 'error');
            return;
        }
        if (pass.length < 6) {
            showToast('La contraseña debe tener al menos 6 caracteres', 'error');
            return;
        }
        if (pass !== pass2) {
            showToast('Las contraseñas no coinciden', 'error');
            return;
        }

        window.showAuthLoading('Creando cuenta...');

        var result = await fetchFromSheets('POST', 'Usuarios', {
            action: 'registerUser',
            nombre: name,
            email: email,
            password: pass
        });

        if (result && result.success && result.user) {
            showToast('¡Cuenta creada exitosamente! Iniciando sesión...', 'success');
            handleUserLogin({
                uid: result.user.uid,
                name: result.user.name,
                email: result.user.email,
                photo: null,
                provider: 'credentials',
                rol: result.user.rol
            });
        } else {
            var msg = 'Error al crear la cuenta.';
            if (result && result.error === 'EMAIL_EXISTS') {
                msg = 'Ese correo ya está registrado. Intenta iniciar sesión.';
            } else if (result && result.error === 'EMAIL_Y_CONTRASENA_REQUERIDOS') {
                msg = 'Ingresa un correo y contraseña válidos.';
            } else if (result && result.error) {
                msg = 'Error: ' + result.error;
            }
            window.hideAuthLoading();
            showToast(msg, 'error');
        }
    }

    /** Alterna entre el formulario de Login y el de Registro */
    function toggleAuthForm(mode) {
        var loginSection    = document.getElementById('loginFormSection');
        var registerSection = document.getElementById('registerFormSection');

        if (mode === 'register') {
            loginSection.style.display    = 'none';
            registerSection.style.display = 'block';
        } else {
            loginSection.style.display    = 'block';
            registerSection.style.display = 'none';
        }

        setTimeout(function() {
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }, 100);
    }

    /** Maneja el estado despues de un login exitoso */
    function handleUserLogin(userData) {
        AppState.currentUser = userData;
        AppState.isAdmin = ADMIN_EMAILS.includes(userData.email.toLowerCase()) || userData.rol === 'admin';
        
        // Si el rol ya viene (por ejemplo de ModernApp), no necesitamos checkUserRole,
        // pero igual lo llamamos si no está definido para credenciales locales.
        if (userData.rol) {
            updateUIForLoggedUser();
        } else {
            checkUserRole(userData);
            updateUIForLoggedUser();
        }
        
        closeLoginModal();
        localStorage.setItem('fundacion_session', JSON.stringify(userData));
        
        // Ya no registramos aquí si viene de Google porque ModernApp lo hace
        if (userData.provider !== 'credentials' && !window.ModernApp) {
            registerUserInSheet(userData);
        }
        if (window.hideAuthLoading) window.hideAuthLoading();
        showToast('Bienvenido, ' + userData.name + '! &#127793;', 'success');
    }

    /** Verifica el rol del usuario consultando la hoja de Usuarios */
    async function checkUserRole(userData) {
        try {
            var response = await fetchFromSheets('GET', 'Usuarios');
            if (response && response.data) {
                var userRow = response.data.find(function(row) {
                    return row[0] === userData.uid || row[2] === userData.email;
                });
                if (userRow && userRow[3] === 'admin') {
                    AppState.isAdmin = true;
                }
            }
        } catch (error) {
            console.log('No se pudo verificar rol desde Sheets');
        }
        if (AppState.isAdmin || userData.rol === 'admin') {
            document.body.classList.add('is-admin', 'is-lider', 'is-voluntario');
            document.querySelectorAll('.admin-only').forEach(function(el) { el.classList.remove('hidden'); });
        } else if (userData.rol === 'lider') {
            document.body.classList.add('is-lider', 'is-voluntario');
            document.querySelectorAll('.admin-only').forEach(function(el) { el.classList.add('hidden'); });
        } else if (userData.rol === 'voluntario') {
            document.body.classList.add('is-voluntario');
            document.querySelectorAll('.admin-only').forEach(function(el) { el.classList.add('hidden'); });
        } else {
            document.querySelectorAll('.admin-only').forEach(function(el) { el.classList.add('hidden'); });
        }
        
        // Mostrar elementos para usuarios logueados (Mi Perfil, etc)
        document.body.classList.add('is-logged-in');
    }

    /** Actualiza la interfaz cuando un usuario inicia sesion */
    function updateUIForLoggedUser() {
        var user = AppState.currentUser;
        var initials = getInitials(user.name);

        document.getElementById('userAvatarSidebar').innerHTML = user.photo
            ? '<img src="' + user.photo + '" alt="' + user.name + '">'
            : initials;
        document.getElementById('userNameSidebar').textContent = user.name;
        
        let roleText = 'Usuario';
        if (AppState.isAdmin || user.rol === 'admin') roleText = 'Administrador';
        else if (user.rol === 'lider') roleText = 'Líder';
        else if (user.rol === 'voluntario') roleText = 'Voluntario';
        
        document.getElementById('userRoleSidebar').textContent = roleText;

        document.getElementById('userAvatarTop').innerHTML = user.photo
            ? '<img src="' + user.photo + '" alt="' + user.name + '" style="width:100%;height:100%;object-fit:cover;border-radius:50%;">'
            : initials;

        document.getElementById('sidebarLoginPrompt').classList.add('hidden');
        document.getElementById('sidebarUserInfo').classList.remove('hidden');
        document.getElementById('topBarLoginBtn').classList.add('hidden');
        document.getElementById('topBarUserAvatar').classList.remove('hidden');

        document.getElementById('forumAvatar').textContent = initials;
        document.getElementById('forumTextarea').removeAttribute('onclick');

        if (document.getElementById('donorName')) {
            document.getElementById('donorName').value = user.name;
        }
        if (document.getElementById('donorEmail') && user.email) {
            document.getElementById('donorEmail').value = user.email;
        }

        if (AppState.isAdmin) {
            document.body.classList.add('is-admin');
        }

        setTimeout(function() {
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }, 100);
    }

    /** Cierra la sesion del usuario */
    async function logout() {
        window.showAuthLoading('Cerrando sesión...');
        
        try {
            if (firebaseAuth) await firebaseAuth.signOut();
        } catch (e) { /* ignorar */ }
        
        // Pequeño delay para que se vea el modal
        await new Promise(r => setTimeout(r, 600));

        AppState.currentUser = null;
        AppState.isAdmin = false;
        localStorage.removeItem('fundacion_session');
        
        updateUIForLogout();

        navigateTo('inicio');
        window.hideAuthLoading();
        showToast('Sesion cerrada', 'info');
    }

    function updateUIForLogout() {
        document.getElementById('sidebarLoginPrompt').classList.remove('hidden');
        document.getElementById('sidebarUserInfo').classList.add('hidden');
        document.getElementById('topBarLoginBtn').classList.remove('hidden');
        document.getElementById('topBarUserAvatar').classList.add('hidden');

        document.getElementById('forumAvatar').textContent = '?';
        document.getElementById('forumTextarea').setAttribute('onclick', 'document.getElementById(\'loginModal\').classList.add(\'active\')');

        document.body.classList.remove('is-admin', 'is-lider', 'is-voluntario', 'is-logged-in');
        document.querySelectorAll('.admin-only').forEach(function(el) {
            el.classList.add('hidden');
        });
        
        if (window.ModernApp && window.ModernApp.auth) {
            window.ModernApp.auth.clearAuthState();
        }

        if (AppState.currentPage === 'panel') {
            switchPage('inicio');
        }
    }

    function checkLocalSession() {
        var saved = localStorage.getItem('fundacion_session');
        if (saved) {
            try {
                var userData = JSON.parse(saved);
                handleUserLogin(userData);
            } catch (e) {
                localStorage.removeItem('fundacion_session');
            }
        }
    }

    function requireLogin(action) {
        if (!AppState.currentUser) {
            showToast('Inicia sesion para ' + action, 'info');
            openLoginModal();
            return false;
        }
        return true;
    }

    // =====================================================================
    // GOOGLE SHEETS API (Base de Datos)
    // =====================================================================

    /**
     * Funcion generica para comunicarse con Google Sheets via Apps Script.
     * @param {string} method - 'GET' para leer o 'POST' para escribir
     * @param {string} sheet - Nombre de la hoja (pestana)
     * @param {object} data - Datos a enviar (solo para POST)
     * @returns {object} Respuesta del servidor
     */
    async function fetchFromSheets(method, sheet, data) {
        data = data || null;

        try {
            var url = GOOGLE_SHEETS_API_URL + '?sheet=' + encodeURIComponent(sheet);

            if (method === 'GET') {
                var response = await fetch(url);
                return await response.json();
            } else {
                var response = await fetch(GOOGLE_SHEETS_API_URL, {
                    method: 'POST',
                    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                    body: JSON.stringify(Object.assign({ sheet: sheet }, data))
                });
                return await response.json();
            }
        } catch (error) {
            console.error('Error al comunicarse con Google Sheets (' + sheet + '):', error);
            return null;
        }
    }

    async function registerUserInSheet(userData) {
        await fetchFromSheets('POST', 'Usuarios', {
            action: 'register',
            row: [
                userData.uid,
                userData.name,
                userData.email,
                AppState.isAdmin ? 'admin' : 'normal',
                new Date().toLocaleDateString('es-CO')
            ]
        });
    }

    async function registerVisit() {
        await fetchFromSheets('POST', 'Metricas', {
            action: 'increment',
            key: 'Visitas'
        });
    }

    // =====================================================================
    // FEED DE LABORES (Noticias)
    // =====================================================================

    /** Carga las publicaciones del feed */
    async function loadFeedData() {
        var response = await fetchFromSheets('GET', 'Feed');
        var feedGrid = document.getElementById('feedGrid');
        var feedEmpty = document.getElementById('feedEmpty');

        if (response && response.data && response.data.length > 0) {
            AppState.feedData = response.data;
            feedGrid.innerHTML = '';
            feedEmpty.classList.add('hidden');

            var gradients = [
                'linear-gradient(135deg, #059669, #34d399)',
                'linear-gradient(135deg, #0891b2, #67e8f9)',
                'linear-gradient(135deg, #7c3aed, #a78bfa)',
                'linear-gradient(135deg, #db2777, #f472b6)',
                'linear-gradient(135deg, #ea580c, #fb923c)',
                'linear-gradient(135deg, #0284c7, #38bdf8)'
            ];
            var icons = ['\u{1F392}', '\u{1F3E5}', '\u{1F4AA}', '\u{1F333}', '\u{1F37D}', '\u{1F384}'];

            response.data.forEach(function(post, index) {
                var isObj = typeof post === 'object' && !Array.isArray(post);
                var title = isObj ? post.title : (post[0] || 'Sin titulo');
                var desc = isObj ? post.description : (post[1] || '');
                var imgUrl = isObj ? post.imageUrl : (post[2] || '');
                var date = isObj ? post.date : (post[3] || '');
                var likes = isObj ? (post.likes || 0) : (parseInt(post[4]) || 0);
                var likedByStr = isObj ? (post.likedBy || []) : (post[5] || '');
                var likedBy = Array.isArray(likedByStr) ? likedByStr : (likedByStr ? likedByStr.split(',') : []);

                var isLiked = AppState.currentUser ? likedBy.includes(AppState.currentUser.uid) : false;
                var gradientIndex = index % gradients.length;

                var imgUrls = imgUrl ? imgUrl.split(',').map(function(u) {
                    var url = u.trim();
                    // Transformar URLs nativas de drive para usar el endpoint de thumbnail de alta resolucion
                    // Esto evita el bloqueo de CORS y cookies de 3ros
                    if (url.includes('drive.google.com/uc?id=')) {
                        return url.replace('uc?id=', 'thumbnail?sz=w1000&id=');
                    }
                    return url;
                }).filter(u => u !== '') : [];
                var coverUrl = imgUrls[0] || '';

                var imageHtml = coverUrl
                    ? '<img src="' + coverUrl + '" alt="' + escapeHtml(title) + '" loading="lazy" onerror="this.parentElement.innerHTML=\'<div class=placeholder style=background:' + gradients[gradientIndex].replace(/\s/g, '') + '><span style=font-size:3rem>' + icons[gradientIndex] + '</span></div>\'">'
                    : '<div class="placeholder" style="background:' + gradients[gradientIndex] + ';"><span style="font-size:3rem;">' + icons[gradientIndex] + '</span></div>';

                var card = document.createElement('div');
                card.className = 'feed-card';
                card.onclick = function(e) { 
                    // Si se hizo clic en el boton de like o sus hijos o botones de admin, no abrir modal
                    if (e.target.closest('.like-btn') || e.target.closest('.admin-post-actions')) return;
                    openPostModal(index); 
                };

                var adminHtml = '';
                if (AppState.isAdmin) {
                    adminHtml = '<div class="admin-post-actions">' +
                        '<button class="admin-post-action-btn" onclick="editPost(' + index + ')" title="Editar">' +
                            '<i data-lucide="pencil" width="16" height="16"></i>' +
                        '</button>' +
                        '<button class="admin-post-action-btn delete" onclick="deletePost(' + index + ')" title="Eliminar">' +
                            '<i data-lucide="trash-2" width="16" height="16"></i>' +
                        '</button>' +
                    '</div>';
                }

                card.innerHTML = adminHtml +
                    '<div class="feed-card-image">' + imageHtml + '</div>' +
                    '<div class="feed-card-body">' +
                        '<div class="feed-card-date">' + formatDate(date) + '</div>' +
                        '<h3>' + escapeHtml(title) + '</h3>' +
                        '<p>' + escapeHtml(desc) + '</p>' +
                        '<div class="feed-card-actions">' +
                            '<button class="like-btn ' + (isLiked ? 'liked' : '') + '" onclick="toggleLike(' + index + ', this)" id="likeBtn_' + index + '">' +
                                '<i data-lucide="heart" width="16" height="16" ' + (isLiked ? 'style="fill:#ef4444;"' : '') + '></i>' +
                                '<span id="likeCount_' + index + '">' + likes + '</span>' +
                            '</button>' +
                            '<span style="font-size:0.82rem; color:var(--text-tertiary); margin-left:auto;">' +
                                '<i data-lucide="calendar" width="14" height="14" style="display:inline;vertical-align:-2px;"></i> ' + date +
                            '</span>' +
                        '</div>' +
                    '</div>';
                feedGrid.appendChild(card);
            });
        } else {
            feedGrid.innerHTML = '';
            feedEmpty.classList.remove('hidden');
        }

        setTimeout(function() {
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }, 150);
    }

    /** Alterna el "Me gusta" de una publicacion */
    async function toggleLike(postIndex, buttonEl) {
        if (!requireLogin('dar Me gusta')) return;

        var userId = AppState.currentUser.uid;
        var post = AppState.feedData[postIndex];
        var isObj = typeof post === 'object' && !Array.isArray(post);

        var likedBy = [];
        if (isObj) {
            likedBy = Array.isArray(post.likedBy) ? post.likedBy.slice() : (post.likedBy ? post.likedBy.split(',') : []);
        }

        var isLiked = likedBy.includes(userId);

        if (isLiked) {
            likedBy = likedBy.filter(function(id) { return id !== userId; });
        } else {
            likedBy.push(userId);
        }

        var countEl = document.getElementById('likeCount_' + postIndex);
        var currentCount = parseInt(countEl.textContent);

        if (isLiked) {
            buttonEl.classList.remove('liked');
            countEl.textContent = Math.max(0, currentCount - 1);
        } else {
            buttonEl.classList.add('liked');
            countEl.textContent = currentCount + 1;
            buttonEl.style.transform = 'scale(1.2)';
            setTimeout(function() { buttonEl.style.transform = 'scale(1)'; }, 200);
        }

        if (isObj) {
            post.likedBy = likedBy;
            post.likes = likedBy.length;
        }

        await fetchFromSheets('POST', 'Feed', {
            action: 'updateLike',
            rowIndex: postIndex,
            likedBy: likedBy.join(','),
            likes: likedBy.length
        });

        setTimeout(function() {
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }, 100);
    }

    // =====================================================================
    // GESTION DE PUBLICACIONES (ADMIN)
    // =====================================================================

    var pendingDeleteIndex = -1;
    var pendingDeleteType = '';

    window.deletePost = function(index) {
        pendingDeleteIndex = index;
        pendingDeleteType = 'Feed';
        showDeleteModal();
    };

    window.deleteForumPost = function(index) {
        pendingDeleteIndex = index;
        pendingDeleteType = 'Foro';
        showDeleteModal();
    };

    function showDeleteModal() {
        var modal = document.getElementById('deleteConfirmModal');
        document.getElementById('deleteStateConfirm').classList.remove('hidden');
        document.getElementById('deleteStateLoading').classList.add('hidden');
        document.getElementById('deleteStateSuccess').classList.add('hidden');
        
        var titleText = '¿Eliminar Elemento?';
        var descText = 'El elemento se eliminó correctamente.';
        
        if (pendingDeleteType === 'Foro') {
            titleText = '¿Eliminar Comentario?';
            descText = 'El comentario se eliminó correctamente.';
        } else if (pendingDeleteType === 'Feed') {
            titleText = '¿Eliminar Publicación?';
            descText = 'La publicación se eliminó correctamente.';
        } else if (pendingDeleteType === 'Voluntarios') {
            titleText = '¿Eliminar Voluntario?';
            descText = 'El voluntario se eliminó correctamente.';
        } else if (pendingDeleteType === 'Donaciones') {
            titleText = '¿Eliminar Donación?';
            descText = 'La donación se eliminó correctamente.';
        }
        
        document.getElementById('deleteConfirmTitle').textContent = titleText;
        document.getElementById('deleteSuccessDesc').textContent = descText;

        var loadingH2 = document.querySelector('#deleteStateLoading h2');
        if (loadingH2) {
            loadingH2.textContent = pendingDeleteType === 'Donaciones' ? 'Eliminando donación...' : 'Eliminando...';
        }

        var btnConfirm = document.getElementById('btnConfirmDelete');
        btnConfirm.onclick = executeDelete;

        if (modal) {
            modal.classList.remove('hidden');
            document.body.style.overflow = 'hidden';
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }
    }

    window.closeDeleteModal = function() {
        var modal = document.getElementById('deleteConfirmModal');
        if (modal) {
            modal.classList.add('hidden');
            document.body.style.overflow = '';
        }
        pendingDeleteIndex = -1;
    };

    async function executeDelete() {
        if (pendingDeleteIndex === -1) return;

        // Estado 2: Cargando
        document.getElementById('deleteStateConfirm').classList.add('hidden');
        document.getElementById('deleteStateLoading').classList.remove('hidden');
        
        var result = await fetchFromSheets('POST', pendingDeleteType, {
            action: 'delete',
            rowIndex: pendingDeleteIndex
        });

        if (result && result.success !== false) {
            // Estado 3: Exito
            document.getElementById('deleteStateLoading').classList.add('hidden');
            document.getElementById('deleteStateSuccess').classList.remove('hidden');
            if (typeof lucide !== 'undefined') lucide.createIcons();
            
            if (pendingDeleteType === 'Feed') await loadFeedData();
            if (pendingDeleteType === 'Foro') {
                if (window.ModernApp && window.ModernApp.forum && window.ModernApp.forum.loadForumData) {
                    await window.ModernApp.forum.loadForumData();
                } else {
                    await loadForumData(); // fallback
                }
            }
            if (pendingDeleteType === 'Voluntarios') await loadVolunteers();
        } else {
            closeDeleteModal();
            showToast('Error al eliminar', 'error');
        }
    }

    window.editPost = function(index) {
        var post = AppState.feedData[index];
        if (!post) return;

        var isObj = typeof post === 'object' && !Array.isArray(post);
        var title = isObj ? post.title : (post[0] || '');
        var desc = isObj ? post.description : (post[1] || '');
        var imgUrl = isObj ? post.imageUrl : (post[2] || '');
        var date = isObj ? post.date : (post[3] || '');

        // Rellenar modal
        document.getElementById('editPostIndex').value = index;
        document.getElementById('editPostTitle').value = title;
        document.getElementById('editPostDescription').value = desc;
        // Para la fecha se requiere formato YYYY-MM-DD
        var dateVal = date;
        if (date.includes('/')) {
            var parts = date.split('/');
            if (parts.length === 3) dateVal = parts[2] + '-' + parts[1] + '-' + parts[0];
        }
        document.getElementById('editPostDate').value = dateVal;
        document.getElementById('editPostImageUrl').value = imgUrl;

        var modal = document.getElementById('editPostModal');
        if (modal) {
            modal.classList.remove('hidden');
            document.body.style.overflow = 'hidden';
        }
    };

    window.closeEditModal = function() {
        var modal = document.getElementById('editPostModal');
        if (modal) {
            modal.classList.add('hidden');
            document.body.style.overflow = '';
        }
    };

    window.saveEditPost = async function(event) {
        event.preventDefault();
        
        var index = document.getElementById('editPostIndex').value;
        var title = document.getElementById('editPostTitle').value.trim();
        var desc = document.getElementById('editPostDescription').value.trim();
        var date = document.getElementById('editPostDate').value;
        var imgUrl = document.getElementById('editPostImageUrl').value;
        
        // Obtener likes y likedBy actuales
        var post = AppState.feedData[index];
        var isObj = typeof post === 'object' && !Array.isArray(post);
        var likes = isObj ? (post.likes || 0) : (parseInt(post[4]) || 0);
        var likedBy = isObj ? (post.likedBy || '') : (post[5] || '');
        if (Array.isArray(likedBy)) likedBy = likedBy.join(',');

        var loader = document.getElementById('page-loader');
        if (loader) loader.style.display = 'flex';
        closeEditModal();

        var result = await fetchFromSheets('POST', 'Feed', {
            action: 'updateRow',
            rowIndex: parseInt(index),
            row: [title, desc, imgUrl, date, likes, likedBy]
        });

        if (loader) loader.style.display = 'none';

        if (result && result.success !== false) {
            showToast('Publicación actualizada correctamente', 'success');
            await loadFeedData();
        } else {
            showToast('Error al actualizar', 'error');
        }
    };

    // =====================================================================
    // VISOR DE POST (MODAL / LIGHTBOX)
    // =====================================================================

    window.openPostModal = function(index) {
        var post = AppState.feedData[index];
        if (!post) return;

        AppState.currentModalPostIndex = index;
        AppState.currentModalImageIndex = 0;

        var isObj = typeof post === 'object' && !Array.isArray(post);
        var title = isObj ? post.title : (post[0] || 'Sin titulo');
        var desc = isObj ? post.description : (post[1] || '');
        var imgUrl = isObj ? post.imageUrl : (post[2] || '');
        var date = isObj ? post.date : (post[3] || '');

        document.getElementById('postViewerTitle').textContent = title;
        document.getElementById('postViewerDate').textContent = formatDate(date);
        document.getElementById('postViewerDescription').textContent = desc;

        var imgUrls = imgUrl ? imgUrl.split(',').map(function(u) {
            var url = u.trim();
            if (url.includes('drive.google.com/uc?id=')) {
                return url.replace('uc?id=', 'thumbnail?sz=w1000&id=');
            }
            return url;
        }).filter(u => u !== '') : [];

        if (imgUrls.length === 0) {
            document.getElementById('postViewerImage').src = 'data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><rect width="100" height="100" fill="%230f766e"/><text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" font-size="40">🎒</text></svg>';
        }

        updateModalCarousel(imgUrls);

        var shareBtn = document.getElementById('sharePostBtn');
        if (shareBtn) shareBtn.onclick = function() { sharePost(index); };

        var modal = document.getElementById('postViewerModal');
        if (modal) {
            modal.classList.remove('hidden');
            document.body.style.overflow = 'hidden';
        }
        
        setTimeout(function() {
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }, 50);
    };

    window.closePostModal = function() {
        var modal = document.getElementById('postViewerModal');
        if (modal) {
            modal.classList.add('hidden');
            document.body.style.overflow = '';
        }
        AppState.currentModalPostIndex = -1;
    };

    window.changeModalImage = function(direction) {
        var post = AppState.feedData[AppState.currentModalPostIndex];
        if (!post) return;

        var isObj = typeof post === 'object' && !Array.isArray(post);
        var imgUrl = isObj ? post.imageUrl : (post[2] || '');
        var imgUrls = imgUrl ? imgUrl.split(',').map(function(u) {
            var url = u.trim();
            if (url.includes('drive.google.com/uc?id=')) {
                return url.replace('uc?id=', 'thumbnail?sz=w1000&id=');
            }
            return url;
        }).filter(u => u !== '') : [];

        if (imgUrls.length <= 1) return;

        AppState.currentModalImageIndex += direction;
        
        if (AppState.currentModalImageIndex < 0) {
            AppState.currentModalImageIndex = imgUrls.length - 1;
        } else if (AppState.currentModalImageIndex >= imgUrls.length) {
            AppState.currentModalImageIndex = 0;
        }

        updateModalCarousel(imgUrls);
    };

    function updateModalCarousel(imgUrls) {
        var imgElement = document.getElementById('postViewerImage');
        var prevBtn = document.getElementById('carouselPrevBtn');
        var nextBtn = document.getElementById('carouselNextBtn');
        var indicators = document.getElementById('carouselIndicators');

        if (imgUrls.length > 0) {
            imgElement.src = imgUrls[AppState.currentModalImageIndex];
        }

        if (imgUrls.length <= 1) {
            if (prevBtn) prevBtn.style.display = 'none';
            if (nextBtn) nextBtn.style.display = 'none';
            if (indicators) indicators.innerHTML = '';
        } else {
            if (prevBtn) prevBtn.style.display = 'flex';
            if (nextBtn) nextBtn.style.display = 'flex';
            
            if (indicators) {
                indicators.innerHTML = '';
                for (var i = 0; i < imgUrls.length; i++) {
                    var dot = document.createElement('div');
                    dot.className = 'carousel-indicator' + (i === AppState.currentModalImageIndex ? ' active' : '');
                    (function(index) {
                        dot.onclick = function() {
                            AppState.currentModalImageIndex = index;
                            updateModalCarousel(imgUrls);
                        };
                    })(i);
                    indicators.appendChild(dot);
                }
            }
        }
    }

    window.currentShareUrl = '';
    window.currentShareText = '';

    window.sharePost = function(index) {
        var post = AppState.feedData[index];
        var isObj = typeof post === 'object' && !Array.isArray(post);
        var title = isObj ? post.title : (post[0] || 'Publicación');

        var url = new URL(window.location.href);
        url.searchParams.set('post', index);
        window.currentShareUrl = url.toString();
        window.currentShareText = 'Mira esta labor de la Fundación GAP: ' + title;

        var modal = document.getElementById('shareMenuModal');
        if (modal) {
            modal.classList.remove('hidden');
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }
    };

    window.closeShareMenu = function() {
        var modal = document.getElementById('shareMenuModal');
        if (modal) {
            modal.classList.add('hidden');
        }
    };

    window.shareTo = function(platform) {
        var url = encodeURIComponent(window.currentShareUrl);
        var text = encodeURIComponent(window.currentShareText);

        function copyFallback() {
            navigator.clipboard.writeText(window.currentShareUrl).then(function() {
                showToast('¡Enlace copiado al portapapeles!', 'success');
            }).catch(function() {
                showToast('Error al copiar el enlace', 'error');
            });
        }

        if (platform === 'whatsapp') {
            window.open('https://api.whatsapp.com/send?text=' + text + '%0A%0A' + url, '_blank');
        } else if (platform === 'facebook') {
            window.open('https://www.facebook.com/sharer/sharer.php?u=' + url, '_blank');
        } else if (platform === 'instagram') {
            navigator.clipboard.writeText(window.currentShareUrl).then(function() {
                showToast('¡Link copiado! Abriendo Instagram...', 'success');
                setTimeout(function() {
                    window.open('https://www.instagram.com/', '_blank');
                }, 1500);
            }).catch(function() {
                showToast('Error al copiar el enlace', 'error');
            });
        } else if (platform === 'twitter') {
            window.open('https://twitter.com/intent/tweet?url=' + url + '&text=' + text, '_blank');
        } else if (platform === 'email') {
            window.location.href = 'mailto:?subject=Fundación GAP&body=' + text + '%0A%0A' + window.currentShareUrl;
        } else if (platform === 'copy') {
            copyFallback();
        } else if (platform === 'native') {
            // Prevenir crash del navegador (RESULT_CODE_KILLED_BAD_MESSAGE) en PC al intentar
            // compartir URLs locales (file:///) usando la API nativa de Windows.
            if (navigator.share && window.location.protocol !== 'file:') {
                navigator.share({
                    title: 'Fundación GAP',
                    text: window.currentShareText,
                    url: window.currentShareUrl
                }).catch(function(err) { 
                    console.log('Error compartiendo o cancelado', err); 
                    if (err.name !== 'AbortError') {
                        copyFallback();
                    }
                });
            } else {
                copyFallback();
            }
        }
        closeShareMenu();
    };

    // =====================================================================
    // DONACIONES
    // =====================================================================

    window.selectAmount = function(type, amount) {
        var form = document.getElementById('donationForm' + type);
        if(form) {
            form.querySelectorAll('.amount-option').forEach(function(opt) { opt.classList.remove('selected'); });
        }
        if(window.event && window.event.target && window.event.target.classList.contains('amount-option')) {
            window.event.target.classList.add('selected');
        }
        document.getElementById('donationAmount' + type).value = amount;
    }

    window.submitDonationOnline = function(event) {
        event.preventDefault();
        
        var name = document.getElementById('donorNameOnline').value.trim();
        var email = document.getElementById('donorEmailOnline').value.trim();
        var amount = document.getElementById('donationAmountOnline').value;
        var message = document.getElementById('donationNoteOnline').value.trim();
        
        if (!name || !email || !amount) {
            showToast('Por favor completa los campos requeridos', 'error');
            return;
        }

        var btn = event.target.querySelector('button[type="submit"]');
        var originalBtnHtml = btn.innerHTML;
        btn.disabled = true;
        btn.innerHTML = '<i data-lucide="loader" class="spin-anim"></i> Abriendo Wompi...';
        if (typeof lucide !== 'undefined') lucide.createIcons();

        // Wompi Widget Integration
        var checkout = new WidgetCheckout({
            currency: 'COP',
            amountInCents: parseInt(amount) * 100,
            reference: 'DON-' + Date.now(),
            publicKey: 'pub_test_X0zDA9ooKE8CEEEqwgepmDP6rceJ0123', // Demo Key
            redirectUrl: window.location.href, // Or redirect to a thank you page
            customerData: {
                email: email,
                fullName: name
            }
        });

        checkout.open(function (result) {
            var transaction = result.transaction;
            if(transaction && transaction.status === 'APPROVED') {
                showToast('¡Pago exitoso! Registrando tu donación...', 'success');
                var userId = AppState.currentUser ? AppState.currentUser.uid : 'anon';
                var date = new Date().toLocaleDateString('es-CO');
                fetchFromSheets('POST', 'Donaciones', {
                    action: 'add',
                    row: [name, email, amount, 'Wompi En Linea', date, message, userId, 'Confirmada', 'Pagado por Wompi']
                }).then(function() {
                    document.getElementById('donationFormOnline').reset();
                    document.getElementById('donationAmountOnline').value = '100000';
                    window.selectAmount('Online', 100000);
                });
            } else {
                showToast('El pago no fue aprobado o fue cancelado.', 'warning');
            }
            btn.disabled = false;
            btn.innerHTML = originalBtnHtml;
            if (typeof lucide !== 'undefined') lucide.createIcons();
        });
    };

    window.submitDonationManual = async function(event) {
        event.preventDefault();

        var name = document.getElementById('donorNameManual').value.trim();
        var amount = document.getElementById('donationAmountManual').value;
        var fileInput = document.getElementById('donationReceiptManual');
        var userId = AppState.currentUser ? AppState.currentUser.uid : 'anon';
        var email = AppState.currentUser ? AppState.currentUser.email : 'N/A';

        if (!name || !amount || !fileInput.files[0]) {
            showToast('Por favor completa todos los campos requeridos, incluyendo el comprobante', 'error');
            return;
        }

        var btn = event.target.querySelector('button[type="submit"]');
        var originalBtnHtml = btn.innerHTML;
        if (btn) {
            btn.disabled = true;
            btn.innerHTML = '<i data-lucide="loader" class="spin-anim"></i> Procesando comprobante...';
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        var comprobanteUrl = '';
        var file = fileInput.files[0];
        
        try {
            showToast('Subiendo comprobante...', 'info');
            var base64Data = await new Promise(function(resolve, reject) {
                var reader = new FileReader();
                reader.onload = function() { resolve(reader.result); };
                reader.onerror = reject;
                reader.readAsDataURL(file);
            });

            var uploadResult = await fetchFromSheets('POST', 'Donaciones', {
                action: 'uploadImage',
                base64: base64Data,
                fileName: file.name,
                mimeType: file.type
            });

            if (uploadResult && uploadResult.success && uploadResult.url) {
                comprobanteUrl = uploadResult.url;
            } else {
                showToast('Error al subir el comprobante. Intenta de nuevo.', 'error');
                if (btn) { btn.disabled = false; btn.innerHTML = originalBtnHtml; if (typeof lucide !== 'undefined') lucide.createIcons(); }
                return;
            }
        } catch (err) {
            showToast('Error leyendo el archivo', 'error');
            if (btn) { btn.disabled = false; btn.innerHTML = originalBtnHtml; if (typeof lucide !== 'undefined') lucide.createIcons(); }
            return;
        }

        showToast('Registrando donación...', 'info');
        var date = new Date().toLocaleDateString('es-CO');

        var result = await fetchFromSheets('POST', 'Donaciones', {
            action: 'add',
            row: [name, email, amount, 'Transferencia Manual', date, '', userId, 'Pendiente', comprobanteUrl]
        });

        if (result && result.success !== false) {
            showToast('¡Comprobante subido! Gracias por tu generosidad. Validaremos la donación pronto.', 'success');
            document.getElementById('donationFormManual').reset();
        } else {
            showToast('Error al registrar la donacion. Intenta de nuevo.', 'error');
        }

        if (btn) {
            btn.disabled = false;
            btn.innerHTML = originalBtnHtml;
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }
    }

    // =====================================================================
    // VOLUNTARIADO
    // =====================================================================

    window.submitVolunteer = async function(event) {
        event.preventDefault();
        var name = document.getElementById('volName').value.trim();
        var email = document.getElementById('volEmail').value.trim();
        var phone = document.getElementById('volPhone').value.trim();
        var skills = document.getElementById('volSkills').value.trim();
        var availability = document.getElementById('volAvailability').value;
        var message = document.getElementById('volMessage').value.trim();

        if(!name || !email || !phone || !skills || !availability) {
            showToast('Por favor completa los campos requeridos', 'error');
            return;
        }

        var btn = event.target.querySelector('button[type="submit"]');
        var originalBtnHtml = btn.innerHTML;
        btn.disabled = true;
        btn.innerHTML = '<i data-lucide="loader" class="spin-anim"></i> Enviando solicitud...';
        if (typeof lucide !== 'undefined') lucide.createIcons();

        var date = new Date().toLocaleDateString('es-CO');
        var result = await fetchFromSheets('POST', 'Voluntarios', {
            action: 'add',
            row: [date, name, email, phone, skills, availability, message, 'Pendiente']
        });

        if (result && result.success !== false) {
            showToast('¡Solicitud enviada! Nos pondremos en contacto pronto.', 'success');
            document.getElementById('volunteerForm').reset();
        } else {
            showToast('Error al enviar. Intenta de nuevo más tarde.', 'error');
        }

        btn.disabled = false;
        btn.innerHTML = originalBtnHtml;
        if (typeof lucide !== 'undefined') lucide.createIcons();
    };

    window.loadVolunteers = async function() {
        var tbodyPending = document.getElementById('volunteersPendingBody');
        var tbodyAccepted = document.getElementById('volunteersAcceptedBody');
        if(!tbodyPending || !tbodyAccepted) return;
        
        switchVolunteerTab('pending');

        tbodyPending.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px;">Cargando...</td></tr>';
        tbodyAccepted.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px;">Cargando...</td></tr>';
        
        var response = await fetchFromSheets('GET', 'Voluntarios');
        if (response && Array.isArray(response.data)) {
            tbodyPending.innerHTML = '';
            tbodyAccepted.innerHTML = '';
            
            var pendingCount = 0;
            var acceptedCount = 0;
            
            for (var i = 0; i < response.data.length; i++) {
                var v = response.data[i];
                var status = v[7] || 'Pendiente';
                var isAccepted = (status === 'Aceptado' || status === 'Activo' || status === 'Inactivo');
                var badgeClass = '';
                
                if (status === 'Aceptado' || status === 'Activo') badgeClass = 'success';
                else if (status === 'Inactivo') badgeClass = 'warning';
                else if (status === 'Rechazado') badgeClass = 'danger';
                else badgeClass = 'warning';
                
                var tr = document.createElement('tr');
                
                var actionsHtml = '';
                if (!isAccepted) {
                    actionsHtml = `
                        <button class="btn btn-sm btn-ghost" onclick="updateVolunteerStatus(${i}, 'Aceptado')" style="color:var(--success);" title="Aceptar"><i data-lucide="check" width="16" height="16"></i></button>
                        <button class="btn btn-sm btn-ghost" onclick="updateVolunteerStatus(${i}, 'Rechazado')" style="color:var(--danger);" title="Rechazar"><i data-lucide="x" width="16" height="16"></i></button>
                    `;
                } else {
                    var isCurrentlyActive = (status === 'Aceptado' || status === 'Activo');
                    var toggleStatus = isCurrentlyActive ? 'Inactivo' : 'Activo';
                    var toggleIcon = isCurrentlyActive ? 'pause' : 'play';
                    var toggleTitle = isCurrentlyActive ? 'Marcar como inactivo' : 'Marcar como activo';
                    var toggleColor = isCurrentlyActive ? 'var(--txt3)' : 'var(--success)';
                    
                    actionsHtml = `
                        <button class="btn btn-sm btn-ghost" onclick="updateVolunteerStatus(${i}, '${toggleStatus}')" style="color:${toggleColor};" title="${toggleTitle}"><i data-lucide="${toggleIcon}" width="16" height="16"></i></button>
                        <button class="btn btn-sm btn-ghost" onclick="deleteVolunteer(${i})" style="color:var(--danger);" title="Eliminar"><i data-lucide="trash-2" width="16" height="16"></i></button>
                    `;
                }

                var estadoHtml = isAccepted ? `<td><span class="badge ${badgeClass}">${status}</span></td>` : '';
                
                tr.innerHTML = `
                    <td>${v[0] || ''}</td>
                    <td style="font-weight:600;">${v[1] || ''}<br><span style="font-size:0.8rem;color:var(--txt3);font-weight:normal;">${v[2] || ''}</span></td>
                    <td>${v[3] || ''}</td>
                    <td>${v[4] || ''}</td>
                    <td>${v[5] || ''}</td>
                    ${estadoHtml}
                    <td style="display:flex; gap: 8px;">
                        ${actionsHtml}
                    </td>
                `;
                
                if (isAccepted) {
                    tbodyAccepted.appendChild(tr);
                    acceptedCount++;
                } else {
                    tbodyPending.appendChild(tr);
                    pendingCount++;
                }
            }
            
            if (pendingCount === 0) {
                tbodyPending.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px;color:var(--txt3);">No hay solicitudes pendientes.</td></tr>';
            }
            if (acceptedCount === 0) {
                tbodyAccepted.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px;color:var(--txt3);">No hay voluntarios aceptados todavía.</td></tr>';
            }
            
            document.getElementById('pendingCountBadge').textContent = pendingCount;
            document.getElementById('acceptedCountBadge').textContent = acceptedCount;
            
            if (typeof lucide !== 'undefined') lucide.createIcons();
        } else {
            tbodyPending.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px;color:var(--danger);">Error cargando voluntarios.</td></tr>';
            tbodyAccepted.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:20px;color:var(--danger);">Error cargando voluntarios.</td></tr>';
        }
    };

    window.switchVolunteerTab = function(tabName) {
        var tabs = ['Pending', 'Accepted', 'Activities'];
        tabs.forEach(function(t) {
            var btn = document.getElementById('tab' + t + 'Btn');
            var view = document.getElementById('view' + t);
            if (btn) {
                btn.classList.remove('active');
                btn.style.borderBottomColor = 'transparent';
                btn.style.color = 'var(--txt2)';
                btn.style.fontWeight = '500';
            }
            if (view) view.classList.add('hidden');
        });

        if (tabName === 'pending') {
            var btn = document.getElementById('tabPendingBtn');
            if (btn) {
                btn.classList.add('active');
                btn.style.borderBottomColor = 'var(--p500)';
                btn.style.color = 'var(--txt1)';
                btn.style.fontWeight = '600';
            }
            document.getElementById('viewPending').classList.remove('hidden');
            document.getElementById('volunteerExportButtons').classList.add('hidden');
        } else if (tabName === 'accepted') {
            var btn = document.getElementById('tabAcceptedBtn');
            if (btn) {
                btn.classList.add('active');
                btn.style.borderBottomColor = 'var(--p500)';
                btn.style.color = 'var(--txt1)';
                btn.style.fontWeight = '600';
            }
            document.getElementById('viewAccepted').classList.remove('hidden');
            document.getElementById('volunteerExportButtons').classList.remove('hidden');
        } else if (tabName === 'activities') {
            var btn = document.getElementById('tabActivitiesBtn');
            if (btn) {
                btn.classList.add('active');
                btn.style.borderBottomColor = 'var(--p500)';
                btn.style.color = 'var(--txt1)';
                btn.style.fontWeight = '600';
            }
            document.getElementById('viewActivities').classList.remove('hidden');
            document.getElementById('volunteerExportButtons').classList.add('hidden'); // Activities has its own export buttons per row
            
            // Cargar actividades si no están cargadas
            loadVolunteerActivitiesAdmin();
        }
    };

    window.deleteVolunteer = function(rowIndex) {
        pendingDeleteIndex = rowIndex;
        pendingDeleteType = 'Voluntarios';
        showDeleteModal();
    };

    window.updateVolunteerStatus = async function(rowIndex, status) {
        showToast('Actualizando estado...', 'info');
        var response = await fetchFromSheets('GET', 'Voluntarios');
        if (response && Array.isArray(response.data)) {
            var rowData = response.data[rowIndex];
            rowData[7] = status;

            var updateResult = await fetchFromSheets('POST', 'Voluntarios', {
                action: 'updateRow',
                rowIndex: rowIndex,
                row: rowData
            });

            if (updateResult && updateResult.success !== false) {
                showToast('Estado actualizado.', 'success');
                loadVolunteers();
            } else {
                showToast('Error al actualizar.', 'error');
            }
        }
    };

    // =====================================================================
    // ASISTENCIA A ACTIVIDADES (ADMIN)
    // =====================================================================
    window.loadVolunteerActivitiesAdmin = async function() {
        var tbody = document.getElementById('volunteersActivitiesBody');
        if (!tbody) return;
        
        tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;padding:20px;">Cargando...</td></tr>';
        
        var response = await fetchFromSheets('GET', 'Actividades');
        if (response && Array.isArray(response.data)) {
            tbody.innerHTML = '';
            var count = 0;
            
            // Reverse so newest are first
            for (var i = response.data.length - 1; i >= 0; i--) {
                var v = response.data[i];
                var actId = v[0];
                var title = v[1];
                var date = v[3];
                
                var asistentesRaw = String(v[4] || '');
                var asistentesArr = [];
                if (asistentesRaw) {
                    try { asistentesArr = JSON.parse(asistentesRaw); } catch(e) { asistentesArr = []; }
                }
                if (!Array.isArray(asistentesArr)) asistentesArr = [];
                
                var countAsis = asistentesArr.length;
                
                var tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${new Date(date).toLocaleString('es-CO')}</td>
                    <td style="font-weight:600;">${title}</td>
                    <td><span class="badge success">${countAsis} confirmados</span></td>
                    <td style="display:flex; gap: 8px;">
                        <button class="btn btn-sm btn-outline" onclick="exportActivityAttendanceExcel('${actId}')" style="color: #047857; border-color: #047857;" title="Exportar Excel">
                            <i data-lucide="table" width="14" height="14"></i> Excel
                        </button>
                        <button class="btn btn-sm btn-outline" onclick="exportActivityAttendancePDF('${actId}')" style="color: #b91c1c; border-color: #b91c1c;" title="Exportar PDF (Hoja de firmas)">
                            <i data-lucide="file-text" width="14" height="14"></i> PDF
                        </button>
                    </td>
                `;
                tbody.appendChild(tr);
                count++;
            }
            
            if (count === 0) {
                tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;padding:20px;color:var(--txt3);">No hay actividades registradas.</td></tr>';
            }
            
            if (typeof lucide !== 'undefined') lucide.createIcons();
        } else {
            tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;padding:20px;color:var(--danger);">Error cargando actividades.</td></tr>';
        }
    };

    window.exportActivityAttendanceExcel = async function(actId) {
        showToast('Generando Excel...', 'info');
        var response = await fetchFromSheets('GET', 'Actividades');
        if (!response || !Array.isArray(response.data)) {
            showToast('Error al obtener datos', 'error');
            return;
        }
        
        var actData = response.data.find(r => String(r[0]) === String(actId));
        if (!actData) return showToast('Actividad no encontrada', 'error');
        
        var asistentesRaw = String(actData[4] || '');
        var asistentesArr = [];
        try { asistentesArr = JSON.parse(asistentesRaw); } catch(e) { }
        if (!Array.isArray(asistentesArr)) asistentesArr = [];
        
        var csvContent = "data:text/csv;charset=utf-8,";
        csvContent += "Nombre del Voluntario,Correo Electronico,Firma\n";
        
        asistentesArr.forEach(function(a) {
            var row = '"' + (a.name || '') + '","' + (a.email || '') + '",""';
            csvContent += row + "\n";
        });
        
        var encodedUri = encodeURI(csvContent);
        var link = document.createElement("a");
        link.setAttribute("href", encodedUri);
        link.setAttribute("download", "Asistencia_" + actData[1].replace(/[^a-z0-9]/gi, '_') + ".csv");
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    };

    window.exportActivityAttendancePDF = async function(actId) {
        showToast('Generando PDF...', 'info');
        var response = await fetchFromSheets('GET', 'Actividades');
        if (!response || !Array.isArray(response.data)) {
            showToast('Error al obtener datos', 'error');
            return;
        }
        
        var actData = response.data.find(r => String(r[0]) === String(actId));
        if (!actData) return showToast('Actividad no encontrada', 'error');
        
        var asistentesRaw = String(actData[4] || '');
        var asistentesArr = [];
        try { asistentesArr = JSON.parse(asistentesRaw); } catch(e) { }
        if (!Array.isArray(asistentesArr)) asistentesArr = [];
        
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();
        
        const startY = await addInstitutionalHeader(doc, "REPORTE DE ASISTENCIA A ACTIVIDADES", "GAP: EMPRENDE", "Este documento certifica el listado oficial de asistencia a la actividad programada por la fundación.");
        
        // Draw info box
        doc.setDrawColor(187, 247, 208); // border #bbf7d0
        doc.setFillColor(240, 253, 244); // bg #f0fdf4
        doc.roundedRect(14, startY, 182, 25, 2, 2, 'FD');
        
        doc.setTextColor(51, 51, 51);
        doc.setFontSize(11);
        
        doc.setFont(undefined, 'bold');
        doc.text('Actividad:', 18, startY + 7);
        doc.setFont(undefined, 'normal');
        doc.text(actData[1] || '', 40, startY + 7);
        
        doc.setFont(undefined, 'bold');
        doc.text('Fecha y Hora:', 18, startY + 14);
        doc.setFont(undefined, 'normal');
        doc.text(new Date(actData[3]).toLocaleString('es-CO'), 45, startY + 14);
        
        doc.setFont(undefined, 'bold');
        doc.text('Total Confirmados:', 18, startY + 21);
        doc.setFont(undefined, 'normal');
        doc.text(asistentesArr.length.toString(), 55, startY + 21);
        
        const tableData = asistentesArr.map((a, idx) => [
            idx + 1,
            a.name || 'Sin nombre',
            a.email || 'Sin correo',
            '' // Firma
        ]);
        
        for (let i = 0; i < 3; i++) {
            tableData.push(['-', '', '', '']);
        }
        
        doc.autoTable({
            startY: startY + 30,
            head: [['#', 'Nombre del Voluntario', 'Correo Electrónico', 'Firma']],
            body: tableData,
            theme: 'grid',
            headStyles: { fillColor: [4, 120, 87] },
            columnStyles: {
                0: { cellWidth: 10, halign: 'center' },
                1: { cellWidth: 60 },
                2: { cellWidth: 60 },
                3: { cellWidth: 'auto' }
            },
            didParseCell: function(data) {
                if (data.section === 'body') {
                    data.cell.styles.minCellHeight = 15;
                }
            }
        });
        
        doc.save("Asistencia_" + actData[1].replace(/[^a-z0-9]/gi, '_') + ".pdf");
    };

    // =====================================================================
    // INFORMACION BANCARIA (CARGA Y EDICION)
    // =====================================================================
    async function loadBankInfo() {
        var metricsResponse = await fetchFromSheets('GET', 'Metricas');
        if (metricsResponse && Array.isArray(metricsResponse.data)) {
            var bankData = {};
            metricsResponse.data.forEach(function(row) {
                if (Array.isArray(row)) {
                    bankData[row[0]] = row[1];
                }
            });

            // Actualizar vista
            if (bankData['Banco']) document.getElementById('displayBanco').textContent = bankData['Banco'];
            if (bankData['TipoCuenta']) document.getElementById('displayTipoCuenta').textContent = bankData['TipoCuenta'];
            if (bankData['NumCuenta']) document.getElementById('displayNumCuenta').textContent = bankData['NumCuenta'];
            if (bankData['Titular']) document.getElementById('displayTitular').textContent = bankData['Titular'];
            if (bankData['NIT']) document.getElementById('displayNIT').textContent = bankData['NIT'];
            if (bankData['Nequi']) document.getElementById('displayNequi').textContent = bankData['Nequi'];
            
            // Guardar para el modal
            AppState.bankData = bankData;
        }
    }

    window.openEditBankInfoModal = function() {
        var data = AppState.bankData || {};
        document.getElementById('editBanco').value = data['Banco'] || 'Bancolombia';
        document.getElementById('editTipoCuenta').value = data['TipoCuenta'] || 'Ahorros';
        document.getElementById('editNumCuenta').value = data['NumCuenta'] || '123-456789-00';
        document.getElementById('editTitular').value = data['Titular'] || 'Fundación GAP';
        document.getElementById('editNIT').value = data['NIT'] || '900.123.456-7';
        document.getElementById('editNequi').value = data['Nequi'] || '300 123 4567';
        
        document.getElementById('editBankInfoModal').classList.remove('hidden');
    };

    window.closeEditBankInfoModal = function() {
        document.getElementById('editBankInfoModal').classList.add('hidden');
    };

    window.saveBankInfo = async function() {
        var btn = document.getElementById('btnSaveBankInfo');
        btn.disabled = true;
        btn.innerHTML = '<i data-lucide="loader-2" class="spin-anim"></i> Guardando...';
        if (typeof lucide !== 'undefined') lucide.createIcons();

        var newSettings = {
            'Banco': document.getElementById('editBanco').value.trim(),
            'TipoCuenta': document.getElementById('editTipoCuenta').value.trim(),
            'NumCuenta': document.getElementById('editNumCuenta').value.trim(),
            'Titular': document.getElementById('editTitular').value.trim(),
            'NIT': document.getElementById('editNIT').value.trim(),
            'Nequi': document.getElementById('editNequi').value.trim()
        };

        var res = await fetchFromSheets('POST', 'Metricas', {
            action: 'updateSettings',
            settings: newSettings
        });

        btn.disabled = false;
        btn.innerHTML = '<i data-lucide="save"></i> Guardar Cambios';
        if (typeof lucide !== 'undefined') lucide.createIcons();

        if (res && res.success) {
            showToast('Información bancaria actualizada correctamente', 'success');
            closeEditBankInfoModal();
            loadBankInfo();
        } else {
            showToast('Error al actualizar: ' + (res.error || 'Desconocido'), 'error');
        }
    };

    function copyToClipboard(text) {
        navigator.clipboard.writeText(text).then(function() {
            showToast('Copiado al portapapeles', 'success');
        }).catch(function() {
            var textarea = document.createElement('textarea');
            textarea.value = text;
            document.body.appendChild(textarea);
            textarea.select();
            document.execCommand('copy');
            document.body.removeChild(textarea);
            showToast('Copiado al portapapeles', 'success');
        });
    }

    // =====================================================================
    // FORO DE COMUNIDAD
    // =====================================================================

    /** Carga los comentarios del foro */
    async function loadForumData() {
        // Obsoleto: ahora manejado por forum.js
    }

    async function submitForumPost() {
        // Obsoleto: ahora manejado por forum.js
    }

    // (La función de eliminar foro ahora usa la misma modal: deleteForumPost)

    function likeForumPost(btn) {
        // Obsoleto: ahora manejado por forum.js
    }

    // =====================================================================
    // PANEL DE ADMINISTRACION
    // =====================================================================

    async function loadAdminDashboard() {
        var metricsResponse = await fetchFromSheets('GET', 'Metricas');
        var donationsResponse = await fetchFromSheets('GET', 'Donaciones');
        var forumResponse = await fetchFromSheets('GET', 'Foro');

        var donations = (donationsResponse && donationsResponse.data) ? donationsResponse.data : [];
        var forumPosts = (forumResponse && forumResponse.data) ? forumResponse.data : [];

        var visits = 0;
        if (metricsResponse && Array.isArray(metricsResponse.data)) {
            metricsResponse.data.forEach(function(row) {
                if (Array.isArray(row) && row[0] === 'Visitas') visits = parseInt(row[1]) || 0;
            });
        }

        var totalDonations = donations.reduce(function(sum, d) {
            var amount = (typeof d === 'object' && !Array.isArray(d)) ? d.amount : (d[2] || 0);
            return sum + (parseInt(amount) || 0);
        }, 0);

        var emailSet = {};
        donations.forEach(function(d) {
            var email = (typeof d === 'object' && !Array.isArray(d)) ? d.email : (d[1] || '');
            emailSet[email] = true;
        });
        var uniqueDonors = Object.keys(emailSet).length;

        document.getElementById('metricVisits').textContent = visits.toLocaleString('es-CO');
        document.getElementById('metricDonations').textContent = '$' + totalDonations.toLocaleString('es-CO');
        document.getElementById('metricDonors').textContent = uniqueDonors;
        document.getElementById('metricPosts').textContent = forumPosts.length;

        renderDonationsChart(donations);
        renderMethodsChart(donations);
    }

    function renderDonationsChart(donations) {
        var ctx = document.getElementById('donationsChart');
        if (!ctx) return;

        if (window.donationsChartInstance) window.donationsChartInstance.destroy();

        var monthlyData = {};
        var months = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];

        donations.forEach(function(d) {
            var isObj = typeof d === 'object' && !Array.isArray(d);
            var dateStr = isObj ? d.date : (d[4] || '');
            var amount = isObj ? d.amount : (parseInt(d[2]) || 0);
            var monthIndex = new Date().getMonth();
            try {
                var parts = dateStr.split(/[/\-]/);
                if (parts.length >= 2) {
                    monthIndex = parseInt(parts[1]) - 1;
                    if (isNaN(monthIndex) || monthIndex < 0 || monthIndex > 11) monthIndex = new Date().getMonth();
                }
            } catch (e) {}
            var monthKey = months[monthIndex];
            monthlyData[monthKey] = (monthlyData[monthKey] || 0) + (parseInt(amount) || 0);
        });

        var currentMonth = new Date().getMonth();
        var labels = [];
        var values = [];
        for (var i = 5; i >= 0; i--) {
            var monthIdx = (currentMonth - i + 12) % 12;
            labels.push(months[monthIdx]);
            values.push((monthlyData[months[monthIdx]] || 0) / 1000);
        }

        window.donationsChartInstance = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{
                    label: 'Donaciones (miles $)',
                    data: values,
                    backgroundColor: [
                        'rgba(16, 185, 129, 0.2)',
                        'rgba(16, 185, 129, 0.3)',
                        'rgba(16, 185, 129, 0.4)',
                        'rgba(16, 185, 129, 0.5)',
                        'rgba(16, 185, 129, 0.6)',
                        'rgba(16, 185, 129, 0.8)',
                    ],
                    borderColor: 'rgba(16, 185, 129, 1)',
                    borderWidth: 2,
                    borderRadius: 8,
                    borderSkipped: false,
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        backgroundColor: '#064e3b',
                        titleColor: '#fff',
                        bodyColor: '#d1fae5',
                        padding: 12,
                        cornerRadius: 8,
                        callbacks: {
                            label: function(context) {
                                return '$' + (context.parsed.y * 1000).toLocaleString('es-CO');
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        grid: { display: false },
                        ticks: { color: '#7a9488', font: { weight: '500' } }
                    },
                    y: {
                        grid: { color: 'rgba(16,185,129,0.08)' },
                        ticks: {
                            color: '#7a9488',
                            callback: function(value) { return '$' + value + 'K'; }
                        }
                    }
                }
            }
        });
    }

    function renderMethodsChart(donations) {
        var ctx = document.getElementById('methodsChart');
        if (!ctx) return;
        if (window.methodsChartInstance) window.methodsChartInstance.destroy();

        var methods = {};
        donations.forEach(function(d) {
            var isObj = typeof d === 'object' && !Array.isArray(d);
            var method = isObj ? d.method : (d[3] || 'otro');
            methods[method] = (methods[method] || 0) + 1;
        });

        var labels = Object.keys(methods).map(function(m) { return m.charAt(0).toUpperCase() + m.slice(1); });
        var values = Object.values(methods);
        var colors = ['#10b981', '#f97316', '#3b82f6', '#a855f7', '#ef4444'];

        window.methodsChartInstance = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: labels,
                datasets: [{
                    data: values,
                    backgroundColor: colors.slice(0, labels.length),
                    borderWidth: 0,
                    hoverOffset: 8
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                cutout: '65%',
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            padding: 16,
                            usePointStyle: true,
                            pointStyleWidth: 10,
                            font: { size: 12, weight: '500' },
                            color: '#4a6355'
                        }
                    }
                }
            }
        });
    }

    async function loadDonorsTable() {
        var response = await fetchFromSheets('GET', 'Donaciones');
        var tbody = document.getElementById('donorsTableBody');
        var tableElement = document.getElementById('donorsTableBody').parentElement;
        var thead = tableElement.querySelector('thead tr');
        if (thead && !thead.innerHTML.includes('ACCIONES')) {
            var th = document.createElement('th');
            th.textContent = 'ACCIONES';
            thead.appendChild(th);
        }

        if (response && response.data && response.data.length > 0) {
            AppState.donationsData = response.data; // Guardar datos para edicion
            tbody.innerHTML = '';
            response.data.forEach(function(d, index) {
                var isObj = typeof d === 'object' && !Array.isArray(d);
                var name = isObj ? d.name : (d[0] || '');
                var email = isObj ? d.email : (d[1] || '');
                var amount = isObj ? d.amount : (d[2] || 0);
                var method = isObj ? d.method : (d[3] || '');
                var date = isObj ? d.date : (d[4] || '');
                var status = isObj ? d.status : (d[7] || 'Pendiente');
                var comprobanteUrl = isObj ? d.comprobante : (d[8] || '');
                
                var initials = getInitials(name);
                var statusClass = status === 'Confirmada' ? 'success' : 'pending';

                var row = document.createElement('tr');
                
                var actionsHtml = '<td style="display:flex; gap:8px;">' +
                    '<button class="admin-post-action-btn edit" onclick="openEditDonationModal(' + index + ')" title="Editar estado" style="position:relative; width:32px; height:32px; right:0; top:0; background:var(--bg2);"><i data-lucide="pencil" style="width:16px; height:16px;"></i></button>';
                
                if (comprobanteUrl && comprobanteUrl.trim() !== '') {
                    actionsHtml += '<button class="admin-post-action-btn edit" onclick="viewDonationReceipt(\'' + index + '\')" title="Ver comprobante" style="position:relative; width:32px; height:32px; right:0; top:0; background:var(--bg2); color:var(--p600);"><i data-lucide="file-text" style="width:16px; height:16px;"></i></button>';
                } else {
                    actionsHtml += '<button class="admin-post-action-btn edit" disabled title="Sin comprobante" style="position:relative; width:32px; height:32px; right:0; top:0; background:transparent; border:1px solid var(--b2); opacity:0.5; cursor:not-allowed;"><i data-lucide="file-text" style="width:16px; height:16px; color:var(--txt3);"></i></button>';
                }
                actionsHtml += '</td>';

                row.innerHTML =
                    '<td><div class="donor-name-cell"><div class="donor-avatar-sm">' + initials + '</div><span style="font-weight:600;color:var(--text-primary);">' + escapeHtml(name) + '</span></div></td>' +
                    '<td>' + escapeHtml(email) + '</td>' +
                    '<td style="font-weight:700;color:var(--primary-700);">$' + parseInt(amount).toLocaleString('es-CO') + '</td>' +
                    '<td>' + escapeHtml(method.charAt(0).toUpperCase() + method.slice(1)) + '</td>' +
                    '<td>' + date + '</td>' +
                    '<td><span class="status-badge ' + statusClass + '">' + status + '</span></td>' +
                    actionsHtml;
                tbody.appendChild(row);
            });
            if (typeof lucide !== 'undefined') lucide.createIcons();
        } else {
            AppState.donationsData = [];
            tbody.innerHTML = '<tr><td colspan="7" class="text-center" style="padding:40px;color:var(--text-tertiary);">No hay donaciones registradas</td></tr>';
        }
    }

    // =====================================================================
    // LOGICA DE EDICION DE DONACIONES (ADMIN)
    // =====================================================================

    window.deleteDonation = function(index) {
        pendingDeleteIndex = index;
        pendingDeleteType = 'Donaciones';
        if (typeof showDeleteModal !== 'undefined') showDeleteModal();
    };

    window.openEditDonationModal = function(index) {
        var donation = AppState.donationsData[index];
        if (!donation) return;

        AppState.currentEditingDonationIndex = index;
        var isObj = typeof donation === 'object' && !Array.isArray(donation);
        
        var name = isObj ? donation.name : (donation[0] || '');
        var amount = isObj ? donation.amount : (donation[2] || 0);
        var status = isObj ? donation.status : (donation[7] || 'Pendiente');

        document.getElementById('editDonationName').textContent = name;
        document.getElementById('editDonationAmount').textContent = '$' + parseInt(amount).toLocaleString('es-CO');
        
        var statusSelect = document.getElementById('editDonationStatus');
        if (Array.from(statusSelect.options).some(opt => opt.value === status)) {
            statusSelect.value = status;
        } else {
            statusSelect.value = 'Pendiente'; // fallback
        }

        document.getElementById('editDonationModal').classList.remove('hidden');
    };

    window.closeEditDonationModal = function() {
        document.getElementById('editDonationModal').classList.add('hidden');
        AppState.currentEditingDonationIndex = -1;
    };

    window.saveDonationStatus = async function() {
        var index = AppState.currentEditingDonationIndex;
        if (index < 0 || !AppState.donationsData || !AppState.donationsData[index]) return;

        var btn = document.getElementById('btnSaveDonation');
        var newStatus = document.getElementById('editDonationStatus').value;
        var originalRow = AppState.donationsData[index];

        // Asegurar que el array tenga al menos 8 o 9 elementos
        var updatedRowArray = Array.isArray(originalRow) ? originalRow.slice() : [];
        if (!Array.isArray(originalRow)) {
            // Si por alguna razon es objeto (no deberia pasar con App Script GET)
            updatedRowArray = [
                originalRow.name || '', originalRow.email || '', originalRow.amount || '', 
                originalRow.method || '', originalRow.date || '', originalRow.message || '', 
                originalRow.userId || '', newStatus, originalRow.comprobante || ''
            ];
        } else {
            // Rellenar hasta el indice 8 si es necesario
            while(updatedRowArray.length < 9) {
                updatedRowArray.push('');
            }
            updatedRowArray[7] = newStatus;
        }

        btn.disabled = true;
        btn.innerHTML = '<i data-lucide="loader-2" class="spin-anim"></i> Guardando...';
        if (typeof lucide !== 'undefined') lucide.createIcons();

        var res = await fetchFromSheets('POST', '', {
            action: 'updateRow',
            sheet: 'Donaciones',
            rowIndex: index,
            row: updatedRowArray
        });

        btn.disabled = false;
        btn.innerHTML = '<i data-lucide="save"></i> Guardar Cambios';
        if (typeof lucide !== 'undefined') lucide.createIcons();

        if (res && res.success) {
            showToast('Estado de donación actualizado correctamente', 'success');
            closeEditDonationModal();
            loadDonorsTable();
        } else {
            showToast('Error al actualizar: ' + (res.error || 'Desconocido'), 'error');
        }
    };

    window.viewDonationReceipt = function(index) {
        var donation = AppState.donationsData[index];
        if (!donation) return;
        var isObj = typeof donation === 'object' && !Array.isArray(donation);
        var comprobanteUrl = isObj ? donation.comprobante : (donation[8] || '');
        
        if (comprobanteUrl && comprobanteUrl.trim() !== '') {
            window.open(comprobanteUrl, '_blank');
        } else {
            showToast('No hay comprobante para esta donación', 'error');
        }
    };

    async function loadForumPostsAdmin() {
        var response = await fetchFromSheets('GET', 'Foro');
        var container = document.getElementById('adminForumList');

        if (response && response.data && response.data.length > 0) {
            container.innerHTML = '';
            response.data.forEach(function(post, index) {
                var isObj = typeof post === 'object' && !Array.isArray(post);
                var author = isObj ? post.author : (post[0] || 'Anonimo');
                var message = isObj ? post.message : (post[1] || '');
                var date = isObj ? post.date : (post[2] || '');

                var item = document.createElement('div');
                item.style.cssText = 'padding:12px;border:1px solid var(--border-light);border-radius:var(--radius-sm);margin-bottom:8px;display:flex;align-items:center;justify-content:space-between;gap:12px;';
                item.innerHTML =
                    '<div style="flex:1;min-width:0;">' +
                        '<div style="font-weight:600;font-size:0.85rem;">' + escapeHtml(author) + ' <span style="font-weight:400;color:var(--text-tertiary);font-size:0.75rem;"> - ' + date + '</span></div>' +
                        '<div style="font-size:0.82rem;color:var(--text-secondary);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">' + escapeHtml(message) + '</div>' +
                    '</div>' +
                    '<button class="mod-btn delete" onclick="deleteForumPost(' + index + ')">Eliminar</button>';
                container.appendChild(item);
            });
        } else {
            container.innerHTML = '<div style="text-align:center;padding:20px;color:var(--text-tertiary);">No hay comentarios en el foro</div>';
        }
    }

    window.previewPostImage = function(event) {
        var newFiles = event.target.files;
        var container = document.getElementById('postImagePreviewContainer');
        
        // Agregar los nuevos archivos al array acumulativo
        if (newFiles && newFiles.length > 0) {
            for (var i = 0; i < newFiles.length; i++) {
                AppState.pendingImageFiles.push(newFiles[i]);
            }
        }

        renderPendingImages();
    };

    window.removePendingImage = function(index) {
        AppState.pendingImageFiles.splice(index, 1);
        renderPendingImages();
    };

    function renderPendingImages() {
        var container = document.getElementById('postImagePreviewContainer');
        container.innerHTML = '';
        
        if (AppState.pendingImageFiles.length > 0) {
            container.style.display = 'grid';
            for (var i = 0; i < AppState.pendingImageFiles.length; i++) {
                (function(file, index) {
                    var reader = new FileReader();
                    reader.onload = function(e) {
                        var wrapper = document.createElement('div');
                        wrapper.style.cssText = 'position:relative; width:100%; height:100px; border-radius:4px; overflow:hidden; border:1px solid var(--border-light);';
                        
                        var img = document.createElement('img');
                        img.src = e.target.result;
                        img.style.cssText = 'width:100%; height:100%; object-fit:cover;';
                        
                        var btn = document.createElement('button');
                        btn.innerHTML = '&times;';
                        btn.style.cssText = 'position:absolute; top:4px; right:4px; background:rgba(0,0,0,0.6); color:white; border:none; border-radius:50%; width:24px; height:24px; cursor:pointer; font-size:16px; display:flex; align-items:center; justify-content:center;';
                        btn.onclick = function(ev) {
                            ev.preventDefault();
                            removePendingImage(index);
                        };

                        wrapper.appendChild(img);
                        wrapper.appendChild(btn);
                        container.appendChild(wrapper);
                    };
                    reader.readAsDataURL(file);
                })(AppState.pendingImageFiles[i], i);
            }
        } else {
            container.style.display = 'none';
        }
        
        // Limpiar el input para que pueda volver a seleccionar el mismo archivo si quiere
        var fileInput = document.getElementById('postImageFile');
        if (fileInput) fileInput.value = '';
    }

    async function submitNewPost(event) {
        event.preventDefault();

        var title = document.getElementById('postTitle').value.trim();
        var description = document.getElementById('postDescription').value.trim();
        var date = document.getElementById('postDate').value;
        var imageFiles = AppState.pendingImageFiles; 
        var submitBtn = event.target.querySelector('button[type="submit"]') || event.target.querySelector('.btn-primary');

        if (!title || !description || !date) {
            showToast('Completa todos los campos requeridos', 'error');
            return;
        }

        // --- INICIO DE ESTADO DE CARGA ---
        if (submitBtn) {
            submitBtn.disabled = true;
            submitBtn.innerHTML = '<i data-lucide="loader" class="spin"></i> Publicando...';
        }
        var loader = document.getElementById('page-loader');
        if (loader) loader.style.display = 'flex';
        // ---------------------------------

        var imageUrls = [];

        if (imageFiles && imageFiles.length > 0) {
            showToast('Subiendo ' + imageFiles.length + ' foto(s)...', 'info');
            try {
                for (var i = 0; i < imageFiles.length; i++) {
                    var imgFile = imageFiles[i];
                    showToast('Subiendo foto ' + (i+1) + ' de ' + imageFiles.length + '...', 'info');
                    
                    var base64Data = await new Promise(function(resolve, reject) {
                        var reader = new FileReader();
                        reader.onload = function() { resolve(reader.result); };
                        reader.onerror = reject;
                        reader.readAsDataURL(imgFile);
                    });

                    var uploadResult = await fetchFromSheets('POST', 'Feed', {
                        action: 'uploadImage',
                        base64: base64Data,
                        fileName: imgFile.name,
                        mimeType: imgFile.type
                    });

                    if (uploadResult && uploadResult.success && uploadResult.url) {
                        imageUrls.push(uploadResult.url);
                    } else {
                        showToast('Error al subir la imagen ' + (i+1), 'error');
                    }
                }
            } catch (err) {
                showToast('Error procesando las imagenes', 'error');
                return;
            }
        }

        var finalImageUrl = imageUrls.join(',');
        showToast('Guardando publicacion en la base de datos...', 'info');

        var result = await fetchFromSheets('POST', 'Feed', {
            action: 'add',
            row: [title, description, finalImageUrl, date, 0, '']
        });

        // --- FIN DE ESTADO DE CARGA ---
        if (loader) loader.style.display = 'none';
        if (submitBtn) {
            submitBtn.disabled = false;
            submitBtn.innerHTML = '<i data-lucide="upload"></i> Publicar en el Feed';
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }
        // ------------------------------

        if (result && result.success !== false) {
            showToast('Publicacion creada exitosamente!', 'success');
            document.getElementById('newPostForm').reset();
            document.getElementById('postDate').value = new Date().toISOString().split('T')[0];
            AppState.pendingImageFiles = []; // Limpiar acumulador
            renderPendingImages();
            await loadFeedData();
        } else {
            showToast('Error al crear la publicacion', 'error');
        }
    }

    // =====================================================================
    // UTILIDADES
    // =====================================================================

    function getInitials(name) {
        if (!name) return '?';
        return name.split(' ').map(function(n) { return n[0]; }).join('').toUpperCase().slice(0, 2);
    }

    function formatDate(dateStr) {
        if (!dateStr) return '';
        try {
            var options = { year: 'numeric', month: 'long', day: 'numeric' };
            var date = new Date(dateStr + 'T00:00:00');
            if (isNaN(date.getTime())) return dateStr;
            return date.toLocaleDateString('es-CO', options);
        } catch (e) {
            return dateStr;
        }
    }

    function escapeHtml(text) {
        if (!text) return '';
        var div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    function showToast(message, type) {
        type = type || 'info';
        var container = document.getElementById('toastContainer');
        var toast = document.createElement('div');
        toast.className = 'toast ' + type;

        var icons = { success: '&#9989;', error: '&#10060;', info: '&#8505;' };

        toast.innerHTML =
            '<span style="font-size:1.1rem;">' + (icons[type] || '&#8505;') + '</span>' +
            '<span style="flex:1;">' + message + '</span>' +
            '<button onclick="this.parentElement.remove()" style="opacity:0.5;padding:4px;cursor:pointer;">&#10005;</button>';

        container.appendChild(toast);

        setTimeout(function() {
            toast.classList.add('hiding');
            setTimeout(function() { toast.remove(); }, 300);
        }, 4000);
    }

/* ============================================================
   EXPORTACIONES A PDF Y EXCEL
   ============================================================ */

window.cachedLogoBase64 = null;
async function getLogoBase64() {
    if (window.cachedLogoBase64) return window.cachedLogoBase64;
    return new Promise((resolve) => {
        window.cachedLogoBase64 = 'data:image/jpeg;base64,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';
        resolve(window.cachedLogoBase64);
    });
}

async function getValidLogoBase64() {
    return 'data:image/jpeg;base64,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';
}

async function addInstitutionalHeader(doc, subtitle, reportName, contextText) {
    const logoData = await getValidLogoBase64();
    
    doc.autoTable({
        startY: 15,
        margin: { top: 15, left: 14, right: 14 },
        theme: 'grid',
        styles: { 
            fontSize: 8, 
            textColor: [0, 0, 0], 
            lineColor: [0, 0, 0], 
            lineWidth: 0.1, 
            halign: 'center', 
            valign: 'middle' 
        },
        body: [
            [
                { content: '', rowSpan: 3, styles: { minCellHeight: 30, minCellWidth: 45 } },
                { content: 'FUNDACIÓN GRUPO DE APOYO PARA PROYECTOS\nNIT: 900.123.456-7', styles: { fontStyle: 'bold', fontSize: 10 } },
                { content: 'CODIGO FOR-GAP-003' }
            ],
            [
                { content: reportName, styles: { fontStyle: 'bold', fontSize: 10 } },
                { content: 'Versión 001\nActualizado: 01/MAYO/26' }
            ],
            [
                { content: subtitle, styles: { fontSize: 9 } },
                { content: 'Pág. 1 de 1' }
            ]
        ],
        columnStyles: {
            0: { cellWidth: 45 },
            1: { cellWidth: 'auto' },
            2: { cellWidth: 45 }
        },
        didDrawCell: function(data) {
            if (data.row.index === 0 && data.column.index === 0 && logoData && data.cell.section === 'body') {
                const dim = data.cell.height - 4;
                doc.addImage(logoData, 'JPEG', data.cell.x + (data.cell.width - dim)/2, data.cell.y + 2, dim, dim);
            }
        }
    });
    
    doc.setFontSize(9);
    doc.setFont(undefined, 'italic');
    const finalY = doc.lastAutoTable.finalY + 10;
    if (contextText) {
        doc.text(contextText, 14, finalY);
    }
    return finalY + 5;
}

window.exportDonorsExcel = function() {
    if (!AppState.donationsData || AppState.donationsData.length <= 1) return showToast('No hay datos', 'warning');
    
    // slice(1) to skip header row
    const data = AppState.donationsData.slice(1).map(d => {
        var isObj = typeof d === 'object' && !Array.isArray(d);
        return {
            "Fecha": isObj ? d.date : (d[4] || ''),
            "Donante": isObj ? d.name : (d[0] || ''),
            "Email": isObj ? d.email : (d[1] || ''),
            "Monto": isObj ? d.amount : (d[2] || ''),
            "Método": isObj ? d.method : (d[3] || ''),
            "Estado": isObj ? d.status : (d[7] || '')
        };
    });
    
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Donantes");
    XLSX.writeFile(wb, "Reporte_Donantes.xlsx");
};

window.exportDonorsPDF = async function() {
    if (!AppState.donationsData || AppState.donationsData.length === 0) return showToast('No hay datos', 'warning');
    showToast('Generando PDF...', 'info');
    
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    
    const startY = await addInstitutionalHeader(doc, "REPORTE DE DONACIONES RECIBIDAS", "GAP: EMPRENDE", "Este documento contiene el listado oficial de donaciones recibidas a la fecha.");
    
    // slice(1) to skip header row
    const tableData = AppState.donationsData.slice(1).map(d => {
        var isObj = typeof d === 'object' && !Array.isArray(d);
        return [
            isObj ? d.date : (d[4] || ''), // Fecha
            isObj ? d.name : (d[0] || ''), // Donante
            isObj ? d.amount : (d[2] || ''), // Monto
            isObj ? d.method : (d[3] || ''), // Método
            isObj ? d.status : (d[7] || 'Pendiente') // Estado
        ];
    });
    
    doc.autoTable({
        startY: startY,
        head: [['Fecha', 'Donante', 'Monto', 'Método', 'Estado']],
        body: tableData,
        theme: 'striped',
        headStyles: { fillColor: [4, 120, 87] }
    });
    
    doc.save("Reporte_Donantes.pdf");
};

window.exportVolunteersExcel = async function() {
    const res = await fetchFromSheets('GET', 'Voluntarios');
    if (!res || !res.data || res.data.length === 0) return showToast('No hay datos', 'warning');
    
    const accepted = res.data.filter(v => v[7] === 'Aceptado' || v[7] === 'Activo');
    if (accepted.length === 0) return showToast('No hay voluntarios aceptados', 'warning');
    
    const data = accepted.map(v => ({
        "Fecha": v[0] || '',
        "Nombre": v[1] || '',
        "Email": v[2] || '',
        "Teléfono": v[3] || '',
        "Habilidades": v[4] || '',
        "Disponibilidad": v[5] || '',
        "Estado": v[7] || ''
    }));
    
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Voluntarios");
    XLSX.writeFile(wb, "Reporte_Voluntarios_Aceptados.xlsx");
};

window.exportVolunteersPDF = async function() {
    showToast('Generando PDF...', 'info');
    const res = await fetchFromSheets('GET', 'Voluntarios');
    if (!res || !res.data || res.data.length === 0) return showToast('No hay datos', 'warning');
    
    const accepted = res.data.filter(v => v[7] === 'Aceptado' || v[7] === 'Activo');
    if (accepted.length === 0) return showToast('No hay voluntarios aceptados', 'warning');
    
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    
    const startY = await addInstitutionalHeader(doc, "LISTADO DE VOLUNTARIOS ACEPTADOS", "GAP: EMPRENDE", "Este documento certifica el listado oficial de voluntarios activos y aceptados por la fundación.");
    
    const tableData = accepted.map(v => [
        v[0] || '', // Fecha
        v[1] || '', // Nombre
        v[3] || '', // Teléfono
        v[4] || '', // Habilidades
        v[7] || ''  // Estado
    ]);
    
    doc.autoTable({
        startY: startY,
        head: [['Fecha', 'Nombre', 'Teléfono', 'Habilidades', 'Estado']],
        body: tableData,
        theme: 'striped',
        headStyles: { fillColor: [4, 120, 87] }
    });
    
    doc.save("Reporte_Voluntarios.pdf");
};

window.exportDashboardExcel = function() {
    const totalDonations = document.getElementById('metricDonations') ? document.getElementById('metricDonations').textContent : '0';
    const totalDonors = document.getElementById('metricDonors') ? document.getElementById('metricDonors').textContent : '0';
    
    const data = [
        { "Métrica": "Total Donaciones Recaudadas", "Valor": totalDonations },
        { "Métrica": "Donantes Únicos", "Valor": totalDonors }
    ];
    
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Dashboard");
    XLSX.writeFile(wb, "Resumen_Dashboard.xlsx");
};

window.exportDashboardPDF = async function() {
    showToast('Generando PDF...', 'info');
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    
    const startY = await addInstitutionalHeader(doc, "RESUMEN DE IMPACTO (DASHBOARD)", "GAP: EMPRENDE", "Este documento presenta un resumen oficial del impacto y las métricas clave alcanzadas por la fundación.");
    
    const totalDonations = document.getElementById('metricDonations') ? document.getElementById('metricDonations').textContent : '0';
    const totalDonors = document.getElementById('metricDonors') ? document.getElementById('metricDonors').textContent : '0';
    
    const tableData = [
        ["Total Donaciones Recaudadas", totalDonations],
        ["Donantes Únicos", totalDonors]
    ];
    
    doc.autoTable({
        startY: startY,
        head: [['Métrica', 'Valor']],
        body: tableData,
        theme: 'striped',
        headStyles: { fillColor: [4, 120, 87] }
    });
    
    doc.save("Resumen_Dashboard.pdf");
};

// =====================================================================
// AÑADIR USUARIO DESDE PANEL DE ADMIN
// =====================================================================

window.openAddUserModal = function() {
    var nameInput = document.getElementById('addUserName');
    var emailInput = document.getElementById('addUserEmail');
    var passInput = document.getElementById('addUserPassword');
    var roleInput = document.getElementById('addUserRole');
    
    if (nameInput) nameInput.value = '';
    if (emailInput) emailInput.value = '';
    if (passInput) passInput.value = '';
    if (roleInput) roleInput.value = 'Admin';
    
    var modal = document.getElementById('addUserModal');
    if (modal) {
        modal.classList.remove('hidden');
        if (typeof lucide !== 'undefined') lucide.createIcons();
    }
};

window.closeAddUserModal = function() {
    var modal = document.getElementById('addUserModal');
    if (modal) {
        modal.classList.add('hidden');
    }
};

window.submitAddUser = async function(event) {
    event.preventDefault();
    
    var name = document.getElementById('addUserName').value.trim();
    var email = document.getElementById('addUserEmail').value.trim();
    var password = document.getElementById('addUserPassword').value;
    var role = document.getElementById('addUserRole').value;
    
    if (!name || !email || !password || !role) {
        return showToast('Por favor completa todos los campos', 'error');
    }
    
    var btn = document.getElementById('btnSubmitAddUser');
    var originalBtnHtml = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = '<i data-lucide="loader" class="spin-anim"></i> Guardando...';
    if (typeof lucide !== 'undefined') lucide.createIcons();
    
    try {
        var date = new Date().toISOString();
        var newUid = 'usr_' + Date.now() + Math.floor(Math.random() * 1000);
        
        // row: [uid, name, email, password, fecha_registro, rol]
        var rowData = [newUid, name, email, password, date, role];
        
        await fetchFromSheets('POST', 'Usuarios', {
            action: 'add',
            row: rowData
        });
        
        showToast('Usuario creado con éxito', 'success');
        window.closeAddUserModal();
        
        if (window.ModernApp && window.ModernApp.adminUsers && window.ModernApp.adminUsers.loadUsersAdmin) {
            window.ModernApp.adminUsers.loadUsersAdmin();
        }
    } catch (error) {
        console.error('Error al añadir usuario:', error);
        showToast('Error al procesar la solicitud', 'error');
    } finally {
        btn.disabled = false;
        btn.innerHTML = originalBtnHtml;
        if (typeof lucide !== 'undefined') lucide.createIcons();
    }
};
