
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

        loadForumData();
        loadBankInfo(); // <-- NUEVO: Cargar informacion bancaria dinámica
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
            'admin-dashboard': 'Dashboard',
            'admin-donantes': 'Gestion de Donantes',
            'admin-contenido': 'Gestion de Contenido'
        };
        document.getElementById('pageTitle').textContent = pageTitles[pageId] || 'Inicio';

        AppState.currentPage = pageId;

        if (pageId === 'admin-dashboard' && AppState.isAdmin) loadAdminDashboard();
        if (pageId === 'admin-donantes' && AppState.isAdmin) loadDonorsTable();
        if (pageId === 'admin-contenido' && AppState.isAdmin) loadForumPostsAdmin();

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
                handleUserLogin({
                    uid:   user.uid,
                    name:  user.displayName,
                    email: user.email,
                    photo: user.photoURL,
                    provider: 'google'
                });
            } else {
                showToast('Google Sign-In no está configurado. Usa correo y contraseña.', 'error');
            }
        } catch (error) {
            console.error('Error en login con Google:', error);
            if (error.code === 'auth/popup-closed-by-user') {
                showToast('Login cancelado', 'info');
            } else {
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

        showToast('Creando tu cuenta...', 'info');

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
        checkUserRole(userData);
        updateUIForLoggedUser();
        closeLoginModal();
        localStorage.setItem('fundacion_session', JSON.stringify(userData));
        // Solo registrar en la hoja si NO es login por credenciales
        // (el backend de registerUser/loginUser ya maneja la hoja Usuarios)
        if (userData.provider !== 'credentials') {
            registerUserInSheet(userData);
        }
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
        if (AppState.isAdmin) {
            document.body.classList.add('is-admin');
        }
    }

    /** Actualiza la interfaz cuando un usuario inicia sesion */
    function updateUIForLoggedUser() {
        var user = AppState.currentUser;
        var initials = getInitials(user.name);

        document.getElementById('userAvatarSidebar').innerHTML = user.photo
            ? '<img src="' + user.photo + '" alt="' + user.name + '">'
            : initials;
        document.getElementById('userNameSidebar').textContent = user.name;
        document.getElementById('userRoleSidebar').textContent = AppState.isAdmin ? 'Administrador' : 'Miembro';

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
        try {
            if (firebaseAuth) await firebaseAuth.signOut();
        } catch (e) { /* ignorar */ }

        AppState.currentUser = null;
        AppState.isAdmin = false;
        localStorage.removeItem('fundacion_session');
        document.body.classList.remove('is-admin');

        document.getElementById('sidebarLoginPrompt').classList.remove('hidden');
        document.getElementById('sidebarUserInfo').classList.add('hidden');
        document.getElementById('topBarLoginBtn').classList.remove('hidden');
        document.getElementById('topBarUserAvatar').classList.add('hidden');
        document.getElementById('forumAvatar').textContent = '?';
        document.getElementById('forumTextarea').setAttribute('onclick', "requireLogin('comentar en el foro')");

        navigateTo('inicio');
        showToast('Sesion cerrada', 'info');
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
        if (!AppState.isAdmin) return;
        pendingDeleteIndex = index;
        pendingDeleteType = 'Foro';
        showDeleteModal();
    };

    function showDeleteModal() {
        var modal = document.getElementById('deleteConfirmModal');
        document.getElementById('deleteStateConfirm').classList.remove('hidden');
        document.getElementById('deleteStateLoading').classList.add('hidden');
        document.getElementById('deleteStateSuccess').classList.add('hidden');
        
        var titleText = pendingDeleteType === 'Foro' ? '¿Eliminar Comentario?' : '¿Eliminar Publicación?';
        var descText = pendingDeleteType === 'Foro' ? 'El comentario se eliminó correctamente.' : 'La publicación se eliminó correctamente.';
        
        document.getElementById('deleteConfirmTitle').textContent = titleText;
        document.getElementById('deleteSuccessDesc').textContent = descText;

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
            if (pendingDeleteType === 'Foro') await loadForumData();
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
        window.currentShareText = 'Mira esta labor de la Fundación Esperanza Viva: ' + title;

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
            window.location.href = 'mailto:?subject=Fundación Esperanza Viva&body=' + text + '%0A%0A' + window.currentShareUrl;
        } else if (platform === 'copy') {
            copyFallback();
        } else if (platform === 'native') {
            // Prevenir crash del navegador (RESULT_CODE_KILLED_BAD_MESSAGE) en PC al intentar
            // compartir URLs locales (file:///) usando la API nativa de Windows.
            if (navigator.share && window.location.protocol !== 'file:') {
                navigator.share({
                    title: 'Fundación Esperanza Viva',
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

    function selectAmount(el, amount) {
        document.querySelectorAll('.amount-option').forEach(function(opt) { opt.classList.remove('selected'); });
        el.classList.add('selected');
        document.getElementById('donationAmount').value = amount;
    }

    async function submitDonation(event) {
        event.preventDefault();

        var name = document.getElementById('donorName').value.trim();
        var email = document.getElementById('donorEmail').value.trim();
        var amount = document.getElementById('donationAmount').value;
        var method = document.getElementById('donationMethod').value;
        var message = document.getElementById('donationNote').value.trim();
        var fileInput = document.getElementById('donationReceipt');
        var userId = AppState.currentUser ? AppState.currentUser.uid : 'anon';

        if (!name || !email || !amount || !method || !fileInput.files[0]) {
            showToast('Por favor completa todos los campos requeridos, incluyendo el comprobante', 'error');
            return;
        }

        var btn = event.target.querySelector('button[type="submit"]');
        if (btn) {
            btn.disabled = true;
            btn.innerHTML = '<i data-lucide="loader" class="spin-anim"></i> Procesando donación...';
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
                if (btn) {
                    btn.disabled = false;
                    btn.innerHTML = '<i data-lucide="heart" width="18" height="18"></i> Registrar mi donacion';
                    if (typeof lucide !== 'undefined') lucide.createIcons();
                }
                return;
            }
        } catch (err) {
            showToast('Error leyendo el archivo', 'error');
            if (btn) {
                btn.disabled = false;
                btn.innerHTML = '<i data-lucide="heart" width="18" height="18"></i> Registrar mi donacion';
                if (typeof lucide !== 'undefined') lucide.createIcons();
            }
            return;
        }

        showToast('Registrando donación...', 'info');
        var date = new Date().toLocaleDateString('es-CO');

        var result = await fetchFromSheets('POST', 'Donaciones', {
            action: 'add',
            row: [name, email, amount, method, date, message, userId, 'Pendiente', comprobanteUrl]
        });

        if (result && result.success !== false) {
            showToast('Donacion registrada exitosamente! Gracias por tu generosidad!', 'success');
            document.getElementById('donationForm').reset();
            document.getElementById('donationAmount').value = '100000';
            document.querySelectorAll('.amount-option').forEach(function(opt) { opt.classList.remove('selected'); });
            document.querySelectorAll('.amount-option')[2].classList.add('selected');

            if (AppState.currentUser) {
                document.getElementById('donorName').value = AppState.currentUser.name;
                document.getElementById('donorEmail').value = AppState.currentUser.email;
            }
        } else {
            showToast('Error al registrar la donacion. Intenta de nuevo.', 'error');
        }

        if (btn) {
            btn.disabled = false;
            btn.innerHTML = '<i data-lucide="heart" width="18" height="18"></i> Registrar mi donacion';
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }
    }

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
        document.getElementById('editTitular').value = data['Titular'] || 'Fundacion Esperanza Viva';
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
        var response = await fetchFromSheets('GET', 'Foro');
        var forumPosts = document.getElementById('forumPosts');
        var forumEmpty = document.getElementById('forumEmpty');

        if (response && response.data && response.data.length > 0) {
            AppState.forumData = response.data;
            forumPosts.innerHTML = '';
            forumEmpty.classList.add('hidden');

            var avatarColors = ['#059669', '#0891b2', '#7c3aed', '#db2777', '#ea580c', '#0284c7'];

            response.data.forEach(function(post, index) {
                var isObj = typeof post === 'object' && !Array.isArray(post);
                var author = isObj ? post.author : (post[0] || 'Anonimo');
                var message = isObj ? post.message : (post[1] || '');
                var date = isObj ? post.date : (post[2] || '');
                var avatar = isObj ? post.avatar : (post[4] || getInitials(author));
                var colorIndex = author.charCodeAt(0) % avatarColors.length;

                var postEl = document.createElement('div');
                postEl.className = 'forum-post';
                postEl.id = 'forum-post-' + index;
                postEl.innerHTML =
                    '<div class="forum-post-header">' +
                        '<div class="forum-post-avatar" style="background:' + avatarColors[colorIndex] + '; color:#fff;">' + avatar + '</div>' +
                        '<div class="forum-post-info">' +
                            '<div class="post-author">' + escapeHtml(author) + '</div>' +
                            '<div class="post-date">' + formatDate(date) + '</div>' +
                        '</div>' +
                        '<div class="mod-actions">' +
                            '<button class="mod-btn delete" onclick="deleteForumPost(' + index + ')">Eliminar</button>' +
                        '</div>' +
                    '</div>' +
                    '<div class="forum-post-content">' + escapeHtml(message) + '</div>' +
                    '<div class="forum-post-footer">' +
                        '<button class="post-action-btn" onclick="likeForumPost(this)">' +
                            '<i data-lucide="thumbs-up" width="14" height="14"></i> Me gusta' +
                        '</button>' +
                        '<button class="post-action-btn" onclick="showToast(\'Funcion de respuesta proximamente\', \'info\')">' +
                            '<i data-lucide="reply" width="14" height="14"></i> Responder' +
                        '</button>' +
                    '</div>';
                forumPosts.appendChild(postEl);
            });

            var badge = document.getElementById('forumBadge');
            badge.textContent = response.data.length;
            badge.style.display = 'inline';
        } else {
            forumPosts.innerHTML = '';
            forumEmpty.classList.remove('hidden');
        }

        setTimeout(function() {
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }, 150);
    }

    async function submitForumPost() {
        if (!requireLogin('publicar en el foro')) return;

        var textarea = document.getElementById('forumTextarea');
        var message = textarea.value.trim();

        if (!message) {
            showToast('Escribe un mensaje antes de publicar', 'error');
            return;
        }

        var user = AppState.currentUser;
        var date = new Date().toLocaleDateString('es-CO');
        var initials = getInitials(user.name);

        var result = await fetchFromSheets('POST', 'Foro', {
            action: 'add',
            row: [user.name, message, date, user.uid, initials]
        });

        if (result && result.success !== false) {
            textarea.value = '';
            showToast('Comentario publicado!', 'success');
            await loadForumData();
        } else {
            showToast('Error al publicar el comentario', 'error');
        }
    }

    // (La función de eliminar foro ahora usa la misma modal: deleteForumPost)

    function likeForumPost(btn) {
        if (!requireLogin('dar Me gusta a un comentario')) return;
        btn.classList.toggle('liked');
        if (btn.classList.contains('liked')) {
            btn.style.color = 'var(--primary-600)';
            btn.style.background = 'var(--primary-50)';
        } else {
            btn.style.color = '';
            btn.style.background = '';
        }
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

        // Modificar el encabezado para agregar ACCIONES si no existe
        var thead = document.querySelector('.table-responsive table thead tr');
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

