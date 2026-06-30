// js/views/forum.js
import { State } from '../core/state.js';
import { apiGet, apiPost } from '../services/api.js';

let forumData = [];
const avatarColors = ['#059669', '#0891b2', '#7c3aed', '#db2777', '#ea580c', '#0284c7'];

function getInitials(name) {
    if (!name) return '?';
    return name.substring(0, 2).toUpperCase();
}

function escapeHtml(text) {
    if (!text) return '';
    const map = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' };
    return String(text).replace(/[&<>"']/g, function(m) { return map[m]; });
}

export async function loadForumData() {
    try {
        const data = await apiGet('Foro');
        forumData = data;
        renderThreads();
    } catch (error) {
        console.error('Error al cargar foro:', error);
    }
}

function renderThreads() {
    const forumPosts = document.getElementById('forumPosts');
    const forumEmpty = document.getElementById('forumEmpty');
    const badge = document.getElementById('forumBadge');
    
    if (forumData.length === 0) {
        forumPosts.innerHTML = '';
        forumEmpty.classList.remove('hidden');
        if (badge) badge.style.display = 'none';
        return;
    }
    
    forumEmpty.classList.add('hidden');
    if (badge) {
        badge.textContent = forumData.length;
        badge.style.display = 'inline';
    }
    
    forumPosts.innerHTML = '';
    
    // Organizar por ParentID. Las que no tienen ParentID son hilos principales.
    const threads = {};
    const replies = {};
    
    // Asumimos formato: [Autor, Mensaje, Fecha, UserID, AvatarInitial, ParentID, ID]
    // Pero la hoja Foro actual quizás no devuelva ID. Para evitar complicaciones de IDs, 
    // usaremos el index original del arreglo de Google Sheets (row = index + 2).
    
    forumData.forEach((post, index) => {
        post.originalIndex = index;
        const isObj = typeof post === 'object' && !Array.isArray(post);
        const parentId = isObj ? post.parentId : (post[5] !== undefined && post[5] !== '' ? post[5] : null);
        
        if (parentId === null) {
            threads[index] = { post, children: [] };
        } else {
            if (!replies[parentId]) replies[parentId] = [];
            replies[parentId].push(post);
        }
    });
    
    // Unir
    for (let pId in replies) {
        if (threads[pId]) {
            threads[pId].children = replies[pId];
        }
    }
    
    // Renderizar
    Object.values(threads).reverse().forEach(threadObj => {
        const post = threadObj.post;
        const numReplies = threadObj.children.length;
        const postEl = createPostElement(post, false, numReplies);
        
        // Contenedor de respuestas
        const repliesContainerWrapper = document.createElement('div');
        repliesContainerWrapper.id = 'replies-wrapper-' + post.originalIndex;
        repliesContainerWrapper.className = 'hidden';
        repliesContainerWrapper.setAttribute('data-count', numReplies);

        const repliesContainer = document.createElement('div');
        repliesContainer.className = 'forum-replies';
        repliesContainer.style.marginLeft = '40px';
        repliesContainer.style.marginTop = '10px';
        repliesContainer.style.borderLeft = '2px solid var(--border)';
        repliesContainer.style.paddingLeft = '15px';
        repliesContainer.id = 'replies-' + post.originalIndex;
        
        threadObj.children.forEach(replyPost => {
            repliesContainer.appendChild(createPostElement(replyPost, true, 0));
        });
        
        // Caja de respuesta
        const replyBox = document.createElement('div');
        replyBox.className = 'reply-box hidden';
        replyBox.id = 'reply-box-' + post.originalIndex;
        replyBox.innerHTML = `
            <div style="display:flex; gap:10px; margin-top:10px;">
                <input type="text" id="reply-input-${post.originalIndex}" class="form-input" placeholder="Escribe una respuesta..." style="flex:1;">
                <button class="btn btn-primary" onclick="window.ModernApp.forum.submitForumReply(${post.originalIndex})">Enviar</button>
            </div>
        `;
        repliesContainer.appendChild(replyBox);
        repliesContainerWrapper.appendChild(repliesContainer);
        
        postEl.appendChild(repliesContainerWrapper);
        forumPosts.appendChild(postEl);
    });
    
    if (window.lucide) window.lucide.createIcons();
}

window.ModernApp = window.ModernApp || {};
window.ModernApp.forum = window.ModernApp.forum || {};

window.ModernApp.forum.toggleReplies = function(index) {
    const wrapper = document.getElementById('replies-wrapper-' + index);
    const textSpan = document.getElementById('toggle-replies-text-' + index);
    if (!wrapper) return;
    
    if (wrapper.classList.contains('hidden')) {
        wrapper.classList.remove('hidden');
        if (textSpan) textSpan.textContent = 'Ocultar respuestas';
    } else {
        wrapper.classList.add('hidden');
        if (textSpan) {
            const count = wrapper.getAttribute('data-count');
            textSpan.textContent = `Ver ${count} respuesta${count > 1 ? 's' : ''}`;
        }
    }
};

window.ModernApp.forum.showReplyBox = function(index) {
    const wrapper = document.getElementById('replies-wrapper-' + index);
    if (wrapper && wrapper.classList.contains('hidden')) {
        window.ModernApp.forum.toggleReplies(index);
    }
    const replyBox = document.getElementById('reply-box-' + index);
    if (replyBox) {
        replyBox.classList.remove('hidden');
        const input = document.getElementById('reply-input-' + index);
        if (input) input.focus();
    }
};

window.ModernApp.forum.likePost = async function(index) {
    if (!window.AppState || !window.AppState.currentUser) {
        if (window.requireLogin) window.requireLogin('dar me gusta');
        return;
    }
    const uid = window.AppState.currentUser.uid;
    const post = forumData[index];
    if (!post) return;

    let likesCount = parseInt(post[6]) || 0;
    let likedBy = [];
    try { likedBy = JSON.parse(post[7] || '[]'); } catch(e) {}

    const isLiked = likedBy.includes(uid);
    if (isLiked) {
        likedBy = likedBy.filter(id => id !== uid);
        likesCount = Math.max(0, likesCount - 1);
    } else {
        likedBy.push(uid);
        likesCount++;
    }

    // Update local data
    post[6] = likesCount;
    post[7] = JSON.stringify(likedBy);

    // Update UI
    const likeBtn = document.getElementById('like-btn-' + index);
    const countSpan = document.getElementById('like-count-' + index);
    if (likeBtn) {
        if (!isLiked) {
            likeBtn.classList.add('liked');
            likeBtn.style.color = 'var(--p600)';
        } else {
            likeBtn.classList.remove('liked');
            likeBtn.style.color = '';
        }
    }
    if (countSpan) countSpan.textContent = likesCount > 0 ? likesCount : '';

    // Send to backend
    try {
        await apiPost({
            action: 'updateRow',
            sheet: 'Foro',
            rowIndex: index,
            row: post
        });
    } catch(e) {
        console.error("Error al guardar like:", e);
    }
};

function createPostElement(post, isReply, numReplies = 0) {
    const isObj = typeof post === 'object' && !Array.isArray(post);
    const author = isObj ? post.author : (post[0] || 'Anonimo');
    const message = isObj ? post.message : (post[1] || '');
    const date = isObj ? post.date : (post[2] || '');
    const uid = isObj ? post.userId : (post[3] || '');
    const avatar = isObj ? post.avatar : (post[4] || getInitials(author));
    const parentId = isObj ? post.parentId : (post[5] || '');
    const likesCount = isObj ? post.likes : parseInt(post[6]) || 0;
    let likedBy = [];
    try { likedBy = JSON.parse(post[7] || '[]'); } catch(e) {}
    
    const colorIndex = author.charCodeAt(0) % avatarColors.length;
    const index = post.originalIndex;
    
    const currentUserUid = window.AppState?.currentUser?.uid;
    const isLiked = currentUserUid && likedBy.includes(currentUserUid);
    
    const postEl = document.createElement('div');
    postEl.className = 'forum-post';
    postEl.style.marginBottom = isReply ? '10px' : '20px';
    if (isReply) postEl.style.padding = '10px';
    
    let deleteBtn = '';
    const user = window.AppState ? window.AppState.currentUser : null;
    if (user && (window.AppState.isAdmin || user.email === author)) {
        deleteBtn = `
        <div class="mod-actions">
            <button class="mod-btn delete" onclick="window.deleteForumPost(${index})">Eliminar</button>
        </div>`;
    }
    
    let repliesToggleHtml = '';
    if (!isReply && numReplies > 0) {
        repliesToggleHtml = `
        <button id="toggle-replies-btn-${index}" class="post-action-btn" onclick="window.ModernApp.forum.toggleReplies(${index})">
            <i data-lucide="message-square" width="14" height="14"></i> <span id="toggle-replies-text-${index}">Ver ${numReplies} respuesta${numReplies > 1 ? 's' : ''}</span>
        </button>`;
    }
    
    postEl.innerHTML = `
        <div class="forum-post-header">
            <div class="forum-post-avatar" style="background:${avatarColors[colorIndex]}; color:#fff;">${avatar}</div>
            <div class="forum-post-info">
                <div class="post-author">${escapeHtml(author)}</div>
                <div class="post-date">${date}</div>
            </div>
            ${deleteBtn}
        </div>
        <div class="forum-post-content">${escapeHtml(message)}</div>
        <div class="forum-post-footer">
            <button id="like-btn-${index}" class="post-action-btn ${isLiked ? 'liked' : ''}" style="${isLiked ? 'color: var(--p600);' : ''}" onclick="window.ModernApp.forum.likePost(${index})">
                <i data-lucide="thumbs-up" width="14" height="14"></i> Me gusta <span id="like-count-${index}">${likesCount > 0 ? likesCount : ''}</span>
            </button>
            ${repliesToggleHtml}
            ${!isReply ? `
            <button class="post-action-btn" onclick="window.ModernApp.forum.showReplyBox(${index})">
                <i data-lucide="reply" width="14" height="14"></i> Responder
            </button>` : ''}
        </div>
    `;
    return postEl;
}

export async function submitForumPost() {
    const user = window.AppState ? window.AppState.currentUser : null;
    if (!user) {
        if (window.requireLogin) window.requireLogin('publicar en el foro');
        return;
    }
    
    const textarea = document.getElementById('forumTextarea');
    const message = textarea.value.trim();
    if (!message) return;
    
    showLoadingModal('Subiendo comentario...');
    
    const date = new Date().toLocaleDateString('es-CO');
    textarea.value = '';
    
    await apiPost({
        action: 'add',
        sheet: 'Foro',
        row: [user.name, message, date, user.uid, getInitials(user.name), '']
    });
    
    finishLoadingModal();
    loadForumData();
}

export async function submitForumReply(parentId) {
    const user = window.AppState ? window.AppState.currentUser : null;
    if (!user) {
        if (window.requireLogin) window.requireLogin('responder en el foro');
        return;
    }
    
    const input = document.getElementById('reply-input-' + parentId);
    const message = input.value.trim();
    if (!message) return;
    
    showLoadingModal('Subiendo respuesta...');
    
    const date = new Date().toLocaleDateString('es-CO');
    input.value = '';
    
    await apiPost({
        action: 'add',
        sheet: 'Foro',
        row: [user.name, message, date, user.uid, getInitials(user.name), parentId.toString()]
    });
    
    finishLoadingModal();
    loadForumData();
}

function showLoadingModal(titleText) {
    let modal = document.getElementById('forumLoadingModal');
    if (!modal) {
        modal = document.createElement('div');
        modal.id = 'forumLoadingModal';
        modal.className = 'post-viewer-overlay hidden';
        modal.style.cssText = 'z-index: 10001; align-items: center; justify-content: center; display: flex;';
        modal.innerHTML = `
            <div class="delete-modal-content" style="background: var(--bg1); border-radius: var(--r4); padding: 40px 32px; width: 90%; max-width: 450px; text-align: center; box-shadow: var(--sh4); transform: scale(1) !important; opacity: 1 !important;">
                
                <div id="forumStateLoading">
                    <div class="delete-icon loading" style="margin-bottom: 20px; color: var(--primary-500);">
                        <i data-lucide="loader-2" class="spin-anim" width="48" height="48"></i>
                    </div>
                    <h2 id="forumLoadingTitle" style="color: var(--txt1); font-size: 1.5rem; margin-bottom: 8px;">Subiendo comentario...</h2>
                    <p style="color: var(--txt2);">Por favor, espera un momento.</p>
                </div>
                
                <div id="forumStateSuccess" class="hidden">
                    <div class="delete-icon success" style="margin-bottom: 20px; color: #10b981;">
                        <i data-lucide="check-circle" width="48" height="48"></i>
                    </div>
                    <h2 style="color: var(--txt1); font-size: 1.5rem; margin-bottom: 8px;">Comentario compartido</h2>
                    <p style="color: var(--txt2); margin-bottom: 24px;">Tu comentario ha sido publicado exitosamente.</p>
                    <button onclick="document.getElementById('forumLoadingModal').classList.add('hidden');" class="btn btn-primary" style="width: 100%;">Okay</button>
                </div>
                
            </div>
        `;
        document.body.appendChild(modal);
    }
    
    document.getElementById('forumStateLoading').classList.remove('hidden');
    document.getElementById('forumStateSuccess').classList.add('hidden');
    document.getElementById('forumLoadingTitle').textContent = titleText;
    
    setTimeout(() => {
        modal.classList.remove('hidden');
        if (typeof window.lucide !== 'undefined') window.lucide.createIcons();
    }, 10);
}

function finishLoadingModal() {
    const modal = document.getElementById('forumLoadingModal');
    if (modal) {
        document.getElementById('forumStateLoading').classList.add('hidden');
        document.getElementById('forumStateSuccess').classList.remove('hidden');
    }
}
