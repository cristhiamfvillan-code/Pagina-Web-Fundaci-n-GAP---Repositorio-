// js/services/api.js

const GOOGLE_SHEETS_API_URL = 'https://script.google.com/macros/s/AKfycbx1AYeZGDxjF5LGKKFSC7EXmr0CwEpUQuT7gnEIIQF4MIIj9jhmoZUg15tjKd2vsfSJOg/exec';

/**
 * Función genérica para enviar peticiones POST al backend.
 * @param {Object} data 
 * @returns {Promise<Object>}
 */
export async function apiPost(data) {
    try {
        const response = await fetch(GOOGLE_SHEETS_API_URL, {
            method: 'POST',
            body: JSON.stringify(data),
            headers: {
                'Content-Type': 'text/plain;charset=utf-8',
            }
        });
        return await response.json();
    } catch (error) {
        console.error('Error en apiPost:', error);
        return { success: false, error };
    }
}

/**
 * Función genérica para obtener datos de una hoja (GET).
 * @param {string} sheetName 
 * @returns {Promise<Array>}
 */
export async function apiGet(sheetName) {
    try {
        const response = await fetch(`${GOOGLE_SHEETS_API_URL}?sheet=${sheetName}`);
        const result = await response.json();
        if (result.data) {
            return result.data;
        }
        return [];
    } catch (error) {
        console.error('Error en apiGet:', error);
        return [];
    }
}
