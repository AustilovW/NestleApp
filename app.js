/**
 * NestleApp - Available Quota Module
 * Displays and manages user's available quota information
 */

// Mock data for demonstration purposes
const mockQuotaData = {
    totalQuota: 50000.00,
    usedQuota: 12500.00,
    cutDate: '2025-12-15',
    currency: 'MXN'
};

/**
 * Formats a number as currency
 * @param {number} amount - The amount to format
 * @returns {string} Formatted currency string
 */
function formatCurrency(amount) {
    return new Intl.NumberFormat('es-MX', {
        style: 'currency',
        currency: 'MXN',
        minimumFractionDigits: 2
    }).format(amount);
}

/**
 * Formats a date string to localized format
 * @param {string} dateString - ISO date string
 * @returns {string} Formatted date
 */
function formatDate(dateString) {
    const date = new Date(dateString);
    return new Intl.DateTimeFormat('es-MX', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
    }).format(date);
}

/**
 * Calculates available quota
 * @param {number} total - Total quota
 * @param {number} used - Used quota
 * @returns {number} Available quota
 */
function calculateAvailableQuota(total, used) {
    return Math.max(0, total - used);
}

/**
 * Calculates usage percentage
 * @param {number} total - Total quota
 * @param {number} used - Used quota
 * @returns {number} Usage percentage (0-100)
 */
function calculateUsagePercentage(total, used) {
    if (total <= 0) return 0;
    return Math.min(100, (used / total) * 100);
}

/**
 * Gets the appropriate CSS class for progress bar based on usage
 * @param {number} percentage - Usage percentage
 * @returns {string} CSS class name
 */
function getProgressClass(percentage) {
    if (percentage >= 80) return 'danger';
    if (percentage >= 60) return 'warning';
    return '';
}

/**
 * Updates the UI with quota information
 * @param {Object} quotaData - Quota data object
 */
function updateQuotaDisplay(quotaData) {
    const availableQuota = calculateAvailableQuota(quotaData.totalQuota, quotaData.usedQuota);
    const usagePercentage = calculateUsagePercentage(quotaData.totalQuota, quotaData.usedQuota);
    
    // Update main quota display
    const availableQuotaElement = document.getElementById('available-quota');
    if (availableQuotaElement) {
        availableQuotaElement.textContent = formatCurrency(availableQuota);
    }
    
    // Update quota details
    const totalQuotaElement = document.getElementById('total-quota');
    if (totalQuotaElement) {
        totalQuotaElement.textContent = formatCurrency(quotaData.totalQuota);
    }
    
    const usedQuotaElement = document.getElementById('used-quota');
    if (usedQuotaElement) {
        usedQuotaElement.textContent = formatCurrency(quotaData.usedQuota);
    }
    
    const cutDateElement = document.getElementById('cut-date');
    if (cutDateElement) {
        cutDateElement.textContent = formatDate(quotaData.cutDate);
    }
    
    // Update progress bar
    const progressFillElement = document.getElementById('quota-progress');
    if (progressFillElement) {
        progressFillElement.style.width = usagePercentage + '%';
        progressFillElement.className = 'progress-fill ' + getProgressClass(usagePercentage);
    }
    
    const usagePercentageElement = document.getElementById('usage-percentage');
    if (usagePercentageElement) {
        usagePercentageElement.textContent = Math.round(usagePercentage) + '% utilizado';
    }
}

/**
 * Simulates fetching quota data from a server
 * @returns {Promise<Object>} Promise resolving to quota data
 */
function fetchQuotaData() {
    return new Promise((resolve) => {
        // Simulate network delay
        setTimeout(() => {
            resolve(mockQuotaData);
        }, 500);
    });
}

/**
 * Refreshes the quota display by fetching new data
 */
async function refreshQuota() {
    const refreshBtn = document.getElementById('refresh-btn');
    if (refreshBtn) {
        refreshBtn.disabled = true;
        refreshBtn.textContent = 'Actualizando...';
    }
    
    try {
        const quotaData = await fetchQuotaData();
        updateQuotaDisplay(quotaData);
    } catch (error) {
        console.error('Error fetching quota data:', error);
    } finally {
        if (refreshBtn) {
            refreshBtn.disabled = false;
            refreshBtn.textContent = 'Actualizar Cuota';
        }
    }
}

/**
 * Initializes the application
 */
function initApp() {
    // Initial load of quota data
    refreshQuota();
    
    // Set up refresh button click handler
    const refreshBtn = document.getElementById('refresh-btn');
    if (refreshBtn) {
        refreshBtn.addEventListener('click', refreshQuota);
    }
}

// Initialize app when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initApp);
} else {
    initApp();
}

// Export functions for testing (if in a module environment)
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        formatCurrency,
        formatDate,
        calculateAvailableQuota,
        calculateUsagePercentage,
        getProgressClass,
        updateQuotaDisplay,
        mockQuotaData
    };
}
