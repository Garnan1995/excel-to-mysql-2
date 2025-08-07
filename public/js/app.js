// ===================================================================
// js/app.js - Application entry point
// ===================================================================
document.addEventListener('DOMContentLoaded', async () => {
    console.log('Starting Excel Dashboard...');
    
    const dashboard = new ExcelDashboard({
        apiUrl: 'api/excel-api.php',
        maxRows: 100,
        maxCols: 50
    });
    
    // Make available globally for toolbar buttons
    window.ExcelDashboard = dashboard;
    
    try {
        await dashboard.init();
    } catch (error) {
        console.error('Failed to start dashboard:', error);
    }
});