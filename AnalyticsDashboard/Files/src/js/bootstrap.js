(function() {
    require.config({
        paths: {
            "AnalyticsDashboard": Terrasoft.getFileContentUrl("AnalyticsDashboard", "src/js/analytics-dashboard.js"),
           
        },
        shim: {
            "AnalyticsDashboard": {
                deps: ["ng-core"]
            }
        }
    });
})();
