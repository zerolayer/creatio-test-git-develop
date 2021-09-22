(function() {
    require.config({
        paths: {
            "ForecastComponent": Terrasoft.getFileContentUrl("CoreForecast", "src/js/forecast-component.js"),
           
        },
        shim: {
            "ForecastComponent": {
                deps: ["ng-core"]
            }
        }
    });
})();
