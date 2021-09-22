(function() {
    require.config({
        paths: {
            "PivotTableComponent": Terrasoft.getFileContentUrl("PivotTable", "src/js/pivot-table-component.js"),
           
        },
        shim: {
            "PivotTableComponent": {
                deps: ["ng-core"]
            }
        }
    });
})();
