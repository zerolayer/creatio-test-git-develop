(function() {
    require.config({
        paths: {
            "StructureExplorerComponent": Terrasoft.getFileContentUrl("NUI", "src/js/structure-explorer-component.js"),
			"ErrorListDialogComponent": Terrasoft.getFileContentUrl("NUI", "src/js/error-list-dialog-component.js")
        },
        shim: {
            "StructureExplorerComponent": {
                deps: ["ng-core"]
            },
			"ErrorListDialogComponent": {
                deps: ["ng-core"]
            }
        }
    });
})();
