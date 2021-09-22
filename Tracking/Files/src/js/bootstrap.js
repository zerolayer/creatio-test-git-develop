(function() {
    require.config({
        paths: {
            "TrackingEventsFeedElement": Terrasoft.getFileContentUrl("Tracking", "src/js/tracking-ng-elements.js")
		},
        shim: {
            "TrackingEventsFeedElement": {
                deps: ["ng-core"]
            }
        }
    });
})();