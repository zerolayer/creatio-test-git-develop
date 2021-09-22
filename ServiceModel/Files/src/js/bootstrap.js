(function() {
	require.config({
		paths: {
			"ServiceModelNetworkComponent": Terrasoft.getFileContentUrl("ServiceModel", "src/js/service-model-network-component.js"),
		},
		shim: {
			"ServiceModelNetworkComponent": {
				deps: ["ng-core"]
			}
		}
	});
})();