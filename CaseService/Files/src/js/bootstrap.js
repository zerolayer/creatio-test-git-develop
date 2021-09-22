(function() {
	require.config({
		paths: {
			"TermCalculationComponent": Terrasoft.getFileContentUrl("CaseService", "src/js/term-calculation-component.js"),
		},
		shim: {
			"TermCalculationComponent": {
				deps: ["ng-core"]
			}
		}
	});
})();