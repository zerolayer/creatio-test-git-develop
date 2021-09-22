(function() {
	require.config({
		paths: {
			"DriverJs": Terrasoft.getFileContentUrl("Base", "src/js/driver.min.js"),
			"DriverCSS": Terrasoft.getFileContentUrl("Base", "src/css/driver.min.css"),
		}
	});
})();
