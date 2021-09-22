(function() {
    require.config({
        paths: {
            "VanillaboxMin" :Terrasoft.getFileContentUrl("UIv2", "src/js/jquery.vanillabox-0.1.7.min.js"),
            "VanillaboxCSS" :Terrasoft.getFileContentUrl("UIv2", "src/css/vanillabox.css"),
            "OrgChart" :Terrasoft.getFileContentUrl("UIv2", "src/js/jquery.orgchart.min.js"),
            "OrgChartCSS" :Terrasoft.getFileContentUrl("UIv2", "src/css/jquery.orgchart.min.css")
        },
        shim: {
            "VanillaboxMin": {
                deps: ["jQuery"]
            },
            "OrgChart": {
                deps: ["jQuery"]
            }
        }
    });
})();