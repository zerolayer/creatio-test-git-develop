(function() {
    require.config({
        paths: {
            "MjmlCore": Terrasoft.getFileContentUrl("ContentBuilder", "src/js/mjml.min.js"),
            "CSSLint": Terrasoft.getFileContentUrl("ContentBuilder", "src/js/csslint.min.js"),
            "HTMLHint": Terrasoft.getFileContentUrl("ContentBuilder", "src/js/htmlhint.min.js")
        },
        shim: {
            "MjmlCore": {
                deps: [""]
            },
            "CSSLint": {
                deps: [""]
            },
            "HTMLHint": {
                deps: ["CSSLint"]
            }
        }
    });
})();