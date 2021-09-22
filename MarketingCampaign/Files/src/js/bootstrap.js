(function() {
    require.config({
        paths: {
            "EmailTemplateLookupGalleryElement": Terrasoft.getFileContentUrl("MarketingCampaign", "src/js/marketing-campaign-ng-elements.js")
		},
        shim: {
            "EmailTemplateLookupGalleryElement": {
                deps: ["ng-core"]
            }
        }
    });
})();