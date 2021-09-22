(function () {
	require.config({
		paths: {
			"SimpleDiagram": Terrasoft.getFileContentUrl("CampaignDesigner", "src/js/cytoscape.min.js"),
			"CampaignDesignerComponent": Terrasoft.getFileContentUrl("CampaignDesigner",
				"src/js/campaign-designer-component/campaign-designer-component.js"),
			"SvgRenderer": Terrasoft.getFileContentUrl("CampaignDesigner", "src/js/canvg.min.js"),
			"CampaignGallery": Terrasoft.getFileContentUrl("CampaignDesigner",
				"src/js/campaign-gallery/campaign-gallery-element.js")
		},
		shim: {
			"SimpleDiagram": {
				deps: [""]
			},
			"CampaignDesignerComponent": {
				deps: ["ng-core"]
			},
			"CampaignGallery": {
				deps: ["ng-core"]
			},
			"SvgRenderer": {
				deps: [""]
			}
		}
	});
})();
