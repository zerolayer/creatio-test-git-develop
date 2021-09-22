(function() {
	require.config({
		paths: {
			"OmnichannelMessagingComponent": Terrasoft.getFileContentUrl("OmnichannelMessaging", "src/js/omnichannel-messaging-component.js"),
		},
		shim: {
			"OmnichannelMessagingComponent": {
				deps: ["ng-core"]
			}
		}
	});
})();