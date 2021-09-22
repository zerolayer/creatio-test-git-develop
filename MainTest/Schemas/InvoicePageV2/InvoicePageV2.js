 define("InvoicePageV2", ["BusinessRuleModule", "InvoiceConfigurationConstants"],
	function(BusinessRuleModule, InvoiceConfigurationConstants) {
		return {
			entitySchemaName: "Invoice",
			methods: {
				onDetailChanged: function() {
					this.Terrasoft.showInformation("Test");
				}
			},
			details: /**SCHEMA_DETAILS*/{}/**SCHEMA_DETAILS*/,
			diff: /**SCHEMA_DIFF*/[]/**SCHEMA_DIFF*/
		};
	}
);
