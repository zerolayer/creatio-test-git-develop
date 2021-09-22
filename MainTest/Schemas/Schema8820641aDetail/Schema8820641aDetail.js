define("Schema8820641aDetail", [], function() {
	return {
		entitySchemaName: "PartInPCV2",
		details: /**SCHEMA_DETAILS*/{}/**SCHEMA_DETAILS*/,
		diff: /**SCHEMA_DIFF*/[]/**SCHEMA_DIFF*/,
		methods: {
			onDeleted: function() {
				this.callParents(arguments);
				this.Terrasoft.showInformation();
			}
		}
	};
});
