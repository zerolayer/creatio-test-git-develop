define("InvoiceProductPageV2", [], function() {
	return {
		entitySchemaName: "InvoiceProduct",
		attributes: {
            "Product": {
                lookupListConfig: {
                    columns: ["Price"]
                }        
            },
            "PrimaryAmount" :{
                dependencies: [
                    {
                        columns: ["Product"],
                        methodName: "calcPrimaryAmount"
                    }
                ]   
            }
		},
		modules: /**SCHEMA_MODULES*/{}/**SCHEMA_MODULES*/,
		details: /**SCHEMA_DETAILS*/{}/**SCHEMA_DETAILS*/,
		businessRules: /**SCHEMA_BUSINESS_RULES*/{}/**SCHEMA_BUSINESS_RULES*/,
		methods: {
             // Sets PrimaryAmount by Product Price
             calcPrimaryAmount: function() {
                const product = this.get("Product");
                if(product != undefined && product != null && product.value && product.Price > 0) {
                    this.set("PrimaryAmount", product.Price);
                }
            },
        },
		dataModels: /**SCHEMA_DATA_MODELS*/{}/**SCHEMA_DATA_MODELS*/,
		diff: /**SCHEMA_DIFF*/[]/**SCHEMA_DIFF*/
	};
});