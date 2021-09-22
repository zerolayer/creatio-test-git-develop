define("PC1Page", [], function() {
	return {
		entitySchemaName: "PC",
		attributes: {
			"FirstCommitField": {
				dataValueType: this.Terrasoft.DataValueType.TEXT,
				type: this.Terrasoft.ViewModelColumnType.VIRTUAL_COLUMN,
				value: "test-field-value",
				caption: "tesdt-filed-22"
			},
			"SecondCommitField": {
				dataValueType: this.Terrasoft.DataValueType.TEXT,
				type: this.Terrasoft.ViewModelColumnType.VIRTUAL_COLUMN,
				value: "test-field-value",
				caption: "tesdt-filed-33"
			}
		},
		modules: /**SCHEMA_MODULES*/{}/**SCHEMA_MODULES*/,
		details: /**SCHEMA_DETAILS*/{
			"Files": {
				"schemaName": "FileDetailV2",
				"entitySchemaName": "PCFile",
				"filter": {
					"masterColumn": "Id",
					"detailColumn": "PC"
				}
			},
			"Schema8820641aDetail6cf68247": {
				"schemaName": "Schema8820641aDetail",
				"entitySchemaName": "PartInPCV2",
				"filter": {
					"detailColumn": "PC",
					"masterColumn": "Id"
				}
			}
		}/**SCHEMA_DETAILS*/,
		businessRules: /**SCHEMA_BUSINESS_RULES*/{}/**SCHEMA_BUSINESS_RULES*/,
		methods: {},
		dataModels: /**SCHEMA_DATA_MODELS*/{}/**SCHEMA_DATA_MODELS*/,
		diff: /**SCHEMA_DIFF*/[
			{
				"operation": "insert",
				"name": "FirstCommitField",
				"values": {
					"layout": {
						"colSpan": 24,
						"rowSpan": 1,
						"column": 0,
						"row": 2,
						"layoutName": "ProfileContainer"
					},
					"bindTo": "FirstCommitField" 
				}
			},
			{
				"operation": "insert",
				"name": "Name79a983f7-217c-41f1-925b-828300a82eea",
				"values": {
					"layout": {
						"colSpan": 24,
						"rowSpan": 1,
						"column": 0,
						"row": 0,
						"layoutName": "ProfileContainer"
					},
					"bindTo": "Name"
				},
				"parentName": "ProfileContainer",
				"propertyName": "items",
				"index": 0
			},
			{
				"operation": "insert",
				"name": "Tab31ebad12TabLabel",
				"values": {
					"caption": {
						"bindTo": "Resources.Strings.Tab31ebad12TabLabelTabCaption"
					},
					"items": [],
					"order": 0
				},
				"parentName": "Tabs",
				"propertyName": "tabs",
				"index": 0
			},
			{
				"operation": "insert",
				"name": "Tab31ebad12TabLabelGroupe3f7e190",
				"values": {
					"caption": {
						"bindTo": "Resources.Strings.Tab31ebad12TabLabelGroupe3f7e190GroupCaption"
					},
					"itemType": 15,
					"markerValue": "added-group",
					"items": []
				},
				"parentName": "Tab31ebad12TabLabel",
				"propertyName": "items",
				"index": 0
			},
			{
				"operation": "insert",
				"name": "Tab31ebad12TabLabelGridLayout1a5c0dfd",
				"values": {
					"itemType": 0,
					"items": []
				},
				"parentName": "Tab31ebad12TabLabelGroupe3f7e190",
				"propertyName": "items",
				"index": 0
			},
			{
				"operation": "insert",
				"name": "Account49c5b48f-0eb8-4d10-b5b7-583156292374",
				"values": {
					"layout": {
						"colSpan": 12,
						"rowSpan": 1,
						"column": 0,
						"row": 0,
						"layoutName": "Tab31ebad12TabLabelGridLayout1a5c0dfd"
					},
					"bindTo": "Account"
				},
				"parentName": "Tab31ebad12TabLabelGridLayout1a5c0dfd",
				"propertyName": "items",
				"index": 0
			},
			{
				"operation": "insert",
				"name": "Schema8820641aDetail6cf68247",
				"values": {
					"itemType": 2,
					"markerValue": "added-detail"
				},
				"parentName": "Tab31ebad12TabLabel",
				"propertyName": "items",
				"index": 1
			},
			{
				"operation": "insert",
				"name": "NotesAndFilesTab",
				"values": {
					"caption": {
						"bindTo": "Resources.Strings.NotesAndFilesTabCaption"
					},
					"items": [],
					"order": 1
				},
				"parentName": "Tabs",
				"propertyName": "tabs",
				"index": 1
			},
			{
				"operation": "insert",
				"name": "Files",
				"values": {
					"itemType": 2
				},
				"parentName": "NotesAndFilesTab",
				"propertyName": "items",
				"index": 0
			},
			{
				"operation": "insert",
				"name": "NotesControlGroup",
				"values": {
					"itemType": 15,
					"caption": {
						"bindTo": "Resources.Strings.NotesGroupCaption"
					},
					"items": []
				},
				"parentName": "NotesAndFilesTab",
				"propertyName": "items",
				"index": 1
			},
			{
				"operation": "insert",
				"name": "Notes",
				"values": {
					"bindTo": "Notes",
					"dataValueType": 1,
					"contentType": 4,
					"layout": {
						"column": 0,
						"row": 0,
						"colSpan": 24
					},
					"labelConfig": {
						"visible": false
					},
					"controlConfig": {
						"imageLoaded": {
							"bindTo": "insertImagesToNotes"
						},
						"images": {
							"bindTo": "NotesImagesCollection"
						}
					}
				},
				"parentName": "NotesControlGroup",
				"propertyName": "items",
				"index": 0
			},
			{
				"operation": "merge",
				"name": "ESNTab",
				"values": {
					"order": 2
				}
			}
		]/**SCHEMA_DIFF*/
	};
});
