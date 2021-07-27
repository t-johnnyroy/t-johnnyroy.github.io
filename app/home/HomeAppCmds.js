/// <reference path="../App.js" />

var i = 0;
var k = 0;
var j = 100;
(function () {
	"use strict";

	// The initialize function must be run each time a new page is loaded.
	Office.initialize = function (reason) {
		// Office.onReady(function (reason) {
		// Office.onReady()
		// .then(function() {
		$(document).ready(function () {
			app.initialize();

			$("#writeDataBtn").click(function (event) {
				writeData();
			});
			$("#writeDataBindingBtn").click(function (event) {
				writeDataBindingBtn();
			});

			$("#readDataBtn").click(function (event) {
				readData();
			});

			$("#enableCmdBtn").click(function (event) {
				enableCommandButtons();
			});
			$("#disableCmdBtn").click(function (event) {
				disableCommandButtons();
			});

			// Ribbon
			$("#showRibbonTabBtn").click(function (event) {
				showContextualTabs();
			});

			$("#hideRibbonTabBtn").click(function (event) {
				hideContextualTabs();
			});

			$("#generateRibbonBtn").click(function (event) {
				generateRibbonButton();
			});

			$("#generateRandomRibbonDefinitionBtn").click(function (event) {
				generateRandomRibbonDefinition();
			});

			$("#addSelectionChangedEvent").click(function (event) {
				addSelectionChangedEvent();
			});
			$("#removeSelectionChangedEvent").click(function (event) {
				removeSelectionChangedEvent();
			});

			$("#bindDataBtn").click(function (event) {
				bindData();
			});
			$("#getbindsBtn").click(function (event) {
				getallbindings();
			});

			$("#readBoundDataBtn").click(function (event) {
				readBoundData();
			});
			$("#setBoundDataBtn").click(function (event) {
				setBoundData();
			});

			$("#releaseBindingBtn").click(function (event) {
				releaseBinding();
			});
			$("#setsettingBtn").click(function (event) {
				setseting();
			});
			$("#getsettingBtn").click(function (event) {
				getsetting();
			});
			$("#setSettingDataBtn").click(function (event) {
				setSettingData();
			});
			$("#getSettingDataBtn").click(function (event) {
				getSettingData();
			});
			$("#RemoveSettingDataBtn").click(function (event) {
				removeSettingData();
			});
			$("#saveAutoShowTaskpaneBtn").click(function (event) {
				saveAutoShowTaskpaneWithDocument();
			});

			$("#addBDataChgEventBtn").click(function (event) {
				addBDataChgEventBtn();
			});
			$("#removeBDataChgEventBtn").click(function (event) {
				removeBDataChgEventBtn();
			});

			$("#addBSltChgEventBtn").click(function (event) {
				addBSltChgEventBtn();
			});
			$("#removeBSltChgEventBtn").click(function (event) {
				removeBSltChgEventBtn();
			});

			$("#addDocSltChgEventBtn").click(function (event) {
				addDocSltChgEventBtn();
			});
			$("#removeDocSltChgEventBtn").click(function (event) {
				removeDocSltChgEventBtn();
			});

			$("#goToNextSlideBtn").click(function (event) {
				goToNextSlide();
			});

			$('#runPerfBtn').click(function (event) {

				/* 				var myVar = setInterval(function () {
										if (k == 0) {
											runperf();
											k = 1;
										}
										if (i == 1) {
											j--;
											runperf();
										};
										if (j < 0)
											clearInterval(myVar);
									}, 400); */

				// window.setTimeout(function () { Office.context.document.setSelectedDataAsync("Crash");}, 15000);
			});
			//debugger

			colorDiv(Office.context.document.settings.get("backgroundColor"));

			showQueryString();

			// $('#ribbonDefinitionId').val(JSON.stringify(dynamic_ribbon_sample_data));
		});
	};
	// });

	//run perf
	function runperf() {
		i = 0;
		var start = new Date().getTime();
		Office.context.document.setSelectedDataAsync([["red"], ["green"], ["blue"], ["red"], ["green"], ["blue"], ["red"], ["green"], ["blue"], ["red"], ["green"], ["blue"]], function (asyncResult) {
			if (asyncResult.status === "failed") {
				writeToPage('Error: ' + asyncResult.error.message);
			}
			appResult(new Date().getTime() - start);
			appResult('<br/>');
			i = 1;
		});

		/*         Office.context.document.getSelectedDataAsync("matrix", function (asyncResult) {
		if (asyncResult.status === "failed") {
		writeToPage('Error: ' + asyncResult.error.message);
		}
		else {
		writeToPage('Selected data: ' + asyncResult.value);
		}
		appResult(new Date().getTime() - start);
		appResult('<br/>');
		i = 1;
		}); */

		//Office.select("bindings#myBinding").getDataAsync({ coercionType: "matrix" },
		//    function (asyncResult) {
		//        if (asyncResult.status === "failed") {
		//            //writeToPage('Error: ' + asyncResult.error.message);
		//        } else {
		//            //writeToPage('Selected data: ' + asyncResult.value);
		//        }
		//                appResult(new Date().getTime() - start);
		//                appResult('<br/>');
		//    });

		//Office.context.document.setSelectedDataAsync("", function (e) { });

	}

	function sleep(delay) {
		var start = new Date().getTime();
		while (new Date().getTime() < start + delay) ;
	}

	function appResult(txt) {
		$('#results').append(txt);
	}

	// Reads data from current document selection and displays a notification
	function getDataFromSelection() {
		Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
			function (result) {
				if (result.status === Office.AsyncResultStatus.Succeeded) {
					app.showNotification('The selected text is:', '"' + result.value + '"');
				} else {
					app.showNotification('Error:', result.error.message);
				}
			});
	}

	function writeData() {
		Office.context.document.setSelectedDataAsync([["red"], ["green"], ["blue"], ["red"], ["green"], ["blue"], ["red"], ["green"], ["blue"], ["red"], ["green"], ["blue"]], function (asyncResult) {
			if (asyncResult.status === "failed") {
				writeToPage('Error: ' + asyncResult.error.message);
			}
		});

	}

	function enableCommandButtons() {
		updateCommandButtonState(true);
	}

	function disableCommandButtons() {
		updateCommandButtonState(false);
	}

	async function updateCommandButtonState(enabled) {
		var commandButtonIds = document.getElementById('commandButtonID').value.split(';');
		var commandControls = [];

		for (var i = 0; i < commandButtonIds.length; i++) {
			var commandButtonID = commandButtonIds[i];
			if (commandButtonID != "") {
				commandControls.push({id: commandButtonID, enabled: Boolean(enabled)});
			}
		}

		// var tab = {id: "OfficeAppTab1", controls: commandControls};
		// var data = {tabs: [tab]};

		// try {
		// 	Office.ribbon.requestUpdate(data)
		// 	.then(function(ret2) {})
		// 	.catch(function (err) { writeToPage('Error:' + JSON.stringify(err))});
		// }
		// catch (error) {
		// 	writeToPage('Error:' + error.toString());
		// }
		// if (Office.context.requirements.isSetSupported('RibbonApi', '1.1')) {
		const parentGroup = {id: "Group1Id12", controls: commandControls};
		const parentTab = {id: "OfficeAppTab1", groups: [parentGroup]};
		const ribbonUpdater = {tabs: [parentTab]};
		// @ts-ignore
		await Office.ribbon.requestUpdate(ribbonUpdater);
		// }

	}

	function updateContextualTabsVisibility(visible) {
		var tabIds = document.getElementById('ribbonTabId').value.split(';');
		var btn = {};
		var commandtabs = [];

		for (var i = 0; i < tabIds.length; i++) {
			var tabId = tabIds[i];
			if (tabId != "") {
				commandtabs.push({id: tabId, visible: Boolean(visible), controls: [btn]});
			}
		}

		var data = {tabs: commandtabs};
		try {
			Office.ribbon.requestUpdate(data)
				.then(function (ret2) {
				})
				.catch(function (err) {
					writeToPage('Error:' + JSON.stringify(err))
				});
		} catch (error) {
			OfficeRuntime.ui.getRibbon().then(function (ret) {
				var dr = ret;
				dr.requestUpdate(data)
					.then(function (ret2) {
					})
					.catch(function (err) {
						writeToPage('Error:' + JSON.stringify(err))
					});
			});
		}
	}

	var dynamic_ribbon_sample_data =
		{
			"actions": [
				{
					"id": "executeWriteData",
					"type": "ExecuteFunction",
					"functionFile": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/HomeAppCmds.html",
					"functionName": "writeData"
				},
				{
					"id": "executeWriteformula",
					"type": "ExecuteFunction",
					"functionFile": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/HomeAppCmds.html",
					"functionName": "writeformula"
				},
				{
					"id": "showTaskpanewiki",
					"type": "ShowTaskpane",
					"sourceLocation": "https://wikipedia.firstpartyapps.oaspapps.com/wikipedia/wikipedia_dev.html",
					"taskpaneId": "Taskpane1",
					"title": "TakspaneTitle",
					"supportPinning": false
				},
				{
					"id": "showTaskpaneResCFSample",
					"type": "ShowTaskpane",
					"sourceLocation": "https://officedev.github.io/custom-functions/addins/cfsample2/sharedapp.html",
					"taskpaneId": "Taskpane1",
					"title": "TakspaneTitle",
					"supportPinning": false
				},
				{
					"id": "showTaskpaneResAppCmds",
					"type": "ShowTaskpane",
					"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/HomeAppCmds.html",
					"taskpaneId": "Taskpane1",
					"title": "TakspaneTitle",
					"supportPinning": false
				}
			],
			"tabs": [
				{
					"id": "CtxTab1",
					"label": "CtxTab1",
					"visible": false,
					"groups": [
						{
							"id": "CustomGroup111",
							"label": "Group11Title",
							"icon": [
								{
									"size": 16,
									"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
								},
								{
									"size": 32,
									"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
								},
								{
									"size": 80,
									"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
								}
							],
							"controls": [
								{
									"type": "Button",
									"id": "CtxBt111",
									"enabled": true,
									"icon": [
										{
											"size": 16,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 32,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 80,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										}
									],
									"label": "STP_CtxBt111",
									"toolTip": "Btn111ToolTip",
									"superTip": {
										"title": "Btn111SupertTipeTitle",
										"description": "Btn111SuperTipDesc"
									},
									"actionId": "showTaskpaneResAppCmds"
								},
								{
									"type": "Button",
									"id": "CtxBt112",
									"enabled": true,
									"icon": [
										{
											"size": 16,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 32,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 80,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										}
									],
									"label": "ExeFunc_CtxBt112",
									"toolTip": "Btn112ToolTip",
									"superTip": {
										"title": "Btn112SupertTipeTitle",
										"description": "Btn112SuperTipDesc"
									},
									"actionId": "executeWriteData"
								},
								{
									"type": "Button",
									"id": "CtxBt113",
									"enabled": false,
									"icon": [
										{
											"size": 16,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 32,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 80,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										}
									],
									"label": "STP_CtxBt113",
									"toolTip": "STP_CtxBt113",
									"superTip": {
										"title": "Btn111SupertTipeTitle",
										"description": "Btn111SuperTipDesc"
									},
									"actionId": "showTaskpanewiki"
								},
								{
									"type": "Button",
									"id": "CtxBt114",
									"enabled": true,
									"icon": [
										{
											"size": 16,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 32,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 80,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										}
									],
									"label": "ExeFunc_CtxBt114",
									"toolTip": "Btn114ToolTip",
									"superTip": {
										"title": "Btn114SupertTipeTitle",
										"description": "Btn114SuperTipDesc"
									},
									"actionId": "executeWriteData"
								},
								{
									"type": "Menu",
									"id": "CustomRibbonTab1Menu1",
									"label": "Menu111Label",
									"toolTip": "Btn112ToolTip",
									"superTip": {
										"title": "Btn112SupertTipeTitle",
										"description": "Btn112SuperTipDesc"
									},
									"icon": [
										{
											"size": 16,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 32,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 80,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										}
									],
									"items": [
										{
											"type": "MenuItem",
											"id": "CtxMi111",
											"enabled": true,
											"icon": [
												{
													"size": 16,
													"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
												},
												{
													"size": 32,
													"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
												},
												{
													"size": 80,
													"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
												}
											],
											"label": "STP_CtxMi111",
											"toolTip": "Btn111ToolTip",
											"superTip": {
												"title": "Btn111SupertTipeTitle",
												"description": "Btn111SuperTipDesc"
											},
											"actionId": "showTaskpanewiki"
										},
										{
											"type": "MenuItem",
											"id": "CtxMi112",
											"enabled": false,
											"icon": [
												{
													"size": 16,
													"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
												},
												{
													"size": 32,
													"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
												},
												{
													"size": 80,
													"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
												}
											],
											"label": "ExeFunc_CtxMi112",
											"toolTip": "Btn111ToolTip",
											"superTip": {
												"title": "Btn111SupertTipeTitle",
												"description": "Btn111SuperTipDesc"
											},
											"actionId": "executeWriteformula"
										}
									]
								}
							]
						}
					]
				},
				{
					"id": "CtxTab2",
					"label": "CtxTab2",
					"visible": true,
					"groups": [
						{
							"id": "CustomGroup211",
							"label": "Group211Title",
							"icon": [
								{
									"size": 16,
									"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
								},
								{
									"size": 32,
									"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
								},
								{
									"size": 80,
									"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
								}
							],
							"controls": [
								{
									"type": "Button",
									"id": "CtxBt211",
									"enabled": false,
									"icon": [
										{
											"size": 16,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 32,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 80,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										}
									],
									"label": "STP_CtxBt211",
									"toolTip": "Btn211ToolTip",
									"superTip": {
										"title": "Btn211SuperTipTitle",
										"description": "Btn211SuperTipDesc"
									},
									"actionId": "showTaskpaneResCFSample"
								},
								{
									"type": "Button",
									"id": "CtxBt212",
									"enabled": true,
									"icon": [
										{
											"size": 16,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 32,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 80,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										}
									],
									"label": "ExeFunc_CtxBt212",
									"toolTip": "Btn212ToolTip",
									"superTip": {
										"title": "Btn212SuperTipTitle",
										"description": "Btn212SuperTipDesc"
									},
									"actionId": "executeWriteformula"
								}
							]
						}
					]
				}
			]
		};

	function generateRibbonButton() {
		var ribbonTabDefinition = document.getElementById('ribbonDefinitionId').value;
		if (!ribbonTabDefinition || ribbonTabDefinition === "") {
			ribbonTabDefinition = JSON.stringify(dynamic_ribbon_sample_data);
		}

		Office.ribbon.requestCreateControls(JSON.parse(ribbonTabDefinition))
			.then(function (ret2) {
			})
			.catch(function (err) {
				writeToPage('Error:' + JSON.stringify(err))
			});
	}

	function generateRandomRibbonDefinition() {

		var getRandomInt = function (max) {
			return Math.floor(Math.random() * Math.floor(max));
		};

		//Generate a single tab
		var generateTab = function (id) {
			var newTab = '{"id":"AddinTab${id}","title":"Addin Tab ${id}","isOfficeTab":false,"visible":true,"groups":[{"id":"AddinControl${id}${id}","title":"TestOffice","controls":[{"id":"AddinControl${id}Button1","title":"Tab${id}Button1","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_1.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button2","title":"Tab${id}Button2","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_2.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button3","title":"Tab${id}Button3","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_3.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button4","title":"Tab${id}Button4","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_4.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button5","title":"Tab${id}Button5","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_5.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button6","title":"Tab${id}Button6","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_6.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button7","title":"Tab${id}Button7","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_7.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button8","title":"Tab${id}Button8","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_8.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button9","title":"Tab${id}Button9","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_9.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button10","title":"Tab${id}Button10","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_10.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button11","title":"Tab${id}Button11","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_11.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button12","title":"Tab${id}Button12","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_12.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button13","title":"Tab${id}Button13","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_13.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button14","title":"Tab${id}Button14","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_14.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button15","title":"Tab${id}Button15","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_15.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button16","title":"Tab${id}Button16","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_16.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button17","title":"Tab${id}Button17","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_17.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button18","title":"Tab${id}Button18","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_18.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button19","title":"Tab${id}Button19","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_19.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}},{"id":"AddinControl${id}Button20","title":"Tab${id}Button20","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icons/icon32_20.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}}]}]}';
			return newTab;
		};

		var generateTabs = function (count) {
			var tabs = "";
			for (var i = 0; i < count; ++i) {
				tabs += ',';
				tabs += generateTab(i);
			}
			return tabs;
		};

		var generateTabDefinition = function (count) {
			var generatedTabs = generateTabs(count);
			var tabDefinition = '[{"id":"TabHome","title":null,"isOfficeTab":true,"visible":true,"groups":[{"id":"AddinControlHome","title":"Addin Tab Home","controls":[{"id":"AddinControlHome1","title":"STP_Shared_TabHomeButton1","iconLocation":"https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png","enabled":true,"controlType":0,"attributes":{"Command":1402209562}}]}]}${generatedTabs}]'
			return tabDefinition;
		};

		var generatedRibbonDefinition = generateTabDefinition(getRandomInt(20) + 1);
		document.getElementById('ribbonDefinitionId').value = generatedRibbonDefinition;

		const requestContext = new OfficeCore.RequestContext();
		requestContext._customData = "WacPartition";
		requestContext.ribbon.executeRequestUpdate(generatedRibbonDefinition);
		requestContext.sync();
		return;
	}


	function addSelectionChangedEvent() {

		var sheet = context.workbook.worksheets.getItem("Sheet1");
		sheet.onSelectionChanged.add(function (event) {
			return Excel.run(function (context) {
				var range = context.workbook.getSelectedRange();
				range.load(['address', 'values']);
				var firstSelectedCellValue = range.values[0][0];

				var ribbonTabId = document.getElementById('ribbonTabId').value;
				if (firstSelectedCellValue === "true") {
					OfficeRuntime.ui.getRibbon().then(function (ret) {
						var dr = ret;
						var btn = {};
						var tab = {id: ribbonTabId, visible: true, controls: [btn]};
						var data = {tabs: [tab]};
						dr.requestUpdate(data);
					});
				} else if (firstSelectedCellValue === "false") {
					OfficeRuntime.ui.getRibbon().then(function (ret) {
						var dr = ret;
						var btn = {};
						var tab = {id: ribbonTabId, visible: false, controls: [btn]};
						var data = {tabs: [tab]};
						dr.requestUpdate(data);
					});
				}
			});
		});
	}

	function removeSelectionChangedEvent() {
		var ribbonTabId = document.getElementById('ribbonTabId').value;
		OfficeRuntime.ui.getRibbon().then(function (ret) {
			var dr = ret;
			var btn = {};
			var tab = {id: ribbonTabId, visible: false, controls: [btn]};
			var data = {tabs: [tab]};
			dr.requestUpdate(data);
		});
	}


	function writeDataBindingBtn() {
		var bindingName = document.getElementById('bindingName').value;
		Office.context.document.setSelectedDataAsync([["red"], ["green"], ["blue"], ["red"], ["green"], ["blue"], ["red"], ["green"], ["blue"], ["red"], ["green"], ["blue"]], function (asyncResult) {
			if (asyncResult.status === "failed") {
				writeToPage('Error: ' + asyncResult.error.message);
			} else {
				Office.context.document.bindings.addFromSelectionAsync(
					Office.BindingType.Matrix,
					{id: bindingName},
					function (asyncResult) {
						if (asyncResult.status === "failed") {
							writeToPage('Error: ' + asyncResult.error.message);
						} else {
							writeToPage('Added binding with type: ' + asyncResult.value.type + ' and id: ' +
								asyncResult.value.id);
						}
					});

			}
		});
	}


	function readData() {
		Office.context.document.getSelectedDataAsync("matrix", function (asyncResult) {
			if (asyncResult.status === "failed") {
				writeToPage('Error: ' + asyncResult.error.message);
			} else {
				writeToPage('Selected data: ' + asyncResult.value);
			}
		});
	}

	function bindData() {
		//addFromSelectionAsync
		var bindingName = document.getElementById('bindingName').value;

		Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Matrix, {
				id: bindingName
			},

			//addFromPromptAsync
			/* 		Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Matrix, {
			id : 'myBinding',
			promptText : 'Select text to bind to.'
			}, */

			//addFromNamedItemAsync
			/* 		Office.context.document.bindings.addFromNamedItemAsync("matrix1",
			Office.BindingType.Matrix, {
			id : 'myBinding'
			},
			 */
			function (asyncResult) {
				if (asyncResult.status === "failed") {
					writeToPage('Error: ' + asyncResult.error.message);
				} else {
					writeToPage('Added binding with type: ' + asyncResult.value.type + ' and id: ' +
						asyncResult.value.id);
				}
			});
	}

	function getallbindings() {
		Office.context.document.bindings.getAllAsync(function (asyncResult) {
			var bindingString = '';
			for (var i in asyncResult.value) {
				bindingString += asyncResult.value[i].id + '\n';
			}
			writeToPage('Existing bindings: ' + bindingString);
		});

	}

	function readBoundData() {
		var bindingName = "bindings#" + document.getElementById('bindingName').value;
		Office.select(bindingName).getDataAsync({
				coercionType: Office.BindingType.Matrix
			},
			function (asyncResult) {
				if (asyncResult.status === "failed") {
					writeToPage('Error: ' + asyncResult.error.message);
				} else {
					writeToPage('Selected data: ' + asyncResult.value);
				}
			});
	}

	function setBoundData() {
		var bindingName = "bindings#" + document.getElementById('bindingName').value;
		Office.select(bindingName).setDataAsync([['Berlin'], ['Munich'], ['Duisburg']], {
				coercionType: "matrix"
			},
			function (asyncResult) {
				if (asyncResult.status === "failed") {
					writeToPage('Error: ' + asyncResult.error.message);
				} else {
					writeToPage('Bound data: ' + asyncResult.value);
				}
			});
	}

	function setseting() {
		// Set a setting in the document
		Office.context.document.settings.set("backgroundColor", "7");
		writeToPage("backgroundColor set as 7.");
	}

	function getsetting() {
		//Get a setting previously set in the document
		var settingsValue = Office.context.document.settings.get("backgroundColor");
		writeToPage("backgroundColor value is: " + settingsValue);
	}

	function setSettingData() {
		//addSelectionChangedEventHandler();
		var settingname = document.getElementById("settingName").value;
		var settingvalue = document.getElementById("settingvalue").value;
		Office.context.document.settings.set(settingname, settingvalue);
		writeToPage('Set setting: ' + settingname + '->' + settingvalue);
		//Save a setting in the document to make it available in future sessions
		Office.context.document.settings.saveAsync(function (asyncResult) {
			if (asyncResult.status == "failed") {
				writeToPage("Action failed with error: " + asyncResult.error.message);
			} else {
				writeToPage("Settings saved with status: " + asyncResult.status);
			}
		});
	}

	function saveAutoShowTaskpaneWithDocument() {
		var settingname = 'Office.AutoShowTaskpaneWithDocument';
		var settingvalue = document.getElementById("showtaskpane").checked;
		Office.context.document.settings.set(settingname, settingvalue);
		writeToPage('Set setting: ' + settingname + '->' + settingvalue);
		//Save a setting in the document to make it available in future sessions
		Office.context.document.settings.saveAsync(function (asyncResult) {
			if (asyncResult.status == "failed") {
				writeToPage("Action failed with error: " + asyncResult.error.message);
			} else {
				writeToPage("Settings saved with status: " + asyncResult.status);
			}
		});
	}

	function addSelectionChangedEventHandler() {
		Office.context.document.settings.addHandlerAsync(Office.EventType.SettingsChanged, MySettingHandler, function (asyncResult) {
			if (asyncResult.status === "failed") {
				writeToPage('Error: ' + asyncResult.error.message);
			} else {
				writeToPage('Added event handler');
			}
		});
	}

	function MySettingHandler(eventArgs) {
		var getsettingName = document.getElementById('getsettingName').value;
		Office.context.document.settings.refreshAsync(function (asyncResult) {
			if (asyncResult.status === "failed") {
				writeToPage('Error: ' + asyncResult.error.message);
			} else {
				//var settingsValue = Office.context.document.settings.get(getsettingName);
				var settingsValue = asyncResult.value.get(getsettingName);
				writeToPage(getsettingName + " value is: " + settingsValue);
				colorDiv(settingsValue);
			}
		});
	}

	function colorDiv(coloroption) {
		//    div.style.backgroundColor = 'green';
		var color;
		switch (coloroption) {
			case "1":
				color = 'grey';
				break;
			case "2":
				color = 'lightblue';
				break;
			case "3":
				color = 'blue';
				break;
			case "4":
				color = 'DarkOrange';
				break;
			case "5":
				color = 'green';
				break;
			case "6":
				color = 'FireBrick';
				break;
			case "7":
				color = 'DarkOliveGreen ';
				break;
		}
		var div = document.getElementById('content-header');
		div.style.backgroundColor = color;

	}

	function showQueryString() {
		writeToPage(document.URL);
	}

	function getSettingData() {
		/*         var getsettingName = document.getElementById('getsettingName').value;
		Office.context.document.settings.refreshAsync(function (asyncResult) {
		if (asyncResult.status === "failed") {
		writeToPage('Error: ' + asyncResult.error.message);
		} else {
		//var settingsValue = Office.context.document.settings.get(getsettingName);
		var settingsValue = asyncResult.value;
		writeToPage(getsettingName + " value is: " + settingsValue[getsettingName]);
		}
		}); */

		addSelectionChangedEventHandler();

	}

	function removeSettingData() {
		Office.context.document.settings.removeHandlerAsync(Office.EventType.SettingsChanged, function (asyncResult) {
			if (asyncResult.status === "failed") {
				writeToPage('Error: ' + asyncResult.error.message);
			} else {
				writeToPage('Removed SettingsChanged event handler.');
			}
		});
	}

	function addBDataChgEventBtn() {
		var bindingName = "bindings#" + document.getElementById('bindingName').value;
		Office.select(bindingName).addHandlerAsync(Office.EventType.BindingDataChanged, myHandler, function (asyncResult) {
			if (asyncResult.status === "failed") {
				writeToPage('Error: ' + asyncResult.error.message);
			} else {
				writeToPage('Added BindingDataChanged handler');
			}
		});
	}

	function removeBDataChgEventBtn() {
		var bindingName = "bindings#" + document.getElementById('bindingName').value;
		Office.select(bindingName).removeHandlerAsync(Office.EventType.BindingDataChanged, function (asyncResult) {
			if (asyncResult.status === "failed") {
				writeToPage('Error: ' + asyncResult.error.message);
			} else {
				writeToPage('Removed bindingDataChanged event handler');
			}
		});

	}

	function addBSltChgEventBtn() {
		var bindingName = "bindings#" + document.getElementById('bindingName').value;
		Office.select(bindingName).addHandlerAsync(Office.EventType.BindingSelectionChanged, myHandler2, function (asyncResult) {
			if (asyncResult.status === "failed") {
				writeToPage('Error: ' + asyncResult.error.message);
			} else {
				writeToPage('Added BindingSelectionChanged handler');
			}
		});

	}

	function removeBSltChgEventBtn() {
		var bindingName = "bindings#" + document.getElementById('bindingName').value;
		Office.select(bindingName).removeHandlerAsync(Office.EventType.BindingSelectionChanged, function (asyncResult) {
			if (asyncResult.status === "failed") {
				writeToPage('Error: ' + asyncResult.error.message);
			} else {
				writeToPage('Removed BindingSelectionChanged event handler');
			}
		});

	}

	function addDocSltChgEventBtn() {
		Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, documentSelectionHandler, function (asyncResult) {
			if (asyncResult.status === "failed") {
				writeToPage('Error: ' + asyncResult.error.message);
			} else {
				writeToPage('Added DocumentSelectionChanged handler');
			}
		});

	}

	function removeDocSltChgEventBtn() {
		Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, function (asyncResult) {
			if (asyncResult.status === "failed") {
				writeToPage('Error: ' + asyncResult.error.message);
			} else {
				writeToPage('Removed DocumentSelectionChanged handler');
			}
		});

	}

	function goToNextSlide() {

		var data = "100";
		$.get("http://jackychenaddin.azurewebsites.net/AddInService.svc/GetData", function (response) {
			data = response;
		}).error(function () {
			writeToPage('go To Next Slide Error');
		});

		// Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index,
		// 	function (asyncResult) {
		// 		if (asyncResult.status == "failed") {
		// 			writeToPage("Error", asyncResult.error.message);
		// 		} else {
		// 			writeToPage('go To Next Slide');
		// 		}
		// 	});

	}


	function documentSelectionHandler(eventArgs) {
		Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
			if (asyncResult.status != "succeeded") {
				return;
			}

			var rows = asyncResult.value.split(/\r?\n/);
			var columnValues = rows[0].split("	");
			var ribbonTabId = columnValues[0];
			var isVisible = (columnValues[1] == "true");

			OfficeRuntime.ui.getRibbon().then(function (ret) {
				var dr = ret;
				var btn = {};
				var tab = {id: ribbonTabId, visible: isVisible, controls: [btn]};
				var data = {tabs: [tab]};
				dr.requestUpdate(data);
			});

			writeToPage("DocumentSelectionChanged: " + asyncResult.value);
		});
	}


	function myHandler2(eventArgs) {
		var commandTaab1ButtonIDs = ["Tab1Button1", "Tab1Button2", "Tab1Button3", "Tab1Button4", "Tab1Button5", "Tab1Button6", "Tab1Menu1Item1", "Tab1Menu1Item2"];
		var commandTaab2ButtonIDs = ["Tab2Button1"];
		var customTabIds = ["OfficeAppTab1", "OfficeAppTab2"];

		var tab1CommandControls = [];
		var tab2CommandControls = [];
		var commandtabs = [];

		for (var i = 0; i < commandTaab1ButtonIDs.length; i++) {
			var commandButtonID = commandTaab1ButtonIDs[i];
			if (commandButtonID != "") {
				tab1CommandControls.push({id: commandButtonID, enabled: (Math.random() >= 0.5)});
			}
		}

		for (var i = 0; i < commandTaab2ButtonIDs.length; i++) {
			var commandButtonID = commandTaab2ButtonIDs[i];
			if (commandButtonID != "") {
				tab2CommandControls.push({id: commandButtonID, enabled: (Math.random() >= 0.5)});
			}
		}

		commandtabs.push({id: customTabIds[0], visible: true/*(Math.random() >= 0.5)*/, controls: tab1CommandControls});
		commandtabs.push({id: customTabIds[1], visible: (Math.random() >= 0.5), controls: tab2CommandControls});

		var data = {tabs: commandtabs};
		Office.ribbon.requestUpdate(data)
			.then(function (ret2) {
			})
			.catch(function (err) {
				writeToPage('Error:' + JSON.stringify(err))
			});
	}


	function myHandler(eventArgs) {
		eventArgs.binding.getDataAsync({
			coerciontype: "matrix"
		}, function (asyncResult) {

			if (asyncResult.status === "failed") {
				writeToPage('Error: ' + asyncResult.error.message);
			} else {
				writeToPage('Bound data: ' + asyncResult.value);
			}
		});
	}

	//Trigger on selection change, get partial data from the matrix
	function onBindingSelectionChanged(eventArgs) {
		eventArgs.binding.getDataAsync({
				CoercionType: Office.CoercionType.Matrix,
				startRow: eventArgs.startRow,
				startColumn: 0,
				rowCount: 1,
				columnCount: 1
			},
			function (asyncResult) {
				if (asyncResult.status == "failed") {
					writeToPage("Action failed with error: " + asyncResult.error.message);
				} else {
					writeToPage(asyncResult.value[0].toString());
				}
			});
	}

	function releaseBinding() {
		var bindingName = document.getElementById('bindingName').value;
		Office.context.document.bindings.releaseByIdAsync(bindingName, function (asyncResult) {
			if (asyncResult.status === "failed") {
				writeToPage('Error: ' + asyncResult.error.message);
			} else {
				writeToPage(bindingName + ' released.');
			}
		});
	}

	function writeToPage(text) {
		// var elmResults =  document.getElementById('results');
		// text = elmResults.innerHTML + text;
		// elmResults.innerHTML = text;
		// elmResults.appendChild(document.createElement("br"));
		// console.log(text)
	}

})();

function writeData(event) {
	// Office.context.document.setSelectedDataAsync([["red"], ["green"], ["blue"], ["red"], ["green"], ["blue"], ["red"], ["green"], ["blue"], ["red"], ["green"], ["blue"]], function (asyncResult) {
	// 	if (asyncResult.status === "failed") {
	// 		writeToPage('Error: ' + asyncResult.error.message);
	// 	}
	// });

	Office.context.document.setSelectedDataAsync("Start sleeping...", function (asyncResult) {
		if (asyncResult.status !== "failed") {
			var start = new Date().getTime();
			while (new Date().getTime() < start + 10000) ;

			Office.context.document.setSelectedDataAsync("Changing control state by Ui-less.");
		}
	});


	// var commandTaab1ButtonIDs = ["Tab1Button2", "Tab1Button3", "Tab1Button4", "Tab1Button5", "Tab1Button6", "Tab1Menu1Item1", "Tab1Menu1Item2"];
	// var commandTaab2ButtonIDs = ["Tab2Button1"];
	// var customTabIds = ["OfficeAppTab1", "OfficeAppTab2"];

	// var tab1CommandControls = [];
	// var tab2CommandControls = [];
	// var commandtabs = [];

	// for (var i = 0; i < commandTaab1ButtonIDs.length; i++) {
	// 	var commandButtonID = commandTaab1ButtonIDs[i];
	// 	if (commandButtonID != "") {
	// 		tab1CommandControls.push({id: commandButtonID, enabled: (Math.random() >= 0.5)});
	// 	}
	// }

	// for (var i = 0; i < commandTaab2ButtonIDs.length; i++) {
	// 	var commandButtonID = commandTaab2ButtonIDs[i];
	// 	if (commandButtonID != "") {
	// 		tab2CommandControls.push({id: commandButtonID, enabled: (Math.random() >= 0.5)});
	// 	}
	// }

	// commandtabs.push({id: customTabIds[0], visible: true/*(Math.random() >= 0.5)*/, controls: tab1CommandControls});
	// commandtabs.push({id: customTabIds[1], visible: (Math.random() >= 0.5), controls: tab2CommandControls});

	// OfficeRuntime.ui.getRibbon().then(function (ret) {
	// 	var dr = ret;
	// 	var data = {tabs: commandtabs};
	// 	dr.requestUpdate(data)
	// 	.then(function(ret2) { Office.context.document.setSelectedDataAsync('Done:' + JSON.stringify(ret2))})
	// 	.catch(function (err) { Office.context.document.setSelectedDataAsync('Error:' + JSON.stringify(err))});
	// });

	/* 		var bindingName = "MyUILessBinding";
			Office.context.document.setSelectedDataAsync([["red"], ["green"], ["blue"], ["red"], ["green"], ["blue"], ["red"], ["green"], ["blue"], ["red"], ["green"], ["blue"]], function (asyncResult) {
				if (asyncResult.status === "failed") {
					// writeToPage('Error: ' + asyncResult.error.message);
				} else {
						Office.context.document.bindings.addFromSelectionAsync(
							Office.BindingType.Matrix,
							{ id : bindingName},
							function (asyncResult) {
								if (asyncResult.status === "failed") {
								} else {
								}
							});

				}
			}); */

	//Office.context.document.setSelectedDataAsync("insert test from UI-less function.");
	// Office.context.document.setSelectedDataAsync("");
	event.completed();

}

function writeformula() {
	Office.context.document.setSelectedDataAsync("=CFSample2.GetValue()");
	//setTimeout("", 10000);
	event.completed();

}


function ChangetoRed(event) {
	Office.context.document.settings.set("backgroundColor", "6");
	//Save a setting in the document to make it available in future sessions
	Office.context.document.settings.saveAsync(function (asyncResult) {
		if (asyncResult.status == "failed") {
			// writeToPage("Action failed with error: " + asyncResult.error.message);
		} else {
			Office.context.document.setSelectedDataAsync("backgroundColor changed.");
			// writeToPage("Settings saved with status: " + asyncResult.status);
		}
	});

	event.completed();
}

async function BoldText(event) {
	console.log("changing text")
	await Word.run(async (context) => {
		const range = context.document.getSelection();
		range.font.color = "red";
		range.load("text");

		await context.sync();

		console.log(`The selected text was "${range.text}".`);
		event.completed();
	});
}

function ShowAlertTest(event) {
	console.log("QQQQ")
	// var DialogElements = document.querySelectorAll(".ms-Dialog");
	// var DialogComponents = [];
	// for (var i = 0; i < DialogElements.length; i++) {
	// 	(function() {
	// 		DialogComponents[i] = new fabric['Dialog'](DialogElements[i]);
	// 	}());
	// }
	// //debugger;
	var currentUrl = document.URL;
	var dialogPage = currentUrl.substring(0, currentUrl.lastIndexOf("/")) + "/Dialog.html";
	Office.context.ui.displayDialogAsync(dialogPage, {height: 30, width: 20, displayInIframe: true});
	event.completed();
}

function ChangeBindingData(event) {
	var bindingName = "bindings#" + document.getElementById('bindingName').value;
	Office.select(bindingName).setDataAsync([['Berlin'], ['Munich'], ['Duisburg']], {
			coercionType: "matrix"
		},
		function (asyncResult) {
			if (asyncResult.status === "failed") {
				writeToPage('Error: ' + asyncResult.error.message);
			} else {
				writeToPage('Bound data: ' + asyncResult.value);
			}
		});

	event.completed();
}

function createContextualTab(event) {
	var dynamic_ribbon_sample_data =
		{
			"actions": [
				{
					"id": "executeWriteData",
					"type": "ExecuteFunction",
					"functionFile": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/HomeAppCmds.html",
					"functionName": "writeData"
				},
				{
					"id": "executeWriteformula",
					"type": "ExecuteFunction",
					"functionFile": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/HomeAppCmds.html",
					"functionName": "writeformula"
				},
				{
					"id": "showTaskpanewiki",
					"type": "ShowTaskpane",
					"sourceLocation": "https://wikipedia.firstpartyapps.oaspapps.com/wikipedia/wikipedia_dev.html",
					"taskpaneId": "Taskpane1",
					"title": "TakspaneTitle",
					"supportPinning": false
				},
				{
					"id": "showTaskpaneResCFSample",
					"type": "ShowTaskpane",
					"sourceLocation": "https://officedev.github.io/custom-functions/addins/cfsample2/sharedapp.html",
					"taskpaneId": "Taskpane1",
					"title": "TakspaneTitle",
					"supportPinning": false
				},
				{
					"id": "showTaskpaneResAppCmds",
					"type": "ShowTaskpane",
					"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/HomeAppCmds.html",
					"taskpaneId": "Taskpane1",
					"title": "TakspaneTitle",
					"supportPinning": false
				}
			],
			"tabs": [
				{
					"id": "CtxTab1",
					"label": "CtxTab1",
					"visible": true,
					"groups": [
						{
							"id": "CustomGroup111",
							"label": "Group11Title",
							"icon": [
								{
									"size": 16,
									"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
								},
								{
									"size": 32,
									"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
								},
								{
									"size": 80,
									"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
								}
							],
							"controls": [
								{
									"type": "Button",
									"id": "CtxBt111",
									"enabled": true,
									"icon": [
										{
											"size": 16,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 32,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 80,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										}
									],
									"label": "STP_CtxBt111",
									"toolTip": "Btn111ToolTip",
									"superTip": {
										"title": "Btn111SupertTipeTitle",
										"description": "Btn111SuperTipDesc"
									},
									"actionId": "showTaskpaneResAppCmds"
								},
								{
									"type": "Button",
									"id": "CtxBt112",
									"enabled": true,
									"icon": [
										{
											"size": 16,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 32,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 80,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										}
									],
									"label": "ExeFunc_CtxBt112",
									"toolTip": "Btn112ToolTip",
									"superTip": {
										"title": "Btn112SupertTipeTitle",
										"description": "Btn112SuperTipDesc"
									},
									"actionId": "executeWriteData"
								},
								{
									"type": "Button",
									"id": "CtxBt113",
									"enabled": false,
									"icon": [
										{
											"size": 16,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 32,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 80,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										}
									],
									"label": "STP_CtxBt113",
									"toolTip": "STP_CtxBt113",
									"superTip": {
										"title": "Btn111SupertTipeTitle",
										"description": "Btn111SuperTipDesc"
									},
									"actionId": "showTaskpanewiki"
								},
								{
									"type": "Menu",
									"id": "CustomRibbonTab1Menu1",
									"label": "Menu111Label",
									"toolTip": "Btn112ToolTip",
									"superTip": {
										"title": "Btn112SupertTipeTitle",
										"description": "Btn112SuperTipDesc"
									},
									"icon": [
										{
											"size": 16,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 32,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 80,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										}
									],
									"items": [
										{
											"type": "MenuItem",
											"id": "CtxMi111",
											"enabled": true,
											"icon": [
												{
													"size": 16,
													"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
												},
												{
													"size": 32,
													"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
												},
												{
													"size": 80,
													"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
												}
											],
											"label": "STP_CtxMi111",
											"toolTip": "Btn111ToolTip",
											"superTip": {
												"title": "Btn111SupertTipeTitle",
												"description": "Btn111SuperTipDesc"
											},
											"actionId": "showTaskpanewiki"
										},
										{
											"type": "MenuItem",
											"id": "CtxMi112",
											"enabled": false,
											"icon": [
												{
													"size": 16,
													"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
												},
												{
													"size": 32,
													"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
												},
												{
													"size": 80,
													"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
												}
											],
											"label": "ExeFunc_CtxMi112",
											"toolTip": "Btn111ToolTip",
											"superTip": {
												"title": "Btn111SupertTipeTitle",
												"description": "Btn111SuperTipDesc"
											},
											"actionId": "executeWriteformula"
										}
									]
								}
							]
						}
					]
				},
				{
					"id": "CtxTab2",
					"label": "CtxTab2",
					"visible": true,
					"groups": [
						{
							"id": "CustomGroup211",
							"label": "Group211Title",
							"icon": [
								{
									"size": 16,
									"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
								},
								{
									"size": 32,
									"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
								},
								{
									"size": 80,
									"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
								}
							],
							"controls": [
								{
									"type": "Button",
									"id": "CtxBt211",
									"enabled": false,
									"icon": [
										{
											"size": 16,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 32,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 80,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										}
									],
									"label": "STP_CtxBt211",
									"toolTip": "Btn211ToolTip",
									"superTip": {
										"title": "Btn211SuperTipTitle",
										"description": "Btn211SuperTipDesc"
									},
									"actionId": "showTaskpaneResCFSample"
								},
								{
									"type": "Button",
									"id": "CtxBt212",
									"enabled": true,
									"icon": [
										{
											"size": 16,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 32,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										},
										{
											"size": 80,
											"sourceLocation": "https://jackychen.azurewebsites.net/agaves/testapps/app/home/icon32.png"
										}
									],
									"label": "ExeFunc_CtxBt212",
									"toolTip": "Btn212ToolTip",
									"superTip": {
										"title": "Btn212SuperTipTitle",
										"description": "Btn212SuperTipDesc"
									},
									"actionId": "executeWriteformula"
								}
							]
						}
					]
				}
			]
		};
//	debugger;
	Office.ribbon.requestCreateControls(dynamic_ribbon_sample_data)
		.then(function (ret2) {
			console.log("Contextal Tabs created.");
			Office.context.document.setSelectedDataAsync("Contextal Tabs created.")
			var data = {
				"tabs": [{"id": "CtxTab1", "visible": true, "controls": [{}]}, {
					"id": "CtxTab2",
					"visible": true,
					"controls": [{}]
				}]
			};
			Office.ribbon.requestUpdate(data)
				.then(function (ret3) {
					console.log("Contextal Tabs set visible.");
				})
				.catch(function (err2) {
					console.log('Error:' + JSON.stringify(err2))
				});
		})
		.catch(function (err) {
			console.log('Error:' + JSON.stringify(err));
		});

	event.completed();
}

function myHandler(eventArgs) {
	// Office.context.document.settings.set("backgroundColor", "6");
	//Save a setting in the document to make it available in future sessions
	Office.context.document.setSelectedDataAsync("Hello World!", function (asyncResult) {
		if (asyncResult.status == "failed") {
			// writeToPage("Action failed with error: " + asyncResult.error.message);
		} else {
			// writeToPage("Settings saved with status: " + asyncResult.status);
		}
	});
}

function bindSelectedData(event) {
	//addFromSelectionAsync
	Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Matrix, {
			id: 'myBinding'
		},

		//addFromPromptAsync
		/* 		Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Matrix, {
		id : 'myBinding',
		promptText : 'Select text to bind to.'
		}, */

		//addFromNamedItemAsync
		/* 		Office.context.document.bindings.addFromNamedItemAsync("matrix1",
		Office.BindingType.Matrix, {
		id : 'myBinding'
		},
		 */
		function (asyncResult) {
			if (asyncResult.status === "failed") {
				/* writeToPage('Error: ' + asyncResult.error.message); */
			} else {
				/* 				writeToPage('Added binding with type: ' + asyncResult.value.type + ' and id: ' +
				asyncResult.value.id); */
			}
		});

	event.completed();
}

function ChangetoGreen(event) {
	Office.context.document.settings.set("backgroundColor", "5");
	//Save a setting in the document to make it available in future sessions
	Office.context.document.settings.saveAsync(function (asyncResult) {
		if (asyncResult.status == "failed") {
			//  writeToPage("Action failed with error: " + asyncResult.error.message);
		} else {
			//  writeToPage("Settings saved with status: " + asyncResult.status);
		}
	});
	event.completed();

}
