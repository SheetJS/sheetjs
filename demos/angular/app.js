/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
var app = angular.module('app', ['ngAnimate', 'ngTouch', 'ui.grid', 'ui.grid.selection', 'ui.grid.exporter']);

/* Inject SheetJSExportService */
app.factory('SheetJSExportService', SheetJSExportService);
SheetJSExportService.inject = ['uiGridExporterService'];

app.controller('MainCtrl', ['$scope', '$http','SheetJSExportService', function ($scope, $http, SheetJSExportService) {
	$scope.gridOptions = {
		columnDefs: [
			{ field: 'name' },
			{ field: 'gender', visible: false},
			{ field: 'company' }
		],
		enableGridMenu: true,
		enableSelectAll: true,
		exporterMenuPdf: false,
		exporterMenuCsv: false,
		showHeader: true,
		onRegisterApi: function(gridApi){
			$scope.gridApi = gridApi;
		},
		/* SheetJS Service setup */
		filename: "SheetJSAngular",
		sheetname: "ng-SheetJS",
		gridMenuCustomItems: [
			{
				title: 'Export all data as XLSX',
				action: function ($event) { SheetJSExportService.exportXLSX($scope.gridApi); },
				order: 200
			},
			{
				title: 'Export all data as XLSB',
				action: function ($event) { SheetJSExportService.exportXLSB($scope.gridApi); },
				order: 201
			}
		]
	};

	$http.get('https://cdn.rawgit.com/angular-ui/ui-grid.info/gh-pages/data/100.json').success(function(data) { $scope.gridOptions.data = data; });

}]);
app.directive("importSheetJs", [SheetJSImportDirective]);
