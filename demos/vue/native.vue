<!-- xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com -->
<template>
	<div class="container">
		<image :src="logoUrl" class="logo"></image>
		<text class="welcome">SheetJS WeeX Demo {{version}}</text>
		<text class="instructions">Import Data</text>
		<text :style="{ color: '#841584' }" @click="importFile">Download spreadsheet</text>
		<text class="instructions">Export Data</text>
		<text :style="{ color: data.length ? '#841584' : '#CDCDCD', disabled: !data.length }" @click="exportFile">Upload XLSX</text>
		<text style="instructions">Current Data</text>
		<scroller class="scroller">
			<div class="row" v-for="(row, ridx) in data">
				<text>ROW {{ridx + 1}}</text>
				<text v-for="(cell, cidx) in row">CELL {{get_label(ridx, cidx)}}:{{cell}}</text>
			</div>
		</scroller>
	</div>
</template>

<style>
.container { height: 100%; flex: 1; justify-content: center; align-items: center; background-color: '#F5FCFF'; }
.logo { width: 256px; height: 256px; }
.welcome { font-size: 40; text-align: 'center'; margin: 10; }
.instructions { padding-top: 20px; color:#888; font-size: 24px;}
.scroller { height: 500px; border-width: 3px; width: 700px; }
.loading { justify-content: center; }
</style>

<script>
import XLSX from 'xlsx';
const modal = weex.requireModule('modal');
const stream = weex.requireModule('stream');
export default {
	data: {
		data: [[1,2,3],[4,5,6]],
		logoUrl: 'http://oss.sheetjs.com/assets/img/logo.png',
		version: XLSX.version,
		fileUrl: 'http://sheetjs.com/sheetjs.xlsx.b64',
		binUrl: 'https://hastebin.com/documents'
	},
	methods: {
		importFile: function (e) {
			modal.toast({ message: 'getting ' + this.fileUrl, duration: 1 });
			var self = this;
			stream.fetch({method:'GET', type:'text', url:this.fileUrl}, function(res){
				const wb = XLSX.read(res.data, {type:'base64'});
				const ws = wb.Sheets[wb.SheetNames[0]];
				self.data = XLSX.utils.sheet_to_json(ws, {header:1});
			});
		},
		exportFile: function (e) {
			var self = this;
			const ws = XLSX.utils.aoa_to_sheet(this.data);
			const wb = XLSX.utils.book_new();
			XLSX.utils.book_append_sheet(wb, ws, "SheetJS");
			const wbout = XLSX.write(wb, {type:"base64", bookType:"xlsx"});
			const body = wbout;
			stream.fetch({method:'POST', type:'json', url:this.binUrl, body:body}, function(res) {
				modal.toast({ message: 'KEY: ' + res.data.key, duration: 10 });
				self.version = res.data.key;
			});
		},
		get_label: function(r, c) { return XLSX.utils.encode_cell({r:r, c:c})}
	}
}
</script>
