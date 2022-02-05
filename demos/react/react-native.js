/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
import XLSX from 'xlsx';
import React, { Component } from 'react';
import {
	AppRegistry,
	StyleSheet,
	Text,
	View,
	Button,
	Alert,
	Image,
	ScrollView,
	TouchableWithoutFeedback
} from 'react-native';
import { Table, Row, Rows, TableWrapper } from 'react-native-table-component';

// react-native-file-access
var Base64 = function() {
  var map = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
  return {
    encode: function(input) {
      var o = "";
      var c1 = 0, c2 = 0, c3 = 0, e1 = 0, e2 = 0, e3 = 0, e4 = 0;
      for (var i = 0; i < input.length; ) {
        c1 = input.charCodeAt(i++);
        e1 = c1 >> 2;
        c2 = input.charCodeAt(i++);
        e2 = (c1 & 3) << 4 | c2 >> 4;
        c3 = input.charCodeAt(i++);
        e3 = (c2 & 15) << 2 | c3 >> 6;
        e4 = c3 & 63;
        if (isNaN(c2)) {
          e3 = e4 = 64;
        } else if (isNaN(c3)) {
          e4 = 64;
        }
        o += map.charAt(e1) + map.charAt(e2) + map.charAt(e3) + map.charAt(e4);
      }
      return o;
    },
    decode: function(input) {
      var o = "";
      var c1 = 0, c2 = 0, c3 = 0, e1 = 0, e2 = 0, e3 = 0, e4 = 0;
      input = input.replace(/[^\w\+\/\=]/g, "");
      for (var i = 0; i < input.length; ) {
        e1 = map.indexOf(input.charAt(i++));
        e2 = map.indexOf(input.charAt(i++));
        c1 = e1 << 2 | e2 >> 4;
        o += String.fromCharCode(c1);
        e3 = map.indexOf(input.charAt(i++));
        c2 = (e2 & 15) << 4 | e3 >> 2;
        if (e3 !== 64) {
          o += String.fromCharCode(c2);
        }
        e4 = map.indexOf(input.charAt(i++));
        c3 = (e3 & 3) << 6 | e4;
        if (e4 !== 64) {
          o += String.fromCharCode(c3);
        }
      }
      return o;
    }
  };
}();

import { Dirs, FileSystem } from 'react-native-file-access';
const DDP = Dirs.DocumentDir + "/";
const readFile = (path, enc) => FileSystem.readFile(path, "base64");
const writeFile = (path, data, enc) => FileSystem.writeFile(path, data, "base64");
const input = res => Base64.decode(res);
const output = str => Base64.encode(str);

// react-native-fs
/*
import { writeFile, readFile, DocumentDirectoryPath } from 'react-native-fs';
const DDP = DocumentDirectoryPath + "/";
const input = res => res;
const output = str => str;
*/
// react-native-fetch-blob
/*
import RNFetchBlob from 'react-native-fetch-blob';
const { writeFile, readFile, dirs:{ DocumentDir } } = RNFetchBlob.fs;
const DDP = DocumentDir + "/";
const input = res => res.map(x => String.fromCharCode(x)).join("");
const output = str => str.split("").map(x => x.charCodeAt(0));
*/

const make_cols = refstr => Array.from({length: XLSX.utils.decode_range(refstr).e.c + 1}, (x,i) => XLSX.utils.encode_col(i));
const make_width = refstr => Array.from({length: XLSX.utils.decode_range(refstr).e.c + 1}, () => 60);

export default class SheetJS extends Component {
	constructor(props) {
		super(props);
		this.state = {
			data: [[2,3,4],[3,4,5]],
			widthArr: [60, 60, 60],
			cols: make_cols("A1:C2")
		};
		this.importFile = this.importFile.bind(this);
		this.exportFile = this.exportFile.bind(this);
	};
	importFile() {
		Alert.alert("Rename file to sheetjs.xlsx", "Copy to " + DDP, [
			{text: 'Cancel', onPress: () => {}, style: 'cancel' },
			{text: 'Import', onPress: () => {
				readFile(DDP + "sheetjs.xlsx", 'ascii').then((res) => {
					/* parse file */
					const wb = XLSX.read(input(res), {type:'binary'});

					/* convert first worksheet to AOA */
					const wsname = wb.SheetNames[0];
					const ws = wb.Sheets[wsname];
					const data = XLSX.utils.sheet_to_json(ws, {header:1});

					/* update state */
					this.setState({ data: data, cols: make_cols(ws['!ref']), widthArr: make_width(ws['!ref']) });
				}).catch((err) => { Alert.alert("importFile Error", "Error " + err.message); });
			}}
		]);
	}
	exportFile() {
		/* convert AOA back to worksheet */
		const ws = XLSX.utils.aoa_to_sheet(this.state.data);

		/* build new workbook */
		const wb = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(wb, ws, "SheetJS");

		/* write file */
		const wbout = XLSX.write(wb, {type:'binary', bookType:"xlsx"});
		const file = DDP + "sheetjsw.xlsx";
		writeFile(file, output(wbout), 'ascii').then((res) =>{
				Alert.alert("exportFile success", "Exported to " + file);
		}).catch((err) => { Alert.alert("exportFile Error", "Error " + err.message); });
	};

	render() {
		return (
			<ScrollView contentContainerStyle={styles.container} vertical={true}>
				<Text style={styles.welcome}> </Text>
				<Text style={styles.welcome}>SheetJS React Native Demo</Text>
				<Text style={styles.instructions}>Import Data</Text>
				<Button onPress={this.importFile} title="Import data from a spreadsheet" color="#841584" />
				<Text style={styles.instructions}>Export Data</Text>
				<Button disabled={!this.state.data.length} onPress={this.exportFile} title="Export data to XLSX" color="#841584" />

				<Text style={styles.instructions}>Current Data</Text>

				<ScrollView style={styles.table} horizontal={true} >
					<Table style={styles.table}>
						<TableWrapper>
							<Row data={this.state.cols} style={styles.thead} textStyle={styles.text} widthArr={this.state.widthArr}/>
						</TableWrapper>
						<TouchableWithoutFeedback>
							<ScrollView vertical={true}>
								<TableWrapper>
									<Rows data={this.state.data} style={styles.tr} textStyle={styles.text} widthArr={this.state.widthArr}/>
								</TableWrapper>
							</ScrollView>
						</TouchableWithoutFeedback>
					</Table>
				</ScrollView>
			</ScrollView>
		);
	};
};

const styles = StyleSheet.create({
	container: { flex: 1, justifyContent: 'center', alignItems: 'center', backgroundColor: '#F5FCFF' },
	welcome: { fontSize: 20, textAlign: 'center', margin: 10 },
	instructions: { textAlign: 'center', color: '#333333', marginBottom: 5 },
	thead: { height: 40, backgroundColor: '#f1f8ff' },
	tr: { height: 30 },
	text: { marginLeft: 5 },
	table: { width: "100%" }
});

AppRegistry.registerComponent('SheetJS', () => SheetJS);
