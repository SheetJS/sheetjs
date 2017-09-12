import Head from 'next/head'
import SheetJSApp from './sheetjs.js'
export default () => (
<div>
	<Head>
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
		<title>SheetJS React Demo</title>
		<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" />
		<script src="https://unpkg.com/babel-standalone@6/babel.min.js"></script>
		<script src="https://cdnjs.cloudflare.com/ajax/libs/react/15.3.1/react.min.js"></script>
		<script src="https://cdnjs.cloudflare.com/ajax/libs/react/15.3.1/react-dom.min.js"></script>
		<script src="https://unpkg.com/xlsx/dist/xlsx.full.min.js"></script>
		<script src="https://unpkg.com/file-saver/FileSaver.js"></script>
		<style jsx>{`
			body, #app { height: 100%; };
		`}</style>
	</Head>
	<div class="container-fluid">
		<h1><a href="http://sheetjs.com">SheetJS React Demo</a></h1>
		<br />
		<a href="https://github.com/SheetJS/js-xlsx">Source Code Repo</a><br />
		<a href="https://github.com/SheetJS/js-xlsx/issues">Issues?  Something look weird?  Click here and report an issue</a><br /><br />
	</div>
	<SheetJSApp />
</div>
)
