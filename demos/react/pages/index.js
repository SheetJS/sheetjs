import Head from 'next/head'
import SheetJSApp from './sheetjs.js'
export default () => (
<div>
	<Head>
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
		<title>SheetJS React Demo</title>
		<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" />
		<style jsx>{`
			body, #app { height: 100%; };
		`}</style>
	</Head>
	<div className="container-fluid">
		<h1><a href="http://sheetjs.com">SheetJS React Demo</a></h1>
		<br />
		<a href="https://github.com/SheetJS/js-xlsx">Source Code Repo</a><br />
		<a href="https://github.com/SheetJS/js-xlsx/issues">Issues?  Something look weird?  Click here and report an issue</a><br /><br />
	</div>
	<SheetJSApp />
</div>
)
