module.exports = {
	head: {
		script: [
			// { src: "https://unpkg.com/xlsx/dist/shim.min.js" }, // CDN
			// { src: "https://unpkg.com/xlsx/dist/xlsx.full.min.js" } // CDN
			{ src: "shim.js" }, // development
			{ src: "xlsx.full.min.js" } // development
		]
	}
};
