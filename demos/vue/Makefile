.PHONY: vue
vue: ## Simple server for vue
	python -mSimpleHTTPServer

.PHONY: nuxt
nuxt: ## nuxt.js demo
	npm i https://cdn.sheetjs.com/xlsx-latest/xlsx-latest.tgz @nuxt/content
	npx nuxt

.PHONY: weex
weex: ## Build weex project
	bash ./weex.sh

.PHONY: ios
ios: weex ## weex ios sim
	cd SheetJS; weexpack run ios; cd -
