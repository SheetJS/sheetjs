.PHONY: vue
vue: ## Simple server for vue
	python -mSimpleHTTPServer

.PHONY: nuxt
nuxt: ## nuxt.js demo
	npm i xlsx @nuxt/content
	npx nuxt

.PHONY: weex
weex: ## Build weex project
	bash ./weex.sh

.PHONY: ios
ios: weex ## weex ios sim
	cd SheetJS; weexpack run ios; cd -
