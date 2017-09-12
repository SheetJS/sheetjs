.PHONY: vue
vue: ## Simple server for vue
	python -mSimpleHTTPServer

.PHONY: nuxt
nuxt: ## nuxt.js demo
	mkdir -p node_modules
	cd node_modules; if [ ! -e xlsx ]; then ln -s ../../../ xlsx; fi; cd ..
	nuxt

.PHONY: weex
weex: ## Build weex project
	bash ./weex.sh

.PHONY: ios
ios: weex ## weex ios sim
	cd SheetJS; weexpack run ios; cd -
