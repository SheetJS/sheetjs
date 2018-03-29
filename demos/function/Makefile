.PHONY: microcule
microcule: mcstream.js
	microcule $<

.PHONY: aws
aws: lambda-proxy

.PHONY: lambda-proxy
lambda-proxy:
	cd LambdaProxy; mkdir -p node_modules; npm install xlsx busboy; sam local start-api; cd -

.PHONY: azure
azure:
	func start

.PHONY: azure-server
azure-server:
	mkdir -p /tmp/azurite
	azurite -l /tmp/azurite

FILES=$(filter-out xlsx.full.min.js,$(wildcard *.js)) $(wildcard *.html)
.PHONY: lint
lint: $(FILES)
	eslint $(FILES)

.PHONY: clean
clean:
	rm -f *.db *.xlsx *.csv
