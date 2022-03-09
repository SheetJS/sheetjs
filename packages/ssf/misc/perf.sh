#!/bin/bash
ARGS="--trace_opt --trace_deopt --trace_inlining --code_comments"
#ARGS="--trace_opt --trace_deopt --code_comments"
SCPT=misc/perf.js

echo 1
make && jshint --show-non-errors ssf.js && make lint &&
MINTEST=1 mocha -b && time node $SCPT && {
node $ARGS $SCPT > perf.log
node --prof $SCPT
echo 1; time node $SCPT >/dev/null
echo 1; time node $SCPT >/dev/null
} && grep disabled perf.log
