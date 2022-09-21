#!/bin/bash
set -eo pipefail
INF=${1:-test.numbers}
OUTF=${2:-reframed.numbers}
cp "$INF" "$OUTF"
chmod a-w "$OUTF"
sleep 0.1
# open "$OUTF"
unzip -l "$OUTF"
base64 "$OUTF" | tr -d '\n' > xlsx.zahl.js
sed -i.bak 's/^/var XLSX_ZAHL_PAYLOAD = "/g;s/$/";\n/g' xlsx.zahl.js
cp xlsx.zahl.js xlsx.zahl.mjs
cat >> xlsx.zahl.js <<EOF
if(typeof module !== "undefined") module.exports = XLSX_ZAHL_PAYLOAD;
else if(typeof global !== "undefined") global.XLSX_ZAHL_PAYLOAD = XLSX_ZAHL_PAYLOAD;
else if(typeof window !== "undefined") window.XLSX_ZAHL_PAYLOAD = XLSX_ZAHL_PAYLOAD;
EOF
cat >> xlsx.zahl.mjs <<EOF
export default XLSX_ZAHL_PAYLOAD;
EOF
