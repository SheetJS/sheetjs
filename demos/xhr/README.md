# XMLHttpRequest and Friends

`XMLHttpRequest` and `fetch` browser APIs enable binary data transfer between
web browser clients and web servers.  Since this library works in web browsers,
server conversion work can be offloaded to the client!  This demo shows a few
common scenarios involving browser APIs and popular wrapper libraries.

## Sample Server

The `server.js` nodejs server serves static files on `GET` request.  On a `POST`
request to `/upload`, the server processes the body and looks for the `file` and
`data` fields.  It will write the Base64-decoded data from `data` to the file
name specified in `file`.

To start the demo, run `npm start` and navigate to <http://localhost:7262/>

## XMLHttpRequest (xhr.html)

For downloading data, the `arraybuffer` response type generates an `ArrayBuffer`
that can be viewed as an `Uint8Array` and fed to `XLSX.read` using `array` type.

For uploading data, this demo populates a `FormData` object with string data
generated with the `base64` output type.

## axios (axios.html) and superagent (superagent.html)

The codes are structurally similar to the XMLHttpRequest example.  `axios` uses
a Promise-based API while `superagent` opts for a more traditional chain.

## fetch (fetch.html)

For downloading data, `response.blob()` resolves to a `Blob` object that can be
converted to `ArrayBuffer` using a `FileReader`.

