/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */

const Koa = require('koa'), app = new Koa();
const { sprintf } = require('printj');
const { IncomingForm } = require('formidable');
const { fork } = require('child_process');
const logit = require('./_logit');
const subprocess = fork('koasub.js');

const get_data = async (ctx, type) => {
	await new Promise((resolve, reject) => {
		const cb = (data) => {
			ctx.response.body = Buffer(data);
			subprocess.removeListener('message', cb);
			resolve();
		};
		subprocess.on('message', cb);
		subprocess.send(['get data', type]);
	});
};

const get_file = async (ctx, file) => {
	await new Promise((resolve, reject) => {
		const cb = (data) => {
			ctx.response.body = Buffer(data);
			subprocess.removeListener('message', cb);
			resolve();
		};
		subprocess.on('message', cb);
		subprocess.send(['get file', file]);
	});
};

const load_data = async (ctx, file) => {
	await new Promise((resolve, reject) => {
		const cb = (data) => {
			ctx.response.body = "ok\n";
			subprocess.removeListener('message', cb);
			resolve();
		};
		subprocess.on('message', cb);
		subprocess.send(['load data', file]);
	});
};

const post_data = async (ctx) => {
	const keys = Object.keys(ctx.request._files), k = keys[0];
	await load_data(ctx, ctx.request._files[k].path);
};

app.use(async (ctx, next) => { logit(ctx.req, ctx.res); await next(); });
app.use(async (ctx, next) => {
	const form = new IncomingForm();
	await new Promise((resolve, reject) => {
		form.parse(ctx.req, (err, fields, files) => {
			if(err) return reject(err);
			ctx.request._fields = fields;
			ctx.request._files = files;
			resolve();
		});
	});
	await next();
});
app.use(async (ctx, next) => {
	if(ctx.request.method !== 'GET') await next();
	else if(ctx.request.path !== '/') await next();
	else if(ctx.request.query.t) await get_data(ctx, ctx.request.query.t);
	else if(ctx.request.query.f) await get_file(ctx, ctx.request.query.f);
	else ctx.throw(403, "Forbidden");
});
app.use(async (ctx, next) => {
	if(ctx.request.method !== 'POST') await next();
	else if(ctx.request.path !== '/') await next();
	else if(ctx.request.query.f) await load_data(ctx, ctx.request.query.f);
	else await post_data(ctx);
});

const port = +process.argv[2] || +process.env.PORT || 7262;
app.listen(port, () => { console.log('Serving HTTP on port ' + port); });
