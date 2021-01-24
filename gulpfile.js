/**
 * xlsx-js-style
 */
const pkg = require("./package.json");
const gulp = require("gulp"),
	concat = require("gulp-concat"),
	ignore = require("gulp-ignore"),
	insert = require("gulp-insert"),
	source = require("gulp-sourcemaps"),
	uglify = require("gulp-uglify");

gulp.task("min", () => {
	return gulp
		.src(["./src/xlsx.js"])
		.pipe(concat("xlsx.min.js"))
		.pipe(uglify())
		.pipe(insert.prepend("/* xlsx-js-style " + pkg.version + " @ " + new Date().toISOString() + " */\n"))
		.pipe(source.init())
		.pipe(ignore.exclude(["**/*.map"]))
		.pipe(source.write("./"))
		.pipe(gulp.dest("./dist/"));
});

gulp.task("bundle", () => {
	return gulp
		.src(["./libs/*", "./src/xlsx.js"])
		.pipe(concat("xlsx.bundle.js"))
		.pipe(uglify())
		.pipe(insert.prepend("/* xlsx-js-style " + pkg.version + " @ " + new Date().toISOString() + " */\n"))
		.pipe(source.init())
		.pipe(ignore.exclude(["**/*.map"]))
		.pipe(source.write("./"))
		.pipe(gulp.dest("./dist/"));
});

// Build/Deploy (ad-hoc, no watch)
gulp.task("ship", gulp.series("min", "bundle"), () => {
	console.log("... ./dist/*.js files created!");
});
