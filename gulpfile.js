/**
 * xlsx-js-style
 */
const pkg = require('./package.json')
const gulp = require('gulp'),
	concat = require('gulp-concat'),
	ignore = require('gulp-ignore'),
	insert = require('gulp-insert'),
	source = require('gulp-sourcemaps'),
	uglify = require('gulp-uglify')
