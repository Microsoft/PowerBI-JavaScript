var gulp = require('gulp-help')(require('gulp'));
var rename = require('gulp-rename'),
    concat = require('gulp-concat'),
    uglify = require('gulp-uglify'),
    jshint = require('gulp-jshint'),
    replace = require('gulp-replace'),
    sourcemaps = require('gulp-sourcemaps'),
    tslint = require("gulp-tslint"),
    ts = require('gulp-typescript'),
    rimraf = require('rimraf'),
    merge = require('merge2'),
    karma = require('karma'),
    paths = require('./paths'),
    webpack = require('webpack-stream'),
    webpackConfig = require('./webpack.config'),
    runSequence = require('run-sequence'),
    argv = require('yargs').argv;
    ;

gulp.task('watch', 'Watches for changes', ['lint'], function () {
    gulp.watch(['./src/**/*.ts', './test/**/*.ts'], ['lint:ts']);
    gulp.watch(['./test/**/*.ts'], ['test']);
});

gulp.task('lint', 'Lints all files', function (done) {
    runSequence(
        'lint:ts',
        done
    );
});

gulp.task('test', 'Runs all tests', function (done) {
    runSequence(
        'clean',
        ['compile:ts', 'compile:spec'],
        ['test:js', 'copy', 'min:js'],
        done
    );
});

gulp.task('build', 'Runs a full build', function (done) {
    runSequence(
        'lint',
        'clean',
        ['compile:ts', 'compile:dts'],
        'min:js',
        done
    );
});

gulp.task('clean', 'Cleans destination folder', function(done) {
    rimraf(paths.jsDest, done);
});

gulp.task('lint:ts', 'Lints all TypeScript', function() {
    return gulp.src(['./src/**/*.ts', './test/**/*.ts'])
        .pipe(tslint())
        .pipe(tslint.report("verbose"));
});

gulp.task('min:js', 'Creates minified JavaScript file', function() {
    return gulp.src(paths.jsDest + paths.outFilename)
        .pipe(sourcemaps.init({ debug: true }))
        .pipe(uglify())
        .pipe(sourcemaps.write())
        .pipe(rename({ suffix: '.min' }))
        .pipe(gulp.dest(paths.jsDest));
});

gulp.task('test:js', 'Runs unit tests', function(done) {
    new karma.Server.start({
        configFile: __dirname + '/karma.conf.js',
        singleRun: argv.debug ? false : true,
        captureTimeout: argv.timeout || 60000
    }, done);
});

gulp.task('compile:ts', 'Compile typescript for powerbi library', function() {
    return gulp.src(['./src/powerbi.ts'])
        .pipe(webpack(webpackConfig))
        .pipe(gulp.dest('dist/'));
});

gulp.task('compile:dts', 'Generate single dts file from modules', function (done) {
    var tsResult = gulp.src(['./src/**/*.ts'])
        .pipe(ts({
            outDir: 'dts',
            declaration: true
        }));
    
    return tsResult.dts
        .pipe(replace(/import[^;]+;/, ''))
        .pipe(concat('powerbi.d.ts'))
        .pipe(gulp.dest('dist/'));
});

gulp.task('compile:spec', 'Compile typescript for tests', function () {
    var tsProject = ts.createProject('tsconfig.json');
    
    var tsResult = gulp.src(['./test/**/*.ts', './typings/browser/**/*.d.ts'])
        .pipe(ts(tsProject))
        ;
        
    return tsResult.js.pipe(gulp.dest('./test'));
});