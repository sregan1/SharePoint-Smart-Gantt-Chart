'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
const webpack = require('webpack');

// pptxgenjs v4 uses "node:*" dynamic imports for its Node.js save paths.
// These code paths never run in the browser, but webpack 5's module resolver
// chokes on the "node:" URI scheme. Strip the prefix so webpack resolves them
// as plain built-ins, then stub those built-ins out with empty modules.
build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    // Replace "node:X" requests with "X" so existing fallback logic applies.
    generatedConfiguration.plugins.push(
      new webpack.NormalModuleReplacementPlugin(/^node:/, (resource) => {
        resource.request = resource.request.replace(/^node:/, '');
      })
    );

    // Stub out Node.js built-ins that pptxgenjs uses only in its Node.js paths.
    generatedConfiguration.resolve = generatedConfiguration.resolve || {};
    generatedConfiguration.resolve.fallback = Object.assign(
      {},
      generatedConfiguration.resolve.fallback,
      { fs: false, https: false, http: false, path: false, zlib: false }
    );

    return generatedConfiguration;
  }
});

build.initialize(gulp);

// In SPFx 1.20+, the classic local serve was renamed to 'serve-deprecated'.
// Re-expose it as 'serve' so the familiar command keeps working.
gulp.task('serve', gulp.series('serve-deprecated'));
