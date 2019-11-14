// https://www.tiny.cloud/docs/advanced/usage-with-module-loaders/
const merge = require("webpack-merge"),
  TerserPlugin = require("terser-webpack-plugin"),
  { CleanWebpackPlugin } = require("clean-webpack-plugin"),
  common = require("./config.common.js");

module.exports = merge(common, {
  mode: "production",
  devtool: false,
  optimization: {
    minimize: true,
    minimizer: [
      new TerserPlugin({
        terserOptions: {
          output: { ascii_only: true }
        }
      })
    ]
  },
  plugins: [new CleanWebpackPlugin()]
});
