const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");

module.exports = (env, argv) => {
  const isDev = argv.mode === "development";
  return {
    entry: {
      taskpane: "./src/index.tsx",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].bundle.js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.(ts|tsx)$/,
          use: "ts-loader",
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: "./src/index.html",
        chunks: ["taskpane"],
        filename: "taskpane.html",
      }),
    ],
    devServer: {
      port: 3000,
      https: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      static: {
        directory: path.join(__dirname, "dist"),
      },
    },
    devtool: isDev ? "source-map" : false,
  };
};
