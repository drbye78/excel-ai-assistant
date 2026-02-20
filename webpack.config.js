const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
require("dotenv").config({ path: path.resolve(__dirname, ".env") });

const urlDev = "https://localhost:3000/";
const urlProd = process.env.PRODUCTION_URL || process.env.OPENAI_API_URL?.replace("/v1", "") || "https://localhost:3000/";

const getHttpsOptions = () => {
  const devCerts = require("office-addin-dev-certs");
  return devCerts.getHttpsServerOptions();
};

module.exports = async (env, options) => {
  const isDev = options.mode === "development";
  const config = {
    // devtool: isDev ? "source-map" : false,
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: "./src/taskpane/index.tsx",
      commands: "./src/commands/commands.ts"
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
      alias: {
        "@": path.resolve(__dirname, "src"),
        "@components": path.resolve(__dirname, "src/components"),
        "@services": path.resolve(__dirname, "src/services"),
        "@utils": path.resolve(__dirname, "src/utils"),
        "@types": path.resolve(__dirname, "src/types")
      }
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: [
                ["@babel/preset-env", { targets: { ie: "11" } }],
                "@babel/preset-react",
                "@babel/preset-typescript"
              ]
            }
          }
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"]
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader"
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]"
          }
        }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"]
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "./assets",
            to: "assets",
            noErrorOnMissing: true
          },
          {
            from: "./manifest.xml",
            to: "manifest.xml"
          }
        ]
      })
    ],
    devServer: {
      static: {
        directory: path.join(__dirname, "dist"),
        publicPath: "/"
      },
      headers: {
        "Access-Control-Allow-Origin": "*"
      },
      server: {
        type: "https",
        options: isDev ? await getHttpsOptions() : undefined
      },
      port: 3000,
      hot: true
    },
    output: {
      clean: true,
      filename: "[name].js",
      path: path.resolve(__dirname, "dist")
    }
  };

  return config;
};
