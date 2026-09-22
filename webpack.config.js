/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");

const urlDev = "https://localhost:3000/";
const defaultProductionBaseUrl = "https://your-org.github.io/your-repo/";
const defaultProductionOrigin = "https://your-org.github.io";

function normalizeBaseUrl(value) {
  return value.endsWith("/") ? value : `${value}/`;
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  env = env || {};
  const dev = options.mode === "development";
  const productionBaseUrl = normalizeBaseUrl(
    env.productionUrl || process.env.CLASSIFYME_PRODUCTION_BASE_URL || defaultProductionBaseUrl
  );
  const productionOrigin = new URL(productionBaseUrl).origin;
  const outputDirectory = env.githubPages ? "docs" : "dist";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/ui/taskpane/taskpane.ts", "./src/ui/taskpane/taskpane.html"],
      commands: "./src/commands/commands.ts",
      autoopen: "./src/hosts/outlook/outlookLaunchEvents.ts",
    },
    output: {
      path: path.resolve(__dirname, outputDirectory),
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader"
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/ui/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              }

              return content
                .toString()
                .replace(new RegExp(escapeRegExp(urlDev), "g"), productionBaseUrl)
                .replace(new RegExp(escapeRegExp(defaultProductionBaseUrl), "g"), productionBaseUrl)
                .replace(new RegExp(escapeRegExp(defaultProductionOrigin), "g"), productionOrigin);
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new HtmlWebpackPlugin({
        filename: "autoopen.html",
        template: "./src/hosts/outlook/outlookLaunchEvents.html",
        chunks: ["autoopen"],
        // Le runtime web doit enregistrer les handlers avant de pouvoir recevoir les événements.
        scriptLoading: "blocking",
        // Le hash évite qu'Outlook Web réutilise un ancien bundle pendant le diagnostic.
        hash: true,
      }),
    ],
    devServer: {
      /*
       * Outlook Classic charge autoopen.js dans un runtime JavaScript-only.
       * Le client HMR de webpack-dev-server utilise des API de navigateur et était
       * exécuté avant les handlers, ce qui empêchait leur enregistrement.
       */
      client: false,
      hot: false,
      liveReload: false,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
