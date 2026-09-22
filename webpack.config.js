/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");

const urlDev = "https://localhost:3000/";
const templateProductionBaseUrl = "https://your-org.github.io/your-repo/";
const templateProductionOrigin = "https://your-org.github.io";
const defaultProductionBaseUrl = "https://MindGroup29.github.io/ClassifyMe/";

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

function getTypeScriptRule() {
  return {
    test: /\.ts$/,
    exclude: /node_modules/,
    use: {
      loader: "babel-loader",
    },
  };
}

module.exports = async (env, options) => {
  env = env || {};
  const dev = options.mode === "development";
  const productionBaseUrl = normalizeBaseUrl(
    env.productionUrl || process.env.CLASSIFYME_PRODUCTION_BASE_URL || defaultProductionBaseUrl
  );
  const productionOrigin = new URL(productionBaseUrl).origin;
  const outputDirectory = env.githubPages ? "docs" : "dist";
  const outputPath = path.resolve(__dirname, outputDirectory);

  const applicationConfig = {
    name: "application",
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/ui/taskpane/taskpane.ts", "./src/ui/taskpane/taskpane.html"],
      commands: "./src/commands/commands.ts",
    },
    output: {
      path: outputPath,
      // En mode serveur, un rebuild du task pane ne doit pas effacer le bundle Outlook.
      clean: !process.env.WEBPACK_SERVE,
    },
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },
    module: {
      rules: [
        getTypeScriptRule(),
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
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
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
                .replace(
                  new RegExp(escapeRegExp(templateProductionBaseUrl), "g"),
                  productionBaseUrl
                )
                .replace(
                  new RegExp(escapeRegExp(templateProductionOrigin), "g"),
                  productionOrigin
                );
            },
          },
        ],
      }),
    ],
    devServer: {
      // Le runtime Outlook est précompilé sur disque avant le démarrage du serveur.
      static: {
        directory: outputPath,
      },
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options:
          env.WEBPACK_BUILD || options.https !== undefined
            ? options.https
            : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  const reminderRuntimeConfig = {
    name: "outlook-classification-reminder",
    dependencies: ["application"],
    devtool: "source-map",
    entry: {
      "outlook-classification-reminder":
        "./src/hosts/outlook/outlookClassificationReminder.ts",
    },
    output: {
      path: outputPath,
    },
    resolve: {
      extensions: [".ts", ".js"],
    },
    module: {
      rules: [getTypeScriptRule()],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "outlook-classification-reminder.html",
        template: "./src/hosts/outlook/outlookClassificationReminder.html",
        chunks: ["outlook-classification-reminder"],
        // Les handlers doivent être associés dès le chargement du runtime Web/New Outlook.
        scriptLoading: "blocking",
        // Le hash empêche la réutilisation d'un ancien bundle par le cache Outlook.
        hash: true,
      }),
    ],
    /*
     * Outlook Classic exécute ce bundle dans un runtime JavaScript-only. Il est compilé
     * avant le lancement du serveur, puis servi comme ressource statique. L'exclusion de
     * webpack-dev-server empêche toute injection WebSocket/HMR dans ce bundle.
     */
    devServer: false,
  };

  return [applicationConfig, reminderRuntimeConfig];
};
