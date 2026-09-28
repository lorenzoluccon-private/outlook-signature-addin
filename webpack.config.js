/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3001/";
const urlProd = "https://signature.fgautomazioni.it/addin-graph-pilot/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();

  return {
    ca: httpsOptions.ca,
    key: httpsOptions.key,
    cert: httpsOptions.cert,
  };
}

module.exports = async (env, options) => {
  const isDevelopment = options.mode === "development";

  return {
    mode: options.mode || "development",

    devtool: isDevelopment ? "source-map" : false,

    entry: {
      polyfill: [
        "core-js/stable",
        "regenerator-runtime/runtime",
      ],

      taskpane: [
        "./src/taskpane/taskpane.ts",
        "./src/taskpane/taskpane.html",
      ],

      commands: "./src/commands/commands.ts",

      autorun: "./src/runtime/autorun.ts",
    },

    output: {
      clean: true,
      filename: "[name].js",
    },

    resolve: {
      extensions: [".ts", ".js", ".html"],
    },

    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: {
            loader: "html-loader",
          },
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico|svg)$/i,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },

    plugins: [
      /*
       * Task pane dell'add-in.
       */
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: [
          "polyfill",
          "taskpane",
        ],
        inject: "body",
      }),

      /*
       * Runtime dei pulsanti manuali Outlook.
       */
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: [
          "polyfill",
          "commands",
        ],
        inject: "body",
      }),

      /*
       * Runtime leggero usato esclusivamente
       * dall'evento OnNewMessageCompose.
       *
       * Non include polyfill.js, così riduciamo
       * il numero di file e i dati caricati quando
       * viene creata una nuova email.
       */
      new HtmlWebpackPlugin({
        filename: "autorun.html",
        template: "./src/runtime/autorun.html",
        chunks: [
          "autorun",
        ],
        inject: "body",
      }),

      /*
       * Copia esclusivamente le risorse grafiche.
       *
       * I manifest sono conservati separatamente
       * e non devono essere pubblicati dentro dist.
       */
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "public/.well-known/microsoft-officeaddins-allowed.json",
            to: ".well-known/microsoft-officeaddins-allowed.json",
          },          
        ],
      }),
    ],

    optimization: {
      /*
       * Mantiene separati i bundle commands,
       * taskpane e autorun.
       */
      splitChunks: false,
      runtimeChunk: false,

      /*
       * Rimuove codice non utilizzato durante
       * la build di produzione.
       */
      usedExports: true,
    },

    devServer: {
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

      port:
        process.env.npm_package_config_dev_server_port ||
        3001,

      hot: false,

      liveReload: true,

      proxy: [
        {
          context: ["/signature-api"],
          target: "https://stackrapid.fgautomazioni.it",
          changeOrigin: true,
          secure: true,
          pathRewrite: {
            "^/signature-api": "/stackrapid/api",
          },
        },
      ],

      devMiddleware: {
        writeToDisk: false,
      },
    },
  };
};