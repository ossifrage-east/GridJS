const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const MiniCssExtractPlugin = require('mini-css-extract-plugin');
const MiniSVGDataURI = require("mini-svg-data-uri");  // SVG 文件转换为 URI 编码后，与 base64 相比，体积会更小
const fs = require('fs');
const { resolve } = require('path')

const DIST = resolve(__dirname, './dist');

// 主配置与 Worker 配置都输出到 dist，若各自开 output.clean 会互相删掉对方的产物，
// 故统一在此处「构建前一次性清空」（dev-server 下不清，避免影响热更新）。
if (!process.env.WEBPACK_SERVE) {
  try { fs.rmSync(DIST, { recursive: true, force: true }); } catch (e) { /* 忽略 */ }
}

// —— 主应用配置 ——
const mainConfig = {
  name: 'main',
  mode: "development",   // 指定 development(开发环境) 和 production(生产环境)
  entry: resolve(__dirname, './src/index.ts'),  // 项目入口
  output: {
    filename: '[name].[hash:8].js', // 输出文件名
    path: DIST, // 路径
    assetModuleFilename: 'assets/[hash:10][ext][query]',  // 静态文件打包后的路径及文件名
    clean: false, // 已在上方统一清空，避免与 worker 配置互相删除产物
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: resolve(__dirname, "./src/index.html"), // 模板文件路径
      inject: 'body'// 在body中生成script标签，默认是head标签内
    }),
    new MiniCssExtractPlugin({
      filename: '[name].[contenthash:8].css',
    }),
    // Pyodide 运行时资产：非 webpack 模块，由 Worker 在运行时按 indexURL 自行拉取，
    // 因此必须原样拷贝到 dist 下的**固定路径** dist/pyodide/（不能带 hash）。
    // 只保留运行时真正会用到的 5 个文件：
    //   pyodide.mjs（ESM 入口，Worker 用动态 import() 加载；classic 脚本 pyodide.js 不要，
    //               Pyodide 314 在 classic worker 内会直接拒绝运行）
    //   pyodide.asm.mjs（Pyodide 自身以动态 import() 拉取）
    //   pyodide.asm.wasm、python_stdlib.zip、pyodide-lock.json
    // 其余（类型声明 / sourcemap / 文档 / 包描述 / pyodide.js / console 示例页）都剔除。
    new CopyWebpackPlugin({
      patterns: [{
        from: resolve(__dirname, './node_modules/pyodide'),
        to: 'pyodide',
        globOptions: {
          ignore: [
            '**/*.d.ts', '**/*.map', '**/README.md', '**/package.json',
            '**/console*.html', '**/pyodide.js'
          ],
        },
      }],
    })
   ],
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: 'ts-loader',
        exclude: /node_modules/
      },
      {
        test: /\.css$/i,
        use: [
          {
            // Adds CSS to the DOM by injecting a `<style>` tag
            loader: MiniCssExtractPlugin.loader
          },
          {
            // Interprets `@import` and `url()` like `import/require()` and will resolve them
            loader: 'css-loader'
          },
          {
            // Loader for webpack to process CSS with PostCSS
            loader: 'postcss-loader',
            options: {
              postcssOptions: {
                plugins: [
                  "autoprefixer"
                ]
              }
            }
          }
        ]
      },
      {
        test: /\.(scss)$/,
        use: [
          {
            // Adds CSS to the DOM by injecting a `<style>` tag
            loader: MiniCssExtractPlugin.loader
          },
          {
            // Interprets `@import` and `url()` like `import/require()` and will resolve them
            loader: 'css-loader'
          },
          {
            // Loader for webpack to process CSS with PostCSS
            loader: 'postcss-loader',
            options: {
              postcssOptions: {
                plugins: [
                  "autoprefixer"
                ]
              }
            }
          },
          {
            // Loads a SASS/SCSS file and compiles it to CSS
            loader: 'sass-loader'
          }
        ]
      },
      {	// 图片文件
        test: /\.(jpe?g|png|gif|bmp|tiff|webp)$/i,
        type: "asset", // 一般会转换为 "asset/resource"
        parser: {
          dataUrlCondition: {
            maxSize: 8 * 1024 // 8kb （低于8kb都会压缩成 base64）
          }
        }
      },
      // 图标文件
      {
        test: /\.(ico|icon)$/i,
        type: "asset/resource", // 一般会转换为 "asset/resource"
      },
      // svg文件
      {
        test: /\.svg$/i,
        type: "asset/inline",
        generator: {
          dataUrl: content => {
            content = content.toString();
            return MiniSVGDataURI(content);
          }
        },
        parser: {
          dataUrlCondition: {
            maxSize: 2 * 1024 // 2kb （低于2kb都会压缩）
          }
        }
      },
      // 字体文件
      {
        test: /\.(otf|eot|woff2?|ttf|svg)$/i,
        type: "asset", // 一般会转换为 "asset/inline"
      },
      // 处理视频文件
      {
        test: /\.(mp4|mp3|wav|webm|ogg)$/,
        type: 'asset/resource',
        generator: {
          filename: '[name][ext]'
        }
      }
    ]
  },
  resolve: {
    extensions: ['.ts', '.js'],
    fallback: {
      'fs': false,
      'path': false,
      'vm': false
    }
  },
};

// —— Excel 解析 Worker 配置 ——
// 只打包 parser.worker.ts 这一个入口，输出为固定文件名 parser.worker.js
// （parserClient 以相对文档路径 `new Worker('parser.worker.js')` 加载）。
// target: 'webworker' 让 webpack 按 Worker 环境生成运行时（不注入 document 相关代码）。
const workerConfig = {
  name: 'parser-worker',
  mode: "development",
  target: 'webworker',
  entry: resolve(__dirname, './src/utils/parser/parser.worker.ts'),
  output: {
    filename: 'parser.worker.js',
    path: DIST,
    clean: false, // 与主配置共用 dist，清空已在上方统一完成
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: 'ts-loader',
        exclude: /node_modules/
      }
    ]
  },
  resolve: {
    extensions: ['.ts', '.js'],
    fallback: {
      'fs': false,
      'path': false,
      'vm': false
    }
  },
};

// —— Pyodide 执行 Worker 配置 ——
// 与解析 Worker 同构：只打包 python.worker.ts 这一个入口，输出固定文件名
// python.worker.js（pythonClient 以 `new Worker('python.worker.js', { type: 'module' })`
// 加载）。Pyodide 运行时（wasm / stdlib）不参与打包，由 Worker 在运行时按
// indexURL 从 dist/pyodide/ 拉取（见上方 CopyWebpackPlugin）。
//
// **必须产出为 ES module**：Pyodide 314 会拒绝在 classic worker 中初始化
// （见 python.worker.ts 注释），而 classic 与否由 Worker 的构造方式决定，
// 故这里开 outputModule，让 python.worker.js 是 ESM 文件。
const pythonWorkerConfig = {
  name: 'python-worker',
  mode: "development",
  target: 'webworker',
  experiments: {
    outputModule: true,
  },
  entry: resolve(__dirname, './src/utils/python/python.worker.ts'),
  output: {
    filename: 'python.worker.js',
    path: DIST,
    clean: false, // 与主配置共用 dist，清空已在上方统一完成
    module: true,
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: 'ts-loader',
        exclude: /node_modules/
      }
    ]
  },
  resolve: {
    extensions: ['.ts', '.js'],
    fallback: {
      'fs': false,
      'path': false,
      'vm': false
    }
  },
};

module.exports = [mainConfig, workerConfig, pythonWorkerConfig];