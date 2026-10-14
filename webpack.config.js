const HtmlWebpackPlugin = require('html-webpack-plugin');
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

module.exports = [mainConfig, workerConfig];