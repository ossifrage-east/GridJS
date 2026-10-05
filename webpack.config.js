const HtmlWebpackPlugin = require('html-webpack-plugin');
const MiniCssExtractPlugin = require('mini-css-extract-plugin');
const MiniSVGDataURI = require("mini-svg-data-uri");  // SVG 文件转换为 URI 编码后，与 base64 相比，体积会更小
const { resolve } = require('path')
module.exports = {
  mode: "development",   // 指定 development(开发环境) 和 production(生产环境)
  entry: resolve(__dirname, './src/index.ts'),  // 项目入口
  output: {
    filename: '[name].[hash:8].js', // 输出文件名
    path: resolve(__dirname, './dist'), // 路径
    assetModuleFilename: 'assets/[hash:10][ext][query]',  // 静态文件打包后的路径及文件名
    clean: true, // 打包清除缓存
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
        test: /\.(jpe?g|png|ico|gif|bmp|tiff|webp)$/i,
        type: "asset", // 一般会转换为 "asset/resource"
        parser: {
          dataUrlCondition: {
            maxSize: 8 * 1024 // 8kb （低于8kb都会压缩成 base64）
          }
        }
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
}