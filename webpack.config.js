/* eslint-disable no-undef */
/* eslint-disable no-unused-vars */
const devCerts = require("office-addin-dev-certs");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const ExtractTextPlugin = require('extract-text-webpack-plugin');
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require('webpack');
const path = require("path");

module.exports = async (env, options) => {
    const dev = options.mode === "development";
    const config = {
        devtool: "source-map",
        entry: {
            vendor: [
                'react',
                'react-dom',
                'core-js',
                'office-ui-fabric-react'
              ],
              taskpane: [
                  'react-hot-loader/patch',
                  './src/taskpane/index.tsx',
              ],
            commands: "./src/commands/commands.ts",
            fallbackauthdialog: "./src/helpers/fallbackauthdialog.ts",
            polyfill: "@babel/polyfill",
            //taskpane: "./src/taskpane/taskpane.ts",
        },
        output: {
            path: path.resolve(process.cwd(), 'dist'),
        },
        resolve: {
            extensions: [".ts", ".tsx", ".html", ".js"]
        },
        module: {
            rules: [
                {
                    test: /\.ts$/,
                    exclude: /node_modules/,
                    use: "babel-loader"
                },
                {
                    test: /\.tsx?$/,
                    exclude: /node_modules/,
                    use: "ts-loader"
                },
                {
                    test: /\.html$/,
                    exclude: /node_modules/,
                    use: "html-loader"
                },
                {
                    test: /\.(png|jpe?g|gif|svg|woff|woff2|ttf|eot|ico)$/,
                    use: {
                        loader: 'file-loader',
                        query: {
                            name: 'assets/[name].[ext]'
                          }
                        }  
                },
//                {
//                    test: /\.(png|jpg|jpeg|gif)$/,
//                    use: "file-loader"
//                },
                {
                    test: /\.css$/,
                    use: ['style-loader', 'css-loader']
                }
            ]
        },
        plugins: [
            new CleanWebpackPlugin(),
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
            new HtmlWebpackPlugin({
                filename: "fallbackauthdialog.html",
                template: "./src/helpers/fallbackauthdialog.html",
                chunks: ["polyfill", "fallbackauthdialog"]
            }),
            new CopyWebpackPlugin([
                {
                    to: "taskpane.css",
                    from: "./src/taskpane/taskpane.css"
                }
            ]),
            new ExtractTextPlugin('[name].[hash].css'),
            new HtmlWebpackPlugin({
                filename: "taskpane.html",
                template: './src/taskpane/taskpane.html',
                chunks: ['taskpane', 'vendor', 'polyfills']
            }),
            new CopyWebpackPlugin([
                {
                    from: './assets',
                    ignore: ['*.scss'],
                    to: 'assets',
                }
            ]),
            new webpack.ProvidePlugin({
                Promise: ["es6-promise", "Promise"]
            })
        ]
    };

    return config;
};