/**
 * @license
 * Copyright 2024 Google LLC.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

const path = require('path');
const GasPlugin = require('gas-webpack-plugin');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const HtmlInlineScriptPlugin = require('html-inline-script-webpack-plugin');

module.exports = {
  entry : {
    'curvesmith' : path.resolve(__dirname, 'curvesmith.ts'),
    'sidebar' : path.resolve(__dirname, 'assets/sidebar.ts'),
    'upload' : path.resolve(__dirname, 'assets/upload.ts')
  },
  mode : 'production',
  module : {
    rules :
          [
            {
              test : /\.tsx?$/,
              use : 'ts-loader',
              exclude : /node_modules/,
            },
            {test : /\.css$/, use : ['style-loader', 'css-loader']},
            {
              test : /\.html$/,
              loader : 'html-loader',
            },
            {
              test : /\.m?js/,  // Fix for issue with safevalues library
              resolve : {
                fullySpecified : false,
              },
            },
          ],
  },
  resolve : {
    extensions : ['.tsx', '.ts', '.js'],
  },
  output : {
    libraryTarget : 'this',
    filename : '[name].js',
    path : path.resolve(__dirname, 'dist'),
  },
  optimization : {minimize : false},
  plugins :
          [
            new GasPlugin({autoGlobalExportsFiles : ['**/*.ts']}),
            new HtmlWebpackPlugin({
              template : 'assets/sidebar.html',
              filename : 'sidebar.html',
              inject : 'head',
              chunks : ['sidebar'],
              scriptLoading : 'blocking',
            }),
            new HtmlWebpackPlugin({
              template : 'assets/upload.html',
              filename : 'upload.html',
              inject : 'head',
              chunks : ['upload'],
              scriptLoading : 'blocking',
            }),
            new HtmlInlineScriptPlugin({
              htmlMatchPattern : [/(sidebar|upload).html$/],
            }),
          ],
};