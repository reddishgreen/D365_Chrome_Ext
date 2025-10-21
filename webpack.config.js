const path = require('path');
const CopyPlugin = require('copy-webpack-plugin');
const HtmlWebpackPlugin = require('html-webpack-plugin');

module.exports = {
  mode: 'production',
  entry: {
    content: './src/content/index.tsx',
    injected: './src/content/injected.ts',
    'webapi-viewer': './src/webapi-viewer/index.tsx',
    popup: './src/popup/index.tsx'
  },
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: '[name].bundle.js',
    clean: true
  },
  devtool: false,
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: 'ts-loader',
        exclude: /node_modules/
      },
      {
        test: /\.css$/,
        use: ['style-loader', 'css-loader']
      }
    ]
  },
  resolve: {
    extensions: ['.tsx', '.ts', '.js']
  },
  plugins: [
    new CopyPlugin({
      patterns: [
        { from: 'manifest.json', to: 'manifest.json' },
        { from: 'src/content/styles.css', to: 'content.css' },
        { from: 'icons', to: 'icons', noErrorOnMissing: true }
      ]
    }),
    new HtmlWebpackPlugin({
      template: './src/webapi-viewer/index.html',
      filename: 'webapi-viewer.html',
      chunks: ['webapi-viewer']
    }),
    new HtmlWebpackPlugin({
      template: './src/popup/index.html',
      filename: 'popup.html',
      chunks: ['popup']
    })
  ]
};
