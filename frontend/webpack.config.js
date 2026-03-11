const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');

module.exports = {
  mode: 'development',
  devtool: 'cheap-module-source-map',
  entry: './src/taskpane.tsx',
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: 'taskpane.js',
    clean: true,
  },
  resolve: {
    extensions: ['.tsx', '.ts', '.jsx', '.js'],
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: 'ts-loader',
        exclude: /node_modules/,
      },
      {
        test: /\.module\.css$/,
        use: [
          'style-loader',
          { loader: 'css-loader', options: { modules: { localIdentName: '[name]__[local]--[hash:base64:5]' } } },
        ],
      },
      {
        test: /\.css$/,
        exclude: /\.module\.css$/,
        use: ['style-loader', 'css-loader'],
      },
    ],
  },
  devServer: {
    server: 'http',
    port: 3000,
    hot: true,
    static: {
      directory: path.join(__dirname, 'dist'),
    },
    headers: {
      'Access-Control-Allow-Origin': '*',
    },
  },
  performance: false,
  plugins: [
    new HtmlWebpackPlugin({
      template: './assets/taskpane.html',
      filename: 'index.html',
      inject: false,
    }),
  ],
};
