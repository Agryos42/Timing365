const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');

module.exports = (env, argv) => {
  const mode = argv.mode || 'production';
  return {
    mode,
    entry: './src/index.tsx',
    devtool: mode === 'development' ? 'inline-source-map' : false,
    resolve: {
      extensions: ['.ts', '.tsx', '.js']
    },
    output: {
      filename: 'bundle.js',
      path: path.resolve(__dirname, 'build'),
      clean: true
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
    plugins: [
      new HtmlWebpackPlugin({
        template: path.resolve(__dirname, 'public/taskpane.html'),
        filename: 'taskpane.html'
      })
    ],
    devServer: {
      static: {
        directory: path.join(__dirname, 'public')
      },
      port: 3000
    }
  };
};
