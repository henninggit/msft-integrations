const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = (env, options) => {
  const isDev = options.mode === 'development';
  
  return {
    entry: {
      taskpane: './src/taskpane/taskpane.js'
    },
    
    output: {
      path: path.resolve(__dirname, 'dist'),
      filename: '[name].js',
      clean: true
    },
    
    devServer: {
      port: 3000,
      host: 'localhost',
      hot: true,
      open: false,
      // server: 'https', // Commented out for SSL troubleshooting
      headers: {
        'Access-Control-Allow-Origin': '*'
      }
    },
    
    plugins: [
      new HtmlWebpackPlugin({
        template: './src/taskpane/taskpane.html',
        filename: 'taskpane.html',
        chunks: ['taskpane']
      }),
      
      new HtmlWebpackPlugin({
        template: './src/commands/commands.html',
        filename: 'commands.html',
        chunks: []
      }),
      
      new CopyWebpackPlugin({
        patterns: [
          {
            from: 'assets',
            to: 'assets'
          },
          {
            from: 'manifest.xml',
            to: 'manifest.xml'
          }
        ]
      })
    ],
    
    resolve: {
      extensions: ['.js', '.html']
    },
    
    module: {
      rules: [
        {
          test: /\.css$/,
          use: ['style-loader', 'css-loader']
        },
        {
          test: /\.html$/i,
          loader: 'html-loader',
        }
      ]
    }
  };
};
