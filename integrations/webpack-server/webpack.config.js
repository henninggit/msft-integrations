const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const fs = require('fs');

module.exports = (env, options) => {
  const isDev = options.mode === 'development';
  
  // SSL certificate paths for Office dev
  const certPath = path.join(require('os').homedir(), '.office-addin-dev-certs', 'localhost.crt');
  const keyPath = path.join(require('os').homedir(), '.office-addin-dev-certs', 'localhost.key');
  
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
      host: '0.0.0.0', // Bind to all interfaces
      hot: false, // Disable hot reload to avoid WebSocket issues
      liveReload: false,
      open: false,
      server: {
        type: 'https',
        options: fs.existsSync(certPath) && fs.existsSync(keyPath) ? {
          cert: fs.readFileSync(certPath),
          key: fs.readFileSync(keyPath),
          ca: fs.existsSync(path.join(require('os').homedir(), '.office-addin-dev-certs', 'ca.crt')) ? 
              fs.readFileSync(path.join(require('os').homedir(), '.office-addin-dev-certs', 'ca.crt')) : undefined
        } : {} // Webpack will generate self-signed cert if Office certs missing
      },
      headers: {
        'Access-Control-Allow-Origin': '*'
      },
      allowedHosts: 'all'
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
