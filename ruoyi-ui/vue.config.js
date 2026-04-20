const TEXT_TYPES = /(text|javascript|json|xml|svg|css)/i

module.exports = {
  lintOnSave: false,
  devServer: {
    port: 8080,
    host: '0.0.0.0',
    headers: {
      'Access-Control-Allow-Origin': '*'
    },
    before(app) {
      app.use((req, res, next) => {
        const type = res.getHeader('Content-Type')
        if (type && TEXT_TYPES.test(type) && !/charset/i.test(type)) {
          res.setHeader('Content-Type', `${type}; charset=utf-8`)
        }
        next()
      })
    }
  },
  chainWebpack: config => {
    config.plugin('html').tap(args => {
      args[0].title = 'RAG 智能助手'
      args[0].meta = {
        charset: { charset: 'UTF-8' },
        viewport: 'width=device-width,initial-scale=1.0',
        description: '本地 RAG 智能助手'
      }
      return args
    })
  }
}

