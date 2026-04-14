module.exports = {
  apps: [
    {
      name: 'mogal-complain-sys',
      script: 'server.js',
      node_args: '',
      env: {
        NODE_ENV: 'production',
        PORT: 3050,
        DATA_DIR: '/var/www/mogal-complain-sys/data',
      },
      watch: false,
      max_memory_restart: '300M',
      restart_delay: 3000,
      log_file: '/var/log/mogal/combined.log',
      out_file: '/var/log/mogal/out.log',
      error_file: '/var/log/mogal/error.log',
      log_date_format: 'YYYY-MM-DD HH:mm:ss',
    },
  ],
};
