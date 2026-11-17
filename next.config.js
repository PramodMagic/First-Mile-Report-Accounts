module.exports = {
  poweredByHeader: false,
  webpack: (config) => { config.resolve.alias.canvas = false; return config; }, // pdf.js optional node dependency
};
