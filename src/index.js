const nativeModule = require('../build/Release/node_ole');
const environment = new nativeModule.Environment();

process.on('beforeExit', function () {
    console.log('cleanup');
    nativeModule.cleanup();
});

module.exports = environment;