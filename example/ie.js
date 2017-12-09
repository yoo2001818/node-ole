const nativeModule = require('../build/Release/node_ole');
const environment = new nativeModule.Environment();
environment.create('InternetExplorer.Application')
// ole.create('{8DEC7B3B-3332-4B59-AF2B-DDEBF6419DD7}')
.then(v => {
    console.log(v);
    for (let key in v) {
        v[key].types.forEach(type => {
            console.log(type);
        });
    }
    nativeModule.cleanup();
}, e => console.error(e));

/*
(async function() {
  let ie = await ole.create('InternetExplorer.Application');
  await ie.Navigate('http://google.com');
  ie.on('NaviageComplete', (url) => {
    console.log(url);
  });
})();
*/