const ole = require('../build/Release/node_ole');
const environment = new ole.Environment();

console.log(environment);
console.log(environment.create);
console.log(environment.create('InternetExplorer.Application'));

(async function() {
  let ie = await environment.create('InternetExplorer.Application');
  await ie.Navigate('http://google.com');
  ie.on('NaviageComplete', (url) => {
    console.log(url);
  });
})();
