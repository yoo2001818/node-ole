const ole = require('../src');

console.log(ole);
console.log(ole.create);
// ole.create('InternetExplorer.Application')
ole.create('{8DEC7B3B-3332-4B59-AF2B-DDEBF6419DD7}')
.then(v => {
    console.log(v);
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