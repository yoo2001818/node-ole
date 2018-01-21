const environment = require('../src');
environment.create('InternetExplorer.Application')
// ole.create('{8DEC7B3B-3332-4B59-AF2B-DDEBF6419DD7}')
  .then(v => {
    console.log(v);
    for (let key in v) {
      v[key].types.forEach(type => {
        // Format the type information.
        let result = `[${type.type}] ${type.name}`;
        let argInfo = type.args.map(v => {
          return `${v.name}${v.optional ? '?' : ''}: ${v.type}`;
        });
        result += '(' + argInfo.join(', ') + ')';
        result += ': ' + type.returns.type;
        console.log(result);
      });
    }
    environment.close();
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
