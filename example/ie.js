const environment = require('../src');
(async function() {
  let ie = await environment.create('InternetExplorer.Application');
  // ole.create('{8DEC7B3B-3332-4B59-AF2B-DDEBF6419DD7}')
  for (let key in ie) {
    ie[key].types.forEach(type => {
      // Format the type information.
      let result = `[${type.type}] ${type.name}`;
      let argInfo = type.args.map(v => {
        return `${v.out ? '[out] ' : ''}${v.pointer ? '*' : ''}${v.name}${v.optional ? '?' : ''}: ${v.type}`;
      });
      result += '(' + argInfo.join(', ') + ')';
      result += ': ' + type.returns.type;
      console.log(result);
    });
  }
  await ie.Navigate2("http://google.com/");
  await ie.Visible(true);
  await ie.StatusBar(true);
  setTimeout(() => ie.StatusText('Go to sleep'), 1000);
  console.log(await ie.Path());
  console.log(await ie.ReadyState());
  console.log(await ie.Application());
  setTimeout(() => ie.Quit(), 10000);
})()
  .catch(v => console.error(v));
