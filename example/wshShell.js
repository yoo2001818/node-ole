const environment = require('../src');

(async function() {
  let wshShell = await environment.create('Wscript.Shell');
  for (let key in wshShell) {
    wshShell[key].types.forEach(type => {
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
  console.log(wshShell.Popup('Hello, world on node-ole!', true, 'Hello!', 16));
  // environment.close();
})()
  .catch(v => console.error(v));
