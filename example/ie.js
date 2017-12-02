const ole = require('../src');

(async function() {
  let ie = await ole('InternetExplorer.Application');
  await ie.Navigate('http://google.com');
  ie.on('NaviageComplete', (url) => {
    console.log(url);
  });
})();
