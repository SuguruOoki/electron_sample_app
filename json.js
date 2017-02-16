function download_json(){

var fs = require('fs');

	var data = {
    	hoge: 100,
    	foo: 'a',
    	bar: true,
		};
	fs.writeFile('hoge.json', JSON.stringify(data, null, '    '));
}