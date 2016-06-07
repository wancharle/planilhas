var mainBowerFiles = require('main-bower-files');
var concat = require('concat');
var js_files = mainBowerFiles("**/*.js");

concat(js_files,"tests/bowerdeps.js",function (error) { 
    if (error) 
        console.error(error);
    else 
        console.log("Concatenando dependencias do bower:",js_files, "\n no arquivo tests/bowerdeps.js\n");
});



