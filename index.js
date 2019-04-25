if (typeof require !== 'undefined') XLSX = require('xlsx');

var fs = require('fs'),
  path = require('path'),
  filePath = path.join(__dirname, 'files/speakers.xls'),
  fileNames = [];
jsonFile = {},
  text = "";

// make Promise version of fs.readdir()
fs.readdirAsync = function(dirname) {
  return new Promise(function(resolve, reject) {
    fs.readdir(dirname, function(err, filenames) {
      if (err)
        reject(err);
      else
        resolve(filenames);
    });
  });
};

// make Promise version of fs.readFile()
fs.readFileAsync = function(filename, enc) {
  return new Promise(function(resolve, reject) {
    fs.readFile(filename, enc, function(err, data) {
      if (err)
        reject(err);
      else
        resolve(data);
    });
  });
};

var readXLSX = function(filename) {
  return new Promise(function(resolve, reject) {

    var workbook = XLSX.readFile(filename)
    var sheet_name_list = workbook.SheetNames;
    var data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
    if (data) {
      resolve(data)
    } else {
      reject("xlsx error");
    }
  }).catch(err => console.log(err));
}

// utility function, return Promise
function getFile(filename) {
  filePath = path.join(__dirname, 'files/' + filename);
  if (filename.split('.')[1] == 'xls') {
    return readXLSX(filePath);
  }
  return fs.readFileAsync(filePath, 'utf8');
}

// example of using promised version of getFile
// getFile('./fish1.json', 'utf8').then(function (data){
// console.log(data);
// });


// a function specific to my project to filter out the files I need to read and process, you can pretty much ignore or write your own filter function.
function isDataFile(filename) {
  return (filename.split('.')[1] == 'txt' ||
    filename.split('.')[1] == 'xls')
}

// start a blank fishes.json file
//fs.writeFile('.files/fishes.json', '', function(){console.log('done')});

// read all json files in the directory, filter out those needed to process, and using Promise.all to time when all async readFiles has completed.
fs.readdirAsync('./files/').then(function(filenames) {
  filenames = filenames.filter(isDataFile);
  fileNames = filenames;
  return Promise.all(filenames.map(getFile));
}).then(function(files) {
  //console.log(files);
  //console.log(fileNames);

  for (var i = 0; i < fileNames.length; i++) {
    var name = fileNames[i];
    if (name.split('.')[1] == 'xls') {

      var templateData = "";

      var bioData = "";

      for (var j = 0; j < files[i].length; j++) {

        var file = files[i][j];
        var firstname = file['First Name'];
        var lastname = file['Last Name'];
        var title = file.Title;
        var bio = file.Bio;
        var org = file['Organization Name'];
        var linkedin = file['Linkedin URL'];


        console.log(file);

        var template = "<div class='col-md-3'> \
			      <div class='speaker'> \
			        <div class='speakerImage'><a href='#' data-featherlight='#" + firstname + "-" + lastname + "-bio'><img src='https://www2.arccorp.com/globalassets/home2/2019/speakers/" + firstname.toLowerCase() + "-" + lastname.toLowerCase() + ".jpg' alt='" + firstname + "-" + lastname + "'></a></div> \
			        <div class='speakerName'>" + firstname + " " + lastname + "</div> \
			        <div class='speakerTitle'>" + title + "</div> \
			        <div class='speakerCompany'>" + org + "</div> \
							<a class='ctaBtn' href='#' data-featherlight='#" + firstname + "-" + lastname + "-bio'>View Profile</a> \
			      </div> \
			    </div>";

        var bioTemplate = "<div id='" + firstname + "-" + lastname + "-bio' class='speakerProfile'> \
				  <div class='speaker'> \
					<div class='row'>\
					<div class='col-md-4' style='text-align:center'>\
						<div class='speakerImage'><img src='https://www2.arccorp.com/globalassets/home2/2019/speakers/" + firstname.toLowerCase() + "-" + lastname.toLowerCase() + ".jpg' alt='" + firstname + "-" + lastname + "'></div> \
						<div class='speakerName'>" + firstname + " " + lastname + "</div> \
						<div class='speakerTitle'>" + title + "</div> \
						<div class='speakerCompany'>" + org + "</div> \
				    <div class='speakerSocial'> \
				      <div class='speakerLinkedin'> \
				        <a target='_blank' href='" + linkedin + "'><img src='https://www.arctravelconnect.com/globalassets/home2/2019/svg/LINKEDIN-ICON.svg' alt=''></a> \
				      </div> \
				    </div> \
					</div> \
			    <div class='col-md-8'><div class='speakerBio'>" + bio.replace(/(?:\r\n|\r|\n)/g, "<br/>") + "</div></div> \
				  </div> \
					</div> \
				</div>";

        //console.log(template);

        templateData += "\n" + template;

        bioData += "\n" + bioTemplate;

      }

      templateData = "<div class='speakersSection'><div class='row no-gutters'>" + templateData + "</div></div>";

      var totalData = templateData + "\n" + bioData;

      fs.writeFile('./' + name.split('.')[0] + '.html', totalData, function() {
        console.log('done')
      });
    }
  }
});
