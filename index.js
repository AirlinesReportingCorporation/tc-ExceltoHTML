if (typeof require !== 'undefined') XLSX = require('xlsx');

const order = ["Kirby", "Simi", "Webb", "Katz", "Harvey", "Best", "Ali", "Lamping", "D'Astolfo" , "Moscovici", "Radcliffe", "Morin", "Hegeman", "Barth", "Tonnessen", "Heywood", "Swanston", "Armstrong", "Share", "Premo", "Reishus", "Oliver", "Barber", "Gillespie", "Younger", "Nass", "Hattingh", "Coleman"
 ];

var moment = require('moment');

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

function findWithAttr(array, attr, value) {
  for (var i = 0; i < array.length; i += 1) {
    console.log(attr);
    console.log(value);
    if (array[i][attr] === value) {
      console.log(array[i][attr]);
      return i;
    }
  }
  return -1;
}

function compareColumn( a, b ) {
  if ( a['Start Time'] < b['Start Time'] ){
    return -1;
  }
  if ( a['Start Time'] > b['Start Time'] ){
    return 1;
  }
  return 0;
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
    var filename = name.split('.')[0];
    var extension = name.split('.')[1];

    if (filename == 'speakers' && extension == 'xls') {

      var templateData = "";

      var bioData = "";

      for (var j = 0; j < order.length; j++) {

        console.log(order.length);

        var x = findWithAttr(files[i], 'Last Name', order[j]);
        console.log(x);

        var file = files[i][x];
        var firstname = file['First Name'];
        var lastname = file['Last Name'];
        var title = file.Title;
        var bio = file.Bio;
        var org = file['Organization Name'] ? file['Organization Name'] : "&nbsp;";
        var linkedin = file['Linkedin URL'];
        var website1 = file['Website'] ? file['Website'] : "&nbsp;";
        var website2 = file['Website 2'] ? file['Website 2'] : "&nbsp;";


        var template = "<div class='col-md-3'> \
			      <div id='" + firstname.toLowerCase() + "-" + lastname.replace("'", "").toLowerCase() + "' class='speaker'> \
			        <div class='speakerImage'><a href='#' data-featherlight='#" + firstname + "-" + lastname.replace("'", "") + "-bio'><img src='https://www2.arccorp.com/globalassets/home2/2019/speakers/" + firstname.toLowerCase() + "-" + lastname.replace("'", "").toLowerCase() + ".jpg' alt='" + firstname + "-" + lastname.replace("'", "") + "'></a></div> \
			        <div class='speakerName'>" + firstname + " " + lastname.replace("'", "&apos;") + "</div> \
			        <div class='speakerTitle'>" + title + "</div> \
			        <div class='speakerCompany'>" + org + "</div> \
							<a class='ctaBtn' href='#' data-featherlight='#" + firstname + "-" + lastname.replace("'", "") + "-bio'>View Profile</a> \
			      </div> \
			    </div>";

        var bioTemplate = "<div id='" + firstname + "-" + lastname.replace("'", "") + "-bio' class='speakerProfile'> \
				  <div class='speaker'> \
					<div class='row'>\
					<div class='col-md-4' style='text-align:center'>\
						<div class='speakerImage'><img src='https://www2.arccorp.com/globalassets/home2/2019/speakers/" + firstname.toLowerCase() + "-" + lastname.replace("'", "").toLowerCase() + ".jpg' alt='" + firstname + "-" + lastname.replace("'", "") + "'></div> \
						<div class='speakerName'>" + firstname + " " + lastname.replace("'", "&apos;") + "</div> \
						<div class='speakerTitle'>" + title + "</div> \
						<div class='speakerCompany'>" + org + "</div> \
            <div class='website1'><a target='_blank' href='" + website1 + "'>" + website1 + "</a></div> \
            <div class='website1'><a target='_blank' href='" + website2 + "'>" + website2 + "</a></div> \
				    <div class='speakerSocial " + (linkedin ? " " : "hideSocial") + "'> \
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
        console.log('./' + name.split('.')[0] + '.html written to root directory');
        console.log('================================');
      });
    }

    if (filename == 'agenda' && extension == 'xls') {

      //files[0].sort(compareColumn());
      files[i].sort(compareColumn);

      var templateData = "";

      for (var j = 0; j < files[i].length; j++) {

        var file = files[i][j];

        var uniqueid = file['Unique ID'];

        var sessionName = file['Name'];
        var sessionDescription = file['Description'];

        var tags = file['Tags (comma-separated)'];

        var startTime = file['Start Time'];
        var startDate = startTime.split(" ")[0];
        startTime = startTime.split(" ")[1].replace(":", "");
        startTime = moment(startTime, "hmm").format("h:mm");

        var endTime = file['End Time'];
        var endDate = endTime.split(" ")[0];
        endTime = endTime.split(" ")[1].replace(":", "");
        endTime = moment(endTime, "hmm").format("h:mm a");

        var speaker1first = file['Speaker 1 First Name'];
        var speaker1last = file['Speaker 1 Last Name'];
        var speaker1role = file['Speaker 1 Role'];
        var speaker1Title = file['Speaker 1 Title'];
        var speaker1Bio = file['Speaker 1 Bio'];
        var speaker1Org = file['Speaker 1 Organization Name'];

        var speaker2first = file['Speaker 2 First Name'];
        var speaker2last = file['Speaker 2 Last Name'];
        var speaker2role = file['Speaker 2 Role'];
        var speaker2Title = file['Speaker 2 Title'];
        var speaker2Bio = file['Speaker 2 Bio'];
        var speaker2Org = file['Speaker 2 Organization Name'];

        var speaker3first = file['Speaker 3 First Name'];
        var speaker3last = file['Speaker 3 Last Name'];
        var speaker3role = file['Speaker 3 Role'];
        var speaker3Title = file['Speaker 3 Title'];
        var speaker3Bio = file['Speaker 3 Bio'];
        var speaker3Org = file['Speaker 3 Organization Name'];

        var speaker4first = file['Speaker 4 First Name'];
        var speaker4last = file['Speaker 4 Last Name'];
        var speaker4role = file['Speaker 4 Role'];
        var speaker4Title = file['Speaker 4 Title'];
        var speaker4Bio = file['Speaker 4 Bio'];
        var speaker4Org = file['Speaker 4 Organization Name'];

        var speaker5first = file['Speaker 5 First Name'];
        var speaker5last = file['Speaker 5 Last Name'];
        var speaker5role = file['Speaker 5 Role'];
        var speaker5Title = file['Speaker 5 Title'];
        var speaker5Bio = file['Speaker 5 Bio'];
        var speaker5Org = file['Speaker 5 Organization Name'];

        var location = file['Location Name'];

        var template = '<div class="row schedule-item" id="' + uniqueid + '"> \
            <div class="col-md-3"> \
                <div class="schedule-time">' + startTime + ' - ' + endTime + '</div> \
            </div> \
             <div class="col-md-3"> \
                <div class="schedule-type">' + (tags ? tags : " ") + '</div> \
            </div> \
            <div class="col-md-6"> \
                <div class="schedule-name">' + sessionName + '</div> \
                <p class="schedule-participants">' + (sessionDescription ? sessionDescription : " ") + '</p>  \
                <p class="schedule-participants"> ' + (speaker1first ? "<strong>Speakers:</strong><br/>" +

                  (speaker1first ? "<a href='/tc2019/speakers#" + speaker1first.toLowerCase() + "-" + speaker1last.replace("'", "").toLowerCase() + "'><strong>" + speaker1first + " " + speaker1last + "</strong></a>, " + speaker1Title + ", " + speaker1Org + "<br/>" : " ") +

                  (speaker2first ? "<a href='/tc2019/speakers#" + speaker2first.toLowerCase() + "-" + speaker2last.replace("'", "").toLowerCase() + "'><strong>" + speaker2first + " " + speaker2last + "</strong></a>, " + speaker2Title + ", " + speaker2Org + "<br/>" : " ") +

                  (speaker3first ? "<a href='/tc2019/speakers#" + speaker3first.toLowerCase() + "-" + speaker3last.replace("'", "").toLowerCase() + "'><strong>" + speaker3first + " " + speaker3last + "</strong></a>, " + speaker3Title + ", " + speaker3Org + "<br/>" : " ") +

                  (speaker4first ? "<a href='/tc2019/speakers#" + speaker4first.toLowerCase() + "-" + speaker4last.replace("'", "").toLowerCase() + "'><strong>" + speaker4first + " " + speaker4last + "</strong></a>, " + speaker4Title + ", " + speaker4Org + "<br/>" : " ") +

                  (speaker5first ? "<a href='/tc2019/speakers#" + speaker5first.toLowerCase() + "-" + speaker5last.replace("'", "").toLowerCase() + "'><strong>" + speaker5first + " " + speaker5last + "</strong></a>, " + speaker5Title + ", " + speaker5Org + "<br/>" : " ") : " ") +
                 '</p>  \
                <p class="schedule-participants"> ' + (location ? "<strong>Location:</strong><br/>" + location: "") + '</p> \
            </div>  \
        </div>';

        templateData += "\n" + template;

        var preconf = "";
        var day1 = "";
        var day2 = "";
      }

      fs.writeFile('./' + filename + '.html', templateData, function() {
        console.log('./' + filename + '.html written to root directory');
        console.log('================================');
      });

    }

  }
});
