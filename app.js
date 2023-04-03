const express = require('express');
const cors = require('cors');
const fs = require('fs');
const request = require('request');

require('dotenv').config();

const app = express();
const port = 5000;

app.use(express.json({ limit: '2gb' }));
app.use("/tmp", express.static("tmp"));

const corsOpts = {
  origin: process.env.ORIGIN_DOMAIN,
  credentials: true,
  methods: [
    'GET',
    'POST'
  ],
  allowedHeaders: [
    "Origin",
    "X-Requested-With",
    'Content-Type',
    "Accept"
  ]
};
app.use(cors(corsOpts));

app.listen(port, () => {
  console.log(`Listening on port ${port}`)
});

// /-/-/- START SharePoint integration -/-/-/-/
const { SPPull } = require('sppull');
const spsave = require('spsave').spsave;

const SPSiteUrl = process.env.SHAREPOINT_URL;
const SPCredentials = {
  username: process.env.SHAREPOINT_USER,
  password: process.env.SHAREPOINT_PASSWORD
};

app.post('/ask-to-upload', (req, res) => {
  let type = "docs";
      
  const coreOptions = {
    siteUrl: SPSiteUrl,
    notification: true,
    checkin: true,
    checkinType: 1,
    filesMetaData: [{
      fileName: req.body.file_name,
      metadata: {
        __metadata:{'type':'SP.ListItem'},
        "Title": req.body.file_title
      }
    }]
  };

  // get file content from temporary link
  const filePath = encodeURIComponent(req.body.file_name);
  const host = process.env.ORIGIN_DOMAIN;
  const reqPath = host + process.env.ORIGIN_API + filePath;
  
  // temporary store for binary file (create tmp dir if not exists)
  const tmpDir = __dirname + "/tmp/";
  const tmpPath =  tmpDir + req.body.file_name;
  if (!fs.existsSync(tmpDir)){
    fs.mkdirSync(tmpDir, { recursive: true });
  }

  const writeStream = fs.createWriteStream(tmpPath, (err) => {
    if (err) console.log(err);
  });
  
  const stream = request.get(reqPath, (error, response) => {
    if (error || response.statusCode != 200) {
      console.log(error);
    }
  }).pipe(writeStream);

  stream.on('finish', () => {
    // sharepoint upload options
    const fileOptions = {
      folder: (
        type + "/" + 
        req.body.year + "/" + 
        req.body.month
      ),
      fileName: req.body.file_name,
      fileContent: fs.readFileSync(tmpPath, (err) => {
        if (err) console.log(err);
      })
    };
    
    // upload file to sharepoint
    spsave(coreOptions, SPCredentials, fileOptions)
    .then(function(){
      fs.unlink(tmpPath, (err) => {
        if (err) console.log(err);
      });
      console.log("upload success");
      res.send("1");
    })
    .catch(function(err){
      fs.unlink(tmpPath, (err) => {
        if (err) console.log(err);
      });
      console.log(err);
    });
  });
});

app.post('/download', (req, res) => {
  let type = "docs";

  const context = {
    siteUrl: SPSiteUrl,
    ...SPCredentials
  };
  
  const options = {
    spRootFolder: type + "/" + req.body.folder,
    dlRootFolder: "./tmp",
    strictObjects: [
      req.body.file
    ]
  };

  SPPull.download(context, options)
    .then((downloadResults) => {
      console.log("Files are downloaded");
      console.log("For more, please check the results", JSON.stringify(downloadResults));
      
      const path = __dirname + "/tmp/" + req.body.file;
      res.sendFile(path, (err) => {
        if (err) {
          console.log(err);
          res.send("0");
        }
        fs.unlinkSync(path);
      });
    })
    .catch((err) => {
      console.log("Core error has happened", err);
      res.send("0");
    });
});

// /-/-/- END SharePoint integration -/-/-/-/