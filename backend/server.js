const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const app = express();
const nodemailer = require("nodemailer");

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(cors());

const transport = {
  host: "smtp.gmail.com",
  auth: {
    user: "", //Username
    pass: "", //Password
  },
};

const transporter = nodemailer.createTransport(transport);

transporter.verify((error, success) => {
  if (error) {
    console.log(error);
  } else {
    console.log("All works fine, congratz!");
  }
});


var data = [];

var XLSX = require('xlsx');
var workbook = XLSX.readFile('./test.xlsx');
var sheet_name_list = workbook.SheetNames;
sheet_name_list.forEach(function(y) {
    var worksheet = workbook.Sheets[y];
    var headers = {};
    for(z in worksheet) {1
        if(z[0] === '!') continue;
        var tt = 0;
        for (var i = 0; i < z.length; i++) {
            if (!isNaN(z[i])) {
                tt = i;
                break;
            }
        };
        var col = z.substring(0,tt);
        var row = parseInt(z.substring(tt));
        var value = worksheet[z].v;

        if(row == 1 && value) {
            headers[col] = value;
            continue;
        }

        if(!data[row]) data[row]={};
        data[row][headers[col]] = value;
    }
    data.shift();
    data.shift();
});









app.post("/", (req, res) => {
  data.forEach((e)=> {
    // console.log(e);
    const name = "Krishna Khatri";
    const email = e.email;
    const company = e.company;
    const salutation = e.salutation;
    if (email.length < 5 || company.length < 3) {
      console.log("Len error at", email, " ", length)
      return;
    } else {
      var mail = {
        from: name,
        to: email,
        subject: `Application with ${company}`,
        attachments: [
          {
            filename: "Resume_Krishna_SDE.pdf", //File name
            path: "./Resume_Krishna_SDE.pdf", //File Path
          },
        ],
        html: `<p><span style="color: rgb(34, 34, 34); font-family: arial, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; display: inline !important; float: none;">Hello ${salutation},</span></p>
          <div style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial;">Hope you are doing well.</div>
          <div style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial;"><br>I am Krishna Khatri, final year student of IIT(ISM) Dhanbad, India, graduating in May 2021, doing my majors in Computer Science and Engineering. I&apos;m writing to you regarding the full-time <strong>Software Developer&nbsp;</strong>role at&nbsp;${company}.</div>
          <div style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial;"><br>${salutation}, I have spent my last summer working with one of the hottest start-ups in Portugal, Smartex.Ai as a <strong>frontend web developer intern</strong>. There I was responsible for developing the factory dashboard <span style="letter-spacing: 0px;">for the clients of the company and helped them to present the visuals of factory analysis.</span></div>
          <div style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial;"><span style="letter-spacing: 0px;"><br></span><span style="letter-spacing: 0px;">Before that, I have worked on developing an online community where people can ask and answer business-related queries. I took this whole website was taken from zero to one in two-month single-handedly, which includes both <strong>frontend and backend development</strong> along with <strong>hosting on AWS</strong>.</span></div>
          <div style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial;"><br>In the winter of 2017, I have done a <strong>data analysis internship</strong> at Saddacampus, where I analysed around 25,000 food orders, and the resulting outcomes were used for the business development purpose.</div>
          <div style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial;"><br>Along with this, I have served as a <strong>Tech. Head</strong> for my college entrepreneurship cell, where I developed the official website of the cell (<strong><a data-saferedirecturl="https://www.google.com/url?q=http://www.ecelliitism.org/&source=gmail&ust=1603805274437000&usg=AFQjCNG6rDiiWfVX4Vlfrs4AzepGktznQA" href="http://www.ecelliitism.org/" rel="noreferrer" style="color: rgb(17, 85, 204);" target="_blank">www.ecelliitism.org</a></strong>). It is designed on <strong>Adobe XD</strong> and developed with <strong>HTML, CSS, Javascript and Bootstrap</strong>. Along with this, I have developed the internship portal for the cell, where students from the first year and second year can come and grab an internship (<strong><a data-saferedirecturl="https://www.google.com/url?q=http://www.ecelliitism.org/internship&source=gmail&ust=1603805274437000&usg=AFQjCNGYz2Zjeo_gbab2PtEReTSuUBa4jw" href="http://www.ecelliitism.org/internship" rel="noreferrer" style="color: rgb(17, 85, 204);" target="_blank">www.ecelliitism.org/<wbr>internship</a></strong>).</div>
          <div style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial;"><br>Currently, I am mastering my skills in sports programming(data structures and algorithms), where I have solved more than <strong>800 problems</strong> on the various platforms, and along with this, I have participated in multiple online competitions. I have stood <strong>46th among 4000 participants</strong> in February Lunch Time hosted by CodeChef. Along with that, I have grabbed the <strong>99th position out of 3500 participants</strong> in September Mega Cook-Off challenge. Also, my team secured<strong>&nbsp;31st rank among 782 teams</strong> in Inter-University programming contest hosted by LNMIIT Jaipur, and many more such achievements are in my account. Right now <span style="letter-spacing: 0px;">I am <strong>1914 rated</strong> on Codechef, <strong>1915 rated</strong> on Codeforces and <strong>888 rated</strong> on AtCoder.</span></div>
          <div style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial;"><span style="letter-spacing: 0px;"><br></span><span style="letter-spacing: 0px;">I believe my experience and skill will put good use to your organisation, and in case you feel the same, then please consider me. I am attaching my resume with this.</span></div>
          <div style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial;"><span style="letter-spacing: 0px;"><br></span><span style="letter-spacing: 0px;">Always happy to hear from you.</span></div>
          <div style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial;"><span style="letter-spacing: 0px;"><br></span>With Regards,</div>
          <div style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial;">Krishna Khatri</div>
          <div style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial;">B.Tech. CSE,</div>
          <div style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial;">IIT(ISM) Dhanbad</div>
          <div style="color: rgb(34, 34, 34); font-family: Arial, Helvetica, sans-serif; font-size: small; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial;">+91-9631610549</div>`,
      };
  
      transporter.sendMail(mail, (err, data) => {
        if (err) {
          console.log("Error at", email, " ")
        } else {
          console.log("Mailed at", email);
        }
      });
    }
  })
  res.sendStatus(200)
});

app.listen(process.env.PORT || 3332, function () {
  console.log(`Server started on port ${process.env.PORT || 3332}.`);
});
