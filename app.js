const express = require("express");
const bodyParser = require("body-parser");
const ejs = require("ejs");
const mongoose = require("mongoose");
const session = require("express-session");
const passport =  require("passport");
const passportLocalMongoose = require("passport-local-mongoose");
const fileUpload = require('express-fileupload');
const xlsx = require("xlsx");
const convert = require("xml-js");
const XMLWriter = require('xml-writer');
const fs = require("fs");
const builder = require('xmlbuilder');

const app = express();

app.use(express.static("public"));
app.set("view engine", "ejs");
app.use(bodyParser.urlencoded({extended: true}));
app.use(fileUpload());
app.use(session({
  secret: "no one finds war; war finds them",
  resave: false,
  saveUninitialized: false
}));
app.use(passport.initialize());
app.use(passport.session());

mongoose.connect("mongodb+srv://admin-alsan:RO1Yzyn2PO6pvjMU@clusterxml.kn4ro.mongodb.net/xmluserDB", {useNewUrlParser: true, useUnifiedTopology: true});
mongoose.set("useCreateIndex", true);

const userSchema = new mongoose.Schema({
  email: String,
  password: String
});

userSchema.plugin(passportLocalMongoose);

const User = new mongoose.model("User", userSchema);

passport.use(User.createStrategy());
passport.serializeUser(User.serializeUser());
passport.deserializeUser(User.deserializeUser());

app.get("/", function(req, res){
  res.render("home");
});

app.get("/login", function(req, res){
  res.render("login");
});

app.get("/register", function(req, res){
  res.render("register");
});

app.get("/xmlView", function(req, res){
  res.render("xmlView");
});

app.get("/converter", function(req, res){
  if (req.isAuthenticated()) {
    res.render("converter");
  } else {
    res.redirect("/login");
  }
});

app.get("/logout", function(req, res){
  req.logout();
  res.redirect("/");
});

app.post("/register", function(req, res){
  User.register({username: req.body.username}, req.body.password, function(err, user){
    if(err){
      console.log(err);
      res.redirect("/register");
    } else {
      passport.authenticate("local")(req, res, function(){
        res.redirect("/converter");
      });
    }
  });
});

app.post("/login", function(req, res){
  const user = new User({
    username: req.body.username,
    password: req.body.password
  });
  req.login(user, function(err){
    if (err) {
      console.log(err);
    } else {
      passport.authenticate("local")(req, res, function(){
        res.redirect("/converter");
      });
    }
  });
});

/************************ WORKBOOK HANDLING ************************/
app.post("/converter", function(req, res){
  if (req.files) {
    console.log(req.files);
    var file = req.files.arrendamientoArchivo;
    var fileName = file.name;
    console.log(fileName);
    file.mv('./uploads/' + fileName, function(err){
      if (err) {
        console.log(err);
      } else {
        console.log("file uploaded succesfully!");
        var wb = xlsx.readFile("./uploads/" + fileName, {cellDates:true});
        var first_sheetName = wb.SheetNames[1];
        var worksheet = wb.Sheets[first_sheetName];

        let a5="A5"; let a5Adress=worksheet[a5]; let a5Value=(a5Adress?a5Adress.v:undefined);
        let b5 = "B5"; let b5Adress = worksheet[b5]; let b5Value = (b5Adress ? b5Adress.v : undefined);
        let c5 = "C5"; let c5Adress = worksheet[c5]; let c5Value = (c5Adress ? c5Adress.v : undefined);
        let d5 = "D5"; let d5Adress = worksheet[d5]; let d5Value = (d5Adress ? d5Adress.v : undefined);
        let e5 = "E5"; let e5Adress = worksheet[e5]; let e5Value = (e5Adress ? e5Adress.v : undefined);
        let f5 = "F5"; let f5Adress = worksheet[f5]; let f5Value = (f5Adress ? f5Adress.v : undefined);
        let g5 = "G5"; let g5Adress = worksheet[g5]; let g5Value = (g5Adress ? g5Adress.v : undefined);
        let h5 = "H5"; let h5Adress = worksheet[h5]; let h5Value = (h5Adress ? h5Adress.v : undefined);
        let i5 = "I5"; let i5Adress = worksheet[i5]; let i5Value = (i5Adress ? i5Adress.v : undefined);
        let j5 = "J5"; let j5Adress = worksheet[j5]; let j5Value = (j5Adress ? j5Adress.v : undefined);
        let k5 = "K5"; let k5Adress = worksheet[k5]; let k5Value = (k5Adress ? k5Adress.v : undefined);
        let l5 = "L5"; let l5Adress = worksheet[l5]; let l5Value = (l5Adress ? l5Adress.v : undefined);
        let m5 = "M5"; let m5Adress = worksheet[m5]; let m5Value = (m5Adress ? m5Adress.v : undefined);
        let n5 = "N5"; let n5Adress = worksheet[n5]; let n5Value = (n5Adress ? n5Adress.v : undefined);
        let o5 = "O5"; let o5Adress = worksheet[o5]; let o5Value = (o5Adress ? o5Adress.v : undefined);
        let p5 = "P5"; let p5Adress = worksheet[p5]; let p5Value = (p5Adress ? p5Adress.v : undefined);
        let q5 = "Q5"; let q5Adress = worksheet[q5]; let q5Value = (q5Adress ? q5Adress.v : undefined);
        let r5 = "R5"; let r5Adress = worksheet[r5]; let r5Value = (r5Adress ? r5Adress.v : undefined);
        let s5 = "S5"; let s5Adress = worksheet[s5]; let s5Value = (s5Adress ? s5Adress.v : undefined);
        let t5 = "T5"; let t5Adress = worksheet[t5]; let t5Value = (t5Adress ? t5Adress.v : undefined);
        let u5 = "U5"; let u5Adress = worksheet[u5]; let u5Value = (u5Adress ? u5Adress.v : undefined);
        let v5 = "V5"; let v5Adress = worksheet[v5]; let v5Value = (v5Adress ? v5Adress.v : undefined);
        let w5 = "W5"; let w5Adress = worksheet[w5]; let w5Value = (w5Adress ? w5Adress.v : undefined);
        let x5 = "X5"; let x5Adress = worksheet[x5]; let x5Value = (x5Adress ? x5Adress.v : undefined);
        let y5 = "Y5"; let y5Adress = worksheet[y5]; let y5Value = (y5Adress ? y5Adress.v : undefined);
        let z5 = "Z5"; let z5Adress = worksheet[z5]; let z5Value = (z5Adress ? z5Adress.v : undefined);
        let aa5 = "AA5"; let aa5Adress = worksheet[aa5]; let aa5Value = (aa5Adress ? aa5Adress.v : undefined);
        let ab5 = "AB5"; let ab5Adress = worksheet[ab5]; let ab5Value = (ab5Adress ? ab5Adress.v : undefined);
        let ac5 = "AC5"; let ac5Adress = worksheet[ac5]; let ac5Value = (ac5Adress ? ac5Adress.v : undefined);
        let ad5 = "AD5"; let ad5Adress = worksheet[ad5]; let ad5Value = (ad5Adress ? ad5Adress.v : undefined);
        let ae5 = "AE5"; let ae5Adress = worksheet[ae5]; let ae5Value = (ae5Adress ? ae5Adress.v : undefined);
        let af5 = "AF5"; let af5Adress = worksheet[af5]; let af5Value = (af5Adress ? af5Adress.v : undefined);
        let ag5 = "AG5"; let ag5Adress = worksheet[ag5]; let ag5Value = (ag5Adress ? ag5Adress.v : undefined);
        let ah5 = "AH5"; let ah5Adress = worksheet[ah5]; let ah5Value = (ah5Adress ? ah5Adress.v : undefined);
        let ai5 = "AI5"; let ai5Adress = worksheet[ai5]; let ai5Value = (ai5Adress ? ai5Adress.v : undefined);
        let aj5 = "AJ5"; let aj5Adress = worksheet[aj5]; let aj5Value = (aj5Adress ? aj5Adress.v : undefined);
        let ak5 = "AK5"; let ak5Adress = worksheet[ak5]; let ak5Value = (ak5Adress ? ak5Adress.v : undefined);
        let al5 = "AL5"; let al5Adress = worksheet[al5]; let al5Value = (al5Adress ? al5Adress.v : undefined);
        let am5 = "AM5"; let am5Adress = worksheet[am5]; let am5Value = (am5Adress ? am5Adress.v : undefined);
        let an5 = "AN5"; let an5Adress = worksheet[an5]; let an5Value = (an5Adress ? an5Adress.v : undefined);
        let ao5 = "AO5"; let ao5Adress = worksheet[ao5]; let ao5Value = (ao5Adress ? ao5Adress.v : undefined);
        let ap5 = "AP5"; let ap5Adress = worksheet[ap5]; let ap5Value = (ap5Adress ? ap5Adress.v : undefined);
        let aq5 = "AQ5"; let aq5Adress = worksheet[aq5]; let aq5Value = (aq5Adress ? aq5Adress.v : undefined);
        let ar5 = "AR5"; let ar5Adress = worksheet[ar5]; let ar5Value = (ar5Adress ? ar5Adress.v : undefined);
        let as5 = "AS5"; let as5Adress = worksheet[as5]; let as5Value = (as5Adress ? as5Adress.v : undefined);
        let at5 = "AT5"; let at5Adress = worksheet[at5]; let at5Value = (at5Adress ? at5Adress.v : undefined);
        let au5 = "AU5"; let au5Adress = worksheet[au5]; let au5Value = (au5Adress ? au5Adress.v : undefined);
        let av5 = "AV5"; let av5Adress = worksheet[av5]; let av5Value = (av5Adress ? av5Adress.v : undefined);
        let aw5 = "AW5"; let aw5Adress = worksheet[aw5]; let aw5Value = (aw5Adress ? aw5Adress.v : undefined);
        let ax5 = "AX5"; let ax5Adress = worksheet[ax5]; let ax5Value = (ax5Adress ? ax5Adress.v : undefined);
        let ay5 = "AY5"; let ay5Adress = worksheet[ay5]; let ay5Value = (ay5Adress ? ay5Adress.v : undefined);
        let az5 = "AZ5"; let az5Adress = worksheet[az5]; let az5Value = (az5Adress ? az5Adress.v : undefined);
        let ba5 = "BA5"; let ba5Adress = worksheet[ba5]; let ba5Value = (ba5Adress ? ba5Adress.v : undefined);
        let bb5 = "BB5"; let bb5Adress = worksheet[bb5]; let bb5Value = (bb5Adress ? bb5Adress.v : undefined);
        let bc5 = "BC5"; let bc5Adress = worksheet[bc5]; let bc5Value = (bc5Adress ? bc5Adress.v : undefined);
        let bd5 = "BD5"; let bd5Adress = worksheet[bd5]; let bd5Value = (bd5Adress ? bd5Adress.v : undefined);
        let be5 = "BE5"; let be5Adress = worksheet[be5]; let be5Value = (be5Adress ? be5Adress.v : undefined);
        let bf5 = "BF5"; let bf5Adress = worksheet[bf5]; let bf5Value = (bf5Adress ? bf5Adress.v : undefined);
        let bg5 = "BG5"; let bg5Adress = worksheet[bg5]; let bg5Value = (bg5Adress ? bg5Adress.v : undefined);
        let bh5 = "BH5"; let bh5Adress = worksheet[bh5]; let bh5Value = (bh5Adress ? bh5Adress.v : undefined);
        let bi5 = "BI5"; let bi5Adress = worksheet[bi5]; let bi5Value = (bi5Adress ? bi5Adress.v : undefined);
        let bj5 = "BJ5"; let bj5Adress = worksheet[bj5]; let bj5Value = (bj5Adress ? bj5Adress.v : undefined);
        let bk5 = "BK5"; let bk5Adress = worksheet[bk5]; let bk5Value = (bk5Adress ? bk5Adress.v : undefined);
        let bl5 = "BL5"; let bl5Adress = worksheet[bl5]; let bl5Value = (bl5Adress ? bl5Adress.v : undefined);
        let bm5 = "BM5"; let bm5Adress = worksheet[bm5]; let bm5Value = (bm5Adress ? bm5Adress.v : undefined);
        let bn5 = "BN5"; let bn5Adress = worksheet[bn5]; let bn5Value = (bn5Adress ? bn5Adress.v : undefined);
        let bo5 = "BO5"; let bo5Adress = worksheet[bo5]; let bo5Value = (bo5Adress ? bo5Adress.v : undefined);
        let bp5 = "BP5"; let bp5Adress = worksheet[bp5]; let bp5Value = (bp5Adress ? bp5Adress.v : undefined);
        let bq5 = "BQ5"; let bq5Adress = worksheet[bq5]; let bq5Value = (bq5Adress ? bq5Adress.v : undefined);
        let br5 = "BR5"; let br5Adress = worksheet[br5]; let br5Value = (br5Adress ? br5Adress.v : undefined);
        let bs5 = "BS5"; let bs5Adress = worksheet[bs5]; let bs5Value = (bs5Adress ? bs5Adress.v : undefined);
        let bt5 = "BT5"; let bt5Adress = worksheet[bt5]; let bt5Value = (bt5Adress ? bt5Adress.v : undefined);
        let bu5 = "BU5"; let bu5Adress = worksheet[bu5]; let bu5Value = (bu5Adress ? bu5Adress.v : undefined);
        let bv5 = "BV5"; let bv5Adress = worksheet[bv5]; let bv5Value = (bv5Adress ? bv5Adress.v : undefined);
        let bw5 = "BW5"; let bw5Adress = worksheet[bw5]; let bw5Value = (bw5Adress ? bw5Adress.v : undefined);

        let a6 = "A6"; let a6Adress = worksheet[a6]; let a6Value = (a6Adress ? a6Adress.v : undefined);
        let b6 = "B6"; let b6Adress = worksheet[b6]; let b6Value = (b6Adress ? b6Adress.v : undefined);
        let c6 = "C6"; let c6Adress = worksheet[c6]; let c6Value = (c6Adress ? c6Adress.v : undefined);
        let d6 = "D6"; let d6Adress = worksheet[d6]; let d6Value = (d6Adress ? d6Adress.v : undefined);
        let e6 = "E6"; let e6Adress = worksheet[e6]; let e6Value = (e6Adress ? e6Adress.v : undefined);
        let f6 = "F6"; let f6Adress = worksheet[f6]; let f6Value = (f6Adress ? f6Adress.v : undefined);
        let g6 = "G6"; let g6Adress = worksheet[g6]; let g6Value = (g6Adress ? g6Adress.v : undefined);
        let h6 = "H6"; let h6Adress = worksheet[h6]; let h6Value = (h6Adress ? h6Adress.v : undefined);
        let i6 = "I6"; let i6Adress = worksheet[i6]; let i6Value = (i6Adress ? i6Adress.v : undefined);
        let j6 = "J6"; let j6Adress = worksheet[j6]; let j6Value = (j6Adress ? j6Adress.v : undefined);
        let k6 = "K6"; let k6Adress = worksheet[k6]; let k6Value = (k6Adress ? k6Adress.v : undefined);
        let l6 = "L6"; let l6Adress = worksheet[l6]; let l6Value = (l6Adress ? l6Adress.v : undefined);
        let m6 = "M6"; let m6Adress = worksheet[m6]; let m6Value = (m6Adress ? m6Adress.v : undefined);
        let n6 = "N6"; let n6Adress = worksheet[n6]; let n6Value = (n6Adress ? n6Adress.v : undefined);
        let o6 = "O6"; let o6Adress = worksheet[o6]; let o6Value = (o6Adress ? o6Adress.v : undefined);
        let p6 = "P6"; let p6Adress = worksheet[p6]; let p6Value = (p6Adress ? p6Adress.v : undefined);
        let q6 = "Q6"; let q6Adress = worksheet[q6]; let q6Value = (q6Adress ? q6Adress.v : undefined);
        let r6 = "R6"; let r6Adress = worksheet[r6]; let r6Value = (r6Adress ? r6Adress.v : undefined);
        let s6 = "S6"; let s6Adress = worksheet[s6]; let s6Value = (s6Adress ? s6Adress.v : undefined);
        let t6 = "T6"; let t6Adress = worksheet[t6]; let t6Value = (t6Adress ? t6Adress.v : undefined);
        let u6 = "U6"; let u6Adress = worksheet[u6]; let u6Value = (u6Adress ? u6Adress.v : undefined);
        let v6 = "V6"; let v6Adress = worksheet[v6]; let v6Value = (v6Adress ? v6Adress.v : undefined);
        let w6 = "W6"; let w6Adress = worksheet[w6]; let w6Value = (w6Adress ? w6Adress.v : undefined);
        let x6 = "X6"; let x6Adress = worksheet[x6]; let x6Value = (x6Adress ? x6Adress.v : undefined);
        let y6 = "Y6"; let y6Adress = worksheet[y6]; let y6Value = (y6Adress ? y6Adress.v : undefined);
        let z6 = "Z6"; let z6Adress = worksheet[z6]; let z6Value = (z6Adress ? z6Adress.v : undefined);
        let aa6 = "AA6"; let aa6Adress = worksheet[aa6]; let aa6Value = (aa6Adress ? aa6Adress.v : undefined);
        let ab6 = "AB6"; let ab6Adress = worksheet[ab6]; let ab6Value = (ab6Adress ? ab6Adress.v : undefined);
        let ac6 = "AC6"; let ac6Adress = worksheet[ac6]; let ac6Value = (ac6Adress ? ac6Adress.v : undefined);
        let ad6 = "AD6"; let ad6Adress = worksheet[ad6]; let ad6Value = (ad6Adress ? ad6Adress.v : undefined);
        let ae6 = "AE6"; let ae6Adress = worksheet[ae6]; let ae6Value = (ae6Adress ? ae6Adress.v : undefined);
        let af6 = "AF6"; let af6Adress = worksheet[af6]; let af6Value = (af6Adress ? af6Adress.v : undefined);
        let ag6 = "AG6"; let ag6Adress = worksheet[ag6]; let ag6Value = (ag6Adress ? ag6Adress.v : undefined);
        let ah6 = "AH6"; let ah6Adress = worksheet[ah6]; let ah6Value = (ah6Adress ? ah6Adress.v : undefined);
        let ai6 = "AI6"; let ai6Adress = worksheet[ai6]; let ai6Value = (ai6Adress ? ai6Adress.v : undefined);
        let aj6 = "AJ6"; let aj6Adress = worksheet[aj6]; let aj6Value = (aj6Adress ? aj6Adress.v : undefined);
        let ak6 = "AK6"; let ak6Adress = worksheet[ak6]; let ak6Value = (ak6Adress ? ak6Adress.v : undefined);
        let al6 = "AL6"; let al6Adress = worksheet[al6]; let al6Value = (al6Adress ? al6Adress.v : undefined);
        let am6 = "AM6"; let am6Adress = worksheet[am6]; let am6Value = (am6Adress ? am6Adress.v : undefined);
        let an6 = "AN6"; let an6Adress = worksheet[an6]; let an6Value = (an6Adress ? an6Adress.v : undefined);
        let ao6 = "AO6"; let ao6Adress = worksheet[ao6]; let ao6Value = (ao6Adress ? ao6Adress.v : undefined);
        let ap6 = "AP6"; let ap6Adress = worksheet[ap6]; let ap6Value = (ap6Adress ? ap6Adress.v : undefined);
        let aq6 = "AQ6"; let aq6Adress = worksheet[aq6]; let aq6Value = (aq6Adress ? aq6Adress.v : undefined);
        let ar6 = "AR6"; let ar6Adress = worksheet[ar6]; let ar6Value = (ar6Adress ? ar6Adress.v : undefined);
        let as6 = "AS6"; let as6Adress = worksheet[as6]; let as6Value = (as6Adress ? as6Adress.v : undefined);
        let at6 = "AT6"; let at6Adress = worksheet[at6]; let at6Value = (at6Adress ? at6Adress.v : undefined);
        let au6 = "AU6"; let au6Adress = worksheet[au6]; let au6Value = (au6Adress ? au6Adress.v : undefined);
        let av6 = "AV6"; let av6Adress = worksheet[av6]; let av6Value = (av6Adress ? av6Adress.v : undefined);
        let aw6 = "AW6"; let aw6Adress = worksheet[aw6]; let aw6Value = (aw6Adress ? aw6Adress.v : undefined);
        let ax6 = "AX6"; let ax6Adress = worksheet[ax6]; let ax6Value = (ax6Adress ? ax6Adress.v : undefined);
        let ay6 = "AY6"; let ay6Adress = worksheet[ay6]; let ay6Value = (ay6Adress ? ay6Adress.v : undefined);
        let az6 = "AZ6"; let az6Adress = worksheet[az6]; let az6Value = (az6Adress ? az6Adress.v : undefined);
        let ba6 = "BA6"; let ba6Adress = worksheet[ba6]; let ba6Value = (ba6Adress ? ba6Adress.v : undefined);
        let bb6 = "BB6"; let bb6Adress = worksheet[bb6]; let bb6Value = (bb6Adress ? bb6Adress.v : undefined);
        let bc6 = "BC6"; let bc6Adress = worksheet[bc6]; let bc6Value = (bc6Adress ? bc6Adress.v : undefined);
        let bd6 = "BD6"; let bd6Adress = worksheet[bd6]; let bd6Value = (bd6Adress ? bd6Adress.v : undefined);
        let be6 = "BE6"; let be6Adress = worksheet[be6]; let be6Value = (be6Adress ? be6Adress.v : undefined);
        let bf6 = "BF6"; let bf6Adress = worksheet[bf6]; let bf6Value = (bf6Adress ? bf6Adress.v : undefined);
        let bg6 = "BG6"; let bg6Adress = worksheet[bg6]; let bg6Value = (bg6Adress ? bg6Adress.v : undefined);
        let bh6 = "BH6"; let bh6Adress = worksheet[bh6]; let bh6Value = (bh6Adress ? bh6Adress.v : undefined);
        let bi6 = "BI6"; let bi6Adress = worksheet[bi6]; let bi6Value = (bi6Adress ? bi6Adress.v : undefined);
        let bj6 = "BJ6"; let bj6Adress = worksheet[bj6]; let bj6Value = (bj6Adress ? bj6Adress.v : undefined);
        let bk6 = "BK6"; let bk6Adress = worksheet[bk6]; let bk6Value = (bk6Adress ? bk6Adress.v : undefined);
        let bl6 = "BL6"; let bl6Adress = worksheet[bl6]; let bl6Value = (bl6Adress ? bl6Adress.v : undefined);
        let bm6 = "BM6"; let bm6Adress = worksheet[bm6]; let bm6Value = (bm6Adress ? bm6Adress.v : undefined);
        let bn6 = "BN6"; let bn6Adress = worksheet[bn6]; let bn6Value = (bn6Adress ? bn6Adress.v : undefined);
        let bo6 = "BO6"; let bo6Adress = worksheet[bo6]; let bo6Value = (bo6Adress ? bo6Adress.v : undefined);
        let bp6 = "BP6"; let bp6Adress = worksheet[bp6]; let bp6Value = (bp6Adress ? bp6Adress.v : undefined);
        let bq6 = "BQ6"; let bq6Adress = worksheet[bq6]; let bq6Value = (bq6Adress ? bq6Adress.v : undefined);
        let br6 = "BR6"; let br6Adress = worksheet[br6]; let br6Value = (br6Adress ? br6Adress.v : undefined);
        let bs6 = "BS6"; let bs6Adress = worksheet[bs6]; let bs6Value = (bs6Adress ? bs6Adress.v : undefined);
        let bt6 = "BT6"; let bt6Adress = worksheet[bt6]; let bt6Value = (bt6Adress ? bt6Adress.v : undefined);
        let bu6 = "BU6"; let bu6Adress = worksheet[bu6]; let bu6Value = (bu6Adress ? bu6Adress.v : undefined);
        let bv6 = "BV6"; let bv6Adress = worksheet[bv6]; let bv6Value = (bv6Adress ? bv6Adress.v : undefined);
        let bw6 = "BW6"; let bw6Adress = worksheet[bw6]; let bw6Value = (bw6Adress ? bw6Adress.v : undefined);

        let a7 = "A7"; let a7Adress = worksheet[a7]; let a7Value = (a7Adress ? a7Adress.v : undefined);
        let b7 = "B7"; let b7Adress = worksheet[b7]; let b7Value = (b7Adress ? b7Adress.v : undefined);
        let c7 = "C7"; let c7Adress = worksheet[c7]; let c7Value = (c7Adress ? c7Adress.v : undefined);
        let d7 = "D7"; let d7Adress = worksheet[d7]; let d7Value = (d7Adress ? d7Adress.v : undefined);
        let e7 = "E7"; let e7Adress = worksheet[e7]; let e7Value = (e7Adress ? e7Adress.v : undefined);
        let f7 = "F7"; let f7Adress = worksheet[f7]; let f7Value = (f7Adress ? f7Adress.v : undefined);
        let g7 = "G7"; let g7Adress = worksheet[g7]; let g7Value = (g7Adress ? g7Adress.v : undefined);
        let h7 = "H7"; let h7Adress = worksheet[h7]; let h7Value = (h7Adress ? h7Adress.v : undefined);
        let i7 = "I7"; let i7Adress = worksheet[i7]; let i7Value = (i7Adress ? i7Adress.v : undefined);
        let j7 = "J7"; let j7Adress = worksheet[j7]; let j7Value = (j7Adress ? j7Adress.v : undefined);
        let k7 = "K7"; let k7Adress = worksheet[k7]; let k7Value = (k7Adress ? k7Adress.v : undefined);
        let l7 = "L7"; let l7Adress = worksheet[l7]; let l7Value = (l7Adress ? l7Adress.v : undefined);
        let m7 = "M7"; let m7Adress = worksheet[m7]; let m7Value = (m7Adress ? m7Adress.v : undefined);
        let n7 = "N7"; let n7Adress = worksheet[n7]; let n7Value = (n7Adress ? n7Adress.v : undefined);
        let o7 = "O7"; let o7Adress = worksheet[o7]; let o7Value = (o7Adress ? o7Adress.v : undefined);
        let p7 = "P7"; let p7Adress = worksheet[p7]; let p7Value = (p7Adress ? p7Adress.v : undefined);
        let q7 = "Q7"; let q7Adress = worksheet[q7]; let q7Value = (q7Adress ? q7Adress.v : undefined);
        let r7 = "R7"; let r7Adress = worksheet[r7]; let r7Value = (r7Adress ? r7Adress.v : undefined);
        let s7 = "S7"; let s7Adress = worksheet[s7]; let s7Value = (s7Adress ? s7Adress.v : undefined);
        let t7 = "T7"; let t7Adress = worksheet[t7]; let t7Value = (t7Adress ? t7Adress.v : undefined);
        let u7 = "U7"; let u7Adress = worksheet[u7]; let u7Value = (u7Adress ? u7Adress.v : undefined);
        let v7 = "V7"; let v7Adress = worksheet[v7]; let v7Value = (v7Adress ? v7Adress.v : undefined);
        let w7 = "W7"; let w7Adress = worksheet[w7]; let w7Value = (w7Adress ? w7Adress.v : undefined);
        let x7 = "X7"; let x7Adress = worksheet[x7]; let x7Value = (x7Adress ? x7Adress.v : undefined);
        let y7 = "Y7"; let y7Adress = worksheet[y7]; let y7Value = (y7Adress ? y7Adress.v : undefined);
        let z7 = "Z7"; let z7Adress = worksheet[z7]; let z7Value = (z7Adress ? z7Adress.v : undefined);
        let aa7 = "AA7"; let aa7Adress = worksheet[aa7]; let aa7Value = (aa7Adress ? aa7Adress.v : undefined);
        let ab7 = "AB7"; let ab7Adress = worksheet[ab7]; let ab7Value = (ab7Adress ? ab7Adress.v : undefined);
        let ac7 = "AC7"; let ac7Adress = worksheet[ac7]; let ac7Value = (ac7Adress ? ac7Adress.v : undefined);
        let ad7 = "AD7"; let ad7Adress = worksheet[ad7]; let ad7Value = (ad7Adress ? ad7Adress.v : undefined);
        let ae7 = "AE7"; let ae7Adress = worksheet[ae7]; let ae7Value = (ae7Adress ? ae7Adress.v : undefined);
        let af7 = "AF7"; let af7Adress = worksheet[af7]; let af7Value = (af7Adress ? af7Adress.v : undefined);
        let ag7 = "AG7"; let ag7Adress = worksheet[ag7]; let ag7Value = (ag7Adress ? ag7Adress.v : undefined);
        let ah7 = "AH7"; let ah7Adress = worksheet[ah7]; let ah7Value = (ah7Adress ? ah7Adress.v : undefined);
        let ai7 = "AI7"; let ai7Adress = worksheet[ai7]; let ai7Value = (ai7Adress ? ai7Adress.v : undefined);
        let aj7 = "AJ7"; let aj7Adress = worksheet[aj7]; let aj7Value = (aj7Adress ? aj7Adress.v : undefined);
        let ak7 = "AK7"; let ak7Adress = worksheet[ak7]; let ak7Value = (ak7Adress ? ak7Adress.v : undefined);
        let al7 = "AL7"; let al7Adress = worksheet[al7]; let al7Value = (al7Adress ? al7Adress.v : undefined);
        let am7 = "AM7"; let am7Adress = worksheet[am7]; let am7Value = (am7Adress ? am7Adress.v : undefined);
        let an7 = "AN7"; let an7Adress = worksheet[an7]; let an7Value = (an7Adress ? an7Adress.v : undefined);
        let ao7 = "AO7"; let ao7Adress = worksheet[ao7]; let ao7Value = (ao7Adress ? ao7Adress.v : undefined);
        let ap7 = "AP7"; let ap7Adress = worksheet[ap7]; let ap7Value = (ap7Adress ? ap7Adress.v : undefined);
        let aq7 = "AQ7"; let aq7Adress = worksheet[aq7]; let aq7Value = (aq7Adress ? aq7Adress.v : undefined);
        let ar7 = "AR7"; let ar7Adress = worksheet[ar7]; let ar7Value = (ar7Adress ? ar7Adress.v : undefined);
        let as7 = "AS7"; let as7Adress = worksheet[as7]; let as7Value = (as7Adress ? as7Adress.v : undefined);
        let at7 = "AT7"; let at7Adress = worksheet[at7]; let at7Value = (at7Adress ? at7Adress.v : undefined);
        let au7 = "AU7"; let au7Adress = worksheet[au7]; let au7Value = (au7Adress ? au7Adress.v : undefined);
        let av7 = "AV7"; let av7Adress = worksheet[av7]; let av7Value = (av7Adress ? av7Adress.v : undefined);
        let aw7 = "AW7"; let aw7Adress = worksheet[aw7]; let aw7Value = (aw7Adress ? aw7Adress.v : undefined);
        let ax7 = "AX7"; let ax7Adress = worksheet[ax7]; let ax7Value = (ax7Adress ? ax7Adress.v : undefined);
        let ay7 = "AY7"; let ay7Adress = worksheet[ay7]; let ay7Value = (ay7Adress ? ay7Adress.v : undefined);
        let az7 = "AZ7"; let az7Adress = worksheet[az7]; let az7Value = (az7Adress ? az7Adress.v : undefined);
        let ba7 = "BA7"; let ba7Adress = worksheet[ba7]; let ba7Value = (ba7Adress ? ba7Adress.v : undefined);
        let bb7 = "BB7"; let bb7Adress = worksheet[bb7]; let bb7Value = (bb7Adress ? bb7Adress.v : undefined);
        let bc7 = "BC7"; let bc7Adress = worksheet[bc7]; let bc7Value = (bc7Adress ? bc7Adress.v : undefined);
        let bd7 = "BD7"; let bd7Adress = worksheet[bd7]; let bd7Value = (bd7Adress ? bd7Adress.v : undefined);
        let be7 = "BE7"; let be7Adress = worksheet[be7]; let be7Value = (be7Adress ? be7Adress.v : undefined);
        let bf7 = "BF7"; let bf7Adress = worksheet[bf7]; let bf7Value = (bf7Adress ? bf7Adress.v : undefined);
        let bg7 = "BG7"; let bg7Adress = worksheet[bg7]; let bg7Value = (bg7Adress ? bg7Adress.v : undefined);
        let bh7 = "BH7"; let bh7Adress = worksheet[bh7]; let bh7Value = (bh7Adress ? bh7Adress.v : undefined);
        let bi7 = "BI7"; let bi7Adress = worksheet[bi7]; let bi7Value = (bi7Adress ? bi7Adress.v : undefined);
        let bj7 = "BJ7"; let bj7Adress = worksheet[bj7]; let bj7Value = (bj7Adress ? bj7Adress.v : undefined);
        let bk7 = "BK7"; let bk7Adress = worksheet[bk7]; let bk7Value = (bk7Adress ? bk7Adress.v : undefined);
        let bl7 = "BL7"; let bl7Adress = worksheet[bl7]; let bl7Value = (bl7Adress ? bl7Adress.v : undefined);
        let bm7 = "BM7"; let bm7Adress = worksheet[bm7]; let bm7Value = (bm7Adress ? bm7Adress.v : undefined);
        let bn7 = "BN7"; let bn7Adress = worksheet[bn7]; let bn7Value = (bn7Adress ? bn7Adress.v : undefined);
        let bo7 = "BO7"; let bo7Adress = worksheet[bo7]; let bo7Value = (bo7Adress ? bo7Adress.v : undefined);
        let bp7 = "BP7"; let bp7Adress = worksheet[bp7]; let bp7Value = (bp7Adress ? bp7Adress.v : undefined);
        let bq7 = "BQ7"; let bq7Adress = worksheet[bq7]; let bq7Value = (bq7Adress ? bq7Adress.v : undefined);
        let br7 = "BR7"; let br7Adress = worksheet[br7]; let br7Value = (br7Adress ? br7Adress.v : undefined);
        let bs7 = "BS7"; let bs7Adress = worksheet[bs7]; let bs7Value = (bs7Adress ? bs7Adress.v : undefined);
        let bt7 = "BT7"; let bt7Adress = worksheet[bt7]; let bt7Value = (bt7Adress ? bt7Adress.v : undefined);
        let bu7 = "BU7"; let bu7Adress = worksheet[bu7]; let bu7Value = (bu7Adress ? bu7Adress.v : undefined);
        let bv7 = "BV7"; let bv7Adress = worksheet[bv7]; let bv7Value = (bv7Adress ? bv7Adress.v : undefined);
        let bw7 = "BW7"; let bw7Adress = worksheet[bw7]; let bw7Value = (bw7Adress ? bw7Adress.v : undefined);

        xw = new XMLWriter(true);
        xw.startDocument("1.0", "UTF-8");
        xw.startElement("archivo").writeAttribute("xsi:schemaLocation", "http://www.uif.shcp.gob.mx/recepcion/ari ari.xsd").writeAttribute("xmlns", "http://www.uif.shcp.gob.mx/recepcion/ari").writeAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
          xw.startElement("informe");
            /************************************** INICIO DE FILA **************************************/
            if(b5Value != undefined){
              if(b5Value != undefined){xw.startElement("mes_reportado").text(b5Value).endElement();}
            }//if b5Value
            if(a5Value != undefined){
              xw.startElement("sujeto_obligado");
                xw.startElement("clave_sujeto_obligado").text(a5Value.toUpperCase().trim()).endElement();
                xw.startElement("clave_actividad").text("ARI").endElement();
              xw.endElement(); //sujeto obligado
            }
            if(a5Value != undefined){ xw.startElement("aviso");
              if(c5Value != undefined){xw.startElement("referencia_aviso").text(c5Value).endElement();}
              if(e5Value != undefined){xw.startElement("prioridad").text(e5Value).endElement();}
              if(d5Value != undefined){
                xw.startElement("alerta");
                  xw.startElement("tipo_alerta").text(d5Value).endElement();
                  if(f5Value != undefined){xw.startElement("descripcion_alerta").text(f5Value.trim()).endElement();}
                xw.endElement(); //alerta
              } //if d5Value
              xw.startElement("persona_aviso");
                xw.startElement("tipo_persona");
                  if(g5Value != undefined){
                    xw.startElement("persona_física");
                      xw.startElement("nombre").text(g5Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();
                      if(h5Value != undefined){xw.startElement("apellido_paterno").text(h5Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                      if(i5Value != undefined){xw.startElement("apellido_materno").text(i5Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                      if(j5Value != undefined){
                        let j5Month = j5Value.getMonth(); j5Month++; if (j5Month < 10) { j5Month = "0" + j5Month; }
                        let j5Day = j5Value.getDate(); if (j5Day < 10) { j5Day = "0" + j5Day; }
                        xw.startElement("fecha_nacimiento").text(j5Value.getFullYear() + "" + j5Month + "" + j5Day).endElement();
                      } //ifj5Value
                      if(k5Value != undefined){xw.startElement("rfc").text(k5Value.toUpperCase().trim()).endElement();}
                      if(l5Value != undefined){xw.startElement("curp").text(l5Value.toUpperCase().trim()).endElement();}
                      if(m5Value != undefined){xw.startElement("pais_nacionalidad").text(m5Value.toUpperCase().trim()).endElement();}
                      if(n5Value != undefined){xw.startElement("actividad_economica").text(n5Value).endElement();}
                    xw.endElement(); //persona física
                  } /*if g5Value*/ else if(o5Value != undefined){
                    xw.startElement("persona_moral");
                      if(o5Value != undefined){xw.startElement("denominacion_razon").text(o5Value.toUpperCase().trim()).endElement();}
                      if(p5Value != undefined){
                        let p5Month = p5Value.getMonth(); p5Month++; if (p5Month < 10) { p5Month = "0" + p5Month; }
                        let p5Day = p5Value.getDate(); if (p5Day < 10) { p5Day = "0" + p5Day; }
                        xw.startElement("fecha_constitucion").text(p5Value.getFullYear() + "" + p5Month + "" + p5Day).endElement();
                      } //if p5Value
                      if(q5Value != undefined){xw.startElement("rfc").text(q5Value.toUpperCase().trim()).endElement();}
                      if(r5Value != undefined){xw.startElement("pais_nacionalidad").text(r5Value.toUpperCase().trim()).endElement();}
                      if(s5Value != undefined){xw.startElement("giro_mercanril").text(s5Value).endElement();}
                      xw.startElement("representante_apoderado");
                        if(w5Value != undefined){xw.startElement("nombre").text(w5Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(x5Value != undefined){xw.startElement("apellido_paterno").text(x5Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(y5Value != undefined){xw.startElement("apellido_materno").text(y5Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(z5Value != undefined){
                          let z5Month = z5Value.getMonth(); z5Month++; if (z5Month < 10) { z5Month = "0" + z5Month; }
                          let z5Day = z5Value.getDate(); if (z5Day < 10) { z5Day = "0" + z5Day; }
                          xw.startElement("fecha_nacimiento").text(z5Value.getFullYear() + "" + z5Month + "" + z5Day).endElement();
                        } //if z5Value
                        if(aa5Value != undefined){xw.startElement("rfc").text(aa5Value.toUpperCase().trim()).endElement();}
                        if(ab5Value != undefined){xw.startElement("curp").text(ab5Value.toUpperCase().trim()).endElement();}
                      xw.endElement(); //representante apoderado persona moral
                    xw.endElement(); //persona moral
                  } /*if o5Value*/ else if(t5Value != undefined){
                    xw.startElement("fideicomiso");
                      if(t5Value != undefined){xw.startElement("denominacion_razon").text(t5Value.toUpperCase().trim()).endElement();}
                      if(u5Value != undefined){xw.startElement("rfc").text(u5Value.toUpperCase().trim()).endElement();}
                      if(v5Value != undefined){xw.startElement("identificador_fideicomiso").text(v5Value).endElement();}
                      xw.startElement("representante_apoderado");
                        if(w5Value != undefined){xw.startElement("nombre").text(w5Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(x5Value != undefined){xw.startElement("apellido_paterno").text(x5Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(y5Value != undefined){xw.startElement("apellido_materno").text(y5Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(z5Value != undefined){
                          let z5Month = z5Value.getMonth(); z5Month++; if (z5Month < 10) { z5Month = "0" + z5Month; }
                          let z5Day = z5Value.getDate(); if (z5Day < 10) { z5Day = "0" + z5Day; }
                          xw.startElement("fecha_nacimiento").text(z5Value.getFullYear() + "" + z5Month + "" + z5Day).endElement();
                        } //if z5Value
                        if(aa5Value != undefined){xw.startElement("rfc").text(aa5Value.toUpperCase().trim()).endElement();}
                        if(ab5Value != undefined){xw.startElement("curp").text(ab5Value.toUpperCase().trim()).endElement();}
                      xw.endElement(); //representante apoderado fideicomiso
                    xw.endElement(); //fideicomiso
                  }
                xw.endElement(); //tipo persona
                xw.startElement("tipo_domicilio");
                  if(ac5Value != undefined){
                    xw.startElement("nacional");
                      if(ad5Value != undefined){xw.startElement("colonia").text(ad5Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                      if(ae5Value != undefined){xw.startElement("calle").text(ae5Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                      if(af5Value != undefined){xw.startElement("numero_exterior").text(af5Value).endElement();}
                      if(ag5Value != undefined){xw.startElement("numero_interior").text(ag5Value).endElement();}
                      if(ac5Value != undefined){xw.startElement("codigo_postal").text(ac5Value).endElement();}
                    xw.endElement();
                  }else if(ah5Value != undefined) {
                    xw.startElement("extranjero");
                      if(ah5Value != undefined){xw.startElement("pais").text(ah5Value.toUpperCase()).endElement();}
                      if(ai5Value != undefined){xw.startElement("estado_provincia").text(ai5Value).endElement();}
                      if(aj5Value != undefined){xw.startElement("ciudad_poblacion").text(aj5Value).endElement();}
                      if(ak5Value != undefined){xw.startElement("colonia").text(ak5Value).endElement();}
                      if(al5Value != undefined){xw.startElement("calle").text(al5Value).endElement();}
                      if(am5Value != undefined){xw.startElement("numero_exterior").text(am5Value).endElement();}
                      if(an5Value != undefined){xw.startElement("numero_interior").text(an5Value).endElement();}
                      if(ao5Value != undefined){xw.startElement("codigo_postal").text(ao5Value).endElement();}
                    xw.endElement();
                  }
                xw.endElement(); //tipo domicilio
                if(ap5Value != undefined || aq5Value != undefined || ar5Value != undefined){
                  xw.startElement("telefono");
                    if(ap5Value != undefined){xw.startElement("clave_pais").text(ap5Value.toUpperCase().trim()).endElement();}
                    if(aq5Value != undefined){xw.startElement("numero_telefono").text(aq5Value.toString().replace("    ","").replace("  ","").replace(" ","")).endElement();}
                    if(ar5Value != undefined){xw.startElement("correo_electronico").text(ar5Value.toUpperCase().trim()).endElement();}
                  xw.endElement(); //telefono
                } //if ap5Value || aq5Value || ar5Value
              xw.endElement(); //persona aviso
              if(as5Value != undefined || az5Value != undefined || bd5Value != undefined){
                xw.startElement("dueno_beneficiario");
                  xw.startElement("tipo_persona");
                    if(as5Value != undefined){
                      xw.startElement("persona_fisica");
                        xw.startElement("nombre").text(as5Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();
                        if(at5Value != undefined){xw.startElement("apellido_paterno").text(at5Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(au5Value != undefined){xw.startElement("apellido_materno").text(au5Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(av5Value != undefined){
                          let av5Month = av5Value.getMonth(); av5Month++; if (av5Month < 10) { av5Month = "0" + av5Month; }
                          let av5Day = av5Value.getDate(); if (av5Day < 10) { av5Day = "0" + av5Day; }
                          xw.startElement("fecha_nacimiento").text(av5Value.getFullYear() + "" + av5Month + "" + av5Day).endElement();
                        } //if av5Value
                        if(aw5Value != undefined){xw.startElement("rfc").text(aw5Value.toUpperCase().trim()).endElement();}
                        if(ax5Value != undefined){xw.startElement("curp").text(ax5Value.toUpperCase().trim()).endElement();}
                        if(ay5Value != undefined){xw.startElement("pais_nacionalidad").text(ay5Value.toUpperCase()).endElement();}
                      xw.endElement(); //persona física
                    } /*if as5Value*/ else if(az5Value != undefined){
                      xw.startElement("persona_moral");
                        xw.startElement("denominacion_razon").text(az5Value.toUpperCase().trim()).endElement();
                        if(ba5Value != undefined){
                          let ba5Month = ba5Value.getMonth(); ba5Month++; if (ba5Month < 10) { ba5Month = "0" + ba5Month; }
                          let ba5Day = ba5Value.getDate(); if (ba5Day < 10) { ba5Day = "0" + ba5Day; }
                          xw.startElement("fecha_constitucion").text(ba5Value.getFullYear() + "" + ba5Month + "" + ba5Day).endElement();
                        } //if ba5Value
                        if(bb5Value != undefined){xw.startElement("rfc").text(bb5Value.toUpperCase().trim()).endElement();}
                        if(bc5Value != undefined){xw.startElement("pais_nacionalidad").text(bc5Value.toUpperCase()).endElement();}
                      xw.endElement(); //persona moral
                    } /*if az5Value*/ else if(bd5Value != undefined){
                      xw.startElement("fideicomiso");
                        xw.startElement("denominacion_razon").text(bd5Value.toUpperCase().trim()).endElement();
                        if(be5Value != undefined){xw.startElement("rfc").text(be5Value.toUpperCase().trim()).endElement();}
                        if(bf5Value != undefined){xw.startElement("identificador_fideicomiso").text(bf5Value).endElement();}
                      xw.endElement(); //fideicomiso
                    } //if bd5Value
                  xw.endElement(); //tipo persona beneficiario
                xw.endElement(); //dueño beneficiario
              } //if as5Value || az5Value || bd5Value
              if(bg5Value != undefined){
                xw.startElement("detalle_operaciones");
                  xw.startElement("datos_operacion");
                    if(bg5Value != undefined){
                      let bg5Month = bg5Value.getMonth(); bg5Month++; if (bg5Month < 10) { bg5Month = "0" + bg5Month; }
                      let bg5Day = bg5Value.getDate(); if (bg5Day < 10) { bg5Day = "0" + bg5Day; }
                      xw.startElement("fecha_operacion").text(bg5Value.getFullYear() + "" + bg5Month + "" + bg5Day).endElement();
                    } //if bg5Value
                    if(bh5Value != undefined){xw.startElement("tipo_operacion").text(bh5Value).endElement();}
                    xw.startElement("caracteristicas");
                      if(bq5Value != undefined){
                        let bq5Month = bq5Value.getMonth(); bq5Month++; if (bq5Month < 10) { bq5Month = "0" + bq5Month; }
                        let bq5Day = bq5Value.getDate(); if (bq5Day < 10) { bq5Day = "0" + bq5Day; }
                        xw.startElement("fecha_inicio").text(bq5Value.getFullYear() + "" + bq5Month + "" + bq5Day).endElement();
                      } //if bq5Value
                      if(br5Value != undefined){
                        let br5Month = br5Value.getMonth(); br5Month++; if (br5Month < 10) { br5Month = "0" + br5Month; }
                        let br5Day = br5Value.getDate(); if (br5Day < 10) { br5Day = "0" + br5Day; }
                        xw.startElement("fecha_termino").text(br5Value.getFullYear() + "" + br5Month + "" + br5Day).endElement();
                      } //if br5Value
                      if(bi5Value != undefined){xw.startElement("tipo_inmueble").text(bi5Value).endElement();}
                      if(bj5Value != undefined){xw.startElement("valor_referencia").text(bj5Value.toFixed(2)).endElement();}
                      if(bl5Value != undefined){xw.startElement("colonia").text(bl5Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                      if(bm5Value != undefined){xw.startElement("calle").text(bm5Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                      if(bn5Value != undefined){xw.startElement("numero_exterior").text(bn5Value).endElement();}
                      if(bo5Value != undefined){xw.startElement("numero_interior").text(bo5Value).endElement();}
                      if(bk5Value != undefined){xw.startElement("codigo_postal").text(bk5Value).endElement();}
                      if(bp5Value != undefined){xw.startElement("folio_real").text(bp5Value).endElement();}
                    xw.endElement(); //caracteristicas
                    if(bt5Value != undefined){
                      xw.startElement("datos_liquidacion");
                        if(bs5Value != undefined){
                          let bs5Month = bs5Value.getMonth(); bs5Month++; if (bs5Month < 10) { bs5Month = "0" + bs5Month; }
                          let bs5Day = bs5Value.getDate(); if (bs5Day < 10) { bs5Day = "0" + bs5Day; }
                          xw.startElement("fecha_pago").text(bs5Value.getFullYear() + "" + bs5Month + "" + bs5Day).endElement();
                        } //if bs5Value
                        if(bt5Value != undefined){xw.startElement("forma_pago").text(bt5Value).endElement();}
                        if(bu5Value != undefined){xw.startElement("instrumento_monetario").text(bu5Value).endElement();}
                        if(bv5Value != undefined){xw.startElement("moneda").text(bv5Value).endElement();}
                        if(bw5Value != undefined){xw.startElement("monto_operacion").text(bw5Value.toFixed(2)).endElement();}
                      xw.endElement(); //datos liquidación
                    } //if bt5Value
                  xw.endElement(); //datos operacion
                xw.endElement(); //detalle operaciones
              } //if bg5Value
            xw.endElement();} //aviso
            /************************************** FIN DE FILA **************************************/
            /************************************** INICIO DE FILA 2 **************************************/
            if(a6Value != undefined){ xw.startElement("aviso");
              if(c6Value != undefined){xw.startElement("referencia_aviso").text(c6Value).endElement();}
              if(e6Value != undefined){xw.startElement("prioridad").text(e6Value).endElement();}
              if(d6Value != undefined){
                xw.startElement("alerta");
                  xw.startElement("tipo_alerta").text(d6Value).endElement();
                  if(f6Value != undefined){xw.startElement("descripcion_alerta").text(f6Value.trim()).endElement();}
                xw.endElement(); //alerta
              } //if d6Value
              xw.startElement("persona_aviso");
                xw.startElement("tipo_persona");
                  if(g6Value != undefined){
                    xw.startElement("persona_física");
                      xw.startElement("nombre").text(g6Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();
                      if(h6Value != undefined){xw.startElement("apellido_paterno").text(h6Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                      if(i6Value != undefined){xw.startElement("apellido_materno").text(i6Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                      if(j6Value != undefined){
                        let j6Month = j6Value.getMonth(); j6Month++; if (j6Month < 10) { j6Month = "0" + j6Month; }
                        let j6Day = j6Value.getDate(); if (j6Day < 10) { j6Day = "0" + j6Day; }
                        xw.startElement("fecha_nacimiento").text(j6Value.getFullYear() + "" + j6Month + "" + j6Day).endElement();
                      } //ifj6Value
                      if(k6Value != undefined){xw.startElement("rfc").text(k6Value.toUpperCase().trim()).endElement();}
                      if(l6Value != undefined){xw.startElement("curp").text(l6Value.toUpperCase().trim()).endElement();}
                      if(m6Value != undefined){xw.startElement("pais_nacionalidad").text(m6Value.toUpperCase().trim()).endElement();}
                      if(n6Value != undefined){xw.startElement("actividad_economica").text(n6Value).endElement();}
                    xw.endElement(); //persona física
                  } /*if g6Value*/ else if(o6Value != undefined){
                    xw.startElement("persona_moral");
                      if(o6Value != undefined){xw.startElement("denominacion_razon").text(o6Value.toUpperCase().trim()).endElement();}
                      if(p6Value != undefined){
                        let p6Month = p6Value.getMonth(); p6Month++; if (p6Month < 10) { p6Month = "0" + p6Month; }
                        let p6Day = p6Value.getDate(); if (p6Day < 10) { p6Day = "0" + p6Day; }
                        xw.startElement("fecha_constitucion").text(p6Value.getFullYear() + "" + p6Month + "" + p6Day).endElement();
                      } //if p6Value
                      if(q6Value != undefined){xw.startElement("rfc").text(q6Value.toUpperCase().trim()).endElement();}
                      if(r6Value != undefined){xw.startElement("pais_nacionalidad").text(r6Value.toUpperCase().trim()).endElement();}
                      if(s6Value != undefined){xw.startElement("giro_mercanril").text(s6Value).endElement();}
                      xw.startElement("representante_apoderado");
                        if(w6Value != undefined){xw.startElement("nombre").text(w6Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(x6Value != undefined){xw.startElement("apellido_paterno").text(x6Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(y6Value != undefined){xw.startElement("apellido_materno").text(y6Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(z6Value != undefined){
                          let z6Month = z6Value.getMonth(); z6Month++; if (z6Month < 10) { z6Month = "0" + z6Month; }
                          let z6Day = z6Value.getDate(); if (z6Day < 10) { z6Day = "0" + z6Day; }
                          xw.startElement("fecha_nacimiento").text(z6Value.getFullYear() + "" + z6Month + "" + z6Day).endElement();
                        } //if z6Value
                        if(aa6Value != undefined){xw.startElement("rfc").text(aa6Value.toUpperCase().trim()).endElement();}
                        if(ab6Value != undefined){xw.startElement("curp").text(ab6Value.toUpperCase().trim()).endElement();}
                      xw.endElement(); //representante apoderado persona moral
                    xw.endElement(); //persona moral
                  } /*if o6Value*/ else if(t6Value != undefined){
                    xw.startElement("fideicomiso");
                      if(t6Value != undefined){xw.startElement("denominacion_razon").text(t6Value.toUpperCase().trim()).endElement();}
                      if(u6Value != undefined){xw.startElement("rfc").text(u6Value.toUpperCase().trim()).endElement();}
                      if(v6Value != undefined){xw.startElement("identificador_fideicomiso").text(v6Value).endElement();}
                      xw.startElement("representante_apoderado");
                        if(w6Value != undefined){xw.startElement("nombre").text(w6Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(x6Value != undefined){xw.startElement("apellido_paterno").text(x6Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(y6Value != undefined){xw.startElement("apellido_materno").text(y6Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(z6Value != undefined){
                          let z6Month = z6Value.getMonth(); z6Month++; if (z6Month < 10) { z6Month = "0" + z6Month; }
                          let z6Day = z6Value.getDate(); if (z6Day < 10) { z6Day = "0" + z6Day; }
                          xw.startElement("fecha_nacimiento").text(z6Value.getFullYear() + "" + z6Month + "" + z6Day).endElement();
                        } //if z6Value
                        if(aa6Value != undefined){xw.startElement("rfc").text(aa6Value.toUpperCase().trim()).endElement();}
                        if(ab6Value != undefined){xw.startElement("curp").text(ab6Value.toUpperCase().trim()).endElement();}
                      xw.endElement(); //representante apoderado fideicomiso
                    xw.endElement(); //fideicomiso
                  }
                xw.endElement(); //tipo persona
                xw.startElement("tipo_domicilio");
                  if(ac6Value != undefined){
                    xw.startElement("nacional");
                      if(ad6Value != undefined){xw.startElement("colonia").text(ad6Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                      if(ae6Value != undefined){xw.startElement("calle").text(ae6Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                      if(af6Value != undefined){xw.startElement("numero_exterior").text(af6Value).endElement();}
                      if(ag6Value != undefined){xw.startElement("numero_interior").text(ag6Value).endElement();}
                      if(ac6Value != undefined){xw.startElement("codigo_postal").text(ac6Value).endElement();}
                    xw.endElement();
                  }else if(ah6Value != undefined) {
                    xw.startElement("extranjero");
                      if(ah6Value != undefined){xw.startElement("pais").text(ah6Value.toUpperCase()).endElement();}
                      if(ai6Value != undefined){xw.startElement("estado_provincia").text(ai6Value).endElement();}
                      if(aj6Value != undefined){xw.startElement("ciudad_poblacion").text(aj6Value).endElement();}
                      if(ak6Value != undefined){xw.startElement("colonia").text(ak6Value).endElement();}
                      if(al6Value != undefined){xw.startElement("calle").text(al6Value).endElement();}
                      if(am6Value != undefined){xw.startElement("numero_exterior").text(am6Value).endElement();}
                      if(an6Value != undefined){xw.startElement("numero_interior").text(an6Value).endElement();}
                      if(ao6Value != undefined){xw.startElement("codigo_postal").text(ao6Value).endElement();}
                    xw.endElement();
                  }
                xw.endElement(); //tipo domicilio
                if(ap6Value != undefined || aq6Value != undefined || ar6Value != undefined){
                  xw.startElement("telefono");
                    if(ap6Value != undefined){xw.startElement("clave_pais").text(ap6Value.toUpperCase().trim()).endElement();}
                    if(aq6Value != undefined){xw.startElement("numero_telefono").text(aq6Value.toString().replace("    ","").replace("  ","").replace(" ","")).endElement();}
                    if(ar6Value != undefined){xw.startElement("correo_electronico").text(ar6Value.toUpperCase().trim()).endElement();}
                  xw.endElement(); //telefono
                } //if ap6Value || aq6Value || ar6Value
              xw.endElement(); //persona aviso
              if(as6Value != undefined || az6Value != undefined || bd6Value != undefined){
                xw.startElement("dueno_beneficiario");
                  xw.startElement("tipo_persona");
                    if(as6Value != undefined){
                      xw.startElement("persona_fisica");
                        xw.startElement("nombre").text(as6Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();
                        if(at6Value != undefined){xw.startElement("apellido_paterno").text(at6Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(au6Value != undefined){xw.startElement("apellido_materno").text(au6Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(av6Value != undefined){
                          let av6Month = av6Value.getMonth(); av6Month++; if (av6Month < 10) { av6Month = "0" + av6Month; }
                          let av6Day = av6Value.getDate(); if (av6Day < 10) { av6Day = "0" + av6Day; }
                          xw.startElement("fecha_nacimiento").text(av6Value.getFullYear() + "" + av6Month + "" + av6Day).endElement();
                        } //if av6Value
                        if(aw6Value != undefined){xw.startElement("rfc").text(aw6Value.toUpperCase().trim()).endElement();}
                        if(ax6Value != undefined){xw.startElement("curp").text(ax6Value.toUpperCase().trim()).endElement();}
                        if(ay6Value != undefined){xw.startElement("pais_nacionalidad").text(ay6Value.toUpperCase()).endElement();}
                      xw.endElement(); //persona física
                    } /*if as6Value*/ else if(az6Value != undefined){
                      xw.startElement("persona_moral");
                        xw.startElement("denominacion_razon").text(az6Value.toUpperCase().trim()).endElement();
                        if(ba6Value != undefined){
                          let ba6Month = ba6Value.getMonth(); ba6Month++; if (ba6Month < 10) { ba6Month = "0" + ba6Month; }
                          let ba6Day = ba6Value.getDate(); if (ba6Day < 10) { ba6Day = "0" + ba6Day; }
                          xw.startElement("fecha_constitucion").text(ba6Value.getFullYear() + "" + ba6Month + "" + ba6Day).endElement();
                        } //if ba6Value
                        if(bb6Value != undefined){xw.startElement("rfc").text(bb6Value.toUpperCase().trim()).endElement();}
                        if(bc6Value != undefined){xw.startElement("pais_nacionalidad").text(bc6Value.toUpperCase()).endElement();}
                      xw.endElement(); //persona moral
                    } /*if az6Value*/ else if(bd6Value != undefined){
                      xw.startElement("fideicomiso");
                        xw.startElement("denominacion_razon").text(bd6Value.toUpperCase().trim()).endElement();
                        if(be6Value != undefined){xw.startElement("rfc").text(be6Value.toUpperCase().trim()).endElement();}
                        if(bf6Value != undefined){xw.startElement("identificador_fideicomiso").text(bf6Value).endElement();}
                      xw.endElement(); //fideicomiso
                    } //if bd6Value
                  xw.endElement(); //tipo persona beneficiario
                xw.endElement(); //dueño beneficiario
              } //if as6Value || az6Value || bd6Value
              if(bg6Value != undefined){
                xw.startElement("detalle_operaciones");
                  xw.startElement("datos_operacion");
                    if(bg6Value != undefined){
                      let bg6Month = bg6Value.getMonth(); bg6Month++; if (bg6Month < 10) { bg6Month = "0" + bg6Month; }
                      let bg6Day = bg6Value.getDate(); if (bg6Day < 10) { bg6Day = "0" + bg6Day; }
                      xw.startElement("fecha_operacion").text(bg6Value.getFullYear() + "" + bg6Month + "" + bg6Day).endElement();
                    } //if bg6Value
                    if(bh6Value != undefined){xw.startElement("tipo_operacion").text(bh6Value).endElement();}
                    xw.startElement("caracteristicas");
                      if(bq6Value != undefined){
                        let bq6Month = bq6Value.getMonth(); bq6Month++; if (bq6Month < 10) { bq6Month = "0" + bq6Month; }
                        let bq6Day = bq6Value.getDate(); if (bq6Day < 10) { bq6Day = "0" + bq6Day; }
                        xw.startElement("fecha_inicio").text(bq6Value.getFullYear() + "" + bq6Month + "" + bq6Day).endElement();
                      } //if bq6Value
                      if(br6Value != undefined){
                        let br6Month = br6Value.getMonth(); br6Month++; if (br6Month < 10) { br6Month = "0" + br6Month; }
                        let br6Day = br6Value.getDate(); if (br6Day < 10) { br6Day = "0" + br6Day; }
                        xw.startElement("fecha_termino").text(br6Value.getFullYear() + "" + br6Month + "" + br6Day).endElement();
                      } //if br6Value
                      if(bi6Value != undefined){xw.startElement("tipo_inmueble").text(bi6Value).endElement();}
                      if(bj6Value != undefined){xw.startElement("valor_referencia").text(bj6Value.toFixed(2)).endElement();}
                      if(bl6Value != undefined){xw.startElement("colonia").text(bl6Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                      if(bm6Value != undefined){xw.startElement("calle").text(bm6Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                      if(bn6Value != undefined){xw.startElement("numero_exterior").text(bn6Value).endElement();}
                      if(bo6Value != undefined){xw.startElement("numero_interior").text(bo6Value).endElement();}
                      if(bk6Value != undefined){xw.startElement("codigo_postal").text(bk6Value).endElement();}
                      if(bp6Value != undefined){xw.startElement("folio_real").text(bp6Value).endElement();}
                    xw.endElement(); //caracteristicas
                    if(bt6Value != undefined){
                      xw.startElement("datos_liquidacion");
                        if(bs6Value != undefined){
                          let bs6Month = bs6Value.getMonth(); bs6Month++; if (bs6Month < 10) { bs6Month = "0" + bs6Month; }
                          let bs6Day = bs6Value.getDate(); if (bs6Day < 10) { bs6Day = "0" + bs6Day; }
                          xw.startElement("fecha_pago").text(bs6Value.getFullYear() + "" + bs6Month + "" + bs6Day).endElement();
                        } //if bs6Value
                        if(bt6Value != undefined){xw.startElement("forma_pago").text(bt6Value).endElement();}
                        if(bu6Value != undefined){xw.startElement("instrumento_monetario").text(bu6Value).endElement();}
                        if(bv6Value != undefined){xw.startElement("moneda").text(bv6Value).endElement();}
                        if(bw6Value != undefined){xw.startElement("monto_operacion").text(bw6Value.toFixed(2)).endElement();}
                      xw.endElement(); //datos liquidación
                    } //if bt6Value
                  xw.endElement(); //datos operacion
                xw.endElement(); //detalle operaciones
              } //if bg6Value
            xw.endElement();} //aviso
            /************************************** FIN DE FILA **************************************/
            /************************************** INICIO DE FILA 3 **************************************/
            if(a7Value != undefined){ xw.startElement("aviso");
              if(c7Value != undefined){xw.startElement("referencia_aviso").text(c7Value).endElement();}
              if(e7Value != undefined){xw.startElement("prioridad").text(e7Value).endElement();}
              if(d7Value != undefined){
                xw.startElement("alerta");
                  xw.startElement("tipo_alerta").text(d7Value).endElement();
                  if(f7Value != undefined){xw.startElement("descripcion_alerta").text(f7Value.trim()).endElement();}
                xw.endElement(); //alerta
              } //if d7Value
              xw.startElement("persona_aviso");
                xw.startElement("tipo_persona");
                  if(g7Value != undefined){
                    xw.startElement("persona_física");
                      xw.startElement("nombre").text(g7Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();
                      if(h7Value != undefined){xw.startElement("apellido_paterno").text(h7Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                      if(i7Value != undefined){xw.startElement("apellido_materno").text(i7Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                      if(j7Value != undefined){
                        let j7Month = j7Value.getMonth(); j7Month++; if (j7Month < 10) { j7Month = "0" + j7Month; }
                        let j7Day = j7Value.getDate(); if (j7Day < 10) { j7Day = "0" + j7Day; }
                        xw.startElement("fecha_nacimiento").text(j7Value.getFullYear() + "" + j7Month + "" + j7Day).endElement();
                      } //ifj7Value
                      if(k7Value != undefined){xw.startElement("rfc").text(k7Value.toUpperCase().trim()).endElement();}
                      if(l7Value != undefined){xw.startElement("curp").text(l7Value.toUpperCase().trim()).endElement();}
                      if(m7Value != undefined){xw.startElement("pais_nacionalidad").text(m7Value.toUpperCase().trim()).endElement();}
                      if(n7Value != undefined){xw.startElement("actividad_economica").text(n7Value).endElement();}
                    xw.endElement(); //persona física
                  } /*if g7Value*/ else if(o7Value != undefined){
                    xw.startElement("persona_moral");
                      if(o7Value != undefined){xw.startElement("denominacion_razon").text(o7Value.toUpperCase().trim()).endElement();}
                      if(p7Value != undefined){
                        let p7Month = p7Value.getMonth(); p7Month++; if (p7Month < 10) { p7Month = "0" + p7Month; }
                        let p7Day = p7Value.getDate(); if (p7Day < 10) { p7Day = "0" + p7Day; }
                        xw.startElement("fecha_constitucion").text(p7Value.getFullYear() + "" + p7Month + "" + p7Day).endElement();
                      } //if p7Value
                      if(q7Value != undefined){xw.startElement("rfc").text(q7Value.toUpperCase().trim()).endElement();}
                      if(r7Value != undefined){xw.startElement("pais_nacionalidad").text(r7Value.toUpperCase().trim()).endElement();}
                      if(s7Value != undefined){xw.startElement("giro_mercanril").text(s7Value).endElement();}
                      xw.startElement("representante_apoderado");
                        if(w7Value != undefined){xw.startElement("nombre").text(w7Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(x7Value != undefined){xw.startElement("apellido_paterno").text(x7Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(y7Value != undefined){xw.startElement("apellido_materno").text(y7Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(z7Value != undefined){
                          let z7Month = z7Value.getMonth(); z7Month++; if (z7Month < 10) { z7Month = "0" + z7Month; }
                          let z7Day = z7Value.getDate(); if (z7Day < 10) { z7Day = "0" + z7Day; }
                          xw.startElement("fecha_nacimiento").text(z7Value.getFullYear() + "" + z7Month + "" + z7Day).endElement();
                        } //if z7Value
                        if(aa7Value != undefined){xw.startElement("rfc").text(aa7Value.toUpperCase().trim()).endElement();}
                        if(ab7Value != undefined){xw.startElement("curp").text(ab7Value.toUpperCase().trim()).endElement();}
                      xw.endElement(); //representante apoderado persona moral
                    xw.endElement(); //persona moral
                  } /*if o7Value*/ else if(t7Value != undefined){
                    xw.startElement("fideicomiso");
                      if(t7Value != undefined){xw.startElement("denominacion_razon").text(t7Value.toUpperCase().trim()).endElement();}
                      if(u7Value != undefined){xw.startElement("rfc").text(u7Value.toUpperCase().trim()).endElement();}
                      if(v7Value != undefined){xw.startElement("identificador_fideicomiso").text(v7Value).endElement();}
                      xw.startElement("representante_apoderado");
                        if(w7Value != undefined){xw.startElement("nombre").text(w7Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(x7Value != undefined){xw.startElement("apellido_paterno").text(x7Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(y7Value != undefined){xw.startElement("apellido_materno").text(y7Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(z7Value != undefined){
                          let z7Month = z7Value.getMonth(); z7Month++; if (z7Month < 10) { z7Month = "0" + z7Month; }
                          let z7Day = z7Value.getDate(); if (z7Day < 10) { z7Day = "0" + z7Day; }
                          xw.startElement("fecha_nacimiento").text(z7Value.getFullYear() + "" + z7Month + "" + z7Day).endElement();
                        } //if z7Value
                        if(aa7Value != undefined){xw.startElement("rfc").text(aa7Value.toUpperCase().trim()).endElement();}
                        if(ab7Value != undefined){xw.startElement("curp").text(ab7Value.toUpperCase().trim()).endElement();}
                      xw.endElement(); //representante apoderado fideicomiso
                    xw.endElement(); //fideicomiso
                  }
                xw.endElement(); //tipo persona
                xw.startElement("tipo_domicilio");
                  if(ac7Value != undefined){
                    xw.startElement("nacional");
                      if(ad7Value != undefined){xw.startElement("colonia").text(ad7Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                      if(ae7Value != undefined){xw.startElement("calle").text(ae7Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                      if(af7Value != undefined){xw.startElement("numero_exterior").text(af7Value).endElement();}
                      if(ag7Value != undefined){xw.startElement("numero_interior").text(ag7Value).endElement();}
                      if(ac7Value != undefined){xw.startElement("codigo_postal").text(ac7Value).endElement();}
                    xw.endElement();
                  }else if(ah7Value != undefined) {
                    xw.startElement("extranjero");
                      if(ah7Value != undefined){xw.startElement("pais").text(ah7Value.toUpperCase()).endElement();}
                      if(ai7Value != undefined){xw.startElement("estado_provincia").text(ai7Value).endElement();}
                      if(aj7Value != undefined){xw.startElement("ciudad_poblacion").text(aj7Value).endElement();}
                      if(ak7Value != undefined){xw.startElement("colonia").text(ak7Value).endElement();}
                      if(al7Value != undefined){xw.startElement("calle").text(al7Value).endElement();}
                      if(am7Value != undefined){xw.startElement("numero_exterior").text(am7Value).endElement();}
                      if(an7Value != undefined){xw.startElement("numero_interior").text(an7Value).endElement();}
                      if(ao7Value != undefined){xw.startElement("codigo_postal").text(ao7Value).endElement();}
                    xw.endElement();
                  }
                xw.endElement(); //tipo domicilio
                if(ap7Value != undefined || aq7Value != undefined || ar7Value != undefined){
                  xw.startElement("telefono");
                    if(ap7Value != undefined){xw.startElement("clave_pais").text(ap7Value.toUpperCase().trim()).endElement();}
                    if(aq7Value != undefined){xw.startElement("numero_telefono").text(aq7Value.toString().replace("    ","").replace("  ","").replace(" ","")).endElement();}
                    if(ar7Value != undefined){xw.startElement("correo_electronico").text(ar7Value.toUpperCase().trim()).endElement();}
                  xw.endElement(); //telefono
                } //if ap7Value || aq7Value || ar7Value
              xw.endElement(); //persona aviso
              if(as7Value != undefined || az7Value != undefined || bd7Value != undefined){
                xw.startElement("dueno_beneficiario");
                  xw.startElement("tipo_persona");
                    if(as7Value != undefined){
                      xw.startElement("persona_fisica");
                        xw.startElement("nombre").text(as7Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();
                        if(at7Value != undefined){xw.startElement("apellido_paterno").text(at7Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(au7Value != undefined){xw.startElement("apellido_materno").text(au7Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                        if(av7Value != undefined){
                          let av7Month = av7Value.getMonth(); av7Month++; if (av7Month < 10) { av7Month = "0" + av7Month; }
                          let av7Day = av7Value.getDate(); if (av7Day < 10) { av7Day = "0" + av7Day; }
                          xw.startElement("fecha_nacimiento").text(av7Value.getFullYear() + "" + av7Month + "" + av7Day).endElement();
                        } //if av7Value
                        if(aw7Value != undefined){xw.startElement("rfc").text(aw7Value.toUpperCase().trim()).endElement();}
                        if(ax7Value != undefined){xw.startElement("curp").text(ax7Value.toUpperCase().trim()).endElement();}
                        if(ay7Value != undefined){xw.startElement("pais_nacionalidad").text(ay7Value.toUpperCase()).endElement();}
                      xw.endElement(); //persona física
                    } /*if as7Value*/ else if(az7Value != undefined){
                      xw.startElement("persona_moral");
                        xw.startElement("denominacion_razon").text(az7Value.toUpperCase().trim()).endElement();
                        if(ba7Value != undefined){
                          let ba7Month = ba7Value.getMonth(); ba7Month++; if (ba7Month < 10) { ba7Month = "0" + ba7Month; }
                          let ba7Day = ba7Value.getDate(); if (ba7Day < 10) { ba7Day = "0" + ba7Day; }
                          xw.startElement("fecha_constitucion").text(ba7Value.getFullYear() + "" + ba7Month + "" + ba7Day).endElement();
                        } //if ba7Value
                        if(bb7Value != undefined){xw.startElement("rfc").text(bb7Value.toUpperCase().trim()).endElement();}
                        if(bc7Value != undefined){xw.startElement("pais_nacionalidad").text(bc7Value.toUpperCase()).endElement();}
                      xw.endElement(); //persona moral
                    } /*if az7Value*/ else if(bd7Value != undefined){
                      xw.startElement("fideicomiso");
                        xw.startElement("denominacion_razon").text(bd7Value.toUpperCase().trim()).endElement();
                        if(be7Value != undefined){xw.startElement("rfc").text(be7Value.toUpperCase().trim()).endElement();}
                        if(bf7Value != undefined){xw.startElement("identificador_fideicomiso").text(bf7Value).endElement();}
                      xw.endElement(); //fideicomiso
                    } //if bd7Value
                  xw.endElement(); //tipo persona beneficiario
                xw.endElement(); //dueño beneficiario
              } //if as7Value || az7Value || bd7Value
              if(bg7Value != undefined){
                xw.startElement("detalle_operaciones");
                  xw.startElement("datos_operacion");
                    if(bg7Value != undefined){
                      let bg7Month = bg7Value.getMonth(); bg7Month++; if (bg7Month < 10) { bg7Month = "0" + bg7Month; }
                      let bg7Day = bg7Value.getDate(); if (bg7Day < 10) { bg7Day = "0" + bg7Day; }
                      xw.startElement("fecha_operacion").text(bg7Value.getFullYear() + "" + bg7Month + "" + bg7Day).endElement();
                    } //if bg7Value
                    if(bh7Value != undefined){xw.startElement("tipo_operacion").text(bh7Value).endElement();}
                    xw.startElement("caracteristicas");
                      if(bq7Value != undefined){
                        let bq7Month = bq7Value.getMonth(); bq7Month++; if (bq7Month < 10) { bq7Month = "0" + bq7Month; }
                        let bq7Day = bq7Value.getDate(); if (bq7Day < 10) { bq7Day = "0" + bq7Day; }
                        xw.startElement("fecha_inicio").text(bq7Value.getFullYear() + "" + bq7Month + "" + bq7Day).endElement();
                      } //if bq7Value
                      if(br7Value != undefined){
                        let br7Month = br7Value.getMonth(); br7Month++; if (br7Month < 10) { br7Month = "0" + br7Month; }
                        let br7Day = br7Value.getDate(); if (br7Day < 10) { br7Day = "0" + br7Day; }
                        xw.startElement("fecha_termino").text(br7Value.getFullYear() + "" + br7Month + "" + br7Day).endElement();
                      } //if br7Value
                      if(bi7Value != undefined){xw.startElement("tipo_inmueble").text(bi7Value).endElement();}
                      if(bj7Value != undefined){xw.startElement("valor_referencia").text(bj7Value.toFixed(2)).endElement();}
                      if(bl7Value != undefined){xw.startElement("colonia").text(bl7Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                      if(bm7Value != undefined){xw.startElement("calle").text(bm7Value.toUpperCase().replace("É", "E").replace("Á", "A").replace("Ó", "O").replace("Í", "I").replace("Ú", "U").trim()).endElement();}
                      if(bn7Value != undefined){xw.startElement("numero_exterior").text(bn7Value).endElement();}
                      if(bo7Value != undefined){xw.startElement("numero_interior").text(bo7Value).endElement();}
                      if(bk7Value != undefined){xw.startElement("codigo_postal").text(bk7Value).endElement();}
                      if(bp7Value != undefined){xw.startElement("folio_real").text(bp7Value).endElement();}
                    xw.endElement(); //caracteristicas
                    if(bt7Value != undefined){
                      xw.startElement("datos_liquidacion");
                        if(bs7Value != undefined){
                          let bs7Month = bs7Value.getMonth(); bs7Month++; if (bs7Month < 10) { bs7Month = "0" + bs7Month; }
                          let bs7Day = bs7Value.getDate(); if (bs7Day < 10) { bs7Day = "0" + bs7Day; }
                          xw.startElement("fecha_pago").text(bs7Value.getFullYear() + "" + bs7Month + "" + bs7Day).endElement();
                        } //if bs7Value
                        if(bt7Value != undefined){xw.startElement("forma_pago").text(bt7Value).endElement();}
                        if(bu7Value != undefined){xw.startElement("instrumento_monetario").text(bu7Value).endElement();}
                        if(bv7Value != undefined){xw.startElement("moneda").text(bv7Value).endElement();}
                        if(bw7Value != undefined){xw.startElement("monto_operacion").text(bw7Value.toFixed(2)).endElement();}
                      xw.endElement(); //datos liquidación
                    } //if bt7Value
                  xw.endElement(); //datos operacion
                xw.endElement(); //detalle operaciones
              } //if bg7Value
            xw.endElement();} //aviso
            /************************************** FIN DE FILA **************************************/


          xw.endElement(); //informe
        xw.endElement(); //archivo
        xw.endDocument();

        //res.redirect("/xmlView");
        // let xmlFile = xw.toString();
        // fs.writeFile("./uploads/arrendamiento.txt", xmlFile);
        //res.send(xw.toString());
        let wstream = fs.createWriteStream("./uploads/arrendamiento.xml");
        wstream.write(xw.toString(), (err)=>{ //xw.toString()
          if (err) {
            console.log(err);
          } else {
            console.log("data written succesfully");
            res.download("./uploads/arrendamiento.xml");
          }
        });
        // fs.unlink("./uploads/arrendamiento.xml", (err)=>{
        //   if (err) {
        //     console.log(err);
        //   } else {
        //     console.log("file deleted succesfully");
        //   }
        // });
      } //else
    });
  }
});
/*******************************************************************/

let port = process.env.PORT;
if (port == null || port == "") {
  port = 3000;
}

app.listen(port, function(){
  console.log("Server started succesfully");
});
