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
        var first_sheetName = wb.SheetNames[0];
        var worksheet = wb.Sheets[first_sheetName];

        let a2 = "A2"; let a2Adress = worksheet[a2]; let a2Value = (a2Adress ? a2Adress.v : undefined);
        let b2 = "B2"; let b2Adress = worksheet[b2]; let b2Value = (b2Adress ? b2Adress.v : undefined);
        let c2 = "C2"; let c2Adress = worksheet[c2]; let c2Value = (c2Adress ? c2Adress.v : undefined);
        let d2 = "D2"; let d2Adress = worksheet[d2]; let d2Value = (d2Adress ? d2Adress.v : undefined);
        let e2 = "E2"; let e2Adress = worksheet[e2]; let e2Value = (e2Adress ? e2Adress.v : undefined);
        let f2 = "F2"; let f2Adress = worksheet[f2]; let f2Value = (f2Adress ? f2Adress.v : undefined);
        let g2 = "G2"; let g2Adress = worksheet[g2]; let g2Value = (g2Adress ? g2Adress.v : undefined);
        /*************************** FALTAN ***************************/
        let o2 = "O2"; let o2Adress = worksheet[o2]; let o2Value = (o2Adress ? o2Adress.v : undefined);
        let p2 = "P2"; let p2Adress = worksheet[p2]; let p2Value = (p2Adress ? p2Adress.v : undefined);
        const p2Date = new Date((p2Value - (25567 + 2)) * 86400 * 1000);
        let p2Year = p2Date.getFullYear();
        let p2Month = (p2Date.getMonth())+1;
        if (p2Month < 10) { p2Month = "0" + p2Month; }
        let p2Day = (p2Date.getDate())+1;
        let q2 = "Q2"; let q2Adress = worksheet[q2]; let q2Value = (q2Adress ? q2Adress.v : undefined);
        let r2 = "R2"; let r2Adress = worksheet[r2]; let r2Value = (r2Adress ? r2Adress.v : undefined);
        let s2 = "S2"; let s2Adress = worksheet[s2]; let s2Value = (s2Adress ? s2Adress.v : undefined);
        let w2 = "W2"; let w2Adress = worksheet[w2]; let w2Value = (w2Adress ? w2Adress.v : undefined);
        let x2 = "X2"; let x2Adress = worksheet[x2]; let x2Value = (x2Adress ? x2Adress.v : undefined);
        let y2 = "Y2"; let y2Adress = worksheet[y2]; let y2Value = (y2Adress ? y2Adress.v : undefined);
        let z2 = "Z2"; let z2Adress = worksheet[z2]; let z2Value = (z2Adress ? z2Adress.v : undefined);
        const z2Date = new Date((z2Value - (25567 + 2)) * 86400 * 1000);
        let z2Year = z2Date.getFullYear(); let z2Month = (z2Date.getMonth())+1;
        if (z2Month < 10) { z2Month = "0" + z2Month; }
        let z2Day = (z2Date.getDate())+1;
        let aa2 = "AA2"; let aa2Adress = worksheet[aa2]; let aa2Value = (aa2Adress ? aa2Adress.v : undefined);
        let ab2 = "AB2"; let ab2Adress = worksheet[ab2]; let ab2Value = (ab2Adress ? ab2Adress.v : undefined);
        let ac2 = "AC2"; let ac2Adress = worksheet[ac2]; let ac2Value = (ac2Adress ? ac2Adress.v : undefined);
        let ad2 = "AD2"; let ad2Adress = worksheet[ad2]; let ad2Value = (ad2Adress ? ad2Adress.v : undefined);
        let ae2 = "AE2"; let ae2Adress = worksheet[ae2]; let ae2Value = (ae2Adress ? ae2Adress.v : undefined);
        let af2 = "AF2"; let af2Adress = worksheet[af2]; let af2Value = (af2Adress ? af2Adress.v : undefined);
        let ag2 = "AG2"; let ag2Adress = worksheet[ag2]; let ag2Value = (ag2Adress ? ag2Adress.v : undefined);
        /*************************** FALTAN ***************************/
        let ap2 = "AP2"; let ap2Adress = worksheet[ap2]; let ap2Value = (ap2Adress ? ap2Adress.v : undefined);
        let aq2 = "AQ2"; let aq2Adress = worksheet[aq2]; let aq2Value = (aq2Adress ? aq2Adress.v : undefined);
        let ar2 = "AR2"; let ar2Adress = worksheet[ar2]; let ar2Value = (ar2Adress ? ar2Adress.v : undefined);
        /*************************** FALTAN ***************************/
        let bg2 = "BG2"; let bg2Adress = worksheet[bg2]; let bg2Value = (bg2Adress ? bg2Adress.v : undefined);
        let bg2Month = bg2Value.getMonth(); bg2Month++; if (bg2Month < 10) { bg2Month = "0" + bg2Month; }
        let bg2Day = bg2Value.getDate(); if (bg2Day < 10) { bg2Day = "0" + bg2Day; }
        let bh2 = "BH2"; let bh2Adress = worksheet[bh2]; let bh2Value = (bh2Adress ? bh2Adress.v : undefined);
        let bi2 = "BI2"; let bi2Adress = worksheet[bi2]; let bi2Value = (bi2Adress ? bi2Adress.v : undefined);
        let bj2 = "BJ2"; let bj2Adress = worksheet[bj2]; let bj2Value = (bj2Adress ? bj2Adress.v : undefined);
        let bk2 = "BK2"; let bk2Adress = worksheet[bk2]; let bk2Value = (bk2Adress ? bk2Adress.v : undefined);
        let bl2 = "BL2"; let bl2Adress = worksheet[bl2]; let bl2Value = (bl2Adress ? bl2Adress.v : undefined);
        let bm2 = "BM2"; let bm2Adress = worksheet[bm2]; let bm2Value = (bm2Adress ? bm2Adress.v : undefined);
        let bn2 = "BN2"; let bn2Adress = worksheet[bn2]; let bn2Value = (bn2Adress ? bn2Adress.v : undefined);
        let bo2 = "BO2"; let bo2Adress = worksheet[bo2]; let bo2Value = (bo2Adress ? bo2Adress.v : undefined);
        let bp2 = "BP2"; let bp2Adress = worksheet[bp2]; let bp2Value = (bp2Adress ? bp2Adress.v : undefined);
        let bq2 = "BQ2"; let bq2Adress = worksheet[bq2]; let bq2Value = (bq2Adress ? bq2Adress.v : undefined);
        let bq2Month = bq2Value.getMonth(); bq2Month++; if (bq2Month < 10) { bq2Month = "0" + bq2Month; }
        let bq2Day = bq2Value.getDate(); if (bq2Day < 10) { bq2Day = "0" + bq2Day; }
        let br2 = "BR2"; let br2Adress = worksheet[br2]; let br2Value = (br2Adress ? br2Adress.v : undefined);
        let br2Month = br2Value.getMonth(); br2Month++; if (br2Month < 10) { br2Month = "0" + br2Month; }
        let br2Day = br2Value.getDate(); if (br2Day < 10) { br2Day = "0" + br2Day; }
        let bs2 = "BS2"; let bs2Adress = worksheet[bs2]; let bs2Value = (bs2Adress ? bs2Adress.v : undefined);
        let bs2Month = bs2Value.getMonth(); bs2Month++; if (bs2Month < 10) { bs2Month = "0" + bs2Month; }
        let bs2Day = bs2Value.getDate(); if (bs2Day < 10) { bs2Day = "0" + bs2Day; }
        let bt2 = "BT2"; let bt2Adress = worksheet[bt2]; let bt2Value = (bt2Adress ? bt2Adress.v : undefined);
        let bu2 = "BU2"; let bu2Adress = worksheet[bu2]; let bu2Value = (bu2Adress ? bu2Adress.v : undefined);
        let bv2 = "BV2"; let bv2Adress = worksheet[bv2]; let bv2Value = (bv2Adress ? bv2Adress.v : undefined);
        let bw2 = "BW2"; let bw2Adress = worksheet[bw2]; let bw2Value = (bw2Adress ? bw2Adress.v : undefined);

        //xw.startElement(""); xw.endElement();
        xw = new XMLWriter(true);
        xw.startDocument("1.0", "UTF-8");
        xw.startElement("archivo").writeAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
          xw.startElement("informe");
            xw.startElement("aviso");
              if (a2Value != undefined) {
                xw.startElement("rfc").text(a2Value).endElement();
                xw.startElement("periodo").text(b2Value).endElement();
                xw.startElement("referencia").text(c2Value).endElement();
                xw.startElement("alerta");
                  xw.startElement("tipo_alerta").text(d2Value).endElement();
                xw.endElement();
                xw.startElement("prioridad").text(e2Value).endElement();
                  if (f2Value != undefined) {
                    xw.startElement("descripcion_alerta").text(f2Value).endElement();
                  }
                xw.startElement("persona_aviso");
                xw.startElement("tipo_persona");
                  if (g2Value === undefined) {
                    xw.startElement("persona_moral");
                      xw.startElement("denominacion_razon").text(o2Value.toUpperCase()).endElement();
                      xw.startElement("fecha_constitución").text(p2Year+""+p2Month+""+p2Day).endElement();
                      xw.startElement("rfc").text(q2Value).endElement();
                      xw.startElement("pais_nacionalidad").text(r2Value).endElement();
                      xw.startElement("giro_mercanril").text(s2Value).endElement();
                      xw.startElement("representante_apoderado");
                        xw.startElement("nombre").text(w2Value.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u").toUpperCase().trim()).endElement();
                        xw.startElement("apellido_paterno").text(x2Value.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u").toUpperCase().trim()).endElement();
                        xw.startElement("apellido_materno").text(y2Value.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u").toUpperCase().trim()).endElement();
                        xw.startElement("fecha_nacimiento").text(z2Year+""+z2Month+""+z2Day).endElement();
                        xw.startElement("rfc").text(aa2Value).endElement();
                        xw.startElement("curp").text(ab2Value).endElement();
                      xw.endElement();
                    xw.endElement();
                  } else {
                    console.log("falta proceso de person física y fiduciario");
                  }
                xw.endElement();
                xw.startElement("tipo_domicilio");
                  if (ac2Value != undefined) {
                    xw.startElement("nacional");
                      xw.startElement("codigo_postal").text(ac2Value).endElement();
                      xw.startElement("colonia").text(ad2Value.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u").toUpperCase().trim()).endElement();
                      xw.startElement("calle").text(ae2Value.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u").toUpperCase().trim()).endElement();
                      xw.startElement("numero_interior").text(af2Value).endElement();
                      xw.startElement("numero_exterior").text(ag2Value).endElement();
                    xw.endElement();
                  }
                xw.endElement();
                xw.startElement("telefono");
                  xw.startElement("clave_pais").text(ap2Value.toUpperCase()).endElement();
                  xw.startElement("numero_telefono").text(aq2Value).endElement();
                  xw.startElement("correo_electronico").text(ar2Value.toUpperCase()).endElement();
                xw.endElement();
              xw.endElement();
              xw.startElement("detalle_operaciones");
                xw.startElement("datos_operacion");
                  xw.startElement("fecha_operacion").text(bg2Value.getFullYear() + "" + bg2Month + "" + bg2Day).endElement();
                  xw.startElement("tipo_operacion").text(bh2Value).endElement();
                  xw.startElement("características");
                    xw.startElement("tipo_inmueble").text(bi2Value).endElement();
                    xw.startElement("valor_referencia").text(bj2Value.toFixed(2)).endElement();
                    xw.startElement("codigo_postal").text(bk2Value).endElement();
                    xw.startElement("colonia").text(bl2Value.toUpperCase().trim()).endElement();
                    xw.startElement("calle").text(bm2Value.toUpperCase().trim()).endElement();
                    xw.startElement("numero_exterior").text(bn2Value).endElement();
                    xw.startElement("numero_interior").text(bo2Value).endElement();
                    xw.startElement("folio_real").text(bp2Value).endElement();
                    xw.startElement("fecha_inicio").text(bq2Value.getFullYear() + "" + bq2Month + "" + bq2Day).endElement();
                    xw.startElement("fecha_termino").text(br2Value.getFullYear() + "" + br2Month + "" + br2Day).endElement();
                  xw.endElement();
                  xw.startElement("datos_liquidacion");
                    xw.startElement("fecha_pago").text(bs2Value.getFullYear() + "" + bs2Month + "" + bs2Day).endElement();
                    xw.startElement("forma_pago").text(bt2Value).endElement();
                    xw.startElement("instrumento_monetario").text(bu2Value).endElement();
                    xw.startElement("moneda").text(bv2Value).endElement();
                    xw.startElement("mont_operacion").text(bw2Value.toFixed(2)).endElement();
                  xw.endElement();
                xw.endElement();
              xw.endElement();
              }
            xw.endElement(); //aviso
          xw.endElement(); //informe
        xw.endElement(); //archivo
        xw.endDocument();

        //res.redirect("/xmlView");
        // let xmlFile = xw.toString();
        // fs.writeFile("./uploads/arrendamiento.txt", xmlFile);
        //res.send(xw.toString());
        let wstream = fs.createWriteStream("./uploads/arrendamiento.xml");
        wstream.write(xw.toString(), (err)=>{
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
