/* Escrito por Ramsés Mendoza García
   Fecha de inicio: 12/12/2017 
   Descipción: Conversor .xml a .xlsx */

var express = require('express');
var app = express();

var multer = require('multer');
var upload = multer({dest: './uploads/'});

var fs = require('fs');
var xml2js = require('xml2js'); // Transforma .xml a JSON para facilitar limpieza del código en JS con semiParser()
var parser = new xml2js.Parser();

var nombre = ""; // El nombre del archivo
var obj; // Almacena un un objeto JSON

app.use(express.static('public'));

app.get('/', function(req, res) {
  res.sendfile('./index.html');
});

app.post('/upload', upload.array('archivo', 1), function(req, res, next) {
    for(var x=0;x<req.files.length;x++) {
        //copiamos el archivo a la carpeta definitiva de fotos
       fs.createReadStream('./uploads/'+req.files[x].filename).pipe(fs.createWriteStream('./facturas/'+req.files[x].originalname)); 
       //borramos el archivo temporal creado
       nombre = req.files[x].originalname;
       fs.unlink('./uploads/'+req.files[x].filename); 
    }  
    setTimeout(function(){
      var pagina='<!doctype html><html><head></head><body>'+
               '<p>¡Conversión exitosa!</p>'+
               '<br><a href="/">Retornar</a><br/><br/><a href="/factura.xlsx">Descargar</a></body></html>';
      res.send(pagina); 
    },2000);
        
    setTimeout(function(){ // 
      fs.readFile('./facturas/'+nombre, function(err, data){
        parser.parseString(data, function(err, result){
          var res = JSON.stringify(result, null, 1); // Del objeto .JSON pasamos a una cadena
            var cadenajson = semiParser(res); // Limpia a la cadena 'res' para poder convertir a .xlsx

            obj = JSON.parse(cadenajson); //De la cadena de texto pasamos a un objeto JSON

            //fs.writeFile('genial.json', cadenajson);
    },2000);

    setTimeout(function(){ 
      var json2xls = require('json2xls');
      //var result = require('./genial.json');
      var xls = json2xls(obj,{});

      fs.writeFileSync('./public/factura.xlsx',xls,'binary');
    },2000);
        
      
      });
    });
});

// La función semiParser quita de un string JSON (res) los caracteres { } y [] 
// además de otros conjuntos de caracteres para facilitar su transformación a un archivo .xlsx

function semiParser(res){ 
  var ser = res.replace("\"cfdi:Comprobante\":","");
        ser = ser.replace(/"\$":/g,"");

        ser = ser.replace("\"cfdi:Emisor\":","");

        ser = ser.replace("\"cfdi:ExpedidoEn\":","");

        ser = ser.replace("\"cfdi:DomicilioFiscal\":","");

        ser = ser.replace("\"cfdi:RegimenFiscal\":","");

        ser = ser.replace("\"cfdi:Receptor\":","");

        ser = ser.replace("\"cfdi:Domicilio\":","");

        ser = ser.replace("\"cfdi:Conceptos\":","");

        ser = ser.replace("\"cfdi:Concepto\":","");

        ser = ser.replace("\"cfdi:Impuestos\":","");

        ser = ser.replace("\"cfdi:Traslados\":","");

        ser = ser.replace("\"cfdi:Traslado\":","");

        ser = ser.replace("\"cfdi:Complemento\":","");

        ser = ser.replace("\"cfdi:Traslados\":","");

        ser = ser.replace("\"tfd:TimbreFiscalDigital\":","");

        ser = ser.replace(/\n/g,"");
        ser = ser.replace("\n","");

        ser = ser.replace(/{/g,"");

        ser = ser.replace(/}/g,"");

        ser = ser.replace(/.[*+?^{}()|[\]\\]/g,"");

        ser = ser.replace("rfc","rfcEmisor");
        ser = ser.replace("nombre","nombreEmisor");
        ser = ser.replace("calle","calleEmisor");
        ser = ser.replace("noExterior","noExteriorEmisor");
        ser = ser.replace("colonia","coloniaEmisor");
        ser = ser.replace("municipio","municipioEmisor");
        ser = ser.replace("estado","estadoEmisor");
        ser = ser.replace("pais","estadoEmisor");
        ser = ser.replace("estado","paisEmisor");
        ser = ser.replace("codigoPostal","codigoPostalEmisor");

        var final = "{"+ser+"}";

        return final;
}

app.listen(8080);
console.log("Escuchando...");