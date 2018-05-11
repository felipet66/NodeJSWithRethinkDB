var express = require('express'); 
var app = express(); 
var bodyParser = require('body-parser');
var multer = require('multer');
var xlstojson = require("xls-to-json-lc");
var xlsxtojson = require("xlsx-to-json-lc");
var r = require('rethinkdb');

app.use(bodyParser.json());  
var connection = null;
r.connect({ host: 'localhost', port: 28015 }, function(err, conn) {
  if(err) throw err;
  connection = conn;
    r.connect( {host: 'localhost', port: 28015}, function(err, conn) {
        if (err) throw err;
        connection = conn;
    })
});

var storage = multer.diskStorage({ //Configurações de armazenamento
    destination: function (req, file, cb) {
        cb(null, './uploads/')
    },
    filename: function (req, file, cb) {
        var datetimestamp = Date.now();
        cb(null, file.fieldname + '-' + datetimestamp + '.' + file.originalname.split('.')[file.originalname.split('.').length -1])
    }
});

var upload = multer({ //configurações do multer
                storage: storage,
                fileFilter : function(req, file, callback) { //Filtro arquivo
                    if (['xls', 'xlsx'].indexOf(file.originalname.split('.')[file.originalname.split('.').length-1]) === -1) {
                        return callback(new Error('Tipo de extensão incorreto!'));
                    }
                    callback(null, true);
                }
            }).single('file');

/** Caminho onde a API fará os UPLOADS */
app.post('/upload', function(req, res) {
    var exceltojson;
    upload(req,res,function(err){
        if(err){
                res.json({error_code:1,err_desc:err});
                return;
        }
        /** Multer trás informações do arquivo no objeto req.file */
        if(!req.file){
            res.json({error_code:1,err_desc:"Arquivo não inserido"});
            return;
        }
        /** Verifique a extensão do arquivo recebido
         *  e utilize o modulo aproproiado
         */
        if(req.file.originalname.split('.')[req.file.originalname.split('.').length-1] === 'xlsx'){
            exceltojson = xlsxtojson;
        } else {
            exceltojson = xlstojson;
        }
        console.log(req.file.path);
        try {
            exceltojson({
                input: req.file.path,
                output: null, //não precisamos do arquivo output.json
                lowerCaseHeaders:true
            }, function(err,result){
                if(err) {
                    return res.json({error_code:1,err_desc:err, data: null});
                }
                r.table('contacts').insert([
                     result 
                ]).run(connection, err);
                r.table('contacts');
                res.json({
                    error_code:0,err_desc:null, data: result
                });
            });
        } catch (e){
            res.json({error_code:1,err_desc:"Excel corrumpido"});
        }
    })
    
});

app.get('/',function(req,res){
    res.sendFile(__dirname + "/index.html");
});

app.listen('5000', function(){
    console.log('executando na porta 5000...');
});