var express = require('express');
var router = express.Router();
var formidable = require('formidable');
var xlsx = require('node-xlsx');
var fs =  require('fs');
var request = require('request');
/* GET home page. */
router.get('/', function(req, res, next) {
  res.render('index', { title: 'Express' });
});

router.post("/upload",function(req,res){  
    var form = new formidable.IncomingForm();   //创建上传表单 
    form.encoding = 'utf-8';        //设置编辑 
    form.uploadDir = './public';     //设置上传目录 文件会自动保存在这里 
    form.keepExtensions = true;     //保留后缀 
    form.maxFieldsSize = 5 * 1024 * 1024 ;   //文件大小5M 
    form.parse(req, function (err, fields, files) { 
        if(err){ 
            console.log(err); 
            return;
        }
        let rsjson = JSON.parse(JSON.stringify(files));
        console.log(rsjson)
        let path = rsjson.avatar.path;
        var fileExt = path.substring(path.lastIndexOf('.'));
        console.log(fileExt);
        console.log("&77777777777");
        console.log(fileExt.indexOf(".xlsx"));
        if(fileExt.indexOf(".xlsx")===0){
        let rsjson = JSON.parse(JSON.stringify(files));
        let path = rsjson.avatar.path;
        let obj = xlsx.parse(path);
        let excelObj  = obj[0].data;
        writeFile(rsjson.avatar.name+".json",JSON.stringify(obj));
        function writeFile(fileName,data)
        {  
          fs.writeFile(fileName,data,'utf-8',complete);
          function complete(err)
          {
              if(!err)
              {
                res.download(rsjson.avatar.name+".json");
              }   
          } 
        }
        }
        if(fileExt.indexOf(".xlsx")===-1){
          req.session.notice = {type:'error', message:'操作失败，请导入excel文件'};
          res.redirect('/devices');
        }
    }) 
});  
module.exports = router;
