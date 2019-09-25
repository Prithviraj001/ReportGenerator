const {app,BrowserWindow,ipcMain,dialog}= require('electron');
const generateDocs = require('generate-docx');
const path=require('path')
const os=require('os')
let fs = require("fs");
global.sharedObj = {mail: false,wrdFile:null,xlFile:null, outputLoc:'./Output'};
let senderId='xyzs@'+'organization.com';

function createWindow() {
	let appWindow = new BrowserWindow({
		width:850,
		height:400,
		icon:'icon.ico',
		webPreferences: {
			nodeIntegration: true
		}
	});
	appWindow.setMenuBarVisibility(false);
	appWindow.loadFile("./index.html");
	try{
	fs.mkdirSync('./Output');
	}
	catch (e) {
    }
}

app.once("ready",function () {
	createWindow();
});

ipcMain.on('openDialog1',(event, args) => {
	dialog.showOpenDialog({ properties: [ 'openFile' ],filters:[{name:'Template File (.docx)',extensions:['docx']}]})
		.then(result=>{
			global.sharedObj.wrdFile=result.filePaths.toString();
			event.reply('wordReply');
		})

});

ipcMain.on('openDialog2',(event, args) => {
	dialog.showOpenDialog({ properties: [ 'openFile' ],filters:[{name:'Data File (.xlsx,.csv,.xls)',extensions:['xlsx','csv','xls']}]})
		.then(result=>{
			global.sharedObj.xlFile=result.filePaths.toString();
			event.reply('xlReply');
		})
});

ipcMain.on('saveDialog',(event, args)=> {
	let output;
	dialog.showOpenDialog({properties: ['openDirectory']}).then(result => {
		output = result.filePaths;
		global.sharedObj.outputLoc = output.toString();
		event.reply('saveLocation');
	});
});

ipcMain.on('generate',(event, args)=>{
	try {
		let wrdFile = global.sharedObj.wrdFile, xlFile = global.sharedObj.xlFile,outputLoc = global.sharedObj.outputLoc;
		var startTime = process.hrtime();
		let parser = new (require('simple-excel-to-json').XlsParser)();
		let doc = parser.parseXls2Json(xlFile);
		for (let i = doc[0].length-2; i >= 0; i--) {

			const options = {
				template: {
					filePath: wrdFile,
					data: doc[0][i]
				},
				save: {
					filePath:path.format({
						root: '/ignored',
						dir: outputLoc,
						base: doc[0][i].Fname + '_' + doc[0][i].Id + '.docx'
					})
						//outputLoc + path.deli + doc[0][i].Fname + '_' + doc[0][i].Id + '.docx'
				}
			};
			generateDocs(options).then().catch(console.error);
		}

		let report_msg = '<html><body> Template file: ' + wrdFile + '<br/> Excel File : ' + xlFile + ' <br/> Output Location : ' + outputLoc + '<br/> Generated on : ' + os.hostname() + '<br/> Generated Files : ' + doc[0].length + '</body> </html>';
		if (global.sharedObj.mail)
			notify(report_msg)

		return dialog.showMessageBox(new BrowserWindow({
			title:' ',
			show: false,
			alwaysOnTop: true,
			center: true
		}),{
		message:"Generated Successfully !!!",
		buttons:["OK"],
		defaultId: 0
		},function(index){
		})
	}

	catch (e) {
	}
});


function notify(report_msg) {
	let recpId=	fs.readFileSync('mail.txt', 'utf8', function(err, data) {
		if (err) throw err;
		return data;
	})
	const sendmail = require('sendmail')();
sendmail({
    from: senderId,
    to: recpId,
    subject: 'Generation Report',
    html: report_msg,
  }, function(	err, reply) {
    console.log(err && err.stack);
    console.dir(reply);
});

}
