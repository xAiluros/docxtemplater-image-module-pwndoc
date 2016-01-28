XmlTemplater=require('docxtemplater').XmlTemplater

vm=require('vm')
JSZip=require('jszip')

QrCode=require('qrcode-reader')

module.exports= class DocxQrCode
	constructor:(imageData, @xmlTemplater,@imgName="",@num,@getDataFromString)->
		@callbacked=false
		@data=imageData
		if @data==undefined then throw new Error("data of qrcode can't be undefined")
		@ready=false
		@result=null
	decode:(@callback) ->
		_this= this
		@qr= new QrCode()
		@qr.callback= () ->
			_this.ready= true
			_this.result= @result
			testdoc= new XmlTemplater @result,
				fileTypeConfig:_this.xmlTemplater.fileTypeConfig
				tags:_this.xmlTemplater.tags
				Tags:_this.xmlTemplater.Tags
				parser:_this.xmlTemplater.parser
			testdoc.render()
			_this.result=testdoc.content
			_this.searchImage()
		@qr.decode({width:@data.width,height:@data.height},@data.decoded)
	searchImage:() ->
		cb=(err,@data=@data.data)=>
			if err then console.error err
			@callback(this,@imgName,@num)
		if !@result? then return cb()
		@getDataFromString(@result,cb)
