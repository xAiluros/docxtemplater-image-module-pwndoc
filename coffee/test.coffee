fs=require('fs')
DocxGen=require('docxtemplater')
expect=require('chai').expect

fileNames=[
	'imageExample.docx',
	'imageLoopExample.docx',
	'imageInlineExample.docx',
	'imageHeaderFooterExample.docx',
	'qrExample.docx',
	'noImage.docx',
	'qrExample2.docx',
	'qrHeader.docx',
	'qrHeaderNoImage.docx',
]

opts=null
beforeEach ()->
	opts={}
	opts.getImage=(tagValue)->
		fs.readFileSync(tagValue,'binary')
	opts.getSize=(imgBuffer,tagValue)->
		[150,150]
	opts.centered = false

ImageModule=require('../js/index.js')

docX={}

loadFile=(name)->
	if fs.readFileSync? then return fs.readFileSync(__dirname+"/../examples/"+name,"binary")
	xhrDoc= new XMLHttpRequest()
	xhrDoc.open('GET',"../examples/"+name,false)
	if (xhrDoc.overrideMimeType)
		xhrDoc.overrideMimeType('text/plain; charset=x-user-defined')
	xhrDoc.send()
	xhrDoc.response

for name in fileNames
	content=loadFile(name)
	docX[name]=new DocxGen()
	docX[name].loadedContent=content

describe 'image adding with {% image} syntax', ()->
	it 'should work with one image',()->
		name='imageExample.docx'
		imageModule=new ImageModule(opts)
		docX[name].attachModule(imageModule)
		out=docX[name]
			.load(docX[name].loadedContent)
			.setData({image:'examples/image.png'})
			.render()

		zip=out.getZip()
		fs.writeFile("test7.docx",zip.generate({type:"nodebuffer"}))

		imageFile=zip.files['word/media/image_generated_1.png']
		expect(imageFile?, "No image file found").to.equal(true)
		expect(imageFile.asText().length).to.be.within(17417,17440)

		relsFile=zip.files['word/_rels/document.xml.rels']
		expect(relsFile?, "No rels file found").to.equal(true)
		relsFileContent=relsFile.asText()
		expect(relsFileContent).to.equal("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="hId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header0.xml"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image_generated_1.png"/></Relationships>""")

		documentFile=zip.files['word/document.xml']
		expect(documentFile?, "No document file found").to.equal(true)
		documentContent=documentFile.asText()


	it 'should work with centering',()->
		d=new DocxGen()
		name='imageExample.docx'
		opts.centered = true
		imageModule=new ImageModule(opts)
		d.attachModule(imageModule)
		out=d
			.load(docX[name].loadedContent)
			.setData({image:'examples/image.png'})
			.render()

		zip=out.getZip()
		fs.writeFile("test_center.docx",zip.generate({type:"nodebuffer"}))
		imageFile=zip.files['word/media/image_generated_1.png']
		expect(imageFile?).to.equal(true)
		expect(imageFile.asText().length).to.be.within(17417,17440)

		relsFile=zip.files['word/_rels/document.xml.rels']
		expect(relsFile?).to.equal(true)
		relsFileContent=relsFile.asText()
		expect(relsFileContent).to.equal("""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering\" Target=\"numbering.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/><Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes\" Target=\"footnotes.xml\"/><Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes\" Target=\"endnotes.xml\"/><Relationship Id=\"hId0\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header0.xml\"/><Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image_generated_1.png\"/></Relationships>""")

		documentFile=zip.files['word/document.xml']
		expect(documentFile?).to.equal(true)
		documentContent=documentFile.asText()

	it 'should work with loops',()->
		name='imageLoopExample.docx'

		opts.centered = true
		imageModule=new ImageModule(opts)
		docX[name].attachModule(imageModule)

		out=docX[name]
			.load(docX[name].loadedContent)
			.setData({images:['examples/image.png','examples/image2.png']})

		out
			.render()

		zip=out.getZip()

		imageFile=zip.files['word/media/image_generated_1.png']
		expect(imageFile?).to.equal(true)
		expect(imageFile.asText().length).to.be.within(17417,17440)

		imageFile2=zip.files['word/media/image_generated_2.png']
		expect(imageFile2?).to.equal(true)
		expect(imageFile2.asText().length).to.be.within(7177,7181)

		relsFile=zip.files['word/_rels/document.xml.rels']
		expect(relsFile?).to.equal(true)
		relsFileContent=relsFile.asText()
		expect(relsFileContent).to.equal("""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering\" Target=\"numbering.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/><Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes\" Target=\"footnotes.xml\"/><Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes\" Target=\"endnotes.xml\"/><Relationship Id=\"hId0\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header0.xml\"/><Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image_generated_1.png\"/><Relationship Id=\"rId7\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image_generated_2.png\"/></Relationships>""")

		documentFile=zip.files['word/document.xml']
		expect(documentFile?).to.equal(true)
		documentContent=documentFile.asText()

		buffer=zip.generate({type:"nodebuffer"})
		fs.writeFile("test_multi.docx",buffer)

	it 'should work with image in header/footer',()->
		name='imageHeaderFooterExample.docx'
		imageModule=new ImageModule(opts)
		docX[name].attachModule(imageModule)
		out=docX[name]
			.load(docX[name].loadedContent)
			.setData({image:'examples/image.png'})
			.render()

		zip=out.getZip()

		imageFile=zip.files['word/media/image_generated_1.png']
		expect(imageFile?).to.equal(true)
		expect(imageFile.asText().length).to.be.within(17417,17440)

		imageFile2=zip.files['word/media/image_generated_2.png']
		expect(imageFile2?).to.equal(true)
		expect(imageFile2.asText().length).to.be.within(17417,17440)

		relsFile=zip.files['word/_rels/document.xml.rels']
		expect(relsFile?).to.equal(true)
		relsFileContent=relsFile.asText()
		expect(relsFileContent).to.equal("""<?xml version="1.0" encoding="UTF-8"?>
			<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
			</Relationships>""")

		headerRelsFile=zip.files['word/_rels/header1.xml.rels']
		expect(headerRelsFile?).to.equal(true)
		headerRelsFileContent=headerRelsFile.asText()
		expect(headerRelsFileContent).to.equal("""<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
			<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image_generated_2.png"/></Relationships>""")

		footerRelsFile=zip.files['word/_rels/footer1.xml.rels']
		expect(footerRelsFile?).to.equal(true)
		footerRelsFileContent=footerRelsFile.asText()
		expect(footerRelsFileContent).to.equal("""<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
			<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image_generated_1.png"/></Relationships>""")

		documentFile=zip.files['word/document.xml']
		expect(documentFile?).to.equal(true)
		documentContent=documentFile.asText()
		expect(documentContent).to.equal("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<w:document xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body><w:p><w:pPr><w:pStyle w:val="Normal"/><w:rPr></w:rPr></w:pPr><w:r><w:rPr></w:rPr></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:footerReference w:type="default" r:id="rId3"/><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="1800" w:right="1800" w:header="720" w:top="2810" w:footer="1440" w:bottom="2003" w:gutter="0"/><w:pgNumType w:fmt="decimal"/><w:formProt w:val="false"/><w:textDirection w:val="lrTb"/><w:docGrid w:type="default" w:linePitch="249" w:charSpace="2047"/></w:sectPr></w:body></w:document>""")

		fs.writeFile("test_header_footer.docx",zip.generate({type:"nodebuffer"}))

describe 'qrcode replacing',->
	describe 'shoud work without loops',->

		it 'should work with simple',(done)->
			name='qrExample.docx'
			opts.qrCode = true
			imageModule=new ImageModule(opts)

			imageModule.finished=()->
				zip=docX[name].getZip()
				buffer=zip.generate({type:"nodebuffer"})
				fs.writeFileSync("test_qr.docx",buffer)
				images=zip.file(/media\/.*.png/)
				expect(images.length).to.equal(2)
				expect(images[0].asText().length).to.equal(826)
				expect(images[1].asText().length).to.be.within(17417,17440)
				done()

			docX[name]=docX[name]
				.load(docX[name].loadedContent)
				.setData({image:'examples/image'})

			docX[name].attachModule(imageModule)

			docX[name]
				.render()

	describe 'should work with two',->

		it 'should work',(done)->
			name='qrExample2.docx'

			opts.qrCode = true
			imageModule=new ImageModule(opts)

			imageModule.finished=()->
				zip=docX[name].getZip()
				buffer=zip.generate({type:"nodebuffer"})
				fs.writeFileSync("test_qr3.docx",buffer)
				images=zip.file(/media\/.*.png/)
				expect(images.length).to.equal(4)
				expect(images[0].asText().length).to.equal(859)
				expect(images[1].asText().length).to.equal(826)
				expect(images[2].asText().length).to.be.within(17417,17440)
				expect(images[3].asText().length).to.be.within(7177,7181)
				done()

			docX[name]=docX[name]
				.load(docX[name].loadedContent)
				.setData({image:'examples/image',image2:'examples/image2.png'})

			docX[name].attachModule(imageModule)

			docX[name]
				.render()

	describe 'should work qr in headers without extra images',->
		it 'should work in a header too',(done)->
			name='qrHeaderNoImage.docx'
			opts.qrCode=true
			imageModule=new ImageModule(opts)

			imageModule.finished=()->
				zip=docX[name].getZip()
				buffer=zip.generate({type:"nodebuffer"})
				fs.writeFile("test_qr_header_no_image.docx",buffer)
				images=zip.file(/media\/.*.png/)
				expect(images.length).to.equal(3)
				expect(images[0].asText().length).to.equal(826)
				expect(images[1].asText().length).to.be.within(12888, 12900)
				expect(images[2].asText().length).to.be.within(17417,17440)
				done()

			docX[name]=docX[name]
				.load(docX[name].loadedContent)
				.setData({image:'examples/image',image2:'examples/image2.png'})

			docX[name].attachModule(imageModule)

			docX[name]
				.render()

	describe 'should work qr in headers with extra images',->
		it 'should work in a header too',(done)->
			name='qrHeader.docx'

			opts.qrCode = true
			imageModule=new ImageModule(opts)

			imageModule.finished=()->
				zip=docX[name].getZip()
				buffer=zip.generate({type:"nodebuffer"})
				fs.writeFile("test_qr_header.docx",buffer)
				images=zip.file(/media\/.*.png/)
				expect(images.length).to.equal(3)
				expect(images[0].asText().length).to.equal(826)
				expect(images[1].asText().length).to.be.within(12888, 12900)
				expect(images[2].asText().length).to.be.within(17417,17440)
				done()

			docX[name]=docX[name]
				.load(docX[name].loadedContent)
				.setData({image:'examples/image',image2:'examples/image2.png'})

			docX[name].attachModule(imageModule)

			docX[name]
				.render()
