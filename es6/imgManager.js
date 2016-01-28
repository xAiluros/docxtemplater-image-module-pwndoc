"use strict";

var DocUtils = require("./docUtils");

var imageExtensions = ["gif", "jpeg", "jpg", "emf", "png"];

module.exports = class ImgManager {
	constructor(zip, fileName) {
		this.zip = zip;
		this.fileName = fileName;
		this.endFileName = this.fileName.replace(/^.*?([a-z0-9]+)\.xml$/, "$1");
	}
	getImageList() {
		var regex = /		[^.]+		\.		([^.]+)		/;
		var imageList = [];
		Object.keys(this.zip.files).forEach(function (path) {
			var extension = path.replace(regex, "$1");
			if (imageExtensions.indexOf(extension) >= 0) {
				imageList.push({path: path, files: this.zip.files[path]});
			}
		});
		return imageList;
	}
	setImage(fileName, data, options = {}) {
		this.zip.remove(fileName);
		return this.zip.file(fileName, data, options);
	}
	hasImage(fileName) {
		return this.zip.files[fileName] != null;
	}
	loadImageRels() {
		var file = this.zip.files[`word/_rels/${this.endFileName}.xml.rels`] || this.zip.files["word/_rels/document.xml.rels"];
		if (file === undefined) { return; }
		var content = DocUtils.decodeUtf8(file.asText());
		this.xmlDoc = DocUtils.str2xml(content);
		// Get all Rids
		var RidArray = [];
		var iterable = this.xmlDoc.getElementsByTagName("Relationship");
		for (var i = 0, tag; i < iterable.length; i++) {
			tag = iterable[i];
			RidArray.push(parseInt(tag.getAttribute("Id").substr(3), 10));
		}
		this.maxRid = DocUtils.maxArray(RidArray);
		this.imageRels = [];
		return this;
	}
// Add an extension type in the [Content_Types.xml], is used if for example you want word to be able to read png files (for every extension you add you need a contentType)
	addExtensionRels(contentType, extension) {
		var content = this.zip.files["[Content_Types].xml"].asText();
		var xmlDoc = DocUtils.str2xml(content);
		var addTag = true;
		var defaultTags = xmlDoc.getElementsByTagName("Default");
		for (var i = 0, tag; i < defaultTags.length; i++) {
			tag = defaultTags[i];
			if (tag.getAttribute("Extension") === extension) { addTag = false; }
		}
		if (addTag) {
			var types = xmlDoc.getElementsByTagName("Types")[0];
			var newTag = xmlDoc.createElement("Default");
			newTag.namespaceURI = null;
			newTag.setAttribute("ContentType", contentType);
			newTag.setAttribute("Extension", extension);
			types.appendChild(newTag);
			return this.setImage("[Content_Types].xml", DocUtils.encodeUtf8(DocUtils.xml2Str(xmlDoc)));
		}
	}
	// Adding an image and returns it's Rid
	addImageRels(imageName, imageData, i = 0) {
		var realImageName = i === 0 ? imageName : imageName + `(${i})`;
		if ((this.zip.files[`word/media/${realImageName}`] != null)) {
			return this.addImageRels(imageName, imageData, i + 1);
		}
		this.maxRid++;
		var file = {
			name: `word/media/${realImageName}`,
			data: imageData,
			options: {
				base64: false,
				binary: true,
				compression: null,
				date: new Date(),
				dir: false,
			},
		};
		this.zip.file(file.name, file.data, file.options);
		var extension = realImageName.replace(/[^.]+\.([^.]+)/, "$1");
		this.addExtensionRels(`image/${extension}`, extension);
		var relationships = this.xmlDoc.getElementsByTagName("Relationships")[0];
		var newTag = this.xmlDoc.createElement("Relationship");
		newTag.namespaceURI = null;
		newTag.setAttribute("Id", `rId${this.maxRid}`);
		newTag.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
		newTag.setAttribute("Target", `media/${realImageName}`);
		relationships.appendChild(newTag);
		this.setImage(`word/_rels/${this.endFileName}.xml.rels`, DocUtils.encodeUtf8(DocUtils.xml2Str(this.xmlDoc)));
		return this.maxRid;
	}
	getImageName(id = 0) {
		var nameCandidate = "Copie_" + id + ".png";
		var fullPath = this.getFullPath(nameCandidate);
		if (this.hasImage(fullPath)) {
			return this.getImageName(id + 1);
		}
		return nameCandidate;
	}
	getFullPath(imgName) { return `word/media/${imgName}`; }
	// This is to get an image by it's rId (returns null if no img was found)
	getImageByRid(rId) {
		var relationships = this.xmlDoc.getElementsByTagName("Relationship");
		for (var i = 0, relationship; i < relationships.length; i++) {
			relationship = relationships[i];
			var cRId = relationship.getAttribute("Id");
			if (rId === cRId) {
				var path = relationship.getAttribute("Target");
				if (path.substr(0, 6) === "media/") {
					return this.zip.files[`word/${path}`];
				}
				throw new Error("Rid is not an image");
			}
		}
		throw new Error("No Media with this Rid found");
	}
};
