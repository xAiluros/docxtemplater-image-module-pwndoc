"use strict";

var DOMParser = require("xmldom").DOMParser;
var XMLSerializer = require("xmldom").XMLSerializer;

var DocUtils = {};

DocUtils.xml2Str = function (xmlNode) {
	var a = new XMLSerializer();
	return a.serializeToString(xmlNode);
};

DocUtils.str2xml = function (str, errorHandler) {
	var parser = new DOMParser({errorHandler});
	return parser.parseFromString(str, "text/xml");
};

DocUtils.maxArray = function (a) { return Math.max.apply(null, a); };

DocUtils.decodeUtf8 = function (s) {
	try {
		if (s === undefined) { return undefined; }
		// replace Ascii 160 space by the normal space, Ascii 32
		return decodeURIComponent(escape(DocUtils.convertSpaces(s)));
	}
	catch (e) {
		var err = new Error("End");
		err.properties.data = s;
		err.properties.explanation = "Could not decode string to UFT8";
		throw err;
	}
};

DocUtils.encodeUtf8 = function (s) {
	return unescape(encodeURIComponent(s));
};

DocUtils.convertSpaces = function (s) {
	return s.replace(new RegExp(String.fromCharCode(160), "g"), " ");
};

DocUtils.pregMatchAll = function (regex, content) {
	/* regex is a string, content is the content. It returns an array of all matches with their offset, for example:
	regex=la
	content=lolalolilala
	returns: [{0:'la',offset:2},{0:'la',offset:8},{0:'la',offset:10}]
	*/
	if (typeof regex !== "object") {
		regex = (new RegExp(regex, "g"));
	}
	var matchArray = [];
	var replacer = function (...pn) {
		pn.pop();
		var offset = pn.pop();
		// add match so that pn[0] = whole match, pn[1]= first parenthesis,...
		pn.offset = offset;
		return matchArray.push(pn);
	};
	content.replace(regex, replacer);
	return matchArray;
};

module.exports = DocUtils;
