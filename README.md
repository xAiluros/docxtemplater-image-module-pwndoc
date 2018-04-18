Open source docxtemplater image module
==========================================
This repository holds a updated version of docxtemplater image module.

This package is open source. There is also a [paid version](https://docxtemplater.com/modules/image/) maintained by docxtemplater author.

Note this version is compatible with docxtemplater 3.x.

Installation
=============
You first need to install docxtemplater by following its [installation guide](https://docxtemplater.readthedocs.io/en/latest/installation.html#node).

For Node.js install this package
```
npm install open-docxtemplater-image-module
```

For the browser find builds in `build/` directory.

Usage
=====
To render an image, your **docx** or **pptx** template should contain the text: `{%image}`

```javascript
//Node.js example
var ImageModule = require('open-docxtemplater-image-module');

//Below the options that will be passed to ImageModule instance
var opts = {}
opts.centered = false; //Set to true to always center images
opts.fileType = "docx"; //Or pptx

//Pass your image loader
opts.getImage = function(tagValue, tagName) {
    //tagValue is 'examples/image.png'
    //tagName is 'image'
    return fs.readFileSync(tagValue);
}

//Pass the function that return image size
opts.getSize = function(img, tagValue, tagName) {
    //img is the image returned by opts.getImage()
    //tagValue is 'examples/image.png'
    //tagName is 'image'
    //tip: you can use node module 'image-size' here
    return [150, 150];
}

var imageModule = new ImageModule(opts);

var zip = new JSZip(content);
var doc = new Docxtemplater()
    .attachModule(imageModule)
    .loadZip(zip)
    .setData({image: 'examples/image.png'})
    .render();

var buffer = doc
        .getZip()
        .generate({type:"nodebuffer"});

fs.writeFile("test.docx",buffer);
```

Some notes regarding the template:
* **docx** files: the placeholder `{%image}` must be in a dedicated paragraph.
* **pptx** files: the placeholder `{%image}` must be in a dedicated text cell.

Centering images
================
You can center all images using either passing the option `opts.centered = true` or one by one using `{%%image}` instead of `{%image}` in your templates.

In **pptx** generated documents, images are centered vertically and horizontally relative to the parent cell.
