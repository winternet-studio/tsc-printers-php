TSC label printers
==================

Library for handling TSC label printers

!!! PROJECT ONLY JUST STARTED - NOT YET READY FOR USE !!!

- If images are used PHP's GD extension is required.
- If USB connection is used PHP's COM (com_dotnet) extension is required.
- If the unofficial web interface method is used PHP's curl extension is required, unless you implement your own HTTP client.

So far only the very basic features for making a label has been abstracted into clean methods, for example like writing scalable text, lines, boxes, and images.
Even these might be limited to just exactly our usage scenario, so you are welcome to do Pull Requests with enhancements.
You can always do custom stuff using the `customCommand` method or use the `callActiveX` method for your own COM object calls.

A TSC DA220 printer has been used in the development of this class.


Installation
------------

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
php composer.phar require --prefer-dist winternet-studio/tsc-printers-php "*"
```

or add

```
"winternet-studio/tsc-printers-php": "*"
```

to the require section of your `composer.json` file.


Usage
-----

See the documentation with full example in `LabelPrinting.php`.

See the source code to fully understand how to use this library.
