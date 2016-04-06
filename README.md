# MSAccessUtilities
## AccessObjectExport.py
###Description
I had/have the need to capture after-the-fact code modifications in an Access database.  Specifically, an ADP file.  These are long deprecated and obsolete, but there is no out-of-the-box way of distilling plain textual files for the representative Access object.  Tables, forms, reports, etc. are held in binary form, and are very difficult to track via source control.  This utility allows you to export in a known and consistent way all of the internal Access objects so that you can then apply source control.

This uses COM, so of course, is only Windows-specific.
###Usage
```$python AccessObjectExport.py -file=blah.adp -exportpath=exports```

* -file (required): a valid Access file.  This is tested with Access 2010 ADP files.
* -exportpath (optional): A valid folder location where you want the exports to go to.  Defaults to 'export' in the current directory.

Will export the following:
* forms: '.form' extension
* modules: '.bas' extension
* macros: '.mac' extension
* report: '.report' extension

