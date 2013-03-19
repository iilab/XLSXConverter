XLSXConverter
-------------

This is a webpage that converts XLSX files into json form defintions for ODK Survey.
It is avaialbe online [here](http://uw-ictd.github.com/XLSConverter.js/) via this repo's gh-pages branch.

These are some potential uses for it:

- Doing xlsx conversion within the ODK Survey code enabling users to simply drag XLSX files onto the sdcard
without converting them through a separate application.
- Offline conversions (just download the zip and you can run it from your filesystem).
- The converter could be included in a web interface for uploading form defs to ODK Aggregate.

XLSXConverter uses the [js-xlsx library](https://github.com/Niggler/js-xlsx) which is fairly new,
and I have encountered a few bugs in it. However, it is very well supported,
all the bugs I've reported have been fixed in about a day.
[The js-xlsx demo page](http://niggler.github.com/js-xlsx/)
provides a Litmus test of whether a bug is occuring in XLSConverter or js-xlsx.
