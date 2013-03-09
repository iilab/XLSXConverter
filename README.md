XLSConverter.js
---------------

This is a webpage that converts XLSX files into json form defintions for ODK Survey.
It is avaialbe online [here] via this repo's gh-pages branch.

These are some potential uses for it:

- xlsx conversion within the ODK Survey code enabling users to simply drag xlsx files onto the sdcard
without converting them through a separate application.
- Offline conversions (the page can be used from the local file system).
- The converter could be included in a web interface for uploading form defs to ODK Aggregate.

The python XLSConverter may still be useful for previewing forms in Chrome, and XLS conversion.
However, with some work, it should be possible to add these features to the Javascript XLSConverter.
