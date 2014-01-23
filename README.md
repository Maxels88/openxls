openxls
=======

This is a copy of the LGPL shared openxls library.  It's no longer supported by the owners, but we're using it and are fixing things as we
encounter any problems.  If you find an issue, please open an issue on Github and/or provide a pull request.  For any issues logged, please
provide a small simple test case.  Small simple Excel files are also useful in dealing with the multitude of file formats and versions out
there.

## Plan
Initial commits are going to be code cleanup - bringing the code base up to a more common JDK7 level.  We've noticed several bugs while
doing this and will attempt to fix the ones we encounter as we go. The aim is to add unit tests as we fix these issues, but please bear
with us as the code base is many years old and has lots of fixes and workarounds for oddities encountered in the field.  We're also moving
the logging to SLF4J, and ensuring that all exceptions are logged at WARN or ERROR level (depending on severity).

A big issue for us right now is the corruption of the two TreeMaps held within Boundsheet - this is an ordered tree, and instance variables
that are part of the ordering are being modified in place.  This prevents the nodes from being found in subsequent searches and deletions.

## 2014-01-22
Loggin has been moved to SLF4J.  There will likely be many log lines that are being recorded at the wrong level (i.e. too much noise).
Please feel free to send a pull request with your level changes, or log an issue with the file name and line number, the current level and
what you think it should be.

## 2014-01-23
Tracking down some weird and wonderful initialization issue with Mulblank cells...
