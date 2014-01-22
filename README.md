openxls
=======

This is a copy of the LGPL shared openxls library.  It's no longer supported by the owners, but we're using it and are fixing things as we
encounter any problems.  If you find an issue, please open an issue on Github and/or provide a pull request.  For any issues logged, please
provide a small simple test case.  Small simple Excel files are also useful in dealing with the multitude of file formats and versions out
there.

**** Plan ****
Initial commits are going to be code cleanup - bringing the code base up to a more common JDK7 standard.  We've noticed several bugs while
doing this and will attempt to fix the ones we encounter as we go, while adding unit tests for the lib as we go.

A big issue for us right now is the corruption of the two TreeMaps held within Boundsheet - this is an ordered tree, and instance variables
that are part of the ordering are being modified in place.  This prevents the nodes from being found in subsequent searches and deletions,
and triggers errors in the other parts of the code base.