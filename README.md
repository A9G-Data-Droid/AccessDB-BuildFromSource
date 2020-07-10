AccessDB-BuildFromSource
========================

**NOTE: The upstream project now includes the capability to compile the code exported back in to a new Access file. Please use this instead: https://github.com/joyfullservice/msaccess-vcs-integration 

# About

 Builds a new AccessDB from source files exported from msaccess-vcs-integration. 
 I am working with code from both the Add-In: https://github.com/joyfullservice/msaccess-vcs-integration
 
 And the upstream: https://github.com/timabell/msaccess-vcs-integration
 
Currently imports:
 
* Queries
* Forms
* Reports
* Macros
* Modules
* Table Data
* Table Definitions
* Table Data Macros
* Database Properties

# WARNING:

  In my testing I have found that the size of the compiled database is smaller than the original database where the source was exported from, even after a compact\repair. There are some things that are known to not survive the round trip export-compile. We are working to resolve each of those in the upstream project. This version is for historical purposes. I made this just to show that it could be done.
