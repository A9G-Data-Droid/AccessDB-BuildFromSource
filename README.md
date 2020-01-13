AccessDB-BuildFromSource
========================

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

  In my testing I have found that the size of the compiled database is smaller than the original database where the source was exported from, even after a compact\repair. At this time I do not know what the source of the size mismatch is. I am not sure if the missing data is needed or not. It will take some time and testing to see if anything is missing from the new DB.

# Testers wanted!

  I am looking for feedback from anyone who has ran this against their database exports. Is there anything missing? Did you find a bug?

