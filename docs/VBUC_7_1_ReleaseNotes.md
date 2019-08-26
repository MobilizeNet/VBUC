# Visual Basic Upgrade Companion Version 7.1

# Release Notes

- The VBUC is now able to execute the upgrade step of the migration in parallel, taking advantage 
of modern multicore computers. To enable parallel migration, check the option on the Upgrade 
tab of the VBUC UI. When enabled, if the migration solution has multiple projects, the VBUC will 
divide the number of projects among threads that will run concurrently on separate processor 
cores, significantly reducing the migration time.  
Actual performance improvements will vary depending on the migration solution properties and 
the physical capacities of the machine itself.  Some tests conducted have yield results in which 
the upgrade step time has been reduced to 60% ‐ 40% of the original sequential (single core) time.   

- Several enhancements that improve the code readability and reduce the amount of manual 
work required after the automatic migration, by avoiding compilation errors or by achieving 
higher functional equivalence: 
  - a. Improved recognition and removal of unnecessary `By Ref` parameters. 
  - b. Fixed compilation errors related to the migration of parameterless events. 
  - c. Shared files migration has been improved to take into consideration some cases where identifiers can be resolved in different ways for different projects, particularly in the case of parameters resolution. 
  - d. Arrays are declared and initialized in different ways in VB6, VB.NET and C#.  The VBUC 7.1 introduces enhancements in the recognition of array dimensions, their declarations and initializations. 
  - e. Improved functional equivalence through the helper classes by covering additional cases of runtime behavior. 
  - f. A few bugs were fixed when referencing old COM/ActiveX libraries, particularly related to enums, `OLE_Color` properties, and specific cases of property get to methods. 
  - g. The recognition of late binding cases has been improved to reduce the usage of a helper class and the occurrence of related compilation errors. 
  - h. String concatenations with combinations of `+` operators and multiple operand types have been enhanced to cover all known cases achieving compilation and functional equivalence. 
  - i. Fixed cases of for loops that were wrongly migrated to VB.Net. 
  - j. Several improvements related to the conversion of the Cursor enum to a .Net class, and its corresponding coercions in various use cases. 
  - k. Several other fixes related to compilation, functional equivalence or code readability. 

- The following libraries are now converted in a better way to .Net corresponding libraries: 
  - a. `SSDataWidgets`: both design properties and code member references. 
  - b. `TrueDBGrid`: both design properties and code member references. 
  - c. Data Access migration (`ADO`, `DAO`, `RDO`) 
  - d. DataSource member from multiple controls. 
  - e. `CommonDialog` patterns. 
  - f. Fonts: Design and code member references. 
  - g. `SSPanel` to Label/PictureBox optional feature added to the Upgrade Options. 
  - h. `SSActiveTab` mappings added to the Upgrade Options. 
  - i. BorderStyle mappings were improved. 
  - j. Several other details where improved in the migration of common VB6 Libraries to .Net equivalent ones. 

- Reduced compile errors: Using real‐world code, the VBUC team further improved code generation which in turn reduced the number of compile errors needing attention. Following on the significant improvements in the last several releases, an additional 2.5M lines of VB6 code 
were used as a test bed, with post‐migration compile errors reduced from 33,000 to 2600, or about 1 fixup required per KLOC.  

- Included with VBUC 7.1 is a standalone assessment tool that can create metadata about your 
VB6 project for our consulting services engineers. It will provide details like number of lines of 
code, references, project names, etc. It does not collect source code.  