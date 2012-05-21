#PhoneReport

##Dependencies
PhoneReport is a desktop Windows application built entirely in HTML and JavaScript. It depends on the HTML Application interpreter (mshta.exe) with Internet Explorer 5.0+ and the windows scripting host (wscript.exe). The currently implemented importer & exporters depend on Microsoft Office Excel, but alternate import and export formats could be designed to avoid that dependency.

##Basic Functionality
PhoneReport is designed for tracking usage of organizational phones. It correlates a database of phone call information with a series of directories specifying who is assigned to a given phone number at any given time and what their position in the organizational hierarchy is and then generates reports based on a configuration file giving the rules for phone usage within the organization. Reports can be generated for arbitrary time periods, for all phone numbers or all users in the organization, for a specific phone or subset of phone numbers or for a specific user. Partial reports can also be generated for phone numbers and users outside of the organization by consolidating information on calls between internal numbers and the external number or numbers.

The directory and settings formats and the settings editor interface were designed for use in LDS missions. They could be easily modified to support more generic hierarchical organizational structures.

##Import/Export
PhoneReport makes use of its own text-based storage format for call data to provide a compact and standardized target for importing. Currently, only a single importer is implemented (for the UkrMTS monthly invoice format), so this mainly serves the purpose of reducing storage requirements compared to the raw Excel files.

There is no storage format for calculated report data; report data is held in memory and must immediately exported to some presentation format.