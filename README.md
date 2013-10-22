# I'll be Shelling out to Access

A very hacky little test project to demonstrate a Windows application that shells-out to the Access 2007 Runtime to open an `.accdb` project and hook into a button-click event.

## Warning(s)

* There are some hardcoded paths in the `F#` source.
* This assumes that you have only the [Access 2007 Runtime](http://www.microsoft.com/en-gb/download/details.aspx?id=4438) installed. Other versions of Access may leave registry information for COM to load types that are not version-aligned.
* The code will likely break if you do not click the Access security risk popup quickly enough... either that or you should add the `.accdb` file path to a trusted location in your registry.