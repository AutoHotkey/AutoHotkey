// This script minimizes the size of AutoHotkey.exe.manifest.
// Using JS because it doesn't depend on AutoHotkey having already been compiled...

try
{
	var args = WScript.Arguments;
	if (args.length < 2)
		throw new Error("Too few arguments!");
	var source = args(0);
	var dest = args(1);

	var fs = new ActiveXObject("Scripting.FileSystemObject");

	var xml = fs.OpenTextFile(source).ReadAll();

	xml = xml
	.replace(/<!--[\s\S]*?-->/g, "")	// Remove comments
	.replace(/>\s*</g, "><")			// Remove space between elements
	.replace(/\r?\n\s*/g, " ")			// Replace line breaks+indent with one space
	.replace(/<\?xml.*?\?>/, "")		// The VS manifest tool seems to strip this out, so we will too

	fs.CreateTextFile(dest, true).Write(xml);

	//WScript.Echo("Finished processing manifest '"+dest+"'");
}
catch (ex)
{
	WScript.Echo("Error in minman.js: " + ex.message);
	WScript.Quit(1);
}