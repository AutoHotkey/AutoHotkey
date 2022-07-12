/*
This script exports a PNG file for each size of each icon from icons.svg.
It requires Inkscape to be somewhere on the PATH.
It does not currently automate conversion of the groups of PNG files to ICO.
*/

#Requires AutoHotkey v2.0-beta.6

; Using WScript.Shell & Exec didn't work out well...
cmds := FileOpen(tmpname := 'inkscape-commands.tmp', 'w')

for size in [16, 20, 24, 28, 32, 40, 48, 64] {
    export size, 'main'
}

for size in [16, 20, 24, 28, 32] {
    export size, 'p', 'bg, left, right'
    export size, 's', 'bg-h'
    export size, 'ps', 'bg-ii'
}

cmds.WriteLine 'quit-immediate'
cmds.Close()
cmds := ''
Run 'cmd /k "(type ' tmpname ' | inkscape icons.svg --shell) && del ' tmpname '"'

export(size, prefix, ids?) {
    if IsSet(ids)
        cmds.WriteLine Format(
            'unhide-all; select-clear; select-by-id: {1}; select-invert; selection-hide'
        , ids)
    cmds.WriteLine Format('
    (
        export-filename: png\{2}{1}.png
        export-area-page
        export-width: {1}; export-height: {1}
        export-do
    )', size, prefix)
}