"""Stamp a build id into the pages so the browser cannot serve stale scripts,
and so the running version is visible on screen. glass.html and welding.html
carry the same stamp as index.html: the three share graph.js and
station-core.js, and a tablet left open for a fortnight is exactly where a
stale script would hurt most.

Note the order inside SCRIPTS: `station-core` and `station-ui` come before
`station`, and `welding-core` before `welding`, because the alternation is
left-to-right and `station` would otherwise match the front of
`station-core.js` and leave `-core.js` unstamped."""
import io, re, time
BUILD = time.strftime("%Y%m%d-%H%M")
SCRIPTS = r'(?:parser|graph|checkpoints|station-core|station-ui|station|welding-core|welding|export|app)'

def stamp(name):
    h = io.open(name, encoding='utf8').read()
    h = re.sub(r'(<script src="' + SCRIPTS + r'\.js)(\?v=[^"]*)?(")', r'\1?v=' + BUILD + r'\3', h)
    h = re.sub(r'<span id="build">[^<]*</span>', '<span id="build">build ' + BUILD + '</span>', h)
    if 'id="build"' not in h and '<span id="srcinfo"></span>' in h:
        h = h.replace('<span id="srcinfo"></span>',
                      '<span><span id="srcinfo"></span> &middot; <span id="build">build ' + BUILD + '</span></span>')
    io.open(name, 'w', encoding='utf8').write(h)

stamp('index.html')
stamp('glass.html')
stamp('welding.html')
io.open('version.json', 'w', encoding='utf8').write('{"build":"' + BUILD + '"}\n')   # checkBuild() polls this
print("stamped build " + BUILD)
