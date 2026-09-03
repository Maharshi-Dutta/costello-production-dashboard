"""Stamp a build id into index.html so the browser cannot serve stale scripts,
and so the running version is visible on screen."""
import io, re, time
BUILD = time.strftime("%Y%m%d-%H%M")
h = io.open('index.html', encoding='utf8').read()
h = re.sub(r'(<script src="(?:parser|graph|app)\.js)(\?v=[^"]*)?(")', r'\1?v=' + BUILD + r'\3', h)
h = re.sub(r'<span id="build">[^<]*</span>', '<span id="build">build ' + BUILD + '</span>', h)
if 'id="build"' not in h:
    h = h.replace('<span id="srcinfo"></span>',
                  '<span><span id="srcinfo"></span> &middot; <span id="build">build ' + BUILD + '</span></span>')
io.open('index.html', 'w', encoding='utf8').write(h)
print("stamped build " + BUILD)
