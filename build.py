"""Stamp a build id into index.html so the browser cannot serve stale scripts,
and so the running version is visible on screen."""
import io, re, time
BUILD = time.strftime("%Y%m%d-%H%M")
h = io.open('index.html', encoding='utf8').read()
h = re.sub(r'(<script src="(?:parser|graph|checkpoints|app)\.js)(\?v=[^"]*)?(")', r'\1?v=' + BUILD + r'\3', h)
h = re.sub(r'<span id="build">[^<]*</span>', '<span id="build">build ' + BUILD + '</span>', h)
if 'id="build"' not in h:
    h = h.replace('<span id="srcinfo"></span>',
                  '<span><span id="srcinfo"></span> &middot; <span id="build">build ' + BUILD + '</span></span>')
io.open('index.html', 'w', encoding='utf8').write(h)
io.open(chr(118)+chr(101)+chr(114)+chr(115)+chr(105)+chr(111)+chr(110)+chr(46)+chr(106)+chr(115)+chr(111)+chr(110), chr(119), encoding=chr(117)+chr(116)+chr(102)+chr(56)).write(chr(123)+chr(34)+chr(98)+chr(117)+chr(105)+chr(108)+chr(100)+chr(34)+chr(58)+chr(34)+BUILD+chr(34)+chr(125)+chr(10))
print("stamped build " + BUILD)
