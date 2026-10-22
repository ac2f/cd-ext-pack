#!/usr/bin/env python3
"""ac2f pack - VBA derleyicisinin yakaladigi, vbapy/lint'in kacirdigi hatalar.

vbapy kodu Python'a cevirip calistirir; ama gercek VBA derleyicisi daha
katidir ve CorelDRAW'in tur kutuphanesine baglidir. Bu arac kaynagi
derlemeden tarar ve su hatalari arar (hepsi CorelDRAW'da modulu hic
calistirmadan "Compile error" ya da ilk cagrida calisma hatasi verir):

  * ByRef argument type mismatch: ByRef (varsayilan!) bir parametreye
    baska turden bir degisken ya da dizi ogesi gecmek
  * dizi bekleyen parametreye dizi olmayan arguman
  * Sub'i Call olmadan parantez ve birden cok argumanla cagirmak
  * arguman sayisi, baska modulun Private yordamini cagirmak
  * ayni kapsamda iki kez bildirilen ad, parametreyle ayni adda yerel
  * erken bagli CorelDRAW nesnesinde olmayan uye ("Method or data
    member not found") - bilinen uyelerin listesi asagida
  * nesne degiskenine Set'siz atama / nesne olmayana Set
  * For Each degiskeni Variant ya da nesne degil
  * Exit Sub/Function yanlis yordamda, tanimsiz GoTo etiketi
  * 1023 karakteri asan satir, 24'ten fazla satir devami
  * Integer * Integer gibi ara sonucu tasabilecek islemler (uyari)

Kullanim:  python3 tools/vbacheck.py [src]
"""
import os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vbapy import logical_lines, split_colons, split_commas, tokenize  # noqa: E402

# ---------------------------------------------------------------------
# CorelDRAW object model: class -> {member: return class or '' (scalar)}
# Only members this project may touch are listed; anything else on an
# early bound variable is reported so it gets checked against the docs.
# ---------------------------------------------------------------------
COREL = {
    'application': {
        'activedocument': 'document', 'activepage': 'page', 'activelayer': 'layer',
        'activeselection': 'shape', 'activeselectionrange': 'shaperange',
        'activeshape': 'shape', 'activewindow': '', 'optimization': '',
        'refresh': '', 'version': '', 'versionmajor': '', 'versionminor': '',
        'eventsenabled': '', 'documents': '', 'createshaperange': 'shaperange',
        'vba': '', 'gmsmanager': '', 'path': '', 'setuppath': '', 'activetool': '',
        'createdocument': 'document', 'opendocument': 'document', 'unit': '',
        'vgcore': '', 'activeview': '', 'visible': '', 'appwindow': '',
    },
    'document': {
        'unit': '', 'begincommandgroup': '', 'endcommandgroup': '', 'undo': '',
        'redo': '', 'clearselection': '', 'activepage': 'page', 'activelayer': 'layer',
        'selection': 'shape', 'selectionrange': 'shaperange', 'pages': '',
        'name': '', 'fullfilename': '', 'filepath': '', 'filename': '',
        'save': '', 'saveas': '', 'exportbitmap': '', 'export': '', 'exportex': '',
        'dirty': '', 'referencepoint': '', 'worldscale': '', 'resolution': '',
        'activate': '', 'close': '', 'addpages': '', 'insertpages': '',
        'drawingoriginx': '', 'drawingoriginy': '', 'activeshape': 'shape',
        'createshaperangefromarray': 'shaperange', 'createcurve': 'curve', 'title': '', 'index': '',
        'masterpage': 'page', 'printsettings': '', 'properties': '', 'rulerunits': '',
        'savesettings': '', 'restoresettings': '', 'sourceformat': '',
    },
    'page': {
        'sizewidth': '', 'sizeheight': '', 'layers': 'layers', 'shapes': 'shapes',
        'activelayer': 'layer', 'createlayer': 'layer', 'activate': '', 'name': '',
        'index': '', 'setsize': '', 'boundingbox': '', 'getboundingbox': '',
        'selectshapesfromrectangle': 'shaperange', 'findshapes': 'shaperange',
        'findshape': 'shape', 'orientation': '', 'centerx': '', 'centery': '',
        'leftx': '', 'rightx': '', 'topy': '', 'bottomy': '', 'color': '',
    },
    'layer': {
        'name': '', 'shapes': 'shapes', 'activate': '', 'createrectangle': 'shape',
        'createrectangle2': 'shape', 'createellipse': 'shape', 'createellipse2': 'shape',
        'createlinesegment': 'shape', 'createcurve': 'shape', 'createcurvesegment': 'shape',
        'createartistictext': 'shape', 'createparagraphtext': 'shape',
        'createpolygon': 'shape', 'visible': '', 'printable': '', 'editable': '',
        'delete': '', 'index': '', 'page': 'page', 'ismasterlayer': '', 'color': '',
        'findshapes': 'shaperange', 'findshape': 'shape', 'import': '',
        'selectableshapes': '', 'override': '', 'moveabove': '', 'movebelow': '',
        'createguide': 'shape', 'createguideangle': 'shape', 'pastespecial': '',
        'paste': '', 'createsymbol': 'shape', 'iconoverride': '',
    },
    'layers': {'count': '', 'item': 'layer', '_index': 'layer', 'find': 'layer'},
    'shapes': {'count': '', 'item': 'shape', '_index': 'shape', 'all': 'shaperange',
               'range': 'shaperange', 'findshapes': 'shaperange', 'findshape': 'shape',
               'first': 'shape', 'last': 'shape', 'allexcluding': 'shaperange'},
    'shape': {
        'type': '', 'name': '', 'shapes': 'shapes', 'curve': 'curve', 'displaycurve': 'curve',
        'getboundingbox': '', 'setboundingbox': '', 'leftx': '', 'rightx': '', 'topy': '',
        'bottomy': '', 'centerx': '', 'centery': '', 'positionx': '', 'positiony': '',
        'sizewidth': '', 'sizeheight': '', 'setposition': '', 'getposition': '',
        'setsize': '', 'getsize': '', 'move': '', 'rotate': '', 'rotateex': '', 'flip': '',
        'stretch': '', 'stretchex': '', 'skew': '', 'duplicate': 'shape', 'delete': '',
        'convertto curves': '', 'converttocurves': '', 'group': 'shape', 'ungroup': '',
        'ungroupex': 'shaperange', 'ungroupall': '', 'ungroupallex': 'shaperange',
        'addtoselection': '', 'removefromselection': '', 'createselection': '',
        'selected': '', 'layer': 'layer', 'page': 'page', 'outline': 'outline',
        'fill': 'fill', 'text': '', 'locked': '', 'visible': '', 'parent': 'shape',
        'rotationangle': '', 'rotationcenterx': '', 'rotationcentery': '',
        'setrotationcenter': '', 'staticid': '', 'objectdata': '', 'combine': 'shape',
        'breakapart': '', 'breakapartex': 'shaperange', 'orderfronttoback': '',
        'ordertofront': '', 'ordertoback': '', 'orderforwardone': '',
        'orderbackone': '', 'orderisinfrontof': '', 'orderfrontof': '', 'orderbackof': '',
        'transformationmatrix': '', 'gettransformationmatrix': '',
        'settransformationmatrix': '', 'applyeffectinvert': '', 'clearrotation': '',
        'powerclip': '', 'powerclipparent': 'shape', 'addtopowerclip': '',
        'isonshape': '', 'createcontour': '', 'createblend': '', 'weld': 'shape',
        'trim': 'shape', 'intersect': 'shape', 'createboundary': 'shape',
        'copy': '', 'cut': '', 'clone': 'shape', 'getlinkedshapes': 'shaperange',
        'rectangle': '', 'ellipse': '', 'polygon': '', 'bitmap': '', 'effects': '',
        'effect': '', 'dimension': '', 'connector': '', 'symbol': '', 'snappoints': '',
        'index': '', 'zorder': '', 'tag': '', 'absoluteh': '', 'absolutev': '',
        'originalwidth': '', 'originalheight': '', 'boundingbox': '',
        'alignandistribute': '', 'aligntoshape': '', 'aligntopage': '',
        'aligntopagecenter': '', 'aligntoshaperange': '', 'aligntopoint': '',
        'customcommand': '', 'printable': '', 'transparency': '', 'style': '',
        'drawingoriginx': '', 'drawingoriginy': '', 'virtual': '', 'wrapText': '',
    },
    'shaperange': {
        'count': '', 'item': 'shape', '_index': 'shape', 'add': '', 'addrange': '',
        'remove': '', 'removeall': '', 'removerange': '', 'group': 'shape',
        'ungroup': '', 'ungroupex': 'shaperange', 'ungroupall': '',
        'ungroupallex': 'shaperange', 'getboundingbox': '', 'setboundingbox': '',
        'delete': '', 'move': '', 'rotate': '', 'flip': '', 'stretch': '',
        'duplicate': 'shaperange', 'createselection': '', 'addtoselection': '',
        'removefromselection': '', 'converttocurves': '', 'combine': 'shape',
        'breakapart': '', 'breakapartex': 'shaperange', 'leftx': '', 'rightx': '',
        'topy': '', 'bottomy': '', 'centerx': '', 'centery': '', 'sizewidth': '',
        'sizeheight': '', 'positionx': '', 'positiony': '', 'setposition': '',
        'setsize': '', 'sort': '', 'firstshape': 'shape', 'lastshape': 'shape',
        'shapes': 'shapes', 'all': 'shaperange', 'reverserange': '', 'exists': '',
        'indexof': '', 'outline': 'outline', 'fill': 'fill', 'applyfill': '',
        'setoutlineproperties': '', 'orderfronttoback': '', 'ordertofront': '',
        'ordertoback': '', 'reverseorder': '', 'getlinkedshapes': 'shaperange',
        'copy': '', 'cut': '', 'clone': 'shaperange', 'weld': 'shape',
        'alignanddistribute': '', 'aligntopage': '', 'aligntopagecenter': '',
        'aligntoshape': '', 'aligntoshaperange': '', 'aligntopoint': '', 'distribute': '',
        'range': 'shaperange', 'createdocumentfrom': '', 'getpositions': '',
        'skew': '', 'rotateex': '',
    },
    'curve': {
        'subpaths': 'subpaths', 'nodes': 'nodes', 'segments': 'segments',
        'length': '', 'closed': '', 'createsubpath': 'subpath', 'appendcurve': '',
        'reversedirection': '', 'getboundingbox': '', 'selection': '',
        'isclockwise': '', 'area': '', 'isonsubpath': '', 'ispointinside': '',
        'getpolyline': '', 'getcopy': 'curve', 'copyassign': '', 'clearselection': '',
        'autoreduce': '', 'findsubpathatpoint': 'subpath', 'findnodeatpoint': 'node',
        'findsegmentatpoint': 'segment', 'getintersections': '', 'bezierto': '',
        'lineto': '', 'moveto': '', 'closepath': '', 'getcurvesegment': '',
        'applytransformmatrix': '', 'issimple': '', 'startnode': 'node',
        'endnode': 'node', 'subpathcount': '', 'nodecount': '', 'segmentcount': '',
        'getsegmentsinrange': '', 'getnodesinrange': '', 'joinsubpaths': '',
    },
    'subpaths': {'count': '', 'item': 'subpath', '_index': 'subpath', 'first': 'subpath',
                 'last': 'subpath', 'all': ''},
    'subpath': {
        'segments': 'segments', 'nodes': 'nodes', 'closed': '', 'length': '',
        'startnode': 'node', 'endnode': 'node', 'firstnode': 'node', 'lastnode': 'node',
        'firstsegment': 'segment', 'lastsegment': 'segment', 'reversedirection': '',
        'getpointpositionat': '', 'getboundingbox': '', 'index': '', 'parent': 'curve',
        'delete': '', 'breakapartat': 'subpath', 'appendlinesegment': 'segment',
        'appendcurvesegment': 'segment', 'appendcurvesegment2': 'segment',
        'isonsubpath': '', 'ispointinside': '', 'positionx': '', 'positiony': '',
        'getperpendicularat': '', 'gettangentat': '', 'getcurvatureat': '',
        'move': '', 'extract': 'shape', 'getcopy': 'curve', 'selectall': '',
        'area': '', 'isclockwise': '', 'getsegmentat': 'segment', 'findnodeatpoint': 'node',
        'findsegmentatpoint': 'segment', 'getintersections': '', 'isrectangle': '',
        'equaldivide': '', 'shape': 'shape', 'insertlinesegmentat': 'segment',
        'insertcurvesegmentat': 'segment', 'getposition': '', 'setposition': '',
    },
    'segments': {'count': '', 'item': 'segment', '_index': 'segment', 'first': 'segment',
                 'last': 'segment', 'all': '', 'range': ''},
    'segment': {
        'startnode': 'node', 'endnode': 'node', 'length': '', 'type': '',
        'getpointpositionat': '', 'addnodeat': 'node', 'subpath': 'subpath',
        'index': '', 'absoluteindex': '', 'next': 'segment', 'previous': 'segment',
        'startingcontrolpointx': '', 'startingcontrolpointy': '',
        'endingcontrolpointx': '', 'endingcontrolpointy': '',
        'getstartingcontrolpointposition': '', 'getendingcontrolpointposition': '',
        'setstartingcontrolpointposition': '', 'setendingcontrolpointposition': '',
        'getperpendicularat': '', 'gettangentat': '', 'getcurvatureat': '',
        'getintersections': '', 'isonsegment': '', 'bounds': '', 'getboundingbox': '',
        'getcopy': 'curve', 'getpointpositionat2': '', 'getdistance': '',
        'startingcontrolpointangle': '', 'endingcontrolpointangle': '',
        'startingcontrolpointlength': '', 'endingcontrolpointlength': '',
        'startingcontrolpoint': '', 'endingcontrolpoint': '', 'parent': 'curve',
        'selected': '', 'shape': 'shape', 'breakapartat': '', 'deletenodes': '',
    },
    'nodes': {'count': '', 'item': 'node', '_index': 'node', 'first': 'node',
              'last': 'node', 'all': '', 'range': ''},
    'node': {
        'positionx': '', 'positiony': '', 'getposition': '', 'setposition': '',
        'breakapart': 'subpath', 'joinwith': '', 'delete': '', 'type': '',
        'index': '', 'absoluteindex': '', 'subpath': 'subpath', 'segment': 'segment',
        'prevsegment': 'segment', 'nextsegment': 'segment', 'next': 'node',
        'previous': 'node', 'isending': '', 'selected': '', 'move': '',
        'parent': 'curve', 'connectwith': '', 'shape': 'shape', 'getdistancefrom': '',
        'isselected': '', 'makesmooth': '', 'makecusp': '', 'makesymmetrical': '',
        'nextnode': 'node', 'prevnode': 'node', 'subpathindex': '',
    },
    'outline': {'width': '', 'color': 'color', 'style': '', 'type': '',
                'setproperties': '', 'setnooutline': '', 'endarrow': '', 'startarrow': ''},
    'fill': {'applynofill': '', 'applyuniformfill': '', 'uniformcolor': 'color', 'type': ''},
    'color': {'rgbassign': '', 'cmykassign': '', 'rgbred': '', 'rgbgreen': '', 'rgbblue': '',
              'hexvalue': '', 'name': '', 'type': '', 'grayassign': '', 'convertto': ''},
}
COREL_GLOBALS = {
    'activedocument': 'document', 'activepage': 'page', 'activelayer': 'layer',
    'activeselection': 'shape', 'activeselectionrange': 'shaperange',
    'activeshape': 'shape', 'application': 'application',
}
COREL_CREATABLE = {'shaperange', 'collection'}
OBJECT_TYPES = set(COREL) | {'object', 'collection', 'variant'}

VBA_BUILTINS = set('''
abs atn cbool cbyte ccur cdate cdbl cdec cint clng clnglng clngptr csng cstr cvar cverr
chr chr$ chrw chrw$ asc ascw ascb cos sin tan exp log sqr sgn int fix rnd randomize round
len lenb left left$ right right$ mid mid$ trim trim$ ltrim ltrim$ rtrim rtrim$ ucase ucase$
lcase lcase$ instr instrrev replace replace$ split join strcomp strconv string string$
space space$ str str$ val format format$ formatnumber formatcurrency formatdatetime
formatpercent hex hex$ oct oct$ iif choose switch isnumeric isarray isdate isempty
iserror ismissing isnull isobject typename vartype ubound lbound array filter strreverse
now date date$ time time$ timer dateadd datediff datepart dateserial datevalue timeserial
timevalue day month year weekday hour minute second monthname weekdayname
msgbox inputbox getsetting savesetting deletesetting getallsettings environ environ$
dir dir$ curdir curdir$ filelen filedatetime getattr setattr kill mkdir rmdir chdir
chdrive filecopy name freefile eof lof loc seek input input$ createobject getobject
doevents shell sendkeys appactivate beep error error$ cverr strptr objptr varptr
callbyname rgb qbcolor nz spc tab width collection
'''.split())
VBA_CONSTS = set('''
true false nothing empty null vbcrlf vblf vbcr vbtab vbnullstring vbnullchar vbnewline
vbtextcompare vbbinarycompare vbyes vbno vbok vbcancel vbabort vbretry vbignore
vbquestion vbyesno vbinformation vbexclamation vbcritical vbokonly vbokcancel
vbyesnocancel vbretrycancel vbabortretryignore vbdefaultbutton1 vbdefaultbutton2
vbdefaultbutton3 vbsystemmodal vbapplicationmodal vbmsgboxsetforeground vbdirectory
vbnormal vbhidden vbsystem vbreadonly vbarchive vbmonday vbsunday vbuseSystem
vbusesystemdayofweek vbfirstjan1 vbstring vblong vbinteger vbdouble vbempty vbnull
vbobject vbarray vbboolean vbdate vbvariant vbsingle vbbyte vbcurrency vbdecimal
vbobjecterror vbblack vbwhite vbred vbgreen vbblue vbbinary vbunicode vbfromunicode
vbproppercase vbuppercase vblowercase vbtrue vbfalse vbusedefault vbgeneraldate
vblongdate vbshortdate vblongtime vbshorttime vbback vbformfeed vbverticaltab
vbmaximizedfocus vbminimizedfocus vbnormalfocus vbhide vbnormalnofocus
vbminimizednofocus err debug
'''.split())
VBA_TYPES = {'long', 'integer', 'double', 'single', 'string', 'boolean', 'byte',
             'variant', 'object', 'date', 'currency', 'longlong', 'longptr', 'decimal'}
NUMERIC = {'long', 'integer', 'double', 'single', 'byte', 'currency', 'date', 'longlong'}

KEYWORDS = set('''
if then else elseif end for to step next each in do loop while until wend exit sub
function select case is dim redim preserve as set let call on error resume goto gosub
return option explicit private public static const new byval byref optional paramarray
with not and or xor mod eqv imp like typeof get put open close output append binary
random access read write lock shared print line stop type enum property friend global
declare lib alias attribute me erase
'''.split())


class Issue(object):
    def __init__(self, sev, path, line, msg):
        self.sev, self.path, self.line, self.msg = sev, path, line, msg

    def __str__(self):
        return '%s:%d: %s: %s' % (os.path.basename(self.path), self.line, self.sev, self.msg)


class Var(object):
    __slots__ = ('name', 'typ', 'arr', 'fixed', 'const', 'line')

    def __init__(self, name, typ, arr=False, fixed=False, const=False, line=0):
        self.name, self.typ, self.arr, self.fixed, self.const, self.line = name, typ, arr, fixed, const, line


class Param(object):
    __slots__ = ('name', 'typ', 'arr', 'byref', 'optional', 'paramarray')

    def __init__(self, name, typ, arr, byref, optional, paramarray):
        self.name, self.typ, self.arr, self.byref = name, typ, arr, byref
        self.optional, self.paramarray = optional, paramarray


class Proc(object):
    def __init__(self, mod, name, kind, params, ret, public, line, body):
        self.mod, self.name, self.kind, self.params = mod, name, kind, params
        self.ret, self.public, self.line, self.body = ret, public, line, body


def norm(t):
    t = (t or 'variant').lower()
    if t.startswith('cdr'):
        return 'long'
    return t


DECL_RE = re.compile(r'^(\w+)(\$)?\s*(\(([^)]*)\))?\s*(?:As\s+(New\s+)?([\w.]+)(\s*\*\s*\d+)?)?\s*$', re.I)


def parse_decl_list(text):
    out = []
    for part in split_commas(text):
        m = DECL_RE.match(part.strip())
        if not m:
            out.append(None)
            continue
        typ = norm(m.group(6)) if m.group(6) else ('string' if m.group(2) else 'variant')
        arr = m.group(3) is not None
        fixed = arr and bool((m.group(4) or '').strip())
        out.append(Var(m.group(1).lower(), typ, arr, fixed))
    return out


def parse_params(text):
    out = []
    for part in split_commas(text):
        p = part.strip()
        m = re.match(r'^(Optional\s+)?(ParamArray\s+)?(ByVal\s+|ByRef\s+)?(\w+)(\$)?\s*(\(\s*\))?\s*'
                     r'(?:As\s+([\w.]+))?\s*(?:=\s*(.+))?$', p, re.I)
        if not m:
            out.append(None)
            continue
        byref = not (m.group(3) and m.group(3).strip().lower() == 'byval')
        typ = norm(m.group(7)) if m.group(7) else ('string' if m.group(5) else 'variant')
        out.append(Param(m.group(4).lower(), typ, bool(m.group(6)), byref,
                         bool(m.group(1)), bool(m.group(2))))
    return out


class Module(object):
    def __init__(self, path):
        self.path = path
        self.name = os.path.splitext(os.path.basename(path))[0].lower()
        with open(path, encoding='latin-1') as fh:
            self.src = fh.read().replace('\r\n', '\n')
        self.vars = {}
        self.procs = {}
        self.types = {}
        self.enums = {}
        self.issues = []
        self.parse()

    def issue(self, sev, line, msg):
        self.issues.append(Issue(sev, self.path, line, msg))

    def parse(self):
        L = logical_lines(self.src)
        i = 0
        while i < len(L):
            ln, code = L[i]
            s = code.strip()
            i += 1
            if not s or re.match(r'^(Attribute|Option)\b', s, re.I):
                continue
            m = re.match(r'^(?:(Public|Private|Global)\s+)?Const\s+(\w+)\s*(?:As\s+(\w+))?\s*=\s*(.+)$', s, re.I)
            if m:
                n = m.group(2).lower()
                if n in self.vars:
                    self.issue('ERROR', ln, 'duplicate declaration of %s at module level' % n)
                self.vars[n] = Var(n, norm(m.group(3)) if m.group(3) else 'const', const=True, line=ln)
                self.vars[n].fixed = (m.group(1) or 'private').lower() in ('public', 'global')
                continue
            m = re.match(r'^(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function)\s+(\w+)\s*\((.*)\)'
                         r'\s*(?:As\s+([\w.]+)(\(\))?)?\s*$', s, re.I)
            if m:
                name = m.group(3).lower()
                kind = m.group(2).lower()
                body = []
                end_re = re.compile(r'^End\s+' + kind + r'\s*$', re.I)
                while i < len(L) and not end_re.match(L[i][1].strip()):
                    body.append(L[i])
                    i += 1
                i += 1
                params = parse_params(m.group(4))
                if None in params:
                    self.issue('ERROR', ln, 'cannot parse the parameters of %s' % name)
                    params = [p for p in params if p]
                seen_opt = False
                for p in params:
                    if p.optional or p.paramarray:
                        seen_opt = True
                    elif seen_opt:
                        self.issue('ERROR', ln, '%s: required parameter %s after an Optional one' % (name, p.name))
                if name in self.procs or name in self.vars:
                    self.issue('ERROR', ln, 'ambiguous name: %s declared twice in the module' % name)
                pub = (m.group(1) or 'public').lower() == 'public'
                self.procs[name] = Proc(self, name, kind, params, norm(m.group(5)) if kind == 'function' else None,
                                        pub, ln, body)
                continue
            m = re.match(r'^(?:(Public|Private)\s+)?(Type|Enum)\s+(\w+)\s*$', s, re.I)
            if m:
                fields = {}
                end_re = re.compile(r'^End\s+' + m.group(2) + r'\s*$', re.I)
                while i < len(L) and not end_re.match(L[i][1].strip()):
                    f = L[i][1].strip()
                    if f:
                        fm = re.match(r'^(\w+)', f)
                        if fm:
                            fields[fm.group(1).lower()] = f
                    i += 1
                i += 1
                if m.group(2).lower() == 'type':
                    self.types[m.group(3).lower()] = fields
                else:
                    self.enums[m.group(3).lower()] = fields
                    for k in fields:
                        self.vars[k] = Var(k, 'long', const=True, line=ln)
                        self.vars[k].fixed = (m.group(1) or 'public').lower() == 'public'
                continue
            m = re.match(r'^(Public|Private|Dim|Global)\s+(.+)$', s, re.I)
            if m:
                if re.match(r'^Declare\b', m.group(2), re.I):
                    continue
                for v in parse_decl_list(m.group(2)):
                    if v is None:
                        self.issue('ERROR', ln, 'cannot parse declaration %r' % s)
                        continue
                    if v.name in self.vars:
                        self.issue('ERROR', ln, 'duplicate declaration of %s at module level' % v.name)
                    v.line = ln
                    v.const = False
                    # public flag kept in .fixed's place for module vars? no: separate set
                    self.vars[v.name] = v
                    if m.group(1).lower() in ('public', 'global'):
                        v.__class__ = PubVar
                continue
            self.issue('ERROR', ln, 'statement outside a procedure: %r' % s[:60])


class PubVar(Var):
    __slots__ = ()


def is_public_var(v):
    return isinstance(v, PubVar) or (v.const and v.fixed)


# ---------------------------------------------------------------------
# procedure checks
# ---------------------------------------------------------------------
class Checker(object):
    def __init__(self, mods):
        self.mods = mods
        self.modnames = set(m.name for m in mods)
        self.modmap = dict((m.name, m) for m in mods)
        self.project = 'ac2fpack'
        self.procs = {}       # name -> [Proc]
        self.pubvars = {}     # name -> [(Module, Var)]
        for m in mods:
            for p in m.procs.values():
                self.procs.setdefault(p.name, []).append(p)
            for v in m.vars.values():
                if is_public_var(v):
                    self.pubvars.setdefault(v.name, []).append((m, v))
        for n, lst in self.pubvars.items():
            if len(lst) > 1:
                for m, v in lst:
                    m.issue('ERROR', v.line, 'ambiguous name: public %s is also declared in %s' %
                            (n, ', '.join(x[0].name for x in lst if x[0] is not m)))
        for n, lst in self.procs.items():
            pubs = [p for p in lst if p.public]
            if len(pubs) > 1:
                for p in pubs:
                    p.mod.issue('WARN', p.line, 'public %s also exists in another module (ambiguous when called)' % n)
            if n in self.pubvars:
                for p in lst:
                    p.mod.issue('ERROR', p.line, 'procedure %s has the name of a public variable/constant' % n)

    # --- name resolution ---
    def lookup_var(self, proc, local, name):
        if name in local:
            return local[name]
        m = proc.mod
        if name in m.vars:
            return m.vars[name]
        lst = self.pubvars.get(name)
        if lst:
            return lst[0][1]
        return None

    def lookup_proc(self, proc, name):
        m = proc.mod
        if name in m.procs:
            return m.procs[name]
        for p in self.procs.get(name, []):
            if p.public:
                return p
        return None

    def check(self):
        for m in self.mods:
            for p in m.procs.values():
                self.check_proc(p)
            self.check_lines(m)

    def check_lines(self, m):
        cont = 0
        start = 0
        for i, ln in enumerate(m.src.split('\n'), 1):
            if len(ln) > 1023:
                m.issue('ERROR', i, 'line longer than 1023 characters (%d)' % len(ln))
            r = ln.rstrip()
            code = r
            if len(code) > 1 and code.endswith(' _'):
                if cont == 0:
                    start = i
                cont += 1
            else:
                if cont > 24:
                    m.issue('ERROR', start, 'too many line continuations (%d > 24)' % cont)
                cont = 0

    def check_proc(self, p):
        m = p.mod
        local = {}
        for prm in p.params:
            if prm.name in local:
                m.issue('ERROR', p.line, '%s: parameter %s twice' % (p.name, prm.name))
            local[prm.name] = Var(prm.name, prm.typ, prm.arr, False, False, p.line)
            local[prm.name].fixed = False
        labels = set()
        gotos = []
        stmts = []
        size = 0
        for ln, code in p.body:
            size += len(code)
            s = code.strip()
            if re.match(r'^Attribute\s', s, re.I):
                continue
            lm = re.match(r'^([A-Za-z]\w*):(\s|$)', s)
            if lm and lm.group(1).lower() not in KEYWORDS:
                labels.add(lm.group(1).lower())
                s = s[lm.end():].strip()
            for st in split_colons(s):
                stmts.extend(self.expand_if(st, ln))
        if size > 40000:
            m.issue('WARN', p.line, '%s is %d characters long; VBA refuses procedures over 64K compiled' % (p.name, size))
        # declarations first (VBA scopes them to the whole procedure)
        for ln, st in stmts:
            dm = re.match(r'^(Dim|Static)\s+(.+)$', st, re.I)
            if dm:
                for v in parse_decl_list(dm.group(2)):
                    if v is None:
                        m.issue('ERROR', ln, 'cannot parse declaration %r' % st)
                        continue
                    if v.name in local:
                        m.issue('ERROR', ln, 'duplicate declaration in current scope: %s' % v.name)
                    if v.name == p.name:
                        m.issue('ERROR', ln, 'local %s has the name of its procedure' % v.name)
                    v.line = ln
                    local[v.name] = v
                continue
            cm = re.match(r'^Const\s+(\w+)\s*(?:As\s+(\w+))?\s*=', st, re.I)
            if cm:
                n = cm.group(1).lower()
                if n in local:
                    m.issue('ERROR', ln, 'duplicate declaration in current scope: %s' % n)
                local[n] = Var(n, norm(cm.group(2)) if cm.group(2) else 'const', const=True, line=ln)
        for ln, st in stmts:
            self.check_stmt(p, local, ln, st, gotos)
        for ln, lab in gotos:
            if lab not in labels and not lab.isdigit():
                m.issue('ERROR', ln, 'label not defined: %s' % lab)

    def expand_if(self, st, ln):
        """Splits a single-line If into its parts so each is checked."""
        m = re.match(r'^If\s+(.+?)\s+Then\s+(.+)$', st, re.I)
        if not m:
            return [(ln, st)]
        cond, rest = m.group(1), m.group(2)
        out = [(ln, '__COND__ ' + cond)]
        # split on a top level Else
        toks = re.split(r'\s+Else\s+', rest, flags=re.I)
        for part in toks:
            for sub in split_colons(part):
                out.extend(self.expand_if(sub, ln))
        return out

    # --- statements ---
    def check_stmt(self, p, local, ln, st, gotos):
        m = p.mod
        low = st.lower()
        if re.match(r'^(dim|static|const)\b', low):
            # initialisers do not exist in VBA, nothing else to check
            return
        g = re.match(r'^(?:on\s+error\s+goto|goto|gosub)\s+(\w+)\s*$', low)
        if g:
            if g.group(1) != '0':
                gotos.append((ln, g.group(1)))
            return
        if re.match(r'^on\s+error\s+resume\s+next$', low) or re.match(r'^resume\b', low):
            return
        e = re.match(r'^exit\s+(sub|function|for|do|property)$', low)
        if e:
            if e.group(1) in ('sub', 'function') and e.group(1) != p.kind:
                m.issue('ERROR', ln, 'Exit %s inside a %s' % (e.group(1).title(), p.kind.title()))
            return
        if low.startswith('__cond__ '):
            self.check_expr(p, local, ln, st[9:])
            return
        bi = re.match(r'^(?:if|elseif)\s+(.+?)\s+then$', st, re.I)
        if bi:
            self.check_expr(p, local, ln, bi.group(1))
            return
        if re.match(r'^(else|end\s+if|end\s+select|loop|next(\s+\w+)?|wend|end\s+with|do)$', low):
            return
        dw = re.match(r'^(?:do|loop)\s+(?:while|until)\s+(.+)$', st, re.I) or re.match(r'^while\s+(.+)$', st, re.I)
        if dw:
            self.check_expr(p, local, ln, dw.group(1))
            return
        sc = re.match(r'^select\s+case\s+(.+)$', st, re.I)
        if sc:
            self.check_expr(p, local, ln, sc.group(1))
            return
        if re.match(r'^case\b', low):
            rest = st[4:].strip()
            if rest.lower() != 'else':
                for part in split_commas(rest):
                    part = re.sub(r'^Is\s*(<=|>=|<>|<|>|=)', '', part.strip(), flags=re.I)
                    for side in re.split(r'\s+To\s+', part, flags=re.I):
                        self.check_expr(p, local, ln, side)
            return
        fe = re.match(r'^for\s+each\s+(\w+)\s+in\s+(.+)$', st, re.I)
        if fe:
            v = self.lookup_var(p, local, fe.group(1).lower())
            if v is not None and v.typ not in ('variant', 'object') and v.typ not in COREL:
                m.issue('ERROR', ln, 'For Each control variable %s must be Variant or an object (is %s)' %
                        (v.name, v.typ))
            self.check_expr(p, local, ln, fe.group(2))
            return
        fo = re.match(r'^for\s+(\w+)\s*=\s*(.+?)\s+to\s+(.+?)(?:\s+step\s+(.+))?$', st, re.I)
        if fo:
            v = self.lookup_var(p, local, fo.group(1).lower())
            if v is None:
                m.issue('ERROR', ln, 'variable not defined: %s' % fo.group(1))
            elif v.typ not in NUMERIC and v.typ != 'variant':
                m.issue('ERROR', ln, 'For counter %s is %s' % (v.name, v.typ))
            for x in fo.group(2, 3, 4):
                if x:
                    self.check_expr(p, local, ln, x)
            return
        rd = re.match(r'^redim\s+(preserve\s+)?(.+)$', st, re.I)
        if rd:
            for part in split_commas(rd.group(2)):
                nm = re.match(r'^(\w+)\s*\((.*)\)\s*(?:As\s+\w+)?$', part.strip(), re.I)
                if not nm and re.match(r'^\w+\.\w+\s*\(', part.strip()):
                    continue      # a UDT field array
                if not nm:
                    m.issue('ERROR', ln, 'cannot parse ReDim %r' % part)
                    continue
                v = self.lookup_var(p, local, nm.group(1).lower())
                if v is None:
                    m.issue('ERROR', ln, 'ReDim of an undeclared array %s' % nm.group(1))
                elif not v.arr:
                    m.issue('ERROR', ln, 'ReDim of %s, which is not declared as an array' % v.name)
                elif v.fixed:
                    m.issue('ERROR', ln, 'array already dimensioned: %s has fixed bounds' % v.name)
                for b in split_commas(nm.group(2)):
                    for side in re.split(r'\s+To\s+', b, flags=re.I):
                        self.check_expr(p, local, ln, side)
            return
        if re.match(r'^erase\s+', low):
            return
        op = re.match(r'^(open|close|print|line\s+input|input|put|get|write)\s*#?', low)
        if op and re.match(r'^(open\s|close\b|print\s*#|line\s+input\s*#|input\s*#|put\s*#|get\s*#|write\s*#)', low):
            if low.startswith('print'):
                rest = st.split(',', 1)
                if len(rest) > 1:
                    self.check_expr(p, local, ln, rest[1].rstrip(';'))
            return
        if re.match(r'^(kill|mkdir|rmdir|chdir|filecopy|name|savesetting|deletesetting|randomize|'
                    r'mid\$?\s*\(|debug\.print|stop|beep|err\.raise|err\.clear|doevents)\b', low):
            rest = re.sub(r'^\w+\$?\s*', '', st)
            if rest and not low.startswith('mid'):
                for a in split_commas(rest):
                    self.check_expr(p, local, ln, a)
            return
        if re.match(r'^with\b', low):
            self.check_expr(p, local, ln, st[4:])
            return
        # assignment
        am = re.match(r'^(Set\s+|Let\s+)?([\w.]+(?:\s*\([^=]*\))?)\s*=\s*(.+)$', st, re.I)
        if am and not re.match(r'^(if|elseif|do|loop|while|case)\b', low):
            lhs, rhs = am.group(2).strip(), am.group(3)
            is_set = bool(am.group(1)) and am.group(1).strip().lower() == 'set'
            self.check_assign(p, local, ln, lhs, rhs, is_set)
            return
        cl = re.match(r'^Call\s+(.+)$', st, re.I)
        if cl:
            self.check_expr(p, local, ln, cl.group(1), call_stmt=True)
            return
        self.check_call_stmt(p, local, ln, st)

    def check_assign(self, p, local, ln, lhs, rhs, is_set):
        m = p.mod
        base = re.match(r'^(\w+)', lhs).group(1).lower()
        if '.' not in lhs:
            v = self.lookup_var(p, local, base)
            if v is None:
                if base == p.name and p.kind == 'function':
                    rt = p.ret or 'variant'
                    if is_set and rt not in OBJECT_TYPES:
                        m.issue('ERROR', ln, 'Set on function %s that returns %s' % (base, rt))
                    if not is_set and rt in COREL:
                        m.issue('ERROR', ln, 'object result %s assigned without Set' % base)
                elif base in ('mid', 'mid$'):
                    pass
                else:
                    m.issue('ERROR', ln, 'assignment to undeclared %s' % base)
            else:
                if v.const:
                    m.issue('ERROR', ln, 'assignment to constant %s' % base)
                if is_set and v.typ not in OBJECT_TYPES:
                    m.issue('ERROR', ln, 'Set on %s, which is %s' % (base, v.typ))
                if not is_set and v.typ in COREL and not (v.arr and '(' not in lhs):
                    m.issue('ERROR', ln, 'object %s assigned without Set' % base)
                if v.arr and '(' not in lhs and not is_set:
                    # whole array assignment: needs a dynamic array
                    if v.fixed:
                        m.issue('ERROR', ln, "can't assign to array %s with fixed bounds" % base)
                    src = self.lookup_var(p, local, rhs.strip().lower())
                    if src is not None and src.arr and src.typ != v.typ:
                        m.issue('ERROR', ln, 'array %s (%s) = array %s (%s): types differ' %
                                (base, v.typ, src.name, src.typ))
            idx = lhs[len(base):].strip()
            if idx.startswith('('):
                self.check_expr(p, local, ln, idx)
        else:
            self.check_expr(p, local, ln, lhs, lvalue=True)
        self.check_expr(p, local, ln, rhs)

    def check_call_stmt(self, p, local, ln, st):
        m = p.mod
        toks = self.toks(m, ln, st)
        if toks is None:
            return
        if not toks:
            return
        if toks[0][0] != 'id':
            m.issue('WARN', ln, 'unrecognised statement %r' % st[:60])
            return
        if toks[0][1].lower() in self.modnames and len(toks) > 2 and toks[1] == ('op', '.'):
            self.check_tokens(p, local, ln, toks, stmt_call=True)
            return
        # method call on an object: obj.member args
        j = 0
        while j + 2 < len(toks) and toks[j + 1] == ('op', '.') and toks[j + 2][0] == 'id':
            j += 2
        if j > 0 or (len(toks) > 1 and toks[1] == ('op', '(') and self.is_object_chain(p, local, toks)):
            # check the chain and the args
            self.check_tokens(p, local, ln, toks, stmt_call=True)
            return
        name = toks[0][1].lower()
        if name in self.modnames and self.lookup_var(p, local, name) is None:
            m.issue('ERROR', ln, '%s is also a module name: call it as %s.%s '
                    '(compile error "Expected variable or procedure, not module")' % (toks[0][1], toks[0][1], toks[0][1]))
        if name == self.project:
            m.issue('ERROR', ln, '%s is the project name ("Expected variable or procedure, not project")' % toks[0][1])
        target = self.lookup_proc(p, name)
        rest = toks[1:]
        if target is None:
            if name in VBA_BUILTINS:
                self.check_tokens(p, local, ln, rest)
                return
            v = self.lookup_var(p, local, name)
            if v is None:
                lst = self.procs.get(name)
                if lst:
                    m.issue('ERROR', ln, '%s is Private in module %s: "Sub or Function not defined"' %
                            (name, lst[0].mod.name))
                else:
                    m.issue('ERROR', ln, 'Sub or Function not defined: %s' % name)
            self.check_tokens(p, local, ln, toks)
            return
        # Name(a, b) without Call is a syntax error
        if rest and rest[0] == ('op', '('):
            close = self.match_paren(rest, 0)
            if close == len(rest) - 1:
                inner = rest[1:close]
                if self.top_level_commas(inner) > 0:
                    m.issue('ERROR', ln, '%s(...) with several arguments needs Call or no parentheses '
                            '(compile error "Expected: =")' % name)
                    args = self.split_args(inner)
                    self.check_args(p, local, ln, target, args)
                    return
        args = self.split_args(rest)
        self.check_args(p, local, ln, target, args)

    def is_object_chain(self, p, local, toks):
        v = self.lookup_var(p, local, toks[0][1].lower())
        return v is not None and v.typ in COREL

    def toks(self, m, ln, text):
        try:
            return tokenize(text)
        except Exception as e:
            m.issue('WARN', ln, 'cannot tokenize %r (%s)' % (text[:50], e))
            return None

    @staticmethod
    def match_paren(toks, i):
        depth = 0
        for k in range(i, len(toks)):
            if toks[k] == ('op', '('):
                depth += 1
            elif toks[k] == ('op', ')'):
                depth -= 1
                if depth == 0:
                    return k
        return -1

    @staticmethod
    def top_level_commas(toks):
        depth, n = 0, 0
        for t in toks:
            if t == ('op', '('):
                depth += 1
            elif t == ('op', ')'):
                depth -= 1
            elif t == ('op', ',') and depth == 0:
                n += 1
        return n

    @staticmethod
    def split_args(toks):
        args, cur, depth = [], [], 0
        for t in toks:
            if t == ('op', '('):
                depth += 1
            elif t == ('op', ')'):
                depth -= 1
            if t == ('op', ',') and depth == 0:
                args.append(cur)
                cur = []
                continue
            cur.append(t)
        if cur or args:
            args.append(cur)
        return args

    def check_expr(self, p, local, ln, text, call_stmt=False, lvalue=False):
        toks = self.toks(p.mod, ln, text)
        if toks:
            self.check_tokens(p, local, ln, toks, stmt_call=call_stmt)

    def arg_kind(self, p, local, arg):
        """(var, element) when the argument is a plain variable or array
        element - the only cases a ByRef parameter binds to directly."""
        if not arg or arg[0][0] != 'id':
            return None, False
        name = arg[0][1].lower()
        v = self.lookup_var(p, local, name)
        if v is None:
            return None, False
        if len(arg) == 1:
            return v, False
        if arg[1] == ('op', '(') and self.match_paren(arg, 1) == len(arg) - 1 and v.arr:
            return v, True
        return None, False

    def check_args(self, p, local, ln, target, args):
        m = p.mod
        if target.mod is not p.mod and not target.public:
            m.issue('ERROR', ln, '%s is Private in %s' % (target.name, target.mod.name))
        params = target.params
        has_pa = any(x.paramarray for x in params)
        req = len([x for x in params if not x.optional and not x.paramarray])
        if len(args) > len(params) and not has_pa:
            m.issue('ERROR', ln, '%s: %d arguments, takes at most %d' % (target.name, len(args), len(params)))
        empties = [k for k, a in enumerate(args) if not a]
        if len(args) < req or any(k < req for k in empties):
            m.issue('ERROR', ln, '%s: %d arguments, needs %d ("Argument not optional")' %
                    (target.name, len(args), req))
        for k, a in enumerate(args):
            if k < len(params) and a:
                prm = params[k]
                v, elem = self.arg_kind(p, local, a)
                if prm.arr:
                    if v is None or elem or not v.arr:
                        m.issue('ERROR', ln, '%s: parameter %s() needs an array variable' % (target.name, prm.name))
                    elif prm.typ != 'variant' and v.typ != prm.typ:
                        m.issue('ERROR', ln, '%s: array %s is %s(), parameter %s() is %s()' %
                                (target.name, v.name, v.typ, prm.name, prm.typ))
                elif prm.byref and v is not None and not v.const:
                    if v.arr and not elem:
                        if prm.typ != 'variant':
                            m.issue('ERROR', ln, '%s: array %s passed to scalar parameter %s' %
                                    (target.name, v.name, prm.name))
                    elif prm.typ != 'variant' and v.typ != prm.typ and v.typ != 'const':
                        if not (prm.typ == 'object' and v.typ in COREL):
                            m.issue('ERROR', ln, 'ByRef argument type mismatch: %s(%s As %s) gets %s %s' %
                                    (target.name, prm.name, prm.typ, v.typ, v.name))
            self.check_tokens(p, local, ln, a)

    def check_tokens(self, p, local, ln, toks, stmt_call=False):
        m = p.mod
        n = len(toks)
        i = 0
        while i < n:
            kind, val = toks[i]
            if kind != 'id':
                i += 1
                continue
            name = val.lower()
            prev = toks[i - 1] if i > 0 else None
            if prev == ('op', '.'):
                i += 1
                continue
            if name == 'new' and i + 1 < n and toks[i + 1][0] == 'id':
                cls = toks[i + 1][1].lower()
                if cls not in COREL_CREATABLE:
                    m.issue('ERROR', ln, 'New %s: not a creatable class' % toks[i + 1][1])
                i += 2
                continue
            if name in KEYWORDS or name in VBA_CONSTS:
                i += 1
                continue
            nxt0 = toks[i + 1] if i + 1 < n else None
            if name in self.modnames and nxt0 != ('op', '.') and self.lookup_var(p, local, name) is None:
                m.issue('ERROR', ln, '%s is also a module name: call it as %s.%s '
                        '(compile error "Expected variable or procedure, not module")' % (val, val, val))
            if name == self.project and nxt0 != ('op', '.'):
                m.issue('ERROR', ln, '%s is the project name ("Expected variable or procedure, not project")' % val)
            if name in self.modnames and nxt0 == ('op', '.') and i + 2 < n and toks[i + 2][0] == 'id':
                # Module.Procedure: check the call against that module
                tm = self.modmap[name]
                pn = toks[i + 2][1].lower()
                target = tm.procs.get(pn)
                if target is None or (not target.public and tm is not p.mod):
                    m.issue('ERROR', ln, '%s.%s is not a public procedure' % (val, toks[i + 2][1]))
                    i += 3
                    continue
                j = i + 3
                if j < n and toks[j] == ('op', '('):
                    close = self.match_paren(toks, j)
                    if not (stmt_call and close == n - 1 and target.kind == 'sub' and
                            self.top_level_commas(toks[j + 1:close]) > 0):
                        self.check_args(p, local, ln, target, self.split_args(toks[j + 1:close]))
                        i = close + 1
                        continue
                    m.issue('ERROR', ln, '%s(...) with several arguments needs Call' % toks[i + 2][1])
                    i = close + 1
                    continue
                if stmt_call and i == 0:
                    self.check_args(p, local, ln, target, self.split_args(toks[j:]))
                    return
                i = j
                continue
            # member chain on a typed object
            nxt = toks[i + 1] if i + 1 < n else None
            v = self.lookup_var(p, local, name)
            typ = None
            if v is not None:
                typ = v.typ if v.typ in COREL else None
                if v.arr and nxt == ('op', '(') and typ:
                    close = self.match_paren(toks, i + 1)
                    if close < 0:
                        return
                    self.check_tokens(p, local, ln, toks[i + 2:close])
                    i = close
            elif name in COREL_GLOBALS:
                typ = COREL_GLOBALS[name]
            elif name.startswith('cdr') or name.startswith('ac2f_') or name in VBA_BUILTINS:
                pass
            else:
                target = self.lookup_proc(p, name)
                if target is not None:
                    if target.kind == 'sub' and not stmt_call:
                        m.issue('ERROR', ln, 'Sub %s used as a value ("Expected Function or variable")' % name)
                    if nxt == ('op', '('):
                        close = self.match_paren(toks, i + 1)
                        args = self.split_args(toks[i + 2:close])
                        self.check_args(p, local, ln, target, args)
                        i = close + 1
                        continue
                    else:
                        if target.params and len([x for x in target.params if not x.optional]) > 0 \
                                and not stmt_call:
                            m.issue('ERROR', ln, '%s called without its arguments' % name)
                else:
                    lst = self.procs.get(name)
                    if lst:
                        m.issue('ERROR', ln, '%s is Private in module %s' % (name, lst[0].mod.name))
                    elif not re.match(r'^[a-z]$', name) or True:
                        m.issue('ERROR', ln, 'not defined: %s' % val)
                i += 1
                continue
            if typ:
                i = self.walk_members(p, local, ln, toks, i, typ)
                continue
            i += 1

    def walk_members(self, p, local, ln, toks, i, typ):
        """toks[i] is a typed object; follows .member chains, checking each."""
        m = p.mod
        n = len(toks)
        i += 1
        # collection default member: x(1)
        while i < n:
            t = toks[i]
            if t == ('op', '(') and typ in COREL and '_index' in COREL[typ]:
                close = self.match_paren(toks, i)
                self.check_tokens(p, local, ln, toks[i + 1:close])
                typ = COREL[typ]['_index']
                i = close + 1
                continue
            if t == ('op', '(') :
                close = self.match_paren(toks, i)
                self.check_tokens(p, local, ln, toks[i + 1:close])
                i = close + 1
                typ = None
                continue
            if t == ('op', '.') and i + 1 < n and toks[i + 1][0] == 'id':
                mem = toks[i + 1][1].lower()
                if typ is not None:
                    members = COREL.get(typ)
                    if members is not None and mem not in members:
                        m.issue('CHECK', ln, '%s.%s: not in the known CorelDRAW members '
                                '(compile error "Method or data member not found" if it does not exist)' %
                                (typ.title(), toks[i + 1][1]))
                        typ = None
                    elif members is not None:
                        typ = members[mem] or None
                    else:
                        typ = None
                i += 2
                continue
            break
        return i


def main(argv):
    src = argv[1] if len(argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'src')
    paths = sorted(os.path.join(src, f) for f in os.listdir(src) if f.lower().endswith('.bas'))
    mods = [Module(pth) for pth in paths]
    ch = Checker(mods)
    ch.check()
    issues = [x for m in mods for x in m.issues]
    order = {'ERROR': 0, 'CHECK': 1, 'WARN': 2}
    issues.sort(key=lambda x: (order.get(x.sev, 3), x.path, x.line))
    for x in issues:
        print(x)
    errs = sum(1 for x in issues if x.sev == 'ERROR')
    print('%d module(s), %d error(s), %d to check, %d warning(s)' % (
        len(mods), errs, sum(1 for x in issues if x.sev == 'CHECK'), sum(1 for x in issues if x.sev == 'WARN')))
    return 1 if errs else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
