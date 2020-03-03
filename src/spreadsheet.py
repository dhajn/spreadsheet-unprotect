from ruamel.std.zipfile import ZipFile, InMemoryZipFile, BadZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET
import re
from collections import namedtuple


class BadSpreadsheetError(Exception):
    """Error reading spreadsheet"""

    pass


class SpreadsheetReader:
    """Wrapper around ZipFile to read xml files from xlsx/xlsm spreadsheets and check for protection.

    Attributes:
        infile (zipfile.ZipFile): input xlsx/xlsm file
        wbProt (bool): True if workbook structure is protected, otherwise False
        protectedSheets (List[Tuple[str, str]]): list of sheets with protection as [path to sheet in archive, sheet name]
        unprotectedSheets (List[Tuple[str, str]]): list of sheets without protection as [path to sheet in archive, sheet name]
    """
    def __init__(self, infile):
        """Initialize reader with spreadsheet file"""
        try:
            self.infile = ZipFile(infile, mode="r")
        except BadZipFile:
            raise BadSpreadsheetError("Not a zip file")

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, exc_traceback):
        self.close()

    def parseWbSheets(self):
        """Check for wb protection, parse sheet names and paths and check for sheet protection.
        
        Initialize wbProt, protectedSheets and unprotectedSheets.
        """
        xmlnsRegEx = re.compile(r"\{.*?\}")
        try:
            workbookRoot = ET.fromstring(
                self.infile.read("xl/workbook.xml").decode("utf-8"))
        except KeyError:
            raise BadSpreadsheetError("workbook.xml not found or corrupt")
        xmlnsWb = xmlnsRegEx.match(workbookRoot.tag).group()
        sheetsTag = workbookRoot.find(xmlnsWb + "sheets")
        try:
            xmlnsWbRel = xmlnsRegEx.match(
                [i for i in sheetsTag[0].attrib if i[0] == "{"][0]).group()
        except KeyError:
            raise BadSpreadsheetError("Failed to parse workbook.xml")
        sheetRels = dict([(i.attrib[xmlnsWbRel + "id"], i.attrib["name"])
                          for i in sheetsTag])

        try:
            relsRoot = ET.fromstring(
                self.getFile("xl/_rels/workbook.xml.rels").decode("utf-8"))
        except KeyError:
            raise BadSpreadsheetError("workbook.xml.rels not found or corrupt")
        try:
            sheetPaths = dict([(i.attrib["Id"], i.attrib["Target"])
                               for i in relsRoot])
        except KeyError:
            raise BadSpreadsheetError("Failed to parse workbook.xml.rels")

        AllSheetsTuple = namedtuple("AllSheetsTuple",
                                    ["path", "name", "protected"])
        SheetsTuple = namedtuple("SheetsTuple", ["path", "name"])

        def getPath(p):
            return p[1:] if p[:4] == "/xl/" else "xl/" + p

        try:
            sheets = [
                AllSheetsTuple(
                    getPath(sheetPaths[i]),
                    sheetRels[i],
                    "<sheetProtection" in self.getFile(getPath(
                        sheetPaths[i])).decode("utf-8"),
                ) for i in sheetRels.keys()
            ]

        except KeyError:
            raise BadSpreadsheetError("Sheet expected but not found")
        self.protectedSheets = [
            SheetsTuple(i.path, i.name) for i in sheets if i.protected
        ]
        self.unprotectedSheets = [
            SheetsTuple(i.path, i.name) for i in sheets if not i.protected
        ]

        wbProtTag = workbookRoot.find(xmlnsWb + "workbookProtection")
        if wbProtTag is not None:
            self.wbProt = True if len(wbProtTag.attrib) > 0 else False
        else:
            self.wbProt = False

    @property
    def zipInfolist(self):
        """zipfile.ZipInfo: ZipInfo object with information about the file."""
        return self.infile.infolist()

    def getFile(self, zippedFileName):
        """Return the bytes of given file in spreadsheet zip archive"""
        return self.infile.read(zippedFileName)

    @property
    def hasVba(self):
        """bool: True if spreadsheet contains a VBA project, otherwise false"""
        return True if "xl/vbaProject.bin" in self.infile.namelist() else False

    @property
    def path(self):
        """str: path to spreadsheet file"""
        return self.infile.filename

    def close(self):
        self.infile.close()


class SpreadsheetWriter:
    """Writer of unprotected xlsx/xlsm files from initial source files.

    The zip file is first created in memory, actual write is delayed until __exit__() or until writeClose() is called.

    Attributes:
        outFilePath (path-like): path to write output file
        imz (ruamel.std.zipfile.InMemoryZipFile): zip file created in memory
    """
    def __init__(self, outFilePath):
        self.imz = InMemoryZipFile(compression=ZIP_DEFLATED)
        self.outFilePath = outFilePath

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, exc_traceback):
        self.writeClose()

    def writeClose(self):
        """Write output file"""
        self.imz.write_to_file(self.outFilePath)

    def loadUnprotect(self,
                      reader,
                      workbook=False,
                      sheets=None,
                      dumpVba=False):
        """Create output file in memory, unprotecting given components.

        Args:
            reader (SpreadsheetReader): source file to unprotect
            workbook (bool): indicates if workbook structure is to be unprotected
            sheets (List[Tuple[str, str]]): list of sheets to be unprotected as [path to sheet in archive, sheet name]
            dumpVba (bool): indicates if VBA project is to be removed from source file
        """
        if sheets is not None:
            if isinstance(sheets, str):
                sheetsToUnprotect = [sheets]
            else:
                if len(sheets) == 0:
                    sheetsToUnprotect = []
                else:
                    try:
                        sheetsToUnprotect = [i.path for i in sheets]
                    except AttributeError:
                        sheetsToUnprotect = sheets
        else:
            sheetsToUnprotect = []

        for zippedFile in reader.zipInfolist:
            if workbook and zippedFile.filename == "xl/workbook.xml":
                self.imz.append(
                    zippedFile.filename,
                    self._getUnprotectedXml(
                        type="workbook",
                        xml=reader.getFile(zippedFile).decode("utf-8")),
                )
            elif zippedFile.filename in sheetsToUnprotect:
                self.imz.append(
                    zippedFile.filename,
                    self._getUnprotectedXml(
                        type="sheet",
                        xml=reader.getFile(zippedFile).decode("utf-8")),
                )
            elif dumpVba and zippedFile.filename == "xl/vbaProject.bin":
                continue
            else:
                self.imz.append(zippedFile.filename,
                                reader.getFile(zippedFile))

    def _getUnprotectedXml(self, type, xml):
        if type == "sheet":
            unprotRe = re.compile(r"<sheetProtection[\s\S]*?>")
        elif type == "workbook":
            unprotRe = re.compile(r"<workbookProtection[\s\S]*?>")
        else:
            raise ValueError("unknown type")
        return unprotRe.sub(string=xml, repl="")
