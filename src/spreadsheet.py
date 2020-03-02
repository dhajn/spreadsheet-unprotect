from ruamel.std.zipfile import ZipFile, InMemoryZipFile, BadZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET
import re
from collections import namedtuple


class BadSpreadsheetError(Exception):
    pass

# Named tuple for sheets
SheetsTuple = namedtuple("SheetsTuple", ["path", "name"])

class SpreadsheetReader:
    def __init__(self, infile):
        try:
            self.infile = ZipFile(infile, mode="r")
        except BadZipFile:
            raise BadSpreadsheetError("Not a zip file")

    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_value, exc_traceback):
        self.close()

    def parseWbSheets(self):
        """
        Check for wb protection, parse sheet names and paths and check for sheet protection
        """
        xmlnsRegEx = re.compile(r"\{.*?\}")
        try:        
            workbookRoot = ET.fromstring(self.infile.read("xl/workbook.xml").decode("utf-8"))
        except KeyError:
            raise BadSpreadsheetError("workbook.xml not found or corrupt")
        xmlnsWb = xmlnsRegEx.match(workbookRoot.tag).group()
        sheetsTag = workbookRoot.find(xmlnsWb+"sheets")
        try:
            xmlnsWbRel = xmlnsRegEx.match([i for i in sheetsTag[0].attrib if i[0] == "{"][0]).group()
        except KeyError:
            raise BadSpreadsheetError("Failed to parse workbook.xml")
        sheetRels = dict([(i.attrib[xmlnsWbRel+"id"], i.attrib["name"]) for i in sheetsTag])

        try:
            relsRoot = ET.fromstring(self.infile.read("xl/_rels/workbook.xml.rels").decode("utf-8"))
        except KeyError:
            raise BadSpreadsheetError("workbook.xml.rels not found or corrupt")
        try:
            sheetPaths = dict([(i.attrib["Id"], i.attrib["Target"]) for i in relsRoot])
        except KeyError:
            raise BadSpreadsheetError("Failed to parse workbook.xml.rels")

        AllSheetsTuple = namedtuple("AllSheetsTuple", ["path", "name", "protected"])
        try:
            sheets = [AllSheetsTuple("xl/" + sheetPaths[i], sheetRels[i], "<sheetProtection" in self.infile.read("xl/"+sheetPaths[i]).decode("utf-8")) for i in sheetRels.keys()]
        except KeyError:
            raise BadSpreadsheetError("Sheet expected but not found")
        self.protectedSheets = [SheetsTuple(i.path, i.name) for i in sheets if i.protected]
        self.unprotectedSheets = [SheetsTuple(i.path, i.name) for i in sheets if not i.protected]

        wbProtTag = workbookRoot.find(xmlnsWb+"workbookProtection")
        if wbProtTag is not None:
            self.wbProt = True if len(wbProtTag.attrib) > 0 else False
        else:
            self.wbProt = False

    @property
    def zipInfolist(self):
        return self.infile.infolist()

    def getFile(self, zippedFileName):
        return self.infile.read(zippedFileName)

    @property
    def hasVba(self):
        return True if "xl/vbaProject.bin" in self.infile.namelist() else False

    @property
    def path(self):
        return self.infile.filename

    def close(self):
        self.infile.close()

class SpreadsheetWriter:
    def __init__(self, outFilePath):
        self.imz = InMemoryZipFile(compression=ZIP_DEFLATED)
        self.outFilePath = outFilePath
    
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, exc_traceback):
        self.writeClose()

    def writeClose(self):
        self.imz.write_to_file(self.outFilePath)

    def loadUnprotect(self, reader, workbook=False, sheets=None, dumpVba=False):
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
                self.imz.append(zippedFile.filename, self._getUnprotectedXml(type="workbook", xml=reader.getFile(zippedFile).decode("utf-8")))
            elif zippedFile.filename in sheetsToUnprotect:
                self.imz.append(zippedFile.filename, self._getUnprotectedXml(type="sheet", xml=reader.getFile(zippedFile).decode("utf-8")))
            elif dumpVba and zippedFile.filename == "xl/vbaProject.bin":
                continue
            else:
                self.imz.append(zippedFile.filename, reader.getFile(zippedFile))         

    def _getUnprotectedXml(self, type, xml):
        if type == "sheet":
            unprotRe = re.compile(r"<sheetProtection[\s\S]*?>")
        elif type == "workbook":
            unprotRe = re.compile(r"<workbookProtection[\s\S]*?>")
        else:
            raise ValueError("unknown type")
        return unprotRe.sub(string=xml, repl="")


    