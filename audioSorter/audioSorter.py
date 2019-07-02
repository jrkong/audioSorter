import argparse
import mutagen
from pathlib import Path, PureWindowsPath
from win32com.client import Dispatch
import shutil
import os

shortcutStore = "../moved-files"
logStore = []

""" FileType holds the definitions and extensions for file types """
class FileTypes():
    def __init__(self, typeIn = {}):
        self.types = typeIn
    # end of FileType constructor

    def getTypes(self):
        return self.types
    # end of getTypes

    """ setTypes(typesIn) Setter for types. Overrides the current types dictionary completely"""
    def setTypes(self, typesIn):
        self.Types = typesIn
    # end of setTypes

    """ addTypes(typesIn) adds a single type to the type dictionary (Epects a dictionary key) """
    def addTypes(self, typesIn):
        self.types.update(typesIn)
    # end of addTypes

    """ removeTypes(typesIn) removes a single type to the type dictionary (Epects a dictionary key) """
    def removeTypes(self, typesIn):
        self.types.update(typesIn)
    # end of removeTypes

    """
    isType(pathIn) takes a Path to a file or directory and checks if it is a type stored in the type dictionary. If it is, return the type name, if not return "other" 
    """
    def isType(self, pathIn):
        if pathIn.is_file():
            for type in self.types:
                suffixList = pathIn.suffixes
                if len(suffixList) == 0:
                    return "other"
                for suffix in suffixList:
                    typeList = self.types[type]
                    if suffix in typeList:
                        return type
            # end of type for loop
        # end of if
        elif pathIn.is_dir():
            typeCount = {}
            blnTyped = False
            if len(os.listdir(pathIn)) == 0:
                return "empty"
            for file in pathIn.iterdir():
                blnTyped = False
                for type in self.types:
                    for suffix in file.suffixes:
                        if suffix in self.types[type]:
                            if type not in typeCount.keys():
                                typeCount[type] = 1
                            else:
                                typeCount[type] = typeCount[type] + 1
                            blnTyped = True
                        elif file.is_dir():
                            # ignore directories
                            blnTyped = True
                # end of type check for loop
                if blnTyped is False:
                    if "other" not in typeCount.keys():
                        typeCount["other"] = 1
                    else:
                        typeCount["other"] = typeCount["other"] + 1
            # end of file check for loop
            highestTypeCount = max(typeCount.values())
            typeOfDirectory = [typeKey for typeKey in typeCount.keys() if typeCount[typeKey] == highestTypeCount]

            return typeOfDirectory[0]
    # end of isType
# end of FileTypes

""" AudioFileType holds the definition for audio filetype extensions """
class AudioFileType(FileTypes):
    def __init__(self):
        super().__init__()
        self.audioTypes = [".flac", ".alac", ".aiff", ".wav", ".mp3", ".acc", ".m4a"]
        self.types = {"audio" : self.audioTypes }
    # end of AudioFileType constructor
# end of AudioFileTypes

""" ImageFileType holds the definition for image filetype extensions """
class ImageFileType(FileTypes):
    def __init__(self):
        super().__init__()
        self.imageTypes = [".jpg", ".png", ".bmp", ".tiff"]
        self.types = {"image" : self.imageTypes }
    # end of ImageFileType constructor
# end of ImageFileTypes

""" VideoFileType holds the definition for Video filetype extensions """
class VideoFileType(FileTypes):
    def __init__(self):
        super().__init__()
        self.videoTypes = [".mp4", ".mpeg", ".m4a", ".mov", ".wmv",".webm", ".mkv"]
        self.types = {"video" : self.videoTypes }
    # end of VideoFileType constructor
# end of VideoFileTypes

""" TextFileType holds the definition for Text filetype extensions """
class TextFileType(FileTypes):
    def __init__(self):
        super().__init__()
        self.textTypes = [".txt", ".rtf" ,".doc", ".log", ".md"]
        self.types = {"text" : self.textTypes }
    # end of TextFileType constructor
# end of TextFileTypes

""" 
The FileCrawler class scans directories and performs directory related functions for audioSorter 

Each FileCrawler handles one base directory
"""
class FileCrawler:
    def __init__(self, directoryIn = ""):
        self.directory = Path(directoryIn)
        self.dirPath = self.directory.cwd()

    # end of FileCrawler constructor

    """ exists() will find the target directory exists """
    def exists(self, searchDirectory):
        # fullPath = Path(str(self.dirPath) + "/" + searchTerm)

        return Path(searchDirectory).exists()
    # end of search
    
    """ getRoot() returns the main directory the current instance of FileCrawler works on """
    def getRoot(self):
        return self.direPath
    # end of getRoot()

    """
    findAllDirectoriesofType(typeIn) finds files of the input type and returns a list of directories which contain a majority of said type 
        - typeIn is a FileType object
    """
    def findAllDirectoriesofType(self, directoryIn = None, dirListIn = []):
        typeIn = AudioFileType()
        if directoryIn == None:
            directoryIn = self.rootSource
        currentDirectory = directoryIn
        dirList = dirListIn
        dirCount = 0
        typeCount = 0
        dirFileCount = 0
        for file in directoryIn.iterdir():
            # check if inner element is directory
            if file.is_dir():
                dirCount = dirCount + 1
                # recurse through all directories
                if file != currentDirectory:
                    self.findAllDirectoriesofType(file, dirListIn)  
            else:
                if typeIn.isType(file):
                    typeCount = typeCount + 1;
        # end of for loop

        # TODO: this math sucks, it'll work but makes assumptions
        if typeCount >= ((dirFileCount - dirCount)/2):
            dirList.append(directoryIn)
        else:
            return None

        return dirList
    # end of findAllDirectoriesofType

# end of FileCrawler

""" The FileManipulator class is responsible for performing actions on files and directories """ 
class FileManipulator:
    def __init__(self, directoryIn = "", typesIn = FileTypes):
        directory = directoryIn
        # types is a list of types a FileManipulator can recognize during a search
        types = typesIn
    # end of FileManipulator constructor

    def getDirectory(self):
        return self.directory
    # end of getDirectory

    # getters and setters
    def setDirectory(self, directoryIn):
        self.directory = directoryIn
    # end of setDirectory

    """ move(sourceIn) will move the source file or directory into the directory FileManipulator is working on """
    def move(self, sourceIn, directoryIn):
        print("Moving: " + str(directoryIn))
        try:
            shutil.move(sourceIn, directoryIn)
        except:
            logStore.append("Failed to move: " + str(sourceIn))
            logStore.append("Move target directory: " + str(directoryIn))    # end of moveFile

    # Directory functions
    def createDirectory(self, dirName):
        # make new directory logic goes here
        os.makedirs(dirName)
    # end createDirectory

    """ copyDirectory(sourceIn) will copy everything in the source directory into the directory FileManipulator is working on """
    def copyDirectory(self, sourceIn, destinationIn, allowedTypes = None):
        # Python copy does not recursivly copy the directory structure, do it yourself
        # This needs to be overrided 
        if destinationIn.exists() == False:
            shutil.copytree(sourceIn, destinationIn)
        else:
            for file in sourceIn.iterdir():
                if file.is_file():
                    filename = file.name
                    destinationFile = Path(str(destinationIn) + "/" + filename)
                    self.copy(file, destinationFile)
                elif destination.is_dir():
                    if allowedTypes == None:
                        self.copyDirectory(sourceIn, destinationIn)
                    else:
                        dirType = allowedTypes.isType(destinationIn)
                        for type in allowedTypes:
                            if dirType in allowedTypes[type]:
                                self.copyDirectory(sourceIn, destinationIn)
                            
    # end of copyDirectory

    """ 
    deleteDirectory(directoryIn, mode) will delete everything in the target directory and delete the target directory. 
     - If no inputs are given then the working directory FileManipulator is working on is deleted
     - If no inputs are given then the working directory FileManipulator is working on is deleted
    """
    def deleteDirectory(self, directoryIn = None):
        if directoryIn is None:
            directoryIn = self.directory
        shutil.rmtree(directoryIn)
    # end of deleteDirectory

    # File related functions
    """ copy(sourceIn) will copy the source file or directory into the directory FileManipulator is working on """
    def copy(self, sourceIn, directoryIn = None):
        if directoryIn is None:
            directoryIn = self.directory
        print("Copying: " + str(directoryIn))
        try:
            shutil.copyfile(sourceIn, directoryIn)
        except:
            logStore.append("Failed to copy: " + str(sourceIn))
            logStore.append("Copy target directory: " + str(directoryIn))
    # end of copyFile

    """ deleteFile(file) will delete the target file in directory being worked on. """
    def deleteFile(self, file):
        file.unlink()
    # end of deleteFile
# end of FileManipulator

"""
The AudioFile class holds the data of a singular audio file in a managable manner
"""
class AudioFile():
    def __init__(self):
        # all the tags that'll be used to sort the file goes here!
        self.title = ""
        self.album = ""
        self.artist = ""
        self.releaseType = []
        self.releaseYear = ""
        self.label = ""
    # end of AudioScanner constructor
    
    def scanAudio(self, filePath):
        # TODO: Test mapping
        audioFile = mutagen.File(filePath)
        
        # print(audioFile)

        self.title = self.title.join(audioFile["title"])
        self.album = self.album.join(audioFile["album"])
        # save artist
        if "albumartist" in audioFile.keys():
            self.artist = self.artist.join(audioFile["albumartist"])
        elif "artist" in audioFile.keys():
            self.artist = self.artist.join(audioFile["artist"])
        # release type stays as a list!
        if "releasetype" in audioFile.keys():
            self.releaseType = audioFile["releasetype"]
        if "originalyear" in audioFile.keys():
            self.releaseYear = self.releaseYear.join(audioFile["originalyear"])
        if self.releaseYear == "" and "year" in audioFile.keys():
            self.releaseYear = self.releaseYear.join(audioFile["year"])
        
        # attempt to save label (mostly used for Monstercat albums)
        if "releaselabel"in audioFile.keys():
            self.label = self.label.join(audioFile["releaselabel"])
        if "label"in audioFile.keys():
            self.label = self.label.join(audioFile["label"])
    # end of scanAudio
# end of AudioFile

""" FindAudio() takes the root source directory and finds all audio directories """
class FindAudio():
    def __init__(self, sourceIn = None):
        self.rootSource = sourceIn
    # end of FindAudio constructor

    def setRoot(self, sourceIn):
        self.rootSource = sourceIn
    # end of setRoot

    def findAllAudioDirectories(self, directoryIn = None, dirListIn = []):
        typeIn = AudioFileType()
        if directoryIn == None:
            directoryIn = self.rootSource
        currentDirectory = directoryIn
        dirList = dirListIn
        dirCount = 0
        typeCount = 0
        dirFileCount = 0
        # print(currentDirectory)

        for file in directoryIn.iterdir():
            # check if inner element is directory
            if file.is_dir():
                dirCount = dirCount + 1
                # recurse through all directories
                if file != currentDirectory:
                    # print("file: " + str(file))
                    # print("dir : " + str(currentDirectory))
                    self.findAllAudioDirectories(file, dirListIn)  
            else:
                if typeIn.isType(file) in typeIn.types.keys():
                    typeCount = typeCount + 1

            # print(file)
        # end of for loop

        # TODO: this math sucks, it'll work but makes assumptions
        if typeCount >= 1:
            dirList.append(directoryIn)
        else:
            if currentDirectory == self.rootSource:
                return dirList
            else:
                return None

        return dirList
    # end of FindAudio

""" 
AudioOrganizer() organizes audio according to the structure defined in buildPath (should update and abstract more later)

V1 assumptions:
 - AudioOrganizer assumes the first audio file it scans represents the contents of the rest of the directory 
    - Everything in the directory belongs in the same album 
"""
class AudioOrganizer():
    def __init__(self, sourceDirIn = None, targetDirectoryIn = None, backupDirectoriesIn = None):
        self.audioFileType = AudioFileType()
        self.audioIn = None
        for file in sourceDirIn.iterdir():
            if self.audioFileType.isType(file) == "audio":
                # yes file is an audio file
                audioIn = file
                break
        self.audio = AudioFile()
        self.audio.scanAudio(audioIn)

        self.sourceDir = sourceDirIn
        self.artistDirName = ""
        self.releaseTypeDirName = ""
        self.albumDirName = ""

        self.targetDirectory = targetDirectoryIn
        self.backupDirectoryList = backupDirectoriesIn

        self.builtArtistDirectory = None
        self.builtTargetDirectory = None

        # set permitted types
        self.permittedTypes = FileTypes()
        self.permittedTypes.addTypes(AudioFileType().getTypes())
        self.permittedTypes.addTypes(ImageFileType().getTypes())
        self.permittedTypes.addTypes(VideoFileType().getTypes())

        # file manipulator here
        self.manipulator = FileManipulator()
    # end of AudioOrganizer constructor

    """ pathLinter(self, stringIn) lints the inputted string to ensure it is path safe  """
    def pathLinter(self, stringIn):
        invalidChars = ['/', '\\', ':', '*', '?', '<', '>', '|']
        strReturn = stringIn
        for char in invalidChars:
            if char in stringIn:
                stringIn = stringIn.replace(char, ' -')
        if '"' in stringIn:
            stringIn = stringIn.replace('"', "'")
        strReturn = stringIn
        return strReturn

    def organizeMusic(self):
        # build directory structure
        # create backups
        self.createBackups(self.builtTargetDirectory)

        # create structure for target directory (this is done second because moving is more destructive)
        self.buildDirectoryStructure()
        # self.organizeFiles(mode="copy")
        self.organizeFiles()
        
        # delete original folder
        # self.manipulator.deleteDirectory(self.sourceDir)

    # end of organizeMusic

    """ 
    buildDirectoryStructure(self, structureTarget) creates the directory structure for the music library
        - structureTarget: can be either "target" or "backup"
    """
    def buildDirectoryStructure(self, structureTarget = "target"):
        # build artist directory
        if structureTarget == "target":
            self.builtArtistDirectory = self.createArtistDirectory(self.targetDirectory)
        
            # assume structureTarget will always be populated in backup mode
            self.builtTargetDirectory = self.createAlbumDirectory(self.builtArtistDirectory)
        
        elif structureTarget == "backup":
            builtBackupArtistDirectory = None
            builtBackupTargetDirectory = None
            for backup in self.backupDirectoryList:
                builtBackupArtistDirectory = self.createArtistDirectory(backup)
        
                # assume structureTarget will always be populated in backup mode
                builtBackupTargetDirectory = self.createAlbumDirectory(builtBackupArtistDirectory)
        # print("leaving buildDirectoryStructure")
        
        return True
    # end of buildDirectoryStructure

    def createArtistDirectory(self, targetDir):
        # targetDir is used so this can be reused for backup and target directory generation

        # monstercat check here:
        if self.audio.label == "Monstercat" and self.audio.artist != "Infected Mushroom" :
            destArtistDir = Path(str(targetDir) + "/Monstercat")
            self.artistDirName = "Monstercat"
        else:
            strArtist = self.pathLinter(self.audio.artist)
            destArtistDir = Path(str(targetDir) + "/" + strArtist)
            self.artistDirName = strArtist
            # print(destArtistDir)
        crawler = FileCrawler(targetDir)
        if not crawler.exists(destArtistDir):
            manipulator = FileManipulator(destArtistDir)
            manipulator.createDirectory(destArtistDir)
        else:
            print ("log: Artist path: " + str(destArtistDir) + " already exists")

        # print("leaving createArtistDirectory")
        return destArtistDir
    # end of createArtistDirectory

    def createAlbumDirectory(self, targetDir):
        # targetDir is used so this can be reused for backup and target directory generation
        destinationDir = ""
        destinationType = ""

        # determine which release type (Soundtrack, Album, Single) this album belongs in
        if "soundtrack" in self.audio.releaseType:
            destinationType = "Soundtrack"
        elif "single" in self.audio.releaseType:
            destinationType = "Single"
        else:
            destinationType = "Album"
        # end of if tree

        # create the upper folder if it doesn't exist
        destinationDir = Path(str(targetDir) + "/" + destinationType)
        self.releaseTypeDir = destinationType

        crawler = FileCrawler(targetDir)
        if not crawler.exists(destinationDir):
            manipulator = FileManipulator(destinationDir)
            manipulator.createDirectory(destinationDir)
        else:
            print ("log: Upper directory: " + destinationType + " already exists")
        # end of if tree

        # add current directory structure to the chain
        strAlbum = self.pathLinter(self.audio.album)
        self.albumDirName = "(" + self.audio.releaseYear + ") " + strAlbum
        destinationDir = Path(str(destinationDir) + "/" + self.albumDirName)
        crawler = FileCrawler(targetDir)
        if not crawler.exists(destinationDir):
            manipulator = FileManipulator(destinationDir)
            manipulator.createDirectory(destinationDir)
        else:
            print ("log: Album path: " + str(destinationDir) + " already exists")
        # end of if tree

        # print("leaving createAlbumDirectory")
        return destinationDir
    # end of createAlbumDirectory

    """ 
    organizeFiles(mode) moves or copies all music related files into the new folder
        - mode can be either "move" or "copy"
        - destination can be either "target" or "backup"
    """
    def organizeFiles(self, mode = "move", destinationType = "target", directoryIn = None):
        # move all files from source directory to new directory
        sourcePath = Path(self.sourceDir)
        scMaker = ShortcutMaker()

        if mode == "move":
            scMaker.setTarget(self.builtTargetDirectory)
            scMaker.setShortcutLocation(shortcutStore)

        # moved to organizeMusic to decouple from organizeFiles
        # build path
        # self.buildDirectoryStructure()
        
        print("Working on: " + str(sourcePath))
        destination = None
        for source in sourcePath.iterdir():
            # check if source file is permitted type
           if self.permittedTypes.isType(pathIn = source) in self.permittedTypes.types.keys() or source.is_dir():
                if destinationType == "target":
                    destination = Path(str(self.builtTargetDirectory) + "/" + source.name)
                elif destinationType == "backup":
                    backupDirectoryPath = str(directoryIn) + "/" + self.artistDirName + "/" + self.releaseTypeDir + "/" + self.albumDirName
                    destination = Path(backupDirectoryPath + "/" + source.name)
                if mode == "move":
                    self.manipulator.move(source, destination)
                elif mode == "copy":
                    allowedDirectoryTypes = FileTypes()
                    allowedDirectoryTypes.addTypes(ImageFileType().getTypes())
                    allowedDirectoryTypes.addTypes(VideoFileType().getTypes())
                    if source.is_file():
                        self.manipulator.copy(source, destination)
                    elif source.is_dir():
                        self.manipulator.copyDirectory(source, destination, allowedDirectoryTypes)
            # end of if statement
        # end of for loop

        if mode == "move":
            # create shortcut
            scMaker.genShortcut()

        # print("leaving organizeFiles")
        return True
    # end of organizeFiles

    def createBackups(self, sourceIn):
        self.buildDirectoryStructure("backup")
        for directory in self.backupDirectoryList:
            self.organizeFiles("copy", "backup", directory)
        # end of for loop

        # print("leaving createBackups")
        return True
    # end of createBackups
# end of AudioOrganizer

""" ShortcutMaker generates shortcuts """
class ShortcutMaker():
    def __init__(self):
        self.shortcutLocation = ""
        self.target = ""
    # end of ShortcutMaker constructor

    def setShortcutLocation(self, pathIn):
        shortcutString = pathIn + "/" + self.target.name + ".lnk" 
        shortcutPath = Path(shortcutString)
        self.shortcutLocation = shortcutPath
    # end of setShortcutLocation

    def setTarget(self, pathIn):
        self.target = pathIn
    # end of setTarget

    """
    genSortcut() generates a shortcut to the moved directory 

    This exists as a "KONG WOULD LIKE THIS" just to make adding new playlists easier
    """
    def genShortcut(self):
        manipulator = FileManipulator(shortcutStore)
        crawler = FileCrawler(shortcutStore)

        # check if shortcut storage location exists
        if not crawler.exists(Path(shortcutStore)):
            manipulator = FileManipulator(shortcutStore)
            manipulator.createDirectory(shortcutStore)

        shell = Dispatch("WScript.Shell")
        strShortcut = str(self.shortcutLocation)
        strTarget = str(self.target)
        shortcut = shell.CreateShortCut(strShortcut)
        try:
            shortcut.Targetpath = strTarget
            print("Creating shortcut to: " + strTarget)
            shortcut.save()
        except:
            print("no go with shortcut for: " + strTarget)
        '''
        # we don't need a working directory here
        shortcut.WorkingDirectory = wDir 

        # no need for icons either
        if icon == '':
            pass
        else:
            shortcut.IconLocation = icon
        '''
    # end of genShortcut
# end of ShortcutMaker

# temp function to test functionality before conf file implementation
def findAndMoveAudio():
    # TODO: later we'll need to dig for the sourcePath

    # backup and organize download folders
    sourcePath = Path("K:\\Users\\Alex\Desktop\\To add\\english")
    targetPath = Path("A:\\Users\\Alex\\Music\\FLAC Files")
    backupPaths = [Path("D:\\Backup\\A Storage Drive\\Music\\FLAC Files")]
    sourceList = FindAudio(sourcePath).findAllAudioDirectories()
    for source in sourceList:
        AudioOrganizer(source, targetPath, backupPaths).organizeMusic()
    
    print("DONE?!?!?!?!")
# end of findAndMoveAudio()

def main():
    # command line argument parsing
    argparser = argparse.ArgumentParser()
    argparser.add_argument("-l", "--logging", help = "Disable logging", action = "store_false")
    argparser.add_argument("-t", "--target", help = "Target directory", action = "store")
    argparser.add_argument("-b", "--backup", help = "Backup directory or backup directories. Accepts multiple inputs", action = "append")
    argparser.add_argument("-i", "--interactive", help = "Interactive sorting", action = "store_true")

    findAndMoveAudio()

    logFile = open("Audio Organizer.log", "w+")
    for line in logStore:
        logFile.write(line + "\n")
# end of main()

if __name__ == '__main__':
    main()
# end of __main__ call