# -*- coding: utf-8 -*-



import os, sys
import subprocess as subp
import threading
import time
import GlobalCongif


def WalkTargetFiles(tDir, tExt):
    iList = []
    for (dirpath, dirnames, filenames) in os.walk(tDir):
        for itorFile in filenames:
            if os.path.splitext(itorFile)[1] == tExt:
                iList.append(itorFile)
    return iList


def BatchRename(tDir, tFormat):
    lList = WalkTargetFiles(tDir, ".png")
    lList.sort(key= str.lower)
    iCount = len(lList)
    for itor in range(iCount):
        lStrName = str.format(tFormat%itor)
        os.rename(lList[itor],lStrName)


def FindPythonPath():
    lList = ["C:\\Python27\\python.exe","D:\\Python27\\python.exe","C:\\Program Files (x86)\\Python27\\python.exe","C:\\Program Files\\Python27\\python.exe","D:\\Program Files (x86)\\Python27\\python.exe","D:\\Program Files\\Python27\\python.exe"]
    for itor in lList:
        if os.path.exists(itor):
            return itor
    return None


def SetupProcess(tSrc, tDest):
    if GlobalCongif.g_exeFile.find(".py") !=-1:
        if tDest is None:
            subp.call([FindPythonPath(),GlobalCongif.g_exeFile,"-sf",tSrc])
        else:
            subp.call([FindPythonPath(),GlobalCongif.g_exeFile,"-sf",tSrc,"-df",tDest])
    else:
        if tDest is None:
            subp.call([GlobalCongif.g_exeFile,"-sf",tSrc])
        else:
            subp.call([GlobalCongif.g_exeFile,"-sf",tSrc,"-df",tDest])


def BeginFilterVideo():
    for iExt in GlobalCongif.g_extList:
        iList = WalkTargetFiles(GlobalCongif.g_sourceDir, iExt)
        for itor in iList:
            print u"Setup Process for dealing filename :%s"%itor
            lThread =threading.Thread(target=SetupProcess,args=(itor,GetDestFileName(itor)))
            lThread.start()
            # EnsureProcessBalance(GlobalCongif.g_exeFile,GlobalCongif.g_processCount)

def GetDestFileName(tFile):
    if os.path.exists(GlobalCongif.g_destDir) is False:
        os.mkdir(GlobalCongif.g_destDir)
    if os.path.exists(tFile) is False:
        raise "could not find file %s"%tFile
    lOutFile = str.format("out_%s.mkv"%(os.path.basename(tFile)))
    lOutFile = os.path.join(GlobalCongif.g_destDir,lOutFile)


def QueryProcessCounts(tProcess):
    if isinstance(tProcess, unicode) is False:
        tProcess = tProcess.decode("GBK")
    if tProcess.find(u".py") !=  -1:
        tProcess = u"python.exe"
    if tProcess.find(u".exe") == -1:
        tProcess+= u".exe"
    lMap = {}
    import psutil
    pList = psutil.process_iter()
    for itor in pList:
        try:
            if itor.name() == tProcess:
                lMap[itor.pid] = itor.name()
        except Exception:
            continue
    return  len(lMap)


def EnsureProcessBalance(tProcess, tCount):
    lCount = QueryProcessCounts(tProcess)
    while lCount >= tCount:
        time.sleep(0.05)
        lCount = QueryProcessCounts(tProcess)





if __name__ == "__main__":
    # BeginFilterVideo()
    # BatchRename(os.getcwd(),"face_in_%04d.png")
    # QueryProcessCounts(1)
    # EnsureProcessBalance("python.exe",1)

    pass



