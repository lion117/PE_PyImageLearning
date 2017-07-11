# -*- coding: utf-8 -*-



import os, sys
import argparse
import GlobalCongif
from VideoManager import BeginFilterVideo , SetupProcess , GetDestFileName
import VideoDealing

def RunUserTerminator():
    # 创建参数解析器并解析参数
    ap = argparse.ArgumentParser()
    ap.add_argument("-a", "--area", type=int, default= 150, help=u"分析的最小区域")
    ap.add_argument("-g","--gray",type=int, default= 60,help=u"亮度提取灰度值")
    ap.add_argument("-n","--count",type=int, default=1,help=u"同时允许启动的处理视频进程数")
    ap.add_argument("-e","--effect",type=bool,default=False,help=u"是否使用直方图均衡化处理视频")
    ap.add_argument("-sd","--sourcedir",help=u"视频源文件夹路径,默认为软件启动当前文件夹下")
    ap.add_argument("-dd","--destdir",help=u"生成处理后的视频存储路径,默认与源路径相同")
    # ap.add_argument("-p","--process",help=u"启动独立的进程处理视频")
    ap.add_argument("-sf","--sourcefile",help=u"处理的视频源文件路径,需要配合'-p'才能使用")
    ap.add_argument("-df","--destfile",help=u"处理后存储的文件路径,需要配合'-p'才能使用")
    # ap.add_argument("-vf", "--video", help=u"分析视频文件的路径")
    args = vars(ap.parse_args())

    if args.get("sourcedir", None):
        GlobalCongif.g_sourceDir = args.get("sourcedir")
    else:
        GlobalCongif.g_sourceDir = os.getcwd()
    if args.get("destdir", None):
        GlobalCongif.g_destDir = args.get("destdir")
        if GlobalCongif.g_destDir == GlobalCongif.g_sourceDir: # 避免同级目录下陷入死循环
            GlobalCongif.g_destDir = os.path.jion(GlobalCongif.g_sourceDir,"output")
    else:
        GlobalCongif.g_destDir = os.path.join(GlobalCongif.g_sourceDir,"output")
    GlobalCongif.g_miniArea = args.get("area")
    GlobalCongif.g_grabValue = args.get("gray")
    GlobalCongif.g_bEffect = args.get("effect")

    lSrcFile = args.get("sourcefile",None)

    if lSrcFile  and os.path.exists(lSrcFile) :
        if args.get("destfile",None) is None:
            lDestFile = GetDestFileName(lSrcFile)
        else:
            lDestFile = args.get("destfile",None)
        print u"begin to dealing file %s"%(args.get("sourcefile"))
        # SetupProcess(lSrcFile,lDestFile)
        VideoDealing.ParseVideo(lSrcFile, lDestFile)
    else:
        print u"begin to filter video"
        BeginFilterVideo()



if __name__ == "__main__":
    RunUserTerminator()


