# -*- coding: utf-8 -*-



import os, sys
import time

import argparse
import imutils
import cv2


g_miniArea = 150
g_grabValue = 80
# g_updateEclipse = 24*60*30
g_updateEclipse = 10
import GlobalCongif


def GetDestFileName(tFile):
    # if os.path.exists(os.path.dirname(tFile)) is False:
    #     os.makedirs(os.path.join(os.path.dirname(tFile)))
    lDestPath = os.path.join(os.path.dirname(tFile),"output")
    if os.path.exists(lDestPath) is False:
        os.mkdir(lDestPath)
    lOutFile = str.format("out_%s.mkv"%(os.path.basename(tFile)))
    lOutFile = os.path.join(lDestPath,lOutFile)
    if os.path.exists(lOutFile):
        os.remove(lOutFile)
    return  lOutFile


def ParseVideo(tSrcFile, tDestFile):
    if os.path.exists(tSrcFile):
        camera = cv2.VideoCapture(tSrcFile)
        print u"load file into opencv %s" % tSrcFile
    else:
        print u"file is not exist %s" % tSrcFile
        return
    # 初始化视频流的第一帧
    dWidth = int(camera.get(cv2.CAP_PROP_FRAME_WIDTH ))
    dHeight =int(camera.get(cv2.CAP_PROP_FRAME_HEIGHT))
    firstFrame = None
    lIndex = 0
    fourcc = cv2.VideoWriter_fourcc(*'XVID')
    if tDestFile is None:
        lFileName = GetDestFileName(tSrcFile)
    lOutPutWriter = cv2.VideoWriter(lFileName,fourcc, 20.0, (dWidth,dHeight))

    while True:
        # 获取当前帧并初始化occupied/unoccupied文本
        (grabbed, frame) = camera.read()
        lIndex += 1
        # print "\r",
        print "index  %d %s"%(lIndex, tSrcFile)
        # 如果不能抓取到一帧，说明我们到了视频的结尾
        if not grabbed:
            print u"end of video"
            break

        # 调整该帧的大小，转换为灰阶图像并且对其进行高斯模糊
        frame = imutils.resize(frame, width=dWidth,height=dHeight)
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (21, 21), 0)

        # 如果第一帧是None，对其进行初始化
        if firstFrame is None:
            firstFrame = gray
            print u"第一帧是None，对其进行初始化"
            continue
        # 每个1小时自动更新对比图片
        if lIndex%g_updateEclipse == 0:
            firstFrame = gray
            # print u"update bkimage %d"%lIndex

        # 计算当前帧和第一帧的不同
        frameDelta = cv2.absdiff(firstFrame, gray)
        thresh = cv2.threshold(frameDelta, g_grabValue, 255, cv2.THRESH_BINARY)[1]

        # 扩展阀值图像填充孔洞，然后找到阀值图像上的轮廓
        thresh = cv2.dilate(thresh, None, iterations=2)
        (image, contours, hierarchy ) = cv2.findContours(thresh.copy(), cv2.RETR_EXTERNAL,
                                     cv2.CHAIN_APPROX_SIMPLE)

        # 遍历轮廓
        for c in contours:
            # if the contour is too small, ignore it
            area = cv2.contourArea(c)
            if area < g_miniArea:
                # print "\r",
                # print u"ignore contour as too small skip: %d "%area
                continue
            # compute the bounding box for the contour, draw it on the frame,
            # and update the text
            # 计算轮廓的边界框，在当前帧中画出该框
            # (x, y, w, h) = cv2.boundingRect(c)
            # cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
            text = "index: %d page: %d diff: %7d  "%(lIndex,lIndex/g_updateEclipse,area,)
            # draw the text and timestamp on the frame
            # 在当前帧上写文字以及时间戳
            cv2.putText(frame, "{}".format(text), (10, 15),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 0, 255), 2)
            # cv2.putText(frame, datetime.datetime.now().strftime("%A %d %B %Y %I:%M:%S%p"),
            #             (10, frame.shape[0] - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.35, (0, 0, 255), 1)
            lOutPutWriter.write(frame)
            #显示当前帧并记录用户是否按下按键
            # cv2.imshow("Security Feed", frame)
            # # cv2.imshow("Thresh", thresh)
            # cv2.imshow("Frame Delta", frameDelta)
            # print u"record index %d"%lIndex
            # 如果q键被按下，跳出循环
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break

    # 清理摄像机资源并关闭打开的窗口
    camera.release()
    lOutPutWriter.release()
    cv2.destroyAllWindows()
    print u"done release source %s" % tSrcFile




def ImgToEQWithWindow(tFile):
    if os.path.exists(tFile) is False:
        print "file is not exist %s"%tFile
        return
    image = cv2.imread(tFile, 0)
    result = cv2.equalizeHist(image)
    cv2.imshow(tFile, result)
    lFileName = str.format("EQ_%s"%tFile)
    cv2.imwrite(lFileName, result)
    cv2.waitKey(0)
    cv2.destroyAllWindows()
    print "done"


def CVImgToEQ(tFile):
    if os.path.exists(tFile) is False:
        print u"file is not exist %s" % tFile
        return
    image = cv2.imread(tFile, 0)
    result = cv2.equalizeHist(image)
    lFileName = str.format("EQ_%s" % tFile)
    cv2.imwrite(lFileName, result)



def ImgConvertor(tImg):
    if tImg is None:
        print u"timg is none"
        return
    return cv2.equalizeHist(tImg)

def WalkTargetFiles(tDir, tExt):
    iList = []
    for (dirpath, dirnames, filenames) in os.walk(tDir):
        for itorFile in filenames:
            if os.path.splitext(itorFile)[1] == tExt:
                iList.append(itorFile)
    return iList

def BatchEQImg():
    lList = WalkTargetFiles(os.getcwd(), ".png")
    for itor in lList:
        CVImgToEQ(itor)


if __name__ == "__main__":
    # lFile = ParseCmdLine()
    # if lFile is not None:
    #     print u"%s"%lFile
    #     ParseVideo(lFile)
    # else:
    #     print "cmd is null"
    # ImgToEQWithWindow("face_in_0000.png")
    # BatchEQImg()
    GetDestFileName(sys.argv[0])